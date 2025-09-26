const logger = require("../config/logger");
class AppError extends Error {
  constructor(message, statusCode = 500, isOperational = true) {
    super(typeof message === "string" ? message : JSON.stringify(message));
    this.statusCode = statusCode;
    this.isOperational = isOperational;
    this.timestamp = new Date().toISOString();
    Error.captureStackTrace?.(this, this.constructor);
  }
}

const maskSecrets = (obj) => {
  if (!obj || typeof obj !== "object") return obj;
  const clone = JSON.parse(JSON.stringify(obj));
  const mask = (v) =>
    typeof v === "string" && v.length > 8
      ? v.slice(0, 4) + "***" + v.slice(-2)
      : v;

  const keysToMask = [
    "access_token",
    "refresh_token",
    "client_secret",
    "Authorization",
    "authorization",
    "x-api-key",
    "apiKey",
  ];

  const walk = (node) => {
    if (node && typeof node === "object") {
      Object.keys(node).forEach((k) => {
        if (keysToMask.includes(k)) node[k] = mask(node[k]);
        else if (typeof node[k] === "object") walk(node[k]);
      });
    }
  };

  walk(clone);
  return clone;
};


const handleGraphError = (error) => {
  const status = error?.response?.status;
  const graphError = error?.response?.data?.error;
  const retryAfter = error?.response?.headers?.["retry-after"];

  const messageFrom = () => {
    if (graphError?.message) return graphError.message;
    if (typeof error?.response?.data === "string") return error.response.data;
    return "Graph API error";
  };

  if (status === 429 || graphError?.code === "TooManyRequests") {
    const appErr = new AppError(
      "Rate limit exceeded. Please try again later",
      429
    );
    if (retryAfter) appErr.retryAfter = retryAfter;
    return appErr;
  }

  if (status === 401 || graphError?.code === "Unauthorized") {
    return new AppError("Authentication failed", 401);
  }
  if (status === 403 || graphError?.code === "Forbidden") {
    return new AppError("Access denied to the requested resource", 403);
  }
  if (status === 404 || graphError?.code === "NotFound") {
    return new AppError("Requested resource not found", 404);
  }
  if (status === 400 || graphError?.code === "BadRequest") {
    return new AppError(`Invalid request: ${messageFrom()}`, 400);
  }
  if ((status && status >= 500) || graphError?.code === "InternalServerError") {
    return new AppError("Microsoft Graph service error", 502);
  }

  return new AppError(messageFrom(), status || 500);
};

const handleValidationError = (error) => {
  const message = error.details
    ? error.details.map((d) => d.message).join(", ")
    : error.message || "Validation failed";
  return new AppError(message, 400);
};


const handleAuthError = (error) => {
  if (error?.message?.includes("AADSTS")) {
    return new AppError("Azure AD authentication failed", 401);
  }
  return new AppError("Authentication error", 401);
};

const sanitizeError = (err) => {
  const sanitized = {
    message: err.message || "Unknown error",
    name: err.name,
    stack: err.stack,
  };

  // Axios/HTTP response details
  if (err.response?.data) {
    sanitized.responseData = maskSecrets(err.response.data);
    sanitized.statusCode = err.response.status;
    sanitized.responseHeaders = maskSecrets(err.response.headers);
  }

  // Axios request details
  if (err.config?.url) {
    sanitized.requestUrl = err.config.url;
    sanitized.requestMethod = err.config.method;
    sanitized.requestHeaders = maskSecrets(err.config.headers);
  }

  return sanitized;
};

const sendErrorDev = (err, res, requestId) => {
  res.status(err.statusCode || 500).json({
    status: "error",
    error: {
      code: err.statusCode || 500,
      message: err.message,
      stack: err.stack,
      ...(err.retryAfter ? { retryAfter: err.retryAfter } : {}),
    },
    requestId,
    timestamp: err.timestamp || new Date().toISOString(),
  });
};

const sendErrorProd = (err, res, requestId) => {
  if (err.isOperational) {
    res.status(err.statusCode || 500).json({
      status: "error",
      error: {
        code: err.statusCode || 500,
        message: err.message,
        ...(err.retryAfter ? { retryAfter: err.retryAfter } : {}),
      },
      requestId,
      timestamp: err.timestamp || new Date().toISOString(),
    });
  } else {
    // Unknown/unexpected error
    res.status(500).json({
      status: "error",
      error: { code: 500, message: "Internal server error" },
      requestId,
      timestamp: new Date().toISOString(),
    });
  }
};


const globalErrorHandler = (err, req, res, next) => {
  let error =
    err instanceof AppError
      ? err
      : new AppError(
          err?.message || "Internal error",
          err?.statusCode || 500,
          false
        );

  // Log sanitized error with context
  const sanitizedError = sanitizeError(err);
  logger.error("Error occurred:", {
    ...sanitizedError,
    url: req.originalUrl,
    method: req.method,
    ip: req.ip,
    userAgent: req.get("User-Agent"),
    requestId: req.id,
    timestamp: new Date().toISOString(),
  });

  // Map known error shapes
  if (err.response && err.response.status) {
    error = handleGraphError(err);
  } else if (err.name === "ValidationError" || err.isJoi) {
    error = handleValidationError(err);
  } else if (err.message && err.message.includes("Authentication")) {
    error = handleAuthError(err);
  } else if (err.code === "ENOTFOUND" || err.code === "ECONNREFUSED") {
    error = new AppError("Service temporarily unavailable", 503);
  } else if (err.name === "SyntaxError" && err.message.includes("JSON")) {
    error = new AppError("Invalid JSON in request body", 400);
  }

  if (process.env.NODE_ENV === "development") {
    return sendErrorDev(error, res, req.id);
  }
  return sendErrorProd(error, res, req.id);
};

const handleNotFound = (req, res, next) => {
  const err = new AppError(`Route ${req.originalUrl} not found`, 404);
  next(err);
};


const catchAsync = (fn) => {
  return (req, res, next) => {
    Promise.resolve(fn(req, res, next)).catch(next);
  };
};

const handleUnhandledRejections = () => {
  const isServerless =
    process.env.SERVERLESS === "true" ||
    process.env.VERCEL === "1" ||
    !!process.env.AWS_REGION;

  if (isServerless) return;

  process.on("unhandledRejection", (err) => {
    logger.error("Unhandled Promise Rejection:", err);
    // Do not exit; allow process manager to decide (PM2/K8s/etc.)
  });

  // in middleware/errorHandler.js
  process.on("uncaughtException", (err) => {
    logger.error("Uncaught Exception:", err);
    if (process.env.NODE_ENV === "production") {
      process.exit(1);
    }
  });
};

module.exports = {
  AppError,
  globalErrorHandler,
  handleNotFound,
  catchAsync,
  handleUnhandledRejections,
  handleGraphError,
  handleValidationError,
  handleAuthError,
};
