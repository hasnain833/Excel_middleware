/**
 * Error Handling Middleware
 * Centralized error handling for the application
 */

const logger = require('../config/logger');

/**
 * Custom error class for application-specific errors
 */
class AppError extends Error {
    constructor(message, statusCode = 500, isOperational = true) {
        super(message);
        this.statusCode = statusCode;
        this.isOperational = isOperational;
        this.timestamp = new Date().toISOString();
        
        Error.captureStackTrace(this, this.constructor);
    }
}

/**
 * Handle Microsoft Graph API errors
 * @param {Object} error - Graph API error
 * @returns {Object} Formatted error response
 */
const handleGraphError = (error) => {
    const graphError = error.response?.data?.error;
    
    if (graphError) {
        switch (graphError.code) {
            case 'Forbidden':
                return new AppError('Access denied to the requested resource', 403);
            case 'NotFound':
                return new AppError('Requested resource not found', 404);
            case 'BadRequest':
                return new AppError(`Invalid request: ${graphError.message}`, 400);
            case 'Unauthorized':
                return new AppError('Authentication failed', 401);
            case 'TooManyRequests':
                return new AppError('Rate limit exceeded. Please try again later', 429);
            case 'InternalServerError':
                return new AppError('Microsoft Graph service error', 502);
            default:
                return new AppError(`Graph API error: ${graphError.message}`, 500);
        }
    }
    
    return new AppError('Unknown Graph API error', 500);
};

/**
 * Handle validation errors
 * @param {Object} error - Validation error
 * @returns {Object} Formatted error response
 */
const handleValidationError = (error) => {
    const message = error.details ? 
        error.details.map(detail => detail.message).join(', ') :
        'Validation failed';
    
    return new AppError(message, 400);
};

/**
 * Handle authentication errors
 * @param {Object} error - Authentication error
 * @returns {Object} Formatted error response
 */
const handleAuthError = (error) => {
    if (error.message.includes('AADSTS')) {
        return new AppError('Azure AD authentication failed', 401);
    }
    
    return new AppError('Authentication error', 401);
};

/**
 * Development error response (includes stack trace)
 * @param {Object} err - Error object
 * @param {Object} res - Express response object
 */
const sendErrorDev = (err, res) => {
    res.status(err.statusCode || 500).json({
        status: 'error',
        error: {
            code: err.statusCode || 500,
            message: err.message,
            stack: err.stack
        },
        timestamp: err.timestamp || new Date().toISOString()
    });
};

/**
 * Production error response (sanitized)
 * @param {Object} err - Error object
 * @param {Object} res - Express response object
 */
const sendErrorProd = (err, res) => {
    // Operational, trusted error: send message to client
    if (err.isOperational) {
        res.status(err.statusCode || 500).json({
            status: 'error',
            error: {
                code: err.statusCode || 500,
                message: err.message
            },
            timestamp: err.timestamp || new Date().toISOString()
        });
    } else {
        // Programming or other unknown error: don't leak error details
        logger.error('Unknown error:', {
            message: err.message,
            stack: err.stack,
            name: err.name
        });
        
        res.status(500).json({
            status: 'error',
            error: {
                code: 500,
                message: 'Internal server error'
            },
            timestamp: new Date().toISOString()
        });
    }
};

/**
 * Sanitize error object to prevent circular references
 * @param {Object} err - Error object
 * @returns {Object} Sanitized error data
 */
const sanitizeError = (err) => {
    const sanitized = {
        message: err.message || 'Unknown error',
        name: err.name,
        stack: err.stack
    };

    // Extract response data if it exists (from axios errors)
    if (err.response?.data) {
        sanitized.responseData = err.response.data;
        sanitized.statusCode = err.response.status;
    }

    // Extract request info if it exists
    if (err.config?.url) {
        sanitized.requestUrl = err.config.url;
        sanitized.requestMethod = err.config.method;
    }

    return sanitized;
};

/**
 * Global error handling middleware
 * @param {Object} err - Error object
 * @param {Object} req - Express request object
 * @param {Object} res - Express response object
 * @param {Function} next - Next middleware function
 */
const globalErrorHandler = (err, req, res, next) => {
    let error = { ...err };
    error.message = err.message;

    // Log sanitized error details (no circular references)
    const sanitizedError = sanitizeError(err);
    logger.error('Error occurred:', {
        ...sanitizedError,
        url: req.originalUrl,
        method: req.method,
        ip: req.ip,
        userAgent: req.get('User-Agent'),
        timestamp: new Date().toISOString()
    });

    // Handle specific error types
    if (err.response && err.response.status) {
        // Axios/HTTP errors (likely from Graph API)
        error = handleGraphError(err);
    } else if (err.name === 'ValidationError' || err.isJoi) {
        // Joi validation errors
        error = handleValidationError(err);
    } else if (err.message && err.message.includes('Authentication')) {
        // Authentication errors
        error = handleAuthError(err);
    } else if (err.code === 'ENOTFOUND' || err.code === 'ECONNREFUSED') {
        // Network errors
        error = new AppError('Service temporarily unavailable', 503);
    } else if (err.name === 'SyntaxError' && err.message.includes('JSON')) {
        // JSON parsing errors
        error = new AppError('Invalid JSON in request body', 400);
    }

    // Send error response
    if (process.env.NODE_ENV === 'development') {
        sendErrorDev(error, res);
    } else {
        sendErrorProd(error, res);
    }
};

/**
 * Handle unhandled routes (404 errors)
 * @param {Object} req - Express request object
 * @param {Object} res - Express response object
 * @param {Function} next - Next middleware function
 */
const handleNotFound = (req, res, next) => {
    const err = new AppError(`Route ${req.originalUrl} not found`, 404);
    next(err);
};

/**
 * Async error wrapper to catch errors in async route handlers
 * @param {Function} fn - Async function to wrap
 * @returns {Function} Wrapped function
 */
const catchAsync = (fn) => {
    return (req, res, next) => {
        Promise.resolve(fn(req, res, next)).catch(next);
    };
};

/**
 * Handle promise rejections and uncaught exceptions
 */
const handleUnhandledRejections = () => {
    process.on('unhandledRejection', (err) => {
        logger.error('Unhandled Promise Rejection:', err);
        // In production, you might want to gracefully shut down the server
        if (process.env.NODE_ENV === 'production') {
            process.exit(1);
        }
    });

    process.on('uncaughtException', (err) => {
        logger.error('Uncaught Exception:', err);
        // Always exit on uncaught exceptions
        process.exit(1);
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
    handleAuthError
};
