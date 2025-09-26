const winston = require("winston");
const DailyRotateFile = require("winston-daily-rotate-file");
const path = require("path");

const logDir = process.env.LOG_DIR || "./logs";

const logFormat = winston.format.combine(
  winston.format.timestamp({ format: "YYYY-MM-DD HH:mm:ss" }),
  winston.format.errors({ stack: true }),
  winston.format.json()
);

const consoleFormat = winston.format.combine(
  winston.format.colorize(),
  winston.format.timestamp({ format: "YYYY-MM-DD HH:mm:ss" }),
  winston.format.printf(({ timestamp, level, message, ...meta }) => {
    let msg = `${timestamp} [${level}]: ${message}`;
    if (Object.keys(meta).length > 0) {
      msg += ` ${JSON.stringify(meta)}`;
    }
    return msg;
  })
);

const transports = [];

// 🚀 On Vercel → console logging only (no file writes)
if (process.env.VERCEL) {
  transports.push(
    new winston.transports.Console({
      level: process.env.LOG_LEVEL || "info",
      format: consoleFormat,
    })
  );
} else {
  // 💻 Local/dev → write logs to files + console
  transports.push(
    new DailyRotateFile({
      filename: path.join(logDir, "error-%DATE%.log"),
      datePattern: "YYYY-MM-DD",
      level: "error",
      handleExceptions: true,
      maxSize: "20m",
      maxFiles: "14d",
    }),
    new DailyRotateFile({
      filename: path.join(logDir, "combined-%DATE%.log"),
      datePattern: "YYYY-MM-DD",
      handleExceptions: true,
      maxSize: "20m",
      maxFiles: "14d",
    }),
    new DailyRotateFile({
      filename: path.join(logDir, "audit-%DATE%.log"),
      datePattern: "YYYY-MM-DD",
      level: "info",
      maxSize: "20m",
      maxFiles: "30d",
      format: winston.format.combine(
        winston.format.timestamp(),
        winston.format.json()
      ),
    }),
    new winston.transports.Console({ format: consoleFormat })
  );

  // Handle exceptions in local/dev (writes to file)
  transports.push(
    new winston.transports.File({
      filename: path.join(logDir, "exceptions.log"),
    })
  );
}

const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || "info",
  format: logFormat,
  defaultMeta: { service: "excel-gpt-middleware" },
  transports,
});

// Handle unhandled promise rejections
process.on("unhandledRejection", (ex) => {
  logger.error("Unhandled promise rejection", { error: ex });
  throw ex;
});

module.exports = logger;
