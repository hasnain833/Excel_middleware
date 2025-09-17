/**
 * Rate Limiting Middleware
 * Protects the API from abuse and ensures fair usage
 */

const rateLimit = require('express-rate-limit');
const logger = require('../config/logger');

/**
 * General API rate limiter
 */
const generalLimiter = rateLimit({
    windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000, // 15 minutes
    max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS) || 100, // Limit each IP to 100 requests per windowMs
    message: {
        error: 'Too many requests',
        message: 'Too many requests from this IP, please try again later',
        retryAfter: Math.ceil((parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000) / 1000),
        timestamp: new Date().toISOString()
    },
    standardHeaders: true, // Return rate limit info in the `RateLimit-*` headers
    legacyHeaders: false, // Disable the `X-RateLimit-*` headers
    handler: (req, res) => {
        logger.warn('Rate limit exceeded', {
            ip: req.ip,
            userAgent: req.get('User-Agent'),
            url: req.originalUrl,
            method: req.method
        });
        
        res.status(429).json({
            error: 'Too many requests',
            message: 'Rate limit exceeded. Please try again later.',
            retryAfter: Math.ceil((parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000) / 1000),
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Strict rate limiter for write operations
 */
const writeLimiter = rateLimit({
    windowMs: 5 * 60 * 1000, // 5 minutes
    max: 20, // Limit write operations to 20 per 5 minutes
    message: {
        error: 'Write rate limit exceeded',
        message: 'Too many write operations from this IP, please try again later',
        retryAfter: 5 * 60, // 5 minutes
        timestamp: new Date().toISOString()
    },
    standardHeaders: true,
    legacyHeaders: false,
    handler: (req, res) => {
        logger.warn('Write rate limit exceeded', {
            ip: req.ip,
            userAgent: req.get('User-Agent'),
            url: req.originalUrl,
            method: req.method
        });
        
        res.status(429).json({
            error: 'Write rate limit exceeded',
            message: 'Too many write operations. Please try again later.',
            retryAfter: 5 * 60,
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Authentication rate limiter (more restrictive)
 */
const authLimiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 5, // Limit auth attempts to 5 per 15 minutes
    message: {
        error: 'Authentication rate limit exceeded',
        message: 'Too many authentication attempts from this IP, please try again later',
        retryAfter: 15 * 60, // 15 minutes
        timestamp: new Date().toISOString()
    },
    standardHeaders: true,
    legacyHeaders: false,
    skipSuccessfulRequests: true, // Don't count successful requests
    handler: (req, res) => {
        logger.warn('Authentication rate limit exceeded', {
            ip: req.ip,
            userAgent: req.get('User-Agent'),
            url: req.originalUrl
        });
        
        res.status(429).json({
            error: 'Authentication rate limit exceeded',
            message: 'Too many authentication attempts. Please try again later.',
            retryAfter: 15 * 60,
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Create a custom rate limiter with specific configuration
 * @param {Object} options - Rate limiter options
 * @returns {Function} Rate limiter middleware
 */
const createCustomLimiter = (options) => {
    const defaultOptions = {
        windowMs: 15 * 60 * 1000,
        max: 100,
        message: {
            error: 'Rate limit exceeded',
            message: 'Too many requests, please try again later',
            timestamp: new Date().toISOString()
        },
        standardHeaders: true,
        legacyHeaders: false
    };

    return rateLimit({ ...defaultOptions, ...options });
};

/**
 * Dynamic rate limiter based on user type
 * @param {Object} req - Express request object
 * @param {Object} res - Express response object
 * @param {Function} next - Next middleware function
 */
const dynamicLimiter = (req, res, next) => {
    const userType = req.headers['x-user-type'] || 'standard';
    const userId = req.headers['x-user-id'];

    // Different limits based on user type
    let windowMs = 15 * 60 * 1000; // 15 minutes
    let maxRequests = 100;

    switch (userType) {
        case 'premium':
            maxRequests = 500;
            break;
        case 'enterprise':
            maxRequests = 1000;
            break;
        case 'admin':
            maxRequests = 2000;
            break;
        default:
            maxRequests = 100;
    }

    // Create dynamic limiter
    const limiter = rateLimit({
        windowMs: windowMs,
        max: maxRequests,
        keyGenerator: (req) => {
            // Use user ID if available, otherwise fall back to IP
            return userId || req.ip;
        },
        message: {
            error: 'Rate limit exceeded',
            message: `Rate limit exceeded for ${userType} user. Please try again later.`,
            retryAfter: Math.ceil(windowMs / 1000),
            timestamp: new Date().toISOString()
        },
        handler: (req, res) => {
            logger.warn('Dynamic rate limit exceeded', {
                userId: userId,
                userType: userType,
                ip: req.ip,
                url: req.originalUrl,
                method: req.method
            });
            
            res.status(429).json({
                error: 'Rate limit exceeded',
                message: `Rate limit exceeded for ${userType} user. Please try again later.`,
                retryAfter: Math.ceil(windowMs / 1000),
                timestamp: new Date().toISOString()
            });
        }
    });

    limiter(req, res, next);
};

/**
 * Skip rate limiting for certain conditions
 * @param {Object} req - Express request object
 * @returns {boolean} True to skip rate limiting
 */
const skipRateLimit = (req) => {
    // Skip rate limiting for health checks
    if (req.path === '/health' || req.path === '/health/detailed') {
        return true;
    }

    // Skip for localhost in development
    if (process.env.NODE_ENV === 'development' && req.ip === '127.0.0.1') {
        return true;
    }

    // Skip for admin users (if properly authenticated)
    const userType = req.headers['x-user-type'];
    if (userType === 'admin' && req.headers['x-admin-token'] === process.env.ADMIN_TOKEN) {
        return true;
    }

    return false;
};

/**
 * Rate limiter with skip logic
 */
const conditionalLimiter = rateLimit({
    windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000,
    max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS) || 100,
    skip: skipRateLimit,
    message: {
        error: 'Rate limit exceeded',
        message: 'Too many requests, please try again later',
        timestamp: new Date().toISOString()
    }
});

module.exports = {
    generalLimiter,
    writeLimiter,
    authLimiter,
    dynamicLimiter,
    conditionalLimiter,
    createCustomLimiter,
    skipRateLimit
};
