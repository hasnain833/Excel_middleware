/**
 * Authentication Middleware
 * Handles request authentication and authorization
 */

const azureAuth = require('./azureAuth');
const logger = require('../config/logger');

/**
 * Middleware to ensure valid Azure AD token
 */
const ensureAuthenticated = async (req, res, next) => {
    try {
        // Get access token (will use cached token if valid)
        const token = await azureAuth.getAccessToken();
        
        // Add token to request object for use in controllers
        req.accessToken = token;
        req.tokenInfo = azureAuth.getTokenInfo();
        
        logger.debug('Request authenticated successfully');
        next();
    } catch (error) {
        logger.error('Authentication failed:', error);
        return res.status(401).json({
            error: 'Authentication failed',
            message: 'Unable to authenticate with Microsoft Graph API',
            timestamp: new Date().toISOString()
        });
    }
};


/**
 * Middleware to log authenticated requests
 */
const logAuthenticatedRequest = (req, res, next) => {
    if (logger.isLevelEnabled('info')) {
        logger.info('Authenticated request', {
            method: req.method,
            url: req.originalUrl,
            ip: req.ip,
            tokenValid: req.tokenInfo?.isValid || false
        });
    }
    next();
};

module.exports = {
    ensureAuthenticated,
    logAuthenticatedRequest
};
