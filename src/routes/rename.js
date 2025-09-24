/**
 * Rename API Routes
 * Defines all renaming-related endpoints
 */

const express = require('express');
const router = express.Router();

// Controllers
const renameController = require('../controllers/renameController');

// Middleware
const { ensureAuthenticated, logAuthenticatedRequest } = require('../auth/middleware');
const { validateRequest, sanitizeRequest } = require('../middleware/validation');
const { writeLimiter, generalLimiter } = require('../middleware/rateLimiter');
const auditLogger = require('../middleware/auditLogger');

// Apply common middleware to all routes
router.use(sanitizeRequest);
router.use(ensureAuthenticated);
router.use(logAuthenticatedRequest);
router.use(generalLimiter);

/**
 * @route POST /api/excel/rename-file
 * @desc Rename an Excel file
 * @access Private
 */
router.post('/rename-file', 
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameFile', 'body'),
    renameController.renameFile
);

/**
 * @route POST /api/excel/rename-folder
 * @desc Rename a folder
 * @access Private
 */
router.post('/rename-folder', 
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameFolder', 'body'),
    renameController.renameFolder
);

/**
 * @route POST /api/excel/rename-sheet
 * @desc Rename an Excel worksheet
 * @access Private
 */
router.post('/rename-sheet', 
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameSheet', 'body'),
    renameController.renameSheet
);

/**
 * @route POST /api/excel/rename-suggestions
 * @desc Get intelligent rename suggestions for related items
 * @access Private
 */
router.post('/rename-suggestions', 
    validateRequest('renameSuggestions', 'body'),
    renameController.getRenameSuggestions
);

/**
 * @route POST /api/excel/batch-rename
 * @desc Perform multiple rename operations in batch
 * @access Private
 */
router.post('/batch-rename', 
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('batchRename', 'body'),
    renameController.batchRename
);

module.exports = router;
