/**
 * Excel API Routes
 * Defines all Excel-related endpoints
 */

const express = require('express');
const router = express.Router();

// Controllers
const excelController = require('../controllers/excelController');

// Middleware
const { ensureAuthenticated, logAuthenticatedRequest } = require('../auth/middleware');
const { validateRequest, validateRangeValuesCompatibility, sanitizeRequest } = require('../middleware/validation');
const { writeLimiter, generalLimiter } = require('../middleware/rateLimiter');
const rangeValidator = require('../middleware/rangeValidator');
const auditLogger = require('../middleware/auditLogger');

// Apply common middleware to all routes
router.use(sanitizeRequest);
router.use(ensureAuthenticated);
router.use(logAuthenticatedRequest);
router.use(generalLimiter);

/**
 * @route GET /api/excel/workbooks
 * @desc Get all accessible workbooks
 * @access Private
 */
router.get('/workbooks', excelController.getWorkbooks);

/**
 * @route GET /api/excel/worksheets
 * @desc Get worksheets in a workbook
 * @access Private
 */
router.get('/worksheets', 
    validateRequest('getWorksheets', 'query'),
    excelController.getWorksheets
);

/**
 * @route POST /api/excel/read
 * @desc Read data from Excel range
 * @access Private
 */
router.post('/read', 
    validateRequest('readRange', 'body'),
    excelController.readRange
);

/**
 * @route POST /api/excel/write
 * @desc Write data to Excel range
 * @access Private
 */
router.post('/write', 
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all write operations
    rangeValidator.middleware(), // Validate range permissions
    validateRequest('writeRange', 'body'),
    validateRangeValuesCompatibility,
    excelController.writeRange
);

/**
 * @route POST /api/excel/read-table
 * @desc Read data from Excel table
 * @access Private
 */
router.post('/read-table', 
    validateRequest('readTable', 'body'),
    excelController.readTable
);

/**
 * @route POST /api/excel/add-table-rows
 * @desc Add rows to Excel table
 * @access Private
 */
router.post('/add-table-rows', 
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all write operations
    rangeValidator.middleware(), // Validate range permissions
    validateRequest('addTableRows', 'body'),
    excelController.addTableRows
);

/**
 * @route POST /api/excel/batch
 * @desc Perform batch Excel operations
 * @access Private
 */
router.post('/batch', 
    writeLimiter, // Apply stricter rate limiting since this can include writes
    excelController.batchOperations
);

/**
 * @route GET /api/excel/metadata
 * @desc Get Excel file metadata
 * @access Private
 */
router.get('/metadata', 
    excelController.getFileMetadata
);

/**
 * @route GET /api/excel/logs
 * @desc Get audit logs
 * @access Private
 */
router.get('/logs', 
    require('../controllers/auditController').getAuditLogs
);

module.exports = router;
