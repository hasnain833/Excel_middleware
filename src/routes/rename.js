const express = require('express');
const router = express.Router();
const renameController = require('../controllers/renameController');
const { ensureAuthenticated, logAuthenticatedRequest } = require('../auth/middleware');
const { validateRequest, sanitizeRequest } = require('../middleware/validation');
const { writeLimiter, generalLimiter } = require('../middleware/rateLimiter');
const auditLogger = require('../middleware/auditLogger');

// Apply common middleware to all routes
router.use(sanitizeRequest);
router.use(ensureAuthenticated);
router.use(logAuthenticatedRequest);
router.use(generalLimiter);


router.post('/rename-file',
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameFile', 'body'),
    renameController.renameFile
);
router.post('/rename-folder',
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameFolder', 'body'),
    renameController.renameFolder
);


router.post('/rename-sheet',
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameSheet', 'body'),
    renameController.renameSheet
);


router.post('/rename-suggestions',
    validateRequest('renameSuggestions', 'body'),
    renameController.getRenameSuggestions
);

router.post('/batch-rename',
    writeLimiter, // Apply stricter rate limiting for write operations
    auditLogger.middleware(), // Log all rename operations
    validateRequest('batchRename', 'body'),
    renameController.batchRename
);

module.exports = router;
