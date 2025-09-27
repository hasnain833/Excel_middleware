const express = require('express');
const router = express.Router();
const renameController = require('../controllers/renameController');
const { ensureAuthenticated, logAuthenticatedRequest } = require('../auth/middleware');
const { validateRequest, sanitizeRequest } = require('../middleware/validation');
const auditLogger = require('../middleware/auditLogger');

// Apply common middleware to all routes
router.use(sanitizeRequest);
router.use(ensureAuthenticated);
router.use(logAuthenticatedRequest);


router.post('/rename-file',
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameFile', 'body'),
    renameController.renameFile
);
router.post('/rename-folder',
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameFolder', 'body'),
    renameController.renameFolder
);


router.post('/rename-sheet',
    auditLogger.middleware(), // Log all rename operations
    validateRequest('renameSheet', 'body'),
    renameController.renameSheet
);


router.post('/rename-suggestions',
    validateRequest('renameSuggestions', 'body'),
    renameController.getRenameSuggestions
);

router.post('/batch-rename',
    auditLogger.middleware(), // Log all rename operations
    validateRequest('batchRename', 'body'),
    renameController.batchRename
);

module.exports = router;
