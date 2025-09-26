const express = require("express");
const router = express.Router();
const excelController = require("../controllers/excelController");
const findReplaceController = require("../controllers/findReplaceController");
const excelEngineController = require("../controllers/excelEngineController");
const {
  ensureAuthenticated,
  logAuthenticatedRequest,
} = require("../auth/middleware");
const {
  validateRequest,
  validateRangeValuesCompatibility,
  sanitizeRequest,
} = require("../middleware/validation");
const { writeLimiter, generalLimiter } = require("../middleware/rateLimiter");
const rangeValidator = require("../middleware/rangeValidator");
const auditLogger = require("../middleware/auditLogger");

// Apply common middleware to all routes
router.use(sanitizeRequest);
router.use(ensureAuthenticated);
router.use(logAuthenticatedRequest);
router.use(generalLimiter);

// All Routes
router.get("/workbooks", excelController.getWorkbooks);

router.get(
  "/worksheets",
  validateRequest("getWorksheets", "query"),
  excelController.getWorksheets
);

router.post(
  "/read",
  validateRequest("readRange", "body"),
  excelController.readRange
);

router.post(
  "/write",
  writeLimiter, // Apply stricter rate limiting for write operations
  auditLogger.middleware(), // Log all write operations
  rangeValidator.middleware(), // Validate range permissions
  validateRequest("writeRange", "body"),
  validateRangeValuesCompatibility,
  excelController.writeRange
);

router.post(
  "/batch",
  writeLimiter, // Apply stricter rate limiting since this can include writes
  excelController.batchOperations
);

router.get(
  "/search",
  validateRequest("searchFiles", "query"),
  excelController.searchFiles
);

router.post(
  "/find-replace",
  writeLimiter, // Apply stricter rate limiting for write operations
  auditLogger.middleware(), // Log all find-replace operations
  validateRequest("findReplace", "body"),
  findReplaceController.findReplace
);

router.post(
  "/search-text",
  validateRequest("searchText", "body"),
  findReplaceController.searchText
);

router.get(
  "/analyze-scope",
  validateRequest("analyzeScope", "query"),
  findReplaceController.analyzeScope
);

router.post(
  "/format",
  writeLimiter, // Apply stricter rate limiting for write operations
  auditLogger.middleware(), // Log all formatting operations
  validateRequest("excelFormat", "body"),
  excelEngineController.applyFormatting
);

// File and worksheet management
router.post(
  "/create-file",
  writeLimiter,
  auditLogger.middleware(),
  validateRequest("createFile", "body"),
  excelController.createFile
);

router.post(
  "/create-sheet",
  writeLimiter,
  auditLogger.middleware(),
  validateRequest("createSheet", "body"),
  excelController.createSheet
);

router.delete(
  "/delete-file",
  writeLimiter,
  auditLogger.middleware(),
  validateRequest("deleteFile", "queryOrBody"),
  excelController.deleteFile
);

router.delete(
  "/delete-sheet",
  writeLimiter,
  auditLogger.middleware(),
  validateRequest("deleteSheet", "body"),
  excelController.deleteSheet
);

module.exports = router;
