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
const auditLogger = require("../middleware/auditLogger");

// Apply common middleware to all routes
router.use(sanitizeRequest);
router.use(ensureAuthenticated);
router.use(logAuthenticatedRequest);

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
  auditLogger.middleware(), // Log all write operations
  validateRequest("writeRange", "body"),
  validateRangeValuesCompatibility,
  excelController.writeRange
);

router.post(
  "/batch",
  excelController.batchOperations
);

router.get(
  "/search",
  validateRequest("searchFiles", "query"),
  excelController.searchFiles
);

router.post(
  "/find-replace",
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
  auditLogger.middleware(), // Log all formatting operations
  validateRequest("excelFormat", "body"),
  excelEngineController.applyFormatting
);

// Clear data (range or whole sheet)
router.post(
  "/clear-data",
  auditLogger.middleware(),
  validateRequest("clearData", "body"),
  excelController.clearData
);

// File and worksheet management
router.post(
  "/create-file",
  auditLogger.middleware(),
  validateRequest("createFile", "body"),
  excelController.createFile
);

router.post(
  "/create-sheet",
  auditLogger.middleware(),
  validateRequest("createSheet", "body"),
  excelController.createSheet
);

router.delete(
  "/delete-file",
  auditLogger.middleware(),
  validateRequest("deleteFile", "queryOrBody"),
  excelController.deleteFile
);

router.delete(
  "/delete-sheet",
  auditLogger.middleware(),
  validateRequest("deleteSheet", "body"),
  excelController.deleteSheet
);

module.exports = router;
