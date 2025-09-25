const express = require("express");
const cors = require("cors");
const app = express();
// Routers (make sure these paths match your project structure)
const healthRoutes = require("./routes/health");
const excelRoutes = require("./routes/excel");
const renameRoutes = require("./routes/rename");
// Basic middleware
app.use(cors());
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

// Root endpoint
app.get("/", (req, res) => {
  res.json({
    service: "Excel GPT Middleware",
    version: "1.0.0",
    status: "running",
    environment: "serverless",
    timestamp: new Date().toISOString(),
    endpoints: {
      health: "/health",
      functions: "/api/excel/functions",
      documentation: "/api/docs",
    },
  });
});

// Health check
app.use("/health", healthRoutes);
// Excel operations
app.use("/api/excel", excelRoutes);
// Rename operations
app.use("/api/excel", renameRoutes);

// app.get("/health", (req, res) => {
//   res.json({
//     success: true,
//     data: {
//       status: "ok",
//       time: new Date().toISOString(),
//       service: "Excel GPT Middleware",
//       environment: "serverless",
//     },
//   });
// });

// Excel functions data
const excelFunctions = {
  arithmetic: [
    "SUM",
    "AVERAGE",
    "COUNT",
    "MIN",
    "MAX",
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "ABS",
    "POWER",
    "SQRT",
  ],
  lookup: [
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "XLOOKUP",
    "FILTER",
    "SORT",
    "UNIQUE",
  ],
  text: [
    "CONCATENATE",
    "CONCAT",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "UPPER",
    "LOWER",
    "PROPER",
    "TRIM",
  ],
  logical: [
    "IF",
    "IFS",
    "AND",
    "OR",
    "NOT",
    "TRUE",
    "FALSE",
    "IFERROR",
    "IFNA",
  ],
  date: [
    "TODAY",
    "NOW",
    "DATE",
    "TIME",
    "YEAR",
    "MONTH",
    "DAY",
    "WEEKDAY",
    "DATEDIF",
  ],
  financial: ["PMT", "PV", "FV", "RATE", "NPER", "NPV", "IRR", "XIRR"],
  statistical: [
    "STDEV",
    "STDEVP",
    "VAR",
    "VARP",
    "CORREL",
    "MEDIAN",
    "MODE",
    "PERCENTILE",
  ],
  engineering: [
    "CONVERT",
    "BIN2DEC",
    "BIN2HEX",
    "DEC2BIN",
    "DEC2HEX",
    "HEX2BIN",
    "HEX2DEC",
  ],
  information: [
    "ISBLANK",
    "ISERROR",
    "ISNA",
    "ISNUMBER",
    "ISTEXT",
    "TYPE",
    "CELL",
  ],
};

// Excel functions endpoint
app.get("/api/excel/functions", (req, res) => {
  const { category } = req.query;

  if (category && excelFunctions[category.toLowerCase()]) {
    res.json({
      status: "success",
      data: {
        category: category.toLowerCase(),
        functions: excelFunctions[category.toLowerCase()],
      },
    });
  } else {
    res.json({
      status: "success",
      data: {
        categories: Object.keys(excelFunctions),
        functions: excelFunctions,
        totalFunctions: Object.values(excelFunctions).flat().length,
      },
    });
  }
});

// API documentation
app.get("/api/docs", (req, res) => {
  res.json({
    service: "Excel GPT Middleware API",
    version: "1.0.0",
    environment: "serverless",
    description:
      "Comprehensive Excel operations via Microsoft Graph API with universal name-based resolution",

    features: [
      "Universal name-to-ID resolution (use driveName, fileName, sheetName instead of IDs)",
      "Deep recursive file search through all folders and subfolders",
      "Intelligent find & replace with scope-aware operations",
      "Comprehensive Excel engine with 200+ formulas",
      "Cell formatting: colors, fonts, borders, styles",
      "Advanced features: pivot tables, conditional formatting, data validation",
      "Rename functionality for files, folders, and sheets",
      "Formula validation and syntax checking",
      "Backward compatibility with existing ID-based calls",
    ],

    endpoints: {
      // Core Excel Operations
      read: {
        method: "POST",
        path: "/api/excel/read",
        description: "Read data from Excel range with name-based resolution",
        body: ["driveName", "fileName", "sheetName", "range"],
        example: {
          driveName: "Documents",
          fileName: "Budget.xlsx",
          sheetName: "Summary",
          range: "A1:D10",
        },
      },
      write: {
        method: "POST",
        path: "/api/excel/write",
        description: "Write data to Excel range",
        body: ["driveName", "fileName", "sheetName", "range", "values"],
      },

      // Find & Replace Operations
      findReplace: {
        method: "POST",
        path: "/api/excel/find-replace",
        description:
          "Find and replace text with intelligent scoping and preview",
        body: ["driveName", "fileName", "searchTerm", "replaceTerm", "scope"],
        scopes: ["header_only", "specific_range", "entire_sheet", "all_sheets"],
      },
      searchText: {
        method: "POST",
        path: "/api/excel/search-text",
        description: "Search for text without replacement",
        body: ["driveName", "fileName", "searchTerm", "scope"],
      },

      // Excel Engine Operations
      format: {
        method: "POST",
        path: "/api/excel/format",
        description:
          "Apply comprehensive Excel formatting, formulas, and advanced features",
        body: ["driveName", "fileName", "sheetName", "operations"],
        operationTypes: [
          "highlight",
          "textStyle",
          "borders",
          "resizeColumn",
          "mergeCells",
          "formula",
          "conditionalFormatting",
          "pivotTable",
        ],
      },
      validateFormula: {
        method: "POST",
        path: "/api/excel/validate-formula",
        description: "Validate Excel formula syntax before insertion",
        body: ["driveName", "fileName", "sheetName", "formula"],
      },

      // Information Endpoints
      functions: {
        method: "GET",
        path: "/api/excel/functions",
        description:
          "Get available Excel functions by category (no auth required)",
        parameters: ["category (optional)"],
        categories: Object.keys(excelFunctions),
      },
      cellInfo: {
        method: "GET",
        path: "/api/excel/cell-info",
        description:
          "Get comprehensive cell information (value, formula, formatting)",
        parameters: ["driveName", "fileName", "sheetName", "cellAddress"],
      },
      worksheetInfo: {
        method: "GET",
        path: "/api/excel/worksheet-info",
        description: "Get worksheet structure and metadata",
        parameters: ["driveName", "fileName", "sheetName (optional)"],
      },

      // Rename Operations
      renameFile: {
        method: "POST",
        path: "/api/excel/rename-file",
        description: "Rename Excel files with intelligent duplicate handling",
        body: ["driveName", "fileName", "newName"],
      },
      renameSheet: {
        method: "POST",
        path: "/api/excel/rename-sheet",
        description: "Rename Excel worksheets",
        body: ["driveName", "fileName", "sheetName", "newName"],
      },
      renameFolder: {
        method: "POST",
        path: "/api/excel/rename-folder",
        description: "Rename folders at any hierarchy level",
        body: ["driveName", "folderName", "newName"],
      },
    },

    authentication: {
      type: "Azure AD Client Credentials",
      description:
        "Most endpoints require authentication via Azure AD service principal",
      publicEndpoints: ["/health", "/", "/api/docs", "/api/excel/functions"],
      authRequiredEndpoints: "All other endpoints require valid access token",
    },

    nameResolution: {
      description: "Use natural names instead of complex IDs",
      examples: {
        legacy: {
          driveId: "b!abc123",
          itemId: "def456",
          worksheetId: "ghi789",
        },
        modern: {
          driveName: "Documents",
          fileName: "Budget.xlsx",
          sheetName: "Summary",
        },
        hybrid: {
          driveId: "b!abc123",
          fileName: "Budget.xlsx",
          sheetName: "Summary",
        },
      },
      pathResolution: {
        fullPath: "/Projects/2024/Budget.xlsx",
        itemPath: "/Folder1/Subfolder/file.xlsx (for duplicate disambiguation)",
      },
    },
  });
});

// Mock protected endpoints (return 401 for auth required)
const protectedEndpoints = [
  "/api/excel/read",
  "/api/excel/write",
  "/api/excel/find-replace",
  "/api/excel/search-text",
  "/api/excel/format",
  "/api/excel/validate-formula",
  "/api/excel/cell-info",
  "/api/excel/worksheet-info",
  "/api/excel/analyze-scope",
  "/api/excel/rename-file",
  "/api/excel/rename-folder",
  "/api/excel/rename-sheet",
];

protectedEndpoints.forEach((endpoint) => {
  app.all(endpoint, (req, res) => {
    res.status(401).json({
      status: "error",
      error: {
        code: 401,
        message:
          "Authentication required. Please provide a valid Azure AD access token.",
        endpoint: endpoint,
        method: req.method,
        hint: "This endpoint requires Azure AD authentication with appropriate SharePoint permissions",
      },
      timestamp: new Date().toISOString(),
    });
  });
});

// 404 handler
app.use("*", (req, res) => {
  res.status(404).json({
    status: "error",
    error: {
      code: 404,
      message: `Endpoint not found: ${req.method} ${req.originalUrl}`,
    },
    timestamp: new Date().toISOString(),
  });
});

// Export for Vercel
module.exports = app;
