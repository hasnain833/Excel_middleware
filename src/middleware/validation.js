const Joi = require("joi");
const logger = require("../config/logger");

// Common validation schemas
const schemas = {
  range: Joi.string()
    .pattern(
      /^(?:[^!\n\r]+!)?(?:[A-Z]+\d+:[A-Z]+\d+|[A-Z]+\d+|[A-Z]+:[A-Z]+|\d+:\d+)$/
    )
    .required(),
  workbookId: Joi.string().min(1).max(255),
  worksheetId: Joi.string().min(1).max(255),
  driveId: Joi.string().min(1).max(255),
  driveName: Joi.string().min(1).max(255),
  itemName: Joi.string().min(1).max(255),
  worksheetName: Joi.string().min(1).max(255),
  values: Joi.array().items(Joi.array()).min(1),
  userId: Joi.string().email().optional(),
  xlsxFileName: Joi.string()
    .pattern(/\.xlsx$/i)
    .message("fileName must end with .xlsx"),
  parentPath: Joi.string().pattern(/^\//).optional(),
  position: Joi.number().integer().min(0).optional(),
  itemPath: Joi.string().pattern(/^\//).optional(),
  force: Joi.boolean().optional(),
};

// Helper to validate .xlsx filenames
schemas.xlsxFileName = Joi.string()
  .pattern(/\.xlsx$/i)
  .message("fileName must end with .xlsx");
schemas.parentPath = Joi.string().pattern(/^\//).optional();
schemas.position = Joi.number().integer().min(0).optional();
schemas.itemPath = Joi.string().pattern(/^\//).optional();

// Names-only base object (IDs are not accepted anymore)
const namesOnlyBase = Joi.object({
  driveName: schemas.driveName.required(),
  itemName: schemas.itemName.optional(),
  itemPath: schemas.itemPath.optional(),
  fullPath: Joi.string().pattern(/^\//).optional(),
}).or("itemName", "fullPath");

const requestSchemas = {
  readRange: namesOnlyBase.concat(
    Joi.object({
      worksheetName: schemas.worksheetName.optional(),
      range: Joi.string().min(1).optional(),
    })
  ),

  writeRange: namesOnlyBase.concat(
    Joi.object({
      worksheetName: schemas.worksheetName.optional(),
      range: Joi.string().min(1).optional(),
      values: schemas.values.required(),
    })
  ),

  getWorksheets: Joi.object({
    driveName: schemas.driveName.required(),
    itemName: schemas.itemName.required(),
    itemPath: schemas.itemPath.optional(),
  }),

  searchFiles: Joi.object({
    driveName: schemas.driveName.required(),
    fileName: Joi.string().min(1).required(),
  }),

  createFile: Joi.object({
    driveName: schemas.driveName.required(),
    parentPath: schemas.parentPath,
    fileName: schemas.xlsxFileName.required(),
    template: Joi.string().valid("blank").default("blank"),
  }),

  createSheet: Joi.object({
    driveName: schemas.driveName.required(),
    itemName: schemas.itemName.required(),
    itemPath: schemas.itemPath.optional(),
    sheetName: schemas.worksheetName.required(),
    position: schemas.position,
  }),

  deleteFile: Joi.object({
    driveId: schemas.driveId.optional(),
    driveName: schemas.driveName.optional(),
    itemId: schemas.itemId.optional(),
    itemName: schemas.itemName.optional(),
    itemPath: schemas.itemPath.optional(),
    force: schemas.force,
  })
    .or("driveId", "driveName")
    .or("itemId", "itemName"),

  deleteSheet: Joi.object({
    driveName: schemas.driveName.required(),
    itemName: schemas.itemName.required(),
    itemPath: schemas.itemPath.optional(),
    sheetName: schemas.worksheetName.required(),
  }),

  analyzeScope: Joi.object({
    driveName: schemas.driveName.required(),
    itemName: schemas.itemName.required(),
    itemPath: schemas.itemPath.optional(),
  }),

  findReplace: Joi.object({
    driveName: schemas.driveName.required(),
    itemName: schemas.itemName.required(),
    itemPath: schemas.itemPath.optional(),
    searchTerm: Joi.string().min(1).required(),
    replaceTerm: Joi.string().allow("").optional(),
    // New optional workflow fields (backward compatible)
    mode: Joi.string().valid("preview", "apply").optional(),
    strategy: Joi.string().valid("text", "entityName").default("text"),
    sheetScope: Joi.string()
      .optional()
      .description("ALL or a specific sheet name"),
    selection: Joi.array().items(Joi.string()).optional(),
    selectAll: Joi.boolean().optional(),
    scope: Joi.string()
      .valid("header_only", "specific_range", "entire_sheet", "all_sheets")
      .default("entire_sheet"),
    rangeSpec: Joi.string().optional(),
    highlightChanges: Joi.boolean().optional(),
    logChanges: Joi.boolean().optional(),
    confirm: Joi.boolean().optional(),
    previewId: Joi.string().optional(),
  }),

  searchText: Joi.object({
    driveName: schemas.driveName.required(),
    itemName: schemas.itemName.required(),
    itemPath: schemas.itemPath.optional(),
    searchTerm: Joi.string().min(1).required(),
    scope: Joi.string()
      .valid("header_only", "specific_range", "entire_sheet", "all_sheets")
      .default("entire_sheet"),
    rangeSpec: Joi.string().optional(),
  }),

  excelFormat: Joi.object({
    driveName: schemas.driveName.required(),
    itemName: schemas.itemName.required(),
    itemPath: schemas.itemPath.optional(),
    sheetName: schemas.worksheetName.optional(),
    operations: Joi.array().items(Joi.object()).optional(),
    formula: Joi.object({
      expression: Joi.string().min(1).required(),
      targetCell: Joi.string().min(1).required(),
      overwrite: Joi.boolean().optional(),
    }).optional(),
  }),

  // New: rename file with optional driveName and support for selectedItemId
  renameFile: Joi.object({
    driveName: schemas.driveName.optional(),
    itemName: schemas.itemName.optional(),
    itemPath: schemas.itemPath.optional(),
    oldName: Joi.string().min(1).optional(),
    newName: Joi.string().min(1).required(),
    selectedItemId: Joi.string().min(1).optional(),
  }).or("itemName", "selectedItemId"),
};

// Clear data request: clear whole sheet (usedRange) or a specific range
requestSchemas.clearData = Joi.object({
  driveName: schemas.driveName.required(),
  itemName: schemas.itemName.required(),
  itemPath: schemas.itemPath.optional(),
  worksheetName: schemas.worksheetName.optional(),
  sheetName: schemas.worksheetName.optional(),
  range: Joi.string().min(1).optional(),
});

const isValidRange = (range) => {
  const rangeRegex =
    /^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/;
  return rangeRegex.test(range);
};
const validateValuesArray = (values) => {
  if (!Array.isArray(values)) {
    return { valid: false, message: "Values must be an array" };
  }

  if (values.length === 0) {
    return { valid: false, message: "Values array cannot be empty" };
  }

  // Check if all rows are arrays and have consistent length
  const firstRowLength = Array.isArray(values[0]) ? values[0].length : 1;

  for (let i = 0; i < values.length; i++) {
    if (!Array.isArray(values[i])) {
      return { valid: false, message: `Row ${i} must be an array` };
    }

    if (values[i].length !== firstRowLength) {
      return {
        valid: false,
        message: `All rows must have the same length. Row ${i} has ${values[i].length} columns, expected ${firstRowLength}`,
      };
    }
  }

  return { valid: true, rows: values.length, columns: firstRowLength };
};

const validateRangeValuesCompatibility = (req, res, next) => {
  const { range, values } = req.body;

  if (!range || !values) {
    return next();
  }

  try {
    const rangeParts = range.split(":");
    if (rangeParts.length === 2) {
      const startCell = rangeParts[0];
      const endCell = rangeParts[1];
      const startColMatch = startCell.match(/[A-Z]+/);
      const startRowMatch = startCell.match(/\d+/);
      const endColMatch = endCell.match(/[A-Z]+/);
      const endRowMatch = endCell.match(/\d+/);

      if (startColMatch && startRowMatch && endColMatch && endRowMatch) {
        const expectedRows =
          parseInt(endRowMatch[0]) - parseInt(startRowMatch[0]) + 1;
        const expectedCols =
          columnToNumber(endColMatch[0]) - columnToNumber(startColMatch[0]) + 1;

        if (
          values.length !== expectedRows ||
          (values[0] && values[0].length !== expectedCols)
        ) {
          return res.status(400).json({
            error: "Range-values mismatch",
            message: `Range expects ${expectedRows}x${expectedCols}, got ${
              values.length
            }x${values[0]?.length || 0}`,
            timestamp: new Date().toISOString(),
          });
        }
      }
    }

    next();
  } catch (error) {
    // Continue - let Graph API handle validation
    next();
  }
};

function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - "A".charCodeAt(0) + 1);
  }
  return result;
}

const sanitizeInput = (input) => {
  if (typeof input !== "string") {
    return input;
  }
  return input.replace(/[<>"';]/g, "");
};

const sanitizeRequest = (req, res, next) => {
  const sanitizeObject = (obj) => {
    if (typeof obj === "string") {
      return sanitizeInput(obj);
    } else if (Array.isArray(obj)) {
      return obj.map(sanitizeObject);
    } else if (obj && typeof obj === "object") {
      const sanitized = {};
      for (const [key, value] of Object.entries(obj)) {
        sanitized[key] = sanitizeObject(value);
      }
      return sanitized;
    }
    return obj;
  };
  req.body = sanitizeObject(req.body);
  req.query = sanitizeObject(req.query);
  req.params = sanitizeObject(req.params);
  next();
};

const validateRequest = (schemaName, source = "body") => {
  return (req, res, next) => {
    const schema = requestSchemas[schemaName];
    if (!schema) {
      logger.error(`Validation schema '${schemaName}' not found`);
      return res.status(500).json({
        error: "Validation configuration error",
        message: `Schema '${schemaName}' is not defined`,
        timestamp: new Date().toISOString(),
      });
    }

    const dataToValidate =
      source === "queryOrBody"
        ? { ...(req.query || {}), ...(req.body || {}) }
        : req[source];
    const { error, value } = schema.validate(dataToValidate, {
      abortEarly: false,
      stripUnknown: true,
    });

    if (error) {
      const errorDetails = error.details.map((detail) => ({
        field: detail.path.join("."),
        message: detail.message,
        value: detail.context?.value,
      }));
      logger.warn("Request validation failed", {
        schema: schemaName,
        errors: errorDetails,
      });
      return res.status(400).json({
        status: "error",
        error: {
          code: 400,
          message: "Request data is invalid",
          details: errorDetails,
        },
        timestamp: new Date().toISOString(),
      });
    }

    if (source === "queryOrBody") {
      req.body = value; // normalize merged payload into body for controllers
    } else {
      req[source] = value;
    }
    next();
  };
};

module.exports = {
  validateRequest,
  validateRangeValuesCompatibility,
  sanitizeRequest,
  isValidRange,
  validateValuesArray,
  schemas,
  requestSchemas,
};
