/**
 * Request Validation Middleware
 * Validates incoming requests using Joi schemas
 */

const Joi = require('joi');
const logger = require('../config/logger');

// Common validation schemas
const schemas = {
    // Range validation schema (supports optional Sheet! prefix)
    // Examples: "A1:B2", "Sheet1!A1:D10"
    range: Joi.string()
        .pattern(/^(?:[^!\n\r]+!)?(?:[A-Z]+\d+:[A-Z]+\d+|[A-Z]+\d+|[A-Z]+:[A-Z]+|\d+:\d+)$/)
        .required(),
    
    // Workbook ID validation
    workbookId: Joi.string().min(1).max(255),
    
    // Worksheet ID validation
    worksheetId: Joi.string().min(1).max(255),
    
    // Drive ID validation
    driveId: Joi.string().min(1).max(255),
    
    // Drive Name validation
    driveName: Joi.string().min(1).max(255),
    
    // Item Name (file name) validation
    itemName: Joi.string().min(1).max(255),
    
    // Worksheet Name validation
    worksheetName: Joi.string().min(1).max(255),
    
    // Table name validation
    tableName: Joi.string().min(1).max(255),
    
    // Values array validation (2D array)
    values: Joi.array().items(Joi.array()).min(1),
    
    // Single row validation
    rows: Joi.array().items(Joi.array()).min(1),
    
    // User ID validation
    userId: Joi.string().email().optional()
};

// Helper to require either IDs or names
const idOrName = Joi.alternatives().try(
    Joi.object({ driveId: schemas.driveId.required(), itemId: schemas.workbookId.required() }),
    Joi.object({ driveName: schemas.driveName.required(), itemName: schemas.itemName.required() })
);

// Request validation schemas
const requestSchemas = {
    // Read range request
    readRange: idOrName.concat(Joi.object({
        // worksheet can be provided via worksheetId or worksheetName or inferred from range prefix
        worksheetId: schemas.worksheetId.optional(),
        worksheetName: schemas.worksheetName.optional(),
        range: schemas.range.required()
    })),
    
    // Write range request
    writeRange: idOrName.concat(Joi.object({
        worksheetId: schemas.worksheetId.optional(),
        worksheetName: schemas.worksheetName.optional(),
        range: schemas.range.required(),
        values: schemas.values.required()
    })),
    
    // Read table request
    readTable: idOrName.concat(Joi.object({
        worksheetId: schemas.worksheetId.optional(),
        worksheetName: schemas.worksheetName.optional(),
        tableName: schemas.tableName.required()
    })),
    
    // Add table rows request
    addTableRows: idOrName.concat(Joi.object({
        worksheetId: schemas.worksheetId.optional(),
        worksheetName: schemas.worksheetName.optional(),
        tableName: schemas.tableName.required(),
        rows: schemas.rows.required()
    })),
    
    // Get worksheets request
    getWorksheets: idOrName
};

/**
 * Create validation middleware for a specific schema
 * @param {string} schemaName - Name of the schema to use
 * @param {string} source - Source of data to validate ('body', 'query', 'params')
 * @returns {Function} Express middleware function
 */
const validateRequest = (schemaName, source = 'body') => {
    return (req, res, next) => {
        const schema = requestSchemas[schemaName];
        if (!schema) {
            logger.error(`Validation schema '${schemaName}' not found`);
            return res.status(500).json({
                error: 'Internal server error',
                message: 'Validation configuration error',
                timestamp: new Date().toISOString()
            });
        }

        const dataToValidate = req[source];
        const { error, value } = schema.validate(dataToValidate, {
            abortEarly: false,
            stripUnknown: true
        });

        if (error) {
            const errorDetails = error.details.map(detail => ({
                field: detail.path.join('.'),
                message: detail.message,
                value: detail.context?.value
            }));

            logger.warn('Request validation failed', {
                schema: schemaName,
                errors: errorDetails,
                originalData: dataToValidate
            });

            return res.status(400).json({
                error: 'Validation failed',
                message: 'Request data is invalid',
                details: errorDetails,
                timestamp: new Date().toISOString()
            });
        }

        // Replace the original data with the validated and sanitized data
        req[source] = value;
        next();
    };
};

/**
 * Validate Excel range format
 * @param {string} range - Range string
 * @returns {boolean} True if valid
 */
const isValidRange = (range) => {
    const rangeRegex = /^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/;
    return rangeRegex.test(range);
};

/**
 * Validate 2D array structure for Excel values
 * @param {Array} values - Values array
 * @returns {Object} Validation result
 */
const validateValuesArray = (values) => {
    if (!Array.isArray(values)) {
        return { valid: false, message: 'Values must be an array' };
    }

    if (values.length === 0) {
        return { valid: false, message: 'Values array cannot be empty' };
    }

    // Check if all rows are arrays and have consistent length
    const firstRowLength = Array.isArray(values[0]) ? values[0].length : 1;
    
    for (let i = 0; i < values.length; i++) {
        if (!Array.isArray(values[i])) {
            return { valid: false, message: `Row ${i} must be an array` };
        }
        
        if (values[i].length !== firstRowLength) {
            return { valid: false, message: `All rows must have the same length. Row ${i} has ${values[i].length} columns, expected ${firstRowLength}` };
        }
    }

    return { valid: true, rows: values.length, columns: firstRowLength };
};

/**
 * Middleware to validate range and values compatibility
 */
const validateRangeValuesCompatibility = (req, res, next) => {
    const { range, values } = req.body;
    
    if (!range || !values) {
        return next();
    }

    try {
        const rangeParts = range.split(':');
        if (rangeParts.length === 2) {
            const startCell = rangeParts[0];
            const endCell = rangeParts[1];
            
            const startColMatch = startCell.match(/[A-Z]+/);
            const startRowMatch = startCell.match(/\d+/);
            const endColMatch = endCell.match(/[A-Z]+/);
            const endRowMatch = endCell.match(/\d+/);
            
            if (startColMatch && startRowMatch && endColMatch && endRowMatch) {
                const expectedRows = parseInt(endRowMatch[0]) - parseInt(startRowMatch[0]) + 1;
                const expectedCols = columnToNumber(endColMatch[0]) - columnToNumber(startColMatch[0]) + 1;
                
                if (values.length !== expectedRows || (values[0] && values[0].length !== expectedCols)) {
                    return res.status(400).json({
                        error: 'Range-values mismatch',
                        message: `Range expects ${expectedRows}x${expectedCols}, got ${values.length}x${values[0]?.length || 0}`,
                        timestamp: new Date().toISOString()
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

/**
 * Convert column letter to number (A=1, B=2, etc.)
 * @param {string} column - Column letter(s)
 * @returns {number} Column number
 */
function columnToNumber(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
        result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
}

/**
 * Sanitize user input to prevent injection attacks
 * @param {string} input - Input string
 * @returns {string} Sanitized string
 */
const sanitizeInput = (input) => {
    if (typeof input !== 'string') {
        return input;
    }
    
    // Remove potentially dangerous characters but preserve Excel formulas
    return input.replace(/[<>"';]/g, '');
};

/**
 * Middleware to sanitize all string inputs in request
 */
const sanitizeRequest = (req, res, next) => {
    const sanitizeObject = (obj) => {
        if (typeof obj === 'string') {
            return sanitizeInput(obj);
        } else if (Array.isArray(obj)) {
            return obj.map(sanitizeObject);
        } else if (obj && typeof obj === 'object') {
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

module.exports = {
    validateRequest,
    validateRangeValuesCompatibility,
    sanitizeRequest,
    isValidRange,
    validateValuesArray,
    schemas,
    requestSchemas
};
