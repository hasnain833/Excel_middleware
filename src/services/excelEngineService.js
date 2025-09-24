/**
 * Excel Engine Service
 * Provides comprehensive Excel functionality including formulas, formatting, and advanced features
 * Acts as a true Excel engine via Microsoft Graph API
 */

const { Client } = require('@microsoft/microsoft-graph-client');
const logger = require('../config/logger');
const resolverService = require('./resolverService');
const auditService = require('./auditService');
const { AppError } = require('../middleware/errorHandler');

class ExcelEngineService {
  constructor() {
    // Cache for workbook sessions and formula validation
    this.sessionCache = new Map();
    this.formulaCache = new Map();
    this.ttlMs = 5 * 60 * 1000; // 5 minutes TTL
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }

  /**
   * Apply comprehensive formatting and formulas to Excel worksheet
   */
  async applyFormatting(accessToken, driveId, itemId, sheetName, operations, auditContext) {
    if (!driveId || !itemId || !operations || !Array.isArray(operations)) {
      throw new AppError('driveId, itemId, and operations array are required', 400);
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const results = [];
      const errors = [];

      // Get worksheet ID
      const worksheetId = sheetName 
        ? await resolverService.resolveWorksheetIdByName(accessToken, driveId, itemId, sheetName)
        : null;

      // Process operations in batches for performance
      const batches = this.groupOperationsByType(operations);
      
      for (const [operationType, ops] of batches.entries()) {
        try {
          const batchResults = await this.processBatchOperations(
            graphClient, 
            driveId, 
            itemId, 
            worksheetId, 
            sheetName,
            operationType, 
            ops
          );
          results.push(...batchResults);
        } catch (batchErr) {
          logger.error(`Failed to process ${operationType} operations`, { error: batchErr.message });
          errors.push({
            operationType,
            error: batchErr.message,
            operations: ops.length
          });
        }
      }

      // Log the operation
      auditService.logSystemEvent({
        event: 'EXCEL_FORMATTING_APPLIED',
        details: {
          driveId,
          itemId,
          sheetName,
          operationsCount: operations.length,
          successfulOperations: results.length,
          failedOperations: errors.length,
          requestedBy: auditContext.user
        }
      });

      return {
        results,
        errors,
        summary: {
          total: operations.length,
          successful: results.length,
          failed: errors.length
        }
      };

    } catch (err) {
      logger.error('Failed to apply Excel formatting', {
        driveId,
        itemId,
        sheetName,
        error: err.message
      });
      throw err;
    }
  }

  /**
   * Group operations by type for batch processing
   */
  groupOperationsByType(operations) {
    const groups = new Map();
    
    operations.forEach(op => {
      const type = op.type || 'unknown';
      if (!groups.has(type)) {
        groups.set(type, []);
      }
      groups.get(type).push(op);
    });

    return groups;
  }

  /**
   * Process batch operations by type
   */
  async processBatchOperations(graphClient, driveId, itemId, worksheetId, sheetName, operationType, operations) {
    const results = [];

    switch (operationType) {
      case 'highlight':
      case 'backgroundColor':
        const highlightResults = await this.applyBackgroundColor(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...highlightResults);
        break;

      case 'textStyle':
      case 'font':
        const fontResults = await this.applyTextFormatting(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...fontResults);
        break;

      case 'borders':
        const borderResults = await this.applyBorders(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...borderResults);
        break;

      case 'resizeColumn':
        const columnResults = await this.resizeColumns(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...columnResults);
        break;

      case 'resizeRow':
        const rowResults = await this.resizeRows(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...rowResults);
        break;

      case 'mergeCells':
        const mergeResults = await this.mergeCells(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...mergeResults);
        break;

      case 'unmergeCells':
        const unmergeResults = await this.unmergeCells(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...unmergeResults);
        break;

      case 'formula':
        const formulaResults = await this.insertFormulas(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...formulaResults);
        break;

      case 'conditionalFormatting':
        const conditionalResults = await this.applyConditionalFormatting(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...conditionalResults);
        break;

      case 'pivotTable':
        const pivotResults = await this.createPivotTable(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...pivotResults);
        break;

      case 'namedRange':
        const namedRangeResults = await this.createNamedRange(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...namedRangeResults);
        break;

      case 'dataValidation':
        const validationResults = await this.applyDataValidation(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...validationResults);
        break;

      case 'sort':
        const sortResults = await this.sortRange(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...sortResults);
        break;

      case 'filter':
        const filterResults = await this.applyFilter(graphClient, driveId, itemId, worksheetId, sheetName, operations);
        results.push(...filterResults);
        break;

      default:
        logger.warn(`Unknown operation type: ${operationType}`);
        break;
    }

    return results;
  }

  /**
   * Apply background color highlighting
   */
  async applyBackgroundColor(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, color } = op;
        if (!range || !color) {
          throw new Error('range and color are required for highlight operation');
        }

        const colorCode = this.normalizeColor(color);
        const rangeAddress = this.normalizeRange(range);

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/format/fill`)
          .patch({
            color: colorCode
          });

        results.push({
          type: 'highlight',
          range: rangeAddress,
          color: colorCode,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to apply background color', { operation: op, error: err.message });
        results.push({
          type: 'highlight',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Apply text formatting (bold, italic, underline, font size, color)
   */
  async applyTextFormatting(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, style, fontSize, fontColor, fontName } = op;
        if (!range) {
          throw new Error('range is required for text formatting operation');
        }

        const rangeAddress = this.normalizeRange(range);
        const formatUpdates = {};

        // Handle text styles
        if (style && Array.isArray(style)) {
          if (style.includes('bold')) formatUpdates.bold = true;
          if (style.includes('italic')) formatUpdates.italic = true;
          if (style.includes('underline')) formatUpdates.underline = 'Single';
        }

        // Handle font properties
        if (fontSize) formatUpdates.size = fontSize;
        if (fontColor) formatUpdates.color = this.normalizeColor(fontColor);
        if (fontName) formatUpdates.name = fontName;

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/format/font`)
          .patch(formatUpdates);

        results.push({
          type: 'textStyle',
          range: rangeAddress,
          formatting: formatUpdates,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to apply text formatting', { operation: op, error: err.message });
        results.push({
          type: 'textStyle',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Apply borders to cells
   */
  async applyBorders(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, borderStyle, borderColor, sides } = op;
        if (!range) {
          throw new Error('range is required for border operation');
        }

        const rangeAddress = this.normalizeRange(range);
        const borderUpdates = {};

        const style = borderStyle || 'Continuous';
        const color = this.normalizeColor(borderColor || '#000000');

        // Apply borders to specified sides
        const borderSides = sides || ['top', 'bottom', 'left', 'right'];
        borderSides.forEach(side => {
          borderUpdates[side] = {
            style: style,
            color: color
          };
        });

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/format/borders`)
          .patch(borderUpdates);

        results.push({
          type: 'borders',
          range: rangeAddress,
          borders: borderUpdates,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to apply borders', { operation: op, error: err.message });
        results.push({
          type: 'borders',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Resize columns
   */
  async resizeColumns(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { column, width, autoFit } = op;
        if (!column) {
          throw new Error('column is required for resize operation');
        }

        if (autoFit) {
          // Auto-fit column
          await graphClient
            .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/columns('${column}')/resizeToFit`)
            .post({});
        } else if (width) {
          // Set specific width
          await graphClient
            .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/columns('${column}')`)
            .patch({
              columnWidth: width
            });
        }

        results.push({
          type: 'resizeColumn',
          column: column,
          width: autoFit ? 'auto-fit' : width,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to resize column', { operation: op, error: err.message });
        results.push({
          type: 'resizeColumn',
          column: op.column,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Resize rows
   */
  async resizeRows(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { row, height, autoFit } = op;
        if (!row) {
          throw new Error('row is required for resize operation');
        }

        if (autoFit) {
          // Auto-fit row
          await graphClient
            .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/rows('${row}')/resizeToFit`)
            .post({});
        } else if (height) {
          // Set specific height
          await graphClient
            .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/rows('${row}')`)
            .patch({
              rowHeight: height
            });
        }

        results.push({
          type: 'resizeRow',
          row: row,
          height: autoFit ? 'auto-fit' : height,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to resize row', { operation: op, error: err.message });
        results.push({
          type: 'resizeRow',
          row: op.row,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Merge cells
   */
  async mergeCells(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, across } = op;
        if (!range) {
          throw new Error('range is required for merge operation');
        }

        const rangeAddress = this.normalizeRange(range);

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/merge`)
          .post({
            across: across || false
          });

        results.push({
          type: 'mergeCells',
          range: rangeAddress,
          across: across || false,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to merge cells', { operation: op, error: err.message });
        results.push({
          type: 'mergeCells',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Unmerge cells
   */
  async unmergeCells(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range } = op;
        if (!range) {
          throw new Error('range is required for unmerge operation');
        }

        const rangeAddress = this.normalizeRange(range);

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/unmerge`)
          .post({});

        results.push({
          type: 'unmergeCells',
          range: rangeAddress,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to unmerge cells', { operation: op, error: err.message });
        results.push({
          type: 'unmergeCells',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Insert formulas with validation
   */
  async insertFormulas(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { expression, targetCell, overwrite } = op;
        if (!expression || !targetCell) {
          throw new Error('expression and targetCell are required for formula operation');
        }

        // Validate formula syntax
        const isValid = await this.validateFormula(graphClient, driveId, itemId, worksheetId, sheetName, expression);
        if (!isValid.valid) {
          throw new Error(`Invalid formula: ${isValid.error}`);
        }

        // Check if cell has existing data
        if (!overwrite) {
          const existingData = await this.getCellValue(graphClient, driveId, itemId, worksheetId, sheetName, targetCell);
          if (existingData && existingData.value !== null && existingData.value !== '') {
            throw new Error(`Cell ${targetCell} contains data. Set overwrite: true to replace.`);
          }
        }

        // Insert formula
        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${targetCell}')`)
          .patch({
            formulas: [[expression]]
          });

        results.push({
          type: 'formula',
          cell: targetCell,
          expression: expression,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to insert formula', { operation: op, error: err.message });
        results.push({
          type: 'formula',
          cell: op.targetCell,
          expression: op.expression,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Validate formula syntax
   */
  async validateFormula(graphClient, driveId, itemId, worksheetId, sheetName, formula) {
    try {
      // Use a temporary cell to test formula validity
      const testCell = 'ZZ1000'; // Use a cell unlikely to contain data
      
      await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${testCell}')`)
        .patch({
          formulas: [[formula]]
        });

      // If no error, formula is valid - clear the test cell
      await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${testCell}')`)
        .patch({
          values: [['']]
        });

      return { valid: true };

    } catch (err) {
      return { 
        valid: false, 
        error: err.message || 'Invalid formula syntax'
      };
    }
  }

  /**
   * Get cell value
   */
  async getCellValue(graphClient, driveId, itemId, worksheetId, sheetName, cellAddress) {
    try {
      const response = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${cellAddress}')`)
        .get();

      return {
        value: response.values?.[0]?.[0] || null,
        formula: response.formulas?.[0]?.[0] || null
      };
    } catch (err) {
      return { value: null, formula: null };
    }
  }

  /**
   * Apply conditional formatting
   */
  async applyConditionalFormatting(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, rule, format } = op;
        if (!range || !rule) {
          throw new Error('range and rule are required for conditional formatting');
        }

        const rangeAddress = this.normalizeRange(range);
        
        const conditionalFormat = {
          type: rule.type || 'cellValue',
          cellValue: rule.condition ? {
            formula1: rule.condition.value,
            operator: rule.condition.operator || 'greaterThan'
          } : undefined,
          format: {
            fill: format?.backgroundColor ? { color: this.normalizeColor(format.backgroundColor) } : undefined,
            font: format?.fontColor ? { color: this.normalizeColor(format.fontColor) } : undefined
          }
        };

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/conditionalFormats`)
          .post(conditionalFormat);

        results.push({
          type: 'conditionalFormatting',
          range: rangeAddress,
          rule: rule,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to apply conditional formatting', { operation: op, error: err.message });
        results.push({
          type: 'conditionalFormatting',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Create pivot table
   */
  async createPivotTable(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { sourceRange, destinationRange, rows, columns, values, filters } = op;
        if (!sourceRange || !destinationRange) {
          throw new Error('sourceRange and destinationRange are required for pivot table');
        }

        const pivotTable = {
          source: {
            range: this.normalizeRange(sourceRange)
          },
          destination: {
            range: this.normalizeRange(destinationRange)
          },
          rows: rows || [],
          columns: columns || [],
          data: values || [],
          filters: filters || []
        };

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/pivotTables`)
          .post(pivotTable);

        results.push({
          type: 'pivotTable',
          sourceRange: sourceRange,
          destinationRange: destinationRange,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to create pivot table', { operation: op, error: err.message });
        results.push({
          type: 'pivotTable',
          sourceRange: op.sourceRange,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Create named range
   */
  async createNamedRange(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { name, range, comment } = op;
        if (!name || !range) {
          throw new Error('name and range are required for named range');
        }

        const namedRange = {
          name: name,
          reference: `${sheetName}!${this.normalizeRange(range)}`,
          comment: comment || ''
        };

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/names`)
          .post(namedRange);

        results.push({
          type: 'namedRange',
          name: name,
          range: range,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to create named range', { operation: op, error: err.message });
        results.push({
          type: 'namedRange',
          name: op.name,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Apply data validation
   */
  async applyDataValidation(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, rule } = op;
        if (!range || !rule) {
          throw new Error('range and rule are required for data validation');
        }

        const rangeAddress = this.normalizeRange(range);
        const validation = {
          type: rule.type || 'list',
          criteria: rule.criteria || {},
          errorAlert: rule.errorAlert || {},
          inputMessage: rule.inputMessage || {}
        };

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/dataValidation`)
          .patch(validation);

        results.push({
          type: 'dataValidation',
          range: rangeAddress,
          rule: rule,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to apply data validation', { operation: op, error: err.message });
        results.push({
          type: 'dataValidation',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Sort range
   */
  async sortRange(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, sortFields, hasHeaders } = op;
        if (!range || !sortFields) {
          throw new Error('range and sortFields are required for sort operation');
        }

        const rangeAddress = this.normalizeRange(range);
        const sortData = {
          fields: sortFields,
          hasHeaders: hasHeaders !== false
        };

        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/sort/apply`)
          .post(sortData);

        results.push({
          type: 'sort',
          range: rangeAddress,
          sortFields: sortFields,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to sort range', { operation: op, error: err.message });
        results.push({
          type: 'sort',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Apply filter
   */
  async applyFilter(graphClient, driveId, itemId, worksheetId, sheetName, operations) {
    const results = [];

    for (const op of operations) {
      try {
        const { range, criteria } = op;
        if (!range) {
          throw new Error('range is required for filter operation');
        }

        const rangeAddress = this.normalizeRange(range);

        if (criteria) {
          // Apply specific filter criteria
          await graphClient
            .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${rangeAddress}')/filter/apply`)
            .post({
              criteria: criteria
            });
        } else {
          // Apply auto filter
          await graphClient
            .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/autoFilter/apply`)
            .post({
              range: rangeAddress
            });
        }

        results.push({
          type: 'filter',
          range: rangeAddress,
          criteria: criteria,
          status: 'success'
        });

      } catch (err) {
        logger.error('Failed to apply filter', { operation: op, error: err.message });
        results.push({
          type: 'filter',
          range: op.range,
          status: 'error',
          error: err.message
        });
      }
    }

    return results;
  }

  /**
   * Normalize color values to hex format
   */
  normalizeColor(color) {
    if (!color) return '#000000';
    
    // Color name to hex mapping
    const colorMap = {
      'red': '#FF0000',
      'green': '#00FF00',
      'blue': '#0000FF',
      'yellow': '#FFFF00',
      'orange': '#FFA500',
      'purple': '#800080',
      'pink': '#FFC0CB',
      'brown': '#A52A2A',
      'gray': '#808080',
      'grey': '#808080',
      'black': '#000000',
      'white': '#FFFFFF'
    };

    const lowerColor = color.toLowerCase();
    if (colorMap[lowerColor]) {
      return colorMap[lowerColor];
    }

    // If already hex format, return as is
    if (color.startsWith('#')) {
      return color;
    }

    // Default to black if unrecognized
    return '#000000';
  }

  /**
   * Normalize range addresses
   */
  normalizeRange(range) {
    if (!range) return 'A1';
    
    // Handle single cell references
    if (/^[A-Z]+\d+$/.test(range)) {
      return range;
    }

    // Handle range references
    if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(range)) {
      return range;
    }

    // Default fallback
    return range;
  }
}

module.exports = new ExcelEngineService();
