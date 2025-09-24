/**
 * Excel Engine Controller
 * Handles HTTP requests for comprehensive Excel formatting, formulas, and advanced features
 */

const excelEngineService = require('../services/excelEngineService');
const resolverService = require('../services/resolverService');
const auditService = require('../services/auditService');
const logger = require('../config/logger');
const { catchAsync } = require('../middleware/errorHandler');
const { AppError } = require('../middleware/errorHandler');

class ExcelEngineController {
  
  /**
   * Apply comprehensive Excel formatting and formulas
   * POST /api/excel/format
   */
  applyFormatting = catchAsync(async (req, res) => {
    const { 
      driveId, 
      itemId, 
      driveName, 
      itemName, 
      itemPath,
      sheetName,
      operations,
      formula
    } = req.body;
    
    const auditContext = auditService.createAuditContext(req);

    // Validate required parameters
    if (!operations && !formula) {
      throw new AppError('Either operations array or formula object is required', 400);
    }

    // Resolve drive and item IDs if names provided
    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;

    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    if (!resolvedItemId && itemName) {
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(req.accessToken, resolvedDriveId, itemName);
      } catch (err) {
        if (err.isMultipleMatches) {
          if (itemPath) {
            resolvedItemId = await resolverService.resolveItemIdByPath(req.accessToken, resolvedDriveId, itemName, itemPath);
          } else {
            return res.status(409).json({
              status: 'multiple_matches',
              message: 'Multiple files found with the same name. Please specify itemPath or select from the list.',
              matches: err.matches.map(match => ({
                id: match.id,
                name: match.name,
                path: match.path,
                parentId: match.parentId
              }))
            });
          }
        } else {
          throw err;
        }
      }
    }

    if (!resolvedDriveId || !resolvedItemId) {
      throw new AppError('Could not resolve drive or item. Please provide valid identifiers.', 400);
    }

    try {
      let allOperations = [];

      // Handle operations array
      if (operations && Array.isArray(operations)) {
        allOperations.push(...operations);
      }

      // Handle single formula object
      if (formula) {
        allOperations.push({
          type: 'formula',
          expression: formula.expression,
          targetCell: formula.targetCell,
          overwrite: formula.overwrite
        });
      }

      // Validate operations
      this.validateOperations(allOperations);

      // Apply formatting and formulas
      const result = await excelEngineService.applyFormatting(
        req.accessToken,
        resolvedDriveId,
        resolvedItemId,
        sheetName,
        allOperations,
        auditContext
      );

      // Log the operation
      auditService.logSystemEvent({
        event: 'EXCEL_ENGINE_OPERATION_COMPLETED',
        details: {
          driveId: resolvedDriveId,
          itemId: resolvedItemId,
          sheetName,
          operationsCount: allOperations.length,
          successful: result.summary.successful,
          failed: result.summary.failed,
          requestedBy: auditContext.user
        }
      });

      res.json({
        status: 'success',
        message: `Successfully completed ${result.summary.successful} of ${result.summary.total} operations`,
        data: {
          summary: result.summary,
          results: result.results,
          errors: result.errors.length > 0 ? result.errors : undefined
        }
      });

    } catch (err) {
      logger.error('Excel engine operation failed', {
        driveId: resolvedDriveId,
        itemId: resolvedItemId,
        sheetName,
        error: err.message
      });
      throw err;
    }
  });

  /**
   * Validate formula syntax
   * POST /api/excel/validate-formula
   */
  validateFormula = catchAsync(async (req, res) => {
    const { 
      driveId, 
      itemId, 
      driveName, 
      itemName, 
      itemPath,
      sheetName,
      formula
    } = req.body;
    
    if (!formula) {
      throw new AppError('formula is required', 400);
    }

    // Resolve drive and item IDs
    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;

    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    if (!resolvedItemId && itemName) {
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(req.accessToken, resolvedDriveId, itemName);
      } catch (err) {
        if (err.isMultipleMatches && itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(req.accessToken, resolvedDriveId, itemName, itemPath);
        } else {
          throw err;
        }
      }
    }

    if (!resolvedDriveId || !resolvedItemId) {
      throw new AppError('Could not resolve drive or item. Please provide valid identifiers.', 400);
    }

    try {
      const graphClient = excelEngineService.createGraphClient(req.accessToken);
      const worksheetId = sheetName 
        ? await resolverService.resolveWorksheetIdByName(req.accessToken, resolvedDriveId, resolvedItemId, sheetName)
        : null;

      const validation = await excelEngineService.validateFormula(
        graphClient,
        resolvedDriveId,
        resolvedItemId,
        worksheetId,
        sheetName,
        formula
      );

      res.json({
        status: 'success',
        data: {
          formula,
          valid: validation.valid,
          error: validation.error || null
        }
      });

    } catch (err) {
      logger.error('Formula validation failed', {
        driveId: resolvedDriveId,
        itemId: resolvedItemId,
        formula,
        error: err.message
      });
      throw err;
    }
  });

  /**
   * Get cell information (value, formula, formatting)
   * GET /api/excel/cell-info
   */
  getCellInfo = catchAsync(async (req, res) => {
    const { 
      driveId, 
      itemId, 
      driveName, 
      itemName, 
      itemPath,
      sheetName,
      cellAddress
    } = req.query;
    
    if (!cellAddress) {
      throw new AppError('cellAddress is required', 400);
    }

    // Resolve drive and item IDs
    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;

    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    if (!resolvedItemId && itemName) {
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(req.accessToken, resolvedDriveId, itemName);
      } catch (err) {
        if (err.isMultipleMatches && itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(req.accessToken, resolvedDriveId, itemName, itemPath);
        } else {
          throw err;
        }
      }
    }

    if (!resolvedDriveId || !resolvedItemId) {
      throw new AppError('Could not resolve drive or item. Please provide valid identifiers.', 400);
    }

    try {
      const graphClient = excelEngineService.createGraphClient(req.accessToken);
      const worksheetId = sheetName 
        ? await resolverService.resolveWorksheetIdByName(req.accessToken, resolvedDriveId, resolvedItemId, sheetName)
        : null;

      // Get comprehensive cell information
      const cellData = await graphClient
        .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${worksheetId || `'${sheetName}'`}/range(address='${cellAddress}')`)
        .select('values,formulas,format,numberFormat')
        .get();

      res.json({
        status: 'success',
        data: {
          cellAddress,
          value: cellData.values?.[0]?.[0] || null,
          formula: cellData.formulas?.[0]?.[0] || null,
          format: cellData.format || null,
          numberFormat: cellData.numberFormat?.[0]?.[0] || null
        }
      });

    } catch (err) {
      logger.error('Failed to get cell info', {
        driveId: resolvedDriveId,
        itemId: resolvedItemId,
        cellAddress,
        error: err.message
      });
      throw err;
    }
  });

  /**
   * Get available Excel functions and formulas
   * GET /api/excel/functions
   */
  getExcelFunctions = catchAsync(async (req, res) => {
    const { category } = req.query;

    // Comprehensive list of Excel functions by category
    const excelFunctions = {
      arithmetic: [
        'SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS',
        'MIN', 'MAX', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS',
        'POWER', 'SQRT', 'MOD', 'QUOTIENT', 'RAND', 'RANDBETWEEN'
      ],
      lookup: [
        'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'CHOOSE',
        'LOOKUP', 'XLOOKUP', 'FILTER', 'SORT', 'UNIQUE'
      ],
      text: [
        'CONCATENATE', 'CONCAT', 'TEXT', 'LEFT', 'RIGHT', 'MID',
        'LEN', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE',
        'UPPER', 'LOWER', 'PROPER', 'TRIM', 'VALUE'
      ],
      logical: [
        'IF', 'IFS', 'AND', 'OR', 'NOT', 'TRUE', 'FALSE',
        'IFERROR', 'IFNA', 'SWITCH'
      ],
      date: [
        'TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY',
        'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'DATEDIF',
        'WORKDAY', 'NETWORKDAYS', 'EOMONTH'
      ],
      financial: [
        'PMT', 'PV', 'FV', 'RATE', 'NPER', 'NPV', 'IRR',
        'XIRR', 'XNPV', 'DB', 'DDB', 'SLN', 'SYD'
      ],
      statistical: [
        'STDEV', 'STDEVP', 'VAR', 'VARP', 'CORREL', 'COVAR',
        'FORECAST', 'TREND', 'LINEST', 'LOGEST', 'PERCENTILE',
        'QUARTILE', 'MEDIAN', 'MODE'
      ],
      engineering: [
        'CONVERT', 'BIN2DEC', 'BIN2HEX', 'DEC2BIN', 'DEC2HEX',
        'HEX2BIN', 'HEX2DEC', 'BITAND', 'BITOR', 'BITXOR'
      ],
      information: [
        'ISBLANK', 'ISERROR', 'ISNA', 'ISNUMBER', 'ISTEXT',
        'TYPE', 'CELL', 'INFO', 'N', 'NA'
      ]
    };

    if (category && excelFunctions[category.toLowerCase()]) {
      res.json({
        status: 'success',
        data: {
          category: category.toLowerCase(),
          functions: excelFunctions[category.toLowerCase()]
        }
      });
    } else {
      res.json({
        status: 'success',
        data: {
          categories: Object.keys(excelFunctions),
          functions: excelFunctions,
          totalFunctions: Object.values(excelFunctions).flat().length
        }
      });
    }
  });

  /**
   * Get worksheet structure and metadata
   * GET /api/excel/worksheet-info
   */
  getWorksheetInfo = catchAsync(async (req, res) => {
    const { 
      driveId, 
      itemId, 
      driveName, 
      itemName, 
      itemPath,
      sheetName
    } = req.query;

    // Resolve drive and item IDs
    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;

    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    if (!resolvedItemId && itemName) {
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(req.accessToken, resolvedDriveId, itemName);
      } catch (err) {
        if (err.isMultipleMatches && itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(req.accessToken, resolvedDriveId, itemName, itemPath);
        } else {
          throw err;
        }
      }
    }

    if (!resolvedDriveId || !resolvedItemId) {
      throw new AppError('Could not resolve drive or item. Please provide valid identifiers.', 400);
    }

    try {
      const graphClient = excelEngineService.createGraphClient(req.accessToken);
      
      if (sheetName) {
        // Get specific worksheet info
        const worksheetId = await resolverService.resolveWorksheetIdByName(req.accessToken, resolvedDriveId, resolvedItemId, sheetName);
        
        const [worksheet, usedRange, tables, pivotTables, charts] = await Promise.all([
          graphClient.api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${worksheetId}`).get(),
          graphClient.api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${worksheetId}/usedRange`).get().catch(() => null),
          graphClient.api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${worksheetId}/tables`).get().catch(() => ({ value: [] })),
          graphClient.api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${worksheetId}/pivotTables`).get().catch(() => ({ value: [] })),
          graphClient.api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${worksheetId}/charts`).get().catch(() => ({ value: [] }))
        ]);

        res.json({
          status: 'success',
          data: {
            worksheet: {
              id: worksheet.id,
              name: worksheet.name,
              position: worksheet.position,
              visibility: worksheet.visibility
            },
            usedRange: usedRange ? {
              address: usedRange.address,
              rowCount: usedRange.rowCount,
              columnCount: usedRange.columnCount
            } : null,
            tables: tables.value?.map(t => ({ id: t.id, name: t.name, range: t.range })) || [],
            pivotTables: pivotTables.value?.map(p => ({ id: p.id, name: p.name })) || [],
            charts: charts.value?.map(c => ({ id: c.id, name: c.name, chartType: c.chartType })) || []
          }
        });
      } else {
        // Get all worksheets info
        const worksheets = await graphClient
          .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets`)
          .get();

        const worksheetInfo = await Promise.all(
          (worksheets.value || []).map(async (ws) => {
            try {
              const usedRange = await graphClient
                .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${ws.id}/usedRange`)
                .get()
                .catch(() => null);

              return {
                id: ws.id,
                name: ws.name,
                position: ws.position,
                visibility: ws.visibility,
                usedRange: usedRange ? {
                  address: usedRange.address,
                  rowCount: usedRange.rowCount,
                  columnCount: usedRange.columnCount
                } : null
              };
            } catch (err) {
              return {
                id: ws.id,
                name: ws.name,
                position: ws.position,
                visibility: ws.visibility,
                usedRange: null,
                error: err.message
              };
            }
          })
        );

        res.json({
          status: 'success',
          data: {
            worksheets: worksheetInfo,
            totalSheets: worksheetInfo.length
          }
        });
      }

    } catch (err) {
      logger.error('Failed to get worksheet info', {
        driveId: resolvedDriveId,
        itemId: resolvedItemId,
        sheetName,
        error: err.message
      });
      throw err;
    }
  });

  /**
   * Validate operations array
   */
  validateOperations(operations) {
    if (!Array.isArray(operations)) {
      throw new AppError('operations must be an array', 400);
    }

    const validOperationTypes = [
      'highlight', 'backgroundColor', 'textStyle', 'font', 'borders',
      'resizeColumn', 'resizeRow', 'mergeCells', 'unmergeCells',
      'formula', 'conditionalFormatting', 'pivotTable', 'namedRange',
      'dataValidation', 'sort', 'filter'
    ];

    operations.forEach((op, index) => {
      if (!op.type) {
        throw new AppError(`Operation at index ${index} is missing 'type' field`, 400);
      }

      if (!validOperationTypes.includes(op.type)) {
        throw new AppError(`Invalid operation type '${op.type}' at index ${index}. Valid types: ${validOperationTypes.join(', ')}`, 400);
      }

      // Type-specific validation
      switch (op.type) {
        case 'highlight':
        case 'backgroundColor':
          if (!op.range || !op.color) {
            throw new AppError(`Operation '${op.type}' at index ${index} requires 'range' and 'color' fields`, 400);
          }
          break;

        case 'textStyle':
        case 'font':
          if (!op.range) {
            throw new AppError(`Operation '${op.type}' at index ${index} requires 'range' field`, 400);
          }
          break;

        case 'formula':
          if (!op.expression || !op.targetCell) {
            throw new AppError(`Operation 'formula' at index ${index} requires 'expression' and 'targetCell' fields`, 400);
          }
          break;

        case 'resizeColumn':
          if (!op.column) {
            throw new AppError(`Operation 'resizeColumn' at index ${index} requires 'column' field`, 400);
          }
          break;

        case 'resizeRow':
          if (!op.row) {
            throw new AppError(`Operation 'resizeRow' at index ${index} requires 'row' field`, 400);
          }
          break;

        case 'mergeCells':
        case 'unmergeCells':
          if (!op.range) {
            throw new AppError(`Operation '${op.type}' at index ${index} requires 'range' field`, 400);
          }
          break;

        case 'pivotTable':
          if (!op.sourceRange || !op.destinationRange) {
            throw new AppError(`Operation 'pivotTable' at index ${index} requires 'sourceRange' and 'destinationRange' fields`, 400);
          }
          break;

        case 'namedRange':
          if (!op.name || !op.range) {
            throw new AppError(`Operation 'namedRange' at index ${index} requires 'name' and 'range' fields`, 400);
          }
          break;
      }
    });
  }
}

module.exports = new ExcelEngineController();
