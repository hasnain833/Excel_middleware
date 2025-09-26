const excelEngineService = require('../services/excelEngineService');
const resolverService = require('../services/resolverService');
const auditService = require('../services/auditService');
const logger = require('../config/logger');
const { catchAsync } = require('../middleware/errorHandler');
const { AppError } = require('../middleware/errorHandler');

class ExcelEngineController {

  applyFormatting = catchAsync(async (req, res) => {
    const { 
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

    // Resolve drive and item IDs from names only
    const resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    let resolvedItemId;
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
            matches: err.matches.map(match => ({ id: match.id, name: match.name, path: match.path, parentId: match.parentId }))
          });
        }
      } else {
        throw err;
      }
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
      logger.error('Excel engine operation failed', { driveId: resolvedDriveId, itemId: resolvedItemId, sheetName, error: err.message });
      throw err;
    }
  });


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
