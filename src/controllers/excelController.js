// Excel Controller Handles HTTP requests for Excel operations

const excelService = require("../services/excelService");
const resolverService = require('../services/resolverService');
const nameResolverService = require('../services/nameResolverService');
const nameResolutionMixin = require('../middleware/nameResolutionMixin');
const auditService = require('../services/auditService');
const logger = require('../config/logger');
const { catchAsync } = require('../middleware/errorHandler');
const { AppError } = require('../middleware/errorHandler');

class ExcelController {
  // Get all accessible workbooks
  getWorkbooks = catchAsync(async (req, res) => {
    const auditContext = auditService.createAuditContext(req);

    const workbooksResponse = await excelService.getWorkbooks(
      req.accessToken,
      auditContext
    );

    // ✅ Extract only `value` if Graph returns an object with circular refs
    const safeData = Array.isArray(workbooksResponse?.value)
      ? workbooksResponse.value
      : workbooksResponse;

    res.json({
      status: "success",
      data: safeData,
    });
  });

  /**
   * Get worksheets in a workbook
   */
  getWorksheets = catchAsync(async (req, res) => {
    const { driveId, itemId, driveName, itemName, itemPath } = req.query;
    const auditContext = auditService.createAuditContext(req);

    // Resolve IDs if names are provided
    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;

    if ((!resolvedDriveId || !resolvedItemId) && (driveName && itemName)) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
      
      try {
        // Try to resolve by name first, then by path if multiple matches
        resolvedItemId = await resolverService.resolveItemIdByName(req.accessToken, resolvedDriveId, itemName);
      } catch (err) {
        if (err.isMultipleMatches && itemPath) {
          // If multiple matches and path is provided, resolve by path
          resolvedItemId = await resolverService.resolveItemIdByPath(req.accessToken, resolvedDriveId, itemName, itemPath);
        } else {
          throw err; // Re-throw the original error
        }
      }
    }

    const worksheets = await excelService.getWorksheets(
      req.accessToken,
      resolvedDriveId,
      resolvedItemId,
      auditContext
    );

    res.json({
      status: "success",
      data: worksheets,
    });
  });

  // Read data from Excel range
  readRange = catchAsync(async (req, res) => {
    const auditContext = auditService.createAuditContext(req);
    
    // Extract name-based parameters
    const nameParams = nameResolutionMixin.extractNameParams(req);
    
    // Validate input parameters
    nameResolutionMixin.validateNameInput(nameParams);
    
    try {
      // Resolve names to IDs with backward compatibility
      const resolution = await nameResolutionMixin.resolveNames(req, nameParams);
      
      // Handle multiple matches error
      if (!resolution.itemId) {
        throw new AppError('Could not resolve file. Please check the file name and path.', 404);
      }

      // Log name resolution for audit
      nameResolutionMixin.logNameResolution(resolution, 'READ_RANGE', { 
        range: req.body.range,
        worksheetName: req.body.worksheetName 
      });

      // Extract worksheet from range if provided like Sheet1!A1:D10
      const { sheetName, address } = resolverService.parseSheetAndAddress(req.body.range);
      let resolvedWorksheetId = resolution.sheetId;
      const effectiveWorksheetName = req.body.worksheetName || sheetName || resolution.sheetName;
      
      if (!resolvedWorksheetId && effectiveWorksheetName) {
        resolvedWorksheetId = await resolverService.resolveWorksheetIdByName(
          req.accessToken,
          resolution.driveId,
          resolution.itemId,
          effectiveWorksheetName
        );
      }

      const data = await excelService.readRange({
        accessToken: req.accessToken,
        driveId: resolution.driveId,
        itemId: resolution.itemId,
        worksheetId: resolvedWorksheetId,
        range: address,
        auditContext,
      });

      res.json({
        status: "success",
        data: {
          range: data.address,
          values: data.values,
          formulas: data.formulas,
          text: data.text,
          dimensions: {
            rows: data.rowCount,
            columns: data.columnCount,
          },
        },
        resolution: nameResolutionMixin.getResolutionSummary(resolution)
      });

    } catch (err) {
      // Handle multiple matches with user-friendly response
      if (err.isMultipleMatches) {
        return nameResolutionMixin.handleMultipleMatches(res, err, 'file');
      }
      throw err;
    }
  });

  /**
   * Write data to Excel range
   */
  writeRange = catchAsync(async (req, res) => {
    const { driveId, itemId, driveName, itemName, itemPath, worksheetId, worksheetName, range, values } = req.body;
    const auditContext = auditService.createAuditContext(req);

    // Resolve drive/item
    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;
    if ((!resolvedDriveId || !resolvedItemId) && (driveName && itemName)) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
      
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

    // Resolve worksheet and address
    const { sheetName, address } = resolverService.parseSheetAndAddress(range);
    let resolvedWorksheetId = worksheetId;
    const effectiveWorksheetName = worksheetName || sheetName;
    if (!resolvedWorksheetId && effectiveWorksheetName) {
      resolvedWorksheetId = await resolverService.resolveWorksheetIdByName(
        req.accessToken,
        resolvedDriveId,
        resolvedItemId,
        effectiveWorksheetName
      );
    }

    const data = await excelService.writeRange({
      accessToken: req.accessToken,
      driveId: resolvedDriveId,
      itemId: resolvedItemId,
      worksheetId: resolvedWorksheetId,
      range: address,
      values,
      auditContext,
    });

    res.json({
      status: "success",
      data: {
        range: data.address,
        values: data.values,
        dimensions: {
          rows: data.rowCount,
          columns: data.columnCount,
        },
      },
    });
  });

  /**
   * Read data from Excel table
   */
  readTable = catchAsync(async (req, res) => {
    const { driveId, itemId, driveName, itemName, itemPath, worksheetId, worksheetName, tableName } = req.body;
    const auditContext = auditService.createAuditContext(req);

    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;
    if ((!resolvedDriveId || !resolvedItemId) && (driveName && itemName)) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
      
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

    let resolvedWorksheetId = worksheetId;
    if (!resolvedWorksheetId && worksheetName) {
      resolvedWorksheetId = await resolverService.resolveWorksheetIdByName(
        req.accessToken,
        resolvedDriveId,
        resolvedItemId,
        worksheetName
      );
    }

    const data = await excelService.readTable({
      accessToken: req.accessToken,
      driveId: resolvedDriveId,
      itemId: resolvedItemId,
      worksheetId: resolvedWorksheetId,
      tableName,
      auditContext,
    });

    res.json({
      status: "success",
      data: {
        table: {
          id: data.id,
          name: data.name,
          address: data.address,
        },
        headers: data.headers,
        rows: data.rows,
        values: data.values,
        dimensions: {
          rows: data.rowCount,
          columns: data.columnCount,
        },
      },
    });
  });

  /**
   * Add rows to Excel table
   */
  addTableRows = catchAsync(async (req, res) => {
    const { driveId, itemId, driveName, itemName, itemPath, worksheetId, worksheetName, tableName, rows } = req.body;
    const auditContext = auditService.createAuditContext(req);

    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;
    if ((!resolvedDriveId || !resolvedItemId) && (driveName && itemName)) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
      
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

    let resolvedWorksheetId = worksheetId;
    if (!resolvedWorksheetId && worksheetName) {
      resolvedWorksheetId = await resolverService.resolveWorksheetIdByName(
        req.accessToken,
        resolvedDriveId,
        resolvedItemId,
        worksheetName
      );
    }

    const result = await excelService.addTableRows({
      accessToken: req.accessToken,
      driveId: resolvedDriveId,
      itemId: resolvedItemId,
      worksheetId: resolvedWorksheetId,
      tableName,
      rows,
      auditContext,
    });

    res.json({
      status: "success",
      data: {
        rowsAdded: rows.length,
        result: result,
      },
    });
  });

  /**
   * Batch operations - perform multiple Excel operations in sequence
   */
  batchOperations = catchAsync(async (req, res) => {
    const { operations } = req.body;
    const auditContext = auditService.createAuditContext(req);

    if (!Array.isArray(operations) || operations.length === 0) {
      return res.status(400).json({
        success: false,
        error: "Invalid operations array",
        timestamp: new Date().toISOString(),
      });
    }

    const results = [];
    const errors = [];

    for (let i = 0; i < operations.length; i++) {
      const operation = operations[i];

      try {
        let result;

        switch (operation.type) {
          case "READ_range":
            result = await excelService.readRange({
              accessToken: req.accessToken,
              driveId: operation.driveId,
              itemId: operation.itemId,
              worksheetId: operation.worksheetId,
              range: operation.range,
              auditContext,
            });
            break;

          case "write_range":
            result = await excelService.writeRange({
              accessToken: req.accessToken,
              driveId: operation.driveId,
              itemId: operation.itemId,
              worksheetId: operation.worksheetId,
              range: operation.range,
              values: operation.values,
              auditContext,
            });
            break;

          case "read_table":
            result = await excelService.readTable({
              accessToken: req.accessToken,
              driveId: operation.driveId,
              itemId: operation.itemId,
              worksheetId: operation.worksheetId,
              tableName: operation.tableName,
              auditContext,
            });
            break;

          case "add_table_rows":
            result = await excelService.addTableRows({
              accessToken: req.accessToken,
              driveId: operation.driveId,
              itemId: operation.itemId,
              worksheetId: operation.worksheetId,
              tableName: operation.tableName,
              rows: operation.rows,
              auditContext,
            });
            break;

          default:
            throw new Error(`Unknown operation type: ${operation.type}`);
        }

        results.push({
          index: i,
          operation: operation.type,
          success: true,
          data: result,
        });
      } catch (error) {
        logger.error(`Batch operation ${i} failed:`, error);
        errors.push({
          index: i,
          operation: operation.type,
          error: error.message,
        });

        // Continue with other operations unless it's a critical error
        if (error.message.includes("Authentication")) {
          break; // Stop if authentication fails
        }
      }
    }

    const response = {
      status: errors.length === 0 ? "success" : "partial_success",
      data: {
        results: results,
        errors: errors,
        summary: {
          total: operations.length,
          successful: results.length,
          failed: errors.length,
        },
      },
    };

    // Return 207 Multi-Status if there were partial failures
    const statusCode =
      errors.length > 0 && results.length > 0
        ? 207
        : errors.length === 0
        ? 200
        : 400;

    res.status(statusCode).json(response);
  });

  /**
   * Get Excel file metadata
   */
  getFileMetadata = catchAsync(async (req, res) => {
    const { driveId, itemId } = req.query;
    const auditContext = auditService.createAuditContext(req);

    // This would typically get file metadata from Graph API
    // For now, return basic structure
    res.json({
      status: "success",
      data: {
        fileId: itemId,
        driveId: driveId,
      },
    });
  });

  /**
   * Search for files recursively and return all matches with their paths
   * This endpoint helps users find files when they don't know the exact location
   */
  searchFiles = catchAsync(async (req, res) => {
    const { driveName, fileName } = req.query;
    const auditContext = auditService.createAuditContext(req);

    if (!driveName || !fileName) {
      return res.status(400).json({
        status: "error",
        error: {
          code: 400,
          message: "Both driveName and fileName are required"
        },
        timestamp: new Date().toISOString()
      });
    }

    try {
      // Resolve drive ID
      const driveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
      
      // Create graph client and perform recursive search
      const graphClient = resolverService.createGraphClient(req.accessToken);
      const matches = await resolverService.recursiveSearchForFile(graphClient, driveId, fileName);
      
      auditService.logSystemEvent({
        event: "FILE_SEARCH",
        details: { 
          driveName, 
          fileName, 
          matchCount: matches.length,
          requestedBy: auditContext.user 
        },
      });

      res.json({
        status: "success",
        data: {
          fileName,
          driveName,
          matches: matches.map(match => ({
            id: match.id,
            name: match.name,
            path: match.path,
            parentId: match.parentId
          })),
          totalMatches: matches.length
        }
      });
    } catch (err) {
      logger.error('File search failed:', { driveName, fileName, error: err.message });
      throw err;
    }
  });

  /**
   * Get audit logs
   */
  getLogs = catchAsync(async (req, res) => {
    const { startDate, endDate, operation, user, limit = 100 } = req.query;
    const auditContext = auditService.createAuditContext(req);

    // In a real implementation, this would query a database or log files
    // For now, we'll return a sample response showing the log structure
    const logs = {
      logs: [
        {
          id: "audit-001",
          timestamp: new Date().toISOString(),
          operation: "READ",
          user: auditContext.user,
          workbookId: "sample-workbook-id",
          worksheetId: "Sheet1",
          range: "A1:C10",
          success: true,
          requestId: auditContext.requestId,
          ipAddress: auditContext.ipAddress,
        },
      ],
      filters: {
        startDate: startDate || null,
        endDate: endDate || null,
        operation: operation || null,
        user: user || null,
        limit: parseInt(limit),
      },
      count: 1,
      totalCount: 1,
    };

    auditService.logSystemEvent({
      event: "AUDIT_LOG_REQUEST",
      details: { filters: logs.filters, requestedBy: auditContext.user },
    });

    res.json({
      status: "success",
      data: logs,
    });
  });
}

module.exports = new ExcelController();
