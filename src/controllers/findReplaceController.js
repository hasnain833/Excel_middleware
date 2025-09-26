const findReplaceService = require("../services/findReplaceService");
const resolverService = require("../services/resolverService");
const auditService = require("../services/auditService");
const logger = require("../config/logger");
const { catchAsync } = require("../middleware/errorHandler");
const { AppError } = require("../middleware/errorHandler");

class FindReplaceController {
  findReplace = catchAsync(async (req, res) => {
    const {
      driveName,
      itemName,
      itemPath,
      searchTerm,
      replaceTerm,
      scope = "entire_sheet",
      rangeSpec,
      highlightChanges = false,
      logChanges = true,
      confirm = false,
      previewId,
    } = req.body;

    const auditContext = auditService.createAuditContext(req);

    // Validate required parameters
    if (!searchTerm) {
      throw new AppError("searchTerm is required", 400);
    }

    if (!replaceTerm && confirm) {
      throw new AppError(
        "replaceTerm is required for replacement operation",
        400
      );
    }

    // Resolve drive and item IDs via names only
    const resolvedDriveId = await resolverService.resolveDriveIdByName(
      req.accessToken,
      driveName
    );
    let resolvedItemId;
    try {
      resolvedItemId = await resolverService.resolveItemIdByName(
        req.accessToken,
        resolvedDriveId,
        itemName
      );
    } catch (err) {
      if (err.isMultipleMatches) {
        if (itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(
            req.accessToken,
            resolvedDriveId,
            itemName,
            itemPath
          );
        } else {
          return res.status(409).json({
            status: "multiple_matches",
            message:
              "Multiple files found with the same name. Please specify itemPath or select from the list.",
            matches: err.matches.map((match) => ({
              id: match.id,
              name: match.name,
              path: match.path,
              parentId: match.parentId,
            })),
          });
        }
      } else {
        throw err;
      }
    }

    // Validate scope and range
    if (scope === "specific_range" && !rangeSpec) {
      throw new AppError(
        'rangeSpec is required when scope is "specific_range"',
        400
      );
    }

    try {
      // Step 1: Find all occurrences
      const matches = await findReplaceService.findOccurrences(
        req.accessToken,
        resolvedDriveId,
        resolvedItemId,
        searchTerm,
        scope,
        rangeSpec
      );

      // If no matches found
      if (matches.length === 0) {
        return res.json({
          status: "no_matches",
          message: `No occurrences of '${searchTerm}' found.`,
          data: {
            searchTerm,
            scope,
            rangeSpec,
            totalMatches: 0,
          },
        });
      }
      if (!confirm) {
        const preview = findReplaceService.generatePreview(matches, searchTerm);
        const previewSessionId = `preview_${Date.now()}_${Math.random()
          .toString(36)
          .substr(2, 9)}`;
        return res.json({
          status: "preview",
          message: this.generatePreviewMessage(preview, replaceTerm),
          data: {
            previewId: previewSessionId,
            searchTerm,
            replaceTerm,
            scope,
            rangeSpec,
            preview,
            confirmationRequired: true,
            instructions:
              'To proceed with replacement, send the same request with "confirm": true and include this previewId',
          },
        });
      }

      if (!replaceTerm) {
        throw new AppError(
          "replaceTerm is required for confirmed replacement",
          400
        );
      }

      const result = await findReplaceService.performReplace(
        req.accessToken,
        resolvedDriveId,
        resolvedItemId,
        searchTerm,
        replaceTerm,
        matches,
        { highlightChanges, logChanges }
      );

      // Log the operation
      auditService.logSystemEvent({
        event: "FIND_REPLACE_COMPLETED",
        details: {
          driveId: resolvedDriveId,
          itemId: resolvedItemId,
          searchTerm,
          replaceTerm,
          scope,
          rangeSpec,
          totalMatches: matches.length,
          successful: result.summary.successful,
          failed: result.summary.failed,
          highlightChanges,
          logChanges,
          requestedBy: auditContext.user,
        },
      });

      res.json({
        status: "success",
        message: `Successfully replaced ${result.summary.successful} occurrences of '${searchTerm}' with '${replaceTerm}'`,
        data: {
          searchTerm,
          replaceTerm,
          scope,
          rangeSpec,
          summary: result.summary,
          changes: logChanges ? result.changes : undefined,
          errors: result.errors.length > 0 ? result.errors : undefined,
          highlightChanges,
          logChanges,
        },
      });
    } catch (err) {
      logger.error("Find and replace operation failed", {
        driveId: resolvedDriveId,
        itemId: resolvedItemId,
        searchTerm,
        replaceTerm,
        scope,
        error: err.message,
      });
      throw err;
    }
  });

  searchText = catchAsync(async (req, res) => {
    const {
      driveName,
      itemName,
      itemPath,
      searchTerm,
      scope = "entire_sheet",
      rangeSpec,
    } = req.body;

    const auditContext = auditService.createAuditContext(req);

    if (!searchTerm) {
      throw new AppError("searchTerm is required", 400);
    }

    // Resolve drive and item IDs via names only
    const resolvedDriveId = await resolverService.resolveDriveIdByName(
      req.accessToken,
      driveName
    );
    let resolvedItemId;
    try {
      resolvedItemId = await resolverService.resolveItemIdByName(
        req.accessToken,
        resolvedDriveId,
        itemName
      );
    } catch (err) {
      if (err.isMultipleMatches) {
        if (itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(
            req.accessToken,
            resolvedDriveId,
            itemName,
            itemPath
          );
        } else {
          return res.status(409).json({
            status: "multiple_matches",
            message:
              "Multiple files found with the same name. Please specify itemPath or select from the list.",
            matches: err.matches.map((match) => ({
              id: match.id,
              name: match.name,
              path: match.path,
              parentId: match.parentId,
            })),
          });
        }
      } else {
        throw err;
      }
    }

    const matches = await findReplaceService.findOccurrences(
      req.accessToken,
      resolvedDriveId,
      resolvedItemId,
      searchTerm,
      scope,
      rangeSpec
    );

    const preview = findReplaceService.generatePreview(matches, searchTerm);

    // Log search operation
    auditService.logSystemEvent({
      event: "TEXT_SEARCH_PERFORMED",
      details: {
        driveId: resolvedDriveId,
        itemId: resolvedItemId,
        searchTerm,
        scope,
        rangeSpec,
        matchCount: matches.length,
        requestedBy: auditContext.user,
      },
    });

    res.json({
      status: "success",
      message:
        matches.length > 0
          ? `Found ${matches.length} occurrences of '${searchTerm}'`
          : `No occurrences of '${searchTerm}' found`,
      data: {
        searchTerm,
        scope,
        rangeSpec,
        preview,
        matches: matches.slice(0, 50), // Limit response size
      },
    });
  });

  analyzeScope = catchAsync(async (req, res) => {
    const { driveName, itemName, itemPath } = req.query;

    const auditContext = auditService.createAuditContext(req);

    // Resolve drive and item IDs via names only
    const resolvedDriveId = await resolverService.resolveDriveIdByName(
      req.accessToken,
      driveName
    );
    let resolvedItemId;
    try {
      resolvedItemId = await resolverService.resolveItemIdByName(
        req.accessToken,
        resolvedDriveId,
        itemName
      );
    } catch (err) {
      if (err.isMultipleMatches && itemPath) {
        resolvedItemId = await resolverService.resolveItemIdByPath(
          req.accessToken,
          resolvedDriveId,
          itemName,
          itemPath
        );
      } else {
        throw err;
      }
    }

    try {
      const graphClient = findReplaceService.createGraphClient(req.accessToken);

      // Get worksheet information
      const worksheetsResp = await graphClient
        .api(
          `/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets`
        )
        .get();

      const scopeAnalysis = {
        worksheets: [],
        totalSheets: worksheetsResp.value?.length || 0,
        availableScopes: [
          "header_only",
          "specific_range",
          "entire_sheet",
          "all_sheets",
        ],
      };

      // Analyze each worksheet
      for (const worksheet of worksheetsResp.value || []) {
        try {
          const usedRangeResp = await graphClient
            .api(
              `/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/${worksheet.id}/usedRange`
            )
            .get();

          scopeAnalysis.worksheets.push({
            name: worksheet.name,
            id: worksheet.id,
            usedRange: usedRangeResp?.address || "Empty",
            rowCount: usedRangeResp?.rowCount || 0,
            columnCount: usedRangeResp?.columnCount || 0,
          });
        } catch (sheetErr) {
          scopeAnalysis.worksheets.push({
            name: worksheet.name,
            id: worksheet.id,
            usedRange: "Error accessing sheet",
            rowCount: 0,
            columnCount: 0,
            error: sheetErr.message,
          });
        }
      }

      res.json({
        status: "success",
        data: scopeAnalysis,
      });
    } catch (err) {
      logger.error("Failed to analyze scope", {
        driveId: resolvedDriveId,
        itemId: resolvedItemId,
        error: err.message,
      });
      throw err;
    }
  });

  generatePreviewMessage(preview, replaceTerm) {
    const { totalMatches, breakdown, bySheet } = preview;

    let message = `I found ${totalMatches} cells containing '${preview.searchTerm}':\n`;

    if (breakdown.headers > 0) {
      message += `• ${breakdown.headers} in header rows\n`;
    }

    if (breakdown.dataRows > 0) {
      message += `• ${breakdown.dataRows} in data rows\n`;
    }

    if (Object.keys(bySheet).length > 1) {
      message += "\nBy sheet:\n";
      Object.entries(bySheet).forEach(([sheet, count]) => {
        message += `• ${sheet}: ${count} occurrences\n`;
      });
    }

    if (replaceTerm) {
      message += `\nWould you like to replace all occurrences with '${replaceTerm}'?`;
      message += '\nTo proceed, send the same request with "confirm": true';
    }

    return message;
  }
}

module.exports = new FindReplaceController();
