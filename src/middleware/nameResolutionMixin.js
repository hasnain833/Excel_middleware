const nameResolverService = require("../services/nameResolverService");
const { AppError } = require("./errorHandler");
const logger = require("../config/logger");

class NameResolutionMixin {
  async resolveNames(req, inputParams) {
    const {
      // Legacy ID-based parameters (for backward compatibility)
      driveId,
      itemId,
      worksheetId,

      // New name-based parameters
      driveName,
      folderName,
      fileName,
      sheetName,

      // Path-based resolution
      fullPath,
      itemPath,
    } = inputParams;

    // If legacy IDs are provided, use them directly (backward compatibility)
    if (driveId && itemId) {
      logger.info("Using legacy ID-based resolution", {
        driveId,
        itemId,
        worksheetId,
      });
      return {
        driveId,
        itemId,
        worksheetId,
        driveName: driveName || "Unknown",
        fileName: fileName || "Unknown",
        sheetName: sheetName || "Unknown",
        legacy: true,
        resolvedAt: new Date().toISOString(),
      };
    }

    // If only driveId is provided but names are also provided, use hybrid approach
    if (driveId && !itemId && (fileName || folderName)) {
      logger.info("Using hybrid ID/name resolution", {
        driveId,
        fileName,
        folderName,
        sheetName,
      });

      const resolution = {
        driveId,
        driveName: driveName || "Unknown",
        folderId: null,
        folderName,
        itemId: null,
        fileName,
        sheetId: null,
        sheetName,
        hybrid: true,
        resolvedAt: new Date().toISOString(),
      };

      // Resolve remaining names to IDs
      if (fileName) {
        const fileResult = await nameResolverService.resolveFileId(
          req.accessToken,
          driveId,
          "root",
          fileName
        );
        resolution.itemId = fileResult.id;
        resolution.filePath = fileResult.path;
      }

      if (sheetName && resolution.itemId) {
        resolution.sheetId = await nameResolverService.resolveSheetId(
          req.accessToken,
          driveId,
          resolution.itemId,
          sheetName
        );
      }

      return resolution;
    }

    // Full name-based resolution (new approach)
    if (!driveName) {
      throw new AppError("Either driveId or driveName must be provided", 400);
    }

    // Handle full path resolution
    if (fullPath) {
      logger.info("Using full path resolution", { driveName, fullPath });
      const pathResult = await nameResolverService.resolveByPath(
        req.accessToken,
        driveName,
        fullPath
      );

      return {
        driveId: pathResult.driveId,
        itemId: pathResult.itemId,
        driveName,
        fileName: pathResult.itemName,
        filePath: fullPath,
        fullPathResolution: true,
        resolvedAt: new Date().toISOString(),
      };
    }

    // Handle itemPath resolution (for duplicate disambiguation)
    if (itemPath && fileName) {
      logger.info("Using item path resolution for disambiguation", {
        driveName,
        fileName,
        itemPath,
      });

      try {
        const driveId = await nameResolverService.resolveDriveId(
          req.accessToken,
          driveName
        );
        const matches = await nameResolverService.recursiveSearchForFile(
          nameResolverService.createGraphClient(req.accessToken),
          driveId,
          fileName
        );

        const match = matches.find((m) => m.path === itemPath);
        if (!match) {
          const availablePaths = matches.map((m) => m.path);
          throw new AppError(
            `File '${fileName}' not found at path '${itemPath}'. Available paths: ${availablePaths.join(
              ", "
            )}`,
            404
          );
        }

        const resolution = {
          driveId,
          itemId: match.id,
          driveName,
          fileName,
          filePath: itemPath,
          itemPathResolution: true,
          resolvedAt: new Date().toISOString(),
        };

        // Resolve sheet if provided
        if (sheetName) {
          resolution.sheetId = await nameResolverService.resolveSheetId(
            req.accessToken,
            driveId,
            match.id,
            sheetName
          );
          resolution.sheetName = sheetName;
        }

        return resolution;
      } catch (err) {
        throw new AppError(
          `Failed to resolve item path: ${err.message}`,
          err.status || 500
        );
      }
    }

    // Standard name-based resolution
    logger.info("Using standard name-based resolution", {
      driveName,
      folderName,
      fileName,
      sheetName,
    });

    try {
      const resolution = await nameResolverService.resolveIdByName(
        req.accessToken,
        driveName,
        folderName,
        fileName,
        sheetName
      );

      return resolution;
    } catch (err) {
      // Handle multiple matches error with user-friendly response
      if (err.isMultipleMatches) {
        throw err; // Let controller handle this with 409 response
      }

      throw new AppError(
        `Name resolution failed: ${err.message}`,
        err.status || 500
      );
    }
  }
  handleMultipleMatches(res, error, entityType = "item") {
    const matches = error.matches.map((match, index) => ({
      index: index + 1,
      id: match.id,
      name: match.name,
      path: match.path,
      parentId: match.parentId,
    }));

    return res.status(409).json({
      status: "multiple_matches",
      message: `Multiple ${entityType}s found with the same name. Please specify the full path or select from the list.`,
      entityType,
      matches,
      instructions: {
        useFullPath: `Provide 'fullPath' parameter with the complete path (e.g., '/Folder1/Subfolder/file.xlsx')`,
        useItemPath: `Provide 'itemPath' parameter with the specific path from the matches above`,
        selectById: `Use the 'id' from one of the matches above in your next request`,
      },
    });
  }

  validateNameInput(inputParams) {
    const {
      driveId,
      driveName,
      folderName,
      fileName,
      sheetName,
      fullPath,
      itemPath,
    } = inputParams;

    if (!driveId && !driveName) {
      throw new AppError("Either driveId or driveName must be provided", 400);
    }

    if (fullPath && !fullPath.startsWith("/")) {
      throw new AppError(
        "fullPath must start with / (e.g., /Folder/file.xlsx)",
        400
      );
    }

    if (itemPath && !fileName) {
      throw new AppError("fileName is required when using itemPath", 400);
    }
    if (sheetName && !fileName && !inputParams.itemId) {
      throw new AppError(
        "fileName (or itemId) is required when specifying sheetName",
        400
      );
    }

    return true;
  }

  logNameResolution(resolution, operation, additionalDetails = {}) {
    const logEntry = {
      event: "NAME_RESOLUTION_AUDIT",
      operation,
      resolution: {
        driveName: resolution.driveName,
        driveId: resolution.driveId,
        folderName: resolution.folderName,
        folderId: resolution.folderId,
        fileName: resolution.fileName,
        itemId: resolution.itemId,
        sheetName: resolution.sheetName,
        sheetId: resolution.sheetId,
        resolvedAt: resolution.resolvedAt,
      },
      additionalDetails,
      timestamp: new Date().toISOString(),
    };

    logger.info("Name resolution audit", logEntry);
    return logEntry;
  }

  extractNameParams(req) {
    const body = req.body || {};
    const query = req.query || {};

    // Combine body and query parameters, with body taking precedence
    return {
      // Legacy ID parameters
      driveId: body.driveId || query.driveId,
      itemId: body.itemId || query.itemId,
      worksheetId: body.worksheetId || query.worksheetId,

      // Name-based parameters
      driveName: body.driveName || query.driveName,
      folderName: body.folderName || query.folderName,
      fileName:
        body.fileName || body.itemName || query.fileName || query.itemName,
      sheetName:
        body.sheetName ||
        body.worksheetName ||
        query.sheetName ||
        query.worksheetName,

      // Path-based parameters
      fullPath: body.fullPath || query.fullPath,
      itemPath: body.itemPath || query.itemPath,
    };
  }

  createNameResolutionError(entityType, entityName, availableOptions = []) {
    let message = `${entityType} '${entityName}' not found.`;

    if (availableOptions.length > 0) {
      message += ` Available ${entityType.toLowerCase()}s: ${availableOptions.join(
        ", "
      )}`;
    }

    return new AppError(message, 404);
  }

  getResolutionSummary(resolution) {
    const summary = {
      resolvedAt: resolution.resolvedAt,
      resolutionType: "name-based",
    };

    if (resolution.legacy) {
      summary.resolutionType = "legacy-id-based";
    } else if (resolution.hybrid) {
      summary.resolutionType = "hybrid-id-name";
    } else if (resolution.fullPathResolution) {
      summary.resolutionType = "full-path";
    } else if (resolution.itemPathResolution) {
      summary.resolutionType = "item-path";
    }

    summary.resolved = {
      drive: resolution.driveName,
      folder: resolution.folderName,
      file: resolution.fileName,
      sheet: resolution.sheetName,
    };

    return summary;
  }
}

module.exports = new NameResolutionMixin();
