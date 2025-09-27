const renameService = require('../services/renameService');
const resolverService = require('../services/resolverService');
const auditService = require('../services/auditService');
const logger = require('../config/logger');
const { catchAsync } = require('../middleware/errorHandler');
const { AppError } = require('../middleware/errorHandler');

class RenameController {

  renameFile = catchAsync(async (req, res) => {
    const {
      driveName,
      itemName,
      itemPath,
      oldName,
      newName,
      selectedItemId
    } = req.body;

    const auditContext = auditService.createAuditContext(req);

    if (!newName) {
      throw new AppError('newName is required', 400);
    }

    // Helper to build display paths from Graph parentReference.path
    const toDisplayPath = (drvName, graphParentPath, fileNm) => {
      let folderPath = '';
      if (graphParentPath) {
        const idx = graphParentPath.indexOf('root:');
        folderPath = idx >= 0 ? graphParentPath.substring(idx + 'root:'.length) : graphParentPath;
      }
      if (!folderPath || folderPath === '/') folderPath = '';
      // Ensure no double slashes
      const prefix = drvName ? `${drvName}` : '';
      const combined = [prefix, folderPath.replace(/^\/?/, ''), fileNm].filter(Boolean).join('/');
      return combined;
    };

    let resolvedDriveId;
    let resolvedDriveName = driveName || null;
    let resolvedItemId;

    // 1) If selectedItemId is provided (second-call flow), resolve drive automatically if needed
    if (selectedItemId) {
      if (driveName) {
        resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
        resolvedDriveName = driveName;
      } else {
        const located = await resolverService.findDriveForItemId(req.accessToken, selectedItemId);
        resolvedDriveId = located.driveId;
        resolvedDriveName = located.driveName;
      }
      resolvedItemId = selectedItemId;
    } else if (driveName) {
      // 2) Backward compatible path: driveName provided
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
      resolvedDriveName = driveName;
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(req.accessToken, resolvedDriveId, itemName);
      } catch (err) {
        if (err.isMultipleMatches) {
          if (itemPath) {
            resolvedItemId = await resolverService.resolveItemIdByPath(req.accessToken, resolvedDriveId, itemName, itemPath);
          } else {
            const matches = (err.matches || []).map((m) => {
              const breadcrumb = `${resolvedDriveName} `.replace(/\u000b/g, '›'); // ensure separator
              const fullPath = `${resolvedDriveName}${m.path}`.replace(/\\/g, '/');
              return {
                id: m.id,
                name: m.name,
                driveName: resolvedDriveName,
                breadcrumb: `${resolvedDriveName} › ${m.path.split('/').filter(Boolean).slice(0, -1).join(' › ')}`,
                fullPath,
                path: m.path,
                parentId: m.parentId,
              };
            });
            return res.status(409).json({
              status: 'multiple_matches',
              message: 'Multiple files found with that name. Please pick one.',
              matches,
            });
          }
        } else {
          throw err;
        }
      }
    } else {
      // 3) No driveName: search across all drives
      if (!itemName) {
        throw new AppError('Either itemName or selectedItemId is required', 400);
      }
      const allMatches = await resolverService.searchAllDrivesForFileByExactName(req.accessToken, itemName);
      if (allMatches.length === 0) {
        throw new AppError(`File '${itemName}' not found across any accessible drives`, 404);
      }
      if (allMatches.length > 1) {
        const matches = allMatches.map((m) => ({
          id: m.id,
          name: m.name,
          driveName: m.driveName,
          breadcrumb: `${m.driveName} › ${m.path.split('/').filter(Boolean).slice(0, -1).join(' › ')}`,
          fullPath: `${m.driveName}${m.path}`,
          path: m.path,
          parentId: m.parentId,
        }));
        return res.status(409).json({
          status: 'multiple_matches',
          message: 'Multiple files found with that name. Please pick one.',
          matches,
        });
      }
      const only = allMatches[0];
      resolvedDriveId = only.driveId;
      resolvedDriveName = only.driveName;
      resolvedItemId = only.id;
    }

    const result = await renameService.renameFile(
      req.accessToken,
      resolvedDriveId,
      resolvedItemId,
      oldName,
      newName,
      auditContext
    );

    // Build pathBefore and pathAfter for message
    const pathBefore = toDisplayPath(resolvedDriveName, result.path, result.oldName);
    const pathAfter = toDisplayPath(resolvedDriveName, result.path, result.newName);

    res.json({
      status: 'success',
      data: {
        file: result,
        message: `Renamed: ${pathBefore} -> ${pathAfter}`,
        pathBefore,
        pathAfter,
      }
    });
  });

  renameFolder = catchAsync(async (req, res) => {
    const {
      driveName,
      folderName,
      folderPath,
      oldName,
      newName
    } = req.body;

    const auditContext = auditService.createAuditContext(req);

    if (!newName) {
      throw new AppError('newName is required', 400);
    }

    // Resolve drive ID from name
    const resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);

    // Resolve folder ID via name, optionally disambiguated by folderPath
    let effectiveFolderName = folderName;
    if (!effectiveFolderName && folderPath) {
      const parts = folderPath.split('/').filter(Boolean);
      effectiveFolderName = parts[parts.length - 1];
    }

    let resolvedFolderId;
    if (effectiveFolderName) {
      const graphClient = resolverService.createGraphClient(req.accessToken);
      const matches = await this.findFoldersByName(graphClient, resolvedDriveId, effectiveFolderName);

      if (matches.length === 0) {
        throw new AppError(`Folder '${effectiveFolderName}' not found`, 404);
      }

      if (matches.length > 1) {
        if (folderPath) {
          const match = matches.find(m => m.path === folderPath);
          if (!match) {
            throw new AppError(`Folder '${effectiveFolderName}' not found at path '${folderPath}'`, 404);
          }
          resolvedFolderId = match.id;
        } else {
          return res.status(409).json({
            status: 'multiple_matches',
            message: 'Multiple folders found with the same name. Please specify folderPath or select from the list.',
            matches: matches.map(match => ({
              id: match.id,
              name: match.name,
              path: match.path,
              parentId: match.parentId
            }))
          });
        }
      } else {
        resolvedFolderId = matches[0].id;
      }
    }

    if (!resolvedDriveId || !resolvedFolderId) {
      throw new AppError('Could not resolve drive or folder. Please provide valid names.', 400);
    }

    const result = await renameService.renameFolder(
      req.accessToken,
      resolvedDriveId,
      resolvedFolderId,
      oldName,
      newName,
      auditContext
    );

    res.json({
      status: 'success',
      data: {
        folder: result,
        message: `Folder renamed from '${result.oldName}' to '${result.newName}'`
      }
    });
  });


  renameSheet = catchAsync(async (req, res) => {
    const {
      driveName,
      itemName,
      itemPath,
      oldSheetName,
      newSheetName
    } = req.body;

    const auditContext = auditService.createAuditContext(req);

    if (!oldSheetName || !newSheetName) {
      throw new AppError('oldSheetName and newSheetName are required', 400);
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

    const result = await renameService.renameSheet(
      req.accessToken,
      resolvedDriveId,
      resolvedItemId,
      oldSheetName,
      newSheetName,
      auditContext
    );

    res.json({
      status: 'success',
      data: {
        worksheet: result,
        message: `Sheet renamed from '${result.oldSheetName}' to '${result.newSheetName}' in file '${result.fileName}'`
      }
    });
  });

  getRenameSuggestions = catchAsync(async (req, res) => {
    const {
      driveName,
      oldTerm,
      newTerm
    } = req.body;

    const auditContext = auditService.createAuditContext(req);

    if (!oldTerm || !newTerm) {
      throw new AppError('oldTerm and newTerm are required', 400);
    }

    const resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);

    if (!resolvedDriveId) {
      throw new AppError('Could not resolve drive. Please provide valid drive name.', 400);
    }

    const suggestions = await renameService.findRelatedItems(
      req.accessToken,
      resolvedDriveId,
      oldTerm,
      newTerm
    );

    auditService.logSystemEvent({
      event: 'RENAME_SUGGESTIONS_REQUESTED',
      details: {
        driveId: resolvedDriveId,
        oldTerm,
        newTerm,
        suggestionsCount: suggestions.length,
        requestedBy: auditContext.user
      }
    });

    res.json({
      status: 'success',
      data: {
        oldTerm,
        newTerm,
        suggestions,
        message: suggestions.length > 0
          ? `Found ${suggestions.length} items that might need renaming`
          : 'No related items found that need renaming'
      }
    });
  });


  batchRename = catchAsync(async (req, res) => {
    const {
      driveName,
      operations
    } = req.body;

    const auditContext = auditService.createAuditContext(req);

    if (!Array.isArray(operations) || operations.length === 0) {
      throw new AppError('operations array is required and must not be empty', 400);
    }

    let resolvedDriveId = driveId;
    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    if (!resolvedDriveId) {
      throw new AppError('Could not resolve drive. Please provide valid drive identifier.', 400);
    }

    const result = await renameService.batchRename(
      req.accessToken,
      resolvedDriveId,
      operations,
      auditContext
    );

    const statusCode = result.errors.length > 0 && result.results.length > 0
      ? 207 // Multi-Status
      : result.errors.length === 0
        ? 200
        : 400;

    res.status(statusCode).json({
      status: result.errors.length === 0 ? 'success' : 'partial_success',
      data: result
    });
  });

  async findFoldersByName(graphClient, driveId, folderName) {
    try {
      // Use the recursive search from renameService but filter for folders only
      const allItems = await renameService.recursiveSearch(graphClient, driveId, folderName);

      return allItems.filter(item =>
        item.folder &&
        item.name.toLowerCase() === folderName.toLowerCase()
      );

    } catch (err) {
      logger.error('Failed to find folders by name', {
        driveId,
        folderName,
        error: err.message
      });
      return [];
    }
  }
}

module.exports = new RenameController();
