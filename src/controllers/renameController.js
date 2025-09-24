/**
 * Rename Controller
 * Handles HTTP requests for renaming files, folders, and sheets
 */

const renameService = require('../services/renameService');
const resolverService = require('../services/resolverService');
const auditService = require('../services/auditService');
const logger = require('../config/logger');
const { catchAsync } = require('../middleware/errorHandler');
const { AppError } = require('../middleware/errorHandler');

class RenameController {
  
  /**
   * Rename a file
   * POST /api/excel/rename-file
   */
  renameFile = catchAsync(async (req, res) => {
    const { 
      driveId, 
      itemId, 
      driveName, 
      itemName, 
      itemPath,
      oldName, 
      newName 
    } = req.body;
    
    const auditContext = auditService.createAuditContext(req);

    if (!newName) {
      throw new AppError('newName is required', 400);
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
            // Return multiple matches for user selection
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

    const result = await renameService.renameFile(
      req.accessToken,
      resolvedDriveId,
      resolvedItemId,
      oldName,
      newName,
      auditContext
    );

    res.json({
      status: 'success',
      data: {
        file: result,
        message: `File renamed from '${result.oldName}' to '${result.newName}'`
      }
    });
  });

  /**
   * Rename a folder
   * POST /api/excel/rename-folder
   */
  renameFolder = catchAsync(async (req, res) => {
    const { 
      driveId, 
      folderId, 
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

    // Resolve drive ID if name provided
    let resolvedDriveId = driveId;
    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    // For folders, we need to implement folder resolution similar to file resolution
    let resolvedFolderId = folderId;
    if (!resolvedFolderId && folderName) {
      // Search for folder by name (similar to file search but for folders)
      const graphClient = resolverService.createGraphClient(req.accessToken);
      const matches = await this.findFoldersByName(graphClient, resolvedDriveId, folderName);
      
      if (matches.length === 0) {
        throw new AppError(`Folder '${folderName}' not found`, 404);
      }
      
      if (matches.length > 1) {
        if (folderPath) {
          const match = matches.find(m => m.path === folderPath);
          if (!match) {
            throw new AppError(`Folder '${folderName}' not found at path '${folderPath}'`, 404);
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
      throw new AppError('Could not resolve drive or folder. Please provide valid identifiers.', 400);
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

  /**
   * Rename an Excel worksheet
   * POST /api/excel/rename-sheet
   */
  renameSheet = catchAsync(async (req, res) => {
    const { 
      driveId, 
      itemId, 
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

  /**
   * Get intelligent rename suggestions
   * POST /api/excel/rename-suggestions
   */
  getRenameSuggestions = catchAsync(async (req, res) => {
    const { 
      driveId, 
      driveName, 
      oldTerm, 
      newTerm 
    } = req.body;
    
    const auditContext = auditService.createAuditContext(req);

    if (!oldTerm || !newTerm) {
      throw new AppError('oldTerm and newTerm are required', 400);
    }

    let resolvedDriveId = driveId;
    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    if (!resolvedDriveId) {
      throw new AppError('Could not resolve drive. Please provide valid drive identifier.', 400);
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

  /**
   * Batch rename multiple items
   * POST /api/excel/batch-rename
   */
  batchRename = catchAsync(async (req, res) => {
    const { 
      driveId, 
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

  /**
   * Helper method to find folders by name
   */
  async findFoldersByName(graphClient, driveId, folderName) {
    const matches = [];
    
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
