/**
 * Rename Service
 * Handles renaming of files, folders, and Excel sheets
 * Includes intelligent rename suggestions for related items
 */

const { Client } = require('@microsoft/microsoft-graph-client');
const logger = require('../config/logger');
const resolverService = require('./resolverService');
const auditService = require('./auditService');
const { AppError } = require('../middleware/errorHandler');

class RenameService {
  constructor() {
    // Cache for rename suggestions to avoid duplicate API calls
    this.suggestionCache = new Map();
    this.ttlMs = 5 * 60 * 1000; // 5 minutes TTL for suggestions
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }

  /**
   * Rename an Excel file
   */
  async renameFile(accessToken, driveId, itemId, oldName, newName, auditContext) {
    if (!driveId || !itemId || !newName) {
      throw new AppError('driveId, itemId, and newName are required', 400);
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      
      // Validate file exists and get current info
      const currentFile = await graphClient
        .api(`/drives/${driveId}/items/${itemId}`)
        .get();

      if (!currentFile) {
        throw new AppError('File not found', 404);
      }

      // Perform rename operation
      const updatedFile = await graphClient
        .api(`/drives/${driveId}/items/${itemId}`)
        .patch({ name: newName });

      // Log the rename operation
      auditService.logSystemEvent({
        event: 'FILE_RENAMED',
        details: {
          driveId,
          itemId,
          oldName: currentFile.name,
          newName: updatedFile.name,
          path: currentFile.parentReference?.path || '/',
          requestedBy: auditContext.user,
          timestamp: new Date().toISOString()
        }
      });

      logger.info(`File renamed successfully`, {
        driveId,
        itemId,
        oldName: currentFile.name,
        newName: updatedFile.name
      });

      return {
        id: updatedFile.id,
        oldName: currentFile.name,
        newName: updatedFile.name,
        path: updatedFile.parentReference?.path || '/',
        lastModified: updatedFile.lastModifiedDateTime
      };

    } catch (err) {
      logger.error('Failed to rename file', {
        driveId,
        itemId,
        oldName,
        newName,
        error: err.message
      });
      
      if (err.code === 'nameAlreadyExists') {
        throw new AppError(`A file named '${newName}' already exists in this location`, 409);
      }
      if (err.code === 'accessDenied') {
        throw new AppError('Access denied. You may not have permission to rename this file', 403);
      }
      if (err.code === 'resourceLocked') {
        throw new AppError('File is currently locked and cannot be renamed', 423);
      }
      
      throw err;
    }
  }

  /**
   * Rename a folder
   */
  async renameFolder(accessToken, driveId, folderId, oldName, newName, auditContext) {
    if (!driveId || !folderId || !newName) {
      throw new AppError('driveId, folderId, and newName are required', 400);
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      
      // Validate folder exists
      const currentFolder = await graphClient
        .api(`/drives/${driveId}/items/${folderId}`)
        .get();

      if (!currentFolder || !currentFolder.folder) {
        throw new AppError('Folder not found', 404);
      }

      // Perform rename operation
      const updatedFolder = await graphClient
        .api(`/drives/${driveId}/items/${folderId}`)
        .patch({ name: newName });

      // Log the rename operation
      auditService.logSystemEvent({
        event: 'FOLDER_RENAMED',
        details: {
          driveId,
          folderId,
          oldName: currentFolder.name,
          newName: updatedFolder.name,
          path: currentFolder.parentReference?.path || '/',
          requestedBy: auditContext.user,
          timestamp: new Date().toISOString()
        }
      });

      logger.info(`Folder renamed successfully`, {
        driveId,
        folderId,
        oldName: currentFolder.name,
        newName: updatedFolder.name
      });

      return {
        id: updatedFolder.id,
        oldName: currentFolder.name,
        newName: updatedFolder.name,
        path: updatedFolder.parentReference?.path || '/',
        lastModified: updatedFolder.lastModifiedDateTime
      };

    } catch (err) {
      logger.error('Failed to rename folder', {
        driveId,
        folderId,
        oldName,
        newName,
        error: err.message
      });
      
      if (err.code === 'nameAlreadyExists') {
        throw new AppError(`A folder named '${newName}' already exists in this location`, 409);
      }
      if (err.code === 'accessDenied') {
        throw new AppError('Access denied. You may not have permission to rename this folder', 403);
      }
      
      throw err;
    }
  }

  /**
   * Rename an Excel worksheet
   */
  async renameSheet(accessToken, driveId, itemId, oldSheetName, newSheetName, auditContext) {
    if (!driveId || !itemId || !oldSheetName || !newSheetName) {
      throw new AppError('driveId, itemId, oldSheetName, and newSheetName are required', 400);
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      
      // Get file info for logging
      const fileInfo = await graphClient
        .api(`/drives/${driveId}/items/${itemId}`)
        .get();

      // Get current worksheet info
      const worksheets = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
        .get();

      const targetSheet = worksheets.value.find(ws => ws.name === oldSheetName);
      if (!targetSheet) {
        throw new AppError(`Worksheet '${oldSheetName}' not found`, 404);
      }

      // Perform rename operation
      const updatedSheet = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${targetSheet.id}`)
        .patch({ name: newSheetName });

      // Log the rename operation
      auditService.logSystemEvent({
        event: 'WORKSHEET_RENAMED',
        details: {
          driveId,
          itemId,
          fileName: fileInfo.name,
          worksheetId: targetSheet.id,
          oldSheetName,
          newSheetName: updatedSheet.name,
          requestedBy: auditContext.user,
          timestamp: new Date().toISOString()
        }
      });

      logger.info(`Worksheet renamed successfully`, {
        driveId,
        itemId,
        fileName: fileInfo.name,
        oldSheetName,
        newSheetName: updatedSheet.name
      });

      return {
        worksheetId: updatedSheet.id,
        oldSheetName,
        newSheetName: updatedSheet.name,
        fileName: fileInfo.name,
        fileId: itemId
      };

    } catch (err) {
      logger.error('Failed to rename worksheet', {
        driveId,
        itemId,
        oldSheetName,
        newSheetName,
        error: err.message
      });
      
      if (err.code === 'nameAlreadyExists') {
        throw new AppError(`A worksheet named '${newSheetName}' already exists in this workbook`, 409);
      }
      if (err.code === 'accessDenied') {
        throw new AppError('Access denied. You may not have permission to rename worksheets in this file', 403);
      }
      
      throw err;
    }
  }

  /**
   * Find related files and folders that might need renaming
   * Used for intelligent rename suggestions
   */
  async findRelatedItems(accessToken, driveId, oldTerm, newTerm) {
    const cacheKey = `${driveId}:${oldTerm}`;
    const cached = this.suggestionCache.get(cacheKey);
    
    if (cached && (Date.now() - cached.ts) < this.ttlMs) {
      return cached.items;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const relatedItems = [];

      // Search for files and folders containing the old term
      const searchResults = await this.recursiveSearch(graphClient, driveId, oldTerm);
      
      for (const item of searchResults) {
        if (item.name.toLowerCase().includes(oldTerm.toLowerCase())) {
          const suggestedName = item.name.replace(
            new RegExp(oldTerm, 'gi'), 
            newTerm
          );
          
          relatedItems.push({
            id: item.id,
            currentName: item.name,
            suggestedName,
            path: item.path,
            type: item.folder ? 'folder' : 'file',
            parentId: item.parentId
          });
        }
      }

      // Cache the results
      this.suggestionCache.set(cacheKey, {
        items: relatedItems,
        ts: Date.now()
      });

      return relatedItems;

    } catch (err) {
      logger.error('Failed to find related items', {
        driveId,
        oldTerm,
        newTerm,
        error: err.message
      });
      return [];
    }
  }

  /**
   * Recursive search for items containing a term
   */
  async recursiveSearch(graphClient, driveId, searchTerm, folderId = 'root', currentPath = '', depth = 0, maxDepth = 10) {
    if (depth > maxDepth) return [];

    const items = [];
    
    try {
      const resp = await graphClient
        .api(`/drives/${driveId}/items/${folderId}/children`)
        .select('id,name,folder,parentReference')
        .top(999)
        .get();

      for (const item of resp.value || []) {
        const itemPath = currentPath ? `${currentPath}/${item.name}` : `/${item.name}`;
        
        if (item.name.toLowerCase().includes(searchTerm.toLowerCase())) {
          items.push({
            id: item.id,
            name: item.name,
            path: itemPath,
            folder: item.folder,
            parentId: folderId
          });
        }

        // Recursively search subfolders
        if (item.folder) {
          const subItems = await this.recursiveSearch(
            graphClient,
            driveId,
            searchTerm,
            item.id,
            itemPath,
            depth + 1,
            maxDepth
          );
          items.push(...subItems);
        }
      }
    } catch (err) {
      logger.warn(`Failed to search in folder ${folderId}`, { error: err.message });
    }

    return items;
  }

  /**
   * Batch rename multiple items
   */
  async batchRename(accessToken, driveId, renameOperations, auditContext) {
    const results = [];
    const errors = [];

    for (let i = 0; i < renameOperations.length; i++) {
      const operation = renameOperations[i];
      
      try {
        let result;
        
        switch (operation.type) {
          case 'file':
            result = await this.renameFile(
              accessToken,
              driveId,
              operation.itemId,
              operation.oldName,
              operation.newName,
              auditContext
            );
            break;
            
          case 'folder':
            result = await this.renameFolder(
              accessToken,
              driveId,
              operation.itemId,
              operation.oldName,
              operation.newName,
              auditContext
            );
            break;
            
          case 'sheet':
            result = await this.renameSheet(
              accessToken,
              driveId,
              operation.fileId,
              operation.oldName,
              operation.newName,
              auditContext
            );
            break;
            
          default:
            throw new Error(`Unknown operation type: ${operation.type}`);
        }

        results.push({
          index: i,
          operation: operation.type,
          success: true,
          data: result
        });

      } catch (error) {
        logger.error(`Batch rename operation ${i} failed:`, error);
        errors.push({
          index: i,
          operation: operation.type,
          error: error.message,
          itemId: operation.itemId || operation.fileId
        });
      }
    }

    return {
      results,
      errors,
      summary: {
        total: renameOperations.length,
        successful: results.length,
        failed: errors.length
      }
    };
  }
}

module.exports = new RenameService();
