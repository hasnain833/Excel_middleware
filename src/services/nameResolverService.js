/**
 * Universal Name-to-ID Resolution Service
 * Resolves drive names, folder names, file names, and sheet names to their corresponding IDs
 * Supports deep recursive search and handles duplicate resolution
 */

const { Client } = require('@microsoft/microsoft-graph-client');
const logger = require('../config/logger');
const excelService = require('./excelService');
const auditService = require('./auditService');
const { AppError } = require('../middleware/errorHandler');

class NameResolverService {
  constructor() {
    // Enhanced caching for name resolution
    this.driveCache = new Map(); // driveName -> { id, ts }
    this.folderCache = new Map(); // driveId:folderPath -> { id, ts }
    this.fileCache = new Map(); // driveId:folderPath:fileName -> { id, ts }
    this.sheetCache = new Map(); // itemId:sheetName -> { id, ts }
    this.pathCache = new Map(); // driveId:fullPath -> { id, type, ts }
    this.ttlMs = 10 * 60 * 1000; // 10 minutes TTL
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }

  /**
   * Universal name-to-ID resolution function
   * @param {string} accessToken - Access token for Graph API
   * @param {string} driveName - Name of the drive (required)
   * @param {string} folderName - Name of the folder (optional, defaults to root)
   * @param {string} fileName - Name of the file (optional)
   * @param {string} sheetName - Name of the worksheet (optional, requires fileName)
   * @returns {Object} Resolved IDs and metadata
   */
  async resolveIdByName(accessToken, driveName, folderName = null, fileName = null, sheetName = null) {
    if (!driveName) {
      throw new AppError('driveName is required for name resolution', 400);
    }

    const resolution = {
      driveId: null,
      driveName: driveName,
      folderId: null,
      folderName: folderName,
      folderPath: null,
      itemId: null,
      fileName: fileName,
      filePath: null,
      sheetId: null,
      sheetName: sheetName,
      resolvedAt: new Date().toISOString()
    };

    try {
      // Step 1: Resolve drive name to ID
      resolution.driveId = await this.resolveDriveId(accessToken, driveName);

      // Step 2: Resolve folder name to ID (if provided)
      if (folderName) {
        const folderResult = await this.resolveFolderId(accessToken, resolution.driveId, folderName);
        resolution.folderId = folderResult.id;
        resolution.folderPath = folderResult.path;
      } else {
        resolution.folderId = 'root';
        resolution.folderPath = '/';
      }

      // Step 3: Resolve file name to ID (if provided)
      if (fileName) {
        const fileResult = await this.resolveFileId(accessToken, resolution.driveId, resolution.folderId, fileName, resolution.folderPath);
        resolution.itemId = fileResult.id;
        resolution.filePath = fileResult.path;
      }

      // Step 4: Resolve sheet name to ID (if provided and file exists)
      if (sheetName && resolution.itemId) {
        resolution.sheetId = await this.resolveSheetId(accessToken, resolution.driveId, resolution.itemId, sheetName);
      }

      // Log the resolution for audit purposes
      auditService.logSystemEvent({
        event: 'NAME_RESOLUTION_COMPLETED',
        details: {
          input: { driveName, folderName, fileName, sheetName },
          resolution: resolution,
          timestamp: resolution.resolvedAt
        }
      });

      return resolution;

    } catch (err) {
      logger.error('Name resolution failed', {
        driveName,
        folderName,
        fileName,
        sheetName,
        error: err.message
      });
      throw err;
    }
  }

  /**
   * Resolve drive name to drive ID
   */
  async resolveDriveId(accessToken, driveName) {
    const cached = this.driveCache.get(driveName);
    if (this.isFresh(cached)) {
      return cached.id;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const siteId = await excelService.getSiteId(graphClient);
      const drives = await excelService.getDrives(graphClient, siteId);
      
      const driveNameLc = String(driveName).toLowerCase();
      const match = (drives || []).find((d) => String(d.name).toLowerCase() === driveNameLc);

      if (!match) {
        const available = (drives || []).map(d => d.name);
        throw new AppError(`Drive '${driveName}' not found. Available drives: ${available.join(', ')}`, 404);
      }

      this.driveCache.set(driveName, { id: match.id, ts: Date.now() });
      return match.id;

    } catch (err) {
      throw new AppError(`Failed to resolve drive '${driveName}': ${err.message}`, err.status || 500);
    }
  }

  /**
   * Resolve folder name to folder ID with deep search
   */
  async resolveFolderId(accessToken, driveId, folderName) {
    const cacheKey = `${driveId}:${folderName}`;
    const cached = this.folderCache.get(cacheKey);
    if (this.isFresh(cached)) {
      return cached;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      
      // Search for folder recursively
      const matches = await this.recursiveSearchForFolder(graphClient, driveId, folderName);
      
      if (matches.length === 0) {
        throw new AppError(`Folder '${folderName}' not found in drive`, 404);
      }

      if (matches.length === 1) {
        const match = matches[0];
        const result = { id: match.id, path: match.path };
        this.folderCache.set(cacheKey, { ...result, ts: Date.now() });
        return result;
      }

      // Multiple matches found
      const pathOptions = matches.map((match, index) => `${index + 1}. ${match.path}`).join('\n');
      const error = new AppError(
        `Multiple folders named '${folderName}' found. Please specify the full path:\n${pathOptions}`,
        409
      );
      error.matches = matches;
      error.isMultipleMatches = true;
      throw error;

    } catch (err) {
      throw new AppError(`Failed to resolve folder '${folderName}': ${err.message}`, err.status || 500);
    }
  }

  /**
   * Resolve file name to file ID with folder context
   */
  async resolveFileId(accessToken, driveId, folderId, fileName, folderPath = '/') {
    const cacheKey = `${driveId}:${folderPath}:${fileName}`;
    const cached = this.fileCache.get(cacheKey);
    if (this.isFresh(cached)) {
      return cached;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      
      // First try to find in specified folder
      if (folderId && folderId !== 'root') {
        const folderItems = await graphClient
          .api(`/drives/${driveId}/items/${folderId}/children`)
          .select('id,name,folder')
          .top(999)
          .get();

        const fileNameLc = String(fileName).toLowerCase();
        const directMatch = (folderItems.value || []).find(
          item => !item.folder && String(item.name).toLowerCase() === fileNameLc
        );

        if (directMatch) {
          const result = { 
            id: directMatch.id, 
            path: `${folderPath}${folderPath.endsWith('/') ? '' : '/'}${directMatch.name}` 
          };
          this.fileCache.set(cacheKey, { ...result, ts: Date.now() });
          return result;
        }
      }

      // If not found in specific folder, search recursively from root
      const matches = await this.recursiveSearchForFile(graphClient, driveId, fileName);
      
      if (matches.length === 0) {
        throw new AppError(`File '${fileName}' not found in drive`, 404);
      }

      if (matches.length === 1) {
        const match = matches[0];
        const result = { id: match.id, path: match.path };
        this.fileCache.set(cacheKey, { ...result, ts: Date.now() });
        return result;
      }

      // Multiple matches found
      const pathOptions = matches.map((match, index) => `${index + 1}. ${match.path}`).join('\n');
      const error = new AppError(
        `Multiple files named '${fileName}' found. Please specify the full path:\n${pathOptions}`,
        409
      );
      error.matches = matches;
      error.isMultipleMatches = true;
      throw error;

    } catch (err) {
      throw new AppError(`Failed to resolve file '${fileName}': ${err.message}`, err.status || 500);
    }
  }

  /**
   * Resolve sheet name to sheet ID
   */
  async resolveSheetId(accessToken, driveId, itemId, sheetName) {
    const cacheKey = `${itemId}:${sheetName}`;
    const cached = this.sheetCache.get(cacheKey);
    if (this.isFresh(cached)) {
      return cached.id;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const worksheets = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
        .get();

      const sheetNameLc = String(sheetName).toLowerCase();
      const match = (worksheets.value || []).find(
        ws => String(ws.name).toLowerCase() === sheetNameLc
      );

      if (!match) {
        const available = (worksheets.value || []).map(ws => ws.name);
        throw new AppError(
          `Sheet '${sheetName}' not found. Available sheets: ${available.join(', ')}`,
          404
        );
      }

      this.sheetCache.set(cacheKey, { id: match.id, ts: Date.now() });
      return match.id;

    } catch (err) {
      throw new AppError(`Failed to resolve sheet '${sheetName}': ${err.message}`, err.status || 500);
    }
  }

  /**
   * Recursively search for folders
   */
  async recursiveSearchForFolder(graphClient, driveId, folderName, currentPath = '', folderId = 'root', depth = 0, maxDepth = 20) {
    if (depth > maxDepth) {
      logger.warn(`Maximum search depth reached for folder '${folderName}'`, { driveId, currentPath });
      return [];
    }

    const matches = [];
    const folderNameLc = String(folderName).toLowerCase();
    
    try {
      const resp = await graphClient
        .api(`/drives/${driveId}/items/${folderId}/children`)
        .select('id,name,folder,parentReference')
        .top(999)
        .get();

      const items = resp.value || [];
      
      // Check for folder matches in current directory
      for (const item of items) {
        if (item.folder && String(item.name).toLowerCase() === folderNameLc) {
          const fullPath = currentPath ? `${currentPath}/${item.name}` : `/${item.name}`;
          matches.push({
            id: item.id,
            name: item.name,
            path: fullPath,
            parentId: folderId
          });
        }
      }
      
      // Recursively search in subfolders
      const subfolders = items.filter(item => item.folder);
      for (const subfolder of subfolders) {
        const subPath = currentPath ? `${currentPath}/${subfolder.name}` : `/${subfolder.name}`;
        
        try {
          const subMatches = await this.recursiveSearchForFolder(
            graphClient, 
            driveId, 
            folderName, 
            subPath, 
            subfolder.id, 
            depth + 1, 
            maxDepth
          );
          matches.push(...subMatches);
        } catch (subErr) {
          logger.warn(`Failed to search in subfolder: ${subPath}`, { error: subErr.message });
        }
      }
      
    } catch (err) {
      logger.error(`Error during folder search`, { 
        driveId, 
        folderId, 
        currentPath, 
        depth, 
        error: err.message 
      });
    }
    
    return matches;
  }

  /**
   * Recursively search for files (reuse existing implementation)
   */
  async recursiveSearchForFile(graphClient, driveId, fileName, currentPath = '', folderId = 'root', depth = 0, maxDepth = 20) {
    if (depth > maxDepth) {
      logger.warn(`Maximum search depth reached for file '${fileName}'`, { driveId, currentPath });
      return [];
    }

    const matches = [];
    const fileNameLc = String(fileName).toLowerCase();
    
    try {
      const resp = await graphClient
        .api(`/drives/${driveId}/items/${folderId}/children`)
        .select('id,name,folder,parentReference')
        .top(999)
        .get();

      const items = resp.value || [];
      
      // Check for file matches in current folder
      for (const item of items) {
        if (!item.folder && String(item.name).toLowerCase() === fileNameLc) {
          const fullPath = currentPath ? `${currentPath}/${item.name}` : `/${item.name}`;
          matches.push({
            id: item.id,
            name: item.name,
            path: fullPath,
            parentId: folderId
          });
        }
      }
      
      // Recursively search in subfolders
      const folders = items.filter(item => item.folder);
      for (const folder of folders) {
        const subPath = currentPath ? `${currentPath}/${folder.name}` : `/${folder.name}`;
        
        try {
          const subMatches = await this.recursiveSearchForFile(
            graphClient, 
            driveId, 
            fileName, 
            subPath, 
            folder.id, 
            depth + 1, 
            maxDepth
          );
          matches.push(...subMatches);
        } catch (subErr) {
          logger.warn(`Failed to search in subfolder: ${subPath}`, { error: subErr.message });
        }
      }
      
    } catch (err) {
      logger.error(`Error during file search`, { 
        driveId, 
        folderId, 
        currentPath, 
        depth, 
        error: err.message 
      });
    }
    
    return matches;
  }

  /**
   * Resolve by full path (e.g., "/Folder1/Subfolder/file.xlsx")
   */
  async resolveByPath(accessToken, driveName, fullPath) {
    if (!fullPath || !fullPath.startsWith('/')) {
      throw new AppError('Full path must start with /', 400);
    }

    const pathParts = fullPath.split('/').filter(part => part.length > 0);
    if (pathParts.length === 0) {
      throw new AppError('Invalid path provided', 400);
    }

    const driveId = await this.resolveDriveId(accessToken, driveName);
    
    let currentFolderId = 'root';
    let currentPath = '';
    
    // Navigate through folders
    for (let i = 0; i < pathParts.length - 1; i++) {
      const folderName = pathParts[i];
      currentPath += `/${folderName}`;
      
      const folderResult = await this.findItemInFolder(accessToken, driveId, currentFolderId, folderName, true);
      if (!folderResult) {
        throw new AppError(`Folder '${folderName}' not found in path '${currentPath}'`, 404);
      }
      currentFolderId = folderResult.id;
    }
    
    // Find the final item (file or folder)
    const finalItemName = pathParts[pathParts.length - 1];
    const finalItem = await this.findItemInFolder(accessToken, driveId, currentFolderId, finalItemName, false);
    
    if (!finalItem) {
      throw new AppError(`Item '${finalItemName}' not found in path '${fullPath}'`, 404);
    }
    
    return {
      driveId,
      itemId: finalItem.id,
      itemName: finalItem.name,
      fullPath,
      isFolder: finalItem.folder !== undefined
    };
  }

  /**
   * Find item in specific folder
   */
  async findItemInFolder(accessToken, driveId, folderId, itemName, foldersOnly = false) {
    try {
      const graphClient = this.createGraphClient(accessToken);
      const resp = await graphClient
        .api(`/drives/${driveId}/items/${folderId}/children`)
        .select('id,name,folder')
        .top(999)
        .get();

      const itemNameLc = String(itemName).toLowerCase();
      return (resp.value || []).find(item => {
        const nameMatch = String(item.name).toLowerCase() === itemNameLc;
        if (foldersOnly) {
          return nameMatch && item.folder;
        }
        return nameMatch;
      });

    } catch (err) {
      logger.error('Failed to find item in folder', { driveId, folderId, itemName, error: err.message });
      return null;
    }
  }

  /**
   * Check if cached item is fresh
   */
  isFresh(cached) {
    return cached && (Date.now() - cached.ts) < this.ttlMs;
  }

  /**
   * Clear all caches
   */
  clearCache() {
    this.driveCache.clear();
    this.folderCache.clear();
    this.fileCache.clear();
    this.sheetCache.clear();
    this.pathCache.clear();
    logger.info('Name resolver cache cleared');
  }

  /**
   * Get cache statistics
   */
  getCacheStats() {
    return {
      drives: this.driveCache.size,
      folders: this.folderCache.size,
      files: this.fileCache.size,
      sheets: this.sheetCache.size,
      paths: this.pathCache.size,
      ttlMinutes: this.ttlMs / (60 * 1000)
    };
  }
}

module.exports = new NameResolverService();
