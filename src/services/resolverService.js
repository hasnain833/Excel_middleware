/**
 * Resolver Service
 * Resolves driveName -> driveId, itemName -> itemId, and worksheetName from range
 * Implements lightweight in-memory caches to minimize Graph lookups
 */

const { Client } = require('@microsoft/microsoft-graph-client');
const logger = require('../config/logger');
const excelService = require('./excelService');
const { AppError } = require('../middleware/errorHandler');

class ResolverService {
  constructor() {
    // Simple in-memory caches
    this.driveCache = new Map(); // key: driveName -> { id, ts }
    this.itemCache = new Map();  // key: `${driveId}:${itemName}` -> { id, ts }
    this.worksheetCache = new Map(); // key: `${itemId}:${worksheetName}` -> { id, ts }
    this.ttlMs = 10 * 60 * 1000; // 10 minutes TTL
    
    // Cache for recursive search results to improve performance
    this.recursiveSearchCache = new Map(); // key: `${driveId}:${itemName}` -> { matches: [], ts }
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }

  async resolveDriveIdByName(accessToken, driveName) {
    if (!driveName) throw new AppError('driveName is required', 400);

    const cached = this.driveCache.get(driveName);
    if (this.isFresh(cached)) {
      return cached.id;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const siteId = await excelService.getSiteId(graphClient);
      const drives = await excelService.getDrives(graphClient, siteId);
      const available = (drives || []).map(d => d.name);
      const driveNameLc = String(driveName).toLowerCase();
      const match = (drives || []).find((d) => String(d.name).toLowerCase() === driveNameLc);

      if (!match) {
        const msg = `Drive not found. Available drives: ${JSON.stringify(available)}`;
        logger.warn(msg);
        throw new AppError(msg, 404);
      }

      this.driveCache.set(driveName, { id: match.id, ts: Date.now() });
      return match.id;
    } catch (err) {
      if (!err.status) logger.error('Failed resolving driveId by name', { driveName, error: err.message });
      throw err;
    }
  }

  async resolveItemIdByName(accessToken, driveId, itemName) {
    if (!driveId) throw new AppError('driveId is required', 400);
    if (!itemName) throw new AppError('itemName is required', 400);

    const cacheKey = `${driveId}:${itemName}`;
    const cached = this.itemCache.get(cacheKey);
    if (this.isFresh(cached)) {
      return cached.id;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      
      // First, try to find in root folder for backward compatibility
      const rootResp = await graphClient
        .api(`/drives/${driveId}/root/children`)
        .select('id,name,folder')
        .top(999)
        .get();

      const rootItems = rootResp.value || [];
      const itemNameLc = String(itemName).toLowerCase();
      const rootMatch = rootItems.find((it) => String(it.name).toLowerCase() === itemNameLc && !it.folder);
      
      if (rootMatch) {
        // Found in root folder - cache and return immediately
        this.itemCache.set(cacheKey, { id: rootMatch.id, ts: Date.now() });
        return rootMatch.id;
      }

      // If not found in root, perform recursive search
      logger.info(`File '${itemName}' not found in root folder. Starting recursive search...`, { driveId });
      const matches = await this.recursiveSearchForFile(graphClient, driveId, itemName);
      
      if (matches.length === 0) {
        const rootAvailable = rootItems.filter(it => !it.folder).map(it => it.name);
        const msg = `File '${itemName}' not found in drive. Files in root folder: ${JSON.stringify(rootAvailable)}`;
        logger.warn(msg, { driveId });
        throw new AppError(msg, 404);
      }
      
      if (matches.length === 1) {
        // Single match found - cache and return
        const match = matches[0];
        this.itemCache.set(cacheKey, { id: match.id, ts: Date.now() });
        logger.info(`File '${itemName}' found at path: ${match.path}`, { driveId, itemId: match.id });
        return match.id;
      }
      
      // Multiple matches found - throw error with selection options
      const pathOptions = matches.map((match, index) => `${index + 1}. ${match.path}`).join('\n');
      const msg = `Multiple files named '${itemName}' found. Please specify which file you want to use by providing the file path or use the selection endpoint:\n${pathOptions}`;
      logger.warn(msg, { driveId, matchCount: matches.length });
      
      // Create a special error that includes the matches for potential UI handling
      const error = new AppError(msg, 409); // 409 Conflict for multiple choices
      error.matches = matches;
      error.isMultipleMatches = true;
      throw error;
      
    } catch (err) {
      if (!err.status) logger.error('Failed resolving itemId by name', { driveId, itemName, error: err.message });
      throw err;
    }
  }

  /**
   * Recursively search for files with the given name in all folders and subfolders
   * @param {Object} graphClient - Microsoft Graph client
   * @param {string} driveId - Drive ID to search in
   * @param {string} fileName - Name of the file to search for
   * @param {string} currentPath - Current folder path (for building full paths)
   * @param {string} folderId - Current folder ID (defaults to 'root')
   * @param {number} depth - Current search depth (for preventing infinite recursion)
   * @param {number} maxDepth - Maximum search depth (default: 20)
   * @returns {Array} Array of matching files with their paths and IDs
   */
  async recursiveSearchForFile(graphClient, driveId, fileName, currentPath = '', folderId = 'root', depth = 0, maxDepth = 20) {
    // Prevent infinite recursion
    if (depth > maxDepth) {
      logger.warn(`Maximum search depth (${maxDepth}) reached. Stopping recursive search.`, { driveId, fileName, currentPath });
      return [];
    }

    const matches = [];
    const fileNameLc = String(fileName).toLowerCase();
    
    try {
      // Get all items in current folder
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
          logger.debug(`Found file match: ${fullPath}`, { driveId, itemId: item.id });
        }
      }
      
      // Recursively search in subfolders
      const folders = items.filter(item => item.folder);
      for (const folder of folders) {
        const subPath = currentPath ? `${currentPath}/${folder.name}` : `/${folder.name}`;
        logger.debug(`Searching in folder: ${subPath}`, { driveId, folderId: folder.id, depth });
        
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
          // Log but don't fail the entire search if one subfolder fails
          logger.warn(`Failed to search in subfolder: ${subPath}`, { 
            driveId, 
            folderId: folder.id, 
            error: subErr.message 
          });
        }
      }
      
    } catch (err) {
      logger.error(`Error during recursive search in folder`, { 
        driveId, 
        folderId, 
        currentPath, 
        depth, 
        error: err.message 
      });
      // Don't throw - return partial results
    }
    
    return matches;
  }

  /**
   * Resolve item ID by providing a specific path when multiple files exist
   * @param {string} accessToken - Access token for Graph API
   * @param {string} driveId - Drive ID
   * @param {string} itemName - File name
   * @param {string} itemPath - Full path to the specific file
   * @returns {string} Item ID
   */
  async resolveItemIdByPath(accessToken, driveId, itemName, itemPath) {
    if (!driveId) throw new AppError('driveId is required', 400);
    if (!itemName) throw new AppError('itemName is required', 400);
    if (!itemPath) throw new AppError('itemPath is required', 400);

    const cacheKey = `${driveId}:${itemPath}`;
    const cached = this.itemCache.get(cacheKey);
    if (this.isFresh(cached)) {
      return cached.id;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      
      // Search for the file and find the one with matching path
      const matches = await this.recursiveSearchForFile(graphClient, driveId, itemName);
      const match = matches.find(m => m.path === itemPath);
      
      if (!match) {
        const availablePaths = matches.map(m => m.path);
        const msg = `File '${itemName}' not found at path '${itemPath}'. Available paths: ${JSON.stringify(availablePaths)}`;
        logger.warn(msg, { driveId });
        throw new AppError(msg, 404);
      }
      
      this.itemCache.set(cacheKey, { id: match.id, ts: Date.now() });
      logger.info(`File '${itemName}' resolved by path: ${itemPath}`, { driveId, itemId: match.id });
      return match.id;
      
    } catch (err) {
      if (!err.status) logger.error('Failed resolving itemId by path', { driveId, itemName, itemPath, error: err.message });
      throw err;
    }
  }

  isFresh(cached) {
    return cached && (Date.now() - cached.ts) < this.ttlMs;
  }

  /**
   * Parse sheet name and address from range specification
   * @param {string} rangeSpec - Range specification like "Sheet1!A2:D20" or "A2:D20"
   * @returns {Object} Object with sheetName and address properties
   */
  parseSheetAndAddress(rangeSpec) {
    if (!rangeSpec) {
      return { sheetName: null, address: null };
    }

    const parts = rangeSpec.split('!');
    if (parts.length === 2) {
      return {
        sheetName: parts[0].replace(/'/g, ''), // Remove quotes if present
        address: parts[1]
      };
    } else {
      return {
        sheetName: null,
        address: rangeSpec
      };
    }
  }

  async resolveWorksheetIdByName(accessToken, driveId, itemId, worksheetName) {
    if (!worksheetName) {
      const msg = 'worksheetName is required to resolve worksheetId';
      throw new AppError(msg, 400);
    }

    const cacheKey = `${itemId}:${worksheetName}`;
    const cached = this.worksheetCache.get(cacheKey);
    if (this.isFresh(cached)) {
      return cached.id;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const resp = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
        .get();

      const match = (resp.value || []).find((ws) => ws.name === worksheetName);
      if (!match) {
        const msg = `Worksheet not found: ${worksheetName}`;
        throw new AppError(msg, 404);
      }

      this.worksheetCache.set(cacheKey, { id: match.id, ts: Date.now() });
      return match.id;
    } catch (err) {
      if (!err.status) logger.error('Failed resolving worksheetId by name', { driveId, itemId, worksheetName, error: err.message });
      throw err;
    }
  }
}

module.exports = new ResolverService();
