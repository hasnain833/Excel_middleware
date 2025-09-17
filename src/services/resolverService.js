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
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }

  isFresh(entry) {
    return entry && (Date.now() - entry.ts) < this.ttlMs;
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
      // As per requirement, list children of root then find by exact name
      const resp = await graphClient
        .api(`/drives/${driveId}/root/children`)
        .select('id,name')
        .top(999)
        .get();

      const items = resp.value || [];
      const available = items.map(it => it.name);
      const itemNameLc = String(itemName).toLowerCase();
      const match = items.find((it) => String(it.name).toLowerCase() === itemNameLc);
      if (!match) {
        const msg = `File not found in this drive. Available items: ${JSON.stringify(available)}`;
        logger.warn(msg, { driveId });
        throw new AppError(msg, 404);
      }

      this.itemCache.set(cacheKey, { id: match.id, ts: Date.now() });
      return match.id;
    } catch (err) {
      if (!err.status) logger.error('Failed resolving itemId by name', { driveId, itemName, error: err.message });
      throw err;
    }
  }

  parseSheetAndAddress(maybeQualifiedRange) {
    // Supports formats: 'Sheet1!A1:D10' or just 'A1:D10'
    // Returns { sheetName, address }
    const str = String(maybeQualifiedRange || '');
    const bangIdx = str.indexOf('!');
    if (bangIdx > 0) {
      const sheetName = str.substring(0, bangIdx);
      const address = str.substring(bangIdx + 1);
      return { sheetName, address };
    }
    return { sheetName: null, address: str };
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
