/**
 * Audit Logger Middleware
 * Logs all Excel write operations to audit-log.json file
 */

const fs = require('fs').promises;
const path = require('path');
const logger = require('../config/logger');
const { v4: uuidv4 } = require('uuid');

class AuditLogger {
    constructor() {
        this.auditLogPath = path.join(__dirname, '../../audit-log.json');
        this.ensureAuditLogExists();
    }

    /**
     * Ensure audit log file exists
     */
    async ensureAuditLogExists() {
        try {
            await fs.access(this.auditLogPath);
        } catch (error) {
            // File doesn't exist, create it with empty array
            await fs.writeFile(this.auditLogPath, '[]', 'utf8');
            logger.info('Created audit log file:', this.auditLogPath);
        }
    }

    /**
     * Read existing audit log entries
     * @returns {Array} Existing log entries
     */
    async readAuditLog() {
        try {
            const data = await fs.readFile(this.auditLogPath, 'utf8');
            return JSON.parse(data);
        } catch (error) {
            logger.error('Failed to read audit log:', error);
            return [];
        }
    }

    /**
     * Write audit log entries back to file
     * @param {Array} entries - Log entries to write
     */
    async writeAuditLog(entries) {
        try {
            await fs.writeFile(this.auditLogPath, JSON.stringify(entries, null, 2), 'utf8');
        } catch (error) {
            logger.error('Failed to write audit log:', error);
            throw error;
        }
    }

    /**
     * Add new audit entry
     * @param {Object} entry - Audit entry to add
     */
    async addAuditEntry(entry) {
        try {
            const entries = await this.readAuditLog();
            entries.push(entry);
            
            // Keep only last 1000 entries to prevent file from growing too large
            if (entries.length > 1000) {
                entries.splice(0, entries.length - 1000);
            }
            
            await this.writeAuditLog(entries);
            logger.debug('Audit entry added', { entryId: entry.id });
        } catch (error) {
            logger.error('Failed to add audit entry:', error);
            // Don't throw error to avoid breaking the main operation
        }
    }

    /**
     * Read current values from Excel before write operation
     * @param {Object} graphClient - Microsoft Graph client
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @param {string} worksheetId - Worksheet ID
     * @param {string} range - Range to read
     * @returns {Array|null} Current values or null if read fails
     */
    async readCurrentValues(graphClient, driveId, itemId, worksheetId, range) {
        try {
            const response = await graphClient
                .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`)
                .get();
            return response.values;
        } catch (error) {
            logger.warn('Could not read current values for audit:', error.message);
            return null;
        }
    }

    /**
     * Get file name from Graph API
     * @param {Object} graphClient - Microsoft Graph client
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @returns {string} File name or itemId if not found
     */
    async getFileName(graphClient, driveId, itemId) {
        try {
            const response = await graphClient
                .api(`/drives/${driveId}/items/${itemId}`)
                .select('name')
                .get();
            return response.name;
        } catch (error) {
            logger.warn('Could not get file name for audit:', error.message);
            return itemId; // Fallback to itemId
        }
    }

    /**
     * Express middleware function for audit logging
     * @param {Object} req - Express request
     * @param {Object} res - Express response
     * @param {Function} next - Next middleware
     */
    middleware() {
        return async (req, res, next) => {
            // Only log write operations
            if (req.method !== 'POST' && req.method !== 'PATCH') {
                return next();
            }

            // Only log Excel write endpoints
            if (!req.path.includes('/excel/write') && !req.path.includes('/excel/add-table-rows')) {
                return next();
            }

            const startTime = Date.now();
            const auditId = uuidv4();

            // Store original res.json to intercept response
            const originalJson = res.json;
            let responseData = null;
            let statusCode = null;

            res.json = function(data) {
                responseData = data;
                statusCode = res.statusCode;
                return originalJson.call(this, data);
            };

            // Store request data for audit
            const auditData = {
                id: auditId,
                timestamp: new Date().toISOString(),
                user: req.headers['x-user-id'] || 'anonymous',
                operation: req.path.includes('/add-table-rows') ? 'ADD_TABLE_ROWS' : 'WRITE_RANGE',
                driveId: req.body.driveId,
                itemId: req.body.itemId,
                worksheetId: req.body.worksheetId,
                range: req.body.range,
                tableName: req.body.tableName,
                newValues: req.body.values || req.body.rows,
                oldValues: null,
                fileName: null,
                success: false,
                error: null,
                duration: 0,
                ipAddress: req.ip || req.connection.remoteAddress,
                userAgent: req.get('User-Agent')
            };

            // Try to get current values and file name before the operation
            if (req.accessToken && req.body.driveId && req.body.itemId) {
                try {
                    const { Client } = require('@microsoft/microsoft-graph-client');
                    const graphClient = Client.init({
                        authProvider: (done) => {
                            done(null, req.accessToken);
                        }
                    });

                    // Get file name
                    auditData.fileName = await this.getFileName(graphClient, req.body.driveId, req.body.itemId);

                    // Get current values for range operations
                    if (req.body.range && req.body.worksheetId) {
                        auditData.oldValues = await this.readCurrentValues(
                            graphClient, 
                            req.body.driveId, 
                            req.body.itemId, 
                            req.body.worksheetId, 
                            req.body.range
                        );
                    }
                } catch (error) {
                    logger.warn('Could not fetch pre-operation data for audit:', error.message);
                }
            }

            // Continue with the request
            next();

            // Log after response is sent
            res.on('finish', async () => {
                try {
                    auditData.duration = Date.now() - startTime;
                    auditData.success = statusCode >= 200 && statusCode < 300;
                    
                    if (!auditData.success && responseData?.error) {
                        auditData.error = responseData.error.message || 'Unknown error';
                    }

                    await this.addAuditEntry(auditData);
                    
                    logger.info('Excel operation audited', {
                        auditId: auditData.id,
                        operation: auditData.operation,
                        user: auditData.user,
                        fileName: auditData.fileName,
                        success: auditData.success,
                        duration: auditData.duration
                    });
                } catch (error) {
                    logger.error('Failed to complete audit logging:', error);
                }
            });
        };
    }

    /**
     * Get audit log entries with optional filtering
     * @param {Object} filters - Filter options
     * @returns {Array} Filtered audit entries
     */
    async getAuditEntries(filters = {}) {
        try {
            const entries = await this.readAuditLog();
            let filtered = entries;

            // Apply filters
            if (filters.user) {
                filtered = filtered.filter(entry => entry.user === filters.user);
            }

            if (filters.operation) {
                filtered = filtered.filter(entry => entry.operation === filters.operation);
            }

            if (filters.fileName) {
                filtered = filtered.filter(entry => 
                    entry.fileName && entry.fileName.toLowerCase().includes(filters.fileName.toLowerCase())
                );
            }

            if (filters.startDate) {
                const startDate = new Date(filters.startDate);
                filtered = filtered.filter(entry => new Date(entry.timestamp) >= startDate);
            }

            if (filters.endDate) {
                const endDate = new Date(filters.endDate);
                filtered = filtered.filter(entry => new Date(entry.timestamp) <= endDate);
            }

            if (filters.success !== undefined) {
                filtered = filtered.filter(entry => entry.success === filters.success);
            }

            // Sort by timestamp (newest first)
            filtered.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

            // Apply limit
            if (filters.limit) {
                filtered = filtered.slice(0, parseInt(filters.limit));
            }

            return filtered;
        } catch (error) {
            logger.error('Failed to get audit entries:', error);
            return [];
        }
    }
}

module.exports = new AuditLogger();
