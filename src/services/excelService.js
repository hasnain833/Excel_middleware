/**
 * Excel Service
 * Production-grade service for Excel operations using Microsoft Graph Client SDK
 * Supports Client Credentials Flow with proper site/drive discovery
 */

const { Client } = require('@microsoft/microsoft-graph-client');
const auditService = require('./auditService');
const logger = require('../config/logger');
const permissions = require('../config/permissions');

class ExcelService {
    constructor() {
        this.auditService = auditService;
        // Configuration for SharePoint site discovery
        this.hostname = process.env.SHAREPOINT_HOSTNAME || 'yourtenant.sharepoint.com';
        this.siteName = process.env.SHAREPOINT_SITE_NAME || 'Documents';
    }

    /**
     * Create Microsoft Graph Client with access token
     * @param {string} accessToken - Access token for Client Credentials Flow
     * @returns {Client} Microsoft Graph Client instance
     */
    createGraphClient(accessToken) {
        return Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
    }

    /**
     * Get SharePoint site ID from tenant domain and site name
     * @param {Client} graphClient - Graph client instance
     * @returns {Promise<string>} Site ID
     */
    async getSiteId(graphClient) {
        try {
            logger.debug('🔍 Fetching SharePoint site ID', { 
                hostname: this.hostname, 
                siteName: this.siteName 
            });

            // Get site by hostname and site path
            const siteResponse = await graphClient
                .api(`/sites/${this.hostname}:/sites/${this.siteName}`)
                .get();

            const siteId = siteResponse.id;
            logger.debug('✅ Site ID retrieved successfully', { siteId });
            
            return siteId;
        } catch (error) {
            logger.error('❌ Failed to get site ID', { 
                error: error.message,
                hostname: this.hostname,
                siteName: this.siteName
            });
            
            if (error.code === 'itemNotFound') {
                throw new Error(`SharePoint site not found: ${this.hostname}/sites/${this.siteName}`);
            } else if (error.code === 'Forbidden') {
                throw new Error('Access denied to SharePoint site. Check application permissions.');
            } else if (error.code === 'Unauthorized') {
                throw new Error('Authentication failed. Check access token.');
            }
            
            throw new Error(`Failed to retrieve site ID: ${error.message}`);
        }
    }

    /**
     * Get all drives (document libraries) in the SharePoint site
     * @param {Client} graphClient - Graph client instance
     * @param {string} siteId - SharePoint site ID
     * @returns {Promise<Array>} List of drives
     */
    async getDrives(graphClient, siteId) {
        try {
            logger.debug('🔍 Fetching drives from site', { siteId });

            const drivesResponse = await graphClient
                .api(`/sites/${siteId}/drives`)
                .get();

            const drives = drivesResponse.value.map(drive => ({
                id: drive.id,
                name: drive.name,
                description: drive.description,
                driveType: drive.driveType,
                webUrl: drive.webUrl
            }));

            logger.debug('✅ Drives retrieved successfully', { 
                driveCount: drives.length,
                drives: drives.map(d => ({ id: d.id, name: d.name }))
            });

            return drives;
        } catch (error) {
            logger.error('❌ Failed to get drives', { error: error.message, siteId });
            
            if (error.code === 'Forbidden') {
                throw new Error('Access denied to site drives. Check application permissions.');
            }
            
            throw new Error(`Failed to retrieve drives: ${error.message}`);
        }
    }

    /**
     * Get Excel workbooks from all drives
     * @param {Client} graphClient - Graph client instance
     * @param {Array} drives - List of drives
     * @returns {Promise<Array>} List of Excel workbooks
     */
    async getWorkbooksFromDrives(graphClient, drives) {
        try {
            logger.debug('🔍 Searching for Excel workbooks in drives', { driveCount: drives.length });

            const allWorkbooks = [];

            for (const drive of drives) {
                try {
                    logger.debug(`🔍 Searching drive: ${drive.name}`, { driveId: drive.id });

                    // Search for Excel files in this drive
                    const searchResponse = await graphClient
                        .api(`/drives/${drive.id}/root/search(q='.xlsx')`)
                        .filter('file ne null')
                        .top(50)
                        .get();

                    const workbooks = searchResponse.value.map(item => ({
                        id: item.id,
                        name: item.name,
                        driveId: drive.id,
                        driveName: drive.name,
                        webUrl: item.webUrl,
                        parentReference: item.parentReference,
                        lastModifiedDateTime: item.lastModifiedDateTime,
                        size: item.size,
                        createdDateTime: item.createdDateTime
                    }));

                    allWorkbooks.push(...workbooks);
                    logger.debug(`✅ Found ${workbooks.length} workbooks in drive: ${drive.name}`);

                } catch (driveError) {
                    logger.warn(`⚠️ Failed to search drive: ${drive.name}`, { 
                        driveId: drive.id, 
                        error: driveError.message 
                    });
                    // Continue with other drives even if one fails
                }
            }

            logger.debug('✅ Total workbooks found across all drives', { totalCount: allWorkbooks.length });
            return allWorkbooks;

        } catch (error) {
            logger.error('❌ Failed to get workbooks from drives', { error: error.message });
            throw new Error(`Failed to search for workbooks: ${error.message}`);
        }
    }

    /**
     * Get all accessible workbooks (main entry point)
     * @param {string} accessToken - Access token
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Array>} List of workbooks
     */
    async getWorkbooks(accessToken, auditContext) {
        try {
            const graphClient = this.createGraphClient(accessToken);
            
            // Step 1: Get SharePoint site ID
            const siteId = await this.getSiteId(graphClient);
            
            // Step 2: Get all drives in the site
            const drives = await this.getDrives(graphClient, siteId);
            
            // Step 3: Search for Excel workbooks in all drives
            const workbooks = await this.getWorkbooksFromDrives(graphClient, drives);
            
            // Step 4: Apply permission filtering
            const filteredWorkbooks = workbooks.filter(workbook => 
                permissions.canAccessWorkbook(auditContext.user, workbook.id)
            );

            // Log audit event
            auditService.logSystemEvent({
                event: 'WORKBOOKS_RETRIEVED',
                details: { 
                    totalFound: workbooks.length,
                    accessibleCount: filteredWorkbooks.length,
                    user: auditContext.user,
                    siteId,
                    driveCount: drives.length
                }
            });

            logger.info('📊 Workbooks retrieval completed', {
                totalFound: workbooks.length,
                accessibleCount: filteredWorkbooks.length,
                user: auditContext.user
            });

            return filteredWorkbooks;

        } catch (error) {
            logger.error('❌ Excel service - failed to get workbooks:', error);
            throw error;
        }
    }

    /**
     * Get worksheets in a workbook
     * @param {string} accessToken - Access token
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Array>} List of worksheets
     */
    async getWorksheets(accessToken, driveId, itemId, auditContext) {
        try {
            // Check workbook access permission
            if (!permissions.canAccessWorkbook(auditContext.user, itemId)) {
                throw new Error('Access denied to workbook');
            }

            const graphClient = this.createGraphClient(accessToken);
            
            logger.debug('🔍 Fetching worksheets', { driveId, itemId });

            const response = await graphClient
                .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
                .get();
            
            const worksheets = response.value.map(sheet => ({
                id: sheet.id,
                name: sheet.name,
                position: sheet.position,
                visibility: sheet.visibility
            }));

            // Filter worksheets based on permissions
            const filteredWorksheets = worksheets.filter(worksheet => 
                permissions.canAccessWorksheet(auditContext.user, itemId, worksheet.id)
            );

            logger.debug('✅ Worksheets retrieved successfully', { 
                workbookId: itemId,
                totalCount: worksheets.length,
                accessibleCount: filteredWorksheets.length
            });

            return filteredWorksheets;

        } catch (error) {
            logger.error('❌ Excel service - failed to get worksheets:', error);
            
            if (error.code === 'itemNotFound') {
                throw new Error('Workbook not found or not accessible');
            } else if (error.code === 'Forbidden') {
                throw new Error('Access denied to workbook');
            }
            
            throw error;
        }
    }

    /**
     * Read data from Excel range with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Range data
     */
    async readRange(params) {
        const { accessToken, driveId, itemId, worksheetId, range, auditContext } = params;

        try {
            // Check permissions
            const hasPermission = permissions.canReadRange(
                auditContext.user, 
                itemId, 
                worksheetId, 
                range
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    range: range,
                    requestedPermission: 'READ',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Read access denied: ${hasPermission.reason}`);
            }

            const graphClient = this.createGraphClient(accessToken);
            
            logger.debug('🔍 Reading Excel range', { driveId, itemId, worksheetId, range });

            const response = await graphClient
                .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`)
                .get();

            const rangeData = {
                address: response.address,
                values: response.values,
                formulas: response.formulas,
                text: response.text,
                rowCount: response.rowCount,
                columnCount: response.columnCount
            };

            // Log successful permission check and operation
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                requestedPermission: 'READ',
                granted: true
            });

            auditService.logReadOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                cellCount: rangeData.rowCount * rangeData.columnCount,
                success: true
            });

            logger.debug('✅ Range read successfully', { 
                range,
                rowCount: rangeData.rowCount,
                columnCount: rangeData.columnCount
            });

            return rangeData;

        } catch (error) {
            logger.error('❌ Excel service - failed to read range:', error);
            
            if (error.code === 'InvalidArgument') {
                throw new Error(`Invalid range format: ${range}`);
            } else if (error.code === 'itemNotFound') {
                throw new Error('Worksheet or range not found');
            }
            
            throw error;
        }
    }

    /**
     * Write data to Excel range with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Updated range data
     */
    async writeRange(params) {
        const { accessToken, driveId, itemId, worksheetId, range, values, auditContext } = params;

        try {
            // Validate input data
            if (!Array.isArray(values) || values.length === 0) {
                throw new Error('Values must be a non-empty array');
            }

            // Check permissions
            const hasPermission = permissions.canWriteRange(
                auditContext.user, 
                itemId, 
                worksheetId, 
                range
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    range: range,
                    requestedPermission: 'WRITE',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Write access denied: ${hasPermission.reason}`);
            }

            const graphClient = this.createGraphClient(accessToken);
            
            logger.debug('🔍 Writing to Excel range', { driveId, itemId, worksheetId, range });

            // Read current values for audit trail
            let oldValues = null;
            try {
                const currentResponse = await graphClient
                    .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`)
                    .get();
                oldValues = currentResponse.values;
            } catch (readError) {
                logger.warn('Could not read current values for audit trail:', readError.message);
            }

            // Write new values
            const response = await graphClient
                .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`)
                .patch({ values: values });

            const updatedData = {
                address: response.address,
                values: response.values,
                rowCount: response.rowCount,
                columnCount: response.columnCount
            };

            // Log successful permission check and operation
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                requestedPermission: 'WRITE',
                granted: true
            });

            auditService.logWriteOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                oldValues: oldValues,
                newValues: values,
                cellsModified: updatedData.rowCount * updatedData.columnCount,
                success: true
            });

            logger.debug('✅ Range written successfully', { 
                range,
                rowCount: updatedData.rowCount,
                columnCount: updatedData.columnCount
            });

            return updatedData;

        } catch (error) {
            logger.error('❌ Excel service - failed to write range:', error);
            
            if (error.code === 'InvalidArgument') {
                throw new Error(`Invalid range format or data: ${range}`);
            } else if (error.code === 'itemNotFound') {
                throw new Error('Worksheet or range not found');
            }
            
            throw error;
        }
    }

    /**
     * Read data from Excel table with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Table data
     */
    async readTable(params) {
        const { accessToken, driveId, itemId, worksheetId, tableName, auditContext } = params;

        try {
            // Check permissions
            const hasPermission = permissions.canReadTable(
                auditContext.user, 
                itemId, 
                worksheetId, 
                tableName
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    table: tableName,
                    requestedPermission: 'READ_TABLE',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Table read access denied: ${hasPermission.reason}`);
            }

            const graphClient = this.createGraphClient(accessToken);
            
            logger.debug('🔍 Reading Excel table', { driveId, itemId, worksheetId, tableName });

            // Get table info
            const tableResponse = await graphClient
                .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/tables/${tableName}`)
                .get();
            
            // Get table data
            const dataResponse = await graphClient
                .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/tables/${tableName}/range`)
                .get();

            const tableData = {
                id: tableResponse.id,
                name: tableResponse.name,
                address: dataResponse.address,
                values: dataResponse.values,
                headers: dataResponse.values[0], // First row is typically headers
                rows: dataResponse.values.slice(1), // Data rows
                rowCount: dataResponse.rowCount,
                columnCount: dataResponse.columnCount
            };

            // Log successful permission check and operation
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                requestedPermission: 'READ_TABLE',
                granted: true
            });

            auditService.logReadOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                cellCount: tableData.rowCount * tableData.columnCount,
                success: true
            });

            logger.debug('✅ Table read successfully', { 
                tableName,
                rowCount: tableData.rowCount,
                columnCount: tableData.columnCount
            });

            return tableData;

        } catch (error) {
            logger.error('❌ Excel service - failed to read table:', error);
            
            if (error.code === 'itemNotFound') {
                throw new Error(`Table '${tableName}' not found in worksheet`);
            }
            
            throw error;
        }
    }

    /**
     * Add rows to Excel table with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Result
     */
    async addTableRows(params) {
        const { accessToken, driveId, itemId, worksheetId, tableName, rows, auditContext } = params;

        try {
            // Validate input data
            if (!Array.isArray(rows) || rows.length === 0) {
                throw new Error('Rows must be a non-empty array');
            }

            // Check permissions
            const hasPermission = permissions.canWriteTable(
                auditContext.user, 
                itemId, 
                worksheetId, 
                tableName
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    table: tableName,
                    requestedPermission: 'WRITE_TABLE',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Table write access denied: ${hasPermission.reason}`);
            }

            const graphClient = this.createGraphClient(accessToken);
            
            logger.debug('🔍 Adding rows to Excel table', { driveId, itemId, worksheetId, tableName, rowCount: rows.length });

            const response = await graphClient
                .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/tables/${tableName}/rows`)
                .post({ values: rows });

            // Log successful permission check and operation
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                requestedPermission: 'WRITE_TABLE',
                granted: true
            });

            auditService.logWriteOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                newValues: rows,
                cellsModified: rows.length * (rows[0]?.length || 0),
                success: true
            });

            logger.debug('✅ Rows added to table successfully', { 
                tableName,
                rowsAdded: rows.length
            });

            return response;

        } catch (error) {
            logger.error('❌ Excel service - failed to add table rows:', error);
            
            if (error.code === 'itemNotFound') {
                throw new Error(`Table '${tableName}' not found in worksheet`);
            } else if (error.code === 'InvalidArgument') {
                throw new Error('Invalid row data format');
            }
            
            throw error;
        }
    }

    /**
     * Validate range format
     * @param {string} range - Range string (e.g., 'A1:C10')
     * @returns {boolean} True if valid
     */
    validateRangeFormat(range) {
        // Basic range validation regex
        const rangeRegex = /^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/;
        return rangeRegex.test(range);
    }

    /**
     * Parse range to get dimensions
     * @param {string} range - Range string
     * @returns {Object} Range dimensions
     */
    parseRange(range) {
        try {
            // This is a simplified parser - production would need more robust parsing
            const parts = range.split(':');
            if (parts.length === 1) {
                // Single cell
                return { startCell: parts[0], endCell: parts[0], isSingleCell: true };
            } else if (parts.length === 2) {
                return { startCell: parts[0], endCell: parts[1], isSingleCell: false };
            }
            throw new Error('Invalid range format');
        } catch (error) {
            logger.error('Failed to parse range:', error);
            throw new Error(`Invalid range format: ${range}`);
        }
    }
}

module.exports = new ExcelService();
