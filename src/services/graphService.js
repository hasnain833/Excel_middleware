/**
 * Microsoft Graph Service
 * Handles all interactions with Microsoft Graph API
 */

const axios = require('axios');
const logger = require('../config/logger');
const auditService = require('./auditService');

class GraphService {
    constructor() {
        this.baseURL = process.env.GRAPH_API_BASE_URL || 'https://graph.microsoft.com/v1.0';
        this.timeout = 15000; // Reduced to 15 seconds
        this.retryAttempts = 2;
    }

    /**
     * Create axios instance with authentication
     * @param {string} accessToken - Access token
     * @returns {Object} Axios instance
     */
    createAuthenticatedClient(accessToken) {
        return axios.create({
            baseURL: this.baseURL,
            timeout: this.timeout,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
    }

    /**
     * Get all accessible workbooks
     * @param {string} accessToken - Access token
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Array>} List of workbooks
     */
    async getWorkbooks(accessToken, auditContext) {
        try {
            const client = this.createAuthenticatedClient(accessToken);
            
            // Optimized: Get workbooks from OneDrive only initially
            const response = await client.get('/me/drive/root/search(q=\'.xlsx\')?$filter=file ne null&$top=50');
            
            const workbooks = response.data.value.map(item => ({
                id: item.id,
                name: item.name,
                webUrl: item.webUrl,
                parentReference: item.parentReference,
                lastModifiedDateTime: item.lastModifiedDateTime,
                size: item.size
            }));

            auditService.logSystemEvent({
                event: 'WORKBOOKS_RETRIEVED',
                details: { count: workbooks.length, user: auditContext.user }
            });

            return workbooks;

        } catch (error) {
            logger.error('Failed to get workbooks:', {
                message: error.message,
                name: error.name,
                stack: error.stack
            });
            throw new Error(`Failed to retrieve workbooks: ${error.message}`);
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
            const client = this.createAuthenticatedClient(accessToken);
            const response = await client.get(`/drives/${driveId}/items/${itemId}/workbook/worksheets`);
            
            const worksheets = response.data.value.map(sheet => ({
                id: sheet.id,
                name: sheet.name,
                position: sheet.position,
                visibility: sheet.visibility
            }));

            logger.info(`Retrieved ${worksheets.length} worksheets from workbook ${itemId}`);
            return worksheets;

        } catch (error) {
            logger.error('Failed to get worksheets:', error);
            throw new Error(`Failed to retrieve worksheets: ${error.message}`);
        }
    }

    /**
     * Read data from Excel range
     * @param {string} accessToken - Access token
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @param {string} worksheetId - Worksheet ID
     * @param {string} range - Range (e.g., 'A1:C10')
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Object>} Range data
     */
    async readRange(accessToken, driveId, itemId, worksheetId, range, auditContext) {
        try {
            const client = this.createAuthenticatedClient(accessToken);
            const response = await client.get(
                `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`
            );

            const rangeData = {
                address: response.data.address,
                values: response.data.values,
                formulas: response.data.formulas,
                text: response.data.text,
                rowCount: response.data.rowCount,
                columnCount: response.data.columnCount
            };

            // Log audit entry
            auditService.logReadOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                cellCount: rangeData.rowCount * rangeData.columnCount,
                success: true
            });

            logger.info(`Successfully read range ${range} from worksheet ${worksheetId}`);
            return rangeData;

        } catch (error) {
            // Log failed audit entry
            auditService.logReadOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                success: false,
                error: error.message
            });

            logger.error('Failed to read range:', error);
            throw new Error(`Failed to read range ${range}: ${error.message}`);
        }
    }

    /**
     * Write data to Excel range
     * @param {string} accessToken - Access token
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @param {string} worksheetId - Worksheet ID
     * @param {string} range - Range (e.g., 'A1:C10')
     * @param {Array} values - 2D array of values
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Object>} Updated range data
     */
    async writeRange(accessToken, driveId, itemId, worksheetId, range, values, auditContext) {
        let oldValues = null;
        
        try {
            const client = this.createAuthenticatedClient(accessToken);
            
            // First, read the current values for audit trail
            try {
                const currentData = await this.readRange(accessToken, driveId, itemId, worksheetId, range, auditContext);
                oldValues = currentData.values;
            } catch (readError) {
                logger.warn('Could not read current values for audit trail:', readError.message);
            }

            // Write new values
            const response = await client.patch(
                `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`,
                { values: values }
            );

            const updatedData = {
                address: response.data.address,
                values: response.data.values,
                rowCount: response.data.rowCount,
                columnCount: response.data.columnCount
            };

            // Log audit entry
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

            logger.info(`Successfully wrote to range ${range} in worksheet ${worksheetId}`);
            return updatedData;

        } catch (error) {
            // Log failed audit entry
            auditService.logWriteOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                oldValues: oldValues,
                newValues: values,
                success: false,
                error: error.message
            });

            logger.error('Failed to write range:', error);
            throw new Error(`Failed to write to range ${range}: ${error.message}`);
        }
    }

    /**
     * Read data from Excel table
     * @param {string} accessToken - Access token
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @param {string} worksheetId - Worksheet ID
     * @param {string} tableName - Table name
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Object>} Table data
     */
    async readTable(accessToken, driveId, itemId, worksheetId, tableName, auditContext) {
        try {
            const client = this.createAuthenticatedClient(accessToken);
            
            // Get table info
            const tableResponse = await client.get(
                `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/tables/${tableName}`
            );
            
            // Get table data
            const dataResponse = await client.get(
                `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/tables/${tableName}/range`
            );

            const tableData = {
                id: tableResponse.data.id,
                name: tableResponse.data.name,
                address: dataResponse.data.address,
                values: dataResponse.data.values,
                headers: dataResponse.data.values[0], // First row is typically headers
                rows: dataResponse.data.values.slice(1), // Data rows
                rowCount: dataResponse.data.rowCount,
                columnCount: dataResponse.data.columnCount
            };

            // Log audit entry
            auditService.logReadOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                cellCount: tableData.rowCount * tableData.columnCount,
                success: true
            });

            logger.info(`Successfully read table ${tableName} from worksheet ${worksheetId}`);
            return tableData;

        } catch (error) {
            // Log failed audit entry
            auditService.logReadOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                success: false,
                error: error.message
            });

            logger.error('Failed to read table:', error);
            throw new Error(`Failed to read table ${tableName}: ${error.message}`);
        }
    }

    /**
     * Add rows to Excel table
     * @param {string} accessToken - Access token
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @param {string} worksheetId - Worksheet ID
     * @param {string} tableName - Table name
     * @param {Array} rows - Array of row data
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Object>} Result
     */
    async addTableRows(accessToken, driveId, itemId, worksheetId, tableName, rows, auditContext) {
        try {
            const client = this.createAuthenticatedClient(accessToken);
            
            const response = await client.post(
                `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/tables/${tableName}/rows`,
                { values: rows }
            );

            // Log audit entry
            auditService.logWriteOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                newValues: rows,
                cellsModified: rows.length * (rows[0]?.length || 0),
                success: true
            });

            logger.info(`Successfully added ${rows.length} rows to table ${tableName}`);
            return response.data;

        } catch (error) {
            // Log failed audit entry
            auditService.logWriteOperation({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                newValues: rows,
                success: false,
                error: error.message
            });

            logger.error('Failed to add table rows:', error);
            throw new Error(`Failed to add rows to table ${tableName}: ${error.message}`);
        }
    }
}

module.exports = new GraphService();
