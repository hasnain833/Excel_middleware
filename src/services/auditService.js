const logger = require('../config/logger');
const { v4: uuidv4 } = require('uuid');

class AuditService {
    constructor() {
        this.auditLogger = logger.child({ component: 'audit' });
    }
    logReadOperation(params) {
        const auditEntry = {
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            operation: 'READ',
            user: params.user || 'system',
            workbookId: params.workbookId,
            worksheetId: params.worksheetId,
            range: params.range,
            success: params.success,
            requestId: params.requestId,
            ipAddress: params.ipAddress
        };

        this.auditLogger.info('Excel read operation', auditEntry);
        return auditEntry.id;
    }

    logWriteOperation(params) {
        const auditEntry = {
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            operation: 'WRITE',
            user: params.user || 'system',
            workbookId: params.workbookId,
            workbookName: params.workbookName,
            worksheetId: params.worksheetId,
            worksheetName: params.worksheetName,
            range: params.range,
            table: params.table,
            oldValues: params.oldValues,
            newValues: params.newValues,
            cellsModified: params.cellsModified,
            success: params.success,
            error: params.error,
            requestId: params.requestId,
            ipAddress: params.ipAddress,
            userAgent: params.userAgent
        };

        this.auditLogger.info('Excel write operation', auditEntry);
        return auditEntry.id;
    }

    logPermissionCheck(params) {
        const auditEntry = {
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            operation: 'PERMISSION_CHECK',
            user: params.user || 'system',
            workbookId: params.workbookId,
            worksheetId: params.worksheetId,
            range: params.range,
            requestedPermission: params.requestedPermission,
            granted: params.granted,
            reason: params.reason,
            requestId: params.requestId,
            ipAddress: params.ipAddress
        };

        this.auditLogger.info('Permission check', auditEntry);
        return auditEntry.id;
    }

    logAuthEvent(params) {
        const auditEntry = {
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            operation: 'AUTHENTICATION',
            event: params.event, // 'TOKEN_ACQUIRED', 'TOKEN_REFRESH', 'AUTH_FAILED'
            success: params.success,
            error: params.error,
            tokenExpiry: params.tokenExpiry,
            requestId: params.requestId,
            ipAddress: params.ipAddress
        };

        this.auditLogger.info('Authentication event', auditEntry);
        return auditEntry.id;
    }

    logSystemEvent(params) {
        const auditEntry = {
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            operation: 'SYSTEM',
            event: params.event, // 'SERVER_START', 'SERVER_STOP', 'ERROR'
            details: params.details,
            severity: params.severity || 'info'
        };

        this.auditLogger.info('System event', auditEntry);
        return auditEntry.id;
    }

    generateAuditReport(startDate, endDate, filters = {}) {
        // This would typically query a database or log files
        // For now, we'll log the report request
        const reportRequest = {
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            operation: 'AUDIT_REPORT',
            startDate: startDate.toISOString(),
            endDate: endDate.toISOString(),
            filters: filters
        };

        this.auditLogger.info('Audit report requested', reportRequest);
        return reportRequest.id;
    }
    createAuditContext(req) {
        return {
            requestId: req.id || uuidv4(),
            ipAddress: req.ip || req.connection.remoteAddress,
            userAgent: req.get('User-Agent'),
            user: req.user?.id || req.headers['x-user-id'] || 'anonymous',
            timestamp: new Date().toISOString()
        };
    }
}

module.exports = new AuditService();
