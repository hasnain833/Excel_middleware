const auditLogger = require('../middleware/auditLogger');
const { catchAsync } = require('../middleware/errorHandler');

class AuditController {
    getAuditLogs = catchAsync(async (req, res) => {
        const {
            user,
            operation,
            fileName,
            startDate,
            endDate,
            success,
            limit = 50
        } = req.query;

        const filters = {
            user,
            operation,
            fileName,
            startDate,
            endDate,
            success: success !== undefined ? success === 'true' : undefined,
            limit: Math.min(parseInt(limit) || 50, 100) // Max 100 entries
        };

        const entries = await auditLogger.getAuditEntries(filters);

        res.json({
            status: 'success',
            data: {
                entries,
                count: entries.length,
                filters: filters
            }
        });
    });
}

module.exports = new AuditController();
