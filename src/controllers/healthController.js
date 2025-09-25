/**
 * Health Check Controller
 * Provides health status endpoint for monitoring
 */

const { catchAsync } = require('../middleware/errorHandler');

class HealthController {
    /**
     * Basic health check
     */
    basicHealth = catchAsync(async (req, res) => {
        res.json({
            status: 'success',
            data: {
                health: 'healthy',
                service: 'excel-gpt-middleware',
                timestamp: new Date().toISOString(),
                uptime: process.uptime(),
                version: process.env.npm_package_version || '1.0.0'
            }
        });
    });
}

module.exports = new HealthController();
