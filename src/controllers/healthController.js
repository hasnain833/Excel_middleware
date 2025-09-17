/**
 * Health Check Controller
 * Provides health status endpoints for monitoring
 */

const azureAuth = require('../auth/azureAuth');
const logger = require('../config/logger');
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

    /**
     * Detailed health check including dependencies
     */
    detailedHealth = catchAsync(async (req, res) => {
        const healthChecks = {
            service: 'healthy',
            azure_auth: 'unknown',
            graph_api: 'unknown',
            memory: 'healthy',
            disk: 'healthy'
        };

        // Check Azure authentication
        try {
            const tokenInfo = azureAuth.getTokenInfo();
            healthChecks.azure_auth = tokenInfo.isValid ? 'healthy' : 'degraded';
        } catch (error) {
            healthChecks.azure_auth = 'unhealthy';
        }

        // Check memory usage
        const memoryUsage = process.memoryUsage();
        const memoryUsageMB = {
            rss: Math.round(memoryUsage.rss / 1024 / 1024),
            heapTotal: Math.round(memoryUsage.heapTotal / 1024 / 1024),
            heapUsed: Math.round(memoryUsage.heapUsed / 1024 / 1024),
            external: Math.round(memoryUsage.external / 1024 / 1024)
        };

        // Memory health check (alert if heap used > 500MB)
        if (memoryUsageMB.heapUsed > 500) {
            healthChecks.memory = 'degraded';
        }

        // Overall health status
        const unhealthyServices = Object.values(healthChecks).filter(status => status === 'unhealthy');
        const degradedServices = Object.values(healthChecks).filter(status => status === 'degraded');
        
        let overallStatus = 'healthy';
        if (unhealthyServices.length > 0) {
            overallStatus = 'unhealthy';
        } else if (degradedServices.length > 0) {
            overallStatus = 'degraded';
        }

        const statusCode = overallStatus === 'healthy' ? 200 : 
                          overallStatus === 'degraded' ? 200 : 503;

        res.status(statusCode).json({
            status: overallStatus === 'healthy' ? 'success' : 'error',
            data: {
                health: overallStatus,
                service: 'excel-gpt-middleware',
                timestamp: new Date().toISOString(),
                uptime: process.uptime(),
                version: process.env.npm_package_version || '1.0.0',
                checks: healthChecks,
                system: {
                    memory: memoryUsageMB,
                    nodeVersion: process.version,
                    platform: process.platform,
                    arch: process.arch
                }
            }
        });
    });

    /**
     * Readiness probe for Kubernetes/container orchestration
     */
    readiness = catchAsync(async (req, res) => {
        try {
            // Check if Azure auth is working
            const tokenInfo = azureAuth.getTokenInfo();
            
            if (!tokenInfo.hasToken) {
                // Try to get a token
                await azureAuth.getAccessToken();
            }

            res.json({
                status: 'success',
                data: {
                    readiness: 'ready',
                    timestamp: new Date().toISOString()
                }
            });
        } catch (error) {
            logger.error('Readiness check failed:', error);
            res.status(503).json({
                status: 'error',
                error: {
                    code: 503,
                    message: 'Authentication service unavailable'
                },
                timestamp: new Date().toISOString()
            });
        }
    });

    /**
     * Liveness probe for Kubernetes/container orchestration
     */
    liveness = catchAsync(async (req, res) => {
        // Simple liveness check - if we can respond, we're alive
        res.json({
            status: 'success',
            data: {
                liveness: 'alive',
                timestamp: new Date().toISOString()
            }
        });
    });
}

module.exports = new HealthController();
