/**
 * Main Server File
 * Express server setup with all middleware and routes
 */

require('dotenv').config();
require('express-async-errors');

const express = require('express');
const helmet = require('helmet');
const cors = require('cors');

// Import middleware
const { globalErrorHandler, handleNotFound, handleUnhandledRejections } = require('./middleware/errorHandler');
const { generalLimiter } = require('./middleware/rateLimiter');

// Import routes
const excelRoutes = require('./routes/excel');
const healthRoutes = require('./routes/health');

// Import services
const logger = require('./config/logger');
const auditService = require('./services/auditService');

// Handle unhandled rejections and exceptions
handleUnhandledRejections();

class Server {
    constructor() {
        this.app = express();
        this.port = process.env.PORT || 3000;
        this.setupMiddleware();
        this.setupRoutes();
        this.setupErrorHandling();
    }

    /**
     * Setup Express middleware
     */
    setupMiddleware() {
        // Trust proxy if behind reverse proxy
        if (process.env.TRUST_PROXY === 'true') {
            this.app.set('trust proxy', 1);
        }

        // Security middleware
        this.app.use(helmet({
            contentSecurityPolicy: false, // Simplified for API
            crossOriginEmbedderPolicy: false,
            hsts: {
                maxAge: 31536000,
                includeSubDomains: true,
                preload: true
            }
        }));

        // CORS configuration
        const corsOptions = {
            origin: process.env.ALLOWED_ORIGINS?.split(',').map(origin => origin.trim()) || ['http://localhost:3000'],
            credentials: true,
            optionsSuccessStatus: 200
        };
        this.app.use(cors(corsOptions));

        // Body parsing middleware
        this.app.use(express.json({ 
            limit: process.env.MAX_REQUEST_SIZE || '10mb',
            strict: true
        }));
        this.app.use(express.urlencoded({ 
            extended: true, 
            limit: process.env.MAX_REQUEST_SIZE || '10mb'
        }));

        // Request ID middleware
        this.app.use((req, res, next) => {
            req.id = require('uuid').v4();
            res.setHeader('X-Request-ID', req.id);
            next();
        });

        // Request logging middleware
        this.app.use((req, res, next) => {
            const start = Date.now();
            
            res.on('finish', () => {
                const duration = Date.now() - start;
                logger.info('HTTP Request', {
                    requestId: req.id,
                    method: req.method,
                    url: req.originalUrl,
                    statusCode: res.statusCode,
                    duration: `${duration}ms`,
                    ip: req.ip,
                    userAgent: req.get('User-Agent')
                });
            });
            
            next();
        });

        // Apply general rate limiting
        this.app.use(generalLimiter);
    }

    /**
     * Setup application routes
     */
    setupRoutes() {
        // Health check routes (no authentication required)
        this.app.use('/health', healthRoutes);

        // API routes
        this.app.use('/api/excel', excelRoutes);

        // Root endpoint
        this.app.get('/', (req, res) => {
            res.json({
                service: 'Excel GPT Middleware',
                version: process.env.npm_package_version || '1.0.0',
                status: 'running',
                timestamp: new Date().toISOString(),
                endpoints: {
                    health: '/health',
                    api: '/api/excel',
                    documentation: '/api/docs'
                }
            });
        });

        // API documentation endpoint
        this.app.get('/api/docs', (req, res) => {
            res.json({
                service: 'Excel GPT Middleware API',
                version: '1.0.0',
                endpoints: {
                    workbooks: {
                        method: 'GET',
                        path: '/api/excel/workbooks',
                        description: 'Get all accessible workbooks'
                    },
                    worksheets: {
                        method: 'GET',
                        path: '/api/excel/worksheets',
                        description: 'Get worksheets in a workbook',
                        parameters: ['driveId', 'itemId']
                    },
                    readRange: {
                        method: 'POST',
                        path: '/api/excel/read',
                        description: 'Read data from Excel range',
                        body: ['driveId', 'itemId', 'worksheetId', 'range']
                    },
                    writeRange: {
                        method: 'POST',
                        path: '/api/excel/write',
                        description: 'Write data to Excel range',
                        body: ['driveId', 'itemId', 'worksheetId', 'range', 'values']
                    },
                    readTable: {
                        method: 'POST',
                        path: '/api/excel/read-table',
                        description: 'Read data from Excel table',
                        body: ['driveId', 'itemId', 'worksheetId', 'tableName']
                    },
                    addTableRows: {
                        method: 'POST',
                        path: '/api/excel/add-table-rows',
                        description: 'Add rows to Excel table',
                        body: ['driveId', 'itemId', 'worksheetId', 'tableName', 'rows']
                    },
                    batch: {
                        method: 'POST',
                        path: '/api/excel/batch',
                        description: 'Perform batch Excel operations',
                        body: ['operations']
                    }
                },
                authentication: {
                    type: 'Azure AD Client Credentials',
                    description: 'Automatic authentication using Azure AD service principal'
                }
            });
        });

        // Handle 404 for undefined routes
        this.app.use(handleNotFound);
    }

    /**
     * Setup error handling
     */
    setupErrorHandling() {
        this.app.use(globalErrorHandler);
    }

    /**
     * Start the server
     */
    async start() {
        try {
            // Log system startup
            auditService.logSystemEvent({
                event: 'SERVER_START',
                details: {
                    port: this.port,
                    nodeVersion: process.version,
                    environment: process.env.NODE_ENV || 'development'
                }
            });

            this.server = this.app.listen(this.port, () => {
                logger.info(`🚀 Excel GPT Middleware server started`, {
                    port: this.port,
                    environment: process.env.NODE_ENV || 'development',
                    nodeVersion: process.version,
                    timestamp: new Date().toISOString()
                });

                logger.info('📋 Available endpoints:', {
                    health: `http://localhost:${this.port}/health`,
                    api: `http://localhost:${this.port}/api/excel`,
                    docs: `http://localhost:${this.port}/api/docs`
                });
            });

            // Graceful shutdown handling
            this.setupGracefulShutdown();

        } catch (error) {
            logger.error('Failed to start server:', error);
            auditService.logSystemEvent({
                event: 'SERVER_START_FAILED',
                details: { error: error.message },
                severity: 'error'
            });
            process.exit(1);
        }
    }

    /**
     * Setup graceful shutdown
     */
    setupGracefulShutdown() {
        const gracefulShutdown = (signal) => {
            logger.info(`Received ${signal}. Starting graceful shutdown...`);
            
            auditService.logSystemEvent({
                event: 'SERVER_SHUTDOWN',
                details: { signal }
            });

            this.server.close((err) => {
                if (err) {
                    logger.error('Error during server shutdown:', err);
                    process.exit(1);
                }

                logger.info('Server closed successfully');
                process.exit(0);
            });

            // Force close after 10 seconds
            setTimeout(() => {
                logger.error('Could not close connections in time, forcefully shutting down');
                process.exit(1);
            }, 10000);
        };

        // Listen for termination signals
        process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
        process.on('SIGINT', () => gracefulShutdown('SIGINT'));
    }

    /**
     * Stop the server
     */
    async stop() {
        return new Promise((resolve, reject) => {
            if (this.server) {
                this.server.close((err) => {
                    if (err) {
                        reject(err);
                    } else {
                        resolve();
                    }
                });
            } else {
                resolve();
            }
        });
    }
}

// Create and start server if this file is run directly
if (require.main === module) {
    const server = new Server();
    server.start();
}

module.exports = Server;
