/**
 * Health Check Routes
 * Provides health monitoring endpoints
 */

const express = require('express');
const router = express.Router();

// Controllers
const healthController = require('../controllers/healthController');

/**
 * @route GET /health
 * @desc Basic health check
 * @access Public
 */
router.get('/', healthController.basicHealth);

/**
 * @route GET /health/detailed
 * @desc Detailed health check with dependency status
 * @access Public
 */
router.get('/detailed', healthController.detailedHealth);

/**
 * @route GET /health/ready
 * @desc Readiness probe for container orchestration
 * @access Public
 */
router.get('/ready', healthController.readiness);

/**
 * @route GET /health/live
 * @desc Liveness probe for container orchestration
 * @access Public
 */
router.get('/live', healthController.liveness);

module.exports = router;
