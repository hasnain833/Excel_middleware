const express = require("express");
const router = express.Router();

// Controllers
const healthController = require("../controllers/healthController");

// Routers
router.get("/", healthController.basicHealth);

module.exports = router;
