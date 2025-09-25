const express = require("express");
const cors = require("cors");

const app = express();

// Routers
const healthRoutes = require("../src/routes/health.js");
const excelRoutes = require("../src/routes/excel.js");
const renameRoutes = require("../src/routes/rename.js");

// Basic middleware
app.use(cors());
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

// Root (lightweight heartbeat; optional to keep)
app.get("/", (req, res) => {
  res.json({
    service: "Excel GPT Middleware",
    version: process.env.npm_package_version || "1.0.0",
    status: "running",
    environment: "serverless",
    timestamp: new Date().toISOString(),
  });
});

// Mount routers
app.use("/health", healthRoutes);
app.use("/api/excel", excelRoutes);
app.use("/api/excel", renameRoutes);

// 404 (keep last)
app.use("*", (req, res) => {
  res.status(404).json({
    status: "error",
    error: {
      code: 404,
      message: `Endpoint not found: ${req.method} ${req.originalUrl}`,
    },
    timestamp: new Date().toISOString(),
  });
});

module.exports = app;
