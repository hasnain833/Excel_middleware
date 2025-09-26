const fs = require("fs").promises;
const path = require("path");
const logger = require("../config/logger");

class RangeValidator {
  constructor() {
    this.configPath = path.join(__dirname, "../../rangePermissions.json");
    this.allowedRanges = [];
    this.lockedRanges = [];
    this.loadConfig();
  }

  async loadConfig() {
    try {
      const configData = await fs.readFile(this.configPath, "utf8");
      const config = JSON.parse(configData);

      this.allowedRanges = config.allowedRanges || [];
      this.lockedRanges = config.lockedRanges || [];

      logger.info("Range permissions loaded", {
        allowedCount: this.allowedRanges.length,
        lockedCount: this.lockedRanges.length,
      });
    } catch (error) {
      logger.error("Failed to load range permissions config:", error);
      // Use empty arrays as fallback - deny all writes
      this.allowedRanges = [];
      this.lockedRanges = [];
    }
  }

  parseRange(range) {
    try {
      let sheetName = "";
      let rangeAddress = range;

      // Check if range includes sheet name
      if (range.includes("!")) {
        const parts = range.split("!");
        sheetName = parts[0];
        rangeAddress = parts[1];
      }

      // Parse range address (e.g., "A1:B10")
      const rangeParts = rangeAddress.split(":");
      if (rangeParts.length !== 2) {
        throw new Error("Invalid range format");
      }

      const startCell = this.parseCellAddress(rangeParts[0]);
      const endCell = this.parseCellAddress(rangeParts[1]);

      return {
        sheetName,
        startCell,
        endCell,
        fullRange: range,
      };
    } catch (error) {
      logger.error("Failed to parse range:", { range, error: error.message });
      throw new Error(`Invalid range format: ${range}`);
    }
  }

  parseCellAddress(cellAddress) {
    const match = cellAddress.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`Invalid cell address: ${cellAddress}`);
    }

    const colLetters = match[1];
    const rowNumber = parseInt(match[2]);

    // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
    let colNumber = 0;
    for (let i = 0; i < colLetters.length; i++) {
      colNumber = colNumber * 26 + (colLetters.charCodeAt(i) - 64);
    }

    return {
      col: colNumber,
      row: rowNumber,
      address: cellAddress,
    };
  }

  rangesOverlap(range1, range2) {
    // Must be on the same sheet to overlap
    if (
      range1.sheetName &&
      range2.sheetName &&
      range1.sheetName !== range2.sheetName
    ) {
      return false;
    }
    const r1StartCol = range1.startCell.col;
    const r1EndCol = range1.endCell.col;
    const r1StartRow = range1.startCell.row;
    const r1EndRow = range1.endCell.row;
    const r2StartCol = range2.startCell.col;
    const r2EndCol = range2.endCell.col;
    const r2StartRow = range2.startCell.row;
    const r2EndRow = range2.endCell.row;

    // Check for overlap
    return !(
      r1EndCol < r2StartCol ||
      r2EndCol < r1StartCol ||
      r1EndRow < r2StartRow ||
      r2EndRow < r1StartRow
    );
  }

  validateRange(requestedRange, worksheetId = "") {
    try {
      // Reload config to get latest permissions
      this.loadConfig();

      // Construct full range with worksheet if not included
      let fullRange = requestedRange;
      if (!requestedRange.includes("!") && worksheetId) {
        fullRange = `${worksheetId}!${requestedRange}`;
      }

      const parsedRequested = this.parseRange(fullRange);

      // Check if range is explicitly locked
      for (const lockedRange of this.lockedRanges) {
        const parsedLocked = this.parseRange(lockedRange);
        if (this.rangesOverlap(parsedRequested, parsedLocked)) {
          return {
            allowed: false,
            reason: `Range overlaps with locked range: ${lockedRange}`,
            code: "RANGE_LOCKED",
          };
        }
      }

      // Check if range is within allowed ranges
      for (const allowedRange of this.allowedRanges) {
        const parsedAllowed = this.parseRange(allowedRange);
        if (this.rangesOverlap(parsedRequested, parsedAllowed)) {
          // Additional check: ensure requested range is fully contained within allowed range
          if (this.isRangeContained(parsedRequested, parsedAllowed)) {
            return {
              allowed: true,
              reason: `Range is within allowed range: ${allowedRange}`,
              code: "RANGE_ALLOWED",
            };
          }
        }
      }

      return {
        allowed: false,
        reason: "Range is not within any allowed ranges",
        code: "RANGE_NOT_ALLOWED",
      };
    } catch (error) {
      logger.error("Range validation error:", error);
      return {
        allowed: false,
        reason: `Validation error: ${error.message}`,
        code: "VALIDATION_ERROR",
      };
    }
  }

  isRangeContained(range1, range2) {
    // Must be on the same sheet
    if (
      range1.sheetName &&
      range2.sheetName &&
      range1.sheetName !== range2.sheetName
    ) {
      return false;
    }

    return (
      range1.startCell.col >= range2.startCell.col &&
      range1.endCell.col <= range2.endCell.col &&
      range1.startCell.row >= range2.startCell.row &&
      range1.endCell.row <= range2.endCell.row
    );
  }

  middleware() {
    return async (req, res, next) => {
      // Only validate write operations
      if (req.method !== "POST" && req.method !== "PATCH") {
        return next();
      }

      // Only validate Excel write endpoints
      if (
        !req.path.includes("/excel/write") &&
        !req.path.includes("/excel/add-table-rows")
      ) {
        return next();
      }

      try {
        const { range, worksheetId } = req.body;

        if (!range) {
          return res.status(400).json({
            status: "error",
            error: {
              code: 400,
              message: "Range is required for write operations",
            },
          });
        }

        const validation = this.validateRange(range, worksheetId);

        if (!validation.allowed) {
          logger.warn("Range access denied", {
            user: req.headers["x-user-id"] || "anonymous",
            range,
            worksheetId,
            reason: validation.reason,
            ip: req.ip,
          });

          return res.status(403).json({
            status: "error",
            error: {
              code: 403,
              message: "Range access denied",
              details: validation.reason,
              allowedRanges: this.allowedRanges,
            },
          });
        }

        // Range is allowed, continue to next middleware
        logger.debug("Range validation passed", {
          user: req.headers["x-user-id"] || "anonymous",
          range,
          worksheetId,
          reason: validation.reason,
        });

        next();
      } catch (error) {
        logger.error("Range validation middleware error:", error);
        return res.status(500).json({
          status: "error",
          error: {
            code: 500,
            message: "Range validation failed",
            details: error.message,
          },
        });
      }
    };
  }
}

module.exports = new RangeValidator();
