const { Client } = require("@microsoft/microsoft-graph-client");
const logger = require("../config/logger");
const resolverService = require("./resolverService");
const auditService = require("./auditService");
const { AppError } = require("../middleware/errorHandler");

class FindReplaceService {
  constructor() {
    // Cache for search results to avoid duplicate API calls
    this.searchCache = new Map();
    this.ttlMs = 2 * 60 * 1000; // 2 minutes TTL for search results
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }

  async findOccurrences(
    accessToken,
    driveId,
    itemId,
    searchTerm,
    scope = "entire_sheet",
    rangeSpec = null
  ) {
    if (!driveId || !itemId || !searchTerm) {
      throw new AppError("driveId, itemId, and searchTerm are required", 400);
    }

    const cacheKey = `${driveId}:${itemId}:${searchTerm}:${scope}:${rangeSpec}`;
    const cached = this.searchCache.get(cacheKey);

    if (cached && Date.now() - cached.ts < this.ttlMs) {
      return cached.results;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const matches = [];

      switch (scope) {
        case "header_only":
          const headerMatches = await this.findInHeaders(
            graphClient,
            driveId,
            itemId,
            searchTerm
          );
          matches.push(...headerMatches);
          break;

        case "specific_range":
          if (!rangeSpec) {
            throw new AppError(
              "rangeSpec is required for specific_range scope",
              400
            );
          }
          const rangeMatches = await this.findInRange(
            graphClient,
            driveId,
            itemId,
            searchTerm,
            rangeSpec
          );
          matches.push(...rangeMatches);
          break;

        case "entire_sheet":
          const sheetMatches = await this.findInSheet(
            graphClient,
            driveId,
            itemId,
            searchTerm
          );
          matches.push(...sheetMatches);
          break;

        case "all_sheets":
          const allSheetsMatches = await this.findInAllSheets(
            graphClient,
            driveId,
            itemId,
            searchTerm
          );
          matches.push(...allSheetsMatches);
          break;

        default:
          throw new AppError(`Invalid scope: ${scope}`, 400);
      }

      // Cache the results
      this.searchCache.set(cacheKey, {
        results: matches,
        ts: Date.now(),
      });

      logger.info(`Found ${matches.length} occurrences of '${searchTerm}'`, {
        driveId,
        itemId,
        scope,
        matchCount: matches.length,
      });

      return matches;
    } catch (err) {
      logger.error("Failed to find occurrences", {
        driveId,
        itemId,
        searchTerm,
        scope,
        error: err.message,
      });
      throw err;
    }
  }

  async findInHeaders(graphClient, driveId, itemId, searchTerm) {
    const matches = [];

    try {
      // Get all worksheets
      const worksheetsResp = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
        .get();

      for (const worksheet of worksheetsResp.value || []) {
        // Get the used range to determine how many columns to check
        const usedRangeResp = await graphClient
          .api(
            `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange`
          )
          .get();

        if (usedRangeResp && usedRangeResp.columnCount > 0) {
          // Read first row only
          const headerRange = `A1:${this.getColumnLetter(
            usedRangeResp.columnCount
          )}1`;
          const headerResp = await graphClient
            .api(
              `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/range(address='${headerRange}')`
            )
            .get();

          if (headerResp.values && headerResp.values[0]) {
            headerResp.values[0].forEach((cellValue, colIndex) => {
              if (
                cellValue &&
                String(cellValue)
                  .toLowerCase()
                  .includes(searchTerm.toLowerCase())
              ) {
                matches.push({
                  sheet: worksheet.name,
                  sheetId: worksheet.id,
                  cell: `${this.getColumnLetter(colIndex + 1)}1`,
                  value: cellValue,
                  oldValue: cellValue,
                  isHeader: true,
                });
              }
            });
          }
        }
      }
    } catch (err) {
      logger.warn("Failed to search in headers", { error: err.message });
    }

    return matches;
  }

  async findInRange(graphClient, driveId, itemId, searchTerm, rangeSpec) {
    const matches = [];

    try {
      // Parse range specification (e.g., "Sheet1!A2:D20" or "A2:D20")
      const { sheetName, address } =
        resolverService.parseSheetAndAddress(rangeSpec);

      if (sheetName) {
        // Specific sheet and range
        const rangeResp = await graphClient
          .api(
            `/drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/range(address='${address}')`
          )
          .get();

        this.extractMatchesFromRange(rangeResp, searchTerm, sheetName, matches);
      } else {
        // Range in first worksheet
        const worksheetsResp = await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
          .get();

        if (worksheetsResp.value && worksheetsResp.value.length > 0) {
          const firstSheet = worksheetsResp.value[0];
          const rangeResp = await graphClient
            .api(
              `/drives/${driveId}/items/${itemId}/workbook/worksheets/${firstSheet.id}/range(address='${address}')`
            )
            .get();

          this.extractMatchesFromRange(
            rangeResp,
            searchTerm,
            firstSheet.name,
            matches
          );
        }
      }
    } catch (err) {
      logger.warn("Failed to search in specific range", {
        rangeSpec,
        error: err.message,
      });
    }

    return matches;
  }

  async findInSheet(
    graphClient,
    driveId,
    itemId,
    searchTerm,
    sheetName = null
  ) {
    const matches = [];

    try {
      let worksheets = [];

      if (sheetName) {
        // Search in specific sheet
        const worksheetResp = await graphClient
          .api(
            `/drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')`
          )
          .get();
        worksheets = [worksheetResp];
      } else {
        // Search in first sheet
        const worksheetsResp = await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
          .get();
        worksheets = worksheetsResp.value ? [worksheetsResp.value[0]] : [];
      }

      for (const worksheet of worksheets) {
        const usedRangeResp = await graphClient
          .api(
            `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange`
          )
          .get();

        if (usedRangeResp && usedRangeResp.values) {
          this.extractMatchesFromRange(
            usedRangeResp,
            searchTerm,
            worksheet.name,
            matches
          );
        }
      }
    } catch (err) {
      logger.warn("Failed to search in sheet", {
        sheetName,
        error: err.message,
      });
    }

    return matches;
  }

  async findInAllSheets(graphClient, driveId, itemId, searchTerm) {
    const matches = [];

    try {
      const worksheetsResp = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
        .get();

      for (const worksheet of worksheetsResp.value || []) {
        try {
          const usedRangeResp = await graphClient
            .api(
              `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange`
            )
            .get();

          if (usedRangeResp && usedRangeResp.values) {
            this.extractMatchesFromRange(
              usedRangeResp,
              searchTerm,
              worksheet.name,
              matches
            );
          }
        } catch (sheetErr) {
          logger.warn(`Failed to search in sheet ${worksheet.name}`, {
            error: sheetErr.message,
          });
        }
      }
    } catch (err) {
      logger.warn("Failed to search in all sheets", { error: err.message });
    }

    return matches;
  }

  extractMatchesFromRange(rangeResp, searchTerm, sheetName, matches) {
    if (!rangeResp.values) return;

    const startRow = rangeResp.address.match(/:([A-Z]+)(\d+)/)?.[2] || 1;
    const startCol = rangeResp.address.match(/([A-Z]+)(\d+)/)?.[1] || "A";
    const startColIndex = this.getColumnIndex(startCol);
    const startRowIndex = parseInt(startRow);

    rangeResp.values.forEach((row, rowIndex) => {
      row.forEach((cellValue, colIndex) => {
        if (
          cellValue &&
          String(cellValue).toLowerCase().includes(searchTerm.toLowerCase())
        ) {
          const actualRow = startRowIndex + rowIndex;
          const actualCol = this.getColumnLetter(startColIndex + colIndex);

          matches.push({
            sheet: sheetName,
            cell: `${actualCol}${actualRow}`,
            value: cellValue,
            oldValue: cellValue,
            isHeader: actualRow === 1,
          });
        }
      });
    });
  }

  async performReplace(
    accessToken,
    driveId,
    itemId,
    searchTerm,
    replaceTerm,
    matches,
    options = {}
  ) {
    const { highlightChanges = false, logChanges = true } = options;
    const changes = [];
    const errors = [];

    try {
      const graphClient = this.createGraphClient(accessToken);

      // Group matches by sheet for batch operations
      const matchesBySheet = this.groupMatchesBySheet(matches);

      for (const [sheetName, sheetMatches] of matchesBySheet.entries()) {
        try {
          const sheetChanges = await this.replaceInSheet(
            graphClient,
            driveId,
            itemId,
            sheetName,
            sheetMatches,
            searchTerm,
            replaceTerm,
            highlightChanges
          );

          changes.push(...sheetChanges);
        } catch (sheetErr) {
          logger.error(`Failed to replace in sheet ${sheetName}`, {
            error: sheetErr.message,
          });
          errors.push({
            sheet: sheetName,
            error: sheetErr.message,
          });
        }
      }

      // Log the operation
      if (logChanges && changes.length > 0) {
        auditService.logSystemEvent({
          event: "FIND_REPLACE_OPERATION",
          details: {
            driveId,
            itemId,
            searchTerm,
            replaceTerm,
            changesCount: changes.length,
            highlightChanges,
            changes: changes.slice(0, 100), // Limit logged changes to prevent huge logs
          },
        });
      }

      return {
        changes,
        errors,
        summary: {
          totalMatches: matches.length,
          successful: changes.length,
          failed: errors.length,
        },
      };
    } catch (err) {
      logger.error("Failed to perform replace operation", {
        driveId,
        itemId,
        searchTerm,
        replaceTerm,
        error: err.message,
      });
      throw err;
    }
  }

  async replaceInSheet(
    graphClient,
    driveId,
    itemId,
    sheetName,
    matches,
    searchTerm,
    replaceTerm,
    highlightChanges
  ) {
    const changes = [];

    // Batch update values
    const updates = matches.map((match) => {
      const newValue = String(match.value).replace(
        new RegExp(searchTerm, "gi"),
        replaceTerm
      );

      return {
        cell: match.cell,
        oldValue: match.value,
        newValue: newValue,
      };
    });

    // Update cell values in batches
    for (let i = 0; i < updates.length; i += 10) {
      // Process 10 cells at a time
      const batch = updates.slice(i, i + 10);

      for (const update of batch) {
        try {
          // Update cell value
          await graphClient
            .api(
              `/drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/range(address='${update.cell}')`
            )
            .patch({
              values: [[update.newValue]],
            });

          // Apply highlighting if requested
          if (highlightChanges) {
            await this.highlightCell(
              graphClient,
              driveId,
              itemId,
              sheetName,
              update.cell
            );
          }

          changes.push({
            sheet: sheetName,
            cell: update.cell,
            oldValue: update.oldValue,
            newValue: update.newValue,
          });
        } catch (cellErr) {
          logger.warn(`Failed to update cell ${update.cell}`, {
            error: cellErr.message,
          });
        }
      }
    }

    return changes;
  }

  async highlightCell(graphClient, driveId, itemId, sheetName, cellAddress) {
    try {
      await graphClient
        .api(
          `/drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/range(address='${cellAddress}')/format/fill`
        )
        .patch({
          color: "#FFFF00", // Yellow background
        });
    } catch (err) {
      logger.warn(`Failed to highlight cell ${cellAddress}`, {
        error: err.message,
      });
    }
  }

  groupMatchesBySheet(matches) {
    const grouped = new Map();

    matches.forEach((match) => {
      if (!grouped.has(match.sheet)) {
        grouped.set(match.sheet, []);
      }
      grouped.get(match.sheet).push(match);
    });

    return grouped;
  }

  getColumnLetter(columnIndex) {
    let result = "";
    while (columnIndex > 0) {
      columnIndex--;
      result = String.fromCharCode(65 + (columnIndex % 26)) + result;
      columnIndex = Math.floor(columnIndex / 26);
    }
    return result;
  }

  getColumnIndex(columnLetter) {
    let result = 0;
    for (let i = 0; i < columnLetter.length; i++) {
      result = result * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    return result;
  }

  generatePreview(matches, searchTerm) {
    const preview = {
      searchTerm,
      totalMatches: matches.length,
      breakdown: {
        headers: matches.filter((m) => m.isHeader).length,
        dataRows: matches.filter((m) => !m.isHeader).length,
      },
      bySheet: {},
      samples: matches.slice(0, 10), // Show first 10 matches as samples
    };

    // Group by sheet for breakdown
    matches.forEach((match) => {
      if (!preview.bySheet[match.sheet]) {
        preview.bySheet[match.sheet] = 0;
      }
      preview.bySheet[match.sheet]++;
    });

    return preview;
  }

  async findOccurrencesByName(
    accessToken,
    {
      driveName,
      fileName,
      searchTerm,
      scope = "entire_sheet",
      rangeSpec = null,
      itemPath = null,
    }
  ) {
    if (!driveName || !fileName || !searchTerm) {
      throw new AppError(
        "driveName, fileName, and searchTerm are required",
        400
      );
    }

    // Resolve IDs from names
    const driveId = await resolverService.resolveDriveIdByName(
      accessToken,
      driveName
    );

    let itemId;
    if (itemPath) {
      itemId = await resolverService.resolveItemIdByPath(
        accessToken,
        driveId,
        fileName,
        itemPath
      );
    } else {
      itemId = await resolverService.resolveItemIdByName(
        accessToken,
        driveId,
        fileName
      );
    }

    return this.findOccurrences(
      accessToken,
      driveId,
      itemId,
      searchTerm,
      scope,
      rangeSpec
    );
  }

  async performReplaceByName(
    accessToken,
    {
      driveName,
      fileName,
      searchTerm,
      replaceTerm,
      matches,
      options = {},
      itemPath = null,
    }
  ) {
    if (!driveName || !fileName || !searchTerm || replaceTerm === undefined) {
      throw new AppError(
        "driveName, fileName, searchTerm, and replaceTerm are required",
        400
      );
    }

    const driveId = await resolverService.resolveDriveIdByName(
      accessToken,
      driveName
    );

    let itemId;
    if (itemPath) {
      itemId = await resolverService.resolveItemIdByPath(
        accessToken,
        driveId,
        fileName,
        itemPath
      );
    } else {
      itemId = await resolverService.resolveItemIdByName(
        accessToken,
        driveId,
        fileName
      );
    }

    return this.performReplace(
      accessToken,
      driveId,
      itemId,
      searchTerm,
      replaceTerm,
      matches,
      options
    );
  }
}

module.exports = new FindReplaceService();
