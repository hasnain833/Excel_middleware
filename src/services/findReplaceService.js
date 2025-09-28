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

  // New: Generic label matcher utility
  _normalizeLabelText(s, stripColons = true) {
    if (s == null) return "";
    let out = String(s);
    if (stripColons) out = out.replace(/:+\s*$/, "");
    return out.trim();
  }

  _fuzzySimilarity(a, b) {
    // Simple normalized Levenshtein similarity (1 - distance/maxLen)
    const s1 = a.toLowerCase();
    const s2 = b.toLowerCase();
    const m = s1.length;
    const n = s2.length;
    if (m === 0 && n === 0) return 1;
    const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
        dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
      }
    }
    const dist = dp[m][n];
    return 1 - dist / Math.max(m, n);
  }

  _labelMatches(cellText, labels, { labelMode = "exact", caseSensitiveLabel = false, stripColons = true, fuzzyThreshold = 0.85 }) {
    if (!cellText) return false;
    const txt = this._normalizeLabelText(cellText, stripColons);
    const candidates = Array.isArray(labels) ? labels : [labels];
    if (candidates.length === 0) return false;
    for (const lab of candidates) {
      if (lab == null || lab === "") continue;
      const normLab = this._normalizeLabelText(lab, stripColons);
      if (labelMode === "regex") {
        try {
          const flags = caseSensitiveLabel ? "" : "i";
          const re = new RegExp(normLab, flags);
          if (re.test(txt)) return true;
        } catch (_) {
          // ignore invalid regex; treat as non-match
        }
      } else if (labelMode === "fuzzy") {
        const sim = this._fuzzySimilarity(caseSensitiveLabel ? txt : txt.toLowerCase(), caseSensitiveLabel ? normLab : normLab.toLowerCase());
        if (sim >= fuzzyThreshold) return true;
      } else {
        // exact
        if (caseSensitiveLabel) {
          if (txt === normLab) return true;
        } else if (txt.toLowerCase() === normLab.toLowerCase()) {
          return true;
        }
      }
    }
    return false;
  }

  async _readUsedRange(graphClient, driveId, itemId, worksheetId) {
    const used = await graphClient
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/usedRange(valuesOnly=true)`) // values only
      .get();
    const values = used.values || [];
    const startRow = used.address?.match(/:([A-Z]+)(\d+)/)?.[2] || used.address?.match(/([A-Z]+)(\d+)/)?.[2] || 1;
    const startCol = used.address?.match(/([A-Z]+)(\d+)/)?.[1] || "A";
    const startColIndex = this.getColumnIndex(startCol);
    const startRowIndex = parseInt(startRow);
    return { values, startColIndex, startRowIndex, address: used.address };
  }

  // New: labelNeighbor strategy
  async findLabelNeighborMatches(accessToken, driveId, itemId, sheetScope, opts = {}) {
    const {
      label = [],
      labelMode = "exact",
      caseSensitiveLabel = false,
      stripColons = true,
      fuzzyThreshold = 0.85,
      directions = ["down", "right"],
      maxDown = 3,
      maxRight = 3,
      valueSearchTerm,
    } = opts;

    const graphClient = this.createGraphClient(accessToken);
    const matches = [];
    const worksheetsResp = await graphClient
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
      .get();

    const worksheets = worksheetsResp.value || [];
    const targetSheets = (() => {
      if (!sheetScope || sheetScope === "ALL") return worksheets;
      const byName = worksheets.find((w) => w.name === sheetScope);
      return byName ? [byName] : [];
    })();

    for (const ws of targetSheets) {
      try {
        const { values, startColIndex, startRowIndex } = await this._readUsedRange(
          graphClient,
          driveId,
          itemId,
          ws.id
        );
        for (let r = 0; r < values.length; r++) {
          const row = values[r] || [];
          for (let c = 0; c < row.length; c++) {
            const cellVal = row[c];
            if (!this._labelMatches(cellVal, label, { labelMode, caseSensitiveLabel, stripColons, fuzzyThreshold })) continue;

            // search neighbors per directions
            let target = null;
            for (const dir of directions) {
              if (dir === "down") {
                for (let k = 1; k <= maxDown; k++) {
                  const rr = r + k;
                  if (rr >= values.length) break;
                  const nbrVal = (values[rr] || [])[c];
                  if (nbrVal !== undefined && nbrVal !== null && String(nbrVal).length > 0) {
                    target = { rr, cc: c, nbrVal };
                    break;
                  }
                }
              } else if (dir === "right") {
                for (let k = 1; k <= maxRight; k++) {
                  const cc = c + k;
                  const nbrVal = (values[r] || [])[cc];
                  if (nbrVal !== undefined && nbrVal !== null && String(nbrVal).length > 0) {
                    target = { rr: r, cc, nbrVal };
                    break;
                  }
                }
              }
              if (target) break;
            }

            if (!target) continue; // no neighbor found within limits

            // Optional filter by current value content
            if (valueSearchTerm) {
              const hay = String(target.nbrVal);
              if (!hay.toLowerCase().includes(String(valueSearchTerm).toLowerCase())) {
                continue;
              }
            }

            const actualRow = startRowIndex + (target.rr ?? r);
            const labelCol = this.getColumnLetter(startColIndex + c);
            const targetCol = this.getColumnLetter(startColIndex + (target.cc ?? c));
            matches.push({
              sheet: ws.name,
              sheetId: ws.id,
              cell: `${targetCol}${actualRow}`,
              value: target.nbrVal,
              oldValue: target.nbrVal,
              labelText: this._normalizeLabelText(cellVal, stripColons),
              labelAddress: `${labelCol}${startRowIndex + r}`,
              context: `Near ${labelCol}${startRowIndex + r}`,
            });
          }
        }
      } catch (e) {
        logger.warn("labelNeighbor scan failed for sheet", { sheet: ws.name, error: e.message });
      }
    }

    return matches;
  }

  async performLabelNeighborUpdate(accessToken, driveId, itemId, matches, newValue, options = {}) {
    const { highlightChanges = false } = options;
    const graphClient = this.createGraphClient(accessToken);
    const changes = [];
    const errors = [];

    // Build $batch requests in chunks of 20
    const chunkSize = 20;
    for (let i = 0; i < matches.length; i += chunkSize) {
      const batchChunk = matches.slice(i, i + chunkSize);
      const requests = batchChunk.map((m, idx) => ({
        id: String(idx + 1),
        method: "PATCH",
        url: `/drives/${driveId}/items/${itemId}/workbook/worksheets('${m.sheet}')/range(address='${m.cell}')`,
        headers: { "Content-Type": "application/json" },
        body: { values: [[newValue]] },
      }));
      try {
        const resp = await graphClient.api("/$batch").post({ requests });
        // Record changes; Graph returns responses array in same order
        const responses = resp?.responses || [];
        responses.forEach((r, idx) => {
          const m = batchChunk[idx];
          if (r.status >= 200 && r.status < 300) {
            changes.push({ sheet: m.sheet, cell: m.cell, oldValue: m.oldValue, newValue });
          } else {
            errors.push({ sheet: m.sheet, cell: m.cell, error: r.body?.error?.message || `HTTP ${r.status}` });
          }
        });
      } catch (err) {
        // If batch fails altogether, attempt individual updates to salvage some
        for (const m of batchChunk) {
          try {
            await graphClient
              .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets('${m.sheet}')/range(address='${m.cell}')`)
              .patch({ values: [[newValue]] });
            changes.push({ sheet: m.sheet, cell: m.cell, oldValue: m.oldValue, newValue });
          } catch (e) {
            errors.push({ sheet: m.sheet, cell: m.cell, error: e.message });
          }
        }
      }
    }

    // Optional highlighting (non-batched)
    if (highlightChanges && changes.length > 0) {
      for (const ch of changes) {
        try {
          await this.highlightCell(graphClient, driveId, itemId, ch.sheet, ch.cell);
        } catch (_) {}
      }
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
    rangeSpec = null,
    sheetName = null
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
            searchTerm,
            sheetName || null
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
            `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange(valuesOnly=true)`
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
              `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange(valuesOnly=true)`
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

  async getWorksheetsMap(graphClient, driveId, itemId) {
    const resp = await graphClient
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
      .get();
    const byName = new Map();
    const byId = new Map();
    for (const ws of resp.value || []) {
      byName.set(ws.name, ws.id);
      byId.set(ws.id, ws.name);
    }
    return { byName, byId };
  }

  buildMatchId(sheetId, address) {
    // Standardized selectable match id for preview/apply
    // Will be expanded with sheet name by generateSelectablePreview()
    return `m:${sheetId}:${address}`;
  }

  generateSelectablePreview(matches, searchTerm, sheetsByName) {
    // Ensure consistent shape and include matchId; try to attach sheetId when available
    const items = matches.map((m) => {
      const sheetId = m.sheetId || (sheetsByName && sheetsByName.get(m.sheet));
      const sheetName = m.sheet;
      const base = {
        matchId: sheetId
          ? `m:${sheetId}:${sheetName}:${m.cell || m.address}`
          : `m:${sheetName}:${m.cell || m.address}`,
        sheet: m.sheet,
        sheetId: sheetId,
        address: m.cell || m.address,
        currentValue: m.value ?? m.currentValue ?? m.oldValue,
      };
      if (m.labelText) {
        base.labelText = m.labelText;
        base.labelAddress = m.labelAddress;
        base.context = m.context;
      }
      return base;
    });
    return {
      searchTerm,
      totalMatches: items.length,
      matches: items,
    };
  }

  async findEntityNameMatches(accessToken, driveId, itemId, sheetScope) {
    const graphClient = this.createGraphClient(accessToken);
    const matches = [];
    const labelRegex = /^(entity|entity\s*name)$/i;
    const worksheetsResp = await graphClient
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
      .get();
    const worksheets = worksheetsResp.value || [];

    const targetSheets = (() => {
      if (!sheetScope || sheetScope === "ALL") return worksheets;
      const byName = worksheets.find((w) => w.name === sheetScope);
      return byName ? [byName] : [];
    })();

    for (const ws of targetSheets) {
      try {
        const used = await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${ws.id}/usedRange`)
          .get();
        const values = used.values || [];
        const startRow = used.address.match(/:([A-Z]+)(\d+)/)?.[2] || 1;
        const startCol = used.address.match(/([A-Z]+)(\d+)/)?.[1] || "A";
        const startColIndex = this.getColumnIndex(startCol);
        const startRowIndex = parseInt(startRow);
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < (values[r] || []).length; c++) {
            const cellVal = values[r][c];
            if (cellVal && String(cellVal).trim().match(labelRegex)) {
              // Neighbor to the right is target value cell
              const nbrC = c + 1;
              const targetVal = values[r][nbrC];
              const actualRow = startRowIndex + r;
              const labelCol = this.getColumnLetter(startColIndex + c);
              const targetCol = this.getColumnLetter(startColIndex + nbrC);
              matches.push({
                sheet: ws.name,
                sheetId: ws.id,
                cell: `${targetCol}${actualRow}`,
                value: targetVal,
                oldValue: targetVal,
                labelText: String(cellVal),
                labelAddress: `${labelCol}${actualRow}`,
                context: `Row ${actualRow} near ${labelCol}${actualRow}`,
              });
            }
          }
        }
      } catch (e) {
        logger.warn("EntityName scan failed for sheet", { sheet: ws.name, error: e.message });
      }
    }
    return matches;
  }

  async performEntityValueUpdate(accessToken, driveId, itemId, matches, newValue, options = {}) {
    const { highlightChanges = false } = options;
    const graphClient = this.createGraphClient(accessToken);
    const changes = [];
    const errors = [];
    for (const m of matches) {
      try {
        await graphClient
          .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets('${m.sheet}')/range(address='${m.cell}')`)
          .patch({ values: [[newValue]] });
        if (highlightChanges) {
          await this.highlightCell(graphClient, driveId, itemId, m.sheet, m.cell);
        }
        changes.push({ sheet: m.sheet, cell: m.cell, oldValue: m.oldValue, newValue });
      } catch (err) {
        errors.push({ sheet: m.sheet, cell: m.cell, error: err.message });
      }
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
            highlightChanges,
            options
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
    highlightChanges,
    options = {}
  ) {
    const changes = [];

    const {
      caseSensitive = false,
      wholeWord = false,
      replaceInside = true,
      replaceMode = "all",
    } = options;

    // Build regex from searchTerm and knobs
    const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    let pattern = escapeRegExp(String(searchTerm));
    if (wholeWord) {
      pattern = `\\b${pattern}\\b`;
    }
    if (!replaceInside) {
      pattern = `^${pattern}$`;
    }
    let flags = caseSensitive ? "" : "i";
    if (replaceMode === "all") flags += "g";
    const re = new RegExp(pattern, flags);

    // Batch update values
    const updates = matches.map((match) => {
      const strVal = match.value == null ? "" : String(match.value);
      const newValue = strVal.replace(re, replaceTerm);

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
