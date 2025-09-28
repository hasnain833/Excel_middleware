const excelService = require("../services/excelService");
const resolverService = require("../services/resolverService");
const nameResolutionMixin = require("../middleware/nameResolutionMixin");
const auditService = require("../services/auditService");
const logger = require("../config/logger");
const { catchAsync } = require("../middleware/errorHandler");
const { AppError } = require("../middleware/errorHandler");

class ExcelController {
  getWorkbooks = catchAsync(async (req, res) => {
    const auditContext = auditService.createAuditContext(req);
    const workbooksResponse = await excelService.getWorkbooks(
      req.accessToken,
      auditContext
    );
    const safeData = Array.isArray(workbooksResponse?.value)
      ? workbooksResponse.value
      : workbooksResponse;

    res.json({
      status: "success",
      data: safeData,
    });
  });

  clearData = catchAsync(async (req, res) => {
    const auditContext = auditService.createAuditContext(req);
    const nameParams = nameResolutionMixin.extractNameParams(req);
    nameResolutionMixin.validateNameInput(nameParams);

    const resolution = await nameResolutionMixin.resolveNames(req, nameParams);
    if (!resolution.itemId) {
      throw new AppError('Could not resolve file. Please check the file name and path.', 404);
    }

    // Optional worksheet and range
    let resolvedWorksheetId = resolution.sheetId;
    if (!resolvedWorksheetId && (req.body?.worksheetName || req.body?.sheetName)) {
      const wsName = req.body.worksheetName || req.body.sheetName;
      resolvedWorksheetId = await resolverService.resolveWorksheetIdByName(
        req.accessToken,
        resolution.driveId,
        resolution.itemId,
        wsName
      );
    }

    const result = await excelService.clearData({
      accessToken: req.accessToken,
      driveId: resolution.driveId,
      itemId: resolution.itemId,
      worksheetId: resolvedWorksheetId,
      range: req.body?.range,
      auditContext,
    });

    res.json({ status: 'success', data: result, resolution: nameResolutionMixin.getResolutionSummary(resolution) });
  });

  getWorksheets = catchAsync(async (req, res) => {
    const { driveId, itemId, driveName, itemName, itemPath } = req.query;
    const auditContext = auditService.createAuditContext(req);
    let resolvedDriveId = driveId;
    let resolvedItemId = itemId;

    if ((!resolvedDriveId || !resolvedItemId) && driveName && itemName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(
        req.accessToken,
        driveName
      );

      try {
        resolvedItemId = await resolverService.resolveItemIdByName(
          req.accessToken,
          resolvedDriveId,
          itemName
        );
      } catch (err) {
        if (err.isMultipleMatches && itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(
            req.accessToken,
            resolvedDriveId,
            itemName,
            itemPath
          );
        } else {
          throw err; // Re-throw the original error
        }
      }
    }

    const worksheets = await excelService.getWorksheets(
      req.accessToken,
      resolvedDriveId,
      resolvedItemId,
      auditContext
    );

    res.json({
      status: "success",
      data: worksheets,
    });
  });
  readRange = catchAsync(async (req, res) => {
    const auditContext = auditService.createAuditContext(req);
    const nameParams = nameResolutionMixin.extractNameParams(req);
    nameResolutionMixin.validateNameInput(nameParams);

    // Helper: normalize sheet names
    const normalizeSheetName = (s = "") =>
      String(s)
        .replace(/\s+/g, " ")
        .replace(/\.+$/, "")
        .trim()
        .toLowerCase();

    // A1 helpers from findReplaceService
    const findReplaceService = require("../services/findReplaceService");

    try {
      // Resolve names to IDs with backward compatibility
      const resolution = await nameResolutionMixin.resolveNames(req, nameParams);

      if (!resolution.itemId) {
        throw new AppError(
          "Could not resolve file. Please check the file name and path.",
          404
        );
      }

      // Log name resolution
      nameResolutionMixin.logNameResolution(resolution, "READ_RANGE", {
        range: req.body.range,
        worksheetName: req.body.worksheetName || req.body.sheetName,
      });

      const explicitMode = req.body.mode;
      const explicitProjection = req.body.projection;
      const includeFormulas = req.body.includeFormulas === true;
      const includeText = req.body.includeText === true;
      const includeFormats = req.body.includeFormats === true; // reserved for future use
      const valuesOnly = req.body.valuesOnly !== false; // default true
      const summary = req.body.summary === true;
      const paginate = req.body.paginate || {};

      // Extract range parts if provided like Sheet1!A1:D10
      const rawRange = req.body.range;
      let parsedSheetFromRange = undefined;
      let addressFromRange = rawRange;
      if (typeof rawRange === "string" && rawRange.length > 0) {
        const parsed = resolverService.parseSheetAndAddress(rawRange);
        parsedSheetFromRange = parsed.sheetName;
        addressFromRange = parsed.address;
      }

      const requestedSheetName =
        req.body.worksheetName || req.body.sheetName || parsedSheetFromRange || resolution.sheetName;

      // Backward compatibility: if request has sheet + range and no mode/projection → behave as before
      const isLegacyRange =
        !!(requestedSheetName && rawRange && !explicitMode && !explicitProjection);
      if (isLegacyRange) {
        let resolvedWorksheetId = resolution.sheetId;
        if (!resolvedWorksheetId && requestedSheetName) {
          resolvedWorksheetId = await resolverService.resolveWorksheetIdByName(
            req.accessToken,
            resolution.driveId,
            resolution.itemId,
            requestedSheetName
          );
        }
        const legacyData = await excelService.readRange({
          accessToken: req.accessToken,
          driveId: resolution.driveId,
          itemId: resolution.itemId,
          worksheetId: resolvedWorksheetId,
          range: addressFromRange,
          auditContext,
        });
        return res.json({
          status: "success",
          data: {
            range: legacyData.address,
            values: legacyData.values,
            formulas: legacyData.formulas,
            text: legacyData.text,
            dimensions: {
              rows: legacyData.rowCount,
              columns: legacyData.columnCount,
            },
          },
          resolution: nameResolutionMixin.getResolutionSummary(resolution),
        });
      }

      // Determine effective mode and projection
      let mode = explicitMode;
      if (!mode) {
        mode = rawRange ? "range" : requestedSheetName ? "sheet" : "workbook";
      }
      const projection = explicitProjection || "matrix";

      const graphClient = excelService.createGraphClient(req.accessToken);

      // Utility: sheet map with normalization and 409 candidates if not found
      const ensureWorksheet = async () => {
        const { byName, byId } = await findReplaceService.getWorksheetsMap(
          graphClient,
          resolution.driveId,
          resolution.itemId
        );
        // Build normalized lookup
        const normToActual = new Map();
        for (const name of byName.keys()) {
          normToActual.set(normalizeSheetName(name), name);
        }
        const candidates = Array.from(byName.keys());
        if (!requestedSheetName) {
          return { worksheetId: null, worksheetName: candidates[0] || null, maps: { byName, byId } };
        }
        const actual = normToActual.get(normalizeSheetName(requestedSheetName));
        if (!actual) {
          // 409 with candidates array
          return res.status(409).json({
            status: "multiple_matches",
            data: { candidates },
          });
        }
        return { worksheetId: byName.get(actual), worksheetName: actual, maps: { byName, byId } };
      };

      // Utility: read a range from a sheet (sheet-scoped)
      const readSheetRange = async (worksheetId, address) => {
        // If caller passed a sheet-qualified address while we have worksheetId, strip the prefix
        const addr = typeof address === "string" ? resolverService.parseSheetAndAddress(address).address : undefined;
        const select = ["address", "values"]; // always need values
        if (includeFormulas) select.push("formulas");
        if (includeText) select.push("text");
        const qs = select.length ? `?$select=${select.join(",")}` : "";
        const url = addr
          ? `/drives/${resolution.driveId}/items/${resolution.itemId}/workbook/worksheets('${worksheetId}')/range(address='${addr}')${qs}`
          : `/drives/${resolution.driveId}/items/${resolution.itemId}/workbook/worksheets/${worksheetId}/usedRange(valuesOnly=${valuesOnly})${qs}`;
        return await graphClient.api(url).get();
      };

      // Build projections
      const buildMatrix = (resp) => {
        const out = {
          usedRange: resp.address,
          values: resp.values || [],
        };
        if (includeFormulas) out.formulas = resp.formulas || [];
        if (includeText) out.text = resp.text || [];
        return out;
      };

      const buildCells = (resp) => {
        const values = resp.values || [];
        const start = findReplaceService._parseStartFromAddress(resp.address);
        const cells = [];
        for (let r = 0; r < values.length; r++) {
          const row = values[r] || [];
          for (let c = 0; c < row.length; c++) {
            const rowA1 = (start.startRowIndex || 1) + r;
            const colA1 = findReplaceService.getColumnLetter((start.startColIndex || 1) + c);
            const cell = {
              address: `${colA1}${rowA1}`,
              value: row[c],
            };
            if (includeFormulas && Array.isArray(resp.formulas) && resp.formulas[r]) {
              cell.formula = resp.formulas[r][c] ?? null;
            } else {
              cell.formula = null;
            }
            if (includeText && Array.isArray(resp.text) && resp.text[r]) {
              cell.text = resp.text[r][c] ?? null;
            } else {
              cell.text = null;
            }
            cells.push(cell);
          }
        }
        // Pagination
        const pageSize = Math.max(1, Math.min(10000, paginate.pageSize || cells.length));
        const offset = paginate.pageToken ? parseInt(paginate.pageToken, 10) || 0 : 0;
        const slice = cells.slice(offset, offset + pageSize);
        const nextOffset = offset + pageSize;
        const hasMore = nextOffset < cells.length;
        return { cells: slice, page: { hasMore, nextPageToken: hasMore ? String(nextOffset) : null } };
      };

      const buildRecords = (resp) => {
        const values = (resp.values || []).slice();
        if (!values.length) return { records: [] };
        // Find first non-empty row as header
        let header = [];
        let startIdx = 0;
        for (let i = 0; i < values.length; i++) {
          const row = values[i] || [];
          const any = row.some((v) => v !== null && v !== undefined && String(v).trim() !== "");
          if (any) {
            header = row.map((k) => String(k || "").replace(/\s+/g, " ").trim());
            startIdx = i + 1;
            break;
          }
        }
        const records = [];
        for (let i = startIdx; i < values.length; i++) {
          const row = values[i] || [];
          const obj = {};
          for (let j = 0; j < header.length; j++) {
            const key = header[j] || `col_${j + 1}`;
            obj[key] = row[j];
          }
          records.push(obj);
        }
        return { records };
      };

      const buildKv = (resp) => {
        const values = resp.values || [];
        const start = findReplaceService._parseStartFromAddress(resp.address);
        const kv = [];
        for (let r = 0; r < values.length; r++) {
          const row = values[r] || [];
          for (let c = 0; c < row.length; c++) {
            const v = row[c];
            if (typeof v === "string" && v.trim().length > 0) {
              // Prefer below, else right
              let value = null;
              let valR = r + 1;
              let valC = c;
              if (valR < values.length && (values[valR] || [])[valC] != null && String((values[valR] || [])[valC]).trim() !== "") {
                value = (values[valR] || [])[valC];
              } else if ((row[c + 1] != null) && String(row[c + 1]).trim() !== "") {
                value = row[c + 1];
                valR = r;
                valC = c + 1;
              }
              if (value != null) {
                const labelAddress = `${findReplaceService.getColumnLetter((start.startColIndex || 1) + c)}${(start.startRowIndex || 1) + r}`;
                const valueAddress = `${findReplaceService.getColumnLetter((start.startColIndex || 1) + valC)}${(start.startRowIndex || 1) + valR}`;
                kv.push({ label: String(v).replace(/:+\s*$/, "").trim(), labelAddress, value, valueAddress });
              }
            }
          }
        }
        return { kv };
      };

      // Execute per mode
      if (mode === "workbook") {
        const { byName } = await findReplaceService.getWorksheetsMap(
          graphClient,
          resolution.driveId,
          resolution.itemId
        );
        const sheets = [];
        const workbookSummary = [];
        for (const [sheetActualName, wsId] of byName.entries()) {
          const resp = await readSheetRange(wsId, undefined);
          const base = { sheet: sheetActualName, usedRange: resp.address };
          if (summary) {
            const rows = Array.isArray(resp.values) ? resp.values.length : 0;
            const cols = rows > 0 ? Math.max(...resp.values.map((r) => (r || []).length)) : 0;
            workbookSummary.push({ sheet: sheetActualName, usedRange: resp.address, rows, cols });
          }
          if (projection === "matrix") sheets.push({ sheet: sheetActualName, usedRange: resp.address, ...buildMatrix(resp) });
          else if (projection === "cells") {
            const c = buildCells(resp);
            sheets.push({ sheet: sheetActualName, usedRange: resp.address, ...c });
          } else if (projection === "records") sheets.push({ sheet: sheetActualName, usedRange: resp.address, ...buildRecords(resp) });
          else if (projection === "kv") sheets.push({ sheet: sheetActualName, usedRange: resp.address, ...buildKv(resp) });
        }
        return res.json({ status: "success", data: { workbookSummary: summary ? workbookSummary : undefined, sheets } });
      }

      if (mode === "sheet") {
        const ensured = await ensureWorksheet();
        if (!ensured || ensured.status) return; // response already sent (409)
        const resp = await readSheetRange(ensured.worksheetId, undefined);
        const base = { sheet: ensured.worksheetName, usedRange: resp.address };
        let payload;
        if (projection === "matrix") payload = { ...base, ...buildMatrix(resp) };
        else if (projection === "cells") payload = { ...base, ...buildCells(resp) };
        else if (projection === "records") payload = { ...base, ...buildRecords(resp) };
        else if (projection === "kv") payload = { ...base, ...buildKv(resp) };
        if (summary) {
          const rows = Array.isArray(resp.values) ? resp.values.length : 0;
          const cols = rows > 0 ? Math.max(...resp.values.map((r) => (r || []).length)) : 0;
          payload.rows = rows;
          payload.cols = cols;
        }
        return res.json({ status: "success", data: payload });
      }

      // mode === 'range'
      {
        const ensured = await ensureWorksheet();
        if (!ensured || ensured.status) return; // 409 already sent
        const resp = await readSheetRange(ensured.worksheetId, addressFromRange);
        const base = { sheet: ensured.worksheetName, usedRange: resp.address };
        let payload;
        if (projection === "matrix") payload = { ...base, ...buildMatrix(resp) };
        else if (projection === "cells") payload = { ...base, ...buildCells(resp) };
        else if (projection === "records") payload = { ...base, ...buildRecords(resp) };
        else if (projection === "kv") payload = { ...base, ...buildKv(resp) };
        if (summary) {
          const rows = Array.isArray(resp.values) ? resp.values.length : 0;
          const cols = rows > 0 ? Math.max(...resp.values.map((r) => (r || []).length)) : 0;
          payload.rows = rows;
          payload.cols = cols;
        }
        return res.json({ status: "success", data: payload });
      }
    } catch (err) {
      if (err.isMultipleMatches) {
        return nameResolutionMixin.handleMultipleMatches(res, err, "file");
      }
      throw err;
    }
  });

  /**
   * Write data to Excel range
   */
  writeRange = catchAsync(async (req, res) => {
    const { driveName, itemName, itemPath, worksheetName, range, values } =
      req.body;
    const auditContext = auditService.createAuditContext(req);

    // Resolve drive/item by names only
    const resolvedDriveId = await resolverService.resolveDriveIdByName(
      req.accessToken,
      driveName
    );
    let resolvedItemId;
    try {
      resolvedItemId = await resolverService.resolveItemIdByName(
        req.accessToken,
        resolvedDriveId,
        itemName
      );
    } catch (err) {
      if (err.isMultipleMatches && itemPath) {
        resolvedItemId = await resolverService.resolveItemIdByPath(
          req.accessToken,
          resolvedDriveId,
          itemName,
          itemPath
        );
      } else {
        throw err;
      }
    }

    // Resolve worksheet and address
    const { sheetName, address } = resolverService.parseSheetAndAddress(range);
    let resolvedWorksheetId = null;
    const effectiveWorksheetName = worksheetName || sheetName;
    if (effectiveWorksheetName) {
      resolvedWorksheetId = await resolverService.resolveWorksheetIdByName(
        req.accessToken,
        resolvedDriveId,
        resolvedItemId,
        effectiveWorksheetName
      );
    }

    const data = await excelService.writeRange({
      accessToken: req.accessToken,
      driveId: resolvedDriveId,
      itemId: resolvedItemId,
      worksheetId: resolvedWorksheetId,
      range: address,
      values,
      auditContext,
    });

    res.json({
      status: "success",
      data: {
        range: data.address,
        values: data.values,
        dimensions: {
          rows: data.rowCount,
          columns: data.columnCount,
        },
      },
    });
  });

  /**
   * Batch operations - perform multiple Excel operations in sequence
   */
  batchOperations = catchAsync(async (req, res) => {
    const { operations } = req.body;
    const auditContext = auditService.createAuditContext(req);

    if (!Array.isArray(operations) || operations.length === 0) {
      return res.status(400).json({
        success: false,
        error: "Invalid operations array",
        timestamp: new Date().toISOString(),
      });
    }

    const results = [];
    const errors = [];

    for (let i = 0; i < operations.length; i++) {
      const operation = operations[i];

      try {
        let result;
        // Resolve names-only per operation
        const opDriveName = operation.driveName;
        const opItemName = operation.itemName;
        const opItemPath = operation.itemPath;
        const opWorksheetName = operation.worksheetName || operation.sheetName;
        const opRange = operation.range;

        if (!opDriveName || !opItemName) {
          throw new AppError(
            `Operation ${i} must include driveName and itemName`,
            400
          );
        }
        const opDriveId = await resolverService.resolveDriveIdByName(
          req.accessToken,
          opDriveName
        );
        let opItemId;
        try {
          opItemId = await resolverService.resolveItemIdByName(
            req.accessToken,
            opDriveId,
            opItemName
          );
        } catch (err) {
          if (err.isMultipleMatches && opItemPath) {
            opItemId = await resolverService.resolveItemIdByPath(
              req.accessToken,
              opDriveId,
              opItemName,
              opItemPath
            );
          } else {
            throw err;
          }
        }
        let opWorksheetId = null;
        let opAddress = opRange;
        if (opRange) {
          const parsed = resolverService.parseSheetAndAddress(opRange);
          opAddress = parsed.address;
          if (parsed.sheetName) {
            opWorksheetId = await resolverService.resolveWorksheetIdByName(
              req.accessToken,
              opDriveId,
              opItemId,
              parsed.sheetName
            );
          }
        }
        if (!opWorksheetId && opWorksheetName) {
          opWorksheetId = await resolverService.resolveWorksheetIdByName(
            req.accessToken,
            opDriveId,
            opItemId,
            opWorksheetName
          );
        }

        switch (operation.type) {
          case "READ_range":
            result = await excelService.readRange({
              accessToken: req.accessToken,
              driveId: opDriveId,
              itemId: opItemId,
              worksheetId: opWorksheetId,
              range: opAddress,
              auditContext,
            });
            break;

          case "write_range":
            result = await excelService.writeRange({
              accessToken: req.accessToken,
              driveId: opDriveId,
              itemId: opItemId,
              worksheetId: opWorksheetId,
              range: opAddress,
              values: operation.values,
              auditContext,
            });
            break;

          default:
            throw new Error(`Unknown operation type: ${operation.type}`);
        }

        results.push({
          index: i,
          operation: operation.type,
          success: true,
          data: result,
        });
      } catch (error) {
        logger.error(`Batch operation ${i} failed:`, error);
        errors.push({
          index: i,
          operation: operation.type,
          error: error.message,
        });

        // Continue with other operations unless it's a critical error
        if (error.message.includes("Authentication")) {
          break; // Stop if authentication fails
        }
      }
    }

    const response = {
      status: errors.length === 0 ? "success" : "partial_success",
      data: {
        results: results,
        errors: errors,
        summary: {
          total: operations.length,
          successful: results.length,
          failed: errors.length,
        },
      },
    };

    // Return 207 Multi-Status if there were partial failures
    const statusCode =
      errors.length > 0 && results.length > 0
        ? 207
        : errors.length === 0
          ? 200
          : 400;

    res.status(statusCode).json(response);
  });


  searchFiles = catchAsync(async (req, res) => {
    const { driveName, fileName, matchMode, excelOnly = false, pathsOnly = false } = req.query;
    const auditContext = auditService.createAuditContext(req);

    if (!fileName) {
      return res.status(400).json({
        status: "error",
        error: {
          code: 400,
          message: "fileName is required",
        },
        timestamp: new Date().toISOString(),
      });
    }

    try {
      // Decide effective matchMode default: if fileName contains a dot -> exact else contains
      const hasDot = String(fileName).includes(".");
      const effMatchMode = matchMode || (hasDot ? "exact" : "contains");

      let allMatches = [];
      const options = { matchMode: effMatchMode, excelOnly: String(excelOnly) === "true" || excelOnly === true };

      if (driveName) {
        // Search within a specific drive only
        const driveId = await resolverService.resolveDriveIdByName(
          req.accessToken,
          driveName
        );
        const graphClient = resolverService.createGraphClient(req.accessToken);
        const matches = await resolverService.recursiveSearchForFile(
          graphClient,
          driveId,
          fileName,
          "",
          "root",
          0,
          20,
          options
        );
        allMatches = matches.map((m) => ({ ...m, driveId, driveName }));
      } else {
        // Enumerate all drives in current site and search each
        const drives = await resolverService.listAllDrives(req.accessToken);
        for (const drv of drives) {
          try {
            const graphClient = resolverService.createGraphClient(req.accessToken);
            const matches = await resolverService.recursiveSearchForFile(
              graphClient,
              drv.id,
              fileName,
              "",
              "root",
              0,
              20,
              options
            );
            allMatches.push(
              ...matches.map((m) => ({ ...m, driveId: drv.id, driveName: drv.name }))
            );
          } catch (e) {
            logger.warn("Drive search failed", { driveId: drv.id, error: e.message });
          }
        }
      }

      // Normalize result shape and paths
      const normalized = allMatches.map((m) => ({
        id: m.id,
        name: m.name,
        path: resolverService.normalizePath(m.path),
        parentId: m.parentId,
        driveId: m.driveId,
        driveName: m.driveName,
      }));

      auditService.logSystemEvent({
        event: "FILE_SEARCH",
        details: {
          scope: driveName ? "single_drive" : "all_drives",
          driveName: driveName || null,
          fileName,
          matchMode: effMatchMode,
          excelOnly: !!options.excelOnly,
          matchCount: normalized.length,
          requestedBy: auditContext.user,
        },
      });

      if (String(pathsOnly) === "true" || pathsOnly === true) {
        return res.json({
          status: "success",
          data: normalized.map((m) => m.path),
        });
      }

      res.json({
        status: "success",
        data: {
          fileName,
          scope: driveName ? { driveName } : { allDrives: true },
          matches: normalized,
          totalMatches: normalized.length,
        },
      });
    } catch (err) {
      logger.error("File search failed:", {
        driveName,
        fileName,
        error: err.message,
      });
      throw err;
    }
  });



  createFile = catchAsync(async (req, res) => {
    const {
      driveId,
      driveName,
      parentPath = "",
      fileName,
    } = req.body;
    const auditContext = auditService.createAuditContext(req);

    // Resolve drive
    let resolvedDriveId = driveId;
    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(
        req.accessToken,
        driveName
      );
    }
    if (!resolvedDriveId) {
      throw new AppError("driveId or driveName is required", 400);
    }

    const graphClient = excelService.createGraphClient(req.accessToken);

    try {
      const basePath = parentPath ? parentPath.replace(/\/$/, "") : "";
      const parentSegment = basePath ? `root:${basePath}:` : "root";
      const resp = await graphClient
        .api(`/drives/${resolvedDriveId}/${parentSegment}/children`)
        .post({
          name: fileName,
          file: {},
          "@microsoft.graph.conflictBehavior": "fail",
        });

      auditService.logSystemEvent({
        event: "FILE_CREATED",
        details: {
          driveId: resolvedDriveId,
          fileName,
          parentPath: basePath,
          requestedBy: auditContext.user,
        },
      });

      return res.status(200).json({
        status: "success",
        data: {
          driveId: resolvedDriveId,
          itemId: resp.id,
          path: `${basePath || ""}/${resp.name}`
            .replace(/\\+/g, "/")
            .replace(/^\/(?=\/)/, "/"),
          name: resp.name,
        },
      });
    } catch (err) {
      if (
        err.statusCode === 409 ||
        err.code === "nameAlreadyExists" ||
        err.status === 409
      ) {
        return res.status(409).json({
          status: "error",
          error: {
            code: 409,
            message: "File already exists at the target path.",
          },
        });
      }
      throw err;
    }
  });

  createSheet = catchAsync(async (req, res) => {
    const {
      driveId,
      driveName,
      itemId,
      itemName,
      itemPath,
      sheetName,
      position,
    } = req.body;
    const auditContext = auditService.createAuditContext(req);

    // Resolve drive and item
    let resolvedDriveId = driveId;
    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(
        req.accessToken,
        driveName
      );
    }

    let resolvedItemId = itemId;
    if (!resolvedItemId && itemName) {
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(
          req.accessToken,
          resolvedDriveId,
          itemName
        );
      } catch (err) {
        if (err.isMultipleMatches && itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(
            req.accessToken,
            resolvedDriveId,
            itemName,
            itemPath
          );
        } else {
          throw err;
        }
      }
    }

    if (!resolvedDriveId || !resolvedItemId) {
      throw new AppError(
        "Unable to resolve drive or file. Provide driveId/driveName and itemId/itemName.",
        400
      );
    }

    const graphClient = excelService.createGraphClient(req.accessToken);

    try {
      const body =
        position === undefined
          ? { name: sheetName }
          : { name: sheetName, position };
      const resp = await graphClient
        .api(
          `/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets/add`
        )
        .post(body);

      auditService.logSystemEvent({
        event: "SHEET_CREATED",
        details: {
          driveId: resolvedDriveId,
          itemId: resolvedItemId,
          sheetName,
          position,
          requestedBy: auditContext.user,
        },
      });

      return res.status(200).json({
        status: "success",
        data: {
          worksheet: { id: resp.id, name: resp.name, position: resp.position },
          file: { itemId: resolvedItemId, name: itemName || undefined },
        },
      });
    } catch (err) {
      if (
        err.statusCode === 400 ||
        err.code === "invalidRequest" ||
        err.status === 400
      ) {
        return res.status(400).json({
          status: "error",
          error: {
            code: 400,
            message: err.message || "Unable to create worksheet.",
          },
        });
      }
      throw err;
    }
  });


  deleteFile = catchAsync(async (req, res) => {
    const { driveId, driveName, itemId, itemName, itemPath, force } = req.body;
    const auditContext = auditService.createAuditContext(req);

    let resolvedDriveId = driveId;
    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(
        req.accessToken,
        driveName
      );
    }

    let resolvedItemId = itemId;
    if (!resolvedItemId && itemName) {
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(
          req.accessToken,
          resolvedDriveId,
          itemName
        );
      } catch (err) {
        if (err.isMultipleMatches && !itemPath) {
          return res.status(409).json({
            status: "multiple_matches",
            data: {
              matches: (err.matches || []).map((m) => ({
                id: m.id,
                name: m.name,
                path: m.path,
                parentId: m.parentId,
              })),
            },
          });
        }
        if (err.isMultipleMatches && itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(
            req.accessToken,
            resolvedDriveId,
            itemName,
            itemPath
          );
        } else {
          throw err;
        }
      }
    }

    if (!resolvedDriveId || !resolvedItemId) {
      throw new AppError("Unable to resolve file to delete", 400);
    }

    const graphClient = excelService.createGraphClient(req.accessToken);

    try {
      await graphClient
        .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}`)
        .delete();

      auditService.logSystemEvent({
        event: "FILE_DELETED",
        details: {
          driveId: resolvedDriveId,
          itemId: resolvedItemId,
          requestedBy: auditContext.user,
          force: !!force,
        },
      });

      return res.status(200).json({
        status: "success",
        data: {
          deleted: true,
          itemId: resolvedItemId,
          name: itemName,
          path: itemPath || null,
        },
      });
    } catch (err) {
      if (err.statusCode === 423 || err.code === "resourceLocked") {
        return res.status(423).json({
          status: "error",
          error: {
            code: 423,
            message:
              "File is locked or in use. Retry later or use force if supported.",
          },
        });
      }
      if (err.statusCode === 403) {
        return res.status(403).json({
          status: "error",
          error: { code: 403, message: "Insufficient permissions to delete this file." },
        });
      }
      if (err.statusCode === 404) {
        return res.status(404).json({
          status: "error",
          error: { code: 404, message: "File not found (it may have been moved or already deleted)." },
        });
      }
      throw err;
    }
  });


  deleteSheet = catchAsync(async (req, res) => {
    const { driveId, driveName, itemId, itemName, itemPath, sheetName } = req.body;
    const auditContext = auditService.createAuditContext(req);

    // Local normalizer (no shared module)
    const normalize = (s) => (s || "").replace(/\.+$/, "").replace(/\s+/g, " ").trim().toLowerCase();

    // Resolve drive and item (prefer IDs when provided)
    let resolvedDriveId = driveId;
    if (!resolvedDriveId && driveName) {
      resolvedDriveId = await resolverService.resolveDriveIdByName(req.accessToken, driveName);
    }

    let resolvedItemId = itemId;
    if (!resolvedItemId && itemName) {
      try {
        resolvedItemId = await resolverService.resolveItemIdByName(
          req.accessToken,
          resolvedDriveId,
          itemName
        );
      } catch (err) {
        if (err.isMultipleMatches && itemPath) {
          resolvedItemId = await resolverService.resolveItemIdByPath(
            req.accessToken,
            resolvedDriveId,
            itemName,
            itemPath
          );
        } else {
          throw err;
        }
      }
    }

    if (!resolvedDriveId || !resolvedItemId) {
      throw new AppError("Unable to resolve workbook for sheet deletion", 400);
    }

    const graphClient = excelService.createGraphClient(req.accessToken);

    try {
      // List worksheets and guard on last-sheet
      const wsResp = await graphClient
        .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets`)
        .get();
      const sheets = wsResp.value || [];
      if (sheets.length <= 1) {
        return res.status(400).json({
          status: "error",
          error: { code: 400, message: "Cannot delete the last remaining worksheet in a workbook." },
        });
      }

      // Normalize and find target
      const byName = new Map();
      for (const s of sheets) byName.set(s.name, s.id);
      const normMap = new Map(Array.from(byName.keys()).map((n) => [normalize(n), n]));
      const targetActual = normMap.get(normalize(sheetName));
      if (!targetActual) {
        return res.status(409).json({ status: "multiple_matches", data: { candidates: Array.from(byName.keys()) } });
      }
      const targetWorksheetId = byName.get(targetActual);

      // Create workbook session (persistChanges: true)
      const sess = await graphClient
        .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/createSession`)
        .post({ persistChanges: true });
      const sessionId = sess && (sess.id || sess.sessionId || sess["id"]);

      // Delete via POST .../delete with session header
      await graphClient
        .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/worksheets('${targetWorksheetId}')/delete`)
        .header("workbook-session-id", sessionId)
        .post({});

      // Best-effort close session
      try {
        await graphClient
          .api(`/drives/${resolvedDriveId}/items/${resolvedItemId}/workbook/closeSession`)
          .header("workbook-session-id", sessionId)
          .post({});
      } catch (e) {
        logger.warn("closeSession failed", { error: e.message });
      }

      auditService.logSystemEvent({
        event: "SHEET_DELETED",
        details: { driveId: resolvedDriveId, itemId: resolvedItemId, sheetName: targetActual, requestedBy: auditContext.user },
      });

      return res.status(200).json({
        status: "success",
        data: { deleted: true, sheetName: targetActual, file: { itemId: resolvedItemId, name: itemName || undefined } },
      });
    } catch (err) {
      if (err.statusCode === 403) {
        return res.status(403).json({ status: "error", error: { code: 403, message: "Insufficient permissions to delete this worksheet." } });
      }
      if (err.statusCode === 404) {
        return res.status(404).json({ status: "error", error: { code: 404, message: "Worksheet not found." } });
      }
      if (err.statusCode === 409 || err.statusCode === 422 || /session/i.test(err.message || "")) {
        return res.status(409).json({ status: "error", error: { code: 409, message: "Workbook session required. Please retry." } });
      }
      throw err;
    }
  });
}

module.exports = new ExcelController();
