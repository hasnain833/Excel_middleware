import express from "express";
import fetch from "node-fetch";
import "dotenv/config";

// Serverless-compatible Express app for Vercel
const app = express();
app.use(express.json({ limit: "1mb" }));

// Health check
app.get("/health", (req, res) => {
  res.status(200).json({
    success: true,
    data: { status: "ok", time: new Date().toISOString() },
  });
});

// Helper: get access token via client credentials
let cachedToken = null;
let tokenExpiresAt = 0; // epoch ms
let refreshPromise = null; // to avoid concurrent refreshes

async function getAccessToken() {
  // Prefer AZURE_* envs, fallback to legacy names if present
  const TENANT_ID = process.env.AZURE_TENANT_ID || process.env.TENANT_ID;
  const CLIENT_ID = process.env.AZURE_CLIENT_ID || process.env.CLIENT_ID;
  const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || process.env.CLIENT_SECRET;

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error(
      "Missing required environment variables: AZURE_TENANT_ID/AZURE_CLIENT_ID/AZURE_CLIENT_SECRET (or TENANT_ID/CLIENT_ID/CLIENT_SECRET)"
    );
  }


  const now = Date.now();
  const safetyWindowMs = 60_000; // refresh 60s before expiry
  if (cachedToken && now < tokenExpiresAt - safetyWindowMs) {
    return cachedToken;
  }

  if (refreshPromise) {
    // Another request is already refreshing the token; await it
    return refreshPromise;
  }

  refreshPromise = (async () => {
    const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    console.log(`[Auth] Fetching new Graph token for tenant ${TENANT_ID}, client ${CLIENT_ID}.`);
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });

    if (!resp.ok) {
      const text = await resp.text();
      console.error(`[Auth] Token request failed (${resp.status}).`);
      throw new Error(`Token request failed (${resp.status}): ${text}`);
    }

    const json = await resp.json();
    const expiresInSec = Number(json.expires_in) || 3600;
    cachedToken = json.access_token;
    tokenExpiresAt = Date.now() + expiresInSec * 1000;
    console.log(`[Auth] Token acquired. Expires in ~${expiresInSec}s.`);
    return cachedToken;
  })();

  try {
    const token = await refreshPromise;
    return token;
  } finally {
    // Ensure we clear the promise so future refreshes can occur
    refreshPromise = null;
  }
}

// Helper: Build Graph base URL for a workbook
function buildWorkbookBase({ driveId, itemId }) {
  if (!driveId || !itemId) {
    throw new Error("driveId and itemId are required");
  }
  return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    driveId
  )}/items/${encodeURIComponent(itemId)}/workbook`;
}

// Helper: Graph fetch with auto token handling and single retry on 401
async function graphFetch(url, options = {}) {
  const makeRequest = async () => {
    const token = await getAccessToken();
    return fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });
  };

  let resp = await makeRequest();
  if (resp.status === 401) {
    // Clear cache and retry once
    console.warn(`[Auth] Received 401 from Graph. Clearing cached token and retrying once...`);
    cachedToken = null;
    tokenExpiresAt = 0;
    resp = await makeRequest();
  }

  const contentType = resp.headers.get("content-type") || "";
  const isJson = contentType.includes("application/json");
  const data = isJson ? await resp.json() : await resp.text();
  if (!resp.ok) {
    const msg = typeof data === "string" ? data : JSON.stringify(data);
    throw new Error(`Graph error (${resp.status}): ${msg}`);
  }
  return data;
}

// Helpers to resolve driveId/itemId from names (case-insensitive)
function getSiteContext(req) {
  // Allows passing site context in either body or query
  const body = req.body || {};
  const query = req.query || {};
  return {
    siteId: body.siteId || query.siteId,
    siteUrl: body.siteUrl || query.siteUrl,
    hostname: body.sharepointHostname || body.hostname || query.sharepointHostname || query.hostname,
    siteName: body.sharepointSiteName || body.siteName || query.sharepointSiteName || query.siteName,
  };
}

async function resolveSiteId(ctx = {}) {
  // 1) Preferred: explicit site ID
  const SITE_ID = ctx.siteId || process.env.SHAREPOINT_SITE_ID || process.env.SITE_ID;
  if (SITE_ID) {
    return { id: SITE_ID };
  }

  // 2) SITE URL: e.g. https://tenant.sharepoint.com/sites/MySite
  const SITE_URL = ctx.siteUrl || process.env.SHAREPOINT_SITE_URL || process.env.SITE_URL;
  if (SITE_URL) {
    try {
      const u = new URL(SITE_URL);
      const hostname = u.hostname; // tenant.sharepoint.com
      // Expect path like /sites/MySite or /teams/MyTeam
      const parts = u.pathname.split('/').filter(Boolean); // ["sites", "MySite"]
      const collection = parts[0]; // sites | teams | etc
      const siteName = parts.slice(1).join('/');
      if (!hostname || !collection || !siteName) {
        throw new Error('Invalid SHAREPOINT_SITE_URL format');
      }
      const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(hostname)}:/${encodeURIComponent(collection)}/${encodeURIComponent(siteName)}?$select=id`;
      return graphFetch(url, { method: "GET" });
    } catch (e) {
      throw new Error(`Invalid SHAREPOINT_SITE_URL. Expected like https://tenant.sharepoint.com/sites/SiteName. Details: ${e.message}`);
    }
  }

  // 3) Legacy: hostname + site name
  const hostname = ctx.hostname || process.env.SHAREPOINT_HOSTNAME;
  const siteName = ctx.siteName || process.env.SHAREPOINT_SITE_NAME;
  if (!hostname || !siteName) {
    const missing = [];
    if (!SITE_ID) missing.push('siteId/SHAREPOINT_SITE_ID');
    if (!SITE_URL) missing.push('siteUrl/SHAREPOINT_SITE_URL');
    if (!hostname) missing.push('hostname/SHAREPOINT_HOSTNAME');
    if (!siteName) missing.push('siteName/SHAREPOINT_SITE_NAME');
    const msg = `Missing SharePoint site configuration. Provide one of: (1) siteId/SHAREPOINT_SITE_ID, or (2) siteUrl/SHAREPOINT_SITE_URL (e.g., https://tenant.sharepoint.com/sites/SiteName), or (3) hostname+siteName / SHAREPOINT_HOSTNAME + SHAREPOINT_SITE_NAME. Missing: ${missing.join(', ')}`;
    const err = new Error(msg);
    err.status = 500;
    throw err;
  }
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(
    hostname
  )}:/sites/${encodeURIComponent(siteName)}?$select=id`;
  return graphFetch(url, { method: "GET" });
}

async function listDrives(ctx = {}) {
  const site = await resolveSiteId(ctx);
  console.log(`[Graph] Fetching drives for site: ${ctx.siteUrl || ctx.siteId || site.id}`);
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(
    site.id
  )}/drives`;
  const data = await graphFetch(url, { method: "GET" });
  const drives = (data.value || []).map((d) => ({ id: d.id, name: d.name }));
  return drives;
}

async function resolveDriveIdByName(driveName, ctx = {}) {
  const key = String(driveName || "").toLowerCase();
  const drives = await listDrives(ctx);
  const match = drives.find((d) => String(d.name).toLowerCase() === key);
  if (!match) return { id: null, available: drives.map((d) => d.name) };
  return { id: match.id, available: drives.map((d) => d.name) };
}

async function listItems(driveId) {
  console.log(`[Graph] Fetching items for drive: ${driveId}`);
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    driveId
  )}/root/children?$select=id,name&$top=999`;
  const data = await graphFetch(url, { method: "GET" });
  return (data.value || []).map((it) => ({ id: it.id, name: it.name }));
}

async function resolveItemIdByName(driveId, itemName) {
  const items = await listItems(driveId);
  const match = items.find((it) => String(it.name).toLowerCase() === String(itemName).toLowerCase());
  if (!match) return { id: null, available: items.map((i) => i.name) };
  return { id: match.id, available: items.map((i) => i.name) };
}

// Public helpers that throw with helpful messages
async function resolveDriveId(driveName, ctx = {}) {
  // Retry once on empty response (still hits Graph directly)
  let drives = await listDrives(ctx);
  if (!drives.length) {
    console.warn(`[WARN] No drives found. Retrying in 1s...`);
    await new Promise((r) => setTimeout(r, 1000));
    drives = await listDrives(ctx);
  }
  console.log(`[Debug] resolveDriveId: looking for "${driveName}". Available: ${drives.map((d) => d.name).join(', ')}`);
  const drive = drives.find((d) => String(d.name).toLowerCase() === String(driveName).toLowerCase());
  if (!drive) {
    const list = JSON.stringify(drives.map((d) => d.name));
    const err = new Error(`Drive not found. Available drives: ${list}`);
    err.status = 404;
    throw err;
  }
  return drive.id;
}

async function resolveItemId(driveId, itemName) {
  const itemRes = await resolveItemIdByName(driveId, itemName);
  if (!itemRes.id) {
    const list = JSON.stringify(itemRes.available || []);
    const err = new Error(`File not found in this drive. Available items: ${list}`);
    err.status = 404;
    throw err;
  }
  return itemRes.id;
}

// Worksheets helpers
async function listWorksheets(driveId, itemId) {
  const base = buildWorkbookBase({ driveId, itemId });
  const url = `${base}/worksheets`;
  const data = await graphFetch(url, { method: "GET" });
  return (data.value || []).map((ws) => ({ id: ws.id, name: ws.name }));
}

async function resolveWorksheetIdByName(driveId, itemId, sheetName) {
  const key = String(sheetName || "").toLowerCase();
  const sheets = await listWorksheets(driveId, itemId);
  const match = sheets.find((ws) => String(ws.name).toLowerCase() === key);
  return match?.id || null;
}

function parseSheetAndAddress(range) {
  const str = String(range || "");
  const idx = str.indexOf("!");
  if (idx > 0) {
    return { sheetName: str.slice(0, idx), address: str.slice(idx + 1) };
  }
  return { sheetName: null, address: str };
}

// POST /excel/read
// Body: { driveName, itemName, range (optionally Sheet!A1:B2) }
app.post("/excel/read", async (req, res) => {
  try {
    const ctx = getSiteContext(req);
    let { driveName, itemName, sheetName, range } = req.body || {};
    if (!driveName || !itemName) {
      return res.status(400).json({ success: false, error: "Missing body. Required: driveName, itemName" });
    }

    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);

    // Support Sheet!A1:B2
    if (range) {
      const parsed = parseSheetAndAddress(range);
      if (parsed.sheetName && !sheetName) sheetName = parsed.sheetName;
      range = parsed.address;
    }

    // Determine sheetName dynamically if missing
    const sheets = await listWorksheets(driveId, itemId);
    const availableSheets = sheets.map((s) => s.name);
    if (!sheetName) {
      if (sheets.length === 1) {
        sheetName = sheets[0].name;
      } else {
        return res.status(400).json({ success: false, error: "Multiple sheets found. Please specify sheetName.", availableSheets });
      }
    } else {
      const exists = sheets.some((s) => String(s.name).toLowerCase() === String(sheetName).toLowerCase());
      if (!exists) {
        return res.status(404).json({ success: false, error: "Sheet not found.", availableSheets });
      }
    }
    console.log(`[Debug] Using sheet: ${sheetName}`);

    const base = buildWorkbookBase({ driveId, itemId });
    if (range && range.trim().length > 0) {
      console.log(`[Debug] Reading range: ${sheetName}!${range}`);
      const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(range)}')`;
      const data = await graphFetch(url, { method: "GET" });
      return res.json({ success: true, data });
    }

    // No range provided: return usedRange values
    const usedUrl = `${base}/worksheets('${encodeURIComponent(sheetName)}')/usedRange`;
    const used = await graphFetch(usedUrl, { method: "GET" });
    return res.json({ success: true, data: { message: "No range provided. Returning full sheet contents.", values: used?.values || [] } });
  } catch (err) {
    const status = err.status || 500;
    return res.status(status).json({ success: false, error: err.message });
  }
});

// ---- Compatibility aliases (preserve legacy /api/* paths) ----
// Health
app.get("/api/health", (req, res) => res.redirect(307, "/health"));

// Discovery
app.get("/api/list-drives", (req, res) => res.redirect(307, "/list-drives"));
app.get("/api/list-items", (req, res) => res.redirect(307, "/list-items"));

// Excel ops
app.post("/api/excel/read", (req, res) => res.redirect(307, "/excel/read"));
app.post("/api/excel/write", (req, res) => res.redirect(307, "/excel/write"));
app.post("/api/excel/delete", (req, res) => res.redirect(307, "/excel/delete"));
app.post("/api/excel/create-sheet", (req, res) => res.redirect(307, "/excel/create-sheet"));
app.post("/api/excel/delete-sheet", (req, res) => res.redirect(307, "/excel/delete-sheet"));

// POST /excel/write
// Body: { driveName, itemName, range (may be Sheet!A1:B2), values (2D array) }
app.post("/excel/write", async (req, res) => {
  try {
    const ctx = getSiteContext(req);
    let { driveName, itemName, sheetName, range, values } = req.body || {};

    if (!driveName || !itemName || !Array.isArray(values)) {
      return res.status(400).json({ success: false, error: "Missing body. Required: driveName, itemName, values(2D array)" });
    }

    // Parse possible Sheet!A1:B2
    if (range) {
      const parsed = parseSheetAndAddress(range);
      if (parsed.sheetName && !sheetName) sheetName = parsed.sheetName;
      range = parsed.address;
    }

    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);

    // Determine sheetName dynamically
    const sheets = await listWorksheets(driveId, itemId);
    const availableSheets = sheets.map((s) => s.name);
    if (!sheetName) {
      if (sheets.length === 1) sheetName = sheets[0].name;
      else return res.status(400).json({ success: false, error: "Multiple sheets found. Please specify sheetName.", availableSheets });
    } else {
      const exists = sheets.some((s) => String(s.name).toLowerCase() === String(sheetName).toLowerCase());
      if (!exists) return res.status(404).json({ success: false, error: "Sheet not found.", availableSheets });
    }
    console.log(`[Debug] Using sheet: ${sheetName}`);

    const base = buildWorkbookBase({ driveId, itemId });

    // Helper to convert number to Excel column letters (1-based)
    const numToCol = (n) => {
      let s = "";
      while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
      return s;
    };

    if (range && range.trim().length > 0) {
      console.log(`[Debug] Writing to explicit range: ${sheetName}!${range}`);
      const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(range)}')`;
      const data = await graphFetch(url, { method: "PATCH", body: JSON.stringify({ values }) });
      return res.json({ success: true, data });
    }

    // No range provided → append after used range
    const usedUrl = `${base}/worksheets('${encodeURIComponent(sheetName)}')/usedRange`;
    const used = await graphFetch(usedUrl, { method: "GET" });
    const rowIndex = Number(used?.rowIndex ?? 0); // 0-based
    const rowCount = Number(used?.rowCount ?? 0);
    const nextRow = rowIndex + rowCount + 1; // Excel addresses are 1-based
    const cols = Array.isArray(values[0]) ? values[0].length : 1;
    const rows = Array.isArray(values) ? values.length : 1;
    const endCol = numToCol(cols);
    const endRow = nextRow + rows - 1;
    const autoRange = `A${nextRow}:${endCol}${endRow}`;
    console.log(`[Debug] Appending data after row ${rowIndex + rowCount}`);
    console.log(`[Debug] Auto range: ${sheetName}!${autoRange}`);
    const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(autoRange)}')`;
    const data = await graphFetch(url, { method: "PATCH", body: JSON.stringify({ values }) });
    return res.json({ success: true, data: { message: `No range provided. Appending data after row ${rowIndex + rowCount}.`, writtenTo: `${autoRange}` } });
  } catch (err) {
    const status = err.status || 500;
    return res.status(status).json({ success: false, error: err.message });
  }
});

// POST /excel/create-sheet
// Body: { driveName, itemName, name }
app.post("/excel/create-sheet", async (req, res) => {
  try {
    const ctx = getSiteContext(req);
    const { driveName, itemName, name } = req.body || {};
    if (!driveName || !itemName || !name) {
      return res.status(400).json({
        success: false,
        error: "Missing body. Required: driveName, itemName, name",
      });
    }

    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);
    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets/add`;
    const data = await graphFetch(url, {
      method: "POST",
      body: JSON.stringify({ name }),
    });
    return res.json({ success: true, data });
  } catch (err) {
    const status = err.status || 500;
    return res.status(status).json({ success: false, error: err.message });
  }
});

// POST /excel/delete
// Clears data in a range
// Body: { driveName, itemName, sheetName, range, applyTo? }
app.post("/excel/delete", async (req, res) => {
  try {
    const ctx = getSiteContext(req);
    let { driveName, itemName, sheetName, range, applyTo = "contents" } = req.body || {};
    if (!driveName || !itemName) {
      return res.status(400).json({ success: false, error: "Missing body. Required: driveName, itemName" });
    }

    // Parse possible Sheet!A1:B2
    if (range) {
      const parsed = parseSheetAndAddress(range);
      if (parsed.sheetName && !sheetName) sheetName = parsed.sheetName;
      range = parsed.address;
    }

    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);
    const sheets = await listWorksheets(driveId, itemId);
    const availableSheets = sheets.map((s) => s.name);
    if (!sheetName) {
      if (sheets.length === 1) sheetName = sheets[0].name;
      else return res.status(400).json({ success: false, error: "Multiple sheets found. Please specify sheetName.", availableSheets });
    } else {
      const exists = sheets.some((s) => String(s.name).toLowerCase() === String(sheetName).toLowerCase());
      if (!exists) return res.status(404).json({ success: false, error: "Sheet not found.", availableSheets });
    }
    console.log(`[Debug] Using sheet: ${sheetName}`);

    const base = buildWorkbookBase({ driveId, itemId });
    if (range && range.trim().length > 0) {
      console.log(`[Debug] Clearing range: ${sheetName}!${range}`);
      const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(range)}')/clear`;
      const data = await graphFetch(url, { method: "POST", body: JSON.stringify({ applyTo }) });
      return res.json({ success: true, data });
    }

    // No range provided
    if (String(applyTo).toLowerCase() === "all") {
      console.log(`[Debug] Clearing entire used range for sheet: ${sheetName}`);
      const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/usedRange/clear`;
      const data = await graphFetch(url, { method: "POST", body: JSON.stringify({ applyTo: "contents" }) });
      return res.json({ success: true, data: { message: "Cleared entire sheet used range." } });
    }
    return res.status(400).json({ success: false, error: "No range provided. To clear entire sheet, call again with applyTo=all." });
  } catch (err) {
    const status = err.status || 500;
    return res.status(status).json({ success: false, error: err.message });
  }
});

// POST /excel/delete-sheet
// Body: { driveName, itemName, sheetName }
app.post("/excel/delete-sheet", async (req, res) => {
  try {
    const ctx = getSiteContext(req);
    const { driveName, itemName, sheetName } = req.body || {};
    if (!driveName || !itemName || !sheetName) {
      return res.status(400).json({ success: false, error: "Missing body. Required: driveName, itemName, sheetName" });
    }
    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);
    const worksheetId = await resolveWorksheetIdByName(driveId, itemId, sheetName);
    if (!worksheetId) {
      // fetch available worksheets for helpful error
      const sheets = await listWorksheets(driveId, itemId);
      const available = JSON.stringify(sheets.map(s => s.name));
      return res.status(404).json({ success: false, error: `Worksheet not found. Available sheets: ${available}` });
    }
    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets/${encodeURIComponent(worksheetId)}`;
    await graphFetch(url, { method: "DELETE" });
    return res.json({ success: true });
  } catch (err) {
    const status = err.status || 500;
    return res.status(status).json({ success: false, error: err.message });
  }
});

// GET /list-drives (supports per-request site overrides)
app.get("/list-drives", async (req, res) => {
  try {
    const ctx = getSiteContext(req);
    // Always resolve site and list drives via Graph
    const site = await resolveSiteId(ctx);
    const drives = await listDrives(ctx);
    console.log(`[Debug] /list-drives siteUrl=${ctx.siteUrl || '(none)'} siteId=${site.id || '(unknown)'} drives.count=${drives.length}`);
    return res.json({ success: true, drives });
  } catch (err) {
    const status = err.status || 500;
    return res.status(status).json({ success: false, error: err.message });
  }
});

// GET /list-items?driveName=Documents (supports per-request site overrides)
app.get("/list-items", async (req, res) => {
  try {
    const ctx = getSiteContext(req);
    const { driveName } = req.query || {};
    if (!driveName) {
      return res.status(400).json({ success: false, error: "driveName is required" });
    }
    const driveId = await resolveDriveId(driveName, ctx);
    const items = await listItems(driveId);
    return res.json({ success: true, items });
  } catch (err) {
    const status = err.status || 500;
    return res.status(status).json({ success: false, error: err.message });
  }
});

// Important for Vercel: export the app as default
import { createServer } from "http";

// Convert Express app into a request handler for Vercel
export default function handler(req, res) {
  app(req, res);
}

// export default app;

