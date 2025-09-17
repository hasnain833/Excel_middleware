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

// Simple in-memory caches with 10-minute TTL
const NAME_CACHE_TTL_MS = 10 * 60 * 1000;
const driveCache = new Map(); // key: driveNameLower -> { id, ts }
const itemCache = new Map();  // key: `${driveId}:${itemNameLower}` -> { id, ts }

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
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(
    site.id
  )}/drives`;
  const data = await graphFetch(url, { method: "GET" });
  const drives = (data.value || []).map((d) => ({ id: d.id, name: d.name }));
  return drives;
}

async function resolveDriveIdByName(driveName, ctx = {}) {
  const key = String(driveName || "").toLowerCase();
  const cached = driveCache.get(key);
  if (cached && Date.now() - cached.ts < NAME_CACHE_TTL_MS) return cached.id;

  const drives = await listDrives(ctx);
  const match = drives.find((d) => String(d.name).toLowerCase() === key);
  if (!match) return { id: null, available: drives.map((d) => d.name) };
  driveCache.set(key, { id: match.id, ts: Date.now() });
  return { id: match.id, available: drives.map((d) => d.name) };
}

async function listItems(driveId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    driveId
  )}/root/children?$select=id,name&$top=999`;
  const data = await graphFetch(url, { method: "GET" });
  return (data.value || []).map((it) => ({ id: it.id, name: it.name }));
}

async function resolveItemIdByName(driveId, itemName) {
  const key = `${driveId}:${String(itemName || "").toLowerCase()}`;
  const cached = itemCache.get(key);
  if (cached && Date.now() - cached.ts < NAME_CACHE_TTL_MS) return cached.id;

  const items = await listItems(driveId);
  const match = items.find((it) => String(it.name).toLowerCase() === String(itemName).toLowerCase());
  if (!match) return { id: null, available: items.map((i) => i.name) };
  itemCache.set(key, { id: match.id, ts: Date.now() });
  return { id: match.id, available: items.map((i) => i.name) };
}

// Public helpers that throw with helpful messages
async function resolveDriveId(driveName, ctx = {}) {
  const driveRes = await resolveDriveIdByName(driveName, ctx);
  if (!driveRes.id) {
    const list = JSON.stringify(driveRes.available || []);
    const err = new Error(`Drive not found. Available drives: ${list}`);
    err.status = 404;
    throw err;
  }
  return driveRes.id;
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
    if (!driveName || !itemName || !range) {
      return res.status(400).json({
        success: false,
        error: "Missing body. Required: driveName, itemName, range",
      });
    }

    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);

    // Support Sheet!A1:B2
    const parsed = parseSheetAndAddress(range);
    if (parsed.sheetName && !sheetName) sheetName = parsed.sheetName;
    const address = parsed.address;
    if (!sheetName) {
      return res.status(400).json({ success: false, error: "sheetName is required (or prefix range as Sheet!A1:B2)" });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(
      sheetName
    )}')/range(address='${encodeURIComponent(address)}')`;
    const data = await graphFetch(url, { method: "GET" });
    return res.json({ success: true, data });
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
    const parsed = parseSheetAndAddress(range);
    if (parsed.sheetName && !sheetName) sheetName = parsed.sheetName;
    const address = parsed.address;

    if (!driveName || !itemName || !sheetName || !address || !Array.isArray(values)) {
      return res.status(400).json({
        success: false,
        error:
          "Missing body. Required: driveName, itemName, sheetName (or prefix range), range, values(2D array)",
      });
    }

    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(
      sheetName
    )}')/range(address='${encodeURIComponent(address)}')`;
    const data = await graphFetch(url, {
      method: "PATCH",
      body: JSON.stringify({ values }),
    });
    return res.json({ success: true, data });
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
    if (!driveName || !itemName || !sheetName || !range) {
      return res.status(400).json({
        success: false,
        error: "Missing body. Required: driveName, itemName, sheetName, range",
      });
    }

    const driveId = await resolveDriveId(driveName, ctx);
    const itemId = await resolveItemId(driveId, itemName);
    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(
      sheetName
    )}')/range(address='${encodeURIComponent(range)}')/clear`;
    const data = await graphFetch(url, {
      method: "POST",
      body: JSON.stringify({ applyTo }),
    });
    return res.json({ success: true, data });
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
    const drives = await listDrives(ctx);
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
export default app;

