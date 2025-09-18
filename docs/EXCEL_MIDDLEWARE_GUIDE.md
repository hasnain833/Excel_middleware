# Excel GPT Middleware Guide

This guide explains how to use the Excel GPT Middleware to interact with SharePoint/OneDrive Excel files via the Microsoft Graph API. It documents the available endpoints, the auto-selection behavior for drives and sheets, and provides end-to-end examples for common tasks.

## Overview

- Bridges your SharePoint/OneDrive environment with Microsoft Graph to perform Excel operations from a simple HTTP API.
- Supports Azure AD Client Credentials flow and works with both site-bound SharePoint document libraries and personal/OneDrive drives (via site context).
- Adds smart auto-selection logic to minimize required inputs while still providing helpful prompts when multiple options exist.

## Available Endpoints

- GET `/health`
- GET `/list-drives`
- GET `/list-items`
- POST `/excel/read`
- POST `/excel/write`
- POST `/excel/delete`

Compatibility aliases are also available under `/api/*` (e.g., `/api/excel/read` => `/excel/read`).

## Step-by-Step Usage Flow

### 1) List Drives

- Purpose: Discover drives (e.g., SharePoint document libraries like "Documents", "Shared Documents").
- Request:
  - Method: GET
  - URL: `/list-drives`
  - Optional per-request site context via query parameters: `siteId`, `siteUrl`, `hostname`, `siteName`.
- Example:
```http
GET /list-drives
```
- Sample Response:
```json
{
  "success": true,
  "drives": [
    { "id": "b!XYZ...", "name": "Documents" },
    { "id": "b!ABC...", "name": "Shared Documents" }
  ]
}
```

### 2) List Items

- Purpose: List files/folders at the root of a drive.
- Request:
  - Method: GET
  - URL: `/list-items?driveName=Documents`
  - If `driveName` is omitted:
    - If only one drive exists: it will be auto-selected.
    - If multiple drives exist: returns 400 with a helpful error and the list of `availableDrives`.
- Examples:
  - With explicit drive:
```http
GET /list-items?driveName=Documents
```
  - Without drive (auto-select if only one):
```http
GET /list-items
```
- Sample Responses:
  - Success:
```json
{
  "success": true,
  "items": [
    { "id": "01ABCDEF...", "name": "Sales.xlsx" },
    { "id": "01ABCDEFG...", "name": "Archive" }
  ]
}
```
  - Multiple drives exist:
```json
{
  "success": false,
  "error": "Multiple drives found. Please specify driveName.",
  "availableDrives": ["Documents", "Shared Documents"]
}
```

### 3) Read Data

- Purpose: Read a specific range or the entire used range from a worksheet.
- Request:
  - Method: POST
  - URL: `/excel/read`
  - Body fields:
    - `itemName` (required)
    - `driveName` (optional if only one drive)
    - `sheetName` (optional if only one sheet)
    - `range` (optional; can be address only like `A1:B10` or include sheet `Sheet1!A1:B10`)
  - Sheet behavior:
    - If `sheetName` is omitted and only one sheet exists → auto-selected.
    - If multiple sheets exist → returns 400 with `availableSheets`.
- Examples:
  - Explicit sheet and range:
```json
POST /excel/read
{
  "driveName": "Documents",
  "itemName": "Sales.xlsx",
  "sheetName": "Data",
  "range": "A1:C10"
}
```
  - Let middleware auto-select sheet (only one) and read entire used range:
```json
POST /excel/read
{
  "itemName": "Sales.xlsx"
}
```
- Sample Responses:
  - Range provided:
```json
{
  "success": true,
  "data": { "values": [["Region", "Qty"], ["West", 42]] }
}
```
  - No range (full used range values):
```json
{
  "success": true,
  "data": {
    "message": "No range provided. Returning full sheet contents.",
    "values": [["Region", "Qty"], ["West", 42]]
  }
}
```
  - Multiple sheets exist:
```json
{
  "success": false,
  "error": "Multiple sheets found. Please specify sheetName.",
  "availableSheets": ["Data", "Summary"]
}
```

### 4) Write Data

- Purpose: Write values to a specific range, or append after the last used row when no range is provided.
- Request:
  - Method: POST
  - URL: `/excel/write`
  - Body fields:
    - `itemName` (required)
    - `values` (required; 2D array like `[["Header1","Header2"],["Row1","Row2"]]`)
    - `driveName` (optional if only one drive)
    - `sheetName` (optional if only one sheet)
    - `range` (optional; `Sheet!A1:B2` or `A1:B2`)
  - Behavior:
    - With `range`: writes exactly to that address.
    - Without `range`: appends after the used range starting at column A. The middleware computes the destination and logs `Auto range`.
  - Validation:
    - If `itemName` is missing → 400 with: `Missing itemName (the workbook filename is required).`
    - If `values` is missing or not an array → 400 with: `Missing values. Must be a 2D array, e.g. [["Header1","Header2"],["Row1","Row2"]]`
    - `driveName` is not required; the middleware will auto-select a single available drive or return a 400 with `availableDrives` if multiple exist.
- Examples:
  - Explicit range:
```json
POST /excel/write
{
  "driveName": "Documents",
  "itemName": "Sales.xlsx",
  "sheetName": "Data",
  "range": "B2:C3",
  "values": [["West", 42], ["South", 13]]
}
```
  - Append mode (no range):
```json
POST /excel/write
{
  "itemName": "Sales.xlsx",
  "sheetName": "Data",
  "values": [["North", 10]]
}
```
- Sample Responses:
```json
{
  "success": true,
  "data": {
    "message": "No range provided. Appending data after row 12.",
    "writtenTo": "A13:B13"
  }
}
```

- Error Examples:
  - Missing itemName
```json
{
  "success": false,
  "error": "Missing itemName (the workbook filename is required)."
}
```
  - Missing/invalid values (not a 2D array)
```json
{
  "success": false,
  "error": "Missing values. Must be a 2D array, e.g. [[\"Header1\",\"Header2\"],[\"Row1\",\"Row2\"]]"
}
```
  - Multiple drives exist and driveName not specified
```json
{
  "success": false,
  "error": "Multiple drives found. Please specify driveName.",
  "availableDrives": ["Documents", "Shared Documents"]
}
```

### 5) Delete Data

- Purpose: Clear a range or clear the entire used range of a sheet.
- Request:
  - Method: POST
  - URL: `/excel/delete`
  - Body fields:
    - `itemName` (required)
    - `driveName` (optional if only one drive)
    - `sheetName` (optional if only one sheet)
    - `range` (optional; `Sheet!A1:B10` or `A1:B10`)
    - `applyTo` (optional; `contents` by default; use `all` to clear entire used range when no range provided)
  - Behavior:
    - With `range`: clears that range (default `applyTo` = `contents`).
    - Without `range` + `applyTo=all`: clears the sheet's entire used range.
    - Without `range` and without `applyTo=all`: returns 400 with a helpful message.
- Examples:
  - Clear a specific range:
```json
POST /excel/delete
{
  "itemName": "Sales.xlsx",
  "sheetName": "Data",
  "range": "B2:C10"
}
```
  - Clear entire sheet used range:
```json
POST /excel/delete
{
  "itemName": "Sales.xlsx",
  "sheetName": "Data",
  "applyTo": "all"
}
```
- Sample Responses:
```json
{
  "success": true,
  "data": { "message": "Cleared entire sheet used range." }
}
```

## Auto-Selection Logic

- **Drive auto-selection**
  - If `driveName` is omitted:
    - If only one drive exists, middleware auto-selects it and proceeds. Debug log: `Auto-selected drive: <name> (<id>)`.
    - If multiple drives exist, returns 400 with:
```json
{
  "success": false,
  "error": "Multiple drives found. Please specify driveName.",
  "availableDrives": ["Documents", "Shared Documents"]
}
```
- **Sheet auto-selection**
  - If `sheetName` is omitted:
    - If only one sheet exists, it is auto-selected. Debug log: `Using sheet: <sheetName>`.
    - If multiple sheets exist, returns 400 with:
```json
{
  "success": false,
  "error": "Multiple sheets found. Please specify sheetName.",
  "availableSheets": ["Data", "Summary"]
}
```
- **Range parsing**
  - You can send `range` as `Sheet!A1:B2`. The middleware will extract `sheetName` if omitted and set `range` to the address only.

## Site Context (Per Request)

You can override the SharePoint site per request via either body or query parameters. Supported fields:

- `siteId`
- `siteUrl` (e.g., `https://tenant.sharepoint.com/sites/MySite`)
- `hostname` (aka `sharepointHostname`)
- `siteName` (aka `sharepointSiteName`)

If none are provided, the middleware falls back to environment variables in this order: `SHAREPOINT_SITE_ID`, `SHAREPOINT_SITE_URL`, or `SHAREPOINT_HOSTNAME` + `SHAREPOINT_SITE_NAME`.

## Examples Summary

- **List Drives**
```http
GET /list-drives
```
- **List Items (auto-select drive if single)**
```http
GET /list-items
```
- **Read (auto-select drive/sheet, full used range)**
```json
POST /excel/read
{
  "itemName": "Sales.xlsx"
}
```
- **Write (append)**
```json
POST /excel/write
{
  "itemName": "Sales.xlsx",
  "sheetName": "Data",
  "values": [["North", 10]]
}
```
- **Delete (clear full sheet)**
```json
POST /excel/delete
{
  "itemName": "Sales.xlsx",
  "sheetName": "Data",
  "applyTo": "all"
}
```

## Notes

- All endpoints maintain verbose debug logs in `api/index.js` to aid troubleshooting:
  - Drive selection, sheet determination, range used for read/write/delete.
- Authentication uses Azure AD Client Credentials. Ensure your app registration has Graph permissions for Files and Sites as required.
- For production, review the existing security middleware, rate limiting, and logging in the broader codebase.
