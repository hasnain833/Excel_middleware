# Excel GPT Middleware - API Testing Guide

This guide shows how to test the current endpoints exposed by the middleware.

Base URL: `http://localhost:3000`

## Endpoints

- `GET /health` — Basic health check
- `POST /excel/read` — Read values from a range
- `POST /excel/write` — Write values to a range
- `POST /excel/add-sheet` — Add a new worksheet
- `POST /excel/delete` — Clear data in a range

All endpoints return JSON in the shape `{ success: true, data }` or `{ success: false, error }`.

## Environment Variables

Set these in `.env` before testing:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-app-client-id
CLIENT_SECRET=your-app-client-secret
```

## Quick Tests (PowerShell)

```powershell
# Health
Invoke-RestMethod -Uri "http://localhost:3000/health" -Method GET

# Read range
$readBody = @{ driveId="<driveId>"; itemId="<itemId>"; sheetName="Sheet1"; range="A1:B2" } | ConvertTo-Json
Invoke-RestMethod -Uri "http://localhost:3000/excel/read" -Method POST -Body $readBody -ContentType "application/json"

# Write range (2x2)
$writeBody = @{ driveId="<driveId>"; itemId="<itemId>"; sheetName="Sheet1"; range="A1:B2"; values=@(@(1,2),@(3,4)) } | ConvertTo-Json -Depth 5
Invoke-RestMethod -Uri "http://localhost:3000/excel/write" -Method POST -Body $writeBody -ContentType "application/json"

# Add sheet
$addSheetBody = @{ driveId="<driveId>"; itemId="<itemId>"; sheetName="NewSheet" } | ConvertTo-Json
Invoke-RestMethod -Uri "http://localhost:3000/excel/add-sheet" -Method POST -Body $addSheetBody -ContentType "application/json"

# Clear range
$deleteBody = @{ driveId="<driveId>"; itemId="<itemId>"; sheetName="Sheet1"; range="A1:B10"; applyTo="contents" } | ConvertTo-Json
Invoke-RestMethod -Uri "http://localhost:3000/excel/delete" -Method POST -Body $deleteBody -ContentType "application/json"
```

## Quick Tests (curl)

```bash
# Health
curl http://localhost:3000/health

# Read range
curl -X POST http://localhost:3000/excel/read \
  -H "Content-Type: application/json" \
  -d '{"driveId":"<driveId>","itemId":"<itemId>","sheetName":"Sheet1","range":"A1:B2"}'

# Write range
curl -X POST http://localhost:3000/excel/write \
  -H "Content-Type: application/json" \
  -d '{"driveId":"<driveId>","itemId":"<itemId>","sheetName":"Sheet1","range":"A1:B2","values":[[1,2],[3,4]]}'

# Add sheet
curl -X POST http://localhost:3000/excel/add-sheet \
  -H "Content-Type: application/json" \
  -d '{"driveId":"<driveId>","itemId":"<itemId>","sheetName":"NewSheet"}'

# Clear range
curl -X POST http://localhost:3000/excel/delete \
  -H "Content-Type: application/json" \
  -d '{"driveId":"<driveId>","itemId":"<itemId>","sheetName":"Sheet1","range":"A1:B10","applyTo":"contents"}'
```

## Notes

- Replace `<driveId>` and `<itemId>` with your real values from OneDrive/SharePoint.
- Ensure your Azure app registration has Graph application permissions (e.g., `Files.ReadWrite.All`) and admin consent.
