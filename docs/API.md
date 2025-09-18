# Excel GPT Middleware API (Minimal)

Base URL: `http://localhost:3000`

The middleware handles Microsoft Graph OAuth2 client_credentials automatically. No token needs to be provided by the client.

All responses have the form: `{ success: true, data }` or `{ success: false, error }`.

## Endpoints

### GET `/health`
Quick health check.

Example response:
```json
{ "success": true, "data": { "status": "ok", "time": "2024-01-01T00:00:00.000Z" } }
```

### POST `/excel/read`
Read values from a range.

Request body:
```json
{ "driveId": "<driveId>", "itemId": "<itemId>", "sheetName": "Sheet1", "range": "A1:B2" }
```

### POST `/excel/write`
Write values to a range.

Request body:
```json
{ "driveId": "<driveId>", "itemId": "<itemId>", "sheetName": "Sheet1", "range": "A1:B2", "values": [[1,2],[3,4]] }
```

### POST `/excel/add-sheet`
Add a new worksheet to the workbook.

Request body:
```json
{ "driveId": "<driveId>", "itemId": "<itemId>", "sheetName": "NewSheet" }
```

### POST `/excel/delete`
Clear data in a range.

Request body:
```json
{ "driveId": "<driveId>", "itemId": "<itemId>", "sheetName": "Sheet1", "range": "A1:B10", "applyTo": "contents" }
```

## Environment Variables

Define any of the following in `.env` (AZURE_* names preferred):

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-app-client-id
AZURE_CLIENT_SECRET=your-app-client-secret

# Optional legacy aliases also supported by the middleware
TENANT_ID=your-tenant-id
CLIENT_ID=your-app-client-id
CLIENT_SECRET=your-app-client-secret
```

## Notes

- Ensure your Azure App Registration has Microsoft Graph application permissions, e.g. `Files.ReadWrite.All`, with admin consent.
- Replace `<driveId>` and `<itemId>` with your OneDrive/SharePoint values.
