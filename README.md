# Excel GPT Middleware

A secure, enterprise-grade middleware solution that connects custom GPTs to Excel files stored on SharePoint/OneDrive using Microsoft Graph API with Azure AD authentication.

## ✅ Status: Production Ready

This middleware is fully functional and tested with:
- ✅ Azure AD Client Credentials authentication
- ✅ Microsoft Graph API integration with SharePoint site discovery
- ✅ Role-based range protection for secure GPT operations
- ✅ Enterprise audit logging with complete operation tracking
- ✅ Comprehensive error handling and validation
- ✅ Rate limiting and security measures

## 🚀 Enterprise Features

- **🔐 Azure AD Authentication**: Automatic service-to-service authentication
- **📊 Excel Integration**: Full read/write access to Excel ranges and tables
- **🏢 SharePoint Discovery**: Automatic site and drive enumeration
- **🔒 Range Protection**: Configurable allowed/locked ranges for GPT safety
- **📋 Audit Logging**: Complete operation history with pre/post values
- **🛡️ Security**: Built-in rate limiting, validation, and error handling
- **📈 Production Ready**: Monitoring, health checks, and graceful shutdown

## Project Structure

```
excel-gpt-middleware/
├── src/
│   ├── auth/
│   │   ├── azureAuth.js          # Azure AD authentication
│   │   └── middleware.js         # Authentication middleware
│   ├── controllers/
│   │   ├── excelController.js    # Excel operations controller
│   │   └── healthController.js   # Health check endpoints
│   ├── services/
│   │   ├── excelService.js       # Excel API service
│   │   ├── graphService.js       # Microsoft Graph service
│   │   └── auditService.js       # Audit logging service
│   ├── middleware/
│   │   ├── validation.js         # Request validation
│   │   ├── errorHandler.js       # Error handling
│   │   └── rateLimiter.js        # Rate limiting
│   ├── config/
│   │   ├── database.js           # Database configuration
│   │   ├── logger.js             # Winston logger setup
│   │   └── permissions.js        # Role-based permissions
│   ├── routes/
│   │   ├── excel.js              # Excel API routes
│   │   └── health.js             # Health check routes
│   ├── utils/
│   │   ├── helpers.js            # Utility functions
│   │   └── constants.js          # Application constants
│   └── server.js                 # Main server file
├── tests/
│   ├── unit/
│   └── integration/
├── docs/
│   ├── API.md                    # API documentation
│   ├── SETUP.md                  # Setup instructions
│   └── DEPLOYMENT.md             # Deployment guide
├── logs/                         # Log files directory
├── .env.example                  # Environment variables template
├── .gitignore
├── package.json
└── README.md
```

## Quick Start

1. **Configure Azure AD**: Set up your organizational Azure AD app with required permissions
2. **Environment Setup**: Copy `.env.example` to `.env` and add your Azure AD credentials:
   ```bash
   cp .env.example .env
   # Edit .env with your Azure AD values
   ```
3. **Install & Run**:
   ```bash
   npm install
   npm start
   ```
4. **Verify**: Server runs on `http://localhost:3000`

## Environment Files

- **`.env`** - Your actual Azure AD credentials (not in git)
- **`.env.example`** - Template showing required variables

## API Endpoints

### Authentication
All endpoints require authentication via Azure AD Client Credentials Flow.

### Excel Operations
- `GET /api/excel/workbooks` - List accessible workbooks from SharePoint
- `GET /api/excel/worksheets` - Get worksheets in a workbook
- `POST /api/excel/read` - Read data from Excel ranges
- `POST /api/excel/write` - Write data to Excel ranges (🔒 range-protected)
- `POST /api/excel/read-table` - Read Excel table data
- `POST /api/excel/add-table-rows` - Add rows to Excel tables (🔒 range-protected)
- `POST /api/excel/batch` - Perform batch operations

### Monitoring & Audit
- `GET /health` - Health check endpoint
- `GET /api/excel/logs` - Get audit logs with filtering
- `GET /api/docs` - API documentation

## Documentation

- [Client Setup Guide](CLIENT_SETUP_GUIDE.md) - Easy setup instructions for clients
- [API Documentation](docs/API.md) - Complete API reference
- [Deployment Guide](docs/DEPLOYMENT.md) - Production deployment instructions

## Authentication

This middleware uses **Azure AD Client Credentials Flow** for automatic authentication:
- No API keys required
- No JWT tokens needed  
- Service-to-service authentication handled automatically
- Enterprise-grade security through Azure AD

## Security Features

- Azure AD Client Credentials authentication
- Automatic token management and refresh
- Rate limiting (100 requests per 15 minutes)
- Input validation and sanitization
- Comprehensive audit logging
- CORS protection and security headers

## RBAC (Role-Based Access Control)

This project includes a config-driven RBAC system that controls which roles can perform write, create, and delete operations on Excel resources.

### Roles and Permissions

Defined in `config/roles.js`:

```
admin  -> canWrite, canCreate, canDelete
editor -> canWrite
viewer -> (read-only)
```

### Role Extraction

- Development: provide `x-user-role: admin|editor|viewer` header. If omitted, defaults to `viewer`.
- Production: role is extracted from a verified JWT in the `Authorization: Bearer <token>` header. The middleware looks for role hints in `decoded.roles`, `appRole`, or `groups`. If none found, defaults to `viewer`.

JWTs are verified using Azure AD JWKS for the configured tenant (`AZURE_TENANT_ID`). Optionally, a symmetric `JWT_SECRET` can be used for local testing.

Environment variables:

```
AZURE_TENANT_ID=...          # required for JWKS
AZURE_CLIENT_ID=...          # default JWT audience
AZURE_AUDIENCE=...           # optional override audience
JWT_SECRET=...               # optional for local testing
```

### Protected Routes

The following routes enforce RBAC:

- POST `/excel/write`          -> requires `canWrite`
- POST `/excel/create-file`    -> requires `canCreate`
- POST `/excel/delete-file`    -> requires `canDelete`
- POST `/excel/create-sheet`   -> requires `canCreate`
- POST `/excel/delete`         -> requires `canDelete`
- POST `/excel/delete-sheet`   -> requires `canDelete`

Unauthorized or forbidden requests return structured errors:

- 401 Unauthorized: missing/invalid token in production
- 403 Forbidden: role lacks permissions for the requested action

### Testing

- Unit tests: `npm test` includes RBAC tests under `tests/rbac.unit.test.js`.
- Local dev: send `x-user-role` header to simulate roles.

## Integration Example

```javascript
// Simple API call - no authentication headers needed
const response = await fetch('http://localhost:3000/api/excel/workbooks');
const workbooks = await response.json();

// Read Excel data
const data = await fetch('http://localhost:3000/api/excel/read', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    driveId: "your-drive-id",
    itemId: "your-workbook-id",
    worksheetId: "Sheet1", 
    range: "A1:C10"
  })
});
```

## License

MIT License - see LICENSE file for details
