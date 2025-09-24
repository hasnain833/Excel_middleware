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

- Azure AD Client Credentials authentication
- Automatic token management and refresh
- Rate limiting (100 requests per 15 minutes)
- Input validation and sanitization
- Comprehensive audit logging
- CORS protection and security headers

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
