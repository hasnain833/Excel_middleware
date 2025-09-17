# Excel GPT Middleware - Deployment Checklist

## ✅ Pre-Deployment Setup

### 1. Azure AD Configuration
- [ ] Create Azure AD App Registration
- [ ] Configure API permissions for Microsoft Graph
- [ ] Generate client secret
- [ ] Note down: Tenant ID, Client ID, Client Secret

### 2. SharePoint Configuration
- [ ] Identify SharePoint site URL
- [ ] Ensure app has access to document libraries
- [ ] Test with sample Excel files

### 3. Environment Configuration
- [ ] Copy `.env.example` to `.env`
- [ ] Fill in Azure AD credentials
- [ ] Set SharePoint hostname and site name
- [ ] Configure logging and rate limiting

### 4. Range Permissions Setup
- [ ] Edit `rangePermissions.json`
- [ ] Define allowed ranges for GPT operations
- [ ] Set locked ranges for protection
- [ ] Test range validation

## 🚀 Deployment Steps

### 1. Install Dependencies
```bash
npm install
```

### 2. Start Server
```bash
npm start
# or for development
npm run dev
```

### 3. Verify Health
```bash
curl http://localhost:3000/health
```

### 4. Test Authentication
```bash
curl -H "x-user-id: test@company.com" http://localhost:3000/api/excel/workbooks
```

## 📋 Post-Deployment Verification

### Core Functionality
- [ ] Health endpoint responds
- [ ] Authentication works
- [ ] Workbooks endpoint returns data
- [ ] Range validation blocks unauthorized writes
- [ ] Audit logging creates entries

### Security Features
- [ ] Range protection active
- [ ] Audit logs being written
- [ ] Rate limiting functional
- [ ] Error handling working

### Monitoring
- [ ] Logs directory created
- [ ] Audit log file created
- [ ] Winston logging operational

## 🔧 Production Considerations

### Performance
- [ ] Configure appropriate rate limits
- [ ] Set up log rotation
- [ ] Monitor memory usage

### Security
- [ ] Review allowed ranges
- [ ] Set up log monitoring
- [ ] Configure CORS properly

### Maintenance
- [ ] Set up automated backups
- [ ] Configure log retention
- [ ] Plan for credential rotation

## 📞 Support

For issues or questions:
1. Check logs in `./logs/` directory
2. Review audit entries in `audit-log.json`
3. Verify environment configuration
4. Test with Postman collection
