# Excel GPT Middleware - Client Setup Guide

**Version:** 1.0  
**Date:** September 2025  
**Document Type:** Client Installation & Configuration Guide

---

## 📋 Table of Contents

1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [Azure AD Application Setup](#azure-ad-application-setup)
4. [Project Installation](#project-installation)
5. [Configuration](#configuration)
6. [Testing & Verification](#testing--verification)
7. [Troubleshooting](#troubleshooting)
8. [Support & Contact](#support--contact)

---

## 🎯 Overview

The Excel GPT Middleware is a secure bridge that connects your GPT applications to Excel files stored in SharePoint/OneDrive. This guide will walk you through the complete setup process in simple, easy-to-follow steps.

### What You'll Achieve
- ✅ Secure connection to your organizational Excel files
- ✅ Read and write access to Excel data via API
- ✅ Integration with GPT applications
- ✅ Enterprise-grade security through Azure AD

### Estimated Setup Time
**30-45 minutes** (including Azure AD configuration)

---

## 🔧 Prerequisites

Before starting, ensure you have:

### Required Access
- [ ] **Azure Admin Access** - Ability to create Azure AD applications
- [ ] **SharePoint/OneDrive Access** - Access to organizational files
- [ ] **Local Development Environment** - Windows/Mac/Linux with Node.js

### Required Software
- [ ] **Node.js** (version 16 or higher) - [Download here](https://nodejs.org/)
- [ ] **Git** (optional) - For version control
- [ ] **Text Editor** - VS Code, Notepad++, or similar

### Required Information
- [ ] **Azure Tenant ID** - Your organization's Azure directory ID
- [ ] **Admin Permissions** - To grant API permissions in Azure AD

---

## 🔐 Azure AD Application Setup

### Step 1: Access Azure Portal

1. Open your web browser and go to: **https://portal.azure.com**
2. Sign in with your **Azure administrator account**
3. Ensure you're in the correct organizational directory

### Step 2: Create New App Registration

1. In the Azure Portal, search for **"Azure Active Directory"**
2. Click on **"App registrations"** in the left menu
3. Click **"+ New registration"** button
4. Fill in the application details:
   - **Name:** `Excel GPT Middleware`
   - **Supported account types:** Select **"Accounts in this organizational directory only"**
   - **Redirect URI:** Leave blank (not needed)
5. Click **"Register"**

### Step 3: Note Important Values

After registration, you'll see the **Overview** page. **Copy and save** these values:

```
Application (client) ID: ________________________________
Directory (tenant) ID:   ________________________________
```

⚠️ **Important:** Keep these values secure - you'll need them later!

### Step 4: Create Client Secret

1. In your app registration, click **"Certificates & secrets"** (left menu)
2. Click **"+ New client secret"**
3. Add description: `Excel GPT Middleware Secret`
4. Set expiration: **24 months** (recommended)
5. Click **"Add"**
6. **Immediately copy the secret Value** (not the Secret ID):

```
Client Secret: ________________________________
```

⚠️ **Critical:** Copy this value immediately - you cannot see it again!

### Step 5: Configure API Permissions

1. Click **"API permissions"** in the left menu
2. Click **"+ Add a permission"**
3. Select **"Microsoft Graph"**
4. Choose **"Application permissions"** (not Delegated)
5. Add these permissions:
   - `Files.Read.All`
   - `Files.ReadWrite.All`
   - `Sites.Read.All`
   - `Sites.ReadWrite.All`

### Step 6: Grant Admin Consent

1. Click **"Grant admin consent for [Your Organization]"**
2. Click **"Yes"** to confirm
3. Verify all permissions show **"Granted for [Your Organization]"** with green checkmarks

✅ **Azure AD Setup Complete!**

---

## 💻 Project Installation

### Step 1: Download the Project

1. **Extract** the Excel GPT Middleware project to your desired location
2. **Navigate** to the project folder:
   ```
   C:\your-path\excel-gpt-middleware\
   ```

### Step 2: Install Dependencies

1. **Open Command Prompt** or **PowerShell** in the project folder
2. **Run the installation command:**
   ```bash
   npm install
   ```
3. **Wait** for installation to complete (may take 2-3 minutes)

✅ **Installation Complete!**

---

## ⚙️ Configuration

### Step 1: Create Environment File

1. **Locate** the file named `.env.example` in the project folder
2. **Copy** this file and **rename** the copy to `.env`
3. **Open** the `.env` file in a text editor

### Step 2: Configure Azure AD Settings

**Replace** the placeholder values with your Azure AD information:

```env
# Azure AD Configuration (Required)
AZURE_TENANT_ID=your-tenant-id-from-step-3
AZURE_CLIENT_ID=your-client-id-from-step-3
AZURE_CLIENT_SECRET=your-client-secret-from-step-4

# Server Configuration
PORT=3000
NODE_ENV=development

# Microsoft Graph API Configuration
GRAPH_API_BASE_URL=https://graph.microsoft.com/v1.0

# Logging Configuration
LOG_LEVEL=info
LOG_DIR=./logs

# Rate Limiting
RATE_LIMIT_WINDOW_MS=900000
RATE_LIMIT_MAX_REQUESTS=100

# CORS Configuration
ALLOWED_ORIGINS=http://localhost:3000,https://your-gpt-domain.com
```

### Step 3: Save Configuration

1. **Save** the `.env` file
2. **Verify** all three Azure AD values are correctly entered
3. **Close** the text editor

✅ **Configuration Complete!**

---

## 🧪 Testing & Verification

### Step 1: Start the Server

1. **Open Command Prompt** in the project folder
2. **Run the start command:**
   ```bash
   npm start
   ```
3. **Look for success messages:**
   ```
   ✅ Azure AD client initialized successfully
   🚀 Excel GPT Middleware server started
   📋 Available endpoints: http://localhost:3000
   ```

### Step 2: Test Basic Functionality

1. **Open your web browser**
2. **Navigate to:** `http://localhost:3000`
3. **You should see:**
   ```json
   {
     "service": "Excel GPT Middleware",
     "status": "running",
     "version": "1.0.0"
   }
   ```

### Step 3: Test Health Check

1. **Navigate to:** `http://localhost:3000/health`
2. **You should see:**
   ```json
   {
     "status": "healthy",
     "service": "excel-gpt-middleware"
   }
   ```

### Step 4: Test Excel Access

1. **Navigate to:** `http://localhost:3000/api/excel/workbooks`
2. **You should see:**
   ```json
   {
     "success": true,
     "data": {
       "workbooks": [],
       "count": 0
     }
   }
   ```

✅ **All Tests Passed!** Your middleware is working correctly.

---

## 🔧 Troubleshooting

### Common Issues & Solutions

#### Issue: "Authentication failed"
**Symptoms:** Error messages about Azure AD authentication
**Solutions:**
1. ✅ Verify your Azure AD credentials in `.env` file
2. ✅ Ensure admin consent was granted for all permissions
3. ✅ Check that the Azure AD app is in the correct tenant

#### Issue: "Port already in use"
**Symptoms:** Error about port 3000 being occupied
**Solutions:**
1. ✅ Stop any other applications using port 3000
2. ✅ Change the PORT value in `.env` file to 3001 or 8080
3. ✅ Restart the server

#### Issue: "Module not found"
**Symptoms:** Error about missing Node.js modules
**Solutions:**
1. ✅ Run `npm install` again
2. ✅ Delete `node_modules` folder and run `npm install`
3. ✅ Ensure Node.js version 16+ is installed

#### Issue: "Permission denied"
**Symptoms:** Cannot access Excel files or SharePoint
**Solutions:**
1. ✅ Verify Azure AD permissions are granted
2. ✅ Ensure you have access to the SharePoint/OneDrive files
3. ✅ Check that the service principal has proper permissions

### Getting Help

If you encounter issues not covered here:

1. **Check the logs** in the `./logs/` folder for detailed error messages
2. **Verify your Azure AD configuration** matches the setup steps
3. **Ensure all prerequisites** are properly installed
4. **Contact support** with specific error messages

---

## 📞 Support & Contact

### Technical Support
- **Email:** [Your Support Email]
- **Phone:** [Your Support Phone]
- **Hours:** Monday-Friday, 9 AM - 5 PM

### Documentation
- **API Documentation:** `docs/API.md`
- **Deployment Guide:** `docs/DEPLOYMENT.md`
- **Setup Instructions:** `docs/SETUP.md`

### Quick Reference

#### Server URLs
- **Main Service:** `http://localhost:3000`
- **Health Check:** `http://localhost:3000/health`
- **API Documentation:** `http://localhost:3000/api/docs`
- **Excel API:** `http://localhost:3000/api/excel/`

#### Important Files
- **Configuration:** `.env`
- **Main Server:** `src/server.js`
- **Logs:** `./logs/` folder

---

## ✅ Setup Checklist

Use this checklist to ensure everything is configured correctly:

### Azure AD Setup
- [ ] Created Azure AD application
- [ ] Noted Application (client) ID
- [ ] Noted Directory (tenant) ID  
- [ ] Created client secret
- [ ] Added Microsoft Graph permissions
- [ ] Granted admin consent

### Project Setup
- [ ] Downloaded/extracted project files
- [ ] Installed Node.js dependencies (`npm install`)
- [ ] Created `.env` file from `.env.example`
- [ ] Configured Azure AD credentials in `.env`
- [ ] Started server (`npm start`)

### Verification
- [ ] Server starts without errors
- [ ] Health check responds correctly
- [ ] Excel API endpoints are accessible
- [ ] Authentication is working

### Final Steps
- [ ] Documented configuration for your team
- [ ] Tested with sample Excel files
- [ ] Ready for GPT integration

---

**🎉 Congratulations! Your Excel GPT Middleware is now ready for use.**

This middleware provides secure, enterprise-grade access to your Excel files through a simple API that can be integrated with any GPT application.

---

*Document Version: 1.0 | Last Updated: September 2025*
