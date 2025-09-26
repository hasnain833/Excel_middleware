const { Client } = require("@microsoft/microsoft-graph-client");
const auditService = require("./auditService");
const logger = require("../config/logger");
const permissions = require("../config/permissions");

class ExcelService {
  constructor() {
    this.auditService = auditService;
    // Configuration for SharePoint site discovery
    this.hostname =
      process.env.SHAREPOINT_HOSTNAME || "yourtenant.sharepoint.com";
    this.siteName = process.env.SHAREPOINT_SITE_NAME || "Documents";
    this.siteUrl = process.env.SHAREPOINT_SITE_URL || "";
    this.siteId = process.env.SHAREPOINT_SITE_ID || "";
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }

  async getSiteId(graphClient) {
    try {
      // Fast path: if explicit Site ID is provided, use it directly
      if (this.siteId) {
        logger.debug("🔍 Using explicit SHAREPOINT_SITE_ID", { siteId: this.siteId });
        return this.siteId;
      }

      // Prefer explicit SHAREPOINT_SITE_URL when provided
      let siteEndpoint;
      if (this.siteUrl) {
        // Normalize if user provided a full https URL
        let normalized = this.siteUrl;
        try {
          if (/^https?:\/\//i.test(this.siteUrl)) {
            const u = new URL(this.siteUrl);
            // Graph expects host:/sites/Path (no scheme)
            const path = u.pathname.replace(/\/+$/, '');
            normalized = `${u.host}:${path}`;
          }
        } catch (_) {
          // keep original if URL parsing fails
        }
        siteEndpoint = `/sites/${normalized}`;
        logger.debug("🔍 Fetching SharePoint site ID via SHAREPOINT_SITE_URL", { raw: this.siteUrl, normalized });
      } else {
        siteEndpoint = `/sites/${this.hostname}:/sites/${this.siteName}`;
        logger.debug("🔍 Fetching SharePoint site ID via hostname/siteName", {
          hostname: this.hostname,
          siteName: this.siteName,
        });
      }

      // Get site by URL or by hostname and site path
      const siteResponse = await graphClient.api(siteEndpoint).get();

      const siteId = siteResponse.id;
      logger.debug("✅ Site ID retrieved successfully", { siteId });

      return siteId;
    } catch (error) {
      logger.error("❌ Failed to get site ID", {
        error: error.message,
        hostname: this.hostname,
        siteName: this.siteName,
        siteUrl: this.siteUrl,
        note: "Set SHAREPOINT_SITE_ID to bypass lookup if your environment blocks /sites resolution or requires special DNS/proxy."
      });

      if (error.code === "itemNotFound") {
        const siteHint = this.siteUrl
          ? this.siteUrl
          : `${this.hostname}/sites/${this.siteName}`;
        throw new Error(`SharePoint site not found: ${siteHint}`);
      } else if (error.code === "Forbidden") {
        throw new Error(
          "Access denied to SharePoint site. Check application permissions."
        );
      } else if (error.code === "Unauthorized") {
        throw new Error("Authentication failed. Check access token.");
      }

      throw new Error(`Failed to retrieve site ID: ${error.message}`);
    }
  }
  async getDrives(graphClient, siteId) {
    try {
      logger.debug("🔍 Fetching drives from site", { siteId });

      const drivesResponse = await graphClient
        .api(`/sites/${siteId}/drives`)
        .get();

      const drives = drivesResponse.value.map((drive) => ({
        id: drive.id,
        name: drive.name,
        description: drive.description,
        driveType: drive.driveType,
        webUrl: drive.webUrl,
      }));

      logger.debug("✅ Drives retrieved successfully", {
        driveCount: drives.length,
        drives: drives.map((d) => ({ id: d.id, name: d.name })),
      });

      return drives;
    } catch (error) {
      logger.error("❌ Failed to get drives", { error: error.message, siteId });

      if (error.code === "Forbidden") {
        throw new Error(
          "Access denied to site drives. Check application permissions."
        );
      }

      throw new Error(`Failed to retrieve drives: ${error.message}`);
    }
  }

  async getWorkbooksFromDrives(graphClient, drives) {
    try {
      logger.debug("🔍 Searching for Excel workbooks in drives", {
        driveCount: drives.length,
      });

      const allWorkbooks = [];

      for (const drive of drives) {
        try {
          logger.debug(`🔍 Searching drive: ${drive.name}`, {
            driveId: drive.id,
          });

          // Search for likely Excel files in this drive. Note: Graph search does not support $filter chaining reliably.
          // We'll search for 'xlsx' and post-filter by file presence and name extension.
          const searchResponse = await graphClient
            .api(`/drives/${drive.id}/root/search(q='xlsx')`)
            .top(200)
            .get();

          const workbooks = (searchResponse.value || [])
            .filter(
              (item) =>
                item.file &&
                typeof item.name === "string" &&
                item.name.toLowerCase().endsWith(".xlsx")
            )
            .map((item) => ({
              id: item.id,
              name: item.name,
              driveId: drive.id,
              driveName: drive.name,
            }));

          allWorkbooks.push(...workbooks);
          logger.debug(
            `✅ Found ${workbooks.length} workbooks in drive: ${drive.name}`
          );
        } catch (driveError) {
          logger.warn(`⚠️ Failed to search drive: ${drive.name}`, {
            driveId: drive.id,
            error: driveError.message,
          });
          // Continue with other drives even if one fails
        }
      }

      logger.debug("✅ Total workbooks found across all drives", {
        totalCount: allWorkbooks.length,
      });
      if (allWorkbooks.length === 0) {
        logger.info(
          "ℹ️ No workbooks found via search. Consider verifying site/drives or using explicit drive scanning.",
          {
            siteUrl: this.siteUrl || `${this.hostname}/sites/${this.siteName}`,
          }
        );
      }
      return allWorkbooks;
    } catch (error) {
      logger.error("❌ Failed to get workbooks from drives", {
        error: error.message,
      });
      throw new Error(`Failed to search for workbooks: ${error.message}`);
    }
  }

  async getWorkbooks(accessToken, auditContext) {
    try {
      const graphClient = this.createGraphClient(accessToken);

      // Step 1: Get SharePoint site ID
      const siteId = await this.getSiteId(graphClient);

      // Step 2: Get all drives in the site
      const drives = await this.getDrives(graphClient, siteId);

      // Step 3: Search for Excel workbooks in all drives
      const workbooks = await this.getWorkbooksFromDrives(graphClient, drives);

      // Step 4: Apply permission filtering
      const filteredWorkbooks = workbooks.filter((workbook) =>
        permissions.canAccessWorkbook(auditContext.user, workbook.id)
      );

      // Log audit event
      auditService.logSystemEvent({
        event: "WORKBOOKS_RETRIEVED",
        details: {
          totalFound: workbooks.length,
          accessibleCount: filteredWorkbooks.length,
          user: auditContext.user,
          siteId,
          driveCount: drives.length,
        },
      });

      logger.info("📊 Workbooks retrieval completed", {
        totalFound: workbooks.length,
        accessibleCount: filteredWorkbooks.length,
        user: auditContext.user,
      });

      return filteredWorkbooks;
    } catch (error) {
      logger.error("❌ Excel service - failed to get workbooks:", error);
      throw error;
    }
  }

  async getWorksheets(accessToken, driveId, itemId, auditContext) {
    try {
      // Check workbook access permission
      if (!permissions.canAccessWorkbook(auditContext.user, itemId)) {
        throw new Error("Access denied to workbook");
      }

      const graphClient = this.createGraphClient(accessToken);

      logger.debug("🔍 Fetching worksheets", { driveId, itemId });

      const response = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
        .get();

      const worksheets = response.value.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        position: sheet.position,
        visibility: sheet.visibility,
      }));

      // Filter worksheets based on permissions
      const filteredWorksheets = worksheets.filter((worksheet) =>
        permissions.canAccessWorksheet(auditContext.user, itemId, worksheet.id)
      );

      logger.debug("✅ Worksheets retrieved successfully", {
        workbookId: itemId,
        totalCount: worksheets.length,
        accessibleCount: filteredWorksheets.length,
      });

      return filteredWorksheets;
    } catch (error) {
      logger.error("❌ Excel service - failed to get worksheets:", error);

      if (error.code === "itemNotFound") {
        throw new Error("Workbook not found or not accessible");
      } else if (error.code === "Forbidden") {
        throw new Error("Access denied to workbook");
      }

      throw error;
    }
  }
  async readRange(params) {
    const { accessToken, driveId, itemId, worksheetId, range, auditContext } =
      params;

    try {
      // Check permissions
      const hasPermission = permissions.canReadRange(
        auditContext.user,
        itemId,
        worksheetId,
        range
      );

      if (!hasPermission.allowed) {
        auditService.logPermissionCheck({
          ...auditContext,
          workbookId: itemId,
          worksheetId: worksheetId,
          range: range,
          requestedPermission: "READ",
          granted: false,
          reason: hasPermission.reason,
        });
        throw new Error(`Read access denied: ${hasPermission.reason}`);
      }

      const graphClient = this.createGraphClient(accessToken);

      logger.debug("🔍 Reading Excel range", {
        driveId,
        itemId,
        worksheetId,
        range,
      });

      const response = await graphClient
        .api(
          `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`
        )
        .get();

      const rangeData = {
        address: response.address,
        values: response.values,
        formulas: response.formulas,
        text: response.text,
        rowCount: response.rowCount,
        columnCount: response.columnCount,
      };

      // Log successful permission check and operation
      auditService.logPermissionCheck({
        ...auditContext,
        workbookId: itemId,
        worksheetId: worksheetId,
        range: range,
        requestedPermission: "READ",
        granted: true,
      });

      auditService.logReadOperation({
        ...auditContext,
        workbookId: itemId,
        worksheetId: worksheetId,
        range: range,
        cellCount: rangeData.rowCount * rangeData.columnCount,
        success: true,
      });

      logger.debug("✅ Range read successfully", {
        range,
        rowCount: rangeData.rowCount,
        columnCount: rangeData.columnCount,
      });

      return rangeData;
    } catch (error) {
      logger.error("❌ Excel service - failed to read range:", error);

      if (error.code === "InvalidArgument") {
        throw new Error(`Invalid range format: ${range}`);
      } else if (error.code === "itemNotFound") {
        throw new Error("Worksheet or range not found");
      }

      throw error;
    }
  }

  async writeRange(params) {
    const {
      accessToken,
      driveId,
      itemId,
      worksheetId,
      range,
      values,
      auditContext,
    } = params;

    try {
      // Validate input data
      if (!Array.isArray(values) || values.length === 0) {
        throw new Error("Values must be a non-empty array");
      }

      // Check permissions
      const hasPermission = permissions.canWriteRange(
        auditContext.user,
        itemId,
        worksheetId,
        range
      );

      if (!hasPermission.allowed) {
        auditService.logPermissionCheck({
          ...auditContext,
          workbookId: itemId,
          worksheetId: worksheetId,
          range: range,
          requestedPermission: "WRITE",
          granted: false,
          reason: hasPermission.reason,
        });
        throw new Error(`Write access denied: ${hasPermission.reason}`);
      }

      const graphClient = this.createGraphClient(accessToken);

      logger.debug("🔍 Writing to Excel range", {
        driveId,
        itemId,
        worksheetId,
        range,
      });

      // Read current values for audit trail
      let oldValues = null;
      try {
        const currentResponse = await graphClient
          .api(
            `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`
          )
          .get();
        oldValues = currentResponse.values;
      } catch (readError) {
        logger.warn(
          "Could not read current values for audit trail:",
          readError.message
        );
      }

      // Write new values
      const response = await graphClient
        .api(
          `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${range}')`
        )
        .patch({ values: values });

      const updatedData = {
        address: response.address,
        values: response.values,
        rowCount: response.rowCount,
        columnCount: response.columnCount,
      };

      // Log successful permission check and operation
      auditService.logPermissionCheck({
        ...auditContext,
        workbookId: itemId,
        worksheetId: worksheetId,
        range: range,
        requestedPermission: "WRITE",
        granted: true,
      });

      auditService.logWriteOperation({
        ...auditContext,
        workbookId: itemId,
        worksheetId: worksheetId,
        range: range,
        oldValues: oldValues,
        newValues: values,
        cellsModified: updatedData.rowCount * updatedData.columnCount,
        success: true,
      });

      logger.debug("✅ Range written successfully", {
        range,
        rowCount: updatedData.rowCount,
        columnCount: updatedData.columnCount,
      });

      return updatedData;
    } catch (error) {
      logger.error("❌ Excel service - failed to write range:", error);

      if (error.code === "InvalidArgument") {
        throw new Error(`Invalid range format or data: ${range}`);
      } else if (error.code === "itemNotFound") {
        throw new Error("Worksheet or range not found");
      }

      throw error;
    }
  }
}

module.exports = new ExcelService();
