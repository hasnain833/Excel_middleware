const { Client } = require("@microsoft/microsoft-graph-client");
const logger = require("../config/logger");
const resolverService = require("./resolverService");
const auditService = require("./auditService");
const { AppError } = require("../middleware/errorHandler");

class RenameService {
  constructor() {
    // Cache for rename suggestions to avoid duplicate API calls
    this.suggestionCache = new Map();
    this.ttlMs = 5 * 60 * 1000; // 5 minutes TTL for suggestions
  }

  createGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }
  async renameFile(
    accessToken,
    driveId,
    itemId,
    oldName,
    newName,
    auditContext
  ) {
    if (!driveId || !itemId || !newName) {
      throw new AppError("driveId, itemId, and newName are required", 400);
    }

    try {
      const graphClient = this.createGraphClient(accessToken);

      // Validate file exists and get current info
      const currentFile = await graphClient
        .api(`/drives/${driveId}/items/${itemId}`)
        .get();

      if (!currentFile) {
        throw new AppError("File not found", 404);
      }

      // Perform rename operation
      const updatedFile = await graphClient
        .api(`/drives/${driveId}/items/${itemId}`)
        .patch({ name: newName });

      // Log the rename operation
      auditService.logSystemEvent({
        event: "FILE_RENAMED",
        details: {
          driveId,
          itemId,
          oldName: currentFile.name,
          newName: updatedFile.name,
          path: currentFile.parentReference?.path || "/",
          requestedBy: auditContext.user,
          timestamp: new Date().toISOString(),
        },
      });

      logger.info(`File renamed successfully`, {
        driveId,
        itemId,
        oldName: currentFile.name,
        newName: updatedFile.name,
      });

      return {
        id: updatedFile.id,
        oldName: currentFile.name,
        newName: updatedFile.name,
        path: updatedFile.parentReference?.path || "/",
        lastModified: updatedFile.lastModifiedDateTime,
      };
    } catch (err) {
      logger.error("Failed to rename file", {
        driveId,
        itemId,
        oldName,
        newName,
        error: err.message,
      });

      if (err.code === "nameAlreadyExists") {
        throw new AppError(
          `A file named '${newName}' already exists in this location`,
          409
        );
      }
      if (err.code === "accessDenied") {
        throw new AppError(
          "Access denied. You may not have permission to rename this file",
          403
        );
      }
      if (err.code === "resourceLocked") {
        throw new AppError(
          "File is currently locked and cannot be renamed",
          423
        );
      }

      throw err;
    }
  }

  async renameFolder(
    accessToken,
    driveId,
    folderId,
    oldName,
    newName,
    auditContext
  ) {
    if (!driveId || !folderId || !newName) {
      throw new AppError("driveId, folderId, and newName are required", 400);
    }

    try {
      const graphClient = this.createGraphClient(accessToken);

      // Validate folder exists
      const currentFolder = await graphClient
        .api(`/drives/${driveId}/items/${folderId}`)
        .get();

      if (!currentFolder || !currentFolder.folder) {
        throw new AppError("Folder not found", 404);
      }

      // Perform rename operation
      const updatedFolder = await graphClient
        .api(`/drives/${driveId}/items/${folderId}`)
        .patch({ name: newName });

      // Log the rename operation
      auditService.logSystemEvent({
        event: "FOLDER_RENAMED",
        details: {
          driveId,
          folderId,
          oldName: currentFolder.name,
          newName: updatedFolder.name,
          path: currentFolder.parentReference?.path || "/",
          requestedBy: auditContext.user,
          timestamp: new Date().toISOString(),
        },
      });

      logger.info(`Folder renamed successfully`, {
        driveId,
        folderId,
        oldName: currentFolder.name,
        newName: updatedFolder.name,
      });

      return {
        id: updatedFolder.id,
        oldName: currentFolder.name,
        newName: updatedFolder.name,
        path: updatedFolder.parentReference?.path || "/",
        lastModified: updatedFolder.lastModifiedDateTime,
      };
    } catch (err) {
      logger.error("Failed to rename folder", {
        driveId,
        folderId,
        oldName,
        newName,
        error: err.message,
      });

      if (err.code === "nameAlreadyExists") {
        throw new AppError(
          `A folder named '${newName}' already exists in this location`,
          409
        );
      }
      if (err.code === "accessDenied") {
        throw new AppError(
          "Access denied. You may not have permission to rename this folder",
          403
        );
      }

      throw err;
    }
  }

  async renameSheet(
    accessToken,
    driveId,
    itemId,
    oldSheetName,
    newSheetName,
    auditContext
  ) {
    if (!driveId || !itemId || !oldSheetName || !newSheetName) {
      throw new AppError(
        "driveId, itemId, oldSheetName, and newSheetName are required",
        400
      );
    }

    try {
      const graphClient = this.createGraphClient(accessToken);

      // Get file info for logging
      const fileInfo = await graphClient
        .api(`/drives/${driveId}/items/${itemId}`)
        .get();

      // Get current worksheet info
      const worksheets = await graphClient
        .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
        .get();

      const targetSheet = worksheets.value.find(
        (ws) => ws.name === oldSheetName
      );
      if (!targetSheet) {
        throw new AppError(`Worksheet '${oldSheetName}' not found`, 404);
      }

      // Perform rename operation
      const updatedSheet = await graphClient
        .api(
          `/drives/${driveId}/items/${itemId}/workbook/worksheets/${targetSheet.id}`
        )
        .patch({ name: newSheetName });

      // Log the rename operation
      auditService.logSystemEvent({
        event: "WORKSHEET_RENAMED",
        details: {
          driveId,
          itemId,
          fileName: fileInfo.name,
          worksheetId: targetSheet.id,
          oldSheetName,
          newSheetName: updatedSheet.name,
          requestedBy: auditContext.user,
          timestamp: new Date().toISOString(),
        },
      });

      logger.info(`Worksheet renamed successfully`, {
        driveId,
        itemId,
        fileName: fileInfo.name,
        oldSheetName,
        newSheetName: updatedSheet.name,
      });

      return {
        worksheetId: updatedSheet.id,
        oldSheetName,
        newSheetName: updatedSheet.name,
        fileName: fileInfo.name,
        fileId: itemId,
      };
    } catch (err) {
      logger.error("Failed to rename worksheet", {
        driveId,
        itemId,
        oldSheetName,
        newSheetName,
        error: err.message,
      });

      if (err.code === "nameAlreadyExists") {
        throw new AppError(
          `A worksheet named '${newSheetName}' already exists in this workbook`,
          409
        );
      }
      if (err.code === "accessDenied") {
        throw new AppError(
          "Access denied. You may not have permission to rename worksheets in this file",
          403
        );
      }

      throw err;
    }
  }

  async findRelatedItems(accessToken, driveId, oldTerm, newTerm) {
    const cacheKey = `${driveId}:${oldTerm}`;
    const cached = this.suggestionCache.get(cacheKey);

    if (cached && Date.now() - cached.ts < this.ttlMs) {
      return cached.items;
    }

    try {
      const graphClient = this.createGraphClient(accessToken);
      const relatedItems = [];

      // Search for files and folders containing the old term
      const searchResults = await this.recursiveSearch(
        graphClient,
        driveId,
        oldTerm
      );

      for (const item of searchResults) {
        if (item.name.toLowerCase().includes(oldTerm.toLowerCase())) {
          const suggestedName = item.name.replace(
            new RegExp(oldTerm, "gi"),
            newTerm
          );

          relatedItems.push({
            id: item.id,
            currentName: item.name,
            suggestedName,
            path: item.path,
            type: item.folder ? "folder" : "file",
            parentId: item.parentId,
          });
        }
      }

      // Cache the results
      this.suggestionCache.set(cacheKey, {
        items: relatedItems,
        ts: Date.now(),
      });

      return relatedItems;
    } catch (err) {
      logger.error("Failed to find related items", {
        driveId,
        oldTerm,
        newTerm,
        error: err.message,
      });
      return [];
    }
  }

  async recursiveSearch(
    graphClient,
    driveId,
    searchTerm,
    folderId = "root",
    currentPath = "",
    depth = 0,
    maxDepth = 10
  ) {
    if (depth > maxDepth) return [];

    const items = [];

    try {
      const resp = await graphClient
        .api(`/drives/${driveId}/items/${folderId}/children`)
        .select("id,name,folder,parentReference")
        .top(999)
        .get();

      for (const item of resp.value || []) {
        const itemPath = currentPath
          ? `${currentPath}/${item.name}`
          : `/${item.name}`;

        if (item.name.toLowerCase().includes(searchTerm.toLowerCase())) {
          items.push({
            id: item.id,
            name: item.name,
            path: itemPath,
            folder: item.folder,
            parentId: folderId,
          });
        }

        // Recursively search subfolders
        if (item.folder) {
          const subItems = await this.recursiveSearch(
            graphClient,
            driveId,
            searchTerm,
            item.id,
            itemPath,
            depth + 1,
            maxDepth
          );
          items.push(...subItems);
        }
      }
    } catch (err) {
      logger.warn(`Failed to search in folder ${folderId}`, {
        error: err.message,
      });
    }

    return items;
  }

  async batchRename(accessToken, driveId, renameOperations, auditContext) {
    const results = [];
    const errors = [];

    for (let i = 0; i < renameOperations.length; i++) {
      const operation = renameOperations[i];

      try {
        let result;

        switch (operation.type) {
          case "file":
            result = await this.renameFile(
              accessToken,
              driveId,
              operation.itemId,
              operation.oldName,
              operation.newName,
              auditContext
            );
            break;

          case "folder":
            result = await this.renameFolder(
              accessToken,
              driveId,
              operation.itemId,
              operation.oldName,
              operation.newName,
              auditContext
            );
            break;

          case "sheet":
            result = await this.renameSheet(
              accessToken,
              driveId,
              operation.fileId,
              operation.oldName,
              operation.newName,
              auditContext
            );
            break;

          default:
            throw new Error(`Unknown operation type: ${operation.type}`);
        }

        results.push({
          index: i,
          operation: operation.type,
          success: true,
          data: result,
        });
      } catch (error) {
        logger.error(`Batch rename operation ${i} failed:`, error);
        errors.push({
          index: i,
          operation: operation.type,
          error: error.message,
          itemId: operation.itemId || operation.fileId,
        });
      }
    }

    return {
      results,
      errors,
      summary: {
        total: renameOperations.length,
        successful: results.length,
        failed: errors.length,
      },
    };
  }

  // ========================
  // Name-based convenience APIs (backward compatible wrappers)
  // ========================
  async renameFileByName(
    accessToken,
    { driveName, fileName, newName, itemPath = null, auditContext = {} }
  ) {
    if (!driveName || !fileName || !newName) {
      throw new AppError(
        "driveName, fileName, and newName are required",
        400
      );
    }

    const driveId = await resolverService.resolveDriveIdByName(
      accessToken,
      driveName
    );

    let itemId;
    if (itemPath) {
      itemId = await resolverService.resolveItemIdByPath(
        accessToken,
        driveId,
        fileName,
        itemPath
      );
    } else {
      itemId = await resolverService.resolveItemIdByName(
        accessToken,
        driveId,
        fileName
      );
    }

    return this.renameFile(
      accessToken,
      driveId,
      itemId,
      fileName,
      newName,
      auditContext
    );
  }

  async renameFolderByName(
    accessToken,
    { driveName, folderName, newName, folderPath = null, auditContext = {} }
  ) {
    if (!driveName || !folderName || !newName) {
      throw new AppError(
        "driveName, folderName, and newName are required",
        400
      );
    }

    const driveId = await resolverService.resolveDriveIdByName(
      accessToken,
      driveName
    );

    // Resolve folderId by name (and optional exact path)
    const graphClient = this.createGraphClient(accessToken);
    const candidates = await this.recursiveSearch(
      graphClient,
      driveId,
      folderName
    );

    const folderCandidates = (candidates || []).filter((c) => c.folder);
    if (folderCandidates.length === 0) {
      throw new AppError(
        `Folder '${folderName}' not found in drive '${driveName}'`,
        404
      );
    }

    let target;
    if (folderPath) {
      target = folderCandidates.find((c) => c.path === folderPath);
      if (!target) {
        const paths = folderCandidates.map((c, i) => `${i + 1}. ${c.path}`).join("\n");
        const err = new AppError(
          `Folder '${folderName}' not found at path '${folderPath}'. Available paths:\n${paths}`,
          404
        );
        err.matches = folderCandidates;
        throw err;
      }
    } else if (folderCandidates.length === 1) {
      target = folderCandidates[0];
    } else {
      const paths = folderCandidates.map((c, i) => `${i + 1}. ${c.path}`).join("\n");
      const err = new AppError(
        `Multiple folders named '${folderName}' found. Please specify folderPath or choose one:\n${paths}`,
        409
      );
      err.matches = folderCandidates;
      err.isMultipleMatches = true;
      throw err;
    }

    return this.renameFolder(
      accessToken,
      driveId,
      target.id,
      folderName,
      newName,
      auditContext
    );
  }

  async renameSheetByName(
    accessToken,
    {
      driveName,
      fileName,
      oldSheetName,
      newSheetName,
      itemPath = null,
      auditContext = {},
    }
  ) {
    if (!driveName || !fileName || !oldSheetName || !newSheetName) {
      throw new AppError(
        "driveName, fileName, oldSheetName, and newSheetName are required",
        400
      );
    }

    const driveId = await resolverService.resolveDriveIdByName(
      accessToken,
      driveName
    );

    let itemId;
    if (itemPath) {
      itemId = await resolverService.resolveItemIdByPath(
        accessToken,
        driveId,
        fileName,
        itemPath
      );
    } else {
      itemId = await resolverService.resolveItemIdByName(
        accessToken,
        driveId,
        fileName
      );
    }

    return this.renameSheet(
      accessToken,
      driveId,
      itemId,
      oldSheetName,
      newSheetName,
      auditContext
    );
  }

  async batchRenameByName(
    accessToken,
    driveName,
    renameOperations,
    auditContext = {}
  ) {
    const driveId = await resolverService.resolveDriveIdByName(
      accessToken,
      driveName
    );

    const mappedOps = [];

    for (const op of renameOperations) {
      if (op.type === "file") {
        const itemId = op.itemPath
          ? await resolverService.resolveItemIdByPath(
              accessToken,
              driveId,
              op.fileName || op.oldName,
              op.itemPath
            )
          : await resolverService.resolveItemIdByName(
              accessToken,
              driveId,
              op.fileName || op.oldName
            );
        mappedOps.push({
          type: "file",
          itemId,
          oldName: op.oldName || op.fileName,
          newName: op.newName,
        });
      } else if (op.type === "folder") {
        // Resolve folder similar to renameFolderByName
        const graphClient = this.createGraphClient(accessToken);
        const candidates = await this.recursiveSearch(
          graphClient,
          driveId,
          op.folderName || op.oldName
        );
        const folderCandidates = (candidates || []).filter((c) => c.folder);
        let target;
        if (op.folderPath) {
          target = folderCandidates.find((c) => c.path === op.folderPath);
        } else if (folderCandidates.length === 1) {
          target = folderCandidates[0];
        } else {
          const paths = folderCandidates.map((c) => c.path);
          const err = new AppError(
            `Multiple or zero folders found for '${op.folderName || op.oldName}'. Provide folderPath. Available paths: ${JSON.stringify(
              paths
            )}`,
            409
          );
          err.matches = folderCandidates;
          throw err;
        }

        mappedOps.push({
          type: "folder",
          itemId: target.id,
          oldName: op.oldName || op.folderName,
          newName: op.newName,
        });
      } else if (op.type === "sheet") {
        const itemId = op.itemPath
          ? await resolverService.resolveItemIdByPath(
              accessToken,
              driveId,
              op.fileName,
              op.itemPath
            )
          : await resolverService.resolveItemIdByName(
              accessToken,
              driveId,
              op.fileName
            );
        mappedOps.push({
          type: "sheet",
          fileId: itemId,
          oldName: op.oldSheetName || op.oldName,
          newName: op.newSheetName || op.newName,
        });
      } else {
        throw new AppError(`Unknown operation type: ${op.type}`, 400);
      }
    }

    return this.batchRename(accessToken, driveId, mappedOps, auditContext);
  }
}

module.exports = new RenameService();
