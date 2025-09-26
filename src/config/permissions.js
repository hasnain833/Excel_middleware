const logger = require("./logger");

class PermissionService {
  constructor() {
    this.permissions = {
      // Global admin has access to everything
      admins: ["admin@company.com", "system"],

      // Workbook-level permissions
      workbooks: {
        // Example: specific workbook permissions
        "workbook-id": {
          readers: ["user1@company.com", "user2@company.com"],
          writers: ["user1@company.com"],
          admins: ["admin@company.com"],
        },
      },

      // Worksheet-level permissions
      worksheets: {
        // Example: specific worksheet permissions
        "workbook-id:worksheet-id": {
          readers: ["user1@company.com"],
          writers: ["user1@company.com"],
          locked: false,
        },
      },

      // Range-level permissions (most granular)
      ranges: {
        // Example: specific range permissions
        "workbook-id:worksheet-id:A1:C10": {
          readers: ["user1@company.com"],
          writers: [], // Read-only range
          locked: true,
        },
      },

      // Table-level permissions
      tables: {
        // Example: specific table permissions
        "workbook-id:worksheet-id:TableName": {
          readers: ["user1@company.com", "user2@company.com"],
          writers: ["user1@company.com"],
          locked: false,
        },
      },

      // Default permissions for new resources
      defaults: {
        allowReadAll: true, // Allow read access to all users by default
        allowWriteAll: false, // Require explicit write permissions
        inheritFromParent: true, // Inherit permissions from parent resource
      },
    };
  }

  isAdmin(userId) {
    return this.permissions.admins.includes(userId);
  }

  canAccessWorkbook(userId, workbookId) {
    // Admins can access everything
    if (this.isAdmin(userId)) {
      return true;
    }

    const workbookPerms = this.permissions.workbooks[workbookId];

    if (workbookPerms) {
      return (
        workbookPerms.readers?.includes(userId) ||
        workbookPerms.writers?.includes(userId) ||
        workbookPerms.admins?.includes(userId)
      );
    }

    return this.permissions.defaults.allowReadAll;
  }

  canAccessWorksheet(userId, workbookId, worksheetId) {
    // Admins can access everything
    if (this.isAdmin(userId)) {
      return true;
    }

    const worksheetKey = `${workbookId}:${worksheetId}`;
    const worksheetPerms = this.permissions.worksheets[worksheetKey];

    if (worksheetPerms) {
      return (
        worksheetPerms.readers?.includes(userId) ||
        worksheetPerms.writers?.includes(userId)
      );
    }
    if (this.permissions.defaults.inheritFromParent) {
      return this.canAccessWorkbook(userId, workbookId);
    }

    return this.permissions.defaults.allowReadAll;
  }

  canReadRange(userId, workbookId, worksheetId, range) {
    // Admins can read everything
    if (this.isAdmin(userId)) {
      return { allowed: true, reason: "Admin access" };
    }

    const rangeKey = `${workbookId}:${worksheetId}:${range}`;
    const rangePerms = this.permissions.ranges[rangeKey];

    if (rangePerms) {
      if (
        rangePerms.readers?.includes(userId) ||
        rangePerms.writers?.includes(userId)
      ) {
        return { allowed: true, reason: "Explicit range permission" };
      } else {
        return { allowed: false, reason: "No permission for this range" };
      }
    }

    // Check worksheet-level permissions
    if (this.permissions.defaults.inheritFromParent) {
      if (this.canAccessWorksheet(userId, workbookId, worksheetId)) {
        return { allowed: true, reason: "Inherited from worksheet" };
      }
    }

    // Default behavior
    if (this.permissions.defaults.allowReadAll) {
      return { allowed: true, reason: "Default read access" };
    }

    return { allowed: false, reason: "No read permission" };
  }

  canWriteRange(userId, workbookId, worksheetId, range) {
    // Admins can write everything
    if (this.isAdmin(userId)) {
      return { allowed: true, reason: "Admin access" };
    }

    const rangeKey = `${workbookId}:${worksheetId}:${range}`;
    const rangePerms = this.permissions.ranges[rangeKey];

    if (rangePerms) {
      if (rangePerms.locked) {
        return { allowed: false, reason: "Range is locked" };
      }
      if (rangePerms.writers?.includes(userId)) {
        return { allowed: true, reason: "Explicit range write permission" };
      } else {
        return { allowed: false, reason: "No write permission for this range" };
      }
    }

    // Check worksheet-level permissions
    const worksheetKey = `${workbookId}:${worksheetId}`;
    const worksheetPerms = this.permissions.worksheets[worksheetKey];

    if (worksheetPerms) {
      if (worksheetPerms.locked) {
        return { allowed: false, reason: "Worksheet is locked" };
      }
      if (worksheetPerms.writers?.includes(userId)) {
        return { allowed: true, reason: "Worksheet write permission" };
      }
    }

    // Check workbook-level permissions
    const workbookPerms = this.permissions.workbooks[workbookId];
    if (workbookPerms && workbookPerms.writers?.includes(userId)) {
      return { allowed: true, reason: "Workbook write permission" };
    }

    // Default behavior - no write access unless explicitly granted
    return { allowed: false, reason: "No write permission" };
  }

  canReadTable(userId, workbookId, worksheetId, tableName) {
    // Admins can read everything
    if (this.isAdmin(userId)) {
      return { allowed: true, reason: "Admin access" };
    }

    const tableKey = `${workbookId}:${worksheetId}:${tableName}`;
    const tablePerms = this.permissions.tables[tableKey];

    if (tablePerms) {
      if (
        tablePerms.readers?.includes(userId) ||
        tablePerms.writers?.includes(userId)
      ) {
        return { allowed: true, reason: "Explicit table permission" };
      } else {
        return { allowed: false, reason: "No permission for this table" };
      }
    }

    // Inherit from worksheet permissions
    if (this.permissions.defaults.inheritFromParent) {
      if (this.canAccessWorksheet(userId, workbookId, worksheetId)) {
        return { allowed: true, reason: "Inherited from worksheet" };
      }
    }

    return {
      allowed: this.permissions.defaults.allowReadAll,
      reason: "Default access",
    };
  }

  canWriteTable(userId, workbookId, worksheetId, tableName) {
    // Admins can write everything
    if (this.isAdmin(userId)) {
      return { allowed: true, reason: "Admin access" };
    }

    const tableKey = `${workbookId}:${worksheetId}:${tableName}`;
    const tablePerms = this.permissions.tables[tableKey];

    if (tablePerms) {
      if (tablePerms.locked) {
        return { allowed: false, reason: "Table is locked" };
      }
      if (tablePerms.writers?.includes(userId)) {
        return { allowed: true, reason: "Explicit table write permission" };
      } else {
        return { allowed: false, reason: "No write permission for this table" };
      }
    }

    // Check worksheet-level permissions
    const worksheetKey = `${workbookId}:${worksheetId}`;
    const worksheetPerms = this.permissions.worksheets[worksheetKey];

    if (worksheetPerms && worksheetPerms.writers?.includes(userId)) {
      return { allowed: true, reason: "Worksheet write permission" };
    }

    // Default behavior - no write access unless explicitly granted
    return { allowed: false, reason: "No write permission" };
  }

  addPermission(resourceType, resourceId, userId, permission) {
    try {
      const permissionMap = this.permissions[`${resourceType}s`];
      if (!permissionMap) {
        throw new Error(`Invalid resource type: ${resourceType}`);
      }

      if (!permissionMap[resourceId]) {
        permissionMap[resourceId] = { readers: [], writers: [], admins: [] };
      }

      const permissionList = permissionMap[resourceId][`${permission}s`];
      if (permissionList && !permissionList.includes(userId)) {
        permissionList.push(userId);
        logger.info(
          `Added ${permission} permission for ${userId} to ${resourceType} ${resourceId}`
        );
      }
    } catch (error) {
      logger.error("Failed to add permission:", error);
      throw error;
    }
  }

  removePermission(resourceType, resourceId, userId, permission) {
    try {
      const permissionMap = this.permissions[`${resourceType}s`];
      if (!permissionMap || !permissionMap[resourceId]) {
        return;
      }

      const permissionList = permissionMap[resourceId][`${permission}s`];
      if (permissionList) {
        const index = permissionList.indexOf(userId);
        if (index > -1) {
          permissionList.splice(index, 1);
          logger.info(
            `Removed ${permission} permission for ${userId} from ${resourceType} ${resourceId}`
          );
        }
      }
    } catch (error) {
      logger.error("Failed to remove permission:", error);
      throw error;
    }
  }

  getUserPermissions(userId) {
    const userPermissions = {
      isAdmin: this.isAdmin(userId),
      workbooks: { read: [], write: [], admin: [] },
      worksheets: { read: [], write: [] },
      ranges: { read: [], write: [] },
      tables: { read: [], write: [] },
    };

    // Collect all permissions for the user
    Object.entries(this.permissions.workbooks).forEach(([id, perms]) => {
      if (perms.readers?.includes(userId))
        userPermissions.workbooks.read.push(id);
      if (perms.writers?.includes(userId))
        userPermissions.workbooks.write.push(id);
      if (perms.admins?.includes(userId))
        userPermissions.workbooks.admin.push(id);
    });

    Object.entries(this.permissions.worksheets).forEach(([id, perms]) => {
      if (perms.readers?.includes(userId))
        userPermissions.worksheets.read.push(id);
      if (perms.writers?.includes(userId))
        userPermissions.worksheets.write.push(id);
    });

    Object.entries(this.permissions.ranges).forEach(([id, perms]) => {
      if (perms.readers?.includes(userId)) userPermissions.ranges.read.push(id);
      if (perms.writers?.includes(userId))
        userPermissions.ranges.write.push(id);
    });

    Object.entries(this.permissions.tables).forEach(([id, perms]) => {
      if (perms.readers?.includes(userId)) userPermissions.tables.read.push(id);
      if (perms.writers?.includes(userId))
        userPermissions.tables.write.push(id);
    });

    return userPermissions;
  }
}

module.exports = new PermissionService();