// Role definitions and permission map
// Extendable config-driven permissions

export const rolePermissions = {
  admin:  { canWrite: true, canDelete: true, canCreate: true },
  editor: { canWrite: true, canDelete: false, canCreate: false },
  viewer: { canWrite: false, canDelete: false, canCreate: false },
};

// Helper to check if a role is allowed for a given action
// action: 'write' | 'delete' | 'create'
export function isAllowed(role = 'viewer', action) {
  const safeRole = (role || 'viewer').toLowerCase();
  const perms = rolePermissions[safeRole] || rolePermissions.viewer;
  switch (action) {
    case 'write':
      return !!perms.canWrite;
    case 'delete':
      return !!perms.canDelete;
    case 'create':
      return !!perms.canCreate;
    default:
      return false;
  }
}
