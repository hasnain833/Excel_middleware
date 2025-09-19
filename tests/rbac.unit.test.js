import { isAllowed, rolePermissions } from '../config/roles.js';
import { extractRoleFromDecoded, extractUserRole, requirePermission } from '../api/middleware/extractUserRole.js';

function mockReq(headers = {}) { return { headers, method: 'POST', path: '/test' }; }
function mockRes() {
  const res = {};
  res.statusCode = 200;
  res.status = (code) => { res.statusCode = code; return res; };
  res.body = null;
  res.json = (obj) => { res.body = obj; return res; };
  return res;
}
function nextSpy() { const fn = jest.fn(); return fn; }

describe('RBAC config', () => {
  test('rolePermissions defined correctly', () => {
    expect(rolePermissions.admin.canWrite).toBe(true);
    expect(rolePermissions.editor.canWrite).toBe(true);
    expect(rolePermissions.viewer.canWrite).toBe(false);
  });
  test('isAllowed checks actions', () => {
    expect(isAllowed('admin', 'delete')).toBe(true);
    expect(isAllowed('editor', 'delete')).toBe(false);
    expect(isAllowed('viewer', 'write')).toBe(false);
  });
});

describe('extractRoleFromDecoded', () => {
  test('picks admin from roles array', () => {
    expect(extractRoleFromDecoded({ roles: ['Admin'] })).toBe('admin');
  });
  test('picks editor from appRole', () => {
    expect(extractRoleFromDecoded({ appRole: 'editor' })).toBe('editor');
  });
  test('fallback to viewer when unknown', () => {
    expect(extractRoleFromDecoded({ roles: ['unknown'] })).toBe('viewer');
  });
});

describe('extractUserRole middleware (dev mode)', () => {
  const OLD_ENV = process.env.NODE_ENV;
  beforeAll(() => { process.env.NODE_ENV = 'development'; });
  afterAll(() => { process.env.NODE_ENV = OLD_ENV; });

  test('accepts x-user-role header', async () => {
    const req = mockReq({ 'x-user-role': 'editor' });
    const res = mockRes();
    const next = nextSpy();
    await extractUserRole(req, res, next);
    expect(req.userRole).toBe('editor');
    expect(next).toHaveBeenCalled();
  });

  test('missing role header falls back to viewer and requires token in prod only', async () => {
    const req = mockReq();
    const res = mockRes();
    const next = nextSpy();
    await extractUserRole(req, res, next);
    // In dev, if no header is provided, we default to viewer and continue
    expect(req.userRole).toBe('viewer');
    expect(next).toHaveBeenCalled();
  });
});

describe('requirePermission middleware', () => {
  test('forbids when role lacks permission', () => {
    const req = { userRole: 'viewer', method: 'POST', path: '/excel/write' };
    const res = mockRes();
    const next = nextSpy();
    const mw = requirePermission('write');
    mw(req, res, next);
    expect(res.statusCode).toBe(403);
    expect(next).not.toHaveBeenCalled();
  });
  test('allows when role has permission', () => {
    const req = { userRole: 'admin', method: 'POST', path: '/excel/write' };
    const res = mockRes();
    const next = nextSpy();
    const mw = requirePermission('write');
    mw(req, res, next);
    expect(res.statusCode).toBe(200);
    expect(next).toHaveBeenCalled();
  });
});
