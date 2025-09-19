import jwt from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import { isAllowed } from '../../config/roles.js';

// Helper: build JWKS client for Azure AD
function buildJwksClient() {
  const TENANT_ID = process.env.AZURE_TENANT_ID || process.env.TENANT_ID;
  if (!TENANT_ID) return null;
  const jwksUri = `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`;
  return jwksClient({
    jwksUri,
    cache: true,
    cacheMaxEntries: 5,
    cacheMaxAge: 10 * 60 * 1000,
    rateLimit: true,
    jwksRequestsPerMinute: 10,
  });
}

const client = buildJwksClient();

// Resolve signing key for JWT verification
function getKey(header, callback) {
  if (!client || !header.kid) {
    return callback(new Error('JWKS client not configured or missing KID'));
  }
  client.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    const signingKey = key?.getPublicKey?.() || key?.publicKey || key?.rsaPublicKey;
    callback(null, signingKey);
  });
}

// Extract role from a decoded token payload
export function extractRoleFromDecoded(decoded) {
  if (!decoded || typeof decoded !== 'object') return 'viewer';

  // Possible locations for roles in Azure AD tokens
  const candidates = [];
  if (Array.isArray(decoded.roles)) candidates.push(...decoded.roles);
  if (decoded.appRole) candidates.push(decoded.appRole);
  if (Array.isArray(decoded.groups)) candidates.push(...decoded.groups);
  if (typeof decoded.group === 'string') candidates.push(decoded.group);

  // Normalize to lower-case strings
  const names = candidates
    .filter(Boolean)
    .map((v) => String(v).toLowerCase());

  if (names.includes('admin')) return 'admin';
  if (names.includes('editor')) return 'editor';
  if (names.includes('viewer')) return 'viewer';

  // Fallback to most restrictive
  return 'viewer';
}

// Middleware to extract user role
export async function extractUserRole(req, res, next) {
  try {
    const env = process.env.NODE_ENV || 'development';

    // Non-production: allow testing header x-user-role
    if (env !== 'production') {
      const headerRole = (req.headers['x-user-role'] || req.headers['x-userrole'] || '').toString().toLowerCase();
      if (headerRole === 'admin' || headerRole === 'editor' || headerRole === 'viewer') {
        req.userRole = headerRole;
        return next();
      }
      // Default to most restrictive if not provided
      req.userRole = 'viewer';
      return next();
    }

    // Production: must validate Authorization: Bearer <token>
    const auth = req.headers.authorization || '';
    const parts = auth.split(' ');
    if (parts.length !== 2 || parts[0] !== 'Bearer') {
      console.warn(`[RBAC] Missing or malformed Authorization header for ${req.method} ${req.path}`);
      return res.status(401).json({ success: false, error: 'Unauthorized: Bearer token required.' });
    }
    const token = parts[1];

    const TENANT_ID = process.env.AZURE_TENANT_ID || process.env.TENANT_ID;
    const AUDIENCE = process.env.AZURE_CLIENT_ID || process.env.AZURE_AUDIENCE || process.env.CLIENT_ID;
    const ISSUER = TENANT_ID ? new RegExp(`^https://login.microsoftonline.com/${TENANT_ID}/v2.0/?$`) : undefined;

    const verifyOptions = {
      algorithms: ['RS256', 'RS384', 'RS512', 'HS256'],
      audience: AUDIENCE ? [AUDIENCE] : undefined,
      issuer: ISSUER,
    };

    // Prefer RS256 via JWKS if configured; fallback to symmetric secret if provided
    const JWT_SECRET = process.env.JWT_SECRET;

    const verifyWithKey = () => new Promise((resolve, reject) => {
      if (client) {
        jwt.verify(token, getKey, verifyOptions, (err, decoded) => {
          if (err) return reject(err);
          return resolve(decoded);
        });
      } else if (JWT_SECRET) {
        jwt.verify(token, JWT_SECRET, verifyOptions, (err, decoded) => {
          if (err) return reject(err);
          return resolve(decoded);
        });
      } else {
        return reject(new Error('No JWKS client or JWT_SECRET configured for token verification'));
      }
    });

    let decoded;
    try {
      decoded = await verifyWithKey();
    } catch (e) {
      console.warn(`[RBAC] Token verification failed for ${req.method} ${req.path}: ${e.message}`);
      return res.status(401).json({ success: false, error: 'Unauthorized: Invalid token.' });
    }

    const role = extractRoleFromDecoded(decoded);
    req.userRole = role || 'viewer';
    return next();
  } catch (err) {
    console.error('[RBAC] Unexpected error in extractUserRole:', err);
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }
}

// Factory middleware to enforce permission for an action
export function requirePermission(action) {
  return function (req, res, next) {
    const role = (req.userRole || 'viewer').toLowerCase();
    const allowed = isAllowed(role, action);
    if (!allowed) {
      console.warn(`[RBAC] Forbidden: role='${role}' action='${action}' path='${req.method} ${req.path}'`);
      return res.status(403).json({ success: false, error: `Forbidden: role '${role}' is not allowed to ${action}.` });
    }
    return next();
  };
}
