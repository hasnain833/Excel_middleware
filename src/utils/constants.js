// HTTP Status Codes
const HTTP_STATUS = {
    OK: 200,
    CREATED: 201,
    ACCEPTED: 202,
    NO_CONTENT: 204,
    MULTI_STATUS: 207,
    BAD_REQUEST: 400,
    UNAUTHORIZED: 401,
    FORBIDDEN: 403,
    NOT_FOUND: 404,
    METHOD_NOT_ALLOWED: 405,
    CONFLICT: 409,
    UNPROCESSABLE_ENTITY: 422,
    TOO_MANY_REQUESTS: 429,
    INTERNAL_SERVER_ERROR: 500,
    BAD_GATEWAY: 502,
    SERVICE_UNAVAILABLE: 503,
    GATEWAY_TIMEOUT: 504
};

// Excel-specific constants
const EXCEL = {
    MAX_ROWS: 1048576,
    MAX_COLUMNS: 16384,
    MAX_COLUMN_LETTER: 'XFD',
    SUPPORTED_FILE_EXTENSIONS: ['.xlsx', '.xlsm', '.xltx', '.xltm'],
    MAX_CELL_LENGTH: 32767,
    MAX_WORKSHEET_NAME_LENGTH: 31,
    MAX_TABLE_NAME_LENGTH: 255
};

// Microsoft Graph API constants
const GRAPH_API = {
    BASE_URL: 'https://graph.microsoft.com/v1.0',
    SCOPES: {
        FILES_READ: 'https://graph.microsoft.com/Files.Read',
        FILES_READWRITE: 'https://graph.microsoft.com/Files.ReadWrite',
        SITES_READ: 'https://graph.microsoft.com/Sites.Read.All',
        SITES_READWRITE: 'https://graph.microsoft.com/Sites.ReadWrite.All'
    },
    ENDPOINTS: {
        ME: '/me',
        DRIVES: '/drives',
        SITES: '/sites',
        WORKBOOKS: '/workbook',
        WORKSHEETS: '/worksheets',
        TABLES: '/tables',
        RANGES: '/range'
    },
    LIMITS: {
        REQUEST_TIMEOUT: 30000, // 30 seconds
        MAX_BATCH_SIZE: 20,
        RATE_LIMIT_REQUESTS: 10000, // per hour
        MAX_FILE_SIZE: 250 * 1024 * 1024 // 250MB
    }
};

// Authentication constants
const AUTH = {
    TOKEN_BUFFER_TIME: 5 * 60 * 1000, // 5 minutes buffer before token expiry
    MAX_TOKEN_RETRIES: 3,
    TOKEN_REFRESH_THRESHOLD: 10 * 60 * 1000, // Refresh if expires within 10 minutes
    CLIENT_CREDENTIAL_SCOPE: 'https://graph.microsoft.com/.default'
};

// Logging levels
const LOG_LEVELS = {
    ERROR: 'error',
    WARN: 'warn',
    INFO: 'info',
    HTTP: 'http',
    VERBOSE: 'verbose',
    DEBUG: 'debug',
    SILLY: 'silly'
};

// Audit operation types
const AUDIT_OPERATIONS = {
    READ: 'READ',
    WRITE: 'WRITE',
    READ_TABLE: 'READ_TABLE',
    WRITE_TABLE: 'WRITE_TABLE',
    PERMISSION_CHECK: 'PERMISSION_CHECK',
    AUTHENTICATION: 'AUTHENTICATION',
    SYSTEM: 'SYSTEM',
    BATCH: 'BATCH_OPERATIONS'
};

// Permission types
const PERMISSIONS = {
    READ: 'read',
    WRITE: 'write',
    ADMIN: 'admin',
    NONE: 'none'
};

// User roles
const USER_ROLES = {
    ADMIN: 'admin',
    PREMIUM: 'premium',
    ENTERPRISE: 'enterprise',
    STANDARD: 'standard',
    GUEST: 'guest'
};

// Rate limiting constants
const RATE_LIMITS = {
    GENERAL: {
        WINDOW_MS: 15 * 60 * 1000, // 15 minutes
        MAX_REQUESTS: 100
    },
    WRITE: {
        WINDOW_MS: 5 * 60 * 1000, // 5 minutes
        MAX_REQUESTS: 20
    },
    AUTH: {
        WINDOW_MS: 15 * 60 * 1000, // 15 minutes
        MAX_REQUESTS: 5
    }
};

// Error types
const ERROR_TYPES = {
    VALIDATION_ERROR: 'ValidationError',
    AUTHENTICATION_ERROR: 'AuthenticationError',
    AUTHORIZATION_ERROR: 'AuthorizationError',
    GRAPH_API_ERROR: 'GraphApiError',
    NETWORK_ERROR: 'NetworkError',
    INTERNAL_ERROR: 'InternalError'
};

// Excel operation types for batch processing
const EXCEL_OPERATIONS = {
    READ_RANGE: 'read_range',
    WRITE_RANGE: 'write_range',
    READ_TABLE: 'read_table',
    ADD_TABLE_ROWS: 'add_table_rows',
    GET_WORKSHEETS: 'get_worksheets',
    GET_WORKBOOKS: 'get_workbooks'
};

// Validation patterns
const VALIDATION_PATTERNS = {
    EXCEL_RANGE: /^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/,
    EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    UUID: /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i,
    WORKSHEET_NAME: /^[^\\\/\?\*\[\]]{1,31}$/,
    TABLE_NAME: /^[a-zA-Z_][a-zA-Z0-9_]{0,254}$/
};

// Default configuration values
const DEFAULTS = {
    PORT: 3000,
    LOG_LEVEL: 'info',
    NODE_ENV: 'development',
    REQUEST_TIMEOUT: 30000,
    MAX_REQUEST_SIZE: '10mb',
    CORS_ORIGIN: '*',
    TRUST_PROXY: false
};

// Cache settings
const CACHE = {
    TOKEN_TTL: 60 * 60, // 1 hour in seconds
    WORKBOOK_LIST_TTL: 5 * 60, // 5 minutes
    WORKSHEET_LIST_TTL: 10 * 60, // 10 minutes
    MAX_CACHE_SIZE: 100 // Maximum number of cached items
};

// File size limits
const FILE_LIMITS = {
    MAX_UPLOAD_SIZE: 250 * 1024 * 1024, // 250MB
    MAX_CELLS_PER_REQUEST: 10000,
    MAX_BATCH_OPERATIONS: 20,
    MAX_TABLE_ROWS_PER_REQUEST: 1000
};

// Environment types
const ENVIRONMENTS = {
    DEVELOPMENT: 'development',
    STAGING: 'staging',
    PRODUCTION: 'production',
    TEST: 'test'
};

module.exports = {
    HTTP_STATUS,
    EXCEL,
    GRAPH_API,
    AUTH,
    LOG_LEVELS,
    AUDIT_OPERATIONS,
    PERMISSIONS,
    USER_ROLES,
    RATE_LIMITS,
    ERROR_TYPES,
    EXCEL_OPERATIONS,
    VALIDATION_PATTERNS,
    DEFAULTS,
    CACHE,
    FILE_LIMITS,
    ENVIRONMENTS
};
