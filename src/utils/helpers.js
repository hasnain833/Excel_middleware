/**
 * Utility Helper Functions
 * Common utility functions used throughout the application
 */

const { v4: uuidv4 } = require('uuid');
const logger = require('../config/logger');

/**
 * Generate a unique request ID
 * @returns {string} Unique request ID
 */
const generateRequestId = () => {
    return uuidv4();
};

/**
 * Convert Excel column letter to number (A=1, B=2, etc.)
 * @param {string} column - Column letter(s)
 * @returns {number} Column number
 */
const columnLetterToNumber = (column) => {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
        result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
};

/**
 * Convert column number to Excel letter (1=A, 2=B, etc.)
 * @param {number} num - Column number
 * @returns {string} Column letter(s)
 */
const columnNumberToLetter = (num) => {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
};

/**
 * Parse Excel range into components
 * @param {string} range - Range string (e.g., 'A1:C10')
 * @returns {Object} Parsed range components
 */
const parseExcelRange = (range) => {
    try {
        const parts = range.split(':');
        
        if (parts.length === 1) {
            // Single cell
            const match = parts[0].match(/^([A-Z]+)(\d+)$/);
            if (!match) throw new Error('Invalid cell format');
            
            return {
                startColumn: match[1],
                startRow: parseInt(match[2]),
                endColumn: match[1],
                endRow: parseInt(match[2]),
                isSingleCell: true,
                rowCount: 1,
                columnCount: 1
            };
        } else if (parts.length === 2) {
            // Range
            const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
            const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
            
            if (!startMatch || !endMatch) throw new Error('Invalid range format');
            
            const startCol = columnLetterToNumber(startMatch[1]);
            const endCol = columnLetterToNumber(endMatch[1]);
            const startRow = parseInt(startMatch[2]);
            const endRow = parseInt(endMatch[2]);
            
            return {
                startColumn: startMatch[1],
                startRow: startRow,
                endColumn: endMatch[1],
                endRow: endRow,
                isSingleCell: false,
                rowCount: endRow - startRow + 1,
                columnCount: endCol - startCol + 1
            };
        }
        
        throw new Error('Invalid range format');
    } catch (error) {
        logger.error('Failed to parse Excel range:', error);
        throw new Error(`Invalid range format: ${range}`);
    }
};

/**
 * Validate Excel range format
 * @param {string} range - Range string
 * @returns {boolean} True if valid
 */
const isValidExcelRange = (range) => {
    const rangeRegex = /^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/;
    return rangeRegex.test(range);
};

/**
 * Convert 2D array to CSV string
 * @param {Array} data - 2D array of data
 * @returns {string} CSV string
 */
const arrayToCSV = (data) => {
    if (!Array.isArray(data) || data.length === 0) {
        return '';
    }
    
    return data.map(row => 
        row.map(cell => {
            const cellStr = String(cell || '');
            // Escape quotes and wrap in quotes if contains comma, quote, or newline
            if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                return `"${cellStr.replace(/"/g, '""')}"`;
            }
            return cellStr;
        }).join(',')
    ).join('\n');
};

/**
 * Deep clone an object
 * @param {*} obj - Object to clone
 * @returns {*} Cloned object
 */
const deepClone = (obj) => {
    if (obj === null || typeof obj !== 'object') {
        return obj;
    }
    
    if (obj instanceof Date) {
        return new Date(obj.getTime());
    }
    
    if (Array.isArray(obj)) {
        return obj.map(item => deepClone(item));
    }
    
    const cloned = {};
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            cloned[key] = deepClone(obj[key]);
        }
    }
    
    return cloned;
};

/**
 * Sanitize filename for safe storage
 * @param {string} filename - Original filename
 * @returns {string} Sanitized filename
 */
const sanitizeFilename = (filename) => {
    return filename
        .replace(/[<>:"/\\|?*]/g, '_')
        .replace(/\s+/g, '_')
        .toLowerCase();
};

/**
 * Format file size in human readable format
 * @param {number} bytes - Size in bytes
 * @returns {string} Formatted size
 */
const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

/**
 * Check if a string is a valid email
 * @param {string} email - Email string
 * @returns {boolean} True if valid email
 */
const isValidEmail = (email) => {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
};

/**
 * Retry function with exponential backoff
 * @param {Function} fn - Function to retry
 * @param {number} maxRetries - Maximum number of retries
 * @param {number} baseDelay - Base delay in milliseconds
 * @returns {Promise} Promise that resolves with function result
 */
const retryWithBackoff = async (fn, maxRetries = 3, baseDelay = 1000) => {
    let lastError;
    
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
            return await fn();
        } catch (error) {
            lastError = error;
            
            if (attempt === maxRetries) {
                break;
            }
            
            const delay = baseDelay * Math.pow(2, attempt);
            logger.warn(`Attempt ${attempt + 1} failed, retrying in ${delay}ms:`, error.message);
            
            await new Promise(resolve => setTimeout(resolve, delay));
        }
    }
    
    throw lastError;
};

/**
 * Create a debounced function
 * @param {Function} func - Function to debounce
 * @param {number} wait - Wait time in milliseconds
 * @returns {Function} Debounced function
 */
const debounce = (func, wait) => {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
};

/**
 * Create a throttled function
 * @param {Function} func - Function to throttle
 * @param {number} limit - Time limit in milliseconds
 * @returns {Function} Throttled function
 */
const throttle = (func, limit) => {
    let inThrottle;
    return function(...args) {
        if (!inThrottle) {
            func.apply(this, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
};

/**
 * Sleep for specified milliseconds
 * @param {number} ms - Milliseconds to sleep
 * @returns {Promise} Promise that resolves after delay
 */
const sleep = (ms) => {
    return new Promise(resolve => setTimeout(resolve, ms));
};

module.exports = {
    generateRequestId,
    columnLetterToNumber,
    columnNumberToLetter,
    parseExcelRange,
    isValidExcelRange,
    arrayToCSV,
    deepClone,
    sanitizeFilename,
    formatFileSize,
    isValidEmail,
    retryWithBackoff,
    debounce,
    throttle,
    sleep
};
