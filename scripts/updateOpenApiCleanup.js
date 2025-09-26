const fs = require('fs');
const path = require('path');

(function main() {
  const file = path.resolve(__dirname, '..', 'openAPI.json');
  if (!fs.existsSync(file)) {
    console.error('openAPI.json not found at', file);
    process.exit(1);
  }

  const backup = path.resolve(__dirname, '..', 'openAPI.backup.json');
  fs.copyFileSync(file, backup);

  const doc = JSON.parse(fs.readFileSync(file, 'utf8'));

  // Keep only the routes that exist in src/routes/*.js
  const allowedPaths = new Set([
    '/health',
    '/api/excel/workbooks',
    '/api/excel/worksheets',
    '/api/excel/read',
    '/api/excel/write',
    '/api/excel/batch',
    '/api/excel/search',
    '/api/excel/find-replace',
    '/api/excel/search-text',
    '/api/excel/analyze-scope',
    '/api/excel/format',
    '/api/excel/create-file',
    '/api/excel/create-sheet',
    '/api/excel/delete-file',
    '/api/excel/delete-sheet',
    '/api/excel/rename-file',
    '/api/excel/rename-folder',
    '/api/excel/rename-sheet',
    '/api/excel/rename-suggestions',
    '/api/excel/batch-rename',
  ]);

  const beforePaths = Object.keys(doc.paths || {}).length;
  Object.keys(doc.paths || {}).forEach((k) => {
    if (!allowedPaths.has(k)) {
      delete doc.paths[k];
    }
  });
  const afterPaths = Object.keys(doc.paths || {}).length;

  // Remove unused schemas
  const removeSchemas = [
    'ReadTableRequest',
    'ReadTableSuccess',
    'AddTableRowsRequest',
    'ValidateFormulaRequest',
    'ValidateFormulaSuccess',
  ];
  let removedSchemas = 0;
  if (doc.components && doc.components.schemas) {
    for (const s of removeSchemas) {
      if (s in doc.components.schemas) {
        delete doc.components.schemas[s];
        removedSchemas++;
      }
    }
  }

  fs.writeFileSync(file, JSON.stringify(doc, null, 2));
  console.log(
    `openAPI.json cleaned. Paths: ${beforePaths} -> ${afterPaths}. Schemas removed: ${removedSchemas}. Backup: ${path.basename(backup)}`
  );
})();
