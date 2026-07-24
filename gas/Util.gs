// Util.gs
// Shared constants and low-level sheet helpers used by every other .gs file.
// Paste this into the Apps Script editor FIRST — every other file calls into it.

const SHEET_NAMES = {
  MEMBERS: 'Members',
  FAMILIES: 'Families',
  MONTHS: 'Months',
  EXPENSES: 'Expenses',
  USERS: 'Users',
  PENDING_USERS: 'PendingUsers',
  SIGNATURES: 'Signatures',
  SETTINGS: 'Settings',
  AUDIT_LOG: 'AuditLog'
};

// Columns on the Members sheet that are NOT a month payment column.
const FIXED_MEMBER_COLUMNS = ['id', 'name', 'parent', 'sum'];

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const sheet = getSS_().getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" not found. Run setupSpreadsheet() first.');
  }
  return sheet;
}

// Reads a sheet into {headers, rows}. Each row object carries a "_row" field
// (1-based sheet row number) so callers can address it for updates/deletes.
function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) return { headers: [], rows: [] };
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const rowValues = values[i];
    if (rowValues.every(function (v) { return v === '' || v === null; })) continue;
    const obj = { _row: i + 1 };
    headers.forEach(function (h, idx) { obj[h] = rowValues[idx]; });
    rows.push(obj);
  }
  return { headers: headers, rows: rows };
}

function headerIndex_(headers, name) {
  const idx = headers.indexOf(name);
  if (idx === -1) throw new Error('Column "' + name + '" not found in headers: ' + headers.join(', '));
  return idx;
}

function nextId_(rows, idField) {
  let max = 0;
  rows.forEach(function (r) {
    const n = Number(r[idField || 'id']);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

// Appends one row built from `obj` in the sheet's current header order.
function appendRowFromObject_(sheet, headers, obj) {
  const row = headers.map(function (h) { return (h in obj) ? obj[h] : ''; });
  sheet.appendRow(row);
}

// Overwrites an existing row (1-based sheet row number) with `obj`, in the
// sheet's current header order.
function writeRowFromObject_(sheet, headers, rowNumber, obj) {
  const row = headers.map(function (h) { return (h in obj) ? obj[h] : ''; });
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
}

function findRowById_(rows, id, idField) {
  const field = idField || 'id';
  const match = rows.filter(function (r) { return String(r[field]) === String(id); });
  return match.length ? match[0] : null;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Arabic name normalization — mirrors the frontend's normalizeArabicName() in
// app.js exactly, so a family head name typed with slightly different letter
// forms (أ/إ/آ vs ا, ة vs ه, ى vs ي) still matches server-side.
function normalizeArabicName_(name) {
  if (!name) return '';
  return String(name).trim().toLowerCase()
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    .replace(/\s+/g, ' ');
}
