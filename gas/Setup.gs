// Setup.gs
// Run setupSpreadsheet() once, manually, from the Apps Script editor: select
// "setupSpreadsheet" in the function dropdown next to the Run button, then
// click Run. Idempotent — existing sheets and non-empty seed data are left
// untouched, so re-running it later (e.g. after a new .gs file is added) is
// always safe and never destroys real data.

const DEFAULT_MONTHS_ = [
  { key: 'أبريل 2026', id: 4 },
  { key: 'مايو 2026', id: 5 },
  { key: 'يونيو 2026', id: 6 },
  { key: 'يوليو 2026', id: 7 },
  { key: 'أغسطس 2026', id: 8 },
  { key: 'سبتمبر 2026', id: 9 },
  { key: 'أكتوبر 2026', id: 10 },
  { key: 'نوفمبر 2026', id: 11 },
  { key: 'ديسمبر 2026', id: 12 },
  { key: 'يناير 2027', id: 1 },
  { key: 'فبراير 2027', id: 2 },
  { key: 'مارس 2027', id: 3 }
];

const DEFAULT_USERS_ = [
  { key: 'admin1', name: 'جهاد زكري إسماعيل اطفيحة', role: 'رئيس اللجنة المالية للعشيرة', email: 'jehadgml@gmail.com', pass: 'ABC12345' },
  { key: 'admin2', name: 'أشرف يوسف محمود اطفيحة', role: 'نائب رئيس اللجنة المالية', email: 'ashraf.atfihah@gmail.com', pass: 'ABC12345' },
  { key: 'admin3', name: 'إبراهيم محمد إبراهيم اطفيحة', role: 'أمين صندوق العشيرة', email: 'ibrahim.atfihah@gmail.com', pass: 'ABC12345' },
  { key: 'admin4', name: 'معتز أمين محمد محمود اطفيحة', role: 'المراقب المالي للصندوق', email: 'moataz.atfihah@gmail.com', pass: 'ABC12345' },
  { key: 'admin5', name: 'حبيب محمود محمد اطفيحة', role: 'منسق التحصيل والوصولات', email: 'habib.atfihah@gmail.com', pass: 'ABC12345' }
];

const MASTER_ADMIN_EMAIL_ = 'jehadgml@gmail.com';
const ATTACHMENTS_FOLDER_NAME_ = 'CEMS-ATFIHAH Attachments';
const ATTACHMENTS_FOLDER_PROP_ = 'ATTACHMENTS_FOLDER_ID';

function setupSpreadsheet() {
  const ss = getSS_();

  ensureSheetWithHeaders_(ss, SHEET_NAMES.MEMBERS,
    ['id', 'name', 'parent'].concat(DEFAULT_MONTHS_.map(function (m) { return m.key; })).concat(['sum']));
  ensureSheetWithHeaders_(ss, SHEET_NAMES.FAMILIES,
    ['familyId', 'headName', 'memberCount', 'subscription', 'totalPaid', 'membersList']);
  const monthsSheet = ensureSheetWithHeaders_(ss, SHEET_NAMES.MONTHS, ['key', 'id', 'selected']);
  ensureSheetWithHeaders_(ss, SHEET_NAMES.EXPENSES,
    ['id', 'date', 'amount', 'reason', 'category', 'authorized', 'attachmentName', 'attachmentType', 'attachmentUrl']);
  const usersSheet = ensureSheetWithHeaders_(ss, SHEET_NAMES.USERS,
    ['key', 'name', 'role', 'email', 'pass', 'isBuiltIn', 'isMaster', 'isActive', 'firstLoginDone']);
  ensureSheetWithHeaders_(ss, SHEET_NAMES.PENDING_USERS, ['id', 'firstName', 'lastName', 'email', 'pass', 'requestedAt']);
  ensureSheetWithHeaders_(ss, SHEET_NAMES.SIGNATURES, ['month', 'slotIndex', 'signatureUrl', 'updatedAt']);
  ensureSheetWithHeaders_(ss, SHEET_NAMES.SETTINGS, ['key', 'value']);

  seedMonthsIfEmpty_(monthsSheet);
  seedUsersIfEmpty_(usersSheet);
  ensureAttachmentsFolder_();

  Logger.log('setupSpreadsheet() finished successfully.');
}

// Creates the sheet (with the given header row) only if it doesn't exist yet.
// Never touches an existing sheet's data.
function ensureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function seedMonthsIfEmpty_(monthsSheet) {
  const data = sheetToObjects_(monthsSheet);
  if (data.rows.length > 0) return;
  DEFAULT_MONTHS_.forEach(function (m) {
    monthsSheet.appendRow([m.key, m.id, true]);
  });
}

function seedUsersIfEmpty_(usersSheet) {
  const data = sheetToObjects_(usersSheet);
  if (data.rows.length > 0) return;
  DEFAULT_USERS_.forEach(function (u) {
    usersSheet.appendRow([u.key, u.name, u.role, u.email, u.pass, true, u.email === MASTER_ADMIN_EMAIL_, true, false]);
  });
}

function ensureAttachmentsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(ATTACHMENTS_FOLDER_PROP_);
  if (existingId) {
    try {
      DriveApp.getFolderById(existingId);
      return;
    } catch (e) {
      // stored id no longer valid — fall through and recreate
    }
  }
  const folder = DriveApp.createFolder(ATTACHMENTS_FOLDER_NAME_);
  props.setProperty(ATTACHMENTS_FOLDER_PROP_, folder.getId());
}
