# Google Sheets Backend Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the fund-tracker app's `localStorage` persistence with a Google Sheet, accessed through a Google Apps Script (GAS) Web App JSON API, so every piece of data (members, families, months, expenses, accounts, pending registrations, signatures) lives in the Sheet.

**Architecture:** A set of `.gs` files (kept in this repo under `gas/` for the user to paste into the Sheet's Apps Script editor) implement a single `doPost` JSON-RPC dispatcher over one sheet per entity, with large binary attachments (signatures, expense receipts) stored in a Drive folder and referenced by URL. `app.js` is rewritten to call this API instead of `localStorage` for every piece of app data.

**Tech Stack:** Google Apps Script (`SpreadsheetApp`, `DriveApp`, `LockService`, `ContentService`), vanilla JS `fetch` on the frontend. No test framework exists in this repo for either language — every task is verified manually (Apps Script editor execution log + browser).

**Spec:** [docs/superpowers/specs/2026-07-20-google-sheets-backend-migration-design.md](../specs/2026-07-20-google-sheets-backend-migration-design.md)

---

## Before you start

- This plan produces files under `gas/*.gs` that are **not executed by anything in this repo**. They only become real once the user manually creates a Google Sheet, opens Extensions → Apps Script, creates a matching `.gs` file per task and pastes the content in, then runs `setupSpreadsheet()` and deploys. Task "Backend 14" is that manual deployment guide.
- Because of this, "verify" steps for backend tasks describe what the *user* (or an operator with access to the real Google Sheet) checks in the Apps Script editor / Spreadsheet UI — they cannot be run by an automated agent in this repo. Say so plainly when executing those tasks rather than pretending to have run them.
- Frontend (`app.js`) tasks *can* be verified in-repo up to the point of needing a live `API_URL` — static/syntax checks and a browser load can be done without a deployment; full click-through verification needs the user's deployed URL (see Frontend 11).
- Every task ends with a `git commit`. Commit small — one task, one commit.

---

## Backend Task 1: `gas/Util.gs` — shared constants and sheet helpers

**Files:**
- Create: `gas/Util.gs`

- [ ] **Step 1: Write the file**

```javascript
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
  SETTINGS: 'Settings'
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
```

- [ ] **Step 2: Commit**

```bash
git add gas/Util.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Util.gs shared sheet helpers

Foundational helper layer for the Apps Script backend — sheet-to-object
mapping, header lookups, id generation, and Arabic name normalization
matching the frontend's existing normalizeArabicName().

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 2: `gas/Setup.gs` — one-time idempotent spreadsheet setup

**Files:**
- Create: `gas/Setup.gs`

- [ ] **Step 1: Write the file**

```javascript
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
```

Note: Members/Families are intentionally **not** seeded with any default roster — those sheets start empty. The user repopulates them via the app's existing Excel-import feature (Frontend Task 10) or by adding members through the UI. This keeps `Setup.gs` from having to embed the ~73KB `data.js` dataset.

- [ ] **Step 2: Commit**

```bash
git add gas/Setup.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Setup.gs idempotent spreadsheet initializer

Creates all 8 backend sheets with headers, seeds Months (12 default
months) and Users (the 5 committee accounts) only on first run, and
creates the Drive attachments folder.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 3: `gas/DriveHelper.gs` — attachment upload helper

**Files:**
- Create: `gas/DriveHelper.gs`

- [ ] **Step 1: Write the file**

```javascript
// DriveHelper.gs
// Shared helper for uploading a data-URL payload (as produced by the
// frontend's FileReader-based readAttachmentFile() / the signature canvas's
// toDataURL()) into the attachments Drive folder, returning a URL usable
// directly as an <img src> or download link.

function getAttachmentsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(ATTACHMENTS_FOLDER_PROP_);
  if (!id) throw new Error('Attachments folder not set up yet. Run setupSpreadsheet() first.');
  return DriveApp.getFolderById(id);
}

// payload: { name, type, data } where data is a data-URL string
// ("data:<mime>;base64,<...>"). Returns a direct-content URL — NOT
// file.getUrl()'s "view" page, which does not work as an <img src> — using
// the uc?export=view form so it renders inline for images and embeds.
function uploadDataUrlToDrive_(payload, filenamePrefix) {
  const match = /^data:([^;]+);base64,(.+)$/.exec(payload.data);
  if (!match) throw new Error('uploadDataUrlToDrive_: payload.data is not a base64 data URL.');
  const mimeType = match[1];
  const base64 = match[2];
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, filenamePrefix + '_' + (payload.name || 'file'));

  const folder = getAttachmentsFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/DriveHelper.gs
git commit -m "$(cat <<'EOF'
feat(gas): add DriveHelper.gs attachment upload helper

Uploads a base64 data-URL (signatures, expense attachments) to the
Drive attachments folder and returns a directly-embeddable URL, since
Sheets cells can't hold long base64 strings.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 4: `gas/Months.gs`

**Files:**
- Create: `gas/Months.gs`

- [ ] **Step 1: Write the file**

```javascript
// Months.gs

function getMonths() {
  const sheet = getSheet_(SHEET_NAMES.MONTHS);
  const data = sheetToObjects_(sheet);
  return data.rows.map(function (r) {
    return { key: r.key, id: r.id, selected: r.selected === true || r.selected === 'TRUE' };
  });
}

// Appends a new month: adds a row to Months AND a matching column to
// Members (inserted right before "sum"), defaulted to 0 for every member.
function addMonth(payload) {
  const monthsSheet = getSheet_(SHEET_NAMES.MONTHS);
  const monthsData = sheetToObjects_(monthsSheet);
  if (monthsData.rows.some(function (r) { return r.key === payload.key; })) {
    throw new Error('الشهر "' + payload.key + '" موجود مسبقاً.');
  }
  monthsSheet.appendRow([payload.key, payload.id, true]);

  const membersSheet = getSheet_(SHEET_NAMES.MEMBERS);
  const headers = membersSheet.getRange(1, 1, 1, membersSheet.getLastColumn()).getValues()[0];
  const sumCol = headerIndex_(headers, 'sum') + 1;
  membersSheet.insertColumnBefore(sumCol);
  membersSheet.getRange(1, sumCol).setValue(payload.key);
  const lastRow = membersSheet.getLastRow();
  if (lastRow > 1) {
    membersSheet.getRange(2, sumCol, lastRow - 1, 1).setValue(0);
  }

  return { months: getMonths(), members: getMembers() };
}

// Removes a month: deletes its Members column (recomputing every member's
// sum, since the column shift would otherwise misalign old cached sums) and
// its Months row.
function deleteMonth(payload) {
  const membersSheet = getSheet_(SHEET_NAMES.MEMBERS);
  const headers = membersSheet.getRange(1, 1, 1, membersSheet.getLastColumn()).getValues()[0];
  const colIdx = headers.indexOf(payload.key);
  if (colIdx !== -1) {
    membersSheet.deleteColumn(colIdx + 1);
    recalculateAllMemberSums_();
  }

  const monthsSheet = getSheet_(SHEET_NAMES.MONTHS);
  const data = sheetToObjects_(monthsSheet);
  const row = data.rows.find(function (r) { return r.key === payload.key; });
  if (row) monthsSheet.deleteRow(row._row);

  const familiesSheet = getSheet_(SHEET_NAMES.FAMILIES);
  sheetToObjects_(familiesSheet).rows.forEach(function (r) { recalculateFamily_(r.headName); });

  return { months: getMonths(), members: getMembers() };
}

function updateMonth(payload) {
  const monthsSheet = getSheet_(SHEET_NAMES.MONTHS);
  const data = sheetToObjects_(monthsSheet);
  const row = data.rows.find(function (r) { return r.key === payload.oldKey; });
  if (!row) throw new Error('الشهر غير موجود.');

  if (payload.newKey && payload.newKey !== payload.oldKey) {
    const membersSheet = getSheet_(SHEET_NAMES.MEMBERS);
    const headers = membersSheet.getRange(1, 1, 1, membersSheet.getLastColumn()).getValues()[0];
    const colIdx = headers.indexOf(payload.oldKey);
    if (colIdx !== -1) membersSheet.getRange(1, colIdx + 1).setValue(payload.newKey);
    monthsSheet.getRange(row._row, 1).setValue(payload.newKey);
  }
  if (payload.id !== undefined) monthsSheet.getRange(row._row, 2).setValue(payload.id);

  return { months: getMonths(), members: getMembers() };
}

function setSelectedMonths(payload) {
  const monthsSheet = getSheet_(SHEET_NAMES.MONTHS);
  const data = sheetToObjects_(monthsSheet);
  const selectedSet = {};
  (payload.selected || []).forEach(function (k) { selectedSet[k] = true; });
  data.rows.forEach(function (r) {
    monthsSheet.getRange(r._row, 3).setValue(!!selectedSet[r.key]);
  });
  return getMonths();
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Months.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Months.gs CRUD handlers

Add/update/delete a month, keeping the Members sheet's dynamic month
columns in sync (insert/rename/delete column, recompute sums), plus
setSelectedMonths for the ledger's month filter.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 5: `gas/Families.gs`

**Files:**
- Create: `gas/Families.gs`

- [ ] **Step 1: Write the file**

```javascript
// Families.gs

function getFamilies() {
  const sheet = getSheet_(SHEET_NAMES.FAMILIES);
  const data = sheetToObjects_(sheet);
  return data.rows.map(function (r) {
    return {
      familyId: r.familyId,
      headName: r.headName,
      memberCount: Number(r.memberCount) || 0,
      subscription: Number(r.subscription) || 0,
      totalPaid: Number(r.totalPaid) || 0,
      membersList: r.membersList
    };
  });
}

function addFamily(payload) {
  const sheet = getSheet_(SHEET_NAMES.FAMILIES);
  const data = sheetToObjects_(sheet);
  const familyId = nextId_(data.rows, 'familyId');
  appendRowFromObject_(sheet, data.headers, {
    familyId: familyId, headName: payload.headName, memberCount: 0, subscription: 0, totalPaid: 0, membersList: ''
  });
  return getFamilies();
}

function updateFamily(payload) {
  const sheet = getSheet_(SHEET_NAMES.FAMILIES);
  const data = sheetToObjects_(sheet);
  const existing = data.rows.find(function (r) { return String(r.familyId) === String(payload.familyId); });
  if (!existing) throw new Error('العائلة غير موجودة.');
  const updated = Object.assign({}, existing, {
    headName: payload.headName !== undefined ? payload.headName : existing.headName
  });
  delete updated._row;
  writeRowFromObject_(sheet, data.headers, existing._row, updated);
  return getFamilies();
}

function deleteFamily(payload) {
  const sheet = getSheet_(SHEET_NAMES.FAMILIES);
  const data = sheetToObjects_(sheet);
  const existing = data.rows.find(function (r) { return String(r.familyId) === String(payload.familyId); });
  if (!existing) throw new Error('العائلة غير موجودة.');
  sheet.deleteRow(existing._row);
  return getFamilies();
}

// Creates a Families row for `headName` if one doesn't already exist yet —
// mirrors the old client-side auto-create-family-on-new-member behavior.
function ensureFamilyExists_(headName) {
  if (!headName) return;
  const sheet = getSheet_(SHEET_NAMES.FAMILIES);
  const data = sheetToObjects_(sheet);
  const normTarget = normalizeArabicName_(headName);
  if (data.rows.some(function (r) { return normalizeArabicName_(r.headName) === normTarget; })) return;
  const familyId = nextId_(data.rows, 'familyId');
  appendRowFromObject_(sheet, data.headers, {
    familyId: familyId, headName: headName, memberCount: 0, subscription: 0, totalPaid: 0, membersList: ''
  });
}

// Recomputes memberCount/subscription/totalPaid/membersList for the family
// identified by `headName` (matched via normalizeArabicName_, same as the
// original client code), from the current Members sheet. Deletes the family
// row entirely if it ends up with zero members — matching the original
// app's "if family has no more individuals, delete family completely".
// Called after any member add/edit/delete/payment change.
function recalculateFamily_(headName) {
  if (!headName) return;
  const familiesSheet = getSheet_(SHEET_NAMES.FAMILIES);
  const familiesData = sheetToObjects_(familiesSheet);
  const normTarget = normalizeArabicName_(headName);
  const familyRow = familiesData.rows.find(function (r) { return normalizeArabicName_(r.headName) === normTarget; });
  if (!familyRow) return;

  const membersSheet = getSheet_(SHEET_NAMES.MEMBERS);
  const membersData = sheetToObjects_(membersSheet);
  const familyMembers = membersData.rows.filter(function (r) { return normalizeArabicName_(r.parent) === normTarget; });

  const memberCount = familyMembers.length;
  if (memberCount === 0) {
    familiesSheet.deleteRow(familyRow._row);
    return;
  }

  const totalPaid = familyMembers.reduce(function (s, m) { return s + (Number(m.sum) || 0); }, 0);
  const membersList = familyMembers.map(function (m) { return m.name; }).join('، ');
  const subscription = memberCount * 10;

  familiesSheet.getRange(familyRow._row, headerIndex_(familiesData.headers, 'memberCount') + 1).setValue(memberCount);
  familiesSheet.getRange(familyRow._row, headerIndex_(familiesData.headers, 'subscription') + 1).setValue(subscription);
  familiesSheet.getRange(familyRow._row, headerIndex_(familiesData.headers, 'totalPaid') + 1).setValue(totalPaid);
  familiesSheet.getRange(familyRow._row, headerIndex_(familiesData.headers, 'membersList') + 1).setValue(membersList);
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Families.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Families.gs CRUD + recalculation

recalculateFamily_ is the single source of truth for memberCount,
subscription (memberCount * 10, matching the app's per-member flat fee
invariant), totalPaid, and membersList — replacing the scattered
increment/decrement logic the old client code had at every mutation
site. Auto-deletes a family once its last member is removed, matching
prior behavior.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 6: `gas/Members.gs`

**Files:**
- Create: `gas/Members.gs`

- [ ] **Step 1: Write the file**

```javascript
// Members.gs

function monthColumns_(headers) {
  return headers.filter(function (h) { return FIXED_MEMBER_COLUMNS.indexOf(h) === -1; });
}

function memberRowToObject_(headers, row) {
  const months = monthColumns_(headers);
  const payments = {};
  months.forEach(function (m) { payments[m] = Number(row[m]) || 0; });
  return { id: row.id, name: row.name, parent: row.parent, payments: payments, sum: Number(row.sum) || 0 };
}

function getMembers() {
  const sheet = getSheet_(SHEET_NAMES.MEMBERS);
  const data = sheetToObjects_(sheet);
  return data.rows.map(function (r) { return memberRowToObject_(data.headers, r); });
}

function addMember(payload) {
  const sheet = getSheet_(SHEET_NAMES.MEMBERS);
  const data = sheetToObjects_(sheet);
  const id = nextId_(data.rows, 'id');
  const months = monthColumns_(data.headers);
  const obj = { id: id, name: payload.name, parent: payload.parent, sum: 0 };
  months.forEach(function (m) { obj[m] = 0; });
  appendRowFromObject_(sheet, data.headers, obj);

  ensureFamilyExists_(payload.parent);
  recalculateFamily_(payload.parent);
  return { members: getMembers(), families: getFamilies() };
}

function updateMember(payload) {
  const sheet = getSheet_(SHEET_NAMES.MEMBERS);
  const data = sheetToObjects_(sheet);
  const existing = findRowById_(data.rows, payload.id);
  if (!existing) throw new Error('العضو غير موجود.');

  const oldParent = existing.parent;
  const updated = Object.assign({}, existing, { name: payload.name, parent: payload.parent });
  delete updated._row;
  writeRowFromObject_(sheet, data.headers, existing._row, updated);

  ensureFamilyExists_(payload.parent);
  recalculateFamily_(oldParent);
  if (normalizeArabicName_(payload.parent) !== normalizeArabicName_(oldParent)) recalculateFamily_(payload.parent);
  return { members: getMembers(), families: getFamilies() };
}

function deleteMember(payload) {
  const sheet = getSheet_(SHEET_NAMES.MEMBERS);
  const data = sheetToObjects_(sheet);
  const existing = findRowById_(data.rows, payload.id);
  if (!existing) throw new Error('العضو غير موجود.');

  sheet.deleteRow(existing._row);
  recalculateFamily_(existing.parent);
  return { members: getMembers(), families: getFamilies() };
}

// Sets one member's payment for one month, recomputes that member's `sum`,
// and recalculates the owning family's totals. Used by the ledger's
// per-cell payment checkbox.
function setMemberPayment(payload) {
  const sheet = getSheet_(SHEET_NAMES.MEMBERS);
  const data = sheetToObjects_(sheet);
  const existing = findRowById_(data.rows, payload.id);
  if (!existing) throw new Error('العضو غير موجود.');

  const colIdx = headerIndex_(data.headers, payload.month);
  sheet.getRange(existing._row, colIdx + 1).setValue(payload.amount);

  const months = monthColumns_(data.headers);
  let sum = 0;
  months.forEach(function (m) {
    sum += (m === payload.month) ? Number(payload.amount) || 0 : Number(existing[m]) || 0;
  });
  sheet.getRange(existing._row, headerIndex_(data.headers, 'sum') + 1).setValue(sum);

  recalculateFamily_(existing.parent);
  return { members: getMembers(), families: getFamilies() };
}

// Sets a flat 10-shekel payment for every given month, for every member
// belonging to `headName` (matches either as a member's `parent`, or by
// name for a household head with no separate member row) — mirrors the
// receipt workspace's "record payment" bulk action.
function setFamilyPayments(payload) {
  const sheet = getSheet_(SHEET_NAMES.MEMBERS);
  const data = sheetToObjects_(sheet);
  const months = monthColumns_(data.headers);
  const sumCol = headerIndex_(data.headers, 'sum') + 1;
  const normTarget = normalizeArabicName_(payload.headName);
  const targetMonths = payload.months || [];

  data.rows.forEach(function (r) {
    const matches = normalizeArabicName_(r.parent) === normTarget || normalizeArabicName_(r.name) === normTarget;
    if (!matches) return;

    targetMonths.forEach(function (monthKey) {
      const colIdx = data.headers.indexOf(monthKey);
      if (colIdx !== -1) sheet.getRange(r._row, colIdx + 1).setValue(10);
    });

    let sum = 0;
    months.forEach(function (m) {
      sum += targetMonths.indexOf(m) !== -1 ? 10 : (Number(r[m]) || 0);
    });
    sheet.getRange(r._row, sumCol).setValue(sum);
  });

  recalculateFamily_(payload.headName);
  return { members: getMembers(), families: getFamilies() };
}

// Recomputes every member's `sum` from scratch. Used after a month column
// is removed (deleteMonth), since removing a column shifts values.
function recalculateAllMemberSums_() {
  const sheet = getSheet_(SHEET_NAMES.MEMBERS);
  const data = sheetToObjects_(sheet);
  const months = monthColumns_(data.headers);
  const sumCol = headerIndex_(data.headers, 'sum') + 1;
  data.rows.forEach(function (r) {
    let sum = 0;
    months.forEach(function (m) { sum += Number(r[m]) || 0; });
    sheet.getRange(r._row, sumCol).setValue(sum);
  });
}

// Full resync: recomputes every member's sum AND every family's totals from
// scratch. Backs the "recalculate everything" admin button — a safety net,
// since every mutation already keeps sums/totals correct incrementally.
function recalculateEverything() {
  recalculateAllMemberSums_();
  const familiesSheet = getSheet_(SHEET_NAMES.FAMILIES);
  sheetToObjects_(familiesSheet).rows.forEach(function (r) { recalculateFamily_(r.headName); });
  return { members: getMembers(), families: getFamilies() };
}

// Replaces all Members/Families/Months data from a parsed Excel import,
// matching the previous client-side processWorkbook()'s full-replace
// behavior. payload: { members: [{id,name,parent,payments,sum}], families:
// [{familyId,headName,memberCount,subscription,totalPaid,membersList}],
// months: [{key,id}] }.
function bulkImportMembers(payload) {
  const monthsSheet = getSheet_(SHEET_NAMES.MONTHS);
  const oldMonthsLastRow = monthsSheet.getLastRow();
  if (oldMonthsLastRow > 1) monthsSheet.getRange(2, 1, oldMonthsLastRow - 1, 3).clearContent();
  (payload.months || []).forEach(function (m, i) {
    monthsSheet.getRange(2 + i, 1, 1, 3).setValues([[m.key, m.id, true]]);
  });

  const membersSheet = getSheet_(SHEET_NAMES.MEMBERS);
  const monthKeys = (payload.months || []).map(function (m) { return m.key; });
  const newHeaders = ['id', 'name', 'parent'].concat(monthKeys).concat(['sum']);
  membersSheet.clear();
  membersSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  membersSheet.setFrozenRows(1);

  const memberRows = (payload.members || []).map(function (m) {
    const row = [m.id, m.name, m.parent];
    monthKeys.forEach(function (k) { row.push(Number((m.payments || {})[k]) || 0); });
    row.push(Number(m.sum) || 0);
    return row;
  });
  if (memberRows.length) {
    membersSheet.getRange(2, 1, memberRows.length, newHeaders.length).setValues(memberRows);
  }

  const familiesSheet = getSheet_(SHEET_NAMES.FAMILIES);
  const oldFamiliesLastRow = familiesSheet.getLastRow();
  if (oldFamiliesLastRow > 1) familiesSheet.getRange(2, 1, oldFamiliesLastRow - 1, 6).clearContent();
  const familyRows = (payload.families || []).map(function (f) {
    return [f.familyId, f.headName, f.memberCount, f.subscription, f.totalPaid, f.membersList];
  });
  if (familyRows.length) {
    familiesSheet.getRange(2, 1, familyRows.length, 6).setValues(familyRows);
  }

  return { members: getMembers(), families: getFamilies(), months: getMonths() };
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Members.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Members.gs CRUD, payments, and bulk import

Covers per-member CRUD, single-cell payment updates (ledger checkbox),
bulk family-wide payment recording (receipt workspace), the
recalculate-everything safety net, and the full-replace bulk Excel
import path.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 7: `gas/Expenses.gs`

**Files:**
- Create: `gas/Expenses.gs`

- [ ] **Step 1: Write the file**

```javascript
// Expenses.gs

function getExpenses() {
  const sheet = getSheet_(SHEET_NAMES.EXPENSES);
  const data = sheetToObjects_(sheet);
  return data.rows.map(function (r) {
    return {
      id: r.id, date: r.date, amount: Number(r.amount) || 0, reason: r.reason,
      category: r.category, authorized: r.authorized,
      attachment: r.attachmentUrl ? { name: r.attachmentName, type: r.attachmentType, url: r.attachmentUrl } : null
    };
  });
}

// payload: { date, amount, reason, category, authorized, attachment:
// {name, type, data} | null } — attachment.data, if present, is a data-URL
// string uploaded to Drive first.
function addExpense(payload) {
  const sheet = getSheet_(SHEET_NAMES.EXPENSES);
  const data = sheetToObjects_(sheet);
  const id = 'exp_' + Date.now();

  let attachmentUrl = '', attachmentName = '', attachmentType = '';
  if (payload.attachment && payload.attachment.data) {
    attachmentUrl = uploadDataUrlToDrive_(payload.attachment, 'expense');
    attachmentName = payload.attachment.name;
    attachmentType = payload.attachment.type;
  }

  appendRowFromObject_(sheet, data.headers, {
    id: id, date: payload.date, amount: payload.amount, reason: payload.reason, category: payload.category,
    authorized: payload.authorized, attachmentName: attachmentName, attachmentType: attachmentType, attachmentUrl: attachmentUrl
  });
  return getExpenses();
}

function deleteExpense(payload) {
  const sheet = getSheet_(SHEET_NAMES.EXPENSES);
  const data = sheetToObjects_(sheet);
  const existing = findRowById_(data.rows, payload.id);
  if (!existing) throw new Error('أمر الصرف غير موجود.');
  sheet.deleteRow(existing._row);
  return getExpenses();
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Expenses.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Expenses.gs CRUD with Drive attachment upload

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 8: `gas/Users.gs`

**Files:**
- Create: `gas/Users.gs`

- [ ] **Step 1: Write the file**

```javascript
// Users.gs

function sanitizeUser_(r) {
  return {
    key: r.key, name: r.name, role: r.role, email: r.email,
    isBuiltIn: r.isBuiltIn === true || r.isBuiltIn === 'TRUE',
    isMaster: r.isMaster === true || r.isMaster === 'TRUE',
    isActive: r.isActive === true || r.isActive === 'TRUE',
    firstLoginDone: r.firstLoginDone === true || r.firstLoginDone === 'TRUE'
  };
}

function getUsers() {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  return sheetToObjects_(sheet).rows.map(sanitizeUser_);
}

// payload: { email, pass }. Returns { user } on success or { error } on any
// expected failure (wrong password, inactive account) — never throws for
// bad credentials, only for unexpected sheet errors.
function login(payload) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const data = sheetToObjects_(sheet);
  const email = String(payload.email || '').trim().toLowerCase();
  const row = data.rows.find(function (r) { return String(r.email).toLowerCase() === email; });

  if (!row) return { error: 'لا يوجد حساب بهذا البريد. تحقق من أنك مسجّل أو أرسل طلب للتسجيل.' };
  if (String(row.pass) !== String(payload.pass)) return { error: 'كلمة المرور غير صحيحة. يرجى المحاولة مجدداً.' };
  if (row.isActive === false || row.isActive === 'FALSE') return { error: 'حسابك معطّل حالياً. تواصل مع المسؤول لفعل حسابك.' };

  return { user: sanitizeUser_(row) };
}

// payload: { key, newPass }. Used for the forced first-login password
// change dialog.
function changePassword(payload) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const data = sheetToObjects_(sheet);
  const row = data.rows.find(function (r) { return r.key === payload.key; });
  if (!row) throw new Error('المستخدم غير موجود.');
  sheet.getRange(row._row, headerIndex_(data.headers, 'pass') + 1).setValue(payload.newPass);
  sheet.getRange(row._row, headerIndex_(data.headers, 'firstLoginDone') + 1).setValue(true);
  return { user: sanitizeUser_(row) };
}

// payload: { email }. Resets to the default password and clears
// firstLoginDone, matching the login screen's "forgot password" link.
function resetPassword(payload) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const data = sheetToObjects_(sheet);
  const emailLower = String(payload.email || '').trim().toLowerCase();
  const row = data.rows.find(function (r) { return String(r.email).toLowerCase() === emailLower; });
  if (!row) throw new Error('لا يوجد حساب مشرف مالي مسجل بهذا البريد الإلكتروني.');
  sheet.getRange(row._row, headerIndex_(data.headers, 'pass') + 1).setValue('ABC12345');
  sheet.getRange(row._row, headerIndex_(data.headers, 'firstLoginDone') + 1).setValue(false);
  return { user: sanitizeUser_(row) };
}

// Adds a user account directly (admin "add supervisor" screen — distinct
// from the public registration-request flow in PendingUsers.gs).
function addSupervisor(payload) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const data = sheetToObjects_(sheet);
  const emailLower = String(payload.email || '').trim().toLowerCase();
  if (data.rows.some(function (r) { return String(r.email).toLowerCase() === emailLower; })) {
    throw new Error('هذا البريد الإلكتروني مسجل بالفعل لمشرف آخر.');
  }
  appendRowFromObject_(sheet, data.headers, {
    key: 'user_' + Date.now(), name: payload.name, role: payload.role || 'مستخدم', email: payload.email,
    pass: payload.pass, isBuiltIn: false, isMaster: false, isActive: true, firstLoginDone: true
  });
  return getUsers();
}

function updateUser(payload) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const data = sheetToObjects_(sheet);
  const row = data.rows.find(function (r) { return r.key === payload.key; });
  if (!row) throw new Error('المستخدم غير موجود.');
  const updated = Object.assign({}, row, {
    name: payload.name !== undefined ? payload.name : row.name,
    role: payload.role !== undefined ? payload.role : row.role,
    email: payload.email !== undefined ? payload.email : row.email
  });
  delete updated._row;
  writeRowFromObject_(sheet, data.headers, row._row, updated);
  return getUsers();
}

function setUserActive(payload) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const data = sheetToObjects_(sheet);
  const row = data.rows.find(function (r) { return r.key === payload.key; });
  if (!row) throw new Error('المستخدم غير موجود.');
  if ((row.isMaster === true || row.isMaster === 'TRUE') && !payload.active) {
    throw new Error('لا يمكن تعطيل الحساب الرئيسي.');
  }
  sheet.getRange(row._row, headerIndex_(data.headers, 'isActive') + 1).setValue(!!payload.active);
  return getUsers();
}

function deleteUser(payload) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const data = sheetToObjects_(sheet);
  const row = data.rows.find(function (r) { return r.key === payload.key; });
  if (!row) throw new Error('المستخدم غير موجود.');
  if (row.isMaster === true || row.isMaster === 'TRUE') throw new Error('لا يمكن حذف الحساب الرئيسي.');
  sheet.deleteRow(row._row);
  return getUsers();
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Users.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Users.gs auth and account management

login/changePassword/resetPassword back the app's existing login,
forced-first-login, and forgot-password flows; addSupervisor/
updateUser/setUserActive/deleteUser back the admin user-management
screen, guarding the master account from deactivation/deletion.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 9: `gas/PendingUsers.gs`

**Files:**
- Create: `gas/PendingUsers.gs`

- [ ] **Step 1: Write the file**

```javascript
// PendingUsers.gs

function getPendingUsers() {
  const sheet = getSheet_(SHEET_NAMES.PENDING_USERS);
  return sheetToObjects_(sheet).rows.map(function (r) {
    return { id: r.id, firstName: r.firstName, lastName: r.lastName, email: r.email, requestedAt: r.requestedAt };
  });
}

// payload: { firstName, lastName, email, pass }
function requestRegistration(payload) {
  const usersSheet = getSheet_(SHEET_NAMES.USERS);
  const emailLower = String(payload.email || '').trim().toLowerCase();
  if (sheetToObjects_(usersSheet).rows.some(function (r) { return String(r.email).toLowerCase() === emailLower; })) {
    throw new Error('هذا البريد الإلكتروني مسجل بالفعل لمشرف آخر.');
  }

  const sheet = getSheet_(SHEET_NAMES.PENDING_USERS);
  const data = sheetToObjects_(sheet);
  if (data.rows.some(function (r) { return String(r.email).toLowerCase() === emailLower; })) {
    throw new Error('يوجد بالفعل طلب تسجيل معلق لهذا البريد الإلكتروني.');
  }

  appendRowFromObject_(sheet, data.headers, {
    id: 'pend_' + Date.now(), firstName: payload.firstName, lastName: payload.lastName,
    email: payload.email, pass: payload.pass, requestedAt: new Date().toISOString()
  });
  return getPendingUsers();
}

function approveUser(payload) {
  const pendingSheet = getSheet_(SHEET_NAMES.PENDING_USERS);
  const pendingData = sheetToObjects_(pendingSheet);
  const row = pendingData.rows.find(function (r) { return r.id === payload.id; });
  if (!row) throw new Error('الطلب غير موجود.');

  const usersSheet = getSheet_(SHEET_NAMES.USERS);
  const usersData = sheetToObjects_(usersSheet);
  appendRowFromObject_(usersSheet, usersData.headers, {
    key: 'user_' + Date.now(), name: (row.firstName + ' ' + row.lastName).trim(), role: 'مشرف مالي',
    email: row.email, pass: row.pass, isBuiltIn: false, isMaster: false, isActive: true, firstLoginDone: true
  });

  pendingSheet.deleteRow(row._row);
  return { users: getUsers(), pendingUsers: getPendingUsers() };
}

function rejectUser(payload) {
  const sheet = getSheet_(SHEET_NAMES.PENDING_USERS);
  const data = sheetToObjects_(sheet);
  const row = data.rows.find(function (r) { return r.id === payload.id; });
  if (!row) throw new Error('الطلب غير موجود.');
  sheet.deleteRow(row._row);
  return getPendingUsers();
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/PendingUsers.gs
git commit -m "$(cat <<'EOF'
feat(gas): add PendingUsers.gs registration request flow

requestRegistration/approveUser/rejectUser back the public
registration form and the admin approve/reject screen. Rows are
addressed by a generated id rather than email, unlike the old
localStorage array.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 10: `gas/Signatures.gs`

**Files:**
- Create: `gas/Signatures.gs`

- [ ] **Step 1: Write the file**

```javascript
// Signatures.gs

// payload: { month }. Returns { "<slotIndex>": signatureUrl, ... }.
function getSignatures(payload) {
  const sheet = getSheet_(SHEET_NAMES.SIGNATURES);
  const rows = sheetToObjects_(sheet).rows.filter(function (r) { return r.month === payload.month; });
  const result = {};
  rows.forEach(function (r) { result[r.slotIndex] = r.signatureUrl; });
  return result;
}

// payload: { month, slotIndex, data } — data is the signature canvas's
// toDataURL() PNG string.
function saveSignature(payload) {
  const url = uploadDataUrlToDrive_(
    { name: 'signature.png', type: 'image/png', data: payload.data },
    'sig_' + payload.month + '_' + payload.slotIndex
  );

  const sheet = getSheet_(SHEET_NAMES.SIGNATURES);
  const data = sheetToObjects_(sheet);
  const existing = data.rows.find(function (r) { return r.month === payload.month && String(r.slotIndex) === String(payload.slotIndex); });

  if (existing) {
    writeRowFromObject_(sheet, data.headers, existing._row,
      { month: payload.month, slotIndex: payload.slotIndex, signatureUrl: url, updatedAt: new Date().toISOString() });
  } else {
    appendRowFromObject_(sheet, data.headers,
      { month: payload.month, slotIndex: payload.slotIndex, signatureUrl: url, updatedAt: new Date().toISOString() });
  }
  return getSignatures({ month: payload.month });
}

function clearSignature(payload) {
  const sheet = getSheet_(SHEET_NAMES.SIGNATURES);
  const data = sheetToObjects_(sheet);
  const existing = data.rows.find(function (r) { return r.month === payload.month && String(r.slotIndex) === String(payload.slotIndex); });
  if (existing) sheet.deleteRow(existing._row);
  return getSignatures({ month: payload.month });
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Signatures.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Signatures.gs per-month signature slots

Fetched on demand per report month rather than bundled into
getAllData, matching how the frontend already only displays one
month's signatures at a time.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 11: `gas/Settings.gs`

**Files:**
- Create: `gas/Settings.gs`

- [ ] **Step 1: Write the file**

```javascript
// Settings.gs

function getSettings() {
  const sheet = getSheet_(SHEET_NAMES.SETTINGS);
  const result = {};
  sheetToObjects_(sheet).rows.forEach(function (r) { result[r.key] = r.value; });
  return result;
}

function setSetting(payload) {
  const sheet = getSheet_(SHEET_NAMES.SETTINGS);
  const data = sheetToObjects_(sheet);
  const existing = data.rows.find(function (r) { return r.key === payload.key; });
  if (existing) {
    sheet.getRange(existing._row, headerIndex_(data.headers, 'value') + 1).setValue(payload.value);
  } else {
    appendRowFromObject_(sheet, data.headers, { key: payload.key, value: payload.value });
  }
  return getSettings();
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Settings.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Settings.gs generic key-value store

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 12: `gas/Bootstrap.gs`

**Files:**
- Create: `gas/Bootstrap.gs`

- [ ] **Step 1: Write the file**

```javascript
// Bootstrap.gs
// One aggregating call used once on app load, so the frontend doesn't have
// to make 7 separate round trips. Signatures are excluded — fetched
// per-month on demand instead (see Signatures.gs).

function getAllData() {
  return {
    members: getMembers(),
    families: getFamilies(),
    months: getMonths(),
    expenses: getExpenses(),
    users: getUsers(),
    pendingUsers: getPendingUsers(),
    settings: getSettings()
  };
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Bootstrap.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Bootstrap.gs getAllData aggregator

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 13: `gas/Api.gs` — the HTTP entry point

**Files:**
- Create: `gas/Api.gs`

- [ ] **Step 1: Write the file**

```javascript
// Api.gs
// Single HTTP entry point. Deploy this project as a Web App (Deploy > New
// deployment > Web app; Execute as: Me; Who has access: Anyone) and paste
// the resulting URL into API_URL at the top of app.js. This must be the
// LAST .gs file pasted in, since it references every handler function.

const ACTIONS = {
  getAllData: getAllData,

  getMembers: getMembers,
  addMember: addMember,
  updateMember: updateMember,
  deleteMember: deleteMember,
  setMemberPayment: setMemberPayment,
  setFamilyPayments: setFamilyPayments,
  recalculateEverything: recalculateEverything,
  bulkImportMembers: bulkImportMembers,

  getFamilies: getFamilies,
  addFamily: addFamily,
  updateFamily: updateFamily,
  deleteFamily: deleteFamily,

  getMonths: getMonths,
  addMonth: addMonth,
  updateMonth: updateMonth,
  deleteMonth: deleteMonth,
  setSelectedMonths: setSelectedMonths,

  getExpenses: getExpenses,
  addExpense: addExpense,
  deleteExpense: deleteExpense,

  login: login,
  changePassword: changePassword,
  resetPassword: resetPassword,
  getUsers: getUsers,
  addSupervisor: addSupervisor,
  updateUser: updateUser,
  setUserActive: setUserActive,
  deleteUser: deleteUser,

  getPendingUsers: getPendingUsers,
  requestRegistration: requestRegistration,
  approveUser: approveUser,
  rejectUser: rejectUser,

  getSignatures: getSignatures,
  saveSignature: saveSignature,
  clearSignature: clearSignature,

  getSettings: getSettings,
  setSetting: setSetting
};

// Actions that only read data — skip the script lock for these.
const READ_ONLY_ACTIONS = {
  getAllData: true, getMembers: true, getFamilies: true, getMonths: true,
  getExpenses: true, getUsers: true, getPendingUsers: true,
  getSignatures: true, getSettings: true, login: true
};

function doPost(e) {
  let action = '';
  try {
    const body = JSON.parse(e.postData.contents);
    action = body.action;
    const payload = body.payload || {};
    const handler = ACTIONS[action];
    if (!handler) return jsonResponse_({ error: 'إجراء غير معروف: ' + action });

    if (READ_ONLY_ACTIONS[action]) {
      return jsonResponse_({ result: handler(payload) });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return jsonResponse_({ result: handler(payload) });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse_({ error: (action ? '[' + action + '] ' : '') + err.message });
  }
}

function doGet(e) {
  return jsonResponse_({ status: 'ok', message: 'CEMS-ATFIHAH API is running.' });
}
```

- [ ] **Step 2: Commit**

```bash
git add gas/Api.gs
git commit -m "$(cat <<'EOF'
feat(gas): add Api.gs doPost/doGet dispatcher

Routes {action, payload} JSON bodies to the matching handler across
every entity file, serializing mutating actions through
LockService so concurrent requests can't corrupt a row. doGet is a
trivial health check for visiting the deployment URL directly.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Backend Task 14: `gas/DEPLOYMENT.md` — manual deployment guide

**Files:**
- Create: `gas/DEPLOYMENT.md`

- [ ] **Step 1: Write the file**

```markdown
# Deploying the Google Sheets backend

These steps are manual — done once by whoever owns the Google account this
runs under (the family committee). An AI agent cannot perform them; they
require the Google Sheets/Apps Script UI and real Google credentials.

## 5. What "empty Members/Families" means on first run

`setupSpreadsheet()` deliberately does not seed a member roster (see the
note at the end of Backend Task 2). To load the real ~134-member roster:
open the app, log in, go to "بوابة تحميل وتخزين قاعدة البيانات" (the Excel
upload tab) and upload the same Excel file the app used before — this now
pushes the parsed data into the Sheet via `bulkImportMembers` instead of
`localStorage`.

## 1. Create the Sheet and paste in the code

1. Go to https://sheets.google.com and create a new blank spreadsheet.
   Rename it (e.g. "صندوق تحصيل عائلة آل اطفيحة — قاعدة البيانات").
2. Extensions → Apps Script. This opens a bound Apps Script project.
3. Delete the default empty `Code.gs` file's contents (or delete the file).
4. For each file in this repo's `gas/` folder — in this exact order —
   create a matching file in the Apps Script editor (File → New → Script
   file, name it without the `.gs` extension, e.g. `Util`) and paste the
   contents in:
   1. `Util.gs`
   2. `Setup.gs`
   3. `DriveHelper.gs`
   4. `Months.gs`
   5. `Families.gs`
   6. `Members.gs`
   7. `Expenses.gs`
   8. `Users.gs`
   9. `PendingUsers.gs`
   10. `Signatures.gs`
   11. `Settings.gs`
   12. `Bootstrap.gs`
   13. `Api.gs`
5. Save the project (Ctrl+S / Cmd+S).

## 2. Run the one-time setup

1. In the function dropdown (top toolbar, next to Run/Debug), select
   `setupSpreadsheet`.
2. Click **Run**. The first run will prompt for authorization — review the
   permissions (it needs access to this Sheet and to create a Drive
   folder) and allow them.
3. Check the execution log (View → Logs, or Ctrl+Enter) — it should end
   with `setupSpreadsheet() finished successfully.`
4. Switch back to the Sheet tab: confirm 8 new tabs now exist — Members,
   Families, Months, Expenses, Users, PendingUsers, Signatures, Settings —
   each with a header row. Months should have 12 rows and Users should
   have 5 rows (the committee accounts); the rest start empty.
5. Check Google Drive for a new folder named "CEMS-ATFIHAH Attachments".

Re-running `setupSpreadsheet` later (e.g. after adding a new `.gs` file) is
always safe — it never overwrites existing sheet data.

## 3. Deploy as a Web App

1. In the Apps Script editor: Deploy → New deployment.
2. Click the gear icon next to "Select type" → Web app.
3. Description: anything (e.g. "CEMS API v1").
4. Execute as: **Me** (your account).
5. Who has access: **Anyone**.
6. Click Deploy. Authorize again if prompted.
7. Copy the **Web app URL** shown (ends in `/exec`).

## 4. Wire it into the app

1. Open `app.js` in this repo, find the `API_URL` constant near the top.
2. Paste the Web app URL in as its value.
3. Serve the app (`npx serve .`) and open it in a browser — the login
   screen should appear, and logging in with one of the 5 default accounts
   (email from `gas/Setup.gs`'s `DEFAULT_USERS_`, password `ABC12345`)
   should work end-to-end against the live Sheet.

## Redeploying after a code change

Apps Script Web App URLs are pinned to a specific deployment. After
editing any `.gs` file:
1. Deploy → Manage deployments.
2. Click the pencil/edit icon on the active deployment.
3. Version: **New version**. Deploy.
4. The URL stays the same — no need to update `app.js` again.
```

- [ ] **Step 2: Commit**

```bash
git add gas/DEPLOYMENT.md
git commit -m "$(cat <<'EOF'
docs(gas): add manual deployment guide

Step-by-step instructions for the parts of this migration that can
only be done by a human with the real Google account — creating the
Sheet, pasting in the gas/*.gs files, running setupSpreadsheet, and
deploying the Web App.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 1: API client foundation in `app.js`

**Files:**
- Modify: `app.js:1-6` (top of file, before `const MASTER_ADMIN_EMAIL`)
- Modify: `app.js:115-122` (the `state` object)

This task is purely additive (no removals) so it can be verified standalone.

- [ ] **Step 1: Add `API_URL` and `callApi()` at the top of the file**

Find this at the very top of `app.js`:

```javascript
// app.js

// Admin credentials — Palestinian Family Financial Committee
// Master admin: jehadgml@gmail.com (full access, cannot be removed)
const MASTER_ADMIN_EMAIL = "jehadgml@gmail.com";
```

Replace it with:

```javascript
// app.js

// Google Apps Script Web App URL — paste the deployment URL from
// gas/DEPLOYMENT.md step 3 here. Every piece of app data is read from and
// written to the Google Sheet behind this URL; there is no local fallback.
const API_URL = "PASTE_YOUR_WEB_APP_URL_HERE";

// Calls one backend action. Uses text/plain as the Content-Type so the
// browser treats this as a CORS-simple request and skips the preflight
// OPTIONS request — Apps Script Web Apps don't answer OPTIONS.
async function callApi(action, payload) {
    let res;
    try {
        res = await fetch(API_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({ action, payload: payload || {} })
        });
    } catch (networkErr) {
        throw new Error("تعذر الاتصال بالخادم. تحقق من اتصالك بالإنترنت وحاول مجدداً. (" + networkErr.message + ")");
    }
    if (!res.ok) {
        throw new Error("خطأ من الخادم (HTTP " + res.status + "). حاول مجدداً لاحقاً.");
    }
    const body = await res.json();
    if (body.error) {
        throw new Error(body.error);
    }
    return body.result;
}
```

- [ ] **Step 2: Add the new state fields**

Find this (state object definition):

```javascript
// Application State
let state = {
    currentUser: null,
    members: [],      // from 'التحصيل الشهري'
    families: [],     // from 'وصولات العائلات'
    monthsList: JSON.parse(JSON.stringify(DEFAULT_MONTHS)),
    expenses: [],     // Expenditure / withdrawal records
    originalWorkbook: null // SheetJS raw object
};
```

Replace it with:

```javascript
// Application State — populated from the Google Sheet backend via
// loadDataFromServer() on startup; nothing here is read from localStorage.
let state = {
    currentUser: null,
    members: [],
    families: [],
    monthsList: [],
    selectedLedgerMonths: [],
    expenses: [],
    users: [],
    pendingUsers: [],
    settings: {},
    originalWorkbook: null // SheetJS raw object, session-only (never persisted)
};
```

(`DEFAULT_MONTHS` stays defined lower in the file for now — Frontend Task 2 removes it along with the other now-dead defaults.)

- [ ] **Step 3: Verify**

This step only needs a syntax check, since `API_URL` is still a placeholder and nothing calls `callApi` yet:

```bash
node --check app.js
```

Expected: no output (exit code 0) — confirms the file still parses as valid JavaScript.

- [ ] **Step 4: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: add API client foundation for Google Sheets backend

Adds API_URL + callApi(), and the new state fields (users,
pendingUsers, settings, selectedLedgerMonths) that will be populated
from the backend instead of localStorage. Purely additive — no
existing behavior changes yet.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 2: Bootstrap rewrite — remove local account storage, load from server

**Files:**
- Modify: `app.js` (several non-adjacent regions, listed per step)

This is the foundational rewiring task: after it, the app no longer has a
hardcoded `ADMINS` object or any `localStorage`-backed data loading. Every
later frontend task builds on `state.users` / `loadDataFromServer()` this
task introduces.

- [ ] **Step 1: Remove the hardcoded `ADMINS` object and its loader**

Find (right after the `API_URL`/`callApi` block added in Task 1):

```javascript
const ADMINS = {
    admin1: {
        name: "جهاد زكري إسماعيل اطفيحة",
        role: "رئيس اللجنة المالية للعشيرة",
        pass: "ABC12345",
        email: "jehadgml@gmail.com",
        firstLoginDone: false
    },
    admin2: {
        name: "أشرف يوسف محمود اطفيحة",
        role: "نائب رئيس اللجنة المالية",
        pass: "ABC12345",
        email: "ashraf.atfihah@gmail.com",
        firstLoginDone: false
    },
    admin3: {
        name: "إبراهيم محمد إبراهيم اطفيحة",
        role: "أمين صندوق العشيرة",
        pass: "ABC12345",
        email: "ibrahim.atfihah@gmail.com",
        firstLoginDone: false
    },
    admin4: {
        name: "معتز أمين محمد محمود اطفيحة",
        role: "المراقب المالي للصندوق",
        pass: "ABC12345",
        email: "moataz.atfihah@gmail.com",
        firstLoginDone: false
    },
    admin5: {
        name: "حبيب محمود محمد اطفيحة",
        role: "منسق التحصيل والوصولات",
        pass: "ABC12345",
        email: "habib.atfihah@gmail.com",
        firstLoginDone: false
    }
};

// Load any saved admin passwords/emails/states from localStorage
function loadAdminPasswords() {
    Object.keys(ADMINS).forEach(key => {
        const savedPass = localStorage.getItem(`cems_admin_pass_${key}`);
        const savedEmail = localStorage.getItem(`cems_admin_email_${key}`);
        const savedFirstLogin = localStorage.getItem(`cems_admin_firstlogin_${key}`);
        const savedActive = localStorage.getItem(`cems_admin_active_${key}`);
        if (savedPass) ADMINS[key].pass = savedPass;
        if (savedEmail) ADMINS[key].email = savedEmail;
        if (savedFirstLogin === 'done') ADMINS[key].firstLoginDone = true;
        if (savedActive === 'false') {
            ADMINS[key].isActive = false;
        } else {
            ADMINS[key].isActive = true;
        }
    });
}
loadAdminPasswords();

// Find user info crosschecking built-in ADMINS and approved dynamic users
function findUserByEmail(email) {
    if (!email) return null;
    const lowerEmail = email.trim().toLowerCase();
    
    // 1. Check built-in admins
    for (const key of Object.keys(ADMINS)) {
        if (ADMINS[key].email && ADMINS[key].email.toLowerCase() === lowerEmail) {
            return {
                key: key,
                name: ADMINS[key].name,
                role: ADMINS[key].role,
                pass: ADMINS[key].pass,
                email: ADMINS[key].email,
                firstLoginDone: ADMINS[key].firstLoginDone,
                isActive: ADMINS[key].isActive !== false,
                isBuiltIn: true
            };
        }
    }

    // 2. Check approved dynamic users in localStorage
    const dynamicUsers = JSON.parse(localStorage.getItem("cems_dynamic_users") || "[]");
    const foundUser = dynamicUsers.find(u => u.email.toLowerCase() === lowerEmail);
    if (foundUser) {
        return {
            ...foundUser,
            isBuiltIn: false
        };
    }

    return null;
}
```

Replace it with:

```javascript
// Looks up a user in the already-loaded state.users (populated from the
// Sheet by loadDataFromServer()). Used only for instant client-side
// pre-checks (e.g. "this email is already registered") — the server
// performs the authoritative check on every mutating call.
function findUserInState_(email) {
    if (!email) return null;
    const lower = email.trim().toLowerCase();
    return state.users.find(u => u.email.toLowerCase() === lower) || null;
}
```

- [ ] **Step 2: Replace `handleExcelFile` / add `arrayBufferToBase64` removal**

Find:

```javascript
function handleExcelFile(file) {
    uploadStatus.textContent = "جاري معالجة وتدقيق ملف Excel...";
    uploadStatus.style.color = "var(--accent)";

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            state.originalWorkbook = workbook;

            // Process sheets
            processWorkbook(workbook);

            uploadStatus.innerHTML = `<strong>✅ تم تحميل الملف "${file.name}" بنجاح!</strong><br>تم استخراج البيانات وحفظها في المتصفح بشكل دائم — لن تحتاج لإعادة رفعه مجدداً.`;
            uploadStatus.style.color = "var(--success)";

            // Save filename only (we do NOT save the huge raw binary to prevent storage crashes/freezes)
            try {
                localStorage.setItem("cems_excel_filename", file.name);
                // Clean up any old binary to free space
                localStorage.removeItem("cems_excel_binary");
            } catch (storageErr) {
                console.warn("Could not store Excel filename:", storageErr);
            }

            // Save JSON data as well
            saveToLocalMemory();

            // Direct transfer to dashboard
            setTimeout(() => {
                switchTab("dashboard-tab");
            }, 1200);

        } catch (err) {
            console.error(err);
            uploadStatus.textContent = "فشلت قراءة الملف المالي: " + err.message;
            uploadStatus.style.color = "var(--danger)";
        }
    };
    reader.readAsArrayBuffer(file);
}

// Helper: convert ArrayBuffer to Base64 string
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

// Helper: convert Base64 back to Uint8Array
function base64ToUint8Array(base64) {
    const binary = atob(base64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
}
```

Replace it with:

```javascript
function handleExcelFile(file) {
    uploadStatus.textContent = "جاري معالجة وتدقيق ملف Excel...";
    uploadStatus.style.color = "var(--accent)";

    const reader = new FileReader();
    reader.onload = async function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            state.originalWorkbook = workbook;

            const parsed = processWorkbook(workbook);

            uploadStatus.textContent = `جاري رفع بيانات الملف "${file.name}" إلى جوجل شيت...`;

            const result = await callApi("bulkImportMembers", {
                members: parsed.members,
                families: parsed.families,
                months: state.monthsList
            });
            state.members = result.members;
            state.families = result.families;
            state.monthsList = result.months;

            uploadStatus.innerHTML = `<strong>✅ تم تحميل الملف "${file.name}" ورفعه لجوجل شيت بنجاح!</strong>`;
            uploadStatus.style.color = "var(--success)";

            populateMonthFilters();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
            renderLedgerTable();
            renderFamiliesTable();

            setTimeout(() => {
                switchTab("dashboard-tab");
            }, 1200);

        } catch (err) {
            console.error(err);
            uploadStatus.textContent = "فشلت قراءة الملف المالي أو رفعه: " + err.message;
            uploadStatus.style.color = "var(--danger)";
        }
    };
    reader.readAsArrayBuffer(file);
}
```

(The two base64/binary helper functions are deleted outright — they were
already dead code: nothing in the app ever called `localStorage.setItem`
with `cems_excel_binary`, only stale `getItem`/`removeItem` calls, which
are removed in Step 4 below.)

- [ ] **Step 3: Make `processWorkbook` return data instead of mutating `state` directly**

Find (the two lines that assign directly into `state`):

```javascript
            parsedMembers.push(member);
        });
        state.members = parsedMembers;
    }
```

Replace with:

```javascript
            parsedMembers.push(member);
        });
    }
```

Find:

```javascript
        Object.keys(grouped).forEach((headName, index) => {
            const mCount = grouped[headName].members.length;
            parsedFamilies.push({
                familyId: index + 1,
                headName: headName,
                memberCount: mCount,
                subscription: mCount * 10,
                totalPaid: grouped[headName].totalPaid,
                membersList: grouped[headName].members.join('، '),
                membersArr: grouped[headName].members
            });
        });
        state.families = parsedFamilies;
    }
}
```

Replace with:

```javascript
        Object.keys(grouped).forEach((headName, index) => {
            const mCount = grouped[headName].members.length;
            parsedFamilies.push({
                familyId: index + 1,
                headName: headName,
                memberCount: mCount,
                subscription: mCount * 10,
                totalPaid: grouped[headName].totalPaid,
                membersList: grouped[headName].members.join('، '),
                membersArr: grouped[headName].members
            });
        });
    }

    return { members: parsedMembers, families: parsedFamilies };
}
```

Also find the other assignment in the middle of the same function:

```javascript
            });
        });
        state.families = parsedFamilies;
    } else {
        // Fallback: group dynamically from التحصيل الشهري if no receipts sheet
```

Replace with:

```javascript
            });
        });
    } else {
        // Fallback: group dynamically from التحصيل الشهري if no receipts sheet
```

- [ ] **Step 4: Replace `saveToLocalMemory`/`loadCachedData` with `loadDataFromServer`**

Find:

```javascript
// Local Storage Sync (with safety error handling for browser storage limit)
function saveToLocalMemory() {
    try {
        localStorage.setItem("cems_data_members", JSON.stringify(state.members));
        localStorage.setItem("cems_data_families", JSON.stringify(state.families));
        localStorage.setItem("cems_data_months", JSON.stringify(state.monthsList));
        localStorage.setItem("cems_data_expenses", JSON.stringify(state.expenses));
        localStorage.setItem("cems_selected_ledger_months", JSON.stringify(state.selectedLedgerMonths || []));
    } catch (err) {
        console.error("Storage error:", err);
        if (err.name === "QuotaExceededError" || err.code === 22) {
            alert("⚠️ تنبيه: مساحة التخزين في المتصفح ممتلئة! يرجى إزالة المرفقات الكبيرة أو تنظيف سجل الصرف للمتابعة دون فقدان البيانات.");
        }
    }
}

function loadCachedData() {
    // Load months list cache first
    const cachedMonths = localStorage.getItem("cems_data_months");
    if (cachedMonths) {
        state.monthsList = JSON.parse(cachedMonths);
    } else {
        state.monthsList = JSON.parse(JSON.stringify(DEFAULT_MONTHS));
    }

    // Initialize/load selected Ledger months
    const cachedSelected = localStorage.getItem("cems_selected_ledger_months");
    if (cachedSelected) {
        try {
            state.selectedLedgerMonths = JSON.parse(cachedSelected);
        } catch(e) {
            state.selectedLedgerMonths = state.monthsList.map(m => m.key);
        }
    } else {
        // By default, select all months
        state.selectedLedgerMonths = state.monthsList.map(m => m.key);
    }

    // Load expenses
    const cachedExpenses = localStorage.getItem("cems_data_expenses");
    state.expenses = cachedExpenses ? JSON.parse(cachedExpenses) : [];
    
    // Ensure all expenses have valid IDs to enable working Delete options
    let idUpdated = false;
    state.expenses.forEach((e, idx) => {
        if (!e.id) {
            e.id = "exp_" + Date.now() + "_" + idx;
            idUpdated = true;
        }
    });
    if (idUpdated) {
        saveToLocalMemory();
    }

    // Load JSON dataset directly (very fast, no freeze)
    const cachedMembers = localStorage.getItem("cems_data_members");
    const cachedFamilies = localStorage.getItem("cems_data_families");
    const cachedFilename = localStorage.getItem("cems_excel_filename");

    // Clean up large raw Excel binary if it exists to free up localStorage space
    if (localStorage.getItem("cems_excel_binary")) {
        localStorage.removeItem("cems_excel_binary");
    }

    if (cachedMembers && cachedFamilies) {
        state.members = JSON.parse(cachedMembers);
        state.families = JSON.parse(cachedFamilies);
        
        // Update upload status info
        const statusEl = document.getElementById("upload-status");
        if (statusEl) {
            statusEl.innerHTML = `<strong>✅ البيانات محملة من الذاكرة الدائمة</strong><br>آخر ملف: "${cachedFilename || 'احصاء ال اطفيحه'}". يمكنك رفع ملف جديد لتحديث البيانات.`;
            statusEl.style.color = "var(--success)";
        }
        updateIndicators();
    } else {
        // No local cache yet, load the hardcoded default database seed
        state.members = JSON.parse(JSON.stringify(DEFAULT_MEMBERS));
        state.families = JSON.parse(JSON.stringify(DEFAULT_FAMILIES));
        
        // Save it so that localStorage has baseline values
        saveToLocalMemory();
        updateIndicators();
        
        const statusEl = document.getElementById("upload-status");
        if (statusEl) {
            statusEl.innerHTML = `<strong>✅ تم تحميل كشف الأسماء الافتراضي تلقائياً بالتطبيق</strong><br>البيانات مضمنة وجاهزة للعمل والتعديل مباشرة دون حاجة لرفع Excel.`;
            statusEl.style.color = "var(--success)";
        }
    }
}

// Clear browser cache
document.getElementById("btn-clear-cache").addEventListener("click", () => {
    if (confirm("هل أنت متأكد من رغبتك في مسح كافة التعديلات الحالية وإعادة تعيين كشف الصندوق إلى البيانات الافتراضية؟")) {
        localStorage.removeItem("cems_data_members");
        localStorage.removeItem("cems_data_families");
        localStorage.removeItem("cems_data_months");
        localStorage.removeItem("cems_data_expenses");
        localStorage.removeItem("cems_excel_binary");
        localStorage.removeItem("cems_excel_filename");
        
        // Reload default data
        state.monthsList = JSON.parse(JSON.stringify(DEFAULT_MONTHS));
        state.members = JSON.parse(JSON.stringify(DEFAULT_MEMBERS));
        state.families = JSON.parse(JSON.stringify(DEFAULT_FAMILIES));
        state.expenses = [];
        state.originalWorkbook = null;
        
        saveToLocalMemory();
        updateIndicators();
        
        uploadStatus.textContent = "تمت إعادة تعيين كشف الصندوق إلى الأسماء الافتراضية بنجاح ونظيفة.";
        uploadStatus.style.color = "var(--success)";
        
        // Refresh tables/views
        populateMonthFilters();
        populateAddMemberFamilies();
        populateReceiptFamilies();
        
        renderDashboard();
        renderLedgerTable();
        renderFamiliesTable();
        
        alert("تمت إعادة تعيين كشف الصندوق للبيانات الافتراضية بنجاح.");
    }
});
```

Replace it with:

```javascript
// Loads every piece of app data from the Google Sheet backend in one round
// trip and populates `state`. Called once on startup, and again by the
// "reload from sheet" button (see below) or after an Excel bulk import.
async function loadDataFromServer() {
    const data = await callApi("getAllData", {});
    state.members = data.members;
    state.families = data.families;
    state.monthsList = data.months.map(m => ({ key: m.key, id: m.id }));
    state.selectedLedgerMonths = data.months.filter(m => m.selected).map(m => m.key);
    state.expenses = data.expenses;
    state.users = data.users;
    state.pendingUsers = data.pendingUsers;
    state.settings = data.settings;

    const statusEl = document.getElementById("upload-status");
    if (statusEl) {
        if (state.members.length > 0) {
            statusEl.innerHTML = `<strong>✅ البيانات محملة من جوجل شيت</strong><br>${state.members.length} فرد، ${state.families.length} عائلة. يمكنك رفع ملف Excel لاستبدال البيانات بالكامل.`;
        } else {
            statusEl.innerHTML = `<strong>لا توجد بيانات أعضاء بعد.</strong><br>ارفع ملف Excel هنا لتعبئة قاعدة البيانات على جوجل شيت لأول مرة.`;
        }
        statusEl.style.color = "var(--success)";
    }
    updateIndicators();
}

// Reload button — re-fetches everything from the Sheet (e.g. after another
// user, or a direct edit on the Sheet itself, changed something).
document.getElementById("btn-clear-cache").addEventListener("click", async () => {
    if (!confirm("سيتم إعادة تحميل كل البيانات من جوجل شيت، وسيتم فقدان أي تعديل لم يُحفظ بعد. متابعة؟")) return;
    try {
        uploadStatus.textContent = "جاري إعادة التحميل من جوجل شيت...";
        uploadStatus.style.color = "var(--accent)";
        await loadDataFromServer();

        populateMonthFilters();
        populateAddMemberFamilies();
        populateReceiptFamilies();
        renderCharts();
        renderLedgerTable();
        renderFamiliesTable();

        alert("✅ تمت إعادة تحميل البيانات من جوجل شيت بنجاح.");
    } catch (err) {
        alert("تعذّرت إعادة التحميل: " + err.message);
    }
});
```

- [ ] **Step 5: Fix the dead `renderDashboard()` references**

`renderDashboard` is called in 6 places in the current file but was never
defined anywhere — a pre-existing bug (every one of those calls throws
`ReferenceError: renderDashboard is not defined`, silently breaking
whatever ran after it, since it's always the last line of its handler).
Since this task touches surrounding code at several of those call sites
anyway, fix all six by replacing the call with `renderCharts()` (which
already calls `updateIndicators()` internally — the same "refresh the
dashboard" effect the dead call was clearly meant to have). Use this exact
substitution everywhere `renderDashboard();` appears standalone on its own
line in `app.js`:

```javascript
        renderDashboard();
```
→
```javascript
        renderCharts();
```

This plan's later frontend tasks (4, 5) rewrite the surrounding handlers
at 4 of those 6 call sites anyway (add member, add month, edit member, edit
month/delete month) and will carry this fix along automatically — only
apply this substitution standalone for any occurrence that's still
`renderDashboard()` by the time you reach the end of Frontend Task 5. Run
`grep -n renderDashboard app.js` at that point; it must return nothing.

- [ ] **Step 6: Update the DOMContentLoaded bootstrap to call the new loader**

Find:

```javascript
// Initial Setup
document.addEventListener("DOMContentLoaded", () => {
    // Current date fill
    const today = new Date().toISOString().split("T")[0];
    liveDateSpan.textContent = today;
    document.getElementById("p-receipt-date").textContent = today;
    document.getElementById("r-date").textContent = today;

    // *** CRITICAL: Load cached data FIRST ***
    loadCachedData();
    populateMonthFilters();
    populateAddMemberFamilies();

    // Check if user is logged in (session persisted by email)
    const cachedEmail = localStorage.getItem("cems_logged_email");
    if (cachedEmail) {
        const userObj = findUserByEmail(cachedEmail);
        if (userObj) {
            doLogin(userObj);
        } else {
            localStorage.removeItem("cems_logged_email");
        }
    }
```

Replace it with:

```javascript
// Initial Setup
document.addEventListener("DOMContentLoaded", async () => {
    // Current date fill
    const today = new Date().toISOString().split("T")[0];
    liveDateSpan.textContent = today;
    document.getElementById("p-receipt-date").textContent = today;
    document.getElementById("r-date").textContent = today;

    // *** CRITICAL: Load data from the Google Sheet backend FIRST ***
    try {
        await loadDataFromServer();
    } catch (err) {
        alert("تعذّر الاتصال بقاعدة البيانات (جوجل شيت): " + err.message + "\n\nتحقق من رابط API_URL في app.js ومن اتصالك بالإنترنت.");
    }
    populateMonthFilters();
    populateAddMemberFamilies();

    // Check if user is logged in (session persisted by email) — the email
    // is only a client-side "stay logged in" convenience; the account
    // itself is re-validated against the freshly-loaded state.users.
    const cachedEmail = localStorage.getItem("cems_logged_email");
    if (cachedEmail) {
        const userObj = findUserInState_(cachedEmail);
        if (userObj) {
            doLogin(userObj);
        } else {
            localStorage.removeItem("cems_logged_email");
        }
    }
```

- [ ] **Step 7: Verify**

```bash
node --check app.js
```

Expected: no output. Then run `grep -n "loadCachedData\|saveToLocalMemory\|ADMINS\[" app.js` — expected: no matches (both functions and every remaining `ADMINS[...]` lookup are gone; the next task, Frontend Task 3, still has a few plain `ADMINS` / `MASTER_ADMIN_EMAIL` identifier references left to clean up in the login/registration flow, so don't expect zero matches for those two words yet).

- [ ] **Step 8: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: load app data from Google Sheets backend instead of localStorage

Removes the hardcoded ADMINS object, loadAdminPasswords(),
findUserByEmail(), saveToLocalMemory(), and loadCachedData().
loadDataFromServer() now populates all of `state` from the backend's
getAllData action on startup, and the Excel import path pushes parsed
rows to the Sheet via bulkImportMembers instead of writing to
localStorage. Also fixes six pre-existing dead calls to an undefined
renderDashboard() by pointing them at renderCharts().

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 3: Auth flows — login, password reset, registration, logout

**Files:**
- Modify: `app.js` (the login/registration block, originally around lines 260–530)

- [ ] **Step 1: Rewrite the login click handler**

Find:

```javascript
// Login Handler — email-based
btnSubmitLogin.addEventListener("click", () => {
    const email = loginEmailInput.value.trim().toLowerCase();
    const pass = loginPassInput.value;

    if (!email) {
        showLoginError("الرجاء إدخال بريدك الإلكتروني.");
        return;
    }
    if (!pass) {
        showLoginError("الرجاء إدخال كلمة المرور.");
        return;
    }

    const userObj = findUserByEmail(email);
    if (!userObj) {
        showLoginError("لا يوجد حساب بهذا البريد. تحقق من أنك مسجّل أو أرسل طلب للتسجيل.");
        return;
    }

    if (userObj.pass !== pass) {
        showLoginError("كلمة المرور غير صحيحة. يرجى المحاولة مجدداً.");
        return;
    }

    if (userObj.isActive === false) {
        showLoginError("حسابك معطّل حالياً. تواصل مع المسؤول لفعل حسابك.");
        return;
    }

    localStorage.setItem("cems_logged_email", email);

    // Force password change on first login for built-in admins
    if (userObj.firstLoginDone === false && userObj.isBuiltIn) {
        showPasswordChangeDialog(userObj, () => doLogin(userObj));
    } else {
        doLogin(userObj);
    }
});
```

Replace it with:

```javascript
// Login Handler — email-based, verified against the Google Sheet backend
btnSubmitLogin.addEventListener("click", async () => {
    const email = loginEmailInput.value.trim().toLowerCase();
    const pass = loginPassInput.value;

    if (!email) {
        showLoginError("الرجاء إدخال بريدك الإلكتروني.");
        return;
    }
    if (!pass) {
        showLoginError("الرجاء إدخال كلمة المرور.");
        return;
    }

    btnSubmitLogin.disabled = true;
    try {
        const result = await callApi("login", { email, pass });
        if (result.error) {
            showLoginError(result.error);
            return;
        }
        const userObj = result.user;
        localStorage.setItem("cems_logged_email", email);

        // Force password change on first login for built-in admins
        if (userObj.firstLoginDone === false && userObj.isBuiltIn) {
            showPasswordChangeDialog(userObj, () => doLogin(userObj));
        } else {
            doLogin(userObj);
        }
    } catch (err) {
        showLoginError("تعذّر تسجيل الدخول: " + err.message);
    } finally {
        btnSubmitLogin.disabled = false;
    }
});
```

- [ ] **Step 2: Rewrite the password reset handler**

Find:

```javascript
// Password Reset Handler (Email-based)
document.getElementById("link-reset-pass").addEventListener("click", (e) => {
    e.preventDefault();
    const email = loginEmailInput.value.trim().toLowerCase();
    if (!email) {
        alert("الرجاء إدخال البريد الإلكتروني للحساب المراد إعادة تعيين رمزه أولاً.");
        return;
    }

    const userObj = findUserByEmail(email);
    if (!userObj) {
        alert("لا يوجد حساب مشرف مالي مسجل بهذا البريد الإلكتروني.");
        return;
    }

    const confirmReset = confirm(`هل أنت متأكد من رغبتك في إعادة تعيين رمز المرور لحساب المشرف (${userObj.name})؟ \n\nسيتم استخدام رمز المرور الافتراضي (ABC12345) وسيُطلب منك تغييره عند الدخول.`);
    if (confirmReset) {
        if (userObj.isBuiltIn) {
            const key = userObj.key;
            localStorage.removeItem(`cems_admin_pass_${key}`);
            localStorage.removeItem(`cems_admin_firstlogin_${key}`);
            ADMINS[key].pass = "ABC12345";
            ADMINS[key].firstLoginDone = false;
        } else {
            const dynamicUsers = JSON.parse(localStorage.getItem("cems_dynamic_users") || "[]");
            const idx = dynamicUsers.findIndex(u => u.email.toLowerCase() === email);
            if (idx !== -1) {
                dynamicUsers[idx].pass = "ABC12345";
                dynamicUsers[idx].firstLoginDone = false;
                localStorage.setItem("cems_dynamic_users", JSON.stringify(dynamicUsers));
            }
        }
        alert(`تمت إعادة تعيين رمز مرور الحساب بنجاح إلى: ABC12345 \nيرجى تسجيل الدخول فيه الآن وتعديله.`);
        loginPassInput.value = "";
        loginErrorMsg.style.display = "none";
    }
});
```

Replace it with:

```javascript
// Password Reset Handler (Email-based)
document.getElementById("link-reset-pass").addEventListener("click", async (e) => {
    e.preventDefault();
    const email = loginEmailInput.value.trim().toLowerCase();
    if (!email) {
        alert("الرجاء إدخال البريد الإلكتروني للحساب المراد إعادة تعيين رمزه أولاً.");
        return;
    }

    const userObj = findUserInState_(email);
    if (!userObj) {
        alert("لا يوجد حساب مشرف مالي مسجل بهذا البريد الإلكتروني.");
        return;
    }

    const confirmReset = confirm(`هل أنت متأكد من رغبتك في إعادة تعيين رمز المرور لحساب المشرف (${userObj.name})؟ \n\nسيتم استخدام رمز المرور الافتراضي (ABC12345) وسيُطلب منك تغييره عند الدخول.`);
    if (!confirmReset) return;

    try {
        await callApi("resetPassword", { email });
        alert(`تمت إعادة تعيين رمز مرور الحساب بنجاح إلى: ABC12345 \nيرجى تسجيل الدخول فيه الآن وتعديله.`);
        loginPassInput.value = "";
        loginErrorMsg.style.display = "none";
    } catch (err) {
        alert("تعذّرت إعادة تعيين رمز المرور: " + err.message);
    }
});
```

- [ ] **Step 3: Rewrite the registration submit handler**

Find:

```javascript
    if (findUserByEmail(email)) {
        errorEl.textContent = "هذا البريد الإلكتروني مسجل بالفعل لمشرف آخر.";
        errorEl.style.display = "block";
        return;
    }

    const pendingList = JSON.parse(localStorage.getItem("cems_pending_users") || "[]");
    if (pendingList.some(r => r.email === email)) {
        errorEl.textContent = "يوجد بالفعل طلب تسجيل معلق لهذا البريد الإلكتروني.";
        errorEl.style.display = "block";
        return;
    }

    const name = firstName + " " + lastName;
    const role = "مشرف مالي";

    // Add to pending
    pendingList.push({
        name,
        email,
        role,
        pass,
        date: new Date().toISOString().split("T")[0]
    });
    localStorage.setItem("cems_pending_users", JSON.stringify(pendingList));

    alert(`✅ تم إرسال طلب تسجيلك بنجاح باسم (${name})! يرجى انتظار موافقة المسؤول جهاد زكري لتتمكن من تسجيل الدخول.`);
    
    // Clear
    document.getElementById("reg-firstname").value = "";
    document.getElementById("reg-lastname").value = "";
    document.getElementById("reg-email").value = "";
    document.getElementById("reg-pass").value = "";
    document.getElementById("reg-pass2").value = "";

    document.getElementById("panel-register").style.display = "none";
    document.getElementById("panel-login").style.display = "block";
});
```

Replace it with:

```javascript
    if (findUserInState_(email) || state.pendingUsers.some(r => r.email.toLowerCase() === email)) {
        errorEl.textContent = "هذا البريد الإلكتروني مسجل بالفعل أو له طلب تسجيل معلق.";
        errorEl.style.display = "block";
        return;
    }

    try {
        state.pendingUsers = await callApi("requestRegistration", { firstName, lastName, email, pass });

        alert(`✅ تم إرسال طلب تسجيلك بنجاح باسم (${firstName} ${lastName})! يرجى انتظار موافقة المسؤول جهاد زكري لتتمكن من تسجيل الدخول.`);

        document.getElementById("reg-firstname").value = "";
        document.getElementById("reg-lastname").value = "";
        document.getElementById("reg-email").value = "";
        document.getElementById("reg-pass").value = "";
        document.getElementById("reg-pass2").value = "";

        document.getElementById("panel-register").style.display = "none";
        document.getElementById("panel-login").style.display = "block";
    } catch (err) {
        errorEl.textContent = "تعذّر إرسال الطلب: " + err.message;
        errorEl.style.display = "block";
    }
});
```

- [ ] **Step 4: Rewrite the forced password-change dialog's save handler**

Find:

```javascript
        // Save for built-in admin via ADMINS key
        userObj.pass = p1;
        userObj.firstLoginDone = true;
        const adminKey = Object.keys(ADMINS).find(k => ADMINS[k].email.toLowerCase() === userObj.email.toLowerCase());
        if (adminKey) {
            localStorage.setItem(`cems_admin_pass_${adminKey}`, p1);
            localStorage.setItem(`cems_admin_firstlogin_${adminKey}`, 'done');
        }

        overlay.remove();
        onSuccess();
    });
}
```

Replace it with:

```javascript
        try {
            const result = await callApi("changePassword", { key: userObj.key, newPass: p1 });
            userObj.pass = undefined;
            userObj.firstLoginDone = true;
            Object.assign(userObj, result.user);
            overlay.remove();
            onSuccess();
        } catch (err) {
            errEl.style.display = 'block';
            errEl.textContent = 'تعذّر حفظ كلمة المرور: ' + err.message;
        }
    });
}
```

Also find (a few lines above, in the same `showPasswordChangeDialog` function) the button click listener's opening line, since it must become `async` to `await` inside it:

```javascript
    document.getElementById('btn-confirm-pwd').addEventListener('click', () => {
```

Replace it with:

```javascript
    document.getElementById('btn-confirm-pwd').addEventListener('click', async () => {
```

- [ ] **Step 5: Update `doLogin`'s master-admin check**

Find:

```javascript
    // Show User Management tab only for master admin
    const isMaster = userObj.email && userObj.email.toLowerCase() === MASTER_ADMIN_EMAIL.toLowerCase();
    const navMgmt = document.getElementById("nav-usermgmt");
```

Replace it with:

```javascript
    // Show User Management tab only for the master admin (per the Users
    // sheet's isMaster flag, not a hardcoded email compare)
    const isMaster = userObj.isMaster === true;
    const navMgmt = document.getElementById("nav-usermgmt");
```

- [ ] **Step 6: Verify**

```bash
node --check app.js
```

Expected: no output. Then `grep -n "findUserByEmail\|MASTER_ADMIN_EMAIL\b" app.js` — expected: no matches (the constant and function no longer exist anywhere; `renderActiveUsers`'s own use of `MASTER_ADMIN_EMAIL` is cleaned up in Frontend Task 9, so if this grep still shows a hit inside `renderActiveUsers` at this point, that's expected and will be resolved later — but there should be zero hits inside the login/registration block modified in this task).

- [ ] **Step 7: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire login, registration, and password reset to the API

Login now calls the backend's login action instead of checking a
local ADMINS object; password reset and the forced first-login
change now call resetPassword/changePassword; registration requests
go through requestRegistration instead of a localStorage array.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 4: Members & Families — payment toggle, add, edit, delete

**Files:**
- Modify: `app.js` (ledger checkbox handlers, add-member handler, member edit
  modal, delete-member)

- [ ] **Step 1: Replace both ledger payment-checkbox handlers with `setMemberPayment`**

`renderLedgerTable()` builds this same checkbox `change` handler twice —
once inside the `filterMonth === "all"` branch (iterating `visibleMonths`)
and once in the `else` branch (single `filterMonth`). Both copies are
byte-for-byte identical except which `m.key` / `filterMonth` variable they
close over. Replace **both** occurrences of this exact block:

```javascript
                chk.addEventListener("change", (e) => {
                    const isChecked = e.target.checked;
                    const paymentVal = isChecked ? 10 : 0;
                    
                    member.payments[m.key] = paymentVal;
                    
                    let memberTotal = 0;
                    state.monthsList.forEach(mon => {
                        memberTotal += (member.payments[mon.key] || 0);
                    });
                    member.sum = memberTotal;
                    tdSum.textContent = memberTotal + " شيكل";

                    recalculateFamilyTotals(member.parent);
                    saveToLocalMemory();
                    renderCharts();
                    populateMonthFilters();
                });
```

with:

```javascript
                chk.addEventListener("change", async (e) => {
                    const isChecked = e.target.checked;
                    const paymentVal = isChecked ? 10 : 0;
                    chk.disabled = true;

                    try {
                        const result = await callApi("setMemberPayment", { id: member.id, month: m.key, amount: paymentVal });
                        state.members = result.members;
                        state.families = result.families;
                        renderLedgerTable();
                        renderCharts();
                        populateMonthFilters();
                    } catch (err) {
                        e.target.checked = !isChecked;
                        alert("تعذّر حفظ الدفعة: " + err.message);
                        chk.disabled = false;
                    }
                });
```

and this exact block (the single-`filterMonth` branch — note it closes over
`filterMonth` directly, not `m.key`):

```javascript
            chk.addEventListener("change", (e) => {
                const isChecked = e.target.checked;
                const paymentVal = isChecked ? 10 : 0;
                
                member.payments[filterMonth] = paymentVal;
                
                let memberTotal = 0;
                state.monthsList.forEach(mon => {
                    memberTotal += (member.payments[mon.key] || 0);
                });
                member.sum = memberTotal;
                tdSum.textContent = memberTotal + " شيكل";

                recalculateFamilyTotals(member.parent);
                saveToLocalMemory();
                renderCharts();
                populateMonthFilters();
            });
```

with:

```javascript
            chk.addEventListener("change", async (e) => {
                const isChecked = e.target.checked;
                const paymentVal = isChecked ? 10 : 0;
                chk.disabled = true;

                try {
                    const result = await callApi("setMemberPayment", { id: member.id, month: filterMonth, amount: paymentVal });
                    state.members = result.members;
                    state.families = result.families;
                    renderLedgerTable();
                    renderCharts();
                    populateMonthFilters();
                } catch (err) {
                    e.target.checked = !isChecked;
                    alert("تعذّر حفظ الدفعة: " + err.message);
                    chk.disabled = false;
                }
            });
```

Note this now calls `renderLedgerTable()` itself instead of relying on the
caller to re-render — which is correct and necessary, since the whole
table (including this exact checkbox and its `tdSum` sibling) is rebuilt
from the fresh `state.members` returned by the API, replacing the old
direct-DOM-write-then-continue approach.

- [ ] **Step 2: Rewrite the "add member" handler**

Find:

```javascript
// Add Member Event
document.getElementById("btn-add-member").addEventListener("click", () => {
    const nameInput = document.getElementById("add-member-name");
    const selectFamily = document.getElementById("add-member-select-family");
    const newFamilyInput = document.getElementById("add-member-new-family");

    const name = nameInput.value.trim();
    if (!name) {
        alert("الرجاء إدخال اسم العضو الجديد.");
        return;
    }

    // Determine family head (parent)
    let parent = "";
    const selectedHead = selectFamily.value;
    const newHead = newFamilyInput.value.trim();

    if (newHead) {
        parent = newHead;
    } else if (selectedHead) {
        parent = selectedHead;
    } else {
        alert("الرجاء اختيار رب عائلة من القائمة أو كتابة اسم رب عائلة جديد.");
        return;
    }

    // Check if member already exists
    const exists = state.members.some(m => m.name.trim().toLowerCase() === name.toLowerCase());
    if (exists) {
        alert("هذا العضو مسجل بالفعل في القائمة.");
        return;
    }

    // Generate new member ID
    const newId = state.members.length > 0 ? Math.max(...state.members.map(m => m.id)) + 1 : 1;

    // Initialize payments
    const payments = {};
    state.monthsList.forEach(m => {
        payments[m.key] = 0;
    });

    const newMember = {
        id: newId,
        name: name,
        parent: parent,
        payments: payments,
        sum: 0
    };

    // Add member to state
    state.members.push(newMember);

    // Sync state.families
    let family = state.families.find(f => f.headName.trim().toLowerCase() === parent.toLowerCase());
    if (family) {
        // Increment existing family count and sub
        family.memberCount += 1;
        family.subscription += 10;
        if (!family.membersArr) family.membersArr = [];
        family.membersArr.push(name);
        family.membersList = family.membersArr.join("، ");
    } else {
        // Create new family
        const newFamId = state.families.length > 0 ? Math.max(...state.families.map(f => f.familyId)) + 1 : 1;
        const newFamily = {
            familyId: newFamId,
            headName: parent,
            memberCount: 1,
            subscription: 10,
            totalPaid: 0,
            membersList: name,
            membersArr: [name]
        };
        state.families.push(newFamily);
    }

    // Persist
    saveToLocalMemory();

    // Rerender and clean inputs
    nameInput.value = "";
    newFamilyInput.value = "";
    
    // Refresh family dropdown list in addition to other views
    populateAddMemberFamilies();
    populateReceiptFamilies();
    
    renderDashboard();
    renderLedgerTable();
    renderFamiliesTable();
    
    alert(`✅ تم إضافة العضو (${name}) بنجاح وإلحاقه بعائلة (${parent}).`);
});
```

Replace it with:

```javascript
// Add Member Event
document.getElementById("btn-add-member").addEventListener("click", async () => {
    const nameInput = document.getElementById("add-member-name");
    const selectFamily = document.getElementById("add-member-select-family");
    const newFamilyInput = document.getElementById("add-member-new-family");

    const name = nameInput.value.trim();
    if (!name) {
        alert("الرجاء إدخال اسم العضو الجديد.");
        return;
    }

    let parent = "";
    const selectedHead = selectFamily.value;
    const newHead = newFamilyInput.value.trim();

    if (newHead) {
        parent = newHead;
    } else if (selectedHead) {
        parent = selectedHead;
    } else {
        alert("الرجاء اختيار رب عائلة من القائمة أو كتابة اسم رب عائلة جديد.");
        return;
    }

    const exists = state.members.some(m => m.name.trim().toLowerCase() === name.toLowerCase());
    if (exists) {
        alert("هذا العضو مسجل بالفعل في القائمة.");
        return;
    }

    const btn = document.getElementById("btn-add-member");
    btn.disabled = true;
    try {
        const result = await callApi("addMember", { name, parent });
        state.members = result.members;
        state.families = result.families;

        nameInput.value = "";
        newFamilyInput.value = "";

        populateAddMemberFamilies();
        populateReceiptFamilies();

        renderCharts();
        renderLedgerTable();
        renderFamiliesTable();

        alert(`✅ تم إضافة العضو (${name}) بنجاح وإلحاقه بعائلة (${parent}).`);
    } catch (err) {
        alert("تعذّرت إضافة العضو: " + err.message);
    } finally {
        btn.disabled = false;
    }
});
```

- [ ] **Step 3: Rewrite the member-edit save handler**

Find:

```javascript
// Save edited member info
const btnSaveMemberEdit = document.getElementById("btn-save-member-edit");
if (btnSaveMemberEdit) {
    btnSaveMemberEdit.addEventListener("click", () => {
        const id = document.getElementById("edit-member-id").value;
        const newName = document.getElementById("edit-member-name").value.trim();
        const selectedFam = document.getElementById("edit-member-select-family").value;
        const newFamInput = document.getElementById("edit-member-new-family").value.trim();

        if (!newName) {
            alert("الرجاء إدخال اسم العضو.");
            return;
        }

        let parent = "";
        if (newFamInput) {
            parent = newFamInput;
        } else if (selectedFam) {
            parent = selectedFam;
        } else {
            alert("الرجاء اختيار رب عائلة أو كتابة اسم رب عائلة جديد.");
            return;
        }

        const member = state.members.find(m => String(m.id) === String(id));
        if (!member) return;

        const oldName = member.name;
        const oldParent = member.parent;

        // 1. Update member details
        member.name = newName;
        member.parent = parent;

        // 2. Adjust families list
        // Remove name from old family members registry
        if (oldParent) {
            const oldFamily = state.families.find(f => normalizeArabicName(f.headName) === normalizeArabicName(oldParent));
            if (oldFamily) {
                if (oldFamily.membersArr) {
                    oldFamily.membersArr = oldFamily.membersArr.filter(n => normalizeArabicName(n) !== normalizeArabicName(oldName));
                    oldFamily.membersList = oldFamily.membersArr.join("، ");
                    oldFamily.memberCount = oldFamily.membersArr.length;
                    oldFamily.subscription = oldFamily.memberCount * 10;
                } else {
                    oldFamily.memberCount = Math.max(0, oldFamily.memberCount - 1);
                    oldFamily.subscription = oldFamily.memberCount * 10;
                }
                if (oldFamily.memberCount === 0) {
                    state.families = state.families.filter(f => normalizeArabicName(f.headName) !== normalizeArabicName(oldParent));
                } else {
                    recalculateFamilyTotals(oldParent);
                }
            }
        }

        // Add name to new family register
        let newFamily = state.families.find(f => normalizeArabicName(f.headName) === normalizeArabicName(parent));
        if (newFamily) {
            if (!newFamily.membersArr) newFamily.membersArr = [];
            if (!newFamily.membersArr.some(n => normalizeArabicName(n) === normalizeArabicName(newName))) {
                newFamily.membersArr.push(newName);
            }
            newFamily.membersList = newFamily.membersArr.join("، ");
            newFamily.memberCount = newFamily.membersArr.length;
            newFamily.subscription = newFamily.memberCount * 10;
            recalculateFamilyTotals(parent);
        } else {
            // Create new family registry
            const newFamId = state.families.length > 0 ? Math.max(...state.families.map(f => f.familyId)) + 1 : 1;
            state.families.push({
                familyId: newFamId,
                headName: parent,
                memberCount: 1,
                subscription: 10,
                totalPaid: 0,
                membersList: newName,
                membersArr: [newName]
            });
            recalculateFamilyTotals(parent);
        }

        // Save changes
        saveToLocalMemory();

        // Close Modal and Refresh
        closeEditMemberModal();
        renderLedgerTable();
        renderFamiliesTable();
        populateAddMemberFamilies();
        populateReceiptFamilies();
        renderDashboard();
    });
}
```

Replace it with:

```javascript
// Save edited member info
const btnSaveMemberEdit = document.getElementById("btn-save-member-edit");
if (btnSaveMemberEdit) {
    btnSaveMemberEdit.addEventListener("click", async () => {
        const id = document.getElementById("edit-member-id").value;
        const newName = document.getElementById("edit-member-name").value.trim();
        const selectedFam = document.getElementById("edit-member-select-family").value;
        const newFamInput = document.getElementById("edit-member-new-family").value.trim();

        if (!newName) {
            alert("الرجاء إدخال اسم العضو.");
            return;
        }

        let parent = "";
        if (newFamInput) {
            parent = newFamInput;
        } else if (selectedFam) {
            parent = selectedFam;
        } else {
            alert("الرجاء اختيار رب عائلة أو كتابة اسم رب عائلة جديد.");
            return;
        }

        btnSaveMemberEdit.disabled = true;
        try {
            const result = await callApi("updateMember", { id, name: newName, parent });
            state.members = result.members;
            state.families = result.families;

            closeEditMemberModal();
            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
        } catch (err) {
            alert("تعذّر حفظ تعديل العضو: " + err.message);
        } finally {
            btnSaveMemberEdit.disabled = false;
        }
    });
}
```

- [ ] **Step 4: Rewrite `window.deleteMember`**

Find:

```javascript
window.deleteMember = function(id, name) {
    showConfirm(`هل أنت متأكد من حذف العضو (${name}) نهائياً من الصندوق؟\nسيتأثر مجموع الأسرة والتقارير المالية بهذا التغيير.`, () => {
        const member = state.members.find(m => String(m.id) === String(id));
        if (!member) {
            console.error("Member to delete not found, ID was:", id);
            return;
        }
        
        const parentName = member.parent;
        
        // Remove from state.members
        state.members = state.members.filter(m => String(m.id) !== String(id));
        
        // Update family details
        if (parentName) {
            const family = state.families.find(f => normalizeArabicName(f.headName) === normalizeArabicName(parentName));
            if (family) {
                if (family.membersArr) {
                    family.membersArr = family.membersArr.filter(n => normalizeArabicName(n) !== normalizeArabicName(name));
                    family.membersList = family.membersArr.join("، ");
                    family.memberCount = family.membersArr.length;
                    family.subscription = family.memberCount * 10;
                } else {
                    family.memberCount = Math.max(0, family.memberCount - 1);
                    family.subscription = family.memberCount * 10;
                }
                
                // If family has no more individuals, delete family completely
                if (family.memberCount === 0) {
                    state.families = state.families.filter(f => normalizeArabicName(f.headName) !== normalizeArabicName(parentName));
                } else {
                    recalculateFamilyTotals(parentName);
                }
            }
        }
        
        // Persist changes
        saveToLocalMemory();
        
        // Refresh all views
        renderLedgerTable();
        renderFamiliesTable();
        populateAddMemberFamilies();
        populateReceiptFamilies();
        renderCharts();
        
        alert(`🗑️ تم حذف العضو (${name}) بنجاح.`);
    });
};
```

Replace it with:

```javascript
window.deleteMember = function(id, name) {
    showConfirm(`هل أنت متأكد من حذف العضو (${name}) نهائياً من الصندوق؟\nسيتأثر مجموع الأسرة والتقارير المالية بهذا التغيير.`, async () => {
        try {
            const result = await callApi("deleteMember", { id });
            state.members = result.members;
            state.families = result.families;

            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();

            alert(`🗑️ تم حذف العضو (${name}) بنجاح.`);
        } catch (err) {
            alert("تعذّر حذف العضو: " + err.message);
        }
    });
};
```

- [ ] **Step 5: Verify**

```bash
node --check app.js
```

Expected: no output.

- [ ] **Step 6: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire member/family CRUD and ledger payments to the API

The ledger's per-cell payment checkbox, add-member, edit-member, and
delete-member now all call the matching backend action and refresh
state.members/state.families from its response instead of mutating
local state directly and calling saveToLocalMemory().

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 5: Months — add, edit, delete, and the multi-month selector

**Files:**
- Modify: `app.js` (add-month handler, multi-month checkbox handler,
  edit-month save handler, delete-month)

- [ ] **Step 1: Rewrite the "add month" handler**

Find:

```javascript
// Add Month Event
document.getElementById("btn-add-month").addEventListener("click", () => {
    const monthNameInput = document.getElementById("add-month-name");
    const monthIdSelect = document.getElementById("add-month-id");

    const monthKey = monthNameInput.value.trim();
    if (!monthKey) {
        alert("الرجاء إدخال اسم الشهر والسنة (مثال: يناير 2027).");
        return;
    }

    const monthId = parseInt(monthIdSelect.value);

    // Check if month already exists
    const exists = state.monthsList.some(mon => mon.key.trim().toLowerCase() === monthKey.toLowerCase());
    if (exists) {
        alert("هذا الشهر مضاف بالفعل في القائمة.");
        return;
    }

    // Add month to state.monthsList
    state.monthsList.push({
        key: monthKey,
        id: monthId
    });

    // Initialize this month's payment to 0 for all existing members
    state.members.forEach(m => {
        if (m.payments[monthKey] === undefined) {
            m.payments[monthKey] = 0;
        }
    });

    // Persist
    saveToLocalMemory();

    // Rerender and clean inputs
    monthNameInput.value = "";

    // Refresh dynamically populated month components
    populateMonthFilters();
    renderDashboard();
    renderLedgerTable();
    renderFamiliesTable();
    
    alert(`✅ تم إضافة الشهر (${monthKey}) بنجاح إلى جدول الاشتراكات والتحصيل.`);
});
```

Replace it with:

```javascript
// Add Month Event
document.getElementById("btn-add-month").addEventListener("click", async () => {
    const monthNameInput = document.getElementById("add-month-name");
    const monthIdSelect = document.getElementById("add-month-id");

    const monthKey = monthNameInput.value.trim();
    if (!monthKey) {
        alert("الرجاء إدخال اسم الشهر والسنة (مثال: يناير 2027).");
        return;
    }

    const monthId = parseInt(monthIdSelect.value);

    const exists = state.monthsList.some(mon => mon.key.trim().toLowerCase() === monthKey.toLowerCase());
    if (exists) {
        alert("هذا الشهر مضاف بالفعل في القائمة.");
        return;
    }

    const btn = document.getElementById("btn-add-month");
    btn.disabled = true;
    try {
        const result = await callApi("addMonth", { key: monthKey, id: monthId });
        state.monthsList = result.months.map(m => ({ key: m.key, id: m.id }));
        state.selectedLedgerMonths.push(monthKey);
        state.members = result.members;

        monthNameInput.value = "";

        populateMonthFilters();
        renderCharts();
        renderLedgerTable();
        renderFamiliesTable();

        alert(`✅ تم إضافة الشهر (${monthKey}) بنجاح إلى جدول الاشتراكات والتحصيل.`);
    } catch (err) {
        alert("تعذّرت إضافة الشهر: " + err.message);
    } finally {
        btn.disabled = false;
    }
});
```

- [ ] **Step 2: Rewrite the multi-month selector checkbox handler**

Find (inside `populateMonthFilters()`, the "Multi-Month Checkboxes Panel"
section):

```javascript
                    chk.addEventListener("change", (e) => {
                        const activeChecked = e.target.checked;
                        if (activeChecked) {
                            if (!state.selectedLedgerMonths.includes(m.key)) {
                                state.selectedLedgerMonths.push(m.key);
                            }
                            label.classList.add("checked");
                        } else {
                            state.selectedLedgerMonths = state.selectedLedgerMonths.filter(k => k !== m.key);
                            label.classList.remove("checked");
                        }
                        saveToLocalMemory();
                        renderLedgerTable();
                    });
```

Replace it with:

```javascript
                    chk.addEventListener("change", async (e) => {
                        const activeChecked = e.target.checked;
                        const previous = state.selectedLedgerMonths.slice();
                        if (activeChecked) {
                            if (!state.selectedLedgerMonths.includes(m.key)) {
                                state.selectedLedgerMonths.push(m.key);
                            }
                            label.classList.add("checked");
                        } else {
                            state.selectedLedgerMonths = state.selectedLedgerMonths.filter(k => k !== m.key);
                            label.classList.remove("checked");
                        }
                        chk.disabled = true;
                        try {
                            await callApi("setSelectedMonths", { selected: state.selectedLedgerMonths });
                            renderLedgerTable();
                        } catch (err) {
                            state.selectedLedgerMonths = previous;
                            e.target.checked = !activeChecked;
                            label.classList.toggle("checked", previous.includes(m.key));
                            alert("تعذّر حفظ خيارات الأشهر: " + err.message);
                        } finally {
                            chk.disabled = false;
                        }
                    });
```

- [ ] **Step 3: Rewrite the "select all / select none" month buttons**

Find (inside the `DOMContentLoaded` handler rewritten in Frontend Task 2 —
this block sits right after the login-check `if (cachedEmail) { ... }`):

```javascript
    // Multi-month select all/none controls
    const btnSelectAll = document.getElementById("btn-select-all-months");
    if (btnSelectAll) {
        btnSelectAll.addEventListener("click", () => {
            state.selectedLedgerMonths = state.monthsList.map(m => m.key);
            saveToLocalMemory();
            populateMonthFilters();
            renderLedgerTable();
        });
    }

    const btnSelectNone = document.getElementById("btn-select-none-months");
    if (btnSelectNone) {
        btnSelectNone.addEventListener("click", () => {
            state.selectedLedgerMonths = [];
            saveToLocalMemory();
            populateMonthFilters();
            renderLedgerTable();
        });
    }
});
```

Replace it with:

```javascript
    // Multi-month select all/none controls
    const btnSelectAll = document.getElementById("btn-select-all-months");
    if (btnSelectAll) {
        btnSelectAll.addEventListener("click", async () => {
            const next = state.monthsList.map(m => m.key);
            try {
                await callApi("setSelectedMonths", { selected: next });
                state.selectedLedgerMonths = next;
                populateMonthFilters();
                renderLedgerTable();
            } catch (err) {
                alert("تعذّر تحديد كل الأشهر: " + err.message);
            }
        });
    }

    const btnSelectNone = document.getElementById("btn-select-none-months");
    if (btnSelectNone) {
        btnSelectNone.addEventListener("click", async () => {
            try {
                await callApi("setSelectedMonths", { selected: [] });
                state.selectedLedgerMonths = [];
                populateMonthFilters();
                renderLedgerTable();
            } catch (err) {
                alert("تعذّر إلغاء تحديد الأشهر: " + err.message);
            }
        });
    }
});
```

- [ ] **Step 4: Rewrite the "save month edit" handler**

Find:

```javascript
const btnSaveMonthEdit = document.getElementById("btn-save-month-edit");
if (btnSaveMonthEdit) {
    btnSaveMonthEdit.addEventListener("click", () => {
        const oldKey = document.getElementById("edit-month-old-key").value;
        const newKey = document.getElementById("edit-month-new-key").value.trim();

        if (!newKey) {
            alert("الرجاء إدخال الاسم الجديد للشهر.");
            return;
        }
        if (oldKey === newKey) {
            closeEditMonthModal();
            return;
        }

        // Check if new month key already exists
        const exists = state.monthsList.some(m => m.key.toLowerCase() === newKey.toLowerCase());
        if (exists) {
            alert("هذا الاسم مسجل بالفعل لشهر آخر.");
            return;
        }

        // Find and update month key
        const monthObj = state.monthsList.find(m => m.key === oldKey);
        if (!monthObj) return;
        monthObj.key = newKey;

        // Update payments key for all members
        state.members.forEach(member => {
            if (member.payments[oldKey] !== undefined) {
                member.payments[newKey] = member.payments[oldKey];
                delete member.payments[oldKey];
            }
        });

        // Save changes
        saveToLocalMemory();
        closeEditMonthModal();

        // Reload filters
        populateMonthFilters();
        
        // Retain the newly updated month in filter select
        const monthSelect = document.getElementById("ledger-filter-month");
        if (monthSelect) {
            monthSelect.value = newKey;
        }

        renderLedgerTable();
        renderFamiliesTable();
        populateAddMemberFamilies();
        populateReceiptFamilies();
        renderDashboard();
    });
}
```

Replace it with:

```javascript
const btnSaveMonthEdit = document.getElementById("btn-save-month-edit");
if (btnSaveMonthEdit) {
    btnSaveMonthEdit.addEventListener("click", async () => {
        const oldKey = document.getElementById("edit-month-old-key").value;
        const newKey = document.getElementById("edit-month-new-key").value.trim();

        if (!newKey) {
            alert("الرجاء إدخال الاسم الجديد للشهر.");
            return;
        }
        if (oldKey === newKey) {
            closeEditMonthModal();
            return;
        }

        const exists = state.monthsList.some(m => m.key.toLowerCase() === newKey.toLowerCase());
        if (exists) {
            alert("هذا الاسم مسجل بالفعل لشهر آخر.");
            return;
        }

        btnSaveMonthEdit.disabled = true;
        try {
            const result = await callApi("updateMonth", { oldKey, newKey });
            state.monthsList = result.months.map(m => ({ key: m.key, id: m.id }));
            state.selectedLedgerMonths = state.selectedLedgerMonths.map(k => k === oldKey ? newKey : k);
            state.members = result.members;

            closeEditMonthModal();
            populateMonthFilters();

            const monthSelect = document.getElementById("ledger-filter-month");
            if (monthSelect) monthSelect.value = newKey;

            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
        } catch (err) {
            alert("تعذّر حفظ تعديل الشهر: " + err.message);
        } finally {
            btnSaveMonthEdit.disabled = false;
        }
    });
}
```

- [ ] **Step 5: Rewrite `window.deleteMonth`**

Find:

```javascript
window.deleteMonth = function() {
    const monthSelect = document.getElementById("ledger-filter-month");
    if (!monthSelect) return;
    const monthKey = monthSelect.value;
    if (monthKey === "all") return;

    showConfirm(`هل أنت متأكد من حذف شهر (${monthKey}) نهائياً من الصندوق؟\nسيتم إزالة كافة اشتراكات هذا الشهر المسجلة لجميع الأعضاء.`, () => {
        // 1. Remove from state.monthsList
        state.monthsList = state.monthsList.filter(m => m.key !== monthKey);

        // 2. Remove the payments field from all members
        state.members.forEach(member => {
            delete member.payments[monthKey];

            // Re-calculate member sum
            let memberTotal = 0;
            state.monthsList.forEach(m => {
                memberTotal += (member.payments[m.key] || 0);
            });
            member.sum = memberTotal;
        });

        // 3. Re-calculate family totals
        state.families.forEach(fam => {
            recalculateFamilyTotals(fam.headName);
        });

        // 4. Save and reload
        saveToLocalMemory();

        // 5. Reset month filter dropdown selection
        populateMonthFilters();
        const select = document.getElementById("ledger-filter-month");
        if (select) select.value = "all";
        
        // Hide ops controls
        const ops = document.getElementById("month-ops-controls");
        if (ops) ops.style.display = "none";

        renderLedgerTable();
        renderFamiliesTable();
        populateAddMemberFamilies();
        populateReceiptFamilies();
        renderDashboard();
    });
};
```

Replace it with:

```javascript
window.deleteMonth = function() {
    const monthSelect = document.getElementById("ledger-filter-month");
    if (!monthSelect) return;
    const monthKey = monthSelect.value;
    if (monthKey === "all") return;

    showConfirm(`هل أنت متأكد من حذف شهر (${monthKey}) نهائياً من الصندوق؟\nسيتم إزالة كافة اشتراكات هذا الشهر المسجلة لجميع الأعضاء.`, async () => {
        try {
            const result = await callApi("deleteMonth", { key: monthKey });
            state.monthsList = result.months.map(m => ({ key: m.key, id: m.id }));
            state.selectedLedgerMonths = state.selectedLedgerMonths.filter(k => k !== monthKey);
            state.members = result.members;

            populateMonthFilters();
            const select = document.getElementById("ledger-filter-month");
            if (select) select.value = "all";

            const ops = document.getElementById("month-ops-controls");
            if (ops) ops.style.display = "none";

            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
        } catch (err) {
            alert("تعذّر حذف الشهر: " + err.message);
        }
    });
};
```

- [ ] **Step 6: Verify**

```bash
node --check app.js
grep -n "renderDashboard\|saveToLocalMemory" app.js
```

Expected: `node --check` prints nothing; the `grep` prints nothing at all
— every `renderDashboard()`/`saveToLocalMemory()` reference in the whole
file has now been removed or replaced (Frontend Tasks 2, 4, and 5 covered
all six original `renderDashboard()` call sites between them).

- [ ] **Step 7: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire month CRUD and ledger month-selection to the API

Add/edit/delete month and the multi-month filter checkboxes now call
the matching backend action. This removes the last remaining
saveToLocalMemory()/renderDashboard() call sites in the file.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 6: Receipt workspace bulk payment recording

**Files:**
- Modify: `app.js` (the `btn-save-receipt-payment` click handler)

- [ ] **Step 1: Rewrite the handler**

Find:

```javascript
// Record payment from receipt generator directly
document.getElementById("btn-save-receipt-payment").addEventListener("click", () => {
    const headName = receiptSelectFamily.value;
    if (!headName) {
        alert("الرجاء اختيار العائلة أولاً.");
        return;
    }

    const checkedChks = document.querySelectorAll(".receipt-month-chk:checked");
    if (checkedChks.length === 0) {
        alert("الرجاء تحديد شهر واحد على الأقل لتسجيل عملية الدفع.");
        return;
    }

    const family = state.families.find(f => f.headName === headName);
    if (!family) {
        alert("خطأ: لم يتم العثور على سجل العائلة.");
        return;
    }

    // Collect checked months keys
    const checkedMonthsKeys = [];
    checkedChks.forEach(chk => {
        const m = state.monthsList.find(mon => mon.id == chk.value || mon.key == chk.value);
        if (m) checkedMonthsKeys.push(m.key);
    });

    let updatedCount = 0;
    state.members.forEach(m => {
        // Matches if it's the family head himself or parent equals family head
        const matchesFamily = (m.parent && m.parent.trim().toLowerCase() === headName.trim().toLowerCase()) 
                            || (m.name.trim().toLowerCase() === headName.trim().toLowerCase());
        
        if (matchesFamily) {
            checkedMonthsKeys.forEach(mKey => {
                m.payments[mKey] = 10; // set 10 shikel payment
            });
            // Recalculate member total
            let memberTotal = 0;
            state.monthsList.forEach(mon => {
                memberTotal += (m.payments[mon.key] || 0);
            });
            m.sum = memberTotal;
            updatedCount++;
        }
    });

    // Recalculate family totalPaid
    recalculateFamilyTotals(headName);

    // Save state to localStorage
    saveToLocalMemory();

    // Trigger complete dashboards and charts refresh
    renderCharts();

    // Show success dialog
    alert(`✅ تم بنجاح تسجيل دفعة العائلة: (${headName}) لشهور: [${checkedMonthsKeys.join("، ")}]. تم تحديث حسابات (${updatedCount}) أفراد وتحديث الداشبورد والمخططات البيانية!`);

    // Reset input fields
    checkedChks.forEach(chk => chk.checked = false);
    receiptAmount.value = 0;
    receiptNotes.value = "";
    
    // Update preview with empty state
    updatePrintPreview();
});
```

Replace it with:

```javascript
// Record payment from receipt generator directly
document.getElementById("btn-save-receipt-payment").addEventListener("click", async () => {
    const headName = receiptSelectFamily.value;
    if (!headName) {
        alert("الرجاء اختيار العائلة أولاً.");
        return;
    }

    const checkedChks = document.querySelectorAll(".receipt-month-chk:checked");
    if (checkedChks.length === 0) {
        alert("الرجاء تحديد شهر واحد على الأقل لتسجيل عملية الدفع.");
        return;
    }

    const family = state.families.find(f => f.headName === headName);
    if (!family) {
        alert("خطأ: لم يتم العثور على سجل العائلة.");
        return;
    }

    const checkedMonthsKeys = [];
    checkedChks.forEach(chk => {
        const m = state.monthsList.find(mon => mon.id == chk.value || mon.key == chk.value);
        if (m) checkedMonthsKeys.push(m.key);
    });

    const btn = document.getElementById("btn-save-receipt-payment");
    btn.disabled = true;
    try {
        const result = await callApi("setFamilyPayments", { headName, months: checkedMonthsKeys });
        state.members = result.members;
        state.families = result.families;

        renderCharts();

        alert(`✅ تم بنجاح تسجيل دفعة العائلة: (${headName}) لشهور: [${checkedMonthsKeys.join("، ")}]. تم تحديث الداشبورد والمخططات البيانية!`);

        checkedChks.forEach(chk => chk.checked = false);
        receiptAmount.value = 0;
        receiptNotes.value = "";
        updatePrintPreview();
    } catch (err) {
        alert("تعذّر تسجيل الدفعة: " + err.message);
    } finally {
        btn.disabled = false;
    }
});
```

- [ ] **Step 2: Verify**

```bash
node --check app.js
```

Expected: no output.

- [ ] **Step 3: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire receipt workspace bulk payments to the API

The receipt tab's "save payment" button now calls setFamilyPayments
instead of mutating every matching member locally.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 7: Expenses — add, delete, and the attachment viewer

**Files:**
- Modify: `app.js` (`window.viewAttachment`, `window.deleteExpense`, the
  add-expense click handler)

The backend's `getExpenses()` returns `attachment: {name, type, url}` — a
Drive URL — where the old client shape was `attachment: {name, type, size,
data}` with `data` as an inline base64 string. `viewAttachment` and the
download button both read `.data` today; both need to read `.url` instead.

- [ ] **Step 1: Fix `window.viewAttachment`'s image/PDF src and download button**

Find:

```javascript
    title.textContent = `عرض المرفق: ${expense.attachment.name}`;
    body.innerHTML = "";

    const type = expense.attachment.type;
    const data = expense.attachment.data;

    if (type.startsWith("image/")) {
        const image = document.createElement("img");
        image.src = data;
        image.alt = expense.attachment.name;
        image.style.maxWidth = "100%";
        image.style.maxHeight = "400px";
        image.style.borderRadius = "8px";
        image.style.boxShadow = "var(--shadow)";
        body.appendChild(image);
    } else if (type === "application/pdf") {
        const embed = document.createElement("object");
        embed.data = data;
        embed.type = "application/pdf";
        embed.style.width = "100%";
        embed.style.height = "380px";
        
        const fallbackText = document.createElement("div");
        fallbackText.style.textAlign = "center";
        fallbackText.style.padding = "20px";
        fallbackText.innerHTML = `
            <p style="margin-bottom:12px; color:var(--text-muted);">لا تدعم المعاينة الفورية لـ PDF في هذا المتصفح.</p>
            <a href="${data}" download="${expense.attachment.name}" class="btn btn-primary" style="font-size:0.85rem;">📥 تحميل ملف الـ PDF مباشرة</a>
        `;
        embed.appendChild(fallbackText);
        body.appendChild(embed);
    } else {
        const infoDiv = document.createElement("div");
        infoDiv.style.textAlign = "center";
        infoDiv.style.padding = "30px 20px";
        infoDiv.innerHTML = `
            <div style="font-size:3rem; margin-bottom:10px;">📄</div>
            <h4 style="font-size:1.05rem; color:var(--text-main); margin-bottom:6px;">مستند خارجي: ${expense.attachment.name}</h4>
            <p style="font-size:0.82rem; color:var(--text-muted); margin-bottom:14px;">لا تتوفر معاينة مباشرة لهذا النوع من الملفات (${type}).</p>
            <p style="font-weight:600; color:var(--primary); font-size:0.85rem;">انقر فوق زر التحميل أدناه لحفظ المستند على جهازك.</p>
        `;
        body.appendChild(infoDiv);
    }

    // Download action hook
    downloadBtn.onclick = function() {
        const link = document.createElement("a");
        link.href = data;
        link.download = expense.attachment.name;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    modal.style.display = "flex";
};
```

Replace it with:

```javascript
    title.textContent = `عرض المرفق: ${expense.attachment.name}`;
    body.innerHTML = "";

    const type = expense.attachment.type;
    const url = expense.attachment.url;

    if (type.startsWith("image/")) {
        const image = document.createElement("img");
        image.src = url;
        image.alt = expense.attachment.name;
        image.style.maxWidth = "100%";
        image.style.maxHeight = "400px";
        image.style.borderRadius = "8px";
        image.style.boxShadow = "var(--shadow)";
        body.appendChild(image);
    } else if (type === "application/pdf") {
        const embed = document.createElement("object");
        embed.data = url;
        embed.type = "application/pdf";
        embed.style.width = "100%";
        embed.style.height = "380px";
        
        const fallbackText = document.createElement("div");
        fallbackText.style.textAlign = "center";
        fallbackText.style.padding = "20px";
        fallbackText.innerHTML = `
            <p style="margin-bottom:12px; color:var(--text-muted);">لا تدعم المعاينة الفورية لـ PDF في هذا المتصفح.</p>
            <a href="${url}" target="_blank" rel="noopener" class="btn btn-primary" style="font-size:0.85rem;">📥 فتح ملف الـ PDF في تبويب جديد</a>
        `;
        embed.appendChild(fallbackText);
        body.appendChild(embed);
    } else {
        const infoDiv = document.createElement("div");
        infoDiv.style.textAlign = "center";
        infoDiv.style.padding = "30px 20px";
        infoDiv.innerHTML = `
            <div style="font-size:3rem; margin-bottom:10px;">📄</div>
            <h4 style="font-size:1.05rem; color:var(--text-main); margin-bottom:6px;">مستند خارجي: ${expense.attachment.name}</h4>
            <p style="font-size:0.82rem; color:var(--text-muted); margin-bottom:14px;">لا تتوفر معاينة مباشرة لهذا النوع من الملفات (${type}).</p>
            <p style="font-weight:600; color:var(--primary); font-size:0.85rem;">انقر فوق زر التحميل أدناه لفتح المستند.</p>
        `;
        body.appendChild(infoDiv);
    }

    // Download action hook — url is a cross-origin Drive link, so a
    // download-attribute anchor won't force a save-as; open it instead.
    downloadBtn.onclick = function() {
        window.open(url, "_blank", "noopener");
    };

    modal.style.display = "flex";
};
```

- [ ] **Step 2: Rewrite `window.deleteExpense`**

Find:

```javascript
window.deleteExpense = function(id) {
    showConfirm("هل أنت متأكد من حذف أمر الصرف هذا؟ سيتم استعادة المبلغ للرصيد الصافي.", () => {
        state.expenses = state.expenses.filter(e => String(e.id) !== String(id));
        saveToLocalMemory();
        renderExpensesTable();
        updateIndicators();
    });
};
```

Replace it with:

```javascript
window.deleteExpense = function(id) {
    showConfirm("هل أنت متأكد من حذف أمر الصرف هذا؟ سيتم استعادة المبلغ للرصيد الصافي.", async () => {
        try {
            state.expenses = await callApi("deleteExpense", { id });
            renderExpensesTable();
            updateIndicators();
        } catch (err) {
            alert("تعذّر حذف أمر الصرف: " + err.message);
        }
    });
};
```

- [ ] **Step 3: Rewrite the "add expense" click handler**

Find:

```javascript
    const newExpense = {
        id: "exp_" + Date.now(),
        date: dateVal,
        amount: amountVal,
        reason: reasonVal,
        category: categoryVal,
        authorized: authorizedVal || (state.currentUser ? state.currentUser.name : "اللجنة المالية"),
        attachment: attachmentObj
    };

    state.expenses.push(newExpense);
    saveToLocalMemory();

    // Clear form (keep date and authorized)
    document.getElementById("exp-amount").value = "";
    document.getElementById("exp-reason").value = "";
    if (fileInput) fileInput.value = ""; // Clear file selector

    renderExpensesTable();
    updateIndicators();

    alert(`✅ تم تسجيل أمر الصرف بنجاح!\nالمبلغ: ${amountVal} شيكل\nالسبب: ${reasonVal}\nالرصيد الصافي الجديد: ${netBalance - amountVal} شيكل`);
});
```

Replace it with:

```javascript
    const btn = document.getElementById("btn-add-expense");
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `<span style="display:inline-block; width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; margin-left:8px; vertical-align:middle;"></span> جاري الحفظ على جوجل شيت...`;

    try {
        state.expenses = await callApi("addExpense", {
            date: dateVal,
            amount: amountVal,
            reason: reasonVal,
            category: categoryVal,
            authorized: authorizedVal || (state.currentUser ? state.currentUser.name : "اللجنة المالية"),
            attachment: attachmentObj
        });

        document.getElementById("exp-amount").value = "";
        document.getElementById("exp-reason").value = "";
        if (fileInput) fileInput.value = "";

        renderExpensesTable();
        updateIndicators();

        alert(`✅ تم تسجيل أمر الصرف بنجاح!\nالمبلغ: ${amountVal} شيكل\nالسبب: ${reasonVal}\nالرصيد الصافي الجديد: ${netBalance - amountVal} شيكل`);
    } catch (err) {
        alert("تعذّر تسجيل أمر الصرف: " + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
});
```

Note: this block sits inside the same `btn-add-expense` click handler whose
opening line already reads
`document.getElementById("btn-add-expense").addEventListener("click", async () => {`
— it was already `async` (for the pre-existing `await readAttachmentFile(file)`
call above this block), so no signature change is needed here.

- [ ] **Step 4: Verify**

```bash
node --check app.js
```

Expected: no output.

- [ ] **Step 5: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire expense CRUD and attachments to the API

Add/delete expense now call the backend, which uploads any attachment
to Drive and returns a URL. viewAttachment and the download button
are updated to use attachment.url instead of the old inline
attachment.data base64 string.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 8: Signatures — save, clear, display

**Files:**
- Modify: `app.js` (`saveSignature`, `window.clearSignature`,
  `window.updateSignaturesDisplay`)

Signatures are fetched per-month on demand (not part of `getAllData`), so
`updateSignaturesDisplay` becomes `async`. It's called from two places —
`renderMonthlyReport()` (already reads `if (window.updateSignaturesDisplay)
window.updateSignaturesDisplay();`, which works fine as a fire-and-forget
call against an async function) and the signature save/clear handlers
below.

- [ ] **Step 1: Rewrite `saveSignature`**

Find:

```javascript
function saveSignature() {
    if (!currentSigIndex) return;
    
    // Check if canvas has drawing in it
    const blank = document.createElement("canvas");
    blank.width = sigCanvas.width;
    blank.height = sigCanvas.height;
    if (sigCanvas.toDataURL() === blank.toDataURL()) {
        alert("الرجاء رسم التوقيع أولاً قبل الحفظ.");
        return;
    }
    
    const dataUrl = sigCanvas.toDataURL();
    
    // Save signature under monthly key
    const monthVal = document.getElementById("report-select-month").value;
    const savedSigsKey = `cems_report_sigs_${monthVal}`;
    const sigs = JSON.parse(localStorage.getItem(savedSigsKey) || "{}");
    sigs[currentSigIndex] = dataUrl;
    localStorage.setItem(savedSigsKey, JSON.stringify(sigs));
    
    window.updateSignaturesDisplay();
    closeSignatureModal();
}
```

Replace it with:

```javascript
async function saveSignature() {
    if (!currentSigIndex) return;

    const blank = document.createElement("canvas");
    blank.width = sigCanvas.width;
    blank.height = sigCanvas.height;
    if (sigCanvas.toDataURL() === blank.toDataURL()) {
        alert("الرجاء رسم التوقيع أولاً قبل الحفظ.");
        return;
    }

    const dataUrl = sigCanvas.toDataURL();
    const monthVal = document.getElementById("report-select-month").value;

    try {
        await callApi("saveSignature", { month: monthVal, slotIndex: currentSigIndex, data: dataUrl });
        await window.updateSignaturesDisplay();
        closeSignatureModal();
    } catch (err) {
        alert("تعذّر حفظ التوقيع: " + err.message);
    }
}
```

- [ ] **Step 2: Rewrite `window.clearSignature`**

Find:

```javascript
window.clearSignature = function(index) {
    showConfirm("هل أنت متأكد من إزالة هذا التوقيع الرقمي؟", () => {
        const monthVal = document.getElementById("report-select-month").value;
        const savedSigsKey = `cems_report_sigs_${monthVal}`;
        const sigs = JSON.parse(localStorage.getItem(savedSigsKey) || "{}");
        delete sigs[index];
        localStorage.setItem(savedSigsKey, JSON.stringify(sigs));
        window.updateSignaturesDisplay();
    });
};
```

Replace it with:

```javascript
window.clearSignature = function(index) {
    showConfirm("هل أنت متأكد من إزالة هذا التوقيع الرقمي؟", async () => {
        const monthVal = document.getElementById("report-select-month").value;
        try {
            await callApi("clearSignature", { month: monthVal, slotIndex: index });
            await window.updateSignaturesDisplay();
        } catch (err) {
            alert("تعذّر إزالة التوقيع: " + err.message);
        }
    });
};
```

- [ ] **Step 3: Rewrite `window.updateSignaturesDisplay`**

Find:

```javascript
window.updateSignaturesDisplay = function() {
    const reportSelect = document.getElementById("report-select-month");
    if (!reportSelect) return;
    const monthVal = reportSelect.value;
    const savedSigsKey = `cems_report_sigs_${monthVal}`;
    const sigs = JSON.parse(localStorage.getItem(savedSigsKey) || "{}");
    
    for (let i = 1; i <= 5; i++) {
```

Replace it with:

```javascript
window.updateSignaturesDisplay = async function() {
    const reportSelect = document.getElementById("report-select-month");
    if (!reportSelect) return;
    const monthVal = reportSelect.value;
    let sigs = {};
    try {
        sigs = await callApi("getSignatures", { month: monthVal });
    } catch (err) {
        console.error("Could not load signatures:", err);
    }

    for (let i = 1; i <= 5; i++) {
```

- [ ] **Step 4: Verify**

```bash
node --check app.js
grep -n "cems_report_sigs" app.js
```

Expected: `node --check` prints nothing; the `grep` prints nothing (every
`cems_report_sigs_*` localStorage reference is gone).

- [ ] **Step 5: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire signature pad to the API

Save/clear/display now go through the backend's Signatures actions,
which upload the signature PNG to Drive and store its URL, instead of
localStorage keyed by report month.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 9: User management — pending requests, active users, add supervisor

**Files:**
- Modify: `app.js` (`updatePendingBadge`, `renderPendingUsers`,
  `window.approveUser`, `window.rejectUser`, `renderActiveUsers`,
  `window.toggleAdminStatus`, `window.toggleDynamicUserStatus`,
  `window.deleteDynamicUser`, the `btn-add-supervisor` handler)

The old pending-request records had no `id` field (matched by `email`
instead); the backend's `PendingUsers` rows always have a generated `id`.
This task switches every pending-user action to address rows by `id`.

- [ ] **Step 1: Rewrite `updatePendingBadge` and `renderPendingUsers`**

Find:

```javascript
function updatePendingBadge() {
    const pendingList = JSON.parse(localStorage.getItem("cems_pending_users") || "[]");
    const badge = document.getElementById("pending-badge");
    const label = document.getElementById("pending-count-label");
    
    if (badge) {
        if (pendingList.length > 0) {
            badge.textContent = pendingList.length;
            badge.style.display = "flex";
        } else {
            badge.style.display = "none";
        }
    }
    
    if (label) {
        label.textContent = `${pendingList.length} طلب${pendingList.length === 1 ? '' : 'ات'}`;
    }
}

function renderPendingUsers() {
    const tbody = document.getElementById("pending-user-rows");
    if (!tbody) return;
    
    tbody.innerHTML = "";
    const pendingList = JSON.parse(localStorage.getItem("cems_pending_users") || "[]");
    
    if (pendingList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted);">لا توجد طلبات تسجيل معلّقة.</td></tr>`;
        return;
    }
    
    pendingList.forEach((req, idx) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td><strong>${req.name}</strong></td>
            <td>${req.email}</td>
            <td>${req.role}</td>
            <td>${req.date || "-"}</td>
            <td>
                <div style="display: flex; gap: 8px;">
                    <button class="btn btn-primary btn-sm" onclick="approveUser('${req.email}')">✔️ موافقة</button>
                    <button class="btn btn-danger btn-sm" onclick="rejectUser('${req.email}')">❌ رفض</button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
}
```

Replace it with:

```javascript
function updatePendingBadge() {
    const badge = document.getElementById("pending-badge");
    const label = document.getElementById("pending-count-label");
    const count = state.pendingUsers.length;

    if (badge) {
        badge.textContent = count;
        badge.style.display = count > 0 ? "flex" : "none";
    }
    if (label) {
        label.textContent = `${count} طلب${count === 1 ? '' : 'ات'}`;
    }
}

function renderPendingUsers() {
    const tbody = document.getElementById("pending-user-rows");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (state.pendingUsers.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted);">لا توجد طلبات تسجيل معلّقة.</td></tr>`;
        return;
    }

    state.pendingUsers.forEach((req, idx) => {
        const tr = document.createElement("tr");
        const name = `${req.firstName} ${req.lastName}`;
        const requestedDate = req.requestedAt ? req.requestedAt.split("T")[0] : "-";
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td><strong>${name}</strong></td>
            <td>${req.email}</td>
            <td>مشرف مالي</td>
            <td>${requestedDate}</td>
            <td>
                <div style="display: flex; gap: 8px;">
                    <button class="btn btn-primary btn-sm" onclick="approveUser('${req.id}')">✔️ موافقة</button>
                    <button class="btn btn-danger btn-sm" onclick="rejectUser('${req.id}')">❌ رفض</button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
}
```

- [ ] **Step 2: Rewrite `window.approveUser` and `window.rejectUser`**

Find:

```javascript
window.approveUser = function(email) {
    const pendingList = JSON.parse(localStorage.getItem("cems_pending_users") || "[]");
    const userIdx = pendingList.findIndex(u => u.email.toLowerCase() === email.toLowerCase());
    if (userIdx === -1) return;
    
    const user = pendingList[userIdx];
    pendingList.splice(userIdx, 1);
    localStorage.setItem("cems_pending_users", JSON.stringify(pendingList));
    
    const dynamicUsers = JSON.parse(localStorage.getItem("cems_dynamic_users") || "[]");
    dynamicUsers.push({
        name: user.name,
        email: user.email,
        role: user.role,
        pass: user.pass,
        firstLoginDone: true,
        isActive: true
    });
    
    localStorage.setItem("cems_dynamic_users", JSON.stringify(dynamicUsers));
    
    alert(`✔️ تم تفعيل حساب المشرف: ${user.name} بنجاح!`);
    updatePendingBadge();
    renderPendingUsers();
    renderActiveUsers();
};

window.rejectUser = function(email) {
    const pendingList = JSON.parse(localStorage.getItem("cems_pending_users") || "[]");
    const userIdx = pendingList.findIndex(u => u.email.toLowerCase() === email.toLowerCase());
    if (userIdx === -1) return;
    
    const user = pendingList[userIdx];
    if (confirm(`هل أنت متأكد من رفض طلب تسجيل المشرف: ${user.name}؟`)) {
        pendingList.splice(userIdx, 1);
        localStorage.setItem("cems_pending_users", JSON.stringify(pendingList));
        updatePendingBadge();
        renderPendingUsers();
    }
};
```

Replace it with:

```javascript
window.approveUser = async function(id) {
    const user = state.pendingUsers.find(u => u.id === id);
    if (!user) return;
    const name = `${user.firstName} ${user.lastName}`;

    try {
        const result = await callApi("approveUser", { id });
        state.users = result.users;
        state.pendingUsers = result.pendingUsers;

        alert(`✔️ تم تفعيل حساب المشرف: ${name} بنجاح!`);
        updatePendingBadge();
        renderPendingUsers();
        renderActiveUsers();
    } catch (err) {
        alert("تعذّرت الموافقة على الطلب: " + err.message);
    }
};

window.rejectUser = function(id) {
    const user = state.pendingUsers.find(u => u.id === id);
    if (!user) return;
    const name = `${user.firstName} ${user.lastName}`;

    if (!confirm(`هل أنت متأكد من رفض طلب تسجيل المشرف: ${name}؟`)) return;

    callApi("rejectUser", { id }).then(pendingUsers => {
        state.pendingUsers = pendingUsers;
        updatePendingBadge();
        renderPendingUsers();
    }).catch(err => {
        alert("تعذّر رفض الطلب: " + err.message);
    });
};
```

- [ ] **Step 3: Rewrite `renderActiveUsers`**

Find:

```javascript
function renderActiveUsers() {
    const tbody = document.getElementById("active-user-rows");
    if (!tbody) return;
    
    tbody.innerHTML = "";
    let index = 1;
    
    Object.keys(ADMINS).forEach(key => {
        const admin = ADMINS[key];
        const tr = document.createElement("tr");
        const isMaster = admin.email.toLowerCase() === MASTER_ADMIN_EMAIL.toLowerCase();
        
        tr.innerHTML = `
            <td>${index++}</td>
            <td><strong>${admin.name}</strong></td>
            <td>${admin.email}</td>
            <td>${admin.role}</td>
            <td><span class="badge badge-success">أساسي</span></td>
            <td>
                ${isMaster ? '<span style="color: var(--text-muted); font-size: 0.82rem;">وصول كامل للوحة</span>' : `
                    <button class="btn btn-secondary btn-sm" onclick="toggleAdminStatus('${key}')">
                        ${admin.isActive !== false ? '❌ تعطيل' : '✔️ تفعيل'}
                    </button>
                `}
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    const dynamicUsers = JSON.parse(localStorage.getItem("cems_dynamic_users") || "[]");
    dynamicUsers.forEach((user, idx) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${index++}</td>
            <td><strong>${user.name}</strong></td>
            <td>${user.email}</td>
            <td>${user.role}</td>
            <td><span class="badge ${user.isActive ? 'badge-success' : 'badge-danger'}">${user.isActive ? 'نشط' : 'معطل'}</span></td>
            <td>
                <div style="display: flex; gap: 8px;">
                    <button class="btn btn-secondary btn-sm" onclick="toggleDynamicUserStatus('${user.email}')">
                        ${user.isActive ? '❌ تعطيل' : '✔️ تفعيل'}
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteDynamicUser('${user.email}')">🗑️ حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    const totalCount = Object.keys(ADMINS).length + dynamicUsers.length;
    const activeCountLabel = document.getElementById("active-users-count");
    if (activeCountLabel) {
        activeCountLabel.textContent = `${totalCount} مستخدم`;
    }
}
```

Replace it with:

```javascript
function renderActiveUsers() {
    const tbody = document.getElementById("active-user-rows");
    if (!tbody) return;

    tbody.innerHTML = "";

    state.users.forEach((user, idx) => {
        const tr = document.createElement("tr");
        let actionsHtml;
        if (user.isMaster) {
            actionsHtml = '<span style="color: var(--text-muted); font-size: 0.82rem;">وصول كامل للوحة</span>';
        } else if (user.isBuiltIn) {
            actionsHtml = `
                <button class="btn btn-secondary btn-sm" onclick="toggleUserStatus('${user.key}')">
                    ${user.isActive ? '❌ تعطيل' : '✔️ تفعيل'}
                </button>
            `;
        } else {
            actionsHtml = `
                <div style="display: flex; gap: 8px;">
                    <button class="btn btn-secondary btn-sm" onclick="toggleUserStatus('${user.key}')">
                        ${user.isActive ? '❌ تعطيل' : '✔️ تفعيل'}
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteAppUser('${user.key}')">🗑️ حذف</button>
                </div>
            `;
        }

        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td><strong>${user.name}</strong></td>
            <td>${user.email}</td>
            <td>${user.role}</td>
            <td><span class="badge ${user.isActive ? 'badge-success' : 'badge-danger'}">${user.isBuiltIn ? (user.isActive ? 'أساسي' : 'معطل') : (user.isActive ? 'نشط' : 'معطل')}</span></td>
            <td>${actionsHtml}</td>
        `;
        tbody.appendChild(tr);
    });

    const activeCountLabel = document.getElementById("active-users-count");
    if (activeCountLabel) {
        activeCountLabel.textContent = `${state.users.length} مستخدم`;
    }
}
```

- [ ] **Step 4: Replace the three per-account-type toggle/delete functions with unified ones**

Find:

```javascript
window.toggleAdminStatus = function(key) {
    const admin = ADMINS[key];
    if (!admin) return;
    
    const currentIsActive = admin.isActive !== false;
    const nextIsActive = !currentIsActive;
    
    showConfirm(`هل أنت متأكد من ${nextIsActive ? 'تفعيل' : 'تعطيل'} حساب المشرف المالي: ${admin.name}؟`, () => {
        admin.isActive = nextIsActive;
        localStorage.setItem(`cems_admin_active_${key}`, nextIsActive ? 'true' : 'false');
        renderActiveUsers();
    });
};

window.toggleDynamicUserStatus = function(email) {
    const dynamicUsers = JSON.parse(localStorage.getItem("cems_dynamic_users") || "[]");
    const userIdx = dynamicUsers.findIndex(u => u.email.toLowerCase() === email.toLowerCase());
    if (userIdx === -1) return;
    
    const user = dynamicUsers[userIdx];
    const nextIsActive = !user.isActive;
    
    showConfirm(`هل أنت متأكد من ${nextIsActive ? 'تفعيل' : 'تعطيل'} حساب المشرف المالي: ${user.name}؟`, () => {
        user.isActive = nextIsActive;
        localStorage.setItem("cems_dynamic_users", JSON.stringify(dynamicUsers));
        renderActiveUsers();
    });
};

window.deleteDynamicUser = function(email) {
    const dynamicUsers = JSON.parse(localStorage.getItem("cems_dynamic_users") || "[]");
    const userIdx = dynamicUsers.findIndex(u => u.email.toLowerCase() === email.toLowerCase());
    if (userIdx === -1) return;
    
    const user = dynamicUsers[userIdx];
    showConfirm(`⚠️ تحذير: هل أنت متأكد من الحذف النهائي لحساب المشرف: ${user.name}؟ لن يتمكن من تسجيل الدخول بعد الآن.`, () => {
        dynamicUsers.splice(userIdx, 1);
        localStorage.setItem("cems_dynamic_users", JSON.stringify(dynamicUsers));
        alert("🗑️ تم حذف حساب المشرف بنجاح.");
        renderActiveUsers();
    });
};
```

Replace it with:

```javascript
window.toggleUserStatus = function(key) {
    const user = state.users.find(u => u.key === key);
    if (!user) return;
    const nextIsActive = !user.isActive;

    showConfirm(`هل أنت متأكد من ${nextIsActive ? 'تفعيل' : 'تعطيل'} حساب المشرف المالي: ${user.name}؟`, async () => {
        try {
            state.users = await callApi("setUserActive", { key, active: nextIsActive });
            renderActiveUsers();
        } catch (err) {
            alert("تعذّر تغيير حالة الحساب: " + err.message);
        }
    });
};

window.deleteAppUser = function(key) {
    const user = state.users.find(u => u.key === key);
    if (!user) return;

    showConfirm(`⚠️ تحذير: هل أنت متأكد من الحذف النهائي لحساب المشرف: ${user.name}؟ لن يتمكن من تسجيل الدخول بعد الآن.`, async () => {
        try {
            state.users = await callApi("deleteUser", { key });
            alert("🗑️ تم حذف حساب المشرف بنجاح.");
            renderActiveUsers();
        } catch (err) {
            alert("تعذّر حذف الحساب: " + err.message);
        }
    });
};
```

- [ ] **Step 5: Rewrite the "add supervisor" handler**

Find:

```javascript
        if (findUserByEmail(email)) {
            alert("هذا البريد الإلكتروني مسجل بالفعل لمشرف آخر.");
            return;
        }

        const dynamicUsers = JSON.parse(localStorage.getItem("cems_dynamic_users") || "[]");
        dynamicUsers.push({
            name,
            email,
            role,
            pass,
            firstLoginDone: true,
            isActive: true
        });
        localStorage.setItem("cems_dynamic_users", JSON.stringify(dynamicUsers));

        alert(`✅ تم إضافة المشرف المالي: ${name} وتفعيل حسابه فوراً وبنجاح!`);
        
        document.getElementById("su-name").value = "";
        document.getElementById("su-email").value = "";
        document.getElementById("su-role").value = "";
        document.getElementById("su-pass").value = "";

        renderActiveUsers();
    });
}
```

Replace it with:

```javascript
        if (findUserInState_(email)) {
            alert("هذا البريد الإلكتروني مسجل بالفعل لمشرف آخر.");
            return;
        }

        try {
            state.users = await callApi("addSupervisor", { name, email, role, pass });

            alert(`✅ تم إضافة المشرف المالي: ${name} وتفعيل حسابه فوراً وبنجاح!`);

            document.getElementById("su-name").value = "";
            document.getElementById("su-email").value = "";
            document.getElementById("su-role").value = "";
            document.getElementById("su-pass").value = "";

            renderActiveUsers();
        } catch (err) {
            alert("تعذّرت إضافة المشرف: " + err.message);
        }
    });
}
```

Also find the opening line of this same handler (a few lines above the
`Find` block in Step 5) and make it `async`:

```javascript
    btnAddSupervisor.addEventListener("click", () => {
```

Replace it with:

```javascript
    btnAddSupervisor.addEventListener("click", async () => {
```

- [ ] **Step 6: Verify**

```bash
node --check app.js
grep -n "cems_pending_users\|cems_dynamic_users\|ADMINS\b\|MASTER_ADMIN_EMAIL\b" app.js
```

Expected: `node --check` prints nothing; the `grep` prints nothing — every
last reference to the old localStorage-backed account system is gone from
`app.js`.

- [ ] **Step 7: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire user management to the API, remove ADMINS entirely

Pending-request approve/reject now address rows by id instead of
email. Active-user toggle/delete are unified across built-in and
registered accounts (toggleUserStatus/deleteAppUser) since both kinds
now live in the same Users sheet. This removes the last references to
ADMINS, MASTER_ADMIN_EMAIL, and the cems_pending_users/
cems_dynamic_users localStorage keys.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 10: Recalculate-everything safety net

**Files:**
- Modify: `app.js` (`doRecalculateAll`)

- [ ] **Step 1: Rewrite `doRecalculateAll`**

Find:

```javascript
function doRecalculateAll() {
    // 1. Recalculate each member's total payments (sum)
    state.members.forEach(member => {
        let sum = 0;
        state.monthsList.forEach(m => {
            sum += (member.payments[m.key] || 0);
        });
        member.sum = sum;
    });

    // 2. Rebuild/recalculate families summaries to match member sums with Arabic normalization
    state.families.forEach(f => {
        const normHead = normalizeArabicName(f.headName);
        let familySum = 0;
        state.members.forEach(m => {
            if (normalizeArabicName(m.parent) === normHead) {
                familySum += m.sum;
            }
        });
        f.totalPaid = familySum;
    });

    // 3. Save to localStorage to persist
    saveToLocalMemory();

    // 4. Update core UI displays
    updateIndicators();
    renderCharts();

    // 5. Update active view tables
    const activeLink = document.querySelector(".nav-link.active");
    if (activeLink) {
        const activeTab = activeLink.getAttribute("data-tab");
        if (activeTab === "ledger-tab") {
            renderLedgerTable();
        } else if (activeTab === "families-tab") {
            renderFamiliesTable();
        } else if (activeTab === "receipt-tab") {
            populateReceiptFamilies();
        } else if (activeTab === "reports-tab") {
            renderMonthlyReport();
        } else if (activeTab === "expenses-tab") {
            renderExpensesTable();
        }
    }
}

const btnRecalculateAll = document.getElementById("btn-recalculate-all");
if (btnRecalculateAll) {
    btnRecalculateAll.addEventListener("click", () => {
        doRecalculateAll();
        alert("✅ تم إعادة عملية فحص ومطابقة جميع الحسابات وتحديث أرقام الميزانية والمصروفات بنجاح!");
    });
}
```

Replace it with:

```javascript
async function doRecalculateAll() {
    const result = await callApi("recalculateEverything", {});
    state.members = result.members;
    state.families = result.families;

    updateIndicators();
    renderCharts();

    const activeLink = document.querySelector(".nav-link.active");
    if (activeLink) {
        const activeTab = activeLink.getAttribute("data-tab");
        if (activeTab === "ledger-tab") {
            renderLedgerTable();
        } else if (activeTab === "families-tab") {
            renderFamiliesTable();
        } else if (activeTab === "receipt-tab") {
            populateReceiptFamilies();
        } else if (activeTab === "reports-tab") {
            renderMonthlyReport();
        } else if (activeTab === "expenses-tab") {
            renderExpensesTable();
        }
    }
}

const btnRecalculateAll = document.getElementById("btn-recalculate-all");
if (btnRecalculateAll) {
    btnRecalculateAll.addEventListener("click", async () => {
        btnRecalculateAll.disabled = true;
        try {
            await doRecalculateAll();
            alert("✅ تم إعادة عملية فحص ومطابقة جميع الحسابات وتحديث أرقام الميزانية والمصروفات بنجاح!");
        } catch (err) {
            alert("تعذّرت إعادة الحساب: " + err.message);
        } finally {
            btnRecalculateAll.disabled = false;
        }
    });
}
```

Note `normalizeArabicName()` (the client-side copy) stays in `app.js`
untouched — it's still used by ledger/family search filtering
(`renderLedgerTable`, `renderFamiliesTable`) and by `recalculateFamilyTotals()`,
which is now dead code after Frontend Tasks 4–6 removed its only call
sites. Leave `recalculateFamilyTotals()` defined but unused rather than
deleting it in this task — Frontend Task 11's verification pass double
-checks it's truly unreachable before any final cleanup.

- [ ] **Step 2: Verify**

```bash
node --check app.js
grep -n "recalculateFamilyTotals(" app.js
```

Expected: `node --check` prints nothing. The `grep` should show exactly
one match — the function's own `function recalculateFamilyTotals(familyName) {`
definition line — confirming every call site has been migrated to the API
in earlier tasks.

- [ ] **Step 3: If `recalculateFamilyTotals` truly has zero call sites, delete it**

Find:

```javascript
function recalculateFamilyTotals(familyName) {
    if (!familyName) return;
    
    const normFamilyName = normalizeArabicName(familyName);
    let familySum = 0;
    state.members.forEach(m => {
        if (normalizeArabicName(m.parent) === normFamilyName) {
            familySum += m.sum;
        }
    });

    // Update inside family record
    const familyRecord = state.families.find(f => normalizeArabicName(f.headName) === normFamilyName);
    if (familyRecord) {
        familyRecord.totalPaid = familySum;
    }
}
```

Delete this whole function (replace it with nothing / remove the block).

- [ ] **Step 4: Verify again**

```bash
node --check app.js
```

Expected: no output.

- [ ] **Step 5: Commit**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
feat: wire recalculate-everything to the API, remove dead code

doRecalculateAll now calls recalculateEverything on the backend
instead of recomputing locally and saving to localStorage. Also
removes recalculateFamilyTotals(), left orphaned once every call site
was migrated to server-side recalculation in earlier tasks.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Frontend Task 11: Full end-to-end manual verification

**Files:** none (verification only)

This task requires a live deployment (Backend Task 14 completed by the
user) — it cannot be run by an automated agent without one. If no
deployment exists yet, stop here and report that the remaining work is
blocked on the user completing `gas/DEPLOYMENT.md`.

- [ ] **Step 1: Remove now-dead default-data code**

`DEFAULT_MONTHS` (defined directly in `app.js`, right after the
`findUserInState_` helper added in Frontend Task 2) and `DEFAULT_MEMBERS`
/ `DEFAULT_FAMILIES` (defined in `data.js`, loaded as a ~74KB `<script>` by
`index.html`) were only ever read by `loadCachedData()` and the old
`btn-clear-cache` handler — both fully replaced in Frontend Task 2. Confirm
and clean up:

```bash
grep -n "DEFAULT_MONTHS\|DEFAULT_MEMBERS\|DEFAULT_FAMILIES" app.js
```

Expected: the only hits left are each constant's own declaration line
(`DEFAULT_MONTHS` in `app.js`; `DEFAULT_MEMBERS`/`DEFAULT_FAMILIES` don't
appear in `app.js` at all, only in `data.js`). If so:

1. Delete the `const DEFAULT_MONTHS = [...]` block from `app.js` (the 12-entry
   array literal, immediately following `findUserInState_`).
2. In `index.html`, remove the `<script src="data.js"></script>` tag (or
   equivalent) that loads `data.js` — grep for `data.js` in `index.html` to
   find it.
3. Delete `data.js` from the repo (`git rm data.js`) — its ~134-member seed
   dataset now permanently lives in the Google Sheet once the user has run
   the Excel import (Frontend Task 2 / `gas/DEPLOYMENT.md` §5), not in a
   static file shipped to every browser.

Re-run `node --check app.js` — expected: no output.

- [ ] **Step 2: Static sanity pass**

```bash
node --check app.js
grep -n "localStorage\." app.js
```

Expected: `node --check` prints nothing. The only remaining
`localStorage.*` hits should be exactly these two — both intentionally
kept as session-only, non-app-data conveniences (per the design spec §6):
`localStorage.setItem("cems_logged_email", ...)`,
`localStorage.getItem("cems_logged_email")`, and
`localStorage.removeItem("cems_logged_email")`. If anything else shows up,
find which earlier task's `Find`/`Replace` was missed and go back to it.

- [ ] **Step 3: Serve and open the app**

```bash
npx serve .
```

Open the printed local URL in a browser. Confirm the login screen renders
with no console errors before logging in (an `API_URL` error alert is
expected only if Backend Task 14 hasn't been completed yet).

- [ ] **Step 4: Click through every migrated flow**

With a live `API_URL` deployed and pointed at a real Google Sheet, verify
each of these against the actual Sheet in a second browser tab open to it:

1. **Login** with one of the 5 seeded accounts (`gas/Setup.gs`'s
   `DEFAULT_USERS_`, password `ABC12345`) — succeeds, forces a password
   change on first login for a built-in admin.
2. **Excel import**: upload a roster file — Members/Families sheets
   populate; dashboard/ledger/families tables reflect it without a page
   reload.
3. **Ledger payment checkbox**: toggle a cell — the Members sheet's
   corresponding cell and `sum` column update; the Families sheet's
   `totalPaid` updates.
4. **Add / edit / delete member** — Members and Families sheets reflect
   each change, including auto-creating a family row for a brand-new
   family head name.
5. **Add / edit / delete month** — Months sheet row and Members sheet
   column both change; existing payment data survives a rename.
6. **Receipt workspace**: record a family payment for several months —
   every matching member's row updates in one action.
7. **Add / delete expense with an attachment** — Expenses sheet gets a
   new row; the attachment appears as a new file in the "CEMS-ATFIHAH
   Attachments" Drive folder, viewable from the app's attachment modal.
8. **Draw / save / clear a signature** on the monthly report — Signatures
   sheet gets a row with a Drive URL; the image renders in the report.
9. **Registration request → approve** — PendingUsers gets a row, then
   Users gets a row and PendingUsers loses it after approval; the new
   account can log in afterward.
10. **Add supervisor directly / toggle active / delete** (as the master
    admin) — Users sheet reflects each action; a deactivated account is
    correctly rejected on next login attempt.
11. **Recalculate everything** button — no visible data change (since
    everything should already be correct), confirming it's a true no-op
    safety net on a healthy dataset.
12. **Reload from Sheet** button (repurposed `btn-clear-cache`) — edit a
    cell directly in the Google Sheet UI, click the button in the app,
    confirm the app now shows that edit.
13. **Full page reload** after several of the above changes — confirms
    nothing was relying on any leftover `localStorage` state; everything
    reloads correctly from the Sheet via `loadDataFromServer()`.

- [ ] **Step 5 (offline check, no live deployment needed)**

If a live deployment genuinely isn't available in this environment, at
minimum confirm `API_URL`'s placeholder string is the only reachability
blocker — i.e. every other piece of Step 2's static check passes — and
report to the user that Steps 3–4 above are the outstanding manual
verification they need to run themselves once they've completed
`gas/DEPLOYMENT.md`.

- [ ] **Step 6: Final commit (only if any fixups were needed above)**

```bash
git add app.js
git commit -m "$(cat <<'EOF'
fix: address issues found during end-to-end verification

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

If nothing needed fixing, skip this commit — there's nothing to commit.
