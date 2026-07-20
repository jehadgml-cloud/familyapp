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
  const updated = Object.assign({}, existing, {
    name: payload.name !== undefined ? payload.name : existing.name,
    parent: payload.parent !== undefined ? payload.parent : existing.parent
  });
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
