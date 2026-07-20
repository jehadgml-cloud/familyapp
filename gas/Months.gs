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
