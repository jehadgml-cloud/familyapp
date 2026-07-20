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
