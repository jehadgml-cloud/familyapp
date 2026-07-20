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
