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
