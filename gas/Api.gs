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
  setSetting: setSetting,

  getAuditLog: getAuditLog
};

// Actions that only read data — skip the script lock for these.
const READ_ONLY_ACTIONS = {
  getAllData: true, getMembers: true, getFamilies: true, getMonths: true,
  getExpenses: true, getUsers: true, getPendingUsers: true,
  getSignatures: true, getSettings: true, login: true, getAuditLog: true
};

function doPost(e) {
  let action = '';
  try {
    const body = JSON.parse(e.postData.contents);
    action = body.action;
    const payload = body.payload || {};
    const handler = Object.prototype.hasOwnProperty.call(ACTIONS, action) ? ACTIONS[action] : null;
    if (!handler) return jsonResponse_({ error: 'إجراء غير معروف: ' + action });

    if (READ_ONLY_ACTIONS[action]) {
      return jsonResponse_({ result: handler(payload) });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const result = handler(payload);
      // Every mutating action gets logged automatically here — individual
      // handlers (addMember, setMemberPayment, deleteExpense, ...) don't
      // need to know the audit log exists.
      logAction_(action, payload);
      return jsonResponse_({ result: result });
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
