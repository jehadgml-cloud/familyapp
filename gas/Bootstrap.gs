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
