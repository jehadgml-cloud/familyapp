// AuditLog.gs
// Generic "who did what, and when" trail for financial transparency.
// Api.gs calls logAction_() once, automatically, after every successful
// *mutating* action (see doPost) — individual handlers never need to know
// about it. Never let a logging failure break the real operation.

// Fields that must never end up in the audit trail (passwords, raw
// attachment payloads) even though they pass through mutating actions.
const AUDIT_STRIP_FIELDS_ = ['by', 'pass', 'newPass', 'oldPass', 'password'];

function logAction_(action, payload) {
  try {
    const sheet = getSheet_(SHEET_NAMES.AUDIT_LOG);
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    const actor = (payload && payload.by) ? String(payload.by) : 'غير معروف';
    const id = 'log_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    sheet.appendRow([id, now, actor, action, summarizePayload_(payload)]);
  } catch (e) {
    // Audit logging is best-effort — a failure here must never surface to
    // the user or roll back the real action that already succeeded.
  }
}

function summarizePayload_(payload) {
  if (!payload) return '';
  const clone = {};
  Object.keys(payload).forEach(function (k) {
    if (AUDIT_STRIP_FIELDS_.indexOf(k) !== -1) return;
    clone[k] = (k === 'attachment' && payload[k]) ? '[مرفق]' : payload[k];
  });
  try {
    return JSON.stringify(clone);
  } catch (e) {
    return '';
  }
}

// Returns the most recent entries first. Capped at 500 rows so the sheet
// never becomes a performance problem for the admin panel that displays it.
function getAuditLog() {
  const sheet = getSheet_(SHEET_NAMES.AUDIT_LOG);
  const data = sheetToObjects_(sheet);
  return data.rows
    .map(function (r) {
      return { id: r.id, timestamp: r.timestamp, actor: r.actor, action: r.action, details: r.details };
    })
    .reverse()
    .slice(0, 500);
}
