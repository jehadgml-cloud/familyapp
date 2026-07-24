// PaymentRecorders.gs
// Tracks, per (member, month) cell in the ledger, which logged-in supervisor
// last checked it as paid — so the frontend can color-code each payment box
// by who filled it in. A payment that's unchecked back to 0 loses its
// attribution (nobody "recorded" an empty box).

function paymentRecorderKey_(memberId, month) {
  return memberId + '|' + month;
}

// Returns a flat map { "memberId|month": "recordedBy label" } — cheap for
// the frontend to look up per checkbox while rendering the ledger.
function getPaymentRecorders() {
  const sheet = getSheet_(SHEET_NAMES.PAYMENT_RECORDERS);
  const data = sheetToObjects_(sheet);
  const map = {};
  data.rows.forEach(function (r) {
    map[r.key] = r.recordedBy;
  });
  return map;
}

// Upserts (or clears) the attribution for one payment cell. Called
// automatically from setMemberPayment()/setFamilyPayments() — never exposed
// as its own API action.
function setPaymentRecorder_(memberId, month, recordedBy) {
  const sheet = getSheet_(SHEET_NAMES.PAYMENT_RECORDERS);
  const data = sheetToObjects_(sheet);
  const key = paymentRecorderKey_(memberId, month);
  const existing = findRowById_(data.rows, key, 'key');

  if (!recordedBy) {
    if (existing) sheet.deleteRow(existing._row);
    return;
  }

  if (existing) {
    sheet.getRange(existing._row, headerIndex_(data.headers, 'recordedBy') + 1).setValue(recordedBy);
  } else {
    appendRowFromObject_(sheet, data.headers, { key: key, memberId: memberId, month: month, recordedBy: recordedBy });
  }
}
