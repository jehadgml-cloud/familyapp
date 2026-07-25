// Signatures.gs

// Ignores stray leading/trailing whitespace when matching a month string —
// a common source of "signature saved but never shows up" bugs when the
// month name was typed by hand in different places with a tiny, invisible
// whitespace difference.
function normalizeMonthKey_(m) {
  return String(m || '').trim().replace(/\s+/g, ' ');
}

// payload: { month }. Returns { "<slotIndex>": signatureUrl, ... }.
function getSignatures(payload) {
  const sheet = getSheet_(SHEET_NAMES.SIGNATURES);
  const targetMonth = normalizeMonthKey_(payload.month);
  const rows = sheetToObjects_(sheet).rows.filter(function (r) { return normalizeMonthKey_(r.month) === targetMonth; });
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
  const targetMonth = normalizeMonthKey_(payload.month);
  const existing = data.rows.find(function (r) {
    return normalizeMonthKey_(r.month) === targetMonth && String(r.slotIndex) === String(payload.slotIndex);
  });

  const cleanMonth = targetMonth;

  if (existing) {
    writeRowFromObject_(sheet, data.headers, existing._row,
      { month: cleanMonth, slotIndex: payload.slotIndex, signatureUrl: url, updatedAt: new Date().toISOString() });
  } else {
    appendRowFromObject_(sheet, data.headers,
      { month: cleanMonth, slotIndex: payload.slotIndex, signatureUrl: url, updatedAt: new Date().toISOString() });
  }
  return getSignatures({ month: payload.month });
}

function clearSignature(payload) {
  const sheet = getSheet_(SHEET_NAMES.SIGNATURES);
  const data = sheetToObjects_(sheet);
  const targetMonth = normalizeMonthKey_(payload.month);
  const existing = data.rows.find(function (r) {
    return normalizeMonthKey_(r.month) === targetMonth && String(r.slotIndex) === String(payload.slotIndex);
  });
  if (existing) sheet.deleteRow(existing._row);
  return getSignatures({ month: payload.month });
}
