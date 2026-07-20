// Signatures.gs

// payload: { month }. Returns { "<slotIndex>": signatureUrl, ... }.
function getSignatures(payload) {
  const sheet = getSheet_(SHEET_NAMES.SIGNATURES);
  const rows = sheetToObjects_(sheet).rows.filter(function (r) { return r.month === payload.month; });
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
  const existing = data.rows.find(function (r) { return r.month === payload.month && String(r.slotIndex) === String(payload.slotIndex); });

  if (existing) {
    writeRowFromObject_(sheet, data.headers, existing._row,
      { month: payload.month, slotIndex: payload.slotIndex, signatureUrl: url, updatedAt: new Date().toISOString() });
  } else {
    appendRowFromObject_(sheet, data.headers,
      { month: payload.month, slotIndex: payload.slotIndex, signatureUrl: url, updatedAt: new Date().toISOString() });
  }
  return getSignatures({ month: payload.month });
}

function clearSignature(payload) {
  const sheet = getSheet_(SHEET_NAMES.SIGNATURES);
  const data = sheetToObjects_(sheet);
  const existing = data.rows.find(function (r) { return r.month === payload.month && String(r.slotIndex) === String(payload.slotIndex); });
  if (existing) sheet.deleteRow(existing._row);
  return getSignatures({ month: payload.month });
}
