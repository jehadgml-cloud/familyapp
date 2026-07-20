// Settings.gs

function getSettings() {
  const sheet = getSheet_(SHEET_NAMES.SETTINGS);
  const result = {};
  sheetToObjects_(sheet).rows.forEach(function (r) { result[r.key] = r.value; });
  return result;
}

function setSetting(payload) {
  const sheet = getSheet_(SHEET_NAMES.SETTINGS);
  const data = sheetToObjects_(sheet);
  const existing = data.rows.find(function (r) { return r.key === payload.key; });
  if (existing) {
    sheet.getRange(existing._row, headerIndex_(data.headers, 'value') + 1).setValue(payload.value);
  } else {
    appendRowFromObject_(sheet, data.headers, { key: payload.key, value: payload.value });
  }
  return getSettings();
}
