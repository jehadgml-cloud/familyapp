// Expenses.gs

function getExpenses() {
  const sheet = getSheet_(SHEET_NAMES.EXPENSES);
  const data = sheetToObjects_(sheet);
  return data.rows.map(function (r) {
    return {
      id: r.id, date: r.date, amount: Number(r.amount) || 0, reason: r.reason,
      category: r.category, authorized: r.authorized,
      attachment: r.attachmentUrl ? { name: r.attachmentName, type: r.attachmentType, url: r.attachmentUrl } : null
    };
  });
}

// payload: { date, amount, reason, category, authorized, attachment:
// {name, type, data} | null } — attachment.data, if present, is a data-URL
// string uploaded to Drive first.
function addExpense(payload) {
  const sheet = getSheet_(SHEET_NAMES.EXPENSES);
  const data = sheetToObjects_(sheet);
  const id = 'exp_' + Date.now();

  let attachmentUrl = '', attachmentName = '', attachmentType = '';
  if (payload.attachment && payload.attachment.data) {
    attachmentUrl = uploadDataUrlToDrive_(payload.attachment, 'expense');
    attachmentName = payload.attachment.name;
    attachmentType = payload.attachment.type;
  }

  appendRowFromObject_(sheet, data.headers, {
    id: id, date: payload.date, amount: payload.amount, reason: payload.reason, category: payload.category,
    authorized: payload.authorized, attachmentName: attachmentName, attachmentType: attachmentType, attachmentUrl: attachmentUrl
  });
  return getExpenses();
}

function deleteExpense(payload) {
  const sheet = getSheet_(SHEET_NAMES.EXPENSES);
  const data = sheetToObjects_(sheet);
  const existing = findRowById_(data.rows, payload.id);
  if (!existing) throw new Error('أمر الصرف غير موجود.');
  sheet.deleteRow(existing._row);
  return getExpenses();
}
