// ═══════════════════════════════════════════════
// Google Apps Script — Nuntă Vlad & Iana
// Lipește acest cod în Extensions → Apps Script
// ═══════════════════════════════════════════════

const SECRET_KEY = 'SCHIMBA_ACEASTA_PAROLA'; // parola pentru panoul mirilor

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    data.name || '',
    data.attendance || '',
    data.phone || '',
    data.partner || '',
    data.children || '',
    data.allergies || '',
    data.drinks || '',
    data.note || '',
    data.timestamp || new Date().toISOString()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action;
  const key = e.parameter.key;

  if (action === 'login') {
    return jsonResponse({ success: key === SECRET_KEY });
  }

  if (key !== SECRET_KEY) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }

  if (action === 'list') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonResponse({ success: true, data: [] });

    const headers = data[0];
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return jsonResponse({ success: true, data: rows });
  }

  if (action === 'delete') {
    const rowIndex = parseInt(e.parameter.row);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.deleteRow(rowIndex + 2); // +2: header + 0-index
    return jsonResponse({ success: true });
  }

  return jsonResponse({ success: false, error: 'Unknown action' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
