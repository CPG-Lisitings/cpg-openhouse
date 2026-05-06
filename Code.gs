// ─────────────────────────────────────────────────────────────
// Chelsea Phillips Group | Open House Schedule
// Google Apps Script Backend
// Paste this entire file into your Apps Script editor
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = 'Open Houses';

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'get') {
    return getListings(e);
  } else if (action === 'claim') {
    return claimSlot(e);
  }

  return jsonResponse({ error: 'Unknown action' });
}

function getListings(e) {
  const month = parseInt(e.parameter.month);
  const year  = parseInt(e.parameter.year);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();

  const results = [];

  // Row 0 is headers, start at row 1
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[0];
    if (!dateVal) continue;

    const date = new Date(dateVal);
    if (date.getMonth() + 1 === month && date.getFullYear() === year) {
      results.push({
        row:         i + 1,                                        // 1-indexed sheet row
        date:        Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        address:     row[1] || '',
        neighborhood:row[2] || '',
        time:        row[3] || '',
        price:       row[4] || '',
        agentName:   row[5] || '',
        agentEmail:  row[6] || ''
      });
    }
  }

  return jsonResponse(results);
}

function claimSlot(e) {
  const rowIndex  = parseInt(e.parameter.row);
  const agentName = e.parameter.name  || '';
  const agentEmail= e.parameter.email || '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const currentName = sheet.getRange(rowIndex, 6).getValue();

  // Prevent overwriting an already-claimed slot
  if (currentName && currentName.toString().trim() !== '') {
    return jsonResponse({ success: false, reason: 'already_claimed' });
  }

  sheet.getRange(rowIndex, 6).setValue(agentName);
  sheet.getRange(rowIndex, 7).setValue(agentEmail);
  sheet.getRange(rowIndex, 8).setValue(new Date()); // timestamp

  return jsonResponse({ success: true });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
