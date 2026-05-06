// ─────────────────────────────────────────────────────────────
// Chelsea Phillips Group | Open House Schedule
// Google Apps Script Backend
// Paste this entire file into your Apps Script editor
// ─────────────────────────────────────────────────────────────
//
// COLUMN STRUCTURE (update your sheet to match):
//   A: Date
//   B: Address
//   C: Time
//   D: MLS Link
//   E: Agent Name     (auto-filled, leave blank)
//   F: Agent Email    (auto-filled, leave blank)
//   G: Claimed At     (auto-filled, leave blank)
//
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = 'Open Houses';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'get')   return getListings(e);
  if (action === 'claim') return claimSlot(e);
  return jsonResponse({ error: 'Unknown action' });
}

function getListings(e) {
  const month = parseInt(e.parameter.month);
  const year  = parseInt(e.parameter.year);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[0];
    if (!dateVal) continue;

    const date = new Date(dateVal);
    if (date.getMonth() + 1 === month && date.getFullYear() === year) {
      results.push({
        row:        i + 1,
        date:       Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        address:    row[1] || '',
        time:       row[2] || '',
        mlsLink:    row[3] || '',
        agentName:  row[4] || '',
        agentEmail: row[5] || ''
      });
    }
  }

  return jsonResponse(results);
}

function claimSlot(e) {
  const rowIndex   = parseInt(e.parameter.row);
  const agentName  = e.parameter.name  || '';
  const agentEmail = e.parameter.email || '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const currentName = sheet.getRange(rowIndex, 5).getValue();

  if (currentName && currentName.toString().trim() !== '') {
    return jsonResponse({ success: false, reason: 'already_claimed' });
  }

  sheet.getRange(rowIndex, 5).setValue(agentName);
  sheet.getRange(rowIndex, 6).setValue(agentEmail);
  sheet.getRange(rowIndex, 7).setValue(new Date());

  return jsonResponse({ success: true });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
