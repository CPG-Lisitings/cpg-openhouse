// ─────────────────────────────────────────────────────────────
// Chelsea Phillips Group | Open House Schedule
// Google Apps Script Backend
// ─────────────────────────────────────────────────────────────
//
// COLUMN STRUCTURE:
//   A: Date
//   B: Address
//   C: Time
//   D: MLS Link
//   E: Agent Name     (auto-filled by app)
//   F: Agent Email    (auto-filled by app)
//   G: Claimed At     (auto-filled by app)
//   H: Added to Cal   (auto-filled by calendar sync, leave blank)
//
// ─────────────────────────────────────────────────────────────

const SHEET_NAME    = 'Open Houses';
const CALENDAR_NAME = 'Six Bricks General';

// ── WEB APP HANDLER ───────────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'get')   return getListings(e);
  if (action === 'claim') return claimSlot(e);
  return jsonResponse({ error: 'Unknown action' });
}

// ── GET LISTINGS ──────────────────────────────────────────────

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

// ── CLAIM SLOT ────────────────────────────────────────────────

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

// ── CALENDAR SYNC ─────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Open Houses')
    .addItem('Sync to Six Bricks General Calendar', 'syncToCalendar')
    .addToUi();
}

function syncToCalendar() {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data     = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];
  const today    = new Date();
  today.setHours(0, 0, 0, 0);

  if (!calendar) {
    SpreadsheetApp.getUi().alert(
      'Calendar not found. Make sure "' + CALENDAR_NAME + '" is shared with your Google account and the name matches exactly.'
    );
    return;
  }

  let added   = 0;
  let skipped = 0;

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const dateVal    = row[0];
    const address    = row[1] || '';
    const timeStr    = row[2] || '';
    const agentName  = row[4] || '';
    const agentEmail = row[5] || '';
    const addedToCal = row[7] || '';

    if (!dateVal || !address) continue;

    const eventDate = new Date(dateVal);
    eventDate.setHours(0, 0, 0, 0);

    // Skip past dates
    if (eventDate < today) { skipped++; continue; }

    // Skip already synced rows
    if (addedToCal.toString().trim() !== '') { skipped++; continue; }

    // Build event title
    const agentPart = agentName ? ` (${agentName})` : '';
    const title     = `OH - ${address}${agentPart}`;

    // Parse time range
    const times = parseTimeRange(timeStr, eventDate);

    let event;
    if (times) {
      event = calendar.createEvent(title, times.start, times.end);
    } else {
      event = calendar.createAllDayEvent(title, eventDate);
    }

    // Invite the agent if we have their email
    if (agentEmail && agentEmail.trim() !== '') {
      try {
        event.addGuest(agentEmail.trim());
      } catch (err) {
        // Non-fatal: log but continue
        Logger.log('Could not add guest ' + agentEmail + ': ' + err.message);
      }
    }

    // Mark row as synced in column H
    sheet.getRange(i + 1, 8).setValue(
      'Yes - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy')
    );

    added++;
  }

  SpreadsheetApp.getUi().alert(
    added + ' open house(s) added to Six Bricks General.\n' +
    skipped + ' skipped (past dates or already synced).\n\n' +
    (added > 0 ? 'Agents with claimed slots have been sent a calendar invite.' : '')
  );
}

// ── PARSE TIME RANGE ──────────────────────────────────────────
// Handles "1:00 PM – 3:00 PM" or "1:00 PM - 3:00 PM"

function parseTimeRange(timeStr, baseDate) {
  if (!timeStr) return null;
  const parts = timeStr.split(/\s*[–\-]\s*/);
  if (parts.length < 2) return null;
  const start = parseTime(parts[0].trim(), baseDate);
  const end   = parseTime(parts[1].trim(), baseDate);
  if (!start || !end) return null;
  return { start, end };
}

function parseTime(str, baseDate) {
  const match = str.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!match) return null;
  let hours  = parseInt(match[1]);
  const mins = parseInt(match[2]);
  const ampm = match[3].toUpperCase();
  if (ampm === 'PM' && hours !== 12) hours += 12;
  if (ampm === 'AM' && hours === 12) hours = 0;
  const dt = new Date(baseDate);
  dt.setHours(hours, mins, 0, 0);
  return dt;
}

// ── HELPERS ───────────────────────────────────────────────────

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
