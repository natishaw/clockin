// ============================================================
//  Clock In / Out — Google Apps Script Backend
//  Paste this entire file into your Apps Script editor,
//  then deploy as a Web App (Execute as: Me, Access: Anyone).
// ============================================================

// ── CONFIG ───────────────────────────────────────────────────
// The script auto-creates a sheet named SHEET_NAME if it doesn't exist.
var SHEET_NAME = 'Clock Records';

// Optional: set to true to send email summaries (configure below)
var SEND_EMAIL_SUMMARY = false;
var SUMMARY_EMAIL      = 'your@email.com'; // change if using email


// ── ENTRY POINTS ─────────────────────────────────────────────

/**
 * Handles GET requests (used to verify the deployment is live).
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Clock In/Out API is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles POST requests from the clock-in app.
 * Expected JSON body: { "name": "Alice", "action": "IN" | "OUT" }
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var name    = (payload.name   || '').toString().trim();
    var action  = (payload.action || '').toString().trim().toUpperCase();

    if (!name || (action !== 'IN' && action !== 'OUT')) {
      return jsonResponse({ status: 'error', message: 'Invalid payload.' });
    }

    var now       = new Date();
    var sheet     = getOrCreateSheet();
    var rowData   = buildRow(name, action, now);

    sheet.appendRow(rowData);
    formatLastRow(sheet);

    if (SEND_EMAIL_SUMMARY) {
      sendEmailNotification(name, action, now);
    }

    return jsonResponse({
      status    : 'ok',
      name      : name,
      action    : action,
      timestamp : Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
    });

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}


// ── HELPERS ──────────────────────────────────────────────────

/**
 * Returns (or creates) the clock records sheet.
 */
function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headers = ['Date', 'Day', 'Time', 'Name', 'Action', 'Notes'];
    sheet.appendRow(headers);

    // Style the header row
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
               .setBackground('#1A1A1A')
               .setFontColor('#FFFFFF')
               .setFontSize(11);

    // Freeze header
    sheet.setFrozenRows(1);

    // Set column widths
    sheet.setColumnWidth(1, 110); // Date
    sheet.setColumnWidth(2, 90);  // Day
    sheet.setColumnWidth(3, 90);  // Time
    sheet.setColumnWidth(4, 130); // Name
    sheet.setColumnWidth(5, 90);  // Action
    sheet.setColumnWidth(6, 200); // Notes
  }

  return sheet;
}

/**
 * Builds the row array for a clock record.
 */
function buildRow(name, action, timestamp) {
  var tz      = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(timestamp, tz, 'yyyy-MM-dd');
  var dayStr  = Utilities.formatDate(timestamp, tz, 'EEEE');
  var timeStr = Utilities.formatDate(timestamp, tz, 'HH:mm:ss');

  return [dateStr, dayStr, timeStr, name, action, ''];
}

/**
 * Applies alternating row color and action-based highlight to the last row.
 */
function formatLastRow(sheet) {
  var lastRow = sheet.getLastRow();
  var range   = sheet.getRange(lastRow, 1, 1, 6);

  // Alternating row background
  var rowBg = (lastRow % 2 === 0) ? '#F9F9F7' : '#FFFFFF';
  range.setBackground(rowBg);

  // Color the Action cell (column 5)
  var actionCell = sheet.getRange(lastRow, 5);
  var actionVal  = actionCell.getValue();
  if (actionVal === 'IN') {
    actionCell.setBackground('#E8F5EE').setFontColor('#1A6B3C').setFontWeight('bold');
  } else if (actionVal === 'OUT') {
    actionCell.setBackground('#FAE8E8').setFontColor('#7A2020').setFontWeight('bold');
  }
}

/**
 * Wraps a JS object as a JSON ContentService response.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Optional: sends a quick email notification on each clock event.
 */
function sendEmailNotification(name, action, timestamp) {
  try {
    var tz      = Session.getScriptTimeZone();
    var timeStr = Utilities.formatDate(timestamp, tz, 'EEEE, MMMM d · h:mm a');
    var subject = name + ' clocked ' + action + ' — ' + timeStr;
    var body    = name + ' clocked ' + action + ' at ' + timeStr + '.\n\nThis is an automated notification from your Clock In/Out app.';
    MailApp.sendEmail(SUMMARY_EMAIL, subject, body);
  } catch (mailErr) {
    Logger.log('Email error: ' + mailErr.message);
  }
}


// ── WEEKLY SUMMARY (optional, attach to time trigger) ────────

/**
 * Generates a weekly hours summary per employee.
 * To use: Apps Script → Triggers → Add trigger → weeklySummary → Week timer.
 */
function weeklySummary() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;

  var data    = sheet.getDataRange().getValues();
  var totals  = {}; // { name: totalMinutes }
  var lastIn  = {}; // { name: Date }

  // Skip header row (index 0)
  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var date   = row[0]; // yyyy-MM-dd string
    var time   = row[2]; // HH:mm:ss string
    var name   = row[3];
    var action = row[4];

    var dtStr  = date + ' ' + time;
    var dt     = new Date(dtStr);

    if (action === 'IN') {
      lastIn[name] = dt;
    } else if (action === 'OUT' && lastIn[name]) {
      var diff = (dt - lastIn[name]) / 60000; // minutes
      totals[name] = (totals[name] || 0) + diff;
      delete lastIn[name];
    }
  }

  // Build summary text
  var lines = ['Weekly Hours Summary\n'];
  for (var emp in totals) {
    var hrs  = Math.floor(totals[emp] / 60);
    var mins = Math.round(totals[emp] % 60);
    lines.push(emp + ': ' + hrs + 'h ' + mins + 'm');
  }

  var summaryText = lines.join('\n');
  Logger.log(summaryText);

  if (SEND_EMAIL_SUMMARY) {
    MailApp.sendEmail(SUMMARY_EMAIL, 'Weekly Hours Summary', summaryText);
  }
}
