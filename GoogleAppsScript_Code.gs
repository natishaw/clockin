// EJPM Clock Log — Google Apps Script
// Paste this in Extensions → Apps Script, then Deploy → Web App (Anyone)

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // ── CLOCK IN/OUT (existing) ──
    if (!data.action_type || data.action_type === 'clock') {
      sheet.appendRow([
        new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }),
        data.name,
        data.action
      ]);
      return ok('clocked');
    }

    // ── ADD ROW ──
    if (data.action_type === 'addRow') {
      sheet.appendRow([data.timestamp, data.name, data.action]);
      return ok('added');
    }

    // ── DELETE ROW ──
    if (data.action_type === 'deleteRow') {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.timestamp) &&
            rows[i][1] === data.name &&
            rows[i][2] === data.action) {
          sheet.deleteRow(i + 1);
          return ok('deleted');
        }
      }
      return ok('not_found');
    }

    // ── EDIT ROW ──
    if (data.action_type === 'editRow') {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.old_timestamp) &&
            rows[i][1] === data.old_name &&
            rows[i][2] === data.old_action) {
          sheet.getRange(i + 1, 1).setValue(data.new_timestamp);
          sheet.getRange(i + 1, 2).setValue(data.new_name);
          sheet.getRange(i + 1, 3).setValue(data.new_action);
          return ok('edited');
        }
      }
      return ok('not_found');
    }

    return ok('unknown_action');
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
