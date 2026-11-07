// EJPM Clock Log — Google Apps Script backend (v3, for app v1.17.0+)
// Paste this in Extensions → Apps Script, then Deploy → Manage deployments →
// Edit → Version: New version → Deploy (keeps the same URL).
//
// Security model
// - The kiosk (clock in/out) is public by design: anyone at the tablet can punch.
//   The server checks the name is on the staff list and blocks double clock-ins.
// - Everything else (full log, edits, deletes, staff list, passcode) needs an admin
//   token. A token comes from `login` with the passcode and lasts 6 hours.
// - The passcode lives in Script Properties (ADMIN_PASS) and is never sent out.
// - Repeated wrong passcodes lock admin login for 15 minutes.
// - Staff can ASK for a punch correction from the kiosk. It goes to the
//   "Correction Requests" tab and changes nothing until the admin approves it.

var TOKEN_TTL_SEC = 6 * 60 * 60;
var MAX_FAILS = 8;           // wrong passcodes allowed per lockout window
var LOCKOUT_SEC = 15 * 60;
var KIOSK_DAYS = 35;         // history the kiosk gets for one employee (covers a pay period)
var MAX_PENDING_PER_EMP = 5; // open correction requests one employee can have at once
var REQ_HEADERS = ['ID', 'Submitted', 'Name', 'Type', 'Punch Time', 'Punch Action',
                   'New Time', 'New Action', 'Reason', 'Status', 'Decided', 'Note'];

// ───────────────────────── entry points ─────────────────────────

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return json({ status: 'error', msg: 'bad_json' }); }
  var type = data.action_type || 'clock';
  try {
    migrate_();
    if (type === 'clock') return withLock_(function () { return clock_(data); });
    if (type === 'requestFix') return withLock_(function () { return requestFix_(data); });
    if (type === 'login') return login_(data);
    if (type === 'logout') { if (data.token) CacheService.getScriptCache().remove('tok_' + data.token); return ok('logged_out'); }

    // Everything below needs a valid admin token.
    if (!validToken_(data.token)) return json({ status: 'error', msg: 'unauthorized' });

    if (type === 'getLog') return json({ status: 'ok', rows: readRows_() });
    if (type === 'changePass') {
      if (!/^\d{4}$/.test(String(data.new_pass || ''))) return json({ status: 'error', msg: 'bad_pass' });
      props_().setProperty('ADMIN_PASS', String(data.new_pass));
      return ok('pass_changed');
    }
    if (type === 'employee') return withLock_(function () { return employeeOp_(data); });
    if (type === 'addRow') return withLock_(function () { return addRow_(data); });
    if (type === 'deleteRow') return withLock_(function () { return deleteRow_(data); });
    if (type === 'editRow') return withLock_(function () { return editRow_(data); });
    if (type === 'listRequests') return json({ status: 'ok', requests: listRequests_() });
    if (type === 'decideRequest') return withLock_(function () { return decideRequest_(data); });
    return json({ status: 'error', msg: 'unknown_action' });
  } catch (err) {
    return json({ status: 'error', msg: String(err) });
  }
}

function doGet(e) {
  migrate_();
  var p = (e && e.parameter) || {};
  // Public: staff list + break rules (the kiosk needs the names). No passcode.
  if (p.type === 'config') {
    var cfg = getConfig_();
    return json({ employees: cfg.employees, breaks: cfg.breaks });
  }
  // Public: one employee's recent punches (kiosk status + "hours this period").
  if (p.type === 'mine') return json({ status: 'ok', rows: rowsFor_(String(p.name || ''), KIOSK_DAYS) });
  // Public: one employee's own correction requests (so they can see the status).
  if (p.type === 'myRequests') {
    var nm = String(p.name || '');
    var mineReq = readTable_(reqSheet_()).filter(function (r) { return r.Name === nm; }).slice(-10).reverse();
    return json({ status: 'ok', requests: mineReq.map(publicReq_) });
  }
  // The full log used to be public here. It now needs an admin token (POST getLog).
  return json({ status: 'error', msg: 'unauthorized' });
}

// ───────────────────────── kiosk ─────────────────────────

function clock_(data) {
  var name = String(data.name || ''), action = String(data.action || '');
  if (action !== 'IN' && action !== 'OUT') return json({ status: 'error', msg: 'bad_action' });
  var emps = getConfig_().employees;
  if (emps && emps.indexOf(name) === -1) return json({ status: 'error', msg: 'unknown_employee' });
  var mine = rowsFor_(name, 0);
  var last = mine.length ? mine[mine.length - 1].Action : null;
  if (action === 'IN' && last === 'IN') return json({ status: 'error', msg: 'already_in' });
  if (action === 'OUT' && last !== 'IN') return json({ status: 'error', msg: 'not_in' });
  logSheet_().appendRow([
    new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }),
    name, action
  ]);
  SpreadsheetApp.flush();
  return json({ status: 'ok', msg: 'clocked', rows: rowsFor_(name, KIOSK_DAYS) });
}

// ───────────────────────── admin ─────────────────────────

function login_(data) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('login_fails') || 0);
  if (fails >= MAX_FAILS) return json({ status: 'error', msg: 'locked' });
  var pass = props_().getProperty('ADMIN_PASS');
  if (!pass || String(data.pass) !== pass) {
    cache.put('login_fails', String(fails + 1), LOCKOUT_SEC);
    return json({ status: 'error', msg: 'wrong_pass', left: Math.max(0, MAX_FAILS - fails - 1) });
  }
  cache.remove('login_fails');
  var token = Utilities.getUuid();
  cache.put('tok_' + token, '1', TOKEN_TTL_SEC);
  return json({ status: 'ok', token: token, ttl: TOKEN_TTL_SEC });
}

function validToken_(token) {
  return !!token && CacheService.getScriptCache().get('tok_' + token) === '1';
}

// Change ONE thing in the staff list on the server, under a lock, then return
// the new list. Two devices editing at once no longer overwrite each other.
function employeeOp_(data) {
  var cfg = getConfig_();
  var emps = cfg.employees || [], breaks = cfg.breaks || {};
  var name = String(data.name || '').trim();
  if (!name) return json({ status: 'error', msg: 'no_name' });
  if (data.op === 'add') {
    var exists = emps.some(function (x) { return x.toLowerCase() === name.toLowerCase(); });
    if (exists) return json({ status: 'error', msg: 'exists', config: cfg });
    emps.push(name);
    var b = parseInt(data.break_min, 10);
    if (b > 0) breaks[name] = Math.min(b, 240);
  } else if (data.op === 'remove') {
    emps = emps.filter(function (x) { return x !== name; });
    delete breaks[name];
  } else if (data.op === 'setBreak') {
    var m = parseInt(data.break_min, 10);
    if (m > 0) breaks[name] = Math.min(m, 240); else delete breaks[name];
  } else {
    return json({ status: 'error', msg: 'bad_op' });
  }
  cfg = { employees: emps, breaks: breaks };
  props_().setProperty('CONFIG', JSON.stringify(cfg));
  return json({ status: 'ok', config: cfg });
}

function addRow_(data) {
  if (data.action !== 'IN' && data.action !== 'OUT') return json({ status: 'error', msg: 'bad_action' });
  if (!data.name || isNaN(new Date(data.timestamp).getTime())) return json({ status: 'error', msg: 'bad_row' });
  logSheet_().appendRow([data.timestamp, data.name, data.action]);
  return ok('added');
}

function deleteRow_(data) {
  var sheet = logSheet_(), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (sameTs(rows[i][0], data.timestamp) && rows[i][1] === data.name && rows[i][2] === data.action) {
      sheet.deleteRow(i + 1);
      return ok('deleted');
    }
  }
  return json({ status: 'error', msg: 'not_found' });
}

function editRow_(data) {
  if (data.new_action !== 'IN' && data.new_action !== 'OUT') return json({ status: 'error', msg: 'bad_action' });
  if (!data.new_name || isNaN(new Date(data.new_timestamp).getTime())) return json({ status: 'error', msg: 'bad_row' });
  var sheet = logSheet_(), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (sameTs(rows[i][0], data.old_timestamp) && rows[i][1] === data.old_name && rows[i][2] === data.old_action) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[data.new_timestamp, data.new_name, data.new_action]]);
      return ok('edited');
    }
  }
  return json({ status: 'error', msg: 'not_found' });
}

// ───────────────────────── correction requests ─────────────────────────

// Kiosk: an employee asks for a fix. Nothing in the log changes here.
// type: 'edit'   → punch (punch_ts, punch_action) should be at new_ts
//       'add'    → a missed punch: new_action at new_ts
//       'delete' → punch (punch_ts, punch_action) should not exist
function requestFix_(data) {
  var name = String(data.name || ''), type = String(data.type || '');
  var reason = String(data.reason || '').trim().slice(0, 300);
  var emps = getConfig_().employees;
  if (!name || (emps && emps.indexOf(name) === -1)) return json({ status: 'error', msg: 'unknown_employee' });
  if (['edit', 'add', 'delete'].indexOf(type) === -1) return json({ status: 'error', msg: 'bad_type' });
  if (!reason) return json({ status: 'error', msg: 'no_reason' });

  var punchTs = '', punchAction = '', newTs = '', newAction = '';
  if (type === 'edit' || type === 'delete') {
    punchAction = String(data.punch_action || '');
    var target = findRow_(data.punch_ts, name, punchAction);
    if (!target) return json({ status: 'error', msg: 'punch_not_found' });
    punchTs = new Date(target.value).toISOString(); // exact instant of the stored punch
  }
  if (type === 'edit' || type === 'add') {
    var t = new Date(data.new_ts);
    if (isNaN(t.getTime())) return json({ status: 'error', msg: 'bad_time' });
    if (t.getTime() > Date.now() + 5 * 60000) return json({ status: 'error', msg: 'future_time' });
    if (t.getTime() < Date.now() - KIOSK_DAYS * 86400000) return json({ status: 'error', msg: 'too_old' });
    newTs = String(data.new_ts);
    newAction = type === 'add' ? String(data.new_action || '') : punchAction;
    if (newAction !== 'IN' && newAction !== 'OUT') return json({ status: 'error', msg: 'bad_action' });
  }
  var sheet = reqSheet_();
  var pending = readTable_(sheet).filter(function (r) { return r.Name === name && r.Status === 'PENDING'; }).length;
  if (pending >= MAX_PENDING_PER_EMP) return json({ status: 'error', msg: 'too_many_pending' });
  var id = Utilities.getUuid().slice(0, 8);
  sheet.appendRow([id, new Date().toISOString(), name, type, punchTs, punchAction, newTs, newAction, reason, 'PENDING', '', '']);
  return json({ status: 'ok', id: id });
}

// Admin: open requests first, then the 30 most recent decisions.
function listRequests_() {
  var all = readTable_(reqSheet_());
  var pending = all.filter(function (r) { return r.Status === 'PENDING'; });
  var done = all.filter(function (r) { return r.Status !== 'PENDING'; }).slice(-30).reverse();
  return pending.concat(done);
}

// Admin: approve (apply the change to the log) or reject a request.
function decideRequest_(data) {
  var sheet = reqSheet_(), values = sheet.getDataRange().getValues();
  var idx = -1;
  for (var i = 1; i < values.length; i++) if (String(values[i][0]) === String(data.id)) { idx = i; break; }
  if (idx === -1) return json({ status: 'error', msg: 'not_found' });
  var r = rowObj_(values[0], values[idx]);
  if (r.Status !== 'PENDING') return json({ status: 'error', msg: 'already_decided' });
  var note = String(data.note || '').slice(0, 300);

  if (data.decision === 'approve') {
    var log = logSheet_();
    if (r.Type === 'add') {
      log.appendRow([r['New Time'], r.Name, r['New Action']]);
    } else {
      var target = findRow_(r['Punch Time'], r.Name, r['Punch Action']);
      if (!target) return json({ status: 'error', msg: 'punch_changed' }); // edited/deleted since the request
      if (r.Type === 'delete') log.deleteRow(target.row);
      else log.getRange(target.row, 1).setValue(r['New Time']);
    }
  } else if (data.decision !== 'reject') {
    return json({ status: 'error', msg: 'bad_decision' });
  }
  sheet.getRange(idx + 1, 10, 1, 3).setValues([[data.decision === 'approve' ? 'APPROVED' : 'REJECTED', new Date().toISOString(), note]]);
  return json({ status: 'ok', msg: data.decision === 'approve' ? 'approved' : 'rejected' });
}

// What the kiosk may see about its own requests.
function publicReq_(r) {
  return { ID: r.ID, Submitted: r.Submitted, Type: r.Type, 'Punch Time': r['Punch Time'], 'Punch Action': r['Punch Action'],
           'New Time': r['New Time'], 'New Action': r['New Action'], Status: r.Status, Note: r.Note };
}

// The log row matching (timestamp, name, action): {row: sheet row number, value: stored cell}.
function findRow_(ts, name, action) {
  var rows = logSheet_().getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (sameTs(rows[i][0], ts) && rows[i][1] === name && rows[i][2] === action) return { row: i + 1, value: rows[i][0] };
  }
  return null;
}

// ───────────────────────── storage ─────────────────────────

function props_() { return PropertiesService.getScriptProperties(); }

// One-time moves from the old setup:
// - the passcode used to sit inside the public CONFIG blob → move it to ADMIN_PASS
// - pin the log tab by its ID so adding or reordering tabs can't redirect punches
function migrate_() {
  var p = props_();
  var raw = p.getProperty('CONFIG');
  if (raw) {
    try {
      var cfg = JSON.parse(raw);
      if (cfg && cfg.pass !== undefined) {
        if (!p.getProperty('ADMIN_PASS') && /^\d{4}$/.test(String(cfg.pass))) p.setProperty('ADMIN_PASS', String(cfg.pass));
        delete cfg.pass;
        p.setProperty('CONFIG', JSON.stringify(cfg));
      }
    } catch (err) {}
  }
  if (!p.getProperty('LOG_SHEET_ID')) {
    p.setProperty('LOG_SHEET_ID', String(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetId()));
  }
}

function getConfig_() {
  var raw = props_().getProperty('CONFIG'), cfg = {};
  if (raw) { try { cfg = JSON.parse(raw) || {}; } catch (err) {} }
  return { employees: Array.isArray(cfg.employees) ? cfg.employees : null, breaks: cfg.breaks || {} };
}

// The clock-log tab, found by its pinned ID (survives renames and tab reordering).
function logSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = Number(props_().getProperty('LOG_SHEET_ID'));
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (sheets[i].getSheetId() === id) return sheets[i];
  throw new Error('log_sheet_missing');
}

function readRows_() { return readTable_(logSheet_()); }

function readTable_(sheet) {
  var rows = sheet.getDataRange().getValues(), out = [];
  for (var i = 1; i < rows.length; i++) out.push(rowObj_(rows[0], rows[i]));
  return out;
}
function rowObj_(headers, row) {
  var o = {};
  for (var j = 0; j < headers.length; j++) o[headers[j]] = row[j];
  return o;
}

// The "Correction Requests" tab: created at the end on first use, then found by ID.
function reqSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), p = props_();
  var id = p.getProperty('REQ_SHEET_ID'), sheets = ss.getSheets();
  if (id) for (var i = 0; i < sheets.length; i++) if (String(sheets[i].getSheetId()) === id) return sheets[i];
  var sh = ss.insertSheet('Correction Requests', sheets.length);
  // Plain text: otherwise Sheets turns a reason like "May 5" into a date cell.
  sh.getRange('A:L').setNumberFormat('@');
  sh.appendRow(REQ_HEADERS);
  sh.setFrozenRows(1);
  p.setProperty('REQ_SHEET_ID', String(sh.getSheetId()));
  return sh;
}

// One employee's rows, oldest first. days>0 keeps only recent rows, but always
// includes the latest row so the kiosk knows whether they're clocked in.
function rowsFor_(name, days) {
  var mine = readRows_().filter(function (r) { return r.Name === name && r.Timestamp; });
  mine.sort(function (a, b) { return new Date(a.Timestamp) - new Date(b.Timestamp); });
  if (!days || !mine.length) return mine;
  var cutoff = Date.now() - days * 86400000, last = mine[mine.length - 1];
  var recent = mine.filter(function (r) { return new Date(r.Timestamp).getTime() >= cutoff; });
  if (recent.indexOf(last) === -1) recent.push(last);
  return recent;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json({ status: 'error', msg: 'busy' });
  try { return fn(); } finally { lock.releaseLock(); }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function ok(msg) { return json({ status: 'ok', msg: msg }); }

// Match timestamps robustly. Google Sheets stores the timestamp as a real
// date, so a stored cell (a Date) and the ISO string the app sends back never
// match as raw strings. Compare them as the same instant instead, with an
// exact-string fallback for cells Sheets kept as plain text.
function sameTs(cell, incoming) {
  if (String(cell) === String(incoming)) return true;
  var a = new Date(cell).getTime(), b = new Date(incoming).getTime();
  return !isNaN(a) && !isNaN(b) && a === b;
}
