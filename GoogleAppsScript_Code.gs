// EJPM Clock Log — Google Apps Script backend (v2, for app v1.16.0+)
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

var TOKEN_TTL_SEC = 6 * 60 * 60;
var MAX_FAILS = 8;           // wrong passcodes allowed per lockout window
var LOCKOUT_SEC = 15 * 60;
var KIOSK_DAYS = 35;         // history the kiosk gets for one employee (covers a pay period)

// ───────────────────────── entry points ─────────────────────────

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return json({ status: 'error', msg: 'bad_json' }); }
  var type = data.action_type || 'clock';
  try {
    migrate_();
    if (type === 'clock') return withLock_(function () { return clock_(data); });
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

function readRows_() {
  var rows = logSheet_().getDataRange().getValues(), headers = rows[0], out = [];
  for (var i = 1; i < rows.length; i++) {
    var o = {};
    for (var j = 0; j < headers.length; j++) o[headers[j]] = rows[i][j];
    out.push(o);
  }
  return out;
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
