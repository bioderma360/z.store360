// ===================== Config =====================
/**
 * Z. STORE Portal (Lite) — Config.gs
 * This version works directly on YOUR existing ORDERS spreadsheet (monthly tabs). Your columns, formulas and colours stay as they are.
 * The app only ADDS: 3 helper columns in each month tab, and a few helper tabs (Users, Sessions, Settings, AuditLog, Riders, CommissionGuide, PayLog).
 */
var CFG = {
  TZ: 'Asia/Karachi',
  PBKDF2_ITER: 1000,
  MAX_FAILED_LOGINS: 5,
  LOCKOUT_SECONDS: 900,
  LOCK_WAIT_MS: 25000,
  PAGE_SIZE: 25
};

/** Helper tabs created by setupLite(). "Name:n" = numeric column. */
var SCHEMA = {
  Users: 'UserID,Username,PasswordHash,Role,DisplayName,Active,CreatedAt,LastLogin,PasswordChangedAt,MustChangePassword',
  Sessions: 'SessionHash,UserID,CreatedAt,ExpiresAtMs:n,Revoked',
  Settings: 'Key,Value,Description',
  AuditLog: 'EventID,Timestamp,ActorID,ActorName,Role,Action,EntityType,EntityID,Details',
  Riders: 'RiderID,Name,Phone,Note,Active',
  CommissionGuide: 'GuideID,Product,PriceSlab,Commission:n,Sort:n',
  PayLog: 'LogID,PaidDate,Salesman,OrderCount:n,Amount:n,Note,CreatedBy,CreatedAt,Status,RequestID,Refs'
};

var DEFAULT_SETTINGS = [
  ['BusinessName', 'Z. STORE', 'Shown in the app header.'],
  ['BusinessPhone', '03318441416', 'Business number shown to salesmen.'],
  ['OrderTabs', 'AUTO', 'AUTO = every tab named like "OCTOBER 2026". Or list tab names separated by commas (e.g. OCTOBER 2026,NOVEMBER 2026).'],
  ['SalesmanCanChangeStatus', 'No', 'Yes = salesman apne Pending orders ko Delivered/Return kar sakta hai. No = sirf admin/staff status badalte hain.'],
  ['CycleCutoffHour', '12', 'Weekly payment cycle starts every Sunday at this hour (24h, Karachi time). 12 = Sunday noon.'],
  ['SessionHoursSalesman', '8', 'Salesman login lifetime in hours.'],
  ['SessionHoursAdmin', '4', 'Admin login lifetime in hours.'],
  ['PolicyText', 'Commission sirf Delivered orders par milta hai.\nReturn aur Cancel orders par commission nahi.\nPayment har Sunday ko hoti hai.\nJo Delivered orders ki payment baqi rahe woh agle hafte mein khud shamil ho jate hain.\n(Admin: yahan apni policy aur rules likhein — Policy page se badal sakte hain.)', 'Rules/policy text shown to salesmen.']
];

class AppError extends Error {
  constructor(message, code) { super(message); this.code = code || 'ERR'; this.isApp = true; }
}

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function pad2(n) { return ('0' + n).slice(-2); }
function toNum(v) { if (v === '' || v === null || v === undefined) return ''; var n = Number(v); return isNaN(n) ? '' : n; }
function rnd(len) { return (Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '')).slice(0, len); }
function list(s) { return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
function prop(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function setProp(k, v) { PropertiesService.getScriptProperties().setProperty(k, v); }
function normName(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }

/* ---------- Time (Asia/Karachi) ---------- */
function nowStr() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm:ss'); }
function todayStr() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }
function hourNow() { return Number(Utilities.formatDate(new Date(), CFG.TZ, 'H')); }
function dayNum(s) { var p = s.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; }
function fromDayNum(n) { var d = new Date(n * 86400000); return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate()); }
function isDate(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s) && fromDayNum(dayNum(s)) === s; }
/** Start of the current payment cycle: the most recent Sunday at CycleCutoffHour (default 12:00 noon). */
function cycleStart() {
  var cut = Number(getSettings().CycleCutoffHour); if (isNaN(cut) || cut < 0 || cut > 23) cut = 12;
  var n = dayNum(todayStr()), dow = new Date(n * 86400000).getUTCDay(), sunday = n - dow;
  if (dow === 0 && hourNow() < cut) sunday -= 7;
  return fromDayNum(sunday) + ' ' + pad2(cut) + ':00:00';
}

function reserveSeq(key, count) { var cur = Number(prop(key) || '0'); setProp(key, String(cur + count)); return cur + 1; }
function fmtId(prefix, n, pad) { return prefix + ('0000000000' + n).slice(-pad); }

var _lockDepth = 0;
function withLock(fn) {
  if (_lockDepth > 0) return fn();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(CFG.LOCK_WAIT_MS)) throw new AppError('System busy hai. Thori der baad dobara try karein.', 'BUSY');
  _lockDepth++;
  try { return fn(); } finally { _lockDepth--; SpreadsheetApp.flush(); lock.releaseLock(); }
}

var _settings = null;
function getSettings() {
  if (_settings) return _settings;
  var o = {};
  DEFAULT_SETTINGS.forEach(function (d) { o[d[0]] = d[1]; });
  readTable('Settings').forEach(function (r) { if (r.Key) o[r.Key] = r.Value; });
  _settings = o; return o;
}
function resetSettingsCache() { _settings = null; }
function isDemo() { return prop('MODE') === 'DEMO'; }

// ===================== Db =====================
/**
 * Db.gs — thin data-access layer over Google Sheets.
 * Rows are addressed by permanent ID (never by row number alone). All writes must be wrapped in withLock().
 */
var _ss = null;
function ss() {
  if (_ss) return _ss;
  var id = prop('SPREADSHEET_ID');
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_ss) throw new Error('Spreadsheet is not configured. Set the SPREADSHEET_ID script property.');
  return _ss;
}
function sheet(name) {
  var s = ss().getSheetByName(name);
  if (!s) throw new Error('Missing sheet "' + name + '". Run setupDatabase().');
  return s;
}
function cols(name) {
  return SCHEMA[name].split(',').map(function (c) { var p = c.split(':'); return { name: p[0], type: p[1] || 's' }; });
}
function colIndex(name, col) {
  var c = cols(name);
  for (var i = 0; i < c.length; i++) if (c[i].name === col) return i + 1;
  throw new Error('Unknown column ' + col);
}

function rowToObj(name, r, rowNo) {
  var c = cols(name), o = { _row: rowNo };
  for (var j = 0; j < c.length; j++) {
    var v = r[j];
    if (c[j].type === 'n') o[c[j].name] = toNum(v);
    else if (v instanceof Date) o[c[j].name] = Utilities.formatDate(v, CFG.TZ, 'yyyy-MM-dd HH:mm:ss');
    else o[c[j].name] = (v === null || v === undefined) ? '' : String(v);
  }
  return o;
}
function objToRow(name, o) {
  return cols(name).map(function (c) {
    var v = o[c.name];
    if (v === undefined || v === null) return '';
    if (c.type === 'n') { return (v === '' || isNaN(Number(v))) ? '' : Number(v); }
    return String(v);
  });
}

function idIndex(name) { return 1; }   // helper tabs keep their ID in column A

function readTable(name) {
  var s = sheet(name), last = s.getLastRow(), n = cols(name).length;
  if (last < 2) return [];
  var vals = s.getRange(2, 1, last - 1, n).getValues(), out = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i].every(function (v) { return v === '' || v === null; })) continue;   // skip fully blank rows
    out.push(rowToObj(name, vals[i], i + 2));
  }
  return out;
}

/** Find one row by its ID in column A. Reads one column, then one row. */
function getById(name, id) {
  id = String(id || '');
  if (!id) return null;
  var s = sheet(name), last = s.getLastRow();
  if (last < 2) return null;
  var ids = s.getRange(2, idIndex(name), last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      return rowToObj(name, s.getRange(i + 2, 1, 1, cols(name).length).getValues()[0], i + 2);
    }
  }
  return null;
}

function applyFormats(name, fromRow, numRows) {
  if (numRows < 1) return;
  var s = sheet(name);
  cols(name).forEach(function (c, i) {
    if (c.type === 's') s.getRange(fromRow, i + 1, numRows, 1).setNumberFormat('@');   // plain text: keeps leading zeros, stops formula/date conversion
  });
}
function ensureRows(s, name, neededLastRow) {
  var max = s.getMaxRows();
  if (neededLastRow > max) {
    var add = neededLastRow - max + 500;
    s.insertRowsAfter(max, add);
    applyFormats(name, max + 1, add);
  }
}

/** Append many objects in a single write. Returns the first row number written. */
function appendObjects(name, objs) {
  if (!objs.length) return 0;
  var s = sheet(name), start = s.getLastRow() + 1, n = cols(name).length;
  ensureRows(s, name, start + objs.length - 1);
  s.getRange(start, 1, objs.length, n).setValues(objs.map(function (o) { return objToRow(name, o); }));
  return start;
}
function updateObj(name, o) {
  if (!o._row) throw new Error('updateObj needs _row');
  sheet(name).getRange(o._row, 1, 1, cols(name).length).setValues([objToRow(name, o)]);
}
function deleteRowAt(name, rowNo) { sheet(name).deleteRow(rowNo); }

/** Overwrite one column for many rows with a single read + single write. items = [{_row, value}] */
function setColumn(name, colName, items) {
  if (!items.length) return;
  var s = sheet(name), last = s.getLastRow(), ci = colIndex(name, colName);
  var rng = s.getRange(2, ci, last - 1, 1), vals = rng.getValues();
  items.forEach(function (it) { vals[it._row - 2][0] = it.value; });
  rng.setValues(vals);
}

/* ---------- Audit + status history ---------- */
function audit(ctx, action, entityType, entityId, details) {
  var u = (ctx && ctx.user) || {};
  withLock(function () {
    appendObjects('AuditLog', [{
      EventID: 'AU-' + rnd(10), Timestamp: nowStr(), ActorID: u.userId || '', ActorName: u.username || '',
      Role: u.role || '', Action: action, EntityType: entityType || '', EntityID: entityId || '',
      Details: String(details || '').slice(0, 800)
    }]);
  });
}

// ===================== Auth =====================
/**
 * Auth.gs — custom username/password authentication, sessions, rate limiting, user management.
 * Password storage: PBKDF2-HMAC-SHA256 (RFC 8018) built on Utilities.computeHmacSha256Signature,
 * keyed with a server-side pepper (Script Properties). See README for the honest limits of this approach.
 */
var GENERIC_LOGIN_ERROR = 'Invalid username or password.';

/* ---------- Crypto helpers ---------- */
function bytesToHex(b) { return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join(''); }
function hexToBytes(h) {
  var out = [];
  for (var i = 0; i < h.length; i += 2) { var v = parseInt(h.substr(i, 2), 16); out.push(v > 127 ? v - 256 : v); }
  return out;
}
function strBytes(s) { return Utilities.newBlob(s).getBytes(); }
function sha256hex(s) { return bytesToHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }
function hmacHex(key, msg) { return bytesToHex(Utilities.computeHmacSha256Signature(msg, key)); }
function pepper() {
  var p = prop('PEPPER');
  if (!p) throw new Error('PEPPER missing. Run setupDatabase().');
  return p;
}
function constEq(a, b) {
  a = String(a); b = String(b);
  var d = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return d === 0;
}
/** PBKDF2-HMAC-SHA256, dkLen = 32 (a single block, so no block loop is required). */
function pbkdf2Hex(password, saltHex, iter) {
  var key = strBytes(password);
  var u = Utilities.computeHmacSha256Signature(hexToBytes(saltHex).concat([0, 0, 0, 1]), key);
  var t = u.slice();
  for (var i = 1; i < iter; i++) {
    u = Utilities.computeHmacSha256Signature(u, key);
    for (var j = 0; j < 32; j++) t[j] ^= u[j];
  }
  return bytesToHex(t);
}
function hashPassword(pw) {
  var salt = rnd(32), pre = hmacHex(pepper(), pw);
  return 'p1$' + CFG.PBKDF2_ITER + '$' + salt + '$' + pbkdf2Hex(pre, salt, CFG.PBKDF2_ITER);
}
function verifyPassword(pw, stored) {
  var p = String(stored || '').split('$');
  if (p.length !== 4 || p[0] !== 'p1') return false;
  var iter = Number(p[1]);
  if (!(iter >= 100 && iter <= 100000)) return false;
  return constEq(pbkdf2Hex(hmacHex(pepper(), pw), p[2], iter), p[3]);
}
function checkPasswordPolicy(pw) {
  if (typeof pw !== 'string' || pw.length < 10 || pw.length > 100 || !/[A-Za-z]/.test(pw) || !/\d/.test(pw))
    throw new AppError('Password must be 10–100 characters and include letters and numbers.');
}
function generatePassword() {
  var cs = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
  for (var tries = 0; tries < 20; tries++) {
    var h = hmacHex(pepper(), Utilities.getUuid() + Utilities.getUuid() + Date.now()), pw = '';
    for (var i = 0; i < 14; i++) pw += cs.charAt(parseInt(h.substr(i * 2, 2), 16) % cs.length);
    if (/[A-Za-z]/.test(pw) && /\d/.test(pw)) return pw;
  }
  throw new Error('Could not generate password');
}
function sessionHash(token) { return hmacHex(pepper(), 'session:' + token); }

/* ---------- Users ---------- */
function userPublic(u) {
  return { userId: u.UserID, username: u.Username, role: u.Role, displayName: u.DisplayName,
           active: u.Active === 'Yes', createdAt: u.CreatedAt, lastLogin: u.LastLogin, mustChange: u.MustChangePassword === 'Yes' };
}
function findUserByName(uname) {
  var all = readTable('Users');
  for (var i = 0; i < all.length; i++) if (all[i].Username === uname) return all[i];
  return null;
}
function clearUserCache(userId) { CacheService.getScriptCache().remove('U:' + userId); }

/* ---------- Login / sessions ---------- */
function login(p) {
  var uname = String(p.username || '').trim().toLowerCase(), pw = String(p.password || '');
  if (!uname || !pw || uname.length > 40 || pw.length > 200) throw new AppError(GENERIC_LOGIN_ERROR, 'LOGIN');
  var cache = CacheService.getScriptCache(), fkey = 'F:' + sha256hex(uname);
  var fails = Number(cache.get(fkey) || '0');
  if (fails >= CFG.MAX_FAILED_LOGINS) {
    throw new AppError('Too many failed attempts. Please wait 15 minutes and try again, or ask the admin.', 'LOGIN');
  }
  var u = findUserByName(uname);
  var ok = u ? verifyPassword(pw, u.PasswordHash) : (verifyPassword(pw, 'p1$' + CFG.PBKDF2_ITER + '$' + rnd(32) + '$' + rnd(64)), false);
  if (!u || !ok || u.Active !== 'Yes') {
    cache.put(fkey, String(fails + 1), CFG.LOCKOUT_SECONDS);
    if (u) audit({ user: { userId: u.UserID, username: u.Username, role: u.Role } },
                 fails + 1 >= CFG.MAX_FAILED_LOGINS ? 'LOGIN_LOCKOUT' : 'LOGIN_FAILED', 'User', u.UserID, 'attempt ' + (fails + 1));
    throw new AppError(GENERIC_LOGIN_ERROR, 'LOGIN');
  }
  cache.remove(fkey);
  var s = getSettings(), hours = Number(u.Role === 'admin' ? s.SessionHoursAdmin : s.SessionHoursSalesman) || 8;
  var token = hmacHex(pepper(), Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid() + Date.now());
  var exp = Date.now() + hours * 3600000, h = sessionHash(token);
  withLock(function () {
    purgeSessions();
    appendObjects('Sessions', [{ SessionHash: h, UserID: u.UserID, CreatedAt: nowStr(), ExpiresAtMs: exp, Revoked: 'No' }]);
    var fresh = getById('Users', u.UserID); fresh.LastLogin = nowStr(); updateObj('Users', fresh);
  });
  clearUserCache(u.UserID);
  audit({ user: { userId: u.UserID, username: u.Username, role: u.Role } }, 'LOGIN', 'User', u.UserID, '');
  return { token: token, expiresAtMs: exp, user: userPublic(u) };
}

function purgeSessions() {
  var s = sheet('Sessions'), last = s.getLastRow();
  if (last < 200) return;
  var rows = readTable('Sessions'), cutoff = Date.now() - 86400000;
  var keep = rows.filter(function (r) { return r.ExpiresAtMs > cutoff; });
  s.getRange(2, 1, last - 1, cols('Sessions').length).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, cols('Sessions').length).setValues(keep.map(function (r) { return objToRow('Sessions', r); }));
}

/** Validates a token and returns { user, sessionHash }. Throws AppError(code 'AUTH') on any problem. */
function authenticate(token) {
  var fail = function () { throw new AppError('Your session has expired. Please sign in again.', 'AUTH'); };
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/.test(token)) fail();
  var h = sessionHash(token), cache = CacheService.getScriptCache(), sess = null, raw = cache.get('S:' + h);
  if (raw) sess = JSON.parse(raw);
  else {
    var row = getById('Sessions', h);
    if (!row || row.Revoked === 'Yes') fail();
    sess = { u: row.UserID, e: row.ExpiresAtMs };
    cache.put('S:' + h, JSON.stringify(sess), 300);
  }
  if (!(sess.e > Date.now())) fail();
  var user = null, uraw = cache.get('U:' + sess.u);
  if (uraw) user = JSON.parse(uraw);
  else {
    var ur = getById('Users', sess.u);
    if (!ur) fail();
    user = { userId: ur.UserID, username: ur.Username, role: ur.Role, displayName: ur.DisplayName, active: ur.Active === 'Yes', mustChange: ur.MustChangePassword === 'Yes' };
    cache.put('U:' + sess.u, JSON.stringify(user), 60);
  }
  if (!user.active) fail();
  return { user: user, sessionHash: h };
}

function revokeSessions(userId, exceptHash) {
  var cache = CacheService.getScriptCache();
  withLock(function () {
    var changed = false;
    readTable('Sessions').forEach(function (r) {
      if (r.UserID === userId && r.Revoked !== 'Yes' && r.SessionHash !== exceptHash) {
        r.Revoked = 'Yes'; updateObj('Sessions', r); cache.remove('S:' + r.SessionHash); changed = true;
      }
    });
    return changed;
  });
}
function logout(ctx) {
  withLock(function () {
    var r = getById('Sessions', ctx.sessionHash);
    if (r) { r.Revoked = 'Yes'; updateObj('Sessions', r); }
  });
  CacheService.getScriptCache().remove('S:' + ctx.sessionHash);
  audit(ctx, 'LOGOUT', 'User', ctx.user.userId, '');
  return { done: true };
}

function changePassword(ctx, p) {
  var oldPw = String(p.oldPassword || ''), newPw = String(p.newPassword || '');
  checkPasswordPolicy(newPw);
  var u = getById('Users', ctx.user.userId);
  if (!u || !verifyPassword(oldPw, u.PasswordHash)) throw new AppError('Current password is incorrect.');
  if (oldPw === newPw) throw new AppError('Choose a different password.');
  withLock(function () {
    var fresh = getById('Users', u.UserID);
    fresh.PasswordHash = hashPassword(newPw); fresh.PasswordChangedAt = nowStr(); fresh.MustChangePassword = 'No';
    updateObj('Users', fresh);
  });
  clearUserCache(u.UserID);
  revokeSessions(u.UserID, ctx.sessionHash);
  audit(ctx, 'PASSWORD_CHANGED', 'User', u.UserID, '');
  return { done: true };
}

/* ---------- Admin: user management ---------- */
function adminListUsers() {
  return readTable('Users').map(userPublic);
}
function adminCreateUser(ctx, p) {
  var uname = String(p.username || '').trim().toLowerCase(), dn = String(p.displayName || '').trim().replace(/\s+/g, ' ');
  var role = p.role === 'admin' ? 'admin' : (p.role === 'staff' ? 'staff' : 'salesman');
  if (!/^[a-z0-9._-]{3,30}$/.test(uname)) throw new AppError('Username must be 3–30 characters: letters, numbers, dot, dash, underscore.');
  if (dn.length < 2 || dn.length > 60) throw new AppError('Naam 2–60 characters ka hona chahiye.');
  var tempPw = generatePassword();
  var created = withLock(function () {
    var users = readTable('Users');
    users.forEach(function (x) {
      if (x.Username === uname) throw new AppError('That username is already taken.');
      if (x.DisplayName.toLowerCase() === dn.toLowerCase()) throw new AppError('Another user already has that display name.');
    });
    var id = fmtId('U-', reserveSeq('USER_SEQ', 1), 4);
    var u = { UserID: id, Username: uname, PasswordHash: hashPassword(tempPw), Role: role, DisplayName: dn, Active: 'Yes',
              CreatedAt: nowStr(), LastLogin: '', PasswordChangedAt: nowStr(), MustChangePassword: 'Yes' };
    appendObjects('Users', [u]);
    return u;
  });
  audit(ctx, 'USER_CREATED', 'User', created.UserID, 'username=' + uname + ' role=' + role);
  return { user: userPublic(created), tempPassword: tempPw };
}
function adminUpdateUser(ctx, p) {
  var dn = String(p.displayName || '').trim().replace(/\s+/g, ' ');
  if (dn.length < 2 || dn.length > 60) throw new AppError('Naam 2–60 characters ka hona chahiye.');
  var res = withLock(function () {
    var u = getById('Users', p.userId);
    if (!u) throw new AppError('User not found.');
    readTable('Users').forEach(function (x) { if (x.UserID !== u.UserID && x.DisplayName.toLowerCase() === dn.toLowerCase()) throw new AppError('Another user already has that display name.'); });
    var old = u.DisplayName; u.DisplayName = dn; updateObj('Users', u);
    return { old: old, u: u };
  });
  clearUserCache(res.u.UserID);
  audit(ctx, 'USER_RENAMED', 'User', res.u.UserID, res.old + ' -> ' + dn);
  return userPublic(res.u);
}
function adminResetPassword(ctx, p) {
  var tempPw = generatePassword();
  var u = withLock(function () {
    var x = getById('Users', p.userId);
    if (!x) throw new AppError('User not found.');
    x.PasswordHash = hashPassword(tempPw); x.PasswordChangedAt = nowStr(); x.MustChangePassword = 'Yes';
    updateObj('Users', x); return x;
  });
  clearUserCache(u.UserID);
  revokeSessions(u.UserID, null);
  audit(ctx, 'PASSWORD_RESET', 'User', u.UserID, '');
  return { user: userPublic(u), tempPassword: tempPw };
}
function adminSetActive(ctx, p) {
  var active = p.active === true;
  var u = withLock(function () {
    var x = getById('Users', p.userId);
    if (!x) throw new AppError('User not found.');
    if (x.UserID === ctx.user.userId && !active) throw new AppError('You cannot deactivate your own account.');
    x.Active = active ? 'Yes' : 'No'; updateObj('Users', x); return x;
  });
  clearUserCache(u.UserID);
  if (!active) revokeSessions(u.UserID, null);
  audit(ctx, active ? 'USER_ACTIVATED' : 'USER_DEACTIVATED', 'User', u.UserID, '');
  return userPublic(u);
}
function adminRevokeSessions(ctx, p) {
  revokeSessions(String(p.userId || ''), null);
  clearUserCache(String(p.userId || ''));
  audit(ctx, 'SESSIONS_REVOKED', 'User', p.userId, '');
  return { done: true };
}

// ===================== Sheet =====================
/**
 * Sheet.gs — reads and writes YOUR existing order tabs (e.g. "OCTOBER 2026").
 * Columns are found by their header text, so your layout can change. Formulas (DC, Profit) are never overwritten.
 * Orders have no ID column: an order is addressed as "TAB|row|check", where check detects that the row moved.
 */
var TAB_RE = /^(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+(\d{4})$/i;
var MONTHS_UP = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
var HEAD_KEYS = {
  product: 'product', price: 'price', customername: 'customer', name: 'customer', customer: 'customer', address: 'address',
  phonenumber: 'phone', phone: 'phone', phoneno: 'phone', saleman: 'salesman', salesman: 'salesman', courier: 'courier',
  tracking: 'tracking', trackingnumber: 'tracking', trackingno: 'tracking', status: 'status', payment: 'payment', date: 'date',
  productcost: 'pcost', adscost: 'ads', deliverycharges: 'dc', salemanc: 'comm', salesmanc: 'comm', salemancommission: 'comm',
  salesmancommission: 'comm', commission: 'comm', profit: 'profit',
  orderno: 'orderno', ordernumber: 'orderno', orderid: 'orderno',
  salemanpayment: 'spay', salesmanpayment: 'spay', salemanpaymentdate: 'spaydate', salesmanpaymentdate: 'spaydate', statusdate: 'sdate'
};
var STATUS_STYLE = {                       // same colours you use by hand
  Delivered: { bg: '#00ff00', fg: '#434343' },
  Return: { bg: '#ff0000', fg: '#ffffff' },
  Cancel: { bg: '#00ffff', fg: '#434343' },
  Pending: { bg: '#ffffff', fg: '#434343' }
};
var YELLOW = '#ffff00', WHITE = '#ffffff';

function headKey(s) { return HEAD_KEYS[String(s).toLowerCase().replace(/[^a-z]/g, '')] || ''; }
function colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/** Column map of a tab from its header row, or null if it is not an orders tab. */
function headerMapOf(sh) {
  var lc = Math.min(Math.max(sh.getLastColumn(), 1), 40);
  var row = sh.getRange(1, 1, 1, lc).getValues()[0], cols = {}, width = 0;
  row.forEach(function (h, i) { var k = headKey(h); if (k && !cols[k]) { cols[k] = i + 1; width = Math.max(width, i + 1); } });
  if (!(cols.product && cols.price && cols.customer && cols.salesman && cols.status)) return null;
  return { cols: cols, width: width };
}
function tabSortKey(n) { var m = String(n).trim().match(TAB_RE); return m ? Number(m[2]) * 100 + MONTHS_UP.indexOf(m[1].toUpperCase()) : 0; }

/** Order tabs in chronological order. */
function orderTabs() {
  var setting = String(getSettings().OrderTabs || 'AUTO').trim(), wanted = null, out = [];
  if (setting.toUpperCase() !== 'AUTO') wanted = list(setting).map(function (x) { return x.toLowerCase(); });
  ss().getSheets().forEach(function (sh, i) {
    var n = sh.getName().trim();
    if (wanted ? wanted.indexOf(n.toLowerCase()) < 0 : !TAB_RE.test(n)) return;
    var h = headerMapOf(sh); if (!h) return;
    out.push({ name: sh.getName(), sheet: sh, h: h, key: tabSortKey(n) || (i + 1) });
  });
  out.sort(function (a, b) { return a.key - b.key; });
  return out;
}
function tabByName(name) {
  var t = orderTabs().filter(function (x) { return x.name === name; })[0];
  if (!t) throw new AppError('Yeh tab nahi mila.', 'BAD');
  return t;
}

/* ---------- parsing ---------- */
function parseMoney(x) {
  var s = String(x == null ? '' : x).replace(/rs\.?/ig, '').replace(/[,\s]/g, '');
  if (s === '') return null;
  return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
}
function money0(x) { var n = parseMoney(x); return (n === null || isNaN(n)) ? 0 : n; }
function normStatus(raw) {
  var t = String(raw || '').trim().toLowerCase();
  if (t === 'delivered' || t === 'delivery') return 'Delivered';
  if (t === 'return' || t === 'returned') return 'Return';
  if (t === 'cancel' || t === 'cancelled' || t === 'canceled') return 'Cancel';
  return 'Pending';          // blank (or anything unfinished)
}
function hash32(s) { var h = 5381; for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return ('00000000' + (h >>> 0).toString(16)).slice(-8); }
function rowCheck(customer, phone, price, product) {
  return hash32([normName(customer), String(phone).replace(/\D/g, ''), String(money0(price)), normName(product)].join('|'));
}

function rowToOrder(tab, rowNo, r) {
  var c = tab.h.cols, g = function (k) { return c[k] ? String(r[c[k] - 1] == null ? '' : r[c[k] - 1]).trim() : ''; };
  var o = { tab: tab.name, row: rowNo, product: g('product'), price: money0(g('price')), customer: g('customer'), address: g('address'), phone: g('phone'),
            salesman: g('salesman'), courier: g('courier'), tracking: g('tracking'), statusRaw: g('status'), payment: g('payment'), date: g('date'),
            pcost: g('pcost'), ads: g('ads'), dc: g('dc'), comm: money0(g('comm')), commRaw: g('comm'), profit: g('profit'),
            spay: g('spay'), spayDate: g('spaydate'), sdate: g('sdate'), orderNo: g('orderno'), hasNoCol: !!c.orderno };
  o.status = normStatus(o.statusRaw);
  o.paid = o.spay.toLowerCase() === 'paid';
  o.ref = tab.name + '|' + rowNo + '|' + rowCheck(o.customer, o.phone, o.price, o.product);
  return o;
}
function isBlankRow(o) { return o.product === '' && o.customer === '' && o.phone === '' && o.price === 0; }

function readTab(tab) {
  var sh = tab.sheet, last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, tab.h.width).getDisplayValues(), out = [];
  for (var i = 0; i < vals.length; i++) { var o = rowToOrder(tab, i + 2, vals[i]); if (!isBlankRow(o)) out.push(o); }
  return out;
}
function readAllOrders() {
  var all = [];
  orderTabs().forEach(function (t) { all = all.concat(readTab(t)); });
  return all;
}
function readOneRow(tab, rowNo) {
  if (!(rowNo >= 2 && rowNo <= tab.sheet.getLastRow())) return null;
  var o = rowToOrder(tab, rowNo, tab.sheet.getRange(rowNo, 1, 1, tab.h.width).getDisplayValues()[0]);
  return isBlankRow(o) ? null : o;
}
/** Resolves "TAB|row|check" and proves the row is still the same order. */
function locate(ref) {
  var p = String(ref || '').split('|');
  if (p.length !== 3) throw new AppError('Order nahi mila. Page refresh karein.', 'NOTFOUND');
  var tab = tabByName(p[0]), row = Number(p[1]), o = readOneRow(tab, row);
  if (!o || o.ref !== ref) throw new AppError('Sheet mein yeh order badal chuka hai. Page refresh karein.', 'CONFLICT');
  return { tab: tab, o: o };
}

/** Newest first: latest month tab first, bottom rows first. */
function sortNewest(list) {
  var key = {}; orderTabs().forEach(function (t) { key[t.name] = t.key; });
  return list.slice().sort(function (a, b) { return (key[b.tab] - key[a.tab]) || (b.row - a.row); });
}

/* ---------- writing ---------- */
function setTextCell(sh, row, col, value) {
  if (!col) return;
  var cell = sh.getRange(row, col); cell.setNumberFormat('@'); cell.setValue(String(value));
}
function setNumCell(sh, row, col, value) { if (col) sh.getRange(row, col).setValue(value === '' ? '' : Number(value)); }
function paintStatus(tab, row, status, stamp) {
  var c = tab.h.cols, st = STATUS_STYLE[status] || STATUS_STYLE.Pending, cell = tab.sheet.getRange(row, c.status);
  cell.setNumberFormat('@'); cell.setValue(status === 'Pending' ? '' : status); cell.setBackground(st.bg).setFontColor(st.fg);
  if (c.sdate) { var d = tab.sheet.getRange(row, c.sdate); d.setNumberFormat('@'); d.setValue(status === 'Pending' ? '' : (stamp || nowStr())); }
}
function paintPaymentCell(tab, row, paid) {
  var c = tab.h.cols; if (!c.payment) return;
  var cell = tab.sheet.getRange(row, c.payment); cell.setNumberFormat('@'); cell.setValue(paid ? 'PAID' : '');
  cell.setBackground(paid ? YELLOW : WHITE).setFontColor('#434343');
}
/** Marks/unmarks many rows of one tab as paid to the salesman: one read + one write per column. */
function writeSalesmanPaid(tab, rows, paid, dateStr) {
  var c = tab.h.cols; if (!c.spay || !c.spaydate) throw new AppError('Is tab mein "Saleman Payment" columns nahi hain. setupLite() chalayein.');
  var sh = tab.sheet, last = sh.getLastRow(), n = last - 1, set = {}; rows.forEach(function (r) { set[r] = 1; });
  [c.spay, c.spaydate].forEach(function (col) {
    var rng = sh.getRange(2, col, n, 1); rng.setNumberFormat('@');
    var vals = rng.getValues(), bgs = rng.getBackgrounds();
    Object.keys(set).forEach(function (r) {
      var i = Number(r) - 2; if (i < 0 || i >= n) return;
      if (col === c.spay) { vals[i][0] = paid ? 'Paid' : ''; bgs[i][0] = paid ? YELLOW : WHITE; }
      else { vals[i][0] = paid ? dateStr : ''; }
    });
    rng.setValues(vals); if (col === c.spay) rng.setBackgrounds(bgs);
  });
}

/** Values of the dropdown on a column (data validation), else the distinct values already used. */
function dropdownValues(tab, key, fallbackFromOrders) {
  var col = tab.h.cols[key], out = [];
  try {
    var rule = tab.sheet.getRange(Math.min(3, tab.sheet.getMaxRows()), col).getDataValidation();
    if (rule) {
      var type = String(rule.getCriteriaType()), v = rule.getCriteriaValues();
      if (type.indexOf('VALUE_IN_LIST') >= 0) out = v[0].slice();
      else if (type.indexOf('VALUE_IN_RANGE') >= 0) out = v[0].getValues().map(function (r) { return String(r[0]).trim(); }).filter(Boolean);
    }
  } catch (e) { /* no dropdown */ }
  if (!out.length && fallbackFromOrders) {
    var seen = {}; fallbackFromOrders.forEach(function (o) { var x = o[key]; if (x && !seen[normName(x)]) { seen[normName(x)] = 1; out.push(x); } });
    out.sort();
  }
  return out;
}

/** Creates a new month tab ("NOVEMBER 2026") as a clean copy of the latest month tab: header, widths, formats, dropdowns and DC/Profit formulas, but no orders. */
function createMonthTab(name) {
  name = String(name || '').trim().replace(/\s+/g, ' ').toUpperCase();
  if (!TAB_RE.test(name)) throw new AppError('Tab ka naam "NOVEMBER 2026" jaisa hona chahiye (mahina + saal).', 'VALIDATION');
  if (ss().getSheetByName(name)) throw new AppError('Yeh tab pehle se maujood hai: ' + name, 'VALIDATION');
  var tabs = orderTabs(); if (!tabs.length) throw new AppError('Koi order tab nahi mila.');
  var src = tabs[tabs.length - 1], w = src.h.width, ns = ss().insertSheet(name), rows = 999;
  var tr = Math.min(firstFreeRow(src), src.sheet.getMaxRows());           // a blank-looking row of the old tab: its look, dropdowns and formulas are the template
  src.sheet.getRange(1, 1, 1, w).copyTo(ns.getRange(1, 1));
  for (var c = 1; c <= w; c++) ns.setColumnWidth(c, src.sheet.getColumnWidth(c));
  ns.setFrozenRows(1);
  var tpl = src.sheet.getRange(tr, 1, 1, w), body = ns.getRange(2, 1, rows, w);
  tpl.copyTo(body, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  tpl.copyTo(body, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  var tf = templateFormulas(src, tr);
  ['dc', 'profit'].forEach(function (k) { if (src.h.cols[k] && tf[k]) ns.getRange(2, src.h.cols[k], rows, 1).setFormulaR1C1(tf[k]); });
  try { ss().setActiveSheet(ns); ss().moveActiveSheet(src.sheet.getIndex() + 1); } catch (e) { /* position is cosmetic */ }
  return tabByName(name);
}
/** The tab for the current month ("OCTOBER 2026"); created from the latest tab if it does not exist yet. */
function currentMonthTab() {
  var name = MONTHS_UP[Number(Utilities.formatDate(new Date(), CFG.TZ, 'M')) - 1] + ' ' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyy');
  var hit = orderTabs().filter(function (t) { return t.name.trim().toUpperCase() === name; })[0];
  return hit || createMonthTab(name);
}
/* ---------- Order numbers: GL-00001, GL-00002 ... one counter for all months ---------- */
function orderNoStr(n) { return 'GL-' + ('00000' + n).slice(-5); }
function hasOrderContent(r, c) { return !!(String(r[c.product - 1]).trim() || String(r[c.customer - 1]).trim() || String(r[c.price - 1]).trim()); }
function maxOrderNoAll() {
  var max = 0;
  orderTabs().forEach(function (t) {
    var col = t.h.cols.orderno, last = t.sheet.getLastRow(); if (!col || last < 2) return;
    t.sheet.getRange(2, col, last - 1, 1).getValues().forEach(function (r) { var m = String(r[0]).trim().match(/^GL-(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
  });
  return max;
}
/** Next free order number. Call while holding the lock. */
function nextOrderNo() {
  var cur = prop('ORDERNO_SEQ'); var n = (cur === null ? maxOrderNoAll() : Number(cur)) + 1;
  setProp('ORDERNO_SEQ', String(n)); return orderNoStr(n);
}
/** Gives a number to every order that has none (oldest month first) and renumbers accidental duplicates. Returns how many were assigned. */
function backfillOrderNumbers() {
  return withLock(function () {
    var tabs = orderTabs().filter(function (t) { return t.h.cols.orderno; }), seen = {}, plan = [], max = Math.max(Number(prop('ORDERNO_SEQ') || '0'), maxOrderNoAll()), assigned = 0;
    tabs.forEach(function (t) { var last = t.sheet.getLastRow(); if (last >= 2) plan.push({ t: t, vals: t.sheet.getRange(2, 1, last - 1, t.h.width).getValues() }); });
    plan.forEach(function (p) {
      var c = p.t.h.cols, out = [], changed = false;
      p.vals.forEach(function (r) {
        var cur = String(r[c.orderno - 1]).trim();
        if (!hasOrderContent(r, c)) { out.push([cur]); return; }
        if (cur && !seen[cur]) { seen[cur] = 1; out.push([cur]); return; }
        max++; var n = orderNoStr(max); seen[n] = 1; out.push([n]); changed = true; assigned++;
      });
      if (changed) { var rng = p.t.sheet.getRange(2, c.orderno, out.length, 1); rng.setNumberFormat('@'); rng.setValues(out); }
    });
    setProp('ORDERNO_SEQ', String(max));
    return assigned;
  });
}
/** For rows just typed or pasted in the sheet: number the ones without a number; after a wide paste also fix copied (duplicate) numbers. */
function numberEditedRows(tab, r1, r2, wide) {
  var c = tab.h.cols; if (!c.orderno || r2 < r1) return;
  withLock(function () {
    var sh = tab.sheet, n = r2 - r1 + 1, vals = sh.getRange(r1, 1, n, tab.h.width).getValues(), counts = null, out = [], changed = false;
    if (wide) { counts = {}; orderTabs().forEach(function (t) { var col = t.h.cols.orderno, last = t.sheet.getLastRow(); if (col && last >= 2) t.sheet.getRange(2, col, last - 1, 1).getValues().forEach(function (x) { var k = String(x[0]).trim(); if (k) counts[k] = (counts[k] || 0) + 1; }); }); }
    vals.forEach(function (r) {
      var cur = String(r[c.orderno - 1]).trim();
      if (!hasOrderContent(r, c)) { out.push([cur]); return; }
      if (!cur || (counts && counts[cur] > 1)) { cur = nextOrderNo(); changed = true; }
      out.push([cur]);
    });
    if (changed) { var rng = sh.getRange(r1, c.orderno, n, 1); rng.setNumberFormat('@'); rng.setValues(out); }
  });
}

function firstFreeRow(tab) {
  var sh = tab.sheet, last = sh.getLastRow(); if (last < 2) return 2;
  var c = tab.h.cols, p = sh.getRange(2, c.product, last - 1, 1).getValues(), cu = sh.getRange(2, c.customer, last - 1, 1).getValues(), pr = sh.getRange(2, c.price, last - 1, 1).getValues();
  for (var i = last - 2; i >= 0; i--) if (String(p[i][0]).trim() || String(cu[i][0]).trim() || String(pr[i][0]).trim()) return i + 3;
  return 2;
}
/** Formulas for DC / Profit copied from your latest row (so new orders calculate exactly like yours). */
function templateFormulas(tab, beforeRow) {
  var c = tab.h.cols, sh = tab.sheet, out = {};
  ['dc', 'profit'].forEach(function (k) {
    if (!c[k]) return;
    for (var r = beforeRow - 1; r >= Math.max(2, beforeRow - 60); r--) {
      var f = sh.getRange(r, c[k]).getFormulaR1C1();
      if (f && String(f).charAt(0) === '=') { out[k] = f; return; }
    }
  });
  return out;
}

// ===================== Portal =====================
/**
 * Portal.gs — everything the salesman and admin screens need. Each function re-checks the role/ownership itself.
 * A salesman's rows are the rows whose "Saleman" cell equals the account's display name (case/space-insensitive).
 */
function activeRiders() { return readTable('Riders').filter(function (r) { return r.Active !== 'No'; }); }
function riderPhoneOf(name, riders) { var n = normName(name); for (var i = 0; i < riders.length; i++) if (normName(riders[i].Name) === n) return riders[i].Phone || ''; return null; }

function mineOf(ctx, all) { var n = normName(ctx.user.displayName); return all.filter(function (o) { return normName(o.salesman) === n; }); }

function salesmanView(o, riders) {
  var rp = riderPhoneOf(o.courier, riders);
  return { ref: o.ref, orderNo: o.orderNo, tab: o.tab, product: o.product, price: o.price, customer: o.customer, address: o.address, phone: o.phone,
           courier: o.courier, isRider: rp !== null, riderPhone: rp || '', tracking: o.tracking, status: o.status, date: o.date,
           comm: o.status === 'Delivered' ? o.comm : 0, paid: o.paid && o.status === 'Delivered', paidDate: o.spayDate, statusDate: o.sdate, canUpdate: o.status === 'Pending' };
}
function adminView(o) {
  return { ref: o.ref, orderNo: o.orderNo, tab: o.tab, row: o.row, product: o.product, price: o.price, customer: o.customer, address: o.address, phone: o.phone, salesman: o.salesman,
           courier: o.courier, tracking: o.tracking, status: o.status, statusRaw: o.statusRaw, payment: o.payment, date: o.date, pcost: o.pcost, ads: o.ads, dc: o.dc,
           comm: o.comm, commRaw: o.commRaw, profit: o.profit, spay: o.spay, paid: o.paid, spayDate: o.spayDate, sdate: o.sdate };
}

/** What the hired STAFF login may see: no costs, ads, DC, commission, profit or salesman-payment data. */
function staffView(o) {
  return { ref: o.ref, orderNo: o.orderNo, tab: o.tab, row: o.row, product: o.product, price: o.price, customer: o.customer, address: o.address, phone: o.phone, salesman: o.salesman,
           courier: o.courier, tracking: o.tracking, status: o.status, statusRaw: o.statusRaw, payment: o.payment, date: o.date, sdate: o.sdate };
}
function viewFor(ctx, o) { return ctx.user.role === 'admin' ? adminView(o) : staffView(o); }
function statusSwitchOn() { return String(getSettings().SalesmanCanChangeStatus) === 'Yes'; }

function statsOf(list) {
  var cs = cycleStart(), s = { total: list.length, pending: 0, delivered: 0, returns: 0, cancelled: 0, earned: 0, paid: 0, unpaid: 0, unpaidCount: 0, cycleDelivered: 0, cycleEarned: 0, cycleReturns: 0, sales: 0, profit: 0 };
  list.forEach(function (o) {
    if (o.status === 'Pending') s.pending++;
    else if (o.status === 'Cancel') s.cancelled++;
    else if (o.status === 'Return') { s.returns++; if (o.sdate && o.sdate >= cs) s.cycleReturns++; }
    else {
      s.delivered++; s.earned += o.comm; s.sales += o.price;
      if (o.paid) s.paid += o.comm; else { s.unpaid += o.comm; s.unpaidCount++; }
      if (o.sdate && o.sdate >= cs) { s.cycleDelivered++; s.cycleEarned += o.comm; }
    }
    if (o.status === 'Delivered' || o.status === 'Return') s.profit += money0(o.profit);
  });
  ['earned', 'paid', 'unpaid', 'cycleEarned', 'sales', 'profit'].forEach(function (k) { s[k] = round2(s[k]); });
  return s;
}

function filterList(list, f, isAdmin) {
  var q = String(f.q || '').trim().toLowerCase(), ph = q.replace(/\D/g, '');
  return list.filter(function (o) {
    if (f.status && o.status !== f.status) return false;
    if (f.courier && normName(o.courier) !== normName(f.courier)) return false;
    if (f.tab && o.tab !== f.tab) return false;
    if (isAdmin && f.salesman && normName(o.salesman) !== normName(f.salesman)) return false;
    if (f.pay === 'paid' && !(o.status === 'Delivered' && o.paid)) return false;
    if (f.pay === 'unpaid' && !(o.status === 'Delivered' && !o.paid)) return false;
    if (isAdmin && f.cpay === 'paid' && o.payment.toLowerCase() !== 'paid') return false;
    if (isAdmin && f.cpay === 'unpaid' && o.payment.toLowerCase() === 'paid') return false;
    if (q && (o.customer + ' ' + o.tracking + ' ' + o.product + ' ' + o.address + ' ' + o.orderNo).toLowerCase().indexOf(q) < 0 && !(ph.length >= 4 && o.phone.replace(/\D/g, '').indexOf(ph) >= 0)) return false;
    return true;
  });
}
function cleanFilter(f) {
  f = f || {}; var o = {};
  ['q', 'status', 'courier', 'tab', 'salesman', 'pay', 'cpay'].forEach(function (k) { o[k] = String(f[k] == null ? '' : f[k]).slice(0, 80); });
  return o;
}
function paginate(items, page, size, max) {
  size = Math.min(Math.max(Number(size) || CFG.PAGE_SIZE, 1), max || 100);
  var pages = Math.max(1, Math.ceil(items.length / size)), pg = Math.min(Math.max(Number(page) || 1, 1), pages);
  return { items: items.slice((pg - 1) * size, pg * size), page: pg, pages: pages, total: items.length };
}

/* ======================= SALESMAN ======================= */
function salesmanDashboard(ctx) {
  var mine = mineOf(ctx, readAllOrders()), s = statsOf(mine), cs = cycleStart();
  var recent = {}; mine.forEach(function (o) { if (o.paid && o.status === 'Delivered' && o.spayDate) recent[o.spayDate] = (recent[o.spayDate] || 0) + o.comm; });
  var pays = Object.keys(recent).sort().reverse().slice(0, 5).map(function (d) { return { date: d, amount: round2(recent[d]) }; });
  var info = getInfo();
  return { stats: s, cycleStart: cs, recentPayments: pays, businessPhone: info.businessPhone, name: ctx.user.displayName };
}
function salesmanOrders(ctx, p) {
  var f = cleanFilter(p.filter), riders = activeRiders(), mine = mineOf(ctx, readAllOrders());
  var fl = sortNewest(filterList(mine, f, false)), res = paginate(fl, p.page, p.pageSize, 100);
  res.commTotal = round2(fl.reduce(function (a, o) { return a + (o.status === 'Delivered' ? o.comm : 0); }, 0));
  var on = statusSwitchOn();
  res.items = res.items.map(function (o) { var v = salesmanView(o, riders); v.canUpdate = v.canUpdate && on; return v; });
  res.statusSwitch = on;
  res.tabs = orderTabs().map(function (t) { return t.name; }).reverse();
  res.couriers = distinct(mine.map(function (o) { return o.courier; }));
  return res;
}
function distinct(arr) { var seen = {}, out = []; arr.forEach(function (x) { var k = normName(x); if (x && !seen[k]) { seen[k] = 1; out.push(x); } }); return out.sort(); }
function salesmanSetStatus(ctx, p) {
  if (!statusSwitchOn()) throw new AppError('Abhi status badalna band hai. Status admin/staff update karte hain.', 'LOCKED');
  var st = String(p.status || '');
  if (st !== 'Delivered' && st !== 'Return') throw new AppError('Status sirf Delivered ya Return ho sakta hai.');
  var o = withLock(function () {
    var loc = locate(p.ref), o = loc.o;
    if (normName(o.salesman) !== normName(ctx.user.displayName)) throw new AppError('Order nahi mila.', 'NOTFOUND');   // same message as a missing order
    if (o.status !== 'Pending') throw new AppError('Yeh order pehle se ' + o.status + ' hai. Ab sirf admin badal sakta hai.', 'LOCKED');
    paintStatus(loc.tab, o.row, st, nowStr());
    return readOneRow(loc.tab, o.row);
  });
  audit(ctx, 'STATUS_SET', 'Order', o.tab + '#' + o.row, st);
  return salesmanView(o, activeRiders());
}
function salesmanPayments(ctx) {
  var mine = mineOf(ctx, readAllOrders()).filter(function (o) { return o.status === 'Delivered'; }), by = {};
  mine.forEach(function (o) { if (o.paid) { var d = o.spayDate || 'Date nahi likhi'; (by[d] = by[d] || { date: d, amount: 0, count: 0 }); by[d].amount += o.comm; by[d].count++; } });
  var rows = Object.keys(by).sort().reverse().map(function (d) { by[d].amount = round2(by[d].amount); return by[d]; });
  var s = statsOf(mine);
  return { payments: rows, unpaid: s.unpaid, unpaidCount: s.unpaidCount, paidTotal: s.paid };
}

/* ======================= INFO (riders, policy, commission guide) ======================= */
function getInfo() {
  var s = getSettings(), guide = {}, order = [];
  readTable('CommissionGuide').sort(function (a, b) { return (a.Sort || 0) - (b.Sort || 0); }).forEach(function (g) {
    if (!guide[g.Product]) { guide[g.Product] = []; order.push(g.Product); }
    guide[g.Product].push({ slab: g.PriceSlab, comm: g.Commission });
  });
  return { policy: s.PolicyText, businessPhone: s.BusinessPhone, businessName: s.BusinessName,
           riders: activeRiders().map(function (r) { return { id: r.RiderID, name: r.Name, phone: r.Phone, note: r.Note }; }),
           guide: order.map(function (p) { return { product: p, slabs: guide[p] }; }) };
}

/* ======================= ADMIN ======================= */
function adminMeta() {
  var tabs = orderTabs(), all = readAllOrders(), cur = tabs.length ? tabs[tabs.length - 1] : null;
  var salesmen = cur ? dropdownValues(cur, 'salesman', all) : [], couriers = cur ? dropdownValues(cur, 'courier', all) : [];
  var nowKey = Number(Utilities.formatDate(new Date(), CFG.TZ, 'yyyy')) * 100 + Number(Utilities.formatDate(new Date(), CFG.TZ, 'M')) - 1;
  return { tabsBeforeCurrent: tabs.filter(function (t) { return t.key < nowKey; }).map(function (t) { return t.name; }).reverse(), statusSwitch: statusSwitchOn(),
           tabs: tabs.map(function (t) { return t.name; }).reverse(), salesmen: salesmen, couriers: couriers,
           allSalesmen: distinct(all.map(function (o) { return o.salesman; })), hasPayCols: !!(cur && cur.h.cols.spay && cur.h.cols.spaydate) };
}
/* ---------- graphs data (dashboard) ---------- */
function orderDay(o) {
  var s = String(o.date || '').trim(), m;
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return m[1] + '-' + pad2(m[2]) + '-' + pad2(m[3]);
  if ((m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/.exec(s))) { var y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + pad2(m[2]) + '-' + pad2(m[1]); }
  if (s) { var d = new Date(s); if (!isNaN(d.getTime())) return Utilities.formatDate(d, CFG.TZ, 'yyyy-MM-dd'); }
  var sd = String(o.sdate || '').slice(0, 10); return /^\d{4}-\d{2}-\d{2}$/.test(sd) ? sd : '';
}
function buildCharts(all, withMoney) {
  var cm = {}, dm = {}, mm = {}, undated = 0;
  function bump(x, o) { x.total++; if (o.status === 'Delivered') x.delivered++; else if (o.status === 'Return') x.returns++; else if (o.status === 'Cancel') x.cancel++; else x.pending++; }
  all.forEach(function (o) {
    var cn = o.courier || '(Courier nahi)', cx = cm[cn] = cm[cn] || { name: cn, total: 0, delivered: 0, returns: 0, pending: 0, cancel: 0 }; bump(cx, o);
    var day = orderDay(o); if (day) { var dx = dm[day] = dm[day] || { d: day, total: 0, delivered: 0, returns: 0, pending: 0, cancel: 0 }; bump(dx, o); } else undated++;
    var mx = mm[o.tab] = mm[o.tab] || { tab: o.tab, total: 0, delivered: 0, returns: 0, pending: 0, cancel: 0, sales: 0, profit: 0 }; bump(mx, o);
    if (o.status === 'Delivered') mx.sales += o.price;
    if (o.status === 'Delivered' || o.status === 'Return') mx.profit += money0(o.profit);
  });
  var days = [];
  for (var i = 13; i >= 0; i--) { var ds = Utilities.formatDate(new Date(Date.now() - i * 86400000), CFG.TZ, 'yyyy-MM-dd'); days.push(dm[ds] || { d: ds, total: 0, delivered: 0, returns: 0, pending: 0, cancel: 0 }); }
  var couriers = Object.keys(cm).map(function (k) { return cm[k]; }).sort(function (a, b) { return b.total - a.total; }).slice(0, 8);
  function mkey(t) { var q = /^([A-Z]+)\s+(\d{4})$/.exec(t); return q ? Number(q[2]) * 12 + MONTHS_UP.indexOf(q[1]) : 0; }
  var months = Object.keys(mm).map(function (k) { var x = mm[k]; x.sales = round2(x.sales); x.profit = round2(x.profit); if (!withMoney) { delete x.sales; delete x.profit; } return x; }).sort(function (a, b) { return mkey(a.tab) - mkey(b.tab); }).slice(-6);
  return { days: days, couriers: couriers, months: months, undated: undated };
}
function adminDashboard(ctx, p) {
  var tab = String(p.tab || ''), all = readAllOrders().filter(function (o) { return !tab || o.tab === tab; }), s = statsOf(all);
  var users = {}; readTable('Users').forEach(function (u) { if (u.Role === 'salesman') users[normName(u.DisplayName)] = 1; });
  var by = {}, names = {};
  all.forEach(function (o) { var k = normName(o.salesman) || '(blank)'; (by[k] = by[k] || []).push(o); names[k] = names[k] || o.salesman || '(blank)'; });
  var rows = Object.keys(by).map(function (k) { var x = statsOf(by[k]); x.name = names[k]; x.hasLogin = !!users[k]; return x; }).sort(function (a, b) { return b.unpaid - a.unpaid; });
  var cp = 0, cu = 0; all.forEach(function (o) { if (o.status !== 'Cancel' && o.payment.toLowerCase() !== 'paid' && o.status === 'Delivered') { cp += o.price; cu++; } });
  return { stats: s, salesmen: rows, cycleStart: cycleStart(), courierDue: { count: cu, amount: round2(cp) }, charts: buildCharts(all, true) };
}
function adminOrders(ctx, p) {
  var f = cleanFilter(p.filter); if (ctx.user.role !== 'admin') f.pay = '';          // staff never filter by salesman payment
  var all0 = readAllOrders();
  if (ctx.user.role === 'admin' && all0.some(function (o) { return o.hasNoCol && !o.orderNo; })) { backfillOrderNumbers(); all0 = readAllOrders(); }   // give new rows their Order No
  var list = sortNewest(filterList(all0, f, true));
  var res = p.forExport ? { items: list.slice(0, 5000), page: 1, pages: 1, total: list.length } : paginate(list, p.page, p.pageSize, 100);
  res.items = res.items.map(function (o) { return viewFor(ctx, o); });
  if (ctx.user.role === 'admin') res.commTotal = round2(list.reduce(function (a, o) { return a + (o.status === 'Delivered' ? o.comm : 0); }, 0));
  return res;
}

var ORDER_FIELDS = ['Product', 'Price', 'CustomerName', 'Address', 'Phone', 'Salesman', 'Courier', 'Tracking', 'Status', 'Payment', 'Date', 'ProductCost', 'AdsCost', 'Comm'];
function cleanText(v, max, label, req) {
  var x = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  if (req && !x) throw new AppError(label + ' zaroori hai.', 'VALIDATION');
  if (x.length > max) throw new AppError(label + ' bohat lamba hai.', 'VALIDATION');
  return x;
}
function pickFromList(val, values, label, req) {
  var x = cleanText(val, 80, label, req); if (!x) return '';
  if (!values.length) return x;
  for (var i = 0; i < values.length; i++) if (normName(values[i]) === normName(x)) return values[i];
  throw new AppError(label + ' "' + x + '" dropdown mein nahi hai.', 'VALIDATION');
}
function optMoney(v, label) {
  var n = parseMoney(v); if (n === null) return '';
  if (isNaN(n) || n < 0) throw new AppError(label + ' sahi number hona chahiye.', 'VALIDATION');
  return n;
}
function adminSaveOrder(ctx, p) {
  var isEdit = !!p.ref, staff = ctx.user.role !== 'admin';
  var result = withLock(function () {
    var tab, o = null, free = 0, loc;
    if (isEdit) { loc = locate(p.ref); tab = loc.tab; o = loc.o; free = o.row; } else { tab = currentMonthTab(); free = firstFreeRow(tab); }
    var all = readAllOrders(), sal = dropdownValues(tab, 'salesman'), cou = dropdownValues(tab, 'courier')   /* sirf sheet ka asli dropdown rokta hai; warna koi bhi naam likh sakte hain */, c = tab.h.cols, sh = tab.sheet;
    var price = parseMoney(p.Price); if (price === null || isNaN(price) || price <= 0) throw new AppError('Price sahi likhein.', 'VALIDATION');
    var status = String(p.Status || '').trim() ? normStatus(p.Status) : 'Pending';
    var v = { product: cleanText(p.Product, 120, 'Product', true), customer: cleanText(p.CustomerName, 100, 'Customer name', true), address: cleanText(p.Address, 300, 'Address', true),
              phone: cleanText(p.Phone, 60, 'Phone', true), salesman: pickFromList(p.Salesman, sal, 'Saleman', true), courier: pickFromList(p.Courier, cou, 'Courier', false),
              tracking: cleanText(p.Tracking, 60, 'Tracking', false), date: cleanText(p.Date, 30, 'Date', false), payment: /^paid$/i.test(String(p.Payment || '').trim()) ? 'PAID' : '',
              pcost: optMoney(p.ProductCost, 'Product cost'), ads: optMoney(p.AdsCost, 'Ads cost'), comm: optMoney(p.Comm, 'Commission') };
    if (staff) {   // staff may not touch money columns: keep what is already there (or blank for a new order)
      var oldO = o || {};
      v.pcost = (oldO.pcost === undefined || parseMoney(oldO.pcost) === null) ? '' : parseMoney(oldO.pcost);
      v.ads = (oldO.ads === undefined || parseMoney(oldO.ads) === null) ? '' : parseMoney(oldO.ads);
      v.comm = (oldO.commRaw === undefined || parseMoney(oldO.commRaw) === null) ? '' : parseMoney(oldO.commRaw);
    }
    var r = free, ed = isEdit;
    var txt = function (col, nv, ov) { if (!ed || String(nv) !== String(ov)) setTextCell(sh, r, col, nv); };
    var num = function (col, nv, ovRaw) { var on = parseMoney(ovRaw); if (!ed || (nv === '' ? String(ovRaw).trim() !== '' : on !== nv)) setNumCell(sh, r, col, nv); };
    var oo = o || {};
    txt(c.product, v.product, oo.product); num(c.price, price, oo.price === 0 ? '' : oo.price); txt(c.customer, v.customer, oo.customer); txt(c.address, v.address, oo.address);
    txt(c.phone, v.phone, oo.phone); txt(c.salesman, v.salesman, oo.salesman); txt(c.courier, v.courier, oo.courier); txt(c.tracking, v.tracking, oo.tracking);
    txt(c.date, v.date, oo.date); num(c.pcost, v.pcost, oo.pcost); num(c.ads, v.ads, oo.ads); num(c.comm, v.comm, oo.commRaw);
    if (!ed || (oo.payment.toLowerCase() === 'paid') !== (v.payment === 'PAID')) paintPaymentCell(tab, r, v.payment === 'PAID');
    if (!ed || oo.status !== status) paintStatus(tab, r, status, nowStr());
    if (!isEdit) {
      if (c.orderno) setTextCell(sh, r, c.orderno, nextOrderNo());
      var tf = templateFormulas(tab, r);
      ['dc', 'profit'].forEach(function (k) { if (c[k] && tf[k] && !String(sh.getRange(r, c[k]).getFormulaR1C1() || '')) sh.getRange(r, c[k]).setFormulaR1C1(tf[k]); });
      if (c.spay) { sh.getRange(r, c.spay).setValue(''); sh.getRange(r, c.spay).setBackground(WHITE); }
    } else if (!staff && c.spay && c.spaydate && typeof p.SalesmanPaid === 'string') {
      var wantPaid = /^paid$/i.test(p.SalesmanPaid);
      if (wantPaid !== o.paid) { var pd = String(p.SalesmanPaidDate || todayStr()); if (wantPaid && !isDate(pd)) throw new AppError('Payment date sahi likhein (yyyy-mm-dd).', 'VALIDATION'); writeSalesmanPaid(tab, [r], wantPaid, pd); }
    }
    return { tab: tab, row: r, was: o };
  });
  audit(ctx, isEdit ? 'ORDER_EDITED' : 'ORDER_ADDED', 'Order', result.tab.name + '#' + result.row, 'status=' + (p.Status || 'Pending') + ' salesman=' + p.Salesman);
  return viewFor(ctx, readOneRow(result.tab, result.row));
}
function adminSetStatus(ctx, p) {
  var st = String(p.status || ''); if (['Pending', 'Delivered', 'Return', 'Cancel'].indexOf(st) < 0) throw new AppError('Status galat hai.');
  var o = withLock(function () {
    var loc = locate(p.ref); paintStatus(loc.tab, loc.o.row, st, (loc.o.status === st && loc.o.sdate) ? loc.o.sdate : nowStr()); return readOneRow(loc.tab, loc.o.row);
  });
  audit(ctx, 'STATUS_SET', 'Order', o.tab + '#' + o.row, st);
  return viewFor(ctx, o);
}
function adminSetCourierPaid(ctx, p) {
  var paid = p.paid === true;
  var o = withLock(function () { var loc = locate(p.ref); paintPaymentCell(loc.tab, loc.o.row, paid); return readOneRow(loc.tab, loc.o.row); });
  audit(ctx, paid ? 'COURIER_PAID' : 'COURIER_UNPAID', 'Order', o.tab + '#' + o.row, '');
  return viewFor(ctx, o);
}

/** Counts only (no money) for the staff login. */
function staffDashboard(ctx, p) {
  var tab = String(p.tab || ''), all = readAllOrders().filter(function (o) { return !tab || o.tab === tab; }), s = { total: all.length, pending: 0, delivered: 0, returns: 0, cancelled: 0 }, by = {};
  all.forEach(function (o) {
    if (o.status === 'Pending') { s.pending++; var k = o.courier || '(courier nahi)'; by[k] = (by[k] || 0) + 1; } else if (o.status === 'Delivered') s.delivered++; else if (o.status === 'Return') s.returns++; else s.cancelled++;
  });
  var pend = Object.keys(by).map(function (k) { return { courier: k, count: by[k] }; }).sort(function (a, b) { return b.count - a.count; });
  return { stats: s, pendingByCourier: pend, charts: buildCharts(all, false) };
}
function adminSetStatusSwitch(ctx, p) {
  var on = p.on === true;
  withLock(function () { upsertSetting('SalesmanCanChangeStatus', on ? 'Yes' : 'No'); });
  audit(ctx, on ? 'SALESMAN_STATUS_ON' : 'SALESMAN_STATUS_OFF', 'Settings', 'SalesmanCanChangeStatus', '');
  return { on: on };
}

function adminCreateMonthTab(ctx, p) {
  var t = withLock(function () { return createMonthTab(p.name); });
  audit(ctx, 'MONTH_TAB_CREATED', 'Tab', t.name, '');
  return { name: t.name };
}

/* ---------- salesman payments ---------- */
function unpaidOrders(name) {
  var n = normName(name), key = {}; orderTabs().forEach(function (t) { key[t.name] = t.key; });
  return readAllOrders().filter(function (o) { return o.status === 'Delivered' && !o.paid && o.comm > 0 && normName(o.salesman) === n; })
    .sort(function (a, b) { return (key[a.tab] - key[b.tab]) || (a.row - b.row); });
}
function adminPayments(ctx) {
  var d = adminDashboard(ctx, {}), logs = readTable('PayLog').sort(function (a, b) { return a.CreatedAt < b.CreatedAt ? 1 : -1; }).slice(0, 40)
    .map(function (l) { delete l._row; delete l.Refs; return l; });
  return { salesmen: d.salesmen.filter(function (s) { return s.name !== '(blank)'; }), cycleStart: d.cycleStart, logs: logs };
}
function adminPayDetail(ctx, p) {
  var list = unpaidOrders(p.salesman), cs = cycleStart(), total = 0;
  var items = list.map(function (o) { total += o.comm; return { ref: o.ref, orderNo: o.orderNo, tab: o.tab, product: o.product, customer: o.customer, phone: o.phone, tracking: o.tracking, price: o.price, comm: o.comm, delivered: String(o.sdate || '').slice(0, 10) }; });
  return { salesman: p.salesman, items: items, total: round2(total) };
}
function adminPayConfirm(ctx, p) {
  var reqId = String(p.requestId || ''); if (!/^[A-Za-z0-9-]{8,60}$/.test(reqId)) throw new AppError('Request galat hai. Page refresh karein.');
  var pdate = String(p.paymentDate || ''); if (!isDate(pdate) || pdate > fromDayNum(dayNum(todayStr()) + 1)) throw new AppError('Payment date sahi nahi hai.');
  var refs = (p.refs || []).slice(0, 600), note = String(p.note || '').slice(0, 200), name = String(p.salesman || '');
  if (!refs.length) throw new AppError('Koi order select nahi kiya.');
  var res = withLock(function () {
    var dup = readTable('PayLog').filter(function (l) { return l.RequestID === reqId; })[0];
    if (dup) return { log: dup, replayed: true };
    var byTab = {}, total = 0, seen = {};
    refs.forEach(function (ref) {
      if (seen[ref]) return; seen[ref] = 1;
      var loc = locate(ref), o = loc.o;
      if (normName(o.salesman) !== normName(name)) throw new AppError('Ek order is salesman ka nahi hai. Page refresh karein.', 'CONFLICT');
      if (o.status !== 'Delivered' || o.paid) throw new AppError('Ek order pehle se paid hai ya Delivered nahi. Page refresh karein.', 'CONFLICT');
      (byTab[o.tab] = byTab[o.tab] || { tab: loc.tab, rows: [] }).rows.push(o.row); total += o.comm;
    });
    var logId = fmtId('PL-', reserveSeq('LOG_SEQ', 1), 5);
    var log = { LogID: logId, PaidDate: pdate, Salesman: name, OrderCount: Object.keys(seen).length, Amount: round2(total), Note: note, CreatedBy: ctx.user.username, CreatedAt: nowStr(), Status: 'Processing', RequestID: reqId, Refs: JSON.stringify(Object.keys(seen)).slice(0, 45000) };
    appendObjects('PayLog', [log]);
    Object.keys(byTab).forEach(function (k) { writeSalesmanPaid(byTab[k].tab, byTab[k].rows, true, pdate); });
    SpreadsheetApp.flush();
    var chk = readAllOrders().filter(function (o) { return seen[o.ref] && o.paid; }).length;     // verify the write before saying "done"
    if (chk !== log.OrderCount) throw new Error('Payment verification failed for ' + logId);
    var stored = getById('PayLog', logId); stored.Status = 'Paid'; updateObj('PayLog', stored);
    return { log: stored, replayed: false };
  });
  if (!res.replayed) audit(ctx, 'SALESMAN_PAID', 'PayLog', res.log.LogID, name + ' orders=' + res.log.OrderCount + ' amount=' + res.log.Amount);
  var l = res.log; delete l._row; delete l.Refs; return { log: l, replayed: res.replayed };
}
function adminUndoPayment(ctx, p) {
  var out = withLock(function () {
    var log = getById('PayLog', p.logId); if (!log) throw new AppError('Payment record nahi mila.');
    if (log.Status !== 'Paid') throw new AppError('Yeh payment undo nahi ho sakti (status: ' + log.Status + ').');
    var refs = []; try { refs = JSON.parse(log.Refs || '[]'); } catch (e) {}
    if (!refs.length) throw new AppError('Is payment ke orders ki list mehfooz nahi, isliye undo nahi ho sakti.');
    var byTab = {}, undone = 0, skipped = 0;
    refs.forEach(function (ref) {
      try { var loc = locate(ref); if (loc.o.paid && loc.o.spayDate === log.PaidDate) { (byTab[loc.o.tab] = byTab[loc.o.tab] || { tab: loc.tab, rows: [] }).rows.push(loc.o.row); undone++; } else skipped++; } catch (e) { skipped++; }
    });
    Object.keys(byTab).forEach(function (k) { writeSalesmanPaid(byTab[k].tab, byTab[k].rows, false, ''); });
    log.Status = 'Undone'; log.Note = (log.Note + ' | undone ' + undone + ', skipped ' + skipped).slice(0, 200); updateObj('PayLog', log);
    return { undone: undone, skipped: skipped, logId: log.LogID };
  });
  audit(ctx, 'PAYMENT_UNDONE', 'PayLog', out.logId, 'undone=' + out.undone + ' skipped=' + out.skipped);
  return out;
}
/** Dry run for the opening-balance tool: how many orders / how much per salesman would be marked as already paid. */
function adminBulkPreview(ctx, p) {
  var tabsWanted = (p.tabs || []).map(String), only = normName(p.salesman || ''), by = {}, total = 0, count = 0;
  orderTabs().filter(function (t) { return tabsWanted.indexOf(t.name) >= 0; }).forEach(function (t) {
    readTab(t).forEach(function (o) {
      if (o.status !== 'Delivered' || o.paid || !(o.comm > 0)) return;
      if (only && normName(o.salesman) !== only) return;
      var k = o.salesman || '(blank)'; by[k] = by[k] || { name: k, count: 0, amount: 0 }; by[k].count++; by[k].amount += o.comm; total += o.comm; count++;
    });
  });
  return { rows: Object.keys(by).map(function (k) { by[k].amount = round2(by[k].amount); return by[k]; }).sort(function (a, b) { return b.amount - a.amount; }), count: count, amount: round2(total) };
}
/** Opening balance: mark unpaid Delivered orders of chosen (older) tabs as already paid (optionally for one salesman). */
function adminBulkPaid(ctx, p) {
  if (p.confirm !== 'YES') throw new AppError('Confirm ke liye YES likhein.');
  var date = String(p.paymentDate || ''); if (!isDate(date)) throw new AppError('Date sahi likhein (yyyy-mm-dd).');
  var tabsWanted = (p.tabs || []).map(String); if (!tabsWanted.length) throw new AppError('Koi tab select nahi kiya.');
  var out = withLock(function () {
    var total = 0, count = 0, only = normName(p.salesman || ''), tabs = orderTabs().filter(function (t) { return tabsWanted.indexOf(t.name) >= 0; });
    tabs.forEach(function (t) {
      var rows = readTab(t).filter(function (o) { return o.status === 'Delivered' && !o.paid && o.comm > 0 && (!only || normName(o.salesman) === only); });
      rows.forEach(function (o) { total += o.comm; count++; });
      if (rows.length) writeSalesmanPaid(t, rows.map(function (o) { return o.row; }), true, date);
    });
    var logId = fmtId('PL-', reserveSeq('LOG_SEQ', 1), 5);
    appendObjects('PayLog', [{ LogID: logId, PaidDate: date, Salesman: p.salesman ? String(p.salesman).slice(0, 60) : 'ALL', OrderCount: count, Amount: round2(total), Note: 'Opening balance (purani tabs): ' + tabsWanted.join(', ').slice(0, 120), CreatedBy: ctx.user.username, CreatedAt: nowStr(), Status: 'Historical', RequestID: 'bulk-' + rnd(10), Refs: '' }]);
    return { count: count, amount: round2(total), logId: logId };
  });
  audit(ctx, 'BULK_PAID', 'PayLog', out.logId, 'orders=' + out.count + ' amount=' + out.amount);
  return out;
}

/* ---------- riders, policy, settings ---------- */
function adminRiders() { return readTable('Riders').map(function (r) { delete r._row; return r; }); }
function adminSaveRider(ctx, p) {
  var name = cleanText(p.Name, 60, 'Naam', true), phone = cleanText(p.Phone, 60, 'Phone', false), note = cleanText(p.Note, 120, 'Note', false), active = p.Active === 'No' ? 'No' : 'Yes';
  var r = withLock(function () {
    var x;
    if (p.RiderID) { x = getById('Riders', p.RiderID); if (!x) throw new AppError('Rider nahi mila.'); } else x = { RiderID: fmtId('R-', reserveSeq('RIDER_SEQ', 1), 4) };
    x.Name = name; x.Phone = phone; x.Note = note; x.Active = active;
    if (x._row) updateObj('Riders', x); else appendObjects('Riders', [x]); return x;
  });
  audit(ctx, 'RIDER_SAVED', 'Rider', r.RiderID, name);
  return { RiderID: r.RiderID };
}
function adminSavePolicy(ctx, p) {
  var text = String(p.text || '').slice(0, 8000);
  withLock(function () { upsertSetting('PolicyText', text); });
  audit(ctx, 'POLICY_SAVED', 'Settings', 'PolicyText', 'chars=' + text.length);
  return { saved: true };
}
function upsertSetting(key, val) {
  var rows = readTable('Settings'), r = rows.filter(function (x) { return x.Key === key; })[0];
  if (r) { r.Value = val; updateObj('Settings', r); } else appendObjects('Settings', [{ Key: key, Value: val, Description: '' }]);
  resetSettingsCache();
}
var EDITABLE_SETTINGS = ['BusinessName', 'BusinessPhone', 'OrderTabs', 'SessionHoursSalesman', 'SessionHoursAdmin'];
function adminGetSettings() {
  var s = getSettings(); return DEFAULT_SETTINGS.filter(function (d) { return EDITABLE_SETTINGS.indexOf(d[0]) >= 0; }).map(function (d) { return { key: d[0], value: s[d[0]], description: d[2] }; });
}
function adminSaveSettings(ctx, p) {
  var v = p.values || {}, upd = {};
  if ('BusinessName' in v) { var b = String(v.BusinessName).trim(); if (!b || b.length > 60) throw new AppError('Business name 1–60 characters ka ho.'); upd.BusinessName = b; }
  if ('BusinessPhone' in v) { var ph = String(v.BusinessPhone).trim(); if (!/^[0-9+\- ]{6,20}$/.test(ph)) throw new AppError('Phone number sahi likhein.'); upd.BusinessPhone = ph; }
  if ('OrderTabs' in v) { var t = String(v.OrderTabs).trim() || 'AUTO'; if (t.length > 400) throw new AppError('OrderTabs bohat lamba hai.'); upd.OrderTabs = t; }
  if ('CycleCutoffHour' in v) { var h = Number(v.CycleCutoffHour); if (!(h >= 0 && h <= 23 && Math.floor(h) === h)) throw new AppError('Cutoff hour 0 se 23 ho.'); upd.CycleCutoffHour = String(h); }
  ['SessionHoursSalesman', 'SessionHoursAdmin'].forEach(function (k) { if (k in v) { var n = Number(v[k]); if (!(n >= 1 && n <= 72)) throw new AppError(k + ' 1 se 72 ke beech ho.'); upd[k] = String(n); } });
  withLock(function () { Object.keys(upd).forEach(function (k) { upsertSetting(k, upd[k]); }); });
  audit(ctx, 'SETTINGS_CHANGED', 'Settings', '', Object.keys(upd).join(','));
  return adminGetSettings();
}
function adminAudit(ctx, p) {
  var all = readTable('AuditLog').reverse().map(function (a) { delete a._row; return a; }), q = String(p.q || '').toLowerCase().trim();
  if (q) all = all.filter(function (a) { return (a.Action + ' ' + a.ActorName + ' ' + a.EntityID + ' ' + a.Details).toLowerCase().indexOf(q) >= 0; });
  return paginate(all, p.page, 50, 100);
}

// ===================== Code =====================
/**
 * Code.gs — web entry point and the single server API used by the browser: api(token, action, payload).
 * Each action declares the role it needs; the role comes from the server-side session, never from the browser.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Z. STORE Portal').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

var ACTIONS = {
  login:            { role: 'public', fn: function (c, p) { return login(p); } },
  logout:           { role: 'any', mustChangeOk: true, fn: function (c) { return logout(c); } },
  meta:             { role: 'any', mustChangeOk: true, fn: function (c) { return buildMeta(c); } },
  changePassword:   { role: 'any', mustChangeOk: true, fn: function (c, p) { return changePassword(c, p); } },
  info:             { role: 'any', fn: function () { return getInfo(); } },

  sm_dashboard:     { role: 'salesman', fn: function (c) { return salesmanDashboard(c); } },
  sm_orders:        { role: 'salesman', fn: function (c, p) { return salesmanOrders(c, p); } },
  sm_setStatus:     { role: 'salesman', fn: function (c, p) { return salesmanSetStatus(c, p); } },
  sm_payments:      { role: 'salesman', fn: function (c) { return salesmanPayments(c); } },

  ad_dashboard:     { role: 'admin', fn: function (c, p) { return adminDashboard(c, p); } },
  ad_orders:        { roles: ['admin', 'staff'], fn: function (c, p) { return adminOrders(c, p); } },
  ad_saveOrder:     { roles: ['admin', 'staff'], fn: function (c, p) { return adminSaveOrder(c, p); } },
  ad_setStatus:     { roles: ['admin', 'staff'], fn: function (c, p) { return adminSetStatus(c, p); } },
  ad_setCourierPaid:{ roles: ['admin', 'staff'], fn: function (c, p) { return adminSetCourierPaid(c, p); } },
  ad_createMonthTab:{ role: 'admin', fn: function (c, p) { return adminCreateMonthTab(c, p); } },
  ad_payments:      { role: 'admin', fn: function (c) { return adminPayments(c); } },
  ad_payDetail:     { role: 'admin', fn: function (c, p) { return adminPayDetail(c, p); } },
  ad_payConfirm:    { role: 'admin', fn: function (c, p) { return adminPayConfirm(c, p); } },
  ad_undoPayment:   { role: 'admin', fn: function (c, p) { return adminUndoPayment(c, p); } },
  st_dashboard:     { roles: ['admin', 'staff'], fn: function (c, p) { return staffDashboard(c, p); } },
  ad_setStatusSwitch:{ role: 'admin', fn: function (c, p) { return adminSetStatusSwitch(c, p); } },
  ad_bulkPreview:   { role: 'admin', fn: function (c, p) { return adminBulkPreview(c, p); } },
  ad_bulkPaid:      { role: 'admin', fn: function (c, p) { return adminBulkPaid(c, p); } },
  ad_riders:        { role: 'admin', fn: function () { return adminRiders(); } },
  ad_saveRider:     { role: 'admin', fn: function (c, p) { return adminSaveRider(c, p); } },
  ad_savePolicy:    { role: 'admin', fn: function (c, p) { return adminSavePolicy(c, p); } },
  ad_users:         { role: 'admin', fn: function () { return adminListUsers(); } },
  ad_createUser:    { role: 'admin', fn: function (c, p) { return adminCreateUser(c, p); } },
  ad_updateUser:    { role: 'admin', fn: function (c, p) { return adminUpdateUser(c, p); } },
  ad_resetPassword: { role: 'admin', fn: function (c, p) { return adminResetPassword(c, p); } },
  ad_setActive:     { role: 'admin', fn: function (c, p) { return adminSetActive(c, p); } },
  ad_revokeSessions:{ role: 'admin', fn: function (c, p) { return adminRevokeSessions(c, p); } },
  ad_settings:      { role: 'admin', fn: function () { return adminGetSettings(); } },
  ad_saveSettings:  { role: 'admin', fn: function (c, p) { return adminSaveSettings(c, p); } },
  ad_audit:         { role: 'admin', fn: function (c, p) { return adminAudit(c, p); } }
};

function api(token, action, payload) {
  try {
    var def = Object.prototype.hasOwnProperty.call(ACTIONS, action) ? ACTIONS[action] : null;
    if (!def) throw new AppError('Unknown request.', 'BAD');
    if (payload === null || payload === undefined) payload = {};
    if (typeof payload !== 'object' || Array.isArray(payload)) throw new AppError('Bad request.', 'BAD');
    var ctx = null;
    if (def.role !== 'public') {
      ctx = authenticate(token);
      var allowed = def.roles ? def.roles.indexOf(ctx.user.role) >= 0 : (def.role === 'any' || def.role === ctx.user.role);
      if (!allowed) throw new AppError('Aap ko is ki ijazat nahi.', 'FORBIDDEN');
      if (ctx.user.mustChange && !def.mustChangeOk) throw new AppError('Pehle naya password rakhein.', 'MUSTCHANGE');
    }
    return { ok: true, data: def.fn(ctx, payload) };
  } catch (err) {
    if (err && err.isApp) return { ok: false, error: err.message, code: err.code };
    var ref = Utilities.getUuid().slice(0, 8);
    console.error('[' + ref + '] ' + action + ': ' + (err && err.stack ? err.stack : err));
    return { ok: false, error: 'Server par masla aa gaya. Reference ' + ref + '.', code: 'SERVER' };
  }
}

function buildMeta(ctx) {
  var s = getSettings(), m = { business: s.BusinessName, businessPhone: s.BusinessPhone, demo: isDemo(), today: todayStr(),
    user: { userId: ctx.user.userId, username: ctx.user.username, role: ctx.user.role, displayName: ctx.user.displayName, mustChange: ctx.user.mustChange } };
  if (ctx.user.role === 'admin' || ctx.user.role === 'staff') { var a = adminMeta(); m.tabs = a.tabs; m.salesmen = a.salesmen; m.couriers = a.couriers; m.allSalesmen = a.allSalesmen; m.hasPayCols = a.hasPayCols; m.tabsBeforeCurrent = a.tabsBeforeCurrent; m.statusSwitch = a.statusSwitch; }
  return m;
}

// ===================== Setup =====================
/**
 * Setup.gs — run from the Apps Script editor (never from the web app).
 *   1. setupLite()          adds helper tabs + 3 helper columns to your month tabs. Safe to run again; never deletes data.
 *   2. initializeAdmin()    creates the first admin from Script properties, then deletes the password property.
 *   3. validateLite()       read-only health check.
 * Also: resetAdminPassword(), backupCopy(), installOrderTrigger().
 */
var NEW_COLS = [['spay', 'Saleman Payment'], ['spaydate', 'Saleman Payment Date'], ['sdate', 'Status Date'], ['orderno', 'Order No']];

function setupLite() {
  var log = [], book = ss();
  Object.keys(SCHEMA).forEach(function (name) {
    var c = cols(name), s = book.getSheetByName(name);
    if (!s) { s = book.insertSheet(name); log.push('Naya tab bana: ' + name); }
    var head = s.getRange(1, 1, 1, c.length).getValues()[0];
    if (head.every(function (h) { return h === ''; })) {
      s.getRange(1, 1, 1, c.length).setValues([c.map(function (x) { return x.name; })]);
      s.setFrozenRows(1);
      s.getRange(1, 1, 1, c.length).setFontWeight('bold').setBackground('#111111').setFontColor('#C9A24B');
      applyFormats(name, 2, Math.max(s.getMaxRows() - 1, 1));
    } else c.forEach(function (x, i) {
      if (String(head[i]) !== x.name) throw new Error('Tab "' + name + '" column ' + (i + 1) + ' "' + head[i] + '" hai, "' + x.name + '" chahiye. Header manually theek karein.');
    });
  });
  if (!prop('PEPPER')) { setProp('PEPPER', rnd(64) + rnd(64)); log.push('PEPPER ban gaya (Project Settings > Script properties mein hai; kisi safe jagah copy kar lein)'); }
  setProp('SPREADSHEET_ID', book.getId());
  var have = {}; readTable('Settings').forEach(function (r) { have[r.Key] = 1; });
  var add = DEFAULT_SETTINGS.filter(function (d) { return !have[d[0]]; }).map(function (d) { return { Key: d[0], Value: d[1], Description: d[2] }; });
  if (add.length) appendObjects('Settings', add);
  resetSettingsCache();
  seedRiders(log); seedGuide(log);
  ['USER_SEQ', 'RIDER_SEQ', 'GUIDE_SEQ', 'LOG_SEQ'].forEach(function (k) { if (!prop(k)) setProp(k, '0'); });
  var tabs = orderTabs();
  if (!tabs.length) throw new Error('Koi order tab nahi mila. Tab ka naam "OCTOBER 2026" jaisa aur header mein Product, Price, Customer Name, Saleman, Status hona chahiye.');
  tabs.forEach(function (t) { ensureHelperColumns(t, log); });
  installOrderTrigger(); log.push('Edit trigger lag gaya (Status Date aur Order No khud bhar jate hain)');
  var numbered = backfillOrderNumbers(); log.push('Order No: ' + numbered + ' orders ko number mil gaya');
  var all = readAllOrders(), st = statsOf(all);
  log.push('Order tabs: ' + orderTabs().map(function (t) { return t.name; }).join(', '));
  log.push('Total orders: ' + all.length + ' | Delivered: ' + st.delivered + ' | Salesman payment abhi blank (unpaid maani jayengi): ' + st.unpaidCount + ' orders, Rs ' + st.unpaid);
  log.push('IMPORTANT: purane (pehle se paid) mahino ke orders ko Admin > Payments > "Purana hisaab" se Paid mark karein.');
  log.push('setupLite finished.');
  Logger.log(log.join('\n'));
  return log;
}

function ensureHelperColumns(tab, log) {
  var sh = tab.sheet, headRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0], lastHead = 0;
  headRow.forEach(function (h, i) { if (String(h).trim() !== '') lastHead = i + 1; });
  var last = Math.max(sh.getLastRow(), 1), changed = false;
  NEW_COLS.forEach(function (nc) {
    if (tab.h.cols[nc[0]]) return;
    var col = lastHead + 1;
    for (var tries = 0; tries < 10; tries++, col++) {
      var vals = sh.getRange(1, col, last, 1).getValues();
      if (vals.every(function (r) { return r[0] === ''; })) break;
    }
    if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
    var hc = sh.getRange(1, col);
    if (tab.h.cols.profit) sh.getRange(1, tab.h.cols.profit).copyTo(hc);
    hc.setValue(nc[1]); lastHead = col; changed = true;
    var body = sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1);
    if (nc[0] !== 'spay') body.setNumberFormat('@');
    else body.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Paid'], true).setAllowInvalid(true).build());
    log.push(tab.name + ': column "' + nc[1] + '" add hui (' + colLetter(col) + ')');
  });
  if (changed) tab.h = headerMapOf(sh);
}

function installOrderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'onOrderEdit') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onOrderEdit').forSpreadsheet(ss()).onEdit().create();
}
/** Installable edit trigger: stamps "Status Date" when you type Delivered / Return / Cancel in the sheet. */
function onOrderEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet(), tab = orderTabs().filter(function (t) { return t.name === sh.getName(); })[0];
    if (!tab) return;
    var c = tab.h.cols, r = e.range, r1 = Math.max(2, r.getRow()), r2 = Math.min(r.getLastRow(), r1 + 499);
    if (r2 < r1) return;
    if (c.status && c.sdate && r.getColumn() <= c.status && r.getLastColumn() >= c.status) {
      var st = sh.getRange(r1, c.status, r2 - r1 + 1, 1).getValues(), rng = sh.getRange(r1, c.sdate, r2 - r1 + 1, 1), sd = rng.getValues(), now = nowStr(), changed = false;
      for (var i = 0; i < st.length; i++) {
        var s = normStatus(st[i][0]);
        if (s === 'Pending') { if (sd[i][0] !== '') { sd[i][0] = ''; changed = true; } }
        else if (!sd[i][0]) { sd[i][0] = now; changed = true; }
      }
      if (changed) { rng.setNumberFormat('@'); rng.setValues(sd); }
    }
    numberEditedRows(tab, r1, r2, (r.getLastColumn() - r.getColumn() + 1) >= 8);
  } catch (err) { console.error('onOrderEdit: ' + (err && err.message)); }
}

/* ---------- seed data ---------- */
function seedRiders(log) {
  if (readTable('Riders').length) return;
  var names = ['Local', 'Ahmad Ullah', 'Hanif', 'Asim Mama', 'Molana'], first = reserveSeq('RIDER_SEQ', names.length);
  appendObjects('Riders', names.map(function (n, i) { return { RiderID: fmtId('R-', first + i, 4), Name: n, Phone: '', Note: 'Number Admin > Riders se likhein', Active: 'Yes' }; }));
  log.push('Riders tab mein 5 naam daal diye (phone numbers aap portal se likhein)');
}
var G_A = [['2500', 500], ['2100–2400', 400], ['Less than 2100', 300]], G_B = [['2100–2500', 400], ['Less than 2100', 300]], G_C = [['1600–1800', 300], ['Less than 1600', 200]],
    G_D = [['3000', 400], ['Less than 2900', 300]], G_E = [['2500', 400], ['Less than 2100', 300]], G_F = [['3500', 500], ['3000–3400', 400], ['Less than 3000', 300]];
var GUIDE = [
  ['Stretch Marks Cream', G_A], ['Hip Firming Cream', G_A], ['Breast Marks Cream', G_A], ['Sehat Mand Majoon', [['2000', 400], ['1900 & less', 300]]],
  ['Basha Weight Gain', [['2100–2500', 400], ['2000 & less', 300]]], ['Basha W.G Tablets', [['2100–2500', 400], ['2000 & less', 300]]], ['Slimming Tea', G_C],
  ['Plora Weight Loss Cap', [['2100–2500', 400], ['2000 & less', 300]]], ['Plora Weight Gain Cap', [['2100–2500', 400], ['2000 & less', 300]]], ['Unaani Booster W.G', [['2100–2500', 400], ['2000 & less', 300]]],
  ['Cartplus Capsule 30 Cap', G_C], ['Cartplus Oil 30 ml', G_C], ['Biotin Capsule 30 Cap', G_D], ['Hair Food Capsule 30 Cap', G_D], ['Glutathione Whitening Capsule', G_D],
  ['Stress Care 60 Cap', G_E], ['Plora Stomach 30 Cap', [['2000', 400], ['Less than 2000', 300]]], ['Expert Majoon', [['3500', 500], ['3100–3400', 400], ['Less than 3100', 300]]],
  ['Mini Majoon', G_A], ['Peak Plus / Timing Capsule', G_A], ['Tilla Sanda Oil', [['1600', 400], ['Less than 1600', 300]]], ['Zahir 2in1 Oil & Shampoo', G_B], ['Seven Effect Hair Oil', G_C],
  ['Romax Cream + Zahir Melasma', G_B], ['Zahir Zaffrani Cream', G_B], ['Seven Effect Cream + F.W', G_B], ['Derma Cream + F.W', G_B], ['Free Beauty Serum + F.W', G_B], ['Derma Gluta Serum + F.W', G_B],
  ['Face Care Kit (4 pcs)', [['6000', 500], ['5500–5900', 400], ['Less than 5500', 300]]], ['Intence Anti Aging Serum & F.W', G_B], ['Yardlie Serum', G_A], ['Teeth White Powder', G_C],
  ['Papaya Hair Color', [['3400–3800', 400], ['3000–3300', 300]]], ['Enzo Hair Serum', G_A], ['Sydonic Hair Remove Cream', G_F], ['Hair Inhibitor Lotion', G_F], ['Romax Hair Removal Cream', G_B]
];
function seedGuide(log) {
  if (readTable('CommissionGuide').length) return;
  var rows = [], n = 0; GUIDE.forEach(function (g) { n += g[1].length; });
  var first = reserveSeq('GUIDE_SEQ', n), i = 0;
  GUIDE.forEach(function (g) { g[1].forEach(function (s) { rows.push({ GuideID: fmtId('G-', first + i, 5), Product: g[0], PriceSlab: s[0], Commission: s[1], Sort: i }); i++; }); });
  appendObjects('CommissionGuide', rows);
  log.push('Commission guide (catalogue) ke ' + rows.length + ' rows daal diye (CommissionGuide tab mein edit kar sakte hain)');
}

/* ---------- admin account ---------- */
function initializeAdmin() {
  var uname = String(prop('INITIAL_ADMIN_USERNAME') || '').trim().toLowerCase(), pw = prop('INITIAL_ADMIN_PASSWORD'), name = prop('INITIAL_ADMIN_NAME') || 'Owner';
  if (!uname || !pw) throw new Error('Pehle Script properties mein INITIAL_ADMIN_USERNAME aur INITIAL_ADMIN_PASSWORD add karein.');
  if (!/^[a-z0-9._-]{3,30}$/.test(uname)) throw new Error('Username sirf chhote letters/numbers ho.');
  checkPasswordPolicy(pw);
  if (readTable('Users').some(function (u) { return u.Role === 'admin'; })) throw new Error('Admin pehle se maujood hai. Password bhool gaye hon to resetAdminPassword chalayein.');
  withLock(function () {
    appendObjects('Users', [{ UserID: fmtId('U-', reserveSeq('USER_SEQ', 1), 4), Username: uname, PasswordHash: hashPassword(pw), Role: 'admin', DisplayName: name, Active: 'Yes', CreatedAt: nowStr(), LastLogin: '', PasswordChangedAt: nowStr(), MustChangePassword: 'Yes' }]);
  });
  PropertiesService.getScriptProperties().deleteProperty('INITIAL_ADMIN_PASSWORD');
  audit({ user: { username: 'setup', role: 'system' } }, 'ADMIN_INITIALIZED', 'User', uname, '');
  Logger.log('Admin "' + uname + '" ban gaya. Password property delete ho gayi. Pehli dafa login par naya password rakhna hoga.');
}
/** Forgot the admin password: add RESET_USERNAME and RESET_PASSWORD in Script properties, then run this. */
function resetAdminPassword() {
  var uname = String(prop('RESET_USERNAME') || '').trim().toLowerCase(), pw = prop('RESET_PASSWORD');
  if (!uname || !pw) throw new Error('Pehle Script properties mein RESET_USERNAME aur RESET_PASSWORD add karein.');
  checkPasswordPolicy(pw);
  var u = findUserByName(uname); if (!u) throw new Error('Username "' + uname + '" nahi mila (Users tab dekhein).');
  withLock(function () { var x = getById('Users', u.UserID); x.PasswordHash = hashPassword(pw); x.PasswordChangedAt = nowStr(); x.MustChangePassword = 'Yes'; x.Active = 'Yes'; updateObj('Users', x); });
  clearUserCache(u.UserID); CacheService.getScriptCache().remove('F:' + sha256hex(uname)); revokeSessions(u.UserID, null);
  PropertiesService.getScriptProperties().deleteProperty('RESET_PASSWORD');
  audit({ user: { username: 'setup', role: 'system' } }, 'ADMIN_PASSWORD_RESET', 'User', u.UserID, 'via editor');
  Logger.log('Done. "' + uname + '" ka password naya ho gaya. Ab app mein login karein.');
}

function validateLite() {
  var issues = [], tabs = orderTabs();
  Object.keys(SCHEMA).forEach(function (name) {
    var s = ss().getSheetByName(name); if (!s) { issues.push('Tab nahi mila: ' + name); return; }
    var head = s.getRange(1, 1, 1, cols(name).length).getValues()[0];
    cols(name).forEach(function (c, i) { if (String(head[i]) !== c.name) issues.push(name + ': header ' + (i + 1) + ' "' + c.name + '" hona chahiye'); });
  });
  if (!tabs.length) issues.push('Koi order tab nahi mila');
  tabs.forEach(function (t) { NEW_COLS.forEach(function (nc) { if (!t.h.cols[nc[0]]) issues.push(t.name + ': "' + nc[1] + '" column nahi'); }); });
  readTable('PayLog').forEach(function (l) { if (l.Status === 'Processing') issues.push('PayLog ' + l.LogID + ' adhoori ("Processing") hai. Orders ke Saleman Payment columns check karein.'); });
  var users = readTable('Users'), names = {}; readAllOrders().forEach(function (o) { names[normName(o.salesman)] = 1; });
  users.forEach(function (u) { if (u.Role === 'salesman' && !names[normName(u.DisplayName)]) issues.push('User "' + u.Username + '": sheet mein "' + u.DisplayName + '" naam ka koi order nahi (spelling check karein)'); });
  Logger.log(issues.length ? issues.join('\n') : 'Koi masla nahi mila.');
  return issues;
}
function benchmarkPasswordHash() { var t = Date.now(); hashPassword('Benchmark123456'); Logger.log('Ek password hash ' + (Date.now() - t) + ' ms.'); }
function backupCopy() {
  var copy = ss().copy('ORDERS backup ' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HHmm'));
  Logger.log('Backup bana: ' + copy.getUrl());
}


/* =====================================================================
   FIX: "Koi order tab nahi mila"  ->  Apps Script mein yeh function Run karein: fixOrderTab
   Yeh aaj ke mahine ka tab (jaise NOVEMBER 2026) theek karta hai:
   - tab na ho to bana deta hai
   - tab ki pehli row khali ho to zaroori column ke naam likh deta hai
   - phir setupLite() chala deta hai (baqi helper columns khud lag jate hain)
   Agar pehli row mein pehle se kuch likha ho magar naam alag hon, to kuch nahi badalta;
   Execution log mein dikha deta hai ke kya mila, woh mujhe bhej dein.
   ===================================================================== */
function fixOrderTab() {
  var log = [], name = MONTHS_UP[Number(Utilities.formatDate(new Date(), CFG.TZ, 'M')) - 1] + ' ' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyy');
  var book = ss(), sheets = book.getSheets();
  log.push('Sheet: ' + book.getName());
  sheets.forEach(function (sh) {
    var lc = Math.min(Math.max(sh.getLastColumn(), 1), 30);
    log.push('Tab "' + sh.getName() + '" | row 1 = ' + JSON.stringify(sh.getRange(1, 1, 1, lc).getValues()[0]) + ' | order tab? ' + (headerMapOf(sh) ? 'HAAN' : 'NAHI'));
  });
  var sh = book.getSheetByName(name);
  if (!sh) {
    // koi tab mahine ke naam ka nahi: pehla khali tab (jaise Sheet1) ko rename karo, warna naya banao
    var empty = sheets.filter(function (s) { return s.getLastRow() <= 1 && s.getLastColumn() <= 1 && String(s.getRange(1, 1).getValue()).trim() === ''; })[0];
    if (empty && !sheets.some(function (s) { return TAB_RE.test(s.getName().trim()) && headerMapOf(s); })) { empty.setName(name); sh = empty; log.push('Khali tab ka naam badal kar "' + name + '" rakha.'); }
    else if (!orderTabs().length) { sh = book.insertSheet(name); log.push('Naya tab bana: ' + name); }
  }
  if (sh && !headerMapOf(sh)) {
    var lc2 = Math.max(sh.getLastColumn(), 1), r1 = sh.getRange(1, 1, 1, lc2).getValues()[0];
    var rowEmpty = r1.every(function (x) { return String(x).trim() === ''; });
    if (rowEmpty) {
      var heads = ['Date', 'Product', 'Price', 'Customer Name', 'Phone Number', 'Address', 'Saleman', 'Courier', 'Tracking', 'Status', 'Payment', 'Product Cost', 'Ads Cost', 'Delivery Charges', 'Saleman C', 'Profit'];
      sh.getRange(1, 1, 1, heads.length).setValues([heads]).setFontWeight('bold').setBackground('#d9ead3');
      sh.setFrozenRows(1);
      log.push('Pehli row mein column ke naam likh diye: ' + heads.join(', '));
    } else {
      log.push('DHYAN: tab "' + sh.getName() + '" ki pehli row khali nahi hai magar zaroori naam (Product, Price, Customer Name, Saleman, Status) nahi mile. Kuch nahi badla. Upar row 1 ka text mujhe bhej dein.');
    }
  }
  // helper tabs (Users, PayLog ...) ke header mein khali khane ho to sirf wahi bhar do (jo likha hai wo nahi chhoora jata)
  Object.keys(SCHEMA).forEach(function (tn) {
    var hs = book.getSheetByName(tn); if (!hs) return;
    var exp = SCHEMA[tn].split(',').map(function (x) { return x.split(':')[0].trim(); });
    var cur = hs.getRange(1, 1, 1, exp.length).getValues()[0], fixed = [];
    exp.forEach(function (e, i) { if (String(cur[i]).trim() === '') { hs.getRange(1, i + 1).setValue(e); fixed.push(e); } });
    if (fixed.length) log.push('Tab "' + tn + '" ke khali header bhar diye: ' + fixed.join(', '));
  });
  log.push('Order tabs ab: ' + (orderTabs().map(function (t) { return t.name; }).join(', ') || 'KOI NAHI'));
  if (orderTabs().length) { try { setupLite(); log.push('setupLite chal gaya.'); } catch (e) { log.push('setupLite error: ' + e.message); } }
  Logger.log(log.join('\n'));
  return log;
}
