/**
 * שיווק — ניהול משלוחים | Backend (Google Apps Script)
 * ההוראות: פותחים גוגל שיטס חדש → תוספים (Extensions) → Apps Script →
 * מדביקים את הקובץ הזה → פריסה (Deploy) → New deployment → Web app →
 * Execute as: Me | Who has access: Anyone → Deploy → מעתיקים את כתובת ה-exec.
 */

function doGet(e) {
  return handle(e);
}

function doPost(e) {
  return handle(e);
}

function handle(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || '';
  var out;
  try {
    if (action === 'ping') out = { ok: true, app: 'shivuk' };
    else if (action === 'getTable') out = { ok: true, rows: readSheet(p.sheet) };
    else if (action === 'saveRow') out = saveRow(p.sheet, JSON.parse(p.data));
    else if (action === 'saveRows') out = saveRows(p.sheet, JSON.parse(p.data), p.clear === '1');
    else if (action === 'deleteRow') out = deleteRow(p.sheet, p.id);
    else out = { ok: false, error: 'unknown action' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// קורא את שורת הכותרות בפועל; מוסיף עמודות חסרות לפי המפתחות של הרשומה
function ensureHeaders(sh, record) {
  var lastCol = sh.getLastColumn();
  var headers = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  headers = headers.filter(function (h) { return h !== ''; });
  var keys = Object.keys(record || {});
  var added = false;
  keys.forEach(function (k) {
    if (headers.indexOf(k) === -1) { headers.push(k); added = true; }
  });
  if (headers.length === 0) headers = ['id'];
  if (added || lastCol === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}

function readSheet(name) {
  var sh = getSheet(name);
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(String);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {}, empty = true;
    for (var j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      var v = values[i][j];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      row[headers[j]] = v;
      if (v !== '' && v != null) empty = false;
    }
    if (!empty) rows.push(row);
  }
  return rows;
}

function saveRow(name, record) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet(name);
    var headers = ensureHeaders(sh, record);
    var idCol = headers.indexOf('id') + 1;
    var rowValues = headers.map(function (h) {
      return record.hasOwnProperty(h) ? record[h] : '';
    });
    var lastRow = sh.getLastRow();
    if (idCol > 0 && record.id != null && lastRow > 1) {
      var ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(record.id)) {
          sh.getRange(i + 2, 1, 1, headers.length).setValues([rowValues]);
          return { ok: true, updated: true };
        }
      }
    }
    sh.appendRow(rowValues);
    return { ok: true, created: true };
  } finally {
    lock.releaseLock();
  }
}

function saveRows(name, records, clear) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet(name);
    if (clear && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 1)).clearContent();
    }
    if (!records || !records.length) return { ok: true, count: 0 };
    var headers = ensureHeaders(sh, records[0]);
    var data = records.map(function (r) {
      return headers.map(function (h) { return r.hasOwnProperty(h) ? r[h] : ''; });
    });
    sh.getRange(sh.getLastRow() + 1, 1, data.length, headers.length).setValues(data);
    return { ok: true, count: data.length };
  } finally {
    lock.releaseLock();
  }
}

function deleteRow(name, id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet(name);
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2) return { ok: false, error: 'not found' };
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    var idCol = headers.indexOf('id') + 1;
    if (idCol < 1) return { ok: false, error: 'no id column' };
    var ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sh.deleteRow(i + 2);
        return { ok: true };
      }
    }
    return { ok: false, error: 'not found' };
  } finally {
    lock.releaseLock();
  }
}
