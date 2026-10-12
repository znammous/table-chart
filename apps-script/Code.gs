// خادم بسيط لجدول مسح الطاولة — يُلصق في Google Apps Script المرتبط بجدول Google Sheet
// الأعمدة: id | date | meal | name | createdAt | approved
// ورقة History: سجل دائم للمسحات الموافق عليها (لا يُمسح) — منه تُحسب النقاط والشارات وأبطال الشهر

var PARENT_PASSWORD = '12345';   // كلمة سر ولي الأمر — غيّرها من هنا متى شئت
var KEEP_DAYS = 45;              // التسجيلات الأقدم من هذا تُمسح تلقائياً
var NAMES = ['غنون', 'ملوك'];
var MEALS = ['فطور', 'غداء', 'مسح خفيف'];
var REPEATABLE = ['مسح خفيف'];  // يمكن تسجيله أكثر من مرة في اليوم
var SHEET_NAME = 'Entries';
var HISTORY_NAME = 'History';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'date', 'meal', 'name', 'createdAt', 'approved']);
    sheet.getRange('A:F').setNumberFormat('@');
  }
  return sheet;
}

function getHistory_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_NAME);
    sheet.appendRow(['id', 'date', 'meal', 'name']);
    sheet.getRange('A:D').setNumberFormat('@');
  }
  return sheet;
}

function readHistory_(hist) {
  var values = hist.getDataRange().getDisplayValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) rows.push({ id: values[i][0], date: values[i][1], meal: values[i][2], name: values[i][3], row: i + 1 });
  }
  return rows;
}

// ينسخ كل تسجيل موافق عليه إلى السجل الدائم إن لم يكن فيه
function syncHistory_(sheet, hist) {
  var known = {};
  readHistory_(hist).forEach(function (h) { known[h.id] = true; });
  var add = readAll_(sheet).filter(function (r) { return r.approved && !known[r.id]; })
    .map(function (r) { return [r.id, r.date, r.meal, r.name]; });
  if (add.length) hist.getRange(hist.getLastRow() + 1, 1, add.length, 4).setNumberFormat('@').setValues(add);
}

function prizes_() {
  var all = PropertiesService.getScriptProperties().getProperties();
  var out = {};
  Object.keys(all).forEach(function (k) { if (k.indexOf('prize_') === 0) out[k.slice(6)] = all[k]; });
  return out;
}

function response_(sheet, hist) {
  syncHistory_(sheet, hist);
  return json_({
    ok: true,
    entries: list_(sheet),
    history: readHistory_(hist).map(function (h) { return { date: h.date, meal: h.meal, name: h.name }; }),
    prizes: prizes_()
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readAll_(sheet) {
  var values = sheet.getDataRange().getDisplayValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var v = values[i];
    if (!v[0]) continue;
    rows.push({ id: v[0], date: v[1], meal: v[2], name: v[3], createdAt: v[4], approved: v[5] === 'TRUE', row: i + 1 });
  }
  return rows;
}

function cutoff_() {
  var d = new Date();
  d.setDate(d.getDate() - KEEP_DAYS);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// يمسح التسجيلات الأقدم من KEEP_DAYS يوماً
function cleanup_(sheet) {
  var limit = cutoff_();
  var rows = readAll_(sheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].date < limit) sheet.deleteRow(rows[i].row);
  }
}

function list_(sheet) {
  return readAll_(sheet).map(function (r) {
    return { id: r.id, date: r.date, meal: r.meal, name: r.name, approved: r.approved };
  });
}

function doGet() {
  var sheet = getSheet_();
  cleanup_(sheet);
  return response_(sheet, getHistory_());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var hist = getHistory_();
    cleanup_(sheet);
    var body = JSON.parse(e.postData.contents || '{}');

    if (body.action === 'add') {
      if (NAMES.indexOf(body.name) < 0 || MEALS.indexOf(body.meal) < 0 || !/^\d{4}-\d{2}-\d{2}$/.test(body.date)) {
        return json_({ ok: false, error: 'بيانات غير صحيحة' });
      }
      var taken = REPEATABLE.indexOf(body.meal) < 0 && readAll_(sheet).some(function (r) { return r.date === body.date && r.meal === body.meal; });
      if (taken) return json_({ ok: false, error: 'هذه الوجبة مسجلة مسبقاً في هذا التاريخ' });
      var row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 6).setNumberFormat('@').setValues([[
        Utilities.getUuid(), body.date, body.meal, body.name, new Date().toISOString(), 'FALSE'
      ]]);
      return response_(sheet, hist);
    }

    if (body.action === 'approve' || body.action === 'delete') {
      if (body.password !== PARENT_PASSWORD) return json_({ ok: false, error: 'كلمة السر غير صحيحة' });
      var target = readAll_(sheet).filter(function (r) { return r.id === body.id; })[0];
      if (!target) return json_({ ok: false, error: 'التسجيل غير موجود' });
      if (body.action === 'approve') {
        sheet.getRange(target.row, 6).setValue('TRUE');
      } else {
        sheet.deleteRow(target.row);
        var h = readHistory_(hist).filter(function (x) { return x.id === body.id; })[0];
        if (h) hist.deleteRow(h.row);
      }
      return response_(sheet, hist);
    }

    if (body.action === 'prize') {
      if (body.password !== PARENT_PASSWORD) return json_({ ok: false, error: 'كلمة السر غير صحيحة' });
      if (!/^\d{4}-\d{2}$/.test(body.month)) return json_({ ok: false, error: 'بيانات غير صحيحة' });
      var text = String(body.text || '').slice(0, 100);
      var props = PropertiesService.getScriptProperties();
      if (text) props.setProperty('prize_' + body.month, text); else props.deleteProperty('prize_' + body.month);
      return response_(sheet, hist);
    }

    return json_({ ok: false, error: 'طلب غير معروف' });
  } finally {
    lock.releaseLock();
  }
}
