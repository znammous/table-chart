// خادم بسيط لجدول مسح الطاولة — يُلصق في Google Apps Script المرتبط بجدول Google Sheet
// الأعمدة: id | date | meal | name | createdAt | approved

var PARENT_PASSWORD = '12345';   // كلمة سر ولي الأمر — غيّرها من هنا متى شئت
var KEEP_DAYS = 45;              // التسجيلات الأقدم من هذا تُمسح تلقائياً
var NAMES = ['غنون', 'ملوك'];
var MEALS = ['فطور', 'غداء', 'مسح خفيف'];
var REPEATABLE = ['مسح خفيف'];  // يمكن تسجيله أكثر من مرة في اليوم
var SHEET_NAME = 'Entries';

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
  return json_({ ok: true, entries: list_(sheet) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
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
      return json_({ ok: true, entries: list_(sheet) });
    }

    if (body.action === 'approve' || body.action === 'delete') {
      if (body.password !== PARENT_PASSWORD) return json_({ ok: false, error: 'كلمة السر غير صحيحة' });
      var target = readAll_(sheet).filter(function (r) { return r.id === body.id; })[0];
      if (!target) return json_({ ok: false, error: 'التسجيل غير موجود' });
      if (body.action === 'approve') sheet.getRange(target.row, 6).setValue('TRUE');
      else sheet.deleteRow(target.row);
      return json_({ ok: true, entries: list_(sheet) });
    }

    return json_({ ok: false, error: 'طلب غير معروف' });
  } finally {
    lock.releaseLock();
  }
}
