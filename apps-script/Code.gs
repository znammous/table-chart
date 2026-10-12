// خادم جدول المهام المنزلية — يُلصق في Google Apps Script المرتبط بجدول Google Sheet
// ورقة Entries: id | date | meal (اسم المهمة) | name (اسم الطفل) | createdAt | approved
// ورقة History: سجل دائم للمهام الموافق عليها (لا يُمسح) — منه تُحسب النقاط والشارات وأبطال الشهر
// إعدادات العائلة (الأطفال والمهام وهدف الشهر) محفوظة في Script Properties تحت المفتاح config

var PARENT_PASSWORD = '12345';   // كلمة سر ولي الأمر — غيّرها من هنا متى شئت
var KEEP_DAYS = 45;              // التسجيلات الأقدم من هذا تُمسح تلقائياً
var SHEET_NAME = 'Entries';
var HISTORY_NAME = 'History';

var DEFAULT_CONFIG = {
  title: 'جدول الأبطال',
  goal: 40,
  kids: [
    { id: 'k1', name: 'غنون', gender: 'f', color: '#c2457d', icon: '🌸' },
    { id: 'k2', name: 'ملوك', gender: 'f', color: '#7a4fa3', icon: '🦋' }
  ],
  tasks: [
    { id: 't1', name: 'فطور', icon: '🍳', points: 2, repeat: false },
    { id: 't2', name: 'غداء', icon: '🍲', points: 3, repeat: false },
    { id: 't3', name: 'مسح خفيف', icon: '🧽', points: 1, repeat: true }
  ]
};

// ---------- الإعدادات ----------
function getConfig_() {
  var raw = PropertiesService.getScriptProperties().getProperty('config');
  var c = raw ? JSON.parse(raw) : DEFAULT_CONFIG;
  if (!c.title) c.title = DEFAULT_CONFIG.title;
  return c;
}

function cleanText_(s, max) { return String(s || '').replace(/\s+/g, ' ').trim().slice(0, max); }

// يتحقق من الإعدادات الجديدة ويعيد نسخة نظيفة، أو رسالة خطأ
function validateConfig_(c) {
  if (!c || !Array.isArray(c.kids) || !Array.isArray(c.tasks)) return 'بيانات غير صحيحة';
  var kids = c.kids.map(function (k) {
    return { id: cleanText_(k.id, 12), name: cleanText_(k.name, 20), gender: k.gender === 'm' ? 'm' : 'f',
             color: /^#[0-9a-f]{6}$/i.test(k.color) ? k.color : '#c2457d', icon: cleanText_(k.icon, 8) || '⭐' };
  });
  var tasks = c.tasks.map(function (t) {
    return { id: cleanText_(t.id, 12), name: cleanText_(t.name, 20), icon: cleanText_(t.icon, 8) || '✨',
             points: Math.max(1, Math.min(10, Math.round(+t.points) || 1)), repeat: !!t.repeat };
  });
  if (kids.length < 1 || kids.length > 8) return 'عدد الأطفال من 1 إلى 8';
  if (tasks.length < 1 || tasks.length > 10) return 'عدد المهام من 1 إلى 10';
  var all = kids.concat(tasks);
  if (all.some(function (x) { return !x.id || !x.name; })) return 'كل اسم يجب أن يكون مكتوباً';
  function unique(list) { var seen = {}; return list.every(function (x) { if (seen[x]) return false; seen[x] = true; return true; }); }
  if (!unique(kids.map(function (k) { return k.name; }))) return 'يوجد اسم طفل مكرر';
  if (!unique(tasks.map(function (t) { return t.name; }))) return 'يوجد اسم مهمة مكرر';
  if (!unique(all.map(function (x) { return x.id; }))) return 'بيانات غير صحيحة';
  return { title: cleanText_(c.title, 40) || 'جدول الأبطال', goal: Math.max(5, Math.min(500, Math.round(+c.goal) || 40)), kids: kids, tasks: tasks };
}

// عند تغيير اسم طفل أو مهمة يُحدَّث الاسم في كل التسجيلات القديمة حتى لا تضيع النقاط
function applyRenames_(oldList, newList, col, sheets) {
  var map = {};
  oldList.forEach(function (o) {
    var n = newList.filter(function (x) { return x.id === o.id; })[0];
    if (n && n.name !== o.name) map[o.name] = n.name;
  });
  if (!Object.keys(map).length) return;
  sheets.forEach(function (sh) {
    var last = sh.getLastRow();
    if (last < 2) return;
    var range = sh.getRange(2, col, last - 1, 1);
    var vals = range.getDisplayValues().map(function (r) { return [map[r[0]] || r[0]]; });
    range.setNumberFormat('@').setValues(vals);
  });
}

// ---------- الأوراق ----------
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

function prizes_() {
  var all = PropertiesService.getScriptProperties().getProperties();
  var out = {};
  Object.keys(all).forEach(function (k) { if (k.indexOf('prize_') === 0) out[k.slice(6)] = all[k]; });
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function response_(sheet, hist) {
  syncHistory_(sheet, hist);
  return json_({
    ok: true,
    config: getConfig_(),
    entries: readAll_(sheet).map(function (r) { return { id: r.id, date: r.date, meal: r.meal, name: r.name, approved: r.approved }; }),
    history: readHistory_(hist).map(function (h) { return { date: h.date, meal: h.meal, name: h.name }; }),
    prizes: prizes_()
  });
}

// ---------- الطلبات ----------
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
    var config = getConfig_();

    if (body.action === 'add') {
      var kid = config.kids.filter(function (k) { return k.name === body.name; })[0];
      var task = config.tasks.filter(function (t) { return t.name === body.meal; })[0];
      if (!kid || !task || !/^\d{4}-\d{2}-\d{2}$/.test(body.date)) {
        return json_({ ok: false, error: 'بيانات غير صحيحة' });
      }
      var taken = !task.repeat && readAll_(sheet).some(function (r) { return r.date === body.date && r.meal === body.meal; });
      if (taken) return json_({ ok: false, error: 'هذه المهمة مسجلة مسبقاً في هذا التاريخ' });
      var row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 6).setNumberFormat('@').setValues([[
        Utilities.getUuid(), body.date, body.meal, body.name, new Date().toISOString(), 'FALSE'
      ]]);
      return response_(sheet, hist);
    }

    if (body.password !== PARENT_PASSWORD) return json_({ ok: false, error: 'كلمة السر غير صحيحة' });

    if (body.action === 'approve' || body.action === 'delete') {
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
      if (!/^\d{4}-\d{2}$/.test(body.month)) return json_({ ok: false, error: 'بيانات غير صحيحة' });
      var text = String(body.text || '').slice(0, 100);
      var props = PropertiesService.getScriptProperties();
      if (text) props.setProperty('prize_' + body.month, text); else props.deleteProperty('prize_' + body.month);
      return response_(sheet, hist);
    }

    if (body.action === 'config') {
      var next = validateConfig_(body.config);
      if (typeof next === 'string') return json_({ ok: false, error: next });
      applyRenames_(config.kids, next.kids, 4, [sheet, hist]);
      applyRenames_(config.tasks, next.tasks, 3, [sheet, hist]);
      PropertiesService.getScriptProperties().setProperty('config', JSON.stringify(next));
      return response_(sheet, hist);
    }

    if (body.action === 'check') return json_({ ok: true });

    return json_({ ok: false, error: 'طلب غير معروف' });
  } finally {
    lock.releaseLock();
  }
}
