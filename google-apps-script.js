// Cadeaux Johann & Elsa - serveur de synchronisation (Google Apps Script)
// À coller dans un nouveau projet sur script.google.com (la feuille de données se crée toute seule).
var SHEET_NAME = 'data';
var COLLS = ['people', 'gifts', 'budgets', 'shares', 'likes'];

function getSpreadsheet_() {
  // Fonctionne que le script soit lié à une feuille ou créé seul sur script.google.com
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create('Super Cadeau - données');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['coll', 'id', 'u', 'json']); }
  return sh;
}

function readAll_(sh) {
  var rows = sh.getDataRange().getValues();
  var out = {}, idx = {};
  COLLS.forEach(function (c) { out[c] = {}; });
  for (var i = 1; i < rows.length; i++) {
    var c = rows[i][0], id = rows[i][1], j = rows[i][3];
    if (!c || !id) continue;
    try { if (!out[c]) out[c] = {}; out[c][id] = JSON.parse(j); idx[c + '|' + id] = i + 1; } catch (e) {}
  }
  return { out: out, idx: idx };
}

function handle_(incoming) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    var all = readAll_(sh), out = all.out, idx = all.idx;
    if (incoming) {
      var upd = [], add = [];
      COLLS.forEach(function (c) {
        var m = incoming[c] || {};
        Object.keys(m).forEach(function (id) {
          var rec = m[id];
          if (!rec || typeof rec !== 'object') return;
          var cur = out[c][id];
          if (!cur || (rec.u || 0) > (cur.u || 0)) {
            out[c][id] = rec;
            var row = [c, id, rec.u || 0, JSON.stringify(rec)], k = c + '|' + id;
            if (idx[k]) upd.push([idx[k], row]); else add.push(row);
          }
        });
      });
      upd.forEach(function (x) { sh.getRange(x[0], 1, 1, 4).setValues([x[1]]); });
      if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setValues(add);
    }
    return { ok: true, data: out, t: Date.now() };
  } finally {
    lock.releaseLock();
  }
}

function reply_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try { return reply_(handle_(null)); } catch (err) { return reply_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return reply_(handle_(body && body.data ? body.data : null));
  } catch (err) { return reply_({ ok: false, error: String(err) }); }
}
