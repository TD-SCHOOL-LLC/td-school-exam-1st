/* Admin parolu (Script Properties-də ADMIN_PASSWORD yoxdursa, bu istifadə olunur) */
var DEFAULT_ADMIN_PASSWORD = 'admin2026';

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---- Müəllimlər siyahısı (Script Properties-də saxlanılır) ---- */
function getTeachers_() {
  var s = PropertiesService.getScriptProperties().getProperty('TEACHERS');
  if (!s) return [];
  try { return JSON.parse(s); } catch (e) { return []; }
}
function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || DEFAULT_ADMIN_PASSWORD;
}
function adminOk_(pw) {
  var real = getAdminPassword_();
  return !!real && String(pw) === String(real);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);

  if (data.action === 'checkAdmin') {
    return json_({ ok: adminOk_(data.password) });
  }

  if (data.action === 'setTeachers') {
    if (!adminOk_(data.password)) return json_({ ok: false, error: 'Parol yanlışdır' });
    var list = (data.teachers || [])
      .map(function (x) { return String(x).trim().substring(0, 60); })
      .filter(function (x) { return x.length > 0; })
      .slice(0, 60);
    if (!list.length) return json_({ ok: false, error: 'Siyahı boş ola bilməz' });
    PropertiesService.getScriptProperties().setProperty('TEACHERS', JSON.stringify(list));
    return json_({ ok: true, teachers: list });
  }

  var sheet = getSheet_();

  if (data.action === 'delete') {
    var values = sheet.getDataRange().getValues();
    for (var i = 0; i < values.length; i++) {
      var rowId = String(values[i][0]) || ('row' + (i + 1));
      if (rowId === String(data.id)) { sheet.deleteRow(i + 1); break; }
    }
    return json_({ status: 'deleted' });
  }

  var id = Utilities.getUuid();
  var details = JSON.stringify(data.details || []);
  if (details.length > 45000) { details = '[]'; }
  sheet.appendRow([
    id,
    data.first || '', data.last || '', data.teacher || '', data.group || '',
    data.score !== undefined ? data.score : '',
    data.max !== undefined ? data.max : '',
    data.pct !== undefined ? data.pct : '',
    new Date().toISOString(),
    details
  ]);
  return json_({ status: 'ok', id: id });
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'teachers') {
    return json_({ teachers: getTeachers_() });
  }
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var start = (values.length > 0 && String(values[0][0]).toUpperCase() === 'ID') ? 1 : 0;
  var rows = [];
  for (var i = start; i < values.length; i++) {
    var row = values[i];
    var isEmpty = true;
    for (var c = 0; c < row.length; c++) { if (row[c] !== '' && row[c] !== null) { isEmpty = false; break; } }
    if (isEmpty) continue;
    var details = [];
    try { details = JSON.parse(row[9] || '[]'); } catch (err) { details = []; }
    var dateVal = row[8];
    if (dateVal instanceof Date) { dateVal = dateVal.toISOString(); }
    rows.push({
      id: String(row[0]) || ('row' + (i + 1)),
      first: row[1], last: row[2], teacher: row[3], group: row[4],
      score: row[5], max: row[6], pct: row[7], date: dateVal, details: details
    });
  }
  return json_({ results: rows });
}
