// ============================================================
// 10x Exam Artifact Builder System
// Version: 6.7 (Auto Kick + Absent Result Label + Auto-Resumable Delete-All + Scoring Fix)
// Author: 10x Programs Management Department
// ============================================================
// PREREQUISITE: Enable "Drive API" Advanced Service
//   → Services (+) → Drive API → Add
// ============================================================


// ─────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────

var MODULE_NAMES = [
  'პიროვნული უნარები და ინტერპერსონალური კომუნიკაციები',
  'მეწარმეობა',
  'უცხოური ენა',
  'ვერსიების კონტროლი Git ტექნოლოგიის გამოყენებით',
  'ვებ გვერდის მარკირება და სტილებით გაფორმება',
  'ვებ გვერდის სტილიზაცია Sass/scss პრე პროცესორის და Tailwind ფრეიმვორკის გამოყენებით',
  'ხელოვნური ინტელექტის გამოყენება',
  'ვებგვერდის ინტერაქტიულობისა და ეფექტების შექმნა JavaScript-ის საშუალებით',
  'ვებგვერდის ინტერაქტიულობისა და ეფექტების შემუშავება React-ის საშუალებით',
  'დინამიური ვებგვერდის  შექმნა typescript-ის საშუალებით',
  'ვებ გვერდის ოპტიმიზაცია',
  'დარგობრივი ინგლისური ენა Front-end დეველოპმენტისთვის',
  'პრაქტიკული პროექტი'
];

var CONFIG_SHEET_NAME = 'Config';
var AUDIT_SHEET_NAME = 'AuditLog';
var ACTIVE_STUDENTS_SHEET = 'ActiveStudents';
var DELETE_ALL_GENERATED_MARKER = '__ALL_GENERATED__';
var ABSENT_RESULT_LABEL = 'არ გამოცხადდა';

var CONFIG_KEYS = {
  WORKING_FOLDER_ID: 'WORKING_FOLDER_ID',
  EVAL_SHEETS_FOLDER_ID: 'EVAL_SHEETS_FOLDER_ID',
  INDIVIDUAL_EVAL_FOLDER_ID: 'INDIVIDUAL_EVAL_FOLDER_ID',
  TEMPLATES_FOLDER_ID: 'TEMPLATES_FOLDER_ID',
  EVAL_DOC_TEMPLATE_ID: 'EVAL_DOC_TEMPLATE_ID',
  EVAL_PAPER_TEMPLATE_ID: 'EVAL_PAPER_TEMPLATE_ID',
  DOC_VARS_SHEET_ID: 'DOC_VARS_SHEET_ID',
  PAPER_VARS_SHEET_ID: 'PAPER_VARS_SHEET_ID',
  STUDENT_FOLDER_IDS: 'STUDENT_FOLDER_IDS',
  LAST_BATCH_OPERATION: 'LAST_BATCH_OPERATION',
  LAST_BATCH_PROGRESS: 'LAST_BATCH_PROGRESS',
  LAST_BATCH_TOTAL: 'LAST_BATCH_TOTAL',
  LAST_BATCH_RANGE: 'LAST_BATCH_RANGE',
  LAST_BATCH_MODULES: 'LAST_BATCH_MODULES',
  AUTO_RUN_ACTIVE: 'AUTO_RUN_ACTIVE',
  DETAILED_LOGGING: 'DETAILED_LOGGING'
};


// ─────────────────────────────────────────────────────────────
// MENU
// ─────────────────────────────────────────────────────────────

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  var folderMenu = ui.createMenu('📁 ფოლდერები')
    .addItem('აირჩიე სამუშაო ფოლდერი', 'setWorkingFolder')
    .addSeparator()
    .addItem('შექმენი სრული სტრუქტურა', 'createFullStructure')
    .addItem('შექმენი სტუდენტების ფოლდერები', 'createStudentFolders')
    .addItem('შექმენი მოდულების ქვეფოლდერები', 'createModuleSubfolders');

  var docMenu = ui.createMenu('📄 დოკუმენტები')
    .addItem('შექმენი ჯგუფური ფურცელი', 'generateGroupSheet')
    .addItem('შექმენი ყველა ჯგუფური ფურცელი', 'generateAllBlankDocs')
    .addSeparator()
    .addItem('შექმენი ინდივიდუალური დოკუმენტები', 'generateIndividualBlankDocs');

  var batchMenu = ui.createMenu('📦 ბეჩ ოპერაციები')
    .addItem('ფოლდერების შექმნა (დიაპაზონით)', 'batchCreateFolders')
    .addSeparator()
    .addItem('დოკუმენტების გენერაცია (დიაპაზონი)', 'batchGenerateDocs')
    .addItem('დოკუმენტების წაშლა (დიაპაზონი)', 'batchDeleteDocs')
    .addItem('ყველა გენერირებული დოკუმენტის წაშლა', 'deleteAllGeneratedDocs')
    .addItem('წაშლის პროგრესი', 'showDeleteProgress')
    .addSeparator()
    .addItem('🌙 ავტო-გენერაცია (ღამის რეჟიმი)', 'startAutoGenerate')
    .addItem('🔁 გააღვიძე ავტო-გენერაცია', 'kickAutoGenerate')
    .addItem('⏹️ შეაჩერე ავტო-გენერაცია', 'stopAutoGenerate')
    .addSeparator()
    .addItem('▶ გააგრძელე შეწყვეტილი ოპერაცია', 'resumeLastOperation')
    .addItem('📊 ბეჩის პროგრესი', 'showBatchProgress');

  var settingsMenu = ui.createMenu('⚙️ პარამეტრები')
    .addItem('სამუშაო ფოლდერის ID', 'showWorkingFolderId')
    .addItem('სისტემის სტატუსი', 'showSystemStatus')
    .addSeparator()
    .addItem('ჩართე დეტალური ლოგი', 'enableDetailedLogging')
    .addItem('გამორთე დეტალური ლოგი', 'disableDetailedLogging')
    .addSeparator()
    .addItem('განაახლე ცვლადების ფაილები', 'refreshVariableSheets')
    .addItem('აუდიტის ლოგი', 'openAuditLog')
    .addItem('კონფიგურაციის გადატვირთვა', 'resetConfig');

  ui.createMenu('10x Exam System')
    .addSubMenu(folderMenu)
    .addSubMenu(docMenu)
    .addSubMenu(batchMenu)
    .addSubMenu(settingsMenu)
    .addToUi();
}


// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────

function getConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.getRange('A1').setValue('Key').setFontWeight('bold');
    sheet.getRange('B1').setValue('Value').setFontWeight('bold');
    sheet.getRange('C1').setValue('Description').setFontWeight('bold');
    var rows = [];
    var keys = Object.keys(CONFIG_KEYS);
    for (var i = 0; i < keys.length; i++) {
      var k = CONFIG_KEYS[keys[i]];
      var def = (k === 'STUDENT_FOLDER_IDS' || k === 'LAST_BATCH_MODULES') ? '{}' : (k === 'DETAILED_LOGGING') ? 'true' : '';
      rows.push([k, def, keys[i]]);
    }
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.setColumnWidth(1, 250); sheet.setColumnWidth(2, 400); sheet.setColumnWidth(3, 300);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:C1').setBackground('#1F4E79').setFontColor('#FFFFFF');
  }
  return sheet;
}

function getConfig(key) {
  var data = getConfigSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0] === key) return data[i][1]; }
  return null;
}

function setConfig(key, value) {
  var sheet = getConfigSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value, '']);
}

function _isDetailedLogging() {
  return String(getConfig(CONFIG_KEYS.DETAILED_LOGGING)) === 'true';
}


// ─────────────────────────────────────────────────────────────
// AUDIT LOG
// ─────────────────────────────────────────────────────────────

function getAuditLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Operation', 'Target', 'Result', 'Details', 'Duration (s)']]).setFontWeight('bold');
    sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 80); sheet.setColumnWidth(5, 500); sheet.setColumnWidth(6, 80);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:F1').setBackground('#1F4E79').setFontColor('#FFFFFF');
  }
  return sheet;
}

function logAudit(operation, target, result, details, durationSec) {
  var sheet = getAuditLogSheet();
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 6).setValues([[new Date(), operation, target || '', result || '', details || '', durationSec || '']]);
}

function logAuditUpdate(rowIndex, details, duration) {
  var sheet = getAuditLogSheet();
  sheet.getRange(rowIndex, 5).setValue(details);
  if (duration !== undefined) sheet.getRange(rowIndex, 6).setValue(duration);
}

function logAuditGetRow() {
  getAuditLogSheet().insertRowAfter(1);
  return 2;
}


// ─────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────

function getStudentList() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_STUDENTS_SHEET).getRange('A2:C').getValues();
  var students = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][1] && data[i][2]) {
      students.push({ name: String(data[i][0]).trim(), surname: String(data[i][1]).trim(),
        id: String(data[i][2]).replace(/\.0$/, '').trim(), row: i + 2 });
    }
  }
  return students;
}

function getStudentFolderName(s) { return s.name + ' ' + s.surname + ' - ' + s.id; }

function _compactModuleName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/[\s\/\\._\-–—()]+/g, '')
    .trim();
}

function _moduleCandidatesForHeader(header) {
  var h = _compactModuleName(header);
  if (!h) return [];

  var candidates = [];
  for (var i = 0; i < MODULE_NAMES.length; i++) {
    var m = _compactModuleName(MODULE_NAMES[i]);
    if (m === h || m.indexOf(h) === 0 || h.indexOf(m) === 0) {
      candidates.push(MODULE_NAMES[i]);
    }
  }
  return candidates;
}

function _matchModuleName(header, moduleName) {
  var h = _compactModuleName(header);
  var m = _compactModuleName(moduleName);
  if (!h || !m) return false;
  if (h === m) return true;

  var candidates = _moduleCandidatesForHeader(header);
  return candidates.length === 1 && _compactModuleName(candidates[0]) === m;
}

function _isConfirmedStatus(value) {
  var v = String(value || '').trim();
  return v === 'დადასტურდა' || v === 'ჩაეთვალა';
}

function _isDeniedStatus(value) {
  var v = String(value || '').trim();
  return v === 'არ დადასტურდა' || v === 'არ ჩაეთვალა';
}

function _isAbsentStatus(value) {
  var v = String(value || '').trim().toLowerCase();
  return v === '' || v === 'n/a' || v === 'na' || v === 'n.a.' || v === ABSENT_RESULT_LABEL;
}

function _paperResultMark(value) {
  if (_isConfirmedStatus(value)) return '✓';
  if (_isDeniedStatus(value)) return '✗';
  if (_isAbsentStatus(value)) return ABSENT_RESULT_LABEL;
  return '';
}

function _applyModuleStatusToOutcomeResults(outcomeNames, results, moduleStatus) {
  if (_isConfirmedStatus(moduleStatus)) {
    return outcomeNames.map(function() { return 'დადასტურდა'; });
  }
  if (_isDeniedStatus(moduleStatus) && results.length === 0) {
    return outcomeNames.map(function() { return 'არ დადასტურდა'; });
  }
  return results;
}

function _formatError(e) {
  if (!e) return 'Unknown error';
  return (e.message || String(e)) + (e.stack ? '\n\n' + e.stack : '');
}


// ─────────────────────────────────────────────────────────────
// PROMPTS
// ─────────────────────────────────────────────────────────────

function promptForRange() {
  var ui = SpreadsheetApp.getUi(), students = getStudentList();
  var r = ui.prompt('დიაპაზონი', 'მაგ: 1-30.\nსულ: ' + students.length, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  var m = r.getResponseText().trim().match(/^(\d+)\s*-\s*(\d+)$/);
  if (!m) { ui.alert('შეცდომა', 'ფორმატი: 1-30', ui.ButtonSet.OK); return null; }
  var s = parseInt(m[1]), e = parseInt(m[2]);
  if (s < 1 || e > students.length || s > e) { ui.alert('შეცდომა', '1-' + students.length, ui.ButtonSet.OK); return null; }
  return { start: s, end: e };
}

function promptForModule() {
  var ui = SpreadsheetApp.getUi();
  var list = ''; MODULE_NAMES.forEach(function(n, i) { list += (i+1) + '. ' + n + '\n'; });
  var r = ui.prompt('მოდული', '1-13:\n\n' + list, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  var n = parseInt(r.getResponseText().trim());
  if (isNaN(n) || n < 1 || n > 13) { ui.alert('შეცდომა', '1-13', ui.ButtonSet.OK); return null; }
  return MODULE_NAMES[n - 1];
}

function promptForModules() {
  var ui = SpreadsheetApp.getUi();
  var list = ''; MODULE_NAMES.forEach(function(n, i) { list += (i+1) + '. ' + n + '\n'; });
  var r = ui.prompt('მოდულები', 'ნომრები (მაგ: 4,5,6) ან "all":\n\n' + list, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  var text = r.getResponseText().trim().toLowerCase();
  if (text === 'all' || text === 'ყველა') return MODULE_NAMES.slice();
  var selected = [];
  text.split(/[,\s]+/).forEach(function(s) {
    var n = parseInt(s);
    if (!isNaN(n) && n >= 1 && n <= 13 && selected.indexOf(MODULE_NAMES[n-1]) === -1) selected.push(MODULE_NAMES[n-1]);
  });
  if (selected.length === 0) { ui.alert('შეცდომა', 'არცერთი არ არჩეულა.', ui.ButtonSet.OK); return null; }
  var cl = selected.map(function(n, i) { return (i+1) + '. ' + n; }).join('\n');
  if (ui.alert('მოდულები (' + selected.length + ')', cl + '\n\nთითო სტუდენტზე: ' + (selected.length*2) + ' დოკუმენტი.\nგნებავთ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return null;
  return selected;
}


// ─────────────────────────────────────────────────────────────
// FOLDER MANAGEMENT
// ─────────────────────────────────────────────────────────────

function setWorkingFolder() {
  var ui = SpreadsheetApp.getUi();
  var cur = getConfig(CONFIG_KEYS.WORKING_FOLDER_ID);
  var p = 'შეიყვანეთ 2026 ფოლდერის URL ან ID:' + (cur ? '\n\nმიმდინარე: ' + cur : '');
  var r = ui.prompt('სამუშაო ფოლდერი', p, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var input = r.getResponseText().trim(), fid = input;
  var um = input.match(/folders\/([a-zA-Z0-9_-]+)/); if (um) fid = um[1];
  var t = new Date();
  try {
    var folder = DriveApp.getFolderById(fid);
    setConfig(CONFIG_KEYS.WORKING_FOLDER_ID, fid);
    var evalId = '', indivId = '';
    var subs = folder.getFolders();
    while (subs.hasNext()) {
      var sub = subs.next(), nm = sub.getName();
      if (nm.indexOf('Evaluation Sheets') > -1 || nm.indexOf('შეფასების ფურცლები') > -1) evalId = sub.getId();
      if (nm.indexOf('Individual Evaluation') > -1 || nm.indexOf('ინდივიდუალური შეფასება') > -1) indivId = sub.getId();
    }
    setConfig(CONFIG_KEYS.EVAL_SHEETS_FOLDER_ID, evalId);
    setConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID, indivId);
    var tmpl = _discoverTemplates(fid);
    var msg = '✅ ' + folder.getName() + '\n\nEval Sheets: ' + (evalId ? '✅' : '❌') +
      '\nIndividual: ' + (indivId ? '✅' : '❌') + '\nTemplates: ' + (tmpl.found ? '✅' : '❌');
    if (tmpl.found) msg += '\n  Doc Template: ' + (tmpl.docTemplate ? '✅' : '❌') +
      '\n  Paper Template: ' + (tmpl.paperTemplate ? '✅' : '❌') +
      '\n  Doc Vars: ' + (tmpl.docVarsSheet ? '✅' : '❌') +
      '\n  Paper Vars: ' + (tmpl.paperVarsSheet ? '✅' : '❌');
    ui.alert('სამუშაო ფოლდერი', msg, ui.ButtonSet.OK);
    logAudit('setWorkingFolder', fid, 'success', 'Done', (new Date() - t) / 1000);
  } catch (e) {
    ui.alert('შეცდომა', e.message, ui.ButtonSet.OK);
    logAudit('setWorkingFolder', fid, 'error', e.message, (new Date() - t) / 1000);
  }
}

function _discoverTemplates(wid) {
  var folder = DriveApp.getFolderById(wid), tid = '';
  var subs = folder.getFolders();
  while (subs.hasNext()) { var s = subs.next(); if (s.getName().indexOf('Templates') > -1 || s.getName().indexOf('შაბლონები') > -1) { tid = s.getId(); break; } }
  if (!tid) return { found: false };
  setConfig(CONFIG_KEYS.TEMPLATES_FOLDER_ID, tid);
  var files = DriveApp.getFolderById(tid).getFiles();
  var r = { found: true, docTemplate: '', paperTemplate: '', docVarsSheet: '', paperVarsSheet: '' };
  while (files.hasNext()) {
    var f = files.next(), fn = f.getName().toLowerCase();
    if (fn.indexOf('evaluation_document_template') > -1 || fn.indexOf('evaluation-document-template') > -1) { r.docTemplate = f.getId(); setConfig(CONFIG_KEYS.EVAL_DOC_TEMPLATE_ID, f.getId()); }
    else if (fn.indexOf('eval-paper-template') > -1 || fn.indexOf('evaluation-paper-template') > -1) { r.paperTemplate = f.getId(); setConfig(CONFIG_KEYS.EVAL_PAPER_TEMPLATE_ID, f.getId()); }
    else if (fn.indexOf('evaluation-document-variables') > -1) { r.docVarsSheet = f.getId(); setConfig(CONFIG_KEYS.DOC_VARS_SHEET_ID, f.getId()); }
    else if (fn.indexOf('evaluation-paper-variables') > -1) { r.paperVarsSheet = f.getId(); setConfig(CONFIG_KEYS.PAPER_VARS_SHEET_ID, f.getId()); }
  }
  return r;
}


// ─────────────────────────────────────────────────────────────
// VARIABLE READERS
// ─────────────────────────────────────────────────────────────

function _loadAllDocVariables() {
  var id = getConfig(CONFIG_KEYS.DOC_VARS_SHEET_ID);
  if (!id) throw new Error('Doc variables not found.');
  return SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
}

function _loadAllPaperVariables() {
  var id = getConfig(CONFIG_KEYS.PAPER_VARS_SHEET_ID);
  if (!id) throw new Error('Paper variables not found.');
  return SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
}

function _readDocVarsFromCache(data, mod) {
  var col = -1;
  for (var c = 2; c < data[0].length; c++) { if (_matchModuleName(data[0][c], mod)) { col = c; break; } }
  if (col < 0) return {};
  var vars = {};
  for (var r = 1; r < data.length; r++) { var v = String(data[r][1] || '').trim(); if (v && v.indexOf('{{') === 0) vars[v] = data[r][col] != null ? String(data[r][col]) : ''; }
  return vars;
}

function _readPaperVarsFromCache(data, mod) {
  var col = -1;
  for (var c = 2; c < data[1].length; c++) { if (_matchModuleName(data[1][c], mod)) { col = c; break; } }
  if (col < 0) return {};
  var vars = {};
  for (var r = 2; r < data.length; r++) { var v = String(data[r][1] || '').trim(); if (v && v.indexOf('{{') === 0) vars[v] = data[r][col] != null ? String(data[r][col]) : ''; }
  return vars;
}

function _readPaperVariables(mod) { return _readPaperVarsFromCache(_loadAllPaperVariables(), mod); }


// ─────────────────────────────────────────────────────────────
// STUDENT OUTCOMES
// ─────────────────────────────────────────────────────────────

function _getStudentOutcomes(moduleName, studentId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), moduleSheet = null;
  var skip = ['ActiveStudents', 'examResultsRAW', 'Config', 'AuditLog', 'roles'];
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sn = sheets[s].getName();
    if (skip.indexOf(sn) === -1 && _matchModuleName(sn, moduleName)) { moduleSheet = sheets[s]; break; }
  }
  if (!moduleSheet) return { outcomeNames: [], results: [], moduleStatus: '' };

  var data = moduleSheet.getDataRange().getValues();
  var names = [], lastCol = 2;
  for (var c = 3; c < data[2].length; c++) {
    var val = String(data[2][c] || '').trim();
    if (val === '' || val.indexOf('ჩათვლა') > -1 || val.indexOf('მოდულის') > -1 || val.indexOf('დადასტურება') > -1) break;
    names.push(val); lastCol = c;
  }

  var row = -1, cid = String(studentId).replace(/\.0$/, '').trim();
  for (var r = 3; r < data.length; r++) {
    if (String(data[r][2] || '').replace(/\.0$/, '').trim() === cid) { row = r; break; }
  }
  if (row < 0) return { outcomeNames: names, results: names.map(function() { return 'N/A'; }), moduleStatus: '' };

  var results = [];
  for (var rc = 3; rc <= lastCol; rc++) results.push(String(data[row][rc] || 'N/A').trim());

  var moduleStatusCol = lastCol + 1;
  var moduleStatus = moduleStatusCol < data[row].length ? String(data[row][moduleStatusCol] || '').trim() : '';
  results = _applyModuleStatusToOutcomeResults(names, results, moduleStatus);

  return { outcomeNames: names, results: results, moduleStatus: moduleStatus };
}


// ─────────────────────────────────────────────────────────────
// FILE GENERATORS
// ─────────────────────────────────────────────────────────────

function _isGeneratedDocBlankOutcomeRow(row) {
  if (row.getNumCells() < 2) return false;
  var first = row.getCell(0).getText().trim();
  if (first.indexOf('შეფასების შედეგები') === -1) return false;
  for (var c = 1; c < row.getNumCells(); c++) {
    if (row.getCell(c).getText().trim() !== '') return false;
  }
  return true;
}

function _generateEvalDocWithVars(student, moduleName, folderId, vars) {
  var tid = getConfig(CONFIG_KEYS.EVAL_DOC_TEMPLATE_ID);
  var fileName = 'შეფასების ინსტრუმენტი - ' + moduleName;
  var folder = DriveApp.getFolderById(folderId);
  var ex = folder.getFilesByName(fileName); if (ex.hasNext()) return ex.next().getId();
  var copy = DriveApp.getFileById(tid).makeCopy(fileName, folder);
  var doc = DocumentApp.openById(copy.getId()), body = doc.getBody();
  var keys = Object.keys(vars);
  for (var v = 0; v < keys.length; v++) body.replaceText(keys[v].replace(/[\{\}]/g, '\\$&'), vars[keys[v]] || '');
  var tables = body.getTables();
  for (var t = tables.length - 1; t >= 0; t--) {
    var table = tables[t];
    for (var r = table.getNumRows() - 1; r > 0; r--) {
      var row = table.getRow(r), empty = true;
      for (var c = 0; c < row.getNumCells(); c++) { if (row.getCell(c).getText().trim() !== '') { empty = false; break; } }
      if (empty || _isGeneratedDocBlankOutcomeRow(row)) table.removeRow(r);
    }
  }
  doc.saveAndClose();
  return copy.getId();
}

function _generateEvalPaperWithVars(student, moduleName, folderId, vars) {
  var tid = getConfig(CONFIG_KEYS.EVAL_PAPER_TEMPLATE_ID);
  var fileName = 'შეფასების ფურცელი - ' + moduleName;
  var folder = DriveApp.getFolderById(folderId);
  var ex = folder.getFilesByName(fileName); if (ex.hasNext()) return ex.next().getId();
  var copy = DriveApp.getFileById(tid).makeCopy(fileName, folder);
  var ss = SpreadsheetApp.openById(copy.getId()), sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues(), nr = data.length, nc = data[0].length;
  var keys = Object.keys(vars);
  for (var r = 0; r < nr; r++) {
    for (var c = 0; c < nc; c++) {
      var cell = String(data[r][c] || '');
      if (cell.indexOf('{{') > -1) { for (var v = 0; v < keys.length; v++) { if (cell.indexOf(keys[v]) > -1) cell = cell.replace(keys[v], vars[keys[v]] || ''); } data[r][c] = cell; }
    }
  }
  sheet.getRange(1, 1, nr, nc).setValues(data);
  var upd = sheet.getDataRange().getValues();
  for (var rr = upd.length - 1; rr >= 5; rr--) {
    var a = String(upd[rr][0] || '').trim(), b = String(upd[rr][1] || '').trim();
    if (a.indexOf('პროფესიული მასწავლებელი') > -1 || a.indexOf('ვალიდატორი') > -1) continue;
    if (a === '' && b === '') sheet.deleteRow(rr + 1);
  }
  SpreadsheetApp.flush();
  return copy.getId();
}


// ─────────────────────────────────────────────────────────────
// BUILD VARS FOR ONE STUDENT×MODULE
// ─────────────────────────────────────────────────────────────

function _buildDocVars(student, moduleName, cachedDocVars, outcomes) {
  var vars = _readDocVarsFromCache(cachedDocVars, moduleName);
  vars['{{STUDENT_NAME}}'] = student.name + ' ' + student.surname;
  for (var oi = 0; oi < 15; oi++) {
    var num = oi + 1;
    if (oi < outcomes.results.length) {
      var result = outcomes.results[oi];
      vars['{{OUTCOME_CONFIRMED_' + num + '}}'] = _isConfirmedStatus(result) ? 'დადასტურდა' : '';
      vars['{{OUTCOME_DENIED_' + num + '}}'] = _isDeniedStatus(result) ? 'არ დადასტურდა' : (_isAbsentStatus(result) ? ABSENT_RESULT_LABEL : '');
    } else { vars['{{OUTCOME_CONFIRMED_' + num + '}}'] = ''; vars['{{OUTCOME_DENIED_' + num + '}}'] = ''; }
  }
  return vars;
}

function _buildPaperVars(student, moduleName, cachedPaperVars, outcomes) {
  var vars = _readPaperVarsFromCache(cachedPaperVars, moduleName);
  vars['{{STUDENT_NAME}}'] = student.name + ' ' + student.surname;
  if (!vars['{{MODULE_NAME}}']) vars['{{MODULE_NAME}}'] = moduleName;
  var map = {};
  for (var i = 0; i < outcomes.results.length; i++) {
    map[i + 1] = _paperResultMark(outcomes.results[i]);
  }
  for (var n = 1; n <= 13; n++) {
    var oName = vars['{{OUTCOME_NAME_' + n + '}}'] || '';
    vars['{{RESULT_' + n + '}}'] = oName ? (map[n] || '') : '';
    for (var m = 1; m <= 10; m++) {
      var cName = vars['{{CRITERION_' + n + '_' + m + '}}'] || '';
      vars['{{RESULT_' + n + '_' + m + '}}'] = (cName && oName) ? (map[n] || '') : '';
    }
  }
  return vars;
}


// ─────────────────────────────────────────────────────────────
// CORE GENERATION (shared by manual + auto)
// ─────────────────────────────────────────────────────────────

function _runGeneration(startIdx, endIdx, modules, showUI) {
  var startTime = new Date(), MAX = (showUI ? 4.5 : 3.5) * 60 * 1000;
  var detailed = _isDetailedLogging();

  var docTid = getConfig(CONFIG_KEYS.EVAL_DOC_TEMPLATE_ID);
  var paperTid = getConfig(CONFIG_KEYS.EVAL_PAPER_TEMPLATE_ID);
  if (!docTid || !paperTid || !getConfig(CONFIG_KEYS.DOC_VARS_SHEET_ID) || !getConfig(CONFIG_KEYS.PAPER_VARS_SHEET_ID)) {
    if (showUI) SpreadsheetApp.getUi().alert('შეცდომა', 'Templates არ არის კონფიგურირებული.', SpreadsheetApp.getUi().ButtonSet.OK);
    else logAudit('generate:error', '', 'error', 'Templates not configured');
    return;
  }

  var mods = modules || MODULE_NAMES;
  var docsPerStudent = mods.length * 2;

  logAudit('generate', startIdx + '-' + endIdx, 'running', 'Modules: ' + mods.length + '. Loading vars...', 0);
  var masterRow = 2;

  var cachedDocVars = _loadAllDocVariables();
  var cachedPaperVars = _loadAllPaperVariables();
  logAuditUpdate(masterRow, 'Vars loaded. Modules: ' + mods.length + '. Starting...', (new Date() - startTime) / 1000);

  var students = getStudentList();
  var ids = JSON.parse(getConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS) || '{}');
  var origRange = String(getConfig(CONFIG_KEYS.LAST_BATCH_RANGE) || '').replace(/^'/, '');
  var origStart = parseInt(origRange.split('-')[0]) || startIdx;

  var created = 0, processed = 0, skipped = 0, errors = 0, errList = [];

  for (var i = startIdx - 1; i < endIdx && i < students.length; i++) {
    if (new Date() - startTime > MAX) {
      setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, (i + 1) - origStart);
      getAuditLogSheet().getRange(masterRow, 4).setValue('timeout');
      logAuditUpdate(masterRow, 'TIMEOUT #' + (i+1) + '. Students: ' + processed + '. Docs: ' + created + '. Errors: ' + errors +
        (errList.length > 0 ? '. ERR: ' + errList.join(' | ') : ''), (new Date() - startTime) / 1000);
      if (showUI) SpreadsheetApp.getUi().alert('⏸️', 'სტუდენტები: ' + processed + '\nდოკუმენტები: ' + created + '\nშეჩერდა: ' + students[i].surname +
        '\n\nგააგრძელე: ბეჩ ოპერაციები → გააგრძელე', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var student = students[i], sfid = ids[student.id], label = student.surname;
    if (!sfid) { skipped++; processed++; continue; }

    var sf;
    try { sf = DriveApp.getFolderById(sfid); } catch (e) { errors++; processed++; errList.push(label + ':folder'); continue; }

    var sRow = -1;
    if (detailed) {
      sRow = logAuditGetRow(); masterRow++;
      var as = getAuditLogSheet();
      as.getRange(sRow, 1).setValue(new Date());
      as.getRange(sRow, 2).setValue('#' + (i+1));
      as.getRange(sRow, 3).setValue(label);
      as.getRange(sRow, 4).setValue('running');
    }

    var sDocs = 0, sErrs = 0, prog = [];

    for (var m = 0; m < mods.length; m++) {
      var mod = mods[m], modShort = mod.substring(0, 20);
      try {
        var mfi = sf.getFoldersByName(mod);
        if (!mfi.hasNext()) { prog.push(modShort + ':NO_DIR'); continue; }
        var mfid = mfi.next().getId();
        var outcomes = _getStudentOutcomes(mod, student.id);

        try { _generateEvalDocWithVars(student, mod, mfid, _buildDocVars(student, mod, cachedDocVars, outcomes)); created++; sDocs++; }
        catch (de) { errors++; sErrs++; errList.push(label + '/' + modShort + ':doc'); }

        try { _generateEvalPaperWithVars(student, mod, mfid, _buildPaperVars(student, mod, cachedPaperVars, outcomes)); created++; sDocs++; }
        catch (pe) { errors++; sErrs++; errList.push(label + '/' + modShort + ':paper'); }

        prog.push(modShort + ':✅');
      } catch (e) { errors++; sErrs++; prog.push(modShort + ':❌'); errList.push(label + '/' + modShort); }

      Utilities.sleep(1000);

      if (detailed && sRow > -1 && ((m+1) % 3 === 0 || m === mods.length - 1)) {
        logAuditUpdate(sRow, (m+1) + '/' + mods.length + '. Docs: ' + sDocs + '. ' + prog.join(', '), (new Date() - startTime) / 1000);
      }
    }

    if (detailed && sRow > -1) {
      getAuditLogSheet().getRange(sRow, 4).setValue(sErrs > 0 ? 'partial' : 'done');
      logAuditUpdate(sRow, 'Docs: ' + sDocs + '/' + docsPerStudent + '. ' + prog.join(', '), (new Date() - startTime) / 1000);
    }

    processed++;
    setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, (i + 1) - origStart + 1);

    if (detailed || processed % 10 === 0) {
      logAuditUpdate(masterRow, 'Students: ' + processed + '/' + (endIdx - startIdx + 1) + '. Docs: ' + created + '. Err: ' + errors,
        (new Date() - startTime) / 1000);
    }
  }

  setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, endIdx - origStart + 1);
  var dur = (new Date() - startTime) / 1000;
  getAuditLogSheet().getRange(masterRow, 4).setValue(errors > 0 ? 'partial' : 'success');
  logAuditUpdate(masterRow, 'DONE. Students: ' + processed + '. Docs: ' + created + '. Skip: ' + skipped + '. Err: ' + errors +
    (errList.length > 0 ? '. ' + errList.slice(0, 10).join(' | ') : ''), dur);

  if (showUI) {
    var info = '';
    if ((endIdx - startIdx + 1) <= 10) { for (var j = startIdx-1; j < endIdx && j < students.length; j++) info += '\n• ' + students[j].name + ' ' + students[j].surname; }
    var msg = 'სტუდენტები: ' + processed + '\nდოკუმენტები: ' + created + '\nმოდულები: ' + mods.length +
      (skipped > 0 ? '\nგამოიტოვა: ' + skipped : '') + (errors > 0 ? '\nშეცდომები: ' + errors : '') +
      info + '\n\nდრო: ' + dur.toFixed(1) + ' წამი';
    SpreadsheetApp.getUi().alert('✅', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// ─────────────────────────────────────────────────────────────
// MANUAL GENERATION
// ─────────────────────────────────────────────────────────────

function _generateDocsForRange(startIdx, endIdx, mods) {
  setConfig(CONFIG_KEYS.LAST_BATCH_OPERATION, 'generateDocs');
  setConfig(CONFIG_KEYS.LAST_BATCH_TOTAL, endIdx - startIdx + 1);
  setConfig(CONFIG_KEYS.LAST_BATCH_RANGE, "'" + startIdx + '-' + endIdx);
  setConfig(CONFIG_KEYS.LAST_BATCH_MODULES, JSON.stringify(mods || MODULE_NAMES));
  setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, 0);
  _runGeneration(startIdx, endIdx, mods, true);
}

function generateIndividualBlankDocs() {
  var ui = SpreadsheetApp.getUi();
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { ui.alert('შეცდომა', 'ფოლდერი არ მოიძებნა.', ui.ButtonSet.OK); return; }
  var mods = promptForModules(); if (!mods) return;
  var students = getStudentList();
  if (ui.alert('დაწყება', students.length + ' × ' + mods.length + ' × 2 = ' + (students.length*mods.length*2) + '\nგნებავთ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _generateDocsForRange(1, students.length, mods);
}

function batchGenerateDocs() {
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { SpreadsheetApp.getUi().alert('შეცდომა', 'ფოლდერი არ მოიძებნა.', SpreadsheetApp.getUi().ButtonSet.OK); return; }
  var mods = promptForModules(); if (!mods) return;
  var range = promptForRange(); if (!range) return;
  _generateDocsForRange(range.start, range.end, mods);
}


// ─────────────────────────────────────────────────────────────
// AUTO-GENERATE (OVERNIGHT)
// ─────────────────────────────────────────────────────────────

function startAutoGenerate() {
  var ui = SpreadsheetApp.getUi();
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { ui.alert('შეცდომა', 'ფოლდერი არ მოიძებნა.', ui.ButtonSet.OK); return; }
  var mods = promptForModules(); if (!mods) return;
  var range = promptForRange(); if (!range) return;

  var total = range.end - range.start + 1;
  var est = Math.ceil(total / 12) * 7;
  if (ui.alert('🌙 ავტო-გენერაცია',
    'სტუდენტები: ' + range.start + '-' + range.end + ' (' + total + ')\n' +
    'მოდულები: ' + mods.length + '\nდოკუმენტები: ~' + (total * mods.length * 2) + '\n' +
    'სავარაუდო: ~' + est + ' წუთი\n\nშეგიძლიათ დახუროთ ბრაუზერი.\nდაიწყოს?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  setConfig(CONFIG_KEYS.LAST_BATCH_OPERATION, 'generateDocs');
  setConfig(CONFIG_KEYS.LAST_BATCH_TOTAL, total);
  setConfig(CONFIG_KEYS.LAST_BATCH_RANGE, "'" + range.start + '-' + range.end);
  setConfig(CONFIG_KEYS.LAST_BATCH_MODULES, JSON.stringify(mods));
  setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, 0);
  setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, 'true');

  logAudit('auto:start', range.start + '-' + range.end, 'started', 'Modules: ' + mods.length + '. Total: ' + total, 0);
  _autoGenerateBatch();
}

function _autoGenerateBatch() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    _scheduleAutoGenerate(2 * 60 * 1000);
    logAudit('auto:locked', '', 'waiting', 'Another auto-generation run is still active. Next retry in 2 min.', 0);
    return;
  }

  try {
    if (String(getConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE)) !== 'true') return;
    var progress = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS) || '0');
    var total = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL) || '0');
    var range = String(getConfig(CONFIG_KEYS.LAST_BATCH_RANGE) || '').replace(/^'/, '');

    if (progress >= total || !range) {
      setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
      _clearAutoTriggers();
      logAudit('auto:complete', range, 'success', 'All done. ' + progress + '/' + total, 0);
      return;
    }

    var parts = range.split('-'), start = parseInt(parts[0]), end = parseInt(parts[1]);
    var resumeFrom = start + progress;
    var mods = null;
    try { var j = getConfig(CONFIG_KEYS.LAST_BATCH_MODULES); if (j) mods = JSON.parse(j); } catch (e) {}

    _runGeneration(resumeFrom, end, mods, false);

    var newProgress = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS) || '0');
    if (newProgress >= total) {
      setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
      _clearAutoTriggers();
      logAudit('auto:complete', range, 'success', 'Finished. ' + newProgress + '/' + total, 0);
    } else {
      _scheduleAutoGenerate(90 * 1000);
      logAudit('auto:scheduled', range, 'waiting', 'Progress: ' + newProgress + '/' + total + '. Next in ~90 sec.', 0);
    }
  } catch (e) {
    _scheduleAutoGenerate(3 * 60 * 1000);
    logAudit('auto:error', '', 'error', _formatError(e) + '\n\nRetry scheduled in 3 min.', 0);
  } finally {
    lock.releaseLock();
  }
}

function kickAutoGenerate() {
  var ui = SpreadsheetApp.getUi();
  var op = getConfig(CONFIG_KEYS.LAST_BATCH_OPERATION);
  var progress = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS) || '0');
  var total = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL) || '0');
  var range = String(getConfig(CONFIG_KEYS.LAST_BATCH_RANGE) || '').replace(/^'/, '');

  if (op !== 'generateDocs' || !range || !total) {
    ui.alert('ინფო', 'გასაღვიძებელი ავტო-გენერაცია ვერ მოიძებნა.', ui.ButtonSet.OK);
    return;
  }
  if (progress >= total) {
    setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
    _clearAutoTriggers();
    ui.alert('✅', 'გენერაცია უკვე დასრულებულია: ' + progress + '/' + total, ui.ButtonSet.OK);
    return;
  }

  setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, 'true');
  _clearAutoTriggers();
  logAudit('auto:kick', range, 'running', 'Manual wake. Progress: ' + progress + '/' + total, 0);
  ui.alert('🔁', 'ავტო-გენერაცია გაეშვება ახლავე.\nპროგრესი: ' + progress + '/' + total + '\nშემდეგი სტუდენტი: ' + (parseInt(range.split('-')[0]) + progress), ui.ButtonSet.OK);
  _autoGenerateBatch();
}

function _scheduleAutoGenerate(delayMs) {
  _clearAutoTriggers();
  ScriptApp.newTrigger('_autoGenerateBatch').timeBased().after(delayMs || 90 * 1000).create();
}

function _countAutoTriggers() {
  var count = 0;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === '_autoGenerateBatch') count++;
  }
  return count;
}

function _clearAutoTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) { if (triggers[i].getHandlerFunction() === '_autoGenerateBatch') ScriptApp.deleteTrigger(triggers[i]); }
}

function stopAutoGenerate() {
  setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
  _clearAutoTriggers();
  _clearDeleteTriggers();
  logAudit('auto:stopped', '', 'stopped', 'Manual stop');
  SpreadsheetApp.getUi().alert('⏹️', 'ავტო-ოპერაცია შეჩერდა.\nპროგრესი შენახულია.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ─────────────────────────────────────────────────────────────
// BULK DELETE
// ─────────────────────────────────────────────────────────────

function _promptForDeleteModules() {
  var ui = SpreadsheetApp.getUi();
  var list = ''; MODULE_NAMES.forEach(function(n, i) { list += (i+1) + '. ' + n + '\n'; });
  var mr = ui.prompt('წასაშლელი მოდულები', 'ნომრები (მაგ: 4,5,6) ან "all":\n\n' + list, ui.ButtonSet.OK_CANCEL);
  if (mr.getSelectedButton() !== ui.Button.OK) return null;
  var text = mr.getResponseText().trim().toLowerCase(), mods;
  if (text === 'all' || text === 'ყველა') { mods = MODULE_NAMES.slice(); }
  else { mods = []; text.split(/[,\s]+/).forEach(function(s) { var n = parseInt(s); if (!isNaN(n) && n >= 1 && n <= 13) mods.push(MODULE_NAMES[n-1]); }); }
  if (mods.length === 0) { ui.alert('შეცდომა', 'არცერთი.', ui.ButtonSet.OK); return null; }
  return mods;
}

function batchDeleteDocs() {
  var ui = SpreadsheetApp.getUi();
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { ui.alert('შეცდომა', 'ფოლდერი არ მოიძებნა.', ui.ButtonSet.OK); return; }
  var mods = _promptForDeleteModules(); if (!mods) return;
  var range = promptForRange(); if (!range) return;
  if (ui.alert('⚠️ წაშლა',
    (range.end - range.start + 1) + ' სტუდენტი × ' + mods.length + ' მოდული\n' +
    'წაიშლება მხოლოდ სისტემის გენერირებული ფაილები:\n' +
    '• შეფასების ინსტრუმენტი - [მოდული]\n' +
    '• შეფასების ფურცელი - [მოდული]\n\n' +
    'ფოლდერები და სხვა ფაილები რჩება.\nგნებავთ?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  setConfig(CONFIG_KEYS.LAST_BATCH_OPERATION, 'deleteDocs');
  setConfig(CONFIG_KEYS.LAST_BATCH_TOTAL, range.end - range.start + 1);
  setConfig(CONFIG_KEYS.LAST_BATCH_RANGE, "'" + range.start + '-' + range.end);
  setConfig(CONFIG_KEYS.LAST_BATCH_MODULES, JSON.stringify(mods));
  setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, 0);
  setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, 'deleteDocs');
  _deleteDocsForRange(range.start, range.end, mods, true);
}

function deleteAllGeneratedDocs() {
  try {
    _deleteAllGeneratedDocs();
  } catch (e) {
    logAudit('deleteAll:error', '', 'error', _formatError(e).substring(0, 45000), 0);
    SpreadsheetApp.getUi().alert('წაშლის შეცდომა', _formatError(e).substring(0, 1800), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function _deleteAllGeneratedDocs() {
  var ui = SpreadsheetApp.getUi();
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { ui.alert('შეცდომა', 'ფოლდერი არ მოიძებნა.', ui.ButtonSet.OK); return; }
  var students = getStudentList();
  if (ui.alert('⚠️ ყველა გენერირებული დოკუმენტის წაშლა',
    'სტუდენტები: ' + students.length + '\n' +
    'მოდულები: ყველა სტუდენტის ყველა ქვეფოლდერი\n\n' +
    'წაიშლება მხოლოდ სისტემის გენერირებული ფაილები, რომლებიც იწყება ასე:\n' +
    '• შეფასების ინსტრუმენტი - [მოდული]\n' +
    '• შეფასების ფურცელი - [მოდული]\n\n' +
    'თუ დროის ლიმიტი გაჩერდება, გამოიყენეთ:\n' +
    'ბეჩ ოპერაციები → გააგრძელე შეწყვეტილი ოპერაცია\n\n' +
    'დავიწყოთ?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  setConfig(CONFIG_KEYS.LAST_BATCH_OPERATION, 'deleteDocs');
  setConfig(CONFIG_KEYS.LAST_BATCH_TOTAL, students.length);
  setConfig(CONFIG_KEYS.LAST_BATCH_RANGE, "'1-" + students.length);
  setConfig(CONFIG_KEYS.LAST_BATCH_MODULES, JSON.stringify([DELETE_ALL_GENERATED_MARKER]));
  setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, 0);
  setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, 'deleteDocs');
  _deleteDocsForRange(1, students.length, [DELETE_ALL_GENERATED_MARKER], true);
}

function _isGeneratedArtifactName(fileName, moduleName) {
  return fileName === 'შეფასების ინსტრუმენტი - ' + moduleName ||
    fileName === 'შეფასების ფურცელი - ' + moduleName;
}

function _isGeneratedArtifactFileName(fileName) {
  return fileName.indexOf('შეფასების ინსტრუმენტი - ') === 0 ||
    fileName.indexOf('შეფასების ფურცელი - ') === 0;
}

function _isDeleteAllGeneratedMode(mods) {
  return mods && mods.length === 1 && mods[0] === DELETE_ALL_GENERATED_MARKER;
}

function _deleteGeneratedFilesInModuleFolder(moduleFolder, moduleName) {
  var deleted = 0, errors = 0;
  var files = moduleFolder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (_isGeneratedArtifactName(file.getName(), moduleName)) {
      try { file.setTrashed(true); deleted++; } catch(e) { errors++; }
    }
  }
  return { deleted: deleted, errors: errors };
}

function _deleteGeneratedFilesByPrefix(folder) {
  var deleted = 0, errors = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (_isGeneratedArtifactFileName(file.getName())) {
      try { file.setTrashed(true); deleted++; } catch(e) { errors++; }
    }
  }
  return { deleted: deleted, errors: errors };
}

function _deleteAllGeneratedFilesInStudentFolder(studentFolder) {
  var totals = _deleteGeneratedFilesByPrefix(studentFolder);
  var subfolders = studentFolder.getFolders();
  while (subfolders.hasNext()) {
    var sub = subfolders.next();
    var r = _deleteGeneratedFilesByPrefix(sub);
    totals.deleted += r.deleted;
    totals.errors += r.errors;
  }
  return totals;
}

function _deleteDocsForRange(startIdx, endIdx, mods, showUI) {
  var t = new Date(), MAX = 4.5 * 60 * 1000;
  var ids = JSON.parse(getConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS) || '{}');
  var students = getStudentList(), deleted = 0, processed = 0, skipped = 0, missingModules = 0, errors = 0;
  var origRange = String(getConfig(CONFIG_KEYS.LAST_BATCH_RANGE) || '').replace(/^'/, '');
  var origStart = parseInt(origRange.split('-')[0]) || startIdx;
  var deleteAllMode = _isDeleteAllGeneratedMode(mods);
  logAudit('delete', startIdx + '-' + endIdx, 'running', deleteAllMode ? 'All generated files' : 'Modules: ' + mods.length, 0);
  var mRow = 2;

  for (var i = startIdx - 1; i < endIdx && i < students.length; i++) {
    if (new Date() - t > MAX) {
      setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, (i + 1) - origStart);
      getAuditLogSheet().getRange(mRow, 4).setValue('timeout');
      logAuditUpdate(mRow, 'TIMEOUT. Students: ' + processed + '. Deleted: ' + deleted + '. Missing module folders: ' + missingModules + '. Errors: ' + errors, (new Date() - t) / 1000);
      if (String(getConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE)) === 'deleteDocs') {
        _clearDeleteTriggers();
        ScriptApp.newTrigger('_autoDeleteBatch').timeBased().after(2 * 60 * 1000).create();
        logAudit('delete:auto-scheduled', origRange || (startIdx + '-' + endIdx), 'waiting', 'Progress: ' + ((i + 1) - origStart) + '/' + (parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL) || '0')) + '. Next in 2 min.', 0);
      }
      if (showUI) SpreadsheetApp.getUi().alert('⏸️',
        'დროის ლიმიტი.\n' +
        'დამუშავებული სტუდენტები ამ რანში: ' + processed + '\n' +
        'წაშლილი ფაილები ამ რანში: ' + deleted + '\n' +
        'პროგრესი: ' + ((i + 1) - origStart) + '/' + (parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL) || '0')) + '\n\n' +
        (String(getConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE)) === 'deleteDocs'
          ? 'ავტომატურად გაგრძელდება დაახლოებით 2 წუთში.\nშეგიძლიათ დახუროთ ეს ფანჯარა.'
          : 'გასაგრძელებლად:\nბეჩ ოპერაციები → გააგრძელე შეწყვეტილი ოპერაცია'),
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var sid = ids[students[i].id];
    if (!sid) { skipped++; processed++; continue; }
    try {
      var sf = DriveApp.getFolderById(sid);
      if (deleteAllMode) {
        var allResult = _deleteAllGeneratedFilesInStudentFolder(sf);
        deleted += allResult.deleted;
        errors += allResult.errors;
      } else {
        for (var m = 0; m < mods.length; m++) {
          var mfi = sf.getFoldersByName(mods[m]);
          if (!mfi.hasNext()) { missingModules++; continue; }
          var modResult = _deleteGeneratedFilesInModuleFolder(mfi.next(), mods[m]);
          deleted += modResult.deleted;
          errors += modResult.errors;
        }
      }
    } catch (e) { errors++; }
    processed++;
    setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, (i + 1) - origStart + 1);
    if (processed % 10 === 0) logAuditUpdate(mRow, 'Students: ' + processed + '. Deleted: ' + deleted + '. Missing module folders: ' + missingModules + '. Errors: ' + errors, (new Date() - t) / 1000);
    Utilities.sleep(200);
  }
  setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, endIdx - origStart + 1);
  if (String(getConfig(CONFIG_KEYS.LAST_BATCH_OPERATION)) === 'deleteDocs') {
    setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
    _clearDeleteTriggers();
  }
  getAuditLogSheet().getRange(mRow, 4).setValue('success');
  logAuditUpdate(mRow, 'DONE. Students: ' + processed + '. Deleted: ' + deleted + '. Skip: ' + skipped + '. Missing module folders: ' + missingModules + '. Errors: ' + errors, (new Date() - t) / 1000);
  if (showUI) SpreadsheetApp.getUi().alert('✅',
    'დასრულდა ეს რანი.\n' +
    'დამუშავებული სტუდენტები: ' + processed + '\n' +
    'წაშლილი ფაილები: ' + deleted + '\n' +
    'გამოტოვებული სტუდენტები: ' + skipped + '\n' +
    'არარსებული მოდულის ფოლდერები: ' + missingModules + '\n' +
    'შეცდომები: ' + errors + '\n' +
    'დრო: ' + ((new Date()-t)/1000).toFixed(1) + ' წამი\n\n' +
    '⚠️ ფაილები Trash-შია.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function _autoDeleteBatch() {
  if (String(getConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE)) !== 'deleteDocs') return;
  if (String(getConfig(CONFIG_KEYS.LAST_BATCH_OPERATION)) !== 'deleteDocs') {
    setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
    _clearDeleteTriggers();
    logAudit('delete:auto-error', '', 'error', 'LAST_BATCH_OPERATION is not deleteDocs', 0);
    return;
  }

  var progress = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS) || '0');
  var total = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL) || '0');
  var range = String(getConfig(CONFIG_KEYS.LAST_BATCH_RANGE) || '').replace(/^'/, '');
  if (progress >= total || !range) {
    setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
    _clearDeleteTriggers();
    logAudit('delete:auto-complete', range, 'success', 'Finished. ' + progress + '/' + total, 0);
    return;
  }

  var parts = range.split('-'), start = parseInt(parts[0]), end = parseInt(parts[1]);
  var resumeFrom = start + progress;
  var mods = null;
  try { var j = getConfig(CONFIG_KEYS.LAST_BATCH_MODULES); if (j) mods = JSON.parse(j); } catch(e) {}
  if (!mods || mods.length === 0) {
    setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
    _clearDeleteTriggers();
    logAudit('delete:auto-error', range, 'error', 'No modules saved for delete resume.', 0);
    return;
  }

  _deleteDocsForRange(resumeFrom, end, mods, false);

  var newProgress = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS) || '0');
  if (newProgress >= total) {
    setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, '');
    _clearDeleteTriggers();
    logAudit('delete:auto-complete', range, 'success', 'Finished. ' + newProgress + '/' + total, 0);
  }
}

function _clearDeleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) { if (triggers[i].getHandlerFunction() === '_autoDeleteBatch') ScriptApp.deleteTrigger(triggers[i]); }
}

function showDeleteProgress() {
  var ui = SpreadsheetApp.getUi();
  var op = getConfig(CONFIG_KEYS.LAST_BATCH_OPERATION);
  var progress = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS) || '0');
  var total = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL) || '0');
  var range = String(getConfig(CONFIG_KEYS.LAST_BATCH_RANGE) || '').replace(/^'/, '');
  var modules = [];
  try { modules = JSON.parse(getConfig(CONFIG_KEYS.LAST_BATCH_MODULES) || '[]'); } catch(e) {}
  if (op !== 'deleteDocs') {
    ui.alert('წაშლის პროგრესი', 'ბოლო ოპერაცია არ არის წაშლა.\nბოლო ოპერაცია: ' + (op || '—'), ui.ButtonSet.OK);
    return;
  }
  var parts = range ? range.split('-') : [];
  var start = parseInt(parts[0]) || 1;
  var next = start + progress;
  var pct = total ? Math.min(100, Math.round((progress / total) * 100)) : 0;
  var moduleText = _isDeleteAllGeneratedMode(modules) ? 'ყველა გენერირებული ფაილი' : String(modules.length);
  var msg = 'ოპერაცია: deleteDocs\n' +
    'დიაპაზონი: ' + (range || '?') + '\n' +
    'პროგრესი: ' + progress + '/' + total + ' (' + pct + '%)\n' +
    'შემდეგი სტუდენტი: ' + (progress >= total ? 'დასრულებულია' : next) + '\n' +
    'დარჩა: ' + Math.max(0, total - progress) + '\n' +
    'მოდულები: ' + moduleText + '\n\n' +
    (progress >= total ? '✅ დასრულებულია' : '⏸️ გასაგრძელებლად გამოიყენეთ: ბეჩ ოპერაციები → გააგრძელე შეწყვეტილი ოპერაცია');
  ui.alert('წაშლის პროგრესი', msg, ui.ButtonSet.OK);
}


// ─────────────────────────────────────────────────────────────
// GROUP SHEETS
// ─────────────────────────────────────────────────────────────

function generateGroupSheet() {
  var ui = SpreadsheetApp.getUi();
  var eid = getConfig(CONFIG_KEYS.EVAL_SHEETS_FOLDER_ID);
  if (!eid) { ui.alert('შეცდომა', 'Eval Sheets ფოლდერი არ მოიძებნა.', ui.ButtonSet.OK); return; }
  var mod = promptForModule(); if (!mod) return;
  try { _createGroupSheet(mod, eid); ui.alert('✅', mod, ui.ButtonSet.OK); }
  catch (e) { ui.alert('შეცდომა', e.message, ui.ButtonSet.OK); }
}

function generateAllBlankDocs() {
  var ui = SpreadsheetApp.getUi();
  var eid = getConfig(CONFIG_KEYS.EVAL_SHEETS_FOLDER_ID);
  if (!eid) { ui.alert('შეცდომა', 'ფოლდერი არ მოიძებნა.', ui.ButtonSet.OK); return; }
  if (ui.alert('ჯგუფური', '13 ფურცელი.\nგნებავთ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var c = 0, errs = [];
  for (var m = 0; m < MODULE_NAMES.length; m++) { try { _createGroupSheet(MODULE_NAMES[m], eid); c++; } catch (e) { errs.push(MODULE_NAMES[m].substring(0,20)); } }
  ui.alert('✅', c + '/13' + (errs.length > 0 ? '\n\nშეცდომები: ' + errs.join(', ') : ''), ui.ButtonSet.OK);
}

function _createGroupSheet(moduleName, evalFolderId) {
  var folder = DriveApp.getFolderById(evalFolderId);
  var fn = moduleName + ' - შეფასების ფურცელი';
  var ex = folder.getFilesByName(fn); if (ex.hasNext()) throw new Error('არსებობს: ' + fn);
  var students = getStudentList(), outcomes = _getModuleOutcomeNames(moduleName);
  var ss = SpreadsheetApp.create(fn), sheet = ss.getSheets()[0]; sheet.setName('შეფასების ფურცელი');
  sheet.getRange(1, students.length + 3).setValue('დანართი N1').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(2, 1).setValue('პრაქტიკული დავალების შეფასების ფურცელი').setFontWeight('bold').setFontSize(12);
  sheet.getRange(2, 1, 1, 3).merge();
  sheet.getRange(3, 1).setValue('შეფასების კრიტერიუმები').setFontWeight('bold');
  for (var s = 0; s < students.length; s++) sheet.getRange(3, s+3).setValue((s+1)+'. '+students[s].name+' '+students[s].surname).setFontWeight('bold').setFontSize(8).setTextRotation(90);
  for (var c = 0; c < outcomes.length; c++) { sheet.getRange(5+c, 1).setValue(c+1).setHorizontalAlignment('center'); sheet.getRange(5+c, 2).setValue(outcomes[c]).setWrap(true); }
  sheet.getRange(5 + outcomes.length + 1, 1).setValue('პროფესიული მასწავლებელი/');
  sheet.setColumnWidth(1, 40); sheet.setColumnWidth(2, 400);
  for (var s = 0; s < students.length; s++) sheet.setColumnWidth(s+3, 30);
  sheet.getRange(3, 1, outcomes.length+2, students.length+2).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(4); sheet.setFrozenColumns(2);
  var file = DriveApp.getFileById(ss.getId()); folder.addFile(file); DriveApp.getRootFolder().removeFile(file);
  return ss.getId();
}

function _getModuleOutcomeNames(mod) {
  try { var vars = _readPaperVariables(mod); var n = []; for (var i = 1; i <= 13; i++) { var v = vars['{{OUTCOME_NAME_'+i+'}}']||''; if (v) n.push(v); } if (n.length > 0) return n; } catch(e) {}
  return _getStudentOutcomes(mod, 'DUMMY').outcomeNames;
}


// ─────────────────────────────────────────────────────────────
// FOLDER CREATION
// ─────────────────────────────────────────────────────────────

function createFullStructure() {
  var ui = SpreadsheetApp.getUi();
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { ui.alert('შეცდომა', 'ფოლდერი არ არის.', ui.ButtonSet.OK); return; }
  var students = getStudentList();
  if (ui.alert('სტრუქტურა', students.length + ' სტუდენტი + ' + MODULE_NAMES.length + ' ქვეფოლდერი.\nგნებავთ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _createFoldersForRange(1, students.length, true);
}

function createStudentFolders() {
  var ui = SpreadsheetApp.getUi();
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { ui.alert('შეცდომა', 'ფოლდერი არ არის.', ui.ButtonSet.OK); return; }
  var students = getStudentList();
  if (ui.alert('ფოლდერები', students.length + '\nგნებავთ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _createFoldersForRange(1, students.length, false);
}

function createModuleSubfolders() {
  var ui = SpreadsheetApp.getUi();
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { ui.alert('შეცდომა', 'ფოლდერი არ არის.', ui.ButtonSet.OK); return; }
  var ids = JSON.parse(getConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS) || '{}');
  if (Object.keys(ids).length === 0) { ui.alert('შეცდომა', 'ჯერ შექმენით ფოლდერები.', ui.ButtonSet.OK); return; }
  if (ui.alert('ქვეფოლდერები', MODULE_NAMES.length + ' თითოეულში.\nგნებავთ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _addModuleSubfoldersForRange(1, getStudentList().length);
}

function batchCreateFolders() {
  if (!getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)) { SpreadsheetApp.getUi().alert('შეცდომა', 'ფოლდერი არ არის.', SpreadsheetApp.getUi().ButtonSet.OK); return; }
  var range = promptForRange(); if (!range) return;
  _createFoldersForRange(range.start, range.end, true);
}

function _createFoldersForRange(startIdx, endIdx, withModules) {
  var t = new Date(), MAX = 5*60*1000;
  var students = getStudentList(), indiv = DriveApp.getFolderById(getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID));
  var ids = JSON.parse(getConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS) || '{}');
  var existing = {}; var iter = indiv.getFolders(); while (iter.hasNext()) { var f = iter.next(); existing[f.getName()] = f.getId(); }
  var created = 0, skipped = 0, mods = 0, last = startIdx - 1;
  var op = withModules ? 'createFullStructure' : 'createStudentFolders';
  setConfig(CONFIG_KEYS.LAST_BATCH_OPERATION, op); setConfig(CONFIG_KEYS.LAST_BATCH_TOTAL, endIdx-startIdx+1); setConfig(CONFIG_KEYS.LAST_BATCH_RANGE, "'"+startIdx+'-'+endIdx);
  for (var i = startIdx-1; i < endIdx && i < students.length; i++) {
    if (new Date()-t > MAX) { setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, last-startIdx+1); setConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS, JSON.stringify(ids));
      logAudit(op, startIdx+'-'+endIdx, 'timeout', 'Created: '+created, (new Date()-t)/1000);
      SpreadsheetApp.getUi().alert('⏸️', 'შეიქმნა '+created+'.\nგააგრძელე.', SpreadsheetApp.getUi().ButtonSet.OK); return; }
    var s = students[i], fn = getStudentFolderName(s);
    if (existing[fn]) { ids[s.id] = existing[fn]; skipped++; if (withModules) mods += _ensureModuleSubfolders(existing[fn]); }
    else { var nf = indiv.createFolder(fn); ids[s.id] = nf.getId(); existing[fn] = nf.getId(); created++;
      if (withModules) { for (var m = 0; m < MODULE_NAMES.length; m++) { nf.createFolder(MODULE_NAMES[m]); mods++; } } }
    last = i+1; Utilities.sleep(500);
    if (created % 20 === 0 && created > 0) { setConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS, JSON.stringify(ids)); setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, last-startIdx+1); }
  }
  setConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS, JSON.stringify(ids)); setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, endIdx-startIdx+1);
  logAudit(op, startIdx+'-'+endIdx, 'success', 'Created: '+created+', Skip: '+skipped+', Mods: '+mods, (new Date()-t)/1000);
  SpreadsheetApp.getUi().alert('✅', 'შეიქმნა '+created+'\nგამოიტოვა '+skipped+'\nმოდულები: '+mods+'\nდრო: '+((new Date()-t)/1000).toFixed(1)+' წამი', SpreadsheetApp.getUi().ButtonSet.OK);
}

function _ensureModuleSubfolders(fid) {
  var folder = DriveApp.getFolderById(fid), ex = {};
  var subs = folder.getFolders(); while (subs.hasNext()) { var s = subs.next(); ex[s.getName()] = true; }
  var c = 0; for (var m = 0; m < MODULE_NAMES.length; m++) { if (!ex[MODULE_NAMES[m]]) { folder.createFolder(MODULE_NAMES[m]); c++; Utilities.sleep(200); } } return c;
}

function _addModuleSubfoldersForRange(startIdx, endIdx) {
  var t = new Date(), MAX = 5*60*1000, students = getStudentList(), ids = JSON.parse(getConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS)||'{}');
  var mods = 0, proc = 0;
  setConfig(CONFIG_KEYS.LAST_BATCH_OPERATION, 'createModuleSubfolders'); setConfig(CONFIG_KEYS.LAST_BATCH_TOTAL, endIdx-startIdx+1); setConfig(CONFIG_KEYS.LAST_BATCH_RANGE, "'"+startIdx+'-'+endIdx);
  for (var i = startIdx-1; i < endIdx && i < students.length; i++) {
    if (new Date()-t > MAX) { setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, proc);
      logAudit('moduleSubfolders', startIdx+'-'+endIdx, 'timeout', 'Mods: '+mods, (new Date()-t)/1000);
      SpreadsheetApp.getUi().alert('⏸️', 'გააგრძელე.', SpreadsheetApp.getUi().ButtonSet.OK); return; }
    var fid = ids[students[i].id]; if (fid) mods += _ensureModuleSubfolders(fid);
    proc++; if (proc % 20 === 0) setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, proc);
  }
  setConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS, proc);
  logAudit('moduleSubfolders', startIdx+'-'+endIdx, 'success', 'Mods: '+mods, (new Date()-t)/1000);
  SpreadsheetApp.getUi().alert('✅', 'მოდულები: '+mods+'\nდრო: '+((new Date()-t)/1000).toFixed(1)+' წამი', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ─────────────────────────────────────────────────────────────
// RESUME / SETTINGS
// ─────────────────────────────────────────────────────────────

function resumeLastOperation() {
  var ui = SpreadsheetApp.getUi();
  var op = getConfig(CONFIG_KEYS.LAST_BATCH_OPERATION), progress = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS)||'0');
  var total = parseInt(getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL)||'0'), range = String(getConfig(CONFIG_KEYS.LAST_BATCH_RANGE)||'').replace(/^'/,'');
  if (!op || !range) { ui.alert('ინფო', 'არაფერი გასაგრძელებელი.', ui.ButtonSet.OK); return; }
  if (progress >= total) { ui.alert('ინფო', 'დასრულებულია.', ui.ButtonSet.OK); return; }
  var parts = range.split('-'), start = parseInt(parts[0]), end = parseInt(parts[1]), from = start + progress;
  if (ui.alert('▶', op + '\n' + progress + '/' + total + '\nგაგრძელდება: ' + from + '-' + end, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  if (op === 'createFullStructure') _createFoldersForRange(from, end, true);
  else if (op === 'createStudentFolders') _createFoldersForRange(from, end, false);
  else if (op === 'createModuleSubfolders') _addModuleSubfoldersForRange(from, end);
  else if (op === 'generateDocs') { var m = null; try { m = JSON.parse(getConfig(CONFIG_KEYS.LAST_BATCH_MODULES)); } catch(e) {} _runGeneration(from, end, m, true); }
  else if (op === 'deleteDocs') { var dm = null; try { dm = JSON.parse(getConfig(CONFIG_KEYS.LAST_BATCH_MODULES)); } catch(e) {} if (!dm || dm.length === 0) dm = _promptForDeleteModules(); if (dm) { setConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE, 'deleteDocs'); _deleteDocsForRange(from, end, dm, true); } }
  else ui.alert('შეცდომა', 'უცნობი: ' + op, ui.ButtonSet.OK);
}

function showBatchProgress() {
  var op = getConfig(CONFIG_KEYS.LAST_BATCH_OPERATION), p = getConfig(CONFIG_KEYS.LAST_BATCH_PROGRESS),
    t = getConfig(CONFIG_KEYS.LAST_BATCH_TOTAL), r = getConfig(CONFIG_KEYS.LAST_BATCH_RANGE);
  var auto = getConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE);
  var autoMsg = String(auto) === 'true' ? '\n\n🌙 ავტო-გენერაცია აქტიურია\nTrigger: ' + _countAutoTriggers() : (String(auto) === 'deleteDocs' ? '\n\n🧹 ავტო-წაშლა აქტიურია' : '');
  var msg = !op ? 'ბეჩი არ შესრულებულა.' : op + '\nდიაპაზონი: ' + (r||'?') + '\nპროგრესი: ' + (p||0) + '/' + (t||0) +
    '\n' + (p==t ? '✅ დასრულდა' : '⏸️ შეწყვეტილი') + autoMsg;
  SpreadsheetApp.getUi().alert('ბეჩი', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function openAuditLog() { SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getAuditLogSheet()); }

function showWorkingFolderId() {
  var id = getConfig(CONFIG_KEYS.WORKING_FOLDER_ID);
  SpreadsheetApp.getUi().alert('ფოლდერი', !id ? 'არ არის.' : id + '\nhttps://drive.google.com/drive/folders/' + id +
    '\n\nEval: ' + (getConfig(CONFIG_KEYS.EVAL_SHEETS_FOLDER_ID)||'❌') +
    '\nIndiv: ' + (getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)||'❌') +
    '\nTemplates: ' + (getConfig(CONFIG_KEYS.TEMPLATES_FOLDER_ID)||'❌'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function showSystemStatus() {
  var sl = getStudentList(), fc = 0;
  try { fc = Object.keys(JSON.parse(getConfig(CONFIG_KEYS.STUDENT_FOLDER_IDS)||'{}')).length; } catch(e) {}
  var msg = '═══ სტატუსი ═══\n';
  msg += '\nფოლდერი: '+(getConfig(CONFIG_KEYS.WORKING_FOLDER_ID)?'✅':'❌');
  msg += '\nEval Sheets: '+(getConfig(CONFIG_KEYS.EVAL_SHEETS_FOLDER_ID)?'✅':'❌');
  msg += '\nIndividual: '+(getConfig(CONFIG_KEYS.INDIVIDUAL_EVAL_FOLDER_ID)?'✅':'❌');
  msg += '\nTemplates: '+(getConfig(CONFIG_KEYS.TEMPLATES_FOLDER_ID)?'✅':'❌');
  msg += '\n\nDoc Template: '+(getConfig(CONFIG_KEYS.EVAL_DOC_TEMPLATE_ID)?'✅':'❌');
  msg += '\nPaper Template: '+(getConfig(CONFIG_KEYS.EVAL_PAPER_TEMPLATE_ID)?'✅':'❌');
  msg += '\nDoc Vars: '+(getConfig(CONFIG_KEYS.DOC_VARS_SHEET_ID)?'✅':'❌');
  msg += '\nPaper Vars: '+(getConfig(CONFIG_KEYS.PAPER_VARS_SHEET_ID)?'✅':'❌');
  msg += '\n\nLogging: '+(String(getConfig(CONFIG_KEYS.DETAILED_LOGGING))==='true'?'🔍 Detailed':'⚡ Light');
  var auto = String(getConfig(CONFIG_KEYS.AUTO_RUN_ACTIVE));
  msg += '\nAuto-run: '+(auto==='true'?'🌙 Generate Active':(auto==='deleteDocs'?'🧹 Delete Active':'⏹️ Off'));
  msg += '\n\nსტუდენტები: '+sl.length+'\nფოლდერები: '+fc+'\nმოდულები: '+MODULE_NAMES.length;
  SpreadsheetApp.getUi().alert('სტატუსი', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function enableDetailedLogging() {
  setConfig(CONFIG_KEYS.DETAILED_LOGGING, 'true');
  SpreadsheetApp.getUi().alert('✅', 'დეტალური ლოგი ჩართულია.\nთითოეული სტუდენტი ცალკე სტრიქონში.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function disableDetailedLogging() {
  setConfig(CONFIG_KEYS.DETAILED_LOGGING, 'false');
  SpreadsheetApp.getUi().alert('✅', 'ლაით ლოგი.\nმხოლოდ ყოველ 10 სტუდენტზე განახლდება.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function refreshVariableSheets() {
  var ui = SpreadsheetApp.getUi(), wid = getConfig(CONFIG_KEYS.WORKING_FOLDER_ID);
  if (!wid) { ui.alert('შეცდომა', 'ფოლდერი არ არის.', ui.ButtonSet.OK); return; }
  var r = _discoverTemplates(wid);
  ui.alert('✅', 'Doc Vars: '+(r.docVarsSheet?'✅':'❌')+'\nPaper Vars: '+(r.paperVarsSheet?'✅':'❌'), ui.ButtonSet.OK);
}

function resetConfig() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️', 'ყველა პარამეტრი წაიშლება.\nDrive რჩება.\nგნებავთ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  if (ui.alert('⚠️ დარწმუნებული?', 'შეუქცევადი.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _clearAutoTriggers();
  _clearDeleteTriggers();
  Object.values(CONFIG_KEYS).forEach(function(k) { setConfig(k, (k==='STUDENT_FOLDER_IDS'||k==='LAST_BATCH_MODULES')?'{}': k==='DETAILED_LOGGING'?'true':''); });
  logAudit('reset', 'all', 'success', 'Cleared');
  ui.alert('✅', 'გადატვირთულია.', ui.ButtonSet.OK);
}

function initializeSystem() {
  getConfigSheet(); getAuditLogSheet();
  logAudit('init', 'system', 'success', 'Ready');
  SpreadsheetApp.getUi().alert('✅', 'მზადაა. შემდეგი: აირჩიეთ სამუშაო ფოლდერი.', SpreadsheetApp.getUi().ButtonSet.OK);
}
