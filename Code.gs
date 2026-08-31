function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏫 ប្រព័ន្ធគ្រប់គ្រងសិស្ស')
    .addItem('🚀 បើកប្រព័ន្ធគ្រប់គ្រង (Web App)', 'openWebApp')
    .addToUi();
}

function openWebApp() {
  var url = getWebAppUrl();
  var html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:sans-serif;text-align:center;padding:40px;">' +
    '<p style="font-size:14px;color:#334155;">ចុចប៊ូតុងខាងក្រោមដើម្បីបើកប្រព័ន្ធគ្រប់គ្រងសិស្ស៖</p>' +
    '<a href="' + url + '" target="_blank" ' +
    'style="display:inline-block;margin-top:14px;padding:12px 26px;background:#2563eb;' +
    'color:#fff;border-radius:10px;text-decoration:none;font-weight:bold;">🚀 បើកប្រព័ន្ធគ្រប់គ្រង</a>' +
    '<p style="margin-top:16px;font-size:11px;color:#94a3b8;">ប្រសិនបើលើកទីមួយ សូម Deploy ជា Web App សិន</p>' +
    '</body></html>'
  ).setWidth(420).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'បើកប្រព័ន្ធគ្រប់គ្រង');
}

// ==========================================================================
// 📚 KEY ARCHITECTURE & CONSTANTS
// ==========================================================================

var MONTHLY_KEYS = [
  // ភាសាខ្មែរ (4)
  'kh_listen','kh_write','kh_read','kh_speak',
  // គណិតវិទ្យា (5)
  'm_num','m_measure','m_geo','m_algebra','m_stat',
  // វិទ្យាសាស្ត្រ + សិក្សាសង្គម (2)
  'science','social',
  // អប់រំ/សិល្បៈ (4)
  'a_craft','a_pe','a_health','a_env',
  // ភាសាបរទេស (1)
  'foreign'
];

var SEMESTER_KEYS = [
  // ភាសាខ្មែរ (4)
  'read','listen','dictation','essay',
  // គណិតវិទ្យា (1)
  'math',
  // វិទ្យាសាស្ត្រ + សិក្សាសង្គម (2)
  'science','social',
  // អប់រំ/សិល្បៈ (3)
  'homeArts','pe','ethics',
  // ភាសាបរទេស (1)
  'foreign'
];

var DROPOUT_STATUS_COL = 16;
var SEMESTER1_MONTHS   = ['វិច្ឆិកា','ធ្នូ','មករា','កុម្ភៈ','មីនា','មេសា'];
var SEMESTER2_MONTHS   = ['ឧសភា','មិថុនា','កក្កដា','សីហា','កញ្ញា','តុលា'];
var KHMER_MONTH_NAMES  = ['មករា','កុម្ភៈ','មីនា','មេសា','ឧសភា','មិថុនា','កក្កដា','សីហា','កញ្ញា','តុលា','វិច្ឆិកា','ធ្នូ'];

// ==========================================================================
// 🚪 Dropout helpers
// ==========================================================================
function getDropoutMap() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  var map = {};
  if (!sheet || sheet.getLastRow() < 8) return map;
  var data = sheet.getRange(8, 2, sheet.getLastRow() - 7, DROPOUT_STATUS_COL - 1).getValues();
  for (var i = 0; i < data.length; i++) {
    var id = data[i][0] ? data[i][0].toString().trim() : '';
    if (id === '') continue;
    map[id] = data[i][DROPOUT_STATUS_COL - 2] ? data[i][DROPOUT_STATUS_COL - 2].toString().trim() : '';
  }
  return map;
}

function getSemesterOfMonth(month) {
  if (SEMESTER1_MONTHS.indexOf(month) !== -1) return 1;
  if (SEMESTER2_MONTHS.indexOf(month) !== -1) return 2;
  return 0;
}

function isStudentActiveForMonth(dropoutSemester, month) {
  if (!dropoutSemester) return true;
  var dropSem = dropoutSemester === 'ឆមាសទី១' ? 1 : dropoutSemester === 'ឆមាសទី២' ? 2 : 0;
  if (dropSem === 0) return true;
  var monthSem = getSemesterOfMonth(month);
  if (monthSem === 0) return true;
  return monthSem <= dropSem;
}

function isStudentActiveForSemesterKey(dropoutSemester, semesterKey) {
  if (!dropoutSemester) return true;
  var dropSem = dropoutSemester === 'ឆមាសទី១' ? 1 : dropoutSemester === 'ឆមាសទី២' ? 2 : 0;
  if (dropSem === 0) return true;
  var targetSem = semesterKey === 'ឆមាសទី១' ? 1 : semesterKey === 'ឆមាសទី២' ? 2 : 0;
  if (targetSem === 0) return true;
  return targetSem <= dropSem;
}

function setStudentDropout(studentId, semester) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    if (!sheet || sheet.getLastRow() < 8) throw new Error('រកមិនឃើញទិន្នន័យសិស្សទេ');
    var ids = sheet.getRange(8, 2, sheet.getLastRow() - 7, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if ((ids[i][0] || '').toString().trim() === studentId.toString().trim()) {
        sheet.getRange(i + 8, DROPOUT_STATUS_COL).setValue(semester || '');
        return semester
          ? 'បានកត់ត្រាថាសិស្សបានបោះបង់ការសិក្សា (' + semester + ') រួចរាល់!'
          : 'បានលុបចោលស្ថានភាពបោះបង់ — សិស្សត្រឡប់ជាសកម្មវិញ!';
      }
    }
    throw new Error('រកមិនឃើញអត្តលេខ «' + studentId + '»');
  } catch (err) { throw new Error('ការកត់ត្រាបរាជ័យ៖ ' + err.toString()); }
}

// ==========================================================================
// 📋 Student list helpers
// ==========================================================================
function getStudentsByClass(selectedClass) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    if (!sheet || sheet.getLastRow() < 8) return [];
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, 8).getValues();
    var list = [], target = selectedClass ? selectedClass.toString().toLowerCase().replace(/\s+/g,'') : '';
    for (var i = 0; i < data.length; i++) {
      var rawGrade = data[i][6] ? data[i][6].toString() : '';
      var sg = rawGrade.toLowerCase().replace(/\s+/g,'');
      if (!selectedClass || target==='' || target==='all' || target==='ទាំងអស់' || sg===target) {
        if (data[i][2] && data[i][2].toString().trim() !== '') {
          list.push({ id:data[i][1]?data[i][1].toString().trim():'', name:data[i][2].toString().trim(),
            sex:data[i][3]?data[i][3].toString().trim():'', grade:rawGrade.trim() });
        }
      }
    }
    return list;
  } catch(err) { throw new Error('កំហុសក្នុងការទាញយកទិន្នន័យ៖ '+err.toString()); }
}

function getStudentsByClassWithStatus(selectedClass) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    if (!sheet || sheet.getLastRow() < 8) return [];
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, DROPOUT_STATUS_COL).getValues();
    var list = [], target = selectedClass ? selectedClass.toString().toLowerCase().replace(/\s+/g,'') : '';
    for (var i = 0; i < data.length; i++) {
      var rawGrade = data[i][6] ? data[i][6].toString() : '';
      var sg = rawGrade.toLowerCase().replace(/\s+/g,'');
      if (!selectedClass || target==='' || target==='all' || target==='ទាំងអស់' || sg===target) {
        if (data[i][2] && data[i][2].toString().trim() !== '') {
          list.push({ id:data[i][1]?data[i][1].toString().trim():'', name:data[i][2].toString().trim(),
            sex:data[i][3]?data[i][3].toString().trim():'', grade:rawGrade.trim(),
            dropoutSemester:data[i][15]?data[i][15].toString().trim():'' });
        }
      }
    }
    return list;
  } catch(err) { throw new Error('កំហុសក្នុងការទាញយកទិន្នន័យ៖ '+err.toString()); }
}

function getStudentsByClassForMonth(selectedClass, month) {
  var all = getStudentsByClassWithStatus(selectedClass);
  return all.filter(function(s){ return isStudentActiveForMonth(s.dropoutSemester, month); });
}

function getStudentsForScoreEntry(selectedClass, period) {
  try {
    var list = getStudentsByClassWithStatus(selectedClass);
    return list.filter(function(s) {
      if (period === 'ឆមាសទី១' || period === 'ឆមាសទី២') {
        return isStudentActiveForSemesterKey(s.dropoutSemester, period);
      }
      return isStudentActiveForMonth(s.dropoutSemester, period);
    });
  } catch(err) { throw new Error('កំហុស getStudentsForScoreEntry: '+err.toString()); }
}

// ==========================================================================
// 💾 Save / Read scores (រក្សាទុកជួរឈរតាម Form ពិតប្រាកដ)
// ==========================================================================
// ── SAVE SCORES (រក្សាទុកតែតារាងពិន្ទុស្អាត គ្មានជួរឈរមេគុណលើស) ──
function saveMonthlyScores(payload) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var cls  = payload.className || 'ទូទៅ';
    var period = payload.month || payload.period || 'មិនបានកំណត់';
    var scoreType = (period === 'ឆមាសទី១' || period === 'ឆមាសទី២' || payload.scoreType === 'semester') ? 'semester' : 'monthly';
    var sheetName = 'ពិន្ទុ_ថ្នាក់ទី' + cls + '_' + period;

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName); else sheet.clear();

    var activeKeys = (scoreType === 'semester') ? SEMESTER_KEYS : MONTHLY_KEYS;
    var subjectLabels = payload.subjectLabels || activeKeys;
    var totalTableWidth = 4 + activeKeys.length + 3; // 4 ព័ត៌មាន + N មុខវិជ្ជា + 3 លទ្ធផល

    // Header block (rows 1-6)
    var hb = sheet.getRange(1, 1, 6, totalTableWidth);
    hb.breakApart(); hb.clearContent();

    sheet.getRange(1,1,1,totalTableWidth).merge().setValue('ព្រះរាជាណាចក្រកម្ពុជា')
      .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
    sheet.getRange(2,1,1,totalTableWidth).merge().setValue('ជាតិ សាសនា ព្រះមហាក្សត្រ')
      .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
    sheet.getRange('A3').setValue('មន្ទីរអប់រំ យុវជន និងកីឡា រាជធានី/ខេត្ត បាត់ដំបង').setFontWeight('bold');
    sheet.getRange('A4').setValue('ការិយាល័យអប់រំ យុវជន និងកីឡានៃរដ្ឋបាលក្រុង/ស្រុក/ខណ្ឌ សំឡូត').setFontWeight('bold');
    sheet.getRange('A5').setValue('សាលាបឋមសិក្សា ៖ កំពង់ល្ពៅ').setFontWeight('bold');

    var titleType = scoreType === 'semester' ? 'ប្រឡងឆមាស' : 'ប្រចាំខែ';
    sheet.getRange(6,1,1,totalTableWidth).merge()
      .setValue('តារាងស្រង់ពិន្ទុ' + titleType + ' ' + period + ' (ថ្នាក់ទី ' + cls + ')')
      .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    // Row 7: ចំណងជើងជួរឈរ (Headers)
    var headers = ['ល.រ','អត្តលេខ','ឈ្មោះសិស្ស','ភេទ']
      .concat(subjectLabels, ['ពិន្ទុសរុប','មធ្យមភាគ','ចំណាត់ថ្នាក់']);

    sheet.getRange(7, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#facc15')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,true,true);
    sheet.setRowHeight(7, 100);

    // រក្សាទុកមេគុណក្នុងប្រព័ន្ធសម្ងាត់ Background (មិនឱ្យលេចចេញលើ Sheet ឡើយ)
    if (payload.coefficients) {
      PropertiesService.getDocumentProperties().setProperty('COEF_' + sheetName, JSON.stringify(payload.coefficients));
    }

    // Data rows
    var incomingCoef = payload.coefficients || {};
    var rows = payload.scores.map(function(s) {
      var values = activeKeys.map(function(k) {
        var raw = s[k];
        if (raw === undefined || raw === null || raw === '') return null;
        var num = Number(raw);
        return isNaN(num) ? null : num;
      });

      var wTotal = 0, coefSum = 0;
      values.forEach(function(v, idx) {
        if (v !== null) {
          var k = activeKeys[idx];
          var coef = (scoreType === 'monthly') ? 1 : ((isNaN(Number(incomingCoef[k])) || Number(incomingCoef[k]) < 0) ? 1 : Number(incomingCoef[k]));
          wTotal  += v * coef;
          coefSum += coef;
        }
      });
      var avg = coefSum > 0 ? (wTotal / coefSum) : 0;

      return { id: s.id, name: s.name, sex: s.sex, values: values,
               total: parseFloat(wTotal.toFixed(2)), avg: parseFloat(avg.toFixed(2)) };
    });

    // ចំណាត់ថ្នាក់
    var sorted = rows.slice().sort(function(a,b){ return b.avg - a.avg; });
    var rankMap = {};
    sorted.forEach(function(r,i) {
      if (i > 0 && r.avg === sorted[i-1].avg) rankMap[r.id] = rankMap[sorted[i-1].id];
      else rankMap[r.id] = i + 1;
    });

    var outputRows = rows.map(function(r, idx) {
      var vals = r.values.map(function(v){ return v === null ? '' : v; });
      return [idx + 1, r.id, r.name, r.sex].concat(vals, [r.total, r.avg, rankMap[r.id]]);
    });

    if (outputRows.length > 0) {
      var dr = sheet.getRange(8, 1, outputRows.length, headers.length);
      dr.setValues(outputRows).setBorder(true,true,true,true,true,true).setHorizontalAlignment('center');
      sheet.getRange(8, 3, outputRows.length, 1).setHorizontalAlignment('left');
    }

    return 'បានរក្សាទុកពិន្ទុ «' + sheetName + '» រួចរាល់!';
  } catch(err) { throw new Error('ការរក្សាទុកបរាជ័យ៖ ' + err.message); }
}

// ── GET SCORES (ទាញយកពិន្ទុមកវិញដោយស្អាត) ──
function getSavedScores(className, period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'ពិន្ទុ_ថ្នាក់ទី' + className + '_' + period;
    var sheet = ss.getSheetByName(sheetName);
    var isSem = (period === 'ឆមាសទី១' || period === 'ឆមាសទី២');
    var keys = isSem ? SEMESTER_KEYS : MONTHLY_KEYS;
    var result = { labels:[], scores:{}, coefficients:{}, scoreType: isSem ? 'semester' : 'monthly' };
    if (!sheet || sheet.getLastRow() < 8) return result;

    var width = 4 + keys.length + 3;

    // ទាញយកមេគុណដែលបានលាក់ទុកក្នុង Background
    var rawCoef = PropertiesService.getDocumentProperties().getProperty('COEF_' + sheetName);
    if (rawCoef) {
      try { result.coefficients = JSON.parse(rawCoef); } catch(e){}
    }

    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, width).getValues();
    data.forEach(function(row) {
      var sid = row[1] ? row[1].toString().trim() : '';
      if (!sid) return;
      var entry = {};
      keys.forEach(function(k, idx) {
        var v = row[4 + idx];
        entry[k] = (v === '' || v === null || v === undefined) ? '' : v;
      });
      entry.total = row[width - 3];
      entry.avg   = row[width - 2];
      result.scores[sid] = entry;
    });
    return result;
  } catch(err) {
    Logger.log('getSavedScores error: ' + err);
    return { labels:[], scores:{}, coefficients:{}, scoreType:'monthly' };
  }
}

function getSavedScores(className, period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + className + '_' + period);
    var isSem = (period === 'ឆមាសទី១' || period === 'ឆមាសទី២');
    var defaultKeys = isSem ? SEMESTER_KEYS : MONTHLY_KEYS;
    var result = { labels:[], scores:{}, coefficients:{}, scoreType: isSem ? 'semester' : 'monthly' };
    if (!sheet || sheet.getLastRow() < 8) return result;

    var numSubj = defaultKeys.length;
    var width = 4 + numSubj + 3;
    var coefColStart = width + 2;

    // អាន Keys និង មេគុណដែលបានរក្សាទុក
    var keys = defaultKeys;
    if (sheet.getLastColumn() >= coefColStart + numSubj) {
      var storedKeys = sheet.getRange(7, coefColStart + numSubj, 1, numSubj).getValues()[0];
      if (storedKeys && storedKeys[0]) keys = storedKeys.map(function(k){ return String(k).trim(); });
      var coefVals = sheet.getRange(7, coefColStart, 1, numSubj).getValues()[0];
      keys.forEach(function(k, idx) {
        var v = Number(coefVals[idx]);
        result.coefficients[k] = isNaN(v) ? 1 : v;
      });
    }

    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, width).getValues();
    data.forEach(function(row) {
      var sid = row[1] ? row[1].toString().trim() : '';
      if (!sid) return;
      var entry = {};
      keys.forEach(function(k, idx) {
        var v = row[4 + idx];
        entry[k] = (v === '' || v === null || v === undefined) ? '' : v;
      });
      entry.total = row[width - 3];
      entry.avg   = row[width - 2];
      result.scores[sid] = entry;
    });
    return result;
  } catch(err) {
    Logger.log('getSavedScores error: ' + err);
    return { labels:[], scores:{}, coefficients:{}, scoreType:'monthly' };
  }
}

// ==========================================================================
// 📊 Rank computation helpers
// ==========================================================================
function calculateRankAndGrade(list) {
  list.sort(function(a,b){ return Number(b.avg) - Number(a.avg); });
  list.forEach(function(s,idx) {
    s.rank = (idx > 0 && Number(s.avg) === Number(list[idx-1].avg)) ? list[idx-1].rank : idx + 1;
    var avgNum = Number(s.avg);
    if      (avgNum >= 9.00){s.grade='ល្អប្រសើរ';s.gradeCode='A';s.color='#0277bd';}
    else if (avgNum >= 8.00){s.grade='ល្អណាស់'; s.gradeCode='B';s.color='#8bc34a';}
    else if (avgNum >= 7.00){s.grade='ល្អ';      s.gradeCode='C';s.color='#ffeb3b';}
    else if (avgNum >= 6.00){s.grade='បង្គួរ';   s.gradeCode='D';s.color='#673ab7';}
    else if (avgNum >= 5.00){s.grade='មធ្យម';    s.gradeCode='E';s.color='#00bcd4';}
    else                  {s.grade='ខ្សោយ';    s.gradeCode='F';s.color='#f44336';}
    s.result      = avgNum >= 5 ? 'ជាប់' : 'ធ្លាក់';
    s.resultColor = avgNum >= 5 ? '#000000' : '#dc2626';
  });
  return list;
}

function getRankData(className, month) {
  try {
    if (month === 'ប្រចាំឆមាសទី១') return getWeightedSemesterRankData(className, SEMESTER1_MONTHS, 'ឆមាសទី១');
    if (month === 'ប្រចាំឆមាសទី២') return getWeightedSemesterRankData(className, SEMESTER2_MONTHS, 'ឆមាសទី២');
    if (month === 'ប្រចាំឆ្នាំ')    return getYearlyRankData(className);

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + className + '_' + month);
    if (!sheet || sheet.getLastRow() < 8) return [];

    var isSem = (month === 'ឆមាសទី១' || month === 'ឆមាសទី២');
    var width = 4 + (isSem ? SEMESTER_KEYS.length : MONTHLY_KEYS.length) + 3;
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, width).getValues();
    var list = [];
    var dropoutMap = getDropoutMap();

    data.forEach(function(row) {
      var id   = row[1] ? row[1].toString().trim() : '';
      var name = row[2] ? row[2].toString().trim() : '';
      var sex  = row[3] ? row[3].toString().trim() : 'ប្រុស';
      if (!name) return;
      if (!isStudentActiveForMonth(dropoutMap[id], month)) return;
      list.push({ id:id, name:name, sex:sex,
        total: Number(row[width - 3]) || 0,
        avg:   Number(row[width - 2]) || 0 });
    });

    return calculateRankAndGrade(list.map(function(s){
      return Object.assign({}, s, { total: s.total.toFixed(2), avg: s.avg.toFixed(2) });
    }));
  } catch(err){ Logger.log('getRankData err: ' + err); return []; }
}

function getExamOnlyRankList(className, semesterKey) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + className + '_' + semesterKey);
    if (!sheet || sheet.getLastRow() < 8) return [];
    var width = 4 + SEMESTER_KEYS.length + 3;
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, width).getValues();
    var list = [], dropoutMap = getDropoutMap();
    data.forEach(function(row){
      var id = row[1] ? row[1].toString().trim() : '', name = row[2] ? row[2].toString().trim() : '', sex = row[3] ? row[3].toString().trim() : 'ប្រុស';
      if (!name) return;
      if (!isStudentActiveForSemesterKey(dropoutMap[id], semesterKey)) return;
      list.push({ id:id, name:name, sex:sex, total:(Number(row[width - 3])||0).toFixed(2), avg:(Number(row[width - 2])||0).toFixed(2) });
    });
    return calculateRankAndGrade(list);
  } catch(err){ Logger.log('getExamOnlyRankList err: ' + err); return []; }
}

function getExamOnlyYearlyRankList(className) {
  var sem1 = getExamOnlyRankList(className, 'ឆមាសទី១'), sem2 = getExamOnlyRankList(className, 'ឆមាសទី២');
  if (!sem1.length && !sem2.length) return [];
  var map = {};
  sem1.forEach(function(s){ map[s.id] = { id:s.id, name:s.name, sex:s.sex, avg1:parseFloat(s.avg) || 0 }; });
  sem2.forEach(function(s){ if(!map[s.id]) map[s.id] = { id:s.id, name:s.name, sex:s.sex }; map[s.id].avg2 = parseFloat(s.avg) || 0; });
  var list = [];
  for (var id in map) {
    var s = map[id], vals = [];
    if (s.avg1 !== undefined) vals.push(s.avg1);
    if (s.avg2 !== undefined) vals.push(s.avg2);
    var sum = vals.reduce(function(a,b){ return a + b; }, 0);
    list.push({ id:s.id, name:s.name, sex:s.sex, total:sum.toFixed(2), avg:(vals.length ? sum / vals.length : 0).toFixed(2) });
  }
  return calculateRankAndGrade(list);
}

function getExamBasedRankData(className, period) {
  if (period === 'ប្រចាំឆមាសទី១') return getExamOnlyRankList(className, 'ឆមាសទី១');
  if (period === 'ប្រចាំឆមាសទី២') return getExamOnlyRankList(className, 'ឆមាសទី២');
  if (period === 'ប្រចាំឆ្នាំ')    return getExamOnlyYearlyRankList(className);
  return getRankData(className, period);
}

function getWeightedSemesterRankData(className, monthList, examSuffix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), studentMap = {}, dropoutMap = getDropoutMap();
  monthList.forEach(function(m){
    var sheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + className + '_' + m);
    if (sheet && sheet.getLastRow() >= 8) {
      var w = 4 + MONTHLY_KEYS.length + 3;
      sheet.getRange(8, 1, sheet.getLastRow() - 7, w).getValues().forEach(function(row){
        var id = row[1] ? row[1].toString().trim() : '', name = row[2] ? row[2].toString().trim() : '', sex = row[3] ? row[3].toString().trim() : 'ប្រុស', avg = parseFloat(row[w - 2]) || 0;
        if (id && name) {
          if (!studentMap[id]) studentMap[id] = { id:id, name:name, sex:sex, monthTotal:0, monthCount:0, examAvg:null };
          studentMap[id].monthTotal += avg;
          studentMap[id].monthCount++;
        }
      });
    }
  });
  var examSheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + className + '_' + examSuffix);
  if (examSheet && examSheet.getLastRow() >= 8) {
    var ew = 4 + SEMESTER_KEYS.length + 3;
    examSheet.getRange(8, 1, examSheet.getLastRow() - 7, ew).getValues().forEach(function(row){
      var id = row[1] ? row[1].toString().trim() : '', name = row[2] ? row[2].toString().trim() : '', sex = row[3] ? row[3].toString().trim() : 'ប្រុស', avg = parseFloat(row[ew - 2]) || 0;
      if (id && name) {
        if (!studentMap[id]) studentMap[id] = { id:id, name:name, sex:sex, monthTotal:0, monthCount:0, examAvg:null };
        studentMap[id].examAvg = avg;
      }
    });
  }
  var list = [];
  for (var id in studentMap) {
    if (!isStudentActiveForSemesterKey(dropoutMap[id], examSuffix)) continue;
    var s = studentMap[id], mAvg = s.monthCount > 0 ? s.monthTotal / s.monthCount : null, finalAvg, totalSum;
    if (mAvg !== null && s.examAvg !== null) { finalAvg = (mAvg + s.examAvg) / 2; totalSum = mAvg + s.examAvg; }
    else if (mAvg !== null) { finalAvg = mAvg; totalSum = mAvg; }
    else if (s.examAvg !== null) { finalAvg = s.examAvg; totalSum = s.examAvg; }
    else { finalAvg = 0; totalSum = 0; }
    list.push({ id:s.id, name:s.name, sex:s.sex, total:totalSum.toFixed(2), avg:finalAvg.toFixed(2) });
  }
  return calculateRankAndGrade(list);
}

function getYearlyRankData(selectedClass) {
  var s1 = getWeightedSemesterRankData(selectedClass, SEMESTER1_MONTHS, 'ឆមាសទី១');
  var s2 = getWeightedSemesterRankData(selectedClass, SEMESTER2_MONTHS, 'ឆមាសទី២');
  if (!s1.length && !s2.length) return [];
  var map = {};
  s1.forEach(function(s){ map[s.id] = { id:s.id, name:s.name, sex:s.sex, avg1:parseFloat(s.avg) || 0 }; });
  s2.forEach(function(s){ if(!map[s.id]) map[s.id] = { id:s.id, name:s.name, sex:s.sex }; map[s.id].avg2 = parseFloat(s.avg) || 0; });
  var list = [];
  for (var id in map) {
    var s = map[id], vals = [];
    if (s.avg1 !== undefined) vals.push(s.avg1);
    if (s.avg2 !== undefined) vals.push(s.avg2);
    var sum = vals.reduce(function(a,b){ return a + b; }, 0);
    list.push({ id:s.id, name:s.name, sex:s.sex, total:sum.toFixed(2), avg:(vals.length ? sum / vals.length : 0).toFixed(2) });
  }
  return calculateRankAndGrade(list);
}

// ==========================================================================
// 📋 Roster & Dashboard
// ==========================================================================
function getStudentRosterByClass(selectedClass) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    if (!sheet || sheet.getLastRow() < 8) return [];
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, DROPOUT_STATUS_COL).getValues();
    var list = [], target = selectedClass ? selectedClass.toString().toLowerCase().replace(/\s+/g,'') : '';
    for (var i = 0; i < data.length; i++) {
      var rg = data[i][6] ? data[i][6].toString() : '', sg = rg.toLowerCase().replace(/\s+/g,'');
      if (!selectedClass || target === '' || target === 'all' || target === 'ទាំងអស់' || sg === target) {
        if (data[i][2] && data[i][2].toString().trim() !== '') {
          list.push({ id: data[i][1] ? data[i][1].toString().trim() : '', name: data[i][2].toString().trim(),
            sex: data[i][3] ? data[i][3].toString().trim() : '', dob: formatDobKhmer(data[i][4]), grade: rg.trim(),
            dropoutSemester: data[i][15] ? data[i][15].toString().trim() : '' });
        }
      }
    }
    return list;
  } catch(err) { throw new Error('កំហុស getStudentRosterByClass: ' + err.toString()); }
}

function getDashboardStats() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    var stats = { totalStudents:0, maleCount:0, femaleCount:0, totalClasses:0, classBreakdown:[], dropoutCount:0, activeCount:0 };
    if (!sheet || sheet.getLastRow() < 8) return stats;
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, DROPOUT_STATUS_COL).getValues();
    var classMap = {};
    data.forEach(function(row) {
      var name = row[2] ? row[2].toString().trim() : ''; if (!name) return;
      var sex = row[3] ? row[3].toString().trim() : '', grade = row[6] ? row[6].toString().trim() : 'គ្មានថ្នាក់', dropout = row[15] ? row[15].toString().trim() : '';
      stats.totalStudents++;
      if (sex.indexOf('ស្រី') !== -1) stats.femaleCount++; else stats.maleCount++;
      if (dropout) stats.dropoutCount++; else stats.activeCount++;
      if (!classMap[grade]) classMap[grade] = { className:grade, count:0, male:0, female:0 };
      classMap[grade].count++;
      if (sex.indexOf('ស្រី') !== -1) classMap[grade].female++; else classMap[grade].male++;
    });
    stats.classBreakdown = Object.keys(classMap).map(function(k){ return classMap[k]; }).sort(function(a,b){ return a.className.localeCompare(b.className,'en',{numeric:true}); });
    stats.totalClasses = stats.classBreakdown.length;
    return stats;
  } catch(err) { throw new Error('ស្ថិតិ: ' + err.toString()); }
}

// ==========================================================================
// 🏆 Top students / Certificate
// ==========================================================================
function getTopStudentsData(className, period, limit) {
  try {
    if (!className) return { success:false, message:'សូមជ្រើសរើសថ្នាក់!' };
    var lim = parseInt(limit) || 3;
    var rankList = getExamBasedRankData(className, period);
    if (!rankList || !rankList.length) return { success:false, message:'រកមិនឃើញទិន្នន័យ' };
    var top = rankList.slice(0, lim);
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sts = ss.getSheetByName('ព័ត៌មានសិស្ស');
    var sd = sts && sts.getLastRow() >= 8 ? sts.getRange(8, 1, sts.getLastRow() - 7, 15).getValues() : [];
    top.forEach(function(student) {
      student.studentName = student.name; student.gender = student.sex; student.class = className; student.imageUrl = ''; student.dob = '';
      for (var i = 0; i < sd.length; i++) {
        if ((sd[i][1] || '').toString().trim() === student.id.toString().trim()) {
          var cell = sts.getRange(i + 8, 15), rf = cell.getFormula(), pv = cell.getValue();
          if (rf && rf.indexOf('IMAGE("') !== -1) { var m = rf.match(/IMAGE\("([^"]+)"\)/i); if (m) student.imageUrl = m[1]; }
          else if (pv && typeof pv === 'string' && pv.substring(0, 4) === 'http') student.imageUrl = pv;
          student.dob = formatDobKhmer(sd[i][4]); break;
        }
      }
    });
    return { success:true, scores:top, classes:getClassList() };
  } catch(err) { return { success:false, message:err.toString() }; }
}

function getTopStudentsForHonorRoll(className, month) {
  var allRanks = getRankData(className, month);
  if (!allRanks || !allRanks.length) return { topStudents:[], lunarDate:'' };
  var top5 = allRanks.slice(0, 5);
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sts = ss.getSheetByName('ព័ត៌មានសិស្ស');
  var sd = sts && sts.getLastRow() >= 8 ? sts.getRange(8, 1, sts.getLastRow() - 7, 15).getValues() : [];
  top5.forEach(function(student) {
    student.imageUrl = '';
    for (var i = 0; i < sd.length; i++) {
      if ((sd[i][1] || '').toString().trim() === student.id.toString().trim()) {
        var cell = sts.getRange(i + 8, 15), rf = cell.getFormula(), pv = cell.getValue();
        if (rf && rf.indexOf('IMAGE("') !== -1) { var m = rf.match(/IMAGE\("([^"]+)"\)/i); if (m) student.imageUrl = m[1]; }
        else if (pv && typeof pv === 'string' && pv.substring(0, 4) === 'http') student.imageUrl = pv;
        break;
      }
    }
  });
  return { topStudents:top5, lunarDate:'' };
}

// ==========================================================================
// 📅 Attendance
// ==========================================================================
var ATT_DAY_COL_START = 5, ATT_SHEET_WIDTH = 38;

function getAttendanceSheetName(cls, month) { return 'វត្តមាន_ថ្នាក់ទី' + cls + '_' + month; }

function getMonthlyAttendanceForClass(className, month) {
  try {
    var students = getStudentsByClassForMonth(className, month);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getAttendanceSheetName(className, month));
    var savedMap = {};
    if (sheet && sheet.getLastRow() >= 8) {
      sheet.getRange(8, 1, sheet.getLastRow() - 7, ATT_SHEET_WIDTH).getValues().forEach(function(row) {
        var id = row[1] ? row[1].toString().trim() : ''; if (!id) return;
        var days = {};
        for (var d = 1; d <= 31; d++) {
          var val = row[3 + d];
          if (val === 'វត្តមាន') days[d] = 'P'; else if (val === 'ច្បាប់') days[d] = 'C'; else if (val === 'អវត្តមាន') days[d] = 'A'; else if (val === 'សម្រាក') days[d] = '-';
        }
        savedMap[id] = days;
      });
    }
    return students.map(function(s){ return { id:s.id, name:s.name, sex:s.sex, grade:s.grade, dropoutSemester:s.dropoutSemester || '', days:savedMap[s.id] || {} }; });
  } catch(err) { throw new Error('វត្តមាន: ' + err.toString()); }
}

function saveMonthlyAttendanceForClass(payload) {
  try {
    var cls = payload.className, month = payload.month, records = payload.records || [];
    if (!cls) throw new Error('សូមជ្រើសរើសថ្នាក់!');
    if (!month) throw new Error('សូមជ្រើសខែ!');
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sheetName = getAttendanceSheetName(cls, month);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var hb = sheet.getRange(1, 1, 6, ATT_SHEET_WIDTH); hb.breakApart(); hb.clearContent();
      sheet.getRange(1,1,1,ATT_SHEET_WIDTH).merge().setValue('ព្រះរាជាណាចក្រកម្ពុជា').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
      sheet.getRange(2,1,1,ATT_SHEET_WIDTH).merge().setValue('ជាតិ សាសនា ព្រះមហាក្សត្រ').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
      sheet.getRange('A3').setValue('មន្ទីរអប់រំ យុវជន និងកីឡា រាជធានី/ខេត្ត បាត់ដំបង').setFontWeight('bold');
      sheet.getRange('A4').setValue('ការិយាល័យអប់រំ យុវជន និងកីឡានៃរដ្ឋបាលក្រុង/ស្រុក/ខណ្ឌ សំឡូត').setFontWeight('bold');
      sheet.getRange('A5').setValue('សាលាបឋមសិក្សា ៖ កំពង់ល្ពៅ').setFontWeight('bold');
      sheet.getRange(6,1,1,ATT_SHEET_WIDTH).merge().setValue('តារាងវត្តមានប្រចាំខែ' + month + ' (ថ្នាក់ទី ' + cls + ')').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
      var headers = ['ល.រ','អត្តលេខ','ឈ្មោះសិស្ស','ភេទ'];
      for (var d = 1; d <= 31; d++) headers.push('ថ្ងៃទី ' + d);
      headers = headers.concat(['វត្តមានសរុប (P)','ច្បាប់សរុប (C)','អវត្តមានសរុប (A)']);
      sheet.getRange(7, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('white').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.setFrozenRows(7);
    }
    var existingMap = {};
    if (sheet.getLastRow() >= 8) {
      sheet.getRange(8, 2, sheet.getLastRow() - 7, 1).getValues().forEach(function(r, i) {
        var id = r[0] ? r[0].toString().trim() : ''; if (id) existingMap[id] = i + 8;
      });
    }
    var sw = { 'P':'វត្តមាន', 'C':'ច្បាប់', 'A':'អវត្តមាន', '-':'សម្រាក' };
    records.forEach(function(rec) {
      var dv = [], pc = 0, cc = 0, ac = 0;
      for (var d = 1; d <= 31; d++) {
        var code = (rec.days && rec.days[d]) ? rec.days[d] : '-', word = sw[code] || 'សម្រាក';
        dv.push(word);
        if (word === 'វត្តមាន') pc++; else if (word === 'ច្បាប់') cc++; else if (word === 'អវត្តមាន') ac++;
      }
      var rn = existingMap[rec.id];
      if (!rn) {
        rn = Math.max(sheet.getLastRow() + 1, 8);
        sheet.getRange(rn, 1, 1, 4 + 31 + 3).setValues([[rn - 7, rec.id, rec.name, rec.sex].concat(dv, [pc, cc, ac])]);
        existingMap[rec.id] = rn;
      } else {
        sheet.getRange(rn, 3).setValue(rec.name);
        sheet.getRange(rn, 4).setValue(rec.sex);
        sheet.getRange(rn, ATT_DAY_COL_START, 1, 31).setValues([dv]);
        sheet.getRange(rn, ATT_DAY_COL_START + 31, 1, 3).setValues([[pc, cc, ac]]);
      }
    });
    var fl = sheet.getLastRow();
    if (fl >= 8) {
      var iv = []; for (var i2 = 1; i2 <= fl - 7; i2++) iv.push([i2]);
      sheet.getRange(8, 1, fl - 7, 1).setValues(iv);
    }
    return 'បានរក្សាទុកវត្តមានជោគជ័យ!';
  } catch(err) { throw new Error('ការរក្សាទុកបរាជ័យ: ' + err.toString()); }
}

function getMonthlyAttendanceReport(className, month, year) {
  try {
    if (!className || className === 'ALL') throw new Error('សូមជ្រើសថ្នាក់ជាក់លាក់!');
    year = parseInt(year) || new Date().getFullYear();
    var monthIdx = KHMER_MONTH_NAMES.indexOf(month) + 1, daysCount = monthIdx > 0 ? new Date(year, monthIdx, 0).getDate() : 31;
    var dayLabels = []; for (var d = 1; d <= daysCount; d++) { var dow = monthIdx > 0 ? new Date(year, monthIdx - 1, d).getDay() : 0; dayLabels.push({ day:d, dow:dow, isWeekend:dow === 0 }); }
    var empty = { totalStudents:0, femaleCount:0, maleCount:0, totalPresent:0, totalPermission:0, totalAbsent:0, totalSlots:0, totalAbsenceAll:0, totalPresentActual:0, percentAbsent:'0.00', schoolDays:0 };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getAttendanceSheetName(className, month));
    if (!sheet || sheet.getLastRow() < 8) return { students:[], daysCount:daysCount, dayLabels:dayLabels, dailyPermission:[], dailyAbsent:[], dailyFemaleAbsence:[], summary:empty };
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, ATT_SHEET_WIDTH).getValues(), dropoutMap = getDropoutMap();
    var students = [], dP = new Array(daysCount).fill(0), dA = new Array(daysCount).fill(0), dF = new Array(daysCount).fill(0);
    var tP = 0, tC = 0, tA = 0, tS = 0, fC = 0, mC = 0;
    data.forEach(function(row) {
      var id = row[1] ? row[1].toString().trim() : ''; if (!id) return;
      if (!isStudentActiveForMonth(dropoutMap[id], month)) return;
      var name = row[2] ? row[2].toString().trim() : '', sex = row[3] ? row[3].toString().trim() : '', isF = sex.indexOf('ស្រី') !== -1;
      if (isF) fC++; else mC++;
      var days = [];
      for (var d = 1; d <= daysCount; d++) {
        var val = row[3 + d], code = '-';
        if (val === 'វត្តមាន') { code = 'P'; tP++; tS++; }
        else if (val === 'ច្បាប់') { code = 'C'; tC++; tS++; dP[d - 1]++; if (isF) dF[d - 1]++; }
        else if (val === 'អវត្តមាន') { code = 'A'; tA++; tS++; dA[d - 1]++; if (isF) dF[d - 1]++; }
        days.push(code);
      }
      students.push({ id:id, name:name, sex:sex, totalPresent:Number(row[35]) || 0, totalPermission:Number(row[36]) || 0, totalAbsent:Number(row[37]) || 0, days:days });
    });
    students.sort(function(a,b){ return a.id.localeCompare(b.id, 'en', { numeric:true }); });
    var absAll = tC + tA, presA = tS - absAll, pct = tS > 0 ? (absAll / tS * 100) : 0;
    return { students:students, daysCount:daysCount, dayLabels:dayLabels, dailyPermission:dP, dailyAbsent:dA, dailyFemaleAbsence:dF,
      summary:{ totalStudents:students.length, femaleCount:fC, maleCount:mC, totalPresent:tP, totalPermission:tC, totalAbsent:tA, totalSlots:tS, totalAbsenceAll:absAll, totalPresentActual:presA, percentAbsent:pct.toFixed(2), schoolDays:students.length > 0 ? Math.round(tS / students.length) : 0 } };
  } catch(err) { throw new Error('វត្តមានរបាយការណ៍: ' + err.toString()); }
}

// ==========================================================================
// 👤 Student CRUD
// ==========================================================================
function submitStudentData(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  if (!sheet) return 'Error: រកមិនឃើញ ព័ត៌មានសិស្ស';
  var lastRow = sheet.getLastRow(), newId = formData.studentId ? String(formData.studentId).trim() : '';
  if (lastRow >= 8 && newId !== '') {
    var eids = sheet.getRange(8, 2, lastRow - 7, 1).getValues();
    for (var i = 0; i < eids.length; i++) {
      if (eids[i][0] !== '' && String(eids[i][0]).trim() === newId) throw new Error('អត្តលេខ «' + newId + '» មានក្នុងប្រព័ន្ធរួចហើយ!');
    }
  }
  var nextRow = Math.max(lastRow + 1, 8), autoId = nextRow > 8 ? Number(sheet.getRange(nextRow - 1, 1).getValue()) + 1 : 1;
  var age = calculateAge(formData.dob), photoFormula = savePhoto(formData);
  sheet.getRange(nextRow, 1, 1, 15).setValues([[autoId, formData.studentId, formData.name, formData.gender || formData.sex, formData.dob, age, formData.grade || formData.className, formData.phone, formData.pob, formData.fatherName || formData.father, formData.fatherJob, formData.motherName || formData.mother, formData.motherJob, formData.address, photoFormula]]);
  sheet.setRowHeight(nextRow, 80); sortStudentsAndReindex(sheet);
  return 'រក្សាទុក និងតម្រៀបទិន្នន័យជោគជ័យ!';
}

function updateStudentData(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  var row = Number(formData.rowNum), lastRow = sheet.getLastRow(), upId = formData.studentId ? String(formData.studentId).trim() : '';
  if (lastRow >= 8 && upId !== '') {
    var eids = sheet.getRange(8, 2, lastRow - 7, 1).getValues();
    for (var i = 0; i < eids.length; i++) {
      if (i + 8 !== row && eids[i][0] !== '' && String(eids[i][0]).trim() === upId) throw new Error('អត្តលេខ «' + upId + '» ត្រូវបានប្រើដោយសិស្សផ្សេង!');
    }
  }
  var age = calculateAge(formData.dob);
  if (formData.studentId) sheet.getRange(row, 2).setValue(formData.studentId);
  if (formData.name) sheet.getRange(row, 3).setValue(formData.name);
  if (formData.gender || formData.sex) sheet.getRange(row, 4).setValue(formData.gender || formData.sex);
  if (formData.dob) sheet.getRange(row, 5).setValue(formData.dob);
  sheet.getRange(row, 6).setValue(age);
  if (formData.grade || formData.className) sheet.getRange(row, 7).setValue(formData.grade || formData.className);
  if (formData.phone) sheet.getRange(row, 8).setValue(formData.phone);
  if (formData.pob) sheet.getRange(row, 9).setValue(formData.pob);
  if (formData.fatherName || formData.father) sheet.getRange(row, 10).setValue(formData.fatherName || formData.father);
  if (formData.fatherJob) sheet.getRange(row, 11).setValue(formData.fatherJob);
  if (formData.motherName || formData.mother) sheet.getRange(row, 12).setValue(formData.motherName || formData.mother);
  if (formData.motherJob) sheet.getRange(row, 13).setValue(formData.motherJob);
  if (formData.address) sheet.getRange(row, 14).setValue(formData.address);
  if (formData.photoFile && (formData.photoFile.base64 || (formData.photoFile.bytes && formData.photoFile.bytes.length > 0))) {
    var pu = formData.photoFile.base64 ? uploadStudentPhoto(formData.photoFile.base64, 'សិស្ស_' + formData.studentId) : uploadStudentPhoto(formData.photoFile.bytes, 'សិស្ស_' + formData.studentId);
    sheet.getRange(row, 15).setValue('=IMAGE("' + pu + '")');
  } else if (formData.imageUrl && formData.imageUrl.substring(0, 4) === 'http') {
    sheet.getRange(row, 15).setValue('=IMAGE("' + formData.imageUrl + '")');
  }
  sortStudentsAndReindex(sheet);
  return 'ធ្វើបច្ចុប្បន្នភាពជោគជ័យ!';
}

function deleteStudentData(rowNum) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  sheet.deleteRow(Number(rowNum));
  var lr = sheet.getLastRow();
  if (lr >= 8) {
    var iv = []; for (var i = 1; i <= lr - 7; i++) iv.push([i]);
    sheet.getRange(8, 1, lr - 7, 1).setValues(iv);
  }
  return 'លុបទិន្នន័យសិស្សរួចរាល់!';
}

function sortStudentsAndReindex(sheet) {
  var lr = sheet.getLastRow();
  if (lr >= 8) {
    var n = lr - 7; sheet.getRange(8, 1, n, 17).sort({ column:3, ascending:true });
    var iv = []; for (var i = 1; i <= n; i++) iv.push([i]);
    sheet.getRange(8, 1, n, 1).setValues(iv);
  }
}

function getStudentDetailsByRow(rowNum) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    var row = Number(rowNum);
    var no = sheet.getRange(row, 1).getValue(), sid = sheet.getRange(row, 2).getValue(), name = sheet.getRange(row, 3).getValue();
    var sex = sheet.getRange(row, 4).getValue(), dob = sheet.getRange(row, 5).getValue(), cls = sheet.getRange(row, 7).getValue();
    var phone = sheet.getRange(row, 8).getValue(), pob = sheet.getRange(row, 9).getValue();
    var father = sheet.getRange(row, 10).getValue(), fj = sheet.getRange(row, 11).getValue();
    var mother = sheet.getRange(row, 12).getValue(), mj = sheet.getRange(row, 13).getValue(), addr = sheet.getRange(row, 14).getValue();
    var pcell = sheet.getRange(row, 15), rf = pcell.getFormula(), pv = pcell.getValue(), imgUrl = '';
    if (rf && rf.indexOf('IMAGE("') !== -1) { var m = rf.match(/IMAGE\("([^"]+)"\)/i); if (m) imgUrl = m[1]; }
    else if (pv && typeof pv === 'string' && pv.substring(0, 4) === 'http') imgUrl = pv;
    var dobFmt = dob instanceof Date ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd') : dob;
    return { rowNum:row, no:no ? no.toString().trim() : '', studentId:sid ? sid.toString().trim() : '', id:sid ? sid.toString().trim() : '', name:name ? name.toString().trim() : '', sex:sex ? sex.toString().trim() : 'ប្រុស', gender:sex ? sex.toString().trim() : 'ប្រុស', dob:dobFmt, className:cls ? cls.toString().trim() : '', grade:cls ? cls.toString().trim() : '', phone:phone ? phone.toString().trim() : '', pob:pob ? pob.toString().trim() : '', father:father ? father.toString().trim() : '', fatherName:father ? father.toString().trim() : '', fatherJob:fj ? fj.toString().trim() : '', mother:mother ? mother.toString().trim() : '', motherName:mother ? mother.toString().trim() : '', motherJob:mj ? mj.toString().trim() : '', address:addr ? addr.toString().trim() : '', imageUrl:imgUrl, photoUrl:imgUrl };
  } catch(e) { Logger.log('getStudentDetailsByRow: ' + e); return null; }
}

// ==========================================================================
// 📸 Photo helpers
// ==========================================================================
function uploadStudentPhoto(base64Data, fileName) {
  try {
    var ct = 'image/jpeg', rawBytes;
    if (typeof base64Data === 'string' && base64Data.indexOf(',') !== -1) { var sp = base64Data.split(','), m = sp[0].match(/:(.*?);/); if (m) ct = m[1]; rawBytes = Utilities.base64Decode(sp[1]); }
    else if (typeof base64Data === 'string') rawBytes = Utilities.base64Decode(base64Data);
    else rawBytes = base64Data;
    var blob = Utilities.newBlob(rawBytes, ct, fileName || 'photo');
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://lh3.googleusercontent.com/d/' + file.getId();
  } catch(e) { throw new Error('រូបភាព: ' + e.toString()); }
}

function savePhoto(formData) {
  if (formData.photoFile) {
    try {
      var url = '';
      if (formData.photoFile.base64) url = uploadStudentPhoto(formData.photoFile.base64, 'សិស្ស_' + (formData.studentId || 'photo'));
      else if (formData.photoFile.bytes && formData.photoFile.bytes.length > 0) {
        var b = Utilities.newBlob(formData.photoFile.bytes, formData.photoFile.mimeType, 'សិស្ស_' + (formData.name || 'photo'));
        var f = DriveApp.createFile(b); f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        url = 'https://lh3.googleusercontent.com/d/' + f.getId();
      }
      if (url) return '=IMAGE("' + url + '")';
    } catch(e) { return 'Error: ' + e.toString(); }
  }
  return '';
}

function savePhotoGeneric_(photoFile, fileNamePrefix) {
  if (photoFile && photoFile.base64) { try { var u = uploadStudentPhoto(photoFile.base64, fileNamePrefix); if (u) return '=IMAGE("' + u + '")'; } catch(e) { return ''; } }
  return '';
}

function extractImageUrl_(cell) {
  var rf = cell.getFormula(), val = cell.getValue();
  if (rf && rf.indexOf('IMAGE("') !== -1) { var m = rf.match(/IMAGE\("([^"]+)"\)/i); if (m) return m[1]; }
  else if (val && typeof val === 'string' && val.substring(0, 4) === 'http') return val;
  return '';
}

// ==========================================================================
// 🏫 Roles & Groups
// ==========================================================================
function getClassLeadership(className) {
  try {
    className = className ? className.toString().trim() : ''; if (!className) return null;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('តួនាទីប្រធានថ្នាក់');
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim() !== className) continue;
      function ms(k){ var key = k ? k.toString().trim() : ''; if (!key) return null; var students = getStudentsByClass(className); for (var j = 0; j < students.length; j++) { if ((students[j].id || students[j].name).toString().trim() === key || students[j].name.toString().trim() === key) return students[j]; } return { id:key, name:key, sex:'' }; }
      return { president:ms(data[i][1]), deputy1:ms(data[i][2]), deputy2:ms(data[i][3]) };
    }
    return null;
  } catch(err) { throw new Error('ចំណាត់ភ្ញៀវ: ' + err.message); }
}

function saveClassLeadership(payload) {
  try {
    if (!payload || !payload.className) throw new Error('សូមជ្រើសថ្នាក់!');
    var cls = payload.className.toString().trim(), p = payload.president || null, d1 = payload.deputy1 || null, d2 = payload.deputy2 || null;
    var ids = []; [p, d1, d2].forEach(function(s){ if (s) ids.push((s.id || s.name || '').toString().trim()); });
    if (ids.length !== ids.filter(function(v,i,a){ return a.indexOf(v) === i; }).length) throw new Error('សិស្សម្នាក់មិនអាចកាន់ ២ ​តួនាទី!');
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName('តួនាទីប្រធានថ្នាក់');
    if (!sheet) {
      sheet = ss.insertSheet('តួនាទីប្រធានថ្នាក់');
      sheet.getRange(1, 1, 1, 4).setValues([['ថ្នាក់','ប្រធានថ្នាក់','អនុប្រធានទី១','អនុប្រធានទី២']]).setFontWeight('bold').setBackground('#7c3aed').setFontColor('white').setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }
    var last = sheet.getLastRow(), found = -1;
    if (last >= 2) {
      var vals = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) { if ((vals[i][0] || '').toString().trim() === cls) { found = i + 2; break; } }
    }
    var row = [[cls, p ? (p.id || p.name || '') : '', d1 ? (d1.id || d1.name || '') : '', d2 ? (d2.id || d2.name || '') : '']];
    if (found > -1) sheet.getRange(found, 1, 1, 4).setValues(row); else sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues(row);
    sheet.autoResizeColumns(1, 4);
    return 'បានរក្សាទុកតួនាទីថ្នាក់ ' + cls + ' រួចរាល់!';
  } catch(err) { throw new Error('ការរក្សាទុករបរាជ័យ: ' + err.message); }
}

function saveGroupsToSheet(payload) {
  try {
    if (!payload || !payload.className) throw new Error('មិនបានកំណត់ថ្នាក់');
    var groups = payload.groups || []; if (!groups.length) throw new Error('មិនមានក្រុមត្រូវរក្សាទុក');
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sheetName = 'កាលវិភាគក្រុម_' + payload.className;
    var sheet = ss.getSheetByName(sheetName); if (!sheet) sheet = ss.insertSheet(sheetName); else sheet.clear();
    var days = ['ថ្ងៃចន្ទ','ថ្ងៃអង្គារ','ថ្ងៃពុធ','ថ្ងៃព្រហស្បតិ៍','ថ្ងៃសុក្រ','ថ្ងៃសៅរ៍'];
    var headers = ['ក្រុម','ថ្ងៃអនុវត្ត','ប្រធានប្រចាំថ្ងៃ','អនុប្រធានប្រចាំថ្ងៃ','ល.រ','អត្តលេខ','ឈ្មោះ','ភេទ'];
    sheet.getRange(1, 1, 1, headers.length).merge().setValue('កាលវិភាគ - ថ្នាក់ទី ' + payload.className).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    var cr = payload.classRoles || {};
    sheet.getRange(2, 1, 1, 6).setValues([['ប្រធានថ្នាក់', cr.president || '', 'អនុប្រធានទី១', cr.deputy1 || '', 'អនុប្រធានទី២', cr.deputy2 || '']]).setFontWeight('bold');
    sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#b91c1c').setFontColor('white').setHorizontalAlignment('center');
    var rowsData = [];
    groups.forEach(function(group, idx) {
      if (!group || !group.length) return;
      var dn = days[idx] || ('ក្រុមទី ' + (idx + 1)), dp = group[0] ? group[0].name : '', dd = group[1] ? group[1].name : '', gn = 'ក្រុមទី ' + (idx + 1);
      group.forEach(function(s, si) { rowsData.push([gn, dn, dp, dd, si + 1, s.id || '', s.name || '', s.sex || '']); });
    });
    if (rowsData.length) {
      sheet.getRange(5, 1, rowsData.length, headers.length).setValues(rowsData).setBorder(true,true,true,true,true,true).setVerticalAlignment('middle');
      sheet.getRange(5, 1, rowsData.length, 4).setHorizontalAlignment('center');
      sheet.autoResizeColumns(1, headers.length);
    }
    sheet.setFrozenRows(4);
    return 'បានរក្សាទុកកាលវិភាគក្រុមរួចរាល់!';
  } catch(err) { throw new Error('ការរក្សាទុកបរាជ័យ: ' + err.message); }
}

// ==========================================================================
// 📑 Report & Semester details
// ==========================================================================
function getStudentReportData(studentId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sts = ss.getSheetByName('ព័ត៌មានសិស្ស');
    if (!sts || sts.getLastRow() < 8) return { error:'រកមិនឃើញទិន្នន័យ' };
    var sd = sts.getRange(8, 1, sts.getLastRow() - 7, 15).getValues(), si = null, rn = 0;
    for (var i = 0; i < sd.length; i++) {
      if ((sd[i][1] || '').toString().trim() === studentId.toString().trim()) {
        rn = i + 8; var pc = sts.getRange(rn, 15), rf = pc.getFormula(), pv = pc.getValue(), imgUrl = '';
        if (rf && rf.indexOf('IMAGE("') !== -1) { var m = rf.match(/IMAGE\("([^"]+)"\)/i); if (m) imgUrl = m[1]; }
        else if (pv && typeof pv === 'string' && pv.substring(0, 4) === 'http') imgUrl = pv;
        si = { id:sd[i][1].toString().trim(), no:sd[i][0], name:sd[i][2], sex:sd[i][3], dob:sd[i][4] instanceof Date ? Utilities.formatDate(sd[i][4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : sd[i][4], grade:sd[i][6], phone:sd[i][7], father:sd[i][9], mother:sd[i][11], address:sd[i][13], imageUrl:imgUrl };
        break;
      }
    }
    if (!si) return { error:'រកមិនឃើញ «' + studentId + '»' };
    var scoresList = [];
    ss.getSheets().forEach(function(sh) {
      var sn = sh.getName(); if (sn.indexOf('ពិន្ទុ_ថ្នាក់ទី') !== 0) return;
      var mn = sn.split('_').pop(); if (sh.getLastRow() < 8) return;
      var isSem = (mn === 'ឆមាសទី១' || mn === 'ឆមាសទី២');
      var keys = isSem ? SEMESTER_KEYS : MONTHLY_KEYS;
      var w = 4 + keys.length + 3;
      sh.getRange(8, 1, sh.getLastRow() - 7, w).getValues().forEach(function(row) {
        if ((row[1] || '').toString().trim() !== studentId.toString().trim()) return;
        var entry = { month:mn, total:row[w - 3], avg:row[w - 2], rank:row[w - 1] };
        keys.forEach(function(k, idx){ entry[k] = row[4 + idx]; });
        scoresList.push(entry);
      });
    });
    var attList = [];
    ss.getSheets().forEach(function(sh) {
      var sn = sh.getName(); if (sn.indexOf('វត្តមាន_ថ្នាក់ទី') !== 0) return;
      var mn = sn.split('_').pop(); if (sh.getLastRow() < 8) return;
      sh.getRange(8, 1, sh.getLastRow() - 7, ATT_SHEET_WIDTH).getValues().forEach(function(row) {
        if ((row[1] || '').toString().trim() !== studentId.toString().trim()) return;
        attList.push({ month:mn, totalPresent:row[35], totalPermission:row[36], totalAbsent:row[37] });
      });
    });
    return { success:true, student:si, scores:scoresList, attendance:attList };
  } catch(err) { return { error:'ទាញទិន្នន័យ: ' + err.toString() }; }
}

// ==========================================================================
// 🔐 Auth / Session / Teacher management
// ==========================================================================
var TEACHER_SHEET_NAME = 'ព័ត៌មានគ្រូ', SESSION_DURATION_SEC = 21600, TEACHER_DATA_START_ROW = 3;

function getTeacherSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(TEACHER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TEACHER_SHEET_NAME);
    sheet.getRange(1, 1, 1, 12).merge().setValue('ព័ត៌មានគ្រូបង្រៀន').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
    sheet.appendRow(['ល.រ','អត្តលេខ','ឈ្មោះគ្រូ','ភេទ','លេខទូរសព្ទ','ថ្ងៃខែឆ្នាំ','ទីកន្លែង','Username','password','តួនាទី','ថ្នាក់','រូបថត']);
    sheet.getRange(2, 1, 1, 12).setFontWeight('bold').setBackground('#6d9eeb').setFontColor('white').setHorizontalAlignment('center');
  }
  return sheet;
}

function readAllTeachers_() {
  var sheet = getTeacherSheet_(), lr = sheet.getLastRow();
  if (lr < TEACHER_DATA_START_ROW) return [];
  var values = sheet.getRange(TEACHER_DATA_START_ROW, 1, lr - TEACHER_DATA_START_ROW + 1, 12).getValues(), list = [];
  values.forEach(function(row, idx) {
    if (!row[0] && !row[2]) return;
    var dob = row[5], dobIso = '', dobD = '';
    if (dob instanceof Date) { dobIso = Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd'); dobD = Utilities.formatDate(dob, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
    else { dobD = String(dob || ''); dobIso = dobD; }
    var rn = TEACHER_DATA_START_ROW + idx, pu = extractImageUrl_(sheet.getRange(rn, 12));
    list.push({ id:String(row[1] || row[0]), no:row[0], code:row[1], name:row[2], sex:row[3] || '', phone:row[4], dob:dobD, dobIso:dobIso, address:row[6], username:row[7], role:row[9] ? String(row[9]).trim().toLowerCase() : 'teacher', assignedClass:row[10] ? String(row[10]).trim() : '', photoUrl:pu, rowIndex:rn });
  });
  return list;
}

function createSessionToken_(teacher) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, JSON.stringify(teacher), SESSION_DURATION_SEC);
  return token;
}

function getSessionFromToken_(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get('session_' + token);
  return raw ? JSON.parse(raw) : null;
}

function checkLogin(username, password) {
  var sheet = getTeacherSheet_(), lr = sheet.getLastRow();
  if (lr < TEACHER_DATA_START_ROW) return null;
  var range = sheet.getRange(TEACHER_DATA_START_ROW, 1, lr - TEACHER_DATA_START_ROW + 1, 12).getValues();
  for (var i = 0; i < range.length; i++) {
    var row = range[i], su = row[7] ? String(row[7]).trim() : '', sp = row[8] ? String(row[8]).trim() : '';
    if (!su && !sp) continue;
    if (su === String(username).trim() && sp === String(password).trim()) {
      var rn = TEACHER_DATA_START_ROW + i;
      var teacher = { id:row[1] ? String(row[1]) : String(rn), code:row[1], name:row[2], sex:row[3] || '', username:su, role:row[9] ? String(row[9]).trim().toLowerCase() : 'teacher', assignedClass:row[10] ? String(row[10]).trim() : '', photoUrl:extractImageUrl_(sheet.getRange(rn, 12)) };
      return { token:createSessionToken_(teacher), teacher:teacher };
    }
  }
  return null;
}

function logoutSession(token) { if (token) CacheService.getScriptCache().remove('session_' + token); return true; }
function getSessionOrThrow_(token) { var s = getSessionFromToken_(token); if (!s) throw new Error('សូម Login ជាមុន (Session ផុតកំណត់)'); return s; }
function requireAdmin_(token) { var s = getSessionOrThrow_(token); if ((s.role || 'teacher') !== 'admin') throw new Error('Admin ប៉ុណ្ណោះ'); return s; }

function assertClassAccess_(session, className) {
  if ((session.role || 'teacher') === 'admin') return;
  var mine = (session.assignedClass || '').toString().trim().toLowerCase().replace(/\s+/g,'');
  var target = (className || '').toString().trim().toLowerCase().replace(/\s+/g,'');
  if (!mine || target !== mine) throw new Error('គ្មានសិទ្ធិថ្នាក់នេះ — អ្នកទទួលបន្ទុកថ្នាក់ ' + (session.assignedClass || '?'));
}

function getCurrentSession(token) {
  var s = getSessionFromToken_(token); if (!s) return { loggedIn:false };
  return { loggedIn:true, name:s.name || '', code:s.code || '', sex:s.sex || '', username:s.username || '', role:s.role || 'teacher', assignedClass:s.assignedClass || '', photoUrl:s.photoUrl || '' };
}

function getClassListForSession(token) {
  var s = getSessionFromToken_(token); if (!s) return [];
  if ((s.role || 'teacher') === 'admin') return getClassList();
  return s.assignedClass ? [s.assignedClass] : [];
}

function getClassList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  if (!sheet || sheet.getLastRow() < 8) return [];
  var data = sheet.getRange(8, 7, sheet.getLastRow() - 7, 1).getValues(), classes = [];
  for (var i = 0; i < data.length; i++) {
    var c = data[i][0] ? data[i][0].toString().trim() : '';
    if (c && classes.indexOf(c) === -1) classes.push(c);
  }
  return classes.sort();
}

function calculateAge(dobString) {
  if (!dobString) return '';
  var b = new Date(dobString), t = new Date(), age = t.getFullYear() - b.getFullYear(), m = t.getMonth() - b.getMonth();
  if (m < 0 || (m === 0 && t.getDate() < b.getDate())) age--;
  return age;
}

function formatDobKhmer(dateVal) {
  if (!dateVal) return '';
  var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return '';
  return d.getDate() + ' ខែ' + (KHMER_MONTH_NAMES[d.getMonth()] || '') + ' ឆ្នាំ ' + d.getFullYear();
}

// Secured Wrappers
function getStudentsForScoreEntrySecured(t, c, m) { var s = getSessionOrThrow_(t); return getStudentsForScoreEntry(s.role === 'admin' ? c : s.assignedClass, m); }
function saveMonthlyScoresSecured(t, p) { var s = getSessionOrThrow_(t); assertClassAccess_(s, p.className); return saveMonthlyScores(p); }
function getStudentsByClassSecured(t, c) { var s = getSessionOrThrow_(t); return getStudentsByClass(s.role === 'admin' ? c : s.assignedClass); }
function getStudentRosterByClassSecured(t, c) { var s = getSessionOrThrow_(t); return getStudentRosterByClass(s.role === 'admin' ? c : s.assignedClass); }
function getRankDataSecured(t, c, m) { var s = getSessionOrThrow_(t); return getRankData(s.role === 'admin' ? c : s.assignedClass, m); }
function getMonthlyAttendanceForClassSecured(t, c, m) { var s = getSessionOrThrow_(t); return getMonthlyAttendanceForClass(s.role === 'admin' ? c : s.assignedClass, m); }
function saveMonthlyAttendanceForClassSecured(t, p) { var s = getSessionOrThrow_(t); assertClassAccess_(s, p.className); return saveMonthlyAttendanceForClass(p); }
function submitStudentDataSecured(t, fd) { var s = getSessionOrThrow_(t); if (s.role !== 'admin') { assertClassAccess_(s, fd.grade || fd.className); fd.grade = s.assignedClass; fd.className = s.assignedClass; } return submitStudentData(fd); }
function updateStudentDataSecured(t, fd) { var s = getSessionOrThrow_(t); if (s.role !== 'admin') { var ex = getStudentDetailsByRow(fd.rowNum); if (ex) assertClassAccess_(s, ex.className || ex.grade); if (fd.grade || fd.className) assertClassAccess_(s, fd.grade || fd.className); fd.grade = s.assignedClass; fd.className = s.assignedClass; } return updateStudentData(fd); }
function deleteStudentDataSecured(t, rn) { var s = getSessionOrThrow_(t); if (s.role !== 'admin') { var ex = getStudentDetailsByRow(rn); if (ex) assertClassAccess_(s, ex.className || ex.grade); } return deleteStudentData(rn); }

// ==========================================================================
// 🌐 Web App entry point
// ==========================================================================
var PAGE_FILE_MAP = { roster:'RosterForm', addStudent:'Form', editStudent:'EditForm', attendance:'AttendancePage', scores:'ScoreForm', rank:'RankPage', honorRoll:'HonorRoll', certificate:'Certificate', groups:'GroupForm', wheel:'WheelForm', cards:'CardForm', report:'Report', semesterRecord:'SemesterRecord', teachers:'TeacherManagement' };

function getWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  if (!url) url = 'https://script.google.com/macros/s/' + ScriptApp.getScriptId() + '/exec';
  if (url && url.indexOf('/dev') !== -1) url = url.replace('/dev','/exec');
  return url;
}

function doGet(e) {
  var action = e && e.parameter.action ? e.parameter.action : '', studentId = e && e.parameter.id ? e.parameter.id : '', page = e && e.parameter.page ? e.parameter.page : '', token = e && e.parameter.token ? e.parameter.token : '';
  if (action === 'card') { var t = HtmlService.createTemplateFromFile('CardView'); t.studentId = studentId; return t.evaluate().setTitle('កាតសិស្ស').addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
  if (action === 'parentView' || (studentId && !page)) { var pv = HtmlService.createTemplateFromFile('ParentView'); pv.studentId = studentId; return pv.evaluate().setTitle('តាមដានការសិក្សា').addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
  var session = getSessionFromToken_(token);
  if (page === 'login' || !session) { var lt = HtmlService.createTemplateFromFile('Login'); lt.scriptUrl = getWebAppUrl(); return lt.evaluate().setTitle('ចូលប្រើប្រាស់ប្រព័ន្ធ').addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
  if (page && PAGE_FILE_MAP[page]) {
    var fn = PAGE_FILE_MAP[page], out;
    try { var st = HtmlService.createTemplateFromFile(fn); st.token = token; st.userRole = session.role || 'teacher'; st.userClass = session.assignedClass || ''; out = st.evaluate(); }
    catch (er) { out = HtmlService.createHtmlOutputFromFile(fn); }
    return out.addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var template = HtmlService.createTemplateFromFile('Index');
  template.scriptUrl = getWebAppUrl(); template.token = token; template.teacherName = session.name || ''; template.teacherPhotoUrl = session.photoUrl || ''; template.userRole = session.role || 'teacher'; template.userClass = session.assignedClass || '';
  return template.evaluate().setTitle('ប្រព័ន្ធគ្រប់គ្រងសិស្ស - សាលាបឋមសិក្សាកំពង់ល្ពៅ').addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
