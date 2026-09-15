// ==========================================================================
// 🌐 API ROUTER សម្រាប់បម្រើទិន្នន័យទៅ GITHUB PAGES
// ==========================================================================
function doPost(e) {
  var output = { success: false };
  try {
    var req = JSON.parse(e.postData.contents);
    var action = req.action;
    var data = req.data || {};

    switch (action) {
      case 'checkLogin':
        output = checkLogin(data.username, data.password);
        break;
      case 'getCurrentSession':
        output = getCurrentSession(data.token);
        break;
      case 'logoutSession':
        output = { success: logoutSession(data.token) };
        break;
      case 'getMyProfile':
        output = getMyProfile(data.token);
        break;
      case 'updateMyProfile':
        output = updateMyProfile(data.token, data.profileData);
        break;
      case 'getDashboardStats':
        output = getDashboardStats();
        break;
      case 'getClassList':
        output = getClassListForSession(data.token);
        break;
      case 'getStudentsByClass':
        output = getStudentsByClassSecured(data.token, data.className);
        break;
      case 'getStudentRosterByClass':
        output = getStudentRosterByClassSecured(data.token, data.className);
        break;
      case 'getStudentsForScoreEntry':
        output = getStudentsForScoreEntrySecured(data.token, data.className, data.period);
        break;
      case 'saveMonthlyScores':
        output = { message: saveMonthlyScoresSecured(data.token, data.payload) };
        break;
      case 'getSavedScores':
        output = getSavedScores(data.className, data.period);
        break;
      case 'getRankData':
        output = getRankDataSecured(data.token, data.className, data.period);
        break;
      case 'getTopStudentsData':
        output = getTopStudentsDataSecured(data.token, data.className, data.period, data.limit);
        break;
      case 'getTopStudentsForHonorRoll':
        output = getTopStudentsForHonorRoll(data.className, data.month);
        break;
      case 'getMonthlyAttendanceForClass':
        output = getMonthlyAttendanceForClassSecured(data.token, data.className, data.month);
        break;
      case 'saveMonthlyAttendanceForClass':
        output = { message: saveMonthlyAttendanceForClassSecured(data.token, data.payload) };
        break;
      case 'getMonthlyAttendanceReport':
        output = getMonthlyAttendanceReport(data.className, data.month, data.year);
        break;
      case 'getYearlyAttendanceReport':
        output = getYearlyAttendanceReportSecured(data.token, data.className, data.year);
        break;
      case 'submitStudentData':
        output = { message: submitStudentDataSecured(data.token, data.formData) };
        break;
      case 'updateStudentData':
        output = { message: updateStudentDataSecured(data.token, data.formData) };
        break;
      case 'deleteStudentData':
        output = { message: deleteStudentDataSecured(data.token, data.rowNum) };
        break;
      case 'getStudentDetailsByRow':
        output = getStudentDetailsByRowSecured(data.token, data.rowNum);
        break;
      case 'getAllStudentsCardData':
        output = getAllStudentsCardDataSecured(data.token, data.classFilter);
        break;
      case 'getStudentsDropdownList':
        output = getStudentsDropdownListSecured(data.token);
        break;
      case 'getClassLeadership':
        output = getClassLeadership(data.className);
        break;
      case 'saveClassLeadership':
        output = { message: saveClassLeadership(data.payload) };
        break;
      case 'saveGroupsToSheet':
        output = { message: saveGroupsToSheet(data.payload) };
        break;
      case 'getReportBundle':
        output = getReportBundleSecured(data.token, data.className, data.month);
        break;
      case 'getSemesterRecordBundle':
        output = getSemesterRecordBundleSecured(data.token, data.className);
        break;
      case 'getSchoolHolidays':
        output = getSchoolHolidays();
        break;
      case 'saveSchoolHolidays':
        output = { success: saveSchoolHolidays(data.holidaysMap) };
        break;
      case 'getStudentAbsenceMap':
        output = getStudentAbsenceMapSecured(data.token, data.className, data.period);
        break;
      case 'setStudentDropout':
        output = { message: setStudentDropout(data.studentId, data.semester) };
        break;
      case 'clearAllSystemData':
        output = clearAllSystemData(data.token);
        break;
      default:
        output = { success: false, error: 'មិនស្គាល់ Action: ' + action };
    }
  } catch (err) {
    output = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================================================
// 📚 CONSTANTS & CORE LOGIC
// ==========================================================================
var MONTHLY_KEYS = [
  'kh_listen','kh_write','kh_read','kh_speak',
  'm_num','m_measure','m_geo','m_algebra','m_stat',
  'science','social',
  'a_craft','a_pe','a_health','a_env',
  'foreign'
];

var SEMESTER_KEYS = [
  'read','listen','dictation','essay',
  'math',
  'science','social',
  'homeArts','pe','ethics',
  'foreign'
];

var DROPOUT_STATUS_COL = 16;
var SEMESTER1_MONTHS   = ['វិច្ឆិកា','ធ្នូ','មករា','កុម្ភៈ','មីនា','មេសា'];
var SEMESTER2_MONTHS   = ['ឧសភា','មិថុនា','កក្កដា','សីហា','កញ្ញា','តុលា'];
var KHMER_MONTH_NAMES  = ['មករា','កុម្ភៈ','មីនា','មេសា','ឧសភា','មិថុនា','កក្កដា','សីហា','កញ្ញា','តុលា','វិច្ឆិកា','ធ្នូ'];

function getCleanClassName_(cls) {
  return (cls || '').toString().trim().replace(/^ថ្នាក់ទី\s*/, '').replace(/^ថ្នាក់\s*/, '');
}

function matchClass_(rawGrade, selectedClass) {
  if (!selectedClass) return true;
  var target = selectedClass.toString().toLowerCase().replace(/\s+/g,'').replace(/^(ថ្នាក់ទី|ថ្នាក់)/,'');
  if (target === '' || target === 'all' || target === 'ទាំងអស់') return true;
  var sg = (rawGrade || '').toString().toLowerCase().replace(/\s+/g,'').replace(/^(ថ្នាក់ទី|ថ្នាក់)/,'');
  return sg === target;
}

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

function getStudentsByClass(selectedClass) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    if (!sheet || sheet.getLastRow() < 8) return [];
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, 8).getValues();
    var list = [];
    for (var i = 0; i < data.length; i++) {
      var rawGrade = data[i][6] ? data[i][6].toString() : '';
      if (matchClass_(rawGrade, selectedClass)) {
        if (data[i][2] && data[i][2].toString().trim() !== '') {
          list.push({
            id: data[i][1] ? data[i][1].toString().trim() : '',
            name: data[i][2].toString().trim(),
            sex: data[i][3] ? data[i][3].toString().trim() : '',
            grade: rawGrade.trim()
          });
        }
      }
    }
    return list;
  } catch(err) { throw new Error('កំហុសក្នុងការទាញយកទិន្នន័យ៖ ' + err.toString()); }
}

function getStudentsByClassWithStatus(selectedClass) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    if (!sheet || sheet.getLastRow() < 8) return [];
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, DROPOUT_STATUS_COL).getValues();
    var list = [];
    for (var i = 0; i < data.length; i++) {
      var rawGrade = data[i][6] ? data[i][6].toString() : '';
      if (matchClass_(rawGrade, selectedClass)) {
        if (data[i][2] && data[i][2].toString().trim() !== '') {
          list.push({
            id: data[i][1] ? data[i][1].toString().trim() : '',
            name: data[i][2].toString().trim(),
            sex: data[i][3] ? data[i][3].toString().trim() : '',
            grade: rawGrade.trim(),
            dropoutSemester: data[i][15] ? data[i][15].toString().trim() : ''
          });
        }
      }
    }
    return list;
  } catch(err) { throw new Error('កំហុសក្នុងការទាញយកទិន្នន័យ៖ ' + err.toString()); }
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
  } catch(err) { throw new Error('កំហុស getStudentsForScoreEntry: ' + err.toString()); }
}

function getStudentRosterByClass(selectedClass) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
    if (!sheet || sheet.getLastRow() < 8) return [];
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, DROPOUT_STATUS_COL).getValues();
    var list = [];
    for (var i = 0; i < data.length; i++) {
      var rg = data[i][6] ? data[i][6].toString() : '';
      if (matchClass_(rg, selectedClass)) {
        if (data[i][2] && data[i][2].toString().trim() !== '') {
          list.push({
            id: data[i][1] ? data[i][1].toString().trim() : '',
            name: data[i][2].toString().trim(),
            sex: data[i][3] ? data[i][3].toString().trim() : '',
            dob: formatDobKhmer(data[i][4]),
            grade: rg.trim(),
            phone: data[i][7] ? data[i][7].toString().trim() : '',
            pob: data[i][8] ? data[i][8].toString().trim() : '',
            father: data[i][9] ? data[i][9].toString().trim() : '',
            fatherJob: data[i][10] ? data[i][10].toString().trim() : '',
            mother: data[i][11] ? data[i][11].toString().trim() : '',
            motherJob: data[i][12] ? data[i][12].toString().trim() : '',
            address: data[i][13] ? data[i][13].toString().trim() : '',
            dropoutSemester: data[i][15] ? data[i][15].toString().trim() : ''
          });
        }
      }
    }
    return list;
  } catch(err) { throw new Error('កំហុស getStudentRosterByClass: ' + err.toString()); }
}

function saveMonthlyScores(payload) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var cls  = getCleanClassName_(payload.className || 'ទូទៅ');
    var period = payload.month || payload.period || 'មិនបានកំណត់';
    var scoreType = (period === 'ឆមាសទី១' || period === 'ឆមាសទី២' || payload.scoreType === 'semester') ? 'semester' : 'monthly';
    var sheetName = 'ពិន្ទុ_ថ្នាក់ទី' + cls + '_' + period;

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName); else sheet.clear();

    var activeKeys = (scoreType === 'semester') ? SEMESTER_KEYS : MONTHLY_KEYS;
    var subjectLabels = payload.subjectLabels || activeKeys;
    var totalTableWidth = 4 + activeKeys.length + 3;

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

    var headers = ['ល.រ','អត្តលេខ','ឈ្មោះសិស្ស','ភេទ']
      .concat(subjectLabels, ['ពិន្ទុសរុប','មធ្យមភាគ','ចំណាត់ថ្នាក់']);

    sheet.getRange(7, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#facc15')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,true,true);
    sheet.setRowHeight(7, 100);

    if (payload.coefficients) {
      PropertiesService.getDocumentProperties()
        .setProperty('COEF_' + sheetName, JSON.stringify(payload.coefficients));
      PropertiesService.getDocumentProperties()
        .setProperty('KEYS_' + sheetName, JSON.stringify(activeKeys));
    }

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

function getSavedScores(className, period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleanCls = getCleanClassName_(className);
    var sheetName = 'ពិន្ទុ_ថ្នាក់ទី' + cleanCls + '_' + period;
    var sheet = ss.getSheetByName(sheetName);
    var isSem = (period === 'ឆមាសទី១' || period === 'ឆមាសទី២');
    var defaultKeys = isSem ? SEMESTER_KEYS : MONTHLY_KEYS;
    var result = { labels:[], scores:{}, coefficients:{}, scoreType: isSem ? 'semester' : 'monthly' };
    if (!sheet || sheet.getLastRow() < 8) return result;

    var props = PropertiesService.getDocumentProperties();
    var storedKeysRaw = props.getProperty('KEYS_' + sheetName);
    var keys = storedKeysRaw ? JSON.parse(storedKeysRaw) : defaultKeys;

    var storedCoefRaw = props.getProperty('COEF_' + sheetName);
    if (storedCoefRaw) {
      var storedCoef = JSON.parse(storedCoefRaw);
      keys.forEach(function(k) {
        var v = Number(storedCoef[k]);
        result.coefficients[k] = isNaN(v) ? 1 : v;
      });
    } else {
      keys.forEach(function(k) { result.coefficients[k] = 1; });
    }

    var numSubj = keys.length;
    var width = 4 + numSubj + 3;

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
    return { labels:[], scores:{}, coefficients:{}, scoreType:'monthly' };
  }
}

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
    else                    {s.grade='ខ្សោយ';    s.gradeCode='F';s.color='#f44336';}
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
    var cleanCls = getCleanClassName_(className);
    var sheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + cleanCls + '_' + month);
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
  } catch(err){ return []; }
}

function getExamOnlyRankList(className, semesterKey) {
  try {
    var cleanCls = getCleanClassName_(className);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + cleanCls + '_' + semesterKey);
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
  } catch(err){ return []; }
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
  var cleanCls = getCleanClassName_(className);
  monthList.forEach(function(m){
    var sheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + cleanCls + '_' + m);
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
  var examSheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + cleanCls + '_' + examSuffix);
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

function getTopStudentsData(className, period, limit) {
  try {
    if (!className) return { success:false, message:'សូមជ្រើសរើសថ្នាក់!' };
    var targetPeriod = period || 'ប្រចាំឆ្នាំ';
    var rankList = getExamBasedRankData(className, targetPeriod);
    if (!rankList || !rankList.length) return { success:false, message:'រកមិនឃើញទិន្នន័យ' };
    
    var top = rankList.filter(function(student) {
      return Number(student.rank) <= (limit || 5);
    });

    var ss = SpreadsheetApp.getActiveSpreadsheet(), sts = ss.getSheetByName('ព័ត៌មានសិស្ស');
    var sd = sts && sts.getLastRow() >= 8 ? sts.getRange(8, 1, sts.getLastRow() - 7, 15).getValues() : [];
    
    top.forEach(function(student) {
      student.studentName = student.name; 
      student.gender = student.sex; 
      student.class = className; 
      student.imageUrl = ''; 
      student.dob = '';
      for (var i = 0; i < sd.length; i++) {
        if ((sd[i][1] || '').toString().trim() === student.id.toString().trim()) {
          var cell = sts.getRange(i + 8, 15), rf = cell.getFormula(), pv = cell.getValue();
          if (rf && rf.indexOf('IMAGE("') !== -1) { 
            var m = rf.match(/IMAGE\("([^"]+)"\)/i); 
            if (m) student.imageUrl = m[1]; 
          } else if (pv && typeof pv === 'string' && pv.substring(0, 4) === 'http') {
            student.imageUrl = pv;
          }
          student.dob = formatDobKhmer(sd[i][4]); 
          break;
        }
      }
    });
    return { success:true, scores:top, classes:getClassList() };
  } catch(err) { 
    return { success:false, message:err.toString() }; 
  }
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
// 📅 ATTENDANCE
// ==========================================================================
var ATT_DAY_COL_START = 5;
var ATT_SHEET_WIDTH = 38;

function getAttendanceSheetName(cls, month) {
  var cleanCls = getCleanClassName_(cls);
  return 'វត្តមាន_ថ្នាក់ទី' + cleanCls + '_' + (month || '').toString().trim();
}

function getMonthlyAttendanceForClass(className, month) {
  try {
    var cleanCls = getCleanClassName_(className);
    var students = getStudentsByClassForMonth(cleanCls, month);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getAttendanceSheetName(cleanCls, month));
    var savedMap = {};
    
    if (sheet && sheet.getLastRow() >= 8) {
      var numRows = sheet.getLastRow() - 7;
      var data = sheet.getRange(8, 1, numRows, ATT_SHEET_WIDTH).getValues();
      
      data.forEach(function(row) {
        var id = row[1] ? row[1].toString().trim() : '';
        if (!id) return;
        
        var days = {};
        for (var d = 1; d <= 31; d++) {
          var val = row[4 + d - 1];
          if (val === 'វត្តមាន') days[d] = 'P';
          else if (val === 'ច្បាប់') days[d] = 'C';
          else if (val === 'អវត្តមាន') days[d] = 'A';
          else if (val === 'យឺត') days[d] = 'L';
          else if (val === 'សម្រាក') days[d] = '-';
          else days[d] = '-';
        }
        savedMap[id] = days;
      });
    }
    
    return students.map(function(s) {
      return {
        id: s.id,
        name: s.name,
        sex: s.sex,
        grade: s.grade,
        dropoutSemester: s.dropoutSemester || '',
        days: savedMap[s.id] || {}
      };
    });
  } catch(err) {
    throw new Error('កំហុសទាញវត្តមាន៖ ' + err.toString());
  }
}

function saveMonthlyAttendanceForClass(payload) {
  try {
    var rawCls = payload.className;
    var month = payload.month;
    var records = payload.records || [];
    
    if (!rawCls) throw new Error('សូមជ្រើសរើសថ្នាក់!');
    if (!month) throw new Error('សូមជ្រើសខែ!');
    
    var cls = getCleanClassName_(rawCls);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = getAttendanceSheetName(cls, month);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var hb = sheet.getRange(1, 1, 6, ATT_SHEET_WIDTH);
      hb.breakApart();
      hb.clearContent();
      
      sheet.getRange(1, 1, 1, ATT_SHEET_WIDTH).merge().setValue('ព្រះរាជាណាចក្រកម្ពុជា').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
      sheet.getRange(2, 1, 1, ATT_SHEET_WIDTH).merge().setValue('ជាតិ សាសនា ព្រះមហាក្សត្រ').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
      sheet.getRange('A3').setValue('មន្ទីរអប់រំ យុវជន និងកីឡា រាជធានី/ខេត្ត បាត់ដំបង').setFontWeight('bold');
      sheet.getRange('A4').setValue('ការិយាល័យអប់រំ យុវជន និងកីឡានៃរដ្ឋបាលក្រុង/ស្រុក/ខណ្ឌ សំឡូត').setFontWeight('bold');
      sheet.getRange('A5').setValue('សាលាបឋមសិក្សា ៖ កំពង់ល្ពៅ').setFontWeight('bold');
      sheet.getRange(6, 1, 1, ATT_SHEET_WIDTH).merge().setValue('តារាងវត្តមានប្រចាំខែ' + month + ' (ថ្នាក់ទី ' + cls + ')').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
      
      var headers = ['ល.រ', 'អត្តលេខ', 'ឈ្មោះសិស្ស', 'ភេទ'];
      for (var d = 1; d <= 31; d++) headers.push('ថ្ងៃទី ' + d);
      headers.push('វត្តមានសរុប (P)', 'ច្បាប់សរុប (C)', 'អវត្តមានសរុប (A)');
      
      sheet.getRange(7, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold')
        .setBackground('#1F4E78')
        .setFontColor('white')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      sheet.setFrozenRows(7);
    }
    
    var existingMap = {};
    if (sheet.getLastRow() >= 8) {
      var existingIds = sheet.getRange(8, 2, sheet.getLastRow() - 7, 1).getValues();
      for (var i = 0; i < existingIds.length; i++) {
        var sid = existingIds[i][0] ? existingIds[i][0].toString().trim() : '';
        if (sid) existingMap[sid] = i + 8;
      }
    }
    
    var sw = { 'P':'វត្តមាន', 'C':'ច្បាប់', 'A':'អវត្តមាន', 'L':'យឺត', '-':'សម្រាក' };
    
    records.forEach(function(rec) {
      var dv = [];
      var pc = 0, cc = 0, ac = 0;
      
      for (var d = 1; d <= 31; d++) {
        var code = (rec.days && rec.days[d]) ? rec.days[d] : '-';
        var word = sw[code] || 'សម្រាក';
        dv.push(word);
        
        if (word === 'វត្តមាន' || word === 'យឺត') pc++;
        else if (word === 'ច្បាប់') cc++;
        else if (word === 'អវត្តមាន') ac++;
      }
      
      var rn = existingMap[rec.id];
      if (!rn) {
        rn = Math.max(sheet.getLastRow() + 1, 8);
        var newRow = [rn - 7, rec.id, rec.name, rec.sex].concat(dv, [pc, cc, ac]);
        sheet.getRange(rn, 1, 1, ATT_SHEET_WIDTH).setValues([newRow]);
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
      var iv = [];
      for (var i2 = 1; i2 <= fl - 7; i2++) iv.push([i2]);
      sheet.getRange(8, 1, fl - 7, 1).setValues(iv);
      sheet.getRange(8, 1, fl - 7, ATT_SHEET_WIDTH).setBorder(true, true, true, true, true, true).setHorizontalAlignment('center');
      sheet.getRange(8, 3, fl - 7, 1).setHorizontalAlignment('left');
    }
    
    return 'បានរក្សាទុកវត្តមានជោគជ័យ!';
  } catch(err) {
    throw new Error('ការរក្សាទុកវត្តមានបរាជ័យ៖ ' + err.toString());
  }
}

function getMonthlyAttendanceReport(className, month, year) {
  try {
    if (!className || className === 'ALL') throw new Error('សូមជ្រើសថ្នាក់ជាក់លាក់!');
    year = parseInt(year) || new Date().getFullYear();
    var cleanCls = getCleanClassName_(className);
    var monthIdx = KHMER_MONTH_NAMES.indexOf(month) + 1;
    var daysCount = monthIdx > 0 ? new Date(year, monthIdx, 0).getDate() : 31;
    
    var dayLabels = [];
    for (var d = 1; d <= daysCount; d++) {
      var dow = monthIdx > 0 ? new Date(year, monthIdx - 1, d).getDay() : 0;
      dayLabels.push({ day: d, dow: dow, isWeekend: dow === 0 });
    }
    
    var empty = { totalStudents:0, femaleCount:0, maleCount:0, totalPresent:0, totalPermission:0, totalAbsent:0, totalSlots:0, totalAbsenceAll:0, totalPresentActual:0, percentAbsent:'0.00', schoolDays:0 };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getAttendanceSheetName(cleanCls, month));
    if (!sheet || sheet.getLastRow() < 8) return { students:[], daysCount:daysCount, dayLabels:dayLabels, dailyPermission:[], dailyAbsent:[], dailyFemaleAbsence:[], summary:empty };
    
    var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, ATT_SHEET_WIDTH).getValues();
    var dropoutMap = getDropoutMap();
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
        if (val === 'វត្តមាន' || val === 'យឺត') { code = 'P'; tP++; tS++; }
        else if (val === 'ច្បាប់') { code = 'C'; tC++; tS++; dP[d - 1]++; if (isF) dF[d - 1]++; }
        else if (val === 'អវត្តមាន') { code = 'A'; tA++; tS++; dA[d - 1]++; if (isF) dF[d - 1]++; }
        days.push(code);
      }
      students.push({ id:id, name:name, sex:sex, totalPresent:Number(row[35]) || 0, totalPermission:Number(row[36]) || 0, totalAbsent:Number(row[37]) || 0, days:days });
    });
    
    students.sort(function(a,b){ return a.id.localeCompare(b.id, 'en', { numeric:true }); });
    var absAll = tC + tA, presA = tS - absAll, pct = tS > 0 ? (absAll / tS * 100) : 0;
    
    return {
      students: students,
      daysCount: daysCount,
      dayLabels: dayLabels,
      dailyPermission: dP,
      dailyAbsent: dA,
      dailyFemaleAbsence: dF,
      summary: {
        totalStudents: students.length,
        femaleCount: fC,
        maleCount: mC,
        totalPresent: tP,
        totalPermission: tC,
        totalAbsent: tA,
        totalSlots: tS,
        totalAbsenceAll: absAll,
        totalPresentActual: presA,
        percentAbsent: pct.toFixed(2),
        schoolDays: students.length > 0 ? Math.round(tS / students.length) : 0
      }
    };
  } catch(err) {
    throw new Error('កំហុសរបាយការណ៍វត្តមាន៖ ' + err.toString());
  }
}

function getYearlyAttendanceReport(className, year) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleanCls = getCleanClassName_(className);
    var students = getStudentsByClass(cleanCls);
    var allMonths = SEMESTER1_MONTHS.concat(SEMESTER2_MONTHS);
    var monthDataMap = {};

    allMonths.forEach(function(m) {
      monthDataMap[m] = {};
      var sheet = ss.getSheetByName(getAttendanceSheetName(cleanCls, m));
      if (sheet && sheet.getLastRow() >= 8) {
        var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, ATT_SHEET_WIDTH).getValues();
        data.forEach(function(row) {
          var sid = row[1] ? row[1].toString().trim() : '';
          if (sid) {
            monthDataMap[m][sid] = {
              c: Number(row[36]) || 0,
              a: Number(row[37]) || 0
            };
          }
        });
      }
    });

    return {
      success: true,
      students: students,
      monthData: monthDataMap,
      sem1Months: SEMESTER1_MONTHS,
      sem2Months: SEMESTER2_MONTHS
    };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ==========================================================================
// 👤 STUDENT CRUD & HELPERS
// ==========================================================================
function submitStudentData(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  if (!sheet) throw new Error('រកមិនឃើញសន្លឹក «ព័ត៌មានសិស្ស» ទេ');
  var lastRow = sheet.getLastRow(), newId = formData.studentId ? String(formData.studentId).trim() : '';
  if (lastRow >= 8 && newId !== '') {
    var eids = sheet.getRange(8, 2, lastRow - 7, 1).getValues();
    for (var i = 0; i < eids.length; i++) {
      if (eids[i][0] !== '' && String(eids[i][0]).trim() === newId) throw new Error('អត្តលេខ «' + newId + '» មានក្នុងប្រព័ន្ធរួចហើយ!');
    }
  }
  var nextRow = Math.max(lastRow + 1, 8), autoId = nextRow > 8 ? Number(sheet.getRange(nextRow - 1, 1).getValue()) + 1 : 1;
  var age = calculateAge(formData.dob);
  var photoFormula = savePhoto(formData);
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
      if (eids[i][0] !== '' && String(eids[i][0]).trim() === upId && (i + 8) !== row) throw new Error('អត្តលេខ «' + upId + '» ត្រូវបានប្រើដោយសិស្សផ្សេង!');
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
  } catch(e) { return null; }
}

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
    } catch(e) { return ''; }
  }
  return '';
}

function extractImageUrl_(cell) {
  var rf = cell.getFormula(), val = cell.getValue();
  if (rf && rf.indexOf('IMAGE("') !== -1) { var m = rf.match(/IMAGE\("([^"]+)"\)/i); if (m) return m[1]; }
  else if (val && typeof val === 'string' && val.substring(0, 4) === 'http') return val;
  return '';
}

function getClassLeadership(className) {
  try {
    className = className ? className.toString().trim() : ''; if (!className) return null;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('តួនាទីប្រធានថ្នាក់');
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < data.length; i++) {
      if (!matchClass_(data[i][0], className)) continue;
      function ms(k){ var key = k ? k.toString().trim() : ''; if (!key) return null; var students = getStudentsByClass(className); for (var j = 0; j < students.length; j++) { if ((students[j].id || students[j].name).toString().trim() === key || students[j].name.toString().trim() === key) return students[j]; } return { id:key, name:key, sex:'' }; }
      return { president:ms(data[i][1]), deputy1:ms(data[i][2]), deputy2:ms(data[i][3]) };
    }
    return null;
  } catch(err) { throw new Error('កំហុសទាញយកតួនាទីថ្នាក់៖ ' + err.message); }
}

function saveClassLeadership(payload) {
  try {
    if (!payload || !payload.className) throw new Error('សូមជ្រើសថ្នាក់!');
    var cls = getCleanClassName_(payload.className.toString().trim()), p = payload.president || null, d1 = payload.deputy1 || null, d2 = payload.deputy2 || null;
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName('តួនាទីប្រធានថ្នាក់');
    if (!sheet) {
      sheet = ss.insertSheet('តួនាទីប្រធានថ្នាក់');
      sheet.getRange(1, 1, 1, 4).setValues([['ថ្នាក់','ប្រធានថ្នាក់','អនុប្រធានទី១','អនុប្រធានទី២']]).setFontWeight('bold').setBackground('#7c3aed').setFontColor('white').setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }
    var last = sheet.getLastRow(), found = -1;
    if (last >= 2) {
      var vals = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) { if (matchClass_(vals[i][0], cls)) { found = i + 2; break; } }
    }
    var row = [[cls, p ? (p.id || p.name || '') : '', d1 ? (d1.id || d1.name || '') : '', d2 ? (d2.id || d2.name || '') : '']];
    if (found > -1) sheet.getRange(found, 1, 1, 4).setValues(row); else sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues(row);
    return 'បានរក្សាទុកតួនាទីថ្នាក់ ' + cls + ' រួចរាល់!';
  } catch(err) { throw new Error('ការរក្សាទុករបរាជ័យ: ' + err.message); }
}

function saveGroupsToSheet(payload) {
  try {
    if (!payload || !payload.className) throw new Error('មិនបានកំណត់ថ្នាក់');
    var groups = payload.groups || []; if (!groups.length) throw new Error('មិនមានក្រុមត្រូវរក្សាទុក');
    var cls = getCleanClassName_(payload.className);
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sheetName = 'កាលវិភាគក្រុម_' + cls;
    var sheet = ss.getSheetByName(sheetName); if (!sheet) sheet = ss.insertSheet(sheetName); else sheet.clear();
    var days = ['ថ្ងៃចន្ទ','ថ្ងៃអង្គារ','ថ្ងៃពុធ','ថ្ងៃព្រហស្បតិ៍','ថ្ងៃសុក្រ','ថ្ងៃសៅរ៍'];
    var headers = ['ក្រុម','ថ្ងៃអនុវត្ត','ប្រធានប្រចាំថ្ងៃ','អនុប្រធានប្រចាំថ្ងៃ','ល.រ','អត្តលេខ','ឈ្មោះ','ភេទ'];
    sheet.getRange(1, 1, 1, headers.length).merge().setValue('កាលវិភាគ - ថ្នាក់ទី ' + cls).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
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
    }
    return 'បានរក្សាទុកកាលវិភាគក្រុមរួចរាល់!';
  } catch(err) { throw new Error('ការរក្សាទុកបរាជ័យ: ' + err.message); }
}

// ==========================================================================
// 🔐 AUTH & USER PROFILES
// ==========================================================================
var TEACHER_SHEET_NAME = 'ព័ត៌មានគ្រូ', TEACHER_DATA_START_ROW = 3;

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
  var teacherJson = JSON.stringify(teacher);
  try { PropertiesService.getScriptProperties().setProperty('session_' + token, teacherJson); } catch(e){}
  try { CacheService.getScriptCache().put('session_' + token, teacherJson, 21600); } catch(e){}
  return token;
}

function getSessionFromToken_(token) {
  if (!token) return null;
  try {
    var raw = CacheService.getScriptCache().get('session_' + token);
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  try {
    var persistentRaw = PropertiesService.getScriptProperties().getProperty('session_' + token);
    if (persistentRaw) {
      CacheService.getScriptCache().put('session_' + token, persistentRaw, 21600);
      return JSON.parse(persistentRaw);
    }
  } catch(e) {}
  return null;
}

function checkLogin(username, password) {
  try {
    var u = String(username || '').trim().toLowerCase();
    var p = String(password || '').trim();
    if (!u || !p) return { success: false, message: 'សូមបញ្ចូល Username និង Password!' };

    var sheet = getTeacherSheet_();
    var lr = sheet.getLastRow();
    if (lr < TEACHER_DATA_START_ROW) return { success: false, message: 'មិនទាន់មានគណនីគ្រូក្នុងប្រព័ន្ធឡើយ' };

    var values = sheet.getRange(TEACHER_DATA_START_ROW, 1, lr - TEACHER_DATA_START_ROW + 1, 12).getValues();
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var su = String(row[7] || '').trim().toLowerCase();
      var sp = String(row[8] || '').trim();
      if (su === u && sp === p) {
        var rn = TEACHER_DATA_START_ROW + i;
        var role = String(row[9] || 'teacher').trim().toLowerCase();
        var teacher = {
          id: row[1] ? String(row[1]) : String(rn),
          code: row[1] ? String(row[1]) : '',
          name: String(row[2] || ''),
          sex: String(row[3] || 'ប្រុស'),
          username: su,
          role: role === 'admin' ? 'admin' : 'teacher',
          assignedClass: String(row[10] || ''),
          photoUrl: extractImageUrl_(sheet.getRange(rn, 12))
        };
        var token = createSessionToken_(teacher);
        return { success: true, token: token, teacher: teacher, message: 'ចូលប្រើប្រាស់ជោគជ័យ' };
      }
    }
    return { success: false, message: 'Username ឬ Password មិនត្រឹមត្រូវទេ!' };
  } catch (err) {
    return { success: false, message: 'កំហុស៖ ' + err.message };
  }
}

function logoutSession(token) {
  if (token) {
    try { CacheService.getScriptCache().remove('session_' + token); } catch(e) {}
    try { PropertiesService.getScriptProperties().deleteProperty('session_' + token); } catch(e) {}
  }
  return true;
}

function getSessionOrThrow_(token) { var s = getSessionFromToken_(token); if (!s) throw new Error('Session ផុតកំណត់ សូម Login ឡើងវិញ'); return s; }
function requireAdmin_(token) { var s = getSessionOrThrow_(token); if ((s.role || 'teacher') !== 'admin') throw new Error('សិទ្ធិ Admin ប៉ុណ្ណោះ'); return s; }

function getCurrentSession(token) {
  var s = getSessionFromToken_(token); if (!s) return { loggedIn:false };
  return { loggedIn:true, name:s.name || '', code:s.code || '', sex:s.sex || '', username:s.username || '', role:s.role || 'teacher', assignedClass:s.assignedClass || '', photoUrl:s.photoUrl || '' };
}

function getClassListForSession(token) {
  var s = getSessionFromToken_(token);
  if (!s || (s.role || 'teacher') === 'admin') return getClassList();
  return s.assignedClass ? [s.assignedClass] : getClassList();
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
  var b = dobString instanceof Date ? dobString : new Date(dobString);
  if (isNaN(b.getTime())) return '';
  var t = new Date();
  var age = t.getFullYear() - b.getFullYear();
  var m = t.getMonth() - b.getMonth();
  if (m < 0 || (m === 0 && t.getDate() < b.getDate())) age--;
  return age;
}

function formatDobKhmer(dateVal) {
  if (!dateVal) return '';
  if (typeof dateVal === 'string' && dateVal.indexOf('ខែ') !== -1) return dateVal;
  var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return String(dateVal);
  return d.getDate() + ' ខែ' + (KHMER_MONTH_NAMES[d.getMonth()] || '') + ' ឆ្នាំ ' + d.getFullYear();
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

function clearAllSystemData(token) {
  requireAdmin_(token);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var clearedRows = 0, deletedSheets = 0;
  var stSheet = ss.getSheetByName('ព័ត៌មានសិស្ស');
  if (stSheet && stSheet.getLastRow() >= 8) {
    clearedRows += (stSheet.getLastRow() - 7);
    stSheet.getRange(8, 1, stSheet.getLastRow() - 7, 17).clearContent();
  }
  var sheets = ss.getSheets();
  sheets.forEach(function(sh) {
    var name = sh.getName();
    if (name.indexOf('ពិន្ទុ_ថ្នាក់ទី') === 0 || name.indexOf('វត្តមាន_ថ្នាក់ទី') === 0 || name.indexOf('កាលវិភាគក្រុម_') === 0) {
      ss.deleteSheet(sh);
      deletedSheets++;
    }
  });
  return { success: true, clearedRows: clearedRows, deletedSheets: deletedSheets };
}

function getMyProfile(token) {
  var s = getSessionOrThrow_(token);
  var teachers = readAllTeachers_();
  for (var i = 0; i < teachers.length; i++) {
    var t = teachers[i];
    if (t.username === s.username || (s.code && t.code === s.code) || t.id === s.id) return t;
  }
  return { id: s.id, name: s.name, sex: s.sex || 'ប្រុស', phone: '', dobIso: '', address: '', username: s.username, role: s.role, assignedClass: s.assignedClass, photoUrl: s.photoUrl };
}

function updateMyProfile(token, data) {
  var s = getSessionOrThrow_(token);
  var sheet = getTeacherSheet_(), lr = sheet.getLastRow();
  if (lr < TEACHER_DATA_START_ROW) throw new Error('រកមិនឃើញទិន្នន័យគ្រូទេ');
  var values = sheet.getRange(TEACHER_DATA_START_ROW, 1, lr - TEACHER_DATA_START_ROW + 1, 12).getValues(), foundRow = -1;
  for (var i = 0; i < values.length; i++) {
    var su = values[i][7] ? String(values[i][7]).trim() : '';
    if (su === s.username) { foundRow = TEACHER_DATA_START_ROW + i; break; }
  }
  if (foundRow === -1) throw new Error('រកមិនឃើញគណនីគ្រូនេះទេ');
  if (data.name) sheet.getRange(foundRow, 3).setValue(data.name);
  if (data.sex) sheet.getRange(foundRow, 4).setValue(data.sex);
  if (data.phone) sheet.getRange(foundRow, 5).setValue(data.phone);
  if (data.dob) sheet.getRange(foundRow, 6).setValue(data.dob);
  if (data.address) sheet.getRange(foundRow, 7).setValue(data.address);
  if (data.password && String(data.password).trim() !== '') sheet.getRange(foundRow, 9).setValue(String(data.password).trim());
  var photoUrl = s.photoUrl || '';
  if (data.photoFile && data.photoFile.base64) {
    var uploadedUrl = uploadStudentPhoto(data.photoFile.base64, 'គ្រូ_' + (data.name || s.username));
    sheet.getRange(foundRow, 12).setValue('=IMAGE("' + uploadedUrl + '")');
    photoUrl = uploadedUrl;
  }
  var updatedTeacher = { id: s.id, code: s.code, name: data.name || s.name, sex: data.sex || s.sex, username: s.username, role: s.role, assignedClass: s.assignedClass, photoUrl: photoUrl };
  var json = JSON.stringify(updatedTeacher);
  CacheService.getScriptCache().put('session_' + token, json, 21600);
  PropertiesService.getScriptProperties().setProperty('session_' + token, json);
  return { success: true, teacher: updatedTeacher };
}

// ==========================================================================
// 🛡️ SECURED WRAPPERS
// ==========================================================================
function getTopStudentsDataSecured(token, className, period, limit) {
  var s = getSessionFromToken_(token), cls = className;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getTopStudentsData(cls, period, limit);
}
function getStudentsForScoreEntrySecured(t, c, m) { 
  var s = getSessionFromToken_(t), cls = c;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getStudentsForScoreEntry(cls, m); 
}
function saveMonthlyScoresSecured(t, p) { return saveMonthlyScores(p); }
function getStudentsByClassSecured(t, c) { 
  var s = getSessionFromToken_(t), cls = c;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getStudentsByClass(cls); 
}
function getStudentRosterByClassSecured(t, c) { 
  var s = getSessionFromToken_(t), cls = c;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getStudentRosterByClass(cls); 
}
function getRankDataSecured(t, c, m) { return getRankData(c, m); }
function getMonthlyAttendanceForClassSecured(t, c, m) { return getMonthlyAttendanceForClass(c, m); }
function saveMonthlyAttendanceForClassSecured(t, p) { return saveMonthlyAttendanceForClass(p); }
function submitStudentDataSecured(t, fd) { return submitStudentData(fd); }
function updateStudentDataSecured(t, fd) { return updateStudentData(fd); }
function deleteStudentDataSecured(t, rn) { return deleteStudentData(rn); }
function getAllStudentsCardDataSecured(token, classFilter) {
  var s = getSessionFromToken_(token), cls = classFilter;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getAllStudentsCardData(cls);
}
function getStudentsDropdownListSecured(token) {
  var s = getSessionFromToken_(token), list = getStudentsDropdownList();
  if (!s || (s.role || 'teacher') === 'admin' || !s.assignedClass) return list;
  return list.filter(function (st) { return matchClass_(st.grade, s.assignedClass); });
}
function getStudentDetailsByRowSecured(token, rowNum) { return getStudentDetailsByRow(Number(rowNum)); }
function getYearlyAttendanceReportSecured(token, className, year) {
  var s = getSessionFromToken_(token), cls = className;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getYearlyAttendanceReport(cls, year);
}
function getStudentAbsenceMapSecured(token, className, period) {
  var s = getSessionFromToken_(token), cls = className;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getStudentAbsenceMap(cls, period);
}

function getStudentsDropdownList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  if (!sheet || sheet.getLastRow() < 8) return [];
  var data = sheet.getRange(8, 1, sheet.getLastRow() - 7, 7).getValues(), list = [];
  for (var i = 0; i < data.length; i++) {
    var name = data[i][2] ? data[i][2].toString().trim() : '';
    if (!name) continue;
    list.push({ rowNum: i + 8, id: data[i][1] ? data[i][1].toString().trim() : '', name: name, sex: data[i][3] ? data[i][3].toString().trim() : '', grade: data[i][6] ? data[i][6].toString().trim() : '' });
  }
  return list;
}

function getAllStudentsCardData(classFilter) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ព័ត៌មានសិស្ស');
  if (!sheet || sheet.getLastRow() < 8) return [];
  var n = sheet.getLastRow() - 7;
  var data = sheet.getRange(8, 1, n, 14).getValues();
  var photoFormulas = sheet.getRange(8, 15, n, 1).getFormulas();
  var photoValues = sheet.getRange(8, 15, n, 1).getValues(), list = [];
  for (var i = 0; i < n; i++) {
    var name = data[i][2] ? data[i][2].toString().trim() : '';
    if (!name) continue;
    var grade = data[i][6] ? data[i][6].toString() : '';
    if (!matchClass_(grade, classFilter)) continue;
    var rf = photoFormulas[i][0], pv = photoValues[i][0], imgUrl = '';
    if (rf && rf.indexOf('IMAGE("') !== -1) { var m = rf.match(/IMAGE\("([^"]+)"\)/i); if (m) imgUrl = m[1]; }
    else if (pv && typeof pv === 'string' && pv.substring(0, 4) === 'http') imgUrl = pv;
    list.push({ rowNum: i + 8, id: data[i][1] ? data[i][1].toString().trim() : '', name: name, sex: data[i][3] ? data[i][3].toString().trim() : '', dob: formatDobKhmer(data[i][4]), grade: grade, pob: data[i][8] ? data[i][8].toString().trim() : '', father: data[i][9] ? data[i][9].toString().trim() : '', mother: data[i][11] ? data[i][11].toString().trim() : '', imageUrl: imgUrl });
  }
  return list.sort(function (a, b) { return a.grade.localeCompare(b.grade, 'en', { numeric: true }) || a.name.localeCompare(b.name); });
}

function getReportBundle(className, month) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), cleanCls = getCleanClassName_(className);
    var stSheet = ss.getSheetByName('ព័ត៌មានសិស្ស'), students = [];
    if (stSheet && stSheet.getLastRow() >= 8) {
      var n = stSheet.getLastRow() - 7, stData = stSheet.getRange(8, 1, n, DROPOUT_STATUS_COL).getValues();
      for (var i = 0; i < n; i++) {
        var name = stData[i][2] ? stData[i][2].toString().trim() : ''; if (!name) continue;
        var rawGrade = stData[i][6] ? stData[i][6].toString() : '';
        if (matchClass_(rawGrade, cleanCls)) {
          var dropout = stData[i][15] ? stData[i][15].toString().trim() : '';
          if (!isStudentActiveForMonth(dropout, month)) continue;
          students.push({ id: stData[i][1] ? stData[i][1].toString().trim() : '', name: name, sex: stData[i][3] ? stData[i][3].toString().trim() : '', dob: formatDobKhmer(stData[i][4]), grade: rawGrade.trim(), phone: stData[i][7] ? stData[i][7].toString().trim() : '', pob: stData[i][8] ? stData[i][8].toString().trim() : '', father: stData[i][9] ? stData[i][9].toString().trim() : '', fatherJob: stData[i][10] ? stData[i][10].toString().trim() : '', mother: stData[i][11] ? stData[i][11].toString().trim() : '', motherJob: stData[i][12] ? stData[i][12].toString().trim() : '', address: stData[i][13] ? stData[i][13].toString().trim() : '' });
        }
      }
    }
    var scoreSheet = ss.getSheetByName('ពិន្ទុ_ថ្នាក់ទី' + cleanCls + '_' + month), scoresMap = {};
    if (scoreSheet && scoreSheet.getLastRow() >= 8) {
      var isSem = (month === 'ឆមាសទី១' || month === 'ឆមាសទី២'), keys = isSem ? SEMESTER_KEYS : MONTHLY_KEYS, width = 4 + keys.length + 3;
      scoreSheet.getRange(8, 1, scoreSheet.getLastRow() - 7, width).getValues().forEach(function(row) {
        var sid = row[1] ? row[1].toString().trim() : ''; if (!sid) return;
        var entry = {}; keys.forEach(function(k, idx) { var v = row[4 + idx]; entry[k] = (v === '' || v === null || v === undefined) ? '' : v; });
        entry.total = row[width - 3]; entry.avg = row[width - 2]; scoresMap[sid] = entry;
      });
    }
    var attSheet = ss.getSheetByName(getAttendanceSheetName(cleanCls, month)), attMap = {};
    if (attSheet && attSheet.getLastRow() >= 8) {
      attSheet.getRange(8, 1, attSheet.getLastRow() - 7, ATT_SHEET_WIDTH).getValues().forEach(function(row) {
        var sid = row[1] ? row[1].toString().trim() : '';
        if (sid) attMap[sid] = { totalPresent: Number(row[35]) || 0, totalPermission: Number(row[36]) || 0, totalAbsent: Number(row[37]) || 0 };
      });
    }
    return { success: true, students: students, scores: scoresMap, attendance: attMap };
  } catch(err) { return { success: false, error: err.toString() }; }
}

function getReportBundleSecured(token, className, month) {
  var s = getSessionFromToken_(token), cls = className;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getReportBundle(cls, month);
}

function getSemesterRecordBundle(className) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!className) { var allCls = getClassList(); className = allCls.length > 0 ? allCls[0] : ''; }
    if (!className) return { success: false, error: 'រកមិនឃើញថ្នាក់ឡើយ' };

    var cleanCls = getCleanClassName_(className), students = getStudentRosterByClass(cleanCls);
    var sem1Scores = getSavedScores(cleanCls, 'ឆមាសទី១'), sem2Scores = getSavedScores(cleanCls, 'ឆមាសទី២');
    var s1Map = sem1Scores.scores || {}, s2Map = sem2Scores.scores || {};
    var sem1Rank = getWeightedSemesterRankData(cleanCls, SEMESTER1_MONTHS, 'ឆមាសទី១') || [];
    var sem2Rank = getWeightedSemesterRankData(cleanCls, SEMESTER2_MONTHS, 'ឆមាសទី២') || [];
    var yearRank = getYearlyRankData(cleanCls) || [];
    
    var sem1RankMap = {}, sem2RankMap = {}, yearRankMap = {};
    sem1Rank.forEach(function(s){ if(s && s.id) sem1RankMap[s.id] = s; });
    sem2Rank.forEach(function(s){ if(s && s.id) sem2RankMap[s.id] = s; });
    yearRank.forEach(function(s){ if(s && s.id) yearRankMap[s.id] = s; });

    return {
      success: true, className: cleanCls, classes: getClassList(), students: students,
      sem1Scores: s1Map, sem2Scores: s2Map, sem1RankMap: sem1RankMap, sem2RankMap: sem2RankMap,
      yearRankMap: yearRankMap, schoolName: 'កំពង់ល្ពៅ'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function getSemesterRecordBundleSecured(token, className) {
  var s = getSessionFromToken_(token), cls = className;
  if (s && s.role !== 'admin' && s.assignedClass) cls = s.assignedClass;
  return getSemesterRecordBundle(cls);
}

function getSchoolHolidays() {
  try { var raw = PropertiesService.getScriptProperties().getProperty('SCHOOL_HOLIDAYS'); return raw ? JSON.parse(raw) : null; } catch(e) { return null; }
}
function saveSchoolHolidays(holidaysMap) {
  try { PropertiesService.getScriptProperties().setProperty('SCHOOL_HOLIDAYS', JSON.stringify(holidaysMap)); return true; } catch(e) { return false; }
}
function getStudentAbsenceMap(className, period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), cleanCls = getCleanClassName_(className), absMap = {};
    var months = (period === 'ឆមាសទី១') ? SEMESTER1_MONTHS : (period === 'ឆមាសទី២') ? SEMESTER2_MONTHS : [period];
    months.forEach(function(m) {
      var sheet = ss.getSheetByName(getAttendanceSheetName(cleanCls, m));
      if (sheet && sheet.getLastRow() >= 8) {
        sheet.getRange(8, 1, sheet.getLastRow() - 7, ATT_SHEET_WIDTH).getValues().forEach(function(row) {
          var sid = row[1] ? row[1].toString().trim() : '';
          if (sid) absMap[sid] = (absMap[sid] || 0) + (Number(row[36])||0) + (Number(row[37])||0);
        });
      }
    });
    return absMap;
  } catch(e) { return {}; }
}
