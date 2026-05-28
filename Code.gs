const LOG_SHEET = 'Logg';
const COL_NAME  = 1;
const COL_LEFT  = 2;
const COL_RET   = 3;
const COL_EPOCH = 4;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Toalettbesök')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getClasses() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName())
    .filter(n => n !== LOG_SHEET);
}

function getStudents(className) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  if (!sheet) return { students: [], serverTime: Date.now() };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { students: [], serverTime: Date.now() };
  const data   = sheet.getRange(2, COL_NAME, lastRow - 1, 4).getValues();
  const visits = _visitCounts(className);
  const students = data
    .filter(r => r[0])
    .map(r => ({
      name:      String(r[0]),
      left:      r[1] ? String(r[1]) : '',
      returned:  r[2] ? String(r[2]) : '',
      leftEpoch: r[3] ? Number(r[3]) : 0,
      visits:    visits[String(r[0])] || 0
    }));
  return { students, serverTime: Date.now() };
}

function logLeave(className, studentName, testName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  const now   = new Date();
  const tz    = Session.getScriptTimeZone();
  const time  = Utilities.formatDate(now, tz, 'HH:mm');
  const epoch = now.getTime();
  const row   = _findRow(sheet, studentName);
  if (row < 0) return { success: false };
  sheet.getRange(row, COL_LEFT).setValue(time);
  sheet.getRange(row, COL_RET).setValue('');
  sheet.getRange(row, COL_EPOCH).setValue(epoch);
  _appendLog(className, testName, studentName, 'Gick', time, tz);
  return { success: true, time, epoch };
}

function logReturn(className, studentName, testName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  const now   = new Date();
  const tz    = Session.getScriptTimeZone();
  const time  = Utilities.formatDate(now, tz, 'HH:mm');
  const row   = _findRow(sheet, studentName);
  if (row < 0) return { success: false };
  sheet.getRange(row, COL_RET).setValue(time);
  sheet.getRange(row, COL_EPOCH).setValue('');
  _appendLog(className, testName, studentName, 'Tillbaka', time, tz);
  return { success: true, time };
}

function resetStudent(className, studentName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  const row   = _findRow(sheet, studentName);
  if (row < 0) return { success: false };
  sheet.getRange(row, COL_LEFT, 1, 3).clearContent();
  return { success: true };
}

function clearSession(className) {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, COL_LEFT, lastRow - 1, 3).clearContent();
  return { success: true };
}

function _findRow(sheet, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const names = sheet.getRange(2, COL_NAME, lastRow - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (names[i][0] === name) return i + 2;
  }
  return -1;
}

function _visitCounts(className) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log || log.getLastRow() < 2) return {};
  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data  = log.getRange(2, 1, log.getLastRow() - 1, 5).getValues();
  const counts = {};
  data.forEach(r => {
    const d = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
      : String(r[0]).slice(0, 10);
    if (d === today && String(r[1]) === className && String(r[4]) === 'Gick') {
      const n = String(r[3]);
      counts[n] = (counts[n] || 0) + 1;
    }
  });
  return counts;
}

function _appendLog(className, testName, studentName, type, time, tz) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    const hdr = log.getRange(1, 1, 1, 6);
    hdr.setValues([['Datum', 'Klass', 'Prov', 'Elev', 'Händelse', 'Tid']]);
    hdr.setFontWeight('bold');
  }
  const date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  log.appendRow([date, className, testName || '', studentName, type, time]);
}
