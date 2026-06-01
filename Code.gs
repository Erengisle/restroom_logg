const LOG_SHEET  = 'Logg';
const COL_NAME   = 1;
const COL_LEFT   = 2;
const COL_RET    = 3;
const COL_EPOCH  = 4;
const QUEUE_PFX  = 'q_';
const MAXOUT_PFX = 'm_';
const DEFAULT_MAX = 2;

function doGet(e) {
  const view = e && e.parameter && e.parameter.view;

  if (view === 'manifest') {
    let startUrl = '';
    try { startUrl = ScriptApp.getService().getUrl(); } catch(_) {}
    const manifest = {
      name: 'Toalettbesök',
      short_name: 'Toalettbesök',
      start_url: startUrl,
      display: 'standalone',
      background_color: '#f8f9fa',
      theme_color: '#1a73e8',
      icons: [{
        src: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' rx='20' fill='%231a73e8'/%3E%3Ctext y='.9em' font-size='80' x='10'%3E%F0%9F%9A%BB%3C/text%3E%3C/svg%3E",
        sizes: 'any',
        type: 'image/svg+xml'
      }]
    };
    return ContentService.createTextOutput(JSON.stringify(manifest))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (view === 'elev') {
    const tmpl = HtmlService.createTemplateFromFile('Elev');
    tmpl.presetKlass = (e.parameter.klass) || '';
    return tmpl.evaluate()
      .setTitle('Toalettbesök – Elev')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const tmpl = HtmlService.createTemplateFromFile('Index');
  try { tmpl.webAppUrl = ScriptApp.getService().getUrl(); }
  catch (_) { tmpl.webAppUrl = ''; }
  return tmpl.evaluate()
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

function logLeave(className, studentName, testName, reason) {
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
  _appendLog(className, testName, studentName, 'Gick', time, tz, null, reason || 'Toalettbesök');
  return { success: true, time, epoch };
}

function logReturn(className, studentName, testName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  const now   = new Date();
  const tz    = Session.getScriptTimeZone();
  const time  = Utilities.formatDate(now, tz, 'HH:mm');
  const row   = _findRow(sheet, studentName);
  if (row < 0) return { success: false };
  const epochVal    = sheet.getRange(row, COL_EPOCH).getValue();
  const durationMs  = epochVal ? (now.getTime() - Number(epochVal)) : 0;
  const durationMin = durationMs > 0 ? Math.round(durationMs / 6000) / 10 : null;
  sheet.getRange(row, COL_RET).setValue(time);
  sheet.getRange(row, COL_EPOCH).setValue('');
  _appendLog(className, testName, studentName, 'Tillbaka', time, tz, durationMin, null);
  return { success: true, time, durationMin };
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
  _saveQueue(className, []);
  return { success: true };
}

// ── Kö & inställningar ────────────────────────────

function getMaxOut(className) {
  const v = PropertiesService.getScriptProperties().getProperty(MAXOUT_PFX + className);
  return v ? parseInt(v) : DEFAULT_MAX;
}

function setMaxOut(className, max) {
  PropertiesService.getScriptProperties()
    .setProperty(MAXOUT_PFX + className, String(parseInt(max)));
  return { success: true };
}

function _getQueue(className) {
  const raw = PropertiesService.getScriptProperties().getProperty(QUEUE_PFX + className);
  return raw ? JSON.parse(raw) : [];
}

function _saveQueue(className, queue) {
  PropertiesService.getScriptProperties()
    .setProperty(QUEUE_PFX + className, JSON.stringify(queue));
}

function getPendingRequests(className) {
  return { queue: _getQueue(className), maxOut: getMaxOut(className) };
}

function submitRequest(className, studentName, reason) {
  const queue    = _getQueue(className);
  const existing = queue.find(r => r.name === studentName);
  if (existing && existing.status === 'waiting') return { success: false, reason: 'already_queued' };
  const data     = getStudents(className);
  const student  = data.students.find(s => s.name === studentName);
  if (student && student.left && !student.returned) return { success: false, reason: 'already_out' };
  const outCount = data.students.filter(s => s.left && !s.returned).length;
  if (outCount >= getMaxOut(className)) return { success: false, reason: 'max_reached' };
  const clean = queue.filter(r => r.name !== studentName);
  clean.push({ name: studentName, time: Date.now(), status: 'waiting', reason: reason || 'Toalettbesök' });
  _saveQueue(className, clean);
  return { success: true };
}

function approveRequest(className, studentName, testName) {
  const data     = getStudents(className);
  const outCount = data.students.filter(s => s.left && !s.returned).length;
  if (outCount >= getMaxOut(className)) return { success: false, reason: 'max_reached' };
  const queue  = _getQueue(className);
  const entry  = queue.find(r => r.name === studentName);
  const reason = entry ? (entry.reason || 'Toalettbesök') : 'Toalettbesök';
  _saveQueue(className, queue.filter(r => r.name !== studentName));
  return logLeave(className, studentName, testName, reason);
}

function denyRequest(className, studentName) {
  const queue = _getQueue(className);
  const idx   = queue.findIndex(r => r.name === studentName);
  if (idx >= 0) queue[idx].status = 'denied';
  _saveQueue(className, queue);
  return { success: true };
}

function cancelRequest(className, studentName) {
  _saveQueue(className, _getQueue(className).filter(r => r.name !== studentName));
  return { success: true };
}

function getStudentStatus(className, studentName) {
  const queue    = _getQueue(className);
  const entry    = queue.find(r => r.name === studentName) || null;
  const data     = getStudents(className);
  const student  = data.students.find(s => s.name === studentName) || null;
  const outCount = data.students.filter(s => s.left && !s.returned).length;
  return { entry, student, maxOut: getMaxOut(className), outCount, serverTime: data.serverTime };
}

// ── Statistik ─────────────────────────────────────

function getStats(className, weeks) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log || log.getLastRow() < 2) return { rows: [] };
  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  let cutoffStr = today;
  if (weeks > 0) {
    const cutoff = new Date(new Date().getTime() - weeks * 7 * 24 * 60 * 60 * 1000);
    cutoffStr = Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd');
  }
  const numCols = Math.max(log.getLastColumn(), 8);
  const data    = log.getRange(2, 1, log.getLastRow() - 1, numCols).getValues();
  const rows    = [];
  data.forEach(r => {
    const d = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
      : String(r[0]).slice(0, 10);
    const inRange = weeks === 0 ? (d === today) : (d >= cutoffStr && d <= today);
    if (!inRange || String(r[1]) !== className) return;
    rows.push({
      date:     d,
      student:  String(r[3]),
      event:    String(r[4]),
      time:     String(r[5]),
      duration: (r[6] !== '' && r[6] != null) ? Number(r[6]) : null,
      reason:   String(r[7] || '')
    });
  });
  return { rows };
}

// ── Privata hjälpfunktioner ───────────────────────

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

function _appendLog(className, testName, studentName, type, time, tz, durationMin, reason) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    const hdr = log.getRange(1, 1, 1, 8);
    hdr.setValues([['Datum', 'Klass', 'Prov/Lektion', 'Elev', 'Händelse', 'Tid', 'Minuter', 'Orsak']]);
    hdr.setFontWeight('bold');
  }
  const date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  log.appendRow([
    date,
    className,
    testName || 'Lektion',
    studentName,
    type,
    time,
    durationMin != null ? durationMin : '',
    reason || ''
  ]);
}
