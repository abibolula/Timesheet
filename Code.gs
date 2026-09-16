/**
 * DAILY TIMEKEEPER — Google Apps Script (V8)
 * 1. Set SHEET_ID from the /d/…/ part of your Google Sheets URL.
 * 2. Set SHEET_NAME, timezone and task rules to match index.html CONFIG.
 * 3. Deploy > New deployment > Web app; execute as yourself; access Anyone.
 *    Authorize Sheets access. Paste the /exec URL into index.html CONFIG.
 * 4. After edits: Deploy > Manage deployments > Edit > New version.
 *
 * This date-only schema represents ONE employee per sheet/deployment.
 * employeeName is a display label, not authentication. A public endpoint is
 * callable by anyone who has its URL. For confidential multi-user attendance,
 * add authenticated server-side identity and an employee column before use.
 * Do not put passwords, private keys or a supposed secret token in the HTML.
 * Existing date cells may be Sheets dates or ISO text; ambiguous locale text
 * is rejected. Keep sheet timezone aligned with TIMEZONE. Overnight shifts
 * are not supported by this same-date schema.
 * Sources: https://developers.google.com/apps-script/guides/web
 * https://developers.google.com/apps-script/guides/content
 * https://developers.google.com/apps-script/reference/lock/lock-service
 */
const CONNECTION_PROPERTY = 'TIMEKEEPER_SPREADSHEET_ID';
const SHEET_NAME = 'Timesheet';
const TIMEZONE = 'Asia/Jakarta';
const TASK_REQUIRED = true;
const MAX_TASK_LENGTH = 1000;
const HISTORY_BATCH_SIZE = 300;
const HEADERS = ['Date', 'Clock In', 'Clock Out', 'Task', 'Last Updated'];

/** Run ONCE from the editor attached to the target Google Sheet.
 * getActiveSpreadsheet is unavailable in Web App execution, so the detected
 * ID is saved server-side in Script Properties, never embedded in source/HTML.
 * https://developers.google.com/apps-script/guides/bound#special_methods
 * This function is not exposed through doGet/doPost.
 */
function setupTimekeeper() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Buka Google Sheet > Extensions > Apps Script, lalu jalankan setupTimekeeper dari editor tersebut.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Database sedang digunakan. Coba setup kembali sebentar lagi.');
  try {
    PropertiesService.getScriptProperties().setProperty(CONNECTION_PROPERTY, book.getId());
    console.log('Timekeeper berhasil dihubungkan ke spreadsheet: ' + book.getName());
    return 'Koneksi tersimpan. Deploy Web App, lalu tempel URL /exec di Pengaturan Timekeeper.';
  } finally { lock.releaseLock(); }
}
function openTimekeeperSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(CONNECTION_PROPERTY);
  if (!id) fail_('NOT_CONFIGURED','Jalankan setupTimekeeper sekali dari Extensions > Apps Script pada Google Sheet Anda, lalu coba kembali.');
  return SpreadsheetApp.openById(id);
}

// Read-only GET endpoints. Writes remain POST-only.
function doGet(e) { const p = Object.assign({}, e && e.parameter); p.action = p.action === 'getHistory' ? 'getHistory' : 'getStatus'; return handleRequest_(p); }
function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > 20000) return json_({success:false, code:'INVALID_REQUEST', message:'Request tidak valid.'});
    return handleRequest_(JSON.parse(raw));
  } catch (error) { return json_({success:false, code:'INVALID_JSON', message:'Format JSON tidak valid.'}); }
}
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
function fail_(code, message) { const error = new Error(message); error.code = code; throw error; }
function validDate_(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const d = new Date(value + 'T12:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0,10) === value && value >= '1900-01-01' && value <= '9999-12-31';
}
function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd');
  if (value === '') return '';
  if (validDate_(value)) return value;
  fail_('INVALID_SHEET_DATE', 'Kolom Date harus berupa tanggal Google Sheets atau YYYY-MM-DD.');
}
function normalizeTime_(value) {
  if (value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, TIMEZONE, 'HH:mm');
  if (typeof value === 'number' && value >= 0 && value < 1) {
    const n = Math.round(value * 1440) % 1440;
    return String(Math.floor(n/60)).padStart(2,'0') + ':' + String(n%60).padStart(2,'0');
  }
  if (typeof value === 'string' && /^([01]\d|2[0-3]):[0-5]\d$/.test(value)) return value;
  fail_('INVALID_SHEET_TIME','Jam di Google Sheet harus menggunakan format HH:mm.');
}
function handleRequest_(request) {
  let lock;
  try {
    if (!request || !['clockIn','clockOut','getStatus','getHistory'].includes(request.action)) fail_('INVALID_ACTION','Action tidak dikenali.');
    if (request.action !== 'getHistory' && !validDate_(request.date)) fail_('INVALID_DATE','Tanggal tidak valid.');
    if (['clockIn','clockOut'].includes(request.action)) {
      const time = request.action === 'clockIn' ? request.clockIn : request.clockOut;
      if (typeof time !== 'string' || !/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) fail_('INVALID_TIME','Jam harus menggunakan format HH:mm.');
      if (request.action === 'clockOut' && (typeof request.task !== 'string' || request.task.length > MAX_TASK_LENGTH || (TASK_REQUIRED && !request.task.trim()))) fail_('INVALID_TASK','Isi tugas dengan maksimal ' + MAX_TASK_LENGTH + ' karakter.');
    }
    lock = LockService.getScriptLock();
    if (!lock.tryLock(['getStatus','getHistory'].includes(request.action) ? 1000 : 10000)) fail_('BUSY','Database sedang sibuk. Silakan coba lagi.');
    const book = openTimekeeperSpreadsheet_();
    if (book.getSpreadsheetTimeZone() !== TIMEZONE) fail_('TIMEZONE_MISMATCH','Samakan timezone Google Sheet dengan ' + TIMEZONE + ' melalui File > Settings.');
    const sheet = book.getSheetByName(SHEET_NAME) || book.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) { sheet.getRange(1,1,1,5).setValues([HEADERS]); sheet.setFrozenRows(1); }
    if (sheet.getRange(1,1,1,5).getValues()[0].some((v,i) => v !== HEADERS[i])) fail_('INVALID_HEADERS','Header harus: ' + HEADERS.join(' | '));
    const totalRows = Math.max(0, sheet.getLastRow()-1);
    if (request.action === 'getHistory') {
      const paged = request.offset !== undefined;
      const offset = paged ? Number(request.offset) : 0;
      if (!Number.isSafeInteger(offset) || offset < 0) fail_('INVALID_PAGE','Offset riwayat tidak valid.');
      if (request.totalRows !== undefined && Number(request.totalRows) !== totalRows) fail_('HISTORY_CHANGED','Jumlah baris berubah saat memuat. Klik Refresh untuk mengambil ulang data.');
      const count = Math.min(HISTORY_BATCH_SIZE, Math.max(0,totalRows-offset));
      const rows = count ? sheet.getRange(offset+2,1,count,5).getValues() : [];
      const records = rows.map((row, i) => ({
        rowNumber:offset+i+2, date:normalizeDate_(row[0]), clockIn:normalizeTime_(row[1]),
        clockOut:normalizeTime_(row[2]), task:String(row[3] || ''),
        lastUpdated:row[4] instanceof Date ? row[4].toISOString() : String(row[4] || '')
      })).filter(record => record.date || record.clockIn || record.clockOut || record.task || record.lastUpdated);
      if (!paged && totalRows > HISTORY_BATCH_SIZE) fail_('CLIENT_UPDATE_REQUIRED','Gunakan index.html terbaru untuk memuat seluruh riwayat secara bertahap.');
      records.sort((a,b) => b.date.localeCompare(a.date) || b.rowNumber-a.rowNumber);
      return json_({success:true,message:'History loaded',data:records,pagination:{totalRows,nextOffset:offset+count<totalRows?offset+count:null}});
    }
    // Status only scans column A; task text for unrelated dates is not loaded.
    const rows = totalRows ? sheet.getRange(2,1,totalRows,1).getValues() : [];
    const matches = [];
    rows.forEach((row,i) => { if (normalizeDate_(row[0]) === request.date) matches.push(i); });
    if (matches.length > 1) fail_('DUPLICATE_DATA','Ada duplikasi tanggal di Sheet. Gabungkan baris tersebut sebelum melanjutkan.');
    const index = matches.length ? matches[0] : -1;
    if (index >= 0) rows[index] = sheet.getRange(index+2,1,1,5).getValues()[0];
    let record = index < 0 ? null : {date:request.date, clockIn:normalizeTime_(rows[index][1]), clockOut:normalizeTime_(rows[index][2]), task:String(rows[index][3] || ''), lastUpdated:rows[index][4] instanceof Date ? rows[index][4].toISOString() : String(rows[index][4] || '')};
    if (request.action === 'getStatus') return json_({success:true, message:'Status loaded', data:record});
    if (request.action === 'clockIn') {
      if (record) fail_('ALREADY_CLOCKED_IN','Timesheet untuk tanggal ini sudah memiliki Clock In.');
      const row = sheet.getLastRow()+1;
      sheet.getRange(row,1,1,4).setNumberFormat('@');
      sheet.getRange(row,1,1,5).setValues([[request.date,request.clockIn,'','',new Date()]]);
      record = {date:request.date,clockIn:request.clockIn,clockOut:'',task:''};
    } else {
      if (!record || !record.clockIn) fail_('NO_CLOCK_IN','Clock In untuk tanggal ini belum ditemukan.');
      if (request.clockOut < record.clockIn) fail_('EARLY_CLOCK_OUT','Clock Out tidak boleh lebih awal dari Clock In.');
      // RichText writes literal text, preventing spreadsheet formula injection.
      const row = index + 2;
      const task = request.task.trim();
      // Write task first: a partial failure must not claim a completed session.
      sheet.getRange(row,4).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(task).build());
      sheet.getRange(row,3).setNumberFormat('@').setValue(request.clockOut);
      sheet.getRange(row,5).setValue(new Date());
      record.clockOut = request.clockOut; record.task = task;
    }
    SpreadsheetApp.flush();
    record.lastUpdated = new Date().toISOString();
    return json_({success:true,message:request.action === 'clockIn' ? 'Clock In recorded' : 'Clock Out recorded',data:record});
  } catch (error) {
    console.error(error);
    return json_({success:false,code:error.code || 'SERVER_ERROR',message:error.code ? error.message : 'Database tidak dapat diakses. Periksa konfigurasi dan izin Apps Script.'});
  } finally { if (lock && lock.hasLock()) lock.releaseLock(); }
}
