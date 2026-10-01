/* ============================================================
 * ตัวช่วยจัดการชีต
 * ============================================================ */

function getSS_() {
  const ss = findSpreadsheet_();
  if (!ss) throw new Error('ยังไม่ได้ตั้งค่าระบบ กรุณารันฟังก์ชัน setup ใน Apps Script ก่อน');
  return ss;
}

/** ชีตที่สคริปต์ผูกอยู่ → SPREADSHEET_ID ในโค้ด → ชีตที่ setup() บันทึกไว้ */
function findSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : null;
}

function sheet_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error(`ไม่พบชีต "${name}" กรุณารันเมนู "ตั้งค่าเริ่มต้น" ก่อน`);
  return sh;
}

function readTable_(name, cols) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  const out = [];
  values.forEach((row, i) => {
    if (row[0] === '' || row[0] == null) return;
    const o = { _row: i + 2 };
    cols.forEach((c, j) => { o[c[0]] = row[j]; });
    out.push(o);
  });
  return out;
}

function findById_(name, cols, key, id) {
  if (!id) return null;
  return readTable_(name, cols).find(o => o[key] === id) || null;
}

function writeRow_(name, cols, obj, rowIndex) {
  const sh = sheet_(name);
  const row = rowIndex || sh.getLastRow() + 1;
  const range = sh.getRange(row, 1, 1, cols.length);
  // ตั้งรูปแบบก่อนเขียน เพื่อไม่ให้ 081... หาย 0 หรือ 1/2569 กลายเป็นวันที่
  range.setNumberFormats([cols.map(c => numberFormat_(c[2]))]);
  range.setValues([cols.map(c => {
    const val = obj[c[0]];
    if (val === undefined || val === null) return '';
    return c[2] === 't' ? String(val) : val;
  })]);
}

function deleteRows_(name, rows) {
  const sh = sheet_(name);
  rows.slice().sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}

function cleanInput_(input, cols) {
  const o = {};
  cols.forEach(c => {
    const val = input ? input[c[0]] : undefined;
    o[c[0]] = typeof val === 'string' ? val.trim() : (val == null ? '' : val);
  });
  return o;
}

/** แปลงเป็นข้อมูลที่ส่งให้หน้าเว็บได้ (google.script.run ส่ง Date ไม่ได้) */
function toClient_(obj, cols) {
  const o = {};
  cols.forEach(c => {
    const val = obj[c[0]];
    if (c[2] === 'd') o[c[0]] = toIso_(val);
    else if (c[2] === 'ts') o[c[0]] = val instanceof Date ? Utilities.formatDate(val, tz_(), 'yyyy-MM-dd HH:mm') : String(val || '');
    else if (c[2] === 'n') o[c[0]] = val === '' || val == null ? '' : Number(val);
    else o[c[0]] = val == null ? '' : String(val);
  });
  return o;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function readSettingsRaw_(sh) {
  const out = {};
  const last = sh.getLastRow();
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    if (r[0] !== '') out[String(r[0]).trim()] = r[1] instanceof Date ? r[1] : String(r[1]).trim();
  });
  return out;
}

function getSettings_() {
  const defaults = {};
  DEFAULT_SETTINGS.forEach(s => { defaults[s[0]] = s[1]; });
  const sh = getSS_().getSheetByName(SHEETS.SETTINGS);
  return Object.assign(defaults, sh ? readSettingsRaw_(sh) : {});
}

function getLogoBlob_(s) {
  if (!s.logoFileId) return null;
  try {
    return DriveApp.getFileById(s.logoFileId).getBlob();
  } catch (e) {
    console.warn('โหลดโลโก้ไม่ได้: ' + e.message);
    return null;
  }
}

function getOutputFolder_(s) {
  if (s.docFolderId) return DriveApp.getFolderById(s.docFolderId);
  const name = 'ใบรับรองการซ่อมรถ';
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/* ============================================================
 * ตัวช่วยทั่วไป
 * ============================================================ */

function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Bangkok';
}

function norm_(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/[\s\-.]/g, '');
}

function stripZeros_(jobNo) {
  return String(jobNo).replace(/^0+(?=\d)/, '');
}

function toNumber_(v) {
  if (v === '' || v == null) return '';
  const n = Number(String(v).replace(/,/g, ''));
  if (isNaN(n)) throw new Error('ตัวเลขไม่ถูกต้อง: ' + v);
  return n;
}

/** 'yyyy-MM-dd' → Date (เวลาเที่ยงคืนตามโซนเวลาสคริปต์) */
function parseDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v));
  if (!m) throw new Error('รูปแบบวันที่ไม่ถูกต้อง: ' + v);
  let y = Number(m[1]);
  if (y > 2400) y -= 543; // เผื่อกรอกปี พ.ศ.
  return new Date(y, Number(m[2]) - 1, Number(m[3]));
}

function toIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}

function toYearBE_(d) {
  return Number(Utilities.formatDate(d, tz_(), 'yyyy')) + 543;
}

function nextRunNo_(repairs, yearBE) {
  let max = 0;
  repairs.forEach(r => {
    if (Number(r.yearBE) === Number(yearBE)) max = Math.max(max, Number(r.runNo) || 0);
  });
  return max + 1;
}

function formatJobNo_(runNo, yearBE) {
  const digits = Math.max(1, Number(getSettings_().jobNoDigits) || 1);
  return String(runNo).padStart(digits, '0') + '/' + yearBE;
}

/** '2026-09-29' → '29 กันยายน 2569' */
function thaiDate_(iso) {
  if (!iso) return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso));
  if (!m) return String(iso);
  return `${Number(m[3])} ${THAI_MONTHS[Number(m[2]) - 1]} ${Number(m[1]) + 543}`;
}
