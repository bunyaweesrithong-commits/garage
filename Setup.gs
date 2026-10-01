/* ============================================================
 * เมนู / ติดตั้ง
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚗 ระบบอู่ซ่อมรถ')
    .addItem('1) ตั้งค่าเริ่มต้น (สร้างชีต)', 'setup')
    .addItem('2) เพิ่มข้อมูลตัวอย่าง', 'seedSampleData')
    .addSeparator()
    .addItem('เปิดหน้าเว็บ', 'showWebAppLink')
    .addToUi();
}

/** สร้างชีตและหัวคอลัมน์ (รันซ้ำได้ ไม่ลบข้อมูลเดิม) */
function setup() {
  requireOwner_();
  let ss = findSpreadsheet_();
  let created = false;
  if (!ss) {
    ss = SpreadsheetApp.create('ฐานข้อมูลอู่ซ่อมรถ');
    created = true;
  }
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  setupTable_(ss, SHEETS.VEHICLES, VEHICLE_COLS);
  setupTable_(ss, SHEETS.REPAIRS, REPAIR_COLS);

  let st = ss.getSheetByName(SHEETS.SETTINGS);
  if (!st) st = ss.insertSheet(SHEETS.SETTINGS);
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 1, 3).setValues([['key', 'value', 'คำอธิบาย']]).setFontWeight('bold').setBackground('#1f3a60').setFontColor('#ffffff');
    st.setFrozenRows(1);
  }
  const existing = readSettingsRaw_(st);
  const missing = DEFAULT_SETTINGS.filter(s => !(s[0] in existing));
  if (missing.length) {
    const start = st.getLastRow() + 1;
    st.getRange(start, 1, missing.length, 3).setNumberFormat('@').setValues(missing);
  }
  st.setColumnWidth(1, 170); st.setColumnWidth(2, 420); st.setColumnWidth(3, 420);

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  console.log((created ? 'สร้างชีตฐานข้อมูลใหม่: ' : 'ตั้งค่าชีตฐานข้อมูล: ') + ss.getUrl());
  try {
    SpreadsheetApp.getUi().alert('ตั้งค่าเรียบร้อย\n\nแก้ไขข้อมูลร้านได้ที่ชีต "Settings" แล้ว Deploy เป็น Web app เพื่อใช้งานหน้าเว็บ');
  } catch (e) { /* รันจาก editor */ }
}

function setupTable_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length)
    .setValues([cols.map(c => c[1])])
    .setFontWeight('bold').setBackground('#1f3a60').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);
  const maxRows = sh.getMaxRows();
  cols.forEach((c, i) => {
    if (maxRows > 1) sh.getRange(2, i + 1, maxRows - 1, 1).setNumberFormat(numberFormat_(c[2]));
  });
}

function numberFormat_(type) {
  switch (type) {
    case 'd': return 'dd/MM/yyyy';
    case 'ts': return 'dd/MM/yyyy HH:mm';
    case 'n': return '#,##0.##';
    default: return '@';
  }
}

function showWebAppLink() {
  const url = ScriptApp.getService().getUrl();
  const html = url
    ? `<p style="font-family:sans-serif">เปิดหน้าเว็บ: <a href="${url}" target="_blank">${url}</a></p>`
    : '<p style="font-family:sans-serif">ยังไม่ได้ Deploy เป็น Web app<br>ไปที่ Extensions → Apps Script → Deploy → New deployment → Web app</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(120), 'หน้าเว็บระบบอู่ซ่อมรถ');
}
