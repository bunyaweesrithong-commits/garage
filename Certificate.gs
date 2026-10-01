/* ============================================================
 * ใบรับรอง
 * ============================================================ */

/** ข้อมูลทั้งหมดที่ใช้พิมพ์ใบรับรอง (ใช้ทั้งหน้า HTML และ Google Doc) */
function getCertificateData_(repairId) {
  const r = findById_(SHEETS.REPAIRS, REPAIR_COLS, 'repairId', repairId);
  if (!r) throw new Error('ไม่พบงานซ่อม');
  const v = findById_(SHEETS.VEHICLES, VEHICLE_COLS, 'vehicleId', r.vehicleId);
  if (!v) throw new Error('ไม่พบข้อมูลรถของงานซ่อมนี้');
  const s = getSettings_();
  const repair = toClient_(r, REPAIR_COLS);
  const vehicle = toClient_(v, VEHICLE_COLS);

  let logoDataUri = '';
  const logo = getLogoBlob_(s);
  if (logo) logoDataUri = 'data:' + logo.getContentType() + ';base64,' + Utilities.base64Encode(logo.getBytes());

  return {
    shop: {
      shopName: s.shopName || '', shopAddress: s.shopAddress || '', shopPhone: s.shopPhone || '',
      shopTaxId: s.shopTaxId || '', logoDataUri: logoDataUri,
    },
    text: {
      title: s.certificateTitle || 'ใบรับรองการซ่อมรถยนต์',
      subtitle: s.certificateSubtitle || '',
      certificate: s.certificateText || '',
      warranty: s.warrantyText || '',
    },
    vehicle: vehicle,
    repair: repair,
    fields: certificateFields_(s, vehicle, repair),
  };
}

/** ค่าที่จัดรูปแบบแล้ว (ใช้แทน {{placeholder}} ในแม่แบบ Google Doc ด้วย) */
function certificateFields_(s, v, r) {
  return {
    shopName: s.shopName || '', shopAddress: s.shopAddress || '', shopPhone: s.shopPhone || '', shopTaxId: s.shopTaxId || '',
    title: s.certificateTitle || '', subtitle: s.certificateSubtitle || '',
    certificateText: s.certificateText || '', warrantyText: s.warrantyText || '',
    jobNo: r.jobNo || '', runNo: String(r.runNo || ''), yearBE: String(r.yearBE || ''),
    issueDate: thaiDate_(toIso_(new Date())),
    customerName: v.customerName || '', phone: v.phone || '', otherContact: v.otherContact || '',
    plate: v.plate || '', province: v.province || '',
    plateFull: [v.plate, v.province].filter(String).join(' '),
    brand: v.brand || '', model: v.model || '',
    brandModel: [v.brand, v.model].filter(String).join(' '),
    color: v.color || '', chassisNo: v.chassisNo || '',
    mileage: r.mileage !== '' && r.mileage != null ? Number(r.mileage).toLocaleString('en-US') + ' กม.' : '',
    receivedDate: thaiDate_(r.receivedDate), completedDate: thaiDate_(r.completedDate),
    details: r.details || '', technician: r.technician || '',
    cost: r.cost !== '' && r.cost != null ? Number(r.cost).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' บาท' : '',
    note: r.note || '',
  };
}

/**
 * สร้างใบรับรองเป็น Google Doc + PDF เก็บใน Google Drive
 * ถ้าตั้งค่า templateDocId จะคัดลอกแม่แบบแล้วแทนที่ {{placeholder}} ถ้าไม่ตั้งจะสร้างตามแบบมาตรฐาน
 */
function createCertificateDoc_(repairId) {
  const data = getCertificateData_(repairId);
  const s = getSettings_();
  const f = data.fields;
  const folder = getOutputFolder_(s);
  const name = `ใบรับรองการซ่อม ${f.jobNo} ${f.plateFull}`.trim();

  let doc;
  if (s.templateDocId) {
    const copy = DriveApp.getFileById(s.templateDocId).makeCopy(name, folder);
    doc = DocumentApp.openById(copy.getId());
    fillTemplate_(doc, f);
  } else {
    doc = DocumentApp.create(name);
    DriveApp.getFileById(doc.getId()).moveTo(folder);
    buildDefaultCertificate_(doc, data, s);
  }
  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdf = folder.createFile(docFile.getAs(MimeType.PDF)).setName(name + '.pdf');

  withLock_(() => {
    const r = findById_(SHEETS.REPAIRS, REPAIR_COLS, 'repairId', repairId);
    if (!r) return;
    r.certDocUrl = docFile.getUrl();
    r.certPdfUrl = pdf.getUrl();
    writeRow_(SHEETS.REPAIRS, REPAIR_COLS, r, r._row);
  });

  return { name: name, docUrl: docFile.getUrl(), pdfUrl: pdf.getUrl() };
}

function fillTemplate_(doc, fields) {
  const parts = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(Boolean);
  Object.keys(fields).forEach(k => {
    const value = String(fields[k] == null ? '' : fields[k]);
    parts.forEach(p => p.replaceText('\\{\\{\\s*' + k + '\\s*\\}\\}', value));
  });
}

/** แบบใบรับรองมาตรฐาน (ขนาด A4) */
function buildDefaultCertificate_(doc, data, s) {
  const FONT = s.docFont || 'Sarabun';
  const NAVY = '#1f3a60', GREY = '#5b6573', LINE = '#c9ced6', SHADE = '#eef2f7';
  const A = DocumentApp.HorizontalAlignment;
  const f = data.fields;

  const body = doc.getBody();
  body.clear();
  body.setPageWidth(595.28).setPageHeight(841.89)
    .setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);
  const contentWidth = 595.28 - 108;

  let cursor = body.getParagraphs()[0];
  const style = (el, o) => {
    o = o || {};
    if (el.setAlignment) el.setAlignment(o.align || A.LEFT);
    if (el.setSpacingBefore) el.setSpacingBefore(o.before || 0).setSpacingAfter(o.after || 0).setLineSpacing(o.lineSpacing || 1.1);
    const t = el.editAsText();
    if (t.getText().length) {
      t.setFontFamily(FONT).setFontSize(o.size || 12).setBold(!!o.bold).setItalic(!!o.italic)
        .setForegroundColor(o.color || '#000000');
    }
    return el;
  };
  const P = (text, o) => {
    let p;
    if (cursor) { p = cursor; p.setText(text); cursor = null; } else { p = body.appendParagraph(text); }
    return style(p, o);
  };
  const setCellText = (cell, text, o) => {
    const lines = String(text).split(/\r?\n/);
    // เซลล์ที่สร้างจาก appendTable มีย่อหน้าเดียวอยู่แล้ว ใช้ย่อหน้านั้นเป็นบรรทัดแรก
    const first = cell.getChild(0).asParagraph();
    first.setText(lines[0]);
    style(first, o);
    for (let i = 1; i < lines.length; i++) style(cell.appendParagraph(lines[i]), o);
  };
  const padCell = (cell, pad) => cell.setPaddingTop(pad).setPaddingBottom(pad).setPaddingLeft(6).setPaddingRight(6);

  // ----- หัวกระดาษ: โลโก้ + ข้อมูลร้าน -----
  const logo = getLogoBlob_(s);
  if (logo) {
    const lp = cursor; cursor = null;
    const img = lp.appendInlineImage(logo);
    const h = 64, ratio = img.getWidth() / img.getHeight();
    img.setHeight(h).setWidth(Math.round(h * ratio));
    lp.setAlignment(A.CENTER).setSpacingAfter(2);
  }
  P(f.shopName, { align: A.CENTER, size: 18, bold: true, color: NAVY });
  if (f.shopAddress) P(f.shopAddress, { align: A.CENTER, size: 11, color: GREY });
  const contact = [f.shopPhone && 'โทร ' + f.shopPhone, f.shopTaxId && 'เลขประจำตัวผู้เสียภาษี ' + f.shopTaxId].filter(Boolean).join('   ');
  if (contact) P(contact, { align: A.CENTER, size: 11, color: GREY });
  body.appendHorizontalRule();

  // ----- หัวเรื่อง -----
  P(f.title, { align: A.CENTER, size: 20, bold: true, color: NAVY, before: 6 });
  if (f.subtitle) P(f.subtitle, { align: A.CENTER, size: 10, color: GREY, after: 4 });

  const meta = body.appendTable([[`เลขที่  ${f.jobNo}`, `วันที่ออกเอกสาร  ${f.issueDate}`]]);
  meta.setBorderWidth(0);
  setCellText(meta.getCell(0, 0), `เลขที่  ${f.jobNo}`, { size: 12, bold: true });
  setCellText(meta.getCell(0, 1), `วันที่ออกเอกสาร  ${f.issueDate}`, { size: 12, align: A.RIGHT });

  // ----- ตารางข้อมูล -----
  const section = (title, rows) => {
    P(title, { size: 13, bold: true, color: NAVY, before: 8, after: 3 });
    const t = body.appendTable(rows.map(r => [r[0], '']));
    t.setBorderColor(LINE).setBorderWidth(0.75);
    t.setColumnWidth(0, 150).setColumnWidth(1, contentWidth - 150);
    rows.forEach((r, i) => {
      const lc = t.getCell(i, 0), vc = t.getCell(i, 1);
      lc.setBackgroundColor(SHADE);
      setCellText(lc, r[0], { size: 12, bold: true, color: '#26303d' });
      setCellText(vc, r[1] || '-', { size: 12 });
      padCell(lc, 3); padCell(vc, 3);
    });
  };

  section('ข้อมูลลูกค้า', [
    ['ชื่อลูกค้า', f.customerName],
    ['เบอร์โทร', f.phone],
    ['ช่องทางติดต่ออื่น', f.otherContact],
  ]);
  section('ข้อมูลรถ', [
    ['ทะเบียนรถ', f.plateFull],
    ['ยี่ห้อ / รุ่น', f.brandModel],
    ['สี', f.color],
    ['เลขตัวถัง (คัสซี)', f.chassisNo],
    ['เลขไมล์', f.mileage],
  ]);
  const repairRows = [
    ['วันที่รับรถ', f.receivedDate],
    ['วันที่ซ่อมเสร็จ', f.completedDate],
    ['รายการซ่อม', f.details],
  ];
  if (f.technician) repairRows.push(['ช่างผู้รับผิดชอบ', f.technician]);
  if (f.note) repairRows.push(['หมายเหตุ', f.note]);
  section('รายละเอียดการซ่อม', repairRows);

  // ----- ข้อความรับรอง -----
  if (f.certificateText) P(f.certificateText, { size: 12, before: 12, after: 2, lineSpacing: 1.2 });
  if (f.warrantyText) P('เงื่อนไขการรับประกัน: ' + f.warrantyText, { size: 11, color: GREY, italic: true, after: 6 });

  // ----- ลายเซ็น -----
  P('', { before: 14 });
  const sign = ['ช่างผู้ซ่อม', 'ผู้รับรอง / ผู้จัดการ', 'ลูกค้าผู้รับรถ'];
  const st = body.appendTable([sign.map(() => '')]);
  st.setBorderWidth(0);
  sign.forEach((label, i) => {
    const who = i === 0 && f.technician ? f.technician : '                                        ';
    setCellText(st.getCell(0, i),
      `ลงชื่อ ................................\n(${who})\n${label}\nวันที่ ......../......../..........`,
      { size: 11, align: A.CENTER });
  });
}

/* ============================================================
 * ข้อมูลตัวอย่าง
 * ============================================================ */

function seedSampleData() {
  requireOwner_();
  if (['MR0HA3CD100123456', 'MRHGN6640MP012345'].some(c => searchVehicles_(c).total)) {
    try { SpreadsheetApp.getUi().alert('มีข้อมูลตัวอย่างอยู่แล้ว'); } catch (e) { /* ignore */ }
    return;
  }
  const id1 = saveVehicle_({ plate: 'กข 1234', province: 'กรุงเทพมหานคร', brand: 'Toyota', model: 'Hilux Revo 2.4', color: 'ขาว',
    chassisNo: 'MR0HA3CD100123456', customerName: 'สมชาย ใจดี', phone: '081-111-2222', otherContact: 'LINE: somchai.j' });
  saveRepair_({ vehicleId: id1, receivedDate: '2026-01-10', completedDate: '2026-01-12', mileage: 85200,
    details: 'เปลี่ยนน้ำมันเครื่อง + กรองน้ำมันเครื่อง\nเปลี่ยนผ้าเบรกหน้า\nตรวจเช็คช่วงล่าง', technician: 'ช่างเอก', cost: 4500 });
  saveRepair_({ vehicleId: id1, receivedDate: '2026-06-02', completedDate: '2026-06-05', mileage: 93850,
    details: 'ซ่อมระบบแอร์ เปลี่ยนคอมเพรสเซอร์\nเติมน้ำยาแอร์', technician: 'ช่างบอย', cost: 12800 });
  const id2 = saveVehicle_({ plate: '2กค 5678', province: 'เชียงใหม่', brand: 'Honda', model: 'City 1.0 Turbo', color: 'เทา',
    chassisNo: 'MRHGN6640MP012345', customerName: 'มาลี สวยงาม', phone: '089-333-4444', otherContact: 'Facebook: Malee S.' });
  saveRepair_({ vehicleId: id2, receivedDate: '2026-03-15', completedDate: '2026-03-15', mileage: 32010,
    details: 'เปลี่ยนแบตเตอรี่\nตรวจเช็คระบบไฟ', technician: 'ช่างเอก', cost: 3200 });
  try { SpreadsheetApp.getUi().alert('เพิ่มข้อมูลตัวอย่างแล้ว'); } catch (e) { /* ignore */ }
}
