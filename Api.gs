/* ============================================================
 * Web app
 * ============================================================ */

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.print) {
    const t = HtmlService.createTemplateFromFile('Certificate');
    let data;
    try {
      requireAuth_(params.t);
      data = getCertificateData_(params.print);
    } catch (err) {
      data = { error: err.message.indexOf('AUTH_REQUIRED') === 0
        ? 'ลิงก์หมดอายุ กรุณากลับไปหน้าหลัก ใส่รหัสผ่าน แล้วกดพิมพ์ใหม่' : err.message };
    }
    t.dataJson = safeJson_(data);
    return t.evaluate()
      .setTitle(data.error ? 'ไม่พบข้อมูล' : `ใบรับรอง ${data.repair.jobNo}`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle('ระบบประวัติการซ่อมรถ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** JSON ที่ฝังใน <script> ได้อย่างปลอดภัย */
function safeJson_(obj) {
  return JSON.stringify(obj)
    .replace(/</g, '\\u003c').replace(/>/g, '\\u003e')
    .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}

/* ============================================================
 * ตัวจริงของ API (หน้าเว็บเรียกผ่านฟังก์ชันใน Auth.gs ซึ่งตรวจรหัสผ่านก่อน)
 * ============================================================ */

function getAppInfo_() {
  const s = getSettings_();
  return {
    shopName: s.shopName || '',
    baseUrl: ScriptApp.getService().getUrl() || '',
    user: Session.getActiveUser().getEmail() || '',
  };
}

/**
 * ค้นหารถจาก ทะเบียน / ชื่อลูกค้า / เบอร์ / เลขตัวถัง / รุ่น / ช่องทางติดต่อ / เลขที่งาน (เช่น 12/2569)
 * เว้นว่าง = แสดงรายการล่าสุด
 */
function searchVehicles_(query) {
  const q = norm_(query);
  const vehicles = readTable_(SHEETS.VEHICLES, VEHICLE_COLS);
  const repairs = readTable_(SHEETS.REPAIRS, REPAIR_COLS);

  const stats = {};
  repairs.forEach(r => {
    const s = stats[r.vehicleId] || (stats[r.vehicleId] = { count: 0, lastDate: '', lastJobNo: '', jobNos: [] });
    s.count++;
    s.jobNos.push(norm_(r.jobNo));
    const d = toIso_(r.receivedDate);
    if (d >= s.lastDate) { s.lastDate = d; s.lastJobNo = r.jobNo; }
  });

  const fields = ['plate', 'province', 'brand', 'model', 'color', 'chassisNo', 'customerName', 'phone', 'otherContact'];
  const isJobNo = /^\d+\/\d+/.test(q);
  const matched = vehicles.filter(v => {
    if (!q) return true;
    if (isJobNo) {
      const s = stats[v.vehicleId];
      if (s && s.jobNos.some(j => stripZeros_(j) === stripZeros_(q))) return true;
    }
    return fields.some(f => norm_(v[f]).indexOf(q) !== -1);
  });

  matched.sort((a, b) => {
    const la = (stats[a.vehicleId] && stats[a.vehicleId].lastDate) || toIso_(a.updatedAt);
    const lb = (stats[b.vehicleId] && stats[b.vehicleId].lastDate) || toIso_(b.updatedAt);
    return la < lb ? 1 : la > lb ? -1 : 0;
  });

  return {
    total: matched.length,
    items: matched.slice(0, 200).map(v => {
      const s = stats[v.vehicleId] || { count: 0, lastDate: '', lastJobNo: '' };
      return Object.assign(toClient_(v, VEHICLE_COLS), {
        repairCount: s.count, lastRepairDate: s.lastDate, lastJobNo: s.lastJobNo,
      });
    }),
  };
}

/** ข้อมูลรถ 1 คัน พร้อมประวัติการซ่อมทั้งหมด (ใหม่ → เก่า) */
function getVehicle_(vehicleId) {
  const v = findById_(SHEETS.VEHICLES, VEHICLE_COLS, 'vehicleId', vehicleId);
  if (!v) throw new Error('ไม่พบข้อมูลรถ');
  const repairs = readTable_(SHEETS.REPAIRS, REPAIR_COLS)
    .filter(r => r.vehicleId === vehicleId)
    .map(r => toClient_(r, REPAIR_COLS))
    .sort((a, b) => (b.receivedDate + b.jobNo).localeCompare(a.receivedDate + a.jobNo));
  return { vehicle: toClient_(v, VEHICLE_COLS), repairs: repairs };
}

/** บันทึก (เพิ่ม/แก้ไข) ข้อมูลรถ */
function saveVehicle_(input) {
  const v = cleanInput_(input, VEHICLE_COLS);
  if (!v.plate) throw new Error('กรุณากรอกทะเบียนรถ');
  if (!v.customerName) throw new Error('กรุณากรอกชื่อลูกค้า');

  return withLock_(() => {
    const all = readTable_(SHEETS.VEHICLES, VEHICLE_COLS);
    const dup = all.find(x => x.vehicleId !== v.vehicleId && (
      (norm_(x.plate) === norm_(v.plate) && norm_(x.province) === norm_(v.province)) ||
      (v.chassisNo && norm_(x.chassisNo) === norm_(v.chassisNo))));
    if (dup) {
      throw new Error(`มีรถคันนี้ในระบบแล้ว (ทะเบียน ${dup.plate}${dup.province ? ' ' + dup.province : ''}, ลูกค้า ${dup.customerName})`);
    }

    const now = new Date();
    if (v.vehicleId) {
      const old = all.find(x => x.vehicleId === v.vehicleId);
      if (!old) throw new Error('ไม่พบข้อมูลรถที่ต้องการแก้ไข');
      v.createdAt = old.createdAt;
      v.updatedAt = now;
      writeRow_(SHEETS.VEHICLES, VEHICLE_COLS, v, old._row);
    } else {
      v.vehicleId = 'V' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
      v.createdAt = now;
      v.updatedAt = now;
      writeRow_(SHEETS.VEHICLES, VEHICLE_COLS, v);
    }
    return v.vehicleId;
  });
}

/** ลบรถ พร้อมประวัติการซ่อมทั้งหมดของรถคันนั้น */
/**
 * เพิ่มรถใหม่พร้อมงานซ่อมครั้งแรกในครั้งเดียว (ฟอร์ม "เพิ่มรถ / ลูกค้าใหม่")
 * ถ้าบันทึกงานซ่อมไม่ผ่าน (เช่น เลขที่ซ้ำ) จะลบรถที่เพิ่งเพิ่มออก เพื่อไม่ให้ข้อมูลค้างครึ่งเดียว
 */
function saveVehicleWithRepair_(vehicleInput, repairInput) {
  if (vehicleInput && vehicleInput.vehicleId) throw new Error('ใช้สำหรับเพิ่มรถใหม่เท่านั้น');
  if (!repairInput || !repairInput.receivedDate) throw new Error('กรุณากรอกวันที่รับรถ');
  const vehicleId = saveVehicle_(vehicleInput);
  try {
    const repair = saveRepair_(Object.assign({}, repairInput, { vehicleId: vehicleId, repairId: '' }));
    return { vehicleId: vehicleId, jobNo: repair.jobNo };
  } catch (err) {
    deleteVehicle_(vehicleId);
    throw err;
  }
}

function deleteVehicle_(vehicleId) {
  return withLock_(() => {
    const v = findById_(SHEETS.VEHICLES, VEHICLE_COLS, 'vehicleId', vehicleId);
    if (!v) throw new Error('ไม่พบข้อมูลรถ');
    const rows = readTable_(SHEETS.REPAIRS, REPAIR_COLS).filter(r => r.vehicleId === vehicleId).map(r => r._row);
    deleteRows_(SHEETS.REPAIRS, rows);
    deleteRows_(SHEETS.VEHICLES, [v._row]);
    return true;
  });
}

/** เลขที่/ปี ถัดไปของปีที่รับรถ (ใช้แสดงตัวอย่างในฟอร์ม) */
/** ลำดับถัดไปของงวดของวันที่รับรถ: { runNo, jobNo, template, digits } (template ใช้แสดงตัวอย่างในฟอร์ม) */
function getNextJobNo_(receivedDateIso) {
  const s = getSettings_();
  const period = periodOf_(parseDate_(receivedDateIso) || new Date());
  const runNo = nextRunNo_(readTable_(SHEETS.REPAIRS, REPAIR_COLS), period, s);
  return {
    runNo: runNo, jobNo: formatJobNo_(runNo, period, s),
    template: jobNoTemplate_(period, s), digits: Math.max(1, Number(s.jobNoDigits) || 1),
  };
}

/**
 * บันทึก (เพิ่ม/แก้ไข) ประวัติการซ่อม
 * งานใหม่ได้เลขที่อัตโนมัติ = ลำดับรถที่เข้าซ่อมในเดือน (หรือปี ตาม jobNoReset) ของวันที่รับรถ
 * ส่ง runNo มาเองได้ (เช่น ให้ตรงกับสมุดเดิม) แต่ห้ามซ้ำกับงานอื่นในงวดเดียวกัน
 */
function saveRepair_(input) {
  const r = cleanInput_(input, REPAIR_COLS);
  if (!r.vehicleId) throw new Error('ไม่ระบุรถ');
  if (!r.receivedDate) throw new Error('กรุณากรอกวันที่รับรถ');
  r.receivedDate = parseDate_(r.receivedDate);
  r.completedDate = parseDate_(r.completedDate);
  if (r.completedDate && r.completedDate < r.receivedDate) throw new Error('วันที่ซ่อมเสร็จต้องไม่ก่อนวันที่รับรถ');
  r.mileage = toNumber_(r.mileage);
  r.cost = toNumber_(r.cost);

  return withLock_(() => {
    if (!findById_(SHEETS.VEHICLES, VEHICLE_COLS, 'vehicleId', r.vehicleId)) throw new Error('ไม่พบข้อมูลรถ');
    const all = readTable_(SHEETS.REPAIRS, REPAIR_COLS);
    const now = new Date();

    if (r.repairId) {
      const old = all.find(x => x.repairId === r.repairId);
      if (!old) throw new Error('ไม่พบงานซ่อมที่ต้องการแก้ไข');
      // เลขที่/ปี ออกแล้วไม่เปลี่ยน เพื่อให้ตรงกับใบรับรองที่พิมพ์ไปแล้ว
      ['jobNo', 'runNo', 'yearBE', 'periodKey', 'certDocUrl', 'certPdfUrl', 'createdAt'].forEach(k => { r[k] = old[k]; });
      r.updatedAt = now;
      writeRow_(SHEETS.REPAIRS, REPAIR_COLS, r, old._row);
    } else {
      r.repairId = 'R' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
      const s = getSettings_();
      r.periodKey = periodOf_(r.receivedDate);
      r.yearBE = toYearBE_(r.receivedDate);
      if (r.runNo === '' || r.runNo == null) {
        r.runNo = nextRunNo_(all, r.periodKey, s);
      } else {
        r.runNo = Number(r.runNo);
        if (!Number.isInteger(r.runNo) || r.runNo < 1) throw new Error('ลำดับเลขที่ต้องเป็นตัวเลขจำนวนเต็มตั้งแต่ 1 ขึ้นไป');
        const used = all.find(x => samePeriod_(repairPeriod_(x), r.periodKey, s) && Number(x.runNo) === r.runNo);
        if (used) throw new Error(`เลขที่ ${formatJobNo_(r.runNo, r.periodKey, s)} ถูกใช้แล้ว (${used.jobNo}) กรุณาใช้เลขอื่น`);
      }
      r.jobNo = formatJobNo_(r.runNo, r.periodKey, s);
      r.certDocUrl = '';
      r.certPdfUrl = '';
      r.createdAt = now;
      r.updatedAt = now;
      writeRow_(SHEETS.REPAIRS, REPAIR_COLS, r);
    }
    return { repairId: r.repairId, jobNo: r.jobNo };
  });
}

function deleteRepair_(repairId) {
  return withLock_(() => {
    const r = findById_(SHEETS.REPAIRS, REPAIR_COLS, 'repairId', repairId);
    if (!r) throw new Error('ไม่พบงานซ่อม');
    deleteRows_(SHEETS.REPAIRS, [r._row]);
    return true;
  });
}
