/* ============================================================
 * รหัสผ่าน (PIN) หน้าเว็บ
 *
 * - ตั้งรหัสที่แท็บ Settings แถว accessPin
 * - ใส่รหัสถูก → ได้ token ที่เบราว์เซอร์จำไว้ loginDays วัน
 * - token เซ็นด้วย HMAC ที่ผูกกับรหัสปัจจุบัน เปลี่ยนรหัสเมื่อไร token เดิมใช้ไม่ได้ทันที (ทุกคนต้องใส่รหัสใหม่)
 * - ใส่ผิดเกิน AUTH_MAX_FAILS ครั้ง ระบบล็อกการเข้าสู่ระบบ AUTH_LOCK_SECONDS วินาที
 * ============================================================ */

const AUTH_MAX_FAILS = 10;
const AUTH_LOCK_SECONDS = 15 * 60;
const AUTH_ERROR = 'AUTH_REQUIRED: กรุณาใส่รหัสผ่านอีกครั้ง';

/** ข้อมูลที่แสดงบนหน้าใส่รหัส (ไม่ต้องมี token) */
function getLoginInfo() {
  return { shopName: getSettings_().shopName || '' };
}

/** ตรวจรหัส แล้วคืน token */
function login(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('auth_fails') || 0);
  if (fails >= AUTH_MAX_FAILS) throw new Error('ใส่รหัสผิดหลายครั้งเกินไป กรุณารอ 15 นาทีแล้วลองใหม่');

  const s = getSettings_();
  const real = currentPin_(s);
  if (!real) throw new Error('ยังไม่ได้ตั้งรหัสผ่าน: ให้เจ้าของระบบใส่ accessPin (อย่างน้อย 4 ตัว) ในแท็บ Settings');

  Utilities.sleep(300); // ชะลอการเดารหัส
  if (String(pin == null ? '' : pin).trim() !== real) {
    cache.put('auth_fails', String(fails + 1), AUTH_LOCK_SECONDS);
    throw new Error('รหัสผ่านไม่ถูกต้อง');
  }
  const days = Math.max(1, Number(s.loginDays) || 30);
  const exp = String(Date.now() + days * 24 * 60 * 60 * 1000);
  return exp + '.' + signToken_(exp, real);
}

function requireAuth_(token) {
  const m = /^(\d+)\.([A-Za-z0-9_\-]+=*)$/.exec(String(token || ''));
  if (!m || Number(m[1]) < Date.now()) throw new Error(AUTH_ERROR);
  const pin = currentPin_(getSettings_());
  if (!pin || signToken_(m[1], pin) !== m[2]) throw new Error(AUTH_ERROR);
}

/** ฟังก์ชันสำหรับเจ้าของ (setup, ข้อมูลตัวอย่าง) รันได้จาก Apps Script/เมนูชีตเท่านั้น ไม่ใช่จากหน้าเว็บ */
function requireOwner_() {
  if (!Session.getActiveUser().getEmail()) throw new Error('ฟังก์ชันนี้รันได้จาก Apps Script หรือเมนูในชีตเท่านั้น');
}

function currentPin_(s) {
  const pin = String(s.accessPin == null ? '' : s.accessPin).trim();
  return pin.length >= 4 ? pin : '';
}

function signToken_(exp, pin) {
  const sig = Utilities.computeHmacSha256Signature(exp + '|' + pin, getAuthSecret_());
  return Utilities.base64EncodeWebSafe(sig);
}

/** กุญแจลับสำหรับเซ็น token สร้างครั้งแรกแล้วเก็บใน Script Properties */
function getAuthSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('AUTH_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('AUTH_SECRET', secret);
  }
  return secret;
}

/* ============================================================
 * API ที่หน้าเว็บเรียกผ่าน google.script.run
 * ทุกฟังก์ชันต้องได้ token ที่ถูกต้องก่อนจึงทำงาน
 * ============================================================ */

function getAppInfo(token) { requireAuth_(token); return getAppInfo_(); }
function searchVehicles(token, query) { requireAuth_(token); return searchVehicles_(query); }
function getVehicle(token, vehicleId) { requireAuth_(token); return getVehicle_(vehicleId); }
function saveVehicle(token, input) { requireAuth_(token); return saveVehicle_(input); }
function saveVehicleWithRepair(token, vehicle, repair) { requireAuth_(token); return saveVehicleWithRepair_(vehicle, repair); }
function deleteVehicle(token, vehicleId) { requireAuth_(token); return deleteVehicle_(vehicleId); }
function getNextJobNo(token, receivedDateIso) { requireAuth_(token); return getNextJobNo_(receivedDateIso); }
function saveRepair(token, input) { requireAuth_(token); return saveRepair_(input); }
function deleteRepair(token, repairId) { requireAuth_(token); return deleteRepair_(repairId); }
function createCertificateDoc(token, repairId) { requireAuth_(token); return createCertificateDoc_(repairId); }
