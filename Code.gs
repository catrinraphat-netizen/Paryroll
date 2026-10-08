/**
 * ระบบเงินเดือน (Payroll Template) — Google Apps Script (Code.gs)
 * รับผล OCR จากบอท LINE (doPost) → เซฟรูปเข้า Drive + เขียนลง Sheet → แสดงหน้าเว็บ (doGet)
 * ฟีเจอร์: แก้ไข/เตือนค่าผิดปกติ/กันซ้ำ · โบนัส-หัก · สลิป PDF ครบฟิลด์ + ออกทุกคน ·
 *          ส่งออก CSV · ประวัติรายบุคคล · คำนวณเงินเดือน/OT
 *
 * ติดตั้ง: สร้างโปรเจกต์ Apps Script "ใหม่" → วาง Code.gs / TimesheetOcr.gs / Index.html (HTML ชื่อ Index)
 *          → Run setup() 1 ครั้ง (กดอนุญาตสิทธิ์) → Deploy เป็น Web app (Execute as Me, Access Anyone)
 *          → เปิด URL /exec → เมนู "ตั้งค่าระบบ" ใส่ชื่อบริษัท + อัปโหลดโลโก้เอง
 * ไม่มีชื่อบริษัท/โลโก้/ID ฮาร์ดโค้ด — setup() สร้างโฟลเดอร์ + สเปรดชีตใหม่ของโปรเจกต์นี้เอง
 */

// ชื่อโฟลเดอร์/สเปรดชีตใน Drive — getRoot_() ของเทมเพลตนี้ "สร้างใหม่" แล้วจำ ID ไว้ใน Script Properties
// (ไม่ค้นด้วยชื่อ) ติดตั้งหลายระบบในบัญชีเดียวกันจึงไม่หยิบโฟลเดอร์ของกันและกันมาใช้
var FOLDER_ROOT = 'Payroll_System';
var SUBFOLDERS = ['01_Inbox', '02_Processing', '03_Completed', '04_Error', '05_Payslips', '06_Reports'];
var SS_NAME = 'Payroll_Master';
var TZ = 'Asia/Bangkok';

var SHEETS = {
  Attendance: ['id', 'received_at', 'line_user_id', 'display_name', 'employee_id', 'employee_name',
    'work_date', 'check_in', 'check_out', 'break_minutes', 'normal_hours', 'ot_hours',
    'status', 'source_file_url', 'summary'],
  // คอลัมน์ใหม่ให้ "ต่อท้ายเสมอ" — api ที่เขียนแถวใช้ลำดับตายตัว ถ้าแทรกกลางจะเหลื่อมกับชีตเดิมที่ migrate มา
  Employees: ['employee_id', 'employee_name', 'type', 'hourly_rate', 'ot_rate', 'status', 'position', 'start_date', 'phone',
    'national_id', 'bank_name', 'bank_account'],
  Positions: ['position_id', 'position_name', 'department', 'hourly_rate', 'ot_rate', 'status', 'note'],
  Schedule: ['employee_id', 'weekday', 'shift_start', 'shift_end', 'note'],
  Adjustments: ['employee_id', 'period_from', 'period_to', 'addition', 'bonus', 'deduction', 'note', 'diligence', 'wht'],
  Import_Log: ['imported_at', 'line_user_id', 'display_name', 'file_url', 'status', 'note'],
  Settings: ['key', 'value'],
};

var SCHEMA_V = '5';   // เพิ่มตอนแก้โครงสร้างชีต → apiBootstrap จะ migrate อัตโนมัติ 1 ครั้ง
                      // 4 = จัดลำดับคอลัมน์ Master ใหม่ + ซ่อมข้อมูลเก่าที่เหลื่อมคอลัมน์
                      // 5 = ฟิลด์สำหรับสลิป/ทะเบียนเงินได้ (เลขบัตร ปชช./ธนาคาร/เบี้ยขยัน/ภาษีหัก ณ ที่จ่าย)

// ที่อยู่ / เลขผู้เสียภาษี / โทร / อีเมล / เว็บ — ยังไม่ได้รับจากลูกค้า ห้ามเดา → ว่างไว้ กรอกที่หน้า "ตั้งค่าระบบ"
// (สลิป/ทะเบียนเงินได้ซ่อนบรรทัดที่ค่าว่างให้เองอยู่แล้ว)
var DEFAULT_SETTINGS = {
  company_name: '', company_name_en: '',   // ผู้ใช้ตั้งเองที่หน้า "ตั้งค่าระบบ"
  company_address: '',
  tax_id: '', company_phone: '', company_email: '', company_website: '',
  ot_multiplier: '1.5', standard_day_hours: '8', default_break_minutes: '60',
  ot_min_minutes: '0',    // ทำเกินชั่วโมงปกติน้อยกว่านี้ = ไม่นับเป็น OT (0 = จ่ายทุกเศษ ตามที่ผู้ใช้เลือก)
  slip_prefix: 'SLIP', pay_date: '', slip_note: 'ขอให้ท่านตรวจสอบรายการข้างต้นให้ถูกต้อง',
};

/* ================= โลโก้บริษัท (ผู้ใช้อัปโหลดเองที่หน้า "ตั้งค่าระบบ") =================
   เก็บเป็นไฟล์ในโฟลเดอร์ root ของระบบ · จำ file id ไว้ใน Script Properties LOGO_FILE_ID
   ส่งให้หน้าเว็บเป็น data URI (ห้ามส่งลิงก์ Drive — คนเปิดเว็บแบบไม่ล็อกอินไม่มีสิทธิ์ดูไฟล์)
   หน้าเว็บย่อรูปก่อนส่ง (ด้านยาวไม่เกิน 600px) → ไฟล์เล็ก เก็บแคชได้ ไม่ต้องอ่าน Drive ทุกคำขอ */
var LOGO_MAX = 1024 * 1024;
var LOGO_CACHE_KEY = 'COMPANY_LOGO_URI';
var LOGO_MEMO_ = null;   // แคชต่อคำขอ (สลิปหลายคนเรียกซ้ำ)
// โลโก้ตั้งต้น SPS (PNG 526x201 ฝังเป็น data URI) — แสดงเมื่อยังไม่มีโลโก้ที่อัปโหลด · แหล่งไฟล์: assets/sps-logo.png
var DEFAULT_LOGO_URI = 'data:image/png;base64,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';
function getLogo_() {
  if (LOGO_MEMO_ !== null) return LOGO_MEMO_;
  var cache = CacheService.getScriptCache();
  var hit = cache.get(LOGO_CACHE_KEY);
  if (hit !== null) { LOGO_MEMO_ = hit === '-' ? '' : hit; return LOGO_MEMO_; }
  var uri = '';
  var id = props_().getProperty('LOGO_FILE_ID');
  if (id) {
    try {
      var blob = DriveApp.getFileById(id).getBlob();
      uri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (e) { uri = ''; }
  }
  if (!uri) uri = DEFAULT_LOGO_URI;   // ยังไม่อัปโหลด/ลบโลโก้ = ใช้โลโก้ SPS ที่ฝังไว้
  if (uri.length < 95000) cache.put(LOGO_CACHE_KEY, uri || '-', 21600);   // CacheService รับได้ไม่เกิน 100KB ต่อค่า
  LOGO_MEMO_ = uri;
  return uri;
}
/** รับโลโก้จากหน้าตั้งค่า: dataUri = "data:image/png;base64,..." (PNG / JPEG / WEBP ≤ 1 MB)
    เว็บเปิดแบบไม่ล็อกอิน → ตรวจ magic bytes จริง ไม่เชื่อ mime ที่ส่งมา */
function apiSaveLogo(pin, dataUri) {
  guard_(pin);
  var m = String(dataUri || '').match(/^data:image\/(?:png|jpeg|webp);base64,([A-Za-z0-9+\/=\s]+)$/);
  if (!m) return { ok: false, error: 'รองรับเฉพาะรูป PNG / JPG / WEBP' };
  var bytes;
  try { bytes = Utilities.base64Decode(m[1].replace(/\s+/g, '')); } catch (e) { return { ok: false, error: 'ไฟล์รูปเสียหาย' }; }
  if (!bytes.length || bytes.length > LOGO_MAX) return { ok: false, error: 'รูปว่าง หรือใหญ่เกิน 1 MB' };
  var b = function (i) { return bytes[i] & 255; };
  var isPng = b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4E && b(3) === 0x47;
  var isJpg = b(0) === 0xFF && b(1) === 0xD8 && b(2) === 0xFF;
  var isWebp = b(0) === 0x52 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x46 && b(8) === 0x57 && b(9) === 0x45;
  if (!isPng && !isJpg && !isWebp) return { ok: false, error: 'ไฟล์นี้ไม่ใช่รูปภาพ' };
  var mime = isPng ? 'image/png' : (isJpg ? 'image/jpeg' : 'image/webp');
  var file = getRoot_().createFile(Utilities.newBlob(bytes, mime, 'company-logo.' + mime.split('/')[1]));
  var old = props_().getProperty('LOGO_FILE_ID');
  props_().setProperty('LOGO_FILE_ID', file.getId());
  if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
  CacheService.getScriptCache().remove(LOGO_CACHE_KEY);
  LOGO_MEMO_ = null;
  return { ok: true, logo: getLogo_() };
}
function apiRemoveLogo(pin) {
  guard_(pin);
  var old = props_().getProperty('LOGO_FILE_ID');
  props_().deleteProperty('LOGO_FILE_ID');
  if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
  CacheService.getScriptCache().remove(LOGO_CACHE_KEY);
  LOGO_MEMO_ = null;
  return { ok: true, logo: getLogo_() };   // กลับไปใช้โลโก้ตั้งต้น (DEFAULT_LOGO_URI)
}

// ตำแหน่งตั้งต้น (ค่าตัวอย่าง — ผู้ใช้แก้/เพิ่ม/ลบเองที่เมนู "ตำแหน่งงาน" · ปุ่ม "โหลดค่าตั้งต้น" ใส่ชุดนี้ให้)
// เรียงตาม SHEETS.Positions: รหัส · ตำแหน่ง · แผนก · ค่าแรง/ชม. · OT/ชม. · สถานะ · หมายเหตุ
// (OT = ค่าแรง × 1.5 ตาม ot_multiplier ที่หน้าตั้งค่าระบบ — ตัวคูณเป็นค่ากลาง ไม่ได้แยกรายตำแหน่ง)
var DEFAULT_POSITIONS = [
  ['PT001', 'พนักงานทั่วไป', 'General', 80, 120, 'ใช้งาน', 'ตัวอย่าง — แก้ค่าแรงให้ตรงบริษัท'],
  ['PT002', 'พนักงานขาย', 'Sales', 90, 135, 'ใช้งาน', 'ตัวอย่าง'],
  ['PT003', 'พนักงานคลังสินค้า', 'Warehouse', 90, 135, 'ใช้งาน', 'ตัวอย่าง'],
  ['PT004', 'พนักงานธุรการ', 'Office', 85, 127.5, 'ใช้งาน', 'ตัวอย่าง'],
];
function seedPositions_(ss) {
  var pos = ss.getSheetByName('Positions');
  if (!pos || pos.getLastRow() > 1) return;
  DEFAULT_POSITIONS.forEach(function (p) { pos.appendRow(p); });
}

/* ================= Init / accessors ================= */
function props_() { return PropertiesService.getScriptProperties(); }

function getRoot_() {
  var id = props_().getProperty('ROOT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var root = DriveApp.createFolder(FOLDER_ROOT);   // ไม่ค้นด้วยชื่อ — กันหยิบโฟลเดอร์ของระบบอื่นในบัญชีเดียวกัน
  SUBFOLDERS.forEach(function (n) { ensureFolder_(root, n); });
  props_().setProperty('ROOT_FOLDER_ID', root.getId());
  return root;
}
function ensureFolder_(parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function subFolder_(name) { return ensureFolder_(getRoot_(), name); }

/* แคชต่อ "การเรียก 1 ครั้ง" (Apps Script เริ่มตัวแปรส่วนกลางใหม่ทุกคำขอ จึงไม่มีข้อมูลค้างข้ามคำขอ)
   เดิมทุก readSheet_/readSettings_ เปิดสเปรดชีตใหม่ และ calcRowHours_ อ่านชีต Settings ใหม่ "ทุกแถว"
   → หน้าภาพรวม/บันทึกเวลาใช้ 10–14 วินาที ทั้งที่ข้อมูลไม่กี่สิบแถว */
var SS_CACHE_ = null, SETTINGS_CACHE_ = null;
function getSs_() {
  if (SS_CACHE_) return SS_CACHE_;
  var id = props_().getProperty('SPREADSHEET_ID');
  if (id) { try { SS_CACHE_ = SpreadsheetApp.openById(id); return SS_CACHE_; } catch (e) {} }
  var ss = SpreadsheetApp.create(SS_NAME); SS_CACHE_ = ss;
  try { DriveApp.getFileById(ss.getId()).moveTo(getRoot_()); } catch (e) {}
  props_().setProperty('SPREADSHEET_ID', ss.getId());
  initSheets_(ss);
  return ss;
}
function initSheets_(ss) {
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(SHEETS[name]);
  });
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  var settings = ss.getSheetByName('Settings');
  if (settings.getLastRow() <= 1) { Object.keys(DEFAULT_SETTINGS).forEach(function (k) { settings.appendRow([k, DEFAULT_SETTINGS[k]]); }); SETTINGS_CACHE_ = null; }
  // ไม่สร้างข้อมูลตัวอย่างอัตโนมัติ เพื่อไม่ให้ PT001/PT002 ปนกับ Master จริง
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/** เพิ่มชีต/คอลัมน์ที่ยังไม่มี (สำหรับสเปรดชีตเดิมที่สร้างก่อนอัปเดต) — ปลอดภัย ไม่ลบข้อมูลเก่า */
function migrate_() {
  var ss = getSs_();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(SHEETS[name]); return; }
    var lc = Math.max(1, sh.getLastColumn());
    var head = sh.getRange(1, 1, 1, lc).getValues()[0];
    var missing = SHEETS[name].filter(function (c) { return head.indexOf(c) < 0; });
    if (missing.length) sh.getRange(1, lc + 1, 1, missing.length).setValues([missing]);
  });
  repairMasterSheets_(ss);
  renameStaleCompany_();
}


/** ซ่อมและจัดลำดับคอลัมน์ Master ให้ตรงกับ SHEETS เสมอ
 *  รองรับชีตเดิมที่เคยเติมหัวคอลัมน์ต่อท้ายจนข้อมูลเหลื่อม
 *  และลบเฉพาะข้อมูลตัวอย่างเดิมที่ระบบเคยสร้างอัตโนมัติ */
function repairMasterSheets_(ss) {
  rebuildMasterSheet_(ss, 'Positions', normalizePositionRow_);
  rebuildMasterSheet_(ss, 'Employees', normalizeEmployeeRow_);
}

function rebuildMasterSheet_(ss, name, normalizer) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var values = sh.getDataRange().getDisplayValues();
  var head = values.length ? values[0] : [];
  var out = [];
  var at = {};   // รหัสซ้ำ (เช่นเคยบันทึก "EMP031" กับ "EMP031␣") → เหลือแถวเดียว ใช้ข้อมูลแถวล่างสุด (แก้ล่าสุด) วางที่ตำแหน่งแรก
  for (var r = 1; r < values.length; r++) {
    var row = normalizer(values[r], head);
    if (!row) continue;
    row[0] = String(row[0] == null ? '' : row[0]).trim();
    if (row[0] && at[row[0]] !== undefined) { out[at[row[0]]] = row; continue; }
    if (row[0]) at[row[0]] = out.length;
    out.push(row);
  }
  if (name === 'Employees') out.forEach(function (r) { r[8] = phoneTxt_(r[8]); r[9] = txt_(r[9]); r[11] = txt_(r[11]); });  // เขียนกลับต้องคงเป็นข้อความ
  sh.clearContents();
  sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
  if (out.length) sh.getRange(2, 1, out.length, SHEETS[name].length).setValues(out);
  sh.setFrozenRows(1);
}

function cellByHead_(raw, head, aliases) {
  for (var i = 0; i < head.length; i++) {
    var h = normHead_(head[i]);
    for (var a = 0; a < aliases.length; a++) if (h === normHead_(aliases[a])) return String(raw[i] == null ? '' : raw[i]).trim();
  }
  return '';
}
function isNumText_(v) { return /^-?\d+(?:\.\d+)?$/.test(String(v || '').replace(/,/g, '').trim()); }
function cleanNumText_(v) { return isNumText_(v) ? parseFloat(String(v).replace(/,/g, '')) : 0; }
function isDateText_(v) { return !!normDate_(v); }
function isPhoneText_(v) { return /^0?\d{8,10}$/.test(String(v || '').replace(/\D/g, '')); }

function normalizePositionRow_(raw, head) {
  var id = cellByHead_(raw, head, ['position_id','รหัส','รหัสตำแหน่ง']) || String(raw[0] || '').trim();
  var name = cellByHead_(raw, head, ['position_name','ตำแหน่ง','ชื่อตำแหน่ง']) || String(raw[1] || '').trim();
  if (!id && !name) return null;

  // ลบเฉพาะข้อมูลตัวอย่างเดิมที่ฝังมากับระบบ
  var demo = {PT001:'ผู้ช่วยบัญชี',PT002:'ผู้ช่วยภาษี',PT003:'ผู้ช่วย Payroll',PT004:'ผู้ช่วย OCR'};
  if (demo[id] && name === demo[id]) return null;

  var dep = cellByHead_(raw, head, ['department','แผนก','ฝ่าย']);
  var rate = cellByHead_(raw, head, ['hourly_rate','ค่าแรง/ชม.','ค่าแรงต่อชั่วโมง','ค่าจ้าง/ชม.']);
  var ot = cellByHead_(raw, head, ['ot_rate','OT/ชม.','โอที/ชม.']);
  var status = cellByHead_(raw, head, ['status','สถานะ']);
  var note = cellByHead_(raw, head, ['note','หมายเหตุ']);

  var strings = [], nums = [];
  raw.slice(2).forEach(function(v){
    v = String(v == null ? '' : v).trim(); if (!v) return;
    if (isNumText_(v)) nums.push(cleanNumText_(v)); else strings.push(v);
  });
  var statuses = ['ใช้งาน','ไม่ใช้งาน','หยุดใช้งาน','พักใช้งาน','ยกเลิก'];
  if (!status || statuses.indexOf(status) < 0) {
    for (var i=0;i<strings.length;i++) if (statuses.indexOf(strings[i]) >= 0) { status = strings[i]; break; }
  }
  if (!isNumText_(rate) || cleanNumText_(rate) <= 0) rate = nums.length ? nums[0] : 0;
  if (!isNumText_(ot) || cleanNumText_(ot) <= 0) ot = nums.length > 1 ? nums[1] : round2_(cleanNumText_(rate) * 1.5);

  // หาแผนก/หมายเหตุจากข้อความที่เหลือ กรณีหัวคอลัมน์เดิมสลับตำแหน่ง
  var rest = strings.filter(function(s){ return s !== status; });
  if (!dep || statuses.indexOf(dep) >= 0 || isNumText_(dep)) dep = rest.length ? rest[0] : '';
  if (!note || note === dep || statuses.indexOf(note) >= 0 || isNumText_(note)) note = rest.length > 1 ? rest[rest.length-1] : '';

  return [id, name || id, dep || '', cleanNumText_(rate), cleanNumText_(ot), status || 'ใช้งาน', note || ''];
}

function normalizeEmployeeRow_(raw, head) {
  var id = cellByHead_(raw, head, ['employee_id','รหัส','รหัสพนักงาน']) || String(raw[0] || '').trim();
  var name = cellByHead_(raw, head, ['employee_name','ชื่อ-สกุล','ชื่อ','ชื่อพนักงาน']) || String(raw[1] || '').trim();
  if (!id && !name) return null;

  // ลบเฉพาะข้อมูลตัวอย่างเดิมที่ระบบสร้างไว้
  if ((id === 'PT001' && name === 'สมชาย ใจดี') || (id === 'PT002' && (name === 'สมหญิง ดีมาก' || name === 'สมหญิง อินทา'))) return null;

  var type = cellByHead_(raw, head, ['type','ประเภท','ประเภทพนักงาน']);
  var rate = cellByHead_(raw, head, ['hourly_rate','ค่าแรง/ชม.','ค่าแรงต่อชั่วโมง']);
  var ot = cellByHead_(raw, head, ['ot_rate','OT/ชม.','โอที/ชม.']);
  var status = cellByHead_(raw, head, ['status','สถานะ']);
  var position = cellByHead_(raw, head, ['position','ตำแหน่ง','รหัสตำแหน่ง']);
  var start = cellByHead_(raw, head, ['start_date','วันเริ่มงาน','วันที่เริ่มงาน']);
  var phone = cellByHead_(raw, head, ['phone','เบอร์โทร','โทรศัพท์']);
  var nid = cellByHead_(raw, head, ['national_id','เลขประจำตัวประชาชน','เลขบัตรประชาชน','บัตรประชาชน']);
  var bank = cellByHead_(raw, head, ['bank_name','ธนาคาร','ชื่อธนาคาร']);
  var bankAcc = cellByHead_(raw, head, ['bank_account','เลขที่บัญชี','เลขบัญชี','บัญชีธนาคาร']);

  var nums = [], strings = [];
  raw.slice(2).forEach(function(v){
    v = String(v == null ? '' : v).trim(); if (!v) return;
    if (isDateText_(v)) { if (!start) start = normDate_(v); return; }
    if (isPhoneText_(v) && String(v).replace(/\D/g,'').length >= 9) { if (!phone) phone = String(v).replace(/\D/g,''); return; }
    // เลขบัตร ปชช. 13 หลัก — กันไม่ให้หลุดไปเป็นค่าแรง/OT ตอนเดาคอลัมน์
    if (/^\d{13}$/.test(v.replace(/[-\s]/g, ''))) { if (!nid) nid = v; return; }
    if (isNumText_(v)) nums.push(cleanNumText_(v)); else strings.push(v);
  });
  if (!isNumText_(rate) || cleanNumText_(rate) <= 0) rate = nums.length ? nums[0] : 0;
  if (!isNumText_(ot) || cleanNumText_(ot) <= 0) ot = nums.length > 1 ? nums[1] : round2_(cleanNumText_(rate) * 1.5);
  if (!type) type = strings.filter(function(s){return /part|full|รายชั่วโมง|ชั่วคราว/i.test(s);})[0] || 'Part-time';
  if (!status) status = strings.filter(function(s){return ['ทำงาน','ลาออก','พักงาน','ยกเลิก'].indexOf(s)>=0;})[0] || 'ทำงาน';
  if (!position) position = strings.filter(function(s){return /^POS\d+$/i.test(s);})[0] || '';
  if (/^POS\d{2}$/i.test(position)) position = 'POS0' + position.slice(3); // POS01 → POS001
  start = normDate_(start) || start || '';
  phone = String(phone || '').replace(/\D/g,'');
  if (phone && phone.charAt(0) !== '0' && phone.length === 9) phone = '0' + phone;

  return [id, name || id, type || 'Part-time', cleanNumText_(rate), cleanNumText_(ot), status || 'ทำงาน', position || '', start, phone,
    nid || '', bank || '', bankAcc || ''];
}

/** เรียกจากหน้าเว็บหรือ Run เองครั้งเดียว เพื่อซ่อม Master ที่มีอยู่ทันที */
function apiRepairMaster(pin) {
  guard_(pin);
  var ss = getSs_();
  repairMasterSheets_(ss);
  return {ok:true, positions:readSheet_('Positions').length, employees:readSheet_('Employees').length};
}

// (เทมเพลต) ไม่มีชื่อบริษัทตั้งต้นให้เปลี่ยน — คงชื่อฟังก์ชันไว้เพราะ migrate_ เรียกอยู่
function renameStaleCompany_() { }

/** รันครั้งเดียว — สร้างโฟลเดอร์+ชีต (ไม่มีรหัสผ่านแล้ว: ใครมีลิงก์ /exec ก็เข้าใช้ได้) */
function setup() {
  getRoot_(); var ss = getSs_(); initSheets_(ss); migrate_(); props_().setProperty('SCHEMA_V', SCHEMA_V);
  props_().deleteProperty('ACCESS_PIN');   // ล้างรหัสเดิมทิ้ง เผื่อเคยตั้งไว้ก่อนถอดระบบล็อก
  return 'พร้อมแล้ว! เปิดหน้าเว็บได้เลย (ไม่ต้องใส่รหัส)\nSheet: ' + ss.getUrl();
}

/* ================= ล็อกด้วยรหัส ================= */
/* ระบบล็อกด้วยรหัสผ่านถูกถอดออกแล้ว (ทุกคนที่มีลิงก์ /exec เข้าใช้ได้)
   คง guard_() ไว้เป็น no-op เพื่อไม่ต้องแก้ลายเซ็น api ทุกตัว (ทุกตัวยังรับ pin เป็นพารามิเตอร์แรก)
   ความปลอดภัยที่เหลือ = Web app เป็นแบบ Execute as owner + ลิงก์รู้กันเฉพาะภายใน */
function guard_(pin) { }

/* ================= รับข้อมูลจากบอท (doPost) ================= */
function doPost(e) {
  var out = { ok: false };
  try {
    var body = JSON.parse(e.postData.contents);

    /* ----- ช่องทางบอท: พนักงานขอสลิปของตัวเอง -----
       เส้นเดิม (ใบลงเวลา) ไม่มี action → ต้องตกลงไปทางเดิมเสมอ ห้ามเปลี่ยนพฤติกรรมของเดิม
       เส้นใหม่ต้องมีรหัสลับทุกครั้ง เพราะ URL /exec นี้ใครยิงก็ได้ */
    if (body.action) {
      if (!botAuthOk_(body.secret)) return jsonOut_({ ok: false, reason: 'unauthorized' });
      if (body.action === 'payslip_list') return jsonOut_(botPayslipList_(body.lineUserId, { month: body.month }));
      if (body.action === 'payslip_file') return jsonOut_(botPayslipFile_(body.lineUserId, body.from, body.to));
      return jsonOut_({ ok: false, reason: 'unknown_action' });
    }

    // ฟอร์ม TS-01 1 ใบ = หลายวัน → บอทส่ง rows[] มา; ถ้าไม่มี ใช้ fields เดิม (1 แถว) เหมือนเดิม
    var rows = (body.rows && body.rows.length) ? body.rows : [body.fields || {}];
    var head = rows[0] || {};
    var ss = getSs_(); var att = ss.getSheetByName('Attendance');
    var who = String(head.employee_id || head.employee_name || 'guest').replace(/[^\w฀-๿]+/g, '_').slice(0, 24) || 'guest';
    var stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
    var base = who + '_' + stamp;
    var fileUrl = '';
    if (body.image) {
      var ct = body.contentType || 'image/jpeg';
      var ext = ct.indexOf('png') >= 0 ? 'png' : (ct.indexOf('pdf') >= 0 ? 'pdf' : 'jpg');
      var file = subFolder_('01_Inbox').createFile(Utilities.newBlob(Utilities.base64Decode(body.image), ct, base + '.' + ext));
      fileUrl = file.getUrl();
      subFolder_('01_Inbox').createFile(Utilities.newBlob(JSON.stringify(body, null, 2), 'application/json', base + '_ocr.json'));
    }
    out = appendAttendanceRows_(ss, att, body, rows, head, fileUrl);
  } catch (err) {
    out = { ok: false, error: String(err) };
    try { getSs_().getSheetByName('Import_Log').appendRow([new Date().toISOString(), '', '', '', 'ผิดพลาด', String(err)]); } catch (e2) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** เขียนแถวใบลงเวลาลงชีต Attendance + Import_Log — "เส้นเดียว" ที่ทั้ง doPost (บอท LINE) และหน้าอัปรูปใบลงเวลา
 *  (apiSaveTimesheetRows) ใช้ร่วมกัน · แยกออกมาจาก doPost ตรง ๆ ทุกบรรทัด ห้ามเปลี่ยนพฤติกรรม (บอทยิง rows[] เข้ามาทางนี้)
 *  พิสูจน์ด้วย scripts/harness-dopost.js (เทียบกับ _backup_clone/Code.before-round1b.gs)
 *  rows = แถวที่จะเขียน (อย่างน้อย 1) · head = rows[0] · fileUrl = ลิงก์รูปต้นฉบับใน Drive ('' ถ้าไม่มี) */
function appendAttendanceRows_(ss, att, body, rows, head, fileUrl) {
  // ทุกแถวจากรูปเดียวกันชี้ไฟล์ต้นฉบับเดียวกัน → กดดูรูปจากแถวไหนก็เปิดฟอร์มใบเดิม
  var recvAt = body.receivedAt || new Date().toISOString();
  var ids = [];
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var futureCount = 0;
  for (var i = 0; i < rows.length; i++) {
    var f = rows[i] || {};
    var workDate = normDate_(f.work_date) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    // วันทำงานล่วงหน้า = OCR อ่านวัน/เดือนพลาด (การ์ดลงเวลาย่อมเป็นวันที่ผ่านมาแล้ว)
    // ไม่แก้วันที่ให้เอง เพราะเดาไม่ได้ว่าที่ถูกคือวันไหน — ติดหมายเหตุไว้ให้คนตรวจตัดสิน
    var futureNote = workDate > today ? ' ⚠ วันที่ล่วงหน้า — ต้องตรวจสอบ' : '';
    if (futureNote) futureCount++;
    var id = 'ATT' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    att.appendRow([id, recvAt, body.lineUserId || '', body.displayName || '',
      f.employee_id || head.employee_id || '', f.employee_name || head.employee_name || '',
      workDate, f.check_in || '', f.check_out || '',
      num_(f.break_minutes), num_(f.normal_hours), num_(f.ot_hours), 'รอตรวจสอบ', fileUrl, (body.summary || '') + futureNote]);
    ids.push(id);
  }
  var note = (body.summary || '') + (rows.length > 1 ? ' [' + rows.length + ' แถว]' : '')
    + (futureCount ? ' ⚠ วันที่ล่วงหน้า ' + futureCount + ' แถว — ต้องตรวจสอบ' : '');
  ss.getSheetByName('Import_Log').appendRow([new Date().toISOString(), body.lineUserId || '', body.displayName || '', fileUrl, 'สำเร็จ', note]);
  return { ok: true, id: ids[0], ids: ids, count: ids.length, fileUrl: fileUrl };
}

/* ================= หน้าเว็บ: อัปรูปใบลงเวลา → AI อ่าน → ตรวจแก้ → บันทึก =================
   ความปลอดภัย: เว็บเปิด ANYONE_ANONYMOUS + รันสิทธิ์เจ้าของ → ใครก็เรียก 2 api นี้ได้
   - รับเฉพาะรูป JPEG/PNG จริง (เช็ค magic bytes) + จำกัดขนาด · รหัสพนักงานต้องมีในทะเบียน
   - ห้ามส่งลิงก์/URL ของ Drive กลับหน้าเว็บ (คนดูไม่มีสิทธิ์ในไดรฟ์) → ส่งแค่ file_id ไว้อ้างอิงตอนบันทึก
   - ตอนบันทึก file_id ต้องเป็นไฟล์ในโฟลเดอร์ 01_Inbox จริง ไม่งั้นใครเดา id ไฟล์อื่นในไดรฟ์ได้ก็ผูกลิงก์มั่วเข้าชีตได้ */
var WEB_TS_IMAGE_MAX = 8 * 1024 * 1024;   // หน้าเว็บย่อรูปเหลือ ~0.3–1 MB แล้ว — เผื่อไว้ แต่กันยิงไฟล์ยักษ์มาถมไดรฟ์
var WEB_TS_MAX_ROWS = 62;                  // ใบลงเวลา 1 แผ่น = 16 แถว · งวดยาวสุด ~31 วัน → เผื่อ 2 เท่า

function findEmployee_(id) {
  id = String(id || '').trim();
  if (!id) return null;
  var list = readSheet_('Employees');
  for (var i = 0; i < list.length; i++) if (String(list[i].employee_id).trim() === id) return list[i];
  return null;
}
function inFolder_(file, folderName) {
  var target = subFolder_(folderName).getId();
  var ps = file.getParents();
  while (ps.hasNext()) { if (ps.next().getId() === target) return true; }
  return false;
}
function webTsStr_(v, max) { return String(v == null ? '' : v).slice(0, max || 40); }

/** อ่านรูปใบลงเวลาด้วย AI
 *  payload = {employee_id, image_b64, mime:'image/jpeg'|'image/png', period_from, period_to}
 *  1) เก็บรูปเข้า Drive 01_Inbox (แบบเดียวกับ doPost) 2) เรียก ocrTimesheetImage_ (TimesheetOcr.gs)
 *  คืน JSON string: {ok, file_id, employee, model, header, rows[], warnings[]} หรือ {ok:false, error, file_id?} */
function apiOcrTimesheet(pin, payload) {
  guard_(pin); payload = payload || {};
  var emp = findEmployee_(payload.employee_id);
  if (!emp) return JSON.stringify({ ok: false, error: 'ไม่พบรหัสพนักงานนี้ในทะเบียน — เลือกพนักงานใหม่' });
  var b64 = String(payload.image_b64 || '').replace(/^data:[^,]*,/, '');
  var mime = payload.mime === 'image/png' ? 'image/png' : 'image/jpeg';
  var bytes;
  try { bytes = Utilities.base64Decode(b64); } catch (e) { return JSON.stringify({ ok: false, error: 'ไฟล์รูปเสียหาย ส่งใหม่อีกครั้ง' }); }
  if (!bytes || !bytes.length) return JSON.stringify({ ok: false, error: 'ไม่พบรูปใบลงเวลา' });
  if (bytes.length > WEB_TS_IMAGE_MAX) return JSON.stringify({ ok: false, error: 'รูปใหญ่เกิน ' + (WEB_TS_IMAGE_MAX / 1048576) + ' MB' });
  var b0 = bytes[0] & 255, b1 = bytes[1] & 255, b2 = bytes[2] & 255, b3 = bytes[3] & 255;
  var isJpeg = b0 === 0xFF && b1 === 0xD8 && b2 === 0xFF, isPng = b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47;
  if (!isJpeg && !isPng) return JSON.stringify({ ok: false, error: 'ไฟล์นี้ไม่ใช่รูป JPG/PNG' });
  mime = isPng ? 'image/png' : 'image/jpeg';
  var pFrom = normDate_(payload.period_from), pTo = normDate_(payload.period_to);

  // เก็บรูปต้นฉบับไว้ก่อนเสมอ (อ่านไม่สำเร็จก็ยังมีรูปให้คนกรอกเอง/ย้อนดูได้) — ตั้งชื่อแบบเดียวกับ doPost + นำหน้า WEB_
  var who = String(emp.employee_id).replace(/[^\w฀-๿]+/g, '_').slice(0, 24) || 'guest';
  var base = 'WEB_' + who + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  var inbox = subFolder_('01_Inbox');
  var file = inbox.createFile(Utilities.newBlob(bytes, mime, base + (isPng ? '.png' : '.jpg')));

  var res;
  if (typeof ocrTimesheetImage_ !== 'function') {
    res = { ok: false, error: 'ยังไม่ได้ติดตั้งตัวอ่านใบลงเวลา (ไฟล์ TimesheetOcr.gs) — กรอกเองในตารางได้' };
  } else {
    try { res = ocrTimesheetImage_(b64, mime, { period_from: pFrom, period_to: pTo }); }
    catch (e) { res = { ok: false, error: 'AI อ่านรูปไม่สำเร็จ: ' + String(e && e.message || e) }; }
  }
  if (!res || typeof res !== 'object') res = { ok: false, error: 'ตัวอ่านใบลงเวลาไม่ส่งผลกลับมา' };
  try { inbox.createFile(Utilities.newBlob(JSON.stringify({ source: 'web', employee_id: emp.employee_id, period_from: pFrom, period_to: pTo, result: res }, null, 2), 'application/json', base + '_ocr.json')); } catch (e) {}

  var out = { ok: !!res.ok, file_id: file.getId(), employee: { employee_id: emp.employee_id, employee_name: emp.employee_name || '' } };
  if (!res.ok) { out.error = String(res.error || 'อ่านรูปไม่สำเร็จ'); return JSON.stringify(out); }
  var h = res.header || {};
  out.model = webTsStr_(res.model, 60);
  out.header = { name: webTsStr_(h.name, 120), employee_id: webTsStr_(h.employee_id, 20), period_from: webTsStr_(h.period_from, 10), period_to: webTsStr_(h.period_to, 10) };
  out.rows = (res.rows || []).slice(0, WEB_TS_MAX_ROWS).map(function (r) {
    r = r || {};
    return { work_date: webTsStr_(r.work_date, 10), check_in: webTsStr_(r.check_in, 5), check_out: webTsStr_(r.check_out, 5),
      raw_date: webTsStr_(r.raw_date, 30), raw_in: webTsStr_(r.raw_in, 20), raw_out: webTsStr_(r.raw_out, 20),
      unreadable: !!r.unreadable, future: !!r.future, note: webTsStr_(r.note, 200) };
  });
  out.warnings = (res.warnings || []).slice(0, 20).map(function (w) { return webTsStr_(w, 300); });
  return JSON.stringify(out);
}

/** บันทึกแถวที่ตรวจแก้แล้วลงชีต Attendance — ผ่าน appendAttendanceRows_ เส้นเดียวกับ doPost ของบอท
 *  payload = {employee_id, file_id, period_from, period_to, rows:[{work_date, check_in, check_out}]}
 *  ไม่รับ พัก/ชม.ปกติ/OT จากหน้าเว็บ (ส่งว่างเหมือนบอท) → calcRowHours_ คำนวณเองจากเวลาเข้า-ออก + Settings
 *  คืน JSON string {ok, count, ids} — ไม่ส่ง fileUrl กลับ */
function apiSaveTimesheetRows(pin, payload) {
  guard_(pin); payload = payload || {};
  var emp = findEmployee_(payload.employee_id);
  if (!emp) return JSON.stringify({ ok: false, error: 'ไม่พบรหัสพนักงานนี้ในทะเบียน' });
  var src = payload.rows || [];
  if (!src.length) return JSON.stringify({ ok: false, error: 'ไม่มีแถวให้บันทึก' });
  if (src.length > WEB_TS_MAX_ROWS) return JSON.stringify({ ok: false, error: 'แถวเยอะเกิน ' + WEB_TS_MAX_ROWS + ' แถวต่อครั้ง' });
  var rows = [], bad = [];
  src.forEach(function (r, i) {
    r = r || {};
    var d = normDate_(r.work_date), ci = r.check_in ? hhmmStrict_(r.check_in) : '', co = r.check_out ? hhmmStrict_(r.check_out) : '';
    if (!d || ci === null || co === null || (!ci && !co)) { bad.push(i + 1); return; }
    rows.push({ employee_id: String(emp.employee_id), employee_name: emp.employee_name || '', work_date: d, check_in: ci, check_out: co,
      break_minutes: '', normal_hours: '', ot_hours: '' });   // รูปแบบเดียวกับแถวที่บอทส่ง (vision.js cleanRows)
  });
  if (bad.length) return JSON.stringify({ ok: false, error: 'แถวที่ ' + bad.join(', ') + ' วันที่หรือเวลาไม่ถูกต้อง' });

  var fileUrl = '';
  if (payload.file_id) {
    var file = null;
    try { file = DriveApp.getFileById(String(payload.file_id)); } catch (e) {}
    if (!file || !inFolder_(file, '01_Inbox')) return JSON.stringify({ ok: false, error: 'ไม่พบรูปใบลงเวลาที่อัปไว้ — อัปรูปใหม่อีกครั้ง' });
    fileUrl = file.getUrl();   // เก็บในชีตเท่านั้น (หน้า "บันทึกเวลา" เปิดรูปได้เหมือนแถวจากบอท) ไม่ส่งกลับหน้าเว็บ
  }
  var body = {
    lineUserId: '', displayName: 'หน้าเว็บ · อัปรูปใบลงเวลา',
    summary: 'อัปรูปใบลงเวลาผ่านหน้าเว็บ · ' + (emp.employee_name || emp.employee_id),
    rows: rows, receivedAt: new Date().toISOString(),
  };
  var ss = getSs_(); var att = ss.getSheetByName('Attendance');
  var out = appendAttendanceRows_(ss, att, body, rows, rows[0], fileUrl);
  return JSON.stringify({ ok: true, count: out.count, ids: out.ids });
}
/** "8:05" / "08.05" / "0805" → "08:05" · ว่าง → '' · รูปแบบผิด/เวลาเป็นไปไม่ได้ → null */
function hhmmStrict_(v) {
  var t = String(v == null ? '' : v).trim().replace(/\s+/g, '');
  if (!t) return '';
  var h, m, x = t.match(/^(\d{1,2})[:.](\d{2})$/);
  if (x) { h = +x[1]; m = +x[2]; }
  else if (/^\d{3,4}$/.test(t)) { h = +t.slice(0, t.length - 2); m = +t.slice(-2); }
  else return null;
  if (h > 23 || m > 59) return null;
  return pad_(h) + ':' + pad_(m);
}

/* ================= หน้าเว็บ (doGet) ================= */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบเงินเดือน' + (readSettings_().company_name ? ' · ' + readSettings_().company_name : ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= API ================= */
function apiBootstrap(pin) {
  guard_(pin);
  if (props_().getProperty('SCHEMA_V') !== SCHEMA_V) { migrate_(); props_().setProperty('SCHEMA_V', SCHEMA_V); }
  var ss = getSs_();
  // ส่งเป็น JSON string เพื่อป้องกัน google.script.run คืนค่า null/undefined
  // เมื่อ object มีข้อมูลจำนวนมากหรือมีค่าที่ serialize ไม่สมบูรณ์
  return JSON.stringify({
    ok: true,
    settings: readSettings_(),
    employees: readSheet_('Employees'),
    positions: readSheet_('Positions'),
    ssUrl: ss.getUrl(),
    logo: getLogo_(),
    today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')
  });
}

/* ---- บันทึก + ส่งข้อมูลล่าสุดกลับในคำขอเดียว (หน้าเว็บเรียกผ่าน run() ให้เอง) ----
   เดิมกดบันทึก 1 ครั้ง = ไป Google 2 รอบ (บันทึก ~2 วิ แล้วโหลดหน้าใหม่อีก 2.5–4 วิ) → ผู้ใช้รอ ~5 วิ
   apiDo ทำคำสั่งเขียนแล้วอ่านข้อมูลที่หน้านั้นต้องใช้ต่อทันทีในรอบเดียวกัน · อนุญาตเฉพาะชื่อที่อยู่ในรายการ */
var API_DO_WRITE_ = { apiSaveEmployee: 1, apiDeleteEmployee: 1, apiDeleteEmployees: 1, apiSavePosition: 1, apiDeletePosition: 1,
  apiDeletePositions: 1, apiSaveSchedule: 1, apiDeleteSchedule: 1, apiSaveAdjustment: 1, apiDeleteAdjustment: 1,
  apiSetStatus: 1, apiUpdateAttendance: 1 };
var API_DO_READ_ = { apiBootstrap: 1, apiPositions: 1, apiSchedule: 1, apiAdjustments: 1, apiAttendance: 1, apiDashboard: 1 };
function apiDo(pin, fn, args, fresh) {
  guard_(pin);
  var G = (typeof globalThis !== 'undefined') ? globalThis : this;
  if (!API_DO_WRITE_[fn] || typeof G[fn] !== 'function') throw new Error('คำสั่งนี้ใช้ผ่าน apiDo ไม่ได้: ' + fn);
  var result = G[fn].apply(null, [pin].concat(args || []));
  var out = {};
  (fresh || []).forEach(function (f) {
    if (!f || !API_DO_READ_[f[0]] || typeof G[f[0]] !== 'function') return;
    try { out[f[0] + '|' + JSON.stringify(f.slice(1))] = G[f[0]].apply(null, [pin].concat(f.slice(1))); } catch (e) {}
  });
  return JSON.stringify({ result: result, fresh: out });
}

/** โหลดค่าตั้งต้น: ล้างตำแหน่งเดิม → ใส่ DEFAULT_POSITIONS + ตั้งชื่อบริษัท
 *  มีไว้เพราะ seed ปกติจะไม่ทับชีตที่มีข้อมูลแล้ว (กันข้อมูลจริงหาย) — อันนี้ผู้ใช้สั่งทับเอง
 *  ไม่แตะ Employees / Attendance / Adjustments */
function apiSeedDefaults(pin) {
  guard_(pin);
  var ss = getSs_();
  var pos = ss.getSheetByName('Positions');
  var last = pos.getLastRow();
  // deleteRows (ไม่ใช่ clearContent) — clearContent ทิ้งแถวว่างไว้ getLastRow ยังนับอยู่ seed จะไม่ทำงาน
  if (last > 1) pos.deleteRows(2, last - 1);
  DEFAULT_POSITIONS.forEach(function (p) { pos.appendRow(p); });
  return { ok: true, positions: readSheet_('Positions'), settings: readSettings_() };
}

var TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

function apiDashboard(pin) {
  guard_(pin);
  var rows = readSheet_('Attendance'); var emps = readSheet_('Employees'); var adjs = readSheet_('Adjustments');
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); var year = today.slice(0, 4);
  var rateMap = {}; emps.forEach(function (e) { rateMap[e.employee_id] = e; });
  var hours = 0, ot = 0, base = 0, otPay = 0;
  var monthly = []; for (var i = 0; i < 12; i++) monthly.push(0);
  rows.forEach(function (r) {
    var h = calcRowHours_(r); hours += h.normal + h.ot; ot += h.ot;
    var e = rateMap[r.employee_id]; var rate = e ? num_(e.hourly_rate) : 0; var otr = e ? otRate_(e) : 0;
    var pay = h.normal * rate + h.ot * otr; base += h.normal * rate; otPay += h.ot * otr;
    if (String(r.work_date).slice(0, 4) === year) { var mi = parseInt(String(r.work_date).slice(5, 7), 10) - 1; if (mi >= 0 && mi < 12) monthly[mi] += pay; }
  });
  var addition = 0, bonus = 0, deduction = 0;
  adjs.forEach(function (a) { addition += num_(a.addition); bonus += num_(a.bonus); deduction += num_(a.deduction); });
  var payEstimate = base + otPay + addition + bonus - deduction;
  var st = { 'รอตรวจสอบ': 0, 'ตรวจสอบแล้ว': 0, 'อนุมัติแล้ว': 0, 'ผิดพลาด': 0 };
  rows.forEach(function (r) { if (st[r.status] != null) st[r.status]++; });
  var todays = rows.filter(function (r) { return r.work_date === today; });
  var recent = rows.slice(-6).reverse().map(function (r) {
    return { work_date: r.work_date, employee_name: r.employee_name || r.display_name || '-', employee_id: r.employee_id || '', status: r.status };
  });
  var cost = [
    { label: 'ค่าจ้างพื้นฐาน', value: round2_(base) }, { label: 'ค่าโอที', value: round2_(otPay) },
    { label: 'เงินเพิ่ม', value: round2_(addition) }, { label: 'โบนัส', value: round2_(bonus) },
  ].filter(function (x) { return x.value > 0; });
  return {
    employees: emps.length,
    workedToday: unique_(todays.map(function (r) { return r.employee_id || r.display_name; })).length,
    hoursTotal: round2_(hours), otTotal: round2_(ot), payEstimate: round2_(payEstimate),
    avgHours: emps.length ? round2_(hours / emps.length) : 0, avgPay: emps.length ? round2_(payEstimate / emps.length) : 0,
    slipsIssued: parseInt(props_().getProperty('SLIP_SEQ') || '0', 10),
    pending: st['รอตรวจสอบ'], verified: st['ตรวจสอบแล้ว'] + st['อนุมัติแล้ว'], error: st['ผิดพลาด'], records: rows.length,
    trend: monthly.map(function (v, i) { return { m: TH_MON[i], total: round2_(v) }; }),
    cost: cost, recent: recent, year: year,
  };
}

function apiAttendance(pin, filter) {
  guard_(pin); filter = filter || {};
  var emps = readSheet_('Employees'); var empIds = {}; emps.forEach(function (e) { empIds[e.employee_id] = e; });
  var all = readSheet_('Attendance');
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  // นับคู่ (พนักงาน+วัน) เพื่อตรวจซ้ำ
  var seen = {};
  all.forEach(function (r) { var k = (r.employee_id || r.display_name) + '|' + r.work_date; seen[k] = (seen[k] || 0) + 1; });
  return all.reverse().filter(function (r) {
    if (filter.q) { var q = String(filter.q).toLowerCase();
      if ((r.employee_id + ' ' + r.employee_name + ' ' + r.display_name).toLowerCase().indexOf(q) < 0) return false; }
    if (filter.status && r.status !== filter.status) return false;
    if (filter.from && r.work_date < filter.from) return false;
    if (filter.to && r.work_date > filter.to) return false;
    return true;
  }).map(function (r) {
    var h = calcRowHours_(r); r.normal_calc = h.normal; r.ot_calc = h.ot;
    var w = [];
    if (h.normal + h.ot > 16) w.push('ชั่วโมงเกิน 16');
    var ci = toMin_(r.check_in), co = toMin_(r.check_out);
    if (ci != null && co != null && co <= ci) w.push('เวลาออก ≤ เวลาเข้า');
    if (!r.employee_id || !empIds[r.employee_id]) w.push('ไม่พบรหัสพนักงาน');
    if (seen[(r.employee_id || r.display_name) + '|' + r.work_date] > 1) w.push('วันซ้ำ');
    if (r.employee_id && empIds[r.employee_id] && num_(empIds[r.employee_id].hourly_rate) <= 0) w.push('ยังไม่ตั้งค่าแรง');
    // การ์ดลงเวลาย่อมเป็นวันที่ผ่านมาแล้ว — ล่วงหน้า = OCR อ่านวัน/เดือนพลาด ต้องให้คนยืนยันก่อน
    if (r.work_date && r.work_date > today) w.push('วันที่ล่วงหน้า');
    r.warnings = w;
    return r;
  });
}

function apiSetStatus(pin, id, status) { guard_(pin); return updateAtt_(id, { status: status }); }

function apiUpdateAttendance(pin, id, patch) { guard_(pin); return updateAtt_(id, patch); }

function updateAtt_(id, patch) {
  var sh = getSs_().getSheetByName('Attendance');
  var data = sh.getDataRange().getValues(); var head = data[0];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      Object.keys(patch).forEach(function (k) {
        var c = head.indexOf(k); if (c >= 0) {
          var v = patch[k];
          if (k === 'work_date') v = normDate_(v) || v;
          if (['break_minutes', 'normal_hours', 'ot_hours'].indexOf(k) >= 0) v = num_(v);
          sh.getRange(i + 1, c + 1).setValue(v);
        }
      });
      return { ok: true };
    }
  }
  return { ok: false };
}

/* ================= นำเข้าจาก Excel / CSV =================
   .xlsx อ่านตรง ๆ ใน Apps Script ไม่ได้ → อัปขึ้น Drive ให้ Google แปลงเป็นสเปรดชีตก่อน แล้วอ่าน แล้วลบทิ้ง
   (ใช้ Drive API ผ่าน UrlFetchApp + โทเคนของสคริปต์เอง — ไม่ต้องเปิด Advanced Service)
   วิธีนี้รับ .csv ได้ด้วยเพราะ Drive แปลงให้เหมือนกัน */

// ชื่อคอลัมน์ที่ยอมรับ → คีย์จริงในชีต (รับทั้งไทย/อังกฤษ เผื่อไฟล์มาจากหลายที่)
var IMPORT_ALIAS = {
  positions: {
    position_id: ['position_id', 'รหัส', 'รหัสตำแหน่ง'],
    position_name: ['position_name', 'ตำแหน่ง', 'ชื่อตำแหน่ง'],
    department: ['department', 'แผนก', 'ฝ่าย', 'dept'],
    hourly_rate: ['hourly_rate', 'ค่าแรง', 'ค่าแรงชม', 'ค่าแรงตอชม', 'อัตรา'],
    ot_rate: ['ot_rate', 'ot', 'otชม', 'otตอชม', 'คาลวงเวลา'],
    status: ['status', 'สถานะ'],
    note: ['note', 'หมายเหตุ', 'รายละเอียด'],
  },
  employees: {
    employee_id: ['employee_id', 'รหัส', 'รหัสพนักงาน'],
    employee_name: ['employee_name', 'ชื่อ', 'ชื่อสกุล', 'ชื่อพนักงาน'],
    position: ['position', 'ตำแหน่ง', 'รหัสตำแหน่ง'],
    hourly_rate: ['hourly_rate', 'ค่าแรง', 'ค่าแรงชม', 'ค่าแรงตอชม'],
    ot_rate: ['ot_rate', 'ot', 'otชม', 'otตอชม'],
    start_date: ['start_date', 'วันเริ่มงาน', 'เริ่มงาน'],
    phone: ['phone', 'เบอร์โทร', 'เบอร์', 'โทร'],
    national_id: ['national_id', 'เลขบัตรประชาชน', 'เลขประจำตัวประชาชน', 'บัตรประชาชน', 'เลขบัตร'],
    // ไฟล์ต้นแบบใส่ "ธนาคารกรุงไทย | 154-0-58424-6" มาช่องเดียว → แยกเป็นชื่อธนาคาร/เลขบัญชีตอน upsert
    bank: ['bank', 'บัญชีธนาคาร', 'ธนาคารบัญชี'],
    bank_name: ['bank_name', 'ธนาคาร', 'ชื่อธนาคาร'],
    bank_account: ['bank_account', 'เลขที่บัญชี', 'เลขบัญชี', 'บัญชี'],
    status: ['status', 'สถานะ'],
  },
};

/** "ธนาคารกรุงไทย | 154-0-58424-6" → {name:'ธนาคารกรุงไทย', acc:'154-0-58424-6'}
 *  รองรับตัวคั่น | / - เว้นวรรค และกรณีเขียนสลับข้าง (เลขบัญชีมาก่อน) */
function splitBank_(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return { name: '', acc: '' };
  var parts = s.split(/\s*[|/·]\s*|\s{2,}/).filter(function (x) { return x !== ''; });
  if (parts.length < 2) {
    // ไม่มีตัวคั่น: ตัดเอาก้อนตัวเลข-ขีดยาว ๆ ท้ายสุดเป็นเลขบัญชี ที่เหลือเป็นชื่อธนาคาร
    var m = s.match(/([\d\-\s]{8,})$/);
    if (!m) return { name: s, acc: '' };
    return { name: s.slice(0, m.index).trim(), acc: m[1].trim() };
  }
  var isAcc = function (x) { return /^[\d\-\s]+$/.test(x); };
  if (isAcc(parts[0]) && !isAcc(parts[1])) return { name: parts[1].trim(), acc: parts[0].trim() };
  return { name: parts[0].trim(), acc: parts.slice(1).join(' ').trim() };
}

// เทียบหัวคอลัมน์แบบหลวม ๆ: "ค่าแรง/ชม. (บาท)" → "ค่าแรงชม"
// ต้องตัดข้อความในวงเล็บก่อน แล้วค่อยตัดเครื่องหมาย — สลับลำดับจะเหลือ "ค่าแรงชมบาท" แล้วจับคู่ไม่ติด
function normHead_(v) {
  return String(v == null ? '' : v).toLowerCase()
    .replace(/\([^)]*\)/g, '')
    .replace(/[\s./_\-()฿:]/g, '');
}

/** อัปไฟล์ขึ้น Drive แบบสั่งแปลงเป็น Google Sheet → คืน fileId */
function toSheet_(blob, name) {
  var b = '----payrollimport' + Date.now();
  var meta = { name: name || 'import', mimeType: MimeType.GOOGLE_SHEETS };
  var head = Utilities.newBlob('--' + b + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(meta) + '\r\n--' + b + '\r\nContent-Type: ' + (blob.getContentType() || 'application/octet-stream') + '\r\n\r\n').getBytes();
  var tail = Utilities.newBlob('\r\n--' + b + '--\r\n').getBytes();
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + b,
    payload: head.concat(blob.getBytes()).concat(tail),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) throw new Error('แปลงไฟล์ไม่สำเร็จ (' + res.getResponseCode() + ') — ไฟล์อาจไม่ใช่ Excel/CSV');
  var id = JSON.parse(res.getContentText()).id;
  if (!id) throw new Error('แปลงไฟล์แล้วแต่ไม่ได้ไฟล์กลับมา');
  return id;
}

/** อ่านตาราง (แผ่นแรก) จากไฟล์ที่แปลงแล้ว → [{คีย์: ค่า}] ตาม alias ของ kind */
function readImport_(fileId, kind) {
  var alias = IMPORT_ALIAS[kind];
  var ss = SpreadsheetApp.openById(fileId);
  var wanted = kind === 'positions' ? ['Positions', 'Position', 'ตำแหน่งงาน'] : ['Employees', 'Employee', 'พนักงาน'];
  var sheet = null;
  wanted.some(function (n) { sheet = ss.getSheetByName(n); return !!sheet; });
  if (!sheet) sheet = ss.getSheets()[0];
  // ใช้ค่าที่แสดงในเซลล์ เพื่อรักษาเลข 0 นำหน้าในเบอร์โทรและรูปแบบวันที่จาก Excel
  var vals = sheet.getDataRange().getDisplayValues();
  if (vals.length < 2) return { rows: [], headErr: 'ไฟล์ว่าง หรือมีแต่หัวตาราง' };
  // หาว่าหัวตารางแต่ละคอลัมน์ตรงกับคีย์ไหน
  var head = vals[0].map(normHead_), col = {};
  Object.keys(alias).forEach(function (key) {
    for (var i = 0; i < head.length; i++) {
      if (!head[i]) continue;
      if (alias[key].some(function (a) { return normHead_(a) === head[i]; })) { col[key] = i; return; }
    }
  });
  var idKey = kind === 'positions' ? 'position_id' : 'employee_id';
  var nameKey = kind === 'positions' ? 'position_name' : 'employee_name';
  if (col[idKey] == null && col[nameKey] == null) {
    return { rows: [], headErr: 'ไม่พบคอลัมน์ "' + (kind === 'positions' ? 'รหัส หรือ ชื่อตำแหน่ง' : 'รหัส หรือ ชื่อ-สกุล') + '" ในแถวแรก (แถวแรกต้องเป็นหัวตาราง)' };
  }
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    var o = {}, any = false;
    Object.keys(col).forEach(function (k) {
      var v = vals[r][col[k]];
      v = String(v == null ? '' : v).trim();
      o[k] = v; if (v) any = true;
    });
    if (any) rows.push(o);
  }
  return { rows: rows, cols: Object.keys(col) };
}

/** payload = {kind:'positions'|'employees', name, mimeType, base64} */
function apiImport(pin, payload) {
  guard_(pin);
  var kind = payload && payload.kind;
  if (!IMPORT_ALIAS[kind]) return { ok: false, error: 'ประเภทข้อมูลไม่ถูกต้อง' };
  var tmpId = null;
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), payload.mimeType || 'application/octet-stream', payload.name || 'import');
    tmpId = toSheet_(blob, 'IMPORT_TMP_' + kind);
    var got = readImport_(tmpId, kind);
    if (got.headErr) return { ok: false, error: got.headErr };
    if (!got.rows.length) return { ok: false, error: 'ไม่พบข้อมูลในไฟล์' };
    var res = kind === 'positions' ? upsertPositions_(got.rows) : upsertEmployees_(got.rows);
    res.ok = true; res.cols = got.cols;
    return res;
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  } finally {
    if (tmpId) { try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e2) {} }
  }
}

/** ตรวจชนิดไฟล์จากชื่อชีต/หัวตาราง รองรับ Employee_Master_import.xlsx และ Position_Master_import.xlsx */
function detectImportKind_(fileId, fileName) {
  var ss = SpreadsheetApp.openById(fileId);
  var names = ss.getSheets().map(function (sh) { return normHead_(sh.getName()); });
  var fn = normHead_(fileName || '');
  if (names.indexOf('positions') >= 0 || names.indexOf('position') >= 0 || fn.indexOf('position') >= 0 || fn.indexOf('ตำแหน่ง') >= 0) return 'positions';
  if (names.indexOf('employees') >= 0 || names.indexOf('employee') >= 0 || fn.indexOf('employee') >= 0 || fn.indexOf('พนักงาน') >= 0) return 'employees';
  var sh = ss.getSheets()[0];
  var head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(normHead_);
  var scorePos = 0, scoreEmp = 0;
  Object.keys(IMPORT_ALIAS.positions).forEach(function (k) { if (IMPORT_ALIAS.positions[k].some(function (a) { return head.indexOf(normHead_(a)) >= 0; })) scorePos++; });
  Object.keys(IMPORT_ALIAS.employees).forEach(function (k) { if (IMPORT_ALIAS.employees[k].some(function (a) { return head.indexOf(normHead_(a)) >= 0; })) scoreEmp++; });
  return scoreEmp > scorePos ? 'employees' : 'positions';
}

/** นำเข้า Master 2 ไฟล์ในครั้งเดียว โดยบังคับประมวลผลตำแหน่งก่อนพนักงาน */
function apiImportMasterPair(pin, payloads) {
  guard_(pin);
  if (!payloads || !payloads.length) return { ok: false, error: 'ไม่พบไฟล์ที่เลือก' };
  var tmp = [], grouped = { positions: [], employees: [] };
  try {
    payloads.forEach(function (p, i) {
      var blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mimeType || 'application/octet-stream', p.name || ('import_' + i));
      var id = toSheet_(blob, 'IMPORT_TMP_MASTER_' + i); tmp.push(id);
      var kind = detectImportKind_(id, p.name);
      var got = readImport_(id, kind);
      if (got.headErr) throw new Error((p.name || 'ไฟล์') + ': ' + got.headErr);
      if (!got.rows.length) throw new Error((p.name || 'ไฟล์') + ': ไม่พบข้อมูล');
      grouped[kind] = grouped[kind].concat(got.rows);
    });
    if (!grouped.positions.length || !grouped.employees.length)
      return { ok: false, error: 'กรุณาเลือกทั้ง Position_Master_import.xlsx และ Employee_Master_import.xlsx พร้อมกัน' };
    var rp = upsertPositions_(grouped.positions);
    var re = upsertEmployees_(grouped.employees);
    return { ok: true, positions: rp, employees: re,
      added: rp.added + re.added, updated: rp.updated + re.updated, skipped: rp.skipped + re.skipped,
      unknownPos: re.unknownPos || [] };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  } finally {
    tmp.forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e2) {} });
  }
}

// รหัสซ้ำ = อัปเดตทับ, รหัสใหม่ = เพิ่ม, ไม่มีรหัส = ออกรหัสให้อัตโนมัติ
function upsertPositions_(rows) {
  var sh = getSs_().getSheetByName('Positions');
  var data = sh.getDataRange().getValues();
  var idx = {}; for (var i = 1; i < data.length; i++) idx[String(data[i][0]).trim()] = i + 1;
  var add = 0, upd = 0, skip = 0, seq = data.length;
  rows.forEach(function (r) {
    if (!r.position_name && !r.position_id) { skip++; return; }
    var id = r.position_id || ('POS' + ('000' + (++seq)).slice(-3));
    var rate = num_(r.hourly_rate);
    var ot = r.ot_rate ? num_(r.ot_rate) : round2_(rate * (parseFloat(readSettings_().ot_multiplier) || 1.5));
    var row = [id, r.position_name || id, r.department || '', rate, ot, r.status || 'ใช้งาน', r.note || ''];
    if (idx[id]) { sh.getRange(idx[id], 1, 1, row.length).setValues([row]); upd++; }
    else { sh.appendRow(row); idx[id] = sh.getLastRow(); add++; }
  });
  if (add) sortMasterSheet_(sh);   // แถวใหม่ถูก append ต่อท้าย → เรียงใหม่ให้รหัสไล่ลำดับเสมอ
  return { added: add, updated: upd, skipped: skip };
}

/** เรียงชีต Master ตามรหัส (คอลัมน์ A) — รหัสเติมศูนย์ 3 หลักอยู่แล้ว เรียงแบบข้อความจึงได้ลำดับถูก */
function sortMasterSheet_(sh) {
  var last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, sh.getLastColumn()).sort({ column: 1, ascending: true });
}

/** เรียงชีตพนักงาน/ตำแหน่งใหม่ทั้งสองชีต โดยไม่แตะค่าในแถว (ไม่ normalize ไม่คำนวณอะไรใหม่) */
function apiSortMaster(pin) {
  guard_(pin);
  var ss = getSs_(), out = {};
  ['Employees', 'Positions'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) return;
    sortMasterSheet_(sh);
    out[n] = Math.max(0, sh.getLastRow() - 1);
  });
  return out;
}

/* เบอร์โทร/เลขบัตร/เลขบัญชี ต้องเก็บเป็น "ข้อความ" — ไม่งั้น Sheet แปลงเป็นตัวเลขแล้วเลข 0 ข้างหน้าหาย
   (0812345678 → 812345678) · ใส่ ' นำหน้าตอนเขียน = บังคับเป็นข้อความ (ตอนอ่านกลับไม่มี ' ติดมา) */
function txt_(v) { v = (v === null || v === undefined) ? '' : String(v).trim(); return v ? "'" + v.replace(/^'/, '') : ''; }
function phoneTxt_(v) {
  var s = (v === null || v === undefined) ? '' : String(v).trim().replace(/^'/, '');
  if (/^[1-9]\d{8}$/.test(s)) s = '0' + s;   // เบอร์มือถือ 10 หลักที่ 0 หน้าหายไปแล้ว → เติมคืน
  return txt_(s);
}

function upsertEmployees_(rows) {
  var sh = getSs_().getSheetByName('Employees');
  var data = sh.getDataRange().getValues();
  var idx = {}; for (var i = 1; i < data.length; i++) idx[String(data[i][0]).trim()] = i + 1;
  // ให้ใส่ "ชื่อตำแหน่ง" ในไฟล์ได้ ไม่ต้องรู้รหัส POSxx → แปลงชื่อกลับเป็นรหัสให้
  var pos = readSheet_('Positions'), byName = {}, byId = {};
  pos.forEach(function (p) { byName[normHead_(p.position_name)] = p; byId[String(p.position_id).trim()] = p; });
  var add = 0, upd = 0, skip = 0, seq = data.length, unknownPos = [];
  rows.forEach(function (r) {
    if (!r.employee_name && !r.employee_id) { skip++; return; }
    var id = r.employee_id || ('EMP' + ('000' + (++seq)).slice(-3));
    var p = byId[String(r.position || '').trim()] || byName[normHead_(r.position)] || null;
    if (r.position && !p && unknownPos.indexOf(r.position) < 0) unknownPos.push(r.position);
    var rate = r.hourly_rate ? num_(r.hourly_rate) : (p ? num_(p.hourly_rate) : 0);
    var ot = r.ot_rate ? num_(r.ot_rate) : (p ? num_(p.ot_rate) : round2_(rate * 1.5));
    var old = idx[id] ? data[idx[id] - 1] : null;
    // ธนาคาร: รับได้ทั้งช่องรวม ("ธนาคาร | เลขบัญชี") และแยก 2 ช่อง — ช่องแยกมีสิทธิ์ก่อน
    var bk = splitBank_(r.bank);
    var bankName = r.bank_name || bk.name, bankAcc = r.bank_account || bk.acc;
    var row = [id, r.employee_name || (old ? old[1] : id), 'Part-time', rate, ot,
      r.status || (old ? old[5] : '') || 'ทำงาน', p ? p.position_id : (old ? old[6] : ''),
      normDate_(r.start_date) || r.start_date || (old ? old[7] : '') || '', phoneTxt_(r.phone || (old ? old[8] : '')),
      txt_(r.national_id || (old ? old[9] : '')), bankName || (old ? old[10] : '') || '', txt_(bankAcc || (old ? old[11] : ''))];
    if (idx[id]) { sh.getRange(idx[id], 1, 1, row.length).setValues([row]); upd++; }
    else { sh.appendRow(row); idx[id] = sh.getLastRow(); add++; }
  });
  if (add) sortMasterSheet_(sh);   // แถวใหม่ถูก append ต่อท้าย → เรียงใหม่ให้รหัสไล่ลำดับเสมอ
  return { added: add, updated: upd, skipped: skip, unknownPos: unknownPos };
}

function apiSaveEmployee(pin, emp) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Employees'); var data = sh.getDataRange().getValues();
  // ตัดช่องว่างหัว-ท้ายรหัสเสมอ — เคยพิมพ์ "EMP031␣" แล้วบันทึกซ้ำกลายเป็นพนักงานซ้ำ 2 แถว
  var eid = String(emp.employee_id == null ? '' : emp.employee_id).trim();
  if (!eid) return { ok: false, error: 'กรุณาใส่รหัสพนักงาน' };
  var row = [eid, String(emp.employee_name == null ? '' : emp.employee_name).trim(), emp.type || 'Part-time', num_(emp.hourly_rate), num_(emp.ot_rate),
    emp.status || 'ทำงาน', emp.position || '', normDate_(emp.start_date) || emp.start_date || '', phoneTxt_(emp.phone),
    txt_(emp.national_id), emp.bank_name || '', txt_(emp.bank_account)];
  for (var i = 1; i < data.length; i++) if (String(data[i][0]).trim() === eid) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: true }; }
  sh.appendRow(row); return { ok: true, updated: false };
}
function apiDeleteEmployee(pin, id) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Employees'); var data = sh.getDataRange().getValues();
  id = String(id == null ? '' : id).trim();
  for (var i = data.length - 1; i >= 1; i--) if (String(data[i][0]).trim() === id) { sh.deleteRow(i + 1); return { ok: true }; }
  return { ok: false };
}

// ลบพนักงานหลายรายการในครั้งเดียว
function apiDeleteEmployees(pin, ids) {
  guard_(pin);
  if (!Array.isArray(ids) || !ids.length) return { ok: false, deleted: 0, error: 'ไม่ได้เลือกรายการ' };
  var wanted = {}; ids.forEach(function(id){ id=String(id||'').trim(); if(id) wanted[id]=true; });
  var sh=getSs_().getSheetByName('Employees'); var data=sh.getDataRange().getValues();
  if(data.length<=1) return {ok:true,deleted:0};
  var kept=[data[0]],deleted=0;
  for(var i=1;i<data.length;i++){var id=String(data[i][0]||'').trim();if(wanted[id])deleted++;else kept.push(data[i]);}
  // เขียนทั้งชีตกลับ → เบอร์/เลขบัตร/เลขบัญชีต้องคงเป็นข้อความ ไม่งั้นเลข 0 ข้างหน้าหายทั้งชีต
  for(var k=1;k<kept.length;k++){kept[k][8]=phoneTxt_(kept[k][8]);kept[k][9]=txt_(kept[k][9]);kept[k][11]=txt_(kept[k][11]);}
  sh.clearContents(); sh.getRange(1,1,kept.length,kept[0].length).setValues(kept); sh.setFrozenRows(1);
  return {ok:true,deleted:deleted,requested:Object.keys(wanted).length};
}

/* ================= ตำแหน่งงาน ================= */
function apiPositions(pin) {
  guard_(pin);
  var pos = readSheet_('Positions'); var emps = readSheet_('Employees');
  var cnt = {}; emps.forEach(function (e) { if (e.position) cnt[e.position] = (cnt[e.position] || 0) + 1; });
  return pos.map(function (p) { p.emp_count = cnt[p.position_id] || 0; return p; });
}
function apiSavePosition(pin, p) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Positions'); var data = sh.getDataRange().getValues();
  // ต้องเรียงตาม SHEETS.Positions เป๊ะ: รหัส · ตำแหน่ง · แผนก · ค่าแรง · OT · สถานะ · หมายเหตุ
  var row = [p.position_id, p.position_name, p.department || '', num_(p.hourly_rate), num_(p.ot_rate), p.status || 'ใช้งาน', p.note || ''];
  for (var i = 1; i < data.length; i++) if (data[i][0] === p.position_id) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: true }; }
  sh.appendRow(row); return { ok: true };
}
function apiDeletePosition(pin, id) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Positions'); var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) if (data[i][0] === id) { sh.deleteRow(i + 1); return { ok: true }; }
  return { ok: false };
}

// ลบตำแหน่งงานหลายรายการในครั้งเดียว โดยเขียนข้อมูลที่เหลือกลับเพียงรอบเดียว
function apiDeletePositions(pin, ids) {
  guard_(pin);
  if (!Array.isArray(ids) || !ids.length) return { ok: false, deleted: 0, error: 'ไม่ได้เลือกรายการ' };
  var wanted = {};
  ids.forEach(function (id) { id = String(id || '').trim(); if (id) wanted[id] = true; });
  var sh = getSs_().getSheetByName('Positions');
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, deleted: 0 };
  var kept = [data[0]], deleted = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || '').trim();
    if (wanted[id]) deleted++; else kept.push(data[i]);
  }
  sh.clearContents();
  sh.getRange(1, 1, kept.length, kept[0].length).setValues(kept);
  sh.setFrozenRows(1);
  return { ok: true, deleted: deleted, requested: Object.keys(wanted).length };
}

/* ================= ตารางกะการทำงาน ================= */
function apiSchedule(pin) { guard_(pin); return readSheet_('Schedule'); }
function apiSaveSchedule(pin, s) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Schedule'); var data = sh.getDataRange().getValues();
  var row = [s.employee_id, s.weekday, s.shift_start || '', s.shift_end || '', s.note || ''];
  for (var i = 1; i < data.length; i++)
    if (data[i][0] === s.employee_id && String(data[i][1]) === String(s.weekday)) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: true }; }
  sh.appendRow(row); return { ok: true };
}
function apiDeleteSchedule(pin, employeeId, weekday) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Schedule'); var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--)
    if (data[i][0] === employeeId && String(data[i][1]) === String(weekday)) { sh.deleteRow(i + 1); return { ok: true }; }
  return { ok: false };
}

/* ================= รายการปรับเงิน (list/delete) ================= */
function apiAdjustments(pin) {
  guard_(pin);
  var nameMap = {}; readSheet_('Employees').forEach(function (e) { nameMap[e.employee_id] = e.employee_name; });
  return readSheet_('Adjustments').map(function (a) {
    a.employee_name = nameMap[a.employee_id] || '';
    a.net_adj = round2_(num_(a.addition) + num_(a.bonus) + num_(a.diligence) - num_(a.deduction) - num_(a.wht));
    return a;
  }).reverse();
}
function apiDeleteAdjustment(pin, employeeId, from, to) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Adjustments'); var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--)
    if (data[i][0] === employeeId && String(data[i][1]) === String(from || '') && String(data[i][2]) === String(to || '')) { sh.deleteRow(i + 1); return { ok: true }; }
  return { ok: false };
}

/* ================= ทะเบียนเงินได้และเงินเดือนพนักงาน (จัดกลุ่มตามแผนก) ================= */
var REG_KEYS = ['normal_hours', 'ot_hours', 'days', 'normal_pay', 'ot_pay', 'diligence', 'addition', 'bonus',
  'gross', 'wht', 'deduction', 'total_deduct', 'net'];
function sumRows_(rows) {
  var t = {}; REG_KEYS.forEach(function (k) { t[k] = 0; });
  rows.forEach(function (r) { REG_KEYS.forEach(function (k) { t[k] += num_(r[k]); }); });
  REG_KEYS.forEach(function (k) { t[k] = round2_(t[k]); });
  t.hours = round2_(t.normal_hours + t.ot_hours);
  return t;
}
function apiPayrollReport(pin, period) {
  guard_(pin); period = period || {};
  var rows = calcPayroll_(period);
  // จัดกลุ่มตามแผนก (มาจากตำแหน่งของพนักงาน) — คนที่ยังไม่ผูกตำแหน่งไปรวมกลุ่มท้ายสุด
  var byDep = {}, order = [];
  rows.forEach(function (r) {
    var d = r.department || 'ไม่ระบุแผนก';
    if (!byDep[d]) { byDep[d] = []; order.push(d); }
    byDep[d].push(r);
  });
  order.sort(function (a, b) {
    if (a === 'ไม่ระบุแผนก') return 1; if (b === 'ไม่ระบุแผนก') return -1;
    return a.localeCompare(b, 'th');
  });
  var groups = order.map(function (d) { return { department: d, rows: byDep[d], totals: sumRows_(byDep[d]) }; });
  var tot = sumRows_(rows);
  var byPosition = {};                                   // สรุปตามตำแหน่ง (ใช้ในการ์ดสรุป)
  rows.forEach(function (r) {
    var p = r.position || 'ไม่ระบุตำแหน่ง';
    if (!byPosition[p]) byPosition[p] = { position: p, count: 0, hours: 0, net: 0 };
    byPosition[p].count++; byPosition[p].hours += num_(r.normal_hours) + num_(r.ot_hours); byPosition[p].net += num_(r.net);
  });
  return JSON.stringify({
    period: period, rows: rows, groups: groups, totals: tot, headcount: rows.length,
    pay_date: readSettings_().pay_date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    byPosition: Object.keys(byPosition).map(function (k) {
      var o = byPosition[k]; o.hours = round2_(o.hours); o.net = round2_(o.net); return o;
    }).sort(function (a, b) { return b.net - a.net; }),
  });
}

/** รับ PDF ที่สร้างจากฝั่งเบราว์เซอร์ (html2canvas+jsPDF — สีตรงหน้าเว็บ 100%) มาเก็บลง Drive
    folder เลือกได้เฉพาะที่อนุญาต: 06_Reports (ทะเบียน/รายงาน) หรือ 05_Payslips (สลิป) */
var REPORT_PDF_MAX = 25 * 1024 * 1024;   // สลิปรวมทุกคน ~หลาย MB — เผื่อไว้ แต่กันคนยิงไฟล์ยักษ์มาถมไดรฟ์เจ้าของ
function apiSaveReportPdf(pin, name, base64, folder, slipCount) {
  guard_(pin);
  var dest = folder === '05_Payslips' ? '05_Payslips' : '06_Reports';
  var safe = String(name || 'Report.pdf').replace(/[^\w.\-฀-๿]/g, '_');
  /* hardening (โคลน): เว็บเปิด ANYONE_ANONYMOUS + รันสิทธิ์เจ้าของ → ใครก็ยิง api นี้ได้
     - ต้องเป็น PDF จริง (ขึ้นต้น %PDF) ไม่งั้นใครก็ฝากไฟล์อะไรก็ได้ไว้ในไดรฟ์เจ้าของ
       และไฟล์ใน 05_Payslips ถูกบอทหยิบไปส่งให้พนักงานเป็น "สลิปจริง" (findIssuedSlip_)
     - บังคับนามสกุล .pdf + จำกัดขนาด */
  if (!/\.pdf$/i.test(safe)) safe += '.pdf';
  var bytes = Utilities.base64Decode(String(base64 || ''));
  if (!bytes.length || bytes.length > REPORT_PDF_MAX) return { ok: false, error: 'ไฟล์ว่าง หรือใหญ่เกิน ' + (REPORT_PDF_MAX / 1048576) + ' MB' };
  if (!(bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70)) return { ok: false, error: 'ไฟล์ไม่ใช่ PDF' };   // '%PDF'
  var blob = Utilities.newBlob(bytes, 'application/pdf', safe);
  subFolder_(dest).createFile(blob);
  // สลิปสร้างฝั่ง browser แล้ว ตัวนับ "สลิปที่ออกแล้ว" เลยต้องบวกที่นี่ (ไฟล์รวมหลายคน = บวกตามจำนวนคน)
  if (dest === '05_Payslips') {
    var n = parseInt(props_().getProperty('SLIP_SEQ') || '0', 10) + (parseInt(slipCount, 10) || 1);
    props_().setProperty('SLIP_SEQ', String(n));
  }
  return { ok: true, name: safe };
}

/** ล้างตัวนับ "สลิปที่ออกแล้ว" บนแดชบอร์ด (เริ่มนับใหม่จาก 0) */
function apiResetSlipCount(pin) {
  guard_(pin);
  props_().setProperty('SLIP_SEQ', '0');
  return { ok: true };
}

/* ================= ประวัติสลิปที่ออกไปแล้ว ================= */

/** รายการสลิปที่เคยออก — อ่านจาก "ไฟล์จริง" ใน Drive → 05_Payslips ไม่ใช่ตัวนับ SLIP_SEQ
    (ตัวนับล้างได้และไม่ผูกกับไฟล์ ถ้าอ่านจากตัวนับจะได้ตัวเลขที่ไม่ตรงกับของที่มีอยู่จริง)

    ชื่อไฟล์ที่หน้าเว็บตั้งไว้ (slipOne / slipAllS ใน Index.html):
      รายคน   Payslip_<รหัสพนักงาน>_<วันเริ่มงวด>.pdf
      รวมเล่ม Payslips_ALL_<from>_<to>.pdf
    ไฟล์ชื่ออื่น (ของเก่า/อัปเข้าไปเอง) ก็คืนมาด้วย แค่แยกรหัส-งวดไม่ออก — ยังกดดู/ดาวน์โหลดได้ปกติ

    opt: {q, from, to, limit}  โดย from/to = ช่วง "วันที่ออกสลิป" ไม่ใช่งวดค่าจ้าง */
function apiListPayslips(pin, opt) {
  guard_(pin); opt = opt || {};
  var q = String(opt.q || '').trim().toLowerCase();
  var from = String(opt.from || ''), to = String(opt.to || '');
  var limit = parseInt(opt.limit, 10) || 400;

  var nameById = {};
  readSheet_('Employees').forEach(function (e) {
    var id = String(e.employee_id || '').trim();
    if (id) nameById[id] = e.employee_name || '';
  });

  /* performance: DriveApp วนไฟล์ทีละใบ (getName/getSize/getDateCreated = คนละเรียก)
     หลักร้อยใบยังเร็วพอ ถ้าโตถึงหลักพันค่อยย้ายไป Drive API แบบ query+fields
     CAP กันวนไม่จบจนชน execution limit แล้วหน้าเว็บค้าง */
  var CAP = 2000;
  var it = subFolder_('05_Payslips').getFiles(), rows = [], scanned = 0;
  while (it.hasNext() && scanned < CAP) {
    scanned++;
    var f = it.next(), nm = f.getName(), created = f.getDateCreated();
    var day = Utilities.formatDate(created, TZ, 'yyyy-MM-dd');
    if (from && day < from) continue;
    if (to && day > to) continue;

    var kind = 'other', empId = '', pFrom = '', pTo = '';
    var mAll = nm.match(/^Payslips_ALL_(\d{4}-\d{2}-\d{2})?_(\d{4}-\d{2}-\d{2})?\.pdf$/i);
    var mOne = nm.match(/^Payslip_(.*?)_(\d{4}-\d{2}-\d{2})?\.pdf$/i);
    if (mAll) { kind = 'all'; pFrom = mAll[1] || ''; pTo = mAll[2] || ''; }
    else if (mOne) { kind = 'one'; empId = String(mOne[1] || '').trim(); pFrom = mOne[2] || ''; }

    var empName = empId ? (nameById[empId] || '') : '';
    if (q && (nm + ' ' + empId + ' ' + empName).toLowerCase().indexOf(q) < 0) continue;

    rows.push({
      id: f.getId(), name: nm, kind: kind,
      employee_id: empId, employee_name: empName,
      period_from: pFrom, period_to: pTo,
      created: Utilities.formatDate(created, TZ, 'yyyy-MM-dd HH:mm'),
      ts: created.getTime(), size: f.getSize(),
    });
  }
  var truncated = it.hasNext();

  // เรียงใหม่→เก่า · ts เท่ากันได้ถ้าออกสลิปรัวๆ ในวินาทีเดียว → ต่อ tiebreaker ด้วยชื่อไฟล์
  // ให้ลำดับคงที่ทุกครั้งที่โหลด (ไม่งั้นแถวสลับที่เองเวลากดค้นหาซ้ำ)
  rows.sort(function (a, b) { return (b.ts - a.ts) || (a.name < b.name ? -1 : a.name > b.name ? 1 : 0); });

  /* ออกสลิปคนเดิมงวดเดิมซ้ำ = Drive สร้างไฟล์ "ชื่อซ้ำอีกใบ" ไม่ทับของเดิม
     เรียงใหม่→เก่าแล้ว ใบแรกของแต่ละชื่อคือใบล่าสุด ที่เหลือติดป้ายให้รู้ว่าเป็นของเก่า
     (ตั้งใจไม่ลบให้อัตโนมัติ — ไฟล์ที่ออกไปแล้วอาจส่งให้พนักงานไปแล้ว) */
  var seq = {};
  rows.forEach(function (r) { r.seq = seq[r.name] = (seq[r.name] || 0) + 1; });
  rows.forEach(function (r) { r.copies = seq[r.name]; });

  return JSON.stringify({
    ok: true, rows: rows.slice(0, limit), total: rows.length,
    scanned: scanned, truncated: truncated,
  });
}

var PAYSLIP_INLINE_MAX = 20 * 1024 * 1024;   // เกินนี้ส่งผ่าน google.script.run แล้วเบราว์เซอร์อืด/ล่ม

/** ส่งไฟล์สลิปกลับเป็น base64 ให้เว็บเปิดดู/ดาวน์โหลดเอง
    ตั้งใจไม่ส่ง "ลิงก์ Drive" กลับไป เพราะเว็บนี้เปิดแบบ ANYONE_ANONYMOUS
    คนดูไม่มีสิทธิ์ในไดรฟ์ กดลิงก์แล้วจะเจอหน้า Google Sign in / 403

    ความปลอดภัย: เว็บรันด้วยสิทธิ์เจ้าของ (USER_DEPLOYING) ถ้ารับ fileId มาดึงดื้อๆ
    ใครก็ตามที่เดา id ของไฟล์อื่นในไดรฟ์เจ้าของได้ ก็ดูดไฟล์นั้นออกไปได้ทันที
    → ต้องยืนยันก่อนเสมอว่าไฟล์อยู่ในโฟลเดอร์ 05_Payslips จริง */
function apiPayslipFile(pin, fileId) {
  guard_(pin);
  if (!fileId) return { ok: false, error: 'ไม่ได้ระบุไฟล์' };
  var file;
  try { file = DriveApp.getFileById(String(fileId)); }
  catch (e) { return { ok: false, error: 'ไม่พบไฟล์นี้แล้ว (อาจถูกลบหรือย้ายออกจากโฟลเดอร์)' }; }
  if (!inPayslipFolder_(file)) return { ok: false, error: 'ไฟล์นี้ไม่ได้อยู่ในโฟลเดอร์สลิป' };

  var size = file.getSize();
  if (size > PAYSLIP_INLINE_MAX) {
    return { ok: false, error: 'ไฟล์ใหญ่ ' + Math.round(size / 1048576) + ' MB เปิดผ่านหน้าเว็บไม่ไหว — เปิดจาก Drive โฟลเดอร์ 05_Payslips แทน' };
  }
  return { ok: true, name: file.getName(), size: size, base64: Utilities.base64Encode(file.getBlob().getBytes()) };
}

function inPayslipFolder_(file) {
  var target = subFolder_('05_Payslips').getId();
  var ps = file.getParents();
  while (ps.hasNext()) { if (ps.next().getId() === target) return true; }
  return false;
}

/* ================= ช่องทางบอท LINE — พนักงานขอสลิปของ "ตัวเอง" =================
   หลักการเดียวที่ทำให้ปลอดภัย: **ห้ามรับรหัสพนักงานจากภายนอกเด็ดขาด**
   ตัวตนต้องมาจาก line_user_id ที่บอทยืนยันมาแล้วเท่านั้น (LINE เซ็นลายเซ็นทุกข้อความ ปลอมไม่ได้)
   ถ้ารับ employee_id มาด้วย ใครก็ขอสลิปคนอื่นได้แค่เปลี่ยนตัวเลขที่ส่งมา */

/** ตั้ง/ดูรหัสลับระหว่างบอทกับสคริปต์นี้ — เปิดหน้า Apps Script แล้วกด Run ฟังก์ชันนี้ 1 ครั้ง
    ค่าที่ได้จะโผล่ใน Execution log → เอาไปตั้งเป็น GAS_SHARED_SECRET ที่ฝั่งบอท
    (เว็บแอปนี้เปิดแบบ ANYONE_ANONYMOUS ถ้าไม่มีรหัสลับ ใครยิง doPost ก็ขอสลิปคนอื่นได้) */
function setupBotSecret() {
  var p = props_();
  var v = p.getProperty('BOT_SECRET');
  var made = false;
  if (!v) {
    v = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    p.setProperty('BOT_SECRET', v);
    made = true;
  }
  console.log('===== รหัสลับสำหรับบอท (' + (made ? 'สร้างใหม่' : 'ของเดิมที่ตั้งไว้แล้ว') + ') =====');
  console.log(v);
  console.log('เอาค่าข้างบนไปตั้งเป็น GAS_SHARED_SECRET ที่ฝั่งบอท (Railway → Variables)');
  // ไม่ return ค่ากลับ (15 ก.ย. 69): ฟังก์ชันที่ไม่มี _ ท้ายชื่อ ถูกเรียกผ่าน google.script.run จากหน้าเว็บได้
  // และเว็บนี้เปิดแบบ ANYONE_ANONYMOUS — ถ้าคืนค่า ใครมีลิงก์เว็บก็ดึงรหัสลับไปขอสลิปคนอื่นได้ · ดูค่าได้จาก Execution log เท่านั้น
}

/** ลบข้อมูลทดสอบของบอท — ใช้ตอนทดสอบส่งใบลงเวลาเข้าเว็บจริงแล้วเก็บกวาด
    ลบได้เฉพาะแถวที่ line_user_id ขึ้นต้นด้วย TEST_BOT_ (LINE id จริงขึ้นต้นด้วย U เสมอ จึงแตะข้อมูลจริงไม่ได้)
    ไฟล์: ทิ้งลงถังขยะเฉพาะไฟล์ใน 01_Inbox ที่ชื่อขึ้นต้น TEST_BOT_ / WEB_ และสร้างไม่เกิน 3 ชั่วโมง */
function apiPurgeBotTest(pin, testUserId, fileIds) {
  guard_(pin);
  var uid = String(testUserId || '');
  if (!/^TEST_BOT_[A-Za-z0-9_]{4,40}$/.test(uid)) return { ok: false, error: 'ลบได้เฉพาะข้อมูลทดสอบ (TEST_BOT_…)' };
  var ss = getSs_(), out = { ok: true, attendance: 0, log: 0, files: 0 };
  [['Attendance', 'attendance'], ['Import_Log', 'log']].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return;
    var vals = sh.getDataRange().getValues(), col = vals[0].indexOf('line_user_id');
    if (col < 0) return;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][col]) === uid) { sh.deleteRow(r + 1); out[x[1]]++; }
    }
  });
  var inbox = subFolder_('01_Inbox'), inboxId = inbox.getId(), since = Date.now() - 3 * 3600 * 1000;
  var okFile = function (f) {
    if (f.getDateCreated().getTime() < since || !/^(TEST_BOT_|WEB_)/.test(f.getName())) return false;
    var ps = f.getParents();
    while (ps.hasNext()) if (ps.next().getId() === inboxId) return true;
    return false;
  };
  var trash = function (f) { if (okFile(f) && !f.isTrashed()) { f.setTrashed(true); out.files++; } };
  var it = inbox.searchFiles("title contains 'TEST_BOT_'");
  while (it.hasNext()) trash(it.next());
  (fileIds || []).forEach(function (id) {
    try {
      var f = DriveApp.getFileById(String(id));
      trash(f);
      // หน้าอัปรูปเก็บผลอ่านคู่กันเป็น <ชื่อรูป>_ocr.json (ไม่คืน id) → ตามลบตัวคู่ด้วย
      var sib = inbox.getFilesByName(f.getName().replace(/\.[a-z]+$/i, '') + '_ocr.json');
      while (sib.hasNext()) trash(sib.next());
    } catch (e) {}
  });
  return out;
}

/** เมนู "อัปรูปใบลงเวลา" พร้อมใช้ไหม (มีคีย์ AI ใน Script Properties) — ไม่คืนตัวคีย์ */
/*  test=true → ลองเรียก Gemini จริง 1 ครั้ง (ปุ่ม "ทดสอบคีย์ AI" หน้าตั้งค่า) */
function apiTsOcrReady(pin, test) {
  guard_(pin);
  var ready = !!String(props_().getProperty('GEMINI_API_KEY') || '').trim();
  if (!test) return { ready: ready };
  if (typeof tsocrPing_ !== 'function') return { ready: ready, ok: false, error: 'ยังไม่ได้ติดตั้งไฟล์ TimesheetOcr.gs' };
  var r = tsocrPing_();
  return { ready: ready, ok: !!r.ok, model: r.model || '', error: r.error || '' };
}

/** เทียบรหัสลับแบบไม่แพ้เวลา (Apps Script ไม่มี timingSafeEqual ให้ใช้)
    ยังไม่ตั้งรหัส = ปฏิเสธไว้ก่อน — ปิดตายปลอดภัยกว่าเปิดโล่งรอคนมาตั้ง */
function botAuthOk_(given) {
  var want = props_().getProperty('BOT_SECRET') || '';
  if (!want) return false;
  var got = String(given || '');
  if (got.length !== want.length) return false;
  var diff = 0;
  for (var i = 0; i < want.length; i++) diff |= want.charCodeAt(i) ^ got.charCodeAt(i);
  return diff === 0;
}

/** LINE id → พนักงานคนนั้น โดยดูจากใบลงเวลาที่ "เจ้าตัวส่งเข้ามาเอง" ผ่านบอท
    (ทะเบียนพนักงานไม่มีช่องเก็บ LINE id การจับคู่จึงมาทางนี้ทางเดียว)
    ไล่จากแถวใหม่ไปเก่า เผื่อเคยกรอกรหัสผิดแล้วแก้ทีหลัง — เอาการผูกล่าสุดเป็นหลัก */
function payslipEmployeeOf_(lineUserId) {
  var uid = String(lineUserId || '').trim();
  if (!uid) return null;
  var rows = readSheet_('Attendance');
  var found = null;
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i].line_user_id || '').trim() !== uid) continue;
    var id = String(rows[i].employee_id || '').trim();
    if (id) { found = { employee_id: id, employee_name: rows[i].employee_name || rows[i].display_name || '' }; break; }
  }
  if (!found) return null;
  var emp = readSheet_('Employees').filter(function (e) { return String(e.employee_id || '').trim() === found.employee_id; })[0];
  if (emp && emp.employee_name) found.employee_name = emp.employee_name;   // ชื่อในทะเบียนถือเป็นชื่อจริง
  return found;
}

/** 'yyyy-MM' → ช่วงวันที่ต้น-ปลายเดือน · ไม่ส่งมา = เดือนปัจจุบัน */
function monthPeriod_(ym) {
  var m = String(ym || '').match(/^(\d{4})-(\d{2})$/);
  var now = new Date();
  var y = m ? parseInt(m[1], 10) : parseInt(Utilities.formatDate(now, TZ, 'yyyy'), 10);
  var mo = m ? parseInt(m[2], 10) : parseInt(Utilities.formatDate(now, TZ, 'MM'), 10);
  var last = new Date(y, mo, 0).getDate();          // วันที่ 0 ของเดือนถัดไป = วันสุดท้ายของเดือนนี้
  var p2 = function (n) { return (n < 10 ? '0' : '') + n; };
  return { from: y + '-' + p2(mo) + '-01', to: y + '-' + p2(mo) + '-' + p2(last) };
}

/** ใบที่แอดมินกดออกไว้แล้วของงวดนี้ (ออกซ้ำหลายใบ → เอาใบล่าสุด) หรือ null */
function findIssuedSlip_(empId, from) {
  var it = subFolder_('05_Payslips').getFilesByName('Payslip_' + empId + '_' + (from || '') + '.pdf');
  var best = null;
  while (it.hasNext()) {
    var f = it.next();
    if (!best || f.getDateCreated() > best.getDateCreated()) best = f;
  }
  return best;
}

/** เลขที่สลิปสำหรับใบที่พนักงานขอเอง — ตั้งใจไม่เรียก slipNo_() เพราะตัวนั้นบวกตัวนับ
    "สลิปที่ออกแล้ว" บนแดชบอร์ด ใบที่พนักงานกดขอเองไม่ควรไปปนกับยอดที่แอดมินออก */
function botSlipNo_(s, period, empId) {
  return (s.slip_prefix || 'SLIP') + '-' + String(period.from || '').replace(/-/g, '').slice(2, 6) + '-' + empId;
}

/** แถบเตือนบนใบชั่วคราว — ใช้เส้นขอบ+สีตัวอักษรเท่านั้น
    เพราะตัวแปลง HTML→PDF ของ Drive ไม่ลงสีพื้น (พิสูจน์ด้วยภาพจริงแล้ว) แถบสีพื้นจะหายเกลี้ยง */
function draftBanner_(asOf) {
  return '<table style="width:100%;border-collapse:collapse;margin-bottom:9px"><tr>' +
    '<td style="border:2px solid #C0392B;padding:7px 12px;color:#C0392B;font-size:12.5px;font-weight:bold">' +
    'ฉบับชั่วคราว — งวดนี้ยังไม่ปิด ข้อมูลคำนวณถึงวันที่ ' + esc_(thaiDate_(asOf)) +
    ' เท่านั้น ยอดจริงจะเปลี่ยนเมื่อปิดงวด</td></tr></table>';
}

/** งวดที่พนักงานคนนี้ขอสลิปได้ + ยอดคร่าว ๆ (ยังไม่สร้าง PDF — ให้การ์ดในไลน์ขึ้นเร็ว) */
function botPayslipList_(lineUserId, opt) {
  var who = payslipEmployeeOf_(lineUserId);
  if (!who) return { ok: false, reason: 'unknown_employee' };
  opt = opt || {};
  var months = [];
  if (opt.month) months.push(opt.month);
  else {
    var now = new Date();
    months.push(Utilities.formatDate(now, TZ, 'yyyy-MM'));
    months.push(Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() - 1, 1), TZ, 'yyyy-MM'));
  }
  var today = today_(), out = [];
  months.forEach(function (ym) {
    var period = monthPeriod_(ym);
    var row = calcPayroll_(period).filter(function (r) { return String(r.employee_id || '').trim() === who.employee_id; })[0];
    if (!row) return;                                  // เดือนนั้นไม่มีใบลงเวลา = ไม่มีอะไรให้ออกสลิป
    out.push({
      month: ym, from: period.from, to: period.to, label: thaiPeriod_(period.from, period.to),
      days: row.days, hours: round2_(num_(row.normal_hours) + num_(row.ot_hours)),
      net: round2_(num_(row.net)),
      closed: period.to < today,
      issued: !!findIssuedSlip_(who.employee_id, period.from),
    });
  });
  return { ok: true, employee: who, periods: out };
}

/** ไฟล์สลิปของงวดที่ขอ — งวดปิดแล้วและแอดมินออกใบไว้ = ส่งใบนั้น (ดีไซน์เต็ม สีครบ)
    ไม่งั้นสร้างสดจากข้อมูลลงเวลา ณ ตอนนั้น
    ใบที่สร้างสดตั้งใจ "ไม่เซฟลง Drive" — ไม่งั้นโฟลเดอร์บวมและไปปนกับใบจริงที่แอดมินออก */
function botPayslipFile_(lineUserId, from, to) {
  var who = payslipEmployeeOf_(lineUserId);
  if (!who) return { ok: false, reason: 'unknown_employee' };
  var period = (from && to) ? { from: String(from), to: String(to) } : monthPeriod_('');
  var closed = period.to < today_();

  if (closed) {
    var issued = findIssuedSlip_(who.employee_id, period.from);
    if (issued) {
      return { ok: true, name: issued.getName(), size: issued.getSize(), draft: false, issued: true,
        employee_name: who.employee_name, label: thaiPeriod_(period.from, period.to),
        base64: Utilities.base64Encode(issued.getBlob().getBytes()) };
    }
  }
  var row = calcPayroll_(period).filter(function (r) { return String(r.employee_id || '').trim() === who.employee_id; })[0];
  if (!row) return { ok: false, reason: 'no_data' };

  var s = readSettings_();
  var body = (closed ? '' : draftBanner_(today_())) + payslipBody_(s, row, period, botSlipNo_(s, period, who.employee_id));
  var pdf = Utilities.newBlob(wrapHtml_(body), 'text/html', 'slip.html').getAs('application/pdf');
  var bytes = pdf.getBytes();
  return { ok: true, draft: !closed, issued: false,
    name: 'Payslip_' + who.employee_id + '_' + period.from + (closed ? '' : '_draft') + '.pdf',
    size: bytes.length, employee_name: who.employee_name, label: thaiPeriod_(period.from, period.to),
    base64: Utilities.base64Encode(bytes) };
}

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** ทะเบียนเงินได้เป็น PDF (A4 แนวนอน) — เก็บลง 06_Reports ให้ด้วย
    หมายเหตุ: ปุ่ม PDF บนเว็บเลิกใช้ตัวนี้แล้ว (converter ของ Drive ไม่ render พื้นสีตาราง)
    คงไว้เผื่อเรียกจาก flow อื่น */
function apiPayrollRegisterPdf(pin, period) {
  guard_(pin); period = period || {};
  var data = JSON.parse(apiPayrollReport(pin, period));
  if (!data.rows.length) return { ok: false, error: 'ไม่มีข้อมูลในงวดนี้' };
  var html = '<html><head><meta charset="utf-8"><style>' + DOC_CSS +
    '@page{size:A4 landscape;margin:8mm}</style></head><body>' + registerBody_(readSettings_(), data) + '</body></html>';
  var pdf = Utilities.newBlob(html, 'text/html', 'register.html').getAs('application/pdf');
  var name = 'PayrollRegister_' + (period.from || '') + '_' + (period.to || '') + '.pdf'; pdf.setName(name);
  try { subFolder_('06_Reports').createFile(pdf); } catch (e) {}
  return { ok: true, name: name, base64: Utilities.base64Encode(pdf.getBytes()), count: data.rows.length };
}

/* cell ทะเบียนเงินได้ — หมายเหตุ (พิสูจน์ 21 ก.ค. 2569): converter HTML→PDF ของ Drive
   "ไม่ render พื้นสีของ cell เลย" ไม่ว่าเขียนแบบไหน (ลองครบ bgcolor/inline/th/tr/div ฯลฯ)
   สีที่เห็นใน PDF ฝั่ง server มีแค่เส้นขอบกับตัวอักษร → PDF สีจริงต้องสร้างฝั่ง browser (regPdfBlob ใน Index.html) */
function regCell_(v, cls, bg) {
  var al = cls === 'right' ? 'right' : (cls === 'center' ? 'center' : 'left');
  if (bg) {
    return '<td bgcolor="' + bg + '" align="' + al + '" style="background:' + bg + ';border:1px solid ' + bg + ';padding:4px 6px' +
      ';text-align:' + al + ';color:#1B2430;font-weight:bold">' + v + '</td>';
  }
  return '<td align="' + al + '" style="border:1px solid #C9D6E8;padding:4px 6px;text-align:' + al + '">' + v + '</td>';
}
function registerBody_(s, data) {
  var t = data.totals;
  var head = '<table><tr>' +
    '<td style="width:150px">' + (getLogo_() ? '<img src="' + getLogo_() + '" style="height:44px">' : '') +
      '<div class="muted" style="font-size:9px;margin-top:2px">' + esc_(s.company_name_en || '') + '</div></td>' +
    '<td class="center"><div style="font-size:17px;font-weight:bold;color:#123A6B">' + esc_(s.company_name) + '</div>' +
      '<div style="font-size:13px;font-weight:bold;color:#123A6B">รายงานทะเบียนเงินได้และเงินเดือนพนักงาน (Part-time)</div></td>' +
    '<td style="width:230px" class="right"><table style="border:1px solid #C9D6E8;border-radius:8px"><tr><td bgcolor="#EEF3FC" align="left" style="background:#EEF3FC;border:1px solid #EEF3FC;padding:7px 10px;font-size:10.5px;color:#123A6B;text-align:left">' +
      'งวดค่าจ้าง : ' + esc_(thaiPeriod_(data.period.from, data.period.to)) +
      '<br>วันที่จ่าย&nbsp;&nbsp; : ' + esc_(thaiDate_(data.pay_date)) + '</td></tr></table></td></tr></table>';

  /* หัวตาราง 2 ชั้นแบบหน้าเว็บ: navy + กลุ่ม รายได้(น้ำเงิน)/รายการหัก(แดง)/รายได้สุทธิ(เขียว)
     converter ไม่อ่าน CSS class → สีต้องเป็น bgcolor attribute + inline background เท่านั้น */
  var th = function (label, bg, span, w) {
    return '<th' + (span || '') + ' bgcolor="' + bg + '" style="background:' + bg + ';color:#fff;border:1px solid ' + bg +
      ';padding:5px 6px;font-size:10px;font-weight:bold;vertical-align:middle;text-align:center' + (w ? ';width:' + w : '') + '">' + label + '</th>';
  };
  var NAVY = '#123A6B', BLUE = '#1F4E9C', RED = '#C0392B', GREEN = '#1E7A45';
  var thead = '<tr>' +
    th('ลำดับ', NAVY, ' rowspan="2"', '34px') + th('รหัสพนักงาน', NAVY, ' rowspan="2"', '68px') +
    th('ชื่อ-สกุล', NAVY, ' rowspan="2"') + th('ตำแหน่ง', NAVY, ' rowspan="2"', '96px') +
    th('จำนวนชั่วโมง<br>(ชม.)', NAVY, ' rowspan="2"', '58px') + th('อัตราค่าจ้าง<br>(บาท/ชม.)', NAVY, ' rowspan="2"', '58px') +
    th('รายได้', BLUE, ' colspan="4"') + th('รายการหัก', RED, ' colspan="2"') +
    th('รายได้สุทธิ', GREEN, ' rowspan="2"', '72px') + '</tr>' +
    '<tr>' + th('ค่าจ้างรายชั่วโมง', BLUE, '', '76px') + th('ค่าล่วงเวลา', BLUE, '', '64px') +
    th('รายได้อื่น', BLUE, '', '60px') + th('รายได้รวม', BLUE, '', '68px') +
    th('ภาษี ณ ที่จ่าย', RED, '', '64px') + th('อื่นๆ', RED, '', '52px') + '</tr>';

  var body = '', n = 0;
  data.groups.forEach(function (g) {
    body += '<tr><td colspan="13" bgcolor="#EAE5FE" align="left" style="border:1px solid #EAE5FE;background:#EAE5FE;padding:5px 8px;font-weight:bold;color:#1B2430;text-align:left">' +
      'แผนก : ' + esc_(g.department) + '</td></tr>';
    g.rows.forEach(function (r) {
      n++;
      body += '<tr>' + regCell_(n, 'center') + regCell_(esc_(r.employee_id)) + regCell_(esc_(r.employee_name)) +
        regCell_(esc_(r.position || '-')) +
        regCell_(fmtBaht_(num_(r.normal_hours) + num_(r.ot_hours)), 'right') + regCell_(fmtBaht_(r.hourly_rate), 'right') +
        regCell_(fmtBaht_(r.normal_pay), 'right') + regCell_(fmtBaht_(r.ot_pay), 'right') +
        regCell_(fmtBaht_(num_(r.diligence) + num_(r.addition) + num_(r.bonus)), 'right') +
        regCell_(fmtBaht_(r.gross), 'right') + regCell_(fmtBaht_(r.wht), 'right') + regCell_(fmtBaht_(r.deduction), 'right') +
        regCell_('<b>' + fmtBaht_(r.net) + '</b>', 'right') + '</tr>';
    });
    var gt = g.totals;
    var SUBBG = '#F2EFFE';
    body += '<tr style="font-weight:bold">' +
      '<td colspan="4" bgcolor="' + SUBBG + '" align="right" style="border:1px solid ' + SUBBG + ';background:' + SUBBG + ';padding:4px 6px;text-align:right;color:#1B2430;font-weight:bold">รวมแผนก' + esc_(g.department) + '</td>' +
      regCell_(fmtBaht_(gt.hours), 'right', SUBBG) + regCell_('', '', SUBBG) +
      regCell_(fmtBaht_(gt.normal_pay), 'right', SUBBG) + regCell_(fmtBaht_(gt.ot_pay), 'right', SUBBG) +
      regCell_(fmtBaht_(gt.diligence + gt.addition + gt.bonus), 'right', SUBBG) + regCell_(fmtBaht_(gt.gross), 'right', SUBBG) +
      regCell_(fmtBaht_(gt.wht), 'right', SUBBG) + regCell_(fmtBaht_(gt.deduction), 'right', SUBBG) +
      regCell_(fmtBaht_(gt.net), 'right', SUBBG) + '</tr>';
  });
  var GRANDBG = '#E3F7EC';
  body += '<tr style="font-weight:bold">' +
    '<td colspan="4" bgcolor="' + GRANDBG + '" align="right" style="border:1px solid ' + GRANDBG + ';background:' + GRANDBG + ';padding:5px 6px;text-align:right;color:#1B2430;font-weight:bold">รวมทั้งสิ้น ( ' + data.headcount + ' คน )</td>' +
    regCell_(fmtBaht_(t.hours), 'right', GRANDBG) + regCell_('', '', GRANDBG) +
    regCell_(fmtBaht_(t.normal_pay), 'right', GRANDBG) + regCell_(fmtBaht_(t.ot_pay), 'right', GRANDBG) +
    regCell_(fmtBaht_(t.diligence + t.addition + t.bonus), 'right', GRANDBG) + regCell_(fmtBaht_(t.gross), 'right', GRANDBG) +
    regCell_(fmtBaht_(t.wht), 'right', GRANDBG) + regCell_(fmtBaht_(t.deduction), 'right', GRANDBG) +
    regCell_(fmtBaht_(t.net), 'right', GRANDBG) + '</tr>';

  var foot = '<table style="margin-top:12px"><tr>' +
    '<td class="muted" style="font-size:10px">หมายเหตุ<br>รายงานนี้จัดทำขึ้นเพื่อการตรวจสอบข้อมูลค่าจ้างพนักงาน Part-time เท่านั้น</td>' +
    '<td class="right" style="font-size:10px">.............................................. ผู้จัดทำรายงาน<br>' +
      '<span class="muted">วันที่พิมพ์ : ' + esc_(thaiDate_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'))) +
      ' เวลา ' + Utilities.formatDate(new Date(), TZ, 'HH:mm') + ' น.</span></td></tr></table>';

  return '<div>' + head + '<table style="margin-top:9px;font-size:10px">' + thead + body + '</table>' + foot + '</div>';
}

/* ================= สรุปค่าใช้จ่าย (รายปี) ================= */
function apiCostSummary(pin, year) {
  guard_(pin);
  var rows = readSheet_('Attendance'); var adjs = readSheet_('Adjustments');
  var rateMap = {}; readSheet_('Employees').forEach(function (e) { rateMap[e.employee_id] = e; });
  year = String(year || Utilities.formatDate(new Date(), TZ, 'yyyy'));
  var mBase = [], mOt = []; for (var i = 0; i < 12; i++) { mBase.push(0); mOt.push(0); }
  var base = 0, otPay = 0;
  rows.forEach(function (r) {
    if (String(r.work_date).slice(0, 4) !== year) return;
    var h = calcRowHours_(r); var e = rateMap[r.employee_id];
    var rate = e ? num_(e.hourly_rate) : 0, otr = e ? otRate_(e) : 0;
    var b = h.normal * rate, o = h.ot * otr; base += b; otPay += o;
    var mi = parseInt(String(r.work_date).slice(5, 7), 10) - 1; if (mi >= 0 && mi < 12) { mBase[mi] += b; mOt[mi] += o; }
  });
  var addition = 0, bonus = 0, deduction = 0, diligence = 0;
  adjs.forEach(function (a) {
    if (a.period_from && String(a.period_from).slice(0, 4) !== year) return;
    addition += num_(a.addition); bonus += num_(a.bonus); diligence += num_(a.diligence);
    deduction += num_(a.deduction) + num_(a.wht);            // ภาษีหัก ณ ที่จ่าย นับเป็นรายการหักด้วย
  });
  addition += diligence;                                      // เบี้ยขยันรวมอยู่ในกลุ่ม "เงินเพิ่ม" ของกราฟ
  var cost = [
    { label: 'ค่าจ้างพื้นฐาน', value: round2_(base) }, { label: 'ค่าโอที', value: round2_(otPay) },
    { label: 'เงินเพิ่ม', value: round2_(addition) }, { label: 'โบนัส', value: round2_(bonus) },
  ].filter(function (x) { return x.value > 0; });
  return {
    year: year, cost: cost, deduction: round2_(deduction), total: round2_(base + otPay + addition + bonus - deduction),
    base: round2_(base), otPay: round2_(otPay), addition: round2_(addition), bonus: round2_(bonus),
    trend: mBase.map(function (v, i) { return { m: TH_MON[i], base: round2_(v), ot: round2_(mOt[i]), total: round2_(v + mOt[i]) }; }),
  };
}

/* ================= บันทึกกิจกรรม (Import_Log) ================= */
function apiActivityLog(pin) { guard_(pin); return readSheet_('Import_Log').slice(-120).reverse(); }

function apiSaveSettings(pin, obj) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Settings'); var data = sh.getDataRange().getValues();
  var idx = {}; for (var i = 1; i < data.length; i++) idx[data[i][0]] = i + 1;
  Object.keys(obj).forEach(function (k) { if (idx[k]) sh.getRange(idx[k], 2).setValue(obj[k]); else sh.appendRow([k, obj[k]]); });
  SETTINGS_CACHE_ = null;
  return { ok: true };
}

/* ---- เงินเพิ่ม/โบนัส/หัก ต่อคนต่องวด ---- */
/** ช่วงเวลาเข้า-ออกที่พนักงานทำบ่อยที่สุดในงวด → "09:00 – 18:30"
 *  ถ้าในงวดมีหลายช่วงเวลา ต่อท้ายว่าเป็นเวลาส่วนใหญ่กี่วันจากกี่วัน
 *  ("09:00 – 18:30 (ส่วนใหญ่ 15/20 วัน)") กันคนอ่านสลิปเข้าใจว่าทำเวลานี้ทุกวัน */
function topShift_(shifts) {
  var best = '', n = 0, total = 0, kinds = 0;
  Object.keys(shifts || {}).forEach(function (k) {
    total += shifts[k]; kinds++;
    if (shifts[k] > n) { n = shifts[k]; best = k; }
  });
  if (!best) return '';
  var p = best.split('|'), a = hhmmPad_(p[0]), b = hhmmPad_(p[1]);
  var txt = a && b ? a + ' – ' + b : (a || b || '');
  if (!txt) return '';
  return kinds > 1 ? txt + ' (ส่วนใหญ่ ' + n + '/' + total + ' วัน)' : txt;
}
/** "9:00" → "09:00" (ในชีตพิมพ์มาไม่เหมือนกัน แต่บนสลิปควรเรียงตรงกัน) */
function hhmmPad_(v) {
  var m = String(v == null ? '' : v).match(/(\d{1,2})[:.](\d{2})/);
  return m ? pad_(m[1]) + ':' + m[2] : '';
}

function adjKey_(id, from, to) { return id + '|' + (from || '') + '|' + (to || ''); }
function readAdjustments_() { var m = {}; readSheet_('Adjustments').forEach(function (a) { m[adjKey_(a.employee_id, a.period_from, a.period_to)] = a; }); return m; }
function apiSaveAdjustment(pin, a) {
  guard_(pin);
  var sh = getSs_().getSheetByName('Adjustments'); var data = sh.getDataRange().getValues();
  var row = [a.employee_id, a.period_from || '', a.period_to || '', num_(a.addition), num_(a.bonus), num_(a.deduction), a.note || '',
    num_(a.diligence), num_(a.wht)];
  for (var i = 1; i < data.length; i++)
    if (data[i][0] === a.employee_id && String(data[i][1]) === String(a.period_from || '') && String(data[i][2]) === String(a.period_to || '')) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: true };
    }
  sh.appendRow(row); return { ok: true };
}

/* ---- คำนวณเงินเดือน ---- */
function calcPayroll_(period) {
  var emps = readSheet_('Employees'); var empMap = {}; emps.forEach(function (e) { empMap[e.employee_id] = e; });
  var posMap = {}; readSheet_('Positions').forEach(function (p) { posMap[p.position_id] = p; });
  var adj = readAdjustments_();
  var rows = readSheet_('Attendance').filter(function (r) {
    if (period.from && r.work_date < period.from) return false;
    if (period.to && r.work_date > period.to) return false;
    if (period.onlyApproved && r.status !== 'อนุมัติแล้ว') return false;
    return true;
  });
  var byEmp = {};
  rows.forEach(function (r) {
    var key = r.employee_id || r.employee_name || '(ไม่ระบุ)';
    if (!byEmp[key]) byEmp[key] = { normal: 0, ot: 0, days: {}, name: r.employee_name || r.display_name, shifts: {} };
    var h = calcRowHours_(r); byEmp[key].normal += h.normal; byEmp[key].ot += h.ot;
    if (r.work_date) byEmp[key].days[r.work_date] = 1;
    // เก็บช่วงเวลาเข้า-ออกไว้นับความถี่ → สลิปแสดงเวลาที่ทำประจำในงวดนั้น
    if (r.check_in || r.check_out) {
      var sk = (r.check_in || '') + '|' + (r.check_out || '');
      byEmp[key].shifts[sk] = (byEmp[key].shifts[sk] || 0) + 1;
    }
  });
  var result = [];
  Object.keys(byEmp).forEach(function (key) {
    var b = byEmp[key]; var e = empMap[key] || {};
    var rate = num_(e.hourly_rate), otr = otRate_(e);
    var a = adj[adjKey_(key, period.from, period.to)] || {};
    var addition = num_(a.addition), bonus = num_(a.bonus), deduction = num_(a.deduction);
    var diligence = num_(a.diligence), wht = num_(a.wht);
    var normalPay = b.normal * rate, otPay = b.ot * otr;
    var gross = normalPay + otPay + diligence + addition + bonus;   // รายได้รวม (ก่อนหัก)
    var totalDeduct = wht + deduction;                              // รายการหักทั้งสิ้น
    var pos = posMap[e.position] || {};
    result.push({ employee_id: key, employee_name: e.employee_name || b.name || key, days: Object.keys(b.days).length,
      normal_hours: round2_(b.normal), ot_hours: round2_(b.ot), hourly_rate: rate, ot_rate: otr,
      normal_pay: round2_(normalPay), ot_pay: round2_(otPay), diligence: round2_(diligence), addition: round2_(addition),
      bonus: round2_(bonus), gross: round2_(gross), wht: round2_(wht), deduction: round2_(deduction),
      total_deduct: round2_(totalDeduct), note: a.note || '', net: round2_(gross - totalDeduct),
      // ข้อมูลพนักงานที่สลิป/ทะเบียนเงินได้ต้องใช้ (ดึงมาให้ครบตั้งแต่ที่นี่ ฝั่งเรียกจะได้ไม่ต้องอ่านชีตซ้ำ)
      position: pos.position_name || '', department: pos.department || '', national_id: e.national_id || '',
      bank_name: e.bank_name || '', bank_account: e.bank_account || '', start_date: e.start_date || '',
      work_time: topShift_(b.shifts) });
  });
  result.sort(function (a, b) { return String(a.employee_id).localeCompare(String(b.employee_id)); });
  return result;
}
function apiCalcPayroll(pin, period) { guard_(pin); period = period || {}; return { period: period, rows: calcPayroll_(period) }; }

/* ---- สลิป PDF ---- */
function nextSlipNo_() { var n = parseInt(props_().getProperty('SLIP_SEQ') || '0', 10) + 1; props_().setProperty('SLIP_SEQ', String(n)); return n; }
function slipNo_(s) { return (s.slip_prefix || 'SLIP') + '-' + Utilities.formatDate(new Date(), TZ, 'yyMM') + '-' + ('000' + nextSlipNo_()).slice(-4); }

function apiPayslipPdf(pin, payload) {
  guard_(pin);
  var s = readSettings_(); var row = payload.row; var period = payload.period || {};
  var html = wrapHtml_(payslipBody_(s, row, period, slipNo_(s)));
  var pdf = Utilities.newBlob(html, 'text/html', 'slip.html').getAs('application/pdf');
  var name = 'Payslip_' + (row.employee_id || 'x') + '_' + (period.from || '') + '.pdf'; pdf.setName(name);
  try { subFolder_('05_Payslips').createFile(pdf); } catch (e) {}
  return { ok: true, name: name, base64: Utilities.base64Encode(pdf.getBytes()) };
}
function apiPayslipAll(pin, period) {
  guard_(pin);
  var s = readSettings_(); var rows = calcPayroll_(period || {});
  if (!rows.length) return { ok: false, error: 'ไม่มีข้อมูลในงวดนี้' };
  var body = rows.map(function (r) { return payslipBody_(s, r, period, slipNo_(s)); }).join('<div style="page-break-after:always"></div>');
  var pdf = Utilities.newBlob(wrapHtml_(body), 'text/html', 'slips.html').getAs('application/pdf');
  var name = 'Payslips_ALL_' + (period.from || '') + '_' + (period.to || '') + '.pdf'; pdf.setName(name);
  try { subFolder_('05_Payslips').createFile(pdf); } catch (e) {}
  return { ok: true, name: name, base64: Utilities.base64Encode(pdf.getBytes()), count: rows.length };
}

/* ---- ประวัติรายบุคคล ---- */
function apiEmployeeReport(pin, employeeId, period) {
  guard_(pin); period = period || {};
  var emp = readSheet_('Employees').filter(function (e) { return e.employee_id === employeeId; })[0] || { employee_id: employeeId };
  var history = readSheet_('Attendance').filter(function (r) {
    if (r.employee_id !== employeeId) return false;
    if (period.from && r.work_date < period.from) return false;
    if (period.to && r.work_date > period.to) return false;
    return true;
  }).map(function (r) { var h = calcRowHours_(r); r.normal_calc = h.normal; r.ot_calc = h.ot; return r; })
    .sort(function (a, b) { return (b.work_date || '').localeCompare(a.work_date || ''); });
  var pay = calcPayroll_({ from: period.from, to: period.to }).filter(function (p) { return p.employee_id === employeeId; })[0] || null;
  return { employee: emp, history: history, payroll: pay };
}

/* ---- ส่งออก CSV ---- */
function apiExportCsv(pin, type, arg) {
  guard_(pin);
  var rows, head;
  if (type === 'payroll') {
    head = ['รหัส', 'ชื่อ', 'ตำแหน่ง', 'แผนก', 'วัน', 'ชม.ปกติ', 'OT', 'ค่าแรง/ชม.', 'ค่าแรงปกติ', 'ค่าOT',
      'เบี้ยขยัน', 'เงินเพิ่ม', 'โบนัส', 'รายได้รวม', 'ภาษีหัก ณ ที่จ่าย', 'หักอื่นๆ', 'สุทธิ'];
    rows = calcPayroll_(arg || {}).map(function (r) {
      return [r.employee_id, r.employee_name, r.position, r.department, r.days, r.normal_hours, r.ot_hours, r.hourly_rate,
        r.normal_pay, r.ot_pay, r.diligence, r.addition, r.bonus, r.gross, r.wht, r.deduction, r.net];
    });
  } else {
    head = SHEETS.Attendance;
    rows = apiAttendance(pin, arg || {}).map(function (r) { return head.map(function (k) { return r[k]; }); });
  }
  var csv = [head].concat(rows).map(function (r) {
    return r.map(function (c) { c = (c == null ? '' : String(c)); return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c; }).join(',');
  }).join('\r\n');
  return { ok: true, name: type + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '.csv',
    base64: Utilities.base64Encode('﻿' + csv, Utilities.Charset.UTF_8) };
}

/* ================= helpers ================= */
// getValues() คืนคอลัมน์วันที่เป็น Date object — google.script.run serialize ไม่ผ่าน แล้วคืน null
// ให้ฝั่งเว็บ (เช่น d.employees พัง). แปลง Date → string ตั้งแต่ต้นทาง ทุก api ที่อ่าน sheet จึงปลอดภัยพร้อมกัน
function cellVal_(v) {
  if (v instanceof Date) {
    // ใช้ตัวอ่านส่วนประกอบตรง ๆ (ไม่ผ่าน formatDate+TZ) กันบั๊ก offset ประวัติศาสตร์ของ Asia/Bangkok
    // (ค่า "เวลาอย่างเดียว" ใน Sheet = epoch 1899-12-30 ซึ่ง Bangkok LMT เพี้ยน +6:42 → เวลาเลื่อน)
    var Y = v.getFullYear(), mo = pad_(v.getMonth() + 1), d = pad_(v.getDate());
    var h = v.getHours(), mi = v.getMinutes(), s = v.getSeconds();
    if (Y <= 1900) return pad_(h) + ':' + pad_(mi);                                   // เวลาอย่างเดียว → HH:mm
    if (h || mi || s) return Y + '-' + mo + '-' + d + ' ' + pad_(h) + ':' + pad_(mi) + ':' + pad_(s);
    return Y + '-' + mo + '-' + d;                                                    // วันที่อย่างเดียว → yyyy-MM-dd
  }
  return v;
}
function readSheet_(name) {
  var ss = getSs_(); var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); if (SHEETS[name]) sh.appendRow(SHEETS[name]); }
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var head = data[0], out = [];
  for (var i = 1; i < data.length; i++) { var o = {}; for (var c = 0; c < head.length; c++) o[head[c]] = cellVal_(data[i][c]); out.push(o); }
  return out;
}
function readSettings_() {
  if (!SETTINGS_CACHE_) {
    var o = {}; readSheet_('Settings').forEach(function (r) { o[r.key] = r.value; });
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) { if (o[k] == null || o[k] === '') o[k] = DEFAULT_SETTINGS[k]; });
    SETTINGS_CACHE_ = o;
  }
  return Object.assign({}, SETTINGS_CACHE_);   // คืนสำเนา กันผู้เรียกแก้ค่าในแคช
}
function otRate_(e) { var r = num_(e.ot_rate); if (r > 0) return r; var s = readSettings_(); return num_(e.hourly_rate) * (parseFloat(s.ot_multiplier) || 1.5); }
function calcRowHours_(r) {
  var normal = num_(r.normal_hours), ot = num_(r.ot_hours);
  if (normal > 0 || ot > 0) return { normal: normal, ot: ot };
  var ci = toMin_(r.check_in), co = toMin_(r.check_out);
  if (ci == null || co == null || co <= ci) return { normal: 0, ot: 0 };
  var s = readSettings_(); var brk = num_(r.break_minutes); if (!brk) brk = parseFloat(s.default_break_minutes) || 0;
  var worked = (co - ci - brk) / 60; if (worked < 0) worked = 0;
  var std = parseFloat(s.standard_day_hours) || 8;
  // ส่วนที่เกินชั่วโมงปกติจะนับเป็น OT ต่อเมื่อเกินขั้นต่ำที่ตั้งไว้ (ot_min_minutes)
  // กันเศษจากการออกช้านิดหน่อยกลายเป็น OT ทั้งที่ไม่มีใครขอ — ตั้ง 0 = นับทุกนาทีเหมือนเดิม
  var extra = Math.max(0, worked - std);
  var minOt = (parseFloat(s.ot_min_minutes) || 0) / 60;
  return { normal: round2_(Math.min(worked, std)), ot: extra >= minOt ? round2_(extra) : 0 };
}
function toMin_(x) { if (!x) return null; var m = String(x).match(/(\d{1,2})[:.](\d{2})/); return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : null; }
function normDate_(d) {
  if (!d) return ''; var m = String(d).match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) { var y = parseInt(m[1], 10); if (y > 2500) y -= 543; return y + '-' + pad_(m[2]) + '-' + pad_(m[3]); }
  m = String(d).match(/(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m) { var y2 = parseInt(m[3], 10); if (y2 > 2500) y2 -= 543; return y2 + '-' + pad_(m[2]) + '-' + pad_(m[1]); }
  return '';
}
function pad_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
function num_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function round2_(n) { return Math.round(n * 100) / 100; }
function unique_(a) { var s = {}; a.forEach(function (x) { if (x) s[x] = 1; }); return Object.keys(s); }
function fmtBaht_(n) { return num_(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function esc_(v) { return String(v == null ? '' : v).replace(/[&<>]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]; }); }

/* ---- วันที่แบบไทย (พ.ศ.) สำหรับเอกสารพิมพ์ ---- */
var TH_MON_FULL = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
function thaiDate_(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); if (!m) return String(iso || '-');
  return parseInt(m[3], 10) + ' ' + TH_MON_FULL[parseInt(m[2], 10) - 1] + ' ' + (parseInt(m[1], 10) + 543);
}
/** "1 – 15 กรกฎาคม 2567" ถ้าอยู่เดือนเดียวกัน · ไม่งั้นเขียนเต็มทั้งสองฝั่ง */
function thaiPeriod_(from, to) {
  var a = String(from || '').match(/^(\d{4})-(\d{2})-(\d{2})/), b = String(to || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!a || !b) return thaiDate_(from) + ' – ' + thaiDate_(to);
  if (a[1] === b[1] && a[2] === b[2]) return parseInt(a[3], 10) + ' – ' + thaiDate_(to);
  return thaiDate_(from) + ' – ' + thaiDate_(to);
}

/* ================= เอกสารพิมพ์ (สลิป / ทะเบียนเงินได้) =================
   ตัวแปลง HTML→PDF ของ Apps Script รองรับ flex/grid ไม่ครบ → จัดหน้าด้วย <table> + inline style เท่านั้น */
var DOC_CSS =
  'body{font-family:Arial,"Sarabun",sans-serif;color:#1B2430;font-size:11.5px;margin:0}' +
  'table{border-collapse:collapse;width:100%}td,th{vertical-align:top}' +
  '.right{text-align:right}.center{text-align:center}.muted{color:#5C6470}.b{font-weight:bold}' +
  '.navy{background:#123A6B;color:#fff}.blue{background:#1F4E9C;color:#fff}.red{background:#C0392B;color:#fff}' +
  '.green{color:#1E8449}.box{border:1px solid #C9D6E8;border-radius:8px}' +
  '.hd{padding:6px 9px;font-weight:bold;font-size:11.5px}' +
  '.ln td{border-bottom:1px solid #E4EAF3;padding:4px 9px}' +
  '.kv td{padding:2.5px 0;font-size:11.5px}' +
  '@page{size:A4 portrait;margin:10mm}';

/* ---------- สลิปเงินเดือน (ดีไซน์ v2) ----------
   ตัวแปลง HTML→PDF ของ Drive "ไม่อ่าน CSS class ใน <style>" (พื้นสี/แถบสีหายหมด)
   → กล่องและแถบสีทุกจุดในสลิปต้องเขียนเป็น inline style + attribute bgcolor เท่านั้น */
var SC = { navy: '#123A6B', blue: '#1F4E9C', red: '#C0392B', green: '#1E8449', line: '#C9D6E8', soft: '#E4EAF3', muted: '#5C6470', ink: '#1B2430', tint: '#EEF3FC' };

function wrapHtml_(body) {
  return '<html><head><meta charset="utf-8"><style>' + DOC_CSS +
    '@page{size:A4 landscape;margin:9mm}</style></head><body>' + body + '</body></html>';
}

/** แถบหัวกล่องสี (เต็มความกว้าง) — ใช้ bgcolor ด้วย เพราะ background ใน CSS class ไม่ถูกอ่าน */
function slipHead_(bg, left, right) {
  return '<tr><td bgcolor="' + bg + '" colspan="2" style="background:' + bg + ';color:#fff;font-weight:bold;font-size:13px;padding:8px 12px">' + left + '</td>' +
    '<td bgcolor="' + bg + '" align="right" style="background:' + bg + ';color:#fff;font-weight:bold;font-size:10.5px;padding:8px 12px;white-space:nowrap">' + (right || '') + '</td></tr>';
}
/** แถวรายการในกล่องรายได้/รายการหัก */
function slipRow_(no, label, amount) {
  return '<tr><td style="width:20px;color:#8A94A6;font-size:12.5px;border-bottom:1px solid ' + SC.soft + ';padding:6px 12px">' + no + '</td>' +
    '<td style="white-space:nowrap;font-size:12.5px;border-bottom:1px solid ' + SC.soft + ';padding:6px 4px">' + esc_(label) + '</td>' +
    '<td align="right" style="white-space:nowrap;font-size:12.5px;border-bottom:1px solid ' + SC.soft + ';padding:6px 12px">' + amount + '</td></tr>';
}
/** แถว "รวม…" ท้ายกล่อง */
function slipTot_(label, amount, color) {
  var st = 'font-weight:bold;font-size:13px;color:' + color + ';border-top:1px solid ' + SC.line + ';padding:9px 12px;white-space:nowrap';
  return '<tr><td colspan="2" style="' + st + '">' + esc_(label) + '</td>' +
    '<td align="right" style="' + st + '">' + amount + '</td></tr>';
}
/** ป้าย : ค่า ในการ์ดข้อมูลพนักงาน */
function slipKv_(label, value) {
  return '<tr><td style="width:132px;color:' + SC.muted + ';font-size:12.5px;padding:4px 0">' + esc_(label) + '</td>' +
    '<td style="width:10px;color:' + SC.muted + ';font-size:12.5px;padding:4px 0">:</td>' +
    '<td style="font-weight:bold;font-size:12.5px;padding:4px 0">' + esc_(value == null || value === '' ? '-' : value) + '</td></tr>';
}
/** กล่องขาวขอบฟ้า (ใช้แทน .box ที่ converter ไม่อ่าน) */
function slipBox_(inner, extra) {
  return '<table style="border:1px solid ' + SC.line + ';border-collapse:collapse;width:100%;' + (extra || '') + '">' + inner + '</table>';
}

function payslipBody_(s, row, period, slipNo) {
  var payDate = s.pay_date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var hours = num_(row.normal_hours) + num_(row.ot_hours);

  // ---- หัวเอกสาร: โลโก้ + ชื่อบริษัท | ชื่อเอกสาร + กล่องงวด/วันที่จ่าย ----
  var head = '<table style="width:100%;border-collapse:collapse"><tr>' +
    '<td style="width:96px;vertical-align:top">' + (getLogo_() ? '<img src="' + getLogo_() + '" style="height:56px">' : '') + '</td>' +
    '<td style="padding-left:8px;vertical-align:top">' +
      '<div style="font-size:21px;font-weight:bold;color:' + SC.navy + '">' + esc_(s.company_name) + '</div>' +
      (s.company_name_en ? '<div style="font-size:13px;font-weight:bold;color:' + SC.navy + ';margin-top:1px">' + esc_(s.company_name_en) + '</div>' : '') +
      '<div style="color:' + SC.muted + ';font-size:11px;margin-top:5px">' + esc_(s.company_address || '') + '</div>' +
      '<div style="color:' + SC.muted + ';font-size:11px;margin-top:2px">' + (s.tax_id ? 'เลขประจำตัวผู้เสียภาษีอากร&nbsp; ' + esc_(s.tax_id) : '') +
        (s.company_phone ? '&nbsp;&nbsp;&nbsp; โทร. ' + esc_(s.company_phone) : '') + '</div>' +
    '</td>' +
    '<td style="width:330px;vertical-align:top" align="right">' +
      '<div style="font-size:25px;font-weight:bold;color:' + SC.navy + ';line-height:1.15">สลิปเงินเดือนพนักงาน</div>' +
      '<div style="font-size:16px;font-weight:bold;color:' + SC.navy + '">(Part-time)</div>' +
      '<table style="width:100%;border-collapse:collapse;margin-top:9px"><tr>' +
        '<td bgcolor="' + SC.navy + '" style="background:' + SC.navy + ';color:#fff;padding:10px 14px;font-size:12.5px;line-height:1.7">' +
        'งวดวันที่ : ' + esc_(thaiPeriod_(period.from, period.to)) + '<br>วันที่จ่าย&nbsp; : ' + esc_(thaiDate_(payDate)) +
      '</td></tr></table>' +
    '</td></tr></table>';

  // ---- การ์ดข้อมูลพนักงาน + กล่องรายได้สุทธิ ----
  var netCard = '<table style="width:100%;border-collapse:collapse;border:2px solid ' + SC.navy + '">' +
    '<tr><td bgcolor="' + SC.navy + '" align="center" style="background:' + SC.navy + ';color:#fff;padding:8px 6px">' +
      '<b style="font-size:14px">รายได้สุทธิ</b><br><span style="font-size:11px">(Net Pay)</span></td></tr>' +
    '<tr><td align="center" style="padding:12px 6px">' +
      '<span style="color:' + SC.green + ';font-weight:bold;font-size:27px">' + fmtBaht_(row.net) + '</span><br>' +
      '<span style="color:' + SC.green + ';font-size:12px">บาท</span></td></tr></table>';

  var info = slipBox_('<tr>' +
    '<td style="padding:14px 18px;width:36%;vertical-align:top"><table style="width:100%;border-collapse:collapse">' +
      slipKv_('รหัสพนักงาน', row.employee_id) + slipKv_('ชื่อ-นามสกุล', row.employee_name) +
      slipKv_('ตำแหน่ง', row.position) + slipKv_('แผนก', row.department) +
      slipKv_('วันเริ่มงาน', row.start_date ? thaiDate_(row.start_date) : '') +
    '</table></td>' +
    '<td style="padding:14px 18px;width:39%;vertical-align:top"><table style="width:100%;border-collapse:collapse">' +
      slipKv_('เลขประจำตัวประชาชน', row.national_id) +
      slipKv_('อัตราค่าจ้าง', fmtBaht_(row.hourly_rate) + ' บาท/ชั่วโมง') +
      slipKv_('จำนวนชั่วโมงทำงาน', fmtBaht_(hours) + ' ชั่วโมง') +
      slipKv_('จำนวนวันทำงาน', row.days + ' วัน') +
      slipKv_('เลขที่สลิป', slipNo) +
    '</table></td>' +
    '<td style="padding:12px 14px;width:25%;vertical-align:top">' + netCard + '</td></tr>', 'margin-top:12px');

  // ---- 3 กล่อง: รายได้ / รายการหัก / สรุป ----
  var earn = slipBox_(
    slipHead_(SC.blue, 'รายได้ (Earnings)', 'จำนวนเงิน (บาท)') +
    slipRow_(1, 'ค่าจ้างรายชั่วโมง', fmtBaht_(row.normal_pay)) +
    slipRow_(2, 'ค่าล่วงเวลา', fmtBaht_(row.ot_pay)) +
    slipRow_(3, 'ค่าเบี้ยขยัน', fmtBaht_(row.diligence)) +
    slipRow_(4, 'ค่าอื่นๆ', fmtBaht_(num_(row.addition) + num_(row.bonus))) +
    '<tr><td colspan="3" style="height:34px"></td></tr>' +
    slipTot_('รวมรายได้ทั้งสิ้น', fmtBaht_(row.gross), SC.blue));

  var deduct = slipBox_(
    slipHead_(SC.red, 'รายการหัก (Deductions)', 'จำนวนเงิน (บาท)') +
    slipRow_(1, 'ภาษีเงินได้ ณ ที่จ่าย', fmtBaht_(row.wht)) +
    slipRow_(2, 'อื่นๆ', fmtBaht_(row.deduction)) +
    '<tr><td colspan="3" style="height:34px"></td></tr>' +
    slipTot_('รวมรายการหักทั้งสิ้น', fmtBaht_(row.total_deduct), SC.red));

  var sumLine = function (label, amount) {
    return '<tr><td style="white-space:nowrap;font-size:12.5px;border-bottom:1px solid ' + SC.soft + ';padding:8px 12px">' + label + '</td>' +
      '<td align="right" style="white-space:nowrap;font-size:12.5px;border-bottom:1px solid ' + SC.soft + ';padding:8px 12px">' + amount + '</td></tr>';
  };
  var bank = (row.bank_account || row.bank_name) ?
    '<tr><td colspan="2" style="padding:10px 12px">' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #D6E2F5"><tr>' +
      '<td bgcolor="' + SC.tint + '" style="background:' + SC.tint + ';padding:10px 12px">' +
        '<b style="color:' + SC.navy + ';font-size:12.5px">เงินโอนเข้าบัญชี</b>' +
        '<div style="color:' + SC.muted + ';font-size:11.5px;margin-top:3px">' + esc_(row.bank_name || '-') + '&nbsp; เลขที่บัญชี</div>' +
        '<div style="color:' + SC.ink + ';font-size:12.5px;font-weight:bold">' + esc_(row.bank_account || '-') + '</div>' +
      '</td></tr></table></td></tr>' : '';

  var summary = slipBox_(
    '<tr><td bgcolor="' + SC.navy + '" colspan="2" align="center" style="background:' + SC.navy + ';color:#fff;font-weight:bold;font-size:13px;padding:8px 12px">สรุปเงินได้ - รายการหัก</td></tr>' +
    sumLine('รวมรายได้ทั้งสิ้น', fmtBaht_(row.gross)) +
    sumLine('รวมรายการหักทั้งสิ้น', fmtBaht_(row.total_deduct)) +
    '<tr><td style="font-weight:bold;color:' + SC.green + ';font-size:13.5px;padding:10px 12px;border-top:2px solid ' + SC.navy + ';white-space:nowrap">รายได้สุทธิ (Net Pay)</td>' +
      '<td align="right" style="font-weight:bold;color:' + SC.green + ';font-size:16px;padding:10px 12px;border-top:2px solid ' + SC.navy + ';white-space:nowrap">' + fmtBaht_(row.net) + '</td></tr>' +
    bank);

  var cols = '<table style="width:100%;border-collapse:separate;border-spacing:0;margin-top:12px"><tr>' +
    '<td style="width:34%;padding-right:10px;vertical-align:top">' + earn + '</td>' +
    '<td style="width:33%;padding-right:10px;vertical-align:top">' + deduct + '</td>' +
    '<td style="width:33%;vertical-align:top">' + summary + '</td></tr></table>';

  // ---- หมายเหตุ + ลงชื่อผู้รับเงิน ----
  var note = slipBox_('<tr>' +
    '<td style="padding:14px 18px;width:55%;vertical-align:top"><b style="color:' + SC.navy + ';font-size:13px">หมายเหตุ</b>' +
      '<div style="color:' + SC.muted + ';margin-top:6px;font-size:11.5px;line-height:1.8">' + esc_(row.note || s.slip_note || '') +
      (s.company_phone ? '<br>หากมีข้อสงสัยกรุณาติดต่อฝ่ายบุคคล โทร. ' + esc_(s.company_phone) : '') + '</div></td>' +
    '<td align="center" style="padding:20px 18px 14px;font-size:12.5px;line-height:2">' +
      'ลงชื่อผู้รับเงิน ................................................................<br>' +
      '<span style="color:' + SC.muted + '">(&nbsp; ................................................&nbsp; )</span><br>' +
      'วันที่ ............... / ............... / ...............</td>' +
    '</tr>', 'margin-top:12px');

  var footBits = [s.company_address, s.company_phone, s.company_email, s.company_website]
    .filter(function (x) { return x; }).map(function (x) { return esc_(x); }).join('&nbsp;&nbsp; · &nbsp;&nbsp;');
  var foot = footBits ? '<table style="width:100%;border-collapse:collapse;margin-top:12px"><tr>' +
    '<td bgcolor="' + SC.navy + '" align="center" style="background:' + SC.navy + ';color:#fff;padding:9px;font-size:11px">' + footBits + '</td></tr></table>' : '';

  return '<div style="padding:2px">' + head + info + cols + note + foot + '</div>';
}
function authorizeUrlFetch() {
  const response = UrlFetchApp.fetch(
    "https://www.google.com",
    { muteHttpExceptions: true }
  );

  Logger.log(response.getResponseCode());
}
