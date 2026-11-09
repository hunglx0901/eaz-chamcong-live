/**
 * ============================================================
 *  CHẤM CÔNG & CA LIVE — Google Apps Script Web App (Backend)
 *  File: Code.gs
 *
 *  - Tự tạo Google Sheets database ở lần chạy đầu tiên
 *    (lưu ID vào Script Properties) => deploy là chạy ngay.
 *  - CRUD đầy đủ: addShift / getShifts / updateShift / deleteShift
 *  - Mọi hàm trả về { success: Boolean, data?, error? } => không nuốt lỗi.
 * ============================================================
 */

const SHEET_NAME = 'CaLive';
const HEADERS = ['ID', 'Ngày', 'Tên', 'Vị trí', 'Kênh', 'Bắt đầu', 'Kết thúc', 'Giờ live', 'GMV', 'Ghi chú', 'Thời gian tạo', 'GMV đầu ca', 'GMV cuối ca', 'Ảnh đầu ca', 'Ảnh cuối ca'];
const PHOTO_RETENTION_DAYS = 45;

const STAFF_SHEET = 'NhanSu';
const STAFF_HEADERS = ['Tên', 'Thời gian tạo', 'Mã NV'];

const CHANNEL_SHEET = 'Kenh';
const CHANNEL_HEADERS = ['Tên kênh', 'Thời gian tạo'];
const DEFAULT_CHANNELS = ['Gooby Vietnam', 'Gooby HN', 'Rouya Vietnam'];

const SALARY_SHEET = 'Luong';
const SALARY_HEADERS = ['Tên', 'Vị trí', 'Đơn giá giờ', '% GMV'];

// Yêu cầu sửa ca (nhân viên gửi -> Owner/Lead duyệt)
const REQ_SHEET = 'YeuCau';
const REQ_HEADERS = ['ID', 'Mã ca', 'Người gửi', 'Username', 'Kênh', 'Ngày ca', 'Thời gian gửi',
  'KT cũ', 'KT mới', 'GMV đầu cũ', 'GMV đầu mới', 'GMV cuối cũ', 'GMV cuối mới',
  'Ghi chú cũ', 'Ghi chú mới', 'Trạng thái', 'Người duyệt', 'Thời gian duyệt', 'Lý do từ chối',
  'Vị trí cũ', 'Vị trí mới', 'Kênh cũ', 'Kênh mới',
  'BĐ cũ', 'BĐ mới'];

/** Phục vụ trang web. Nếu có ?code=... (callback TikTok OAuth) thì xử lý riêng. */
function doGet(e) {
  // ===== Callback từ TikTok sau khi cấp quyền (Live Data) =====
  if (e && e.parameter && e.parameter.code) {
    return handleTikTokOAuthCallback_(e.parameter.code, e.parameter.state);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Chấm công & Ca live')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Cho phép nhúng CSS.html / JavaScript.html vào Index.html. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * (Tuỳ chọn) Chạy thủ công 1 lần trong editor để cấp quyền + tạo sẵn DB
 * trước khi deploy. Không bắt buộc.
 */
function setup() {
  const sheet = getSheet_();
  return 'DB sẵn sàng: ' + sheet.getParent().getUrl();
}

/* ----------------------- DATABASE ----------------------- */

/* TỐI ƯU: nhớ Spreadsheet + các sheet đã mở TRONG 1 LẦN GỌI (mỗi lần gọi từ trình duyệt là 1 lần chạy mới nên không bị cũ).
   Trước đây mỗi helper đều mở lại file bằng openById + đọc Script Properties + getLastRow... hàng chục lần mỗi thao tác. */
var MEMO_SS_ = null;
var MEMO_SH_ = {};
var MEMO_SECRET_ = null;
function ss_() { return getSheet_().getParent(); }

/** Lấy (hoặc tạo) sheet dữ liệu. Tự tạo spreadsheet nếu chưa có. */
function getSheet_() {
  if (MEMO_SH_['__main__']) return MEMO_SH_['__main__'];
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  let ss = null;

  if (id) {
    try { ss = MEMO_SS_ || SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('DB - Chấm công Livestream');
    props.setProperty('SPREADSHEET_ID', ss.getId());
    // Dùng luôn sheet đầu tiên làm sheet dữ liệu (không để sheet rác).
    ss.getSheets()[0].setName(SHEET_NAME);
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  let isNew = false;
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); isNew = true; }

  // Định dạng cột chỉ thiết lập khi sheet còn trống (tránh ghi đè định dạng đã có).
  if (sheet.getLastRow() < 1) {
    isNew = true;
    sheet.setFrozenRows(1);
    sheet.getRange('B:B').setNumberFormat('@');   // Ngày
    sheet.getRange('F:G').setNumberFormat('@');   // Bắt đầu / Kết thúc
    sheet.getRange('K:K').setNumberFormat('@');   // Thời gian tạo
    sheet.getRange('I:I').setNumberFormat('#,##0'); // GMV
    sheet.getRange('L:M').setNumberFormat('#,##0'); // GMV đầu ca / cuối ca
    sheet.setColumnWidth(1, 230);
  }
  // TỐI ƯU: chỉ đồng bộ tiêu đề khi sheet MỚI TẠO hoặc THIẾU CỘT (thêm cột mới sau update code) —
  // trước đây ghi lại header + định dạng trên MỌI lần gọi (hàm này chạy hàng chục lần mỗi lần tải trang), rất lãng phí.
  if (isNew || sheet.getLastColumn() < HEADERS.length) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#1a1d26').setFontColor('#ffffff');
  }
  MEMO_SS_ = ss;
  MEMO_SH_['__main__'] = sheet;
  return sheet;
}

/* ----------------------- HELPERS ----------------------- */

function computeHours_(start, end) {
  start = asTimeStr_(start); end = asTimeStr_(end);
  if (!start || !end) return 0;
  const a = String(start).split(':').map(Number);
  const b = String(end).split(':').map(Number);
  if (isNaN(a[0]) || isNaN(b[0])) return 0;
  let diff = (b[0] * 60 + (b[1] || 0)) - (a[0] * 60 + (a[1] || 0));
  if (diff < 0) diff += 1440; // ca qua đêm (trùng giờ = 0, không phải 24)
  return Math.round(diff / 60 * 100) / 100;
}

function validate_(rec) {
  if (!rec) return 'Dữ liệu rỗng.';
  if (!rec.date) return 'Thiếu ngày live.';
  if (!rec.name || !String(rec.name).trim()) return 'Thiếu tên nhân sự.';
  if (!rec.position) return 'Thiếu vị trí.';
  if (!rec.channel) return 'Thiếu kênh.';
  if (!rec.start || !rec.end) return 'Thiếu giờ bắt đầu hoặc kết thúc.';
  if ((Number(rec.gmvEnd) || 0) < (Number(rec.gmvStart) || 0)) return 'GMV cuối ca phải ≥ GMV đầu ca.';
  return '';
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function nowString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/** Chuẩn hoá ô NGÀY về 'yyyy-MM-dd' dù ô lưu là Date hay chuỗi. */
function asDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}

/** Chuẩn hoá ô GIỜ về 'HH:mm' dù ô lưu là Date hay chuỗi. */
function asTimeStr_(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  var s = String(v).trim();
  // Đã đúng dạng HH:mm hoặc H:mm -> chuẩn hoá 2 chữ số
  var m = s.match(/^(\d{1,2}):(\d{1,2})(?::\d{1,2})?$/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + ('0' + m[2]).slice(-2);
  // Dữ liệu cũ có thể là chuỗi Date đầy đủ, vd:
  //  "Sat Dec 30 1899 21:03:00 GMT+0700 (GMT+07:00)"  hoặc  "1899-12-30T14:03:00.000Z"
  // -> tách HH:mm ở bất kỳ vị trí nào trong chuỗi.
  var m2 = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?(?!\d)/);
  if (m2) {
    var hh = Math.max(0, Math.min(23, parseInt(m2[1], 10)));
    var mm = Math.max(0, Math.min(59, parseInt(m2[2], 10)));
    return ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2);
  }
  return s;
}

/* ----------------------- SHEET NHÂN SỰ ----------------------- */

/** Lấy (hoặc tạo) sheet NhanSu. */
function getStaffSheet_() {
  const ss = ss_(); // đảm bảo spreadsheet đã tồn tại
  let sheet = ss.getSheetByName(STAFF_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(STAFF_SHEET);
    sheet.getRange(1, 1, 1, STAFF_HEADERS.length).setValues([STAFF_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, STAFF_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a1d26').setFontColor('#ffffff');
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('B:B').setNumberFormat('@');
    sheet.setColumnWidth(1, 220);
  }
  return sheet;
}

/** Tìm dòng của 1 nhân sự theo tên (không phân biệt hoa thường). */
function staffRow_(sheet, name) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const target = String(name || '').trim().toLowerCase();
  const names = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

/** READ — danh sách nhân sự (lọc theo quyền). */
function getStaff(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const sheet = getStaffSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, data: [], full: [] };
    const seen = {};
    let full = sheet.getRange(2, 1, last - 1, 3).getValues()
      .map(function (r) { return { name: String(r[0]).trim(), code: String(r[2] || '').trim() }; })
      .filter(function (x) { if (!x.name || seen[x.name.toLowerCase()]) return false; seen[x.name.toLowerCase()] = true; return true; })
      .sort(function (a, b) { return a.name.localeCompare(b.name, 'vi'); });
    if (auth.role === 'staff') {
      const nm = String(auth.staff).toLowerCase();
      full = full.filter(function (x) { return x.name.toLowerCase() === nm; });
    }
    return { success: true, data: full.map(function (x) { return x.name; }), full: full };
  } catch (e) {
    return { success: false, error: 'Lỗi đọc nhân sự: ' + e.message };
  }
}

/** CREATE — thêm nhân sự (Owner hoặc Lead). */
function addStaff(token, name, code) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner' && auth.role !== 'lead') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    name = String(name || '').trim();
    code = String(code || '').trim();
    if (!name) return { success: false, error: 'Tên nhân sự không được để trống.' };
    const sheet = getStaffSheet_();
    if (staffRow_(sheet, name) > 0) {
      // Trùng tên: cần mã để phân biệt -> lưu định danh "Tên (Mã)"
      if (!code) return { success: false, error: 'Tên "' + name + '" đã tồn tại. Nhập Mã NV để phân biệt (hệ thống sẽ lưu thành "' + name + ' (Mã)").' };
      name = name + ' (' + code + ')';
      if (staffRow_(sheet, name) > 0) return { success: false, error: 'Nhân sự "' + name + '" đã tồn tại.' };
    }
    sheet.appendRow([name, nowString_(), code]);
    return { success: true, name: name };
  } catch (e) {
    return { success: false, error: 'Lỗi thêm nhân sự: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** DELETE — xoá nhân sự (chỉ Owner). */
function deleteStaff(token, name) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getStaffSheet_();
    const row = staffRow_(sheet, name);
    if (row < 0) return { success: false, error: 'Không tìm thấy nhân sự.' };
    sheet.deleteRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi xoá nhân sự: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Tự đăng ký nhân sự mới khi chấm công lần đầu (không chặn việc lưu ca). */
function ensureStaff_(name) {
  try {
    name = String(name || '').trim();
    if (!name) return;
    const sheet = getStaffSheet_();
    if (staffRow_(sheet, name) < 0) sheet.appendRow([name, nowString_()]);
  } catch (e) { /* bỏ qua, không ảnh hưởng lưu ca */ }
}

/** Tự thêm kênh vào sheet Kenh nếu chưa có (dùng khi import). */
function ensureChannel_(name) {
  try {
    name = String(name || '').trim();
    if (!name) return;
    const sheet = getChannelSheet_();
    if (channelRow_(sheet, name) < 0) sheet.appendRow([name, nowString_()]);
  } catch (e) { /* bỏ qua */ }
}

/* ----------------------- SHEET KÊNH ----------------------- */

/** Lấy (hoặc tạo) sheet Kenh, seed sẵn các kênh mặc định ở lần đầu. */
function getChannelSheet_() {
  const ss = ss_();
  let sheet = ss.getSheetByName(CHANNEL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CHANNEL_SHEET);
    sheet.getRange(1, 1, 1, CHANNEL_HEADERS.length).setValues([CHANNEL_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CHANNEL_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a1d26').setFontColor('#ffffff');
    sheet.getRange('A:B').setNumberFormat('@');
    sheet.setColumnWidth(1, 220);
    DEFAULT_CHANNELS.forEach(function (c) { sheet.appendRow([c, nowString_()]); });
  }
  return sheet;
}

function channelRow_(sheet, name) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const target = String(name || '').trim().toLowerCase();
  const names = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

/** READ — danh sách kênh (Lead chỉ thấy kênh của mình). */
function getChannels(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const sheet = getChannelSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, data: [] };
    let names = sheet.getRange(2, 1, last - 1, 1).getValues()
      .map(function (r) { return String(r[0]).trim(); })
      .filter(function (v, i, a) { return v && a.indexOf(v) === i; })
      .sort(function (a, b) { return a.localeCompare(b, 'vi'); });
    if (auth.role === 'lead') {
      const set = {};
      leadChannelsLower_(auth).forEach(function (c) { set[c] = true; });
      names = names.filter(function (n) { return set[n.toLowerCase()]; });
    }
    return { success: true, data: names };
  } catch (e) {
    return { success: false, error: 'Lỗi đọc kênh: ' + e.message };
  }
}

/** CREATE — thêm kênh (chỉ Owner). */
function addChannel(token, name) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    name = String(name || '').trim();
    if (!name) return { success: false, error: 'Tên kênh không được để trống.' };
    const sheet = getChannelSheet_();
    if (channelRow_(sheet, name) > 0) return { success: false, error: 'Kênh "' + name + '" đã tồn tại.' };
    sheet.appendRow([name, nowString_()]);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi thêm kênh: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** DELETE — xoá kênh (chỉ Owner). */
function deleteChannel(token, name) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getChannelSheet_();
    const row = channelRow_(sheet, name);
    if (row < 0) return { success: false, error: 'Không tìm thấy kênh.' };
    sheet.deleteRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi xoá kênh: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/* ----------------------- SHEET LƯƠNG ----------------------- */

/** Lấy (hoặc tạo) sheet Luong (đơn giá giờ + % GMV theo từng người + vị trí). */
function getSalarySheet_() {
  const ss = ss_();
  let sheet = ss.getSheetByName(SALARY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SALARY_SHEET);
    sheet.getRange(1, 1, 1, SALARY_HEADERS.length).setValues([SALARY_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SALARY_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a1d26').setFontColor('#ffffff');
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('B:B').setNumberFormat('@');
    sheet.getRange('C:C').setNumberFormat('#,##0');
    sheet.getRange('D:D').setNumberFormat('0.###');
    sheet.setColumnWidth(1, 220);
    return sheet;
  }
  // Di trú từ định dạng cũ 3 cột [Tên, Đơn giá giờ, % GMV] -> 4 cột (chèn Vị trí, để trống = áp dụng chung)
  const hdr = sheet.getRange(1, 1, 1, Math.max(3, sheet.getLastColumn())).getValues()[0];
  let migrated = false;
  if (String(hdr[1]).indexOf('Đơn giá') >= 0) {
    sheet.insertColumnAfter(1); // cột Vị trí trống cho các dòng cũ
    migrated = true;
  }
  // TỐI ƯU: chỉ ghi lại header khi vừa migrate — không lặp lại trên mọi lần gọi
  if (migrated) {
    sheet.getRange(1, 1, 1, SALARY_HEADERS.length).setValues([SALARY_HEADERS]);
    sheet.getRange('B:B').setNumberFormat('@');
  }
  return sheet;
}

function salaryRow_(sheet, name, position) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const tn = String(name || '').trim().toLowerCase();
  const tp = String(position || '').trim().toLowerCase();
  const vals = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === tn &&
        String(vals[i][1]).trim().toLowerCase() === tp) return i + 2;
  }
  return -1;
}

/** READ — cấu hình lương (lọc theo quyền). */
function getSalary(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const sheet = getSalarySheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, data: [] };
    const vals = sheet.getRange(2, 1, last - 1, 4).getValues();
    let data = vals.filter(function (r) { return r[0]; }).map(function (r) {
      return { name: String(r[0]).trim(), position: String(r[1] || '').trim(), rate: Number(r[2]) || 0, pct: Number(r[3]) || 0 };
    });
    if (auth.role === 'lead') {
      const set = leadStaffSet_(auth);
      data = data.filter(function (x) { return set[x.name.toLowerCase()]; });
    } else if (auth.role === 'staff') {
      const nm = String(auth.staff).toLowerCase();
      data = data.filter(function (x) { return x.name.toLowerCase() === nm; });
    }
    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: 'Lỗi đọc lương: ' + e.message };
  }
}

/** UPSERT — lưu công thức lương (Owner mọi người; Lead chỉ nhân sự thuộc kênh mình). */
function saveSalary(token, name, position, rate, pct) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role === 'admin' || auth.role === 'staff') return denyErr_();
  if (auth.role === 'lead') {
    const set = leadStaffSet_(auth);
    if (!set[String(name).trim().toLowerCase()]) return denyErr_();
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    name = String(name || '').trim();
    position = String(position || '').trim();
    if (!name) return { success: false, error: 'Thiếu tên nhân sự.' };
    rate = Number(rate) || 0; pct = Number(pct) || 0;
    const sheet = getSalarySheet_();
    const row = salaryRow_(sheet, name, position);
    if (row < 0) sheet.appendRow([name, position, rate, pct]);
    else sheet.getRange(row, 1, 1, 4).setValues([[name, position, rate, pct]]);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi lưu lương: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/* ===================== LƯƠNG & THƯỞNG (V2 - theo KPI) ===================== */
// Sheet đơn giá giờ theo vai trò
const LUONGGIO_SHEET = 'LuongGio';
const LUONGGIO_HEADERS = ['Tên', 'Kênh', 'Đơn giá Host', 'Đơn giá Trợ live'];
// Mốc thưởng theo (Tên × Kênh × Vai trò)
const THUONGMOC_SHEET = 'ThuongMoc';
const THUONGMOC_HEADERS = ['Tên', 'Kênh', 'Vai trò', 'Mốc GMV/h', 'Mốc NMV'];
// Cấu hình chung (key-value)
const THUONGCHUNG_SHEET = 'ThuongChung';
const THUONGCHUNG_HEADERS = ['Khóa', 'Giá trị'];
const THUONGCHUNG_DEFAULT = {
  heso_gmvh: 0.4, heso_nmv: 0.4, heso_diem: 0.2,
  tran_gmvh: 0.48, tran_nmv: 0.48, tran_diem: 0.2,
  moc_diem: 4, vat: 1.08, pct_host: 0.5, pct_tro: 0.2,
  hoan_huy: 30
};
// Nhập theo tháng: hoàn huỷ (1 số/tháng) + điểm (theo nhân viên)
const HOANHUY_SHEET = 'HoanHuy';
const HOANHUY_HEADERS = ['Tháng', 'Tỷ lệ hoàn huỷ'];
const DIEM_SHEET = 'DiemThang';
const DIEM_HEADERS = ['Tháng', 'Tên', 'Điểm'];
const HOTRO_SHEET = 'ThuongHoTro';
const HOTRO_HEADERS = ['Tháng', 'Tên', 'Số tiền', 'Ghi chú'];

// Snapshot lương tháng đã chốt (không tính lại)
const LUONGCHOT_SHEET = 'LuongChot';
const LUONGCHOT_HEADERS = ['Tháng', 'Tên', 'Tổng lương giờ', 'Tổng thưởng', 'Thưởng hỗ trợ', 'Ghi chú hỗ trợ', 'Điểm', 'Tổng thu nhập', 'Chi tiết lines (JSON)', 'Chốt bởi', 'Thời gian chốt', 'Lương cứng', 'HoTro Items (JSON)', 'LuongCung Items (JSON)', 'Thử việc', 'ThuViec Items (JSON)'];

// Trạng thái chốt tháng
const THANGCHOT_SHEET = 'ThangChot';
const THANGCHOT_HEADERS = ['Tháng', 'Trạng thái', 'Người chốt', 'Thời gian chốt', 'Ghi chú', 'Tỷ lệ hoàn huỷ chốt'];

// Nhân sự lương cứng (không live)
const LUONGCUNG_STAFF_SHEET = 'NhanSuLuongCung';
const LUONGCUNG_STAFF_HEADERS = ['Tên', 'Vai trò', 'Ghi chú', 'Trạng thái', 'Thời gian tạo'];
// Lương cứng theo (Tháng × Tên × Kênh)
const LUONGCUNG_THANG_SHEET = 'LuongCungThang';
const LUONGCUNG_THANG_HEADERS = ['Tháng', 'Tên', 'Kênh', 'Số tiền', 'Ghi chú'];

function getSimpleSheet_(name, headers) {
  if (MEMO_SH_[name]) return MEMO_SH_[name];
  const ss = ss_();
  let sheet = ss.getSheetByName(name);
  let isNew = false;
  if (!sheet) {
    try { sheet = ss.insertSheet(name); }
    catch (e) { sheet = ss.getSheetByName(name); } // tránh race "sheet đã tồn tại"
    if (!sheet) sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    isNew = true;
  }
  // TỐI ƯU: chỉ ghi header + định dạng khi sheet MỚI TẠO hoặc thiếu cột (cần thêm cột mới sau update code).
  // Trước đây ghi lại header trên MỌI lần gọi (kể cả khi không đổi gì) — hàm này được gọi 7-8 lần
  // mỗi lần tính lương (ThuongChung, ThuongMoc, HoanHuy, DiemThang, LuongCungThang...), gây chậm không cần thiết.
  if (isNew || sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1d26').setFontColor('#fff');
    // Cột "Tháng" (nếu có) phải là TEXT — nếu để mặc định, Google sẽ tự parse "2026-06" thành Date
    // -> khi đọc lại vals[i][0] là Date object -> lookup fail -> lưu ra dòng mới -> ghi trùng, đọc sai
    if (headers[0] === 'Tháng') sheet.getRange('A:A').setNumberFormat('@');
  }
  MEMO_SH_[name] = sheet;
  return sheet;
}

/** Chuẩn hoá giá trị "Tháng" đọc từ Sheets về dạng "yyyy-MM" (dù ô lưu là Date hay text). */
function asMonthStr_(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  var s = String(v).trim();
  // Đã đúng dạng yyyy-MM
  var m = s.match(/^(\d{4})-(\d{2})$/);
  if (m) return s;
  // yyyy-MM-dd -> lấy 7 ký tự đầu
  var m2 = s.match(/^(\d{4})-(\d{2})-\d{2}/);
  if (m2) return s.slice(0, 7);
  return s;
}
function getLuongGioSheet_() {
  const ss = ss_();
  const existed = ss.getSheetByName(LUONGGIO_SHEET);
  let migrated = false;
  // MIGRATION: sheet cũ 3 cột (Tên, Đơn giá Host, Đơn giá Trợ live) -> chèn cột Kênh (rỗng = mặc định)
  if (existed) {
    const hdr = existed.getRange(1, 1, 1, Math.max(existed.getLastColumn(), 1)).getValues()[0];
    // Nếu tiêu đề cột 2 là "Đơn giá Host" -> chưa migrate -> chèn cột Kênh ở vị trí 2
    if (String(hdr[1] || '').trim().indexOf('Đơn giá Host') === 0) {
      existed.insertColumnBefore(2); // chèn cột trống ở vị trí 2
      existed.getRange('B:B').setNumberFormat('@');
      migrated = true;
    }
  }
  const sheet = getSimpleSheet_(LUONGGIO_SHEET, LUONGGIO_HEADERS);
  // TỐI ƯU: chỉ ép định dạng cột Kênh khi sheet MỚI TẠO hoặc VỪA migrate — không lặp lại trên mọi lần gọi
  // (hàm này chạy ít nhất 1 lần mỗi lần tính lương; ghi định dạng thừa mỗi lần rất lãng phí).
  if (!existed || migrated) sheet.getRange('B:B').setNumberFormat('@');
  // Lần đầu tạo: cố gắng chuyển dữ liệu từ sheet Luong cũ (Tên|Vị trí|Đơn giá giờ|%GMV)
  if (!existed) {
    try {
      const old = ss.getSheetByName(SALARY_SHEET);
      if (old && old.getLastRow() >= 2) {
        const vals = old.getRange(2, 1, old.getLastRow() - 1, 4).getValues();
        const byName = {};
        vals.forEach(function (r) {
          const nm = String(r[0] || '').trim(); if (!nm) return;
          const pos = String(r[1] || '').toLowerCase();
          const rate = Number(r[2]) || 0;
          byName[nm] = byName[nm] || { host: 0, tro: 0 };
          if (pos.indexOf('host') >= 0) byName[nm].host = rate;
          else if (pos.indexOf('trợ') >= 0 || pos.indexOf('tro') >= 0 || pos === 'la') byName[nm].tro = rate;
        });
        Object.keys(byName).forEach(function (nm) {
          sheet.appendRow([nm, '', byName[nm].host, byName[nm].tro]); // kênh trống = áp dụng chung
        });
      }
    } catch (e) { /* bỏ qua nếu lỗi migrate */ }
  }
  return sheet;
}

/** Đọc cấu hình chung (seed mặc định nếu thiếu khóa). */
function readThuongChung_() {
  const sheet = getSimpleSheet_(THUONGCHUNG_SHEET, THUONGCHUNG_HEADERS);
  const last = sheet.getLastRow();
  const cfg = {};
  if (last >= 2) {
    const vals = sheet.getRange(2, 1, last - 1, 2).getValues();
    vals.forEach(function (r) { if (r[0]) cfg[String(r[0]).trim()] = Number(r[1]) || 0; });
  }
  // Bổ sung khóa thiếu bằng mặc định + ghi lại
  let changed = false;
  Object.keys(THUONGCHUNG_DEFAULT).forEach(function (k) {
    if (cfg[k] === undefined) { cfg[k] = THUONGCHUNG_DEFAULT[k]; changed = true; }
  });
  if (changed) {
    const rows = Object.keys(cfg).map(function (k) { return [k, cfg[k]]; });
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  return cfg;
}

/* ----- LƯU CẤU HÌNH (Owner; Lead chỉ nhân sự thuộc kênh mình) ----- */

function payrollCanEditName_(auth, name) {
  if (auth.role === 'owner') return true;
  if (auth.role === 'lead') { const set = leadStaffSet_(auth); return !!set[String(name).trim().toLowerCase()]; }
  return false;
}

/** Lưu đơn giá giờ theo vai trò cho 1 nhân viên. */
/** Lưu đơn giá giờ theo (Tên × Kênh × Vai trò). channel rỗng = mặc định chung cho mọi kênh. */
function saveLuongGio(token, name, channel, rateHost, rateTro) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (!payrollCanEditName_(auth, name)) return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    name = String(name || '').trim();
    channel = String(channel || '').trim();
    if (!name) return { success: false, error: 'Thiếu tên nhân sự.' };
    if (channel && channelRow_(getChannelSheet_(), channel) < 0) {
      return { success: false, error: 'Kênh "' + channel + '" không tồn tại trong danh sách kênh. Vui lòng chọn từ dropdown thay vì nhập tay.' };
    }
    const sheet = getLuongGioSheet_();
    const last = sheet.getLastRow();
    let row = -1;
    if (last >= 2) {
      const vals = sheet.getRange(2, 1, last - 1, 2).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim().toLowerCase() === name.toLowerCase()
            && String(vals[i][1]).trim().toLowerCase() === channel.toLowerCase()) {
          row = i + 2; break;
        }
      }
    }
    const data = [name, channel, Number(rateHost) || 0, Number(rateTro) || 0];
    if (row < 0) sheet.appendRow(data); else sheet.getRange(row, 1, 1, 4).setValues([data]);
    return { success: true };
  } catch (e) { return { success: false, error: 'Lỗi lưu đơn giá: ' + e.message }; }
  finally { lock.releaseLock(); }
}

/** Xoá 1 dòng đơn giá riêng theo (Tên × Kênh). Dùng để dọn dữ liệu sai/rác (VD kênh bị lưu nhầm thành số).
 *  KHÔNG cho xoá dòng "mặc định" (channel rỗng) — dòng đó chỉ có thể sửa về 0, không xoá, vì mỗi người luôn cần có dòng mặc định. */
function deleteLuongGio(token, name, channel) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (!payrollCanEditName_(auth, name)) return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    name = String(name || '').trim();
    channel = String(channel || '').trim();
    if (!name) return { success: false, error: 'Thiếu tên nhân sự.' };
    if (!channel) return { success: false, error: 'Không thể xoá dòng mặc định — chỉ có thể sửa về 0.' };
    const sheet = getLuongGioSheet_();
    const last = sheet.getLastRow();
    if (last >= 2) {
      const vals = sheet.getRange(2, 1, last - 1, 2).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim().toLowerCase() === name.toLowerCase()
            && String(vals[i][1]).trim().toLowerCase() === channel.toLowerCase()) {
          sheet.deleteRow(i + 2);
          return { success: true };
        }
      }
    }
    return { success: false, error: 'Không tìm thấy dòng để xoá.' };
  } catch (e) { return { success: false, error: 'Lỗi xoá đơn giá: ' + e.message }; }
  finally { lock.releaseLock(); }
}

/** Lưu mốc thưởng cho (Tên × Kênh × Vai trò). role: 'Host' | 'Tro'. */
function saveThuongMoc(token, name, channel, role, mocGmvh, mocNmv) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (!payrollCanEditName_(auth, name)) return denyErr_();
  if (auth.role === 'lead' && !leadHasChannel_(auth, channel)) return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    name = String(name || '').trim(); channel = String(channel || '').trim();
    role = (String(role) === 'Host') ? 'Host' : 'Tro';
    if (!name || !channel) return { success: false, error: 'Thiếu tên hoặc kênh.' };
    const sheet = getSimpleSheet_(THUONGMOC_SHEET, THUONGMOC_HEADERS);
    const last = sheet.getLastRow();
    let row = -1;
    if (last >= 2) {
      const vals = sheet.getRange(2, 1, last - 1, 3).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim().toLowerCase() === name.toLowerCase() &&
            String(vals[i][1]).trim().toLowerCase() === channel.toLowerCase() &&
            String(vals[i][2]) === role) { row = i + 2; break; }
      }
    }
    const data = [name, channel, role, Number(mocGmvh) || 0, Number(mocNmv) || 0];
    if (row < 0) sheet.appendRow(data); else sheet.getRange(row, 1, 1, 5).setValues([data]);
    return { success: true };
  } catch (e) { return { success: false, error: 'Lỗi lưu mốc thưởng: ' + e.message }; }
  finally { lock.releaseLock(); }
}

/** Lưu cấu hình chung (chỉ Owner). cfg = object khóa→giá trị. */
function saveThuongChung(token, cfg) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSimpleSheet_(THUONGCHUNG_SHEET, THUONGCHUNG_HEADERS);
    const merged = readThuongChung_();
    Object.keys(cfg || {}).forEach(function (k) { if (merged[k] !== undefined) merged[k] = Number(cfg[k]) || 0; });
    const rows = Object.keys(merged).map(function (k) { return [k, merged[k]]; });
    if (sheet.getLastRow() >= 2) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    return { success: true };
  } catch (e) { return { success: false, error: 'Lỗi lưu cấu hình: ' + e.message }; }
  finally { lock.releaseLock(); }
}

/** Lưu tỷ lệ hoàn huỷ cho 1 tháng (YYYY-MM). rate dạng % (vd 5 = 5%). */
function saveHoanHuy(token, month, ratePct) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner' && auth.role !== 'lead') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ (YYYY-MM).' };
    const sheet = getSimpleSheet_(HOANHUY_SHEET, HOANHUY_HEADERS);
    const last = sheet.getLastRow();
    let row = -1;
    if (last >= 2) {
      const months = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < months.length; i++) if (asMonthStr_(months[i][0]) === month) { row = i + 2; break; }
    }
    const data = [month, Number(ratePct) || 0];
    if (row < 0) sheet.appendRow(data); else sheet.getRange(row, 1, 1, 2).setValues([data]);
    return { success: true };
  } catch (e) { return { success: false, error: 'Lỗi lưu hoàn huỷ: ' + e.message }; }
  finally { lock.releaseLock(); }
}

/** Lưu điểm phòng live của 1 nhân viên cho 1 tháng. */
function saveDiem(token, month, name, diem) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (!payrollCanEditName_(auth, name)) return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    month = String(month || '').trim(); name = String(name || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month) || !name) return { success: false, error: 'Thiếu tháng hoặc tên.' };
    const sheet = getSimpleSheet_(DIEM_SHEET, DIEM_HEADERS);
    const last = sheet.getLastRow();
    let row = -1;
    if (last >= 2) {
      const vals = sheet.getRange(2, 1, last - 1, 2).getValues();
      for (let i = 0; i < vals.length; i++)
        if (asMonthStr_(vals[i][0]) === month && String(vals[i][1]).trim().toLowerCase() === name.toLowerCase()) { row = i + 2; break; }
    }
    const data = [month, name, Number(diem) || 0];
    if (row < 0) sheet.appendRow(data); else sheet.getRange(row, 1, 1, 3).setValues([data]);
    return { success: true };
  } catch (e) { return { success: false, error: 'Lỗi lưu điểm: ' + e.message }; }
  finally { lock.releaseLock(); }
}

/**
 * BATCH: lưu hoàn huỷ + toàn bộ điểm trong 1 lần rồi trả về kết quả tính lương.
 * Thay cho việc gọi nhiều lần saveHoanHuy + saveDiem + computePayroll (đỡ ~90% thời gian).
 * @param diemList: [{name, diem}, ...]
 * @param hoTroList: [{name, amount, note}, ...]
 */
function saveMonthUpdate(token, month, ratePct, diemList, hoTroList, filterChannels) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner' && auth.role !== 'lead') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ (YYYY-MM).' };

    // ★ Chặn nếu tháng đã chốt (Owner vẫn được sửa)
    const monthLock = checkMonthLocked_(month);
    if (monthLock && monthLock.locked && auth.role !== 'owner') {
      return { success: false, error: 'Tháng ' + month + ' đã chốt bởi ' + monthLock.by + '. Chỉ Owner mới có thể sửa.' };
    }

    diemList = Array.isArray(diemList) ? diemList : [];
    hoTroList = Array.isArray(hoTroList) ? hoTroList : [];

    // 1) Lưu hoàn huỷ (1 read + 1 write)
    const hh = getSimpleSheet_(HOANHUY_SHEET, HOANHUY_HEADERS);
    {
      const last = hh.getLastRow();
      let row = -1;
      if (last >= 2) {
        const months = hh.getRange(2, 1, last - 1, 1).getValues();
        for (let i = 0; i < months.length; i++) if (asMonthStr_(months[i][0]) === month) { row = i + 2; break; }
      }
      const data = [month, Number(ratePct) || 0];
      if (row < 0) hh.appendRow(data); else hh.getRange(row, 1, 1, 2).setValues([data]);
    }

    // 2) BATCH lưu điểm: đọc tất cả 1 lần, cập nhật in-memory, ghi lại 1 lần
    if (diemList.length) {
      // Lọc quyền trên các item (Lead: chỉ NV thuộc phạm vi)
      const validItems = diemList.filter(function (it) {
        return it && it.name && payrollCanEditName_(auth, it.name);
      }).map(function (it) {
        return { name: String(it.name).trim(), diem: Number(it.diem) || 0 };
      });

      const dm = getSimpleSheet_(DIEM_SHEET, DIEM_HEADERS);
      const last = dm.getLastRow();
      // Đọc toàn bộ dữ liệu (nếu có) 1 lần duy nhất
      let existing = [];
      if (last >= 2) existing = dm.getRange(2, 1, last - 1, 3).getValues();

      // Map key (tháng|tên) -> index của dòng hiện có
      const rowByKey = {};
      for (let i = 0; i < existing.length; i++) {
        if (!existing[i][0]) continue;
        const k = asMonthStr_(existing[i][0]) + '|' + String(existing[i][1]).trim().toLowerCase();
        rowByKey[k] = i; // 0-based trong `existing`
      }

      const newRows = [];
      validItems.forEach(function (it) {
        const k = month + '|' + it.name.toLowerCase();
        const row = [month, it.name, it.diem];
        if (rowByKey[k] !== undefined) {
          existing[rowByKey[k]] = row; // cập nhật in-memory
        } else {
          newRows.push(row);
        }
      });

      // Ghi lại vùng cũ (nếu có thay đổi) — Sheets API cho phép setValues nguyên khối
      if (existing.length) dm.getRange(2, 1, existing.length, 3).setValues(existing);
      // Thêm dòng mới ở cuối (nếu có)
      if (newRows.length) dm.getRange(existing.length + 2, 1, newRows.length, 3).setValues(newRows);
    }

    // 2b) BATCH lưu thưởng hỗ trợ (v2: ID, Tháng, Tên, Kênh, Số tiền, Ghi chú)
    // Format cũ hỗ trợ: {name, amount, note} -> ghi channel = ''
    // Format mới: {name, channel, amount, note}
    if (hoTroList.length) {
      const hs = ensureHoTroV2_();
      // Với v2: cách xử lý đơn giản là XOÁ toàn bộ dòng hoTro của (tháng, tên) trong hoTroList
      // rồi ghi lại các dòng non-empty. Điều này cho phép user có nhiều dòng/người.
      const namesTouched = {};
      hoTroList.forEach(function (it) {
        if (!it || !it.name) return;
        if (!payrollCanEditName_(auth, it.name)) return;
        namesTouched[String(it.name).trim().toLowerCase()] = true;
      });
      // Xoá dòng cũ của các tên có trong hoTroList
      if (hs.getLastRow() >= 2) {
        const all = hs.getRange(2, 1, hs.getLastRow() - 1, 6).getValues();
        for (let i = all.length - 1; i >= 0; i--) {
          if (asMonthStr_(all[i][1]) !== month) continue;
          const nm = String(all[i][2] || '').trim().toLowerCase();
          if (namesTouched[nm]) hs.deleteRow(i + 2);
        }
      }
      // Ghi dòng mới (không rỗng)
      const rowsToAppend = [];
      hoTroList.forEach(function (it) {
        if (!it || !it.name) return;
        const nm = String(it.name).trim();
        if (!payrollCanEditName_(auth, nm)) return;
        const amt = Number(it.amount) || 0;
        const note = String(it.note || '');
        const ch = String(it.channel || '').trim();
        if (amt === 0 && !note && !ch) return; // bỏ dòng rỗng
        rowsToAppend.push([Utilities.getUuid(), month, nm, ch, amt, note]);
      });
      if (rowsToAppend.length) hs.getRange(hs.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    }

    // 3) Tính lương xong luôn -> chỉ 1 round-trip cho frontend
    // Nếu tháng đã chốt (Owner đang sửa) -> tính RAW từ config mới, cập nhật snapshot, rồi trả về
    if (monthLock && monthLock.locked && auth.role === 'owner') {
      const rawResult = _computePayrollRaw_(auth, month);
      if (rawResult.success && rawResult.data && rawResult.data.length) {
        // Cập nhật snapshot với dữ liệu mới
        try {
          const snapSheet = getSimpleSheet_(LUONGCHOT_SHEET, LUONGCHOT_HEADERS);
          // Xoá dòng cũ của tháng
          if (snapSheet.getLastRow() >= 2) {
            const monthCol = snapSheet.getRange(2, 1, snapSheet.getLastRow() - 1, 1).getValues();
            const rowsToDelete = [];
            for (let i = 0; i < monthCol.length; i++) if (asMonthStr_(monthCol[i][0]) === month) rowsToDelete.push(i + 2);
            for (let i = rowsToDelete.length - 1; i >= 0; i--) snapSheet.deleteRow(rowsToDelete[i]);
          }
          const now = nowString_();
          const snapRows = rawResult.data.map(function (p) {
            return [month, p.name, Number(p.hoursPay) || 0, Number(p.bonus) || 0,
              Number(p.hoTro) || 0, String(p.hoTroNote || ''), Number(p.diem) || 0,
              Number(p.total) || 0, JSON.stringify(p.lines || []),
              auth.username + ' (sửa sau chốt)', now,
              Number(p.luongCung) || 0,
              JSON.stringify(p.hoTroItems || []),
              JSON.stringify(p.luongCungItems || []),
              Number(p.probation) || 0,
              JSON.stringify(p.probationItems || [])];
          });
          snapSheet.getRange(snapSheet.getLastRow() + 1, 1, snapRows.length, LUONGCHOT_HEADERS.length).setValues(snapRows);
        } catch (e) { /* bỏ qua lỗi re-snapshot */ }
      }
      // Đọc lại theo computePayroll (sẽ đọc snapshot mới)
      return computePayroll(token, month, filterChannels);
    }
    return computePayroll(token, month, filterChannels);
  } catch (e) {
    return { success: false, error: 'Lỗi cập nhật tháng: ' + e.message };
  } finally { lock.releaseLock(); }
}

/* ==================== CHỐT LƯƠNG THÁNG (SNAPSHOT) ==================== */

/** Kiểm tra tháng đã chốt chưa. Trả về {locked, by, at, note, rate} hoặc null nếu chưa. */
function checkMonthLocked_(month) {
  try {
    const sheet = getSimpleSheet_(THANGCHOT_SHEET, THANGCHOT_HEADERS);
    if (sheet.getLastRow() < 2) return null;
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (asMonthStr_(rows[i][0]) === month && String(rows[i][1]).trim() === 'da_chot') {
        return {
          locked: true,
          by: String(rows[i][2] || ''),
          at: String(rows[i][3] || ''),
          note: String(rows[i][4] || ''),
          rate: Number(rows[i][5]) || 0
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

/** Public wrapper: trả về trạng thái chốt của tháng để frontend dùng. */
function getMonthLockStatus(token, month) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };
    const lock = checkMonthLocked_(month);
    return { success: true, month: month, lock: lock };
  } catch (e) { return { success: false, error: e.message }; }
}

/** Owner CHỐT lương tháng: snapshot toàn bộ kết quả vào sheet LuongChot, đánh dấu tháng "da_chot". */
function chotLuongThang(token, month, note) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return { success: false, error: 'Chỉ Owner mới có quyền chốt lương tháng.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ (YYYY-MM).' };

    // Kiểm chưa chốt
    const existed = checkMonthLocked_(month);
    if (existed && existed.locked) return { success: false, error: 'Tháng ' + month + ' đã được chốt trước đó (bởi ' + existed.by + ').' };

    // Tính lương hiện tại
    const payResult = computePayroll(token, month);
    if (!payResult.success) return payResult;
    const data = payResult.data || [];
    if (!data.length) return { success: false, error: 'Không có dữ liệu lương để chốt.' };

    // Ghi snapshot vào LuongChot (xoá dữ liệu cũ của tháng này nếu có)
    const snapSheet = getSimpleSheet_(LUONGCHOT_SHEET, LUONGCHOT_HEADERS);
    // Xoá dòng cũ của tháng (nếu có, do lần chốt trước đã mở)
    if (snapSheet.getLastRow() >= 2) {
      const monthCol = snapSheet.getRange(2, 1, snapSheet.getLastRow() - 1, 1).getValues();
      const rowsToDelete = [];
      for (let i = 0; i < monthCol.length; i++) if (asMonthStr_(monthCol[i][0]) === month) rowsToDelete.push(i + 2);
      // Xoá từ dưới lên
      for (let i = rowsToDelete.length - 1; i >= 0; i--) snapSheet.deleteRow(rowsToDelete[i]);
    }

    const now = nowString_();
    const snapRows = data.map(function (p) {
      return [
        month,
        p.name,
        Number(p.hoursPay) || 0,
        Number(p.bonus) || 0,
        Number(p.hoTro) || 0,
        String(p.hoTroNote || ''),
        Number(p.diem) || 0,
        Number(p.total) || 0,
        JSON.stringify(p.lines || []),
        auth.username,
        now,
        Number(p.luongCung) || 0,
        JSON.stringify(p.hoTroItems || []),
        JSON.stringify(p.luongCungItems || []),
        Number(p.probation) || 0,
        JSON.stringify(p.probationItems || [])
      ];
    });
    snapSheet.getRange(snapSheet.getLastRow() + 1, 1, snapRows.length, LUONGCHOT_HEADERS.length).setValues(snapRows);

    // Ghi/cập nhật trạng thái tháng
    const statusSheet = getSimpleSheet_(THANGCHOT_SHEET, THANGCHOT_HEADERS);
    const rate = Number(payResult.returnRate) || 0;
    let row = -1;
    if (statusSheet.getLastRow() >= 2) {
      const months = statusSheet.getRange(2, 1, statusSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < months.length; i++) if (asMonthStr_(months[i][0]) === month) { row = i + 2; break; }
    }
    const statusData = [month, 'da_chot', auth.username, now, String(note || ''), rate];
    if (row < 0) statusSheet.appendRow(statusData);
    else statusSheet.getRange(row, 1, 1, 6).setValues([statusData]);

    // Telegram
    try {
      sendTelegram_('🔒 <b>CHỐT LƯƠNG THÁNG ' + esc_(month) + '</b>\n👤 ' + esc_(auth.username)
        + '\n👥 ' + data.length + ' nhân viên · Tổng: ' + fmtVnd_(data.reduce(function (s, x) { return s + (x.total || 0); }, 0)) + 'đ'
        + (note ? '\n📝 ' + esc_(String(note)) : ''));
    } catch (e) { /* bỏ qua */ }

    return { success: true, count: data.length, at: now };
  } catch (e) {
    return { success: false, error: 'Lỗi chốt lương: ' + e.message };
  } finally { lock.releaseLock(); }
}

/** Owner MỞ CHỐT lương tháng để chỉnh sửa. */
function moChotLuongThang(token, month, reason) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return { success: false, error: 'Chỉ Owner mới có quyền mở chốt lương tháng.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };

    const statusSheet = getSimpleSheet_(THANGCHOT_SHEET, THANGCHOT_HEADERS);
    let row = -1;
    if (statusSheet.getLastRow() >= 2) {
      const months = statusSheet.getRange(2, 1, statusSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < months.length; i++) if (asMonthStr_(months[i][0]) === month) { row = i + 2; break; }
    }
    if (row < 0) return { success: false, error: 'Tháng ' + month + ' chưa được chốt.' };

    // Đổi trạng thái sang mo_chot, giữ lịch sử người mở + lý do
    const oldNote = String(statusSheet.getRange(row, 5).getValue() || '');
    const newNote = oldNote + '\n[Mở chốt lúc ' + nowString_() + ' bởi ' + auth.username + (reason ? ': ' + reason : '') + ']';
    statusSheet.getRange(row, 2).setValue('mo_chot');
    statusSheet.getRange(row, 5).setValue(newNote);

    try {
      sendTelegram_('🔓 <b>MỞ CHỐT LƯƠNG THÁNG ' + esc_(month) + '</b>\n👤 ' + esc_(auth.username) + (reason ? '\n📝 ' + esc_(String(reason)) : ''));
    } catch (e) { }

    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi mở chốt: ' + e.message };
  } finally { lock.releaseLock(); }
}

/* ==================== NHÂN SỰ LƯƠNG CỨNG ==================== */

/** Lấy danh sách toàn bộ nhân sự lương cứng. */
function listLuongCungStaff(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const sheet = getSimpleSheet_(LUONGCUNG_STAFF_SHEET, LUONGCUNG_STAFF_HEADERS);
    let list = [];
    if (sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(function (r) {
        if (!r[0]) return;
        list.push({
          name: String(r[0]).trim(),
          role: String(r[1] || ''),
          note: String(r[2] || ''),
          status: String(r[3] || 'active'),
          createdAt: String(r[4] || '')
        });
      });
    }
    list.sort(function (a, b) { return a.name.localeCompare(b.name, 'vi'); });
    return { success: true, data: list };
  } catch (e) { return { success: false, error: e.message }; }
}

/** Thêm nhân sự lương cứng. */
function addLuongCungStaff(token, name, role, note) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  try {
    name = String(name || '').replace(/\s+/g, ' ').trim();
    if (!name) return { success: false, error: 'Thiếu tên.' };
    const sheet = getSimpleSheet_(LUONGCUNG_STAFF_SHEET, LUONGCUNG_STAFF_HEADERS);
    // Check trùng
    if (sheet.getLastRow() >= 2) {
      const existed = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < existed.length; i++)
        if (String(existed[i][0]).trim().toLowerCase() === name.toLowerCase())
          return { success: false, error: 'Nhân sự "' + name + '" đã tồn tại.' };
    }
    sheet.appendRow([name, String(role || ''), String(note || ''), 'active', nowString_()]);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

/** Cập nhật nhân sự lương cứng (đổi vai trò/ghi chú/trạng thái). */
function updateLuongCungStaff(token, name, role, note, status) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  try {
    name = String(name || '').trim();
    if (!name) return { success: false, error: 'Thiếu tên.' };
    const sheet = getSimpleSheet_(LUONGCUNG_STAFF_SHEET, LUONGCUNG_STAFF_HEADERS);
    if (sheet.getLastRow() < 2) return { success: false, error: 'Không tìm thấy.' };
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < rows.length; i++)
      if (String(rows[i][0]).trim().toLowerCase() === name.toLowerCase()) { row = i + 2; break; }
    if (row < 0) return { success: false, error: 'Không tìm thấy "' + name + '".' };
    if (role != null) sheet.getRange(row, 2).setValue(String(role));
    if (note != null) sheet.getRange(row, 3).setValue(String(note));
    if (status != null) sheet.getRange(row, 4).setValue(String(status));
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

/** Xoá nhân sự lương cứng (và toàn bộ lương cứng theo tháng của người này). */
function deleteLuongCungStaff(token, name) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  try {
    name = String(name || '').trim();
    if (!name) return { success: false, error: 'Thiếu tên.' };
    const sheet = getSimpleSheet_(LUONGCUNG_STAFF_SHEET, LUONGCUNG_STAFF_HEADERS);
    if (sheet.getLastRow() >= 2) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = rows.length - 1; i >= 0; i--)
        if (String(rows[i][0]).trim().toLowerCase() === name.toLowerCase()) sheet.deleteRow(i + 2);
    }
    // Xoá luôn dữ liệu lương cứng theo tháng
    const thangSheet = getSimpleSheet_(LUONGCUNG_THANG_SHEET, LUONGCUNG_THANG_HEADERS);
    if (thangSheet.getLastRow() >= 2) {
      const rows = thangSheet.getRange(2, 2, thangSheet.getLastRow() - 1, 1).getValues();
      for (let i = rows.length - 1; i >= 0; i--)
        if (String(rows[i][0]).trim().toLowerCase() === name.toLowerCase()) thangSheet.deleteRow(i + 2);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

/** Lấy lương cứng theo tháng: trả list [{name, channel, amount, note}]. */
function getLuongCungByMonth(token, month) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };
    const sheet = getSimpleSheet_(LUONGCUNG_THANG_SHEET, LUONGCUNG_THANG_HEADERS);
    const list = [];
    if (sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(function (r) {
        if (!r[0] || !r[1]) return;
        if (asMonthStr_(r[0]) !== month) return;
        list.push({
          name: String(r[1]).trim(),
          channel: String(r[2] || '').trim(),
          amount: Number(r[3]) || 0,
          note: String(r[4] || '')
        });
      });
    }
    return { success: true, data: list };
  } catch (e) { return { success: false, error: e.message }; }
}

/** Batch lưu lương cứng cho tháng (thay thế toàn bộ dòng của tháng đó với danh sách mới). */
function saveLuongCungBatch(token, month, items) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };
    // Chặn nếu tháng đã chốt
    const monthLock = checkMonthLocked_(month);
    if (monthLock && monthLock.locked) return { success: false, error: 'Tháng ' + month + ' đã chốt lương, không sửa được lương cứng.' };
    items = Array.isArray(items) ? items : [];
    const sheet = getSimpleSheet_(LUONGCUNG_THANG_SHEET, LUONGCUNG_THANG_HEADERS);
    // Xoá tất cả dòng của tháng cũ
    if (sheet.getLastRow() >= 2) {
      const monthCol = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = monthCol.length - 1; i >= 0; i--)
        if (asMonthStr_(monthCol[i][0]) === month) sheet.deleteRow(i + 2);
    }
    // Ghi lại dòng mới (chỉ ghi item có amount > 0 hoặc note không rỗng)
    const rowsToAppend = [];
    items.forEach(function (it) {
      const nm = String(it.name || '').trim();
      const ch = String(it.channel || '').trim();
      const amt = Number(it.amount) || 0;
      const note = String(it.note || '');
      if (!nm || !ch) return;
      if (amt === 0 && !note) return; // bỏ qua dòng rỗng
      rowsToAppend.push([month, nm, ch, amt, note]);
    });
    if (rowsToAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
    return { success: true, saved: rowsToAppend.length };
  } catch (e) {
    return { success: false, error: 'Lỗi lưu lương cứng: ' + e.message };
  } finally { lock.releaseLock(); }
}

/* ==================== THƯỞNG HỖ TRỢ NHIỀU DÒNG (v2) ==================== */

/** Cập nhật HOTRO sheet để có cột ID + Kênh (v2). Migration tự động. */
const HOTRO_V2_HEADERS = ['ID', 'Tháng', 'Tên', 'Kênh', 'Số tiền', 'Ghi chú'];

/** Ensure HOTRO sheet đúng schema v2. */
function ensureHoTroV2_() {
  const ss = ss_();
  const sheet = ss.getSheetByName(HOTRO_SHEET);
  if (!sheet) {
    // Tạo mới với v2 headers
    const s = ss.insertSheet(HOTRO_SHEET);
    s.getRange(1, 1, 1, HOTRO_V2_HEADERS.length).setValues([HOTRO_V2_HEADERS]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HOTRO_V2_HEADERS.length).setFontWeight('bold').setBackground('#1a1d26').setFontColor('#fff');
    s.getRange('A:A').setNumberFormat('@');
    s.getRange('B:B').setNumberFormat('@');
    return s;
  }
  // Migration nếu cần
  const hdr = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const hasIdCol = String(hdr[0] || '').trim().toUpperCase() === 'ID';
  if (!hasIdCol) {
    // Cấu trúc cũ: Tháng, Tên, Số tiền, Ghi chú -> chèn cột ID ở đầu + cột Kênh sau Tên
    sheet.insertColumnBefore(1); // ID (empty)
    sheet.insertColumnAfter(3);  // Kênh sau Tên (col C là Tên, sau C là D = Kênh)
    // Set headers
    sheet.getRange(1, 1, 1, HOTRO_V2_HEADERS.length).setValues([HOTRO_V2_HEADERS]);
    // Gán ID cho dữ liệu cũ
    if (sheet.getLastRow() >= 2) {
      const n = sheet.getLastRow() - 1;
      const ids = [];
      for (let i = 0; i < n; i++) ids.push([Utilities.getUuid()]);
      sheet.getRange(2, 1, n, 1).setValues(ids);
    }
  }
  // TỐI ƯU: chỉ ép định dạng khi vừa migrate (thêm cột) — không lặp lại định dạng trên mọi lần gọi
  // (hàm này chạy ít nhất 1 lần mỗi lần tính lương).
  if (!hasIdCol) {
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('B:B').setNumberFormat('@');
  }
  return sheet;
}

/** Lấy tất cả dòng thưởng hỗ trợ của tháng. Trả về list [{id, name, channel, amount, note}]. */
function getHoTroByMonth(token, month) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };
    const sheet = ensureHoTroV2_();
    const list = [];
    if (sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().forEach(function (r) {
        if (!r[1]) return;
        if (asMonthStr_(r[1]) !== month) return;
        list.push({
          id: String(r[0] || ''),
          name: String(r[2]).trim(),
          channel: String(r[3] || '').trim(),
          amount: Number(r[4]) || 0,
          note: String(r[5] || '')
        });
      });
    }
    return { success: true, data: list };
  } catch (e) { return { success: false, error: e.message }; }
}

/** Batch lưu thưởng hỗ trợ (thay toàn bộ dòng của (tháng, tên) với list mới). Cho phép nhiều dòng/người/kênh. */
function saveHoTroBatch(token, month, items) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner' && auth.role !== 'lead') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };
    const monthLock = checkMonthLocked_(month);
    if (monthLock && monthLock.locked && auth.role !== 'owner') return { success: false, error: 'Tháng đã chốt, chỉ Owner sửa được.' };
    items = Array.isArray(items) ? items : [];
    const sheet = ensureHoTroV2_();
    // Nhóm items theo tên -> mỗi tên xoá dòng cũ (của cùng tháng) rồi ghi mới
    const namesTouched = {};
    items.forEach(function (it) { const nm = String(it.name || '').trim().toLowerCase(); if (nm) namesTouched[nm] = true; });
    if (sheet.getLastRow() >= 2) {
      const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      for (let i = all.length - 1; i >= 0; i--) {
        const rowMonth = asMonthStr_(all[i][1]);
        const rowName = String(all[i][2] || '').trim().toLowerCase();
        if (rowMonth === month && namesTouched[rowName]) {
          if (auth.role === 'lead' && !payrollCanEditName_(auth, all[i][2])) continue;
          sheet.deleteRow(i + 2);
        }
      }
    }
    // Ghi dòng mới
    const rowsToAppend = [];
    items.forEach(function (it) {
      const nm = String(it.name || '').trim();
      if (!nm) return;
      if (auth.role === 'lead' && !payrollCanEditName_(auth, nm)) return;
      const ch = String(it.channel || '').trim();
      const amt = Number(it.amount) || 0;
      const note = String(it.note || '');
      // Bỏ dòng rỗng
      if (amt === 0 && !note && !ch) return;
      rowsToAppend.push([Utilities.getUuid(), month, nm, ch, amt, note]);
    });
    if (rowsToAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    return { success: true, saved: rowsToAppend.length };
  } catch (e) {
    return { success: false, error: 'Lỗi lưu thưởng hỗ trợ: ' + e.message };
  } finally { lock.releaseLock(); }
}

/** Đọc snapshot lương đã chốt của tháng. Trả về danh sách theo cùng format với computePayroll.data. */
function readLuongChotSnapshot_(month) {
  const sheet = getSimpleSheet_(LUONGCHOT_SHEET, LUONGCHOT_HEADERS);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, LUONGCHOT_HEADERS.length).getValues();
  const list = [];
  rows.forEach(function (r) {
    if (asMonthStr_(r[0]) !== month) return;
    let lines = [], hoTroItems = [], luongCungItems = [];
    try { lines = JSON.parse(r[8] || '[]'); } catch (e) { }
    try { hoTroItems = JSON.parse(r[12] || '[]'); } catch (e) { }
    try { luongCungItems = JSON.parse(r[13] || '[]'); } catch (e) { }
    let probationItems = [];
    try { probationItems = JSON.parse(r[15] || '[]'); } catch (e) { }
    list.push({
      name: String(r[1]),
      hoursPay: Number(r[2]) || 0,
      bonus: Number(r[3]) || 0,
      hoTro: Number(r[4]) || 0,
      hoTroNote: String(r[5] || ''),
      hoTroItems: hoTroItems,
      diem: Number(r[6]) || 0,
      total: Number(r[7]) || 0,
      luongCung: Number(r[11]) || 0,
      luongCungItems: luongCungItems,
      probation: Number(r[14]) || 0,
      probationItems: probationItems,
      isProbation: (!lines.length && Number(r[14]) > 0),
      lines: lines,
      isFixed: (Number(r[11]) > 0 && Number(r[2]) === 0 && Number(r[3]) === 0)
    });
  });
  return list;
}

/* ----- ĐỌC CẤU HÌNH (cho frontend) ----- */
function getPayrollConfig(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const chung = readThuongChung_();
    // Phạm vi Lead: chỉ nhân viên + kênh mình quản lý
    let scopeNames = null, scopeChannels = null;
    if (auth.role === 'lead') {
      scopeNames = leadStaffSet_(auth);
      scopeChannels = {};
      leadChannelsLower_(auth).forEach(function (c) { scopeChannels[c] = true; });
    }
    const lg = getLuongGioSheet_();
    const hourly = [];
    if (lg.getLastRow() >= 2) {
      lg.getRange(2, 1, lg.getLastRow() - 1, 4).getValues().forEach(function (r) {
        if (!r[0]) return;
        const nm = String(r[0]).trim();
        const ch = String(r[1] || '').trim();
        if (scopeNames && !scopeNames[nm.toLowerCase()]) return;
        if (ch && scopeChannels && !scopeChannels[ch.toLowerCase()]) return;
        hourly.push({ name: nm, channel: ch, host: Number(r[2]) || 0, tro: Number(r[3]) || 0 });
      });
    }
    const mocSheet = getSimpleSheet_(THUONGMOC_SHEET, THUONGMOC_HEADERS);
    const moc = [];
    if (mocSheet.getLastRow() >= 2) {
      mocSheet.getRange(2, 1, mocSheet.getLastRow() - 1, 5).getValues().forEach(function (r) {
        if (!r[0]) return;
        const nm = String(r[0]).trim(), ch = String(r[1]).trim();
        if (scopeNames && !scopeNames[nm.toLowerCase()]) return;
        if (scopeChannels && !scopeChannels[ch.toLowerCase()]) return;
        moc.push({ name: nm, channel: ch, role: String(r[2]), mocGmvh: Number(r[3]) || 0, mocNmv: Number(r[4]) || 0 });
      });
    }

    // MỚI: Danh sách kênh từng nhân viên đã LIVE (để hiện config theo kênh)
    // Quét TẤT CẢ ca (kể cả ca đang mở) — vì mục đích liệt kê kênh
    const staffChannels = {}; // nameLower -> {name, channels: {}}
    const sheetCa = getSheet_();
    if (sheetCa.getLastRow() >= 2) {
      // Đọc thêm cột G (Kết thúc) để log, và cột đầy đủ cho robust
      const rows = sheetCa.getRange(2, 3, sheetCa.getLastRow() - 1, 3).getValues(); // C: Tên, D: Vị trí, E: Kênh
      rows.forEach(function (r) {
        var nm = String(r[0] == null ? '' : r[0]).replace(/\s+/g, ' ').trim();
        var ch = String(r[2] == null ? '' : r[2]).replace(/\s+/g, ' ').trim();
        if (!nm || !ch) return;
        var nl = nm.toLowerCase();
        // Nếu Lead: giới hạn tên NV theo scope
        if (scopeNames && !scopeNames[nl]) return;
        // Nếu Lead: giới hạn kênh theo scope kênh của Lead
        if (scopeChannels && !scopeChannels[ch.toLowerCase()]) return;
        if (!staffChannels[nl]) staffChannels[nl] = { name: nm, channels: {} };
        staffChannels[nl].channels[ch] = true;
      });
    }
    const staffChannelList = Object.keys(staffChannels).map(function (k) {
      return { name: staffChannels[k].name, channels: Object.keys(staffChannels[k].channels).sort() };
    }).sort(function (a, b) { return a.name.localeCompare(b.name, 'vi'); });

    return {
      success: true, chung: chung, hourly: hourly, moc: moc,
      staffChannels: staffChannelList,
      scopeNames: scopeNames ? Object.keys(scopeNames) : null,
      scopeChannels: scopeChannels ? Object.keys(scopeChannels) : null
    };
  } catch (e) { return { success: false, error: 'Lỗi đọc cấu hình lương: ' + e.message }; }
}

/* ----- ENGINE: TÍNH LƯƠNG + THƯỞNG CHO 1 THÁNG ----- */
function computePayroll(token, month, filterChannels) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ (YYYY-MM).' };

    // ★ KIỂM TRA THÁNG ĐÃ CHỐT — nếu có, đọc snapshot thay vì tính lại
    const monthLock = checkMonthLocked_(month);
    if (monthLock && monthLock.locked) {
      let snapshot = readLuongChotSnapshot_(month);
      // Lọc theo quyền
      if (auth.role === 'lead') { const set = leadStaffSet_(auth); snapshot = snapshot.filter(function (x) { return set[x.name.toLowerCase()]; }); }
      else if (auth.role === 'staff') { const nm = String(auth.staff).toLowerCase(); snapshot = snapshot.filter(function (x) { return x.name.toLowerCase() === nm; }); }
      // ★ LỌC THEO KÊNH: áp dụng cả cho dữ liệu đã chốt (snapshot lưu "channel" trên từng dòng chi tiết)
      snapshot = filterPayrollList_(snapshot, buildAllowedChanSet_(auth, filterChannels));
      snapshot.sort(function (a, b) { return b.total - a.total; });
      return {
        success: true, month: month,
        returnRate: Math.round((monthLock.rate || 0) * 10000) / 100,
        data: snapshot,
        locked: true, lockInfo: monthLock
      };
    }
    // Không chốt -> tính raw
    return _computePayrollRaw_(auth, month, null, filterChannels);
  } catch (e) {
    return { success: false, error: 'Lỗi tính lương: ' + e.message };
  }
}

/** Tháng liền trước dạng "yyyy-MM" (bản backend của prevMonthStr phía frontend). */
function prevMonthStr_(m) {
  const p = String(m || '').split('-');
  let y = Number(p[0]), mo = Number(p[1]);
  if (!y || !mo) return '';
  mo -= 1;
  if (mo < 1) { mo = 12; y -= 1; }
  return y + '-' + ('0' + mo).slice(-2);
}

/** TỐI ƯU: tính lương tháng NÀY + tháng TRƯỚC (để so sánh) trong 1 LƯỢT GỌI DUY NHẤT —
 *  chỉ đọc toàn bộ CaLive 1 LẦN (dùng chung cho cả 2 tháng) thay vì 2 lần như gọi computePayroll() riêng rẽ.
 *  Giảm ~50% thời gian mở tab Bảng lương (bớt 1 lượt quét sheet lớn + 1 lượt round-trip client-server). */
function computePayrollWithPrev(token, month, filterChannels) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ (YYYY-MM).' };

    const allRows = readAllShiftsRaw_(); // ĐỌC CaLive 1 LẦN DUY NHẤT
    const allowedChanSet = buildAllowedChanSet_(auth, filterChannels);

    // --- Tháng hiện tại (tôn trọng snapshot nếu đã chốt) ---
    const curLock = checkMonthLocked_(month);
    let current;
    if (curLock && curLock.locked) {
      let snapshot = readLuongChotSnapshot_(month);
      if (auth.role === 'lead') { const set = leadStaffSet_(auth); snapshot = snapshot.filter(function (x) { return set[x.name.toLowerCase()]; }); }
      else if (auth.role === 'staff') { const nm = String(auth.staff).toLowerCase(); snapshot = snapshot.filter(function (x) { return x.name.toLowerCase() === nm; }); }
      snapshot = filterPayrollList_(snapshot, allowedChanSet);
      snapshot.sort(function (a, b) { return b.total - a.total; });
      current = { returnRate: Math.round((curLock.rate || 0) * 10000) / 100, data: snapshot, locked: true, lockInfo: curLock };
    } else {
      const raw = _computePayrollRaw_(auth, month, allRows, filterChannels);
      if (!raw.success) return raw;
      current = { returnRate: raw.returnRate, data: raw.data, locked: false, lockInfo: null };
    }

    // --- Tháng trước (chỉ lấy để so sánh — không cần trả returnRate/lock) ---
    const prevMonth = prevMonthStr_(month);
    let prevData = [];
    if (prevMonth) {
      const prevLock = checkMonthLocked_(prevMonth);
      if (prevLock && prevLock.locked) {
        prevData = readLuongChotSnapshot_(prevMonth);
        if (auth.role === 'lead') { const set = leadStaffSet_(auth); prevData = prevData.filter(function (x) { return set[x.name.toLowerCase()]; }); }
        else if (auth.role === 'staff') { const nm = String(auth.staff).toLowerCase(); prevData = prevData.filter(function (x) { return x.name.toLowerCase() === nm; }); }
        prevData = filterPayrollList_(prevData, allowedChanSet);
      } else {
        const prevRaw = _computePayrollRaw_(auth, prevMonth, allRows, filterChannels);
        prevData = prevRaw.success ? prevRaw.data : [];
      }
    }

    return {
      success: true, month: month,
      returnRate: current.returnRate, data: current.data, locked: current.locked, lockInfo: current.lockInfo,
      prevMonth: prevMonth, prevData: prevData
    };
  } catch (e) {
    return { success: false, error: 'Lỗi tính lương: ' + e.message };
  }
}

/* ==================== NHÂN SỰ THỬ VIỆC (nhập tay theo tháng, không chấm công) ==================== */
const PROBATION_SHEET = 'LuongThuViec';
const PROBATION_HEADERS = ['Tháng', 'Tên', 'Kênh', 'Số tiền', 'Ghi chú'];

/** Đọc các khoản lương thử việc của 1 tháng: [{name, channel, amount, note}]. */
function readProbationByMonth_(month) {
  const sheet = getSimpleSheet_(PROBATION_SHEET, PROBATION_HEADERS);
  const out = [];
  if (sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(function (r) {
    if (!r[1] || asMonthStr_(r[0]) !== month) return;
    out.push({
      name: String(r[1]).replace(/\s+/g, ' ').trim(),
      channel: String(r[2] || '').trim(),
      amount: Number(r[3]) || 0,
      note: String(r[4] || '')
    });
  });
  return out;
}

/** (Owner) Lấy danh sách lương thử việc của 1 tháng để chỉnh sửa. */
function getProbationByMonth(token, month) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  try {
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };
    return { success: true, data: readProbationByMonth_(month) };
  } catch (e) { return { success: false, error: e.message }; }
}

/** (Owner) Lưu toàn bộ lương thử việc của 1 tháng (thay thế các dòng cũ của tháng đó). */
function saveProbationBatch(token, month, items) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ.' };
    const monthLock = checkMonthLocked_(month);
    if (monthLock && monthLock.locked) return { success: false, error: 'Tháng ' + month + ' đã chốt lương. Mở chốt trước khi sửa lương thử việc.' };
    items = Array.isArray(items) ? items : [];
    const chanSheet = getChannelSheet_();
    const rowsToAppend = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i] || {};
      const nm = String(it.name || '').replace(/\s+/g, ' ').trim();
      const ch = String(it.channel || '').trim();
      const amt = Number(it.amount) || 0;
      const note = String(it.note || '');
      if (!nm && amt === 0 && !note) continue; // dòng trống hoàn toàn
      if (!nm) return { success: false, error: 'Dòng thứ ' + (i + 1) + ' thiếu tên nhân sự.' };
      if (amt === 0 && !note) continue;
      if (ch && channelRow_(chanSheet, ch) < 0) return { success: false, error: 'Kênh "' + ch + '" không có trong danh sách kênh.' };
      rowsToAppend.push([month, nm, ch, amt, note]);
    }
    const sheet = getSimpleSheet_(PROBATION_SHEET, PROBATION_HEADERS);
    if (sheet.getLastRow() >= 2) {
      const monthCol = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = monthCol.length - 1; i >= 0; i--) {
        if (asMonthStr_(monthCol[i][0]) === month) sheet.deleteRow(i + 2);
      }
    }
    if (rowsToAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
    return { success: true, saved: rowsToAppend.length };
  } catch (e) {
    return { success: false, error: 'Lỗi lưu lương thử việc: ' + e.message };
  } finally { lock.releaseLock(); }
}

/* ==================== NGÀY LỄ — NHÂN LƯƠNG GIỜ ==================== */
const HOLIDAY_SHEET = 'NgayLe';
const HOLIDAY_HEADERS = ['Ngày', 'Hệ số', 'Ghi chú'];

/** Đọc toàn bộ ngày lễ: [{date:'yyyy-MM-dd', mult, note}] đã sắp theo ngày. */
function readHolidaysAll_() {
  const sheet = getSimpleSheet_(HOLIDAY_SHEET, HOLIDAY_HEADERS);
  const out = [];
  if (sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function (r) {
    const d = asDateStr_(r[0]);
    const m = Number(r[1]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !(m > 0)) return;
    out.push({ date: d, mult: m, note: String(r[2] || '') });
  });
  out.sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  return out;
}

/** Bản đồ ngày lễ của 1 tháng: { 'yyyy-MM-dd': hệ số }. */
function readHolidayMapByMonth_(month) {
  const map = {};
  readHolidaysAll_().forEach(function (h) { if (h.date.slice(0, 7) === month) map[h.date] = h.mult; });
  return map;
}

/** (Owner) Danh sách ngày lễ đã cài. */
function getHolidays(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  try { return { success: true, data: readHolidaysAll_() }; }
  catch (e) { return { success: false, error: e.message }; }
}

/** (Owner) Lưu toàn bộ ngày lễ (thay thế danh sách cũ). Không cho đổi ngày lễ thuộc tháng ĐÃ CHỐT lương. */
function saveHolidays(token, items) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    items = Array.isArray(items) ? items : [];
    const seen = {};
    const rows = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i] || {};
      const d = String(it.date || '').trim();
      const m = Number(it.mult);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return { success: false, error: 'Dòng ' + (i + 1) + ': ngày không hợp lệ.' };
      if (!(m > 0) || m > 10) return { success: false, error: 'Ngày ' + d + ': hệ số phải lớn hơn 0 và không quá 10.' };
      if (seen[d]) return { success: false, error: 'Ngày ' + d + ' bị nhập trùng.' };
      seen[d] = true;
      rows.push({ date: d, mult: m, note: String(it.note || '') });
    }
    // Chặn đổi ngày lễ của tháng đã chốt (số liệu đã đóng băng)
    const canon = function (list, month) {
      return list.filter(function (h) { return h.date.slice(0, 7) === month; })
        .map(function (h) { return h.date + '|' + h.mult; }).sort().join(';');
    };
    const old = readHolidaysAll_();
    const months = {};
    old.concat(rows).forEach(function (h) { months[h.date.slice(0, 7)] = true; });
    const monthKeys = Object.keys(months);
    for (let i = 0; i < monthKeys.length; i++) {
      const mo = monthKeys[i];
      if (canon(old, mo) !== canon(rows, mo)) {
        const lk = checkMonthLocked_(mo);
        if (lk && lk.locked) return { success: false, error: 'Tháng ' + mo + ' đã chốt lương — Mở chốt trước khi đổi ngày lễ của tháng đó.' };
      }
    }
    rows.sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
    const sheet = getSimpleSheet_(HOLIDAY_SHEET, HOLIDAY_HEADERS);
    sheet.getRange('A:A').setNumberFormat('@'); // giữ ngày dạng chữ yyyy-MM-dd
    if (sheet.getLastRow() >= 2) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows.map(function (h) { return [h.date, h.mult, h.note]; }));
    return { success: true, saved: rows.length };
  } catch (e) {
    return { success: false, error: 'Lỗi lưu ngày lễ: ' + e.message };
  } finally { lock.releaseLock(); }
}

/** Đọc TOÀN BỘ sheet CaLive 1 lần — dùng chung khi cần tính lương nhiều tháng trong cùng 1 lượt gọi backend
 *  (VD tháng này + tháng trước để so sánh), tránh quét lại sheet nhiều lần không cần thiết. */
function readAllShiftsRaw_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
}

/** Tập kênh được phép xem: bắt buộc theo phạm vi Lead (nếu có) giao với lựa chọn lọc kênh tuỳ chọn (nếu có).
 *  Trả về null nếu không giới hạn gì (Owner, không lọc kênh). */
function buildAllowedChanSet_(auth, filterChannels) {
  let allowedChanSet = null;
  if (auth.role === 'lead') {
    allowedChanSet = {};
    leadChannelsLower_(auth).forEach(function (c) { allowedChanSet[c] = true; });
  }
  if (filterChannels && filterChannels.length) {
    const filterSet = {};
    filterChannels.forEach(function (c) { filterSet[String(c).toLowerCase()] = true; });
    if (allowedChanSet) {
      Object.keys(allowedChanSet).forEach(function (c) { if (!filterSet[c]) delete allowedChanSet[c]; });
    } else {
      allowedChanSet = filterSet;
    }
  }
  return allowedChanSet;
}

/** Lọc danh sách lương theo tập kênh cho phép (allowedChanSet: {kênh_thường: true} hoặc null = không giới hạn).
 *  Tính LẠI đầy đủ hoursPay/bonus/hoTro/luongCung/total chỉ từ phần dữ liệu (lines/hoTroItems/luongCungItems)
 *  thuộc các kênh cho phép — dùng chung cho cả lương tính mới (_computePayrollRaw_) lẫn snapshot đã chốt
 *  (readLuongChotSnapshot_), vì cả 2 đều lưu "channel" trên từng dòng/mục chi tiết.
 *  Mục KHÔNG gắn kênh cụ thể (channel rỗng — VD thưởng hỗ trợ chung) luôn được giữ lại ở mọi kênh. */
function filterPayrollList_(list, allowedChanSet) {
  if (!allowedChanSet) return list;
  const out = [];
  list.forEach(function (p) {
    const passChan = function (ch) { return !ch || allowedChanSet[String(ch).toLowerCase()]; };
    const lines = (p.lines || []).filter(function (l) { return passChan(l.channel); });
    const hoTroItems = (p.hoTroItems || []).filter(function (it) { return passChan(it.channel); });
    const luongCungItems = (p.luongCungItems || []).filter(function (it) { return passChan(it.channel); });
    const probationItems = (p.probationItems || []).filter(function (it) { return it.channel && allowedChanSet[String(it.channel).toLowerCase()]; });
    if (!lines.length && !hoTroItems.length && !luongCungItems.length && !probationItems.length) return; // không còn gì trong phạm vi -> ẩn khỏi danh sách
    const hoursPay = Math.round(lines.reduce(function (s, l) { return s + (Number(l.hoursPay) || 0); }, 0));
    const bonus = Math.round(lines.reduce(function (s, l) { return s + (Number(l.thuong) || 0); }, 0));
    const hoTro = Math.round(hoTroItems.reduce(function (s, it) { return s + (Number(it.amount) || 0); }, 0));
    const luongCung = Math.round(luongCungItems.reduce(function (s, it) { return s + (Number(it.amount) || 0); }, 0));
    const probation = Math.round(probationItems.reduce(function (s, it) { return s + (Number(it.amount) || 0); }, 0));
    const copy = {};
    Object.keys(p).forEach(function (k) { copy[k] = p[k]; }); // Object.assign polyfill-free (Apps Script V8 hỗ trợ Object.assign nhưng giữ đơn giản)
    copy.lines = lines; copy.hoTroItems = hoTroItems; copy.luongCungItems = luongCungItems;
    copy.hoursPay = hoursPay; copy.bonus = bonus; copy.hoTro = hoTro; copy.luongCung = luongCung;
    copy.probationItems = probationItems; copy.probation = probation;
    copy.total = hoursPay + bonus + hoTro + luongCung + probation;
    out.push(copy);
  });
  return out;
}

/** Tính lương THÔ từ config hiện tại (bỏ qua snapshot). Dùng nội bộ khi Owner sửa tháng chốt để re-snapshot.
 *  allRowsCache (tuỳ chọn): mảng dòng CaLive đã đọc sẵn (readAllShiftsRaw_()) — truyền vào để tránh đọc lại sheet.
 *  filterChannels (tuỳ chọn): mảng tên kênh muốn lọc riêng (Owner/Lead chủ động chọn ở UI "Lọc kênh"). */
function _computePayrollRaw_(auth, month, allRowsCache, filterChannels) {
  try {
    const cfg = readThuongChung_();

    // Đơn giá giờ theo (tên|kênh) — kênh rỗng "" = mặc định chung cho mọi kênh chưa có riêng
    const hourly = {}; // "name|channel" -> {host, tro}
    const lg = getLuongGioSheet_();
    if (lg.getLastRow() >= 2) lg.getRange(2, 1, lg.getLastRow() - 1, 4).getValues().forEach(function (r) {
      if (!r[0]) return;
      const key = String(r[0]).trim().toLowerCase() + '|' + String(r[1] || '').trim().toLowerCase();
      hourly[key] = { host: Number(r[2]) || 0, tro: Number(r[3]) || 0 };
    });
    // Mốc thưởng theo (tên|kênh|vai trò)
    const mocMap = {};
    const mocSheet = getSimpleSheet_(THUONGMOC_SHEET, THUONGMOC_HEADERS);
    if (mocSheet.getLastRow() >= 2) mocSheet.getRange(2, 1, mocSheet.getLastRow() - 1, 5).getValues().forEach(function (r) {
      if (!r[0]) return;
      const key = String(r[0]).trim().toLowerCase() + '|' + String(r[1]).trim().toLowerCase() + '|' + String(r[2]);
      mocMap[key] = { mocGmvh: Number(r[3]) || 0, mocNmv: Number(r[4]) || 0 };
    });
    // Hoàn huỷ tháng: nếu chưa có -> lấy mặc định trong ThuongChung (mặc định 30%)
    let returnRate = null;
    const hh = getSimpleSheet_(HOANHUY_SHEET, HOANHUY_HEADERS);
    if (hh.getLastRow() >= 2) hh.getRange(2, 1, hh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (asMonthStr_(r[0]) === month) returnRate = (Number(r[1]) || 0) / 100;
    });
    if (returnRate === null) returnRate = (Number(cfg.hoan_huy) || 0) / 100;

    // Điểm theo tên (tháng)
    const diemMap = {};
    const dm = getSimpleSheet_(DIEM_SHEET, DIEM_HEADERS);
    if (dm.getLastRow() >= 2) dm.getRange(2, 1, dm.getLastRow() - 1, 3).getValues().forEach(function (r) {
      if (asMonthStr_(r[0]) === month) diemMap[String(r[1]).trim().toLowerCase()] = Number(r[2]) || 0;
    });

    // ★ MẶC ĐỊNH ĐIỂM: tháng CHƯA chốt và nhân sự CHƯA có điểm nhập riêng cho tháng này
    //   -> lấy điểm phòng live của THÁNG TRƯỚC nếu tháng trước ĐÃ CHỐT lương (đọc từ snapshot, không đổi được).
    //   Điểm đã nhập riêng (kể cả 0) luôn được ưu tiên. diemDefaultSet đánh dấu ai đang dùng điểm mặc định.
    const diemDefaultSet = {};
    if (!checkMonthLocked_(month)) {
      const prevM = prevMonthStr_(month);
      if (prevM && checkMonthLocked_(prevM)) {
        readLuongChotSnapshot_(prevM).forEach(function (sp) {
          const k = String(sp.name).trim().toLowerCase();
          if (diemMap[k] === undefined && sp.diem != null) { diemMap[k] = Number(sp.diem) || 0; diemDefaultSet[k] = true; }
        });
      }
    }

    // Thưởng hỗ trợ theo tên (v2 — nhiều dòng/người/tháng, có kênh + note)
    // Tổng tiền/tên; giữ chi tiết từng dòng để xuất file
    const hoTroMap = {}; // nameLower -> { amount, note (gộp), items:[{channel, amount, note}] }
    const hs = ensureHoTroV2_();
    if (hs.getLastRow() >= 2) {
      hs.getRange(2, 1, hs.getLastRow() - 1, 6).getValues().forEach(function (r) {
        if (asMonthStr_(r[1]) !== month) return;
        const nm = String(r[2] || '').trim();
        if (!nm) return;
        const nl = nm.toLowerCase();
        const ch = String(r[3] || '').trim();
        const amt = Number(r[4]) || 0;
        const note = String(r[5] || '');
        if (!hoTroMap[nl]) hoTroMap[nl] = { amount: 0, note: '', items: [] };
        hoTroMap[nl].amount += amt;
        hoTroMap[nl].items.push({ channel: ch, amount: amt, note: note });
        // Ghi chú tổng hợp cho backward-compat: nối các note không rỗng
        if (note) hoTroMap[nl].note = hoTroMap[nl].note ? (hoTroMap[nl].note + ' · ' + note) : note;
      });
    }

    // Lương cứng theo tháng (v2 — theo tên + kênh)
    // Trả cả list chi tiết để hiển thị + xuất file
    const luongCungMap = {}; // nameLower -> { total, items:[{channel, amount, note}] }
    const lcs = getSimpleSheet_(LUONGCUNG_THANG_SHEET, LUONGCUNG_THANG_HEADERS);
    if (lcs.getLastRow() >= 2) {
      lcs.getRange(2, 1, lcs.getLastRow() - 1, 5).getValues().forEach(function (r) {
        if (asMonthStr_(r[0]) !== month) return;
        const nm = String(r[1] || '').trim();
        if (!nm) return;
        const nl = nm.toLowerCase();
        const ch = String(r[2] || '').trim();
        const amt = Number(r[3]) || 0;
        const note = String(r[4] || '');
        if (!luongCungMap[nl]) luongCungMap[nl] = { total: 0, name: nm, items: [] };
        luongCungMap[nl].total += amt;
        luongCungMap[nl].items.push({ channel: ch, amount: amt, note: note });
      });
    }

    // Gom ca trong tháng theo (tên|kênh|vai trò)
    let leadChanSet = null;
    if (auth.role === 'lead') { leadChanSet = {}; leadChannelsLower_(auth).forEach(function (c) { leadChanSet[c] = true; }); }
    const groups = {}; // key -> {name, channel, role, gmv, hours, holidayExtraHours, holidayDays}
    const holMap = readHolidayMapByMonth_(month); // ngày lễ của tháng: nhân lương GIỜ theo hệ số (thưởng KPI không đổi)
    const rows = allRowsCache || readAllShiftsRaw_();
    {
      rows.forEach(function (r) {
        const end = String(r[6]).trim(); if (!end) return; // chỉ ca đã đóng
        const dateStr = asDateStr_(r[1]); if (dateStr.slice(0, 7) !== month) return;
        const name = String(r[2]).trim(); if (!name) return;
        const channel = String(r[4]).trim();
        if (leadChanSet && !leadChanSet[channel.toLowerCase()]) return; // Lead: chỉ kênh quản lý
        const role = (String(r[3]) === 'Host Live') ? 'Host' : 'Tro';
        const hrs = Number(r[7]) || computeHours_(asTimeStr_(r[5]), end);
        const gmv = Number(r[8]) || 0;
        const key = name.toLowerCase() + '|' + channel.toLowerCase() + '|' + role;
        const g = groups[key] = groups[key] || { name: name, channel: channel, role: role, gmv: 0, hours: 0, holidayExtraHours: 0, holidayDays: {} };
        g.gmv += gmv; g.hours += hrs;
        // Ca bắt đầu vào ngày lễ: giờ làm được tính thêm (hệ số − 1) khi nhân lương giờ
        const hm = holMap[dateStr];
        if (hm && hm !== 1) {
          g.holidayExtraHours += hrs * (hm - 1);
          const hd = g.holidayDays[dateStr] = g.holidayDays[dateStr] || { date: dateStr, mult: hm, hours: 0 };
          hd.hours += hrs;
        }
      });
    }

    // Tính từng nhân viên
    const people = {}; // nameLower -> {name, hoursPay, bonus, lines:[]}
    function kpi(actual, moc, heso, tran) {
      if (!moc || moc <= 0) return 0;
      return Math.min((actual / moc) * heso, tran);
    }
    Object.keys(groups).forEach(function (key) {
      const g = groups[key];
      const nmL = g.name.toLowerCase();
      const p = people[nmL] = people[nmL] || { name: g.name, hoursPay: 0, bonus: 0, lines: [] };
      // Lương giờ: ưu tiên đơn giá theo (tên|kênh cụ thể), fallback về (tên|"") = mặc định chung
      const channelKey = nmL + '|' + g.channel.toLowerCase();
      const defaultKey = nmL + '|';
      const rateObj = hourly[channelKey] || hourly[defaultKey] || { host: 0, tro: 0 };
      const rate = rateObj[g.role === 'Host' ? 'host' : 'tro'];
      const hoursPay = (g.hours + (g.holidayExtraHours || 0)) * rate; // giờ thường + phần tăng thêm của ngày lễ
      p.hoursPay += hoursPay;
      // Thưởng
      const gmvh = g.hours > 0 ? g.gmv / g.hours : 0;
      const nmv = (g.gmv / (cfg.vat || 1.08)) * (1 - returnRate);
      const diem = diemMap[nmL] || 0;
      const mocK = mocMap[key] || { mocGmvh: 0, mocNmv: 0 };
      const k1 = kpi(gmvh, mocK.mocGmvh, cfg.heso_gmvh, cfg.tran_gmvh);
      const k2 = kpi(nmv, mocK.mocNmv, cfg.heso_nmv, cfg.tran_nmv);
      const k3 = kpi(diem, cfg.moc_diem, cfg.heso_diem, cfg.tran_diem);
      const tongHeSo = k1 + k2 + k3;
      const pct = (g.role === 'Host' ? cfg.pct_host : cfg.pct_tro) / 100;
      const thuong = pct * tongHeSo * nmv;
      p.bonus += thuong;
      p.lines.push({
        channel: g.channel, role: g.role, gmv: g.gmv, hours: Math.round(g.hours * 100) / 100,
        gmvh: Math.round(gmvh), nmv: Math.round(nmv), diem: diem,
        k1: Math.round(k1 * 1000) / 1000, k2: Math.round(k2 * 1000) / 1000, k3: Math.round(k3 * 1000) / 1000,
        tongHeSo: Math.round(tongHeSo * 1000) / 1000, hoursPay: Math.round(hoursPay), thuong: Math.round(thuong),
        holidayExtra: Math.round((g.holidayExtraHours || 0) * rate),
        holidayDays: Object.keys(g.holidayDays || {}).sort().map(function (k) {
          const hd = g.holidayDays[k]; return { date: hd.date, mult: hd.mult, hours: Math.round(hd.hours * 100) / 100 };
        }),
        mocGmvh: mocK.mocGmvh, mocNmv: mocK.mocNmv, mocDiem: cfg.moc_diem,
        pctApplied: pct, hesoGmvh: cfg.heso_gmvh, tranGmvh: cfg.tran_gmvh,
        hesoNmv: cfg.heso_nmv, tranNmv: cfg.tran_nmv, hesoDiem: cfg.heso_diem, tranDiem: cfg.tran_diem
      });
    });

    // Lọc theo quyền — đính kèm điểm + thưởng hỗ trợ + lương cứng ở cấp nhân viên
    let list = Object.keys(people).map(function (k) {
      const p = people[k];
      const hoTro = hoTroMap[k] || { amount: 0, note: '', items: [] };
      const lc = luongCungMap[k] || { total: 0, items: [] };
      return {
        name: p.name,
        hoursPay: Math.round(p.hoursPay),
        bonus: Math.round(p.bonus),
        hoTro: Math.round(hoTro.amount),
        hoTroNote: hoTro.note,
        hoTroItems: hoTro.items,       // chi tiết từng dòng thưởng hỗ trợ [{channel, amount, note}]
        luongCung: Math.round(lc.total),
        luongCungItems: lc.items,      // chi tiết lương cứng theo kênh
        total: Math.round(p.hoursPay + p.bonus + hoTro.amount + lc.total),
        diem: diemMap[k] || 0, diemDefault: !!diemDefaultSet[k],
        lines: p.lines,
        isFixed: false                  // NV có ca live (không phải NV lương cứng thuần)
      };
    });
    // Thêm cả nhân viên chỉ có thưởng hỗ trợ / lương cứng (không có ca live) vào list
    const extraNames = {};
    Object.keys(hoTroMap).forEach(function (k) { if (!people[k]) extraNames[k] = true; });
    Object.keys(luongCungMap).forEach(function (k) { if (!people[k]) extraNames[k] = true; });
    Object.keys(extraNames).forEach(function (k) {
      const ht = hoTroMap[k] || { amount: 0, note: '', items: [] };
      const lc = luongCungMap[k] || { total: 0, items: [], name: '' };
      if (!ht.amount && !ht.note && !lc.total) return; // rỗng
      // Ưu tiên tên chuẩn từ luongCungMap (đã lưu name gốc)
      const nameGuess = lc.name || (k.charAt(0).toUpperCase() + k.slice(1));
      list.push({
        name: nameGuess,
        hoursPay: 0, bonus: 0,
        hoTro: Math.round(ht.amount), hoTroNote: ht.note, hoTroItems: ht.items,
        luongCung: Math.round(lc.total), luongCungItems: lc.items,
        total: Math.round(ht.amount + lc.total),
        diem: diemMap[k] || 0, diemDefault: !!diemDefaultSet[k],
        lines: [],
        isFixed: (lc.total > 0 && !ht.amount) // đánh dấu để UI phân biệt
      });
    });
    // Sửa lại tên nhân viên chỉ có thưởng hỗ trợ — đọc thẳng từ sheet HoTro v2 để lấy tên chuẩn
    if (hs.getLastRow() >= 2) {
      const htRows = hs.getRange(2, 1, hs.getLastRow() - 1, 6).getValues(); // v2: [ID, Tháng, Tên, Kênh, Số tiền, Ghi chú]
      const nameByLower = {};
      htRows.forEach(function (r) {
        if (asMonthStr_(r[1]) === month) {
          const nm = String(r[2] || '').trim();
          if (nm) nameByLower[nm.toLowerCase()] = nm;
        }
      });
      list.forEach(function (x) {
        if (nameByLower[x.name.toLowerCase()]) x.name = nameByLower[x.name.toLowerCase()];
      });
    }
    if (auth.role === 'lead') { const set = leadStaffSet_(auth); list = list.filter(function (x) { return set[x.name.toLowerCase()]; }); }
    else if (auth.role === 'staff') { const nm = String(auth.staff).toLowerCase(); list = list.filter(function (x) { return x.name.toLowerCase() === nm; }); }

    // ★ LỌC THEO KÊNH: bắt buộc với Lead (chỉ kênh mình quản lý) + tuỳ chọn lọc thêm (Owner/Lead chọn "Lọc kênh").
    // Áp dụng cho TOÀN BỘ thành phần thu nhập (lương giờ, thưởng KPI, thưởng hỗ trợ, lương cứng) — trước đây
    // chỉ lương giờ/KPI được lọc theo kênh Lead quản lý, còn thưởng hỗ trợ/lương cứng bị lộ số liệu kênh khác.
    // ★ NHÂN SỰ THỬ VIỆC: nhập tay theo tháng, cộng thẳng vào lương (không chấm công). Gộp vào người đã có cùng tên, hoặc tạo dòng riêng.
    const probRows = readProbationByMonth_(month);
    if (probRows.length) {
      const byLower = {};
      list.forEach(function (x) { byLower[x.name.toLowerCase()] = x; });
      const probBy = {};
      probRows.forEach(function (it) {
        const k = it.name.toLowerCase();
        if (auth.role === 'staff' && k !== String(auth.staff).toLowerCase()) return;
        (probBy[k] = probBy[k] || { name: it.name, items: [] }).items.push({ channel: it.channel, amount: it.amount, note: it.note });
      });
      Object.keys(probBy).forEach(function (k) {
        const items = probBy[k].items;
        const sum = Math.round(items.reduce(function (s, it) { return s + (Number(it.amount) || 0); }, 0));
        let person = byLower[k];
        if (!person) {
          person = { name: probBy[k].name, hoursPay: 0, bonus: 0, hoTro: 0, hoTroNote: '', hoTroItems: [],
            luongCung: 0, luongCungItems: [], total: 0, diem: 0, lines: [], isFixed: false, isProbation: true };
          list.push(person);
          byLower[k] = person;
        }
        person.probationItems = items;
        person.probation = sum;
        person.total = Math.round((person.total || 0) + sum);
      });
    }

    list = filterPayrollList_(list, buildAllowedChanSet_(auth, filterChannels));

    list.sort(function (a, b) { return b.total - a.total; });

    return { success: true, month: month, returnRate: Math.round(returnRate * 10000) / 100, data: list };
  } catch (e) { return { success: false, error: 'Lỗi tính lương: ' + e.message }; }
}

/* ----------------------- TÀI KHOẢN & PHÂN QUYỀN ----------------------- */

const USERS_SHEET = 'Users';
const USERS_HEADERS = ['Username', 'Hash', 'Salt', 'Vai trò', 'Kênh', 'Nhân sự', 'Thời gian tạo', 'Đã thiết lập', 'Kênh quản lý'];

/** Lấy (hoặc tạo) sheet Users; seed sẵn 1 tài khoản owner mặc định ở lần đầu. */
function getUsersSheet_() {
  const ss = ss_();
  let sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET);
    sheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, USERS_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a1d26').setFontColor('#ffffff');
    sheet.getRange('A:I').setNumberFormat('@');
    sheet.setColumnWidth(1, 160);
    const salt = Utilities.getUuid();
    sheet.appendRow(['hunglx0901', hashPw_('Hunghung1@3', salt), salt, 'owner', '', '', nowString_(), 'x', '']);
  } else if (sheet.getLastColumn() < USERS_HEADERS.length) {
    // Migration: bổ sung cột mới cho sheet cũ (chỉ ghi hàng tiêu đề, không phá dữ liệu)
    sheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS])
      .setFontWeight('bold').setBackground('#1a1d26').setFontColor('#ffffff');
  }
  return sheet;
}

function hashPw_(pw, salt) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(salt) + '|' + String(pw), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function authSecret_() {
  if (MEMO_SECRET_) return MEMO_SECRET_;
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('AUTH_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('AUTH_SECRET', s); }
  MEMO_SECRET_ = s;
  return s;
}

function makeToken_(id) {
  const exp = Date.now() + 12 * 3600 * 1000; // 12 giờ
  const chans = id.channels ? id.channels : splitChannels_(id.channel);
  const managed = id.managed ? id.managed : [];
  const json = JSON.stringify(
    { u: id.username, r: id.role, c: chans.join('|'), s: id.staff || '', m: managed.join('|'), e: exp });
  // Mã hoá theo UTF-8 để giữ nguyên dấu tiếng Việt (â, ê...) trong token
  const payload = Utilities.base64EncodeWebSafe(json, Utilities.Charset.UTF_8);
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(payload, authSecret_()));
  return payload + '.' + sig;
}

function verifyToken_(token) {
  try {
    if (!token) return null;
    const parts = String(token).split('.');
    if (parts.length !== 2) return null;
    const sig = Utilities.base64EncodeWebSafe(
      Utilities.computeHmacSha256Signature(parts[0], authSecret_()));
    if (sig !== parts[1]) return null;
    const data = JSON.parse(Utilities.newBlob(
      Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
    if (Date.now() > Number(data.e)) return null;
    const chans = splitChannels_(data.c);
    return { username: data.u, role: data.r, channel: chans[0] || '', channels: chans, staff: data.s, managed: splitChannels_(data.m) };
  } catch (e) { return null; }
}

/* --- Helper đa kênh cho Lead --- */
function splitChannels_(s) {
  return String(s || '').split('|').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
}
function leadChannelsLower_(auth) {
  return (auth.channels || []).map(function (x) { return String(x).toLowerCase(); });
}
function leadHasChannel_(auth, channel) {
  return leadChannelsLower_(auth).indexOf(String(channel || '').toLowerCase()) >= 0;
}
/** Union tên nhân sự (lowercase) có ca trong CÁC kênh của Lead. */
function leadStaffSet_(auth) {
  const set = {};
  const leadChans = {};
  (auth.channels || []).forEach(function (ch) { leadChans[String(ch).toLowerCase()] = true; });
  // 1) Tên nhân sự từ ca live (CaLive) thuộc kênh của lead
  (auth.channels || []).forEach(function (ch) {
    const s = channelStaffNames_(ch);
    for (const k in s) set[k] = true;
  });
  // 2) Tên nhân sự từ tài khoản (Users) có kênh trùng với lead
  //    -> nhân viên mới được lead tạo (chưa có ca live) vẫn nằm trong phạm vi
  try {
    const usersSheet = getUsersSheet_();
    const last = usersSheet.getLastRow();
    if (last >= 2) {
      const vals = usersSheet.getRange(2, 1, last - 1, USERS_HEADERS.length).getValues();
      vals.forEach(function (r) {
        if (String(r[3]) !== 'staff') return;
        const staffName = String(r[5] || '').trim().toLowerCase();
        if (!staffName) return;
        const userChans = splitChannels_(String(r[4] || ''));
        for (let i = 0; i < userChans.length; i++) {
          if (leadChans[String(userChans[i]).toLowerCase()]) {
            set[staffName] = true;
            return;
          }
        }
      });
    }
  } catch (e) { /* ignore */ }
  return set;
}

function authErr_() { return { success: false, code: 'AUTH', error: 'Phiên đăng nhập đã hết. Đăng nhập lại nhé.' }; }
function denyErr_() { return { success: false, error: 'Bạn không có quyền thực hiện thao tác này.' }; }

function usersRow_(sheet, username) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const target = String(username || '').trim().toLowerCase();
  const vals = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

/** ĐĂNG NHẬP — trả token nếu đúng. */
function login(username, password) {
  try {
    const sheet = getUsersSheet_();
    const row = usersRow_(sheet, username);
    if (row < 0) return { success: false, error: 'Tài khoản không tồn tại.' };
    const r = sheet.getRange(row, 1, 1, USERS_HEADERS.length).getValues()[0];
    if (hashPw_(password, String(r[2])) !== String(r[1]))
      return { success: false, error: 'Sai mật khẩu.' };
    const staffName = String(r[5] || '');
    let chans = splitChannels_(String(r[4] || ''));
    // Nhân viên: tự thêm kênh từ ca đã có (check-in cũ hoặc import trước đây)
    if (r[3] === 'staff' && staffName) {
      const synced = syncStaffChannels_(String(r[0]), staffName);
      if (synced) chans = synced;
    }
    // Popup chọn kênh chỉ hiện 1 lần (lần đăng nhập đầu) cho Nhân viên
    const managed = splitChannels_(String(r[8] || ''));
    const firstLogin = (r[3] === 'staff') && !String(r[7] || '').trim();
    if (firstLogin) sheet.getRange(row, 8).setValue('x'); // đánh dấu để không hiện lại
    const id = { username: String(r[0]), role: String(r[3]), channel: chans[0] || '', channels: chans, staff: staffName, managed: managed };
    return { success: true, token: makeToken_(id), role: id.role, channel: id.channel, channels: id.channels, staff: id.staff, username: id.username, firstLogin: firstLogin, managed: managed };
  } catch (e) {
    return { success: false, error: 'Lỗi đăng nhập: ' + e.message };
  }
}

/** Khôi phục phiên từ token đã lưu. */
function whoami(token) {
  const id = verifyToken_(token);
  if (!id) return { success: false, code: 'AUTH' };
  return { success: true, role: id.role, channel: id.channel, channels: id.channels, staff: id.staff, username: id.username, managed: id.managed };
}

/** Đổi mật khẩu của chính mình. */
function changeMyPassword(token, oldPw, newPw) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id = verifyToken_(token);
    if (!id) return authErr_();
    if (!newPw || String(newPw).length < 4) return { success: false, error: 'Mật khẩu mới tối thiểu 4 ký tự.' };
    const sheet = getUsersSheet_();
    const row = usersRow_(sheet, id.username);
    if (row < 0) return authErr_();
    const r = sheet.getRange(row, 1, 1, USERS_HEADERS.length).getValues()[0];
    if (hashPw_(oldPw, String(r[2])) !== String(r[1])) return { success: false, error: 'Mật khẩu hiện tại không đúng.' };
    const salt = Utilities.getUuid();
    sheet.getRange(row, 2, 1, 2).setValues([[hashPw_(newPw, salt), salt]]);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi đổi mật khẩu: ' + e.message };
  } finally { lock.releaseLock(); }
}

/* --- Quản lý tài khoản (chỉ Owner) --- */

function listUsers(token) {
  const id = verifyToken_(token);
  if (!id) return authErr_();
  if (id.role !== 'owner' && id.role !== 'lead') return denyErr_();
  try {
    const sheet = getUsersSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, data: [] };
    const vals = sheet.getRange(2, 1, last - 1, USERS_HEADERS.length).getValues();
    let data = vals.filter(function (r) { return r[0]; }).map(function (r) {
      return { username: String(r[0]), role: String(r[3]), channel: String(r[4] || ''), staff: String(r[5] || ''), managed: String(r[8] || '') };
    });
    if (id.role === 'lead') {
      const set = leadStaffSet_(id);
      data = data.filter(function (u) { return u.role === 'staff' && set[String(u.staff).toLowerCase()]; });
    }
    return { success: true, data: data };
  } catch (e) { return { success: false, error: 'Lỗi đọc tài khoản: ' + e.message }; }
}

function createUser(token, username, password, role, channel, staff, managed) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id = verifyToken_(token);
    if (!id) return authErr_();
    if (id.role !== 'owner' && id.role !== 'lead') return denyErr_();
    username = String(username || '').trim();
    role = String(role || '').trim();
    // Lead chỉ được tạo tài khoản Nhân viên (được phép gán cho bất kỳ nhân sự nào, kể cả nhân sự mới)
    if (id.role === 'lead') {
      if (role !== 'staff') return denyErr_();
    }
    if (!username) return { success: false, error: 'Thiếu tên đăng nhập.' };
    if (!password || String(password).length < 4) return { success: false, error: 'Mật khẩu tối thiểu 4 ký tự.' };
    if (['owner', 'admin', 'lead', 'staff'].indexOf(role) < 0) return { success: false, error: 'Vai trò không hợp lệ.' };
    const chans = splitChannels_(channel);
    if (role === 'lead' && chans.length === 0) return { success: false, error: 'Lead cần gán ít nhất 1 kênh.' };
    if (role === 'staff' && !String(staff || '').trim()) return { success: false, error: 'Nhân viên cần gán 1 nhân sự.' };
    const managedChans = (role === 'staff') ? splitChannels_(managed) : [];
    const sheet = getUsersSheet_();
    if (usersRow_(sheet, username) > 0) return { success: false, error: 'Tên đăng nhập đã tồn tại.' };
    const salt = Utilities.getUuid();
    if (role === 'staff' && String(staff || '').trim()) ensureStaff_(String(staff).trim()); // nhân sự mới -> tự thêm vào NhanSu
    sheet.appendRow([username, hashPw_(password, salt), salt, role,
      (role === 'lead' || role === 'staff') ? chans.join('|') : '',
      (role === 'staff') ? String(staff).trim() : '', nowString_(), '', managedChans.join('|')]);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi tạo tài khoản: ' + e.message };
  } finally { lock.releaseLock(); }
}

function updateUser(token, username, role, channel, staff, managed) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id = verifyToken_(token);
    if (!id) return authErr_();
    if (id.role !== 'owner' && id.role !== 'lead') return denyErr_();
    const sheet = getUsersSheet_();
    const row = usersRow_(sheet, username);
    if (row < 0) return { success: false, error: 'Không tìm thấy tài khoản.' };
    const curRole = String(sheet.getRange(row, 4).getValue());
    const curStaff = String(sheet.getRange(row, 6).getValue());

    // LEAD: chỉ được sửa kênh của Nhân viên thuộc phạm vi mình quản lý (không đổi vai trò/nhân sự)
    if (id.role === 'lead') {
      if (curRole !== 'staff') return denyErr_();
      const set = leadStaffSet_(id);
      if (!set[curStaff.trim().toLowerCase()]) return denyErr_();
      sheet.getRange(row, 5).setValue(splitChannels_(channel).join('|'));
      return { success: true };
    }

    // OWNER: cập nhật đầy đủ
    role = String(role || '').trim();
    if (['owner', 'admin', 'lead', 'staff'].indexOf(role) < 0) return { success: false, error: 'Vai trò không hợp lệ.' };
    const chans = splitChannels_(channel);
    if (role === 'lead' && chans.length === 0) return { success: false, error: 'Lead cần gán ít nhất 1 kênh.' };
    if (role === 'staff' && !String(staff || '').trim()) return { success: false, error: 'Nhân viên cần gán 1 nhân sự.' };
    sheet.getRange(row, 4, 1, 3).setValues([[role,
      (role === 'lead' || role === 'staff') ? chans.join('|') : '',
      (role === 'staff') ? String(staff).trim() : '']]);
    sheet.getRange(row, 9).setValue((role === 'staff') ? splitChannels_(managed).join('|') : '');
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi cập nhật tài khoản: ' + e.message };
  } finally { lock.releaseLock(); }
}

function resetPassword(token, username, newPw) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id = verifyToken_(token);
    if (!id) return authErr_();
    if (id.role !== 'owner') return denyErr_();
    if (!newPw || String(newPw).length < 4) return { success: false, error: 'Mật khẩu tối thiểu 4 ký tự.' };
    const sheet = getUsersSheet_();
    const row = usersRow_(sheet, username);
    if (row < 0) return { success: false, error: 'Không tìm thấy tài khoản.' };
    const salt = Utilities.getUuid();
    sheet.getRange(row, 2, 1, 2).setValues([[hashPw_(newPw, salt), salt]]);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi đặt lại mật khẩu: ' + e.message };
  } finally { lock.releaseLock(); }
}

function deleteUser(token, username) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id = verifyToken_(token);
    if (!id) return authErr_();
    if (id.role !== 'owner') return denyErr_();
    username = String(username || '').trim();
    if (username.toLowerCase() === String(id.username).toLowerCase())
      return { success: false, error: 'Không thể xoá tài khoản đang đăng nhập.' };
    const sheet = getUsersSheet_();
    const last = sheet.getLastRow();
    const vals = last >= 2 ? sheet.getRange(2, 1, last - 1, 4).getValues() : [];
    const owners = vals.filter(function (r) { return String(r[3]) === 'owner'; });
    const row = usersRow_(sheet, username);
    if (row < 0) return { success: false, error: 'Không tìm thấy tài khoản.' };
    if (String(sheet.getRange(row, 4).getValue()) === 'owner' && owners.length <= 1)
      return { success: false, error: 'Phải còn ít nhất 1 Owner.' };
    sheet.deleteRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi xoá tài khoản: ' + e.message };
  } finally { lock.releaseLock(); }
}

/** Tập tên nhân sự (lowercase) có ca trong 1 kênh — dùng cho Lead. */
function channelStaffNames_(channel) {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  const set = {};
  if (last < 2) return set;
  const vals = sheet.getRange(2, 3, last - 1, 3).getValues(); // C: Tên, D: Vị trí, E: Kênh
  const ch = String(channel || '').toLowerCase();
  vals.forEach(function (r) {
    if (String(r[2]).toLowerCase() === ch) set[String(r[0]).trim().toLowerCase()] = true;
  });
  return set;
}

/* ----------------------- CRUD ----------------------- */

/** READ — lấy ca live (lọc theo quyền). */
function getShifts(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, data: [] };

    const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
    let data = values
      .filter(function (r) { return r[0]; })
      .map(function (r) {
        var start = asTimeStr_(r[5]);
        var end = asTimeStr_(r[6]);
        var closed = !!String(r[6]).trim();
        return {
          id: String(r[0]),
          date: asDateStr_(r[1]),
          name: String(r[2]),
          position: String(r[3]),
          channel: String(r[4]),
          start: start,
          end: end,
          hours: closed ? computeHours_(start, end) : (Number(r[7]) || 0),
          gmv: Number(r[8]) || 0,
          note: String(r[9] || ''),
          createdAt: String(r[10] || ''),
          gmvStart: Number(r[11]) || 0,
          gmvEnd: Number(r[12]) || 0,
          open: !closed,
          hasStartPhoto: !!String(r[13] || '').trim(),
          hasEndPhoto: !!String(r[14] || '').trim()
        };
      });
    if (auth.role === 'lead') {
      const set = {};
      leadChannelsLower_(auth).forEach(function (c) { set[c] = true; });
      data = data.filter(function (x) { return set[x.channel.toLowerCase()]; });
    } else if (auth.role === 'staff') {
      // Nhân viên: ca của chính mình + (nếu có quyền Lead Live) hiệu suất các KÊNH QUẢN LÝ được giao
      const nm = String(auth.staff).toLowerCase();
      const mset = {}; (auth.managed || []).forEach(function (c) { mset[String(c).toLowerCase()] = true; });
      data = data.filter(function (x) { return x.name.toLowerCase() === nm || mset[x.channel.toLowerCase()]; });
    }
    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: 'Lỗi đọc dữ liệu: ' + e.message };
  }
}

/** CREATE — thêm 1 ca live. */
function addShift(token, rec) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role === 'admin') return denyErr_();
  rec = rec || {};
  if (auth.role === 'lead' && !leadHasChannel_(auth, rec.channel)) return denyErr_(); // chỉ kênh của Lead
  if (auth.role === 'staff') rec.name = auth.staff;          // ép tên theo Nhân viên
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const err = validate_(rec);
    if (err) return { success: false, error: err };

    const sheet = getSheet_();
    const id = Utilities.getUuid();
    const hours = computeHours_(rec.start, rec.end);
    const gmvStart = Number(rec.gmvStart) || 0;
    const gmvEnd = Number(rec.gmvEnd) || 0;
    const gmv = gmvEnd - gmvStart;

    sheet.appendRow([
      id, rec.date, String(rec.name).trim(), rec.position, rec.channel,
      rec.start, rec.end, hours, gmv,
      String(rec.note || '').trim(), nowString_(), gmvStart, gmvEnd, '', ''
    ]);
    ensureStaff_(rec.name);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: 'Lỗi khi thêm: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Phân tích văn bản dán vào thành các dòng ca. Dùng chung cho preview & import. */
function parseImportRows_(text, mode) {
  mode = (mode === 'hours') ? 'hours' : 'times';
  const lines = String(text || '').split(/\r?\n/);
  const rows = [];
  const errors = [];
  let headerSkipped = false;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const cols = line.indexOf('\t') >= 0 ? line.split('\t') : line.split(',');
    if (!headerSkipped && (/ng[aà]y/i.test(cols[0] || '') || /t[eê]n/i.test(cols[1] || ''))) {
      headerSkipped = true; continue;
    }
    const date = parseImportDate_(String(cols[0] || '').trim());
    const name = String(cols[1] || '').trim();
    const channel = String(cols[2] || '').trim();
    const position = String(cols[3] || '').trim();
    if (!date || !name) { errors.push('Dòng ' + (i + 1) + ': thiếu Ngày hoặc Tên.'); continue; }
    let start, end, hours, gmv;
    if (mode === 'times') {
      start = parseImportTime_(cols[4]);
      end = parseImportTime_(cols[5]);
      gmv = parseImportNumber_(cols[6]);
      if (!start || !end) { errors.push('Dòng ' + (i + 1) + ': giờ bắt đầu/kết thúc không hợp lệ.'); continue; }
      hours = computeHours_(start, end);
    } else {
      hours = parseImportHours_(cols[4]);
      gmv = parseImportNumber_(cols[5]);
      start = '00:00'; end = hoursToHHMM_(hours);
    }
    rows.push([
      Utilities.getUuid(), date, name, position, channel,
      start, end, hours, gmv, 'Import', nowString_(), 0, gmv, '', ''
    ]);
  }
  return { rows: rows, errors: errors };
}

/** PREVIEW — đọc dữ liệu, trả về tóm tắt để Owner xác nhận (KHÔNG ghi gì). */
function previewImport(token, text, mode) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  text = String(text || '').trim();
  if (!text) return { success: false, error: 'Chưa có dữ liệu.' };
  try {
    const parsed = parseImportRows_(text, mode);
    const rows = parsed.rows;
    if (!rows.length) return { success: false, error: 'Không có dòng hợp lệ. ' + parsed.errors.slice(0, 3).join(' ') };
    // Danh sách nhân sự & kênh đang có
    const existStaff = {}; (getStaff(token).data || []).forEach(function (s) { existStaff[String(s).toLowerCase()] = true; });
    const existChan = {}; (getChannels(token).data || []).forEach(function (c) { existChan[String(c).toLowerCase()] = true; });
    const byName = {}; const chanSet = {}; let minD = null, maxD = null;
    rows.forEach(function (r) {
      const nm = r[2], ch = r[4], hrs = Number(r[7]) || 0, gmv = Number(r[8]) || 0, d = r[1];
      const o = byName[nm] = byName[nm] || { count: 0, hours: 0, gmv: 0, chans: {} };
      o.count++; o.hours += hrs; o.gmv += gmv; if (ch) o.chans[ch] = true;
      if (ch) chanSet[ch] = true;
      if (!minD || d < minD) minD = d; if (!maxD || d > maxD) maxD = d;
    });
    const employees = Object.keys(byName).sort().map(function (nm) {
      return {
        name: nm, count: byName[nm].count,
        hours: Math.round(byName[nm].hours * 100) / 100, gmv: byName[nm].gmv,
        channels: Object.keys(byName[nm].chans).sort(),
        isNew: !existStaff[nm.toLowerCase()]
      };
    });
    const newChannels = Object.keys(chanSet).filter(function (c) { return !existChan[c.toLowerCase()]; }).sort();
    return {
      success: true, valid: rows.length, skipped: parsed.errors.length,
      errors: parsed.errors.slice(0, 5),
      employees: employees, newChannels: newChannels,
      dateFrom: minD, dateTo: maxD
    };
  } catch (e) {
    return { success: false, error: 'Lỗi đọc dữ liệu: ' + e.message };
  }
}

/** IMPORT — nhập nhiều ca lịch sử cùng lúc (chỉ Owner). */
function importShifts(token, text, mode) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  text = String(text || '').trim();
  if (!text) return { success: false, error: 'Chưa có dữ liệu để nhập.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = getSheet_();
    const parsed = parseImportRows_(text, mode);
    const rows = parsed.rows, errors = parsed.errors;
    if (!rows.length) return { success: false, error: 'Không có dòng hợp lệ. ' + errors.slice(0, 3).join(' ') };
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
    sheet.getRange(startRow, 2, rows.length, 1).setNumberFormat('@'); // Ngày dạng text
    sheet.getRange(startRow, 6, rows.length, 2).setNumberFormat('@'); // Bắt đầu/Kết thúc dạng text
    // Tự thêm nhân sự & kênh mới xuất hiện trong dữ liệu
    const names = {}, chans = {};
    rows.forEach(function (r) { names[r[2]] = true; if (r[4]) chans[r[4]] = true; });
    Object.keys(names).forEach(function (n) { ensureStaff_(n); });
    Object.keys(chans).forEach(function (c) { ensureChannel_(c); });
    return { success: true, imported: rows.length, skipped: errors.length, errors: errors.slice(0, 5) };
  } catch (e) {
    return { success: false, error: 'Lỗi import: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function parseImportTime_(v) {
  let s = String(v == null ? '' : v).trim().toLowerCase().replace(/h/g, ':').replace(/\./g, ':');
  if (/:$/.test(s)) s += '00';
  let m = s.match(/^(\d{1,2}):(\d{1,2})$/);
  if (m) {
    let hh = parseInt(m[1], 10) % 24, mm = Math.min(59, parseInt(m[2], 10)); // 24:00 -> 00:00
    return ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2);
  }
  m = s.match(/^(\d{1,2})$/);
  if (m) { const hh = parseInt(m[1], 10) % 24; return ('0' + hh).slice(-2) + ':00'; }
  return '';
}

function parseImportDate_(s) {
  s = String(s || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);       // yyyy-mm-dd
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);           // dd/mm/yyyy
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2})$/);           // dd/mm/yy
  if (m) return '20' + m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s;
}
function parseImportHours_(v) {
  let s = String(v == null ? '' : v).replace(/[^\d.,]/g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}
function parseImportNumber_(v) {
  let s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  if (/e/i.test(s)) { const n = parseFloat(s); return isNaN(n) ? 0 : Math.round(n); } // 7.16E7
  s = s.replace(/\s/g, '');
  const seps = (s.match(/[.,]/g) || []).length;
  if (seps >= 2) return parseInt(s.replace(/[^\d]/g, ''), 10) || 0; // 71.600.000 / 71,600,000
  if (seps === 1) {
    const m = s.match(/[.,](\d+)$/);
    if (m && m[1].length === 3) return parseInt(s.replace(/[^\d]/g, ''), 10) || 0; // 71.600 -> 71600
    return Math.round(parseFloat(s.replace(',', '.'))) || 0;                      // 10433333.33 -> 10433333
  }
  return parseInt(s.replace(/[^\d]/g, ''), 10) || 0;
}
function hoursToHHMM_(h) {
  const totalMin = Math.round((Number(h) || 0) * 60);
  const hh = Math.floor(totalMin / 60), mm = totalMin % 60;
  return ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2);
}

/** UPDATE — sửa 1 ca live theo id. */
function updateShift(token, id, rec) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role === 'admin' || auth.role === 'staff') return denyErr_();
  if (auth.role === 'lead' && !leadHasChannel_(auth, rec.channel)) return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!id) return { success: false, error: 'Thiếu ID bản ghi.' };
    const err = validate_(rec);
    if (err) return { success: false, error: err };

    const sheet = getSheet_();
    const row = findRow_(sheet, id);
    if (row < 0) return { success: false, error: 'Không tìm thấy bản ghi để sửa.' };
    if (auth.role === 'lead' && !leadHasChannel_(auth, String(sheet.getRange(row, 5).getValue())))
      return denyErr_();

    const hours = computeHours_(rec.start, rec.end);
    const createdAt = sheet.getRange(row, 11).getValue() || nowString_();
    const gmvStart = Number(rec.gmvStart) || 0;
    const gmvEnd = Number(rec.gmvEnd) || 0;
    const gmv = gmvEnd - gmvStart;
    const photoStart = sheet.getRange(row, 14).getValue() || '';  // giữ ảnh cũ
    const photoEnd = sheet.getRange(row, 15).getValue() || '';

    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
      id, rec.date, String(rec.name).trim(), rec.position, rec.channel,
      rec.start, rec.end, hours, gmv,
      String(rec.note || '').trim(), createdAt, gmvStart, gmvEnd, photoStart, photoEnd
    ]]);
    ensureStaff_(rec.name);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi khi sửa: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** DELETE — xoá 1 ca live theo id. */
function deleteShift(token, id) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role === 'admin' || auth.role === 'staff') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!id) return { success: false, error: 'Thiếu ID bản ghi.' };

    const sheet = getSheet_();
    const row = findRow_(sheet, id);
    if (row < 0) return { success: false, error: 'Không tìm thấy bản ghi để xoá.' };
    if (auth.role === 'lead' && !leadHasChannel_(auth, String(sheet.getRange(row, 5).getValue())))
      return denyErr_();

    deletePhoto_(String(sheet.getRange(row, 14).getValue() || ''));
    deletePhoto_(String(sheet.getRange(row, 15).getValue() || ''));
    sheet.deleteRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi khi xoá: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteShiftsBatch(token, ids) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role === 'admin' || auth.role === 'staff') return denyErr_();
  ids = Array.isArray(ids) ? ids : [];
  if (!ids.length) return { success: false, error: 'Không có ca nào để xoá.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: false, error: 'Không có dữ liệu.' };

    const idCol = sheet.getRange(2, 1, last - 1, 1).getValues();
    const rowById = {};
    for (let i = 0; i < idCol.length; i++) {
      if (idCol[i][0]) rowById[String(idCol[i][0])] = i + 2;
    }

    const idSet = {};
    ids.forEach(function (id) { if (id) idSet[String(id)] = true; });
    const toDelete = [];
    const errors = [];
    let skipped = 0;
    Object.keys(idSet).forEach(function (id) {
      const row = rowById[id];
      if (!row) { skipped++; errors.push('Không tìm thấy ca: ' + id.slice(0, 8)); return; }
      const rowData = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      const channel = String(rowData[4]);
      if (auth.role === 'lead' && !leadHasChannel_(auth, channel)) { skipped++; errors.push('Không có quyền: ' + channel); return; }
      toDelete.push({ id: id, row: row, photoStart: String(rowData[13] || ''), photoEnd: String(rowData[14] || '') });
    });

    if (!toDelete.length) return { success: false, error: 'Không có ca hợp lệ để xoá.', errors: errors };

    toDelete.forEach(function (x) {
      try { deletePhoto_(x.photoStart); } catch (e) { }
      try { deletePhoto_(x.photoEnd); } catch (e) { }
    });

    toDelete.sort(function (a, b) { return b.row - a.row; });
    toDelete.forEach(function (x) { sheet.deleteRow(x.row); });

    return { success: true, deleted: toDelete.length, skipped: skipped, errors: errors.slice(0, 5) };
  } catch (e) {
    return { success: false, error: 'Lỗi khi xoá hàng loạt: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/* ----------------------- ẢNH CHẤM CÔNG (Drive) ----------------------- */

function getPhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('PHOTO_FOLDER_ID');
  let folder = null;
  if (id) { try { folder = DriveApp.getFolderById(id); } catch (e) { folder = null; } }
  if (!folder) {
    folder = DriveApp.createFolder('EAZ - Ảnh chấm công');
    props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  }
  return folder;
}

/** Lưu ảnh base64 (jpeg) vào Drive, trả file ID. */
function savePhoto_(b64, nameHint) {
  if (!b64) return '';
  const clean = String(b64).indexOf(',') >= 0 ? String(b64).split(',')[1] : String(b64);
  const bytes = Utilities.base64Decode(clean);
  const blob = Utilities.newBlob(bytes, 'image/jpeg', (nameHint || 'cc') + '_' + Date.now() + '.jpg');
  return getPhotoFolder_().createFile(blob).getId();
}

function deletePhoto_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
}

/** Đọc 1 ảnh của ca (kiểm tra quyền theo ca). which = 'start' | 'end'. */
function getShiftPhoto(token, shiftId, which) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const sheet = getSheet_();
    const row = findRow_(sheet, shiftId);
    if (row < 0) return { success: false, error: 'Không tìm thấy ca.' };
    const name = String(sheet.getRange(row, 3).getValue());
    const channel = String(sheet.getRange(row, 5).getValue());
    if (auth.role === 'lead' && !leadHasChannel_(auth, channel)) return denyErr_();
    if (auth.role === 'staff') {
      const isMine = name.toLowerCase() === String(auth.staff).toLowerCase();
      const inManaged = (auth.managed || []).some(function (c) {
        return String(c).toLowerCase() === channel.toLowerCase();
      });
      if (!isMine && !inManaged) return denyErr_();
    }
    const fileId = String(sheet.getRange(row, which === 'end' ? 15 : 14).getValue() || '');
    if (!fileId) return { success: false, error: 'Chưa có ảnh.' };
    let blob;
    try { blob = DriveApp.getFileById(fileId).getBlob(); }
    catch (e) { return { success: false, error: 'Ảnh đã hết hạn lưu trữ (quá ' + PHOTO_RETENTION_DAYS + ' ngày).' }; }
    const data = 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes());
    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: 'Lỗi đọc ảnh: ' + e.message };
  }
}

/* ===================== YÊU CẦU SỬA CA (DUYỆT 2 CẤP) ===================== */

function getRequestsSheet_() {
  const ss = ss_();
  let sheet = ss.getSheetByName(REQ_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REQ_SHEET);
    sheet.getRange(1, 1, 1, REQ_HEADERS.length).setValues([REQ_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, REQ_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a1d26').setFontColor('#ffffff');
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('B:B').setNumberFormat('@');
    sheet.getRange('F:F').setNumberFormat('@');
    sheet.getRange('H:I').setNumberFormat('@');
    sheet.getRange('J:M').setNumberFormat('#,##0');
    sheet.setColumnWidth(1, 220);
    return sheet;
  }
  if (sheet.getMaxColumns() < REQ_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), REQ_HEADERS.length - sheet.getMaxColumns());
    // Đồng bộ tiêu đề CHỈ khi vừa thêm cột mới (migrate) — không ghi lại trên mọi lần gọi.
    sheet.getRange(1, 1, 1, REQ_HEADERS.length).setValues([REQ_HEADERS]);
  }
  return sheet;
}

function pendingRequestRow_(sheet, shiftId) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const vals = sheet.getRange(2, 1, last - 1, REQ_HEADERS.length).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][1]) === String(shiftId) && String(vals[i][15]) === 'cho_duyet') return i + 2;
  }
  return -1;
}

function submitEditRequest(token, shiftId, payload) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'staff') return { success: false, error: 'Chỉ Nhân viên mới gửi yêu cầu sửa.' };
  payload = payload || {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet_();
    const row = findRow_(sheet, shiftId);
    if (row < 0) return { success: false, error: 'Không tìm thấy ca.' };
    const r = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    const name = String(r[2]);
    if (name.trim().toLowerCase() !== String(auth.staff).trim().toLowerCase())
      return { success: false, error: 'Bạn chỉ được yêu cầu sửa ca của mình.' };

    const reqSheet = getRequestsSheet_();
    if (pendingRequestRow_(reqSheet, shiftId) > 0)
      return { success: false, error: 'Ca này đã có 1 yêu cầu đang chờ duyệt.' };

    const startOld = asTimeStr_(r[5]);
    const endOld = asTimeStr_(r[6]);
    const gmvStartOld = Number(r[11]) || 0;
    const gmvEndOld = Number(r[12]) || 0;
    const noteOld = String(r[9] || '');
    const positionOld = String(r[3] || '');
    const channelOld = String(r[4] || '');

    const startNew = (payload.start != null && String(payload.start).trim()) ? String(payload.start).trim() : startOld;
    const endNew = (payload.end != null && String(payload.end).trim()) ? String(payload.end).trim() : endOld;
    const gmvStartNew = (payload.gmvStart != null && String(payload.gmvStart) !== '') ? (Number(payload.gmvStart) || 0) : gmvStartOld;
    const gmvEndNew = (payload.gmvEnd != null && String(payload.gmvEnd) !== '') ? (Number(payload.gmvEnd) || 0) : gmvEndOld;
    const noteNew = (payload.note != null) ? String(payload.note) : noteOld;
    const positionNew = (payload.position != null && String(payload.position).trim()) ? String(payload.position).trim() : positionOld;
    const channelNew = (payload.channel != null && String(payload.channel).trim()) ? String(payload.channel).trim() : channelOld;

    if (!/^\d{1,2}:\d{2}$/.test(startNew)) return { success: false, error: 'Giờ bắt đầu không hợp lệ (định dạng HH:MM).' };
    if (!/^\d{1,2}:\d{2}$/.test(endNew)) return { success: false, error: 'Giờ kết thúc không hợp lệ (định dạng HH:MM).' };
    if (gmvEndNew < gmvStartNew) return { success: false, error: 'GMV cuối ca phải ≥ GMV đầu ca.' };
    if (['Host Live', 'LA'].indexOf(positionNew) < 0) return { success: false, error: 'Vị trí không hợp lệ (chỉ nhận Host Live hoặc LA).' };

    if (startNew === startOld && endNew === endOld && gmvStartNew === gmvStartOld && gmvEndNew === gmvEndOld
        && noteNew === noteOld && positionNew === positionOld && channelNew === channelOld)
      return { success: false, error: 'Bạn chưa thay đổi thông tin nào.' };

    const id = Utilities.getUuid();
    reqSheet.appendRow([
      id, String(shiftId), name, String(auth.username), String(r[4]), asDateStr_(r[1]), nowString_(),
      endOld, endNew, gmvStartOld, gmvStartNew, gmvEndOld, gmvEndNew,
      noteOld, noteNew, 'cho_duyet', '', '', '',
      positionOld, positionNew, channelOld, channelNew,
      startOld, startNew
    ]);
    try {
      sendTelegram_('📝 <b>YÊU CẦU SỬA CA</b> (chờ duyệt)\n👤 ' + esc_(name)
        + '\n📺 ' + esc_(String(r[4])) + '\n📅 ' + asDateStr_(r[1]));
    } catch (e) { /* bỏ qua */ }
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi gửi yêu cầu: ' + e.message };
  } finally { lock.releaseLock(); }
}

function canApproveChannel_(auth, channel) {
  if (auth.role === 'owner') return true;
  if (auth.role === 'lead') return leadHasChannel_(auth, channel);
  return false;
}

function canViewChannelReq_(auth, channel) {
  if (canApproveChannel_(auth, channel)) return true;
  if (auth.role === 'staff') {
    return (auth.managed || []).some(function (c) {
      return String(c).toLowerCase() === String(channel).toLowerCase();
    });
  }
  return false;
}

function listRequests(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  try {
    const sheet = getRequestsSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, data: [] };
    const vals = sheet.getRange(2, 1, last - 1, REQ_HEADERS.length).getValues();

    const shiftSheet = getSheet_();
    const shiftMap = {};
    if (shiftSheet.getLastRow() >= 2) {
      const shifts = shiftSheet.getRange(2, 1, shiftSheet.getLastRow() - 1, HEADERS.length).getValues();
      shifts.forEach(function (v) {
        if (!v[0]) return;
        shiftMap[String(v[0])] = {
          date: asDateStr_(v[1]),
          name: String(v[2]),
          position: String(v[3]),
          channel: String(v[4]),
          start: asTimeStr_(v[5]),
          end: asTimeStr_(v[6]),
          hours: Number(v[7]) || 0,
          gmv: Number(v[8]) || 0,
          note: String(v[9] || ''),
          gmvStart: Number(v[11]) || 0,
          gmvEnd: Number(v[12]) || 0
        };
      });
    }

    const out = [];
    vals.forEach(function (v) {
      const channel = String(v[4]);
      const status = String(v[15]);
      const mineSent = String(v[3]).toLowerCase() === String(auth.username).toLowerCase();
      const canApprove = canApproveChannel_(auth, channel);
      const canView = canViewChannelReq_(auth, channel);
      if (!mineSent && !(canView && status === 'cho_duyet')) return;

      const endOld = asTimeStr_(v[7]);
      const endNew = asTimeStr_(v[8]);
      const startOld = (v[23] != null && v[23] !== '') ? asTimeStr_(v[23]) : '';
      const startNew = (v[24] != null && v[24] !== '') ? asTimeStr_(v[24]) : '';

      out.push({
        id: String(v[0]), shiftId: String(v[1]), requester: String(v[2]), username: String(v[3]),
        channel: channel, date: asDateStr_(v[5]), submittedAt: String(v[6]),
        startOld: startOld, startNew: startNew,
        endOld: endOld, endNew: endNew,
        gmvStartOld: Number(v[9]) || 0, gmvStartNew: Number(v[10]) || 0,
        gmvEndOld: Number(v[11]) || 0, gmvEndNew: Number(v[12]) || 0,
        noteOld: String(v[13]), noteNew: String(v[14]),
        status: status, approver: String(v[16]), decidedAt: String(v[17]), rejectReason: String(v[18]),
        positionOld: String(v[19] || ''), positionNew: String(v[20] || ''),
        channelOld: String(v[21] || ''), channelNew: String(v[22] || ''),
        canApprove: canApprove, canView: canView, mine: mineSent,
        shift: shiftMap[String(v[1])] || null
      });
    });
    out.sort(function (a, b) { return b.submittedAt.localeCompare(a.submittedAt); });
    return { success: true, data: out };
  } catch (e) {
    return { success: false, error: 'Lỗi đọc yêu cầu: ' + e.message };
  }
}

function approveRequest(token, requestId) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const reqSheet = getRequestsSheet_();
    const rrow = findRow_(reqSheet, requestId);
    if (rrow < 0) return { success: false, error: 'Không tìm thấy yêu cầu.' };
    const rq = reqSheet.getRange(rrow, 1, 1, REQ_HEADERS.length).getValues()[0];
    if (String(rq[15]) !== 'cho_duyet') return { success: false, error: 'Yêu cầu đã được xử lý.' };
    const channel = String(rq[4]);
    if (!canApproveChannel_(auth, channel)) return denyErr_();

    const sheet = getSheet_();
    const srow = findRow_(sheet, String(rq[1]));
    if (srow < 0) return { success: false, error: 'Ca đã bị xoá, không thể duyệt.' };
    // TỐI ƯU: đọc cả dòng 1 lần thay vì 3 lệnh getValue() rời rạc (Bắt đầu cũ, Vị trí cũ, Kênh cũ)
    const srowVals = sheet.getRange(srow, 1, 1, HEADERS.length).getValues()[0];
    const oldStart = asTimeStr_(srowVals[5]);
    const positionOld = String(srowVals[3]);
    const channelOld = String(srowVals[4]);
    const createdAt = srowVals[10];   // giữ nguyên "Thời gian tạo"
    const endNew = asTimeStr_(rq[8]);
    const gmvStartNew = Number(rq[10]) || 0;
    const gmvEndNew = Number(rq[12]) || 0;
    const noteNew = String(rq[14]);
    const positionNew = String(rq[20] || '') || positionOld;
    const channelNew = String(rq[22] || '') || channelOld;
    const startNewRaw = rq[24];
    const startNew = (startNewRaw != null && startNewRaw !== '') ? asTimeStr_(startNewRaw) : oldStart;
    const start = startNew || oldStart;
    const hours = computeHours_(start, endNew);
    const gmv = gmvEndNew - gmvStartNew;

    // TỐI ƯU: ghi 1 lần duy nhất cột D→M (Vị trí, Kênh, Bắt đầu, Kết thúc, Giờ live, GMV, Ghi chú,
    // Thời gian tạo, GMV đầu ca, GMV cuối ca) thay vì 9 lệnh setValue() riêng lẻ.
    sheet.getRange(srow, 4, 1, 10).setValues([[
      positionNew, channelNew, start, endNew, hours, gmv, noteNew, createdAt, gmvStartNew, gmvEndNew
    ]]);

    // TỐI ƯU: ghi Trạng thái/Người duyệt/Thời gian duyệt trong 1 lệnh thay vì 3
    reqSheet.getRange(rrow, 16, 1, 3).setValues([[ 'da_duyet', String(auth.username), nowString_() ]]);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi duyệt: ' + e.message };
  } finally { lock.releaseLock(); }
}

function rejectRequest(token, requestId, reason) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const reqSheet = getRequestsSheet_();
    const rrow = findRow_(reqSheet, requestId);
    if (rrow < 0) return { success: false, error: 'Không tìm thấy yêu cầu.' };
    const rq = reqSheet.getRange(rrow, 1, 1, REQ_HEADERS.length).getValues()[0];
    if (String(rq[15]) !== 'cho_duyet') return { success: false, error: 'Yêu cầu đã được xử lý.' };
    if (!canApproveChannel_(auth, String(rq[4]))) return denyErr_();
    reqSheet.getRange(rrow, 16).setValue('tu_choi');
    reqSheet.getRange(rrow, 17).setValue(String(auth.username));
    reqSheet.getRange(rrow, 18).setValue(nowString_());
    reqSheet.getRange(rrow, 19).setValue(String(reason || ''));
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi từ chối: ' + e.message };
  } finally { lock.releaseLock(); }
}

/* ----------------------- CHECK-IN / CHECK-OUT ----------------------- */

function todayStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function timeStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'); }

function roundInTime_(hhmm) {
  var p = String(hhmm || '').split(':');
  var h = Number(p[0]) || 0, m = Number(p[1]) || 0;
  var totalMin = h * 60 + m;
  if (totalMin % 15 === 0) {
    return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
  }
  var rounded = Math.ceil(totalMin / 15) * 15;
  if (rounded >= 1440) rounded -= 1440;
  var rh = Math.floor(rounded / 60), rm = rounded % 60;
  return ('0' + rh).slice(-2) + ':' + ('0' + rm).slice(-2);
}

function roundOutTime_(hhmm) {
  var p = String(hhmm || '').split(':');
  var h = Number(p[0]) || 0, m = Number(p[1]) || 0;
  var totalMin = h * 60 + m;
  var rounded = Math.floor(totalMin / 15) * 15;
  var rh = Math.floor(rounded / 60), rm = rounded % 60;
  return ('0' + rh).slice(-2) + ':' + ('0' + rm).slice(-2);
}

function openShiftRow_(sheet, name) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const vals = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const nm = String(name).toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][2]).toLowerCase() === nm && !String(vals[i][6]).trim()) return i + 2;
  }
  return -1;
}

function checkIn(token, rec, photoB64) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'staff') return { success: false, error: 'Chức năng chấm công chỉ dành cho Nhân viên.' };
  rec = rec || {};
  rec.name = auth.staff;
  if (!photoB64) return { success: false, error: 'Cần chụp ảnh đầu ca.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!String(rec.name || '').trim()) return { success: false, error: 'Thiếu nhân sự.' };
    if (!rec.channel) return { success: false, error: 'Vui lòng chọn kênh.' };
    if (!rec.position) return { success: false, error: 'Vui lòng chọn vị trí.' };
    const sheet = getSheet_();
    if (openShiftRow_(sheet, rec.name) > 0)
      return { success: false, error: 'Đang có một ca chưa check-out. Hãy check-out trước.' };
    const photoId = savePhoto_(photoB64, 'in');
    const id = Utilities.getUuid();
    const gmvStart = Number(rec.gmvStart) || 0;
    const inTime = roundInTime_(timeStr_());
    sheet.appendRow([
      id, todayStr_(), String(rec.name).trim(), rec.position, rec.channel,
      inTime, '', 0, 0,
      String(rec.note || '').trim(), nowString_(), gmvStart, 0, photoId, ''
    ]);
    ensureStaff_(rec.name);
    var newChans = (auth.role === 'staff') ? addUserChannel_(auth.username, rec.channel) : null;
    try {
      sendTelegram_('✅ <b>CHECK-IN</b>\n👤 ' + esc_(String(rec.name).trim()) + ' · ' + esc_(rec.position)
        + '\n📺 ' + esc_(rec.channel) + '\n🕐 ' + inTime);
      var conflicts = hostLiveConflicts_(sheet, id, String(rec.name).trim(), rec.channel, rec.position);
      if (conflicts.length) sendTelegram_('🔴 <b>CẢNH BÁO TRÙNG HOST LIVE</b>\n' + conflicts.join('\n'));
    } catch (e) { /* bỏ qua */ }
    return { success: true, id: id, channels: newChans };
  } catch (e) {
    return { success: false, error: 'Lỗi check-in: ' + e.message };
  } finally { lock.releaseLock(); }
}

function syncStaffChannels_(username, staffName) {
  try {
    const nm = String(staffName || '').trim().toLowerCase();
    if (!nm) return null;
    const sh = getSheet_();
    const last = sh.getLastRow();
    const found = {};
    if (last >= 2) {
      const vals = sh.getRange(2, 3, last - 1, 3).getValues();
      vals.forEach(function (r) {
        if (String(r[0]).trim().toLowerCase() === nm) {
          const ch = String(r[2]).trim();
          if (ch) found[ch] = true;
        }
      });
    }
    const usheet = getUsersSheet_();
    const urow = usersRow_(usheet, username);
    if (urow < 0) return null;
    const cur = splitChannels_(usheet.getRange(urow, 5).getValue());
    const lower = {}; cur.forEach(function (c) { lower[c.toLowerCase()] = true; });
    let changed = false;
    Object.keys(found).forEach(function (ch) {
      if (!lower[ch.toLowerCase()]) { cur.push(ch); changed = true; }
    });
    if (changed) usheet.getRange(urow, 5).setValue(cur.join('|'));
    return cur;
  } catch (e) { return null; }
}

function addUserChannel_(username, channel) {
  try {
    channel = String(channel || '').trim();
    if (!channel) return null;
    const sheet = getUsersSheet_();
    const row = usersRow_(sheet, username);
    if (row < 0) return null;
    const cur = splitChannels_(sheet.getRange(row, 5).getValue());
    const lower = cur.map(function (c) { return c.toLowerCase(); });
    if (lower.indexOf(channel.toLowerCase()) < 0) {
      cur.push(channel);
      sheet.getRange(row, 5).setValue(cur.join('|'));
    }
    return cur;
  } catch (e) { return null; }
}

function setMyChannels(token, channels) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'staff') return { success: false, error: 'Chỉ nhân viên cập nhật kênh của mình.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let chans = Array.isArray(channels) ? channels : splitChannels_(channels);
    chans = chans.map(function (c) { return String(c).trim(); }).filter(Boolean);
    const sheet = getUsersSheet_();
    const row = usersRow_(sheet, auth.username);
    if (row < 0) return { success: false, error: 'Không tìm thấy tài khoản.' };
    sheet.getRange(row, 5).setValue(chans.join('|'));
    return { success: true, channels: chans };
  } catch (e) {
    return { success: false, error: 'Lỗi cập nhật kênh: ' + e.message };
  } finally { lock.releaseLock(); }
}

function checkOut(token, shiftId, gmvEnd, photoB64) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'staff') return { success: false, error: 'Chức năng chấm công chỉ dành cho Nhân viên.' };
  if (!photoB64) return { success: false, error: 'Cần chụp ảnh cuối ca.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getSheet_();
    const row = findRow_(sheet, shiftId);
    if (row < 0) return { success: false, error: 'Không tìm thấy ca.' };
    const name = String(sheet.getRange(row, 3).getValue());
    const channel = String(sheet.getRange(row, 5).getValue());
    if (name.toLowerCase() !== String(auth.staff).toLowerCase()) return denyErr_();
    if (String(sheet.getRange(row, 7).getValue()).trim()) return { success: false, error: 'Ca này đã check-out rồi.' };
    const start = asTimeStr_(sheet.getRange(row, 6).getValue());
    const end = roundOutTime_(timeStr_());
    const gmvStart = Number(sheet.getRange(row, 12).getValue()) || 0;
    const gEnd = Number(gmvEnd) || 0;
    if (gEnd < gmvStart) return { success: false, error: 'GMV cuối ca phải ≥ GMV đầu ca.' };
    var finalEnd = end;
    if (start === end) {
      finalEnd = timeStr_();
    }
    const photoId = savePhoto_(photoB64, 'out');
    sheet.getRange(row, 7).setValue(finalEnd);
    sheet.getRange(row, 8).setValue(computeHours_(start, finalEnd));
    sheet.getRange(row, 9).setValue(gEnd - gmvStart);
    sheet.getRange(row, 13).setValue(gEnd);
    sheet.getRange(row, 15).setValue(photoId);
    try {
      sendTelegram_('🏁 <b>CHECK-OUT</b>\n👤 ' + esc_(name) + '\n📺 ' + esc_(channel)
        + '\n🕐 ' + esc_(start) + ' – ' + esc_(finalEnd) + '\n💰 GMV ca: ' + fmtVnd_(gEnd - gmvStart) + 'đ');
    } catch (e) { /* bỏ qua */ }
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi check-out: ' + e.message };
  } finally { lock.releaseLock(); }
}

/* ===================== THÔNG BÁO TELEGRAM ===================== */

const LATE_SHIFT_HOURS = 6; // ca mở quá X giờ => cảnh báo check-out muộn

function tgConfig_() {
  const p = PropertiesService.getScriptProperties();
  return { token: p.getProperty('TELEGRAM_TOKEN') || '', chatId: p.getProperty('TELEGRAM_CHAT_ID') || '' };
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function fmtVnd_(n) {
  n = Math.round(Number(n) || 0);
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function parseDateTime_(dateStr, timeStr) {
  try {
    const d = String(dateStr).split('-');
    const t = String(timeStr).split(':');
    if (d.length < 3) return null;
    return new Date(Number(d[0]), Number(d[1]) - 1, Number(d[2]), Number(t[0]) || 0, Number(t[1]) || 0, 0);
  } catch (e) { return null; }
}

function sendTelegram_(text) {
  try {
    const cfg = tgConfig_();
    if (!cfg.token || !cfg.chatId) return;
    UrlFetchApp.fetch('https://api.telegram.org/bot' + cfg.token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: cfg.chatId, text: text, parse_mode: 'HTML', disable_web_page_preview: true })
    });
  } catch (e) { /* nuốt lỗi */ }
}

function _show_(msg) { Logger.log(msg); return msg; }

function CAU_HINH_TELEGRAM() {
  var TOKEN = '';
  var CHAT_ID = '';

  var p = PropertiesService.getScriptProperties();
  if (TOKEN) p.setProperty('TELEGRAM_TOKEN', String(TOKEN).trim());
  if (CHAT_ID) p.setProperty('TELEGRAM_CHAT_ID', String(CHAT_ID).trim());
  var cfg = tgConfig_();
  if (!cfg.token) return _show_('❌ Chưa có TOKEN. Dán Bot Token vào biến TOKEN rồi chạy lại.');
  if (!cfg.chatId) return _show_('⚠️ Đã lưu TOKEN. Giờ: thêm bot vào nhóm → trong nhóm gõ /start → rồi chạy hàm LAY_CHAT_ID để lấy Chat ID.');
  sendTelegram_('✅ <b>EAZ Chấm công</b> đã kết nối Telegram thành công!');
  return _show_('✅ Đã lưu cấu hình + gửi tin test vào nhóm. Kiểm tra Telegram nhé.');
}

function LAY_CHAT_ID() {
  var cfg = tgConfig_();
  if (!cfg.token) return _show_('❌ Chưa lưu TOKEN. Chạy CAU_HINH_TELEGRAM (với TOKEN) trước.');
  try {
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + cfg.token + '/getUpdates', { muteHttpExceptions: true });
    var data = JSON.parse(res.getContentText());
    if (!data.ok || !data.result || !data.result.length)
      return _show_('Chưa thấy tin nhắn nào. Hãy: thêm bot vào nhóm → trong nhóm gõ /start (hoặc /id) → chạy lại hàm này.');
    var ids = {};
    data.result.forEach(function (u) {
      var chat = (u.message && u.message.chat) || (u.my_chat_member && u.my_chat_member.chat) || (u.channel_post && u.channel_post.chat);
      if (chat) ids[chat.id] = (chat.title || chat.username || chat.first_name || '');
    });
    var lines = Object.keys(ids).map(function (id) { return 'Chat ID = ' + id + '   (' + ids[id] + ')'; });
    return _show_('✅ Tìm được:\n' + lines.join('\n') + '\n\n→ Dán Chat ID (nhóm thường bắt đầu bằng dấu -) vào CAU_HINH_TELEGRAM rồi chạy lại.');
  } catch (e) { return _show_('Lỗi: ' + e.message); }
}

function hostLiveConflicts_(sheet, selfId, name, channel, position) {
  if (String(position) !== 'Host Live') return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const vals = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const nm = String(name).toLowerCase(), ch = String(channel).toLowerCase();
  let sameChannelOther = null, samePersonOtherChan = null;
  vals.forEach(function (r) {
    if (String(r[0]) === String(selfId)) return;
    if (String(r[6]).trim()) return;
    if (String(r[3]) !== 'Host Live') return;
    const rName = String(r[2]), rChan = String(r[4]);
    if (rChan.toLowerCase() === ch && rName.toLowerCase() !== nm) sameChannelOther = rName;
    if (rName.toLowerCase() === nm && rChan.toLowerCase() !== ch) samePersonOtherChan = rChan;
  });
  const msgs = [];
  if (sameChannelOther) msgs.push('Kênh <b>' + esc_(channel) + '</b> đang có 2 Host Live cùng lúc: ' + esc_(sameChannelOther) + ' &amp; ' + esc_(name) + '.');
  if (samePersonOtherChan) msgs.push('<b>' + esc_(name) + '</b> đang Host Live ở 2 kênh cùng lúc: ' + esc_(samePersonOtherChan) + ' &amp; ' + esc_(channel) + '.');
  return msgs;
}

function checkLateShifts_() {
  try {
    const cfg = tgConfig_();
    if (!cfg.token || !cfg.chatId) return;
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return;
    const vals = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
    const now = new Date();
    const p = PropertiesService.getScriptProperties();
    let warned = {};
    try { warned = JSON.parse(p.getProperty('LATE_WARNED') || '{}'); } catch (e) { warned = {}; }
    const stillOpen = {};
    vals.forEach(function (r) {
      const id = String(r[0]);
      if (String(r[6]).trim()) return;
      const dateStr = asDateStr_(r[1]), startStr = asTimeStr_(r[5]);
      const startDt = parseDateTime_(dateStr, startStr);
      if (!startDt) return;
      const hoursOpen = (now.getTime() - startDt.getTime()) / 3600000;
      if (hoursOpen >= LATE_SHIFT_HOURS) {
        stillOpen[id] = true;
        if (!warned[id]) {
          sendTelegram_('⚠️ <b>CẢNH BÁO CHECK-OUT MUỘN</b>\n'
            + '👤 ' + esc_(r[2]) + ' · ' + esc_(r[3]) + '\n'
            + '📺 ' + esc_(r[4]) + '\n'
            + '🕐 Vào ca ' + esc_(startStr) + ' (' + dateStr + ') — đã mở <b>' + hoursOpen.toFixed(1) + ' giờ</b> mà chưa check-out!');
          warned[id] = true;
        }
      }
    });
    const cleaned = {};
    Object.keys(warned).forEach(function (id) { if (stillOpen[id]) cleaned[id] = true; });
    p.setProperty('LATE_WARNED', JSON.stringify(cleaned));
  } catch (e) { /* nuốt lỗi */ }
}

function createTelegramTrigger() {
  const existing = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'checkLateShifts_') ScriptApp.deleteTrigger(existing[i]);
  }
  ScriptApp.newTrigger('checkLateShifts_').timeBased().everyMinutes(30).create();
  return '✅ Đã tạo lịch quét check-out muộn mỗi 30 phút.';
}

/* ----------------------- DỌN ẢNH CŨ (45 NGÀY) ----------------------- */

function purgeOldPhotos() {
  const folder = getPhotoFolder_();
  const cutoff = new Date(Date.now() - PHOTO_RETENTION_DAYS * 24 * 3600 * 1000);
  const files = folder.getFiles();
  let n = 0;
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated() < cutoff) { f.setTrashed(true); n++; }
  }
  return n;
}

function createPhotoPurgeTrigger() {
  const existing = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'purgeOldPhotos') ScriptApp.deleteTrigger(existing[i]);
  }
  ScriptApp.newTrigger('purgeOldPhotos').timeBased().everyDays(1).atHour(3).create();
  return 'Đã tạo lịch dọn ảnh hằng ngày (3h sáng).';
}

function getDatabaseUrl(token) {
  const auth = verifyToken_(token);
  if (!auth || auth.role !== 'owner') return { success: false };
  try {
    return { success: true, url: getSheet_().getParent().getUrl() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
//  BACKUP TỰ ĐỘNG HẰNG NGÀY
// ============================================================

const BACKUP_KEEP_DAYS = 30;

function getBackupFolder_() {
  const folderName = 'EAZ Backup - Chấm công';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function backupDatabase() {
  try {
    const ss = ss_();
    const fileName = 'Backup_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd_HH-mm');
    const folder = getBackupFolder_();
    DriveApp.getFileById(ss.getId()).makeCopy(fileName, folder);
    const cutoff = new Date(Date.now() - BACKUP_KEEP_DAYS * 24 * 3600 * 1000);
    const files = folder.getFiles();
    let deleted = 0;
    while (files.hasNext()) {
      const f = files.next();
      if (f.getDateCreated() < cutoff) { f.setTrashed(true); deleted++; }
    }
    Logger.log('✅ Backup thành công: ' + fileName + ' | Đã xóa ' + deleted + ' bản cũ.');
    return { success: true, name: fileName, deleted: deleted };
  } catch (e) {
    Logger.log('❌ Backup thất bại: ' + e.message);
    return { success: false, error: e.message };
  }
}

function createBackupTrigger() {
  const existing = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'backupDatabase') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  ScriptApp.newTrigger('backupDatabase').timeBased().everyDays(1).atHour(2).create();
  backupDatabase();
  return '✅ Đã tạo lịch backup hằng ngày (2h sáng). Bản backup đầu tiên vừa được tạo trong folder "EAZ Backup - Chấm công" trên Drive.';
}

function listBackups(token) {
  const auth = verifyToken_(token);
  if (!auth || auth.role !== 'owner') return denyErr_();
  try {
    const folder = getBackupFolder_();
    const files = folder.getFiles();
    const list = [];
    while (files.hasNext()) {
      const f = files.next();
      list.push({
        name: f.getName(),
        date: Utilities.formatDate(f.getDateCreated(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm'),
        url: f.getUrl()
      });
    }
    list.sort(function(a, b) { return b.name.localeCompare(a.name); });
    return { success: true, data: list };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * (CHẠY 1 LẦN trong editor) Chuẩn hoá dữ liệu giờ/GMV của các ca CŨ.
 */
function CHUAN_HOA_CA_CU() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return _show_('Không có dữ liệu.');
  const range = sheet.getRange(2, 1, last - 1, HEADERS.length);
  const vals = range.getValues();
  let fixed = 0, skipped = 0, opens = 0;
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0]) { skipped++; continue; }
    const endRaw = r[6];
    if (!endRaw || !String(endRaw).trim()) { opens++; continue; }
    const startStr = asTimeStr_(r[5]);
    const endStr = asTimeStr_(endRaw);
    if (!/^\d{2}:\d{2}$/.test(startStr) || !/^\d{2}:\d{2}$/.test(endStr)) { skipped++; continue; }
    const hours = computeHours_(startStr, endStr);
    const gmvStart = Number(r[11]) || 0;
    const gmvEnd = Number(r[12]) || 0;
    let gmv = Number(r[8]) || 0;
    if (gmvEnd > 0 || gmvStart > 0) gmv = gmvEnd - gmvStart;
    const curStart = String(r[5]);
    const curEnd = String(r[6]);
    if (curStart === startStr && curEnd === endStr && Number(r[7]) === hours && Number(r[8]) === gmv) continue;
    r[5] = startStr;
    r[6] = endStr;
    r[7] = hours;
    r[8] = gmv;
    fixed++;
  }
  range.setValues(vals);
  sheet.getRange(2, 6, last - 1, 2).setNumberFormat('@');
  return _show_('✅ Đã chuẩn hoá xong. Fix: ' + fixed + ' ca. Ca đang mở (bỏ qua): ' + opens + '. Bỏ qua khác: ' + skipped + '.');
}

function DON_DIEM_HOAN_HUY() {
  const results = [];
  const dm = getSimpleSheet_(DIEM_SHEET, DIEM_HEADERS);
  results.push(cleanupMonthKeyed_(dm, 3, function (r) {
    return asMonthStr_(r[0]) + '|' + String(r[1]).trim().toLowerCase();
  }, function (r) {
    return [asMonthStr_(r[0]), String(r[1]).trim(), Number(r[2]) || 0];
  }, 'DiemThang'));
  const hh = getSimpleSheet_(HOANHUY_SHEET, HOANHUY_HEADERS);
  results.push(cleanupMonthKeyed_(hh, 2, function (r) {
    return asMonthStr_(r[0]);
  }, function (r) {
    return [asMonthStr_(r[0]), Number(r[1]) || 0];
  }, 'HoanHuy'));
  return _show_(results.join('\n'));
}

function cleanupMonthKeyed_(sheet, cols, keyFn, normFn, label) {
  const last = sheet.getLastRow();
  if (last < 2) return '📋 ' + label + ': trống, không cần dọn.';
  const range = sheet.getRange(2, 1, last - 1, cols);
  const vals = range.getValues();
  const byKey = {};
  vals.forEach(function (r, idx) {
    if (!r[0]) return;
    const k = keyFn(r);
    if (!k || k === '|') return;
    byKey[k] = normFn(r);
  });
  const kept = Object.keys(byKey).map(function (k) { return byKey[k]; });
  const before = vals.filter(function (r) { return r[0]; }).length;
  const after = kept.length;
  range.clearContent();
  if (kept.length) {
    sheet.getRange(2, 1, kept.length, cols).setValues(kept);
  }
  sheet.getRange('A:A').setNumberFormat('@');
  return '✅ ' + label + ': ' + before + ' dòng cũ -> ' + after + ' dòng (đã gộp trùng).';
}

function LAM_TRON_CA_CU() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return _show_('Không có dữ liệu.');
  const range = sheet.getRange(2, 1, last - 1, HEADERS.length);
  const vals = range.getValues();
  let rounded = 0, unchanged = 0, opens = 0, imported = 0, skippedCollide = 0;
  const changes = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0]) continue;
    const endRaw = r[6];
    if (!endRaw || !String(endRaw).trim()) { opens++; continue; }

    const note = String(r[9] || '');
    if (/^import/i.test(note.trim())) { imported++; continue; }

    const startOld = asTimeStr_(r[5]);
    const endOld = asTimeStr_(endRaw);
    if (!/^\d{2}:\d{2}$/.test(startOld) || !/^\d{2}:\d{2}$/.test(endOld)) { unchanged++; continue; }

    const startNew = roundInTime_(startOld);
    const endNew = roundOutTime_(endOld);

    if (startNew === startOld && endNew === endOld) { unchanged++; continue; }

    const startMin = Number(startNew.split(':')[0]) * 60 + Number(startNew.split(':')[1]);
    let endMin = Number(endNew.split(':')[0]) * 60 + Number(endNew.split(':')[1]);
    const startOldMin = Number(startOld.split(':')[0]) * 60 + Number(startOld.split(':')[1]);
    const endOldMin = Number(endOld.split(':')[0]) * 60 + Number(endOld.split(':')[1]);
    const wasOvernight = endOldMin <= startOldMin;
    if (!wasOvernight && endMin <= startMin) { skippedCollide++; continue; }

    r[5] = startNew;
    r[6] = endNew;
    r[7] = computeHours_(startNew, endNew);
    rounded++;
    if (changes.length < 10) {
      changes.push('· ' + asDateStr_(r[1]) + ' ' + String(r[2]).trim() + ': ' + startOld + '-' + endOld + ' → ' + startNew + '-' + endNew);
    }
  }
  range.setValues(vals);
  sheet.getRange(2, 6, last - 1, 2).setNumberFormat('@');

  var summary = '✅ Đã làm tròn ' + rounded + ' ca.\n'
    + '· Giữ nguyên (đã đúng mốc hoặc không đổi): ' + unchanged + '\n'
    + '· Ca đang mở (bỏ qua): ' + opens + '\n'
    + '· Ca import lịch sử (bỏ qua): ' + imported + '\n'
    + '· Ca bị bỏ qua vì làm tròn sẽ mất giờ: ' + skippedCollide;
  if (changes.length) summary += '\n\n10 ca đầu đã sửa:\n' + changes.join('\n');
  return _show_(summary);
}

/* ==================== XUẤT BẢNG LƯƠNG EXCEL (batch, tối ưu) ==================== */
function exportPayrollXlsx(token, month, filterChannels) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner' && auth.role !== 'lead') return denyErr_();
  try {
    month = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, error: 'Tháng không hợp lệ (YYYY-MM).' };

    const payResult = computePayroll(token, month, filterChannels);
    if (!payResult.success) return payResult;
    const data = payResult.data || [];
    const chanFilter = Array.isArray(filterChannels) ? filterChannels.filter(function (c) { return !!c; }) : [];
    const chanFilterSet = {};
    chanFilter.forEach(function (c) { chanFilterSet[String(c).toLowerCase()] = true; });
    function passChanFilter(ch) { return !chanFilter.length || chanFilterSet[String(ch || '').toLowerCase()]; }

    const cfgRes = getPayrollConfig(token);
    if (!cfgRes.success) return cfgRes;
    const cfg = cfgRes.chung || {};
    const hourlyList = cfgRes.hourly || [];
    const mocList = cfgRes.moc || [];
    const mocDiem = Number(cfg.moc_diem) || 0;
    const pctHost = (Number(cfg.pct_host) || 0) / 100;
    const pctTro = (Number(cfg.pct_tro) || 0) / 100;

    const rateMap = {};
    hourlyList.forEach(function (h) {
      rateMap[String(h.name).toLowerCase() + '|' + String(h.channel || '').toLowerCase()] = { host: h.host || 0, tro: h.tro || 0 };
    });
    function getRate(name, ch) {
      const nl = String(name).toLowerCase();
      return rateMap[nl + '|' + String(ch).toLowerCase()] || rateMap[nl + '|'] || { host: 0, tro: 0 };
    }
    const mocMap = {};
    mocList.forEach(function (m) {
      mocMap[String(m.name).toLowerCase() + '|' + String(m.channel).toLowerCase() + '|' + m.role] = { mocGmvh: m.mocGmvh || 0, mocNmv: m.mocNmv || 0 };
    });
    function getMoc(name, ch, role) {
      return mocMap[String(name).toLowerCase() + '|' + String(ch).toLowerCase() + '|' + role] || { mocGmvh: 0, mocNmv: 0 };
    }

    const byChannel = {};
    const channelSet = {};
    function ensureChan(ch) { if (!byChannel[ch]) byChannel[ch] = []; channelSet[ch] = true; }

    data.forEach(function (p) {
      const hoTroByChan = {};
      (p.hoTroItems || []).forEach(function (it) {
        const ch = String(it.channel || '').trim();
        const chL = ch.toLowerCase();
        if (!hoTroByChan[chL]) hoTroByChan[chL] = { channel: ch, amount: 0, notes: [] };
        hoTroByChan[chL].amount += Number(it.amount) || 0;
        if (it.note) hoTroByChan[chL].notes.push(String(it.note));
      });
      const luongCungByChan = {};
      const luongCungSrc = (p.luongCungItems || []).concat((p.probationItems || []).map(function (it) {
        return { channel: it.channel, amount: it.amount, note: 'Thử việc' + (it.note ? ': ' + it.note : '') };
      }));
      luongCungSrc.forEach(function (it) {
        const ch = String(it.channel || '').trim();
        const chL = ch.toLowerCase();
        if (!luongCungByChan[chL]) luongCungByChan[chL] = { channel: ch, amount: 0, notes: [] };
        luongCungByChan[chL].amount += Number(it.amount) || 0;
        if (it.note) luongCungByChan[chL].notes.push(String(it.note));
      });

      if (!p.lines || !p.lines.length) {
        const allChanKeys = Object.keys(Object.assign({}, hoTroByChan, luongCungByChan));
        if (allChanKeys.length === 0) return;
        allChanKeys.forEach(function (chL) {
          const chName = (hoTroByChan[chL] && hoTroByChan[chL].channel) || (luongCungByChan[chL] && luongCungByChan[chL].channel) || 'Khác';
          if (!passChanFilter(chName)) return;
          const ch = chName || 'Khác';
          ensureChan(ch);
          const ht = hoTroByChan[chL] || { amount: 0, notes: [] };
          const lc = luongCungByChan[chL] || { amount: 0, notes: [] };
          const allNotes = ht.notes.concat(lc.notes).filter(Boolean).join(' · ');
          byChannel[ch].push({
            name: p.name, channel: ch, position: p.isProbation ? 'Thử việc' : (p.isFixed ? 'Lương cứng' : '—'),
            luongTroLive: 0, luongLive: 0, gioTroLive: 0, gioLive: 0,
            tongLuong: lc.amount,
            gmvhMoc: 0, gmvhThucTe: 0, k1: 0,
            nmvMoc: 0, nmvThucTe: 0, k2: 0,
            diemMoc: mocDiem, diemThucTe: p.diem || 0, k3: 0,
            tongHeSo: 0, pctLiveChinh: 0, pctTroLive: 0, thuong: 0,
            hoTro: ht.amount, hoTroNote: allNotes,
            tongThuNhap: lc.amount + ht.amount
          });
        });
        return;
      }
      const sorted = p.lines.slice().sort(function (a, b) { return b.gmv - a.gmv; });
      const top = sorted[0];
      const hoTroApplied = {}, luongCungApplied = {};
      sorted.forEach(function (l) {
        const ch = l.channel;
        if (!passChanFilter(ch)) return;
        ensureChan(ch);
        const isHost = l.role === 'Host';
        const rate = getRate(p.name, ch);
        const moc = getMoc(p.name, ch, l.role);
        const chL = ch.toLowerCase();
        let hoTroCh = 0, hoTroNoteCh = '';
        if (hoTroByChan[chL] && !hoTroApplied[chL]) {
          hoTroCh = hoTroByChan[chL].amount;
          hoTroNoteCh = hoTroByChan[chL].notes.join(' · ');
          hoTroApplied[chL] = true;
        }
        let lcCh = 0, lcNoteCh = '';
        if (luongCungByChan[chL] && !luongCungApplied[chL]) {
          lcCh = luongCungByChan[chL].amount;
          lcNoteCh = luongCungByChan[chL].notes.join(' · ');
          luongCungApplied[chL] = true;
        }
        const holNote = (l.holidayDays && l.holidayDays.length)
          ? 'Lễ: ' + l.holidayDays.map(function (d) { return d.date.slice(8, 10) + '/' + d.date.slice(5, 7) + ' x' + d.mult + ' (' + d.hours + 'h)'; }).join(', ') + ' (+' + fmtVnd_(l.holidayExtra) + 'đ)'
          : '';
        const finalNote = [hoTroNoteCh, lcNoteCh, holNote].filter(Boolean).join(' · ');
        byChannel[ch].push({
          hoursPayLine: l.hoursPay,
          name: p.name, channel: ch, position: isHost ? 'Host live' : 'LA',
          luongTroLive: rate.tro, luongLive: rate.host,
          gioTroLive: isHost ? 0 : l.hours, gioLive: isHost ? l.hours : 0,
          tongLuong: l.hoursPay + lcCh,
          gmvhMoc: moc.mocGmvh, gmvhThucTe: l.gmvh, k1: l.k1,
          nmvMoc: moc.mocNmv, nmvThucTe: l.nmv, k2: l.k2,
          diemMoc: mocDiem, diemThucTe: l.diem || 0, k3: l.k3,
          tongHeSo: l.tongHeSo,
          pctLiveChinh: pctHost, pctTroLive: pctTro, thuong: l.thuong,
          hoTro: hoTroCh, hoTroNote: finalNote,
          tongThuNhap: l.hoursPay + l.thuong + hoTroCh + lcCh
        });
      });
      const allExtraKeys = Object.keys(Object.assign({}, hoTroByChan, luongCungByChan));
      allExtraKeys.forEach(function (chL) {
        if (hoTroApplied[chL] || luongCungApplied[chL]) return;
        const chName = (hoTroByChan[chL] && hoTroByChan[chL].channel) || (luongCungByChan[chL] && luongCungByChan[chL].channel);
        if (!chName) return;
        if (!passChanFilter(chName)) return;
        ensureChan(chName);
        const ht = hoTroByChan[chL] || { amount: 0, notes: [] };
        const lc = luongCungByChan[chL] || { amount: 0, notes: [] };
        byChannel[chName].push({
          name: p.name, channel: chName, position: '—',
          luongTroLive: 0, luongLive: 0, gioTroLive: 0, gioLive: 0,
          tongLuong: lc.amount,
          gmvhMoc: 0, gmvhThucTe: 0, k1: 0, nmvMoc: 0, nmvThucTe: 0, k2: 0,
          diemMoc: 0, diemThucTe: 0, k3: 0, tongHeSo: 0, pctLiveChinh: 0, pctTroLive: 0, thuong: 0,
          hoTro: ht.amount,
          hoTroNote: ht.notes.concat(lc.notes).filter(Boolean).join(' · '),
          tongThuNhap: ht.amount + lc.amount
        });
      });
      if (top && hoTroByChan[''] && passChanFilter(top.channel)) {
        const rows = byChannel[top.channel] || [];
        for (let i = rows.length - 1; i >= 0; i--) {
          if (rows[i].name === p.name && rows[i].position === (top.role === 'Host' ? 'Host live' : 'LA')) {
            rows[i].hoTro += hoTroByChan[''].amount;
            const noteAdd = hoTroByChan[''].notes.join(' · ');
            if (noteAdd) rows[i].hoTroNote = rows[i].hoTroNote ? (rows[i].hoTroNote + ' · ' + noteAdd) : noteAdd;
            rows[i].tongThuNhap += hoTroByChan[''].amount;
            break;
          }
        }
      }
      if (top && luongCungByChan[''] && passChanFilter(top.channel)) {
        const rows = byChannel[top.channel] || [];
        for (let i = rows.length - 1; i >= 0; i--) {
          if (rows[i].name === p.name && rows[i].position === (top.role === 'Host' ? 'Host live' : 'LA')) {
            rows[i].tongLuong += luongCungByChan[''].amount;
            rows[i].tongThuNhap += luongCungByChan[''].amount;
            break;
          }
        }
      }
    });
    const channelOrder = Object.keys(channelSet).sort();

    // Mã NV theo tên nhân sự (sheet NhanSu, cột C) — điền vào cột "MÃ NV" của bảng lương
    const codeByName = {};
    try {
      const ss_ = getStaffSheet_();
      if (ss_.getLastRow() >= 2) {
        ss_.getRange(2, 1, ss_.getLastRow() - 1, 3).getValues().forEach(function (r) {
          const nm = String(r[0] || '').trim().toLowerCase();
          const cd = String(r[2] || '').trim();
          if (nm && cd) codeByName[nm] = cd;
        });
      }
    } catch (e) { /* không có mã thì để trống, không làm hỏng việc xuất file */ }

    const NCOLS = 34;
    const monthLabel = month.split('-').reverse().join('/');
    const aoa = [];
    aoa.push(padRow_(['Tên Đơn vị: Công ty Cổ phần EAZ Việt Nam'], NCOLS));
    aoa.push(new Array(NCOLS).fill(''));
    const titleRow = new Array(NCOLS).fill('');
    titleRow[5] = 'THANH TOÁN CHI PHÍ CỘNG TÁC — Tháng ' + monthLabel;
    aoa.push(titleRow);
    aoa.push(new Array(NCOLS).fill(''));
    const sh = new Array(NCOLS).fill('');
    sh[0] = 'THÔNG TIN NHÂN SỰ';
    sh[6] = 'VỊ TRÍ';
    sh[7] = 'Lương cứng';
    sh[8] = 'LƯƠNG THEO GIỜ';
    sh[13] = 'GMV/h'; sh[17] = 'NMV/tháng'; sh[21] = 'Điểm phòng live';
    sh[25] = 'TỔNG HỆ SỐ'; sh[26] = 'THƯỞNG THỰC TẾ';
    sh[29] = 'THƯỞNG HỖ TRỢ'; sh[30] = 'TỔNG THU NHẬP'; sh[33] = 'Ghi chú';
    aoa.push(sh);
    aoa.push([
      'STT', 'KHỐI', 'PHÒNG/ BỘ PHẬN', 'MÃ NV\n(000000)', 'TÊN NHÂN VIÊN', 'KÊNH', '', '',
      'Lương trợ live', 'Lương Live', 'Giờ trợ live', 'Giờ Live', 'TỔNG LƯƠNG',
      'Hệ số', 'Mốc thưởng', 'Thực tế', 'Kết quả hệ số',
      'Hệ số', 'Mốc thưởng', 'Thực tế', 'Kết quả hệ số',
      'Hệ số', 'Mốc thưởng', 'Thực tế', 'Kết quả hệ số',
      '', 'LIVE CHÍNH', 'TRỢ LIVE', 'THƯỞNG', '', 'TỔNG THU NHẬP', 'PIT', 'THỰC NHẬN', ''
    ]);

    let stt = 1;
    const channelHeaderRows = [];
    channelOrder.forEach(function (ch) {
      const chRow = new Array(NCOLS).fill('');
      chRow[0] = ch;
      aoa.push(chRow);
      channelHeaderRows.push(aoa.length);
      byChannel[ch].forEach(function (r) {
        const pureLuongGio = (r.hoursPayLine != null) ? r.hoursPayLine : ((r.gioTroLive || 0) * (r.luongTroLive || 0) + (r.gioLive || 0) * (r.luongLive || 0));
        const luongCungCell = Math.max(0, (r.tongLuong || 0) - pureLuongGio);
        aoa.push([
          stt++, '', '', (codeByName[String(r.name).trim().toLowerCase()] || ''), r.name, r.channel, r.position,
          luongCungCell,
          r.luongTroLive, r.luongLive, r.gioTroLive, r.gioLive, pureLuongGio,
          Number((r.k1 || 0).toFixed(2)), r.gmvhMoc, r.gmvhThucTe, Number((r.k1 || 0).toFixed(4)),
          Number((r.k2 || 0).toFixed(2)), r.nmvMoc, r.nmvThucTe, Number((r.k2 || 0).toFixed(4)),
          Number((r.k3 || 0).toFixed(2)), r.diemMoc, r.diemThucTe, Number((r.k3 || 0).toFixed(4)),
          Number((r.tongHeSo || 0).toFixed(4)),
          r.pctLiveChinh, r.pctTroLive, r.thuong,
          r.hoTro,
          pureLuongGio + r.thuong + r.hoTro + luongCungCell,
          0,
          pureLuongGio + r.thuong + r.hoTro + luongCungCell,
          r.hoTroNote
        ]);
      });
    });

    const totalRows = aoa.length;
    if (totalRows < 6) return { success: false, error: 'Không có dữ liệu để xuất.' };

    const fileName = 'BangLuong_' + month + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    const newSs = SpreadsheetApp.create(fileName);
    const sheet = newSs.getActiveSheet();
    sheet.setName('Part time');

    if (sheet.getMaxColumns() > NCOLS) sheet.deleteColumns(NCOLS + 1, sheet.getMaxColumns() - NCOLS);
    if (sheet.getMaxRows() < totalRows) sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());

    if (totalRows >= 7) sheet.getRange(7, 4, totalRows - 6, 1).setNumberFormat('@'); // cột Mã NV dạng chữ
    sheet.getRange(1, 1, totalRows, NCOLS).setValues(aoa);

    const wholeRange = sheet.getRange(1, 1, totalRows, NCOLS);
    wholeRange.setFontFamily('Times New Roman').setVerticalAlignment('middle');

    const mergeRanges = [
      'A1:K1', 'F3:P3', 'A5:F5', 'G5:G6', 'H5:H6', 'I5:M5',
      'N5:Q5', 'R5:U5', 'V5:Y5', 'Z5:Z6', 'AA5:AC5', 'AD5:AD6', 'AE5:AG5', 'AH5:AH6'
    ];
    mergeRanges.forEach(function (rng) {
      try { sheet.getRange(rng).merge(); } catch (e) { }
    });

    sheet.getRange('A1').setFontSize(11).setFontWeight('bold').setHorizontalAlignment('left');
    sheet.getRange('F3').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setRowHeight(3, 30);

    sheet.getRangeList(['A5:F5', 'G5:G6', 'AH5:AH6']).setBackground('#D9E2F3');
    sheet.getRange('H5:H6').setBackground('#BDD7EE');
    sheet.getRange('I5:M5').setBackground('#F7CAAC');
    sheet.getRangeList(['N5:Q5', 'R5:U5', 'V5:Y5', 'Z5:Z6', 'AA5:AC5']).setBackground('#FFE598');
    sheet.getRangeList(['AD5:AD6', 'AE5:AG5']).setBackground('#C5E0B3');

    sheet.getRange(5, 1, 2, NCOLS)
      .setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setWrap(true)
      .setBorder(true, true, true, true, true, true);

    sheet.setRowHeight(5, 30);
    sheet.setRowHeight(6, 40);

    sheet.getRange('A6:F6').setBackground('#D9E2F3');
    sheet.getRange('I6:M6').setBackground('#F7CAAC');
    sheet.getRange('N6:Y6').setBackground('#FFE598');
    sheet.getRange('AA6:AC6').setBackground('#FFE598');
    sheet.getRange('AE6:AG6').setBackground('#C5E0B3');

    if (totalRows >= 7) {
      const nData = totalRows - 6;
      const dataRange = sheet.getRange(7, 1, nData, NCOLS);
      dataRange.setFontSize(10).setBorder(true, true, true, true, true, true);

      sheet.getRange(7, 1, nData, 1).setNumberFormat('0');
      sheet.getRange(7, 8, nData, 1).setNumberFormat('#,##0');
      sheet.getRange(7, 9, nData, 2).setNumberFormat('#,##0');
      sheet.getRange(7, 11, nData, 2).setNumberFormat('0.0');
      sheet.getRange(7, 13, nData, 1).setNumberFormat('#,##0');
      sheet.getRange(7, 14, nData, 1).setNumberFormat('0.00');
      sheet.getRange(7, 15, nData, 2).setNumberFormat('#,##0');
      sheet.getRange(7, 17, nData, 1).setNumberFormat('0.0000');
      sheet.getRange(7, 18, nData, 1).setNumberFormat('0.00');
      sheet.getRange(7, 19, nData, 2).setNumberFormat('#,##0');
      sheet.getRange(7, 21, nData, 1).setNumberFormat('0.0000');
      sheet.getRange(7, 22, nData, 1).setNumberFormat('0.00');
      sheet.getRange(7, 23, nData, 2).setNumberFormat('0.0');
      sheet.getRange(7, 25, nData, 1).setNumberFormat('0.0000');
      sheet.getRange(7, 26, nData, 1).setNumberFormat('0.0000');
      sheet.getRange(7, 27, nData, 2).setNumberFormat('0.00%');
      sheet.getRange(7, 29, nData, 5).setNumberFormat('#,##0');

      sheet.getRange(7, 1, nData, 1).setHorizontalAlignment('center');
      sheet.getRange(7, 5, nData, 1).setHorizontalAlignment('left');
      sheet.getRange(7, 6, nData, 2).setHorizontalAlignment('center');
      sheet.getRange(7, 34, nData, 1).setHorizontalAlignment('left');
    }

    if (channelHeaderRows.length) {
      const ranges = channelHeaderRows.map(function (r) { return 'A' + r + ':AH' + r; });
      sheet.getRangeList(ranges).setBackground('#FFE598').setFontSize(11).setFontWeight('bold');
    }

    const widths = [40, 90, 120, 100, 200, 110, 90,
      110,
      90, 90, 70, 70, 100,
      70, 90, 90, 80,
      70, 90, 90, 80,
      70, 70, 60, 80,
      80,
      80, 80, 100,
      100, 110, 80, 110,
      150
    ];
    for (let i = 0; i < widths.length; i++) sheet.setColumnWidth(i + 1, widths[i]);

    sheet.setFrozenRows(6);

    SpreadsheetApp.flush();

    const fileId = newSs.getId();
    try {
      DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { }
    try {
      const folderName = 'EAZ - Bảng lương xuất';
      const folders = DriveApp.getFoldersByName(folderName);
      const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      const file = DriveApp.getFileById(fileId);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (e) { }

    return {
      success: true,
      url: 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx',
      viewUrl: 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit',
      fileName: fileName + '.xlsx'
    };
  } catch (e) {
    return { success: false, error: 'Lỗi xuất bảng lương: ' + e.message };
  }
}

function padRow_(arr, ncols) {
  const r = arr.slice(0, ncols);
  while (r.length < ncols) r.push('');
  return r;
}

function CHAN_DOAN_BANG_LUONG() {
  const THANG_CAN_KIEM_TRA = '2026-05';

  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) { Logger.log('❌ Sheet CaLive KHÔNG CÓ DÒNG NÀO (rỗng hoàn toàn).'); return; }

  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  Logger.log('Tổng số dòng trong CaLive: ' + rows.length);

  let matchMonth = 0, hasEnd = 0, matchMonthAndEnd = 0;
  const sampleDates = {};
  const sampleRows = [];

  rows.forEach(function (r, idx) {
    const rawDate = r[1];
    const dateStr = asDateStr_(rawDate);
    const monthStr = dateStr.slice(0, 7);
    const end = String(r[6]).trim();

    if (monthStr === THANG_CAN_KIEM_TRA) {
      matchMonth++;
      if (end) matchMonthAndEnd++;
      if (sampleRows.length < 5) {
        sampleRows.push('  Dòng ' + (idx + 2) + ': rawDate=' + JSON.stringify(rawDate) + ' (type=' + (rawDate instanceof Date ? 'Date' : typeof rawDate) + ')'
          + ' -> dateStr=' + dateStr + ' | Tên=' + r[2] + ' | Kênh=' + r[4] + ' | Kết thúc=' + JSON.stringify(r[6]) + ' (end=' + JSON.stringify(end) + ')');
      }
    }
    if (end) hasEnd++;

    const typeKey = (rawDate instanceof Date) ? ('Date:' + monthStr) : ('String:' + monthStr);
    sampleDates[typeKey] = (sampleDates[typeKey] || 0) + 1;
  });

  Logger.log('=== KẾT QUẢ CHẨN ĐOÁN cho tháng ' + THANG_CAN_KIEM_TRA + ' ===');
  Logger.log('Số dòng có Ngày thuộc tháng này: ' + matchMonth);
  Logger.log('Số dòng CÓ Kết thúc (đã đóng ca) trong tháng này: ' + matchMonthAndEnd);
  Logger.log('Tổng số dòng có Kết thúc (toàn sheet, mọi tháng): ' + hasEnd);
  Logger.log('');
  Logger.log('--- 5 dòng mẫu của tháng ' + THANG_CAN_KIEM_TRA + ' (nếu có) ---');
  if (sampleRows.length) sampleRows.forEach(function (s) { Logger.log(s); });
  else Logger.log('  (không có dòng nào khớp tháng này)');
  Logger.log('');
  Logger.log('--- Phân bố kiểu dữ liệu cột Ngày theo tháng (toàn sheet) ---');
  Object.keys(sampleDates).sort().forEach(function (k) {
    Logger.log('  ' + k + ': ' + sampleDates[k] + ' dòng');
  });

  const lockInfo = checkMonthLocked_(THANG_CAN_KIEM_TRA);
  Logger.log('');
  Logger.log('--- Trạng thái chốt tháng ---');
  Logger.log(lockInfo ? JSON.stringify(lockInfo) : '(chưa từng chốt)');

  const snap = readLuongChotSnapshot_(THANG_CAN_KIEM_TRA);
  Logger.log('Số dòng trong LuongChot (snapshot) cho tháng này: ' + snap.length);

  Logger.log('');
  Logger.log('=== HẾT CHẨN ĐOÁN ===');
}

/* ==================== TIKTOK SHOP — LIVE DATA (Preview cho Owner) ==================== */

const TT_SHEET = 'TikTokAuth';
const TT_HEADERS = ['Kênh', 'Open ID', 'Access Token', 'Access Token hết hạn lúc', 'Refresh Token', 'Refresh Token hết hạn lúc', 'Kết nối lúc', 'User Type', 'Granted Scopes'];

function CAU_HINH_TIKTOK() {
  var APP_KEY = '';
  var APP_SECRET = '';

  var p = PropertiesService.getScriptProperties();
  if (APP_KEY) p.setProperty('TT_APP_KEY', String(APP_KEY).trim());
  if (APP_SECRET) p.setProperty('TT_APP_SECRET', String(APP_SECRET).trim());

  var cfg = ttConfig_();
  if (!cfg.appKey || !cfg.appSecret) {
    return _show_('❌ Chưa đủ thông tin. Điền cả APP_KEY và APP_SECRET rồi chạy lại.');
  }
  return _show_('✅ Đã lưu cấu hình TikTok. Vào app (mục dành cho Owner) để kết nối từng kênh.');
}

function ttConfig_() {
  var p = PropertiesService.getScriptProperties();
  return {
    appKey: p.getProperty('TT_APP_KEY') || '',
    appSecret: p.getProperty('TT_APP_SECRET') || ''
  };
}

function getWebAppUrl_() {
  return ScriptApp.getService().getUrl();
}

function getTikTokChannelsStatus(token) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  try {
    const channelSet = {};
    const sheetCa = getSheet_();
    if (sheetCa.getLastRow() >= 2) {
      sheetCa.getRange(2, 5, sheetCa.getLastRow() - 1, 1).getValues().forEach(function (r) {
        const ch = String(r[0] || '').replace(/\s+/g, ' ').trim();
        if (ch) channelSet[ch] = true;
      });
    }
    const channels = Object.keys(channelSet).sort();

    const connMap = {};
    const sheet = getSimpleSheet_(TT_SHEET, TT_HEADERS);
    if (sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, TT_HEADERS.length).getValues().forEach(function (r) {
        if (!r[0]) return;
        connMap[String(r[0])] = {
          openId: r[1], expireAt: r[3], connectedAt: r[6],
          userType: r[7], grantedScopes: r[8]
        };
      });
    }

    const list = channels.map(function (ch) {
      const c = connMap[ch];
      return {
        channel: ch,
        connected: !!c,
        openId: c ? c.openId : '',
        expireAt: c ? c.expireAt : '',
        connectedAt: c ? c.connectedAt : '',
        grantedScopes: c ? c.grantedScopes : ''
      };
    });
    return { success: true, data: list };
  } catch (e) { return { success: false, error: e.message }; }
}

function getTikTokAuthUrl(token, channel) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  channel = String(channel || '').trim();
  if (!channel) return { success: false, error: 'Thiếu tên kênh.' };
  const cfg = ttConfig_();
  if (!cfg.appKey) return { success: false, error: 'Chưa cấu hình TikTok. Chạy hàm CAU_HINH_TIKTOK() trong editor trước.' };
  const stateId = Utilities.getUuid();
  const state = stateId + '|' + Utilities.base64EncodeWebSafe(channel, Utilities.Charset.UTF_8);
  PropertiesService.getScriptProperties().setProperty('TT_OAUTH_STATE_' + stateId, channel);
  const url = 'https://shop.tiktok.com/alliance/creator/auth?app_key=' + encodeURIComponent(cfg.appKey) + '&state=' + encodeURIComponent(state);
  return { success: true, url: url };
}

function handleTikTokOAuthCallback_(code, state) {
  try {
    const cfg = ttConfig_();
    if (!cfg.appKey || !cfg.appSecret) {
      return HtmlService.createHtmlOutput('<p>Lỗi: chưa cấu hình App Key/Secret. Chạy CAU_HINH_TIKTOK() trong editor trước.</p>');
    }
    let channel = '';
    try {
      const parts = String(state || '').split('|');
      if (parts.length === 2) channel = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
    } catch (e) { /* bỏ qua */ }
    if (!channel) {
      return HtmlService.createHtmlOutput('<p>❌ Không xác định được kênh đang cấp quyền (state không hợp lệ). Vui lòng thử lại từ app.</p>');
    }

    const url = 'https://auth.tiktok-shops.com/api/v2/token/get'
      + '?app_key=' + encodeURIComponent(cfg.appKey)
      + '&app_secret=' + encodeURIComponent(cfg.appSecret)
      + '&auth_code=' + encodeURIComponent(code)
      + '&grant_type=authorized_code';
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(resp.getContentText());

    if (!data || !data.data || !data.data.access_token) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif;padding:30px;">'
        + '<h3>❌ Kết nối TikTok thất bại cho kênh ' + esc_(channel) + '</h3>'
        + '<pre style="background:#f3f4f6;padding:12px;border-radius:8px;white-space:pre-wrap;">' + esc_(JSON.stringify(data, null, 2)) + '</pre>'
        + '<p><a href="' + getWebAppUrl_() + '">← Quay lại app</a></p>'
        + '</div>'
      );
    }

    const d = data.data;
    const userType = d.user_type;
    const grantedScopes = Array.isArray(d.granted_scopes) ? d.granted_scopes.join(', ') : String(d.granted_scopes || '');

    const sheet = getSimpleSheet_(TT_SHEET, TT_HEADERS);
    if (sheet.getLastRow() >= 2) {
      const chCol = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = chCol.length - 1; i >= 0; i--) {
        if (String(chCol[i][0]).trim() === channel) sheet.deleteRow(i + 2);
      }
    }
    const now = nowString_();
    sheet.appendRow([
      channel,
      d.open_id || '',
      d.access_token || '',
      Utilities.formatDate(new Date(Date.now() + (d.access_token_expire_in || 0) * 1000), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      d.refresh_token || '',
      Utilities.formatDate(new Date(Date.now() + (d.refresh_token_expire_in || 0) * 1000), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      now,
      userType != null ? String(userType) : '',
      grantedScopes
    ]);

    const scopeWarning = grantedScopes && grantedScopes.indexOf('creator.data.live') < 0
      ? '<p style="color:#b45309;background:#fef3c7;padding:10px 14px;border-radius:8px;">⚠️ Kênh chưa cấp đủ quyền Live Data (granted_scopes: ' + esc_(grantedScopes) + '). Cần ngắt kết nối và cấp quyền lại, đảm bảo tick chọn Live Data.</p>'
      : '';

    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:30px;text-align:center;">'
      + '<h2>✅ Đã kết nối kênh "' + esc_(channel) + '" thành công!</h2>'
      + '<p>Có thể đóng tab này và quay lại app.</p>'
      + scopeWarning
      + '<p><a href="' + getWebAppUrl_() + '" style="display:inline-block;margin-top:12px;padding:10px 20px;background:#7c3aed;color:#fff;border-radius:8px;text-decoration:none;">← Quay lại app</a></p>'
      + '</div>'
    );
  } catch (e) {
    return HtmlService.createHtmlOutput('<p>Lỗi xử lý callback: ' + esc_(e.message) + '</p>');
  }
}

function disconnectTikTokChannel(token, channel) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  channel = String(channel || '').trim();
  if (!channel) return { success: false, error: 'Thiếu tên kênh.' };
  try {
    const sheet = getSimpleSheet_(TT_SHEET, TT_HEADERS);
    if (sheet.getLastRow() >= 2) {
      const chCol = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = chCol.length - 1; i >= 0; i--) {
        if (String(chCol[i][0]).trim() === channel) sheet.deleteRow(i + 2);
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function ttGetValidAccessToken_(channel) {
  const sheet = getSimpleSheet_(TT_SHEET, TT_HEADERS);
  if (sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, TT_HEADERS.length).getValues();
  let rowIdx = -1, r = null;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === channel) { rowIdx = i + 2; r = rows[i]; break; }
  }
  if (!r) return null;
  const accessToken = r[2], expireAt = r[3], refreshToken = r[4];
  if (!accessToken) return null;

  const expireDate = new Date(expireAt);
  if (expireDate.getTime() - Date.now() > 5 * 60 * 1000) return accessToken;

  const cfg = ttConfig_();
  const url = 'https://auth.tiktok-shops.com/api/v2/token/refresh'
    + '?app_key=' + encodeURIComponent(cfg.appKey)
    + '&app_secret=' + encodeURIComponent(cfg.appSecret)
    + '&refresh_token=' + encodeURIComponent(refreshToken)
    + '&grant_type=refresh_token';
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data = JSON.parse(resp.getContentText());
  if (!data || !data.data || !data.data.access_token) return null;

  const d = data.data;
  sheet.getRange(rowIdx, 3, 1, 4).setValues([[
    d.access_token,
    Utilities.formatDate(new Date(Date.now() + (d.access_token_expire_in || 0) * 1000), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    d.refresh_token || refreshToken,
    Utilities.formatDate(new Date(Date.now() + (d.refresh_token_expire_in || 0) * 1000), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  ]]);
  return d.access_token;
}

function ttSignRequest_(path, params) {
  const cfg = ttConfig_();
  params = params || {};
  params.app_key = cfg.appKey;
  params.timestamp = Math.floor(Date.now() / 1000);

  const sortedKeys = Object.keys(params).sort();
  let baseString = path;
  sortedKeys.forEach(function (k) { baseString += k + params[k]; });
  const signSrc = cfg.appSecret + baseString + cfg.appSecret;

  const rawHmac = Utilities.computeHmacSha256Signature(signSrc, cfg.appSecret);
  const sign = rawHmac.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');

  params.sign = sign;
  return params;
}

function ttCallApi_(channel, path, extraParams) {
  const accessToken = ttGetValidAccessToken_(channel);
  if (!accessToken) throw new Error('Kênh "' + channel + '" chưa kết nối TikTok hoặc token đã hết hạn. Vào app kết nối lại.');
  const params = ttSignRequest_(path, extraParams || {});
  const query = Object.keys(params).sort().map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  const url = 'https://open-api.tiktokglobalshop.com' + path + '?' + query;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'x-tts-access-token': accessToken },
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText());
}

function getTikTokLiveDataPreview(token, channel) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  channel = String(channel || '').trim();
  if (!channel) return { success: false, error: 'Thiếu tên kênh.' };
  try {
    const traffic = ttCallApi_(channel, '/live/202502/traffic_performance', {});
    const interactive = ttCallApi_(channel, '/live/202502/interactive_performance', {});
    return { success: true, channel: channel, traffic: traffic, interactive: interactive };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/* ============================================================
 * QUẢN LÝ NHÂN SỰ NÂNG CAO — MERGE + EDIT MÃ NV
 * mergeStaff:      gộp toàn bộ dữ liệu tên cũ → tên mới (Owner only)
 * updateStaffCode: cập nhật Mã NV; phát hiện trùng mã (Owner only)
 * ============================================================ */

/**
 * Gộp nhân sự: chuyển tất cả dữ liệu (ca live, đơn giá, mốc thưởng, điểm, hỗ trợ,
 * lương cứng, account) từ oldName sang newName, sau đó xoá oldName khỏi NhanSu.
 * Conflict-safe: nếu newName đã có cùng key (ch/role/month), giữ nguyên bản của newName.
 */
function mergeStaff(token, oldName, newName) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    oldName = String(oldName || '').trim();
    newName = String(newName || '').trim();
    if (!oldName || !newName) return { success: false, error: 'Thiếu tên cũ hoặc tên mới.' };
    if (oldName.toLowerCase() === newName.toLowerCase()) return { success: false, error: 'Tên cũ và tên mới trùng nhau.' };

    const counts = { shifts: 0, hourly: 0, moc: 0, diem: 0, users: 0, hoTro: 0, luongCung: 0, salary: 0 };
    const oldL = oldName.toLowerCase();
    const newL = newName.toLowerCase();

    // 1. CaLive — cột C (Tên)
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last >= 2) {
      const names = sheet.getRange(2, 3, last - 1, 1).getValues();
      for (let i = 0; i < names.length; i++) {
        if (String(names[i][0]).trim().toLowerCase() === oldL) {
          sheet.getRange(i + 2, 3).setValue(newName);
          counts.shifts++;
        }
      }
    }

    // 2. LuongGio — conflict theo (Tên × Kênh)
    const lg = getLuongGioSheet_();
    if (lg.getLastRow() >= 2) {
      const vals = lg.getRange(2, 1, lg.getLastRow() - 1, 4).getValues();
      const tgtKeys = {}; // channelLower -> true
      const srcRows = [];
      for (let i = 0; i < vals.length; i++) {
        const nm = String(vals[i][0]).trim().toLowerCase();
        const ch = String(vals[i][1] || '').trim().toLowerCase();
        if (nm === newL) tgtKeys[ch] = true;
        if (nm === oldL) srcRows.push({ row: i + 2, ch: ch });
      }
      const toDelete = [];
      for (let i = 0; i < srcRows.length; i++) {
        if (tgtKeys[srcRows[i].ch]) toDelete.push(srcRows[i].row);
        else { lg.getRange(srcRows[i].row, 1).setValue(newName); counts.hourly++; }
      }
      toDelete.sort(function (a, b) { return b - a; });
      toDelete.forEach(function (r) { lg.deleteRow(r); counts.hourly++; });
    }

    // 3. ThuongMoc — conflict theo (Tên × Kênh × Vai trò)
    const mocSheet = getSimpleSheet_(THUONGMOC_SHEET, THUONGMOC_HEADERS);
    if (mocSheet.getLastRow() >= 2) {
      const vals = mocSheet.getRange(2, 1, mocSheet.getLastRow() - 1, 5).getValues();
      const tgtKeys = {};
      const srcRows = [];
      for (let i = 0; i < vals.length; i++) {
        const nm = String(vals[i][0]).trim().toLowerCase();
        const k = String(vals[i][1]).trim().toLowerCase() + '|' + String(vals[i][2]);
        if (nm === newL) tgtKeys[k] = true;
        if (nm === oldL) srcRows.push({ row: i + 2, key: k });
      }
      const toDelete = [];
      for (let i = 0; i < srcRows.length; i++) {
        if (tgtKeys[srcRows[i].key]) toDelete.push(srcRows[i].row);
        else { mocSheet.getRange(srcRows[i].row, 1).setValue(newName); counts.moc++; }
      }
      toDelete.sort(function (a, b) { return b - a; });
      toDelete.forEach(function (r) { mocSheet.deleteRow(r); counts.moc++; });
    }

    // 4. DiemThang — conflict theo (Tháng)
    const dm = getSimpleSheet_(DIEM_SHEET, DIEM_HEADERS);
    if (dm.getLastRow() >= 2) {
      const vals = dm.getRange(2, 1, dm.getLastRow() - 1, 3).getValues();
      const tgtMonths = {};
      const srcRows = [];
      for (let i = 0; i < vals.length; i++) {
        const nm = String(vals[i][1]).trim().toLowerCase();
        const month = asMonthStr_(vals[i][0]);
        if (nm === newL) tgtMonths[month] = true;
        if (nm === oldL) srcRows.push({ row: i + 2, month: month });
      }
      const toDelete = [];
      for (let i = 0; i < srcRows.length; i++) {
        if (tgtMonths[srcRows[i].month]) toDelete.push(srcRows[i].row);
        else { dm.getRange(srcRows[i].row, 2).setValue(newName); counts.diem++; }
      }
      toDelete.sort(function (a, b) { return b - a; });
      toDelete.forEach(function (r) { dm.deleteRow(r); counts.diem++; });
    }

    // 5. ThuongHoTro v2 — col C = Tên, không có conflict (mỗi dòng có ID riêng)
    const hs = ensureHoTroV2_();
    if (hs.getLastRow() >= 2) {
      const vals = hs.getRange(2, 3, hs.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim().toLowerCase() === oldL) {
          hs.getRange(i + 2, 3).setValue(newName);
          counts.hoTro++;
        }
      }
    }

    // 6. LuongCungThang — conflict theo (Tháng × Kênh)
    const lcs = getSimpleSheet_(LUONGCUNG_THANG_SHEET, LUONGCUNG_THANG_HEADERS);
    if (lcs.getLastRow() >= 2) {
      const vals = lcs.getRange(2, 1, lcs.getLastRow() - 1, 5).getValues();
      const tgtKeys = {};
      const srcRows = [];
      for (let i = 0; i < vals.length; i++) {
        const nm = String(vals[i][1]).trim().toLowerCase();
        const month = asMonthStr_(vals[i][0]);
        const ch = String(vals[i][2] || '').trim().toLowerCase();
        const k = month + '|' + ch;
        if (nm === newL) tgtKeys[k] = true;
        if (nm === oldL) srcRows.push({ row: i + 2, key: k });
      }
      const toDelete = [];
      for (let i = 0; i < srcRows.length; i++) {
        if (tgtKeys[srcRows[i].key]) toDelete.push(srcRows[i].row);
        else { lcs.getRange(srcRows[i].row, 2).setValue(newName); counts.luongCung++; }
      }
      toDelete.sort(function (a, b) { return b - a; });
      toDelete.forEach(function (r) { lcs.deleteRow(r); counts.luongCung++; });
    }

    // 7. NhanSuLuongCung — col A = Tên (nếu có thì rename hoặc gộp)
    try {
      const lcStaff = getSimpleSheet_(LUONGCUNG_STAFF_SHEET, LUONGCUNG_STAFF_HEADERS);
      if (lcStaff.getLastRow() >= 2) {
        const vals = lcStaff.getRange(2, 1, lcStaff.getLastRow() - 1, 1).getValues();
        let src = -1, tgt = -1;
        for (let i = 0; i < vals.length; i++) {
          const nm = String(vals[i][0]).trim().toLowerCase();
          if (nm === oldL) src = i + 2;
          else if (nm === newL) tgt = i + 2;
        }
        if (src > 0) {
          if (tgt > 0) lcStaff.deleteRow(src);
          else lcStaff.getRange(src, 1).setValue(newName);
        }
      }
    } catch (e) { /* bỏ qua */ }

    // 8. SALARY_SHEET (Luong legacy) — conflict theo (Vị trí)
    try {
      const oldSalary = getSalarySheet_();
      if (oldSalary.getLastRow() >= 2) {
        const vals = oldSalary.getRange(2, 1, oldSalary.getLastRow() - 1, 2).getValues();
        const tgtPos = {};
        const srcRows = [];
        for (let i = 0; i < vals.length; i++) {
          const nm = String(vals[i][0]).trim().toLowerCase();
          const pos = String(vals[i][1] || '').trim().toLowerCase();
          if (nm === newL) tgtPos[pos] = true;
          if (nm === oldL) srcRows.push({ row: i + 2, pos: pos });
        }
        const toDelete = [];
        for (let i = 0; i < srcRows.length; i++) {
          if (tgtPos[srcRows[i].pos]) toDelete.push(srcRows[i].row);
          else { oldSalary.getRange(srcRows[i].row, 1).setValue(newName); counts.salary++; }
        }
        toDelete.sort(function (a, b) { return b - a; });
        toDelete.forEach(function (r) { oldSalary.deleteRow(r); counts.salary++; });
      }
    } catch (e) { /* bỏ qua */ }

    // 9. Users — cột Nhân sự (col F, index 6 in 1-based)
    const usersSheet = getUsersSheet_();
    if (usersSheet.getLastRow() >= 2) {
      const vals = usersSheet.getRange(2, 6, usersSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim().toLowerCase() === oldL) {
          usersSheet.getRange(i + 2, 6).setValue(newName);
          counts.users++;
        }
      }
    }

    // 10. NhanSu — xoá tên cũ, đảm bảo tên mới
    const staffSheet = getStaffSheet_();
    const oldStaffRow = staffRow_(staffSheet, oldName);
    if (oldStaffRow > 0) staffSheet.deleteRow(oldStaffRow);
    if (staffRow_(staffSheet, newName) < 0) staffSheet.appendRow([newName, nowString_(), '']);

    return { success: true, counts: counts };
  } catch (e) {
    return { success: false, error: 'Lỗi gộp nhân sự: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cập nhật Mã NV cho 1 nhân sự.
 * Nếu newCode đã dùng cho nhân sự khác → KHÔNG ghi, trả về duplicate: {otherName, otherCode}.
 * Frontend sẽ dùng thông tin duplicate để hỏi user có muốn gộp không.
 */
function updateStaffCode(token, name, newCode) {
  const auth = verifyToken_(token);
  if (!auth) return authErr_();
  if (auth.role !== 'owner') return denyErr_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    name = String(name || '').trim();
    newCode = String(newCode || '').trim();
    if (!name) return { success: false, error: 'Thiếu tên nhân sự.' };

    const sheet = getStaffSheet_();
    const row = staffRow_(sheet, name);
    if (row < 0) return { success: false, error: 'Không tìm thấy nhân sự.' };

    // Kiểm tra trùng mã với nhân sự KHÁC (không tính chính mình)
    let duplicate = null;
    if (newCode) {
      const last = sheet.getLastRow();
      if (last >= 2) {
        const vals = sheet.getRange(2, 1, last - 1, 3).getValues();
        const codeL = newCode.toLowerCase();
        for (let i = 0; i < vals.length; i++) {
          if (i + 2 === row) continue;
          const otherCode = String(vals[i][2] || '').trim();
          if (otherCode && otherCode.toLowerCase() === codeL) {
            duplicate = { otherName: String(vals[i][0]).trim(), otherCode: otherCode };
            break;
          }
        }
      }
    }

    if (duplicate) {
      // KHÔNG ghi mã ngay — chờ user quyết định (gộp hay đổi mã khác)
      return { success: true, duplicate: duplicate };
    }

    // Không trùng — ghi mã bình thường
    sheet.getRange(row, 3).setValue(newCode);
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Lỗi cập nhật mã: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
