import { sql, Auth, ok, fail, authErr, denyErr, verifyToken, leadHasChannel, nowString, asTimeStr, computeHours, esc, sendTelegram } from "./lib.ts";
import { getStaff, getChannels } from "./group1.ts";

const L = (s: unknown) => String(s ?? "").trim().toLowerCase();
const num = (v: unknown) => Number(v) || 0;
const dstr = (v: unknown) => (v ? (v instanceof Date ? v.toISOString().slice(0, 10) : String(v).slice(0, 10)) : "");

function canApproveChannel(auth: Auth, channel: string) {
  if (auth.role === "owner") return true;
  if (auth.role === "lead") return leadHasChannel(auth, channel);
  return false;
}
function canViewChannelReq(auth: Auth, channel: string) {
  if (canApproveChannel(auth, channel)) return true;
  if (auth.role === "staff") return auth.managed.some((c) => c.toLowerCase() === String(channel).toLowerCase());
  return false;
}

export async function submitEditRequest(token: string, shiftId: string, payload: any) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "staff") return fail("Chỉ Nhân viên mới gửi yêu cầu sửa.");
  payload = payload || {};
  try {
    const s = await sql`select * from chamcong.ca_live where id=${shiftId}`;
    if (!s.length) return fail("Không tìm thấy ca.");
    const r = s[0];
    const name = String(r.ten);
    if (L(name) !== L(auth.staff)) return fail("Bạn chỉ được yêu cầu sửa ca của mình.");
    const pend = await sql`select 1 from chamcong.yeu_cau where ma_ca=${shiftId} and trang_thai='cho_duyet'`;
    if (pend.length) return fail("Ca này đã có 1 yêu cầu đang chờ duyệt.");

    const startOld = asTimeStr(r.bat_dau), endOld = asTimeStr(r.ket_thuc);
    const gmvStartOld = num(r.gmv_dau), gmvEndOld = num(r.gmv_cuoi);
    const noteOld = String(r.ghi_chu || ""), positionOld = String(r.vi_tri || ""), channelOld = String(r.kenh || "");
    const has = (v: unknown) => v != null && String(v).trim() !== "";
    const startNew = has(payload.start) ? String(payload.start).trim() : startOld;
    const endNew = has(payload.end) ? String(payload.end).trim() : endOld;
    const gmvStartNew = (payload.gmvStart != null && String(payload.gmvStart) !== "") ? num(payload.gmvStart) : gmvStartOld;
    const gmvEndNew = (payload.gmvEnd != null && String(payload.gmvEnd) !== "") ? num(payload.gmvEnd) : gmvEndOld;
    const noteNew = payload.note != null ? String(payload.note) : noteOld;
    const positionNew = has(payload.position) ? String(payload.position).trim() : positionOld;
    const channelNew = has(payload.channel) ? String(payload.channel).trim() : channelOld;

    if (!/^\d{1,2}:\d{2}$/.test(startNew)) return fail("Giờ bắt đầu không hợp lệ (định dạng HH:MM).");
    if (!/^\d{1,2}:\d{2}$/.test(endNew)) return fail("Giờ kết thúc không hợp lệ (định dạng HH:MM).");
    if (gmvEndNew < gmvStartNew) return fail("GMV cuối ca phải ≥ GMV đầu ca.");
    if (!["Host Live", "LA"].includes(positionNew)) return fail("Vị trí không hợp lệ (chỉ nhận Host Live hoặc LA).");
    if (startNew === startOld && endNew === endOld && gmvStartNew === gmvStartOld && gmvEndNew === gmvEndOld && noteNew === noteOld && positionNew === positionOld && channelNew === channelOld)
      return fail("Bạn chưa thay đổi thông tin nào.");

    const id = crypto.randomUUID(), date = dstr(r.ngay);
    await sql`insert into chamcong.yeu_cau(id,ma_ca,nguoi_gui,username,kenh,ngay_ca,gui_luc,kt_cu,kt_moi,gmv_dau_cu,gmv_dau_moi,gmv_cuoi_cu,gmv_cuoi_moi,ghi_chu_cu,ghi_chu_moi,trang_thai,nguoi_duyet,duyet_luc,ly_do_tu_choi,vi_tri_cu,vi_tri_moi,kenh_cu,kenh_moi,bd_cu,bd_moi)
      values(${id},${shiftId},${name},${auth.username},${String(r.kenh)},${date},${nowString()},${endOld},${endNew},${gmvStartOld},${gmvStartNew},${gmvEndOld},${gmvEndNew},${noteOld},${noteNew},'cho_duyet','','','',${positionOld},${positionNew},${channelOld},${channelNew},${startOld},${startNew})`;
    try { await sendTelegram(`📝 <b>YÊU CẦU SỬA CA</b> (chờ duyệt)\n👤 ${esc(name)}\n📺 ${esc(r.kenh)}\n📅 ${date}`); } catch { /* bỏ qua */ }
    return ok();
  } catch (e) { return fail("Lỗi gửi yêu cầu: " + (e as Error).message); }
}

export async function listRequests(token: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    const rows = await sql`select * from chamcong.yeu_cau where trang_thai='cho_duyet' or lower(trim(username))=${L(auth.username)} order by gui_luc desc`;
    const ids = [...new Set(rows.map((v: any) => String(v.ma_ca)))];
    const shiftMap: Record<string, any> = {};
    if (ids.length) {
      const ss = await sql`select * from chamcong.ca_live where id = any(${ids})`;
      ss.forEach((v: any) => {
        shiftMap[String(v.id)] = { date: dstr(v.ngay), name: String(v.ten), position: String(v.vi_tri ?? ""), channel: String(v.kenh ?? ""), start: asTimeStr(v.bat_dau), end: asTimeStr(v.ket_thuc),
          hours: num(v.gio_live), gmv: num(v.gmv), note: String(v.ghi_chu || ""), gmvStart: num(v.gmv_dau), gmvEnd: num(v.gmv_cuoi) };
      });
    }
    const out: any[] = [];
    rows.forEach((v: any) => {
      const channel = String(v.kenh), status = String(v.trang_thai);
      const mine = L(v.username) === L(auth.username);
      const canApprove = canApproveChannel(auth, channel), canView = canViewChannelReq(auth, channel);
      if (!mine && !(canView && status === "cho_duyet")) return;
      out.push({
        id: String(v.id), shiftId: String(v.ma_ca), requester: String(v.nguoi_gui), username: String(v.username), channel, date: String(v.ngay_ca || ""), submittedAt: String(v.gui_luc || ""),
        startOld: v.bd_cu ? asTimeStr(v.bd_cu) : "", startNew: v.bd_moi ? asTimeStr(v.bd_moi) : "",
        endOld: asTimeStr(v.kt_cu), endNew: asTimeStr(v.kt_moi),
        gmvStartOld: num(v.gmv_dau_cu), gmvStartNew: num(v.gmv_dau_moi), gmvEndOld: num(v.gmv_cuoi_cu), gmvEndNew: num(v.gmv_cuoi_moi),
        noteOld: String(v.ghi_chu_cu || ""), noteNew: String(v.ghi_chu_moi || ""),
        status, approver: String(v.nguoi_duyet || ""), decidedAt: String(v.duyet_luc || ""), rejectReason: String(v.ly_do_tu_choi || ""),
        positionOld: String(v.vi_tri_cu || ""), positionNew: String(v.vi_tri_moi || ""), channelOld: String(v.kenh_cu || ""), channelNew: String(v.kenh_moi || ""),
        canApprove, canView, mine, shift: shiftMap[String(v.ma_ca)] || null,
      });
    });
    return ok({ data: out });
  } catch (e) { return fail("Lỗi đọc yêu cầu: " + (e as Error).message); }
}

export async function approveRequest(token: string, requestId: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    return await sql.begin(async (tx: any) => {
      const q = await tx`select * from chamcong.yeu_cau where id=${requestId} for update`;
      if (!q.length) return fail("Không tìm thấy yêu cầu.");
      const rq = q[0];
      if (String(rq.trang_thai) !== "cho_duyet") return fail("Yêu cầu đã được xử lý.");
      if (!canApproveChannel(auth, String(rq.kenh))) return denyErr();
      const s = await tx`select * from chamcong.ca_live where id=${String(rq.ma_ca)} for update`;
      if (!s.length) return fail("Ca đã bị xoá, không thể duyệt.");
      const sv = s[0];
      const oldStart = asTimeStr(sv.bat_dau);
      const positionNew = String(rq.vi_tri_moi || "") || String(sv.vi_tri);
      const channelNew = String(rq.kenh_moi || "") || String(sv.kenh);
      const start = (rq.bd_moi ? asTimeStr(rq.bd_moi) : oldStart) || oldStart;
      const endNew = asTimeStr(rq.kt_moi), gs = num(rq.gmv_dau_moi), ge = num(rq.gmv_cuoi_moi);
      await tx`update chamcong.ca_live set vi_tri=${positionNew}, kenh=${channelNew}, bat_dau=${start}, ket_thuc=${endNew}, gio_live=${computeHours(start, endNew)},
        gmv=${ge - gs}, ghi_chu=${String(rq.ghi_chu_moi || "")}, gmv_dau=${gs}, gmv_cuoi=${ge} where id=${String(rq.ma_ca)}`;
      await tx`update chamcong.yeu_cau set trang_thai='da_duyet', nguoi_duyet=${auth.username}, duyet_luc=${nowString()} where id=${requestId}`;
      return ok();
    });
  } catch (e) { return fail("Lỗi duyệt: " + (e as Error).message); }
}

export async function rejectRequest(token: string, requestId: string, reason: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    const q = await sql`select kenh,trang_thai from chamcong.yeu_cau where id=${requestId}`;
    if (!q.length) return fail("Không tìm thấy yêu cầu.");
    if (String(q[0].trang_thai) !== "cho_duyet") return fail("Yêu cầu đã được xử lý.");
    if (!canApproveChannel(auth, String(q[0].kenh))) return denyErr();
    await sql`update chamcong.yeu_cau set trang_thai='tu_choi', nguoi_duyet=${auth.username}, duyet_luc=${nowString()}, ly_do_tu_choi=${String(reason || "")} where id=${requestId}`;
    return ok();
  } catch (e) { return fail("Lỗi từ chối: " + (e as Error).message); }
}

/* ======================= IMPORT CA LIVE ======================= */
const p2 = (s: string | number) => ("0" + s).slice(-2);
function parseImportDate(s0: unknown) {
  const s = String(s0 ?? "").trim(); if (!s) return "";
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/); if (m) return m[1] + "-" + p2(m[2]) + "-" + p2(m[3]);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/); if (m) return m[3] + "-" + p2(m[2]) + "-" + p2(m[1]);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2})$/); if (m) return "20" + m[3] + "-" + p2(m[2]) + "-" + p2(m[1]);
  return s;
}
function parseImportTime(v: unknown) {
  let s = String(v ?? "").trim().toLowerCase().replace(/h/g, ":").replace(/\./g, ":");
  if (/:$/.test(s)) s += "00";
  let m = s.match(/^(\d{1,2}):(\d{1,2})$/);
  if (m) return p2(parseInt(m[1], 10) % 24) + ":" + p2(Math.min(59, parseInt(m[2], 10)));
  m = s.match(/^(\d{1,2})$/);
  if (m) return p2(parseInt(m[1], 10) % 24) + ":00";
  return "";
}
function parseImportNumber(v: unknown) {
  let s = String(v ?? "").trim(); if (!s) return 0;
  if (/e/i.test(s)) { const n = parseFloat(s); return isNaN(n) ? 0 : Math.round(n); }
  s = s.replace(/\s/g, "");
  const seps = (s.match(/[.,]/g) || []).length;
  if (seps >= 2) return parseInt(s.replace(/[^\d]/g, ""), 10) || 0;
  if (seps === 1) {
    const m = s.match(/[.,](\d+)$/);
    if (m && m[1].length === 3) return parseInt(s.replace(/[^\d]/g, ""), 10) || 0;
    return Math.round(parseFloat(s.replace(",", "."))) || 0;
  }
  return parseInt(s.replace(/[^\d]/g, ""), 10) || 0;
}
function parseImportHours(v: unknown) { const n = parseFloat(String(v ?? "").replace(/[^\d.,]/g, "").replace(",", ".")); return isNaN(n) ? 0 : Math.round(n * 100) / 100; }
function hoursToHHMM(h: number) { const t = Math.round((Number(h) || 0) * 60); return p2(Math.floor(t / 60)) + ":" + p2(t % 60); }

type ImpRow = { id: string; date: string; name: string; position: string; channel: string; start: string; end: string; hours: number; gmv: number };
function parseImportRows(text: string, mode: string) {
  mode = mode === "hours" ? "hours" : "times";
  const lines = String(text || "").split(/\r?\n/);
  const rows: ImpRow[] = [], errors: string[] = [];
  let headerSkipped = false;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim(); if (!line) continue;
    const cols = line.includes("\t") ? line.split("\t") : line.split(",");
    if (!headerSkipped && (/ng[aà]y/i.test(cols[0] || "") || /t[eê]n/i.test(cols[1] || ""))) { headerSkipped = true; continue; }
    const date = parseImportDate(String(cols[0] || "").trim());
    const name = String(cols[1] || "").trim(), channel = String(cols[2] || "").trim(), position = String(cols[3] || "").trim();
    if (!date || !name) { errors.push(`Dòng ${i + 1}: thiếu Ngày hoặc Tên.`); continue; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) { errors.push(`Dòng ${i + 1}: ngày không hợp lệ.`); continue; }
    let start: string, end: string, hours: number, gmv: number;
    if (mode === "times") {
      start = parseImportTime(cols[4]); end = parseImportTime(cols[5]); gmv = parseImportNumber(cols[6]);
      if (!start || !end) { errors.push(`Dòng ${i + 1}: giờ bắt đầu/kết thúc không hợp lệ.`); continue; }
      hours = computeHours(start, end);
    } else { hours = parseImportHours(cols[4]); gmv = parseImportNumber(cols[5]); start = "00:00"; end = hoursToHHMM(hours); }
    rows.push({ id: crypto.randomUUID(), date, name, position, channel, start, end, hours, gmv });
  }
  return { rows, errors };
}

export async function previewImport(token: string, text: string, mode: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  text = String(text || "").trim(); if (!text) return fail("Chưa có dữ liệu.");
  try {
    const parsed = parseImportRows(text, mode), rows = parsed.rows;
    if (!rows.length) return fail("Không có dòng hợp lệ. " + parsed.errors.slice(0, 3).join(" "));
    const existStaff: Record<string, boolean> = {}; ((await getStaff(token) as any).data || []).forEach((s: string) => (existStaff[L(s)] = true));
    const existChan: Record<string, boolean> = {}; ((await getChannels(token) as any).data || []).forEach((c: string) => (existChan[L(c)] = true));
    const byName: Record<string, any> = {}, chanSet: Record<string, boolean> = {}; let minD = "", maxD = "";
    rows.forEach((r) => {
      const o = (byName[r.name] = byName[r.name] || { count: 0, hours: 0, gmv: 0, chans: {} });
      o.count++; o.hours += r.hours; o.gmv += r.gmv; if (r.channel) { o.chans[r.channel] = true; chanSet[r.channel] = true; }
      if (!minD || r.date < minD) minD = r.date; if (!maxD || r.date > maxD) maxD = r.date;
    });
    const employees = Object.keys(byName).sort().map((nm) => ({ name: nm, count: byName[nm].count, hours: Math.round(byName[nm].hours * 100) / 100, gmv: byName[nm].gmv, channels: Object.keys(byName[nm].chans).sort(), isNew: !existStaff[L(nm)] }));
    const newChannels = Object.keys(chanSet).filter((c) => !existChan[L(c)]).sort();
    return ok({ valid: rows.length, skipped: parsed.errors.length, errors: parsed.errors.slice(0, 5), employees, newChannels, dateFrom: minD, dateTo: maxD });
  } catch (e) { return fail("Lỗi đọc dữ liệu: " + (e as Error).message); }
}

export async function importShifts(token: string, text: string, mode: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  text = String(text || "").trim(); if (!text) return fail("Chưa có dữ liệu để nhập.");
  try {
    const parsed = parseImportRows(text, mode), rows = parsed.rows;
    if (!rows.length) return fail("Không có dòng hợp lệ. " + parsed.errors.slice(0, 3).join(" "));
    const now = nowString();
    const objs = rows.map((r) => ({ id: r.id, ngay: r.date, ten: r.name, vi_tri: r.position, kenh: r.channel, bat_dau: r.start, ket_thuc: r.end, gio_live: r.hours, gmv: r.gmv, ghi_chu: "Import", tao_luc: now, gmv_dau: 0, gmv_cuoi: r.gmv, anh_dau: "", anh_cuoi: "" }));
    await sql.begin(async (tx: any) => {
      for (let i = 0; i < objs.length; i += 500)
        await tx`insert into chamcong.ca_live ${tx(objs.slice(i, i + 500))}`;
      const names = [...new Set(rows.map((r) => r.name))], chans = [...new Set(rows.map((r) => r.channel).filter(Boolean))];
      for (const n of names) if (!(await tx`select 1 from chamcong.nhan_su where lower(trim(ten))=${L(n)}`).length) await tx`insert into chamcong.nhan_su(ten,tao_luc,ma_nv) values(${n},${now},'')`;
      for (const c of chans) if (!(await tx`select 1 from chamcong.kenh where lower(trim(ten))=${L(c)}`).length) await tx`insert into chamcong.kenh(ten,tao_luc) values(${c},${now})`;
    });
    return ok({ imported: rows.length, skipped: parsed.errors.length, errors: parsed.errors.slice(0, 5) });
  } catch (e) { return fail("Lỗi import: " + (e as Error).message); }
}
