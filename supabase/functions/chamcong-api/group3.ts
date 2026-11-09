import { sql, Auth, ok, fail, authErr, denyErr, verifyToken, leadChannelsLower, leadHasChannel, nowString, esc, fmtVnd, sendTelegram } from "./lib.ts";
import { leadStaffSet } from "./group1.ts";
import { buildCfg, computeRaw, filterPayrollList, buildAllowedChanSet, PayInput } from "./payroll.ts";

const L = (s: unknown) => String(s ?? "").trim().toLowerCase();
const num = (v: unknown) => Number(v) || 0;
const MONTH_RE = /^\d{4}-\d{2}$/;
const lockTx = (tx: any) => tx`select pg_advisory_xact_lock(7001)`;

export function prevMonthStr(m: string) {
  const p = String(m || "").split("-"); let y = Number(p[0]), mo = Number(p[1]);
  if (!y || !mo) return "";
  mo -= 1; if (mo < 1) { mo = 12; y -= 1; }
  return y + "-" + ("0" + mo).slice(-2);
}
function nextMonthStr(m: string) {
  const p = m.split("-"); let y = Number(p[0]), mo = Number(p[1]) + 1; if (mo > 12) { mo = 1; y++; }
  return y + "-" + ("0" + mo).slice(-2);
}

async function canEditName(auth: Auth, name: string) {
  if (auth.role === "owner") return true;
  if (auth.role === "lead") { const s = await leadStaffSet(auth); return !!s[L(name)]; }
  return false;
}

/* ---------- đọc dữ liệu ---------- */
type Lock = { locked: true; by: string; at: string; note: string; rate: number };
async function readLocks(): Promise<Record<string, Lock>> {
  const r = await sql`select thang, trang_thai, nguoi_chot, chot_luc, ghi_chu, ty_le_hoan_huy from chamcong.thang_chot`;
  const m: Record<string, Lock> = {};
  r.forEach((x: any) => {
    const k = String(x.thang).slice(0, 7);
    if (String(x.trang_thai).trim() === "da_chot" && !m[k]) m[k] = { locked: true, by: String(x.nguoi_chot || ""), at: String(x.chot_luc || ""), note: String(x.ghi_chu || ""), rate: num(x.ty_le_hoan_huy) };
  });
  return m;
}
async function readSnapshot(month: string) {
  const r = await sql`select * from chamcong.luong_chot where thang=${month} order by ctid`;
  return r.map((x: any) => {
    const J = (s: unknown) => { try { return JSON.parse(String(s || "[]")); } catch { return []; } };
    const lines = J(x.lines_json);
    return {
      name: String(x.ten), hoursPay: num(x.tong_luong_gio), bonus: num(x.tong_thuong), hoTro: num(x.thuong_ho_tro), hoTroNote: String(x.ghi_chu_ho_tro || ""),
      hoTroItems: J(x.ho_tro_items_json), diem: num(x.diem), total: num(x.tong_thu_nhap), luongCung: num(x.luong_cung), luongCungItems: J(x.luong_cung_items_json),
      probation: num(x.thu_viec), probationItems: J(x.thu_viec_items_json), isProbation: !lines.length && num(x.thu_viec) > 0, lines,
      isFixed: num(x.luong_cung) > 0 && num(x.tong_luong_gio) === 0 && num(x.tong_thuong) === 0,
    };
  });
}
async function loadBase() {
  const [chung, hourly, moc, hoanHuy, diem, hoTro, luongCung, probation, holidays] = await Promise.all([
    sql`select khoa, gia_tri from chamcong.thuong_chung`,
    sql`select ten, kenh, don_gia_host, don_gia_tro from chamcong.luong_gio`,
    sql`select ten, kenh, vai_tro, moc_gmvh, moc_nmv from chamcong.thuong_moc`,
    sql`select thang, ty_le from chamcong.hoan_huy`,
    sql`select thang, ten, diem from chamcong.diem_thang`,
    sql`select thang, ten, kenh, so_tien, ghi_chu from chamcong.thuong_ho_tro order by ctid`,
    sql`select thang, ten, kenh, so_tien, ghi_chu from chamcong.luong_cung_thang order by ctid`,
    sql`select thang, ten, kenh, so_tien, ghi_chu from chamcong.luong_thu_viec order by ctid`,
    sql`select ngay, he_so from chamcong.ngay_le`,
  ]);
  return {
    cfg: buildCfg(chung as any), hourly: hourly as any, moc: moc as any, hoanHuy: hoanHuy as any, diem: diem as any,
    hoTro: hoTro as any, luongCung: luongCung as any, probation: probation as any, holidays: holidays as any,
  };
}
async function loadShifts(months: string[]) {
  const from = months.slice().sort()[0] + "-01", to = nextMonthStr(months.slice().sort().slice(-1)[0]) + "-01";
  const r = await sql`select ngay::text ngay, ten, vi_tri, kenh, bat_dau, ket_thuc, gio_live, gmv from chamcong.ca_live
    where ngay >= ${from}::date and ngay < ${to}::date and coalesce(trim(ket_thuc),'') <> '' order by ngay, id`;
  return r as any[];
}
type Ctx = { base: Awaited<ReturnType<typeof loadBase>>; shifts: any[]; locks: Record<string, Lock>; leadSet: Record<string, boolean> };
async function makeCtx(auth: Auth, months: string[]): Promise<Ctx> {
  const [base, shifts, locks] = await Promise.all([loadBase(), loadShifts(months), readLocks()]);
  const leadSet = auth.role === "lead" ? await leadStaffSet(auth) : {};
  return { base, shifts, locks, leadSet };
}
async function rawFor(auth: Auth, month: string, c: Ctx, filterChannels?: string[]) {
  let prevSnap: any[] | null = null;
  if (!c.locks[month]) { const pm = prevMonthStr(month); if (pm && c.locks[pm]) prevSnap = await readSnapshot(pm); }
  const inp: PayInput = { ...c.base, shifts: c.shifts, prevSnapshotForDiem: prevSnap, leadStaff: c.leadSet };
  return computeRaw(auth as any, month, inp, filterChannels);
}
async function snapshotFor(auth: Auth, month: string, c: Ctx, filterChannels?: string[]) {
  let s = await readSnapshot(month);
  if (auth.role === "lead") s = s.filter((x: any) => c.leadSet[L(x.name)]);
  else if (auth.role === "staff") s = s.filter((x: any) => L(x.name) === L(auth.staff));
  s = filterPayrollList(s, buildAllowedChanSet(auth as any, filterChannels));
  s.sort((a: any, b: any) => b.total - a.total);
  return s;
}

export async function getMonthLockStatus(token: string, month: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ.");
  try { const l = (await readLocks())[month]; return ok({ month, lock: l || null }); } catch (e) { return fail((e as Error).message); }
}

export async function computePayroll(token: string, month: string, filterChannels?: string[]) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ (YYYY-MM).");
    const c = await makeCtx(auth, [month]);
    const lk = c.locks[month];
    if (lk) {
      const data = await snapshotFor(auth, month, c, filterChannels);
      return { success: true, month, returnRate: lk.rate, data, locked: true, lockInfo: lk };
    }
    return await rawFor(auth, month, c, filterChannels);
  } catch (e) { return fail("Lỗi tính lương: " + (e as Error).message); }
}

export async function computePayrollWithPrev(token: string, month: string, filterChannels?: string[]) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ (YYYY-MM).");
    const prevMonth = prevMonthStr(month);
    const c = await makeCtx(auth, prevMonth ? [prevMonth, month] : [month]);
    let cur: any;
    if (c.locks[month]) cur = { returnRate: c.locks[month].rate, data: await snapshotFor(auth, month, c, filterChannels), locked: true, lockInfo: c.locks[month] };
    else { const raw: any = await rawFor(auth, month, c, filterChannels); if (!raw.success) return raw; cur = { returnRate: raw.returnRate, data: raw.data, locked: false, lockInfo: null }; }
    let prevData: any[] = [];
    if (prevMonth) {
      if (c.locks[prevMonth]) prevData = await snapshotFor(auth, prevMonth, c, filterChannels).then((x) => x); // đã lọc + sắp xếp
      else { const pr: any = await rawFor(auth, prevMonth, c, filterChannels); prevData = pr.success ? pr.data : []; }
    }
    return { success: true, month, returnRate: cur.returnRate, data: cur.data, locked: cur.locked, lockInfo: cur.lockInfo, prevMonth, prevData };
  } catch (e) { return fail("Lỗi tính lương: " + (e as Error).message); }
}

/* ---------- cấu hình ---------- */
export async function getPayrollConfig(token: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    const base = await loadBase();
    let scopeNames: Record<string, boolean> | null = null, scopeChannels: Record<string, boolean> | null = null;
    if (auth.role === "lead") { scopeNames = await leadStaffSet(auth); scopeChannels = {}; leadChannelsLower(auth).forEach((c) => (scopeChannels![c] = true)); }
    const hourly: any[] = [];
    base.hourly.forEach((r: any) => {
      if (!r.ten) return; const nm = String(r.ten).trim(), ch = String(r.kenh || "").trim();
      if (scopeNames && !scopeNames[nm.toLowerCase()]) return;
      if (ch && scopeChannels && !scopeChannels[ch.toLowerCase()]) return;
      hourly.push({ name: nm, channel: ch, host: num(r.don_gia_host), tro: num(r.don_gia_tro) });
    });
    const moc: any[] = [];
    base.moc.forEach((r: any) => {
      if (!r.ten) return; const nm = String(r.ten).trim(), ch = String(r.kenh).trim();
      if (scopeNames && !scopeNames[nm.toLowerCase()]) return;
      if (scopeChannels && !scopeChannels[ch.toLowerCase()]) return;
      moc.push({ name: nm, channel: ch, role: String(r.vai_tro), mocGmvh: num(r.moc_gmvh), mocNmv: num(r.moc_nmv) });
    });
    const sc = await sql`select distinct regexp_replace(trim(ten),'\s+',' ','g') nm, regexp_replace(trim(kenh),'\s+',' ','g') ch from chamcong.ca_live where coalesce(trim(ten),'')<>'' and coalesce(trim(kenh),'')<>''`;
    const staffChannels: Record<string, { name: string; channels: Record<string, boolean> }> = {};
    sc.forEach((r: any) => {
      const nm = String(r.nm), ch = String(r.ch), nl = nm.toLowerCase();
      if (scopeNames && !scopeNames[nl]) return;
      if (scopeChannels && !scopeChannels[ch.toLowerCase()]) return;
      (staffChannels[nl] = staffChannels[nl] || { name: nm, channels: {} }).channels[ch] = true;
    });
    const staffChannelList = Object.values(staffChannels).map((v) => ({ name: v.name, channels: Object.keys(v.channels).sort() })).sort((a, b) => a.name.localeCompare(b.name, "vi"));
    return { success: true, chung: base.cfg, hourly, moc, staffChannels: staffChannelList,
      scopeNames: scopeNames ? Object.keys(scopeNames) : null, scopeChannels: scopeChannels ? Object.keys(scopeChannels) : null };
  } catch (e) { return fail("Lỗi đọc cấu hình lương: " + (e as Error).message); }
}

async function channelExists(ch: string) { return (await sql`select 1 from chamcong.kenh where lower(trim(ten))=${L(ch)}`).length > 0; }

export async function saveLuongGio(token: string, name: string, channel: string, rateHost: number, rateTro: number) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (!(await canEditName(auth, name))) return denyErr();
  try {
    name = String(name || "").trim(); channel = String(channel || "").trim();
    if (!name) return fail("Thiếu tên nhân sự.");
    if (channel && !(await channelExists(channel))) return fail(`Kênh "${channel}" không tồn tại trong danh sách kênh. Vui lòng chọn từ dropdown thay vì nhập tay.`);
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      const u = await tx`update chamcong.luong_gio set ten=${name}, kenh=${channel}, don_gia_host=${num(rateHost)}, don_gia_tro=${num(rateTro)}
        where lower(trim(ten))=${L(name)} and lower(trim(coalesce(kenh,'')))=${L(channel)} returning 1`;
      if (!u.length) await tx`insert into chamcong.luong_gio(ten,kenh,don_gia_host,don_gia_tro) values(${name},${channel},${num(rateHost)},${num(rateTro)})`;
    });
    return ok();
  } catch (e) { return fail("Lỗi lưu đơn giá: " + (e as Error).message); }
}

export async function deleteLuongGio(token: string, name: string, channel: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (!(await canEditName(auth, name))) return denyErr();
  try {
    name = String(name || "").trim(); channel = String(channel || "").trim();
    if (!name) return fail("Thiếu tên nhân sự.");
    if (!channel) return fail("Không thể xoá dòng mặc định — chỉ có thể sửa về 0.");
    const r = await sql`delete from chamcong.luong_gio where lower(trim(ten))=${L(name)} and lower(trim(coalesce(kenh,'')))=${L(channel)} returning 1`;
    return r.length ? ok() : fail("Không tìm thấy dòng để xoá.");
  } catch (e) { return fail("Lỗi xoá đơn giá: " + (e as Error).message); }
}

export async function saveThuongMoc(token: string, name: string, channel: string, role: string, mocGmvh: number, mocNmv: number) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (!(await canEditName(auth, name))) return denyErr();
  if (auth.role === "lead" && !leadHasChannel(auth, channel)) return denyErr();
  try {
    name = String(name || "").trim(); channel = String(channel || "").trim();
    role = String(role) === "Host" ? "Host" : "Tro";
    if (!name || !channel) return fail("Thiếu tên hoặc kênh.");
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      const u = await tx`update chamcong.thuong_moc set ten=${name}, kenh=${channel}, moc_gmvh=${num(mocGmvh)}, moc_nmv=${num(mocNmv)}
        where lower(trim(ten))=${L(name)} and lower(trim(kenh))=${L(channel)} and vai_tro=${role} returning 1`;
      if (!u.length) await tx`insert into chamcong.thuong_moc(ten,kenh,vai_tro,moc_gmvh,moc_nmv) values(${name},${channel},${role},${num(mocGmvh)},${num(mocNmv)})`;
    });
    return ok();
  } catch (e) { return fail("Lỗi lưu mốc thưởng: " + (e as Error).message); }
}

export async function saveThuongChung(token: string, cfg: Record<string, unknown>) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      const cur = await tx`select khoa, gia_tri from chamcong.thuong_chung`;
      const merged = buildCfg(cur as any);
      Object.keys(cfg || {}).forEach((k) => { if (merged[k] !== undefined) merged[k] = num((cfg as any)[k]); });
      await tx`delete from chamcong.thuong_chung`;
      for (const k of Object.keys(merged)) await tx`insert into chamcong.thuong_chung(khoa,gia_tri) values(${k},${String(merged[k])})`;
    });
    return ok();
  } catch (e) { return fail("Lỗi lưu cấu hình: " + (e as Error).message); }
}

export async function saveHoanHuy(token: string, month: string, ratePct: number) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner" && auth.role !== "lead") return denyErr();
  month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ (YYYY-MM).");
  try {
    await sql.begin(async (tx: any) => { await lockTx(tx); await setHoanHuy(tx, month, ratePct); });
    return ok();
  } catch (e) { return fail("Lỗi lưu hoàn huỷ: " + (e as Error).message); }
}
async function setHoanHuy(tx: any, month: string, ratePct: number) {
  const u = await tx`update chamcong.hoan_huy set ty_le=${num(ratePct)} where thang=${month} returning 1`;
  if (!u.length) await tx`insert into chamcong.hoan_huy(thang,ty_le) values(${month},${num(ratePct)})`;
}
async function setDiem(tx: any, month: string, name: string, diem: number) {
  const u = await tx`update chamcong.diem_thang set diem=${num(diem)}, ten=${name} where thang=${month} and lower(trim(ten))=${L(name)} returning 1`;
  if (!u.length) await tx`insert into chamcong.diem_thang(thang,ten,diem) values(${month},${name},${num(diem)})`;
}
export async function saveDiem(token: string, month: string, name: string, diem: number) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (!(await canEditName(auth, name))) return denyErr();
  month = String(month || "").trim(); name = String(name || "").trim();
  if (!MONTH_RE.test(month) || !name) return fail("Thiếu tháng hoặc tên.");
  try { await sql.begin(async (tx: any) => { await lockTx(tx); await setDiem(tx, month, name, diem); }); return ok(); }
  catch (e) { return fail("Lỗi lưu điểm: " + (e as Error).message); }
}

/* ---------- ghi snapshot chốt ---------- */
async function writeSnapshot(tx: any, month: string, data: any[], by: string, now: string) {
  await tx`delete from chamcong.luong_chot where thang=${month}`;
  for (const p of data) {
    await tx`insert into chamcong.luong_chot(thang,ten,tong_luong_gio,tong_thuong,thuong_ho_tro,ghi_chu_ho_tro,diem,tong_thu_nhap,lines_json,chot_boi,chot_luc,luong_cung,ho_tro_items_json,luong_cung_items_json,thu_viec,thu_viec_items_json)
      values(${month},${p.name},${num(p.hoursPay)},${num(p.bonus)},${num(p.hoTro)},${String(p.hoTroNote || "")},${num(p.diem)},${num(p.total)},${JSON.stringify(p.lines || [])},${by},${now},${num(p.luongCung)},${JSON.stringify(p.hoTroItems || [])},${JSON.stringify(p.luongCungItems || [])},${num(p.probation)},${JSON.stringify(p.probationItems || [])})`;
  }
}

export async function saveMonthUpdate(token: string, month: string, ratePct: number, diemList: any[], hoTroList: any[], filterChannels?: string[]) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner" && auth.role !== "lead") return denyErr();
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ (YYYY-MM).");
    const lk = (await readLocks())[month];
    if (lk && auth.role !== "owner") return fail(`Tháng ${month} đã chốt bởi ${lk.by}. Chỉ Owner mới có thể sửa.`);
    diemList = Array.isArray(diemList) ? diemList : []; hoTroList = Array.isArray(hoTroList) ? hoTroList : [];
    const leadSet = auth.role === "lead" ? await leadStaffSet(auth) : {};
    const can = (n: string) => auth.role === "owner" || (auth.role === "lead" && !!leadSet[L(n)]);

    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      await setHoanHuy(tx, month, ratePct);
      for (const it of diemList) { if (it && it.name && can(it.name)) await setDiem(tx, month, String(it.name).trim(), num(it.diem)); }
      if (hoTroList.length) {
        const touched: Record<string, boolean> = {};
        hoTroList.forEach((it) => { if (it && it.name && can(it.name)) touched[L(it.name)] = true; });
        const names = Object.keys(touched);
        if (names.length) await tx`delete from chamcong.thuong_ho_tro where thang=${month} and lower(trim(ten)) = any(${names})`;
        for (const it of hoTroList) {
          if (!it || !it.name) continue;
          const nm = String(it.name).trim(); if (!can(nm)) continue;
          const amt = num(it.amount), note = String(it.note || ""), ch = String(it.channel || "").trim();
          if (amt === 0 && !note && !ch) continue;
          await tx`insert into chamcong.thuong_ho_tro(id,thang,ten,kenh,so_tien,ghi_chu) values(${crypto.randomUUID()},${month},${nm},${ch},${amt},${note})`;
        }
      }
      // Tháng đã chốt (Owner sửa): tính lại từ cấu hình mới và cập nhật snapshot
      if (lk && auth.role === "owner") {
        const [base, shifts] = await Promise.all([loadBaseTx(tx), loadShiftsTx(tx, month)]);
        const raw: any = computeRaw(auth as any, month, { ...base, shifts, prevSnapshotForDiem: null, leadStaff: {} } as PayInput, undefined);
        if (raw.success && raw.data && raw.data.length) await writeSnapshot(tx, month, raw.data, auth.username + " (sửa sau chốt)", nowString());
      }
    });
    return await computePayroll(token, month, filterChannels);
  } catch (e) { return fail("Lỗi cập nhật tháng: " + (e as Error).message); }
}
// bản đọc trong transaction (thấy dữ liệu vừa ghi)
async function loadBaseTx(tx: any) {
  const [chung, hourly, moc, hoanHuy, diem, hoTro, luongCung, probation, holidays] = await Promise.all([
    tx`select khoa, gia_tri from chamcong.thuong_chung`, tx`select ten, kenh, don_gia_host, don_gia_tro from chamcong.luong_gio`,
    tx`select ten, kenh, vai_tro, moc_gmvh, moc_nmv from chamcong.thuong_moc`, tx`select thang, ty_le from chamcong.hoan_huy`,
    tx`select thang, ten, diem from chamcong.diem_thang`, tx`select thang, ten, kenh, so_tien, ghi_chu from chamcong.thuong_ho_tro order by ctid`,
    tx`select thang, ten, kenh, so_tien, ghi_chu from chamcong.luong_cung_thang order by ctid`, tx`select thang, ten, kenh, so_tien, ghi_chu from chamcong.luong_thu_viec order by ctid`,
    tx`select ngay, he_so from chamcong.ngay_le`]);
  return { cfg: buildCfg(chung as any), hourly, moc, hoanHuy, diem, hoTro, luongCung, probation, holidays } as any;
}
async function loadShiftsTx(tx: any, month: string) {
  return await tx`select ngay::text ngay, ten, vi_tri, kenh, bat_dau, ket_thuc, gio_live, gmv from chamcong.ca_live
    where ngay >= ${month + "-01"}::date and ngay < ${nextMonthStr(month) + "-01"}::date and coalesce(trim(ket_thuc),'') <> '' order by ngay, id`;
}

/* ---------- chốt / mở chốt ---------- */
export async function chotLuongThang(token: string, month: string, note: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return fail("Chỉ Owner mới có quyền chốt lương tháng.");
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ (YYYY-MM).");
    const ex = (await readLocks())[month];
    if (ex) return fail(`Tháng ${month} đã được chốt trước đó (bởi ${ex.by}).`);
    const pay: any = await computePayroll(token, month);
    if (!pay.success) return pay;
    const data: any[] = pay.data || [];
    if (!data.length) return fail("Không có dữ liệu lương để chốt.");
    const now = nowString(), rate = num(pay.returnRate);
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      await writeSnapshot(tx, month, data, auth.username, now);
      const u = await tx`update chamcong.thang_chot set trang_thai='da_chot', nguoi_chot=${auth.username}, chot_luc=${now}, ghi_chu=${String(note || "")}, ty_le_hoan_huy=${rate} where thang=${month} returning 1`;
      if (!u.length) await tx`insert into chamcong.thang_chot(thang,trang_thai,nguoi_chot,chot_luc,ghi_chu,ty_le_hoan_huy) values(${month},'da_chot',${auth.username},${now},${String(note || "")},${rate})`;
    });
    try {
      await sendTelegram(`🔒 <b>CHỐT LƯƠNG THÁNG ${esc(month)}</b>\n👤 ${esc(auth.username)}\n👥 ${data.length} nhân viên · Tổng: ${fmtVnd(data.reduce((s, x) => s + (x.total || 0), 0))}đ` + (note ? `\n📝 ${esc(String(note))}` : ""));
    } catch { /* bỏ qua */ }
    return { success: true, count: data.length, at: now };
  } catch (e) { return fail("Lỗi chốt lương: " + (e as Error).message); }
}

export async function moChotLuongThang(token: string, month: string, reason: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return fail("Chỉ Owner mới có quyền mở chốt lương tháng.");
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ.");
    const r = await sql`select ghi_chu from chamcong.thang_chot where thang=${month}`;
    if (!r.length) return fail(`Tháng ${month} chưa được chốt.`);
    const newNote = String(r[0].ghi_chu || "") + `\n[Mở chốt lúc ${nowString()} bởi ${auth.username}${reason ? ": " + reason : ""}]`;
    await sql`update chamcong.thang_chot set trang_thai='mo_chot', ghi_chu=${newNote} where thang=${month}`;
    try { await sendTelegram(`🔓 <b>MỞ CHỐT LƯƠNG THÁNG ${esc(month)}</b>\n👤 ${esc(auth.username)}` + (reason ? `\n📝 ${esc(String(reason))}` : "")); } catch { /* bỏ qua */ }
    return ok();
  } catch (e) { return fail("Lỗi mở chốt: " + (e as Error).message); }
}

/* ---------- lương cứng ---------- */
export async function listLuongCungStaff(token: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    const r = await sql`select ten, vai_tro, ghi_chu, trang_thai, tao_luc from chamcong.nhan_su_luong_cung`;
    const list = r.filter((x: any) => x.ten).map((x: any) => ({ name: String(x.ten).trim(), role: String(x.vai_tro || ""), note: String(x.ghi_chu || ""), status: String(x.trang_thai || "active"), createdAt: String(x.tao_luc || "") }))
      .sort((a: any, b: any) => a.name.localeCompare(b.name, "vi"));
    return ok({ data: list });
  } catch (e) { return fail((e as Error).message); }
}
export async function addLuongCungStaff(token: string, name: string, role: string, note: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    name = String(name || "").replace(/\s+/g, " ").trim(); if (!name) return fail("Thiếu tên.");
    if ((await sql`select 1 from chamcong.nhan_su_luong_cung where lower(trim(ten))=${L(name)}`).length) return fail(`Nhân sự "${name}" đã tồn tại.`);
    await sql`insert into chamcong.nhan_su_luong_cung(ten,vai_tro,ghi_chu,trang_thai,tao_luc) values(${name},${String(role || "")},${String(note || "")},'active',${nowString()})`;
    return ok();
  } catch (e) { return fail((e as Error).message); }
}
export async function updateLuongCungStaff(token: string, name: string, role: string | null, note: string | null, status: string | null) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    name = String(name || "").trim(); if (!name) return fail("Thiếu tên.");
    const r = await sql`select ten from chamcong.nhan_su_luong_cung where lower(trim(ten))=${L(name)}`;
    if (!r.length) return fail(`Không tìm thấy "${name}".`);
    if (role != null) await sql`update chamcong.nhan_su_luong_cung set vai_tro=${String(role)} where lower(trim(ten))=${L(name)}`;
    if (note != null) await sql`update chamcong.nhan_su_luong_cung set ghi_chu=${String(note)} where lower(trim(ten))=${L(name)}`;
    if (status != null) await sql`update chamcong.nhan_su_luong_cung set trang_thai=${String(status)} where lower(trim(ten))=${L(name)}`;
    return ok();
  } catch (e) { return fail((e as Error).message); }
}
export async function deleteLuongCungStaff(token: string, name: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    name = String(name || "").trim(); if (!name) return fail("Thiếu tên.");
    await sql.begin(async (tx: any) => {
      await tx`delete from chamcong.nhan_su_luong_cung where lower(trim(ten))=${L(name)}`;
      await tx`delete from chamcong.luong_cung_thang where lower(trim(ten))=${L(name)}`;
    });
    return ok();
  } catch (e) { return fail((e as Error).message); }
}
export async function getLuongCungByMonth(token: string, month: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ.");
  try {
    const r = await sql`select ten, kenh, so_tien, ghi_chu from chamcong.luong_cung_thang where thang=${month} order by ctid`;
    return ok({ data: r.filter((x: any) => x.ten).map((x: any) => ({ name: String(x.ten).trim(), channel: String(x.kenh || "").trim(), amount: num(x.so_tien), note: String(x.ghi_chu || "") })) });
  } catch (e) { return fail((e as Error).message); }
}
export async function saveLuongCungBatch(token: string, month: string, items: any[]) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ.");
    if ((await readLocks())[month]) return fail(`Tháng ${month} đã chốt lương, không sửa được lương cứng.`);
    items = Array.isArray(items) ? items : [];
    const rows: any[] = [];
    items.forEach((it) => {
      const nm = String(it.name || "").trim(), ch = String(it.channel || "").trim(), amt = num(it.amount), note = String(it.note || "");
      if (!nm || !ch) return; if (amt === 0 && !note) return;
      rows.push([nm, ch, amt, note]);
    });
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      await tx`delete from chamcong.luong_cung_thang where thang=${month}`;
      for (const r of rows) await tx`insert into chamcong.luong_cung_thang(thang,ten,kenh,so_tien,ghi_chu) values(${month},${r[0]},${r[1]},${r[2]},${r[3]})`;
    });
    return { success: true, saved: rows.length };
  } catch (e) { return fail("Lỗi lưu lương cứng: " + (e as Error).message); }
}

/* ---------- thử việc ---------- */
export async function getProbationByMonth(token: string, month: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ.");
  try {
    const r = await sql`select ten, kenh, so_tien, ghi_chu from chamcong.luong_thu_viec where thang=${month} order by ctid`;
    return ok({ data: r.filter((x: any) => x.ten).map((x: any) => ({ name: String(x.ten).replace(/\s+/g, " ").trim(), channel: String(x.kenh || "").trim(), amount: num(x.so_tien), note: String(x.ghi_chu || "") })) });
  } catch (e) { return fail((e as Error).message); }
}
export async function saveProbationBatch(token: string, month: string, items: any[]) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ.");
    if ((await readLocks())[month]) return fail(`Tháng ${month} đã chốt lương. Mở chốt trước khi sửa lương thử việc.`);
    items = Array.isArray(items) ? items : [];
    const rows: any[] = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i] || {};
      const nm = String(it.name || "").replace(/\s+/g, " ").trim(), ch = String(it.channel || "").trim(), amt = num(it.amount), note = String(it.note || "");
      if (!nm && amt === 0 && !note) continue;
      if (!nm) return fail(`Dòng thứ ${i + 1} thiếu tên nhân sự.`);
      if (amt === 0 && !note) continue;
      if (ch && !(await channelExists(ch))) return fail(`Kênh "${ch}" không có trong danh sách kênh.`);
      rows.push([nm, ch, amt, note]);
    }
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      await tx`delete from chamcong.luong_thu_viec where thang=${month}`;
      for (const r of rows) await tx`insert into chamcong.luong_thu_viec(thang,ten,kenh,so_tien,ghi_chu) values(${month},${r[0]},${r[1]},${r[2]},${r[3]})`;
    });
    return { success: true, saved: rows.length };
  } catch (e) { return fail("Lỗi lưu lương thử việc: " + (e as Error).message); }
}

/* ---------- ngày lễ ---------- */
export async function getHolidays(token: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try { return ok({ data: await readHolidays() }); } catch (e) { return fail((e as Error).message); }
}
async function readHolidays() {
  const r = await sql`select ngay, he_so, ghi_chu from chamcong.ngay_le`;
  return r.map((x: any) => ({ date: String(x.ngay).slice(0, 10), mult: num(x.he_so), note: String(x.ghi_chu || "") }))
    .filter((h: any) => /^\d{4}-\d{2}-\d{2}$/.test(h.date) && h.mult > 0).sort((a: any, b: any) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}
export async function saveHolidays(token: string, items: any[]) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    items = Array.isArray(items) ? items : [];
    const seen: Record<string, boolean> = {}, rows: { date: string; mult: number; note: string }[] = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i] || {}, d = String(it.date || "").trim(), m = Number(it.mult);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return fail(`Dòng ${i + 1}: ngày không hợp lệ.`);
      if (!(m > 0) || m > 10) return fail(`Ngày ${d}: hệ số phải lớn hơn 0 và không quá 10.`);
      if (seen[d]) return fail(`Ngày ${d} bị nhập trùng.`);
      seen[d] = true; rows.push({ date: d, mult: m, note: String(it.note || "") });
    }
    const canon = (list: any[], mo: string) => list.filter((h) => h.date.slice(0, 7) === mo).map((h) => h.date + "|" + h.mult).sort().join(";");
    const old = await readHolidays(), locks = await readLocks();
    const months: Record<string, boolean> = {}; old.concat(rows as any).forEach((h: any) => (months[h.date.slice(0, 7)] = true));
    for (const mo of Object.keys(months)) if (canon(old, mo) !== canon(rows, mo) && locks[mo]) return fail(`Tháng ${mo} đã chốt lương — Mở chốt trước khi đổi ngày lễ của tháng đó.`);
    rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      await tx`delete from chamcong.ngay_le`;
      for (const h of rows) await tx`insert into chamcong.ngay_le(ngay,he_so,ghi_chu) values(${h.date},${h.mult},${h.note})`;
    });
    return { success: true, saved: rows.length };
  } catch (e) { return fail("Lỗi lưu ngày lễ: " + (e as Error).message); }
}

/* ---------- Mã NV & gộp nhân sự ---------- */
export async function updateStaffCode(token: string, name: string, newCode: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    name = String(name || "").trim(); newCode = String(newCode || "").trim();
    if (!name) return fail("Thiếu tên nhân sự.");
    const all = await sql`select ten, coalesce(ma_nv,'') ma from chamcong.nhan_su`;
    const me = all.find((x: any) => L(x.ten) === L(name));
    if (!me) return fail("Không tìm thấy nhân sự.");
    if (newCode) {
      const dup = all.find((x: any) => L(x.ten) !== L(name) && String(x.ma || "").trim() && L(x.ma) === L(newCode));
      if (dup) return { success: true, duplicate: { otherName: String(dup.ten).trim(), otherCode: String(dup.ma).trim() } };
    }
    await sql`update chamcong.nhan_su set ma_nv=${newCode} where lower(trim(ten))=${L(name)}`;
    return ok();
  } catch (e) { return fail("Lỗi cập nhật mã: " + (e as Error).message); }
}

export async function mergeStaff(token: string, oldName: string, newName: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  try {
    oldName = String(oldName || "").trim(); newName = String(newName || "").trim();
    if (!oldName || !newName) return fail("Thiếu tên cũ hoặc tên mới.");
    if (L(oldName) === L(newName)) return fail("Tên cũ và tên mới trùng nhau.");
    const oL = L(oldName), nL = L(newName);
    const counts = { shifts: 0, hourly: 0, moc: 0, diem: 0, users: 0, hoTro: 0, luongCung: 0, salary: 0 };
    await sql.begin(async (tx: any) => {
      await lockTx(tx);
      counts.shifts = (await tx`update chamcong.ca_live set ten=${newName} where lower(trim(ten))=${oL} returning 1`).length;
      // xoá dòng cũ trùng khoá với tên mới, còn lại đổi tên
      counts.hourly += (await tx`delete from chamcong.luong_gio s where lower(trim(s.ten))=${oL} and exists (select 1 from chamcong.luong_gio t where lower(trim(t.ten))=${nL} and lower(trim(coalesce(t.kenh,'')))=lower(trim(coalesce(s.kenh,'')))) returning 1`).length;
      counts.hourly += (await tx`update chamcong.luong_gio set ten=${newName} where lower(trim(ten))=${oL} returning 1`).length;
      counts.moc += (await tx`delete from chamcong.thuong_moc s where lower(trim(s.ten))=${oL} and exists (select 1 from chamcong.thuong_moc t where lower(trim(t.ten))=${nL} and lower(trim(t.kenh))=lower(trim(s.kenh)) and t.vai_tro=s.vai_tro) returning 1`).length;
      counts.moc += (await tx`update chamcong.thuong_moc set ten=${newName} where lower(trim(ten))=${oL} returning 1`).length;
      counts.diem += (await tx`delete from chamcong.diem_thang s where lower(trim(s.ten))=${oL} and exists (select 1 from chamcong.diem_thang t where lower(trim(t.ten))=${nL} and t.thang=s.thang) returning 1`).length;
      counts.diem += (await tx`update chamcong.diem_thang set ten=${newName} where lower(trim(ten))=${oL} returning 1`).length;
      counts.hoTro = (await tx`update chamcong.thuong_ho_tro set ten=${newName} where lower(trim(ten))=${oL} returning 1`).length;
      counts.luongCung += (await tx`delete from chamcong.luong_cung_thang s where lower(trim(s.ten))=${oL} and exists (select 1 from chamcong.luong_cung_thang t where lower(trim(t.ten))=${nL} and t.thang=s.thang and lower(trim(coalesce(t.kenh,'')))=lower(trim(coalesce(s.kenh,'')))) returning 1`).length;
      counts.luongCung += (await tx`update chamcong.luong_cung_thang set ten=${newName} where lower(trim(ten))=${oL} returning 1`).length;
      const lc = await tx`select ten from chamcong.nhan_su_luong_cung where lower(trim(ten)) in (${oL},${nL})`;
      const hasOld = lc.some((x: any) => L(x.ten) === oL), hasNew = lc.some((x: any) => L(x.ten) === nL);
      if (hasOld) { if (hasNew) await tx`delete from chamcong.nhan_su_luong_cung where lower(trim(ten))=${oL}`; else await tx`update chamcong.nhan_su_luong_cung set ten=${newName} where lower(trim(ten))=${oL}`; }
      counts.salary += (await tx`delete from chamcong.luong_cu s where lower(trim(s.ten))=${oL} and exists (select 1 from chamcong.luong_cu t where lower(trim(t.ten))=${nL} and lower(trim(coalesce(t.vi_tri,'')))=lower(trim(coalesce(s.vi_tri,'')))) returning 1`).length;
      counts.salary += (await tx`update chamcong.luong_cu set ten=${newName} where lower(trim(ten))=${oL} returning 1`).length;
      counts.users = (await tx`update chamcong.users set nhan_su=${newName} where lower(trim(nhan_su))=${oL} returning 1`).length;
      await tx`delete from chamcong.nhan_su where lower(trim(ten))=${oL}`;
      const has = await tx`select 1 from chamcong.nhan_su where lower(trim(ten))=${nL}`;
      if (!has.length) await tx`insert into chamcong.nhan_su(ten,tao_luc,ma_nv) values(${newName},${nowString()},'')`;
    });
    return { success: true, counts };
  } catch (e) { return fail("Lỗi gộp nhân sự: " + (e as Error).message); }
}
