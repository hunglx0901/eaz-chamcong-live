import { sql, Auth, ok, fail, authErr, denyErr, verifyToken, makeToken, hashPw, splitChannels, leadChannelsLower, leadHasChannel,
  nowString, todayStr, timeStr, asTimeStr, computeHours, roundInTime, roundOutTime, esc, fmtVnd, sendTelegram, savePhoto, readPhoto, deletePhoto } from "./lib.ts";

const L = (s: unknown) => String(s ?? "").trim().toLowerCase();
const num = (v: unknown) => Number(v) || 0;

/* ======================= ĐĂNG NHẬP / TÀI KHOẢN ======================= */
async function syncStaffChannels(username: string, staffName: string, cur: string[]): Promise<string[] | null> {
  try {
    const nm = L(staffName); if (!nm) return null;
    const rows = await sql`select distinct kenh from chamcong.ca_live where lower(trim(ten))=${nm} and kenh<>''`;
    const lower: Record<string, boolean> = {}; cur.forEach((c) => (lower[c.toLowerCase()] = true));
    let changed = false; const out = [...cur];
    rows.forEach((r) => { const ch = String(r.kenh).trim(); if (ch && !lower[ch.toLowerCase()]) { out.push(ch); changed = true; } });
    if (changed) await sql`update chamcong.users set kenh=${out.join("|")} where lower(trim(username))=${L(username)}`;
    return out;
  } catch { return null; }
}

export async function login(username: string, password: string) {
  try {
    const r = await sql`select * from chamcong.users where lower(trim(username))=${L(username)}`;
    if (!r.length) return fail("Tài khoản không tồn tại.");
    const u = r[0];
    if ((await hashPw(password, String(u.salt))) !== String(u.hash)) return fail("Sai mật khẩu.");
    const staffName = String(u.nhan_su || "");
    let chans = splitChannels(u.kenh);
    if (u.vai_tro === "staff" && staffName) { const s = await syncStaffChannels(u.username, staffName, chans); if (s) chans = s; }
    const managed = splitChannels(u.kenh_quan_ly);
    const firstLogin = u.vai_tro === "staff" && !String(u.da_thiet_lap || "").trim();
    if (firstLogin) await sql`update chamcong.users set da_thiet_lap='x' where username=${u.username}`;
    const id: Auth = { username: String(u.username), role: String(u.vai_tro), channel: chans[0] || "", channels: chans, staff: staffName, managed };
    return { success: true, token: await makeToken(id), role: id.role, channel: id.channel, channels: id.channels, staff: id.staff, username: id.username, firstLogin, managed };
  } catch (e) { return fail("Lỗi đăng nhập: " + (e as Error).message); }
}

export async function whoami(token: string) {
  const id = await verifyToken(token);
  if (!id) return { success: false, code: "AUTH" };
  return { success: true, role: id.role, channel: id.channel, channels: id.channels, staff: id.staff, username: id.username, managed: id.managed };
}

export async function changeMyPassword(token: string, oldPw: string, newPw: string) {
  const id = await verifyToken(token); if (!id) return authErr();
  if (!newPw || String(newPw).length < 4) return fail("Mật khẩu mới tối thiểu 4 ký tự.");
  const r = await sql`select hash,salt from chamcong.users where lower(trim(username))=${L(id.username)}`;
  if (!r.length) return authErr();
  if ((await hashPw(oldPw, r[0].salt)) !== r[0].hash) return fail("Mật khẩu hiện tại không đúng.");
  const salt = crypto.randomUUID();
  await sql`update chamcong.users set hash=${await hashPw(newPw, salt)}, salt=${salt} where lower(trim(username))=${L(id.username)}`;
  return ok();
}

export async function leadStaffSet(auth: Auth): Promise<Record<string, boolean>> {
  const set: Record<string, boolean> = {};
  const chans = leadChannelsLower(auth);
  if (!chans.length) return set;
  const a = await sql`select distinct lower(trim(ten)) n from chamcong.ca_live where lower(kenh) = any(${chans})`;
  a.forEach((r) => (set[r.n] = true));
  const u = await sql`select nhan_su, kenh from chamcong.users where vai_tro='staff'`;
  u.forEach((r) => {
    const n = L(r.nhan_su); if (!n) return;
    if (splitChannels(r.kenh).some((c) => chans.includes(c.toLowerCase()))) set[n] = true;
  });
  return set;
}

export async function listUsers(token: string) {
  const id = await verifyToken(token); if (!id) return authErr();
  if (id.role !== "owner" && id.role !== "lead") return denyErr();
  try {
    const r = await sql`select username,vai_tro,kenh,nhan_su,kenh_quan_ly from chamcong.users order by tao_luc`;
    let data = r.map((x) => ({ username: String(x.username), role: String(x.vai_tro), channel: String(x.kenh || ""), staff: String(x.nhan_su || ""), managed: String(x.kenh_quan_ly || "") }));
    if (id.role === "lead") { const set = await leadStaffSet(id); data = data.filter((u) => u.role === "staff" && set[L(u.staff)]); }
    return ok({ data });
  } catch (e) { return fail("Lỗi đọc tài khoản: " + (e as Error).message); }
}

async function ensureStaff(name: string) {
  name = String(name || "").trim(); if (!name) return;
  try {
    const e = await sql`select 1 from chamcong.nhan_su where lower(trim(ten))=${L(name)}`;
    if (!e.length) await sql`insert into chamcong.nhan_su(ten,tao_luc,ma_nv) values(${name},${nowString()},'') on conflict do nothing`;
  } catch { /* bỏ qua */ }
}

export async function createUser(token: string, username: string, password: string, role: string, channel: string, staff: string, managed: string) {
  const id = await verifyToken(token); if (!id) return authErr();
  if (id.role !== "owner" && id.role !== "lead") return denyErr();
  username = String(username || "").trim(); role = String(role || "").trim();
  if (id.role === "lead" && role !== "staff") return denyErr();
  if (!username) return fail("Thiếu tên đăng nhập.");
  if (!password || String(password).length < 4) return fail("Mật khẩu tối thiểu 4 ký tự.");
  if (!["owner", "admin", "lead", "staff"].includes(role)) return fail("Vai trò không hợp lệ.");
  const chans = splitChannels(channel);
  if (role === "lead" && !chans.length) return fail("Lead cần gán ít nhất 1 kênh.");
  if (role === "staff" && !String(staff || "").trim()) return fail("Nhân viên cần gán 1 nhân sự.");
  const managedChans = role === "staff" ? splitChannels(managed) : [];
  const ex = await sql`select 1 from chamcong.users where lower(trim(username))=${L(username)}`;
  if (ex.length) return fail("Tên đăng nhập đã tồn tại.");
  const salt = crypto.randomUUID();
  if (role === "staff") await ensureStaff(String(staff).trim());
  await sql`insert into chamcong.users(username,hash,salt,vai_tro,kenh,nhan_su,tao_luc,da_thiet_lap,kenh_quan_ly)
    values(${username},${await hashPw(password, salt)},${salt},${role},${role === "lead" || role === "staff" ? chans.join("|") : ""},${role === "staff" ? String(staff).trim() : ""},${nowString()},'',${managedChans.join("|")})`;
  return ok();
}

export async function updateUser(token: string, username: string, role: string, channel: string, staff: string, managed: string) {
  const id = await verifyToken(token); if (!id) return authErr();
  if (id.role !== "owner" && id.role !== "lead") return denyErr();
  const r = await sql`select vai_tro,nhan_su from chamcong.users where lower(trim(username))=${L(username)}`;
  if (!r.length) return fail("Không tìm thấy tài khoản.");
  if (id.role === "lead") {
    if (r[0].vai_tro !== "staff") return denyErr();
    const set = await leadStaffSet(id); if (!set[L(r[0].nhan_su)]) return denyErr();
    await sql`update chamcong.users set kenh=${splitChannels(channel).join("|")} where lower(trim(username))=${L(username)}`;
    return ok();
  }
  role = String(role || "").trim();
  if (!["owner", "admin", "lead", "staff"].includes(role)) return fail("Vai trò không hợp lệ.");
  const chans = splitChannels(channel);
  if (role === "lead" && !chans.length) return fail("Lead cần gán ít nhất 1 kênh.");
  if (role === "staff" && !String(staff || "").trim()) return fail("Nhân viên cần gán 1 nhân sự.");
  await sql`update chamcong.users set vai_tro=${role}, kenh=${role === "lead" || role === "staff" ? chans.join("|") : ""},
    nhan_su=${role === "staff" ? String(staff).trim() : ""}, kenh_quan_ly=${role === "staff" ? splitChannels(managed).join("|") : ""}
    where lower(trim(username))=${L(username)}`;
  return ok();
}

export async function resetPassword(token: string, username: string, newPw: string) {
  const id = await verifyToken(token); if (!id) return authErr();
  if (id.role !== "owner") return denyErr();
  if (!newPw || String(newPw).length < 4) return fail("Mật khẩu tối thiểu 4 ký tự.");
  const salt = crypto.randomUUID();
  const r = await sql`update chamcong.users set hash=${await hashPw(newPw, salt)}, salt=${salt} where lower(trim(username))=${L(username)} returning 1`;
  return r.length ? ok() : fail("Không tìm thấy tài khoản.");
}

export async function deleteUser(token: string, username: string) {
  const id = await verifyToken(token); if (!id) return authErr();
  if (id.role !== "owner") return denyErr();
  username = String(username || "").trim();
  if (L(username) === L(id.username)) return fail("Không thể xoá tài khoản đang đăng nhập.");
  const r = await sql`select vai_tro from chamcong.users where lower(trim(username))=${L(username)}`;
  if (!r.length) return fail("Không tìm thấy tài khoản.");
  const owners = await sql`select count(*)::int n from chamcong.users where vai_tro='owner'`;
  if (r[0].vai_tro === "owner" && owners[0].n <= 1) return fail("Phải còn ít nhất 1 Owner.");
  await sql`delete from chamcong.users where lower(trim(username))=${L(username)}`;
  return ok();
}

export async function setMyChannels(token: string, channels: unknown) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "staff") return fail("Chỉ nhân viên cập nhật kênh của mình.");
  let chans = Array.isArray(channels) ? channels as string[] : splitChannels(channels);
  chans = chans.map((c) => String(c).trim()).filter(Boolean);
  const r = await sql`update chamcong.users set kenh=${chans.join("|")} where lower(trim(username))=${L(auth.username)} returning 1`;
  return r.length ? ok({ channels: chans }) : fail("Không tìm thấy tài khoản.");
}

/* ======================= NHÂN SỰ & KÊNH ======================= */
export async function getStaff(token: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    const r = await sql`select ten, coalesce(ma_nv,'') ma from chamcong.nhan_su`;
    const seen: Record<string, boolean> = {};
    let full = r.map((x) => ({ name: String(x.ten).trim(), code: String(x.ma || "").trim() }))
      .filter((x) => { if (!x.name || seen[L(x.name)]) return false; seen[L(x.name)] = true; return true; })
      .sort((a, b) => a.name.localeCompare(b.name, "vi"));
    if (auth.role === "staff") full = full.filter((x) => L(x.name) === L(auth.staff));
    return ok({ data: full.map((x) => x.name), full });
  } catch (e) { return fail("Lỗi đọc nhân sự: " + (e as Error).message); }
}

export async function addStaff(token: string, name: string, code: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner" && auth.role !== "lead") return denyErr();
  name = String(name || "").trim(); code = String(code || "").trim();
  if (!name) return fail("Tên nhân sự không được để trống.");
  const ex = async (n: string) => (await sql`select 1 from chamcong.nhan_su where lower(trim(ten))=${L(n)}`).length > 0;
  if (await ex(name)) {
    if (!code) return fail(`Tên "${name}" đã tồn tại. Nhập Mã NV để phân biệt (hệ thống sẽ lưu thành "${name} (Mã)").`);
    name = `${name} (${code})`;
    if (await ex(name)) return fail(`Nhân sự "${name}" đã tồn tại.`);
  }
  await sql`insert into chamcong.nhan_su(ten,tao_luc,ma_nv) values(${name},${nowString()},${code})`;
  return ok({ name });
}

export async function deleteStaff(token: string, name: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  const r = await sql`delete from chamcong.nhan_su where lower(trim(ten))=${L(name)} returning 1`;
  return r.length ? ok() : fail("Không tìm thấy nhân sự.");
}

export async function getChannels(token: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    const r = await sql`select ten from chamcong.kenh`;
    let names: string[] = ([...new Set(r.map((x: any) => String(x.ten).trim()).filter(Boolean))] as string[]).sort((a, b) => a.localeCompare(b, "vi"));
    if (auth.role === "lead") { const s = leadChannelsLower(auth); names = names.filter((n) => s.includes(n.toLowerCase())); }
    return ok({ data: names });
  } catch (e) { return fail("Lỗi đọc kênh: " + (e as Error).message); }
}

export async function addChannel(token: string, name: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  name = String(name || "").trim();
  if (!name) return fail("Tên kênh không được để trống.");
  if ((await sql`select 1 from chamcong.kenh where lower(trim(ten))=${L(name)}`).length) return fail(`Kênh "${name}" đã tồn tại.`);
  await sql`insert into chamcong.kenh(ten,tao_luc) values(${name},${nowString()})`;
  return ok();
}

export async function deleteChannel(token: string, name: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner") return denyErr();
  const r = await sql`delete from chamcong.kenh where lower(trim(ten))=${L(name)} returning 1`;
  return r.length ? ok() : fail("Không tìm thấy kênh.");
}

/* ======================= CA LIVE ======================= */
export async function getShifts(token: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  try {
    const rows = await sql`select * from chamcong.ca_live order by ngay, bat_dau`;
    let data = rows.map((r) => {
      const start = asTimeStr(r.bat_dau), end = asTimeStr(r.ket_thuc);
      const closed = !!String(r.ket_thuc ?? "").trim();
      return {
        id: String(r.id), date: r.ngay ? new Date(r.ngay).toISOString().slice(0, 10) : "", name: String(r.ten), position: String(r.vi_tri ?? ""), channel: String(r.kenh ?? ""),
        start, end, hours: closed ? computeHours(start, end) : num(r.gio_live), gmv: num(r.gmv), note: String(r.ghi_chu || ""), createdAt: String(r.tao_luc || ""),
        gmvStart: num(r.gmv_dau), gmvEnd: num(r.gmv_cuoi), open: !closed, hasStartPhoto: !!String(r.anh_dau || "").trim(), hasEndPhoto: !!String(r.anh_cuoi || "").trim(),
      };
    });
    if (auth.role === "lead") { const s = leadChannelsLower(auth); data = data.filter((x) => s.includes(x.channel.toLowerCase())); }
    else if (auth.role === "staff") {
      const nm = L(auth.staff), m = auth.managed.map((c) => c.toLowerCase());
      data = data.filter((x) => L(x.name) === nm || m.includes(x.channel.toLowerCase()));
    }
    return ok({ data });
  } catch (e) { return fail("Lỗi đọc dữ liệu: " + (e as Error).message); }
}

function validate(rec: any): string {
  if (!rec) return "Dữ liệu rỗng.";
  if (!rec.date) return "Thiếu ngày live.";
  if (!rec.name || !String(rec.name).trim()) return "Thiếu tên nhân sự.";
  if (!rec.position) return "Thiếu vị trí.";
  if (!rec.channel) return "Thiếu kênh.";
  if (!rec.start || !rec.end) return "Thiếu giờ bắt đầu hoặc kết thúc.";
  if (num(rec.gmvEnd) < num(rec.gmvStart)) return "GMV cuối ca phải ≥ GMV đầu ca.";
  return "";
}

export async function addShift(token: string, rec: any) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role === "admin") return denyErr();
  rec = rec || {};
  if (auth.role === "lead" && !leadHasChannel(auth, rec.channel)) return denyErr();
  if (auth.role === "staff") rec.name = auth.staff;
  const err = validate(rec); if (err) return fail(err);
  const id = crypto.randomUUID();
  const hours = computeHours(rec.start, rec.end), gs = num(rec.gmvStart), ge = num(rec.gmvEnd);
  await sql`insert into chamcong.ca_live(id,ngay,ten,vi_tri,kenh,bat_dau,ket_thuc,gio_live,gmv,ghi_chu,tao_luc,gmv_dau,gmv_cuoi,anh_dau,anh_cuoi)
    values(${id},${rec.date},${String(rec.name).trim()},${rec.position},${rec.channel},${asTimeStr(rec.start)},${asTimeStr(rec.end)},${hours},${ge - gs},${String(rec.note || "").trim()},${nowString()},${gs},${ge},'','')`;
  await ensureStaff(rec.name);
  return ok({ id });
}

export async function updateShift(token: string, id: string, rec: any) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role === "admin" || auth.role === "staff") return denyErr();
  if (auth.role === "lead" && !leadHasChannel(auth, rec?.channel)) return denyErr();
  if (!id) return fail("Thiếu ID bản ghi.");
  const err = validate(rec); if (err) return fail(err);
  const cur = await sql`select kenh from chamcong.ca_live where id=${id}`;
  if (!cur.length) return fail("Không tìm thấy bản ghi để sửa.");
  if (auth.role === "lead" && !leadHasChannel(auth, cur[0].kenh)) return denyErr();
  const hours = computeHours(rec.start, rec.end), gs = num(rec.gmvStart), ge = num(rec.gmvEnd);
  await sql`update chamcong.ca_live set ngay=${rec.date}, ten=${String(rec.name).trim()}, vi_tri=${rec.position}, kenh=${rec.channel},
    bat_dau=${asTimeStr(rec.start)}, ket_thuc=${asTimeStr(rec.end)}, gio_live=${hours}, gmv=${ge - gs}, ghi_chu=${String(rec.note || "").trim()},
    gmv_dau=${gs}, gmv_cuoi=${ge} where id=${id}`;
  await ensureStaff(rec.name);
  return ok();
}

export async function deleteShift(token: string, id: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role === "admin" || auth.role === "staff") return denyErr();
  if (!id) return fail("Thiếu ID bản ghi.");
  const cur = await sql`select kenh,anh_dau,anh_cuoi from chamcong.ca_live where id=${id}`;
  if (!cur.length) return fail("Không tìm thấy bản ghi để xoá.");
  if (auth.role === "lead" && !leadHasChannel(auth, cur[0].kenh)) return denyErr();
  await deletePhoto(String(cur[0].anh_dau || "")); await deletePhoto(String(cur[0].anh_cuoi || ""));
  await sql`delete from chamcong.ca_live where id=${id}`;
  return ok();
}

export async function getShiftPhoto(token: string, shiftId: string, which: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  const r = await sql`select ten,kenh,anh_dau,anh_cuoi from chamcong.ca_live where id=${shiftId}`;
  if (!r.length) return fail("Không tìm thấy ca.");
  const { ten, kenh } = r[0];
  if (auth.role === "lead" && !leadHasChannel(auth, kenh)) return denyErr();
  if (auth.role === "staff") {
    const mine = L(ten) === L(auth.staff), inManaged = auth.managed.some((c) => c.toLowerCase() === String(kenh).toLowerCase());
    if (!mine && !inManaged) return denyErr();
  }
  const ref = String(which === "end" ? r[0].anh_cuoi : r[0].anh_dau || "");
  if (!ref) return fail("Chưa có ảnh.");
  if (!ref.includes("/")) return fail("Ảnh này được chụp trước khi chuyển hệ thống (lưu trên Google Drive) nên không xem được ở bản mới.");
  const data = await readPhoto(ref);
  return data ? ok({ data }) : fail("Ảnh đã hết hạn lưu trữ.");
}

export async function checkIn(token: string, rec: any, photoB64: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "staff") return fail("Chức năng chấm công chỉ dành cho Nhân viên.");
  rec = rec || {}; rec.name = auth.staff;
  if (!photoB64) return fail("Cần chụp ảnh đầu ca.");
  if (!String(rec.name || "").trim()) return fail("Thiếu nhân sự.");
  if (!rec.channel) return fail("Vui lòng chọn kênh.");
  if (!rec.position) return fail("Vui lòng chọn vị trí.");
  const open = await sql`select 1 from chamcong.ca_live where lower(ten)=${L(rec.name)} and trim(coalesce(ket_thuc,''))=''`;
  if (open.length) return fail("Đang có một ca chưa check-out. Hãy check-out trước.");
  const photo = await savePhoto(photoB64, "in");
  const id = crypto.randomUUID(), gs = num(rec.gmvStart), inTime = roundInTime(timeStr());
  await sql`insert into chamcong.ca_live(id,ngay,ten,vi_tri,kenh,bat_dau,ket_thuc,gio_live,gmv,ghi_chu,tao_luc,gmv_dau,gmv_cuoi,anh_dau,anh_cuoi)
    values(${id},${todayStr()},${String(rec.name).trim()},${rec.position},${rec.channel},${inTime},'',0,0,${String(rec.note || "").trim()},${nowString()},${gs},0,${photo},'')`;
  await ensureStaff(rec.name);
  // tự thêm kênh vào tài khoản
  let newChans: string[] | null = null;
  try {
    const u = await sql`select kenh from chamcong.users where lower(trim(username))=${L(auth.username)}`;
    if (u.length) {
      const cur = splitChannels(u[0].kenh);
      if (!cur.map((c) => c.toLowerCase()).includes(String(rec.channel).toLowerCase())) {
        cur.push(String(rec.channel).trim()); await sql`update chamcong.users set kenh=${cur.join("|")} where lower(trim(username))=${L(auth.username)}`;
      }
      newChans = cur;
    }
  } catch { /* bỏ qua */ }
  try {
    await sendTelegram(`✅ <b>CHECK-IN</b>\n👤 ${esc(String(rec.name).trim())} · ${esc(rec.position)}\n📺 ${esc(rec.channel)}\n🕐 ${inTime}`);
    if (rec.position === "Host Live") {
      const o = await sql`select ten,kenh from chamcong.ca_live where vi_tri='Host Live' and trim(coalesce(ket_thuc,''))='' and id<>${id}`;
      const msgs: string[] = []; const nm = L(rec.name), ch = L(rec.channel);
      const sc = o.find((r) => L(r.kenh) === ch && L(r.ten) !== nm); const sp = o.find((r) => L(r.ten) === nm && L(r.kenh) !== ch);
      if (sc) msgs.push(`Kênh <b>${esc(rec.channel)}</b> đang có 2 Host Live cùng lúc: ${esc(sc.ten)} &amp; ${esc(rec.name)}.`);
      if (sp) msgs.push(`<b>${esc(rec.name)}</b> đang Host Live ở 2 kênh cùng lúc: ${esc(sp.kenh)} &amp; ${esc(rec.channel)}.`);
      if (msgs.length) await sendTelegram("🔴 <b>CẢNH BÁO TRÙNG HOST LIVE</b>\n" + msgs.join("\n"));
    }
  } catch { /* bỏ qua */ }
  return ok({ id, channels: newChans });
}

export async function checkOut(token: string, shiftId: string, gmvEnd: number, photoB64: string) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "staff") return fail("Chức năng chấm công chỉ dành cho Nhân viên.");
  if (!photoB64) return fail("Cần chụp ảnh cuối ca.");
  const r = await sql`select * from chamcong.ca_live where id=${shiftId}`;
  if (!r.length) return fail("Không tìm thấy ca.");
  const s = r[0];
  if (L(s.ten) !== L(auth.staff)) return denyErr();
  if (String(s.ket_thuc || "").trim()) return fail("Ca này đã check-out rồi.");
  const start = asTimeStr(s.bat_dau), gs = num(s.gmv_dau), ge = num(gmvEnd);
  if (ge < gs) return fail("GMV cuối ca phải ≥ GMV đầu ca.");
  let end = roundOutTime(timeStr()); if (start === end) end = timeStr();
  const photo = await savePhoto(photoB64, "out");
  await sql`update chamcong.ca_live set ket_thuc=${end}, gio_live=${computeHours(start, end)}, gmv=${ge - gs}, gmv_cuoi=${ge}, anh_cuoi=${photo} where id=${shiftId}`;
  try { await sendTelegram(`🏁 <b>CHECK-OUT</b>\n👤 ${esc(s.ten)}\n📺 ${esc(s.kenh)}\n🕐 ${esc(start)} – ${esc(end)}\n💰 GMV ca: ${fmtVnd(ge - gs)}đ`); } catch { /* bỏ qua */ }
  return ok();
}

export const getDatabaseUrl = async (_t: string) => ({ success: false });
