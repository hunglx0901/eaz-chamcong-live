// Engine tính lương — HÀM THUẦN (không đụng DB), port nguyên văn từ _computePayrollRaw_ của Apps Script.
export type Auth = { username: string; role: string; channel: string; channels: string[]; staff: string; managed: string[] };
const L = (s: unknown) => String(s ?? "").trim().toLowerCase();
const num = (v: unknown) => Number(v) || 0;

export const THUONGCHUNG_DEFAULT: Record<string, number> = {
  heso_gmvh: 0.4, heso_nmv: 0.4, heso_diem: 0.2, tran_gmvh: 0.48, tran_nmv: 0.48, tran_diem: 0.2,
  moc_diem: 4, vat: 1.08, pct_host: 0.5, pct_tro: 0.2, hoan_huy: 30,
};
export function buildCfg(rows: { khoa: string; gia_tri: string }[]) {
  const cfg: Record<string, number> = {};
  rows.forEach((r) => { if (r.khoa) cfg[String(r.khoa).trim()] = Number(r.gia_tri) || 0; });
  Object.keys(THUONGCHUNG_DEFAULT).forEach((k) => { if (cfg[k] === undefined) cfg[k] = THUONGCHUNG_DEFAULT[k]; });
  return cfg;
}

export type PayInput = {
  cfg: Record<string, number>;
  hourly: { ten: string; kenh: string; don_gia_host: unknown; don_gia_tro: unknown }[];
  moc: { ten: string; kenh: string; vai_tro: string; moc_gmvh: unknown; moc_nmv: unknown }[];
  hoanHuy: { thang: string; ty_le: unknown }[];
  diem: { thang: string; ten: string; diem: unknown }[];
  hoTro: { thang: string; ten: string; kenh: string; so_tien: unknown; ghi_chu: string }[];
  luongCung: { thang: string; ten: string; kenh: string; so_tien: unknown; ghi_chu: string }[];
  probation: { thang: string; ten: string; kenh: string; so_tien: unknown; ghi_chu: string }[];
  holidays: { ngay: string; he_so: unknown }[];
  shifts: { ngay: string; ten: string; vi_tri: string; kenh: string; ket_thuc: string; bat_dau: string; gio_live: unknown; gmv: unknown }[];
  prevSnapshotForDiem?: { name: string; diem: unknown }[] | null; // điểm snapshot tháng trước (nếu tháng này chưa chốt & tháng trước đã chốt)
  leadStaff?: Record<string, boolean>; // bộ tên nhân sự thuộc phạm vi Lead
};

function asTimeStr(v: unknown) {
  const s = String(v ?? "").trim(); if (!s) return "";
  const m = s.match(/(\d{1,2}):(\d{2})/); return m ? ("0" + m[1]).slice(-2) + ":" + m[2] : s;
}
function computeHours(start: string, end: string) {
  const a = asTimeStr(start), b = asTimeStr(end); if (!a || !b) return 0;
  const x = a.split(":").map(Number), y = b.split(":").map(Number);
  let d = y[0] * 60 + y[1] - (x[0] * 60 + x[1]); if (d < 0) d += 1440;
  return Math.round((d / 60) * 100) / 100;
}

export function buildAllowedChanSet(auth: Auth, filterChannels?: string[]) {
  let allowed: Record<string, boolean> | null = null;
  if (auth.role === "lead") { allowed = {}; (auth.channels || []).forEach((c) => (allowed![c.toLowerCase()] = true)); }
  if (filterChannels && filterChannels.length) {
    const f: Record<string, boolean> = {}; filterChannels.forEach((c) => (f[String(c).toLowerCase()] = true));
    if (allowed) { Object.keys(allowed).forEach((c) => { if (!f[c]) delete allowed![c]; }); } else allowed = f;
  }
  return allowed;
}

export function filterPayrollList(list: any[], allowed: Record<string, boolean> | null) {
  if (!allowed) return list;
  const out: any[] = [];
  const pass = (ch: unknown) => !ch || allowed[String(ch).toLowerCase()];
  list.forEach((p) => {
    const lines = (p.lines || []).filter((l: any) => pass(l.channel));
    const hoTroItems = (p.hoTroItems || []).filter((it: any) => pass(it.channel));
    const luongCungItems = (p.luongCungItems || []).filter((it: any) => pass(it.channel));
    const probationItems = (p.probationItems || []).filter((it: any) => it.channel && allowed[String(it.channel).toLowerCase()]);
    if (!lines.length && !hoTroItems.length && !luongCungItems.length && !probationItems.length) return;
    const sum = (a: any[], k: string) => Math.round(a.reduce((s, x) => s + (Number(x[k]) || 0), 0));
    const c = { ...p };
    c.lines = lines; c.hoTroItems = hoTroItems; c.luongCungItems = luongCungItems;
    c.hoursPay = sum(lines, "hoursPay"); c.bonus = sum(lines, "thuong"); c.hoTro = sum(hoTroItems, "amount");
    c.luongCung = sum(luongCungItems, "amount"); c.probationItems = probationItems; c.probation = sum(probationItems, "amount");
    c.total = c.hoursPay + c.bonus + c.hoTro + c.luongCung + c.probation;
    out.push(c);
  });
  return out;
}

export function computeRaw(auth: Auth, month: string, d: PayInput, filterChannels?: string[]) {
  const cfg = d.cfg;
  const hourly: Record<string, { host: number; tro: number }> = {};
  d.hourly.forEach((r) => { if (!r.ten) return; hourly[L(r.ten) + "|" + L(r.kenh)] = { host: num(r.don_gia_host), tro: num(r.don_gia_tro) }; });
  const mocMap: Record<string, { mocGmvh: number; mocNmv: number }> = {};
  d.moc.forEach((r) => { if (!r.ten) return; mocMap[L(r.ten) + "|" + L(r.kenh) + "|" + String(r.vai_tro)] = { mocGmvh: num(r.moc_gmvh), mocNmv: num(r.moc_nmv) }; });
  let returnRate: number | null = null;
  d.hoanHuy.forEach((r) => { if (String(r.thang).slice(0, 7) === month) returnRate = num(r.ty_le) / 100; });
  if (returnRate === null) returnRate = num(cfg.hoan_huy) / 100;

  const diemMap: Record<string, number> = {};
  d.diem.forEach((r) => { if (String(r.thang).slice(0, 7) === month) diemMap[L(r.ten)] = num(r.diem); });
  const diemDefaultSet: Record<string, boolean> = {};
  (d.prevSnapshotForDiem || []).forEach((sp) => {
    const k = L(sp.name);
    if (diemMap[k] === undefined && sp.diem != null) { diemMap[k] = num(sp.diem); diemDefaultSet[k] = true; }
  });

  const hoTroMap: Record<string, { amount: number; note: string; items: any[] }> = {};
  d.hoTro.forEach((r) => {
    if (String(r.thang).slice(0, 7) !== month) return;
    const nm = String(r.ten || "").trim(); if (!nm) return;
    const nl = nm.toLowerCase(), amt = num(r.so_tien), note = String(r.ghi_chu || "");
    const o = (hoTroMap[nl] = hoTroMap[nl] || { amount: 0, note: "", items: [] });
    o.amount += amt; o.items.push({ channel: String(r.kenh || "").trim(), amount: amt, note });
    if (note) o.note = o.note ? o.note + " · " + note : note;
  });
  const luongCungMap: Record<string, { total: number; name: string; items: any[] }> = {};
  d.luongCung.forEach((r) => {
    if (String(r.thang).slice(0, 7) !== month) return;
    const nm = String(r.ten || "").trim(); if (!nm) return;
    const nl = nm.toLowerCase(), amt = num(r.so_tien);
    const o = (luongCungMap[nl] = luongCungMap[nl] || { total: 0, name: nm, items: [] });
    o.total += amt; o.items.push({ channel: String(r.kenh || "").trim(), amount: amt, note: String(r.ghi_chu || "") });
  });

  let leadChanSet: Record<string, boolean> | null = null;
  if (auth.role === "lead") { leadChanSet = {}; (auth.channels || []).forEach((c) => (leadChanSet![c.toLowerCase()] = true)); }
  const holMap: Record<string, number> = {};
  d.holidays.forEach((h) => { const dt = String(h.ngay).slice(0, 10); if (dt.slice(0, 7) === month && num(h.he_so) > 0) holMap[dt] = num(h.he_so); });

  const groups: Record<string, any> = {};
  d.shifts.forEach((r) => {
    const end = String(r.ket_thuc ?? "").trim(); if (!end) return;
    const dateStr = String(r.ngay).slice(0, 10); if (dateStr.slice(0, 7) !== month) return;
    const name = String(r.ten).trim(); if (!name) return;
    const channel = String(r.kenh ?? "").trim();
    if (leadChanSet && !leadChanSet[channel.toLowerCase()]) return;
    const role = String(r.vi_tri) === "Host Live" ? "Host" : "Tro";
    const hrs = num(r.gio_live) || computeHours(asTimeStr(r.bat_dau), end);
    const gmv = num(r.gmv);
    const key = name.toLowerCase() + "|" + channel.toLowerCase() + "|" + role;
    const g = (groups[key] = groups[key] || { name, channel, role, gmv: 0, hours: 0, holidayExtraHours: 0, holidayDays: {} });
    g.gmv += gmv; g.hours += hrs;
    const hm = holMap[dateStr];
    if (hm && hm !== 1) {
      g.holidayExtraHours += hrs * (hm - 1);
      const hd = (g.holidayDays[dateStr] = g.holidayDays[dateStr] || { date: dateStr, mult: hm, hours: 0 });
      hd.hours += hrs;
    }
  });

  const people: Record<string, any> = {};
  const kpi = (actual: number, moc: number, heso: number, tran: number) => (!moc || moc <= 0 ? 0 : Math.min((actual / moc) * heso, tran));
  Object.keys(groups).forEach((key) => {
    const g = groups[key], nmL = g.name.toLowerCase();
    const p = (people[nmL] = people[nmL] || { name: g.name, hoursPay: 0, bonus: 0, lines: [] });
    const rateObj = hourly[nmL + "|" + g.channel.toLowerCase()] || hourly[nmL + "|"] || { host: 0, tro: 0 };
    const rate = rateObj[g.role === "Host" ? "host" : "tro"];
    const hoursPay = (g.hours + (g.holidayExtraHours || 0)) * rate;
    p.hoursPay += hoursPay;
    const gmvh = g.hours > 0 ? g.gmv / g.hours : 0;
    const nmv = (g.gmv / (cfg.vat || 1.08)) * (1 - returnRate!);
    const diem = diemMap[nmL] || 0;
    const mocK = mocMap[key] || { mocGmvh: 0, mocNmv: 0 };
    const k1 = kpi(gmvh, mocK.mocGmvh, cfg.heso_gmvh, cfg.tran_gmvh);
    const k2 = kpi(nmv, mocK.mocNmv, cfg.heso_nmv, cfg.tran_nmv);
    const k3 = kpi(diem, cfg.moc_diem, cfg.heso_diem, cfg.tran_diem);
    const tongHeSo = k1 + k2 + k3;
    const pct = (g.role === "Host" ? cfg.pct_host : cfg.pct_tro) / 100;
    const thuong = pct * tongHeSo * nmv;
    p.bonus += thuong;
    p.lines.push({
      channel: g.channel, role: g.role, gmv: g.gmv, hours: Math.round(g.hours * 100) / 100,
      gmvh: Math.round(gmvh), nmv: Math.round(nmv), diem,
      k1: Math.round(k1 * 1000) / 1000, k2: Math.round(k2 * 1000) / 1000, k3: Math.round(k3 * 1000) / 1000,
      tongHeSo: Math.round(tongHeSo * 1000) / 1000, hoursPay: Math.round(hoursPay), thuong: Math.round(thuong),
      holidayExtra: Math.round((g.holidayExtraHours || 0) * rate),
      holidayDays: Object.keys(g.holidayDays || {}).sort().map((k: string) => { const hd = g.holidayDays[k]; return { date: hd.date, mult: hd.mult, hours: Math.round(hd.hours * 100) / 100 }; }),
      mocGmvh: mocK.mocGmvh, mocNmv: mocK.mocNmv, mocDiem: cfg.moc_diem,
      pctApplied: pct, hesoGmvh: cfg.heso_gmvh, tranGmvh: cfg.tran_gmvh, hesoNmv: cfg.heso_nmv, tranNmv: cfg.tran_nmv, hesoDiem: cfg.heso_diem, tranDiem: cfg.tran_diem,
    });
  });

  let list: any[] = Object.keys(people).map((k) => {
    const p = people[k], ht = hoTroMap[k] || { amount: 0, note: "", items: [] }, lc = luongCungMap[k] || { total: 0, items: [] };
    return {
      name: p.name, hoursPay: Math.round(p.hoursPay), bonus: Math.round(p.bonus), hoTro: Math.round(ht.amount), hoTroNote: ht.note, hoTroItems: ht.items,
      luongCung: Math.round(lc.total), luongCungItems: lc.items, total: Math.round(p.hoursPay + p.bonus + ht.amount + lc.total),
      diem: diemMap[k] || 0, diemDefault: !!diemDefaultSet[k], lines: p.lines, isFixed: false,
    };
  });
  const extra: Record<string, boolean> = {};
  Object.keys(hoTroMap).forEach((k) => { if (!people[k]) extra[k] = true; });
  Object.keys(luongCungMap).forEach((k) => { if (!people[k]) extra[k] = true; });
  Object.keys(extra).forEach((k) => {
    const ht = hoTroMap[k] || { amount: 0, note: "", items: [] }, lc: any = luongCungMap[k] || { total: 0, items: [], name: "" };
    if (!ht.amount && !ht.note && !lc.total) return;
    const nameGuess = lc.name || k.charAt(0).toUpperCase() + k.slice(1);
    list.push({ name: nameGuess, hoursPay: 0, bonus: 0, hoTro: Math.round(ht.amount), hoTroNote: ht.note, hoTroItems: ht.items, luongCung: Math.round(lc.total), luongCungItems: lc.items,
      total: Math.round(ht.amount + lc.total), diem: diemMap[k] || 0, diemDefault: !!diemDefaultSet[k], lines: [], isFixed: lc.total > 0 && !ht.amount });
  });
  const nameByLower: Record<string, string> = {};
  d.hoTro.forEach((r) => { if (String(r.thang).slice(0, 7) === month) { const nm = String(r.ten || "").trim(); if (nm) nameByLower[nm.toLowerCase()] = nm; } });
  list.forEach((x) => { if (nameByLower[x.name.toLowerCase()]) x.name = nameByLower[x.name.toLowerCase()]; });
  if (auth.role === "lead") { const set = d.leadStaff || {}; list = list.filter((x) => set[x.name.toLowerCase()]); }
  else if (auth.role === "staff") { const nm = auth.staff.toLowerCase(); list = list.filter((x) => x.name.toLowerCase() === nm); }

  const probRows = d.probation.filter((r) => r.ten && String(r.thang).slice(0, 7) === month)
    .map((r) => ({ name: String(r.ten).replace(/\s+/g, " ").trim(), channel: String(r.kenh || "").trim(), amount: num(r.so_tien), note: String(r.ghi_chu || "") }));
  if (probRows.length) {
    const byLower: Record<string, any> = {}; list.forEach((x) => (byLower[x.name.toLowerCase()] = x));
    const probBy: Record<string, any> = {};
    probRows.forEach((it) => {
      const k = it.name.toLowerCase();
      if (auth.role === "staff" && k !== auth.staff.toLowerCase()) return;
      (probBy[k] = probBy[k] || { name: it.name, items: [] }).items.push({ channel: it.channel, amount: it.amount, note: it.note });
    });
    Object.keys(probBy).forEach((k) => {
      const items = probBy[k].items, sum = Math.round(items.reduce((s: number, it: any) => s + (Number(it.amount) || 0), 0));
      let person = byLower[k];
      if (!person) {
        person = { name: probBy[k].name, hoursPay: 0, bonus: 0, hoTro: 0, hoTroNote: "", hoTroItems: [], luongCung: 0, luongCungItems: [], total: 0, diem: 0, lines: [], isFixed: false, isProbation: true };
        list.push(person); byLower[k] = person;
      }
      person.probationItems = items; person.probation = sum; person.total = Math.round((person.total || 0) + sum);
    });
  }
  list = filterPayrollList(list, buildAllowedChanSet(auth, filterChannels));
  list.sort((a, b) => b.total - a.total);
  return { success: true, month, returnRate: Math.round(returnRate! * 10000) / 100, data: list };
}
