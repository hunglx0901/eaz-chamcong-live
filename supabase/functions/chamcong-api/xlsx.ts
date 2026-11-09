import ExcelJS from "npm:exceljs@4.4.0";
import { sql, ok, fail, authErr, denyErr, verifyToken, fmtVnd } from "./lib.ts";
import { computePayroll, getPayrollConfig } from "./group3.ts";

const L = (s: unknown) => String(s ?? "").trim().toLowerCase();
const MONTH_RE = /^\d{4}-\d{2}$/;
const NCOLS = 34;

function b64(buf: Uint8Array) {
  let s = ""; for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode(...buf.subarray(i, i + 0x8000));
  return btoa(s);
}
function colName(n: number) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

export async function exportPayrollXlsx(token: string, month: string, filterChannels?: string[]) {
  const auth = await verifyToken(token); if (!auth) return authErr();
  if (auth.role !== "owner" && auth.role !== "lead") return denyErr();
  try {
    month = String(month || "").trim(); if (!MONTH_RE.test(month)) return fail("Tháng không hợp lệ (YYYY-MM).");
    const pay: any = await computePayroll(token, month, filterChannels);
    if (!pay.success) return pay;
    const data: any[] = pay.data || [];
    const chanFilter = Array.isArray(filterChannels) ? filterChannels.filter((c) => !!c) : [];
    const chanSet: Record<string, boolean> = {}; chanFilter.forEach((c) => (chanSet[String(c).toLowerCase()] = true));
    const pass = (ch: unknown) => !chanFilter.length || !!chanSet[String(ch || "").toLowerCase()];

    const cfgRes: any = await getPayrollConfig(token); if (!cfgRes.success) return cfgRes;
    const cfg = cfgRes.chung || {};
    const mocDiem = Number(cfg.moc_diem) || 0, pctHost = (Number(cfg.pct_host) || 0) / 100, pctTro = (Number(cfg.pct_tro) || 0) / 100;
    const rateMap: Record<string, any> = {};
    (cfgRes.hourly || []).forEach((h: any) => (rateMap[L(h.name) + "|" + L(h.channel)] = { host: h.host || 0, tro: h.tro || 0 }));
    const getRate = (n: string, ch: string) => rateMap[L(n) + "|" + L(ch)] || rateMap[L(n) + "|"] || { host: 0, tro: 0 };
    const mocMap: Record<string, any> = {};
    (cfgRes.moc || []).forEach((m: any) => (mocMap[L(m.name) + "|" + L(m.channel) + "|" + m.role] = { mocGmvh: m.mocGmvh || 0, mocNmv: m.mocNmv || 0 }));
    const getMoc = (n: string, ch: string, role: string) => mocMap[L(n) + "|" + L(ch) + "|" + role] || { mocGmvh: 0, mocNmv: 0 };

    const byChannel: Record<string, any[]> = {}, channelSet: Record<string, boolean> = {};
    const ensureChan = (ch: string) => { if (!byChannel[ch]) byChannel[ch] = []; channelSet[ch] = true; };

    data.forEach((p) => {
      const hoTroByChan: Record<string, any> = {}, luongCungByChan: Record<string, any> = {};
      (p.hoTroItems || []).forEach((it: any) => {
        const ch = String(it.channel || "").trim(), k = ch.toLowerCase();
        const o = (hoTroByChan[k] = hoTroByChan[k] || { channel: ch, amount: 0, notes: [] });
        o.amount += Number(it.amount) || 0; if (it.note) o.notes.push(String(it.note));
      });
      const lcSrc = (p.luongCungItems || []).concat((p.probationItems || []).map((it: any) => ({ channel: it.channel, amount: it.amount, note: "Thử việc" + (it.note ? ": " + it.note : "") })));
      lcSrc.forEach((it: any) => {
        const ch = String(it.channel || "").trim(), k = ch.toLowerCase();
        const o = (luongCungByChan[k] = luongCungByChan[k] || { channel: ch, amount: 0, notes: [] });
        o.amount += Number(it.amount) || 0; if (it.note) o.notes.push(String(it.note));
      });
      const blank = (o: any) => ({ name: p.name, luongTroLive: 0, luongLive: 0, gioTroLive: 0, gioLive: 0, gmvhMoc: 0, gmvhThucTe: 0, k1: 0, nmvMoc: 0, nmvThucTe: 0, k2: 0, k3: 0, tongHeSo: 0, pctLiveChinh: 0, pctTroLive: 0, thuong: 0, ...o });

      if (!p.lines || !p.lines.length) {
        const keys = Object.keys({ ...hoTroByChan, ...luongCungByChan });
        if (!keys.length) return;
        keys.forEach((k) => {
          const chName = (hoTroByChan[k] && hoTroByChan[k].channel) || (luongCungByChan[k] && luongCungByChan[k].channel) || "Khác";
          if (!pass(chName)) return;
          const ch = chName || "Khác"; ensureChan(ch);
          const ht = hoTroByChan[k] || { amount: 0, notes: [] }, lc = luongCungByChan[k] || { amount: 0, notes: [] };
          byChannel[ch].push(blank({ channel: ch, position: p.isProbation ? "Thử việc" : (p.isFixed ? "Lương cứng" : "—"), tongLuong: lc.amount,
            diemMoc: mocDiem, diemThucTe: p.diem || 0, hoTro: ht.amount, hoTroNote: ht.notes.concat(lc.notes).filter(Boolean).join(" · "), tongThuNhap: lc.amount + ht.amount }));
        });
        return;
      }
      const sorted = p.lines.slice().sort((a: any, b: any) => b.gmv - a.gmv);
      const top = sorted[0];
      const htApplied: Record<string, boolean> = {}, lcApplied: Record<string, boolean> = {};
      sorted.forEach((l: any) => {
        const ch = l.channel; if (!pass(ch)) return; ensureChan(ch);
        const isHost = l.role === "Host", rate = getRate(p.name, ch), moc = getMoc(p.name, ch, l.role), k = ch.toLowerCase();
        let hoTroCh = 0, hoTroNoteCh = "";
        if (hoTroByChan[k] && !htApplied[k]) { hoTroCh = hoTroByChan[k].amount; hoTroNoteCh = hoTroByChan[k].notes.join(" · "); htApplied[k] = true; }
        let lcCh = 0, lcNoteCh = "";
        if (luongCungByChan[k] && !lcApplied[k]) { lcCh = luongCungByChan[k].amount; lcNoteCh = luongCungByChan[k].notes.join(" · "); lcApplied[k] = true; }
        const holNote = (l.holidayDays && l.holidayDays.length)
          ? "Lễ: " + l.holidayDays.map((d: any) => d.date.slice(8, 10) + "/" + d.date.slice(5, 7) + " x" + d.mult + " (" + d.hours + "h)").join(", ") + " (+" + fmtVnd(l.holidayExtra) + "đ)" : "";
        byChannel[ch].push({
          hoursPayLine: l.hoursPay, name: p.name, channel: ch, position: isHost ? "Host live" : "LA",
          luongTroLive: rate.tro, luongLive: rate.host, gioTroLive: isHost ? 0 : l.hours, gioLive: isHost ? l.hours : 0,
          tongLuong: l.hoursPay + lcCh, gmvhMoc: moc.mocGmvh, gmvhThucTe: l.gmvh, k1: l.k1, nmvMoc: moc.mocNmv, nmvThucTe: l.nmv, k2: l.k2,
          diemMoc: mocDiem, diemThucTe: l.diem || 0, k3: l.k3, tongHeSo: l.tongHeSo, pctLiveChinh: pctHost, pctTroLive: pctTro, thuong: l.thuong,
          hoTro: hoTroCh, hoTroNote: [hoTroNoteCh, lcNoteCh, holNote].filter(Boolean).join(" · "), tongThuNhap: l.hoursPay + l.thuong + hoTroCh + lcCh,
        });
      });
      Object.keys({ ...hoTroByChan, ...luongCungByChan }).forEach((k) => {
        if (htApplied[k] || lcApplied[k]) return;
        const chName = (hoTroByChan[k] && hoTroByChan[k].channel) || (luongCungByChan[k] && luongCungByChan[k].channel);
        if (!chName || !pass(chName)) return; ensureChan(chName);
        const ht = hoTroByChan[k] || { amount: 0, notes: [] }, lc = luongCungByChan[k] || { amount: 0, notes: [] };
        byChannel[chName].push(blank({ channel: chName, position: "—", tongLuong: lc.amount, hoTro: ht.amount, hoTroNote: ht.notes.concat(lc.notes).filter(Boolean).join(" · "), tongThuNhap: ht.amount + lc.amount }));
      });
      const posTop = top && (top.role === "Host" ? "Host live" : "LA");
      if (top && hoTroByChan[""] && pass(top.channel)) {
        const rows = byChannel[top.channel] || [];
        for (let i = rows.length - 1; i >= 0; i--) if (rows[i].name === p.name && rows[i].position === posTop) {
          rows[i].hoTro += hoTroByChan[""].amount; const add = hoTroByChan[""].notes.join(" · ");
          if (add) rows[i].hoTroNote = rows[i].hoTroNote ? rows[i].hoTroNote + " · " + add : add;
          rows[i].tongThuNhap += hoTroByChan[""].amount; break;
        }
      }
      if (top && luongCungByChan[""] && pass(top.channel)) {
        const rows = byChannel[top.channel] || [];
        for (let i = rows.length - 1; i >= 0; i--) if (rows[i].name === p.name && rows[i].position === posTop) {
          rows[i].tongLuong += luongCungByChan[""].amount; rows[i].tongThuNhap += luongCungByChan[""].amount; break;
        }
      }
    });
    const channelOrder = Object.keys(channelSet).sort();

    const codeByName: Record<string, string> = {};
    try { (await sql`select ten, coalesce(ma_nv,'') ma from chamcong.nhan_su`).forEach((r: any) => { const n = L(r.ten), c = String(r.ma || "").trim(); if (n && c) codeByName[n] = c; }); } catch { /* để trống */ }

    const monthLabel = month.split("-").reverse().join("/");
    const aoa: any[][] = [];
    const empty = () => new Array(NCOLS).fill("");
    const r1 = empty(); r1[0] = "Tên Đơn vị: Công ty Cổ phần EAZ Việt Nam"; aoa.push(r1);
    aoa.push(empty());
    const t = empty(); t[5] = "THANH TOÁN CHI PHÍ CỘNG TÁC — Tháng " + monthLabel; aoa.push(t);
    aoa.push(empty());
    const sh = empty();
    sh[0] = "THÔNG TIN NHÂN SỰ"; sh[6] = "VỊ TRÍ"; sh[7] = "Lương cứng"; sh[8] = "LƯƠNG THEO GIỜ"; sh[13] = "GMV/h"; sh[17] = "NMV/tháng"; sh[21] = "Điểm phòng live";
    sh[25] = "TỔNG HỆ SỐ"; sh[26] = "THƯỞNG THỰC TẾ"; sh[29] = "THƯỞNG HỖ TRỢ"; sh[30] = "TỔNG THU NHẬP"; sh[33] = "Ghi chú";
    aoa.push(sh);
    aoa.push(["STT", "KHỐI", "PHÒNG/ BỘ PHẬN", "MÃ NV\n(000000)", "TÊN NHÂN VIÊN", "KÊNH", "", "", "Lương trợ live", "Lương Live", "Giờ trợ live", "Giờ Live", "TỔNG LƯƠNG",
      "Hệ số", "Mốc thưởng", "Thực tế", "Kết quả hệ số", "Hệ số", "Mốc thưởng", "Thực tế", "Kết quả hệ số", "Hệ số", "Mốc thưởng", "Thực tế", "Kết quả hệ số",
      "", "LIVE CHÍNH", "TRỢ LIVE", "THƯỞNG", "", "TỔNG THU NHẬP", "PIT", "THỰC NHẬN", ""]);

    let stt = 1; const chanRows: number[] = [];
    const f = (n: number, d: number) => Number((n || 0).toFixed(d));
    channelOrder.forEach((ch) => {
      const cr = empty(); cr[0] = ch; aoa.push(cr); chanRows.push(aoa.length);
      byChannel[ch].forEach((r) => {
        const pure = r.hoursPayLine != null ? r.hoursPayLine : ((r.gioTroLive || 0) * (r.luongTroLive || 0) + (r.gioLive || 0) * (r.luongLive || 0));
        const lcCell = Math.max(0, (r.tongLuong || 0) - pure);
        const tot = pure + r.thuong + r.hoTro + lcCell;
        aoa.push([stt++, "", "", codeByName[L(r.name)] || "", r.name, r.channel, r.position, lcCell,
          r.luongTroLive, r.luongLive, r.gioTroLive, r.gioLive, pure,
          f(r.k1, 2), r.gmvhMoc, r.gmvhThucTe, f(r.k1, 4), f(r.k2, 2), r.nmvMoc, r.nmvThucTe, f(r.k2, 4), f(r.k3, 2), r.diemMoc, r.diemThucTe, f(r.k3, 4),
          f(r.tongHeSo, 4), r.pctLiveChinh, r.pctTroLive, r.thuong, r.hoTro, tot, 0, tot, r.hoTroNote]);
      });
    });
    const totalRows = aoa.length;
    if (totalRows < 6) return fail("Không có dữ liệu để xuất.");

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Part time", { views: [{ state: "frozen", ySplit: 6 }] });
    aoa.forEach((row) => ws.addRow(row));
    const fill = (rng: string, argb: string) => {
      const [a, b] = rng.split(":"); const cell = (s: string) => { const m = s.match(/^([A-Z]+)(\d+)$/)!; return { c: m[1], r: +m[2] }; };
      const A = cell(a), B = cell(b || a);
      const toN = (c: string) => c.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      for (let r = A.r; r <= B.r; r++) for (let c = toN(A.c); c <= toN(B.c); c++) ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
    };
    const border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } } as any;
    for (let r = 1; r <= totalRows; r++) for (let c = 1; c <= NCOLS; c++) ws.getCell(r, c).font = { name: "Times New Roman", size: 10 };
    ["A1:K1", "F3:P3", "A5:F5", "G5:G6", "H5:H6", "I5:M5", "N5:Q5", "R5:U5", "V5:Y5", "Z5:Z6", "AA5:AC5", "AD5:AD6", "AE5:AG5", "AH5:AH6"].forEach((m) => { try { ws.mergeCells(m); } catch { /* bỏ qua */ } });
    ws.getCell("A1").font = { name: "Times New Roman", size: 11, bold: true }; ws.getCell("A1").alignment = { horizontal: "left", vertical: "middle" };
    ws.getCell("F3").font = { name: "Times New Roman", size: 14, bold: true }; ws.getCell("F3").alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(3).height = 22.5;
    ["A5:F5", "G5:G6", "AH5:AH6", "A6:F6"].forEach((x) => fill(x, "FFD9E2F3")); fill("H5:H6", "FFBDD7EE");
    ["I5:M5", "I6:M6"].forEach((x) => fill(x, "FFF7CAAC"));
    ["N5:Q5", "R5:U5", "V5:Y5", "Z5:Z6", "AA5:AC5", "N6:Y6", "AA6:AC6"].forEach((x) => fill(x, "FFFFE598"));
    ["AD5:AD6", "AE5:AG5", "AE6:AG6"].forEach((x) => fill(x, "FFC5E0B3"));
    for (let r = 5; r <= 6; r++) for (let c = 1; c <= NCOLS; c++) { const ce = ws.getCell(r, c); ce.font = { name: "Times New Roman", size: 10, bold: true }; ce.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; ce.border = border; }
    ws.getRow(5).height = 22.5; ws.getRow(6).height = 30;
    const fmts: Record<number, string> = { 1: "0", 8: "#,##0", 9: "#,##0", 10: "#,##0", 11: "0.0", 12: "0.0", 13: "#,##0", 14: "0.00", 15: "#,##0", 16: "#,##0", 17: "0.0000", 18: "0.00", 19: "#,##0", 20: "#,##0", 21: "0.0000", 22: "0.00", 23: "0.0", 24: "0.0", 25: "0.0000", 26: "0.0000", 27: "0.00%", 28: "0.00%", 29: "#,##0", 30: "#,##0", 31: "#,##0", 32: "#,##0", 33: "#,##0", 4: "@" };
    const chanSetRows = new Set(chanRows);
    for (let r = 7; r <= totalRows; r++) {
      for (let c = 1; c <= NCOLS; c++) { const ce = ws.getCell(r, c); ce.border = border; if (!chanSetRows.has(r) && fmts[c]) ce.numFmt = fmts[c]; }
      if (!chanSetRows.has(r)) {
        ws.getCell(r, 1).alignment = { horizontal: "center", vertical: "middle" };
        ws.getCell(r, 5).alignment = { horizontal: "left", vertical: "middle" };
        ws.getCell(r, 6).alignment = { horizontal: "center", vertical: "middle" }; ws.getCell(r, 7).alignment = { horizontal: "center", vertical: "middle" };
        ws.getCell(r, 34).alignment = { horizontal: "left", vertical: "middle" };
      }
    }
    chanRows.forEach((r) => { fill(`A${r}:AH${r}`, "FFFFE598"); for (let c = 1; c <= NCOLS; c++) ws.getCell(r, c).font = { name: "Times New Roman", size: 11, bold: true }; });
    [40, 90, 120, 100, 200, 110, 90, 110, 90, 90, 70, 70, 100, 70, 90, 90, 80, 70, 90, 90, 80, 70, 70, 60, 80, 80, 80, 80, 100, 100, 110, 80, 110, 150]
      .forEach((w, i) => (ws.getColumn(i + 1).width = Math.round(w / 7)));

    const buf = new Uint8Array(await wb.xlsx.writeBuffer());
    const d = new Date(Date.now() + 7 * 3600 * 1000).toISOString();
    const fileName = `BangLuong_${month}_${d.slice(0, 10).replace(/-/g, "")}_${d.slice(11, 16).replace(":", "")}.xlsx`;
    return ok({ fileName, base64: b64(buf) });
  } catch (e) { return fail("Lỗi xuất bảng lương: " + (e as Error).message); }
}
