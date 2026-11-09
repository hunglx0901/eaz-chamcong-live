import postgres from "npm:postgres@3.4.4";

export const sql = postgres(Deno.env.get("SUPABASE_DB_URL")!, { max: 3, idle_timeout: 20, prepare: false });
export const TZ = "Asia/Ho_Chi_Minh";
export const PHOTO_BUCKET = "chamcong-photos";

export type Auth = { username: string; role: string; channel: string; channels: string[]; staff: string; managed: string[] };
export const ok = (o: Record<string, unknown> = {}) => ({ success: true, ...o });
export const fail = (error: string) => ({ success: false, error });
export const authErr = () => ({ success: false, code: "AUTH", error: "Phiên đăng nhập đã hết. Đăng nhập lại nhé." });
export const denyErr = () => ({ success: false, error: "Bạn không có quyền thực hiện thao tác này." });

/* ---------- thời gian (múi giờ VN) ---------- */
function parts(d = new Date()) {
  const f = new Intl.DateTimeFormat("en-CA", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false });
  const o: Record<string, string> = {};
  f.formatToParts(d).forEach((p) => (o[p.type] = p.value));
  if (o.hour === "24") o.hour = "00";
  return o;
}
export const todayStr = () => { const p = parts(); return `${p.year}-${p.month}-${p.day}`; };
export const timeStr = () => { const p = parts(); return `${p.hour}:${p.minute}`; };
export const nowString = () => { const p = parts(); return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`; };
const pad2 = (n: number) => ("0" + n).slice(-2);

export function asTimeStr(v: unknown): string {
  if (v == null || v === "") return "";
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2}):(\d{1,2})(?::\d{1,2})?$/);
  if (m) return pad2(+m[1]) + ":" + pad2(+m[2]);
  m = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?(?!\d)/);
  if (m) return pad2(Math.min(23, +m[1])) + ":" + pad2(Math.min(59, +m[2]));
  return s;
}
export function computeHours(start: unknown, end: unknown): number {
  const a = asTimeStr(start), b = asTimeStr(end);
  if (!a || !b) return 0;
  const x = a.split(":").map(Number), y = b.split(":").map(Number);
  if (isNaN(x[0]) || isNaN(y[0])) return 0;
  let diff = y[0] * 60 + (y[1] || 0) - (x[0] * 60 + (x[1] || 0));
  if (diff < 0) diff += 1440;
  return Math.round((diff / 60) * 100) / 100;
}
export function roundInTime(hhmm: string) {
  const p = String(hhmm || "").split(":"); const h = +p[0] || 0, m = +p[1] || 0; const t = h * 60 + m;
  if (t % 15 === 0) return pad2(h) + ":" + pad2(m);
  let r = Math.ceil(t / 15) * 15; if (r >= 1440) r -= 1440;
  return pad2(Math.floor(r / 60)) + ":" + pad2(r % 60);
}
export function roundOutTime(hhmm: string) {
  const p = String(hhmm || "").split(":"); const t = (+p[0] || 0) * 60 + (+p[1] || 0);
  const r = Math.floor(t / 15) * 15;
  return pad2(Math.floor(r / 60)) + ":" + pad2(r % 60);
}
export const splitChannels = (s: unknown): string[] => String(s || "").split("|").map((x) => x.trim()).filter(Boolean);
export const leadChannelsLower = (a: Auth) => (a.channels || []).map((x) => x.toLowerCase());
export const leadHasChannel = (a: Auth, ch: unknown) => leadChannelsLower(a).includes(String(ch || "").toLowerCase());

/* ---------- mật khẩu & token (tương thích hash/salt cũ của Apps Script) ---------- */
const enc = new TextEncoder();
const hex = (b: ArrayBuffer) => [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join("");
export async function hashPw(pw: string, salt: string) {
  return hex(await crypto.subtle.digest("SHA-256", enc.encode(String(salt) + "|" + String(pw))));
}
const b64u = (b: ArrayBuffer | Uint8Array) => btoa(String.fromCharCode(...new Uint8Array(b as ArrayBuffer))).replace(/\+/g, "-").replace(/\//g, "_");
const b64uStr = (s: string) => b64u(enc.encode(s));
const unb64uStr = (s: string) => new TextDecoder().decode(Uint8Array.from(atob(s.replace(/-/g, "+").replace(/_/g, "/")), (c) => c.charCodeAt(0)));

let SECRET: string | null = null;
async function secret() {
  if (SECRET) return SECRET;
  const r = await sql`select value from chamcong.app_config where key='auth_secret'`;
  if (r.length) { SECRET = r[0].value; return SECRET!; }
  const s = crypto.randomUUID() + crypto.randomUUID();
  await sql`insert into chamcong.app_config(key,value) values('auth_secret',${s}) on conflict (key) do nothing`;
  SECRET = (await sql`select value from chamcong.app_config where key='auth_secret'`)[0].value;
  return SECRET!;
}
async function hmac(payload: string) {
  const k = await crypto.subtle.importKey("raw", enc.encode(await secret()), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  return b64u(await crypto.subtle.sign("HMAC", k, enc.encode(payload)));
}
export async function makeToken(id: Auth) {
  const json = JSON.stringify({ u: id.username, r: id.role, c: id.channels.join("|"), s: id.staff || "", m: id.managed.join("|"), e: Date.now() + 12 * 3600 * 1000 });
  const payload = b64uStr(json);
  return payload + "." + (await hmac(payload));
}
export async function verifyToken(token: unknown): Promise<Auth | null> {
  try {
    if (!token) return null;
    const p = String(token).split(".");
    if (p.length !== 2) return null;
    if ((await hmac(p[0])) !== p[1]) return null;
    const d = JSON.parse(unb64uStr(p[0]));
    if (Date.now() > Number(d.e)) return null;
    const chans = splitChannels(d.c);
    return { username: d.u, role: d.r, channel: chans[0] || "", channels: chans, staff: d.s, managed: splitChannels(d.m) };
  } catch { return null; }
}

/* ---------- Telegram (cấu hình trong bảng app_config: telegram_token, telegram_chat_id) ---------- */
export const esc = (s: unknown) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
export const fmtVnd = (n: number) => String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
export async function sendTelegram(text: string) {
  try {
    const r = await sql`select key,value from chamcong.app_config where key in ('telegram_token','telegram_chat_id')`;
    const c: Record<string, string> = {}; r.forEach((x) => (c[x.key] = x.value));
    if (!c.telegram_token || !c.telegram_chat_id) return;
    await fetch(`https://api.telegram.org/bot${c.telegram_token}/sendMessage`, {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ chat_id: c.telegram_chat_id, text, parse_mode: "HTML", disable_web_page_preview: true }),
    });
  } catch { /* nuốt lỗi */ }
}

/* ---------- ảnh chấm công: Supabase Storage (bucket private) ---------- */
const SB_URL = Deno.env.get("SUPABASE_URL")!, SB_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
export async function savePhoto(b64: string, hint: string): Promise<string> {
  if (!b64) return "";
  const clean = b64.includes(",") ? b64.split(",")[1] : b64;
  const bytes = Uint8Array.from(atob(clean), (c) => c.charCodeAt(0));
  const path = `${todayStr().slice(0, 7)}/${hint}_${Date.now()}_${crypto.randomUUID().slice(0, 8)}.jpg`;
  const r = await fetch(`${SB_URL}/storage/v1/object/${PHOTO_BUCKET}/${path}`, {
    method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, "content-type": "image/jpeg", "x-upsert": "true" }, body: bytes,
  });
  if (!r.ok) throw new Error("Không lưu được ảnh: " + (await r.text()).slice(0, 120));
  return path;
}
export async function readPhoto(path: string): Promise<string | null> {
  const r = await fetch(`${SB_URL}/storage/v1/object/${PHOTO_BUCKET}/${path}`, { headers: { Authorization: `Bearer ${SB_KEY}` } });
  if (!r.ok) return null;
  const buf = new Uint8Array(await r.arrayBuffer());
  let s = ""; for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode(...buf.subarray(i, i + 0x8000));
  return "data:image/jpeg;base64," + btoa(s);
}
export async function deletePhoto(path: string) {
  if (!path || !path.includes("/")) return;
  try { await fetch(`${SB_URL}/storage/v1/object/${PHOTO_BUCKET}/${path}`, { method: "DELETE", headers: { Authorization: `Bearer ${SB_KEY}` } }); } catch { /* bỏ qua */ }
}
