import * as g1 from "./group1.ts";

const CORS = {
  "access-control-allow-origin": "*",
  "access-control-allow-headers": "content-type, authorization, apikey, x-client-info",
  "access-control-allow-methods": "POST, OPTIONS",
};
const PUBLIC = ["login","whoami","changeMyPassword","listUsers","createUser","updateUser","resetPassword","deleteUser","setMyChannels",
  "getStaff","addStaff","deleteStaff","getChannels","addChannel","deleteChannel","getShifts","addShift","updateShift","deleteShift",
  "getShiftPhoto","checkIn","checkOut","getDatabaseUrl"];
const FN: Record<string, (...a: any[]) => Promise<unknown>> = {};
PUBLIC.forEach((n) => { if ((g1 as any)[n]) FN[n] = (g1 as any)[n]; });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const { fn, args } = await req.json();
    const f = FN[fn];
    if (!f) return json({ success: false, error: "Chức năng '" + fn + "' chưa được chuyển sang hệ thống mới." });
    return json(await f(...(Array.isArray(args) ? args : [])));
  } catch (e) {
    return json({ success: false, error: "Lỗi máy chủ: " + (e as Error).message });
  }
});
function json(o: unknown) { return new Response(JSON.stringify(o), { headers: { ...CORS, "content-type": "application/json" } }); }
