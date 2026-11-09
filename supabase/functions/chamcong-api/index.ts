import * as g1 from "./group1.ts";
import * as g2 from "./group2.ts";
import * as g3 from "./group3.ts";
import * as g4 from "./xlsx.ts";

const CORS = {
  "access-control-allow-origin": "*",
  "access-control-allow-headers": "content-type, authorization, apikey, x-client-info",
  "access-control-allow-methods": "POST, OPTIONS",
};
const PUBLIC = ["login","whoami","changeMyPassword","listUsers","createUser","updateUser","resetPassword","deleteUser","setMyChannels",
  "getStaff","addStaff","deleteStaff","getChannels","addChannel","deleteChannel","getShifts","addShift","updateShift","deleteShift",
  "getShiftPhoto","checkIn","checkOut","getDatabaseUrl",
  "submitEditRequest","listRequests","approveRequest","rejectRequest","previewImport","importShifts",
  "getPayrollConfig","saveLuongGio","deleteLuongGio","saveThuongMoc","saveThuongChung","saveHoanHuy","saveDiem","saveMonthUpdate","getMonthLockStatus",
  "chotLuongThang","moChotLuongThang","listLuongCungStaff","addLuongCungStaff","updateLuongCungStaff","deleteLuongCungStaff","getLuongCungByMonth","saveLuongCungBatch",
  "computePayroll","computePayrollWithPrev","getProbationByMonth","saveProbationBatch","getHolidays","saveHolidays","mergeStaff","updateStaffCode","exportPayrollXlsx"];
const FN: Record<string, (...a: any[]) => Promise<unknown>> = {};
PUBLIC.forEach((n) => { const f = (g1 as any)[n] || (g2 as any)[n] || (g3 as any)[n] || (g4 as any)[n]; if (f) FN[n] = f; });

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
