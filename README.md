# EAZ Chấm công Livestream

Hệ thống chấm công & tính lương livestream (Gooby Vietnam / EAZ).

| Thư mục | Nội dung |
|---|---|
| `web/` | Giao diện web tĩnh (chạy được trên GitHub Pages). `assets/shim.js` thay `google.script.run` bằng lời gọi tới Supabase |
| `supabase/functions/chamcong-api/` | Máy chủ (Edge Function). Hiện đã có nhóm 1: đăng nhập, tài khoản, nhân sự, kênh, ca live, check-in/out |
| `supabase/migrations/` | Cấu trúc bảng (schema `chamcong`) |
| `apps-script-cu/` | Bản Apps Script cũ (vẫn chạy song song trong lúc chuyển đổi) |

Dữ liệu: Supabase project `jnzeawdffiwnblrlwtgx`, schema `chamcong` (RLS bật, chỉ Edge Function truy cập).

## Trạng thái chuyển đổi
- [x] Giai đoạn 1: đăng nhập, nhân sự, kênh, ca live, check-in/out, tài khoản
- [x] Giai đoạn 2: yêu cầu sửa ca, import ca
- [x] Giai đoạn 3: bảng lương, chốt tháng, xuất Excel, lương cứng/thử việc/ngày lễ
- [ ] Telegram, TikTok Live Data, chuyển tên miền live.eaztiktok.online
