// Cấu hình bản quyền ExcelLike (được đưa vào app và lic.exe khi đóng gói).
// PUBLIC_KEY: khóa công khai Ed25519 — khớp với khóa ký trong app Quản lý Key. KHÔNG đặt khóa bí mật ở đây.
// REVOKED: số thứ tự các key đã thu hồi (xuất từ app Quản lý Key ▸ Thu hồi) — có hiệu lực từ bản cài kế tiếp.
module.exports = {
  PUBLIC_KEY: 'vWaTBOFTa3/PYRhRGpD7wM/qITQkcvNavbJ6Xo/bc6A=',
  REVOKED: [],
};
