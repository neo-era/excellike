// Bảo vệ Trang tính / Sổ làm việc (0.9): mật khẩu băm SHA-512 có muối, lặp 100.000 lần như Excel 2013+,
// nên tệp lưu ra Excel mở khóa được bằng cùng mật khẩu và ngược lại.
export type HashAlg = 'SHA-512' | 'SHA-384' | 'SHA-256' | 'SHA-1';
/** alg: loại băm; thuật toán lạ trong tệp vẫn giữ nguyên tên để lưu lại (không mở khóa được, nhưng không bị mất bảo vệ) */
export interface PwHash { alg: HashAlg | 'legacy' | (string & {}); hash: string; salt: string; spin: number }
/** tên thuật toán trong tệp (algorithmName) → loại băm hỗ trợ; null nếu không hỗ trợ */
export function algOf(name: unknown): HashAlg | null { const n = String(name ?? '').toUpperCase().replace(/[^A-Z0-9]/g, ''); return n === 'SHA512' ? 'SHA-512' : n === 'SHA384' ? 'SHA-384' : n === 'SHA256' ? 'SHA-256' : n === 'SHA1' ? 'SHA-1' : null; }
/** số vòng tối đa chấp nhận khi mở tệp (chuẩn ECMA-376: ≤ 10.000.000) */
export const MAX_SPIN = 10_000_000;
/** quyền khi trang được bảo vệ (true = được phép) — tên theo hộp thoại Excel */
export interface ProtectAllow {
  selectLocked: boolean; selectUnlocked: boolean; formatCells: boolean; formatColumns: boolean; formatRows: boolean;
  insertColumns: boolean; insertRows: boolean; insertHyperlinks: boolean; deleteColumns: boolean; deleteRows: boolean;
  sort: boolean; autoFilter: boolean; pivotTables: boolean; objects: boolean; scenarios: boolean;
}
export const DEFAULT_ALLOW: ProtectAllow = { selectLocked: true, selectUnlocked: true, formatCells: false, formatColumns: false, formatRows: false, insertColumns: false, insertRows: false, insertHyperlinks: false, deleteColumns: false, deleteRows: false, sort: false, autoFilter: false, pivotTables: false, objects: false, scenarios: false };
export const ALLOW_LABELS: [keyof ProtectAllow, string][] = [
  ['selectLocked', 'Chọn ô bị khóa'], ['selectUnlocked', 'Chọn ô không bị khóa'], ['formatCells', 'Định dạng ô'], ['formatColumns', 'Định dạng cột'], ['formatRows', 'Định dạng hàng'],
  ['insertColumns', 'Chèn cột'], ['insertRows', 'Chèn hàng'], ['insertHyperlinks', 'Chèn siêu kết nối'], ['deleteColumns', 'Xóa cột'], ['deleteRows', 'Xóa hàng'],
  ['sort', 'Sắp xếp'], ['autoFilter', 'Dùng Bộ lọc Tự động'], ['pivotTables', 'Dùng PivotTable'], ['objects', 'Sửa đối tượng'], ['scenarios', 'Sửa kịch bản'],
];
export interface SheetProtect { pw?: PwHash; allow: ProtectAllow }

const b64 = (u: Uint8Array) => { let s = ''; for (const x of u) s += String.fromCharCode(x); return btoa(s); };
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
async function dig(alg: HashAlg, d: Uint8Array): Promise<Uint8Array> { return new Uint8Array(await crypto.subtle.digest(alg, d)); }
/** băm như Excel: H0 = SHA512(muối + UTF-16LE(mật khẩu)); Hi = SHA512(H(i-1) + UInt32LE(i)) */
export async function hashPw(pw: string, salt?: string, spin = 100000, alg: HashAlg = 'SHA-512'): Promise<PwHash> {
  const sb = salt ? unb64(salt) : crypto.getRandomValues(new Uint8Array(16));
  const pwb = new Uint8Array(pw.length * 2);
  for (let i = 0; i < pw.length; i++) { const c = pw.charCodeAt(i); pwb[2 * i] = c & 255; pwb[2 * i + 1] = c >> 8; }
  const first = new Uint8Array(sb.length + pwb.length); first.set(sb); first.set(pwb, sb.length);
  let h = await dig(alg, first);
  const L = h.length, buf = new Uint8Array(L + 4);
  for (let i = 0; i < spin; i++) {
    buf.set(h); buf[L] = i & 255; buf[L + 1] = (i >>> 8) & 255; buf[L + 2] = (i >>> 16) & 255; buf[L + 3] = (i >>> 24) & 255;
    h = await dig(alg, buf);
  }
  return { alg, hash: b64(h), salt: b64(sb), spin };
}
/** mật khẩu kiểu cũ (Excel 97–2010, thuộc tính password="CC1A") */
export function legacyHash(pw: string): string {
  let h = 0;
  for (let i = pw.length - 1; i >= 0; i--) { h = ((h >> 14) & 1) | ((h << 1) & 0x7fff); h ^= pw.charCodeAt(i); }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff); h ^= pw.length; h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}
export async function checkPw(pw: string, h: PwHash): Promise<boolean> {
  if (h.alg === 'legacy') return legacyHash(pw) === h.hash.toUpperCase().padStart(4, '0');
  if (!algOf(h.alg) || !(h.spin >= 0 && h.spin <= MAX_SPIN)) return false;
  return (await hashPw(pw, h.salt, h.spin, algOf(h.alg)!)).hash === h.hash;
}
