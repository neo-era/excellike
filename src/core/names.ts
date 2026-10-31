// Tên vùng (Defined Names) — phần xử lý chuỗi công thức.
// HyperFormula chỉ nhận tên gồm chữ Latin tới U+02AF, nên tên tiếng Việt có dấu thanh (ạ, ả, ấ…) được mã hóa
// thành dạng _x1EA1_ khi đưa vào bộ tính và giải mã lại khi đọc ra; người dùng luôn thấy tên gốc.

/** chữ cái dùng được trong tên (Excel cho phép Unicode) */
const ID_START = /[A-Za-z_\\À-￿]/;
const ID_CHAR = /[A-Za-z0-9_.\\À-￿]/;
const HIGH = /[^\u0000-ʯ]/;

/** mã hóa một tên để HyperFormula chấp nhận (không phân biệt hoa thường như Excel) */
export function encName(name: string): string {
  if (!HIGH.test(name)) return name;
  let out = '';
  for (const ch of name) {
    if (HIGH.test(ch)) out += '_x' + ch.toLowerCase().codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0') + '_';
    else out += ch;
  }
  return out;
}
export const hasEncoded = (s: string) => /_x[0-9A-F]{4,5}_/.test(s);
export function decName(name: string): string {
  return name.replace(/_x([0-9A-F]{4,5})_/g, (_m, h: string) => String.fromCodePoint(parseInt(h, 16)));
}

/**
 * Duyệt các định danh (identifier) ngoài chuỗi "…" và tên trang tính '…' của công thức.
 * f(token, start, nextChar) trả về chuỗi thay thế (hoặc undefined để giữ nguyên).
 */
export function mapIdents(formula: string, f: (tok: string, next: string, prev: string) => string | undefined): string {
  let out = '', i = 0;
  const n = formula.length;
  while (i < n) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      // chuỗi hoặc tên trang tính trong ngoặc: "" / '' là ký tự thoát
      let j = i + 1;
      while (j < n) { if (formula[j] === ch) { if (formula[j + 1] === ch) { j += 2; continue; } break; } j++; }
      out += formula.slice(i, j + 1); i = j + 1; continue;
    }
    if (ID_START.test(ch) && !(i > 0 && /[A-Za-z0-9_.$À-￿]/.test(formula[i - 1]))) {
      let j = i + 1;
      while (j < n && ID_CHAR.test(formula[j])) j++;
      const tok = formula.slice(i, j);
      const r = f(tok, formula[j] ?? '', formula[i - 1] ?? '');
      out += r === undefined ? tok : r;
      i = j; continue;
    }
    out += ch; i++;
  }
  return out;
}

/** tên trong công thức (có chữ có dấu thanh) → dạng mã hóa cho bộ tính */
export function encFormulaNames(f: string): string {
  if (!HIGH.test(f)) return f;
  return mapIdents(f, (tok, next) => (HIGH.test(tok) && next !== '(' && next !== '!' ? encName(tok) : undefined));
}
/** dạng mã hóa → tên gốc (display: bảng tên hiển thị theo tên mã hóa viết hoa) */
export function decFormulaNames(f: string, display?: Map<string, string>): string {
  if (!hasEncoded(f)) return f;
  return mapIdents(f, (tok) => (hasEncoded(tok) ? display?.get(tok.toUpperCase()) ?? decName(tok) : undefined));
}

/** đổi tên trong công thức (không phân biệt hoa thường), bỏ qua chuỗi và tên trang tính */
export function renameInFormula(f: string, oldName: string, newName: string): string {
  const o = oldName.toUpperCase();
  return mapIdents(f, (tok, next, prev) => (tok.toUpperCase() === o && next !== '(' && next !== '!' && prev !== '!' ? newName : undefined));
}

/** Kiểm tra tên hợp lệ theo quy tắc Excel; trả về thông báo lỗi hoặc null */
export function nameError(name: string): string | null {
  const n = name.trim();
  if (!n) return 'Tên không được để trống.';
  if (n.length > 255) return 'Tên dài quá 255 ký tự.';
  if (!ID_START.test(n[0])) return 'Tên phải bắt đầu bằng chữ cái, dấu gạch dưới (_) hoặc dấu \\.';
  for (const ch of n) if (!ID_CHAR.test(ch)) return `Tên không được chứa “${ch === ' ' ? 'dấu cách' : ch}”.`;
  if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^[Rr]\d*[Cc]\d*$/.test(n) || /^[RrCc]$/.test(n)) return 'Tên không được trùng dạng địa chỉ ô (vd A1, R1C1).';
  if (/^(TRUE|FALSE)$/i.test(n)) return 'Tên này đã được dùng sẵn.';
  return null;
}
