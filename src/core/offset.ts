// OFFSET với đối số động (vd. OFFSET(K12;0;MATCH(...);1;2)).
// HyperFormula chỉ nhận OFFSET khi các số hàng/cột/cao/rộng là hằng số; nếu không → lỗi phân tích (#ERROR!).
// Cách làm: khi đưa công thức vào bộ tính, đổi OFFSET động thành hàm nội bộ XOFSn(...; <vùng phụ thuộc>)
// (n = số đối số gốc) và đổi ngược lại khi hiển thị / lưu tệp, nên người dùng luôn thấy OFFSET như Excel.
// <vùng phụ thuộc> là khung bao các ô mà OFFSET nhiều khả năng trỏ tới (ô gốc + các vùng xuất hiện trong đối số):
// nó giúp bộ tính sắp đúng thứ tự tính và tính lại khi bảng tra đổi. Hàm còn là "volatile" (giống Excel)
// nên ô đích nằm ngoài khung vẫn đọc được giá trị.
import { FunctionPlugin, FunctionArgumentType, CellError, ErrorType, SimpleRangeValue } from 'hyperformula';

const NAME = /^XOFS[345]$/;
const REF = /^(?:((?:'(?:[^']|'')+'|[^\s'!;()"]+)!))?(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?$/;
const NUM = /^\s*[-+]?\d+(?:[.,]\d+)?\s*$/;

function colNum(s: string) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
function colName(n: number) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/** Tách lời gọi hàm bắt đầu ở vị trí `open` (dấu "(") → danh sách đối số + vị trí ")" */
function splitCall(f: string, open: number): { args: string[]; close: number } | null {
  let depth = 0, cur = '', i = open + 1;
  const args: string[] = [];
  for (; i < f.length; i++) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; } j++; }
      cur += f.slice(i, j + 1); i = j; continue;
    }
    if (ch === '(' || ch === '{') depth++;
    else if (ch === ')' || ch === '}') {
      if (depth === 0 && ch === ')') { args.push(cur); return { args, close: i }; }
      depth--;
    } else if (ch === ';' && depth === 0) { args.push(cur); cur = ''; continue; }
    cur += ch;
  }
  return null;
}

/** Tìm tên hàm (không phân biệt hoa thường) ngoài chuỗi; gọi cb cho từng lời gọi, cb trả về chuỗi thay thế hoặc null */
function mapCalls(f: string, name: RegExp, cb: (fn: string, args: string[]) => string | null): string {
  let out = '', i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1); i = j + 1; continue;
    }
    if (/[A-Za-z_]/.test(ch) && !/[A-Za-z0-9_.$!]/.test(f[i - 1] ?? '')) {
      const m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(f.slice(i))!;
      const fn = m[0];
      if (f[i + fn.length] === '(' && name.test(fn)) {
        const sp = splitCall(f, i + fn.length);
        if (sp) {
          const rep = cb(fn, sp.args);
          if (rep !== null) { out += rep; i = sp.close + 1; continue; }
        }
      }
      out += fn; i += fn.length; continue;
    }
    out += ch; i++;
  }
  return out;
}

type Box = { r1: number; c1: number; r2: number; c2: number };
const parseRef = (s: string) => {
  const m = REF.exec(s.trim());
  if (!m) return null;
  const r1 = +m[5] - 1, c1 = colNum(m[3]);
  const r2 = m[9] ? +m[9] - 1 : r1, c2 = m[7] ? colNum(m[7]) : c1;
  return { sheet: m[1] ?? '', box: { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) } };
};
const sheetKey = (p: string) => p.replace(/!$/, '').replace(/^'(.*)'$/, '$1').replace(/''/g, "'").toLowerCase();

/**
 * Đổi OFFSET động → XOFSn. `self` = ô chứa công thức (để khung phụ thuộc không chứa chính nó → tránh vòng lặp),
 * `selfSheet` = tên sheet của ô đó.
 */
export function offsetIn(f: string, self?: { r: number; c: number }, selfSheet = ''): string {
  if (typeof f !== 'string' || f[0] !== '=' || !/OFFSET\s*\(/i.test(f)) return f;
  return mapCalls(f, /^OFFSET$/i, (_fn, raw) => {
    const args = raw.map((a) => offsetIn('=' + a, self, selfSheet).slice(1));
    if (args.length < 3 || args.length > 5) return null;
    if (args.slice(1).every((a) => a.trim() === '' || NUM.test(a))) return null; // HyperFormula tự xử lý
    const ref = parseRef(args[0]);
    if (!ref) return null;
    // khung bao: ô gốc + các vùng cùng sheet trong đối số
    const b: Box = { ...ref.box };
    const refText = /(?:'(?:[^']|'')+'!|[^\s'!;()"=<>&+\-*/^,]+!)?\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?/g;
    for (const a of args.slice(1)) {
      for (const m of a.match(refText) ?? []) {
        const p = parseRef(m);
        if (!p || !m.includes(':') || sheetKey(p.sheet) !== sheetKey(ref.sheet)) continue; // chỉ lấy vùng (ô đơn thường là giá trị tra)
        b.r1 = Math.min(b.r1, p.box.r1); b.c1 = Math.min(b.c1, p.box.c1);
        b.r2 = Math.max(b.r2, p.box.r2); b.c2 = Math.max(b.c2, p.box.c2);
      }
    }
    const num = (a: string | undefined) => (a !== undefined && NUM.test(a) ? Math.trunc(+a.replace(',', '.')) : null);
    const dr = num(args[1]), dc = num(args[2]);
    if (dr !== null) b.r2 = Math.max(b.r2, ref.box.r2 + dr), b.r1 = Math.max(0, Math.min(b.r1, ref.box.r1 + dr));
    if (dc !== null) b.c2 = Math.max(b.c2, ref.box.c2 + dc), b.c1 = Math.max(0, Math.min(b.c1, ref.box.c1 + dc));
    const h = num(args[3]) ?? (args[3] ? 1 : ref.box.r2 - ref.box.r1 + 1);
    const w = num(args[4]) ?? (args[4] ? 1 : ref.box.c2 - ref.box.c1 + 1);
    b.r2 += Math.max(0, h - 1); b.c2 += Math.max(0, w - 1);
    let bound = `${ref.sheet}$${colName(b.c1)}$${b.r1 + 1}:$${colName(b.c2)}$${b.r2 + 1}`;
    const same = sheetKey(ref.sheet) === '' || sheetKey(ref.sheet) === sheetKey(selfSheet);
    if (self && same && self.r >= b.r1 && self.r <= b.r2 && self.c >= b.c1 && self.c <= b.c2) bound = args[0].trim(); // tránh vòng lặp
    return `XOFS${args.length}(${args.join(';')};${bound})`;
  });
}

/** Đổi ngược XOFSn → OFFSET (bỏ đối số khung phụ thuộc) */
export function offsetOut(f: string): string {
  if (typeof f !== 'string' || !/XOFS[345]\s*\(/i.test(f)) return f;
  return mapCalls(f, NAME, (_fn, args) => `OFFSET(${args.slice(0, -1).map((a) => offsetOut(a)).join(';')})`);
}

// ---------- hàm trong bộ tính ----------
const P = { argumentType: FunctionArgumentType.ANY };
export class DynamicOffsetPlugin extends FunctionPlugin {
  static implementedFunctions = {
    XOFS3: { method: 'xofs', parameters: [P, P, P, P], isVolatile: true },
    XOFS4: { method: 'xofs', parameters: [P, P, P, P, P], isVolatile: true },
    XOFS5: { method: 'xofs', parameters: [P, P, P, P, P, P], isVolatile: true },
  };
  xofs(ast: any, state: any) {
    const self = this as any;
    const a = ast.args as any[];
    const n = a.length - 1; // bỏ khung phụ thuộc
    const ref = a[0];
    let start: any, h0 = 1, w0 = 1;
    if (ref?.type === 'CELL_REFERENCE') start = ref.reference.toSimpleCellAddress(state.formulaAddress);
    else if (ref?.type === 'CELL_RANGE') {
      const s = ref.start.toSimpleCellAddress(state.formulaAddress), e = ref.end.toSimpleCellAddress(state.formulaAddress);
      start = { sheet: s.sheet, row: Math.min(s.row, e.row), col: Math.min(s.col, e.col) };
      h0 = Math.abs(e.row - s.row) + 1; w0 = Math.abs(e.col - s.col) + 1;
    } else return new CellError(ErrorType.VALUE);
    // khung phụ thuộc: đánh giá để bộ tính biết thứ tự (giá trị không dùng)
    self.evaluateAst(a[n], state);
    const numArg = (i: number, def: number): number | CellError => {
      if (i >= n || a[i]?.type === 'EMPTY') return def;
      let v = self.evaluateAst(a[i], state);
      if (v instanceof SimpleRangeValue) v = (v as any).data?.[0]?.[0];
      if (v instanceof CellError) return v;
      const x = self.coerceScalarToNumberOrError(v);
      return x instanceof CellError ? x : Math.trunc(typeof x === 'number' ? x : (x as any).val ?? Number(x));
    };
    const dr = numArg(1, 0), dc = numArg(2, 0), h = numArg(3, h0), w = numArg(4, w0);
    for (const x of [dr, dc, h, w]) if (x instanceof CellError) return x;
    const R = start.row + (dr as number), C = start.col + (dc as number), H = h as number, W = w as number;
    if (H === 0 || W === 0) return new CellError(ErrorType.REF);
    const r1 = Math.min(R, R + H + (H < 0 ? 1 : -1)), c1 = Math.min(C, C + W + (W < 0 ? 1 : -1));
    const rows = Math.abs(H), cols = Math.abs(W);
    if (r1 < 0 || c1 < 0) return new CellError(ErrorType.REF);
    const g = self.dependencyGraph;
    const read = (row: number, col: number) => {
      try { const v = g.getScalarValue({ sheet: start.sheet, row, col }); return v === undefined ? null : v; } catch { return null; }
    };
    if (rows === 1 && cols === 1) { const v = read(r1, c1); return v === null || typeof v === 'symbol' ? 0 : v; }
    const data: any[][] = [];
    for (let i = 0; i < rows; i++) { const row: any[] = []; for (let j = 0; j < cols; j++) row.push(read(r1 + i, c1 + j)); data.push(row); }
    return SimpleRangeValue.onlyValues(data as any);
  }
}
