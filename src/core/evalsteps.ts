// Đánh giá Công thức (Công thức ▸ Đánh giá Công thức): tính công thức từng bước như Excel.
// Mỗi bước: thay một tham chiếu ô/vùng bằng giá trị của nó, hoặc tính một lời gọi hàm trong cùng
// (khi các đối số đã là giá trị); bước cuối tính cả biểu thức. Vùng lớn (>2.000 ô) không thay bằng mảng
// mà để nguyên trong lời gọi hàm.
import { compileRefs, type Workbook, type XRef } from './workbook';

const REF_FN = new Set(['ROW', 'ROWS', 'COLUMN', 'COLUMNS', 'OFFSET', 'XOFS', 'ISREF', 'CELL', 'AREAS', 'INDIRECT', 'ISFORMULA', 'FORMULATEXT', 'SHEET', 'SHEETS']);

export interface Step { start: number; end: number; text: string; kind: 'ref' | 'call' | 'all'; ref?: XRef }

/** giá trị → chữ trong công thức (dấu thập phân “,”; mảng {a;b|c;d}) */
export function litOf(v: unknown): string {
  if (v === null || v === undefined || v === '') return v === '' ? '""' : '0';
  if (typeof v === 'number') return isFinite(v) ? v.toLocaleString('en-US', { useGrouping: false, maximumSignificantDigits: 15 }).replace('.', ',') : '#NUM!';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return '"' + v.replace(/"/g, '""') + '"';
  if (Array.isArray(v)) return '{' + v.map((row: unknown[]) => row.map((x) => litOf(x === null ? 0 : x)).join(';')).join('|') + '}';
  const e = (v as { value?: unknown })?.value;
  return typeof e === 'string' && e.startsWith('#') ? e : '#VALUE!';
}

/** các lời gọi hàm (vị trí tên → dấu ngoặc đóng), theo thứ tự đóng ngoặc: trong cùng, bên trái trước */
function calls(f: string): { start: number; end: number; name: string }[] {
  const out: { start: number; end: number; name: string }[] = [];
  const st: { start: number; name: string }[] = [];
  for (let i = 0; i < f.length; i++) {
    const ch = f[i];
    if (ch === '"' || ch === "'") { const q = ch; i++; while (i < f.length && !(f[i] === q && f[i + 1] !== q)) { if (f[i] === q) i++; i++; } continue; }
    if (ch === '(') {
      let k = i; while (k > 0 && /[A-Za-z0-9_.]/.test(f[k - 1])) k--;
      st.push({ start: k, name: f.slice(k, i).toUpperCase() });
    } else if (ch === ')') {
      const o = st.pop();
      if (o && o.name) out.push({ start: o.start, end: i + 1, name: o.name });
    }
  }
  return out;
}

/** bước kế tiếp cho biểu thức f (không có dấu =); null khi f đã là một giá trị */
export function nextStep(wb: Workbook, s: number, f: string): Step | null {
  const sheetOf = (n: string | undefined) => (n === undefined ? s : wb.sheetIdByName(n));
  const refsIn = (a: number, b: number) => {
    const toks = compileRefs(f, sheetOf);
    let pos = 0;
    for (const t of toks) {
      const len = typeof t === 'string' ? t.length : t.text.length;
      if (typeof t !== 'string' && pos >= a && pos + len <= b && (t.r2 - t.r1 + 1) * (t.c2 - t.c1 + 1) <= 2000) return { t, pos };
      pos += len;
    }
    return null;
  };
  const cs = calls(f);
  for (const c of cs) {
    // hàm cần chính tham chiếu (ROW, OFFSET…): không thay ô bằng giá trị, tính cả lời gọi
    if (REF_FN.has(c.name)) return { start: c.start, end: c.end, text: f.slice(c.start, c.end), kind: 'call' };
    const r = refsIn(c.start, c.end);
    if (r) return { start: r.pos, end: r.pos + r.t.text.length, text: r.t.text, kind: 'ref', ref: r.t };
    return { start: c.start, end: c.end, text: f.slice(c.start, c.end), kind: 'call' };
  }
  const r = refsIn(0, f.length);
  if (r) return { start: r.pos, end: r.pos + r.t.text.length, text: r.t.text, kind: 'ref', ref: r.t };
  // chỉ còn hằng/phép tính: tính cả biểu thức (nếu chưa phải một giá trị đơn)
  if (/^\s*(-?\d+(,\d+)?(E[+-]?\d+)?|"(?:[^"]|"")*"|TRUE|FALSE|#[A-Z/0!?]+[!?]?|\{[^}]*\})\s*$/i.test(f)) return null;
  return { start: 0, end: f.length, text: f, kind: 'all' };
}

/** thực hiện bước: trả về biểu thức mới */
export function applyStep(wb: Workbook, s: number, f: string, st: Step): string {
  let val: unknown;
  if (st.kind === 'ref' && st.ref) {
    const g = st.ref;
    if (g.r1 === g.r2 && g.c1 === g.c2) {
      val = wb.value(g.s, g.r1, g.c1);
      // ô trống: là "" khi ghép chuỗi (&), là 0 khi tính số — như Excel
      if (val === null || val === undefined) { const before = f.slice(0, st.start).trimEnd().slice(-1), after = f.slice(st.end).trimStart()[0]; if (before === '&' || after === '&') val = ''; }
    }
    else {
      const rows: unknown[][] = [];
      for (let r = g.r1; r <= g.r2; r++) { const row: unknown[] = []; for (let c = g.c1; c <= g.c2; c++) row.push(wb.value(g.s, r, c)); rows.push(row); }
      val = rows;
    }
  } else {
    try { val = wb.hf.calculateFormula('=' + st.text, s); } catch { val = { value: '#VALUE!' }; }
    if (val && typeof val === 'object' && !Array.isArray(val) && 'value' in (val as object)) val = { value: (val as { value: unknown }).value };
  }
  return f.slice(0, st.start) + litOf(val) + f.slice(st.end);
}
