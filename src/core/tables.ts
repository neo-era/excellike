// Bảng (Excel Table, Ctrl+T) và tham chiếu có cấu trúc: Bang1[Thành tiền], [@SL], Bang1[[#Totals],[Tiền]]…
//
// HyperFormula không hiểu tham chiếu có cấu trúc, nên ở lớp vá (workbook.ts › patchOffset) công thức được đổi:
//   • vùng cột/bảng  → tên ẩn XT_<mã bảng>_<vùng>[_<mã cột đầu>_<mã cột cuối>] (named expression tuyệt đối).
//     Khi bảng nở/co chỉ cần sửa biểu thức của tên, mọi công thức tự tính lại; đổi tên bảng/cột không phải sửa công thức.
//   • [@Cột] (hàng này) → XTHISROW($C5): tham chiếu cột tuyệt đối, hàng tương đối nên sao chép/điền vẫn đúng.
// Khi đọc ra, các dạng mã này được dịch ngược theo bảng hiện hành (tên bảng, tên cột mới nhất).
import type { Range } from './address';
import type { Style } from './workbook';

export interface TableCol { id: number; name: string; fn?: TotalFn; label?: string }
export type TotalFn = 'sum' | 'average' | 'count' | 'counta' | 'max' | 'min' | 'stdev' | 'var' | 'none';
export interface TableDef {
  id: number; name: string;
  r1: number; c1: number; r2: number; c2: number;   // gồm hàng tiêu đề và hàng tổng
  header: boolean; totals: boolean;
  style: string;                                    // TableStyleMedium2…
  bandR: boolean; bandC: boolean; first: boolean; last: boolean; btn: boolean;
  cols: TableCol[];
}
/** bảng kèm tên sheet — để dịch công thức ở lớp bộ tính */
export interface TableReg extends TableDef { sheet: number; sheetName: string }
export type Area = 'D' | 'A' | 'H' | 'T' | 'HD' | 'DT';

export const dataRows = (t: TableDef) => ({ a: t.r1 + (t.header ? 1 : 0), b: t.r2 - (t.totals ? 1 : 0) });
export const dataRange = (t: TableDef): Range => { const d = dataRows(t); return { r1: d.a, c1: t.c1, r2: d.b, c2: t.c2 }; };
export const fullRange = (t: TableDef): Range => ({ r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 });
export const inTable = (t: TableDef, r: number, c: number) => r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2;

const quote = (n: string) => (/^[A-Za-z_][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
const colL = (c: number) => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const colN = (L: string) => { let n = 0; for (const ch of L.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const A1 = (sheet: string, g: Range) => `${quote(sheet)}!$${colL(g.c1)}$${g.r1 + 1}:$${colL(g.c2)}$${g.r2 + 1}`;

/** vùng ô của một phần bảng; null khi không tồn tại (vd hàng tổng đang tắt) */
export function areaRange(t: TableDef, area: Area, ca?: number, cb?: number): Range | null {
  const i1 = ca === undefined ? 0 : t.cols.findIndex((x) => x.id === ca);
  const i2 = cb === undefined ? t.cols.length - 1 : t.cols.findIndex((x) => x.id === cb);
  if (i1 < 0 || i2 < 0) return null;
  const c1 = t.c1 + Math.min(i1, i2), c2 = t.c1 + Math.max(i1, i2);
  const d = dataRows(t);
  switch (area) {
    case 'D': return { r1: d.a, c1, r2: d.b, c2 };
    case 'A': return { r1: t.r1, c1, r2: t.r2, c2 };
    case 'H': return t.header ? { r1: t.r1, c1, r2: t.r1, c2 } : null;
    case 'T': return t.totals ? { r1: t.r2, c1, r2: t.r2, c2 } : null;
    case 'HD': return { r1: t.r1, c1, r2: d.b, c2 };
    case 'DT': return { r1: d.a, c1, r2: t.r2, c2 };
  }
}
export const xtName = (id: number, area: Area, ca?: number, cb?: number) => `XT_${id}_${area}` + (ca !== undefined ? `_${ca}_${cb ?? ca}` : '');
export function xtExpr(t: TableReg, area: Area, ca?: number, cb?: number): string {
  const g = areaRange(t, area, ca, cb);
  return g ? '=' + A1(t.sheetName, g) : '=#REF!';
}

// ---------- đọc tham chiếu có cấu trúc ----------
const SPECIAL: Record<string, string> = { '#data': 'D', '#all': 'A', '#headers': 'H', '#totals': 'T', '#this row': 'R' };
const unesc = (s: string) => s.replace(/'(.)/g, '$1');
const esc = (s: string) => s.replace(/(['#[\]@])/g, "'$1");
/** tách nội dung một cặp [...] (đã bỏ ngoặc ngoài): trả về các mục, dấu nối giữa chúng */
function items(inner: string): { list: string[]; seps: string[] } | null {
  const list: string[] = [], seps: string[] = [];
  let i = 0;
  while (i < inner.length) {
    while (inner[i] === ' ') i++;
    if (inner[i] !== '[') return null;
    let j = i + 1, s = '';
    while (j < inner.length && inner[j] !== ']') { if (inner[j] === "'") { s += inner[j + 1] ?? ''; j += 2; continue; } s += inner[j]; j++; }
    if (inner[j] !== ']') return null;
    list.push(s.trim() === s ? s : s.trim());
    i = j + 1;
    while (inner[i] === ' ') i++;
    if (i < inner.length) { if (!/[,;:]/.test(inner[i])) return null; seps.push(inner[i]); i++; }
  }
  return { list, seps };
}
interface Parsed { area: Area | 'R'; c1?: string; c2?: string }
function parseInner(inner: string): Parsed | null {
  const t = inner.trim();
  if (!t) return { area: 'D' };
  if (t[0] === '@') {
    const rest = t.slice(1).trim();
    if (!rest) return null;
    if (rest[0] === '[') { const it = items(rest); if (!it || !it.list.length) return null; return { area: 'R', c1: it.list[0], c2: it.list[it.list.length - 1] }; }
    return { area: 'R', c1: unesc(rest) };
  }
  if (t[0] === '#') { const a = SPECIAL[t.toLowerCase()]; return a ? { area: a as Area | 'R' } : null; }
  if (t[0] !== '[') return { area: 'D', c1: unesc(t) };
  const it = items(t);
  if (!it || !it.list.length) return null;
  const sp = new Set<string>(); const cols: string[] = [];
  it.list.forEach((x) => { if (x[0] === '#') { const a = SPECIAL[x.toLowerCase()]; if (a) sp.add(a); else cols.push('\u0000'); } else cols.push(x); });
  if (cols.includes('\u0000') || cols.length > 2) return null;
  let area: Area | 'R' = 'D';
  if (sp.has('R')) area = 'R';
  else if (sp.has('A')) area = 'A';
  else if (sp.has('H') && sp.has('D')) area = 'HD';
  else if (sp.has('D') && sp.has('T')) area = 'DT';
  else if (sp.has('H')) area = 'H';
  else if (sp.has('T')) area = 'T';
  return { area, c1: cols[0], c2: cols[1] };
}
const IDENT = /[A-Za-z0-9_.\\À-￿]/;

export interface EncodeCtx {
  tables: TableReg[];
  self?: { sheet: string; r: number; c: number };
  need: (name: string, expr: string) => void;
  /** chỉ tính tạm (hộp Đối số Hàm, F9 khi sửa, kiểm tra cú pháp): thay bằng địa chỉ ô, không tạo tên ẩn */
  inline?: boolean;
}
/** công thức có tham chiếu có cấu trúc → dạng bộ tính hiểu được; trả về nguyên văn nếu không có */
export function encodeStructured(f: string, ctx: EncodeCtx): string {
  if (f[0] !== '=' || !f.includes('[') || !ctx.tables.length) return f;
  let out = '', i = 0;
  const byName = (n: string) => ctx.tables.find((t) => t.name.toLowerCase() === n.toLowerCase());
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length) { if (f[j] === ch && f[j + 1] === ch) { j += 2; continue; } if (f[j] === ch) break; j++; }
      out += f.slice(i, j + 1); i = j + 1; continue;
    }
    if (ch !== '[') { out += ch; i++; continue; }
    // tìm ngoặc đóng cân bằng (dấu ' thoát ký tự kế)
    let depth = 0, j = i;
    for (; j < f.length; j++) {
      if (f[j] === "'") { j++; continue; }
      if (f[j] === '[') depth++;
      else if (f[j] === ']') { depth--; if (!depth) break; }
    }
    if (j >= f.length) { out += f.slice(i); break; }
    const inner = f.slice(i + 1, j);
    // tên bảng đứng ngay trước
    let k = out.length; while (k > 0 && IDENT.test(out[k - 1])) k--;
    const tname = out.slice(k);
    // [1]Sheet!A1 (liên kết ngoài) không phải bảng
    if (/^\d+$/.test(inner) && !tname) { out += f.slice(i, j + 1); i = j + 1; continue; }
    const t = tname ? byName(tname) : ctx.tables.find((x) => ctx.self && x.sheetName === ctx.self.sheet && inTable(x, ctx.self.r, ctx.self.c));
    const p = t ? parseInner(inner) : null;
    if (!t || !p) { out += f.slice(i, j + 1); i = j + 1; continue; }
    const id = (n?: string) => (n === undefined ? undefined : t.cols.find((x) => x.name.toLowerCase() === n.toLowerCase())?.id ?? -1);
    const a = id(p.c1), b = p.c2 !== undefined ? id(p.c2) : a;
    let rep: string;
    if (a === -1 || b === -1) rep = '#REF!';
    else if (p.area === 'R') {
      const ia = t.cols.findIndex((x) => x.id === (a ?? t.cols[0].id)), ib = t.cols.findIndex((x) => x.id === (b ?? t.cols[t.cols.length - 1].id));
      const d = dataRows(t);
      if (!ctx.self || ctx.self.r < d.a || ctx.self.r > d.b || ia !== ib) rep = '#VALUE!';
      else {
        const sh = ctx.self.sheet === t.sheetName ? '' : quote(t.sheetName) + '!';
        rep = `XTHISROW(${sh}$${colL(t.c1 + ia)}${ctx.self.r + 1})`;
      }
    } else {
      if (ctx.inline) rep = xtExpr(t, p.area, a, b).slice(1);
      else {
        const n = xtName(t.id, p.area, a, b);
        ctx.need(n, xtExpr(t, p.area, a, b));
        rep = n;
      }
    }
    out = out.slice(0, k) + rep;
    i = j + 1;
  }
  return out;
}

export type DecodeMode = 'ui' | 'file' | 'a1';
const XT_RE = /\bXT_(\d+)_(HD|DT|D|A|H|T)(?:_(\d+)_(\d+))?\b/gi;
const XR_RE = /XTHISROW\(((?:'(?:[^']|'')+'|[^\s'!;()"]+)!)?\$([A-Z]{1,3})\$?(\d+)\)/gi;
const colRef = (n: string) => '[' + esc(n) + ']';
/** dạng mã → tham chiếu có cấu trúc (ui: như Excel hiển thị; file: dạng lưu trong .xlsx; a1: đổi thành địa chỉ ô) */
export function decodeStructured(f: string, tables: TableReg[], at?: { sheet: number; r: number; c: number }, mode: DecodeMode = 'ui', only?: number): string {
  if (!/XT_|XTHISROW/i.test(f)) return f;
  // không đụng chuỗi "…" trong công thức
  if (f.includes('"')) return f.split(/("(?:[^"]|"")*")/).map((p, i) => (i % 2 ? p : decodeStructured(p, tables, at, mode, only))).join('');
  const own = (t: TableReg) => mode === 'ui' && !!at && at.sheet === t.sheet && inTable(t, at.r, at.c);
  f = f.replace(XT_RE, (m, sid, area: Area, ca, cb) => {
    if (only !== undefined && +sid !== only) return m;
    const t = tables.find((x) => x.id === +sid);
    if (!t) return '#REF!';
    area = area.toUpperCase() as Area;
    const a = ca !== undefined ? +ca : undefined, b = cb !== undefined ? +cb : undefined;
    if ((a !== undefined && !t.cols.some((x) => x.id === a)) || (b !== undefined && !t.cols.some((x) => x.id === b))) return '#REF!';
    if (mode === 'a1') { const g = areaRange(t, area, a, b); return g ? A1(t.sheetName, g) : '#REF!'; }
    const nm = (id?: number) => t.cols.find((x) => x.id === id)?.name ?? '';
    const pre = own(t) ? '' : t.name;
    const sep = mode === 'file' ? ',' : ';';
    const specs = ({ D: ['#Data'], A: ['#All'], H: ['#Headers'], T: ['#Totals'], HD: ['#Headers', '#Data'], DT: ['#Data', '#Totals'] } as Record<Area, string[]>)[area];
    const sp = specs.map((x) => `[${x}]`).join(sep);
    if (a === undefined) return specs.length === 1 ? `${pre}[${specs[0]}]` : `${pre}[${sp}]`;
    const cols = a === b ? colRef(nm(a)) : `${colRef(nm(a))}:${colRef(nm(b))}`;
    if (area === 'D') return a === b ? `${pre}${cols}` : `${pre}[${cols}]`;
    return `${pre}[${sp}${sep}${cols}]`;
  });
  f = f.replace(XR_RE, (m, sh, L, row) => {
    const sheetName = sh ? sh.slice(0, -1).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : undefined;
    const c = colN(L), r = +row - 1;
    const t = tables.find((x) => (sheetName ? x.sheetName === sheetName : at ? x.sheet === at.sheet : false) && c >= x.c1 && c <= x.c2 && r >= dataRows(x).a && r <= dataRows(x).b);
    if (!t || (mode === 'a1' && (only === undefined || t.id === only))) return `${sh ?? ''}$${L}${row}`;
    if (mode === 'a1') return m;
    const n = t.cols[c - t.c1]?.name ?? '';
    if (mode === 'file') return `${t.name}[[#This Row],${colRef(n)}]`;
    const body = /^[\p{L}\p{N}_]+$/u.test(n) ? esc(n) : colRef(n);
    return `${own(t) ? '' : t.name}[@${body}]`;
  });
  return f;
}

// ---------- tên bảng ----------
export function tableNameError(n: string, taken: (n: string) => boolean): string | null {
  if (!n) return 'Tên bảng không được để trống.';
  if (n.length > 255) return 'Tên bảng quá dài.';
  if (!/^[A-Za-z_\\À-￿][\w.À-￿]*$/.test(n)) return 'Tên bảng phải bắt đầu bằng chữ cái hoặc dấu gạch dưới và không chứa dấu cách, ký tự đặc biệt.';
  if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^[RC]\d*$/i.test(n) || /^R\d+C\d+$/i.test(n)) return 'Tên bảng không được trông giống địa chỉ ô.';
  if (taken(n)) return 'Tên này đã được dùng cho một bảng hoặc tên vùng khác.';
  return null;
}

// ---------- kiểu bảng ----------
const ACCENTS = ['#000000', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const mix = (hex: string, w: number) => {
  const n = parseInt(hex.slice(1), 16), r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  const f = (x: number) => Math.round(x + (255 - x) * w).toString(16).padStart(2, '0');
  return '#' + f(r) + f(g) + f(b);
};
const dark = (hex: string, w: number) => {
  const n = parseInt(hex.slice(1), 16), f = (x: number) => Math.round(x * (1 - w)).toString(16).padStart(2, '0');
  return '#' + f((n >> 16) & 255) + f((n >> 8) & 255) + f(n & 255);
};
export interface TableLook { head: Style; band: string | null; body?: string; total: Style; line: string; headLine?: string }
/** bảng màu kiểu bảng: Light/Medium/Dark 1–7 theo màu chủ đề như Excel */
export function tableLook(style: string): TableLook | null {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(style);
  if (!m) return null;
  const fam = m[1], idx = (+m[2] - 1) % 7, acc = ACCENTS[idx], grey = idx === 0;
  if (fam === 'Light') return { head: { b: true, fg: grey ? '#000000' : dark(acc, 0.25) }, band: mix(grey ? '#808080' : acc, 0.8), total: { b: true }, line: grey ? '#000000' : acc, headLine: grey ? '#000000' : acc };
  if (fam === 'Dark') return { head: { b: true, fg: '#FFFFFF', bg: grey ? '#000000' : dark(acc, 0.5) }, band: grey ? '#404040' : dark(acc, 0.25), body: grey ? '#737373' : acc, total: { b: true, fg: '#FFFFFF', bg: grey ? '#000000' : dark(acc, 0.5) }, line: '#FFFFFF' };
  return { head: { b: true, fg: '#FFFFFF', bg: acc }, band: mix(grey ? '#808080' : acc, 0.8), total: { b: true }, line: grey ? '#000000' : acc };
}
export const TABLE_STYLES = (['Light', 'Medium', 'Dark'] as const).flatMap((f) => [1, 2, 3, 4, 5, 6, 7].map((i) => `TableStyle${f}${i}`));
/** định dạng hiển thị do bảng tạo cho ô (dưới định dạng riêng của ô) */
export function tableCellStyle(t: TableDef, r: number, c: number): Style | null {
  const L = tableLook(t.style);
  if (!L || !inTable(t, r, c)) return null;
  const d = dataRows(t);
  const fc = t.first && c === t.c1, lc = t.last && c === t.c2;
  if (t.header && r === t.r1) return { ...L.head, bd: L.headLine ? { b: { s: 'thin', c: L.headLine } } : undefined };
  if (t.totals && r === t.r2) return { ...L.total, bd: { t: { s: 'double', c: L.line } } };
  const st: Style = {};
  const k = r - d.a, j = c - t.c1;
  if (L.body) { st.bg = L.body; st.fg = '#FFFFFF'; }
  if (t.bandR && k % 2 === 0 && L.band) st.bg = L.band;
  if (t.bandC && j % 2 === 0 && L.band) st.bg = L.band;
  if (fc || lc) st.b = true;
  if (r === d.b && !t.totals && L.headLine) st.bd = { b: { s: 'thin', c: L.headLine } };
  return st;
}
/** công thức hàng tổng cho một hàm (SUBTOTAL 1xx bỏ qua hàng ẩn như Excel) */
export const TOTAL_FNS: [TotalFn, string, number][] = [
  ['none', 'Không có', 0], ['average', 'Trung bình', 101], ['count', 'Đếm', 103], ['counta', 'Đếm Số', 102], ['max', 'Lớn nhất', 104],
  ['min', 'Nhỏ nhất', 105], ['sum', 'Tổng', 109], ['stdev', 'Độ lệch Chuẩn', 107], ['var', 'Phương sai', 110],
];
export const totalFormula = (fn: TotalFn, col: string) => { const x = TOTAL_FNS.find((y) => y[0] === fn); return x && x[2] ? `=SUBTOTAL(${x[2]};[${esc(col)}])` : null; };
