import { DynamicOffsetPlugin, offsetIn, offsetOut } from './offset';
import { DynArrayPlugin, DYN_NAMES } from './dynarray';
import { areaRange, type Area, encodeStructured, decodeStructured, tableCellStyle, tableNameError, dataRows, inTable, xtExpr, totalFormula, type TableDef, type TableReg, type TotalFn } from './tables';
import { encFormulaNames, decFormulaNames, encName, nameError, renameInFormula } from './names';
import { HyperFormula, FunctionPlugin, FunctionArgumentType, CellError, ErrorType } from 'hyperformula';
import { shiftFormula, fixSheetQuotes, needsQuote, quoteSheet, transposeFormula } from './formula';
import type { ConfigParams, RawCellContent, SimpleCellAddress } from 'hyperformula';
import { MAX_COLS, MAX_ROWS, type Range, rangeSize, toA1 } from './address';
import { normalizeInput } from './input';
import { editString, formatValue, isDateFmt, type Align } from './numfmt';
import { AxisSizes } from './sizes';
import { fillPlan, seriesValues, type SrcCell, type SeriesOpts } from './series';
import { toTSV, parseTSV } from './tsv';
import { parseCSV, toCSV } from './csv';
import { parseRef, parseRange } from './address';
import type { Chart, Ref } from './chart';
import { computePivot, type PivotSpec } from './pivot';
import { cfEval, cfStats, dvDescribe, inR, numOf, compare, type CFRule, type CFResult, type CFStats, type DVRule } from './rules';

export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
export interface Border { s: BorderStyle; c?: string }
export interface Style {
  b?: boolean; i?: boolean; u?: boolean; s?: boolean;
  fg?: string; bg?: string; al?: Align; fmt?: string;
  fn?: string; fs?: number;          // phông chữ, cỡ chữ (pt)
  wr?: boolean; va?: 't' | 'm' | 'b'; // xuống dòng, căn dọc
  bd?: { t?: Border; r?: Border; b?: Border; l?: Border; d?: Border; u?: Border }; // d: chéo xuống (↘), u: chéo lên (↗)
  ind?: number;                       // thụt lề (bậc)
  ud?: boolean;                       // gạch chân kép (đi cùng u)
  vp?: 'sup' | 'sub';                 // chỉ số trên / dưới
  rot?: number;                       // xoay chữ: -90…90 độ, 255 = chữ dọc
  shk?: boolean;                      // co chữ vừa ô
  hx?: 'fill' | 'just' | 'cas' | 'dist'; // căn ngang mở rộng: Điền, Căn đều, Giữa qua vùng chọn, Phân phối
  nt?: string;                        // Ghi chú ô (Shift+F2) — đi theo ô khi chèn/xóa/sắp xếp/sao chép như liên kết
  cm?: CommentItem[];                 // Chú thích dạng chuỗi (Ctrl+Shift+F2): bình luận đầu + các trả lời
  cmDone?: boolean;                   // chuỗi chú thích đã giải quyết
  lk?: false;                         // bỏ khóa ô (mặc định ô bị khóa khi bảo vệ trang)
  hid?: boolean;                      // ẩn công thức khi bảo vệ trang
  ln?: string;                        // siêu kết nối (URL, mailto:, hoặc #Sheet!A1)
}
export interface PasteOpts {
  what?: 'all' | 'formulas' | 'values' | 'formats' | 'comments' | 'validation' | 'theme' | 'noBorders' | 'colWidths' | 'formulasNum' | 'valuesNum' | 'allMerge';
  op?: 'none' | 'add' | 'sub' | 'mul' | 'div';
  skipBlanks?: boolean; transpose?: boolean; link?: boolean;
}
export interface FilterCond { values?: string[]; text?: string }
export interface SheetFilter { range: Range; conds: Record<number, FilterCond> }
export interface PivotDef extends PivotSpec { id?: string; src: Ref; at: { r: number; c: number }; area?: Range | null }
export type PaperName = 'A3' | 'A4' | 'A5' | 'B4' | 'B5' | 'Letter' | 'Legal' | 'Tabloid';
/** Lề trang (inch, như Excel): trên, dưới, trái, phải, đầu trang, chân trang */
export interface Margins { t: number; b: number; l: number; r: number; h: number; f: number }
export interface PrintSetup {
  orientation: 'portrait' | 'landscape'; paper: PaperName; gridlines: boolean;
  fitWidth?: boolean; titleRows?: number;           // dạng cũ (0.3–0.4)
  margins?: 'normal' | 'narrow' | 'wide' | 'custom'; mg?: Margins; area?: Range | null;
  scale?: number; fit?: { w: number; h: number } | null;   // % hoặc vừa w × h trang (0 = tự động)
  titleRowRange?: [number, number] | null; titleColRange?: [number, number] | null;
  headings?: boolean; center?: { h?: boolean; v?: boolean }; firstPage?: number | null; order?: 'down' | 'over';
  header?: string; footer?: string;                 // mã Excel &L &C &R &P &N…
}
export interface SheetMeta {
  styles: Map<number, Style>;
  colW: AxisSizes;
  rowH: AxisSizes;
  freeze: { rows: number; cols: number };
  merges: Range[];
  cf: CFRule[];
  dv: DVRule[];
  filter: SheetFilter | null;
  charts: Chart[];
  pivots: PivotDef[];
  print: PrintSetup | null;
  breaks: { rows: number[]; cols: number[] };   // ngắt trang thủ công (trước hàng/cột)
  tabColor?: string;
  hidden?: boolean;                              // trang tính ẩn
  /** Nhóm hàng/cột (Dữ liệu ▸ Đại cương); bất biến: mỗi thay đổi tạo đối tượng mới (dễ hoàn tác) */
  outline: Outline;
  /** Bảng (Ctrl+T); bất biến như outline */
  tables: TableDef[];
}
/** cấp nhóm (1–7) theo hàng/cột; x: các nhóm đang thu gọn, khóa "r|c:cấp:hàng/cột tổng" */
export interface Outline { r: Record<number, number>; c: Record<number, number>; x: string[] }
const NO_OUTLINE: Outline = { r: {}, c: {}, x: [] };
export type BorderKind = 'all' | 'outer' | 'none' | 'top' | 'bottom' | 'left' | 'right' | 'inner';
export interface SortKey { col: number; desc?: boolean }
const COLLATOR = new Intl.Collator('vi', { sensitivity: 'accent' });
interface Entry { hf: boolean; undo?: () => void; redo?: () => void; parts?: Entry[] }
interface Clip { sheet: number; range: Range; cut: boolean }
export interface Changed { r: number; c: number }

const NO_STYLE: Style = Object.freeze({}) as Style;
export const DEFAULT_COL_W = 64;   // 8,43 ký tự như Excel
export const DEFAULT_ROW_H = 20;   // 15 pt
const HISTORY = 300;

export const HF_CONFIG: Partial<ConfigParams> = {
  licenseKey: 'gpl-v3',
  functionArgSeparator: ';',
  decimalSeparator: ',',
  thousandSeparator: '',
  arrayColumnSeparator: ';', // hằng mảng trong app: {1;2,5|3;4} — cột “;”, hàng “|”, thập phân “,” (tệp: {1,2.5;3,4})
  arrayRowSeparator: '|',
  dateFormats: ['DD/MM/YYYY', 'DD/MM/YY'],
  timeFormats: ['hh:mm', 'hh:mm:ss'],
  currencySymbol: ['₫', 'đ', '$'],
  language: 'enGB',
  undoLimit: HISTORY,
  maxRows: MAX_ROWS,
  maxColumns: MAX_COLS,
  useColumnIndex: false,
  smartRounding: true,
  precisionRounding: 15, // giống Excel: 15 chữ số có nghĩa (mặc định của HyperFormula chỉ 10)
};

// INT của HyperFormula cắt phần thập phân; Excel làm tròn xuống (INT(-2,5) = -3)
class ExcelIntPlugin extends FunctionPlugin {
  static implementedFunctions = {
    INT: { method: 'int', parameters: [{ argumentType: FunctionArgumentType.NUMBER }] },
    MOD: { method: 'mod', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }] },
  };
  int(ast: any, state: any) { return (this as any).runFunction(ast.args, state, (this as any).metadata('INT'), (x: number) => Math.floor(x)); }
  // MOD của Excel lấy dấu theo số chia: MOD(-7;3) = 2
  mod(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('MOD'), (n: number, d: number) =>
      d === 0 ? new CellError(ErrorType.DIV_BY_ZERO) : n - d * Math.floor(n / d));
  }
}
/** XTHISROW(ô): giá trị ô cùng hàng của cột bảng — mã nội bộ của [@Cột] */
class TableRowPlugin extends FunctionPlugin {
  static implementedFunctions = { XTHISROW: { method: 'xthisrow', parameters: [{ argumentType: FunctionArgumentType.SCALAR }] } };
  xthisrow(ast: any, state: any) { return (this as any).runFunction(ast.args, state, (this as any).metadata('XTHISROW'), (x: unknown) => x); }
}
// FORECAST / FORECAST.LINEAR / INTERCEPT (HyperFormula chưa có) — hồi quy tuyến tính như Excel
const numOfCell = (v: any): number | null => (typeof v === 'number' ? v : v && typeof v === 'object' && typeof v.val === 'number' ? v.val : null);
function pairs(ys: any, xs: any): [number[], number[]] | CellError {
  const a = ys.valuesFromTopLeftCorner ? ys.valuesFromTopLeftCorner() : [ys];
  const b = xs.valuesFromTopLeftCorner ? xs.valuesFromTopLeftCorner() : [xs];
  if (a.length !== b.length) return new CellError(ErrorType.NA);
  const Y: number[] = [], X: number[] = [];
  for (let i = 0; i < a.length; i++) {
    if (a[i] instanceof CellError) return a[i];
    if (b[i] instanceof CellError) return b[i];
    const y = numOfCell(a[i]), x = numOfCell(b[i]);
    if (y !== null && x !== null) { Y.push(y); X.push(x); }
  }
  return [Y, X];
}
function linfit(ys: any, xs: any): { a: number; b: number } | CellError {
  const p = pairs(ys, xs);
  if (p instanceof CellError) return p;
  const [Y, X] = p, n = Y.length;
  if (n < 2) return new CellError(ErrorType.DIV_BY_ZERO);
  const mx = X.reduce((s, v) => s + v, 0) / n, my = Y.reduce((s, v) => s + v, 0) / n;
  let sxy = 0, sxx = 0;
  for (let i = 0; i < n; i++) { sxy += (X[i] - mx) * (Y[i] - my); sxx += (X[i] - mx) ** 2; }
  if (sxx === 0) return new CellError(ErrorType.DIV_BY_ZERO);
  const b = sxy / sxx;
  return { a: my - b * mx, b };
}
class ExcelForecastPlugin extends FunctionPlugin {
  static implementedFunctions = {
    FORECAST: { method: 'forecast', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }] },
    'FORECAST.LINEAR': { method: 'forecast', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }] },
    INTERCEPT: { method: 'intercept', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }] },
  };
  forecast(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('FORECAST'), (x: number, ys: any, xs: any) => { const f = linfit(ys, xs); return f instanceof CellError ? f : f.a + f.b * x; });
  }
  intercept(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('INTERCEPT'), (ys: any, xs: any) => { const f = linfit(ys, xs); return f instanceof CellError ? f : f.a; });
  }
}
let pluginsReady = false;
function ensurePlugins() {
  if (pluginsReady) return;
  HyperFormula.unregisterFunction('INT');
  HyperFormula.unregisterFunction('MOD');
  HyperFormula.registerFunctionPlugin(ExcelIntPlugin as any, { enGB: { INT: 'INT', MOD: 'MOD' } } as any);
  HyperFormula.registerFunctionPlugin(ExcelForecastPlugin as any, { enGB: { FORECAST: 'FORECAST', 'FORECAST.LINEAR': 'FORECAST.LINEAR', INTERCEPT: 'INTERCEPT' } } as any);
  HyperFormula.unregisterFunction('FILTER');
  HyperFormula.registerFunctionPlugin(DynArrayPlugin as any, { enGB: Object.fromEntries(DYN_NAMES.map((n) => [n, n])) } as any);
  HyperFormula.registerFunctionPlugin(TableRowPlugin as any, { enGB: { XTHISROW: 'XTHISROW' } } as any);
  HyperFormula.registerFunctionPlugin(DynamicOffsetPlugin as any, { enGB: { XOFS3: 'XOFS3', XOFS4: 'XOFS4', XOFS5: 'XOFS5' } } as any);
  pluginsReady = true;
}
/** OFFSET động: đổi công thức khi vào bộ tính và đổi ngược khi đọc ra (xem offset.ts) */
/** công thức vào bộ tính: OFFSET động + tên có dấu thanh */
/** bối cảnh bảng cho lớp vá: danh sách bảng hiện hành + cách tạo tên ẩn XT_… */
interface TblCtx { tables: () => TableReg[]; need: (name: string, expr: string) => void; inline?: boolean }
const fin = (v: string, self: { r: number; c: number } | undefined, sheet: string, ctx?: TblCtx) => {
  if (v[0] !== '=') return v;
  let x = offsetIn(v, self, sheet);
  if (ctx && x.includes('[')) x = encodeStructured(x, { tables: ctx.tables(), self: self && { sheet, r: self.r, c: self.c }, need: ctx.need, inline: ctx.inline });
  return encFormulaNames(x);
};
const NEEDS_FIN = /OFFSET\s*\(|[^\u0000-\u02AF]|\[/i;
function mapRows(cells: RawCellContent[][], sheet: string, r0 = 0, c0 = 0, ctx?: TblCtx): RawCellContent[][] {
  let hit = false;
  for (const row of cells) { if (row) for (const v of row) if (typeof v === 'string' && v[0] === '=' && NEEDS_FIN.test(v)) { hit = true; break; } if (hit) break; }
  if (!hit) return cells;
  return cells.map((row, i) => row && row.map((v, j) => (typeof v === 'string' ? (fin(v, { r: r0 + i, c: c0 + j }, sheet, ctx) as RawCellContent) : v)));
}
/** bảng tên hiển thị (tên gốc) theo tên mã hóa viết hoa, gắn vào từng bộ tính */
export const nameDisplay = (hf: HyperFormula): Map<string, string> => ((hf as any).__nameDisp ??= new Map<string, string>());
/** bảng hiện hành gắn với bộ tính (Workbook gán hàm lấy danh sách) */
const hfTables = (hf: HyperFormula): TableReg[] => ((hf as any).__tables?.() ?? []);
function tblCtx(hf: HyperFormula): TblCtx {
  return {
    tables: () => hfTables(hf),
    need: (n, e) => { try { if (hf.getNamedExpression(n) === undefined) hf.addNamedExpression(n, e); } catch { /* tên đã có */ } },
  };
}
function patchOffset(hf: HyperFormula): HyperFormula {
  const h: any = hf;
  const set = h.setCellContents.bind(hf), sheetSet = h.setSheetContent.bind(hf);
  const gf = h.getCellFormula.bind(hf), gs = h.getCellSerialized.bind(hf), gss = h.getSheetSerialized.bind(hf);
  h.setCellContents = (a: any, c: any) => {
    const name = hf.getSheetName(a.sheet) ?? '';
    const ctx = tblCtx(hf);
    if (Array.isArray(c)) return set(a, mapRows(c, name, a.row, a.col, ctx));
    return set(a, typeof c === 'string' ? fin(c, { r: a.row, c: a.col }, name, ctx) : c);
  };
  const out = (f: string, at?: { sheet: number; r: number; c: number }) => decodeStructured(decFormulaNames(offsetOut(f), nameDisplay(hf)), hfTables(hf), at);
  const calc = h.calculateFormula.bind(hf);
  h.calculateFormula = (f: string, sheet: number) => calc(fin(f, undefined, hf.getSheetName(sheet) ?? '', { ...tblCtx(hf), inline: true }), sheet);
  const vf = h.validateFormula.bind(hf);
  h.validateFormula = (f: string) => vf(f.includes('[') ? encodeStructured(f, { tables: hfTables(hf), need: () => {}, inline: true }).replace(/(?<![\w'!])\[(?:[^[\]]|\[[^\]]*\])*\]/g, '1') : f);
  h.setSheetContent = (id: number, cells: RawCellContent[][]) => sheetSet(id, mapRows(cells, hf.getSheetName(id) ?? '', 0, 0, tblCtx(hf)));
  h.__rawFormula = (a: any) => gf(a);
  h.getCellFormula = (a: any) => { const f = gf(a); return f === undefined ? f : out(f, { sheet: a.sheet, r: a.row, c: a.col }); };
  h.getCellSerialized = (a: any) => { const v = gs(a); return typeof v === 'string' && v[0] === '=' ? out(v, { sheet: a.sheet, r: a.row, c: a.col }) : v; };
  h.getSheetSerialized = (id: number) => gss(id).map((row: any[], i: number) => row.map((v, j) => (typeof v === 'string' && v[0] === '=' ? out(v, { sheet: id, r: i, c: j }) : v)));
  return hf;
}
// Excel cho phép viết TRUE/FALSE không có ngoặc
const NAMES = [{ name: 'TRUE', expression: '=TRUE()' }, { name: 'FALSE', expression: '=FALSE()' }];
function buildHF(sheets?: { name: string; cells: RawCellContent[][] }[], tables: TableReg[] = []) {
  ensurePlugins();
  // buildFromSheets nhanh hơn ~2,5 lần khi mở sổ lớn; thứ tự key của object phải giữ đúng thứ tự sheet
  if (sheets && sheets.length && !sheets.some((x) => /^\d+$/.test(x.name))) {
    const obj: Record<string, RawCellContent[][]> = Object.create(null);
    const extra = new Map<string, string>();
    const ctx: TblCtx = { tables: () => tables, need: (n, e) => { extra.set(n, e); } };
    for (const x of sheets) obj[x.name] = mapRows(x.cells, x.name, 0, 0, ctx);
    const hf = patchOffset(HyperFormula.buildFromSheets(obj, HF_CONFIG, [...NAMES, ...[...extra].map(([name, expression]) => ({ name, expression }))]));
    (hf as any).__tables = () => tables;
    return hf;
  }
  const hf = patchOffset(HyperFormula.buildEmpty(HF_CONFIG));
  (hf as any).__tables = () => tables;
  for (const n of NAMES) hf.addNamedExpression(n.name, n.expression);
  if (sheets) {
    hf.suspendEvaluation();
    for (const x of sheets) hf.addSheet(x.name);
    for (const x of sheets) if (x.cells.length) hf.setSheetContent(hf.getSheetId(x.name)!, x.cells);
    hf.resumeEvaluation();
  }
  return hf;
}

const key = (r: number, c: number) => r * MAX_COLS + c;
const unkey = (k: number) => ({ r: Math.floor(k / MAX_COLS), c: k % MAX_COLS });
const BAD_NAME = /[\\/?*[\]:]/;

export class Workbook {
  hf!: HyperFormula;
  private metas = new Map<number, SheetMeta>();
  private past: Entry[] = [];
  private future: Entry[] = [];
  private listeners = new Set<() => void>();
  clip: Clip | null = null;
  name = 'Sổ tính 1';

  private constructor(hf: HyperFormula) { this.attach(hf); }
  /** gắn bộ tính + danh sách bảng sống cho lớp dịch tham chiếu có cấu trúc */
  private attach(hf: HyperFormula) { this.hf = hf; (hf as any).__tables = () => this.tableReg(); }

  static create(): Workbook {
    const wb = new Workbook(buildHF());
    const n = wb.hf.addSheet('Sheet1');
    wb.metas.set(wb.hf.getSheetId(n)!, newMeta());
    wb.hf.clearUndoStack();
    return wb;
  }

  // ---------- sự kiện ----------
  onChange(f: () => void) { this.listeners.add(f); return () => this.listeners.delete(f); }
  /** Tăng mỗi khi sổ thay đổi — dùng làm khóa bộ nhớ đệm (CF) */
  ver = 0;
  private emit() { this.ver++; if (this.iterate.on) this.resolveIter(); this.listeners.forEach((f) => f()); }

  // ---------- lịch sử ----------
  private hfDepth(): number | null {
    const st = (this.hf as any)?._crudOperations?.undoRedo?.undoStack;
    return Array.isArray(st) ? st.length : null;
  }
  /**
   * Ghi một thao tác vào lịch sử. hfOp chạy trước; nếu nó lỗi thì không có gì thay đổi.
   * post chạy sau khi hfOp thành công (đổi định dạng/kích thước đi kèm).
   */
  private record(hfOp: (() => unknown) | null, custom?: { undo: () => void; redo: () => void }, opt: { batch?: boolean; post?: () => void } = {}): unknown[] {
    let hfStep = false;
    let changes: unknown[] = [];
    if (hfOp) {
      const before = this.hfDepth();
      const res = opt.batch === false ? hfOp() : this.hf.batch(hfOp);
      if (Array.isArray(res)) changes = res;
      const after = this.hfDepth();
      hfStep = before === null || after === null ? true : after > before || (after === before && after === HISTORY);
    }
    opt.post?.();
    if (!hfStep && !custom) return changes;
    this.past.push({ hf: hfStep, ...custom });
    if (this.past.length > HISTORY) this.past.shift();
    this.future = [];
    this.emit();
    return changes;
  }
  /** Tính lại toàn bộ (F9). HyperFormula xóa lịch sử của nó khi dựng lại nên lịch sử chung cũng phải làm mới. */
  recalc() {
    this.hf.rebuildAndRecalculate();
    this.past = []; this.future = [];
    this.emit();
  }
  /** Xóa lịch sử Hoàn tác (sau khi dựng sổ từ mẫu) */
  clearHistory() { this.past = []; this.future = []; this.hf.clearUndoStack(); this.emit(); }
  canUndo() { return this.past.length > 0; }
  canRedo() { return this.future.length > 0; }
  undo() {
    const e = this.past.pop();
    if (!e) return false;
    const un = (x: Entry) => { if (x.parts) { [...x.parts].reverse().forEach(un); return; } x.undo?.(); if (x.hf && this.hf.isThereSomethingToUndo()) this.hf.undo(); };
    un(e);
    this.future.push(e);
    this.emit();
    return true;
  }
  redo() {
    const e = this.future.pop();
    if (!e) return false;
    const re = (x: Entry) => { if (x.parts) { x.parts.forEach(re); return; } if (x.hf && this.hf.isThereSomethingToRedo()) this.hf.redo(); x.redo?.(); };
    re(e);
    this.past.push(e);
    this.emit();
    return true;
  }

  // ---------- sheet ----------
  /** Sheet người dùng thấy (bỏ sheet đệm giá trị của liên kết ngoài “[n]Tên”) */
  sheetIds(): number[] {
    const ids = this.hf.getSheetNames().filter((n) => !isExtName(n)).map((n) => this.hf.getSheetId(n)!);
    if (!this.order.length) return ids;
    const pos = new Map(this.order.map((id, i) => [id, i]));
    return ids.map((id, i) => [id, pos.get(id) ?? this.order.length + i] as const).sort((a, b) => a[1] - b[1]).map((x) => x[0]);
  }
  /** Thứ tự sheet khi đã di chuyển (rỗng = theo thứ tự tạo) */
  private order: number[] = [];
  /** Các trang tính đang hiện (thẻ sheet) */
  visibleSheetIds(): number[] { return this.sheetIds().filter((id) => !this.meta(id).hidden); }
  isSheetHidden(id: number) { return !!this.meta(id).hidden; }
  /** Mọi sheet trong bộ tính, kể cả sheet đệm liên kết ngoài */
  allSheetIds(): number[] { return this.hf.getSheetNames().map((n) => this.hf.getSheetId(n)!); }
  /** Liên kết sang tệp khác (giữ nguyên phần XML để ghi lại khi lưu .xlsx) */
  extLinks: ExtLink[] = [];
  sheetName(id: number) { return this.hf.getSheetName(id)!; }
  meta(id: number): SheetMeta {
    let m = this.metas.get(id);
    if (!m) { m = newMeta(); this.metas.set(id, m); }
    return m;
  }
  private checkName(name: string, except?: number) {
    const t = name.trim();
    if (!t) throw new Error('Tên sheet không được để trống.');
    if (t.length > 31) throw new Error('Tên sheet tối đa 31 ký tự.');
    if (BAD_NAME.test(t)) throw new Error('Tên sheet không được chứa các ký tự \\ / ? * [ ] :');
    if (t.startsWith("'") || t.endsWith("'")) throw new Error('Tên sheet không được bắt đầu hoặc kết thúc bằng dấu nháy đơn.');
    for (const id of this.sheetIds()) {
      if (id !== except && this.sheetName(id).toLowerCase() === t.toLowerCase()) throw new Error(`Đã có sheet tên “${t}”.`);
    }
    return t;
  }
  /** Thêm trang tính (index: vị trí trong thứ tự; bỏ trống = cuối) — một bước hoàn tác */
  addSheet(name?: string, index?: number): number {
    let n = name;
    if (!n) { let i = this.sheetIds().length + 1; while (this.hasName('Sheet' + i)) i++; n = 'Sheet' + i; }
    n = this.checkName(n);
    let id = -1;
    let old: MetaSnap | null = null;
    const fresh = snapMeta(newMeta());
    const before = this.sheetIds(), oOrder = [...this.order];
    let nOrder: number[] = [];
    this.record(() => { this.hf.addSheet(n); id = this.hf.getSheetId(n)!; }, {
      undo: () => { if (old) restoreMeta(this.meta(id), old); this.order = oOrder; },
      redo: () => { restoreMeta(this.meta(id), fresh); this.order = nOrder; },
    }, { post: () => {
      const m = this.meta(id); old = snapMeta(m); restoreMeta(m, fresh);
      // mã sheet có thể được HyperFormula dùng lại: luôn đặt rõ vị trí trong thứ tự
      const ids = before.filter((x) => x !== id);
      ids.splice(index === undefined ? ids.length : Math.max(0, Math.min(index, ids.length)), 0, id);
      this.order = nOrder = ids;
    } });
    return id;
  }
  private hasName(n: string) { return this.sheetIds().some((id) => this.sheetName(id).toLowerCase() === n.toLowerCase()); }
  renameSheet(id: number, name: string) {
    const t = this.checkName(name, id);
    if (t === this.sheetName(id)) return;
    this.record(() => this.hf.renameSheet(id, t));
  }
  removeSheet(id: number) {
    if (this.sheetIds().length <= 1) throw new Error('Sổ tính phải còn ít nhất một sheet.');
    const snap = snapMeta(this.meta(id));
    const xBefore = this.snapRefs();
    let xAfter = xBefore;
    const oOrder = [...this.order], nOrder = this.sheetIds().filter((x) => x !== id);
    this.record(() => this.hf.removeSheet(id), {
      undo: () => { restoreMeta(this.meta(id), snap); this.restoreRefs(xBefore); this.order = oOrder; },
      redo: () => { this.restoreRefs(xAfter); this.order = nOrder; },
    }, { post: () => { this.dropRefsTo(id); xAfter = this.snapRefs(); this.order = nOrder; } });
    if (this.clip?.sheet === id) this.clip = null;
  }

  // ---------- thuộc tính trang tính: màu thẻ, ẩn, thứ tự, sao chép ----------
  private metaProp<K extends 'tabColor' | 'hidden'>(id: number, k: K, v: SheetMeta[K]) {
    const m = this.meta(id), o = m[k];
    if (o === v) return;
    m[k] = v;
    this.record(null, { undo: () => { this.meta(id)[k] = o; }, redo: () => { this.meta(id)[k] = v; } });
  }
  setTabColor(id: number, color: string | undefined) { this.metaProp(id, 'tabColor', color || undefined); }
  setSheetHidden(id: number, hide: boolean) {
    if (hide && this.visibleSheetIds().filter((x) => x !== id).length === 0) throw new Error('Sổ làm việc phải còn ít nhất một trang tính hiển thị.');
    this.metaProp(id, 'hidden', hide || undefined);
  }
  moveSheet(id: number, index: number) {
    const before = this.sheetIds();
    const after = before.filter((x) => x !== id);
    after.splice(Math.max(0, Math.min(index, after.length)), 0, id);
    if (after.join() === before.join()) return;
    const o = [...this.order];
    this.order = after;
    this.record(null, { undo: () => { this.order = o; }, redo: () => { this.order = after; } });
  }
  /** Tạo bản sao (Di chuyển hoặc Sao chép… ▸ Tạo bản sao) */
  copySheet(id: number, index?: number): number {
    const base = this.sheetName(id).replace(/ \(\d+\)$/, '');
    let k = 2, n = '';
    do { const suf = ` (${k++})`; n = base.slice(0, 31 - suf.length) + suf; } while (this.hasName(n));
    const cells = this.hf.getSheetSerialized(id);
    const src = snapMeta(this.meta(id));
    let nid = -1;
    const oOrder = [...this.order];
    let nOrder: number[] = [];
    let fresh: MetaSnap | null = null;
    // bảng của bản sao: mã và tên mới (Bảng1 → Bảng2…) — gán trước khi nạp ô để [@Cột] dịch đúng
    const tcopy = src.tables?.map((t) => ({ ...t, cols: t.cols.map((c) => ({ ...c })) })) ?? [];
    this.record(() => {
      this.hf.addSheet(n); nid = this.hf.getSheetId(n)!;
      for (const t of tcopy) { t.id = ++this.tableSeq; this.meta(nid).tables = tcopy; t.name = this.nextTableName(); }
      this.meta(nid).tables = tcopy;
      this.hf.setSheetContent(nid, cells);
    }, {
      undo: () => { this.order = oOrder; },
      redo: () => { if (fresh) restoreMeta(this.meta(nid), fresh); this.order = nOrder; },
    }, {
      post: () => {
        const m = this.meta(nid);
        const sm = this.meta(id);
        m.colW = new AxisSizes(sm.colW.def, MAX_COLS); m.rowH = new AxisSizes(sm.rowH.def, MAX_ROWS);
        restoreMeta(m, { ...src, styles: new Map([...src.styles].map(([k2, v]) => [k2, { ...v }])), merges: src.merges.map((g) => ({ ...g })), charts: [], pivots: [], tables: tcopy });
        // biểu đồ trỏ vào chính sheet gốc → trỏ sang bản sao
        m.charts = this.meta(id).charts.map((ch) => ({ ...ch, id: this.uid(), series: ch.series.map((x) => { const mv = (r?: Ref) => (r && r.sheet === id ? { ...r, sheet: nid } : r); return { ...x, vals: mv(x.vals)!, name: mv(x.name), cats: mv(x.cats) }; }) }));
        fresh = snapMeta(m);
        const ids = this.sheetIds().filter((x) => x !== nid);
        ids.splice(index ?? ids.indexOf(id) + 1, 0, nid);
        this.order = nOrder = ids;
      },
    });
    return nid;
  }

  // ---------- ẩn / bỏ ẩn hàng, cột ----------
  private hideAxis(s: number, ax: 'r' | 'c', a: number, b: number, hide: boolean) {
    const m = this.meta(s), A = ax === 'r' ? m.rowH : m.colW;
    const before = A.hidden(), sizes = A.entries();
    const set = new Set(before);
    for (let i = a; i <= b; i++) { if (hide) set.add(i); else { set.delete(i); if (A.base(i) === 0) A.set(i, A.def); } }
    const after = [...set].sort((x, y) => x - y), sizesAfter = A.entries();
    A.setHidden(after);
    this.record(null, { undo: () => { A.load(sizes, before); }, redo: () => { A.load(sizesAfter, after); } });
  }
  hideRows(s: number, r1: number, r2: number, hide: boolean) { this.hideAxis(s, 'r', r1, r2, hide); }
  hideCols(s: number, c1: number, c2: number, hide: boolean) { this.hideAxis(s, 'c', c1, c2, hide); }
  isColHidden(s: number, c: number) { const A = this.meta(s).colW; return A.isHidden(c) || A.base(c) === 0; }

  // ---------- ngắt trang ----------
  private breaksChange(s: number, f: (b: { rows: number[]; cols: number[] }) => { rows: number[]; cols: number[] }) {
    const m = this.meta(s), before = m.breaks;
    const n = f({ rows: [...before.rows], cols: [...before.cols] });
    const after = { rows: [...new Set(n.rows)].filter((x) => x > 0).sort((a, b) => a - b), cols: [...new Set(n.cols)].filter((x) => x > 0).sort((a, b) => a - b) };
    if (JSON.stringify(after) === JSON.stringify(before)) return;
    m.breaks = after;
    this.record(null, { undo: () => { this.meta(s).breaks = before; }, redo: () => { this.meta(s).breaks = after; } });
  }
  /** Như Excel: chèn ngắt trang phía trên hàng và bên trái cột của ô hiện hành */
  insertPageBreak(s: number, r: number, c: number) { this.breaksChange(s, (b) => ({ rows: r > 0 ? [...b.rows, r] : b.rows, cols: c > 0 ? [...b.cols, c] : b.cols })); }
  removePageBreak(s: number, r: number, c: number) { this.breaksChange(s, (b) => ({ rows: b.rows.filter((x) => x !== r), cols: b.cols.filter((x) => x !== c) })); }
  resetPageBreaks(s: number) { this.breaksChange(s, () => ({ rows: [], cols: [] })); }

  // ---------- chèn / xóa ô có dồn ----------
  insertCells(s: number, g: Range, dir: 'down' | 'right') { this.shiftCells(s, g, dir === 'down' ? 'r' : 'c', 1); }
  deleteCells(s: number, g: Range, dir: 'up' | 'left') { this.shiftCells(s, g, dir === 'up' ? 'r' : 'c', -1); }
  private shiftCells(s: number, g: Range, ax: 'r' | 'c', sign: 1 | -1) {
    const m = this.meta(s), u = this.used(s);
    const n = ax === 'r' ? g.r2 - g.r1 + 1 : g.c2 - g.c1 + 1;
    // dải bị dồn: cùng các cột (dồn dọc) hoặc cùng các hàng (dồn ngang) từ vùng chọn trở đi
    const band = (x: Range) => (ax === 'r' ? x.c2 >= g.c1 && x.c1 <= g.c2 && x.r2 >= g.r1 : x.r2 >= g.r1 && x.r1 <= g.r2 && x.c2 >= g.c1);
    const inside = (x: Range) => (ax === 'r' ? x.c1 >= g.c1 && x.c2 <= g.c2 && x.r1 >= g.r1 : x.r1 >= g.r1 && x.r2 <= g.r2 && x.c1 >= g.c1);
    for (const mg of m.merges) if (band(mg) && (!inside(mg) || (sign < 0 && (ax === 'r' ? mg.r1 <= g.r2 : mg.c1 <= g.c2)))) throw new Error('Không thể thực hiện thao tác này với ô đã gộp. Hãy bỏ gộp ô trước.');
    for (const p of m.pivots) if (p.area && band(p.area)) throw new Error('Không thể chèn/xóa ô làm dịch chuyển PivotTable.');
    const last = ax === 'r' ? u.rows - 1 : u.cols - 1;
    const max = ax === 'r' ? MAX_ROWS : MAX_COLS;
    if (sign > 0 && last + n >= max) throw new Error('Không thể chèn: dữ liệu sẽ bị đẩy ra ngoài trang tính.');
    const A = (r: number, c: number) => this.A(s, r, c);
    // Như Excel: chèn/xóa cả hàng (hoặc cột) để mọi tham chiếu dời đúng (vùng giãn ra, ô bị xóa → #REF!),
    // rồi đưa phần ngoài dải về chỗ cũ bằng moveCells (tham chiếu tới chúng đi theo).
    const lastR = u.rows - 1, lastC = u.cols - 1;
    const outside: [number, number][] = ax === 'r' ? [[0, g.c1 - 1], [g.c2 + 1, lastC]] : [[0, g.r1 - 1], [g.r2 + 1, lastR]];
    const blocks = outside.filter(([a, b]) => b >= a);
    const lo = ax === 'r' ? g.r1 : g.c1;
    const hfOp = () => {
      const blk = (a: number, b: number, from: number, to: number) => (ax === 'r' ? { start: A(from, a), end: A(to, b) } : { start: A(a, from), end: A(b, to) });
      const at = (a: number, pos: number) => (ax === 'r' ? A(pos, a) : A(a, pos));
      if (sign > 0) {
        if (ax === 'r') this.hf.addRows(s, [lo, n]); else this.hf.addColumns(s, [lo, n]);
        if (last >= lo) for (const [a, b] of blocks) this.hf.moveCells(blk(a, b, lo + n, last + n), at(a, lo));
      } else {
        if (last >= lo) {
          if (last + n >= max) throw new Error('Không thể xóa: trang tính quá đầy.');
          for (const [a, b] of blocks) this.hf.moveCells(blk(a, b, lo, last), at(a, lo + n));
        }
        if (ax === 'r') this.hf.removeRows(s, [lo, n]); else this.hf.removeColumns(s, [lo, n]);
      }
    };
    const before = snapMeta(m);
    let after = before;
    this.record(hfOp, { undo: () => restoreMeta(m, before), redo: () => restoreMeta(m, after) }, {
      post: () => {
        const hit = (r: number, c: number) => (ax === 'r' ? c >= g.c1 && c <= g.c2 && r >= lo : r >= g.r1 && r <= g.r2 && c >= lo);
        setAll(m.styles, remapStyles(m.styles, (r, c) => {
          if (!hit(r, c)) return { r, c };
          const v = ax === 'r' ? r : c;
          const nv = sign > 0 ? v + n : v < lo + n ? -1 : v - n;
          if (nv < 0) return null;
          return ax === 'r' ? { r: nv, c } : { r, c: nv };
        }));
        m.merges = m.merges.map((x) => (band(x) && inside(x) ? (ax === 'r' ? { ...x, r1: x.r1 + sign * n, r2: x.r2 + sign * n } : { ...x, c1: x.c1 + sign * n, c2: x.c2 + sign * n }) : x));
        // quy tắc CF / kiểm tra dữ liệu nằm gọn trong dải: dời theo (bị xóa hết thì bỏ)
        const mv = (x: Range): Range | null => {
          if (!(band(x) && inside(x))) return x;
          const a1 = ax === 'r' ? x.r1 : x.c1, a2 = ax === 'r' ? x.r2 : x.c2;
          let b1 = a1, b2 = a2;
          if (sign > 0) { b1 += n; b2 += n; }
          else { const end = lo + n; if (a2 < end) return null; b1 = a1 < end ? lo : a1 - n; b2 = a2 - n; }
          return ax === 'r' ? { ...x, r1: b1, r2: b2 } : { ...x, c1: b1, c2: b2 };
        };
        m.cf = m.cf.map((x) => { const r = mv(x.range); return r ? { ...x, range: r } : null; }).filter(Boolean) as CFRule[];
        m.dv = m.dv.map((x) => { const r = mv(x.range); return r ? { ...x, range: r } : null; }).filter(Boolean) as DVRule[];
        after = snapMeta(m);
      },
    });
  }

  // ---------- đọc ô ----------
  private A(s: number, r: number, c: number): SimpleCellAddress { return { sheet: s, row: r, col: c }; }
  value(s: number, r: number, c: number) {
    const v = this.hf.getCellValue(this.A(s, r, c));
    // như Excel: công thức trả về ô trống thì hiện 0 (=A1 khi A1 trống)
    // tính lặp: ô thuộc vòng tham chiếu (hoặc phụ thuộc vào nó) lấy giá trị đã hội tụ
    if (this.over.size) { const o = this.over.get(s + ':' + r + ':' + c); if (o !== undefined) return (o === null ? 0 : o) as typeof v; }
    if (v === null) return (this.hf.doesCellHaveFormula(this.A(s, r, c)) ? 0 : null) as typeof v;
    return v;
  }

  // ---------- tính lặp (File → Options → Formulas → Enable iterative calculation) ----------
  iterate = { on: false, count: 100, delta: 0.001 };
  private over = new Map<string, unknown>();
  private seeds = new Map<string, unknown>();
  setIterate(p: Partial<Workbook['iterate']>) {
    const before = { ...this.iterate }, after = { ...this.iterate, ...p };
    this.iterate = after;
    this.record(null, { undo: () => (this.iterate = { ...before }), redo: () => (this.iterate = { ...after }) });
    this.resolveIter();
  }
  /** Giá trị khởi đầu (vd giá trị Excel lưu sẵn trong tệp) để hội tụ ngay như Excel */
  setIterSeeds(list: [number, number, number, unknown][]) { for (const [s, r, c, v] of list) this.seeds.set(s + ':' + r + ':' + c, v); this.resolveIter(); }
  private resolveIter() {
    if (!this.iterate.on) { this.over.clear(); return; }
    type DC = { s: number; r: number; c: number; k: string; parts: (string | XRef)[] };
    const D: DC[] = [];
    const inD = new Set<string>();
    for (const s of this.allSheetIds()) {
      const vals = this.hf.getSheetValues(s);
      for (let r = 0; r < vals.length; r++) {
        const row = vals[r];
        for (let c = 0; c < row.length; c++) {
          const v: any = row[c];
          if (v && typeof v === 'object' && v.type === 'CYCLE') { const k = s + ':' + r + ':' + c; inD.add(k); D.push({ s, r, c, k, parts: [] }); }
        }
      }
    }
    if (!D.length) { this.over.clear(); return; }
    // ô phụ thuộc vòng lặp nhưng không hiện #CYCLE! (vd. bọc trong IFERROR) cũng phải tính lặp
    {
      const q: any[] = D.map((d) => this.A(d.s, d.r, d.c));
      const seenR = new Set<string>();
      while (q.length && D.length <= 20000) {
        const a = q.pop();
        let deps: any[] = [];
        try { deps = this.hf.getCellDependents(a) as any[]; } catch { deps = []; }
        for (const x of deps) {
          if (x.start) { const rk = JSON.stringify(x); if (!seenR.has(rk)) { seenR.add(rk); q.push(x); } continue; }
          const k = x.sheet + ':' + x.row + ':' + x.col;
          if (inD.has(k) || !this.hf.doesCellHaveFormula(x)) continue;
          inD.add(k); D.push({ s: x.sheet, r: x.row, c: x.col, k, parts: [] }); q.push(x);
        }
      }
    }
    if (D.length > 20000) { this.over.clear(); return; }
    for (const k of [...this.over.keys()]) if (!inD.has(k)) this.over.delete(k); // quá lớn: giữ #CYCLE! như cũ
    const sheetOf = (name: string | undefined, s: number) => (name === undefined ? s : this.hf.getSheetId(name));
    for (const d of D) d.parts = compileRefs(this.hf.getCellFormula(this.A(d.s, d.r, d.c)) ?? '', (n) => sheetOf(n, d.s));
    const cur = (s: number, r: number, c: number): unknown => {
      const k = s + ':' + r + ':' + c;
      if (inD.has(k)) return this.over.get(k) ?? this.seeds.get(k) ?? 0;
      return this.hf.getCellValue(this.A(s, r, c));
    };
    const build = (parts: (string | XRef)[]) => parts.map((p) => {
      if (typeof p === 'string') return p;
      let hit = false;
      for (let r = p.r1; r <= p.r2 && !hit; r++) for (let c = p.c1; c <= p.c2; c++) if (inD.has(p.s + ':' + r + ':' + c)) { hit = true; break; }
      if (!hit) return p.text;
      if (p.r1 === p.r2 && p.c1 === p.c2) return '(' + lit(cur(p.s, p.r1, p.c1)) + ')';
      const n = (p.r2 - p.r1 + 1) * (p.c2 - p.c1 + 1);
      if (n > 5000 || (p.r1 !== p.r2 && p.c1 !== p.c2)) return p.text; // vùng 2 chiều: chưa hỗ trợ
      const items: string[] = [];
      for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) items.push(lit(cur(p.s, r, c), true));
      const arr = '{' + items.join('|') + '}';
      return p.r1 === p.r2 ? 'TRANSPOSE(' + arr + ')' : arr;
    }).join('');
    for (const d of D) if (!this.over.has(d.k)) this.over.set(d.k, this.seeds.get(d.k) ?? 0);
    for (let it = 0; it < Math.max(1, this.iterate.count); it++) {
      let change = 0;
      for (const d of D) {
        let v: any;
        try { v = this.hf.calculateFormula(build(d.parts), d.s); } catch { v = this.hf.getCellValue(this.A(d.s, d.r, d.c)); }
        if (Array.isArray(v)) v = v[0]?.[0] ?? null;
        const old = this.over.get(d.k);
        const delta = typeof v === 'number' && typeof old === 'number' ? Math.abs(v - old) : v === old || (v && old && typeof v === 'object' && (v as any).type === (old as any).type) ? 0 : Infinity;
        if (delta > change) change = delta;
        this.over.set(d.k, v);
      }
      if (change < this.iterate.delta) break;
    }
  }
  style(s: number, r: number, c: number): Style { const m = this.meta(s).styles; return m.size ? m.get(key(r, c)) ?? NO_STYLE : NO_STYLE; }
  isFormula(s: number, r: number, c: number) { return this.hf.doesCellHaveFormula(this.A(s, r, c)); }
  /** Định dạng hiệu lực: style.fmt hoặc suy ra từ kiểu giá trị (ngày, %) */
  effFmt(s: number, r: number, c: number, st = this.style(s, r, c)): string | undefined {
    if (st.fmt) return st.fmt;
    const v = this.value(s, r, c);
    if (typeof v !== 'number') return undefined;
    const t = this.hf.getCellValueDetailedType(this.A(s, r, c)) as string;
    if (t === 'NUMBER_DATE') return 'dd/mm/yyyy';
    if (t === 'NUMBER_DATETIME') return 'dd/mm/yyyy hh:mm';
    if (t === 'NUMBER_TIME') return 'hh:mm';
    if (t === 'NUMBER_PERCENT') return '0%';
    if (t === 'NUMBER_CURRENCY') return '#,##0';
    return undefined;
  }
  display(s: number, r: number, c: number) {
    const st = this.style(s, r, c);
    const v = this.value(s, r, c);
    const f = formatValue(v, typeof v === 'number' ? this.effFmt(s, r, c, st) : undefined);
    const error = typeof v === 'object' && v !== null;
    return { text: f.text, align: st.al ?? f.align, style: st, error, value: v };
  }
  /** Chuỗi hiện trong thanh công thức / khi sửa ô */
  rawText(s: number, r: number, c: number): string {
    const a = this.A(s, r, c);
    const f = this.hf.getCellFormula(a);
    if (f !== undefined) return this.outF(f);
    const raw = this.hf.getCellSerialized(a);
    if (typeof raw === 'string') return raw;
    // ô nằm trong vùng tràn của mảng động: không có nội dung riêng (như Excel)
    if (raw === null || raw === undefined) { if (this.hf.isCellPartOfArray(a)) return ''; }
    else if (this.hf.isCellPartOfArray(a) && !this.hf.doesCellHaveFormula(a)) return '';
    return editString(this.value(s, r, c), this.effFmt(s, r, c));
  }
  /** HyperFormula bỏ dấu nháy quanh tên sheet như 2024, Q1, 1Q khi xuất công thức → thêm lại */
  private quoteNames(): string[] { return this.allSheetIds().map((i) => this.sheetName(i)).filter(needsQuote); }
  outF(f: string, names = this.quoteNames()): string { return names.length ? fixSheetQuotes(f, names) : f; }
  used(s: number) { const d = this.hf.getSheetDimensions(s); return { rows: d.height, cols: d.width }; }

  // ---------- ghi ô ----------
  private changedOn(s: number, ch: unknown[]): Changed[] {
    const out: Changed[] = [];
    for (const x of ch as any[]) if (x.address && x.address.sheet === s) out.push({ r: x.address.row, c: x.address.col });
    return out;
  }
  setInput(s: number, r: number, c: number, text: string): Changed[] {
    const { content, fmt } = normalizeInput(text);
    const k = key(r, c);
    const st = this.meta(s).styles;
    const old = st.get(k);
    const tx = fmt && !old?.fmt ? new StyleTxn([[st, k, { ...(old ?? {}), fmt }]]) : null;
    const changes = this.record(
      () => { this.hf.setCellContents(this.A(s, r, c), [[content as RawCellContent]]); },
      tx ? tx.hist : undefined,
      { post: () => tx?.apply() },
    );
    return this.changedOn(s, changes);
  }
  clear(s: number, g: Range) {
    const u = this.used(s);
    const r2 = Math.min(g.r2, u.rows - 1), c2 = Math.min(g.c2, u.cols - 1);
    if (r2 < g.r1 || c2 < g.c1) return;
    const rows = Array.from({ length: r2 - g.r1 + 1 }, () => Array(c2 - g.c1 + 1).fill(null));
    this.record(() => this.hf.setCellContents(this.A(s, g.r1, g.c1), rows));
  }

  // ---------- định dạng ----------
  private clampToUsed(s: number, g: Range): Range {
    const u = this.used(s);
    const r2 = g.r2 >= MAX_ROWS - 1 ? Math.max(g.r1, Math.min(g.r2, u.rows - 1)) : g.r2;
    const c2 = g.c2 >= MAX_COLS - 1 ? Math.max(g.c1, Math.min(g.c2, u.cols - 1)) : g.c2;
    return { ...g, r2, c2 };
  }
  setStyle(s: number, g0: Range, patch: Style) {
    const g = this.clampToUsed(s, g0);
    if (rangeSize(g) > 500_000) throw new Error('Vùng chọn quá lớn để định dạng (tối đa 500.000 ô).');
    const st = this.meta(s).styles;
    const before = new Map<number, Style | undefined>();
    const after = new Map<number, Style | undefined>();
    for (let r = g.r1; r <= g.r2; r++)
      for (let c = g.c1; c <= g.c2; c++) {
        const k = key(r, c);
        const o = st.get(k);
        before.set(k, o);
        after.set(k, clean({ ...(o ?? {}), ...patch }));
      }
    const apply = (m: Map<number, Style | undefined>) => m.forEach((v, k) => setOrDel(st, k, v));
    apply(after);
    this.record(null, { undo: () => apply(before), redo: () => apply(after) });
  }
  /** Định dạng ▸ Độ rộng Mặc định… (px) */
  setDefaultColWidth(s: number, px: number) {
    const A = this.meta(s).colW, old = A.def;
    if (old === px) return;
    const ent = A.entries(), hid = A.hidden();
    const apply = (d: number) => { A.def = d; A.load(ent, hid); };
    apply(px);
    this.record(null, { undo: () => apply(old), redo: () => apply(px) });
  }
  /** Đặt kích thước nhiều hàng/cột trong một bước hoàn tác */
  setSizes(s: number, ax: 'r' | 'c', list: [number, number][]) {
    const A = ax === 'r' ? this.meta(s).rowH : this.meta(s).colW;
    const before = A.entries(), hid = A.hidden();
    for (const [i, v] of list) A.set(i, v);
    const after = A.entries();
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    this.record(null, { undo: () => A.load(before, hid), redo: () => A.load(after, hid) });
  }
  setColWidth(s: number, c: number, w: number) { this.sizeChange(this.meta(s).colW, c, w); }
  setRowHeight(s: number, r: number, h: number) { this.sizeChange(this.meta(s).rowH, r, h); }
  private sizeChange(ax: AxisSizes, i: number, v: number) {
    const o = ax.base(i);
    if (o === Math.round(v)) return;
    ax.set(i, v);
    this.record(null, { undo: () => ax.set(i, o), redo: () => ax.set(i, v) });
  }
  setFreeze(s: number, rows: number, cols: number) {
    const m = this.meta(s);
    const o = { ...m.freeze }, n = { rows, cols };
    m.freeze = n;
    this.record(null, { undo: () => (m.freeze = o), redo: () => (m.freeze = n) });
  }

  // ---------- chèn / xóa ----------
  private structural(s: number, hfOp: () => void, shift: (m: SheetMeta) => void, ax: 'r' | 'c', span: (a: number, b: number) => { a: number; b: number } | null) {
    const m = this.meta(s);
    const before = snapMeta(m);
    const xBefore = this.snapRefs();
    let after = before, xAfter = xBefore;
    const hadTables = m.tables.length > 0;
    this.record(() => {
      hfOp();
      // bảng: dời vùng bảng rồi cập nhật tên ẩn XT_… ngay trong cùng bước (chèn ở hàng dữ liệu đầu/hàng tổng HF chỉ dời, không nở)
      if (hadTables) { shift(m); this.shiftRefs(s, ax, span); this.syncXT(); }
    }, {
      undo: () => { restoreMeta(m, before); this.restoreRefs(xBefore); },
      redo: () => { restoreMeta(m, after); this.restoreRefs(xAfter); },
    }, {
      post: () => { if (!hadTables) { shift(m); this.shiftRefs(s, ax, span); } after = snapMeta(m); xAfter = this.snapRefs(); },
    });
  }
  /** Xóa sheet: bỏ chuỗi biểu đồ và PivotTable trỏ tới sheet đó (mã sheet có thể bị dùng lại) */
  private dropRefsTo(id: number) {
    const ok = (r?: Ref) => !r || r.sheet !== id;
    for (const [sid, m] of this.metas) {
      if (sid === id) continue;
      m.charts = m.charts.map((ch) => {
        if (ch.series.every((x) => ok(x.vals) && ok(x.name) && ok(x.cats))) return ch;
        return { ...ch, series: ch.series.filter((x) => ok(x.vals)).map((x) => ({ ...x, name: ok(x.name) ? x.name : undefined, cats: ok(x.cats) ? x.cats : undefined })) };
      });
      m.pivots = m.pivots.filter((p) => p.src.sheet !== id);
    }
  }
  /** Biểu đồ/Pivot ở mọi sheet có thể tham chiếu tới sheet đang chèn/xóa → dời theo */
  private snapRefs() { return [...this.metas.entries()].map(([id, m]) => [id, m.charts, m.pivots] as const); }
  private restoreRefs(x: ReturnType<Workbook['snapRefs']>) { for (const [id, ch, pv] of x) { const m = this.metas.get(id); if (m) { m.charts = ch; m.pivots = pv; } } }
  private shiftRefs(s: number, ax: 'r' | 'c', span: (a: number, b: number) => { a: number; b: number } | null) {
    const mv = (g: Range): Range | null => {
      const x = ax === 'r' ? span(g.r1, g.r2) : span(g.c1, g.c2);
      return x ? (ax === 'r' ? { ...g, r1: x.a, r2: x.b } : { ...g, c1: x.a, c2: x.b }) : null;
    };
    const mvRef = (r?: Ref) => (r && r.sheet === s ? (mv(r.range) ? { ...r, range: mv(r.range)! } : null) : r);
    for (const m of this.metas.values()) {
      m.charts = m.charts.map((ch) => {
        const series = ch.series.map((x) => { const vals = mvRef(x.vals); return vals ? { ...x, vals, name: mvRef(x.name) ?? undefined, cats: mvRef(x.cats) ?? undefined } : null; }).filter(Boolean) as Chart['series'];
        return series.length === ch.series.length && series.every((x, i) => x === ch.series[i]) ? ch : { ...ch, series };
      });
      m.pivots = m.pivots.map((p) => (p.src.sheet === s ? (mvRef(p.src) ? { ...p, src: mvRef(p.src)! } : p) : p));
    }
  }
  /** Như Excel: không chèn/xóa hàng, cột cắt ngang một PivotTable */
  private guardPivot(s: number, ax: 'r' | 'c', at: number, n: number, del: boolean) {
    for (const p of this.meta(s).pivots) {
      const a = p.area;
      if (!a) continue;
      const lo = ax === 'r' ? a.r1 : a.c1, hi = ax === 'r' ? a.r2 : a.c2;
      const cut = del ? at <= hi && at + n - 1 >= lo && !(at <= lo && at + n - 1 >= hi) : at > lo && at <= hi;
      if (cut) throw new Error(`Không thể ${del ? 'xóa' : 'chèn'} ${ax === 'r' ? 'hàng' : 'cột'} cắt ngang PivotTable. Hãy di chuyển hoặc xóa PivotTable trước.`);
    }
  }
  insertRows(s: number, at: number, n = 1) {
    this.guardPivot(s, 'r', at, n, false);
    this.structural(s, () => this.hf.addRows(s, [at, n]), (m) => {
      shiftAll(m, 'r', (x) => (x >= at ? x + n : x), (a, b) => ({ a: a >= at ? a + n : a, b: b >= at ? b + n : b }));
      setAll(m.styles, remapStyles(m.styles, (r, c) => ({ r: r >= at ? r + n : r, c })));
      m.rowH.shiftInsert(at, n);
    }, 'r', (a, b) => ({ a: a >= at ? a + n : a, b: b >= at ? b + n : b }));
  }
  deleteRows(s: number, at: number, n = 1) {
    this.guardPivot(s, 'r', at, n, true);
    this.structural(s, () => this.hf.removeRows(s, [at, n]), (m) => {
      shiftAll(m, 'r', (x) => (x < at ? x : x >= at + n ? x - n : -1), (a, b) => delSpan(a, b, at, n));
      setAll(m.styles, remapStyles(m.styles, (r, c) => (r < at ? { r, c } : r >= at + n ? { r: r - n, c } : null)));
      m.rowH.shiftDelete(at, n);
    }, 'r', (a, b) => delSpan(a, b, at, n));
  }
  insertCols(s: number, at: number, n = 1) {
    this.guardPivot(s, 'c', at, n, false);
    this.structural(s, () => this.hf.addColumns(s, [at, n]), (m) => {
      shiftAll(m, 'c', (x) => (x >= at ? x + n : x), (a, b) => ({ a: a >= at ? a + n : a, b: b >= at ? b + n : b }));
      setAll(m.styles, remapStyles(m.styles, (r, c) => ({ r, c: c >= at ? c + n : c })));
      m.colW.shiftInsert(at, n);
    }, 'c', (a, b) => ({ a: a >= at ? a + n : a, b: b >= at ? b + n : b }));
  }
  deleteCols(s: number, at: number, n = 1) {
    this.guardPivot(s, 'c', at, n, true);
    this.structural(s, () => this.hf.removeColumns(s, [at, n]), (m) => {
      shiftAll(m, 'c', (x) => (x < at ? x : x >= at + n ? x - n : -1), (a, b) => delSpan(a, b, at, n));
      setAll(m.styles, remapStyles(m.styles, (r, c) => (c < at ? { r, c } : c >= at + n ? { r, c: c - n } : null)));
      m.colW.shiftDelete(at, n);
    }, 'c', (a, b) => delSpan(a, b, at, n));
  }

  // ---------- sao chép / dán ----------
  copy(s: number, g: Range, cut = false): string {
    this.clip = { sheet: s, range: g, cut };
    return this.copyText(s, g);
  }
  copyText(s: number, g0: Range): string {
    const g = this.clampToUsed(s, g0);
    const rows: string[][] = [];
    for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 100_000); r++) {
      const row: string[] = [];
      for (let c = g.c1; c <= g.c2; c++) row.push(this.display(s, r, c).text);
      rows.push(row);
    }
    return toTSV(rows);
  }
  paste(s: number, r: number, c: number): Range | null {
    const cl = this.clip;
    if (!cl) return null;
    const src = this.clampToUsed(cl.sheet, cl.range);
    const h = src.r2 - src.r1, w = src.c2 - src.c1;
    if (r + h >= MAX_ROWS || c + w >= MAX_COLS) throw new Error('Vùng dán vượt quá giới hạn trang tính.');
    if (!this.sheetIds().includes(cl.sheet)) { this.clip = null; return null; }
    const srcSt = this.meta(cl.sheet).styles, dstSt = this.meta(s).styles;
    const ops: StyleOp[] = [];
    if (cl.cut) for (let i = 0; i <= h; i++) for (let j = 0; j <= w; j++) ops.push([srcSt, key(src.r1 + i, src.c1 + j), undefined]);
    for (let i = 0; i <= h; i++) for (let j = 0; j <= w; j++) ops.push([dstSt, key(r + i, c + j), srcSt.get(key(src.r1 + i, src.c1 + j))]);
    const tx = new StyleTxn(ops);
    const range = { sheet: cl.sheet, start: { sheet: cl.sheet, row: src.r1, col: src.c1 }, end: { sheet: cl.sheet, row: src.r2, col: src.c2 } };
    const target = this.A(s, r, c);
    // cắt cả bảng → bảng đi theo (như Excel)
    const moving = cl.cut ? this.meta(cl.sheet).tables.filter((t) => t.r1 >= src.r1 && t.r2 <= src.r2 && t.c1 >= src.c1 && t.c2 <= src.c2) : [];
    if (moving.length && this.meta(s).tables.some((t) => !moving.includes(t) && !(t.r2 < r || t.r1 > r + h || t.c2 < c || t.c1 > c + w))) throw new Error('Bảng không thể chồng lên một bảng khác.');
    this.group(() => {
      this.record(
        () => {
          if (cl.cut) this.hf.cut(range as any); else this.hf.copy(range as any);
          return this.hf.paste(target);
        },
        tx.hist,
        { batch: false, post: () => tx.apply() },
      );
      if (moving.length) {
        const dr = r - src.r1, dc = c - src.c1;
        const moved = moving.map((t) => ({ ...t, r1: t.r1 + dr, r2: t.r2 + dr, c1: t.c1 + dc, c2: t.c2 + dc }));
        const ids = new Set(moving.map((t) => t.id));
        if (s === cl.sheet) this.tableChange(s, (L) => L.map((x) => moved.find((y) => y.id === x.id) ?? x));
        else { this.tableChange(cl.sheet, (L) => L.filter((x) => !ids.has(x.id))); this.tableChange(s, (L) => [...L, ...moved]); }
      }
    });
    if (cl.cut) this.clip = null;
    return { r1: r, c1: c, r2: r + h, c2: c + w };
  }
  pasteText(s: number, r: number, c: number, text: string): Range | null {
    const rows = parseTSV(text);
    if (!rows.length) return null;
    const w = Math.max(...rows.map((x) => x.length));
    if (r + rows.length > MAX_ROWS || c + w > MAX_COLS) throw new Error('Vùng dán vượt quá giới hạn trang tính.');
    const fmts: [number, string][] = [];
    const data = rows.map((row, i) => Array.from({ length: w }, (_, j) => {
      const n = normalizeInput(row[j] ?? '');
      if (n.fmt) fmts.push([key(r + i, c + j), n.fmt]);
      return n.content as RawCellContent;
    }));
    const st = this.meta(s).styles;
    const tx = fmts.length ? new StyleTxn(fmts.map(([k, f]) => [st, k, { ...(st.get(k) ?? {}), fmt: f }])) : null;
    this.record(() => this.hf.setCellContents(this.A(s, r, c), data), tx?.hist, { post: () => tx?.apply() });
    return { r1: r, c1: c, r2: r + rows.length - 1, c2: c + w - 1 };
  }

  /** Gộp nhiều thao tác thành một bước hoàn tác (Ctrl+Z một lần) */
  group<T>(fn: () => T): T {
    const n0 = this.past.length;
    try { return fn(); }
    finally { if (this.past.length - n0 > 1) { const parts = this.past.splice(n0); this.past.push({ hf: false, parts }); } }
  }
  /** Ghi nhiều ô rời (chuỗi như người dùng gõ) trong một bước hoàn tác */
  writeCells(s: number, list: [number, number, string | number | boolean | null][]) {
    if (!list.length) return;
    const st = this.meta(s).styles;
    const fmts: StyleOp[] = [];
    const data = list.map(([r, c, v]) => {
      if (typeof v !== 'string') return [r, c, v] as const;
      const n = normalizeInput(v);
      const k = key(r, c);
      if (n.fmt && !st.get(k)?.fmt) fmts.push([st, k, { ...(st.get(k) ?? {}), fmt: n.fmt }]);
      return [r, c, n.content as RawCellContent] as const;
    });
    const tx = fmts.length ? new StyleTxn(fmts) : null;
    this.record(() => { for (const [r, c, v] of data) this.hf.setCellContents(this.A(s, r, c), [[v as RawCellContent]]); }, tx?.hist, { post: () => tx?.apply() });
  }
  /** Ô cuối cùng có dùng (dữ liệu hoặc định dạng) như Ctrl+End của Excel */
  lastCell(s: number): { r: number; c: number } {
    const u = this.used(s);
    let r = Math.max(0, u.rows - 1), c = Math.max(0, u.cols - 1);
    for (const k of this.meta(s).styles.keys()) { const p = unkey(+k); if (p.r > r) r = p.r; if (p.c > c) c = p.c; }
    for (const m of this.meta(s).merges) { r = Math.max(r, m.r2); c = Math.max(c, m.c2); }
    return { r, c };
  }
  /** Ô nằm trong vùng gộp nhưng không phải ô đầu */
  mergeCovered(s: number, r: number, c: number) { const m = this.mergeAt(s, r, c); return !!m && !(m.r1 === r && m.c1 === c); }
  sheetIdByName(n: string): number | undefined { return this.hf.getSheetId(n); }
  /** Công thức hợp lệ cú pháp không (không ghi vào ô) */
  isValidFormula(f: string, s: number): boolean {
    try { return this.hf.validateFormula(offsetIn(f, undefined, this.sheetName(s))); } catch { return false; }
  }
  /** Truy tìm Ô Tiền lệ: các ô/vùng công thức của ô dùng trực tiếp (cấp 1) */
  precedentsOf(s: number, r: number, c: number): { s: number; range: Range }[] {
    let p: any[] = [];
    try { p = this.hf.getCellPrecedents({ sheet: s, row: r, col: c }) as any[]; } catch { return []; }
    const out = p.filter((x) => (x.start ? x.start.sheet : x.sheet) >= 0).map((x) => x.start
      ? { s: x.start.sheet, range: { r1: x.start.row, c1: x.start.col, r2: x.end.row, c2: x.end.col } }
      : { s: x.sheet, range: { r1: x.row, c1: x.col, r2: x.row, c2: x.col } })
      .filter((x) => !isExtName(this.sheetName(x.s)));
    // vùng của bảng (tên ẩn XT_…) cũng là tiền lệ
    const raw = (this.hf as any).__rawFormula?.({ sheet: s, row: r, col: c }) as string | undefined;
    if (raw && /XT_/i.test(raw)) {
      const reg = this.tableReg();
      for (const m of raw.matchAll(/\bXT_(\d+)_(HD|DT|D|A|H|T)(?:_(\d+)_(\d+))?\b/gi)) {
        const t = reg.find((x) => x.id === +m[1]);
        const g = t && areaRange(t, m[2].toUpperCase() as Area, m[3] !== undefined ? +m[3] : undefined, m[4] !== undefined ? +m[4] : undefined);
        if (t && g) out.push({ s: t.sheet, range: g });
      }
    }
    return out;
  }
  /** Truy tìm Ô Phụ thuộc: các ô có công thức dùng trực tiếp ô này (cấp 1, kể cả qua vùng) */
  directDependents(s: number, r: number, c: number): { s: number; r: number; c: number }[] {
    const out = new Map<string, { s: number; r: number; c: number }>();
    const visit = (a: any, depth: number) => {
      let deps: any[] = [];
      try { deps = this.hf.getCellDependents(a) as any[]; } catch { deps = []; }
      for (const x of deps) {
        if (x.start) { if (depth < 2) visit(x, depth + 1); continue; } // vùng trung gian → ô dùng vùng
        if (x.sheet >= 0) out.set(x.sheet + ':' + x.row + ':' + x.col, { s: x.sheet, r: x.row, c: x.col });
        else if (depth < 2) visit(x, depth + 1); // qua tên (vd vùng bảng) → ô dùng tên
      }
    };
    visit({ sheet: s, row: r, col: c }, 0);
    return [...out.values()];
  }
  /** Ô phụ thuộc trực tiếp vào một ô (kể cả qua vùng) */
  dependentsOf(s: number, r: number, c: number): { s: number; r: number; c: number }[] {
    const out: { s: number; r: number; c: number }[] = [];
    const seen = new Set<string>();
    const q: any[] = [this.A(s, r, c)];
    while (q.length) {
      let deps: any[] = [];
      try { deps = this.hf.getCellDependents(q.pop()) as any[]; } catch { deps = []; }
      for (const x of deps) {
        if (x.start) { const k = JSON.stringify(x); if (!seen.has(k)) { seen.add(k); q.push(x); } continue; }
        const k = x.sheet + ':' + x.row + ':' + x.col;
        if (seen.has(k)) continue;
        seen.add(k); out.push({ s: x.sheet, r: x.row, c: x.col });
      }
    }
    return out;
  }

  /** Dán Đặc biệt (Ctrl+Alt+V) từ vùng đã Sao chép. Trả về vùng đã dán. Một bước hoàn tác. */
  pasteSpecial(s: number, target: Range, o: PasteOpts): Range {
    const cl = this.clip;
    if (!cl) throw new Error('Chưa có gì được sao chép.');
    if (cl.cut) throw new Error('Không dùng được Dán Đặc biệt với vùng đã Cắt — hãy Sao chép (Ctrl+C) rồi dán.');
    if (!this.sheetIds().includes(cl.sheet)) { this.clip = null; throw new Error('Trang tính nguồn không còn.'); }
    const src = this.clampToUsed(cl.sheet, cl.range);
    const sh = src.r2 - src.r1 + 1, sw = src.c2 - src.c1 + 1;
    const bh = o.transpose ? sw : sh, bw = o.transpose ? sh : sw;
    const tH = target.r2 - target.r1 + 1, tW = target.c2 - target.c1 + 1;
    const nr = tH > bh && tH % bh === 0 ? tH / bh : 1, nc = tW > bw && tW % bw === 0 ? tW / bw : 1;
    const out = { r1: target.r1, c1: target.c1, r2: target.r1 + bh * nr - 1, c2: target.c1 + bw * nc - 1 };
    if (out.r2 >= MAX_ROWS || out.c2 >= MAX_COLS) throw new Error('Vùng dán vượt quá giới hạn trang tính.');
    if (sh * sw * nr * nc > 500_000) throw new Error('Vùng dán quá lớn (tối đa 500.000 ô).');
    const what = o.link ? 'link' : o.what ?? 'all';
    const op = o.op && o.op !== 'none' && what !== 'link' ? o.op : null;
    const qn = this.quoteNames();
    // chụp nguồn trước khi ghi (đích có thể chồng nguồn)
    type Snap = { f?: string; k: unknown; v: unknown; st?: Style; blank: boolean };
    const snap: Snap[][] = [];
    const srcSt = this.meta(cl.sheet).styles;
    for (let i = 0; i < sh; i++) {
      const row: Snap[] = [];
      for (let j = 0; j < sw; j++) {
        const a = this.A(cl.sheet, src.r1 + i, src.c1 + j);
        const f0 = this.hf.getCellFormula(a);
        const f = f0 !== undefined ? this.outF(f0, qn) : undefined;
        const k = f === undefined ? this.hf.getCellSerialized(a) : null;
        row.push({ f, k, v: this.value(cl.sheet, src.r1 + i, src.c1 + j), st: srcSt.get(key(src.r1 + i, src.c1 + j)), blank: f === undefined && (k === null || k === '' || k === undefined) });
      }
      snap.push(row);
    }
    const dstSt = this.meta(s).styles;
    const cells: [number, number, RawCellContent][] = [];
    const sops: StyleOp[] = [];
    const sym = { add: '+', sub: '-', mul: '*', div: '/' } as const;
    const single = sh === 1 && sw === 1;
    const prefix = cl.sheet !== s ? quoteSheet(this.sheetName(cl.sheet)) + '!' : '';
    for (let a = 0; a < nr; a++) for (let b = 0; b < nc; b++) for (let i = 0; i < sh; i++) for (let j = 0; j < sw; j++) {
      const x = snap[i][j];
      if (o.skipBlanks && x.blank) continue;
      const sr = src.r1 + i, sc = src.c1 + j;
      const tr = out.r1 + a * bh + (o.transpose ? j : i), tc = out.c1 + b * bw + (o.transpose ? i : j);
      let content: RawCellContent | undefined;
      if (what === 'link') content = '=' + prefix + (single ? '$' + toA1(sr, sc).replace(/(\d)/, '$$$1') : toA1(sr, sc));
      else if (what === 'all' || what === 'formulas' || what === 'noBorders' || what === 'formulasNum' || what === 'allMerge' || what === 'theme')
        content = x.f !== undefined ? (o.transpose ? transposeFormula(x.f, sr, sc, tr, tc) : shiftFormula(x.f, tr - sr, tc - sc)) : x.blank ? null : (x.k as RawCellContent);
      else if (what === 'values' || what === 'valuesNum') content = x.blank ? null : valueContent(x.v);
      if (op && content !== undefined && x.f === undefined && !x.blank && what !== 'values' && what !== 'valuesNum') content = valueContent(x.v);
      if (op && content !== undefined) content = this.combine(s, tr, tc, content, x.blank, sym[op], qn);
      if (content !== undefined) cells.push([tr, tc, content]);
      // định dạng
      const kk = key(tr, tc);
      if (what === 'all' || what === 'theme' || what === 'allMerge') sops.push([dstSt, kk, x.st ? { ...x.st } : undefined]);
      else if (what === 'formats') {
        // chỉ định dạng: giữ ghi chú/chú thích của ô đích
        const d = dstSt.get(kk);
        sops.push([dstSt, kk, clean({ ...(x.st ?? {}), nt: d?.nt, cm: d?.cm, cmDone: d?.cmDone })]);
      } else if (what === 'comments') {
        const d = dstSt.get(kk);
        sops.push([dstSt, kk, clean({ ...(d ?? {}), nt: x.st?.nt, cm: x.st?.cm, cmDone: x.st?.cmDone })]);
      }
      else if (what === 'noBorders') { const n = x.st ? { ...x.st } : undefined; if (n) delete n.bd; sops.push([dstSt, kk, n && Object.keys(n).length ? n : undefined]); }
      else if (what === 'formulasNum' || what === 'valuesNum') { const cur = { ...(dstSt.get(kk) ?? {}) }; if (x.st?.fmt) cur.fmt = x.st.fmt; else delete cur.fmt; sops.push([dstSt, kk, Object.keys(cur).length ? cur : undefined]); }
    }
    this.group(() => {
      if (cells.length || sops.length) {
        const tx = sops.length ? new StyleTxn(sops) : null;
        this.record(cells.length ? () => { for (const [r, c, v] of cells) this.hf.setCellContents(this.A(s, r, c), [[v]]); } : null, tx?.hist ?? { undo: () => {}, redo: () => {} }, { post: () => tx?.apply() });
      }
      if (what === 'colWidths') {
        const src0 = this.meta(cl.sheet).colW, list: [number, number][] = [];
        for (let b = 0; b < nc; b++) for (let j = 0; j < (o.transpose ? sh : sw); j++) if (!o.transpose) list.push([out.c1 + b * bw + j, src0.get(src.c1 + j)]);
        this.setSizes(s, 'c', list);
      }
      if (what === 'validation' || what === 'all' || what === 'allMerge' || what === 'theme' || what === 'noBorders') {
        const add: DVRule[] = [];
        for (const d of this.meta(cl.sheet).dv) {
          const g = { r1: Math.max(d.range.r1, src.r1), c1: Math.max(d.range.c1, src.c1), r2: Math.min(d.range.r2, src.r2), c2: Math.min(d.range.c2, src.c2) };
          if (g.r1 > g.r2 || g.c1 > g.c2 || o.transpose) continue;
          for (let a = 0; a < nr; a++) for (let b = 0; b < nc; b++) {
            const dr = out.r1 + a * bh - src.r1, dc = out.c1 + b * bw - src.c1;
            add.push({ ...d, id: undefined, range: { r1: g.r1 + dr, c1: g.c1 + dc, r2: g.r2 + dr, c2: g.c2 + dc } });
          }
        }
        for (const d of add) this.addDV(s, d);
      }
    });
    return out;
  }
  /** Dán Đặc biệt có phép tính: đích (giá trị/công thức) ⊕ nguồn */
  private combine(s: number, r: number, c: number, content: RawCellContent, srcBlank: boolean, sym: string, qn: string[]): RawCellContent {
    const a = this.A(s, r, c);
    const tf0 = this.hf.getCellFormula(a);
    const tf = tf0 !== undefined ? this.outF(tf0, qn) : undefined;
    const tv = this.value(s, r, c);
    const srcIsF = typeof content === 'string' && content.startsWith('=');
    if (!srcIsF && typeof content === 'string' && content !== '' && !srcBlank) return content; // chữ: dán như thường
    if (srcBlank && tf === undefined && (tv === null || tv === '')) return null;               // trống ⊕ trống = trống
    if (tf === undefined && typeof tv === 'string' && tv !== '') return tv;                    // đích là chữ: giữ
    const sExpr = srcIsF ? '(' + (content as string).slice(1) + ')' : srcBlank || content === null ? '0' : lit(content);
    if (tf !== undefined || srcIsF) return '=(' + (tf !== undefined ? tf.slice(1) : lit(tv ?? 0)) + ')' + sym + sExpr;
    const x = typeof tv === 'number' ? tv : 0, y = srcBlank || content === null ? 0 : typeof content === 'number' ? content : Number(content) || 0;
    if (sym === '/' && y === 0) return '=' + lit(x) + '/0';
    const v = sym === '+' ? x + y : sym === '-' ? x - y : sym === '*' ? x * y : x / y;
    return Math.round(v * 1e12) / 1e12;
  }

  /** Chổi định dạng: dán kiểu ô (lặp khối nguồn cho kín vùng đích), thay toàn bộ định dạng cũ */
  paintStyles(s: number, g0: Range, src: (Style | undefined)[][]) {
    const g = this.clampToUsed(s, g0);
    if (rangeSize(g) > 500_000) throw new Error('Vùng quá lớn (tối đa 500.000 ô).');
    const st = this.meta(s).styles, h = src.length, w = src[0]?.length ?? 1;
    const ops: StyleOp[] = [];
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const k = key(r, c);
      const n = src[(r - g.r1) % h]?.[(c - g.c1) % w];
      // giữ liên kết, ghi chú, chú thích của ô đích (chổi định dạng chỉ chép định dạng)
      const o = st.get(k);
      const keep = { ln: o?.ln, nt: o?.nt, cm: o?.cm, cmDone: o?.cmDone };
      const base = n ? { ...n, ln: undefined, nt: undefined, cm: undefined, cmDone: undefined } : {};
      const v = clean({ ...base, ...keep });
      ops.push([st, k, v]);
    }
    const tx = new StyleTxn(ops);
    tx.apply();
    this.record(null, tx.hist);
  }
  /** Tính một công thức tạm (không ghi vào ô) — Đối số Hàm, F9 khi sửa */
  evalFormula(f: string, s: number): unknown {
    try {
      const v: any = this.hf.calculateFormula(f, s);
      if (Array.isArray(v)) return v[0]?.[0];
      return v && typeof v === 'object' && 'value' in v ? v.value : v;
    } catch { return undefined; }
  }

  /** Trang đầu ▸ Điền ▸ Chuỗi…: dir 'rows' = điền theo hàng (sang phải), 'cols' = theo cột (xuống) */
  fillSeries(s: number, g: Range, o: SeriesOpts & { dir: 'rows' | 'cols'; trend?: boolean }) {
    if (o.type === 'auto') { const src = o.dir === 'cols' ? { ...g, r2: g.r1 } : { ...g, c2: g.c1 }; this.fill(s, src, g); return; }
    const list: [number, number, string | number | boolean | null][] = [];
    const fmts: StyleOp[] = [];
    const lines = o.dir === 'cols' ? [g.c1, g.c2] : [g.r1, g.r2];
    for (let L = lines[0]; L <= lines[1]; L++) {
      const at = (i: number) => (o.dir === 'cols' ? { r: g.r1 + i, c: L } : { r: L, c: g.c1 + i });
      const n = o.dir === 'cols' ? g.r2 - g.r1 + 1 : g.c2 - g.c1 + 1;
      const v0 = this.value(s, at(0).r, at(0).c);
      if (typeof v0 !== 'number') continue;
      let step = o.step ?? 1;
      if (o.trend && o.type !== 'date') { const v1 = this.value(s, at(1).r, at(1).c); if (typeof v1 === 'number') step = o.type === 'growth' ? (v0 ? v1 / v0 : 1) : v1 - v0; }
      const vals = seriesValues(v0, n, { ...o, step });
      const f0 = this.effFmt(s, at(0).r, at(0).c);
      vals.forEach((v, i) => { if (i > 0) { const p = at(i); list.push([p.r, p.c, v]); if (f0) fmts.push([this.meta(s).styles, key(p.r, p.c), { ...(this.meta(s).styles.get(key(p.r, p.c)) ?? {}), fmt: f0 }]); } });
    }
    if (!list.length) throw new Error('Ô đầu tiên của vùng chọn cần là một số hoặc ngày.');
    const tx = fmts.length ? new StyleTxn(fmts) : null;
    this.record(() => { for (const [r, c, v] of list) this.hf.setCellContents(this.A(s, r, c), [[v as RawCellContent]]); }, tx?.hist, { post: () => tx?.apply() });
  }

  // ---------- kéo điền ----------
  fill(s: number, src: Range, tgt: Range) {
    const down = tgt.r2 > src.r2, up = tgt.r1 < src.r1, right = tgt.c2 > src.c2, left = tgt.c1 < src.c1;
    if (!down && !up && !right && !left) return;
    const vertical = down || up;
    const fwd = down || right;
    const count = vertical ? (down ? tgt.r2 - src.r2 : src.r1 - tgt.r1) : (right ? tgt.c2 - src.c2 : src.c1 - tgt.c1);
    if (count * (vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1) > 200_000) throw new Error('Vùng kéo điền quá lớn (tối đa 200.000 ô).');
    const m = this.meta(s);
    const sops: StyleOp[] = [];
    const qn = this.quoteNames();
    const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
    const ops: (() => void)[] = [];
    for (let line = lines[0]; line <= lines[1]; line++) {
      const cells: { r: number; c: number }[] = [];
      if (vertical) for (let r = src.r1; r <= src.r2; r++) cells.push({ r, c: line });
      else for (let c = src.c1; c <= src.c2; c++) cells.push({ r: line, c });
      const snaps: SrcCell[] = cells.map(({ r, c }) => {
        const v = this.value(s, r, c);
        const f = this.hf.getCellFormula(this.A(s, r, c));
        return { raw: f !== undefined ? this.outF(f, qn) : this.hf.getCellSerialized(this.A(s, r, c)), formula: f !== undefined, num: typeof v === 'number' ? v : undefined, isDate: isDateFmt(this.effFmt(s, r, c)) };
      });
      const plan = fillPlan(snaps, count, fwd);
      plan.forEach((it, k0) => {
        const k = k0 + 1;
        const t = vertical ? { r: fwd ? src.r2 + k : src.r1 - k, c: line } : { r: line, c: fwd ? src.c2 + k : src.c1 - k };
        const si = 'copy' in it ? it.copy : fwd ? (k - 1) % cells.length : ((cells.length - (k % cells.length)) % cells.length);
        const from = cells[si];
        sops.push([m.styles, key(t.r, t.c), m.styles.get(key(from.r, from.c))]);
        if ('copy' in it) {
          const sn = snaps[si];
          const content = sn.formula ? shiftFormula(String(sn.raw), t.r - from.r, t.c - from.c) : sn.raw;
          ops.push(() => this.hf.setCellContents(this.A(s, t.r, t.c), [[content as RawCellContent]]));
        } else ops.push(() => this.hf.setCellContents(this.A(s, t.r, t.c), [[it.value as RawCellContent]]));
      });
    }
    const tx = new StyleTxn(sops);
    this.record(() => ops.forEach((f) => f()), tx.hist, { post: () => tx.apply() });
  }

  /** Tổng Tự động (hoặc Trung bình, Đếm Số, Tối đa, Tối thiểu): tìm vùng số liền kề phía trên/bên trái */
  autoSum(s: number, r: number, c: number, fn = 'SUM', write = true): string {
    if (fn !== 'SUM') { const f0 = this.autoSum(s, r, c, 'SUM', false); const f1 = f0.replace('=SUM(', `=${fn}(`); if (write && f1 !== `=${fn}()`) this.setInput(s, r, c, f1); return f1; }
    const isNum = (rr: number, cc: number) => rr >= 0 && cc >= 0 && typeof this.value(s, rr, cc) === 'number';
    let f = '=SUM()';
    if (isNum(r - 1, c)) {
      let t = r - 1;
      while (isNum(t - 1, c)) t--;
      f = `=SUM(${colLetters(c)}${t + 1}:${colLetters(c)}${r})`;
    } else if (isNum(r, c - 1)) {
      let t = c - 1;
      while (isNum(r, t - 1)) t--;
      f = `=SUM(${colLetters(t)}${r + 1}:${colLetters(c - 1)}${r + 1})`;
    }
    if (f !== '=SUM()' && write) this.setInput(s, r, c, f);
    return f;
  }

  // ---------- gộp ô ----------
  private mIdx = new WeakMap<Range[], Map<number, Range>>();
  mergeAt(s: number, r: number, c: number): Range | null {
    const list = this.meta(s).merges;
    if (!list.length) return null;
    let idx = this.mIdx.get(list);
    if (!idx) {
      idx = new Map();
      for (const g of list) for (let rr = g.r1; rr <= g.r2; rr++) for (let cc = g.c1; cc <= g.c2; cc++) idx.set(key(rr, cc), g);
      this.mIdx.set(list, idx);
    }
    return idx.get(key(r, c)) ?? null;
  }
  /** Mở rộng vùng chọn để bao trọn các ô gộp cắt ngang nó */
  expandToMerges(s: number, g: Range): Range {
    const list = this.meta(s).merges;
    if (!list.length) return g;
    let out = { ...g }, changed = true;
    while (changed) {
      changed = false;
      for (const m of list) {
        if (m.r2 < out.r1 || m.r1 > out.r2 || m.c2 < out.c1 || m.c1 > out.c2) continue;
        const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
        if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) { out = n; changed = true; }
      }
    }
    return out;
  }
  merge(s: number, g0: Range, center = true) {
    const g = this.expandToMerges(s, g0);
    if (g.r1 === g.r2 && g.c1 === g.c2) return;
    if (rangeSize(g) > 1_000_000) throw new Error('Vùng gộp quá lớn.');
    const m = this.meta(s);
    const before = m.merges, after = [...m.merges.filter((x) => x.r2 < g.r1 || x.r1 > g.r2 || x.c2 < g.c1 || x.c1 > g.c2), g];
    const u = this.used(s);
    const clears: [number, number][] = [];
    for (let r = g.r1; r <= Math.min(g.r2, u.rows - 1); r++) for (let c = g.c1; c <= Math.min(g.c2, u.cols - 1); c++)
      if ((r !== g.r1 || c !== g.c1) && this.hf.getCellSerialized(this.A(s, r, c)) !== null) clears.push([r, c]);
    const st = m.styles, k0 = key(g.r1, g.c1);
    const tx = center ? new StyleTxn([[st, k0, { ...(st.get(k0) ?? {}), al: 'c' }]]) : null;
    this.record(clears.length ? () => { for (const [r, c] of clears) this.hf.setCellContents(this.A(s, r, c), [[null]]); } : null, {
      undo: () => { this.meta(s).merges = before; tx?.hist.undo(); },
      redo: () => { this.meta(s).merges = after; tx?.hist.redo(); },
    }, { post: () => { m.merges = after; tx?.apply(); } });
    return clears.length;
  }
  unmerge(s: number, g: Range) {
    const m = this.meta(s);
    const before = m.merges, after = m.merges.filter((x) => x.r2 < g.r1 || x.r1 > g.r2 || x.c2 < g.c1 || x.c1 > g.c2);
    if (after.length === before.length) return;
    m.merges = after;
    this.record(null, { undo: () => (this.meta(s).merges = before), redo: () => (this.meta(s).merges = after) });
  }

  // ---------- viền ----------
  setBorder(s: number, g0: Range, kind: BorderKind, style: BorderStyle = 'thin', color?: string) {
    const g = this.clampToUsed(s, g0);
    if (rangeSize(g) > 500_000) throw new Error('Vùng chọn quá lớn để kẻ viền (tối đa 500.000 ô).');
    const st = this.meta(s).styles;
    const b: Border = color ? { s: style, c: color } : { s: style };
    const ops: StyleOp[] = [];
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const k = key(r, c);
      const o = st.get(k) ?? {};
      let bd: Style['bd'] = { ...(o.bd ?? {}) };
      if (kind === 'none') bd = undefined;
      else if (kind === 'all') bd = { t: b, r: b, b, l: b };
      else {
        if ((kind === 'outer' || kind === 'top') && r === g.r1) bd.t = b;
        if ((kind === 'outer' || kind === 'bottom') && r === g.r2) bd.b = b;
        if ((kind === 'outer' || kind === 'left') && c === g.c1) bd.l = b;
        if ((kind === 'outer' || kind === 'right') && c === g.c2) bd.r = b;
        if (kind === 'inner') { if (r > g.r1) bd.t = b; if (r < g.r2) bd.b = b; if (c > g.c1) bd.l = b; if (c < g.c2) bd.r = b; }
      }
      ops.push([st, k, clean({ ...o, bd: bd && Object.keys(bd).length ? bd : undefined })]);
    }
    const tx = new StyleTxn(ops);
    tx.apply();
    this.record(null, tx.hist);
  }

  /**
   * Kẻ viền theo hộp Định dạng Ô ▸ Viền: t/b/l/r = mép ngoài vùng, h/v = đường trong ngang/dọc, d/u = chéo trong từng ô.
   * undefined = giữ nguyên, null = bỏ. Ô kề bên ngoài vùng cũng được bỏ cạnh chung khi bỏ mép (như Excel).
   */
  setBorders(s: number, g0: Range, spec: Partial<Record<'t' | 'b' | 'l' | 'r' | 'h' | 'v' | 'd' | 'u', Border | null>>) {
    const g = this.clampToUsed(s, g0);
    if (rangeSize(g) > 500_000) throw new Error('Vùng chọn quá lớn để kẻ viền (tối đa 500.000 ô).');
    const st = this.meta(s).styles;
    const ops: StyleOp[] = [];
    const put = (bd: NonNullable<Style['bd']>, k: keyof NonNullable<Style['bd']>, v: Border | null | undefined) => { if (v === undefined) return; if (v === null) delete bd[k]; else bd[k] = { ...v }; };
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const k = key(r, c);
      const o = st.get(k) ?? {};
      const bd: NonNullable<Style['bd']> = { ...(o.bd ?? {}) };
      put(bd, 't', r === g.r1 ? spec.t : spec.h);
      put(bd, 'b', r === g.r2 ? spec.b : spec.h);
      put(bd, 'l', c === g.c1 ? spec.l : spec.v);
      put(bd, 'r', c === g.c2 ? spec.r : spec.v);
      put(bd, 'd', spec.d); put(bd, 'u', spec.u);
      ops.push([st, k, clean({ ...o, bd: Object.keys(bd).length ? bd : undefined })]);
    }
    // bỏ mép ngoài: bỏ luôn cạnh chung của ô kề bên (lưới vẽ cả viền của ô kề)
    const nb = (r: number, c: number, side: 't' | 'b' | 'l' | 'r') => {
      if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return;
      const k = key(r, c), o = st.get(k);
      if (!o?.bd?.[side]) return;
      const bd = { ...o.bd }; delete bd[side];
      ops.push([st, k, clean({ ...o, bd: Object.keys(bd).length ? bd : undefined })]);
    };
    if (spec.t === null) for (let c = g.c1; c <= g.c2; c++) nb(g.r1 - 1, c, 'b');
    if (spec.b === null) for (let c = g.c1; c <= g.c2; c++) nb(g.r2 + 1, c, 't');
    if (spec.l === null) for (let r = g.r1; r <= g.r2; r++) nb(r, g.c1 - 1, 'r');
    if (spec.r === null) for (let r = g.r1; r <= g.r2; r++) nb(r, g.c2 + 1, 'l');
    const tx = new StyleTxn(ops);
    tx.apply();
    this.record(null, tx.hist);
  }

  // ---------- vùng dữ liệu, sắp xếp ----------
  private has(s: number, r: number, c: number) { const v = this.hf.getCellSerialized(this.A(s, r, c)); return v !== null && v !== ''; }
  /** Vùng dữ liệu liền kề quanh ô (như Ctrl+A / CurrentRegion của Excel) */
  currentRegion(s: number, r: number, c: number): Range {
    const u = this.used(s);
    const g = { r1: r, c1: c, r2: r, c2: c };
    const any = (r1: number, r2: number, c1: number, c2: number) => {
      for (let rr = Math.max(0, r1); rr <= Math.min(r2, u.rows - 1); rr++) for (let cc = Math.max(0, c1); cc <= Math.min(c2, u.cols - 1); cc++) if (this.has(s, rr, cc)) return true;
      return false;
    };
    for (let guard = 0; guard < 100000; guard++) {
      let ch = false;
      if (g.r1 > 0 && any(g.r1 - 1, g.r1 - 1, g.c1 - 1, g.c2 + 1)) { g.r1--; ch = true; }
      if (g.r2 < u.rows - 1 && any(g.r2 + 1, g.r2 + 1, g.c1 - 1, g.c2 + 1)) { g.r2++; ch = true; }
      if (g.c1 > 0 && any(g.r1 - 1, g.r2 + 1, g.c1 - 1, g.c1 - 1)) { g.c1--; ch = true; }
      if (g.c2 < u.cols - 1 && any(g.r1 - 1, g.r2 + 1, g.c2 + 1, g.c2 + 1)) { g.c2++; ch = true; }
      if (!ch) break;
    }
    return g;
  }
  guessHeader(s: number, g: Range): boolean {
    if (g.r1 === g.r2) return false;
    let allText = true, numBelow = false;
    for (let c = g.c1; c <= g.c2; c++) {
      const v = this.value(s, g.r1, c);
      if (v !== null && typeof v !== 'string') allText = false;
      if (typeof this.value(s, g.r1 + 1, c) === 'number') numBelow = true;
    }
    return allText && (numBelow || !!this.style(s, g.r1, g.c1).b);
  }
  sort(s: number, g0: Range, keys: SortKey[], header: boolean) {
    const u = this.used(s);
    const g = { ...g0, r2: Math.min(g0.r2, u.rows - 1), c2: Math.min(g0.c2, u.cols - 1) };
    const d1 = header ? g.r1 + 1 : g.r1;
    if (d1 >= g.r2 || !keys.length) return;
    if (keys.some((k) => k.col < g.c1 || k.col > g.c2)) throw new Error('Cột sắp xếp nằm ngoài vùng dữ liệu.');
    if (this.meta(s).merges.some((m) => !(m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2)))
      throw new Error('Không sắp xếp được vùng có ô gộp. Hãy bỏ gộp ô trước.');
    if ((g.r2 - d1 + 1) * (g.c2 - g.c1 + 1) > 2_000_000) throw new Error('Vùng sắp xếp quá lớn.');
    const st = this.meta(s).styles;
    const qn = this.quoteNames();
    // khi đang lọc: chỉ sắp xếp các hàng đang hiện (như Excel), hàng ẩn giữ nguyên chỗ
    const rowH = this.meta(s).rowH;
    const slots: number[] = [];
    for (let r = d1; r <= g.r2; r++) if (!rowH.isHidden(r)) slots.push(r);
    type Row = { r: number; cells: { raw: RawCellContent; f: boolean }[]; styles: (Style | undefined)[]; keyv: unknown[] };
    const rows: Row[] = [];
    for (const r of slots) {
      const cells = [], styles = [];
      for (let c = g.c1; c <= g.c2; c++) {
        const a = this.A(s, r, c), f = this.hf.getCellFormula(a);
        cells.push(f !== undefined ? { raw: this.outF(f, qn), f: true } : { raw: this.hf.getCellSerialized(a), f: false });
        styles.push(st.get(key(r, c)));
      }
      rows.push({ r, cells, styles, keyv: keys.map((k) => this.value(s, r, k.col)) });
    }
    const rank = (v: unknown) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
    const blank = (v: unknown) => v === null || v === '';
    rows.sort((x, y) => {
      for (let i = 0; i < keys.length; i++) {
        const a = x.keyv[i], b = y.keyv[i];
        if (blank(a) || blank(b)) { if (blank(a) && blank(b)) continue; return blank(a) ? 1 : -1; } // ô trống luôn cuối
        let d = rank(a) - rank(b);
        if (!d) d = typeof a === 'number' ? (a as number) - (b as number) : typeof a === 'string' ? COLLATOR.compare(a, b as string) : typeof a === 'boolean' ? Number(a) - Number(b) : 0;
        if (d) return keys[i].desc ? -d : d;
      }
      return x.r - y.r;
    });
    const data = rows.map((row, j) => row.cells.map((cell) => (cell.f ? shiftFormula(String(cell.raw), slots[j] - row.r, 0) : cell.raw)));
    const ops: StyleOp[] = [];
    rows.forEach((row, j) => row.styles.forEach((v, i) => ops.push([st, key(slots[j], g.c1 + i), v])));
    const tx = new StyleTxn(ops);
    const contiguous = slots.length === g.r2 - d1 + 1;
    this.record(() => {
      if (contiguous) this.hf.setCellContents(this.A(s, d1, g.c1), data);
      else data.forEach((row, j) => this.hf.setCellContents(this.A(s, slots[j], g.c1), [row]));
    }, tx.hist, { post: () => tx.apply() });
  }

  /**
   * Loại bỏ Trùng lặp: giữ lần xuất hiện đầu, so theo các cột `cols` (chữ không phân biệt hoa/thường như Excel).
   * Chỉ dồn các ô trong vùng lên (không xóa nguyên hàng), ô thừa ở đáy vùng bị xóa. Trả về số hàng bị loại / còn lại.
   */
  removeDuplicates(s: number, g0: Range, cols: number[], header: boolean): { removed: number; kept: number } {
    const u = this.used(s);
    const g = { ...g0, r2: Math.min(g0.r2, u.rows - 1), c2: Math.min(g0.c2, u.cols - 1) };
    const d1 = header ? g.r1 + 1 : g.r1;
    if (!cols.length) throw new Error('Hãy chọn ít nhất một cột.');
    if (d1 > g.r2) return { removed: 0, kept: 0 };
    if (this.meta(s).merges.some((m) => !(m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2)))
      throw new Error('Không loại bỏ trùng lặp được trong vùng có ô gộp. Hãy bỏ gộp ô trước.');
    const norm = (v: unknown) => (typeof v === 'string' ? 's' + v.toLocaleLowerCase('vi') : typeof v === 'number' ? 'n' + v : v === null || v === undefined ? 'e' : typeof v + String(v));
    const seen = new Set<string>(), keep: number[] = [];
    for (let r = d1; r <= g.r2; r++) {
      const k = cols.map((c) => norm(this.value(s, r, c))).join('\u0001');
      if (seen.has(k)) continue;
      seen.add(k); keep.push(r);
    }
    const removed = g.r2 - d1 + 1 - keep.length;
    if (!removed) return { removed: 0, kept: keep.length };
    const st = this.meta(s).styles, qn = this.quoteNames();
    const data: RawCellContent[][] = [], ops: StyleOp[] = [];
    keep.forEach((r, j) => {
      const row: RawCellContent[] = [];
      for (let c = g.c1; c <= g.c2; c++) {
        const a = this.A(s, r, c), f = this.hf.getCellFormula(a);
        row.push(f !== undefined ? shiftFormula(String(this.outF(f, qn)), d1 + j - r, 0) : this.hf.getCellSerialized(a));
        ops.push([st, key(d1 + j, c), st.get(key(r, c))]);
      }
      data.push(row);
    });
    for (let r = d1 + keep.length; r <= g.r2; r++) {
      data.push(new Array(g.c2 - g.c1 + 1).fill(null));
      for (let c = g.c1; c <= g.c2; c++) ops.push([st, key(r, c), undefined]);
    }
    const tx = new StyleTxn(ops);
    const tbl = this.meta(s).tables.find((t) => t.c1 === g.c1 && t.c2 === g.c2 && dataRows(t).a === d1 && dataRows(t).b === g.r2);
    this.group(() => {
      this.record(() => { this.hf.setCellContents(this.A(s, d1, g.c1), data); }, tx.hist, { post: () => tx.apply() });
      // trong bảng: bảng co lại (hàng tổng dời lên) như Excel
      if (tbl) { this.moveTotals(s, tbl, tbl.r2 - removed); this.tableChange(s, (L) => L.map((x) => (x.id === tbl.id ? { ...x, r2: x.r2 - removed } : x))); }
    });
    return { removed, kept: keep.length };
  }
  /**
   * Văn bản thành Cột: tách cột `col` (hàng r1..r2) theo dấu phân cách hoặc độ rộng cố định, ghi từ ô đích.
   * types: 'g' Chung (tự nhận số/ngày), 't' Văn bản, 'x' Bỏ qua cột. Trả về số cột kết quả.
   */
  textToColumns(s: number, col: number, r1: number, r2: number, o: { delims?: string; consecutive?: boolean; qualifier?: string; widths?: number[]; dest?: { r: number; c: number }; types?: string[] }): number {
    const parts = (t: string): string[] => {
      if (o.widths) {
        const out: string[] = []; let p = 0;
        for (const w of [...o.widths, Infinity]) { if (p >= t.length) break; out.push(t.slice(p, w === Infinity ? undefined : w).trimEnd()); if (w === Infinity) break; p = w; }
        return out;
      }
      const d = o.delims ?? '\t', q = o.qualifier ?? '"';
      const out: string[] = []; let cur = '', inQ = false, i = 0;
      while (i < t.length) {
        const ch = t[i];
        if (q && ch === q) { if (inQ && t[i + 1] === q) { cur += q; i += 2; continue; } inQ = !inQ; i++; continue; }
        if (!inQ && d.includes(ch)) {
          out.push(cur); cur = '';
          if (o.consecutive) while (i + 1 < t.length && d.includes(t[i + 1])) i++;
          i++; continue;
        }
        cur += ch; i++;
      }
      out.push(cur);
      return out;
    };
    const dest = o.dest ?? { r: r1, c: col };
    const list: [number, number, string | null][] = [];
    let width = 0;
    for (let r = r1; r <= r2; r++) {
      const t = this.display(s, r, col).text;
      if (!t && this.value(s, r, col) === null) continue;
      const p = parts(t);
      let k = 0;
      p.forEach((v, i) => {
        const ty = o.types?.[i] ?? 'g';
        if (ty === 'x') return;
        list.push([dest.r + r - r1, dest.c + k, v === '' ? null : ty === 't' ? "'" + v : v]);
        k++;
      });
      width = Math.max(width, k);
    }
    this.writeCells(s, list);
    return width;
  }

  // ---------- Bảng (Chèn ▸ Bảng, Ctrl+T) ----------
  tableSeq = 0;
  /** mọi bảng của sổ kèm tên sheet (cho lớp dịch tham chiếu có cấu trúc) */
  tableReg(): TableReg[] {
    const out: TableReg[] = [];
    for (const [id, m] of this.metas) if (m.tables.length) { const n = this.hf.getSheetName(id); if (n === undefined) continue; for (const t of m.tables) out.push({ ...t, sheet: id, sheetName: n }); }
    return out;
  }
  tables(s: number): TableDef[] { return this.meta(s).tables; }
  tableAt(s: number, r: number, c: number): TableDef | undefined { const l = this.meta(s).tables; return l.length ? l.find((t) => inTable(t, r, c)) : undefined; }
  tableByName(n: string): TableReg | undefined { return this.tableReg().find((t) => t.name.toLowerCase() === n.toLowerCase()); }
  /** định dạng hiển thị: định dạng của bảng nằm dưới định dạng riêng của ô */
  viewStyle(s: number, r: number, c: number): Style {
    const st = this.style(s, r, c);
    const l = this.meta(s).tables;
    if (!l.length) return st;
    const t = l.find((x) => inTable(x, r, c));
    const ts = t && tableCellStyle(t, r, c);
    if (!ts) return st;
    const o: Style = { ...ts, ...st };
    if (ts.bd || st.bd) o.bd = { ...(ts.bd ?? {}), ...(st.bd ?? {}) };
    return o;
  }
  private tableNameTaken(n: string, except?: number) {
    const u = n.toLowerCase();
    return this.tableReg().some((t) => t.id !== except && t.name.toLowerCase() === u) || this.listNames().some((x) => x.name.toLowerCase() === u);
  }
  /** thay danh sách bảng của sheet (+ việc trên bộ tính) thành một bước hoàn tác */
  private tableChange(s: number, f: (list: TableDef[]) => TableDef[], hfOp?: () => void) {
    const m = this.meta(s), before = m.tables;
    let after = before;
    this.record(() => { m.tables = f(m.tables); after = m.tables; this.syncXT(); hfOp?.(); }, { undo: () => { m.tables = before; }, redo: () => { m.tables = after; } });
  }
  /** cập nhật biểu thức các tên ẩn XT_… theo vùng bảng hiện tại (gọi trong batch) */
  private syncXT() {
    const reg = this.tableReg();
    for (const n of this.hf.listNamedExpressions()) {
      const x = /^XT_(\d+)_(HD|DT|D|A|H|T)(?:_(\d+)_(\d+))?$/i.exec(n);
      if (!x) continue;
      const t = reg.find((y) => y.id === +x[1]);
      const e = t ? xtExpr(t, x[2].toUpperCase() as any, x[3] !== undefined ? +x[3] : undefined, x[4] !== undefined ? +x[4] : undefined) : '=#REF!';
      const cur = this.hf.getNamedExpression(n)?.expression;
      if (cur !== e) try { this.hf.changeNamedExpression(n, e); } catch { /* bỏ qua */ }
    }
  }
  /** gợi ý tên bảng kế tiếp: Bảng1, Bảng2… */
  nextTableName() { let i = 1; while (this.tableNameTaken('Bảng' + i)) i++; return 'Bảng' + i; }
  /**
   * Tạo bảng từ vùng g. header=false: chèn hàng tiêu đề Cột1… phía trên (như Excel “Bảng có tiêu đề” bỏ chọn).
   * Trả về bảng mới.
   */
  createTable(s: number, g0: Range, header = true, style = 'TableStyleMedium2'): TableDef {
    const u = this.used(s);
    let g = { ...g0, r2: g0.r2 >= MAX_ROWS - 1 ? Math.max(g0.r1, u.rows - 1) : g0.r2, c2: g0.c2 >= MAX_COLS - 1 ? Math.max(g0.c1, u.cols - 1) : g0.c2 };
    if (g.c2 - g.c1 > 500) throw new Error('Bảng quá rộng (tối đa 500 cột).');
    if (this.meta(s).tables.some((t) => !(t.r2 < g.r1 || t.r1 > g.r2 + (header ? 0 : 1) || t.c2 < g.c1 || t.c1 > g.c2))) throw new Error('Bảng không thể chồng lên một bảng khác.');
    if (this.meta(s).merges.some((m) => !(m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2))) throw new Error('Vùng có ô đã gộp. Hãy bỏ gộp ô trước khi tạo bảng.');
    if (this.meta(s).pivots.some((p) => p.area && !(p.area.r2 < g.r1 || p.area.r1 > g.r2 || p.area.c2 < g.c1 || p.area.c1 > g.c2))) throw new Error('Bảng không thể chồng lên PivotTable.');
    let t!: TableDef;
    this.group(() => {
      if (!header) { this.insertRows(s, g.r1, 1); g = { ...g, r2: g.r2 + 1 }; }
      if (g.r2 === g.r1) g = { ...g, r2: g.r1 + 1 };
      const used = new Set<string>();
      const cols: TableDef['cols'] = [];
      for (let c = g.c1; c <= g.c2; c++) {
        let nm = header ? this.display(s, g.r1, c).text.trim() : '';
        if (!nm) { let j = c - g.c1 + 1; nm = `Cột${j}`; while (used.has(nm.toLowerCase())) nm = `Cột${++j}`; }
        let base = nm, k = 2; while (used.has(nm.toLowerCase())) nm = base + k++;
        used.add(nm.toLowerCase()); cols.push({ id: cols.length + 1, name: nm });
      }
      const id = ++this.tableSeq;
      t = { id, name: this.nextTableName(), r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, header: true, totals: false, style, bandR: true, bandC: false, first: false, last: false, btn: true, cols };
      // tiêu đề luôn là chữ cố định
      const head = cols.map((x, i) => [g.r1, g.c1 + i, "'" + x.name] as [number, number, string]).filter(([r, c, v]) => this.rawText(s, r, c) !== v.slice(1));
      if (head.length) this.writeCells(s, head);
      this.tableChange(s, (l) => [...l, t]);
      if (!this.meta(s).filter) this.filterChange(s, () => { this.meta(s).filter = { range: { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, conds: {} }; });
    });
    return t;
  }
  private tableById(s: number, id: number) { const t = this.meta(s).tables.find((x) => x.id === id); if (!t) throw new Error('Không tìm thấy bảng.'); return t; }
  /** đổi tên / kiểu / tùy chọn hiển thị */
  updateTable(s: number, id: number, patch: Partial<Pick<TableDef, 'name' | 'style' | 'bandR' | 'bandC' | 'first' | 'last' | 'btn'>>) {
    const t = this.tableById(s, id);
    if (patch.name !== undefined && patch.name !== t.name) {
      const e = tableNameError(patch.name, (n) => this.tableNameTaken(n, id));
      if (e) throw new Error(e);
    }
    this.group(() => {
      this.tableChange(s, (l) => l.map((x) => (x.id === id ? { ...x, ...patch } : x)));
      if (patch.btn !== undefined) {
        const f = this.meta(s).filter, mine = f && f.range.r1 === t.r1 && f.range.c1 === t.c1;
        if (!patch.btn && mine) this.clearFilter(s);
        if (patch.btn && !f) this.filterChange(s, () => { this.meta(s).filter = { range: { r1: t.r1, c1: t.c1, r2: dataRows(t).b, c2: t.c2 }, conds: {} }; });
      }
    });
  }
  /** Hàng Tổng: bật thêm hàng SUBTOTAL dưới bảng (chèn hàng nếu bên dưới có dữ liệu), tắt thì xóa */
  setTableTotals(s: number, id: number, on: boolean) {
    const t = this.tableById(s, id);
    if (t.totals === on) return;
    this.group(() => {
      if (on) {
        const below = t.r2 + 1;
        let busy = false;
        for (let c = t.c1; c <= t.c2; c++) if (this.has(s, below, c)) busy = true;
        if (busy) this.insertRows(s, below, 1);
        const cur = this.tableById(s, id);
        // mặc định: chữ "Tổng" ở cột đầu, cột cuối: Tổng nếu là số, ngược lại Đếm
        const cols = cur.cols.map((c, i) => {
          if (c.fn || c.label) return c;
          if (i === 0 && cur.cols.length > 1) return { ...c, label: 'Tổng' };
          if (i === cur.cols.length - 1) return { ...c, fn: (typeof this.value(s, dataRows(cur).b, cur.c2) === 'number' ? 'sum' : 'count') as TotalFn };
          return c;
        });
        this.tableChange(s, (l) => l.map((x) => (x.id === id ? { ...x, r2: x.r2 + 1, totals: true, cols } : x)));
        const nt = this.tableById(s, id);
        const list: [number, number, string | null][] = [];
        nt.cols.forEach((c, i) => {
          const f = c.fn ? totalFormula(c.fn, c.name) : null;
          list.push([nt.r2, nt.c1 + i, f ?? (c.label ? "'" + c.label : null)]);
        });
        this.writeCells(s, list);
      } else {
        const r = t.r2;
        this.tableChange(s, (l) => l.map((x) => (x.id === id ? { ...x, r2: x.r2 - 1, totals: false } : x)));
        this.clear(s, { r1: r, c1: t.c1, r2: r, c2: t.c2 });
      }
    });
  }
  /** chọn hàm cho một cột ở hàng tổng */
  setTotalFn(s: number, id: number, col: number, fn: TotalFn) {
    const t = this.tableById(s, id);
    if (!t.totals) return;
    const i = col - t.c1, c = t.cols[i];
    if (!c) return;
    this.group(() => {
      this.tableChange(s, (l) => l.map((x) => (x.id === id ? { ...x, cols: x.cols.map((y, j) => (j === i ? { ...y, fn, label: undefined } : y)) } : x)));
      const f = totalFormula(fn, c.name);
      this.writeCells(s, [[t.r2, col, f]]);
    });
  }
  /** Đổi kích thước Bảng (Thiết kế Bảng ▸ Đổi kích thước): giữ hàng tiêu đề ở chỗ cũ */
  resizeTable(s: number, id: number, g: Range) {
    const t = this.tableById(s, id);
    if (g.r1 !== t.r1) throw new Error('Tiêu đề phải ở cùng hàng và vùng mới phải chồng lên bảng cũ.');
    if (g.r2 <= g.r1 + (t.totals ? 1 : 0)) throw new Error('Bảng cần ít nhất một hàng dữ liệu.');
    if (this.meta(s).tables.some((x) => x.id !== id && !(x.r2 < g.r1 || x.r1 > g.r2 || x.c2 < g.c1 || x.c1 > g.c2))) throw new Error('Bảng không thể chồng lên một bảng khác.');
    this.group(() => {
      const used = new Set<string>(), cols: TableDef['cols'] = [];
      let nid = Math.max(...t.cols.map((x) => x.id)) + 1;
      for (let c = g.c1; c <= g.c2; c++) {
        const old = c >= t.c1 && c <= t.c2 ? t.cols[c - t.c1] : undefined;
        let nm = old?.name ?? (this.display(s, g.r1, c).text.trim() || `Cột${c - g.c1 + 1}`);
        const base = nm; let k = 2; while (used.has(nm.toLowerCase())) nm = base + k++;
        used.add(nm.toLowerCase()); cols.push(old ? { ...old, name: nm } : { id: nid++, name: nm });
      }
      const head = cols.map((x, i) => [g.r1, g.c1 + i, "'" + x.name] as [number, number, string]).filter(([r, c, v]) => this.rawText(s, r, c) !== v.slice(1));
      if (head.length) this.writeCells(s, head);
      this.moveTotals(s, t, g.r2);
      this.tableChange(s, (l) => l.map((x) => (x.id === id ? { ...x, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, cols } : x)));
      const f = this.meta(s).filter;
      if (f && f.range.r1 === t.r1 && f.range.c1 === t.c1) this.filterChange(s, () => { this.meta(s).filter = { range: { ...g, r2: dataRows(this.tableById(s, id)).b }, conds: {} }; });
    });
  }
  /** công thức (dạng bộ tính) trong mọi sheet tham chiếu tới bảng → đổi về dạng địa chỉ ô */
  private unlinkTable(id: number) {
    const reg = this.tableReg(), t = reg.find((x) => x.id === id);
    if (!t) return;
    const re = new RegExp(`XT_${id}_|XTHISROW\\(`, 'i');
    for (const sid of this.allSheetIds()) {
      const d = this.hf.getSheetDimensions(sid);
      for (let r = 0; r < d.height; r++) for (let c = 0; c < d.width; c++) {
        const a = { sheet: sid, row: r, col: c };
        if (!this.hf.doesCellHaveFormula(a)) continue;
        const f = (this.hf as any).__rawFormula?.(a) as string | undefined;
        if (!f || !re.test(f)) continue;
        const a1 = decodeStructured(f, reg, { sheet: sid, r, c }, 'a1', id);
        this.hf.setCellContents(a, [[decodeStructured(decFormulaNames(offsetOut(a1), nameDisplay(this.hf)), this.tableReg(), { sheet: sid, r, c }) as RawCellContent]]);
      }
    }
  }
  /** Chuyển thành Vùng: giữ định dạng của bảng thành định dạng ô, công thức đổi về địa chỉ ô */
  convertTableToRange(s: number, id: number) {
    const t = this.tableById(s, id);
    this.group(() => {
      const st = this.meta(s).styles, ops: StyleOp[] = [];
      for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) { const v = this.viewStyle(s, r, c); const o = st.get(key(r, c)); if (v !== o) ops.push([st, key(r, c), clean(v)]); }
      const tx = new StyleTxn(ops);
      this.record(() => this.unlinkTable(id), tx.hist, { post: () => tx.apply() });
      const f = this.meta(s).filter;
      if (f && f.range.r1 === t.r1 && f.range.c1 === t.c1) this.clearFilter(s);
      this.tableChange(s, (l) => l.filter((x) => x.id !== id));
    });
  }
  /**
   * Sau khi người dùng nhập ô (r, c): đổi tên cột khi sửa tiêu đề, tự nở bảng khi gõ ngay dưới/bên phải,
   * cột tính toán (công thức nhập vào cột trống được điền cả cột, hàng mới kế thừa công thức).
   * Trả về true nếu có thay đổi bảng (để vẽ lại).
   */
  afterEdit(s: number, r: number, c: number): boolean {
    const l = this.meta(s).tables;
    if (!l.length) return false;
    // sửa tiêu đề
    const th = l.find((t) => t.header && r === t.r1 && c >= t.c1 && c <= t.c2);
    if (th) {
      const i = c - th.c1;
      let nm = this.display(s, r, c).text.trim();
      if (!nm) nm = `Cột${i + 1}`;
      const base = nm; let k = 2;
      while (th.cols.some((x, j) => j !== i && x.name.toLowerCase() === nm.toLowerCase())) nm = base + k++;
      if (nm !== th.cols[i].name) this.tableChange(s, (L) => L.map((x) => (x.id === th.id ? { ...x, cols: x.cols.map((y, j) => (j === i ? { ...y, name: nm } : y)) } : x)));
      if (this.rawText(s, r, c) !== nm) this.writeCells(s, [[r, c, "'" + nm]]);
      return true;
    }
    if (!this.has(s, r, c)) return false;
    // gõ ngay dưới bảng (không có hàng tổng) → thêm hàng
    const tb = l.find((t) => !t.totals && r === t.r2 + 1 && c >= t.c1 && c <= t.c2 && !this.stripBusy(s, t.id, { r1: r, c1: t.c1, r2: r, c2: t.c2 }));
    if (tb) {
      this.tableChange(s, (L) => L.map((x) => (x.id === tb.id ? { ...x, r2: x.r2 + 1 } : x)));
      const fill: [number, number, string][] = [];
      for (let cc = tb.c1; cc <= tb.c2; cc++) {
        if (cc === c || this.has(s, r, cc)) continue;
        const up = this.rawText(s, r - 1, cc), up2 = r - 2 > tb.r1 ? this.rawText(s, r - 2, cc) : up;
        if (up.startsWith('=') && up2.startsWith('=') && shiftFormula(up2, 1, 0) === up) fill.push([r, cc, shiftFormula(up, 1, 0)]);
      }
      if (fill.length) this.writeCells(s, fill);
      const f = this.meta(s).filter;
      if (f && f.range.r1 === tb.r1 && f.range.c1 === tb.c1) this.filterChange(s, () => { f.range = { ...f.range, r2: r }; });
      return true;
    }
    // gõ ngay bên phải → thêm cột
    const tr = l.find((t) => c === t.c2 + 1 && r >= t.r1 && r <= t.r2);
    if (tr) {
      if (this.stripBusy(s, tr.id, { r1: tr.r1, c1: c, r2: tr.r2, c2: c })) return false;
      let nm = r === tr.r1 ? this.display(s, r, c).text.trim() : '';
      let j = tr.cols.length + 1;
      if (!nm) nm = `Cột${j}`;
      while (tr.cols.some((x) => x.name.toLowerCase() === nm.toLowerCase())) nm = `Cột${++j}`;
      const nid = Math.max(...tr.cols.map((x) => x.id)) + 1;
      this.tableChange(s, (L) => L.map((x) => (x.id === tr.id ? { ...x, c2: x.c2 + 1, cols: [...x.cols, { id: nid, name: nm }] } : x)));
      if (this.rawText(s, tr.r1, c) !== nm) this.writeCells(s, [[tr.r1, c, "'" + nm]]);
      const f = this.meta(s).filter;
      if (f && f.range.r1 === tr.r1 && f.range.c1 === tr.c1) this.filterChange(s, () => { f.range = { ...f.range, c2: c }; });
      return true;
    }
    // cột tính toán: công thức vào cột dữ liệu đang trống → điền cả cột
    const tc = l.find((t) => inTable(t, r, c));
    if (tc) {
      const d = dataRows(tc), f = this.rawText(s, r, c);
      if (!f.startsWith('=') || r < d.a || r > d.b || d.b === d.a) return false;
      for (let rr = d.a; rr <= d.b; rr++) if (rr !== r && this.has(s, rr, c)) return false;
      const list: [number, number, string][] = [];
      for (let rr = d.a; rr <= d.b; rr++) if (rr !== r) list.push([rr, c, shiftFormula(f, rr - r, 0)]);
      this.writeCells(s, list);
      return true;
    }
    return false;
  }
  /** tên cột theo ô tiêu đề (sau khi xóa/dán/điền đè hàng tiêu đề): ô trống → Cột<n>, trùng → thêm số */
  syncHeaders(s: number) {
    for (const t of this.meta(s).tables) {
      if (!t.header) continue;
      const used = new Set<string>(), fix: [number, number, string][] = [];
      const cols = t.cols.map((col, i) => {
        let nm = this.display(s, t.r1, t.c1 + i).text.trim();
        if (!nm) nm = `Cột${i + 1}`;
        const base = nm; let k = 2; while (used.has(nm.toLowerCase())) nm = base + k++;
        used.add(nm.toLowerCase());
        if (this.rawText(s, t.r1, t.c1 + i) !== nm) fix.push([t.r1, t.c1 + i, "'" + nm]);
        return nm === col.name ? col : { ...col, name: nm };
      });
      if (fix.length) this.writeCells(s, fix);
      if (cols.some((c, i) => c !== t.cols[i])) this.tableChange(s, (L) => L.map((x) => (x.id === t.id ? { ...x, cols } : x)));
    }
  }
  /** chạy thao tác ghi ô rồi đồng bộ tên cột bảng — một bước hoàn tác */
  withHeaders<T>(s: number, f: () => T): T {
    if (!this.meta(s).tables.length) return f();
    let out!: T;
    this.group(() => { out = f(); this.syncHeaders(s); });
    return out;
  }
  /** dời hàng tổng của bảng tới hàng newR (giữ công thức/nhãn) */
  private moveTotals(s: number, t: TableDef, newR: number) {
    if (!t.totals || newR === t.r2) return;
    const row: [number, number, string | null][] = [];
    for (let c = t.c1; c <= t.c2; c++) row.push([newR, c, this.rawText(s, t.r2, c) || null]);
    this.clear(s, { r1: t.r2, c1: t.c1, r2: t.r2, c2: t.c2 });
    this.writeCells(s, row);
  }
  /** vùng (dải mới khi bảng nở) có đụng bảng khác hoặc ô gộp không */
  private stripBusy(s: number, id: number, g: Range) {
    const m = this.meta(s);
    return m.tables.some((x) => x.id !== id && !(x.r2 < g.r1 || x.r1 > g.r2 || x.c2 < g.c1 || x.c1 > g.c2))
      || m.merges.some((x) => !(x.r2 < g.r1 || x.r1 > g.r2 || x.c2 < g.c1 || x.c1 > g.c2));
  }
  /** công thức dạng lưu tệp .xlsx (Bảng1[[#This Row],[Cột]] thay cho [@Cột]) */
  fileFormula(s: number, r: number, c: number, f: string): string {
    if (!f.includes('[')) return f;
    const reg = this.tableReg();
    if (!reg.length) return f;
    const enc = encodeStructured(f, { tables: reg, self: { sheet: this.sheetName(s), r, c }, need: () => {} });
    return decodeStructured(enc, reg, { sheet: s, r, c }, 'file');
  }

  /** vùng tràn (mảng động) chứa ô (r, c): ô gốc có công thức + vùng kết quả; null nếu không thuộc mảng */
  spillRange(s: number, r: number, c: number): { r: number; c: number; range: Range } | null {
    const part = (rr: number, cc: number) => rr >= 0 && cc >= 0 && this.hf.isCellPartOfArray({ sheet: s, row: rr, col: cc });
    const hasF = (rr: number, cc: number) => this.hf.doesCellHaveFormula({ sheet: s, row: rr, col: cc });
    if (!part(r, c)) return null;
    let r0 = r, c0 = c, n = 0;
    while (!hasF(r0, c0) && part(r0 - 1, c0) && n++ < 100000) r0--;
    n = 0;
    while (!hasF(r0, c0) && part(r0, c0 - 1) && n++ < 16384) c0--;
    if (!hasF(r0, c0)) return null;
    let r2 = r0, c2 = c0;
    while (part(r2 + 1, c0) && !hasF(r2 + 1, c0)) r2++;
    while (part(r0, c2 + 1) && !hasF(r0, c2 + 1)) c2++;
    if (r2 === r0 && c2 === c0) return null;
    return { r: r0, c: c0, range: { r1: r0, c1: c0, r2, c2 } };
  }

  // ---------- nhóm hàng/cột (Dữ liệu ▸ Đại cương) ----------
  private axis(s: number, ax: 'r' | 'c') { return ax === 'r' ? this.meta(s).rowH : this.meta(s).colW; }
  /** ghi một thay đổi đại cương + hàng/cột ẩn thành một bước hoàn tác */
  private outlineChange(s: number, f: () => void) {
    const m = this.meta(s);
    const snap = () => ({ o: m.outline, hr: m.rowH.hidden(), hc: m.colW.hidden() });
    const before = snap();
    f();
    const after = snap();
    const put = (x: ReturnType<typeof snap>) => { m.outline = x.o; m.rowH.setHidden(x.hr); m.colW.setHidden(x.hc); };
    this.record(null, { undo: () => put(before), redo: () => put(after) });
  }
  outlineLevel(s: number, ax: 'r' | 'c', i: number) { return this.meta(s).outline[ax][i] ?? 0; }
  outlineMax(s: number, ax: 'r' | 'c') { let m = 0; for (const v of Object.values(this.meta(s).outline[ax])) if (v > m) m = v; return m; }
  /** các khối liền nhau có cấp ≥ L: [đầu, cuối] */
  outlineRuns(s: number, ax: 'r' | 'c', L: number): { a: number; b: number }[] {
    const idx = Object.entries(this.meta(s).outline[ax]).filter(([, v]) => v >= L).map(([k]) => +k).sort((x, y) => x - y);
    const out: { a: number; b: number }[] = [];
    for (const i of idx) { const last = out[out.length - 1]; if (last && last.b === i - 1) last.b = i; else out.push({ a: i, b: i }); }
    return out;
  }
  isCollapsed(s: number, ax: 'r' | 'c', L: number, b: number) { return this.meta(s).outline.x.includes(`${ax}:${L}:${b + 1}`); }
  /** Nhóm (Alt+Shift+→): tăng cấp các hàng/cột a..b (tối đa 7 cấp như Excel) */
  groupLines(s: number, ax: 'r' | 'c', a: number, b: number) {
    const cur = this.meta(s).outline[ax];
    for (let i = a; i <= b; i++) if ((cur[i] ?? 0) >= 7) throw new Error('Đã đạt tối đa 7 cấp nhóm.');
    this.outlineChange(s, () => {
      const m = this.meta(s), lv = { ...m.outline[ax] };
      for (let i = a; i <= b; i++) lv[i] = (lv[i] ?? 0) + 1;
      m.outline = { ...m.outline, [ax]: lv };
    });
  }
  /** Bỏ nhóm (Alt+Shift+←): giảm cấp; hàng/cột hết thuộc nhóm thì hiện lại */
  ungroupLines(s: number, ax: 'r' | 'c', a: number, b: number) {
    const cur = this.meta(s).outline[ax];
    let any = false;
    for (let i = a; i <= b; i++) if (cur[i]) any = true;
    if (!any) throw new Error('Không có nhóm nào trong vùng chọn để bỏ nhóm.');
    this.outlineChange(s, () => {
      const m = this.meta(s), lv = { ...m.outline[ax] }, A = this.axis(s, ax);
      const show: number[] = [];
      for (let i = a; i <= b; i++) { if (!lv[i]) continue; lv[i]--; if (!lv[i]) { delete lv[i]; show.push(i); } }
      const hid = new Set(A.hidden());
      if (show.some((i) => hid.has(i))) A.setHidden([...hid].filter((i) => !show.includes(i)));
      const max = Math.max(0, ...Object.values(lv));
      m.outline = { ...m.outline, [ax]: lv, x: m.outline.x.filter((k) => !k.startsWith(ax + ':') || +k.split(':')[1] <= max) };
    });
  }
  /** Thu gọn (−) nhóm cấp L kết thúc ở b */
  collapseGroup(s: number, ax: 'r' | 'c', L: number, a: number, b: number) {
    this.outlineChange(s, () => {
      const m = this.meta(s), A = this.axis(s, ax), hid = new Set(A.hidden());
      for (let i = a; i <= b; i++) hid.add(i);
      A.setHidden([...hid]);
      const k = `${ax}:${L}:${b + 1}`;
      if (!m.outline.x.includes(k)) m.outline = { ...m.outline, x: [...m.outline.x, k] };
    });
  }
  /** Mở rộng (+): hiện lại, trừ các nhóm con vẫn đang thu gọn */
  expandGroup(s: number, ax: 'r' | 'c', L: number, a: number, b: number) {
    this.outlineChange(s, () => {
      const m = this.meta(s), A = this.axis(s, ax);
      const x = m.outline.x.filter((k) => k !== `${ax}:${L}:${b + 1}`);
      const keep = new Set<number>();
      for (const k of x) {
        const [t, LL, end] = k.split(':');
        if (t !== ax || +LL <= L) continue;
        const run = this.outlineRuns(s, ax, +LL).find((g) => g.b + 1 === +end);
        if (run && run.a >= a && run.b <= b) for (let i = run.a; i <= run.b; i++) keep.add(i);
      }
      const hid = A.hidden().filter((i) => i < a || i > b || keep.has(i));
      A.setHidden(hid);
      m.outline = { ...m.outline, x };
    });
  }
  /** nút cấp 1, 2, 3…: hiện tới cấp n (cấp ≥ n bị thu gọn) */
  showOutlineLevel(s: number, ax: 'r' | 'c', n: number) {
    this.outlineChange(s, () => {
      const m = this.meta(s), A = this.axis(s, ax), lv = m.outline[ax];
      const inGroup = new Set(Object.keys(lv).map(Number));
      const hid = A.hidden().filter((i) => !inGroup.has(i));
      for (const [k, v] of Object.entries(lv)) if (v >= n) hid.push(+k);
      A.setHidden(hid);
      const x = m.outline.x.filter((k) => !k.startsWith(ax + ':'));
      for (let L = n; L <= 7; L++) for (const g of this.outlineRuns(s, ax, L)) x.push(`${ax}:${L}:${g.b + 1}`);
      m.outline = { ...m.outline, x };
    });
  }
  /** Bỏ Đại cương: xóa mọi nhóm theo trục, hiện lại hàng/cột đã thu gọn */
  clearOutline(s: number, ax: 'r' | 'c') {
    this.outlineChange(s, () => {
      const m = this.meta(s), A = this.axis(s, ax), lv = m.outline[ax];
      const inGroup = new Set(Object.keys(lv).map(Number));
      A.setHidden(A.hidden().filter((i) => !inGroup.has(i)));
      m.outline = { ...m.outline, [ax]: {}, x: m.outline.x.filter((k) => !k.startsWith(ax + ':')) };
    });
  }
  /**
   * Tổng phụ (Dữ liệu ▸ Tổng phụ): chèn hàng tổng sau mỗi nhóm giá trị của cột `by`, hàng Tổng Cộng ở cuối,
   * dùng SUBTOTAL(fn; …) và tạo nhóm 3 cấp như Excel. g: vùng dữ liệu kể cả hàng tiêu đề.
   */
  subtotal(s: number, g: Range, o: { by: number; fn: number; cols: number[]; replace?: boolean }) {
    if (g.r2 <= g.r1) throw new Error('Vùng cần có hàng tiêu đề và ít nhất một hàng dữ liệu.');
    if (!o.cols.length) throw new Error('Hãy chọn ít nhất một cột để tính tổng phụ.');
    const FN: Record<number, string> = { 9: 'Tổng', 1: 'Trung bình', 2: 'Đếm', 3: 'Đếm', 4: 'Lớn nhất', 5: 'Nhỏ nhất', 6: 'Tích', 7: 'Độ lệch chuẩn', 10: 'Phương sai' };
    let r2 = g.r2;
    this.group(() => {
      if (o.replace) r2 = this.removeSubtotals(s, g);
      // nhóm theo giá trị liên tiếp của cột `by`
      const groups: { a: number; b: number; label: string }[] = [];
      for (let r = g.r1 + 1; r <= r2; r++) {
        const t = this.display(s, r, o.by).text;
        const last = groups[groups.length - 1];
        if (last && last.label === t) last.b = r; else groups.push({ a: r, b: r, label: t });
      }
      if (!groups.length) return;
      const name = FN[o.fn] ?? 'Tổng';
      const cells = (r: number, a: number, b: number) => o.cols.map((c) => [r, c, `=SUBTOTAL(${o.fn};${colLetters(c)}${a + 1}:${colLetters(c)}${b + 1})`] as [number, number, string]);
      // chèn từ dưới lên: chỉ số các nhóm phía trên không đổi
      for (let i = groups.length - 1; i >= 0; i--) {
        const gr = groups[i];
        this.insertRows(s, gr.b + 1, 1);
        this.writeCells(s, [[gr.b + 1, o.by, `${name} ${gr.label}`.trim()], ...cells(gr.b + 1, gr.a, gr.b)]);
        this.setStyle(s, { r1: gr.b + 1, c1: g.c1, r2: gr.b + 1, c2: g.c2 }, { b: true });
      }
      const last = r2 + groups.length;
      this.insertRows(s, last + 1, 1);
      // Tổng Cộng tính trên các khối chi tiết (không cộng lặp các hàng tổng phụ)
      const blocks: string[] = [];
      let at = g.r1 + 1;
      for (const gr of groups) { const n = gr.b - gr.a + 1; blocks.push(`${at + 1}:${at + n}`); at += n + 1; }
      const grand = o.cols.map((c) => [last + 1, c, `=SUBTOTAL(${o.fn};${blocks.map((b) => { const [x, y] = b.split(':'); return `${colLetters(c)}${x}:${colLetters(c)}${y}`; }).join(';')})`] as [number, number, string]);
      this.writeCells(s, [[last + 1, o.by, `${name} Cộng`], ...grand]);
      this.setStyle(s, { r1: last + 1, c1: g.c1, r2: last + 1, c2: g.c2 }, { b: true });
      // đại cương: chi tiết cấp 2, hàng tổng phụ cấp 1, Tổng Cộng ngoài nhóm
      this.outlineChange(s, () => {
        const m = this.meta(s), lv = { ...m.outline.r };
        let r = g.r1 + 1;
        for (const gr of groups) { const n = gr.b - gr.a + 1; for (let k = 0; k < n; k++) lv[r + k] = 2; lv[r + n] = 1; r += n + 1; }
        m.outline = { ...m.outline, r: lv };
      });
    });
  }
  /** Xóa Tất cả tổng phụ: bỏ các hàng có SUBTOTAL trong vùng và đại cương hàng; trả về hàng cuối mới */
  removeSubtotals(s: number, g: Range): number {
    const u = this.used(s);
    const rows: number[] = [];
    for (let r = g.r1 + 1; r <= Math.min(u.rows - 1, g.r2 + 10000); r++) {
      let hit = false, any = false;
      for (let c = g.c1; c <= g.c2; c++) { const f = this.rawText(s, r, c); if (f) any = true; if (/^=SUBTOTAL\(/i.test(f)) hit = true; }
      if (hit) rows.push(r);
      else if (!any && r > g.r2) break;
    }
    let end = g.r2;
    this.group(() => {
      for (const r of rows.reverse()) { this.deleteRows(s, r, 1); if (r <= end) end--; }
      this.clearOutline(s, 'r');
    });
    return Math.max(g.r1, end - 0);
  }

  // ---------- ghi chú & chú thích ----------
  /** đặt/xóa ghi chú (text rỗng = xóa) */
  setNote(s: number, r: number, c: number, text: string | null) {
    const st = this.meta(s).styles, k = key(r, c), o = st.get(k);
    const v = text && text.trim() ? text.replace(/\s+$/, '') : undefined;
    if ((o?.nt ?? undefined) === v) return;
    const tx = new StyleTxn([[st, k, clean({ ...(o ?? {}), nt: v })]]);
    tx.apply();
    this.record(null, tx.hist);
  }
  /** thêm bình luận vào chuỗi chú thích (tạo mới nếu chưa có) */
  addComment(s: number, r: number, c: number, author: string, text: string, when = new Date()) {
    const t = text.trim();
    if (!t) return;
    const st = this.meta(s).styles, k = key(r, c), o = st.get(k);
    const list = [...(o?.cm ?? []), { a: author.trim() || 'Người dùng', t, d: when.toISOString() }];
    const tx = new StyleTxn([[st, k, clean({ ...(o ?? {}), cm: list, cmDone: undefined })]]);
    tx.apply();
    this.record(null, tx.hist);
  }
  /** sửa / xóa một bình luận trong chuỗi (xóa bình luận đầu = xóa cả chuỗi như Excel) */
  editComment(s: number, r: number, c: number, i: number, text: string | null) {
    const st = this.meta(s).styles, k = key(r, c), o = st.get(k);
    if (!o?.cm?.[i]) return;
    let list: CommentItem[] | undefined = o.cm.map((x) => ({ ...x }));
    if (text === null) list = i === 0 ? undefined : list.filter((_, j) => j !== i);
    else list[i].t = text.trim();
    const tx = new StyleTxn([[st, k, clean({ ...o, cm: list, cmDone: list ? o.cmDone : undefined })]]);
    tx.apply();
    this.record(null, tx.hist);
  }
  resolveComment(s: number, r: number, c: number, done: boolean) {
    const st = this.meta(s).styles, k = key(r, c), o = st.get(k);
    if (!o?.cm) return;
    const tx = new StyleTxn([[st, k, clean({ ...o, cmDone: done || undefined })]]);
    tx.apply();
    this.record(null, tx.hist);
  }
  /** xóa ghi chú + chú thích trong vùng (Trang đầu ▸ Xóa ▸ Xóa Ghi chú và Chú thích) */
  clearNotes(s: number, g0: Range, which: 'all' | 'notes' | 'comments' = 'all') {
    const st = this.meta(s).styles, ops: StyleOp[] = [];
    const g = g0;
    st.forEach((o, k) => {
      const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS;
      if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2) return;
      if (!o.nt && !o.cm) return;
      const n = { ...o };
      if (which !== 'comments') delete n.nt;
      if (which !== 'notes') { delete n.cm; delete n.cmDone; }
      ops.push([st, k, clean(n)]);
    });
    if (!ops.length) return 0;
    const tx = new StyleTxn(ops);
    tx.apply();
    this.record(null, tx.hist);
    return ops.length;
  }
  /** mọi ô có ghi chú / chú thích của trang tính, theo thứ tự đọc (hàng rồi cột) như Excel đi Trước/Tiếp */
  private noteCache = new Map<string, { ver: number; list: { r: number; c: number }[] }>();
  noteCells(s: number, kind: 'note' | 'comment' | 'any' = 'any'): { r: number; c: number }[] {
    // lưới gọi mỗi lần vẽ → nhớ theo phiên bản sổ (duyệt bảng kiểu ô rất lớn tốn thời gian)
    const ck = s + ':' + kind, hit = this.noteCache.get(ck);
    if (hit && hit.ver === this.ver) return hit.list;
    const out: { r: number; c: number }[] = [];
    this.noteCache.set(ck, { ver: this.ver, list: out });
    this.meta(s).styles.forEach((o, k) => { if ((kind !== 'comment' && o.nt) || (kind !== 'note' && o.cm)) out.push({ r: Math.floor(k / MAX_COLS), c: k % MAX_COLS }); });
    return out.sort((a, b) => a.r - b.r || a.c - b.c);
  }

  // ---------- tên vùng (Công thức ▸ Tên đã Xác định) ----------
  /** mọi tên do người dùng xác định (không gồm TRUE/FALSE dựng sẵn) */
  listNames(): NameDef[] {
    const disp = nameDisplay(this.hf);
    return this.hf.getAllNamedExpressionsSerialized()
      .filter((n) => !BUILTIN_NAMES.has(n.name.toUpperCase()) && !/^XT_\d+_/i.test(n.name))
      .map((n) => {
        const o: any = n.options ?? {};
        return { name: disp.get(n.name.toUpperCase()) ?? n.name, expr: decFormulaNames(offsetOut(String(n.expression)), disp), scope: n.scope, comment: o.comment || undefined, hidden: o.hidden || undefined };
      })
      .sort((a, b) => COLLATOR.compare(a.name, b.name) || (a.scope ?? -1) - (b.scope ?? -1));
  }
  /** giá trị hiện tại của tên (để hiện trong Trình Quản lý Tên) */
  nameValue(name: string, scope?: number): unknown {
    return this.hf.getNamedExpressionValue(encName(name), scope);
  }
  private findName(name: string, scope?: number) {
    return this.listNames().find((n) => n.name.toUpperCase() === name.trim().toUpperCase() && (n.scope ?? -1) === (scope ?? -1));
  }
  /** chuẩn hóa “Tham chiếu tới” → công thức cho bộ tính; lỗi thì ném thông báo tiếng Việt */
  private nameExpr(name: string, expr: string, scope?: number) {
    let e = expr.trim();
    if (!e) throw new Error('Hãy nhập ô hoặc công thức mà tên tham chiếu tới.');
    if (e[0] !== '=') e = '=' + e;
    const enc = encName(name.trim());
    const fx = encFormulaNames(offsetIn(e, undefined, ''));
    if (!this.hf.isItPossibleToAddNamedExpression(enc, fx, scope) && !this.hf.getNamedExpression(enc, scope)) throw new Error('Công thức trong “Tham chiếu tới” không hợp lệ.');
    return { enc, fx };
  }
  /** Xác định Tên (Alt+M,M,D): thêm tên mới */
  defineName(name: string, expr: string, scope?: number, comment?: string) {
    const err = nameError(name);
    if (err) throw new Error(err);
    if (this.findName(name, scope)) throw new Error(`Tên “${name.trim()}” đã có trong phạm vi này.`);
    if (this.tableByName(name.trim()) || /^XT_\d+_/i.test(name.trim())) throw new Error(`Tên “${name.trim()}” đã được dùng cho một bảng.`);
    const { enc, fx } = this.nameExpr(name, expr, scope);
    nameDisplay(this.hf).set(enc.toUpperCase(), name.trim());
    this.record(() => this.hf.addNamedExpression(enc, fx, scope, comment ? { comment } : undefined));
  }
  /** Sửa tên (đổi tên, phạm vi, tham chiếu, chú thích); đổi tên thì công thức đang dùng tên cũ được cập nhật như Excel */
  updateName(oldName: string, oldScope: number | undefined, d: { name: string; expr: string; scope?: number; comment?: string }) {
    const cur = this.findName(oldName, oldScope);
    if (!cur) throw new Error(`Không tìm thấy tên “${oldName}”.`);
    const err = nameError(d.name);
    if (err) throw new Error(err);
    const renamed = d.name.trim().toUpperCase() !== cur.name.toUpperCase() || (d.scope ?? -1) !== (cur.scope ?? -1);
    if (renamed && this.findName(d.name, d.scope)) throw new Error(`Tên “${d.name.trim()}” đã có trong phạm vi này.`);
    const { enc, fx } = this.nameExpr(d.name, d.expr, renamed ? d.scope : cur.scope);
    const opts = d.comment ? { comment: d.comment } : undefined;
    this.group(() => {
      if (!renamed) { nameDisplay(this.hf).set(enc.toUpperCase(), d.name.trim()); this.record(() => this.hf.changeNamedExpression(enc, fx, cur.scope, opts)); return; }
      nameDisplay(this.hf).set(enc.toUpperCase(), d.name.trim());
      this.record(() => this.hf.addNamedExpression(enc, fx, d.scope, opts));
      if (d.name.trim().toUpperCase() !== cur.name.toUpperCase()) this.renameRefs(cur.name, d.name.trim());
      this.record(() => this.hf.removeNamedExpression(encName(cur.name), cur.scope));
    });
  }
  deleteName(name: string, scope?: number) {
    const cur = this.findName(name, scope);
    if (!cur) return;
    this.record(() => this.hf.removeNamedExpression(encName(cur.name), cur.scope));
  }
  /** cập nhật mọi công thức đang dùng tên cũ */
  private renameRefs(oldName: string, newName: string) {
    for (const s of this.sheetIds()) {
      const rows = this.hf.getSheetSerialized(s);
      const list: [number, number, string][] = [];
      rows.forEach((row, r) => row?.forEach((v, c) => {
        if (typeof v !== 'string' || v[0] !== '=') return;
        const nv = renameInFormula(v, oldName, newName);
        if (nv !== v) list.push([r, c, nv]);
      }));
      if (list.length) this.record(() => { for (const [r, c, v] of list) this.hf.setCellContents(this.A(s, r, c), [[v]]); });
    }
  }
  /**
   * Tạo từ Vùng chọn (Ctrl+Shift+F3): lấy chữ ở hàng trên/cột trái/hàng dưới/cột phải làm tên.
   * Trả về số tên đã tạo; tên trùng được thay (như khi chọn “Có” ở Excel).
   */
  createNamesFrom(s: number, g: Range, o: { top?: boolean; left?: boolean; bottom?: boolean; right?: boolean }): number {
    const sh = quoteSheet(this.sheetName(s));
    const abs = (r1: number, c1: number, r2: number, c2: number) => `=${sh}!$${colLetters(c1)}$${r1 + 1}` + (r1 === r2 && c1 === c2 ? '' : `:$${colLetters(c2)}$${r2 + 1}`);
    const clean = (t: string) => { let n = t.trim().replace(/\s+/g, '_').replace(/[^A-Za-z0-9_.\\\u00C0-\uFFFF]/g, ''); if (n && !/^[A-Za-z_\\\u00C0-\uFFFF]/.test(n)) n = '_' + n; return nameError(n) ? '' : n; };
    const items: [string, string][] = [];
    const r1 = g.r1 + (o.top ? 1 : 0), r2 = g.r2 - (o.bottom ? 1 : 0), c1 = g.c1 + (o.left ? 1 : 0), c2 = g.c2 - (o.right ? 1 : 0);
    if (r1 > r2 || c1 > c2) return 0;
    for (let c = c1; c <= c2; c++) {
      if (o.top) { const n = clean(this.display(s, g.r1, c).text); if (n) items.push([n, abs(r1, c, r2, c)]); }
      if (o.bottom) { const n = clean(this.display(s, g.r2, c).text); if (n) items.push([n, abs(r1, c, r2, c)]); }
    }
    for (let r = r1; r <= r2; r++) {
      if (o.left) { const n = clean(this.display(s, r, g.c1).text); if (n) items.push([n, abs(r, c1, r, c2)]); }
      if (o.right) { const n = clean(this.display(s, r, g.c2).text); if (n) items.push([n, abs(r, c1, r, c2)]); }
    }
    let made = 0;
    this.group(() => {
      for (const [n, e] of items) {
        const cur = this.findName(n);
        try { if (cur) this.updateName(cur.name, undefined, { name: cur.name, expr: e, comment: cur.comment }); else this.defineName(n, e); made++; } catch { /* bỏ qua tên lỗi */ }
      }
    });
    return made;
  }
  /** tên trỏ tới một vùng ô → vị trí (để Hộp Tên / Đi tới nhảy tới) */
  nameTarget(name: string, scopeSheet?: number): { sheet: number; range: Range } | null {
    const n = (scopeSheet !== undefined ? this.findName(name, scopeSheet) : undefined) ?? this.findName(name);
    if (!n) return null;
    const m = /^=\s*(?:'((?:[^']|'')+)'|([^'!=]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?\s*$/i.exec(n.expr);
    if (!m) return null;
    const sheet = this.hf.getSheetId((m[1] ?? m[2]).replace(/''/g, "'"));
    if (sheet === undefined) return null;
    const g = parseRange(m[3] + m[4] + (m[5] ? ':' + m[5] + m[6] : ''));
    return g ? { sheet, range: g } : null;
  }

  // ---------- lọc ----------
  isRowHidden(s: number, r: number) { return this.meta(s).rowH.isHidden(r); }
  private filterSnap(s: number) { const m = this.meta(s); return { f: m.filter ? JSON.parse(JSON.stringify(m.filter)) as SheetFilter : null, h: m.rowH.hidden() }; }
  private filterRestore(s: number, x: { f: SheetFilter | null; h: number[] }) { const m = this.meta(s); m.filter = x.f ? JSON.parse(JSON.stringify(x.f)) : null; m.rowH.setHidden(x.h); }
  private filterChange(s: number, f: () => void) {
    const before = this.filterSnap(s);
    f();
    const after = this.filterSnap(s);
    this.record(null, { undo: () => this.filterRestore(s, before), redo: () => this.filterRestore(s, after) });
  }
  setFilter(s: number, g0: Range) {
    let g = g0;
    if (g.r1 === g.r2 && g.c1 === g.c2) g = this.currentRegion(s, g.r1, g.c1);
    const u = this.used(s);
    if (g.r2 >= u.rows) g = { ...g, r2: Math.max(g.r1 + 1, u.rows - 1) };
    if (g.c2 >= u.cols && u.cols > 0) g = { ...g, c2: Math.max(g.c1, u.cols - 1) };
    if (g.r1 === g.r2) g = { ...g, r2: Math.min(MAX_ROWS - 1, g.r2 + 1) };
    this.filterChange(s, () => { const m = this.meta(s); if (m.filter) this.setFilterHidden(s, m.filter.range, []); m.filter = { range: g, conds: {} }; this.setFilterHidden(s, g, []); });
  }
  clearFilter(s: number) {
    if (!this.meta(s).filter) return;
    this.filterChange(s, () => { const m = this.meta(s); this.setFilterHidden(s, m.filter!.range, []); m.filter = null; });
  }
  /**
   * Vùng lọc thực tế như Excel: hàng tiêu đề + mọi hàng có dữ liệu liền kề bên dưới (tới hàng trống hoàn toàn).
   * Không sửa vùng đã lưu (chỉ khi áp dụng lọc mới cập nhật, như Excel).
   * Tệp .xlsx đôi khi chỉ ghi vùng lọc là hàng tiêu đề (vd A2:K2) — khi đó vẫn lấy được dữ liệu bên dưới.
   */
  filterRange(s: number): Range | null {
    const f = this.meta(s).filter;
    if (!f) return null;
    const u = this.used(s);
    const g = f.range;
    const rowHas = (r: number) => { for (let c = g.c1; c <= g.c2; c++) if (this.has(s, r, c)) return true; return false; };
    // lọc của Bảng: đúng vùng tiêu đề + dữ liệu (không gồm hàng tổng)
    const t = this.meta(s).tables.find((x) => x.header && x.r1 === g.r1 && x.c1 === g.c1);
    if (t) { const d = dataRows(t); return { r1: t.r1, c1: t.c1, r2: d.b, c2: t.c2 }; }
    let r2 = Math.min(g.r2, Math.max(g.r1, u.rows - 1));
    while (r2 + 1 < u.rows && rowHas(r2 + 1)) r2++;
    return r2 === g.r2 ? g : { ...g, r2 };
  }
  filterValues(s: number, col: number): string[] {
    const f = this.meta(s).filter;
    if (!f) return [];
    const set = new Set<string>();
    const r2 = Math.min(this.filterRange(s)!.r2, this.used(s).rows - 1);
    for (let r = f.range.r1 + 1; r <= r2; r++) set.add(this.display(s, r, col).text);
    return [...set].sort((a, b) => (a === '' ? 1 : b === '' ? -1 : COLLATOR.compare(a, b)));
  }
  filterColumn(s: number, col: number, cond: FilterCond | null) {
    const f = this.meta(s).filter;
    if (!f) return;
    this.filterChange(s, () => {
      if (cond && (cond.values || cond.text)) f.conds[col] = cond; else delete f.conds[col];
      this.applyFilter(s);
    });
  }
  /** Tính lại các hàng ẩn theo điều kiện lọc hiện tại */
  applyFilter(s: number) {
    const m = this.meta(s), f = m.filter;
    if (!f) return;
    // mở rộng vùng lọc xuống các hàng dữ liệu mới liền kề
    f.range = this.filterRange(s)!;
    const r2 = f.range.r2;
    const hidden: number[] = [];
    const conds = Object.entries(f.conds).map(([c, v]) => [+c, v] as const);
    for (let r = f.range.r1 + 1; r <= r2; r++) {
      for (const [c, cd] of conds) {
        const t = this.display(s, r, c).text;
        if ((cd.values && !cd.values.includes(t)) || (cd.text && !t.toLowerCase().includes(cd.text.toLowerCase()))) { hidden.push(r); break; }
      }
    }
    this.setFilterHidden(s, f.range, hidden);
  }
  /** Đặt các hàng ẩn trong vùng lọc, giữ nguyên hàng ẩn thủ công ngoài vùng lọc */
  private setFilterHidden(s: number, g: Range, hidden: number[]) {
    const A = this.meta(s).rowH;
    A.setHidden([...A.hidden().filter((r) => r <= g.r1 || r > g.r2), ...hidden]);
  }
  reapplyFilter(s: number) { if (this.meta(s).filter) this.filterChange(s, () => this.applyFilter(s)); }

  // ---------- định dạng có điều kiện ----------
  private ruleChange(s: number, key2: 'cf' | 'dv', next: (list: any[]) => any[]) {
    const m = this.meta(s);
    const before = m[key2] as any[], after = next(before);
    (m as any)[key2] = after;
    this.record(null, { undo: () => ((this.meta(s) as any)[key2] = before), redo: () => ((this.meta(s) as any)[key2] = after) });
  }
  private uid() { return Math.random().toString(36).slice(2, 10); }
  addCF(s: number, rule: CFRule) { const r = { ...rule, id: rule.id ?? this.uid() }; this.ruleChange(s, 'cf', (l) => [...l, r]); return r.id; }
  removeCF(s: number, id: string) { this.ruleChange(s, 'cf', (l) => l.filter((x: CFRule) => x.id !== id)); }
  clearCF(s: number, g?: Range) { this.ruleChange(s, 'cf', (l) => (g ? l.filter((x: CFRule) => x.range.r2 < g.r1 || x.range.r1 > g.r2 || x.range.c2 < g.c1 || x.range.c1 > g.c2) : [])); }
  private cfCache = new Map<string, { ver: number; st: CFStats }>();
  cfStyle(s: number, r: number, c: number): CFResult {
    const list = this.meta(s).cf;
    if (!list.length) return {};
    const out: CFResult = {};
    let v: unknown, t: string | null = null;
    for (const rule of list) {
      if (rule.raw || !inR(rule.range, r, c)) continue;
      if (t === null) { v = this.value(s, r, c); t = this.display(s, r, c).text; }
      let st: CFStats | null = null;
      if (!['cellIs', 'text'].includes(rule.type)) {
        const ck = s + ':' + rule.id;
        const hit = this.cfCache.get(ck);
        if (hit && hit.ver === this.ver) st = hit.st;
        else {
          const u = this.used(s), vals: unknown[] = [];
          const g = rule.range;
          for (let rr = g.r1; rr <= Math.min(g.r2, u.rows - 1); rr++) for (let cc = g.c1; cc <= Math.min(g.c2, u.cols - 1); cc++) vals.push(this.value(s, rr, cc));
          st = cfStats(rule, vals);
          this.cfCache.set(ck, { ver: this.ver, st });
        }
      }
      const res = cfEval(rule, v, t, st);
      if (!res) continue;
      for (const [k, val] of Object.entries(res)) if ((out as any)[k] === undefined && val !== undefined) (out as any)[k] = val;
    }
    return out;
  }

  // ---------- kiểm tra dữ liệu ----------
  addDV(s: number, rule: DVRule) { const r = { ...rule, id: rule.id ?? this.uid() }; this.ruleChange(s, 'dv', (l) => [...l.filter((x: DVRule) => !(x.range.r1 === r.range.r1 && x.range.c1 === r.range.c1 && x.range.r2 === r.range.r2 && x.range.c2 === r.range.c2)), r]); return r.id; }
  clearDV(s: number, g: Range) { this.ruleChange(s, 'dv', (l) => l.filter((x: DVRule) => x.range.r2 < g.r1 || x.range.r1 > g.r2 || x.range.c2 < g.c1 || x.range.c1 > g.c2)); }
  dvAt(s: number, r: number, c: number): DVRule | null {
    const l = this.meta(s).dv;
    for (let i = l.length - 1; i >= 0; i--) if (inR(l[i].range, r, c)) return l[i];
    return null;
  }
  dvList(s: number, r: number, c: number): string[] {
    const d = this.dvAt(s, r, c);
    if (!d || d.type !== 'list') return [];
    if (d.list) return d.list;
    if (!d.listRef) return [];
    let ref = d.listRef.replace(/^=/, '').replace(/\$/g, '');
    let sheet = s;
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ref);
    if (m) { const id = this.hf.getSheetId((m[1] ?? m[2]).replace(/''/g, "'")); if (id === undefined) return []; sheet = id; ref = m[3]; }
    const g = parseRef(ref);
    if (!g) return [];
    const out: string[] = [];
    const lastRow = Math.min(g.r2, this.used(sheet).rows - 1, g.r1 + 5000);
    for (let rr = g.r1; rr <= lastRow; rr++) for (let cc = g.c1; cc <= g.c2; cc++) { const t = this.display(sheet, rr, cc).text; if (t) out.push(t); }
    return out;
  }
  /** Kiểm tra giá trị người dùng gõ; trả về null nếu hợp lệ, hoặc câu báo lỗi */
  validate(s: number, r: number, c: number, text: string): string | null {
    const d = this.dvAt(s, r, c);
    if (!d || d.raw || d.type === 'custom') return null;
    if (text.trim() === '') return d.allowBlank === false ? d.error || 'Ô này không được để trống.' : null;
    if (text.startsWith('=')) return null;
    const bad = () => d.error || dvDescribe(d);
    if (d.type === 'list') {
      // danh sách lấy từ tên vùng, INDIRECT… mà app chưa đọc được → không chặn
      if (!d.list && d.listRef && !/^'?[^!]*'?!?\$?[A-Za-z]{1,3}\$?\d*(:\$?[A-Za-z]{1,3}\$?\d*)?$/.test(d.listRef.replace(/^=/, ''))) return null;
      const items = this.dvList(s, r, c).map((x) => x.trim().toLowerCase());
      if (!items.length && d.listRef) return null;
      return items.includes(text.trim().toLowerCase()) ? null : bad();
    }
    if (d.type === 'textLength') {
      const a = numOf(d.v1), b = numOf(d.v2);
      return a !== null && compare(d.op ?? 'between', text.length, a, b ?? undefined) ? null : bad();
    }
    const x = numOf(text, dateSerial), a = numOf(d.v1, dateSerial), b = numOf(d.v2, dateSerial);
    if (x === null || a === null) return bad();
    if (d.type === 'whole' && !Number.isInteger(x)) return bad();
    return compare(d.op ?? 'between', x, a, b ?? undefined) ? null : bad();
  }

  // ---------- biểu đồ ----------
  private chartsChange(s: number, next: (l: Chart[]) => Chart[]) {
    const m = this.meta(s);
    const before = m.charts, after = next(before);
    m.charts = after;
    this.record(null, { undo: () => (this.meta(s).charts = before), redo: () => (this.meta(s).charts = after) });
  }
  addChart(s: number, ch: Chart): string {
    const c = { legend: 'b' as const, ...ch, id: ch.id ?? this.uid() };
    this.chartsChange(s, (l) => [...l, c]);
    return c.id;
  }
  updateChart(s: number, id: string, patch: Partial<Chart>) {
    this.chartsChange(s, (l) => l.map((c) => (c.id === id ? { ...c, ...patch, id } : c)));
  }
  removeChart(s: number, id: string) { this.chartsChange(s, (l) => l.filter((c) => c.id !== id)); }

  // ---------- PivotTable ----------
  /** Đọc vùng nguồn thành bảng: hàng đầu là tên cột */
  pivotTable(src: Ref) {
    const u = this.used(src.sheet);
    const g = { ...src.range, r2: Math.min(src.range.r2, u.rows - 1), c2: Math.min(src.range.c2, u.cols - 1) };
    // tên cột trùng hoặc trống → thêm số thứ tự cho phân biệt (như Excel: “Tên2”)
    const header: string[] = [], seen = new Set<string>();
    for (let c = g.c1; c <= g.c2; c++) {
      const base = this.display(src.sheet, g.r1, c).text.trim() || `Cột ${c - g.c1 + 1}`;
      let n = base, k = 2;
      while (seen.has(n.toLowerCase())) n = base + k++;
      seen.add(n.toLowerCase()); header.push(n);
    }
    const rows: unknown[][] = [], text: (string | undefined)[][] = [];
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      const row: unknown[] = [], tx: (string | undefined)[] = [];
      for (let c = g.c1; c <= g.c2; c++) {
        const v = this.value(src.sheet, r, c);
        row.push(v && typeof v === 'object' ? null : v);
        // ngày/giờ: nhãn theo chữ hiển thị (không phải số sê-ri)
        tx.push(typeof v === 'number' && isDateFmt(this.effFmt(src.sheet, r, c)) ? this.display(src.sheet, r, c).text : undefined);
      }
      rows.push(row); text.push(tx);
    }
    return { header, rows, text };
  }
  pivotAt(s: number, r: number, c: number): PivotDef | null {
    return this.meta(s).pivots.find((p) => p.area && r >= p.area.r1 && r <= p.area.r2 && c >= p.area.c1 && c <= p.area.c2) ?? null;
  }
  /** Tạo Pivot: mặc định ở sheet mới “Pivot n”, ô A3 */
  addPivot(def: Omit<PivotDef, 'at'> & { at?: { r: number; c: number } }, target?: number): { sheet: number; id: string } {
    if (!this.sheetIds().includes(def.src.sheet)) throw new Error('Không tìm thấy sheet nguồn.');
    const n0 = this.past.length;
    let sheet = target;
    if (sheet === undefined) { let i = 1; while (this.hasName('Pivot ' + i)) i++; sheet = this.addSheet('Pivot ' + i); }
    const p: PivotDef = { ...def, id: def.id ?? this.uid(), at: def.at ?? { r: 2, c: 0 }, area: null };
    try { this.writePivot(sheet, p, true); }
    catch (e) { if (target === undefined && this.past.length > n0) { this.undo(); this.future.pop(); } throw e; }
    // tạo sheet + ghi Pivot = một bước hoàn tác
    if (this.past.length - n0 > 1) { const parts = this.past.splice(n0); this.past.push({ hf: false, parts }); }
    return { sheet, id: p.id! };
  }
  refreshPivot(s: number, id: string, patch?: Partial<PivotSpec>) {
    const p = this.meta(s).pivots.find((x) => x.id === id);
    if (!p) throw new Error('Không tìm thấy PivotTable.');
    return this.writePivot(s, { ...p, ...(patch ?? {}) }, false);
  }
  removePivot(s: number, id: string) {
    const m = this.meta(s), p = m.pivots.find((x) => x.id === id);
    if (!p) return;
    const before = m.pivots, after = before.filter((x) => x.id !== id);
    const ops: StyleOp[] = [];
    const a = p.area;
    if (a) for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) ops.push([m.styles, key(r, c), undefined]);
    const tx = new StyleTxn(ops);
    this.record(a ? () => this.hf.setCellContents(this.A(s, a.r1, a.c1), Array.from({ length: a.r2 - a.r1 + 1 }, () => Array(a.c2 - a.c1 + 1).fill(null))) : null, {
      undo: () => { tx.hist.undo(); this.meta(s).pivots = before; },
      redo: () => { tx.hist.redo(); this.meta(s).pivots = after; },
    }, { post: () => { tx.apply(); m.pivots = after; } });
  }
  private writePivot(s: number, p: PivotDef, isNew: boolean) {
    if (!this.sheetIds().includes(p.src.sheet)) throw new Error('Sheet nguồn của PivotTable không còn.');
    const res = computePivot(this.pivotTable(p.src), p);
    const rows = res.cells.length, cols = Math.max(...res.cells.map((r) => r.length));
    const area: Range = { r1: p.at.r, c1: p.at.c, r2: p.at.r + rows - 1, c2: p.at.c + cols - 1 };
    if (area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS) throw new Error('PivotTable vượt quá giới hạn trang tính.');
    if (s === p.src.sheet) {
      const q = p.src.range;
      if (area.r1 <= q.r2 && area.r2 >= q.r1 && area.c1 <= q.c2 && area.c2 >= q.c1) throw new Error('PivotTable không được đặt chồng lên vùng dữ liệu nguồn. Hãy chọn ô khác hoặc đặt ở sheet mới.');
    }
    const old = p.area;
    const inOld = (r: number, c: number) => !!old && r >= old.r1 && r <= old.r2 && c >= old.c1 && c <= old.c2;
    for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++)
      if (!inOld(r, c) && this.has(s, r, c)) throw new Error('Vùng đặt PivotTable đang có dữ liệu. Hãy chọn chỗ trống khác.');
    const m = this.meta(s);
    const ops: StyleOp[] = [];
    if (old) for (let r = old.r1; r <= old.r2; r++) for (let c = old.c1; c <= old.c2; c++) ops.push([m.styles, key(r, c), undefined]);
    res.cells.forEach((row, i) => row.forEach((v, j) => {
      const k = res.kinds[i][j];
      const st: Style = {};
      if (k === 'h') { st.b = true; st.bg = '#ddebf7'; }
      if (k === 't') st.b = true;
      if (k === 'g') { st.b = true; st.bg = '#ddebf7'; st.bd = { t: { s: 'thin' } }; }
      if (typeof v === 'number') st.fmt = '#,##0.##';
      ops.push([m.styles, key(area.r1 + i, area.c1 + j), Object.keys(st).length ? st : undefined]);
    }));
    const tx = new StyleTxn(ops);
    // nới rộng cột cho vừa nội dung (không thu hẹp), như Excel tự khớp độ rộng cột của Pivot
    const widen: [number, number, number][] = [];
    for (let j = 0; j < cols; j++) {
      let len = 0;
      for (const row of res.cells) {
        const v = row[j];
        const t = typeof v === 'number' ? Math.round(v).toLocaleString('vi-VN') + (Number.isInteger(v) ? '' : ',00') : v == null ? '' : String(v);
        len = Math.max(len, t.length);
      }
      const c = area.c1 + j, want = Math.min(320, Math.ceil(len * 7.6 + 16)), o = m.colW.base(c);
      if (want > o) widen.push([c, o, want]);
    }
    const np: PivotDef = { ...p, area };
    const before = m.pivots, after = isNew ? [...before, np] : before.map((x) => (x.id === p.id ? np : x));
    this.record(() => {
      if (old) this.hf.setCellContents(this.A(s, old.r1, old.c1), Array.from({ length: old.r2 - old.r1 + 1 }, () => Array(old.c2 - old.c1 + 1).fill(null)));
      this.hf.setCellContents(this.A(s, area.r1, area.c1), res.cells.map((row) => Array.from({ length: cols }, (_, j) => (row[j] ?? null) as RawCellContent)).map((row) => row.map((v) => (typeof v === 'string' && needsText(v) ? "'" + v : v))));
    }, {
      undo: () => { tx.hist.undo(); this.meta(s).pivots = before; for (const [c, o] of widen) this.meta(s).colW.set(c, o); },
      redo: () => { tx.hist.redo(); this.meta(s).pivots = after; for (const [c, , w] of widen) this.meta(s).colW.set(c, w); },
    }, { post: () => { tx.apply(); m.pivots = after; for (const [c, , w] of widen) m.colW.set(c, w); } });
    return res.warnings;
  }

  // ---------- thiết lập in ----------
  setPrint(s: number, ps: PrintSetup | null) {
    const m = this.meta(s), before = m.print, after = ps ? { ...ps } : null;
    m.print = after;
    this.record(null, { undo: () => (this.meta(s).print = before), redo: () => (this.meta(s).print = after) });
  }

  // ---------- CSV ----------
  importCSV(text: string, name?: string): number {
    const { rows, delim } = parseCSV(text);
    let n = name || 'CSV', i = 1;
    while (this.hasName(n)) n = (name || 'CSV') + ' (' + ++i + ')';
    const id = this.addSheet(n);
    if (!rows.length) return id;
    const w = Math.max(...rows.map((x) => x.length));
    if (rows.length > MAX_ROWS || w > MAX_COLS) throw new Error('Tệp CSV vượt quá giới hạn trang tính.');
    const st = this.meta(id).styles;
    const ops: StyleOp[] = [];
    const data = rows.map((row, r) => Array.from({ length: w }, (_, c) => {
      const t = row[c] ?? '';
      const tt = t.trim();
      if (delim !== ';') {
        // CSV quốc tế: "." là dấu thập phân, "," là nhóm nghìn
        if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(tt)) return +tt;
        if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(tt)) return +tt.replace(/,/g, '');
        const pm = /^(-?\d+(?:\.(\d+))?)\s*%$/.exec(tt);
        if (pm) { ops.push([st, key(r, c), { fmt: pm[2] ? '0.' + '0'.repeat(Math.min(pm[2].length, 4)) + '%' : '0%' }]); return +pm[1] / 100; }
        if (/^-?[\d.,]+$/.test(tt)) return "'" + t;
      }
      const nn = normalizeInput(t);
      if (nn.fmt) ops.push([st, key(r, c), { fmt: nn.fmt }]);
      return nn.content as RawCellContent;
    }));
    const tx = new StyleTxn(ops);
    this.record(() => this.hf.setCellContents(this.A(id, 0, 0), data), tx.hist, { post: () => tx.apply() });
    return id;
  }
  exportCSV(s: number, delim = ','): string {
    const u = this.used(s);
    const rows: string[][] = [];
    for (let r = 0; r < u.rows; r++) {
      const row: string[] = [];
      for (let c = 0; c < u.cols; c++) {
        const v = this.value(s, r, c);
        // CSV dấu phẩy: số ghi dạng quốc tế (dấu chấm thập phân, không làm tròn); ngày giữ như hiển thị
        if (delim === ',' && typeof v === 'number' && !isDateFmt(this.effFmt(s, r, c))) row.push(String(v));
        else row.push(this.display(s, r, c).text);
      }
      rows.push(row);
    }
    return toCSV(rows, delim);
  }

  // ---------- lưu / mở ----------
  toJSON() {
    const refOut = (r?: Ref) => (r && this.sheetIds().includes(r.sheet) ? { sheetName: this.sheetName(r.sheet), range: r.range } : undefined);
    const ext = this.allSheetIds().filter((id) => isExtName(this.sheetName(id)));
    const iterVals = [...this.over.entries()].map(([k, v]) => { const [s, r, c] = k.split(':').map(Number); return [this.sheetName(s), r, c, v] as [string, number, number, unknown]; })
      .filter(([, , , v]) => v === null || typeof v !== 'object');
    return {
      v: 1,
      name: this.name,
      calc: this.iterate.on ? { ...this.iterate } : undefined,
      iterVals: iterVals.length ? iterVals : undefined,
      names: (() => { const l = this.listNames().map((n) => ({ name: n.name, expr: n.expr, scope: n.scope !== undefined ? this.sheetName(n.scope) : undefined, comment: n.comment, hidden: n.hidden })); return l.length ? l : undefined; })(),
      ext: this.extLinks.length || ext.length ? { links: this.extLinks, sheets: ext.map((id) => ({ name: this.sheetName(id), cells: this.hf.getSheetSerialized(id) })) } : undefined,
      sheets: this.sheetIds().map((id) => {
        const m = this.meta(id);
        const qn = this.quoteNames();
        const cells = this.hf.getSheetSerialized(id);
        if (qn.length) for (const row of cells) for (let i = 0; i < row.length; i++) { const v = row[i]; if (typeof v === 'string' && v.startsWith('=')) row[i] = this.outF(v, qn); }
        return {
          name: this.sheetName(id),
          cells,
          styles: [...m.styles.entries()].map(([k, v]) => [Number(k), v]),
          colW: m.colW.entries(),
          rowH: m.rowH.entries(),
          freeze: m.freeze,
          colDef: m.colW.def,
          rowDef: m.rowH.def,
          hiddenRows: m.rowH.hidden(),
          hiddenCols: m.colW.hidden().length ? m.colW.hidden() : undefined,
          breaks: m.breaks.rows.length || m.breaks.cols.length ? m.breaks : undefined,
          outline: Object.keys(m.outline.r).length || Object.keys(m.outline.c).length ? m.outline : undefined,
          tables: m.tables.length ? m.tables : undefined,
          tabColor: m.tabColor,
          state: m.hidden ? 'hidden' : undefined,
          merges: m.merges,
          cf: m.cf,
          dv: m.dv,
          filter: m.filter,
          charts: m.charts.map((c) => ({ ...c, series: c.series.map((x) => ({ name: refOut(x.name), cats: refOut(x.cats), vals: refOut(x.vals)! })) })),
          pivots: m.pivots.map((p) => ({ ...p, src: refOut(p.src)! })),
          print: m.print,
        };
      }),
    };
  }
  static fromJSON(d: BookJSON): Workbook {
    const pre: TableReg[] = d.sheets.flatMap((sh, i) => (sh.tables ?? []).map((t) => ({ ...t, sheet: -1 - i, sheetName: sh.name })));
    const wb = new Workbook(buildHF([...d.sheets, ...(d.ext?.sheets ?? [])].map((x) => ({ name: x.name, cells: x.cells as RawCellContent[][] })), pre));
    wb.name = d.name || 'Sổ tính';
    wb.extLinks = d.ext?.links ?? [];
    const ids = d.sheets.map((x) => wb.hf.getSheetId(x.name)!);
    d.sheets.forEach((sh, i) => {
      const m = wb.meta(ids[i]);
      setAll(m.styles, new Map((sh.styles as [number, Style][]).map(([k, v]) => [+k, v])));
      if (sh.colDef) m.colW = new AxisSizes(sh.colDef, MAX_COLS);
      if (sh.rowDef) m.rowH = new AxisSizes(sh.rowDef, MAX_ROWS);
      m.colW.load((sh.colW ?? []) as [number, number][], sh.hiddenCols ?? []);
      m.breaks = { rows: sh.breaks?.rows ?? [], cols: sh.breaks?.cols ?? [] };
      if (sh.outline) m.outline = { r: { ...sh.outline.r }, c: { ...sh.outline.c }, x: [...(sh.outline.x ?? [])] };
      if (sh.tables?.length) { m.tables = sh.tables.map((t) => ({ ...t, cols: t.cols.map((c) => ({ ...c })) })); for (const t of m.tables) wb.tableSeq = Math.max(wb.tableSeq, t.id); }
      if (sh.tabColor) m.tabColor = sh.tabColor;
      if (sh.state === 'hidden') m.hidden = true;
      m.rowH.load((sh.rowH ?? []) as [number, number][], sh.hiddenRows ?? []);
      m.freeze = sh.freeze ?? { rows: 0, cols: 0 };
      m.merges = sh.merges ?? [];
      m.cf = (sh.cf ?? []).map((x) => ({ ...x, id: x.id ?? Math.random().toString(36).slice(2, 10) }));
      m.dv = (sh.dv ?? []).map((x) => ({ ...x, id: x.id ?? Math.random().toString(36).slice(2, 10) }));
      m.filter = sh.filter ?? null;
      m.print = sh.print ?? null;
    });
    // tham chiếu biểu đồ/Pivot theo tên sheet → mã sheet mới
    const refIn = (r: any): Ref | undefined => {
      if (!r) return undefined;
      if (typeof r.sheet === 'number' && r.sheetName === undefined) return r;
      const id = wb.hf.getSheetId(r.sheetName);
      return id === undefined ? undefined : { sheet: id, range: r.range };
    };
    d.sheets.forEach((sh, i) => {
      const m = wb.meta(ids[i]);
      m.charts = (sh.charts ?? []).map((c: any) => ({ ...c, id: c.id ?? Math.random().toString(36).slice(2, 10), series: (c.series ?? []).map((x: any) => ({ name: refIn(x.name), cats: refIn(x.cats), vals: refIn(x.vals) })).filter((x: any) => x.vals) }));
      m.pivots = (sh.pivots ?? []).map((p: any) => ({ ...p, src: refIn(p.src) })).filter((p: any) => p.src);
    });
    if (d.names?.length) {
      const disp = nameDisplay(wb.hf);
      wb.hf.suspendEvaluation();
      for (const n of d.names) {
        const scope = n.scope !== undefined ? wb.hf.getSheetId(n.scope) : undefined;
        if (n.scope !== undefined && scope === undefined) continue;
        const enc = encName(n.name);
        const fx = encFormulaNames(offsetIn(n.expr[0] === '=' ? n.expr : '=' + n.expr, undefined, ''));
        const opts: any = {}; if (n.comment) opts.comment = n.comment; if (n.hidden) opts.hidden = true;
        try { if (wb.hf.isItPossibleToAddNamedExpression(enc, fx, scope)) { wb.hf.addNamedExpression(enc, fx, scope, Object.keys(opts).length ? opts : undefined); disp.set(enc.toUpperCase(), n.name); } } catch { /* tên lỗi: bỏ qua */ }
      }
      wb.hf.resumeEvaluation();
    }
    wb.hf.clearUndoStack();
    if (d.calc) wb.iterate = { on: !!d.calc.on, count: d.calc.count ?? 100, delta: d.calc.delta ?? 0.001 };
    for (const [n, r, c, v] of d.iterVals ?? []) { const id = wb.hf.getSheetId(n); if (id !== undefined) wb.seeds.set(id + ':' + r + ':' + c, v); }
    wb.resolveIter();
    return wb;
  }
  /** Nạp nhanh một bảng lớn (dùng cho nhập dữ liệu, đo hiệu năng) – không ghi lịch sử */
  loadSheet(s: number, rows: RawCellContent[][]) {
    this.hf.setSheetContent(s, rows);
    this.hf.clearUndoStack();
    this.past = []; this.future = [];
    this.emit();
  }
}

function colLetters(c: number) {
  let s = ''; c++;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
export interface SheetJSON {
  name: string; cells: unknown[][]; styles: unknown[]; colW?: [number, number][]; rowH?: [number, number][];
  freeze?: { rows: number; cols: number }; colDef?: number; rowDef?: number; hiddenRows?: number[];
  hiddenCols?: number[]; breaks?: { rows: number[]; cols: number[] }; tabColor?: string; state?: 'hidden';
  merges?: Range[]; cf?: CFRule[]; dv?: DVRule[]; filter?: SheetFilter | null;
  charts?: unknown[]; pivots?: unknown[]; print?: PrintSetup | null;
  outline?: Outline;
  tables?: TableDef[];
  /** biểu đồ đọc từ .xlsx, gắn sau khi sổ đã dựng (finishImport) */
  rawCharts?: unknown[];
}
export interface ExtLink { index: number; path: string; xml: string; rels?: string }
/** một bình luận trong chuỗi chú thích */
export interface CommentItem { a: string; t: string; d: string }
export interface NameDef { name: string; expr: string; scope?: number; comment?: string; hidden?: boolean }
const BUILTIN_NAMES = new Set(['TRUE', 'FALSE']);
export interface BookJSON {
  v: number; name: string; sheets: SheetJSON[];
  /** tên vùng: scope là tên trang tính (không có = cả sổ) */
  names?: { name: string; expr: string; scope?: string; comment?: string; hidden?: boolean }[];
  calc?: { on: boolean; count?: number; delta?: number };
  iterVals?: [string, number, number, unknown][];
  ext?: { links: ExtLink[]; sheets: { name: string; cells: unknown[][] }[] };
}
/** Sheet đệm giá trị của liên kết ngoài: tên dạng “[1]Tên sheet” (tên sheet thường không được có dấu [ ]) */
export const isExtName = (n: string) => /^\[\d+\]/.test(n);

// ---- tách tham chiếu ô/vùng trong công thức (dùng cho tính lặp) ----
export interface XRef { text: string; s: number; r1: number; c1: number; r2: number; c2: number }
const colNum = (L: string) => { let n = 0; for (const ch of L.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const REF_RE = /((?:'(?:[^']|'')+'|[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*)!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?(?![\w(!])/y;
export function compileRefs(f: string, sheetOf: (name: string | undefined) => number | undefined): (string | XRef)[] {
  const out: (string | XRef)[] = [];
  let buf = '', i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') { let j = i + 1; while (j < f.length) { if (f[j] === '"' && f[j + 1] === '"') { j += 2; continue; } if (f[j] === '"') break; j++; } buf += f.slice(i, j + 1); i = j + 1; continue; }
    const prev = i > 0 ? f[i - 1] : '';
    if (!/[\w.\u00C0-\uFFFF$']/.test(prev)) {
      REF_RE.lastIndex = i;
      const m = REF_RE.exec(f);
      if (m) {
        const name = m[1] ? m[1].slice(0, -1).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : undefined;
        const s = sheetOf(name);
        if (s !== undefined) {
          const r1 = +m[3] - 1, c1 = colNum(m[2]), r2 = m[5] ? +m[5] - 1 : r1, c2 = m[4] ? colNum(m[4]) : c1;
          if (buf) { out.push(buf); buf = ''; }
          out.push({ text: m[0], s, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) });
          i += m[0].length; continue;
        }
      }
    }
    buf += ch; i++;
  }
  if (buf) out.push(buf);
  return out;
}
/** Giá trị → hằng trong công thức (dấu thập phân “,” như cấu hình) */
function lit(v: unknown, inArray = false): string {
  if (v === null || v === undefined || v === '') return '0';
  if (typeof v === 'number') {
    if (!isFinite(v)) return '0';
    return v.toLocaleString('en-US', { useGrouping: false, maximumSignificantDigits: 15 }).replace('.', ',');
  }
  if (typeof v === 'boolean') return inArray ? (v ? 'TRUE' : 'FALSE') : v ? 'TRUE()' : 'FALSE()';
  if (typeof v === 'string') return '"' + v.replace(/"/g, '""') + '"';
  const e = (v as any)?.value;
  return typeof e === 'string' && e.startsWith('#') && !inArray ? e : '0';
}

function newMeta(): SheetMeta {
  return { styles: new Map(), colW: new AxisSizes(DEFAULT_COL_W, MAX_COLS), rowH: new AxisSizes(DEFAULT_ROW_H, MAX_ROWS), freeze: { rows: 0, cols: 0 }, merges: [], cf: [], dv: [], filter: null, charts: [], pivots: [], print: null, breaks: { rows: [], cols: [] }, outline: NO_OUTLINE, tables: [] };
}
function clean(st: Style): Style | undefined {
  const o: any = {};
  for (const [k, v] of Object.entries(st)) if (v !== undefined && (v !== false || k === 'lk') && v !== '' && !(k === 'bd' && (!v || !Object.keys(v).length))) o[k] = v;
  return Object.keys(o).length ? o : undefined;
}
/** Ngày dd/mm/yyyy → số sê-ri Excel */
export function dateSerial(t: string): number | null {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(t.trim());
  if (!m) return null;
  let y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900;
  const ms = Date.UTC(y, +m[2] - 1, +m[1]);
  return Math.round((ms - Date.UTC(1899, 11, 30)) / 86400000);
}
/** Chỉ số mới cho một vị trí nằm trong phần bị xóa: dời về đầu phần xóa */
function delSpanStart(v: number, span: (a: number, b: number) => { a: number; b: number } | null) {
  for (let x = v; x >= 0; x--) { const r = span(x, x); if (r) return x === v ? r.a : r.a + 1; }
  return 0;
}
/** Giá trị → nội dung ô (chữ trông như số/công thức thì thêm ' để giữ là chữ) */
function valueContent(v: unknown): RawCellContent {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (typeof v === 'string') return needsText(v) ? "'" + v : v;
  const e = (v as any)?.value;
  return typeof e === 'string' ? e : null;
}
const needsText = (v: string) => /^[=+\-'\d.,]/.test(v) || /^(true|false)$/i.test(v);
/** Dời các vùng (gộp ô, CF, DV, lọc) khi chèn/xóa hàng hoặc cột */
function delSpan(a: number, b: number, at: number, n: number) {
  const na = a < at ? a : a >= at + n ? a - n : at;
  const nb = b < at ? b : b >= at + n ? b - n : at - 1;
  return nb < na ? null : { a: na, b: nb };
}
function shiftAll(m: SheetMeta, ax: 'r' | 'c', pt: (x: number) => number, span: (a: number, b: number) => { a: number; b: number } | null) {
  const mv = (g: Range): Range | null => {
    const x = ax === 'r' ? span(g.r1, g.r2) : span(g.c1, g.c2);
    if (!x) return null;
    return ax === 'r' ? { ...g, r1: x.a, r2: x.b } : { ...g, c1: x.a, c2: x.b };
  };
  m.merges = m.merges.map(mv).filter((g): g is Range => !!g && !(g.r1 === g.r2 && g.c1 === g.c2));
  m.cf = m.cf.map((x) => { const g = mv(x.range); return g ? { ...x, range: g } : null; }).filter(Boolean) as CFRule[];
  m.dv = m.dv.map((x) => { const g = mv(x.range); return g ? { ...x, range: g } : null; }).filter(Boolean) as DVRule[];
  m.charts = m.charts.map((ch) => {
    const v = ax === 'r' ? ch.anchor.r : ch.anchor.c;
    const nv = pt(v) < 0 ? (span(v, v)?.a ?? Math.max(0, v)) : pt(v);
    const fix = nv < 0 ? delSpanStart(v, span) : nv;
    return fix === v ? ch : { ...ch, anchor: { ...ch.anchor, [ax]: fix } };
  });
  m.pivots = m.pivots.map((p) => {
    const at = ax === 'r' ? p.at.r : p.at.c;
    const nat = pt(at) >= 0 ? pt(at) : delSpanStart(at, span);
    return { ...p, at: ax === 'r' ? { ...p.at, r: nat } : { ...p.at, c: nat }, area: p.area ? mv(p.area) : null };
  }).filter((p, i) => !(m.pivots[i].area && !p.area)); // xóa hết vùng Pivot → bỏ luôn PivotTable
  {
    const k = ax === 'r' ? 'rows' : 'cols';
    m.breaks = { ...m.breaks, [k]: [...new Set(m.breaks[k].map(pt).filter((x) => x > 0))].sort((a, b) => a - b) };
  }
  // hàng/cột lặp lại khi in dời theo
  if (m.print) {
    const key2 = ax === 'r' ? 'titleRowRange' : 'titleColRange';
    const t = m.print[key2];
    if (t) { const x = span(t[0], t[1]); m.print = { ...m.print, [key2]: x ? [x.a, x.b] : null }; }
  }
  if (m.print?.area) {
    const g = mv(m.print.area);
    const pr = { ...m.print };
    if (g) pr.area = g; else delete pr.area;
    m.print = pr;
  }
  {
    const key3 = ax === 'r' ? 'r' : 'c';
    const lv = m.outline[key3], nl: Record<number, number> = {};
    for (const [k, v] of Object.entries(lv)) { const nk = pt(+k); if (nk >= 0) nl[nk] = v; }
    const nx = m.outline.x.map((kx) => { const [t, L, at] = kx.split(':'); if (t !== key3) return kx; const n = pt(+at); return n >= 0 ? `${t}:${L}:${n}` : ''; }).filter(Boolean);
    m.outline = { ...m.outline, [key3]: nl, x: nx };
  }
  if (m.tables.length) {
    m.tables = m.tables.flatMap((t) => {
      const g = mv({ r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 });
      if (!g) return [];
      if (ax === 'r') {
        if (t.header && pt(t.r1) < 0) return []; // xóa hàng tiêu đề → bỏ bảng
        if (g.r2 - g.r1 + 1 < (t.header ? 1 : 0) + (t.totals ? 1 : 0) + 1) return [];
        return [{ ...t, ...g }];
      }
      // cột: giữ cột còn lại theo vị trí mới, cột chèn vào giữa đặt tên Cột<n>
      const keep = new Map<number, TableDef['cols'][number]>();
      t.cols.forEach((col, i) => { const n = pt(t.c1 + i); if (n >= 0) keep.set(n, col); });
      let nid = Math.max(0, ...t.cols.map((x) => x.id)) + 1;
      const used = new Set(t.cols.map((x) => x.name.toLowerCase()));
      const cols: TableDef['cols'] = [];
      for (let c = g.c1; c <= g.c2; c++) {
        const k = keep.get(c);
        if (k) { cols.push(k); continue; }
        let j: number = cols.length + 1, nm: string = `Cột${j}`; while (used.has(nm.toLowerCase())) nm = `Cột${++j}`;
        used.add(nm.toLowerCase()); cols.push({ id: nid++, name: nm });
      }
      return cols.length ? [{ ...t, ...g, cols }] : [];
    });
  }
  if (m.filter) {
    const g = mv(m.filter.range);
    if (!g) m.filter = null;
    else {
      const conds: Record<number, FilterCond> = {};
      for (const [c, v] of Object.entries(m.filter.conds)) { const nc = ax === 'c' ? pt(+c) : +c; if (nc >= 0) conds[nc] = v; }
      m.filter = { range: g, conds };
    }
  }
}
function setOrDel(m: Map<number, Style>, k: number, v: Style | undefined) { if (v) m.set(k, { ...v }); else m.delete(k); }
function remapStyles(m: Map<number, Style>, f: (r: number, c: number) => { r: number; c: number } | null) {
  const out = new Map<number, Style>();
  for (const [k, v] of m) { const p = unkey(+k); const q = f(p.r, p.c); if (q) out.set(key(q.r, q.c), v); }
  return out;
}
// Bảng định dạng của mỗi sheet luôn là MỘT đối tượng Map cố định (chỉ sửa tại chỗ),
// để các mục hoàn tác giữ tham chiếu tới nó không bị "lạc" sang bản cũ.
function setAll(m: Map<number, Style>, src: Map<number, Style>) { if (m === src) return; m.clear(); src.forEach((v, k) => m.set(k, v)); }
type MetaSnap = { tables?: TableDef[]; outline?: Outline; styles: Map<number, Style>; colW: [number, number][]; rowH: [number, number][]; hidden: number[]; colHidden: number[]; freeze: { rows: number; cols: number }; merges: Range[]; cf: CFRule[]; dv: DVRule[]; filter: SheetFilter | null; charts: Chart[]; pivots: PivotDef[]; print: PrintSetup | null; breaks: { rows: number[]; cols: number[] }; tabColor?: string; sheetHidden?: boolean };
function snapMeta(m: SheetMeta): MetaSnap {
  return { styles: new Map(m.styles), colW: m.colW.entries(), rowH: m.rowH.entries(), hidden: m.rowH.hidden(), colHidden: m.colW.hidden(), freeze: { ...m.freeze }, merges: m.merges, cf: m.cf, dv: m.dv, filter: m.filter ? JSON.parse(JSON.stringify(m.filter)) : null, charts: m.charts, pivots: m.pivots, print: m.print, breaks: { rows: [...m.breaks.rows], cols: [...m.breaks.cols] }, tabColor: m.tabColor, sheetHidden: m.hidden, outline: m.outline, tables: m.tables };
}
function restoreMeta(m: SheetMeta, s: MetaSnap) {
  setAll(m.styles, new Map(s.styles)); m.colW.load(s.colW, s.colHidden ?? []); m.rowH.load(s.rowH, s.hidden); m.freeze = { ...s.freeze };
  m.merges = s.merges; m.cf = s.cf; m.dv = s.dv; m.filter = s.filter ? JSON.parse(JSON.stringify(s.filter)) : null;
  m.charts = s.charts ?? []; m.pivots = s.pivots ?? []; m.print = s.print ?? null; m.outline = s.outline ?? NO_OUTLINE; m.tables = s.tables ?? [];
  m.breaks = { rows: [...(s.breaks?.rows ?? [])], cols: [...(s.breaks?.cols ?? [])] }; m.tabColor = s.tabColor; m.hidden = s.sheetHidden;
}

/** Một nhóm thay đổi định dạng: chỉ lưu các ô bị đổi (trước/sau), không chép cả bảng */
type StyleOp = [Map<number, Style>, number, Style | undefined];
class StyleTxn {
  private before: StyleOp[] = [];
  private after: StyleOp[] = [];
  constructor(private ops: StyleOp[]) {}
  apply() {
    const seen = new Map<Map<number, Style>, Set<number>>();
    for (const [m, k, v] of this.ops) {
      let ks = seen.get(m);
      if (!ks) { ks = new Set(); seen.set(m, ks); }
      if (!ks.has(k)) { ks.add(k); const o = m.get(k); this.before.push([m, k, o ? { ...o } : undefined]); }
      setOrDel(m, k, v);
    }
    for (const [m, k] of this.before) { const o = m.get(k); this.after.push([m, k, o ? { ...o } : undefined]); }
  }
  get hist() {
    return {
      undo: () => { for (const [m, k, v] of this.before) setOrDel(m, k, v); },
      redo: () => { for (const [m, k, v] of this.after) setOrDel(m, k, v); },
    };
  }
}
