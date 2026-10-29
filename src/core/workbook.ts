import { DynamicOffsetPlugin, offsetIn, offsetOut } from './offset';
import { HyperFormula, FunctionPlugin, FunctionArgumentType, CellError, ErrorType } from 'hyperformula';
import { shiftFormula, fixSheetQuotes, needsQuote } from './formula';
import type { ConfigParams, RawCellContent, SimpleCellAddress } from 'hyperformula';
import { MAX_COLS, MAX_ROWS, type Range, rangeSize } from './address';
import { normalizeInput } from './input';
import { editString, formatValue, isDateFmt, type Align } from './numfmt';
import { AxisSizes } from './sizes';
import { fillPlan, type SrcCell } from './series';
import { toTSV, parseTSV } from './tsv';
import { parseCSV, toCSV } from './csv';
import { parseRef } from './address';
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
  bd?: { t?: Border; r?: Border; b?: Border; l?: Border };
  ind?: number;                       // thụt lề (bậc)
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
  styles: Map<string, Style>;
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
}
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
  arrayColumnSeparator: ',',
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
  HyperFormula.registerFunctionPlugin(DynamicOffsetPlugin as any, { enGB: { XOFS3: 'XOFS3', XOFS4: 'XOFS4', XOFS5: 'XOFS5' } } as any);
  pluginsReady = true;
}
/** OFFSET động: đổi công thức khi vào bộ tính và đổi ngược khi đọc ra (xem offset.ts) */
function mapRows(cells: RawCellContent[][], sheet: string, r0 = 0, c0 = 0): RawCellContent[][] {
  let hit = false;
  for (const row of cells) { if (row) for (const v of row) if (typeof v === 'string' && v[0] === '=' && /OFFSET\s*\(/i.test(v)) { hit = true; break; } if (hit) break; }
  if (!hit) return cells;
  return cells.map((row, i) => row && row.map((v, j) => (typeof v === 'string' ? (offsetIn(v, { r: r0 + i, c: c0 + j }, sheet) as RawCellContent) : v)));
}
function patchOffset(hf: HyperFormula): HyperFormula {
  const h: any = hf;
  const set = h.setCellContents.bind(hf), sheetSet = h.setSheetContent.bind(hf);
  const gf = h.getCellFormula.bind(hf), gs = h.getCellSerialized.bind(hf), gss = h.getSheetSerialized.bind(hf);
  h.setCellContents = (a: any, c: any) => {
    const name = hf.getSheetName(a.sheet) ?? '';
    if (Array.isArray(c)) return set(a, mapRows(c, name, a.row, a.col));
    return set(a, typeof c === 'string' ? offsetIn(c, { r: a.row, c: a.col }, name) : c);
  };
  const calc = h.calculateFormula.bind(hf);
  h.calculateFormula = (f: string, sheet: number) => calc(offsetIn(f, undefined, hf.getSheetName(sheet) ?? ''), sheet);
  h.setSheetContent = (id: number, cells: RawCellContent[][]) => sheetSet(id, mapRows(cells, hf.getSheetName(id) ?? ''));
  h.getCellFormula = (a: any) => { const f = gf(a); return f === undefined ? f : offsetOut(f); };
  h.getCellSerialized = (a: any) => { const v = gs(a); return typeof v === 'string' ? offsetOut(v) : v; };
  h.getSheetSerialized = (id: number) => gss(id).map((row: any[]) => row.map((v) => (typeof v === 'string' ? offsetOut(v) : v)));
  return hf;
}
// Excel cho phép viết TRUE/FALSE không có ngoặc
const NAMES = [{ name: 'TRUE', expression: '=TRUE()' }, { name: 'FALSE', expression: '=FALSE()' }];
function buildHF(sheets?: { name: string; cells: RawCellContent[][] }[]) {
  ensurePlugins();
  // buildFromSheets nhanh hơn ~2,5 lần khi mở sổ lớn; thứ tự key của object phải giữ đúng thứ tự sheet
  if (sheets && sheets.length && !sheets.some((x) => /^\d+$/.test(x.name))) {
    const obj: Record<string, RawCellContent[][]> = Object.create(null);
    for (const x of sheets) obj[x.name] = mapRows(x.cells, x.name);
    return patchOffset(HyperFormula.buildFromSheets(obj, HF_CONFIG, NAMES));
  }
  const hf = patchOffset(HyperFormula.buildEmpty(HF_CONFIG));
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
  hf: HyperFormula;
  private metas = new Map<number, SheetMeta>();
  private past: Entry[] = [];
  private future: Entry[] = [];
  private listeners = new Set<() => void>();
  clip: Clip | null = null;
  name = 'Sổ tính 1';

  private constructor(hf: HyperFormula) { this.hf = hf; }

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
    this.record(() => { this.hf.addSheet(n); nid = this.hf.getSheetId(n)!; this.hf.setSheetContent(nid, cells); }, {
      undo: () => { this.order = oOrder; },
      redo: () => { if (fresh) restoreMeta(this.meta(nid), fresh); this.order = nOrder; },
    }, {
      post: () => {
        const m = this.meta(nid);
        const sm = this.meta(id);
        m.colW = new AxisSizes(sm.colW.def, MAX_COLS); m.rowH = new AxisSizes(sm.rowH.def, MAX_ROWS);
        restoreMeta(m, { ...src, styles: new Map([...src.styles].map(([k2, v]) => [k2, { ...v }])), merges: src.merges.map((g) => ({ ...g })), charts: [], pivots: [] });
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
  style(s: number, r: number, c: number): Style { const m = this.meta(s).styles; return m.size ? m.get(String(key(r, c))) ?? NO_STYLE : NO_STYLE; }
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
    const k = String(key(r, c));
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
    const before = new Map<string, Style | undefined>();
    const after = new Map<string, Style | undefined>();
    for (let r = g.r1; r <= g.r2; r++)
      for (let c = g.c1; c <= g.c2; c++) {
        const k = String(key(r, c));
        const o = st.get(k);
        before.set(k, o);
        after.set(k, clean({ ...(o ?? {}), ...patch }));
      }
    const apply = (m: Map<string, Style | undefined>) => m.forEach((v, k) => setOrDel(st, k, v));
    apply(after);
    this.record(null, { undo: () => apply(before), redo: () => apply(after) });
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
    this.record(hfOp, {
      undo: () => { restoreMeta(m, before); this.restoreRefs(xBefore); },
      redo: () => { restoreMeta(m, after); this.restoreRefs(xAfter); },
    }, {
      post: () => { shift(m); this.shiftRefs(s, ax, span); after = snapMeta(m); xAfter = this.snapRefs(); },
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
    if (cl.cut) for (let i = 0; i <= h; i++) for (let j = 0; j <= w; j++) ops.push([srcSt, String(key(src.r1 + i, src.c1 + j)), undefined]);
    for (let i = 0; i <= h; i++) for (let j = 0; j <= w; j++) ops.push([dstSt, String(key(r + i, c + j)), srcSt.get(String(key(src.r1 + i, src.c1 + j)))]);
    const tx = new StyleTxn(ops);
    const range = { sheet: cl.sheet, start: { sheet: cl.sheet, row: src.r1, col: src.c1 }, end: { sheet: cl.sheet, row: src.r2, col: src.c2 } };
    const target = this.A(s, r, c);
    this.record(
      () => {
        if (cl.cut) this.hf.cut(range as any); else this.hf.copy(range as any);
        return this.hf.paste(target);
      },
      tx.hist,
      { batch: false, post: () => tx.apply() },
    );
    if (cl.cut) this.clip = null;
    return { r1: r, c1: c, r2: r + h, c2: c + w };
  }
  pasteText(s: number, r: number, c: number, text: string): Range | null {
    const rows = parseTSV(text);
    if (!rows.length) return null;
    const w = Math.max(...rows.map((x) => x.length));
    if (r + rows.length > MAX_ROWS || c + w > MAX_COLS) throw new Error('Vùng dán vượt quá giới hạn trang tính.');
    const fmts: [string, string][] = [];
    const data = rows.map((row, i) => Array.from({ length: w }, (_, j) => {
      const n = normalizeInput(row[j] ?? '');
      if (n.fmt) fmts.push([String(key(r + i, c + j)), n.fmt]);
      return n.content as RawCellContent;
    }));
    const st = this.meta(s).styles;
    const tx = fmts.length ? new StyleTxn(fmts.map(([k, f]) => [st, k, { ...(st.get(k) ?? {}), fmt: f }])) : null;
    this.record(() => this.hf.setCellContents(this.A(s, r, c), data), tx?.hist, { post: () => tx?.apply() });
    return { r1: r, c1: c, r2: r + rows.length - 1, c2: c + w - 1 };
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
        sops.push([m.styles, String(key(t.r, t.c)), m.styles.get(String(key(from.r, from.c)))]);
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
    const st = m.styles, k0 = String(key(g.r1, g.c1));
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
      const k = String(key(r, c));
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
        styles.push(st.get(String(key(r, c))));
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
    rows.forEach((row, j) => row.styles.forEach((v, i) => ops.push([st, String(key(slots[j], g.c1 + i)), v])));
    const tx = new StyleTxn(ops);
    const contiguous = slots.length === g.r2 - d1 + 1;
    this.record(() => {
      if (contiguous) this.hf.setCellContents(this.A(s, d1, g.c1), data);
      else data.forEach((row, j) => this.hf.setCellContents(this.A(s, slots[j], g.c1), [row]));
    }, tx.hist, { post: () => tx.apply() });
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
  filterValues(s: number, col: number): string[] {
    const f = this.meta(s).filter;
    if (!f) return [];
    const set = new Set<string>();
    const r2 = Math.min(f.range.r2, this.used(s).rows - 1);
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
    const u = this.used(s);
    let r2 = f.range.r2;
    while (r2 + 1 < u.rows && [...Array(f.range.c2 - f.range.c1 + 1).keys()].some((i) => this.has(s, r2 + 1, f.range.c1 + i))) r2++;
    f.range = { ...f.range, r2 };
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
    if (a) for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) ops.push([m.styles, String(key(r, c)), undefined]);
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
    if (old) for (let r = old.r1; r <= old.r2; r++) for (let c = old.c1; c <= old.c2; c++) ops.push([m.styles, String(key(r, c)), undefined]);
    res.cells.forEach((row, i) => row.forEach((v, j) => {
      const k = res.kinds[i][j];
      const st: Style = {};
      if (k === 'h') { st.b = true; st.bg = '#ddebf7'; }
      if (k === 't') st.b = true;
      if (k === 'g') { st.b = true; st.bg = '#ddebf7'; st.bd = { t: { s: 'thin' } }; }
      if (typeof v === 'number') st.fmt = '#,##0.##';
      ops.push([m.styles, String(key(area.r1 + i, area.c1 + j)), Object.keys(st).length ? st : undefined]);
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
        if (pm) { ops.push([st, String(key(r, c)), { fmt: pm[2] ? '0.' + '0'.repeat(Math.min(pm[2].length, 4)) + '%' : '0%' }]); return +pm[1] / 100; }
        if (/^-?[\d.,]+$/.test(tt)) return "'" + t;
      }
      const nn = normalizeInput(t);
      if (nn.fmt) ops.push([st, String(key(r, c)), { fmt: nn.fmt }]);
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
    const wb = new Workbook(buildHF([...d.sheets, ...(d.ext?.sheets ?? [])].map((x) => ({ name: x.name, cells: x.cells as RawCellContent[][] }))));
    wb.name = d.name || 'Sổ tính';
    wb.extLinks = d.ext?.links ?? [];
    const ids = d.sheets.map((x) => wb.hf.getSheetId(x.name)!);
    d.sheets.forEach((sh, i) => {
      const m = wb.meta(ids[i]);
      setAll(m.styles, new Map((sh.styles as [number, Style][]).map(([k, v]) => [String(k), v])));
      if (sh.colDef) m.colW = new AxisSizes(sh.colDef, MAX_COLS);
      if (sh.rowDef) m.rowH = new AxisSizes(sh.rowDef, MAX_ROWS);
      m.colW.load((sh.colW ?? []) as [number, number][], sh.hiddenCols ?? []);
      m.breaks = { rows: sh.breaks?.rows ?? [], cols: sh.breaks?.cols ?? [] };
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
}
export interface ExtLink { index: number; path: string; xml: string; rels?: string }
export interface BookJSON {
  v: number; name: string; sheets: SheetJSON[];
  calc?: { on: boolean; count?: number; delta?: number };
  iterVals?: [string, number, number, unknown][];
  ext?: { links: ExtLink[]; sheets: { name: string; cells: unknown[][] }[] };
}
/** Sheet đệm giá trị của liên kết ngoài: tên dạng “[1]Tên sheet” (tên sheet thường không được có dấu [ ]) */
export const isExtName = (n: string) => /^\[\d+\]/.test(n);

// ---- tách tham chiếu ô/vùng trong công thức (dùng cho tính lặp) ----
interface XRef { text: string; s: number; r1: number; c1: number; r2: number; c2: number }
const colNum = (L: string) => { let n = 0; for (const ch of L.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const REF_RE = /((?:'(?:[^']|'')+'|[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*)!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?(?![\w(!])/y;
function compileRefs(f: string, sheetOf: (name: string | undefined) => number | undefined): (string | XRef)[] {
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
  return { styles: new Map(), colW: new AxisSizes(DEFAULT_COL_W, MAX_COLS), rowH: new AxisSizes(DEFAULT_ROW_H, MAX_ROWS), freeze: { rows: 0, cols: 0 }, merges: [], cf: [], dv: [], filter: null, charts: [], pivots: [], print: null, breaks: { rows: [], cols: [] } };
}
function clean(st: Style): Style | undefined {
  const o: any = {};
  for (const [k, v] of Object.entries(st)) if (v !== undefined && v !== false && v !== '' && !(k === 'bd' && (!v || !Object.keys(v).length))) o[k] = v;
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
function setOrDel(m: Map<string, Style>, k: string, v: Style | undefined) { if (v) m.set(k, { ...v }); else m.delete(k); }
function remapStyles(m: Map<string, Style>, f: (r: number, c: number) => { r: number; c: number } | null) {
  const out = new Map<string, Style>();
  for (const [k, v] of m) { const p = unkey(+k); const q = f(p.r, p.c); if (q) out.set(String(key(q.r, q.c)), v); }
  return out;
}
// Bảng định dạng của mỗi sheet luôn là MỘT đối tượng Map cố định (chỉ sửa tại chỗ),
// để các mục hoàn tác giữ tham chiếu tới nó không bị "lạc" sang bản cũ.
function setAll(m: Map<string, Style>, src: Map<string, Style>) { if (m === src) return; m.clear(); src.forEach((v, k) => m.set(k, v)); }
type MetaSnap = { styles: Map<string, Style>; colW: [number, number][]; rowH: [number, number][]; hidden: number[]; colHidden: number[]; freeze: { rows: number; cols: number }; merges: Range[]; cf: CFRule[]; dv: DVRule[]; filter: SheetFilter | null; charts: Chart[]; pivots: PivotDef[]; print: PrintSetup | null; breaks: { rows: number[]; cols: number[] }; tabColor?: string; sheetHidden?: boolean };
function snapMeta(m: SheetMeta): MetaSnap {
  return { styles: new Map(m.styles), colW: m.colW.entries(), rowH: m.rowH.entries(), hidden: m.rowH.hidden(), colHidden: m.colW.hidden(), freeze: { ...m.freeze }, merges: m.merges, cf: m.cf, dv: m.dv, filter: m.filter ? JSON.parse(JSON.stringify(m.filter)) : null, charts: m.charts, pivots: m.pivots, print: m.print, breaks: { rows: [...m.breaks.rows], cols: [...m.breaks.cols] }, tabColor: m.tabColor, sheetHidden: m.hidden };
}
function restoreMeta(m: SheetMeta, s: MetaSnap) {
  setAll(m.styles, new Map(s.styles)); m.colW.load(s.colW, s.colHidden ?? []); m.rowH.load(s.rowH, s.hidden); m.freeze = { ...s.freeze };
  m.merges = s.merges; m.cf = s.cf; m.dv = s.dv; m.filter = s.filter ? JSON.parse(JSON.stringify(s.filter)) : null;
  m.charts = s.charts ?? []; m.pivots = s.pivots ?? []; m.print = s.print ?? null;
  m.breaks = { rows: [...(s.breaks?.rows ?? [])], cols: [...(s.breaks?.cols ?? [])] }; m.tabColor = s.tabColor; m.hidden = s.sheetHidden;
}

/** Một nhóm thay đổi định dạng: chỉ lưu các ô bị đổi (trước/sau), không chép cả bảng */
type StyleOp = [Map<string, Style>, string, Style | undefined];
class StyleTxn {
  private before: StyleOp[] = [];
  private after: StyleOp[] = [];
  constructor(private ops: StyleOp[]) {}
  apply() {
    const seen = new Map<Map<string, Style>, Set<string>>();
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
