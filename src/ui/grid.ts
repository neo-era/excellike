import { MAX_COLS, MAX_ROWS, type Range, colName, normRange, toA1, rangeToA1, inRange } from '../core/address';
import type { Workbook, Style, Border } from '../core/workbook';
import { normPrint, paginate } from '../core/paginate';

/** ar/ac: ô hiện hành; fr/fc: góc đối diện; box: khung vùng chọn khi ô hiện hành không ở góc (Enter trong vùng, Ctrl+Space…) */
export interface Selection { ar: number; ac: number; fr: number; fc: number; box?: Range }
export interface EditHost {
  isEditing(): boolean;
  pointMode(): boolean;
  insertRef(ref: string, replaceLast: boolean): void;
  begin(initial: string | null, mode: 'enter' | 'edit'): void;
  commit(dr: number, dc: number): boolean;
}

type Drag =
  | { kind: 'sel' }
  | { kind: 'col'; c0: number }
  | { kind: 'row'; r0: number }
  | { kind: 'resizeC'; c: number; x0: number; w0: number }
  | { kind: 'resizeR'; r: number; y0: number; h0: number }
  | { kind: 'fill'; src: Range; tgt: Range }
  | { kind: 'point'; r0: number; c0: number };

// phông ô và phông tiêu đề lấy từ CSS (--font-cell: Aptos Narrow/Calibri; --font-ui: Segoe UI)
let FONT = "'Aptos Narrow', Calibri, 'Segoe UI', sans-serif";
let UIFONT = "'Segoe UI', system-ui, sans-serif";
const HDR_H = 20;
/** cỡ chữ pt → px (96 dpi) như Excel: 11 pt = 14,67 px */
export const ptPx = (fs?: number) => ((fs ?? 11) * 4) / 3;
const ROW_STEP = 200, COL_STEP = 30;
const MAX_SCROLL_PX = 15_000_000;

export class Grid {
  sheet: number;
  sel: Selection = { ar: 0, ac: 0, fr: 0, fc: 0 };
  /** các vùng đã chọn thêm (Ctrl+nhấp, Shift+F8, Đi tới Đặc biệt) — không gồm vùng hiện hành */
  areas: Range[] = [];
  /** F8: mở rộng vùng chọn; Shift+F8: thêm vùng */
  extMode = false;
  addMode = false;
  /** Ctrl+6: 0 hiện, 1 chỗ giữ, 2 ẩn đối tượng */
  objMode = 0;
  /** khung nét đứt của Chổi định dạng (không đụng bảng tạm) */
  marquee: { sheet: number; range: Range } | null = null;
  zoom = 1;
  gridlines = true;
  readonly scroller: HTMLDivElement;
  private spacer: HTMLDivElement;
  readonly canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private dirty = false;
  private extR = 200;
  private extC = 40;
  private drag: Drag | null = null;
  private pointer = { x: 0, y: 0, id: -1, type: '' };
  private autoTimer = 0;
  private colors: Record<string, string> = {};
  private lastTap = { t: 0, r: -1, c: -1 };
  private longPress = 0;
  onSelect: (s: Selection) => void = () => {};
  onContext: (x: number, y: number, zone: 'cell' | 'row' | 'col') => void = () => {};
  onDraw: () => void = () => {};
  edit!: EditHost;

  constructor(private host: HTMLElement, private wb: Workbook) {
    this.sheet = wb.sheetIds()[0];
    host.classList.add('grid-host');
    this.scroller = document.createElement('div');
    this.scroller.className = 'grid-scroll';
    this.scroller.tabIndex = -1;
    this.spacer = document.createElement('div');
    this.spacer.className = 'grid-spacer';
    this.scroller.appendChild(this.spacer);
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'grid-canvas';
    this.canvas.setAttribute('role', 'img');
    this.canvas.setAttribute('aria-label', 'Lưới bảng tính');
    host.append(this.scroller, this.canvas);
    this.ctx = this.canvas.getContext('2d')!;
    this.readColors();
    new ResizeObserver(() => this.layout()).observe(host);
    this.scroller.addEventListener('scroll', () => { this.growIfNeeded(); this.invalidate(); });
    this.scroller.addEventListener('pointerdown', (e) => this.onDown(e));
    this.scroller.addEventListener('pointermove', (e) => this.onMove(e));
    this.scroller.addEventListener('pointerup', (e) => this.onUp(e));
    this.scroller.addEventListener('pointercancel', () => this.cancelDrag());
    this.scroller.addEventListener('dblclick', (e) => this.onDbl(e));
    // không để lưới giành focus (giữ focus ở ô nhập / thanh công thức)
    this.scroller.addEventListener('mousedown', (e) => e.preventDefault());
    this.scroller.addEventListener('contextmenu', (e) => { e.preventDefault(); this.contextAt(e.clientX, e.clientY); });
    this.scroller.addEventListener('wheel', (e) => {
      // Ctrl+Shift+cuộn: cuộn ngang (như Excel); Ctrl+cuộn: thu phóng
      if (e.ctrlKey && e.shiftKey) { e.preventDefault(); this.scroller.scrollLeft += (e.deltaY || e.deltaX); return; }
      if (e.ctrlKey) { e.preventDefault(); this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); }
    }, { passive: false });
    wb.onChange(() => { this.updateExtent(); this.invalidate(); });
    this.layout();
  }

  // ---------- màu theo theme ----------
  readColors() {
    const cs = getComputedStyle(this.host);
    this.fontCache.clear(); this.mCache?.clear();
    for (const k of ['cell', 'line', 'hdr', 'hdr-fg', 'hdr-on', 'hdr-on-fg', 'accent', 'sel-bg', 'fg', 'muted', 'freeze', 'hdr-line', 'err'])
      this.colors[k] = cs.getPropertyValue('--' + k).trim() || '#888';
    FONT = cs.getPropertyValue('--font-cell').trim() || FONT;
    UIFONT = cs.getPropertyValue('--font-ui').trim() || UIFONT;
    this.invalidate();
  }

  // ---------- hình học ----------
  private get m() { return this.wb.meta(this.sheet); }
  private hdrCache = { key: '', w: 0 };
  /** Hiện đầu đề hàng/cột (Xem ▸ Đầu đề) */
  headings = true;
  /** Xem ▸ Ô Tiêu điểm: tô dải hàng và cột của ô hiện hành */
  focusCell = false;
  /** hàng/cột của vùng (ngăn cố định) đang vẽ — để dải Ô Tiêu điểm không lấn sang ngăn khác */
  private region: { rows: number[]; cols: number[] } | null = null;
  /** Hiện Công thức (Ctrl+`) */
  showFormulas = false;
  /** Dạng xem: Thường / Xem trước Ngắt Trang */
  view: 'normal' | 'break' = 'normal';
  get hdrW() {
    if (!this.headings) return 0;
    if (this.frame) return this.frame.hw;
    const k = this.scroller.scrollTop + '|' + this.zoom + '|' + this.sheet + '|' + this.height;
    if (this.hdrCache.key === k) return this.hdrCache.w;
    const z = this.zoom;
    const lastRow = this.m.rowH.indexAt((this.scroller.scrollTop + this.height) / z + this.m.rowH.offset(this.fz.rows));
    const digits = String(Math.max(lastRow + 1, 100)).length;
    this.hdrCache = { key: k, w: Math.round((10 + digits * 7) * z) };
    return this.hdrCache.w;
  }
  get hdrH() { return this.headings ? Math.round(HDR_H * this.zoom) : 0; }
  colX(c: number) { return this.m.colW.offset(c) * this.zoom; }
  rowY(r: number) { return this.m.rowH.offset(r) * this.zoom; }
  colWpx(c: number) { return this.m.colW.get(c) * this.zoom; }
  rowHpx(r: number) { return this.m.rowH.get(r) * this.zoom; }
  private get fz() { return this.m.freeze; }
  private get frozenW() { return this.colX(this.fz.cols); }
  private get frozenH() { return this.rowY(this.fz.rows); }
  /** Giá trị đọc một lần cho mỗi khung hình (tránh đọc lại bố cục DOM cho từng ô) */
  private frame: { left: number; top: number; w: number; h: number; hw: number } | null = null;
  get width() { return this.frame?.w ?? this.scroller.clientWidth; }
  get height() { return this.frame?.h ?? this.scroller.clientHeight; }

  /** Toạ độ màn hình (trong lưới) của cạnh trái cột c */
  sx(c: number) {
    return c < this.fz.cols ? this.hdrW + this.colX(c) : this.hdrW + this.frozenW + this.colX(c) - this.colX(this.fz.cols) - (this.frame?.left ?? this.scroller.scrollLeft);
  }
  sy(r: number) {
    return r < this.fz.rows ? this.hdrH + this.rowY(r) : this.hdrH + this.frozenH + this.rowY(r) - this.rowY(this.fz.rows) - (this.frame?.top ?? this.scroller.scrollTop);
  }
  colAt(x: number) {
    const z = this.zoom;
    if (x < this.hdrW + this.frozenW) return this.m.colW.indexAt((x - this.hdrW) / z);
    return this.m.colW.indexAt((x - this.hdrW - this.frozenW + this.scroller.scrollLeft) / z + this.m.colW.offset(this.fz.cols));
  }
  rowAt(y: number) {
    const z = this.zoom;
    if (y < this.hdrH + this.frozenH) return this.m.rowH.indexAt((y - this.hdrH) / z);
    return this.m.rowH.indexAt((y - this.hdrH - this.frozenH + this.scroller.scrollTop) / z + this.m.rowH.offset(this.fz.rows));
  }
  cellRect(r: number, c: number) { return { x: this.sx(c), y: this.sy(r), w: this.colWpx(c), h: this.rowHpx(r) }; }

  private scrollCols(): number[] {
    const out: number[] = [];
    let c = this.colAt(this.hdrW + this.frozenW + 0.01);
    const W = this.width;
    while (c < MAX_COLS && this.sx(c) < W) { if (c >= this.fz.cols) out.push(c); c++; }
    return out;
  }
  private scrollRows(): number[] {
    const out: number[] = [];
    let r = this.rowAt(this.hdrH + this.frozenH + 0.01);
    const H = this.height;
    while (r < MAX_ROWS && this.sy(r) < H) { if (r >= this.fz.rows) out.push(r); r++; }
    return out;
  }
  visibleRows() { return [...Array(this.fz.rows).keys(), ...this.scrollRows()]; }
  visibleCols() { return [...Array(this.fz.cols).keys(), ...this.scrollCols()]; }

  // ---------- kích thước vùng cuộn ----------
  private updateExtent() {
    const u = this.wb.used(this.sheet);
    this.extR = Math.min(MAX_ROWS, Math.max(this.extR, u.rows + ROW_STEP, this.sel.fr + 50, this.sel.ar + 50));
    this.extC = Math.min(MAX_COLS, Math.max(this.extC, u.cols + COL_STEP, this.sel.fc + 10, this.sel.ac + 10));
    const w = this.hdrW + this.colX(this.extC) - this.frozenW;
    let h = this.hdrH + this.rowY(this.extR) - this.frozenH;
    if (h > MAX_SCROLL_PX) h = MAX_SCROLL_PX;
    this.spacer.style.width = w + 'px';
    this.spacer.style.height = h + 'px';
  }
  private growIfNeeded() {
    const s = this.scroller;
    let grew = false;
    if (s.scrollTop + s.clientHeight > s.scrollHeight - 400 && this.extR < MAX_ROWS) { this.extR = Math.min(MAX_ROWS, this.extR + ROW_STEP); grew = true; }
    if (s.scrollLeft + s.clientWidth > s.scrollWidth - 300 && this.extC < MAX_COLS) { this.extC = Math.min(MAX_COLS, this.extC + COL_STEP); grew = true; }
    if (grew) this.updateExtent();
  }
  layout() {
    const dpr = window.devicePixelRatio || 1;
    const w = this.width, h = this.height;
    this.canvas.style.width = w + 'px';
    this.canvas.style.height = h + 'px';
    this.canvas.width = Math.max(1, Math.round(w * dpr));
    this.canvas.height = Math.max(1, Math.round(h * dpr));
    this.updateExtent();
    this.invalidate();
  }
  setSheet(id: number) {
    this.sheet = id;
    this.extR = 200; this.extC = 40;
    this.scroller.scrollTo(0, 0);
    this.select(0, 0);
    this.updateExtent();
  }
  setZoom(z: number) {
    this.zoom = Math.min(4, Math.max(0.25, Math.round(z * 100) / 100));
    this.updateExtent();
    this.invalidate();
    this.onSelect(this.sel);
  }

  // ---------- chọn ----------
  get range(): Range { return this.sel.box ?? normRange({ r: this.sel.ar, c: this.sel.ac }, { r: this.sel.fr, c: this.sel.fc }); }
  /** Đặt ô hiện hành bên trong vùng chọn mà không đổi vùng */
  setActive(r: number, c: number) { const box = this.range; this.sel = { ...this.sel, ar: r, ac: c, box }; this.ensureVisible(r, c); this.invalidate(); this.onSelect(this.sel); }
  /** Mọi vùng đang chọn (vùng hiện hành ở cuối) */
  ranges(): Range[] { return [...this.areas, this.range]; }
  /** Đặt vùng chọn nhiều khối (khối cuối là khối hiện hành) */
  setAreas(list: Range[]) {
    if (!list.length) return;
    this.areas = list.slice(0, -1);
    const g = list[list.length - 1];
    this.sel = { ar: g.r1, ac: g.c1, fr: g.r2, fc: g.c2 };
    this.ensureVisible(g.r1, g.c1);
    this.invalidate(); this.onSelect(this.sel);
  }
  /** Bỏ các vùng thêm, giữ vùng hiện hành */
  clearAreas() { if (this.areas.length) { this.areas = []; this.invalidate(); } }
  select(r: number, c: number, extend = false, keepAreas = false) {
    r = clamp(r, 0, MAX_ROWS - 1); c = clamp(c, 0, MAX_COLS - 1);
    if (!extend) {
      if (keepAreas || this.addMode) { const g = this.range; if (!this.areas.some((a) => a.r1 === g.r1 && a.c1 === g.c1 && a.r2 === g.r2 && a.c2 === g.c2)) this.areas.push(g); }
      else this.areas = [];
    }
    if (extend) {
      if (this.sel.box) { const b = this.sel.box; this.sel = { ar: this.sel.ar === b.r1 || this.sel.ar === b.r2 ? this.sel.ar : b.r1, ac: this.sel.ac === b.c1 || this.sel.ac === b.c2 ? this.sel.ac : b.c1, fr: r, fc: c }; }
      this.sel.fr = r; this.sel.fc = c;
      const g = this.wb.expandToMerges(this.sheet, this.range);
      this.sel.fr = this.sel.fr >= this.sel.ar ? g.r2 : g.r1;
      this.sel.fc = this.sel.fc >= this.sel.ac ? g.c2 : g.c1;
      if (this.sel.ar < g.r1 || this.sel.ar > g.r2) this.sel.ar = this.sel.fr >= this.sel.ar ? g.r1 : g.r2;
    } else {
      const m = this.wb.mergeAt(this.sheet, r, c);
      this.sel = m ? { ar: m.r1, ac: m.c1, fr: m.r2, fc: m.c2 } : { ar: r, ac: c, fr: r, fc: c };
    }
    this.updateExtent();
    this.ensureVisible(extend ? r : this.sel.ar, extend ? c : this.sel.ac);
    this.invalidate();
    this.onSelect(this.sel);
  }
  selectRange(g0: Range, keepAreas = false) {
    if (keepAreas) { const cur = this.range; this.areas.push(cur); } else this.areas = [];
    const g = g0.r2 >= MAX_ROWS - 1 || g0.c2 >= MAX_COLS - 1 ? g0 : this.wb.expandToMerges(this.sheet, g0);
    this.sel = { ar: g.r1, ac: g.c1, fr: g.r2, fc: g.c2 };
    this.updateExtent();
    this.ensureVisible(g.r1, g.c1);
    this.invalidate();
    this.onSelect(this.sel);
  }
  /** Bước một ô theo mũi tên: nhảy qua vùng gộp và hàng đang ẩn (lọc) */
  step(dr: number, dc: number, extend = false) {
    const m = this.m;
    let r: number, c: number;
    if (extend) { r = this.sel.fr; c = this.sel.fc; }
    else {
      const mg = this.wb.mergeAt(this.sheet, this.sel.ar, this.sel.ac);
      r = mg ? (dr > 0 ? mg.r2 : mg.r1) : this.sel.ar;
      c = mg ? (dc > 0 ? mg.c2 : mg.c1) : this.sel.ac;
    }
    r += dr; c += dc;
    while (dr && r > 0 && r < MAX_ROWS - 1 && m.rowH.get(r) === 0) r += dr;
    while (dc && c > 0 && c < MAX_COLS - 1 && m.colW.get(c) === 0) c += dc;
    this.select(r, c, extend);
  }
  moveActive(dr: number, dc: number) {
    // Enter/Tab trong vùng chọn nhiều ô: đi vòng trong vùng như Excel
    const g = this.range;
    if (this.areas.length && this.isSingle()) { this.nextArea(dr + dc > 0); return; }
    if (!this.isSingle()) {
      let { ar, ac } = this.sel;
      if (dc) { ac += dc; if (ac > g.c2) { ac = g.c1; ar = ar + 1 > g.r2 ? g.r1 : ar + 1; } if (ac < g.c1) { ac = g.c2; ar = ar - 1 < g.r1 ? g.r2 : ar - 1; } }
      if (dr) { ar += dr; if (ar > g.r2) { ar = g.r1; ac = ac + 1 > g.c2 ? g.c1 : ac + 1; } if (ar < g.r1) { ar = g.r2; ac = ac - 1 < g.c1 ? g.c2 : ac - 1; } }
      // hết khối hiện hành → sang khối kế tiếp (vùng chọn nhiều khối)
      const wrapped = (dr + dc > 0) ? (ar === g.r1 && ac === g.c1) : (ar === g.r2 && ac === g.c2);
      if (wrapped && this.areas.length) { this.nextArea(dr + dc > 0); return; }
      this.sel = { ...this.sel, ar, ac, box: g };
      this.ensureVisible(ar, ac); this.invalidate(); this.onSelect(this.sel);
      return;
    }
    this.step(dr, dc);
  }
  /** Enter/Tab đi hết khối → sang khối kế tiếp (hoặc trước) */
  private nextArea(fwd: boolean) {
    const g = this.range, all = this.ranges();
    const nx = fwd ? all[0] : all[all.length - 2];
    this.areas = fwd ? [...all.slice(1, -1), g] : [g, ...all.slice(0, -2)];
    this.sel = fwd ? { ar: nx.r1, ac: nx.c1, fr: nx.r2, fc: nx.c2 } : { ar: nx.r2, ac: nx.c2, fr: nx.r1, fc: nx.c1 };
    this.ensureVisible(this.sel.ar, this.sel.ac); this.invalidate(); this.onSelect(this.sel);
  }
  ensureVisible(r: number, c: number) {
    const s = this.scroller;
    if (r >= this.fz.rows) {
      const top = this.rowY(r) - this.rowY(this.fz.rows);
      const viewH = this.height - this.hdrH - this.frozenH;
      if (top < s.scrollTop) s.scrollTop = top;
      else if (top + this.rowHpx(r) > s.scrollTop + viewH) s.scrollTop = top + this.rowHpx(r) - viewH;
    }
    if (c >= this.fz.cols) {
      const left = this.colX(c) - this.colX(this.fz.cols);
      const viewW = this.width - this.hdrW - this.frozenW;
      if (left < s.scrollLeft) s.scrollLeft = left;
      else if (left + this.colWpx(c) > s.scrollLeft + viewW) s.scrollLeft = Math.min(left, left + this.colWpx(c) - viewW);
    }
  }
  /** Ctrl+mũi tên: nhảy tới mép vùng dữ liệu */
  jump(dr: number, dc: number, extend: boolean) {
    const from = extend ? { r: this.sel.fr, c: this.sel.fc } : { r: this.sel.ar, c: this.sel.ac };
    const u = this.wb.used(this.sheet);
    const has = (r: number, c: number) => { const v = this.wb.value(this.sheet, r, c); return v !== null && v !== ''; };
    let { r, c } = from;
    const inb = (r: number, c: number) => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS;
    const step = () => { r += dr; c += dc; };
    const limR = dr > 0 ? (r >= u.rows ? MAX_ROWS - 1 : u.rows) : 0;
    const limC = dc > 0 ? (c >= u.cols ? MAX_COLS - 1 : u.cols) : 0;
    const atLimit = () => (dr ? r === limR : c === limC) || !inb(r + dr, c + dc);
    if (has(r, c) && inb(r + dr, c + dc) && has(r + dr, c + dc)) {
      while (inb(r + dr, c + dc) && has(r + dr, c + dc)) step();
    } else {
      step();
      while (inb(r, c) && !has(r, c) && !atLimit()) step();
      if (!inb(r, c)) { r -= dr; c -= dc; }
      if (!has(r, c)) { r = dr > 0 ? MAX_ROWS - 1 : dr < 0 ? 0 : r; c = dc > 0 ? MAX_COLS - 1 : dc < 0 ? 0 : c; }
    }
    this.select(r, c, extend);
  }

  // ---------- vẽ ----------
  invalidate() {
    if (this.dirty) return;
    this.dirty = true;
    requestAnimationFrame(() => { this.dirty = false; this.draw(); });
  }
  draw() {
    this.frame = null;
    const hw = this.hdrW;
    this.frame = { left: this.scroller.scrollLeft, top: this.scroller.scrollTop, w: this.scroller.clientWidth, h: this.scroller.clientHeight, hw };
    try { this.draw0(); } finally { this.frame = null; }
  }
  private draw0() {
    this.lastFont = '';
    const ctx = this.ctx, dpr = window.devicePixelRatio || 1, W = this.width, H = this.height, C = this.colors;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = C.cell;
    ctx.fillRect(0, 0, W, H);
    const fr = this.fz.rows, fc = this.fz.cols;
    const rowsF = [...Array(fr).keys()], colsF = [...Array(fc).keys()];
    const rowsS = this.scrollRows(), colsS = this.scrollCols();
    const x0 = this.hdrW, y0 = this.hdrH, xF = x0 + this.frozenW, yF = y0 + this.frozenH;
    const regions: [number[], number[], number, number, number, number][] = [
      [rowsS, colsS, xF, yF, W - xF, H - yF],
      [rowsF, colsS, xF, y0, W - xF, yF - y0],
      [rowsS, colsF, x0, yF, xF - x0, H - yF],
      [rowsF, colsF, x0, y0, xF - x0, yF - y0],
    ];
    for (const [rows, cols, x, y, w, h] of regions) {
      if (!rows.length || !cols.length || w <= 0 || h <= 0) continue;
      ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
      this.drawCells(rows, cols, x + w);
      if (this.view === 'break') this.drawBreaks(rows, cols);
      this.region = { rows, cols };
      this.drawSelection();
      ctx.restore();
    }
    this.drawHeaders(rowsF.concat(rowsS), colsF.concat(colsS));
    if (fr || fc) {
      ctx.strokeStyle = C.freeze; ctx.lineWidth = 1.5;
      ctx.beginPath();
      if (fr) { ctx.moveTo(0, yF + 0.5); ctx.lineTo(W, yF + 0.5); }
      if (fc) { ctx.moveTo(xF + 0.5, 0); ctx.lineTo(xF + 0.5, H); }
      ctx.stroke();
    }
    this.onDraw();
  }

  // ---------- Xem trước Ngắt Trang ----------
  private brk: { key: string; pages: { r1: number; r2: number; c1: number; c2: number; n: number }[]; area: Range | null; manR: Set<number>; manC: Set<number> } | null = null;
  private breakLayout() {
    const key = this.sheet + ':' + this.wb.ver;
    if (this.brk?.key === key) return this.brk;
    const L = paginate(this.wb, this.sheet, normPrint(this.wb.meta(this.sheet).print));
    const pages = L.pages.map((p, i) => ({ r1: p.rows[0], r2: p.rows[p.rows.length - 1], c1: p.cols[0], c2: p.cols[p.cols.length - 1], n: i + 1 }));
    const b = this.wb.meta(this.sheet).breaks;
    this.brk = { key, pages, area: L.area, manR: new Set(b.rows), manC: new Set(b.cols) };
    return this.brk;
  }
  private drawBreaks(rows: number[], cols: number[]) {
    const ctx = this.ctx, z = this.zoom, W = this.width, H = this.height;
    const L = this.breakLayout();
    void rows; void cols;
    // ngoài vùng in: tô xám như Excel
    ctx.save();
    ctx.fillStyle = 'rgba(128,128,128,.28)';
    if (!L.area) { ctx.fillRect(0, 0, W, H); ctx.restore(); return; }
    const ax1 = this.sx(L.area.c1), ay1 = this.sy(L.area.r1), ax2 = this.sx(L.area.c2) + this.colWpx(L.area.c2), ay2 = this.sy(L.area.r2) + this.rowHpx(L.area.r2);
    ctx.beginPath(); ctx.rect(0, 0, W, H); ctx.rect(ax1, ay1, ax2 - ax1, ay2 - ay1); ctx.fill('evenodd');
    // số trang mờ ở giữa mỗi trang
    ctx.font = `600 ${Math.round(44 * z)}px ${UIFONT}`;
    this.lastFont = '';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillStyle = 'rgba(120,120,120,.35)';
    for (const p of L.pages) {
      const x1 = this.sx(p.c1), y1 = this.sy(p.r1), x2 = this.sx(p.c2) + this.colWpx(p.c2), y2 = this.sy(p.r2) + this.rowHpx(p.r2);
      if (x2 < 0 || y2 < 0 || x1 > W || y1 > H) continue;
      ctx.fillText('Trang ' + p.n, (x1 + x2) / 2, (y1 + y2) / 2);
    }
    // đường ngắt: thủ công/biên vùng in nét liền, tự động nét đứt (màu xanh như Excel)
    ctx.strokeStyle = '#1f5fbf'; ctx.lineWidth = 2;
    ctx.strokeRect(ax1 + 1, ay1 + 1, ax2 - ax1 - 2, ay2 - ay1 - 2);
    const seenR = new Set<number>(), seenC = new Set<number>();
    for (const p of L.pages) {
      if (p.r1 !== L.area.r1 && !seenR.has(p.r1)) { seenR.add(p.r1); const y = Math.round(this.sy(p.r1)); ctx.setLineDash(L.manR.has(p.r1) ? [] : [6, 4]); ctx.beginPath(); ctx.moveTo(ax1, y); ctx.lineTo(ax2, y); ctx.stroke(); }
      if (p.c1 !== L.area.c1 && !seenC.has(p.c1)) { seenC.add(p.c1); const x = Math.round(this.sx(p.c1)); ctx.setLineDash(L.manC.has(p.c1) ? [] : [6, 4]); ctx.beginPath(); ctx.moveTo(x, ay1); ctx.lineTo(x, ay2); ctx.stroke(); }
    }
    ctx.restore();
  }

  /** Kiểu hiệu lực của ô: định dạng ô + định dạng có điều kiện */
  private eff(r: number, c: number) {
    let d = this.wb.display(this.sheet, r, c);
    if (this.showFormulas) { const t = this.wb.rawText(this.sheet, r, c); if (t.startsWith('=')) d = { ...d, text: t, align: 'l', value: t }; }
    const cf = this.wb.cfStyle(this.sheet, r, c);
    const st = cf.fg || cf.bg || cf.b || cf.i ? { ...d.style, ...(cf.fg ? { fg: cf.fg } : {}), ...(cf.bg ? { bg: cf.bg } : {}), ...(cf.b ? { b: true } : {}), ...(cf.i ? { i: true } : {}) } : d.style;
    return { d, st, bar: cf.bar };
  }
  private lastFont = '';
  private mCache = new Map<string, number>();
  /** Đo độ rộng chữ với phông đang đặt (có bộ nhớ đệm) */
  private measure(t: string) {
    const k = this.lastFont + '\u0001' + t;
    let w = this.mCache.get(k);
    if (w === undefined) { w = this.ctx.measureText(t).width; if (this.mCache.size > 20000) this.mCache.clear(); this.mCache.set(k, w); }
    return w;
  }
  private fontCache = new Map<string, { px: number; css: string }>();
  private fontOf(st: Style) {
    const k = (st.fs ?? 11) + '|' + (st.b ? 1 : 0) + (st.i ? 1 : 0) + '|' + (st.fn ?? '') + '|' + this.zoom;
    let f = this.fontCache.get(k);
    if (!f) { f = this.fontOf0(st); if (this.fontCache.size > 500) this.fontCache.clear(); this.fontCache.set(k, f); }
    return f;
  }
  private fontOf0(st: Style) {
    const px = ptPx(st.fs) * this.zoom;
    const fam = st.fn ? `"${st.fn.replace(/"/g, '')}", ${FONT}` : FONT;
    return { px, css: `${st.i ? 'italic ' : ''}${st.b ? 700 : 400} ${px}px ${fam}` };
  }
  /** Hình chữ nhật của ô (gộp thì cả vùng gộp) */
  cellRectM(r: number, c: number) {
    const m = this.wb.mergeAt(this.sheet, r, c);
    if (!m) return this.cellRect(r, c);
    const x = this.sx(m.c1), y = this.sy(m.r1);
    return { x, y, w: this.sx(m.c2) + this.colWpx(m.c2) - x, h: this.sy(m.r2) + this.rowHpx(m.r2) - y };
  }

  private drawCells(rows: number[], cols: number[], regionRight: number) {
    const ctx = this.ctx, C = this.colors, z = this.zoom, s = this.sheet, wb = this.wb;
    const used = wb.used(s);
    const merges = wb.meta(s).merges.filter((m) => rows.length && cols.length && m.r2 >= rows[0] && m.r1 <= rows[rows.length - 1] && m.c2 >= cols[0] && m.c1 <= cols[cols.length - 1]);
    // 1. đường lưới
    if (this.gridlines) {
      ctx.strokeStyle = C.line; ctx.lineWidth = 1;
      ctx.beginPath();
      for (const r of rows) { const h = this.rowHpx(r); if (h <= 0) continue; const y = Math.round(this.sy(r) + h) - 0.5; ctx.moveTo(0, y); ctx.lineTo(this.width, y); }
      for (const c of cols) { const w = this.colWpx(c); if (w <= 0) continue; const x = Math.round(this.sx(c) + w) - 0.5; ctx.moveTo(x, 0); ctx.lineTo(x, this.height); }
      ctx.stroke();
    }
    // 2. nền ô (che đường lưới như Excel), thanh dữ liệu
    const effs = new Map<number, ReturnType<Grid['eff']>>();
    const E = (r: number, c: number) => { const k = r * MAX_COLS + c; let e = effs.get(k); if (!e) { e = this.eff(r, c); effs.set(k, e); } return e; };
    const lastR = Math.min(rows[rows.length - 1] ?? 0, used.rows - 1);
    for (const r of rows) {
      const h = this.rowHpx(r);
      if (h <= 0) continue;
      const y = this.sy(r);
      for (const c of cols) {
        if (merges.length && wb.mergeAt(s, r, c)) continue;
        const st = wb.style(s, r, c);
        const cfOn = wb.meta(s).cf.length > 0 && r <= lastR;
        const e = cfOn ? E(r, c) : null;
        const bg = e ? e.st.bg : st.bg;
        if (bg) { ctx.fillStyle = bg; ctx.fillRect(this.sx(c), y, this.colWpx(c), h); }
        if (e?.bar) { ctx.fillStyle = e.bar.color; ctx.globalAlpha = 0.55; ctx.fillRect(this.sx(c) + 2, y + 3, Math.max(0, (this.colWpx(c) - 4) * e.bar.frac), h - 6); ctx.globalAlpha = 1; }
      }
    }
    for (const m of merges) {
      const rc = this.cellRectM(m.r1, m.c1);
      const e = E(m.r1, m.c1);
      ctx.fillStyle = e.st.bg || C.cell;
      ctx.fillRect(rc.x, rc.y, rc.w - 1, rc.h - 1);
    }
    // 3. viền
    for (const r of rows) {
      const h = this.rowHpx(r);
      if (h <= 0) continue;
      for (const c of cols) {
        const bd = wb.style(s, r, c).bd;
        if (!bd) continue;
        const x = this.sx(c), y = this.sy(r), w = this.colWpx(c);
        const m = merges.length ? wb.mergeAt(s, r, c) : null;
        const edge = (b: Border | undefined, x1: number, y1: number, x2: number, y2: number) => {
          if (!b) return;
          ctx.strokeStyle = b.c || C.fg;
          const lw = b.s === 'thick' ? 3 : b.s === 'medium' ? 2 : 1;
          ctx.lineWidth = lw;
          ctx.setLineDash(b.s === 'dashed' ? [4 * z, 2 * z] : b.s === 'dotted' ? [1, 2] : []);
          const off = lw % 2 ? 0.5 : 0;
          ctx.beginPath();
          if (b.s === 'double') {
            const v = x1 === x2;
            ctx.moveTo(Math.round(x1) + (v ? -1 : 0) + off, Math.round(y1) + (v ? 0 : -1) + off); ctx.lineTo(Math.round(x2) + (v ? -1 : 0) + off, Math.round(y2) + (v ? 0 : -1) + off);
            ctx.moveTo(Math.round(x1) + (v ? 1 : 0) + off, Math.round(y1) + (v ? 0 : 1) + off); ctx.lineTo(Math.round(x2) + (v ? 1 : 0) + off, Math.round(y2) + (v ? 0 : 1) + off);
          } else { ctx.moveTo(Math.round(x1) + off, Math.round(y1) + off); ctx.lineTo(Math.round(x2) + off, Math.round(y2) + off); }
          ctx.stroke();
        };
        if (!m || r === m.r1) edge(bd.t, x - 0.5, y - 0.5, x + w - 0.5, y - 0.5);
        if (!m || r === m.r2) edge(bd.b, x - 0.5, y + h - 1, x + w - 0.5, y + h - 1);
        if (!m || c === m.c1) edge(bd.l, x - 0.5, y - 0.5, x - 0.5, y + h - 0.5);
        if (!m || c === m.c2) edge(bd.r, x + w - 1, y - 0.5, x + w - 1, y + h - 0.5);
        if (bd.d || bd.u) {
          const g = m ? this.cellRectM(m.r1, m.c1) : { x, y, w, h };
          if (!m || (r === m.r1 && c === m.c1)) { edge(bd.d, g.x, g.y, g.x + g.w - 1, g.y + g.h - 1); edge(bd.u, g.x, g.y + g.h - 1, g.x + g.w - 1, g.y); }
        }
      }
    }
    ctx.setLineDash([]);
    // 4. chữ
    ctx.textBaseline = 'alphabetic';
    for (const r of rows) {
      if (r >= used.rows) break;
      const h = this.rowHpx(r);
      if (h <= 0) continue;
      for (let i = 0; i < cols.length; i++) {
        const c = cols[i];
        if (c >= used.cols) break;
        if (merges.length && wb.mergeAt(s, r, c)) continue;
        const e = E(r, c);
        if (!e.d.text) continue;
        this.paintText(e, this.sx(c), this.sy(r), this.colWpx(c), h, cols, i, regionRight, r);
      }
    }
    for (const m of merges) {
      const e = E(m.r1, m.c1);
      if (!e.d.text) continue;
      const rc = this.cellRectM(m.r1, m.c1);
      this.paintText(e, rc.x, rc.y, rc.w, rc.h, null, 0, regionRight, m.r1);
    }
    // 5. nút lọc ở hàng tiêu đề
    const f = wb.meta(s).filter;
    if (f && rows.includes(f.range.r1)) {
      for (const c of cols) {
        if (c < f.range.c1 || c > f.range.c2) continue;
        const b = this.filterBtn(f.range.r1, c);
        if (!b) continue;
        const on = !!f.conds[c];
        ctx.fillStyle = on ? C.accent : C.hdr;
        ctx.strokeStyle = on ? C.accent : C['hdr-line'];
        ctx.lineWidth = 1;
        ctx.beginPath(); ctx.roundRect(b.x + 0.5, b.y + 0.5, b.w - 1, b.h - 1, 3); ctx.fill(); ctx.stroke();
        ctx.fillStyle = on ? C.cell : C['hdr-fg'];
        const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
        ctx.beginPath();
        if (on) { ctx.moveTo(cx - 4 * z, cy - 3 * z); ctx.lineTo(cx + 4 * z, cy - 3 * z); ctx.lineTo(cx + 1 * z, cy + 1 * z); ctx.lineTo(cx + 1 * z, cy + 4 * z); ctx.lineTo(cx - 1 * z, cy + 4 * z); ctx.lineTo(cx - 1 * z, cy + 1 * z); }
        else { ctx.moveTo(cx - 3.5 * z, cy - 1.5 * z); ctx.lineTo(cx + 3.5 * z, cy - 1.5 * z); ctx.lineTo(cx, cy + 2.5 * z); }
        ctx.closePath(); ctx.fill();
      }
    }
  }
  /** Nút mũi tên danh sách cạnh ô đang chọn (nếu ô có kiểm tra dữ liệu dạng danh sách) */
  dvBtn() {
    const { ar, ac } = this.sel;
    const d = this.wb.dvAt(this.sheet, ar, ac);
    if (!d || d.type !== 'list') return null;
    const a = this.cellRectM(ar, ac);
    const sz = Math.max(18 * this.zoom, Math.min(a.h, 24 * this.zoom));
    return { x: a.x + a.w + 1, y: a.y + a.h - sz, w: sz, h: sz };
  }
  /** Vùng chọn chỉ là một ô (hoặc đúng một vùng gộp) */
  isSingle() {
    const g = this.range;
    if (g.r1 === g.r2 && g.c1 === g.c2) return true;
    const m = this.wb.mergeAt(this.sheet, this.sel.ar, this.sel.ac);
    return !!m && m.r1 === g.r1 && m.r2 === g.r2 && m.c1 === g.c1 && m.c2 === g.c2;
  }
  /** Chiều cao hàng cần để hiện đủ chữ (ô xuống dòng, chữ to) — theo đơn vị px ở mức 100% */
  neededHeight(r: number) {
    const ctx = this.ctx, u = this.wb.used(this.sheet);
    let need = 0;
    for (let c = 0; c < Math.min(u.cols, 200); c++) {
      const d = this.wb.display(this.sheet, r, c);
      if (!d.text) continue;
      const st = d.style;
      const px = ptPx(st.fs);
      if (!st.wr) { need = Math.max(need, Math.ceil(px * 1.36)); continue; }
      ctx.font = `${st.i ? 'italic ' : ''}${st.b ? 700 : 400} ${px}px ${st.fn ? `"${st.fn}", ` : ''}${FONT}`;
      this.lastFont = '';
      const m = this.wb.mergeAt(this.sheet, r, c);
      const w = m ? this.m.colW.offset(m.c2 + 1) - this.m.colW.offset(m.c1) : this.m.colW.get(c);
      let lines = 0;
      for (const para of d.text.split('\n')) {
        let line = '';
        lines++;
        for (const word of para.split(/(\s+)/)) {
          const t = line + word;
          if (ctx.measureText(t).width > w - 8 && line.trim()) { lines++; line = word.trimStart(); } else line = t;
        }
      }
      need = Math.max(need, Math.ceil(lines * px * 1.22 + 4));
    }
    return need;
  }
  /** Vị trí nút lọc của một cột tiêu đề */
  filterBtn(r: number, c: number) {
    const w = this.colWpx(c), h = this.rowHpx(r);
    if (w < 18 || h < 12) return null;
    const sz = Math.min(18 * this.zoom, h - 4);
    return { x: this.sx(c) + w - sz - 2, y: this.sy(r) + (h - sz) / 2, w: sz, h: sz };
  }

  private paintText(e: { d: { text: string; align: string; error: boolean; value: unknown }; st: Style }, x: number, y: number, w: number, h: number, cols: number[] | null, i: number, regionRight: number, r: number) {
    const ctx = this.ctx, C = this.colors, z = this.zoom, wb = this.wb, s = this.sheet;
    const { d } = e;
    let st = e.st;
    if (w <= 0) return;
    if (st.rot) { this.paintRotated(e, x, y, w, h); return; }
    if (st.vp) st = { ...st, fs: (st.fs ?? 11) * 0.7 };
    let f = this.fontOf(st);
    // đặt phông chỉ khi đổi (gán ctx.font tốn thời gian phân tích chuỗi)
    if (this.lastFont !== f.css) { ctx.font = f.css; this.lastFont = f.css; }
    const pad = 4 * z;
    const color = d.error ? C.err : st.fg || (st.bg ? autoText(st.bg) : C.fg);
    let al = d.align;
    const isNum = typeof d.value === 'number';
    const indent = (st.ind ?? 0) * 9 * z;
    if (st.hx === 'dist') al = 'c';
    // xuống dòng trong ô (Căn đều / Phân phối cũng xuống dòng như Excel)
    if ((st.wr || st.hx === 'just' || st.hx === 'dist') && !isNum) {
      const lines: string[] = [];
      for (const para of d.text.split('\n')) {
        let line = '';
        for (const word of para.split(/(\s+)/)) {
          const t = line + word;
          if (ctx.measureText(t).width > w - 2 * pad && line.trim()) { lines.push(line.trimEnd()); line = word.trimStart(); }
          else line = t;
        }
        lines.push(line);
      }
      const lh = f.px * 1.25;
      const total = lines.length * lh;
      let ty = st.va === 't' ? y + pad / 2 + f.px : st.va === 'm' ? y + (h - total) / 2 + f.px * 0.95 : y + h - total - pad / 2 + f.px * 0.95;
      ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
      ctx.fillStyle = color;
      for (const ln of lines) {
        const tw = ctx.measureText(ln).width;
        const tx = al === 'r' ? x + w - pad - tw - indent : al === 'c' ? x + (w - tw) / 2 : x + pad + indent;
        ctx.fillText(ln, tx, ty);
        if (st.u) { ctx.fillRect(tx, ty + 2 * z, tw, Math.max(1, z)); if (st.ud) ctx.fillRect(tx, ty + 4 * z, tw, Math.max(1, z)); }
        if (st.s) ctx.fillRect(tx, ty - f.px * 0.32, tw, Math.max(1, z));
        ty += lh;
      }
      ctx.restore();
      return;
    }
    let text = d.text.replace(/\n/g, ' ');
    let tw = this.measure(text);
    let clipW = w;
    // Co chữ vừa ô
    if (st.shk && tw > w - 2 * pad - indent && tw > 0) {
      const k = Math.max(0.2, (w - 2 * pad - indent) / tw);
      st = { ...st, fs: (st.fs ?? 11) * k };
      f = this.fontOf(st); ctx.font = f.css; this.lastFont = f.css;
      tw = this.measure(text);
    }
    // Điền: lặp chữ cho kín ô
    if (st.hx === 'fill' && text && tw > 0 && !isNum) {
      const n = Math.max(1, Math.floor((w - 2 * pad) / tw));
      text = text.repeat(n); tw = this.measure(text); al = 'l';
    }
    // Căn giữa qua vùng chọn: trải sang các ô trống bên phải cùng kiểu
    let spanW = w;
    if (st.hx === 'cas' && cols) {
      let j = i + 1;
      while (j < cols.length) {
        const cc = cols[j], v = wb.value(s, r, cc);
        if ((v !== null && v !== '') || wb.style(s, r, cc).hx !== 'cas') break;
        spanW += this.colWpx(cc); j++;
      }
      clipW = spanW; al = 'c';
    }
    if (isNum && tw > w - 2 * pad) {
      text = '#'.repeat(Math.max(1, Math.floor((w - 2 * pad) / ctx.measureText('#').width)));
      tw = ctx.measureText(text).width;
    } else if (cols && !isNum && al === 'l' && tw > w - 2 * pad) {
      // tràn chữ sang các ô trống bên phải
      let j = i + 1;
      while (j < cols.length && clipW < tw + 2 * pad && x + clipW < regionRight) {
        const cc = cols[j];
        const v = wb.value(s, r, cc);
        if ((v !== null && v !== '') || wb.mergeAt(s, r, cc)) break;
        clipW += this.colWpx(cc);
        j++;
      }
      if (clipW > w) { ctx.fillStyle = st.bg || C.cell; ctx.fillRect(x + w - 1, y, clipW - w, h - 1); }
    }
    const needClip = tw > clipW - 2 * pad || f.px > h;
    if (needClip) { ctx.save(); ctx.beginPath(); ctx.rect(x, y, clipW, h); ctx.clip(); }
    ctx.fillStyle = color;
    const tx = al === 'r' ? x + w - pad - tw - indent : al === 'c' ? x + (spanW - tw) / 2 : x + pad + indent;
    let ty = st.va === 't' ? y + pad / 2 + f.px * 0.95 : st.va === 'm' ? y + h / 2 + f.px * 0.35 : y + h - Math.max(5 * z, (h - f.px) / 2);
    if (e.st.vp === 'sup') ty -= f.px * 0.45; else if (e.st.vp === 'sub') ty += f.px * 0.25;
    ctx.fillText(text, tx, ty);
    if (st.u) { ctx.fillRect(tx, ty + 2 * z, tw, Math.max(1, z)); if (st.ud) ctx.fillRect(tx, ty + 4 * z, tw, Math.max(1, z)); }
    if (st.s) ctx.fillRect(tx, ty - f.px * 0.32, tw, Math.max(1, z));
    if (needClip) ctx.restore();
  }

  /** Chữ xoay (Định dạng Ô ▸ Căn chỉnh ▸ Hướng): góc −90…90, 255 = chữ dọc xếp chồng */
  private paintRotated(e: { d: { text: string; align: string; error: boolean; value: unknown }; st: Style }, x: number, y: number, w: number, h: number) {
    const ctx = this.ctx, C = this.colors, z = this.zoom, { d, st } = e;
    const f = this.fontOf(st);
    ctx.font = f.css; this.lastFont = f.css;
    const color = d.error ? C.err : st.fg || (st.bg ? autoText(st.bg) : C.fg);
    ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
    ctx.fillStyle = color;
    const text = d.text.replace(/\n/g, ' ');
    if (st.rot === 255) {
      const lh = f.px * 1.1, total = [...text].length * lh;
      let ty = st.va === 't' ? y + 2 * z + f.px : st.va === 'm' ? y + (h - total) / 2 + f.px : y + h - total + f.px * 0.8;
      for (const ch of text) { const cw = ctx.measureText(ch).width; ctx.fillText(ch, x + (w - cw) / 2, ty); ty += lh; }
    } else {
      const a = -(st.rot ?? 0) * Math.PI / 180;
      const tw = ctx.measureText(text).width;
      const cx = x + w / 2, cy = y + h / 2;
      const al = d.align;
      const sx = al === 'r' ? x + w - 4 * z : al === 'c' ? cx : x + 4 * z;
      const sy = st.va === 't' ? y + 4 * z : st.va === 'm' ? cy : y + h - 4 * z;
      ctx.translate(sx, sy); ctx.rotate(a);
      const off = al === 'r' ? -tw : al === 'c' ? -tw / 2 : 0;
      ctx.textBaseline = 'middle';
      ctx.fillText(text, (st.rot ?? 0) > 0 && st.va !== 't' ? off : off, 0);
      if (st.u) ctx.fillRect(off, f.px * 0.55, tw, Math.max(1, z));
      ctx.textBaseline = 'alphabetic';
    }
    ctx.restore();
  }

  private drawSelection() {
    const ctx = this.ctx, C = this.colors, z = this.zoom;
    const g = this.range;
    const x1 = this.sx(g.c1), y1 = this.sy(g.r1);
    const x2 = this.sx(g.c2) + this.colWpx(g.c2), y2 = this.sy(g.r2) + this.rowHpx(g.r2);
    if (this.focusCell) {
      const a = this.cellRectM(this.sel.ar, this.sel.ac);
      const rg = this.region, rowIn = !rg || rg.rows.includes(this.sel.ar), colIn = !rg || rg.cols.includes(this.sel.ac);
      const W = this.canvas.width, H = this.canvas.height;
      ctx.save(); ctx.fillStyle = C.accent; ctx.globalAlpha = 0.12;
      if (rowIn) { if (colIn) { ctx.fillRect(0, a.y, a.x, a.h); ctx.fillRect(a.x + a.w, a.y, W, a.h); } else ctx.fillRect(0, a.y, W, a.h); }
      if (colIn) { if (rowIn) { ctx.fillRect(a.x, 0, a.w, a.y); ctx.fillRect(a.x, a.y + a.h, a.w, H); } else ctx.fillRect(a.x, 0, a.w, H); }
      ctx.restore();
    }
    // vùng copy (viền nét đứt)
    const cl = this.marquee ?? (this.wb.clip ? { sheet: this.wb.clip.sheet, range: this.wb.clip.range } : null);
    if (cl && cl.sheet === this.sheet) {
      const a = cl.range;
      ctx.save();
      ctx.setLineDash([5 * z, 3 * z]); ctx.strokeStyle = C.accent; ctx.lineWidth = 1.5;
      ctx.strokeRect(this.sx(a.c1) + 1, this.sy(a.r1) + 1, this.sx(a.c2) + this.colWpx(a.c2) - this.sx(a.c1) - 2, this.sy(a.r2) + this.rowHpx(a.r2) - this.sy(a.r1) - 2);
      ctx.restore();
    }
    const single = this.isSingle();
    if (this.areas.length) {
      // nhiều khối: tô mờ mọi khối, ô hiện hành viền mảnh (như Excel)
      ctx.fillStyle = C['sel-bg'];
      for (const a of this.ranges()) {
        const ax1 = this.sx(a.c1), ay1 = this.sy(a.r1);
        const ax2 = this.sx(a.c2) + this.colWpx(a.c2), ay2 = this.sy(a.r2) + this.rowHpx(a.r2);
        ctx.fillRect(ax1, ay1, ax2 - ax1, ay2 - ay1);
      }
      const a = this.cellRectM(this.sel.ar, this.sel.ac);
      ctx.fillStyle = C.cell; ctx.fillRect(a.x + 1, a.y + 1, a.w - 2, a.h - 2);
      this.redrawCell(this.sel.ar, this.sel.ac);
      ctx.strokeStyle = C.accent; ctx.lineWidth = 1;
      ctx.strokeRect(a.x + 0.5, a.y + 0.5, a.w - 1, a.h - 1);
      return;
    }
    if (!single) {
      ctx.fillStyle = C['sel-bg'];
      ctx.fillRect(x1, y1, x2 - x1, y2 - y1);
      const a = this.cellRectM(this.sel.ar, this.sel.ac);
      ctx.fillStyle = C.cell;
      ctx.globalAlpha = 0.65;
      ctx.fillRect(a.x + 1, a.y + 1, a.w - 2, a.h - 2);
      ctx.globalAlpha = 1;
      this.redrawCell(this.sel.ar, this.sel.ac);
    }
    ctx.strokeStyle = C.accent; ctx.lineWidth = 2;
    ctx.strokeRect(x1 + 0.5, y1 + 0.5, x2 - x1 - 1, y2 - y1 - 1);
    // điểm kéo điền
    const hs = 7 * Math.max(1, z * 0.9);
    ctx.fillStyle = C.accent;
    ctx.fillRect(x2 - hs / 2 - 1, y2 - hs / 2 - 1, hs, hs);
    ctx.strokeStyle = C.cell; ctx.lineWidth = 1.5;
    ctx.strokeRect(x2 - hs / 2 - 1, y2 - hs / 2 - 1, hs, hs);
    // mũi tên chọn từ danh sách (kiểm tra dữ liệu)
    const db = this.dvBtn();
    if (db) {
      ctx.fillStyle = C.hdr; ctx.strokeStyle = C['hdr-line']; ctx.lineWidth = 1;
      ctx.fillRect(db.x, db.y, db.w, db.h); ctx.strokeRect(db.x + 0.5, db.y + 0.5, db.w - 1, db.h - 1);
      ctx.fillStyle = C.fg;
      const cx = db.x + db.w / 2, cy = db.y + db.h / 2;
      ctx.beginPath(); ctx.moveTo(cx - 4 * z, cy - 2 * z); ctx.lineTo(cx + 4 * z, cy - 2 * z); ctx.lineTo(cx, cy + 2.5 * z); ctx.closePath(); ctx.fill();
    }
    if (this.drag?.kind === 'fill') {
      const t = this.drag.tgt;
      ctx.save(); ctx.setLineDash([4, 3]); ctx.strokeStyle = C.muted; ctx.lineWidth = 1.5;
      ctx.strokeRect(this.sx(t.c1) + 1, this.sy(t.r1) + 1, this.sx(t.c2) + this.colWpx(t.c2) - this.sx(t.c1) - 2, this.sy(t.r2) + this.rowHpx(t.r2) - this.sy(t.r1) - 2);
      ctx.restore();
    }
  }
  private redrawCell(r: number, c: number) {
    // vẽ lại ô hiện hành sau lớp phủ vùng chọn
    const e = this.eff(r, c);
    if (!e.d.text) return;
    const a = this.cellRectM(r, c);
    this.paintText(e, a.x, a.y, a.w, a.h, null, 0, a.x + a.w, r);
  }

  private drawHeaders(rows: number[], cols: number[]) {
    const ctx = this.ctx, C = this.colors, z = this.zoom, W = this.width, H = this.height;
    const hw = this.hdrW, hh = this.hdrH, g = this.range;
    const full = { cols: g.r1 === 0 && g.r2 === MAX_ROWS - 1, rows: g.c1 === 0 && g.c2 === MAX_COLS - 1 };
    ctx.fillStyle = C.hdr;
    ctx.fillRect(0, 0, W, hh);
    ctx.fillRect(0, 0, hw, H);
    ctx.font = `400 ${11 * z}px ${UIFONT}`;
    this.lastFont = '';
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    ctx.lineWidth = 1;
    // cột
    ctx.save(); ctx.beginPath(); ctx.rect(hw, 0, W - hw, hh); ctx.clip();
    for (const c of cols) {
      const x = this.sx(c), w = this.colWpx(c);
      if (w <= 0) continue;
      const on = c >= g.c1 && c <= g.c2;
      if (on) { ctx.fillStyle = full.cols ? C.accent : C['hdr-on']; ctx.fillRect(x, 0, w, hh); }
      ctx.fillStyle = on ? (full.cols ? C.cell : C['hdr-on-fg']) : C['hdr-fg'];
      ctx.fillText(colName(c), x + w / 2, hh / 2 + 0.5);
      ctx.strokeStyle = C['hdr-line'];
      ctx.beginPath(); ctx.moveTo(Math.round(x + w) - 0.5, 0); ctx.lineTo(Math.round(x + w) - 0.5, hh); ctx.stroke();
      if (on) { ctx.fillStyle = C.accent; ctx.fillRect(x, hh - 2, w, 2); }
    }
    ctx.restore();
    // hàng
    ctx.save(); ctx.beginPath(); ctx.rect(0, hh, hw, H - hh); ctx.clip();
    for (const r of rows) {
      const y = this.sy(r), h = this.rowHpx(r);
      if (h <= 0) continue;
      const on = r >= g.r1 && r <= g.r2;
      if (on) { ctx.fillStyle = full.rows ? C.accent : C['hdr-on']; ctx.fillRect(0, y, hw, h); }
      ctx.fillStyle = on ? (full.rows ? C.cell : C['hdr-on-fg']) : C['hdr-fg'];
      ctx.fillText(String(r + 1), hw / 2, y + h / 2 + 0.5);
      ctx.strokeStyle = C['hdr-line'];
      ctx.beginPath(); ctx.moveTo(0, Math.round(y + h) - 0.5); ctx.lineTo(hw, Math.round(y + h) - 0.5); ctx.stroke();
      if (on) { ctx.fillStyle = C.accent; ctx.fillRect(hw - 2, y, 2, h); }
    }
    ctx.restore();
    ctx.textAlign = 'left';
    // góc
    ctx.fillStyle = C.hdr; ctx.fillRect(0, 0, hw, hh);
    ctx.fillStyle = C['hdr-fg'];
    ctx.beginPath(); ctx.moveTo(hw - 4, hh - 12 * z); ctx.lineTo(hw - 4, hh - 4); ctx.lineTo(hw - 12 * z, hh - 4); ctx.closePath(); ctx.fill();
    ctx.strokeStyle = C['hdr-line'];
    ctx.beginPath();
    ctx.moveTo(0, hh - 0.5); ctx.lineTo(W, hh - 0.5);
    ctx.moveTo(hw - 0.5, 0); ctx.lineTo(hw - 0.5, H);
    ctx.stroke();
  }

  // ---------- chuột / cảm ứng ----------
  private local(e: { clientX: number; clientY: number }) {
    const b = this.scroller.getBoundingClientRect();
    return { x: e.clientX - b.left, y: e.clientY - b.top };
  }
  hit(x: number, y: number) {
    const hw = this.hdrW, hh = this.hdrH;
    if (x < hw && y < hh) return { zone: 'corner' as const, r: 0, c: 0 };
    if (y < hh) {
      const c = this.colAt(x);
      const right = this.sx(c) + this.colWpx(c);
      if (Math.abs(x - right) <= 4) return { zone: 'resizeC' as const, r: 0, c };
      if (Math.abs(x - this.sx(c)) <= 4 && c > 0) return { zone: 'resizeC' as const, r: 0, c: prevVisible(c, (i) => this.colWpx(i)) };
      return { zone: 'colh' as const, r: 0, c };
    }
    if (x < hw) {
      const r = this.rowAt(y);
      const bottom = this.sy(r) + this.rowHpx(r);
      if (Math.abs(y - bottom) <= 3) return { zone: 'resizeR' as const, r, c: 0 };
      if (Math.abs(y - this.sy(r)) <= 3 && r > 0) return { zone: 'resizeR' as const, r: prevVisible(r, (i) => this.rowHpx(i)), c: 0 };
      return { zone: 'rowh' as const, r, c: 0 };
    }
    const r = this.rowAt(y), c = this.colAt(x);
    const g = this.range;
    const hx = this.sx(g.c2) + this.colWpx(g.c2), hy = this.sy(g.r2) + this.rowHpx(g.r2);
    const tol = this.pointer.type === 'touch' ? 14 : 6;
    if (Math.abs(x - hx) <= tol && Math.abs(y - hy) <= tol) return { zone: 'fill' as const, r, c };
    return { zone: 'cell' as const, r, c };
  }

  private onDown(e: PointerEvent) {
    if (e.button === 2) return;
    const { x, y } = this.local(e);
    if (x > this.width || y > this.height) return; // thanh cuộn
    this.pointer = { x, y, id: e.pointerId, type: e.pointerType };
    const h = this.hit(x, y);
    const touch = e.pointerType === 'touch';
    if (touch) {
      // chạm: chờ nhấc tay mới chọn để không cản việc cuộn; giữ lâu = menu
      clearTimeout(this.longPress);
      this.longPress = window.setTimeout(() => { this.longPress = 0; this.contextAt(e.clientX, e.clientY); }, 550);
      if (h.zone === 'fill') { this.startFill(); this.scroller.setPointerCapture(e.pointerId); e.preventDefault(); }
      return;
    }
    if (this.edit.isEditing()) {
      if (this.edit.pointMode() && h.zone === 'cell') {
        e.preventDefault();
        this.drag = { kind: 'point', r0: h.r, c0: h.c };
        this.edit.insertRef(toA1(h.r, h.c), false);
        this.scroller.setPointerCapture(e.pointerId);
        return;
      }
      if (!this.edit.commit(0, 0)) return;
    }
    e.preventDefault();
    if (this.buttonAt(x, y, e.clientX, e.clientY)) return;
    this.scroller.setPointerCapture(e.pointerId);
    switch (h.zone) {
      case 'corner': this.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }); break;
      case 'resizeC': this.drag = { kind: 'resizeC', c: h.c, x0: x, w0: this.m.colW.get(h.c) }; break;
      case 'resizeR': this.drag = { kind: 'resizeR', r: h.r, y0: y, h0: this.m.rowH.get(h.r) }; break;
      case 'colh': {
        const c0 = e.shiftKey ? this.sel.ac : h.c;
        if (e.ctrlKey || e.metaKey) this.areas.push(this.range); else if (!e.shiftKey) this.areas = [];
        this.sel = { ar: 0, ac: c0, fr: MAX_ROWS - 1, fc: h.c };
        if (!e.shiftKey) this.sel.ac = h.c;
        this.drag = { kind: 'col', c0 };
        this.invalidate(); this.onSelect(this.sel);
        break;
      }
      case 'rowh': {
        const r0 = e.shiftKey ? this.sel.ar : h.r;
        if (e.ctrlKey || e.metaKey) this.areas.push(this.range); else if (!e.shiftKey) this.areas = [];
        this.sel = { ar: r0, ac: 0, fr: h.r, fc: MAX_COLS - 1 };
        this.drag = { kind: 'row', r0 };
        this.invalidate(); this.onSelect(this.sel);
        break;
      }
      case 'fill': this.startFill(); break;
      default:
        this.select(h.r, h.c, e.shiftKey || this.extMode, (e.ctrlKey || e.metaKey) && !e.shiftKey);
        this.drag = { kind: 'sel' };
        this.host.dispatchEvent(new CustomEvent('cell-down', { detail: { r: h.r, c: h.c, ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey }, bubbles: true }));
    }
    this.focusKeys();
  }
  /** Khung vùng chọn trên màn hình (toạ độ trang) */
  selectionRect() {
    const g = this.range, b = this.scroller.getBoundingClientRect();
    const x1 = this.sx(g.c1), y1 = this.sy(g.r1);
    const x2 = Math.min(this.width, this.sx(Math.min(g.c2, g.c1 + 200)) + this.colWpx(Math.min(g.c2, g.c1 + 200)));
    const y2 = Math.min(this.height, this.sy(Math.min(g.r2, g.r1 + 500)) + this.rowHpx(Math.min(g.r2, g.r1 + 500)));
    return { left: b.left + x1, top: b.top + y1, right: b.left + x2, bottom: b.top + y2 };
  }
  /** Bấm vào nút lọc hoặc nút danh sách → báo cho App mở menu */
  private buttonAt(x: number, y: number, cx: number, cy: number) {
    const db = this.dvBtn();
    if (db && x >= db.x && x <= db.x + db.w && y >= db.y && y <= db.y + db.h) {
      this.host.dispatchEvent(new CustomEvent('dv-btn', { detail: { x: cx, y: cy }, bubbles: true }));
      return true;
    }
    const f = this.m.filter;
    if (f && y > this.hdrH && x > this.hdrW) {
      const r = this.rowAt(y), c = this.colAt(x);
      if (r === f.range.r1 && c >= f.range.c1 && c <= f.range.c2) {
        const b = this.filterBtn(r, c);
        if (b && x >= b.x && x <= b.x + b.w && y >= b.y && y <= b.y + b.h) {
          this.host.dispatchEvent(new CustomEvent('filter-btn', { detail: { col: c, x: cx, y: cy }, bubbles: true }));
          return true;
        }
      }
    }
    return false;
  }
  private startFill() { this.drag = { kind: 'fill', src: this.range, tgt: this.range }; }
  focusKeys() { (this.host.querySelector('.cell-editor') as HTMLElement | null)?.focus({ preventScroll: true }); }

  private onMove(e: PointerEvent) {
    const { x, y } = this.local(e);
    if (this.longPress && (Math.abs(x - this.pointer.x) > 8 || Math.abs(y - this.pointer.y) > 8)) { clearTimeout(this.longPress); this.longPress = 0; }
    if (!this.drag) {
      if (e.pointerType === 'mouse') {
        const z = this.hit(x, y).zone;
        this.scroller.style.cursor = z === 'resizeC' ? 'col-resize' : z === 'resizeR' ? 'row-resize' : z === 'fill' ? 'crosshair' : z === 'colh' ? 's-resize' : z === 'rowh' ? 'e-resize' : 'cell';
      }
      return;
    }
    this.pointer.x = x; this.pointer.y = y;
    this.dragTo(x, y);
    this.autoScroll();
  }
  private dragTo(x: number, y: number) {
    const d = this.drag!;
    const r = this.rowAt(Math.max(this.hdrH + 1, Math.min(y, this.height - 1)));
    const c = this.colAt(Math.max(this.hdrW + 1, Math.min(x, this.width - 1)));
    if (d.kind === 'sel') { if (r !== this.sel.fr || c !== this.sel.fc) this.select(r, c, true); }
    else if (d.kind === 'col') { this.sel.fc = c; this.sel.ac = Math.min(this.sel.ac, c); this.sel.ac = d.c0; this.invalidate(); this.onSelect(this.sel); }
    else if (d.kind === 'row') { this.sel.fr = r; this.sel.ar = d.r0; this.invalidate(); this.onSelect(this.sel); }
    else if (d.kind === 'resizeC') { this.m.colW.set(d.c, Math.max(0, d.w0 + (x - d.x0) / this.zoom)); this.updateExtent(); this.invalidate(); }
    else if (d.kind === 'resizeR') { this.m.rowH.set(d.r, Math.max(0, d.h0 + (y - d.y0) / this.zoom)); this.updateExtent(); this.invalidate(); }
    else if (d.kind === 'fill') {
      const s = d.src;
      const dr = r > s.r2 ? r - s.r2 : r < s.r1 ? r - s.r1 : 0;
      const dc = c > s.c2 ? c - s.c2 : c < s.c1 ? c - s.c1 : 0;
      let t: Range = { ...s };
      if (Math.abs(dr) >= Math.abs(dc) && dr) t = dr > 0 ? { ...s, r2: r } : { ...s, r1: r };
      else if (dc) t = dc > 0 ? { ...s, c2: c } : { ...s, c1: c };
      d.tgt = t; this.invalidate();
    } else if (d.kind === 'point') {
      this.edit.insertRef(rangeToA1(normRange({ r: d.r0, c: d.c0 }, { r, c })), true);
    }
  }
  private autoScroll() {
    if (this.autoTimer) return;
    const tick = () => {
      if (!this.drag || this.drag.kind.startsWith('resize')) { this.autoTimer = 0; return; }
      const { x, y } = this.pointer, s = this.scroller;
      let dx = 0, dy = 0;
      if (y > this.height - 6) dy = 24; else if (y < this.hdrH + this.frozenH + 2 && s.scrollTop > 0) dy = -24;
      if (x > this.width - 6) dx = 40; else if (x < this.hdrW + this.frozenW + 2 && s.scrollLeft > 0) dx = -40;
      if (!dx && !dy) { this.autoTimer = 0; return; }
      s.scrollTop += dy; s.scrollLeft += dx;
      this.dragTo(x, y);
      this.autoTimer = window.setTimeout(tick, 40);
    };
    tick();
  }
  private onUp(e: PointerEvent) {
    const lpFired = e.pointerType === 'touch' && this.longPress === 0;
    clearTimeout(this.longPress);
    this.longPress = 0;
    const touch = e.pointerType === 'touch';
    const { x, y } = this.local(e);
    const d = this.drag;
    this.drag = null;
    clearTimeout(this.autoTimer); this.autoTimer = 0;
    if (touch && !d) {
      if (!lpFired && Math.abs(x - this.pointer.x) < 10 && Math.abs(y - this.pointer.y) < 10) this.tap(x, y, e);
      return;
    }
    if (!d) return;
    if (d.kind === 'sel' || d.kind === 'col' || d.kind === 'row') this.host.dispatchEvent(new CustomEvent('sel-end', { detail: { x: e.clientX, y: e.clientY, ctrl: e.ctrlKey || e.metaKey }, bubbles: true }));
    if (d.kind === 'resizeC') {
      const w = this.m.colW.get(d.c);
      this.m.colW.set(d.c, d.w0);
      this.wb.setColWidth(this.sheet, d.c, w);
    } else if (d.kind === 'resizeR') {
      const h = this.m.rowH.get(d.r);
      this.m.rowH.set(d.r, d.h0);
      this.wb.setRowHeight(this.sheet, d.r, h);
    } else if (d.kind === 'fill') {
      const { src, tgt } = d;
      if (tgt.r1 !== src.r1 || tgt.r2 !== src.r2 || tgt.c1 !== src.c1 || tgt.c2 !== src.c2) {
        try { this.wb.fill(this.sheet, src, tgt); this.selectRange(tgt); } catch (err) { this.host.dispatchEvent(new CustomEvent('grid-error', { detail: (err as Error).message, bubbles: true })); }
      }
    }
    this.invalidate();
  }
  private cancelDrag() {
    clearTimeout(this.longPress);
    const d = this.drag;
    if (d?.kind === 'resizeC') this.m.colW.set(d.c, d.w0);
    if (d?.kind === 'resizeR') this.m.rowH.set(d.r, d.h0);
    this.drag = null;
    this.updateExtent();
    this.invalidate();
  }
  private tap(x: number, y: number, e: PointerEvent) {
    const h = this.hit(x, y);
    const now = performance.now();
    if (this.edit.isEditing()) {
      if (this.edit.pointMode() && h.zone === 'cell') { this.edit.insertRef(toA1(h.r, h.c), false); return; }
      if (!this.edit.commit(0, 0)) return;
    }
    if (this.buttonAt(x, y, e.clientX, e.clientY)) return;
    if (h.zone === 'colh') { this.sel = { ar: 0, ac: h.c, fr: MAX_ROWS - 1, fc: h.c }; this.invalidate(); this.onSelect(this.sel); return; }
    if (h.zone === 'rowh') { this.sel = { ar: h.r, ac: 0, fr: h.r, fc: MAX_COLS - 1 }; this.invalidate(); this.onSelect(this.sel); return; }
    if (h.zone === 'corner') { this.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }); return; }
    if (h.zone !== 'cell' && h.zone !== 'fill') return;
    const dbl = now - this.lastTap.t < 350 && this.lastTap.r === h.r && this.lastTap.c === h.c;
    this.lastTap = { t: now, r: h.r, c: h.c };
    this.select(h.r, h.c, false);
    this.host.dispatchEvent(new CustomEvent('sel-end', { detail: { x: e.clientX, y: e.clientY, ctrl: false }, bubbles: true }));
    if (dbl) this.edit.begin(null, 'edit');
    void e;
  }
  private onDbl(e: MouseEvent) {
    const { x, y } = this.local(e);
    const h = this.hit(x, y);
    if (h.zone === 'resizeC') { this.autoFit(h.c); return; }
    if (h.zone === 'cell' && !this.edit.isEditing()) this.edit.begin(null, 'edit');
  }
  autoFit(c: number) { this.wb.setColWidth(this.sheet, c, this.fitWidth(c)); }
  /** Độ rộng vừa nội dung cột (px ở 100%) */
  fitWidth(c: number): number {
    const ctx = this.ctx, u = this.wb.used(this.sheet);
    let w = 0;
    for (let r = 0; r < Math.min(u.rows, 5000); r++) {
      if (this.m.rowH.get(r) <= 0) continue;
      const d = this.wb.display(this.sheet, r, c);
      if (!d.text || this.wb.mergeAt(this.sheet, r, c)) continue;
      const st = d.style;
      ctx.font = `${st.i ? 'italic ' : ''}${st.b ? 700 : 400} ${ptPx(st.fs)}px ${st.fn ? `"${st.fn}", ` : ''}${FONT}`;
      this.lastFont = '';
      w = Math.max(w, ctx.measureText(d.text).width + (st.ind ?? 0) * 9);
    }
    return w ? Math.ceil(w + 10) : this.m.colW.def;
  }
  private contextAt(cx: number, cy: number) {
    const { x, y } = this.local({ clientX: cx, clientY: cy });
    const h = this.hit(x, y);
    const g = this.range;
    // như Excel: chuột phải tiêu đề hàng/cột chọn cả hàng/cột (nếu chưa nằm trong vùng chọn)
    if (h.zone === 'rowh') {
      if (!(g.c1 === 0 && g.c2 === MAX_COLS - 1 && h.r >= g.r1 && h.r <= g.r2)) this.selectRange({ r1: h.r, c1: 0, r2: h.r, c2: MAX_COLS - 1 });
      this.onContext(cx, cy, 'row'); return;
    }
    if (h.zone === 'colh') {
      if (!(g.r1 === 0 && g.r2 === MAX_ROWS - 1 && h.c >= g.c1 && h.c <= g.c2)) this.selectRange({ r1: 0, c1: h.c, r2: MAX_ROWS - 1, c2: h.c });
      this.onContext(cx, cy, 'col'); return;
    }
    if (h.zone === 'cell' && !inRange(this.range, h.r, h.c)) this.select(h.r, h.c);
    this.onContext(cx, cy, 'cell');
  }
}

/** Màu chữ tự động theo nền ô (nền sáng → chữ tối) để đọc được ở cả chế độ tối */
function autoText(bg: string) {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(bg.trim());
  if (!m) return '#18231d';
  const h = m[1].length === 3 ? m[1].split('').map((x) => x + x).join('') : m[1];
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.5 ? '#18231d' : '#ffffff';
}
function clamp(v: number, a: number, b: number) { return Math.max(a, Math.min(b, v)); }
function prevVisible(i: number, size: (i: number) => number) { let j = i - 1; while (j > 0 && size(j) <= 0) j--; return j; }
