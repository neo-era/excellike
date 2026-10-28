// Lớp biểu đồ nổi trên lưới: vẽ, chọn, kéo, đổi cỡ, sửa, xóa + hộp thoại biểu đồ
import { parseRef, rangeToA1, type Range } from '../core/address';
import { quoteSheet } from '../core/formula';
import { CHART_NAMES, chartData, chartSourceRange, seriesFromRange, type Chart, type ChartType, type Ref } from '../core/chart';
import type { Workbook } from '../core/workbook';
import { drawChart, type ChartTheme } from './chartdraw';
import { field, h, input, openDialog, select } from './dialogs';
import type { Grid } from './grid';
import { ic } from './icons';

const MIN_W = 160, MIN_H = 110;
type Box = { el: HTMLDivElement; cv: HTMLCanvasElement; key: string; ch: Chart | null };

export function themeNow(): ChartTheme {
  const cs = getComputedStyle(document.documentElement);
  const v = (k: string, d: string) => cs.getPropertyValue(k).trim() || d;
  return { bg: v('--cell', '#fff'), fg: v('--fg', '#1f2a24'), muted: v('--muted', '#5b6a62'), grid: v('--line', '#e3e8e5'), font: "'Be Vietnam Pro', Arial, sans-serif" };
}

/** "A1:C5" hoặc "'Sheet 1'!A1:C5" → Ref */
export function parseSheetRef(wb: Workbook, text: string, sheet: number): Ref | null {
  const t = text.trim().replace(/^=/, '');
  const i = t.lastIndexOf('!');
  let s = sheet, rest = t;
  if (i >= 0) {
    const name = t.slice(0, i).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
    const id = wb.sheetIds().find((x) => wb.sheetName(x).toLowerCase() === name.toLowerCase());
    if (id === undefined) return null;
    s = id; rest = t.slice(i + 1);
  }
  const g = parseRef(rest);
  return g ? { sheet: s, range: g } : null;
}
export const refText = (wb: Workbook, ref: Ref, cur: number) => (ref.sheet === cur ? '' : quoteSheet(wb.sheetName(ref.sheet)) + '!') + rangeToA1(ref.range);

export class ChartLayer {
  readonly el: HTMLDivElement;
  private boxes = new Map<string, Box>();
  sel: string | null = null;
  private theme = themeNow();

  constructor(host: HTMLElement, private wb: Workbook, private grid: Grid, private onEdit: (id: string) => void, private onMsg: (m: string) => void) {
    this.el = document.createElement('div');
    this.el.className = 'chart-layer';
    host.appendChild(this.el);
    grid.scroller.addEventListener('pointerdown', () => this.select(null), true);
    document.addEventListener('keydown', (e) => {
      if (!this.sel) return;
      const t = e.target as HTMLElement | null;
      if (t && (/^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName) || t.isContentEditable)) return;
      if (document.querySelector('.dlg-back, .menu')) return;
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); e.stopPropagation(); this.remove(this.sel); }
      else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.select(null); grid.focusKeys(); }
      else if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); this.onEdit(this.sel); }
      else if (e.key.startsWith('Arrow')) {
        e.preventDefault(); e.stopPropagation();
        const step = e.shiftKey ? 1 : 10;
        const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0, dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
        const b = this.boxes.get(this.sel);
        if (b) this.commitMove(this.sel, parseFloat(b.el.style.left) + dx, parseFloat(b.el.style.top) + dy);
      }
    }, true);
    new MutationObserver(() => { this.theme = themeNow(); for (const b of this.boxes.values()) b.key = ''; this.sync(); })
      .observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { this.theme = themeNow(); for (const b of this.boxes.values()) b.key = ''; this.sync(); });
  }

  private charts() { return this.wb.meta(this.grid.sheet).charts; }
  remove(id: string) {
    this.wb.removeChart(this.grid.sheet, id);
    if (this.sel === id) this.sel = null;
    this.grid.focusKeys();
    this.onMsg('Đã xóa biểu đồ (Ctrl+Z để hoàn tác).');
  }
  select(id: string | null) {
    if (this.sel === id) return;
    this.sel = id;
    for (const [k, b] of this.boxes) b.el.classList.toggle('on', k === id);
    if (id) (document.activeElement as HTMLElement | null)?.blur?.();
  }

  /** Gọi sau mỗi lần lưới vẽ lại */
  sync() {
    const g = this.grid, z = g.zoom;
    this.el.style.left = g.hdrW + 'px';
    this.el.style.top = g.hdrH + 'px';
    const seen = new Set<string>();
    const dpr = window.devicePixelRatio || 1;
    for (const ch of this.charts()) {
      const id = ch.id!;
      seen.add(id);
      let b = this.boxes.get(id);
      if (!b) { b = this.make(id); this.boxes.set(id, b); }
      if (b.el.dataset.drag) continue;
      const x = g.sx(ch.anchor.c) + ch.anchor.dx * z - g.hdrW, y = g.sy(ch.anchor.r) + ch.anchor.dy * z - g.hdrH;
      const w = ch.w * z, hh = ch.h * z;
      Object.assign(b.el.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: hh + 'px' });
      b.el.classList.toggle('on', this.sel === id);
      // ngoài màn hình → không vẽ lại
      if (x > g.width || y > g.height || x + w < 0 || y + hh < 0) continue;
      const key = `${this.wb.ver}|${z}|${dpr}`;
      if (b.ch !== ch || b.key !== key) {
        b.ch = ch; b.key = key;
        b.cv.width = Math.max(1, Math.round(w * dpr)); b.cv.height = Math.max(1, Math.round(hh * dpr));
        const ctx = b.cv.getContext('2d')!;
        ctx.setTransform(dpr * z, 0, 0, dpr * z, 0, 0);
        drawChart(ctx, ch.w, ch.h, ch, chartData(this.wb, ch), this.theme);
        b.el.setAttribute('aria-label', `Biểu đồ ${CHART_NAMES[ch.type]}${ch.title ? ': ' + ch.title : ''}`);
      }
    }
    for (const [id, b] of this.boxes) if (!seen.has(id)) { b.el.remove(); this.boxes.delete(id); if (this.sel === id) this.sel = null; }
  }

  private make(id: string): Box {
    const el = document.createElement('div');
    el.className = 'chart-box';
    el.dataset.id = id;
    el.tabIndex = -1;
    el.setAttribute('role', 'img');
    const cv = document.createElement('canvas');
    el.appendChild(cv);
    const bar = document.createElement('div');
    bar.className = 'chart-bar';
    bar.innerHTML = `<button data-a="edit" title="Sửa biểu đồ">${ic('pencil', 16)}<span>Sửa</span></button><button data-a="del" title="Xóa biểu đồ (Delete)">${ic('trash', 16)}<span>Xóa</span></button>`;
    bar.addEventListener('pointerdown', (e) => e.stopPropagation());
    bar.addEventListener('click', (e) => {
      const a = (e.target as HTMLElement).closest('button')?.dataset.a;
      if (a === 'edit') this.onEdit(id); else if (a === 'del') this.remove(id);
    });
    el.appendChild(bar);
    for (const k of ['e', 's', 'se', 'w', 'n', 'nw', 'ne', 'sw']) { const hd = document.createElement('i'); hd.className = 'chart-h h-' + k; hd.dataset.h = k; el.appendChild(hd); }
    el.addEventListener('pointerdown', (e) => this.down(e, id, el));
    el.addEventListener('dblclick', () => this.onEdit(id));
    el.addEventListener('wheel', (e) => { this.grid.scroller.scrollBy(e.deltaX, e.deltaY); }, { passive: true });
    el.addEventListener('contextmenu', (e) => { e.preventDefault(); this.select(id); });
    this.el.appendChild(el);
    return { el, cv, key: '', ch: null };
  }

  private down(e: PointerEvent, id: string, el: HTMLDivElement) {
    if (e.button !== 0) { this.select(id); return; }
    e.preventDefault();
    this.select(id);
    const handle = (e.target as HTMLElement).dataset.h ?? '';
    const x0 = e.clientX, y0 = e.clientY;
    const L = parseFloat(el.style.left), T = parseFloat(el.style.top), W = parseFloat(el.style.width), H = parseFloat(el.style.height);
    const z = this.grid.zoom;
    let moved = false;
    el.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      const dx = ev.clientX - x0, dy = ev.clientY - y0;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 4) return;
      moved = true;
      el.dataset.drag = '1';
      if (!handle) { el.style.left = L + dx + 'px'; el.style.top = T + dy + 'px'; return; }
      let l = L, t = T, w = W, hh = H;
      if (handle.includes('e')) w = Math.max(MIN_W * z, W + dx);
      if (handle.includes('s')) hh = Math.max(MIN_H * z, H + dy);
      if (handle.includes('w')) { w = Math.max(MIN_W * z, W - dx); l = L + W - w; }
      if (handle.includes('n')) { hh = Math.max(MIN_H * z, H - dy); t = T + H - hh; }
      Object.assign(el.style, { left: l + 'px', top: t + 'px', width: w + 'px', height: hh + 'px' });
    };
    const up = () => {
      el.removeEventListener('pointermove', move);
      el.removeEventListener('pointerup', up);
      el.removeEventListener('pointercancel', up);
      delete el.dataset.drag;
      if (!moved) return;
      const l = parseFloat(el.style.left), t = parseFloat(el.style.top);
      if (handle) {
        const w = Math.round(parseFloat(el.style.width) / z), hh = Math.round(parseFloat(el.style.height) / z);
        this.commitMove(id, l, t, { w, h: hh });
      } else this.commitMove(id, l, t);
    };
    el.addEventListener('pointermove', move);
    el.addEventListener('pointerup', up);
    el.addEventListener('pointercancel', up);
  }

  /** Tọa độ trong lớp (px màn hình) → neo ô + lệch */
  private commitMove(id: string, l: number, t: number, size?: { w: number; h: number }) {
    const g = this.grid, z = g.zoom;
    const x = Math.max(g.hdrW, l + g.hdrW), y = Math.max(g.hdrH, t + g.hdrH);
    const c = g.colAt(x + 0.01), r = g.rowAt(y + 0.01);
    const dx = Math.max(0, Math.round((x - g.sx(c)) / z)), dy = Math.max(0, Math.round((y - g.sy(r)) / z));
    this.wb.updateChart(g.sheet, id, { anchor: { r, c, dx, dy }, ...(size ?? {}) });
    this.select(id);
  }
}

// ------------------------- Hộp thoại biểu đồ -------------------------
const TYPES: ChartType[] = ['col', 'colStack', 'bar', 'line', 'area', 'pie', 'doughnut'];
const TYPE_IC: Record<ChartType, string> = { col: 'chart', colStack: 'chartStack', bar: 'chartBar', line: 'chartLine', area: 'chartArea', pie: 'chartPie', doughnut: 'chartDonut' };

export interface ChartDlgOpts { wb: Workbook; sheet: number; range?: Range; chart?: Chart; type?: ChartType; onDone: (ch: Chart | null, err?: string) => void }

export function chartDialog(o: ChartDlgOpts) {
  const { wb, sheet } = o;
  const edit = o.chart;
  let type: ChartType = edit?.type ?? o.type ?? 'col';
  const srcRef = edit ? chartSourceRange(edit) : o.range ? { sheet, range: o.range } : null;
  const srcInit = srcRef ? refText(wb, srcRef, sheet) : '';
  const body = h('div', { class: 'stack chart-dlg' });
  const types = h('div', { class: 'chart-types', role: 'radiogroup', 'aria-label': 'Loại biểu đồ' });
  const typeBtns = TYPES.map((t) => {
    const b = h('button', { type: 'button', role: 'radio', 'data-t': t, title: CHART_NAMES[t] });
    b.innerHTML = ic(TYPE_IC[t], 22) + `<span>${CHART_NAMES[t]}</span>`;
    b.onclick = () => { type = t; mark(); preview(); };
    types.append(b);
    return b;
  });
  const mark = () => typeBtns.forEach((b) => { const on = b.dataset.t === type; b.classList.toggle('on', on); b.setAttribute('aria-checked', String(on)); });
  const title = input(edit?.title ?? '', 'Tiêu đề biểu đồ');
  const src = input(srcInit, 'ví dụ: A1:D8');
  const orient = select([['auto', 'Tự động'], ['col', 'Mỗi cột một chuỗi'], ['row', 'Mỗi hàng một chuỗi']], 'auto');
  const legend = select([['b', 'Dưới'], ['r', 'Phải'], ['t', 'Trên'], ['none', 'Ẩn']], edit?.legend ?? 'b');
  const labels = h('input', { type: 'checkbox' }) as HTMLInputElement;
  labels.checked = edit ? !!edit.labels : false;
  const err = h('p', { class: 'err-text', role: 'alert' });
  const cv = h('canvas', { class: 'chart-prev', width: '720', height: '400', 'aria-label': 'Xem trước biểu đồ' }) as HTMLCanvasElement;
  body.append(field('Loại biểu đồ', types), h('div', { class: 'row2' }, field('Tiêu đề', title), field('Vùng dữ liệu', src)),
    h('div', { class: 'row2' }, field('Chuỗi dữ liệu', orient), field('Chú giải', legend)),
    h('label', { class: 'chk' }, labels, ' Hiện nhãn giá trị'), err, cv);
  mark();

  const build = (): Chart | null => {
    err.textContent = '';
    let series = edit?.series ?? [];
    const keep = edit && src.value.trim() === srcInit && (orient as HTMLSelectElement).value === 'auto';
    if (!keep) {
      const ref = parseSheetRef(wb, src.value, sheet);
      if (!ref) { err.textContent = 'Vùng dữ liệu không hợp lệ. Ví dụ đúng: A1:D8 hoặc Sheet2!B2:C10'; return null; }
      const ov = (orient as HTMLSelectElement).value;
      series = seriesFromRange(wb, ref.sheet, ref.range, ov === 'auto' ? undefined : ov === 'row');
      if (!series.length) { err.textContent = 'Vùng đã chọn không có dữ liệu để vẽ.'; return null; }
    }
    return {
      ...(edit ?? { anchor: { r: 0, c: 0, dx: 0, dy: 0 }, w: 480, h: 288 }),
      type, title: title.value.trim() || undefined, series,
      legend: (legend as HTMLSelectElement).value as Chart['legend'], labels: labels.checked,
    };
  };
  const preview = () => {
    const ch = build();
    const ctx = cv.getContext('2d')!;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, cv.width, cv.height);
    if (!ch) return;
    ctx.setTransform(2, 0, 0, 2, 0, 0);
    drawChart(ctx, 360, 200, ch, chartData(wb, ch), themeNow());
  };
  for (const el of [title, src]) el.addEventListener('input', preview);
  for (const el of [orient, legend, labels]) el.addEventListener('change', preview);
  const btns = [{ label: 'Hủy', act: () => { o.onDone(null); } }];
  openDialog(edit ? 'Sửa biểu đồ' : 'Chèn biểu đồ', body, [
    ...btns,
    { label: edit ? 'Cập nhật' : 'Chèn', kind: 'primary' as const, act: () => { const ch = build(); if (!ch) return false; o.onDone(ch); } },
  ]);
  setTimeout(preview, 0);
}
