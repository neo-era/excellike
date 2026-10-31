// Giao diện Đại cương: Nhóm/Bỏ nhóm (Alt+Shift+→/←), nút +/−/cấp ở lề, Tổng phụ, Hiện/Ẩn Chi tiết
import { openDialog, h, field, select } from './dialogs';
import { MAX_COLS, MAX_ROWS, colName } from '../core/address';
import type { App } from './app';

type Ax = 'r' | 'c';
const done = (app: App) => { app.grid.layout(); app.grid.invalidate(); app.ribbon.refresh(); app.scheduleSave(); };
const run = (app: App, f: () => void) => { try { f(); done(app); } catch (e) { app.toast((e as Error).message, 'err'); } };

/** trục suy từ vùng chọn: chọn nguyên hàng → hàng, nguyên cột → cột; còn lại hỏi như Excel */
function pickAxis(app: App, title: string, then: (ax: Ax) => void) {
  const g = app.grid.range;
  if (g.c1 === 0 && g.c2 >= MAX_COLS - 1) return then('r');
  if (g.r1 === 0 && g.r2 >= MAX_ROWS - 1) return then('c');
  const rr = h('input', { type: 'radio', name: 'olax', value: 'r' }) as HTMLInputElement;
  const cc = h('input', { type: 'radio', name: 'olax', value: 'c' }) as HTMLInputElement;
  rr.checked = true;
  openDialog(title, h('div', { class: 'ol-ax' }, h('label', { class: 'chk' }, rr, ' Hàng'), h('label', { class: 'chk' }, cc, ' Cột')), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => { const ax: Ax = cc.checked ? 'c' : 'r'; setTimeout(() => then(ax), 0); } },
  ]);
  setTimeout(() => rr.focus(), 40);
}
const span = (app: App, ax: Ax) => { const g = app.grid.range; return ax === 'r' ? [g.r1, g.r2] : [g.c1, g.c2]; };

export function groupSel(app: App) {
  pickAxis(app, 'Nhóm', (ax) => run(app, () => { const [a, b] = span(app, ax); app.wb.groupLines(app.grid.sheet, ax, a, b); }));
}
export function ungroupSel(app: App) {
  pickAxis(app, 'Bỏ nhóm', (ax) => run(app, () => { const [a, b] = span(app, ax); app.wb.ungroupLines(app.grid.sheet, ax, a, b); }));
}
export function clearOutline(app: App) {
  const s = app.grid.sheet;
  if (!app.wb.outlineMax(s, 'r') && !app.wb.outlineMax(s, 'c')) { app.toast('Trang tính không có đại cương.', 'err'); return; }
  run(app, () => app.wb.group(() => { app.wb.clearOutline(s, 'r'); app.wb.clearOutline(s, 'c'); }));
}
/** nhóm chứa dòng i ở cấp sâu nhất (để Hiện/Ẩn Chi tiết) */
function groupAt(app: App, ax: Ax, i: number) {
  const s = app.grid.sheet, wb = app.wb;
  for (let L = wb.outlineMax(s, ax); L >= 1; L--) for (const g of wb.outlineRuns(s, ax, L)) if (i >= g.a && i <= g.b + 1) return { L, ...g };
  return null;
}
export function showDetail(app: App, show: boolean) {
  const s = app.grid.sheet, g = app.grid.range;
  const ax: Ax = app.wb.outlineMax(s, 'r') || !app.wb.outlineMax(s, 'c') ? 'r' : 'c';
  const at = groupAt(app, ax, ax === 'r' ? g.r1 : g.c1);
  if (!at) { app.toast('Ô đang chọn không thuộc nhóm nào.', 'err'); return; }
  run(app, () => show ? app.wb.expandGroup(s, ax, at.L, at.a, at.b) : app.wb.collapseGroup(s, ax, at.L, at.a, at.b));
}
/** bấm ở dải đại cương bên lề lưới */
export function outlineClick(app: App, kind: 'r' | 'c' | 'lr' | 'lc', L: number, at: number) {
  const s = app.grid.sheet, wb = app.wb;
  if (kind === 'lr' || kind === 'lc') {
    const ax: Ax = kind === 'lr' ? 'r' : 'c';
    if (L <= wb.outlineMax(s, ax) + 1) run(app, () => wb.showOutlineLevel(s, ax, L));
    return;
  }
  const ax: Ax = kind;
  if (L > wb.outlineMax(s, ax)) return;
  // nút +/− nằm ở hàng/cột tổng ngay sau nhóm
  const g = wb.outlineRuns(s, ax, L).find((x) => x.b + 1 === at);
  if (!g) return;
  run(app, () => wb.isCollapsed(s, ax, L, g.b) ? wb.expandGroup(s, ax, L, g.a, g.b) : wb.collapseGroup(s, ax, L, g.a, g.b));
}

const FNS: [string, string][] = [['9', 'Tổng'], ['2', 'Đếm'], ['1', 'Trung bình'], ['4', 'Lớn nhất'], ['5', 'Nhỏ nhất'], ['6', 'Tích'], ['3', 'Đếm Số'], ['7', 'Độ lệch chuẩn'], ['10', 'Phương sai']];
/** Dữ liệu ▸ Tổng phụ */
export function subtotalDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  const g0 = app.grid.range;
  const g = app.grid.isSingle() ? wb.currentRegion(s, app.grid.sel.ar, app.grid.sel.ac) : { ...g0, r2: Math.min(g0.r2, Math.max(0, wb.used(s).rows - 1)), c2: Math.min(g0.c2, Math.max(0, wb.used(s).cols - 1)) };
  if (g.r2 <= g.r1) { app.toast('Hãy chọn vùng dữ liệu có hàng tiêu đề (hoặc đặt con trỏ trong bảng).', 'err'); return; }
  const cols: [string, string][] = [];
  for (let c = g.c1; c <= g.c2; c++) { const t = wb.display(s, g.r1, c).text; cols.push([String(c), t || `Cột ${colName(c)}`]); }
  const by = select(cols, String(g.c1));
  const fn = select(FNS, '9');
  const list = h('div', { class: 'ol-cols' });
  const checks: HTMLInputElement[] = [];
  // mặc định đánh dấu cột số cuối cùng như Excel
  let lastNum = -1;
  for (let c = g.c1; c <= g.c2; c++) if (typeof wb.value(s, g.r1 + 1, c) === 'number') lastNum = c;
  for (const [v, t] of cols) {
    const cb = h('input', { type: 'checkbox', value: v }) as HTMLInputElement;
    cb.checked = +v === lastNum;
    checks.push(cb);
    list.append(h('label', { class: 'chk' }, cb, ' ' + t));
  }
  const mk = (t: string, on: boolean) => { const c = h('input', { type: 'checkbox' }) as HTMLInputElement; c.checked = on; return { c, el: h('label', { class: 'chk' }, c, ' ' + t) }; };
  const rep = mk('Thay thế tổng phụ hiện tại', true);
  const sum = h('label', { class: 'chk' }, (() => { const c = h('input', { type: 'checkbox' }) as HTMLInputElement; c.checked = true; c.disabled = true; return c; })(), ' Tóm tắt bên dưới dữ liệu');
  const body = h('div', { class: 'ol-sub' }, field('Tại mỗi thay đổi trong', by), field('Dùng hàm', fn), h('div', { class: 'fld' }, h('span', {}, 'Thêm tổng phụ vào'), list), rep.el, sum);
  openDialog('Tổng phụ', body, [
    { label: 'Loại bỏ Tất cả', act: () => { run(app, () => { wb.removeSubtotals(s, g); }); } },
    { label: 'Hủy', act: () => {} },
    {
      label: 'OK', kind: 'primary', act: () => {
        const sel = checks.filter((c) => c.checked).map((c) => +c.value);
        if (!sel.length) { app.toast('Hãy chọn ít nhất một cột để thêm tổng phụ.', 'err'); return false; }
        run(app, () => wb.subtotal(s, g, { by: +by.value, fn: +fn.value, cols: sel, replace: rep.c.checked }));
      },
    },
  ]);
  setTimeout(() => by.focus(), 40);
}
