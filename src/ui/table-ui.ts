// Giao diện Bảng: Tạo Bảng (Ctrl+T / Ctrl+L), Định dạng như Bảng, tab ngữ cảnh Thiết kế Bảng, hàm hàng tổng
import { openDialog, h, input } from './dialogs';
import { openPopover, showMenu, closeMenu } from './menu';
import { parseRange, rangeToA1 } from '../core/address';
import { TABLE_STYLES, TOTAL_FNS, tableLook, type TableDef, type TotalFn } from '../core/tables';
import type { Tab, Btn, Item, Group } from './ribbon';
import type { App } from './app';
import { removeDupDialog } from './datatools-ui';

const after = (app: App) => { app.grid.layout(); app.grid.invalidate(); app.syncSelection(); app.syncTableTab(); app.scheduleSave(); };
const absRef = (g: { r1: number; c1: number; r2: number; c2: number }) => '=' + rangeToA1(g).split(':').map((p) => p.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2')).join(':');

/** hộp Tạo Bảng: vùng dữ liệu + “Bảng có tiêu đề” */
export function createTableDialog(app: App, style = 'TableStyleMedium2') {
  const s = app.grid.sheet, wb = app.wb;
  const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac);
  if (cur) { wb.updateTable(s, cur.id, { style }); after(app); return; }
  let g = app.grid.range;
  if (app.grid.isSingle()) g = wb.currentRegion(s, app.grid.sel.ar, app.grid.sel.ac);
  const ref = input(absRef(g));
  ref.classList.add('mono');
  const hc = h('input', { type: 'checkbox' }) as HTMLInputElement;
  hc.checked = wb.guessHeader(s, g) || (g.r2 > g.r1 && [...Array(g.c2 - g.c1 + 1)].every((_, i) => typeof wb.value(s, g.r1, g.c1 + i) === 'string'));
  app.grid.selectRange(g);
  const body = h('div', { class: 'tb-new' }, h('label', { class: 'fld' }, h('span', {}, 'Dữ liệu cho bảng của bạn ở đâu?'), ref), h('label', { class: 'chk' }, hc, ' Bảng có tiêu đề'));
  openDialog('Tạo Bảng', body, [
    { label: 'Hủy', act: () => {} },
    {
      label: 'OK', kind: 'primary', act: () => {
        const r = parseRange(ref.value.replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, ''));
        if (!r) { app.toast('Tham chiếu không hợp lệ.', 'err'); return false; }
        try {
          const t = wb.createTable(s, r, hc.checked, style);
          app.grid.select(t.r1, t.c1);
          after(app);
          app.selectTab('tdesign');
        } catch (e) { app.toast((e as Error).message, 'err'); return false; }
      },
    },
  ]);
  setTimeout(() => { ref.focus(); ref.select(); }, 40);
}

/** ô mẫu kiểu bảng (vẽ bằng CSS) */
function swatch(style: string): HTMLElement {
  const L = tableLook(style)!;
  const el = h('span', { class: 'tb-sw', 'aria-hidden': 'true' });
  for (let i = 0; i < 5; i++) {
    const row = h('i');
    row.style.background = i === 0 ? (L.head.bg ?? 'transparent') : i % 2 ? (L.band ?? L.body ?? 'transparent') : (L.body ?? 'transparent');
    if (i === 0) row.style.borderBottom = `1.5px solid ${L.headLine ?? L.head.bg ?? '#000'}`;
    if (i === 0 && !L.head.bg) row.style.background = 'transparent';
    el.append(row);
  }
  return el;
}
const galCss = (st: string) => {
  const L = tableLook(st)!, hd = L.head.bg ?? '#ffffff', bd = L.band ?? '#ffffff', bo = L.body ?? '#ffffff';
  return `background:linear-gradient(${hd} 0 28%, ${L.headLine ?? hd} 28% 32%, ${bd} 32% 55%, ${bo} 55% 78%, ${bd} 78% 100%);color:transparent;font-size:0;min-width:44px;border:1px solid #c8c8c8`;
};
const famName = (st: string) => st.replace('TableStyleLight', 'Nhạt ').replace('TableStyleMedium', 'Trung bình ').replace('TableStyleDark', 'Đậm ');
/** bảng chọn kiểu bảng (Định dạng như Bảng / Kiểu Bảng) */
export function tableStylePicker(app: App, r: DOMRect, pick: (st: string) => void) {
  const box = h('div', { class: 'tb-gal', role: 'listbox', 'aria-label': 'Kiểu bảng' });
  for (const fam of ['Light', 'Medium', 'Dark']) {
    box.append(h('div', { class: 'tb-gal-h' }, fam === 'Light' ? 'Kiểu Nhạt' : fam === 'Medium' ? 'Kiểu Trung bình' : 'Kiểu Đậm'));
    const row = h('div', { class: 'tb-gal-row' });
    for (const st of TABLE_STYLES.filter((x) => x.includes(fam))) {
      const b = h('button', { class: 'tb-gal-it', title: famName(st), 'data-style': st, role: 'option' }, swatch(st));
      b.onclick = () => { closeMenu(); pick(st); };
      row.append(b);
    }
    box.append(row);
  }
  openPopover(box, r.left, r.bottom + 2);
  void app;
}

/** thả xuống ở ô hàng tổng: chọn hàm */
export function totalFnMenu(app: App, x: number, y: number): boolean {
  const s = app.grid.sheet, { ar, ac } = app.grid.sel;
  const t = app.wb.tableAt(s, ar, ac);
  if (!t || !t.totals || ar !== t.r2) return false;
  const cur = t.cols[ac - t.c1]?.fn ?? 'none';
  showMenu(TOTAL_FNS.map(([fn, label]) => ({ label, check: fn === cur, act: () => { app.safe(() => app.wb.setTotalFn(s, t.id, ac, fn as TotalFn)); after(app); } })), x, y);
  return true;
}

/** tab ngữ cảnh Thiết kế Bảng (chỉ hiện khi ô hiện hành nằm trong bảng) */
export function tableDesignTab(app: App, t: TableDef): Tab {
  const s = app.grid.sheet, wb = app.wb;
  const B = (tt: Btn['t'], id: string, icon: string, label: string, o: Partial<Btn> = {}): Btn => ({ t: tt, id, icon, label, ...o });
  const G = (id: string, label: string, icon: string, items: Item[]): Group => ({ id, label, icon, items });
  const run = (f: () => void) => { try { f(); after(app); } catch (e) { app.toast((e as Error).message, 'err'); } };
  const chk = (id: string, label: string, k: 'header' | 'totals' | 'bandR' | 'bandC' | 'first' | 'last' | 'btn', kt: string): Item => ({
    t: 'chk', id, label, kt,
    get: () => !!(wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac) as any)?.[k],
    set: (v) => run(() => { const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac); if (!cur) return; if (k === 'totals') wb.setTableTotals(s, cur.id, v); else if (k !== 'header') wb.updateTable(s, cur.id, { [k]: v }); }),
  });
  const nameItem: Item = {
    t: 'combo', id: 'td-name', label: 'Tên Bảng', w: 120, kt: 'A',
    get: () => wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac)?.name ?? '',
    set: (v) => run(() => { const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac); if (cur && v.trim() && v.trim() !== cur.name) wb.updateTable(s, cur.id, { name: v.trim() }); }),
    list: () => {},
  };
  return { id: 'tdesign', label: 'Thiết kế Bảng', kt: 'JT', groups: [
    G('tprops', 'Thuộc tính', 'fmtTable', [
      { t: 'col', items: [nameItem, B('sm', 'td-resize', 'autofit', 'Đổi kích thước Bảng', { kt: 'E', act: () => resizeDialog(app) })] },
    ]),
    G('ttools', 'Công cụ', 'removeDup', [
      { t: 'col', items: [
        B('sm', 'td-pivot', 'pivot', 'Tóm tắt bằng PivotTable', { kt: 'P', act: () => { const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac); if (cur) app.grid.selectRange({ r1: cur.r1, c1: cur.c1, r2: cur.r2 - (cur.totals ? 1 : 0), c2: cur.c2 }); app.insertPivot(); } }),
        B('sm', 'td-dup', 'removeDup', 'Loại bỏ Trùng lặp', { kt: 'M', act: () => { const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac); if (cur) app.grid.selectRange({ r1: cur.r1, c1: cur.c1, r2: cur.r2 - (cur.totals ? 1 : 0), c2: cur.c2 }); removeDupDialog(app); } }),
        B('sm', 'td-conv', 'table', 'Chuyển thành Vùng', { kt: 'G', act: () => { const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac); if (cur && confirm('Bạn có muốn chuyển bảng thành vùng thông thường không?')) run(() => wb.convertTableToRange(s, cur.id)); } }),
      ] },
    ]),
    G('topts', 'Tùy chọn Kiểu Bảng', 'check', [
      { t: 'col', items: [chk('td-hdr', 'Hàng Tiêu đề', 'header', 'H'), chk('td-tot', 'Hàng Tổng', 'totals', 'T'), chk('td-bandr', 'Hàng Dải màu', 'bandR', 'R')] },
      { t: 'col', items: [chk('td-first', 'Cột Đầu tiên', 'first', 'M'), chk('td-last', 'Cột Cuối cùng', 'last', 'L'), chk('td-bandc', 'Cột Dải màu', 'bandC', 'B')] },
      { t: 'col', items: [chk('td-btn', 'Nút Lọc', 'btn', 'W')] },
    ]),
    G('tstyles', 'Kiểu Bảng', 'fmtTable', [
      { t: 'gal', id: 'td-styles', label: 'Kiểu Nhanh', kt: 'S',
        items: ['TableStyleLight1', 'TableStyleLight2', 'TableStyleMedium2', 'TableStyleMedium3', 'TableStyleMedium7', 'TableStyleDark2'].map((st) => ({ label: famName(st), css: galCss(st), act: () => run(() => { const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac); if (cur) wb.updateTable(s, cur.id, { style: st }); }) })),
        act: () => {},
        more: (r) => tableStylePicker(app, r, (st) => run(() => { const cur = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac); if (cur) wb.updateTable(s, cur.id, { style: st }); })) },
    ]),
  ] };
  void t;
}

function resizeDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb, t = wb.tableAt(s, app.grid.sel.ar, app.grid.sel.ac);
  if (!t) return;
  const ref = input(absRef(t));
  ref.classList.add('mono');
  openDialog('Đổi kích thước Bảng', h('div', { class: 'tb-new' }, h('label', { class: 'fld' }, h('span', {}, 'Chọn vùng dữ liệu mới cho bảng'), ref), h('p', { class: 'muted' }, 'Lưu ý: tiêu đề phải ở cùng hàng, vùng mới phải chồng lên bảng hiện tại.')), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      const r = parseRange(ref.value.replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, ''));
      if (!r) { app.toast('Tham chiếu không hợp lệ.', 'err'); return false; }
      try { wb.resizeTable(s, t.id, r); after(app); } catch (e) { app.toast((e as Error).message, 'err'); return false; }
    } },
  ]);
  setTimeout(() => { ref.focus(); ref.select(); }, 40);
}
