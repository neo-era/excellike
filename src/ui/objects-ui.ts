// Đối tượng (0.9): Ảnh, Hình dạng, Biểu tượng, Hộp Văn bản, tab Định dạng Hình, Sắp xếp lớp, Ngăn Chọn, Ký hiệu, Hộp kiểm trong ô
import { openDialog, closeDialog, h, input } from './dialogs';
import { showMenu, type MenuItem } from './menu';
import { SHAPE_NAMES, CHART_NAMES, isObj, type Chart, type ShapeKind } from '../core/chart';
import { P as ICONS, ic } from './icons';
import { MAX_COLS } from '../core/address';
import type { Tab, Btn, Item, Group } from './ribbon';
import type { App } from './app';

/** vẽ ảnh (SVG/WebP) ra PNG */
function raster(im: HTMLImageElement, w: number, h: number) {
  const c = document.createElement('canvas'); c.width = Math.min(2000, w); c.height = Math.min(2000, h);
  c.getContext('2d')!.drawImage(im, 0, 0, c.width, c.height);
  return c.toDataURL('image/png');
}
const placeAt = (app: App) => ({ r: app.grid.sel.ar, c: Math.min(app.grid.sel.ac, MAX_COLS - 4), dx: 0, dy: 0 });
function add(app: App, o: Partial<Chart> & Pick<Chart, 'type' | 'w' | 'h'>) {
  const s = app.grid.sheet;
  app.safe(() => {
    const id = app.wb.addChart(s, { series: [], anchor: placeAt(app), legend: 'none', ...o } as Chart);
    app.grid.invalidate();
    requestAnimationFrame(() => app.charts.select(id));
  });
}
/** Chèn ▸ Ảnh: chọn tệp ảnh trên máy (PNG, JPEG, GIF, SVG, WebP; tối đa 5 MB) */
export function insertPicture(app: App) {
  const fi = h('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif,image/svg+xml,image/webp' }) as HTMLInputElement;
  fi.onchange = () => {
    const f = fi.files?.[0];
    if (!f) return;
    if (f.size > 5_000_000) { app.toast('Ảnh quá lớn (tối đa 5 MB).', 'err'); return; }
    const rd = new FileReader();
    rd.onload = () => {
      const src = String(rd.result), im = new Image();
      im.onload = () => {
        const k = Math.min(1, 480 / im.naturalWidth, 360 / im.naturalHeight);
        add(app, { type: 'pic', img: /^data:image\/(svg|webp)/.test(src) ? raster(im, im.naturalWidth || 300, im.naturalHeight || 300) : src, w: Math.max(24, Math.round(im.naturalWidth * k)), h: Math.max(24, Math.round(im.naturalHeight * k)), alt: f.name.replace(/\.[^.]+$/, '') });
      };
      im.onerror = () => app.toast('Không đọc được ảnh này.', 'err');
      im.src = src;
    };
    rd.readAsDataURL(f);
  };
  fi.click();
}
export function shapeMenu(app: App, r: DOMRect) {
  showMenu((Object.keys(SHAPE_NAMES) as ShapeKind[]).map((k) => ({ label: SHAPE_NAMES[k], act: () => add(app, { type: 'shape', shape: k, w: k === 'line' ? 160 : 140, h: k === 'line' ? 80 : 90 }) })), r.left, r.bottom + 2);
}
export function insertTextbox(app: App) {
  add(app, { type: 'text', text: 'Nhập văn bản', w: 200, h: 60, fs: 11 });
  setTimeout(() => { const id = app.charts.sel; if (id) editObject(app, id); }, 60);
}
/** Biểu tượng: chèn một biểu tượng nét (SVG) làm ảnh */
export function iconsDialog(app: App) {
  const names = Object.keys(ICONS).filter((n) => !/^(chev|x$)/.test(n));
  const q = input('', 'Tìm biểu tượng…');
  const gridEl = h('div', { class: 'ic-grid', role: 'listbox', 'aria-label': 'Biểu tượng' });
  let pick = '';
  const draw = () => {
    gridEl.replaceChildren(...names.filter((n) => !q.value || n.toLowerCase().includes(q.value.toLowerCase())).map((n) => {
      const b = h('button', { class: 'ic-it' + (n === pick ? ' on' : ''), title: n, role: 'option' });
      b.innerHTML = ic(n, 28);
      b.onclick = () => { pick = n; draw(); };
      b.ondblclick = () => { pick = n; ok(); closeDialog(); };
      return b;
    }));
  };
  const ok = () => {
    if (!pick) return false;
    const svg = ic(pick, 192).replace('stroke="currentColor"', 'stroke="#1f3864"').replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" ');
    const im = new Image(), name = pick;
    // ảnh PNG (Excel cần ảnh raster khi mở tệp)
    im.onload = () => add(app, { type: 'pic', img: raster(im, 192, 192), w: 96, h: 96, alt: name });
    im.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)));
  };
  q.oninput = draw; draw();
  openDialog('Chèn Biểu tượng', h('div', { class: 'ic-dlg' }, q, gridEl), [{ label: 'Hủy', act: () => {} }, { label: 'Chèn', kind: 'primary', act: () => ok() }]);
}
/** nhấp đúp đối tượng: sửa chữ (hình/hộp văn bản) hoặc văn bản thay thế (ảnh) */
export function editObject(app: App, id: string) {
  const s = app.grid.sheet, o = app.wb.meta(s).charts.find((c) => c.id === id);
  if (!o) return;
  if (o.type === 'pic') { altDialog(app, o); return; }
  const ta = h('textarea', { rows: '4', class: 'obj-ta' }) as HTMLTextAreaElement;
  ta.value = o.text ?? '';
  openDialog(o.type === 'text' ? 'Sửa Hộp Văn bản' : 'Sửa Văn bản trong Hình', h('div', { class: 'an-form' }, ta), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => { app.safe(() => app.wb.updateChart(s, id, { text: ta.value })); app.grid.invalidate(); } },
  ]);
  setTimeout(() => { ta.focus(); ta.select(); }, 40);
}
function altDialog(app: App, o: Chart) {
  const ta = h('textarea', { rows: '3' }) as HTMLTextAreaElement;
  ta.value = o.alt ?? '';
  openDialog('Văn bản Thay thế', h('div', { class: 'an-form' }, h('p', { class: 'muted' }, 'Mô tả nội dung ảnh cho người dùng trình đọc màn hình.'), ta), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => { app.safe(() => app.wb.updateChart(app.grid.sheet, o.id!, { alt: ta.value.trim() || undefined })); } },
  ]);
}
const FILLS: [string, string][] = [['', 'Mặc định'], ['none', 'Không Tô'], ['#ffffff', 'Trắng'], ['#4472c4', 'Xanh dương'], ['#70ad47', 'Xanh lá'], ['#ed7d31', 'Cam'], ['#ffc000', 'Vàng'], ['#c00000', 'Đỏ'], ['#7f7f7f', 'Xám'], ['#000000', 'Đen']];
/** Ngăn Chọn: danh sách đối tượng trên trang để chọn */
export function selectionPane(app: App, r?: DOMRect) {
  const list = app.wb.meta(app.grid.sheet).charts;
  if (!list.length) { app.toast('Trang tính không có đối tượng nổi.'); return; }
  showMenu([...list].reverse().map((c, i) => ({ label: `${isObj(c.type) ? CHART_NAMES[c.type] : 'Biểu đồ'} ${list.length - i}${c.title || c.alt || c.text ? ' — ' + (c.title || c.alt || c.text)!.slice(0, 30) : ''}`, check: c.id === app.charts.sel, act: () => app.charts.select(c.id!) })), r?.left ?? innerWidth / 2, (r?.bottom ?? 160) + 2);
}
export function arrange(app: App, how: 'front' | 'back' | 'fwd' | 'bwd') {
  const id = app.charts.sel;
  if (!id) { app.toast('Hãy chọn một ảnh, hình hoặc biểu đồ trước.', 'err'); return; }
  app.safe(() => app.wb.orderChart(app.grid.sheet, id, how));
  app.grid.invalidate();
}
export const arrangeMenu = (app: App, fwd: boolean) => (r: DOMRect) => showMenu(fwd
  ? [{ label: 'Đưa Lên trên', act: () => arrange(app, 'fwd') }, { label: 'Đưa lên Trên cùng', act: () => arrange(app, 'front') }]
  : [{ label: 'Đưa Xuống dưới', act: () => arrange(app, 'bwd') }, { label: 'Đưa xuống Dưới cùng', act: () => arrange(app, 'back') }], r.left, r.bottom + 2);

/** tab ngữ cảnh Định dạng Hình / Định dạng Ảnh */
export function objectTab(app: App, o: Chart): Tab {
  const s = app.grid.sheet, wb = app.wb;
  const cur = () => wb.meta(s).charts.find((c) => c.id === app.charts.sel);
  const upd = (p: Partial<Chart>) => { const c = cur(); if (!c?.id) return; app.safe(() => wb.updateChart(s, c.id!, p)); app.grid.invalidate(); app.ribbon.refresh(); };
  const B = (t: Btn['t'], id: string, icon: string, label: string, opt: Partial<Btn> = {}): Btn => ({ t, id, icon, label, ...opt });
  const G = (id: string, label: string, icon: string, items: Item[]): Group => ({ id, label, icon, items });
  const menuAt = (f: () => MenuItem[]) => (r: DOMRect) => showMenu(f(), r.left, r.bottom + 2);
  const cm = (px: number) => String(Math.round((px / 96) * 2.54 * 100) / 100).replace('.', ',');
  const fromCm = (t: string) => { const n = parseFloat(t.replace(',', '.')); return isFinite(n) && n > 0.2 ? Math.round((n / 2.54) * 96) : null; };
  const pic = o.type === 'pic';
  const groups: Group[] = [];
  if (!pic) groups.push(G('oshape', 'Kiểu Hình', 'fill', [
    B('big', 'of-fill', 'fill', 'Tô Hình', { kt: 'SF', drop: menuAt(() => FILLS.map(([v, n]) => ({ label: n, check: (cur()?.fill ?? '') === v, act: () => upd({ fill: v || undefined }) })) as MenuItem[]) }),
    B('big', 'of-line', 'borders', 'Viền Hình', { kt: 'SO', drop: menuAt(() => FILLS.map(([v, n]) => ({ label: v === 'none' ? 'Không Viền' : n, check: (cur()?.border ?? '') === v, act: () => upd({ border: v || undefined }) })) as MenuItem[]) }),
    ...(o.type !== 'shape' || o.shape !== 'line' ? [B('big', 'of-text', 'textbox', 'Sửa Văn bản', { kt: 'X', act: () => { const c = cur(); if (c?.id) editObject(app, c.id); } }),
      B('big', 'of-fg', 'font', 'Màu Chữ', { kt: 'TF', drop: menuAt(() => [['', 'Mặc định'], ['#000000', 'Đen'], ['#ffffff', 'Trắng'], ['#1f3864', 'Xanh đậm'], ['#c00000', 'Đỏ']].map(([v, n]) => ({ label: n, check: (cur()?.fg ?? '') === v, act: () => upd({ fg: v || undefined }) })) as MenuItem[]) }),
      { t: 'combo', id: 'of-fs', label: 'Cỡ chữ', w: 50, kt: 'FS', get: () => String(cur()?.fs ?? 11), set: (v) => { const n = parseFloat(v.replace(',', '.')); if (n >= 6 && n <= 96) upd({ fs: n }); }, list: (r) => showMenu([8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((n) => ({ label: String(n), act: () => upd({ fs: n }) })), r.left, r.bottom) } as Item] : []),
  ]));
  else groups.push(G('opic', 'Kiểu Ảnh', 'picture', [
    B('big', 'op-border', 'borders', 'Viền Ảnh', { kt: 'PB', drop: menuAt(() => [['', 'Không viền'], ['#000000', 'Đen'], ['#7f7f7f', 'Xám'], ['#4472c4', 'Xanh dương']].map(([v, n]) => ({ label: n, check: (cur()?.border ?? '') === v, act: () => upd({ border: v || undefined }) })) as MenuItem[]) }),
    B('big', 'op-alt', 'access', 'Văn bản Thay thế', { kt: 'AT', act: () => { const c = cur(); if (c) altDialog(app, c); } }),
    B('big', 'op-reset', 'refresh', 'Đặt lại Kích thước', { kt: 'R', act: () => { const c = cur(); if (!c?.img) return; const im = new Image(); im.onload = () => { const k = Math.min(1, 480 / im.naturalWidth, 360 / im.naturalHeight); upd({ w: Math.round(im.naturalWidth * k), h: Math.round(im.naturalHeight * k) }); }; im.src = c.img; } }),
  ]));
  groups.push(G('oarr', 'Sắp xếp', 'arrange', [
    B('big', 'of-fwd', 'arrange', 'Đưa Lên trên', { kt: 'AF', split: true, act: () => arrange(app, 'fwd'), drop: arrangeMenu(app, true) }),
    B('big', 'of-back', 'arrange', 'Đưa Xuống dưới', { kt: 'AE', split: true, act: () => arrange(app, 'bwd'), drop: arrangeMenu(app, false) }),
    B('big', 'of-pane', 'slicer', 'Ngăn Chọn', { kt: 'AP', drop: (r) => selectionPane(app, r) }),
  ]));
  groups.push(G('osize', 'Kích thước', 'autofit', [{ t: 'col', items: [
    { t: 'combo', id: 'of-h', label: 'Chiều cao (cm)', w: 70, kt: 'HH', get: () => (cur() ? cm(cur()!.h) : ''), set: (v) => { const n = fromCm(v); if (n) upd({ h: Math.min(n, 4000) }); }, list: () => {} },
    { t: 'combo', id: 'of-w', label: 'Chiều rộng (cm)', w: 70, kt: 'HW', get: () => (cur() ? cm(cur()!.w) : ''), set: (v) => { const n = fromCm(v); if (n) upd({ w: Math.min(n, 4000) }); }, list: () => {} },
  ] }]));
  return { id: 'oformat', label: pic ? 'Định dạng Ảnh' : 'Định dạng Hình', kt: 'JP', groups };
}

/** Chèn ▸ Ký hiệu */
const SYMBOLS = '±×÷≤≥≠≈∞√∑∏∆∂∫°′″‰∅∈∩∪⊂⊃∀∃¬∧∨→←↑↓↔⇒⇔•·…–—‘’“”«»§¶©®™†‡№℃℉µΩαβγδεθλπσφψω½¼¾⅓⅔¹²³⁴ⁿ₀₁₂₃₫€$£¥¢₩₹✓✗★☆●○■□▲△▼▽◆◇♠♣♥♦☎✉⌀⌘'.split('');
export function symbolDialog(app: App) {
  let pick = '';
  const prev = h('div', { class: 'sym-prev', 'aria-live': 'polite' });
  const gridEl = h('div', { class: 'sym-grid', 'data-noi18n': '', role: 'listbox', 'aria-label': 'Ký hiệu' });
  const insert = (ch: string) => {
    if (app.ed.isEditing()) { app.ed.refocus(); app.ed.insertText(ch); return; }
    app.ed.begin(app.wb.rawText(app.grid.sheet, app.grid.sel.ar, app.grid.sel.ac) + ch, 'edit');
  };
  for (const ch of SYMBOLS) {
    const b = h('button', { class: 'sym-it', role: 'option', title: 'U+' + ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0') }, ch);
    b.onclick = () => { pick = ch; gridEl.querySelectorAll('.on').forEach((x) => x.classList.remove('on')); b.classList.add('on'); prev.textContent = `${ch}  ${b.title}`; };
    b.ondblclick = () => { closeDialog(); insert(ch); };
    gridEl.append(b);
  }
  openDialog('Ký hiệu', h('div', { class: 'sym-dlg' }, gridEl, prev), [{ label: 'Đóng', act: () => {} }, { label: 'Chèn', kind: 'primary', act: () => { if (!pick) return false; insert(pick); } }]);
}

/** Chèn ▸ Hộp kiểm: ô chứa TRUE/FALSE hiện thành hộp kiểm, bấm (hoặc Space) để đổi */
export function insertCheckbox(app: App) {
  const s = app.grid.sheet, wb = app.wb, g = app.grid.range;
  const r2 = Math.min(g.r2, g.r1 + 9999), c2 = Math.min(g.c2, g.c1 + 99);
  app.safe(() => wb.group(() => {
    wb.setStyle(s, { r1: g.r1, c1: g.c1, r2, c2 }, { cb: true, al: 'c' });
    const list: [number, number, boolean][] = [];
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { const v = wb.value(s, r, c); if (typeof v !== 'boolean') list.push([r, c, false]); }
    wb.writeCells(s, list);
  }));
  app.grid.invalidate();
}
export function toggleCheckbox(app: App, r: number, c: number): boolean {
  const s = app.grid.sheet, wb = app.wb;
  if (!wb.style(s, r, c).cb || wb.rawText(s, r, c).startsWith('=')) return false;
  app.safe(() => wb.writeCells(s, [[r, c, !(wb.value(s, r, c) === true)]]));
  app.grid.invalidate(); app.syncSelection();
  return true;
}
