// Tab ngữ cảnh Thiết kế Biểu đồ / Định dạng (0.9) khi chọn một biểu đồ
import { seriesFromRange, scatterFromRange, chartSourceRange, PALETTES, CHART_NAMES, type Chart } from '../core/chart';
import { MAX_COLS } from '../core/address';
import type { MenuItem } from './menu';
import { showMenu } from './menu';
import type { Tab, Btn, Item, Group } from './ribbon';
import type { App } from './app';

const STYLES: { name: string; fill?: string; border?: string; grid?: boolean }[] = [
  { name: 'Kiểu 1' }, { name: 'Kiểu 2 (viền)', border: '#bfbfbf' }, { name: 'Kiểu 3 (không lưới)', grid: false },
  { name: 'Kiểu 4 (nền xám)', fill: '#f2f2f2' }, { name: 'Kiểu 5 (nền xanh nhạt)', fill: '#eaf1fb', border: '#8eaadb' },
];
const COLORS: [string, string][] = [['', 'Mặc định'], ['#ffffff', 'Trắng'], ['#f2f2f2', 'Xám nhạt'], ['#fff2cc', 'Vàng nhạt'], ['#e2efda', 'Xanh lá nhạt'], ['#ddebf7', 'Xanh dương nhạt'], ['#fce4d6', 'Cam nhạt'], ['#262626', 'Đen']];

export function chartTabs(app: App): Tab[] {
  const s = app.grid.sheet, wb = app.wb;
  const cur = (): Chart | undefined => wb.meta(s).charts.find((c) => c.id === app.charts.sel);
  const upd = (p: Partial<Chart>) => { const c = cur(); if (!c?.id) return; app.safe(() => wb.updateChart(s, c.id!, p)); app.grid.invalidate(); app.ribbon.refresh(); };
  const B = (t: Btn['t'], id: string, icon: string, label: string, o: Partial<Btn> = {}): Btn => ({ t, id, icon, label, ...o });
  const G = (id: string, label: string, icon: string, items: Item[]): Group => ({ id, label, icon, items });
  // hộp nhập trong app (Electron không có window.prompt)
  const ask = (label: string, v: string, f: (x: string) => void) => {
    const i = input(v); const t = label.replace(/:.*$/, '');
    openDialog(t, h('div', { class: 'an-form' }, field(label, i)), [{ label: 'Hủy', act: () => {} }, { label: 'OK', kind: 'primary', act: () => f(i.value.trim()) }]);
    setTimeout(() => { i.focus(); i.select(); }, 40);
  };
  const elemMenu = (): MenuItem[] => {
    const c = cur(); if (!c) return [];
    return [
      { label: 'Tiêu đề Biểu đồ…', check: !!c.title, act: () => ask('Tiêu đề biểu đồ (để trống để ẩn):', c.title ?? '', (t) => upd({ title: t || undefined })) },
      { label: 'Tiêu đề Trục Ngang…', check: !!c.xTitle, act: () => ask('Tiêu đề trục ngang:', c.xTitle ?? '', (t) => upd({ xTitle: t || undefined })) },
      { label: 'Tiêu đề Trục Dọc…', check: !!c.yTitle, act: () => ask('Tiêu đề trục dọc:', c.yTitle ?? '', (t) => upd({ yTitle: t || undefined })) },
      ...(c.type === 'combo' ? [{ label: 'Tiêu đề Trục Phụ…', check: !!c.y2Title, act: () => ask('Tiêu đề trục phụ:', c.y2Title ?? '', (t) => upd({ y2Title: t || undefined })) }] : []),
      'sep',
      { label: 'Nhãn Dữ liệu', check: !!c.labels, act: () => upd({ labels: !c.labels }) },
      { label: 'Đường lưới', check: c.grid !== false, act: () => upd({ grid: c.grid === false ? undefined : false }) },
      'sep',
      { label: 'Chú giải: Không', check: c.legend === 'none', act: () => upd({ legend: 'none' }) },
      { label: 'Chú giải: Phải', check: c.legend === 'r', act: () => upd({ legend: 'r' }) },
      { label: 'Chú giải: Trên', check: c.legend === 't', act: () => upd({ legend: 't' }) },
      { label: 'Chú giải: Dưới', check: (c.legend ?? 'b') === 'b', act: () => upd({ legend: 'b' }) },
    ] as MenuItem[];
  };
  const menuAt = (f: () => MenuItem[]) => (r: DOMRect) => showMenu(f(), r.left, r.bottom + 2);
  const switchRC = () => {
    const c = cur(); if (!c) return;
    const src = chartSourceRange(c);
    if (!src) return;
    const byRow = c.series.length > 1 ? c.series[0].vals.range.r1 === c.series[0].vals.range.r2 : c.series[0].vals.range.c1 !== c.series[0].vals.range.c2;
    const series = c.type === 'scatter' ? scatterFromRange(wb, src.sheet, src.range) : seriesFromRange(wb, src.sheet, src.range, !byRow);
    if (series.length) upd({ series });
  };
  const moveChart = () => {
    const c = cur(); if (!c?.id) return;
    const names = wb.sheetIds().map((id) => [id, wb.sheetName(id)] as [number, string]);
    showMenu([{ label: 'Trang tính mới', act: () => {
      app.safe(() => wb.group(() => { const id = wb.addSheet(wb.uniqueSheetName('Biểu đồ')); wb.removeChart(s, c.id!); wb.addChart(id, { ...c, id: undefined, anchor: { r: 0, c: 0, dx: 10, dy: 10 }, w: 720, h: 420 }); app.grid.setSheet(id); }));
      app.renderSheets(); app.grid.invalidate();
    } }, 'sep', ...names.filter(([id]) => id !== s).map(([id, n]) => ({ label: n, act: () => {
      app.safe(() => wb.group(() => { wb.removeChart(s, c.id!); wb.addChart(id, { ...c, id: undefined }); }));
      app.grid.setSheet(id); app.renderSheets(); app.grid.invalidate();
    } }))] as MenuItem[], innerWidth / 2, 160);
  };
  const design: Tab = { id: 'cdesign', label: 'Thiết kế Biểu đồ', kt: 'JC', groups: [
    G('clayout', 'Bố cục Biểu đồ', 'chart', [B('big', 'cd-elem', 'plus', 'Thêm Thành phần Biểu đồ', { kt: 'A', drop: menuAt(elemMenu) })]),
    G('cstyles', 'Kiểu Biểu đồ', 'palette', [
      B('big', 'cd-colors', 'palette', 'Đổi Màu', { kt: 'H', drop: menuAt(() => PALETTES.map((p, i) => ({ label: p.name, check: (cur()?.colors ?? 0) === i, act: () => upd({ colors: i || undefined }) })) as MenuItem[]) }),
      { t: 'gal', id: 'cd-style', label: 'Kiểu Biểu đồ', kt: 'S', items: STYLES.map((x) => ({ label: x.name, css: `background:${x.fill ?? '#fff'};border:1px solid ${x.border ?? '#d0d0d0'};font-size:11px;color:#333`, act: () => upd({ fill: x.fill, border: x.border, grid: x.grid }) })), act: () => {}, more: (r) => showMenu(STYLES.map((x) => ({ label: x.name, act: () => upd({ fill: x.fill, border: x.border, grid: x.grid }) })), r.left, r.bottom) },
    ]),
    G('cdata', 'Dữ liệu', 'table', [B('big', 'cd-switch', 'transpose', 'Chuyển Hàng/Cột', { kt: 'W', act: switchRC }), B('big', 'cd-select', 'table', 'Chọn Dữ liệu', { kt: 'E', act: () => { const c = cur(); if (c?.id) app.editChart(c.id); } })]),
    G('ctype', 'Loại', 'chart', [B('big', 'cd-type', 'chart', 'Thay đổi Loại Biểu đồ', { kt: 'C', drop: menuAt(() => (Object.keys(CHART_NAMES) as Chart['type'][]).map((t) => ({ label: CHART_NAMES[t], check: cur()?.type === t, act: () => {
      const c = cur(); if (!c) return;
      if (t === 'scatter' || c.type === 'scatter') { const src = chartSourceRange(c); if (src) { const ser = t === 'scatter' ? scatterFromRange(wb, src.sheet, src.range) : seriesFromRange(wb, src.sheet, src.range); upd({ type: t, series: ser }); return; } }
      upd({ type: t, series: t === 'combo' ? c.series.map((x, i) => ({ ...x, kind: i === 0 ? 'col' : 'line' })) : c.series });
    } })) as MenuItem[]) })]),
    G('cloc', 'Vị trí', 'move', [B('big', 'cd-move', 'move', 'Di chuyển Biểu đồ', { kt: 'M', act: moveChart })]),
  ] };
  const cm = (px: number) => String(Math.round((px / 96) * 2.54 * 100) / 100).replace('.', ',');
  const fromCm = (t: string) => { const n = parseFloat(t.replace(',', '.')); return isFinite(n) && n > 0.5 ? Math.round((n / 2.54) * 96) : null; };
  const format: Tab = { id: 'cformat', label: 'Định dạng', kt: 'JA', groups: [
    G('cfill', 'Kiểu Hình', 'fill', [
      B('big', 'cf-fill', 'fill', 'Tô Hình', { kt: 'SF', drop: menuAt(() => COLORS.map(([v, n]) => ({ label: n, check: (cur()?.fill ?? '') === v, act: () => upd({ fill: v || undefined }) })) as MenuItem[]) }),
      B('big', 'cf-line', 'borders', 'Viền Hình', { kt: 'SO', drop: menuAt(() => [['', 'Không viền'], ['#bfbfbf', 'Xám'], ['#000000', 'Đen'], ['#4472c4', 'Xanh dương'], ['#217346', 'Xanh lá']].map(([v, n]) => ({ label: n, check: (cur()?.border ?? '') === v, act: () => upd({ border: v || undefined }) })) as MenuItem[]) }),
    ]),
    G('csize', 'Kích thước', 'autofit', [{ t: 'col', items: [
      { t: 'combo', id: 'cf-h', label: 'Chiều cao Hình (cm)', w: 70, kt: 'HH', get: () => (cur() ? cm(cur()!.h) : ''), set: (v) => { const n = fromCm(v); if (n) upd({ h: Math.min(n, 4000) }); }, list: () => {} },
      { t: 'combo', id: 'cf-w', label: 'Chiều rộng Hình (cm)', w: 70, kt: 'HW', get: () => (cur() ? cm(cur()!.w) : ''), set: (v) => { const n = fromCm(v); if (n) upd({ w: Math.min(n, 4000) }); }, list: () => {} },
    ] }]),
  ] };
  void MAX_COLS;
  return [design, format];
}

// ---------------- Biểu đồ thu nhỏ (Sparkline) ----------------
import { openDialog, h, field, input } from './dialogs';
import { parseSheetRef } from './charts';
import { rangeToA1 } from '../core/address';
import type { SparkGroup } from '../core/workbook';

export function sparkDialog(app: App, type: SparkGroup['type']) {
  const s = app.grid.sheet, wb = app.wb, g = app.grid.range;
  const src = input(rangeToA1(g)), loc = input('');
  src.classList.add('mono'); loc.classList.add('mono');
  // vị trí gợi ý: cột trống ngay bên phải dữ liệu (mỗi hàng một sparkline)
  if (!(g.r1 === g.r2 && g.c1 === g.c2)) loc.value = rangeToA1({ r1: g.r1, c1: g.c2 + 1, r2: g.r2, c2: g.c2 + 1 });
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog('Tạo Biểu đồ thu nhỏ', h('div', { class: 'an-form' }, h('b', {}, 'Chọn dữ liệu bạn muốn'), field('Vùng Dữ liệu', src), h('b', {}, 'Chọn nơi đặt biểu đồ thu nhỏ'), field('Vùng Vị trí', loc), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      const a = parseSheetRef(wb, src.value.trim().replace(/^=/, ''), s), b = parseSheetRef(wb, loc.value.trim().replace(/^=/, ''), s);
      if (!a) { err.textContent = 'Vùng dữ liệu không hợp lệ.'; return false; }
      if (!b || b.sheet !== s) { err.textContent = 'Vùng vị trí phải ở trang tính hiện hành.'; return false; }
      try { wb.addSparklines(s, type, a, b.range); } catch (e) { err.textContent = (e as Error).message; return false; }
      app.grid.select(b.range.r1, b.range.c1); app.grid.invalidate(); app.syncSelection(); app.scheduleSave();
    } },
  ]);
  setTimeout(() => (g.r1 === g.r2 && g.c1 === g.c2 ? src : loc).focus(), 40);
}

export function sparkTab(app: App): Tab {
  const s = app.grid.sheet, wb = app.wb;
  const at = () => ({ r: app.grid.sel.ar, c: app.grid.sel.ac });
  const cur = () => wb.sparkAt(s, at().r, at().c)?.g;
  const upd = (p: Partial<Omit<SparkGroup, 'items'>>) => { app.safe(() => wb.updateSparkGroup(s, at(), p)); app.grid.invalidate(); app.ribbon.refresh(); app.scheduleSave(); };
  const B = (t: Btn['t'], id: string, icon: string, label: string, o: Partial<Btn> = {}): Btn => ({ t, id, icon, label, ...o });
  const G = (id: string, label: string, icon: string, items: Item[]): Group => ({ id, label, icon, items });
  const chk = (id: string, label: string, k: 'hi' | 'lo' | 'negPts' | 'first' | 'last' | 'markers', kt: string): Item => ({ t: 'chk', id, label, kt, get: () => !!cur()?.[k], set: (v) => upd({ [k]: v }) });
  const colors: [string, string][] = [['#4472C4', 'Xanh dương'], ['#217346', 'Xanh lá'], ['#ED7D31', 'Cam'], ['#7F7F7F', 'Xám'], ['#000000', 'Đen'], ['#C00000', 'Đỏ']];
  return { id: 'sdesign', label: 'Biểu đồ thu nhỏ', kt: 'JS', groups: [
    G('stype', 'Loại', 'sparkLine', [
      B('big', 'sd-line', 'sparkLine', 'Đường', { kt: 'L', on: () => cur()?.type === 'line', act: () => upd({ type: 'line' }) }),
      B('big', 'sd-col', 'sparkCol', 'Cột', { kt: 'C', on: () => cur()?.type === 'column', act: () => upd({ type: 'column' }) }),
      B('big', 'sd-wl', 'winLoss', 'Thắng/Thua', { kt: 'W', on: () => cur()?.type === 'stacked', act: () => upd({ type: 'stacked' }) }),
    ]),
    G('sshow', 'Hiện', 'check', [
      { t: 'col', items: [chk('sd-hi', 'Điểm Cao', 'hi', 'H'), chk('sd-lo', 'Điểm Thấp', 'lo', 'O'), chk('sd-neg', 'Điểm Âm', 'negPts', 'N')] },
      { t: 'col', items: [chk('sd-first', 'Điểm Đầu', 'first', 'F'), chk('sd-last', 'Điểm Cuối', 'last', 'P'), chk('sd-mk', 'Dấu', 'markers', 'M')] },
    ]),
    G('sstyle', 'Kiểu', 'palette', [B('big', 'sd-color', 'palette', 'Màu Biểu đồ thu nhỏ', { kt: 'I', drop: (r) => showMenu(colors.map(([v, n]) => ({ label: n, check: cur()?.color?.toUpperCase() === v, act: () => upd({ color: v }) })), r.left, r.bottom + 2) })]),
    G('sgroup', 'Nhóm', 'trash', [B('big', 'sd-clear', 'eraser', 'Xóa', { kt: 'C', split: true, act: () => { app.safe(() => wb.clearSparklines(s, app.grid.range)); app.grid.invalidate(); app.syncSelection(); }, drop: (r) => showMenu([{ label: 'Xóa Biểu đồ thu nhỏ Đã chọn', act: () => { app.safe(() => wb.clearSparklines(s, app.grid.range)); app.grid.invalidate(); app.syncSelection(); } }, { label: 'Xóa Nhóm Biểu đồ thu nhỏ Đã chọn', act: () => { app.safe(() => wb.clearSparklines(s, app.grid.range, true)); app.grid.invalidate(); app.syncSelection(); } }], r.left, r.bottom + 2) })]),
  ] };
}
