// In / Lưu PDF: thiết lập trang + in qua khung ẩn (trình duyệt cho phép "Lưu dưới dạng PDF")
import { rangeToA1, type Range } from '../core/address';
import { chartData } from '../core/chart';
import { buildPrintHtml, DEFAULT_PRINT } from '../core/print';
import type { PrintSetup, Workbook } from '../core/workbook';
import { chartPng } from './chartdraw';
import { parseSheetRef } from './charts';
import { field, h, input, openDialog, select } from './dialogs';

export function printHtmlFor(wb: Workbook, s: number, setup: PrintSetup, range: Range | undefined, withCharts: boolean) {
  const charts = withCharts ? wb.meta(s).charts.map((ch) => ({ src: chartPng(ch, chartData(wb, ch)), w: ch.w, h: ch.h })) : [];
  return buildPrintHtml(wb, s, { setup, range, charts });
}

export function printViaFrame(html: string) {
  document.querySelector('iframe.print-frame')?.remove();
  const f = document.createElement('iframe');
  f.className = 'print-frame';
  f.setAttribute('aria-hidden', 'true');
  f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
  document.body.appendChild(f);
  f.onload = () => {
    const w = f.contentWindow!;
    // đợi ảnh biểu đồ tải xong
    const imgs = [...w.document.images];
    Promise.all(imgs.map((im) => (im.complete ? 1 : new Promise((r) => { im.onload = im.onerror = r; })))).then(() => {
      w.focus();
      w.print();
    });
  };
  f.srcdoc = html;
}

export function printDialog(wb: Workbook, s: number, sel: Range, onMsg: (m: string, kind?: 'err') => void) {
  const cur: PrintSetup = { ...DEFAULT_PRINT, ...(wb.meta(s).print ?? {}) };
  const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
  const what = select([
    ['sheet', cur.area ? `Vùng in đã đặt (${rangeToA1(cur.area)})` : 'Toàn bộ trang tính'],
    ...(!single ? [['sel', `Vùng đang chọn (${rangeToA1(sel)})`] as [string, string]] : []),
    ['custom', 'Vùng tự nhập…'],
  ], 'sheet') as HTMLSelectElement;
  const area = input(cur.area ? rangeToA1(cur.area) : '', 'ví dụ: A1:H40');
  const areaF = field('Vùng in', area, 'Được lưu cùng sổ tính như “vùng in” của Excel.');
  areaF.hidden = true;
  what.onchange = () => { areaF.hidden = what.value !== 'custom'; };
  const orient = select([['portrait', 'Dọc'], ['landscape', 'Ngang']], cur.orientation) as HTMLSelectElement;
  const paper = select([['A4', 'A4'], ['A3', 'A3'], ['Letter', 'Letter']], cur.paper) as HTMLSelectElement;
  const margins = select([['normal', 'Thường'], ['narrow', 'Hẹp'], ['wide', 'Rộng']], cur.margins ?? 'normal') as HTMLSelectElement;
  const titles = input(String(cur.titleRows || 0));
  titles.setAttribute('inputmode', 'numeric');
  const fit = h('input', { type: 'checkbox' }) as HTMLInputElement; fit.checked = cur.fitWidth;
  const grid = h('input', { type: 'checkbox' }) as HTMLInputElement; grid.checked = cur.gridlines;
  const nCharts = wb.meta(s).charts.length;
  const charts = h('input', { type: 'checkbox' }) as HTMLInputElement; charts.checked = nCharts > 0;
  const err = h('p', { class: 'err-text', role: 'alert' });
  const body = h('div', { class: 'stack' },
    field('In', what), areaF,
    h('div', { class: 'row2' }, field('Hướng giấy', orient), field('Khổ giấy', paper)),
    h('div', { class: 'row2' }, field('Lề', margins), field('Lặp lại hàng tiêu đề', titles, 'Số hàng đầu in lại trên mỗi trang (0 = không).')),
    h('label', { class: 'chk' }, fit, ' Thu nhỏ vừa chiều ngang trang'),
    h('label', { class: 'chk' }, grid, ' In đường lưới'),
    ...(nCharts ? [h('label', { class: 'chk' }, charts, ` In kèm ${nCharts} biểu đồ`)] : []),
    h('p', { class: 'muted' }, 'Ở hộp thoại in của trình duyệt, chọn máy in hoặc “Lưu dưới dạng PDF”.'),
    err);
  const collect = (): { setup: PrintSetup; range?: Range } | null => {
    err.textContent = '';
    const tr = Math.max(0, Math.min(50, parseInt(titles.value) || 0));
    const setup: PrintSetup = { orientation: orient.value as PrintSetup['orientation'], paper: paper.value as PrintSetup['paper'], margins: margins.value as PrintSetup['margins'], fitWidth: fit.checked, gridlines: grid.checked, titleRows: tr, area: cur.area ?? null };
    let range: Range | undefined;
    if (what.value === 'sel') range = sel;
    if (what.value === 'custom') {
      const ref = area.value.trim() ? parseSheetRef(wb, area.value, s) : null;
      if (!ref || ref.sheet !== s) { err.textContent = 'Vùng in không hợp lệ (ví dụ: A1:H40).'; return null; }
      setup.area = ref.range;
    }
    if (!setup.area) delete setup.area;
    return { setup, range };
  };
  const same = (a: PrintSetup) => JSON.stringify({ ...DEFAULT_PRINT, ...(wb.meta(s).print ?? {}) }) === JSON.stringify({ ...DEFAULT_PRINT, ...a });
  openDialog('In / Lưu PDF', body, [
    { label: 'Hủy', act: () => {} },
    { label: 'Lưu thiết lập', act: () => { const c = collect(); if (!c) return false; if (!same(c.setup)) wb.setPrint(s, c.setup); onMsg('Đã lưu thiết lập trang (được ghi vào tệp .xlsx).'); } },
    { label: 'In…', kind: 'primary', act: () => {
      const c = collect(); if (!c) return false;
      if (!same(c.setup)) wb.setPrint(s, c.setup);
      const u = wb.used(s);
      if (!u.rows && !nCharts) { onMsg('Trang tính đang trống, không có gì để in.', 'err'); return; }
      try { printViaFrame(printHtmlFor(wb, s, c.setup, c.range, charts.checked)); }
      catch (e) { onMsg((e as Error).message, 'err'); }
    } },
  ]);
}
