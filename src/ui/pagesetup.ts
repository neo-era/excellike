// Hộp thoại Thiết lập Trang như Excel 365: Trang · Lề · Đầu trang/Chân trang · Trang tính
import { colName } from '../core/address';
import { hfJoin, hfParse, hfRender, MARGIN_PRESETS, normPrint, PAPERS, storePrint, type NPrint } from '../core/paginate';
import { openDialog, h } from './dialogs';
import { parseSheetRef } from './charts';
import type { App } from './app';

// mã hiển thị trong ô soạn đầu/chân trang (như Excel: &[Trang], &[Trang tính]…)
const CODES: [string, string, string][] = [['P', '&[Trang]', 'Số trang'], ['N', '&[Số trang]', 'Tổng số trang'], ['D', '&[Ngày]', 'Ngày'], ['T', '&[Giờ]', 'Giờ'], ['F', '&[Tệp]', 'Tên tệp'], ['A', '&[Trang tính]', 'Tên trang tính']];
// Chỉ đổi các mã nội dung (&P &N &D &T &F &A) sang dạng dễ đọc; mọi mã khác (&& , &"phông", &14, &B, &K…) giữ nguyên
// như ô soạn đầu/chân trang của Excel (muốn gõ ký tự & thì gõ &&).
export function rawToFriendly(t: string) {
  let o = '';
  for (let i = 0; i < t.length; i++) {
    if (t[i] !== '&') { o += t[i]; continue; }
    const n = t[i + 1];
    const c = CODES.find(([k]) => k === n);
    if (c) { o += c[1]; i++; } else if (n !== undefined) { o += '&' + n; i++; } else o += '&';
  }
  return o;
}
export function friendlyToRaw(t: string) {
  let o = '';
  for (let i = 0; i < t.length; i++) {
    const c = t[i] === '&' ? CODES.find(([, f]) => t.startsWith(f, i)) : undefined;
    if (c) { o += '&' + c[0]; i += c[1].length - 1; } else o += t[i];
  }
  return o;
}
const PRESETS = (sheet: string, file: string): [string, string][] => [
  ['', '(không có)'],
  ['&CTrang &P', 'Trang 1'],
  ['&CTrang &P / &N', 'Trang 1 / ?'],
  ['&C&A', sheet],
  ['&C&F', file],
  ['&LBí mật&C&D&RTrang &P', `Bí mật, ${new Date().toLocaleDateString('vi-VN')}, Trang 1`],
  ['&C&F&RTrang &P', `${file}, Trang 1`],
  ['&C&A&RTrang &P', `${sheet}, Trang 1`],
  ['&CTrang &P&R&A', `Trang 1, ${sheet}`],
  ['&LNgười lập&CTrang &P&R&D', `Người lập, Trang 1, ${new Date().toLocaleDateString('vi-VN')}`],
];
const cm = (inch: number) => String(Math.round(inch * 2.54 * 100) / 100).replace('.', ',');
const inch = (s: string) => { const n = parseFloat(s.replace(',', '.')); return isNaN(n) || n < 0 ? null : Math.round((n / 2.54) * 1000) / 1000; };

export function pageSetupDialog(app: App, tab: 'page' | 'margins' | 'hf' | 'sheet', onDone: () => void) {
  const wb = app.wb, s = app.grid.sheet;
  const p: NPrint = normPrint(wb.meta(s).print);
  const sheetName = wb.sheetName(s), file = app.fileName || wb.name + '.xlsx';
  const tabs: [typeof tab, string][] = [['page', 'Trang'], ['margins', 'Lề'], ['hf', 'Đầu trang/Chân trang'], ['sheet', 'Trang tính']];
  const bar = h('div', { class: 'ps-tabs', role: 'tablist' });
  const panes: Record<string, HTMLElement> = {};
  const box = h('div', { class: 'ps' }, bar);
  for (const [k, l] of tabs) {
    const b = h('button', { role: 'tab', 'aria-selected': String(k === tab), 'data-t': k }, l);
    b.onclick = () => { bar.querySelectorAll('[role=tab]').forEach((x) => x.setAttribute('aria-selected', String(x === b))); for (const [kk] of tabs) panes[kk].hidden = kk !== k; };
    bar.append(b);
    panes[k] = h('div', { class: 'ps-pane stack', role: 'tabpanel' });
    panes[k].hidden = k !== tab;
    box.append(panes[k]);
  }
  const radio = (name: string, value: string, label: string, checked: boolean, id?: string) => { const i = h('input', { type: 'radio', name, value, ...(id ? { id } : {}) }) as HTMLInputElement; i.checked = checked; return h('label', { class: 'chk' }, i, ' ' + label); };
  const num = (id: string, v: string, w = '90px') => { const i = h('input', { type: 'text', id, inputmode: 'decimal', style: `width:${w};height:32px` }) as HTMLInputElement; i.value = v; return i; };
  const chk = (id: string, label: string, on: boolean) => { const i = h('input', { type: 'checkbox', id }) as HTMLInputElement; i.checked = on; return h('label', { class: 'chk' }, i, ' ' + label); };
  // ---- Trang ----
  const pct = num('ps-scale', String(p.scale));
  const fw = num('ps-fitw', String(p.fit?.w ?? 1), '60px'), fh = num('ps-fith', String(p.fit?.h || ''), '60px');
  fh.placeholder = 'tự động';
  const paper = h('select', { id: 'ps-paper' }) as HTMLSelectElement;
  for (const [k, v] of Object.entries(PAPERS)) paper.append(h('option', { value: k }, v.label));
  paper.value = p.paper;
  const first = num('ps-first', p.firstPage ? String(p.firstPage) : 'Tự động', '120px');
  panes.page.append(
    h('b', {}, 'Hướng'),
    h('div', { class: 'row-end', style: 'justify-content:flex-start;gap:24px' }, radio('ps-or', 'portrait', 'Dọc', p.orientation === 'portrait'), radio('ps-or', 'landscape', 'Ngang', p.orientation === 'landscape')),
    h('b', {}, 'Co giãn'),
    h('div', { class: 'chk' }, radio('ps-sc', 'pct', 'Điều chỉnh thành', !p.fit, 'ps-pct'), pct, '% kích thước thường'),
    h('div', { class: 'chk' }, radio('ps-sc', 'fit', 'Vừa với', !!p.fit, 'ps-fit'), fw, 'trang rộng ×', fh, 'trang cao'),
    h('label', { class: 'fld' }, h('span', {}, 'Khổ giấy'), paper),
    h('label', { class: 'fld' }, h('span', {}, 'Số trang đầu tiên'), first),
  );
  pct.onfocus = () => ((box.querySelector('#ps-pct') as HTMLInputElement).checked = true);
  fw.onfocus = fh.onfocus = () => ((box.querySelector('#ps-fit') as HTMLInputElement).checked = true);
  // ---- Lề ----
  const mg = { t: num('ps-mt', cm(p.mg.t)), h: num('ps-mh', cm(p.mg.h)), l: num('ps-ml', cm(p.mg.l)), r: num('ps-mr', cm(p.mg.r)), b: num('ps-mb', cm(p.mg.b)), f: num('ps-mf', cm(p.mg.f)) };
  const lab = (t: string, i: HTMLElement) => h('label', { class: 'fld' }, h('span', {}, t), i);
  panes.margins.append(
    h('div', { class: 'ps-grid' }, h('span', {}), lab('Trên (cm)', mg.t), lab('Đầu trang (cm)', mg.h), lab('Trái (cm)', mg.l), h('span', {}), lab('Phải (cm)', mg.r), h('span', {}), lab('Dưới (cm)', mg.b), lab('Chân trang (cm)', mg.f)),
    h('b', {}, 'Căn giữa trên trang'),
    h('div', { class: 'row-end', style: 'justify-content:flex-start;gap:24px' }, chk('ps-ch', 'Theo chiều ngang', p.center.h), chk('ps-cv', 'Theo chiều dọc', p.center.v)),
  );
  // ---- Đầu trang/Chân trang ----
  const ctx = { page: 1, pages: 1, sheet: sheetName, file, date: new Date() };
  const hfBlock = (kind: 'header' | 'footer') => {
    const raw = kind === 'header' ? p.header : p.footer;
    const parts = hfParse(raw);
    const pre = h('div', { class: 'ps-prev', 'aria-label': 'Xem trước' }, h('span'), h('span'), h('span'));
    const sel = h('select', { id: kind === 'header' ? 'ps-header' : 'ps-footer' }) as HTMLSelectElement;
    const presets = PRESETS(sheetName, file);
    for (const [v, l] of presets) sel.append(h('option', { value: v }, l));
    sel.append(h('option', { value: '__custom' }, '(tùy chỉnh)'));
    sel.value = presets.some(([v]) => v === raw) ? raw : '__custom';
    const k = kind === 'header' ? 'h' : 'f';
    const ta = (x: 'l' | 'c' | 'r', label: string) => { const t = h('textarea', { id: `ps-${k}${x}`, 'aria-label': label }) as HTMLTextAreaElement; t.value = rawToFriendly(parts[x] ?? ''); return t; };
    const tl = ta('l', 'Phần trái'), tc = ta('c', 'Phần giữa'), tr = ta('r', 'Phần phải');
    let dirty = false;
    // chưa sửa gì thì giữ nguyên chuỗi gốc (kể cả mã phông/cỡ chữ/màu)
    const value = () => (dirty ? hfJoin({ l: friendlyToRaw(tl.value), c: friendlyToRaw(tc.value), r: friendlyToRaw(tr.value) }) : raw);
    const paint = () => { const v = hfParse(value()); [...pre.children].forEach((sp, i) => (sp.textContent = hfRender(v[(['l', 'c', 'r'] as const)[i]] ?? '', ctx))); };
    sel.onchange = () => { if (sel.value === '__custom') return; dirty = true; const v = hfParse(sel.value); tl.value = rawToFriendly(v.l ?? ''); tc.value = rawToFriendly(v.c ?? ''); tr.value = rawToFriendly(v.r ?? ''); paint(); };
    let last: HTMLTextAreaElement = tc;
    for (const t of [tl, tc, tr]) { t.oninput = () => { dirty = true; sel.value = presets.some(([v]) => v === value()) ? value() : '__custom'; paint(); }; t.onfocus = () => (last = t); }
    const codes = h('div', { class: 'ps-codes' });
    for (const [, f, l] of CODES) { const b = h('button', { type: 'button', title: 'Chèn ' + l }, l); b.onmousedown = (e) => e.preventDefault(); b.onclick = () => { const a = last.selectionStart ?? last.value.length; last.value = last.value.slice(0, a) + f + last.value.slice(last.selectionEnd ?? a); last.dispatchEvent(new Event('input')); last.focus(); }; codes.append(b); }
    paint();
    return { el: h('div', { class: 'stack' }, pre, h('label', { class: 'fld' }, h('span', {}, kind === 'header' ? 'Đầu trang' : 'Chân trang'), sel), h('div', { class: 'ps-hf' }, tl, tc, tr), codes), value };
  };
  const hd = hfBlock('header'), ft = hfBlock('footer');
  panes.hf.append(hd.el, h('hr'), ft.el);
  // ---- Trang tính ----
  const areaIn = h('input', { type: 'text', id: 'ps-area', placeholder: 'ví dụ: A1:H40', style: 'height:32px' }) as HTMLInputElement;
  if (p.area) areaIn.value = `${colName(p.area.c1)}${p.area.r1 + 1}:${colName(p.area.c2)}${p.area.r2 + 1}`;
  const rowsIn = h('input', { type: 'text', id: 'ps-rows', placeholder: 'ví dụ: $1:$3', style: 'height:32px' }) as HTMLInputElement;
  if (p.titleRowRange) rowsIn.value = `$${p.titleRowRange[0] + 1}:$${p.titleRowRange[1] + 1}`;
  const colsIn = h('input', { type: 'text', id: 'ps-cols', placeholder: 'ví dụ: $A:$B', style: 'height:32px' }) as HTMLInputElement;
  if (p.titleColRange) colsIn.value = `$${colName(p.titleColRange[0])}:$${colName(p.titleColRange[1])}`;
  panes.sheet.append(
    lab('Vùng in', areaIn),
    h('b', {}, 'In tiêu đề'),
    lab('Hàng lặp lại ở đầu', rowsIn),
    lab('Cột lặp lại ở bên trái', colsIn),
    h('b', {}, 'In'),
    h('div', { class: 'row-end', style: 'justify-content:flex-start;gap:24px;flex-wrap:wrap' }, chk('ps-grid', 'Đường lưới', p.gridlines), chk('ps-heads', 'Đầu đề hàng và cột', p.headings)),
    h('b', {}, 'Thứ tự trang'),
    h('div', { class: 'row-end', style: 'justify-content:flex-start;gap:24px' }, radio('ps-ord', 'down', 'Xuống, rồi ngang', p.order === 'down', 'ps-down'), radio('ps-ord', 'over', 'Ngang, rồi xuống', p.order === 'over', 'ps-over')),
  );
  const err = h('p', { class: 'err-text', role: 'alert' });
  box.append(err);
  const collect = (): NPrint | null => {
    err.textContent = '';
    const q = (sel: string) => box.querySelector(sel) as HTMLInputElement;
    const n: NPrint = { ...p };
    n.orientation = (q('input[name=ps-or]:checked').value as NPrint['orientation']);
    n.paper = paper.value as NPrint['paper'];
    if (q('#ps-fit').checked) {
      const w = parseInt(fw.value) || 0, hh = parseInt(fh.value) || 0;
      if (!w && !hh) { err.textContent = 'Nhập số trang rộng hoặc số trang cao.'; return null; }
      n.fit = { w, h: hh };
    } else {
      const sc = parseInt(pct.value);
      if (!(sc >= 10 && sc <= 400)) { err.textContent = 'Tỷ lệ phải từ 10% đến 400%.'; return null; }
      n.fit = null; n.scale = sc;
    }
    const fp = parseInt(first.value);
    n.firstPage = fp > 0 ? fp : null;
    // giữ nguyên giá trị inch gốc nếu ô không đổi (tránh sai số làm tròn cm ↔ inch)
    const kept = (k: keyof NPrint['mg']) => (mg[k].value === cm(p.mg[k]) ? p.mg[k] : inch(mg[k].value));
    const m2 = { t: kept('t'), b: kept('b'), l: kept('l'), r: kept('r'), h: kept('h'), f: kept('f') };
    if (Object.values(m2).some((v) => v === null)) { err.textContent = 'Lề phải là số không âm (cm).'; return null; }
    n.mg = m2 as NPrint['mg'];
    {
      const [pw, ph] = PAPERS[n.paper].mm.map((x) => x / 25.4);
      const [W, H] = n.orientation === 'landscape' ? [ph, pw] : [pw, ph];
      if (n.mg.l + n.mg.r > W - 0.5 || n.mg.t + n.mg.b > H - 0.5) { err.textContent = 'Lề không vừa khổ giấy.'; return null; }
    }
    n.margins = (Object.entries(MARGIN_PRESETS).find(([, v]) => JSON.stringify(v) === JSON.stringify(n.mg))?.[0] ?? 'custom') as NPrint['margins'];
    n.center = { h: q('#ps-ch').checked, v: q('#ps-cv').checked };
    n.header = hd.value(); n.footer = ft.value();
    n.gridlines = q('#ps-grid').checked; n.headings = q('#ps-heads').checked;
    n.order = q('input[name=ps-ord]:checked').value as NPrint['order'];
    if (areaIn.value.trim()) {
      const r = parseSheetRef(wb, areaIn.value.trim(), s);
      if (!r || r.sheet !== s) { err.textContent = 'Vùng in không hợp lệ (ví dụ: A1:H40).'; return null; }
      n.area = r.range;
    } else n.area = null;
    const rr = /^\$?(\d+):\$?(\d+)$/.exec(rowsIn.value.trim()), cc = /^\$?([A-Z]+):\$?([A-Z]+)$/i.exec(colsIn.value.trim());
    if (rowsIn.value.trim() && !rr) { err.textContent = 'Hàng lặp lại phải dạng $1:$2.'; return null; }
    if (colsIn.value.trim() && !cc) { err.textContent = 'Cột lặp lại phải dạng $A:$B.'; return null; }
    n.titleRowRange = rr ? [Math.min(+rr[1], +rr[2]) - 1, Math.max(+rr[1], +rr[2]) - 1] : null;
    const ci = (x: string) => { let v = 0; for (const ch of x.toUpperCase()) v = v * 26 + ch.charCodeAt(0) - 64; return v - 1; };
    n.titleColRange = cc ? [Math.min(ci(cc[1]), ci(cc[2])), Math.max(ci(cc[1]), ci(cc[2]))] : null;
    return n;
  };
  const save = (): boolean => { const n = collect(); if (!n) return false; app.safeSetPrint(s, storePrint(n)); onDone(); return true; };
  const dlg = openDialog('Thiết lập Trang', box, [
    { label: 'In…', act: () => { if (!save()) return false; app.backstage.open('print'); } },
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => (save() ? undefined : false) },
  ]);
  dlg.querySelector('.dlg')?.classList.add('wide');
}
