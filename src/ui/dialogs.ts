// Hộp thoại dùng chung (modal) + các hộp thoại: sắp xếp, định dạng có điều kiện, kiểm tra dữ liệu
import { colName, rangeToA1, parseRange, type Range } from '../core/address';
import type { Workbook } from '../core/workbook';
import type { CFRule, DVRule } from '../core/rules';
import { dvDescribe } from '../core/rules';
import { ic } from './icons';

export interface DlgButton { label: string; kind?: 'primary' | 'danger' | ''; act: () => boolean | void }

let current: HTMLElement | null = null;
let currentOnClose: (() => void) | undefined;
/** Đóng hộp thoại đang mở; luôn gọi onClose của nó (kể cả khi đóng bằng Esc hoặc bị hộp thoại khác thay) */
export function closeDialog() {
  const cb = currentOnClose;
  currentOnClose = undefined;
  current?.remove(); current = null;
  document.removeEventListener('keydown', onKey, true);
  cb?.();
}
function onKey(e: KeyboardEvent) {
  if (!current) return;
  if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeDialog(); }
  if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT') {
    e.preventDefault();
    (current.querySelector('.dlg-foot .primary') as HTMLButtonElement | null)?.click();
  }
}
export function openDialog(title: string, body: HTMLElement, buttons: DlgButton[], onClose?: () => void) {
  closeDialog();
  const back = document.createElement('div');
  back.className = 'dlg-back';
  back.innerHTML = `<div class="dlg" role="dialog" aria-modal="true" aria-labelledby="dlg-t"><header><h2 id="dlg-t"></h2><button class="dlg-x" aria-label="Đóng">${ic('x', 18)}</button></header><div class="dlg-body"></div><footer class="dlg-foot"></footer></div>`;
  back.querySelector('#dlg-t')!.textContent = title;
  back.querySelector('.dlg-body')!.appendChild(body);
  const foot = back.querySelector('.dlg-foot')!;
  for (const b of buttons) {
    const el = document.createElement('button');
    el.className = 'btn ' + (b.kind ?? '');
    el.textContent = b.label;
    el.onclick = () => { if (current === back && b.act() !== false && current === back) closeDialog(); };
    foot.appendChild(el);
  }
  (back.querySelector('.dlg-x') as HTMLButtonElement).onclick = () => { if (current === back) closeDialog(); };
  back.addEventListener('pointerdown', (e) => { if (e.target === back && current === back) closeDialog(); });
  document.body.appendChild(back);
  current = back;
  currentOnClose = onClose;
  document.addEventListener('keydown', onKey, true);
  setTimeout(() => (back.querySelector('input, select, textarea, .dlg-body button') as HTMLElement | null)?.focus(), 30);
  return back;
}

export function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, string> = {}, ...kids: (Node | string)[]) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
  for (const k of kids) el.append(k);
  return el;
}
export function field(label: string, control: HTMLElement, hint?: string) {
  const id = control.id || 'f' + Math.random().toString(36).slice(2, 8);
  control.id = id;
  const l = h('label', { for: id, class: 'fld' }, h('span', {}, label), control);
  if (hint) l.append(h('small', {}, hint));
  return l;
}
export function select(opts: [string, string][], value?: string) {
  const s = h('select');
  for (const [v, t] of opts) { const o = h('option', { value: v }, t); if (v === value) o.selected = true; s.append(o); }
  return s;
}
export const input = (value = '', ph = '') => { const i = h('input', { type: 'text', placeholder: ph, spellcheck: 'false' }); i.value = value; return i; };

const colLabel = (wb: Workbook, s: number, g: Range, header: boolean, c: number) => {
  const t = header ? wb.display(s, g.r1, c).text : '';
  return t ? `${t} (cột ${colName(c)})` : `Cột ${colName(c)}`;
};

// ------------------------- Sắp xếp -------------------------
export function sortDialog(wb: Workbook, s: number, g: Range, onDone: (err?: string) => void) {
  let header = wb.guessHeader(s, g);
  const body = h('div', { class: 'stack' });
  const range = h('p', { class: 'muted' }, `Vùng sắp xếp: ${rangeToA1(g)}`);
  const hdr = h('input', { type: 'checkbox' }) as HTMLInputElement; hdr.checked = header;
  const levels = h('div', { class: 'stack' });
  const cols = () => [...Array(g.c2 - g.c1 + 1).keys()].map((i) => [String(g.c1 + i), colLabel(wb, s, g, header, g.c1 + i)] as [string, string]);
  const rows: { col: HTMLSelectElement; ord: HTMLSelectElement }[] = [];
  const addLevel = (col?: number) => {
    if (rows.length >= 3) return;
    const c = select(cols(), String(col ?? g.c1)), o = select([['asc', 'Tăng dần (A → Z, nhỏ → lớn)'], ['desc', 'Giảm dần (Z → A, lớn → nhỏ)']]);
    rows.push({ col: c as HTMLSelectElement, ord: o as HTMLSelectElement });
    levels.append(h('div', { class: 'row2' }, field(rows.length === 1 ? 'Sắp xếp theo' : 'Rồi theo', c), field('Thứ tự', o)));
  };
  addLevel();
  hdr.onchange = () => { header = hdr.checked; for (const r of rows) { const v = r.col.value; r.col.replaceChildren(...cols().map(([val, t]) => { const op = h('option', { value: val }, t); if (val === v) op.selected = true; return op; })); } };
  const more = h('button', { class: 'btn link', type: 'button' }, '+ Thêm mức sắp xếp');
  more.onclick = () => addLevel();
  body.append(range, h('label', { class: 'chk' }, hdr, ' Dữ liệu có hàng tiêu đề'), levels, more);
  openDialog('Sắp xếp', body, [
    { label: 'Hủy', act: () => {} },
    { label: 'Sắp xếp', kind: 'primary', act: () => {
      try { wb.sort(s, g, rows.map((r) => ({ col: +r.col.value, desc: r.ord.value === 'desc' })), header); onDone(); }
      catch (e) { onDone((e as Error).message); }
    } },
  ]);
}

// ------------------------- Định dạng có điều kiện -------------------------
export const CF_PRESETS: [string, CFRule['fmt']][] = [
  ['Tô đỏ nhạt, chữ đỏ đậm', { bg: '#ffc7ce', fg: '#9c0006' }],
  ['Tô vàng nhạt, chữ vàng đậm', { bg: '#ffeb9c', fg: '#9c5700' }],
  ['Tô xanh nhạt, chữ xanh đậm', { bg: '#c6efce', fg: '#006100' }],
  ['Chữ đỏ', { fg: '#c00000' }],
  ['Tô đỏ nhạt', { bg: '#ffc7ce' }],
  ['Chữ đậm', { b: true }],
];
export function cfRuleDialog(wb: Workbook, s: number, g: Range, kind: CFRule['type'] | 'gt' | 'lt' | 'between' | 'eq', onDone: () => void) {
  const body = h('div', { class: 'stack' });
  const titles: Record<string, string> = { gt: 'Tô ô có giá trị lớn hơn', lt: 'Tô ô có giá trị nhỏ hơn', between: 'Tô ô có giá trị trong khoảng', eq: 'Tô ô có giá trị bằng', text: 'Tô ô có chứa chữ', top: 'Tô các giá trị lớn nhất', bottom: 'Tô các giá trị nhỏ nhất' };
  const v1 = input('', kind === 'text' ? 'ví dụ: Vật tư' : 'ví dụ: 1.000.000');
  const v2 = input('', 'ví dụ: 5.000.000');
  const rank = input('10');
  const pct = h('input', { type: 'checkbox' }) as HTMLInputElement;
  const preset = select(CF_PRESETS.map(([t], i) => [String(i), t] as [string, string]));
  body.append(h('p', { class: 'muted' }, `Áp dụng cho: ${rangeToA1(g)}`));
  if (kind === 'top' || kind === 'bottom') body.append(h('div', { class: 'row2' }, field('Số mục', rank), h('label', { class: 'chk' }, pct, ' tính theo %')));
  else if (kind === 'between') body.append(h('div', { class: 'row2' }, field('Từ', v1), field('Đến', v2)));
  else body.append(field(kind === 'text' ? 'Ô chứa' : 'Giá trị', v1));
  body.append(field('Định dạng', preset));
  openDialog(titles[kind] ?? 'Định dạng có điều kiện', body, [
    { label: 'Hủy', act: () => {} },
    { label: 'Áp dụng', kind: 'primary', act: () => {
      const fmt = CF_PRESETS[+(preset as HTMLSelectElement).value][1];
      let rule: CFRule;
      if (kind === 'top' || kind === 'bottom') rule = { range: g, type: kind, rank: Math.max(1, parseInt(rank.value) || 10), percent: pct.checked, fmt };
      else if (kind === 'text') { if (!v1.value.trim()) { v1.focus(); return false; } rule = { range: g, type: 'text', text: v1.value.trim(), fmt }; }
      else {
        if (!v1.value.trim() || (kind === 'between' && !v2.value.trim())) { (v1.value.trim() ? v2 : v1).focus(); return false; }
        rule = { range: g, type: 'cellIs', op: kind as CFRule['op'], v1: v1.value.trim(), v2: kind === 'between' ? v2.value.trim() : undefined, fmt };
      }
      wb.addCF(s, rule); onDone();
    } },
  ]);
}
export function describeCF(r: CFRule): string {
  const op: Record<string, string> = { gt: '>', lt: '<', ge: '≥', le: '≤', eq: '=', ne: '≠', between: 'trong khoảng', notBetween: 'ngoài khoảng' };
  if (r.raw) return 'Quy tắc từ tệp Excel (giữ nguyên khi lưu)';
  switch (r.type) {
    case 'cellIs': return r.op === 'between' || r.op === 'notBetween' ? `Giá trị ${op[r.op]} ${r.v1} – ${r.v2}` : `Giá trị ${op[r.op ?? 'gt']} ${r.v1}`;
    case 'text': return `Chứa “${r.text}”`;
    case 'dup': return 'Giá trị trùng lặp';
    case 'unique': return 'Giá trị duy nhất';
    case 'top': return `${r.rank}${r.percent ? '%' : ''} lớn nhất`;
    case 'bottom': return `${r.rank}${r.percent ? '%' : ''} nhỏ nhất`;
    case 'aboveAvg': return 'Trên trung bình';
    case 'belowAvg': return 'Dưới trung bình';
    case 'scale': return `Thang màu ${r.colors?.length ?? 2} màu`;
    case 'bar': return 'Thanh dữ liệu';
  }
  return r.type;
}
export function cfManageDialog(wb: Workbook, s: number, onChange: () => void) {
  const body = h('div', { class: 'stack' });
  const render = () => {
    const list = wb.meta(s).cf;
    body.replaceChildren();
    if (!list.length) { body.append(h('p', { class: 'muted' }, 'Sheet này chưa có quy tắc định dạng có điều kiện.')); return; }
    const ul = h('ul', { class: 'rules' });
    for (const r of list) {
      const sw = h('span', { class: 'sw' }, 'Aa');
      const f = r.fmt ?? {};
      if (f.bg) sw.style.background = f.bg;
      if (f.fg) sw.style.color = f.fg;
      if (f.b) sw.style.fontWeight = '700';
      if (r.type === 'scale') sw.style.background = `linear-gradient(90deg, ${(r.colors ?? []).join(',')})`;
      if (r.type === 'bar') sw.style.background = `linear-gradient(90deg, ${r.colors?.[0] ?? '#638ec6'} 60%, transparent 60%)`;
      const del = h('button', { class: 'icon-btn', 'aria-label': 'Xóa quy tắc' });
      del.innerHTML = ic('trash', 16);
      del.onclick = () => { wb.removeCF(s, r.id!); onChange(); render(); };
      ul.append(h('li', {}, sw, h('span', { class: 'grow' }, h('b', {}, describeCF(r)), h('small', {}, rangeToA1(r.range))), del));
    }
    body.append(ul, h('p', { class: 'muted' }, 'Quy tắc ở trên được ưu tiên khi trùng thuộc tính.'));
  };
  render();
  openDialog('Quản lý định dạng có điều kiện', body, [{ label: 'Đóng', kind: 'primary', act: () => {} }]);
}

// ------------------------- Kiểm tra dữ liệu -------------------------
export function dvDialog(wb: Workbook, s: number, g: Range, onDone: (msg?: string) => void) {
  const cur = wb.dvAt(s, g.r1, g.c1);
  const body = h('div', { class: 'stack' });
  const type = select([['list', 'Danh sách'], ['whole', 'Số nguyên'], ['decimal', 'Số thập phân'], ['date', 'Ngày'], ['textLength', 'Độ dài chữ']], cur?.type === 'custom' ? 'list' : cur?.type ?? 'list') as HTMLSelectElement;
  const op = select([['between', 'trong khoảng'], ['notBetween', 'ngoài khoảng'], ['eq', 'bằng'], ['ne', 'khác'], ['gt', 'lớn hơn'], ['lt', 'nhỏ hơn'], ['ge', 'lớn hơn hoặc bằng'], ['le', 'nhỏ hơn hoặc bằng']], cur?.op ?? 'between') as HTMLSelectElement;
  const v1 = input(cur?.v1 ?? ''), v2 = input(cur?.v2 ?? '');
  const list = input(cur?.list ? cur.list.join('; ') : cur?.listRef ? '=' + cur.listRef.replace(/^=/, '') : '', 'ví dụ: bộ; cột; m   hoặc   =$H$2:$H$20');
  const blank = h('input', { type: 'checkbox' }) as HTMLInputElement; blank.checked = cur?.allowBlank !== false;
  const err = input(cur?.error ?? '', 'Để trống sẽ dùng câu báo mặc định');
  const prompt = input(cur?.prompt ?? '', 'Hiện khi chọn ô (không bắt buộc)');
  const fList = field('Các giá trị cho phép', list, 'Ngăn cách bằng dấu ; hoặc nhập vùng ô bắt đầu bằng =');
  const fOp = field('Điều kiện', op), fV1 = field('Giá trị', v1), fV2 = field('Đến', v2);
  const sync = () => {
    const isList = type.value === 'list';
    fList.hidden = !isList; fOp.hidden = isList; fV1.hidden = isList;
    fV2.hidden = isList || !['between', 'notBetween'].includes(op.value);
    (fV1.querySelector('span') as HTMLElement).textContent = ['between', 'notBetween'].includes(op.value) ? 'Từ' : 'Giá trị';
  };
  type.onchange = op.onchange = sync;
  body.append(h('p', { class: 'muted' }, `Áp dụng cho: ${rangeToA1(g)}`), field('Cho phép', type), fList, h('div', { class: 'row2' }, fOp), h('div', { class: 'row2' }, fV1, fV2), h('label', { class: 'chk' }, blank, ' Cho phép để trống'), field('Câu báo khi nhập sai', err), field('Gợi ý khi chọn ô', prompt));
  sync();
  const btns: DlgButton[] = [{ label: 'Hủy', act: () => {} }];
  if (cur) btns.unshift({ label: 'Xóa kiểm tra', kind: 'danger', act: () => { wb.clearDV(s, g); onDone('Đã xóa kiểm tra dữ liệu.'); } });
  btns.push({ label: 'Lưu', kind: 'primary', act: () => {
    const rule: DVRule = { range: g, type: type.value as DVRule['type'], allowBlank: blank.checked, error: err.value.trim() || undefined, prompt: prompt.value.trim() || undefined };
    if (rule.type === 'list') {
      const t = list.value.trim();
      if (!t) { list.focus(); return false; }
      if (t.startsWith('=')) {
        const ref = t.slice(1).replace(/\$/g, '');
        if (!parseRange(ref.split('!').pop()!)) { list.focus(); onDone('Vùng ô không hợp lệ. Ví dụ: =$H$2:$H$20'); return false; }
        rule.listRef = t.slice(1);
      } else rule.list = t.split(/[;\n]/).map((x) => x.trim()).filter(Boolean);
    } else {
      rule.op = op.value as DVRule['op']; rule.v1 = v1.value.trim(); rule.v2 = v2.value.trim() || undefined;
      if (!rule.v1 || (['between', 'notBetween'].includes(rule.op!) && !rule.v2)) { (rule.v1 ? v2 : v1).focus(); return false; }
    }
    wb.addDV(s, rule);
    onDone(`Đã đặt kiểm tra: ${rule.error || dvDescribe(rule)}`);
  } });
  openDialog('Kiểm tra dữ liệu', body, btns);
}
