// Dữ liệu ▸ Văn bản thành Cột, Loại bỏ Trùng lặp
import { openDialog, h, field, input } from './dialogs';
import { MAX_COLS, MAX_ROWS, colName, parseRange } from '../core/address';
import type { App } from './app';

const chk = (t: string, on = false) => { const c = h('input', { type: 'checkbox' }) as HTMLInputElement; c.checked = on; return { c, el: h('label', { class: 'chk' }, c, ' ' + t) }; };
const after = (app: App) => { app.grid.layout(); app.grid.invalidate(); app.syncSelection(); app.scheduleSave(); };

/** Loại bỏ Trùng lặp */
export function removeDupDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  let g = app.grid.range;
  if (app.grid.isSingle()) g = wb.currentRegion(s, app.grid.sel.ar, app.grid.sel.ac);
  const u = wb.used(s);
  g = { ...g, r2: Math.min(g.r2, Math.max(0, u.rows - 1)), c2: Math.min(g.c2, Math.max(0, u.cols - 1)) };
  if (g.r2 <= g.r1 && g.c2 === g.c1 && wb.value(s, g.r1, g.c1) === null) { app.toast('Hãy chọn vùng dữ liệu cần loại bỏ trùng lặp.', 'err'); return; }
  app.grid.selectRange(g);
  const hdr = chk('Dữ liệu của tôi có tiêu đề', wb.guessHeader(s, g));
  const list = h('div', { class: 'ol-cols' });
  const cbs: { c: HTMLInputElement; col: number; lab: HTMLElement }[] = [];
  for (let c = g.c1; c <= g.c2; c++) { const x = chk('', true); cbs.push({ c: x.c, col: c, lab: x.el }); list.append(x.el); }
  const relabel = () => cbs.forEach((x) => { const t = hdr.c.checked ? wb.display(s, g.r1, x.col).text : ''; x.lab.lastChild!.textContent = ' ' + (t || `Cột ${colName(x.col)}`); });
  relabel();
  hdr.c.onchange = relabel;
  const all = h('button', { class: 'btn' }, 'Chọn Tất cả'), none = h('button', { class: 'btn' }, 'Bỏ chọn Tất cả');
  all.onclick = () => cbs.forEach((x) => (x.c.checked = true));
  none.onclick = () => cbs.forEach((x) => (x.c.checked = false));
  const body = h('div', { class: 'dup-dlg' }, h('p', {}, 'Để xóa giá trị trùng lặp, chọn một hoặc nhiều cột chứa giá trị trùng lặp.'), h('div', { class: 'nm-bar' }, all, none, h('span', { class: 'sp' }), hdr.el), h('div', { class: 'fld' }, h('span', {}, 'Cột'), list));
  openDialog('Loại bỏ Trùng lặp', body, [
    { label: 'Hủy', act: () => {} },
    {
      label: 'OK', kind: 'primary', act: () => {
        const cols = cbs.filter((x) => x.c.checked).map((x) => x.col);
        if (!cols.length) { app.toast('Hãy chọn ít nhất một cột.', 'err'); return false; }
        try {
          const r = wb.removeDuplicates(s, g, cols, hdr.c.checked);
          after(app);
          app.toast(r.removed ? `Đã tìm thấy và xóa ${r.removed} giá trị trùng lặp; còn lại ${r.kept} giá trị duy nhất.` : 'Không tìm thấy giá trị trùng lặp nào.', 'ok');
        } catch (e) { app.toast((e as Error).message, 'err'); }
      },
    },
  ]);
}

/** Văn bản thành Cột (gộp 3 bước của trình hướng dẫn Excel vào một hộp, có xem trước) */
export function textToColumnsDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb, g0 = app.grid.range;
  if (g0.c2 !== g0.c1) { app.toast('Mỗi lần chỉ chuyển đổi được một cột. Hãy chọn một cột.', 'err'); return; }
  const r1 = g0.r1, r2 = Math.min(g0.r2, Math.max(0, wb.used(s).rows - 1)), col = g0.c1;
  if (r2 < r1 || (g0.r2 >= MAX_ROWS - 1 && wb.used(s).rows === 0)) { app.toast('Cột đang chọn không có dữ liệu.', 'err'); return; }
  const modeD = h('input', { type: 'radio', name: 't2c-m' }) as HTMLInputElement, modeF = h('input', { type: 'radio', name: 't2c-m' }) as HTMLInputElement;
  modeD.checked = true;
  const sample = Array.from({ length: Math.min(200, r2 - r1 + 1) }, (_, i) => wb.display(s, r1 + i, col).text).join('\n');
  const guess = sample.includes('\t') ? '\t' : sample.includes(';') ? ';' : sample.includes(',') ? ',' : ' ';
  const dTab = chk('Tab', guess === '\t'), dSemi = chk('Dấu chấm phẩy', guess === ';'), dComma = chk('Dấu phẩy', guess === ','), dSpace = chk('Dấu cách', guess === ' '), dOther = chk('Khác:');
  const other = h('input', { type: 'text', maxlength: '1', class: 'mono t2c-o' }) as HTMLInputElement;
  dOther.el.append(other);
  const cons = chk('Coi dấu phân cách liên tiếp là một');
  const qual = h('select') as HTMLSelectElement;
  for (const [v, t] of [['"', '"'], ["'", "'"], ['', '{không có}']]) qual.append(h('option', { value: v }, t));
  const widths = input('', 'vd: 4; 10; 25');
  widths.classList.add('mono');
  const dest = input('=' + colName(col) + (r1 + 1));
  dest.classList.add('mono');
  const prev = h('table', { class: 'nm-table t2c-prev' });
  let types: string[] = [];
  const opts = () => {
    const delims = (dTab.c.checked ? '\t' : '') + (dSemi.c.checked ? ';' : '') + (dComma.c.checked ? ',' : '') + (dSpace.c.checked ? ' ' : '') + (dOther.c.checked ? other.value : '');
    const w = widths.value.split(/[;,\s]+/).map(Number).filter((x) => x > 0).sort((a, b) => a - b);
    return modeF.checked ? { widths: w, types } : { delims: delims || '\t', consecutive: cons.c.checked, qualifier: qual.value, types };
  };
  const split = (t: string) => {
    // dùng chính hàm tách của lõi trên một sổ tạm để xem trước đúng như kết quả
    const o = opts();
    if ('widths' in o) { const out: string[] = []; let p = 0; for (const w of [...o.widths!, Infinity]) { if (p >= t.length) break; out.push(t.slice(p, w === Infinity ? undefined : w)); if (w === Infinity) break; p = w; } return out; }
    const out: string[] = []; let cur = '', inQ = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (o.qualifier && ch === o.qualifier) { if (inQ && t[i + 1] === ch) { cur += ch; i++; continue; } inQ = !inQ; continue; }
      if (!inQ && o.delims!.includes(ch)) { out.push(cur); cur = ''; if (o.consecutive) while (i + 1 < t.length && o.delims!.includes(t[i + 1])) i++; continue; }
      cur += ch;
    }
    out.push(cur); return out;
  };
  const TY: [string, string][] = [['g', 'Chung'], ['t', 'Văn bản'], ['x', 'Bỏ qua cột']];
  const draw = () => {
    const rows = sample.split('\n').slice(0, 6).map(split);
    const n = Math.max(1, ...rows.map((r) => r.length));
    types = Array.from({ length: n }, (_, i) => types[i] ?? 'g');
    const head = h('tr');
    for (let i = 0; i < n; i++) {
      const sel = h('select', { 'aria-label': `Kiểu cột ${i + 1}` }) as HTMLSelectElement;
      for (const [v, t] of TY) sel.append(h('option', { value: v }, t));
      sel.value = types[i];
      sel.onchange = () => { types[i] = sel.value; draw(); };
      head.append(h('th', {}, sel));
    }
    prev.replaceChildren(h('thead', {}, head), h('tbody', {}, ...rows.map((r) => h('tr', {}, ...Array.from({ length: n }, (_, i) => h('td', { class: types[i] === 'x' ? 'muted' : '' }, r[i] ?? '')))))) ;
    delimBox.style.display = modeF.checked ? 'none' : '';
    fixBox.style.display = modeF.checked ? '' : 'none';
  };
  const delimBox = h('div', { class: 't2c-d' }, h('div', { class: 't2c-row' }, dTab.el, dSemi.el, dComma.el, dSpace.el, dOther.el), h('div', { class: 't2c-row' }, cons.el, field('Bao văn bản', qual)));
  const fixBox = field('Vị trí ngắt cột (số ký tự)', widths);
  for (const x of [modeD, modeF, dTab.c, dSemi.c, dComma.c, dSpace.c, dOther.c, cons.c]) x.addEventListener('change', draw);
  for (const x of [other, widths] as HTMLElement[]) x.addEventListener('input', draw);
  qual.onchange = draw;
  const body = h('div', { class: 't2c-dlg' },
    h('div', { class: 't2c-row' }, h('label', { class: 'chk' }, modeD, ' Phân cách'), h('label', { class: 'chk' }, modeF, ' Độ rộng cố định')),
    delimBox, fixBox, h('div', { class: 'fld' }, h('span', {}, 'Xem trước dữ liệu'), h('div', { class: 'nm-wrap' }, prev)), field('Đích', dest));
  draw();
  openDialog('Văn bản thành Cột', body, [
    { label: 'Hủy', act: () => {} },
    {
      label: 'Hoàn tất', kind: 'primary', act: () => {
        const d = parseRange(dest.value.replace(/^=/, '').replace(/^.*!/, ''));
        if (!d) { app.toast('Ô đích không hợp lệ.', 'err'); return false; }
        const o = { ...opts(), dest: { r: d.r1, c: d.c1 } };
        const n = types.filter((t) => t !== 'x').length;
        // như Excel: cảnh báo khi ghi đè dữ liệu sẵn có (ngoài cột nguồn)
        let busy = false;
        for (let r = r1; r <= r2 && !busy; r++) for (let c = d.c1; c < d.c1 + n && c < MAX_COLS; c++) { if (c === col && d.r1 === r1) continue; if (wb.value(s, d.r1 + r - r1, c) !== null) { busy = true; break; } }
        if (busy && !confirm('Có dữ liệu ở đây rồi. Bạn có muốn thay thế không?')) return false;
        try { wb.textToColumns(s, col, r1, r2, o); after(app); } catch (e) { app.toast((e as Error).message, 'err'); return false; }
      },
    },
  ]);
}
