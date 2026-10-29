// Bản 0.6 — các hộp thoại đi kèm phím tắt, theo Excel 365:
// Tìm & Thay thế (Ctrl+F/H, không chặn trang tính), Đi tới (F5/Ctrl+G), Đi tới Đặc biệt, Dán Đặc biệt (Ctrl+Alt+V),
// Chèn Liên kết (Ctrl+K), Chuỗi…, Đối số Hàm (Ctrl+A), Phân tích Nhanh (Ctrl+Q), Bảng tạm Office, bảng phím tắt (F1).
import { parseRange, rangeToA1, toA1, colName, type Range } from '../core/address';
import { findAll, findNext, replaceAll, replaceText, type FindOpts, type SpecialOpts } from '../core/find';
import type { PasteOpts } from '../core/workbook';
import { quoteSheet } from '../core/formula';
import { openDialog, closeDialog, h, field } from './dialogs';
import { ic } from './icons';
import { KEYMAP } from './keymap';
import { FUNC_INFO } from './funcs';
import type { App } from './app';

// ================= Tìm & Thay thế =================
export interface FindState { opts: FindOpts; repl: string }
export const findState: FindState = { opts: { what: '', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false }, repl: '' };
let frPanel: HTMLElement | null = null;
export const findPanelOpen = () => !!frPanel;
export function closeFind() { frPanel?.remove(); frPanel = null; }

export function findDialog(app: App, mode: 'find' | 'replace') {
  closeFind();
  const o = findState.opts;
  // chọn sẵn chữ trong ô đang chọn nếu ô tìm đang trống (như Excel lấy từ lần tìm trước)
  const p = h('div', { class: 'fr-panel', role: 'dialog', 'aria-label': 'Tìm và Thay thế', id: 'fr' });
  const head = h('header', {}, h('h2', {}, 'Tìm và Thay thế'));
  const x = h('button', { class: 'dlg-x', 'aria-label': 'Đóng' }); x.innerHTML = ic('x', 18); x.onclick = () => { closeFind(); app.grid.focusKeys(); };
  head.append(x);
  const tabs = h('div', { class: 'ps-tabs', role: 'tablist' });
  const tFind = h('button', { role: 'tab', 'data-t': 'find' }, 'Tìm'), tRep = h('button', { role: 'tab', 'data-t': 'replace' }, 'Thay thế');
  tabs.append(tFind, tRep);
  const what = h('input', { type: 'text', id: 'fr-what', spellcheck: 'false' }) as HTMLInputElement; what.value = o.what;
  const repl = h('input', { type: 'text', id: 'fr-repl', spellcheck: 'false' }) as HTMLInputElement; repl.value = findState.repl;
  const replRow = field('Thay bằng:', repl);
  const within = h('select', { id: 'fr-within' }, h('option', { value: 'sheet' }, 'Trang tính'), h('option', { value: 'book' }, 'Sổ làm việc')) as HTMLSelectElement; within.value = o.within;
  const by = h('select', { id: 'fr-by' }, h('option', { value: 'rows' }, 'Theo Hàng'), h('option', { value: 'cols' }, 'Theo Cột')) as HTMLSelectElement; by.value = o.by;
  const look = h('select', { id: 'fr-look' }, h('option', { value: 'formulas' }, 'Công thức'), h('option', { value: 'values' }, 'Giá trị'), h('option', { value: 'notes' }, 'Ghi chú')) as HTMLSelectElement; look.value = o.lookIn;
  const mc = h('input', { type: 'checkbox', id: 'fr-case' }) as HTMLInputElement; mc.checked = o.matchCase;
  const wh = h('input', { type: 'checkbox', id: 'fr-whole' }) as HTMLInputElement; wh.checked = o.whole;
  const more = h('div', { class: 'fr-more' }, field('Trong:', within), field('Tìm kiếm:', by), field('Tìm trong:', look), h('label', { class: 'chk' }, mc, ' Khớp chữ hoa/thường'), h('label', { class: 'chk' }, wh, ' Khớp toàn bộ nội dung ô'));
  more.hidden = !(o.matchCase || o.whole || o.within !== 'sheet' || o.by !== 'rows' || o.lookIn !== 'formulas');
  const optBtn = h('button', { class: 'btn', id: 'fr-opt' }, more.hidden ? 'Tùy chọn >>' : 'Tùy chọn <<');
  optBtn.onclick = () => { more.hidden = !more.hidden; optBtn.textContent = more.hidden ? 'Tùy chọn >>' : 'Tùy chọn <<'; };
  const msg = h('p', { class: 'fr-msg', role: 'status', id: 'fr-msg' });
  const results = h('div', { class: 'fr-results', id: 'fr-results' });
  const read = (): FindOpts => {
    const n: FindOpts = { what: what.value, within: within.value as FindOpts['within'], by: by.value as FindOpts['by'], lookIn: look.value as FindOpts['lookIn'], matchCase: mc.checked, whole: wh.checked };
    findState.opts = n; findState.repl = repl.value;
    return n;
  };
  const scope = () => (app.grid.isSingle() && !app.grid.areas.length ? undefined : app.grid.ranges());
  const notFound = () => { msg.textContent = 'Không tìm thấy dữ liệu bạn đang tìm.'; };
  const goHit = (hh: { s: number; r: number; c: number }) => { app.gotoCell(hh.s, hh.r, hh.c, true); };
  const btn = (id: string, label: string, act: () => void, primary = false) => { const b = h('button', { class: 'btn' + (primary ? ' primary' : ''), id }, label); b.onclick = act; return b; };
  const doNext = (dir: 1 | -1 = 1) => {
    const op = read(); msg.textContent = ''; results.innerHTML = '';
    if (!op.what) return;
    const { sel } = app.grid;
    const hit = findNext(app.wb, app.grid.sheet, { s: app.grid.sheet, r: sel.ar, c: sel.ac }, op, dir, scope());
    if (!hit) { notFound(); return; }
    goHit(hit);
  };
  const doAll = () => {
    const op = read(); results.innerHTML = '';
    if (!op.what) return;
    const hits = findAll(app.wb, app.grid.sheet, op, scope());
    msg.textContent = hits.length ? `Tìm thấy ${hits.length} ô.` : 'Không tìm thấy dữ liệu bạn đang tìm.';
    if (!hits.length) return;
    const t = h('table', {}, h('thead', {}, h('tr', {}, h('th', {}, 'Trang tính'), h('th', {}, 'Ô'), h('th', {}, 'Giá trị'), h('th', {}, 'Công thức'))));
    const tb = h('tbody');
    for (const x0 of hits.slice(0, 2000)) {
      const raw = app.wb.rawText(x0.s, x0.r, x0.c);
      const tr = h('tr', { tabindex: '0' }, h('td', {}, app.wb.sheetName(x0.s)), h('td', {}, '$' + colName(x0.c) + '$' + (x0.r + 1)), h('td', {}, app.wb.display(x0.s, x0.r, x0.c).text), h('td', {}, raw.startsWith('=') ? raw : ''));
      tr.onclick = () => goHit(x0);
      tr.onkeydown = (e) => { if (e.key === 'Enter') goHit(x0); };
      tb.append(tr);
    }
    t.append(tb); results.append(t);
    goHit(hits[0]);
  };
  const doReplace = () => {
    const op = read(); msg.textContent = '';
    if (!op.what) return;
    const { sel } = app.grid, s = app.grid.sheet;
    // thay ô hiện hành nếu khớp rồi tìm tiếp (như Excel)
    const raw = app.wb.rawText(s, sel.ar, sel.ac);
    const t = raw ? replaceText(app.wb, s, raw, op, repl.value) : null;
    if (t !== null) app.wb.writeCells(s, [[sel.ar, sel.ac, t]]);
    const hit = findNext(app.wb, s, { s, r: sel.ar, c: sel.ac }, { ...op, lookIn: 'formulas' }, 1, scope());
    if (hit) goHit(hit); else if (t === null) notFound();
    app.afterEdit();
  };
  const doReplaceAll = () => {
    const op = read();
    if (!op.what) return;
    const n = replaceAll(app.wb, app.grid.sheet, op, repl.value, scope());
    msg.textContent = n ? `Đã xong. Đã thay thế ${n} mục.` : 'Không tìm thấy gì để thay thế. Hãy kiểm tra cài đặt tìm kiếm.';
    app.afterEdit();
  };
  const footR = h('div', { class: 'fr-foot' }, btn('fr-repall', 'Thay thế Tất cả', doReplaceAll), btn('fr-rep', 'Thay thế', doReplace), btn('fr-all', 'Tìm Tất cả', doAll), btn('fr-next', 'Tìm Tiếp', () => doNext(1), true), btn('fr-close', 'Đóng', () => { closeFind(); app.grid.focusKeys(); }));
  const setMode = (m: 'find' | 'replace') => {
    tFind.setAttribute('aria-selected', String(m === 'find')); tRep.setAttribute('aria-selected', String(m === 'replace'));
    replRow.hidden = m === 'find';
    (footR.querySelector('#fr-repall') as HTMLElement).hidden = m === 'find';
    (footR.querySelector('#fr-rep') as HTMLElement).hidden = m === 'find';
    p.dataset.mode = m;
  };
  tFind.onclick = () => setMode('find'); tRep.onclick = () => setMode('replace');
  p.append(head, tabs, h('div', { class: 'fr-body stack' }, field('Tìm gì:', what), replRow, h('div', { class: 'row' }, optBtn), more, msg, results), footR);
  p.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeFind(); app.grid.focusKeys(); }
    else if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT') { e.preventDefault(); e.stopPropagation(); doNext(e.shiftKey ? -1 : 1); }
    else if (e.key === 'Tab' && e.ctrlKey) { e.preventDefault(); e.stopPropagation(); setMode(p.dataset.mode === 'find' ? 'replace' : 'find'); }
    e.stopPropagation();
  });
  setMode(mode);
  document.body.append(p);
  frPanel = p;
  // kéo thanh tiêu đề để di chuyển
  head.onpointerdown = (e) => {
    if ((e.target as HTMLElement).closest('button')) return;
    const r = p.getBoundingClientRect(), x0 = e.clientX, y0 = e.clientY;
    const mv = (ev: PointerEvent) => { p.style.left = Math.max(0, r.left + ev.clientX - x0) + 'px'; p.style.top = Math.max(0, r.top + ev.clientY - y0) + 'px'; p.style.right = 'auto'; };
    const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); };
    document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
  };
  setTimeout(() => { what.focus(); what.select(); }, 20);
}
/** Shift+F4 / Ctrl+Shift+F4: lặp lần tìm cuối */
export function findAgain(app: App, dir: 1 | -1) {
  const o = findState.opts;
  if (!o.what) { findDialog(app, 'find'); return; }
  const { sel } = app.grid, s = app.grid.sheet;
  const hit = findNext(app.wb, s, { s, r: sel.ar, c: sel.ac }, o, dir);
  if (!hit) { app.toast('Không tìm thấy dữ liệu bạn đang tìm.'); return; }
  app.gotoCell(hit.s, hit.r, hit.c, true);
}

// ================= Đi tới / Đi tới Đặc biệt =================
const recent: string[] = [];
export function gotoDialog(app: App) {
  const list = h('div', { class: 'fc-list', role: 'listbox', 'aria-label': 'Đi tới', id: 'goto-list' });
  const ref = h('input', { type: 'text', id: 'goto-ref', spellcheck: 'false' }) as HTMLInputElement;
  for (const r0 of recent) { const b = h('button', {}, r0); b.onclick = () => { ref.value = r0; }; b.ondblclick = () => { ref.value = r0; go(); closeDialog(); }; list.append(b); }
  if (!recent.length) list.append(h('p', { class: 'muted' }, '(Chưa có — tên vùng có ở bản 0.7)'));
  const go = (): boolean => {
    const t = ref.value.trim();
    if (!t) return false;
    const r = app.gotoRef(t);
    if (!r) { app.toast('Tham chiếu không hợp lệ.', 'err'); return false; }
    const i = recent.indexOf(t); if (i >= 0) recent.splice(i, 1); recent.unshift(t); recent.splice(8);
    return true;
  };
  const sp = h('button', { class: 'btn', id: 'goto-special' }, 'Đặc biệt…');
  sp.onclick = () => { closeDialog(); gotoSpecialDialog(app); };
  openDialog('Đi tới', h('div', { class: 'stack' }, h('span', { class: 'muted' }, 'Đi tới:'), list, field('Tham chiếu:', ref), h('div', { class: 'row' }, sp)), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => go() },
  ], () => app.grid.focusKeys());
}
export function gotoSpecialDialog(app: App) {
  const opts: [SpecialOpts['kind'] | 'array', string][] = [
    ['notes', 'Ghi chú'], ['constants', 'Hằng số'], ['formulas', 'Công thức'], ['blanks', 'Ô trống'], ['region', 'Vùng hiện tại'], ['array', 'Mảng hiện tại'], ['objects', 'Đối tượng'],
    ['rowDiff', 'Khác biệt trong hàng'], ['colDiff', 'Khác biệt trong cột'], ['precedents', 'Ô tiền lệ'], ['dependents', 'Ô phụ thuộc'], ['lastcell', 'Ô cuối cùng'], ['visible', 'Chỉ ô hiển thị'], ['cf', 'Định dạng có điều kiện'], ['dv', 'Xác thực dữ liệu'],
  ];
  const box = h('div', { class: 'gs' });
  const radios = h('div', { class: 'gs-grid' });
  for (const [v, l] of opts) {
    const i = h('input', { type: 'radio', name: 'gs', value: v, id: 'gs-' + v }) as HTMLInputElement;
    if (v === 'constants') i.checked = true;
    if (v === 'array') i.disabled = true;
    radios.append(h('label', { class: 'chk' }, i, ' ' + l));
  }
  const ty = ['num', 'text', 'bool', 'err'].map((k) => { const i = h('input', { type: 'checkbox', id: 'gs-t-' + k }) as HTMLInputElement; i.checked = true; return i; });
  const types = h('div', { class: 'gs-sub' }, h('label', { class: 'chk' }, ty[0], ' Số'), h('label', { class: 'chk' }, ty[1], ' Văn bản'), h('label', { class: 'chk' }, ty[2], ' Lôgic'), h('label', { class: 'chk' }, ty[3], ' Lỗi'));
  const lv = h('div', { class: 'gs-sub' });
  const direct = h('input', { type: 'radio', name: 'gs-lv', value: 'd', id: 'gs-direct' }) as HTMLInputElement; direct.checked = true;
  const all = h('input', { type: 'radio', name: 'gs-lv', value: 'a', id: 'gs-all' }) as HTMLInputElement;
  lv.append(h('label', { class: 'chk' }, direct, ' Chỉ trực tiếp'), h('label', { class: 'chk' }, all, ' Mọi cấp'));
  box.append(h('span', { class: 'muted' }, 'Chọn'), radios, h('span', { class: 'muted' }, 'Công thức / Hằng số gồm:'), types, h('span', { class: 'muted' }, 'Tiền lệ / Phụ thuộc:'), lv);
  openDialog('Đi tới Đặc biệt', box, [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      const kind = (box.querySelector('input[name=gs]:checked') as HTMLInputElement).value as SpecialOpts['kind'];
      const t = { num: ty[0].checked, text: ty[1].checked, bool: ty[2].checked, err: ty[3].checked };
      return app.selectSpecial({ kind, types: kind === 'constants' || kind === 'formulas' ? t : undefined, all: all.checked });
    } },
  ], () => app.grid.focusKeys());
}

// ================= Dán Đặc biệt =================
const PS_WHAT: [NonNullable<PasteOpts['what']>, string, string][] = [
  ['all', 'Tất cả', 'A'], ['formulas', 'Công thức', 'F'], ['values', 'Giá trị', 'V'], ['formats', 'Định dạng', 'T'], ['comments', 'Chú thích và ghi chú', 'C'], ['validation', 'Xác thực', 'N'],
  ['theme', 'Tất cả dùng chủ đề nguồn', 'H'], ['noBorders', 'Tất cả trừ viền', 'X'], ['colWidths', 'Độ rộng cột', 'W'], ['formulasNum', 'Công thức và định dạng số', 'R'], ['valuesNum', 'Giá trị và định dạng số', 'U'], ['allMerge', 'Tất cả định dạng có điều kiện hợp nhất', 'G'],
];
const PS_OP: [NonNullable<PasteOpts['op']>, string, string][] = [['none', 'Không', 'O'], ['add', 'Cộng', 'D'], ['sub', 'Trừ', 'S'], ['mul', 'Nhân', 'M'], ['div', 'Chia', 'I']];
export function pasteSpecialDialog(app: App, ext: string | null) {
  if (!app.wb.clip) {
    // nguồn từ ứng dụng khác: chỉ dán được dạng văn bản
    if (!ext) { app.toast('Chưa có gì được sao chép.'); return; }
    const box = h('div', { class: 'stack' }, h('span', { class: 'muted' }, 'Nguồn: bảng tạm của Windows'), h('label', { class: 'chk' }, h('input', { type: 'radio', name: 'pt', checked: '' }), ' Văn bản Unicode'), h('label', { class: 'chk' }, h('input', { type: 'radio', name: 'pt' }), ' Văn bản'));
    openDialog('Dán Đặc biệt', box, [{ label: 'Hủy', act: () => {} }, { label: 'OK', kind: 'primary', act: () => { app.pasteTextAt(ext); } }], () => app.grid.focusKeys());
    return;
  }
  if (app.wb.clip.cut) { app.toast('Dán Đặc biệt chỉ dùng được sau khi Sao chép (Ctrl+C), không dùng với Cắt.', 'err'); return; }
  const box = h('div', { class: 'psd' });
  const grp = (name: string, list: [string, string, string][], def: string, cls: string) => {
    const g = h('div', { class: 'psd-grid ' + cls });
    for (const [v, l, k] of list) {
      const i = h('input', { type: 'radio', name, value: v, id: `ps-${name}-${v}`, 'data-k': k }) as HTMLInputElement;
      if (v === def) i.checked = true;
      g.append(h('label', { class: 'chk', for: i.id }, i, ` ${l} (${k})`));
    }
    return g;
  };
  const skip = h('input', { type: 'checkbox', id: 'ps-skip', 'data-k': 'B' }) as HTMLInputElement;
  const tr = h('input', { type: 'checkbox', id: 'ps-transpose', 'data-k': 'E' }) as HTMLInputElement;
  box.append(h('h4', {}, 'Dán'), grp('w', PS_WHAT, 'all', 'two'), h('h4', {}, 'Phép toán'), grp('op', PS_OP, 'none', 'two'),
    h('div', { class: 'row' }, h('label', { class: 'chk' }, skip, ' Bỏ qua ô trống (B)'), h('label', { class: 'chk' }, tr, ' Chuyển vị (E)')));
  const read = (): PasteOpts => ({
    what: (box.querySelector('input[name=w]:checked') as HTMLInputElement).value as PasteOpts['what'],
    op: (box.querySelector('input[name=op]:checked') as HTMLInputElement).value as PasteOpts['op'],
    skipBlanks: skip.checked, transpose: tr.checked,
  });
  // phím chữ chọn mục như Excel (A F V T C N H X W R U G · O D S M I · B E · L)
  box.addEventListener('keydown', (e) => {
    if (e.ctrlKey || e.altKey || e.metaKey || e.key.length !== 1) return;
    const k = e.key.toUpperCase();
    if (k === 'L') { e.preventDefault(); closeDialog(); app.pasteSpecialRun({ link: true }); return; }
    if (k === 'B') { e.preventDefault(); skip.checked = !skip.checked; return; }
    if (k === 'E') { e.preventDefault(); tr.checked = !tr.checked; return; }
    const i = box.querySelector(`input[type=radio][data-k="${k}"]`) as HTMLInputElement | null;
    if (i) { e.preventDefault(); i.checked = true; i.focus(); }
  });
  const dlg = openDialog('Dán Đặc biệt', box, [
    { label: 'Dán Liên kết', act: () => { app.pasteSpecialRun({ link: true }); } },
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => { app.pasteSpecialRun(read()); } },
  ], () => app.grid.focusKeys());
  dlg.querySelector('.dlg')?.classList.add('wide');
  dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT') { e.preventDefault(); e.stopPropagation(); closeDialog(); app.pasteSpecialRun(read()); } }, true);
  setTimeout(() => (box.querySelector('input[name=w]:checked') as HTMLElement | null)?.focus(), 40);
}

// ================= Liên kết (Ctrl+K) =================
export function hyperlinkDialog(app: App) {
  const s = app.grid.sheet, { ar, ac } = app.grid.sel;
  const st = app.wb.style(s, ar, ac);
  const curText = app.wb.display(s, ar, ac).text;
  const kinds: [string, string][] = [['web', 'Tệp hoặc Trang web'], ['place', 'Vị trí trong Tài liệu này'], ['mail', 'Địa chỉ Email']];
  let kind = st.ln?.startsWith('#') ? 'place' : st.ln?.startsWith('mailto:') ? 'mail' : 'web';
  const text = h('input', { type: 'text', id: 'hl-text' }) as HTMLInputElement; text.value = curText;
  const addr = h('input', { type: 'text', id: 'hl-addr', placeholder: 'https://…', spellcheck: 'false' }) as HTMLInputElement;
  const sheetSel = h('select', { id: 'hl-sheet' }) as HTMLSelectElement;
  for (const id of app.wb.visibleSheetIds()) sheetSel.append(h('option', { value: app.wb.sheetName(id) }, app.wb.sheetName(id)));
  const cell = h('input', { type: 'text', id: 'hl-cell', value: 'A1' }) as HTMLInputElement;
  const mail = h('input', { type: 'text', id: 'hl-mail', placeholder: 'ten@vidu.vn' }) as HTMLInputElement;
  const subj = h('input', { type: 'text', id: 'hl-subj' }) as HTMLInputElement;
  if (st.ln) {
    if (kind === 'place') { const m = /^#(?:'?(.+?)'?!)?(.+)$/.exec(st.ln); if (m) { if (m[1]) sheetSel.value = m[1].replace(/''/g, "'"); cell.value = m[2]; } }
    else if (kind === 'mail') { const m = /^mailto:([^?]*)(?:\?subject=(.*))?$/.exec(st.ln); mail.value = decodeURIComponent(m?.[1] ?? ''); subj.value = decodeURIComponent(m?.[2] ?? ''); }
    else addr.value = st.ln;
  }
  const panes: Record<string, HTMLElement> = {
    web: h('div', { class: 'stack' }, field('Địa chỉ:', addr)),
    place: h('div', { class: 'stack' }, field('Nhập tham chiếu ô:', cell), field('Hoặc chọn trang tính:', sheetSel)),
    mail: h('div', { class: 'stack' }, field('Địa chỉ email:', mail), field('Tiêu đề:', subj)),
  };
  const side = h('div', { class: 'hl-side', role: 'tablist' });
  const show = () => { for (const [k] of kinds) panes[k].hidden = k !== kind; side.querySelectorAll('button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.k === kind))); };
  for (const [k, l] of kinds) { const b = h('button', { role: 'tab', 'data-k': k }, l); b.onclick = () => { kind = k; show(); }; side.append(b); }
  const body = h('div', { class: 'hl' }, side, h('div', { class: 'stack' }, field('Văn bản hiển thị:', text), panes.web, panes.place, panes.mail));
  show();
  const btns = [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary' as const, act: () => {
      let url = '';
      if (kind === 'web') url = addr.value.trim();
      else if (kind === 'place') url = '#' + quoteSheet(sheetSel.value) + '!' + (cell.value.trim() || 'A1');
      else url = mail.value.trim() ? 'mailto:' + mail.value.trim() + (subj.value ? '?subject=' + encodeURIComponent(subj.value) : '') : '';
      if (!url) { app.toast('Hãy nhập địa chỉ liên kết.', 'err'); return false; }
      if (kind === 'web' && !/^[a-z][\w+.-]*:|^\\\\|^[A-Za-z]:\\/.test(url)) url = 'http://' + url;
      app.setLink(url, text.value);
    } },
  ];
  if (st.ln) btns.unshift({ label: 'Gỡ Liên kết', act: () => { app.setLink(null, null); } } as any);
  openDialog(st.ln ? 'Sửa Siêu kết nối' : 'Chèn Siêu kết nối', body, btns, () => app.grid.focusKeys()).querySelector('.dlg')?.classList.add('wide');
}

// ================= Chuỗi… =================
export function seriesDialog(app: App) {
  const g = app.grid.range;
  const dirR = h('input', { type: 'radio', name: 'sd', value: 'rows', id: 'sr-rows' }) as HTMLInputElement;
  const dirC = h('input', { type: 'radio', name: 'sd', value: 'cols', id: 'sr-cols' }) as HTMLInputElement;
  (g.c2 - g.c1 > g.r2 - g.r1 ? dirR : dirC).checked = true;
  const types = (['linear', 'growth', 'date', 'auto'] as const).map((v) => h('input', { type: 'radio', name: 'st', value: v, id: 'sr-' + v }) as HTMLInputElement);
  const isDate = /[dy]/i.test(app.wb.effFmt(app.grid.sheet, g.r1, g.c1) ?? '');
  types[isDate ? 2 : 0].checked = true;
  const units = (['day', 'weekday', 'month', 'year'] as const).map((v) => h('input', { type: 'radio', name: 'su', value: v, id: 'sr-u-' + v }) as HTMLInputElement);
  units[0].checked = true;
  const trend = h('input', { type: 'checkbox', id: 'sr-trend' }) as HTMLInputElement;
  const step = h('input', { type: 'text', id: 'sr-step', value: '1', inputmode: 'decimal' }) as HTMLInputElement;
  const stop = h('input', { type: 'text', id: 'sr-stop', inputmode: 'decimal' }) as HTMLInputElement;
  const L = (i: HTMLInputElement, t: string) => h('label', { class: 'chk' }, i, ' ' + t);
  const unitBox = h('fieldset', { class: 'sr-box' }, h('legend', {}, 'Đơn vị ngày'), L(units[0], 'Ngày'), L(units[1], 'Ngày trong tuần'), L(units[2], 'Tháng'), L(units[3], 'Năm'));
  const syncUnits = () => { unitBox.disabled = !types[2].checked; };
  types.forEach((t) => (t.onchange = syncUnits)); syncUnits();
  const body = h('div', { class: 'sr' },
    h('fieldset', { class: 'sr-box' }, h('legend', {}, 'Chuỗi trong'), L(dirR, 'Hàng'), L(dirC, 'Cột')),
    h('fieldset', { class: 'sr-box' }, h('legend', {}, 'Loại'), L(types[0], 'Tuyến tính'), L(types[1], 'Tăng trưởng'), L(types[2], 'Ngày'), L(types[3], 'Tự động Điền')),
    unitBox, h('label', { class: 'chk' }, trend, ' Xu hướng'), field('Giá trị bước:', step), field('Giá trị dừng:', stop));
  openDialog('Chuỗi', body, [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      const num = (t: string) => { const n = parseFloat(t.replace(/\./g, '').replace(',', '.')); return isNaN(n) ? undefined : n; };
      const type = (types.find((x) => x.checked)!.value) as 'linear' | 'growth' | 'date' | 'auto';
      return app.fillSeries({ dir: dirR.checked ? 'rows' : 'cols', type, unit: units.find((x) => x.checked)!.value as any, step: num(step.value) ?? 1, stop: stop.value.trim() ? num(stop.value) : undefined, trend: trend.checked });
    } },
  ], () => app.grid.focusKeys());
}

// ================= Đối số Hàm (Ctrl+A sau tên hàm) =================
export function argNames(fn: string): string[] {
  const sig = FUNC_INFO[fn.toUpperCase()]?.[1];
  if (!sig) return [];
  const inner = sig.slice(sig.indexOf('(') + 1, sig.lastIndexOf(')'));
  return inner.split(';').map((x) => x.trim()).filter((x) => x && x !== '…');
}
export function funcArgsDialog(app: App, fn: string, onDone: (text: string) => void) {
  const info = FUNC_INFO[fn.toUpperCase()];
  const names = argNames(fn);
  const list = names.length ? names : ['đối số 1', 'đối số 2'];
  const box = h('div', { class: 'stack fa' });
  const inputs: HTMLInputElement[] = [];
  const result = h('p', { class: 'fa-res', id: 'fa-res' });
  const calc = () => {
    const args = inputs.map((i) => i.value.trim());
    while (args.length && !args[args.length - 1]) args.pop();
    const f = `=${fn}(${args.join(';')})`;
    const v = app.wb.evalFormula(f, app.grid.sheet);
    result.textContent = 'Kết quả công thức = ' + (v === undefined ? '' : String(v));
    return f;
  };
  for (const n of list) {
    const i = h('input', { type: 'text', spellcheck: 'false', 'aria-label': n }) as HTMLInputElement;
    i.oninput = calc;
    inputs.push(i);
    box.append(field(n.replace(/[[\]]/g, '') + (n.startsWith('[') ? ' (tùy chọn)' : ''), i));
  }
  box.prepend(h('p', { class: 'muted' }, info ? `${info[1]} — ${info[2]}` : fn));
  box.append(result);
  openDialog('Đối số Hàm', box, [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => { onDone(calc()); } },
  ]);
  calc();
}

// ================= Phân tích Nhanh (Ctrl+Q) =================
export function quickAnalysis(app: App) {
  const g = app.grid.range;
  const r = app.grid.selectionRect();
  const box = h('div', { class: 'qa', id: 'qa', role: 'dialog', 'aria-label': 'Phân tích Nhanh' });
  const tabs = h('div', { class: 'ps-tabs', role: 'tablist' });
  const pane = h('div', { class: 'qa-pane' });
  const item = (icon: string, label: string, act: () => void) => { const b = h('button', { class: 'qa-it' }); b.innerHTML = ic(icon, 22); b.append(h('span', {}, label)); b.onclick = () => { closeQA(); act(); }; return b; };
  const groups: [string, () => HTMLElement[]][] = [
    ['Định dạng', () => [
      item('cf', 'Thanh Dữ liệu', () => app.quickCF('bar')), item('palette', 'Thang Màu', () => app.quickCF('scale')),
      item('cf', 'Lớn hơn', () => app.quickCF('gt')), item('cf', '10% Trên cùng', () => app.quickCF('top')), item('x', 'Xóa Định dạng', () => app.quickCF('clear'))]],
    ['Biểu đồ', () => [item('chartCol', 'Cột', () => app.insertChart('col')), item('chartLine', 'Đường', () => app.insertChart('line')), item('chartPie', 'Tròn', () => app.insertChart('pie')), item('chartBar', 'Thanh', () => app.insertChart('bar'))]],
    ['Tổng', () => [item('sigma', 'Tổng hàng', () => app.quickTotal('SUM', 'row')), item('sigma', 'Trung bình', () => app.quickTotal('AVERAGE', 'row')), item('sigma', 'Đếm', () => app.quickTotal('COUNT', 'row')), item('sigma', 'Tổng cột', () => app.quickTotal('SUM', 'col'))]],
    ['Bảng', () => [item('pivot', 'PivotTable', () => app.insertPivot())]],
  ];
  const show = (i: number) => { pane.innerHTML = ''; pane.append(...groups[i][1]()); tabs.querySelectorAll('button').forEach((b, k) => b.setAttribute('aria-selected', String(k === i))); };
  groups.forEach(([l], i) => { const b = h('button', { role: 'tab' }, l); b.onclick = () => show(i); tabs.append(b); });
  box.append(tabs, pane);
  show(0);
  closeQA();
  document.body.append(box);
  qaEl = box;
  const vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
  const bw = box.offsetWidth, bh = box.offsetHeight;
  box.style.left = Math.max(4, Math.min(vw - bw - 4, r.right)) + 'px';
  box.style.top = Math.max(4, Math.min(vh - bh - 4, r.bottom + 4)) + 'px';
  box.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); closeQA(); app.grid.focusKeys(); } e.stopPropagation(); });
  setTimeout(() => {
    (pane.querySelector('button') as HTMLElement | null)?.focus();
    document.addEventListener('pointerdown', qaOutside, true);
  });
  void g;
}
let qaEl: HTMLElement | null = null;
function qaOutside(e: Event) { if (qaEl && !qaEl.contains(e.target as Node)) closeQA(); }
export function closeQA() { qaEl?.remove(); qaEl = null; document.removeEventListener('pointerdown', qaOutside, true); }

// ================= Bảng tạm Office =================
export interface ClipItem { text: string; at: number }
export const clipRing: ClipItem[] = [];
export function pushClip(text: string) {
  if (!text) return;
  const i = clipRing.findIndex((x) => x.text === text);
  if (i >= 0) clipRing.splice(i, 1);
  clipRing.unshift({ text, at: Date.now() });
  clipRing.splice(24);
  renderClipPane();
}
let clipPane: HTMLElement | null = null;
let clipApp: App | null = null;
export function toggleClipboardPane(app: App) {
  if (clipPane) { clipPane.remove(); clipPane = null; return; }
  clipApp = app;
  clipPane = h('aside', { class: 'clip-pane', id: 'clip-pane', 'aria-label': 'Bảng tạm' });
  document.body.append(clipPane);
  renderClipPane();
}
function renderClipPane() {
  const p = clipPane, app = clipApp;
  if (!p || !app) return;
  p.innerHTML = '';
  const x = h('button', { class: 'dlg-x', 'aria-label': 'Đóng' }); x.innerHTML = ic('x', 16); x.onclick = () => { p.remove(); clipPane = null; };
  const clearAll = h('button', { class: 'btn' }, 'Xóa Tất cả'); clearAll.onclick = () => { clipRing.length = 0; renderClipPane(); };
  p.append(h('header', {}, h('h3', {}, 'Bảng tạm'), x), h('div', { class: 'row' }, clearAll), h('p', { class: 'muted' }, clipRing.length ? 'Bấm một mục để dán:' : 'Bảng tạm trống. Sao chép hoặc cắt để thu thập mục.'));
  for (const it of clipRing) {
    const b = h('button', { class: 'clip-it' }, it.text.length > 120 ? it.text.slice(0, 120) + '…' : it.text);
    b.onclick = () => app.pasteTextAt(it.text);
    p.append(b);
  }
}

// ================= Bảng phím tắt (F1) =================
export function keysDialog(app: App) {
  const q = h('input', { type: 'search', placeholder: 'Tìm phím hoặc lệnh…', id: 'keys-q' }) as HTMLInputElement;
  const list = h('div', { class: 'keys-list', id: 'keys-list' });
  const fold = (t: string) => t.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
  const draw = () => {
    list.innerHTML = '';
    const f = fold(q.value.trim());
    let grp = '';
    for (const k of KEYMAP) {
      if (f && !fold(k.label + ' ' + k.desc).includes(f)) continue;
      if (k.group !== grp) { grp = k.group; list.append(h('h4', {}, grp)); }
      const row = h('div', { class: 'keys-row' + (k.status !== 'ok' ? ' off' : '') }, h('kbd', {}, k.label), h('span', {}, k.desc));
      if (k.status !== 'ok') row.append(h('small', {}, k.status === 'later' ? `có ở bản ${k.note}` : `không làm: ${k.note}`));
      list.append(row);
    }
  };
  q.oninput = draw;
  draw();
  const n = KEYMAP.filter((k) => k.status === 'ok').length;
  openDialog('Phím tắt', h('div', { class: 'stack' }, h('p', { class: 'muted' }, `Giống Excel 365 cho Windows — ${n}/${KEYMAP.length} mục đã có.`), q, list), [{ label: 'Đóng', kind: 'primary', act: () => {} }], () => app.grid.focusKeys())
    .querySelector('.dlg')?.classList.add('wide');
}

export { rangeToA1, toA1, parseRange };
export type { Range };
