// Kiểm nghiệm Công thức (0.9): Đánh giá Công thức, Kiểm tra Lỗi (tam giác xanh, nút !, Alt+Shift+F10), Cửa sổ Theo dõi
import { openDialog, closeDialog, h } from './dialogs';
import { showMenu } from './menu';
import { toA1, type Range } from '../core/address';
import { quoteSheet } from '../core/formula';
import { nextStep, applyStep, type Step } from '../core/evalsteps';
import type { App } from './app';

const ERR_TEXT: Record<string, string> = {
  '#DIV/0!': 'Lỗi Chia cho Không', '#N/A': 'Lỗi Không có Giá trị', '#NAME?': 'Lỗi Tên Không hợp lệ', '#NULL!': 'Lỗi Giao Rỗng',
  '#NUM!': 'Lỗi Số', '#REF!': 'Lỗi Tham chiếu Không hợp lệ', '#VALUE!': 'Lỗi Giá trị', '#SPILL!': 'Lỗi Tràn', '#CYCLE!': 'Tham chiếu Vòng', '#ERROR!': 'Lỗi Cú pháp',
};
const ERR_HINT: Record<string, string> = {
  '#DIV/0!': 'Công thức hoặc hàm đang chia cho 0 hoặc cho ô trống.',
  '#N/A': 'Không tìm thấy giá trị (thường từ VLOOKUP/MATCH/XLOOKUP).',
  '#NAME?': 'Công thức có tên hàm hoặc tên vùng không nhận ra.',
  '#REF!': 'Công thức tham chiếu tới ô không còn tồn tại.',
  '#VALUE!': 'Đối số hoặc toán hạng sai kiểu (vd cộng chữ với số).',
  '#NUM!': 'Giá trị số không hợp lệ cho hàm hoặc kết quả quá lớn.',
  '#SPILL!': 'Vùng tràn của mảng động đang có dữ liệu. Hãy xóa các ô cản.',
};
function issueText(app: App, r: number, c: number) {
  const s = app.grid.sheet, kind = app.wb.cellIssue(s, r, c);
  if (kind === 'numText') return { title: 'Số Được Lưu dưới dạng Văn bản', hint: 'Số trong ô này được định dạng là văn bản hoặc đứng sau dấu nháy đơn.' };
  const e = String((app.wb.value(s, r, c) as { value?: string })?.value ?? '');
  return { title: ERR_TEXT[e] ?? 'Lỗi trong Công thức', hint: ERR_HINT[e] ?? 'Công thức trả về lỗi.' };
}
const refreshGrid = (app: App) => { app.grid.invalidate(); app.syncSelection(); app.scheduleSave(); };

/** Đánh giá Công thức: hộp có biểu thức, phần sắp tính gạch chân; Đánh giá / Bước Vào / Bước Ra / Khởi động lại */
export function evaluateDialog(app: App) {
  const wb = app.wb;
  type Frame = { s: number; r: number; c: number; f: string; step: Step | null; done: boolean };
  const stack: Frame[] = [];
  const s0 = app.grid.sheet, { ar, ac } = app.grid.sel;
  const raw = wb.rawText(s0, ar, ac);
  if (!raw.startsWith('=')) { app.toast('Ô hiện hành không chứa công thức.', 'err'); return; }
  if (wb.formulaHidden(s0, ar, ac)) { app.toast('Không thể đánh giá công thức bị ẩn trên trang tính được bảo vệ.', 'err'); return; }
  const push = (s: number, r: number, c: number, f: string) => { const fr: Frame = { s, r, c, f, step: null, done: false }; fr.step = nextStep(wb, s, f); fr.done = !fr.step; stack.push(fr); };
  push(s0, ar, ac, raw.slice(1));
  const refEl = h('div', { class: 'ev-ref mono' });
  const boxes = h('div', { class: 'ev-stack' });
  const note = h('p', { class: 'muted ev-note' });
  const bEval = h('button', { class: 'btn primary' }, 'Đánh giá') as HTMLButtonElement;
  const bIn = h('button', { class: 'btn' }, 'Bước Vào') as HTMLButtonElement;
  const bOut = h('button', { class: 'btn' }, 'Bước Ra') as HTMLButtonElement;
  const top = () => stack[stack.length - 1];
  const cellRef = (fr: Frame) => `${quoteSheet(wb.sheetName(fr.s))}!${toA1(fr.r, fr.c).replace(/([A-Z]+)(\d+)/, '$$$1$$$2')}`;
  const draw = () => {
    const t = top();
    refEl.textContent = cellRef(t);
    boxes.replaceChildren(...stack.map((fr, i) => {
      const b = h('div', { class: 'ev-box mono' + (i === stack.length - 1 ? ' cur' : ''), 'aria-live': i === stack.length - 1 ? 'polite' : 'off' });
      if (fr.step && !fr.done) b.append('=' + fr.f.slice(0, fr.step.start), h('u', {}, fr.f.slice(fr.step.start, fr.step.end)), fr.f.slice(fr.step.end));
      else b.append('=' + fr.f);
      return b;
    }));
    const st = t.step;
    const intoOk = !t.done && st?.kind === 'ref' && st.ref && st.ref.r1 === st.ref.r2 && st.ref.c1 === st.ref.c2 && wb.rawText(st.ref.s, st.ref.r1, st.ref.c1).startsWith('=') && !wb.formulaHidden(st.ref.s, st.ref.r1, st.ref.c1);
    bIn.disabled = !intoOk;
    bOut.disabled = stack.length < 2;
    bEval.textContent = t.done && stack.length === 1 ? 'Khởi động lại' : 'Đánh giá';
    note.textContent = t.done ? (stack.length > 1 ? 'Đã tính xong ô này. Bấm Bước Ra để quay lại công thức trước.' : 'Đã đánh giá xong công thức.') : st?.kind === 'ref' ? 'Phần gạch chân là tham chiếu sẽ được thay bằng giá trị.' : 'Phần gạch chân sẽ được tính.';
  };
  bEval.onclick = () => {
    const t = top();
    if (t.done) {
      if (stack.length === 1) { stack.length = 0; push(s0, ar, ac, raw.slice(1)); draw(); }
      return;
    }
    t.f = applyStep(wb, t.s, t.f, t.step!);
    t.step = nextStep(wb, t.s, t.f); t.done = !t.step;
    draw();
  };
  bIn.onclick = () => { const st = top().step; if (!st?.ref) return; const g = st.ref; push(g.s, g.r1, g.c1, wb.rawText(g.s, g.r1, g.c1).slice(1)); draw(); };
  bOut.onclick = () => {
    if (stack.length < 2) return;
    stack.pop();
    const t = top();
    t.f = applyStep(wb, t.s, t.f, t.step!);
    t.step = nextStep(wb, t.s, t.f); t.done = !t.step;
    draw();
  };
  const body = h('div', { class: 'ev-dlg' }, h('div', { class: 'ev-head' }, h('span', {}, 'Tham chiếu:'), refEl), h('span', { class: 'ev-lbl' }, 'Đánh giá:'), boxes, note, h('div', { class: 'ev-btns' }, bEval, bIn, bOut));
  draw();
  openDialog('Đánh giá Công thức', body, [{ label: 'Đóng', act: () => {} }]);
  setTimeout(() => bEval.focus(), 40);
}

/** menu nút ! (hoặc Alt+Shift+F10) ở ô có lỗi */
export function errorMenu(app: App, x: number, y: number) {
  const s = app.grid.sheet, { ar, ac } = app.grid.sel, wb = app.wb;
  const kind = wb.cellIssue(s, ar, ac);
  if (!kind) { app.toast('Ô hiện hành không có lỗi cần kiểm tra.'); return; }
  const t = issueText(app, ar, ac);
  const g: Range = app.grid.range;
  showMenu([
    { label: t.title, disabled: true },
    'sep',
    ...(kind === 'numText' ? [{ label: 'Chuyển thành Số', act: () => { app.safe(() => { const n = wb.convertToNumber(s, g); if (n > 1) app.toast(`Đã chuyển ${n} ô thành số.`, 'ok'); }); refreshGrid(app); } }] : []),
    ...(kind === 'err' ? [{ label: 'Hiện các Bước Tính…', act: () => evaluateDialog(app) }, { label: 'Truy tìm Lỗi', act: () => app.trace(true) }] : []),
    { label: 'Bỏ qua Lỗi', act: () => { app.safe(() => wb.ignoreIssue(s, g)); refreshGrid(app); } },
    { label: 'Sửa trong Thanh Công thức', act: () => (document.querySelector('#fx-input') as HTMLElement).focus() },
    'sep',
    { label: 'Tùy chọn Kiểm tra Lỗi…', act: () => app.backstage.open('options') },
  ] as any, x, y);
}

/** Kiểm tra Lỗi…: đi qua từng ô có lỗi trên trang */
export function errorCheckDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  let list = wb.issueCells(s);
  if (!list.length) { app.toast('Đã hoàn tất kiểm tra lỗi cho toàn bộ trang tính.', 'ok'); return; }
  let i = Math.max(0, list.findIndex((p) => p.r > app.grid.sel.ar || (p.r === app.grid.sel.ar && p.c >= app.grid.sel.ac)));
  const where = h('div', { class: 'mono ec-where' }), title = h('b', {}), hint = h('p', {}), fx = h('div', { class: 'mono ec-fx' });
  const bConv = h('button', { class: 'btn' }, 'Chuyển thành Số') as HTMLButtonElement;
  const bSteps = h('button', { class: 'btn' }, 'Hiện các Bước Tính…') as HTMLButtonElement;
  const bIgn = h('button', { class: 'btn' }, 'Bỏ qua Lỗi') as HTMLButtonElement;
  const bEdit = h('button', { class: 'btn' }, 'Sửa trong Thanh Công thức') as HTMLButtonElement;
  const bPrev = h('button', { class: 'btn' }, 'Trước') as HTMLButtonElement, bNext = h('button', { class: 'btn' }, 'Tiếp') as HTMLButtonElement;
  let dirty = false;
  const show = () => {
    if (dirty) { list = wb.issueCells(s); dirty = false; }
    if (!list.length) { closeDialog(); app.toast('Đã hoàn tất kiểm tra lỗi cho toàn bộ trang tính.', 'ok'); return; }
    i = Math.min(i, list.length - 1);
    const p = list[i];
    app.grid.select(p.r, p.c); app.syncSelection();
    const t = issueText(app, p.r, p.c);
    where.textContent = `Lỗi ở ô ${toA1(p.r, p.c)} (${i + 1}/${list.length})`;
    title.textContent = t.title; hint.textContent = t.hint;
    fx.textContent = wb.formulaHidden(s, p.r, p.c) ? '' : wb.rawText(s, p.r, p.c);
    bConv.hidden = p.kind !== 'numText'; bSteps.hidden = p.kind !== 'err';
  };
  bConv.onclick = () => { const p = list[i]; app.safe(() => wb.convertToNumber(s, { r1: p.r, c1: p.c, r2: p.r, c2: p.c })); refreshGrid(app); dirty = true; show(); };
  bSteps.onclick = () => { closeDialog(); evaluateDialog(app); };
  bIgn.onclick = () => { const p = list[i]; app.safe(() => wb.ignoreIssue(s, { r1: p.r, c1: p.c, r2: p.r, c2: p.c })); refreshGrid(app); dirty = true; show(); };
  bEdit.onclick = () => { closeDialog(); (document.querySelector('#fx-input') as HTMLElement).focus(); };
  bPrev.onclick = () => { i = (i - 1 + list.length) % list.length; show(); };
  bNext.onclick = () => { i = (i + 1) % list.length; show(); };
  const body = h('div', { class: 'ec-dlg' }, h('div', { class: 'ec-left' }, where, fx, title, hint), h('div', { class: 'ec-right' }, bConv, bSteps, bIgn, bEdit));
  openDialog('Kiểm tra Lỗi', body, [{ label: 'Trước', act: () => { bPrev.click(); return false; } }, { label: 'Tiếp', kind: 'primary', act: () => { bNext.click(); return false; } }, { label: 'Đóng', act: () => {} }]);
  show();
}

/** Cửa sổ Theo dõi: theo dõi giá trị ô (kể cả ở trang khác) khi làm việc chỗ khác */
interface Watch { s: number; r: number; c: number }
export class WatchWindow {
  private el: HTMLElement;
  private body: HTMLElement;
  private list: Watch[] = [];
  private sel = -1;
  open = false;
  constructor(private app: App) {
    this.el = h('aside', { class: 'watch-pane', role: 'dialog', 'aria-label': 'Cửa sổ Theo dõi', hidden: '' });
    const add = h('button', { class: 'btn' }, 'Thêm Theo dõi…'), del = h('button', { class: 'btn' }, 'Xóa Theo dõi'), x = h('button', { class: 'x', 'aria-label': 'Đóng' }, '×');
    add.onclick = () => this.add();
    del.onclick = () => { if (this.sel >= 0) { this.list.splice(this.sel, 1); this.sel = -1; this.draw(); } };
    x.onclick = () => this.toggle(false);
    this.body = h('tbody');
    this.el.append(h('header', {}, h('b', {}, 'Cửa sổ Theo dõi'), x), h('div', { class: 'nm-bar' }, add, del), h('div', { class: 'nm-wrap' }, h('table', { class: 'nm-table' }, h('thead', {}, h('tr', {}, ...['Sổ', 'Trang', 'Ô', 'Giá trị', 'Công thức'].map((t) => h('th', {}, t)))), this.body)));
    document.body.append(this.el);
    app.wb.onChange(() => { if (this.open) this.draw(); });
  }
  toggle(on = !this.open) { this.open = on; this.el.hidden = !on; if (on) this.draw(); }
  add() {
    const g = this.app.grid.range, s = this.app.grid.sheet;
    let n = 0;
    for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 49); r++) for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 19); c++) {
      if (!this.list.some((w) => w.s === s && w.r === r && w.c === c)) { this.list.push({ s, r, c }); n++; }
    }
    this.toggle(true);
    if (!n) this.app.toast('Ô đã có trong Cửa sổ Theo dõi.');
  }
  private draw() {
    const wb = this.app.wb;
    this.list = this.list.filter((w) => wb.sheetIds().includes(w.s));
    this.body.replaceChildren(...(this.list.length ? this.list.map((w, i) => {
      const tr = h('tr', { tabindex: '0', 'aria-selected': String(i === this.sel) },
        h('td', {}, wb.name), h('td', {}, wb.sheetName(w.s)), h('td', {}, toA1(w.r, w.c)), h('td', {}, wb.display(w.s, w.r, w.c).text), h('td', { class: 'mono' }, !wb.formulaHidden(w.s, w.r, w.c) && wb.rawText(w.s, w.r, w.c).startsWith('=') ? wb.rawText(w.s, w.r, w.c) : ''));
      tr.onclick = () => { this.sel = i; this.draw(); };
      tr.ondblclick = () => { if (w.s !== this.app.grid.sheet) { this.app.grid.setSheet(w.s); this.app.renderSheets(); } this.app.grid.select(w.r, w.c); this.app.syncSelection(); this.app.grid.focusKeys(); };
      return tr;
    }) : [h('tr', {}, h('td', { colspan: '5', class: 'muted' }, 'Chưa có ô nào. Chọn ô rồi bấm “Thêm Theo dõi…”.'))]));
  }
}
