// Giao diện Tên vùng: Trình Quản lý Tên (Ctrl+F3), Xác định Tên, Tạo từ Vùng chọn (Ctrl+Shift+F3), Dán Tên (F3)
import { openDialog, closeDialog, h, field, select, input } from './dialogs';
import { formatValue } from '../core/numfmt';
import { quoteSheet } from '../core/formula';
import { rangeToA1 } from '../core/address';
import type { NameDef } from '../core/workbook';
import type { App } from './app';

/** tham chiếu tuyệt đối của vùng đang chọn, vd ='Dự toán'!$B$2:$B$20 */
export function selRef(app: App): string {
  const g = app.grid.range, s = app.grid.sheet;
  const a = rangeToA1(g).split(':').map((p) => p.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2')).join(':');
  return `=${quoteSheet(app.wb.sheetName(s))}!${a}`;
}
const valueText = (v: unknown) => {
  if (v === undefined) return '';
  if (Array.isArray(v)) return '{' + v.slice(0, 3).map((r: unknown[]) => r.slice(0, 4).map((x) => formatValue(x as never, undefined).text).join(';')).join('|') + (v.length > 3 || (v[0]?.length ?? 0) > 4 ? '…' : '') + '}';
  if (v && typeof v === 'object' && 'value' in (v as object)) return String((v as { value: unknown }).value);
  return formatValue(v as never, undefined).text;
};
/** giá trị hiển thị của tên: vùng ô → {giá trị; …} như Excel, công thức → kết quả */
function nameValueText(app: App, n: NameDef): string {
  const t = app.wb.nameTarget(n.name, n.scope);
  if (!t) return valueText(app.wb.nameValue(n.name, n.scope));
  const g = t.range, out: string[] = [];
  const cells = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
  for (let r = g.r1; r <= g.r2 && out.length < 6; r++) for (let c = g.c1; c <= g.c2 && out.length < 6; c++) out.push(app.wb.display(t.sheet, r, c).text);
  return cells === 1 ? out[0] : '{' + out.join(';') + (cells > 6 ? ';…' : '') + '}';
}
const scopeText = (app: App, n: { scope?: number }) => (n.scope === undefined ? 'Sổ làm việc' : app.wb.sheetName(n.scope));

/** Tên Mới / Sửa Tên */
export function nameDialog(app: App, edit?: NameDef, done?: () => void) {
  const wb = app.wb;
  const nm = input(edit?.name ?? '', 'vd: DonGia');
  const scopes: [string, string][] = [['', 'Sổ làm việc'], ...wb.sheetIds().map((id) => [String(id), wb.sheetName(id)] as [string, string])];
  const sc = select(scopes, edit?.scope !== undefined ? String(edit.scope) : '');
  const cm = h('textarea', { rows: '2', maxlength: '255' }) as HTMLTextAreaElement;
  cm.value = edit?.comment ?? '';
  const ref = input(edit?.expr ?? selRef(app));
  ref.classList.add('mono');
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  if (!edit) {
    // gợi ý tên từ ô tiêu đề phía trên / bên trái như Excel
    const { r1, c1 } = app.grid.range, s = app.grid.sheet;
    const guess = (r1 > 0 ? wb.display(s, r1 - 1, c1).text : '') || (c1 > 0 ? wb.display(s, r1, c1 - 1).text : '');
    const g = guess.trim().replace(/\s+/g, '_').replace(/[^A-Za-z0-9_.À-￿]/g, '');
    if (g && /^[A-Za-z_À-￿]/.test(g) && !/^[A-Za-z]{1,3}\d+$/.test(g)) nm.value = g;
  }
  const body = h('div', { class: 'nm-form' }, field('Tên', nm), field('Phạm vi', sc), field('Chú thích', cm), field('Tham chiếu tới', ref), err);
  openDialog(edit ? 'Sửa Tên' : 'Tên Mới', body, [
    { label: 'Hủy', act: () => { done?.(); } },
    {
      label: 'OK', kind: 'primary', act: () => {
        const scope = sc.value === '' ? undefined : +sc.value;
        try {
          if (edit) wb.updateName(edit.name, edit.scope, { name: nm.value, expr: ref.value, scope, comment: cm.value.trim() || undefined });
          else wb.defineName(nm.value, ref.value, scope, cm.value.trim() || undefined);
        } catch (e) { err.textContent = (e as Error).message; return false; }
        app.afterNames();
        done?.();
      },
    },
  ]);
  setTimeout(() => { nm.focus(); nm.select(); }, 40);
}

/** Trình Quản lý Tên (Ctrl+F3) */
export function nameManager(app: App) {
  const wb = app.wb;
  let sel: NameDef | null = null;
  const tbody = h('tbody');
  const q = h('input', { type: 'search', placeholder: 'Lọc theo tên…', 'aria-label': 'Lọc tên' }) as HTMLInputElement;
  const refBox = input(''); refBox.classList.add('mono'); refBox.disabled = true;
  const bEdit = h('button', { class: 'btn' }, 'Sửa…') as HTMLButtonElement;
  const bDel = h('button', { class: 'btn danger' }, 'Xóa') as HTMLButtonElement;
  const bNew = h('button', { class: 'btn' }, 'Mới…');
  const draw = () => {
    const t = q.value.trim().toLowerCase();
    const list = wb.listNames().filter((n) => !n.hidden && (!t || n.name.toLowerCase().includes(t)));
    tbody.replaceChildren();
    if (!list.length) tbody.append(h('tr', {}, h('td', { colspan: '5', class: 'muted' }, wb.listNames().length ? 'Không có tên phù hợp.' : 'Chưa có tên nào. Bấm “Mới…” để xác định tên cho vùng ô.')));
    for (const n of list) {
      const tr = h('tr', { tabindex: '0', 'aria-selected': String(!!sel && sel.name === n.name && sel.scope === n.scope) },
        h('td', {}, n.name), h('td', {}, nameValueText(app, n)), h('td', { class: 'mono' }, n.expr), h('td', {}, scopeText(app, n)), h('td', {}, n.comment ?? ''));
      const pick = () => { sel = n; refBox.value = n.expr; draw(); };
      tr.onclick = pick;
      tr.ondblclick = () => { pick(); edit(); };
      tr.onkeydown = (e) => { if (e.key === 'Enter') { pick(); edit(); } if (e.key === 'Delete') { pick(); del(); } };
      tbody.append(tr);
    }
    bEdit.disabled = bDel.disabled = !sel;
  };
  const reopen = () => nameManager(app);
  const edit = () => { if (sel) { const s = sel; closeDialog(); nameDialog(app, s, reopen); } };
  const del = () => {
    if (!sel) return;
    if (!confirm(`Xóa tên “${sel.name}”? Công thức đang dùng tên này sẽ báo #NAME?.`)) return;
    wb.deleteName(sel.name, sel.scope); sel = null; refBox.value = ''; app.afterNames(); draw();
  };
  bNew.onclick = () => { closeDialog(); nameDialog(app, undefined, reopen); };
  bEdit.onclick = edit; bDel.onclick = del;
  q.oninput = draw;
  const table = h('table', { class: 'nm-table' }, h('thead', {}, h('tr', {}, h('th', {}, 'Tên'), h('th', {}, 'Giá trị'), h('th', {}, 'Tham chiếu tới'), h('th', {}, 'Phạm vi'), h('th', {}, 'Chú thích'))), tbody);
  const body = h('div', { class: 'nm-mgr' }, h('div', { class: 'nm-bar' }, bNew, bEdit, bDel, h('span', { class: 'sp' }), q), h('div', { class: 'nm-wrap' }, table), field('Tham chiếu tới', refBox));
  draw();
  const back = openDialog('Trình Quản lý Tên', body, [{ label: 'Đóng', kind: 'primary', act: () => {} }]);
  back.querySelector('.dlg')?.classList.add('nm-dlg');
}

/** Tạo Tên từ Vùng chọn (Ctrl+Shift+F3) */
export function createFromSelection(app: App) {
  const g = app.grid.range;
  const mk = (label: string, on: boolean) => { const c = h('input', { type: 'checkbox' }) as HTMLInputElement; c.checked = on; return { c, el: h('label', { class: 'chk' }, c, ' ' + label) }; };
  // như Excel: đoán vị trí tiêu đề theo ô chữ ở hàng trên / cột trái
  const isText = (r: number, c: number) => { const v = app.wb.value(app.grid.sheet, r, c); return typeof v === 'string' && v !== ''; };
  const top = mk('Hàng trên cùng', g.r2 > g.r1 && isText(g.r1, g.c2));
  const left = mk('Cột bên trái', g.c2 > g.c1 && isText(g.r2, g.c1));
  const bottom = mk('Hàng dưới cùng', false), right = mk('Cột bên phải', false);
  const body = h('div', { class: 'nm-create' }, h('p', {}, 'Tạo tên từ các giá trị ở:'), top.el, left.el, bottom.el, right.el);
  openDialog('Tạo Tên từ Vùng chọn', body, [
    { label: 'Hủy', act: () => {} },
    {
      label: 'OK', kind: 'primary', act: () => {
        const n = app.wb.createNamesFrom(app.grid.sheet, g, { top: top.c.checked, left: left.c.checked, bottom: bottom.c.checked, right: right.c.checked });
        app.afterNames();
        app.toast(n ? `Đã tạo ${n} tên.` : 'Không tạo được tên nào (ô tiêu đề trống hoặc không hợp lệ).', n ? 'ok' : 'err');
      },
    },
  ]);
}

/** Dán Tên (F3): đang sửa ô → chèn tên vào công thức; không sửa → dán danh sách tên vào trang tính */
export function pasteNameDialog(app: App, editing: boolean) {
  const list = app.wb.listNames().filter((n) => !n.hidden && (n.scope === undefined || n.scope === app.grid.sheet));
  if (!list.length) { app.toast('Sổ làm việc chưa có tên nào. Dùng Công thức ▸ Xác định Tên.', 'err'); return; }
  const sel = h('select', { size: String(Math.min(10, Math.max(4, list.length))), class: 'nm-list', 'aria-label': 'Tên' }) as HTMLSelectElement;
  for (const n of list) sel.append(h('option', { value: n.name }, n.name));
  sel.selectedIndex = 0;
  const paste = () => { const v = sel.value; closeDialog(); app.pasteName(v); };
  sel.ondblclick = paste;
  sel.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); paste(); } };
  const buttons = [{ label: 'Hủy', act: () => { app.grid.focusKeys(); } }];
  if (!editing) buttons.push({ label: 'Dán Danh sách', act: () => { app.pasteNameList(); } });
  buttons.push({ label: 'OK', kind: 'primary', act: () => { paste(); return false; } } as any);
  openDialog('Dán Tên', h('div', {}, field('Dán tên', sel)), buttons as any);
  setTimeout(() => sel.focus(), 40);
}
