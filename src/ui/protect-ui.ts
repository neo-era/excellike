// Xem lại ▸ Bảo vệ Trang tính / Bảo vệ Sổ làm việc / Cho phép Sửa Vùng (0.9)
import { openDialog, h, field, input } from './dialogs';
import { ALLOW_LABELS, DEFAULT_ALLOW, hashPw, checkPw, type ProtectAllow } from '../core/protect';
import { parseRange, rangeToA1 } from '../core/address';
import type { App } from './app';

const pwInput = () => { const i = h('input', { type: 'password', autocomplete: 'new-password' }) as HTMLInputElement; return i; };
const refresh = (app: App) => { app.grid.invalidate(); app.renderRibbon(); app.syncSelection(); app.scheduleSave(); };

/** hỏi mật khẩu để bỏ bảo vệ */
function askPw(title: string, onOk: (pw: string) => Promise<boolean>) {
  const pw = pwInput();
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog(title, h('div', { class: 'an-form' }, field('Mật khẩu', pw), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => { err.textContent = 'Đang kiểm tra…'; void onOk(pw.value).then((ok) => { if (!ok) { err.textContent = 'Mật khẩu không đúng. Lưu ý mật khẩu phân biệt chữ hoa/thường.'; } }); return false; } },
  ]);
  setTimeout(() => pw.focus(), 40);
}

export function toggleSheetProtect(app: App) {
  const s = app.grid.sheet, wb = app.wb, p = wb.meta(s).protect;
  if (p) {
    if (!p.pw) { wb.unprotectSheet(s); refresh(app); app.toast('Đã bỏ bảo vệ trang tính.', 'ok'); return; }
    askPw('Bỏ bảo vệ Trang tính', async (pw) => { if (!(await checkPw(pw, p.pw!))) return false; wb.unprotectSheet(s); document.querySelector('.dlg-back')?.remove(); refresh(app); app.toast('Đã bỏ bảo vệ trang tính.', 'ok'); return true; });
    return;
  }
  const pw = pwInput(), pw2 = pwInput();
  const boxes = ALLOW_LABELS.map(([k, label]) => { const c = h('input', { type: 'checkbox' }) as HTMLInputElement; c.checked = DEFAULT_ALLOW[k]; c.dataset.k = k; return { c, el: h('label', { class: 'chk' }, c, ' ' + label) }; });
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog('Bảo vệ Trang tính', h('div', { class: 'an-form pr-dlg' },
    field('Mật khẩu để bỏ bảo vệ trang tính (có thể để trống)', pw), field('Nhập lại mật khẩu', pw2),
    h('b', {}, 'Cho phép tất cả người dùng của trang tính này:'), h('div', { class: 'pr-list' }, ...boxes.map((b) => b.el)),
    h('p', { class: 'muted' }, 'Chỉ các ô đã bỏ chọn “Khóa” (Định dạng Ô ▸ Bảo vệ) hoặc thuộc vùng “Cho phép Sửa Vùng” mới sửa được.'), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      if (pw.value !== pw2.value) { err.textContent = 'Hai mật khẩu không khớp.'; return false; }
      const allow: Partial<ProtectAllow> = {};
      for (const b of boxes) allow[b.c.dataset.k as keyof ProtectAllow] = b.c.checked;
      const go = async () => { const h2 = pw.value ? await hashPw(pw.value) : undefined; app.safe(() => wb.protectSheet(s, allow, h2)); document.querySelector('.dlg-back')?.remove(); refresh(app); app.toast('Đã bảo vệ trang tính.', 'ok'); };
      err.textContent = pw.value ? 'Đang mã hóa mật khẩu…' : '';
      void go();
      return false;
    } },
  ]);
  setTimeout(() => pw.focus(), 40);
}

export function toggleBookProtect(app: App) {
  const wb = app.wb, p = wb.bookProtect;
  if (p) {
    if (!p.pw) { wb.unprotectBook(); refresh(app); app.renderSheets(); app.toast('Đã bỏ bảo vệ sổ làm việc.', 'ok'); return; }
    askPw('Bỏ bảo vệ Sổ làm việc', async (pw) => { if (!(await checkPw(pw, p.pw!))) return false; wb.unprotectBook(); document.querySelector('.dlg-back')?.remove(); refresh(app); app.toast('Đã bỏ bảo vệ sổ làm việc.', 'ok'); return true; });
    return;
  }
  const pw = pwInput(), pw2 = pwInput();
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog('Bảo vệ Cấu trúc và Cửa sổ', h('div', { class: 'an-form' }, h('p', {}, 'Bảo vệ sổ làm việc cho: Cấu trúc (không thêm, xóa, đổi tên, di chuyển, ẩn trang tính).'), field('Mật khẩu (có thể để trống)', pw), field('Nhập lại mật khẩu', pw2), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      if (pw.value !== pw2.value) { err.textContent = 'Hai mật khẩu không khớp.'; return false; }
      err.textContent = pw.value ? 'Đang mã hóa mật khẩu…' : '';
      void (async () => { const hh = pw.value ? await hashPw(pw.value) : undefined; wb.protectBook(hh); document.querySelector('.dlg-back')?.remove(); refresh(app); app.toast('Đã bảo vệ cấu trúc sổ làm việc.', 'ok'); })();
      return false;
    } },
  ]);
  setTimeout(() => pw.focus(), 40);
}

/** Cho phép Sửa Vùng: các vùng sửa được khi trang bị bảo vệ (không đặt mật khẩu riêng cho từng vùng) */
export function allowRangesDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  if (wb.meta(s).protect) { app.toast('Không thể sửa Vùng Cho phép khi trang tính đang được bảo vệ. Hãy bỏ bảo vệ trang tính trước.', 'err'); return; }
  const list = [...(wb.meta(s).allowRanges ?? [])];
  const lb = h('select', { size: '6', class: 'nm-list', 'aria-label': 'Vùng' }) as HTMLSelectElement;
  const draw = () => lb.replaceChildren(...(list.length ? list.map((x, i) => h('option', { value: String(i) }, `${x.title} — ${rangeToA1(x.range)}`)) : [h('option', { disabled: '' }, 'Chưa có vùng nào')]));
  const title = input('Vùng ' + (list.length + 1)), ref = input(rangeToA1(app.grid.range));
  ref.classList.add('mono');
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  const bAdd = h('button', { class: 'btn' }, 'Thêm'), bDel = h('button', { class: 'btn' }, 'Xóa');
  bAdd.onclick = () => { const g = parseRange(ref.value.trim().replace(/^=/, '').replace(/\$/g, '')); if (!g) { err.textContent = 'Tham chiếu không hợp lệ.'; return; } err.textContent = ''; list.push({ title: title.value.trim() || 'Vùng ' + (list.length + 1), range: g }); title.value = 'Vùng ' + (list.length + 1); draw(); };
  bDel.onclick = () => { if (lb.selectedIndex >= 0 && list[lb.selectedIndex]) { list.splice(lb.selectedIndex, 1); draw(); } };
  draw();
  openDialog('Cho phép Người dùng Sửa Vùng', h('div', { class: 'an-form' }, h('p', { class: 'muted' }, 'Các vùng này vẫn sửa được khi trang tính đang được bảo vệ.'), lb, bDel, h('div', { class: 'an-row' }, field('Tiêu đề', title), field('Tham chiếu tới ô', ref), bAdd), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => { app.safe(() => wb.setAllowRanges(s, list)); refresh(app); } },
  ]);
}
