// Kích hoạt bản quyền (chỉ bản Windows): khóa giao diện tới khi có key hợp lệ; hiện Mã máy để xin key gắn máy.
import { ic } from './icons';
import { tr } from './i18n';

export interface LicInfo { serial: number; type: 'machine' | 'shared'; issued: string; expires: string | null; machine: string | null; daysLeft?: number }
export interface LicStatus { ok: boolean; code: string; message: string; machine: string; info?: LicInfo }

const desk = () => (window as any).xlDesktop as { license?: () => Promise<LicStatus>; activate?: (k: string) => Promise<LicStatus> } | undefined;
let last: LicStatus | null = null;
export const licenseStatus = () => last;

/** Mô tả key cho trang Tài khoản */
export function licenseText(s: LicStatus | null): string {
  if (!s) return 'Bản web: không cần key.';
  if (!s.ok || !s.info) return 'Chưa kích hoạt — ' + s.message;
  const i = s.info;
  const kind = i.type === 'machine' ? `gắn với máy này (Mã máy ${i.machine})` : 'dùng trên mọi máy';
  const exp = i.expires ? `hết hạn ${i.expires}${i.daysLeft !== undefined ? ` (còn ${i.daysLeft} ngày)` : ''}` : 'không thời hạn';
  return `Đã kích hoạt — key số ${i.serial}, ${kind}, ${exp}.`;
}

/** Gọi khi khởi động: kiểm key, chưa hợp lệ thì hiện màn kích hoạt chặn mọi thao tác */
export async function ensureLicense(toast: (m: string, k?: string) => void): Promise<void> {
  const d = desk();
  if (!d?.license) return;
  try { last = await d.license(); } catch (e) { last = { ok: false, code: 'error', message: String((e as Error).message || e), machine: '' }; }
  if (last.ok) {
    const left = last.info?.daysLeft;
    if (left !== undefined && left <= 14) toast(`Key bản quyền ExcelLike còn ${left} ngày (hết hạn ${last.info!.expires}). Liên hệ nơi cấp để gia hạn.`, 'err');
    return;
  }
  await activationDialog(last, true);
}

/** Màn kích hoạt. locked = chưa có key hợp lệ (không đóng được) */
export function activationDialog(st: LicStatus, locked: boolean): Promise<void> {
  const d = desk();
  return new Promise((resolve) => {
    document.getElementById('lic-back')?.remove();
    const back = document.createElement('div');
    back.id = 'lic-back';
    back.className = 'lic-back';
    back.innerHTML = `<div class="lic" role="dialog" aria-modal="true" aria-labelledby="lic-h">
  <div class="lic-top">${ic('protectBook', 28)}<div><h2 id="lic-h">Kích hoạt ExcelLike</h2><p>Nhập key bản quyền được cấp để sử dụng.</p></div></div>
  <label class="lic-l" for="lic-key">Key bản quyền</label>
  <textarea id="lic-key" rows="4" spellcheck="false" autocomplete="off" placeholder="XXXXX-XXXXX-XXXXX-…"></textarea>
  <p class="lic-msg" role="alert"></p>
  <div class="lic-mc"><span>Mã máy của máy này</span><code></code><button type="button" class="lic-copy">${ic('copy', 15)}<span>Chép</span></button></div>
  <p class="lic-note">Key gắn với máy: gửi Mã máy trên cho nơi cấp. Liên hệ LAVIPCO — lamvietphat2022@gmail.com — 0989 725 507.</p>
  <div class="lic-btns">${locked ? '<button type="button" class="lic-quit">Thoát</button>' : '<button type="button" class="lic-cancel">Hủy</button>'}<button type="button" class="lic-ok">Kích hoạt</button></div>
</div>`;
    document.body.appendChild(back);
    const ta = back.querySelector('#lic-key') as HTMLTextAreaElement;
    const msg = back.querySelector('.lic-msg') as HTMLElement;
    const code = back.querySelector('.lic-mc code') as HTMLElement;
    code.textContent = st.machine || '—';
    if (st.code !== 'missing' && !st.ok) msg.textContent = tr(st.message);
    const close = () => { back.remove(); window.removeEventListener('keydown', trap, true); resolve(); };
    // chặn phím tắt của bảng tính khi màn kích hoạt đang mở
    const trap = (e: KeyboardEvent) => {
      if (!back.isConnected) return;
      if (e.key === 'Escape' && !locked) { e.preventDefault(); close(); return; }
      if (!(e.target instanceof Node) || !back.contains(e.target)) { e.preventDefault(); e.stopPropagation(); ta.focus(); }
      else if (e.key === 'Tab') {
        const f = [...back.querySelectorAll<HTMLElement>('textarea, button')];
        const i = f.indexOf(document.activeElement as HTMLElement);
        e.preventDefault(); f[(i + (e.shiftKey ? f.length - 1 : 1)) % f.length].focus();
      } else e.stopPropagation();
    };
    window.addEventListener('keydown', trap, true);
    (back.querySelector('.lic-copy') as HTMLElement).onclick = async () => {
      try { await navigator.clipboard.writeText(st.machine); msg.className = 'lic-msg ok'; msg.textContent = tr('Đã chép Mã máy.'); }
      catch { const r = document.createRange(); r.selectNodeContents(code); getSelection()?.removeAllRanges(); getSelection()?.addRange(r); }
    };
    const okBtn = back.querySelector('.lic-ok') as HTMLButtonElement;
    const go = async () => {
      if (!ta.value.trim()) { msg.className = 'lic-msg'; msg.textContent = tr('Hãy dán key bản quyền.'); ta.focus(); return; }
      okBtn.disabled = true;
      try {
        const r = await d!.activate!(ta.value);
        if (r.ok) { last = r; close(); return; }
        msg.className = 'lic-msg'; msg.textContent = tr(r.message);
      } catch (e) { msg.className = 'lic-msg'; msg.textContent = String((e as Error).message || e); }
      okBtn.disabled = false;
    };
    okBtn.onclick = go;
    ta.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); go(); } });
    (back.querySelector('.lic-quit') as HTMLElement | null)?.addEventListener('click', () => window.close());
    (back.querySelector('.lic-cancel') as HTMLElement | null)?.addEventListener('click', close);
    setTimeout(() => ta.focus(), 30);
  });
}

/** Tệp ▸ Tài khoản ▸ Đổi key */
export async function changeKey() {
  const d = desk();
  if (!d?.license) return;
  const st = last ?? await d.license();
  await activationDialog({ ...st, code: 'missing' }, false);
}
