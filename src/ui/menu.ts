// Menu thả xuống / menu chuột phải kiểu Excel: mục có biểu tượng, phím tắt, dấu ✓, menu con ▸, thanh mini phía trên
import { ic } from './icons';

export type MenuItem =
  | {
    label: string; act?: () => void; icon?: string; hint?: string; danger?: boolean; check?: boolean; disabled?: boolean;
    side?: { icon: string; label: string; act: () => void };
    /** menu con (mảng mục) hoặc nội dung tự dựng (bảng màu…) */
    sub?: MenuItem[] | (() => HTMLElement);
    /** phần tử thêm ngay dưới mục (hàng biểu tượng Tùy chọn Dán) */
    extra?: () => HTMLElement;
  }
  | 'sep';

let openEl: HTMLElement | null = null;
let subEl: HTMLElement | null = null;
let restoreFocus: HTMLElement | null = null;
export function closeMenu() {
  subEl?.remove(); subEl = null;
  openEl?.remove();
  openEl = null;
  document.removeEventListener('pointerdown', outside, true);
  document.removeEventListener('keydown', escClose, true);
  restoreFocus?.focus({ preventScroll: true });
  restoreFocus = null;
}
export const menuOpen = () => !!openEl;
function outside(e: Event) {
  const t = e.target as Node;
  if (openEl && !openEl.contains(t) && !subEl?.contains(t)) closeMenu();
}
function escClose(e: KeyboardEvent) {
  if (!openEl) return;
  if (e.key === 'Escape') {
    e.preventDefault(); e.stopPropagation();
    if (subEl) { const owner = subEl.dataset.owner; subEl.remove(); subEl = null; (openEl.querySelector(`[data-sub="${owner}"]`) as HTMLElement | null)?.focus(); }
    else closeMenu();
    return;
  }
  const scope = subEl && subEl.contains(document.activeElement) ? subEl : openEl;
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    const bs = [...scope.querySelectorAll<HTMLButtonElement>('button.mi:not([disabled])')];
    if (!bs.length) return;
    e.preventDefault();
    const i = bs.indexOf(document.activeElement as HTMLButtonElement);
    bs[(i + (e.key === 'ArrowDown' ? 1 : -1) + bs.length) % bs.length].focus();
  }
  if (e.key === 'ArrowRight' && (document.activeElement as HTMLElement)?.dataset.sub) { e.preventDefault(); (document.activeElement as HTMLElement).click(); }
  if (e.key === 'ArrowLeft' && scope === subEl) { e.preventDefault(); const owner = subEl!.dataset.owner; subEl!.remove(); subEl = null; (openEl.querySelector(`[data-sub="${owner}"]`) as HTMLElement | null)?.focus(); }
}
function place(el: HTMLElement, x: number, y: number) {
  const r = el.getBoundingClientRect();
  const vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
  el.style.left = Math.max(4, Math.min(x, vw - r.width - 4)) + 'px';
  el.style.top = (y + r.height > vh - 4 ? Math.max(4, Math.min(y - r.height - 4, vh - r.height - 4)) : y) + 'px';
}
export function openPopover(el: HTMLElement, x: number, y: number) {
  const keep = document.activeElement as HTMLElement | null;
  closeMenu();
  restoreFocus = keep;
  el.classList.add('pop');
  document.body.appendChild(el);
  place(el, x, y);
  openEl = el;
  setTimeout(() => { document.addEventListener('pointerdown', outside, true); document.addEventListener('keydown', escClose, true); });
}

function buildMenu(items: MenuItem[], isSub: boolean): HTMLElement {
  const m = document.createElement('div');
  m.className = 'menu' + (isSub ? ' sub' : '');
  m.setAttribute('role', 'menu');
  let n = 0;
  for (const it of items) {
    if (it === 'sep') { m.appendChild(document.createElement('hr')); continue; }
    const key = String(n++);
    const row = document.createElement('div');
    row.className = 'mrow';
    const b = document.createElement('button');
    b.className = 'mi' + (it.danger ? ' danger' : '') + (it.sub ? ' has-sub' : '');
    b.setAttribute('role', it.check !== undefined ? 'menuitemcheckbox' : 'menuitem');
    if (it.check !== undefined) b.setAttribute('aria-checked', String(it.check));
    if (it.disabled) b.disabled = true;
    const lead = it.check ? ic('check', 16) : it.icon ? ic(it.icon, 16) : '<i></i>';
    b.innerHTML = lead + `<span></span>` + (it.hint ? `<kbd></kbd>` : it.sub ? '<kbd class="arr">›</kbd>' : '');
    b.querySelector('span')!.textContent = it.label;
    if (it.hint) b.querySelector('kbd')!.textContent = it.hint;
    if (it.sub) {
      b.dataset.sub = key;
      b.setAttribute('aria-haspopup', 'menu');
      const open = () => openSub(b, key, it.sub!);
      b.onclick = open;
      b.onmouseenter = () => { if (subEl?.dataset.owner !== key) open(); };
    } else {
      b.onmouseenter = () => { if (subEl && !isSub) { subEl.remove(); subEl = null; } };
      b.onclick = () => { closeMenu(); it.act?.(); };
    }
    row.appendChild(b);
    if (it.side) {
      const sd = document.createElement('button');
      sd.className = 'mside';
      sd.innerHTML = ic(it.side.icon, 15);
      sd.title = sd.ariaLabel = it.side.label;
      sd.onclick = () => { closeMenu(); it.side!.act(); };
      row.appendChild(sd);
    }
    m.appendChild(row);
    if (it.extra) m.appendChild(it.extra());
  }
  return m;
}
function openSub(owner: HTMLElement, key: string, sub: MenuItem[] | (() => HTMLElement)) {
  subEl?.remove();
  const el = Array.isArray(sub) ? buildMenu(sub, true) : (() => { const w = document.createElement('div'); w.className = 'menu sub'; w.appendChild(sub()); return w; })();
  el.dataset.owner = key;
  el.classList.add('pop');
  document.body.appendChild(el);
  const r = owner.getBoundingClientRect();
  const vw = document.documentElement.clientWidth;
  const w = el.getBoundingClientRect().width;
  place(el, r.right + w + 2 > vw ? r.left - w - 2 : r.right + 2, r.top - 4);
  subEl = el;
}

/** Hiện menu. opts.mini: thanh định dạng nhỏ đặt phía trên menu (menu chuột phải ô của Excel) */
export function showMenu(items: MenuItem[], x: number, y: number, opts: { mini?: HTMLElement; cls?: string } = {}) {
  const m = buildMenu(items, false);
  if (opts.cls) m.classList.add(...opts.cls.split(' '));
  if (opts.mini) {
    const wrap = document.createElement('div');
    wrap.className = 'ctx-wrap';
    wrap.append(opts.mini, m);
    openPopover(wrap, x, y);
    // thanh mini nằm trên điểm bấm, menu ngay dưới (như Excel)
    const mh = opts.mini.getBoundingClientRect().height + 6;
    const r = wrap.getBoundingClientRect();
    const vh = document.documentElement.clientHeight;
    let top = y - mh;
    if (top < 4) top = 4;
    if (top + r.height > vh - 4) top = Math.max(4, vh - r.height - 4);
    wrap.style.top = top + 'px';
  } else openPopover(m, x, y);
  (m.querySelector('button.mi') as HTMLButtonElement | null)?.focus();
}
