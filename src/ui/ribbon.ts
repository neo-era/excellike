// Ribbon kiểu Excel 365: nhóm lệnh 3 tầng (nút lớn, nút nhỏ có chữ, nút biểu tượng), nút tách đôi, hộp chọn,
// nút ↘ mở hộp thoại, nhóm tự co khi cửa sổ hẹp, KeyTips (Alt/F10), ô Tìm kiếm lệnh (Alt+Q),
// và dạng ribbon đơn giản một hàng cho điện thoại.
import { ic } from './icons';
import { closeMenu, openPopover, showMenu, type MenuItem } from './menu';

export interface Btn {
  t: 'big' | 'sm' | 'ic';
  id: string; icon: string; label: string;
  act?: () => void;
  /** mở menu/bảng chọn thả xuống tại vị trí nút */
  drop?: (r: DOMRect) => void;
  /** nút tách đôi: phần chính chạy act, mũi tên mở drop */
  split?: boolean;
  on?: () => boolean;
  later?: string;
  kt?: string;
  kbd?: string;
  /** vạch màu dưới biểu tượng (màu tô, màu chữ) */
  bar?: () => string;
  /** chữ hiển thị trên nút khác với tên lệnh (tìm kiếm dùng label) */
  text?: string;
  /** nhấp đúp (Chổi định dạng: dùng nhiều lần) */
  dbl?: () => void;
}
export interface Combo { t: 'combo'; id: string; label: string; w: number; get: () => string; set: (v: string) => void; list: (r: DOMRect) => void; kt?: string; icon?: string }
export interface Chk { t: 'chk'; id: string; label: string; get: () => boolean; set: (v: boolean) => void; kt?: string }
export interface Stack { t: 'col' | 'row'; items: Item[] }
export type Item = Btn | Combo | Chk | Stack;
export interface Group { id: string; label: string; icon: string; items: Item[]; launch?: { id: string; tip: string; act: () => void; kt?: string; later?: string }; shrink?: number }
export interface Tab { id: string; label: string; kt: string; groups: Group[] }

export interface RibbonHost {
  tabs(): Tab[];
  current: string;
  /** chạy trước mỗi lệnh (kết thúc nhập ô…) */
  before(): void;
  /** sau lệnh: trả focus về lưới nếu lệnh không mở menu/hộp thoại */
  after(keepFocus: boolean): void;
  later(label: string, ver: string): void;
  selectTab(id: string): void;
  openFile(): void;
  simple: boolean;
}

const flat = (items: Item[]): (Btn | Combo | Chk)[] => items.flatMap((x): (Btn | Combo | Chk)[] => (x.t === 'col' || x.t === 'row' ? flat((x as Stack).items) : [x as Btn | Combo | Chk]));
export const allItems = (tabs: Tab[]) => tabs.flatMap((t) => t.groups.flatMap((g) => flat(g.items).map((it) => ({ it, tab: t, grp: g }))));
const caret = () => `<svg class="caret" width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>`;
const tipOf = (it: { label: string; kbd?: string; later?: string }) => it.label + (it.kbd ? ` (${it.kbd})` : '') + (it.later ? ` — có ở bản ${it.later}` : '');

export class Ribbon {
  readonly el: HTMLElement;
  private collapsed = new Set<string>();
  private ro: ResizeObserver;
  constructor(private host: RibbonHost, el: HTMLElement) {
    this.el = el;
    this.ro = new ResizeObserver(() => this.fit());
    this.ro.observe(el);
  }

  // ---------------- dựng ----------------
  render() {
    const tab = this.host.tabs().find((t) => t.id === this.host.current) ?? this.host.tabs()[1];
    this.el.innerHTML = '';
    this.el.classList.toggle('simple', this.host.simple);
    if (this.host.simple) { this.renderSimple(tab); return; }
    this.collapsed.clear();
    for (const g of tab.groups) this.el.appendChild(this.group(g));
    this.fit();
  }
  /** Cập nhật trạng thái bật/tắt, hộp phông/cỡ chữ theo ô đang chọn (không dựng lại) */
  refresh() {
    const items = new Map(allItems(this.host.tabs()).map((x) => [x.it.id, x.it]));
    this.el.querySelectorAll<HTMLElement>('[data-id]').forEach((el) => {
      const it = items.get(el.dataset.id!);
      if (!it) return;
      if (it.t === 'combo') { const i = el.querySelector('input'); if (i && document.activeElement !== i) i.value = it.get(); }
      else if (it.t === 'chk') { const i = el.querySelector('input'); if (i) i.checked = it.get(); }
      else {
        if (it.on) { const on = !!it.on(); el.classList.toggle('on', on); el.setAttribute('aria-pressed', String(on)); }
        if (it.bar) el.querySelectorAll<HTMLElement>('.cbar').forEach((b) => (b.style.background = it.bar!()));
      }
    });
  }
  private group(g: Group, forPop = false): HTMLElement {
    const el = document.createElement('div');
    el.className = 'grp';
    el.dataset.g = g.id;
    const body = document.createElement('div');
    body.className = 'gb';
    for (const it of g.items) body.appendChild(this.item(it));
    const lbl = document.createElement('div');
    lbl.className = 'lbl';
    lbl.textContent = g.label;
    el.append(body, lbl);
    if (g.launch) {
      const l = g.launch;
      const b = document.createElement('button');
      b.className = 'launch' + (l.later ? ' later' : '');
      b.dataset.launch = l.id;
      b.title = l.tip + (l.later ? ` — có ở bản ${l.later}` : '');
      b.setAttribute('aria-label', b.title);
      b.innerHTML = '<svg width="9" height="9" viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true"><path d="M1 1h6v6M7 7 2 2M9 4v5H4"/></svg>';
      b.onmousedown = (e) => e.preventDefault();
      b.onclick = () => { if (forPop) closeMenu(); if (l.later) { this.host.later(l.tip, l.later); return; } this.host.before(); l.act(); };
      el.appendChild(b);
    }
    return el;
  }
  private item(it: Item): HTMLElement {
    if (it.t === 'col' || it.t === 'row') {
      const d = document.createElement('div');
      d.className = it.t === 'col' ? 'stk' : 'rw';
      for (const x of it.items) d.appendChild(this.item(x));
      return d;
    }
    if (it.t === 'combo') return this.combo(it);
    if (it.t === 'chk') {
      const l = document.createElement('label');
      l.className = 'ck';
      l.dataset.id = it.id;
      const i = document.createElement('input');
      i.type = 'checkbox';
      i.checked = it.get();
      i.onchange = () => { this.host.before(); it.set(i.checked); this.host.after(false); };
      l.append(i, document.createTextNode(it.label));
      return l;
    }
    return this.btn(it as Btn);
  }
  private iconHtml(it: Btn, size: number) {
    return `<span class="ico">${ic(it.icon, size)}${it.bar ? `<i class="cbar" style="background:${it.bar()}"></i>` : ''}</span>`;
  }
  private btn(it: Btn): HTMLElement {
    const tip = tipOf(it);
    const size = it.t === 'big' ? 32 : 16;
    const label = it.text ?? it.label;
    const hasDrop = !!it.drop;
    const run = (el: HTMLElement, part: 'main' | 'dd' | 'all') => {
      if (it.later) { this.host.later(it.label, it.later); return; }
      this.host.before();
      if (part === 'dd' || (part === 'all' && hasDrop && !it.split)) { it.drop!(el.getBoundingClientRect()); return; }
      it.act?.();
      this.host.after(false);
    };
    let el: HTMLElement;
    if (it.split && hasDrop) {
      el = document.createElement('div');
      el.className = `${it.t} split`;
      const main = document.createElement('button');
      main.className = 'main';
      const dd = document.createElement('button');
      dd.className = 'dd';
      if (it.t === 'big') { main.innerHTML = this.iconHtml(it, size); dd.innerHTML = `<span>${label}</span>${caret()}`; }
      else if (it.t === 'sm') { main.innerHTML = this.iconHtml(it, size) + `<span>${label}</span>`; dd.innerHTML = caret(); }
      else { main.innerHTML = this.iconHtml(it, size); dd.innerHTML = caret(); }
      main.title = tip; main.setAttribute('aria-label', tip);
      dd.title = 'Thêm tùy chọn: ' + it.label; dd.setAttribute('aria-label', dd.title); dd.setAttribute('aria-haspopup', 'menu');
      main.onmousedown = dd.onmousedown = (e) => e.preventDefault();
      main.onclick = () => run(main, 'main');
      dd.onclick = () => run(el, 'dd');
      el.append(main, dd);
    } else {
      el = document.createElement('button');
      el.className = it.t;
      el.innerHTML = this.iconHtml(it, size) + (it.t === 'ic' ? '' : `<span>${label}</span>`) + (hasDrop ? caret() : '');
      el.title = tip; el.setAttribute('aria-label', tip);
      if (hasDrop) el.setAttribute('aria-haspopup', 'menu');
      el.onmousedown = (e) => e.preventDefault();
      el.onclick = () => run(el, 'all');
      if (it.dbl) el.ondblclick = () => it.dbl!();
    }
    el.dataset.id = it.id;
    if (it.later) { el.classList.add('later'); el.title = tip; el.setAttribute('aria-disabled', 'true'); }
    if (it.on) { const on = !!it.on(); el.classList.toggle('on', on); el.setAttribute('aria-pressed', String(on)); }
    return el;
  }
  private combo(it: Combo): HTMLElement {
    const el = document.createElement('div');
    el.className = 'combo';
    el.dataset.id = it.id;
    el.style.width = it.w + 'px';
    const inp = document.createElement('input');
    inp.setAttribute('aria-label', it.label);
    inp.title = it.label;
    inp.spellcheck = false;
    inp.value = it.get();
    inp.onfocus = () => inp.select();
    inp.onkeydown = (e) => {
      if (e.key === 'Enter') { e.preventDefault(); this.host.before(); it.set(inp.value.trim()); inp.value = it.get(); this.host.after(false); }
      if (e.key === 'Escape') { e.preventDefault(); inp.value = it.get(); this.host.after(false); }
      if (e.key === 'ArrowDown' && e.altKey) { e.preventDefault(); it.list(el.getBoundingClientRect()); }
    };
    inp.onblur = () => { inp.value = it.get(); };
    const dd = document.createElement('button');
    dd.className = 'dd';
    dd.tabIndex = -1;
    dd.innerHTML = caret();
    dd.setAttribute('aria-label', 'Danh sách ' + it.label);
    dd.onmousedown = (e) => e.preventDefault();
    dd.onclick = () => { this.host.before(); it.list(el.getBoundingClientRect()); };
    el.append(inp, dd);
    return el;
  }

  // ---------------- co nhóm khi hẹp ----------------
  private fit() {
    if (this.host.simple || !this.el.isConnected || !this.el.clientWidth) return;
    const tab = this.host.tabs().find((t) => t.id === this.host.current);
    if (!tab) return;
    // thử mở lại các nhóm đã co nếu đã đủ chỗ
    if (this.collapsed.size) {
      for (const g of tab.groups) if (this.collapsed.has(g.id)) this.replace(g, false);
      this.collapsed.clear();
    }
    // thứ tự co như Excel: nhóm có shrink nhỏ co trước, còn lại co từ phải sang trái
    const order = tab.groups.map((g, i) => ({ g, k: g.shrink ?? 100 - i })).sort((a, b) => a.k - b.k).map((x) => x.g);
    for (const g of order) {
      if (this.el.scrollWidth <= this.el.clientWidth + 1) break;
      this.replace(g, true);
      this.collapsed.add(g.id);
    }
  }
  private replace(g: Group, collapse: boolean) {
    const old = this.el.querySelector(`.grp[data-g="${g.id}"]`);
    if (!old) return;
    old.replaceWith(collapse ? this.collapsedGroup(g) : this.group(g));
  }
  private collapsedGroup(g: Group): HTMLElement {
    const el = document.createElement('div');
    el.className = 'grp collapsed';
    el.dataset.g = g.id;
    const b = document.createElement('button');
    b.className = 'big';
    b.innerHTML = `<span class="ico">${ic(g.icon, 32)}</span><span>${g.label}</span>${caret()}`;
    b.title = g.label;
    b.setAttribute('aria-haspopup', 'true');
    b.onmousedown = (e) => e.preventDefault();
    b.onclick = () => {
      const r = b.getBoundingClientRect();
      const pop = document.createElement('div');
      pop.className = 'grp-pop';
      pop.appendChild(this.group(g, true));
      pop.addEventListener('click', (e) => { if ((e.target as HTMLElement).closest('button.main, button.big, button.sm, button.ic, .ck') && !(e.target as HTMLElement).closest('.dd, [aria-haspopup]')) setTimeout(closeMenu); });
      openPopover(pop, r.left, r.bottom + 2);
    };
    const lbl = document.createElement('div');
    lbl.className = 'lbl';
    lbl.textContent = g.label;
    el.append(b, lbl);
    return el;
  }

  // ---------------- ribbon đơn giản (điện thoại) ----------------
  private renderSimple(tab: Tab) {
    const sel = document.createElement('button');
    sel.id = 'tabsel';
    sel.className = 'tsel';
    sel.innerHTML = `<span>${tab.label}</span>${caret()}`;
    sel.setAttribute('aria-label', 'Chọn thẻ lệnh: ' + tab.label);
    sel.setAttribute('aria-haspopup', 'menu');
    sel.onclick = () => {
      const r = sel.getBoundingClientRect();
      showMenu(this.host.tabs().map((t) => ({ label: t.label, act: () => (t.id === 'file' ? this.host.selectTab('file') : this.host.selectTab(t.id)) })), r.left, r.bottom + 4);
    };
    const row = document.createElement('div');
    row.className = 'srow';
    for (const g of tab.groups) {
      for (const it of flat(g.items)) {
        if (it.t === 'chk') continue;
        if (it.t === 'combo') {
          const b = document.createElement('button');
          b.className = 'ic';
          b.dataset.id = it.id;
          b.innerHTML = ic(it.icon ?? 'font', 20);
          b.title = it.label; b.setAttribute('aria-label', it.label);
          b.onclick = () => { this.host.before(); it.list(b.getBoundingClientRect()); };
          row.appendChild(b);
          continue;
        }
        const b = document.createElement('button');
        b.className = 'ic';
        b.dataset.id = it.id;
        b.innerHTML = this.iconHtml(it, 20);
        b.title = tipOf(it); b.setAttribute('aria-label', b.title);
        if (it.later) b.classList.add('later');
        if (it.on) { const on = !!it.on(); b.classList.toggle('on', on); b.setAttribute('aria-pressed', String(on)); }
        b.onmousedown = (e) => e.preventDefault();
        b.onclick = () => {
          if (it.later) { this.host.later(it.label, it.later); return; }
          this.host.before();
          if (it.drop && !it.split) { it.drop(b.getBoundingClientRect()); return; }
          if (it.act) { it.act(); this.host.after(false); } else it.drop?.(b.getBoundingClientRect());
        };
        row.appendChild(b);
      }
      const sep = document.createElement('span');
      sep.className = 'ssep';
      row.appendChild(sep);
    }
    row.lastElementChild?.remove();
    this.el.append(sel, row);
  }
}

// ======================= KeyTips =======================
export interface KtTarget { kt: string; el: () => HTMLElement | null; run: () => void; level?: number }
export class KeyTips {
  active = false;
  private level = 0;
  private typed = '';
  private layer: HTMLElement | null = null;
  private targets: KtTarget[] = [];
  private altDown = false;
  constructor(private top: () => KtTarget[], private inTab: () => KtTarget[], private canStart: () => boolean) {
    document.addEventListener('keydown', (e) => this.onDown(e), true);
    document.addEventListener('keyup', (e) => this.onUp(e), true);
    document.addEventListener('pointerdown', () => { if (this.active) this.exit(); }, true);
    window.addEventListener('blur', () => { this.altDown = false; if (this.active) this.exit(); });
    window.addEventListener('resize', () => { if (this.active) this.draw(); });
  }
  private onDown(e: KeyboardEvent) {
    if (this.active) {
      if (e.key === 'Alt' || e.key === 'F10') { e.preventDefault(); e.stopPropagation(); this.exit(); return; }
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); if (this.level > 0 && !this.typed) this.show(0); else if (this.typed) { this.typed = ''; this.draw(); } else this.exit(); return; }
      if (e.ctrlKey || e.metaKey) { this.exit(); return; }
      // mũi tên / Tab khi KeyTips đang hiện: chuyển sang điều hướng Ribbon bằng bàn phím (như Excel)
      if (['ArrowLeft', 'ArrowRight', 'ArrowDown', 'ArrowUp', 'Tab'].includes(e.key) && this.level === 0) {
        e.preventDefault(); e.stopPropagation(); this.exit();
        (document.querySelector<HTMLElement>('#tabs [aria-selected="true"]') ?? document.querySelector<HTMLElement>('#tabs [role="tab"]'))?.focus();
        return;
      }
      const k = e.key.length === 1 ? e.key.toUpperCase() : '';
      if (!k || !/[A-Z0-9À-Ỹ]/.test(k)) { if (!['Shift', 'Control'].includes(e.key)) this.exit(); return; }
      e.preventDefault(); e.stopPropagation();
      this.press(k);
      return;
    }
    if (e.key === 'Alt') { this.altDown = !e.repeat; return; }
    this.altDown = false;
    if (e.key === 'F10' && !e.shiftKey && !e.ctrlKey && !e.altKey && this.canStart()) { e.preventDefault(); e.stopPropagation(); this.show(0); return; }
    // Alt+chữ (nhấn giữ Alt): vào thẳng thẻ tương ứng
    if (e.altKey && !e.ctrlKey && !e.metaKey && e.key.length === 1 && /[a-z]/i.test(e.key) && this.canStart()) {
      const k = e.key.toUpperCase();
      const t = this.top().find((x) => x.kt === k);
      if (!t) return;
      e.preventDefault(); e.stopPropagation();
      this.show(0);
      this.press(k);
    }
  }
  private onUp(e: KeyboardEvent) {
    if (e.key !== 'Alt') return;
    if (this.altDown && !this.active && this.canStart()) { e.preventDefault(); this.show(0); }
    else if (this.altDown && this.active && !this.typed && this.level === 0) { /* nhấn Alt lần nữa đã xử lý ở keydown */ }
    this.altDown = false;
  }
  private press(k: string) {
    this.typed += k;
    const m = this.targets.filter((t) => t.kt.startsWith(this.typed));
    if (!m.length) { this.typed = ''; this.draw(); return; }
    const exact = m.find((t) => t.kt === this.typed);
    if (exact && m.length === 1) {
      const lvl = this.level;
      this.typed = '';
      if (lvl === 0 && exact.level === 1) { exact.run(); this.show(1); return; }
      this.exit();
      exact.run();
      return;
    }
    this.draw();
  }
  show(level: number) {
    this.active = true;
    this.level = level;
    this.typed = '';
    this.targets = level === 0 ? this.top() : this.inTab();
    document.documentElement.classList.add('kt-on');
    this.draw();
  }
  exit() {
    this.active = false;
    this.typed = '';
    this.layer?.remove();
    this.layer = null;
    document.documentElement.classList.remove('kt-on');
  }
  private draw() {
    this.layer?.remove();
    const layer = document.createElement('div');
    layer.className = 'kt-layer';
    layer.setAttribute('aria-hidden', 'true');
    for (const t of this.targets) {
      if (this.typed && !t.kt.startsWith(this.typed)) continue;
      const el = t.el();
      if (!el) continue;
      const r = el.getBoundingClientRect();
      if (!r.width || !r.height) continue;
      const b = document.createElement('span');
      b.className = 'kt';
      b.textContent = t.kt;
      const big = el.classList.contains('big') || el.getAttribute('role') === 'tab';
      b.style.left = r.left + r.width / 2 + 'px';
      b.style.top = (big ? r.bottom - 6 : r.top + r.height * 0.7) + 'px';
      layer.appendChild(b);
    }
    document.body.appendChild(layer);
    this.layer = layer;
  }
}

/** Bỏ dấu tiếng Việt để tìm kiếm không dấu */
export const fold = (s: string) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
export type { MenuItem };

/** Điều hướng Ribbon bằng bàn phím: ←/→ giữa các thẻ và lệnh, ↓ vào lệnh, Ctrl+←/→ giữa các nhóm, Alt+↓ mở menu, Esc về trang tính */
export function bindRibbonKeys(onSelectTab: (el: HTMLElement) => void, toGrid: () => void) {
  const focusables = () => [...document.querySelectorAll<HTMLElement>('#ribbon button:not([disabled]), #ribbon input, #ribbon select')].filter((x) => x.offsetParent !== null);
  document.addEventListener('keydown', (e) => {
    const t = e.target as HTMLElement | null;
    if (!t) return;
    const inTabs = !!t.closest('#tabs'), inRb = !!t.closest('#ribbon');
    if (!inTabs && !inRb) return;
    if (e.key === 'Escape') { e.preventDefault(); toGrid(); return; }
    if (inTabs) {
      const tabs = [...document.querySelectorAll<HTMLElement>('#tabs [role="tab"]')].filter((x) => x.offsetParent !== null);
      const i = tabs.indexOf(t);
      if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
        e.preventDefault();
        const n = tabs[(i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
        if (n) { const id = n.dataset.tab; if (id !== 'file') onSelectTab(n); (document.querySelector<HTMLElement>(`#tabs [data-tab="${id}"]`) ?? n).focus(); }
      } else if (e.key === 'ArrowDown' || (e.key === 'Tab' && !e.shiftKey)) { e.preventDefault(); focusables()[0]?.focus(); }
      return;
    }
    if (t.tagName === 'INPUT' && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) return;
    const list = focusables(), i = list.indexOf(t);
    if (e.ctrlKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft')) {
      e.preventDefault();
      const grps = [...document.querySelectorAll<HTMLElement>('#ribbon .grp')];
      const gi = grps.findIndex((g) => g.contains(t));
      const ng = grps[(gi + (e.key === 'ArrowRight' ? 1 : grps.length - 1)) % grps.length];
      (ng?.querySelector<HTMLElement>('button, input, select'))?.focus();
      return;
    }
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') { e.preventDefault(); list[(i + (e.key === 'ArrowRight' ? 1 : list.length - 1)) % list.length]?.focus(); return; }
    if ((e.key === 'ArrowDown' && (e.altKey || t.getAttribute('aria-haspopup') === 'menu')) || (e.key === 'ArrowDown' && t.classList.contains('main'))) {
      e.preventDefault();
      const dd = t.classList.contains('main') ? t.parentElement?.querySelector<HTMLElement>('.dd') : t;
      dd?.click();
      return;
    }
    if (e.key === 'ArrowDown') { e.preventDefault(); list[i + 1]?.focus(); return; }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (i <= 0) document.querySelector<HTMLElement>('#tabs [aria-selected="true"]')?.focus(); else list[i - 1]?.focus();
    }
  });
}
