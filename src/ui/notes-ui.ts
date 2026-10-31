// Ghi chú (Shift+F2) và Chú thích dạng chuỗi (Ctrl+Shift+F2) như Excel 365:
// tam giác đỏ/tím ở góc ô, rê chuột để xem, hộp sửa ghi chú màu vàng, thẻ chú thích có trả lời,
// ngăn Chú thích liệt kê mọi chuỗi trong trang tính, Hiện tất cả Ghi chú.
import { ic } from './icons';
import { toA1 } from '../core/address';
import type { App } from './app';

const AUTHOR_KEY = 'xl-author';
export function authorName(): string {
  try { return localStorage.getItem(AUTHOR_KEY)?.trim() || 'Người dùng'; } catch { return 'Người dùng'; }
}
export function setAuthorName(n: string) { try { localStorage.setItem(AUTHOR_KEY, n.trim()); } catch { /* bỏ qua */ } }
const when = (iso: string) => {
  const d = new Date(iso);
  if (isNaN(+d) || +d === 0) return '';
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
};

export class NoteLayer {
  private tip: HTMLElement | null = null;
  private tipAt = '';
  private timer = 0;
  private showAll = false;
  private pinned: HTMLElement[] = [];
  private editor: HTMLElement | null = null;
  private card: HTMLElement | null = null;
  constructor(private app: App) {
    app.grid.onHover = (r, c) => this.hover(r, c);
    const reflow = () => { this.hideTip(); if (this.showAll) this.drawAll(); if (this.card) this.closeCard(); };
    document.getElementById('gridwrap')?.addEventListener('scroll', reflow, true);
    app.wb.onChange(() => { if (this.showAll) this.drawAll(); this.syncPane(); });
  }
  get allShown() { return this.showAll; }

  // ---------- rê chuột để xem ----------
  private hover(r: number, c: number) {
    if (this.editor || this.card) return;
    const s = this.app.grid.sheet;
    const st = r >= 0 ? this.app.wb.style(s, r, c) : null;
    const k = st && (st.nt || st.cm) ? `${s}:${r}:${c}` : '';
    if (k === this.tipAt) return;
    clearTimeout(this.timer);
    this.hideTip();
    if (!k) return;
    this.timer = window.setTimeout(() => { this.tipAt = k; this.tip = this.bubble(r, c, false); document.body.appendChild(this.tip); }, 250);
  }
  private hideTip() { clearTimeout(this.timer); this.tip?.remove(); this.tip = null; this.tipAt = ''; }

  /** hộp xem ghi chú / tóm tắt chú thích cạnh ô */
  private bubble(r: number, c: number, pinned: boolean): HTMLElement {
    const st = this.app.wb.style(this.app.grid.sheet, r, c);
    const el = document.createElement('div');
    el.className = st.cm ? 'cm-tip' : 'nt-box' + (pinned ? ' pinned' : '');
    el.setAttribute('role', 'note');
    if (st.cm) {
      for (const x of st.cm) {
        const it = document.createElement('div'); it.className = 'cm-it';
        const h = document.createElement('div'); h.className = 'cm-h';
        const b = document.createElement('b'); b.textContent = x.a;
        const t = document.createElement('small'); t.textContent = when(x.d);
        h.append(b, t);
        const p = document.createElement('p'); p.textContent = x.t;
        it.append(h, p); el.appendChild(it);
      }
      if (st.cmDone) { const d = document.createElement('small'); d.className = 'cm-done'; d.textContent = 'Đã giải quyết'; el.appendChild(d); }
      if (st.nt) { const n = document.createElement('p'); n.className = 'cm-nt'; n.textContent = st.nt; el.appendChild(n); }
    } else el.textContent = st.nt ?? '';
    const a = this.app.grid.cellPageRect(r, c);
    el.style.left = Math.min(window.innerWidth - 270, a.right + 8) + 'px';
    el.style.top = Math.max(4, Math.min(window.innerHeight - 160, a.top)) + 'px';
    return el;
  }

  // ---------- Hiện tất cả Ghi chú ----------
  toggleAll() { this.showAll = !this.showAll; this.drawAll(); this.app.ribbon.refresh(); }
  private drawAll() {
    for (const p of this.pinned) p.remove();
    this.pinned = [];
    if (!this.showAll) return;
    const g = this.app.grid;
    for (const p of this.app.wb.noteCells(g.sheet, 'note')) {
      const a = g.cellPageRect(p.r, p.c);
      const box = document.getElementById('gridwrap')!.getBoundingClientRect();
      if (a.right < box.left || a.left > box.right || a.bottom < box.top || a.top > box.bottom) continue;
      const el = this.bubble(p.r, p.c, true);
      document.body.appendChild(el); this.pinned.push(el);
    }
  }

  // ---------- Ghi chú: sửa (Shift+F2) ----------
  editNote(r = this.app.grid.sel.ar, c = this.app.grid.sel.ac) {
    this.closeEditor(); this.hideTip(); this.closeCard();
    const app = this.app, s = app.grid.sheet, st = app.wb.style(s, r, c);
    if (st.cm && !st.nt) { this.openCard(r, c); app.toast('Ô này có chú thích dạng chuỗi — trả lời trong thẻ chú thích.'); return; }
    const box = document.createElement('div');
    box.className = 'nt-edit';
    const ta = document.createElement('textarea');
    ta.setAttribute('aria-label', `Ghi chú ô ${toA1(r, c)}`);
    ta.value = st.nt ?? authorName() + ':\n';
    box.appendChild(ta);
    const a = app.grid.cellPageRect(r, c);
    box.style.left = Math.min(window.innerWidth - 250, a.right + 8) + 'px';
    box.style.top = Math.max(4, Math.min(window.innerHeight - 170, a.top)) + 'px';
    document.body.appendChild(box);
    this.editor = box;
    const save = () => {
      if (this.editor !== box) return;
      const v = ta.value;
      // chỉ còn dòng tên người viết → coi như không có ghi chú
      const empty = !v.trim() || v.trim() === authorName() + ':';
      app.safe(() => app.wb.setNote(s, r, c, empty ? null : v));
      this.closeEditor(); app.grid.invalidate(); app.grid.focusKeys();
    };
    ta.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') { e.preventDefault(); save(); }
    });
    // bấm ra ngoài hộp → lưu như Excel
    setTimeout(() => document.addEventListener('pointerdown', function out(e) { if (!box.contains(e.target as Node)) { document.removeEventListener('pointerdown', out, true); save(); } }, true), 0);
    ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length);
  }
  private closeEditor() { this.editor?.remove(); this.editor = null; }

  // ---------- Chú thích dạng chuỗi (Ctrl+Shift+F2) ----------
  openCard(r = this.app.grid.sel.ar, c = this.app.grid.sel.ac) {
    this.closeEditor(); this.hideTip(); this.closeCard();
    const app = this.app, s = app.grid.sheet;
    const card = document.createElement('div');
    card.className = 'cm-card';
    card.setAttribute('role', 'dialog');
    card.setAttribute('aria-label', `Chú thích ô ${toA1(r, c)}`);
    const render = () => {
      const st = app.wb.style(s, r, c);
      card.innerHTML = '';
      const head = document.createElement('div'); head.className = 'cm-top';
      const ttl = document.createElement('b'); ttl.textContent = (st.cm ? 'Chú thích' : 'Chú thích mới') + ' — ' + toA1(r, c);
      head.appendChild(ttl);
      if (st.cm) {
        const done = document.createElement('button'); done.className = 'cm-act'; done.textContent = st.cmDone ? 'Mở lại' : 'Giải quyết';
        done.onclick = () => { app.wb.resolveComment(s, r, c, !st.cmDone); app.grid.invalidate(); render(); };
        const del = document.createElement('button'); del.className = 'cm-act'; del.innerHTML = ic('trash', 14); del.title = 'Xóa chuỗi'; del.setAttribute('aria-label', 'Xóa chuỗi');
        del.onclick = () => { app.wb.editComment(s, r, c, 0, null); app.grid.invalidate(); this.closeCard(); app.grid.focusKeys(); };
        head.append(done, del);
      }
      const x = document.createElement('button'); x.className = 'cm-act'; x.innerHTML = ic('x', 14); x.setAttribute('aria-label', 'Đóng'); x.onclick = () => { this.closeCard(); app.grid.focusKeys(); };
      head.appendChild(x);
      card.appendChild(head);
      (st.cm ?? []).forEach((it, i) => {
        const row = document.createElement('div'); row.className = 'cm-it';
        const h = document.createElement('div'); h.className = 'cm-h';
        const b = document.createElement('b'); b.textContent = it.a;
        const t = document.createElement('small'); t.textContent = when(it.d);
        h.append(b, t);
        if (i > 0) { const d = document.createElement('button'); d.className = 'cm-act'; d.textContent = 'Xóa'; d.onclick = () => { app.wb.editComment(s, r, c, i, null); render(); }; h.appendChild(d); }
        const p = document.createElement('p'); p.textContent = it.t;
        row.append(h, p); card.appendChild(row);
      });
      const ta = document.createElement('textarea');
      ta.placeholder = st.cm ? 'Trả lời…' : 'Bắt đầu cuộc hội thoại…';
      ta.setAttribute('aria-label', ta.placeholder);
      const send = document.createElement('button'); send.className = 'btn primary cm-send'; send.textContent = st.cm ? 'Trả lời' : 'Đăng';
      const post = () => { if (!ta.value.trim()) return; app.wb.addComment(s, r, c, authorName(), ta.value); app.grid.invalidate(); render(); };
      send.onclick = post;
      ta.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter' && e.ctrlKey) { e.preventDefault(); post(); } if (e.key === 'Escape') { this.closeCard(); app.grid.focusKeys(); } });
      const foot = document.createElement('div'); foot.className = 'cm-foot'; foot.append(ta, send);
      card.appendChild(foot);
      setTimeout(() => ta.focus(), 20);
    };
    render();
    const a = app.grid.cellPageRect(r, c);
    card.style.left = Math.min(window.innerWidth - 330, a.right + 8) + 'px';
    card.style.top = Math.max(4, Math.min(window.innerHeight - 320, a.top)) + 'px';
    document.body.appendChild(card);
    this.card = card;
  }
  closeCard() { this.card?.remove(); this.card = null; }

  // ---------- đi tới Trước / Tiếp ----------
  step(kind: 'note' | 'comment', d: 1 | -1) {
    const app = this.app, list = app.wb.noteCells(app.grid.sheet, kind);
    if (!list.length) { app.toast(kind === 'note' ? 'Trang tính không có ghi chú.' : 'Trang tính không có chú thích.'); return; }
    const { ar, ac } = app.grid.sel;
    const after = (p: { r: number; c: number }) => p.r > ar || (p.r === ar && p.c > ac);
    const before = (p: { r: number; c: number }) => p.r < ar || (p.r === ar && p.c < ac);
    // tới ô kế tiếp/trước theo thứ tự đọc, hết thì quay vòng như Excel
    const p = d > 0 ? list.find(after) ?? list[0] : [...list].reverse().find(before) ?? list[list.length - 1];
    app.grid.select(p.r, p.c); app.syncSelection();
    if (kind === 'comment') this.openCard(p.r, p.c); else this.editNote(p.r, p.c);
  }

  // ---------- ngăn Chú thích ----------
  private pane: HTMLElement | null = null;
  togglePane() {
    if (this.pane) { this.pane.remove(); this.pane = null; this.app.ribbon.refresh(); return; }
    const p = document.createElement('aside');
    p.className = 'side-pane cm-pane';
    p.setAttribute('aria-label', 'Chú thích');
    document.getElementById('gridwrap')!.appendChild(p);
    this.pane = p;
    this.syncPane();
    this.app.ribbon.refresh();
  }
  get paneOpen() { return !!this.pane; }
  private syncPane() {
    const p = this.pane;
    if (!p) return;
    const app = this.app, s = app.grid.sheet;
    p.innerHTML = `<header><b>Chú thích</b><button class="cm-new btn">${ic('comment', 14)}<span>Mới</span></button><button class="x" aria-label="Đóng">${ic('x', 14)}</button></header><div class="cm-list"></div>`;
    (p.querySelector('.x') as HTMLElement).onclick = () => this.togglePane();
    (p.querySelector('.cm-new') as HTMLElement).onclick = () => this.openCard();
    const list = p.querySelector('.cm-list')!;
    const cells = app.wb.noteCells(s, 'comment');
    if (!cells.length) { const e = document.createElement('p'); e.className = 'muted'; e.textContent = 'Trang tính chưa có chú thích. Chọn ô rồi bấm “Mới” (Ctrl+Shift+F2).'; list.appendChild(e); return; }
    for (const c of cells) {
      const st = app.wb.style(s, c.r, c.c);
      const b = document.createElement('button'); b.className = 'cm-row' + (st.cmDone ? ' done' : '');
      const ref = document.createElement('small'); ref.textContent = toA1(c.r, c.c) + (st.cmDone ? ' · Đã giải quyết' : '') + ` · ${st.cm!.length} bình luận`;
      const who = document.createElement('b'); who.textContent = st.cm![0].a;
      const t = document.createElement('span'); t.textContent = st.cm![0].t;
      b.append(ref, who, t);
      b.onclick = () => { app.grid.select(c.r, c.c); app.syncSelection(); this.openCard(c.r, c.c); };
      list.appendChild(b);
    }
  }
}
