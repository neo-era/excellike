// Dạng xem Bố trí Trang (Xem ▸ Bố trí Trang): các trang giấy như khi in, thước đo, lề,
// sửa đầu trang/chân trang ngay trên trang (bấm vào vùng đầu/chân trang).
// Nhấp đúp vào ô trên trang → về dạng Thường tại ô đó để sửa (ô không sửa trực tiếp trên trang).
import { buildJob } from '../core/print';
import { hfFor, hfJoin, hfParse, normPrint, storePrint, type NPrint } from '../core/paginate';
import { chartImgs } from './printui';
import { rawToFriendly, friendlyToRaw } from './pagesetup';
import type { App } from './app';

export class PageLayoutView {
  el: HTMLDivElement;
  private root: ShadowRoot;
  private timer = 0;
  on = false;
  constructor(private app: App) {
    this.el = document.createElement('div');
    this.el.className = 'pl-view';
    this.el.hidden = true;
    this.el.setAttribute('role', 'region');
    this.el.setAttribute('aria-label', 'Bố trí Trang');
    this.root = this.el.attachShadow({ mode: 'open' });
    document.getElementById('gridwrap')!.appendChild(this.el);
    app.wb.onChange(() => this.schedule());
    this.el.addEventListener('dblclick', (e) => {
      const td = (e.composedPath()[0] as HTMLElement).closest?.('td[data-rc]') as HTMLElement | null;
      if (!td) return;
      const [r, c] = td.dataset.rc!.split(',').map(Number);
      this.app.setView('normal');
      this.app.grid.select(r, c);
      this.app.syncSelection();
      this.app.grid.focusKeys();
    });
  }
  /** sổ đổi (kể cả đổi sổ) → vẽ lại */
  rebind() { this.app.wb.onChange(() => this.schedule()); this.schedule(); }
  private schedule() { if (!this.on) return; clearTimeout(this.timer); this.timer = window.setTimeout(() => this.render(), 120); }
  show() { this.on = true; this.el.hidden = false; this.render(); }
  hide() { this.on = false; this.el.hidden = true; this.root.innerHTML = ''; }
  render() {
    if (!this.on) return;
    const app = this.app, wb = app.wb, s = app.grid.sheet;
    const keepTop = this.el.scrollTop, keepLeft = this.el.scrollLeft;
    const job = buildJob(wb, [{ sheet: s, charts: chartImgs(wb, s) }], { file: wb.name });
    const doc = new DOMParser().parseFromString(job.html, 'text/html');
    const style = doc.querySelector('style')!.textContent ?? '';
    const ps: NPrint = normPrint(wb.meta(s).print);
    const z = app.grid.zoom;
    this.root.innerHTML = `<style>${style}
:host { all: initial; }
.wrap { padding: 28px 24px 40px 44px; min-width: max-content; font: 12px 'Segoe UI', system-ui, sans-serif; color: #222; background: #e6e6e6; min-height: 100%; box-sizing: border-box; zoom: ${z}; }
.pg { position: relative; margin: 0 auto 28px; width: max-content; }
.page { margin: 0 !important; box-shadow: 0 1px 4px rgba(0,0,0,.28); }
.ruler-h { position: absolute; left: 0; right: 0; top: -20px; height: 18px; background: #fafafa; border: 1px solid #c8c8c8; box-sizing: border-box; overflow: hidden; }
.ruler-v { position: absolute; top: 0; bottom: 0; left: -22px; width: 18px; background: #fafafa; border: 1px solid #c8c8c8; box-sizing: border-box; overflow: hidden; }
.ruler-h i, .ruler-v i { position: absolute; background: #9a9a9a; }
.ruler-h b, .ruler-v b { position: absolute; font: 9px 'Segoe UI', sans-serif; color: #555; font-weight: 400; }
.ruler-h .m, .ruler-v .m { background: #dcdcdc; position: absolute; }
.hfz { position: absolute; left: 0; right: 0; display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 4px; z-index: 2; }
.hfz div { min-height: 22px; padding: 2px 4px; outline: none; border: 1px dashed transparent; white-space: pre-wrap; font: 11pt 'Aptos Narrow', Calibri, sans-serif; }
.hfz:hover div { border-color: #b8b8b8; }
.hfz div:focus { border-color: #217346; background: #fff; }
.hfz div.l { text-align: left; } .hfz div.c { text-align: center; } .hfz div.r { text-align: right; }
.hfz .ph { color: #8a8a8a; font-style: italic; }
.page .hf { visibility: hidden; }
td[data-rc] { cursor: cell; }
.tip { text-align: center; color: #555; margin: 0 0 34px; }
</style><div class="wrap"><p class="tip">Bố trí Trang — bấm vào vùng đầu/chân trang để sửa; nhấp đúp vào ô để sửa ô.</p></div>`;
    const wrap = this.root.querySelector('.wrap')!;
    const pages = [...doc.querySelectorAll('section.page')];
    if (!pages.length) { wrap.append(Object.assign(document.createElement('p'), { className: 'tip', textContent: 'Trang tính trống — chưa có trang in nào.' })); return; }
    const mm = (x: string) => parseFloat(x);
    for (const sec of pages) {
      const pg = document.createElement('div');
      pg.className = 'pg';
      const page = document.importNode(sec, true) as HTMLElement;
      pg.append(page);
      const wmm = mm(page.style.width), hmm = mm(page.style.height);
      pg.append(this.ruler('h', wmm, ps.mg.l * 25.4, wmm - ps.mg.r * 25.4), this.ruler('v', hmm, ps.mg.t * 25.4, hmm - ps.mg.b * 25.4));
      const idx = +(page.dataset.i ?? 0), num = +(page.dataset.n ?? 1);
      pg.append(this.zone('header', idx, num, ps), this.zone('footer', idx, num, ps));
      wrap.append(pg);
    }
    this.el.scrollTop = keepTop; this.el.scrollLeft = keepLeft;
  }
  /** thước đo (cm) với phần lề tô xám */
  private ruler(ax: 'h' | 'v', lenMm: number, a: number, b: number) {
    const r = document.createElement('div');
    r.className = 'ruler-' + ax;
    const pos = (v: number) => `${(v / lenMm) * 100}%`;
    const m1 = document.createElement('span'); m1.className = 'm';
    const m2 = document.createElement('span'); m2.className = 'm';
    if (ax === 'h') { Object.assign(m1.style, { left: '0', width: pos(a), top: '0', bottom: '0' }); Object.assign(m2.style, { left: pos(b), right: '0', top: '0', bottom: '0' }); }
    else { Object.assign(m1.style, { top: '0', height: pos(a), left: '0', right: '0' }); Object.assign(m2.style, { top: pos(b), bottom: '0', left: '0', right: '0' }); }
    r.append(m1, m2);
    for (let v = 0; v <= lenMm; v += 5) {
      const t = document.createElement('i'), big = v % 10 === 0;
      if (ax === 'h') Object.assign(t.style, { left: pos(v), bottom: '0', width: '1px', height: big ? '7px' : '4px' });
      else Object.assign(t.style, { top: pos(v), right: '0', height: '1px', width: big ? '7px' : '4px' });
      r.append(t);
      if (big && v > 0) {
        const n = document.createElement('b'); n.textContent = String(Math.round((v - a) / 10));
        if (ax === 'h') Object.assign(n.style, { left: `calc(${pos(v)} - 4px)`, top: '1px' }); else Object.assign(n.style, { top: `calc(${pos(v)} - 6px)`, left: '2px' });
        r.append(n);
      }
    }
    return r;
  }
  /** vùng sửa đầu/chân trang của một trang (đúng biến thể: trang đầu / chẵn / thường) */
  private zone(kind: 'header' | 'footer', idx: number, num: number, ps: NPrint) {
    const z = document.createElement('div');
    z.className = 'hfz';
    const inch = (v: number) => `${v}in`;
    if (kind === 'header') Object.assign(z.style, { top: inch(ps.mg.h), left: inch(ps.mg.l), right: inch(ps.mg.r) });
    else Object.assign(z.style, { bottom: inch(ps.mg.f), left: inch(ps.mg.l), right: inch(ps.mg.r) });
    const which: 'first' | 'even' | 'odd' = idx === 0 && ps.hfFirst ? 'first' : num % 2 === 0 && ps.hfEven ? 'even' : 'odd';
    const cur = hfFor(ps, idx, num)[kind];
    const parts = hfParse(cur);
    const label = { first: 'trang đầu', even: 'trang chẵn', odd: '' }[which];
    for (const k of ['l', 'c', 'r'] as const) {
      const d = document.createElement('div');
      d.className = k;
      d.contentEditable = 'true';
      d.spellcheck = false;
      d.setAttribute('role', 'textbox');
      d.setAttribute('aria-label', `${kind === 'header' ? 'Đầu trang' : 'Chân trang'} ${label} – phần ${k === 'l' ? 'trái' : k === 'c' ? 'giữa' : 'phải'}`.replace('  ', ' '));
      const txt = rawToFriendly(parts[k] ?? '');
      if (txt) d.textContent = txt;
      else if (k === 'c') { d.innerHTML = `<span class="ph">${kind === 'header' ? 'Thêm đầu trang' : 'Thêm chân trang'}</span>`; }
      d.addEventListener('focus', () => { if (d.querySelector('.ph')) d.textContent = ''; });
      d.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); this.render(); } if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); d.blur(); } e.stopPropagation(); });
      d.addEventListener('blur', () => {
        const v = d.querySelector('.ph') ? '' : (d.textContent ?? '');
        const np: Record<'l' | 'c' | 'r', string> = { l: parts.l ?? '', c: parts.c ?? '', r: parts.r ?? '' };
        np[k] = friendlyToRaw(v);
        const raw = hfJoin(np);
        if (raw === cur) { if (!v) this.render(); return; }
        const s = this.app.grid.sheet, p = normPrint(this.app.wb.meta(s).print);
        if (which === 'first') p.hfFirst = { ...p.hfFirst!, [kind]: raw };
        else if (which === 'even') p.hfEven = { ...p.hfEven!, [kind]: raw };
        else p[kind] = raw;
        this.app.safeSetPrint(s, storePrint(p));
      });
      z.append(d);
    }
    return z;
  }
}
