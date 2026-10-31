// Trang In trong Backstage (Tệp ▸ In) như Excel 365: Số bản, Máy in, In gì, Trang từ–đến, Đối chiếu, Hướng, Khổ giấy,
// Lề, Co giãn, Thiết lập Trang; bên phải là bản xem trước từng trang (◀ 1 / N ▶, Hiện Lề, Thu phóng Trang).
// Bản Windows in thẳng ra máy in đã chọn (không qua hộp thoại hệ thống) và xuất PDF thật; bản web in qua trình duyệt.
import type { Range } from '../core/address';
import { chartData } from '../core/chart';
import { buildJob, type ChartImg, type JobItem } from '../core/print';
import { normPrint, PAPERS, MARGIN_PRESETS, storePrint, type NPrint } from '../core/paginate';
import type { Workbook } from '../core/workbook';
import { chartPng } from './chartdraw';
import { desk } from './files';
import { ic } from './icons';

export function chartImgs(wb: Workbook, s: number): ChartImg[] {
  return wb.meta(s).charts.map((ch) => ({ src: ch.type === 'pic' && ch.img ? ch.img : chartPng(ch, chartData(wb, ch)), w: ch.w, h: ch.h, anchor: ch.anchor }));
}

/** In qua khung ẩn (bản web): trình duyệt mở hộp thoại in, có “Lưu dưới dạng PDF” */
export function printViaFrame(html: string) {
  document.querySelector('iframe.print-frame')?.remove();
  const f = document.createElement('iframe');
  f.className = 'print-frame';
  f.setAttribute('aria-hidden', 'true');
  f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
  document.body.appendChild(f);
  f.onload = () => {
    const w = f.contentWindow!;
    const imgs = [...w.document.images];
    Promise.all(imgs.map((im) => (im.complete ? 1 : new Promise((r) => { im.onload = im.onerror = r; })))).then(() => { w.focus(); w.print(); setTimeout(() => window.focus(), 0); });
  };
  f.srcdoc = html;
}

export interface PrintHost {
  wb: Workbook;
  sheet: () => number;
  selection: () => Range;
  fileName: () => string;
  toast: (m: string, k?: 'ok' | 'err' | '') => void;
  pageSetup: (tab: 'page' | 'margins' | 'hf' | 'sheet') => void;
}

const SCALE_OPTS: [string, string][] = [['none', 'Không Co giãn'], ['sheet', 'Vừa Trang tính trên Một Trang'], ['cols', 'Vừa Tất cả Cột trên Một Trang'], ['rows', 'Vừa Tất cả Hàng trên Một Trang'], ['custom', 'Tùy chọn Co giãn Tùy chỉnh…']];
const scaleKey = (p: NPrint) => (!p.fit ? (p.scale === 100 ? 'none' : 'custom') : p.fit.w === 1 && p.fit.h === 1 ? 'sheet' : p.fit.w === 1 && !p.fit.h ? 'cols' : !p.fit.w && p.fit.h === 1 ? 'rows' : 'custom');

export class PrintPage {
  readonly el: HTMLElement;
  private frame!: HTMLIFrameElement;
  private cur = 1;
  private total = 0;
  private timer = 0;
  private zoomed = false;
  private margins = false;
  private printers: { name: string; displayName?: string; isDefault?: boolean }[] = [];
  constructor(private host: PrintHost) {
    this.el = document.createElement('div');
    this.el.className = 'bs-print';
    this.build();
  }
  private get s() { return this.host.sheet(); }
  private ps(): NPrint { return normPrint(this.host.wb.meta(this.s).print); }
  private setPs(p: Partial<NPrint>) {
    const n = storePrint({ ...this.ps(), ...p });
    if (JSON.stringify(n) !== JSON.stringify(storePrint(this.ps()))) this.host.wb.setPrint(this.s, n);
    this.sync();
    this.refresh();
  }
  private $<T extends HTMLElement>(id: string) { return this.el.querySelector('#' + id) as T; }
  private build() {
    const paper = Object.entries(PAPERS).map(([k, v]) => `<option value="${k}">${v.label}</option>`).join('');
    this.el.innerHTML = `
<div class="bs-pset">
  <div class="prow"><button class="bigprint" id="bs-print">${ic('printer2', 30)}<span>In</span></button>
    <label class="fld" style="flex:1"><span>Số bản:</span><input type="number" id="bs-copies" min="1" max="999" value="1"></label></div>
  <h3>Máy in</h3>
  <select id="bs-printer" aria-label="Máy in"></select>
  <button class="link" id="bs-props">Thuộc tính Máy in</button>
  <h3>Thiết đặt</h3>
  <select id="bs-what" aria-label="In gì"><option value="active">In Trang tính Hiện hoạt</option><option value="workbook">In Toàn bộ Sổ làm việc</option><option value="selection">In Vùng chọn</option></select>
  <div class="two"><span>Trang:</span><input type="text" inputmode="numeric" id="bs-from" aria-label="Từ trang"><span>đến</span><input type="text" inputmode="numeric" id="bs-to" aria-label="Đến trang"></div>
  <select id="bs-collate" aria-label="Đối chiếu"><option value="1">Đã đối chiếu — 1,2,3 1,2,3 1,2,3</option><option value="0">Chưa đối chiếu — 1,1,1 2,2,2 3,3,3</option></select>
  <select id="bs-orient" aria-label="Hướng"><option value="portrait">Hướng Dọc</option><option value="landscape">Hướng Ngang</option></select>
  <select id="bs-paper" aria-label="Khổ giấy">${paper}</select>
  <select id="bs-margins" aria-label="Lề"><option value="normal">Lề Thường — T 1,91 cm D 1,91 cm Tr 1,78 cm P 1,78 cm</option><option value="wide">Lề Rộng — T 2,54 cm D 2,54 cm Tr 2,54 cm P 2,54 cm</option><option value="narrow">Lề Hẹp — T 1,91 cm D 1,91 cm Tr 0,64 cm P 0,64 cm</option><option value="custom">Lề Tùy chỉnh…</option></select>
  <select id="bs-scale" aria-label="Co giãn">${SCALE_OPTS.map(([v, l]) => `<option value="${v}">${l}</option>`).join('')}</select>
  <button class="link" id="bs-setup">Thiết lập Trang</button>
  <p class="hint" id="bs-hint"></p>
</div>
<div class="bs-preview">
  <iframe title="Xem trước khi in" sandbox="allow-same-origin"></iframe>
  <div class="pv-bar">
    <button id="pv-prev" title="Trang Trước" aria-label="Trang Trước">${ic('back', 16)}</button>
    <input id="pv-num" aria-label="Trang hiện tại" value="1"><span>/ <span id="pv-total">0</span></span>
    <button id="pv-next" title="Trang Tiếp theo" aria-label="Trang Tiếp theo">${ic('fwd', 16)}</button>
    <span class="sp"></span>
    <button id="pv-margins" title="Hiện Lề" aria-label="Hiện Lề" aria-pressed="false">${ic('margins', 16)}</button>
    <button id="pv-zoom" title="Thu phóng Trang" aria-label="Thu phóng Trang" aria-pressed="false">${ic('fitPage', 16)}</button>
  </div>
</div>`;
    this.frame = this.el.querySelector('iframe')!;
    const on = (id: string, ev: string, f: (e: Event) => void) => this.$(id).addEventListener(ev, f);
    on('bs-print', 'click', () => this.print());
    on('bs-props', 'click', () => this.printerProps());
    on('bs-setup', 'click', () => this.host.pageSetup('page'));
    on('bs-what', 'change', () => { this.cur = 1; this.refresh(); });
    on('bs-orient', 'change', () => this.setPs({ orientation: this.$<HTMLSelectElement>('bs-orient').value as NPrint['orientation'] }));
    on('bs-paper', 'change', () => this.setPs({ paper: this.$<HTMLSelectElement>('bs-paper').value as NPrint['paper'] }));
    on('bs-margins', 'change', () => {
      const v = this.$<HTMLSelectElement>('bs-margins').value as NPrint['margins'];
      if (v === 'custom') { this.sync(); this.host.pageSetup('margins'); return; }
      this.setPs({ margins: v, mg: { ...MARGIN_PRESETS[v] } });
    });
    on('bs-scale', 'change', () => {
      const v = this.$<HTMLSelectElement>('bs-scale').value;
      if (v === 'custom') { this.sync(); this.host.pageSetup('page'); return; }
      this.setPs(v === 'none' ? { fit: null, scale: 100 } : { fit: v === 'sheet' ? { w: 1, h: 1 } : v === 'cols' ? { w: 1, h: 0 } : { w: 0, h: 1 } });
    });
    const go = (n: number) => { this.cur = Math.max(1, Math.min(this.total || 1, n)); this.showPage(); };
    on('pv-prev', 'click', () => go(this.cur - 1));
    on('pv-next', 'click', () => go(this.cur + 1));
    on('pv-num', 'change', () => go(parseInt(this.$<HTMLInputElement>('pv-num').value) || 1));
    on('pv-num', 'keydown', (e) => { if ((e as KeyboardEvent).key === 'Enter') go(parseInt(this.$<HTMLInputElement>('pv-num').value) || 1); });
    on('pv-margins', 'click', () => { this.margins = !this.margins; this.$('pv-margins').setAttribute('aria-pressed', String(this.margins)); this.showPage(); });
    on('pv-zoom', 'click', () => { this.zoomed = !this.zoomed; this.$('pv-zoom').setAttribute('aria-pressed', String(this.zoomed)); this.showPage(); });
    this.frame.addEventListener('load', () => this.showPage());
    new ResizeObserver(() => this.showPage()).observe(this.frame);
    this.loadPrinters();
  }
  private async loadPrinters() {
    const sel = this.$<HTMLSelectElement>('bs-printer');
    const d = desk();
    if (d?.printers) {
      try { this.printers = await d.printers(); } catch { this.printers = []; }
    }
    sel.innerHTML = '';
    if (this.printers.length) {
      for (const p of this.printers) { const o = document.createElement('option'); o.value = p.name; o.textContent = (p.displayName || p.name) + (p.isDefault ? ' (mặc định)' : ''); if (p.isDefault) o.selected = true; sel.appendChild(o); }
    } else {
      const o = document.createElement('option');
      o.value = '';
      o.textContent = d ? 'Máy in mặc định của Windows' : 'Máy in của trình duyệt (chọn ở bước sau)';
      sel.appendChild(o);
    }
    if (d) { const o = document.createElement('option'); o.value = '__pdf'; o.textContent = 'Lưu thành PDF'; sel.appendChild(o); }
    this.$('bs-hint').textContent = d ? '' : 'Bản trình duyệt: bấm In để mở hộp thoại in của trình duyệt (chọn máy in hoặc “Lưu dưới dạng PDF”). Số bản chọn ở hộp thoại đó.';
  }
  /** Đồng bộ các ô chọn với thiết lập trang hiện tại */
  sync() {
    const p = this.ps();
    this.$<HTMLSelectElement>('bs-orient').value = p.orientation;
    this.$<HTMLSelectElement>('bs-paper').value = p.paper;
    this.$<HTMLSelectElement>('bs-margins').value = p.margins;
    this.$<HTMLSelectElement>('bs-scale').value = scaleKey(p);
    const custom = this.$<HTMLSelectElement>('bs-scale').querySelector('option[value=custom]')!;
    custom.textContent = scaleKey(p) === 'custom' ? (p.fit ? `Vừa ${p.fit.w || 'tự động'} × ${p.fit.h || 'tự động'} trang` : `Tỷ lệ ${p.scale}%`) : 'Tùy chọn Co giãn Tùy chỉnh…';
  }
  private items(): JobItem[] {
    const wb = this.host.wb, what = this.$<HTMLSelectElement>('bs-what').value;
    const ids = what === 'workbook' ? wb.visibleSheetIds() : [this.s];
    const range = what === 'selection' ? this.host.selection() : undefined;
    return ids.map((id) => ({ sheet: id, range, charts: chartImgs(wb, id) }));
  }
  private job(pages?: [number, number]) { return buildJob(this.host.wb, this.items(), { file: this.host.fileName(), pages }); }
  refresh() {
    clearTimeout(this.timer);
    this.timer = window.setTimeout(() => this.render(), 60);
  }
  private render() {
    this.sync();
    const j = this.job();
    this.total = j.total;
    this.$('pv-total').textContent = String(this.total);
    this.cur = Math.max(1, Math.min(this.total || 1, this.cur));
    this.frame.srcdoc = this.total ? j.html : `<!doctype html><html><body style="margin:0;display:grid;place-items:center;height:100vh;font:14px 'Segoe UI',system-ui,sans-serif;color:#555;background:#e6e6e6">Không tìm thấy nội dung để in.</body></html>`;
  }
  private showPage() {
    const d = this.frame.contentDocument;
    if (!d?.body) return;
    this.$<HTMLInputElement>('pv-num').value = String(this.total ? this.cur : 0);
    const pages = [...d.querySelectorAll<HTMLElement>('.page')];
    d.body.classList.add('one');
    d.body.classList.toggle('show-margins', this.margins);
    pages.forEach((p, i) => p.classList.toggle('cur', i === this.cur - 1));
    const p = pages[this.cur - 1];
    if (!p) return;
    // vừa khung xem trước (Thu phóng Trang: 100%)
    d.body.style.zoom = '1';
    const r = p.getBoundingClientRect();
    const fw = this.frame.clientWidth - 32, fh = this.frame.clientHeight - 32;
    const z = this.zoomed ? 1 : Math.min(fw / r.width, fh / r.height, 1.4);
    d.body.style.zoom = String(Math.max(0.1, z));
  }
  private range(): [number, number] | undefined {
    const a = parseInt(this.$<HTMLInputElement>('bs-from').value), b = parseInt(this.$<HTMLInputElement>('bs-to').value);
    if (!a && !b) return undefined;
    return [a || 1, b || 9999];
  }
  async print() {
    const j = this.job(this.range());
    if (!j.total) { this.host.toast('Không có gì để in.', 'err'); return; }
    if (!/class="page"/.test(j.html)) { this.host.toast('Không có trang nào trong khoảng đã chọn.', 'err'); return; }
    const d = desk();
    const printer = this.$<HTMLSelectElement>('bs-printer').value;
    const copies = Math.max(1, Math.min(999, parseInt(this.$<HTMLInputElement>('bs-copies').value) || 1));
    if (d && printer === '__pdf') { await this.exportPdf(j.html); return; }
    if (d?.print) {
      this.host.toast('Đang gửi tới máy in…');
      const used = j.layouts.filter((L) => L.pages.length);
      const mixed = new Set(used.map((L) => L.setup.paper + L.setup.orientation)).size > 1;
      const st = mixed ? undefined : used[0]?.setup;
      const r = await d.print({ html: j.html, deviceName: printer || undefined, copies, collate: this.$<HTMLSelectElement>('bs-collate').value === '1', silent: true, paper: st?.paper, landscape: st?.orientation === 'landscape' });
      if (r?.ok) this.host.toast(`Đã gửi ${copies > 1 ? copies + ' bản ' : ''}tới ${printer || 'máy in mặc định'}.`, 'ok');
      else if (r?.error !== 'cancelled') this.host.toast('Không in được: ' + (r?.error || 'máy in không phản hồi'), 'err');
      return;
    }
    printViaFrame(j.html);
  }
  async exportPdf(html?: string) {
    const h = html ?? this.job(this.range()).html;
    const d = desk();
    if (d?.pdf) {
      const r = await d.pdf({ html: h, name: (this.host.fileName().replace(/\.[^.]+$/, '') || 'Sổ làm việc') + '.pdf' });
      if (r?.path) this.host.toast('Đã xuất PDF: ' + r.path, 'ok');
      else if (r?.error) this.host.toast('Không xuất được PDF: ' + r.error, 'err');
      return;
    }
    this.host.toast('Ở hộp thoại in, chọn máy in “Lưu dưới dạng PDF”.');
    printViaFrame(h);
  }
  private async printerProps() {
    const d = desk();
    const j = this.job(this.range());
    const st = j.layouts.find((L) => L.pages.length)?.setup;
    if (d?.print) { await d.print({ html: j.html, deviceName: this.$<HTMLSelectElement>('bs-printer').value || undefined, silent: false, paper: st?.paper, landscape: st?.orientation === 'landscape' }); return; }
    printViaFrame(j.html);
  }
}
