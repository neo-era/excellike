import { MAX_COLS, MAX_ROWS, parseRange, rangeToA1, toA1, type Range } from '../core/address';
import { quoteSheet } from '../core/formula';
import { Workbook, type Style } from '../core/workbook';
import { Grid, ptPx, type EditHost } from './grid';
import { ic } from './icons';
import { FUNC_INFO } from './funcs';
import { saveBook, lsGet, lsSet } from './store';
import { sortDialog, cfRuleDialog, cfManageDialog, dvDialog, h, openDialog, closeDialog, field } from './dialogs';
import { pickFile, saveBytes, baseName, desk, deskToPicked } from './files';
import type { BorderKind, BorderStyle } from '../core/workbook';
import { ChartLayer, chartDialog } from './charts';
import type { ChartType } from '../core/chart';
import { PivotPanel, pivotCreateDialog } from './pivotui';
import { Ribbon, KeyTips, allItems, fold, type Tab, type Btn, type KtTarget, type Item, type Group } from './ribbon';
import { closeMenu, openPopover, showMenu, menuOpen, type MenuItem } from './menu';
import { Backstage } from './backstage';
import { pageSetupDialog } from './pagesetup';
import { normPrint, storePrint } from '../core/paginate';

export { closeMenu, openPopover, showMenu };
const $ = <T extends HTMLElement = HTMLElement>(sel: string, root: ParentNode = document) => root.querySelector(sel) as T;
const isTouch = () => matchMedia('(pointer: coarse)').matches;
const COMPACT = '(max-width: 760px), (max-height: 520px)';
import { VERSION } from './version';
import { formatValue } from '../core/numfmt';

// ---- bảng màu như Excel: Màu Chủ đề (Office) + sắc độ, Màu Chuẩn ----
const THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
function shade(hex: string, t: number) {
  const n = parseInt(hex.slice(1), 16), c = [n >> 16, (n >> 8) & 255, n & 255];
  const f = (v: number) => Math.round(t >= 0 ? v + (255 - v) * t : v * (1 + t));
  return '#' + c.map((v) => f(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}
const TINTS: Record<number, number[]> = { 0: [-0.05, -0.15, -0.25, -0.35, -0.5], 1: [0.5, 0.35, 0.25, 0.15, 0.05] };
const tintsOf = (i: number) => (TINTS[i] ?? [0.8, 0.6, 0.4, -0.25, -0.5]);
/** Bảng chọn màu kiểu Excel */
export function colorPalette(onPick: (c: string | undefined) => void, noneLabel: string): HTMLElement {
  const box = document.createElement('div');
  box.className = 'palette';
  const sw = (c: string) => `<button class="sw" style="--sw:${c}" data-c="${c}" title="${c}" aria-label="${c}"></button>`;
  let html = `<button class="pl-none" data-c="">${noneLabel}</button><h4>Màu Chủ đề</h4><div class="pl-grid">${THEME.map(sw).join('')}</div><div class="pl-grid tints">`;
  for (let k = 0; k < 5; k++) html += THEME.map((c, i) => sw(shade(c, tintsOf(i)[k]))).join('');
  html += `</div><h4>Màu Chuẩn</h4><div class="pl-grid">${STANDARD.map(sw).join('')}</div><label class="pl-more">${ic('palette', 16)}<span>Màu Khác…</span><input type="color" aria-label="Màu khác"></label>`;
  box.innerHTML = html;
  box.onclick = (e) => {
    const b = (e.target as HTMLElement).closest('button');
    if (!b) return;
    closeMenu();
    onPick(b.dataset.c || undefined);
  };
  (box.querySelector('input[type=color]') as HTMLInputElement).onchange = (e) => { closeMenu(); onPick((e.target as HTMLInputElement).value); };
  return box;
}

// ---- định dạng số trong hộp Định dạng Số (như danh sách của Excel) ----
const NF: [string, string | undefined, string][] = [
  ['Chung', undefined, 'abc'],
  ['Số', '0.00', '1234,50'],
  ['Tiền tệ', '#,##0" ₫"', '1.235 ₫'],
  ['Kế toán', '_-* #,##0 "₫"_-;-* #,##0 "₫"_-;_-* "-" "₫"_-;_-@_-', ' 1.235 ₫ '],
  ['Ngày dạng Ngắn', 'dd/mm/yyyy', '29/10/2026'],
  ['Ngày dạng Dài', 'dddd, dd mmmm yyyy', 'Thứ Năm, 29 Tháng 10 2026'],
  ['Thời gian', 'hh:mm:ss', '12:00:00'],
  ['Phần trăm', '0.00%', '15,25%'],
  ['Phân số', '# ?/?', '1234 1/2'],
  ['Khoa học', '0.00E+00', '1,23E+03'],
  ['Văn bản', '@', 'abc'],
];
export function nfName(f: string | undefined) {
  if (!f || f === 'General') return 'Chung';
  if (f === '@') return 'Văn bản';
  if (/%/.test(f)) return 'Phần trăm';
  if (/E\+/i.test(f)) return 'Khoa học';
  if (/\?\/[?\d]/.test(f)) return 'Phân số';
  if (/^_-\*|\*\s/.test(f)) return 'Kế toán';
  if (/₫|\$|€|\[\$/.test(f)) return 'Tiền tệ';
  if (/dddd|mmmm|tháng/.test(f)) return 'Ngày dạng Dài';
  if (/[dy]/i.test(f.replace(/"[^"]*"/g, ''))) return 'Ngày dạng Ngắn';
  if (/h.*m|m.*s/i.test(f.replace(/"[^"]*"/g, ''))) return 'Thời gian';
  if (/^#,##0$/.test(f)) return 'Số';
  return /[0#]/.test(f) ? 'Số' : 'Tùy chỉnh';
}
const FONTS = ['Aptos Narrow', 'Aptos', 'Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana', 'Roboto', 'Be Vietnam Pro'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const DEFAULT_FONT = 'Aptos Narrow';
/** Kiểu Ô dựng sẵn (Home ▸ Kiểu Ô) */
const CELL_STYLES: [string, Style, BorderKind?, BorderStyle?, string?][] = [
  ['Thường', { b: undefined, i: undefined, fg: undefined, bg: undefined, fs: undefined }],
  ['Tốt', { bg: '#C6EFCE', fg: '#006100' }],
  ['Xấu', { bg: '#FFC7CE', fg: '#9C0006' }],
  ['Trung tính', { bg: '#FFEB9C', fg: '#9C5700' }],
  ['Đầu vào', { bg: '#FFCC99', fg: '#3F3F76' }, 'all', 'thin', '#7F7F7F'],
  ['Đầu ra', { bg: '#F2F2F2', fg: '#3F3F3F', b: true }, 'all', 'thin', '#3F3F3F'],
  ['Tính toán', { bg: '#F2F2F2', fg: '#FA7D00', b: true }, 'all', 'thin', '#7F7F7F'],
  ['Ghi chú', { bg: '#FFFFCC' }, 'all', 'thin', '#B2B2B2'],
  ['Văn bản Cảnh báo', { fg: '#FF0000' }],
  ['Tựa đề', { fs: 18, fg: '#44546A' }],
  ['Đầu đề 1', { fs: 15, b: true, fg: '#44546A' }, 'bottom', 'thick', '#4472C4'],
  ['Đầu đề 2', { fs: 13, b: true, fg: '#44546A' }, 'bottom', 'thick', '#A9C4EB'],
  ['Đầu đề 3', { b: true, fg: '#44546A' }, 'bottom', 'medium', '#A2B8E1'],
  ['Tổng', { b: true }, 'bottom', 'double', '#4472C4'],
];

export class App {
  wb: Workbook;
  grid!: Grid;
  ed!: CellEditor;
  charts!: ChartLayer;
  pivot!: PivotPanel;
  bookId: string;
  private saveTimer = 0;
  private installEvt: any = null;

  constructor(wb: Workbook, bookId: string) {
    this.wb = wb;
    this.bookId = bookId;
    this.mount();
  }

  // ---------- dựng giao diện ----------
  ribbon!: Ribbon;
  keytips!: KeyTips;
  backstage!: Backstage;
  private compactQ = matchMedia(COMPACT);
  get simple() { return this.compactQ.matches; }
  private mount() {
    document.querySelectorAll<HTMLElement>('[data-ic]').forEach((el) => el.insertAdjacentHTML('afterbegin', ic(el.dataset.ic!, +(el.dataset.s || 18))));
    this.grid = new Grid($('#gridwrap'), this.wb);
    if (isTouch()) this.grid.setZoom(1.25);
    this.ed = new CellEditor(this);
    this.grid.edit = this.ed;
    this.grid.onSelect = () => this.syncSelection();
    this.grid.onContext = (x, y, zone) => this.contextMenu(x, y, zone);
    this.charts = new ChartLayer($('#gridwrap'), this.wb, this.grid, (id) => this.editChart(id), (m) => this.toast(m));
    this.grid.onDraw = () => this.charts.sync();
    this.pivot = new PivotPanel($('#gridwrap'), this.wb, (m, k) => this.toast(m, k), () => this.grid.invalidate());
    $('#gridwrap').addEventListener('grid-error', (e) => this.toast((e as CustomEvent).detail, 'err'));
    $('#gridwrap').addEventListener('filter-btn', (e) => { const d = (e as CustomEvent).detail; this.filterPopover(d.col, d.x, d.y); });
    $('#gridwrap').addEventListener('dv-btn', (e) => { const d = (e as CustomEvent).detail; this.dvPick(d.x, d.y); });
    const warn = sessionStorage.getItem('xl-warn');
    if (warn) { sessionStorage.removeItem('xl-warn'); setTimeout(() => this.toast('Lưu ý khi mở tệp: ' + warn, 'err'), 400); }
    this.wb.onChange(() => this.onBookChange());
    this.backstage = new Backstage(this);
    this.ribbon = new Ribbon(this.ribbonHost(), $('#ribbon'));
    this.renderTabs();
    this.ribbon.render();
    this.compactQ.addEventListener('change', () => { this.ribbonPeek = false; this.renderTabs(); this.ribbon.render(); this.applyRibbon(); });
    this.grid.scroller.addEventListener('pointerdown', () => { if (this.ribbonPeek) { this.ribbonPeek = false; this.applyRibbon(); } }, true);
    this.keytips = new KeyTips(() => this.ktTop(), () => this.ktTab(), () => !document.querySelector('.dlg-back') && this.backstage.isOpen() === false);
    this.renderSheets();
    this.bindTop();
    this.bindTell();
    this.bindFoot();
    this.bindNameBox();
    this.bindClipboard();
    this.bindTheme();
    this.bindInstall();
    this.bindDesktop();
    this.bindMobile();
    this.syncSelection();
    this.grid.focusKeys();
    // lưu ngay khi rời trang / ẩn app; Ctrl+S ở bất kỳ đâu
    const flush = () => { if (this.dirtySave) this.saveNow(false); };
    window.addEventListener('pagehide', flush);
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });
    document.addEventListener('keydown', (e) => {
      const ctrl = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
      if (this.backstage.isOpen()) {
        if (e.key === 'Escape' && !document.querySelector('.dlg-back') && !menuOpen()) { e.preventDefault(); this.backstage.close(); }
        if (ctrl && k === 'p') e.preventDefault();
        return;
      }
      if (ctrl && k === 's') {
        e.preventDefault(); if (this.ed.isEditing()) this.ed.commit(0, 0);
        this.save();
      }
      if (ctrl && k === 'o') { e.preventDefault(); this.openFile(); }
      if (ctrl && k === 'n' && !e.shiftKey) { e.preventDefault(); this.newBook(); }
      const tg = e.target as HTMLElement | null, inField = !!tg && (/^(INPUT|TEXTAREA|SELECT)$/.test(tg.tagName) || tg.isContentEditable);
      if (ctrl && !inField && !document.querySelector('.dlg-back') && (k === 'z' || k === 'y')) {
        e.preventDefault(); if (k === 'y' || e.shiftKey) this.redo(); else this.undo(); return;
      }
      if (ctrl && e.key === 'F1') { e.preventDefault(); this.toggleRibbon(); }
      if (e.key === 'Escape' && this.ribbonPeek && !document.querySelector('.dlg-back, .menu')) { this.ribbonPeek = false; this.applyRibbon(); }
      if (ctrl && k === 'p') { e.preventDefault(); if (this.ed.isEditing()) this.ed.commit(0, 0); this.printDlg(); }
      if (ctrl && e.shiftKey && k === 'l' && !this.ed.isEditing()) { e.preventDefault(); this.toggleFilter(); }
      if (ctrl && e.shiftKey && k === 'u') { e.preventDefault(); $('#app').classList.toggle('fx-tall'); this.grid.layout(); }
      if (e.altKey && !ctrl && k === 'q') { e.preventDefault(); this.focusTell(); }
      if (e.shiftKey && e.key === 'F11') { e.preventDefault(); this.addSheet(); }
      if (ctrl && (e.key === 'PageDown' || e.key === 'PageUp') && !this.ed.isEditing()) { e.preventDefault(); this.stepSheet(e.key === 'PageDown' ? 1 : -1); }
    }, true);
  }
  private dirtySave = false;
  /** Ctrl+S: đang mở tệp .xlsx có quyền ghi → lưu đè tệp; không thì lưu trong app */
  save() { if (this.fileHandle && /\.xlsx$/i.test(this.fileName)) this.saveXlsx(false); else this.saveNow(true); }

  // ---------- thanh tiêu đề: Tự động lưu, QAT, tên tệp ----------
  private qatAll(): { q: string; icon: string; label: string; act: () => void; kbd?: string; def?: boolean }[] {
    return [
      { q: 'new', icon: 'newFile', label: 'Mới', kbd: 'Ctrl+N', act: () => this.newBook() },
      { q: 'open', icon: 'folder', label: 'Mở', kbd: 'Ctrl+O', act: () => this.openFile() },
      { q: 'save', icon: 'save', label: 'Lưu', kbd: 'Ctrl+S', act: () => this.save(), def: true },
      { q: 'qprint', icon: 'printer', label: 'In Nhanh', act: () => this.quickPrint() },
      { q: 'preview', icon: 'pageView', label: 'Xem trước khi In và In', kbd: 'Ctrl+P', act: () => this.printDlg() },
      { q: 'undo', icon: 'undo', label: 'Hoàn tác', kbd: 'Ctrl+Z', act: () => this.undo(), def: true },
      { q: 'redo', icon: 'redo', label: 'Làm lại', kbd: 'Ctrl+Y', act: () => this.redo(), def: true },
      { q: 'sortAsc', icon: 'sortAZ', label: 'Sắp xếp Tăng dần', act: () => this.quickSort(false) },
      { q: 'sortDesc', icon: 'sortZA', label: 'Sắp xếp Giảm dần', act: () => this.quickSort(true) },
    ];
  }
  private qatOn(): string[] {
    try { const v = JSON.parse(lsGet('xl-qat') ?? 'null'); if (Array.isArray(v)) return v; } catch { /* mặc định */ }
    return this.qatAll().filter((x) => x.def).map((x) => x.q);
  }
  private renderQat() {
    const box = $('#qat');
    box.innerHTML = '';
    const on = this.qatOn();
    for (const x of this.qatAll().filter((y) => on.includes(y.q))) {
      const b = document.createElement('button');
      b.dataset.q = x.q;
      b.innerHTML = ic(x.icon, 16);
      b.title = x.label + (x.kbd ? ` (${x.kbd})` : '');
      b.setAttribute('aria-label', b.title);
      b.onmousedown = (e) => e.preventDefault();
      b.onclick = () => { if (this.ed.isEditing()) this.ed.commit(0, 0); x.act(); };
      box.appendChild(b);
    }
    const more = document.createElement('button');
    more.id = 'qat-more';
    more.className = 'qmore';
    more.innerHTML = ic('chevDown', 12);
    more.title = 'Tùy chỉnh Thanh công cụ Truy nhập Nhanh';
    more.setAttribute('aria-label', more.title);
    more.onclick = () => {
      const r = more.getBoundingClientRect();
      const cur = this.qatOn();
      showMenu([
        ...this.qatAll().map((x) => ({ label: x.label, check: cur.includes(x.q), act: () => { const n = cur.includes(x.q) ? cur.filter((q) => q !== x.q) : this.qatAll().map((y) => y.q).filter((q) => q === x.q || cur.includes(q)); lsSet('xl-qat', JSON.stringify(n)); this.renderQat(); this.syncSelection(); } })),
        'sep',
        { label: 'Đặt lại mặc định', act: () => { lsSet('xl-qat', 'null'); this.renderQat(); this.syncSelection(); } },
      ], r.left, r.bottom + 4);
    };
    box.appendChild(more);
  }
  private bindTop() {
    this.renderQat();
    const as = $('#autosave');
    as.onclick = () => this.toggleAutoSave();
    this.updateAutoSave();
    $('#docname').onclick = () => this.docPopover();
    this.updateTitle();
  }
  updateTitle() {
    $('#doc-title').textContent = this.wb.name;
    document.title = this.wb.name + ' - ExcelLike';
  }
  private docPopover() {
    const box = h('div', { class: 'doc-pop' });
    const inp = h('input', { id: 'doc-pop-name', 'aria-label': 'Tên tệp', spellcheck: 'false' }) as HTMLInputElement;
    inp.value = this.wb.name;
    const where = this.fileName ? (desk() ? this.fileName : `${this.fileName} (tệp trên máy)`) : 'Trong ứng dụng (lưu tự động trên máy này)';
    box.append(field('Tên tệp', inp), h('p', { class: 'muted' }, 'Vị trí: ' + where), h('p', { class: 'muted' }, $('#save-state').textContent ?? ''));
    const apply = () => { const v = inp.value.trim(); if (v && v !== this.wb.name) { this.wb.name = v; this.scheduleSave(); this.updateTitle(); } };
    inp.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); apply(); closeMenu(); this.grid.focusKeys(); } };
    inp.onchange = apply;
    const r = $('#docname').getBoundingClientRect();
    openPopover(box, r.left, r.bottom + 6);
    inp.focus(); inp.select();
  }
  // Tự động lưu: ghi thẳng vào tệp .xlsx đang mở (cần tệp có quyền ghi: bản Windows hoặc Chrome/Edge)
  // bật/tắt riêng cho từng sổ (như Excel), không lan sang tệp mở sau
  private get autoKey() { return 'xl-autosave:' + this.bookId; }
  private get autoSave() { return lsGet(this.autoKey) === '1' && !!this.fileHandle && /\.xlsx$/i.test(this.fileName); }
  private updateAutoSave() {
    const on = this.autoSave, as = $('#autosave');
    as.setAttribute('aria-checked', String(on));
    as.querySelector('.as-st')!.textContent = on ? 'Bật' : 'Tắt';
    as.title = on ? `Tự động lưu vào “${this.fileName}”` : 'Tự động lưu: bật để mọi thay đổi được ghi ngay vào tệp .xlsx đang mở';
  }
  private async toggleAutoSave() {
    if (this.autoSave) { lsSet(this.autoKey, '0'); this.updateAutoSave(); return; }
    if (!(this.fileHandle && /\.xlsx$/i.test(this.fileName))) {
      this.toast('Để Tự động lưu, hãy lưu sổ thành tệp .xlsx trên máy trước.');
      if (!(await this.saveXlsx(true)) || !this.fileHandle) { this.updateAutoSave(); return; }
    }
    lsSet(this.autoKey, '1');
    this.updateAutoSave();
    if (this.fileDirty) this.saveXlsx(false);
  }
  private autoTimer = 0;
  private scheduleAuto() {
    if (!this.autoSave) return;
    clearTimeout(this.autoTimer);
    this.autoTimer = window.setTimeout(() => this.runAuto(), 2500);
  }
  /** Tự động lưu: chỉ ghi khi đã có quyền ghi tệp; không bao giờ mở hộp thoại hay tải xuống */
  private async runAuto() {
    if (!this.autoSave || !this.fileDirty || this.leaving) return;
    if (this.saving) { this.saving.finally(() => this.scheduleAuto()); return; }
    if (!desk()) {
      const h = this.fileHandle;
      let ok = false;
      try { ok = typeof h?.queryPermission === 'function' && (await h.queryPermission({ mode: 'readwrite' })) === 'granted'; } catch { ok = false; }
      if (!ok) { lsSet(this.autoKey, '0'); this.updateAutoSave(); this.toast('Đã tắt Tự động lưu: trình duyệt chưa cho ghi vào tệp. Bấm Lưu (Ctrl+S) để cấp quyền rồi bật lại.', 'err'); return; }
    }
    await this.saveXlsx(false, true);
    if (this.fileDirty) this.scheduleAuto();
  }
  recalc() { this.wb.recalc(); this.grid.invalidate(); this.toast('Đã tính lại toàn bộ sổ tính. Lịch sử hoàn tác được làm mới.'); }
  undo() { if (!this.wb.undo()) this.toast('Không còn thao tác để hoàn tác.'); this.afterHistory(); }
  redo() { if (!this.wb.redo()) this.toast('Không còn thao tác để làm lại.'); this.afterHistory(); }
  private afterHistory() {
    if (!this.wb.visibleSheetIds().includes(this.grid.sheet)) this.grid.setSheet(this.wb.visibleSheetIds()[0]);
    this.renderSheets(); this.grid.invalidate(); this.syncSelection();
  }

  // ---------- Tìm kiếm lệnh (Alt+Q) ----------
  focusTell() {
    if (this.simple) $('#titlebar').classList.add('searching');
    const t = $<HTMLInputElement>('#tell');
    t.focus(); t.select();
    this.tellShow();
  }
  private tellSel = 0;
  private tellHits: { it: Btn | Item; path: string; run: () => void; later?: string; label: string; icon: string }[] = [];
  private bindTell() {
    const t = $<HTMLInputElement>('#tell');
    $('#tell-btn').onclick = () => this.focusTell();
    t.onfocus = () => this.tellShow();
    t.oninput = () => { this.tellSel = 0; this.tellShow(); };
    t.onblur = () => setTimeout(() => { if (document.activeElement !== t) this.tellHide(); }, 150);
    t.onkeydown = (e) => {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); const n = this.tellHits.length; if (n) { this.tellSel = (this.tellSel + (e.key === 'ArrowDown' ? 1 : -1) + n) % n; this.tellShow(false); } }
      if (e.key === 'Enter') { e.preventDefault(); const x = this.tellHits[this.tellSel]; if (x) { this.tellHide(); t.value = ''; this.grid.focusKeys(); x.run(); } }
      if (e.key === 'Escape') { e.preventDefault(); t.value = ''; this.tellHide(); this.grid.focusKeys(); }
    };
  }
  private tellHide() { document.getElementById('tell-list')?.remove(); $('#titlebar').classList.remove('searching'); $('#tell').setAttribute('aria-expanded', 'false'); }
  private tellShow(recompute = true) {
    const q = fold($<HTMLInputElement>('#tell').value.trim());
    if (recompute) {
      const seen = new Set<string>();
      const hits: typeof this.tellHits = [];
      for (const { it, tab, grp } of allItems(this.tabs())) {
        if (it.t === 'chk' || seen.has(it.id)) continue;
        const label = it.label;
        if (q && !fold(label).includes(q) && !fold(grp.label).includes(q)) continue;
        seen.add(it.id);
        const btn = it as Btn;
        const run = it.t === 'combo'
          ? () => { this.selectTab(tab.id); requestAnimationFrame(() => (document.querySelector(`#ribbon [data-id="${it.id}"] input`) as HTMLInputElement | null)?.focus()); }
          : () => {
            if (btn.later) { this.later(btn.label, btn.later); return; }
            if (this.ed.isEditing()) this.ed.commit(0, 0);
            if (btn.act) { btn.act(); this.grid.invalidate(); this.syncSelection(); }
            else if (btn.drop) { const r = $('#search').getBoundingClientRect(); btn.drop(new DOMRect(r.left, r.top, r.width, r.height)); }
          };
        hits.push({ it, path: `${tab.label} ▸ ${grp.label}`, run, later: btn.later, label, icon: (it as Btn).icon ?? 'font' });
      }
      // ưu tiên tên bắt đầu bằng từ khóa
      if (q) hits.sort((a, b) => Number(!fold(a.label).startsWith(q)) - Number(!fold(b.label).startsWith(q)));
      this.tellHits = hits.slice(0, q ? 12 : 8);
    }
    let list = document.getElementById('tell-list');
    if (!list) { list = document.createElement('div'); list.id = 'tell-list'; list.className = 'tell-list'; list.setAttribute('role', 'listbox'); document.body.appendChild(list); }
    const r = $('#search').getBoundingClientRect();
    list.style.left = Math.max(4, Math.min(r.left, innerWidth - 470)) + 'px';
    list.style.top = r.bottom + 4 + 'px';
    list.innerHTML = `<h4>${q ? 'Lệnh' : 'Gợi ý'}</h4>`;
    this.tellHits.forEach((x, i) => {
      const b = document.createElement('button');
      b.setAttribute('role', 'option');
      b.setAttribute('aria-selected', String(i === this.tellSel));
      if (x.later) b.classList.add('later');
      b.innerHTML = ic(x.icon, 16) + `<span></span><small></small>`;
      b.querySelector('span')!.textContent = x.label;
      b.querySelector('small')!.textContent = x.later ? `có ở bản ${x.later}` : x.path;
      b.onmousedown = (e) => e.preventDefault();
      b.onclick = () => { this.tellHide(); $<HTMLInputElement>('#tell').value = ''; this.grid.focusKeys(); x.run(); };
      list!.appendChild(b);
    });
    if (!this.tellHits.length) list.insertAdjacentHTML('beforeend', '<p class="muted" style="padding:6px 10px">Không tìm thấy lệnh phù hợp.</p>');
    $('#tell').setAttribute('aria-expanded', 'true');
  }

  // ---------- ribbon ----------
  private tab = 'home';
  private ribbonHost() {
    const app = this;
    return {
      tabs: () => app.tabs(),
      get current() { return app.tab; },
      get simple() { return app.simple; },
      before: () => { if (app.ed.isEditing()) app.ed.commit(0, 0); },
      after: (keep: boolean) => { app.grid.invalidate(); app.syncSelection(); if (!keep && !menuOpen() && !document.querySelector('.dlg-back') && !app.backstage.isOpen()) app.grid.focusKeys(); if (app.ribbonPeek && !menuOpen()) { app.ribbonPeek = false; app.applyRibbon(); } },
      later: (label: string, ver: string) => app.later(label, ver),
      selectTab: (id: string) => app.selectTab(id),
      openFile: () => app.openFile(),
    };
  }
  later(label: string, ver: string) { this.toast(`“${label}” sẽ có ở bản ${ver}.`); }
  selectTab(id: string) {
    if (id === 'file') { this.backstage.open('home'); return; }
    this.tab = id;
    this.renderTabs();
    this.ribbon.render();
    this.applyRibbon();
  }
  renderRibbon() { this.ribbon?.render(); }
  private refreshToggles() { this.ribbon?.refresh(); }
  private ribbonPeek = false;
  private get ribbonPinned() { return lsGet('xl-ribbon') !== '0'; }
  private set ribbonPinned(v: boolean) { lsSet('xl-ribbon', v ? '1' : '0'); }
  private applyRibbon() {
    const shown = this.simple || this.ribbonPinned || this.ribbonPeek;
    $('#app').classList.toggle('rb-collapsed', !shown);
    $('#app').classList.toggle('rb-peek', !this.simple && !this.ribbonPinned && this.ribbonPeek);
    const t = $('#rb-toggle');
    if (t) { t.innerHTML = ic(this.ribbonPinned ? 'chevUp' : 'chevDown', 16); t.title = this.ribbonPinned ? 'Thu gọn Ribbon (Ctrl+F1)' : 'Ghim Ribbon (Ctrl+F1)'; t.setAttribute('aria-label', t.title); t.setAttribute('aria-pressed', String(this.ribbonPinned)); }
    this.grid?.layout();
  }
  toggleRibbon() { this.ribbonPinned = !this.ribbonPinned; this.ribbonPeek = false; this.applyRibbon(); }
  private renderTabs() {
    const nav = $('#tabs');
    nav.innerHTML = '';
    for (const t of this.tabs()) {
      const b = document.createElement('button');
      b.textContent = t.label;
      b.setAttribute('role', 'tab');
      b.dataset.tab = t.id;
      if (t.id === 'file') b.classList.add('file');
      b.setAttribute('aria-selected', String(t.id === this.tab));
      b.onmousedown = (e) => e.preventDefault();
      b.onclick = () => {
        if (t.id === 'file') { this.backstage.open('home'); return; }
        if (!this.ribbonPinned) this.ribbonPeek = !(this.ribbonPeek && t.id === this.tab);
        this.tab = t.id; this.renderTabs(); this.ribbon.render(); this.applyRibbon();
      };
      b.ondblclick = () => { if (t.id !== 'file') this.toggleRibbon(); };
      nav.appendChild(b);
    }
    nav.insertAdjacentHTML('beforeend', '<span class="tsp"></span>');
    const cm = document.createElement('button');
    cm.className = 'pill';
    cm.innerHTML = ic('comment', 15) + '<span>Ghi chú</span>';
    cm.title = 'Ghi chú — có ở bản 0.7';
    cm.onclick = () => this.later('Ghi chú', '0.7');
    nav.appendChild(cm);
    const t = document.createElement('button');
    t.id = 'rb-toggle';
    t.className = 'rb-toggle';
    t.onclick = () => this.toggleRibbon();
    nav.appendChild(t);
    this.applyRibbon();
  }
  private ktTop(): KtTarget[] {
    const tabs = this.tabs().map((t) => ({ kt: t.kt, level: t.id === 'file' ? 0 : 1, el: () => document.querySelector<HTMLElement>(`#tabs [data-tab="${t.id}"]`), run: () => (t.id === 'file' ? this.backstage.open('home') : this.selectTab(t.id)) }));
    const qat = [...document.querySelectorAll<HTMLElement>('#qat [data-q]')].map((b, i) => ({ kt: String(i + 1), el: () => b, run: () => b.click() }));
    return [...tabs, ...qat, { kt: 'Q', el: () => $('#search'), run: () => this.focusTell() }];
  }
  private ktTab(): KtTarget[] {
    const tab = this.tabs().find((t) => t.id === this.tab);
    if (!tab) return [];
    const out: KtTarget[] = [];
    const sel = (id: string) => () => document.querySelector<HTMLElement>(`#ribbon [data-id="${id}"]`);
    for (const g of tab.groups) {
      for (const { it } of allItems([{ ...tab, groups: [g] }])) {
        if (!it.kt) continue;
        out.push({
          kt: it.kt, el: sel(it.id),
          run: () => {
            const el = sel(it.id)();
            if (it.t === 'combo') { (el?.querySelector('input') as HTMLInputElement | null)?.focus(); return; }
            if (it.t === 'chk') { const i = el?.querySelector('input') as HTMLInputElement | null; if (i) { i.checked = !i.checked; i.dispatchEvent(new Event('change')); } return; }
            if (!el) { if (it.act) it.act(); return; }
            // nút tách đôi/thả xuống: KeyTip mở danh sách như Excel
            const dd = el.querySelector<HTMLElement>('.dd');
            if (dd) dd.click(); else el.click();
          },
        });
      }
      if (g.launch?.kt) out.push({ kt: g.launch.kt, el: () => document.querySelector<HTMLElement>(`#ribbon [data-launch="${g.launch!.id}"]`), run: () => document.querySelector<HTMLElement>(`#ribbon [data-launch="${g.launch!.id}"]`)?.click() });
    }
    return out;
  }

  // ---------- nội dung 9 thẻ ribbon (tên nhóm/lệnh theo Excel 365 tiếng Việt) ----------
  private anchor: { x: number; y: number } | null = null;
  /** mở menu tại vị trí nút được bấm */
  private dd(fn: () => void) { return (r: DOMRect) => { this.anchor = { x: r.left, y: r.bottom + 2 }; fn(); }; }
  private menuAt(items: () => MenuItem[]) { return (r: DOMRect) => showMenu(items(), r.left, r.bottom + 2); }
  lastFill = '#FFFF00';
  lastFont = '#FF0000';
  lastBorder: [BorderKind, BorderStyle] = ['bottom', 'thin'];
  tabs(): Tab[] {
    const s = () => this.grid.sheet;
    const cur = () => this.wb.style(s(), this.grid.sel.ar, this.grid.sel.ac);
    const st = (k: keyof Style) => () => !!cur()[k];
    const al = (a: string) => () => cur().al === a;
    const va = (a: string | undefined) => () => cur().va === a;
    const B = (t: Btn['t'], id: string, icon: string, label: string, o: Partial<Btn> = {}): Btn => ({ t, id, icon, label, ...o });
    const later = (v: string) => ({ later: v });
    const col = (...items: Item[]): Item => ({ t: 'col', items });
    const row = (...items: Item[]): Item => ({ t: 'row', items });
    const SHRINK: Record<string, number> = { styles: 1, cells: 2, edit: 3, number: 4, align: 5, clip: 6 };
    const G = (id: string, label: string, icon: string, items: Item[], launch?: Group['launch']): Group => ({ id, label, icon, items, launch, shrink: SHRINK[id] });
    const ps = () => normPrint(this.wb.meta(s()).print);
    const setPs = (p: Partial<ReturnType<typeof normPrint>>) => this.safe(() => { this.wb.setPrint(s(), storePrint({ ...ps(), ...p })); });
    const pastePick = (label: string) => () => this.later(label, '0.6');
    const home: Tab = { id: 'home', label: 'Trang đầu', kt: 'H', groups: [
      G('clip', 'Bảng tạm', 'paste', [
        B('big', 'paste', 'paste', 'Dán', { kt: 'V', kbd: 'Ctrl+V', split: true, act: () => this.pasteBtn(), drop: this.menuAt(() => [
          { label: 'Dán', icon: 'paste', hint: 'Ctrl+V', act: () => this.pasteBtn() },
          { label: 'Dán Giá trị', icon: 'pasteValues', act: pastePick('Dán Giá trị') },
          { label: 'Dán Công thức', icon: 'pasteFormula', act: pastePick('Dán Công thức') },
          { label: 'Chuyển vị', icon: 'transpose', act: pastePick('Chuyển vị') },
          { label: 'Dán Định dạng', icon: 'pasteFmt', act: pastePick('Dán Định dạng') },
          'sep',
          { label: 'Dán Đặc biệt…', hint: 'Ctrl+Alt+V', act: pastePick('Dán Đặc biệt') },
        ]) }),
        col(
          B('sm', 'cut', 'cut', 'Cắt', { kt: 'X', kbd: 'Ctrl+X', act: () => this.copySel(true) }),
          B('sm', 'copy', 'copy', 'Sao chép', { kt: 'C', kbd: 'Ctrl+C', split: true, act: () => this.copySel(false), drop: this.menuAt(() => [{ label: 'Sao chép', icon: 'copy', act: () => this.copySel(false) }, { label: 'Sao chép dưới dạng Ảnh…', icon: 'copyPic', act: () => this.later('Sao chép dưới dạng Ảnh', '0.9') }]) }),
          B('sm', 'painter', 'brush', 'Chổi định dạng', { kt: 'FP', ...later('0.6') }),
        ),
      ], { id: 'clipboard', tip: 'Bảng tạm Office', act: () => {}, kt: 'FO', later: '0.6' }),
      G('font', 'Phông chữ', 'font', [
        col(
          row(
            { t: 'combo', id: 'fn', label: 'Phông', w: 128, kt: 'FF', icon: 'font', get: () => cur().fn ?? DEFAULT_FONT, set: (v) => { if (v) this.applyStyle({ fn: v === DEFAULT_FONT ? undefined : v }); }, list: (r) => this.fontMenu(r) },
            { t: 'combo', id: 'fs', label: 'Cỡ Phông', w: 46, kt: 'FS', icon: 'fontSize', get: () => String(cur().fs ?? 11), set: (v) => { const n = parseFloat(v.replace(',', '.')); if (n >= 1 && n <= 409) { this.applyStyle({ fs: n === 11 ? undefined : n }); this.fitRows(); } else this.toast('Cỡ phông phải từ 1 đến 409.', 'err'); }, list: (r) => this.sizeMenu(r) },
            B('ic', 'grow', 'grow', 'Tăng Cỡ Phông', { kt: 'FG', kbd: 'Ctrl+Shift+>', act: () => this.growFont(1) }),
            B('ic', 'shrink', 'shrink', 'Giảm Cỡ Phông', { kt: 'FK', kbd: 'Ctrl+Shift+<', act: () => this.growFont(-1) }),
          ),
          row(
            B('ic', 'b', 'bold', 'Đậm', { kt: '1', kbd: 'Ctrl+B', act: () => this.toggle('b'), on: st('b') }),
            B('ic', 'i', 'italic', 'Nghiêng', { kt: '2', kbd: 'Ctrl+I', act: () => this.toggle('i'), on: st('i') }),
            B('ic', 'u', 'underline', 'Gạch chân', { kt: '3', kbd: 'Ctrl+U', act: () => this.toggle('u'), on: st('u'), split: true, drop: this.menuAt(() => [{ label: 'Gạch chân', icon: 'underline', act: () => this.toggle('u') }, { label: 'Gạch chân Kép', icon: 'dblUnder', act: () => this.later('Gạch chân Kép', '0.6') }]) }),
            B('ic', 'border', 'borderBottom', 'Viền', { kt: 'B', split: true, act: () => this.safe(() => this.wb.setBorder(s(), this.grid.range, ...this.lastBorder)), drop: this.dd(() => this.borderMenu()) }),
            B('ic', 'bg', 'fill', 'Màu Tô', { kt: 'H', split: true, bar: () => this.lastFill, act: () => this.applyStyle({ bg: this.lastFill }), drop: (r) => this.colorMenu('bg', r) }),
            B('ic', 'fg', 'textColor', 'Màu Phông', { kt: 'FC', split: true, bar: () => this.lastFont, act: () => this.applyStyle({ fg: this.lastFont }), drop: (r) => this.colorMenu('fg', r) }),
          ),
        ),
      ], { id: 'font', tip: 'Định dạng Ô: Phông chữ (Ctrl+Shift+F)', act: () => {}, kt: 'FN', later: '0.6' }),
      G('align', 'Căn chỉnh', 'center', [
        col(
          row(
            B('ic', 'va-t', 'aTop', 'Căn Trên', { kt: 'AT', act: () => this.applyStyle({ va: cur().va === 't' ? undefined : 't' }), on: va('t') }),
            B('ic', 'va-m', 'aMid', 'Căn Giữa', { kt: 'AM', act: () => this.applyStyle({ va: cur().va === 'm' ? undefined : 'm' }), on: va('m') }),
            B('ic', 'va-b', 'aBot', 'Căn Dưới', { kt: 'AB', act: () => this.applyStyle({ va: undefined }), on: () => !cur().va || cur().va === 'b' }),
            B('ic', 'orient', 'orient', 'Hướng', { kt: 'FQ', ...later('0.6') }),
            B('sm', 'wr', 'wrap', 'Ngắt dòng', { kt: 'W', act: () => this.toggleWrap(), on: st('wr') }),
          ),
          row(
            B('ic', 'al-l', 'left', 'Căn Trái', { kt: 'AL', act: () => this.align('l'), on: al('l') }),
            B('ic', 'al-c', 'center', 'Căn giữa', { kt: 'AC', kbd: 'Ctrl+E', act: () => this.align('c'), on: al('c') }),
            B('ic', 'al-r', 'right', 'Căn Phải', { kt: 'AR', act: () => this.align('r'), on: al('r') }),
            B('ic', 'ind-d', 'indL', 'Giảm Thụt lề', { kt: '5', act: () => this.indent(-1) }),
            B('ic', 'ind-i', 'indR', 'Tăng Thụt lề', { kt: '6', act: () => this.indent(1) }),
            B('sm', 'merge', 'merge2', 'Gộp & Căn giữa', { kt: 'M', split: true, act: () => this.doMerge(true), on: () => !!this.wb.mergeAt(s(), this.grid.sel.ar, this.grid.sel.ac), drop: this.dd(() => this.mergeMenu()) }),
          ),
        ),
      ], { id: 'align', tip: 'Định dạng Ô: Căn chỉnh', act: () => {}, kt: 'FA', later: '0.6' }),
      G('number', 'Số', 'percent', [
        col(
          row({ t: 'combo', id: 'nf', label: 'Định dạng Số', w: 132, kt: 'N', icon: 'table', get: () => nfName(cur().fmt), set: (v) => { const x = NF.find(([n]) => fold(n) === fold(v)); if (x) this.applyStyle({ fmt: x[1] }); else if (v) this.applyStyle({ fmt: v }); }, list: (r) => this.nfMenu(r) }),
          row(
            B('ic', 'acct', 'money', 'Định dạng Số Kế toán', { kt: 'AN', split: true, act: () => this.applyStyle({ fmt: NF[3][1] }), drop: this.menuAt(() => [
              { label: '₫ Tiếng Việt', act: () => this.applyStyle({ fmt: NF[3][1] }) },
              { label: '$ Tiếng Anh (Hoa Kỳ)', act: () => this.applyStyle({ fmt: '_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* "-"??_ ;_-@_ ' }) },
              { label: '€ Euro', act: () => this.applyStyle({ fmt: '_-* #,##0.00 [$€-x-euro2]_-;-* #,##0.00 [$€-x-euro2]_-;_-* "-"?? [$€-x-euro2]_-;_-@_-' }) },
            ]) }),
            B('ic', 'pct', 'percent', 'Kiểu Phần trăm', { kt: 'P', kbd: 'Ctrl+Shift+%', act: () => this.applyStyle({ fmt: '0%' }) }),
            B('ic', 'thou', 'comma', 'Kiểu Dấu phẩy', { kt: 'K', act: () => this.applyStyle({ fmt: '#,##0' }) }),
            B('ic', 'dinc', 'decInc', 'Tăng Số thập phân', { kt: '0', act: () => this.decimals(1) }),
            B('ic', 'ddec', 'decDec', 'Giảm Số thập phân', { kt: '9', act: () => this.decimals(-1) }),
          ),
        ),
      ], { id: 'number', tip: 'Định dạng Ô: Số', act: () => {}, kt: 'FM', later: '0.6' }),
      G('styles', 'Kiểu', 'styles', [
        B('big', 'cf', 'cf', 'Định dạng có Điều kiện', { kt: 'L', drop: this.dd(() => this.cfMenu()) }),
        B('big', 'fmtTable', 'fmtTable', 'Định dạng như Bảng', { kt: 'T', ...later('0.7') }),
        B('big', 'cstyle', 'styles', 'Kiểu Ô', { kt: 'J', drop: this.menuAt(() => CELL_STYLES.map(([n, sty, bk, bs, bc]) => ({ label: n, act: () => this.safe(() => { this.wb.setStyle(s(), this.grid.range, sty); if (bk) this.wb.setBorder(s(), this.grid.range, bk, bs, bc); else if (n === 'Thường') this.wb.setBorder(s(), this.grid.range, 'none'); }) }))) }),
      ]),
      G('cells', 'Ô', 'cellIns', [
        B('big', 'ins', 'cellIns', 'Chèn', { kt: 'I', drop: this.menuAt(() => this.insertItems()) }),
        B('big', 'del', 'cellDel', 'Xóa', { kt: 'D', drop: this.menuAt(() => this.deleteItems()) }),
        B('big', 'format', 'cellFmt', 'Định dạng', { kt: 'O', drop: this.menuAt(() => this.formatItems()) }),
      ]),
      G('edit', 'Chỉnh sửa', 'sigma', [
        col(
          B('sm', 'sum', 'sigma', 'Tổng Tự động', { kt: 'U', kbd: 'Alt+=', split: true, act: () => this.autoSum(), drop: this.menuAt(() => [
            ...([['Tổng', 'SUM'], ['Trung bình', 'AVERAGE'], ['Đếm Số', 'COUNT'], ['Tối đa', 'MAX'], ['Tối thiểu', 'MIN']] as const).map(([l, f]) => ({ label: l, act: () => this.autoSum(f) })),
            'sep', { label: 'Hàm Khác…', act: () => this.funcPicker() },
          ]) }),
          B('sm', 'fill', 'fillDown', 'Điền', { kt: 'FI', drop: this.menuAt(() => [
            { label: 'Xuống', hint: 'Ctrl+D', act: () => this.fillDir('down') },
            { label: 'Phải', hint: 'Ctrl+R', act: () => this.fillDir('right') },
            { label: 'Lên', act: () => this.fillDir('up') },
            { label: 'Trái', act: () => this.fillDir('left') },
            'sep',
            { label: 'Chuỗi…', act: () => this.later('Điền Chuỗi', '0.6') },
            { label: 'Điền Nhanh', hint: 'Ctrl+E', act: () => this.later('Điền Nhanh', '0.6') },
          ]) }),
          B('sm', 'clear', 'eraser', 'Xóa', { kt: 'E', drop: this.menuAt(() => [
            { label: 'Xóa Tất cả', icon: 'eraser', act: () => this.clearAll() },
            { label: 'Xóa Định dạng', act: () => this.clearFormats() },
            { label: 'Xóa Nội dung', hint: 'Delete', act: () => this.clearSel() },
            { label: 'Xóa Ghi chú và Chú thích', act: () => this.later('Ghi chú', '0.7') },
            { label: 'Xóa Siêu kết nối', act: () => this.later('Siêu kết nối', '0.9') },
          ]) }),
        ),
        B('big', 'sortf', 'sortAZ', 'Sắp xếp & Lọc', { kt: 'S', drop: this.dd(() => this.sortFilterMenu()) }),
        B('big', 'find', 'find', 'Tìm & Chọn', { kt: 'FD', ...later('0.6') }),
      ]),
    ] };
    const insert: Tab = { id: 'insert', label: 'Chèn', kt: 'N', groups: [
      G('tables', 'Bảng', 'pivot', [
        B('big', 'pivot', 'pivot', 'PivotTable', { kt: 'V', act: () => this.insertPivot() }),
        B('big', 'recPivot', 'recommended', 'PivotTable Đề xuất', { kt: 'SP', ...later('0.9') }),
        B('big', 'table', 'fmtTable', 'Bảng', { kt: 'T', kbd: 'Ctrl+T', ...later('0.7') }),
      ]),
      G('illus', 'Hình minh họa', 'picture', [
        B('big', 'pic', 'picture', 'Ảnh', { kt: 'P', ...later('0.9') }),
        col(B('sm', 'shapes', 'shapes', 'Hình dạng', { kt: 'SH', ...later('0.9') }), B('sm', 'icons', 'effects', 'Biểu tượng', { kt: 'Y1', ...later('0.9') })),
      ]),
      G('charts', 'Biểu đồ', 'chart', [
        B('big', 'chart', 'recommended', 'Biểu đồ Đề xuất', { kt: 'R', act: () => this.insertChart() }),
        col(
          row(
            B('ic', 'i-col', 'chart', 'Chèn Biểu đồ Cột hoặc Thanh', { kt: 'C', act: () => this.insertChart('col') }),
            B('ic', 'i-chart', 'chartLine', 'Chèn Biểu đồ Đường hoặc Vùng', { kt: 'N', act: () => this.insertChart('line') }),
            B('ic', 'i-pie', 'chartPie', 'Chèn Biểu đồ Tròn hoặc Vành khuyên', { kt: 'Q', act: () => this.insertChart('pie') }),
          ),
          row(
            B('ic', 'i-bar', 'chartBar', 'Chèn Biểu đồ Thanh', { kt: 'B', act: () => this.insertChart('bar') }),
            B('ic', 'i-area', 'chartArea', 'Chèn Biểu đồ Vùng', { kt: 'A', act: () => this.insertChart('area') }),
            B('ic', 'i-xy', 'shapes', 'Chèn Biểu đồ Phân tán (X, Y)', { kt: 'D', ...later('0.9') }),
          ),
        ),
        B('big', 'pchart', 'chartStack', 'PivotChart', { kt: 'ZC', ...later('0.9') }),
      ], { id: 'charts', tip: 'Xem Tất cả Biểu đồ', act: () => this.insertChart(), kt: 'K' }),
      G('spark', 'Biểu đồ thu nhỏ', 'sparkLine', [
        col(B('sm', 'sp-line', 'sparkLine', 'Đường', { kt: 'SL', ...later('0.9') }), B('sm', 'sp-col', 'sparkCol', 'Cột', { kt: 'SO', ...later('0.9') }), B('sm', 'sp-wl', 'winLoss', 'Thắng/Thua', { kt: 'SW', ...later('0.9') })),
      ]),
      G('filters', 'Bộ lọc', 'slicer', [B('big', 'slicer', 'slicer', 'Slicer', { kt: 'SF', ...later('0.9') }), B('big', 'tl', 'timeline', 'Dòng thời gian', { kt: 'LT', ...later('0.9') })]),
      G('links', 'Nối kết', 'link', [B('big', 'link', 'link', 'Liên kết', { kt: 'I', kbd: 'Ctrl+K', ...later('0.9') })]),
      G('comments', 'Chú thích', 'comment', [B('big', 'i-comment', 'comment', 'Chú thích', { kt: 'M', ...later('0.7') })]),
      G('text', 'Văn bản', 'textbox', [
        B('big', 'textbox', 'textbox', 'Hộp Văn bản', { kt: 'X', ...later('0.9') }),
        B('big', 'hf', 'hf', 'Đầu trang & Chân trang', { kt: 'H', act: () => this.pageSetup('hf') }),
        col(B('sm', 'wordart', 'wordart', 'WordArt', { kt: 'W', ...later('0.9') }), B('sm', 'sigline', 'signature', 'Dòng Chữ ký', { kt: 'G', ...later('0.9') }), B('sm', 'object', 'object', 'Đối tượng', { kt: 'J', ...later('0.9') })),
      ]),
      G('symbols', 'Ký hiệu', 'symbol', [B('big', 'equation', 'equation', 'Phương trình', { kt: 'E', ...later('0.9') }), B('big', 'symbol', 'symbol', 'Ký hiệu', { kt: 'U', ...later('0.9') })]),
    ] };
    const margins = (): MenuItem[] => ([['normal', 'Thường', 'Trên: 1,91 cm  Dưới: 1,91 cm  Trái: 1,78 cm  Phải: 1,78 cm'], ['wide', 'Rộng', 'Trên: 2,54 cm  Dưới: 2,54 cm  Trái: 2,54 cm  Phải: 2,54 cm'], ['narrow', 'Hẹp', 'Trên: 1,91 cm  Dưới: 1,91 cm  Trái: 0,64 cm  Phải: 0,64 cm']] as const)
      .map(([k, l, d]) => ({ label: `${l} — ${d}`, check: ps().margins === k, act: () => setPs({ margins: k, mg: normPrint({ ...ps(), margins: k, mg: undefined } as any).mg }) }) as MenuItem)
      .concat(['sep', { label: 'Lề Tùy chỉnh…', act: () => this.pageSetup('margins') }]);
    const layout: Tab = { id: 'layout', label: 'Bố trí Trang', kt: 'P', groups: [
      G('themes', 'Chủ đề', 'themes', [
        B('big', 'themes', 'themes', 'Chủ đề', { kt: 'TH', ...later('0.9') }),
        col(B('sm', 'th-colors', 'palette', 'Màu', { kt: 'TC', ...later('0.9') }), B('sm', 'th-fonts', 'aa', 'Phông', { kt: 'TF', ...later('0.9') }), B('sm', 'th-eff', 'effects', 'Hiệu ứng', { kt: 'TE', ...later('0.9') })),
      ]),
      G('setup', 'Thiết lập Trang', 'margins', [
        B('big', 'pl-margins', 'margins', 'Lề', { kt: 'M', drop: this.menuAt(margins) }),
        B('big', 'pl-orient', 'orientation', 'Hướng', { kt: 'O', drop: this.menuAt(() => [{ label: 'Dọc', check: ps().orientation === 'portrait', act: () => setPs({ orientation: 'portrait' }) }, { label: 'Ngang', check: ps().orientation === 'landscape', act: () => setPs({ orientation: 'landscape' }) }]) }),
        B('big', 'pl-size', 'paperSize', 'Kích thước', { kt: 'SZ', drop: this.menuAt(() => [...(['A4', 'A3', 'A5', 'B4', 'B5', 'Letter', 'Legal', 'Tabloid'] as const).map((p) => ({ label: p, check: ps().paper === p, act: () => setPs({ paper: p }) }) as MenuItem), 'sep', { label: 'Thêm Kích thước Giấy…', act: () => this.pageSetup('page') }]) }),
        B('big', 'pl-area', 'printArea', 'Vùng In', { kt: 'R', drop: this.menuAt(() => [
          { label: 'Đặt Vùng In', act: () => setPs({ area: this.grid.range.r2 >= MAX_ROWS - 1 || this.grid.range.c2 >= MAX_COLS - 1 ? this.wb.expandToMerges(s(), { ...this.grid.range, r2: Math.min(this.grid.range.r2, Math.max(0, this.wb.used(s()).rows - 1)), c2: Math.min(this.grid.range.c2, Math.max(0, this.wb.used(s()).cols - 1)) }) : this.grid.range }) },
          { label: 'Xóa Vùng In', act: () => setPs({ area: null }) },
        ]) }),
        B('big', 'pl-breaks', 'breaks', 'Ngắt', { kt: 'B', drop: this.menuAt(() => [
          { label: 'Chèn Ngắt Trang', act: () => this.safe(() => this.wb.insertPageBreak(s(), this.grid.sel.ar, this.grid.sel.ac)) },
          { label: 'Loại bỏ Ngắt Trang', act: () => this.safe(() => this.wb.removePageBreak(s(), this.grid.sel.ar, this.grid.sel.ac)) },
          { label: 'Đặt lại Tất cả Ngắt Trang', act: () => this.safe(() => this.wb.resetPageBreaks(s())) },
        ]) }),
        B('big', 'pl-bg', 'background', 'Nền', { kt: 'G', ...later('0.9') }),
        B('big', 'pl-titles', 'printTitles', 'In Tiêu đề', { kt: 'I', act: () => this.pageSetup('sheet') }),
      ], { id: 'pagesetup', tip: 'Thiết lập Trang', act: () => this.pageSetup('page'), kt: 'SP' }),
      G('fit', 'Co giãn cho Vừa', 'fitPage', [
        col(
          { t: 'combo', id: 'pl-w', label: 'Chiều rộng', w: 110, kt: 'W', icon: 'colW', get: () => (ps().fit?.w ? `${ps().fit!.w} trang` : 'Tự động'), set: (v) => { const n = parseInt(v) || 0; setPs({ fit: n || ps().fit?.h ? { w: n, h: ps().fit?.h ?? 0 } : null }); }, list: (r) => showMenu(['Tự động', '1 trang', '2 trang', '3 trang'].map((t, i) => ({ label: t, act: () => setPs({ fit: i || ps().fit?.h ? { w: i, h: ps().fit?.h ?? 0 } : null }) })), r.left, r.bottom + 2) },
          { t: 'combo', id: 'pl-h', label: 'Chiều cao', w: 110, kt: 'H', icon: 'rowH', get: () => (ps().fit?.h ? `${ps().fit!.h} trang` : 'Tự động'), set: (v) => { const n = parseInt(v) || 0; setPs({ fit: n || ps().fit?.w ? { w: ps().fit?.w ?? 0, h: n } : null }); }, list: (r) => showMenu(['Tự động', '1 trang', '2 trang', '3 trang'].map((t, i) => ({ label: t, act: () => setPs({ fit: i || ps().fit?.w ? { w: ps().fit?.w ?? 0, h: i } : null }) })), r.left, r.bottom + 2) },
          { t: 'combo', id: 'pl-scale', label: 'Tỷ lệ', w: 110, kt: 'SC', icon: 'zoomIn', get: () => (ps().fit ? '—' : ps().scale + '%'), set: (v) => { const n = parseInt(v); if (n >= 10 && n <= 400) setPs({ scale: n, fit: null }); else this.toast('Tỷ lệ phải từ 10% đến 400%.', 'err'); }, list: (r) => showMenu([50, 75, 85, 100, 125, 150].map((n) => ({ label: n + '%', act: () => setPs({ scale: n, fit: null }) })), r.left, r.bottom + 2) },
        ),
      ], { id: 'scale', tip: 'Thiết lập Trang: Trang', act: () => this.pageSetup('page'), kt: 'SF' }),
      G('sheetopt', 'Tùy chọn Trang tính', 'grid', [
        col({ t: 'chk', id: 'gl-view', label: 'Xem đường lưới', kt: 'VG', get: () => this.grid.gridlines, set: (v) => { this.grid.gridlines = v; this.grid.invalidate(); } }, { t: 'chk', id: 'gl-print', label: 'In đường lưới', kt: 'PG', get: () => ps().gridlines, set: (v) => setPs({ gridlines: v }) }),
        col({ t: 'chk', id: 'hd-view', label: 'Xem đầu đề', kt: 'VH', get: () => this.grid.headings, set: (v) => { this.grid.headings = v; this.grid.layout(); } }, { t: 'chk', id: 'hd-print', label: 'In đầu đề', kt: 'PH', get: () => ps().headings, set: (v) => setPs({ headings: v }) }),
      ], { id: 'sheetopts', tip: 'Thiết lập Trang: Trang tính', act: () => this.pageSetup('sheet'), kt: 'SO' }),
      G('arrange', 'Sắp xếp', 'arrange', [
        col(B('sm', 'ar-fwd', 'arrange', 'Đưa Lên trên', { kt: 'AF', ...later('0.9') }), B('sm', 'ar-back', 'arrange', 'Đưa Xuống dưới', { kt: 'AE', ...later('0.9') }), B('sm', 'ar-pane', 'slicer', 'Ngăn Chọn', { kt: 'AP', ...later('0.9') })),
      ]),
    ] };
    const formulas: Tab = { id: 'formulas', label: 'Công thức', kt: 'M', groups: [
      G('lib', 'Thư viện Hàm', 'fx', [
        B('big', 'fnp', 'fx', 'Chèn Hàm', { kt: 'F', kbd: 'Shift+F3', act: () => this.funcPicker() }),
        B('big', 'fsum', 'sigma', 'Tổng Tự động', { kt: 'U', split: true, act: () => this.autoSum(), drop: this.menuAt(() => ([['Tổng', 'SUM'], ['Trung bình', 'AVERAGE'], ['Đếm Số', 'COUNT'], ['Tối đa', 'MAX'], ['Tối thiểu', 'MIN']] as const).map(([l, f]) => ({ label: l, act: () => this.autoSum(f) }))) }),
        ...(([['f-fin', 'Tài chính', 'money', 'I', 'PMT NPV IRR FV PV RATE NPER'], ['f-log', 'Lô-gic', 'validate', 'L', 'IF IFS AND OR NOT IFERROR XOR SWITCH'], ['f-text', 'Văn bản', 'font', 'T', 'TEXT LEFT RIGHT MID LEN CONCAT TRIM UPPER LOWER SUBSTITUTE'], ['f-date', 'Ngày & Giờ', 'calcOpts', 'E', 'TODAY NOW DATE YEAR MONTH DAY EDATE EOMONTH NETWORKDAYS'], ['f-look', 'Tra cứu & Tham chiếu', 'find', 'O', 'VLOOKUP HLOOKUP XLOOKUP INDEX MATCH OFFSET INDIRECT'], ['f-math', 'Toán & Lượng giác', 'sigma', 'G', 'SUM SUMIF SUMIFS SUMPRODUCT ROUND ROUNDUP ROUNDDOWN INT ABS MOD']] as const)
          .map(([id, l, ico, k, fns]) => B('big', id, ico, l, { kt: k, drop: this.menuAt(() => fns.split(' ').map((f) => ({ label: f, act: () => this.ed.begin('=' + f + '(', 'enter') }))) }))),
      ]),
      G('names', 'Tên đã Xác định', 'nameMgr', [
        B('big', 'nm-mgr', 'nameMgr', 'Trình Quản lý Tên', { kt: 'N', kbd: 'Ctrl+F3', ...later('0.7') }),
        col(B('sm', 'nm-def', 'tag', 'Xác định Tên', { kt: 'MD', ...later('0.7') }), B('sm', 'nm-use', 'fx', 'Dùng trong Công thức', { kt: 'S', ...later('0.7') }), B('sm', 'nm-create', 'table', 'Tạo từ Vùng chọn', { kt: 'C', kbd: 'Ctrl+Shift+F3', ...later('0.7') })),
      ]),
      G('audit', 'Kiểm nghiệm Công thức', 'trace', [
        col(B('sm', 'tr-pre', 'trace', 'Truy tìm Ô Tiền lệ', { kt: 'P', ...later('0.7') }), B('sm', 'tr-dep', 'traceDep', 'Truy tìm Ô Phụ thuộc', { kt: 'D', ...later('0.7') }), B('sm', 'tr-rm', 'arrowsX', 'Loại bỏ Mũi tên', { kt: 'AA', ...later('0.7') })),
        col(B('sm', 'show-f', 'showF', 'Hiện Công thức', { kt: 'H', kbd: 'Ctrl+`', on: () => this.grid.showFormulas, act: () => { this.grid.showFormulas = !this.grid.showFormulas; this.grid.invalidate(); } }), B('sm', 'err-chk', 'errorCheck', 'Kiểm tra Lỗi', { kt: 'K', ...later('0.9') }), B('sm', 'eval-f', 'evaluate', 'Đánh giá Công thức', { kt: 'V', ...later('0.9') })),
        B('big', 'watch', 'watch', 'Cửa sổ Theo dõi', { kt: 'W', ...later('0.9') }),
      ]),
      G('calc', 'Tính toán', 'calcOpts', [
        B('big', 'calc-opt', 'calcOpts', 'Tùy chọn Tính toán', { kt: 'X', drop: this.menuAt(() => [{ label: 'Tự động', check: true, act: () => {} }, { label: 'Thủ công', act: () => this.later('Tính toán Thủ công', '0.9') }]) }),
        col(
          B('sm', 'recalc', 'calc', 'Tính Ngay', { kt: 'B', kbd: 'F9', act: () => this.recalc() }),
          B('sm', 'recalc-s', 'calcSheet', 'Tính Trang tính', { kt: 'J', kbd: 'Shift+F9', act: () => this.recalc() }),
          B('sm', 'iter', 'refresh', 'Tính lặp', { act: () => this.setIter(!this.wb.iterate.on), on: () => this.wb.iterate.on }),
        ),
      ]),
    ] };
    const data: Tab = { id: 'data', label: 'Dữ liệu', kt: 'A', groups: [
      G('get', 'Lấy & Chuyển đổi Dữ liệu', 'getData', [
        B('big', 'd-get', 'getData', 'Lấy Dữ liệu', { kt: 'PN', drop: this.menuAt(() => [{ label: 'Từ Sổ làm việc Excel…', icon: 'file', act: () => this.openFile() }, { label: 'Từ Văn bản/CSV…', icon: 'fromText', act: () => this.openFile() }, { label: 'Từ Web…', icon: 'web', act: () => this.later('Lấy dữ liệu từ Web', 'sau 1.0 (ngoài phạm vi)') }]) }),
        col(B('sm', 'd-open', 'fromText', 'Từ Văn bản/CSV', { kt: 'FT', act: () => this.openFile() }), B('sm', 'd-web', 'web', 'Từ Web', { kt: 'FW', ...later('ngoài phạm vi') }), B('sm', 'd-tbl', 'fmtTable', 'Từ Bảng/Vùng', { kt: 'FB', ...later('ngoài phạm vi') })),
      ]),
      G('queries', 'Truy vấn & Kết nối', 'conn', [
        B('big', 'd-refresh', 'refresh', 'Làm mới Tất cả', { kt: 'R', kbd: 'Ctrl+Alt+F5', act: () => { for (const id of this.wb.sheetIds()) for (const p of this.wb.meta(id).pivots) this.safe(() => this.wb.refreshPivot(id, p.id!)); this.recalc(); } }),
        col(B('sm', 'd-links', 'link', 'Liên kết Sổ làm việc', { kt: 'K', act: () => this.toast(this.wb.extLinks.length ? `Sổ có ${this.wb.extLinks.length} liên kết ngoài (giữ giá trị lưu sẵn, ghi lại khi lưu .xlsx).` : 'Sổ không có liên kết ngoài.') })),
      ]),
      G('sortfilter', 'Sắp xếp & Lọc', 'filter', [
        col(B('ic', 'd-az', 'sortAZ', 'Sắp xếp từ A đến Z', { kt: 'SA', act: () => this.quickSort(false) }), B('ic', 'd-za', 'sortZA', 'Sắp xếp từ Z đến A', { kt: 'SD', act: () => this.quickSort(true) })),
        B('big', 'd-sort', 'sortAZ', 'Sắp xếp', { kt: 'SS', act: () => this.sortDlg() }),
        B('big', 'd-f', 'filter', 'Lọc', { kt: 'T', kbd: 'Ctrl+Shift+L', act: () => this.toggleFilter(), on: () => !!this.wb.meta(s()).filter }),
        col(B('sm', 'd-fx', 'filterX', 'Xóa', { kt: 'C', act: () => this.clearFilterConds() }), B('sm', 'd-fr', 'refresh', 'Áp dụng lại', { kt: 'Y', act: () => { this.wb.reapplyFilter(s()); this.grid.layout(); } }), B('sm', 'd-adv', 'filter', 'Nâng cao', { kt: 'Q', ...later('0.9') })),
      ]),
      G('tools', 'Công cụ Dữ liệu', 'textCols', [
        B('big', 'd-t2c', 'textCols', 'Văn bản thành Cột', { kt: 'E', ...later('0.7') }),
        col(B('ic', 'd-flash', 'flash', 'Điền Nhanh', { kt: 'FF', ...later('0.6') }), B('ic', 'd-dup', 'removeDup', 'Loại bỏ Trùng lặp', { kt: 'M', ...later('0.7') }), B('ic', 'd-v', 'validate', 'Xác thực Dữ liệu', { kt: 'V', act: () => this.dvDlg() })),
        col(B('ic', 'd-cf', 'cf', 'Định dạng có Điều kiện', { drop: this.dd(() => this.cfMenu()) }), B('ic', 'd-cons', 'subtotal', 'Hợp nhất', { kt: 'N', ...later('0.9') })),
      ]),
      G('forecast', 'Dự báo', 'forecast', [B('big', 'd-whatif', 'whatIf', 'Phân tích Giả định', { kt: 'W', ...later('0.9') }), B('big', 'd-fc', 'forecast', 'Trang tính Dự báo', { kt: 'FC', ...later('0.9') })]),
      G('outline', 'Đại cương', 'group', [
        B('big', 'd-group', 'group', 'Nhóm', { kt: 'G', kbd: 'Alt+Shift+→', ...later('0.7') }),
        B('big', 'd-ungroup', 'ungroup', 'Bỏ nhóm', { kt: 'U', kbd: 'Alt+Shift+←', ...later('0.7') }),
        B('big', 'd-sub', 'subtotal', 'Tổng phụ', { kt: 'B', ...later('0.7') }),
      ]),
    ] };
    const review: Tab = { id: 'review', label: 'Xem lại', kt: 'R', groups: [
      G('proof', 'Soát lỗi', 'spell', [B('big', 'r-spell', 'spell', 'Chính tả', { kt: 'S', kbd: 'F7', ...later('0.9') }), col(B('sm', 'r-thes', 'thesaurus', 'Từ điển Đồng nghĩa', { kt: 'E', ...later('ngoài phạm vi') }), B('sm', 'r-stats', 'stats', 'Thống kê Sổ làm việc', { kt: 'W', act: () => this.bookStats() }))]),
      G('access', 'Trợ năng', 'access', [B('big', 'r-access', 'access', 'Kiểm tra Trợ năng', { kt: 'A1', act: () => this.toast('Trợ năng: Mọi thứ đều ổn. Ô, thẻ sheet và lệnh đều có nhãn đọc màn hình.', 'ok') })]),
      G('notes', 'Ghi chú', 'note', [B('big', 'r-note', 'note', 'Ghi chú', { kt: 'T', kbd: 'Shift+F2', ...later('0.7') })]),
      G('comments2', 'Chú thích', 'comment', [B('big', 'r-comment', 'comment', 'Chú thích Mới', { kt: 'C', ...later('0.7') })]),
      G('protect', 'Bảo vệ', 'protectSheet', [B('big', 'r-psheet', 'protectSheet', 'Bảo vệ Trang tính', { kt: 'PS', ...later('0.9') }), B('big', 'r-pbook', 'protectBook', 'Bảo vệ Sổ làm việc', { kt: 'PW', ...later('0.9') }), B('big', 'r-allow', 'allowEdit', 'Cho phép Sửa Vùng', { kt: 'U1', ...later('0.9') })]),
    ] };
    const view: Tab = { id: 'view', label: 'Xem', kt: 'W', groups: [
      G('views', 'Dạng xem Sổ làm việc', 'normalView', [
        B('big', 'v-normal', 'normalView', 'Thường', { kt: 'L', act: () => this.setView('normal'), on: () => this.grid.view === 'normal' }),
        B('big', 'v-break', 'breakView', 'Xem trước Ngắt Trang', { kt: 'I', act: () => this.setView('break'), on: () => this.grid.view === 'break' }),
        B('big', 'v-layout', 'pageView', 'Bố trí Trang', { kt: 'P', ...later('0.8') }),
      ]),
      G('show', 'Hiện', 'eye', [
        col(
          { t: 'chk', id: 'gl', label: 'Đường lưới', kt: 'VG', get: () => this.grid.gridlines, set: (v) => { this.grid.gridlines = v; this.grid.invalidate(); } },
          { t: 'chk', id: 'v-fbar', label: 'Thanh Công thức', kt: 'VF', get: () => !$('#app').classList.contains('no-fbar'), set: (v) => { $('#app').classList.toggle('no-fbar', !v); this.grid.layout(); } },
          { t: 'chk', id: 'v-heads', label: 'Đầu đề', kt: 'VH', get: () => this.grid.headings, set: (v) => { this.grid.headings = v; this.grid.layout(); } },
        ),
      ]),
      G('zoom', 'Thu phóng', 'zoomIn', [
        B('big', 'zoomdlg', 'zoomIn', 'Thu phóng', { kt: 'Q', act: () => this.zoomDlg() }),
        B('big', 'z100', 'table', '100%', { kt: 'J', act: () => this.grid.setZoom(1) }),
        B('big', 'zsel', 'zoomSel', 'Thu phóng tới Vùng chọn', { kt: 'G', act: () => this.zoomToSel() }),
      ]),
      G('window', 'Cửa sổ', 'newWin', [
        B('big', 'w-new', 'newWin', 'Cửa sổ Mới', { kt: 'N', ...later('0.9') }),
        B('big', 'w-arr', 'arrange', 'Sắp xếp Tất cả', { kt: 'A', ...later('0.9') }),
        B('big', 'fz', 'freeze', 'Cố định Ngăn', { kt: 'F', drop: this.menuAt(() => [
          { label: this.wb.meta(s()).freeze.rows || this.wb.meta(s()).freeze.cols ? 'Bỏ Cố định Ngăn' : 'Cố định Ngăn', icon: 'freeze', act: () => (this.wb.meta(s()).freeze.rows || this.wb.meta(s()).freeze.cols ? this.freeze(0, 0) : this.freeze(this.grid.sel.ar, this.grid.sel.ac)) },
          { label: 'Cố định Hàng Trên cùng', icon: 'rowIns', act: () => this.freeze(1, 0) },
          { label: 'Cố định Cột Đầu tiên', icon: 'colIns', act: () => this.freeze(0, 1) },
        ]) }),
        col(B('sm', 'w-split', 'splitWin', 'Tách', { kt: 'S', ...later('0.9') }), B('sm', 'w-hide', 'hideWin', 'Ẩn', { kt: 'H', ...later('0.9') }), B('sm', 'w-unhide', 'eye', 'Bỏ ẩn', { kt: 'U', ...later('0.9') })),
      ]),
      G('macros', 'Macro', 'macro', [B('big', 'macros', 'macro', 'Macro', { kt: 'M', ...later('ngoài phạm vi (không chạy VBA)') })]),
    ] };
    const help: Tab = { id: 'help', label: 'Trợ giúp', kt: 'Y', groups: [
      G('help', 'Trợ giúp', 'help', [
        B('big', 'h-help', 'help', 'Trợ giúp', { kt: 'H', act: () => this.helpDlg() }),
        B('big', 'h-keys', 'kbd', 'Phím tắt', { kt: 'K', act: () => this.helpDlg(true) }),
        B('big', 'h-new', 'news', 'Có gì Mới', { kt: 'W', act: () => this.toast(`ExcelLike ${VERSION}: giao diện như Excel 365 (thanh tiêu đề, 9 thẻ, KeyTips, Backstage), in ấn có xem trước từng trang, đầu/chân trang, ngắt trang, Thiết lập Trang.`) }),
        B('big', 'about', 'info', 'Giới thiệu', { kt: 'A', act: () => this.backstage.open('account') }),
      ]),
    ] };
    return [{ id: 'file', label: 'Tệp', kt: 'F', groups: [] }, home, insert, layout, formulas, data, review, view, help];
  }

  // ---------- lệnh định dạng ----------
  private safe(f: () => void) { try { f(); } catch (e) { this.toast((e as Error).message, 'err'); } }
  applyStyle(p: Style) { this.safe(() => this.wb.setStyle(this.grid.sheet, this.grid.range, p)); this.refreshToggles(); }
  toggle(k: 'b' | 'i' | 'u') {
    const cur = !!this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac)[k];
    this.applyStyle({ [k]: !cur } as Style);
  }
  private align(a: 'l' | 'c' | 'r') {
    const cur = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).al;
    this.applyStyle({ al: cur === a ? undefined : a });
  }
  private decimals(d: number) {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel;
    let f = this.wb.effFmt(s, ar, ac) ?? '';
    if (/[dy]/i.test(f)) return;
    if (!f) {
      const v = this.wb.value(s, ar, ac);
      const n = typeof v === 'number' ? (String(v).split('.')[1] || '').length : 0;
      f = '0' + (n ? '.' + '0'.repeat(n) : '');
    }
    const m = /^([^.%"]*?)(?:\.(0+))?(%?)(".*")?$/.exec(f);
    if (!m) return;
    const n = Math.max(0, Math.min(10, (m[2]?.length ?? 0) + d));
    this.applyStyle({ fmt: m[1] + (n ? '.' + '0'.repeat(n) : '') + m[3] + (m[4] ?? '') });
  }
  private at(id: string) { const a = this.anchor; if (a) { this.anchor = null; return a; } const r = document.querySelector(`[data-id="${id}"]`)?.getBoundingClientRect(); return r ? { x: r.left, y: r.bottom + 4 } : { x: 80, y: 120 }; }
  private borderMenu() {
    const p = this.at('border'), s = this.grid.sheet;
    const b = (kind: BorderKind, style: BorderStyle = 'thin', remember = true) => () => this.safe(() => { this.wb.setBorder(s, this.grid.range, kind, style); if (remember) this.lastBorder = [kind, style]; });
    const two = (top: BorderStyle, bot: BorderStyle) => () => this.safe(() => { this.wb.setBorder(s, this.grid.range, 'top', top); this.wb.setBorder(s, this.grid.range, 'bottom', bot); });
    showMenu([
      { label: 'Viền Dưới', icon: 'borderBottom', act: b('bottom') },
      { label: 'Viền Trên', icon: 'borderTop', act: b('top') },
      { label: 'Viền Trái', icon: 'borderLeft', act: b('left') },
      { label: 'Viền Phải', icon: 'borderRight', act: b('right') },
      'sep',
      { label: 'Không Viền', icon: 'x', act: b('none') },
      { label: 'Tất cả Viền', icon: 'borders', act: b('all') },
      { label: 'Viền Ngoài', icon: 'borderOuter', act: b('outer') },
      { label: 'Viền Ngoài Dày', icon: 'borderOuter', act: b('outer', 'medium') },
      { label: 'Viền Trong', icon: 'borders', act: b('inner') },
      'sep',
      { label: 'Viền Dưới Kép', icon: 'borderBottom', act: b('bottom', 'double') },
      { label: 'Viền Dưới Dày', icon: 'borderBottom', act: b('bottom', 'medium') },
      { label: 'Viền Trên và Dưới', icon: 'borders', act: two('thin', 'thin') },
      { label: 'Viền Trên và Dưới Dày', icon: 'borders', act: two('thin', 'medium') },
      { label: 'Viền Trên và Dưới Kép', icon: 'borders', act: two('thin', 'double') },
      'sep',
      { label: 'Viền Nét đứt', icon: 'borders', act: b('all', 'dashed') },
      { label: 'Thêm Viền…', icon: 'cellFmt', act: () => this.later('Hộp Định dạng Ô ▸ Viền', '0.6') },
    ], p.x, p.y);
  }


  private toggleWrap() {
    const on = !this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).wr;
    this.applyStyle({ wr: on });
    if (on) this.fitRows();
  }
  /** Tăng chiều cao hàng cho vừa chữ (ô xuống dòng hoặc chữ to) */
  fitRows(exact = false) {
    const g = this.grid.range, s = this.grid.sheet, u = this.wb.used(s);
    const r2 = Math.min(g.r2, u.rows - 1, g.r1 + 2000);
    const def = this.wb.meta(s).rowH.def;
    for (let r = g.r1; r <= r2; r++) {
      if (this.wb.isRowHidden(s, r)) continue;
      const need = this.grid.neededHeight(r);
      const cur = this.wb.meta(s).rowH.get(r);
      if (need > cur) this.wb.setRowHeight(s, r, need);
      else if (exact && cur !== Math.max(def, need)) this.wb.setRowHeight(s, r, Math.max(def, need));
    }
    this.grid.layout();
  }
  safeSetPrint(s: number, ps: import('../core/workbook').PrintSetup) { this.safe(() => this.wb.setPrint(s, ps)); }
  async switchBook(id: string) {
    await this.saveNow(false);
    lsSet('xl-last', id);
    this.reload();
  }
  private mergeMenu() {
    const p = this.at('merge'), s = this.grid.sheet, g = this.grid.range;
    showMenu([
      { label: 'Gộp & Căn giữa', icon: 'merge2', act: () => this.doMerge(true) },
      { label: 'Gộp Ngang', icon: 'merge2', act: () => this.safe(() => { for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) this.wb.merge(s, { ...g, r1: r, r2: r }, false); }) },
      { label: 'Gộp Ô', icon: 'merge2', act: () => this.doMerge(false) },
      { label: 'Bỏ gộp Ô', icon: 'x', act: () => { this.wb.unmerge(s, g); this.grid.invalidate(); } },
    ], p.x, p.y);
  }

  private cfMenu() {
    const p = this.at('cf').y > 10 ? this.at('cf') : this.at('d-cf'), s = this.grid.sheet;
    const g = this.grid.isSingle() ? this.wb.currentRegion(s, this.grid.sel.ar, this.grid.sel.ac) : this.grid.range;
    const done = () => { this.grid.invalidate(); this.toast('Đã thêm quy tắc định dạng có điều kiện.', 'ok'); };
    const dlg = (k: Parameters<typeof cfRuleDialog>[3]) => () => cfRuleDialog(this.wb, s, g, k, done);
    const add = (r: Parameters<Workbook['addCF']>[1]) => () => { this.wb.addCF(s, r); done(); };
    showMenu([
      { label: 'Lớn hơn…', icon: 'cf', act: dlg('gt') },
      { label: 'Nhỏ hơn…', icon: 'cf', act: dlg('lt') },
      { label: 'Trong khoảng…', icon: 'cf', act: dlg('between') },
      { label: 'Bằng…', icon: 'cf', act: dlg('eq') },
      { label: 'Có chứa chữ…', icon: 'cf', act: dlg('text') },
      { label: 'Giá trị trùng lặp', icon: 'cf', act: add({ range: g, type: 'dup', fmt: { bg: '#ffc7ce', fg: '#9c0006' } }) },
      'sep',
      { label: '10 giá trị lớn nhất…', icon: 'sortZA', act: dlg('top') },
      { label: '10 giá trị nhỏ nhất…', icon: 'sortAZ', act: dlg('bottom') },
      { label: 'Trên trung bình', icon: 'chart', act: add({ range: g, type: 'aboveAvg', fmt: { bg: '#c6efce', fg: '#006100' } }) },
      'sep',
      { label: 'Thanh dữ liệu', icon: 'chart', act: add({ range: g, type: 'bar', colors: ['#5b9bd5'] }) },
      { label: 'Thang màu xanh – vàng – đỏ', icon: 'fill', act: add({ range: g, type: 'scale', colors: ['#63be7b', '#ffeb84', '#f8696b'] }) },
      { label: 'Thang màu đỏ – vàng – xanh', icon: 'fill', act: add({ range: g, type: 'scale', colors: ['#f8696b', '#ffeb84', '#63be7b'] }) },
      'sep',
      { label: 'Quản lý quy tắc…', icon: 'table', act: () => cfManageDialog(this.wb, s, () => this.grid.invalidate()) },
      { label: 'Xóa quy tắc ở vùng chọn', icon: 'eraser', act: () => { this.wb.clearCF(s, this.grid.range); this.grid.invalidate(); } },
      { label: 'Xóa quy tắc cả sheet', icon: 'trash', danger: true, act: () => { this.wb.clearCF(s); this.grid.invalidate(); } },
    ], p.x, p.y);
  }
  private sortFilterMenu() {
    const p = this.at('sortf');
    showMenu([
      { label: 'Sắp xếp A → Z', icon: 'sortAZ', act: () => this.quickSort(false) },
      { label: 'Sắp xếp Z → A', icon: 'sortZA', act: () => this.quickSort(true) },
      { label: 'Sắp xếp tùy chọn…', icon: 'sortAZ', act: () => this.sortDlg() },
      'sep',
      { label: this.wb.meta(this.grid.sheet).filter ? 'Tắt lọc' : 'Bật lọc', icon: 'filter', hint: 'Ctrl+Shift+L', act: () => this.toggleFilter() },
      { label: 'Xóa điều kiện lọc', icon: 'filterX', act: () => this.clearFilterConds() },
    ], p.x, p.y);
  }
  private sortRange() {
    const s = this.grid.sheet, f = this.wb.meta(s).filter, { ar, ac } = this.grid.sel;
    if (f && ar >= f.range.r1 && ar <= f.range.r2 && ac >= f.range.c1 && ac <= f.range.c2) return { g: f.range, header: true };
    const g = this.grid.isSingle() ? this.wb.currentRegion(s, ar, ac) : this.grid.range;
    // bỏ các hàng tiêu đề lớn đã gộp ở đầu vùng (vd tên bảng)
    while (g.r1 < g.r2 && this.wb.meta(s).merges.some((m) => m.r1 === g.r1 && m.r2 === g.r1 && m.c1 <= g.c1 && m.c2 >= g.c2 - 0)) g.r1++;
    return { g, header: this.wb.guessHeader(s, g) };
  }
  quickSort(desc: boolean) {
    const { g, header } = this.sortRange();
    this.safe(() => { this.wb.sort(this.grid.sheet, g, [{ col: this.grid.sel.ac, desc }], header); this.toast(`Đã sắp xếp ${desc ? 'Z → A' : 'A → Z'} theo cột ${this.colLetter(this.grid.sel.ac)}${header ? ' (giữ hàng tiêu đề)' : ''}.`, 'ok', { label: 'Hoàn tác', act: () => this.undo() }); });
  }
  private colLetter(c: number) { return toA1(0, c).replace(/\d+/, ''); }
  private sortDlg() {
    const { g } = this.sortRange();
    sortDialog(this.wb, this.grid.sheet, g, (err) => { if (err) this.toast(err, 'err'); else this.toast('Đã sắp xếp.', 'ok', { label: 'Hoàn tác', act: () => this.undo() }); this.grid.focusKeys(); });
  }
  toggleFilter() {
    const s = this.grid.sheet;
    if (this.wb.meta(s).filter) this.wb.clearFilter(s);
    else this.wb.setFilter(s, this.grid.isSingle() ? { r1: this.grid.sel.ar, c1: this.grid.sel.ac, r2: this.grid.sel.ar, c2: this.grid.sel.ac } : this.grid.range);
    this.grid.layout(); this.renderRibbon();
  }
  private clearFilterConds() {
    const s = this.grid.sheet, f = this.wb.meta(s).filter;
    if (!f) return;
    for (const c of Object.keys(f.conds)) this.wb.filterColumn(s, +c, null);
    this.grid.layout();
  }
  private filterPopover(col: number, x: number, y: number) {
    const s = this.grid.sheet, f = this.wb.meta(s).filter;
    if (!f) return;
    const all = this.wb.filterValues(s, col);
    const cur = f.conds[col];
    const box = h('div', { class: 'filter-pop' });
    const sortA = h('button', { class: 'mi' }); sortA.innerHTML = ic('sortAZ', 16) + '<span>Sắp xếp A → Z</span>';
    const sortZ = h('button', { class: 'mi' }); sortZ.innerHTML = ic('sortZA', 16) + '<span>Sắp xếp Z → A</span>';
    const doSort = (desc: boolean) => () => { closeMenu(); this.safe(() => this.wb.sort(s, f.range, [{ col, desc }], true)); };
    sortA.onclick = doSort(false); sortZ.onclick = doSort(true);
    const q = h('input', { type: 'search', placeholder: 'Tìm…', 'aria-label': 'Tìm giá trị' }) as HTMLInputElement;
    const txt = h('input', { type: 'text', placeholder: 'Lọc theo chữ chứa…', 'aria-label': 'Lọc theo chữ chứa' }) as HTMLInputElement;
    txt.value = cur?.text ?? '';
    const list = h('div', { class: 'fvals', role: 'group', 'aria-label': 'Giá trị' });
    const chosen = new Set(cur?.values ?? all);
    const shown = all.slice(0, 1000);
    const allBox = h('input', { type: 'checkbox' }) as HTMLInputElement;
    const render = () => {
      const qq = q.value.trim().toLowerCase();
      list.replaceChildren(h('label', { class: 'chk' }, allBox, ' (Chọn tất cả)'));
      allBox.checked = shown.every((v) => chosen.has(v));
      for (const v of shown) {
        if (qq && !v.toLowerCase().includes(qq)) continue;
        const cb = h('input', { type: 'checkbox' }) as HTMLInputElement;
        cb.checked = chosen.has(v);
        cb.onchange = () => { if (cb.checked) chosen.add(v); else chosen.delete(v); allBox.checked = shown.every((x) => chosen.has(x)); };
        list.append(h('label', { class: 'chk' }, cb, ' ' + (v === '' ? '(Trống)' : v)));
      }
    };
    allBox.onchange = () => { for (const v of shown) if (allBox.checked) chosen.add(v); else chosen.delete(v); render(); };
    q.oninput = render;
    render();
    const ok = h('button', { class: 'btn primary' }, 'Lọc');
    const clr = h('button', { class: 'btn' }, 'Bỏ lọc cột');
    ok.onclick = () => {
      const values = chosen.size === all.length ? undefined : [...chosen];
      const text = txt.value.trim() || undefined;
      closeMenu();
      this.wb.filterColumn(s, col, values || text ? { values, text } : null);
      this.grid.layout();
      this.grid.focusKeys();
    };
    clr.onclick = () => { closeMenu(); this.wb.filterColumn(s, col, null); this.grid.layout(); this.grid.focusKeys(); };
    box.append(sortA, sortZ, h('hr'), txt, q, list, h('div', { class: 'row-end' }, clr, ok));
    openPopover(box, x - 20, y + 10);
    q.focus();
  }
  private dvPick(x: number, y: number) {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel;
    const items = this.wb.dvList(s, ar, ac);
    if (!items.length) { this.toast('Danh sách trống.'); return; }
    showMenu(items.map((v) => ({ label: v, act: () => { this.safe(() => this.wb.setInput(s, ar, ac, v)); this.syncSelection(); this.grid.focusKeys(); } })), x - 180, y + 8);
  }
  private dvDlg() {
    const s = this.grid.sheet;
    const g = this.grid.range.r2 >= MAX_ROWS - 1 ? { ...this.grid.range, r2: Math.max(this.grid.range.r1, this.wb.used(s).rows + 500) } : this.grid.range;
    dvDialog(this.wb, s, g, (msg) => {
      const err = !!msg && msg.startsWith('Vùng');
      if (msg) this.toast(msg, err ? 'err' : 'ok');
      this.grid.invalidate();
      if (!err) this.grid.focusKeys();
    });
  }

  // ---------- tệp ----------
  fileHandle: any = null;
  fileName = '';
  async openFile() {
    const picked = await pickFile();
    if (picked) await this.openPicked(picked);
  }
  /** Đang chuyển sang sổ khác: không ghi gì thêm (tránh ghi đè “sổ mở gần nhất” bằng sổ cũ) */
  private leaving = false;
  private reload() { this.leaving = true; clearTimeout(this.saveTimer); location.reload(); }
  private opening = false;
  async openPicked(picked: { file: File; handle?: any }) {
    const { file, handle } = picked;
    if (this.leaving) return;
    if (this.opening) { this.toast('Đang mở một tệp khác, vui lòng đợi.', 'err'); return; }
    this.opening = true;
    try {
      if (/\.csv$|\.txt$/i.test(file.name)) {
        const text = await file.text();
        const id = this.wb.importCSV(text, baseName(file.name).slice(0, 31).replace(/[\\/?*[\]:]/g, '_'));
        this.grid.setSheet(id); this.renderSheets();
        this.toast(`Đã nhập ${file.name} vào sheet mới.`, 'ok');
        return;
      }
      this.toast('Đang mở ' + file.name + '…');
      const { readXlsx } = await import('../core/xlsx');
      const { wb, warnings } = await readXlsx(new Uint8Array(await file.arrayBuffer()));
      await this.saveNow(false);
      const id = 'b' + Date.now();
      await saveBook({ id, name: baseName(file.name), updated: Date.now(), data: { ...wb.toJSON(), name: baseName(file.name) }, handle, fileName: file.name });
      lsSet('xl-last', id);
      if (warnings.length) sessionStorage.setItem('xl-warn', warnings.join('\n'));
      this.fileDirty = false;
      this.reload();
    } catch (e) {
      this.toast('Không mở được tệp: ' + (e as Error).message, 'err');
    } finally {
      this.opening = false;
    }
  }
  private saving: Promise<boolean> | null = null;
  /** Lưu .xlsx; đang lưu dở thì không chạy lần thứ hai cùng lúc (giữ Ctrl+S, bấm Lưu liên tục) */
  saveXlsx(asNew: boolean, quiet = false): Promise<boolean> {
    if (this.saving) return this.saving;
    this.saving = this.doSaveXlsx(asNew, quiet).finally(() => { this.saving = null; this.updateAutoSave(); });
    return this.saving;
  }
  private async doSaveXlsx(asNew: boolean, quiet = false): Promise<boolean> {
    try {
      if (this.ed.isEditing()) this.ed.commit(0, 0);
      const ver = this.wb.ver;
      const { writeXlsx } = await import('../core/xlsx');
      const bytes = await writeXlsx(this.wb);
      const name = (this.fileName && /\.xlsx$/i.test(this.fileName) ? this.fileName : this.wb.name + '.xlsx');
      const res = await saveBytes(bytes, name, 'xlsx', asNew || !/\.xlsx$/i.test(this.fileName) ? undefined : this.fileHandle);
      if (!res) return false;
      if (res.handle) {
        const renamed = (res.handle.name ?? name) !== this.fileName;
        this.fileHandle = res.handle; this.fileName = res.handle.name ?? name;
        // như Excel: Lưu thành… đổi tên sổ theo tên tệp mới
        if (renamed) { this.wb.name = baseName(this.fileName); this.updateTitle(); }
      }
      // sửa trong lúc đang ghi thì tệp chưa có các sửa đó → vẫn coi là chưa lưu
      if (res.how === 'file') this.fileDirty = this.wb.ver !== ver;
      await this.saveNow(false);
      if (!quiet) this.toast(res.how === 'file' ? `Đã lưu ${this.fileName || name}.` : `Đã tải xuống ${name}.`, 'ok');
      if (res.how === 'file') $('#save-state').textContent = '• Đã lưu vào ' + (this.fileName || name);
    } catch (e) {
      this.toast('Không lưu được .xlsx: ' + (e as Error).message, 'err');
      return false;
    }
    return true;
  }
  /** Có thay đổi chưa ghi vào tệp .xlsx đang mở (được lưu kèm sổ nháp nên còn sau khi mở lại app) */
  fileDirty = false;
  private bindDesktop() {
    const d = desk();
    if (!d) return;
    document.documentElement.classList.add('desktop');
    d.onOpen((f: any) => this.openPicked(deskToPicked(f)));
    d.startupFile().then((f: any) => { if (f) this.openPicked(deskToPicked(f)); });
    d.onCloseRequest(() => this.closeRequest());
  }
  private closing = false;
  private async closeRequest() {
    const d = desk();
    if (this.closing) { d.closePending(); return; }
    this.closing = true;
    try {
      if (this.ed.isEditing()) this.ed.commit(0, 0);
      await this.saveNow(false);
      if (!(this.fileDirty && this.fileName)) { d.closeOk(); return; }
      d.closePending();
      const xlsx = /\.xlsx$/i.test(this.fileName);
      const body = h('p', {}, xlsx
        ? `Sổ tính có thay đổi chưa lưu vào tệp “${this.fileName}”. Bản nháp vẫn được giữ trong app, nhưng tệp trên máy sẽ chưa có các thay đổi này.`
        : `Sổ tính mở từ “${this.fileName}” có thay đổi chưa lưu ra tệp. Chọn Lưu để lưu thành tệp .xlsx mới (tệp gốc không bị ghi đè). Bản nháp vẫn được giữ trong app.`);
      let done = false;
      openDialog(xlsx ? `Lưu thay đổi vào “${this.fileName}”?` : 'Lưu thay đổi thành tệp .xlsx?', body, [
        { label: 'Không lưu', kind: 'danger', act: () => { done = true; d.closeOk(); } },
        { label: 'Hủy', act: () => {} },
        { label: xlsx ? 'Lưu' : 'Lưu thành…', kind: 'primary', act: () => {
          this.saveXlsx(false).then((ok) => {
            if (ok && !this.fileDirty) { done = true; d.closeOk(); }
            else { closeDialog(); }
          });
          return false; // giữ hộp thoại tới khi lưu xong
        } },
      ], () => { this.closing = false; if (!done) d.closeCancel(); });
    } catch {
      d.closeOk();
    }
  }


  // ---------- lệnh ô ----------
  // Số hàng/cột chèn hoặc xóa = số hàng/cột đang chọn; chỉ giới hạn theo vùng dữ liệu khi chọn cả cột/cả trang
  private rows() {
    const g = this.grid.range, u = this.wb.used(this.grid.sheet);
    return { at: g.r1, n: g.r2 >= MAX_ROWS - 1 ? Math.max(1, u.rows - g.r1) : g.r2 - g.r1 + 1 };
  }
  private cols() {
    const g = this.grid.range, u = this.wb.used(this.grid.sheet);
    return { at: g.c1, n: g.c2 >= MAX_COLS - 1 ? Math.max(1, u.cols - g.c1) : g.c2 - g.c1 + 1 };
  }
  insRows() { const { at, n } = this.rows(); this.safe(() => this.wb.insertRows(this.grid.sheet, at, n)); }
  delRows() { const { at, n } = this.rows(); this.safe(() => this.wb.deleteRows(this.grid.sheet, at, n)); }
  insCols() { const { at, n } = this.cols(); this.safe(() => this.wb.insertCols(this.grid.sheet, at, n)); }
  delCols() { const { at, n } = this.cols(); this.safe(() => this.wb.deleteCols(this.grid.sheet, at, n)); }
  clearSel() { this.safe(() => this.wb.clear(this.grid.sheet, this.grid.range)); }
  autoSum(fn = 'SUM') {
    const { ar, ac } = this.grid.sel;
    const f = this.wb.autoSum(this.grid.sheet, ar, ac, fn);
    if (f === `=${fn}()`) this.ed.begin(`=${fn}(`, 'enter');
    else this.syncSelection();
  }
  fillDir(dir: 'down' | 'right' | 'up' | 'left') {
    const g = this.grid.range;
    if (dir === 'down' && g.r2 > g.r1) this.safe(() => this.wb.fill(this.grid.sheet, { ...g, r2: g.r1 }, g));
    if (dir === 'right' && g.c2 > g.c1) this.safe(() => this.wb.fill(this.grid.sheet, { ...g, c2: g.c1 }, g));
    if (dir === 'up' && g.r2 > g.r1) this.safe(() => this.wb.fill(this.grid.sheet, { ...g, r1: g.r2 }, g));
    if (dir === 'left' && g.c2 > g.c1) this.safe(() => this.wb.fill(this.grid.sheet, { ...g, c1: g.c2 }, g));
  }
  private freeze(rows: number, cols: number) {
    if (rows === 0 && cols === 0 && !this.wb.meta(this.grid.sheet).freeze.rows && !this.wb.meta(this.grid.sheet).freeze.cols) return;
    this.wb.setFreeze(this.grid.sheet, rows, cols);
    this.grid.scroller.scrollTo(0, 0);
    this.grid.layout();
    this.toast(rows || cols ? `Đã cố định ${rows ? rows + ' hàng' : ''}${rows && cols ? ' và ' : ''}${cols ? cols + ' cột' : ''}.` : 'Đã bỏ cố định ngăn.');
  }
  private zoom(d: number) {
    const steps = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2, 3];
    const z = this.grid.zoom;
    const next = d > 0 ? steps.find((s) => s > z + 0.001) ?? 4 : [...steps].reverse().find((s) => s < z - 0.001) ?? 0.25;
    this.grid.setZoom(next);
  }

  // ---------- bảng tạm ----------
  private lastCopyText = '';
  copySel(cut: boolean) {
    const t = this.wb.copy(this.grid.sheet, this.grid.range, cut);
    this.lastCopyText = t;
    navigator.clipboard?.writeText(t).catch(() => {});
    this.grid.invalidate();
    return t;
  }
  private pasteAt(text: string | null) {
    const { ar, ac } = this.grid.sel;
    this.safe(() => {
      let g: Range | null;
      if (this.wb.clip && (text === null || normalizeNL(text) === normalizeNL(this.lastCopyText))) g = this.wb.paste(this.grid.sheet, ar, ac);
      else if (text) g = this.wb.pasteText(this.grid.sheet, ar, ac, text);
      else g = null;
      if (g) this.grid.selectRange(g);
    });
    this.grid.invalidate();
  }
  private async pasteBtn() {
    let text: string | null = null;
    try { text = await navigator.clipboard.readText(); } catch { text = null; }
    if (text === null && !this.wb.clip) { this.toast('Trình duyệt không cho đọc bảng tạm. Hãy bấm Ctrl+V để dán.'); return; }
    this.pasteAt(text);
  }
  private bindClipboard() {
    const target = () => document.activeElement?.classList.contains('cell-editor') && !this.ed.isEditing();
    document.addEventListener('copy', (e) => { if (!target()) return; e.preventDefault(); e.clipboardData?.setData('text/plain', this.copySel(false)); });
    document.addEventListener('cut', (e) => { if (!target()) return; e.preventDefault(); e.clipboardData?.setData('text/plain', this.copySel(true)); });
    document.addEventListener('paste', (e) => { if (!target()) return; e.preventDefault(); this.pasteAt(e.clipboardData?.getData('text/plain') ?? null); });
  }

  // ---------- thẻ trang tính ----------
  addSheet() {
    this.safe(() => {
      // như Excel: trang tính mới đặt ngay sau trang đang mở (một bước hoàn tác)
      const id = this.wb.addSheet(undefined, this.wb.sheetIds().indexOf(this.grid.sheet) + 1);
      this.grid.setSheet(id); this.renderSheets(); this.syncSelection();
    });
  }
  renderSheets() {
    const box = $('#sheets');
    box.innerHTML = '';
    for (const id of this.wb.visibleSheetIds()) {
      const b = document.createElement('button');
      const tc = this.wb.meta(id).tabColor;
      b.className = 'st' + (id === this.grid.sheet ? ' on' : '') + (tc ? ' tc' : '');
      if (tc) b.style.setProperty('--tab', tc);
      b.textContent = this.wb.sheetName(id);
      b.dataset.sheet = String(id);
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(id === this.grid.sheet));
      b.onmousedown = (e) => { if (this.ed.isEditing()) e.preventDefault(); };
      b.onclick = () => {
        // đang gõ công thức ở chế độ trỏ: chuyển sheet để chọn ô tham chiếu, chưa nhập
        if (this.ed.isEditing() && this.ed.pointMode()) { if (id !== this.grid.sheet) { this.grid.setSheet(id); this.renderSheets(); } this.ed.refocus(); return; }
        if (this.ed.isEditing()) this.ed.commit(0, 0);
        if (id !== this.grid.sheet) { this.grid.setSheet(id); this.renderSheets(); this.syncSelection(); }
        this.grid.focusKeys();
      };
      b.ondblclick = () => this.renameSheet(id);
      b.oncontextmenu = (e) => { e.preventDefault(); this.sheetMenu(id, e.clientX, e.clientY); };
      let lp = 0;
      b.onpointerdown = (e) => { if (e.pointerType === 'touch') lp = window.setTimeout(() => this.sheetMenu(id, e.clientX, e.clientY), 550); };
      b.onpointerup = b.onpointercancel = () => clearTimeout(lp);
      box.appendChild(b);
    }
    box.querySelector('.st.on')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    this.sheetSig = this.sigOf();
  }
  private sigOf() { return this.wb.sheetIds().map((i) => i + ':' + this.wb.sheetName(i) + ':' + (this.wb.meta(i).hidden ? 'h' : '') + ':' + (this.wb.meta(i).tabColor ?? '')).join('|'); }
  private sheetMenu(id: number, x: number, y: number) {
    const hidden = this.wb.sheetIds().filter((i) => this.wb.isSheetHidden(i));
    showMenu([
      { label: 'Chèn…', icon: 'sheet', act: () => this.addSheet() },
      { label: 'Xóa', icon: 'trash', act: () => this.removeSheet(id) },
      { label: 'Đổi tên', act: () => this.renameSheet(id) },
      { label: 'Di chuyển hoặc Sao chép…', icon: 'move', act: () => this.moveCopyDlg(id) },
      { label: 'Xem Mã', act: () => this.later('Xem Mã (VBA)', 'ngoài phạm vi') },
      { label: 'Bảo vệ Trang tính…', icon: 'protectSheet', act: () => this.later('Bảo vệ Trang tính', '0.9') },
      { label: 'Màu Tab', icon: 'tabColor', sub: () => colorPalette((c) => { this.wb.setTabColor(id, c); this.renderSheets(); }, 'Không Màu') },
      'sep',
      { label: 'Ẩn', act: () => this.hideSheet(id) },
      { label: 'Bỏ ẩn…', disabled: !hidden.length, act: () => this.unhideSheetDlg() },
      'sep',
      { label: 'Chọn Tất cả Trang tính', act: () => this.later('Nhóm trang tính', '0.9') },
    ], x, y);
  }
  hideSheet(id: number) {
    this.safe(() => {
      this.wb.setSheetHidden(id, true);
      if (id === this.grid.sheet) this.grid.setSheet(this.wb.visibleSheetIds()[0]);
      this.renderSheets(); this.syncSelection();
    });
  }
  unhideSheetDlg() {
    const hidden = this.wb.sheetIds().filter((i) => this.wb.isSheetHidden(i));
    if (!hidden.length) { this.toast('Không có trang tính nào đang ẩn.'); return; }
    const list = h('select', { size: String(Math.min(8, Math.max(3, hidden.length))), 'aria-label': 'Trang tính đang ẩn' }) as HTMLSelectElement;
    for (const i of hidden) list.append(h('option', { value: String(i) }, this.wb.sheetName(i)));
    list.selectedIndex = 0;
    openDialog('Bỏ ẩn', h('div', { class: 'stack' }, field('Bỏ ẩn trang tính:', list)), [
      { label: 'Hủy', act: () => {} },
      { label: 'OK', kind: 'primary', act: () => { const id = +list.value; this.safe(() => this.wb.setSheetHidden(id, false)); this.grid.setSheet(id); this.renderSheets(); this.syncSelection(); } },
    ], () => this.grid.focusKeys());
  }
  moveCopyDlg(id: number) {
    const ids = this.wb.sheetIds();
    const before = h('select', { size: String(Math.min(8, ids.length + 1)), 'aria-label': 'Trước trang tính' }) as HTMLSelectElement;
    ids.forEach((i, k) => before.append(h('option', { value: String(k) }, this.wb.sheetName(i))));
    before.append(h('option', { value: String(ids.length) }, '(di chuyển đến cuối)'));
    before.value = String(Math.min(ids.length, ids.indexOf(id) + 1));
    const copy = h('input', { type: 'checkbox', id: 'mv-copy' }) as HTMLInputElement;
    openDialog('Di chuyển hoặc Sao chép', h('div', { class: 'stack' }, h('p', { class: 'muted' }, `Di chuyển trang tính đã chọn “${this.wb.sheetName(id)}”`), field('Trước trang tính:', before), h('label', { class: 'chk' }, copy, ' Tạo bản sao')), [
      { label: 'Hủy', act: () => {} },
      { label: 'OK', kind: 'primary', act: () => {
        const at = +before.value;
        this.safe(() => {
          if (copy.checked) { const n = this.wb.copySheet(id, at); this.grid.setSheet(n); }
          else { const cur = this.wb.sheetIds().indexOf(id); this.wb.moveSheet(id, at > cur ? at - 1 : at); }
        });
        this.renderSheets(); this.syncSelection();
      } },
    ], () => this.grid.focusKeys());
  }
  removeSheet(id: number) {
    const n = this.wb.sheetName(id);
    if (this.wb.visibleSheetIds().length <= 1 && !this.wb.isSheetHidden(id)) { this.toast('Sổ làm việc phải còn ít nhất một trang tính hiển thị.', 'err'); return; }
    this.safe(() => {
      this.wb.removeSheet(id);
      if (id === this.grid.sheet) this.grid.setSheet(this.wb.visibleSheetIds()[0]);
      this.renderSheets(); this.syncSelection();
      this.toast(`Đã xóa trang tính “${n}”.`, 'ok', { label: 'Hoàn tác', act: () => this.undo() });
    });
  }
  renameSheet(id: number) {
    const el = document.querySelector<HTMLElement>(`#sheets [data-sheet="${id}"]`);
    if (!el) return;
    const inp = document.createElement('input');
    inp.className = 'st-edit';
    inp.value = this.wb.sheetName(id);
    inp.setAttribute('aria-label', 'Tên trang tính');
    el.replaceWith(inp);
    inp.focus(); inp.select();
    let done = false;
    const finish = (ok: boolean) => {
      if (done) return; done = true;
      if (ok) { try { this.wb.renameSheet(id, inp.value); } catch (e) { this.toast((e as Error).message, 'err'); } }
      this.renderSheets(); this.grid.invalidate(); this.syncSelection(); this.grid.focusKeys();
    };
    inp.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); finish(true); } if (e.key === 'Escape') { e.preventDefault(); finish(false); } };
    inp.onblur = () => finish(true);
  }

  // ---------- thanh trạng thái, hộp tên ----------
  private static AGG: [string, string, boolean][] = [['avg', 'Trung bình', true], ['count', 'Số đếm', true], ['ncount', 'Số đếm Số', false], ['min', 'Tối thiểu', false], ['max', 'Tối đa', false], ['sum', 'Tổng', true]];
  private statusOn(): Record<string, boolean> {
    const d: Record<string, boolean> = { mode: true, zoom: true, slider: true, views: true };
    for (const [k, , on] of App.AGG) d[k] = on;
    try { return { ...d, ...JSON.parse(lsGet('xl-status') ?? '{}') }; } catch { return d; }
  }
  private bindFoot() {
    $('#zoom-out').onclick = () => this.zoom(-1);
    $('#zoom-in').onclick = () => this.zoom(1);
    $('#zoom-val').onclick = () => this.zoomDlg();
    const sl = $<HTMLInputElement>('#zoom-slider');
    // thanh trượt như Excel: nửa trái 10–100%, nửa phải 100–400%
    sl.oninput = () => { const v = +sl.value; this.grid.setZoom(v <= 50 ? (10 + (v / 50) * 90) / 100 : (100 + ((v - 50) / 50) * 300) / 100); };
    $('#sheet-prev').onclick = () => this.stepSheet(-1);
    $('#sheet-next').onclick = () => this.stepSheet(1);
    $('#sheet-add').onclick = () => this.addSheet();
    document.querySelectorAll<HTMLElement>('#views [data-view]').forEach((b) => { b.onclick = () => this.setView(b.dataset.view as 'normal'); });
    $('#status').oncontextmenu = (e) => {
      e.preventDefault();
      const on = this.statusOn();
      const tog = (k: string) => () => { lsSet('xl-status', JSON.stringify({ ...this.statusOn(), [k]: !on[k] })); this.applyStatus(); this.updateStatus(); };
      showMenu([
        { label: 'Chế độ Ô', check: on.mode, act: tog('mode') },
        'sep',
        ...App.AGG.map(([k, l]) => ({ label: l, check: on[k], act: tog(k) }) as MenuItem),
        'sep',
        { label: 'Lối tắt Dạng xem', check: on.views, act: tog('views') },
        { label: 'Thanh trượt Thu phóng', check: on.slider, act: tog('slider') },
        { label: 'Thu phóng', check: on.zoom, act: tog('zoom') },
      ], e.clientX, Math.max(8, e.clientY - 330));
    };
    this.applyStatus();
  }
  private applyStatus() {
    const on = this.statusOn();
    $('#mode').hidden = !on.mode;
    $('#views').hidden = !on.views;
    $('#zoom-slider').hidden = $('#zoom-out').hidden = $('#zoom-in').hidden = !on.slider;
    $('#zoom-val').hidden = !on.zoom;
  }
  private syncZoomUi() {
    const z = Math.round(this.grid.zoom * 100);
    $('#zoom-val').textContent = z + '%';
    $<HTMLInputElement>('#zoom-slider').value = String(z <= 100 ? ((z - 10) / 90) * 50 : 50 + ((z - 100) / 300) * 50);
  }
  private bindNameBox() {
    const nb = $<HTMLInputElement>('#namebox');
    nb.onfocus = () => nb.select();
    nb.onkeydown = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); this.syncSelection(); this.grid.focusKeys(); }
      if (e.key !== 'Enter') return;
      e.preventDefault();
      const g = parseRange(nb.value.replace(/\s/g, ''));
      if (!g) { this.toast('Tham chiếu không hợp lệ. Ví dụ: B5 hoặc A1:C10', 'err'); return; }
      this.grid.selectRange(g); this.grid.focusKeys();
    };
    $('#namebox-dd').onclick = () => this.later('Danh sách Tên', '0.7');
    $('#fx-fn').onclick = () => this.funcPicker();
    $('#fx-exp').onclick = () => { $('#app').classList.toggle('fx-tall'); this.grid.layout(); };
  }
  syncSelection() {
    const g = this.grid.range, s = this.grid.sheet, { ar, ac } = this.grid.sel;
    const nb = $<HTMLInputElement>('#namebox');
    if (document.activeElement !== nb) {
      const fullC = g.r1 === 0 && g.r2 === MAX_ROWS - 1, fullR = g.c1 === 0 && g.c2 === MAX_COLS - 1;
      nb.value = fullC && fullR ? toA1(ar, ac) : g.r1 === g.r2 && g.c1 === g.c2 ? toA1(ar, ac) : this.wb.mergeAt(s, ar, ac) && this.grid.isSingle() ? toA1(ar, ac) : toA1(ar, ac);
    }
    if (!this.ed.isEditing()) $<HTMLTextAreaElement>('#fx-input').value = this.wb.rawText(s, ar, ac);
    this.refreshToggles();
    if (!this.bookChanging) this.pivot?.follow(s, ar, ac);
    this.updateStatus();
    this.announce();
    this.syncZoomUi();
    document.querySelector('#qat [data-q="undo"]')?.toggleAttribute('disabled', !this.wb.canUndo());
    document.querySelector('#qat [data-q="redo"]')?.toggleAttribute('disabled', !this.wb.canRedo());
    void rangeToA1;
  }
  private updateStatus() {
    const g = this.grid.range, s = this.grid.sheet, u = this.wb.used(s);
    const r2 = Math.min(g.r2, u.rows - 1), c2 = Math.min(g.c2, u.cols - 1);
    const agg = $('#agg');
    if ((g.r1 === g.r2 && g.c1 === g.c2) || r2 < g.r1 || c2 < g.c1 || (r2 - g.r1 + 1) * (c2 - g.c1 + 1) > 1_000_000) { agg.innerHTML = ''; return; }
    let sum = 0, nums = 0, cnt = 0, mn = Infinity, mx = -Infinity;
    for (let r = g.r1; r <= r2; r++) {
      if (this.wb.isRowHidden(s, r)) continue;
      for (let c = g.c1; c <= c2; c++) {
        const v = this.wb.value(s, r, c);
        if (v === null || v === '') continue;
        cnt++;
        if (typeof v === 'number') { sum += v; nums++; if (v < mn) mn = v; if (v > mx) mx = v; }
      }
    }
    const fmt = this.wb.effFmt(s, g.r1, g.c1);
    const f = (n: number) => (fmt && !/%/.test(fmt) ? formatValue(n, fmt).text : n.toLocaleString('vi-VN', { maximumFractionDigits: 10 }));
    const on = this.statusOn();
    const parts: [string, string, string][] = [];
    if (cnt >= 2) {
      if (nums && on.avg) parts.push(['avg', 'Trung bình', f(sum / nums)]);
      if (on.count) parts.push(['count', 'Số đếm', String(cnt)]);
      if (nums && on.ncount) parts.push(['ncount', 'Số đếm Số', String(nums)]);
      if (nums && on.min) parts.push(['min', 'Tối thiểu', f(mn)]);
      if (nums && on.max) parts.push(['max', 'Tối đa', f(mx)]);
      if (nums && on.sum) parts.push(['sum', 'Tổng', f(sum)]);
    }
    agg.innerHTML = '';
    for (const [k, l, v] of parts) { const sp = document.createElement('span'); sp.dataset.agg = k; sp.textContent = `${l}: ${v}`; agg.appendChild(sp); }
  }

  private annTimer = 0;
  private dvPrompt() {
    const d = this.wb.dvAt(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac);
    if (d?.prompt && !this.ed.isEditing()) this.setMode(d.prompt);
    else if (!this.ed.isEditing()) this.setMode('Sẵn sàng');
  }
  private announce() {
    this.dvPrompt();
    clearTimeout(this.annTimer);
    this.annTimer = window.setTimeout(() => {
      const { ar, ac } = this.grid.sel;
      const d = this.wb.display(this.grid.sheet, ar, ac);
      const t = toA1(ar, ac) + ': ' + (d.text || 'trống');
      $('#sr-cell').textContent = t;
      this.ed.label(t);
    }, 150);
  }
  // ---------- biểu đồ, PivotTable, in ----------
  private dataRange() {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel;
    let g = this.grid.range;
    if (g.r1 === g.r2 && g.c1 === g.c2) g = this.wb.currentRegion(s, ar, ac);
    const u = this.wb.used(s);
    return g.r1 < u.rows && g.c1 < u.cols ? g : null;
  }
  insertChart(type?: ChartType) {
    const s = this.grid.sheet, g = this.dataRange();
    if (!g) { this.toast('Hãy chọn vùng dữ liệu (có tiêu đề và số) trước khi chèn biểu đồ.', 'err'); return; }
    chartDialog({ wb: this.wb, sheet: s, range: g, type, onDone: (ch) => {
      if (!ch) { this.grid.focusKeys(); return; }
      const c = Math.min(g.c2 + 2, MAX_COLS - 8), r = g.r1;
      this.safe(() => {
        // xếp lệch nếu đã có biểu đồ ở cùng chỗ
        let dy = 0;
        while (this.wb.meta(s).charts.some((x) => x.anchor.r === r && x.anchor.c === c && x.anchor.dy === dy)) dy += 28;
        const id = this.wb.addChart(s, { ...ch, anchor: { r, c, dx: dy, dy } });
        this.grid.ensureVisible(r, c);
        this.grid.invalidate();
        requestAnimationFrame(() => this.charts.select(id));
        this.toast('Đã chèn biểu đồ. Kéo để di chuyển, kéo góc để đổi cỡ, nhấp đúp để sửa.', 'ok');
      });
    } });
  }
  editChart(id: string) {
    const s = this.grid.sheet, ch = this.wb.meta(s).charts.find((x) => x.id === id);
    if (!ch) return;
    chartDialog({ wb: this.wb, sheet: s, chart: ch, onDone: (nc) => { if (nc) this.safe(() => this.wb.updateChart(s, id, nc)); } });
  }
  insertPivot() {
    const s = this.grid.sheet, g = this.dataRange();
    if (!g) { this.toast('Hãy chọn vùng dữ liệu có hàng tiêu đề trước khi tạo PivotTable.', 'err'); return; }
    pivotCreateDialog(this.wb, s, g, { r: this.grid.sel.ar, c: this.grid.sel.ac }, (res) => {
      if (!res) { this.grid.focusKeys(); return; }
      const p = this.wb.meta(res.sheet).pivots.find((x) => x.id === res.id)!;
      if (res.sheet !== this.grid.sheet) { this.grid.setSheet(res.sheet); this.renderSheets(); }
      this.grid.select(p.at.r, p.at.c);
      this.pivot.show(res.sheet, res.id);
      this.toast('Đã tạo PivotTable. Tích chọn trường ở ngăn bên phải hoặc kéo vào Hàng / Cột / Giá trị.', 'ok');
    });
  }

  setMode(m: string) { $('#mode').textContent = m; }

  // ---------- thay đổi & lưu ----------
  private bookChanging = false;
  private onBookChange() {
    this.fileDirty = true;
    this.renderSheetsIfNeeded();
    this.pivot.refreshView();
    this.bookChanging = true;
    try { this.syncSelection(); } finally { this.bookChanging = false; }
    this.scheduleSave();
  }
  private sheetSig = '';
  private renderSheetsIfNeeded() {
    const sig = this.wb.sheetIds().map((i) => i + ':' + this.wb.sheetName(i)).join('|');
    if (sig !== this.sheetSig) { this.sheetSig = sig; if (!this.wb.sheetIds().includes(this.grid.sheet)) this.grid.setSheet(this.wb.sheetIds()[0]); this.renderSheets(); }
  }
  private scheduleSave() {
    clearTimeout(this.saveTimer);
    this.scheduleAuto();
    $('#save-state').textContent = '• Chưa lưu';
    this.dirtySave = true;
    let cells = 0;
    for (const id of this.wb.sheetIds()) { const u = this.wb.used(id); cells += u.rows * u.cols; }
    this.saveTimer = window.setTimeout(() => this.saveNow(false), cells > 300_000 ? 8000 : 1200);
  }
  async saveNow(announce: boolean) {
    clearTimeout(this.saveTimer);
    if (this.leaving) return;
    try {
      await saveBook({ id: this.bookId, name: this.wb.name, updated: Date.now(), data: this.wb.toJSON(), handle: this.fileHandle ?? undefined, fileName: this.fileName || undefined, fileDirty: this.fileDirty });
      this.dirtySave = false;
      lsSet('xl-last', this.bookId);
      const t = new Date();
      $('#save-state').textContent = `• Đã lưu lúc ${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}`;
      if (announce) this.toast('Đã lưu trên máy này.', 'ok');
    } catch (e) {
      $('#save-state').textContent = '• Không lưu được';
      this.toast('Không lưu được: ' + (e as Error).message, 'err');
    }
  }
  async newBook() {
    await this.saveNow(false);
    const id = 'b' + Date.now();
    lsSet('xl-last', id);
    location.hash = '';
    this.reload();
  }
  // ---------- chủ đề Office ----------
  themePref(): 'colorful' | 'white' | 'darkgray' | 'black' | 'system' {
    const v = lsGet('xl-theme');
    if (v === 'light') return 'colorful';
    if (v === 'dark') return 'black';
    return (v as any) || 'colorful';
  }
  private bindTheme() {
    this.applyTheme();
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => this.applyTheme());
  }
  setTheme(t: string) {
    const map: Record<string, string> = { light: 'colorful', dark: 'black' };
    lsSet('xl-theme', map[t] ?? t);
    this.applyTheme();
    this.ribbon?.render();
  }
  private applyTheme() {
    const t = this.themePref();
    const eff = t === 'system' ? (matchMedia('(prefers-color-scheme: dark)').matches ? 'black' : 'colorful') : t;
    document.documentElement.dataset.theme = eff;
    requestAnimationFrame(() => {
      const cs = getComputedStyle(document.documentElement);
      const bg = cs.getPropertyValue('--title-bg').trim(), fg = cs.getPropertyValue('--title-fg').trim();
      document.querySelector('meta[name="theme-color"]')?.setAttribute('content', bg);
      desk()?.setTitleBar?.({ color: bg, symbolColor: fg });
      this.grid.readColors();
    });
  }

  // ---------- menu chuột phải (ô, tiêu đề hàng, tiêu đề cột) + thanh mini ----------
  private miniBar(): HTMLElement {
    const bar = document.createElement('div');
    bar.className = 'mini-tb';
    bar.setAttribute('role', 'toolbar');
    bar.setAttribute('aria-label', 'Thanh công cụ mini');
    const st = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac);
    const combo = (m: string, w: number, val: string, list: (r: DOMRect) => void, set: (v: string) => void) => {
      const c = document.createElement('div');
      c.className = 'combo'; c.dataset.m = m; c.style.width = w + 'px';
      c.innerHTML = `<input aria-label="${m === 'fn' ? 'Phông' : 'Cỡ phông'}"><button class="dd" tabindex="-1" aria-label="Danh sách">${ic('chevDown', 10)}</button>`;
      const i = c.querySelector('input')!; i.value = val;
      i.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); closeMenu(); set(i.value.trim()); this.grid.focusKeys(); } };
      (c.querySelector('.dd') as HTMLElement).onclick = () => { const r = c.getBoundingClientRect(); closeMenu(); list(r); };
      return c;
    };
    const btn = (m: string, icon: string, label: string, act: (r: DOMRect) => void, barColor?: string) => {
      const b = document.createElement('button');
      b.dataset.m = m; b.title = label; b.setAttribute('aria-label', label);
      b.innerHTML = `<span class="ico">${ic(icon, 16)}${barColor ? `<i class="cbar" style="background:${barColor}"></i>` : ''}</span>`;
      b.onmousedown = (e) => e.preventDefault();
      b.onclick = () => { const r = b.getBoundingClientRect(); closeMenu(); act(r); this.grid.invalidate(); this.syncSelection(); };
      return b;
    };
    bar.append(
      combo('fn', 118, st.fn ?? DEFAULT_FONT, (r) => this.fontMenu(r), (v) => v && this.applyStyle({ fn: v === DEFAULT_FONT ? undefined : v })),
      combo('fs', 42, String(st.fs ?? 11), (r) => this.sizeMenu(r), (v) => { const n = parseFloat(v); if (n >= 1 && n <= 409) this.applyStyle({ fs: n === 11 ? undefined : n }); }),
      btn('grow', 'grow', 'Tăng Cỡ Phông', () => this.growFont(1)),
      btn('shrink', 'shrink', 'Giảm Cỡ Phông', () => this.growFont(-1)),
      btn('acct', 'money', 'Định dạng Số Kế toán', () => this.applyStyle({ fmt: NF[3][1] })),
      btn('pct', 'percent', 'Kiểu Phần trăm', () => this.applyStyle({ fmt: '0%' })),
      btn('comma', 'comma', 'Kiểu Dấu phẩy', () => this.applyStyle({ fmt: '#,##0' })),
      btn('b', 'bold', 'Đậm', () => this.toggle('b')),
      btn('i', 'italic', 'Nghiêng', () => this.toggle('i')),
      btn('center', 'center', 'Căn giữa', () => this.align('c')),
      btn('bg', 'fill', 'Màu Tô', (r) => this.colorMenu('bg', r), this.lastFill),
      btn('fg', 'textColor', 'Màu Phông', (r) => this.colorMenu('fg', r), this.lastFont),
      btn('border', 'borderBottom', 'Viền', (r) => { this.anchor = { x: r.left, y: r.bottom + 2 }; this.borderMenu(); }),
      btn('merge', 'merge2', 'Gộp & Căn giữa', () => this.doMerge(true)),
      btn('dinc', 'decInc', 'Tăng Số thập phân', () => this.decimals(1)),
      btn('ddec', 'decDec', 'Giảm Số thập phân', () => this.decimals(-1)),
    );
    return bar;
  }
  private pasteRow(): HTMLElement {
    const row = document.createElement('div');
    row.className = 'paste-row';
    const items: [string, string, (() => void) | null][] = [['paste', 'Dán (P)', () => this.pasteBtn()], ['pasteValues', 'Giá trị (V)', null], ['pasteFormula', 'Công thức (F)', null], ['transpose', 'Chuyển vị (T)', null], ['pasteFmt', 'Định dạng (R)', null]];
    for (const [icon, label, act] of items) {
      const b = document.createElement('button');
      b.innerHTML = ic(icon, 16); b.title = label + (act ? '' : ' — có ở bản 0.6'); b.setAttribute('aria-label', b.title);
      if (!act) b.classList.add('later');
      b.onclick = () => { closeMenu(); if (act) act(); else this.later(label.replace(/ \(.\)$/, ''), '0.6'); };
      row.appendChild(b);
    }
    return row;
  }
  private contextMenu(x: number, y: number, zone: 'cell' | 'row' | 'col') {
    const later = (l: string, v: string) => () => this.later(l, v);
    const head: MenuItem[] = [
      { label: 'Cắt', icon: 'cut', hint: 'Ctrl+X', act: () => this.copySel(true) },
      { label: 'Sao chép', icon: 'copy', hint: 'Ctrl+C', act: () => this.copySel(false) },
      { label: 'Tùy chọn Dán:', icon: 'paste', act: () => this.pasteBtn(), extra: () => this.pasteRow() },
      { label: 'Dán Đặc biệt…', act: later('Dán Đặc biệt', '0.6') },
      'sep',
    ];
    let items: MenuItem[];
    if (zone === 'row' || zone === 'col') {
      const R = zone === 'row';
      items = [...head,
        { label: 'Chèn', act: () => (R ? this.insRows() : this.insCols()) },
        { label: 'Xóa', act: () => (R ? this.delRows() : this.delCols()) },
        { label: 'Xóa Nội dung', act: () => this.clearSel() },
        'sep',
        { label: 'Định dạng Ô…', icon: 'cellFmt', act: later('Định dạng Ô', '0.6') },
        R ? { label: 'Chiều cao Hàng…', act: () => this.rowHeightDlg() } : { label: 'Độ rộng Cột…', act: () => this.colWidthDlg() },
        { label: 'Ẩn', act: () => this.hideSel(R ? 'r' : 'c', true) },
        { label: 'Bỏ ẩn', act: () => this.unhideAround(R ? 'r' : 'c') },
      ];
    } else {
      items = [...head,
        { label: 'Thông tin Thông minh', icon: 'search', act: later('Thông tin Thông minh', 'ngoài phạm vi') },
        'sep',
        { label: 'Chèn…', act: () => this.cellsDlg(true) },
        { label: 'Xóa…', act: () => this.cellsDlg(false) },
        { label: 'Xóa Nội dung', act: () => this.clearSel() },
        'sep',
        { label: 'Lọc', icon: 'filter', sub: [
          { label: 'Bật/Tắt Lọc', hint: 'Ctrl+Shift+L', act: () => this.toggleFilter() },
          { label: 'Lọc theo Giá trị của Ô Đã chọn', act: () => this.filterByCell() },
          { label: 'Xóa Bộ lọc', act: () => this.clearFilterConds() },
          { label: 'Áp dụng lại', act: () => { this.wb.reapplyFilter(this.grid.sheet); this.grid.layout(); } },
        ] },
        { label: 'Sắp xếp', icon: 'sortAZ', sub: [
          { label: 'Sắp xếp từ Nhỏ nhất đến Lớn nhất', icon: 'sortAZ', act: () => this.quickSort(false) },
          { label: 'Sắp xếp từ Lớn nhất đến Nhỏ nhất', icon: 'sortZA', act: () => this.quickSort(true) },
          { label: 'Sắp xếp Tùy chỉnh…', act: () => this.sortDlg() },
        ] },
        'sep',
        { label: 'Bình luận Mới', icon: 'comment', act: later('Bình luận', '0.7') },
        { label: 'Ghi chú Mới', icon: 'note', act: later('Ghi chú', '0.7') },
        'sep',
        { label: 'Định dạng Ô…', icon: 'cellFmt', hint: 'Ctrl+1', act: later('Định dạng Ô', '0.6') },
        { label: 'Chọn từ Danh sách Thả xuống…', act: () => this.pickFromList() },
        { label: 'Xác định Tên…', icon: 'tag', act: later('Xác định Tên', '0.7') },
        { label: 'Liên kết', icon: 'link', hint: 'Ctrl+K', act: later('Liên kết', '0.9') },
      ];
    }
    showMenu(items, x, y, { mini: this.miniBar() });
  }
  private unhideAround(ax: 'r' | 'c') {
    const g = this.grid.range;
    // như Excel: chọn hai hàng/cột hai bên vùng ẩn rồi Bỏ ẩn; chọn một hàng/cột thì bỏ ẩn liền kề
    const a = ax === 'r' ? g.r1 : g.c1, b = ax === 'r' ? g.r2 : g.c2;
    this.safe(() => { if (ax === 'r') this.wb.hideRows(this.grid.sheet, Math.max(0, a - (a === b ? 1 : 0)), b + (a === b ? 1 : 0), false); else this.wb.hideCols(this.grid.sheet, Math.max(0, a - (a === b ? 1 : 0)), b + (a === b ? 1 : 0), false); });
    this.grid.layout(); this.syncSelection();
  }
  private filterByCell() {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel;
    if (!this.wb.meta(s).filter) this.wb.setFilter(s, { r1: ar, c1: ac, r2: ar, c2: ac });
    const f = this.wb.meta(s).filter!;
    if (ar === f.range.r1) { this.toast('Hãy chọn một ô dữ liệu (không phải tiêu đề) để lọc theo giá trị.'); return; }
    this.wb.filterColumn(s, ac, { values: [this.wb.display(s, ar, ac).text] });
    this.grid.layout();
  }
  /** Chọn từ Danh sách Thả xuống (Alt+↓): các giá trị đã có trong cột phía trên/dưới */
  private pickFromList() {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel;
    const set = new Set<string>();
    for (let r = ar - 1; r >= 0 && this.wb.rawText(s, r, ac) !== ''; r--) set.add(this.wb.display(s, r, ac).text);
    for (let r = ar + 1; this.wb.rawText(s, r, ac) !== ''; r++) set.add(this.wb.display(s, r, ac).text);
    const vals = [...set].filter(Boolean).sort((a, b) => a.localeCompare(b, 'vi'));
    if (!vals.length) { this.toast('Không có giá trị nào ở các ô liền kề trong cột.'); return; }
    const p = this.grid.cellRectM(ar, ac), b = this.grid.scroller.getBoundingClientRect();
    showMenu(vals.slice(0, 200).map((v) => ({ label: v, act: () => { this.safe(() => this.wb.setInput(s, ar, ac, v)); this.syncSelection(); this.grid.focusKeys(); } })), b.left + p.x, b.top + p.y + p.h);
  }

  // ---------- cài đặt PWA ----------
  private bindInstall() {
    window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); this.installEvt = e; $('#btn-install').hidden = false; });
    $('#btn-install').onclick = () => this.install();
    window.addEventListener('appinstalled', () => { $('#btn-install').hidden = true; this.toast('Đã cài ExcelLike.', 'ok'); });
  }
  async install() {
    if (this.installEvt) { this.installEvt.prompt(); await this.installEvt.userChoice; this.installEvt = null; $('#btn-install').hidden = true; return; }
    const ios = /iphone|ipad|ipod/i.test(navigator.userAgent);
    if (matchMedia('(display-mode: standalone)').matches) this.toast('ExcelLike đang chạy như ứng dụng đã cài.');
    else if (ios) this.toast('Trên iPhone/iPad: bấm nút Chia sẻ rồi chọn “Thêm vào MH chính”.');
    else this.toast('Mở bằng Chrome hoặc Edge rồi bấm biểu tượng cài đặt trên thanh địa chỉ.');
  }

  // ---------- hàm ----------
  funcPicker() {
    const box = document.createElement('div');
    box.className = 'fn-picker';
    box.innerHTML = `<label class="fn-search">${ic('fx', 16)}<input id="fn-q" placeholder="Tìm hàm, ví dụ: tra cứu, SUMIF" aria-label="Tìm hàm"></label><div class="fn-list" role="listbox"></div>`;
    const list = box.querySelector('.fn-list') as HTMLElement;
    const all = this.wb.hf.getRegisteredFunctionNames().sort();
    const render = (q: string) => {
      const qq = q.trim().toLowerCase();
      const names = all.filter((n) => !qq || n.toLowerCase().includes(qq) || (FUNC_INFO[n]?.join(' ').toLowerCase().includes(qq) ?? false));
      const known = names.filter((n) => FUNC_INFO[n]), rest = names.filter((n) => !FUNC_INFO[n]);
      list.innerHTML = [...known, ...rest].slice(0, 80).map((n) => {
        const i = FUNC_INFO[n];
        return `<button role="option" data-n="${n}"><b>${n}</b><span>${i ? i[2] + ' · ' + i[1] : ''}</span></button>`;
      }).join('') || '<p class="empty">Không có hàm phù hợp.</p>';
    };
    render('');
    const q = box.querySelector('#fn-q') as HTMLInputElement;
    q.oninput = () => render(q.value);
    q.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); (list.querySelector('button') as HTMLButtonElement | null)?.click(); } };
    list.onclick = (e) => {
      const b = (e.target as HTMLElement).closest('button');
      if (!b) return;
      closeMenu();
      this.ed.begin('=' + b.dataset.n + '(', 'enter');
    };
    const r = $('#gridwrap').getBoundingClientRect();
    openPopover(box, r.left + 40, r.top + 10);
    q.focus();
  }

  // ---------- lệnh mới của 0.5 ----------
  growFont(d: 1 | -1) {
    const cur = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).fs ?? 11;
    const n = d > 0 ? SIZES.find((x) => x > cur) ?? Math.min(409, cur + 10) : [...SIZES].reverse().find((x) => x < cur) ?? Math.max(1, cur - 1);
    this.applyStyle({ fs: n === 11 ? undefined : n });
    this.fitRows();
  }
  indent(d: 1 | -1) {
    const st = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac);
    const n = Math.max(0, Math.min(15, (st.ind ?? 0) + d));
    this.applyStyle({ ind: n || undefined, al: n && st.al !== 'r' ? st.al ?? 'l' : st.al });
  }
  doMerge(center: boolean) {
    const s = this.grid.sheet, g = this.grid.range;
    if (center && this.wb.mergeAt(s, this.grid.sel.ar, this.grid.sel.ac)) { this.wb.unmerge(s, g); this.grid.invalidate(); return; }
    this.safe(() => {
      const lost = this.wb.merge(s, g, center);
      this.grid.selectRange(this.wb.expandToMerges(s, g));
      if (lost) this.toast(`Đã gộp ô. Chỉ giữ giá trị ô trên cùng bên trái (bỏ ${lost} ô khác).`, 'ok', { label: 'Hoàn tác', act: () => this.undo() });
    });
  }
  private nfMenu(r: DOMRect) {
    const v = this.wb.value(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac);
    showMenu([
      ...NF.map(([l, f, ex]) => ({ label: l, hint: typeof v === 'number' ? formatValue(v, f).text : ex, act: () => this.applyStyle({ fmt: f }) }) as MenuItem),
      'sep',
      { label: 'Định dạng Số Khác…', act: () => this.later('Định dạng Số Khác (hộp Định dạng Ô)', '0.6') },
    ], r.left, r.bottom + 2);
  }
  colorMenu(k: 'fg' | 'bg', r: DOMRect) {
    const box = colorPalette((c) => {
      if (k === 'bg') { if (c) this.lastFill = c; this.applyStyle({ bg: c }); } else { if (c) this.lastFont = c; this.applyStyle({ fg: c }); }
      this.ribbon.refresh(); this.grid.focusKeys();
    }, k === 'bg' ? 'Không Tô' : 'Tự động');
    openPopover(box, r.left, r.bottom + 2);
  }
  fontMenu(r: DOMRect) {
    const cur = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).fn ?? DEFAULT_FONT;
    showMenu(FONTS.map((f) => ({ label: f, check: f === cur, act: () => this.applyStyle({ fn: f === DEFAULT_FONT ? undefined : f }) })), r.left, r.bottom + 2, { cls: 'fonts' });
  }
  sizeMenu(r: DOMRect) {
    const cur = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).fs ?? 11;
    showMenu(SIZES.map((n) => ({ label: String(n), check: n === cur, act: () => { this.applyStyle({ fs: n === 11 ? undefined : n }); this.fitRows(); } })), r.left, r.bottom + 2);
  }
  private insertItems(): MenuItem[] {
    return [
      { label: 'Chèn Ô…', icon: 'cellIns', act: () => this.cellsDlg(true) },
      { label: 'Chèn Hàng Trang tính', icon: 'rowIns', act: () => this.insRows() },
      { label: 'Chèn Cột Trang tính', icon: 'colIns', act: () => this.insCols() },
      'sep',
      { label: 'Chèn Trang tính', icon: 'sheet', act: () => this.addSheet() },
    ];
  }
  private deleteItems(): MenuItem[] {
    return [
      { label: 'Xóa Ô…', icon: 'cellDel', act: () => this.cellsDlg(false) },
      { label: 'Xóa Hàng Trang tính', icon: 'rowDel', act: () => this.delRows() },
      { label: 'Xóa Cột Trang tính', icon: 'colDel', act: () => this.delCols() },
      'sep',
      { label: 'Xóa Trang tính', icon: 'trash', act: () => this.removeSheet(this.grid.sheet) },
    ];
  }
  private formatItems(): MenuItem[] {
    const s = this.grid.sheet;
    return [
      { label: 'Chiều cao Hàng…', icon: 'rowH', act: () => this.rowHeightDlg() },
      { label: 'Tự Khớp Chiều cao Hàng', act: () => this.autofitRows() },
      { label: 'Độ rộng Cột…', icon: 'colW', act: () => this.colWidthDlg() },
      { label: 'Tự Khớp Độ rộng Cột', icon: 'autofit', act: () => this.autofitCols() },
      { label: 'Độ rộng Mặc định…', act: () => this.later('Độ rộng Mặc định', '0.6') },
      'sep',
      { label: 'Ẩn & Bỏ ẩn', icon: 'eye', sub: [
        { label: 'Ẩn Hàng', hint: 'Ctrl+9', act: () => this.hideSel('r', true) },
        { label: 'Ẩn Cột', hint: 'Ctrl+0', act: () => this.hideSel('c', true) },
        { label: 'Ẩn Trang tính', act: () => this.hideSheet(s) },
        'sep',
        { label: 'Bỏ ẩn Hàng', hint: 'Ctrl+Shift+9', act: () => this.hideSel('r', false) },
        { label: 'Bỏ ẩn Cột', act: () => this.hideSel('c', false) },
        { label: 'Bỏ ẩn Trang tính…', act: () => this.unhideSheetDlg() },
      ] },
      'sep',
      { label: 'Đổi tên Trang tính', act: () => this.renameSheet(s) },
      { label: 'Di chuyển hoặc Sao chép Trang tính…', icon: 'move', act: () => this.moveCopyDlg(s) },
      { label: 'Màu Tab', icon: 'tabColor', sub: () => colorPalette((c) => { this.wb.setTabColor(s, c); this.renderSheets(); }, 'Không Màu') },
      'sep',
      { label: 'Bảo vệ Trang tính…', icon: 'protectSheet', act: () => this.later('Bảo vệ Trang tính', '0.9') },
      { label: 'Khóa Ô', icon: 'lock', act: () => this.later('Khóa Ô', '0.9') },
      { label: 'Định dạng Ô…', icon: 'cellFmt', hint: 'Ctrl+1', act: () => this.later('Định dạng Ô', '0.6') },
    ];
  }
  /** Hộp thoại Chèn… / Xóa… (dồn ô) như Excel */
  cellsDlg(insert: boolean) {
    const s = this.grid.sheet, g = this.grid.range;
    const opts = insert ? [['right', 'Dồn ô sang phải'], ['down', 'Dồn ô xuống'], ['row', 'Toàn bộ hàng'], ['col', 'Toàn bộ cột']] : [['left', 'Dồn ô sang trái'], ['up', 'Dồn ô lên'], ['row', 'Toàn bộ hàng'], ['col', 'Toàn bộ cột']];
    // như Excel: vùng chọn cao hơn rộng → mặc định dồn ngang, ngược lại dồn dọc
    const def = g.r2 - g.r1 > g.c2 - g.c1 ? (insert ? 'right' : 'left') : insert ? 'down' : 'up';
    const body = h('div', { class: 'stack' }, ...opts.map(([v, l]) => { const i = h('input', { type: 'radio', name: 'cells', value: v }) as HTMLInputElement; if (v === def) i.checked = true; return h('label', { class: 'chk' }, i, ' ' + l); }));
    openDialog(insert ? 'Chèn' : 'Xóa', body, [
      { label: 'Hủy', act: () => {} },
      { label: 'OK', kind: 'primary', act: () => {
        const v = (body.querySelector('input:checked') as HTMLInputElement).value;
        if (v === 'row') { if (insert) this.insRows(); else this.delRows(); }
        else if (v === 'col') { if (insert) this.insCols(); else this.delCols(); }
        else this.safe(() => { if (insert) this.wb.insertCells(s, g, v as 'down' | 'right'); else this.wb.deleteCells(s, g, v as 'up' | 'left'); });
        this.grid.invalidate(); this.syncSelection();
      } },
    ], () => this.grid.focusKeys());
  }
  private numDlg(title: string, label: string, value: string, apply: (n: number) => string | void) {
    const inp = h('input', { type: 'text', inputmode: 'decimal' }) as HTMLInputElement;
    inp.value = value;
    const err = h('p', { class: 'err-text', role: 'alert' });
    openDialog(title, h('div', { class: 'stack' }, field(label, inp), err), [
      { label: 'Hủy', act: () => {} },
      { label: 'OK', kind: 'primary', act: () => { const n = parseFloat(inp.value.replace(',', '.')); const e = isNaN(n) ? 'Hãy nhập một số.' : apply(n); if (e) { err.textContent = e; return false; } } },
    ], () => this.grid.focusKeys());
    setTimeout(() => inp.select(), 40);
  }
  private selRows() { const g = this.grid.range, u = this.wb.used(this.grid.sheet); return [g.r1, g.r2 >= MAX_ROWS - 1 ? Math.max(g.r1, u.rows - 1) : g.r2] as const; }
  private selCols() { const g = this.grid.range, u = this.wb.used(this.grid.sheet); return [g.c1, g.c2 >= MAX_COLS - 1 ? Math.max(g.c1, u.cols - 1) : g.c2] as const; }
  rowHeightDlg() {
    const s = this.grid.sheet, [a, b] = this.selRows();
    const pt = Math.round(this.wb.meta(s).rowH.get(this.grid.sel.ar) * 0.75 * 100) / 100;
    this.numDlg('Chiều cao Hàng', 'Chiều cao hàng (điểm):', String(pt).replace('.', ','), (n) => {
      if (n < 0 || n > 409) return 'Chiều cao hàng phải từ 0 đến 409.';
      const list: [number, number][] = [];
      for (let r = a; r <= Math.min(b, a + 5000); r++) list.push([r, Math.round(n / 0.75)]);
      this.safe(() => this.wb.setSizes(s, 'r', list));
      this.grid.layout();
    });
  }
  colWidthDlg() {
    const s = this.grid.sheet, [a, b] = this.selCols();
    const px = this.wb.meta(s).colW.get(this.grid.sel.ac);
    const w = Math.round(((px - 5) / 7) * 100) / 100;
    this.numDlg('Độ rộng Cột', 'Độ rộng cột (ký tự):', String(Math.max(0, w)).replace('.', ','), (n) => {
      if (n < 0 || n > 255) return 'Độ rộng cột phải từ 0 đến 255.';
      const list: [number, number][] = [];
      for (let c = a; c <= Math.min(b, a + 2000); c++) list.push([c, n === 0 ? 0 : Math.round(n * 7 + 5)]);
      this.safe(() => this.wb.setSizes(s, 'c', list));
      this.grid.layout();
    });
  }
  autofitRows() { this.fitRows(true); }
  autofitCols() {
    const [a, b] = this.selCols(), list: [number, number][] = [];
    for (let c = a; c <= Math.min(b, a + 500); c++) list.push([c, this.grid.fitWidth(c)]);
    this.safe(() => this.wb.setSizes(this.grid.sheet, 'c', list)); this.grid.layout();
  }
  hideSel(ax: 'r' | 'c', hide: boolean) {
    const s = this.grid.sheet;
    this.safe(() => {
      if (ax === 'r') { const [a, b] = this.selRows(); this.wb.hideRows(s, a, b, hide); }
      else { const [a, b] = this.selCols(); this.wb.hideCols(s, a, b, hide); }
    });
    this.grid.layout(); this.syncSelection();
  }
  setIter(on: boolean) {
    this.wb.setIterate({ on }); this.grid.invalidate(); this.ribbon.refresh();
    this.toast(on ? `Đã bật tính lặp (tối đa ${this.wb.iterate.count} lần, sai số ${String(this.wb.iterate.delta).replace('.', ',')}) — tham chiếu vòng được tính như Excel.` : 'Đã tắt tính lặp — ô có tham chiếu vòng hiện #CYCLE!.');
  }
  clearAll() { const s = this.grid.sheet, g = this.grid.range; this.safe(() => { this.wb.clear(s, g); this.wb.setStyle(s, g, { b: undefined, i: undefined, u: undefined, s: undefined, fg: undefined, bg: undefined, al: undefined, fmt: undefined, fn: undefined, fs: undefined, wr: undefined, va: undefined, bd: undefined, ind: undefined }); }); }
  clearFormats() { const s = this.grid.sheet, g = this.grid.range; this.safe(() => this.wb.setStyle(s, g, { b: undefined, i: undefined, u: undefined, s: undefined, fg: undefined, bg: undefined, al: undefined, fmt: undefined, fn: undefined, fs: undefined, wr: undefined, va: undefined, bd: undefined, ind: undefined })); }
  pageSetup(tab: 'page' | 'margins' | 'hf' | 'sheet') { pageSetupDialog(this, tab, () => { this.backstage.refreshPrint(); this.grid.invalidate(); this.ribbon.refresh(); }); }
  setView(v: 'normal' | 'break' | 'layout') {
    if (v === 'layout') { this.later('Dạng xem Bố trí Trang', '0.8'); return; }
    this.grid.view = v;
    this.grid.invalidate();
    document.querySelectorAll<HTMLElement>('#views [data-view]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.view === v)));
    this.ribbon.refresh();
    if (v === 'break' && !this.wb.used(this.grid.sheet).rows) this.toast('Trang tính trống — chưa có trang in nào.');
  }
  zoomDlg() {
    const cur = Math.round(this.grid.zoom * 100);
    const opts: [string, string][] = [['200', '200%'], ['100', '100%'], ['75', '75%'], ['50', '50%'], ['25', '25%'], ['fit', 'Vừa với vùng chọn'], ['custom', 'Tùy chỉnh:']];
    const custom = h('input', { type: 'text', inputmode: 'numeric', 'aria-label': 'Phần trăm tùy chỉnh' }) as HTMLInputElement;
    custom.value = cur + '%';
    const body = h('div', { class: 'stack' }, ...opts.map(([v, l]) => { const i = h('input', { type: 'radio', name: 'zoom', value: v }) as HTMLInputElement; if (v === String(cur) || (v === 'custom' && ![200, 100, 75, 50, 25].includes(cur))) i.checked = true; return h('label', { class: 'chk' }, i, ' ' + l, ...(v === 'custom' ? [custom] : [])); }));
    custom.onfocus = () => { (body.querySelector('input[value=custom]') as HTMLInputElement).checked = true; };
    openDialog('Thu phóng', body, [
      { label: 'Hủy', act: () => {} },
      { label: 'OK', kind: 'primary', act: () => {
        const v = (body.querySelector('input[name=zoom]:checked') as HTMLInputElement).value;
        if (v === 'fit') this.zoomToSel();
        else { const n = v === 'custom' ? parseInt(custom.value) : +v; if (!(n >= 10 && n <= 400)) { this.toast('Thu phóng phải từ 10% đến 400%.', 'err'); return false; } this.grid.setZoom(n / 100); }
      } },
    ], () => this.grid.focusKeys());
  }
  zoomToSel() {
    const g = this.grid.range, m = this.wb.meta(this.grid.sheet);
    const w = m.colW.offset(g.c2 + 1) - m.colW.offset(g.c1), hh = m.rowH.offset(g.r2 + 1) - m.rowH.offset(g.r1);
    const r = this.grid.scroller.getBoundingClientRect();
    const z = Math.max(0.1, Math.min(4, Math.min((r.width - 60) / Math.max(1, w), (r.height - 40) / Math.max(1, hh))));
    this.grid.setZoom(z);
    this.grid.ensureVisible(g.r1, g.c1);
  }
  bookStats() {
    let cells = 0, formulas = 0;
    for (const id of this.wb.sheetIds()) { const u = this.wb.used(id); for (let r = 0; r < u.rows; r++) for (let c = 0; c < u.cols; c++) { const t = this.wb.rawText(id, r, c); if (t) { cells++; if (t.startsWith('=')) formulas++; } } }
    const body = h('dl', { class: 'bs-props' }, h('dt', {}, 'Trang tính'), h('dd', {}, String(this.wb.sheetIds().length)), h('dt', {}, 'Ô có dữ liệu'), h('dd', {}, cells.toLocaleString('vi-VN')), h('dt', {}, 'Công thức'), h('dd', {}, formulas.toLocaleString('vi-VN')), h('dt', {}, 'Biểu đồ'), h('dd', {}, String(this.wb.sheetIds().reduce((a, id) => a + this.wb.meta(id).charts.length, 0))), h('dt', {}, 'PivotTable'), h('dd', {}, String(this.wb.sheetIds().reduce((a, id) => a + this.wb.meta(id).pivots.length, 0))));
    openDialog('Thống kê Sổ làm việc', body, [{ label: 'OK', kind: 'primary', act: () => {} }], () => this.grid.focusKeys());
  }
  helpDlg(keys = false) {
    const rows: [string, string][] = [['Ctrl+S', 'Lưu'], ['Ctrl+O', 'Mở'], ['Ctrl+N', 'Sổ làm việc mới'], ['Ctrl+P', 'In (xem trước)'], ['Ctrl+Z / Ctrl+Y', 'Hoàn tác / Làm lại'], ['Ctrl+B, I, U', 'Đậm, Nghiêng, Gạch chân'], ['Alt hoặc F10', 'Hiện KeyTips'], ['Alt+Q', 'Tìm kiếm lệnh'], ['Alt+=', 'Tổng Tự động'], ['F2', 'Sửa ô'], ['Alt+Enter', 'Xuống dòng trong ô'], ['Ctrl+Shift+L', 'Bật/tắt Lọc'], ['Ctrl+F1', 'Thu gọn/hiện Ribbon'], ['Shift+F11', 'Trang tính mới'], ['Ctrl+PageDown/PageUp', 'Sang trang tính sau/trước'], ['Ctrl+Shift+U', 'Mở rộng thanh công thức'], ['Shift+F3', 'Chèn Hàm'], ['F9', 'Tính lại']];
    const body = h('div', { class: 'stack' }, ...(keys ? [] : [h('p', {}, `ExcelLike ${VERSION} — bảng tính kiểu Excel 365 chạy trên Windows và trình duyệt. Công thức tiếng Việt dùng dấu chấm phẩy (;) ngăn tham số, dấu phẩy (,) là dấu thập phân.`)]), h('dl', { class: 'bs-props' }, ...rows.flatMap(([k, v]) => [h('dt', {}, k), h('dd', {}, v)])));
    openDialog(keys ? 'Phím tắt' : 'Trợ giúp', body, [{ label: 'Đóng', kind: 'primary', act: () => {} }], () => this.grid.focusKeys());
  }
  printDlg() { this.backstage.open('print'); }
  quickPrint() { this.backstage.quickPrint(); }
  stepSheet(d: number) {
    const ids = this.wb.visibleSheetIds(), i = ids.indexOf(this.grid.sheet);
    const n = ids[Math.max(0, Math.min(ids.length - 1, i + d))];
    if (n !== undefined && n !== this.grid.sheet) { this.grid.setSheet(n); this.renderSheets(); this.syncSelection(); }
  }

  // ---------- điện thoại ----------
  private bindMobile() {
    document.querySelectorAll<HTMLElement>('#mob-cmd [data-m]').forEach((b) => {
      b.onclick = () => {
        const m = b.dataset.m, r = b.getBoundingClientRect();
        if (m === 'edit') this.ed.begin(null, 'edit', false, 'bar');
        if (m === 'fn') this.funcPicker();
        if (m === 'sort') { this.anchor = { x: r.left, y: r.top - 8 }; this.sortFilterMenu(); }
        if (m === 'chart') this.insertChart();
        if (m === 'more') this.backstage.open('home');
      };
    });
  }

  /** Hướng di chuyển sau khi nhấn Enter (Tệp ▸ Tùy chọn ▸ Nâng cao) */
  enterMove(shift: boolean): [number, number] {
    const d = lsGet('xl-enter') ?? 'down';
    const v: [number, number] = d === 'right' ? [0, 1] : d === 'left' ? [0, -1] : d === 'up' ? [-1, 0] : d === 'none' ? [0, 0] : [1, 0];
    return shift ? [-v[0], -v[1]] : v;
  }
  // ---------- thông báo ----------
  toast(msg: string, kind: 'ok' | 'err' | '' = '', action?: { label: string; act: () => void }) {
    const t = document.createElement('div');
    t.className = 'toast ' + kind;
    t.setAttribute('role', kind === 'err' ? 'alert' : 'status');
    t.innerHTML = `<span></span>`;
    t.firstElementChild!.textContent = msg;
    if (action) {
      const b = document.createElement('button');
      b.textContent = action.label;
      b.onclick = () => { action.act(); t.remove(); };
      t.appendChild(b);
    }
    $('#toasts').appendChild(t);
    setTimeout(() => t.classList.add('out'), action ? 5500 : 3200);
    setTimeout(() => t.remove(), action ? 6000 : 3700);
  }
}

// ======================= Bộ sửa ô =======================
class CellEditor implements EditHost {
  private ta: HTMLTextAreaElement;
  private fx: HTMLTextAreaElement;
  private ac: HTMLDivElement;
  private editing = false;
  private mode: 'enter' | 'edit' = 'enter';
  private at = { s: 0, r: 0, c: 0 };
  private orig = '';
  private origin: 'cell' | 'bar' = 'cell';
  private pref: { start: number; len: number; r: number; c: number } | null = null;
  private acItems: string[] = [];
  private acIdx = 0;
  private composing = false;

  constructor(private app: App) {
    const host = $('#gridwrap');
    this.ta = document.createElement('textarea');
    this.ta.className = 'cell-editor idle';
    this.ta.setAttribute('aria-label', 'Nội dung ô');
    this.ta.spellcheck = false;
    this.ta.autocapitalize = 'off';
    this.ta.setAttribute('autocomplete', 'off');
    if (isTouch()) this.ta.inputMode = 'none';
    host.appendChild(this.ta);
    this.ac = document.createElement('div');
    this.ac.className = 'ac-list';
    this.ac.hidden = true;
    this.ac.setAttribute('role', 'listbox');
    host.appendChild(this.ac);
    this.fx = $('#fx-input');
    this.ta.addEventListener('keydown', (e) => this.onKey(e, 'cell'));
    this.fx.addEventListener('keydown', (e) => this.onKey(e, 'bar'));
    this.ta.addEventListener('input', () => this.onInput('cell'));
    this.fx.addEventListener('input', () => this.onInput('bar'));
    this.ta.addEventListener('compositionstart', () => { this.composing = true; if (!this.editing) this.begin('', 'enter', true); });
    this.ta.addEventListener('compositionend', () => { this.composing = false; });
    this.fx.addEventListener('focus', () => { if (!this.editing) this.begin(null, 'edit', false, 'bar'); else { this.origin = 'bar'; this.ta.classList.add('mirror'); this.ta.readOnly = true; } });
    this.ac.addEventListener('mousedown', (e) => e.preventDefault());
    this.ac.addEventListener('click', (e) => { const b = (e.target as HTMLElement).closest('button'); if (b) { this.acIdx = +b.dataset.i!; this.acceptAc(); } });
    $('#fx-ok').onclick = () => this.commit(0, 0);
    $('#fx-cancel').onclick = () => this.cancel();
    window.addEventListener('resize', () => this.place());
    app.grid.scroller.addEventListener('scroll', () => this.place());
  }
  isEditing() { return this.editing; }
  refocus() { this.place(); this.el.focus({ preventScroll: true }); }
  label(t: string) { if (!this.editing) this.ta.setAttribute('aria-label', 'Ô ' + t); }
  private get grid() { return this.app.grid; }
  private get wb() { return this.app.wb; }
  private get el() { return this.origin === 'bar' ? this.fx : this.ta; }

  begin(initial: string | null, mode: 'enter' | 'edit', keepValue = false, origin: 'cell' | 'bar' = isTouch() ? 'bar' : 'cell') {
    if (this.editing) return;
    const { ar, ac } = this.grid.sel;
    this.at = { s: this.grid.sheet, r: ar, c: ac };
    this.orig = this.wb.rawText(this.at.s, ar, ac);
    this.editing = true;
    this.mode = mode;
    this.origin = origin;
    this.pref = null;
    const text = initial === null ? this.orig : initial;
    if (!keepValue) this.ta.value = text;
    this.fx.value = keepValue ? this.ta.value : text;
    this.ta.classList.remove('idle');
    document.body.classList.add('editing');
    $('#fx-ok').removeAttribute('disabled'); $('#fx-cancel').removeAttribute('disabled');
    this.app.setMode(mode === 'enter' ? 'Nhập' : 'Sửa');
    this.place();
    this.ta.classList.toggle('mirror', origin === 'bar');
    if (origin === 'bar') {
      this.ta.readOnly = true;
      if (document.activeElement !== this.fx) this.fx.focus();
      this.fx.setSelectionRange(this.fx.value.length, this.fx.value.length);
    } else {
      this.ta.readOnly = false;
      this.ta.focus({ preventScroll: true });
      if (!keepValue) this.ta.setSelectionRange(this.ta.value.length, this.ta.value.length);
    }
    this.updateAc();
  }
  commit(dr: number, dc: number): boolean {
    if (!this.editing) { if (dr || dc) this.grid.moveActive(dr, dc); return true; }
    const v = this.el.value;
    if (v !== this.orig) {
      const bad = this.wb.validate(this.at.s, this.at.r, this.at.c, v);
      if (bad) {
        const rule = this.wb.dvAt(this.at.s, this.at.r, this.at.c);
        if (!rule?.errorStyle || rule.errorStyle === 'stop') { this.app.toast(bad, 'err'); this.el.focus(); this.el.select(); return false; }
        this.app.toast(bad, 'err');
      }
    }
    this.finish();
    if (v !== this.orig) {
      try { this.wb.setInput(this.at.s, this.at.r, this.at.c, v); } catch (e) { this.app.toast((e as Error).message, 'err'); }
    }
    this.backToCell();
    if (dr || dc) this.grid.moveActive(dr, dc); else this.app.syncSelection();
    this.grid.invalidate();
    return true;
  }
  cancel() { if (!this.editing) return; this.finish(); this.backToCell(); this.app.syncSelection(); }
  /** Nếu đã chuyển sang sheet khác để trỏ ô, quay lại sheet và ô đang sửa */
  private backToCell() {
    if (this.grid.sheet === this.at.s) return;
    this.grid.setSheet(this.at.s);
    this.grid.select(this.at.r, this.at.c);
    this.app.renderSheets();
  }
  private finish() {
    this.editing = false;
    this.pref = null;
    this.closeAc();
    this.ta.classList.add('idle');
    this.ta.classList.remove('mirror');
    this.ta.readOnly = false;
    this.ta.value = '';
    document.body.classList.remove('editing');
    $('#fx-ok').setAttribute('disabled', ''); $('#fx-cancel').setAttribute('disabled', '');
    this.app.setMode('Sẵn sàng');
    this.place();
    if (!isTouch() || document.activeElement === this.fx) this.grid.focusKeys();
    if (document.activeElement === this.fx) this.fx.blur();
  }
  place() {
    const g = this.grid;
    const { r, c } = this.editing ? this.at : { r: g.sel.ar, c: g.sel.ac };
    const rc = g.cellRectM(r, c);
    const s = this.ta.style;
    this.ta.classList.toggle('offsheet', this.editing && g.sheet !== this.at.s);
    if (!this.editing) { s.left = Math.max(0, Math.min(rc.x, g.width - 4)) + 'px'; s.top = Math.max(0, Math.min(rc.y, g.height - 4)) + 'px'; s.width = '2px'; s.height = '2px'; return; }
    const st = this.wb.style(this.at.s, r, c);
    s.font = `${st.b ? 700 : 400} ${ptPx(st.fs) * g.zoom}px ${st.fn ? `"${st.fn}", ` : ''}${getComputedStyle(document.documentElement).getPropertyValue('--font-cell')}`;
    s.fontWeight = st.b ? '600' : '400';
    s.fontStyle = st.i ? 'italic' : 'normal';
    const ctx = document.createElement('canvas').getContext('2d')!;
    ctx.font = s.font;
    const lines = this.ta.value.split('\n');
    const tw = Math.max(...lines.map((l) => ctx.measureText(l).width)) + 16 * g.zoom;
    const w = Math.min(Math.max(rc.w + 1, tw), g.width - rc.x - 2);
    s.left = rc.x + 'px'; s.top = rc.y + 'px';
    s.width = Math.max(rc.w + 1, w) + 'px';
    s.height = Math.max(rc.h + 1, lines.length * 18 * g.zoom + 6) + 'px';
    this.ac.style.left = rc.x + 'px';
    this.ac.style.top = (rc.y + Math.max(rc.h, lines.length * 18 * g.zoom + 6) + 4) + 'px';
  }

  // ---------- chế độ trỏ ô khi gõ công thức ----------
  pointMode(): boolean {
    if (!this.editing) return false;
    const el = this.el;
    if (!el.value.startsWith('=')) return false;
    const pos = el.selectionStart ?? el.value.length;
    if (this.pref && pos === this.pref.start + this.pref.len) return true;
    return /[=(;+\-*/^&<>:]\s*$/.test(el.value.slice(0, pos));
  }
  insertRef(ref: string, replaceLast: boolean) {
    const el = this.el;
    const sheetPrefix = this.grid.sheet !== this.at.s ? quoteSheet(this.wb.sheetName(this.grid.sheet)) + '!' : '';
    const full = sheetPrefix + ref;
    let start: number;
    if ((replaceLast || (this.pref && el.selectionStart === this.pref.start + this.pref.len)) && this.pref) {
      start = this.pref.start;
      el.value = el.value.slice(0, start) + full + el.value.slice(start + this.pref.len);
    } else {
      start = el.selectionStart ?? el.value.length;
      el.value = el.value.slice(0, start) + full + el.value.slice(el.selectionEnd ?? start);
    }
    const m = /([A-Z]+)(\d+)$/.exec(ref.split(':').pop()!)!;
    this.pref = { start, len: full.length, r: +m[2] - 1, c: colIdx(m[1]) };
    el.setSelectionRange(start + full.length, start + full.length);
    this.mirror(el === this.ta ? 'cell' : 'bar');
    this.app.setMode('Trỏ');
    this.place();
    el.focus({ preventScroll: true });
  }
  private movePoint(dr: number, dc: number) {
    const base = this.pref ?? { r: this.at.r, c: this.at.c };
    const r = Math.max(0, Math.min(MAX_ROWS - 1, base.r + dr)), c = Math.max(0, Math.min(MAX_COLS - 1, base.c + dc));
    this.grid.ensureVisible(r, c);
    this.insertRef(toA1(r, c), !!this.pref);
  }

  // ---------- gợi ý hàm ----------
  private updateAc() {
    const el = this.el;
    const pos = el.selectionStart ?? el.value.length;
    const before = el.value.slice(0, pos);
    const m = el.value.startsWith('=') ? /(?:^=|[=(;,+\-*/^&<>:\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(before) : null;
    if (!m) return this.closeAc();
    const q = m[1].toUpperCase();
    const names = this.wb.hf.getRegisteredFunctionNames();
    const hits = names.filter((n) => n.startsWith(q)).sort((a, b) => (FUNC_INFO[b] ? 1 : 0) - (FUNC_INFO[a] ? 1 : 0) || a.length - b.length || a.localeCompare(b)).slice(0, 8);
    if (!hits.length || (hits.length === 1 && hits[0] === q)) return this.closeAc();
    this.acItems = hits;
    this.acIdx = 0;
    this.renderAc();
  }
  private renderAc() {
    this.ac.innerHTML = this.acItems.map((n, i) => {
      const inf = FUNC_INFO[n];
      return `<button role="option" data-i="${i}" aria-selected="${i === this.acIdx}"><b>${n}</b>${inf ? `<span>${inf[2]}</span>` : ''}</button>`;
    }).join('');
    this.ac.hidden = false;
    if (this.origin === 'bar') {
      const r = this.fx.getBoundingClientRect(), h = $('#gridwrap').getBoundingClientRect();
      this.ac.style.left = Math.max(4, r.left - h.left) + 'px';
      this.ac.style.top = '4px';
    }
  }
  private closeAc() { this.ac.hidden = true; this.acItems = []; }
  private acceptAc() {
    const el = this.el;
    const name = this.acItems[this.acIdx];
    const pos = el.selectionStart ?? el.value.length;
    const before = el.value.slice(0, pos).replace(/[A-Za-z][A-Za-z0-9.]*$/, '');
    el.value = before + name + '(' + el.value.slice(pos);
    const p = before.length + name.length + 1;
    el.setSelectionRange(p, p);
    this.closeAc();
    this.mirror(el === this.ta ? 'cell' : 'bar');
    this.showHint(name);
    this.place();
  }
  private showHint(name: string) {
    const inf = FUNC_INFO[name];
    if (inf) this.app.setMode(inf[1]);
  }

  private mirror(from: 'cell' | 'bar') { if (from === 'cell') this.fx.value = this.ta.value; else this.ta.value = this.fx.value; }
  private onInput(from: 'cell' | 'bar') {
    if (!this.editing) {
      if (from === 'cell') {
        const v = this.ta.value;
        if (!v.trim()) { this.ta.value = ''; return; }
        this.begin(v, 'enter', true);
      }
      return;
    }
    this.pref = null;
    this.mirror(from);
    if (this.mode === 'enter') this.app.setMode('Nhập');
    this.place();
    this.updateAc();
  }

  private onKey(e: KeyboardEvent, from: 'cell' | 'bar') {
    if (this.composing || e.isComposing) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (!this.editing) {
      if (from === 'bar') return;
      this.navKey(e, ctrl);
      return;
    }
    if (!this.ac.hidden) {
      if (k === 'ArrowDown' || k === 'ArrowUp') { e.preventDefault(); this.acIdx = (this.acIdx + (k === 'ArrowDown' ? 1 : -1) + this.acItems.length) % this.acItems.length; this.renderAc(); return; }
      if (k === 'Tab') { e.preventDefault(); this.acceptAc(); return; }
      if (k === 'Escape') { e.preventDefault(); this.closeAc(); return; }
    }
    if (k === 'Enter' && e.altKey) {
      e.preventDefault();
      const el = this.el, p = el.selectionStart ?? el.value.length;
      el.value = el.value.slice(0, p) + '\n' + el.value.slice(el.selectionEnd ?? p);
      el.setSelectionRange(p + 1, p + 1);
      this.onInput(from);
      return;
    }
    if (k === 'Enter') { e.preventDefault(); this.commit(...this.app.enterMove(e.shiftKey)); return; }
    if (k === 'Tab') { e.preventDefault(); this.commit(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'Escape') { e.preventDefault(); this.cancel(); return; }
    if (k === 'F2') { e.preventDefault(); this.mode = this.mode === 'enter' ? 'edit' : 'enter'; this.app.setMode(this.mode === 'enter' ? 'Nhập' : 'Sửa'); return; }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k] && this.mode === 'enter' && from === 'cell') {
      e.preventDefault();
      if (this.pointMode()) this.movePoint(...arrows[k]);
      else this.commit(...arrows[k]);
    }
  }

  private navKey(e: KeyboardEvent, ctrl: boolean) {
    const g = this.grid, app = this.app, k = e.key;
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const done = () => e.preventDefault();
    if (arrows[k]) {
      done();
      const [dr, dc] = arrows[k];
      if (ctrl) g.jump(dr, dc, e.shiftKey);
      else g.step(dr, dc, e.shiftKey);
      return;
    }
    if (ctrl) {
      const kk = k.toLowerCase();
      if (kk === 'z') { done(); e.shiftKey ? app.redo() : app.undo(); return; }
      if (kk === 'y') { done(); app.redo(); return; }
      if (kk === 'b') { done(); app.toggle('b'); return; }
      if (kk === 'i') { done(); app.toggle('i'); return; }
      if (kk === 'u') { done(); app.toggle('u'); return; }
      if (kk === 'a') { done(); g.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }); return; }
      if (kk === 'd') { done(); app.fillDir('down'); return; }
      if (kk === 'r') { done(); app.fillDir('right'); return; }
      if (k === 'Home') { done(); g.select(0, 0); return; }
      if (k === ';') { done(); const t = new Date(); this.begin(`${pad(t.getDate())}/${pad(t.getMonth() + 1)}/${t.getFullYear()}`, 'enter'); return; }
      if (kk === 'c' || kk === 'x' || kk === 'v') return; // để sự kiện copy/cut/paste xử lý
      return;
    }
    switch (k) {
      case 'Enter': { done(); const [dr, dc] = app.enterMove(e.shiftKey); if (dr || dc) g.moveActive(dr, dc); return; }
      case 'Tab': done(); g.moveActive(0, e.shiftKey ? -1 : 1); return;
      case 'Home': done(); g.select(g.sel.ar, 0); return;
      case 'PageDown': case 'PageUp': {
        done();
        const n = Math.max(1, g.visibleRows().length - 2) * (k === 'PageDown' ? 1 : -1);
        g.select(g.sel.ar + n, g.sel.ac);
        return;
      }
      case 'Delete': done(); app.clearSel(); return;
      case 'Backspace': done(); this.begin('', 'enter'); return;
      case 'F2': done(); this.begin(null, 'edit'); return;
      case 'F9': done(); app.recalc(); return;
      case 'Escape': done(); if (app.wb.clip) { app.wb.clip = null; g.invalidate(); } return;
    }
    if (k === '=' && e.altKey) { done(); app.autoSum(); return; }
    if (k === 'F3' && e.shiftKey) { done(); app.funcPicker(); return; }
    if (k.length === 1 && !e.altKey) {
      // gõ ký tự: bắt đầu nhập, để ký tự tự rơi vào ô
      this.begin('', 'enter', true);
      this.ta.value = '';
    }
  }
}

function pad(n: number) { return String(n).padStart(2, '0'); }
function colIdx(s: string) { let c = 0; for (const ch of s) c = c * 26 + ch.charCodeAt(0) - 64; return c - 1; }
function normalizeNL(s: string) { return s.replace(/\r\n/g, '\n').replace(/\n$/, ''); }
