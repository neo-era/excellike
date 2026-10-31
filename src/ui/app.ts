import { MAX_COLS, MAX_ROWS, parseRange, rangeToA1, toA1, type Range } from '../core/address';
import { quoteSheet, shiftFormula } from '../core/formula';
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
import { Ribbon, KeyTips, bindRibbonKeys, allItems, fold, type Tab, type Btn, type KtTarget, type Item, type Group } from './ribbon';
import { closeMenu, openPopover, showMenu, menuOpen, type MenuItem } from './menu';
import { Backstage } from './backstage';
import { groupSel, ungroupSel, clearOutline, showDetail, subtotalDialog, outlineClick } from './outline-ui';
import { removeDupDialog, textToColumnsDialog } from './datatools-ui';
import { createTableDialog, tableStylePicker, tableDesignTab, totalFnMenu } from './table-ui';
import { NoteLayer } from './notes-ui';
import { nameManager, nameDialog, createFromSelection, pasteNameDialog, selRef } from './names-ui';
import { pageSetupDialog } from './pagesetup';
import { normPrint, storePrint } from '../core/paginate';

export { closeMenu, openPopover, showMenu };
const $ = <T extends HTMLElement = HTMLElement>(sel: string, root: ParentNode = document) => root.querySelector(sel) as T;
const isTouch = () => matchMedia('(pointer: coarse)').matches;
const COMPACT = '(max-width: 760px), (max-height: 520px)';
import { VERSION } from './version';
import { applyTemplate } from './templates';
import { tr } from './i18n';
import { formatValue, NF_SHORTCUT } from '../core/numfmt';
import { comboOf, keyFor } from './keymap';
import { formatCellsDialog } from './fmtcells';
import { findDialog, findAgain, gotoDialog, gotoSpecialDialog, pasteSpecialDialog, hyperlinkDialog, seriesDialog, funcArgsDialog, argNames, quickAnalysis, keysDialog, pushClip, toggleClipboardPane, closeFind } from './dlg06';
import { goSpecial } from '../core/find';
import { flashFill } from '../core/flashfill';
import { cycleRef } from '../core/formula';
import type { SeriesOpts } from '../core/series';
import type { PasteOpts } from '../core/workbook';
const pad2 = (n: number) => String(n).padStart(2, '0');
const needsApos = (v: string) => /^[=+\-'\d.,]/.test(v) || /^(true|false)$/i.test(v);

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
export const FONTS = ['Aptos Narrow', 'Aptos', 'Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana', 'Roboto', 'Be Vietnam Pro'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export const DEFAULT_FONT = 'Aptos Narrow';
/** Kiểu Ô dựng sẵn (Home ▸ Kiểu Ô) */
/** CSS xem trước một Kiểu Ô trong bộ sưu tập */
function styleCss(st: Style, bk?: BorderKind, bc?: string) {
  const o: string[] = [];
  if (st.bg) o.push(`background:${st.bg}`);
  if (st.fg) o.push(`color:${st.fg}`);
  if (st.b) o.push('font-weight:700');
  if (bk === 'all') o.push(`border:1px solid ${bc}`);
  else if (bk === 'bottom') o.push(`border-bottom:2px solid ${bc}`);
  return o.join(';');
}
export const laterNote = (v: string) => v.startsWith('ngoài') ? v : `có ở bản ${v}`;
const RECENT_DEF = ['SUM', 'AVERAGE', 'IF', 'VLOOKUP', 'COUNT', 'MAX', 'MIN', 'ROUND', 'SUMIF', 'IFERROR'];
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
  notes!: NoteLayer;
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
    this.grid.onOutline = (k, L, at) => outlineClick(this, k, L, at);
    this.grid.onContext = (x, y, zone) => this.contextMenu(x, y, zone);
    this.notes = new NoteLayer(this);
    this.charts = new ChartLayer($('#gridwrap'), this.wb, this.grid, (id) => this.editChart(id), (m) => this.toast(m));
    this.grid.onDraw = () => this.charts.sync();
    this.pivot = new PivotPanel($('#gridwrap'), this.wb, (m, k) => this.toast(m, k), () => this.grid.invalidate());
    $('#gridwrap').addEventListener('grid-error', (e) => this.toast((e as CustomEvent).detail, 'err'));
    $('#gridwrap').addEventListener('filter-btn', (e) => { const d = (e as CustomEvent).detail; this.filterPopover(d.col, d.x, d.y); });
    $('#gridwrap').addEventListener('dv-btn', (e) => { const d = (e as CustomEvent).detail; if (!totalFnMenu(this, d.x, d.y)) this.dvPick(d.x, d.y); });
    // 0.6: chổi định dạng áp khi nhả chuột; nhấp nhanh vào ô có liên kết thì mở liên kết (như Excel)
    let down = { t: 0, r: -1, c: -1 };
    $('#gridwrap').addEventListener('cell-down', (e) => { const d = (e as CustomEvent).detail; down = { t: performance.now(), r: d.r, c: d.c }; this.tabStart = null; if (d.ctrl) down.t = 0; });
    $('#gridwrap').addEventListener('sel-end', () => {
      if (this.painter) { this.painterApply(); return; }
      const g = this.grid;
      if (down.t && performance.now() - down.t < 450 && g.isSingle() && !g.areas.length && g.sel.ar === down.r && g.sel.ac === down.c) this.followLink(down.r, down.c);
    });
    const closed = sessionStorage.getItem('xl-warn-closed');
    if (closed !== null) { sessionStorage.removeItem('xl-warn-closed'); setTimeout(() => this.toast(`Đã đóng “${closed}”. Mở lại trong Tệp ▸ Mở ▸ Gần đây.`), 300); }
    const tpl = sessionStorage.getItem('xl-template');
    if (tpl) {
      sessionStorage.removeItem('xl-template');
      // chỉ áp vào sổ trống vừa tạo (nếu không tạo được sổ mới thì không ghi đè sổ cũ)
      const u0 = this.wb.used(this.grid.sheet);
      if (u0.rows === 0 && this.wb.sheetIds().length === 1) try { const n = applyTemplate(this.wb, this.grid.sheet, tpl); if (n) { this.wb.clearHistory(); this.updateTitle(); this.grid.layout(); this.scheduleSave(); } } catch (e) { this.toast('Không tạo được từ mẫu: ' + (e as Error).message, 'err'); }
    }
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
      if (e.key === 'Escape' && this.ribbonPeek && !document.querySelector('.dlg-back, .menu')) { this.ribbonPeek = false; this.applyRibbon(); }
      void ctrl; void k;
      this.keyGlobal(e);
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
    const grp = this.group.length > 1 ? ' [Nhóm]' : '';
    $('#doc-title').textContent = this.wb.name + grp;
    document.title = this.wb.name + grp + ' - ExcelLike';
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
      b.querySelector('small')!.textContent = x.later ? laterNote(x.later) : x.path;
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
  later(label: string, ver: string) { this.toast(ver.startsWith('ngoài') ? `“${label}”: ${ver}.` : `“${label}” sẽ có ở bản ${ver}.`); }
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
    cm.id = 'btn-comments';
    cm.innerHTML = ic('comment', 15) + '<span>Chú thích</span>';
    cm.title = 'Chú thích — có ở bản 0.7';
    cm.onclick = () => this.notes.togglePane();
    nav.appendChild(cm);
    const sh = document.createElement('button');
    sh.className = 'pill share';
    sh.id = 'btn-share';
    sh.innerHTML = ic('share', 15) + '<span>Chia sẻ</span>';
    sh.title = 'Chia sẻ — cần OneDrive/SharePoint (ngoài phạm vi). Dùng Tệp ▸ Xuất để gửi tệp .xlsx hoặc PDF.';
    sh.onclick = () => this.backstage.open('share');
    nav.appendChild(sh);
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
    const HIDE_FIRST = new Set(['addins', 'controls', 'dtypes', 'ink', 'sheetview', 'macros', 'lang', 'changes', 'perf']);
    const SHRINK: Record<string, number> = { macros: 0.3, views: 0.7, zoom: 0.75, show: 0.8, addins: 0.5, dtypes: 0.5, controls: 0.5, ink: 0.5, lang: 0.5, changes: 0.6, perf: 0.7, sheetview: 0.5, styles: 1, cells: 2, edit: 3, number: 4, align: 5, clip: 6 };
    const G = (id: string, label: string, icon: string, items: Item[], launch?: Group['launch']): Group => ({ id, label, icon, items, launch, shrink: SHRINK[id], hideFirst: HIDE_FIRST.has(id) });
    const ps = () => normPrint(this.wb.meta(s()).print);
    const setPs = (p: Partial<ReturnType<typeof normPrint>>) => this.safe(() => { this.wb.setPrint(s(), storePrint({ ...ps(), ...p })); });
    const pst = (o: PasteOpts) => () => this.pasteSpecialRun(o);
    const home: Tab = { id: 'home', label: 'Trang đầu', kt: 'H', groups: [
      G('clip', 'Bảng tạm', 'paste', [
        B('big', 'paste', 'paste', 'Dán', { kt: 'V', kbd: 'Ctrl+V', split: true, act: () => this.pasteBtn(), drop: this.menuAt(() => [
          { label: 'Dán', icon: 'paste', hint: 'Ctrl+V', act: () => this.pasteBtn() },
          { label: 'Công thức', icon: 'pasteFormula', act: pst({ what: 'formulas' }) },
          { label: 'Công thức và Định dạng Số', icon: 'pasteFormula', act: pst({ what: 'formulasNum' }) },
          { label: 'Không có Viền', icon: 'paste', act: pst({ what: 'noBorders' }) },
          { label: 'Giữ Độ rộng Cột Nguồn', icon: 'paste', act: () => { this.pasteSpecialRun({ what: 'all' }); this.pasteSpecialRun({ what: 'colWidths' }); } },
          { label: 'Chuyển vị', icon: 'transpose', act: pst({ what: 'all', transpose: true }) },
          'sep',
          { label: 'Giá trị', icon: 'pasteValues', hint: 'Ctrl+Shift+V', act: pst({ what: 'values' }) },
          { label: 'Giá trị và Định dạng Số', icon: 'pasteValues', act: pst({ what: 'valuesNum' }) },
          { label: 'Định dạng', icon: 'pasteFmt', act: pst({ what: 'formats' }) },
          { label: 'Dán Liên kết', icon: 'link', act: pst({ link: true }) },
          'sep',
          { label: 'Dán Đặc biệt…', hint: 'Ctrl+Alt+V', act: () => this.pasteSpecialOpen() },
        ]) }),
        col(
          B('sm', 'cut', 'cut', 'Cắt', { kt: 'X', kbd: 'Ctrl+X', act: () => this.copySel(true) }),
          B('sm', 'copy', 'copy', 'Sao chép', { kt: 'C', kbd: 'Ctrl+C', split: true, act: () => this.copySel(false), drop: this.menuAt(() => [{ label: 'Sao chép', icon: 'copy', act: () => this.copySel(false) }, { label: 'Sao chép dưới dạng Ảnh…', icon: 'copyPic', act: () => this.later('Sao chép dưới dạng Ảnh', '0.9') }]) }),
          B('sm', 'painter', 'brush', 'Chổi định dạng', { kt: 'FP', act: () => this.painterToggle(), dbl: () => this.painterOn(true), on: () => !!this.painter }),
        ),
      ], { id: 'clipboard', tip: 'Bảng tạm Office', act: () => toggleClipboardPane(this), kt: 'FO' }),
      G('font', 'Phông chữ', 'font', [
        col(
          row(
            { t: 'combo', id: 'fn', label: 'Phông', w: 128, kt: 'FF', icon: 'font', get: () => cur().fn ?? DEFAULT_FONT, set: (v) => { if (v) this.applyStyle({ fn: v === DEFAULT_FONT ? undefined : v }); }, list: (r) => this.fontMenu(r) },
            { t: 'combo', id: 'fs', label: 'Cỡ Phông', w: 46, kt: 'FS', icon: 'fontSize', get: () => String(cur().fs ?? 11), set: (v) => { const n = parseFloat(v.replace(',', '.')); if (n >= 1 && n <= 409) { this.applyStyle({ fs: n === 11 ? undefined : n }); this.fitRows(); } else this.toast('Cỡ phông phải từ 1 đến 409.', 'err'); }, list: (r) => this.sizeMenu(r) },
            B('ic', 'grow', 'grow', 'Tăng Cỡ Phông', { kt: 'FG', act: () => this.growFont(1) }),
            B('ic', 'shrink', 'shrink', 'Giảm Cỡ Phông', { kt: 'FK', act: () => this.growFont(-1) }),
          ),
          row(
            B('ic', 'b', 'bold', 'Đậm', { kt: '1', kbd: 'Ctrl+B', act: () => this.toggle('b'), on: st('b') }),
            B('ic', 'i', 'italic', 'Nghiêng', { kt: '2', kbd: 'Ctrl+I', act: () => this.toggle('i'), on: st('i') }),
            B('ic', 'u', 'underline', 'Gạch chân', { kt: '3', kbd: 'Ctrl+U', act: () => this.toggle('u'), on: st('u'), split: true, drop: this.menuAt(() => [{ label: 'Gạch chân', icon: 'underline', act: () => this.toggle('u') }, { label: 'Gạch chân Kép', icon: 'dblUnder', act: () => this.act(() => this.applyStyle({ u: true, ud: true })) }]) }),
            B('ic', 'border', 'borderBottom', 'Viền', { kt: 'B', split: true, act: () => this.safe(() => this.wb.setBorder(s(), this.grid.range, ...this.lastBorder)), drop: this.dd(() => this.borderMenu()) }),
            B('ic', 'bg', 'fill', 'Màu Tô', { kt: 'H', split: true, bar: () => this.lastFill, act: () => this.applyStyle({ bg: this.lastFill }), drop: (r) => this.colorMenu('bg', r) }),
            B('ic', 'fg', 'textColor', 'Màu Phông', { kt: 'FC', split: true, bar: () => this.lastFont, act: () => this.applyStyle({ fg: this.lastFont }), drop: (r) => this.colorMenu('fg', r) }),
          ),
        ),
      ], { id: 'font', tip: 'Định dạng Ô: Phông chữ (Ctrl+Shift+F)', act: () => formatCellsDialog(this, 'font'), kt: 'FN' }),
      G('align', 'Căn chỉnh', 'center', [
        col(
          row(
            B('ic', 'va-t', 'aTop', 'Căn Trên', { kt: 'AT', act: () => this.applyStyle({ va: cur().va === 't' ? undefined : 't' }), on: va('t') }),
            B('ic', 'va-m', 'aMid', 'Căn Giữa', { kt: 'AM', act: () => this.applyStyle({ va: cur().va === 'm' ? undefined : 'm' }), on: va('m') }),
            B('ic', 'va-b', 'aBot', 'Căn Dưới', { kt: 'AB', act: () => this.applyStyle({ va: undefined }), on: () => !cur().va || cur().va === 'b' }),
            B('ic', 'orient', 'orient', 'Hướng', { kt: 'FQ', drop: this.menuAt(() => [
              { label: 'Xoay Ngược Chiều Kim đồng hồ', act: () => this.act(() => this.applyStyle({ rot: 45 })) },
              { label: 'Xoay Theo Chiều Kim đồng hồ', act: () => this.act(() => this.applyStyle({ rot: -45 })) },
              { label: 'Chữ Dọc', act: () => this.act(() => this.applyStyle({ rot: 255 })) },
              { label: 'Xoay Chữ Lên', act: () => this.act(() => this.applyStyle({ rot: 90 })) },
              { label: 'Xoay Chữ Xuống', act: () => this.act(() => this.applyStyle({ rot: -90 })) },
              { label: 'Bỏ Xoay', act: () => this.applyStyle({ rot: undefined }) },
              'sep',
              { label: 'Định dạng Căn chỉnh Ô', icon: 'cellFmt', act: () => formatCellsDialog(this, 'align') },
            ]) }),
            B('sm', 'wr', 'wrap', 'Ngắt dòng', { kt: 'W', act: () => this.toggleWrap(), on: st('wr') }),
          ),
          row(
            B('ic', 'al-l', 'left', 'Căn Trái', { kt: 'AL', act: () => this.align('l'), on: al('l') }),
            B('ic', 'al-c', 'center', 'Căn giữa', { kt: 'AC', act: () => this.align('c'), on: al('c') }),
            B('ic', 'al-r', 'right', 'Căn Phải', { kt: 'AR', act: () => this.align('r'), on: al('r') }),
            B('ic', 'ind-d', 'indL', 'Giảm Thụt lề', { kt: '5', act: () => this.indent(-1) }),
            B('ic', 'ind-i', 'indR', 'Tăng Thụt lề', { kt: '6', act: () => this.indent(1) }),
            B('sm', 'merge', 'merge2', 'Gộp & Căn giữa', { kt: 'M', split: true, act: () => this.doMerge(true), on: () => !!this.wb.mergeAt(s(), this.grid.sel.ar, this.grid.sel.ac), drop: this.dd(() => this.mergeMenu()) }),
          ),
        ),
      ], { id: 'align', tip: 'Định dạng Ô: Căn chỉnh', act: () => formatCellsDialog(this, 'align'), kt: 'FA' }),
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
      ], { id: 'number', tip: 'Định dạng Ô: Số', act: () => formatCellsDialog(this, 'num'), kt: 'FM' }),
      G('styles', 'Kiểu', 'styles', [
        B('big', 'cf', 'cf', 'Định dạng có Điều kiện', { kt: 'L', drop: this.dd(() => this.cfMenu()) }),
        B('big', 'fmtTable', 'fmtTable', 'Định dạng như Bảng', { kt: 'T', drop: (r) => tableStylePicker(this, r, (st) => createTableDialog(this, st)) }),
        { t: 'gal', id: 'cstyle', label: 'Kiểu Ô', icon: 'styles', kt: 'J', items: CELL_STYLES.slice(0, 6).map(([n, sty, bk, bs, bc]) => ({ label: n, css: styleCss(sty, bk, bc), act: () => this.applyCellStyle(n, sty, bk, bs, bc) })), act: () => this.cellStyleMenu(), more: (r) => this.cellStyleMenu(r) },
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
            { label: 'Chuỗi…', act: () => seriesDialog(this) },
            { label: 'Điền Nhanh', hint: 'Ctrl+E', act: () => this.flashFill() },
          ]) }),
          B('sm', 'clear', 'eraser', 'Xóa', { kt: 'E', drop: this.menuAt(() => [
            { label: 'Xóa Tất cả', icon: 'eraser', act: () => this.clearAll() },
            { label: 'Xóa Định dạng', act: () => this.clearFormats() },
            { label: 'Xóa Nội dung', hint: 'Delete', act: () => this.clearSel() },
            { label: 'Xóa Ghi chú và Chú thích', act: () => this.act(() => this.eachArea((a) => this.wb.clearNotes(this.grid.sheet, a))) },
            { label: 'Xóa Siêu kết nối', act: () => this.clearLinks() },
          ]) }),
        ),
        B('big', 'sortf', 'sortAZ', 'Sắp xếp & Lọc', { kt: 'S', drop: this.dd(() => this.sortFilterMenu()) }),
        B('big', 'find', 'find', 'Tìm & Chọn', { kt: 'FD', drop: this.menuAt(() => [
          { label: 'Tìm…', icon: 'find', hint: 'Ctrl+F', act: () => findDialog(this, 'find') },
          { label: 'Thay thế…', hint: 'Ctrl+H', act: () => findDialog(this, 'replace') },
          { label: 'Đi tới…', hint: 'Ctrl+G', act: () => gotoDialog(this) },
          { label: 'Đi tới Đặc biệt…', act: () => gotoSpecialDialog(this) },
          'sep',
          { label: 'Công thức', act: () => this.selectSpecial({ kind: 'formulas' }) },
          { label: 'Ghi chú', act: () => this.selectSpecial({ kind: 'notes' }) },
          { label: 'Định dạng có Điều kiện', act: () => this.selectSpecial({ kind: 'cf' }) },
          { label: 'Hằng số', act: () => this.selectSpecial({ kind: 'constants' }) },
          { label: 'Xác thực Dữ liệu', act: () => this.selectSpecial({ kind: 'dv' }) },
          'sep',
          { label: 'Chọn Đối tượng', act: () => { if (!this.charts.cycle(1)) this.toast('Trang tính không có đối tượng nổi.'); } },
        ]) }),
      ]),
      G('addins', 'Phần bổ trợ', 'addins', [B('big', 'addins', 'addins', 'Phần bổ trợ', { kt: 'Y1', ...later('ngoài phạm vi (Office Store)') })]),
    ] };
    const insert: Tab = { id: 'insert', label: 'Chèn', kt: 'N', groups: [
      G('tables', 'Bảng', 'pivot', [
        B('big', 'pivot', 'pivot', 'PivotTable', { kt: 'V', act: () => this.insertPivot() }),
        B('big', 'recPivot', 'recommended', 'PivotTable Đề xuất', { kt: 'SP', ...later('0.9') }),
        B('big', 'table', 'fmtTable', 'Bảng', { kt: 'T', kbd: 'Ctrl+T', act: () => createTableDialog(this) }),
        B('big', 'forms', 'forms', 'Biểu mẫu', { kt: 'FO', ...later('ngoài phạm vi (Microsoft Forms)') }),
      ]),
      G('illus', 'Hình minh họa', 'picture', [
        B('big', 'pic', 'picture', 'Ảnh', { kt: 'P', ...later('0.9') }),
        B('big', 'shapes', 'shapes', 'Hình dạng', { kt: 'SH', ...later('0.9') }),
        B('big', 'icons', 'effects', 'Biểu tượng', { kt: 'Y1', ...later('0.9') }),
        col(B('sm', 'model3d', 'model3d', 'Mô hình 3D', { kt: 'S3', ...later('ngoài phạm vi') }), B('sm', 'smartart', 'smartart', 'SmartArt', { kt: 'S1', ...later('0.9') }), B('sm', 'screenshot', 'screenshot', 'Ảnh chụp màn hình', { kt: 'SC', ...later('0.9') })),
      ]),
      G('controls', 'Điều khiển', 'checkbox', [B('big', 'i-check', 'checkbox', 'Hộp kiểm', { kt: 'Z1', ...later('0.9') })]),
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
        B('big', 'maps', 'maps', 'Bản đồ', { kt: 'L1', ...later('ngoài phạm vi (Bing)') }),
        B('big', 'pchart', 'chartStack', 'PivotChart', { kt: 'ZC', ...later('0.9') }),
      ], { id: 'charts', tip: 'Xem Tất cả Biểu đồ', act: () => this.insertChart(), kt: 'K' }),
      G('spark', 'Biểu đồ thu nhỏ', 'sparkLine', [
        col(B('sm', 'sp-line', 'sparkLine', 'Đường', { kt: 'SL', ...later('0.9') }), B('sm', 'sp-col', 'sparkCol', 'Cột', { kt: 'SO', ...later('0.9') }), B('sm', 'sp-wl', 'winLoss', 'Thắng/Thua', { kt: 'SW', ...later('0.9') })),
      ]),
      G('filters', 'Bộ lọc', 'slicer', [B('big', 'slicer', 'slicer', 'Slicer', { kt: 'SF', ...later('0.9') }), B('big', 'tl', 'timeline', 'Dòng thời gian', { kt: 'LT', ...later('0.9') })]),
      G('links', 'Nối kết', 'link', [B('big', 'link', 'link', 'Liên kết', { kt: 'I', kbd: 'Ctrl+K', act: () => hyperlinkDialog(this) })]),
      G('comments', 'Chú thích', 'comment', [B('big', 'i-comment', 'comment', 'Chú thích', { kt: 'M', kbd: 'Ctrl+Shift+F2', act: () => this.notes.openCard() })]),
      G('text', 'Văn bản', 'textbox', [
        B('big', 'textbox', 'textbox', 'Hộp Văn bản', { kt: 'X', ...later('0.9') }),
        B('big', 'hf', 'hf', 'Đầu & Chân trang', { kt: 'H', act: () => this.pageSetup('hf') }),
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
        B('big', 'ar-fwd', 'arrange', 'Đưa Lên trên', { kt: 'AF', ...later('0.9') }),
        B('big', 'ar-back', 'arrange', 'Đưa Xuống dưới', { kt: 'AE', ...later('0.9') }),
        B('big', 'ar-pane', 'slicer', 'Ngăn Chọn', { kt: 'AP', ...later('0.9') }),
        col(B('sm', 'ar-align', 'align', 'Căn', { kt: 'AA', ...later('0.9') }), B('sm', 'ar-group', 'groupObj', 'Nhóm', { kt: 'AG', ...later('0.9') }), B('sm', 'ar-rotate', 'rotate', 'Xoay', { kt: 'AY', ...later('0.9') })),
      ]),
    ] };
    const formulas: Tab = { id: 'formulas', label: 'Công thức', kt: 'M', groups: [
      G('lib', 'Thư viện Hàm', 'fx', [
        B('big', 'fnp', 'fx', 'Chèn Hàm', { kt: 'F', kbd: 'Shift+F3', act: () => this.funcPicker() }),
        B('big', 'fsum', 'sigma', 'Tổng Tự động', { kt: 'U', split: true, act: () => this.autoSum(), drop: this.menuAt(() => ([['Tổng', 'SUM'], ['Trung bình', 'AVERAGE'], ['Đếm Số', 'COUNT'], ['Tối đa', 'MAX'], ['Tối thiểu', 'MIN']] as const).map(([l, f]) => ({ label: l, act: () => this.autoSum(f) }))) }),
        B('big', 'f-recent', 'recentFn', 'Dùng Gần đây', { kt: 'R', drop: this.menuAt(() => [...this.recentFns().map((f) => ({ label: f, act: () => this.insertFn(f) })), 'sep', { label: 'Chèn Hàm…', act: () => this.funcPicker() }]) }),
        ...(([['f-fin', 'Tài chính', 'money', 'I', 'PMT NPV IRR FV PV RATE NPER'], ['f-log', 'Lô-gic', 'validate', 'L', 'IF IFS AND OR NOT IFERROR XOR SWITCH'], ['f-text', 'Văn bản', 'font', 'T', 'TEXT LEFT RIGHT MID LEN CONCAT TRIM UPPER LOWER SUBSTITUTE'], ['f-date', 'Ngày & Giờ', 'calcOpts', 'E', 'TODAY NOW DATE YEAR MONTH DAY EDATE EOMONTH NETWORKDAYS'], ['f-look', 'Tra cứu & Tham chiếu', 'find', 'O', 'VLOOKUP HLOOKUP XLOOKUP INDEX MATCH OFFSET INDIRECT'], ['f-math', 'Toán & Lượng giác', 'sigma', 'G', 'SUM SUMIF SUMIFS SUMPRODUCT ROUND ROUNDUP ROUNDDOWN INT ABS MOD']] as const)
          .map(([id, l, ico, k, fns]) => B('big', id, ico, l, { kt: k, drop: this.menuAt(() => fns.split(' ').map((f) => ({ label: f, act: () => this.insertFn(f) }))) }))),
        B('big', 'f-more', 'moreFn', 'Hàm Khác', { kt: 'Q', drop: this.menuAt(() => ([['Thống kê', 'AVERAGE AVERAGEIF COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MAXIFS MEDIAN MIN MINIFS MODE RANK STDEV.S VAR.S FORECAST'], ['Kỹ thuật', 'CONVERT DEC2BIN BIN2DEC DEC2HEX HEX2DEC DELTA'], ['Khối', 'CUBEVALUE CUBEMEMBER'], ['Thông tin', 'ISBLANK ISERROR ISNUMBER ISTEXT ISNA TYPE N NA CELL'], ['Tương thích', 'STDEV VAR RANK PERCENTILE QUARTILE'], ['Web', 'ENCODEURL']] as const).map(([cat, fns]) => ({ label: cat, sub: fns.split(' ').map((f) => ({ label: f, act: () => this.insertFn(f) })) }) as MenuItem)) }),
      ]),
      G('names', 'Tên đã Xác định', 'nameMgr', [
        B('big', 'nm-mgr', 'nameMgr', 'Trình Quản lý Tên', { kt: 'N', kbd: 'Ctrl+F3', act: () => nameManager(this) }),
        col(
          B('sm', 'nm-def', 'tag', 'Xác định Tên', { kt: 'MD', act: () => nameDialog(this) }),
          B('sm', 'nm-use', 'fx', 'Dùng trong Công thức', { kt: 'S', drop: this.menuAt(() => {
            const list = this.wb.listNames().filter((n) => !n.hidden && (n.scope === undefined || n.scope === this.grid.sheet));
            return [...(list.length ? list.map((n) => ({ label: n.name, act: () => this.pasteName(n.name) })) : [{ label: '(Chưa có tên)', act: () => {} }]), 'sep', { label: 'Dán Tên…', kbd: 'F3', act: () => pasteNameDialog(this, this.ed.isEditing()) }] as MenuItem[];
          }) }),
          B('sm', 'nm-create', 'table', 'Tạo từ Vùng chọn', { kt: 'C', kbd: 'Ctrl+Shift+F3', act: () => createFromSelection(this) }),
        ),
      ]),
      G('audit', 'Kiểm nghiệm Công thức', 'trace', [
        col(B('sm', 'tr-pre', 'trace', 'Truy tìm Ô Tiền lệ', { kt: 'P', act: () => this.trace(true) }), B('sm', 'tr-dep', 'traceDep', 'Truy tìm Ô Phụ thuộc', { kt: 'D', act: () => this.trace(false) }), B('sm', 'tr-rm', 'arrowsX', 'Loại bỏ Mũi tên', { kt: 'AA', act: () => { this.grid.arrows = []; this.traced = null; this.grid.invalidate(); } })),
        col(B('sm', 'show-f', 'showF', 'Hiện Công thức', { kt: 'H', kbd: 'Ctrl+`', on: () => this.grid.showFormulas, act: () => { this.grid.showFormulas = !this.grid.showFormulas; this.grid.invalidate(); } }), B('sm', 'err-chk', 'errorCheck', 'Kiểm tra Lỗi', { kt: 'K', ...later('0.9') }), B('sm', 'eval-f', 'evaluate', 'Đánh giá Công thức', { kt: 'V', ...later('0.9') })),
        B('big', 'watch', 'watch', 'Cửa sổ Theo dõi', { kt: 'W', ...later('0.9') }),
      ]),
      G('calc', 'Tính toán', 'calcOpts', [
        B('big', 'calc-opt', 'calcOpts', 'Tùy chọn Tính toán', { kt: 'X', drop: this.menuAt(() => [{ label: 'Tự động', check: true, act: () => {} }, { label: 'Thủ công', act: () => this.later('Tính toán Thủ công', '0.9') }]) }),
        col(
          B('sm', 'recalc', 'calc', 'Tính Ngay', { kt: 'B', kbd: 'F9', act: () => this.recalc() }),
          B('sm', 'recalc-s', 'calcSheet', 'Tính Trang tính', { kt: 'J', kbd: 'Shift+F9', act: () => this.recalc() }),
        ),
      ]),
    ] };
    const data: Tab = { id: 'data', label: 'Dữ liệu', kt: 'A', groups: [
      G('get', 'Lấy & Chuyển đổi Dữ liệu', 'getData', [
        B('big', 'd-get', 'getData', 'Lấy Dữ liệu', { kt: 'PN', drop: this.menuAt(() => [{ label: 'Từ Sổ làm việc Excel…', icon: 'file', act: () => this.openFile() }, { label: 'Từ Văn bản/CSV…', icon: 'fromText', act: () => this.openFile() }, { label: 'Từ Web…', icon: 'web', act: () => this.later('Lấy dữ liệu từ Web', 'ngoài phạm vi') }]) }),
        col(B('sm', 'd-open', 'fromText', 'Từ Văn bản/CSV', { kt: 'FT', act: () => this.openFile() }), B('sm', 'd-web', 'web', 'Từ Web', { kt: 'FW', ...later('ngoài phạm vi') }), B('sm', 'd-tbl', 'fmtTable', 'Từ Bảng/Vùng', { kt: 'FB', ...later('ngoài phạm vi') })),
        col(B('sm', 'd-pic', 'fromPic', 'Từ Ảnh', { kt: 'FP', ...later('ngoài phạm vi') }), B('sm', 'd-recent', 'recentSrc', 'Nguồn Gần đây', { kt: 'FR', act: () => this.backstage.open('open') }), B('sm', 'd-conn', 'existConn', 'Kết nối Hiện có', { kt: 'FE', ...later('ngoài phạm vi') })),
      ]),
      G('queries', 'Truy vấn & Kết nối', 'conn', [
        B('big', 'd-refresh', 'refresh', 'Làm mới Tất cả', { kt: 'R', kbd: 'Ctrl+Alt+F5', act: () => { for (const id of this.wb.sheetIds()) for (const p of this.wb.meta(id).pivots) this.safe(() => this.wb.refreshPivot(id, p.id!)); this.recalc(); } }),
        col(B('sm', 'd-qc', 'queries', 'Truy vấn & Kết nối', { kt: 'O', ...later('ngoài phạm vi (Power Query)') }), B('sm', 'd-props', 'props', 'Thuộc tính', { kt: 'PP', ...later('ngoài phạm vi') }), B('sm', 'd-links', 'link', 'Liên kết Sổ làm việc', { kt: 'KL', act: () => this.toast(this.wb.extLinks.length ? `Sổ có ${this.wb.extLinks.length} liên kết ngoài (giữ giá trị lưu sẵn, ghi lại khi lưu .xlsx).` : 'Sổ không có liên kết ngoài.') })),
      ]),
      G('dtypes', 'Kiểu Dữ liệu', 'stocks', [B('big', 'dt-stock', 'stocks', 'Cổ phiếu', { kt: 'K1', ...later('ngoài phạm vi (dịch vụ trực tuyến)') }), B('big', 'dt-cur', 'currency', 'Tiền tệ', { kt: 'K2', ...later('ngoài phạm vi (dịch vụ trực tuyến)') }), B('big', 'dt-geo', 'geo', 'Địa lý', { kt: 'K3', ...later('ngoài phạm vi (dịch vụ trực tuyến)') })]),
      G('sortfilter', 'Sắp xếp & Lọc', 'filter', [
        col(B('ic', 'd-az', 'sortAZ', 'Sắp xếp từ A đến Z', { kt: 'SA', act: () => this.quickSort(false) }), B('ic', 'd-za', 'sortZA', 'Sắp xếp từ Z đến A', { kt: 'SD', act: () => this.quickSort(true) })),
        B('big', 'd-sort', 'sortAZ', 'Sắp xếp', { kt: 'SS', act: () => this.sortDlg() }),
        B('big', 'd-f', 'filter', 'Lọc', { kt: 'T', kbd: 'Ctrl+Shift+L', act: () => this.toggleFilter(), on: () => !!this.wb.meta(s()).filter }),
        col(B('sm', 'd-fx', 'filterX', 'Xóa', { kt: 'C', act: () => this.clearFilterConds() }), B('sm', 'd-fr', 'refresh', 'Áp dụng lại', { kt: 'Y', act: () => { this.wb.reapplyFilter(s()); this.grid.layout(); } }), B('sm', 'd-adv', 'filter', 'Nâng cao', { kt: 'Q', ...later('0.9') })),
      ]),
      G('tools', 'Công cụ Dữ liệu', 'textCols', [
        B('big', 'd-t2c', 'textCols', 'Văn bản thành Cột', { kt: 'E', act: () => textToColumnsDialog(this) }),
        col(B('ic', 'd-flash', 'flash', 'Điền Nhanh', { kt: 'FF', kbd: 'Ctrl+E', act: () => this.flashFill() }), B('ic', 'd-dup', 'removeDup', 'Loại bỏ Trùng lặp', { kt: 'M', act: () => removeDupDialog(this) }), B('ic', 'd-v', 'validate', 'Xác thực Dữ liệu', { kt: 'V', act: () => this.dvDlg() })),
        col(B('sm', 'd-cons', 'subtotal', 'Hợp nhất', { kt: 'N', ...later('0.9') }), B('sm', 'd-model', 'dataModel', 'Mô hình Dữ liệu', { kt: 'DM', ...later('ngoài phạm vi (Power Pivot)') }), B('sm', 'd-analyze', 'analyze', 'Phân tích Dữ liệu', { kt: 'AD', ...later('ngoài phạm vi (dịch vụ AI trực tuyến)') })),
      ]),
      G('forecast', 'Dự báo', 'forecast', [B('big', 'd-whatif', 'whatIf', 'Phân tích Giả định', { kt: 'W', ...later('0.9') }), B('big', 'd-fc', 'forecast', 'Trang tính Dự báo', { kt: 'FC', ...later('0.9') })]),
      G('outline', 'Đại cương', 'group', [
        B('big', 'd-group', 'group', 'Nhóm', { kt: 'G', kbd: 'Alt+Shift+→', split: true, act: () => groupSel(this), drop: this.menuAt(() => [{ label: 'Nhóm…', hint: 'Alt+Shift+→', act: () => groupSel(this) }] as MenuItem[]) }),
        B('big', 'd-ungroup', 'ungroup', 'Bỏ nhóm', { kt: 'U', kbd: 'Alt+Shift+←', split: true, act: () => ungroupSel(this), drop: this.menuAt(() => [{ label: 'Bỏ nhóm…', hint: 'Alt+Shift+←', act: () => ungroupSel(this) }, { label: 'Xóa Đại cương', act: () => clearOutline(this) }] as MenuItem[]) }),
        B('big', 'd-sub', 'subtotal', 'Tổng phụ', { kt: 'B', act: () => subtotalDialog(this) }),
        col(B('sm', 'd-showd', 'plus', 'Hiện Chi tiết', { kt: 'J', act: () => showDetail(this, true) }), B('sm', 'd-hided', 'minus', 'Ẩn Chi tiết', { kt: 'H', act: () => showDetail(this, false) })),
      ]),
    ] };
    const review: Tab = { id: 'review', label: 'Xem lại', kt: 'R', groups: [
      G('proof', 'Soát lỗi', 'spell', [B('big', 'r-spell', 'spell', 'Chính tả', { kt: 'S', kbd: 'F7', ...later('0.9') }), col(B('sm', 'r-thes', 'thesaurus', 'Từ điển Đồng nghĩa', { kt: 'E', ...later('ngoài phạm vi') }), B('sm', 'r-stats', 'stats', 'Thống kê Sổ làm việc', { kt: 'W', act: () => this.bookStats() }))]),
      G('perf', 'Hiệu suất', 'perf', [B('big', 'r-perf', 'perf', 'Kiểm tra Hiệu suất', { kt: 'PF', act: () => this.checkPerformance() })]),
      G('access', 'Trợ năng', 'access', [B('big', 'r-access', 'access', 'Kiểm tra Trợ năng', { kt: 'A1', act: () => this.checkAccessibility() })]),
      G('lang', 'Ngôn ngữ', 'translate', [B('big', 'r-translate', 'translate', 'Dịch', { kt: 'L', ...later('ngoài phạm vi (dịch vụ trực tuyến)') })]),
      G('changes', 'Thay đổi', 'showChanges', [B('big', 'r-changes', 'showChanges', 'Hiện Thay đổi', { kt: 'H1', ...later('ngoài phạm vi (cộng tác trực tuyến)') })]),
      G('comments2', 'Chú thích', 'comment', [
        B('big', 'r-comment', 'comment', 'Chú thích Mới', { kt: 'C', kbd: 'Ctrl+Shift+F2', act: () => this.notes.openCard() }),
        B('big', 'r-cdel', 'commentDel', 'Xóa', { kt: 'D', act: () => this.act(() => this.eachArea((a) => { if (!this.wb.clearNotes(this.grid.sheet, a, 'comments')) this.toast('Vùng chọn không có chú thích.'); })) }),
        B('big', 'r-cprev', 'commentPrev', 'Chú thích Trước', { kt: 'V', act: () => this.notes.step('comment', -1) }),
        B('big', 'r-cnext', 'commentNext', 'Chú thích Tiếp theo', { kt: 'N', act: () => this.notes.step('comment', 1) }),
        B('big', 'r-cshow', 'comments', 'Hiện Chú thích', { kt: 'H2', act: () => this.notes.togglePane(), on: () => this.notes.paneOpen }),
      ]),
      G('notes', 'Ghi chú', 'note', [B('big', 'r-note', 'note', 'Ghi chú', { kt: 'T', drop: this.menuAt(() => {
        const st = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac);
        return [
          { label: st.nt ? 'Sửa Ghi chú' : 'Ghi chú Mới', icon: 'note', kbd: 'Shift+F2', act: () => this.notes.editNote() },
          { label: 'Ghi chú Trước', act: () => this.notes.step('note', -1) },
          { label: 'Ghi chú Tiếp theo', act: () => this.notes.step('note', 1) },
          { label: 'Hiện Tất cả Ghi chú', check: this.notes.allShown, act: () => this.notes.toggleAll() },
          'sep',
          { label: 'Xóa Ghi chú', act: () => this.act(() => this.eachArea((a) => this.wb.clearNotes(this.grid.sheet, a, 'notes'))) },
        ] as MenuItem[];
      }) })]),
      G('protect', 'Bảo vệ', 'protectSheet', [B('big', 'r-psheet', 'protectSheet', 'Bảo vệ Trang tính', { kt: 'PS', ...later('0.9') }), B('big', 'r-pbook', 'protectBook', 'Bảo vệ Sổ làm việc', { kt: 'PW', ...later('0.9') }), B('big', 'r-allow', 'allowEdit', 'Cho phép Sửa Vùng', { kt: 'U1', ...later('0.9') }), B('big', 'r-unshare', 'unshare', 'Hủy chia sẻ Sổ làm việc', { kt: 'UW', ...later('ngoài phạm vi') })]),
      G('ink', 'Mực', 'ink', [B('big', 'r-ink', 'ink', 'Ẩn Mực', { kt: 'K', ...later('ngoài phạm vi (bút cảm ứng)') })]),
    ] };
    const view: Tab = { id: 'view', label: 'Xem', kt: 'W', groups: [
      G('sheetview', 'Dạng xem Trang tính', 'sheetView', [
        col({ t: 'combo', id: 'sv-name', label: 'Dạng xem Trang tính', w: 150, kt: 'KV', icon: 'sheetView', get: () => 'Mặc định', set: () => this.later('Dạng xem Trang tính', 'ngoài phạm vi (cộng tác trực tuyến)'), list: () => this.later('Dạng xem Trang tính', 'ngoài phạm vi (cộng tác trực tuyến)') },
          row(B('sm', 'sv-keep', 'keep', 'Giữ', { kt: 'KK', ...later('ngoài phạm vi') }), B('sm', 'sv-exit', 'exitView', 'Thoát', { kt: 'KX', ...later('ngoài phạm vi') }), B('sm', 'sv-new', 'plus', 'Mới', { kt: 'KN', ...later('ngoài phạm vi') }), B('sm', 'sv-opt', 'gear', 'Tùy chọn', { kt: 'KO', ...later('ngoài phạm vi') }))),
      ]),
      G('views', 'Dạng xem Sổ làm việc', 'normalView', [
        B('big', 'v-normal', 'normalView', 'Thường', { kt: 'L', act: () => this.setView('normal'), on: () => this.grid.view === 'normal' }),
        B('big', 'v-break', 'breakView', 'Xem trước Ngắt Trang', { kt: 'I', act: () => this.setView('break'), on: () => this.grid.view === 'break' }),
        B('big', 'v-layout', 'pageView', 'Bố trí Trang', { kt: 'P', ...later('0.8') }),
        B('big', 'v-custom', 'customViews', 'Dạng xem Tùy chỉnh', { kt: 'C', ...later('0.9') }),
      ]),
      G('show', 'Hiện', 'eye', [
        B('big', 'v-nav', 'navPane', 'Dẫn hướng', { kt: 'VN', act: () => this.navPane(), on: () => !!document.getElementById('nav-pane') }),
        col(
          { t: 'chk', id: 'v-ruler', label: 'Thước', kt: 'VR', get: () => false, set: () => { this.later('Thước (dạng xem Bố trí Trang)', '0.8'); this.ribbon.refresh(); } },
          { t: 'chk', id: 'gl', label: 'Đường lưới', kt: 'VG', get: () => this.grid.gridlines, set: (v) => { this.grid.gridlines = v; this.grid.invalidate(); } },
          { t: 'chk', id: 'v-fbar', label: 'Thanh Công thức', kt: 'VF', get: () => !$('#app').classList.contains('no-fbar'), set: (v) => { $('#app').classList.toggle('no-fbar', !v); this.grid.layout(); } },
        ),
        col(
          { t: 'chk', id: 'v-heads', label: 'Đầu đề', kt: 'VH', get: () => this.grid.headings, set: (v) => { this.grid.headings = v; this.grid.layout(); } },
          { t: 'chk', id: 'v-dticons', label: 'Biểu tượng Kiểu Dữ liệu', kt: 'VD', get: () => true, set: () => { this.later('Kiểu Dữ liệu', 'ngoài phạm vi'); this.ribbon.refresh(); } },
          { t: 'chk', id: 'v-focus', label: 'Ô Tiêu điểm', kt: 'VC', get: () => this.grid.focusCell, set: (v) => { this.grid.focusCell = v; this.grid.invalidate(); } },
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
        col(B('sm', 'w-side', 'sideBySide', 'Xem Cạnh nhau', { kt: 'B', ...later('0.9') }), B('sm', 'w-sync', 'syncScroll', 'Cuộn Đồng bộ', { kt: 'YS', ...later('0.9') }), B('sm', 'w-reset', 'resetWin', 'Đặt lại Vị trí Cửa sổ', { kt: 'T', ...later('0.9') })),
        B('big', 'w-switch', 'switchWin', 'Chuyển Cửa sổ', { kt: 'W', drop: this.menuAt(() => [{ label: this.wb.name, check: true, act: () => {} }]) }),
      ]),
      G('macros', 'Macro', 'macro', [B('big', 'macros', 'macro', 'Macro', { kt: 'M', ...later('ngoài phạm vi (không chạy VBA)') })]),
    ] };
    const automate: Tab = { id: 'automate', label: 'Tự động hóa', kt: 'U', groups: [
      G('scripts', 'Office Scripts', 'scripts', [B('big', 'au-new', 'scripts', 'Tập lệnh Mới', { kt: 'N', ...later('ngoài phạm vi (Office Scripts trên web)') }), B('big', 'au-all', 'scripts', 'Xem Tất cả Tập lệnh', { kt: 'A', ...later('ngoài phạm vi (Office Scripts trên web)') })]),
      G('flows', 'Power Automate', 'powerAuto', [B('big', 'au-flow', 'powerAuto', 'Tự động hóa Công việc', { kt: 'F', ...later('ngoài phạm vi (Power Automate)') })]),
    ] };
    const help: Tab = { id: 'help', label: 'Trợ giúp', kt: 'Y', groups: [
      G('help', 'Trợ giúp', 'help', [
        B('big', 'h-help', 'help', 'Trợ giúp', { kt: 'H', act: () => this.helpDlg() }),
        B('big', 'h-keys', 'kbd', 'Phím tắt', { kt: 'K', act: () => this.helpDlg(true) }),
        B('big', 'h-new', 'news', 'Có gì Mới', { kt: 'W', act: () => this.toast(`ExcelLike ${VERSION}: Tên vùng, Ghi chú & Chú thích, Nhóm hàng/cột & Tổng phụ, Văn bản thành Cột, Loại bỏ Trùng lặp, Bảng (Ctrl+T) với tham chiếu có cấu trúc, mảng động FILTER/SORT/UNIQUE, Truy tìm Ô Tiền lệ/Phụ thuộc.`) }),
        B('big', 'about', 'info', 'Giới thiệu', { kt: 'A', act: () => this.backstage.open('account') }),
      ]),
    ] };
    const tcur = this.wb.tableAt(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac);
    return [{ id: 'file', label: 'Tệp', kt: 'F', groups: [] }, home, insert, layout, formulas, data, review, view, automate, help, ...(tcur ? [tableDesignTab(this, tcur)] : [])];
  }

  // ---------- lệnh định dạng ----------
  safe(f: () => void) { try { f(); } catch (e) { this.toast((e as Error).message, 'err'); } }
  applyStyle(p: Style) {
    const areas = this.grid.ranges(), sheets = this.sheetsTarget();
    this.safe(() => this.wb.group(() => { for (const sh of sheets) for (const a of areas) this.wb.setStyle(sh, a, p); }));
    this.lastAct = () => this.applyStyle(p);
    this.refreshToggles();
  }
  toggle(k: 'b' | 'i' | 'u' | 's') {
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
    const b = (kind: BorderKind, style: BorderStyle = 'thin', remember = true) => () => { if (remember) this.lastBorder = [kind, style]; this.act(() => this.eachArea((a) => { for (const sh of this.sheetsTarget()) this.wb.setBorder(sh, a, kind, style); })); };
    const two = (top: BorderStyle, bot: BorderStyle) => () => this.act(() => this.eachArea((a) => { this.wb.setBorder(s, a, 'top', top); this.wb.setBorder(s, a, 'bottom', bot); }));
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
      { label: 'Thêm Viền…', icon: 'cellFmt', act: () => formatCellsDialog(this, 'border') },
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
  filterPopover(col: number, x: number, y: number) {
    const s = this.grid.sheet, f = this.wb.meta(s).filter;
    if (!f) return;
    const all = this.wb.filterValues(s, col);
    const cur = f.conds[col];
    // điều kiện cũ không còn khớp giá trị nào (vd lọc rỗng) → coi như chưa lọc, chọn tất cả như Excel
    const curVals = cur?.values?.filter((v) => all.includes(v));
    const box = h('div', { class: 'filter-pop' });
    const sortA = h('button', { class: 'mi' }); sortA.innerHTML = ic('sortAZ', 16) + '<span>Sắp xếp A → Z</span>';
    const sortZ = h('button', { class: 'mi' }); sortZ.innerHTML = ic('sortZA', 16) + '<span>Sắp xếp Z → A</span>';
    const doSort = (desc: boolean) => () => { closeMenu(); this.safe(() => this.wb.sort(s, this.wb.filterRange(s) ?? f.range, [{ col, desc }], true)); };
    sortA.onclick = doSort(false); sortZ.onclick = doSort(true);
    const q = h('input', { type: 'search', placeholder: 'Tìm…', 'aria-label': 'Tìm giá trị' }) as HTMLInputElement;
    const txt = h('input', { type: 'text', placeholder: 'Lọc theo chữ chứa…', 'aria-label': 'Lọc theo chữ chứa' }) as HTMLInputElement;
    txt.value = cur?.text ?? '';
    const list = h('div', { class: 'fvals', role: 'group', 'aria-label': 'Giá trị' });
    const chosen = new Set(curVals?.length ? curVals : all);
    const shown = all.slice(0, 1000);
    const allBox = h('input', { type: 'checkbox' }) as HTMLInputElement;
    const ok = h('button', { class: 'btn primary' }, 'Lọc') as HTMLButtonElement;
    // như Excel: không chọn giá trị nào thì không cho bấm Lọc
    const canOk = () => { ok.disabled = chosen.size === 0 && !txt.value.trim(); };
    txt.oninput = canOk;
    const render = () => {
      const qq = q.value.trim().toLowerCase();
      list.replaceChildren(h('label', { class: 'chk' }, allBox, ' (Chọn tất cả)'));
      allBox.checked = shown.every((v) => chosen.has(v));
      for (const v of shown) {
        if (qq && !v.toLowerCase().includes(qq)) continue;
        const cb = h('input', { type: 'checkbox' }) as HTMLInputElement;
        cb.checked = chosen.has(v);
        cb.onchange = () => { if (cb.checked) chosen.add(v); else chosen.delete(v); allBox.checked = shown.every((x) => chosen.has(x)); canOk(); };
        list.append(h('label', { class: 'chk' }, cb, ' ' + (v === '' ? '(Trống)' : v)));
      }
      if (!all.length) list.append(h('p', { class: 'muted' }, 'Không có dữ liệu bên dưới hàng tiêu đề.'));
      canOk();
    };
    allBox.onchange = () => { for (const v of shown) if (allBox.checked) chosen.add(v); else chosen.delete(v); render(); };
    q.oninput = render;
    render();
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
      // đọc thành dữ liệu sổ rồi nạp lại trang: sổ chỉ dựng một lần (khi khởi động), không dựng hai lần như trước
      const { readXlsxBook } = await import('../core/xlsx');
      const { book, warnings } = await readXlsxBook(new Uint8Array(await file.arrayBuffer()));
      await this.saveNow(false);
      const id = 'b' + Date.now();
      await saveBook({ id, name: baseName(file.name), updated: Date.now(), data: { ...book, name: baseName(file.name), pendingImport: true } as any, handle, fileName: file.name });
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
  clearSel() { const areas = this.grid.ranges(), sheets = this.sheetsTarget(); this.safe(() => this.wb.group(() => { for (const sh of sheets) this.wb.withHeaders(sh, () => { for (const a of areas) this.wb.clear(sh, a); }); })); this.lastAct = () => this.clearSel(); }
  autoSum(fn = 'SUM') {
    const { ar, ac } = this.grid.sel;
    const f = this.wb.autoSum(this.grid.sheet, ar, ac, fn);
    if (f === `=${fn}()`) this.ed.begin(`=${fn}(`, 'enter');
    else this.syncSelection();
  }
  fillDir(dir: 'down' | 'right' | 'up' | 'left') {
    const g = this.grid.range;
    if (dir === 'down' && g.r2 > g.r1) this.safe(() => this.wb.withHeaders(this.grid.sheet, () => this.wb.fill(this.grid.sheet, { ...g, r2: g.r1 }, g)));
    if (dir === 'right' && g.c2 > g.c1) this.safe(() => this.wb.withHeaders(this.grid.sheet, () => this.wb.fill(this.grid.sheet, { ...g, c2: g.c1 }, g)));
    if (dir === 'up' && g.r2 > g.r1) this.safe(() => this.wb.withHeaders(this.grid.sheet, () => this.wb.fill(this.grid.sheet, { ...g, r1: g.r2 }, g)));
    if (dir === 'left' && g.c2 > g.c1) this.safe(() => this.wb.withHeaders(this.grid.sheet, () => this.wb.fill(this.grid.sheet, { ...g, c1: g.c2 }, g)));
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
    if (this.grid.areas.length) { this.toast('Không thể dùng lệnh này với vùng chọn nhiều khối.', 'err'); return ''; }
    const t = this.wb.copy(this.grid.sheet, this.grid.range, cut);
    pushClip(t);
    this.lastCopyText = t;
    navigator.clipboard?.writeText(t).catch(() => {});
    this.grid.invalidate();
    return t;
  }
  private pasteAt(text: string | null) {
    const { ar, ac } = this.grid.sel;
    this.safe(() => {
      let g: Range | null;
      if (this.wb.clip && (text === null || normalizeNL(text) === normalizeNL(this.lastCopyText))) g = this.wb.withHeaders(this.grid.sheet, () => this.wb.paste(this.grid.sheet, ar, ac));
      else if (text) g = this.wb.withHeaders(this.grid.sheet, () => this.wb.pasteText(this.grid.sheet, ar, ac, text));
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
      b.className = 'st' + (id === this.grid.sheet ? ' on' : '') + (tc ? ' tc' : '') + (this.group.length > 1 && this.group.includes(id) ? ' grp' : '');
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
        if (this.group.length > 1 && !this.group.includes(id)) { this.group = []; this.updateTitle(); }
        else if (this.group.length > 1 && id !== this.grid.sheet) { /* bấm thẻ trong nhóm: giữ nhóm */ }
        else if (this.group.length > 1) { this.group = []; this.updateTitle(); }
        if (id !== this.grid.sheet) { this.grid.setSheet(id); this.renderSheets(); this.syncSelection(); }
        else this.renderSheets();
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
    $('#st-acc').onclick = () => this.checkAccessibility();
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
      const text = nb.value.trim();
      const g = parseRange(text.replace(/\s/g, ''));
      if (g) { this.grid.selectRange(g); this.grid.focusKeys(); return; }
      // tên đã có → nhảy tới; tên mới hợp lệ → đặt tên cho vùng đang chọn (như Excel)
      if (this.wb.tableByName(text)) { this.gotoName(text); return; }
      const t = this.wb.nameTarget(text, this.grid.sheet);
      if (t) { this.gotoName(text); return; }
      if (this.wb.listNames().some((n) => n.name.toUpperCase() === text.toUpperCase())) { this.toast(`Tên “${text}” không trỏ tới một vùng ô.`, 'err'); return; }
      try { this.wb.defineName(text, selRef(this)); this.afterNames(); this.toast(`Đã đặt tên “${text}” cho ${nb.dataset.ref ?? 'vùng chọn'}.`, 'ok'); this.grid.focusKeys(); }
      catch (err) { this.toast((err as Error).message + ' Ví dụ địa chỉ: B5 hoặc A1:C10; ví dụ tên: DonGia.', 'err'); }
    };
    $('#namebox-dd').onclick = () => {
      const list = this.wb.listNames().filter((n) => !n.hidden && (n.scope === undefined || n.scope === this.grid.sheet) && this.wb.nameTarget(n.name, this.grid.sheet));
      const r = nb.getBoundingClientRect();
      const all = [...this.wb.tableReg().map((t) => t.name), ...list.map((n) => n.name)].sort((a, b) => a.localeCompare(b, 'vi'));
      showMenu(all.length ? all.map((n) => ({ label: n, act: () => this.gotoName(n) })) : [{ label: '(Chưa có tên vùng)', act: () => {} }], r.left, r.bottom + 2);
    };
    $('#fx-fn').onclick = () => this.funcPicker();
    $('#fx-exp').onclick = () => { $('#app').classList.toggle('fx-tall'); this.grid.layout(); };
    bindRibbonKeys((el) => el.click(), () => this.grid.focusKeys());
  }
  /** sau khi thêm/sửa/xóa tên: vẽ lại, cập nhật thanh công thức */
  afterNames() { this.grid.invalidate(); this.syncSelection(); this.ribbon.refresh(); }
  /** nhảy tới vùng của tên (Hộp Tên, F5) */
  gotoName(name: string) {
    const tb = this.wb.tableByName(name);
    if (tb) {
      if (tb.sheet !== this.grid.sheet) { this.grid.setSheet(tb.sheet); this.renderSheets(); }
      const d = { r1: tb.r1 + (tb.header ? 1 : 0), c1: tb.c1, r2: tb.r2 - (tb.totals ? 1 : 0), c2: tb.c2 };
      this.grid.selectRange(d); this.syncSelection(); this.grid.focusKeys(); return;
    }
    const t = this.wb.nameTarget(name, this.grid.sheet);
    if (!t) { this.toast(`Tên “${name}” không trỏ tới một vùng ô.`, 'err'); return; }
    if (t.sheet !== this.grid.sheet) { this.grid.setSheet(t.sheet); this.renderSheets(); }
    this.grid.selectRange(t.range); this.syncSelection(); this.grid.focusKeys();
  }
  /** Dán Tên: đang sửa ô → chèn vào công thức; không thì bắt đầu công thức =Tên */
  pasteName(name: string) {
    if (this.ed.isEditing()) { this.ed.refocus(); this.ed.insertText(name); return; }
    this.ed.begin('=' + name, 'enter');
  }
  /** Dán Danh sách tên vào trang tính tại ô hiện hành (tên | tham chiếu) */
  pasteNameList() {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel;
    const list = this.wb.listNames().filter((n) => !n.hidden && (n.scope === undefined || n.scope === s));
    this.safe(() => this.wb.writeCells(s, list.flatMap((n, i) => [[ar + i, ac, "'" + n.name], [ar + i, ac + 1, "'" + n.expr]] as [number, number, string][])));
    this.grid.focusKeys();
  }
  syncSelection() {
    const g = this.grid.range, s = this.grid.sheet, { ar, ac } = this.grid.sel;
    const nb = $<HTMLInputElement>('#namebox');
    if (document.activeElement !== nb) {
      const fullC = g.r1 === 0 && g.r2 === MAX_ROWS - 1, fullR = g.c1 === 0 && g.c2 === MAX_COLS - 1;
      nb.value = fullC && fullR ? toA1(ar, ac) : g.r1 === g.r2 && g.c1 === g.c2 ? toA1(ar, ac) : this.wb.mergeAt(s, ar, ac) && this.grid.isSingle() ? toA1(ar, ac) : toA1(ar, ac);
    }
    if (!this.ed.isEditing()) {
      const fx = $<HTMLTextAreaElement>('#fx-input');
      let t = this.wb.rawText(s, ar, ac), ghost = false;
      // ô nằm trong vùng tràn của mảng động: hiện công thức ô gốc màu xám như Excel
      if (t === '') { const sp = this.wb.spillRange(s, ar, ac); if (sp && (sp.r !== ar || sp.c !== ac)) { t = this.wb.rawText(s, sp.r, sp.c); ghost = true; } }
      fx.value = t; fx.classList.toggle('ghost', ghost);
    }
    this.refreshToggles();
    this.syncTableTab();
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
    this.navFill();
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
  scheduleSave() {
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
  /** Tạo sổ mới từ mẫu (Tệp ▸ Trang đầu / Mới) */
  newFromTemplate(id: string) { try { sessionStorage.setItem('xl-template', id); } catch { this.toast('Không tạo được từ mẫu (bộ nhớ trình duyệt bị chặn).', 'err'); return; } this.newBook(); }
  async newBook() {
    closeFind();
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
    const items: [string, string, (() => void) | null][] = [['paste', 'Dán (P)', () => this.pasteBtn()], ['pasteValues', 'Giá trị (V)', () => this.pasteSpecialRun({ what: 'values' })], ['pasteFormula', 'Công thức (F)', () => this.pasteSpecialRun({ what: 'formulas' })], ['transpose', 'Chuyển vị (T)', () => this.pasteSpecialRun({ what: 'all', transpose: true })], ['pasteFmt', 'Định dạng (R)', () => this.pasteSpecialRun({ what: 'formats' })], ['link', 'Dán Liên kết (N)', () => this.pasteSpecialRun({ link: true })]];
    for (const [icon, label, act] of items) {
      const b = document.createElement('button');
      b.innerHTML = ic(icon, 16); b.title = label; b.setAttribute('aria-label', b.title);
      b.onclick = () => { closeMenu(); act?.(); };
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
      { label: 'Dán Đặc biệt…', hint: 'Ctrl+Alt+V', act: () => this.pasteSpecialOpen() },
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
        { label: 'Định dạng Ô…', icon: 'cellFmt', hint: 'Ctrl+1', act: () => formatCellsDialog(this, 'num') },
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
        { label: this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).cm ? 'Trả lời Chú thích' : 'Chú thích Mới', icon: 'comment', act: () => this.notes.openCard() },
        { label: this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).nt ? 'Sửa Ghi chú' : 'Ghi chú Mới', icon: 'note', act: () => this.notes.editNote() },
        ...(this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).nt ? [{ label: 'Xóa Ghi chú', act: () => this.act(() => this.eachArea((a) => this.wb.clearNotes(this.grid.sheet, a, 'notes'))) }] : []),
        'sep',
        { label: 'Định dạng Ô…', icon: 'cellFmt', hint: 'Ctrl+1', act: () => formatCellsDialog(this, 'num') },
        { label: 'Chọn từ Danh sách Thả xuống…', act: () => this.pickFromList() },
        { label: 'Xác định Tên…', icon: 'tag', act: () => nameDialog(this) },
        { label: this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).ln ? 'Sửa Liên kết…' : 'Liên kết', icon: 'link', hint: 'Ctrl+K', act: () => hyperlinkDialog(this) },
        ...(this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).ln ? [{ label: 'Mở Liên kết', act: () => this.followLink(this.grid.sel.ar, this.grid.sel.ac) }, { label: 'Loại bỏ Liên kết', act: () => this.setLink(null, null) }] as MenuItem[] : []),
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
  applyCellStyle(n: string, sty: Style, bk?: BorderKind, bs?: BorderStyle, bc?: string) {
    const s = this.grid.sheet;
    // Kiểu Ô thay hẳn phông/màu/viền cũ như Excel (không cộng dồn với kiểu trước)
    const reset: Style = { b: undefined, i: undefined, u: undefined, ud: undefined, s: undefined, fg: undefined, bg: undefined, fs: undefined, fn: undefined };
    this.safe(() => this.wb.group(() => { for (const a of this.grid.ranges()) { this.wb.setStyle(s, a, { ...reset, ...sty }); this.wb.setBorder(s, a, 'none'); if (bk) this.wb.setBorder(s, a, bk, bs, bc); } }));
  }
  cellStyleMenu(r?: DOMRect) {
    const box = document.createElement('div');
    box.className = 'gal-pop';
    box.innerHTML = '<h4>Tốt, Xấu và Trung tính</h4><div class="gal-grid" data-g="0"></div><h4>Dữ liệu và Mô hình</h4><div class="gal-grid" data-g="1"></div><h4>Tựa đề và Đầu đề</h4><div class="gal-grid" data-g="2"></div>';
    const groups = [[0, 1, 2, 3], [4, 5, 6, 7, 8], [9, 10, 11, 12, 13]];
    groups.forEach((ix, g) => {
      const host = box.querySelector(`[data-g="${g}"]`)!;
      for (const i of ix) {
        const [n, sty, bk, bs, bc] = CELL_STYLES[i];
        if (!n) continue;
        const b = document.createElement('button');
        b.className = 'gal-it'; b.textContent = n; b.style.cssText = styleCss(sty, bk, bc);
        b.onclick = () => { closeMenu(); this.applyCellStyle(n, sty, bk, bs, bc); };
        host.appendChild(b);
      }
    });
    const rr = r ?? document.querySelector('[data-id="cstyle"]')?.getBoundingClientRect() ?? $('#gridwrap').getBoundingClientRect();
    openPopover(box, rr.left, rr.bottom + 2);
  }
  /** Hàm dùng gần đây (Công thức ▸ Dùng Gần đây), lưu trên máy */
  recentFns(): string[] {
    try { const v = JSON.parse(localStorage.getItem('xl-recent-fn') || 'null'); if (Array.isArray(v) && v.length) return v.slice(0, 10); } catch { /* bỏ qua */ }
    return RECENT_DEF;
  }
  insertFn(f: string) {
    try { localStorage.setItem('xl-recent-fn', JSON.stringify([f, ...this.recentFns().filter((x) => x !== f)].slice(0, 10))); } catch { /* bỏ qua */ }
    this.ed.begin('=' + f + '(', 'enter');
  }
  /** Xem lại ▸ Kiểm tra Hiệu suất: tìm ô có định dạng nhưng nằm ngoài vùng dữ liệu (như Excel) */
  checkPerformance() {
    const s = this.grid.sheet;
    let extra = 0;
    const u = this.wb.used(s);
    for (const [k, st] of this.wb.meta(s).styles) { const kk = +k, r = Math.floor(kk / MAX_COLS), c = kk % MAX_COLS; if ((r >= u.rows || c >= u.cols) && Object.keys(st).length) extra++; }
    this.toast(extra ? `Kiểm tra Hiệu suất: ${extra} ô có định dạng nằm ngoài vùng dữ liệu (vùng dữ liệu: ${u.rows ? 'A1:' + toA1(u.rows - 1, u.cols - 1) : 'trống'}). Xóa định dạng thừa để tệp nhẹ hơn.` : 'Kiểm tra Hiệu suất: không có ô định dạng thừa trong trang tính này.', extra ? undefined : 'ok');
  }
  /** Xem lại ▸ Kiểm tra Trợ năng */
  checkAccessibility() {
    const issues: string[] = [];
    for (const id of this.wb.sheetIds()) { const nm = this.wb.sheetName(id); if (/^(Sheet|Trang tính)\d+$/.test(nm)) { issues.push(`Đặt tên rõ nghĩa cho trang tính “${nm}”`); break; } }
    const ms = this.wb.meta(this.grid.sheet).merges.length;
    if (ms) issues.push(`${ms} ô đã phối`);
    const n = issues.length;
    const sa = document.querySelector('#st-acc span'); if (sa) sa.textContent = n ? 'Trợ năng: Cần xem xét' : 'Trợ năng: Mọi thứ đều ổn';
    this.toast(n ? `Trợ năng: ${n} gợi ý — ${issues.map(tr).join('; ')}.` : 'Trợ năng: Mọi thứ đều ổn.', n ? undefined : 'ok');
    this.updateFlags();
  }
  /** Xem ▸ Dẫn hướng: ngăn liệt kê trang tính và tên vùng */
  navPane() {
    const old = document.getElementById('nav-pane');
    if (old) { old.remove(); this.ribbon.refresh(); return; }
    const p = document.createElement('aside');
    p.id = 'nav-pane'; p.className = 'side-pane'; p.setAttribute('aria-label', 'Dẫn hướng');
    p.innerHTML = `<header><b>Dẫn hướng</b><button class="x" aria-label="Đóng">${ic('close', 14)}</button></header><div class="nv-list"></div>`;
    p.querySelector('.x')!.addEventListener('click', () => this.navPane());
    p.querySelector('.nv-list')!.addEventListener('click', (e) => {
      const b = (e.target as HTMLElement).closest('button');
      const id = b ? +b.dataset.s! : NaN;
      if (b && this.wb.visibleSheetIds().includes(id)) { this.grid.setSheet(id); this.renderSheets(); this.syncSelection(); }
    });
    $('#gridwrap').appendChild(p);
    this.navFill();
    this.ribbon.refresh();
  }
  /** dựng lại danh sách trong ngăn Dẫn hướng (thêm/xóa/đổi tên/ẩn trang tính) */
  navFill() {
    const list = document.querySelector('#nav-pane .nv-list');
    if (!list) return;
    const ids = this.wb.visibleSheetIds();
    const key = ids.map((id) => id + ':' + this.wb.sheetName(id)).join('|') + '#' + this.grid.sheet;
    if ((list as HTMLElement).dataset.k === key) return;
    (list as HTMLElement).dataset.k = key;
    list.innerHTML = '';
    for (const id of ids) {
      const b = document.createElement('button');
      b.dataset.s = String(id);
      if (id === this.grid.sheet) b.setAttribute('aria-current', 'true');
      b.innerHTML = ic('sheetView', 15) + '<span class="nv-name"></span>';
      b.querySelector('.nv-name')!.textContent = this.wb.sheetName(id);
      list.appendChild(b);
    }
  }
  funcPicker() {
    const box = document.createElement('div');
    box.className = 'fn-picker';
    box.innerHTML = `<label class="fn-search">${ic('fx', 16)}<input id="fn-q" placeholder="Tìm hàm, ví dụ: tra cứu, SUMIF" aria-label="Tìm hàm"></label><div class="fn-list" role="listbox"></div>`;
    const list = box.querySelector('.fn-list') as HTMLElement;
    const all = this.wb.hf.getRegisteredFunctionNames().filter((n) => !/^(XOFS\d|XTHISROW)$/.test(n)).sort();
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
      this.insertFn(b.dataset.n!);
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
      { label: 'Định dạng Số Khác…', act: () => formatCellsDialog(this, 'num') },
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
      { label: 'Độ rộng Mặc định…', act: () => { const s = this.grid.sheet, d = this.wb.meta(s).colW.def; this.numDlg('Độ rộng Chuẩn', 'Độ rộng cột chuẩn (ký tự):', String(Math.round(((d - 5) / 7) * 100) / 100).replace('.', ','), (n) => { if (n < 0 || n > 255) return 'Độ rộng cột phải từ 0 đến 255.'; this.safe(() => this.wb.setDefaultColWidth(s, Math.round(n * 7 + 5))); this.grid.layout(); }); } },
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
      { label: 'Định dạng Ô…', icon: 'cellFmt', hint: 'Ctrl+1', act: () => formatCellsDialog(this, 'num') },
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
  clearAll() { const s = this.grid.sheet, g = this.grid.range; this.safe(() => { this.wb.clear(s, g); this.wb.clearNotes(s, g); this.wb.setStyle(s, g, { b: undefined, i: undefined, u: undefined, s: undefined, fg: undefined, bg: undefined, al: undefined, fmt: undefined, fn: undefined, fs: undefined, wr: undefined, va: undefined, bd: undefined, ind: undefined }); }); }
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
    if (keys) { keysDialog(this); return; }
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

  /** Enter khi không sửa: Tab rồi Enter về cột bắt đầu */
  enterNav(shift: boolean) {
    const g = this.grid, [dr, dc] = this.enterMove(shift);
    if (this.tabStart !== null && dr > 0 && !dc && g.isSingle() && !g.areas.length) { const c = this.tabStart; this.tabStart = null; g.select(g.sel.ar + 1, c); return; }
    this.tabStart = null;
    if (dr || dc) g.moveActive(dr, dc);
  }
  /** Hướng di chuyển sau khi nhấn Enter (Tệp ▸ Tùy chọn ▸ Nâng cao) */
  enterMove(shift: boolean): [number, number] {
    const d = lsGet('xl-enter') ?? 'down';
    const v: [number, number] = d === 'right' ? [0, 1] : d === 'left' ? [0, -1] : d === 'up' ? [-1, 0] : d === 'none' ? [0, 0] : [1, 0];
    return shift ? [-v[0], -v[1]] : v;
  }
  // =====================================================================
  // ---------- 0.6: phím tắt giống hệt Excel 365 + tính năng đi kèm ----------
  // =====================================================================
  /** Chế độ End (End rồi mũi tên) */
  endMode = false;
  /** Bật/tắt chú thích hàm (Ctrl+Alt+P) */
  fTips = true;
  /** Lệnh cuối để F4 / Ctrl+Y lặp lại */
  private lastAct: (() => void) | null = null;
  /** Chổi định dạng: kiểu ô nguồn và có dùng nhiều lần không */
  private painter: { styles: (Style | undefined)[][]; h: number; w: number; sticky: boolean } | null = null;
  /** Ô bắt đầu chuỗi Tab (Tab rồi Enter về cột bắt đầu) */
  tabStart: number | null = null;
  /** nhóm trang tính (Ctrl+Shift+PgDn/PgUp) */
  group: number[] = [];

  /** Ghi nhớ lệnh để lặp (F4/Ctrl+Y) và chạy luôn */
  private act(f: () => void) { this.lastAct = f; f(); }
  repeat() { if (this.lastAct) { this.lastAct(); this.grid.invalidate(); this.syncSelection(); } }

  /** Phím dùng chung (bắt ở document) — trả true nếu đã xử lý */
  keyGlobal(e: KeyboardEvent): boolean {
    const t = e.target as HTMLElement | null;
    if (t?.closest?.('.fr-panel, .qa, .clip-pane') && !(e.key === 'F1')) return false;
    if (document.querySelector('.dlg-back') || menuOpen()) return false;
    const inField = !!t && (/^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName) || t.isContentEditable);
    const combo = comboOf(e);
    const def = keyFor(combo, 'global', inField);
    if (!def) return false;
    if (this.ed.isEditing() && !['save', 'print', 'help', 'fxTall', 'fTips', 'ribbonToggle', 'tell', 'fullScreen', 'winMin', 'winMax', 'keytips', 'ribbonTip'].includes(def.id)) return false;
    const ok = this.runKey(def.id, e);
    if (ok) { e.preventDefault(); e.stopPropagation(); }
    return ok;
  }
  /** Phím khi lưới đang nhận phím (không sửa ô) */
  keyNav(e: KeyboardEvent): boolean {
    const combo = comboOf(e);
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const g = this.grid;
    if (arrows[e.key] || ['Home', 'End', 'PageUp', 'PageDown', 'Escape'].includes(e.key)) this.tabStart = null;
    // Chế độ End: End rồi mũi tên = Ctrl+mũi tên; End rồi Home = ô cuối
    if (this.endMode && arrows[e.key] && !e.ctrlKey && !e.altKey) {
      e.preventDefault(); this.endMode = false; this.updateFlags();
      g.jump(arrows[e.key][0], arrows[e.key][1], e.shiftKey || g.extMode);
      return true;
    }
    if (this.endMode && e.key === 'Home') { e.preventDefault(); this.endMode = false; this.updateFlags(); const lc = this.wb.lastCell(g.sheet); g.select(lc.r, lc.c, e.shiftKey); return true; }
    if (this.endMode && !['End', 'Shift', 'Control', 'Alt'].includes(e.key)) { this.endMode = false; this.updateFlags(); }
    // F8: mũi tên mở rộng vùng chọn
    if (g.extMode && arrows[e.key] && !e.altKey) {
      e.preventDefault();
      if (e.ctrlKey) g.jump(arrows[e.key][0], arrows[e.key][1], true); else g.step(arrows[e.key][0], arrows[e.key][1], true);
      return true;
    }
    if (combo === 'F4') { e.preventDefault(); this.repeat(); return true; }
    if (e.key === 'Escape' && (g.extMode || g.addMode || this.painter)) { e.preventDefault(); g.extMode = g.addMode = false; this.painterOff(); this.updateFlags(); return true; }
    const def = keyFor(combo, 'nav', true);
    if (!def) return false;
    const ok = this.runKey(def.id, e);
    if (ok) e.preventDefault();
    return ok;
  }

  /** Thực thi một phím theo id trong bảng ánh xạ */
  runKey(id: string, e: KeyboardEvent): boolean {
    const g = this.grid, s = g.sheet, { ar, ac } = g.sel;
    const nf = (f: string | undefined) => this.act(() => this.applyStyle({ fmt: f }));
    switch (id) {
      // ---- tệp ----
      case 'save': case 'saveF12': if (this.ed.isEditing()) this.ed.commit(0, 0); this.save(); return true;
      case 'open': case 'openF12': this.openFile(); return true;
      case 'new': this.newBook(); return true;
      case 'saveAs': if (this.ed.isEditing()) this.ed.commit(0, 0); this.saveXlsx(true); return true;
      case 'print': if (this.ed.isEditing()) this.ed.commit(0, 0); this.printDlg(); return true;
      case 'closeBook': case 'closeWin': this.closeBook(); return true;
      case 'winMin': case 'winMax': { const d = desk() as any; if (d?.win) { d.win(id === 'winMin' ? 'min' : 'max'); return true; } return false; }
      // ---- hoàn tác ----
      case 'undo': this.undo(); return true;
      case 'redo': if (this.wb.canRedo()) this.redo(); else this.repeat(); return true;
      // ---- ribbon ----
      case 'tell': this.focusTell(); return true;
      case 'ribbonToggle': this.toggleRibbon(); return true;
      case 'ribbonTip': { const el = document.activeElement as HTMLElement | null; if (el?.closest('#ribbon') && el.title) { this.toast(el.title); return true; } return false; }
      case 'help': keysDialog(this); return true;
      case 'newSheet': this.addSheet(); return true;
      case 'fullScreen': $('#app').classList.toggle('full-screen'); this.grid.layout(); return true;
      case 'fxTall': $('#app').classList.toggle('fx-tall'); this.grid.layout(); return true;
      case 'fTips': this.fTips = !this.fTips; this.toast(this.fTips ? 'Đã bật chú thích hàm khi gõ công thức.' : 'Đã tắt chú thích hàm khi gõ công thức.'); return true;
      case 'pane': this.cyclePane(e.shiftKey ? -1 : 1); return true;
      case 'nextSheet': case 'prevSheet': if (this.ed.isEditing()) return false; this.stepSheet(id === 'nextSheet' ? 1 : -1); return true;
      case 'zoomIn': this.zoom(1); return true;
      case 'zoomOut': this.zoom(-1); return true;
      // ---- định dạng ----
      case 'bold': this.act(() => this.toggle('b')); return true;
      case 'italic': this.act(() => this.toggle('i')); return true;
      case 'under': this.act(() => this.toggle('u')); return true;
      case 'strike': this.act(() => this.toggle('s')); return true;
      case 'fmtCells': formatCellsDialog(this, 'num'); return true;
      case 'fmtFont': formatCellsDialog(this, 'font'); return true;
      case 'nfGeneral': nf(NF_SHORTCUT.general); return true;
      case 'nfCurrency': nf(NF_SHORTCUT.currency); return true;
      case 'nfPercent': nf(NF_SHORTCUT.percent); return true;
      case 'nfSci': nf(NF_SHORTCUT.sci); return true;
      case 'nfDate': nf(NF_SHORTCUT.date); return true;
      case 'nfTime': nf(NF_SHORTCUT.time); return true;
      case 'nfNumber': nf(NF_SHORTCUT.number); return true;
      case 'outlineBd': this.act(() => this.eachArea((a) => this.wb.setBorder(s, a, 'outer'))); return true;
      case 'noBd': this.act(() => this.eachArea((a) => this.wb.setBorders(s, a, { t: null, b: null, l: null, r: null, h: null, v: null }))); return true;
      case 'showF': g.showFormulas = !g.showFormulas; g.invalidate(); this.ribbon.refresh(); return true;
      case 'objects': g.objMode = (g.objMode + 1) % 3; this.charts.setMode(g.objMode); this.toast(['Hiện đối tượng.', 'Hiện chỗ giữ cho đối tượng.', 'Ẩn đối tượng.'][g.objMode]); return true;
      case 'hideRows': this.act(() => this.hideSel('r', true)); return true;
      case 'hideCols': this.act(() => this.hideSel('c', true)); return true;
      case 'unhideRows': this.hideSel('r', false); return true;
      case 'unhideCols': this.hideSel('c', false); return true;
      case 'insertDlg': this.insertKey(); return true;
      case 'deleteDlg': this.deleteKey(); return true;
      case 'link': hyperlinkDialog(this); return true;
      case 'quick': quickAnalysis(this); return true;
      case 'stats': this.bookStats(); return true;
      case 'filter': this.toggleFilter(); return true;
      // ---- dán ----
      case 'pasteSp': this.pasteSpecialOpen(); return true;
      case 'pasteVals': this.pasteSpecialRun({ what: 'values' }); return true;
      case 'del': this.act(() => this.clearSel()); return true;
      case 'ctxMenu': { const r = g.cellRectM(ar, ac), b = g.scroller.getBoundingClientRect(); this.contextMenu(b.left + r.x + r.w / 2, b.top + r.y + r.h / 2, 'cell'); return true; }
      // ---- chọn ----
      case 'selAll': this.selectAllKey(); return true;
      case 'selCol': g.selectRange({ r1: 0, c1: g.range.c1, r2: MAX_ROWS - 1, c2: g.range.c2 }); g.setActive(ar, ac); return true;
      case 'selRow': g.selectRange({ r1: g.range.r1, c1: 0, r2: g.range.r2, c2: MAX_COLS - 1 }); g.setActive(ar, ac); return true;
      case 'region': g.selectRange(this.wb.currentRegion(s, ar, ac)); return true;
      case 'extMode': g.extMode = !g.extMode; g.addMode = false; this.updateFlags(); return true;
      case 'addMode': g.addMode = !g.addMode; g.extMode = false; this.updateFlags(); return true;
      case 'endMode': this.endMode = !this.endMode; this.updateFlags(); return true;
      case 'lastCell': { const lc = this.wb.lastCell(s); g.select(lc.r, lc.c); return true; }
      case 'lastCellExt': { const lc = this.wb.lastCell(s); g.select(lc.r, lc.c, true); return true; }
      case 'top': { const fz = this.wb.meta(s).freeze; g.select(fz.rows, fz.cols); return true; }
      case 'topExt': g.select(0, 0, true); return true;
      case 'home': g.select(ar, 0); return true;
      case 'homeExt': g.select(g.sel.fr, 0, true); return true;
      case 'pgRight': case 'pgLeft': { const n = Math.max(1, g.visibleCols().length - 1); g.select(ar, Math.max(0, ac + (id === 'pgRight' ? n : -n))); return true; }
      case 'pgExt': { const n = Math.max(1, g.visibleRows().length - 2) * (e.key === 'PageDown' ? 1 : -1); g.select(Math.max(0, g.sel.fr + n), g.sel.fc, true); return true; }
      case 'corner': this.nextCorner(); return true;
      case 'collapse': g.select(ar, ac); return true;
      case 'showActive': g.ensureVisible(ar, ac); g.invalidate(); return true;
      case 'visible': return this.selectSpecial({ kind: 'visible' }, true);
      case 'prec': return this.selectSpecial({ kind: 'precedents' }, true);
      case 'precAll': return this.selectSpecial({ kind: 'precedents', all: true }, true);
      case 'dep': return this.selectSpecial({ kind: 'dependents' }, true);
      case 'depAll': return this.selectSpecial({ kind: 'dependents', all: true }, true);
      case 'rowDiff': return this.selectSpecial({ kind: 'rowDiff', active: { r: ar, c: ac } }, true);
      case 'colDiff': return this.selectSpecial({ kind: 'colDiff', active: { r: ar, c: ac } }, true);
      case 'selNotes': return this.selectSpecial({ kind: 'notes' }, true);
      case 'grpNext': case 'grpPrev': this.groupStep(id === 'grpNext' ? 1 : -1); return true;
      case 'dvList': this.altDown(); return true;
      case 'shapes': if (!this.charts.cycle(1)) this.toast('Trang tính không có đối tượng nổi.'); return true;
      // ---- dữ liệu ----
      case 'goto': gotoDialog(this); return true;
      case 'find': findDialog(this, 'find'); return true;
      case 'replace': findDialog(this, 'replace'); return true;
      case 'findNext': findAgain(this, 1); return true;
      case 'findPrev': findAgain(this, -1); return true;
      case 'flash': this.flashFill(); return true;
      case 'fillDown': this.act(() => this.fillDir('down')); return true;
      case 'fillRight': this.act(() => this.fillDir('right')); return true;
      case 'autoSum': this.autoSum(); return true;
      case 'calcSheet': case 'calcFull': this.recalc(); return true;
      case 'chartHere': this.insertChart(); return true;
      case 'chartSheet': this.chartSheet(); return true;
      case 'refresh': case 'refreshAll': this.refreshPivots(id === 'refresh' ? s : null); return true;
      case 'today': case 'nowTime': { const t = new Date(); this.ed.begin(id === 'today' ? `${pad2(t.getDate())}/${pad2(t.getMonth() + 1)}/${t.getFullYear()}` : `${t.getHours()}:${pad2(t.getMinutes())}`, 'enter'); return true; }
      case 'fAbove': case 'vAbove': {
        if (ar === 0) return true;
        const txt = id === 'fAbove' ? this.wb.rawText(s, ar - 1, ac) : this.wb.display(s, ar - 1, ac).text;
        this.ed.begin(txt, 'edit'); return true;
      }
      case 'insFn': this.funcPicker(); return true;
      case 'pasteName': pasteNameDialog(this, false); return true;
      case 'note': this.notes.editNote(); return true;
      case 'comment': this.notes.openCard(); return true;
      case 'nameMgr': nameManager(this); return true;
      case 'table': createTableDialog(this); return true;
      case 'outline': this.grid.showOutline = !this.grid.showOutline; this.grid.layout(); return true;
      case 'groupSel': groupSel(this); return true;
      case 'ungroupSel': ungroupSel(this); return true;
      case 'createNames': createFromSelection(this); return true;
      case 'calcAll': this.recalc(); return true;
      case 'editBar': this.ed.begin(null, 'edit', false, 'bar'); return true;
    }
    return false;
  }

  /** Áp một thao tác cho mọi khối đang chọn (vùng chọn nhiều khối) — một bước hoàn tác */
  eachArea(f: (g: Range) => void) {
    const list = this.grid.ranges();
    this.safe(() => this.wb.group(() => { for (const a of list) f(a); }));
    this.grid.invalidate();
  }
  /** Sau khi thay đổi nội dung từ hộp thoại */
  afterEdit() { this.grid.invalidate(); this.syncSelection(); }
  /** Chuyển tới một ô (có thể ở trang khác) */
  gotoCell(sh: number, r: number, c: number, select = true) {
    if (sh !== this.grid.sheet) { this.grid.setSheet(sh); this.renderSheets(); }
    if (select) this.grid.select(r, c);
    this.grid.ensureVisible(r, c);
    this.syncSelection();
  }
  /** Ô/vùng gõ trong Đi tới hoặc Hộp Tên: A1, B2:C9, Tờ 2!A1 */
  gotoRef(t: string): boolean {
    const m = /^(?:'?(.+?)'?!)?(.+)$/.exec(t.trim());
    if (!m) return false;
    let sh = this.grid.sheet;
    if (m[1]) { const id = this.wb.sheetIdByName(m[1].replace(/''/g, "'")); if (id === undefined) return false; sh = id; }
    const g = parseRange(m[2].replace(/\$/g, '').toUpperCase());
    if (!g) return false;
    if (sh !== this.grid.sheet) { this.grid.setSheet(sh); this.renderSheets(); }
    this.grid.selectRange(g);
    this.syncSelection();
    return true;
  }
  /** Đi tới Đặc biệt / các phím chọn — false nếu không có ô nào */
  selectSpecial(o: import('../core/find').SpecialOpts, fromKey = false): boolean {
    const g = this.grid, s = g.sheet;
    const act = { r: g.sel.ar, c: g.sel.ac };
    const res = goSpecial(this.wb, s, g.ranges(), { ...o, active: o.active ?? act });
    if (!res.length) { this.toast('Không tìm thấy ô nào.', fromKey ? '' : 'err'); return fromKey; }
    if (res.length > 4000) res.splice(4000);
    g.setAreas(res);
    // ô hiện hành giữ nguyên nếu nằm trong kết quả (như Excel)
    const last = res[res.length - 1];
    void last;
    this.syncSelection();
    return true;
  }
  /** Ctrl+A: vùng dữ liệu hiện tại → cả trang tính */
  private selectAllKey() {
    const g = this.grid, s = g.sheet, { ar, ac } = g.sel;
    const reg = this.wb.currentRegion(s, ar, ac), cur = g.range;
    const empty = reg.r1 === reg.r2 && reg.c1 === reg.c2 && this.wb.rawText(s, ar, ac) === '';
    const same = cur.r1 === reg.r1 && cur.r2 === reg.r2 && cur.c1 === reg.c1 && cur.c2 === reg.c2;
    if (empty || same || g.areas.length) g.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 });
    else { g.selectRange(reg); g.setActive(ar, ac); }
    this.syncSelection();
  }
  /** Ctrl+. : ô hiện hành sang góc kế tiếp (theo chiều kim đồng hồ) */
  private nextCorner() {
    const g = this.grid, r = g.range, { ar, ac } = g.sel;
    const corners: [number, number][] = [[r.r1, r.c1], [r.r1, r.c2], [r.r2, r.c2], [r.r2, r.c1]];
    const i = corners.findIndex(([y, x]) => y === ar && x === ac);
    const [nr, nc] = corners[(i + 1) % 4];
    g.sel = { ar: nr, ac: nc, fr: nr === r.r1 ? r.r2 : r.r1, fc: nc === r.c1 ? r.c2 : r.c1 };
    g.ensureVisible(nr, nc); g.invalidate(); this.syncSelection();
  }
  /** Ctrl+Shift++ : cả hàng/cột thì chèn luôn, còn lại hỏi như Excel */
  private insertKey() {
    const g = this.grid.range;
    if (g.c1 === 0 && g.c2 === MAX_COLS - 1) this.act(() => this.insRows());
    else if (g.r1 === 0 && g.r2 === MAX_ROWS - 1) this.act(() => this.insCols());
    else this.cellsDlg(true);
  }
  private deleteKey() {
    const g = this.grid.range;
    if (g.c1 === 0 && g.c2 === MAX_COLS - 1) this.act(() => this.delRows());
    else if (g.r1 === 0 && g.r2 === MAX_ROWS - 1) this.act(() => this.delCols());
    else this.cellsDlg(false);
  }
  /** Alt+↓: danh sách kiểm tra dữ liệu → nút lọc → chọn từ danh sách */
  private altDown() {
    const g = this.grid, s = g.sheet, { ar, ac } = g.sel;
    const b = g.scroller.getBoundingClientRect(), r = g.cellRectM(ar, ac);
    if (this.wb.dvAt(s, ar, ac)?.type === 'list') { this.dvPick(b.left + r.x, b.top + r.y + r.h); return; }
    if (totalFnMenu(this, b.left + r.x + r.w, b.top + r.y + r.h)) return;
    const f = this.wb.meta(s).filter;
    if (f && ar === f.range.r1 && ac >= f.range.c1 && ac <= f.range.c2) { this.filterPopover(ac, b.left + r.x + r.w, b.top + r.y + r.h); return; }
    this.pickFromList();
  }
  /** F6 / Shift+F6: trang tính → Ribbon → thu phóng */
  private cyclePane(d: 1 | -1) {
    const a = document.activeElement as HTMLElement | null;
    const panes: (() => HTMLElement | null)[] = [
      () => $('#gridwrap .cell-editor'),
      () => document.querySelector<HTMLElement>('#tabs [aria-selected="true"]') ?? document.querySelector<HTMLElement>('#tabs [role="tab"]'),
      () => $('#zoom-slider'),
    ];
    const cur = a?.closest('#tabs, #ribbon') ? 1 : a?.closest('#status') ? 2 : 0;
    const next = panes[(cur + d + panes.length) % panes.length]();
    if (next === $('#gridwrap .cell-editor')) this.grid.focusKeys(); else next?.focus();
  }
  /** Ctrl+Shift+PgDn/PgUp: nhóm trang tính */
  private groupStep(d: 1 | -1) {
    const vis = this.wb.visibleSheetIds();
    if (!this.group.length) this.group = [this.grid.sheet];
    const ends = this.group.map((x) => vis.indexOf(x));
    const i = d > 0 ? Math.max(...ends) + 1 : Math.min(...ends) - 1;
    if (i < 0 || i >= vis.length) return;
    this.group.push(vis[i]);
    this.renderSheets(); this.updateTitle();
    this.toast(`Đã nhóm ${this.group.length} trang tính — nhập liệu, định dạng, xóa áp cho mọi trang trong nhóm. Bấm vào một thẻ khác để bỏ nhóm.`);
  }
  /** Trang tính đang chịu tác động (nhóm hoặc trang hiện hành) */
  sheetsTarget(): number[] {
    const ids = this.wb.sheetIds();
    this.group = this.group.filter((x) => ids.includes(x));
    return this.group.length > 1 && this.group.includes(this.grid.sheet) ? this.group : [this.grid.sheet];
  }
  /** Ô nhập: ghi vào mọi trang trong nhóm */
  writeInput(r: number, c: number, v: string) {
    const list = this.sheetsTarget();
    let grew = false;
    this.wb.group(() => { for (const sh of list) { this.wb.setInput(sh, r, c, v); if (this.wb.afterEdit(sh, r, c)) grew = true; } });
    if (grew) { this.grid.invalidate(); this.syncTableTab(); }
  }
  /** Truy tìm Ô Tiền lệ / Phụ thuộc: mỗi lần bấm thêm một cấp (như Excel) */
  private traced: { pre: boolean; sheet: number; frontier: { r: number; c: number }[]; seen: Set<string> } | null = null;
  trace(pre: boolean) {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel, wb = this.wb;
    let st = this.traced;
    if (!st || st.pre !== pre || st.sheet !== s || !st.seen.has(ar + ':' + ac)) {
      if (st && (st.pre !== pre || st.sheet !== s)) { /* đổi kiểu truy tìm: giữ mũi tên cũ */ }
      st = this.traced = { pre, sheet: s, frontier: [{ r: ar, c: ac }], seen: new Set([ar + ':' + ac]) };
    }
    const next: { r: number; c: number }[] = [];
    let added = 0;
    for (const p of st.frontier) {
      if (pre) {
        for (const x of wb.precedentsOf(s, p.r, p.c)) {
          if (x.s !== s) { this.grid.arrows.push({ from: x.range, to: p, ext: 'from' }); added++; continue; }
          this.grid.arrows.push({ from: x.range, to: p }); added++;
          for (let r = x.range.r1; r <= Math.min(x.range.r2, x.range.r1 + 200); r++) for (let c = x.range.c1; c <= Math.min(x.range.c2, x.range.c1 + 50); c++) {
            const k = r + ':' + c;
            if (!st.seen.has(k) && wb.rawText(s, r, c).startsWith('=')) { st.seen.add(k); next.push({ r, c }); }
          }
        }
      } else {
        for (const d of wb.directDependents(s, p.r, p.c)) {
          if (d.s !== s) { this.grid.arrows.push({ from: { r1: p.r, c1: p.c, r2: p.r, c2: p.c }, to: p, ext: 'to' }); added++; continue; }
          this.grid.arrows.push({ from: { r1: p.r, c1: p.c, r2: p.r, c2: p.c }, to: { r: d.r, c: d.c } }); added++;
          const k = d.r + ':' + d.c;
          if (!st.seen.has(k)) { st.seen.add(k); next.push({ r: d.r, c: d.c }); }
        }
      }
    }
    st.frontier = next;
    if (!added) this.toast(pre ? 'Ô hiện hành không có ô tiền lệ nào (hoặc đã truy tìm hết).' : 'Không có công thức nào tham chiếu tới ô hiện hành (hoặc đã truy tìm hết).');
    this.grid.invalidate();
  }
  /** hiện/ẩn tab ngữ cảnh Thiết kế Bảng theo ô hiện hành */
  private tblTabOn = false;
  syncTableTab() {
    const on = !!this.wb.tableAt(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac);
    if (on === this.tblTabOn) { if (on) this.ribbon?.refresh(); return; }
    this.tblTabOn = on;
    if (!on && this.tab === 'tdesign') this.tab = 'home';
    this.renderTabs(); this.ribbon?.render(); this.applyRibbon();
  }
  /** Ctrl+Enter: điền nội dung vào mọi ô đang chọn (công thức dời tham chiếu như Excel) */
  fillSelection(v: string) {
    const list: [number, number, string][] = [];
    const base = { r: this.grid.sel.ar, c: this.grid.sel.ac };
    for (const a of this.grid.ranges()) {
      const u = this.wb.used(this.grid.sheet);
      const r2 = a.r2 >= MAX_ROWS - 1 ? Math.max(a.r1, u.rows - 1) : a.r2, c2 = a.c2 >= MAX_COLS - 1 ? Math.max(a.c1, u.cols - 1) : a.c2;
      if ((r2 - a.r1 + 1) * (c2 - a.c1 + 1) > 200_000) { this.toast('Vùng chọn quá lớn (tối đa 200.000 ô).', 'err'); return; }
      for (let r = a.r1; r <= r2; r++) for (let c = a.c1; c <= c2; c++) {
        if (this.wb.mergeCovered(this.grid.sheet, r, c)) continue;
        list.push([r, c, v.startsWith('=') ? shiftFormula(v, r - base.r, c - base.c) : v]);
      }
    }
    this.safe(() => this.wb.group(() => { for (const sh of this.sheetsTarget()) this.wb.writeCells(sh, list); }));
    this.grid.invalidate(); this.syncSelection();
  }

  // ---------- Dán Đặc biệt ----------
  private async pasteSpecialOpen() {
    let text: string | null = null;
    if (!this.wb.clip) { try { text = await navigator.clipboard.readText(); } catch { text = null; } }
    pasteSpecialDialog(this, text);
  }
  pasteSpecialRun(o: PasteOpts) {
    if (!this.wb.clip) { this.pasteBtn(); return; }
    this.safe(() => {
      const g = this.wb.pasteSpecial(this.grid.sheet, this.grid.range, o);
      this.grid.selectRange(g);
      this.lastAct = () => this.safe(() => { const g2 = this.wb.pasteSpecial(this.grid.sheet, this.grid.range, o); this.grid.selectRange(g2); });
    });
    this.grid.layout(); this.syncSelection();
  }
  pasteTextAt(text: string) { this.pasteAt(text); }

  // ---------- Chổi định dạng ----------
  painterOn(sticky: boolean) {
    const g = this.grid.range, s = this.grid.sheet;
    const h0 = Math.min(g.r2 - g.r1 + 1, 500), w0 = Math.min(g.c2 - g.c1 + 1, 200);
    const styles: (Style | undefined)[][] = [];
    for (let i = 0; i < h0; i++) { const row: (Style | undefined)[] = []; for (let j = 0; j < w0; j++) { const st = this.wb.style(s, g.r1 + i, g.c1 + j); row.push(Object.keys(st).length ? { ...st, ln: undefined } : undefined); } styles.push(row); }
    this.painter = { styles, h: h0, w: w0, sticky };
    $('#gridwrap').classList.add('painting');
    this.grid.marquee = { sheet: s, range: g }; this.grid.invalidate();
    this.updateFlags();
  }
  painterToggle() { if (this.painter) this.painterOff(); else this.painterOn(false); }
  painterOff() { if (!this.painter) return; this.painter = null; $('#gridwrap').classList.remove('painting'); this.grid.marquee = null; this.grid.invalidate(); this.updateFlags(); this.ribbon.refresh(); }
  private painterApply() {
    const p = this.painter;
    if (!p) return;
    const s = this.grid.sheet, g = this.grid.range;
    const tgt = g.r1 === g.r2 && g.c1 === g.c2 ? { r1: g.r1, c1: g.c1, r2: g.r1 + p.h - 1, c2: g.c1 + p.w - 1 } : g;
    this.safe(() => this.wb.group(() => this.wb.paintStyles(s, tgt, p.styles)));
    if (!p.sticky) this.painterOff();
    this.grid.invalidate(); this.syncSelection();
  }

  // ---------- Định dạng Ô ----------
  usedFormats(): string[] {
    const set = new Set<string>();
    for (const id of this.wb.sheetIds()) for (const st of this.wb.meta(id).styles.values()) if (st.fmt) set.add(st.fmt);
    return [...set].slice(0, 200);
  }
  applyFormatCells(areas: Range[], patch: Style | null, bspec: Record<string, import('../core/workbook').Border | null> | null, merge: boolean | null) {
    const s = this.grid.sheet;
    const run = () => this.safe(() => this.wb.group(() => {
      for (const sh of this.sheetsTarget()) for (const a of areas) {
        if (patch) this.wb.setStyle(sh, a, patch);
        if (bspec) this.wb.setBorders(sh, a, bspec as any);
      }
      if (merge === true) for (const a of areas) this.wb.merge(s, a, false);
      if (merge === false) for (const a of areas) this.wb.unmerge(s, a);
    }));
    run();
    this.lastAct = () => { const a0 = this.grid.ranges(); this.safe(() => this.wb.group(() => { for (const a of a0) { if (patch) this.wb.setStyle(s, a, patch); if (bspec) this.wb.setBorders(s, a, bspec as any); } })); };
    if (patch?.wr || patch?.fs || patch?.rot) this.fitRows();
    this.grid.invalidate(); this.syncSelection(); this.ribbon.refresh();
  }

  // ---------- Liên kết ----------
  setLink(url: string | null, text: string | null) {
    const s = this.grid.sheet, { ar, ac } = this.grid.sel;
    this.safe(() => this.wb.group(() => {
      if (url) {
        if (text !== null && text !== this.wb.display(s, ar, ac).text) this.wb.setInput(s, ar, ac, text || url);
        else if (this.wb.rawText(s, ar, ac) === '') this.wb.setInput(s, ar, ac, url.replace(/^mailto:/, '').replace(/^#/, ''));
        this.wb.setStyle(s, { r1: ar, c1: ac, r2: ar, c2: ac }, { ln: url, fg: '#0563C1', u: true });
      } else this.wb.setStyle(s, { r1: ar, c1: ac, r2: ar, c2: ac }, { ln: undefined, fg: undefined, u: undefined });
    }));
    this.grid.invalidate(); this.syncSelection();
  }
  /** Mở liên kết của ô (nhấp vào ô có liên kết) */
  followLink(r: number, c: number) {
    const url = this.wb.style(this.grid.sheet, r, c).ln;
    if (!url) return false;
    if (!/^(https?:|mailto:|#)/i.test(url)) { this.toast('Không mở liên kết này vì lý do an toàn: ' + url, 'err'); return true; }
    if (url.startsWith('#')) { if (!this.gotoRef(url.slice(1))) this.toast('Không tìm thấy vị trí liên kết: ' + url.slice(1), 'err'); return true; }
    const d = desk() as any;
    if (d?.openExternal) d.openExternal(url); else window.open(url, '_blank', 'noopener');
    return true;
  }
  clearLinks() { this.act(() => this.eachArea((a) => this.wb.setStyle(this.grid.sheet, a, { ln: undefined, fg: undefined, u: undefined }))); }

  // ---------- Điền Chuỗi / Điền Nhanh ----------
  fillSeries(o: SeriesOpts & { dir: 'rows' | 'cols'; trend?: boolean }): boolean {
    try { this.wb.fillSeries(this.grid.sheet, this.grid.range, o); } catch (e) { this.toast((e as Error).message, 'err'); return false; }
    this.grid.invalidate(); this.syncSelection(); return true;
  }
  /** Ctrl+E: Điền Nhanh theo cột đang chọn, học từ các ô đã gõ ở đầu cột */
  flashFill() {
    const g = this.grid, s = g.sheet, c = g.sel.ac, u = this.wb.used(s);
    // cột nguồn: các cột liền kề bên trái/phải có dữ liệu
    const reg = this.wb.currentRegion(s, g.sel.ar, c);
    const srcCols: number[] = [];
    for (let cc = reg.c1; cc <= reg.c2; cc++) if (cc !== c) srcCols.push(cc);
    if (!srcCols.length) { this.toast('Điền Nhanh cần một cột dữ liệu liền kề làm mẫu.', 'err'); return; }
    let r1 = reg.r1; if (this.wb.guessHeader(s, reg)) r1++;
    const r2 = Math.min(reg.r2, u.rows - 1);
    const rows: string[][] = [], ex: string[] = [], at: number[] = [];
    for (let r = r1; r <= r2; r++) { rows.push(srcCols.map((cc) => this.wb.display(s, r, cc).text)); at.push(r); }
    // ví dụ = các ô đã có ở đầu cột đích
    for (let i = 0; i < at.length; i++) { const t = this.wb.display(s, at[i], c).text; if (!t) break; ex.push(t); }
    if (!ex.length) { this.toast('Hãy gõ một vài ví dụ ở đầu cột rồi bấm Ctrl+E.', 'err'); return; }
    const out = flashFill(rows, ex);
    if (!out) { this.toast('Điền Nhanh không nhận ra mẫu. Hãy gõ thêm ví dụ rồi thử lại.', 'err'); return; }
    const list: [number, number, string][] = [];
    out.forEach((v, i) => { if (i >= ex.length && v !== null && this.wb.rawText(s, at[i], c) === '') list.push([at[i], c, needsApos(v) ? "'" + v : v]); });
    if (!list.length) { this.toast('Không có ô nào để điền.'); return; }
    this.safe(() => this.wb.writeCells(s, list));
    g.selectRange({ r1: at[0], c1: c, r2: at[at.length - 1], c2: c });
    this.toast(`Điền Nhanh: đã điền ${list.length} ô (Ctrl+Z để hoàn tác).`, 'ok');
  }

  // ---------- Phân tích Nhanh ----------
  quickCF(kind: 'bar' | 'scale' | 'gt' | 'top' | 'clear') {
    const s = this.grid.sheet, g = this.grid.range;
    this.safe(() => {
      if (kind === 'clear') this.wb.clearCF(s, g);
      else if (kind === 'bar') this.wb.addCF(s, { range: g, type: 'bar', colors: ['#638EC6'] });
      else if (kind === 'scale') this.wb.addCF(s, { range: g, type: 'scale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] });
      else if (kind === 'gt') { cfRuleDialog(this.wb, s, g, 'gt', () => this.grid.invalidate()); return; }
      else this.wb.addCF(s, { range: g, type: 'top', rank: 10, percent: true, fmt: { bg: '#FFC7CE', fg: '#9C0006' } });
    });
    this.grid.invalidate();
  }
  quickTotal(fn: string, dir: 'row' | 'col') {
    const s = this.grid.sheet, g = this.grid.range;
    const list: [number, number, string][] = [];
    if (dir === 'row') { for (let c = g.c1; c <= g.c2; c++) list.push([g.r2 + 1, c, `=${fn}(${rangeToA1({ r1: g.r1, c1: c, r2: g.r2, c2: c })})`]); }
    else for (let r = g.r1; r <= g.r2; r++) list.push([r, g.c2 + 1, `=${fn}(${rangeToA1({ r1: r, c1: g.c1, r2: r, c2: g.c2 })})`]);
    this.safe(() => this.wb.writeCells(s, list));
    this.grid.invalidate(); this.syncSelection();
  }

  // ---------- biểu đồ trang riêng, làm mới ----------
  private chartSheet() {
    const s = this.grid.sheet, g = this.dataRange();
    if (!g) { this.toast('Hãy chọn vùng dữ liệu (có tiêu đề và số) trước khi tạo biểu đồ.', 'err'); return; }
    chartDialog({ wb: this.wb, sheet: s, range: g, onDone: (ch) => {
      if (!ch) { this.grid.focusKeys(); return; }
      this.safe(() => this.wb.group(() => {
        let i = 1; while (this.wb.sheetIdByName('Biểu đồ' + i) !== undefined) i++;
        const id = this.wb.addSheet('Biểu đồ' + i, this.wb.sheetIds().indexOf(s));
        this.wb.addChart(id, { ...ch, w: 900, h: 560, anchor: { r: 0, c: 0, dx: 8, dy: 8 } });
        this.grid.setSheet(id); this.renderSheets();
      }));
    } });
  }
  private refreshPivots(sheet: number | null) {
    let n = 0;
    this.safe(() => this.wb.group(() => {
      for (const id of sheet === null ? this.wb.sheetIds() : [sheet]) for (const p of [...this.wb.meta(id).pivots]) { this.wb.refreshPivot(id, p.id!); n++; }
    }));
    this.grid.invalidate();
    this.toast(n ? `Đã làm mới ${n} PivotTable.` : 'Không có dữ liệu nào để làm mới.');
  }
  /** Ctrl+W / Ctrl+F4: đóng sổ (hỏi lưu nếu có thay đổi chưa ghi ra tệp) */
  private closeBook() {
    if (this.ed.isEditing()) this.ed.commit(0, 0);
    const go = () => { sessionStorage.setItem('xl-warn-closed', this.wb.name); this.newBook(); };
    if (!(this.fileDirty && this.fileName)) { go(); return; }
    openDialog(`Lưu thay đổi vào “${this.fileName}”?`, h('p', {}, 'Sổ làm việc có thay đổi chưa lưu ra tệp. Bản nháp vẫn được giữ trong mục Gần đây.'), [
      { label: 'Không lưu', kind: 'danger', act: () => go() },
      { label: 'Hủy', act: () => {} },
      { label: 'Lưu', kind: 'primary', act: () => { this.saveXlsx(false).then((ok) => { if (ok) go(); }); } },
    ]);
  }
  /** Chỉ báo trên thanh trạng thái: Mở rộng Vùng chọn / Thêm vào Vùng chọn / Chế độ End / Chổi định dạng */
  updateFlags() {
    let el = document.getElementById('st-flags');
    if (!el) { el = document.createElement('span'); el.id = 'st-flags'; el.className = 'st-flags'; $('#mode').after(el); }
    const g = this.grid;
    el.textContent = [g.extMode ? 'Mở rộng Vùng chọn' : '', g.addMode ? 'Thêm vào Vùng chọn' : '', this.endMode ? 'Chế độ End' : '', this.painter ? 'Chổi định dạng' + (this.painter.sticky ? ' (nhiều lần)' : '') : ''].filter(Boolean).join(' · ');
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
      try { if (this.at.s === this.grid.sheet) this.app.writeInput(this.at.r, this.at.c, v); else this.wb.setInput(this.at.s, this.at.r, this.at.c, v); } catch (e) { this.app.toast((e as Error).message, 'err'); }
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
    const names = this.wb.hf.getRegisteredFunctionNames().filter((n) => !/^(XOFS\d|XTHISROW)$/.test(n));
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
    if (this.editKey(e)) return;
    if (k === 'Enter' && e.altKey) {
      e.preventDefault();
      const el = this.el, p = el.selectionStart ?? el.value.length;
      el.value = el.value.slice(0, p) + '\n' + el.value.slice(el.selectionEnd ?? p);
      el.setSelectionRange(p + 1, p + 1);
      this.onInput(from);
      return;
    }
    if (k === 'Enter') { e.preventDefault(); this.commitEnter(e.shiftKey); return; }
    if (k === 'Tab') { e.preventDefault(); if (this.app.tabStart === null && this.grid.isSingle()) this.app.tabStart = this.at.c; this.commit(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'Escape') { e.preventDefault(); this.cancel(); return; }
    if (k === 'F2') { e.preventDefault(); this.mode = this.mode === 'enter' ? 'edit' : 'enter'; this.app.setMode(this.mode === 'enter' ? 'Nhập' : 'Sửa'); return; }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k] && this.mode === 'enter' && from === 'cell') {
      e.preventDefault();
      if (this.pointMode()) this.movePoint(...arrows[k]);
      else this.commit(...arrows[k]);
    }
  }

  /** Enter khi sửa: Tab rồi Enter về cột bắt đầu (như Excel) */
  private commitEnter(shift: boolean) {
    const ts = this.app.tabStart;
    const [dr, dc] = this.app.enterMove(shift);
    if (ts !== null && dr > 0 && !dc) {
      const r = this.at.r + 1;
      if (!this.commit(0, 0)) return;
      this.app.tabStart = null;
      this.grid.select(r, ts);
      return;
    }
    this.app.tabStart = null;
    this.commit(dr, dc);
  }
  /** Chèn chữ tại con trỏ của ô đang sửa */
  insertText(t: string) {
    const el = this.el, a = el.selectionStart ?? el.value.length, b = el.selectionEnd ?? a;
    el.value = el.value.slice(0, a) + t + el.value.slice(b);
    el.setSelectionRange(a + t.length, a + t.length);
    this.onInput(el === this.ta ? 'cell' : 'bar');
  }
  /** Tên hàm ngay trước con trỏ (=SUM| hoặc =SUM(|) */
  private fnAtCaret(): { name: string; start: number; paren: boolean } | null {
    const el = this.el, p = el.selectionStart ?? el.value.length;
    if (!el.value.startsWith('=')) return null;
    const m = /([A-Za-z][A-Za-z0-9.]*)(\(?)$/.exec(el.value.slice(0, p));
    if (!m || !this.wb.hf.getRegisteredFunctionNames().includes(m[1].toUpperCase())) return null;
    return { name: m[1].toUpperCase(), start: p - m[0].length, paren: !!m[2] };
  }
  /** Phím riêng khi đang sửa ô (F4, Ctrl+A, Ctrl+Shift+A, Ctrl+Enter, F9, Ctrl+; …) */
  private editKey(e: KeyboardEvent): boolean {
    const combo = comboOf(e), el = this.el;
    const done = () => { e.preventDefault(); e.stopPropagation(); return true; };
    switch (combo) {
      case 'F4': {
        if (!el.value.startsWith('=')) return done();
        const r = cycleRef(el.value, el.selectionStart ?? 0, el.selectionEnd ?? 0);
        el.value = r.text; el.setSelectionRange(r.start, r.end);
        this.mirror(el === this.ta ? 'cell' : 'bar'); this.place();
        return done();
      }
      case 'Ctrl+A': {
        const f = this.fnAtCaret();
        if (!f) return false;
        done();
        const before = el.value.slice(0, f.start), after = el.value.slice(el.selectionStart ?? el.value.length);
        funcArgsDialog(this.app, f.name, (t) => {
          el.value = before + t.slice(1) + after.replace(/^\)/, '');
          const p = before.length + t.length - 1;
          this.refocus(); el.setSelectionRange(p, p); this.mirror(el === this.ta ? 'cell' : 'bar');
        });
        return true;
      }
      case 'Ctrl+Shift+A': {
        const f = this.fnAtCaret();
        if (!f) return false;
        const names = argNames(f.name);
        const p = el.selectionStart ?? el.value.length;
        el.value = el.value.slice(0, f.start) + f.name + '(' + names.join(';') + ')' + el.value.slice(p);
        const q = f.start + f.name.length + 1;
        el.setSelectionRange(q, q + (names[0]?.length ?? 0));
        this.mirror(el === this.ta ? 'cell' : 'bar'); this.closeAc();
        return done();
      }
      case 'Ctrl+Enter': {
        done();
        const v = el.value;
        if (this.grid.isSingle() && !this.grid.areas.length) { this.commit(0, 0); return true; }
        const bad = this.wb.validate(this.at.s, this.at.r, this.at.c, v);
        if (bad) { const rule = this.wb.dvAt(this.at.s, this.at.r, this.at.c); this.app.toast(bad, 'err'); if (!rule?.errorStyle || rule.errorStyle === 'stop') { el.select(); return true; } }
        this.finish();
        this.app.fillSelection(v);
        return true;
      }
      case 'Ctrl+Shift+Enter': this.commitEnter(false); return done();
      case 'F9': {
        // tính phần công thức đang bôi đen (như Excel)
        const a = el.selectionStart ?? 0, b = el.selectionEnd ?? 0;
        if (!el.value.startsWith('=')) return done();
        const part = a === b ? el.value.slice(1) : el.value.slice(a, b);
        const v = this.wb.evalFormula('=' + part, this.at.s);
        if (v === undefined) return done();
        const lit = typeof v === 'string' ? (v.startsWith('#') ? v : '"' + v.replace(/"/g, '""') + '"') : typeof v === 'number' ? String(v).replace('.', ',') : String(v).toUpperCase();
        if (a === b) el.value = '=' + lit; else el.value = el.value.slice(0, a) + lit + el.value.slice(b);
        el.setSelectionRange(a === b ? el.value.length : a, a === b ? el.value.length : a + lit.length);
        this.mirror(el === this.ta ? 'cell' : 'bar');
        return done();
      }
      case 'Ctrl+;': { const t = new Date(); this.insertText(`${pad2(t.getDate())}/${pad2(t.getMonth() + 1)}/${t.getFullYear()}`); return done(); }
      case 'Ctrl+Shift+;': { const t = new Date(); this.insertText(`${t.getHours()}:${pad2(t.getMinutes())}`); return done(); }
      case "Ctrl+'": case "Ctrl+Shift+'": {
        if (this.at.r === 0) return done();
        this.insertText(combo === "Ctrl+'" ? this.wb.rawText(this.at.s, this.at.r - 1, this.at.c).replace(/^=/, el.value ? '' : '=') : this.wb.display(this.at.s, this.at.r - 1, this.at.c).text);
        return done();
      }
      case 'Ctrl+Delete': {
        const a = el.selectionStart ?? 0, nl = el.value.indexOf('\n', a);
        el.value = el.value.slice(0, a) + (nl < 0 ? '' : el.value.slice(nl));
        el.setSelectionRange(a, a); this.mirror(el === this.ta ? 'cell' : 'bar');
        return done();
      }
      case 'Ctrl+F2': {
        // chuyển qua lại ô ↔ thanh công thức
        const p = el.selectionStart ?? el.value.length;
        if (this.origin === 'cell') { this.origin = 'bar'; this.ta.readOnly = true; this.ta.classList.add('mirror'); this.fx.focus(); this.fx.setSelectionRange(p, p); }
        else { this.origin = 'cell'; this.ta.readOnly = false; this.ta.classList.remove('mirror'); this.ta.focus({ preventScroll: true }); this.ta.setSelectionRange(p, p); }
        return done();
      }
      case 'F3': done(); pasteNameDialog(this.app, true); return true;
      case 'Shift+F3': {
        const f = this.fnAtCaret();
        if (f) { done(); funcArgsDialog(this.app, f.name, (t) => { el.value = el.value.slice(0, f.start) + t.slice(1); this.refocus(); this.mirror(el === this.ta ? 'cell' : 'bar'); }); return true; }
        return false;
      }
    }
    return false;
  }

  private navKey(e: KeyboardEvent, ctrl: boolean) {
    const g = this.grid, app = this.app, k = e.key;
    if (app.keyNav(e)) return;
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
      case 'Enter': { done(); app.enterNav(e.shiftKey); return; }
      case 'Tab': done(); if (g.isSingle() && !g.areas.length && app.tabStart === null) app.tabStart = g.sel.ac; g.moveActive(0, e.shiftKey ? -1 : 1); return;
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
