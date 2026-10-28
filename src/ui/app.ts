import { MAX_COLS, MAX_ROWS, parseRange, rangeToA1, toA1, type Range } from '../core/address';
import { quoteSheet } from '../core/formula';
import { Workbook, type Style } from '../core/workbook';
import { Grid, type EditHost } from './grid';
import { ic } from './icons';
import { FUNC_INFO } from './funcs';
import { saveBook, listBooks, deleteBook, lsGet, lsSet } from './store';
import { sortDialog, cfRuleDialog, cfManageDialog, dvDialog, h, openDialog, closeDialog } from './dialogs';
import { pickFile, saveBytes, baseName, desk, deskToPicked } from './files';
import type { BorderKind, BorderStyle } from '../core/workbook';
import { ChartLayer, chartDialog } from './charts';
import { PivotPanel, pivotCreateDialog } from './pivotui';
import { printDialog } from './printui';

type Btn = { id: string; icon: string; label: string; act: () => void; on?: () => boolean; later?: string; kbd?: string };
type Group = { label: string; items: Btn[] };

const MENU_IDS = ['open', 'fmt', 'fg', 'bg', 'fnp', 'fn', 'fs', 'border', 'va', 'merge', 'cf', 'sortf', 'csv', 'd-v', 'd-cf', 'd-sort', 'xlsx-open', 'd-open', 'chart', 'pivot', 'pdf', 'i-chart', 'i-pie'];
const $ = <T extends HTMLElement = HTMLElement>(sel: string, root: ParentNode = document) => root.querySelector(sel) as T;
const isTouch = () => matchMedia('(pointer: coarse)').matches;
const COLORS = ['#000000', '#5b6a62', '#c0392b', '#e67e22', '#f1c40f', '#127a4e', '#1f6fb2', '#6c3fb5', '#ffffff', '#fde2e1', '#fff3cd', '#dcefe5', '#dbeafe', '#ede4ff'];
const NUM_FMTS: [string, string | undefined, string][] = [
  ['Chung', undefined, '1250,5'],
  ['Số', '#,##0', '1.251'],
  ['Số thập phân', '#,##0.00', '1.250,50'],
  ['Tiền tệ (₫)', '#,##0" ₫"', '1.251 ₫'],
  ['Phần trăm', '0%', '15%'],
  ['Phần trăm thập phân', '0.00%', '15,25%'],
  ['Ngày', 'dd/mm/yyyy', '28/10/2026'],
  ['Ngày giờ', 'dd/mm/yyyy hh:mm', '28/10/2026 18:30'],
];

export class App {
  wb: Workbook;
  grid!: Grid;
  ed!: CellEditor;
  charts!: ChartLayer;
  pivot!: PivotPanel;
  bookId: string;
  private tab = 'home';
  private saveTimer = 0;
  private installEvt: any = null;

  constructor(wb: Workbook, bookId: string) {
    this.wb = wb;
    this.bookId = bookId;
    this.mount();
  }

  // ---------- dựng giao diện ----------
  private mount() {
    document.querySelectorAll<HTMLElement>('[data-ic]').forEach((el) => el.insertAdjacentHTML('afterbegin', ic(el.dataset.ic!, +(el.dataset.s || 18))));
    this.grid = new Grid($('#gridwrap'), this.wb);
    this.ed = new CellEditor(this);
    this.grid.edit = this.ed;
    this.grid.onSelect = () => this.syncSelection();
    this.grid.onContext = (x, y) => this.cellMenu(x, y);
    this.charts = new ChartLayer($('#gridwrap'), this.wb, this.grid, (id) => this.editChart(id), (m) => this.toast(m));
    this.grid.onDraw = () => this.charts.sync();
    this.pivot = new PivotPanel($('#gridwrap'), this.wb, (m, k) => this.toast(m, k), () => this.grid.invalidate());
    $('#gridwrap').addEventListener('grid-error', (e) => this.toast((e as CustomEvent).detail, 'err'));
    $('#gridwrap').addEventListener('filter-btn', (e) => { const d = (e as CustomEvent).detail; this.filterPopover(d.col, d.x, d.y); });
    $('#gridwrap').addEventListener('dv-btn', (e) => { const d = (e as CustomEvent).detail; this.dvPick(d.x, d.y); });
    const warn = sessionStorage.getItem('xl-warn');
    if (warn) { sessionStorage.removeItem('xl-warn'); setTimeout(() => this.toast('Lưu ý khi mở tệp: ' + warn, 'err'), 400); }
    this.wb.onChange(() => this.onBookChange());
    this.renderTabs();
    this.renderRibbon();
    this.compactQ.addEventListener('change', () => { this.ribbonPeek = false; this.applyRibbon(); });
    // chạm vào lưới thì ẩn ribbon đang tạm hiện
    this.grid.scroller.addEventListener('pointerdown', () => { if (this.ribbonPeek) { this.ribbonPeek = false; this.applyRibbon(); } }, true);
    this.renderSheets();
    this.bindTop();
    this.bindFoot();
    this.bindNameBox();
    this.bindClipboard();
    this.bindTheme();
    this.bindInstall();
    this.bindDesktop();
    this.syncSelection();
    this.grid.focusKeys();
    // lưu ngay khi rời trang / ẩn app; Ctrl+S ở bất kỳ đâu
    const flush = () => { if (this.dirtySave) this.saveNow(false); };
    window.addEventListener('pagehide', flush);
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });
    document.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault(); if (this.ed.isEditing()) this.ed.commit(0, 0);
        // đang mở một tệp .xlsx (bản máy tính hoặc Chrome có quyền ghi) → lưu đè tệp; không thì lưu trong app
        if (this.fileHandle && /\.xlsx$/i.test(this.fileName)) this.saveXlsx(false); else this.saveNow(true);
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') { e.preventDefault(); this.openFile(); }
      const tg = e.target as HTMLElement | null, inField = !!tg && (/^(INPUT|TEXTAREA|SELECT)$/.test(tg.tagName) || tg.isContentEditable);
      if ((e.ctrlKey || e.metaKey) && !inField && !document.querySelector('.dlg-back') && (e.key.toLowerCase() === 'z' || e.key.toLowerCase() === 'y')) {
        e.preventDefault(); if (e.key.toLowerCase() === 'y' || e.shiftKey) this.redo(); else this.undo(); return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'F1') { e.preventDefault(); this.toggleRibbon(); }
      if (e.key === 'Escape' && this.ribbonPeek && !document.querySelector('.dlg-back, .menu')) { this.ribbonPeek = false; this.applyRibbon(); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') { e.preventDefault(); if (this.ed.isEditing()) this.ed.commit(0, 0); this.printDlg(); }
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'l' && !this.ed.isEditing()) { e.preventDefault(); this.toggleFilter(); }
    }, true);
  }
  private dirtySave = false;

  // ---------- thanh trên ----------
  private bindTop() {
    $('#btn-undo').onclick = () => this.undo();
    $('#btn-redo').onclick = () => this.redo();
    $('#btn-save').onclick = () => this.saveNow(true);
    const name = $<HTMLInputElement>('#doc-name');
    name.value = this.wb.name;
    name.onchange = () => { this.wb.name = name.value.trim() || 'Sổ tính'; name.value = this.wb.name; this.scheduleSave(); document.title = this.wb.name + ' · ExcelLike'; };
    name.onkeydown = (e) => { if (e.key === 'Enter' || e.key === 'Escape') { e.preventDefault(); name.blur(); this.grid.focusKeys(); } };
    document.title = this.wb.name + ' · ExcelLike';
  }
  recalc() { this.wb.recalc(); this.grid.invalidate(); this.toast('Đã tính lại toàn bộ sổ tính. Lịch sử hoàn tác được làm mới.'); }
  undo() { if (!this.wb.undo()) this.toast('Không còn thao tác để hoàn tác.'); this.afterHistory(); }
  redo() { if (!this.wb.redo()) this.toast('Không còn thao tác để làm lại.'); this.afterHistory(); }
  private afterHistory() {
    if (!this.wb.sheetIds().includes(this.grid.sheet)) this.grid.setSheet(this.wb.sheetIds()[0]);
    this.renderSheets(); this.grid.invalidate(); this.syncSelection();
  }

  // ---------- ribbon ----------
  private tabs(): Record<string, [string, Group[]]> {
    const s = () => this.grid.sheet;
    const cur = () => this.wb.style(s(), this.grid.sel.ar, this.grid.sel.ac);
    const st = (k: keyof Style) => () => !!cur()[k];
    const al = (a: string) => () => cur().al === a;
    return {
      file: ['Tệp', [
        { label: 'Sổ tính', items: [
          { id: 'new', icon: 'file', label: 'Sổ mới', act: () => this.newBook() },
          { id: 'open', icon: 'folder', label: 'Sổ đã lưu', act: () => this.openMenu() },
          { id: 'save', icon: 'save', label: 'Lưu', kbd: 'Ctrl+S', act: () => this.saveNow(true) },
        ] },
        { label: 'Tệp Excel & CSV', items: [
          { id: 'xlsx-open', icon: 'upload', label: 'Mở tệp', kbd: 'Ctrl+O', act: () => this.openFile() },
          { id: 'xlsx-save', icon: 'download', label: 'Lưu .xlsx', act: () => this.saveXlsx(false) },
          { id: 'xlsx-saveas', icon: 'download', label: 'Lưu thành…', act: () => this.saveXlsx(true) },
          { id: 'csv', icon: 'table', label: 'Xuất CSV', act: () => this.csvMenu() },
          { id: 'pdf', icon: 'printer', label: 'In / PDF', kbd: 'Ctrl+P', act: () => this.printDlg() },
        ] },
        { label: 'Ứng dụng', items: [
          ...(desk() ? [] : [{ id: 'install', icon: 'install', label: 'Cài đặt app', act: () => this.install() }]),
          { id: 'about', icon: 'info', label: 'Giới thiệu', act: () => this.toast('ExcelLike 0.4.1 · Bảng tính kiểu Excel 365: công thức tiếng Việt, .xlsx, CSV, định dạng, sắp xếp, lọc, định dạng có điều kiện, kiểm tra dữ liệu, biểu đồ, PivotTable, in / PDF. Bản Windows và PWA cùng một mã nguồn.') },
        ] },
      ]],
      home: ['Trang chủ', [
        { label: 'Bảng tạm', items: [
          { id: 'paste', icon: 'paste', label: 'Dán', kbd: 'Ctrl+V', act: () => this.pasteBtn() },
          { id: 'cut', icon: 'cut', label: 'Cắt', kbd: 'Ctrl+X', act: () => this.copySel(true) },
          { id: 'copy', icon: 'copy', label: 'Sao chép', kbd: 'Ctrl+C', act: () => this.copySel(false) },
        ] },
        { label: 'Phông chữ', items: [
          { id: 'fn', icon: 'font', label: 'Phông', act: () => this.fontMenu() },
          { id: 'fs', icon: 'fontSize', label: 'Cỡ chữ', act: () => this.sizeMenu() },
          { id: 'b', icon: 'bold', label: 'Đậm', kbd: 'Ctrl+B', act: () => this.toggle('b'), on: st('b') },
          { id: 'i', icon: 'italic', label: 'Nghiêng', kbd: 'Ctrl+I', act: () => this.toggle('i'), on: st('i') },
          { id: 'u', icon: 'underline', label: 'Gạch dưới', kbd: 'Ctrl+U', act: () => this.toggle('u'), on: st('u') },
          { id: 'fg', icon: 'textColor', label: 'Màu chữ', act: () => this.colorMenu('fg') },
          { id: 'bg', icon: 'fill', label: 'Tô nền', act: () => this.colorMenu('bg') },
          { id: 'border', icon: 'borders', label: 'Viền', act: () => this.borderMenu() },
        ] },
        { label: 'Căn lề', items: [
          { id: 'al-l', icon: 'left', label: 'Trái', act: () => this.align('l'), on: al('l') },
          { id: 'al-c', icon: 'center', label: 'Giữa', act: () => this.align('c'), on: al('c') },
          { id: 'al-r', icon: 'right', label: 'Phải', act: () => this.align('r'), on: al('r') },
          { id: 'va', icon: 'vAlign', label: 'Căn dọc', act: () => this.vAlignMenu() },
          { id: 'wr', icon: 'wrap', label: 'Xuống dòng', act: () => this.toggleWrap(), on: st('wr') },
          { id: 'merge', icon: 'merge', label: 'Gộp ô', act: () => this.mergeMenu(), on: () => !!this.wb.mergeAt(s(), this.grid.sel.ar, this.grid.sel.ac) },
        ] },
        { label: 'Số', items: [
          { id: 'fmt', icon: 'table', label: 'Định dạng', act: () => this.fmtMenu() },
          { id: 'thou', icon: 'thousands', label: 'Số 0.000', act: () => this.applyStyle({ fmt: '#,##0' }) },
          { id: 'pct', icon: 'percent', label: 'Phần trăm', act: () => this.applyStyle({ fmt: '0%' }) },
          { id: 'dinc', icon: 'decInc', label: 'Tăng số lẻ', act: () => this.decimals(1) },
          { id: 'ddec', icon: 'decDec', label: 'Giảm số lẻ', act: () => this.decimals(-1) },
        ] },
        { label: 'Kiểu', items: [
          { id: 'cf', icon: 'cf', label: 'Có điều kiện', act: () => this.cfMenu() },
        ] },
        { label: 'Ô', items: [
          { id: 'ins-r', icon: 'rowIns', label: 'Chèn hàng', act: () => this.insRows() },
          { id: 'del-r', icon: 'rowDel', label: 'Xóa hàng', act: () => this.delRows() },
          { id: 'ins-c', icon: 'colIns', label: 'Chèn cột', act: () => this.insCols() },
          { id: 'del-c', icon: 'colDel', label: 'Xóa cột', act: () => this.delCols() },
        ] },
        { label: 'Chỉnh sửa', items: [
          { id: 'sum', icon: 'sigma', label: 'Tổng', kbd: 'Alt+=', act: () => this.autoSum() },
          { id: 'clear', icon: 'eraser', label: 'Xóa nội dung', kbd: 'Delete', act: () => this.clearSel() },
          { id: 'sortf', icon: 'sortAZ', label: 'Sắp xếp & lọc', act: () => this.sortFilterMenu() },
        ] },
      ]],
      insert: ['Chèn', [
        { label: 'Ô', items: [
          { id: 'i-row', icon: 'rowIns', label: 'Hàng phía trên', act: () => this.insRows() },
          { id: 'i-col', icon: 'colIns', label: 'Cột bên trái', act: () => this.insCols() },
          { id: 'i-sheet', icon: 'sheet', label: 'Sheet mới', act: () => this.addSheet() },
        ] },
        { label: 'Phân tích', items: [
          { id: 'pivot', icon: 'pivot', label: 'PivotTable', act: () => this.insertPivot() },
        ] },
        { label: 'Biểu đồ', items: [
          { id: 'chart', icon: 'chart', label: 'Biểu đồ…', act: () => this.insertChart() },
          { id: 'i-chart', icon: 'chartLine', label: 'Đường', act: () => this.insertChart('line') },
          { id: 'i-pie', icon: 'chartPie', label: 'Tròn', act: () => this.insertChart('pie') },
        ] },
      ]],
      formulas: ['Công thức', [
        { label: 'Thư viện hàm', items: [
          { id: 'fnp', icon: 'fx', label: 'Chèn hàm', kbd: 'Shift+F3', act: () => this.funcPicker() },
          { id: 'fsum', icon: 'sigma', label: 'Tổng tự động', act: () => this.autoSum() },
        ] },
        { label: 'Tính toán', items: [
          { id: 'recalc', icon: 'calc', label: 'Tính lại', kbd: 'F9', act: () => this.recalc() },
          { id: 'iter', icon: 'refresh', label: 'Tính lặp', act: () => { this.wb.setIterate({ on: !this.wb.iterate.on }); this.grid.invalidate(); this.renderRibbon(); this.toast(this.wb.iterate.on ? `Đã bật tính lặp (tối đa ${this.wb.iterate.count} lần, sai số ${String(this.wb.iterate.delta).replace('.', ',')}) — tham chiếu vòng được tính như Excel.` : 'Đã tắt tính lặp — ô có tham chiếu vòng hiện #CYCLE!.'); }, on: () => this.wb.iterate.on },
        ] },
      ]],
      data: ['Dữ liệu', [
        { label: 'Lấy dữ liệu', items: [
          { id: 'd-open', icon: 'upload', label: 'Mở tệp / CSV', act: () => this.openFile() },
        ] },
        { label: 'Sắp xếp & Lọc', items: [
          { id: 'd-az', icon: 'sortAZ', label: 'A → Z', act: () => this.quickSort(false) },
          { id: 'd-za', icon: 'sortZA', label: 'Z → A', act: () => this.quickSort(true) },
          { id: 'd-sort', icon: 'sortAZ', label: 'Sắp xếp…', act: () => this.sortDlg() },
          { id: 'd-f', icon: 'filter', label: 'Lọc', kbd: 'Ctrl+Shift+L', act: () => this.toggleFilter(), on: () => !!this.wb.meta(s()).filter },
          { id: 'd-fx', icon: 'filterX', label: 'Xóa lọc', act: () => this.clearFilterConds() },
          { id: 'd-fr', icon: 'calc', label: 'Áp dụng lại', act: () => { this.wb.reapplyFilter(s()); this.grid.layout(); } },
        ] },
        { label: 'Công cụ dữ liệu', items: [
          { id: 'd-v', icon: 'validate', label: 'Kiểm tra dữ liệu', act: () => this.dvDlg() },
          { id: 'd-cf', icon: 'cf', label: 'Có điều kiện', act: () => this.cfMenu() },
        ] },
      ]],
      view: ['Xem', [
        { label: 'Cố định ngăn', items: [
          { id: 'fz', icon: 'freeze', label: 'Tại ô chọn', act: () => this.freeze(this.grid.sel.ar, this.grid.sel.ac) },
          { id: 'fz-r', icon: 'rowIns', label: 'Hàng đầu', act: () => this.freeze(1, 0) },
          { id: 'fz-c', icon: 'colIns', label: 'Cột đầu', act: () => this.freeze(0, 1) },
          { id: 'fz-x', icon: 'unfreeze', label: 'Bỏ cố định', act: () => this.freeze(0, 0) },
        ] },
        { label: 'Hiển thị', items: [
          { id: 'gl', icon: 'grid', label: 'Đường lưới', act: () => { this.grid.gridlines = !this.grid.gridlines; this.grid.invalidate(); this.renderRibbon(); }, on: () => this.grid.gridlines },
        ] },
        { label: 'Thu phóng', items: [
          { id: 'zo', icon: 'zoomOut', label: 'Thu nhỏ', act: () => this.zoom(-1) },
          { id: 'z100', icon: 'table', label: '100%', act: () => this.grid.setZoom(1) },
          { id: 'zi', icon: 'zoomIn', label: 'Phóng to', act: () => this.zoom(1) },
        ] },
        { label: 'Giao diện', items: [
          { id: 't-l', icon: 'sun', label: 'Sáng', act: () => this.setTheme('light'), on: () => this.themePref() === 'light' },
          { id: 't-d', icon: 'moon', label: 'Tối', act: () => this.setTheme('dark'), on: () => this.themePref() === 'dark' },
          { id: 't-s', icon: 'grid', label: 'Theo máy', act: () => this.setTheme('system'), on: () => this.themePref() === 'system' },
        ] },
      ]],
    };
  }
  // ---------- thu gọn ribbon (như Ctrl+F1 của Excel; mặc định thu gọn trên điện thoại nằm ngang) ----------
  private compactQ = matchMedia('(orientation: landscape) and (max-height: 520px)');
  private ribbonPeek = false;
  private get ribbonPinned() {
    const v = lsGet(this.compactQ.matches ? 'xl-ribbon-c' : 'xl-ribbon');
    return v === null ? !this.compactQ.matches : v === '1';
  }
  private set ribbonPinned(v: boolean) { lsSet(this.compactQ.matches ? 'xl-ribbon-c' : 'xl-ribbon', v ? '1' : '0'); }
  private applyRibbon() {
    const shown = this.ribbonPinned || this.ribbonPeek;
    $('#app').classList.toggle('rb-collapsed', !shown);
    $('#app').classList.toggle('rb-peek', !this.ribbonPinned && this.ribbonPeek);
    const t = $('#rb-toggle');
    if (t) { t.innerHTML = ic(this.ribbonPinned ? 'chevUp' : 'chevDown', 16); t.title = this.ribbonPinned ? 'Thu gọn ribbon (Ctrl+F1)' : 'Ghim ribbon (Ctrl+F1)'; t.setAttribute('aria-label', t.title); t.setAttribute('aria-pressed', String(this.ribbonPinned)); }
  }
  toggleRibbon() { this.ribbonPinned = !this.ribbonPinned; this.ribbonPeek = false; this.applyRibbon(); }
  private renderTabs() {
    const nav = $('#tabs');
    nav.innerHTML = '';
    for (const [k, [label]] of Object.entries(this.tabs())) {
      const b = document.createElement('button');
      b.textContent = label;
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(k === this.tab));
      b.onclick = () => {
        if (!this.ribbonPinned) this.ribbonPeek = !(this.ribbonPeek && k === this.tab);
        this.tab = k; this.renderTabs(); this.renderRibbon(); this.applyRibbon();
        if (!this.ed.isEditing() && !this.ribbonPeek) this.grid.focusKeys();
      };
      b.ondblclick = () => this.toggleRibbon();
      nav.appendChild(b);
    }
    // nút thu gọn đặt ngoài tablist (chỉ chứa thẻ) cho đúng chuẩn trợ năng
    if (!document.getElementById('rb-toggle')) {
      const t = document.createElement('button');
      t.id = 'rb-toggle';
      t.className = 'rb-toggle';
      t.onclick = () => this.toggleRibbon();
      nav.after(t);
    }
    this.applyRibbon();
  }
  renderRibbon() {
    const rb = $('#ribbon');
    rb.innerHTML = '';
    for (const grp of this.tabs()[this.tab][1]) {
      const g = document.createElement('div');
      g.className = 'grp';
      const row = document.createElement('div');
      row.className = 'row';
      for (const it of grp.items) {
        const b = document.createElement('button');
        b.className = 'rb';
        b.dataset.id = it.id;
        b.innerHTML = ic(it.icon) + `<span>${it.label}</span>`;
        const tip = it.label + (it.kbd ? ` (${it.kbd})` : '') + (it.later ? ` — có ở ${it.later}` : '');
        b.title = tip;
        b.setAttribute('aria-label', tip);
        if (it.later) { b.classList.add('later'); b.setAttribute('aria-disabled', 'true'); }
        if (it.on) { const on = !!it.on(); b.classList.toggle('on', on); b.setAttribute('aria-pressed', String(on)); }
        b.onmousedown = (e) => e.preventDefault();
        b.onclick = () => {
          if (it.later) { this.toast(`“${it.label}” sẽ có ở ${it.later}.`); return; }
          if (this.ed.isEditing()) this.ed.commit(0, 0);
          it.act();
          if (!MENU_IDS.includes(it.id)) this.grid.focusKeys();
          if (this.ribbonPeek && !MENU_IDS.includes(it.id) && !it.on) { this.ribbonPeek = false; this.applyRibbon(); }
        };
        row.appendChild(b);
      }
      g.appendChild(row);
      const small = document.createElement('small');
      small.textContent = grp.label;
      g.appendChild(small);
      rb.appendChild(g);
    }
  }
  private refreshToggles() {
    const all = Object.values(this.tabs()).flatMap(([, gs]) => gs.flatMap((g) => g.items));
    document.querySelectorAll<HTMLElement>('#ribbon .rb').forEach((b) => {
      const it = all.find((x) => x.id === b.dataset.id);
      if (it?.on) { const on = !!it.on(); b.classList.toggle('on', on); b.setAttribute('aria-pressed', String(on)); }
    });
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
  private colorMenu(k: 'fg' | 'bg') {
    const b = $(`#ribbon [data-id="${k}"]`).getBoundingClientRect();
    const box = document.createElement('div');
    box.className = 'swatches';
    box.innerHTML = COLORS.map((c) => `<button style="--sw:${c}" data-c="${c}" aria-label="${c}"></button>`).join('') + `<button class="none" data-c="">Không màu</button>`;
    box.onclick = (e) => {
      const t = (e.target as HTMLElement).closest('button');
      if (!t) return;
      this.applyStyle({ [k]: t.dataset.c || undefined } as Style);
      closeMenu(); this.grid.focusKeys();
    };
    openPopover(box, b.left, b.bottom + 4);
  }
  private at(id: string) { const r = $(`#ribbon [data-id="${id}"]`)?.getBoundingClientRect(); return r ? { x: r.left, y: r.bottom + 4 } : { x: 80, y: 120 }; }
  private fontMenu() {
    const p = this.at('fn'), cur = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).fn;
    const fonts: [string, string | undefined][] = [['Mặc định (Be Vietnam Pro)', undefined], ['Arial', 'Arial'], ['Times New Roman', 'Times New Roman'], ['Calibri', 'Calibri'], ['Tahoma', 'Tahoma'], ['Verdana', 'Verdana'], ['Roboto', 'Roboto'], ['Courier New', 'Courier New']];
    showMenu(fonts.map(([l, f]) => ({ label: l + (f === cur ? '  ✓' : ''), act: () => this.applyStyle({ fn: f }) })), p.x, p.y);
  }
  private sizeMenu() {
    const p = this.at('fs'), cur = this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).fs ?? 11;
    showMenu([8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 36].map((n) => ({ label: String(n) + (n === cur ? '  ✓' : ''), act: () => { this.applyStyle({ fs: n === 11 ? undefined : n }); this.fitRows(); } })), p.x, p.y);
  }
  private borderMenu() {
    const p = this.at('border');
    const b = (kind: BorderKind, style: BorderStyle = 'thin') => () => this.safe(() => this.wb.setBorder(this.grid.sheet, this.grid.range, kind, style));
    showMenu([
      { label: 'Tất cả viền', icon: 'borders', act: b('all') },
      { label: 'Viền ngoài', icon: 'borderOuter', act: b('outer') },
      { label: 'Viền ngoài đậm', icon: 'borderOuter', act: b('outer', 'medium') },
      { label: 'Viền trong', icon: 'borders', act: b('inner') },
      'sep',
      { label: 'Viền trên', icon: 'borderTop', act: b('top') },
      { label: 'Viền dưới', icon: 'borderBottom', act: b('bottom') },
      { label: 'Viền dưới kép', icon: 'borderBottom', act: b('bottom', 'double') },
      { label: 'Viền trái', icon: 'borderLeft', act: b('left') },
      { label: 'Viền phải', icon: 'borderRight', act: b('right') },
      'sep',
      { label: 'Viền nét đứt', icon: 'borders', act: b('all', 'dashed') },
      { label: 'Không viền', icon: 'x', act: b('none') },
    ], p.x, p.y);
  }
  private vAlignMenu() {
    const p = this.at('va');
    showMenu([['Trên', 't'], ['Giữa', 'm'], ['Dưới', undefined]].map(([l, v]) => ({ label: l as string, act: () => this.applyStyle({ va: v as Style['va'] }) })), p.x, p.y);
  }
  private toggleWrap() {
    const on = !this.wb.style(this.grid.sheet, this.grid.sel.ar, this.grid.sel.ac).wr;
    this.applyStyle({ wr: on });
    if (on) this.fitRows();
  }
  /** Tăng chiều cao hàng cho vừa chữ (ô xuống dòng hoặc chữ to) */
  fitRows() {
    const g = this.grid.range, s = this.grid.sheet, u = this.wb.used(s);
    const r2 = Math.min(g.r2, u.rows - 1, g.r1 + 2000);
    for (let r = g.r1; r <= r2; r++) {
      if (this.wb.isRowHidden(s, r)) continue;
      const need = this.grid.neededHeight(r);
      if (need > this.wb.meta(s).rowH.get(r)) this.wb.setRowHeight(s, r, need);
    }
  }
  private mergeMenu() {
    const p = this.at('merge'), s = this.grid.sheet, g = this.grid.range;
    const doMerge = (center: boolean) => () => this.safe(() => {
      const lost = this.wb.merge(s, g, center);
      this.grid.selectRange(this.wb.expandToMerges(s, g));
      if (lost) this.toast(`Đã gộp ô. Chỉ giữ giá trị ô trên cùng bên trái (bỏ ${lost} ô khác).`, 'ok', { label: 'Hoàn tác', act: () => this.undo() });
    });
    showMenu([
      { label: 'Gộp và căn giữa', icon: 'merge', act: doMerge(true) },
      { label: 'Gộp ô', icon: 'merge', act: doMerge(false) },
      { label: 'Gộp từng hàng', icon: 'merge', act: () => this.safe(() => { for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) this.wb.merge(s, { ...g, r1: r, r2: r }, false); }) },
      { label: 'Bỏ gộp ô', icon: 'x', act: () => { this.wb.unmerge(s, g); this.grid.invalidate(); } },
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
  saveXlsx(asNew: boolean): Promise<boolean> {
    if (this.saving) return this.saving;
    this.saving = this.doSaveXlsx(asNew).finally(() => { this.saving = null; });
    return this.saving;
  }
  private async doSaveXlsx(asNew: boolean): Promise<boolean> {
    try {
      if (this.ed.isEditing()) this.ed.commit(0, 0);
      const ver = this.wb.ver;
      const { writeXlsx } = await import('../core/xlsx');
      const bytes = await writeXlsx(this.wb);
      const name = (this.fileName && /\.xlsx$/i.test(this.fileName) ? this.fileName : this.wb.name + '.xlsx');
      const res = await saveBytes(bytes, name, 'xlsx', asNew || !/\.xlsx$/i.test(this.fileName) ? undefined : this.fileHandle);
      if (!res) return false;
      if (res.handle) { this.fileHandle = res.handle; this.fileName = res.handle.name ?? name; }
      // sửa trong lúc đang ghi thì tệp chưa có các sửa đó → vẫn coi là chưa lưu
      if (res.how === 'file') this.fileDirty = this.wb.ver !== ver;
      await this.saveNow(false);
      this.toast(res.how === 'file' ? `Đã lưu ${this.fileName || name}.` : `Đã tải xuống ${name}.`, 'ok');
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
  private csvMenu() {
    const p = this.at('csv'), s = this.grid.sheet;
    const go = (d: string) => async () => {
      const res = await saveBytes(this.wb.exportCSV(s, d), `${this.wb.name} - ${this.wb.sheetName(s)}.csv`, 'csv');
      if (res) this.toast('Đã xuất CSV sheet “' + this.wb.sheetName(s) + '”.', 'ok');
    };
    showMenu([
      { label: 'CSV dấu chấm phẩy ( ; ) – Excel tiếng Việt', icon: 'table', act: go(';') },
      { label: 'CSV dấu phẩy ( , ) – quốc tế', icon: 'table', act: go(',') },
    ], p.x, p.y);
  }

  private fmtMenu() {
    const b = $('#ribbon [data-id="fmt"]').getBoundingClientRect();
    showMenu(NUM_FMTS.map(([l, f, ex]) => ({ label: l, hint: ex, act: () => this.applyStyle({ fmt: f }) })), b.left, b.bottom + 4);
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
  autoSum() {
    const { ar, ac } = this.grid.sel;
    const f = this.wb.autoSum(this.grid.sheet, ar, ac);
    if (f === '=SUM()') this.ed.begin('=SUM(', 'enter');
    else this.syncSelection();
  }
  fillDir(dir: 'down' | 'right') {
    const g = this.grid.range;
    if (dir === 'down' && g.r2 > g.r1) this.safe(() => this.wb.fill(this.grid.sheet, { ...g, r2: g.r1 }, g));
    if (dir === 'right' && g.c2 > g.c1) this.safe(() => this.wb.fill(this.grid.sheet, { ...g, c2: g.c1 }, g));
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

  // ---------- sheet ----------
  addSheet() { this.safe(() => { const id = this.wb.addSheet(); this.grid.setSheet(id); this.renderSheets(); }); }
  renderSheets() {
    const box = $('#sheets');
    box.innerHTML = '';
    for (const id of this.wb.sheetIds()) {
      const b = document.createElement('button');
      b.className = 'st' + (id === this.grid.sheet ? ' on' : '');
      b.textContent = this.wb.sheetName(id);
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(id === this.grid.sheet));
      b.onmousedown = (e) => { if (this.ed.isEditing()) e.preventDefault(); };
      b.onclick = () => {
        // đang gõ công thức ở chế độ trỏ: chuyển sheet để chọn ô tham chiếu, chưa nhập
        if (this.ed.isEditing() && this.ed.pointMode()) { if (id !== this.grid.sheet) { this.grid.setSheet(id); this.renderSheets(); } this.ed.refocus(); return; }
        if (this.ed.isEditing()) this.ed.commit(0, 0);
        if (id !== this.grid.sheet) { this.grid.setSheet(id); this.renderSheets(); }
        this.grid.focusKeys();
      };
      b.ondblclick = () => this.renameSheet(id, b);
      b.oncontextmenu = (e) => { e.preventDefault(); this.sheetMenu(id, b, e.clientX, e.clientY); };
      let lp = 0;
      b.onpointerdown = (e) => { if (e.pointerType === 'touch') lp = window.setTimeout(() => this.sheetMenu(id, b, e.clientX, e.clientY), 550); };
      b.onpointerup = b.onpointercancel = () => clearTimeout(lp);
      box.appendChild(b);
    }
    if (!document.getElementById('sheet-add')) {
      const add = document.createElement('button');
      add.id = 'sheet-add';
      add.className = 'add';
      add.innerHTML = ic('plus', 16);
      add.title = add.ariaLabel = 'Thêm sheet';
      add.onclick = () => this.addSheet();
      box.after(add);
    }
  }
  private sheetMenu(id: number, el: HTMLElement, x: number, y: number) {
    void el;
    showMenu([
      { label: 'Đổi tên', icon: 'file', act: () => this.renameSheet(id, document.querySelector<HTMLElement>(`#sheets .st:nth-child(${this.wb.sheetIds().indexOf(id) + 1})`)!) },
      { label: 'Thêm sheet', icon: 'plus', act: () => this.addSheet() },
      { label: 'Xóa sheet', icon: 'trash', danger: true, act: () => this.removeSheet(id) },
    ], x, y);
  }
  private removeSheet(id: number) {
    const n = this.wb.sheetName(id);
    this.safe(() => {
      this.wb.removeSheet(id);
      if (id === this.grid.sheet) this.grid.setSheet(this.wb.sheetIds()[0]);
      this.renderSheets();
      this.toast(`Đã xóa sheet “${n}”.`, 'ok', { label: 'Hoàn tác', act: () => this.undo() });
    });
  }
  private renameSheet(id: number, el: HTMLElement) {
    const inp = document.createElement('input');
    inp.className = 'st-edit';
    inp.value = this.wb.sheetName(id);
    inp.setAttribute('aria-label', 'Tên sheet');
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

  // ---------- thanh dưới, ô tên ----------
  private bindFoot() {
    $('#zoom-out').onclick = () => this.zoom(-1);
    $('#zoom-in').onclick = () => this.zoom(1);
    $('#zoom-val').onclick = () => this.grid.setZoom(1);
  }
  private bindNameBox() {
    const nb = $<HTMLInputElement>('#namebox');
    nb.onfocus = () => nb.select();
    nb.onkeydown = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); this.syncSelection(); this.grid.focusKeys(); }
      if (e.key !== 'Enter') return;
      e.preventDefault();
      const g = parseRange(nb.value.replace(/\s/g, ''));
      if (!g) { this.toast('Địa chỉ không hợp lệ. Ví dụ: B5 hoặc A1:C10', 'err'); return; }
      this.grid.selectRange(g); this.grid.focusKeys();
    };
  }
  syncSelection() {
    const g = this.grid.range, s = this.grid.sheet, { ar, ac } = this.grid.sel;
    const nb = $<HTMLInputElement>('#namebox');
    if (document.activeElement !== nb) {
      const fullC = g.r1 === 0 && g.r2 === MAX_ROWS - 1, fullR = g.c1 === 0 && g.c2 === MAX_COLS - 1;
      nb.value = fullC && fullR ? 'Tất cả' : fullC ? rangeToA1(g).replace(/\d+/g, '') : fullR ? `${g.r1 + 1}:${g.r2 + 1}` : g.r1 === g.r2 && g.c1 === g.c2 ? toA1(ar, ac) : rangeToA1(g);
    }
    if (!this.ed.isEditing()) $<HTMLTextAreaElement>('#fx-input').value = this.wb.rawText(s, ar, ac);
    this.refreshToggles();
    if (!this.bookChanging) this.pivot?.follow(s, ar, ac);
    this.updateStatus();
    this.announce();
    $('#zoom-val').textContent = Math.round(this.grid.zoom * 100) + '%';
    $('#btn-undo').toggleAttribute('disabled', !this.wb.canUndo());
    $('#btn-redo').toggleAttribute('disabled', !this.wb.canRedo());
  }
  private updateStatus() {
    const g = this.grid.range, s = this.grid.sheet, u = this.wb.used(s);
    const r2 = Math.min(g.r2, u.rows - 1), c2 = Math.min(g.c2, u.cols - 1);
    const agg = $('#agg');
    if ((g.r1 === g.r2 && g.c1 === g.c2) || r2 < g.r1 || c2 < g.c1 || (r2 - g.r1 + 1) * (c2 - g.c1 + 1) > 1_000_000) { agg.innerHTML = ''; return; }
    let sum = 0, nums = 0, cnt = 0;
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) {
      const v = this.wb.value(s, r, c);
      if (v === null || v === '') continue;
      cnt++;
      if (typeof v === 'number') { sum += v; nums++; }
    }
    const f = (n: number) => n.toLocaleString('vi-VN', { maximumFractionDigits: 10 });
    agg.innerHTML = cnt < 2 ? '' : nums ? `Trung bình: <b>${f(sum / nums)}</b> <span>Đếm: <b>${cnt}</b></span> Tổng: <b>${f(sum)}</b>` : `Đếm: <b>${cnt}</b>`;
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
  insertChart(type?: 'line' | 'pie') {
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
  printDlg() { printDialog(this.wb, this.grid.sheet, this.grid.range, (m, k) => this.toast(m, k)); }

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
    $('#save-state').textContent = 'Chưa lưu';
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
      $('#save-state').textContent = `Đã lưu lúc ${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}`;
      if (announce) this.toast('Đã lưu trên máy này.', 'ok');
    } catch (e) {
      $('#save-state').textContent = 'Không lưu được';
      this.toast('Không lưu được: ' + (e as Error).message, 'err');
    }
  }
  private async newBook() {
    await this.saveNow(false);
    const id = 'b' + Date.now();
    lsSet('xl-last', id);
    location.hash = '';
    this.reload();
  }
  private async openMenu() {
    let books: Awaited<ReturnType<typeof listBooks>> = [];
    try { books = await listBooks(); } catch { /* bỏ qua */ }
    const b = $('#ribbon [data-id="open"]').getBoundingClientRect();
    if (!books.length) { this.toast('Chưa có sổ tính nào được lưu.'); return; }
    showMenu(books.map((bk) => ({
      label: bk.name + (bk.id === this.bookId ? ' (đang mở)' : ''),
      hint: new Date(bk.updated).toLocaleString('vi-VN', { dateStyle: 'short', timeStyle: 'short' }),
      icon: 'file',
      act: async () => {
        if (bk.id === this.bookId) return;
        await this.saveNow(false);
        lsSet('xl-last', bk.id);
        this.reload();
      },
      side: bk.id === this.bookId ? undefined : { icon: 'trash', label: 'Xóa', act: async () => { await deleteBook(bk.id); this.toast(`Đã xóa “${bk.name}”.`); } },
    })), b.left, b.bottom + 4);
  }

  // ---------- theme ----------
  themePref() { return (lsGet('xl-theme') as 'light' | 'dark' | 'system') || 'system'; }
  private bindTheme() {
    this.applyTheme();
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => this.applyTheme());
    $('#btn-theme').onclick = () => {
      const dark = document.documentElement.dataset.theme ? document.documentElement.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
      this.setTheme(dark ? 'light' : 'dark');
    };
  }
  setTheme(t: 'light' | 'dark' | 'system') { lsSet('xl-theme', t); this.applyTheme(); this.renderRibbon(); }
  private applyTheme() {
    const t = this.themePref();
    if (t === 'system') delete document.documentElement.dataset.theme; else document.documentElement.dataset.theme = t;
    const dark = t === 'dark' || (t === 'system' && matchMedia('(prefers-color-scheme: dark)').matches);
    document.querySelector('meta[name="theme-color"]')?.setAttribute('content', dark ? '#12382a' : '#0f5e3d');
    requestAnimationFrame(() => this.grid.readColors());
  }

  // ---------- cài đặt PWA ----------
  private bindInstall() {
    window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); this.installEvt = e; $('#btn-install').hidden = false; });
    $('#btn-install').onclick = () => this.install();
    window.addEventListener('appinstalled', () => { $('#btn-install').hidden = true; this.toast('Đã cài ExcelLike.', 'ok'); });
  }
  private async install() {
    if (this.installEvt) { this.installEvt.prompt(); await this.installEvt.userChoice; this.installEvt = null; $('#btn-install').hidden = true; return; }
    const ios = /iphone|ipad|ipod/i.test(navigator.userAgent);
    if (matchMedia('(display-mode: standalone)').matches) this.toast('ExcelLike đang chạy như ứng dụng đã cài.');
    else if (ios) this.toast('Trên iPhone/iPad: bấm nút Chia sẻ rồi chọn “Thêm vào MH chính”.');
    else this.toast('Mở bằng Chrome hoặc Edge rồi bấm biểu tượng cài đặt trên thanh địa chỉ.');
  }

  // ---------- menu ô & hàm ----------
  private cellMenu(x: number, y: number) {
    showMenu([
      { label: 'Cắt', icon: 'cut', hint: 'Ctrl+X', act: () => this.copySel(true) },
      { label: 'Sao chép', icon: 'copy', hint: 'Ctrl+C', act: () => this.copySel(false) },
      { label: 'Dán', icon: 'paste', hint: 'Ctrl+V', act: () => this.pasteBtn() },
      'sep',
      { label: 'Chèn hàng phía trên', icon: 'rowIns', act: () => this.insRows() },
      { label: 'Chèn cột bên trái', icon: 'colIns', act: () => this.insCols() },
      { label: 'Xóa hàng', icon: 'rowDel', act: () => this.delRows() },
      { label: 'Xóa cột', icon: 'colDel', act: () => this.delCols() },
      'sep',
      { label: 'Xóa nội dung', icon: 'eraser', hint: 'Delete', act: () => this.clearSel() },
      { label: 'Định dạng số…', icon: 'table', act: () => this.fmtMenuAt(x, y) },
    ], x, y);
  }
  private fmtMenuAt(x: number, y: number) { showMenu(NUM_FMTS.map(([l, f, ex]) => ({ label: l, hint: ex, act: () => this.applyStyle({ fmt: f }) })), x, y); }
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
    s.font = `400 ${(st.fs ? (st.fs * 13) / 11 : 13) * g.zoom}px ${st.fn ? `"${st.fn}", ` : ''}'Be Vietnam Pro', system-ui, sans-serif`;
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
    if (k === 'Enter') { e.preventDefault(); this.commit(e.shiftKey ? -1 : 1, 0); return; }
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
      case 'Enter': done(); g.moveActive(e.shiftKey ? -1 : 1, 0); return;
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

// ======================= menu / popover =======================
type MenuItem = { label: string; act: () => void; icon?: string; hint?: string; danger?: boolean; side?: { icon: string; label: string; act: () => void } } | 'sep';
let openEl: HTMLElement | null = null;
let restoreFocus: HTMLElement | null = null;
export function closeMenu() {
  openEl?.remove();
  openEl = null;
  document.removeEventListener('pointerdown', outside, true);
  document.removeEventListener('keydown', escClose, true);
  restoreFocus?.focus({ preventScroll: true });
  restoreFocus = null;
}
function outside(e: Event) { if (openEl && !openEl.contains(e.target as Node)) closeMenu(); }
function escClose(e: KeyboardEvent) {
  if (!openEl) return;
  if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeMenu(); return; }
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    const bs = [...openEl.querySelectorAll<HTMLButtonElement>('button.mi')];
    if (!bs.length) return;
    e.preventDefault();
    const i = bs.indexOf(document.activeElement as HTMLButtonElement);
    bs[(i + (e.key === 'ArrowDown' ? 1 : -1) + bs.length) % bs.length].focus();
  }
}
export function openPopover(el: HTMLElement, x: number, y: number) {
  const keep = document.activeElement as HTMLElement | null;
  closeMenu();
  restoreFocus = keep;
  el.classList.add('pop');
  document.body.appendChild(el);
  const r = el.getBoundingClientRect();
  const vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
  el.style.left = Math.max(8, Math.min(x, vw - r.width - 8)) + 'px';
  el.style.top = (y + r.height > vh - 8 ? Math.max(8, y - r.height - 8) : y) + 'px';
  openEl = el;
  setTimeout(() => { document.addEventListener('pointerdown', outside, true); document.addEventListener('keydown', escClose, true); });
}
export function showMenu(items: MenuItem[], x: number, y: number) {
  const m = document.createElement('div');
  m.className = 'menu';
  m.setAttribute('role', 'menu');
  for (const it of items) {
    if (it === 'sep') { m.appendChild(Object.assign(document.createElement('hr'), {})); continue; }
    const row = document.createElement('div');
    row.className = 'mrow';
    const b = document.createElement('button');
    b.className = 'mi' + (it.danger ? ' danger' : '');
    b.setAttribute('role', 'menuitem');
    b.innerHTML = (it.icon ? ic(it.icon, 16) : '<i></i>') + `<span></span>` + (it.hint ? `<kbd></kbd>` : '');
    b.querySelector('span')!.textContent = it.label;
    if (it.hint) b.querySelector('kbd')!.textContent = it.hint;
    b.onclick = () => { closeMenu(); it.act(); };
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
  }
  openPopover(m, x, y);
  (m.querySelector('button.mi') as HTMLButtonElement | null)?.focus();
}

function pad(n: number) { return String(n).padStart(2, '0'); }
function colIdx(s: string) { let c = 0; for (const ch of s) c = c * 26 + ch.charCodeAt(0) - 64; return c - 1; }
function normalizeNL(s: string) { return s.replace(/\r\n/g, '\n').replace(/\n$/, ''); }
