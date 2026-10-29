// Backstage (Tệp) toàn màn hình như Excel 365: Trang đầu, Mới, Mở, Thông tin, Lưu, Lưu thành, In, Xuất, Đóng, Tài khoản, Tùy chọn
import { listBooks, deleteBook, lsGet, lsSet } from './store';
import { desk, saveBytes } from './files';
import { ic } from './icons';
import { PrintPage } from './printui';
import { openDialog, h } from './dialogs';
import { buildJob } from '../core/print';
import { chartImgs, printViaFrame } from './printui';
import type { App } from './app';
import { VERSION } from './version';

type Page = 'home' | 'new' | 'open' | 'info' | 'save' | 'saveas' | 'print' | 'export' | 'close' | 'account' | 'options';
const NAV: [Page, string, string][] = [['home', 'Trang đầu', 'home'], ['new', 'Mới', 'newFile'], ['open', 'Mở', 'folder']];
const NAV2: [Page, string, string][] = [['info', 'Thông tin', 'info2'], ['save', 'Lưu', 'save'], ['saveas', 'Lưu thành', 'download'], ['print', 'In', 'printer2'], ['export', 'Xuất', 'export'], ['close', 'Đóng', 'close']];
const NAV3: [Page, string, string][] = [['account', 'Tài khoản', 'account'], ['options', 'Tùy chọn', 'gear']];
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!));

export class Backstage {
  private el: HTMLElement;
  private page: Page = 'home';
  private printPage: PrintPage | null = null;
  constructor(private app: App) {
    this.el = document.getElementById('backstage')!;
  }
  isOpen() { return !this.el.hidden; }
  open(page: Page = 'home') {
    if (this.app.ed.isEditing()) this.app.ed.commit(0, 0);
    this.app.keytips?.exit();
    this.el.hidden = false;
    if (!this.el.firstChild) this.build();
    this.show(page);
  }
  close() {
    this.el.hidden = true;
    this.app.grid.layout();
    this.app.grid.focusKeys();
  }
  refreshPrint() { if (this.isOpen() && this.page === 'print') this.printPage?.refresh(); }
  quickPrint() {
    const s = this.app.grid.sheet;
    const j = buildJob(this.app.wb, [{ sheet: s, charts: chartImgs(this.app.wb, s) }], { file: this.app.fileName || this.app.wb.name });
    if (!j.total) { this.app.toast('Không có gì để in.', 'err'); return; }
    const d = desk();
    const st = j.layouts[0]?.setup;
    if (d?.print) d.print({ html: j.html, silent: true, paper: st?.paper, landscape: st?.orientation === 'landscape' }).then((r: any) => this.app.toast(r?.ok ? 'Đã gửi tới máy in mặc định.' : 'Không in được: ' + (r?.error ?? ''), r?.ok ? 'ok' : 'err'));
    else printViaFrame(j.html);
  }
  private build() {
    const nav = document.createElement('nav');
    nav.className = 'bs-nav';
    nav.setAttribute('aria-label', 'Tệp');
    const back = document.createElement('button');
    back.id = 'bs-back';
    back.innerHTML = ic('arrowL', 20);
    back.title = 'Quay lại (Esc)';
    back.setAttribute('aria-label', 'Quay lại');
    back.onclick = () => this.close();
    nav.appendChild(back);
    const add = (list: [Page, string, string][]) => list.forEach(([k, l, icon]) => {
      const b = document.createElement('button');
      b.dataset.bs = k;
      b.innerHTML = ic(icon, 18) + `<span>${l}</span>`;
      b.onclick = () => this.show(k);
      nav.appendChild(b);
    });
    add(NAV);
    nav.appendChild(document.createElement('hr'));
    add(NAV2);
    const sp = document.createElement('div'); sp.className = 'sp'; nav.appendChild(sp);
    add(NAV3);
    const main = document.createElement('div');
    main.className = 'bs-page';
    this.el.append(nav, main);
  }
  private main() { return this.el.querySelector('.bs-page') as HTMLElement; }
  show(page: Page) {
    const app = this.app;
    // Lưu / Đóng là lệnh chạy ngay như Excel
    if (page === 'save') { this.close(); app.save(); return; }
    if (page === 'close') { this.close(); if (desk()) window.close(); else app.newBook(); return; }
    this.page = page;
    this.el.querySelectorAll<HTMLElement>('.bs-nav [data-bs]').forEach((b) => b.setAttribute('aria-current', b.dataset.bs === page ? 'page' : 'false'));
    const m = this.main();
    m.className = 'bs-page' + (page === 'print' ? ' print' : '');
    m.innerHTML = '';
    const h1 = (t: string) => { const e = document.createElement('h1'); e.textContent = t; m.appendChild(e); };
    const cmd = (icon: string, title: string, sub: string, act: () => void, a?: string) => {
      const b = document.createElement('button');
      b.className = 'bs-cmd';
      if (a) b.dataset.a = a;
      b.innerHTML = `<span class="ico">${ic(icon, 26)}</span><span><b>${esc(title)}</b><small>${esc(sub)}</small></span>`;
      b.onclick = act;
      return b;
    };
    const blank = () => {
      const t = document.createElement('div');
      t.className = 'bs-tiles';
      const b = document.createElement('button');
      b.className = 'bs-tile';
      b.dataset.a = 'blank';
      b.innerHTML = `<span class="thumb">${ic('table', 40)}</span><span>Sổ làm việc trống</span>`;
      b.onclick = () => { this.close(); app.newBook(); };
      t.appendChild(b);
      return t;
    };
    if (page === 'home') {
      const hr = new Date().getHours();
      h1(hr < 11 ? 'Chào buổi sáng' : hr < 18 ? 'Chào buổi chiều' : 'Chào buổi tối');
      m.appendChild(blank());
      const h2 = document.createElement('h2'); h2.textContent = 'Gần đây'; m.appendChild(h2);
      this.recent(m, 8);
    } else if (page === 'new') {
      h1('Mới');
      m.appendChild(blank());
    } else if (page === 'open') {
      h1('Mở');
      const c = document.createElement('div'); c.className = 'bs-cmds';
      c.appendChild(cmd('folder', 'Duyệt', 'Mở tệp .xlsx, .xlsm hoặc .csv trên máy (Ctrl+O)', () => { this.close(); app.openFile(); }, 'browse'));
      m.appendChild(c);
      const h2 = document.createElement('h2'); h2.textContent = 'Sổ làm việc gần đây'; m.appendChild(h2);
      this.recent(m, 50);
    } else if (page === 'info') {
      h1('Thông tin');
      const wb = app.wb;
      let cells = 0;
      for (const id of wb.sheetIds()) { const u = wb.used(id); cells += u.rows * u.cols; }
      const dl = document.createElement('dl');
      dl.className = 'bs-props';
      const rows: [string, string][] = [
        ['Tên', wb.name],
        ['Tệp', app.fileName || 'Chưa lưu ra tệp (chỉ lưu trong ứng dụng)'],
        ['Trang tính', String(wb.sheetIds().length)],
        ['Trang tính ẩn', String(wb.sheetIds().filter((i) => wb.isSheetHidden(i)).length)],
        ['Vùng dữ liệu', cells.toLocaleString('vi-VN') + ' ô'],
        ['Liên kết ngoài', String(wb.extLinks.length)],
        ['Tính lặp', wb.iterate.on ? `Bật (${wb.iterate.count} lần, sai số ${String(wb.iterate.delta).replace('.', ',')})` : 'Tắt'],
        ['Trạng thái', document.getElementById('save-state')?.textContent ?? ''],
      ];
      for (const [k, v] of rows) { const dt = document.createElement('dt'); dt.textContent = k; const dd = document.createElement('dd'); dd.textContent = v; dl.append(dt, dd); }
      m.appendChild(dl);
      const c = document.createElement('div'); c.className = 'bs-cmds'; c.style.marginTop = '20px';
      c.append(
        cmd('protectBook', 'Bảo vệ Sổ làm việc', 'Có ở bản 0.9', () => app.later('Bảo vệ Sổ làm việc', '0.9')),
        cmd('errorCheck', 'Kiểm tra Sổ làm việc', 'Thống kê nội dung sổ làm việc', () => app.bookStats()),
      );
      m.appendChild(c);
    } else if (page === 'saveas' || page === 'export') {
      h1(page === 'saveas' ? 'Lưu thành' : 'Xuất');
      const c = document.createElement('div'); c.className = 'bs-cmds';
      const csv = (d: string) => async () => {
        const s = app.grid.sheet;
        const r = await saveBytes(app.wb.exportCSV(s, d), `${app.wb.name} - ${app.wb.sheetName(s)}.csv`, 'csv');
        if (r) { this.close(); app.toast(`Đã xuất CSV trang tính “${app.wb.sheetName(s)}”.`, 'ok'); }
      };
      if (page === 'saveas') c.append(cmd('download', 'Sổ làm việc Excel (*.xlsx)', 'Chọn nơi lưu và tên tệp', () => { this.close(); app.saveXlsx(true); }, 'xlsx'));
      c.append(
        cmd('pdf', 'Tạo Tài liệu PDF', 'Giữ nguyên bố cục trang in (khổ giấy, lề, đầu/chân trang)', () => this.show('print'), 'pdf'),
        cmd('csvFile', 'CSV UTF-8 (phân tách bằng dấu chấm phẩy)', 'Trang tính hiện hoạt — Excel tiếng Việt', csv(';'), 'csv1'),
        cmd('csvFile', 'CSV UTF-8 (phân tách bằng dấu phẩy)', 'Trang tính hiện hoạt — chuẩn quốc tế', csv(','), 'csv2'),
      );
      if (page === 'export') c.append(cmd('download', 'Thay đổi Loại Tệp: Sổ làm việc Excel (*.xlsx)', 'Lưu một bản .xlsx mới', () => { this.close(); app.saveXlsx(true); }, 'xlsx'));
      m.appendChild(c);
    } else if (page === 'print') {
      h1('In');
      if (!this.printPage) this.printPage = new PrintPage({
        wb: app.wb, sheet: () => app.grid.sheet, selection: () => app.grid.range,
        fileName: () => app.fileName || app.wb.name + '.xlsx',
        toast: (msg, k) => app.toast(msg, k ?? ''), pageSetup: (t) => app.pageSetup(t),
      });
      m.appendChild(this.printPage.el);
      this.printPage.refresh();
      setTimeout(() => (document.getElementById('bs-print') as HTMLElement | null)?.focus(), 30);
    } else if (page === 'account') {
      h1('Tài khoản');
      const f = document.createElement('div'); f.className = 'bs-form';
      f.innerHTML = `<label class="fld"><span>Chủ đề Office</span><select id="bs-theme"><option value="colorful">Nhiều màu</option><option value="darkgray">Xám Đậm</option><option value="black">Đen</option><option value="white">Trắng</option><option value="system">Dùng thiết đặt hệ thống</option></select></label>
<h2>Thông tin Sản phẩm</h2><p><b>ExcelLike</b> phiên bản ${VERSION}<br><span class="muted">Bảng tính kiểu Excel 365 — bản Windows và PWA cùng một mã nguồn. Công thức tiếng Việt, .xlsx, biểu đồ, PivotTable, in ấn có xem trước.</span></p>`;
      m.appendChild(f);
      const sel = f.querySelector('#bs-theme') as HTMLSelectElement;
      sel.value = app.themePref();
      sel.onchange = () => app.setTheme(sel.value);
      if (!desk()) f.appendChild(cmd('install', 'Cài đặt ứng dụng', 'Chạy như ứng dụng riêng, dùng được khi không có mạng', () => app.install(), 'install'));
    } else if (page === 'options') {
      h1('Tùy chọn');
      const f = document.createElement('div'); f.className = 'bs-form';
      const it = app.wb.iterate;
      const dir = lsGet('xl-enter') ?? 'down';
      f.innerHTML = `<h2>Công thức</h2>
<label class="chk"><input type="checkbox" id="opt-iter"${it.on ? ' checked' : ''}> Bật tính toán lặp</label>
<div class="row2"><label class="fld"><span>Số lần lặp Tối đa</span><input id="opt-iter-n" type="text" inputmode="numeric" value="${it.count}"></label><label class="fld"><span>Thay đổi Tối đa</span><input id="opt-iter-d" type="text" inputmode="decimal" value="${String(it.delta).replace('.', ',')}"></label></div>
<h2>Nâng cao</h2>
<label class="chk"><input type="checkbox" id="opt-enter-on"${dir !== 'none' ? ' checked' : ''}> Sau khi nhấn Enter, di chuyển vùng chọn</label>
<label class="fld"><span>Hướng</span><select id="opt-enter"><option value="down">Xuống</option><option value="right">Phải</option><option value="up">Lên</option><option value="left">Trái</option></select></label>
<label class="chk"><input type="checkbox" id="opt-status"${lsGet('xl-status') !== null ? '' : ' checked'}> Thanh trạng thái mặc định (Trung bình, Số đếm, Tổng)</label>`;
      m.appendChild(f);
      const q = <T extends HTMLElement>(id: string) => f.querySelector('#' + id) as T;
      q<HTMLSelectElement>('opt-enter').value = dir === 'none' ? 'down' : dir;
      const iter = () => {
        const n = parseInt(q<HTMLInputElement>('opt-iter-n').value), d = parseFloat(q<HTMLInputElement>('opt-iter-d').value.replace(',', '.'));
        app.wb.setIterate({ on: q<HTMLInputElement>('opt-iter').checked, count: n >= 1 && n <= 32767 ? n : it.count, delta: d > 0 ? d : it.delta });
        app.grid.invalidate(); app.ribbon.refresh();
      };
      q('opt-iter').onchange = iter; q('opt-iter-n').onchange = iter; q('opt-iter-d').onchange = iter;
      const enter = () => lsSet('xl-enter', q<HTMLInputElement>('opt-enter-on').checked ? q<HTMLSelectElement>('opt-enter').value : 'none');
      q('opt-enter').onchange = enter; q('opt-enter-on').onchange = enter;
      q('opt-status').onchange = () => { if (q<HTMLInputElement>('opt-status').checked) localStorage.removeItem('xl-status'); };
    }
  }
  private recent(m: HTMLElement, n: number) {
    const box = document.createElement('div');
    box.className = 'bs-recent';
    m.appendChild(box);
    listBooks().then((books0) => {
      // sổ đang mở luôn có trong danh sách (kể cả khi chưa kịp lưu lần đầu)
      const books = books0.some((b) => b.id === this.app.bookId) ? books0 : [{ id: this.app.bookId, name: this.app.wb.name, updated: Date.now(), fileName: this.app.fileName } as (typeof books0)[number], ...books0];
      if (!books.length) { box.innerHTML = '<p class="muted" style="padding:12px 4px">Chưa có sổ làm việc nào.</p>'; return; }
      for (const bk of books.slice(0, n)) {
        const row = document.createElement('div');
        row.className = 'rr';
        row.dataset.book = bk.id;
        row.setAttribute('role', 'button');
        row.tabIndex = 0;
        row.innerHTML = `${ic('table', 20)}<span><b></b><br><small></small></span><small>${new Date(bk.updated).toLocaleString('vi-VN', { dateStyle: 'short', timeStyle: 'short' })}</small>`;
        row.querySelector('b')!.textContent = bk.name + (bk.id === this.app.bookId ? ' (đang mở)' : '');
        row.querySelector('span small')!.textContent = bk.fileName || 'Trong ứng dụng';
        const open = () => { if (bk.id !== this.app.bookId) this.app.switchBook(bk.id); else this.close(); };
        row.onclick = open;
        row.onkeydown = (e) => { if (e.key === 'Enter') open(); };
        if (bk.id !== this.app.bookId) {
          const del = document.createElement('button');
          del.className = 'del';
          del.innerHTML = ic('trash', 16);
          del.title = 'Xóa sổ này khỏi ứng dụng'; del.setAttribute('aria-label', del.title);
          del.onclick = (e) => {
            e.stopPropagation();
            const msg = bk.fileName ? `Xóa “${bk.name}” khỏi danh sách? Bản nháp trong ứng dụng sẽ bị xóa; tệp “${bk.fileName}” trên máy vẫn còn.` : `Xóa vĩnh viễn sổ “${bk.name}”? Sổ này chưa được lưu ra tệp nên sẽ không khôi phục được.`;
            openDialog('Xóa sổ làm việc', h('p', {}, msg), [
              { label: 'Hủy', act: () => {} },
              { label: 'Xóa', kind: 'danger', act: () => { deleteBook(bk.id).then(() => { row.remove(); this.app.toast(`Đã xóa “${bk.name}”.`); }); } },
            ]);
          };
          row.appendChild(del);
        } else row.appendChild(document.createElement('span'));
        box.appendChild(row);
      }
    }).catch(() => { box.innerHTML = '<p class="muted">Không đọc được danh sách.</p>'; });
  }
}
