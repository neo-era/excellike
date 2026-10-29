import { KEYMAP } from './ui/keymap';
import { startI18n, setLang } from './ui/i18n';
// phông Be Vietnam Pro đóng gói kèm app (chạy offline, không gọi Google Fonts)
import '@fontsource/be-vietnam-pro/400.css';
import '@fontsource/be-vietnam-pro/500.css';
import '@fontsource/be-vietnam-pro/600.css';
import '@fontsource/be-vietnam-pro/700.css';
import './styles.css';
import { Workbook } from './core/workbook';
import { App } from './ui/app';
import { loadBook, lsGet, lsSet } from './ui/store';
import { ensureLicense } from './ui/license';

/** Không khởi động được: hiện thông báo rõ ràng thay vì màn hình trống, kèm cách mở chế độ an toàn */
function bootError(e: unknown) {
  console.error('Lỗi khởi động ExcelLike: ' + ((e as Error)?.stack || e));
  const box = document.createElement('div');
  box.className = 'boot-err';
  box.setAttribute('role', 'alert');
  box.innerHTML = '<h1>ExcelLike không khởi động được</h1><p class="m"></p><p>Mở ở chế độ an toàn sẽ dùng sổ trống và giao diện tiếng Việt. Sổ cũ vẫn được giữ trong Tệp ▸ Mở.</p><div><button class="safe">Mở chế độ an toàn</button> <button class="again">Thử lại</button></div>';
  box.querySelector('.m')!.textContent = String((e as Error)?.message || e);
  (box.querySelector('.safe') as HTMLElement).onclick = () => { location.search = '?safe=1'; };
  (box.querySelector('.again') as HTMLElement).onclick = () => location.reload();
  document.body.appendChild(box);
}
window.addEventListener('error', (e) => console.error('Lỗi: ' + (e.error?.stack || e.message)));
window.addEventListener('unhandledrejection', (e) => console.error('Lỗi (promise): ' + ((e.reason as Error)?.stack || e.reason)));

async function boot() {
  // chế độ an toàn (?safe=1): sổ trống, giao diện tiếng Việt — dùng khi lần mở trước bị treo/lỗi
  const safe = new URLSearchParams(location.search).has('safe');
  if (safe) setLang('vi');
  let id = (!safe && lsGet('xl-last')) || 'b' + Date.now();
  let wb: Workbook | null = null;
  let failed = false;
  let rec: Awaited<ReturnType<typeof loadBook>> | undefined;
  try {
    rec = await loadBook(id);
    if (rec) wb = Workbook.fromJSON(rec.data as any);
  } catch (e) {
    // không ghi đè bản cũ: mở sổ mới với mã khác, bản cũ vẫn nằm trong Tệp → Mở…
    console.warn('Không mở được sổ đã lưu', e);
    failed = true;
    id = 'b' + Date.now();
  }
  if (!wb) { wb = Workbook.create(); lsSet('xl-last', id); }
  const app = new App(wb, id);
  // tệp .xlsx vừa mở: gắn biểu đồ, kiểm công thức (cần sổ đã dựng) rồi lưu lại bản nháp
  if (rec && !failed && (rec.data as any)?.pendingImport) {
    import('./core/xlsx').then(({ finishImport }) => {
      const w = finishImport(app.wb, rec!.data as any);
      if (w.length) app.toast('Lưu ý khi mở tệp: ' + w.join(' '), 'err');
      app.charts?.sync();
      app.grid.invalidate();
      void app.saveNow(false);
    }).catch((e) => console.error('finishImport: ' + e));
  }
  if (safe) app.toast('Đang ở chế độ an toàn: lần mở trước bị lỗi nên ExcelLike mở sổ trống, giao diện tiếng Việt. Sổ cũ vẫn ở Tệp ▸ Mở.', 'err');
  if (rec && !failed) { app.fileHandle = rec.handle ?? null; app.fileName = rec.fileName ?? ''; app.fileDirty = !!rec.fileDirty; }
  if (failed) app.toast('Không mở được sổ tính lần trước nên đã mở sổ mới. Bản cũ vẫn được giữ trong Tệp → Mở…', 'err');
  (window as any).__xl = app;
  (window as any).__XLWorkbook = Workbook;
  (window as any).__xlKeys = KEYMAP;
  startI18n();
  void ensureLicense((m, k) => app.toast(m, k === 'err' ? 'err' : undefined));
  if ('serviceWorker' in navigator && location.protocol !== 'file:' && !(window as any).xlDesktop && !import.meta.env.DEV) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }
  document.fonts?.ready.then(() => app.grid.invalidate());
  void id;
}
boot().catch(bootError);
