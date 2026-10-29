import { KEYMAP } from './ui/keymap';
// phông Be Vietnam Pro đóng gói kèm app (chạy offline, không gọi Google Fonts)
import '@fontsource/be-vietnam-pro/400.css';
import '@fontsource/be-vietnam-pro/500.css';
import '@fontsource/be-vietnam-pro/600.css';
import '@fontsource/be-vietnam-pro/700.css';
import './styles.css';
import { Workbook } from './core/workbook';
import { App } from './ui/app';
import { loadBook, lsGet, lsSet } from './ui/store';

async function boot() {
  let id = lsGet('xl-last') || 'b' + Date.now();
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
  if (rec && !failed) { app.fileHandle = rec.handle ?? null; app.fileName = rec.fileName ?? ''; app.fileDirty = !!rec.fileDirty; }
  if (failed) app.toast('Không mở được sổ tính lần trước nên đã mở sổ mới. Bản cũ vẫn được giữ trong Tệp → Mở…', 'err');
  (window as any).__xl = app;
  (window as any).__XLWorkbook = Workbook;
  (window as any).__xlKeys = KEYMAP;
  if ('serviceWorker' in navigator && location.protocol !== 'file:' && !(window as any).xlDesktop && !import.meta.env.DEV) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }
  document.fonts?.ready.then(() => app.grid.invalidate());
  void id;
}
boot();
