// ExcelLike bản Windows (Electron): nạp bản build web qua giao thức app://, mở/lưu tệp bằng hộp thoại hệ thống
const { app, BrowserWindow, protocol, net, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const crypto = require('node:crypto');
const { pathToFileURL } = require('node:url');

// biến môi trường chỉ dùng cho kiểm thử, bản đóng gói bỏ qua
const TEST = !app.isPackaged ? process.env : {};
if (TEST.XL_USER_DATA) app.setPath('userData', TEST.XL_USER_DATA);
const DIST = path.join(__dirname, '..', 'dist');
const OPENABLE = /\.(xlsx|xlsm|csv|txt)$/i;
const CSP = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'; worker-src 'self' blob:; object-src 'none'; base-uri 'none'";

protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true } }]);

// chỉ một phiên chạy: mở tệp lần sau sẽ chuyển vào cửa sổ đang có
if (!app.requestSingleInstanceLock()) { app.quit(); return; }

let win = null;
let allowClose = false;
let closeReady = false;   // trang đã đăng ký xử lý đóng cửa sổ
let closeTimer = null;
const pending = [];       // tệp chờ mở (đường dẫn)
const fileArg = (argv, cwd = process.cwd()) => {
  const c = argv.slice(1).filter((a) => !a.startsWith('-') && OPENABLE.test(a));
  return c.length ? path.resolve(cwd, c[0]) : undefined;
};

// ---- chỉ cho ghi vào tệp mà người dùng đã chọn (mở / lưu thành / nhấp đúp), nhớ qua các lần chạy ----
const ALLOW_FILE = path.join(app.getPath('userData'), 'allowed-files.json');
let allowed = new Set();
try { allowed = new Set(JSON.parse(fsSync.readFileSync(ALLOW_FILE, 'utf8'))); } catch { /* lần đầu */ }
const key = (p) => path.resolve(p).toLowerCase();
function allow(p) {
  allowed.delete(key(p)); allowed.add(key(p));
  const list = [...allowed].slice(-300);
  allowed = new Set(list);
  fs.writeFile(ALLOW_FILE, JSON.stringify(list)).catch(() => {});
}

async function readFor(p) {
  const data = await fs.readFile(p);
  allow(p);
  return { path: p, name: path.basename(p), data: new Uint8Array(data) };
}
function sendOpen(p) {
  // trang đang nạp/nạp lại → để trang tự lấy khi khởi động xong
  if (!win || win.webContents.isLoading()) { pending.push(p); return; }
  if (win.isMinimized()) win.restore();
  win.focus();
  readFor(p).then((f) => win.webContents.send('xl:open', f)).catch((e) => dialog.showErrorBox('Không mở được tệp', String(e.message || e)));
}

function forceClose() { clearTimeout(closeTimer); if (win && !win.isDestroyed()) { allowClose = true; win.close(); } }

function createWindow() {
  win = new BrowserWindow({
    width: 1366, height: 860, minWidth: 360, minHeight: 480,
    title: 'ExcelLike', backgroundColor: '#eef2ef', show: false,
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: false },
  });
  win.once('ready-to-show', () => win.show());
  win.loadURL('app://excellike/index.html');
  win.webContents.on('did-start-loading', () => { closeReady = false; });
  // tệp nhấp đúp trong lúc trang đang nạp → gửi khi nạp xong
  win.webContents.on('did-finish-load', () => { setTimeout(() => { while (pending.length && win) sendOpen(pending.shift()); }, 300); });
  // liên kết ngoài → trình duyệt mặc định; không cho mở cửa sổ con hay điều hướng ra ngoài
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/i.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('app://excellike/')) { e.preventDefault(); if (/^https?:/i.test(url)) shell.openExternal(url); } });
  win.on('close', (e) => {
    if (allowClose) return;
    // trang hỏng, chưa nạp xong hoặc chưa sẵn sàng → đóng luôn (bản nháp đã tự lưu)
    if (!closeReady || win.webContents.isCrashed() || win.webContents.isLoading()) return;
    e.preventDefault();
    win.webContents.send('xl:close-request');
    // trang không trả lời (treo) trong 6 giây → vẫn cho đóng
    clearTimeout(closeTimer);
    closeTimer = setTimeout(forceClose, 6000);
  });
  win.on('unresponsive', () => { if (closeTimer) forceClose(); });
  win.on('closed', () => { win = null; });
}

ipcMain.on('xl:close-ready', () => { closeReady = true; });
ipcMain.on('xl:close-pending', () => clearTimeout(closeTimer)); // trang đang hỏi người dùng
ipcMain.on('xl:close-ok', () => forceClose());
ipcMain.on('xl:close-cancel', () => { clearTimeout(closeTimer); closeTimer = null; });

ipcMain.handle('xl:startup-file', async () => {
  const p = pending.shift();
  return p ? readFor(p).catch(() => null) : null;
});
ipcMain.handle('xl:open-dialog', async () => {
  let p = TEST.XL_TEST_OPEN;
  if (!p) {
    const r = await dialog.showOpenDialog(win, { title: 'Mở tệp', properties: ['openFile'], filters: [{ name: 'Excel hoặc CSV', extensions: ['xlsx', 'csv', 'txt'] }, { name: 'Mọi tệp', extensions: ['*'] }] });
    if (r.canceled || !r.filePaths[0]) return null;
    p = r.filePaths[0];
  }
  return readFor(p);
});

// ghi tuần tự theo từng tệp: hai lần lưu cùng lúc không chồng lên nhau
const queues = new Map();
function serial(target, job) {
  const k = key(target);
  const prev = queues.get(k) ?? Promise.resolve();
  const next = prev.catch(() => {}).then(job);
  queues.set(k, next);
  next.finally(() => { if (queues.get(k) === next) queues.delete(k); }).catch(() => {});
  return next;
}
async function writeAtomic(target, data) {
  // ghi ra tệp tạm (tên riêng mỗi lần) rồi đổi tên: mất điện giữa chừng không làm hỏng tệp cũ
  const tmp = `${target}.~xl-${crypto.randomUUID().slice(0, 8)}`;
  try {
    await fs.writeFile(tmp, typeof data === 'string' ? data : Buffer.from(data));
    try { await fs.rename(tmp, target); }
    catch { await fs.copyFile(tmp, target); }
  } finally {
    await fs.rm(tmp, { force: true }).catch(() => {});
  }
}
ipcMain.handle('xl:save', async (_e, { path: p, name, kind, data, saveAs }) => {
  // chỉ ghi đè thẳng vào tệp người dùng đã chọn; đường dẫn lạ → hỏi bằng hộp thoại
  let target = !saveAs && p && allowed.has(key(p)) ? p : null;
  if (!target) {
    target = TEST.XL_TEST_SAVE;
    if (!target) {
      const ext = kind === 'csv' ? 'csv' : 'xlsx';
      const safeName = String(name || 'So tinh').replace(/[\\/:*?"<>|]/g, '_');
      const r = await dialog.showSaveDialog(win, { title: 'Lưu thành', defaultPath: p ? path.join(path.dirname(p), safeName) : safeName, filters: [ext === 'csv' ? { name: 'CSV', extensions: ['csv'] } : { name: 'Sổ tính Excel', extensions: ['xlsx'] }] });
      if (r.canceled || !r.filePath) return null;
      target = r.filePath;
    }
  }
  await serial(target, () => writeAtomic(target, data));
  allow(target);
  return { path: target, name: path.basename(target) };
});

app.on('second-instance', (_e, argv, cwd) => { const p = fileArg(argv, cwd); if (p) sendOpen(p); else if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });
app.on('open-file', (e, p) => { e.preventDefault(); sendOpen(p); }); // macOS
app.on('window-all-closed', () => app.quit());

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  protocol.handle('app', async (req) => {
    const u = new URL(req.url);
    const rel = decodeURIComponent(u.pathname).replace(/^\/+/, '') || 'index.html';
    const full = path.normalize(path.join(DIST, rel));
    if (!full.startsWith(DIST + path.sep)) return new Response('Không tìm thấy', { status: 404 });
    const res = await net.fetch(pathToFileURL(full).toString());
    if (!/\.html$/i.test(full)) return res;
    // chính sách bảo mật nội dung cho trang: chỉ chạy mã của app
    const h = new Headers(res.headers);
    h.set('Content-Security-Policy', CSP);
    return new Response(res.body, { status: res.status, headers: h });
  });
  const p = fileArg(process.argv);
  if (p) pending.push(p);
  createWindow();
});
