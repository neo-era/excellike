// ExcelLike bản Windows (Electron): nạp bản build web qua giao thức app://, mở/lưu tệp bằng hộp thoại hệ thống
const { app, BrowserWindow, protocol, net, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const crypto = require('node:crypto');
const { pathToFileURL } = require('node:url');
const os = require('node:os');
const LIC = require('./license.cjs');
const LIC_CFG = require('./license-config.cjs');

// biến môi trường chỉ dùng cho kiểm thử, bản đóng gói bỏ qua
const TEST = !app.isPackaged ? process.env : {};
if (TEST.XL_USER_DATA) app.setPath('userData', TEST.XL_USER_DATA);
const DIST = path.join(__dirname, '..', 'dist');
const OPENABLE = /\.(xlsx|xlsm|csv|txt)$/i;
const CSP = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'; worker-src 'self' blob:; object-src 'none'; base-uri 'none'";

protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true } }]);

// ghi lỗi ra tệp nhật ký trong thư mục dữ liệu của app và báo cho người dùng thay vì im lặng thoát
function logErr(msg) {
  try { fsSync.appendFileSync(path.join(app.getPath('userData'), 'excellike-log.txt'), `[${new Date().toISOString()}] ${msg}\n`); } catch { /* bỏ qua */ }
}
process.on('uncaughtException', (e) => {
  logErr('Lỗi: ' + (e && e.stack || e));
  try { dialog.showErrorBox('ExcelLike gặp lỗi', String(e && e.message || e) + '\n\nChi tiết đã ghi vào: ' + path.join(app.getPath('userData'), 'excellike-log.txt')); } catch { /* bỏ qua */ }
});

// ---- khởi động an toàn ----
// Lần chạy trước không khởi động xong (cờ boot-pending còn) → lần này vào chế độ an toàn (sổ trống, tiếng Việt).
// Tiến trình đồ họa (GPU) hỏng nhiều lần → tắt tăng tốc phần cứng cho các lần sau.
const BOOT_FLAG = path.join(app.getPath('userData'), 'boot-pending');
const NOGPU_FLAG = path.join(app.getPath('userData'), 'disable-gpu');
let safeStart = process.argv.includes('--safe') || fsSync.existsSync(BOOT_FLAG);
if (fsSync.existsSync(NOGPU_FLAG) || process.argv.includes('--disable-gpu')) app.disableHardwareAcceleration();
let gpuCrashes = 0;
logErr(`Khởi động ExcelLike ${app.getVersion()}${safeStart ? ' (chế độ an toàn)' : ''}${fsSync.existsSync(NOGPU_FLAG) ? ' (không tăng tốc đồ họa)' : ''}`);
app.on('child-process-gone', (_e, d) => {
  logErr('Tiến trình con dừng: ' + JSON.stringify(d));
  // đồ họa hỏng 2 lần → các lần sau chạy không dùng card đồ họa; đang khởi động dở thì mở lại ngay
  if (d.type === 'GPU' && d.reason !== 'clean-exit' && ++gpuCrashes >= 2 && !fsSync.existsSync(NOGPU_FLAG)) {
    try { fsSync.writeFileSync(NOGPU_FLAG, '1'); } catch { /* bỏ qua */ }
    if (!rendererReady) { try { fsSync.rmSync(BOOT_FLAG, { force: true }); } catch { /* bỏ qua */ } app.relaunch(); app.exit(0); }
  }
});

// chỉ một phiên chạy: mở tệp lần sau sẽ chuyển vào cửa sổ đang có
if (!app.requestSingleInstanceLock()) { app.quit(); return; }

let win = null;
let allowClose = false;
let closeReady = false;   // trang đã đăng ký xử lý đóng cửa sổ
let closeTimer = null;
const pending = [];       // tệp chờ mở (đường dẫn)
let rendererReady = false; // trang đã khởi động xong và tự hỏi tệp chờ mở
let bootTimer = null;
let intentionalCrash = false, crashCount = 0;
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
  // trang chưa sẵn sàng (đang nạp/nạp lại) → để trang tự lấy khi khởi động xong; mỗi tệp chỉ giao một lần
  if (!win || !rendererReady) { pending.push(p); return; }
  if (win.isMinimized()) win.restore();
  win.focus();
  readFor(p).then((f) => win.webContents.send('xl:open', f)).catch((e) => dialog.showErrorBox('Không mở được tệp', String(e.message || e)));
}

function forceClose() { clearTimeout(closeTimer); if (win && !win.isDestroyed()) { allowClose = true; win.close(); } }

function createWindow() {
  win = new BrowserWindow({
    width: 1366, height: 860, minWidth: 360, minHeight: 480,
    title: 'ExcelLike', backgroundColor: '#f3f3f3', show: false,
    // thanh tiêu đề kiểu Excel 365: app tự vẽ (QAT, tên tệp, Tìm kiếm), nút cửa sổ của Windows nằm đè bên phải
    titleBarStyle: 'hidden',
    titleBarOverlay: { color: '#107c41', symbolColor: '#ffffff', height: 44 },
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: false },
  });
  win.once('ready-to-show', () => win.show());
  // phòng khi trang không báo sẵn sàng (lỗi đồ họa, lỗi nạp trang): vẫn hiện cửa sổ sau 6 giây
  setTimeout(() => { if (win && !win.isDestroyed() && !win.isVisible()) win.show(); }, 6000);
  win.webContents.on('did-fail-load', (_e, code, desc, url) => logErr(`Không nạp được trang ${url}: ${code} ${desc}`));
  win.webContents.on('render-process-gone', (_e, d) => {
    logErr('Trang bị dừng: ' + JSON.stringify(d));
    if (!win || win.isDestroyed()) return;
    // chủ động dừng để vào chế độ an toàn → nạp trang an toàn, không nạp lại trang cũ
    if (intentionalCrash) { intentionalCrash = false; loadApp(); return; }
    if (d.reason === 'clean-exit') return;
    // dừng khi đang khởi động → lần nạp lại vào chế độ an toàn (tránh lặp lỗi với sổ nặng/hỏng)
    if (!rendererReady) safeStart = true;
    if (++crashCount > 3) { logErr('Trang dừng quá nhiều lần, không tự nạp lại'); return; }
    loadApp();
  });
  // trang chưa khởi động xong sau 25 giây (sổ quá nặng, lỗi đồ họa, lỗi dữ liệu) → hỏi người dùng thay vì để màn hình trống
  const watchdog = () => {
    clearTimeout(bootTimer);
    bootTimer = setTimeout(async () => {
      if (rendererReady || !win || win.isDestroyed()) return;
      logErr(`Trang chưa khởi động xong sau 25 giây${safeStart ? ' (chế độ an toàn)' : ''}`);
      if (!win.isVisible()) win.show();
      const r = await dialog.showMessageBox(win, {
        type: 'warning', title: 'ExcelLike', noLink: true, cancelId: 0,
        message: 'ExcelLike đang mở lâu hơn bình thường.',
        detail: 'Có thể sổ tính lần trước quá nặng, hoặc card đồ họa của máy không tương thích.\n\n'
          + '• Chờ thêm: tiếp tục đợi.\n• Chế độ an toàn: mở sổ trống (sổ cũ vẫn giữ trong Tệp ▸ Mở).\n• Tắt tăng tốc đồ họa: khởi động lại ExcelLike không dùng card đồ họa.\n\n'
          + 'Nhật ký: ' + path.join(app.getPath('userData'), 'excellike-log.txt'),
        buttons: ['Chờ thêm', 'Chế độ an toàn', 'Tắt tăng tốc đồ họa', 'Đóng'],
      });
      if (rendererReady || !win || win.isDestroyed()) return;
      logErr('Người dùng chọn: ' + ['chờ thêm', 'chế độ an toàn', 'tắt tăng tốc đồ họa', 'đóng'][r.response]);
      if (r.response === 0) watchdog();
      else if (r.response === 1) {
        safeStart = true;
        // trang có thể đang treo: dừng hẳn tiến trình trang rồi nạp bản an toàn (trong render-process-gone)
        intentionalCrash = true;
        try { win.webContents.forcefullyCrashRenderer(); } catch { intentionalCrash = false; loadApp(); }
      }
      else if (r.response === 2) { try { fsSync.writeFileSync(NOGPU_FLAG, '1'); fsSync.rmSync(BOOT_FLAG, { force: true }); } catch { /* bỏ qua */ } app.relaunch(); app.exit(0); }
      else { allowClose = true; win.close(); }
    }, TEST.XL_TEST_BOOT_WAIT ? +TEST.XL_TEST_BOOT_WAIT : 25000);
  };
  const loadApp = () => { rendererReady = false; watchdog(); win.loadURL('app://excellike/index.html' + (safeStart ? '?safe=1' : '')); };
  try { fsSync.writeFileSync(BOOT_FLAG, String(Date.now())); } catch { /* bỏ qua */ }
  // lỗi trong trang (console.error, lỗi chưa bắt) → ghi nhật ký để chẩn đoán
  win.webContents.on('console-message', (_e, level, message, line, source) => { if (level >= 3) logErr(`Trang: ${message} (${source}:${line})`); });
  win.webContents.on('unresponsive', () => logErr('Trang không phản hồi'));
  loadApp();
  // chỉ khi chính trang app nạp lại (không tính khung con như bản xem trước khi in)
  win.webContents.on('did-start-navigation', (e, _url, isInPlace, isMainFrame) => {
    const main = e?.isMainFrame ?? isMainFrame, same = e?.isSameDocument ?? isInPlace;
    if (main && !same) { closeReady = false; rendererReady = false; }
  });
  // tệp nhấp đúp trong lúc trang đang nạp → gửi khi nạp xong
  // liên kết ngoài → trình duyệt mặc định; không cho mở cửa sổ con hay điều hướng ra ngoài
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/i.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('app://excellike/')) { e.preventDefault(); if (/^https?:/i.test(url)) shell.openExternal(url); } });
  win.on('close', (e) => {
    if (allowClose) return;
    // chưa kích hoạt (màn Kích hoạt đang chặn giao diện) → đóng luôn, bản nháp đã tự lưu
    if (!licensed) return;
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

// ---- Bản quyền: key ký số, kiểm ngay trên máy (không cần mạng) ----
const LIC_FILE = path.join(app.getPath('userData'), 'license.key');
const LIC_STATE = path.join(app.getPath('userData'), 'license-state.json');
let machineCache = null;
function machineGuid() {
  if (TEST.XL_TEST_MACHINE) return TEST.XL_TEST_MACHINE;
  if (process.platform === 'win32') {
    try {
      const out = require('node:child_process').execFileSync('reg', ['query', 'HKLM\\SOFTWARE\\Microsoft\\Cryptography', '/v', 'MachineGuid', '/reg:64'], { encoding: 'utf8', windowsHide: true });
      const m = /MachineGuid\s+REG_SZ\s+(\S+)/i.exec(out);
      if (m) return m[1];
    } catch (e) { logErr('Không đọc được MachineGuid: ' + e); }
  }
  if (process.platform === 'win32') return null; // không đọc được: không tự đoán (tránh lệch Mã máy với bộ cài)
  try { return fsSync.readFileSync('/etc/machine-id', 'utf8').trim(); } catch { return os.hostname(); }
}
function machine() {
  if (machineCache) return machineCache;
  // Windows: dùng chính lic.exe của bộ cài để Mã máy luôn khớp (không phụ thuộc reg.exe bị chặn bởi chính sách)
  if (process.platform === 'win32' && !TEST.XL_TEST_MACHINE) {
    try {
      const out = require('node:child_process').execFileSync(path.join(process.resourcesPath, 'lic.exe'), ['machine'], { encoding: 'utf8', windowsHide: true, timeout: 10000 }).trim();
      const h = LIC.parseMachineCode(out);
      if (h) return (machineCache = { hash: h, code: LIC.machineCode(h) });
    } catch (e) { logErr('lic.exe machine: ' + e); }
  }
  const g = machineGuid();
  if (!g) return (machineCache = { hash: null, code: 'không đọc được' });
  const h = LIC.machineHash(g);
  return (machineCache = { hash: h, code: LIC.machineCode(h) });
}
function licState() { try { return JSON.parse(fsSync.readFileSync(LIC_STATE, 'utf8')); } catch { return {}; } }
const licToday = () => (TEST.XL_TEST_TODAY ? new Date(TEST.XL_TEST_TODAY) : new Date());
function licCheck(key, ignoreLastSeen = false) {
  const st = licState(), m = machine(), today = LIC.dayNum(licToday());
  // ngày dùng gần nhất quá xa trong tương lai (đồng hồ từng bị đặt sai năm) → bỏ qua để không khóa oan
  const lastSeen = ignoreLastSeen || !st.lastSeen || st.lastSeen > today + 366 ? undefined : st.lastSeen;
  const d = LIC.decodeKey(key);
  if (d && d.type === 'machine' && !m.hash) return { ok: false, code: 'machine', message: 'Không đọc được Mã máy của máy này nên không dùng được key gắn máy. Hãy dùng key “dùng mọi máy” hoặc báo người hỗ trợ.', machine: m.code };
  const r = LIC.checkKey(key, { pub: LIC_CFG.PUBLIC_KEY, machine: m.hash || undefined, revoked: LIC_CFG.REVOKED, lastSeen, today: licToday() });
  return { ...r, message: LIC.MESSAGES[r.code], machine: m.code };
}
function licStatus() {
  let key = '';
  try { key = fsSync.readFileSync(LIC_FILE, 'utf8').trim(); } catch { /* chưa có key */ }
  if (!key) return { ok: false, code: 'missing', message: 'Chưa nhập key bản quyền.', machine: machine().code };
  const r = licCheck(key);
  if (r.ok) {
    // ghi ngày dùng gần nhất (chống chỉnh lùi đồng hồ để kéo dài key có hạn)
    const today = LIC.dayNum(licToday()), st = licState();
    if (!st.lastSeen || today > st.lastSeen || st.lastSeen > today + 366) { try { fsSync.writeFileSync(LIC_STATE, JSON.stringify({ ...st, lastSeen: today })); } catch { /* bỏ qua */ } }
  }
  return r;
}
let licensed = false;
ipcMain.handle('xl:license', () => { const r = licStatus(); licensed = r.ok; return r; });
ipcMain.handle('xl:activate', (_e, key) => {
  const k = String(key || '').slice(0, 600).trim();
  // key mới (gia hạn/đổi key) được kiểm theo đồng hồ hiện tại; nhập lại đúng key cũ thì vẫn kiểm chống lùi đồng hồ
  let cur = null; try { cur = LIC.decodeKey(fsSync.readFileSync(LIC_FILE, 'utf8').trim()); } catch { /* chưa có */ }
  const d = LIC.decodeKey(k);
  const r = licCheck(k, !!d && (!cur || cur.serial !== d.serial));
  if (r.ok) {
    licensed = true;
    try { fsSync.writeFileSync(LIC_STATE, JSON.stringify({ ...licState(), lastSeen: LIC.dayNum(licToday()) })); } catch { /* bỏ qua */ }
    try { fsSync.mkdirSync(path.dirname(LIC_FILE), { recursive: true }); fsSync.writeFileSync(LIC_FILE, LIC.decodeKey(k) ? k.replace(/\s+/g, ' ') : k); }
    catch (e) { return { ok: false, code: 'write', message: 'Không lưu được key: ' + e.message, machine: machine().code }; }
    logErr('Đã kích hoạt key số ' + r.info.serial);
  }
  return r;
});

ipcMain.on('xl:close-ready', () => { closeReady = true; });
// màu nút cửa sổ theo chủ đề Office (Nhiều màu / Trắng / Xám Đậm / Đen)
let lastOverlay = null;
ipcMain.on('xl:titlebar', (_e, o) => {
  if (!win || !o || !/^#[0-9a-f]{6}$/i.test(String(o.color)) || !/^#[0-9a-f]{6}$/i.test(String(o.symbolColor))) return;
  lastOverlay = { color: o.color, symbolColor: o.symbolColor };
  try { win.setTitleBarOverlay({ ...lastOverlay, height: 44 }); } catch { /* hệ điều hành không hỗ trợ */ }
  win.setBackgroundColor(o.color);
});
if (!app.isPackaged) global.__xlOverlay = () => lastOverlay;

// ---- In ấn: in thẳng ra máy in đã chọn (không qua hộp thoại), xuất PDF thật ----
async function renderJob(html) {
  const tmp = path.join(os.tmpdir(), `excellike-print-${crypto.randomUUID().slice(0, 8)}.html`);
  await fs.writeFile(tmp, String(html));
  const w = new BrowserWindow({ show: false, width: 1000, height: 1400, webPreferences: { sandbox: true, javascript: false, contextIsolation: true, nodeIntegration: false } });
  try {
    await w.loadFile(tmp);
    // đợi ảnh biểu đồ (data:) và phông tải xong
    await new Promise((r) => setTimeout(r, 150));
    return { w, done: () => { if (!w.isDestroyed()) w.destroy(); fs.rm(tmp, { force: true }).catch(() => {}); } };
  } catch (e) {
    if (!w.isDestroyed()) w.destroy();
    fs.rm(tmp, { force: true }).catch(() => {});
    throw e;
  }
}
const PAPER = { A3: 'A3', A4: 'A4', A5: 'A5', Letter: 'Letter', Legal: 'Legal', Tabloid: 'Tabloid' };
ipcMain.handle('xl:printers', async () => {
  if (!win) return [];
  try { return (await win.webContents.getPrintersAsync()).map((p) => ({ name: p.name, displayName: p.displayName, isDefault: !!p.isDefault })); } catch { return []; }
});
ipcMain.handle('xl:print', async (_e, o) => {
  if (!licensed) return { ok: false, error: 'Chưa kích hoạt bản quyền' };
  if (!o || typeof o.html !== 'string' || o.html.length > 200e6) return { ok: false, error: 'Dữ liệu in không hợp lệ' };
  let job;
  try { job = await renderJob(o.html); } catch (e) { return { ok: false, error: String(e.message || e) }; }
  return new Promise((resolve) => {
    const opts = {
      silent: o.silent !== false, printBackground: true, deviceName: typeof o.deviceName === 'string' ? o.deviceName : '',
      copies: Math.max(1, Math.min(999, o.copies | 0 || 1)), collate: o.collate !== false,
      margins: { marginType: 'none' }, landscape: !!o.landscape,
      ...(PAPER[o.paper] ? { pageSize: PAPER[o.paper] } : {}),
    };
    if (TEST.XL_TEST_PRINT) { fs.writeFile(TEST.XL_TEST_PRINT, JSON.stringify({ ...opts, pages: (o.html.match(/class="page"/g) || []).length })).then(() => { job.done(); resolve({ ok: true }); }); return; }
    try { job.w.webContents.print(opts, (ok, err) => { job.done(); resolve({ ok, error: ok ? undefined : err }); }); }
    catch (e) { job.done(); resolve({ ok: false, error: String(e.message || e) }); }
  });
});
ipcMain.handle('xl:pdf', async (_e, o) => {
  if (!licensed) return { error: 'Chưa kích hoạt bản quyền' };
  if (!o || typeof o.html !== 'string' || o.html.length > 200e6) return { error: 'Dữ liệu không hợp lệ' };
  let target = TEST.XL_TEST_PDF;
  if (!target) {
    const safeName = String(o.name || 'So tinh.pdf').replace(/[\\/:*?"<>|]/g, '_');
    const r = await dialog.showSaveDialog(win, { title: 'Xuất PDF', defaultPath: safeName, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
    if (r.canceled || !r.filePath) return null;
    target = r.filePath;
  }
  let job;
  try {
    job = await renderJob(o.html);
    const pdf = await job.w.webContents.printToPDF({ preferCSSPageSize: true, printBackground: true, margins: { marginType: 'none' } });
    await serial(target, () => writeAtomic(target, pdf));
    return { path: target };
  } catch (e) {
    return { error: String(e.message || e) };
  } finally { job?.done(); }
});
// Ctrl+F9 thu nhỏ, Ctrl+F10 phóng to/khôi phục cửa sổ (như Excel)
ipcMain.on('xl:win', (e, a) => {
  try {
    const w = BrowserWindow.fromWebContents(e.sender);
    if (!w) return;
    if (a === 'min') w.minimize();
    else if (a === 'max') { if (w.isMaximized()) w.unmaximize(); else w.maximize(); }
  } catch (err) { logErr('xl:win ' + err); }
});
// Mở liên kết trong ô (http, https, mailto) bằng trình duyệt/ứng dụng mặc định
ipcMain.on('xl:open-external', (_e, u) => {
  try { if (/^(https?:|mailto:)/i.test(u)) shell.openExternal(u); } catch (err) { logErr('xl:open-external ' + err); }
});
ipcMain.on('xl:close-pending', () => clearTimeout(closeTimer)); // trang đang hỏi người dùng
ipcMain.on('xl:close-ok', () => forceClose());
ipcMain.on('xl:close-cancel', () => { clearTimeout(closeTimer); closeTimer = null; });

ipcMain.handle('xl:startup-file', async () => {
  rendererReady = true;
  clearTimeout(bootTimer);
  try { fsSync.rmSync(BOOT_FLAG, { force: true }); } catch { /* bỏ qua */ }
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
  if (!licensed) throw new Error('Chưa kích hoạt bản quyền ExcelLike');
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

app.on('second-instance', (_e, argv, cwd) => {
  const p = fileArg(argv, cwd);
  // phiên đang chạy nhưng không còn cửa sổ (đang thoát dở) → mở lại cửa sổ
  if (!win || win.isDestroyed()) { if (p) pending.push(p); if (app.isReady()) createWindow(); return; }
  if (p) sendOpen(p); else { if (win.isMinimized()) win.restore(); win.show(); win.focus(); }
});
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
