// ExcelLike Key Manager — app riêng của người cấp key: giữ khóa ký (mã hóa bằng Windows DPAPI), sinh key, lưu sổ key.
'use strict';
const { app, BrowserWindow, ipcMain, dialog, safeStorage, shell, Menu, clipboard } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const L = require('./license.cjs');

const TEST = !app.isPackaged ? process.env : {};
if (TEST.XL_KM_DATA) app.setPath('userData', TEST.XL_KM_DATA);
if (!app.requestSingleInstanceLock()) { app.quit(); return; }

const DATA = () => app.getPath('userData');
const SIGN_FILE = () => path.join(DATA(), 'signing.bin');
const DB_FILE = () => path.join(DATA(), 'keys.json');
const LOG = (m) => { try { fs.appendFileSync(path.join(DATA(), 'keymgr-log.txt'), `[${new Date().toISOString()}] ${m}\n`); } catch { /* bỏ qua */ } };

// ---------- khóa ký ----------
let signing = null; // { pub, priv }
function encrypt(text) {
  if (safeStorage.isEncryptionAvailable()) return Buffer.concat([Buffer.from('DPAPI1'), safeStorage.encryptString(text)]);
  if (app.isPackaged && process.platform === 'win32') throw new Error('Windows không cho mã hóa khóa ký trên tài khoản này');
  return Buffer.from('PLAIN1' + text); // máy không hỗ trợ mã hóa (chỉ gặp khi kiểm thử trên Linux)
}
function decrypt(buf) {
  const tag = buf.subarray(0, 6).toString();
  if (tag === 'DPAPI1') return safeStorage.decryptString(buf.subarray(6));
  if (tag === 'PLAIN1') return buf.subarray(6).toString();
  throw new Error('Tệp khóa ký hỏng');
}
function loadSigning() {
  try { signing = JSON.parse(decrypt(fs.readFileSync(SIGN_FILE()))); } catch (e) { signing = null; if (fs.existsSync(SIGN_FILE())) LOG('Không đọc được khóa ký: ' + e.message); }
}
/** đọc tệp khóa ký (.json do người phát triển giao) */
function parseSigningFile(text) {
  const o = JSON.parse(text);
  if (!o || typeof o.priv !== 'string') throw new Error('Không phải tệp khóa ký ExcelLike');
  const pub = L.pubFromPriv(o.priv);
  if (o.pub && o.pub !== pub) throw new Error('Tệp khóa ký hỏng (khóa công khai không khớp)');
  return { pub, priv: o.priv };
}

// ---------- sổ key ----------
function emptyDb() { return { version: 1, nextSerial: 1, keys: [] }; }
/** Đọc sổ key. Sổ hỏng → dùng bản .bak; cả hai hỏng → báo lỗi (không bao giờ tự về sổ rỗng, tránh cấp trùng số) */
function readDb() {
  const tryRead = (f) => { const d = JSON.parse(fs.readFileSync(f, 'utf8')); if (!Array.isArray(d.keys)) throw new Error('sai cấu trúc'); return d; };
  let d;
  if (!fs.existsSync(DB_FILE())) d = fs.existsSync(DB_FILE() + '.bak') ? tryRead(DB_FILE() + '.bak') : emptyDb();
  else {
    try { d = tryRead(DB_FILE()); }
    catch (e) {
      LOG('Sổ key hỏng: ' + e.message);
      try { d = tryRead(DB_FILE() + '.bak'); LOG('Đã dùng bản dự phòng keys.json.bak'); }
      catch { throw new Error('Sổ key bị hỏng và không có bản dự phòng đọc được. Hãy dùng Cài đặt ▸ Khôi phục sổ key từ tệp sao lưu (thư mục backup).'); }
    }
  }
  // số kế tiếp luôn lớn hơn mọi số đã cấp
  d.nextSerial = Math.max(Number.isInteger(d.nextSerial) ? d.nextSerial : 1, ...d.keys.map((k) => (Number.isInteger(k.serial) ? k.serial + 1 : 1)));
  return d;
}
const DATE_RE = /^\d{4}-\d\d-\d\d$/;
/** kiểm một bản ghi key (khi khôi phục từ tệp) */
function validRec(k) {
  if (!k || !Number.isInteger(k.serial) || k.serial < 1 || k.serial > 0xffffffff || typeof k.key !== 'string') return null;
  const d = L.decodeKey(k.key);
  if (!d || d.serial !== k.serial) return null;
  // thông tin hiển thị lấy từ chính key (không tin dữ liệu trong tệp)
  return {
    serial: k.serial, key: k.key, type: d.type, machine: d.machine,
    issued: iso(L.dayDate(d.issued)), expires: d.expires ? iso(L.dayDate(d.expires)) : null,
    customer: clean(k.customer, 120) || '(không tên)', org: clean(k.org, 160), contact: clean(k.contact, 160), note: clean(k.note, 500),
    status: k.status === 'revoked' ? 'revoked' : 'active', revokedAt: typeof k.revokedAt === 'string' ? clean(k.revokedAt, 40) : undefined,
    created: typeof k.created === 'string' ? clean(k.created, 40) : new Date().toISOString(),
  };
}
function writeDb(d) {
  fs.mkdirSync(DATA(), { recursive: true });
  // giữ bản trước làm bản dự phòng, ghi qua tệp tạm để mất điện không hỏng sổ
  if (fs.existsSync(DB_FILE())) fs.copyFileSync(DB_FILE(), DB_FILE() + '.bak');
  const tmp = DB_FILE() + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(d, null, 1));
  fs.renameSync(tmp, DB_FILE());
  // sao lưu theo ngày (giữ 30 bản gần nhất)
  try {
    const bd = path.join(DATA(), 'backup'); fs.mkdirSync(bd, { recursive: true });
    fs.copyFileSync(DB_FILE(), path.join(bd, `keys-${new Date().toISOString().slice(0, 10)}.json`));
    const old = fs.readdirSync(bd).filter((f) => /^keys-\d{4}-\d\d-\d\d\.json$/.test(f)).sort();
    for (const f of old.slice(0, Math.max(0, old.length - 30))) fs.rmSync(path.join(bd, f));
  } catch (e) { LOG('Sao lưu: ' + e.message); }
}
const today = () => (TEST.XL_TEST_TODAY ? new Date(TEST.XL_TEST_TODAY) : new Date());
const parseDate = (s) => { const m = /^(\d{4})-(\d\d)-(\d\d)$/.exec(String(s || '')); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const clean = (s, n = 200) => String(s ?? '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);

function status(k) {
  if (k.status === 'revoked') return 'revoked';
  if (k.expires && L.dayNum(parseDate(k.expires)) < L.dayNum(today())) return 'expired';
  return 'active';
}

function state() {
  const db = readDb();
  return {
    signing: !!signing, pub: signing?.pub ?? null, version: app.getVersion(), dataDir: DATA(),
    keys: db.keys.map((k) => ({ ...k, state: status(k) })), nextSerial: db.nextSerial,
  };
}

ipcMain.handle('km:state', () => { try { return state(); } catch (e) { return { ...state0(), dbError: e.message }; } });
function state0() { return { signing: !!signing, pub: signing?.pub ?? null, version: app.getVersion(), dataDir: DATA(), keys: [], nextSerial: 0 }; }

ipcMain.handle('km:import-signing', async (_e, text) => {
  try {
    let t = text;
    if (!t) {
      const r = TEST.XL_KM_SIGNING_FILE ? { canceled: false, filePaths: [TEST.XL_KM_SIGNING_FILE] } : await dialog.showOpenDialog(win, { title: 'Chọn tệp khóa ký ExcelLike', filters: [{ name: 'Khóa ký', extensions: ['json', 'xlsign', 'txt'] }], properties: ['openFile'] });
      if (r.canceled || !r.filePaths[0]) return { ok: false, canceled: true };
      t = fs.readFileSync(r.filePaths[0], 'utf8');
    }
    const k = parseSigningFile(t);
    fs.mkdirSync(DATA(), { recursive: true });
    fs.writeFileSync(SIGN_FILE(), encrypt(JSON.stringify(k)));
    signing = k;
    LOG('Đã nhập khóa ký ' + k.pub.slice(0, 8) + '…');
    return { ok: true, pub: k.pub };
  } catch (e) { return { ok: false, error: e.message }; }
});

ipcMain.handle('km:export-signing', async () => {
  if (!signing) return { ok: false, error: 'Chưa có khóa ký' };
  const r = TEST.XL_KM_SAVE ? { canceled: false, filePath: TEST.XL_KM_SAVE } : await dialog.showSaveDialog(win, { title: 'Sao lưu khóa ký', defaultPath: 'ExcelLike-khoa-ky.json', filters: [{ name: 'Khóa ký', extensions: ['json'] }] });
  if (r.canceled || !r.filePath) return { ok: false, canceled: true };
  fs.writeFileSync(r.filePath, JSON.stringify({ product: 'ExcelLike', note: 'KHÓA BÍ MẬT — ai có tệp này đều cấp được key. Cất giữ cẩn thận.', pub: signing.pub, priv: signing.priv }, null, 2));
  return { ok: true, path: r.filePath };
});

ipcMain.handle('km:create', (_e, o) => {
  if (!signing) return { ok: false, error: 'Chưa nhập khóa ký' };
  try {
    const db = readDb();
    const type = o.type === 'machine' ? 'machine' : 'shared';
    const issued = today();
    const expires = o.expires ? parseDate(o.expires) : null;
    if (o.expires && !expires) throw new Error('Ngày hết hạn không hợp lệ');
    if (expires && L.dayNum(expires) < L.dayNum(issued)) throw new Error('Ngày hết hạn phải từ hôm nay trở đi');
    const customer = clean(o.customer, 120);
    if (!customer) throw new Error('Hãy nhập tên khách hàng / người dùng');
    let machine = null;
    if (type === 'machine') {
      const h = L.parseMachineCode(o.machine);
      if (!h) throw new Error('Mã máy không hợp lệ (13 ký tự, dạng XXXX-XXXX-XXXXX)');
      machine = L.machineCode(h);
    }
    const serial = db.nextSerial;
    const key = L.makeKey({ serial, type, machine, issued, expires }, signing.priv);
    const rec = { serial, key, type, machine, customer, org: clean(o.org, 160), contact: clean(o.contact, 160), issued: iso(issued), expires: expires ? iso(expires) : null, note: clean(o.note, 500), status: 'active', created: new Date().toISOString() };
    db.keys.push(rec);
    db.nextSerial = serial + 1;
    writeDb(db);
    LOG(`Cấp key số ${serial} cho ${customer}`);
    return { ok: true, rec: { ...rec, state: status(rec) } };
  } catch (e) { return { ok: false, error: e.message }; }
});

ipcMain.handle('km:update', (_e, serial, patch) => {
  let db;
  try { db = readDb(); } catch (e) { return { ok: false, error: e.message }; }
  const k = db.keys.find((x) => x.serial === serial);
  if (!k) return { ok: false, error: 'Không tìm thấy key' };
  for (const f of ['customer', 'org', 'contact', 'note']) if (f in patch) k[f] = clean(patch[f], f === 'note' ? 500 : 160);
  if (patch.status === 'revoked' || patch.status === 'active') { k.status = patch.status; k.revokedAt = patch.status === 'revoked' ? new Date().toISOString() : undefined; LOG(`Key số ${serial}: ${patch.status}`); }
  writeDb(db);
  return { ok: true };
});

ipcMain.handle('km:verify', (_e, text, machineCode) => {
  const d = L.decodeKey(text);
  if (!d) return { ok: false, code: 'format', message: L.MESSAGES.format };
  const pub = signing?.pub;
  if (!pub) return { ok: false, code: 'nosign', message: 'Chưa nhập khóa ký nên không kiểm được chữ ký.' };
  const m = machineCode ? L.parseMachineCode(machineCode) : null;
  const r = L.checkKey(text, { pub, machine: m || undefined, today: today() });
  let keys = []; try { keys = readDb().keys; } catch { /* sổ hỏng: chỉ kiểm chữ ký */ }
  const rec = keys.find((k) => k.serial === d.serial && L.normalize(k.key) === L.normalize(text));
  return { ...r, message: L.MESSAGES[r.code], rec: rec ? { ...rec, state: status(rec) } : null };
});

ipcMain.handle('km:copy', (_e, t) => { clipboard.writeText(String(t)); return true; });

ipcMain.handle('km:save-text', async (_e, name, text) => {
  const r = TEST.XL_KM_SAVE ? { canceled: false, filePath: TEST.XL_KM_SAVE } : await dialog.showSaveDialog(win, { title: 'Lưu', defaultPath: name, filters: [{ name: name.endsWith('.csv') ? 'CSV' : name.endsWith('.json') ? 'JSON' : 'Văn bản', extensions: [name.split('.').pop()] }] });
  if (r.canceled || !r.filePath) return { ok: false, canceled: true };
  fs.writeFileSync(r.filePath, String(text));
  return { ok: true, path: r.filePath };
});

ipcMain.handle('km:restore', async () => {
  const r = TEST.XL_KM_OPEN ? { canceled: false, filePaths: [TEST.XL_KM_OPEN] } : await dialog.showOpenDialog(win, { title: 'Khôi phục sổ key', filters: [{ name: 'Sổ key', extensions: ['json'] }], properties: ['openFile'] });
  if (r.canceled || !r.filePaths[0]) return { ok: false, canceled: true };
  try {
    const d = JSON.parse(fs.readFileSync(r.filePaths[0], 'utf8'));
    if (!d || !Array.isArray(d.keys)) throw new Error('Không phải tệp sổ key');
    // gộp: giữ key hiện có, thêm key chưa có; bản ghi sai/không khớp key bị bỏ qua và báo lại
    let db;
    try { db = readDb(); } catch { db = emptyDb(); } // khôi phục chính là cách sửa sổ hỏng
    const have = new Map(db.keys.map((k) => [k.serial, k]));
    let added = 0, bad = 0, conflict = 0;
    for (const raw of d.keys) {
      const k = validRec(raw);
      if (!k) { bad++; continue; }
      const cur = have.get(k.serial);
      if (cur) { if (L.normalize(cur.key) !== L.normalize(k.key)) conflict++; continue; }
      db.keys.push(k); have.set(k.serial, k); added++;
    }
    db.keys.sort((a, b) => a.serial - b.serial);
    db.nextSerial = Math.max(db.nextSerial, ...db.keys.map((k) => k.serial + 1));
    writeDb(db);
    if (conflict) LOG(`Khôi phục: ${conflict} key trùng số nhưng khác nội dung (giữ bản đang có)`);
    return { ok: true, added, bad, conflict };
  } catch (e) { return { ok: false, error: e.message }; }
});

ipcMain.handle('km:open-data', () => shell.openPath(DATA()));

let win = null;
function createWindow() {
  win = new BrowserWindow({
    width: 1240, height: 800, minWidth: 380, minHeight: 520, title: 'ExcelLike Key Manager', backgroundColor: '#f5f6f5', show: false,
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: false },
  });
  win.once('ready-to-show', () => win.show());
  setTimeout(() => { if (win && !win.isDestroyed() && !win.isVisible()) win.show(); }, 5000);
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  win.webContents.on('will-navigate', (e) => e.preventDefault());
  win.webContents.on('console-message', (_e, level, msg) => { if (level >= 3) LOG('Trang: ' + msg); });
  win.loadFile(path.join(__dirname, '..', 'ui', 'index.html'));
  win.on('closed', () => { win = null; });
}
app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });
app.on('window-all-closed', () => app.quit());
process.on('uncaughtException', (e) => { LOG('Lỗi: ' + (e?.stack || e)); try { dialog.showErrorBox('Key Manager gặp lỗi', String(e?.message || e)); } catch { /* bỏ qua */ } });
app.whenReady().then(() => { Menu.setApplicationMenu(null); loadSigning(); createWindow(); });
