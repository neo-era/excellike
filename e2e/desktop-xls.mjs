// Bản máy tính (Electron): mở tệp .xls từ dòng lệnh; Lưu không bao giờ ghi đè tệp .xls gốc
// Chạy: xvfb-run -a node e2e/desktop-xls.mjs   (cần npm run build trước)
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { testKey, TEST_GUID } from './testkey.mjs';
const require = createRequire(import.meta.url);
const { _electron: electron } = require('playwright');
const electronBin = require('electron');
const ROOT = new URL('..', import.meta.url).pathname;
const work = mkdtempSync(join(tmpdir(), 'xl-xls-'));
const userData = join(work, 'userdata');
mkdirSync(userData, { recursive: true });
writeFileSync(join(userData, 'license.key'), testKey());
const fx = join(work, 'Bao cao cu.xls');
copyFileSync(join(ROOT, 'tests/fixtures/basic.xls'), fx);
const sha = () => createHash('sha256').update(readFileSync(fx)).digest('hex');
const before = sha();
const results = [];
const check = (name, ok, info = '') => { results.push(ok); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };
// hộp lưu giả trả về chính đường dẫn tệp .xls gốc (người dùng gõ nhầm)
const env = { ...process.env, XL_TEST_MACHINE: TEST_GUID, XL_USER_DATA: userData, XL_TEST_SAVE: fx };
const app = await electron.launch({ executablePath: electronBin, args: [ROOT, fx], env });
const win = await app.firstWindow();
const errors = [];
win.on('pageerror', (e) => errors.push(e.message));
await win.waitForFunction(() => window.__xl && window.__xl.fileName === 'Bao cao cu.xls', null, { timeout: 20000 });
const X = (f, a) => win.evaluate(f, a);
check('Mở tệp .xls truyền qua dòng lệnh', await X(() => { const x = window.__xl; return x.wb.display(x.grid.sheet, 4, 1).text === '2 x 55'; }));
await win.keyboard.press('Control+s');
await win.waitForFunction(() => /\.xlsx$/.test(window.__xl.fileName), null, { timeout: 10000 }).catch(() => {});
check('Ctrl+S lưu thành .xlsx mới', existsSync(fx.replace(/\.xls$/, '.xlsx')) && readFileSync(fx.replace(/\.xls$/, '.xlsx'))[0] === 0x50, await X(() => window.__xl.fileName));
check('Tệp .xls gốc không bị ghi đè', sha() === before);
check('Không lỗi trang', errors.length === 0, errors.join('; '));
await app.close();
console.log(`\n${results.filter(Boolean).length}/${results.length} kiểm thử .xls bản máy tính đạt`);
process.exit(results.every(Boolean) ? 0 : 1);
