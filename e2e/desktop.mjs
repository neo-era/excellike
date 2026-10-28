// Kiểm thử bản máy tính (Electron): mở tệp từ dòng lệnh, lưu đè đúng tệp, Lưu thành…, Ctrl+O, hỏi khi đóng, một cửa sổ duy nhất, an toàn
// Chạy: xvfb-run -a node e2e/desktop.mjs   (cần npm run build trước)
import { createRequire } from 'node:module';
import { execFileSync, spawn } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const require = createRequire(import.meta.url);
const { _electron: electron } = require('playwright');
const electronBin = require('electron');

const ROOT = new URL('..', import.meta.url).pathname;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const work = mkdtempSync(join(tmpdir(), 'xl-desk-'));
const userData = join(work, 'userdata');
const fx = join(work, 'Dự toán mẫu.xlsx');
copyFileSync(join(ROOT, 'tests/fixtures/mau.xlsx'), fx);
const csv = join(work, 'vat-tu.csv');
writeFileSync(csv, '﻿Mã;Tên;Số lượng\nVT01;Cáp đồng;12,5\nVT02;Bóng LED;40\n');
const saveAs = join(work, 'ban-moi.xlsx');

const results = [];
const check = (name, ok, info = '') => { results.push({ name, ok }); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };
const py = (code, ...args) => execFileSync('python3', ['-c', code, ...args]).toString().trim();
const env = { ...process.env, XL_USER_DATA: userData, XL_TEST_OPEN: csv, XL_TEST_SAVE: saveAs };

const app = await electron.launch({ executablePath: electronBin, args: [ROOT, fx], env });
const win = await app.firstWindow();
const errors = [];
win.on('pageerror', (e) => errors.push(e.message));
await win.waitForFunction(() => window.__xl && window.__xl.wb.sheetIds().length > 1, null, { timeout: 20000 });
const X = (f, a) => win.evaluate(f, a);

const info = await X(() => ({ sheets: window.__xl.wb.sheetIds().map((i) => window.__xl.wb.sheetName(i)), file: window.__xl.fileName, desk: !!window.xlDesktop, req: typeof window.require, proc: typeof window.process }));
check('Mở tệp .xlsx truyền qua dòng lệnh (nhấp đúp tệp trong Windows)', info.sheets[0] === 'Dự toán' && info.file === 'Dự toán mẫu.xlsx', JSON.stringify(info));
check('An toàn: trang không truy cập được Node (require/process)', info.desk && info.req === 'undefined' && info.proc === 'undefined');
const title = await win.title();
check('Tiêu đề cửa sổ có tên sổ tính', /Dự toán mẫu/.test(title), title);
await win.screenshot({ path: OUT + 'd4-desktop.png' });

// sửa ô rồi Ctrl+S → ghi đè đúng tệp đã mở (không hỏi)
await X(() => { const x = window.__xl; x.wb.setInput(x.grid.sheet, 0, 0, 'ĐÃ SỬA'); });
await win.keyboard.press('Control+s');
await win.waitForFunction(() => !window.__xl.fileDirty, null, { timeout: 10000 }).catch(() => {});
const a1 = py('import openpyxl,sys;print(openpyxl.load_workbook(sys.argv[1]).worksheets[0]["A1"].value)', fx);
check('Ctrl+S lưu đè vào chính tệp .xlsx đã mở', a1 === 'ĐÃ SỬA', a1);

// Lưu thành… → hộp thoại lưu của Windows (giả lập bằng XL_TEST_SAVE)
await X(() => window.__xl.saveXlsx(true));
await win.waitForFunction(() => window.__xl.fileName === 'ban-moi.xlsx', null, { timeout: 10000 }).catch(() => {});
check('Lưu thành… ghi tệp mới và chuyển sang tệp đó', existsSync(saveAs) && (await X(() => window.__xl.fileName)) === 'ban-moi.xlsx');

// Ctrl+O → hộp thoại mở (giả lập) chọn CSV → sheet mới
await win.keyboard.press('Control+o');
await win.waitForFunction(() => window.__xl.wb.sheetIds().some((i) => window.__xl.wb.sheetName(i) === 'vat-tu'), null, { timeout: 10000 }).catch(() => {});
const c = await X(() => { const x = window.__xl, id = x.wb.sheetIds().find((i) => x.wb.sheetName(i) === 'vat-tu'); return id === undefined ? null : [x.wb.display(id, 1, 1).text, x.wb.value(id, 1, 2)]; });
check('Ctrl+O mở CSV (dấu ; và số thập phân dấu phẩy) thành sheet mới', JSON.stringify(c) === '["Cáp đồng",12.5]', JSON.stringify(c));

// liên kết ngoài mở bằng trình duyệt, không mở cửa sổ mới trong app
const nWin = app.windows().length;
await X(() => window.open('https://example.com', '_blank'));
await win.waitForTimeout(400);
check('Liên kết ngoài không mở cửa sổ mới trong app', app.windows().length === nWin);

// chạy lần hai với tệp khác → mở trong cửa sổ đang chạy (một phiên duy nhất)
const fx2 = join(work, 'bieu-do.xlsx');
copyFileSync(join(ROOT, 'tests/fixtures/bieudo.xlsx'), fx2);
const second = spawn(electronBin, [...(process.getuid?.() === 0 ? ['--no-sandbox'] : []), ROOT, fx2], { env, stdio: 'ignore' });
const secondExit = await new Promise((r) => { second.on('exit', (code) => r(code)); setTimeout(() => r('timeout'), 15000); });
await win.waitForFunction(() => window.__xl && window.__xl.fileName === 'bieu-do.xlsx', null, { timeout: 15000 }).catch(() => {});
const charts = await X(() => (window.__xl ? window.__xl.wb.meta(window.__xl.grid.sheet).charts.length : -1));
check('Mở tệp thứ hai: chuyển vào cửa sổ đang chạy, có biểu đồ', secondExit === 0 && charts > 0, `exit=${secondExit} charts=${charts}`);
check('Chỉ một cửa sổ', app.windows().length === 1);

// lưu hai lần cùng lúc (giữ Ctrl+S) → tệp vẫn hợp lệ, không báo lỗi giả
await X(() => { const x = window.__xl; x.wb.setInput(x.grid.sheet, 31, 0, 'lần 2'); });
const both = await X(() => Promise.all([window.__xl.saveXlsx(false), window.__xl.saveXlsx(false), window.__xl.saveXlsx(false)]));
const okFile = py('import openpyxl,sys;print(openpyxl.load_workbook(sys.argv[1]).worksheets[0]["A32"].value)', join(work, 'bieu-do.xlsx'));
check('Lưu dồn dập: chỉ ghi một lần, tệp hợp lệ', both.every(Boolean) && okFile === 'lần 2', `${JSON.stringify(both)} ${okFile}`);
// trang không tự ghi được vào đường dẫn lạ
const evil = join(work, 'khong-duoc-ghi.xlsx');
await X((p) => window.xlDesktop.save({ path: p, name: 'x.xlsx', kind: 'xlsx', data: new Uint8Array([1, 2, 3]) }), evil);
check('An toàn: không ghi thẳng vào đường dẫn chưa được người dùng chọn', !existsSync(evil));
const noTmp = execFileSync('ls', ['-a', work]).toString().split('\n').filter((f) => f.includes('.~xl'));
check('Không để lại tệp tạm cạnh tệp của người dùng', noTmp.length === 0, noTmp.join(','));

// đóng khi còn thay đổi chưa lưu vào tệp → hỏi
await X(() => { const x = window.__xl; x.wb.setInput(x.grid.sheet, 30, 0, 'chưa lưu'); });
const closeWin = () => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]?.close());
const ask = async () => { await win.waitForSelector('.dlg h2:has-text("Lưu thay đổi")', { timeout: 5000 }).catch(() => {}); return win.locator('.dlg h2:has-text("Lưu thay đổi")').count(); };
await closeWin();
check('Đóng cửa sổ khi chưa lưu: hỏi “Lưu thay đổi…?”', (await ask()) === 1);
await win.screenshot({ path: OUT + 'd4-desktop-close.png' });
await win.click('.dlg .btn:has-text("Hủy")').catch(() => {});
await win.waitForTimeout(300);
check('Bấm Hủy: cửa sổ vẫn mở', app.windows().length === 1 && !(await win.isClosed()));
await closeWin();
await ask();
await win.keyboard.press('Escape');
await win.waitForTimeout(7000); // quá thời gian chờ của tiến trình chính
check('Bấm Esc: cửa sổ vẫn mở (không tự đóng)', !(await win.isClosed()));
await closeWin();
check('Sau Esc vẫn đóng/hỏi lại được (không bị treo)', (await ask()) === 1);
const closed = app.waitForEvent('close', { timeout: 15000 }).then(() => true).catch(() => false);
await win.click('.dlg .btn:has-text("Không lưu")').catch(() => {});
check('Bấm Không lưu: app thoát', await closed);
check('Không có lỗi JavaScript', errors.length === 0, errors.join(' | '));

// mở lại app: vẫn nhớ là tệp chưa có thay đổi → đóng lại vẫn hỏi
const app2 = await electron.launch({ executablePath: electronBin, args: [ROOT], env });
const win2 = await app2.firstWindow();
await win2.waitForFunction(() => window.__xl, null, { timeout: 20000 });
const st = await win2.evaluate(() => ({ dirty: window.__xl.fileDirty, file: window.__xl.fileName, a31: window.__xl.wb.display(window.__xl.grid.sheet, 30, 0).text }));
check('Mở lại app: bản nháp còn thay đổi và vẫn biết tệp chưa được lưu', st.dirty && st.file === 'bieu-do.xlsx' && st.a31 === 'chưa lưu', JSON.stringify(st));
const csp = await win2.evaluate(async () => (await fetch(location.href)).headers.get('content-security-policy') || '');
check('Có chính sách bảo mật nội dung (CSP)', /script-src 'self'/.test(csp));
const closed2 = app2.waitForEvent('close', { timeout: 15000 }).then(() => true).catch(() => false);
await app2.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close());
await win2.waitForSelector('.dlg h2:has-text("Lưu thay đổi")', { timeout: 5000 }).catch(() => {});
await win2.click('.dlg .btn:has-text("Không lưu")').catch(() => {});
check('Mở lại rồi đóng: vẫn hỏi lưu, chọn Không lưu thì thoát', await closed2);

const bad = results.filter((r) => !r.ok);
console.log(`\n${results.length - bad.length}/${results.length} kiểm thử bản máy tính đạt`);
await app.close().catch(() => {}); await app2.close().catch(() => {});
process.exit(bad.length ? 1 : 0);
