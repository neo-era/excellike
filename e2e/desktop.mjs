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
const printOut = join(work, 'print.json'), pdfOut = join(work, 'xuat.pdf');
import { testKey, TEST_GUID } from './testkey.mjs';
mkdirSync(userData, { recursive: true });
writeFileSync(join(userData, 'license.key'), testKey());
const env = { ...process.env, XL_TEST_MACHINE: TEST_GUID, XL_USER_DATA: userData, XL_TEST_OPEN: csv, XL_TEST_SAVE: saveAs, XL_TEST_PRINT: printOut, XL_TEST_PDF: pdfOut };

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

// 0.5: thanh tiêu đề tự vẽ như Excel (nút cửa sổ hệ thống nằm đè), màu theo chủ đề Office
const wco = await X(() => ({ on: !!navigator.windowControlsOverlay?.visible, drag: getComputedStyle(document.querySelector('#titlebar')).getPropertyValue('-webkit-app-region') || getComputedStyle(document.querySelector('#titlebar')).getPropertyValue('app-region'), desk: document.documentElement.classList.contains('desktop') }));
check('Cửa sổ không viền: thanh tiêu đề của app (kéo được) + nút cửa sổ hệ thống đè lên', wco.on && wco.drag === 'drag' && wco.desk, JSON.stringify(wco));
await X(() => window.__xl.setTheme('black'));
await win.waitForTimeout(300);
const ov = await app.evaluate(() => global.__xlOverlay?.());
check('Đổi chủ đề Đen → nút cửa sổ đổi màu theo', ov?.color === '#0a0a0a' && ov?.symbolColor === '#e6e6e6', JSON.stringify(ov));
await win.screenshot({ path: OUT + 'v05-desktop-black.png' });
await X(() => window.__xl.setTheme('colorful'));
await win.waitForTimeout(200);
check('Nhiều màu → thanh xanh #107C41', (await app.evaluate(() => global.__xlOverlay?.()))?.color === '#107c41');
await win.screenshot({ path: OUT + 'v05-desktop.png' });
// In: danh sách máy in của Windows, in thẳng không qua hộp thoại, xuất PDF thật
const printers = await X(() => window.xlDesktop.printers());
check('Lấy danh sách máy in của hệ thống', Array.isArray(printers), JSON.stringify(printers).slice(0, 120));
await win.keyboard.press('Control+p');
await win.waitForSelector('#backstage .bs-preview iframe');
await win.waitForTimeout(500);
const opts = await X(() => [...document.querySelectorAll('#bs-printer option')].map((o) => o.value));
check('Máy in có mục “Lưu thành PDF”', opts.includes('__pdf'), opts.join(','));
await win.fill('#bs-copies', '2');
await win.click('#bs-print');
await win.waitForFunction(() => document.querySelector('.toast')?.textContent.includes('Đã gửi'), null, { timeout: 15000 }).catch(() => {});
const pj = existsSync(printOut) ? JSON.parse(execFileSync('cat', [printOut]).toString()) : null;
check('In thẳng ra máy in (không hộp thoại), 2 bản, khổ giấy theo thiết lập', pj && pj.silent === true && pj.copies === 2 && pj.pages >= 1 && pj.pageSize === 'A4', JSON.stringify(pj));
await win.selectOption('#bs-printer', '__pdf');
await win.click('#bs-print');
await win.waitForFunction(() => document.querySelector('.toast')?.textContent.includes('PDF'), null, { timeout: 20000 }).catch(() => {});
await win.waitForTimeout(300);
const pdfOk = existsSync(pdfOut) ? execFileSync('pdftotext', ['-layout', pdfOut, '-']).toString() : '';
const pdfInfo = existsSync(pdfOut) ? execFileSync('pdfinfo', [pdfOut]).toString() : '';
check('Xuất PDF thật (có chữ, đúng khổ A4)', /Cáp đồng|ĐÃ SỬA/.test(pdfOk) && /Page size:\s+59[45]/.test(pdfInfo), (pdfInfo.match(/Page.*/g) || []).join(' | '));
await win.keyboard.press('Escape');

// liên kết ngoài mở bằng trình duyệt, không mở cửa sổ mới trong app
const nWin = app.windows().length;
await X(() => window.open('https://example.com', '_blank'));
await win.waitForTimeout(400);
check('Liên kết ngoài không mở cửa sổ mới trong app', app.windows().length === nWin);

// 0.6: Ctrl+F10 phóng to/khôi phục, Ctrl+F9 thu nhỏ; liên kết trong ô mở bằng ứng dụng mặc định
await X(() => window.__xl.grid.focusKeys());
// (xvfb không có trình quản lý cửa sổ nên đếm lời gọi maximize/unmaximize thay vì đọc trạng thái)
await app.evaluate(({ BrowserWindow }) => { const w = BrowserWindow.getAllWindows()[0]; global.__mx = []; const mx = w.maximize.bind(w), ux = w.unmaximize.bind(w); w.maximize = () => { global.__mx.push('max'); mx(); }; w.unmaximize = () => { global.__mx.push('unmax'); ux(); }; });
await win.keyboard.press('Control+F10'); await win.waitForTimeout(300);
const mx = await app.evaluate(() => global.__mx);
check('Ctrl+F10: phóng to / khôi phục cửa sổ', mx.length === 1, JSON.stringify(mx));
await app.evaluate(({ BrowserWindow }) => { const w = BrowserWindow.getAllWindows()[0]; global.__minCalls = 0; const m = w.minimize.bind(w); w.minimize = () => { global.__minCalls++; m(); }; });
await win.keyboard.press('Control+F9'); await win.waitForTimeout(300);
const mins = await app.evaluate(({ BrowserWindow }) => { const w = BrowserWindow.getAllWindows()[0]; w.restore(); w.show(); return global.__minCalls; });
check('Ctrl+F9: thu nhỏ cửa sổ', mins === 1, 'gọi minimize ' + mins + ' lần');
await app.evaluate(({ shell }) => { global.__ext = []; shell.openExternal = (u) => { global.__ext.push(u); return Promise.resolve(); }; });
await X(() => { const a = window.__xl; a.grid.select(40, 0); a.setLink('https://sapulico.vn', 'Sapulico'); a.followLink(40, 0); });
await win.waitForTimeout(300);
const ext = await app.evaluate(() => global.__ext);
check('Nhấp liên kết trong ô (Ctrl+K) mở bằng trình duyệt mặc định', ext.includes('https://sapulico.vn'), JSON.stringify(ext));

// chạy lần hai với tệp khác → mở trong cửa sổ đang chạy (một phiên duy nhất)
const fx2 = join(work, 'bieu-do.xlsx');
copyFileSync(join(ROOT, 'tests/fixtures/bieudo.xlsx'), fx2);
const second = spawn(electronBin, [...(process.getuid?.() === 0 ? ['--no-sandbox'] : []), ROOT, fx2], { env, stdio: 'ignore' });
const secondExit = await new Promise((r) => { second.on('exit', (code) => r(code)); setTimeout(() => r('timeout'), 15000); });
await win.waitForFunction(() => window.__xl && window.__xl.fileName === 'bieu-do.xlsx', null, { timeout: 15000 }).catch(() => {});
const charts = await X(() => (window.__xl ? window.__xl.wb.sheetIds().reduce((a, i) => a + window.__xl.wb.meta(i).charts.length, 0) : -1));
check('Mở tệp thứ hai: chuyển vào cửa sổ đang chạy, có biểu đồ', secondExit === 0 && charts > 0, `exit=${secondExit} charts=${charts}`);
check('Chỉ một cửa sổ', app.windows().length === 1);

// lưu hai lần cùng lúc (giữ Ctrl+S) → tệp vẫn hợp lệ, không báo lỗi giả
await X(() => { const x = window.__xl; x.wb.setInput(x.grid.sheet, 31, 0, 'lần 2'); });
const both = await X(() => Promise.all([window.__xl.saveXlsx(false), window.__xl.saveXlsx(false), window.__xl.saveXlsx(false)]));
const okFile = py('import openpyxl,sys;print(openpyxl.load_workbook(sys.argv[1]).worksheets[0]["A32"].value)', join(work, 'bieu-do.xlsx'));
check('Lưu dồn dập: chỉ ghi một lần, tệp hợp lệ', both.every(Boolean) && okFile === 'lần 2', `${JSON.stringify(both)} ${okFile}`);
// Tự động lưu (công tắc trên thanh tiêu đề): sửa ô → tự ghi vào tệp đang mở
await win.click('#autosave');
check('Bật Tự động lưu khi đang mở tệp .xlsx', (await win.getAttribute('#autosave', 'aria-checked')) === 'true');
await X(() => { const x = window.__xl; x.wb.setInput(x.grid.sheet, 32, 0, 'tự lưu'); });
await win.waitForTimeout(4500);
const auto = py('import openpyxl,sys;print(openpyxl.load_workbook(sys.argv[1]).worksheets[0]["A33"].value)', join(work, 'bieu-do.xlsx'));
check('Tự động lưu ghi thay đổi vào tệp sau vài giây', auto === 'tự lưu', auto);
await win.click('#autosave');
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

// nhấp đúp tệp lớn rồi nhấp đúp lần nữa khi app đang khởi động/đang mở → vẫn ra đúng sổ, không quay về sổ trống
{
  const big = join(work, 'tep-lon.xlsx');
  execFileSync('python3', ['-c', `
import openpyxl,sys
wb=openpyxl.Workbook(); ws=wb.active; ws.title='Lớn'
for r in range(1,60001): ws.append([f'Mục {r}', r, r*2, f'=B{r}*C{r}', f'Ghi chú dài cho hạng mục số {r}', r*0.5])
wb.save(sys.argv[1])`, big]);
  const ud = join(work, 'ud-race');
  const envR = { ...env, XL_USER_DATA: ud };
  const a3 = await electron.launch({ executablePath: electronBin, args: [ROOT, big], env: envR });
  const w3 = await a3.firstWindow();
  await w3.waitForTimeout(300);
  const again = spawn(electronBin, [...(process.getuid?.() === 0 ? ['--no-sandbox'] : []), ROOT, big], { env: envR, stdio: 'ignore' });
  await new Promise((r) => { again.on('exit', r); setTimeout(r, 15000); });
  await w3.waitForFunction(() => window.__xl && window.__xl.wb.name === 'tep-lon', null, { timeout: 60000 }).catch(() => {});
  await w3.waitForTimeout(1500);
  const r3 = await w3.evaluate(() => ({ name: window.__xl.wb.name, rows: window.__xl.wb.used(window.__xl.grid.sheet).rows }));
  check('Mở tệp lớn khi nhấp đúp hai lần liền: vẫn ra đúng sổ', r3.name === 'tep-lon' && r3.rows === 60000, JSON.stringify(r3));
  await a3.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close());
  await a3.waitForEvent('close', { timeout: 15000 }).catch(() => {});
  const a4 = await electron.launch({ executablePath: electronBin, args: [ROOT], env: envR });
  const w4 = await a4.firstWindow();
  await w4.waitForFunction(() => window.__xl, null, { timeout: 30000 });
  const n4 = await w4.evaluate(() => window.__xl.wb.name);
  check('Mở lại app: vẫn là sổ vừa mở (không quay về Sổ tính 1)', n4 === 'tep-lon', n4);
  await a4.close().catch(() => {});
}

const bad = results.filter((r) => !r.ok);
console.log(`\n${results.length - bad.length}/${results.length} kiểm thử bản máy tính đạt`);
await app.close().catch(() => {}); await app2.close().catch(() => {});
process.exit(bad.length ? 1 : 0);
