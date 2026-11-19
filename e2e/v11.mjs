// Kiểm thử bản 1.1 trên giao diện thật: mở tệp Excel 97–2003 (.xls), Ctrl+S lưu thành .xlsx mới.
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync, readFileSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');
const PORT = 4185, URL0 = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 6000); });
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1440, height: 860 } });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => { errors.push(e.message); console.log('PAGEERR', e.message); });
page.on('dialog', (d) => d.accept());
await page.goto(URL0);
await page.waitForFunction(() => window.__xl && window.__xl.grid);
const X = (f, a) => page.evaluate(f, a);
const wait = (ms = 80) => page.waitForTimeout(ms);
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) pass++; else { fail++; console.log('✗', m); } };

ok(await X(() => document.querySelector('input[type=file]') === null || true), 'trang sẵn sàng');
const b64 = readFileSync(new URL('./fixtures/basic.xls', import.meta.url)).toString('base64');
await X((b64) => {
  const bin = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  window.__xl.openPicked({ file: new File([bin], 'Bao cao cu.xls') });
}, b64);
await page.waitForFunction(() => window.__xl && window.__xl.fileName === 'Bao cao cu.xls', null, { timeout: 15000 });
await wait(700);
ok(await X(() => /97–2003/.test(document.body.innerText)), 'báo tệp .xls sẽ lưu thành .xlsx');
const tabs = await X(() => [...document.querySelectorAll('#sheets [data-sheet]')].map((e) => e.textContent.trim()));
ok(tabs.length === 1 && tabs[0] === 'Dữ liệu', 'trang ẩn không hiện thẻ: ' + tabs.join(','));
ok(await X(() => { const x = window.__xl, s = x.grid.sheet; return x.wb.display(s, 4, 1).text === '2 x 55' && x.wb.display(s, 3, 1).text === '129,953' && x.wb.display(s, 2, 0).text === '2,00'; }), 'giá trị + định dạng số hiện như Excel');
ok(await X(() => { const x = window.__xl, s = x.grid.sheet, st = x.wb.style(s, 0, 0); return st.b && st.fs === 14 && st.bg?.toLowerCase() === '#ffff00'; }), 'định dạng ô tiêu đề');
await page.screenshot({ path: OUT + 'v11-xls-open.png' });
// Ctrl+S: không ghi đè .xls, mở hộp lưu .xlsx
await X(() => {
  window.__saved = [];
  window.showSaveFilePicker = async (o) => {
    window.__pick = { name: o.suggestedName, types: o.types.map((t) => Object.values(t.accept)[0][0]) };
    const name = 'Bao cao cu.xlsx';
    return { name, createWritable: async () => { const parts = []; return { write: async (b) => parts.push(b), close: async () => window.__saved.push({ name, bytes: new Uint8Array(await new Blob(parts).arrayBuffer()) }) }; } };
  };
});
await page.click('.grid-wrap, #grid, canvas', { position: { x: 300, y: 300 } }).catch(() => {});
await page.keyboard.press('Control+s');
await page.waitForFunction(() => window.__saved && window.__saved.length === 1, null, { timeout: 10000 }).catch(() => {});
const pick = await X(() => window.__pick);
ok(pick && pick.name === 'Bao cao cu.xlsx' && pick.types[0] === '.xlsx', 'Ctrl+S đề xuất tên .xlsx: ' + JSON.stringify(pick));
ok(await X(() => window.__saved.length === 1 && window.__saved[0].bytes[0] === 0x50 && window.__saved[0].bytes[1] === 0x4b), 'ghi ra tệp .xlsx (zip)');
ok(await X(() => window.__xl.fileName === 'Bao cao cu.xlsx'), 'sau khi lưu, sổ gắn với tệp .xlsx mới');
// 1.1.2: đổi độ rộng cột → hàng có chữ xuống dòng tự đổi chiều cao để chữ không bị che
await X(() => { const x = window.__xl; x.newBook ? x.newBook() : null; });
await wait(300);
const fit = await X(async () => {
  const x = window.__xl, g = x.grid, s = g.sheet, wb = x.wb;
  wb.setInput(s, 30, 1, 'Chi phí thẩm định hồ sơ yêu cầu và kết quả lựa chọn nhà thầu tư vấn khảo sát và lập BCNCKT');
  wb.setStyle(s, { r1: 30, c1: 1, r2: 30, c2: 1 }, { wr: true });
  wb.setInput(s, 31, 1, 'Dự phòng cho yếu tố khối lượng, công việc phát sinh');
  wb.setStyle(s, { r1: 31, c1: 1, r2: 31, c2: 1 }, { wr: true });
  g.changeColWidths([1], () => wb.setColWidth(s, 1, 400));
  const def = wb.meta(s).rowH.def;
  wb.setRowHeight(s, 31, 60); // hàng đặt cao tay
  const h0 = [wb.meta(s).rowH.get(30), wb.meta(s).rowH.get(31)];
  g.changeColWidths([1], () => wb.setColWidth(s, 1, 120)); // thu hẹp: chữ phải xuống nhiều dòng
  const h1 = [wb.meta(s).rowH.get(30), wb.meta(s).rowH.get(31)];
  const need1 = [g.neededHeight(30), g.neededHeight(31)];
  g.changeColWidths([1], () => wb.setColWidth(s, 1, 400)); // nới rộng lại
  const h2 = [wb.meta(s).rowH.get(30), wb.meta(s).rowH.get(31)];
  wb.undo();
  const h3 = [wb.meta(s).rowH.get(30), wb.meta(s).colW.get(1)];
  return { def, h0, h1, need1, h2, h3 };
});
ok(fit.h1[0] >= fit.need1[0] && fit.h1[0] > fit.h0[0], 'thu hẹp cột: hàng tự động cao thêm vừa chữ ' + JSON.stringify(fit));
ok(fit.h1[1] >= Math.max(60, fit.need1[1]), 'hàng đặt cao tay không bị co, chỉ giãn khi cần');
ok(fit.h2[0] === fit.h0[0], 'nới rộng cột: hàng tự động co lại như cũ');
ok(fit.h2[1] === 60, 'nới rộng cột: hàng đặt cao tay trở về đúng chiều cao đã đặt (60) ' + JSON.stringify(fit.h2));
ok(fit.h3[0] === fit.h1[0] && fit.h3[1] === 120, 'một lần Hoàn tác trả cả độ rộng cột và chiều cao hàng');
console.log(`v11: ${pass} đạt, ${fail} lỗi, ${errors.length} lỗi trang`);
await browser.close();
try { process.kill(-srv.pid); } catch { /* đã tắt */ }
process.exit(fail || errors.length ? 1 : 0);
