// Kiểm thử bản 1.0 trên giao diện thật: giữ macro .xlsm (nhóm A).
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync, readFileSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');
const PORT = 4184, URL0 = `http://localhost:${PORT}/`;
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

// ================= A. Giữ macro .xlsm =================
const b64 = readFileSync(new URL('./fixtures/macro.xlsm', import.meta.url)).toString('base64');
await X((b64) => {
  const bin = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  window.__xl.openPicked({ file: new File([bin], 'Macro.xlsm') });
}, b64);
await page.waitForFunction(() => window.__xl && window.__xl.wb && window.__xl.wb.hasMacros && window.__xl.wb.hasMacros(), null, { timeout: 15000 });
await wait(700);
ok(await X(() => window.__xl.fileName === 'Macro.xlsm'), 'tên tệp .xlsm');
ok(await X(() => /macro/i.test(document.body.innerText)), 'báo sổ có macro khi mở');
const tabs = await X(() => [...document.querySelectorAll('#sheets [data-sheet]')].map((e) => e.textContent.trim()));
ok(!tabs.some((t) => t.includes('Ẩn sâu')) && tabs.length === 2, 'trang veryHidden không hiện thẻ: ' + tabs.join(','));
ok(await X(() => { const x = window.__xl; return x.wb.value(x.wb.sheetIds()[0], 0, 2) === 'a, b'; }), 'ô hàm VBA hiện giá trị lưu sẵn');
ok(await X(() => window.__xl.wb.unhideableSheetIds().length === 0), 'Bỏ ẩn không liệt kê trang veryHidden');
// menu chuột phải thẻ trang: Bỏ ẩn… bị mờ
await page.click('#sheets [data-sheet]', { button: 'right' }); await wait(150);
ok(await X(() => { const it = [...document.querySelectorAll('button.mi')].find((e) => e.textContent.includes('Bỏ ẩn…')); return !!it && it.disabled; }), 'menu Bỏ ẩn… bị mờ');
await page.keyboard.press('Escape'); await wait(80);
// xóa trang có codeName → hỏi lại
await X(() => { const x = window.__xl; x.removeSheet(x.wb.visibleSheetIds()[1]); });
await wait(150);
ok(await X(() => /Macro VBA trong sổ có thể tham chiếu/.test(document.querySelector('.dlg')?.textContent ?? '')), 'cảnh báo xóa trang có macro');
await page.screenshot({ path: OUT + 'v10-del-macro-sheet.png' });
await page.click('.dlg-foot .btn:not(.primary)'); await wait(80);
ok(await X(() => window.__xl.wb.visibleSheetIds().length === 2), 'Hủy thì không xóa');

// Lưu: hộp lưu giả trả tên .xlsm → giữ macro
await X(() => {
  window.__saved = [];
  window.__pickName = 'Ban sao.xlsm';
  window.showSaveFilePicker = async (o) => {
    window.__pickTypes = o.types.map((t) => Object.values(t.accept)[0][0]);
    const name = window.__pickName;
    return { name, createWritable: async () => { const parts = []; return { write: async (b) => parts.push(b), close: async () => window.__saved.push({ name, bytes: new Uint8Array(await new Blob(parts).arrayBuffer()) }) }; } };
  };
});
await X(() => window.__xl.saveXlsx(true));
await wait(300);
const has = (name) => X(async (name) => { const s = window.__saved.at(-1); const txt = new TextDecoder('latin1').decode(s.bytes); return { name: s.name, vba: txt.includes(name) }; }, name);
let r = await has('xl/vbaProject.bin');
ok(r.name === 'Ban sao.xlsm' && r.vba, 'lưu .xlsm giữ vbaProject.bin');
ok(await X(() => window.__pickTypes.join() === '.xlsm,.xlsx'), 'hộp lưu cho chọn .xlsm hoặc .xlsx');
// Lưu thành .xlsx → hỏi bỏ macro
await X(() => { window.__pickName = 'Khong macro.xlsx'; window.__xl.saveXlsx(true); });
await page.waitForSelector('.dlg', { timeout: 5000 });
ok(await X(() => /bỏ toàn bộ macro/.test(document.querySelector('.dlg').textContent)), 'báo macro sẽ bị bỏ khi lưu .xlsx');
await page.screenshot({ path: OUT + 'v10-save-plain.png' });
await page.click('.dlg-foot .btn.primary'); await wait(500);
r = await has('xl/vbaProject.bin');
ok(r.name === 'Khong macro.xlsx' && !r.vba, 'tệp .xlsx không còn macro');
ok(await X(() => !window.__xl.wb.hasMacros() && window.__xl.fileName === 'Khong macro.xlsx'), 'sổ đang mở không còn macro sau khi lưu .xlsx');
// Lưu thành .xlsx rồi bấm Hủy → không ghi
const n0 = await X(() => window.__saved.length);
await X(() => { window.__xl.wb.vba = { files: [['xl/vbaProject.bin', 'AAAA']], ctDefaultBin: true, overrides: [] }; window.__pickName = 'Huy.xlsx'; window.__xl.saveXlsx(true); });
await page.waitForSelector('.dlg', { timeout: 5000 });
await page.click('.dlg-foot .btn:not(.primary)'); await wait(300);
ok(await X((n0) => window.__saved.length === n0, n0), 'Hủy thì không ghi tệp');

// nút ribbon có thể nằm trong nhóm đã co (cửa sổ hẹp): mở nhóm co rồi bấm
async function rb(id, sub = '') {
  if (await page.locator(`#ribbon [data-id="${id}"]`).count()) { await page.click(`#ribbon [data-id="${id}"]${sub}`); return; }
  const n = await page.locator('#ribbon .grp.collapsed:not([hidden]) > .big').count();
  for (let i = 0; i < n; i++) {
    await page.locator('#ribbon .grp.collapsed:not([hidden]) > .big').nth(i).click(); await wait(80);
    if (await page.locator(`.grp-pop [data-id="${id}"]`).count()) { await page.click(`.grp-pop [data-id="${id}"]${sub}`); return; }
    await page.keyboard.press('Escape'); await wait(60);
  }
  throw new Error('không thấy nút ' + id);
}
// ================= B. Tính toán Thủ công =================
const K = (k) => page.keyboard.press(k);
const A = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return [+m[2] - 1, [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1]; };
const val = (a) => X(([r, c]) => window.__xl.wb.value(window.__xl.grid.sheet, r, c), A(a));
async function goto(a) { await page.click('#namebox'); await page.keyboard.type(a); await K('Enter'); await wait(40); }
const typeAt = async (a, t) => { await goto(a); await page.keyboard.type(t); await K('Enter'); await wait(60); };
const calcBadge = () => X(() => { const b = document.getElementById('st-calc'); return !!b && !b.hidden && b.offsetParent !== null; });
await X(() => { const x = window.__xl, s = x.grid.sheet; x.wb.setInput(s, 0, 0, '1'); x.wb.setInput(s, 0, 1, '=A1*2'); x.wb.setInput(s, 0, 2, '=SUM(A1:A3)'); x.grid.invalidate(); });
await page.click('#tabs [data-tab="formulas"]'); await wait(100);
await rb('calc-opt'); await wait(150);
ok(await X(() => !![...document.querySelectorAll('button.mi')].find((e) => e.textContent.includes('Thủ công'))), 'menu Tùy chọn Tính toán có Thủ công');
await page.click('button.mi:has-text("Thủ công")'); await wait(150);
ok(await X(() => window.__xl.wb.calcMode === 'manual'), 'chuyển sang Thủ công từ ribbon');
await X(() => window.__xl.grid.focusKeys());
await typeAt('A1', '5');
ok((await val('A1')) === 5 && (await val('B1')) === 2, 'Thủ công: sửa ô không tính lại ô phụ thuộc');
ok(await calcBadge(), 'thanh trạng thái hiện “Tính toán”');
await page.screenshot({ path: OUT + 'v10-manual-status.png' });
await typeAt('D1', '=A1+1');
ok((await val('D1')) === 6, 'công thức vừa gõ tự tính ngay');
await page.click('#st-calc'); await wait(120);
ok((await val('B1')) === 10 && !(await calcBadge()), 'bấm “Tính toán” thì tính, chữ tắt đi');
await typeAt('A2', '4');
ok((await val('C1')) === 5, 'tổng chưa đổi trước F9');
await K('F9'); await wait(120);
ok((await val('C1')) === 9, 'F9 tính lại');
ok(await X(() => window.__xl.wb.canUndo()), 'F9 giữ lịch sử Hoàn tác');
await K('Control+z'); await wait(80);
ok((await val('A2')) === null && (await val('C1')) === 9, 'Hoàn tác ở Thủ công: ô về cũ, ô phụ thuộc chờ F9');
await K('Shift+F9'); await wait(100);
ok((await val('C1')) === 5, 'Shift+F9 tính lại');
// thao tác trong danh sách tính ngay: chèn hàng
await typeAt('A1', '7');
await X(() => { const x = window.__xl; x.wb.insertRows(x.grid.sheet, 9, 1); x.grid.invalidate(); });
ok((await val('B1')) === 14, 'chèn hàng: tính lại ngay dù Thủ công');
// lưu: “Tính lại sổ trước khi lưu” (mặc định bật)
await typeAt('A1', '8');
await X(() => { window.__pickName = 'Thu cong.xlsm'; });
await X(() => window.__xl.saveXlsx(true)); await wait(400);
ok((await val('B1')) === 16 && !(await calcBadge()), 'lưu khi bật “Tính lại sổ trước khi lưu” thì tính trước');
await X(() => { window.__xl.wb.calcOnSave = false; });
await typeAt('A1', '9');
await X(() => window.__xl.saveXlsx(true)); await wait(400);
ok((await val('B1')) === 16 && (await calcBadge()), 'tắt “Tính lại sổ trước khi lưu” thì lưu không tính');
// Ctrl+Alt+F9
await K('Control+Alt+F9'); await wait(120);
ok((await val('B1')) === 18, 'Ctrl+Alt+F9 tính tất cả');
// về Tự động
await page.click('#tabs [data-tab="formulas"]'); await wait(80);
await rb('calc-opt'); await wait(150);
await page.click('button.mi:has-text("Tự động")'); await wait(150);
await typeAt('A1', '2');
ok((await val('B1')) === 4 && !(await calcBadge()), 'về Tự động: sửa ô tính ngay');

// ================= C. LET/LAMBDA, màu, vùng in =================
await typeAt('F1', '=LET(x;A1*10;y;x+1;x*y)');
ok((await val('F1')) === 420, 'gõ LET trên lưới: ' + (await val('F1')));
await goto('F1');
ok(await X(() => document.querySelector('#fx-input, .fx-input, #formula')?.value ?? window.__xl.wb.rawText(window.__xl.grid.sheet, 0, 5)) === '=LET(x;A1*10;y;x+1;x*y)', 'thanh công thức hiện đúng LET');
await typeAt('F2', '=LAMBDA(n;n*n)(7)');
ok((await val('F2')) === 49, 'gọi ngay LAMBDA');
ok(await X(() => window.__xl.wb.rawText(window.__xl.grid.sheet, 1, 5)) === '=LAMBDA(n;n*n)(7)', 'thanh công thức hiện LAMBDA(…)(…) không lộ dạng nội bộ');
await typeAt('F3', '=LAMBDA(n;n)');
ok(await X(() => window.__xl.wb.display(window.__xl.grid.sheet, 2, 5).text) === '#CALC!', 'LAMBDA không gọi hiện #CALC!');
await typeAt('G1', '=MAP(SEQUENCE(3);LAMBDA(v;v*2))');
ok((await val('G3')) === 6, 'MAP tràn 3 ô');
ok(await X(() => !window.__xl.wb.listNames().some((n) => /^(x|y|n|v)$/i.test(n.name))), 'tên LET/LAMBDA không vào Trình Quản lý Tên');
// bảng màu: Màu Chủ đề theo sổ + Màu Gần đây
await X(() => { window.__xl.wb.theme = ['FFFFFF', '000000', 'E7E6E6', '44546A', 'C00000', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']; window.__xl.wb.addRecentColor('#123456'); });
await page.click('#tabs [data-tab="home"]'); await wait(80);
await page.click('#ribbon [data-id="bg"] .dd'); await wait(150);
const pal = await X(() => !!document.querySelector('.palette'));
ok(pal, 'mở bảng Màu Tô');
if (pal) {
  ok(await X(() => !!document.querySelector('.palette [data-c="#C00000"]')), 'Màu Chủ đề lấy theo sổ (Accent1 #C00000)');
  ok(await X(() => /Màu Gần đây/.test(document.querySelector('.palette').textContent) && !!document.querySelector('.palette [data-c="#123456"]')), 'có Màu Gần đây');
  await page.screenshot({ path: OUT + 'v10-palette.png' });
  await K('Escape');
}
// vùng in: chọn 2 khối rồi Đặt Vùng In
await X(() => { const x = window.__xl, s = x.grid.sheet; x.wb.setPrint(s, { ...(x.wb.meta(s).print ?? {}), area: { r1: 0, c1: 0, r2: 2, c2: 1 }, more: [{ r1: 5, c1: 4, r2: 7, c2: 6 }] }); });
await page.click('#views [data-view="break"]'); await wait(400);
await page.screenshot({ path: OUT + 'v10-printareas.png' });
await page.click('#views [data-view="normal"]'); await wait(200);
ok(await X(() => { const x = window.__xl, p = x.wb.meta(x.grid.sheet).print; return p.more?.length === 1; }), 'vùng in 2 khối');

console.log(`v10: ${pass} đạt, ${fail} lỗi, ${errors.length} lỗi trang`);
await browser.close();
try { process.kill(-srv.pid); } catch { /* đã tắt */ }
process.exit(fail || errors.length ? 1 : 0);
