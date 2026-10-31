// Kiểm thử bản 0.8–0.9 trên giao diện thật.
// Văn bản thành Cột, Xóa trùng lặp, Bảng, mảng động.
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');
const PORT = 4183, URL0 = `http://localhost:${PORT}/`;
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
const K = (k) => page.keyboard.press(k);
const wait = (ms = 80) => page.waitForTimeout(ms);
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) pass++; else { fail++; console.log('✗', m); } };
const S = () => X(() => window.__xl.grid.sheet);
const A = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return [+m[2] - 1, [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1]; };
const val = (a) => X(([r, c]) => window.__xl.wb.value(window.__xl.grid.sheet, r, c), A(a));
const raw = (a) => X(([r, c]) => window.__xl.wb.rawText(window.__xl.grid.sheet, r, c), A(a));
const put = (cells) => X((cells) => { const x = window.__xl; for (const [a, v] of cells) { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; x.wb.setInput(x.grid.sheet, +m[2] - 1, c, v); } x.grid.invalidate(); }, cells);
async function goto(a) { await page.click('#namebox'); await page.keyboard.type(a); await K('Enter'); await wait(40); }
const focusGrid = () => X(() => window.__xl.grid.focusKeys());
const tab = async (id) => { await page.click(`#tabs [data-tab="${id}"]`); await wait(); };

// nút ribbon có thể nằm trong nhóm đã co (cửa sổ hẹp): mở nhóm co rồi bấm
async function rb(id) {
  if (await page.locator(`#ribbon [data-id="${id}"]`).count()) { await page.click(`#ribbon [data-id="${id}"]`); return; }
  const n = await page.locator('#ribbon .grp.collapsed:not([hidden]) > .big').count();
  for (let i = 0; i < n; i++) {
    await page.locator('#ribbon .grp.collapsed:not([hidden]) > .big').nth(i).click(); await wait(80);
    if (await page.locator(`.grp-pop [data-id="${id}"]`).count()) { await page.click(`.grp-pop [data-id="${id}"]`); return; }
    await page.keyboard.press('Escape'); await wait(60);
  }
  // nhóm bị ẩn hẳn vì thiếu chỗ: mở rộng cửa sổ tạm thời
  await page.setViewportSize({ width: 2400, height: 860 }); await wait(250);
  const found = await page.locator(`#ribbon [data-id="${id}"]`).count();
  if (found) await page.click(`#ribbon [data-id="${id}"]`);
  await page.setViewportSize({ width: 1440, height: 860 }); await wait(150);
  if (!found) throw new Error('không thấy nút ' + id);
}
// ================= 0.8 Bố trí Trang =================
await X(() => { const x = window.__xl, s = x.grid.sheet; for (let r = 0; r < 120; r++) { x.wb.setInput(s, r, 0, 'Hạng mục ' + (r + 1)); x.wb.setInput(s, r, 1, String(r * 10)); } x.grid.invalidate(); });
await page.click('#views [data-view="layout"]'); await wait(400);
const plPages = () => X(() => document.querySelector('.pl-view')?.shadowRoot?.querySelectorAll('section.page').length ?? 0);
ok(await X(() => !document.querySelector('.pl-view').hidden), 'mở Bố trí Trang');
ok((await plPages()) >= 2, 'hiện các trang giấy');
ok(await X(() => !!document.querySelector('.pl-view').shadowRoot.querySelector('.ruler-h')), 'có thước đo');
await page.screenshot({ path: OUT + 'v08-layout.png' });
// sửa đầu trang ngay trên trang
await X(() => { const d = document.querySelector('.pl-view').shadowRoot.querySelector('.hfz div.c'); d.focus(); });
await page.keyboard.type('Dự toán &[Trang]'); await K('Enter'); await wait(300);
ok((await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).print?.header)) === '&CDự toán &P', 'sửa đầu trang tại chỗ');
{ const t = await X(() => [...document.querySelector('.pl-view').shadowRoot.querySelectorAll('.hfz div.c')].map((d) => d.textContent).join('|')); ok(t.split('|')[2] === 'Dự toán &[Trang]', 'trang 2 hiện đầu trang mới: ' + t); }
// nhấp đúp ô → về Thường
await X(() => { const td = document.querySelector('.pl-view').shadowRoot.querySelector('td[data-rc="4,0"]'); td.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true })); });
await wait(150);
ok((await X(() => document.querySelector('.pl-view').hidden)) && (await page.inputValue('#namebox')) === 'A5', 'nhấp đúp ô → dạng Thường tại A5');
// trang đầu khác qua Thiết lập Trang
await X(() => window.__xl.pageSetup('hf')); await wait(200);
await page.click('#ps-difffirst'); await wait(60);
await page.fill('#ps-hc-1', 'BÌA HỒ SƠ');
await page.click('.dlg-foot .primary'); await wait(150);
ok((await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).print?.hfFirst?.header)) === '&CBÌA HỒ SƠ', 'Trang đầu khác');
await page.click('#views [data-view="layout"]'); await wait(400);
{ const t = (await X(() => [...document.querySelector('.pl-view').shadowRoot.querySelectorAll('.hfz div.c')].map((d) => d.textContent).join('|'))).split('|'); ok(t[0] === 'BÌA HỒ SƠ' && t[2] === 'Dự toán &[Trang]', 'trang 1 dùng đầu trang riêng: ' + t.slice(0, 3)); }
await page.screenshot({ path: OUT + 'v08-layout2.png' });
// kéo đường ngắt trang
await page.click('#views [data-view="break"]'); await wait(200);
await X(() => { const g = window.__xl.grid; g.scroller.scrollTop = 20 * 20 * g.zoom; }); await wait(150);
const line = await X(() => { const g = window.__xl.grid; const L = g.breakLayout(); const p = L.pages.find((q) => q.r1 !== L.area.r1); const b = g.canvas.getBoundingClientRect(); return { y: b.top + g.sy(p.r1), x: b.left + g.sx(0) + 30, r: p.r1 }; });
await page.mouse.move(line.x, line.y); await page.mouse.down();
await page.mouse.move(line.x, line.y - 100, { steps: 6 }); await page.mouse.up(); await wait(150);
const br = await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).breaks.rows);
ok(br.length === 1 && br[0] < line.r && br[0] > line.r - 8, 'kéo ngắt trang lên → ngắt thủ công ' + JSON.stringify(br) + ' từ ' + line.r);
await page.screenshot({ path: OUT + 'v08-break.png' });
await page.click('#views [data-view="normal"]'); await wait(100);

// ================= 0.9 Kiểm nghiệm công thức =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await put([['A1', '2'], ['A2', '3'], ['A3', '1,5'], ['B1', '=ROUND(SUM(A1:A2)*A3;1)+1'], ['C1', "'123"], ['D1', '=1/0']]);
await goto('B1'); await tab('formulas');
await page.click('#ribbon [data-id="eval-f"]'); await wait(150);
ok((await X(() => document.querySelector('.ev-box.cur u')?.textContent)) === 'A1:A2', 'Đánh giá: gạch chân A1:A2');
for (let i = 0; i < 4; i++) { await page.click('.ev-btns .primary'); await wait(40); }
ok((await X(() => document.querySelector('.ev-box.cur').textContent)) === '=7,5+1', 'sau 4 bước còn =7,5+1');
await page.click('.ev-btns .primary'); await wait(40);
ok((await X(() => document.querySelector('.ev-box.cur').textContent)) === '=8,5', 'kết quả 8,5');
await page.keyboard.press('Escape'); await wait(80);
// Bước Vào
await goto('E1'); await page.keyboard.type('=B1*2'); await K('Enter'); await goto('E1');
await page.click('#ribbon [data-id="eval-f"]'); await wait(150);
await page.click('.ev-btns button:has-text("Bước Vào")'); await wait(60);
ok((await X(() => document.querySelectorAll('.ev-box').length)) === 2, 'Bước Vào mở công thức ô B1');
await page.click('.ev-btns button:has-text("Bước Ra")'); await wait(60);
ok((await X(() => document.querySelector('.ev-box.cur').textContent)) === '=8,5*2', 'Bước Ra thay B1 bằng 8,5');
await page.keyboard.press('Escape'); await wait(80);
// tam giác xanh + nút ! + chuyển thành số
ok(await X(() => window.__xl.wb.cellIssue(window.__xl.grid.sheet, 0, 2) === 'numText' && window.__xl.wb.cellIssue(window.__xl.grid.sheet, 0, 3) === 'err'), 'nhận ra số dạng văn bản và lỗi công thức');
await goto('C1'); await focusGrid(); await K('Alt+Shift+F10'); await wait(120);
await page.click('.menu [role=menuitem]:has-text("Chuyển thành Số")'); await wait(100);
ok((await val('C1')) === 123, 'Chuyển thành Số');
await goto('D1'); await wait(60);
const eb = await X(() => { const g = window.__xl.grid, b = g.canvas.getBoundingClientRect(), e = g.errBtn(); return e && [b.left + e.x + e.w / 2, b.top + e.y + e.h / 2]; });
ok(!!eb, 'hiện nút ! cạnh ô lỗi');
await page.mouse.click(eb[0], eb[1]); await wait(120);
ok((await X(() => [...document.querySelectorAll('.menu [role=menuitem]')].map((e) => e.textContent).join('|'))).includes('Lỗi Chia cho Không'), 'menu nêu loại lỗi');
await page.click('.menu [role=menuitem]:has-text("Bỏ qua Lỗi")'); await wait(100);
ok(!(await X(() => window.__xl.wb.cellIssue(window.__xl.grid.sheet, 0, 3))), 'Bỏ qua Lỗi');
await page.screenshot({ path: OUT + 'v09-errors.png' });
// Kiểm tra Lỗi… (hộp duyệt)
await put([['C2', "'45"], ['D2', '=NA()']]);
await tab('formulas'); await page.click('#ribbon [data-id="err-chk"] .main, #ribbon [data-id="err-chk"]'); await wait(150);
ok((await X(() => document.querySelector('.ec-where')?.textContent ?? '')).includes('/2'), 'Kiểm tra Lỗi tìm 2 ô');
await page.keyboard.press('Escape'); await wait(80);
// Cửa sổ Theo dõi
await goto('B1'); await tab('formulas'); await page.click('#ribbon [data-id="watch"]'); await wait(100);
await page.click('.watch-pane button:has-text("Thêm Theo dõi")'); await wait(100);
ok((await X(() => document.querySelector('.watch-pane tbody td:nth-child(4)')?.textContent)) === '8,5', 'theo dõi B1 = 8,5');
await put([['A3', '2']]); await wait(80);
ok((await X(() => document.querySelector('.watch-pane tbody td:nth-child(4)')?.textContent)) === '11', 'giá trị theo dõi tự cập nhật');
await page.screenshot({ path: OUT + 'v09-watch.png' });
await page.click('.watch-pane .x'); await wait(60);
ok(await X(() => document.querySelector('.watch-pane').hidden), 'đóng Cửa sổ Theo dõi');

// ================= 0.9 Phân tích =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await put([['A1', '10'], ['B1', '25000'], ['C1', '=A1*B1*1,1']]);
await goto('C1'); await tab('data');
await page.click('#ribbon [data-id="d-whatif"]'); await wait(100);
await page.click('.menu [role=menuitem]:has-text("Tìm Mục tiêu")'); await wait(150);
await page.locator('.dlg .an-form input').nth(1).fill('1000000');
await page.locator('.dlg .an-form input').nth(2).fill('A1');
await page.click('.dlg-foot .primary'); await wait(250);
ok((await X(() => document.querySelector('.dlg h2')?.textContent)) === 'Trạng thái Tìm Mục tiêu', 'Tìm Mục tiêu báo trạng thái');
await page.click('.dlg-foot .primary'); await wait(100);
ok(Math.abs((await val('C1')) - 1000000) < 0.01, 'Tìm Mục tiêu đạt 1.000.000: ' + (await val('C1')));
// Dự báo
await put([['E1', 'Năm'], ['F1', 'Sản lượng'], ['E2', '2020'], ['F2', '100'], ['E3', '2021'], ['F3', '120'], ['E4', '2022'], ['F4', '138'], ['E5', '2023'], ['F5', '161']]);
await goto('E2'); await tab('data');
await page.click('#ribbon [data-id="d-fc"]'); await wait(150);
await page.click('.dlg-foot .primary'); await wait(300);
ok((await X(() => window.__xl.wb.sheetName(window.__xl.grid.sheet))) === 'Dự báo', 'tạo trang Dự báo');
ok(Math.abs((await val('C6')) - 179.5) < 0.6 && (await val('D6')) < (await val('C6')) && (await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).charts.length)) === 1, 'dự báo 2024 ≈ 179,5 + khoảng tin cậy + biểu đồ: ' + (await val('C6')));
await page.screenshot({ path: OUT + 'v09-forecast.png' });

// ================= 0.9 Biểu đồ =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await put([['A1', 'Tháng'], ['B1', 'Doanh thu'], ['C1', 'Tỷ lệ'], ['A2', '1'], ['B2', '100'], ['C2', '0,1'], ['A3', '2'], ['B3', '150'], ['C3', '0,2'], ['A4', '3'], ['B4', '120'], ['C4', '0,15']]);
await X(() => window.__xl.grid.selectRange({ r1: 0, c1: 0, r2: 3, c2: 1 }));
await tab('insert'); await page.click('#ribbon [data-id="i-xy"]'); await wait(200);
ok((await X(() => document.querySelector('.chart-types [data-t="scatter"]')?.getAttribute('aria-checked'))) === 'true', 'hộp biểu đồ chọn sẵn Phân tán XY');
await page.click('.dlg-foot .primary'); await wait(250);
const chs = () => X(() => window.__xl.wb.meta(window.__xl.grid.sheet).charts);
ok((await chs())[0]?.type === 'scatter', 'chèn biểu đồ phân tán');
ok((await X(() => document.querySelector('#tabs [data-tab="cdesign"]')?.getAttribute('aria-selected'))) === 'true', 'chọn biểu đồ → tab Thiết kế Biểu đồ');
await page.click('#ribbon [data-id="cd-elem"]'); await wait(100);
await page.click('.menu [role=menuitemcheckbox]:has-text("Tiêu đề Trục Ngang")'); await wait(150);
await page.fill('.dlg .an-form input', 'Tháng (số)'); await page.click('.dlg-foot .primary'); await wait(150);
ok((await chs())[0]?.xTitle === 'Tháng (số)', 'Thêm Thành phần ▸ Tiêu đề Trục Ngang');
await page.click('#ribbon [data-id="cd-colors"]'); await wait(80);
await page.click('.menu [role=menuitemcheckbox]:has-text("Xanh lá")'); await wait(100);
ok((await chs())[0]?.colors === 3, 'Đổi Màu');
await page.click('#tabs [data-tab="cformat"]'); await wait(80);
await page.click('#ribbon [data-id="cf-fill"]'); await wait(80);
await page.click('.menu [role=menuitemcheckbox]:has-text("Xám nhạt")'); await wait(100);
ok((await chs())[0]?.fill === '#f2f2f2', 'Định dạng ▸ Tô Hình');
await page.screenshot({ path: OUT + 'v09-chart.png' });
await goto('E8'); await wait(80);
ok((await X(() => document.querySelector('#tabs [data-tab="cdesign"]'))) === null, 'bỏ chọn biểu đồ → ẩn tab');
// Kết hợp
await X(() => window.__xl.grid.selectRange({ r1: 0, c1: 0, r2: 3, c2: 2 }));
await tab('insert'); await page.click('#ribbon [data-id="i-combo"]'); await wait(200);
await page.click('.dlg label.chk:has-text("trục phụ") input'); await wait(60);
await page.click('.dlg-foot .primary'); await wait(250);
{ const c = (await chs())[1]; ok(c?.type === 'combo' && c.series[c.series.length - 1].y2 === true && c.series[0].kind === 'col', 'biểu đồ Kết hợp có trục phụ'); }
// Sparkline
await goto('E8'); await focusGrid();
await X(() => window.__xl.grid.selectRange({ r1: 1, c1: 1, r2: 3, c2: 2 }));
await tab('insert'); await page.click('#ribbon [data-id="sp-col"]'); await wait(150);
ok((await X(() => document.querySelectorAll('.dlg .an-form input')[1].value)) === 'D2:D4', 'gợi ý vị trí D2:D4');
await page.click('.dlg-foot .primary'); await wait(150);
ok((await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).sparks?.[0]?.items.length)) === 3, 'tạo 3 sparkline cột');
ok((await X(() => document.querySelector('#tabs [data-tab="sdesign"]'))) !== null, 'tab Biểu đồ thu nhỏ');
await tab('sdesign'); await page.click('#ribbon [data-id="sd-line"]'); await wait(80);
await page.click('#ribbon [data-id="sd-hi"] input'); await wait(80);
ok(await X(() => { const g = window.__xl.wb.meta(window.__xl.grid.sheet).sparks[0]; return g.type === 'line' && g.hi === true; }), 'đổi loại + Điểm Cao');
await page.screenshot({ path: OUT + 'v09-spark.png' });

// ================= 0.9 Đối tượng =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await goto('B2'); await tab('insert');
await page.click('#ribbon [data-id="shapes"]'); await wait(80);
await page.click('.menu [role=menuitem]:has-text("Hình bầu dục")'); await wait(200);
const objs = () => X(() => window.__xl.wb.meta(window.__xl.grid.sheet).charts);
ok((await objs())[0]?.shape === 'ellipse', 'chèn hình bầu dục');
ok((await X(() => document.querySelector('#tabs [data-tab="oformat"]')?.getAttribute('aria-selected'))) === 'true', 'tab Định dạng Hình');
await page.click('#ribbon [data-id="of-fill"]'); await wait(80);
await page.click('.menu [role=menuitemcheckbox]:has-text("Xanh lá")'); await wait(100);
ok((await objs())[0]?.fill === '#70ad47', 'Tô Hình xanh lá');
await page.click('#ribbon [data-id="of-text"]'); await wait(100);
await page.fill('.obj-ta', 'Trụ đèn'); await page.click('.dlg-foot .primary'); await wait(100);
ok((await objs())[0]?.text === 'Trụ đèn', 'sửa chữ trong hình');
await goto('E2'); await tab('insert'); await rb('textbox'); await wait(200);
await page.fill('.obj-ta', 'Ghi chú hồ sơ'); await page.click('.dlg-foot .primary'); await wait(100);
ok((await objs())[1]?.type === 'text' && (await objs())[1]?.text === 'Ghi chú hồ sơ', 'Hộp Văn bản');
await tab('oformat'); await page.click('#ribbon [data-id="of-back"] .main'); await wait(80);
ok((await objs())[0]?.type === 'text', 'Đưa Xuống dưới');
await goto('B10'); await tab('insert'); await rb('icons'); await wait(150);
await page.click('.ic-it[title="home"]'); await page.click('.dlg-foot .primary'); await wait(300);
ok((await objs())[2]?.type === 'pic' && (await objs())[2]?.img?.startsWith('data:image/png'), 'chèn Biểu tượng (ảnh PNG)');
await page.screenshot({ path: OUT + 'v09-objects.png' });
await K('Delete'); await wait(100);
ok((await objs()).length === 2, 'Delete xóa đối tượng đang chọn');
// Hộp kiểm trong ô
await goto('H2:H4'); await tab('insert'); await rb('i-check'); await wait(100);
ok((await val('H2')) === false && (await val('H4')) === false, 'Hộp kiểm: ô = FALSE');
const cb = await X(() => { const g = window.__xl.grid, b = g.canvas.getBoundingClientRect(), r = g.checkRect(2, 7); return [b.left + r.x + r.s / 2, b.top + r.y + r.s / 2]; });
await page.mouse.click(cb[0], cb[1]); await wait(100);
ok((await val('H3')) === true, 'bấm hộp kiểm → TRUE');
await goto('H4'); await focusGrid(); await K('Space'); await wait(80);
ok((await val('H4')) === true, 'Space đổi hộp kiểm');
// Ký hiệu
await goto('J2'); await tab('insert'); await rb('symbol'); await wait(120);
await page.click('.sym-it:has-text("±")'); await page.click('.dlg-foot .primary'); await wait(80);
await K('Enter'); await wait(60);
ok((await val('J2')) === '±', 'chèn ký hiệu ±');

// ================= Kéo viền vùng chọn =================
await X(() => { const x = window.__xl; x.wb.addSheet(); const ids = x.wb.sheetIds(); x.grid.setSheet(ids[ids.length - 1]); x.renderSheets(); x.syncSelection(); });
await wait(150);
await put([['A1', 'a'], ['A2', 'b'], ['A3', 'c'], ['A4', 'd'], ['B1', '=A1&"!"']]);
const edgeOf = (a, side) => X(([r, c, side]) => { const g = window.__xl.grid, b = g.canvas.getBoundingClientRect(), R = g.cellRectM(r, c); return side === 'top' ? [b.left + R.x + R.w / 2, b.top + R.y + 1] : [b.left + R.x + R.w / 2, b.top + R.y + R.h / 2]; }, [...A(a), side]);
await goto('A1');
let p0 = await edgeOf('A1', 'top'), p1 = await edgeOf('A6', 'mid');
await page.mouse.move(p0[0], p0[1]); await wait(30);
ok((await X(() => window.__xl.grid.scroller.style.cursor)) === 'move', 'con trỏ di chuyển trên viền vùng chọn');
await page.mouse.down(); await page.mouse.move(p1[0], p1[1], { steps: 6 }); await page.mouse.up(); await wait(120);
ok((await val('A1')) === null && (await val('A6')) === 'a' && (await raw('B1')) === '=A6&"!"', 'kéo viền: di chuyển ô, công thức đi theo');
await K('Control+z'); await wait(80);
ok((await val('A1')) === 'a' && (await val('A6')) === null, 'hoàn tác kéo viền');
await goto('A1'); p0 = await edgeOf('A1', 'top'); p1 = await edgeOf('A4', 'top');
await page.mouse.move(p0[0], p0[1]); await page.mouse.down(); await page.keyboard.down('Shift');
await page.mouse.move(p1[0], p1[1] + 2, { steps: 6 }); await page.mouse.up(); await page.keyboard.up('Shift'); await wait(120);
ok([await val('A1'), await val('A2'), await val('A3'), await val('A4')].join() === 'b,c,a,d', 'Shift+kéo: chèn ô, các ô dồn lại');
await goto('C1'); p0 = await edgeOf('A1', 'top');
await goto('A1'); p1 = await edgeOf('C1', 'mid');
await page.mouse.move(p0[0], p0[1]); await page.mouse.down(); await page.keyboard.down('Control');
await page.mouse.move(p1[0], p1[1], { steps: 6 }); await page.mouse.up(); await page.keyboard.up('Control'); await wait(120);
ok((await val('A1')) === 'b' && (await val('C1')) === 'b', 'Ctrl+kéo: sao chép');

// ================= Bảo vệ Trang tính =================
await goto('E1'); await focusGrid(); await K('Control+1'); await wait(150);
await page.click('[role=tab][data-t=prot]'); await wait(80);
await X(() => { const c = document.getElementById('fc-lock'); c.checked = false; c.dispatchEvent(new Event('change')); });
await page.click('.dlg-foot .primary'); await wait(100);
ok((await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 0, 4).lk)) === false, 'bỏ khóa ô E1');
await tab('review'); await rb('r-psheet'); await wait(150);
await page.fill('.pr-dlg input[type=password] >> nth=0', 'abc'); await page.fill('.pr-dlg input[type=password] >> nth=1', 'abc');
await page.click('.dlg-foot .primary'); await wait(600);
ok(!!(await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).protect)), 'bảo vệ trang tính có mật khẩu');
await goto('A1'); await focusGrid(); await page.keyboard.type('x'); await wait(150);
ok((await val('A1')) === 'b', 'ô khóa không sửa được');
ok(await X(() => /bảo vệ/i.test(document.body.innerText)), 'báo ô đang được bảo vệ');
await K('Escape');
await goto('E1'); await focusGrid(); await page.keyboard.type('12'); await K('Enter'); await wait(100);
ok((await val('E1')) === 12, 'ô đã bỏ khóa vẫn sửa được');
await tab('review'); await rb('r-psheet'); await wait(150);
await page.fill('.dlg input[type=password]', 'sai'); await page.click('.dlg-foot .primary'); await wait(600);
ok(await X(() => /không đúng/.test(document.querySelector('.dlg-err')?.textContent ?? '')), 'mật khẩu sai bị từ chối');
await page.fill('.dlg input[type=password]', 'abc'); await page.click('.dlg-foot .primary'); await wait(600);
ok(!(await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).protect)), 'bỏ bảo vệ đúng mật khẩu');

// ================= Chính tả F7 =================
await put([['G1', 'Thay đèn chiếu sáng'], ['G2', 'Đường Nguyễn Huệ teh']]);
await goto('G1'); await focusGrid(); await K('F7'); await wait(1500);
ok(await X(() => document.querySelector('.sp-ctx mark')?.textContent === 'teh'), 'F7 tìm từ sai “teh”');
ok(await X(() => [...document.querySelectorAll('.sp-dlg option')].some((o) => o.value === 'the')), 'gợi ý “the”');
await X(() => { const s = document.querySelector('.sp-dlg select'); s.value = 'the'; });
await page.click('.sp-dlg button:has-text("Thay đổi") >> nth=0'); await wait(200);
ok((await val('G2')) === 'Đường Nguyễn Huệ the', 'Thay đổi từ sai');
await page.keyboard.press('Escape'); await wait(80);

console.log(`${pass}/${pass + fail} kiểm thử 0.8–0.9 đạt`);
ok(errors.length === 0, 'không lỗi trang: ' + errors.join(' | '));
await browser.close();
try { process.kill(-srv.pid); } catch { /* đã tắt */ }
process.exit(fail ? 1 : 0);
