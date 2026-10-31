// Kiểm thử bản 0.7 trên giao diện thật (bấm phím, bấm nút): Tên vùng, Ghi chú/Chú thích, Nhóm hàng/cột, Tổng phụ,
// Văn bản thành Cột, Xóa trùng lặp, Bảng, mảng động.
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');
const PORT = 4181, URL0 = `http://localhost:${PORT}/`;
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
const toastText = () => X(() => [...document.querySelectorAll('.toast')].map((t) => t.textContent).join('|'));

// ================= Tên vùng =================
await put([['A1', 'Khối lượng'], ['B1', 'Đơn giá'], ['A2', '2'], ['B2', '100'], ['A3', '3'], ['B3', '200']]);
// Tạo từ vùng chọn (Ctrl+Shift+F3)
await goto('A1:B3'); await focusGrid();
await K('Control+Shift+F3'); await wait(150);
ok(await page.$('.dlg') !== null && (await X(() => document.querySelector('.dlg h2').textContent)) === 'Tạo Tên từ Vùng chọn', 'Ctrl+Shift+F3 mở Tạo Tên từ Vùng chọn');
ok(await X(() => document.querySelector('.nm-create input').checked), 'đoán sẵn Hàng trên cùng');
await page.click('.dlg-foot .btn.primary'); await wait();
ok((await X(() => window.__xl.wb.listNames().map((n) => n.name).join(','))) === 'Đơn_giá,Khối_lượng', 'tạo 2 tên từ tiêu đề');
// dùng tên trong công thức gõ tay
await goto('D1'); await focusGrid();
await page.keyboard.type('=SUMPRODUCT(Khối_lượng;Đơn_giá)'); await K('Enter'); await wait();
ok((await val('D1')) === 800, 'công thức dùng tên tiếng Việt có dấu = 800');
// Hộp Tên: gõ tên mới đặt cho vùng chọn
await goto('B2:B3');
await page.click('#namebox'); await page.keyboard.type('GiaVT'); await K('Enter'); await wait();
ok((await X(() => window.__xl.wb.listNames().some((n) => n.name === 'GiaVT' && n.expr === '=Sheet1!$B$2:$B$3'))), 'Hộp Tên: gõ tên mới → đặt tên vùng chọn');
// Hộp Tên: chọn tên trong danh sách → nhảy tới
await goto('H20');
await page.click('#namebox-dd'); await wait();
await page.click('.menu .mi:has-text("GiaVT")'); await wait();
ok((await page.inputValue('#namebox')) === 'B2', 'danh sách Hộp Tên: nhảy tới vùng của tên');
// gõ tên đã có → nhảy tới
await goto('Khối_lượng');
ok(JSON.stringify(await X(() => window.__xl.grid.range)) === JSON.stringify({ r1: 1, c1: 0, r2: 2, c2: 0 }), 'gõ tên đã có vào Hộp Tên → chọn vùng');
// F3 trong khi gõ công thức
await goto('D2'); await focusGrid();
await page.keyboard.type('=SUM('); await K('F3'); await wait(150);
ok(await page.$('.dlg .nm-list') !== null, 'F3 khi đang gõ: hộp Dán Tên');
await page.selectOption('.dlg .nm-list', 'GiaVT'); await K('Enter'); await wait(100);
await page.keyboard.type(')'); await K('Enter'); await wait();
ok((await raw('D2')) === '=SUM(GiaVT)' && (await val('D2')) === 300, 'F3 chèn tên vào công thức: ' + (await raw('D2')));
// Trình Quản lý Tên: sửa đổi tên → công thức cập nhật
await focusGrid(); await K('Control+F3'); await wait(150);
ok((await X(() => document.querySelectorAll('.nm-table tbody tr').length)) === 3, 'Ctrl+F3: Trình Quản lý Tên liệt kê 3 tên');
await page.screenshot({ path: OUT + 'v07-names.png' });
await page.click('.nm-table tbody tr:has-text("GiaVT")');
await page.click('.nm-bar .btn:has-text("Sửa")'); await wait(100);
await page.fill('.nm-form input >> nth=0', 'GiaVatTu');
await page.click('.dlg-foot .btn.primary'); await wait(150);
ok((await X(() => !!document.querySelector('.nm-table'))), 'sửa xong quay lại Trình Quản lý Tên');
await page.click('.dlg-foot .btn.primary'); await wait();
ok((await raw('D2')) === '=SUM(GiaVatTu)', 'đổi tên cập nhật công thức: ' + (await raw('D2')));
// tên sai báo lỗi
await tab('formulas');
await page.click('#ribbon [data-id="nm-def"]'); await wait(100);
await page.fill('.nm-form input >> nth=0', 'A1');
await page.click('.dlg-foot .btn.primary'); await wait();
ok(/địa chỉ/.test(await X(() => document.querySelector('.dlg-err').textContent)), 'Xác định Tên: tên dạng A1 bị từ chối');
await K('Escape'); await wait();
// hoàn tác
await focusGrid(); await K('Control+z'); await wait();
ok((await raw('D2')) === '=SUM(GiaVT)', 'Ctrl+Z hoàn tác đổi tên');
await K('Control+y'); await wait();

// ================= Ghi chú & Chú thích =================
await X(() => window.__xl.grid.select(5, 1)); await focusGrid();
await K('Shift+F2'); await wait(120);
ok(await page.$('.nt-edit textarea') !== null, 'Shift+F2 mở hộp ghi chú');
ok((await page.inputValue('.nt-edit textarea')).endsWith(':\n'), 'ghi chú mới có sẵn tên người viết');
await page.keyboard.type('Kiểm tra lại khối lượng'); await K('Escape'); await wait(100);
ok(/Kiểm tra lại khối lượng/.test(await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 5, 1).nt ?? '')), 'Esc lưu ghi chú');
// rê chuột hiện ghi chú
const rc = await X(() => window.__xl.grid.cellPageRect(5, 1));
await page.mouse.move(rc.left + rc.width / 2, rc.top + rc.height / 2); await wait(450);
ok(await page.$('.nt-box') !== null, 'rê chuột vào ô hiện ghi chú');
await page.screenshot({ path: OUT + 'v07-note.png' });
await page.mouse.move(5, 300); await wait(100);
ok(await page.$('.nt-box') === null, 'rời chuột thì ẩn');
// chú thích dạng chuỗi
await X(() => window.__xl.grid.select(6, 2)); await focusGrid();
await K('Control+Shift+F2'); await wait(120);
ok(await page.$('.cm-card textarea') !== null, 'Ctrl+Shift+F2 mở thẻ chú thích');
await page.keyboard.type('Đơn giá này lấy từ đâu?'); await page.click('.cm-card .cm-send'); await wait(100);
await page.fill('.cm-card textarea', 'Theo công bố tháng 9'); await K('Control+Enter'); await wait(100);
ok((await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 6, 2).cm?.length)) === 2, 'đăng + trả lời: chuỗi 2 bình luận');
await page.screenshot({ path: OUT + 'v07-comment.png' });
await K('Escape'); await wait();
// ngăn chú thích
await page.click('#btn-comments'); await wait(100);
ok((await X(() => document.querySelectorAll('.cm-pane .cm-row').length)) === 1, 'nút Chú thích mở ngăn liệt kê chuỗi');
await page.click('#btn-comments'); await wait();
// Ctrl+Shift+O chọn ô có ghi chú
await focusGrid(); await K('Control+Shift+O'); await wait();
ok((await X(() => window.__xl.grid.ranges().length)) === 2, 'Ctrl+Shift+O chọn 2 ô có ghi chú/chú thích');
// Trước/Tiếp + xóa
await tab('review');
await X(() => window.__xl.grid.select(0, 0));
await page.click('#ribbon [data-id="r-cnext"]'); await wait(100);
ok((await page.inputValue('#namebox')) === 'C7', 'Chú thích Tiếp theo tới C7');
await K('Escape');
await page.click('#ribbon [data-id="r-cdel"]'); await wait();
ok(!(await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 6, 2).cm)), 'Xóa chú thích');
await focusGrid(); await K('Control+z'); await wait();
ok((await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 6, 2).cm?.length)) === 2, 'hoàn tác xóa chú thích');

// ================= Nhóm hàng/cột & Tổng phụ =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await put([['A1', 'Loại'], ['B1', 'Tiền'], ['A2', 'Cát'], ['B2', '10'], ['A3', 'Cát'], ['B3', '20'], ['A4', 'Đá'], ['B4', '5'], ['A5', 'Đá'], ['B5', '7'], ['A6', 'Đá'], ['B6', '8']]);
const lv = (r) => X((r) => window.__xl.wb.outlineLevel(window.__xl.grid.sheet, 'r', r), r);
const hidden = (r) => X((r) => window.__xl.wb.meta(window.__xl.grid.sheet).rowH.hidden().includes(r), r);
await goto('A2:B3'); await focusGrid();
await K('Alt+Shift+ArrowRight'); await wait(150);
ok((await X(() => document.querySelector('.dlg h2')?.textContent)) === 'Nhóm', 'Alt+Shift+→ hỏi Hàng/Cột');
await page.click('.dlg-foot .primary'); await wait(120);
ok((await lv(1)) === 1 && (await lv(2)) === 1 && (await lv(3)) === 0, 'nhóm hàng 2:3 cấp 1');
ok((await X(() => window.__xl.grid.outW)) > 0, 'hiện dải đại cương bên trái');
await page.screenshot({ path: OUT + 'v07-outline.png' });
// bấm nút − ở hàng tổng (hàng 4)
const btnXY = (r, L) => X(([r, L]) => { const g = window.__xl.grid, b = g.canvas.getBoundingClientRect(); return [b.left + (L - 1) * 14 * g.zoom + 7 * g.zoom + 1, b.top + g.sy(r) + g.rowHpx(r) / 2]; }, [r, L]);
let [bx, by] = await btnXY(3, 1);
await page.mouse.click(bx, by); await wait(120);
ok((await hidden(1)) && (await hidden(2)), 'nút − thu gọn nhóm');
[bx, by] = await btnXY(3, 1);
await page.mouse.click(bx, by); await wait(120);
ok(!(await hidden(1)) && !(await hidden(2)), 'nút + mở lại nhóm');
// nút cấp 1
const lvXY = (L) => X((L) => { const g = window.__xl.grid, b = g.canvas.getBoundingClientRect(); return [b.left + (L - 1) * 14 * g.zoom + 7 * g.zoom + 1, b.top + g.hdrH - 8]; }, L);
[bx, by] = await lvXY(1); await page.mouse.click(bx, by); await wait(120);
ok(await hidden(1), 'nút cấp 1 thu gọn tất cả');
[bx, by] = await lvXY(2); await page.mouse.click(bx, by); await wait(120);
ok(!(await hidden(1)), 'nút cấp 2 hiện tất cả');
await focusGrid(); await K('Control+8'); await wait(80);
ok((await X(() => window.__xl.grid.outW)) === 0, 'Ctrl+8 ẩn ký hiệu đại cương');
await K('Control+8'); await wait(80);
await goto('A2:B3'); await focusGrid(); await K('Alt+Shift+ArrowLeft'); await wait(120);
await page.click('.dlg-foot .primary'); await wait(120);
ok((await lv(1)) === 0 && (await X(() => window.__xl.grid.outW)) === 0, 'Alt+Shift+← bỏ nhóm');
// Tổng phụ
await goto('A3'); await tab('data');
await page.click('#ribbon [data-id="d-sub"]'); await wait(150);
ok((await X(() => document.querySelector('.dlg h2')?.textContent)) === 'Tổng phụ', 'mở hộp Tổng phụ');
await page.click('.dlg-foot .primary'); await wait(200);
ok((await val('A4')) === 'Tổng Cát' && (await val('B4')) === 30, 'Tổng Cát = 30');
ok((await val('B8')) === 20 && (await val('A9')) === 'Tổng Cộng' && (await val('B9')) === 50, 'Tổng Đá 20, Tổng Cộng 50');
ok((await X(() => window.__xl.wb.outlineMax(window.__xl.grid.sheet, 'r'))) === 2, 'tổng phụ tạo 2 cấp');
[bx, by] = await lvXY(2); await page.mouse.click(bx, by); await wait(120);
ok((await hidden(1)) && !(await hidden(3)), 'cấp 2 chỉ hiện hàng tổng');
await page.screenshot({ path: OUT + 'v07-subtotal.png' });
await focusGrid(); await K('Control+z'); await wait(150);
await K('Control+z'); await wait(150);
ok((await val('A4')) === 'Đá' && (await X(() => window.__xl.wb.outlineMax(window.__xl.grid.sheet, 'r'))) === 0, 'hoàn tác tổng phụ');

// ================= Văn bản thành Cột & Loại bỏ Trùng lặp =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await put([['A1', 'Cát;12,5;m3'], ['A2', 'Đá;7;m3'], ['A3', 'Xi măng;1500;kg']]);
await goto('A1:A3'); await tab('data');
await page.click('#ribbon [data-id="d-t2c"]'); await wait(150);
ok((await X(() => document.querySelectorAll('.t2c-prev thead th').length)) === 3, 'xem trước 3 cột (tự đoán dấu ;)');
await page.click('.dlg-foot .primary'); await wait(150);
ok((await val('A3')) === 'Xi măng' && (await val('B3')) === 1500 && (await val('C1')) === 'm3' && (await val('B1')) === 12.5, 'tách cột đúng, nhận số');
await focusGrid(); await K('Control+z'); await wait(100);
ok((await val('A1')) === 'Cát;12,5;m3' && (await val('B1')) === null, 'hoàn tác Văn bản thành Cột');
await put([['A1', 'Vật tư'], ['B1', 'ĐG'], ['A2', 'Cát'], ['B2', '10'], ['A3', 'cát'], ['B3', '10'], ['A4', 'Đá'], ['B4', '5'], ['A5', 'Đá'], ['B5', '5']]);
await goto('A2'); await tab('data');
await page.click('#ribbon [data-id="d-dup"]'); await wait(150);
ok((await X(() => document.querySelectorAll('.ol-cols input:checked').length)) === 2, 'Loại bỏ Trùng lặp: 2 cột được chọn');
ok((await X(() => document.querySelector('.dup-dlg .nm-bar input').checked)), 'tự nhận có tiêu đề');
await page.click('.dlg-foot .primary'); await wait(150);
ok((await val('A3')) === 'Đá' && (await val('A4')) === null && (await val('A1')) === 'Vật tư', 'xóa 2 dòng trùng, giữ tiêu đề');
ok((await toastText()).includes('xóa 2 giá trị trùng lặp'), 'thông báo số dòng đã xóa');

// ================= Bảng (Ctrl+T) =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await put([['A1', 'Vật tư'], ['B1', 'KL'], ['C1', 'Đơn giá'], ['A2', 'Cát'], ['B2', '2'], ['C2', '100'], ['A3', 'Đá'], ['B3', '3'], ['C3', '200']]);
await goto('B2'); await focusGrid();
await K('Control+t'); await wait(150);
ok((await X(() => document.querySelector('.dlg h2')?.textContent)) === 'Tạo Bảng', 'Ctrl+T mở Tạo Bảng');
ok((await X(() => document.querySelector('.tb-new input.mono').value)) === '=$A$1:$C$3', 'tự lấy vùng hiện hành');
await page.click('.dlg-foot .primary'); await wait(200);
ok((await X(() => window.__xl.wb.tables(window.__xl.grid.sheet).length)) === 1, 'đã tạo bảng');
ok((await X(() => document.querySelector('#tabs [data-tab="tdesign"]')?.getAttribute('aria-selected'))) === 'true', 'mở tab Thiết kế Bảng');
ok((await X(() => document.querySelector('#ribbon [data-id="td-name"] input')?.value)) === 'Bảng1', 'tên bảng Bảng1');
await page.screenshot({ path: OUT + 'v07-table.png' });
// gõ tiêu đề bên phải → thêm cột; nhập công thức → cột tính toán
await goto('D1'); await page.keyboard.type('Thành tiền'); await K('Enter'); await wait(80);
ok((await X(() => window.__xl.wb.tables(window.__xl.grid.sheet)[0].c2)) === 3, 'gõ cạnh phải → bảng thêm cột');
await page.keyboard.type('=[@KL]*[@[Đơn giá]]'); await K('Enter'); await wait(120);
ok((await val('D2')) === 200 && (await val('D3')) === 600, 'cột tính toán tự điền');
ok((await raw('D3')) === '=[@KL]*[@[Đơn giá]]', 'hiện tham chiếu có cấu trúc');
// gõ dưới bảng → thêm hàng, công thức kế thừa
await goto('A4'); await page.keyboard.type('Thép'); await K('Tab'); await page.keyboard.type('5'); await K('Tab'); await page.keyboard.type('10'); await K('Enter'); await wait(120);
ok((await X(() => window.__xl.wb.tables(window.__xl.grid.sheet)[0].r2)) === 3 && (await val('D4')) === 50, 'gõ dưới bảng → thêm hàng, kế thừa công thức');
await goto('F1'); await page.keyboard.type('=SUM(Bảng1[Thành tiền])'); await K('Enter'); await wait(100);
ok((await val('F1')) === 850, 'SUM(Bảng1[Thành tiền]) = 850');
// Hàng Tổng
await goto('B2'); await tab('tdesign');
await page.click('#ribbon [data-id="td-tot"] input'); await wait(150);
ok((await val('A5')) === 'Tổng' && (await val('D5')) === 850, 'Hàng Tổng: Tổng 850');
await goto('B5'); await wait(80);
await K('Alt+ArrowDown'); await wait(150);
await page.click(".menu [role=menuitemcheckbox]:has-text(\"Trung bình\")"); await wait(150);
ok(Math.abs((await val('B5')) - 10 / 3) < 1e-9 && (await raw('B5')) === '=SUBTOTAL(101;[KL])', 'chọn hàm Trung bình ở hàng tổng');
// đổi kiểu
await goto('B2'); await tab('tdesign');
await page.click('#ribbon [data-id="td-styles"] .gal-more'); await wait(120);
await page.click('.tb-gal-it[data-style="TableStyleMedium7"]'); await wait(120);
ok((await X(() => window.__xl.wb.tables(window.__xl.grid.sheet)[0].style)) === 'TableStyleMedium7', 'đổi kiểu bảng');
// Hộp tên: chọn bảng
await page.click('#namebox'); await page.keyboard.type('Bảng1'); await K('Enter'); await wait(80);
ok((await X(() => JSON.stringify(window.__xl.grid.range))) === JSON.stringify({ r1: 1, c1: 0, r2: 3, c2: 3 }), 'Hộp Tên tới vùng dữ liệu Bảng1');
await page.screenshot({ path: OUT + 'v07-table2.png' });
// Chuyển thành Vùng
await goto('B2'); await tab('tdesign');
await page.click('#ribbon [data-id="td-conv"]'); await wait(150);
ok((await X(() => window.__xl.wb.tables(window.__xl.grid.sheet).length)) === 0 && (await raw('F1')) === '=SUM(Sheet4!$D$2:$D$4)', 'Chuyển thành Vùng: công thức về địa chỉ ô');
ok((await X(() => document.querySelector('#tabs [data-tab="tdesign"]'))) === null, 'ẩn tab Thiết kế Bảng');
await focusGrid(); await K('Control+z'); await wait(150);
ok((await X(() => window.__xl.wb.tables(window.__xl.grid.sheet).length)) === 1, 'hoàn tác Chuyển thành Vùng');

// ================= Mảng động & Truy tìm =================
await focusGrid(); await K('Shift+F11'); await wait(150);
await put([['A1', 'Cát'], ['B1', '3'], ['A2', 'Đá'], ['B2', '1'], ['A3', 'Thép'], ['B3', '2']]);
await goto('D1'); await page.keyboard.type('=SORT(A1:B3;2;-1)'); await K('Enter'); await wait(120);
ok((await val('D1')) === 'Cát' && (await val('E3')) === 1, 'SORT tràn ra vùng 3×2');
await goto('E2'); await wait(60);
ok((await X(() => document.querySelector('#fx-input').classList.contains('ghost'))) && (await page.inputValue('#fx-input')) === '=SORT(A1:B3;2;-1)', 'ô tràn hiện công thức gốc (mờ)');
await goto('G1'); await page.keyboard.type('=FILTER(A1:A3;B1:B3>1)'); await K('Enter'); await wait(120);
ok((await val('G2')) === 'Thép', 'FILTER theo điều kiện');
await goto('I1'); await page.keyboard.type('={1;2|3;4}'); await K('Enter'); await wait(120);
ok((await val('J2')) === 4, 'hằng mảng 2×2');
await goto('G2'); await page.keyboard.type('x'); await K('Enter'); await wait(120);
ok((await X(() => window.__xl.wb.display(window.__xl.grid.sheet, 0, 6).text)) === '#SPILL!', 'chặn vùng tràn → #SPILL!');
await goto('C5'); await page.keyboard.type('=SUM(B1:B3)*2'); await K('Enter'); await wait(60);
await goto('C5'); await tab('formulas');
await page.click('#ribbon [data-id="tr-pre"]'); await wait(100);
ok((await X(() => window.__xl.grid.arrows.length)) === 1, 'Truy tìm Ô Tiền lệ: 1 mũi tên');
await goto('B2'); await page.click('#ribbon [data-id="tr-dep"]'); await wait(100);
ok((await X(() => window.__xl.grid.arrows.length)) >= 3, 'Truy tìm Ô Phụ thuộc thêm mũi tên');
await page.screenshot({ path: OUT + 'v07-trace.png' });
await page.click('#ribbon [data-id="tr-rm"]'); await wait(80);
ok((await X(() => window.__xl.grid.arrows.length)) === 0, 'Loại bỏ Mũi tên');

console.log(`${pass}/${pass + fail} kiểm thử 0.7 đạt`);
ok(errors.length === 0, 'không lỗi trang: ' + errors.join(' | '));
await browser.close();
try { process.kill(-srv.pid); } catch { /* đã tắt */ }
process.exit(fail ? 1 : 0);
