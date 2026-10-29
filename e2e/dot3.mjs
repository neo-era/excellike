// Kiểm thử trình duyệt cho Đợt 3: biểu đồ (chèn, kéo, đổi cỡ, sửa, xóa, hoàn tác), PivotTable (tạo, chọn trường, làm mới), In / PDF, lưu .xlsx có biểu đồ, di động
import { createRequire } from 'node:module';
import { spawn, execFileSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');

const PORT = 4176, APP_URL = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 5000); });

const results = [];
const check = (name, ok, info = '') => { results.push({ name, ok }); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };
const browser = await chromium.launch();
const DATA = [
  ['Nhóm', 'Hạng mục', 'Thành tiền', 'SL'],
  ['Thiết bị', 'Đèn LED', '204000000', '48'],
  ['Vật tư', 'Cột thép', '326400000', '48'],
  ['Vật tư', 'Cáp ngầm', '148680000', '1260'],
  ['Thiết bị', 'Tủ điều khiển', '85500000', '3'],
  ['Xây lắp', 'Móng cột', '64800000', '48'],
];

async function newPage(opts = {}) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 820 }, acceptDownloads: true, ...opts });
  await ctx.addInitScript(() => {
    delete window.showOpenFilePicker; delete window.showSaveFilePicker;
    // bắt lệnh in trong khung in (không mở hộp thoại hệ thống)
    window.print = () => { window.top.__printed = document.documentElement.outerHTML; };
  });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  await page.goto(APP_URL);
  await page.waitForFunction(() => window.__xl);
  await page.evaluate((rows) => { const x = window.__xl; rows.forEach((row, r) => row.forEach((v, c) => x.wb.setInput(x.grid.sheet, r, c, v))); x.grid.invalidate(); }, DATA);
  return { ctx, page };
}
const X = (page, f, a) => page.evaluate(f, a);
const goto = async (page, a) => { await page.click('#namebox'); await page.keyboard.type(a); await page.keyboard.press('Enter'); };
const tab = (page, t) => page.click(`[role="tab"]:has-text("${t}")`);
const btn = (page, id) => page.click(`#ribbon [data-id="${id}"]`);
const charts = (page) => X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).charts);
const inkVariety = (page, sel) => X(page, (sel) => {
  const cv = document.querySelector(sel); const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data;
  const set = new Set(); for (let i = 0; i < d.length; i += 4 * 37) set.add(d[i] >> 4 << 8 | d[i + 1] >> 4 << 4 | d[i + 2] >> 4); return set.size;
}, sel);

// ---------------- Biểu đồ ----------------
{
  const { ctx, page } = await newPage();
  await goto(page, 'B1:C6');
  await tab(page, 'Chèn');
  await btn(page, 'chart');
  await page.waitForSelector('.dlg .chart-types');
  await page.waitForTimeout(150);
  check('Hộp thoại biểu đồ có xem trước', (await inkVariety(page, '.chart-prev')) > 4);
  await page.screenshot({ path: OUT + 'd3-chart-dialog.png' });
  await page.fill('.dlg input[placeholder="Tiêu đề biểu đồ"]', 'Giá trị theo hạng mục');
  await page.click('.dlg .btn.primary');
  await page.waitForSelector('.chart-box');
  await page.waitForTimeout(200);
  let ch = await charts(page);
  check('Chèn biểu đồ cột từ vùng chọn (tiêu đề cột → tên chuỗi, cột nhãn → danh mục)', ch.length === 1 && ch[0].type === 'col' && ch[0].series.length === 1 && ch[0].series[0].cats?.range.c1 === 1 && ch[0].anchor.c === 4, JSON.stringify(ch[0]?.anchor));
  check('Biểu đồ được vẽ trên lưới', (await inkVariety(page, '.chart-box canvas')) > 4);
  await page.screenshot({ path: OUT + 'd3-chart.png' });

  // kéo di chuyển
  let box = await page.locator('.chart-box').boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 120, box.y + box.height / 2 + 90, { steps: 6 });
  await page.mouse.up();
  const a0 = ch[0].anchor;
  ch = await charts(page);
  const moved = ch[0].anchor;
  check('Kéo biểu đồ đổi vị trí neo', moved.c > a0.c || moved.dx !== a0.dx, JSON.stringify(moved));
  await page.keyboard.press('Control+z');
  ch = await charts(page);
  check('Ctrl+Z hoàn tác di chuyển biểu đồ', JSON.stringify(ch[0].anchor) === JSON.stringify(a0), JSON.stringify(ch[0].anchor));

  // đổi cỡ bằng góc dưới phải
  await page.click('.chart-box', { position: { x: 30, y: 60 } });
  const hd = await page.locator('.chart-box .h-se').boundingBox();
  await page.mouse.move(hd.x + hd.width / 2, hd.y + hd.height / 2);
  await page.mouse.down();
  await page.mouse.move(hd.x + 100, hd.y + 60, { steps: 5 });
  await page.mouse.up();
  ch = await charts(page);
  check('Kéo góc đổi cỡ biểu đồ', ch[0].w > 560 && ch[0].h > 330, `${ch[0].w}×${ch[0].h}`);

  // nhấp đúp → sửa thành biểu đồ tròn, hiện nhãn
  await page.dblclick('.chart-box', { position: { x: 40, y: 80 } });
  await page.waitForSelector('.dlg h2:has-text("Sửa biểu đồ")');
  await page.click('.dlg .chart-types [data-t="pie"]');
  await page.click('.dlg .btn.primary');
  ch = await charts(page);
  check('Sửa biểu đồ qua hộp thoại (đổi sang Tròn)', ch[0].type === 'pie' && ch[0].title === 'Giá trị theo hạng mục');
  await page.waitForTimeout(100);
  await page.screenshot({ path: OUT + 'd3-chart-pie.png' });

  // dữ liệu đổi → biểu đồ vẽ lại
  const before = await X(page, () => document.querySelector('.chart-box canvas').toDataURL());
  await X(page, () => { window.__xl.wb.setInput(0, 1, 2, '900000000'); });
  await page.waitForTimeout(150);
  const after = await X(page, () => document.querySelector('.chart-box canvas').toDataURL());
  check('Biểu đồ cập nhật khi dữ liệu nguồn đổi', before !== after);

  // xóa bằng phím Delete, hoàn tác
  await page.click('.chart-box', { position: { x: 30, y: 60 } });
  await page.keyboard.press('Delete');
  check('Phím Delete xóa biểu đồ đang chọn, ô không bị xóa', (await charts(page)).length === 0 && (await X(page, () => window.__xl.wb.display(0, 1, 1).text)) === 'Đèn LED');
  await page.keyboard.press('Control+z');
  check('Hoàn tác xóa biểu đồ', (await charts(page)).length === 1);

  // chèn hàng phía trên → biểu đồ dời theo
  await page.click('#gridwrap', { position: { x: 80, y: 40 } });
  await X(page, () => window.__xl.wb.insertRows(0, 0, 2));
  ch = await charts(page);
  check('Chèn hàng: vùng dữ liệu & vị trí biểu đồ dời theo', ch[0].series[0].vals.range.r1 === 3 && ch[0].anchor.r === 2, JSON.stringify([ch[0].series[0].vals.range, ch[0].anchor]));
  await X(page, () => window.__xl.wb.undo());

  // chèn nhanh biểu đồ đường
  await goto(page, 'A1');
  await btn(page, 'i-chart');
  await page.waitForSelector('.dlg .chart-types [data-t="line"].on');
  await page.click('.dlg .btn.primary');
  ch = await charts(page);
  check('Chèn nhanh biểu đồ Đường từ vùng liền kề ô hiện hành', ch.length === 2 && ch[1].type === 'line' && ch[1].series.length === 2 && ch[1].series[0].cats?.range.c1 === 1, JSON.stringify(ch[1]?.series.map((s) => s.vals.range)));

  // lưu .xlsx có biểu đồ + thiết lập in ngang
  await X(page, () => { window.__xl.wb.setPrint(0, { orientation: 'landscape', paper: 'A4', fitWidth: true, titleRows: 1, gridlines: false }); window.__xl.wb.setColWidth(0, 1, 110); window.__xl.wb.setColWidth(0, 2, 100); });
  await tab(page, 'Tệp'); await page.click('#backstage [data-bs="saveas"]');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#backstage [data-a="xlsx"]')]);
  const path = OUT + 'd3-charts.xlsx';
  await dl.saveAs(path);
  const py = execFileSync('python3', ['-c', `
import openpyxl,sys
wb=openpyxl.load_workbook(sys.argv[1])
ws=wb.worksheets[0]
print(','.join(type(c).__name__ for c in ws._charts), ws.page_setup.orientation, ws.print_title_rows)`, path]).toString().trim();
  check('Lưu .xlsx: Excel/openpyxl đọc được 2 biểu đồ + in ngang + lặp tiêu đề', py === 'PieChart,LineChart landscape $1:$1', py);

  // In / PDF
  await page.keyboard.press('Escape');
  await goto(page, 'A1');
  await page.keyboard.press('Control+p');
  await page.waitForSelector('#backstage .bs-preview iframe');
  await page.waitForTimeout(400);
  await page.screenshot({ path: OUT + 'd3-print-dialog.png' });
  await page.click('#bs-print');
  await page.waitForFunction(() => window.__printed, null, { timeout: 5000 }).catch(() => {});
  const html = await X(page, () => window.__printed || '');
  check('In: tạo trang in có bảng và ảnh biểu đồ', html.includes('Đèn LED') && (html.match(/<img /g) || []).length === 2 && /size: 297mm 210mm/.test(html));
  if (html) {
    const pp = await ctx.newPage();
    await pp.setContent(html);
    await pp.waitForTimeout(300);
    const pdfPath = OUT + 'd3-print.pdf';
    await pp.pdf({ path: pdfPath, preferCSSPageSize: true, printBackground: true });
    const txt = execFileSync('pdftotext', ['-layout', pdfPath, '-']).toString();
    const info = execFileSync('pdfinfo', [pdfPath]).toString();
    const m = /Page size:\s+([\d.]+) x ([\d.]+)/.exec(info);
    check('PDF: có dữ liệu đã định dạng, khổ A4 ngang', txt.includes('Cột thép') && txt.includes('326400000') && m && +m[1] > +m[2], m?.[0]);
    await pp.screenshot({ path: OUT + 'd3-print-page.png', fullPage: true });
    await pp.close();
  }

  await page.keyboard.press('Escape');
  // giao diện tối
  await X(page, () => window.__xl.setTheme('dark'));
  await page.waitForTimeout(250);
  await page.screenshot({ path: OUT + 'd3-chart-dark.png' });
  await X(page, () => window.__xl.setTheme('light'));
  check('Không có lỗi JavaScript (biểu đồ)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ---------------- PivotTable ----------------
{
  const { ctx, page } = await newPage();
  await goto(page, 'A2');
  await tab(page, 'Chèn');
  await btn(page, 'pivot');
  await page.waitForSelector('.dlg h2:has-text("Tạo PivotTable")');
  const src = await page.inputValue('.dlg input');
  check('Hộp thoại PivotTable tự nhận vùng dữ liệu', src === 'A1:D6', src);
  await page.click('.dlg .btn.primary');
  await page.waitForSelector('.pv-panel:not([hidden])');
  const sh = await X(page, () => window.__xl.wb.sheetName(window.__xl.grid.sheet));
  check('Tạo PivotTable ở sheet mới và mở ngăn trường', sh === 'Pivot 1', sh);
  await page.click('.pv-f:has-text("Nhóm") input');
  await page.click('.pv-f:has-text("Thành tiền") input');
  const v = await X(page, () => { const x = window.__xl, s = x.grid.sheet; return [2, 3, 4, 5, 6].map((r) => [0, 1].map((c) => x.wb.display(s, r, c).text)); });
  check('Tích trường: chữ → Hàng, số → Giá trị (Tổng)', JSON.stringify(v) === JSON.stringify([['Nhóm', 'Tổng Thành tiền'], ['Thiết bị', '289.500.000'], ['Vật tư', '475.080.000'], ['Xây lắp', '64.800.000'], ['Tổng cộng', '829.380.000']]), JSON.stringify(v));
  await page.screenshot({ path: OUT + 'd3-pivot.png' });
  // đổi cách tính sang Đếm
  await page.selectOption('.pv-zone[data-z="values"] .pv-chip select >> nth=0', 'count');
  check('Đổi cách tính giá trị sang Đếm', (await X(page, () => window.__xl.wb.display(window.__xl.grid.sheet, 3, 1).text)) === '2');
  // chuyển Nhóm sang Cột
  await page.selectOption('.pv-zone[data-z="rows"] .pv-chip select', 'cols');
  const hdr = await X(page, () => { const x = window.__xl, s = x.grid.sheet; return [0, 1, 2, 3, 4].map((c) => x.wb.display(s, 3, c).text); });
  check('Chuyển trường sang Cột → bảng ngang', hdr.join('|').startsWith('Thiết bị|Vật tư|Xây lắp|Tổng cộng') || hdr.join('|').includes('Thiết bị|Vật tư|Xây lắp|Tổng cộng'), hdr.join('|'));
  await page.keyboard.press('Control+z');
  // dữ liệu nguồn đổi → Làm mới
  await X(page, () => { const x = window.__xl; x.wb.setInput(x.wb.sheetIds()[0], 5, 2, '100'); });
  await page.click('.pv-foot .btn:has-text("Làm mới")');
  const after = await X(page, () => window.__xl.wb.display(window.__xl.grid.sheet, 3, 1).text);
  check('Làm mới lấy dữ liệu nguồn mới', after === '2', after);
  // lọc báo cáo
  await page.selectOption('.pv-zone[data-z="values"] .pv-chip select >> nth=0', 'sum');
  await X(page, () => { const x = window.__xl; const p = x.wb.meta(x.grid.sheet).pivots[0]; x.wb.refreshPivot(x.grid.sheet, p.id, { filters: [{ field: 'Nhóm', values: ['Vật tư'] }] }); });
  await page.waitForTimeout(50);
  // chọn ô ngoài Pivot → ngăn ẩn; chọn lại → hiện
  await goto(page, 'J20');
  const hid = await page.locator('.pv-panel').isHidden();
  await goto(page, 'A4');
  const vis = await page.locator('.pv-panel').isVisible();
  check('Ngăn trường tự ẩn/hiện theo ô hiện hành', hid && vis);
  await page.click('.pv-foot .btn.danger');
  const left = await X(page, () => ({ n: window.__xl.wb.meta(window.__xl.grid.sheet).pivots.length, a3: window.__xl.wb.display(window.__xl.grid.sheet, 2, 0).text }));
  check('Xóa PivotTable xóa cả ô kết quả', left.n === 0 && left.a3 === '', JSON.stringify(left));
  await page.keyboard.press('Control+z');
  check('Hoàn tác xóa PivotTable', (await X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).pivots.length)) === 1);
  check('Không có lỗi JavaScript (Pivot)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ---------------- Di động ----------------
for (const [name, vp] of [['portrait', { width: 390, height: 844 }], ['landscape', { width: 844, height: 390 }]]) {
  const { ctx, page } = await newPage({ viewport: vp, isMobile: true, hasTouch: true, deviceScaleFactor: 2 });
  await X(page, () => {
    const x = window.__xl;
    x.wb.addChart(0, { type: 'col', title: 'Thành tiền', series: [{ name: { sheet: 0, range: { r1: 0, c1: 2, r2: 0, c2: 2 } }, cats: { sheet: 0, range: { r1: 1, c1: 1, r2: 5, c2: 1 } }, vals: { sheet: 0, range: { r1: 1, c1: 2, r2: 5, c2: 2 } } }], anchor: { r: 7, c: 0, dx: 4, dy: 4 }, w: 340, h: 220 });
    x.grid.invalidate();
  });
  await X(page, () => window.__xl.grid.ensureVisible(7, 0));
  await page.waitForTimeout(200);
  const b = await page.locator('.chart-box').boundingBox();
  await page.touchscreen.tap(b.x + 40, b.y + 8);
  await page.waitForTimeout(100);
  const sel = await page.locator('.chart-box.on .chart-bar').isVisible();
  check(`Di động ${name}: chạm chọn biểu đồ hiện nút Sửa/Xóa`, sel);
  await page.screenshot({ path: OUT + `d3-mobile-chart-${name}.png` });
  await X(page, () => { const x = window.__xl; const r = x.wb.addPivot({ src: { sheet: 0, range: { r1: 0, c1: 0, r2: 5, c2: 3 } }, rows: ['Nhóm'], cols: [], values: [{ field: 'SL', agg: 'sum' }], filters: [] }); x.grid.setSheet(r.sheet); x.renderSheets(); x.grid.select(3, 0); x.pivot.show(r.sheet, r.id); });
  await page.waitForTimeout(200);
  const pb = await page.locator('.pv-panel').boundingBox();
  check(`Di động ${name}: ngăn Pivot nằm trong màn hình`, pb && pb.x >= 0 && pb.x + pb.width <= vp.width + 1 && pb.y + pb.height <= vp.height + 1, JSON.stringify(pb));
  await page.screenshot({ path: OUT + `d3-mobile-pivot-${name}.png` });
  check(`Di động ${name}: không lỗi JavaScript`, page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

await browser.close();
try { process.kill(-srv.pid); } catch { /* đã dừng */ }
const bad = results.filter((r) => !r.ok);
console.log(`\n${results.length - bad.length}/${results.length} kiểm thử Đợt 3 đạt`);
writeFileSync(OUT + 'd3-results.json', JSON.stringify(results, null, 1));
process.exit(bad.length ? 1 : 0);
