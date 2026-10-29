// Kiểm thử trình duyệt cho Đợt 2: mở/lưu .xlsx, CSV, lọc, kiểm tra dữ liệu, gộp ô, viền, sắp xếp, định dạng có điều kiện
import { createRequire } from 'node:module';
import { spawn, execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');

const PORT = 4175, APP_URL = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
const FIX = new URL('../tests/fixtures/mau.xlsx', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 5000); });

const results = [];
const check = (name, ok, info = '') => { results.push({ name, ok }); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };
const browser = await chromium.launch();

async function newPage(opts = {}) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 820 }, acceptDownloads: true, ...opts });
  // dùng hộp chọn tệp thường (Playwright không điều khiển được hộp thoại hệ thống)
  await ctx.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  await page.goto(APP_URL);
  await page.waitForFunction(() => window.__xl);
  return { ctx, page };
}
const X = (page, f, a) => page.evaluate(f, a);
const disp = (page, r, c) => X(page, ({ r, c }) => window.__xl.wb.display(window.__xl.grid.sheet, r, c).text, { r, c });
const pt = (page, r, c) => X(page, ({ r, c }) => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), rc = g.cellRectM(r, c); return { x: b.left + rc.x + rc.w / 2, y: b.top + rc.y + rc.h / 2, x2: b.left + rc.x + rc.w, y2: b.top + rc.y + rc.h, top: b.top + rc.y }; }, { r, c });
const goto = async (page, a) => { await page.click('#namebox'); await page.keyboard.type(a); await page.keyboard.press('Enter'); };
const tab = (page, t) => page.click(`[role="tab"]:has-text("${t}")`);
const btn = (page, id) => page.click(`#ribbon [data-id="${id}"]`);
const menu = (page, label) => page.click(`.menu .mi:has-text("${label}")`);

// ---------------- Mở tệp .xlsx ----------------
{
  const { ctx, page } = await newPage();
  await tab(page, 'Tệp'); await page.click('#backstage [data-bs="open"]');
  const [fc] = await Promise.all([page.waitForEvent('filechooser'), page.click('#backstage [data-a="browse"]')]);
  await fc.setFiles(FIX);
  await page.waitForEvent('load');
  await page.waitForFunction(() => window.__xl);
  const info = await X(page, () => { const x = window.__xl, s = x.grid.sheet; return { sheets: x.wb.sheetIds().map((i) => x.wb.sheetName(i)), f2: x.wb.display(s, 1, 5).text, fz: x.wb.meta(s).freeze, merge: !!x.wb.mergeAt(s, 6, 2), filter: !!x.wb.meta(s).filter, name: x.wb.name, file: x.fileName }; });
  check('Mở tệp .xlsx: sheet, công thức, gộp ô, cố định ngăn, bộ lọc', info.sheets.join(',') === 'Dự toán,2024,Q1' && info.f2 === '204.000.000' && info.fz.rows === 1 && info.merge && info.filter, JSON.stringify(info));
  await page.screenshot({ path: OUT + 'd2-open-xlsx.png' });

  // Lọc: bấm nút lọc ở cột ĐVT (C1), chỉ giữ "bộ"
  const p = await pt(page, 0, 2);
  await page.mouse.click(p.x2 - 10, p.y);
  await page.waitForSelector('.filter-pop');
  await page.screenshot({ path: OUT + 'd2-filter-pop.png' });
  await page.click('.filter-pop .chk:has-text("(Chọn tất cả)") input');
  await page.click('.filter-pop .chk:has-text("bộ") input');
  await page.click('.filter-pop .btn.primary');
  const hidden = await X(page, () => [1, 2, 3].map((r) => window.__xl.wb.isRowHidden(window.__xl.grid.sheet, r)));
  check('Lọc theo giá trị ẩn đúng hàng', JSON.stringify(hidden) === '[false,true,true]', JSON.stringify(hidden));
  await page.screenshot({ path: OUT + 'd2-filtered.png' });
  await page.keyboard.press('Control+z');
  check('Hoàn tác lọc', !(await X(page, () => window.__xl.wb.isRowHidden(window.__xl.grid.sheet, 2))));

  // Kiểm tra dữ liệu: ô C3 có danh sách; chọn từ mũi tên
  await goto(page, 'C3');
  const db = await X(page, () => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), d = g.dvBtn(); return d && { x: b.left + d.x + d.w / 2, y: b.top + d.y + d.h / 2 }; });
  await page.mouse.click(db.x, db.y);
  await page.waitForSelector('.menu');
  await menu(page, 'm');
  check('Chọn giá trị từ danh sách thả xuống', (await disp(page, 2, 2)) === 'm');
  await page.keyboard.type('xyz'); await page.keyboard.press('Enter');
  const stillEditing = await X(page, () => window.__xl.ed.isEditing());
  const toast = await page.textContent('.toast.err').catch(() => '');
  await page.keyboard.press('Escape');
  check('Nhập sai danh sách bị chặn và báo lỗi', stillEditing && /danh sách/i.test(toast) && (await disp(page, 2, 2)) === 'm', toast);

  // Lưu .xlsx (tải xuống) rồi kiểm tra bằng openpyxl
  await tab(page, 'Tệp'); await page.click('#backstage [data-bs="saveas"]');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#backstage [data-a="xlsx"]')]);
  const out = OUT + 'luu-lai.xlsx';
  await dl.saveAs(out);
  const py = execFileSync('python3', ['-c', `
import openpyxl,json
wb=openpyxl.load_workbook('${out}')
ws=wb['Dự toán']
print(json.dumps({'sheets':wb.sheetnames,'F2':ws['F2'].value,'H3':ws['H3'].value,'merged':[str(m) for m in ws.merged_cells.ranges],'freeze':ws.freeze_panes,'bold':ws['A1'].font.b,'fill':ws['A1'].fill.fgColor.rgb,'filter':ws.auto_filter.ref,'dv':len(ws.data_validations.dataValidation),'cf':len(ws.conditional_formatting),'colB':round(ws.column_dimensions['B'].width,1)}))
`]).toString();
  const j = JSON.parse(py);
  check('Lưu .xlsx mở được bằng thư viện Excel khác (openpyxl)', j.sheets.join(',') === 'Dự toán,2024,Q1' && j.F2 === '=D2*E2' && j.H3 === '=ROUND(1.2345,2)' && j.merged.includes('B7:D7') && j.freeze === 'B2' && j.bold === true && /DCEFE5/i.test(j.fill) && j.filter === 'A1:G4' && j.dv >= 2 && j.cf >= 2 && Math.abs(j.colB - 30) < 0.2, py.trim());
  check('Không có lỗi JavaScript (mở/lưu tệp)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ---------------- Định dạng, gộp ô, sắp xếp, CF, CSV trên sổ mới ----------------
{
  const { ctx, page } = await newPage();
  await X(page, () => {
    const x = window.__xl, s = x.grid.sheet;
    [['Hạng mục', 'ĐVT', 'KL', 'Đơn giá'], ['Đèn LED 120W đường phố Quận 1', 'bộ', '48', '4.250.000'], ['Cột thép bát giác 10m', 'cột', '12', '6.800.000'], ['Cáp ngầm Cu/XLPE 4x16', 'm', '1.260', '118.000'], ['Ánh sáng trang trí', 'bộ', '6', '9.500.000']]
      .forEach((row, r) => row.forEach((v, c) => x.wb.setInput(s, r, c, v)));
  });
  // Viền tất cả
  await goto(page, 'A1:D5');
  await page.click('#ribbon [data-id="border"] .dd'); await menu(page, 'Tất cả Viền');
  check('Kẻ viền tất cả', await X(page, () => Object.keys(window.__xl.wb.style(window.__xl.grid.sheet, 2, 2).bd ?? {}).length === 4));
  // Xuống dòng cột A
  await goto(page, 'A2:A5');
  await btn(page, 'wr');
  const h2 = await X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).rowH.get(1));
  check('Xuống dòng và tự tăng chiều cao hàng', h2 > 24, `cao ${h2}px`);
  // Gộp ô tiêu đề
  await page.evaluate(() => window.__xl.wb.insertRows(window.__xl.grid.sheet, 0, 1));
  await page.evaluate(() => window.__xl.wb.setInput(window.__xl.grid.sheet, 0, 0, 'BẢNG DỰ TOÁN VẬT TƯ CHIẾU SÁNG'));
  await goto(page, 'A1:D1');
  await page.click('#ribbon [data-id="merge"] .dd'); await menu(page, 'Gộp & Căn giữa');
  check('Gộp và căn giữa', await X(page, () => { const x = window.__xl; const m = x.wb.mergeAt(x.grid.sheet, 0, 3); return !!m && x.wb.style(x.grid.sheet, 0, 0).al === 'c'; }));
  // Chọn ô trong vùng gộp → nhảy về ô đầu, mũi tên xuống đi ra khỏi vùng
  await goto(page, 'C1');
  const sel = await X(page, () => JSON.stringify(window.__xl.grid.sel));
  await page.keyboard.press('ArrowDown');
  const after = await X(page, () => [window.__xl.grid.sel.ar, window.__xl.grid.sel.ac]);
  check('Điều hướng qua ô gộp', sel === '{"ar":0,"ac":0,"fr":0,"fc":3}' && after[0] === 1, sel + ' → ' + after);
  // Sắp xếp theo đơn giá giảm dần (chọn ô D3 → Dữ liệu → Z→A)
  await goto(page, 'D3');
  await tab(page, 'Dữ liệu');
  await btn(page, 'd-za');
  const order = await X(page, () => [2, 3, 4, 5].map((r) => window.__xl.wb.display(window.__xl.grid.sheet, r, 0).text.slice(0, 6)));
  check('Sắp xếp Z→A theo cột (giữ tiêu đề, bỏ qua hàng gộp)', order.join('|') === 'Ánh sá|Cột th|Đèn LE|Cáp ng', order.join('|'));
  // CF: lớn hơn 1.000.000 qua hộp thoại
  await goto(page, 'D3:D6');
  await btn(page, 'd-cf'); await menu(page, 'Lớn hơn');
  await page.fill('.dlg input', '5.000.000');
  await page.screenshot({ path: OUT + 'd2-cf-dialog.png' });
  await page.click('.dlg .btn.primary');
  const cf = await X(page, () => [2, 3, 4, 5].map((r) => !!window.__xl.wb.cfStyle(window.__xl.grid.sheet, r, 3).bg));
  check('Định dạng có điều kiện "lớn hơn" qua hộp thoại', JSON.stringify(cf) === '[true,true,false,false]', JSON.stringify(cf));
  // Kiểm tra dữ liệu qua hộp thoại: số nguyên 0–100 cho cột KL
  await goto(page, 'C3:C6');
  await btn(page, 'd-v');
  await page.selectOption('.dlg select >> nth=0', 'whole');
  await page.fill('.dlg .fld:has-text("Từ") input', '0');
  await page.fill('.dlg .fld:has-text("Đến") input', '100');
  await page.click('.dlg .btn.primary');
  check('Đặt kiểm tra dữ liệu qua hộp thoại', await X(page, () => window.__xl.wb.validate(window.__xl.grid.sheet, 3, 2, '150') !== null));
  await page.screenshot({ path: OUT + 'd2-format.png' });
  // Xuất CSV
  await tab(page, 'Tệp'); await page.click('#backstage [data-bs="saveas"]');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#backstage [data-a="csv1"]')]);
  const csvPath = OUT + 'xuat.csv';
  await dl.saveAs(csvPath);
  const csv = readFileSync(csvPath, 'utf8');
  check('Xuất CSV (UTF-8 có BOM, dấu ;)', csv.charCodeAt(0) === 0xfeff && csv.includes('Ánh sáng trang trí;bộ;6;9500000'), csv.split('\n')[2]);
  // Nhập CSV
  writeFileSync(OUT + 'nhap.csv', 'Tủ,Số đèn,Công suất\nTủ 1,24,2.88\nTủ 2,30,3.6\n');
  await tab(page, 'Dữ liệu');
  const [fc] = await Promise.all([page.waitForEvent('filechooser'), btn(page, 'd-open')]);
  await fc.setFiles(OUT + 'nhap.csv');
  await page.waitForTimeout(400);
  const imp = await X(page, () => { const x = window.__xl; return [x.wb.sheetName(x.grid.sheet), x.wb.value(x.grid.sheet, 2, 2)]; });
  check('Nhập CSV thành sheet mới', imp[0] === 'nhap' && imp[1] === 3.6, JSON.stringify(imp));
  check('Không có lỗi JavaScript (định dạng & dữ liệu)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ---------------- Mobile: hộp thoại dạng ngăn kéo dưới, popover lọc ----------------
for (const [name, vp] of [['doc', { width: 390, height: 844 }], ['ngang', { width: 844, height: 390 }]]) {
  const { ctx, page } = await newPage({ viewport: vp, hasTouch: true, isMobile: true });
  await X(page, () => { const x = window.__xl, s = x.grid.sheet; [['Loại', 'SL'], ['Vật tư', '1'], ['Thiết bị', '2']].forEach((row, r) => row.forEach((v, c) => x.wb.setInput(s, r, c, v))); x.toggleFilter(); });
  await page.waitForTimeout(200);
  await page.click('#tabsel'); await page.click('.menu .mi:has-text("Dữ liệu")');
  await btn(page, 'd-sort');
  await page.waitForSelector('.dlg');
  const fits = await X(page, () => { const r = document.querySelector('.dlg').getBoundingClientRect(); return r.left >= 0 && r.right <= innerWidth && r.bottom <= innerHeight + 1; });
  await page.screenshot({ path: `${OUT}d2-m-${name}-sort.png` });
  check(`Mobile ${name}: hộp thoại nằm gọn trong màn hình`, fits);
  await page.click('.dlg .btn.primary');
  check(`Mobile ${name}: không có lỗi JavaScript`, page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

await browser.close();
try { process.kill(-srv.pid); } catch { /* đã dừng */ }
const fail = results.filter((r) => !r.ok);
console.log(`\n${results.length - fail.length}/${results.length} đạt`);
process.exit(fail.length ? 1 : 0);
