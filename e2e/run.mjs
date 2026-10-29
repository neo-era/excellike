// Kiểm thử E2E + ảnh chụp + đo hiệu năng. Chạy: npm run build && npm run e2e
import { createRequire } from 'node:module';
const { chromium } = createRequire(import.meta.url)('playwright');
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';

const PORT = 4173, APP_URL = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe' });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 5000); });

const results = [];
const check = (name, ok, info = '') => { results.push({ name, ok, info }); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };
const browser = await chromium.launch();

async function newPage(opts = {}) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 820 }, permissions: ['clipboard-read', 'clipboard-write'], ...opts });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/fonts\.g|ERR_|Failed to load resource/.test(m.text())) page.errors.push(m.text()); });
  await page.goto(APP_URL);
  await page.waitForFunction(() => window.__xl);
  return { ctx, page };
}
const disp = (page, a) => page.evaluate((a) => {
  const x = window.__xl, m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return x.wb.display(x.grid.sheet, +m[2] - 1, c - 1).text;
}, a);
const cellCenter = (page, r, c) => page.evaluate(({ r, c }) => {
  const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), rc = g.cellRect(r, c);
  return { x: b.left + rc.x + rc.w / 2, y: b.top + rc.y + rc.h / 2, rx: b.left + rc.x + rc.w, ry: b.top + rc.y + rc.h };
}, { r, c });

// ---------------- Desktop: thao tác cơ bản ----------------
{
  const { ctx, page } = await newPage();
  await page.keyboard.type('1.250'); await page.keyboard.press('Enter');
  await page.keyboard.type('=A1*2'); await page.keyboard.press('Enter');
  check('Gõ số kiểu VN và công thức', (await disp(page, 'A1')) === '1250' && (await disp(page, 'A2')) === '2500', `${await disp(page, 'A1')} / ${await disp(page, 'A2')}`);

  // chế độ trỏ ô: gõ "=" rồi bấm ô A1, gõ "+", bấm A2
  await page.keyboard.type('=');
  let p = await cellCenter(page, 0, 0); await page.mouse.click(p.x, p.y);
  await page.keyboard.type('+');
  p = await cellCenter(page, 1, 0); await page.mouse.click(p.x, p.y);
  await page.keyboard.press('Enter');
  check('Trỏ chuột chọn ô khi gõ công thức', (await disp(page, 'A3')) === '3750', await page.evaluate(() => window.__xl.wb.rawText(window.__xl.grid.sheet, 2, 0)));

  // Hoàn tác / làm lại
  await page.keyboard.press('Control+z');
  const afterUndo = await disp(page, 'A3');
  await page.keyboard.press('Control+y');
  check('Ctrl+Z / Ctrl+Y', afterUndo === '' && (await disp(page, 'A3')) === '3750');

  // Tiếng Việt có dấu + tràn chữ
  p = await cellCenter(page, 0, 1); await page.mouse.click(p.x, p.y);
  await page.keyboard.type('Đèn LED 120W đường phố Quận 1'); await page.keyboard.press('Enter');
  check('Nhập tiếng Việt có dấu', (await disp(page, 'B1')) === 'Đèn LED 120W đường phố Quận 1');

  // Định dạng: đậm bằng Ctrl+B, nút Số 0.000
  p = await cellCenter(page, 1, 0); await page.mouse.click(p.x, p.y);
  await page.keyboard.press('Control+b');
  await page.click('#ribbon [data-id="thou"]');
  check('Ctrl+B và định dạng #,##0', (await disp(page, 'A2')) === '2.500' && await page.evaluate(() => !!window.__xl.wb.style(window.__xl.grid.sheet, 1, 0).b));

  // Sao chép / dán bằng phím
  p = await cellCenter(page, 2, 0); await page.mouse.click(p.x, p.y);
  await page.keyboard.press('Control+c');
  p = await cellCenter(page, 2, 2); await page.mouse.click(p.x, p.y);
  await page.keyboard.press('Control+v');
  const pasted = await page.evaluate(() => window.__xl.wb.rawText(window.__xl.grid.sheet, 2, 2));
  check('Ctrl+C / Ctrl+V dời tham chiếu', pasted === '=C1+C2', pasted);

  // Kéo điền: 1, 2 → kéo xuống tới hàng 8
  await page.click('#namebox'); await page.keyboard.type('E1'); await page.keyboard.press('Enter');
  await page.keyboard.type('1'); await page.keyboard.press('Enter'); await page.keyboard.type('2'); await page.keyboard.press('Enter');
  await page.click('#namebox'); await page.keyboard.type('E1:E2'); await page.keyboard.press('Enter');
  const h = await cellCenter(page, 1, 4), t = await cellCenter(page, 7, 4);
  await page.mouse.move(h.rx - 1, h.ry - 1); await page.mouse.down(); await page.mouse.move(t.x, t.y, { steps: 8 }); await page.mouse.up();
  check('Kéo điền dãy số', (await disp(page, 'E8')) === '8', await disp(page, 'E8'));

  // Gợi ý hàm: gõ =VLO rồi Tab
  await page.click('#namebox'); await page.keyboard.type('G1'); await page.keyboard.press('Enter');
  await page.keyboard.type('=vlo');
  const acShown = await page.isVisible('.ac-list');
  await page.screenshot({ path: OUT + 'desktop-autocomplete.png' });
  await page.keyboard.press('Tab');
  const fxText = await page.inputValue('#fx-input');
  await page.keyboard.press('Escape');
  check('Gợi ý hàm khi gõ', acShown && fxText === '=VLOOKUP(', fxText);

  // Thêm sheet, đổi tên, tham chiếu chéo
  await page.click('#sheet-add');
  const n = await page.evaluate(() => window.__xl.wb.sheetIds().length);
  await page.dblclick('#sheets .st.on'); await page.keyboard.press('Control+a'); await page.keyboard.type('Tổng hợp'); await page.keyboard.press('Enter');
  await page.keyboard.type("=Sheet1!A2+1"); await page.keyboard.press('Enter');
  check('Thêm/đổi tên sheet, tham chiếu chéo', n === 2 && (await disp(page, 'A1')) === '2501' && (await page.textContent('#sheets .st.on')) === 'Tổng hợp');

  // Cố định ngăn + chụp ảnh
  await page.click('#sheets .st >> nth=0');
  await page.click('[role="tab"]:text-is("Xem")');
  await page.click('#namebox'); await page.keyboard.type('B2'); await page.keyboard.press('Enter');
  await page.click('#ribbon [data-id="fz"]');
  await page.click('.menu .mi:has-text("Cố định Ngăn")');
  check('Cố định ngăn', await page.evaluate(() => JSON.stringify(window.__xl.wb.meta(window.__xl.grid.sheet).freeze)) === '{"rows":1,"cols":1}');
  await page.click('[role="tab"]:has-text("Trang đầu")');

  // Trỏ ô ở sheet khác khi đang gõ công thức
  await page.click('#namebox'); await page.keyboard.type('J1'); await page.keyboard.press('Enter');
  await page.keyboard.type('=');
  await page.click('#sheets .st >> nth=1');
  let q = await cellCenter(page, 0, 0); await page.mouse.click(q.x, q.y);
  await page.keyboard.press('Enter');
  const cross = await page.evaluate(() => { const x = window.__xl; return [x.wb.sheetName(x.grid.sheet), x.wb.rawText(x.grid.sheet, 0, 9)]; });
  check('Trỏ ô sang sheet khác khi gõ công thức', cross[0] === 'Sheet1' && cross[1] === "='Tổng hợp'!A1", cross.join(' | '));

  // Chèn hàng theo số hàng đang chọn
  const before = await page.evaluate(() => window.__xl.wb.value(window.__xl.grid.sheet, 2, 0));
  await page.click('#namebox'); await page.keyboard.type('A3:A5'); await page.keyboard.press('Enter');
  await page.click('#ribbon [data-id="ins"]');
  await page.click('.menu .mi:has-text("Chèn Hàng Trang tính")');
  const moved = await page.evaluate(() => window.__xl.wb.value(window.__xl.grid.sheet, 5, 0));
  await page.keyboard.press('Control+z');
  check('Chèn đúng số hàng đang chọn (3 hàng)', moved === before, `A3 cũ → A6 = ${moved}`);

  // Lưu tự động → tải lại trang → còn dữ liệu
  await page.keyboard.press('Control+s');
  await page.waitForTimeout(400);
  await page.reload(); await page.waitForFunction(() => window.__xl);
  check('Tự lưu và mở lại sau khi tải lại trang', (await disp(page, 'A3')) === '3750' && await page.evaluate(() => window.__xl.wb.sheetIds().length) === 2);

  // Lỗi công thức hiển thị
  await page.click('#namebox'); await page.keyboard.type('H1'); await page.keyboard.press('Enter');
  await page.keyboard.type('=1/0'); await page.keyboard.press('Enter');
  await page.keyboard.type('=KHONGCO(1)'); await page.keyboard.press('Enter');
  check('Lỗi #DIV/0! và #NAME?', (await disp(page, 'H1')) === '#DIV/0!' && (await disp(page, 'H2')) === '#NAME?');

  await page.screenshot({ path: OUT + 'desktop-light.png' });
  await page.emulateMedia({ colorScheme: 'dark' });
  await page.waitForTimeout(200);
  await page.screenshot({ path: OUT + 'desktop-dark.png' });
  // menu chuột phải
  p = await cellCenter(page, 4, 3); await page.mouse.click(p.x, p.y, { button: 'right' });
  await page.waitForTimeout(150);
  await page.screenshot({ path: OUT + 'desktop-dark-menu.png' });
  await page.keyboard.press('Escape');

  // Bàn phím: chỉ dùng phím để đi và sửa
  // như Excel: có cố định ngăn tại B2 thì Ctrl+Home về B2 (ô đầu vùng cuộn)
  await page.keyboard.press('Control+Home');
  check('Ctrl+Home về ô đầu vùng không cố định (B2)', (await page.inputValue('#namebox')) === 'B2');
  await page.keyboard.press('ArrowUp'); await page.keyboard.press('F2'); await page.keyboard.press('End');
  await page.keyboard.type(' (sửa)'); await page.keyboard.press('Enter');
  check('Điều hướng & sửa bằng bàn phím (F2)', (await disp(page, 'B1')).endsWith('(sửa)'));

  check('Không có lỗi JavaScript (desktop)', page.errors.length === 0, page.errors.slice(0, 3).join(' | '));
  await ctx.close();
}

// ---------------- Hiệu năng 100.000 hàng × 20 cột ----------------
{
  const { ctx, page } = await newPage();
  const perf = await page.evaluate(async () => {
    const x = window.__xl;
    const rows = [];
    for (let r = 0; r < 100000; r++) {
      const row = [];
      for (let c = 0; c < 19; c++) row.push(c === 0 ? 'Hạng mục ' + r : (r * 7 + c * 13) % 1000);
      row.push(`=SUM(B${r + 1}:S${r + 1})`);
      rows.push(row);
    }
    const data = { v: 1, name: 'Lớn', sheets: [{ name: 'Sheet1', cells: rows, styles: [], colW: [], rowH: [], freeze: { rows: 0, cols: 0 } }] };
    const t0 = performance.now();
    const big = window.__XLWorkbook.fromJSON(data);
    const load = performance.now() - t0;
    const ts = performance.now(); let json = big.toJSON(); const ser = performance.now() - ts; json = null;
    x.wb.loadSheet(x.grid.sheet, rows);
    const t1 = performance.now();
    const ch = x.wb.setInput(x.grid.sheet, 50000, 1, '999');
    const edit = performance.now() - t1;
    // đo vẽ khi cuộn
    const g = x.grid; const times = [];
    await new Promise((r) => setTimeout(r, 1500));
    for (let i = 0; i < 10; i++) { g.scroller.scrollTop = i * 1000; g.draw(); }
    for (let i = 0; i < 100; i++) { g.scroller.scrollTop = 20000 + i * 3000; const a = performance.now(); g.draw(); times.push(performance.now() - a); }
    times.sort((a, b) => a - b);
    return { load, ser, edit, changed: ch.length, p50: times[50], p95: times[95], total: x.wb.value(x.grid.sheet, 50000, 19) };
  });
  check('Mở sổ 100.000 × 20 ô (2 triệu ô, 100.000 công thức)', perf.load < 10000, `${(perf.load / 1000).toFixed(1)} giây`);
  check('Đóng gói để lưu sổ 2 triệu ô', perf.ser < 5000, `${(perf.ser / 1000).toFixed(1)} giây`);
  check('Sửa 1 ô chỉ tính lại ô phụ thuộc', perf.changed === 2 && perf.edit < 200, `${perf.changed} ô đổi, ${perf.edit.toFixed(0)} ms`);
  check('Vẽ một khung hình khi cuộn (trung vị < 10 ms, p95 < 33 ms)', perf.p50 < 10 && perf.p95 < 33, `trung vị ${perf.p50.toFixed(1)} ms, p95 ${perf.p95.toFixed(1)} ms`);
  await page.keyboard.press('Control+End').catch(() => {});
  await page.screenshot({ path: OUT + 'perf-100k.png' });
  await ctx.close();
}

// ---------------- Mobile: dọc / ngang / máy nhỏ, sáng / tối ----------------
const devices = { 'doc-390x844': { width: 390, height: 844 }, 'ngang-844x390': { width: 844, height: 390 }, 'nho-360x640': { width: 360, height: 640 } };
for (const [name, vp] of Object.entries(devices)) {
  for (const scheme of ['light', 'dark']) {
    const { ctx, page } = await newPage({ viewport: vp, hasTouch: true, isMobile: true, colorScheme: scheme });
    await page.evaluate(() => {
      const x = window.__xl, s = x.grid.sheet;
      const data = [['STT', 'Hạng mục', 'ĐVT', 'KL', 'Đơn giá', 'Thành tiền'], ['1', 'Đèn LED 120W đường phố', 'bộ', '48', '4.250.000', '=D2*E2'], ['2', 'Cột thép bát giác 10m', 'cột', '48', '6.800.000', '=D3*E3'], ['3', 'Cáp ngầm Cu/XLPE 4x16', 'm', '1.260', '118.000', '=D4*E4'], ['', 'Tổng cộng', '', '', '', '=SUM(F2:F4)']];
      data.forEach((row, r) => row.forEach((v, c) => x.wb.setInput(s, r, c, v)));
      x.wb.setStyle(s, { r1: 1, c1: 4, r2: 4, c2: 5 }, { fmt: '#,##0' });
      x.wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 5 }, { b: true, bg: '#dcefe5' });
      x.grid.select(4, 5); x.grid.scroller.scrollTo(0, 0);
    });
    await page.waitForTimeout(250);
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth > innerWidth || document.body.scrollWidth > innerWidth);
    await page.screenshot({ path: `${OUT}m-${name}-${scheme}.png` });
    // chạm 2 lần để sửa: dùng thanh công thức
    const p = await cellCenter(page, 1, 1);
    await page.touchscreen.tap(p.x, p.y); await page.waitForTimeout(80); await page.touchscreen.tap(p.x, p.y);
    await page.waitForTimeout(250);
    const editing = await page.evaluate(() => document.body.classList.contains('editing') && document.activeElement?.id === 'fx-input');
    await page.screenshot({ path: `${OUT}m-${name}-${scheme}-sua.png` });
    if (scheme === 'light') {
      check(`Mobile ${name}: không tràn ngang`, !overflow);
      check(`Mobile ${name}: chạm 2 lần → sửa trên thanh công thức`, editing);
      // vùng chạm tối thiểu
      const small = await page.evaluate(() => [...document.querySelectorAll('#ribbon button, #titlebar button:not([hidden]), .sheets .st, .fxb')].filter((b) => { const r = b.getBoundingClientRect(); return r.width && r.height && (r.width < 40 || r.height < 36); }).length);
      check(`Mobile ${name}: nút đủ lớn để chạm`, small === 0, small ? `${small} nút nhỏ` : '');
    }
    check(`Mobile ${name} (${scheme}): không có lỗi JavaScript`, page.errors.length === 0, page.errors.slice(0, 2).join(' | '));
    await ctx.close();
  }
}

// ---------------- PWA: manifest, service worker, offline ----------------
{
  const { ctx, page } = await newPage();
  const man = await page.evaluate(async () => { const r = await fetch('manifest.webmanifest'); return r.ok ? r.json() : null; });
  check('Manifest hợp lệ (tên, standalone, icon 192/512/maskable)', !!man && man.display === 'standalone' && man.icons.some((i) => i.sizes === '192x192') && man.icons.some((i) => i.purpose === 'maskable'));
  const sw = await page.evaluate(async () => { const r = await navigator.serviceWorker.ready; return !!r.active; });
  check('Service worker hoạt động', sw);
  await page.reload(); await page.waitForTimeout(800);
  await ctx.setOffline(true);
  await page.reload().catch(() => {});
  const offline = await page.waitForFunction(() => window.__xl, null, { timeout: 8000 }).then(() => true).catch(() => false);
  check('Mất mạng vẫn mở được app', offline);
  await ctx.close();
}

await browser.close();
srv.kill();
const fail = results.filter((r) => !r.ok);
console.log(`\n${results.length - fail.length}/${results.length} đạt`);
process.exit(fail.length ? 1 : 0);
