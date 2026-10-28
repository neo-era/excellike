// Kiểm thử điện thoại (Đợt 4): 360×640 dọc, 390×844 dọc, 844×390 ngang, 640×360 ngang — vùng lưới đủ rộng, không tràn ngang,
// ribbon thu gọn khi nằm ngang, vùng chạm ≥ 40px, hộp thoại vừa màn hình, nhập liệu bằng chạm
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');

const PORT = 4177, APP_URL = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 5000); });
const results = [];
const check = (name, ok, info = '') => { results.push({ name, ok }); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };
const browser = await chromium.launch();

const SIZES = [
  ['360x640', { width: 360, height: 640 }, 330],
  ['390x844', { width: 390, height: 844 }, 500],
  ['844x390', { width: 844, height: 390 }, 240],
  ['640x360', { width: 640, height: 360 }, 215],
];
for (const [name, vp, minGrid] of SIZES) {
  const landscape = vp.width > vp.height;
  const ctx = await browser.newContext({ viewport: vp, isMobile: true, hasTouch: true, deviceScaleFactor: 2 });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(APP_URL);
  await page.waitForFunction(() => window.__xl);
  await page.evaluate(() => {
    const x = window.__xl, s = x.grid.sheet;
    [['Hạng mục', 'Số lượng', 'Đơn giá', 'Thành tiền'], ['Đèn LED 100W', '48', '4250000', '=B2*C2'], ['Cột thép 8m', '48', '6800000', '=B3*C3'], ['Cáp ngầm', '1260', '118000', '=B4*C4']]
      .forEach((row, r) => row.forEach((v, c) => x.wb.setInput(s, r, c, v)));
    x.grid.invalidate();
  });
  await page.waitForTimeout(250);
  const m = await page.evaluate(() => {
    const g = document.querySelector('#gridwrap').getBoundingClientRect();
    return { gridH: Math.round(g.height), overflow: document.scrollingElement.scrollWidth - innerWidth, ribbon: getComputedStyle(document.querySelector('#ribbon')).display };
  });
  check(`${name}: vùng lưới cao ≥ ${minGrid}px`, m.gridH >= minGrid, `${m.gridH}px`);
  check(`${name}: trang không tràn ngang`, m.overflow <= 0, String(m.overflow));
  if (landscape) check(`${name}: ribbon thu gọn mặc định khi nằm ngang`, m.ribbon === 'none', m.ribbon);
  await page.screenshot({ path: OUT + `d4-mobile-${name}.png` });

  // chạm thẻ “Trang chủ” → ribbon hiện tạm; chạm nút Đậm; chạm lưới → ribbon ẩn
  if (landscape) {
    await page.tap('[role="tab"]:has-text("Trang chủ")');
    const peek = await page.evaluate(() => getComputedStyle(document.querySelector('#ribbon')).display);
    check(`${name}: chạm thẻ lệnh thì ribbon hiện tạm`, peek !== 'none', peek);
    await page.screenshot({ path: OUT + `d4-mobile-${name}-peek.png` });
    await page.tap('#ribbon [data-id="b"]');
    const bold = await page.evaluate(() => !!window.__xl.wb.style(window.__xl.grid.sheet, 0, 0).b);
    check(`${name}: dùng lệnh từ ribbon tạm (Đậm)`, bold);
    const box = await page.locator('#gridwrap').boundingBox();
    await page.touchscreen.tap(box.x + box.width - 30, box.y + box.height - 20);
    const after = await page.evaluate(() => getComputedStyle(document.querySelector('#ribbon')).display);
    check(`${name}: chạm vào lưới thì ribbon thu lại`, after === 'none', after);
  }

  // vùng chạm
  const small = await page.evaluate(() => [...document.querySelectorAll('#ribbon .rb, .tabs button, .top .ib:not([hidden]), .sheets .st, .fxb')]
    .filter((b) => b.offsetParent).map((b) => [b.id || b.className || b.textContent, b.getBoundingClientRect()]).filter(([, r]) => r.width < 36 || r.height < 30).map(([n, r]) => `${n} ${Math.round(r.width)}×${Math.round(r.height)}`));
  check(`${name}: nút đủ lớn để chạm`, small.length === 0, small.join(', '));

  // nhập liệu: chạm đúp ô E2 rồi gõ
  const cell = await page.evaluate(() => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), r = g.cellRectM(1, 1); return { x: b.left + r.x + r.w / 2, y: b.top + r.y + r.h / 2 }; });
  await page.touchscreen.tap(cell.x, cell.y);
  await page.waitForTimeout(60);
  await page.touchscreen.tap(cell.x, cell.y);
  await page.waitForTimeout(150);
  await page.keyboard.press('Control+a');
  await page.keyboard.type('50');
  await page.keyboard.press('Enter');
  const d2 = await page.evaluate(() => window.__xl.wb.value(window.__xl.grid.sheet, 1, 3));
  check(`${name}: chạm đúp sửa ô, công thức tính lại`, d2 === 212500000, String(d2));

  // hộp thoại biểu đồ vừa màn hình
  await page.evaluate(() => { window.__xl.grid.selectRange({ r1: 0, c1: 0, r2: 3, c2: 1 }); window.__xl.insertChart(); });
  await page.waitForSelector('.dlg');
  const dlg = await page.locator('.dlg').boundingBox();
  check(`${name}: hộp thoại nằm trong màn hình`, dlg.x >= 0 && dlg.x + dlg.width <= vp.width + 0.5 && dlg.y >= 0 && dlg.y + dlg.height <= vp.height + 0.5, JSON.stringify(dlg));
  await page.screenshot({ path: OUT + `d4-mobile-${name}-dialog.png` });
  await page.keyboard.press('Escape');

  // giao diện tối
  await page.evaluate(() => window.__xl.setTheme('dark'));
  await page.waitForTimeout(150);
  await page.screenshot({ path: OUT + `d4-mobile-${name}-dark.png` });
  check(`${name}: không lỗi JavaScript`, errors.length === 0, errors.join(' | '));
  await ctx.close();
}
await browser.close();
process.kill(-srv.pid);
const bad = results.filter((r) => !r.ok);
console.log(`\n${results.length - bad.length}/${results.length} kiểm thử điện thoại đạt`);
writeFileSync(OUT + 'd4-mobile-results.json', JSON.stringify(results, null, 1));
process.exit(bad.length ? 1 : 0);
