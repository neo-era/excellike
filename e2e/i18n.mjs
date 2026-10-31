// Kiểm thử 0.6.1: ribbon/Backstage như ảnh Excel 365 của anh + chọn ngôn ngữ (Tiếng Việt / English) trong Tùy chọn.
// Phần English: duyệt mọi thẻ ribbon, các trang Backstage và vài hộp thoại, liệt kê chữ còn tiếng Việt (ngoài dữ liệu người dùng).
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');

const PORT = 4179, APP_URL = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 5000); });
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1920, height: 1000 } });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => { errors.push(e.message); console.log('PAGEERR', e.message); });
page.on('dialog', (d) => d.dismiss());
await page.goto(APP_URL);
await page.waitForFunction(() => window.__xl && window.__xl.grid);
const X = (f, a) => page.evaluate(f, a);
const wait = (ms = 80) => page.waitForTimeout(ms);
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) pass++; else { fail++; console.log('✗', m); } };
const has = (sel) => X((s) => !!document.querySelector(s), sel);
const tab = async (id) => { await page.click(`#tabs [data-tab="${id}"]`); await wait(); };
async function reloadWait(f) { await Promise.all([page.waitForEvent('load'), X(f)]); await page.waitForFunction(() => window.__xl && window.__xl.grid); }

// ---------- 1. Tiếng Việt: các nhóm/lệnh theo ảnh ----------
const want = {
  home: ['cstyle', 'cf', 'fmtTable', 'addins'],
  insert: ['forms', 'model3d', 'smartart', 'screenshot', 'i-check', 'maps'],
  layout: ['ar-align', 'ar-group', 'ar-rotate'],
  formulas: ['f-recent', 'f-more'],
  data: ['d-pic', 'd-recent', 'd-conn', 'd-qc', 'd-props', 'dt-stock', 'dt-cur', 'dt-geo', 'd-model', 'd-analyze'],
  review: ['r-perf', 'r-access', 'r-translate', 'r-changes', 'r-comment', 'r-cdel', 'r-cprev', 'r-cnext', 'r-cshow', 'r-note', 'r-unshare', 'r-ink'],
  view: ['sv-name', 'sv-keep', 'v-custom', 'v-nav', 'v-ruler', 'v-dticons', 'v-focus', 'w-side', 'w-sync', 'w-reset', 'w-switch'],
  automate: ['au-new', 'au-all', 'au-flow'],
};
for (const [t, ids] of Object.entries(want)) {
  ok(await has(`#tabs [data-tab="${t}"]`), `thẻ ${t}`);
  await tab(t);
  for (const id of ids) ok(await has(`#ribbon [data-id="${id}"]`), `${t}: thiếu ${id}`);
  await page.screenshot({ path: OUT + `i18n-vi-${t}.png`, clip: { x: 0, y: 0, width: 1920, height: 170 } });
}
await tab('formulas');
ok(!(await has('#ribbon [data-id="iter"]')), 'Công thức: bỏ nút Tính lặp');
const tabOrder = await X(() => [...document.querySelectorAll('#tabs [data-tab]')].map((b) => b.dataset.tab).join(','));
ok(tabOrder.endsWith('view,automate,help'), 'thứ tự thẻ: … Xem, Tự động hóa, Trợ giúp — ' + tabOrder);
ok(await has('#btn-comments') && await has('#btn-share'), 'nút Chú thích + Chia sẻ góc phải');
ok((await X(() => document.querySelector('#btn-comments').textContent)).includes('Chú thích'), 'nút Chú thích ghi đúng chữ');
ok(await has('#st-acc'), 'thanh trạng thái có Trợ năng');

// Bộ sưu tập Kiểu Ô: 6 ô + nút mở rộng, bấm “Tốt” tô ô
await tab('home');
ok((await X(() => document.querySelectorAll('#ribbon [data-id="cstyle"] .gal-it').length)) === 6, 'bộ sưu tập Kiểu Ô có 6 ô');
await page.click('#ribbon [data-id="cstyle"] .gal-it:nth-child(2)');
await wait();
ok((await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 0, 0).bg)) === '#C6EFCE', 'Kiểu Ô “Tốt” áp vào A1');
await page.click('#ribbon [data-id="cstyle"] .gal-more');
await wait();
ok((await X(() => document.querySelectorAll('.gal-pop .gal-it').length)) >= 13, 'mở rộng bộ sưu tập Kiểu Ô');
await page.screenshot({ path: OUT + 'i18n-vi-styles-pop.png', clip: { x: 0, y: 0, width: 1920, height: 420 } });
await page.click('.gal-pop .gal-it:has-text("Đầu đề 1")'); await wait();
ok((await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 0, 0).b)) === true, 'Kiểu Đầu đề 1 đậm');
await page.click('#ribbon [data-id="cstyle"] .gal-it:nth-child(2)'); await wait();
const st2 = await X(() => window.__xl.wb.style(window.__xl.grid.sheet, 0, 0));
ok(!st2.b && !st2.fs && !st2.bd?.b && st2.bg === '#C6EFCE', 'Kiểu Ô thay hẳn kiểu cũ (không cộng dồn) ' + JSON.stringify(st2));

// Ô Tiêu điểm
await tab('view');
await page.click('#ribbon [data-id="v-focus"]');
ok(await X(() => window.__xl.grid.focusCell), 'Ô Tiêu điểm bật');
await page.click('#ribbon [data-id="v-focus"]');
ok(!(await X(() => window.__xl.grid.focusCell)), 'Ô Tiêu điểm tắt');
// Dẫn hướng
await page.click('#ribbon [data-id="v-nav"]'); await wait();
ok(await has('#nav-pane .nv-list button'), 'ngăn Dẫn hướng liệt kê trang tính');
await page.click('#ribbon [data-id="v-nav"]'); await wait();
ok(!(await has('#nav-pane')), 'đóng ngăn Dẫn hướng');
// Dùng Gần đây ghi nhớ hàm vừa chèn
await X(() => window.__xl.insertFn('SUMPRODUCT'));
await page.keyboard.press('Escape');
ok((await X(() => window.__xl.recentFns()[0])) === 'SUMPRODUCT', 'Dùng Gần đây nhớ hàm vừa chèn');
// Kiểm tra Hiệu suất
await X(() => window.__xl.wb.setStyle(window.__xl.grid.sheet, { r1: 50, c1: 10, r2: 50, c2: 10 }, { bg: '#FF0000' }));
await tab('review');
await page.click('#ribbon [data-id="r-perf"]'); await wait();
ok((await X(() => [...document.querySelectorAll('.toast')].map((t) => t.textContent).join('|'))).includes('ô có định dạng nằm ngoài'), 'Kiểm tra Hiệu suất phát hiện ô định dạng thừa');

// Backstage Trang đầu
await page.click('#tabs [data-tab="file"]'); await wait(150);
ok(await has('[data-a="tpl-du-toan"]') && await has('[data-a="more-tpl"]'), 'Trang đầu: mẫu + Thêm mẫu');
ok(await has('.bs-pills [data-f="fav"]') && await has('#bs-find'), 'Trang đầu: Gần đây/Ưa thích/Được chia sẻ + Tìm tệp');
ok(await has('.bs-foot [data-a="recover"]'), 'Khôi phục sổ làm việc chưa lưu');
ok(await has('.bs-nav [data-bs="share"]'), 'thanh trái có Chia sẻ');
ok(/Vừa xong|phút trước|Hôm nay/.test(await X(() => document.querySelector('.bs-recent .when')?.textContent ?? '')), 'Ngày sửa đổi dạng tương đối');
await page.screenshot({ path: OUT + 'i18n-vi-backstage.png' });
await page.click('.bs-recent .rr:not(.rh) .fav'); await wait(150);
await page.click('.bs-pills [data-f="fav"]'); await wait(150);
ok((await X(() => document.querySelectorAll('.bs-recent .rr:not(.rh)').length)) === 1, 'Ưa thích: ghim một sổ');
await page.fill('#bs-find', 'zzzkhongco'); await wait(150);
ok(await has('.bs-empty'), 'Tìm tệp: không có kết quả');
// Tạo từ mẫu Dự toán
await reloadWait(() => document.querySelector('[data-a="tpl-du-toan"]').click());
await wait(300);
const tot = await X(() => window.__xl.wb.value(window.__xl.grid.sheet, 13, 8));
ok(typeof tot === 'number' && tot > 30000000, 'mẫu Dự toán: TỔNG CỘNG tính được — ' + tot);
ok((await X(() => window.__xl.wb.name)) === 'Dự toán xây dựng', 'mẫu đặt tên sổ');
ok(!(await X(() => window.__xl.wb.canUndo())), 'mẫu không để lại lịch sử Hoàn tác');
await page.screenshot({ path: OUT + 'i18n-vi-template.png' });

// ---------- 2. English ----------
await page.click('#tabs [data-tab="file"]'); await wait(100);
await page.click('.bs-nav [data-bs="options"]'); await wait(100);
await Promise.all([page.waitForEvent('load'), page.selectOption('#opt-lang', 'en')]);
await page.waitForFunction(() => window.__xl && window.__xl.grid);
await wait(300);
const VI = /[À-ỹĐđ]/;
const SKIP = '[data-noi18n], .t2c-prev, .ol-cols, .ol-sub select, .cell-editor, #fx-input, #namebox, .sheets .st, #doc-title, .fr-results td, .keys-row kbd, .clip-it, .fn-list b, .ac-list b, .bs-recent .rr b, .bs-recent .rr span small, .nv-name, #nav-pane .nv-list';
const scan = (root) => X(([root, SKIP]) => {
  const out = new Set();
  const r = document.querySelector(root); if (!r) return [];
  const VI = /[À-ÖØ-öø-ỹĐđ]/;
  const w = document.createTreeWalker(r, NodeFilter.SHOW_TEXT);
  for (let n = w.nextNode(); n; n = w.nextNode()) { const p = n.parentElement; if (!p || p.closest(SKIP) || !p.getClientRects().length) continue; const t = n.nodeValue.trim(); if (t && VI.test(t) && !/^AaBbCc|^Sample: /.test(t)) out.add(t); }
  r.querySelectorAll('[title],[aria-label],[placeholder]').forEach((el) => { if (el.closest(SKIP)) return; for (const a of ['title', 'aria-label', 'placeholder']) { const v = el.getAttribute(a); if (v && VI.test(v)) out.add('@' + a + ': ' + v); } });
  return [...out];
}, [root, SKIP]);
const missing = new Set();
const add = (list, where) => list.forEach((x) => missing.add(where + ' | ' + x));
add(await scan('#titlebar'), 'title'); add(await scan('#tabs'), 'tabs'); add(await scan('#status'), 'status'); add(await scan('.fbar'), 'fbar');
const tabs = await X(() => [...document.querySelectorAll('#tabs [data-tab]')].map((b) => b.dataset.tab).filter((t) => t !== 'file'));
for (const t of tabs) { await tab(t); add(await scan('#ribbon'), 'ribbon:' + t); await page.screenshot({ path: OUT + `i18n-en-${t}.png`, clip: { x: 0, y: 0, width: 1920, height: 170 } }); }
await tab('home');
ok((await X(() => document.querySelector('#tabs [data-tab="home"]').textContent.trim())) === 'Home', 'EN: thẻ Home');
ok((await X(() => document.querySelector('#tabs [data-tab="automate"]').textContent.trim())) === 'Automate', 'EN: thẻ Automate');
ok((await X(() => document.querySelector('#btn-share').textContent.trim())) === 'Share', 'EN: nút Share');
ok((await X(() => document.querySelector('#btn-comments').textContent.trim())) === 'Comments', 'EN: nút Comments');
ok((await X(() => document.querySelector('#ribbon [data-id="cstyle"] .gal-it').textContent.trim())) === 'Normal', 'EN: Kiểu Ô Normal');
for (const pg of ['home', 'new', 'open', 'info', 'saveas', 'print', 'share', 'export', 'account', 'options']) {
  await page.click('#tabs [data-tab="file"]').catch(() => {}); await wait(80);
  await page.click(`.bs-nav [data-bs="${pg}"]`); await wait(pg === 'print' ? 600 : 120);
  add(await scan('#backstage'), 'file:' + pg);
  if (pg === 'home') await page.screenshot({ path: OUT + 'i18n-en-backstage.png' });
}
await page.keyboard.press('Escape'); await wait(); await X(() => window.__xl.backstage.close());
await X(() => window.__xl.grid.focusKeys());
for (const [k, root, nm] of [['Control+1', '.dlg', 'Format Cells'], ['Control+f', '#fr', 'Find'], ['Control+g', '.dlg', 'Go To'], ['F1', '.dlg', 'Keys']]) {
  await page.keyboard.press(k); await wait(200);
  add(await scan(root), 'dlg:' + nm);
  if (nm === 'Format Cells') for (const i of [1, 2, 3, 4, 5]) { await X((i) => document.querySelectorAll('.dlg [role=tab]')[i]?.click(), i); await wait(60); add(await scan(root), 'dlg:' + nm + i); }
  await page.keyboard.press('Escape'); await wait(80);
  await X(() => document.querySelectorAll('.dlg-back, .fr-panel').forEach((e) => e.remove()));
}
// 0.7: hộp thoại mới + tab Thiết kế Bảng
await X(() => { const x = window.__xl, s = x.grid.sheet; [['Loại', 'Tiền'], ['Cát', '1'], ['Đá', '2']].forEach((row, r) => row.forEach((v, c) => x.wb.setInput(s, 200 + r, 20 + c, v))); x.grid.select(201, 20); x.grid.focusKeys(); });
for (const [k, nm] of [['Control+F3', 'Name Manager'], ['Control+t', 'Create Table'], ['Alt+Shift+ArrowRight', 'Group']]) {
  await page.keyboard.press(k); await wait(200);
  add(await scan('.dlg'), 'dlg:' + nm);
  await page.keyboard.press('Escape'); await wait(80);
  await X(() => { document.querySelectorAll('.dlg-back').forEach((e) => e.remove()); window.__xl.grid.focusKeys(); });
}
await tab('data');
for (const id of ['d-sub', 'd-t2c', 'd-dup']) {
  await X(() => { window.__xl.grid.select(201, 20); });
  if (id === 'd-t2c') await X(() => window.__xl.grid.selectRange({ r1: 200, c1: 20, r2: 202, c2: 20 }));
  await page.click(`#ribbon [data-id="${id}"]`); await wait(200);
  add(await scan('.dlg'), 'dlg:' + id);
  await page.keyboard.press('Escape'); await wait(80);
  await X(() => document.querySelectorAll('.dlg-back').forEach((e) => e.remove()));
}
await X(() => { const x = window.__xl, s = x.grid.sheet; x.wb.createTable(s, { r1: 200, c1: 20, r2: 202, c2: 21 }); x.grid.select(201, 20); x.syncSelection(); });
await wait(150);
await tab('tdesign'); add(await scan('#ribbon'), 'ribbon:tdesign'); add(await scan('#tabs'), 'tabs2');
await page.screenshot({ path: OUT + 'i18n-en-tdesign.png', clip: { x: 0, y: 0, width: 1920, height: 170 } });
// 0.8–0.9: hộp thoại, ngăn và tab ngữ cảnh mới
const dlgAt = async (t, id, nth, nm) => {
  await tab(t); await page.click(`#ribbon [data-id="${id}"]`); await wait(200);
  if (nth !== null) { await page.locator('.menu [role^=menuitem]').nth(nth).click(); await wait(250); }
  add(await scan('.dlg'), 'dlg:' + nm); add(await scan('.watch-pane'), 'pane:' + nm);
  await page.keyboard.press('Escape'); await wait(80);
  await X(() => { document.querySelectorAll('.dlg-back').forEach((e) => e.remove()); window.__xl.grid.focusKeys(); });
};
await X(() => { const x = window.__xl, s = x.grid.sheet; for (let r = 0; r < 8; r++) { x.wb.setInput(s, 300 + r, 0, String(45000 + r * 30)); x.wb.setInput(s, 300 + r, 1, String(10 + r)); } x.wb.setInput(s, 300, 2, '=A301*B301'); x.wb.setInput(s, 301, 2, '=1/0'); x.grid.selectRange({ r1: 300, c1: 0, r2: 307, c2: 1 }); x.syncSelection(); });
for (const [t, id, nth, nm] of [['data', 'd-whatif', 0, 'Scenario'], ['data', 'd-whatif', 1, 'Goal Seek'], ['data', 'd-whatif', 2, 'Data Table'], ['data', 'd-fc', null, 'Forecast'], ['data', 'd-adv', null, 'Advanced Filter'], ['data', 'd-cons', null, 'Consolidate'],
  ['insert', 'sp-line', null, 'Sparklines'], ['insert', 'symbol', null, 'Symbol'], ['insert', 'icons', null, 'Icons'], ['review', 'r-allow', null, 'Allow Ranges'], ['review', 'r-pbook', null, 'Protect Book'], ['review', 'r-psheet', null, 'Protect Sheet']]) await dlgAt(t, id, nth, nm);
await X(() => { const x = window.__xl; x.grid.select(300, 2); x.syncSelection(); });
await dlgAt('formulas', 'eval-f', null, 'Evaluate');
await X(() => { const x = window.__xl; x.grid.select(301, 2); x.syncSelection(); });
await dlgAt('formulas', 'err-chk', null, 'Error Check');
await tab('formulas'); await page.click('#ribbon [data-id="watch"]'); await wait(200); add(await scan('.watch-pane'), 'pane:watch'); await page.click('#ribbon [data-id="watch"]'); await wait(100);
await X(() => { const x = window.__xl, s = x.grid.sheet; x.wb.setInput(s, 310, 0, 'teh'); x.grid.select(310, 0); x.grid.focusKeys(); });
await page.keyboard.press('F7'); await wait(1500); add(await scan('.dlg'), 'dlg:Spelling'); await page.keyboard.press('Escape'); await wait(80);
await X(() => document.querySelectorAll('.dlg-back').forEach((e) => e.remove()));
// tab ngữ cảnh: Biểu đồ, Biểu đồ thu nhỏ, Định dạng Hình, Bố trí Trang
await X(() => { const x = window.__xl, s = x.grid.sheet; const id = x.wb.addChart(s, { type: 'col', series: [{ vals: { sheet: s, range: { r1: 300, c1: 1, r2: 307, c2: 1 } } }], anchor: { r: 295, c: 4, dx: 0, dy: 0 }, w: 360, h: 220 }); x.grid.select(295, 4); x.charts.sync(); x.charts.select(id); });
await wait(250);
for (const t of ['cdesign', 'cformat']) if (await has(`#tabs [data-tab="${t}"]`)) { await tab(t); add(await scan('#ribbon'), 'ribbon:' + t); add(await scan('#tabs'), 'tabs:' + t); } else ok(false, 'thiếu tab ' + t);
await X(() => { const x = window.__xl, s = x.grid.sheet; x.charts.select(null); const id = x.wb.addChart(s, { type: 'shape', shape: 'rect', series: [], anchor: { r: 296, c: 10, dx: 0, dy: 0 }, w: 120, h: 60 }); x.charts.sync(); x.charts.select(id); });
await wait(250);
if (await has('#tabs [data-tab="oformat"]')) { await tab('oformat'); add(await scan('#ribbon'), 'ribbon:oformat'); } else ok(false, 'thiếu tab oformat');
await X(() => { const x = window.__xl, s = x.grid.sheet; x.charts.select(null); x.wb.addSparklines(s, 'line', { sheet: s, range: { r1: 300, c1: 0, r2: 300, c2: 1 } }, { r1: 300, c1: 5, r2: 300, c2: 5 }); x.grid.select(300, 5); x.syncSelection(); });
await wait(250);
const spT = await X(() => [...document.querySelectorAll('#tabs [data-tab]')].map((b) => b.dataset.tab).find((t) => t === 'sdesign'));
if (spT) { await tab(spT); add(await scan('#ribbon'), 'ribbon:' + spT); } else ok(false, 'thiếu tab Sparkline');
await X(() => { window.__xl.grid.select(0, 0); window.__xl.syncSelection(); });
await page.click('#views [data-view="layout"]'); await wait(500);
add(await X(() => { const r = document.querySelector('.pl-view')?.shadowRoot; if (!r) return []; const VI = /[À-ÖØ-öø-ỹĐđ]/, out = []; r.querySelectorAll('.hfz [data-ph], .hfz [title], .pl-tip').forEach((e) => { for (const v of [e.getAttribute('data-ph'), e.getAttribute('title'), e.classList.contains('pl-tip') ? e.textContent : null]) if (v && VI.test(v)) out.push(v); }); return out; }), 'layout');
add(await scan('#status'), 'status:layout');
await page.click('#views [data-view="normal"]'); await wait(200);
writeFileSync(OUT + 'i18n-missing.txt', [...missing].join('\n'));
console.log(`chữ còn tiếng Việt khi chọn English: ${missing.size} (xem e2e/shots/i18n-missing.txt)`);
ok(missing.size === 0, 'English: không còn chữ tiếng Việt trên giao diện');
// quay lại tiếng Việt
await X(() => { localStorage.setItem('xl-lang', 'vi'); });
ok(errors.length === 0, 'không có lỗi trang: ' + errors.join(' | '));
console.log(`${pass}/${pass + fail} kiểm thử giao diện 0.6.1 + ngôn ngữ đạt`);
await browser.close();
try { process.kill(-srv.pid); } catch { /* đã tắt */ }
process.exit(fail ? 1 : 0);
