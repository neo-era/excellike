// Kiểm thử bản 0.6 – BẤM THẬT từng phím trong bảng ánh xạ (keymap.ts) trên giao diện, kiểm kết quả trong sổ.
// Báo: số phím đạt / tổng số phím trong phạm vi. Phím chỉ có ở bản Windows (Alt+F4, Ctrl+F9, Ctrl+F10) kiểm ở e2e/desktop.mjs.
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');

const PORT = 4178, APP_URL = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 5000); });

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1440, height: 860 } });
await ctx.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => { errors.push(e.message); console.log('PAGEERR', e.stack.split('\n').slice(0, 4).join(' / ')); });
page.on('dialog', (d) => d.dismiss());
await page.goto(APP_URL);
await page.waitForFunction(() => window.__xl);

const X = (f, a) => page.evaluate(f, a);
const K = (k) => page.keyboard.press(k);
const wait = (ms = 60) => page.waitForTimeout(ms);
const S = () => X(() => window.__xl.grid.sheet);
const val = (a) => X((a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return window.__xl.wb.value(window.__xl.grid.sheet, +m[2] - 1, c); }, a);
const raw = (a) => X((a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return window.__xl.wb.rawText(window.__xl.grid.sheet, +m[2] - 1, c); }, a);
const st = (a) => X((a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return window.__xl.wb.style(window.__xl.grid.sheet, +m[2] - 1, c); }, a);
const sel = () => X(() => ({ ...window.__xl.grid.sel, areas: window.__xl.grid.areas.length, range: window.__xl.grid.range }));
const nb = () => page.inputValue('#namebox');
const focusGrid = () => X(() => window.__xl.grid.focusKeys());
async function goto(a) { await page.click('#namebox'); await page.keyboard.type(a); await K('Enter'); await wait(20); }
async function reloadWait(f) {
  await Promise.all([page.waitForEvent('load'), X(f)]);
  await page.waitForFunction(() => window.__xl && window.__xl.grid);
}
async function fresh(cells = []) {
  await X(() => { document.querySelectorAll('.dlg-back, .fr-panel, .qa, .menu, .pop').forEach((e) => e.remove()); });
  await reloadWait(() => { window.__xl.newBook(); });
  await X((cells) => { const a = window.__xl, s = a.grid.sheet; for (const [ad, v] of cells) { const m = /^([A-Z]+)(\d+)$/.exec(ad); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; a.wb.setInput(s, +m[2] - 1, c, v); } a.grid.invalidate(); a.syncSelection(); }, cells);
  await goto('A1');
}
const dlgOpen = (title) => X((t) => { const h = document.querySelector('.dlg-back #dlg-t'); return !!h && (!t || h.textContent === t); }, title);
const closeDlg = () => X(() => document.querySelector('.dlg-back .dlg-x')?.click());
const menuShown = () => X(() => !!document.querySelector('.menu, .pop'));

const T = {};
// ================= Thường dùng =================
const closeKey = (key) => async () => { await fresh([['A1', 'x']]); await focusGrid(); await Promise.all([page.waitForEvent('load'), K(key)]); await page.waitForFunction(() => window.__xl && window.__xl.grid); await wait(400); const r = await X(() => ({ v: window.__xl.wb.value(window.__xl.grid.sheet, 0, 0), t: [...document.querySelectorAll('.toast')].map((e) => e.textContent).join('|') })); return [r.v === null && /Đã đóng/.test(r.t), JSON.stringify(r)]; };
T.closeBook = closeKey('Control+w');
T.closeWin = closeKey('Control+F4');
T.open = async () => { await fresh(); let called = 0; await X(() => { window.__opened = 0; const a = window.__xl; a.__o = a.openFile; a.openFile = () => { window.__opened++; }; }); await K('Control+o'); await K('Control+F12'); const n = await X(() => { const a = window.__xl; a.openFile = a.__o; return window.__opened; }); called = n; return [called === 2, 'gọi Mở ' + called + ' lần (Ctrl+O, Ctrl+F12)']; };
T.openF12 = T.open;
T.new = async () => { await fresh([['A1', 'cũ']]); await focusGrid(); const id0 = await X(() => localStorage.getItem('xl-last')); await Promise.all([page.waitForEvent('load'), K('Control+n')]); await page.waitForFunction(() => window.__xl && window.__xl.grid); const r = await X(() => [localStorage.getItem('xl-last'), window.__xl.wb.value(window.__xl.grid.sheet, 0, 0)]); return [r[0] !== id0 && r[1] === null, JSON.stringify(r)]; };
T.save = async () => { await fresh(); await X(() => { window.__sv = 0; const a = window.__xl; a.__s = a.save; a.save = () => { window.__sv++; }; }); await K('Control+s'); await K('Shift+F12'); const n = await X(() => { window.__xl.save = window.__xl.__s; return window.__sv; }); return [n === 2, 'Ctrl+S, Shift+F12 gọi Lưu ' + n + ' lần']; };
T.saveF12 = T.save;
T.saveAs = async () => { await fresh(); await X(() => { window.__sa = null; const a = window.__xl; a.__x = a.saveXlsx; a.saveXlsx = (n) => { window.__sa = n; return Promise.resolve(true); }; }); await K('F12'); const n = await X(() => { window.__xl.saveXlsx = window.__xl.__x; return window.__sa; }); return [n === true, 'F12 → Lưu thành (asNew=' + n + ')']; };
T.print = async () => { await fresh(); await K('Control+p'); await wait(250); const r = await X(() => window.__xl.backstage.isOpen() && !!document.querySelector('#backstage [data-pane="print"], #backstage .bs-print')); await K('Escape'); await wait(50); return [r]; };
T.copy = async () => { await fresh([['A1', '5']]); await goto('A1'); await K('Control+c'); const c = await X(() => !!window.__xl.wb.clip); await goto('B2'); await K('Control+v'); await wait(60); return [c && (await val('B2')) === 5, 'clip=' + c + ' B2=' + (await val('B2'))]; };
T.paste = T.copy;
T.cut = async () => { await fresh([['A1', '7']]); await goto('A1'); await K('Control+x'); await goto('C3'); await K('Control+v'); await wait(60); return [(await val('C3')) === 7 && (await val('A1')) === null]; };
T.undo = async () => { await fresh(); await goto('A1'); await page.keyboard.type('9'); await K('Enter'); await goto('A1'); await K('Control+z'); const a = await val('A1'); await K('Control+y'); const b = await val('A1'); return [a === null && b === 9, `${a} → ${b}`]; };
T.redo = async () => { await fresh([['A1', '1'], ['A2', '2']]); await goto('A1'); await K('Control+b'); await goto('A2'); await K('Control+y'); const s2 = await st('A2'); return [s2.b === true, 'Ctrl+Y không có gì làm lại → lặp lệnh cuối (Đậm) cho A2: ' + JSON.stringify(s2)]; };
T.del = async () => { await fresh([['A1', '5']]); await X(() => window.__xl.wb.setStyle(window.__xl.grid.sheet, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: true })); await goto('A1'); await K('Delete'); return [(await val('A1')) === null && (await st('A1')).b === true, 'Delete giữ định dạng']; };
T.bold = async () => { await fresh([['A1', 'x']]); await goto('A1'); await K('Control+b'); const a = (await st('A1')).b; await K('Control+2'); const b = (await st('A1')).b; return [a === true && !b, `Ctrl+B=${a}, Ctrl+2=${b}`]; };
T.italic = async () => { await fresh([['A1', 'x']]); await goto('A1'); await K('Control+i'); const a = (await st('A1')).i; await K('Control+3'); return [a === true && !(await st('A1')).i]; };
T.under = async () => { await fresh([['A1', 'x']]); await goto('A1'); await K('Control+u'); const a = (await st('A1')).u; await K('Control+4'); return [a === true && !(await st('A1')).u]; };
T.strike = async () => { await fresh([['A1', 'x']]); await goto('A1'); await K('Control+5'); return [(await st('A1')).s === true]; };
T.fillColor = async () => { await fresh(); await goto('A1'); await K('Alt'); await K('h'); await K('h'); await wait(80); const ok = await menuShown(); await K('Escape'); return [ok, 'Alt, H, H mở bảng màu tô']; };
T.centerAl = async () => { await fresh([['A1', 'x']]); await goto('A1'); await K('Alt'); await K('h'); await K('a'); await K('c'); await wait(40); return [(await st('A1')).al === 'c']; };
T.borderKt = async () => { await fresh(); await goto('A1'); await K('Alt'); await K('h'); await K('b'); await wait(80); const ok = await menuShown(); await K('Escape'); return [ok]; };
T.delColKt = async () => { await fresh([['A1', 'a'], ['B1', 'b']]); await goto('A1'); await K('Alt'); await K('h'); await K('d'); await wait(60); const items = await X(() => [...document.querySelectorAll('.menu .mi')].map((e) => e.textContent)); const it = await page.$('.menu .mi:has-text("Xóa Cột")'); if (it) await it.click(); await wait(40); return [(await val('A1')) === 'b', items.slice(0, 4).join('|')]; };
T.ctxMenu = async () => { await fresh(); await goto('B2'); await K('Shift+F10'); await wait(60); const a = await menuShown(); await K('Escape'); await focusGrid(); await K('ContextMenu'); await wait(60); const b = await menuShown(); await K('Escape'); return [a && b]; };
T.hideRows = async () => { await fresh(); await goto('A2:A3'); await K('Control+9'); const h = await X(() => [1, 2].map((r) => window.__xl.wb.isRowHidden(window.__xl.grid.sheet, r))); await goto('A1:A4'); await K('Control+Shift+Digit9'); const u = await X(() => [1, 2].map((r) => window.__xl.wb.isRowHidden(window.__xl.grid.sheet, r))); return [h.every(Boolean) && !u.some(Boolean), `ẩn ${h} → bỏ ẩn ${u}`]; };
T.unhideRows = T.hideRows;
T.hideCols = async () => { await fresh(); await goto('B1'); await K('Control+0'); const h = await X(() => window.__xl.wb.isColHidden(window.__xl.grid.sheet, 1)); await goto('A1:C1'); await K('Control+Shift+Digit0'); const u = await X(() => window.__xl.wb.isColHidden(window.__xl.grid.sheet, 1)); return [h && !u]; };
T.unhideCols = T.hideCols;

// ================= Ribbon =================
const tabKey = (key, label) => async () => { await fresh(); await focusGrid(); await K('Alt+' + key); await wait(60); await K('Escape'); await K('Escape'); const t = await page.$eval('#tabs [aria-selected="true"]', (e) => e.textContent.trim()).catch(() => ''); const bs = await X(() => window.__xl.backstage.isOpen()); if (bs) await X(() => window.__xl.backstage.close()); return [label === 'Tệp' ? bs : t === label, t + (bs ? ' (Backstage)' : '')]; };
T.tabFile = async () => { await fresh(); await focusGrid(); await K('Alt+f'); await wait(150); const bs = await X(() => window.__xl.backstage.isOpen()); await X(() => window.__xl.backstage.close()); return [bs]; };
T.tabHome = tabKey('h', 'Trang đầu'); T.tabInsert = tabKey('n', 'Chèn'); T.tabLayout = tabKey('p', 'Bố trí Trang'); T.tabFormulas = tabKey('m', 'Công thức');
T.tabData = tabKey('a', 'Dữ liệu'); T.tabReview = tabKey('r', 'Xem lại'); T.tabView = tabKey('w', 'Xem');
T.tell = async () => { await fresh(); await focusGrid(); await K('Alt+q'); const f = await X(() => document.activeElement?.id); await K('Escape'); return [f === 'tell', f]; };
T.keytips = async () => { await fresh(); await focusGrid(); await K('F10'); const a = await X(() => document.querySelectorAll('.kt').length); await K('F10'); const b = await X(() => document.querySelectorAll('.kt').length); return [a > 5 && b === 0, `F10: ${a} KeyTips → ${b}`]; };
T.ribbonArrows = async () => { await fresh(); await focusGrid(); await K('Alt'); await wait(40); await K('ArrowRight'); const t1 = await X(() => document.activeElement?.dataset?.tab); await K('ArrowRight'); const t2 = await X(() => document.activeElement?.dataset?.tab); await K('Escape'); return [!!t1 && t1 !== t2, `${t1} → ${t2}`]; };
T.ribbonTab = async () => { await fresh(); await focusGrid(); await K('Alt'); await wait(40); await K('Tab'); await K('Tab'); const a = await X(() => !!document.activeElement?.closest('#ribbon')); await K('Escape'); return [a, 'Tab vào các lệnh Ribbon']; };
T.ribbonGo = async () => { await fresh([['A1', 'x']]); await goto('A1'); await X(() => document.querySelector('#ribbon [data-id="b"]').focus()); await K('Enter'); await focusGrid(); return [(await st('A1')).b === true, 'Enter trên nút Đậm']; };
T.ribbonDrop = async () => { await fresh(); await X(() => document.querySelector('#ribbon [data-id="ins"]').focus()); await K('Alt+ArrowDown'); await wait(60); const m = await menuShown(); await K('Escape'); return [m]; };
T.ribbonGroups = async () => { await fresh(); await X(() => document.querySelector('#ribbon [data-id="cut"]').focus()); const g0 = await X(() => document.activeElement.closest('.grp')?.dataset?.g ?? document.activeElement.closest('.grp')?.querySelector('.lbl')?.textContent); await K('Control+ArrowRight'); const g1 = await X(() => document.activeElement.closest('.grp')?.querySelector('.lbl')?.textContent); await K('Escape'); return [!!g1 && g1 !== g0, `${g0} → ${g1}`]; };
T.ribbonTip = async () => { await fresh(); await X(() => document.querySelector('#ribbon [data-id="b"]').focus()); await K('Control+Shift+F10'); await wait(60); const t = await X(() => [...document.querySelectorAll('.toast')].map((e) => e.textContent).join('|')); await K('Escape'); return [/Đậm/.test(t), t]; };
T.ribbonToggle = async () => { await fresh(); await focusGrid(); const a = await X(() => document.querySelector('#app').className); await K('Control+F1'); const b = await X(() => document.querySelector('#app').className); await K('Control+F1'); return [a !== b, b]; };

// ================= Di chuyển =================
T.arrows = async () => { await fresh(); await goto('B2'); await K('ArrowDown'); await K('ArrowRight'); await K('ArrowUp'); return [(await nb()) === 'C2']; };
T.tab = async () => { await fresh(); await goto('B2'); await K('Tab'); await K('Tab'); await K('Shift+Tab'); return [(await nb()) === 'C2']; };
T.enter = async () => { await fresh(); await goto('B2'); await page.keyboard.type('1'); await K('Tab'); await page.keyboard.type('2'); await K('Tab'); await page.keyboard.type('3'); await K('Enter'); return [(await nb()) === 'B3' && (await val('D2')) === 3, 'Tab rồi Enter về cột bắt đầu: ' + (await nb())]; };
T.enterUp = async () => { await fresh(); await goto('B3'); await page.keyboard.type('1'); await K('Shift+Enter'); return [(await nb()) === 'B2']; };
T.jump = async () => { await fresh([['A1', '1'], ['A2', '2'], ['A3', '3']]); await goto('A1'); await K('Control+ArrowDown'); return [(await nb()) === 'A3']; };
T.endMode = async () => { await fresh([['A1', '1'], ['A2', '2'], ['A3', '3']]); await goto('A1'); await K('End'); const f = await X(() => document.querySelector('#st-flags')?.textContent); await K('ArrowDown'); return [(await nb()) === 'A3' && /End/.test(f), f]; };
T.lastCell = async () => { await fresh([['C5', '1'], ['A2', '2']]); await goto('A1'); await K('Control+End'); return [(await nb()) === 'C5']; };
T.lastCellExt = async () => { await fresh([['C5', '1']]); await goto('B2'); await K('Control+Shift+End'); const s = await sel(); return [s.range.r1 === 1 && s.range.r2 === 4 && s.range.c2 === 2]; };
T.home = async () => { await fresh(); await goto('D4'); await K('Home'); return [(await nb()) === 'A4']; };
T.top = async () => { await fresh(); await goto('D9'); await K('Control+Home'); return [(await nb()) === 'A1']; };
T.pgDn = async () => { await fresh(); await goto('A1'); await K('PageDown'); const a = await sel(); await K('PageUp'); const b = await sel(); return [a.ar > 10 && b.ar === 0, `${a.ar} → ${b.ar}`]; };
T.pgUp = T.pgDn;
T.pgRight = async () => { await fresh(); await goto('A1'); await K('Alt+PageDown'); const a = await sel(); await K('Alt+PageUp'); const b = await sel(); return [a.ac > 5 && b.ac === 0, `${a.ac} → ${b.ac}`]; };
T.pgLeft = T.pgRight;
T.nextSheet = async () => { await fresh(); await X(() => window.__xl.addSheet()); await X(() => window.__xl.stepSheet(-1)); const s0 = await S(); await focusGrid(); await K('Control+PageDown'); const s1 = await S(); await K('Control+PageUp'); const s2 = await S(); return [s1 !== s0 && s2 === s0]; };
T.prevSheet = T.nextSheet;
T.dvList = async () => { await fresh([['A1', 'Đèn'], ['A2', 'Cột']]); await goto('A3'); await K('Alt+ArrowDown'); await wait(60); const items = await X(() => [...document.querySelectorAll('.menu .mi')].map((e) => e.textContent.trim())); await K('Escape'); return [items.includes('Đèn') && items.includes('Cột'), items.join(',')]; };
T.shapes = async () => { await fresh([['A1', 'T'], ['B1', 'S'], ['A2', 'a'], ['B2', '1'], ['A3', 'b'], ['B3', '2']]); await X(() => { const a = window.__xl; a.wb.addChart(a.grid.sheet, { type: 'col', title: 'x', series: [{ vals: { sheet: a.grid.sheet, range: { r1: 1, c1: 1, r2: 2, c2: 1 } }, cats: { sheet: a.grid.sheet, range: { r1: 1, c1: 0, r2: 2, c2: 0 } } }], w: 300, h: 200, anchor: { r: 4, c: 3, dx: 0, dy: 0 } }); a.grid.invalidate(); }); await wait(80); await focusGrid(); await K('Control+Alt+Digit5'); const sel1 = await X(() => window.__xl.charts.sel); await K('Escape'); return [!!sel1, 'chọn biểu đồ ' + sel1]; };
T.hScroll = async () => { await fresh(); const b = await X(() => { const r = window.__xl.grid.scroller.getBoundingClientRect(); return { x: r.left + 300, y: r.top + 200 }; }); await page.mouse.move(b.x, b.y); await page.keyboard.down('Control'); await page.keyboard.down('Shift'); await page.mouse.wheel(0, 300); await page.keyboard.up('Shift'); await page.keyboard.up('Control'); await wait(80); const sl = await X(() => window.__xl.grid.scroller.scrollLeft); return [sl > 0, 'scrollLeft=' + sl]; };
T.zoomIn = async () => { await fresh(); await focusGrid(); const z0 = await X(() => window.__xl.grid.zoom); await K('Control+Alt+Equal'); const z1 = await X(() => window.__xl.grid.zoom); await K('Control+Alt+Minus'); const z2 = await X(() => window.__xl.grid.zoom); return [z1 > z0 && z2 === z0, `${z0} → ${z1} → ${z2}`]; };
T.zoomOut = T.zoomIn;

// ================= Định dạng ô =================
T.fmtCells = async () => { await fresh([['A1', '1234,5']]); await goto('A1'); await K('Control+1'); await wait(80); const open = await dlgOpen('Định dạng Ô'); await page.click('.fc [data-c="num"]'); await page.fill('#fc-dec', '3'); await page.dispatchEvent('#fc-dec', 'input'); await page.check('#fc-thou'); await page.click('.dlg-foot .primary'); const t = await X(() => window.__xl.wb.display(window.__xl.grid.sheet, 0, 0).text); return [open && t === '1.234,500', t]; };
T.fmtFont = async () => { await fresh([['A1', 'x']]); await goto('A1'); await K('Control+Shift+f'); await wait(80); const tab = await X(() => document.querySelector('.fc [role=tab][aria-selected="true"]')?.textContent); await page.selectOption('#fc-under', '2'); await page.check('#fc-sup'); await page.click('.dlg-foot .primary'); const s = await st('A1'); await K('Control+Shift+p'); await wait(60); const tab2 = await X(() => document.querySelector('.fc [role=tab][aria-selected="true"]')?.textContent); await closeDlg(); return [tab === 'Phông chữ' && tab2 === 'Phông chữ' && s.ud && s.vp === 'sup', JSON.stringify(s)]; };
T.dlgTabs = async () => { await fresh(); await goto('A1'); await K('Control+1'); await wait(60); await page.focus('.fc [role=tab][data-t="num"]'); await K('Control+Tab'); const t = await X(() => document.querySelector('.fc [role=tab][aria-selected="true"]')?.textContent); await K('Control+Shift+Tab'); const t2 = await X(() => document.querySelector('.fc [role=tab][aria-selected="true"]')?.textContent); await closeDlg(); return [t === 'Căn chỉnh' && t2 === 'Số', `${t} → ${t2}`]; };
T.edit = async () => { await fresh([['A1', 'abc']]); await goto('A1'); await K('F2'); await page.keyboard.type('d'); await K('Enter'); return [(await val('A1')) === 'abcd']; };
T.insertDlg = async () => { await fresh([['A1', '1'], ['A2', '2']]); await goto('A1'); await K('Control+Shift+Equal'); await wait(60); const o = await dlgOpen('Chèn'); await page.check('input[name=cells][value=down]'); await page.click('.dlg-foot .primary'); const v = await val('A2'); await goto('A1:XFD1'); await X(() => window.__xl.grid.selectRange({ r1: 0, c1: 0, r2: 0, c2: 16383 })); await K('Control+Shift+Equal'); const v2 = await val('A3'); return [o && v === 1 && v2 === 1, `hộp thoại Chèn; dồn xuống A2=${v}; chèn cả hàng A3=${v2}`]; };
T.deleteDlg = async () => { await fresh([['A1', '1'], ['A2', '2']]); await goto('A1'); await K('Control+Minus'); await wait(60); const o = await dlgOpen('Xóa'); await page.check('input[name=cells][value=up]'); await page.click('.dlg-foot .primary'); return [o && (await val('A1')) === 2]; };
T.today = async () => { await fresh(); await goto('A1'); await K('Control+Semicolon'); await K('Enter'); const v = await X(() => window.__xl.wb.display(window.__xl.grid.sheet, 0, 0).text); const d = new Date(); const want = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`; await goto('A2'); await page.keyboard.type('Ngày '); await K('Control+Semicolon'); await K('Enter'); const v2 = await val('A2'); return [v === want && v2 === 'Ngày ' + want, `${v} | ${v2}`]; };
T.nowTime = async () => { await fresh(); await goto('A1'); await K('Control+Shift+Semicolon'); await K('Enter'); const v = await X(() => window.__xl.wb.display(window.__xl.grid.sheet, 0, 0).text); return [/^\d{1,2}:\d{2}/.test(v), v]; };
T.showF = async () => { await fresh([['A1', '2'], ['A2', '=A1*3']]); await focusGrid(); await K('Control+Backquote'); const on = await X(() => window.__xl.grid.showFormulas); await K('Control+Backquote'); return [on === true && !(await X(() => window.__xl.grid.showFormulas))]; };
T.fAbove = async () => { await fresh([['A1', '2'], ['B1', '=A1*3']]); await goto('B2'); await K("Control+Quote"); await K('Enter'); return [(await raw('B2')) === '=A1*3', await raw('B2')]; };
T.vAbove = async () => { await fresh([['A1', '2'], ['B1', '=A1*3']]); await goto('B2'); await K('Control+Shift+Quote'); await K('Enter'); return [(await val('B2')) === 6 && !(await raw('B2')).startsWith('='), await raw('B2')]; };
T.pasteSp = async () => { await fresh([['A1', '5'], ['A2', '=A1*2']]); await goto('A1:A2'); await K('Control+c'); await goto('C1'); await K('Control+Alt+v'); await wait(80); const o = await dlgOpen('Dán Đặc biệt'); await K('v'); await K('Enter'); await wait(60); return [o && (await raw('C2')) === '10', await raw('C2')]; };
T.pasteVals = async () => { await fresh([['A1', '5'], ['A2', '=A1*2']]); await goto('A1:A2'); await K('Control+c'); await goto('D1'); await K('Control+Shift+v'); await wait(60); return [(await raw('D2')) === '10']; };
T.objects = async () => { await T.shapes(); await focusGrid(); await K('Control+6'); const a = await X(() => window.__xl.charts.el.className); await K('Control+6'); const b = await X(() => window.__xl.charts.el.className); await K('Control+6'); const c = await X(() => window.__xl.charts.el.className); return [/obj-ph/.test(a) && /obj-hide/.test(b) && !/obj-/.test(c), `${a} | ${b} | ${c}`]; };
T.outlineBd = async () => { await fresh(); await goto('B2:C3'); await K('Control+Shift+Digit7'); const a = await st('B2'), b = await st('C3'); await K('Control+Shift+Minus'); const c = await st('B2'); return [a.bd?.t && a.bd?.l && b.bd?.b && b.bd?.r && !c.bd, 'kẻ viền ngoài rồi bỏ']; };
T.noBd = T.outlineBd;
T.fillDown = async () => { await fresh([['A1', '=ROW()']]); await goto('A1:A3'); await K('Control+d'); return [(await val('A3')) === 3]; };
T.fillRight = async () => { await fresh([['A1', 'x']]); await goto('A1:C1'); await K('Control+r'); return [(await val('C1')) === 'x']; };
const nfKey = (key, test) => async () => { await fresh([['A1', '1234,5']]); await goto('A1'); await K(key); const s = await st('A1'); const t = await X(() => window.__xl.wb.display(window.__xl.grid.sheet, 0, 0).text); return [test(s.fmt, t), `${s.fmt} → ${t}`]; };
T.nfNumber = nfKey('Control+Shift+Digit1', (f, t) => t === '1.234,50');
T.nfCurrency = nfKey('Control+Shift+Digit4', (f, t) => t === '1.234,50 ₫');
T.nfPercent = nfKey('Control+Shift+Digit5', (f, t) => t === '123450%');
T.nfSci = nfKey('Control+Shift+Digit6', (f, t) => t === '1,23E+03');
T.nfDate = nfKey('Control+Shift+Digit3', (f) => f === 'd-mmm-yy');
T.nfTime = nfKey('Control+Shift+Digit2', (f) => f === 'h:mm AM/PM');
T.nfGeneral = async () => { await fresh([['A1', '0,5']]); await goto('A1'); await K('Control+Shift+Digit5'); await K('Control+Shift+Backquote'); return [!(await st('A1')).fmt]; };
T.link = async () => { await fresh([['A1', 'Sapulico']]); await goto('A1'); await K('Control+k'); await wait(60); const o = await dlgOpen(); await page.fill('#hl-addr', 'sapulico.vn'); await page.click('.dlg-foot .primary'); const s = await st('A1'); return [o && s.ln === 'http://sapulico.vn' && s.u, JSON.stringify(s)]; };
T.quick = async () => { await fresh([['A1', '1'], ['A2', '2'], ['A3', '3']]); await goto('A1:A3'); await K('Control+q'); await wait(60); const o = await X(() => !!document.querySelector('#qa')); await page.click('#qa .qa-it:has-text("Thanh Dữ liệu")'); const cf = await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).cf.map((x) => x.type)); return [o && cf.includes('bar'), cf.join(',')]; };
T.stats = async () => { await fresh([['A1', '1']]); await focusGrid(); await K('Control+Shift+g'); await wait(60); const o = await dlgOpen(); await closeDlg(); return [o]; };
T.filter = async () => { await fresh([['A1', 'T'], ['A2', '1']]); await goto('A1'); await K('Control+Shift+l'); const f = await X(() => !!window.__xl.wb.meta(window.__xl.grid.sheet).filter); await K('Control+Shift+l'); return [f && !(await X(() => !!window.__xl.wb.meta(window.__xl.grid.sheet).filter))]; };

// ================= Dán Đặc biệt (phím chữ trong hộp thoại) =================
const psKey = (letter, check) => async () => {
  await fresh([['A1', '5'], ['A2', '=A1*2']]);
  await X(() => { const a = window.__xl, s = a.grid.sheet; a.wb.setStyle(s, { r1: 0, c1: 0, r2: 1, c2: 0 }, { b: true, fmt: '0.00', bd: { t: { s: 'thin' } } }); a.wb.addDV(s, { range: { r1: 0, c1: 0, r2: 1, c2: 0 }, type: 'whole', op: 'gt', v1: '0' }); a.wb.setColWidth(s, 0, 150); });
  await goto('A1:A2'); await K('Control+c'); await goto('C5'); await K('Control+Alt+v'); await wait(60);
  await K(letter); await K('Enter'); await wait(60);
  return check();
};
T.psA = psKey('a', async () => { const s = await st('C6'); return [(await raw('C6')) === '=C5*2' && s.b === true]; });
T.psH = psKey('h', async () => [(await raw('C6')) === '=C5*2' && (await st('C6')).b === true]);
T.psF = psKey('f', async () => [(await raw('C6')) === '=C5*2' && !(await st('C6')).b]);
T.psV = psKey('v', async () => [(await raw('C6')) === '10' && !(await st('C6')).b]);
T.psT = psKey('t', async () => [(await val('C6')) === null && (await st('C6')).b === true]);
T.psC = psKey('c', async () => [(await val('C6')) === null && !(await st('C6')).b, 'Chú thích/ghi chú: không đổi nội dung (ghi chú có ở 0.7)']);
T.psN = psKey('n', async () => [await X(() => window.__xl.wb.dvAt(window.__xl.grid.sheet, 5, 2)?.type === 'whole') && (await val('C6')) === null]);
T.psX = psKey('x', async () => { const s = await st('C6'); return [s.b === true && !s.bd && (await raw('C6')) === '=C5*2']; });
T.psW = psKey('w', async () => [(await X(() => window.__xl.wb.meta(window.__xl.grid.sheet).colW.get(2))) === 150 && (await val('C6')) === null]);
T.psR = psKey('r', async () => { const s = await st('C6'); return [(await raw('C6')) === '=C5*2' && s.fmt === '0.00' && !s.b, JSON.stringify(s)]; });
T.psU = psKey('u', async () => { const s = await st('C6'); return [(await raw('C6')) === '10' && s.fmt === '0.00' && !s.b, JSON.stringify(s)]; });

// ================= Chọn & thao tác =================
T.selAll = async () => { await fresh([['B2', '1'], ['C3', '2']]); await goto('B2'); await K('Control+a'); const a = await sel(); await K('Control+a'); const b = await sel(); await goto('B2'); await K('Control+Shift+Space'); const c = await sel(); return [a.range.r2 === 2 && a.range.c2 === 2 && b.range.r2 > 1000000 && c.range.c2 === 2, `vùng ${JSON.stringify(a.range)} → cả trang`]; };
T.grpNext = async () => { await fresh(); await X(() => { window.__xl.addSheet(); window.__xl.stepSheet(-1); }); await focusGrid(); await K('Control+Shift+PageDown'); const g = await X(() => window.__xl.group.length); await goto('A1'); await page.keyboard.type('7'); await K('Enter'); const both = await X(() => window.__xl.wb.sheetIds().map((id) => window.__xl.wb.value(id, 0, 0))); return [g === 2 && both.every((v) => v === 7), 'nhập vào nhóm: ' + JSON.stringify(both)]; };
T.grpPrev = async () => { await fresh(); await X(() => { window.__xl.addSheet(); }); await focusGrid(); await K('Control+Shift+PageUp'); const g = await X(() => window.__xl.group.length); await X(() => { window.__xl.group = []; window.__xl.renderSheets(); }); return [g === 2]; };
T.extOne = async () => { await fresh(); await goto('B2'); await K('Shift+ArrowDown'); await K('Shift+ArrowRight'); const s = await sel(); return [s.range.r2 === 2 && s.range.c2 === 2]; };
T.extJump = async () => { await fresh([['A1', '1'], ['A2', '1'], ['A3', '1']]); await goto('A1'); await K('Control+Shift+ArrowDown'); const s = await sel(); return [s.range.r2 === 2]; };
T.extMode = async () => { await fresh(); await goto('B2'); await K('F8'); const f = await X(() => document.querySelector('#st-flags')?.textContent); await K('ArrowDown'); await K('ArrowDown'); await K('ArrowRight'); const s = await sel(); await K('F8'); return [s.range.r2 === 3 && s.range.c2 === 2 && /Mở rộng/.test(f), f]; };
T.addMode = async () => { await fresh(); await goto('A1'); await K('Shift+F8'); const f = await X(() => document.querySelector('#st-flags')?.textContent); const p = await X(() => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), rc = g.cellRectM(4, 3); return { x: b.left + rc.x + 5, y: b.top + rc.y + 5 }; }); await page.mouse.click(p.x, p.y); const s = await sel(); await K('Escape'); return [s.areas === 1 && /Thêm/.test(f), `${f}; ${s.areas + 1} khối`]; };
T.addClick = async () => { await fresh(); await goto('A1'); const p = await X(() => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), rc = g.cellRectM(3, 2); return { x: b.left + rc.x + 5, y: b.top + rc.y + 5 }; }); await page.keyboard.down('Control'); await page.mouse.click(p.x, p.y); await page.keyboard.up('Control'); const s = await sel(); await K('Control+b'); const b = [(await st('A1')).b, (await st('C4')).b]; return [s.areas === 1 && b[0] && b[1], 'Ctrl+nhấp: 2 khối, Đậm áp cả hai']; };
T.newLine = async () => { await fresh(); await goto('A1'); await page.keyboard.type('a'); await K('Alt+Enter'); await page.keyboard.type('b'); await K('Enter'); return [(await val('A1')) === 'a\nb']; };
T.fillSel = async () => { await fresh([['A1', '1'], ['A2', '2'], ['A3', '3']]); await goto('B1:B3'); await page.keyboard.type('=A1*10'); await K('Control+Enter'); const v = [await val('B1'), await val('B3'), await raw('B3')]; const s = await sel(); return [v[0] === 10 && v[1] === 30 && v[2] === '=A3*10' && s.range.r2 === 2, JSON.stringify(v)]; };
T.cse = async () => { await fresh([['A1', '2']]); await goto('B1'); await page.keyboard.type('=A1+1'); await K('Control+Shift+Enter'); return [(await val('B1')) === 3]; };
T.selCol = async () => { await fresh(); await goto('C3'); await K('Control+Space'); const s = await sel(); return [s.range.r1 === 0 && s.range.r2 > 1000000 && s.range.c1 === 2 && s.range.c2 === 2]; };
T.selRow = async () => { await fresh(); await goto('C3'); await K('Shift+Space'); const s = await sel(); return [s.range.c1 === 0 && s.range.c2 === 16383 && s.range.r1 === 2]; };
T.topExt = async () => { await fresh(); await goto('C3'); await K('Control+Shift+Home'); const s = await sel(); return [s.range.r1 === 0 && s.range.c1 === 0 && s.range.r2 === 2]; };
T.homeExt = async () => { await fresh(); await goto('C3'); await K('Shift+Home'); const s = await sel(); return [s.range.c1 === 0 && s.range.c2 === 2 && s.range.r1 === 2 && s.range.r2 === 2]; };
T.pgExt = async () => { await fresh(); await goto('A1'); await K('Shift+PageDown'); const s = await sel(); return [s.range.r1 === 0 && s.range.r2 > 10]; };
T.region = async () => { await fresh([['B2', '1'], ['C3', '2']]); await goto('C3'); await K('Control+Shift+Digit8'); const s = await sel(); return [s.range.r1 === 1 && s.range.c1 === 1 && s.range.r2 === 2 && s.range.c2 === 2]; };
T.repeat = async () => { await fresh([['A1', 'x'], ['A2', 'y']]); await goto('A1'); await K('Control+Shift+Digit7'); await goto('A2'); await K('F4'); return [!!(await st('A2')).bd?.t, 'F4 lặp Kẻ viền ngoài']; };
T.visible = async () => { await fresh(); await X(() => window.__xl.wb.hideRows(window.__xl.grid.sheet, 1, 1, true)); await goto('A1:A3'); await K('Alt+Semicolon'); const s = await sel(); return [s.areas === 1, (s.areas + 1) + ' khối']; };
T.prec = async () => { await fresh([['A1', '1'], ['B1', '=A1*2'], ['C1', '=B1+A1']]); await goto('C1'); await K('Control+BracketLeft'); const s = await sel(); return [s.range.c1 === 0 && s.range.c2 === 1, JSON.stringify(s.range)]; };
T.precAll = async () => { await fresh([['A1', '1'], ['B1', '=A1*2'], ['C1', '=B1'], ['D1', '=C1']]); await goto('D1'); await K('Control+Shift+BracketLeft'); const s = await sel(); return [s.range.c1 === 0 && s.range.c2 === 2]; };
T.dep = async () => { await fresh([['A1', '1'], ['B1', '=A1*2'], ['C1', '=B1']]); await goto('A1'); await K('Control+BracketRight'); const s = await sel(); return [s.range.c1 === 1 && s.range.c2 === 1]; };
T.depAll = async () => { await fresh([['A1', '1'], ['B1', '=A1*2'], ['C1', '=B1']]); await goto('A1'); await K('Control+Shift+BracketRight'); const s = await sel(); return [s.range.c1 === 1 && s.range.c2 === 2]; };
T.rowDiff = async () => { await fresh([['A1', '1'], ['B1', '1'], ['C1', '2']]); await goto('A1:C1'); await K('Control+Backslash'); const s = await sel(); return [s.range.c1 === 2 && s.range.c2 === 2 && s.areas === 0, JSON.stringify(s.range)]; };
T.colDiff = async () => { await fresh([['A1', '1'], ['A2', '1'], ['A3', '5']]); await goto('A1:A3'); await K('Control+Shift+Backslash'); const s = await sel(); return [s.range.r1 === 2 && s.range.r2 === 2, JSON.stringify(s.range)]; };
T.selNotes = async () => { await fresh([['A1', '1']]); await goto('A1:B2'); await K('Control+Shift+o'); const t = await X(() => [...document.querySelectorAll('.toast')].map((e) => e.textContent).join('|')); return [/Không tìm thấy ô nào/.test(t), 'không có ghi chú → báo như Excel']; };
T.corner = async () => { await fresh(); await goto('B2:D4'); await K('Control+Period'); const a = await sel(); await K('Control+Period'); const b = await sel(); return [a.ar === 1 && a.ac === 3 && b.ar === 3 && b.ac === 3, `(${a.ar},${a.ac}) → (${b.ar},${b.ac})`]; };
T.collapse = async () => { await fresh(); await goto('B2:D4'); await K('Shift+Backspace'); const s = await sel(); return [s.range.r1 === s.range.r2 && s.range.c1 === s.range.c2]; };
T.showActive = async () => { await fresh(); await goto('A500'); await X(() => { window.__xl.grid.scroller.scrollTop = 0; }); await focusGrid(); await K('Control+Backspace'); await wait(40); const t = await X(() => window.__xl.grid.scroller.scrollTop); return [t > 1000, 'scrollTop=' + t]; };

// ================= Dữ liệu, hàm, thanh công thức =================
T.fTips = async () => { await fresh(); await focusGrid(); await K('Control+Alt+p'); const a = await X(() => window.__xl.fTips); await K('Control+Alt+p'); return [a === false && (await X(() => window.__xl.fTips)) === true]; };
T.fxTall = async () => { await fresh(); await focusGrid(); await K('Control+Shift+u'); const a = await X(() => document.querySelector('#app').classList.contains('fx-tall')); await K('Control+Shift+u'); return [a]; };
T.cancel = async () => { await fresh([['A1', 'giữ']]); await goto('A1'); await page.keyboard.type('mới'); await K('Escape'); return [(await val('A1')) === 'giữ']; };
T.calcAll = async () => { await fresh([['A1', '=RAND()']]); await focusGrid(); const a = await val('A1'); await K('F9'); const b = await val('A1'); return [a !== b, 'F9 tính lại']; };
T.calcSheet = async () => { await fresh([['A1', '=RAND()']]); await focusGrid(); const a = await val('A1'); await K('Shift+F9'); return [a !== (await val('A1'))]; };
T.calcFull = async () => { await fresh([['A1', '=RAND()']]); await focusGrid(); const a = await val('A1'); await K('Control+Alt+F9'); return [a !== (await val('A1'))]; };
T.fArgs = async () => { await fresh([['A1', '2'], ['A2', '3']]); await goto('B1'); await page.keyboard.type('=SUM('); await K('Control+a'); await wait(80); const o = await dlgOpen('Đối số Hàm'); await page.fill('.fa input >> nth=0', 'A1:A2'); await page.dispatchEvent('.fa input >> nth=0', 'input'); const res = await page.textContent('#fa-res'); await page.click('.dlg-foot .primary'); await wait(40); await K('Enter'); return [o && (await val('B1')) === 5 && /5/.test(res), res]; };
T.fArgNames = async () => { await fresh(); await goto('B1'); await page.keyboard.type('=ROUND'); await K('Control+Shift+a'); const t = await X(() => document.querySelector('.cell-editor').value); await K('Escape'); return [t === '=ROUND(số;số_chữ_số)', t]; };
T.autoSum = async () => { await fresh([['A1', '2'], ['A2', '3']]); await goto('A3'); await K('Alt+Equal'); await K('Enter'); return [(await val('A3')) === 5]; };
T.flash = async () => { await fresh([['A1', 'Nguyễn Văn An'], ['A2', 'Trần Thị Bình'], ['A3', 'Lê Cường'], ['B1', 'An']]); await goto('B2'); await K('Control+e'); return [(await val('B2')) === 'Bình' && (await val('B3')) === 'Cường']; };
T.f4 = async () => { await fresh(); await goto('B1'); await page.keyboard.type('=A1'); await K('F4'); const a = await X(() => document.querySelector('.cell-editor').value); await K('F4'); const b = await X(() => document.querySelector('.cell-editor').value); await K('Escape'); return [a === '=$A$1' && b === '=A$1', `${a} → ${b}`]; };
T.insFn = async () => { await fresh(); await goto('A1'); await K('Shift+F3'); await wait(60); const o = await X(() => !!document.querySelector('.fn-picker')); await K('Escape'); return [o]; };
T.chartHere = async () => { await fresh([['A1', 'T'], ['B1', 'S'], ['A2', 'a'], ['B2', '1'], ['A3', 'b'], ['B3', '2']]); await goto('A1'); await K('Alt+F1'); await wait(100); const o = await X(() => !!document.querySelector('.chart-dlg, .dlg-back')); await closeDlg(); return [o, 'Alt+F1 mở chèn biểu đồ']; };
T.chartSheet = async () => { await fresh([['A1', 'T'], ['B1', 'S'], ['A2', 'a'], ['B2', '1']]); await goto('A1'); await K('F11'); await wait(100); const o = await X(() => !!document.querySelector('.dlg-back')); const btn = await page.$('.dlg-foot .primary'); if (btn) await btn.click(); await wait(100); const n = await X(() => window.__xl.wb.sheetIds().map((i) => window.__xl.wb.sheetName(i))); return [o && n.some((x) => /^Biểu đồ/.test(x)), n.join(',')]; };
T.refresh = async () => { await fresh(); await focusGrid(); await K('Control+F5'); const t = await X(() => [...document.querySelectorAll('.toast')].map((e) => e.textContent).join('|')); return [/làm mới/i.test(t), t]; };
T.refreshAll = async () => { await fresh(); await focusGrid(); await K('Control+Alt+F5'); const t = await X(() => [...document.querySelectorAll('.toast')].map((e) => e.textContent).join('|')); return [/làm mới/i.test(t), t]; };

// ================= Phím chức năng =================
T.help = async () => { await fresh(); await focusGrid(); await K('F1'); await wait(60); const n = await X(() => document.querySelectorAll('.keys-row').length); await closeDlg(); return [n > 150, n + ' dòng']; };
T.newSheet = async () => { await fresh(); await focusGrid(); const a = await X(() => window.__xl.wb.sheetIds().length); await K('Shift+F11'); await K('Alt+Shift+F1'); const b = await X(() => window.__xl.wb.sheetIds().length); return [b === a + 2]; };
T.fullScreen = async () => { await fresh(); await focusGrid(); await K('Control+Shift+F1'); const a = await X(() => getComputedStyle(document.querySelector('#ribbon')).display); await K('Control+Shift+F1'); return [a === 'none']; };
T.editBar = async () => { await fresh([['A1', 'abc']]); await goto('A1'); await K('Control+F2'); const a = await X(() => document.activeElement?.id); await K('Control+F2'); const b = await X(() => document.activeElement?.classList.contains('cell-editor')); await K('Escape'); return [a === 'fx-input' && b, `${a} → ô`]; };
T.goto = async () => { await fresh(); await focusGrid(); await K('Control+g'); await wait(50); await page.fill('#goto-ref', 'D7'); await K('Enter'); const a = await nb(); await focusGrid(); await K('F5'); await wait(50); const o = await dlgOpen('Đi tới'); await page.click('#goto-special'); await wait(50); const o2 = await dlgOpen('Đi tới Đặc biệt'); await closeDlg(); return [a === 'D7' && o && o2]; };
T.find = async () => { await fresh([['C3', 'đèn LED']]); await goto('A1'); await K('Control+f'); await wait(50); await page.fill('#fr-what', 'led'); await K('Enter'); const a = await nb(); await K('Escape'); await focusGrid(); await K('Shift+F5'); const o = await X(() => !!document.querySelector('#fr')); await K('Escape'); return [a === 'C3' && o]; };
T.replace = async () => { await fresh([['A1', 'đèn'], ['A2', 'Đèn cao áp']]); await focusGrid(); await K('Control+h'); await wait(50); await page.fill('#fr-what', 'đèn'); await page.fill('#fr-repl', 'bóng'); await page.click('#fr-repall'); const m = await page.textContent('#fr-msg'); await K('Escape'); return [(await val('A2')) === 'bóng cao áp' && /2/.test(m), m]; };
T.findNext = async () => { await fresh([['A2', 'x'], ['A5', 'x']]); await goto('A1'); await K('Control+f'); await wait(40); await page.fill('#fr-what', 'x'); await K('Enter'); await K('Escape'); await focusGrid(); await K('Shift+F4'); const a = await nb(); await K('Control+Shift+F4'); const b = await nb(); return [a === 'A5' && b === 'A2', `${a} → ${b}`]; };
T.findPrev = T.findNext;
T.pane = async () => { await fresh(); await focusGrid(); await K('F6'); const a = await X(() => !!document.activeElement?.closest('#tabs')); await K('F6'); const b = await X(() => document.activeElement?.id); await K('F6'); const c = await X(() => document.activeElement?.classList.contains('cell-editor')); return [a && b === 'zoom-slider' && c, `Ribbon → ${b} → trang tính`]; };
T.backspace = async () => { await fresh([['A1', 'abc']]); await goto('A1'); await K('Backspace'); const e = await X(() => window.__xl.ed.isEditing()); await page.keyboard.type('z'); await K('Enter'); return [e && (await val('A1')) === 'z']; };
T.painter = async () => { await fresh([['A1', 'x'], ['B3', 'y']]); await X(() => window.__xl.wb.setStyle(window.__xl.grid.sheet, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: true, bg: '#ffff00' })); await goto('A1'); await K('Alt'); await K('h'); await K('f'); await K('p'); const p = await X(() => document.querySelector('#gridwrap').classList.contains('painting')); const xy = await X(() => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), rc = g.cellRectM(2, 1); return { x: b.left + rc.x + 5, y: b.top + rc.y + 5 }; }); await page.mouse.click(xy.x, xy.y); await wait(40); const s = await st('B3'); const off = await X(() => !document.querySelector('#gridwrap').classList.contains('painting')); return [p && s.b && s.bg === '#ffff00' && off, JSON.stringify(s)]; };
T.dateEdit = async () => { await fresh([['A1', 'abcdef']]); await goto('A1'); await K('F2'); await K('Home'); await K('ArrowRight'); await K('ArrowRight'); await K('Control+Delete'); await K('Enter'); return [(await val('A1')) === 'ab']; };

// ================= Chạy =================
const keys = await X(() => window.__xlKeys.map((k) => ({ id: k.id, label: k.label, status: k.status, desc: k.desc })));
const desktopOnly = new Set(['quit', 'winMin', 'winMax']);
const results = [];
for (const k of keys.filter((x) => x.status === 'ok' && !desktopOnly.has(x.id))) {
  const f = T[k.id];
  if (!f) { results.push({ k, ok: false, info: 'CHƯA CÓ TEST' }); console.log(`✗ ${k.label} — ${k.desc} — CHƯA CÓ TEST`); continue; }
  let ok = false, info = '';
  try { [ok, info = ''] = await f(); } catch (e) { info = 'lỗi: ' + e.message.split('\n')[0]; }
  results.push({ k, ok: !!ok, info });
  console.log(`${ok ? '✓' : '✗'} ${k.label} — ${k.desc}${info ? ' — ' + info : ''}`);
}
await page.screenshot({ path: OUT + 'v06-last.png' });
const okN = results.filter((r) => r.ok).length;
const total = keys.filter((x) => x.status === 'ok').length;
console.log(`\nLỗi trang: ${errors.length ? errors.join(' | ') : 'không'}`);
console.log(`Phím trong phạm vi: ${total} (bản Windows riêng: ${[...desktopOnly].join(', ')} — kiểm ở desktop.mjs). Chưa làm/không làm: ${keys.length - total}.`);
console.log(`${okN}/${results.length} phím đạt`);
await browser.close();
try { process.kill(-srv.pid); } catch { /* đã dừng */ }
process.exit(okN === results.length && !errors.length ? 0 : 1);
