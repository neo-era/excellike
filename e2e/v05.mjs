// Kiểm thử trình duyệt bản 0.5 – giao diện giống Excel 365 (phương án A – Nhiều màu)
// Thanh tiêu đề + QAT + Tìm kiếm (Alt+Q), 9 thẻ, ribbon 3 tầng, KeyTips, Backstage, thanh trạng thái,
// menu chuột phải + thanh mini, thẻ sheet, chủ đề Office, điện thoại.
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
const { chromium } = createRequire(import.meta.url)('playwright');

const PORT = 4177, APP_URL = `http://localhost:${PORT}/`;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 5000); });

const results = [];
const check = (name, ok, info = '') => { results.push({ name, ok }); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };
const browser = await chromium.launch();
async function newPage(opts = {}) {
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 860 }, ...opts });
  await ctx.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  await page.goto(APP_URL);
  await page.waitForFunction(() => window.__xl);
  return { ctx, page };
}
const X = (page, f, a) => page.evaluate(f, a);
const st = (page, r, c) => X(page, ({ r, c }) => window.__xl.wb.style(window.__xl.grid.sheet, r, c), { r, c });
const val = (page, r, c) => X(page, ({ r, c }) => window.__xl.wb.value(window.__xl.grid.sheet, r, c), { r, c });
const goto = async (page, a) => { await page.click('#namebox'); await page.keyboard.type(a); await page.keyboard.press('Enter'); };
const type = async (page, a, t) => { await goto(page, a); await page.keyboard.type(t); await page.keyboard.press('Enter'); };
const pt = (page, r, c) => X(page, ({ r, c }) => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), rc = g.cellRectM(r, c); return { x: b.left + rc.x + rc.w / 2, y: b.top + rc.y + rc.h / 2 }; }, { r, c });
const hdr = (page, kind, i) => X(page, ({ kind, i }) => { const g = window.__xl.grid, b = g.scroller.getBoundingClientRect(), rc = g.cellRect(kind === 'r' ? i : 0, kind === 'c' ? i : 0); return kind === 'r' ? { x: b.left + g.hdrW / 2, y: b.top + rc.y + rc.h / 2 } : { x: b.left + rc.x + rc.w / 2, y: b.top + g.hdrH / 2 }; }, { kind, i });
const texts = (page, sel) => page.$$eval(sel, (els) => els.map((e) => e.textContent.trim()));
const visible = (page, sel) => page.$eval(sel, (e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(e).visibility !== 'hidden'; }).catch(() => false);
const ktShown = (page) => page.$$eval('.kt', (els) => els.filter((e) => e.offsetParent).map((e) => e.textContent));
const menuItems = (page) => page.$$eval('.menu .mi span', (els) => els.map((e) => e.textContent.trim()));

// ================= Thanh tiêu đề, thẻ, ribbon =================
{
  const { ctx, page } = await newPage();
  const tb = await X(page, () => { const t = document.querySelector('#titlebar'); return { bg: getComputedStyle(t).backgroundColor, h: t.getBoundingClientRect().height, font: getComputedStyle(document.body).fontFamily }; });
  check('Thanh tiêu đề xanh Excel #107C41, cao 44px', tb.bg === 'rgb(16, 124, 65)' && Math.abs(tb.h - 44) <= 1, JSON.stringify(tb));
  check('Phông giao diện Segoe UI', /^"?Segoe UI/.test(tb.font), tb.font);
  const qat = await X(page, () => ['save', 'undo', 'redo'].map((q) => !!document.querySelector(`#qat [data-q="${q}"]`)).concat([!!document.querySelector('#qat-more'), document.querySelector('#autosave')?.getAttribute('role')]));
  check('QAT: Tự động lưu (công tắc), Lưu, Hoàn tác, Làm lại, Tùy chỉnh', JSON.stringify(qat) === '[true,true,true,true,"switch"]', JSON.stringify(qat));
  check('Ô Tìm kiếm (Alt+Q) trên thanh tiêu đề', (await page.getAttribute('#tell', 'placeholder')) === 'Tìm kiếm (Alt+Q)');
  const tabs = await texts(page, '#tabs [role="tab"]');
  check('Đủ 9 thẻ đúng thứ tự Excel', tabs.join('|') === 'Tệp|Trang đầu|Chèn|Bố trí Trang|Công thức|Dữ liệu|Xem lại|Xem|Trợ giúp', tabs.join('|'));
  const sel = await page.$eval('#tabs [aria-selected="true"]', (e) => e.textContent.trim());
  check('Mặc định mở thẻ Trang đầu', sel === 'Trang đầu', sel);
  const groups = await texts(page, '#ribbon .grp > .lbl');
  check('Trang đầu: 7 nhóm như Excel', groups.join('|') === 'Bảng tạm|Phông chữ|Căn chỉnh|Số|Kiểu|Ô|Chỉnh sửa', groups.join('|'));
  const launch = await X(page, () => [...document.querySelectorAll('#ribbon .grp [data-launch]')].map((e) => e.dataset.launch));
  check('Nút ↘ mở hộp thoại ở Bảng tạm, Phông chữ, Căn chỉnh, Số', ['clipboard', 'font', 'align', 'number'].every((k) => launch.includes(k)), launch.join(','));
  const shapes = await X(page, () => ({ paste: document.querySelector('#ribbon [data-id="paste"]')?.classList.contains('big'), cut: document.querySelector('#ribbon [data-id="cut"]')?.classList.contains('sm'), b: document.querySelector('#ribbon [data-id="b"]')?.classList.contains('ic'), mergeSplit: !!document.querySelector('#ribbon [data-id="merge"] .dd'), rbH: document.querySelector('#ribbon').getBoundingClientRect().height }));
  check('Ribbon 3 tầng: nút lớn (Dán), nút nhỏ có chữ (Cắt), nút biểu tượng (Đậm), nút tách đôi (Gộp)', shapes.paste && shapes.cut && shapes.b && shapes.mergeSplit && shapes.rbH >= 88 && shapes.rbH <= 100, JSON.stringify(shapes));
  await page.screenshot({ path: OUT + 'v05-home.png' });

  // hộp phông / cỡ chữ / định dạng số
  await goto(page, 'A1');
  const fn0 = await page.inputValue('#ribbon [data-id="fn"] input');
  const fs0 = await page.inputValue('#ribbon [data-id="fs"] input');
  check('Hộp phông mặc định Aptos Narrow, cỡ 11', fn0 === 'Aptos Narrow' && fs0 === '11', fn0 + ' ' + fs0);
  await page.fill('#ribbon [data-id="fn"] input', 'Arial'); await page.keyboard.press('Enter');
  await page.fill('#ribbon [data-id="fs"] input', '14'); await page.keyboard.press('Enter');
  await page.click('#ribbon [data-id="grow"]');
  const s1 = await st(page, 0, 0);
  check('Gõ tên phông, cỡ chữ; Tăng Cỡ Phông theo bậc Excel (14 → 16)', s1.fn === 'Arial' && s1.fs === 16, JSON.stringify(s1));
  await type(page, 'B2', '0,25');
  await goto(page, 'B2');
  check('Hộp định dạng số hiện “Chung”', (await page.inputValue('#ribbon [data-id="nf"] input')) === 'Chung');
  await page.click('#ribbon [data-id="pct"]');
  check('Bấm % → hộp định dạng số hiện “Phần trăm”', (await page.inputValue('#ribbon [data-id="nf"] input')) === 'Phần trăm');
  await page.click('#ribbon [data-id="nf"] .dd');
  const nfList = await menuItems(page);
  check('Danh sách định dạng số như Excel', ['Chung', 'Số', 'Tiền tệ', 'Kế toán', 'Ngày dạng Ngắn', 'Ngày dạng Dài', 'Thời gian', 'Phần trăm', 'Phân số', 'Khoa học', 'Văn bản'].every((x) => nfList.includes(x)), nfList.join(','));
  await page.keyboard.press('Escape');

  // nút tách đôi
  await goto(page, 'C3:E3');
  await page.click('#ribbon [data-id="merge"] .main');
  const mg = await X(page, () => window.__xl.wb.mergeAt(window.__xl.grid.sheet, 2, 2));
  check('Gộp & Căn giữa (phần chính của nút tách đôi)', mg && mg.c2 === 4, JSON.stringify(mg));
  await page.click('#ribbon [data-id="merge"] .dd');
  const mItems = await menuItems(page);
  check('Menu Gộp: Gộp & Căn giữa, Gộp Ngang, Gộp Ô, Bỏ gộp Ô', ['Gộp & Căn giữa', 'Gộp Ngang', 'Gộp Ô', 'Bỏ gộp Ô'].every((x) => mItems.includes(x)), mItems.join(','));
  await page.keyboard.press('Escape');
  await type(page, 'F1', '2'); await type(page, 'F2', '4'); await type(page, 'F3', '9');
  await goto(page, 'F4');
  await page.click('#ribbon [data-id="sum"] .dd');
  await page.click('.menu .mi:has-text("Trung bình")');
  await page.keyboard.press('Enter');
  check('Tổng Tự động ▸ Trung bình', (await val(page, 3, 5)) === 5, String(await val(page, 3, 5)));
  await page.click('#ribbon [data-id="format"]');
  const fItems = await menuItems(page);
  check('Định dạng ▸ Chiều cao Hàng…, Độ rộng Cột…, Ẩn & Bỏ ẩn, Đổi tên Trang tính, Màu Tab', ['Chiều cao Hàng…', 'Độ rộng Cột…', 'Ẩn & Bỏ ẩn', 'Đổi tên Trang tính', 'Màu Tab'].every((x) => fItems.includes(x)), fItems.join(','));
  await page.keyboard.press('Escape');
  const later = await X(page, () => [...document.querySelectorAll('#ribbon .later')].map((e) => e.title));
  check('Lệnh chưa làm được đánh dấu rõ “có ở bản 0.x”', later.length > 0 && later.every((t) => /bản 0\.\d/.test(t)), later.slice(0, 3).join(' | '));

  // các thẻ khác
  const want = { 'Chèn': ['Bảng', 'Hình minh họa', 'Biểu đồ'], 'Bố trí Trang': ['Chủ đề', 'Thiết lập Trang', 'Co giãn cho Vừa', 'Tùy chọn Trang tính'], 'Công thức': ['Thư viện Hàm', 'Tên đã Xác định', 'Kiểm nghiệm Công thức', 'Tính toán'], 'Dữ liệu': ['Lấy & Chuyển đổi Dữ liệu', 'Sắp xếp & Lọc', 'Công cụ Dữ liệu', 'Dự báo', 'Đại cương'], 'Xem lại': ['Soát lỗi', 'Ghi chú', 'Bảo vệ'], 'Xem': ['Dạng xem Sổ làm việc', 'Hiện', 'Thu phóng', 'Cửa sổ'], 'Trợ giúp': ['Trợ giúp'] };
  for (const [t, gs] of Object.entries(want)) {
    await page.click(`#tabs [role="tab"]:text-is("${t}")`);
    const got = await texts(page, '#ribbon .grp > .lbl');
    check(`Thẻ ${t}: các nhóm ${gs.join(', ')}`, gs.every((g) => got.includes(g)), got.join('|'));
  }
  await page.click('#tabs [role="tab"]:text-is("Xem")');
  await page.screenshot({ path: OUT + 'v05-view.png' });
  await page.click('#tabs [role="tab"]:has-text("Trang đầu")');

  // tên tệp trên thanh tiêu đề
  await page.click('#docname');
  await page.fill('#doc-pop-name', 'Dự toán Q4');
  await page.keyboard.press('Enter');
  const nm = await X(page, () => ({ wb: window.__xl.wb.name, t: document.querySelector('#docname').textContent }));
  check('Bấm tên tệp → đổi tên sổ làm việc', nm.wb === 'Dự toán Q4' && nm.t.includes('Dự toán Q4'), JSON.stringify(nm));
  await page.click('#qat-more');
  const qItems = await menuItems(page);
  check('Tùy chỉnh QAT: Mới, Mở, In Nhanh, Sắp xếp Tăng dần…', ['Mới', 'Mở', 'Lưu', 'In Nhanh', 'Sắp xếp Tăng dần', 'Sắp xếp Giảm dần'].every((x) => qItems.includes(x)), qItems.join(','));
  await page.click('.menu .mi:has-text("Sắp xếp Tăng dần")');
  check('Thêm Sắp xếp Tăng dần vào QAT', !!(await page.$('#qat [data-q="sortAsc"]')));
  await page.reload(); await page.waitForFunction(() => window.__xl);
  check('QAT tùy chỉnh được nhớ sau khi tải lại', !!(await page.$('#qat [data-q="sortAsc"]')));
  check('Không lỗi trang (phần thanh tiêu đề/ribbon)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ================= KeyTips, Alt+Q =================
{
  const { ctx, page } = await newPage();
  await goto(page, 'A1');
  await page.keyboard.press('Alt');
  let kt = await ktShown(page);
  check('Nhấn Alt → hiện KeyTips thẻ (F, H, N, P, M, A, R, W, Y)', ['F', 'H', 'N', 'P', 'M', 'A', 'R', 'W', 'Y'].every((k) => kt.includes(k)), kt.join(','));
  await page.screenshot({ path: OUT + 'v05-keytips.png' });
  await page.keyboard.press('h');
  kt = await ktShown(page);
  check('Alt, H → KeyTips nút trên Trang đầu (1, 2, 3, V, B, AC…)', ['1', '2', '3', 'V', 'B', 'AC', 'W', 'M'].every((k) => kt.includes(k)), kt.slice(0, 20).join(','));
  await page.screenshot({ path: OUT + 'v05-keytips-h.png' });
  await page.keyboard.press('1');
  check('Alt, H, 1 → Đậm; KeyTips tắt', (await st(page, 0, 0)).b === true && (await ktShown(page)).length === 0);
  await page.keyboard.press('Alt'); await page.keyboard.press('h'); await page.keyboard.press('a'); await page.keyboard.press('c');
  check('Alt, H, A, C → Căn giữa', (await st(page, 0, 0)).al === 'c', JSON.stringify(await st(page, 0, 0)));
  await page.keyboard.press('Alt'); await page.keyboard.press('Escape');
  check('Esc tắt KeyTips', (await ktShown(page)).length === 0);
  await page.keyboard.press('F10');
  check('F10 cũng hiện KeyTips', (await ktShown(page)).includes('H'));
  await page.keyboard.press('n');
  check('Alt, N → sang thẻ Chèn', (await page.$eval('#tabs [aria-selected="true"]', (e) => e.textContent.trim())) === 'Chèn');
  await page.keyboard.press('Escape'); await page.keyboard.press('Escape');
  await page.keyboard.press('Alt+h');
  kt = await ktShown(page);
  check('Alt+H (nhấn giữ) → vào thẳng Trang đầu với KeyTips', (await page.$eval('#tabs [aria-selected="true"]', (e) => e.textContent.trim())) === 'Trang đầu' && kt.includes('1'));
  await page.keyboard.press('Escape'); await page.keyboard.press('Escape');
  await page.keyboard.press('Alt'); await page.keyboard.press('h'); await page.keyboard.press('b');
  check('Alt, H, B → mở menu Viền', (await menuItems(page)).some((x) => /Viền Dưới/.test(x)));
  await page.keyboard.press('Escape');
  // Alt+Q tìm lệnh
  await goto(page, 'B2:D2');
  await page.keyboard.press('Alt+q');
  check('Alt+Q đưa con trỏ vào ô Tìm kiếm', await X(page, () => document.activeElement?.id === 'tell'));
  await page.keyboard.type('gộp');
  const opts = await texts(page, '.tell-list [role="option"]');
  check('Tìm “gộp” ra lệnh Gộp & Căn giữa', opts.some((o) => o.includes('Gộp & Căn giữa')), opts.join(' | '));
  await page.keyboard.press('Enter');
  const mg = await X(page, () => window.__xl.wb.mergeAt(window.__xl.grid.sheet, 1, 1));
  check('Enter chạy lệnh tìm được', !!mg && mg.c2 === 3, JSON.stringify(mg));
  await page.keyboard.press('Alt+q'); await page.keyboard.type('dam');
  const o2 = await texts(page, '.tell-list [role="option"]');
  check('Tìm không dấu “dam” vẫn ra Đậm', o2.some((o) => o.startsWith('Đậm')), o2.join(' | '));
  await page.keyboard.press('Escape');
  // Alt+= và Alt+Enter không bị KeyTips chiếm
  await type(page, 'H1', '5'); await type(page, 'H2', '6'); await goto(page, 'H3');
  await page.keyboard.press('Alt+=');
  await page.keyboard.press('Enter');
  check('Alt+= vẫn là Tổng Tự động', (await val(page, 2, 7)) === 11 && (await ktShown(page)).length === 0);
  const clash = await X(page, () => {
    const bad = [];
    const tabs = window.__xl.tabs();
    const sets = [tabs.map((t) => t.kt).concat(['Q', '1', '2', '3', '4', '5', '6', '7', '8', '9'])];
    const walk = (items, out) => { for (const it of items) { if (it.items) walk(it.items, out); else if (it.kt) out.push(it.kt); } return out; };
    for (const t of tabs) sets.push(walk(t.groups.flatMap((g) => g.items), t.groups.filter((g) => g.launch?.kt).map((g) => g.launch.kt)).map((k) => k));
    sets.forEach((ks, i) => ks.forEach((a, x) => ks.forEach((b, y) => { if (x !== y && b.startsWith(a)) bad.push(`${i}:${a}/${b}`); })));
    return bad;
  });
  check('KeyTips mỗi thẻ không trùng, không là tiền tố của nhau', clash.length === 0, clash.join(' '));
  check('Không lỗi trang (KeyTips)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ================= Backstage =================
{
  const { ctx, page } = await newPage();
  await type(page, 'A1', 'Hạng mục'); await type(page, 'A2', 'Đèn LED 100W');
  await page.click('#tabs [role="tab"]:has-text("Tệp")');
  check('Tệp mở Backstage toàn màn hình', await visible(page, '#backstage'));
  const nav = await X(page, () => [...document.querySelectorAll('#backstage [data-bs]')].map((e) => e.dataset.bs));
  check('Backstage: Trang đầu, Mới, Mở, Thông tin, Lưu, Lưu thành, In, Xuất, Đóng, Tài khoản, Tùy chọn', ['home', 'new', 'open', 'info', 'save', 'saveas', 'print', 'export', 'close', 'account', 'options'].every((k) => nav.includes(k)), nav.join(','));
  await page.screenshot({ path: OUT + 'v05-backstage.png' });
  await page.click('#backstage [data-bs="info"]');
  check('Thông tin: số trang tính, tên sổ', /Trang tính\s*1/.test(await page.textContent('#backstage .bs-page')));
  await page.click('#backstage [data-bs="open"]');
  await page.waitForSelector('#backstage .bs-recent [data-book]', { timeout: 3000 }).catch(() => {});
  check('Mở: danh sách gần đây + nút Duyệt', (await page.$$('#backstage .bs-recent [data-book]')).length >= 1 && !!(await page.$('#backstage [data-a="browse"]')));
  check('Không lỗi trang (Backstage – phần đầu)', page.errors.length === 0, page.errors.join(' | '));
  await page.keyboard.press('Escape');
  await ctx.close();
}

// ================= In như Excel 365 =================
{
  const { ctx, page } = await newPage();
  // 60 hàng × 12 cột (ô 64×20) → A4 dọc 100%: 4 trang
  await X(page, () => { const x = window.__xl, d = []; for (let r = 0; r < 60; r++) { const row = []; for (let c = 0; c < 12; c++) row.push(r * 100 + c); d.push(row); } x.wb.loadSheet(x.grid.sheet, d); x.grid.invalidate(); });
  await page.keyboard.press('Control+p');
  await page.waitForSelector('#backstage .bs-preview iframe');
  const ctl = await X(page, () => ['bs-copies', 'bs-printer', 'bs-what', 'bs-from', 'bs-to', 'bs-collate', 'bs-orient', 'bs-paper', 'bs-margins', 'bs-scale', 'bs-setup', 'bs-print', 'pv-prev', 'pv-next', 'pv-num', 'pv-margins', 'pv-zoom'].filter((id) => !document.getElementById(id)));
  check('Trang In đủ điều khiển như Excel (Số bản, Máy in, In gì, Trang từ–đến, Đối chiếu, Hướng, Khổ, Lề, Co giãn, Thiết lập Trang, xem trước)', ctl.length === 0, 'thiếu: ' + ctl.join(','));
  const whatOpts = await page.$$eval('#bs-what option', (o) => o.map((x) => x.textContent));
  check('In: Trang tính Hiện hoạt / Toàn bộ Sổ làm việc / Vùng chọn', whatOpts.join('|') === 'In Trang tính Hiện hoạt|In Toàn bộ Sổ làm việc|In Vùng chọn', whatOpts.join('|'));
  const scOpts = await page.$$eval('#bs-scale option', (o) => o.map((x) => x.textContent));
  check('Co giãn: Không Co giãn, Vừa Trang tính trên Một Trang, Vừa Tất cả Cột…, Vừa Tất cả Hàng…, Tùy chọn Co giãn Tùy chỉnh…', scOpts.join('|') === 'Không Co giãn|Vừa Trang tính trên Một Trang|Vừa Tất cả Cột trên Một Trang|Vừa Tất cả Hàng trên Một Trang|Tùy chọn Co giãn Tùy chỉnh…', scOpts.join('|'));
  const pv = () => X(page, () => { const f = document.querySelector('#backstage .bs-preview iframe'), d = f.contentDocument; return { total: document.querySelector('#pv-total').textContent, num: document.querySelector('#pv-num').value, pages: d.querySelectorAll('.page').length, cur: [...d.querySelectorAll('.page')].findIndex((p) => p.classList.contains('cur')), src: f.srcdoc, body: d.body.className }; });
  await page.waitForTimeout(300);
  let v = await pv();
  check('Xem trước: 4 trang, đang ở trang 1', v.total === '4' && v.num === '1' && v.pages === 4 && v.cur === 0, JSON.stringify({ ...v, src: undefined }));
  await page.screenshot({ path: OUT + 'v05-print.png' });
  await page.click('#pv-next'); await page.click('#pv-next');
  v = await pv();
  check('▶ sang trang 3', v.num === '3' && v.cur === 2);
  await page.click('#pv-margins');
  check('Hiện lề trên bản xem trước', /show-margins/.test((await pv()).body));
  await page.selectOption('#bs-scale', 'cols');
  await page.waitForTimeout(300);
  v = await pv();
  check('Vừa Tất cả Cột trên Một Trang → 2 trang', v.total === '2', v.total);
  await page.selectOption('#bs-orient', 'landscape');
  await page.waitForTimeout(300);
  check('Khổ ngang cập nhật bản xem trước', /297mm 210mm/.test((await pv()).src));
  await page.selectOption('#bs-margins', 'narrow');
  check('Lề Hẹp được lưu vào thiết lập trang', await X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).print?.margins === 'narrow'));
  // Thiết lập Trang
  await page.click('#bs-setup');
  const psTabs = await texts(page, '.dlg [role="tab"]');
  check('Hộp Thiết lập Trang: Trang, Lề, Đầu trang/Chân trang, Trang tính', psTabs.join('|') === 'Trang|Lề|Đầu trang/Chân trang|Trang tính', psTabs.join('|'));
  await page.click('.dlg [role="tab"]:has-text("Đầu trang/Chân trang")');
  await page.selectOption('#ps-footer', { label: 'Trang 1 / ?' });
  await page.fill('#ps-hc', 'BẢNG KHỐI LƯỢNG');
  await page.click('.dlg [role="tab"]:has-text("Trang tính")');
  await page.fill('#ps-rows', '$1:$2');
  await page.check('#ps-heads');
  await page.click('.dlg [role="tab"]:has-text("Lề")');
  await page.check('#ps-ch');
  await page.click('.dlg .btn.primary');
  await page.waitForTimeout(300);
  v = await pv();
  const ps = await X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).print);
  check('Thiết lập Trang ghi đầu/chân trang, hàng lặp lại, in tiêu đề, căn giữa ngang', ps.footer === '&CTrang &P / &N' && ps.header === '&CBẢNG KHỐI LƯỢNG' && JSON.stringify(ps.titleRowRange) === '[0,1]' && ps.headings === true && ps.center?.h === true, JSON.stringify(ps));
  check('Xem trước hiện đầu trang và “Trang 1 / 2”', v.src.includes('BẢNG KHỐI LƯỢNG') && v.src.includes('Trang 1 / 2'));
  await page.screenshot({ path: OUT + 'v05-print-setup.png' });
  await page.selectOption('#bs-what', 'workbook');
  await page.waitForTimeout(300);
  check('In Toàn bộ Sổ làm việc', (await pv()).total === '2');
  await page.fill('#bs-from', '2'); await page.fill('#bs-to', '2'); await page.press('#bs-to', 'Tab');
  await page.click('#bs-print');
  await page.waitForTimeout(300);
  check('Nút In (bản web) in đúng các trang đã chọn qua khung ẩn', await X(page, () => { const f = document.querySelector('iframe.print-frame'); return !!f && (f.srcdoc.match(/class="page"/g) ?? []).length === 1; }));
  // thẻ Bố trí Trang
  await page.keyboard.press('Escape');
  await page.click('#tabs [role="tab"]:has-text("Bố trí Trang")');
  await goto(page, 'C21');
  await page.click('#ribbon [data-id="pl-breaks"]');
  await page.click('.menu .mi:has-text("Chèn Ngắt Trang")');
  check('Bố trí Trang ▸ Ngắt ▸ Chèn Ngắt Trang', JSON.stringify(await X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).breaks)) === '{"rows":[20],"cols":[2]}');
  await page.click('#ribbon [data-id="pl-area"]');
  await page.click('.menu .mi:has-text("Đặt Vùng In")');
  check('Vùng In ▸ Đặt Vùng In', !!(await X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).print?.area)));
  await page.click('#views [data-view="break"]');
  check('Xem trước Ngắt Trang: lưới hiện số trang và đường ngắt', await X(page, () => window.__xl.grid.view === 'break'));
  await page.screenshot({ path: OUT + 'v05-pagebreak.png' });
  await page.click('#views [data-view="normal"]');
  check('Không lỗi trang (In)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

{
  const { ctx, page } = await newPage();
  await page.click('#tabs [role="tab"]:has-text("Tệp")');
  await page.click('#backstage [data-bs="account"]');
  await page.selectOption('#bs-theme', 'darkgray');
  const dg = await X(page, () => ({ t: document.documentElement.dataset.theme, bg: getComputedStyle(document.querySelector('#titlebar')).backgroundColor }));
  check('Tài khoản ▸ Chủ đề Office: Xám Đậm', dg.t === 'darkgray' && dg.bg !== 'rgb(16, 124, 65)', JSON.stringify(dg));
  const themes = await page.$$eval('#bs-theme option', (o) => o.map((x) => x.textContent));
  check('Có đủ 5 chủ đề: Nhiều màu, Xám Đậm, Đen, Trắng, Dùng thiết đặt hệ thống', themes.join('|') === 'Nhiều màu|Xám Đậm|Đen|Trắng|Dùng thiết đặt hệ thống', themes.join('|'));
  await page.screenshot({ path: OUT + 'v05-darkgray.png' });
  await page.selectOption('#bs-theme', 'colorful');
  await page.click('#backstage [data-bs="options"]');
  await page.check('#opt-iter');
  check('Tùy chọn ▸ Công thức: Bật tính lặp', await X(page, () => window.__xl.wb.iterate.on));
  await page.selectOption('#opt-enter', 'right');
  await page.click('#bs-back');
  check('Nút ← quay về trang tính', !(await visible(page, '#backstage')));
  await type(page, 'C5', 'x');
  check('Tùy chọn “Sau khi nhấn Enter: sang phải”', await X(page, () => window.__xl.grid.sel.ar === 4 && window.__xl.grid.sel.ac === 3));
  await page.keyboard.press('Control+p');
  check('Ctrl+P mở trang In trong Backstage', await visible(page, '#backstage .bs-preview'));
  await page.keyboard.press('Escape');
  check('Esc đóng Backstage', !(await visible(page, '#backstage')));
  check('Không lỗi trang (Backstage)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ================= Thanh trạng thái, thẻ sheet, menu chuột phải =================
{
  const { ctx, page } = await newPage();
  for (const [a, v] of [['A1', '10'], ['A2', '20'], ['A3', '30'], ['B1', 'x']]) await type(page, a, v);
  await goto(page, 'A1:A3');
  const agg = await X(page, () => ({ mode: document.querySelector('#mode').textContent, keys: [...document.querySelectorAll('#agg [data-agg]')].map((e) => e.dataset.agg + '=' + e.textContent) }));
  check('Thanh trạng thái: Sẵn sàng, Trung bình, Số đếm, Tổng', agg.mode === 'Sẵn sàng' && agg.keys.join('|') === 'avg=Trung bình: 20|count=Số đếm: 3|sum=Tổng: 60', JSON.stringify(agg));
  await page.click('#status', { button: 'right' });
  const sItems = await menuItems(page);
  check('Chuột phải thanh trạng thái: Tối thiểu, Tối đa, Số đếm Số…', ['Trung bình', 'Số đếm', 'Số đếm Số', 'Tối thiểu', 'Tối đa', 'Tổng', 'Thanh trượt Thu phóng'].every((x) => sItems.includes(x)), sItems.join(','));
  await page.click('.menu .mi:has-text("Tối đa")');
  check('Bật Tối đa → hiện “Tối đa: 30”', (await page.textContent('#agg [data-agg="max"]').catch(() => '')) === 'Tối đa: 30');
  const views = await X(page, () => [...document.querySelectorAll('#views [data-view]')].map((e) => e.dataset.view + ':' + e.getAttribute('aria-pressed')));
  check('3 nút dạng xem: Thường (đang chọn), Bố trí Trang, Xem trước Ngắt Trang', views.join('|') === 'normal:true|layout:false|break:false', views.join('|'));
  await page.$eval('#zoom-slider', (e) => { e.value = '75'; e.dispatchEvent(new Event('input', { bubbles: true })); });
  const z1 = await X(page, () => window.__xl.grid.zoom);
  check('Thanh trượt thu phóng (giữa = 100%, phải tới 400%)', z1 > 1.9 && z1 < 2.6, String(z1));
  await page.click('#zoom-val');
  await page.click('.dlg label:has-text("75%") input');
  await page.click('.dlg .btn.primary');
  check('Bấm % mở hộp Thu phóng → 75%', Math.abs((await X(page, () => window.__xl.grid.zoom)) - 0.75) < 0.01);
  await X(page, () => window.__xl.grid.setZoom(1));
  await page.screenshot({ path: OUT + 'v05-status.png' });

  // thẻ sheet
  check('Thẻ sheet có nút ◀ ▶ và ⊕', !!(await page.$('#sheet-prev')) && !!(await page.$('#sheet-next')) && !!(await page.$('#sheet-add')));
  await page.click('#sheets .st >> nth=0', { button: 'right' });
  const tItems = await menuItems(page);
  check('Chuột phải thẻ sheet: Chèn…, Xóa, Đổi tên, Di chuyển hoặc Sao chép…, Màu Tab, Ẩn, Bỏ ẩn…', ['Chèn…', 'Xóa', 'Đổi tên', 'Di chuyển hoặc Sao chép…', 'Màu Tab', 'Ẩn', 'Bỏ ẩn…', 'Chọn Tất cả Trang tính'].every((x) => tItems.includes(x)), tItems.join(','));
  await page.hover('.menu .mi:has-text("Màu Tab")');
  await page.click('.menu.sub .palette .sw >> nth=2');
  const tc = await X(page, () => ({ c: window.__xl.wb.meta(window.__xl.grid.sheet).tabColor, el: getComputedStyle(document.querySelector('#sheets .st')).getPropertyValue('--tab').trim() }));
  check('Màu Tab đặt màu cho thẻ', !!tc.c && tc.el === tc.c, JSON.stringify(tc));
  await page.click('#sheets .st >> nth=0', { button: 'right' });
  await page.click('.menu .mi:has-text("Di chuyển hoặc Sao chép…")');
  await page.check('.dlg #mv-copy');
  await page.click('.dlg .btn.primary');
  const sh = await texts(page, '#sheets .st');
  check('Di chuyển hoặc Sao chép… → Tạo bản sao', sh.length === 2 && sh[1] === 'Sheet1 (2)', sh.join(','));
  await page.click('#sheets .st >> nth=0', { button: 'right' });
  await page.click('.menu .mi:has-text("Ẩn")');
  check('Ẩn trang tính', (await texts(page, '#sheets .st')).join() === 'Sheet1 (2)');
  await page.click('#sheets .st >> nth=0', { button: 'right' });
  await page.click('.menu .mi:has-text("Bỏ ẩn…")');
  await page.click('.dlg .btn.primary');
  check('Bỏ ẩn… trang tính', (await texts(page, '#sheets .st')).join() === 'Sheet1,Sheet1 (2)');
  await page.click('#sheets .st >> nth=0');

  // menu chuột phải ô + thanh mini
  const p = await pt(page, 1, 0);
  await page.mouse.click(p.x, p.y, { button: 'right' });
  await page.waitForSelector('.mini-tb');
  const cItems = await menuItems(page);
  check('Menu chuột phải ô đủ mục như Excel', ['Cắt', 'Sao chép', 'Tùy chọn Dán:', 'Dán Đặc biệt…', 'Chèn…', 'Xóa…', 'Xóa Nội dung', 'Lọc', 'Sắp xếp', 'Bình luận Mới', 'Ghi chú Mới', 'Định dạng Ô…', 'Chọn từ Danh sách Thả xuống…', 'Xác định Tên…', 'Liên kết'].every((x) => cItems.includes(x)), cItems.join(','));
  const mini = await X(page, () => { const a = document.querySelector('.mini-tb').getBoundingClientRect(), b = document.querySelector('.ctx-wrap .menu').getBoundingClientRect(); return { above: a.bottom <= b.top + 1, ids: [...document.querySelectorAll('.mini-tb [data-m]')].map((e) => e.dataset.m) }; });
  check('Thanh mini nằm trên menu: phông, cỡ, A↑ A↓, $, %, 000, B, I, căn giữa, tô, màu chữ, viền, gộp', mini.above && ['fn', 'fs', 'grow', 'shrink', 'acct', 'pct', 'comma', 'b', 'i', 'center', 'bg', 'fg', 'border', 'merge'].every((k) => mini.ids.includes(k)), JSON.stringify(mini));
  await page.screenshot({ path: OUT + 'v05-ctx.png' });
  await page.click('.mini-tb [data-m="b"]');
  check('Thanh mini: Đậm', (await st(page, 1, 0)).b === true);
  await goto(page, 'A1:A3');
  const p2 = await pt(page, 1, 0);
  await page.mouse.click(p2.x, p2.y, { button: 'right' });
  await page.hover('.menu .mi:has-text("Sắp xếp")');
  await page.click('.menu.sub .mi:has-text("Sắp xếp từ Lớn nhất đến Nhỏ nhất")');
  check('Sắp xếp ▸ Lớn nhất đến Nhỏ nhất', (await val(page, 0, 0)) === 30 && (await val(page, 2, 0)) === 10);
  await goto(page, 'A1');
  const p3 = await pt(page, 0, 0);
  await page.mouse.click(p3.x, p3.y, { button: 'right' });
  await page.click('.menu .mi:has-text("Chèn…")');
  await page.click('.dlg label:has-text("Dồn ô xuống") input');
  await page.click('.dlg .btn.primary');
  check('Chèn… ▸ Dồn ô xuống', (await val(page, 0, 0)) === null && (await val(page, 1, 0)) === 30 && (await val(page, 0, 1)) === 'x');
  // tiêu đề hàng / cột
  const rh = await hdr(page, 'r', 2);
  await page.mouse.click(rh.x, rh.y, { button: 'right' });
  const rItems = await menuItems(page);
  check('Chuột phải tiêu đề hàng: Chèn, Xóa, Chiều cao Hàng…, Ẩn, Bỏ ẩn', ['Chèn', 'Xóa', 'Chiều cao Hàng…', 'Ẩn', 'Bỏ ẩn'].every((x) => rItems.includes(x)), rItems.join(','));
  await page.click('.menu .mi:has-text("Ẩn")');
  check('Ẩn hàng 3', await X(page, () => window.__xl.wb.isRowHidden(window.__xl.grid.sheet, 2)));
  await goto(page, '2:4');
  const rh2 = await hdr(page, 'r', 1);
  await page.mouse.click(rh2.x, rh2.y, { button: 'right' });
  await page.click('.menu .mi:has-text("Bỏ ẩn")');
  check('Bỏ ẩn hàng', !(await X(page, () => window.__xl.wb.isRowHidden(window.__xl.grid.sheet, 2))));
  const ch = await hdr(page, 'c', 3);
  await page.mouse.click(ch.x, ch.y, { button: 'right' });
  await page.click('.menu .mi:has-text("Độ rộng Cột…")');
  await page.fill('.dlg input', '20');
  await page.click('.dlg .btn.primary');
  check('Độ rộng Cột… theo ký tự như Excel (20 → 145 px)', (await X(page, () => window.__xl.wb.meta(window.__xl.grid.sheet).colW.get(3))) === 145);
  // lưới
  const gv = await X(page, () => { const cs = getComputedStyle(document.querySelector('#gridwrap')); return { hdr: cs.getPropertyValue('--hdr').trim(), line: cs.getPropertyValue('--line').trim(), font: cs.getPropertyValue('--font-cell').trim(), rowH: window.__xl.wb.meta(window.__xl.grid.sheet).rowH.def, colW: window.__xl.wb.meta(window.__xl.grid.sheet).colW.def }; });
  check('Lưới theo Excel: tiêu đề #f5f5f5, đường lưới #e1e1e1, phông ô Aptos Narrow/Calibri, ô 64×20', gv.hdr === '#f5f5f5' && gv.line === '#e1e1e1' && /Aptos Narrow.*Calibri/.test(gv.font) && gv.rowH === 20 && gv.colW === 64, JSON.stringify(gv));
  check('Không lỗi trang (trạng thái/sheet/menu)', page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

// ================= Thu gọn ribbon, nhóm co lại khi hẹp =================
{
  const { ctx, page } = await newPage({ viewport: { width: 1000, height: 760 } });
  const col = await X(page, () => [...document.querySelectorAll('#ribbon .grp.collapsed')].map((e) => e.querySelector('.lbl').textContent));
  check('Cửa sổ hẹp: nhóm bên phải co thành một nút (như Excel)', col.length >= 1 && !(await X(page, () => document.querySelector('#ribbon').scrollWidth > document.querySelector('#ribbon').clientWidth + 1)), col.join(','));
  await page.click('#ribbon .grp.collapsed >> nth=0 >> .big');
  check('Bấm nhóm co → hiện cả nhóm trong bảng thả xuống', !!(await page.$('.grp-pop .grp')));
  await page.keyboard.press('Escape');
  await page.keyboard.press('Control+F1');
  check('Ctrl+F1 thu gọn ribbon chỉ còn thẻ', (await X(page, () => getComputedStyle(document.querySelector('#ribbon')).display)) === 'none');
  await page.keyboard.press('Control+F1');
  await ctx.close();
}

// ================= Điện thoại =================
for (const [name, vp] of [['Dọc 390×844', { width: 390, height: 844 }], ['Ngang 844×390', { width: 844, height: 390 }]]) {
  const { ctx, page } = await newPage({ viewport: vp, isMobile: true, hasTouch: true, deviceScaleFactor: 2 });
  const m = await X(page, () => ({ tabs: getComputedStyle(document.querySelector('#tabs')).display, chip: document.querySelector('#tabsel')?.textContent.trim(), simple: document.querySelector('#ribbon').classList.contains('simple'), cmd: getComputedStyle(document.querySelector('.mob-cmd')).display, gridH: document.querySelector('#gridwrap').getBoundingClientRect().height, overflow: document.scrollingElement.scrollWidth - innerWidth }));
  check(`${name}: dải thẻ ẩn, thay bằng nút chọn thẻ “Trang đầu”`, m.tabs === 'none' && m.chip === 'Trang đầu', JSON.stringify(m));
  check(`${name}: ribbon đơn giản một hàng`, m.simple && (await X(page, () => document.querySelector('#ribbon').getBoundingClientRect().height)) <= 56);
  check(`${name}: thanh lệnh dưới chỉ khi dọc`, name.startsWith('Dọc') ? m.cmd !== 'none' : m.cmd === 'none', m.cmd);
  check(`${name}: lưới đủ cao, không tràn ngang`, m.gridH > (name.startsWith('Dọc') ? 450 : 200) && m.overflow <= 0, JSON.stringify(m));
  await page.tap('#tabsel');
  await page.tap('.menu .mi:has-text("Chèn")');
  check(`${name}: chọn thẻ Chèn`, (await page.textContent('#tabsel')).trim() === 'Chèn' && !!(await page.$('#ribbon [data-id="chart"]')));
  await page.tap('#tabsel'); await page.tap('.menu .mi:has-text("Trang đầu")');
  await goto(page, 'A1');
  await page.tap('#ribbon [data-id="b"]');
  check(`${name}: dùng lệnh Đậm từ ribbon đơn giản`, (await st(page, 0, 0)).b === true);
  const small = await X(page, () => [...document.querySelectorAll('#ribbon button, #titlebar button, #tabsel, .mob-cmd button, .sheets .st')].filter((b) => { const r = b.getBoundingClientRect(); return r.width && r.height && (r.width < 40 || r.height < 40); }).map((b) => b.dataset.id || b.id || b.className));
  check(`${name}: vùng chạm ≥ 40px`, small.length === 0, small.slice(0, 8).join(','));
  await page.screenshot({ path: OUT + `v05-mobile-${vp.width}.png` });
  check(`${name}: không lỗi trang`, page.errors.length === 0, page.errors.join(' | '));
  await ctx.close();
}

await browser.close();
try { process.kill(-srv.pid); } catch { /* đã dừng */ }
const bad = results.filter((r) => !r.ok);
console.log(`\n${results.length - bad.length}/${results.length} đạt`);
process.exit(bad.length ? 1 : 0);
