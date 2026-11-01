// Kiểm thử Key Manager bản iPhone (public/keymgr): màn iPhone dọc/ngang/nhỏ, sáng/tối,
// key cấp ra phải được bản Windows (license.cjs) chấp nhận, khóa bằng mật khẩu, sao lưu/khôi phục, chạy không mạng.
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const require = createRequire(import.meta.url);
const { chromium } = require('playwright');
const L = require('../keymgr/electron/license.cjs');
const OUT = new URL('./shots/', import.meta.url).pathname; mkdirSync(OUT, { recursive: true });
const TMP = mkdtempSync(join(tmpdir(), 'kmw-'));
const kp = L.newKeyPair();
const signFile = join(TMP, 'ExcelLike-khoa-ky.json'); writeFileSync(signFile, JSON.stringify({ product: 'ExcelLike', pub: kp.pub, priv: kp.priv }));
const winDb = join(TMP, 'So-key-win.json');
writeFileSync(winDb, JSON.stringify({ version: 1, nextSerial: 3, keys: [1, 2].map((n) => ({ serial: n, key: L.makeKey({ serial: n, type: 'shared' }, kp.priv), customer: 'Win ' + n, status: 'active' })) }));
const PORT = 4199;
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 6000); });
const URL0 = `http://localhost:${PORT}/keymgr/`;
const b = await chromium.launch();
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) pass++; else { fail++; console.log('✗', m); } };
const ctx = await b.newContext({ viewport: { width: 390, height: 844 }, isMobile: true, hasTouch: true, deviceScaleFactor: 3, acceptDownloads: true, permissions: ['clipboard-read', 'clipboard-write'] });
const p = await ctx.newPage();
const errors = []; p.on('pageerror', (e) => errors.push(e.message));
const wait = (ms = 120) => p.waitForTimeout(ms);
const noOverflow = () => p.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1);
await p.goto(URL0);
await p.waitForSelector('#ob');
ok(await noOverflow(), 'không tràn ngang (màn nhập khóa)');
await p.screenshot({ path: OUT + 'kmw-onboard.png' });
// mật khẩu không khớp
let [fc] = await Promise.all([p.waitForEvent('filechooser'), p.click('#ob-file')]); await fc.setFiles(signFile); await wait();
await p.fill('#ob-p1', 'matkhau1'); await p.fill('#ob-p2', 'khac'); await p.click('#ob-go'); await wait();
ok(/không khớp/.test(await p.textContent('#imp-err')), 'báo hai mật khẩu không khớp');
await p.fill('#ob-p2', 'matkhau1'); await p.click('#ob-go');
await p.waitForSelector('#f', { timeout: 15000 });
ok((await p.textContent('.sub b')) === '100001', 'key đầu tiên trên iPhone là số 100001');
// cấp key dùng mọi máy
await p.fill('#customer', 'Nguyễn Văn A'); await p.click('.seg [data-type="shared"]'); await p.selectOption('#term', '365');
await p.click('#go'); await p.waitForSelector('#keybox');
const k1 = (await p.textContent('#keybox')).trim();
const c1 = L.checkKey(k1, { pub: kp.pub });
ok(c1.ok && c1.info.serial === 100001 && c1.info.type === 'shared' && c1.info.expires, 'key iPhone được bản Windows chấp nhận (dùng mọi máy, 1 năm) ' + JSON.stringify(c1));
await p.screenshot({ path: OUT + 'kmw-issued.png' });
ok(await noOverflow(), 'không tràn ngang (key vừa cấp)');
// key gắn máy: mã máy sai → báo lỗi; đúng → khớp máy
await p.fill('#customer', 'Trần B'); await p.click('.seg [data-type="machine"]'); await p.fill('#machine', 'ABC'); await p.click('#go'); await wait(300);
ok(/Mã máy không hợp lệ/.test(await p.textContent('#err')), 'mã máy sai bị từ chối');
const mh = L.machineHash('{1234-ABCD}'); const mc = L.machineCode(mh);
await p.fill('#machine', mc.replace(/-/g, '').toLowerCase()); await p.click('#go'); await wait(500);
const k2 = (await p.textContent('#keybox')).trim();
ok(L.checkKey(k2, { pub: kp.pub, machine: mh }).ok && !L.checkKey(k2, { pub: kp.pub, machine: L.machineHash('khac') }).ok, 'key gắn máy chỉ chạy trên đúng máy');
// sổ key (dạng thẻ trên điện thoại)
await p.click('.side nav [data-page="list"]'); await wait(200);
ok((await p.locator('#cards .kc').count()) === 2, 'sổ key hiện 2 thẻ');
ok(!(await p.isVisible('.tbl-wrap')), 'điện thoại: dùng thẻ thay bảng');
await p.screenshot({ path: OUT + 'kmw-list.png' });
await p.click('#cards .kc >> nth=0'); await wait(250);
ok(await p.isVisible('.drawer'), 'mở chi tiết key');
p.once('dialog', (d) => d.accept());
await p.click('#d-rv'); await wait(300);
ok((await p.locator('#cards .tag.revoked').count()) === 1, 'thu hồi key');
// kiểm tra key
await p.click('.side nav [data-page="verify"]'); await wait(150);
await p.fill('#v-key', k1.toLowerCase().replace(/-/g, ' ')); await p.click('#v-go'); await wait(300);
ok(/Hợp lệ/.test(await p.textContent('#v-res')), 'kiểm tra key: hợp lệ (gõ thường, có khoảng trắng)');
// sao lưu sổ key → tệp tải về đúng cấu trúc
await p.click('.side nav [data-page="settings"]'); await wait(150);
const [dl] = await Promise.all([p.waitForEvent('download'), p.click('#s-bak')]);
const bak = JSON.parse(readFileSync(await dl.path(), 'utf8'));
ok(bak.keys.length === 2 && bak.keys.every((k) => L.checkKey(k.key, { pub: kp.pub }).info?.serial === k.serial), 'sao lưu sổ key: 2 key đúng');
// khôi phục sổ của bản Windows: gộp key 1, 2; số kế tiếp vẫn trong dải iPhone
[fc] = await Promise.all([p.waitForEvent('filechooser'), p.click('#s-res')]); await fc.setFiles(winDb); await wait(400);
const st = await p.evaluate(() => window.km.state());
ok(st.keys.length === 4 && st.nextSerial === 100003, 'gộp sổ Windows: 4 key, số kế tiếp 100003 — ' + st.nextSerial);
await p.screenshot({ path: OUT + 'kmw-settings.png', fullPage: true });
// khóa / mở khóa
await p.click('#s-lock'); await wait(300);
ok(await p.isVisible('#lk'), 'khóa ngay → màn khóa');
ok(!(await p.evaluate(() => !!document.querySelector('#cards, #f'))), 'đang khóa: không thấy sổ key');
await p.fill('#lk-p', 'sai'); await p.click('#lk-go'); await wait(1500);
ok(/không đúng/.test(await p.textContent('#lk-err')), 'mật khẩu sai bị từ chối');
await p.screenshot({ path: OUT + 'kmw-lock.png' });
await p.reload(); await p.waitForSelector('#lk');
ok(true, 'mở lại app: vẫn khóa');
await p.fill('#lk-p', 'matkhau1'); await p.click('#lk-go'); await p.waitForSelector('#f', { timeout: 15000 });
ok((await p.textContent('.sub b')) === '100003', 'mở khóa: sổ key còn nguyên sau khi tải lại');
// không lộ khóa bí mật trong bộ nhớ trình duyệt
const stored = await p.evaluate(() => new Promise((res) => { const r = indexedDB.open('excellike-keymgr'); r.onsuccess = () => { const q = r.result.transaction('kv').objectStore('kv').get('signing'); q.onsuccess = () => res(JSON.stringify(q.result)); }; }));
ok(!stored.includes(kp.priv) && /"ct"/.test(stored), 'khóa ký lưu dạng mã hóa, không có khóa thô');
// đổi mật khẩu
await p.click('.side nav [data-page="settings"]'); await p.click('#s-pw'); await p.fill('#p0', 'matkhau1'); await p.fill('#p1', 'moi12345'); await p.fill('#p2', 'moi12345'); await p.click('.drawer .btn.primary'); await wait(2500);
await p.click('#s-lock'); await wait(200); await p.fill('#lk-p', 'moi12345'); await p.click('#lk-go'); await p.waitForSelector('#s-lock', { timeout: 15000 }); await p.click('.side nav [data-page="issue"]');
ok(true, 'đổi mật khẩu rồi mở bằng mật khẩu mới');
// màn ngang, máy nhỏ, giao diện tối
await p.setViewportSize({ width: 844, height: 390 }); await wait(200);
ok(await noOverflow(), 'màn ngang không tràn');
ok(await p.evaluate(() => { const n = document.querySelector('.side').getBoundingClientRect(); return n.bottom <= innerHeight + 1 && n.top > innerHeight / 2; }), 'màn ngang: thanh điều hướng ở đáy, không mất');
await p.screenshot({ path: OUT + 'kmw-landscape.png' });
await p.setViewportSize({ width: 360, height: 640 }); await wait(200);
ok(await noOverflow(), 'máy nhỏ 360×640 không tràn');
await p.emulateMedia({ colorScheme: 'dark' }); await wait(150);
await p.screenshot({ path: OUT + 'kmw-dark.png' });
await p.emulateMedia({ colorScheme: 'light' });
// không mạng: mở lại vẫn vào được
await p.evaluate(() => navigator.serviceWorker.ready); await p.reload(); await p.waitForSelector('#lk');
await ctx.setOffline(true); await p.reload(); await wait(800);
ok(await p.isVisible('#lk'), 'mất mạng: app vẫn mở (màn khóa)');
await p.fill('#lk-p', 'moi12345'); await p.click('#lk-go'); await wait(1500); await p.click('.side nav [data-page="issue"]'); await p.waitForSelector('#f', { timeout: 15000 });
await p.fill('#customer', 'Ngoại tuyến'); await p.click('.seg [data-type="shared"]'); await p.click('#go'); await wait(600);
const k3 = (await p.textContent('#keybox')).trim();
ok(L.checkKey(k3, { pub: kp.pub }).info?.serial === 100003, 'mất mạng vẫn cấp được key');
await ctx.setOffline(false);
// trang manifest
const mf = await (await p.request.get(URL0 + 'manifest.webmanifest')).json();
ok(mf.display === 'standalone' && mf.icons.some((i) => i.purpose === 'maskable'), 'manifest cài được');
ok(errors.length === 0, 'không lỗi trang: ' + errors.join(' | '));
console.log(`${pass}/${pass + fail} kiểm thử Key Manager bản iPhone đạt`);
await b.close(); try { process.kill(-srv.pid); } catch { /* đã tắt */ }
process.exit(fail ? 1 : 0);
