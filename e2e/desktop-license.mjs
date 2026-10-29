// Kiểm thử kích hoạt bản quyền + khởi động an toàn trên bản máy tính (Electron).
// Chạy: xvfb-run -a node e2e/desktop-license.mjs   (cần npm run build và tệp khóa ký)
import { createRequire } from 'node:module';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { L, testKey, TEST_GUID } from './testkey.mjs';
const require = createRequire(import.meta.url);
const { _electron: electron } = require('playwright');
const electronBin = require('electron');
const ROOT = new URL('..', import.meta.url).pathname;
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
let pass = 0, fail = 0;
const check = (n, ok, info = '') => { ok ? pass++ : fail++; console.log(`${ok ? '✓' : '✗'} ${n}${info ? ' — ' + info : ''}`); };
const MC = L.machineCode(L.machineHash(TEST_GUID));

async function launch(userData, extra = {}, args = []) {
  const app = await electron.launch({ executablePath: electronBin, args: [ROOT, ...args], env: { ...process.env, XL_USER_DATA: userData, XL_TEST_MACHINE: TEST_GUID, ...extra } });
  const win = await app.firstWindow();
  await win.waitForFunction(() => window.__xl, null, { timeout: 30000 });
  return { app, win };
}
const fresh = () => { const d = join(mkdtempSync(join(tmpdir(), 'xl-lic-')), 'ud'); mkdirSync(d, { recursive: true }); return d; };

// 1. chưa có key → màn kích hoạt chặn thao tác, hiện đúng Mã máy
{
  const ud = fresh();
  const { app, win } = await launch(ud);
  await win.waitForSelector('#lic-back', { timeout: 10000 });
  check('Chưa có key: hiện màn Kích hoạt', true);
  check('Hiện đúng Mã máy của máy này', (await win.textContent('.lic-mc code')) === MC, MC);
  await win.screenshot({ path: OUT + 'lic-activate.png' });
  // phím tắt bảng tính bị chặn
  await win.keyboard.press('Control+n');
  check('Phím tắt bị chặn khi chưa kích hoạt', await win.evaluate(() => !!document.querySelector('#lic-back') && document.activeElement?.id === 'lic-key'));
  const saveBlocked = await win.evaluate(() => window.xlDesktop.save({ name: 'x', kind: 'csv', data: 'a', saveAs: true }).then(() => false, () => true));
  check('Chưa kích hoạt: tiến trình chính từ chối lưu tệp', saveBlocked);
  // key sai
  await win.fill('#lic-key', 'ABCDE-FGHJK');
  await win.click('.lic-ok');
  await win.waitForFunction(() => document.querySelector('.lic-msg').textContent.length > 0);
  check('Key sai định dạng: báo lỗi', /định dạng/.test(await win.textContent('.lic-msg')));
  // key gắn máy khác
  await win.fill('#lic-key', testKey({ serial: 9001, type: 'machine', machine: L.machineCode(L.machineHash('khac')) }));
  await win.click('.lic-ok');
  await win.waitForFunction(() => /máy khác/.test(document.querySelector('.lic-msg').textContent));
  check('Key của máy khác: báo đúng lý do', true);
  // key hết hạn
  await win.fill('#lic-key', testKey({ serial: 9002, issued: new Date(2025, 0, 1), expires: new Date(2025, 1, 1) }));
  await win.click('.lic-ok');
  await win.waitForFunction(() => /hết hạn/.test(document.querySelector('.lic-msg').textContent));
  check('Key hết hạn: báo hết hạn', true);
  // key đúng máy (có dấu cách, chữ thường)
  const good = testKey({ serial: 9003, type: 'machine', machine: MC, expires: new Date(Date.now() + 10 * 86400000) });
  await win.fill('#lic-key', good.toLowerCase().replace(/-/g, ' '));
  await win.click('.lic-ok');
  await win.waitForSelector('#lic-back', { state: 'detached', timeout: 5000 }).catch(() => {});
  check('Key đúng: kích hoạt, mở khóa giao diện', !(await win.$('#lic-back')));
  check('Key được lưu vào license.key', existsSync(join(ud, 'license.key')) && L.decodeKey(readFileSync(join(ud, 'license.key'), 'utf8'))?.serial === 9003);
  const okSave = await win.evaluate(() => window.xlDesktop.license().then((r) => r.ok));
  check('Sau kích hoạt: trạng thái hợp lệ', okSave);
  await win.evaluate(() => window.__xl.backstage.open('account'));
  await win.waitForSelector('#bs-lic');
  const t = await win.textContent('#bs-lic');
  check('Tệp ▸ Tài khoản hiện thông tin key (số, gắn máy, hạn)', /key số 9003/.test(t) && /gắn với máy này/.test(t) && /còn \d+ ngày/.test(t), t);
  await win.screenshot({ path: OUT + 'lic-account.png' });
  await app.close();
  // mở lại: không hỏi key nữa; key sắp hết hạn → nhắc
  const b = await launch(ud);
  await b.win.waitForTimeout(1500);
  check('Mở lại: không hỏi key nữa', !(await b.win.$('#lic-back')));
  check('Key còn ≤ 14 ngày: nhắc gia hạn', /còn \d+ ngày/.test(await b.win.evaluate(() => [...document.querySelectorAll('.toast')].map((x) => x.textContent).join('|'))));
  await b.app.close();
  // đồng hồ bị chỉnh về sau ngày hết hạn → báo hết hạn khi mở
  const c = await launch(ud, { XL_TEST_TODAY: new Date(Date.now() + 40 * 86400000).toISOString() });
  await c.win.waitForSelector('#lic-back', { timeout: 10000 });
  check('Quá hạn khi mở: màn kích hoạt báo hết hạn', /hết hạn/.test(await c.win.textContent('.lic-msg')));
  await c.app.close();
}

// 1b. chưa kích hoạt mà sổ có thay đổi chưa lưu ra tệp → bấm Thoát vẫn đóng được (không kẹt hộp hỏi lưu)
{
  const ud = fresh();
  const { app, win } = await launch(ud);
  await win.waitForSelector('#lic-back', { timeout: 10000 });
  await win.evaluate(() => { window.__xl.fileDirty = true; });
  const closed = app.waitForEvent('close', { timeout: 10000 }).then(() => true, () => false);
  await win.click('.lic-quit');
  check('Màn kích hoạt: Thoát đóng được ứng dụng', await closed);
}

// 1c. đồng hồ từng bị đặt sai (lastSeen ở tương lai xa) không khóa oan; key mới (gia hạn) kích hoạt được
{
  const ud = fresh();
  writeFileSync(join(ud, 'license.key'), testKey({ serial: 9010, expires: new Date(Date.now() + 100 * 86400000) }));
  writeFileSync(join(ud, 'license-state.json'), JSON.stringify({ lastSeen: L.dayNum(new Date()) + 2000 }));
  let { app, win } = await launch(ud);
  await win.waitForTimeout(1500);
  check('lastSeen sai năm (xa > 1 năm): bỏ qua, không khóa', !(await win.$('#lic-back')));
  await app.close();
  writeFileSync(join(ud, 'license-state.json'), JSON.stringify({ lastSeen: L.dayNum(new Date()) + 200 }));
  ({ app, win } = await launch(ud));
  await win.waitForSelector('#lic-back', { timeout: 10000 });
  check('Đồng hồ lùi so với lần dùng trước: báo lỗi đồng hồ / hết hạn', /đồng hồ|hết hạn/i.test(await win.textContent('.lic-msg')));
  await win.fill('#lic-key', testKey({ serial: 9011, expires: new Date(Date.now() + 365 * 86400000) }));
  await win.click('.lic-ok');
  await win.waitForSelector('#lic-back', { state: 'detached', timeout: 5000 }).catch(() => {});
  check('Nhập key mới (gia hạn) thì kích hoạt được, đặt lại ngày dùng', !(await win.$('#lic-back')) && JSON.parse(readFileSync(join(ud, 'license-state.json'), 'utf8')).lastSeen === L.dayNum(new Date()));
  await app.close();
}

// 2. key do bộ cài ghi sẵn (dùng mọi máy) → mở là dùng được
{
  const ud = fresh();
  writeFileSync(join(ud, 'license.key'), testKey({ serial: 9004 }) + '\r\n');
  const { app, win } = await launch(ud);
  await win.waitForTimeout(1500);
  check('Key bộ cài ghi sẵn (có xuống dòng CRLF): tự kích hoạt', !(await win.$('#lic-back')));
  await app.close();
}

// 3. khởi động an toàn: lần trước không khởi động xong → sổ trống, tiếng Việt, có thông báo
{
  const ud = fresh();
  writeFileSync(join(ud, 'license.key'), testKey({ serial: 9005 }));
  writeFileSync(join(ud, 'boot-pending'), '1');
  const { app, win } = await launch(ud);
  await win.waitForTimeout(1200);
  check('Lần trước lỗi khởi động → mở chế độ an toàn (?safe=1)', /safe=1/.test(win.url()), win.url());
  check('Chế độ an toàn có thông báo', /chế độ an toàn/.test(await win.evaluate(() => [...document.querySelectorAll('.toast')].map((x) => x.textContent).join('|'))));
  check('Khởi động xong thì xóa cờ boot-pending', !existsSync(join(ud, 'boot-pending')));
  await app.close();
}
console.log(`${pass}/${pass + fail} kiểm thử bản quyền + khởi động an toàn đạt`);
process.exit(fail ? 1 : 0);
