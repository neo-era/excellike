// Kiểm thử app Quản lý Key: nhập khóa ký, cấp key (gắn máy / mọi máy / có hạn), sổ key, tìm, thu hồi, kiểm tra key,
// xuất CSV, sao lưu/khôi phục; key sinh ra được ExcelLike chấp nhận (cùng khóa công khai).
// Chạy: xvfb-run -a node e2e/keymgr.mjs
import { createRequire } from 'node:module';
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { L, signing } from './testkey.mjs';
const require = createRequire(import.meta.url);
const { _electron: electron } = require('playwright');
const electronBin = require('electron');
const ROOT = new URL('..', import.meta.url).pathname;
const CFG = require('../electron/license-config.cjs');
const OUT = new URL('./shots/', import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });
let pass = 0, fail = 0;
const check = (n, ok, info = '') => { ok ? pass++ : fail++; console.log(`${ok ? '✓' : '✗'} ${n}${info ? ' — ' + info : ''}`); };

const work = mkdtempSync(join(tmpdir(), 'xl-km-'));
const sigFile = join(work, 'khoa-ky.json');
writeFileSync(sigFile, JSON.stringify({ pub: signing.pub, priv: signing.priv }));
const badSig = join(work, 'sai.json');
writeFileSync(badSig, '{"hello":1}');
const saveOut = join(work, 'out.txt');
const env = (o = {}) => ({ ...process.env, XL_KM_DATA: join(work, 'data'), XL_KM_SAVE: saveOut, ...o });
const launch = async (o) => { const app = await electron.launch({ executablePath: electronBin, args: [join(ROOT, 'keymgr')], env: env(o) }); const w = await app.firstWindow(); await w.waitForSelector('#main h1'); w.on('dialog', (d) => d.accept()); return { app, w }; };
const MC = L.machineCode(L.machineHash('ABCDEF01-2345-6789-ABCD-EF0123456789'));

{
  const { app, w } = await launch({ XL_KM_SIGNING_FILE: badSig });
  check('Lần đầu: yêu cầu nhập khóa ký', /Nhập khóa ký/.test(await w.textContent('#main h1')));
  await w.screenshot({ path: OUT + 'km-onboard.png' });
  await w.click('#imp');
  await w.waitForFunction(() => document.querySelector('#imp-err').textContent.length > 0);
  check('Tệp sai: báo lỗi, không nhận', /không phải tệp khóa ký/i.test(await w.textContent('#imp-err')));
  await app.close();
}
{
  const { app, w } = await launch({ XL_KM_SIGNING_FILE: sigFile });
  await w.click('#imp');
  await w.waitForSelector('#f');
  check('Nhập khóa ký đúng → vào màn Cấp key', true);
  const enc = readFileSync(join(work, 'data', 'signing.bin'));
  // Windows: mã hóa DPAPI; Linux không có kho khóa (chỉ khi kiểm thử) → dạng thô có đánh dấu
  const tag = enc.subarray(0, 6).toString();
  check('Khóa ký lưu mã hóa (DPAPI trên Windows)', tag === 'DPAPI1' ? !enc.toString().includes(signing.priv) : process.platform !== 'win32' && tag === 'PLAIN1', tag);
  check('Khóa ký khớp khóa công khai của ExcelLike', (await w.evaluate(() => window.km.state())).pub === CFG.PUBLIC_KEY);

  // thiếu tên → lỗi
  await w.click('#go');
  check('Thiếu tên khách hàng: báo lỗi', /tên khách hàng/i.test(await w.textContent('#err')));
  // gắn máy, mã máy sai
  await w.fill('#customer', 'Nguyễn Văn A <b>x</b>');
  await w.fill('[name=org]', 'XN Chiếu sáng 1');
  await w.fill('#machine', 'ABC');
  await w.click('#go');
  await w.waitForFunction(() => /Mã máy/.test(document.querySelector('#err').textContent));
  check('Mã máy sai: báo lỗi', true);
  // gõ mã máy thường, tự định dạng
  await w.fill('#machine', '');
  await w.type('#machine', MC.replace(/-/g, '').toLowerCase());
  check('Ô Mã máy tự viết hoa và chèn gạch', (await w.inputValue('#machine')) === MC, await w.inputValue('#machine'));
  await w.click('#go');
  await w.waitForSelector('#keybox');
  const k1 = await w.textContent('#keybox');
  const r1 = L.checkKey(k1, { pub: CFG.PUBLIC_KEY, machine: L.machineHash('ABCDEF01-2345-6789-ABCD-EF0123456789') });
  check('Key gắn máy: ExcelLike chấp nhận trên đúng máy', r1.ok && r1.info.serial === 1 && r1.info.type === 'machine', JSON.stringify(r1.info));
  check('Key gắn máy: máy khác bị từ chối', L.checkKey(k1, { pub: CFG.PUBLIC_KEY, machine: L.machineHash('khac') }).code === 'machine');
  check('Tên có ký tự HTML hiển thị nguyên văn (không chèn mã)', (await w.textContent('#res h2')).includes('<b>x</b>') && !(await w.$('#res h2 b')));
  await w.screenshot({ path: OUT + 'km-issue.png' });
  await w.click('#c-msg');
  const msg = await app.evaluate(({ clipboard }) => clipboard.readText());
  check('Chép tin nhắn gửi khách có key và hướng dẫn', msg.includes(k1) && /Kích hoạt bản quyền/.test(msg));
  // dùng mọi máy, 30 ngày
  await w.fill('#customer', 'Trần Thị B');
  await w.click('.seg [data-type=shared]');
  await w.selectOption('#term', '30');
  await w.click('#go');
  await w.waitForFunction(() => /Key số 2/.test(document.querySelector('#res h2').textContent));
  const k2 = await w.textContent('#keybox');
  const r2 = L.checkKey(k2, { pub: CFG.PUBLIC_KEY, machine: L.machineHash('bat-ky') });
  check('Key dùng mọi máy 30 ngày (tính cả hôm nay): hợp lệ trên máy bất kỳ', r2.ok && r2.info.daysLeft === 29, JSON.stringify(r2.info));
  // đến ngày cụ thể
  await w.fill('#customer', 'Lê Văn C');
  await w.click('.seg [data-type=shared]');
  await w.selectOption('#term', 'date');
  await w.fill('#expires', '2027-06-30');
  await w.click('#go');
  await w.waitForFunction(() => /Key số 3/.test(document.querySelector('#res h2').textContent));
  check('Key đến ngày 30/06/2027', L.checkKey(await w.textContent('#keybox'), { pub: CFG.PUBLIC_KEY }).info.expires === '30/06/2027');

  // sổ key
  await w.click('[data-page=list]');
  await w.waitForSelector('#tb tr[data-s]');
  check('Sổ key có 3 key, mới nhất trên cùng', (await w.$$eval('#tb tr[data-s]', (r) => r.map((x) => x.dataset.s).join(','))) === '3,2,1');
  await w.fill('#q', 'chiếu sáng');
  check('Tìm theo đơn vị', (await w.$$eval('#tb tr[data-s]', (r) => r.length)) === 1);
  await w.fill('#q', '');
  await w.click('.stat[data-f=soon]');
  check('Lọc sắp hết hạn (≤ 30 ngày)', (await w.$$eval('#tb tr[data-s]', (r) => r.map((x) => x.dataset.s).join(','))) === '2');
  await w.click('.stat[data-f=all]');
  await w.screenshot({ path: OUT + 'km-list.png' });
  // chi tiết: sửa ghi chú, thu hồi
  await w.click('#tb tr[data-s="1"]');
  await w.fill('#d-note', 'Máy phòng kỹ thuật');
  await w.click('#d-save');
  await w.waitForSelector('.drawer', { state: 'detached' });
  await w.click('#tb tr[data-s="1"]');
  check('Sửa ghi chú được lưu', (await w.inputValue('#d-note')) === 'Máy phòng kỹ thuật');
  await w.click('#d-rv');
  await w.waitForSelector('.drawer', { state: 'detached' });
  check('Thu hồi key số 1', /Đã thu hồi/.test(await w.textContent('#tb tr[data-s="1"]')));
  await w.click('#rv');
  check('Chép danh sách thu hồi để đưa vào bản cài kế tiếp', (await app.evaluate(({ clipboard }) => clipboard.readText())) === 'REVOKED: [1],');
  // xuất CSV
  await w.click('#csv');
  await w.waitForTimeout(300);
  const csv = readFileSync(saveOut, 'utf8');
  check('Xuất CSV (UTF-8 BOM, dấu ;) đủ 3 key', csv.startsWith('﻿') && csv.split('\r\n').length === 4 && csv.includes('Trần Thị B'));
  // kiểm tra key
  await w.click('[data-page=verify]');
  await w.fill('#v-key', k1);
  await w.click('#v-go');
  await w.waitForSelector('#v-res dl');
  const vt = await w.textContent('#v-res');
  check('Kiểm tra key: nhận ra khách hàng và trạng thái thu hồi', /Nguyễn Văn A/.test(vt) && /thu hồi/.test(vt));
  await w.fill('#v-key', k2.slice(0, -2) + (k2.endsWith('AA') ? 'BB' : 'AA'));
  await w.click('#v-go');
  await w.waitForFunction(() => /chữ ký/.test(document.querySelector('#v-res').textContent));
  check('Kiểm tra key sửa ký tự: báo sai chữ ký', true);
  // sao lưu sổ key, khôi phục vào máy mới
  await w.click('[data-page=settings]');
  await w.click('#s-bak');
  await w.waitForTimeout(300);
  const bak = JSON.parse(readFileSync(saveOut, 'utf8'));
  check('Sao lưu sổ key ra .json', bak.keys.length === 3 && bak.nextSerial === 4);
  await w.screenshot({ path: OUT + 'km-settings.png' });
  await app.close();
  writeFileSync(join(work, 'bak.json'), JSON.stringify(bak));
}
{
  // máy mới: khôi phục sổ key, số kế tiếp nối tiếp (không cấp trùng số)
  const app = await electron.launch({ executablePath: electronBin, args: [join(ROOT, 'keymgr')], env: env({ XL_KM_DATA: join(work, 'data2'), XL_KM_SIGNING_FILE: sigFile, XL_KM_OPEN: join(work, 'bak.json') }) });
  const w = await app.firstWindow(); await w.waitForSelector('#main h1');
  await w.click('#imp'); await w.waitForSelector('#f');
  await w.click('[data-page=settings]'); await w.click('#s-res');
  await w.waitForTimeout(400);
  const st = await w.evaluate(() => window.km.state());
  check('Khôi phục sổ key trên máy mới: đủ 3 key, số kế tiếp = 4', st.keys.length === 3 && st.nextSerial === 4);
  await w.click('#s-res'); await w.waitForTimeout(300);
  check('Khôi phục lại lần nữa không nhân đôi', (await w.evaluate(() => window.km.state())).keys.length === 3);
  await app.close();
}
{
  // sổ key hỏng → dùng bản dự phòng; cả hai hỏng → không cấp key (tránh trùng số), hướng dẫn khôi phục
  const d = join(work, 'data');
  const good = readFileSync(join(d, 'keys.json'), 'utf8');
  writeFileSync(join(d, 'keys.json.bak'), good);
  writeFileSync(join(d, 'keys.json'), '{hỏng');
  let { app, w } = await launch();
  let st = await w.evaluate(() => window.km.state());
  check('Sổ key hỏng: tự dùng bản dự phòng .bak', st.keys.length === 3 && st.nextSerial === 4 && !st.dbError);
  await app.close();
  writeFileSync(join(d, 'keys.json'), '{hỏng'); writeFileSync(join(d, 'keys.json.bak'), 'x');
  ({ app, w } = await launch({ XL_KM_OPEN: join(work, 'bak2.json') }));
  check('Sổ và bản dự phòng đều hỏng: báo lỗi, không cho cấp key', /Sổ key bị hỏng/.test(await w.textContent('#main h1')));
  const cr = await w.evaluate(() => window.km.create({ customer: 'X', type: 'shared' }));
  check('… và không cấp được key trùng số', !cr.ok, cr.error);
  // khôi phục từ tệp có bản ghi lỗi/giả
  const bak = JSON.parse(readFileSync(join(work, 'bak.json'), 'utf8'));
  bak.keys.push({ serial: 50, key: 'GIẢ', customer: 'x' }, { serial: 2, key: bak.keys[0].key, customer: 'sai số' }, { ...bak.keys[0], issued: '<img src=x>' });
  bak.nextSerial = 99999999999;
  writeFileSync(join(work, 'bak2.json'), JSON.stringify(bak));
  await w.click('#to-set'); await w.click('#s-res'); await w.waitForTimeout(400);
  st = await w.evaluate(() => window.km.state());
  check('Khôi phục bỏ qua bản ghi giả/sai, số kế tiếp hợp lý', st.keys.length === 3 && st.nextSerial === 4 && !st.dbError, `${st.keys.length} ${st.nextSerial}`);
  check('Ngày cấp lấy từ chính key (không tin dữ liệu tệp)', st.keys.every((k) => /^\d{4}-\d\d-\d\d$/.test(k.issued)));
  await app.close();
}
{
  // giao diện hẹp (điện thoại/cửa sổ nhỏ) không tràn ngang
  const app = await electron.launch({ executablePath: electronBin, args: [join(ROOT, 'keymgr')], env: env() });
  const w = await app.firstWindow(); await w.waitForSelector('#f');
  await w.setViewportSize({ width: 380, height: 760 });
  await w.click('[data-page=list]'); await w.waitForTimeout(200);
  check('Cửa sổ hẹp 380px: không tràn ngang trang', await w.evaluate(() => document.documentElement.scrollWidth <= 381));
  await w.screenshot({ path: OUT + 'km-narrow.png' });
  await app.close();
}
console.log(`${pass}/${pass + fail} kiểm thử Key Manager đạt`);
process.exit(fail ? 1 : 0);
