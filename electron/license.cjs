// Bản quyền ExcelLike: key ký số Ed25519 (khóa bí mật chỉ nằm ở app Quản lý Key; app/bộ cài chỉ có khóa công khai).
// Dùng chung cho: tiến trình chính Electron (kích hoạt), app Quản lý Key (sinh key), test.
//
// Nội dung key (19 byte) rồi tới chữ ký 64 byte, viết bằng Base32 Crockford, nhóm 5 ký tự:
//   [0] phiên bản định dạng = 1   [1] sản phẩm (1 = ExcelLike)   [2] loại: 1 = gắn 1 máy, 2 = dùng mọi máy
//   [3..6] số thứ tự key (uint32)   [7..8] ngày cấp   [9..10] ngày hết hạn (0 = vĩnh viễn) — tính bằng số ngày từ 01/01/2020
//   [11..18] mã băm của máy (8 byte; key dùng mọi máy = 0)
'use strict';
const crypto = require('node:crypto');

const ALPHA = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';
const PRODUCT = { ExcelLike: 1 };
const TYPE = { machine: 1, shared: 2 };
const BODY = 19, SIG = 64;
const EPOCH = Date.UTC(2020, 0, 1);
const DAY = 86400000;

function b32enc(buf) {
  let bits = 0, val = 0, out = '';
  for (const b of buf) { val = (val << 8) | b; bits += 8; while (bits >= 5) { out += ALPHA[(val >>> (bits - 5)) & 31]; bits -= 5; } }
  if (bits > 0) out += ALPHA[(val << (5 - bits)) & 31];
  return out;
}
/** chuẩn hóa chữ người dùng gõ: bỏ khoảng trắng/gạch, chữ thường → hoa, O→0, I/L→1 */
function normalize(s) { return String(s || '').toUpperCase().replace(/[\s\-_.]/g, '').replace(/O/g, '0').replace(/[IL]/g, '1'); }
function b32dec(s, nbytes) {
  const t = normalize(s);
  const out = Buffer.alloc(nbytes);
  let bits = 0, val = 0, i = 0;
  for (const ch of t) {
    const v = ALPHA.indexOf(ch);
    if (v < 0) return null;
    val = ((val << 5) | v) & 0xffff; bits += 5;
    if (bits >= 8) { if (i >= nbytes) return null; out[i++] = (val >>> (bits - 8)) & 255; bits -= 8; }
  }
  return i === nbytes ? out : null;
}
const group = (s, n = 5) => s.match(new RegExp(`.{1,${n}}`, 'g')).join('-');

/** ngày (theo lịch, không tính giờ) → số ngày từ 01/01/2020 */
function dayNum(d) { const x = d instanceof Date ? d : new Date(d); return Math.floor((Date.UTC(x.getFullYear(), x.getMonth(), x.getDate()) - EPOCH) / DAY); }
function dayDate(n) { const d = new Date(EPOCH + n * DAY); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
const ymd = (n) => { const d = dayDate(n); return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`; };

/** mã băm máy từ MachineGuid của Windows (giống hệt lic.exe trong bộ cài) */
function machineHash(guid) {
  return crypto.createHash('sha512').update('ExcelLike|' + String(guid).trim().toUpperCase(), 'utf8').digest().subarray(0, 8);
}
/** Mã máy hiển thị cho người dùng: 13 ký tự, dạng XXXX-XXXX-XXXXX */
function machineCode(hash) { const s = b32enc(hash); return `${s.slice(0, 4)}-${s.slice(4, 8)}-${s.slice(8)}`; }
function parseMachineCode(code) { return b32dec(code, 8); }

/** Khóa công khai/bí mật dạng 32 byte thô (base64) ↔ KeyObject */
const SPKI = Buffer.from('302a300506032b6570032100', 'hex');
const PKCS8 = Buffer.from('302e020100300506032b657004220420', 'hex');
function pubKey(raw) { return crypto.createPublicKey({ key: Buffer.concat([SPKI, Buffer.from(raw, 'base64')]), format: 'der', type: 'spki' }); }
function privKey(raw) { return crypto.createPrivateKey({ key: Buffer.concat([PKCS8, Buffer.from(raw, 'base64')]), format: 'der', type: 'pkcs8' }); }
function newKeyPair() {
  const { publicKey, privateKey } = crypto.generateKeyPairSync('ed25519');
  return {
    pub: publicKey.export({ format: 'der', type: 'spki' }).subarray(12).toString('base64'),
    priv: privateKey.export({ format: 'der', type: 'pkcs8' }).subarray(16).toString('base64'),
  };
}
/** khóa công khai suy ra từ khóa bí mật (để kiểm tra tệp khóa ký có khớp ứng dụng không) */
function pubFromPriv(privRaw) { return crypto.createPublicKey(privKey(privRaw)).export({ format: 'der', type: 'spki' }).subarray(12).toString('base64'); }

/**
 * Sinh key. o = { serial, type: 'machine'|'shared', machine?: mã máy (chuỗi) hoặc Buffer 8 byte, issued?: Date, expires?: Date|null }
 */
function makeKey(o, privRaw) {
  const body = Buffer.alloc(BODY);
  body[0] = 1; body[1] = PRODUCT.ExcelLike;
  body[2] = TYPE[o.type];
  if (!body[2]) throw new Error('Loại key không hợp lệ');
  if (!Number.isInteger(o.serial) || o.serial < 1 || o.serial > 0xffffffff) throw new Error('Số thứ tự key không hợp lệ');
  body.writeUInt32BE(o.serial, 3);
  const iss = dayNum(o.issued || new Date());
  body.writeUInt16BE(iss, 7);
  const exp = o.expires ? dayNum(o.expires) : 0;
  if (o.expires && (exp < iss || exp > 0xffff)) throw new Error('Ngày hết hạn phải sau ngày cấp');
  body.writeUInt16BE(exp, 9);
  if (o.type === 'machine') {
    const h = Buffer.isBuffer(o.machine) ? o.machine : parseMachineCode(o.machine);
    if (!h) throw new Error('Mã máy không hợp lệ (13 ký tự, dạng XXXX-XXXX-XXXXX)');
    h.copy(body, 11);
  }
  const sig = crypto.sign(null, body, privKey(privRaw));
  return group(b32enc(Buffer.concat([body, sig])));
}

/** Đọc key (chưa kiểm chữ ký) */
function decodeKey(text) {
  const all = b32dec(text, BODY + SIG);
  if (!all) return null;
  const body = all.subarray(0, BODY), sig = all.subarray(BODY);
  if (body[0] !== 1) return null;
  const type = body[2] === 1 ? 'machine' : body[2] === 2 ? 'shared' : null;
  if (!type) return null;
  const mid = body.subarray(11, 19);
  return {
    body, sig, product: body[1], type, serial: body.readUInt32BE(3),
    issued: body.readUInt16BE(7), expires: body.readUInt16BE(9) || 0,
    machine: type === 'machine' ? machineCode(mid) : null, mid,
  };
}

/**
 * Kiểm key. ctx = { pub (base64), machine?: Buffer 8 byte của máy này, today?: Date, revoked?: number[], lastSeen?: số ngày }
 * Trả về { ok, code, info }: code = ok | format | signature | product | machine | expired | revoked | clock
 */
function checkKey(text, ctx) {
  const k = decodeKey(text);
  if (!k) return { ok: false, code: 'format' };
  let good = false;
  try { good = crypto.verify(null, k.body, pubKey(ctx.pub), k.sig); } catch { good = false; }
  const info = { serial: k.serial, type: k.type, issued: ymd(k.issued), expires: k.expires ? ymd(k.expires) : null, machine: k.machine, expiresDay: k.expires };
  if (!good) return { ok: false, code: 'signature' };
  if (k.product !== PRODUCT.ExcelLike) return { ok: false, code: 'product', info };
  if ((ctx.revoked || []).includes(k.serial)) return { ok: false, code: 'revoked', info };
  if (k.type === 'machine' && ctx.machine && !k.mid.equals(ctx.machine)) return { ok: false, code: 'machine', info };
  const today = dayNum(ctx.today || new Date());
  if (k.expires) {
    // đồng hồ máy bị chỉnh lùi quá 2 ngày so với lần dùng gần nhất → không tin ngày hiện tại
    if (ctx.lastSeen && today < ctx.lastSeen - 2) return { ok: false, code: 'clock', info };
    if (Math.max(today, ctx.lastSeen || 0) > k.expires) return { ok: false, code: 'expired', info };
    info.daysLeft = k.expires - today;
  }
  return { ok: true, code: 'ok', info };
}

const MESSAGES = {
  ok: 'Key hợp lệ.',
  format: 'Key không đúng định dạng. Hãy dán nguyên key (gồm các nhóm 5 ký tự).',
  signature: 'Key không hợp lệ (sai chữ ký). Kiểm tra lại từng ký tự hoặc liên hệ nơi cấp key.',
  product: 'Key này không dành cho ExcelLike.',
  machine: 'Key này cấp cho máy khác. Gửi Mã máy của máy này để được cấp key mới.',
  expired: 'Key đã hết hạn sử dụng.',
  revoked: 'Key này đã bị thu hồi.',
  clock: 'Ngày giờ của máy đang bị lùi so với lần dùng trước. Chỉnh lại đồng hồ Windows rồi mở lại.',
};

module.exports = { ALPHA, b32enc, b32dec, normalize, dayNum, dayDate, ymd, machineHash, machineCode, parseMachineCode, newKeyPair, pubFromPriv, makeKey, decodeKey, checkKey, MESSAGES };
