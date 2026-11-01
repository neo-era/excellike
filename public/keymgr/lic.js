// Key bản quyền ExcelLike — bản web (WebCrypto Ed25519), cùng định dạng với electron/license.cjs của bản Windows.
// Nội dung key 19 byte + chữ ký 64 byte, Base32 Crockford, nhóm 5 ký tự (xem license.cjs).
'use strict';
(function (G) {
  const ALPHA = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';
  const BODY = 19, SIG = 64, EPOCH = Date.UTC(2020, 0, 1), DAY = 86400000;
  const subtle = () => G.crypto.subtle;
  const b64 = (u) => { let s = ''; for (const x of u) s += String.fromCharCode(x); return btoa(s); };
  const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
  const b64u = (u) => b64(u).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const unb64u = (s) => unb64(s.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - (s.length % 4)) % 4));

  function b32enc(buf) {
    let bits = 0, val = 0, out = '';
    for (const b of buf) { val = (val << 8) | b; bits += 8; while (bits >= 5) { out += ALPHA[(val >>> (bits - 5)) & 31]; bits -= 5; } }
    if (bits > 0) out += ALPHA[(val << (5 - bits)) & 31];
    return out;
  }
  function normalize(s) { return String(s || '').toUpperCase().replace(/[\s\-_.]/g, '').replace(/O/g, '0').replace(/[IL]/g, '1'); }
  function b32dec(s, nbytes) {
    const t = normalize(s), out = new Uint8Array(nbytes);
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
  function dayNum(d) { const x = d instanceof Date ? d : new Date(d); return Math.floor((Date.UTC(x.getFullYear(), x.getMonth(), x.getDate()) - EPOCH) / DAY); }
  function dayDate(n) { const d = new Date(EPOCH + n * DAY); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
  const ymd = (n) => { const d = dayDate(n); return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`; };
  function machineCode(hash) { const s = b32enc(hash); return `${s.slice(0, 4)}-${s.slice(4, 8)}-${s.slice(8)}`; }
  function parseMachineCode(code) { return b32dec(code, 8); }

  /** máy có hỗ trợ Ed25519 trong WebCrypto không (Safari/iOS 17+, Chrome 113+) */
  async function supported() {
    try { await subtle().importKey('raw', new Uint8Array(32), { name: 'Ed25519' }, false, ['verify']); return true; } catch { return false; }
  }
  /** khóa bí mật thô 32 byte (base64) → { key: CryptoKey (không xuất được), pub: base64 } */
  async function importPriv(privRaw) {
    const d = unb64(privRaw);
    if (d.length !== 32) throw new Error('Khóa ký không đúng độ dài');
    const tmp = await subtle().importKey('pkcs8', concat(PKCS8, d), { name: 'Ed25519' }, true, ['sign']);
    const jwk = await subtle().exportKey('jwk', tmp);
    const key = await subtle().importKey('jwk', jwk, { name: 'Ed25519' }, false, ['sign']);
    return { key, pub: b64(unb64u(jwk.x)) };
  }
  const PKCS8 = Uint8Array.from([0x30, 0x2e, 0x02, 0x01, 0x00, 0x30, 0x05, 0x06, 0x03, 0x2b, 0x65, 0x70, 0x04, 0x22, 0x04, 0x20]);
  function concat(...a) { const n = a.reduce((s, x) => s + x.length, 0), o = new Uint8Array(n); let p = 0; for (const x of a) { o.set(x, p); p += x.length; } return o; }
  async function pubKey(raw) { return subtle().importKey('raw', unb64(raw), { name: 'Ed25519' }, false, ['verify']); }

  /** Sinh key. o = { serial, type: 'machine'|'shared', machine?, issued?: Date, expires?: Date|null }; signKey = CryptoKey */
  async function makeKey(o, signKey) {
    const body = new Uint8Array(BODY), dv = new DataView(body.buffer);
    body[0] = 1; body[1] = 1;
    body[2] = o.type === 'machine' ? 1 : o.type === 'shared' ? 2 : 0;
    if (!body[2]) throw new Error('Loại key không hợp lệ');
    if (!Number.isInteger(o.serial) || o.serial < 1 || o.serial > 0xffffffff) throw new Error('Số thứ tự key không hợp lệ');
    dv.setUint32(3, o.serial);
    const iss = dayNum(o.issued || new Date());
    dv.setUint16(7, iss);
    const exp = o.expires ? dayNum(o.expires) : 0;
    if (o.expires && (exp < iss || exp > 0xffff)) throw new Error('Ngày hết hạn phải sau ngày cấp');
    dv.setUint16(9, exp);
    if (o.type === 'machine') {
      const h = o.machine instanceof Uint8Array ? o.machine : parseMachineCode(o.machine);
      if (!h) throw new Error('Mã máy không hợp lệ (13 ký tự, dạng XXXX-XXXX-XXXXX)');
      body.set(h, 11);
    }
    const sig = new Uint8Array(await subtle().sign({ name: 'Ed25519' }, signKey, body));
    return group(b32enc(concat(body, sig)));
  }
  function decodeKey(text) {
    const all = b32dec(text, BODY + SIG);
    if (!all) return null;
    const body = all.subarray(0, BODY), sig = all.subarray(BODY), dv = new DataView(all.buffer, all.byteOffset, BODY);
    if (body[0] !== 1) return null;
    const type = body[2] === 1 ? 'machine' : body[2] === 2 ? 'shared' : null;
    if (!type) return null;
    const mid = body.slice(11, 19);
    return { body, sig, product: body[1], type, serial: dv.getUint32(3), issued: dv.getUint16(7), expires: dv.getUint16(9) || 0, machine: type === 'machine' ? machineCode(mid) : null, mid };
  }
  async function checkKey(text, ctx) {
    const k = decodeKey(text);
    if (!k) return { ok: false, code: 'format' };
    let good = false;
    try { good = await subtle().verify({ name: 'Ed25519' }, await pubKey(ctx.pub), k.sig, k.body); } catch { good = false; }
    const info = { serial: k.serial, type: k.type, issued: ymd(k.issued), expires: k.expires ? ymd(k.expires) : null, machine: k.machine, expiresDay: k.expires };
    if (!good) return { ok: false, code: 'signature' };
    if (k.product !== 1) return { ok: false, code: 'product', info };
    if ((ctx.revoked || []).includes(k.serial)) return { ok: false, code: 'revoked', info };
    if (k.type === 'machine' && ctx.machine && !k.mid.every((x, i) => x === ctx.machine[i])) return { ok: false, code: 'machine', info };
    const today = dayNum(ctx.today || new Date());
    if (k.expires) { if (today > k.expires) return { ok: false, code: 'expired', info }; info.daysLeft = k.expires - today; }
    return { ok: true, code: 'ok', info };
  }
  const MESSAGES = {
    ok: 'Key hợp lệ.',
    format: 'Key không đúng định dạng. Hãy dán nguyên key (gồm các nhóm 5 ký tự).',
    signature: 'Key không hợp lệ (sai chữ ký). Kiểm tra lại từng ký tự hoặc liên hệ nơi cấp key.',
    product: 'Key này không dành cho ExcelLike.',
    machine: 'Key này cấp cho máy khác.',
    expired: 'Key đã hết hạn sử dụng.',
    revoked: 'Key này đã bị thu hồi.',
  };
  G.LIC = { b32enc, b32dec, normalize, dayNum, dayDate, ymd, machineCode, parseMachineCode, supported, importPriv, makeKey, decodeKey, checkKey, MESSAGES, b64, unb64 };
})(typeof window !== 'undefined' ? window : globalThis);
