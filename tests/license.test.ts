import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../electron/license.cjs');
const CFG = require('../electron/license-config.cjs');

// khóa thử riêng cho test (không phải khóa thật)
const K = L.newKeyPair();
const ctx = (o: object = {}) => ({ pub: K.pub, ...o });
const M1 = L.machineHash('EB8490F1-5982-446C-8A60-08487FC2501C');
const M2 = L.machineHash('00000000-1111-2222-3333-444444444444');

describe('Key bản quyền', () => {
  it('Base32 Crockford khứ hồi, chịu chữ thường/gạch/khoảng trắng, O→0, I/L→1', () => {
    const b = Buffer.from([0, 1, 2, 250, 255, 128, 77]);
    const s = L.b32enc(b);
    expect(L.b32dec(s, 7).equals(b)).toBe(true);
    expect(L.b32dec(s.toLowerCase().split('').join(' - '), 7).equals(b)).toBe(true);
    expect(L.normalize('o-i-l')).toBe('011');
    expect(L.b32dec('U', 1)).toBeNull();
  });
  it('Mã máy: 13 ký tự XXXX-XXXX-XXXXX, không phân biệt hoa thường của GUID', () => {
    const c = L.machineCode(M1);
    expect(c).toMatch(/^[0-9A-Z]{4}-[0-9A-Z]{4}-[0-9A-Z]{5}$/);
    expect(L.machineCode(L.machineHash('eb8490f1-5982-446c-8a60-08487fc2501c'))).toBe(c);
    expect(L.parseMachineCode(c.toLowerCase()).equals(M1)).toBe(true);
  });
  it('khớp với lic.exe trong bộ cài (cùng GUID → cùng Mã máy, đã đối chiếu dưới Wine)', () => {
    expect(L.machineCode(L.machineHash('eb8490f1-5982-446c-8a60-08487fc2501c'))).toBe('RKKG-MH8R-7Y5B2');
  });
  it('key dùng mọi máy, vĩnh viễn: hợp lệ trên máy bất kỳ', () => {
    const k = L.makeKey({ serial: 1, type: 'shared' }, K.priv);
    expect(k).toMatch(/^([0-9A-Z]{5}-)+[0-9A-Z]{1,5}$/);
    const r = L.checkKey(k, ctx({ machine: M2 }));
    expect(r.ok).toBe(true);
    expect(r.info).toMatchObject({ serial: 1, type: 'shared', expires: null });
  });
  it('key gắn máy: đúng máy hợp lệ, máy khác báo "machine"', () => {
    const k = L.makeKey({ serial: 2, type: 'machine', machine: L.machineCode(M1) }, K.priv);
    expect(L.checkKey(k, ctx({ machine: M1 })).ok).toBe(true);
    expect(L.checkKey(k, ctx({ machine: M2 })).code).toBe('machine');
    expect(L.decodeKey(k).machine).toBe(L.machineCode(M1));
  });
  it('hạn dùng: còn hạn đến hết ngày hết hạn, qua ngày là hết hạn; còn bao nhiêu ngày', () => {
    const k = L.makeKey({ serial: 3, type: 'shared', issued: new Date(2026, 9, 1), expires: new Date(2026, 11, 31) }, K.priv);
    expect(L.checkKey(k, ctx({ today: new Date(2026, 11, 31, 23, 59) })).ok).toBe(true);
    expect(L.checkKey(k, ctx({ today: new Date(2026, 11, 21) })).info.daysLeft).toBe(10);
    expect(L.checkKey(k, ctx({ today: new Date(2027, 0, 1) })).code).toBe('expired');
  });
  it('chỉnh lùi đồng hồ không kéo dài được key có hạn', () => {
    const k = L.makeKey({ serial: 4, type: 'shared', issued: new Date(2026, 0, 1), expires: new Date(2026, 5, 30) }, K.priv);
    const seen = L.dayNum(new Date(2026, 6, 15));
    expect(L.checkKey(k, ctx({ today: new Date(2026, 5, 1), lastSeen: seen })).code).toBe('clock');
    expect(L.checkKey(k, ctx({ today: new Date(2026, 6, 14), lastSeen: seen })).code).toBe('expired');
    // key vĩnh viễn không bị ảnh hưởng bởi đồng hồ
    const p = L.makeKey({ serial: 5, type: 'shared' }, K.priv);
    expect(L.checkKey(p, ctx({ today: new Date(2020, 0, 5), lastSeen: seen })).ok).toBe(true);
  });
  it('sửa một ký tự → sai chữ ký; chuỗi rác → sai định dạng; khóa ký khác → sai chữ ký', () => {
    const k = L.makeKey({ serial: 6, type: 'shared' }, K.priv);
    const i = 12, ch = k[i] === 'A' ? 'B' : 'A';
    expect(L.checkKey(k.slice(0, i) + ch + k.slice(i + 1), ctx()).code).toBe('signature');
    expect(L.checkKey('ABC', ctx()).code).toBe('format');
    expect(L.checkKey('', ctx()).code).toBe('format');
    const other = L.newKeyPair();
    expect(L.checkKey(L.makeKey({ serial: 7, type: 'shared' }, other.priv), ctx()).code).toBe('signature');
  });
  it('key đã thu hồi', () => {
    const k = L.makeKey({ serial: 99, type: 'shared' }, K.priv);
    expect(L.checkKey(k, ctx({ revoked: [99] })).code).toBe('revoked');
  });
  it('báo lỗi khi sinh key sai dữ liệu', () => {
    expect(() => L.makeKey({ serial: 1, type: 'machine', machine: 'SAI' }, K.priv)).toThrow(/Mã máy/);
    expect(() => L.makeKey({ serial: 0, type: 'shared' }, K.priv)).toThrow();
    expect(() => L.makeKey({ serial: 1, type: 'shared', issued: new Date(2026, 5, 1), expires: new Date(2026, 0, 1) }, K.priv)).toThrow(/hết hạn/);
  });
  it('khóa công khai trong cấu hình là 32 byte, không có khóa bí mật trong mã nguồn app', () => {
    expect(Buffer.from(CFG.PUBLIC_KEY, 'base64').length).toBe(32);
    expect(Object.keys(CFG)).toEqual(['PUBLIC_KEY', 'REVOKED']);
    expect(L.pubFromPriv(K.priv)).toBe(K.pub);
  });
});
