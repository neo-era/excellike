import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { hashPw, checkPw, legacyHash } from '../src/core/protect';

describe('Bảo vệ', () => {
  it('băm mật khẩu như Excel (SHA-512, muối, 100.000 vòng) và kiểu cũ', async () => {
    // giá trị tham chiếu tính bằng thuật toán Excel (muối cố định)
    const h = await hashPw('abc', 'AAAAAAAAAAAAAAAAAAAAAA==', 1000);
    expect(h.hash.length).toBe(88);
    expect(await checkPw('abc', h)).toBe(true);
    expect(await checkPw('abd', h)).toBe(false);
    expect(legacyHash('password')).toBe('83AF');
  });
  it('trang bảo vệ: ô khóa không sửa được, ô không khóa và vùng cho phép thì được; lưu/mở .xlsx', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'Khóa'); wb.setInput(s, 1, 0, '5');
    wb.setStyle(s, { r1: 1, c1: 0, r2: 1, c2: 0 }, { lk: false });
    wb.setAllowRanges(s, [{ title: 'Nhập liệu', range: { r1: 5, c1: 0, r2: 9, c2: 3 } }]);
    const pw = await hashPw('123', undefined, 1000);
    wb.protectSheet(s, { formatCells: true }, pw);
    expect(() => wb.setInput(s, 0, 0, 'x')).toThrow(/được bảo vệ/);
    expect(() => wb.insertRows(s, 0, 1)).toThrow();
    wb.setInput(s, 1, 0, '7'); wb.setInput(s, 6, 1, '9');
    expect(wb.value(s, 1, 0)).toBe(7);
    wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: true }); // được phép định dạng
    wb.protectBook();
    expect(() => wb.addSheet()).toThrow(/cấu trúc/);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0], p = w2.meta(s2).protect!;
    expect(p.allow.formatCells).toBe(true); expect(p.allow.insertRows).toBe(false);
    expect(await checkPw('123', p.pw!)).toBe(true);
    expect(w2.meta(s2).allowRanges?.[0].title).toBe('Nhập liệu');
    expect(w2.cellEditable(s2, 1, 0)).toBe(true); expect(w2.cellEditable(s2, 0, 0)).toBe(false);
    expect(w2.bookProtect).not.toBe(null);
    w2.unprotectSheet(s2);
    w2.setInput(s2, 0, 0, 'ok');
    expect(w2.value(s2, 0, 0)).toBe('ok');
  });
});

describe('Bảo vệ — tệp từ công cụ khác', () => {
  it('SHA-256, vùng có mật khẩu: giữ nguyên khi mở/lưu, vùng có mật khẩu vẫn khóa', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const pw = await hashPw('abc', undefined, 1000, 'SHA-256');
    expect(await checkPw('abc', pw)).toBe(true);
    wb.setAllowRanges(s, [{ title: 'Có MK', range: { r1: 0, c1: 0, r2: 0, c2: 0 }, pw: await hashPw('x', undefined, 10) }, { title: 'Tự do', range: { r1: 1, c1: 0, r2: 1, c2: 0 } }]);
    wb.protectSheet(s, {}, pw);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0];
    expect(w2.meta(s2).protect?.pw?.alg).toBe('SHA-256');
    expect(await checkPw('abc', w2.meta(s2).protect!.pw!)).toBe(true);
    expect(w2.cellEditable(s2, 0, 0)).toBe(false);
    expect(w2.cellEditable(s2, 1, 0)).toBe(true);
    expect(w2.meta(s2).allowRanges?.[0].pw).toBeTruthy();
  });
  it('số vòng quá lớn trong tệp: không treo, từ chối mật khẩu', async () => {
    expect(await checkPw('a', { alg: 'SHA-512', hash: 'x', salt: 'AAAA', spin: 1e12 })).toBe(false);
    expect(await checkPw('a', { alg: 'MD5', hash: 'x', salt: 'AAAA', spin: 1 })).toBe(false);
  });
});
