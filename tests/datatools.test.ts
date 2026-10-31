import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';

const mk = (rows: (string | number)[][]) => {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(s, r, c, String(v))));
  return { wb, s };
};
const col = (wb: Workbook, s: number, c: number, n: number) => Array.from({ length: n }, (_, r) => wb.value(s, r, c));

describe('Loại bỏ Trùng lặp', () => {
  it('giữ lần đầu, không phân biệt hoa/thường, dồn ô lên, hoàn tác được', () => {
    const { wb, s } = mk([['Vật tư', 'ĐG'], ['Cát', 10], ['cát', 10], ['Đá', 5], ['Cát', 11], ['Đá', 5]]);
    const r = wb.removeDuplicates(s, { r1: 0, c1: 0, r2: 5, c2: 1 }, [0, 1], true);
    expect(r).toEqual({ removed: 2, kept: 3 });
    expect(col(wb, s, 0, 6)).toEqual(['Vật tư', 'Cát', 'Đá', 'Cát', null, null]);
    expect(col(wb, s, 1, 4)).toEqual(['ĐG', 10, 5, 11]);
    wb.undo();
    expect(col(wb, s, 0, 6)).toEqual(['Vật tư', 'Cát', 'cát', 'Đá', 'Cát', 'Đá']);
  });
  it('so theo một cột; công thức được dời đúng hàng', () => {
    const { wb, s } = mk([['A', 1], ['A', 2], ['B', 3]]);
    wb.setInput(s, 2, 2, '=B3*2');
    wb.removeDuplicates(s, { r1: 0, c1: 0, r2: 2, c2: 2 }, [0], false);
    expect(col(wb, s, 0, 3)).toEqual(['A', 'B', null]);
    expect(wb.rawText(s, 1, 2)).toBe('=B2*2');
    expect(wb.value(s, 1, 2)).toBe(6);
  });
});

describe('Văn bản thành Cột', () => {
  it('tách theo dấu chấm phẩy, nhận số, ngoặc kép và dấu liền nhau', () => {
    const { wb, s } = mk([['Cát;12,5;"m3; khối"'], ['Đá;;7']]);
    const n = wb.textToColumns(s, 0, 0, 1, { delims: ';' });
    expect(n).toBe(3);
    expect([wb.value(s, 0, 0), wb.value(s, 0, 1), wb.value(s, 0, 2)]).toEqual(['Cát', 12.5, 'm3; khối']);
    expect([wb.value(s, 1, 0), wb.value(s, 1, 1), wb.value(s, 1, 2)]).toEqual(['Đá', null, 7]);
    wb.undo();
    expect(wb.value(s, 0, 0)).toBe('Cát;12,5;"m3; khối"');
    wb.textToColumns(s, 0, 1, 1, { delims: ';', consecutive: true });
    expect([wb.value(s, 1, 0), wb.value(s, 1, 1)]).toEqual(['Đá', 7]);
  });
  it('độ rộng cố định, kiểu Văn bản giữ số 0 đầu, bỏ qua cột, ô đích khác', () => {
    const { wb, s } = mk([['0012ABCxyz']]);
    wb.textToColumns(s, 0, 0, 0, { widths: [4, 7], types: ['t', 'x', 'g'], dest: { r: 0, c: 2 } });
    expect(wb.value(s, 0, 2)).toBe('0012');
    expect(wb.value(s, 0, 3)).toBe('xyz');
    expect(wb.value(s, 0, 0)).toBe('0012ABCxyz');
  });
});
