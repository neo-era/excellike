import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';

const mk = () => {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  [['Cát', 3, 'm3'], ['Đá', 1, 'm3'], ['Thép', 2, 'kg'], ['Cát', 5, 'm3']].forEach((row, r) => row.forEach((v, c) => wb.setInput(s, r, c, String(v))));
  return { wb, s };
};
const block = (wb: Workbook, s: number, r: number, c: number, h: number, w: number) => Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => wb.value(s, r + i, c + j)));
describe('Mảng động', () => {
  it('FILTER 2 chiều theo cột điều kiện, nếu rỗng', () => {
    const { wb, s } = mk();
    wb.setInput(s, 0, 5, '=FILTER(A1:C4;B1:B4>=2)');
    expect(block(wb, s, 0, 5, 3, 3)).toEqual([['Cát', 3, 'm3'], ['Thép', 2, 'kg'], ['Cát', 5, 'm3']]);
    expect(wb.value(s, 3, 5)).toBe(null);
    wb.setInput(s, 6, 5, '=FILTER(A1:A4;C1:C4="tấn";"không có")');
    expect(wb.value(s, 6, 5)).toBe('không có');
    wb.setInput(s, 8, 5, '=FILTER(A1:C4;(C1:C4="m3")*(B1:B4>2))');
    expect(block(wb, s, 8, 5, 2, 2)).toEqual([['Cát', 3], ['Cát', 5]]);
  });
  it('SORT, SORTBY, UNIQUE, TAKE, DROP, CHOOSECOLS, TOCOL', () => {
    const { wb, s } = mk();
    wb.setInput(s, 0, 5, '=SORT(A1:B4;2;-1)');
    expect(block(wb, s, 0, 5, 4, 2)).toEqual([['Cát', 5], ['Cát', 3], ['Thép', 2], ['Đá', 1]]);
    wb.setInput(s, 0, 8, '=SORTBY(A1:A4;B1:B4)');
    expect(block(wb, s, 0, 8, 4, 1).flat()).toEqual(['Đá', 'Thép', 'Cát', 'Cát']);
    wb.setInput(s, 0, 10, '=UNIQUE(A1:A4)');
    expect(block(wb, s, 0, 10, 3, 1).flat()).toEqual(['Cát', 'Đá', 'Thép']);
    wb.setInput(s, 6, 0, '=TAKE(A1:C4;2;-1)');
    expect(block(wb, s, 6, 0, 2, 1).flat()).toEqual(['m3', 'm3']);
    wb.setInput(s, 9, 0, '=DROP(A1:C4;3;1)');
    expect(block(wb, s, 9, 0, 1, 2)).toEqual([[5, 'm3']]);
    wb.setInput(s, 11, 0, '=CHOOSECOLS(A1:C4;3;1)');
    expect(block(wb, s, 11, 0, 1, 2)).toEqual([['m3', 'Cát']]);
    wb.setInput(s, 16, 0, '=TOCOL(A1:B2)');
    expect(block(wb, s, 16, 0, 4, 1).flat()).toEqual(['Cát', 3, 'Đá', 1]);
    wb.setInput(s, 21, 0, '=SUM(FILTER(B1:B4;A1:A4="Cát"))');
    expect(wb.value(s, 21, 0)).toBe(8);
  });
  it('hằng mảng nhiều cột {1;2,5|3;4} và #SPILL!', () => {
    const { wb, s } = mk();
    wb.setInput(s, 0, 5, '={1;2,5|3;4}');
    expect(block(wb, s, 0, 5, 2, 2)).toEqual([[1, 2.5], [3, 4]]);
    wb.setInput(s, 3, 5, '=SUM({1;2|3;4})');
    expect(wb.value(s, 3, 5)).toBe(10);
    wb.setInput(s, 6, 7, 'x');
    wb.setInput(s, 5, 6, '=SEQUENCE(3)');
    wb.setInput(s, 5, 7, '=SEQUENCE(3)');
    expect((wb.value(s, 5, 7) as any).value).toBe('#SPILL!');
  });
});
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import JSZip from 'jszip';
describe('Mảng động trong .xlsx', () => {
  it('ghi t="array" + cm="1" + metadata; đọc lại không bị #SPILL!', async () => {
    const { wb, s } = mk();
    wb.setInput(s, 0, 5, '=FILTER(A1:C4;B1:B4>=2)');
    wb.setInput(s, 6, 0, '=SORT(A1:A4)');
    const buf = await writeXlsx(wb);
    const z = await JSZip.loadAsync(buf);
    const x = await z.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(x).toMatch(/<c r="F1"[^>]* cm="1"><f t="array" ref="F1:H4">_xlfn\._xlws\.FILTER\(/);
    expect(await z.file('xl/metadata.xml')?.async('string')).toMatch(/XLDAPR/);
    // giả lập tệp Excel: ô tràn có giá trị lưu sẵn
    const x2 = x.replace(/<c r="G1"[^>]*\/>|<c r="G1"[^>]*>[\s\S]*?<\/c>/, '<c r="G1"><v>3</v></c>');
    z.file('xl/worksheets/sheet1.xml', x2);
    const { wb: w2 } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    const s2 = w2.sheetIds()[0];
    expect(w2.value(s2, 0, 6)).toBe(3);
    expect(w2.rawText(s2, 0, 5)).toBe('=FILTER(A1:C4;B1:B4>=2)');
    expect(w2.value(s2, 2, 5)).toBe('Cát');
  });
});
