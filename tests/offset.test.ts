import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { offsetIn, offsetOut } from '../src/core/offset';

const table = (wb: Workbook, s: number) => {
  // bảng tra nội suy (như THKP hạng mục): hàng 11 = mốc, hàng 12 = tỷ lệ
  wb.setInput(s, 10, 10, 'Mốc'); [40, 60, 100, 300].forEach((v, i) => wb.setInput(s, 10, 11 + i, String(v)));
  wb.setInput(s, 11, 10, 'Tỷ lệ'); ['5,5', '5,3', '5', '4,8'].forEach((v, i) => wb.setInput(s, 11, 11 + i, v));
};
const F = '=IF(R11<=40;L12;IF(R11>=300;O12;FORECAST(R11;OFFSET(K12;0;MATCH(R11;L11:O11;1);1;2);OFFSET(K11;0;MATCH(R11;L11:O11;1);1;2))))';

describe('OFFSET với đối số động', () => {
  it('đổi qua lại giữ nguyên công thức người dùng', () => {
    const inF = offsetIn(F, { r: 11, c: 17 });
    expect(inF).toMatch(/XOFS5\(K12;0;MATCH\(R11;L11:O11;1\);1;2;\$K\$11:\$P\$12\)/);
    expect(offsetOut(inF)).toBe(F);
    expect(offsetIn('=OFFSET(A1;1;2)')).toBe('=OFFSET(A1;1;2)'); // tĩnh: để HyperFormula xử lý
    expect(offsetIn('="OFFSET(A1;B1;1)"')).toBe('="OFFSET(A1;B1;1)"'); // trong chuỗi
    expect(offsetOut(offsetIn('=SUM(OFFSET(Sheet2!A1;B1;0;3))'))).toBe('=SUM(OFFSET(Sheet2!A1;B1;0;3))');
  });

  it('tính đúng như Excel (FORECAST + OFFSET + MATCH) và tính lại khi đổi dữ liệu', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    table(wb, s);
    wb.setInput(s, 10, 17, '80');
    wb.setInput(s, 11, 17, F);
    expect(wb.rawText(s, 11, 17)).toBe(F);
    expect(wb.value(s, 11, 17)).toBeCloseTo(5.15, 9); // 5,3 + (5−5,3)·(80−60)/(100−60)
    wb.setInput(s, 10, 17, '29,393588923');
    expect(wb.value(s, 11, 17)).toBe(5.5);
    wb.setInput(s, 10, 17, '200');
    expect(wb.value(s, 11, 17)).toBeCloseTo(4.9, 9);
    wb.setInput(s, 11, 13, '6'); // đổi bảng tra → ô tự tính lại
    expect(wb.value(s, 11, 17)).toBeCloseTo(5.4, 9);
    // OFFSET ra ô đơn / vùng trong SUM
    wb.setInput(s, 0, 0, '2'); wb.setInput(s, 0, 1, '=OFFSET(K11;0;A1)'); wb.setInput(s, 0, 2, '=SUM(OFFSET(L12;0;0;1;A1))');
    expect(wb.value(s, 0, 1)).toBe(60);
    expect(wb.value(s, 0, 2)).toBeCloseTo(5.5 + 5.3, 9);
    wb.setInput(s, 0, 0, '9'); // ra ngoài bảng → ô trống = 0
    expect(wb.value(s, 0, 1)).toBe(0);
  });

  it('chèn hàng dời tham chiếu, hoàn tác, lưu .xlsx vẫn là OFFSET', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    table(wb, s); wb.setInput(s, 10, 17, '80'); wb.setInput(s, 11, 17, F);
    wb.insertRows(s, 0, 2);
    expect(wb.rawText(s, 13, 17)).toBe(F.replace(/R11/g, 'R13').replace(/K12/g, 'K14').replace(/K11/g, 'K13').replace(/L11:O11/g, 'L13:O13').replace(/L12/g, 'L14').replace(/O12/g, 'O14'));
    expect(wb.value(s, 13, 17)).toBeCloseTo(5.15, 9);
    wb.undo();
    expect(wb.rawText(s, 11, 17)).toBe(F);
    const x = new ExcelJS.Workbook(); await x.xlsx.load((await writeXlsx(wb)) as any);
    const v: any = x.worksheets[0].getCell('R12').value;
    expect(v.formula).toContain('OFFSET(K12,0,MATCH(R11,L11:O11,1),1,2)');
    expect(v.formula).not.toContain('XOFS');
    const back = (await readXlsx(await writeXlsx(wb))).wb;
    expect(back.value(back.sheetIds()[0], 11, 17)).toBeCloseTo(5.15, 9);
  });

  it('tính lặp: ô phụ thuộc vòng lặp qua IFERROR cũng lấy giá trị hội tụ', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    // A1 = B1/2 + 10 ; B1 = A1 → vòng lặp, hội tụ A1 = 20
    wb.setInput(s, 0, 0, '=B1/2+10'); wb.setInput(s, 0, 1, '=A1');
    // C1 bọc IFERROR (HyperFormula không thấy #CYCLE!) — Excel cho 20, không phải 0
    wb.setInput(s, 0, 2, '=IFERROR(A1;0)'); wb.setInput(s, 0, 3, '=C1*2');
    wb.setIterate({ on: true, count: 200, delta: 1e-9 });
    expect(wb.value(s, 0, 0)).toBeCloseTo(20, 6);
    expect(wb.value(s, 0, 2)).toBeCloseTo(20, 6);
    expect(wb.value(s, 0, 3)).toBeCloseTo(40, 6);
    wb.setInput(s, 0, 2, '7'); // không còn phụ thuộc → giá trị nhập
    expect(wb.value(s, 0, 2)).toBe(7);
    expect(wb.value(s, 0, 3)).toBe(14);
  });
});
