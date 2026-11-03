import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

const W = () => { const wb = Workbook.create(); return { wb, s: wb.sheetIds()[0] }; };

describe('Mảng động tràn đúng kích thước thật (như Excel)', () => {
  it('FILTER chỉ giữ số hàng có kết quả: ô ngay dưới không gây #SPILL!', () => {
    const { wb, s } = W();
    ['a', 'b', 'c', 'd', 'e'].forEach((v, i) => { wb.setInput(s, i, 0, v); wb.setInput(s, i, 1, String(i)); });
    wb.setInput(s, 3, 3, 'chặn'); // D4 — nằm trong vùng tối đa (5 hàng) nhưng ngoài kết quả thật (2 hàng)
    wb.setInput(s, 0, 3, '=FILTER(A1:A5;B1:B5>=3)');
    expect(wb.value(s, 0, 3)).toBe('d');
    expect(wb.value(s, 1, 3)).toBe('e');
    expect(wb.rawText(s, 0, 3)).toBe('=FILTER(A1:A5;B1:B5>=3)');
    expect(wb.spillRange(s, 0, 3)?.range).toEqual({ r1: 0, c1: 3, r2: 1, c2: 3 });
  });
  it('dữ liệu đổi → vùng tràn co/giãn; hoàn tác trả về như cũ', () => {
    const { wb, s } = W();
    [1, 5, 2, 8].forEach((v, i) => wb.setInput(s, i, 0, String(v)));
    wb.setInput(s, 0, 2, '=FILTER(A1:A4;A1:A4>4)');
    expect([wb.value(s, 0, 2), wb.value(s, 1, 2), wb.value(s, 2, 2)]).toEqual([5, 8, null]);
    wb.setInput(s, 0, 0, '9');
    expect([wb.value(s, 0, 2), wb.value(s, 1, 2), wb.value(s, 2, 2)]).toEqual([9, 5, 8]);
    wb.setInput(s, 1, 0, '0'); wb.setInput(s, 3, 0, '0');
    expect([wb.value(s, 0, 2), wb.value(s, 1, 2)]).toEqual([9, null]);
    wb.undo(); wb.undo(); wb.undo();
    expect([wb.value(s, 0, 2), wb.value(s, 1, 2), wb.value(s, 2, 2)]).toEqual([5, 8, null]);
    wb.redo();
    expect([wb.value(s, 0, 2), wb.value(s, 1, 2), wb.value(s, 2, 2)]).toEqual([9, 5, 8]);
    expect(wb.canRedo()).toBe(true);
  });
  it('TOROW(FILTER(HSTACK(...))) như tệp đề xuất đèn LED: tràn ngang đúng số cột', () => {
    const { wb, s } = W();
    const rows = [['P1', 'Carina', 'TM', 2], ['P2', 'Magnolia', 'BL', 3], ['P1', 'Nova', 'TM', 1], ['P1', 'Iris', 'X', 4]];
    rows.forEach((r, i) => r.forEach((v, j) => wb.setInput(s, 10 + i, j, String(v))));
    wb.setInput(s, 0, 0, 'P1');
    wb.setInput(s, 0, 8, 'kế bên'); // I1
    wb.setInput(s, 0, 1, '=TOROW(FILTER(HSTACK(B11:B14&"/"&C11:C14;D11:D14);A11:A14=A1;""))');
    const vals = Array.from({ length: 7 }, (_, j) => wb.value(s, 0, 1 + j));
    expect(vals).toEqual(['Carina/TM', 2, 'Nova/TM', 1, 'Iris/X', 4, null]);
    expect(wb.value(s, 0, 8)).toBe('kế bên');
  });
  it('FILTER không có kết quả dùng giá trị nếu_rỗng (1 ô)', () => {
    const { wb, s } = W();
    wb.setInput(s, 0, 0, '1'); wb.setInput(s, 1, 0, '2');
    wb.setInput(s, 0, 1, '=FILTER(A1:A2;A1:A2>5;"không có")');
    expect(wb.value(s, 0, 1)).toBe('không có');
    expect(wb.spillRange(s, 0, 1)).toBe(null);
  });
  it('lưu .xlsx: không có XFIT, ref đúng kích thước; mở lại vẫn đúng', async () => {
    const { wb, s } = W();
    [1, 5, 2, 8].forEach((v, i) => wb.setInput(s, i, 0, String(v)));
    wb.setInput(s, 0, 2, '=FILTER(A1:A4;A1:A4>1)');
    const buf = await writeXlsx(wb);
    const JSZip = (await import('jszip')).default;
    const x = await (await JSZip.loadAsync(buf)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(x).not.toMatch(/XFIT/);
    expect(x).toMatch(/ref="C1:C3"/);
    const { wb: w2 } = await readXlsx(buf);
    const s2 = w2.sheetIds()[0];
    expect([w2.value(s2, 0, 2), w2.value(s2, 2, 2), w2.value(s2, 3, 2)]).toEqual([5, 8, null]);
    expect(w2.canUndo()).toBe(false);
  });
});
