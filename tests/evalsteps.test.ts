import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { nextStep, applyStep } from '../src/core/evalsteps';

describe('Đánh giá Công thức', () => {
  it('thay tham chiếu rồi tính hàm trong cùng, cuối cùng ra kết quả', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '2'); wb.setInput(s, 1, 0, '3'); wb.setInput(s, 2, 0, '1,5');
    let f = 'ROUND(SUM(A1:A2)*A3;1)+1';
    const seen: string[] = [];
    for (let i = 0; i < 20; i++) { const st = nextStep(wb, s, f); if (!st) break; f = applyStep(wb, s, f, st); seen.push(f); }
    expect(seen[0]).toBe('ROUND(SUM({2|3})*A3;1)+1');
    expect(seen[1]).toBe('ROUND(5*A3;1)+1');
    expect(seen[2]).toBe('ROUND(5*1,5;1)+1');
    expect(seen[3]).toBe('7,5+1');
    expect(f).toBe('8,5');
  });
  it('chữ, lỗi và vùng lớn', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'Cát'); wb.setInput(s, 0, 1, '0');
    let f = 'IFERROR(1/B1;A1&"!")';
    for (let i = 0; i < 10; i++) { const st = nextStep(wb, s, f); if (!st) break; f = applyStep(wb, s, f, st); }
    expect(f).toBe('"Cát!"');
    f = 'SUM(A1:A5000)';
    const st = nextStep(wb, s, f)!;
    expect(st.kind).toBe('call');
  });
});

describe('Đánh giá Công thức — ca rà soát', () => {
  it('ô trống trong phép ghép chuỗi là "" và ROW giữ tham chiếu', async () => {
    const { Workbook } = await import('../src/core/workbook');
    const { nextStep, applyStep } = await import('../src/core/evalsteps');
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    let f = 'A5&"x"';
    for (let i = 0; i < 10; i++) { const st = nextStep(wb, s, f); if (!st) break; f = applyStep(wb, s, f, st); }
    expect(f).toBe('"x"');
    f = 'ROW(B7)+1';
    for (let i = 0; i < 10; i++) { const st = nextStep(wb, s, f); if (!st) break; f = applyStep(wb, s, f, st); }
    expect(f).toBe('8');
  });
});
