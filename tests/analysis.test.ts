import { describe, it, expect } from 'vitest';
import { Workbook, critTest } from '../src/core/workbook';

describe('Phân tích Giả định', () => {
  it('Goal Seek: tìm khối lượng để thành tiền = 1.000.000, hoàn tác 1 bước', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '10'); wb.setInput(s, 0, 1, '25000'); wb.setInput(s, 0, 2, '=A1*B1*1,1');
    const r = wb.goalSeek(s, { r: 0, c: 2 }, 1_000_000, { r: 0, c: 0 });
    expect(r.ok).toBe(true);
    expect(Math.abs(r.x - 1_000_000 / 27500)).toBeLessThan(1e-4);
    expect(wb.value(s, 0, 0)).toBe(10); // chưa ghi
    wb.setInput(s, 0, 0, String(r.x).replace('.', ','));
    expect(Math.abs((wb.value(s, 0, 2) as number) - 1_000_000)).toBeLessThan(0.01);
    wb.undo();
    expect(wb.value(s, 0, 0)).toBe(10);
    expect(wb.canUndo()).toBe(true);
    // phi tuyến
    wb.setInput(s, 5, 0, '1'); wb.setInput(s, 5, 1, '=A6^3+A6');
    const q = wb.goalSeek(s, { r: 5, c: 1 }, 30, { r: 5, c: 0 });
    expect(q.ok).toBe(true); expect(Math.abs(q.x - 3)).toBeLessThan(1e-3);
  });
  it('Bảng Dữ liệu 1 biến và 2 biến', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '2'); wb.setInput(s, 0, 1, '3'); // A1 lãi, B1 năm
    // 1 biến theo cột: D2:D4 giá trị, E1 công thức
    wb.setInput(s, 0, 4, '=A1*10'); ['1', '2', '3'].forEach((v, i) => wb.setInput(s, 1 + i, 3, v));
    wb.dataTable(s, { r1: 0, c1: 3, r2: 3, c2: 4 }, null, { r: 0, c: 0 });
    expect([wb.value(s, 1, 4), wb.value(s, 2, 4), wb.value(s, 3, 4)]).toEqual([10, 20, 30]);
    expect(wb.value(s, 0, 0)).toBe(2);
    // 2 biến: G1 = A1*B1, hàng G1:I1 giá trị B, cột G2:G3 giá trị A
    wb.setInput(s, 0, 6, '=A1*B1'); wb.setInput(s, 0, 7, '10'); wb.setInput(s, 0, 8, '20'); wb.setInput(s, 1, 6, '1'); wb.setInput(s, 2, 6, '2');
    wb.dataTable(s, { r1: 0, c1: 6, r2: 2, c2: 8 }, { r: 0, c: 1 }, { r: 0, c: 0 });
    expect([wb.value(s, 1, 7), wb.value(s, 1, 8), wb.value(s, 2, 7), wb.value(s, 2, 8)]).toEqual([10, 20, 20, 40]);
    wb.undo();
    expect(wb.value(s, 1, 7)).toBe(null);
  });
  it('Kịch bản: hiện và tóm tắt', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '100'); wb.setInput(s, 0, 1, '=A1*2');
    wb.setScenarios(s, [{ name: 'Thấp', cells: [[0, 0]], values: ['50'] }, { name: 'Cao', cells: [[0, 0]], values: ['300'] }]);
    const id = wb.scenarioSummary(s, [{ r: 0, c: 1 }]);
    expect(wb.value(id, 1, 2)).toBe('Thấp');
    expect(wb.value(id, 5, 2)).toBe(100);
    expect(wb.value(id, 5, 3)).toBe(600);
    expect(wb.value(s, 0, 0)).toBe(100);
    wb.showScenario(s, wb.scenarios(s)[1]);
    expect(wb.value(s, 0, 1)).toBe(600);
  });
});
describe('Lọc Nâng cao, Hợp nhất', () => {
  const mk = () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    [['Vật tư', 'KL', 'Đơn vị'], ['Cát', 5, 'm3'], ['Đá', 2, 'm3'], ['Cát vàng', 8, 'm3'], ['Thép', 9, 'kg'], ['Cát', 5, 'm3']].forEach((row, r) => row.forEach((v, c) => wb.setInput(s, r, c, String(v))));
    return { wb, s };
  };
  it('điều kiện VÀ/HOẶC, ký tự đại diện, sao chép, duy nhất', () => {
    const { wb, s } = mk();
    [['Vật tư', 'KL'], ['Cát*', '>4'], ['Thép', '']].forEach((row, r) => row.forEach((v, c) => wb.setInput(s, 10 + r, c, v)));
    const n = wb.advancedFilter(s, { r1: 0, c1: 0, r2: 5, c2: 2 }, { r1: 10, c1: 0, r2: 12, c2: 1 }, { to: { s, r: 20, c: 0 }, unique: true });
    expect(n).toBe(3);
    expect([wb.value(s, 21, 0), wb.value(s, 22, 0), wb.value(s, 23, 0)]).toEqual(['Cát', 'Cát vàng', 'Thép']);
    const m = wb.advancedFilter(s, { r1: 0, c1: 0, r2: 5, c2: 2 }, { r1: 10, c1: 0, r2: 11, c2: 1 }, {});
    expect(m).toBe(3);
    expect(wb.meta(s).rowH.hidden().sort()).toEqual([2, 4]);
    expect(critTest('<>m3')(null, 'kg')).toBe(true);
    expect(critTest('=Cát')(null, 'Cát vàng')).toBe(false);
  });
  it('Hợp nhất theo nhãn', () => {
    const { wb, s } = mk();
    wb.consolidate({ s, r: 0, c: 5 }, [{ sheet: s, range: { r1: 0, c1: 0, r2: 5, c2: 1 } }], 'sum', true, true);
    expect(wb.value(s, 0, 6)).toBe('KL');
    expect([wb.value(s, 1, 5), wb.value(s, 1, 6)]).toEqual(['Cát', 10]);
    expect(wb.value(s, 4, 5)).toBe('Thép');
  });
});

describe('Phân tích Giả định — lịch sử', () => {
  it('Bảng Dữ liệu lớn (350 hàng, 20×20) không làm hỏng ô nhập và lịch sử hoàn tác', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 5, '7'); wb.setInput(s, 0, 6, '3'); // F1, G1
    wb.setInput(s, 0, 9, 'x'); // J1: bước đầu của người dùng
    wb.setInput(s, 1, 1, '=F2*2'); // B2 công thức (1 biến theo cột, ô nhập F2)
    wb.setInput(s, 1, 5, '1');
    wb.writeCells(s, Array.from({ length: 350 }, (_, i) => [2 + i, 0, String(i)] as [number, number, string]));
    wb.dataTable(s, { r1: 1, c1: 0, r2: 351, c2: 1 }, null, { r: 1, c: 5 });
    expect(wb.value(s, 1, 5)).toBe(1);
    expect(wb.value(s, 351, 1)).toBe(698);
    wb.setInput(s, 400, 0, '=F1*G1');
    wb.writeCells(s, Array.from({ length: 20 }, (_, i) => [[401 + i, 0, String(i)], [400, 1 + i, String(i)]] as [number, number, string][]).flat());
    wb.dataTable(s, { r1: 400, c1: 0, r2: 420, c2: 20 }, { r: 0, c: 6 }, { r: 0, c: 5 });
    expect([wb.value(s, 0, 5), wb.value(s, 0, 6)]).toEqual([7, 3]);
    expect(wb.value(s, 420, 20)).toBe(361);
    for (let i = 0; i < 500 && wb.canUndo(); i++) wb.undo();
    expect(wb.value(s, 0, 9)).toBe(null);
  });
});
