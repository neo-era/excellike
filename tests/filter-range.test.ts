import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';

describe('Lọc: vùng lọc chỉ có hàng tiêu đề (tệp .xlsx ghi A2:K2)', () => {
  const mk = () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const rows = [['Phường', 'Gói', 'SL'], ['Phú Thạnh', 'Gói 1', 14], ['', '', 1], ['Tân Phú', 'Gói 1', 6], ['Phú Thạnh', 'Gói 3', 2]];
    rows.forEach((r, i) => r.forEach((v, c) => wb.setInput(s, 2 + i, c, String(v))));
    wb.meta(s).filter = { range: { r1: 1, c1: 0, r2: 1, c2: 2 }, conds: {} }; // tiêu đề trống ở hàng 2, dữ liệu từ hàng 3
    return { wb, s };
  };
  it('danh sách giá trị lấy cả dữ liệu liền kề bên dưới', () => {
    const { wb, s } = mk();
    expect(wb.filterValues(s, 0)).toEqual(['Phú Thạnh', 'Phường', 'Tân Phú', '']);
    expect(wb.filterRange(s)).toMatchObject({ r1: 1, r2: 6 });
    expect(wb.meta(s).filter!.range.r2).toBe(1); // chỉ đọc, không đổi vùng đã lưu
  });
  it('lọc theo giá trị ẩn đúng hàng', () => {
    const { wb, s } = mk();
    wb.filterColumn(s, 0, { values: ['Phú Thạnh'] });
    expect([2, 3, 4, 5, 6].filter((r) => wb.isRowHidden(s, r))).toEqual([2, 4, 5]);
    wb.filterColumn(s, 0, null);
    expect([2, 3, 4, 5, 6].some((r) => wb.isRowHidden(s, r))).toBe(false);
  });
  it('dừng ở hàng trống hoàn toàn (như Excel)', () => {
    const { wb, s } = mk();
    wb.setInput(s, 9, 0, 'Xa phía dưới');
    expect(wb.filterValues(s, 0)).not.toContain('Xa phía dưới');
  });
});
