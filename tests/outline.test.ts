import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

const hidden = (wb: Workbook, s: number) => wb.meta(s).rowH.hidden().sort((a, b) => a - b);
describe('Nhóm hàng/cột', () => {
  it('nhóm, thu gọn, mở rộng, nhóm lồng giữ trạng thái, bỏ nhóm', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.groupLines(s, 'r', 1, 6);
    wb.groupLines(s, 'r', 2, 3);
    expect(wb.outlineMax(s, 'r')).toBe(2);
    expect(wb.outlineRuns(s, 'r', 2)).toEqual([{ a: 2, b: 3 }]);
    wb.collapseGroup(s, 'r', 2, 2, 3);
    expect(hidden(wb, s)).toEqual([2, 3]);
    wb.collapseGroup(s, 'r', 1, 1, 6);
    expect(hidden(wb, s)).toEqual([1, 2, 3, 4, 5, 6]);
    wb.expandGroup(s, 'r', 1, 1, 6);
    expect(hidden(wb, s)).toEqual([2, 3]); // nhóm con vẫn thu gọn
    wb.showOutlineLevel(s, 'r', 1);
    expect(hidden(wb, s)).toEqual([1, 2, 3, 4, 5, 6]);
    wb.showOutlineLevel(s, 'r', 3);
    expect(hidden(wb, s)).toEqual([]);
    wb.undo();
    expect(hidden(wb, s)).toEqual([1, 2, 3, 4, 5, 6]);
    wb.ungroupLines(s, 'r', 1, 6);
    expect(wb.outlineMax(s, 'r')).toBe(1);
    expect(() => wb.ungroupLines(s, 'r', 10, 12)).toThrow(/Không có nhóm/);
  });
  it('chèn hàng dời cấp nhóm theo', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.groupLines(s, 'c', 2, 3);
    wb.insertCols(s, 0, 2);
    expect(wb.outlineRuns(s, 'c', 1)).toEqual([{ a: 4, b: 5 }]);
  });
  it('Tổng phụ theo phường: SUBTOTAL, Tổng Cộng, đại cương 3 cấp; xóa tổng phụ', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const data = [['Phường', 'Gói', 'SL'], ['An Đông', 'G1', '5'], ['An Đông', 'G2', '7'], ['Tân Phú', 'G1', '3'], ['Tân Phú', 'G3', '4'], ['Tân Phú', 'G1', '1']];
    data.forEach((r, i) => r.forEach((v, c) => wb.setInput(s, i, c, v)));
    wb.subtotal(s, { r1: 0, c1: 0, r2: 5, c2: 2 }, { by: 0, fn: 9, cols: [2] });
    expect(wb.display(s, 3, 0).text).toBe('Tổng An Đông');
    expect(wb.value(s, 3, 2)).toBe(12);
    expect(wb.display(s, 7, 0).text).toBe('Tổng Tân Phú');
    expect(wb.value(s, 7, 2)).toBe(8);
    expect(wb.display(s, 8, 0).text).toBe('Tổng Cộng');
    expect(wb.value(s, 8, 2)).toBe(20);
    expect(wb.outlineMax(s, 'r')).toBe(2);
    wb.showOutlineLevel(s, 'r', 2);
    expect(hidden(wb, s)).toEqual([1, 2, 4, 5, 6]);
    wb.undo(); wb.undo();
    expect(wb.used(s).rows).toBe(6);
    wb.redo(); wb.redo();
    wb.removeSubtotals(s, { r1: 0, c1: 0, r2: 8, c2: 2 });
    expect(wb.used(s).rows).toBe(6);
    expect(wb.outlineMax(s, 'r')).toBe(0);
  });
  it('ghi/đọc .xlsx giữ cấp nhóm và trạng thái thu gọn', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    for (let r = 0; r < 6; r++) wb.setInput(s, r, 0, String(r));
    wb.groupLines(s, 'r', 1, 3); wb.collapseGroup(s, 'r', 1, 1, 3);
    wb.groupLines(s, 'c', 2, 3);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0];
    expect(w2.outlineRuns(s2, 'r', 1)).toEqual([{ a: 1, b: 3 }]);
    expect(w2.outlineRuns(s2, 'c', 1)).toEqual([{ a: 2, b: 3 }]);
    expect(w2.isCollapsed(s2, 'r', 1, 3)).toBe(true);
    expect(hidden(w2, s2)).toEqual([1, 2, 3]);
  });
});
