import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { writeXlsx } from '../src/core/xlsx';

describe('Rà soát 0.9', () => {
  it('xóa trang nguồn của sparkline vẫn lưu .xlsx được', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const d = wb.addSheet('Data'); for (let i = 0; i < 4; i++) wb.setInput(d, 0, i, String(i + 1));
    wb.addSparklines(s, 'line', { sheet: d, range: { r1: 0, c1: 0, r2: 0, c2: 3 } }, { r1: 0, c1: 0, r2: 0, c2: 0 });
    wb.removeSheet(d);
    expect(wb.meta(s).sparks?.length ?? 0).toBe(0);
    await expect(writeXlsx(wb)).resolves.toBeTruthy();
    wb.undo();
    expect(wb.meta(s).sparks?.length).toBe(1);
  });
  it('sparkline: chèn hàng ở trang nguồn dời vùng nguồn; chèn ở trang chứa không dời', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const d = wb.addSheet('Data');
    wb.addSparklines(s, 'line', { sheet: d, range: { r1: 2, c1: 0, r2: 2, c2: 3 } }, { r1: 5, c1: 0, r2: 5, c2: 0 });
    wb.insertRows(s, 0, 1);
    let it = wb.meta(s).sparks![0].items[0];
    expect([it.r, it.src.range.r1]).toEqual([6, 2]);
    wb.insertRows(d, 0, 2);
    it = wb.meta(s).sparks![0].items[0];
    expect([it.r, it.src.range.r1]).toEqual([6, 4]);
    wb.undo();
    expect(wb.meta(s).sparks![0].items[0].src.range.r1).toBe(2);
  });
  it('vùng cho phép sửa dời theo khi chèn hàng', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setAllowRanges(s, [{ title: 'V', range: { r1: 0, c1: 0, r2: 0, c2: 0 } }]);
    wb.insertRows(s, 0, 1);
    expect(wb.meta(s).allowRanges![0].range.r1).toBe(1);
    wb.protectSheet(s, {});
    expect(wb.cellEditable(s, 1, 0)).toBe(true);
    expect(wb.cellEditable(s, 0, 0)).toBe(false);
  });
  it('các đường sửa khác bị chặn khi trang được bảo vệ', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'locked'); wb.setStyle(s, { r1: 1, c1: 0, r2: 1, c2: 0 }, { lk: false });
    wb.protectSheet(s, { formatCells: true });
    expect(() => wb.pasteText(s, 0, 0, 'hacked')).toThrow();
    wb.copy(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, true);
    expect(() => wb.paste(s, 1, 0)).toThrow();
    expect(() => wb.insertCells(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, 'down')).toThrow();
    expect(() => wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { lk: false })).toThrow();
    expect(() => wb.setAllowRanges(s, [{ title: 'x', range: { r1: 0, c1: 0, r2: 9, c2: 9 } }])).toThrow();
    expect(wb.value(s, 0, 0)).toBe('locked');
    wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: true }); // định dạng ô vẫn được (đã cho phép)
    expect(wb.style(s, 0, 0).b).toBe(true);
  });
});
