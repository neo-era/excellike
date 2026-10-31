import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

describe('Ghi chú & chú thích', () => {
  it('đặt, sửa, xóa ghi chú; hoàn tác', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setNote(s, 1, 1, 'Kiểm tra lại đơn giá');
    expect(wb.style(s, 1, 1).nt).toBe('Kiểm tra lại đơn giá');
    wb.setNote(s, 1, 1, '');
    expect(wb.style(s, 1, 1).nt).toBeUndefined();
    wb.undo();
    expect(wb.style(s, 1, 1).nt).toBe('Kiểm tra lại đơn giá');
  });
  it('chuỗi chú thích: thêm, trả lời, giải quyết, xóa bình luận đầu xóa cả chuỗi', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.addComment(s, 0, 0, 'Lâm', 'Số này lấy từ đâu?');
    wb.addComment(s, 0, 0, 'Hùng', 'Theo hồ sơ khảo sát');
    expect(wb.style(s, 0, 0).cm!.map((x) => x.a)).toEqual(['Lâm', 'Hùng']);
    wb.resolveComment(s, 0, 0, true);
    expect(wb.style(s, 0, 0).cmDone).toBe(true);
    wb.editComment(s, 0, 0, 1, null);
    expect(wb.style(s, 0, 0).cm!.length).toBe(1);
    wb.editComment(s, 0, 0, 0, null);
    expect(wb.style(s, 0, 0).cm).toBeUndefined();
  });
  it('ghi chú đi theo ô khi chèn hàng; chổi định dạng không chép ghi chú', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 2, 0, 'x'); wb.setNote(s, 2, 0, 'ghi chú');
    wb.insertRows(s, 0, 1);
    expect(wb.style(s, 3, 0).nt).toBe('ghi chú');
    wb.paintStyles(s, { r1: 3, c1: 1, r2: 3, c2: 1 }, [[wb.style(s, 3, 0)]]);
    expect(wb.style(s, 3, 1).nt).toBeUndefined();
    expect(wb.noteCells(s)).toEqual([{ r: 3, c: 0 }]);
    expect(wb.clearNotes(s, { r1: 0, c1: 0, r2: 10, c2: 10 })).toBe(1);
    expect(wb.noteCells(s)).toEqual([]);
  });
  it('ghi/đọc .xlsx: ghi chú giữ nguyên, chú thích ghi thành ghi chú có tên người viết', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '1');
    wb.setNote(s, 0, 0, 'Đơn giá tháng 10');
    wb.addComment(s, 1, 1, 'Lâm', 'Xem lại', new Date(2026, 9, 30));
    const { wb: w2, warnings } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0];
    expect(w2.style(s2, 0, 0).nt).toBe('Đơn giá tháng 10');
    expect(w2.style(s2, 1, 1).nt).toBe('Lâm (30/10/2026):\nXem lại');
    expect(warnings.join(' ')).not.toMatch(/ghi chú/);
  });
});
