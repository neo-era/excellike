import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';

const col = (wb: Workbook, s: number, c: number, n: number) => Array.from({ length: n }, (_, r) => wb.value(s, r, c));

describe('Kéo viền vùng chọn', () => {
  it('di chuyển, công thức tham chiếu đi theo; hoàn tác 1 bước', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '5'); wb.setInput(s, 0, 1, '=A1*2');
    wb.dragCells(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, 3, 0, 'move');
    expect(wb.value(s, 0, 0)).toBe(null); expect(wb.value(s, 3, 0)).toBe(5);
    expect(wb.rawText(s, 0, 1)).toBe('=A4*2');
    wb.undo();
    expect(wb.value(s, 0, 0)).toBe(5); expect(wb.value(s, 3, 0)).toBe(null);
  });
  it('Ctrl = sao chép, không mất clipboard của người dùng', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '7'); wb.copy(s, { r1: 5, c1: 5, r2: 5, c2: 5 });
    wb.dragCells(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, 0, 2, 'copy');
    expect(wb.value(s, 0, 0)).toBe(7); expect(wb.value(s, 0, 2)).toBe(7);
    expect(wb.clip?.range.r1).toBe(5);
  });
  it('Shift = chèn: kéo A1 xuống giữa A3/A4 → các ô dồn lên/xuống như Excel', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    ['a', 'b', 'c', 'd', 'e'].forEach((v, i) => wb.setInput(s, i, 0, v));
    const g = wb.dragCells(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, 3, 0, 'insert');
    expect(col(wb, s, 0, 5)).toEqual(['b', 'c', 'a', 'd', 'e']);
    expect(g).toEqual({ r1: 2, c1: 0, r2: 2, c2: 0 });
    wb.undo();
    expect(col(wb, s, 0, 5)).toEqual(['a', 'b', 'c', 'd', 'e']);
  });
  it('Shift = chèn lên trên: kéo A5 lên A2', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    ['a', 'b', 'c', 'd', 'e'].forEach((v, i) => wb.setInput(s, i, 0, v));
    const g = wb.dragCells(s, { r1: 4, c1: 0, r2: 4, c2: 0 }, 1, 0, 'insert');
    expect(col(wb, s, 0, 5)).toEqual(['a', 'e', 'b', 'c', 'd']);
    expect(g).toEqual({ r1: 1, c1: 0, r2: 1, c2: 0 });
  });
});

describe('Kéo viền — ca chồng lấn', () => {
  it('Shift+kéo A2:A4 xuống 1 hàng (trong vùng nguồn) không mất dữ liệu', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    [1, 2, 3, 4, 5, 6].forEach((v, i) => wb.setInput(s, i, 0, String(v)));
    const g = wb.dragCells(s, { r1: 1, c1: 0, r2: 3, c2: 0 }, 2, 0, 'insert');
    expect(g).toBe(null);
    expect(col(wb, s, 0, 7)).toEqual([1, 2, 3, 4, 5, 6, null]);
  });
  it('Shift+kéo A2:A3 xuống dưới A5 → 1,4,5,2,3,6', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    [1, 2, 3, 4, 5, 6].forEach((v, i) => wb.setInput(s, i, 0, String(v)));
    wb.dragCells(s, { r1: 1, c1: 0, r2: 2, c2: 0 }, 5, 0, 'insert');
    expect(col(wb, s, 0, 7)).toEqual([1, 4, 5, 2, 3, 6, null]);
  });
  it('chồng một phần cột: báo lỗi, không đổi gì', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'a'); wb.setInput(s, 0, 1, 'b');
    expect(() => wb.dragCells(s, { r1: 0, c1: 0, r2: 0, c2: 1 }, 3, 1, 'insert')).toThrow();
    expect([wb.value(s, 0, 0), wb.value(s, 0, 1)]).toEqual(['a', 'b']);
  });
  it('bị chặn giữa chừng (trang bảo vệ) thì hoàn lại hết', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    ['a', 'b', 'c'].forEach((v, i) => wb.setInput(s, i, 0, v));
    wb.protectSheet(s, {});
    expect(() => wb.dragCells(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, 2, 0, 'insert')).toThrow();
    expect(col(wb, s, 0, 3)).toEqual(['a', 'b', 'c']);
  });
});
