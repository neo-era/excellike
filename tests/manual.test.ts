import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import JSZip from 'jszip';

/** sổ nhỏ: A1..A3 số, B1 = A1*2, C1 = B1+1, D1 = SUM(A1:A3), E1 = VLOOKUP */
function book() {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  wb.setInput(s, 0, 0, '1'); wb.setInput(s, 1, 0, '2'); wb.setInput(s, 2, 0, '3');
  wb.setInput(s, 0, 1, '=A1*2'); wb.setInput(s, 0, 2, '=B1+1'); wb.setInput(s, 0, 3, '=SUM(A1:A3)');
  wb.setInput(s, 0, 5, 'k'); wb.setInput(s, 0, 6, '10');
  wb.setInput(s, 0, 4, '=VLOOKUP("k";F1:G1;2;FALSE)');
  return { wb, s };
}
const edit = (wb: Workbook, s: number, r: number, c: number, t: string) => wb.lazy(() => wb.setInput(s, r, c, t));
const snapshot = (wb: Workbook) => wb.sheetIds().map((id) => { const d = wb.used(id); const out: unknown[] = []; for (let r = 0; r < d.rows; r++) for (let c = 0; c < d.cols; c++) out.push(wb.value(id, r, c)); return out; });

describe('Tính toán Thủ công (B)', () => {
  it('B1: sửa giá trị không tính lại ô phụ thuộc; ô vừa sửa hiện ngay; F9 thì tính', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 0, 0, '5');
    expect(wb.value(s, 0, 0)).toBe(5);
    expect(wb.value(s, 0, 1)).toBe(2);
    expect(wb.value(s, 0, 2)).toBe(3);
    expect(wb.value(s, 0, 3)).toBe(6);
    expect(wb.calcDirty).toBe(true);
    wb.calcNow();
    expect(wb.value(s, 0, 1)).toBe(10);
    expect(wb.value(s, 0, 2)).toBe(11);
    expect(wb.value(s, 0, 3)).toBe(10);
    expect(wb.calcDirty).toBe(false);
  });
  it('B1b: đọc mọi kiểu giá trị khi đang Thủ công không báo lỗi tạm dừng', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 1, 0, '7');
    expect(() => { wb.display(s, 0, 1); wb.rawText(s, 0, 1); wb.effFmt(s, 0, 1); wb.hf.getSheetValues(s); wb.hf.getSheetSerialized(s); wb.toJSON(); }).not.toThrow();
  });
  it('B4: công thức vừa gõ tự tính ngay, ô phụ thuộc chờ F9', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 0, 1, '=A1*100');
    expect(wb.value(s, 0, 1)).toBe(100);
    expect(wb.value(s, 0, 2)).toBe(3);
    wb.calcNow();
    expect(wb.value(s, 0, 2)).toBe(101);
  });
  it('F9 sau công thức vừa gõ: ô phụ thuộc vẫn được tính (dù công thức đã có giá trị tạm)', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 0, 1, '=A1+41');
    expect(wb.value(s, 0, 1)).toBe(42);
    wb.calcNow();
    expect(wb.value(s, 0, 2)).toBe(43);
  });
  it('VLOOKUP qua chỉ mục cột vẫn đúng sau khi sửa ở Thủ công rồi F9', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 0, 6, '99');
    expect(wb.value(s, 0, 4)).toBe(10);
    wb.calcNow();
    expect(wb.value(s, 0, 4)).toBe(99);
    edit(wb, s, 0, 6, '=50+1');
    wb.calcNow();
    expect(wb.value(s, 0, 4)).toBe(51);
  });
  it('B3: kết quả sau F9 trùng chế độ Tự động; F9 giữ lịch sử Hoàn tác', () => {
    const a = book(), b = book();
    a.wb.setCalcMode('manual');
    for (let i = 0; i < 50; i++) { edit(a.wb, a.s, i % 3, 0, String(i)); edit(b.wb, b.s, i % 3, 0, String(i)); }
    edit(a.wb, a.s, 3, 3, '=D1*A2'); edit(b.wb, b.s, 3, 3, '=D1*A2');
    a.wb.calcNow();
    expect(snapshot(a.wb)).toEqual(snapshot(b.wb));
    expect(a.wb.canUndo()).toBe(true);
    a.wb.undo(); b.wb.undo();
    a.wb.calcNow();
    expect(snapshot(a.wb)).toEqual(snapshot(b.wb));
  });
  it('Ctrl+Alt+F9 tính lại tất cả, giữ Hoàn tác; Ctrl+Shift+Alt+F9 dựng lại (mất Hoàn tác)', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 0, 0, '4');
    wb.calcAll();
    expect(wb.value(s, 0, 2)).toBe(9);
    expect(wb.canUndo()).toBe(true);
    edit(wb, s, 0, 0, '6');
    wb.recalc();
    expect(wb.value(s, 0, 2)).toBe(13);
    expect(wb.canUndo()).toBe(false);
    expect(wb.calcMode).toBe('manual');
    edit(wb, s, 0, 0, '1');
    expect(wb.value(s, 0, 2)).toBe(13);
  });
  it('B5: chèn hàng, sắp xếp, thêm trang… tính lại ngay dù đang Thủ công', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 0, 0, '8');
    expect(wb.value(s, 0, 1)).toBe(2);
    wb.insertRows(s, 5, 1);
    expect(wb.value(s, 0, 1)).toBe(16);
    expect(wb.calcDirty).toBe(false);
    edit(wb, s, 1, 0, '20');
    wb.addSheet('Mới');
    expect(wb.value(s, 0, 3)).toBe(31);
  });
  it('B7: Hoàn tác/Làm lại sửa ô ở Thủ công; về Tự động thì tính ngay', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 0, 0, '9');
    edit(wb, s, 0, 1, '=A1*3');
    expect(wb.value(s, 0, 1)).toBe(27);
    wb.undo();
    expect(wb.rawText(s, 0, 1)).toBe('=A1*2');
    expect(wb.value(s, 0, 1)).toBe(18); // công thức cũ khôi phục: tự tính từ giá trị đang có
    wb.redo();
    expect(wb.value(s, 0, 1)).toBe(27);
    expect(wb.value(s, 0, 2)).toBe(3);
    wb.setCalcMode('auto');
    expect(wb.value(s, 0, 2)).toBe(28);
    edit(wb, s, 0, 0, '1');
    expect(wb.value(s, 0, 2)).toBe(4);
  });
  it('B1c: hàm biến động không đổi khi sửa ô, chỉ đổi khi F9', () => {
    const { wb, s } = book();
    wb.setInput(s, 5, 0, '=RAND()');
    const r0 = wb.value(s, 5, 0);
    wb.setCalcMode('manual');
    edit(wb, s, 0, 0, '2'); edit(wb, s, 0, 0, '3');
    expect(wb.value(s, 5, 0)).toBe(r0);
    wb.calcNow();
    expect(wb.value(s, 5, 0)).not.toBe(r0);
  });
  it('B1c: tính lặp — F9 ở Thủ công cho kết quả như Tự động', () => {
    const mk = () => { const { wb, s } = book(); wb.setIterate({ on: true, count: 100, delta: 0.0001 }); wb.setInput(s, 6, 0, '=B7/2+1'); wb.setInput(s, 6, 1, '=A7'); return { wb, s }; };
    const a = mk(), b = mk();
    a.wb.setCalcMode('manual');
    edit(a.wb, a.s, 6, 1, '=A7+1'); edit(b.wb, b.s, 6, 1, '=A7+1');
    a.wb.calcNow();
    expect(Math.abs((a.wb.value(a.s, 6, 0) as number) - (b.wb.value(b.s, 6, 0) as number))).toBeLessThan(1e-9);
  });
  it('mảng động vừa gõ ở Thủ công: tràn đúng vùng ngay', () => {
    const { wb, s } = book();
    wb.setCalcMode('manual');
    edit(wb, s, 10, 0, '=SEQUENCE(3)');
    expect([wb.value(s, 10, 0), wb.value(s, 11, 0), wb.value(s, 12, 0)]).toEqual([1, 2, 3]);
  });
  it('ca xấu: sửa 1.000 ô rồi F9 một lần', () => {
    const { wb, s } = book();
    wb.setInput(s, 0, 7, '=SUM(H2:H1001)');
    wb.setCalcMode('manual');
    for (let i = 1; i <= 1000; i++) edit(wb, s, i, 7, String(i));
    expect(wb.value(s, 0, 7)).toBe(0);
    wb.calcNow();
    expect(wb.value(s, 0, 7)).toBe(500500);
  });
  it('B6: lưu/mở giữ calcMode="manual" và calcOnSave; bản lưu tự động (JSON) cũng vậy', async () => {
    const { wb, s } = book();
    wb.setCalcMode('manual'); wb.calcOnSave = false;
    edit(wb, s, 0, 0, '5');
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    const wx = await z.file('xl/workbook.xml')!.async('string');
    expect(wx).toMatch(/<calcPr[^>]*calcMode="manual"/);
    expect(wx).toMatch(/calcOnSave="0"/);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    expect(w2.calcMode).toBe('manual');
    expect(w2.calcOnSave).toBe(false);
    // mở tệp: tính toàn bộ khi dựng sổ rồi mới Thủ công
    expect(w2.value(w2.sheetIds()[0], 0, 1)).toBe(10);
    const w3 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(w3.calcMode).toBe('manual');
    expect(w3.value(w3.sheetIds()[0], 0, 2)).toBe(11); // tắt app khi còn chờ tính → mở lại tính đúng
    const plain = Workbook.create();
    const z2 = await JSZip.loadAsync(await writeXlsx(plain));
    expect(await z2.file('xl/workbook.xml')!.async('string')).not.toMatch(/calcMode/);
  });
  it('rà mã: công thức vừa gõ tính đúng tại ô của nó, đọc được công thức vừa gõ khác; F9 vẫn tính ô phụ thuộc', () => {
    const { wb, s } = book();
    wb.setInput(s, 4, 3, '=E5+1'); // D5 phụ thuộc E5
    wb.setCalcMode('manual');
    edit(wb, s, 4, 0, '=ROW()');
    edit(wb, s, 4, 1, '=COLUMN()*10');
    edit(wb, s, 4, 2, '=A1+1');
    edit(wb, s, 4, 4, '=C5*2');
    expect([wb.value(s, 4, 0), wb.value(s, 4, 1), wb.value(s, 4, 2), wb.value(s, 4, 4)]).toEqual([5, 20, 2, 4]);
    expect(wb.value(s, 4, 3)).toBe(1); // ô phụ thuộc chưa tính
    wb.calcNow();
    expect(wb.value(s, 4, 3)).toBe(5);
    edit(wb, s, 0, 0, '3');
    wb.calcNow();
    expect([wb.value(s, 4, 2), wb.value(s, 4, 4), wb.value(s, 4, 3)]).toEqual([4, 8, 9]);
  });
});
