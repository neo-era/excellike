import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { addPrintArea, normPrint, paginate, storePrint } from '../src/core/paginate';

const R = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
function book() {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  wb.renameSheet(s, 'Trang 1');
  for (let r = 0; r < 200; r++) for (let c = 0; c < 6; c++) wb.setInput(s, r, c, String(r * 10 + c));
  return { wb, s };
}

describe('Vùng in nhiều khối (C3)', () => {
  it('Thêm vào Vùng In: liền kề thì nới khối cũ, nằm trong thì giữ, còn lại thành khối mới', () => {
    expect(addPrintArea([R(0, 0, 9, 2)], R(0, 3, 9, 4))).toEqual([R(0, 0, 9, 4)]);
    expect(addPrintArea([R(0, 0, 9, 2)], R(10, 0, 19, 2))).toEqual([R(0, 0, 19, 2)]);
    expect(addPrintArea([R(0, 0, 9, 2)], R(2, 1, 3, 1))).toEqual([R(0, 0, 9, 2)]);
    expect(addPrintArea([R(0, 0, 9, 2)], R(20, 4, 25, 5))).toEqual([R(0, 0, 9, 2), R(20, 4, 25, 5)]);
  });
  it('mỗi khối bắt đầu trang mới, theo thứ tự trong Print_Area; khối lớn tự chia trang', () => {
    const { wb, s } = book();
    wb.setPrint(s, storePrint({ ...normPrint(null), area: R(0, 3, 4, 4), more: [R(0, 0, 4, 1), R(10, 0, 199, 2)] }));
    const L = paginate(wb, s, normPrint(wb.meta(s).print));
    expect(L.areas).toEqual([R(0, 3, 4, 4), R(0, 0, 4, 1), R(10, 0, 199, 2)]);
    expect(L.pages[0].cols).toEqual([3, 4]);
    expect(L.pages[0].rows).toEqual([0, 1, 2, 3, 4]);
    expect(L.pages[1].cols).toEqual([0, 1]);
    expect(L.pages.length).toBeGreaterThan(3); // khối 190 hàng > 1 trang A4
    expect(L.pages[2].rows[0]).toBe(10);
    expect(L.pages.map((p) => p.index)).toEqual(L.pages.map((_, i) => i));
  });
  it('ghi một tên _xlnm.Print_Area duy nhất, khối cách dấu phẩy, đủ $; mở lại giữ đủ khối', async () => {
    const { wb, s } = book();
    wb.setPrint(s, storePrint({ ...normPrint(null), area: R(0, 0, 19, 5), more: [R(30, 0, 39, 2)] }));
    const bytes = await writeXlsx(wb);
    const wx = await (await JSZip.loadAsync(bytes)).file('xl/workbook.xml')!.async('string');
    const all = [...wx.matchAll(/<definedName[^>]*name="_xlnm\.Print_Area"[^>]*>([^<]*)<\/definedName>/g)];
    expect(all.length).toBe(1);
    expect(all[0][1]).toBe("'Trang 1'!$A$1:$F$20,'Trang 1'!$A$31:$C$40");
    const { wb: w2 } = await readXlsx(bytes);
    const p = w2.meta(w2.sheetIds()[0]).print!;
    expect(p.area).toEqual(R(0, 0, 19, 5));
    expect(p.more).toEqual([R(30, 0, 39, 2)]);
  });
  it('chèn/xóa hàng dời mọi khối', () => {
    const { wb, s } = book();
    wb.setPrint(s, storePrint({ ...normPrint(null), area: R(0, 0, 9, 2), more: [R(20, 0, 29, 2)] }));
    wb.insertRows(s, 5, 2);
    expect(wb.meta(s).print!.area).toEqual(R(0, 0, 11, 2));
    expect(wb.meta(s).print!.more).toEqual([R(22, 0, 31, 2)]);
    wb.deleteRows(s, 22, 10);
    expect(wb.meta(s).print!.more ?? []).toEqual([]);
  });
});
