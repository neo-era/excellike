import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';

const mk = () => {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  const rows = [['Vật tư', 'KL', 'Đơn giá'], ['Cát', 2, 100], ['Đá', 3, 200], ['Xi măng', 4, 50]];
  rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(s, r, c, String(v))));
  return { wb, s };
};
describe('Bảng & tham chiếu có cấu trúc', () => {
  it('tạo bảng, [@Cột] và Bảng1[Cột] tính đúng, hiện lại dạng có cấu trúc', () => {
    const { wb, s } = mk();
    const t = wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    expect(t.name).toBe('Bảng1');
    expect(t.cols.map((c) => c.name)).toEqual(['Vật tư', 'KL', 'Đơn giá']);
    wb.setInput(s, 1, 4, '=Bảng1[@KL]*Bảng1[@[Đơn giá]]');
    expect(wb.value(s, 1, 4)).toBe(200);
    expect(wb.rawText(s, 1, 4)).toBe('=Bảng1[@KL]*Bảng1[@[Đơn giá]]');
    wb.setInput(s, 5, 4, '=Bảng1[@KL]');
    expect((wb.value(s, 5, 4) as any).value).toBe('#VALUE!');
    wb.setInput(s, 6, 0, '=SUM(Bảng1[KL])');
    expect(wb.value(s, 6, 0)).toBe(9);
    expect(wb.rawText(s, 6, 0)).toBe('=SUM(Bảng1[KL])');
    wb.setInput(s, 7, 0, '=ROWS(Bảng1[#All])');
    expect(wb.value(s, 7, 0)).toBe(4);
    wb.setInput(s, 8, 0, '=Bảng1[[#Headers];[Đơn giá]]');
    expect(wb.value(s, 8, 0)).toBe('Đơn giá');
  });
  it('gõ ngay dưới bảng → bảng nở, SUM theo cột tự cập nhật; công thức cột tự điền', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 1, 3, '=[@KL]*[@[Đơn giá]]');
    expect(wb.afterEdit(s, 1, 3)).toBe(true); // vào cột trống → cột tính toán
    // cột D chưa thuộc bảng nên không phải cột tính toán; đưa D vào bảng trước
    wb.undo();
    wb.setInput(s, 0, 3, 'Thành tiền'); wb.afterEdit(s, 0, 3);
    expect(wb.tables(s)[0].c2).toBe(3);
    wb.setInput(s, 1, 3, '=[@KL]*[@[Đơn giá]]'); wb.afterEdit(s, 1, 3);
    expect([wb.value(s, 2, 3), wb.value(s, 3, 3)]).toEqual([600, 200]);
    wb.setInput(s, 6, 0, '=SUM(Bảng1[Thành tiền])');
    expect(wb.value(s, 6, 0)).toBe(1000);
    wb.group(() => { wb.setInput(s, 4, 0, 'Thép'); wb.afterEdit(s, 4, 0); });
    expect(wb.tables(s)[0].r2).toBe(4);
    expect(wb.rawText(s, 4, 3)).toBe('=[@KL]*[@[Đơn giá]]');
    wb.setInput(s, 4, 1, '5'); wb.setInput(s, 4, 2, '10');
    expect(wb.value(s, 6, 0)).toBe(1050);
    wb.undo(); wb.undo(); wb.undo();
    expect(wb.tables(s)[0].r2).toBe(3);
    expect(wb.value(s, 6, 0)).toBe(1000);
  });
  it('đổi tên bảng và tiêu đề cột → công thức hiển thị tên mới; chèn hàng giữa bảng → bảng nở', () => {
    const { wb, s } = mk();
    const t = wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 6, 0, '=SUM(Bảng1[KL])');
    wb.updateTable(s, t.id, { name: 'DuToan' });
    expect(wb.rawText(s, 6, 0)).toBe('=SUM(DuToan[KL])');
    wb.setInput(s, 0, 1, 'Khối lượng'); wb.afterEdit(s, 0, 1);
    expect(wb.rawText(s, 6, 0)).toBe('=SUM(DuToan[Khối lượng])');
    wb.insertRows(s, 2, 1);
    wb.setInput(s, 2, 1, '10');
    expect(wb.tables(s)[0].r2).toBe(4);
    expect(wb.value(s, 7, 0)).toBe(19);
    expect(() => wb.updateTable(s, t.id, { name: 'A1' })).toThrow();
  });
  it('hàng tổng dùng SUBTOTAL(109;[Cột]); chuyển thành vùng → công thức địa chỉ ô', () => {
    const { wb, s } = mk();
    const t = wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setTableTotals(s, t.id, true);
    expect(wb.value(s, 4, 0)).toBe('Tổng');
    expect(wb.rawText(s, 4, 2)).toBe('=SUBTOTAL(109;[Đơn giá])');
    expect(wb.value(s, 4, 2)).toBe(350);
    wb.setTotalFn(s, t.id, 1, 'average');
    expect(wb.value(s, 4, 1)).toBe(3);
    wb.setInput(s, 6, 0, '=Bảng1[[#Totals];[KL]]');
    expect(wb.value(s, 6, 0)).toBe(3);
    wb.convertTableToRange(s, t.id);
    expect(wb.tables(s).length).toBe(0);
    expect(wb.rawText(s, 4, 2)).toBe("=SUBTOTAL(109;Sheet1!$C$2:$C$4)");
    expect(wb.value(s, 4, 2)).toBe(350);
    expect(wb.style(s, 0, 0).b).toBe(true); // giữ định dạng bảng
    wb.undo();
    expect(wb.tables(s).length).toBe(1);
    expect(wb.rawText(s, 4, 2)).toBe('=SUBTOTAL(109;[Đơn giá])');
  });
  it('lưu/mở JSON giữ bảng và công thức; sao chép trang có bảng mới', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 1, 3, '=[@KL]*2');
    wb.setInput(s, 6, 0, '=SUM(Bảng1[KL])');
    const wb2 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    const s2 = wb2.sheetIds()[0];
    expect(wb2.value(s2, 6, 0)).toBe(9);
    expect(wb2.rawText(s2, 6, 0)).toBe('=SUM(Bảng1[KL])');
    // ô D2 ở ngoài bảng (cột D) nên [@KL] dùng được khi D thuộc hàng dữ liệu? Excel: #VALUE! khi ngoài bảng
    const n = wb2.copySheet(s2);
    expect(wb2.tables(n)[0].name).toBe('Bảng2');
    expect(wb2.value(n, 6, 0)).toBe(9);
    expect(wb2.fileFormula(s2, 6, 0, '=SUM(Bảng1[KL])')).toBe('=SUM(Bảng1[KL])');
  });
  it('dạng tệp: [@Cột] → Bảng1[[#This Row],[Cột]]', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    expect(wb.fileFormula(s, 1, 1, '=[@KL]*[@[Đơn giá]]')).toBe('=Bảng1[[#This Row],[KL]]*Bảng1[[#This Row],[Đơn giá]]');
  });
});

import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { writeFileSync } from 'node:fs';
describe('Bảng trong .xlsx', () => {
  it('ghi rồi đọc lại: bảng, hàng tổng, công thức có cấu trúc', async () => {
    const { wb, s } = mk();
    const t = wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 0, 3, 'Thành tiền'); wb.afterEdit(s, 0, 3);
    wb.setInput(s, 1, 3, '=[@KL]*[@[Đơn giá]]'); wb.afterEdit(s, 1, 3);
    wb.setTableTotals(s, t.id, true);
    wb.setInput(s, 7, 0, '=SUM(Bảng1[Thành tiền])');
    const buf = await writeXlsx(wb);
    if (process.env.KEEP) writeFileSync(process.env.KEEP, buf);
    const { wb: wb2 } = await readXlsx(buf);
    const s2 = wb2.sheetIds()[0];
    const t2 = wb2.tables(s2)[0];
    expect(t2.name).toBe('Bảng1');
    expect(t2.totals).toBe(true);
    expect(t2.cols.map((c) => c.name)).toEqual(['Vật tư', 'KL', 'Đơn giá', 'Thành tiền']);
    expect(wb2.rawText(s2, 2, 3)).toBe('=[@KL]*[@[Đơn giá]]');
    expect(wb2.value(s2, 3, 3)).toBe(200);
    expect(wb2.rawText(s2, 4, 3)).toBe('=SUBTOTAL(109;[Thành tiền])');
    expect(wb2.value(s2, 4, 3)).toBe(1000);
    expect(wb2.value(s2, 7, 0)).toBe(1000);
    expect(wb2.meta(s2).filter?.range.r1).toBe(0);
  });
});

describe('Bảng — sửa sau rà soát', () => {
  const two = () => {
    const { wb, s } = mk();
    const t1 = wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 0, 5, 'A'); wb.setInput(s, 1, 5, '4'); wb.setInput(s, 2, 5, '6');
    const t2 = wb.createTable(s, { r1: 0, c1: 5, r2: 2, c2: 5 });
    return { wb, s, t1, t2 };
  };
  it('#1 chuyển một bảng thành vùng không làm hỏng tham chiếu bảng khác', () => {
    const { wb, s, t1 } = two();
    wb.setInput(s, 1, 6, '=Bảng2[@A]/SUM(Bảng2[A])');
    wb.setInput(s, 8, 0, '=SUM(Bảng1[KL])+SUM(Bảng2[A])');
    wb.convertTableToRange(s, t1.id);
    expect(wb.rawText(s, 1, 6)).toBe('=Bảng2[@A]/SUM(Bảng2[A])');
    expect(wb.rawText(s, 8, 0)).toBe('=SUM(Sheet1!$B$2:$B$4)+SUM(Bảng2[A])');
    expect(wb.value(s, 8, 0)).toBe(19);
  });
  it('#2 tính tạm công thức có bảng không làm lệch hoàn tác', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 9, 0, '5');
    expect(wb.hf.calculateFormula('=SUM(Bảng1[Đơn giá])', s)).toBe(350);
    wb.undo();
    expect(wb.value(s, 9, 0)).toBe(null);
    expect(wb.tables(s).length).toBe(1);
  });
  it('#3 chèn hàng ở hàng dữ liệu đầu: tên vùng bảng nở theo', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 8, 0, '=SUM(Bảng1[KL])');
    wb.insertRows(s, 1, 1);
    wb.setInput(s, 1, 1, '10');
    expect(wb.value(s, 9, 0)).toBe(19);
    wb.undo(); wb.undo();
    expect(wb.value(s, 8, 0)).toBe(9);
    expect(wb.tables(s)[0].r2).toBe(3);
  });
  it('#4 tên cột có dấu phẩy, chấm, # qua .xlsx', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    [['Giá, VNĐ', 'Đơn giá (1.000đ)', 'Số #'], ['1', '2', '3']].forEach((row, r) => row.forEach((v, c) => wb.setInput(s, r, c, v)));
    wb.createTable(s, { r1: 0, c1: 0, r2: 1, c2: 2 });
    wb.setInput(s, 3, 0, "=SUM(Bảng1[Giá, VNĐ])+SUM(Bảng1[Đơn giá (1.000đ)])+SUM(Bảng1[Số '#])");
    expect(wb.value(s, 3, 0)).toBe(6);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0];
    expect(w2.value(s2, 3, 0)).toBe(6);
  });
  it('#5 xóa ô tiêu đề → Cột<n>; #14 chuỗi trong công thức không bị dịch', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.withHeaders(s, () => wb.clear(s, { r1: 0, c1: 1, r2: 0, c2: 1 }));
    expect(wb.tables(s)[0].cols[1].name).toBe('Cột2');
    expect(wb.value(s, 0, 1)).toBe('Cột2');
    wb.setInput(s, 8, 0, '="XT_1_D và XTHISROW($B2)"');
    expect(wb.rawText(s, 8, 0)).toBe('="XT_1_D và XTHISROW($B2)"');
  });
  it('#7 không nở đè lên bảng khác; #10 xóa trang có bảng', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 1 });
    wb.setInput(s, 4, 1, 'X'); wb.setInput(s, 5, 1, '1');
    wb.createTable(s, { r1: 4, c1: 1, r2: 5, c2: 1 });
    wb.setInput(s, 4, 0, 'gõ'); wb.afterEdit(s, 4, 0);
    expect(wb.tables(s)[0].r2).toBe(3);
    const n = wb.addSheet();
    wb.removeSheet(s);
    expect(wb.tableReg().length).toBe(0);
    void n;
  });
  it('#8 cột bị xóa → #REF!; #13 truy tìm qua vùng bảng', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 8, 0, '=SUM(Bảng1[Đơn giá])');
    expect(wb.precedentsOf(s, 8, 0)).toEqual([{ s, range: { r1: 1, c1: 2, r2: 3, c2: 2 } }]);
    expect(wb.directDependents(s, 2, 2)).toEqual([{ s, r: 8, c: 0 }]);
    wb.deleteCols(s, 2, 1);
    expect(wb.rawText(s, 8, 0)).toMatch(/#REF!/);
  });
  it('#11 đổi kích thước có hàng tổng; #15 loại trùng trong bảng co bảng', () => {
    const { wb, s } = mk();
    const t = wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setTableTotals(s, t.id, true);
    wb.resizeTable(s, t.id, { r1: 0, c1: 0, r2: 6, c2: 2 });
    expect(wb.rawText(s, 6, 2)).toBe('=SUBTOTAL(109;[Đơn giá])');
    expect(wb.rawText(s, 4, 2)).toBe('');
    const { wb: w, s: s2 } = mk();
    w.setInput(s2, 3, 0, 'Cát'); w.setInput(s2, 3, 1, '2'); w.setInput(s2, 3, 2, '100');
    const t2 = w.createTable(s2, { r1: 0, c1: 0, r2: 3, c2: 2 });
    w.setTableTotals(s2, t2.id, true);
    const res = w.removeDuplicates(s2, { r1: 0, c1: 0, r2: 3, c2: 2 }, [0, 1, 2], true);
    expect(res.removed).toBe(1);
    expect(w.tables(s2)[0].r2).toBe(3);
    expect(w.value(s2, 3, 2)).toBe(300);
  });
  it('#6 cắt-dán cả bảng → bảng đi theo', () => {
    const { wb, s } = mk();
    wb.createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 });
    wb.setInput(s, 8, 6, '=SUM(Bảng1[KL])');
    wb.copy(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, true);
    wb.paste(s, 10, 0);
    expect(wb.tables(s)[0].r1).toBe(10);
    wb.setInput(s, 12, 1, '10');
    expect(wb.value(s, 8, 6)).toBe(16);
  });
});
