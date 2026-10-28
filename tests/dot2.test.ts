// Đợt 2: chuyển đổi công thức, CSV, gộp ô, viền, sắp xếp, lọc, định dạng có điều kiện, kiểm tra dữ liệu, .xlsx
import { readFileSync } from 'node:fs';
import { Workbook } from '../src/core/workbook';
import { toEn, toVi } from '../src/core/locale';
import { parseCSV, toCSV } from '../src/core/csv';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

const one = (r: number, c: number) => ({ r1: r, c1: c, r2: r, c2: c });
const rg = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
function book(rows: string[][]) {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setInput(s, r, c, v); }));
  return { wb, s, t: (r: number, c: number) => wb.display(s, r, c).text, v: (r: number, c: number) => wb.value(s, r, c) };
}

describe('Chuyển công thức Việt ↔ Excel gốc (en)', () => {
  it('dấu ngăn tham số và dấu thập phân', () => {
    expect(toEn('=ROUND(1,2345;2)+SUM(A1:B2;3)')).toBe('=ROUND(1.2345,2)+SUM(A1:B2,3)');
    expect(toVi('=ROUND(1.2345,2)+SUM(A1:B2,3)')).toBe('=ROUND(1,2345;2)+SUM(A1:B2;3)');
  });
  it('không đụng chuỗi, tên hàm có dấu chấm, tên sheet', () => {
    expect(toVi('=IF(A1>1,"1.5, ok",STDEV.S(B1:B3))')).toBe('=IF(A1>1;"1.5, ok";STDEV.S(B1:B3))');
    expect(toEn("='Q1, 2026'!A1+1,5")).toBe("='Q1, 2026'!A1+1.5");
    expect(toVi('=1E-3*2.5E+2')).toBe('=1E-3*2,5E+2');
  });
});

describe('CSV', () => {
  it('nhận đúng dấu phân cách, ngoặc kép, BOM, CRLF', () => {
    expect(parseCSV('﻿a,"b,c",1.5\r\n"x""y",,2\r\n')).toEqual({ rows: [['a', 'b,c', '1.5'], ['x"y', '', '2']], delim: ',' });
    expect(parseCSV('Tên;Giá\nĐèn;1.250,5\n').delim).toBe(';');
  });
  it('xuất CSV có BOM và ngoặc khi cần', () => {
    expect(toCSV([['a', 'b;c'], ['1,5', 'x"y']], ';')).toBe('﻿a;"b;c"\r\n1,5;"x""y"');
  });
  it('nhập CSV vào sheet: số dấu chấm (CSV quốc tế) và số kiểu Việt', () => {
    const wb = Workbook.create();
    const id = wb.importCSV('Đèn,Giá\nLED,1250.5\n', 'Nhập');
    expect(wb.value(id, 1, 1)).toBe(1250.5);
    const id2 = wb.importCSV('Đèn;Giá\nLED;1.250,5\n', 'Nhập 2');
    expect(wb.value(id2, 1, 1)).toBe(1250.5);
  });
});

describe('Gộp ô', () => {
  it('giữ giá trị ô trên trái, xóa ô khác, căn giữa; hoàn tác đầy đủ', () => {
    const { wb, s, v } = book([['a', 'b'], ['c', 'd']]);
    wb.merge(s, rg(0, 0, 1, 1));
    expect(wb.mergeAt(s, 1, 1)).toEqual(rg(0, 0, 1, 1));
    expect(v(0, 0)).toBe('a');
    expect(v(1, 1)).toBeNull();
    expect(wb.style(s, 0, 0).al).toBe('c');
    wb.undo();
    expect(wb.mergeAt(s, 1, 1)).toBeNull();
    expect(v(1, 1)).toBe('d');
  });
  it('chèn hàng phía trên làm vùng gộp dời theo, bỏ gộp được', () => {
    const { wb, s } = book([['x']]);
    wb.merge(s, rg(2, 1, 3, 2));
    wb.insertRows(s, 0, 2);
    expect(wb.mergeAt(s, 4, 1)).toEqual(rg(4, 1, 5, 2));
    wb.unmerge(s, one(5, 2));
    expect(wb.mergeAt(s, 4, 1)).toBeNull();
  });
});

describe('Viền', () => {
  it('viền ngoài chỉ đặt ở cạnh ngoài vùng; tất cả viền đặt mọi cạnh; không viền xóa', () => {
    const { wb, s } = book([]);
    wb.setBorder(s, rg(0, 0, 1, 1), 'outer', 'thin');
    expect(wb.style(s, 0, 0).bd).toEqual({ t: { s: 'thin' }, l: { s: 'thin' } });
    expect(wb.style(s, 1, 1).bd).toEqual({ b: { s: 'thin' }, r: { s: 'thin' } });
    wb.setBorder(s, rg(0, 0, 1, 1), 'all', 'medium');
    expect(Object.keys(wb.style(s, 0, 1).bd!).sort()).toEqual(['b', 'l', 'r', 't']);
    wb.setBorder(s, rg(0, 0, 1, 1), 'none');
    expect(wb.style(s, 0, 0).bd).toBeUndefined();
  });
});

describe('Sắp xếp', () => {
  const data = [['Tên', 'SL', 'Tiền'], ['Đông', '3', '=B2*10'], ['Ánh', '1', '=B3*10'], ['Bình', '', '=B4*10'], ['Dũng', '2', '=B5*10']];
  it('tăng dần theo cột số, ô trống xuống cuối, giữ hàng tiêu đề, công thức theo hàng', () => {
    const { wb, s, t } = book(data);
    wb.setStyle(s, one(1, 0), { b: true });
    wb.sort(s, rg(0, 0, 4, 2), [{ col: 1 }], true);
    expect([1, 2, 3, 4].map((r) => t(r, 0))).toEqual(['Ánh', 'Dũng', 'Đông', 'Bình']);
    expect([1, 2, 3, 4].map((r) => t(r, 2))).toEqual(['10', '20', '30', '0']);
    expect(wb.rawText(s, 1, 2)).toBe('=B2*10');
    expect(wb.style(s, 3, 0).b).toBe(true);
    wb.undo();
    expect(t(1, 0)).toBe('Đông');
    expect(wb.style(s, 1, 0).b).toBe(true);
  });
  it('giảm dần theo chữ tiếng Việt (Đ sau D)', () => {
    const { wb, s, t } = book(data);
    wb.sort(s, rg(0, 0, 4, 2), [{ col: 0, desc: true }], true);
    expect([1, 2, 3, 4].map((r) => t(r, 0))).toEqual(['Đông', 'Dũng', 'Bình', 'Ánh']);
  });
  it('tự nhận vùng liền kề và tiêu đề', () => {
    const { wb, s } = book(data);
    expect(wb.currentRegion(s, 2, 1)).toEqual(rg(0, 0, 4, 2));
    expect(wb.guessHeader(s, rg(0, 0, 4, 2))).toBe(true);
  });
  it('không sắp xếp vùng có ô gộp', () => {
    const { wb, s } = book(data);
    wb.merge(s, rg(1, 0, 2, 0));
    expect(() => wb.sort(s, rg(0, 0, 4, 2), [{ col: 1 }], true)).toThrow();
  });
});

describe('Lọc (AutoFilter)', () => {
  it('ẩn hàng không thỏa, bỏ lọc hiện lại, hoàn tác được', () => {
    const { wb, s } = book([['Loại', 'SL'], ['Vật tư', '1'], ['Thiết bị', '2'], ['Vật tư', '3'], ['Nhân công', '4']]);
    wb.setFilter(s, rg(0, 0, 4, 1));
    expect(wb.filterValues(s, 0)).toEqual(['Nhân công', 'Thiết bị', 'Vật tư']);
    wb.filterColumn(s, 0, { values: ['Vật tư'] });
    expect([1, 2, 3, 4].map((r) => wb.isRowHidden(s, r))).toEqual([false, true, false, true]);
    wb.filterColumn(s, 1, { text: '3' });
    expect([1, 2, 3, 4].map((r) => wb.isRowHidden(s, r))).toEqual([true, true, false, true]);
    wb.undo();
    expect(wb.isRowHidden(s, 1)).toBe(false);
    wb.clearFilter(s);
    expect([1, 2, 3, 4].some((r) => wb.isRowHidden(s, r))).toBe(false);
  });
});

describe('Định dạng có điều kiện', () => {
  it('lớn hơn, trùng lặp, top N, thang màu; tự cập nhật khi đổi dữ liệu', () => {
    const { wb, s } = book([['50'], ['150'], ['150'], ['300']]);
    wb.addCF(s, { range: rg(0, 0, 3, 0), type: 'cellIs', op: 'gt', v1: '100', fmt: { bg: '#ffc7ce', fg: '#9c0006' } });
    expect(wb.cfStyle(s, 1, 0)).toMatchObject({ bg: '#ffc7ce' });
    expect(wb.cfStyle(s, 0, 0)).toEqual({});
    wb.addCF(s, { range: rg(0, 0, 3, 0), type: 'dup', fmt: { b: true } });
    expect(wb.cfStyle(s, 2, 0).b).toBe(true);
    expect(wb.cfStyle(s, 3, 0).b).toBeFalsy();
    wb.addCF(s, { range: rg(0, 0, 3, 0), type: 'top', rank: 1, fmt: { i: true, fg: '#0000ff' } });
    expect(wb.cfStyle(s, 3, 0)).toMatchObject({ i: true, fg: '#9c0006' }); // quy tắc đứng trước được ưu tiên màu chữ
    wb.setInput(s, 0, 0, '999');
    expect(wb.cfStyle(s, 0, 0).i).toBe(true);
    expect(wb.cfStyle(s, 3, 0).i).toBeFalsy();
    const { wb: w2, s: s2 } = book([['0'], ['50'], ['100']]);
    w2.addCF(s2, { range: rg(0, 0, 2, 0), type: 'scale', colors: ['#ff0000', '#00ff00'] });
    expect(w2.cfStyle(s2, 0, 0).bg).toBe('#ff0000');
    expect(w2.cfStyle(s2, 2, 0).bg).toBe('#00ff00');
    expect(w2.cfStyle(s2, 1, 0).bg).toBe('#808000');
  });
});

describe('Kiểm tra dữ liệu', () => {
  it('danh sách, số nguyên trong khoảng, cho phép trống', () => {
    const { wb, s } = book([]);
    wb.addDV(s, { range: rg(1, 2, 9, 2), type: 'list', list: ['bộ', 'cột', 'm'], allowBlank: true });
    wb.addDV(s, { range: rg(1, 3, 9, 3), type: 'whole', op: 'between', v1: '0', v2: '1000', error: 'Khối lượng 0–1000' });
    expect(wb.validate(s, 1, 2, 'cột')).toBeNull();
    expect(wb.validate(s, 1, 2, 'xyz')).toMatch(/danh sách/);
    expect(wb.validate(s, 1, 2, '')).toBeNull();
    expect(wb.validate(s, 2, 3, '1500')).toBe('Khối lượng 0–1000');
    expect(wb.validate(s, 2, 3, '1,5')).toBe('Khối lượng 0–1000');
    expect(wb.validate(s, 2, 3, '500')).toBeNull();
    expect(wb.validate(s, 0, 0, 'gì cũng được')).toBeNull();
    expect(wb.dvList(s, 5, 2)).toEqual(['bộ', 'cột', 'm']);
  });
});

describe('Mở tệp .xlsx (tạo kiểu Excel)', () => {
  it('đọc giá trị, công thức, định dạng, gộp ô, cỡ cột/hàng, cố định ngăn, CF, DV, lọc, nhiều sheet', async () => {
    const { wb, warnings } = await readXlsx(readFileSync('tests/fixtures/mau.xlsx'));
    const [a, b, c] = wb.sheetIds();
    expect(wb.sheetIds().map((i) => wb.sheetName(i))).toEqual(['Dự toán', '2024', 'Q1']);
    const t = (r: number, cc: number) => wb.display(a, r, cc).text;
    expect(wb.value(a, 1, 5)).toBe(204000000);
    expect(t(1, 5)).toBe('204.000.000');
    expect(wb.rawText(a, 1, 5)).toBe('=D2*E2');
    expect(wb.value(a, 4, 5)).toBe(204000000 + 326400000 + 148680000);
    expect(t(1, 6)).toBe('02/02/2026');
    expect(wb.value(a, 1, 7)).toBe('lớn');
    expect(t(2, 7)).toBe('1,23');
    expect(wb.rawText(a, 2, 7)).toBe('=ROUND(1,2345;2)');
    expect(wb.value(a, 3, 7)).toBe(6800000);
    expect(wb.style(a, 0, 0)).toMatchObject({ b: true, fg: '#c00000', bg: '#dcefe5' });
    expect(wb.style(a, 0, 1)).toMatchObject({ i: true, u: true });
    expect(wb.style(a, 1, 2).al).toBe('c');
    expect(wb.style(a, 1, 1).bd).toMatchObject({ l: { s: 'thin' }, b: { s: 'medium' } });
    expect(wb.mergeAt(a, 6, 3)).toEqual(rg(6, 1, 6, 3));
    expect(wb.meta(a).colW.get(1)).toBe(245);
    expect(wb.meta(a).rowH.get(0)).toBe(40);
    expect(wb.meta(a).freeze).toEqual({ rows: 1, cols: 1 });
    expect(wb.cfStyle(a, 3, 3)).toMatchObject({ bg: '#ffc7ce', fg: '#9c0006' });
    expect(wb.cfStyle(a, 1, 3)).toEqual({});
    expect(wb.cfStyle(a, 1, 4).bg).toBeDefined();
    expect(wb.dvList(a, 9, 2)).toEqual(['bộ', 'cột', 'm']);
    expect(wb.validate(a, 3, 3, '5000')).not.toBeNull();
    expect(wb.meta(a).filter?.range).toEqual(rg(0, 0, 3, 6));
    expect(wb.value(b, 1, 0)).toBe(wb.value(a, 4, 5) as number / 2);
    expect(wb.value(c, 0, 0)).toBe(10);
    expect(warnings).toEqual([]);
  });
});

describe('Lưu .xlsx rồi mở lại (khứ hồi)', () => {
  it('giữ nguyên giá trị, công thức, định dạng, gộp ô, kích thước, CF, DV, lọc', async () => {
    const { wb, s } = book([['Tên', 'SL', 'Giá', 'Ngày'], ['LED', '2', '1.250.000,5', '01/02/2026'], ['Cột', '3', '15%', '=D2+1'], ['', 'Tổng', '=SUM(C2:C3)', '']]);
    wb.setStyle(s, rg(0, 0, 0, 3), { b: true, bg: '#dcefe5', fg: '#127a4e', al: 'c', fs: 14, fn: 'Arial' });
    wb.setStyle(s, one(1, 2), { fmt: '#,##0.00" ₫"' });
    wb.setStyle(s, one(2, 3), { fmt: 'dd/mm/yyyy' });
    wb.setStyle(s, one(1, 0), { wr: true, va: 't', i: true, u: true, s: true });
    wb.setBorder(s, rg(0, 0, 3, 3), 'all', 'thin');
    wb.merge(s, rg(5, 0, 5, 3));
    wb.setColWidth(s, 0, 150); wb.setRowHeight(s, 0, 36); wb.setFreeze(s, 1, 0);
    wb.addCF(s, { range: rg(1, 1, 2, 1), type: 'cellIs', op: 'gt', v1: '2', fmt: { bg: '#ffc7ce' } });
    wb.addDV(s, { range: rg(1, 1, 9, 1), type: 'whole', op: 'between', v1: '0', v2: '10' });
    wb.setFilter(s, rg(0, 0, 3, 3));
    const s2 = wb.addSheet('2024'); wb.setInput(s2, 0, 0, "='Sheet1'!C4*2");
    const buf = await writeXlsx(wb);
    const { wb: back } = await readXlsx(buf);
    const [b1, b2] = back.sheetIds();
    for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) expect(back.display(b1, r, c).text).toBe(wb.display(s, r, c).text);
    expect(back.rawText(b1, 3, 2)).toBe('=SUM(C2:C3)');
    expect(back.style(b1, 0, 0)).toMatchObject({ b: true, bg: '#dcefe5', fg: '#127a4e', al: 'c', fs: 14, fn: 'Arial' });
    expect(back.style(b1, 1, 0)).toMatchObject({ wr: true, va: 't', i: true, u: true, s: true });
    expect(back.style(b1, 2, 2).bd).toMatchObject({ t: { s: 'thin' }, r: { s: 'thin' } });
    expect(back.mergeAt(b1, 5, 2)).toEqual(rg(5, 0, 5, 3));
    expect(back.meta(b1).colW.get(0)).toBe(150);
    expect(back.meta(b1).rowH.get(0)).toBe(36);
    expect(back.meta(b1).freeze).toEqual({ rows: 1, cols: 0 });
    expect(back.cfStyle(b1, 2, 1)).toMatchObject({ bg: '#ffc7ce' });
    expect(back.validate(b1, 4, 1, '11')).not.toBeNull();
    expect(back.meta(b1).filter?.range).toEqual(rg(0, 0, 3, 3));
    expect(back.sheetName(b2)).toBe('2024');
    expect(back.value(b2, 0, 0)).toBe((wb.value(s, 3, 2) as number) * 2);
  });
});
