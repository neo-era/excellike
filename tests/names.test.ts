import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { encName, decName, encFormulaNames, renameInFormula, nameError } from '../src/core/names';

const mk = () => { const wb = Workbook.create(); const s = wb.sheetIds()[0]; [[10, 20], [30, 40]].forEach((r, i) => r.forEach((v, c) => wb.setInput(s, i, c, String(v)))); return { wb, s }; };

describe('Tên vùng', () => {
  it('mã hóa tên có dấu thanh, khứ hồi, không phân biệt hoa thường', () => {
    expect(encName('DonGia')).toBe('DonGia');
    expect(decName(encName('Đơn_giả'))).toBe('Đơn_giả');
    expect(encName('GIẢ').toUpperCase()).toBe(encName('giả').toUpperCase());
    expect(encFormulaNames('=SUM(Đơn_giá)*"giá trị"')).toBe('=SUM(' + encName('Đơn_giá') + ')*"giá trị"');
    expect(encFormulaNames("='Dự toán'!A1+Giá")).toBe("='Dự toán'!A1+" + encName('Giá'));
  });
  it('kiểm tra tên như Excel', () => {
    expect(nameError('DonGia')).toBeNull();
    expect(nameError('Đơn_giá')).toBeNull();
    expect(nameError('A1')).toMatch(/địa chỉ/);
    expect(nameError('1abc')).toMatch(/bắt đầu/);
    expect(nameError('a b')).toMatch(/dấu cách/);
    expect(nameError('R1C1')).toMatch(/địa chỉ/);
  });
  it('đổi tên trong công thức, bỏ qua chuỗi và hàm', () => {
    expect(renameInFormula('=SUM(gia)+Gia*2&"Gia"', 'Gia', 'DonGia')).toBe('=SUM(DonGia)+DonGia*2&"Gia"');
  });
  it('xác định tên, dùng trong công thức, tên tiếng Việt có dấu', () => {
    const { wb, s } = mk();
    wb.defineName('Đơn_giá', '=Sheet1!$A$1:$B$2');
    wb.setInput(s, 3, 0, '=SUM(Đơn_giá)');
    expect(wb.value(s, 3, 0)).toBe(100);
    expect(wb.rawText(s, 3, 0)).toBe('=SUM(Đơn_giá)');
    expect(wb.listNames()).toEqual([{ name: 'Đơn_giá', expr: '=Sheet1!$A$1:$B$2', scope: undefined, comment: undefined, hidden: undefined }]);
    expect(() => wb.defineName('đơn_giá', '=Sheet1!$A$1')).toThrow(/đã có/);
    expect(wb.nameTarget('đơn_giá')).toEqual({ sheet: s, range: { r1: 0, c1: 0, r2: 1, c2: 1 } });
  });
  it('công thức nhập trước khi có tên → #NAME?, có tên thì tính lại', () => {
    const { wb, s } = mk();
    wb.setInput(s, 3, 0, '=Gia*2');
    expect((wb.value(s, 3, 0) as any).value).toBe('#NAME?');
    wb.defineName('Gia', '=Sheet1!$B$1');
    expect(wb.value(s, 3, 0)).toBe(40);
  });
  it('đổi tên cập nhật công thức; xóa tên → #NAME?; hoàn tác', () => {
    const { wb, s } = mk();
    wb.defineName('Gia', '=Sheet1!$B$1', undefined, 'giá vật tư');
    wb.setInput(s, 3, 0, '=Gia+1');
    wb.updateName('Gia', undefined, { name: 'GiaVT', expr: '=Sheet1!$B$2', comment: 'mới' });
    expect(wb.rawText(s, 3, 0)).toBe('=GiaVT+1');
    expect(wb.value(s, 3, 0)).toBe(41);
    expect(wb.listNames()[0]).toMatchObject({ name: 'GiaVT', comment: 'mới' });
    wb.undo();
    expect(wb.rawText(s, 3, 0)).toBe('=Gia+1');
    expect(wb.value(s, 3, 0)).toBe(21);
    wb.deleteName('Gia');
    expect((wb.value(s, 3, 0) as any).value).toBe('#NAME?');
    wb.undo();
    expect(wb.value(s, 3, 0)).toBe(21);
  });
  it('tên phạm vi trang tính', () => {
    const { wb, s } = mk();
    const s2 = wb.addSheet('Q1');
    wb.defineName('X', '=Sheet1!$A$1', s);
    wb.defineName('X', '=Sheet1!$B$1', s2);
    wb.setInput(s, 3, 0, '=X'); wb.setInput(s2, 0, 0, '=X');
    expect(wb.value(s, 3, 0)).toBe(10);
    expect(wb.value(s2, 0, 0)).toBe(20);
  });
  it('Tạo từ Vùng chọn (hàng trên) và lưu/mở lại sổ', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    [['Khối lượng', 'Đơn giá'], ['2', '100'], ['3', '200']].forEach((r, i) => r.forEach((v, c) => wb.setInput(s, i, c, v)));
    expect(wb.createNamesFrom(s, { r1: 0, c1: 0, r2: 2, c2: 1 }, { top: true })).toBe(2);
    wb.setInput(s, 4, 0, '=SUMPRODUCT(Khối_lượng;Đơn_giá)');
    expect(wb.value(s, 4, 0)).toBe(800);
    const wb2 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(wb2.value(s, 4, 0)).toBe(800);
    expect(wb2.listNames().map((n) => n.name)).toEqual(['Đơn_giá', 'Khối_lượng']);
    expect(wb2.rawText(s, 4, 0)).toBe('=SUMPRODUCT(Khối_lượng;Đơn_giá)');
  });
});

import { readXlsx, writeXlsx } from '../src/core/xlsx';
import JSZip from 'jszip';
describe('Tên vùng trong .xlsx', () => {
  it('ghi definedNames (cả phạm vi trang tính, chú thích) và đọc lại', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.renameSheet(s, 'Dự toán');
    [['2', '100'], ['3', '200']].forEach((r, i) => r.forEach((v, c) => wb.setInput(s, i, c, v)));
    wb.defineName('KhoiLuong', "='Dự toán'!$A$1:$A$2", undefined, 'khối lượng');
    wb.defineName('Đơn_giá', "='Dự toán'!$B$1:$B$2", s);
    wb.setInput(s, 3, 0, '=SUMPRODUCT(KhoiLuong;Đơn_giá)');
    const buf = await writeXlsx(wb);
    const xml = await (await JSZip.loadAsync(buf)).file('xl/workbook.xml')!.async('string');
    expect(xml).toContain('<definedName name="KhoiLuong" comment="khối lượng">\'Dự toán\'!$A$1:$A$2</definedName>');
    expect(xml).toContain('<definedName name="Đơn_giá" localSheetId="0">');
    const { wb: w2 } = await readXlsx(buf);
    const s2 = w2.sheetIds()[0];
    expect(w2.value(s2, 3, 0)).toBe(800);
    expect(w2.listNames().map((n) => [n.name, n.scope !== undefined])).toEqual([['Đơn_giá', true], ['KhoiLuong', false]]);
  });
});
