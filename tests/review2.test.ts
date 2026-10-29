// Ca kiểm thử bổ sung từ lần review độc lập Đợt 2
import ExcelJS from 'exceljs';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { toEn, toVi } from '../src/core/locale';

const rg = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
async function excelFile(build: (ws: ExcelJS.Worksheet, wb: ExcelJS.Workbook) => void) {
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('S'); build(ws, wb);
  return new Uint8Array(await wb.xlsx.writeBuffer() as ArrayBuffer);
}

describe('Review Đợt 2: dữ liệu chữ không bị đổi khi mở/lưu', () => {
  it('chữ trông như số, công thức, tiền, số mũ, dấu nháy giữ nguyên', async () => {
    const texts = ['=abc text', '+84989725507', '1E5', '$100', "'quoted", '0123', 'TRUE', '01/02/2026', '15%', 'bình thường'];
    const buf = await excelFile((ws) => texts.forEach((t, i) => (ws.getCell(i + 1, 1).value = t)));
    const { wb } = await readXlsx(buf);
    texts.forEach((t, i) => expect(wb.value(0, i, 0)).toBe(t));
    const { wb: back } = await readXlsx(await writeXlsx(wb));
    texts.forEach((t, i) => expect(back.value(0, i, 0)).toBe(t));
  });
});

describe('Review Đợt 2: kiểm tra dữ liệu và CF ghi đúng chuẩn Excel', () => {
  it('ngày, số kiểu Việt, ô tham chiếu, danh sách có dấu phẩy và rất dài', async () => {
    const wb = Workbook.create(); const s = 0;
    wb.addDV(s, { range: rg(0, 0, 9, 0), type: 'date', op: 'between', v1: '01/01/2026', v2: '31/12/2026' });
    wb.addDV(s, { range: rg(0, 1, 9, 1), type: 'decimal', op: 'ge', v1: '1.000.000' });
    wb.addDV(s, { range: rg(0, 2, 9, 2), type: 'list', list: ['Hà Nội, Q1', 'TP.HCM', 'say "hi"'] });
    const long = Array.from({ length: 80 }, (_, i) => 'Mục ' + i);
    wb.addDV(s, { range: rg(0, 3, 9, 3), type: 'list', list: long });
    wb.addCF(s, { range: rg(0, 4, 9, 4), type: 'cellIs', op: 'gt', v1: '1.000.000', fmt: { bg: '#ffc7ce' } });
    wb.addCF(s, { range: rg(0, 5, 9, 5), type: 'text', text: 'a"b', fmt: { b: true } });
    const buf = await writeXlsx(wb);
    const { wb: back } = await readXlsx(buf);
    expect(back.sheetIds().length).toBe(1);
    expect(back.validate(0, 2, 0, '15/07/2026')).toBeNull();
    expect(back.validate(0, 2, 0, '15/07/2025')).not.toBeNull();
    expect(back.validate(0, 2, 1, '999.999')).not.toBeNull();
    expect(back.validate(0, 2, 1, '2.000.000')).toBeNull();
    expect(back.dvList(0, 1, 2)).toEqual(['Hà Nội, Q1', 'TP.HCM', 'say "hi"']);
    expect(back.dvList(0, 1, 3)).toEqual(long);
    back.setInput(0, 0, 4, '2000000'); back.setInput(0, 1, 4, '5');
    expect(back.cfStyle(0, 0, 4).bg).toBe('#ffc7ce');
    expect(back.cfStyle(0, 1, 4).bg).toBeUndefined();
    back.setInput(0, 0, 5, 'xa"bx');
    expect(back.cfStyle(0, 0, 5).b).toBe(true);
    // XML không có NaN, không có lỗi giả
    const zipText = new TextDecoder().decode(buf);
    expect(zipText.includes('NaN')).toBe(false);
  });
  it('kiểm tra dữ liệu cả cột C:C đọc nhanh, đúng vùng', async () => {
    const buf = await excelFile((ws) => {
      ws.getCell('A1').value = 'x';
      for (const col of ['C', 'D', 'E', 'F']) (ws as any).dataValidations.add(`${col}1:${col}1048576`, { type: 'whole', operator: 'between', formulae: [0, 10], allowBlank: true });
    });
    const t = performance.now();
    const { wb } = await readXlsx(buf);
    expect(performance.now() - t).toBeLessThan(3000);
    expect(wb.meta(0).dv.length).toBe(4);
    expect(wb.validate(0, 500000, 2, '11')).not.toBeNull();
  });
  it('danh sách tham chiếu tên vùng không chặn nhập', () => {
    const wb = Workbook.create();
    wb.addDV(0, { range: rg(0, 0, 5, 0), type: 'list', listRef: 'DanhMuc' });
    expect(wb.validate(0, 1, 0, 'gì cũng được')).toBeNull();
  });
});

describe('Review Đợt 2: sắp xếp', () => {
  it('giữ công thức tham chiếu sheet tên số', () => {
    const wb = Workbook.create(); const y = wb.addSheet('2024'); wb.setInput(y, 0, 0, '5');
    wb.setInput(0, 0, 0, 'k'); wb.setInput(0, 1, 0, "='2024'!$A$1+1"); wb.setInput(0, 2, 0, '1');
    wb.sort(0, rg(0, 0, 2, 0), [{ col: 0 }], true);
    expect(wb.value(0, 2, 0)).toBe(6);
  });
  it('khi đang lọc chỉ sắp xếp hàng đang hiện', () => {
    const wb = Workbook.create();
    [['K', 'V'], ['a', '1'], ['b', '9'], ['a', '3'], ['b', '7']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v)));
    wb.setFilter(0, rg(0, 0, 4, 1));
    wb.filterColumn(0, 0, { values: ['a'] });
    wb.sort(0, rg(0, 0, 4, 1), [{ col: 1, desc: true }], true);
    const visible = [1, 2, 3, 4].filter((r) => !wb.isRowHidden(0, r)).map((r) => wb.display(0, r, 0).text + wb.display(0, r, 1).text);
    expect(visible).toEqual(['a3', 'a1']);
  });
});

describe('Review Đợt 2: lọc và kích thước hàng', () => {
  it('bật lọc cho cả cột chỉ lấy vùng có dữ liệu', () => {
    const wb = Workbook.create();
    ['T', '1', '2'].forEach((v, r) => wb.setInput(0, r, 0, v));
    wb.setFilter(0, rg(0, 0, 1048575, 0));
    expect(wb.meta(0).filter!.range.r2).toBe(2);
  });
  it('hoàn tác đổi cỡ hàng không làm ẩn vĩnh viễn hàng đang bị lọc', () => {
    const wb = Workbook.create();
    [['K'], ['a'], ['b']].forEach((row, r) => wb.setInput(0, r, 0, row[0]));
    wb.setFilter(0, rg(0, 0, 2, 0)); wb.filterColumn(0, 0, { values: ['a'] });
    wb.setRowHeight(0, 2, 40); wb.undo();
    wb.clearFilter(0);
    expect(wb.meta(0).rowH.get(2)).toBe(20);
  });
  it('hàng ẩn do lọc trong tệp Excel hiện lại khi bỏ lọc', async () => {
    const buf = await excelFile((ws) => { ['K', 'a', 'b'].forEach((v, i) => (ws.getCell(i + 1, 1).value = v)); ws.autoFilter = 'A1:A3'; ws.getRow(3).hidden = true; });
    const { wb } = await readXlsx(buf);
    expect(wb.isRowHidden(0, 2)).toBe(true);
    wb.clearFilter(0);
    expect(wb.meta(0).rowH.get(2)).toBeGreaterThan(0);
  });
});

describe('Review Đợt 2: ô gộp giữ viền, hằng mảng, CSV', () => {
  it('ô gộp có viền ngoài giữ đủ 4 cạnh sau khi lưu', async () => {
    const wb = Workbook.create();
    wb.setBorder(0, rg(0, 0, 0, 2), 'outer', 'thin');
    wb.merge(0, rg(0, 0, 0, 2));
    const { wb: back } = await readXlsx(await writeXlsx(wb));
    expect(back.style(0, 0, 2).bd?.r).toBeDefined();
    expect(back.style(0, 0, 0).bd?.l).toBeDefined();
  });
  it('hằng mảng đổi đúng dấu hàng/cột', () => {
    expect(toVi('={1,2;3,4}')).toBe('={1,2|3,4}');
    expect(toEn('={1,2|3,4}')).toBe('={1,2;3,4}');
  });
  it('CSV dấu phẩy: "1,500" là 1500; xuất số không làm tròn, dấu chấm', () => {
    const wb = Workbook.create();
    const id = wb.importCSV('A,B\n"1,500",0.333333333333\n');
    expect(wb.value(id, 1, 0)).toBe(1500);
    wb.setInput(id, 2, 0, '=1/3');
    expect(wb.exportCSV(id, ',')).toMatch(/\n0\.33333333333333/);
  });
  it('lỗi chuẩn Excel ghi đúng kiểu, lỗi riêng của app không ghi', async () => {
    const wb = Workbook.create();
    wb.setInput(0, 0, 0, '=1/0'); wb.setInput(0, 1, 0, '=A2+1'); // vòng lặp
    const buf = await writeXlsx(wb);
    const x = new ExcelJS.Workbook(); await x.xlsx.load(buf as any);
    expect((x.worksheets[0].getCell('A1').value as any).result).toEqual({ error: '#DIV/0!' });
    expect((x.worksheets[0].getCell('A2').value as any).result).toBeUndefined();
  });
});
