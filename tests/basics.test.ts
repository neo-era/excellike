import { colName, colIndex, toA1, parseA1, parseRange, rangeToA1 } from '../src/core/address';
import { formatGeneral, formatNumber, formatValue } from '../src/core/numfmt';
import { normalizeInput } from '../src/core/input';
import { toTSV, parseTSV } from '../src/core/tsv';
import { AxisSizes } from '../src/core/sizes';
import { fillPlan } from '../src/core/series';

describe('Địa chỉ A1', () => {
  it('đổi chỉ số cột ↔ chữ', () => {
    expect([0, 25, 26, 51, 701, 702, 16383].map(colName)).toEqual(['A', 'Z', 'AA', 'AZ', 'ZZ', 'AAA', 'XFD']);
    expect(['A', 'Z', 'AA', 'XFD'].map(colIndex)).toEqual([0, 25, 26, 16383]);
  });
  it('toA1 / parseA1, chấp nhận $ và chữ thường', () => {
    expect(toA1(0, 0)).toBe('A1');
    expect(toA1(9, 27)).toBe('AB10');
    expect(parseA1('$b$3')).toEqual({ r: 2, c: 1 });
    expect(parseA1('A0')).toBeNull();
    expect(parseA1('XFE1')).toBeNull();
    expect(parseA1('A1048577')).toBeNull();
    expect(parseA1('hello')).toBeNull();
  });
  it('vùng được chuẩn hóa', () => {
    expect(parseRange('C5:A1')).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
    expect(parseRange('B2')).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
    expect(rangeToA1({ r1: 0, c1: 0, r2: 4, c2: 2 })).toBe('A1:C5');
    expect(rangeToA1({ r1: 1, c1: 1, r2: 1, c2: 1 })).toBe('B2');
  });
});

describe('Hiển thị số kiểu Việt Nam', () => {
  it('General giống Excel (tối đa 11 ký tự)', () => {
    expect(formatGeneral(1.5)).toBe('1,5');
    expect(formatGeneral(1250)).toBe('1250');
    expect(formatGeneral(0.1 + 0.2)).toBe('0,3');
    expect(formatGeneral(1 / 3)).toBe('0,333333333');
    expect(formatGeneral(-0.5)).toBe('-0,5');
    expect(formatGeneral(123456789012)).toBe('1,23457E+11');
    expect(formatGeneral(1e-10)).toBe('1E-10');
    expect(formatGeneral(0)).toBe('0');
  });
  it('định dạng số có nhóm nghìn, %, ngày', () => {
    expect(formatNumber(1234567.891, '#,##0')).toBe('1.234.568');
    expect(formatNumber(1234567.891, '#,##0.00')).toBe('1.234.567,89');
    expect(formatNumber(-1234, '#,##0')).toBe('-1.234');
    expect(formatNumber(0.153, '0%')).toBe('15%');
    expect(formatNumber(0.153, '0.0%')).toBe('15,3%');
    expect(formatNumber(46054, 'dd/mm/yyyy')).toBe('01/02/2026');
    expect(formatNumber(2.5, '0.00')).toBe('2,50');
    expect(formatNumber(1250000, '#,##0" ₫"')).toBe('1.250.000 ₫');
  });
  it('căn lề mặc định theo kiểu dữ liệu', () => {
    expect(formatValue(5).align).toBe('r');
    expect(formatValue('abc').align).toBe('l');
    expect(formatValue(true)).toEqual({ text: 'TRUE', align: 'c' });
    expect(formatValue(null).text).toBe('');
    expect(formatValue({ value: '#DIV/0!' }).text).toBe('#DIV/0!');
  });
});

describe('Chuẩn hóa dữ liệu nhập', () => {
  it('số kiểu Việt Nam có dấu chấm nhóm nghìn', () => {
    expect(normalizeInput('1.250').content).toBe(1250);
    expect(normalizeInput('-1.250.000,5').content).toBe(-1250000.5);
  });
  it('phần trăm và ngày tự gắn định dạng', () => {
    expect(normalizeInput('15%')).toEqual({ content: '15%', fmt: '0%' });
    expect(normalizeInput('12,5%')).toEqual({ content: '12,5%', fmt: '0.0%' });
    expect(normalizeInput('01/02/2026').fmt).toBe('dd/mm/yyyy');
  });
  it('công thức, chữ, ô rỗng', () => {
    expect(normalizeInput('=SUM(A1:A2)').content).toBe('=SUM(A1:A2)');
    expect(normalizeInput('abc').content).toBe('abc');
    expect(normalizeInput('').content).toBeNull();
  });
});

describe('Clipboard dạng TSV', () => {
  it('khứ hồi giữ nguyên dữ liệu có tab, xuống dòng, ngoặc kép', () => {
    const d = [['a', 'b\tc'], ['x"y', 'dòng 1\ndòng 2']];
    expect(parseTSV(toTSV(d))).toEqual(d);
  });
  it('đọc TSV từ Excel (CRLF, dòng cuối trống)', () => {
    expect(parseTSV('1\t2\r\n3\t4\r\n')).toEqual([['1', '2'], ['3', '4']]);
  });
});

describe('Kích thước hàng/cột', () => {
  it('offset và tìm chỉ số theo pixel', () => {
    const a = new AxisSizes(10, 1000);
    a.set(2, 30);
    a.set(5, 0);
    expect(a.offset(0)).toBe(0);
    expect(a.offset(3)).toBe(50);
    expect(a.offset(6)).toBe(70);
    expect(a.indexAt(49)).toBe(2);
    expect(a.indexAt(50)).toBe(3);
    expect(a.indexAt(65)).toBe(4);
    expect(a.indexAt(70)).toBe(6);
    expect(a.total()).toBe(10000 + 20 - 10);
  });
  it('dịch kích thước khi chèn/xóa', () => {
    const a = new AxisSizes(10, 100);
    a.set(5, 40);
    a.shiftInsert(2, 3);
    expect(a.get(8)).toBe(40);
    a.shiftDelete(0, 4);
    expect(a.get(4)).toBe(40);
    a.shiftDelete(4, 1);
    expect(a.get(4)).toBe(10);
  });
});

describe('Kéo điền', () => {
  const num = (n: number) => ({ raw: n, num: n });
  it('dãy số tuyến tính từ 2 giá trị', () => {
    expect(fillPlan([num(1), num(3)], 3, true)).toEqual([{ value: 5 }, { value: 7 }, { value: 9 }]);
    expect(fillPlan([num(10), num(20)], 2, false)).toEqual([{ value: 0 }, { value: -10 }]);
  });
  it('một số đơn lẻ thì sao chép', () => {
    expect(fillPlan([num(7)], 2, true)).toEqual([{ copy: 0 }, { copy: 0 }]);
  });
  it('ngày đơn lẻ tăng từng ngày, chữ có số cuối tăng dần', () => {
    expect(fillPlan([{ raw: 46054, num: 46054, isDate: true }], 2, true)).toEqual([{ value: 46055 }, { value: 46056 }]);
    expect(fillPlan([{ raw: 'Tủ 9' }], 2, true)).toEqual([{ value: 'Tủ 10' }, { value: 'Tủ 11' }]);
  });
  it('công thức và chữ lặp theo mẫu', () => {
    expect(fillPlan([{ raw: '=A1', formula: true }, { raw: 'x' }], 3, true)).toEqual([{ copy: 0 }, { copy: 1 }, { copy: 0 }]);
    expect(fillPlan([{ raw: 'a' }, { raw: 'b' }], 3, false)).toEqual([{ copy: 1 }, { copy: 0 }, { copy: 1 }]);
  });
});

import { shiftFormula } from '../src/core/formula';
describe('Dời tham chiếu công thức', () => {
  it('tương đối dời, tuyệt đối giữ', () => {
    expect(shiftFormula('=A1*$D$1', 1, 0)).toBe('=A2*$D$1');
    expect(shiftFormula('=SUM(A$1:B2)', 2, 1)).toBe('=SUM(B$1:C4)');
    expect(shiftFormula('=$A1+B$1', 3, 3)).toBe('=$A4+E$1');
  });
  it('không đụng chuỗi, tên hàm, tên sheet', () => {
    expect(shiftFormula('="A1"&A1', 1, 0)).toBe('="A1"&A2');
    expect(shiftFormula('=LOG10(A1)', 0, 1)).toBe('=LOG10(B1)');
    expect(shiftFormula("='Dự toán'!B2+Sheet2!C3", 1, 0)).toBe("='Dự toán'!B3+Sheet2!C4");
  });
  it('cột/hàng nguyên và tràn biên thành #REF!', () => {
    expect(shiftFormula('=SUM(A:B)+SUM(1:2)', 1, 1)).toBe('=SUM(B:C)+SUM(2:3)');
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!');
  });
});

describe('Mã định dạng Excel đầy đủ hơn (Đợt 2)', () => {
  it('nhiều phần, màu, chữ cố định, khoảng trống', () => {
    expect(formatNumber(-1234.5, '#,##0.00;[Red]-#,##0.00')).toBe('-1.234,50');
    expect(formatNumber(-1234.5, '#,##0;(#,##0)')).toBe('(1.235)');
    expect(formatNumber(0, '#,##0;-#,##0;"-"')).toBe('-');
    expect(formatNumber(1234, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)')).toBe(' 1.234 ');
    expect(formatNumber(1250000, '#,##0 [$₫-42A]')).toBe('1.250.000 ₫');
    expect(formatNumber(0.5, '0.0#')).toBe('0,5');
    expect(formatNumber(12345.678, '0.00E+00')).toBe('1,23E+04');
    expect(formatNumber(1500000, '#,##0.0,,"tr"')).toBe('1,5tr');
  });
  it('giờ, ngày giờ, tháng chữ, thời gian cộng dồn', () => {
    expect(formatNumber(0.5208333333, 'h:mm:ss')).toBe('12:30:00');
    expect(formatNumber(46218.53125, 'yyyy-mm-dd h:mm:ss')).toBe('2026-07-15 12:45:00');
    expect(formatNumber(46218, 'mmm-yy')).toBe('thg 7-26');
    expect(formatNumber(46218, 'dd/mm/yyyy')).toBe('15/07/2026');
    expect(formatNumber(1.5, '[h]:mm')).toBe('36:00');
    expect(formatNumber(0.75, 'h:mm AM/PM')).toBe('6:00 CH');
  });
});
