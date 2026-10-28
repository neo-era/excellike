// Ca kiểm thử bổ sung từ lần review độc lập (Bước 4)
import { Workbook } from '../src/core/workbook';
import { shiftFormula, fixSheetQuotes } from '../src/core/formula';
import { fillPlan } from '../src/core/series';

const one = (r: number, c: number) => ({ r1: r, c1: c, r2: r, c2: c });

describe('Review #1: hoàn tác định dạng sau thao tác cấu trúc', () => {
  it('đậm A1 → chèn hàng → hoàn tác 2 lần → hết đậm; làm lại được', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'x');
    wb.setStyle(s, one(0, 0), { b: true });
    wb.insertRows(s, 5, 1);
    wb.undo(); wb.undo();
    expect(wb.style(s, 0, 0).b).toBeFalsy();
    wb.redo();
    expect(wb.style(s, 0, 0).b).toBe(true);
  });
  it('ngày tự gắn định dạng → dán → hoàn tác 2 lần → ô sạch định dạng', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '01/02/2024');
    wb.copy(s, one(0, 0)); wb.paste(s, 3, 3);
    wb.undo(); wb.undo();
    expect(wb.style(s, 0, 0).fmt).toBeUndefined();
    wb.setInput(s, 0, 0, '45000');
    expect(wb.display(s, 0, 0).text).toBe('45000');
  });
  it('kéo điền rồi chèn cột, hoàn tác hết vẫn đúng', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '1'); wb.setStyle(s, one(0, 0), { i: true });
    wb.fill(s, one(0, 0), { r1: 0, c1: 0, r2: 3, c2: 0 });
    wb.insertCols(s, 0, 1);
    expect(wb.style(s, 3, 1).i).toBe(true);
    wb.undo(); wb.undo();
    expect(wb.style(s, 3, 0).i).toBeFalsy();
    expect(wb.style(s, 0, 0).i).toBe(true);
  });
});

describe('Review #2: tên sheet đặc biệt khi lưu/mở', () => {
  it('sheet tên __proto__ và tên toàn số vẫn mở lại được, đúng thứ tự', () => {
    const wb = Workbook.create();
    const a = wb.addSheet('__proto__'); wb.setInput(a, 0, 0, '5');
    const b = wb.addSheet('2024'); wb.setInput(b, 0, 0, '7');
    const back = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(back.sheetIds().map((i) => back.sheetName(i))).toEqual(['Sheet1', '__proto__', '2024']);
    expect(back.value(back.sheetIds()[2], 0, 0)).toBe(7);
  });
});

describe('Review #3/#4: tên sheet giống số hoặc địa chỉ ô', () => {
  it('công thức tham chiếu sheet 2024, 1Q, Q1 giữ đúng khi lưu/mở và khi kéo điền', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    for (const n of ['2024', '1Q', 'Q1']) { const id = wb.addSheet(n); wb.setInput(id, 0, 0, '10'); wb.setInput(id, 1, 0, '20'); }
    wb.setInput(s, 0, 0, "='2024'!A1+'1Q'!A1+'Q1'!A1");
    expect(wb.value(s, 0, 0)).toBe(30);
    expect(wb.rawText(s, 0, 0)).toBe("='2024'!A1+'1Q'!A1+'Q1'!A1");
    wb.fill(s, one(0, 0), { r1: 0, c1: 0, r2: 1, c2: 0 });
    expect(wb.value(s, 1, 0)).toBe(60);
    const back = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(back.value(back.sheetIds()[0], 1, 0)).toBe(60);
  });
  it('shiftFormula không đổi tên sheet giống địa chỉ ô', () => {
    expect(shiftFormula('=Q1!A1+T10!B2', 1, 0)).toBe('=Q1!A2+T10!B3');
  });
  it('fixSheetQuotes không đụng chuỗi và tên đã có nháy', () => {
    expect(fixSheetQuotes('=2024!A1&"2024!A1"', ['2024'])).toBe(`='2024'!A1&"2024!A1"`);
    expect(fixSheetQuotes("='Báo cáo Q1'!A1+Q1!B1", ['Q1'])).toBe("='Báo cáo Q1'!A1+'Q1'!B1");
  });
});

describe('Review #5: tính lại toàn bộ', () => {
  it('sau F9 không còn mục hoàn tác lệch', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '1'); wb.setInput(s, 0, 0, '2');
    wb.recalc();
    expect(wb.canUndo()).toBe(false);
    wb.setInput(s, 0, 0, '3'); wb.undo();
    expect(wb.value(s, 0, 0)).toBe(2);
  });
});

describe('Review #6: thao tác lỗi không làm lệch định dạng', () => {
  it('chèn hàng khi đã có dữ liệu ở hàng cuối → báo lỗi, định dạng giữ nguyên', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'x'); wb.setStyle(s, one(0, 0), { b: true });
    wb.setInput(s, 1048575, 0, 'cuối');
    expect(() => wb.insertRows(s, 0, 1)).toThrow();
    expect(wb.style(s, 0, 0).b).toBe(true);
    expect(wb.style(s, 1, 0).b).toBeFalsy();
  });
});

describe('Review #7/#9: xóa sheet', () => {
  it('xóa sheet → thêm sheet (dùng lại mã) → hoàn tác 2 lần → định dạng sheet cũ trở lại', () => {
    const wb = Workbook.create();
    const s2 = wb.addSheet(); wb.setInput(s2, 0, 0, 'x'); wb.setStyle(s2, one(0, 0), { b: true, bg: '#ff0' });
    wb.removeSheet(s2);
    const s3 = wb.addSheet();
    expect(wb.style(s3, 0, 0).b).toBeFalsy();
    wb.undo(); wb.undo();
    const back = wb.sheetIds().find((i) => wb.sheetName(i) === 'Sheet2')!;
    expect(wb.style(back, 0, 0)).toMatchObject({ b: true, bg: '#ff0' });
  });
  it('sao chép trên sheet đã xóa thì không dán nhầm', () => {
    const wb = Workbook.create(); const s1 = wb.sheetIds()[0];
    const s2 = wb.addSheet(); wb.setInput(s2, 0, 0, 'x');
    wb.copy(s2, one(0, 0));
    wb.removeSheet(s2);
    expect(wb.paste(s1, 0, 0)).toBeNull();
  });
});

describe('Review #12: kéo điền giữ số 0 đứng đầu', () => {
  it('A01 → A02, A03', () => {
    expect(fillPlan([{ raw: 'A01' }], 2, true)).toEqual([{ value: 'A02' }, { value: 'A03' }]);
  });
});
