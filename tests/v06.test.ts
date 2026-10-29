// Bản 0.6 – phím tắt giống hệt Excel 365 và các tính năng đi kèm (phần lõi):
// Tìm/Thay, Đi tới Đặc biệt, Dán Đặc biệt, F4 đổi tham chiếu, Điền Nhanh, Điền Chuỗi, liên kết, kiểu ô mở rộng.
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { findAll, findNext, replaceAll, goSpecial, toRuns } from '../src/core/find';
import { cycleRef } from '../src/core/formula';
import { flashFill } from '../src/core/flashfill';
import { seriesValues } from '../src/core/series';
import { formatValue } from '../src/core/numfmt';
import { KEYMAP, comboOf, keysOk } from '../src/ui/keymap';
import { NF_SHORTCUT } from '../src/core/numfmt';

const rg = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
function book(rows: (string | number)[][]) {
  const wb = Workbook.create();
  const s = wb.sheetIds()[0];
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setInput(s, r, c, String(v)); }));
  return { wb, s };
}
const A = (a: string) => { const m = /^([A-Z]+)(\d+)$/.exec(a)!; return { r: +m[2] - 1, c: m[1].charCodeAt(0) - 65 }; };
const val = (wb: Workbook, s: number, a: string) => { const p = A(a); return wb.value(s, p.r, p.c); };
const raw = (wb: Workbook, s: number, a: string) => { const p = A(a); return wb.rawText(s, p.r, p.c); };

describe('0.6 – bảng ánh xạ phím', () => {
  it('không có tổ hợp trùng trong cùng ngữ cảnh', () => {
    const seen = new Map<string, string>();
    for (const k of keysOk()) for (const c of k.combos) for (const ctx of k.ctx === 'any' ? ['nav', 'edit'] : [k.ctx]) {
      const key = ctx + '|' + c;
      expect(seen.get(key), `${c} (${ctx}) trùng giữa ${seen.get(key)} và ${k.id}`).toBeUndefined();
      seen.set(key, k.id);
    }
    expect(KEYMAP.length).toBeGreaterThan(150);
  });
  it('chuẩn hoá phím theo vị trí: Ctrl+Shift+$ = Ctrl+Shift+4', () => {
    const ev = (o: Partial<KeyboardEvent>) => ({ ctrlKey: false, altKey: false, shiftKey: false, metaKey: false, key: '', code: '', ...o }) as KeyboardEvent;
    expect(comboOf(ev({ ctrlKey: true, shiftKey: true, key: '$', code: 'Digit4' }))).toBe('Ctrl+Shift+4');
    expect(comboOf(ev({ ctrlKey: true, key: ';', code: 'Semicolon' }))).toBe('Ctrl+;');
    expect(comboOf(ev({ shiftKey: true, key: 'F10', code: 'F10' }))).toBe('Shift+F10');
    expect(comboOf(ev({ ctrlKey: true, key: ' ', code: 'Space' }))).toBe('Ctrl+Space');
    expect(comboOf(ev({ ctrlKey: true, key: 'g', code: 'KeyG' }))).toBe('Ctrl+G');
  });
  it('định dạng số theo phím Ctrl+Shift+~ $ % ^ # @ !', () => {
    expect(formatValue(1234.5, NF_SHORTCUT.number).text).toBe('1.234,50');
    expect(formatValue(-1234.5, NF_SHORTCUT.currency).text).toBe('(1.234,50 ₫)');
    expect(formatValue(0.256, NF_SHORTCUT.percent).text).toBe('26%');
    expect(formatValue(12345, NF_SHORTCUT.sci).text).toBe('1,23E+04');
    expect(formatValue(45000.75, NF_SHORTCUT.time).text).toBe('6:00 CH');
    expect(formatValue(45000, NF_SHORTCUT.date).text).toMatch(/^15-/);
    expect(NF_SHORTCUT.general).toBeUndefined();
  });
});

describe('0.6 – gộp thao tác thành một bước hoàn tác', () => {
  it('group(): nhiều lệnh → một lần Ctrl+Z', () => {
    const { wb, s } = book([[1, 2]]);
    wb.group(() => { wb.setInput(s, 0, 0, '9'); wb.setStyle(s, rg(0, 0, 0, 0), { b: true }); wb.setInput(s, 0, 1, '8'); });
    expect(val(wb, s, 'A1')).toBe(9);
    wb.undo();
    expect(val(wb, s, 'A1')).toBe(1); expect(val(wb, s, 'B1')).toBe(2); expect(wb.style(s, 0, 0).b).toBeUndefined();
    wb.redo();
    expect(val(wb, s, 'B1')).toBe(8); expect(wb.style(s, 0, 0).b).toBe(true);
  });
  it('writeCells: ghi nhiều ô rời trong một bước', () => {
    const { wb, s } = book([[1, 2, 3]]);
    wb.writeCells(s, [[0, 0, 'x'], [5, 5, '=A1&"y"']]);
    expect(val(wb, s, 'F6')).toBe('xy');
    wb.undo(); expect(val(wb, s, 'A1')).toBe(1); expect(val(wb, s, 'F6')).toBe(null);
  });
});

describe('0.6 – Tìm & Thay thế (Ctrl+F, Ctrl+H)', () => {
  const data = [['Đèn LED 100W', 'đèn led', 5], ['Cột thép', '=A1&" mới"', '=C1*2'], ['ĐÈN', 'x', 'đèn*']];
  it('tìm không phân biệt hoa thường, theo hàng; ký tự đại diện * ?', () => {
    const { wb, s } = book(data);
    const h = findAll(wb, s, { what: 'đèn', within: 'sheet', by: 'rows', lookIn: 'values', matchCase: false, whole: false });
    expect(h.map((x) => [x.r, x.c])).toEqual([[0, 0], [0, 1], [1, 1], [2, 0], [2, 2]]);
    const byCol = findAll(wb, s, { what: 'đèn', within: 'sheet', by: 'cols', lookIn: 'values', matchCase: false, whole: false });
    expect(byCol.map((x) => [x.r, x.c])).toEqual([[0, 0], [2, 0], [0, 1], [1, 1], [2, 2]]);
    expect(findAll(wb, s, { what: 'Đèn', within: 'sheet', by: 'rows', lookIn: 'values', matchCase: true, whole: false }).length).toBe(2);
    expect(findAll(wb, s, { what: 'đèn', within: 'sheet', by: 'rows', lookIn: 'values', matchCase: false, whole: true }).map((x) => [x.r, x.c])).toEqual([[2, 0]]);
    expect(findAll(wb, s, { what: 'đ?n l*', within: 'sheet', by: 'rows', lookIn: 'values', matchCase: false, whole: true }).length).toBe(3);
    expect(findAll(wb, s, { what: 'đèn~*', within: 'sheet', by: 'rows', lookIn: 'values', matchCase: false, whole: true }).map((x) => [x.r, x.c])).toEqual([[2, 2]]);
  });
  it('tìm trong công thức vs giá trị; trong vùng chọn; cả sổ', () => {
    const { wb, s } = book(data);
    expect(findAll(wb, s, { what: 'C1', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false }).map((x) => [x.r, x.c])).toEqual([[1, 2]]);
    expect(findAll(wb, s, { what: 'C1', within: 'sheet', by: 'rows', lookIn: 'values', matchCase: false, whole: false }).length).toBe(0);
    expect(findAll(wb, s, { what: '10', within: 'sheet', by: 'rows', lookIn: 'values', matchCase: false, whole: false }, [rg(1, 0, 2, 2)]).map((x) => [x.r, x.c])).toEqual([[1, 1], [1, 2]]);
    const s2 = wb.addSheet('Tờ 2'); wb.setInput(s2, 3, 3, 'đèn đường');
    const all = findAll(wb, s, { what: 'đèn đ', within: 'book', by: 'rows', lookIn: 'values', matchCase: false, whole: false });
    expect(all.map((x) => [x.s, x.r, x.c])).toEqual([[s2, 3, 3]]);
  });
  it('tìm tiếp / ngược lại, quay vòng', () => {
    const { wb, s } = book(data);
    const o = { what: 'đèn', within: 'sheet' as const, by: 'rows' as const, lookIn: 'values' as const, matchCase: false, whole: false };
    expect(findNext(wb, s, { s, r: 0, c: 0 }, o, 1)).toMatchObject({ r: 0, c: 1 });
    expect(findNext(wb, s, { s, r: 2, c: 2 }, o, 1)).toMatchObject({ r: 0, c: 0 });
    expect(findNext(wb, s, { s, r: 0, c: 0 }, o, -1)).toMatchObject({ r: 2, c: 2 });
    expect(findNext(wb, s, { s, r: 0, c: 0 }, { ...o, what: 'không có' }, 1)).toBeNull();
  });
  it('thay thế tất cả: một bước hoàn tác, giữ công thức, chữ hoa thường', () => {
    const { wb, s } = book(data);
    const n = replaceAll(wb, s, { what: 'đèn', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false }, 'Bóng');
    expect(n).toBe(4);
    expect(val(wb, s, 'A1')).toBe('Bóng LED 100W');
    expect(val(wb, s, 'B1')).toBe('Bóng led');
    expect(val(wb, s, 'B2')).toBe('Bóng LED 100W mới');
    expect(val(wb, s, 'C3')).toBe('Bóng*');
    wb.undo();
    expect(val(wb, s, 'A1')).toBe('Đèn LED 100W'); expect(val(wb, s, 'C3')).toBe('đèn*');
    // thay trong công thức
    expect(replaceAll(wb, s, { what: 'C1', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false }, 'C1+1')).toBe(1);
    expect(raw(wb, s, 'C2')).toBe('=C1+1*2');
    expect(val(wb, s, 'C2')).toBe(7);
    // không đổi được thành công thức lỗi cú pháp → bỏ qua ô đó, không làm hỏng
    expect(replaceAll(wb, s, { what: '~*2', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false }, '*(')).toBe(0);
  });
});

describe('0.6 – Đi tới Đặc biệt và các phím chọn', () => {
  const data = [['A', 'B', 'C'], [1, '', '=A2*2'], ['x', 3, '=1/0'], [4, '', '=A2+B3']];
  it('ô trống / hằng số / công thức (lọc theo loại)', () => {
    const { wb, s } = book(data);
    expect(goSpecial(wb, s, [rg(0, 0, 3, 2)], { kind: 'blanks' })).toEqual([rg(1, 1, 1, 1), rg(3, 1, 3, 1)]);
    expect(goSpecial(wb, s, [rg(0, 0, 0, 0)], { kind: 'blanks' })).toEqual([rg(1, 1, 1, 1), rg(3, 1, 3, 1)]); // một ô → cả vùng đã dùng
    expect(goSpecial(wb, s, [rg(0, 0, 3, 2)], { kind: 'constants', types: { num: true } })).toEqual([rg(1, 0, 1, 0), rg(2, 1, 2, 1), rg(3, 0, 3, 0)]);
    expect(goSpecial(wb, s, [rg(0, 0, 3, 2)], { kind: 'constants', types: { text: true } })).toEqual([rg(0, 0, 0, 2), rg(2, 0, 2, 0)]);
    expect(goSpecial(wb, s, [rg(0, 0, 3, 2)], { kind: 'formulas', types: { err: true } })).toEqual([rg(2, 2, 2, 2)]);
    expect(goSpecial(wb, s, [rg(0, 0, 3, 2)], { kind: 'formulas' })).toEqual([rg(1, 2, 3, 2)]);
  });
  it('ô cuối, vùng hiện tại, chỉ ô đang hiện', () => {
    const { wb, s } = book(data);
    wb.setStyle(s, rg(6, 4, 6, 4), { bg: '#ff0' });
    expect(wb.lastCell(s)).toEqual({ r: 6, c: 4 });
    expect(goSpecial(wb, s, [rg(1, 1, 1, 1)], { kind: 'region' })).toEqual([rg(0, 0, 3, 2)]);
    wb.hideRows(s, 1, 1, true); wb.hideCols(s, 1, 1, true);
    expect(goSpecial(wb, s, [rg(0, 0, 3, 2)], { kind: 'visible' })).toEqual([rg(0, 0, 0, 0), rg(0, 2, 0, 2), rg(2, 0, 3, 0), rg(2, 2, 3, 2)]);
  });
  it('khác biệt trong hàng / cột', () => {
    const { wb, s } = book([[1, 1, 2], [1, 3, 1], [2, 2, 2]]);
    // so với cột A (ô hiện hành ở cột A)
    expect(goSpecial(wb, s, [rg(0, 0, 2, 2)], { kind: 'rowDiff', active: { r: 0, c: 0 } })).toEqual([rg(0, 2, 0, 2), rg(1, 1, 1, 1)]);
    // so với hàng 1
    expect(goSpecial(wb, s, [rg(0, 0, 2, 2)], { kind: 'colDiff', active: { r: 0, c: 0 } })).toEqual([rg(1, 1, 1, 2), rg(2, 0, 2, 1)]);
  });
  it('tiền lệ / phụ thuộc trực tiếp và mọi cấp', () => {
    const { wb, s } = book([[1, '=A1*2', '=B1+A1', '=C1']]);
    expect(goSpecial(wb, s, [rg(0, 2, 0, 2)], { kind: 'precedents' })).toEqual([rg(0, 0, 0, 1)]);
    expect(goSpecial(wb, s, [rg(0, 3, 0, 3)], { kind: 'precedents' })).toEqual([rg(0, 2, 0, 2)]);
    expect(goSpecial(wb, s, [rg(0, 3, 0, 3)], { kind: 'precedents', all: true })).toEqual([rg(0, 0, 0, 2)]);
    expect(goSpecial(wb, s, [rg(0, 0, 0, 0)], { kind: 'dependents' })).toEqual([rg(0, 1, 0, 2)]);
    expect(goSpecial(wb, s, [rg(0, 1, 0, 1)], { kind: 'dependents', all: true })).toEqual([rg(0, 2, 0, 3)]);
  });
  it('định dạng có điều kiện, kiểm tra dữ liệu; không có ô nào → []', () => {
    const { wb, s } = book(data);
    wb.addDV(s, { range: rg(5, 0, 6, 0), type: 'list', list: ['a', 'b'] } as any);
    expect(goSpecial(wb, s, [rg(0, 0, 0, 0)], { kind: 'dv' })).toEqual([rg(5, 0, 6, 0)]);
    expect(goSpecial(wb, s, [rg(0, 0, 0, 0)], { kind: 'cf' })).toEqual([]);
    expect(goSpecial(wb, s, [rg(0, 0, 0, 0)], { kind: 'notes' })).toEqual([]);
  });
  it('toRuns gộp ô rời thành khối chữ nhật', () => {
    expect(toRuns([[0, 0], [1, 0], [2, 0], [0, 1], [1, 1], [2, 1], [5, 5]])).toEqual([rg(0, 0, 2, 1), rg(5, 5, 5, 5)]);
  });
});

describe('0.6 – Dán Đặc biệt (Ctrl+Alt+V)', () => {
  const setup = () => {
    const { wb, s } = book([[1, '=A1*10', 'chữ', ''], [2, '=A2*10', '', 5]]);
    wb.setStyle(s, rg(0, 0, 1, 3), { b: true, fmt: '#,##0.00', bd: { t: { s: 'thin' } } });
    wb.setColWidth(s, 1, 150);
    wb.setInput(s, 5, 0, '100'); wb.setInput(s, 5, 1, '=E1'); wb.setInput(s, 6, 0, '7');
    return { wb, s };
  };
  it('Giá trị (V): ghi giá trị, không định dạng', () => {
    const { wb, s } = setup();
    wb.copy(s, rg(0, 0, 1, 1));
    wb.pasteSpecial(s, rg(10, 0, 10, 0), { what: 'values' });
    expect(raw(wb, s, 'B11')).toBe('10'); expect(raw(wb, s, 'B12')).toBe('20');
    expect(wb.style(s, 10, 1).b).toBeUndefined();
  });
  it('Công thức (F): dời tham chiếu; Định dạng (T); Độ rộng cột (W); Tất cả trừ viền (X)', () => {
    const { wb, s } = setup();
    wb.copy(s, rg(0, 0, 1, 1));
    wb.pasteSpecial(s, rg(10, 2, 10, 2), { what: 'formulas' });
    expect(raw(wb, s, 'D11')).toBe('=C11*10');
    expect(wb.style(s, 10, 3).b).toBeUndefined();
    wb.pasteSpecial(s, rg(20, 0, 20, 0), { what: 'formats' });
    expect(wb.style(s, 20, 1)).toMatchObject({ b: true, fmt: '#,##0.00' });
    expect(wb.value(s, 20, 1)).toBe(null);
    wb.pasteSpecial(s, rg(30, 5, 30, 5), { what: 'colWidths' });
    expect(wb.meta(s).colW.get(6)).toBe(150);
    wb.pasteSpecial(s, rg(40, 0, 40, 0), { what: 'noBorders' });
    expect(wb.style(s, 40, 0).bd).toBeUndefined(); expect(wb.style(s, 40, 0).b).toBe(true);
    expect(raw(wb, s, 'B41')).toBe('=A41*10');
  });
  it('Công thức & định dạng số (R), Giá trị & định dạng số (U)', () => {
    const { wb, s } = setup();
    wb.copy(s, rg(0, 0, 0, 1));
    wb.pasteSpecial(s, rg(10, 0, 10, 0), { what: 'formulasNum' });
    expect(wb.style(s, 10, 1)).toEqual({ fmt: '#,##0.00' }); expect(raw(wb, s, 'B11')).toBe('=A11*10');
    wb.pasteSpecial(s, rg(12, 0, 12, 0), { what: 'valuesNum' });
    expect(wb.style(s, 12, 1)).toEqual({ fmt: '#,##0.00' }); expect(raw(wb, s, 'B13')).toBe('10');
  });
  it('Phép tính cộng/nhân với giá trị và công thức đích; Bỏ qua ô trống', () => {
    const { wb, s } = setup();
    wb.copy(s, rg(0, 0, 1, 0)); // 1, 2
    wb.pasteSpecial(s, rg(5, 0, 5, 0), { what: 'values', op: 'add' }); // 100+1, 7+2
    expect(val(wb, s, 'A6')).toBe(101); expect(val(wb, s, 'A7')).toBe(9);
    wb.copy(s, rg(0, 0, 0, 0)); // 1
    wb.setInput(s, 0, 4, '3');
    wb.pasteSpecial(s, rg(5, 1, 5, 1), { what: 'all', op: 'mul' }); // =E1 * 1
    expect(raw(wb, s, 'B6')).toBe('=(E1)*1'); expect(val(wb, s, 'B6')).toBe(3);
    wb.copy(s, rg(0, 2, 1, 3)); // 'chữ','' / '',5
    wb.setInput(s, 8, 2, 'giữ'); wb.setInput(s, 8, 3, 'giữ2');
    wb.pasteSpecial(s, rg(8, 2, 8, 2), { what: 'all', skipBlanks: true });
    expect(val(wb, s, 'C9')).toBe('chữ'); expect(val(wb, s, 'D9')).toBe('giữ2'); expect(val(wb, s, 'D10')).toBe(5);
  });
  it('Chuyển vị (E), Dán liên kết (L), lặp nguồn cho vùng đích lớn hơn, một bước hoàn tác', () => {
    const { wb, s } = setup();
    wb.copy(s, rg(0, 0, 1, 1));
    wb.pasteSpecial(s, rg(10, 0, 10, 0), { what: 'all', transpose: true });
    expect(val(wb, s, 'A11')).toBe(1); expect(val(wb, s, 'B11')).toBe(2);
    expect(raw(wb, s, 'A12')).toBe('=A11*10'); expect(raw(wb, s, 'B12')).toBe('=B11*10');
    wb.undo();
    expect(val(wb, s, 'A11')).toBe(null); expect(wb.style(s, 10, 0).b).toBeUndefined();
    wb.pasteSpecial(s, rg(14, 0, 14, 0), { link: true });
    expect(raw(wb, s, 'A15')).toBe('=A1'); expect(raw(wb, s, 'B16')).toBe('=B2'); expect(val(wb, s, 'B16')).toBe(20);
    wb.copy(s, rg(0, 0, 0, 0));
    wb.pasteSpecial(s, rg(20, 0, 23, 1), { what: 'values' });
    expect(val(wb, s, 'B24')).toBe(1);
    const s2 = wb.addSheet('Tờ 2');
    wb.copy(s, rg(0, 0, 0, 0));
    wb.pasteSpecial(s2, rg(0, 0, 0, 0), { link: true });
    expect(wb.rawText(s2, 0, 0)).toMatch(/^=.*!\$A\$1$/);
    expect(wb.value(s2, 0, 0)).toBe(1);
  });
  it('Kiểm tra dữ liệu (N) và lỗi khi nguồn là Cắt', () => {
    const { wb, s } = setup();
    wb.addDV(s, { range: rg(0, 0, 1, 0), type: 'list', list: ['1', '2'] } as any);
    wb.copy(s, rg(0, 0, 1, 0));
    wb.pasteSpecial(s, rg(10, 3, 10, 3), { what: 'validation' });
    expect(wb.dvAt(s, 11, 3)?.type).toBe('list'); expect(wb.value(s, 10, 3)).toBe(null);
    wb.copy(s, rg(0, 0, 0, 0), true);
    expect(() => wb.pasteSpecial(s, rg(12, 0, 12, 0), { what: 'values' })).toThrow();
  });
});

describe('0.6 – F4 đổi tham chiếu khi sửa công thức', () => {
  it('A1 → $A$1 → A$1 → $A1 → A1, cả vùng và tên sheet', () => {
    let r = cycleRef('=A1+B2', 3, 3); // con trỏ sau A1
    expect(r.text).toBe('=$A$1+B2'); expect(r.end).toBe(5);
    r = cycleRef(r.text, r.start, r.end); expect(r.text).toBe('=A$1+B2');
    r = cycleRef(r.text, r.start, r.end); expect(r.text).toBe('=$A1+B2');
    r = cycleRef(r.text, r.start, r.end); expect(r.text).toBe('=A1+B2');
    expect(cycleRef('=SUM(A1:B5)', 10, 10).text).toBe('=SUM($A$1:$B$5)');
    expect(cycleRef("='Tờ 2'!C3*2", 10, 10).text).toBe("='Tờ 2'!$C$3*2");
    expect(cycleRef('=SUM(1;2)', 9, 9).text).toBe('=SUM(1;2)'); // không có tham chiếu
    // bôi đen nhiều tham chiếu: đổi tất cả
    expect(cycleRef('=A1+B2', 1, 6).text).toBe('=$A$1+$B$2');
  });
});

describe('0.6 – Điền Nhanh (Ctrl+E)', () => {
  it('tách tên, ghép, đổi hoa thường, chữ cái đầu, lấy số', () => {
    const src = [['Nguyễn Văn An'], ['Trần Thị Bình'], ['Lê Cường']];
    expect(flashFill(src, ['An'])).toEqual(['An', 'Bình', 'Cường']);
    expect(flashFill(src, ['Nguyễn'])).toEqual(['Nguyễn', 'Trần', 'Lê']);
    expect(flashFill(src, ['NGUYỄN VĂN AN'])).toEqual(['NGUYỄN VĂN AN', 'TRẦN THỊ BÌNH', 'LÊ CƯỜNG']);
    expect(flashFill([['an', 'nguyễn'], ['bình', 'trần']], ['Nguyễn, An'])).toEqual(['Nguyễn, An', 'Trần, Bình']);
    expect(flashFill([['Mã ĐL-0123 tuyến 5'], ['Mã ĐL-77 tuyến 12']], ['0123'])).toEqual(['0123', '77']);
    expect(flashFill([['Nguyễn Văn An'], ['Trần Thị Bình'], ['Lê Văn Cường']], ['N.V.A'])).toEqual(['N.V.A', 'T.T.B', 'L.V.C']);
    expect(flashFill([['Nguyễn Văn An'], ['Trần']], ['Văn'])).toEqual(['Văn', null]);
    expect(flashFill([['a b'], ['c d']], ['b', 'c'])).toBeNull(); // mâu thuẫn ví dụ
  });
});

describe('0.6 – Điền Chuỗi (Trang đầu ▸ Điền ▸ Chuỗi…)', () => {
  it('tuyến tính, cấp số nhân, ngày theo tháng/ngày làm việc, giá trị dừng', () => {
    expect(seriesValues(1, 5, { type: 'linear', step: 2 })).toEqual([1, 3, 5, 7, 9]);
    expect(seriesValues(1, 10, { type: 'linear', step: 3, stop: 8 })).toEqual([1, 4, 7]);
    expect(seriesValues(2, 4, { type: 'growth', step: 3 })).toEqual([2, 6, 18, 54]);
    // 31/01/2024 + 1 tháng = 29/02/2024 (năm nhuận), 31/03/2024
    const d = 45322; // 31/01/2024
    expect(seriesValues(d, 3, { type: 'date', unit: 'month', step: 1 })).toEqual([45322, 45351, 45382]);
    // Thứ Sáu 02/02/2024 → Thứ Hai, Thứ Ba
    expect(seriesValues(45324, 3, { type: 'date', unit: 'weekday', step: 1 })).toEqual([45324, 45327, 45328]);
    expect(seriesValues(45322, 2, { type: 'date', unit: 'year', step: 1 })).toEqual([45322, 45688]);
  });
  it('fillSeries ghi vào vùng theo hàng/cột, một bước hoàn tác', () => {
    const { wb, s } = book([[10]]);
    wb.fillSeries(s, rg(0, 0, 4, 0), { dir: 'cols', type: 'linear', step: 5 });
    expect([0, 1, 2, 3, 4].map((r) => wb.value(s, r, 0))).toEqual([10, 15, 20, 25, 30]);
    wb.undo(); expect(wb.value(s, 4, 0)).toBe(null);
    wb.fillSeries(s, rg(0, 0, 0, 3), { dir: 'rows', type: 'growth', step: 2 });
    expect([0, 1, 2, 3].map((c) => wb.value(s, 0, c))).toEqual([10, 20, 40, 80]);
  });
});

describe('0.6 – kiểu ô mở rộng (hộp Định dạng Ô) và liên kết: khứ hồi .xlsx', () => {
  it('gạch chân kép, chỉ số trên, xoay chữ, co vừa ô, căn ngang mở rộng, thụt lề, khóa/ẩn, viền chéo, liên kết', async () => {
    const { wb, s } = book([['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i']]);
    wb.setStyle(s, rg(0, 0, 0, 0), { u: true, ud: true });
    wb.setStyle(s, rg(0, 1, 0, 1), { vp: 'sup' });
    wb.setStyle(s, rg(0, 2, 0, 2), { rot: 45 });
    wb.setStyle(s, rg(0, 3, 0, 3), { shk: true });
    wb.setStyle(s, rg(0, 4, 0, 4), { hx: 'cas' });
    wb.setStyle(s, rg(0, 5, 0, 5), { al: 'l', ind: 2 });
    wb.setStyle(s, rg(0, 6, 0, 6), { lk: false, hid: true });
    wb.setStyle(s, rg(0, 7, 0, 7), { bd: { d: { s: 'thin', c: '#ff0000' }, u: { s: 'thin', c: '#ff0000' } } });
    wb.setStyle(s, rg(0, 8, 0, 8), { ln: 'https://sapulico.vn', fg: '#0563c1', u: true });
    wb.setStyle(s, rg(1, 0, 1, 0), { rot: 255 });
    const back = (await readXlsx(await writeXlsx(wb))).wb;
    const b = back.sheetIds()[0];
    expect(back.style(b, 0, 0)).toMatchObject({ u: true, ud: true });
    expect(back.style(b, 0, 1).vp).toBe('sup');
    expect(back.style(b, 0, 2).rot).toBe(45);
    expect(back.style(b, 0, 3).shk).toBe(true);
    expect(back.style(b, 0, 4).hx).toBe('cas');
    expect(back.style(b, 0, 5)).toMatchObject({ al: 'l', ind: 2 });
    expect(back.style(b, 0, 6)).toMatchObject({ lk: false, hid: true });
    expect(back.style(b, 0, 7).bd?.d).toMatchObject({ s: 'thin', c: '#ff0000' });
    expect(back.style(b, 0, 7).bd?.u).toMatchObject({ s: 'thin' });
    expect(back.style(b, 0, 8).ln).toBe('https://sapulico.vn');
    expect(back.value(b, 0, 8)).toBe('i');
    expect(back.style(b, 1, 0).rot).toBe(255);
  });
});

describe('0.6 – sửa sau review độc lập', () => {
  it('lưu .xlsx: ô số/ô trống có liên kết không bị hỏng; liên kết trong sổ khứ hồi; chặn javascript:', async () => {
    const { wb, s } = book([[123, '', 'đi tới']]);
    wb.setStyle(s, rg(0, 0, 0, 1), { ln: 'https://a.vn' });
    wb.setStyle(s, rg(0, 2, 0, 2), { ln: "#'Sheet1'!B2" });
    const back = (await readXlsx(await writeXlsx(wb))).wb, b = back.sheetIds()[0];
    expect(back.value(b, 0, 0)).toBe(123);
    expect(back.value(b, 0, 1)).toBe(null);
    expect(back.style(b, 0, 2).ln).toMatch(/^#.*B2$/);
    const { safeLink } = await import('../src/core/xlsx');
    expect(safeLink('javascript:alert(1)')).toBe(false);
    expect(safeLink('data:text/html,x')).toBe(false);
    expect(safeLink('https://sapulico.vn')).toBe(true);
    expect(safeLink('#Sheet1!A1')).toBe(true);
  });
  it('bỏ viền ngoài xoá cả cạnh chung của ô kề', () => {
    const { wb, s } = book([]);
    wb.setBorder(s, rg(0, 0, 3, 3), 'all');
    wb.setBorders(s, rg(1, 1, 2, 2), { t: null, b: null, l: null, r: null, h: null, v: null });
    expect(wb.style(s, 0, 1).bd?.b).toBeUndefined();
    expect(wb.style(s, 1, 0).bd?.r).toBeUndefined();
    expect(wb.style(s, 3, 2).bd?.t).toBeUndefined();
    expect(wb.style(s, 0, 1).bd?.t).toBeDefined();
  });
  it('Chuỗi ngày giữ định dạng ngày; Dán phép tính trống ⊕ trống = trống', () => {
    const { wb, s } = book([['31/01/2024']]);
    wb.fillSeries(s, rg(0, 0, 2, 0), { dir: 'cols', type: 'date', unit: 'month', step: 1 });
    expect(wb.display(s, 2, 0).text).toBe('31/03/2024');
    wb.copy(s, rg(5, 5, 5, 5));
    wb.pasteSpecial(s, rg(9, 9, 9, 9), { what: 'values', op: 'add' });
    expect(wb.value(s, 9, 9)).toBe(null);
  });
  it('Điền Nhanh không treo với ví dụ mâu thuẫn; 20.000 hàng nhanh', () => {
    let t = Date.now();
    expect(flashFill([['a b c d e f g h'], ['a b c d e f g h']], ['a b c d e f g h-X', 'a b c d e f g h-Y'])).toBeNull();
    expect(Date.now() - t).toBeLessThan(3000);
    t = Date.now();
    const out = flashFill(Array.from({ length: 20000 }, (_, i) => ['Nguyễn Văn ' + i]), ['Nguyễn']);
    expect(out?.[19999]).toBe('Nguyễn');
    expect(Date.now() - t).toBeLessThan(3000);
  });
});
