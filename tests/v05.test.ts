// Bản 0.5: thao tác sheet/hàng/cột theo menu chuột phải của Excel 365 (màu thẻ, ẩn/hiện, di chuyển/sao chép sheet,
// chèn/xóa ô có dồn) và khứ hồi .xlsx các thuộc tính đó.
import { Workbook, DEFAULT_COL_W, DEFAULT_ROW_H } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

const rg = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
function book() {
  const wb = Workbook.create();
  const s = wb.sheetIds()[0];
  [['Hạng mục', 'SL', 'Đơn giá', 'Thành tiền'], ['Đèn', '2', '100', '=B2*C2'], ['Cột', '3', '200', '=B3*C3'], ['Cộng', '', '', '=SUM(D2:D3)']]
    .forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setInput(s, r, c, v); }));
  return { wb, s };
}
const v = (wb: Workbook, s: number, a: string) => { const m = /^([A-Z]+)(\d+)$/.exec(a)!; return wb.value(s, +m[2] - 1, m[1].charCodeAt(0) - 65); };

describe('0.5 – kích thước mặc định như Excel', () => {
  it('cột 64 px (8,43 ký tự), hàng 20 px (15 pt)', () => {
    expect(DEFAULT_COL_W).toBe(64);
    expect(DEFAULT_ROW_H).toBe(20);
  });
});

describe('0.5 – màu thẻ và ẩn sheet', () => {
  it('đặt màu thẻ, hoàn tác/làm lại, giữ khi lưu JSON', () => {
    const { wb, s } = book();
    wb.setTabColor(s, '#c00000');
    expect(wb.meta(s).tabColor).toBe('#c00000');
    wb.undo(); expect(wb.meta(s).tabColor).toBeUndefined();
    wb.redo(); expect(wb.meta(s).tabColor).toBe('#c00000');
    const w2 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(w2.meta(w2.sheetIds()[0]).tabColor).toBe('#c00000');
  });
  it('ẩn/bỏ ẩn sheet; không ẩn được sheet hiện cuối cùng; công thức vẫn tính', () => {
    const { wb, s } = book();
    const s2 = wb.addSheet('Tổng');
    wb.setInput(s2, 0, 0, `='${wb.sheetName(s)}'!D4`);
    wb.setSheetHidden(s, true);
    expect(wb.isSheetHidden(s)).toBe(true);
    expect(wb.visibleSheetIds()).toEqual([s2]);
    expect(wb.sheetIds()).toEqual([s, s2]);
    expect(wb.value(s2, 0, 0)).toBe(800);
    expect(() => wb.setSheetHidden(s2, true)).toThrow(/ít nhất một/);
    wb.undo(); expect(wb.isSheetHidden(s)).toBe(false);
    wb.redo();
    const w2 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(w2.isSheetHidden(w2.sheetIds()[0])).toBe(true);
  });
});

describe('0.5 – ẩn/bỏ ẩn hàng, cột (giữ kích thước gốc)', () => {
  it('ẩn cột rồi bỏ ẩn trả lại độ rộng cũ, hoàn tác được', () => {
    const { wb, s } = book();
    wb.setColWidth(s, 1, 120);
    wb.hideCols(s, 1, 2, true);
    expect(wb.isColHidden(s, 1)).toBe(true);
    expect(wb.isColHidden(s, 2)).toBe(true);
    expect(wb.meta(s).colW.get(1)).toBe(0);
    wb.hideCols(s, 0, 3, false);
    expect(wb.meta(s).colW.get(1)).toBe(120);
    expect(wb.meta(s).colW.get(2)).toBe(DEFAULT_COL_W);
    wb.undo(); expect(wb.isColHidden(s, 1)).toBe(true);
    wb.undo(); expect(wb.isColHidden(s, 1)).toBe(false);
  });
  it('ẩn hàng thủ công không bị bộ lọc xóa mất khi hàng nằm ngoài vùng lọc', () => {
    const { wb, s } = book();
    wb.setInput(s, 8, 0, 'ghi chú');
    wb.hideRows(s, 8, 8, true);
    wb.setFilter(s, rg(0, 0, 3, 3));
    wb.filterColumn(s, 0, { values: ['Đèn'] });
    expect(wb.isRowHidden(s, 2)).toBe(true);
    expect(wb.isRowHidden(s, 8)).toBe(true);
    wb.hideRows(s, 7, 9, false);
    expect(wb.isRowHidden(s, 8)).toBe(false);
  });
});

describe('0.5 – di chuyển / sao chép sheet', () => {
  it('di chuyển đổi thứ tự, công thức giữ nguyên, hoàn tác', () => {
    const { wb, s } = book();
    const a = wb.addSheet('A'), b = wb.addSheet('B');
    wb.setInput(b, 0, 0, `='${wb.sheetName(s)}'!D2*2`);
    wb.moveSheet(b, 0);
    expect(wb.sheetIds()).toEqual([b, s, a]);
    expect(wb.value(b, 0, 0)).toBe(400);
    wb.undo(); expect(wb.sheetIds()).toEqual([s, a, b]);
    wb.redo(); expect(wb.sheetIds()).toEqual([b, s, a]);
    const w2 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(w2.sheetIds().map((i) => w2.sheetName(i))).toEqual(['B', wb.sheetName(s), 'A']);
  });
  it('sao chép: giá trị, công thức, định dạng, gộp ô, độ rộng; tên “(2)”; hoàn tác xóa bản sao', () => {
    const { wb, s } = book();
    wb.setStyle(s, rg(0, 0, 0, 3), { b: true });
    wb.merge(s, rg(5, 0, 5, 2));
    wb.setColWidth(s, 0, 150);
    wb.setTabColor(s, '#00b050');
    const c = wb.copySheet(s, 1);
    expect(wb.sheetName(c)).toBe(wb.sheetName(s) + ' (2)');
    expect(wb.sheetIds()).toEqual([s, c]);
    expect(v(wb, c, 'D4')).toBe(800);
    expect(wb.rawText(c, 1, 3)).toBe('=B2*C2');
    expect(wb.style(c, 0, 1).b).toBe(true);
    expect(wb.mergeAt(c, 5, 1)).toEqual(rg(5, 0, 5, 2));
    expect(wb.meta(c).colW.get(0)).toBe(150);
    expect(wb.meta(c).tabColor).toBe('#00b050');
    // sửa bản sao không ảnh hưởng bản gốc
    wb.setInput(c, 1, 1, '10');
    expect(v(wb, c, 'D4')).toBe(1600);
    expect(v(wb, s, 'D4')).toBe(800);
    wb.undo(); wb.undo();
    expect(wb.sheetIds()).toEqual([s]);
  });
});

describe('0.5 – chèn / xóa ô có dồn (hộp thoại Chèn… / Xóa…)', () => {
  it('chèn ô dồn xuống: dữ liệu và công thức tham chiếu dời theo, cột khác giữ nguyên', () => {
    const { wb, s } = book();
    wb.setInput(s, 5, 5, '=B3');
    wb.setStyle(s, rg(2, 1, 2, 1), { b: true });
    wb.insertCells(s, rg(1, 1, 1, 1), 'down');
    expect(wb.rawText(s, 1, 1)).toBe('');
    expect(wb.value(s, 2, 1)).toBe(2);
    expect(wb.value(s, 3, 1)).toBe(3);
    expect(wb.style(s, 3, 1).b).toBe(true);
    expect(wb.value(s, 1, 2)).toBe(100);
    expect(wb.rawText(s, 5, 5)).toBe('=B4');
    wb.undo();
    expect(wb.value(s, 1, 1)).toBe(2);
    expect(wb.style(s, 2, 1).b).toBe(true);
  });
  it('chèn ô dồn sang phải, xóa ô dồn lên và dồn sang trái', () => {
    const { wb, s } = book();
    wb.insertCells(s, rg(1, 0, 2, 0), 'right');
    expect(wb.rawText(s, 1, 0)).toBe('');
    expect(wb.value(s, 1, 1)).toBe('Đèn');
    expect(wb.value(s, 1, 4)).toBe(200);
    wb.deleteCells(s, rg(1, 0, 2, 0), 'left');
    expect(wb.value(s, 1, 0)).toBe('Đèn');
    wb.deleteCells(s, rg(1, 0, 1, 0), 'up');
    expect(wb.value(s, 1, 0)).toBe('Cột');
    expect(wb.value(s, 2, 0)).toBe('Cộng');
  });
  it('không cắt ngang vùng gộp ô', () => {
    const { wb, s } = book();
    wb.merge(s, rg(6, 0, 6, 2));
    expect(() => wb.insertCells(s, rg(1, 1, 1, 1), 'down')).toThrow(/gộp/);
  });
});

describe('0.5 – khứ hồi .xlsx', () => {
  it('màu thẻ, sheet ẩn, cột ẩn (giữ độ rộng), hàng ẩn, thứ tự sheet', async () => {
    const { wb, s } = book();
    const t = wb.addSheet('Phụ');
    wb.setTabColor(s, '#ffc000');
    wb.setColWidth(s, 2, 140);
    wb.hideCols(s, 2, 2, true);
    wb.hideRows(s, 2, 2, true);
    wb.setSheetHidden(t, true);
    wb.moveSheet(t, 0);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const [a, b] = w2.sheetIds();
    expect(w2.sheetName(a)).toBe('Phụ');
    expect(w2.isSheetHidden(a)).toBe(true);
    expect(w2.meta(b).tabColor?.toLowerCase()).toBe('#ffc000');
    expect(w2.isColHidden(b, 2)).toBe(true);
    expect(w2.isRowHidden(b, 2)).toBe(true);
    w2.hideCols(b, 2, 2, false);
    expect(Math.abs(w2.meta(b).colW.get(2) - 140)).toBeLessThanOrEqual(2);
  });
});

import { formatNumber } from '../src/core/numfmt';
describe('0.5 – định dạng Phân số trong danh sách định dạng số', () => {
  it('# ?/? và # ?/4 như Excel', () => {
    expect(formatNumber(1234.5, '# ?/?')).toBe('1234 1/2');
    expect(formatNumber(-3.25, '# ?/?')).toBe('-3 1/4');
    expect(formatNumber(0.75, '# ?/4')).toBe('3/4');
    expect(formatNumber(2, '# ?/?')).toBe('2');
  });
});

// ---- Rà soát 0.5 ----
describe('Rà soát 0.5 – chèn/xóa ô dồn như Excel', () => {
  const nums = () => { const wb = Workbook.create(); const s = wb.sheetIds()[0]; for (let r = 0; r < 5; r++) wb.setInput(s, r, 0, String(r + 1)); return { wb, s }; };
  it('chèn ô dồn xuống trong vùng SUM → vùng giãn ra (A1:A6)', () => {
    const { wb, s } = nums();
    wb.setInput(s, 0, 2, '=SUM(A1:A5)');
    wb.setInput(s, 0, 3, '=B3');
    wb.setInput(s, 2, 1, '7');
    wb.insertCells(s, { r1: 1, c1: 0, r2: 1, c2: 0 }, 'down');
    wb.setInput(s, 1, 0, '5');
    expect(wb.rawText(s, 0, 2)).toBe('=SUM(A1:A6)');
    expect(wb.value(s, 0, 2)).toBe(20);
    // ô ngoài dải không bị dời, tham chiếu tới nó giữ nguyên
    expect(wb.value(s, 2, 1)).toBe(7);
    expect(wb.rawText(s, 0, 3)).toBe('=B3');
    wb.undo(); wb.undo();
    expect(wb.rawText(s, 0, 2)).toBe('=SUM(A1:A5)');
    expect(wb.value(s, 1, 0)).toBe(2);
  });
  it('xóa ô dồn lên: tham chiếu tới ô bị xóa thành #REF!, ô bên cạnh giữ nguyên', () => {
    const { wb, s } = nums();
    wb.setInput(s, 0, 2, '=A2');
    wb.setInput(s, 1, 1, 'giữ');
    wb.deleteCells(s, { r1: 1, c1: 0, r2: 1, c2: 0 }, 'up');
    expect(wb.display(s, 0, 2).text).toBe('#REF!');
    expect(wb.value(s, 1, 0)).toBe(3);
    expect(wb.value(s, 1, 1)).toBe('giữ');
  });
  it('quy tắc định dạng có điều kiện trong dải dời theo', () => {
    const { wb, s } = nums();
    wb.addCF(s, { range: { r1: 3, c1: 0, r2: 3, c2: 0 }, type: 'dup', fmt: { bg: '#ff0000' } } as any);
    wb.insertCells(s, { r1: 1, c1: 0, r2: 1, c2: 0 }, 'down');
    expect(wb.meta(s).cf[0].range).toEqual({ r1: 4, c1: 0, r2: 4, c2: 0 });
  });
});

describe('Rà soát 0.5 – khác', () => {
  it('in vùng chọn cả cột chỉ lấy phần có dữ liệu', async () => {
    const { paginate, normPrint } = await import('../src/core/paginate');
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    for (let r = 0; r < 10; r++) wb.setInput(s, r, 0, String(r));
    const L = paginate(wb, s, normPrint(null), { range: { r1: 0, c1: 0, r2: 1048575, c2: 0 } });
    expect(L.pages.length).toBe(1);
    const L2 = paginate(wb, s, normPrint(null), { range: { r1: 0, c1: 0, r2: 1048575, c2: 16383 } });
    expect(L2.pages.length).toBe(1);
  });
  it('Vừa X×Y trang bỏ qua ngắt trang thủ công (như Excel)', async () => {
    const { paginate, normPrint } = await import('../src/core/paginate');
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    for (let r = 0; r < 20; r++) wb.setInput(s, r, 0, String(r));
    wb.insertPageBreak(s, 10, 0);
    const L = paginate(wb, s, normPrint({ ...normPrint(null), fit: { w: 1, h: 1 } }));
    expect(L.pages.length).toBe(1);
    expect(L.scale).toBe(1);
  });
  it('xóa sheet rồi thêm sheet: sheet mới ở cuối (mã sheet dùng lại)', () => {
    const wb = Workbook.create();
    const a = wb.sheetIds()[0], b = wb.addSheet('B'), c = wb.addSheet('C');
    wb.moveSheet(c, 0);
    wb.removeSheet(c);
    const d = wb.addSheet('D');
    expect(wb.sheetIds()).toEqual([a, b, d]);
    wb.undo(); wb.undo();
    expect(wb.sheetIds().map((i) => wb.sheetName(i))).toEqual(['C', wb.sheetName(a), 'B']);
  });
  it('thêm sheet tại vị trí là một bước hoàn tác', () => {
    const wb = Workbook.create();
    const a = wb.sheetIds()[0], b = wb.addSheet('B');
    const n = wb.addSheet('N', 1);
    expect(wb.sheetIds()).toEqual([a, n, b]);
    wb.undo();
    expect(wb.sheetIds()).toEqual([a, b]);
  });
  it('đặt độ rộng nhiều cột là một bước hoàn tác', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setSizes(s, 'c', [[0, 100], [1, 100], [2, 100]]);
    expect(wb.meta(s).colW.get(2)).toBe(100);
    wb.undo();
    expect(wb.meta(s).colW.get(0)).toBe(64);
  });
  it('sao chép sheet giữ độ rộng/chiều cao mặc định', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.meta(s).colW.def = 100;
    const c = wb.copySheet(s);
    expect(wb.meta(c).colW.def).toBe(100);
  });
});

describe('Rà soát 0.5 – đầu/chân trang & in nhiều khổ', () => {
  it('mã phông/cỡ/màu không bị biến thành chữ khi qua hộp Thiết lập Trang', async () => {
    const { rawToFriendly, friendlyToRaw } = await import('../src/ui/pagesetup');
    for (const raw of ['&"Arial,Bold"&14Tiêu đề', '&KFF0000Đỏ', 'Trang &P / &N &&', '&BĐậm&B']) expect(friendlyToRaw(rawToFriendly(raw))).toBe(raw);
    expect(rawToFriendly('Trang &P / &N')).toBe('Trang &[Trang] / &[Số trang]');
  });
  it('In cả sổ có sheet dọc + ngang: mỗi khổ một @page có tên', async () => {
    const { buildJob } = await import('../src/core/print');
    const wb = Workbook.create(); const a = wb.sheetIds()[0], b = wb.addSheet('Ngang');
    wb.setInput(a, 0, 0, 'x'); wb.setInput(b, 0, 0, 'y');
    wb.setPrint(b, { orientation: 'landscape', paper: 'A4', gridlines: false });
    const j = buildJob(wb, [{ sheet: a }, { sheet: b }]);
    expect(j.html).toMatch(/@page pg0 \{ size: 210mm 297mm/);
    expect(j.html).toMatch(/@page pg1 \{ size: 297mm 210mm/);
    expect(j.html).toContain('page:pg1');
  });
  it('hàng lặp lại khi in dời theo khi chèn hàng', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setPrint(s, { orientation: 'portrait', paper: 'A4', gridlines: false, titleRowRange: [2, 3] });
    wb.insertRows(s, 0, 2);
    expect(wb.meta(s).print!.titleRowRange).toEqual([4, 5]);
  });
});
