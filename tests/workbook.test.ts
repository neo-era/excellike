import { Workbook } from '../src/core/workbook';

const R = (a: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a)!;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return { r: +m[2] - 1, c: c - 1 };
};

function wbWith(cells: Record<string, string>) {
  const wb = Workbook.create();
  const s = wb.sheetIds()[0];
  for (const [a, v] of Object.entries(cells)) {
    const { r, c } = R(a);
    wb.setInput(s, r, c, v);
  }
  const val = (a: string) => { const { r, c } = R(a); return wb.value(s, r, c); };
  const txt = (a: string) => { const { r, c } = R(a); return wb.display(s, r, c).text; };
  return { wb, s, val, txt };
}

describe('Công thức – đối chiếu kết quả Excel', () => {
  const close = (a: unknown, b: number) => expect(Math.abs((a as number) - b)).toBeLessThan(1e-9);

  it('toán và làm tròn', () => {
    const { val } = wbWith({
      A1: '=SUM(1;2;3)', A2: '=ROUND(2,345;2)', A3: '=ROUNDDOWN(-2,5;0)', A4: '=INT(-2,5)',
      A5: '=MOD(-7;3)', A6: '=0,1+0,2', A7: '=ROUNDUP(3,21;1)', A8: '=ABS(-4)*POWER(2;3)',
    });
    expect(val('A1')).toBe(6);
    expect(val('A2')).toBe(2.35);
    expect(val('A3')).toBe(-2);
    expect(val('A4')).toBe(-3);
    expect(val('A5')).toBe(2);
    close(val('A6'), 0.3);
    expect(val('A7')).toBe(3.3);
    expect(val('A8')).toBe(32);
  });

  it('logic', () => {
    const { val } = wbWith({
      A1: '5', B1: '=IF(A1>1;"lớn";"nhỏ")', B2: '=IFERROR(1/0;"lỗi")', B3: '=AND(TRUE;FALSE)',
      B4: '=IFS(A1<3;"a";A1<10;"b")', B5: '=NOT(OR(A1=1;A1=2))',
    });
    expect(val('B1')).toBe('lớn');
    expect(val('B2')).toBe('lỗi');
    expect(val('B3')).toBe(false);
    expect(val('B4')).toBe('b');
    expect(val('B5')).toBe(true);
  });

  it('tra cứu VLOOKUP, INDEX/MATCH, XLOOKUP', () => {
    const { val } = wbWith({
      A1: 'LED', B1: '4250000', A2: 'Cột', B2: '6800000', A3: 'Cáp', B3: '118000',
      D1: '=VLOOKUP("Cột";A1:B3;2;FALSE)', D2: '=INDEX(B1:B3;MATCH("Cáp";A1:A3;0))',
      D3: '=XLOOKUP("LED";A1:A3;B1:B3)', D4: '=XLOOKUP("Tủ";A1:A3;B1:B3;"không có")',
      D5: '=VLOOKUP("Tủ";A1:B3;2;FALSE)',
      F1: '0', G1: 'thấp', F2: '100', G2: 'vừa', F3: '1000', G3: 'cao', D6: '=VLOOKUP(450;F1:G3;2;TRUE)',
    });
    expect(val('D1')).toBe(6800000);
    expect(val('D2')).toBe(118000);
    expect(val('D3')).toBe(4250000);
    expect(val('D4')).toBe('không có');
    expect((val('D5') as any).value).toBe('#N/A');
    expect(val('D6')).toBe('vừa');
  });

  it('thống kê có điều kiện', () => {
    const { val } = wbWith({
      A1: 'Vật tư', B1: '10', A2: 'Thiết bị', B2: '20', A3: 'Vật tư', B3: '30', A4: '', B4: '',
      C1: '=COUNTIF(A1:A3;"Vật tư")', C2: '=SUMIF(A1:A3;"Vật tư";B1:B3)', C3: '=SUMIFS(B1:B3;A1:A3;"Vật tư";B1:B3;">15")',
      C4: '=AVERAGEIF(B1:B3;">10")', C5: '=COUNTA(A1:A4)', C6: '=COUNTBLANK(A1:A4)', C7: '=MEDIAN(B1:B3)',
      C8: '=COUNTIFS(A1:A3;"Vật tư";B1:B3;"<20")',
    });
    expect(val('C1')).toBe(2);
    expect(val('C2')).toBe(40);
    expect(val('C3')).toBe(30);
    expect(val('C4')).toBe(25);
    expect(val('C5')).toBe(3);
    expect(val('C6')).toBe(1);
    expect(val('C7')).toBe(20);
    expect(val('C8')).toBe(1);
  });

  it('văn bản', () => {
    const { val } = wbWith({
      A1: 'Chiếu sáng', B1: '=LEFT(A1;5)', B2: '=LEN(A1)', B3: '=UPPER(A1)', B4: '=TRIM("  a   b ")',
      B5: '=A1&" "&"KVTT"', B6: '=TEXTJOIN(", ";TRUE;"a";"";"b")', B7: '=SUBSTITUTE("a-b-c";"-";"/")', B8: '=MID(A1;7;4)',
    });
    expect(val('B1')).toBe('Chiếu');
    expect(val('B2')).toBe(10);
    expect(val('B3')).toBe('CHIẾU SÁNG');
    expect(val('B4')).toBe('a b');
    expect(val('B5')).toBe('Chiếu sáng KVTT');
    expect(val('B6')).toBe('a, b');
    expect(val('B7')).toBe('a/b/c');
    expect(val('B8')).toBe('sáng');
  });

  it('ngày tháng (số sê-ri giống Excel)', () => {
    const { val, txt } = wbWith({
      A1: '=DATE(2026;2;1)', A2: '=EDATE(DATE(2026;1;31);1)', A3: '=NETWORKDAYS(DATE(2026;10;1);DATE(2026;10;31))',
      A4: '01/02/2026', A5: '=YEAR(A4)*100+MONTH(A4)', A6: '=A4+30', A7: '=WEEKDAY(DATE(2026;10;28))',
    });
    expect(val('A1')).toBe(46054);
    expect(val('A2')).toBe(46081);
    expect(val('A3')).toBe(22);
    expect(val('A4')).toBe(46054);
    expect(txt('A4')).toBe('01/02/2026');
    expect(val('A5')).toBe(202602);
    expect(val('A6')).toBe(46084);
    expect(val('A7')).toBe(4);
  });

  it('thống kê và tài chính', () => {
    const { val } = wbWith({
      A1: '2', A2: '4', A3: '4', A4: '4', A5: '5', A6: '5', A7: '7', A8: '9',
      B1: '=STDEV.S(A1:A8)', B2: '=STDEV.P(A1:A8)', B3: '=PMT(0,1/12;12;-1000)', B4: '=AVERAGE(A1:A8)', B5: '=MAX(A1:A8)-MIN(A1:A8)',
    });
    close(val('B1'), 2.1380899352993947);
    expect(val('B2')).toBe(2);
    close(val('B3'), 87.91588723000);
    expect(val('B4')).toBe(5);
    expect(val('B5')).toBe(7);
  });

  it('các lỗi chuẩn Excel', () => {
    const { txt } = wbWith({ A1: '=1/0', A2: '=FOO(1)', A3: '=1+"abc"', B1: '=B2', B2: '=B1' });
    expect(txt('A1')).toBe('#DIV/0!');
    expect(txt('A2')).toBe('#NAME?');
    expect(txt('A3')).toBe('#VALUE!');
    expect(txt('B1')).toBe('#CYCLE!');
  });

  it('nhập số kiểu Việt Nam vào ô', () => {
    const { val, txt } = wbWith({ A1: '1.250.000', A2: '1,5', A3: '15%', A4: '=A1*A2' });
    expect(val('A1')).toBe(1250000);
    expect(val('A2')).toBe(1.5);
    expect(val('A3')).toBe(0.15);
    expect(txt('A3')).toBe('15%');
    expect(val('A4')).toBe(1875000);
  });
});

describe('Tính lại theo phụ thuộc', () => {
  it('sửa ô nguồn chỉ làm thay đổi các ô phụ thuộc', () => {
    const { wb, s, val } = wbWith({ A1: '1', B1: '=A1*2', C1: '=B1+1', D1: '100', E1: '=D1' });
    const changed = wb.setInput(s, 0, 0, '5');
    expect(val('B1')).toBe(10);
    expect(val('C1')).toBe(11);
    const addrs = changed.map((x) => `${x.r},${x.c}`).sort();
    expect(addrs).toEqual(['0,0', '0,1', '0,2']);
  });
});

describe('Hoàn tác / làm lại', () => {
  it('nội dung', () => {
    const { wb, s, val } = wbWith({ A1: '1' });
    wb.setInput(s, 0, 0, '2');
    wb.setInput(s, 0, 0, '3');
    wb.undo();
    expect(val('A1')).toBe(2);
    wb.undo();
    expect(val('A1')).toBe(1);
    wb.redo();
    expect(val('A1')).toBe(2);
  });
  it('ghi cùng giá trị cũ vẫn giữ lịch sử đúng', () => {
    const { wb, s, val } = wbWith({ A1: '1' });
    wb.setInput(s, 0, 0, '1');
    wb.setInput(s, 0, 0, '9');
    wb.undo();
    wb.undo();
    expect(val('A1')).toBe(1);
  });
  it('định dạng ô', () => {
    const { wb, s } = wbWith({ A1: 'x' });
    wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: true });
    expect(wb.style(s, 0, 0).b).toBe(true);
    wb.undo();
    expect(wb.style(s, 0, 0).b).toBeFalsy();
    wb.redo();
    expect(wb.style(s, 0, 0).b).toBe(true);
  });
  it('xen kẽ định dạng và nội dung', () => {
    const { wb, s, val } = wbWith({ A1: '1' });
    wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { i: true });
    wb.setInput(s, 0, 0, '2');
    wb.undo();
    expect(val('A1')).toBe(1);
    expect(wb.style(s, 0, 0).i).toBe(true);
    wb.undo();
    expect(wb.style(s, 0, 0).i).toBeFalsy();
    expect(val('A1')).toBe(1);
  });
});

describe('Chèn / xóa hàng cột', () => {
  it('công thức tự dời tham chiếu, định dạng và cỡ hàng đi theo', () => {
    const { wb, s, val } = wbWith({ A1: '10', A2: '20', A3: '=SUM(A1:A2)' });
    wb.setStyle(s, { r1: 1, c1: 0, r2: 1, c2: 0 }, { b: true });
    wb.setRowHeight(s, 1, 40);
    wb.insertRows(s, 1, 2);
    expect(val('A5')).toBe(30);
    expect(wb.rawText(s, 4, 0)).toBe('=SUM(A1:A4)');
    expect(wb.style(s, 3, 0).b).toBe(true);
    expect(wb.meta(s).rowH.get(3)).toBe(40);
    wb.undo();
    expect(val('A3')).toBe(30);
    expect(wb.style(s, 1, 0).b).toBe(true);
    expect(wb.meta(s).rowH.get(1)).toBe(40);
  });
  it('xóa hàng được tham chiếu → #REF!, hoàn tác khôi phục', () => {
    const { wb, s, txt, val } = wbWith({ A1: '5', B3: '=A1*2' });
    wb.deleteRows(s, 0, 1);
    expect(txt('B2')).toBe('#REF!');
    wb.undo();
    expect(val('B3')).toBe(10);
  });
  it('chèn cột', () => {
    const { wb, s, val } = wbWith({ A1: '2', B1: '=A1*3' });
    wb.insertCols(s, 0, 1);
    expect(val('C1')).toBe(6);
    expect(wb.rawText(s, 0, 2)).toBe('=B1*3');
  });
});

describe('Sao chép / dán / kéo điền', () => {
  it('dán công thức dời tham chiếu tương đối, giữ tuyệt đối', () => {
    const { wb, s, val } = wbWith({ A1: '1', A2: '2', B1: '=A1*$D$1', D1: '10' });
    wb.copy(s, { r1: 0, c1: 1, r2: 0, c2: 1 });
    wb.paste(s, 1, 1);
    expect(wb.rawText(s, 1, 1)).toBe('=A2*$D$1');
    expect(val('B2')).toBe(20);
  });
  it('cắt rồi dán di chuyển ô, công thức tham chiếu đi theo', () => {
    const { wb, s, val } = wbWith({ A1: '7', B1: '=A1' });
    wb.copy(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, true);
    wb.paste(s, 4, 0);
    expect(val('A5')).toBe(7);
    expect(wb.rawText(s, 0, 1)).toBe('=A5');
  });
  it('dán dữ liệu ngoài (TSV)', () => {
    const { wb, s, val } = wbWith({});
    wb.pasteText(s, 0, 0, '1\t2\r\n3\t=A1+B1\r\n');
    expect(val('B2')).toBe(3);
  });
  it('dán định dạng đi kèm', () => {
    const { wb, s } = wbWith({ A1: 'x' });
    wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: true, bg: '#ff0' });
    wb.copy(s, { r1: 0, c1: 0, r2: 0, c2: 0 });
    wb.paste(s, 2, 2);
    expect(wb.style(s, 2, 2)).toMatchObject({ b: true, bg: '#ff0' });
  });
  it('kéo điền xuống: dãy số và công thức', () => {
    const { wb, s, val } = wbWith({ A1: '1', A2: '2', B1: '=A1*10', B2: '=A2*10' });
    wb.fill(s, { r1: 0, c1: 0, r2: 1, c2: 1 }, { r1: 0, c1: 0, r2: 4, c2: 1 });
    expect(val('A5')).toBe(5);
    expect(val('B5')).toBe(50);
    wb.undo();
    expect(val('A5')).toBeNull();
  });
  it('AutoSum lấy vùng số liền phía trên', () => {
    const { wb, s, val } = wbWith({ A1: 'Tiêu đề', A2: '4', A3: '6' });
    wb.autoSum(s, 3, 0);
    expect(wb.rawText(s, 3, 0)).toBe('=SUM(A2:A3)');
    expect(val('A4')).toBe(10);
  });
});

describe('Nhiều sheet', () => {
  it('thêm, đổi tên, tham chiếu chéo, xóa và hoàn tác', () => {
    const wb = Workbook.create();
    const s1 = wb.sheetIds()[0];
    const s2 = wb.addSheet();
    expect(wb.sheetName(s2)).toBe('Sheet2');
    wb.setInput(s2, 0, 0, '42');
    wb.setInput(s1, 0, 0, '=Sheet2!A1+1');
    expect(wb.value(s1, 0, 0)).toBe(43);
    wb.renameSheet(s2, 'Dự toán');
    expect(wb.rawText(s1, 0, 0)).toBe("='Dự toán'!A1+1");
    expect(() => wb.renameSheet(s2, 'Sheet1')).toThrow();
    expect(() => wb.renameSheet(s2, 'a/b')).toThrow();
    wb.removeSheet(s2);
    expect(wb.sheetIds()).toEqual([s1]);
    expect(wb.display(s1, 0, 0).text).toBe('#REF!');
    wb.undo();
    expect(wb.sheetIds().length).toBe(2);
    expect(wb.value(s1, 0, 0)).toBe(43);
  });
  it('không xóa được sheet cuối cùng', () => {
    const wb = Workbook.create();
    expect(() => wb.removeSheet(wb.sheetIds()[0])).toThrow();
  });
});

describe('Lưu / mở lại (JSON)', () => {
  it('khứ hồi giữ công thức, định dạng, cỡ cột, cố định ngăn, nhiều sheet', () => {
    const { wb, s } = wbWith({ A1: '01/02/2026', B1: '=A1+1', C1: '15%' });
    wb.setStyle(s, { r1: 0, c1: 1, r2: 0, c2: 1 }, { b: true, fmt: 'dd/mm/yyyy' });
    wb.setColWidth(s, 1, 150);
    wb.setFreeze(s, 1, 1);
    const s2 = wb.addSheet('Tổng hợp');
    wb.setInput(s2, 0, 0, "='Sheet1'!B1");
    const back = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    const [b1, b2] = back.sheetIds();
    expect(back.display(b1, 0, 0).text).toBe('01/02/2026');
    expect(back.display(b1, 0, 1).text).toBe('02/02/2026');
    expect(back.display(b1, 0, 2).text).toBe('15%');
    expect(back.style(b1, 0, 1).b).toBe(true);
    expect(back.meta(b1).colW.get(1)).toBe(150);
    expect(back.meta(b1).freeze).toEqual({ rows: 1, cols: 1 });
    expect(back.sheetName(b2)).toBe('Tổng hợp');
    expect(back.value(b2, 0, 0)).toBe(46055);
  });
});
