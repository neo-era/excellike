import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { lamIn, lamOut, declaredNames } from '../src/core/lambda';

function wbWith(cells: [string, string][]) {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  for (const [a, v] of cells) { const m = /^([A-Z]+)(\d+)$/.exec(a)!; const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; wb.setInput(s, +m[2] - 1, c, v); }
  return { wb, s, v: (a: string) => { const m = /^([A-Z]+)(\d+)$/.exec(a)!; const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return wb.value(s, +m[2] - 1, c); }, raw: (a: string) => { const m = /^([A-Z]+)(\d+)$/.exec(a)!; const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return wb.rawText(s, +m[2] - 1, c); } };
}
const err = (v: unknown) => (v && typeof v === 'object' ? (v as { value?: string }).value : undefined);

describe('LET (C1.1–C1.2)', () => {
  it('tên dùng được ở giá trị sau và trong phép tính', () => {
    const { v } = wbWith([['A1', '5'], ['B1', '=LET(x;A1*2;y;x+1;x*y)'], ['B2', '=LET(a;1;b;a+1;d;b+1;a+b+d)']]);
    expect(v('B1')).toBe(110);
    expect(v('B2')).toBe(6);
  });
  it('tên LET không thành tên toàn cục; tên lồng che tên ngoài', () => {
    const { wb, v } = wbWith([['A1', '=LET(x;1;LET(x;2;x)+x)']]);
    expect(v('A1')).toBe(3);
    expect(wb.listNames().length).toBe(0);
  });
  it('sai quy tắc → lỗi: số đối số chẵn, tên là tham chiếu ô', () => {
    const { v } = wbWith([['A1', '=LET(x;1)'], ['A2', '=LET(B5;1;B5)'], ['A3', '=LET(x;1;y;2)']]);
    expect(err(v('A1'))).toBe('#VALUE!');
    expect(err(v('A2'))).toBe('#VALUE!');
    expect(err(v('A3'))).toBe('#VALUE!');
  });
  it('tham chiếu trong LET dời khi chèn hàng; tính lại khi ô nguồn đổi', () => {
    const { wb, s, v, raw } = wbWith([['A1', '3'], ['B1', '=LET(x;A1;x*x)']]);
    expect(v('B1')).toBe(9);
    wb.setInput(s, 0, 0, '4');
    expect(v('B1')).toBe(16);
    wb.insertRows(s, 0, 1);
    expect(raw('B2')).toBe('=LET(x;A2;x*x)');
    expect(v('B2')).toBe(16);
  });
  it('LET trả mảng thì tràn đúng kích thước', () => {
    const { v } = wbWith([['A1', '=LET(n;3;SEQUENCE(n))'], ['C1', '1']]);
    expect([v('A1'), v('A2'), v('A3'), v('A4')]).toEqual([1, 2, 3, null]);
  });
});

describe('LAMBDA (C1.3–C1.4)', () => {
  it('gọi ngay, trong LET, đứng một mình → #CALC!', () => {
    const { v, raw } = wbWith([['A1', '=LAMBDA(x;x+1)(5)'], ['A2', '=LET(f;LAMBDA(a;a*2);f(3))'], ['A3', '=LAMBDA(x;x+1)'], ['A4', '=LET(f;LAMBDA(a;a);f)']]);
    expect(v('A1')).toBe(6);
    expect(raw('A1')).toBe('=LAMBDA(x;x+1)(5)');
    expect(v('A2')).toBe(6);
    expect(err(v('A3'))).toBe('#CALC!');
    expect(err(v('A4'))).toBe('#CALC!');
  });
  it('sai số đối số, quá 253 tham số → #VALUE!', () => {
    const many = Array.from({ length: 254 }, (_, i) => 'p' + i).join(';');
    const { v } = wbWith([['A1', '=LAMBDA(x;y;x+y)(1)'], ['A2', `=LAMBDA(${many};1)(1)`]]);
    expect(err(v('A1'))).toBe('#VALUE!');
    expect(err(v('A2'))).toBe('#VALUE!');
  });
  it('tên LAMBDA trong Trình Quản lý Tên; lồng; đệ quy; sửa định nghĩa thì tính lại', () => {
    const { wb, s, v } = wbWith([]);
    wb.defineName('NhanDoi', '=LAMBDA(x;x*2)');
    wb.defineName('TongDen', '=LAMBDA(n;IF(n<=0;0;n+TongDen(n-1)))');
    wb.setInput(s, 0, 0, '=NhanDoi(4)');
    wb.setInput(s, 1, 0, '=NhanDoi(NhanDoi(2))');
    wb.setInput(s, 2, 0, '=TongDen(10)');
    wb.setInput(s, 3, 0, '=TongDen(1000)');
    wb.setInput(s, 4, 0, '=TongDen(100000)');
    expect(v('A1')).toBe(8);
    expect(wb.rawText(s, 0, 0)).toBe('=NhanDoi(4)');
    expect(v('A2')).toBe(8);
    expect(v('A3')).toBe(55);
    expect(v('A4')).toBe(500500);
    expect(err(v('A5'))).toBe('#NUM!');
    const cur = wb.listNames().find((n) => n.name === 'NhanDoi')!;
    wb.updateName(cur.name, cur.scope, { name: cur.name, expr: '=LAMBDA(x;x*3)', scope: cur.scope });
    expect(v('A1')).toBe(12);
  });
  it('MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY', () => {
    const { v } = wbWith([
      ['A1', '1'], ['A2', '2'], ['A3', '3'], ['B1', '10'], ['B2', '20'], ['B3', '30'],
      ['D1', '=MAP(A1:A3;LAMBDA(x;x*x))'],
      ['E1', '=REDUCE(0;A1:B3;LAMBDA(a;x;a+x))'],
      ['F1', '=SCAN(0;A1:A3;LAMBDA(a;x;a+x))'],
      ['G1', '=BYROW(A1:B3;LAMBDA(r;SUM(r)))'],
      ['H1', '=BYCOL(A1:B3;LAMBDA(c;MAX(c)))'],
      ['H5', '=MAKEARRAY(2;3;LAMBDA(r;c;r*10+c))'],
      ['D6', '=MAP(A1:A3;B1:B3;LAMBDA(a;b;a+b))'],
    ]);
    expect([v('D1'), v('D2'), v('D3')]).toEqual([1, 4, 9]);
    expect(v('E1')).toBe(66);
    expect([v('F1'), v('F2'), v('F3')]).toEqual([1, 3, 6]);
    expect([v('G1'), v('G2'), v('G3')]).toEqual([11, 22, 33]);
    expect([v('H1'), v('I1')]).toEqual([3, 30]);
    expect([v('H5'), v('I5'), v('J5'), v('H6'), v('J6')]).toEqual([11, 12, 13, 21, 23]);
    expect([v('D6'), v('D7'), v('D8')]).toEqual([11, 22, 33]);
  });
  it('LAMBDA đóng gói biến của LET (closure)', () => {
    const { v } = wbWith([['A1', '=LET(k;10;f;LAMBDA(x;x*k);MAP({1;2};f))'], ['C1', '=LET(k;10;f;LAMBDA(x;x*k);k+f(2))']]);
    expect([v('A1'), v('B1')]).toEqual([10, 20]);
    expect(v('C1')).toBe(30);
  });
});

describe('Tham chiếu qua tên LET / tham số LAMBDA (duyệt chuyên gia)', () => {
  it('ROW, OFFSET, ISREF thấy tham chiếu như Excel; không bao giờ ra #ERROR!', () => {
    const { v } = wbWith([['A1', '5'], ['A2', '7'], ['B2', '3'],
      ['D1', '=LET(rr;B2;ROW(rr))'], ['D2', '=LET(x;A1;OFFSET(x;1;0))'], ['D3', '=LAMBDA(q;ISREF(q))(A1)'], ['D4', '=LAMBDA(q;ISREF(q))(3)'],
      ['D5', '=ISREF(A1)'], ['D6', '=LET(x;A1:A2;SUM(x))'], ['D7', '=LET(c;1;c)']]);
    expect(v('D1')).toBe(2);
    expect(v('D2')).toBe(7);
    expect(v('D3')).toBe(true);
    expect(v('D4')).toBe(false);
    expect(v('D5')).toBe(true);
    expect(v('D6')).toBe(12);
    expect(err(v('D7'))).toBe('#VALUE!');
  });
  it('_xlpm. chỉ gắn trong phạm vi LET/LAMBDA (tên vùng trùng chữ bên ngoài giữ nguyên)', async () => {
    const { wb } = wbWith([['A1', '=LET(tyle;2;tyle)*tyle'], ['A2', '=SUM(LET(x;1;x);x)']]);
    wb.defineName('tyle', '=10'); wb.defineName('x', '=3');
    const sx = await (await JSZip.loadAsync(await writeXlsx(wb))).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sx).toContain('_xlfn.LET(_xlpm.tyle,2,_xlpm.tyle)*tyle');
    expect(sx).toContain('SUM(_xlfn.LET(_xlpm.x,1,_xlpm.x),x)');
  });
});

describe('Tên LAMBDA (rà mã độc lập)', () => {
  it('mở lại (xlsx, bản nháp JSON) vẫn tính; đổi ô LAMBDA đọc thì tính lại; hoàn tác định nghĩa', async () => {
    const { wb, s, v } = wbWith([['A1', '5']]);
    wb.defineName('Sq', '=LAMBDA(n;n*n)');
    wb.defineName('AddA', '=LAMBDA(x;x+Sheet1!$A$1)');
    wb.setInput(s, 0, 1, '=Sq(3)');
    wb.setInput(s, 1, 1, '=AddA(1)');
    expect([v('B1'), v('B2')]).toEqual([9, 6]);
    wb.setInput(s, 0, 0, '10');
    expect(v('B2')).toBe(11);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    expect(w2.value(w2.sheetIds()[0], 0, 1)).toBe(9);
    const w3 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(w3.value(w3.sheetIds()[0], 1, 1)).toBe(11);
    const { wb: w4, s: s4, v: v4 } = wbWith([]);
    w4.setInput(s4, 0, 0, '=Tam(1)');
    expect(err(v4('A1'))).toBe('#NAME?');
    w4.defineName('Tam', '=LAMBDA(n;n*8)');
    expect(v4('A1')).toBe(8);
    w4.undo();
    expect(err(v4('A1'))).toBe('#NAME?');
    w4.redo();
    expect(v4('A1')).toBe(8);
  });
  it('tên LET trùng tên hàm: =LET(sum;5;SUM(sum;1)) = 6, ghi tệp SUM viết hoa', async () => {
    const { wb, v, raw } = wbWith([['A1', '=LET(sum;5;SUM(sum;1))']]);
    expect(v('A1')).toBe(6);
    expect(raw('A1')).toBe('=LET(sum;5;SUM(sum;1))');
    const sx = await (await JSZip.loadAsync(await writeXlsx(wb))).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sx).toContain('_xlfn.LET(_xlpm.sum,5,SUM(_xlpm.sum,1))');
  });
});

describe('Đọc/ghi .xlsx và chữ công thức (C1.5–C1.6)', () => {
  it('lamIn / lamOut khứ hồi', () => {
    const f = '=LAMBDA(x;LAMBDA(y;x+y)(2))(5)+1';
    expect(lamIn(f)).toBe('=XLAMCALL(LAMBDA(x;XLAMCALL(LAMBDA(y;x+y);2));5)+1');
    expect(lamOut(lamIn(f))).toBe(f);
    expect(lamIn('="LAMBDA(x)(1)"')).toBe('="LAMBDA(x)(1)"');
    expect([...declaredNames('=LET(a;1;b;LAMBDA(x;y;x+y);b(a;2))')].sort()).toEqual(['A', 'B', 'X', 'Y']);
  });
  it('ghi _xlfn.LET/_xlfn.LAMBDA và _xlpm. cho tên tham số; mở lại giống hệt', async () => {
    const { wb, s } = wbWith([['A1', '2'], ['B1', '=LET(x;A1;f;LAMBDA(n;n*x);f(3))'], ['C1', '=LAMBDA(x;x+1)(5)'], ['D1', '=MAP(A1:A1;LAMBDA(v;v*2))']]);
    wb.defineName('NhanDoi', '=LAMBDA(x;x*2)');
    const bytes = await writeXlsx(wb);
    const z = await JSZip.loadAsync(bytes);
    const sx = await z.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sx).toContain('_xlfn.LET(_xlpm.x,A1,_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n*_xlpm.x),_xlpm.f(3))');
    expect(sx).toContain('_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(5)');
    expect(sx).toContain('_xlfn.MAP(A1:A1,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))');
    const wx = await z.file('xl/workbook.xml')!.async('string');
    expect(wx).toContain('_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)');
    const { wb: w2 } = await readXlsx(bytes);
    const s2 = w2.sheetIds()[0];
    expect(w2.rawText(s2, 0, 1)).toBe('=LET(x;A1;f;LAMBDA(n;n*x);f(3))');
    expect(w2.rawText(s2, 0, 2)).toBe('=LAMBDA(x;x+1)(5)');
    expect(w2.value(s2, 0, 1)).toBe(6);
    expect(w2.value(s2, 0, 2)).toBe(6);
    expect(w2.listNames().find((n) => n.name === 'NhanDoi')?.expr).toBe('=LAMBDA(x;x*2)');
    void s;
  });
  it('tên ẩn _xlpm.* (xlm="1") giữ nguyên khi lưu lại', async () => {
    const { wb } = wbWith([['A1', '1']]);
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    let wx = await z.file('xl/workbook.xml')!.async('string');
    const extra = '<definedName name="_xlpm.ltb_H1" hidden="1" xlm="1">#NAME?</definedName>';
    wx = /<definedNames>/.test(wx) ? wx.replace('<definedNames>', '<definedNames>' + extra) : wx.replace('</sheets>', '</sheets><definedNames>' + extra + '</definedNames>');
    z.file('xl/workbook.xml', wx);
    const { wb: w2 } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    const z2 = await JSZip.loadAsync(await writeXlsx(w2));
    expect(await z2.file('xl/workbook.xml')!.async('string')).toContain(extra);
  });
});
