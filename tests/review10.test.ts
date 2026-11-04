import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

describe('Đối chiếu tệp đèn LED do Excel lưu (04/11/2026)', () => {
  it('TEXT đọc mã định dạng theo vùng Việt Nam như Excel: "," thập phân, "." nhóm nghìn', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const t = (f: string) => { wb.setInput(s, 0, 0, f); return wb.value(s, 0, 0); };
    expect(t('=TEXT(29;"#,##0")')).toBe('29,0');
    expect(t('=TEXT(1234,5;"#.##0")')).toBe('1.235');
    expect(t('=TEXT(0,256;"0%")')).toBe('26%');
    expect(t('=TEXT(46000;"dd/mm/yyyy")')).toBe('09/12/2025');
    expect(t('=TEXT("abc";"0")')).toBe('abc');
    expect(t('="Có "&TEXT(5;"00")&" bộ"')).toBe('Có 05 bộ');
  });
  it('công thức mảng kiểu cũ (t="array", không cm): ô trong vùng không chặn mảng; vùng cố định, thiếu thì #N/A; lưu lại vẫn là mảng kiểu cũ', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '1'); wb.setInput(s, 1, 0, '2'); wb.setInput(s, 2, 0, '3');
    wb.setInput(s, 0, 2, 'x'); // để có ô C1 trong tệp
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    let x = await z.file('xl/worksheets/sheet1.xml')!.async('string');
    // E1:H1 = TOROW(FILTER(A1:A3;A1:A3>1)) kiểu cũ; F1:H1 có <f/> rỗng + giá trị lưu sẵn #NAME?
    x = x.replace('</row>', '<c r="E1" t="e"><f t="array" ref="E1:H1">_xlfn.TOROW(_xlfn._xlws.FILTER(A1:A3,A1:A3&gt;1))</f><v>#NAME?</v></c><c r="F1" t="e"><f/><v>#NAME?</v></c><c r="G1" t="e"><f/><v>#NAME?</v></c><c r="H1" t="e"><f/><v>#NAME?</v></c><c r="I1" t="s"><v>0</v></c></row>');
    z.file('xl/worksheets/sheet1.xml', x);
    const { wb: w2 } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    const s2 = w2.sheetIds()[0], v = (c: number) => { const y: any = w2.value(s2, 0, c); return y && typeof y === 'object' ? y.value : y; };
    expect([v(4), v(5), v(6), v(7)]).toEqual([2, 3, '#N/A', '#N/A']);
    expect(w2.rawText(s2, 0, 4)).toBe('=TOROW(FILTER(A1:A3;A1:A3>1))');
    const out = await (await JSZip.loadAsync(await writeXlsx(w2))).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(out).toMatch(/<c r="E1"(?![^>]*cm=)[^>]*><f t="array" ref="E1:H1">/);
  });
  it('rà mã 1.0.1: TEXT giữ lỗi, ô trống = 0, số dạng chữ', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const t = (f: string) => { wb.setInput(s, 0, 0, f); const v: any = wb.value(s, 0, 0); return v && typeof v === 'object' ? v.value : v; };
    expect(t('=TEXT(1/0;"0")')).toBe('#DIV/0!');
    expect(t('=TEXT(NA();"0")')).toBe('#N/A');
    expect(t('=TEXT(Z99;"0")')).toBe('0');
    expect(t('=TEXT("123";"0,0")')).toBe('123,0');
  });
  it('rà mã 1.0.1: mọi công thức mảng kiểu cũ nhiều ô (TRANSPOSE, phép tính trên vùng) giữ giá trị; không sửa được một phần mảng', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '1'); wb.setInput(s, 1, 0, '2'); wb.setInput(s, 2, 0, '3');
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    let x = await z.file('xl/worksheets/sheet1.xml')!.async('string');
    x = x.replace('</row>', '<c r="E1"><f t="array" ref="E1:G1">TRANSPOSE(A1:A3*2)</f><v>2</v></c><c r="F1"><f/><v>4</v></c><c r="G1"><f/><v>6</v></c><c r="H1"><f t="array" ref="H1:J1">A1*2</f><v>2</v></c><c r="I1"><v>2</v></c><c r="J1"><v>2</v></c></row>');
    z.file('xl/worksheets/sheet1.xml', x);
    const { wb: w2 } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    const s2 = w2.sheetIds()[0];
    expect([4, 5, 6, 7, 8, 9].map((c) => w2.value(s2, 0, c))).toEqual([2, 4, 6, 2, 2, 2]);
    expect(() => w2.setInput(s2, 0, 6, '9')).toThrow(/một phần của mảng/);
    const out = await (await JSZip.loadAsync(await writeXlsx(w2))).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(out).toMatch(/<f t="array" ref="E1:G1">TRANSPOSE\(A1:A3\*2\)<\/f>/);
    expect(out).toMatch(/<f t="array" ref="H1:J1">A1\*2<\/f>/);
  });
  it('rà mã 1.0.1: vòng không hội tụ không làm chậm; đổi định dạng không tính lại vòng; công thức mới đọc ô vòng ra đúng', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    for (let r = 0; r < 10; r++) wb.setInput(s, r, 0, String(r + 1));
    wb.setInput(s, 10, 0, '=SUM(A1:A11)'); // tự cộng chính nó: không hội tụ
    const t0 = performance.now();
    for (let i = 0; i < 60; i++) wb.setInput(s, i, 2, '=A11*' + (i + 1));
    for (let i = 0; i < 5; i++) wb.setStyle(s, { r1: 0, c1: 2, r2: 59, c2: 2 }, { b: i % 2 === 0 });
    expect(performance.now() - t0).toBeLessThan(8000);
    const v: any = wb.value(s, 10, 0);
    expect(v?.type).toBe('CYCLE');
    // vòng hội tụ + công thức mới đọc ô vòng
    const w = Workbook.create(); const q = w.sheetIds()[0];
    w.setInput(q, 0, 0, '1000'); w.setInput(q, 0, 1, '=A1+C1'); w.setInput(q, 0, 2, '=ROUND(B1*10%;0)');
    w.setInput(q, 3, 3, '=IFERROR(B1;0)+1');
    expect(w.value(q, 3, 3)).toBe(1112);
  }, 30000);
  it('rà mã 1.0.1 vòng 2: vòng không hội tụ không xóa vòng hội tụ khác; đổi ô đầu vào thì tính lại; bước đều vẫn ra số', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '1000'); wb.setInput(s, 0, 1, '=A1+C1'); wb.setInput(s, 0, 2, '=ROUND(B1*10%;0)');
    expect(wb.value(s, 0, 1)).toBe(1111);
    wb.setInput(s, 9, 0, '=A10+1');
    expect(wb.value(s, 0, 1)).toBe(1111);
    expect((wb.value(s, 9, 0) as any)?.type).toBe('CYCLE');
    wb.setInput(s, 5, 4, '2'); wb.setInput(s, 5, 3, '=D6*E6+1');
    expect((wb.value(s, 5, 3) as any)?.type).toBe('CYCLE');
    wb.setInput(s, 5, 4, '0,5');
    expect(Math.abs((wb.value(s, 5, 3) as number) - 2)).toBeLessThan(0.01);
    wb.setInput(s, 7, 6, '=MIN(G8+1;10)');
    expect(wb.value(s, 7, 6)).toBe(10);
  });
  it('rà mã 1.0.1 vòng 2: không xóa riêng một phần mảng kiểu cũ', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '1'); wb.setInput(s, 1, 0, '2'); wb.setInput(s, 2, 0, '3');
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    let x = await z.file('xl/worksheets/sheet1.xml')!.async('string');
    x = x.replace('</row>', '<c r="E1"><f t="array" ref="E1:G1">TRANSPOSE(A1:A3*2)</f><v>2</v></c><c r="F1"><f/><v>4</v></c><c r="G1"><f/><v>6</v></c></row>');
    z.file('xl/worksheets/sheet1.xml', x);
    const { wb: w2 } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    const s2 = w2.sheetIds()[0];
    expect(() => w2.clear(s2, { r1: 0, c1: 5, r2: 0, c2: 6 })).toThrow(/một phần của mảng/);
    expect(() => w2.clear(s2, { r1: 0, c1: 4, r2: 0, c2: 6 })).not.toThrow();
  });
});
