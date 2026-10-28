// Các lỗi tìm được khi review Đợt 3 — mỗi ca một test để không tái phát
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { Workbook } from '../src/core/workbook';
import { computePivot } from '../src/core/pivot';
import { chartData } from '../src/core/chart';
import { buildPrintHtml } from '../src/core/print';
import { writeXlsx, readXlsx } from '../src/core/xlsx';

const rg = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
function mk(d: string[][] = [['K', 'V'], ['a', '1'], ['b', '2']]) {
  const wb = Workbook.create();
  d.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setInput(0, r, c, v); }));
  return wb;
}

describe('Review Đợt 3', () => {
  it('H1: tên phông độc hại không phá được thẻ style trong trang in; có CSP chặn script', () => {
    const wb = mk([['x']]);
    wb.setStyle(0, rg(0, 0, 0, 0), { fn: 'x"><img src=x onerror=alert(1)>', fs: 14 });
    const html = buildPrintHtml(wb, 0);
    expect(html).not.toMatch(/<img src=x/);
    expect(html).toMatch(/font-size:14pt/); // phần sau font-family không bị cắt
    expect(html).toMatch(/Content-Security-Policy/);
    wb.setStyle(0, rg(0, 0, 0, 0), { fn: 'Times New Roman' });
    expect(buildPrintHtml(wb, 0)).toMatch(/style="[^"]*font-family:'Times New Roman'[^"]*text-align/);
  });

  it('H2: xóa sheet nguồn → biểu đồ/Pivot không bám sang sheet mới cùng mã; hoàn tác trả lại', async () => {
    const wb = mk();
    const s2 = wb.addSheet('Data2');
    wb.setInput(s2, 0, 0, 'X'); wb.setInput(s2, 1, 0, '5'); wb.setInput(s2, 2, 0, '6');
    wb.addChart(0, { type: 'col', series: [{ vals: { sheet: s2, range: rg(1, 0, 2, 0) } }], anchor: { r: 0, c: 5, dx: 0, dy: 0 }, w: 400, h: 300 });
    wb.addPivot({ src: { sheet: s2, range: rg(0, 0, 2, 0) }, rows: [], cols: [], values: [{ field: 'X', agg: 'sum' }], filters: [], at: { r: 10, c: 0 } }, 0);
    wb.removeSheet(s2);
    expect(wb.meta(0).charts[0].series.length).toBe(0);
    expect(wb.meta(0).pivots.length).toBe(0);
    const s3 = wb.addSheet('Other');
    expect(s3).toBe(s2);
    wb.setInput(s3, 1, 0, '999');
    expect(chartData(wb, wb.meta(0).charts[0]).series).toEqual([]);
    // tệp lưu ra không có tham chiếu "undefined!"
    const zip = await JSZip.loadAsync(await writeXlsx(wb));
    const all = await Promise.all(Object.keys(zip.files).filter((f) => f.startsWith('xl/charts/')).map((f) => zip.file(f)!.async('string')));
    expect(all.join('')).not.toMatch(/undefined!/);
    wb.undo(); wb.undo(); wb.undo(); // addSheet, setInput, removeSheet
    expect(wb.sheetName(s2)).toBe('Data2');
    expect(wb.meta(0).charts[0].series[0].vals.sheet).toBe(s2);
    expect(wb.meta(0).pivots.length).toBe(1);
  });

  it('M1: mục tên “*” hoặc chứa “|” không bị gộp nhầm với tổng', () => {
    const t = { header: ['K', 'C', 'V'], rows: [['*', 'x', 10], ['a', 'x', 1], ['a|b', 'c', 1], ['a', 'b|c', 100]] as unknown[][] };
    const r1 = computePivot(t, { rows: ['K'], cols: [], values: [{ field: 'V', agg: 'sum' }], filters: [] });
    expect(r1.cells.find((x) => x[0] === '*')?.[1]).toBe(10);
    const r2 = computePivot(t, { rows: ['K'], cols: ['C'], values: [{ field: 'V', agg: 'sum' }], filters: [] });
    const g = r2.cells[r2.cells.length - 1];
    expect(g[g.length - 1]).toBe(112);
    const ab = r2.cells.find((x) => x[0] === 'a|b')!;
    expect(ab.filter((v) => typeof v === 'number')).toEqual([1, 1]);
  });

  it('M2: trường ngày hiện theo ngày đã định dạng, không phải số sê-ri', () => {
    const wb = mk([['Ngày', 'SL'], ['01/02/2024', '5'], ['02/02/2024', '7']]);
    const { sheet } = wb.addPivot({ src: { sheet: 0, range: rg(0, 0, 2, 1) }, rows: ['Ngày'], cols: [], values: [{ field: 'SL', agg: 'sum' }], filters: [] });
    expect(wb.display(sheet, 3, 0).text).toBe('01/02/2024');
  });

  it('M3: không cho đặt Pivot chồng lên vùng nguồn (kể cả nguồn cả cột)', () => {
    const wb = mk();
    expect(() => wb.addPivot({ src: { sheet: 0, range: rg(0, 0, 1048575, 1) }, rows: ['K'], cols: [], values: [{ field: 'V', agg: 'sum' }], filters: [], at: { r: 6, c: 0 } }, 0)).toThrow(/chồng lên vùng dữ liệu nguồn/);
    expect(wb.meta(0).pivots.length).toBe(0);
  });

  it('M4: không chèn/xóa hàng cắt ngang Pivot; xóa trọn vùng Pivot thì bỏ Pivot', () => {
    const wb = mk();
    wb.addPivot({ src: { sheet: 0, range: rg(0, 0, 2, 1) }, rows: ['K'], cols: [], values: [{ field: 'V', agg: 'sum' }], filters: [], at: { r: 5, c: 0 } }, 0);
    expect(() => wb.insertRows(0, 6, 1)).toThrow(/cắt ngang PivotTable/);
    expect(() => wb.deleteRows(0, 7, 2)).toThrow(/cắt ngang PivotTable/);
    wb.insertRows(0, 5, 1); // chèn ngay trên Pivot: Pivot dời xuống
    expect(wb.meta(0).pivots[0].area!.r1).toBe(6);
    wb.deleteRows(0, 6, 4);
    expect(wb.meta(0).pivots.length).toBe(0);
    wb.undo();
    expect(wb.meta(0).pivots.length).toBe(1);
  });

  it('M6: vùng in dời theo khi chèn hàng', () => {
    const wb = mk();
    wb.setPrint(0, { orientation: 'portrait', paper: 'A4', fitWidth: true, titleRows: 0, gridlines: false, area: rg(0, 0, 2, 1) });
    wb.insertRows(0, 0, 3);
    expect(wb.meta(0).print!.area).toEqual(rg(3, 0, 5, 1));
    wb.undo();
    expect(wb.meta(0).print!.area).toEqual(rg(0, 0, 2, 1));
  });

  it('M7: lặp hàng tiêu đề không cắt ngang ô gộp', () => {
    const wb = mk([['A', 'B'], ['', 'b2'], ['c', 'd']]);
    wb.merge(0, rg(0, 0, 1, 0));
    const html = buildPrintHtml(wb, 0, { setup: { orientation: 'portrait', paper: 'A4', fitWidth: true, titleRows: 1, gridlines: false } });
    const thead = /<thead>([\s\S]*?)<\/thead>/.exec(html)![1];
    expect((thead.match(/<tr/g) || []).length).toBe(2);
    expect(/<tbody>([\s\S]*?)<\/tbody>/.exec(html)![1]).toMatch(/>c<\/td><td[^>]*>d</);
  });

  it('L1: tạo Pivot ở sheet mới là một bước hoàn tác', () => {
    const wb = mk();
    wb.addPivot({ src: { sheet: 0, range: rg(0, 0, 2, 1) }, rows: ['K'], cols: [], values: [{ field: 'V', agg: 'sum' }], filters: [] });
    expect(wb.sheetIds().length).toBe(2);
    wb.undo();
    expect(wb.sheetIds().length).toBe(1);
    wb.redo();
    expect(wb.sheetIds().length).toBe(2);
    expect(wb.meta(wb.sheetIds()[1]).pivots.length).toBe(1);
  });

  it('L4: tên cột trùng được đánh số để chọn riêng', () => {
    const wb = mk([['V', 'V'], ['1', '10'], ['2', '20']]);
    expect(wb.pivotTable({ sheet: 0, range: rg(0, 0, 2, 1) }).header).toEqual(['V', 'V2']);
  });

  it('L5: bộ đệm biểu đồ trong tệp giữ đúng chỉ số khi có hàng ẩn; mỗi chuỗi dùng nhãn của nó', async () => {
    const wb = mk([['K', 'V'], ['a', '1'], ['b', '2'], ['c', '3']]);
    wb.addChart(0, { type: 'col', series: [{ cats: { sheet: 0, range: rg(1, 0, 3, 0) }, vals: { sheet: 0, range: rg(1, 1, 3, 1) } }], anchor: { r: 0, c: 4, dx: 0, dy: 0 }, w: 400, h: 300 });
    wb.meta(0).rowH.setHidden([2]);
    const zip = await JSZip.loadAsync(await writeXlsx(wb));
    const x = await zip.file('xl/charts/chart1.xml')!.async('string');
    expect(x).toMatch(/<c:numCache>.*<c:ptCount val="3"\/>.*<c:pt idx="2"><c:v>3<\/c:v>/);
  });

  it('L7: mở được tệp có Target tuyệt đối (openpyxl) và có ghi chú', async () => {
    const wb0 = mk();
    const zip = await JSZip.loadAsync(await writeXlsx(wb0));
    const rp = 'xl/_rels/workbook.xml.rels';
    const rx = await zip.file(rp)!.async('string');
    zip.file(rp, rx.replace(/Target="(?!\/)([^"]+)"/g, 'Target="/xl/$1"'));
    const { wb } = await readXlsx(await zip.generateAsync({ type: 'uint8array' }));
    expect(wb.display(0, 1, 1).text).toBe('1');
  });
  it('L7b: tệp openpyxl có ghi chú + biểu đồ mở được, báo trước ghi chú sẽ mất', async () => {
    const { wb, warnings } = await readXlsx(readFileSync(new URL('./fixtures/ghichu-bieudo.xlsx', import.meta.url)));
    expect(wb.meta(0).charts.length).toBe(1);
    expect(warnings.join(' ')).toMatch(/1 ghi chú/);
  });
});
