import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { scatterFromRange, seriesFromRange, chartData } from '../src/core/chart';
import { writeFileSync } from 'node:fs';

const mk = () => {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  [['Tháng', 'Doanh thu', 'Tỷ lệ'], ['1', '100', '0,1'], ['2', '150', '0,2'], ['3', '120', '0,15']].forEach((row, r) => row.forEach((v, c) => wb.setInput(s, r, c, v)));
  return { wb, s };
};
describe('Biểu đồ 0.9', () => {
  it('phân tán XY: X lấy từ cột đầu; Kết hợp có trục phụ; giữ khi lưu .xlsx', async () => {
    const { wb, s } = mk();
    const xy = scatterFromRange(wb, s, { r1: 0, c1: 0, r2: 3, c2: 1 });
    expect(chartData(wb, { type: 'scatter', series: xy }).series[0].xs).toEqual([1, 2, 3]);
    wb.addChart(s, { type: 'scatter', series: xy, anchor: { r: 5, c: 0, dx: 0, dy: 0 }, w: 400, h: 240, xTitle: 'Tháng', yTitle: 'Triệu đồng' });
    const ser = seriesFromRange(wb, s, { r1: 0, c1: 0, r2: 3, c2: 2 }).map((x, i) => ({ ...x, kind: i === 0 ? 'col' as const : 'line' as const, y2: i === 1 }));
    wb.addChart(s, { type: 'combo', series: ser, anchor: { r: 5, c: 8, dx: 0, dy: 0 }, w: 400, h: 240, colors: 1, fill: '#f2f2f2', border: '#bfbfbf' });
    const buf = await writeXlsx(wb);
    if (process.env.KEEP) writeFileSync(process.env.KEEP, buf);
    const { wb: w2 } = await readXlsx(buf);
    const cs = w2.meta(w2.sheetIds()[0]).charts;
    expect(cs.map((c) => c.type)).toEqual(['scatter', 'combo']);
    expect(cs[0].xTitle).toBe('Tháng'); expect(cs[0].yTitle).toBe('Triệu đồng');
    expect(cs[1].series.map((x) => [x.kind, !!x.y2])).toEqual([['col', false], ['line', true], ['line', false]]);
  });
});
describe('Sparkline', () => {
  it('thêm theo hàng, lưu/mở .xlsx, dời khi chèn hàng', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    [['1', '3', '2', '5'], ['4', '-1', '2', '0']].forEach((row, r) => row.forEach((v, c) => wb.setInput(s, r, c, v)));
    expect(wb.addSparklines(s, 'line', { sheet: s, range: { r1: 0, c1: 0, r2: 1, c2: 3 } }, { r1: 0, c1: 4, r2: 1, c2: 4 })).toBe(2);
    wb.updateSparkGroup(s, { r: 0, c: 4 }, { hi: true, color: '#217346' });
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const g = w2.meta(w2.sheetIds()[0]).sparks![0];
    expect(g.type).toBe('line'); expect(g.hi).toBe(true); expect(g.color).toBe('#217346');
    expect(g.items.map((x) => [x.r, x.c, x.src.range.r1])).toEqual([[0, 4, 0], [1, 4, 1]]);
    wb.insertRows(s, 0, 1);
    expect(wb.sparkAt(s, 1, 4)?.g.items[0].src.range.r1).toBe(1);
    wb.undo();
    expect(wb.sparkAt(s, 0, 4)).not.toBe(null);
    expect(() => wb.addSparklines(s, 'column', { sheet: s, range: { r1: 0, c1: 0, r2: 1, c2: 3 } }, { r1: 5, c1: 0, r2: 6, c2: 2 })).toThrow();
  });
});
describe('Đối tượng nổi', () => {
  it('ảnh, hình, hộp văn bản lưu/mở .xlsx', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    wb.addChart(s, { type: 'pic', img: png, alt: 'Logo', series: [], anchor: { r: 1, c: 1, dx: 0, dy: 0 }, w: 100, h: 50 });
    wb.addChart(s, { type: 'shape', shape: 'ellipse', text: 'Hạng mục A', fill: '#70ad47', series: [], anchor: { r: 5, c: 1, dx: 0, dy: 0 }, w: 140, h: 90 });
    wb.addChart(s, { type: 'text', text: 'Ghi chú:\nDòng 2', fs: 12, series: [], anchor: { r: 10, c: 1, dx: 0, dy: 0 }, w: 200, h: 60 });
    const id = wb.meta(s).charts[0].id!;
    wb.orderChart(s, id, 'front');
    expect(wb.meta(s).charts[2].id).toBe(id);
    const buf = await writeXlsx(wb);
    if (process.env.KEEP2) writeFileSync(process.env.KEEP2, buf);
    const { wb: w2 } = await readXlsx(buf);
    const o = w2.meta(w2.sheetIds()[0]).charts;
    expect(o.map((x) => x.type)).toEqual(['shape', 'text', 'pic']);
    expect(o[0].shape).toBe('ellipse'); expect(o[0].text).toBe('Hạng mục A'); expect(o[0].fill).toBe('#70ad47');
    expect(o[1].text).toBe('Ghi chú:\nDòng 2'); expect(o[1].fs).toBe(12);
    expect(o[2].img).toBe(png); expect(o[2].alt).toBe('Logo'); expect(o[2].w).toBe(100);
  });
});
