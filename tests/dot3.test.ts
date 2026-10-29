// Đợt 3: PivotTable, dữ liệu biểu đồ, thang trục, biểu đồ & thiết lập in trong .xlsx, trang in
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { Workbook } from '../src/core/workbook';
import { computePivot } from '../src/core/pivot';
import { niceScale, compactNum, seriesFromRange, chartData } from '../src/core/chart';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { buildPrintHtml } from '../src/core/print';

const rg = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
const DATA = [
  ['Nhóm', 'ĐVT', 'Hạng mục', 'Thành tiền', 'SL'],
  ['Thiết bị', 'bộ', 'Đèn LED', '204000000', '48'],
  ['Vật tư', 'cột', 'Cột thép', '326400000', '48'],
  ['Vật tư', 'm', 'Cáp ngầm', '148680000', '1260'],
  ['Thiết bị', 'tủ', 'Tủ điều khiển', '85500000', '3'],
  ['Xây lắp', 'móng', 'Móng cột', '64800000', '48'],
  ['Vật tư', 'm', 'Ống HDPE', '42480000', '1180'],
  ['', 'công', 'Nhân công', '61600000', '160'],
];
function book(rows = DATA) {
  const wb = Workbook.create();
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setInput(0, r, c, v); }));
  return wb;
}
const table = () => ({ header: DATA[0], rows: DATA.slice(1).map((r) => r.map((v, i) => (i >= 3 ? +v : v === '' ? null : v))) });

describe('PivotTable – tính toán (đối chiếu kết quả Pivot của Excel)', () => {
  it('nhóm theo hàng, tổng giá trị, tổng cộng, ô trống ghi (trống)', () => {
    const p = computePivot(table(), { rows: ['Nhóm'], cols: [], values: [{ field: 'Thành tiền', agg: 'sum' }], filters: [] });
    expect(p.cells).toEqual([
      ['Nhóm', 'Tổng Thành tiền'],
      ['Thiết bị', 289500000],
      ['Vật tư', 517560000],
      ['Xây lắp', 64800000],
      ['(trống)', 61600000],
      ['Tổng cộng', 933460000],
    ]);
    expect(p.kinds[0][0]).toBe('h');
    expect(p.kinds[5][1]).toBe('g');
  });
  it('bảng chéo hàng × cột, nhiều phép tổng hợp, lọc', () => {
    const p = computePivot(table(), { rows: ['Nhóm'], cols: ['ĐVT'], values: [{ field: 'SL', agg: 'sum' }], filters: [{ field: 'Nhóm', values: ['Vật tư', 'Thiết bị'] }] });
    expect(p.cells[0]).toEqual(['Tổng SL', 'ĐVT', null, null, null, null]);
    expect(p.cells[1]).toEqual(['Nhóm', 'bộ', 'cột', 'm', 'tủ', 'Tổng cộng']);
    expect(p.cells[2]).toEqual(['Thiết bị', 48, null, null, 3, 51]);
    expect(p.cells[3]).toEqual(['Vật tư', null, 48, 2440, null, 2488]);
    expect(p.cells[4]).toEqual(['Tổng cộng', 48, 48, 2440, 3, 2539]);
    const q = computePivot(table(), { rows: ['Nhóm'], cols: [], values: [{ field: 'Thành tiền', agg: 'count' }, { field: 'Thành tiền', agg: 'avg' }, { field: 'SL', agg: 'max' }], filters: [] });
    expect(q.cells[0]).toEqual(['Nhóm', 'Đếm Thành tiền', 'Trung bình Thành tiền', 'Lớn nhất SL']);
    expect(q.cells[2]).toEqual(['Vật tư', 3, 172520000, 1260]);
    expect(q.cells[5]).toEqual(['Tổng cộng', 7, 933460000 / 7, 1260]);
  });
  it('hai cấp hàng có dòng tổng phụ', () => {
    const p = computePivot(table(), { rows: ['Nhóm', 'ĐVT'], cols: [], values: [{ field: 'SL', agg: 'sum' }], filters: [] });
    expect(p.cells.slice(0, 5)).toEqual([
      ['Nhóm', 'ĐVT', 'Tổng SL'],
      ['Thiết bị', 'bộ', 48],
      [null, 'tủ', 3],
      ['Tổng Thiết bị', null, 51],
      ['Vật tư', 'cột', 48],
    ]);
    expect(p.kinds[3][0]).toBe('t');
    expect(p.cells[p.cells.length - 1]).toEqual(['Tổng cộng', null, 2747]);
  });
});

describe('PivotTable trong sổ tính', () => {
  it('tạo trên sheet mới, làm mới khi dữ liệu đổi, hoàn tác', () => {
    const wb = book();
    const { sheet, id } = wb.addPivot({ src: { sheet: 0, range: rg(0, 0, 7, 4) }, rows: ['Nhóm'], cols: [], values: [{ field: 'Thành tiền', agg: 'sum' }], filters: [] });
    expect(wb.sheetName(sheet)).toBe('Pivot 1');
    expect(wb.display(sheet, 2, 0).text).toBe('Nhóm');
    expect(wb.value(sheet, 3, 1)).toBe(289500000);
    expect(wb.display(sheet, 3, 1).text).toBe('289.500.000');
    expect(wb.style(sheet, 2, 0).b).toBe(true);
    expect(wb.meta(sheet).colW.get(1)).toBeGreaterThanOrEqual(Math.ceil('Tổng Thành tiền'.length * 7.6 + 16)); // tự nới cột
    wb.setInput(0, 1, 3, '1000000');
    wb.refreshPivot(sheet, id);
    expect(wb.value(sheet, 3, 1)).toBe(86500000);
    wb.undo();
    expect(wb.value(sheet, 3, 1)).toBe(289500000);
    expect(wb.pivotAt(sheet, 4, 1)?.id).toBe(id);
  });
  it('làm mới khi vùng nguồn thêm cột mới vẫn giữ đúng trường theo tên', () => {
    const wb = book();
    const { sheet, id } = wb.addPivot({ src: { sheet: 0, range: rg(0, 0, 7, 4) }, rows: ['Nhóm'], cols: [], values: [{ field: 'SL', agg: 'sum' }], filters: [] });
    wb.insertCols(0, 1, 1);
    wb.refreshPivot(sheet, id);
    expect(wb.value(sheet, 3, 1)).toBe(51);
  });
});

describe('Biểu đồ – dữ liệu và trục', () => {
  it('thang trục đẹp', () => {
    expect(niceScale(0, 326400000, 5)).toEqual({ lo: 0, hi: 400000000, step: 100000000 });
    expect(niceScale(-12, 47, 5)).toEqual({ lo: -20, hi: 60, step: 20 });
    expect(niceScale(0, 0, 5)).toEqual({ lo: 0, hi: 1, step: 0.2 });
  });
  it('số rút gọn kiểu Việt Nam', () => {
    expect(compactNum(1250)).toBe('1.250');
    expect(compactNum(1500000)).toBe('1,5 tr');
    expect(compactNum(326400000)).toBe('326,4 tr');
    expect(compactNum(2000000000)).toBe('2 tỷ');
    expect(compactNum(0.25)).toBe('0,25');
  });
  it('tự nhận tiêu đề cột và cột nhãn từ vùng chọn', () => {
    const wb = book();
    const s = seriesFromRange(wb, 0, rg(0, 2, 7, 4));
    expect(s.length).toBe(2);
    expect(s[0].name).toEqual({ sheet: 0, range: rg(0, 3, 0, 3) });
    expect(s[0].cats).toEqual({ sheet: 0, range: rg(1, 2, 7, 2) });
    expect(s[0].vals).toEqual({ sheet: 0, range: rg(1, 3, 7, 3) });
    const d = chartData(wb, { type: 'col', series: s });
    expect(d.cats[0]).toBe('Đèn LED');
    expect(d.series[1].name).toBe('SL');
    expect(d.series[0].values[1]).toBe(326400000);
  });
  it('nhiều cột chữ ở đầu: lấy cột chữ sát số làm nhãn, bỏ cột không có số', () => {
    const wb = book();
    const ss = seriesFromRange(wb, 0, rg(0, 0, 7, 4));
    expect(ss.map((x) => x.vals.range.c1)).toEqual([3, 4]);
    expect(ss[0].cats?.range.c1).toBe(2);
    expect(chartData(wb, { type: 'col', series: ss }).cats[0]).toBe('Đèn LED');
  });
  it('chuỗi theo hàng khi chọn', () => {
    const wb = book([['', 'T1', 'T2', 'T3'], ['LED', '1', '2', '3'], ['Cột', '4', '5', '6']]);
    const s = seriesFromRange(wb, 0, rg(0, 0, 2, 3), true);
    expect(s.map((x) => chartData(wb, { type: 'line', series: [x] }).series[0].name)).toEqual(['LED', 'Cột']);
    expect(chartData(wb, { type: 'line', series: s }).cats).toEqual(['T1', 'T2', 'T3']);
  });
});

describe('Biểu đồ trong sổ tính', () => {
  it('thêm, sửa, xóa, hoàn tác; dời theo khi chèn hàng', () => {
    const wb = book();
    const id = wb.addChart(0, { type: 'col', title: 'Thành tiền', series: seriesFromRange(wb, 0, rg(0, 2, 7, 3)), anchor: { r: 1, c: 6, dx: 0, dy: 0 }, w: 480, h: 280 });
    expect(wb.meta(0).charts.length).toBe(1);
    wb.updateChart(0, id, { type: 'pie' });
    expect(wb.meta(0).charts[0].type).toBe('pie');
    wb.insertRows(0, 0, 2);
    const c = wb.meta(0).charts[0];
    expect(c.anchor.r).toBe(3);
    expect(c.series[0].vals.range).toEqual(rg(3, 3, 9, 3));
    wb.undo();
    expect(wb.meta(0).charts[0].anchor.r).toBe(1);
    wb.removeChart(0, id);
    expect(wb.meta(0).charts.length).toBe(0);
    wb.undo();
    expect(wb.meta(0).charts[0].type).toBe('pie');
  });
});

describe('Biểu đồ trong tệp .xlsx', () => {
  it('đọc biểu đồ cột, đường, tròn do Excel/openpyxl tạo', async () => {
    const { wb, warnings } = await readXlsx(readFileSync('tests/fixtures/bieudo.xlsx'));
    const charts = wb.meta(0).charts;
    expect(charts.map((c) => c.type).sort()).toEqual(['col', 'line', 'pie']);
    const bar = charts.find((c) => c.type === 'col')!;
    expect(bar.title).toBe('Sản lượng theo tháng');
    const d = chartData(wb, bar);
    expect(d.cats).toEqual(['T1', 'T2', 'T3', 'T4']);
    expect(d.series.map((s) => s.name)).toEqual(['Đèn LED', 'Cột thép']);
    expect(d.series[0].values).toEqual([120, 150, 90, 200]);
    expect(bar.anchor).toMatchObject({ r: 1, c: 4 });
    expect(bar.w).toBeGreaterThan(300);
    expect(warnings).toEqual([]);
    expect(wb.meta(0).print).toMatchObject({ orientation: 'landscape', paper: 'A4', titleRowRange: [0, 0] });
  });
  it('ghi biểu đồ vào .xlsx đúng chuẩn (có phần biểu đồ, bản vẽ, liên kết) rồi đọc lại', async () => {
    const wb = book();
    wb.addChart(0, { type: 'bar', title: 'Theo hạng mục', series: seriesFromRange(wb, 0, rg(0, 2, 7, 3)), anchor: { r: 2, c: 6, dx: 10, dy: 5 }, w: 480, h: 300, legend: 'r', labels: true });
    wb.addChart(0, { type: 'doughnut', series: seriesFromRange(wb, 0, rg(0, 2, 7, 3)), anchor: { r: 20, c: 6, dx: 0, dy: 0 }, w: 360, h: 300 });
    const s2 = wb.addSheet('Báo cáo 2026');
    wb.setInput(s2, 0, 0, 'x');
    wb.addChart(s2, { type: 'line', series: seriesFromRange(wb, 0, rg(0, 2, 7, 4)), anchor: { r: 1, c: 1, dx: 0, dy: 0 }, w: 400, h: 250 });
    wb.setPrint(0, { orientation: 'landscape', paper: 'A3', fitWidth: true, titleRows: 1, gridlines: true, margins: 'narrow' });
    const buf = await writeXlsx(wb);
    const zip = await JSZip.loadAsync(buf);
    const names = Object.keys(zip.files);
    expect(names.filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n)).length).toBe(3);
    expect(names.filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).length).toBe(2);
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('drawingml.chart+xml');
    const chart1 = await zip.file('xl/charts/chart1.xml')!.async('string');
    expect(chart1).toContain("<c:f>'Sheet1'!$D$2:$D$8</c:f>".replace("'Sheet1'", 'Sheet1'));
    expect(chart1).toContain('<c:barDir val="bar"/>');
    const { wb: back } = await readXlsx(buf);
    const types = back.sheetIds().flatMap((id) => back.meta(id).charts.map((c) => c.type));
    expect(types.sort()).toEqual(['bar', 'doughnut', 'line']);
    const c0 = back.meta(0).charts.find((c) => c.type === 'bar')!;
    expect(c0).toMatchObject({ title: 'Theo hạng mục', legend: 'r', labels: true, anchor: { r: 2, c: 6 } });
    expect(Math.abs(c0.w - 480)).toBeLessThan(3);
    expect(chartData(back, c0).series[0].values).toEqual(chartData(wb, wb.meta(0).charts[0]).series[0].values);
    expect(back.meta(0).print).toMatchObject({ orientation: 'landscape', paper: 'A3', fit: { w: 1, h: 0 }, titleRowRange: [0, 0], gridlines: true });
  });
});

describe('Trang in', () => {
  it('bảng in có giá trị đã định dạng, ô gộp, cột ẩn bị bỏ, khổ giấy & lặp tiêu đề', () => {
    const wb = book();
    wb.setStyle(0, rg(1, 3, 7, 3), { fmt: '#,##0' });
    wb.setStyle(0, rg(0, 0, 0, 4), { b: true, bg: '#dcefe5' });
    wb.insertRows(0, 0, 1);
    wb.setInput(0, 0, 0, 'BẢNG DỰ TOÁN <thử>');
    wb.merge(0, rg(0, 0, 0, 4));
    wb.setColWidth(0, 4, 0);
    wb.setColWidth(0, 3, 110);
    wb.setPrint(0, { orientation: 'landscape', paper: 'A4', fitWidth: true, titleRows: 2, gridlines: false });
    const html = buildPrintHtml(wb, 0, {});
    expect(html).toContain('size: 297mm 210mm');
    expect(html).toContain('204.000.000');
    expect(html).toContain('BẢNG DỰ TOÁN &lt;thử&gt;');
    expect(html).toMatch(/colspan="4"/);
    expect(html).not.toContain('1260');
    expect((html.match(/class="page"/g) ?? []).length).toBe(1);
    expect(html).toMatch(/background:#dcefe5/);
  });
});
