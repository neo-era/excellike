// Bản 0.5: in ấn như Excel 365 – tự chia trang (xem trước = bản in), co giãn, vừa X×Y trang, ngắt trang thủ công,
// hàng/cột lặp lại, thứ tự trang, đầu trang/chân trang (mã &P &N &D &A &F), căn giữa, in tiêu đề hàng/cột, khứ hồi .xlsx.
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { normPrint, paginate, hfParse, hfRender, MARGIN_PRESETS } from '../src/core/paginate';
import { buildPrintHtml } from '../src/core/print';

const rg = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
/** Sổ 60 hàng × 12 cột (ô mặc định 64×20 px) */
function grid(rows = 60, cols = 12) {
  const wb = Workbook.create();
  const s = wb.sheetIds()[0];
  const data: (string | number)[][] = [];
  for (let r = 0; r < rows; r++) { const row: (string | number)[] = []; for (let c = 0; c < cols; c++) row.push(r * 100 + c); data.push(row); }
  wb.loadSheet(s, data);
  return { wb, s };
}
const rowsOf = (p: { rows: number[] }) => [p.rows[0], p.rows[p.rows.length - 1]];
const colsOf = (p: { cols: number[] }) => [p.cols[0], p.cols[p.cols.length - 1]];

describe('Thiết lập in mặc định như Excel', () => {
  it('Dọc, A4, Không co giãn (100%), lề Thường 0,75/0,7/0,3 inch, xuống rồi sang', () => {
    const p = normPrint(null);
    expect(p.orientation).toBe('portrait');
    expect(p.paper).toBe('A4');
    expect(p.scale).toBe(100);
    expect(p.fit).toBeNull();
    expect(p.mg).toEqual(MARGIN_PRESETS.normal);
    expect(MARGIN_PRESETS.normal).toEqual({ t: 0.75, b: 0.75, l: 0.7, r: 0.7, h: 0.3, f: 0.3 });
    expect(MARGIN_PRESETS.wide).toEqual({ t: 1, b: 1, l: 1, r: 1, h: 0.5, f: 0.5 });
    expect(MARGIN_PRESETS.narrow).toEqual({ t: 0.75, b: 0.75, l: 0.25, r: 0.25, h: 0.3, f: 0.3 });
    expect(p.order).toBe('down');
  });
  it('thiết lập cũ “Thu nhỏ vừa chiều ngang” đọc thành Vừa 1 trang ngang', () => {
    const p = normPrint({ orientation: 'portrait', paper: 'A4', fitWidth: true, titleRows: 2, gridlines: false });
    expect(p.fit).toEqual({ w: 1, h: 0 });
    expect(p.titleRowRange).toEqual([0, 1]);
  });
});

describe('Chia trang', () => {
  it('100%: 48 hàng × 10 cột mỗi trang A4 dọc → 4 trang, thứ tự xuống rồi sang', () => {
    const { wb, s } = grid();
    const L = paginate(wb, s, normPrint(null));
    expect(L.scale).toBe(1);
    expect(L.pages.length).toBe(4);
    expect(L.pages.map((p) => [...rowsOf(p), ...colsOf(p)])).toEqual([[0, 47, 0, 9], [48, 59, 0, 9], [0, 47, 10, 11], [48, 59, 10, 11]]);
  });
  it('thứ tự sang rồi xuống', () => {
    const { wb, s } = grid();
    const L = paginate(wb, s, normPrint({ ...normPrint(null), order: 'over' }));
    expect(L.pages.map((p) => [rowsOf(p)[0], colsOf(p)[0]])).toEqual([[0, 0], [0, 10], [48, 0], [48, 10]]);
  });
  it('khổ ngang: rộng hơn, ít hàng hơn', () => {
    const { wb, s } = grid();
    const L = paginate(wb, s, normPrint({ ...normPrint(null), orientation: 'landscape' }));
    expect(colsOf(L.pages[0])).toEqual([0, 11]);
    expect(rowsOf(L.pages[0])).toEqual([0, 31]);
  });
  it('Vừa tất cả cột trên một trang → 85%, 2 trang', () => {
    const { wb, s } = grid();
    const L = paginate(wb, s, normPrint({ ...normPrint(null), fit: { w: 1, h: 0 } }));
    expect(L.scale).toBeCloseTo(0.85, 5);
    expect(L.pages.length).toBe(2);
    expect(colsOf(L.pages[0])).toEqual([0, 11]);
  });
  it('Vừa trang tính trên một trang → 81%, 1 trang', () => {
    const { wb, s } = grid();
    const L = paginate(wb, s, normPrint({ ...normPrint(null), fit: { w: 1, h: 1 } }));
    expect(L.scale).toBeCloseTo(0.81, 5);
    expect(L.pages.length).toBe(1);
  });
  it('Tỷ lệ tùy chỉnh 50% → 1 trang', () => {
    const { wb, s } = grid();
    const L = paginate(wb, s, normPrint({ ...normPrint(null), scale: 50 }));
    expect(L.pages.length).toBe(1);
  });
  it('ngắt trang thủ công trước hàng 11 và cột E', () => {
    const { wb, s } = grid();
    wb.insertPageBreak(s, 10, 4);
    expect(wb.meta(s).breaks).toEqual({ rows: [10], cols: [4] });
    const L = paginate(wb, s, normPrint(null));
    expect(rowsOf(L.pages[0])).toEqual([0, 9]);
    expect(colsOf(L.pages[0])).toEqual([0, 3]);
    wb.removePageBreak(s, 10, 4);
    expect(wb.meta(s).breaks).toEqual({ rows: [], cols: [] });
    wb.undo();
    expect(wb.meta(s).breaks).toEqual({ rows: [10], cols: [4] });
    wb.resetPageBreaks(s);
    expect(wb.meta(s).breaks).toEqual({ rows: [], cols: [] });
  });
  it('hàng lặp lại ở đầu mỗi trang (1:2) và cột lặp lại bên trái (A)', () => {
    const { wb, s } = grid();
    const L = paginate(wb, s, normPrint({ ...normPrint(null), titleRowRange: [0, 1], titleColRange: [0, 0] }));
    expect(L.pages[0].titleRows).toEqual([]);
    expect(L.pages[1].titleRows).toEqual([0, 1]);
    expect(rowsOf(L.pages[1])).toEqual([48, 59]);
    // trang cột thứ hai có cột A lặp lại, bớt chỗ cho dữ liệu
    expect(L.pages[2].titleCols).toEqual([0]);
    expect(colsOf(L.pages[2])).toEqual([10, 11]);
  });
  it('bỏ qua hàng/cột ẩn, chỉ in vùng in / vùng chọn', () => {
    const { wb, s } = grid();
    wb.hideRows(s, 0, 9, true);
    let L = paginate(wb, s, normPrint(null));
    expect(rowsOf(L.pages[0])).toEqual([10, 57]);
    L = paginate(wb, s, normPrint({ ...normPrint(null), area: rg(20, 2, 30, 5) }));
    expect(L.pages.length).toBe(1);
    expect([...rowsOf(L.pages[0]), ...colsOf(L.pages[0])]).toEqual([20, 30, 2, 5]);
    L = paginate(wb, s, normPrint(null), { range: rg(40, 0, 41, 1) });
    expect(L.pages.length).toBe(1);
  });
  it('trang tính trống → 0 trang', () => {
    const wb = Workbook.create();
    expect(paginate(wb, wb.sheetIds()[0], normPrint(null)).pages.length).toBe(0);
  });
});

describe('Đầu trang / chân trang (mã Excel)', () => {
  it('tách &L &C &R và thay &P &N &D &T &A &F &&', () => {
    const p = hfParse('&LBí mật&C&P / &N&R&A - &F &&');
    expect(p).toEqual({ l: 'Bí mật', c: '&P / &N', r: '&A - &F &&' });
    const d = new Date(2026, 9, 29, 8, 5);
    const ctx = { page: 2, pages: 5, sheet: 'Dự toán', file: 'DT.xlsx', date: d };
    expect(hfRender(p.c!, ctx)).toBe('2 / 5');
    expect(hfRender(p.r!, ctx)).toBe('Dự toán - DT.xlsx &');
    expect(hfRender('&D &T', ctx)).toBe('29/10/2026 08:05');
    expect(hfRender('Trang &P', { ...ctx, page: 2 + 9 })).toBe('Trang 11');
    // bỏ mã phông/cỡ chữ: &"Arial,Bold"&14&B
    expect(hfRender('&"Arial,Bold"&14&BTiêu đề', ctx)).toBe('Tiêu đề');
    expect(hfParse('Không có mã')).toEqual({ c: 'Không có mã' });
  });
  it('số trang bắt đầu từ (First page number)', () => {
    const { wb, s } = grid();
    const html = buildPrintHtml(wb, s, { setup: { ...normPrint(null), footer: '&CTrang &P / &N', firstPage: 5 } });
    expect(html).toContain('Trang 5 / 8');
    expect(html).toContain('Trang 8 / 8');
  });
});

describe('Trang in HTML (xem trước = bản in)', () => {
  it('mỗi trang một khối đúng khổ giấy, @page lề 0, có đầu/chân trang', () => {
    const { wb, s } = grid();
    const html = buildPrintHtml(wb, s, { setup: { ...normPrint(null), header: '&CBẢNG KHỐI LƯỢNG', footer: '&RTrang &P' } });
    expect((html.match(/class="page"/g) ?? []).length).toBe(4);
    expect(html).toMatch(/@page\s*\{\s*size:\s*210mm 297mm;\s*margin:\s*0/);
    expect((html.match(/BẢNG KHỐI LƯỢNG/g) ?? []).length).toBe(4);
    expect(html).toContain('Trang 4');
    expect(html).toContain('>4711<');
  });
  it('khổ ngang, in đường lưới, in tiêu đề hàng/cột, căn giữa trang', () => {
    const { wb, s } = grid(5, 3);
    const html = buildPrintHtml(wb, s, { setup: { ...normPrint(null), orientation: 'landscape', gridlines: true, headings: true, center: { h: true, v: true } } });
    expect(html).toMatch(/size:\s*297mm 210mm/);
    expect(html).toContain('class="hd"');
    expect(html).toMatch(/<th[^>]*>B<\/th>/);
    expect(html).toContain('justify-content:center');
    expect(html).toContain('align-items:center');
  });
});

describe('Khứ hồi .xlsx thiết lập in', () => {
  it('tỷ lệ, vừa trang, lề tùy chỉnh, căn giữa, đầu/chân trang, hàng & cột lặp lại, ngắt trang, tiêu đề, số trang đầu, thứ tự, khổ A5', async () => {
    const { wb, s } = grid();
    wb.setPrint(s, { ...normPrint(null), paper: 'A5', orientation: 'landscape', scale: 75, margins: 'custom', mg: { t: 1.2, b: 1, l: 0.5, r: 0.4, h: 0.4, f: 0.35 }, center: { h: true }, header: '&LCông ty A&C&A', footer: '&CTrang &P / &N', titleRowRange: [0, 1], titleColRange: [0, 1], headings: true, gridlines: true, firstPage: 3, order: 'over' });
    wb.insertPageBreak(s, 20, 0);
    wb.insertPageBreak(s, 0, 6);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0];
    const p = normPrint(w2.meta(s2).print);
    expect(p.paper).toBe('A5');
    expect(p.orientation).toBe('landscape');
    expect(p.scale).toBe(75);
    expect(p.fit).toBeNull();
    expect(p.mg).toEqual({ t: 1.2, b: 1, l: 0.5, r: 0.4, h: 0.4, f: 0.35 });
    expect(p.center).toEqual({ h: true, v: false });
    expect(p.header).toBe('&LCông ty A&C&A');
    expect(p.footer).toBe('&CTrang &P / &N');
    expect(p.titleRowRange).toEqual([0, 1]);
    expect(p.titleColRange).toEqual([0, 1]);
    expect(p.headings).toBe(true);
    expect(p.gridlines).toBe(true);
    expect(p.firstPage).toBe(3);
    expect(p.order).toBe('over');
    expect(w2.meta(s2).breaks).toEqual({ rows: [20], cols: [6] });
  });
  it('vừa 1 trang ngang × tự động dọc', async () => {
    const { wb, s } = grid();
    wb.setPrint(s, { ...normPrint(null), fit: { w: 1, h: 0 } });
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    expect(normPrint(w2.meta(w2.sheetIds()[0]).print).fit).toEqual({ w: 1, h: 0 });
  });
});
