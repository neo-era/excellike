// Biểu đồ: mô hình, lấy dữ liệu từ ô, thang trục, số rút gọn
import type { Range } from './address';
import type { Workbook } from './workbook';

export type ChartType = 'col' | 'colStack' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut';
export interface Ref { sheet: number; range: Range }
export interface Series { name?: Ref; cats?: Ref; vals: Ref }
export interface Chart {
  id?: string;
  type: ChartType;
  title?: string;
  series: Series[];
  anchor: { r: number; c: number; dx: number; dy: number }; // góc trên trái: ô + lệch (px ở 100%)
  w: number; h: number;                                     // kích thước px ở 100%
  legend?: 'b' | 'r' | 't' | 'none';
  labels?: boolean;
}
export const CHART_NAMES: Record<ChartType, string> = { col: 'Cột', colStack: 'Cột chồng', bar: 'Thanh ngang', line: 'Đường', area: 'Vùng', pie: 'Tròn', doughnut: 'Vành khuyên' };
export const PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e', '#636363', '#997300'];

/** Thang trục "đẹp": bước 1/2/5 × 10^n */
export function niceScale(min: number, max: number, ticks = 5): { lo: number; hi: number; step: number } {
  if (min > 0) min = 0;
  if (max < 0) max = 0;
  if (max === min) return { lo: min, hi: min + 1, step: 0.2 };
  const raw = (max - min) / ticks;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 5, 10].map((m) => m * mag).find((s) => s >= raw - 1e-12)!;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  return { lo: clean(lo), hi: clean(hi), step: clean(step) };
}
const clean = (x: number) => +x.toPrecision(12) || 0;

/** Số rút gọn cho trục/nhãn: 1.250 · 1,5 tr · 2 tỷ */
export function compactNum(n: number): string {
  const a = Math.abs(n), s = n < 0 ? '-' : '';
  const fmt = (x: number, d: number) => (+x.toFixed(d)).toLocaleString('vi-VN', { maximumFractionDigits: d });
  if (a >= 1e9) return s + fmt(a / 1e9, 2) + ' tỷ';
  if (a >= 1e6) return s + fmt(a / 1e6, 2) + ' tr';
  if (a >= 1000) return s + fmt(a, 0);
  return s + fmt(a, a < 1 ? 3 : 2);
}

const isNumLike = (v: unknown) => typeof v === 'number';
const isTextLike = (v: unknown) => v === null || v === '' || typeof v === 'string';

/** Từ vùng chọn → danh sách chuỗi: tự nhận hàng tiêu đề và cột nhãn */
export function seriesFromRange(wb: Workbook, sheet: number, g: Range, byRow?: boolean): Series[] {
  const u = wb.used(sheet);
  g = { ...g, r2: Math.min(g.r2, Math.max(g.r1, u.rows - 1)), c2: Math.min(g.c2, Math.max(g.c1, u.cols - 1)) };
  const v = (r: number, c: number) => wb.value(sheet, r, c);
  const nR = g.r2 - g.r1 + 1, nC = g.c2 - g.c1 + 1;
  // hàng đầu là tiêu đề nếu toàn chữ (bỏ qua ô góc) và bên dưới có số
  let hdrRow = false, lblCol = false;
  if (nR > 1) { hdrRow = true; for (let c = g.c1; c <= g.c2; c++) if (!isTextLike(v(g.r1, c))) hdrRow = false; }
  if (nC > 1) { lblCol = true; for (let r = g.r1 + (hdrRow ? 1 : 0); r <= g.r2; r++) if (!isTextLike(v(r, g.c1))) lblCol = false; }
  if (hdrRow && !lblCol && nC > 1) {
    // góc trên trái trống + cột đầu là chữ ở hàng dưới → cột nhãn
    let allText = true;
    for (let r = g.r1 + 1; r <= g.r2; r++) if (isNumLike(v(r, g.c1))) allText = false;
    if (allText) lblCol = true;
  }
  // nhiều cột chữ liền nhau ở đầu (vd. Nhóm, Hạng mục) → cột chữ sát số nhất làm nhãn
  let lc = g.c1;
  if (lblCol) {
    while (lc + 1 < g.c2) {
      let allText = true, any = false;
      for (let r = g.r1 + (hdrRow ? 1 : 0); r <= g.r2; r++) { const x = v(r, lc + 1); if (!isTextLike(x)) allText = false; else if (x !== null && x !== '') any = true; }
      if (!allText || !any) break;
      lc++;
    }
  }
  const d = { r1: g.r1 + (hdrRow ? 1 : 0), c1: lblCol ? lc + 1 : g.c1, r2: g.r2, c2: g.c2 };
  // như Excel: nhiều cột hơn hàng thì mỗi hàng là một chuỗi
  if (byRow === undefined) byRow = d.c2 - d.c1 > d.r2 - d.r1;
  const out: Series[] = [];
  const R = (range: Range): Ref => ({ sheet, range });
  if (!byRow) {
    for (let c = d.c1; c <= d.c2; c++) out.push({
      name: hdrRow ? R({ r1: g.r1, c1: c, r2: g.r1, c2: c }) : undefined,
      cats: lblCol ? R({ r1: d.r1, c1: lc, r2: d.r2, c2: lc }) : undefined,
      vals: R({ r1: d.r1, c1: c, r2: d.r2, c2: c }),
    });
  } else {
    for (let r = d.r1; r <= d.r2; r++) out.push({
      name: lblCol ? R({ r1: r, c1: lc, r2: r, c2: lc }) : undefined,
      cats: hdrRow ? R({ r1: g.r1, c1: d.c1, r2: g.r1, c2: d.c2 }) : undefined,
      vals: R({ r1: r, c1: d.c1, r2: r, c2: d.c2 }),
    });
  }
  const hasNum = (x: Ref) => { for (let r = x.range.r1; r <= x.range.r2; r++) for (let c = x.range.c1; c <= x.range.c2; c++) if (isNumLike(v(r, c))) return true; return false; };
  return out.filter((s) => s.vals.range.r1 <= s.vals.range.r2 && s.vals.range.c1 <= s.vals.range.c2 && hasNum(s.vals));
}

/** Vùng bao các tham chiếu của biểu đồ (để hiển thị trong hộp thoại) */
export function chartSourceRange(ch: Pick<Chart, 'series'>): Ref | null {
  const refs = ch.series.flatMap((s) => [s.name, s.cats, s.vals]).filter(Boolean) as Ref[];
  if (!refs.length) return null;
  const sheet = refs[0].sheet;
  const rs = refs.filter((x) => x.sheet === sheet).map((x) => x.range);
  return { sheet, range: { r1: Math.min(...rs.map((x) => x.r1)), c1: Math.min(...rs.map((x) => x.c1)), r2: Math.max(...rs.map((x) => x.r2)), c2: Math.max(...rs.map((x) => x.c2)) } };
}

/** Giá trị các ô của một tham chiếu (all = gồm cả hàng ẩn, dùng khi ghi tệp) */
export function cellsOf(wb: Workbook, ref: Ref | undefined, kind: 'text' | 'num', all = false): (string | number | null)[] {
  if (!ref) return [];
  if (!wb.sheetIds().includes(ref.sheet)) return [];
  const g = ref.range, out: (string | number | null)[] = [];
  const n = Math.min((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1), 5000);
  for (let r = g.r1; r <= g.r2 && out.length < n; r++) for (let c = g.c1; c <= g.c2 && out.length < n; c++) {
    if (!all && wb.isRowHidden(ref.sheet, r)) continue;
    if (kind === 'text') out.push(wb.display(ref.sheet, r, c).text);
    else { const v = wb.value(ref.sheet, r, c); out.push(typeof v === 'number' ? v : null); }
  }
  return out;
}

export interface ChartData { cats: string[]; series: { name: string; values: (number | null)[] }[] }
export function chartData(wb: Workbook, ch: Pick<Chart, 'type' | 'series'>): ChartData {
  const series = ch.series.map((s, i) => ({
    name: (s.name ? (cellsOf(wb, s.name, 'text')[0] as string) : '') || `Chuỗi ${i + 1}`,
    values: cellsOf(wb, s.vals, 'num') as (number | null)[],
  }));
  const n = Math.max(0, ...series.map((s) => s.values.length));
  let cats = ch.series[0]?.cats ? (cellsOf(wb, ch.series[0].cats, 'text') as string[]) : [];
  if (cats.length < n) cats = [...cats, ...Array.from({ length: n - cats.length }, (_, i) => String(cats.length + i + 1))];
  return { cats: cats.slice(0, n), series };
}
