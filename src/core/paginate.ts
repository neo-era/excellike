// Chia trang in như Excel 365: khổ giấy, lề (inch), co giãn % hoặc vừa X×Y trang, ngắt trang thủ công,
// hàng/cột lặp lại, thứ tự trang (xuống rồi sang / sang rồi xuống). Xem trước và bản in dùng chung kết quả này.
import { MAX_COLS, MAX_ROWS, type Range } from './address';
import type { Margins, PaperName, PrintSetup, Workbook } from './workbook';

export const PX_IN = 96;
/** Khổ giấy dọc (mm) và mã khổ giấy của Excel (pageSetup paperSize) */
export const PAPERS: Record<PaperName, { mm: [number, number]; code: number; label: string }> = {
  Letter: { mm: [215.9, 279.4], code: 1, label: 'Letter 21,59 cm × 27,94 cm' },
  Tabloid: { mm: [279.4, 431.8], code: 3, label: 'Tabloid 27,94 cm × 43,18 cm' },
  Legal: { mm: [215.9, 355.6], code: 5, label: 'Legal 21,59 cm × 35,56 cm' },
  A3: { mm: [297, 420], code: 8, label: 'A3 29,7 cm × 42 cm' },
  A4: { mm: [210, 297], code: 9, label: 'A4 21 cm × 29,7 cm' },
  A5: { mm: [148, 210], code: 11, label: 'A5 14,8 cm × 21 cm' },
  B4: { mm: [257, 364], code: 12, label: 'B4 (JIS) 25,7 cm × 36,4 cm' },
  B5: { mm: [182, 257], code: 13, label: 'B5 (JIS) 18,2 cm × 25,7 cm' },
};
export const paperByCode = (n: number | undefined): PaperName => (Object.entries(PAPERS).find(([, p]) => p.code === n)?.[0] as PaperName) ?? 'A4';
/** Lề có sẵn của Excel (inch): Thường, Rộng, Hẹp */
export const MARGIN_PRESETS: Record<'normal' | 'wide' | 'narrow', Margins> = {
  normal: { t: 0.75, b: 0.75, l: 0.7, r: 0.7, h: 0.3, f: 0.3 },
  wide: { t: 1, b: 1, l: 1, r: 1, h: 0.5, f: 0.5 },
  narrow: { t: 0.75, b: 0.75, l: 0.25, r: 0.25, h: 0.3, f: 0.3 },
};

/** Thiết lập in đầy đủ (đã điền mặc định) */
export interface NPrint extends PrintSetup {
  margins: 'normal' | 'wide' | 'narrow' | 'custom';
  mg: Margins;
  scale: number;
  fit: { w: number; h: number } | null;
  titleRowRange: [number, number] | null;
  titleColRange: [number, number] | null;
  headings: boolean;
  center: { h: boolean; v: boolean };
  firstPage: number | null;
  order: 'down' | 'over';
  header: string;
  footer: string;
  area: Range | null;
}

export function normPrint(ps: PrintSetup | null | undefined): NPrint {
  const p: Partial<PrintSetup> = ps ?? {};
  const preset = p.margins ?? 'normal';
  const mg = p.mg ?? MARGIN_PRESETS[preset === 'custom' ? 'normal' : preset];
  const fit = p.fit !== undefined ? (p.fit && (p.fit.w > 0 || p.fit.h > 0) ? { w: Math.max(0, p.fit.w | 0), h: Math.max(0, p.fit.h | 0) } : null) : p.fitWidth ? { w: 1, h: 0 } : null;
  const tr = p.titleRowRange !== undefined ? p.titleRowRange : p.titleRows ? [0, p.titleRows - 1] as [number, number] : null;
  return {
    orientation: p.orientation === 'landscape' ? 'landscape' : 'portrait',
    paper: p.paper && PAPERS[p.paper] ? p.paper : 'A4',
    gridlines: !!p.gridlines,
    margins: preset,
    mg: { ...mg },
    area: p.area ?? null,
    scale: Math.max(10, Math.min(400, Math.round(p.scale ?? 100))),
    fit,
    titleRowRange: tr ?? null,
    titleColRange: p.titleColRange ?? null,
    headings: !!p.headings,
    center: { h: !!p.center?.h, v: !!p.center?.v },
    firstPage: p.firstPage ?? null,
    order: p.order === 'over' ? 'over' : 'down',
    header: p.header ?? '',
    footer: p.footer ?? '',
  };
}
/** Bản gọn để lưu (bỏ trường cũ fitWidth/titleRows) */
export function storePrint(p: NPrint): PrintSetup {
  const o: any = { ...p };
  delete o.fitWidth; delete o.titleRows;
  return o;
}

export interface PageSpec { index: number; rows: number[]; cols: number[]; titleRows: number[]; titleCols: number[] }
export interface Layout {
  pages: PageSpec[];
  scale: number;
  paper: { w: number; h: number; wmm: number; hmm: number };
  avail: { w: number; h: number };
  area: Range | null;
  setup: NPrint;
}

export function paperSize(p: NPrint) {
  const [a, b] = PAPERS[p.paper].mm;
  const [wmm, hmm] = p.orientation === 'landscape' ? [b, a] : [a, b];
  return { wmm, hmm, w: (wmm / 25.4) * PX_IN, h: (hmm / 25.4) * PX_IN };
}

/** Vùng có dữ liệu cần in (gồm cả biểu đồ) */
export function printRange(wb: Workbook, s: number): Range | null {
  const u = wb.used(s), m = wb.meta(s);
  let r2 = u.rows - 1, c2 = u.cols - 1;
  for (const ch of m.charts) {
    let c = ch.anchor.c, x = ch.anchor.dx + ch.w;
    while (c < 16383 && x > m.colW.get(c)) { x -= m.colW.get(c); c++; }
    let r = ch.anchor.r, y = ch.anchor.dy + ch.h;
    while (r < 1048575 && y > m.rowH.get(r)) { y -= m.rowH.get(r); r++; }
    r2 = Math.max(r2, r); c2 = Math.max(c2, c);
  }
  if (r2 < 0 || c2 < 0) return null;
  return { r1: 0, c1: 0, r2, c2 };
}

function split(items: number[], size: (i: number) => number, cap: (first: number) => number, manual: Set<number>): number[][] {
  const out: number[][] = [];
  let cur: number[] = [], used = 0, lim = 0;
  for (const it of items) {
    const sz = size(it);
    if (cur.length && (manual.has(it) || used + sz > lim + 0.01)) { out.push(cur); cur = []; }
    if (!cur.length) { used = 0; lim = cap(it); }
    cur.push(it); used += sz;
  }
  if (cur.length) out.push(cur);
  return out;
}

export function paginate(wb: Workbook, s: number, setup: NPrint, opts: { range?: Range } = {}): Layout {
  const ps = normPrint(setup);
  const m = wb.meta(s);
  const paper = paperSize(ps);
  const avail = { w: paper.w - (ps.mg.l + ps.mg.r) * PX_IN, h: paper.h - (ps.mg.t + ps.mg.b) * PX_IN };
  const base = { paper, avail, setup: ps };
  let g = opts.range ?? ps.area ?? printRange(wb, s);
  if (!g) return { ...base, pages: [], scale: ps.scale / 100, area: null };
  // chọn cả cột / cả hàng / cả trang tính → chỉ in phần có dữ liệu (như Excel)
  if (g.r2 >= MAX_ROWS - 1 || g.c2 >= MAX_COLS - 1) {
    const pr = printRange(wb, s);
    if (!pr) return { ...base, pages: [], scale: ps.scale / 100, area: null };
    g = { ...g, r2: g.r2 >= MAX_ROWS - 1 ? pr.r2 : g.r2, c2: g.c2 >= MAX_COLS - 1 ? pr.c2 : g.c2 };
    if (g.r2 < g.r1 || g.c2 < g.c1) return { ...base, pages: [], scale: ps.scale / 100, area: null };
  }
  g = wb.expandToMerges(s, g);
  const rows: number[] = [], cols: number[] = [];
  for (let r = g.r1; r <= g.r2; r++) if (m.rowH.get(r) > 0) rows.push(r);
  for (let c = g.c1; c <= g.c2; c++) if (m.colW.get(c) > 0) cols.push(c);
  if (!rows.length || !cols.length) return { ...base, pages: [], scale: ps.scale / 100, area: g };
  const tR: number[] = [], tC: number[] = [];
  // hàng/cột lặp lại không cắt ngang ô gộp
  const ext = (a: [number, number] | null, ax: 'r' | 'c'): [number, number] | null => {
    if (!a) return null;
    let [lo, hi] = a;
    for (let grow = true; grow;) {
      grow = false;
      for (const mg of m.merges) {
        const [m1, m2] = ax === 'r' ? [mg.r1, mg.r2] : [mg.c1, mg.c2];
        if (m1 <= hi && m2 > hi && m1 >= lo) { hi = m2; grow = true; }
      }
    }
    return [lo, hi];
  };
  ps.titleRowRange = ext(ps.titleRowRange, 'r');
  ps.titleColRange = ext(ps.titleColRange, 'c');
  if (ps.titleRowRange) for (let r = ps.titleRowRange[0]; r <= ps.titleRowRange[1]; r++) if (m.rowH.get(r) > 0) tR.push(r);
  if (ps.titleColRange) for (let c = ps.titleColRange[0]; c <= ps.titleColRange[1]; c++) if (m.colW.get(c) > 0) tC.push(c);
  const tRH = tR.reduce((a, r) => a + m.rowH.get(r), 0), tCW = tC.reduce((a, c) => a + m.colW.get(c), 0);
  const tREnd = ps.titleRowRange?.[1] ?? -1, tCEnd = ps.titleColRange?.[1] ?? -1;
  // tiêu đề hàng/cột (A, B, 1, 2) khi in chiếm thêm chỗ
  const hdW = ps.headings ? 30 + 7 * String(g.r2 + 1).length : 0, hdH = ps.headings ? 20 : 0;
  // như Excel: chế độ “Vừa X × Y trang” bỏ qua ngắt trang thủ công
  const brR = new Set(ps.fit ? [] : m.breaks?.rows ?? []), brC = new Set(ps.fit ? [] : m.breaks?.cols ?? []);
  const run = (sc: number) => {
    const rp = split(rows, (r) => m.rowH.get(r), (first) => avail.h / sc - hdH - (first > tREnd && tR.length ? tRH : 0), brR);
    const cp = split(cols, (c) => m.colW.get(c), (first) => avail.w / sc - hdW - (first > tCEnd && tC.length ? tCW : 0), brC);
    return { rp, cp };
  };
  let scale = ps.scale / 100;
  if (ps.fit) {
    const ok = (p: number) => { const { rp, cp } = run(p / 100); return (!ps.fit!.w || cp.length <= ps.fit!.w) && (!ps.fit!.h || rp.length <= ps.fit!.h); };
    let lo = 10, hi = 100;
    if (ok(100)) lo = 100;
    else { while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (ok(mid)) lo = mid; else hi = mid - 1; } }
    scale = lo / 100;
  }
  const { rp, cp } = run(scale);
  const pages: PageSpec[] = [];
  const mk = (rr: number[], cc: number[]) => pages.push({ index: pages.length, rows: rr, cols: cc, titleRows: rr[0] > tREnd ? tR : [], titleCols: cc[0] > tCEnd ? tC : [] });
  if (ps.order === 'over') { for (const rr of rp) for (const cc of cp) mk(rr, cc); }
  else { for (const cc of cp) for (const rr of rp) mk(rr, cc); }
  return { ...base, pages, scale, area: g };
}

// ---------------- đầu trang / chân trang ----------------
export interface HFParts { l?: string; c?: string; r?: string }
/** Tách chuỗi đầu/chân trang Excel (&L…&C…&R…) thành 3 phần; giữ nguyên các mã khác */
export function hfParse(s: string | undefined | null): HFParts {
  const out: Record<'l' | 'c' | 'r', string> = { l: '', c: '', r: '' };
  let sec: 'l' | 'c' | 'r' = 'c';
  const t = s ?? '';
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (ch === '&' && i + 1 < t.length) {
      const n = t[i + 1];
      if (n === 'L' || n === 'C' || n === 'R') { sec = n.toLowerCase() as 'l' | 'c' | 'r'; i++; continue; }
      out[sec] += ch + n; i++; continue;
    }
    out[sec] += ch;
  }
  const r: HFParts = {};
  for (const k of ['l', 'c', 'r'] as const) if (out[k]) r[k] = out[k];
  return r;
}
export function hfJoin(p: HFParts) { return (p.l ? '&L' + p.l : '') + (p.c ? '&C' + p.c : '') + (p.r ? '&R' + p.r : ''); }
export interface HFCtx { page: number; pages: number; sheet: string; file: string; date: Date }
const p2 = (n: number) => String(n).padStart(2, '0');
/** Thay mã &P &N &D &T &A &F &Z &&; bỏ mã định dạng (&"phông", &cỡ, &B &I &U &S &X &Y &E &K) */
export function hfRender(t: string, x: HFCtx): string {
  let o = '';
  for (let i = 0; i < t.length; i++) {
    if (t[i] !== '&') { o += t[i]; continue; }
    const n = t[i + 1];
    if (n === undefined) { o += '&'; break; }
    i++;
    switch (n) {
      case '&': o += '&'; break;
      case 'P': o += String(x.page); break;
      case 'N': o += String(x.pages); break;
      case 'D': o += `${p2(x.date.getDate())}/${p2(x.date.getMonth() + 1)}/${x.date.getFullYear()}`; break;
      case 'T': o += `${p2(x.date.getHours())}:${p2(x.date.getMinutes())}`; break;
      case 'A': o += x.sheet; break;
      case 'F': o += x.file; break;
      case 'Z': case 'G': break;
      case '"': { const e = t.indexOf('"', i + 1); i = e < 0 ? t.length : e; break; }
      case 'K': i += 6; break;
      default:
        if (/\d/.test(n)) { let k = 1; while (k < 3 && /\d/.test(t[i + 1] ?? '')) { i++; k++; } }
        // &B &I &U &E &S &X &Y và mã lạ: bỏ qua
    }
  }
  return o;
}
