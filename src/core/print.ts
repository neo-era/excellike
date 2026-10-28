// Tạo trang in (HTML) cho một sheet: dùng cho In / Lưu PDF
import type { Range } from './address';
import type { PrintSetup, Workbook } from './workbook';

export const DEFAULT_PRINT: PrintSetup = { orientation: 'portrait', paper: 'A4', fitWidth: true, titleRows: 0, gridlines: false, margins: 'normal' };
const PAPER_MM: Record<string, [number, number]> = { A4: [210, 297], A3: [297, 420], Letter: [216, 279] };
const MARGIN_CM: Record<string, [number, number]> = { normal: [1.9, 1.8], narrow: [1.9, 0.6], wide: [2.5, 2.5] };
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!));

export interface PrintOpts { range?: Range; charts?: { src: string; w: number; h: number }[]; setup?: PrintSetup }

export function buildPrintHtml(wb: Workbook, s: number, opts: PrintOpts = {}): string {
  const m = wb.meta(s);
  const ps = { ...DEFAULT_PRINT, ...(opts.setup ?? m.print ?? {}) };
  const u = wb.used(s);
  let g = opts.range ?? ps.area ?? { r1: 0, c1: 0, r2: Math.max(0, u.rows - 1), c2: Math.max(0, u.cols - 1) };
  g = wb.expandToMerges(s, g);
  const cols: number[] = [], rows: number[] = [];
  for (let c = g.c1; c <= g.c2; c++) if (m.colW.get(c) > 0) cols.push(c);
  for (let r = g.r1; r <= g.r2; r++) if (m.rowH.get(r) > 0) rows.push(r);
  const tableW = cols.reduce((a, c) => a + m.colW.get(c), 0);
  // vùng in được (px ở 96 dpi)
  const [pw, ph] = PAPER_MM[ps.paper] ?? PAPER_MM.A4;
  const pageW = (ps.orientation === 'landscape' ? ph : pw) / 25.4 * 96;
  const [mv, mh] = MARGIN_CM[ps.margins ?? 'normal'];
  const printable = pageW - (2 * mh / 2.54) * 96;
  const zoom = ps.fitWidth && tableW > printable ? Math.max(0.1, printable / tableW) : 1;
  const colSet = new Set(cols), rowSet = new Set(rows);
  const css = (st: ReturnType<Workbook['style']>, cf: ReturnType<Workbook['cfStyle']>, align: string) => {
    const o: string[] = [];
    const bg = cf.bg ?? st.bg, fg = cf.fg ?? st.fg;
    if (st.b || cf.b) o.push('font-weight:700');
    if (st.i || cf.i) o.push('font-style:italic');
    const deco = [st.u ? 'underline' : '', st.s ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) o.push('text-decoration:' + deco);
    if (fg) o.push('color:' + fg);
    if (bg) o.push('background:' + bg);
    if (st.fn) { const fn = st.fn.replace(/[^\p{L}\p{N} _.-]/gu, '').trim(); if (fn) o.push(`font-family:'${fn}',Arial,sans-serif`); }
    if (st.fs) o.push(`font-size:${st.fs}pt`);
    o.push('text-align:' + ({ l: 'left', c: 'center', r: 'right' }[align] ?? 'left'));
    o.push('vertical-align:' + (({ t: 'top', m: 'middle' } as Record<string, string>)[st.va ?? ''] ?? 'bottom'));
    if (st.wr) o.push('white-space:pre-wrap');
    const side = (k: 't' | 'r' | 'b' | 'l', name: string) => {
      const b = st.bd?.[k];
      if (!b) return;
      const w = b.s === 'thick' ? 3 : b.s === 'medium' ? 2 : b.s === 'double' ? 3 : 1;
      const style = b.s === 'dashed' ? 'dashed' : b.s === 'dotted' ? 'dotted' : b.s === 'double' ? 'double' : 'solid';
      o.push(`border-${name}:${w}px ${style} ${b.c ?? '#000'}`);
    };
    side('t', 'top'); side('r', 'right'); side('b', 'bottom'); side('l', 'left');
    return o.join(';');
  };
  const rowHtml = (r: number) => {
    let h = `<tr style="height:${m.rowH.get(r)}px">`;
    for (const c of cols) {
      const mg = wb.mergeAt(s, r, c);
      if (mg && (mg.r1 !== r || mg.c1 !== c)) {
        // ô bị gộp: chỉ in ở ô đầu nhìn thấy được
        const firstR = rows.find((x) => x >= mg.r1 && x <= mg.r2), firstC = cols.find((x) => x >= mg.c1 && x <= mg.c2);
        if (!(firstR === r && firstC === c && (!rowSet.has(mg.r1) || !colSet.has(mg.c1)))) continue;
      }
      const d = wb.display(s, mg ? mg.r1 : r, mg ? mg.c1 : c);
      const cf = wb.cfStyle(s, mg ? mg.r1 : r, mg ? mg.c1 : c);
      let span = '';
      if (mg) {
        const rs = rows.filter((x) => x >= mg.r1 && x <= mg.r2).length, cs = cols.filter((x) => x >= mg.c1 && x <= mg.c2).length;
        if (rs > 1) span += ` rowspan="${rs}"`;
        if (cs > 1) span += ` colspan="${cs}"`;
      }
      h += `<td${span} style="${esc(css(d.style, cf, d.align))}">${esc(d.text)}</td>`;
    }
    return h + '</tr>';
  };
  // hàng tiêu đề lặp lại: mở rộng để không cắt ngang ô gộp
  let tEnd = ps.titleRows > 0 && g.r1 === 0 ? Math.min(ps.titleRows, g.r2 + 1) : 0;
  for (let grow = true; grow && tEnd > 0;) {
    grow = false;
    for (const mg of m.merges) if (mg.r1 < tEnd && mg.r2 >= tEnd && mg.c2 >= g.c1 && mg.c1 <= g.c2) { tEnd = mg.r2 + 1; grow = true; }
  }
  const nTitle = rows.filter((r) => r < tEnd);
  const body = rows.filter((r) => !nTitle.includes(r));
  const charts = (opts.charts ?? []).map((c) => {
    const w = Math.min(c.w, printable / zoom);
    return `<div class="chart"><img src="${c.src}" style="width:${w}px;height:${(c.h * w) / c.w}px" alt="Biểu đồ"></div>`;
  }).join('');
  return `<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; font-src data:"><title>${esc(wb.sheetName(s))}</title>
<style>
@page { size: ${ps.paper} ${ps.orientation}; margin: ${mv}cm ${mh}cm; }
* { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { margin: 0; font: 10pt 'Be Vietnam Pro', Arial, sans-serif; color: #000; background: #fff; }
.wrap { zoom: ${zoom.toFixed(4)}; }
table { border-collapse: collapse; table-layout: fixed; width: ${tableW}px; }
td { padding: 1px 4px; overflow: hidden; white-space: nowrap; ${ps.gridlines ? 'border: 1px solid #c8c8c8;' : ''} font-variant-numeric: tabular-nums; }
thead { display: table-header-group; }
tr { break-inside: avoid; }
.chart { break-inside: avoid; margin-top: 16px; }
.chart img { display: block; }
</style></head><body><div class="wrap">
<table><colgroup>${cols.map((c) => `<col style="width:${m.colW.get(c)}px">`).join('')}</colgroup>
${nTitle.length ? `<thead>${nTitle.map(rowHtml).join('')}</thead>` : ''}<tbody>${body.map(rowHtml).join('')}</tbody></table>
${charts}</div></body></html>`;
}
