// Tạo trang in (HTML) đã chia trang sẵn: mỗi trang giấy là một khối đúng khổ, có đầu trang/chân trang.
// Bản xem trước trong Backstage và bản in/PDF dùng cùng một HTML nên luôn khớp nhau.
import type { Range } from './address';
import type { PrintSetup, Workbook } from './workbook';
import { hfParse, hfRender, normPrint, paginate, PX_IN, type Layout, type NPrint, type PageSpec } from './paginate';

/** Giữ để tương thích: thiết lập mặc định (Excel: Không co giãn) */
export const DEFAULT_PRINT: PrintSetup = normPrint(null);
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!));
const FONT = "'Aptos Narrow', Calibri, 'Segoe UI', Arial, sans-serif";

export interface ChartImg { src: string; w: number; h: number; anchor: { r: number; c: number; dx: number; dy: number } }
export interface PrintOpts { range?: Range; charts?: ChartImg[]; setup?: PrintSetup; file?: string; date?: Date; pages?: [number, number] }
export interface JobItem { sheet: number; setup?: PrintSetup; range?: Range; charts?: ChartImg[] }
export interface Job { html: string; total: number; layouts: Layout[] }

/** Một sheet → HTML (dùng cho kiểm thử và in nhanh) */
export function buildPrintHtml(wb: Workbook, s: number, opts: PrintOpts = {}): string {
  return buildJob(wb, [{ sheet: s, setup: opts.setup, range: opts.range, charts: opts.charts }], opts).html;
}

/** Nhiều sheet (In Toàn bộ Sổ làm việc): đánh số trang nối tiếp như Excel */
export function buildJob(wb: Workbook, items: JobItem[], opts: { file?: string; date?: Date; pages?: [number, number] } = {}): Job {
  const date = opts.date ?? new Date();
  const layouts = items.map((it) => paginate(wb, it.sheet, normPrint(it.setup ?? wb.meta(it.sheet).print), { range: it.range }));
  // số trang: mỗi sheet có “Số trang đầu tiên” thì đánh lại từ đó, không thì nối tiếp
  const nums: number[][] = [];
  let n = 0;
  layouts.forEach((L) => { if (L.setup.firstPage) n = L.setup.firstPage - 1; nums.push(L.pages.map(() => ++n)); });
  const last = n;
  const total = layouts.reduce((a, L) => a + L.pages.length, 0);
  const from = opts.pages ? Math.max(1, opts.pages[0]) : 1, to = opts.pages ? Math.min(total, opts.pages[1]) : total;
  let seq = 0, size = '210mm 297mm';
  // khổ/hướng khác nhau giữa các sheet (In Toàn bộ Sổ làm việc): mỗi khổ một @page có tên
  const sizes = new Map<string, string>();
  const out: string[] = [];
  layouts.forEach((L, i) => {
    const s = items[i].sheet;
    L.pages.forEach((p, j) => {
      seq++;
      if (seq < from || seq > to) return;
      const sz = `${fmt(L.paper.wmm)}mm ${fmt(L.paper.hmm)}mm`;
      if (!out.length) size = sz;
      if (!sizes.has(sz)) sizes.set(sz, 'pg' + sizes.size);
 out.push(pageHtml(sizes.get(sz)!, wb, s, L, p, { page: nums[i][j], pages: last, sheet: wb.sheetName(s), file: opts.file ?? wb.name, date }, items[i].charts ?? []));
    });
  });
  const title = items.length === 1 ? wb.sheetName(items[0].sheet) : wb.name;
  const html = `<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; font-src data:"><title>${esc(title)}</title>
<style>
@page { size: ${size}; margin: 0; }
${[...sizes].map(([sz, n]) => `@page ${n} { size: ${sz}; margin: 0; }`).join('\n')}
* { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
html, body { margin: 0; padding: 0; background: #fff; color: #000; }
body { font: 11pt ${FONT}; }
.page { position: relative; overflow: hidden; background: #fff; break-after: page; page-break-after: always; }
.page:last-child { break-after: auto; page-break-after: auto; }
.hf { position: absolute; display: grid; grid-template-columns: 1fr 1fr 1fr; font: 11pt ${FONT}; white-space: pre; line-height: 1.2; }
.hf .l { text-align: left; } .hf .c { text-align: center; } .hf .r { text-align: right; }
.bd { position: absolute; overflow: hidden; display: flex; }
.sc { position: relative; flex: none; }
table { border-collapse: collapse; table-layout: fixed; }
td { padding: 0 2px; overflow: hidden; white-space: nowrap; font-variant-numeric: tabular-nums; vertical-align: bottom; line-height: 1.15; }
td.ov { overflow: visible; }
.grid td { border: 1px solid #c6c6c6; }
th.hd { background: #f2f2f2; border: 1px solid #a6a6a6; font: 9pt ${FONT}; font-weight: 400; text-align: center; padding: 0; }
.chart { position: absolute; }
.chart img { display: block; width: 100%; height: 100%; }
.mg { display: none; position: absolute; inset: 0; pointer-events: none; }
@media screen {
  html { background: #e6e6e6; }
  body { background: transparent; padding: 16px 0; }
  .page { margin: 0 auto 16px; box-shadow: 0 1px 4px rgba(0,0,0,.25); }
  body.one .page { display: none; margin: 0 auto; }
  body.one .page.cur { display: block; }
  body.show-margins .mg { display: block; }
  .mg i { position: absolute; border: 0 dashed #7a7a7a; }
}
</style></head><body>${out.join('\n')}</body></html>`;
  return { html, total, layouts };
}
const fmt = (x: number) => String(Math.round(x * 100) / 100);

function cellCss(st: ReturnType<Workbook['style']>, cf: ReturnType<Workbook['cfStyle']>, align: string) {
  const o: string[] = [];
  const bg = cf.bg ?? st.bg, fg = cf.fg ?? st.fg;
  if (st.b || cf.b) o.push('font-weight:700');
  if (st.i || cf.i) o.push('font-style:italic');
  const deco = [st.u ? 'underline' : '', st.s ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) o.push('text-decoration:' + deco + (st.ud ? ' double' : ''));
  if (fg) o.push('color:' + fg);
  if (bg) o.push('background:' + bg);
  if (st.fn) { const fn = st.fn.replace(/[^\p{L}\p{N} _.-]/gu, '').trim(); if (fn) o.push(`font-family:'${fn}',${FONT}`); }
  if (st.fs) o.push(`font-size:${st.fs}pt`);
  if (st.ind) o.push(`padding-${align === 'r' ? 'right' : 'left'}:${3 + st.ind * 9}px`);
  const ta = st.hx === 'just' ? 'justify' : st.hx === 'dist' || st.hx === 'cas' ? 'center' : ({ l: 'left', c: 'center', r: 'right' }[align] ?? 'left');
  o.push('text-align:' + ta);
  if (st.hx === 'just' || st.hx === 'dist') o.push('white-space:normal');
  if (st.rot === 255) o.push('writing-mode:vertical-rl;text-orientation:upright');
  const diag: string[] = [];
  if (st.bd?.d) diag.push(`linear-gradient(to top right,transparent calc(50% - 0.6px),${st.bd.d.c ?? '#000'} 50%,transparent calc(50% + 0.6px))`);
  if (st.bd?.u) diag.push(`linear-gradient(to bottom right,transparent calc(50% - 0.6px),${st.bd.u.c ?? '#000'} 50%,transparent calc(50% + 0.6px))`);
  if (diag.length) o.push('background-image:' + diag.join(','));
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
}

function pageHtml(named: string, wb: Workbook, s: number, L: Layout, p: PageSpec, ctx: { page: number; pages: number; sheet: string; file: string; date: Date }, charts: ChartImg[]) {
  const ps: NPrint = L.setup, m = wb.meta(s);
  const rows = [...p.titleRows, ...p.rows], cols = [...p.titleCols, ...p.cols];
  const rowSet = new Set(rows), colSet = new Set(cols);
  const colName = (c: number) => { let n = '', x = c + 1; while (x > 0) { const k = (x - 1) % 26; n = String.fromCharCode(65 + k) + n; x = Math.floor((x - 1) / 26); } return n; };
  const hdW = ps.headings ? 30 + 7 * String((L.area?.r2 ?? 0) + 1).length : 0;
  const td = (r: number, c: number, i: number) => {
    const mg = wb.mergeAt(s, r, c);
    if (mg && (mg.r1 !== r || mg.c1 !== c)) {
      // ô bị gộp: chỉ in ở ô đầu tiên nhìn thấy trên trang này
      const fr = rows.find((x) => x >= mg.r1 && x <= mg.r2), fc = cols.find((x) => x >= mg.c1 && x <= mg.c2);
      if (!(fr === r && fc === c && (!rowSet.has(mg.r1) || !colSet.has(mg.c1)))) return '';
    }
    const ar = mg ? mg.r1 : r, ac = mg ? mg.c1 : c;
    const d = wb.display(s, ar, ac);
    const cf = wb.cfStyle(s, ar, ac);
    let attr = '';
    if (mg) {
      const rs = rows.filter((x) => x >= mg.r1 && x <= mg.r2).length, cs = cols.filter((x) => x >= mg.c1 && x <= mg.c2).length;
      if (rs > 1) attr += ` rowspan="${rs}"`;
      if (cs > 1) attr += ` colspan="${cs}"`;
    }
    // chữ tràn sang ô trống bên phải như Excel
    const nc = cols[i + 1];
    if (!mg && d.text && typeof d.value !== 'number' && !d.style.wr && d.align !== 'r' && d.align !== 'c' && nc !== undefined && wb.display(s, r, nc).text === '' && !wb.mergeAt(s, r, nc)) attr += ' class="ov"';
    // số không vừa ô: Excel in ##### (ước lượng độ rộng chữ số Aptos/Calibri 11 pt ≈ 7 px)
    let text = d.text;
    if (!mg && typeof d.value === 'number' && text) {
      const k = (d.style.fs ?? 11) / 11, w = m.colW.get(c) - 5;
      const est = [...text].reduce((a, ch) => a + (/[.,:\s]/.test(ch) ? 3.5 : 7), 0) * k * (d.style.b ? 1.06 : 1);
      if (est > w) text = '#'.repeat(Math.max(1, Math.floor(w / (7 * k))));
    }
    const st = d.style;
    let inner = esc(text);
    if (st.vp) inner = `<${st.vp}>${inner}</${st.vp}>`;
    if (st.rot && st.rot !== 255) inner = `<span style="display:inline-block;transform:rotate(${-st.rot}deg);white-space:nowrap">${inner}</span>`;
    if (st.shk && text) inner = `<span style="display:inline-block;max-width:100%;font-size:${Math.max(0.3, Math.min(1, (m.colW.get(c) - 6) / Math.max(1, [...text].length * 7 * ((st.fs ?? 11) / 11))))}em">${inner}</span>`;
    if (st.hx === 'fill' && text && typeof d.value !== 'number') inner = esc(text.repeat(Math.max(1, Math.floor((m.colW.get(c) - 6) / Math.max(1, [...text].length * 7)))));
    return `<td${attr} style="${esc(cellCss(st, cf, d.align))}">${inner}</td>`;
  };
  let t = `<table class="${ps.gridlines ? 'grid' : ''}" style="width:${cols.reduce((a, c) => a + m.colW.get(c), 0) + hdW}px"><colgroup>${ps.headings ? `<col style="width:${hdW}px">` : ''}${cols.map((c) => `<col style="width:${m.colW.get(c)}px">`).join('')}</colgroup>`;
  if (ps.headings) t += `<tr style="height:20px"><th class="hd"></th>${cols.map((c) => `<th class="hd">${colName(c)}</th>`).join('')}</tr>`;
  for (const r of rows) t += `<tr style="height:${m.rowH.get(r)}px">${ps.headings ? `<th class="hd">${r + 1}</th>` : ''}${cols.map((c, i) => td(r, c, i)).join('')}</tr>`;
  t += '</table>';
  // biểu đồ nằm trên trang này (theo ô neo)
  const xOf = (c: number) => { let x = hdW; for (const k of cols) { if (k === c) return x; x += m.colW.get(k); } return -1; };
  const yOf = (r: number) => { let y = ps.headings ? 20 : 0; for (const k of rows) { if (k === r) return y; y += m.rowH.get(k); } return -1; };
  for (const ch of charts) {
    if (!p.rows.includes(ch.anchor.r) || !p.cols.includes(ch.anchor.c)) continue;
    const x = xOf(ch.anchor.c), y = yOf(ch.anchor.r);
    t += `<div class="chart" style="left:${x + ch.anchor.dx}px;top:${y + ch.anchor.dy}px;width:${ch.w}px;height:${ch.h}px"><img src="${ch.src}" alt="Biểu đồ"></div>`;
  }
  const mg = ps.mg, inch = (v: number) => `${fmt(v)}in`;
  const hf = (code: string, pos: string) => {
    const parts = hfParse(code);
    if (!parts.l && !parts.c && !parts.r) return '';
    const sec = (k: 'l' | 'c' | 'r') => `<span class="${k}">${esc(hfRender(parts[k] ?? '', ctx))}</span>`;
    return `<div class="hf" style="${pos};left:${inch(mg.l)};right:${inch(mg.r)}">${sec('l')}${sec('c')}${sec('r')}</div>`;
  };
  const just = ps.center.h ? 'center' : 'flex-start', items = ps.center.v ? 'center' : 'flex-start';
  const guides = `<div class="mg"><i style="left:${inch(mg.l)};top:0;bottom:0;border-left-width:1px"></i><i style="right:${inch(mg.r)};top:0;bottom:0;border-right-width:1px"></i><i style="top:${inch(mg.t)};left:0;right:0;border-top-width:1px"></i><i style="bottom:${inch(mg.b)};left:0;right:0;border-bottom-width:1px"></i><i style="top:${inch(mg.h)};left:0;right:0;border-top-width:1px;border-color:#b0b0b0"></i><i style="bottom:${inch(mg.f)};left:0;right:0;border-bottom-width:1px;border-color:#b0b0b0"></i></div>`;
  return `<section class="page" data-n="${ctx.page}" style="width:${fmt(L.paper.wmm)}mm;height:${fmt(L.paper.hmm)}mm;page:${named}">`
    + hf(ps.header, `top:${inch(mg.h)}`)
    + `<div class="bd" style="top:${inch(mg.t)};left:${inch(mg.l)};width:${fmt(L.avail.w)}px;height:${fmt(L.avail.h)}px;justify-content:${just};align-items:${items}"><div class="sc" style="zoom:${L.scale}">${t}</div></div>`
    + hf(ps.footer, `bottom:${inch(mg.f)}`)
    + guides + '</section>';
}
export { PX_IN };
