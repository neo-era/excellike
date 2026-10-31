// Đọc nhanh phần dữ liệu ô của .xlsx (sheetData, sharedStrings, styles) bằng regex trên XML gốc,
// thay cho ExcelJS (ExcelJS dựng một đối tượng cho từng ô — chậm gấp ~10 lần với sổ lớn).
// Kiểu ô được dựng theo đúng hình dạng mô hình của ExcelJS (font/fill/alignment/border/numFmt/protection)
// để dùng lại styleIn() có sẵn, và chỉ dựng MỘT lần cho mỗi kiểu (xf) rồi dùng chung.

const XML_ENT: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
export const unescXml = (s: string) => (s.indexOf('&') < 0 ? s : s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-f]+);/gi, (_m, e: string) => {
  const k = e.toLowerCase();
  if (XML_ENT[k]) return XML_ENT[k];
  return String.fromCodePoint(k[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
}));
/** _x000D_ … trong chuỗi OOXML */
const unescOoxml = (s: string) => (s.indexOf('_x') < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16))));
function attrsOf(tag: string): Record<string, string> {
  const o: Record<string, string> = {};
  const re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(tag))) o[m[1].replace(/^\w+:/, '')] = unescXml(m[2] ?? m[3]);
  return o;
}
const stripNs = (xml: string) => xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');
const truthy = (v: string | undefined) => v === undefined || v === '1' || v === 'true';

// ---------- chuỗi dùng chung ----------
export function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  const x = stripNs(xml);
  const re = /<si>([\s\S]*?)<\/si>|<si\/>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(x))) out.push(textOf(m[1] ?? ''));
  return out;
}
/** nối mọi <t> (kể cả các đoạn định dạng <r>), bỏ phiên âm <rPh> */
function textOf(inner: string): string {
  const s = inner.indexOf('<rPh') >= 0 ? inner.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '') : inner;
  let out = '';
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(s))) out += m[1] ?? '';
  return unescOoxml(unescXml(out));
}

// ---------- kiểu ô ----------
const BUILTIN_FMT: Record<number, string> = {
  1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy "h":mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00 ;(#,##0.00)', 40: '#,##0.00 ;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};
function colorIn(a: Record<string, string>): any {
  if (a.rgb) return { argb: a.rgb };
  if (a.theme !== undefined) return a.tint !== undefined ? { theme: +a.theme, tint: +a.tint } : { theme: +a.theme };
  if (a.indexed !== undefined) return { indexed: +a.indexed };
  return undefined;
}
function fontIn(x: string): any {
  const f: any = {};
  const leaf = (tag: string) => { const m = new RegExp(`<${tag}(\\s[^>]*)?/?>`).exec(x); return m ? attrsOf(m[1] ?? '') : null; };
  let a;
  if ((a = leaf('b'))) f.bold = truthy(a.val);
  if ((a = leaf('i'))) f.italic = truthy(a.val);
  if ((a = leaf('strike'))) f.strike = truthy(a.val);
  if ((a = leaf('u'))) f.underline = a.val ?? true;
  if ((a = leaf('vertAlign'))) f.vertAlign = a.val;
  if ((a = leaf('sz'))) f.size = +a.val;
  if ((a = leaf('name'))) f.name = a.val;
  if ((a = leaf('color'))) { const c = colorIn(a); if (c) f.color = c; }
  return f;
}
function fillIn(x: string): any {
  const pm = /<patternFill\b([^>]*?)(?:\/>|>([\s\S]*?)<\/patternFill>)/.exec(x);
  if (!pm) return { type: 'gradient' };
  const a = attrsOf(pm[1]);
  const f: any = { type: 'pattern', pattern: a.patternType ?? 'none' };
  const inner = pm[2] ?? '';
  const fg = /<fgColor\b([^>]*)\/?>/.exec(inner), bg = /<bgColor\b([^>]*)\/?>/.exec(inner);
  if (fg) f.fgColor = colorIn(attrsOf(fg[1]));
  if (bg) f.bgColor = colorIn(attrsOf(bg[1]));
  return f;
}
function borderIn(x: string): any {
  const om = /<border\b([^>]*)>/.exec(x);
  const oa = om ? attrsOf(om[1]) : {};
  const b: any = {};
  for (const side of ['left', 'right', 'top', 'bottom', 'diagonal']) {
    const m = new RegExp(`<${side}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${side}>)`).exec(x);
    if (!m) continue;
    const a = attrsOf(m[1]);
    if (!a.style) continue;
    const s: any = { style: a.style };
    const cm = m[2] ? /<color\b([^>]*)\/?>/.exec(m[2]) : null;
    if (cm) { const c = colorIn(attrsOf(cm[1])); if (c) s.color = c; }
    b[side] = s;
  }
  if (b.diagonal) { b.diagonal.up = oa.diagonalUp === '1' || oa.diagonalUp === 'true'; b.diagonal.down = oa.diagonalDown === '1' || oa.diagonalDown === 'true'; }
  return b;
}
function listOf(x: string, block: string, item: string): string[] {
  const bm = new RegExp(`<${block}\\b[^>]*?(?:/>|>([\\s\\S]*?)</${block}>)`).exec(x);
  if (!bm || !bm[1]) return [];
  return [...bm[1].matchAll(new RegExp(`<${item}\\b[^>]*?(?:/>|>[\\s\\S]*?</${item}>)`, 'g'))].map((m) => m[0]);
}
export interface XfModel { font?: any; fill?: any; border?: any; alignment?: any; protection?: any; numFmt?: string }
/** cellXfs → mô hình kiểu theo dạng ExcelJS (font, fill, border, alignment, protection, numFmt) */
export function parseStyles(xml: string): XfModel[] {
  const x = stripNs(xml);
  const numFmts = new Map<number, string>();
  for (const n of listOf(x, 'numFmts', 'numFmt')) { const a = attrsOf(n); numFmts.set(+a.numFmtId, a.formatCode); }
  const fonts = listOf(x, 'fonts', 'font').map(fontIn);
  const fills = listOf(x, 'fills', 'fill').map(fillIn);
  const borders = listOf(x, 'borders', 'border').map(borderIn);
  return listOf(x, 'cellXfs', 'xf').map((xf) => {
    const om = /<xf\b([^>]*?)\/?>/.exec(xf)!;
    const a = attrsOf(om[1]);
    const m: XfModel = {};
    const nid = a.numFmtId !== undefined ? +a.numFmtId : 0;
    if (nid) { const f = numFmts.get(nid) ?? BUILTIN_FMT[nid]; if (f) m.numFmt = f; }
    if (a.fontId !== undefined && fonts[+a.fontId]) m.font = fonts[+a.fontId];
    if (a.borderId !== undefined && borders[+a.borderId]) m.border = borders[+a.borderId];
    if (a.fillId !== undefined && fills[+a.fillId]) m.fill = fills[+a.fillId];
    const al = /<alignment\b([^>]*?)\/?>/.exec(xf);
    if (al) {
      const aa = attrsOf(al[1]);
      const o: any = {};
      if (aa.horizontal) o.horizontal = aa.horizontal;
      if (aa.vertical) o.vertical = aa.vertical;
      if (aa.wrapText !== undefined) o.wrapText = truthy(aa.wrapText);
      if (aa.shrinkToFit !== undefined) o.shrinkToFit = truthy(aa.shrinkToFit);
      if (aa.indent) o.indent = +aa.indent;
      if (aa.textRotation) o.textRotation = aa.textRotation === '255' ? 'vertical' : +aa.textRotation;
      m.alignment = o;
    }
    const pr = /<protection\b([^>]*?)\/?>/.exec(xf);
    if (pr) { const pa = attrsOf(pr[1]); const o: any = {}; if (pa.locked !== undefined) o.locked = truthy(pa.locked); if (pa.hidden !== undefined) o.hidden = truthy(pa.hidden); m.protection = o; }
    return m;
  });
}

// ---------- dữ liệu ô ----------
export interface FastCell { r: number; c: number; s: number; t: string; v: string | null; f: { text: string | null; t?: string; si?: string; ref?: string } | null; inline: string | null }
export interface FastRow { r: number; ht?: number; hidden?: boolean; lv?: number; collapsed?: boolean }
const A1 = /^([A-Z]+)(\d+)$/;
function a1(ref: string): { r: number; c: number } | null {
  const m = A1.exec(ref);
  if (!m) return null;
  let c = 0;
  for (let i = 0; i < m[1].length; i++) c = c * 26 + m[1].charCodeAt(i) - 64;
  return { r: +m[2] - 1, c: c - 1 };
}
const attrRe = new Map<string, RegExp>();
const attr1 = (tag: string, name: string) => { const i = tag.indexOf(' ' + name + '="'); if (i < 0) { if (tag.indexOf(name) < 0) return undefined; let re = attrRe.get(name); if (!re) attrRe.set(name, (re = new RegExp('\\s' + name + '\\s*=\\s*["\']([^"\']*)'))); const m = re.exec(tag); return m ? m[1] : undefined; } const j = i + name.length + 3; return tag.slice(j, tag.indexOf('"', j)); };

/** Tách và đọc <sheetData>; trả về xml đã bỏ dữ liệu ô (để ExcelJS chỉ đọc phần còn lại) */
export function takeSheetData(xml: string): { xml: string; cells: FastCell[]; rows: FastRow[] } {
  const open = xml.search(/<(?:\w+:)?sheetData\b/);
  if (open < 0) return { xml, cells: [], rows: [] };
  const selfClose = /^<(?:\w+:)?sheetData\b[^>]*\/>/.exec(xml.slice(open));
  if (selfClose) return { xml, cells: [], rows: [] };
  const closeRe = /<\/(?:\w+:)?sheetData>/g;
  closeRe.lastIndex = open;
  const cm = closeRe.exec(xml);
  if (!cm) return { xml, cells: [], rows: [] };
  const body = stripNs(xml.slice(xml.indexOf('>', open) + 1, cm.index));
  const rest = xml.slice(0, open) + '<sheetData/>' + xml.slice(cm.index + cm[0].length);
  const cells: FastCell[] = [], rows: FastRow[] = [];
  const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let rm: RegExpExecArray | null, lastR = -1;
  while ((rm = rowRe.exec(body))) {
    const ra = rm[1];
    const rAttr = attr1(ra, 'r');
    const r = rAttr ? +rAttr - 1 : lastR + 1;
    lastR = r;
    const ht = attr1(ra, 'ht'), hid = attr1(ra, 'hidden'), lv = attr1(ra, 'outlineLevel'), col = attr1(ra, 'collapsed');
    const isHid = hid === '1' || hid === 'true', isCol = col === '1' || col === 'true';
    if (ht !== undefined || isHid || (lv && +lv > 0) || isCol) rows.push({ r, ht: ht !== undefined ? +ht : undefined, hidden: isHid, lv: lv ? +lv : undefined, collapsed: isCol || undefined });
    const inner = rm[2];
    if (!inner) continue;
    cellRe.lastIndex = 0;
    let m: RegExpExecArray | null, lastC = -1;
    while ((m = cellRe.exec(inner))) {
      const ca = m[1];
      const ref = attr1(ca, 'r');
      const p = ref ? a1(ref) : null;
      const c = p ? p.c : lastC + 1;
      lastC = c;
      const s = attr1(ca, 's');
      const t = attr1(ca, 't') ?? 'n';
      const ci = m[2];
      let v: string | null = null, f: FastCell['f'] = null, inline: string | null = null;
      if (ci) {
        const vm = /<v>([\s\S]*?)<\/v>/.exec(ci);
        if (vm) v = vm[1];
        const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(ci);
        if (fm) {
          const fa = fm[1];
          f = { text: fm[2] !== undefined ? unescXml(fm[2]) : null, t: attr1(fa, 't'), si: attr1(fa, 'si'), ref: attr1(fa, 'ref') };
        }
        if (t === 'inlineStr') { const im = /<is>([\s\S]*?)<\/is>/.exec(ci); if (im) inline = textOf(im[1]); }
      }
      cells.push({ r: p ? p.r : r, c, s: s ? +s : 0, t, v, f, inline });
    }
  }
  return { xml: rest, cells, rows };
}
