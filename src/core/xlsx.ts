// Đọc / ghi tệp Excel .xlsx bằng ExcelJS (+ JSZip để tự đọc phần Kiểm tra dữ liệu)
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { colName, parseA1, parseRef, rangeToA1, MAX_ROWS, type Range } from './address';
import { shiftFormula } from './formula';
import { toEn, toVi } from './locale';
import { normalizeInput } from './input';
import { numOf } from './rules';
import { Workbook, dateSerial, type ExtLink, type BookJSON, type Border, type BorderStyle, type SheetJSON, type Style, type PrintSetup } from './workbook';
import type { CFRule, DVRule } from './rules';
import { cellsOf, type Chart, type ChartType, type Ref } from './chart';
import { quoteSheet } from './formula';
import { MARGIN_PRESETS, normPrint, PAPERS, paperByCode, storePrint, type NPrint } from './paginate';

const EPOCH = Date.UTC(1899, 11, 30);
// Độ rộng cột Excel tính theo số ký tự; phông của app rộng hơn Calibri nên lấy 8px/ký tự để số không bị ####
const MDW = 8;
const px2w = (px: number) => Math.round(((px - 5) / MDW) * 100) / 100;
const w2px = (w: number) => Math.round(w * MDW + 5);
const pt2px = (pt: number) => Math.round((pt * 4) / 3);
const px2pt = (px: number) => Math.round(px * 0.75 * 100) / 100;
const DS_SHEET = '_ExcelLike_DS'; // sheet ẩn chứa danh sách dài của Kiểm tra dữ liệu
const STD_ERR = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);

// ---------- màu ----------
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
/** Áp độ sáng (tint) kiểu Excel: theo độ sáng HLS */
function tint(hex: string, t: number) {
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  let h = 0, s = 0, l = (mx + mn) / 2;
  if (mx !== mn) {
    const d = mx - mn;
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = t < 0 ? l * (1 + t) : l * (1 - t) + t;
  const f = (p: number, q: number, x: number) => { if (x < 0) x += 1; if (x > 1) x -= 1; return x < 1 / 6 ? p + (q - p) * 6 * x : x < 1 / 2 ? q : x < 2 / 3 ? p + (q - p) * (2 / 3 - x) * 6 : p; };
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
}
function colorOf(c: any): string | undefined {
  if (!c) return undefined;
  if (typeof c.argb === 'string') return '#' + c.argb.slice(-6).toLowerCase();
  if (typeof c.theme === 'number' && THEME[c.theme]) return '#' + (c.tint ? tint(THEME[c.theme], c.tint) : THEME[c.theme]).toLowerCase();
  if (typeof c.indexed === 'number') { const idx: Record<number, string> = { 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00', 12: '0000FF', 13: 'FFFF00', 64: '000000' }; return idx[c.indexed] ? '#' + idx[c.indexed].toLowerCase() : undefined; }
  return undefined;
}
const argb = (h: string) => ({ argb: 'FF' + h.replace('#', '').toUpperCase().padStart(6, '0') });

// ---------- tiền tố hàm mới ----------
const XLFN = ['XLOOKUP', 'XMATCH', 'IFS', 'SWITCH', 'TEXTJOIN', 'CONCAT', 'MAXIFS', 'MINIFS', 'IFNA', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'UNIQUE', 'SEQUENCE', 'DAYS', 'ISOWEEKNUM', 'CEILING.MATH', 'FLOOR.MATH', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'NORM.DIST', 'NORM.INV', 'NETWORKDAYS.INTL', 'WORKDAY.INTL', 'COVARIANCE.S', 'COVARIANCE.P', 'BITAND', 'BITOR', 'BITXOR'];
const XLWS = ['FILTER', 'SORT', 'SORTBY'];
const fnRe = (names: string[]) => new RegExp(`(^|[^A-Za-z0-9_.])(${names.map((n) => n.replace('.', '\\.')).join('|')})\\(`, 'g');
const RE_XLFN = fnRe(XLFN), RE_XLWS = fnRe(XLWS);
function addPrefixes(f: string) {
  return f.split(/("(?:[^"]|"")*")/).map((p, i) => (i % 2 ? p : p.replace(RE_XLWS, '$1_xlfn._xlws.$2(').replace(RE_XLFN, '$1_xlfn.$2('))).join('');
}
const stripPrefixes = (f: string) => f.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\./g, '');

const BORDER_IN: Record<string, BorderStyle> = { thin: 'thin', hair: 'thin', medium: 'medium', thick: 'thick', dashed: 'dashed', mediumDashed: 'dashed', dashDot: 'dashed', mediumDashDot: 'dashed', dashDotDot: 'dotted', mediumDashDotDot: 'dotted', slantDashDot: 'dashed', dotted: 'dotted', double: 'double' };
const OP_IN: Record<string, CFRule['op']> = { greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le', equal: 'eq', notEqual: 'ne', between: 'between', notBetween: 'notBetween' };
const OP_OUT = Object.fromEntries(Object.entries(OP_IN).map(([k, v]) => [v, k]));

function fixNumFmt(f: string | undefined): string | undefined {
  if (!f || f === 'General') return undefined;
  if (/^m{1,2}[-/]d{1,2}[-/]yy(yy)?$/i.test(f) || f === 'd/m/yyyy' || f === 'd/mm/yyyy') return 'dd/mm/yyyy';
  return f;
}
/** Giá trị trong quy tắc CF/DV đọc từ tệp → chuỗi kiểu app */
function operandIn(x: unknown): string | undefined {
  if (x === undefined || x === null || x === '') return undefined;
  if (typeof x === 'number') return String(x).replace('.', ',');
  if (x instanceof Date) return String(Math.round((x.getTime() - EPOCH) / 86400000));
  const s = String(x);
  if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(s)) return String(+s).replace('.', ',');
  const q = /^"((?:[^"]|"")*)"$/.exec(s);
  if (q) return q[1].replace(/""/g, '"');
  return toVi('=' + stripPrefixes(s)); // công thức / tham chiếu ô
}
/** Giá trị trong quy tắc CF/DV của app → toán hạng ghi vào tệp */
function operandOut(v: string | undefined): string | number | undefined {
  if (v === undefined || v === '') return undefined;
  if (v.startsWith('=')) return addPrefixes(toEn(v)).slice(1);
  const n = numOf(v, dateSerial);
  if (n !== null) return n;
  if (/^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(v)) return v;
  return `"${v.replace(/"/g, '""')}"`;
}
/** Chuỗi đọc từ tệp có cần dấu ' để không bị hiểu thành số/công thức không */
function needsApostrophe(s: string): boolean {
  if (s === '') return false;
  if (/^[='+\-$€£¥₫]/.test(s) || /^\s/.test(s) || /^[\d.,]/.test(s)) return true;
  if (/^(true|false)$/i.test(s.trim()) || /^#[A-Z/0!?]+/.test(s)) return true;
  return normalizeInput(s).content !== s;
}

// ---------- XML tối giản cho phần Kiểm tra dữ liệu ----------
const unesc = (s: string) => s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-f]+);/gi, (_m, e: string) => {
  const map: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
  if (map[e.toLowerCase()]) return map[e.toLowerCase()];
  return String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
});
function attrs(tag: string): Record<string, string> {
  const o: Record<string, string> = {};
  for (const m of tag.matchAll(/([\w:]+)="([^"]*)"/g)) o[m[1].replace(/^\w+:/, '')] = unesc(m[2]);
  return o;
}
interface RawDV { a: Record<string, string>; sqref: string; f1?: string; f2?: string }
function parseDVXml(xml: string): RawDV[] {
  const out: RawDV[] = [];
  for (const m of xml.matchAll(/<(?:\w+:)?dataValidation\b([^>]*?)(\/>|>([\s\S]*?)<\/(?:\w+:)?dataValidation>)/g)) {
    const a = attrs(m[1]);
    const body = m[3] ?? '';
    const f = (n: number) => { const x = new RegExp(`<(?:\\w+:)?formula${n}>(?:\\s*<(?:\\w+:)?f>)?([\\s\\S]*?)(?:</(?:\\w+:)?f>\\s*)?</(?:\\w+:)?formula${n}>`).exec(body); return x ? unesc(x[1]) : undefined; };
    const sq = a.sqref ?? (/<(?:\w+:)?sqref>([\s\S]*?)<\/(?:\w+:)?sqref>/.exec(body)?.[1] ?? '');
    out.push({ a, sqref: unesc(sq), f1: f(1), f2: f(2) });
  }
  return out;
}

// =================== ĐỌC ===================
export async function readXlsx(buf: ArrayBuffer | Uint8Array): Promise<{ wb: Workbook; warnings: string[] }> {
  // 1) Tự đọc Kiểm tra dữ liệu (kể cả dạng cột nguyên C:C và dạng mới trong extLst) rồi gỡ khỏi XML,
  //    tránh ExcelJS trải ra từng ô (rất chậm, tốn bộ nhớ)
  const zip = await JSZip.loadAsync(buf);
  const wbXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
  const relXml = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? '';
  const rels = new Map<string, string>();
  for (const m of relXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) { const a = attrs(m[1]); rels.set(a.Id, a.Target); }
  const dvBySheet = new Map<string, RawDV[]>();
  const breaksBySheet = new Map<string, { rows: number[]; cols: number[] }>();
  const locLinks = new Map<string, [number, number, string][]>();
  const chartsBySheet = new Map<string, RawChart[]>();
  // tính lặp (File → Options → Formulas → Enable iterative calculation)
  const calcTag = /<(?:\w+:)?calcPr\b[^>]*>/.exec(wbXml)?.[0] ?? '';
  const iterOn = /\biterate="(1|true)"/.test(calcTag);
  const calc = iterOn ? { on: true, count: +(/iterateCount="(\d+)"/.exec(calcTag)?.[1] ?? 100), delta: +(/iterateDelta="([^"]+)"/.exec(calcTag)?.[1] ?? 0.001) || 0.001 } : undefined;
  const iterVals: [string, number, number, unknown][] = [];
  // liên kết sang tệp khác: lấy giá trị Excel lưu sẵn trong externalLinkN.xml
  const extLinks: ExtLink[] = [];
  const extSheets: { name: string; cells: unknown[][] }[] = [];
  const erBlock = /<(?:\w+:)?externalReferences>([\s\S]*?)<\/(?:\w+:)?externalReferences>/.exec(wbXml);
  if (erBlock) {
    let index = 0;
    for (const m of erBlock[1].matchAll(/<(?:\w+:)?externalReference\b[^>]*\bid="([^"]+)"/g)) {
      index++;
      const t = rels.get(m[1]);
      if (!t) continue;
      const path = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
      const xml = await zip.file(path)?.async('string');
      if (!xml) continue;
      const relsX = await zip.file(relsPath(path))?.async('string');
      extLinks.push({ index, path, xml, rels: relsX ?? undefined });
      const sx = strip(xml);
      const names = [...sx.matchAll(/<sheetName\b[^>]*\bval="([^"]*)"/g)].map((x) => unesc(x[1]));
      for (const sd of sx.matchAll(/<sheetData\b[^>]*\bsheetId="(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/sheetData>)/g)) {
        const nm = names[+sd[1]];
        if (nm === undefined || !sd[2]) continue;
        const cells: unknown[][] = [];
        for (const cm of sd[2].matchAll(/<cell\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cell>)/g)) {
          const a = attrs(cm[1]);
          const pos = a.r ? parseA1(a.r) : null;
          if (!pos || pos.r > 200000) continue;
          const vm = cm[2] ? /<v>([\s\S]*?)<\/v>/.exec(cm[2]) : null;
          if (!vm) continue;
          const txt = unesc(vm[1]);
          const val = a.t === 'str' || a.t === 's' || a.t === 'inlineStr' ? (needsApostrophe(txt) ? "'" + txt : txt) : a.t === 'b' ? txt === '1' : a.t === 'e' ? (STD_ERR.has(txt) ? txt : null) : Number(txt);
          (cells[pos.r] ??= [])[pos.c] = typeof val === 'number' && !isFinite(val) ? null : val;
        }
        for (let r = 0; r < cells.length; r++) { cells[r] ??= []; for (let c = 0; c < cells[r].length; c++) if (cells[r][c] === undefined) cells[r][c] = null; }
        extSheets.push({ name: `[${index}]${nm}`, cells });
      }
    }
  }
  let pics = 0, unsupportedCharts = 0, shapes = 0, extraNotes = 0;
  let rezip = false;
  // tệp do openpyxl… tạo có Target tuyệt đối (/xl/…) mà ExcelJS không hiểu → đổi sang đường dẫn tương đối
  for (const rp of Object.keys(zip.files).filter((f) => /_rels\/[^/]+\.rels$/.test(f) && f !== '_rels/.rels')) {
    const x = await zip.file(rp)!.async('string');
    if (!/Target="\//.test(x)) continue;
    const srcDir = rp.replace(/_rels\/[^/]+$/, '').split('/').filter(Boolean);
    const y = x.replace(/(<Relationship\b[^>]*?)Target="\/([^"]*)"/g, (all, pre, abs) => {
      if (/TargetMode="External"/.test(all)) return all;
      const t = abs.split('/');
      let k = 0;
      while (k < srcDir.length && k < t.length - 1 && srcDir[k] === t[k]) k++;
      return `${pre}Target="${[...Array(srcDir.length - k).fill('..'), ...t.slice(k)].join('/')}"`;
    });
    zip.file(rp, y); rezip = true;
  }
  for (const m of wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const a = attrs(m[1]);
    const target = rels.get(a.id);
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const file = zip.file(path);
    if (!file) continue;
    let xml = await file.async('string');
    const found: RawDV[] = [];
    const plain = [...xml.matchAll(/<dataValidations\b[\s\S]*?<\/dataValidations>/g)];
    const ext = [...xml.matchAll(/<(\w+):dataValidations\b[\s\S]*?<\/\1:dataValidations>/g)];
    for (const b of [...plain, ...ext]) found.push(...parseDVXml(b[0]));
    if (plain.length || ext.length) {
      xml = xml.replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/g, '').replace(/<(\w+):dataValidations\b[\s\S]*?<\/\1:dataValidations>/g, '');
      zip.file(path, xml);
      rezip = true;
    }
    dvBySheet.set(a.name, found);
    // ngắt trang thủ công (ExcelJS không đọc)
    {
      const br = (tag: string) => [...(new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>`).exec(xml)?.[0] ?? '').matchAll(/<brk\b([^>]*)\/?>/g)].map((b) => +(attrs(b[1]).id ?? 0)).filter((n) => n > 0);
      const rows = br('rowBreaks'), cols = br('colBreaks');
      if (rows.length || cols.length) breaksBySheet.set(a.name, { rows, cols });
      // liên kết trong sổ (chỉ có location, không r:id) — ExcelJS bỏ qua
      const ll: [number, number, string][] = [];
      for (const m of xml.matchAll(/<hyperlink\b([^>]*?)\/?>/g)) {
        const at = attrs(m[1]);
        if (at.location && !Object.keys(at).some((k) => /(^|:)id$/.test(k))) { const p = parseRef(String(at.ref).split(':')[0]); if (p) ll.push([p.r1, p.c1, '#' + at.location]); }
      }
      if (ll.length) locLinks.set(a.name, ll);
    }
    // biểu đồ: tự đọc bản vẽ + biểu đồ rồi gỡ thẻ <drawing> (ExcelJS không đọc được biểu đồ)
    // ghi chú ở vị trí không chuẩn (openpyxl: xl/comments/comment1.xml) làm ExcelJS lỗi → tự đếm rồi gỡ
    {
      const rp = relsPath(path);
      const rx = await zip.file(rp)?.async('string');
      if (rx) {
        let ry = rx;
        for (const rm of rx.matchAll(/<Relationship\b[^>]*\/>/g)) {
          const ra = attrs(rm[0].slice(13, -2));
          if (!/\/(comments|vmlDrawing)$/.test(ra.Type ?? '')) continue;
          const tp = resolve(path, ra.Target ?? '');
          if (/^xl\/comments\d+\.xml$/.test(tp) || (/vmlDrawing$/.test(ra.Type) && /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(tp))) continue;
          if (/comments$/.test(ra.Type)) extraNotes += ((await zip.file(tp)?.async('string')) ?? '').split('<comment ').length - 1;
          ry = ry.replace(rm[0], '');
          if (/vmlDrawing$/.test(ra.Type)) xml = xml.replace(new RegExp(`<legacyDrawing\\b[^>]*r:id="${ra.Id}"[^>]*/>`), '');
          zip.remove(tp);
        }
        if (ry !== rx) { zip.file(rp, ry); zip.file(path, xml); rezip = true; }
      }
    }
    const dm = /<drawing\b[^>]*r:id="([^"]+)"[^>]*\/>/.exec(xml);
    if (dm) {
      const srels = await relsOf(zip, path);
      const dpath = srels.get(dm[1]);
      if (dpath) {
        const got = await readDrawing(zip, dpath);
        chartsBySheet.set(a.name, got.charts);
        pics += got.pics;
        shapes += got.shapes;
        unsupportedCharts += got.unsupported;
        // ExcelJS đọc mọi tệp bản vẽ trong gói và lỗi với biểu đồ → bỏ khỏi bản đưa cho ExcelJS
        for (const p of got.parts) zip.remove(p);
      }
      xml = xml.replace(dm[0], '');
      zip.file(path, xml);
      // gỡ luôn liên kết tới bản vẽ để ExcelJS không đọc phần biểu đồ
      const rp = relsPath(path);
      const rx = await zip.file(rp)?.async('string');
      if (rx) zip.file(rp, rx.replace(new RegExp(`<Relationship\\b[^>]*Id="${dm[1]}"[^>]*/>`), ''));
      rezip = true;
    }
  }
  const data = rezip ? await zip.generateAsync({ type: 'uint8array' }) : buf;

  const x = new ExcelJS.Workbook();
  await x.xlsx.load(data as any);
  const warnings: string[] = [];
  // phông/cỡ chữ phổ biến nhất coi là mặc định (không lưu vào từng ô)
  const fontCount = new Map<string, number>();
  x.eachSheet((ws) => ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: true }, (cell) => {
    const f: any = cell.font; if (f?.name) { const k = f.name + '|' + (f.size ?? 11); fontCount.set(k, (fontCount.get(k) ?? 0) + 1); }
  })));
  const [defName, defSizeS] = ([...fontCount.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 'Calibri|11').split('|');
  const defSize = +defSizeS;
  let images = 0, notes = 0;
  const sheets: SheetJSON[] = [];
  const dsValues = new Map<string, string>(); // ô của sheet danh sách ẩn → giá trị
  x.eachSheet((ws) => {
    if (ws.name === DS_SHEET) {
      ws.eachRow((row, rn) => row.eachCell((cell, cn) => { const v: any = cell.value; dsValues.set(colName(cn - 1) + rn, v?.richText ? v.richText.map((t: any) => t.text).join('') : String(v ?? '')); }));
      return;
    }
    const cells: unknown[][] = [];
    const styles: [number, Style][] = [];
    const masters = new Map<string, { f: string; r: number; c: number }>();
    images += (ws as any).getImages?.().length ?? 0;
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      const r = rn - 1;
      if (!(row as any).cellCount) return;
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        const c = cn - 1;
        const v: any = cell.value;
        let raw: unknown = null;
        let isText = false;
        if (cell.note) notes++;
        if (v === null || v === undefined) raw = null;
        else if (v instanceof Date) raw = Math.round(((v.getTime() - EPOCH) / 86400000) * 1e8) / 1e8;
        else if (typeof v === 'object') {
          if (iterOn && (typeof v.formula === 'string' || typeof v.sharedFormula === 'string') && (typeof v.result === 'number' || typeof v.result === 'string' || typeof v.result === 'boolean')) iterVals.push([ws.name, r, c, v.result]);
          if (typeof v.formula === 'string') { raw = toVi('=' + stripPrefixes(v.formula)); masters.set(cell.address, { f: raw as string, r, c }); }
          else if (typeof v.sharedFormula === 'string') {
            const m = masters.get(v.sharedFormula) ?? (() => { const mc: any = ws.getCell(v.sharedFormula); return mc.formula ? { f: toVi('=' + stripPrefixes(mc.formula)), r: mc.row - 1, c: mc.col - 1 } : null; })();
            raw = m ? shiftFormula(m.f, r - m.r, c - m.c) : (v.result ?? null);
          } else if (Array.isArray(v.richText)) { raw = v.richText.map((t: any) => t.text).join(''); isText = true; }
          else if (typeof v.text === 'string') { raw = v.text; isText = true; }
          else if (v.hyperlink && Array.isArray(v.text?.richText)) { raw = v.text.richText.map((t: any) => t.text).join(''); isText = true; }
          else if (typeof v.error === 'string') raw = STD_ERR.has(v.error) ? v.error : null;
          else raw = null;
        } else { raw = v; isText = typeof v === 'string'; }
        // chữ trông như số/công thức/ngày → giữ là chữ (thêm dấu ' như Excel)
        if (isText && needsApostrophe(raw as string)) raw = "'" + raw;
        if (raw !== null) { (cells[r] ??= [])[c] = raw; }
        const st = styleIn(cell, defName, defSize);
        if (st) styles.push([r * 16384 + c, st]);
        if (st?.ln && !safeLink(st.ln)) delete st.ln;
      });
    });
    for (const row of cells) if (row) for (let i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = null;
    for (let i = 0; i < cells.length; i++) if (!cells[i]) cells[i] = [];
    const props: any = ws.properties ?? {};
    const colDef = props.defaultColWidth ? w2px(props.defaultColWidth) : w2px(8.43);
    const rowDef = props.defaultRowHeight ? pt2px(props.defaultRowHeight) : 20;
    const colW: [number, number][] = [];
    const hiddenCols: number[] = [];
    for (let i = 1; i <= Math.max(ws.columnCount, 1); i++) {
      const col: any = ws.getColumn(i);
      if (col.hidden) hiddenCols.push(i - 1);
      if (col.width && w2px(col.width) !== colDef && w2px(col.width) > 0) colW.push([i - 1, w2px(col.width)]);
    }
    const af: any = ws.autoFilter;
    const afRef = typeof af === 'string' ? af : af?.from ? `${typeof af.from === 'string' ? af.from : colName(af.from.column - 1) + af.from.row}:${typeof af.to === 'string' ? af.to : colName(af.to.column - 1) + af.to.row}` : null;
    const fr = afRef ? parseRef(afRef) : null;
    const rowH: [number, number][] = [];
    const hiddenRows: number[] = [];
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      const r = rn - 1;
      // hàng ẩn trong vùng lọc là do lọc → bỏ lọc sẽ hiện lại
      if (row.hidden) hiddenRows.push(r);
      if (row.height && pt2px(row.height) !== rowDef && pt2px(row.height) > 0) rowH.push([r, pt2px(row.height)]);
    });
    const view: any = ws.views?.[0];
    const freeze = view?.state === 'frozen' ? { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 } : { rows: 0, cols: 0 };
    const merges = ((ws.model as any).merges ?? []).map((a: string) => parseRef(a)).filter(Boolean) as Range[];
    const cf: CFRule[] = [];
    for (const blk of ((ws as any).conditionalFormattings ?? []) as any[]) {
      for (const ref of String(blk.ref).split(/\s+/)) {
        const range = parseRef(ref);
        if (!range) continue;
        for (const rule of blk.rules) cf.push(cfIn(rule, range));
      }
    }
    const dv: DVRule[] = [];
    for (const d of dvBySheet.get(ws.name) ?? []) for (const ref of d.sqref.split(/\s+/)) { const g = parseRef(ref); if (g) dv.push(dvRuleIn(d, g)); }
    const print = printIn(ws);
    const props2: any = ws.properties ?? {};
    const tc = props2.tabColor?.argb ? '#' + String(props2.tabColor.argb).slice(-6).toLowerCase() : undefined;
    for (const [r, c, ln] of locLinks.get(ws.name) ?? []) {
      const k = r * 16384 + c, e = styles.find((x) => x[0] === k);
      if (e) e[1].ln = ln; else styles.push([k, { ln }]);
    }
    sheets.push({ name: ws.name, cells, styles, colW, rowH, freeze, colDef, rowDef, hiddenRows, hiddenCols, merges, cf, dv, filter: fr ? { range: fr, conds: {} } : null, print, breaks: breaksBySheet.get(ws.name), tabColor: tc, state: (ws as any).state === 'hidden' || (ws as any).state === 'veryHidden' ? 'hidden' : undefined });
  });
  if (!sheets.length) throw new Error('Tệp không có trang tính nào.');
  // danh sách dài do app ghi ra sheet ẩn → đưa về danh sách gõ sẵn
  for (const sh of sheets) for (const d of sh.dv ?? []) {
    if (d.type === 'list' && d.listRef && d.listRef.replace(/'/g, '').startsWith(DS_SHEET + '!')) {
      const g = parseRef(d.listRef.split('!')[1]);
      if (g) { d.list = []; for (let r = g.r1; r <= g.r2; r++) d.list.push(dsValues.get(colName(g.c1) + (r + 1)) ?? ''); delete d.listRef; }
    }
  }
  if (images + pics) warnings.push(`Tệp có ${images + pics} hình ảnh — chưa hiển thị và sẽ mất nếu lưu đè.`);
  if (unsupportedCharts) warnings.push(`Tệp có ${unsupportedCharts} biểu đồ loại chưa hỗ trợ (tán xạ, radar…) — sẽ mất nếu lưu đè.`);
  if (shapes) warnings.push(`Tệp có ${shapes} hình vẽ/hộp chữ — chưa hỗ trợ, sẽ mất nếu lưu đè.`);
  notes += extraNotes;
  if (notes) warnings.push(`Tệp có ${notes} ghi chú (comment) — chưa hỗ trợ, sẽ mất nếu lưu đè.`);
  if (extLinks.length) warnings.push(`Tệp có liên kết tới ${extLinks.length} tệp khác — dùng giá trị lưu sẵn trong tệp (như Excel khi chưa cập nhật liên kết); liên kết được giữ nguyên khi lưu.`);
  const book: BookJSON = { v: 2, name: 'Sổ tính', sheets, calc, iterVals: iterVals.length ? iterVals : undefined, ext: extLinks.length ? { links: extLinks, sheets: extSheets } : undefined };
  const wb = Workbook.fromJSON(book);
  // gắn biểu đồ sau khi đã có độ rộng cột/chiều cao hàng
  for (const id of wb.sheetIds()) {
    const raw = chartsBySheet.get(wb.sheetName(id));
    if (raw?.length) wb.meta(id).charts = raw.map((rc, i) => chartIn(wb, id, rc, i)).filter((c): c is Chart => !!c);
  }
  let bad = 0;
  for (const id of wb.sheetIds()) { const u = wb.used(id); for (let r = 0; r < u.rows && bad < 1000; r++) for (let c = 0; c < u.cols; c++) { const v: any = wb.value(id, r, c); if (v && typeof v === 'object' && v.type === 'ERROR' && wb.isFormula(id, r, c)) bad++; } }
  if (bad) warnings.push(`${bad} công thức có cú pháp chưa hỗ trợ (hiện #ERROR!).`);
  return { wb, warnings };
}

function styleIn(cell: ExcelJS.Cell, defName: string, defSize: number): Style | null {
  const st: Style = {};
  const f: any = cell.font;
  if (f) {
    if (f.bold) st.b = true;
    if (f.italic) st.i = true;
    if (f.underline && f.underline !== 'none') { st.u = true; if (f.underline === 'double' || f.underline === 'doubleAccounting') st.ud = true; }
    if (f.strike) st.s = true;
    if (f.vertAlign === 'superscript') st.vp = 'sup'; else if (f.vertAlign === 'subscript') st.vp = 'sub';
    const fc = colorOf(f.color);
    if (fc && !(f.color?.theme === 1 && !f.color?.tint) && fc !== '#000000') st.fg = fc;
    if (f.name && f.name !== defName) st.fn = f.name;
    if (f.size && f.size !== defSize) st.fs = f.size;
  }
  const fl: any = cell.fill;
  if (fl?.type === 'pattern' && fl.pattern === 'solid') { const c = colorOf(fl.fgColor); if (c) st.bg = c; }
  const al: any = cell.alignment;
  if (al) {
    const h = { left: 'l', center: 'c', centerContinuous: 'c', right: 'r' }[al.horizontal as string];
    if (h) st.al = h as Style['al'];
    const hx = { fill: 'fill', justify: 'just', centerContinuous: 'cas', distributed: 'dist' }[al.horizontal as string];
    if (hx) st.hx = hx as Style['hx'];
    const v = { top: 't', middle: 'm' }[al.vertical as string];
    if (v) st.va = v as Style['va'];
    if (al.wrapText) st.wr = true;
    if (al.shrinkToFit) st.shk = true;
    if (al.indent) st.ind = +al.indent;
    const rot = al.textRotation === 'vertical' ? 255 : +al.textRotation;
    if (rot) st.rot = rot > 90 && rot <= 180 ? 90 - rot : rot; // OOXML: 91–180 = âm
  }
  const pr: any = cell.protection;
  if (pr) { if (pr.locked === false) st.lk = false; if (pr.hidden) st.hid = true; }
  const hl: any = (cell as any).hyperlink;
  if (typeof hl === 'string' && hl) st.ln = hl;
  const nf = fixNumFmt(cell.numFmt);
  if (nf) st.fmt = nf;
  const bd: any = cell.border;
  if (bd) {
    const o: NonNullable<Style['bd']> = {};
    for (const [k, side] of [['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left']] as const) {
      const b = bd[side];
      if (b?.style && BORDER_IN[b.style]) { const x: Border = { s: BORDER_IN[b.style] }; const c = colorOf(b.color); if (c && c !== '#000000') x.c = c; o[k] = x; }
    }
    const dg = bd.diagonal;
    if (dg?.style && BORDER_IN[dg.style] && (dg.up || dg.down)) {
      const x: Border = { s: BORDER_IN[dg.style] }; const c = colorOf(dg.color); if (c && c !== '#000000') x.c = c;
      if (dg.down) o.d = { ...x };
      if (dg.up) o.u = { ...x };
    }
    if (Object.keys(o).length) st.bd = o;
  }
  return Object.keys(st).length ? st : null;
}

const RE_TEXT = /^NOT\(ISERROR\(SEARCH\("((?:[^"]|"")*)",\$?[A-Z]+\$?\d+\)\)\)$/;
function cfIn(rule: any, range: Range): CFRule {
  const sty = rule.style ?? {};
  const fmt = { fg: colorOf(sty.font?.color), bg: colorOf(sty.fill?.bgColor ?? sty.fill?.fgColor), b: sty.font?.bold || undefined, i: sty.font?.italic || undefined };
  for (const k of Object.keys(fmt) as (keyof typeof fmt)[]) if (fmt[k] === undefined) delete fmt[k];
  const f0 = String(rule.formulae?.[0] ?? '').replace(/\s/g, '');
  switch (rule.type) {
    case 'cellIs':
      if (OP_IN[rule.operator]) return { range, type: 'cellIs', op: OP_IN[rule.operator], v1: operandIn(rule.formulae?.[0]), v2: operandIn(rule.formulae?.[1]), fmt };
      break;
    case 'containsText': {
      if (rule.text && (!rule.operator || rule.operator === 'containsText')) return { range, type: 'text', text: rule.text, fmt };
      const m = RE_TEXT.exec(f0);
      if (m) return { range, type: 'text', text: m[1].replace(/""/g, '"'), fmt };
      break;
    }
    case 'top10': return { range, type: rule.bottom ? 'bottom' : 'top', rank: rule.rank ?? 10, percent: !!rule.percent, fmt };
    case 'aboveAverage': return { range, type: rule.aboveAverage === false ? 'belowAvg' : 'aboveAvg', fmt };
    case 'colorScale': return { range, type: 'scale', colors: (rule.color ?? []).map((c: any) => colorOf(c) ?? '#ffffff') };
    case 'dataBar': return { range, type: 'bar', colors: [colorOf(rule.color) ?? '#638ec6'] };
    case 'expression': {
      const m = /^COUNTIF\(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+,\$?[A-Z]+\$?\d+\)([>=])1$/.exec(f0);
      if (m) return { range, type: m[1] === '>' ? 'dup' : 'unique', fmt };
      const t = RE_TEXT.exec(f0);
      if (t) return { range, type: 'text', text: t[1].replace(/""/g, '"'), fmt };
      break;
    }
  }
  return { range, type: 'cellIs', raw: rule };
}

function dvRuleIn(d: RawDV, range: Range): DVRule {
  const a = d.a;
  const base = { range, allowBlank: a.allowBlank === '1' || a.allowBlank === 'true', error: a.showErrorMessage === '0' ? undefined : a.error || undefined, errorTitle: a.errorTitle || undefined, errorStyle: (a.errorStyle as DVRule['errorStyle']) || undefined, prompt: a.showInputMessage === '0' ? undefined : a.prompt || undefined, promptTitle: a.promptTitle || undefined };
  const type = a.type ?? 'none';
  if (type === 'list') {
    const s = d.f1 ?? '';
    if (/^".*"$/.test(s)) return { ...base, type: 'list', list: s.slice(1, -1).split(',').map((x) => x.trim()) };
    return { ...base, type: 'list', listRef: s };
  }
  if (type === 'whole' || type === 'decimal' || type === 'date' || type === 'textLength') {
    return { ...base, type, op: OP_IN[a.operator ?? 'between'] ?? 'between', v1: operandIn(d.f1), v2: operandIn(d.f2) };
  }
  // loại chưa hỗ trợ (giờ, công thức tùy chỉnh…): giữ nguyên để ghi lại
  return { ...base, type: 'custom', raw: { type, operator: a.operator, formulae: [d.f1, d.f2].filter((v) => v !== undefined), allowBlank: base.allowBlank, showErrorMessage: a.showErrorMessage !== '0', error: a.error, errorTitle: a.errorTitle, errorStyle: a.errorStyle, showInputMessage: a.showInputMessage === '1', prompt: a.prompt, promptTitle: a.promptTitle } };
}

// =================== GHI ===================
export async function writeXlsx(wb: Workbook): Promise<Uint8Array> {
  const x = new ExcelJS.Workbook();
  x.creator = 'ExcelLike';
  x.created = new Date();
  const json = wb.toJSON();
  const dsLists: string[][] = [];
  json.sheets.forEach((sh, si) => {
    const id = wb.sheetIds()[si];
    const fz = sh.freeze ?? { rows: 0, cols: 0 };
    const ws = x.addWorksheet(sh.name, {
      views: fz.rows || fz.cols ? [{ state: 'frozen', xSplit: fz.cols, ySplit: fz.rows, topLeftCell: colName(fz.cols) + (fz.rows + 1) } as any] : [{}],
      properties: { defaultColWidth: px2w(sh.colDef ?? 64), defaultRowHeight: px2pt(sh.rowDef ?? 20), ...(sh.tabColor ? { tabColor: { argb: 'FF' + sh.tabColor.replace('#', '').toUpperCase() } } : {}) } as any,
      ...(sh.state === 'hidden' ? { state: 'hidden' } : {}),
    } as any);
    for (const [c, w] of sh.colW ?? []) { const col = ws.getColumn(c + 1); if (w === 0) col.hidden = true; else col.width = px2w(w); }
    for (const c of sh.hiddenCols ?? []) { const col = ws.getColumn(c + 1); if (!col.width) col.width = px2w(sh.colDef ?? 64); col.hidden = true; }
    if (sh.colDef) for (let c = 1; c <= Math.max(1, wb.used(id).cols); c++) { const col = ws.getColumn(c); if (!col.width && !col.hidden) col.width = px2w(sh.colDef); }
    for (const [r, h] of sh.rowH ?? []) { const row = ws.getRow(r + 1); if (h === 0) row.hidden = true; else row.height = px2pt(h); }
    for (const r of sh.hiddenRows ?? []) ws.getRow(r + 1).hidden = true;
    for (const r of sh.breaks?.rows ?? []) ws.getRow(r).addPageBreak();
    // giá trị
    sh.cells.forEach((row, r) => row?.forEach((raw, c) => {
      if (raw === null || raw === undefined || raw === '') return;
      const cell = ws.getCell(r + 1, c + 1);
      const v: any = wb.value(id, r, c);
      if (typeof raw === 'string' && raw.startsWith('=') && raw.length > 1) {
        const formula = addPrefixes(toEn(raw)).slice(1);
        const result = v && typeof v === 'object' ? (STD_ERR.has(v.value) ? { error: v.value } : undefined) : v;
        cell.value = (result === undefined ? { formula } : { formula, result }) as any;
      } else if (v && typeof v === 'object') {
        cell.value = STD_ERR.has(v.value) ? ({ error: v.value } as any) : String(raw).replace(/^'/, '');
      } else if (typeof raw === 'string') {
        cell.value = typeof v === 'number' || typeof v === 'boolean' ? v : raw.startsWith("'") ? raw.slice(1) : raw;
      } else cell.value = raw as any;
    }));
    // định dạng
    for (const [k, st] of sh.styles as [number, Style][]) {
      const r = Math.floor(k / 16384), c = k % 16384;
      styleOut(ws.getCell(r + 1, c + 1), st, wb, id, r, c);
    }
    for (const g of sh.merges ?? []) (ws as any).mergeCellsWithoutStyle(g.r1 + 1, g.c1 + 1, g.r2 + 1, g.c2 + 1);
    for (const rule of sh.cf ?? []) {
      const out = cfOut(rule);
      if (out) ws.addConditionalFormatting({ ref: refOut(rule.range), rules: [out] } as any);
    }
    for (const rule of sh.dv ?? []) (ws as any).dataValidations.add(refOut(rule.range), dvOut(rule, dsLists));
    if (sh.filter) ws.autoFilter = rangeToA1(sh.filter.range);
    if (sh.print) printOut(ws, sh.print);
  });
  // trang tính đang mở khi mở tệp: trang hiện đầu tiên (không trỏ vào trang ẩn)
  const firstVis = json.sheets.findIndex((sh) => sh.state !== 'hidden');
  if (firstVis > 0) (x as any).views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: firstVis, activeTab: firstVis, visibility: 'visible' }];
  if (dsLists.length) {
    const ds = x.addWorksheet(DS_SHEET, { state: 'hidden' } as any);
    dsLists.forEach((items, c) => items.forEach((v, r) => { ds.getCell(r + 1, c + 1).value = v; }));
  }
  const buf = new Uint8Array(await x.xlsx.writeBuffer() as ArrayBuffer);
  const withCharts = wb.sheetIds().some((id) => wb.meta(id).charts.length);
  let out = withCharts ? await addChartsToZip(buf, wb) : buf;
  if (wb.extLinks.length || wb.iterate.on) out = await addWorkbookExtras(out, wb);
  if (json.sheets.some((sh) => (sh.styles as [number, Style][]).some(([, st]) => st.ln?.startsWith('#')))) out = await fixLocations(out);
  if (json.sheets.some((sh) => sh.breaks?.cols.length)) out = await addColBreaks(out, json.sheets.map((sh) => sh.breaks?.cols ?? []));
  return out;
}

/** Ghi lại liên kết ngoài (externalLinks) và chế độ tính lặp vào workbook.xml */
async function addWorkbookExtras(buf: Uint8Array, wb: Workbook): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buf);
  let wx = await zip.file('xl/workbook.xml')!.async('string');
  let rx = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  let ct = await zip.file('[Content_Types].xml')!.async('string');
  if (wb.extLinks.length) {
    // giữ đúng thứ tự [1], [2]… như tệp gốc (công thức tham chiếu theo số thứ tự)
    const links = [...wb.extLinks].sort((a, b) => a.index - b.index);
    let refs = '';
    links.forEach((l, i) => {
      const path = `xl/externalLinks/externalLink${i + 1}.xml`;
      zip.file(path, l.xml);
      if (l.rels) zip.file(relsPath(path), l.rels);
      const id = `rIdXLExt${i + 1}`;
      rx = rx.replace('</Relationships>', `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink${i + 1}.xml"/></Relationships>`);
      if (!ct.includes(`/${path}"`)) ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>`);
      refs += `<externalReference r:id="${id}"/>`;
    });
    if (!/xmlns:r=/.test(wx.slice(0, wx.indexOf('>', wx.indexOf('<workbook'))))) wx = wx.replace('<workbook ', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    wx = wx.replace('</sheets>', `</sheets><externalReferences>${refs}</externalReferences>`);
  }
  if (wb.iterate.on) {
    const it = `iterate="1" iterateCount="${wb.iterate.count}" iterateDelta="${wb.iterate.delta}"`;
    if (/<calcPr\b/.test(wx)) wx = wx.replace(/<calcPr\b([^>]*?)(\/?)>/, (_m, a, sl) => `<calcPr${a.replace(/\s(iterate|iterateCount|iterateDelta)="[^"]*"/g, '')} ${it}${sl}>`);
    else wx = wx.replace(/(<\/definedNames>|<\/externalReferences>|<\/sheets>)(?![\s\S]*(<\/definedNames>|<\/externalReferences>))/, `$1<calcPr ${it}/>`);
  }
  zip.file('xl/workbook.xml', wx);
  zip.file('xl/_rels/workbook.xml.rels', rx);
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
/** Vùng cột nguyên ghi gọn kiểu C1:C1048576 */
const refOut = (g: Range) => (g.r2 >= MAX_ROWS - 1 && g.r1 === 0 ? `${colName(g.c1)}1:${colName(g.c2)}${MAX_ROWS}` : rangeToA1(g));

function styleOut(cell: ExcelJS.Cell, st: Style, wb: Workbook, id: number, r: number, c: number) {
  const font: any = {};
  if (st.b) font.bold = true;
  if (st.i) font.italic = true;
  if (st.u) font.underline = st.ud ? 'double' : true;
  if (st.s) font.strike = true;
  if (st.vp) font.vertAlign = st.vp === 'sup' ? 'superscript' : 'subscript';
  if (st.fg) font.color = argb(st.fg);
  if (st.fn) font.name = st.fn;
  if (st.fs) font.size = st.fs;
  if (Object.keys(font).length) cell.font = font;
  if (st.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(st.bg) };
  const al: any = {};
  if (st.al) al.horizontal = { l: 'left', c: 'center', r: 'right' }[st.al];
  if (st.va) al.vertical = { t: 'top', m: 'middle', b: 'bottom' }[st.va];
  if (st.hx) al.horizontal = { fill: 'fill', just: 'justify', cas: 'centerContinuous', dist: 'distributed' }[st.hx];
  if (st.wr) al.wrapText = true;
  if (st.shk) al.shrinkToFit = true;
  if (st.ind) al.indent = st.ind;
  if (st.rot) al.textRotation = st.rot === 255 ? 'vertical' : st.rot;
  if (Object.keys(al).length) cell.alignment = al;
  if (st.lk === false || st.hid) cell.protection = { locked: st.lk !== false, hidden: !!st.hid } as any;
  // liên kết: chỉ ghi cho ô chữ có nội dung (ExcelJS cần text; số/công thức giữ nguyên giá trị)
  if (st.ln && safeLink(st.ln)) {
    const cur: any = cell.value;
    if (typeof cur === 'string' && cur) cell.value = { text: cur, hyperlink: st.ln } as any;
  }
  const fmt = st.fmt ?? wb.effFmt(id, r, c, st);
  if (fmt) cell.numFmt = fmt;
  if (st.bd) {
    const b: any = {};
    for (const [k, side] of [['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left']] as const) {
      const x = st.bd[k];
      if (x) b[side] = { style: x.s, color: argb(x.c ?? '#000000') };
    }
    const dg = st.bd.d ?? st.bd.u;
    if (dg) b.diagonal = { up: !!st.bd.u, down: !!st.bd.d, style: dg.s, color: argb(dg.c ?? '#000000') };
    cell.border = b;
  }
}

function cfFmt(f: CFRule['fmt']) {
  const o: any = {};
  if (f?.fg || f?.b || f?.i) o.font = { ...(f.fg ? { color: argb(f.fg) } : {}), ...(f.b ? { bold: true } : {}), ...(f.i ? { italic: true } : {}) };
  if (f?.bg) o.fill = { type: 'pattern', pattern: 'solid', bgColor: argb(f.bg) };
  return o;
}
function cfOut(rule: CFRule): any {
  if (rule.raw) return rule.raw;
  const style = cfFmt(rule.fmt);
  const a = rangeToA1(rule.range).split(':');
  const abs = (s: string) => s.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  switch (rule.type) {
    case 'cellIs': return { type: 'cellIs', operator: OP_OUT[rule.op ?? 'gt'], formulae: [operandOut(rule.v1), operandOut(rule.v2)].filter((v) => v !== undefined).map(String), style };
    case 'text': return { type: 'expression', formulae: [`NOT(ISERROR(SEARCH("${(rule.text ?? '').replace(/"/g, '""')}",${a[0]})))`], style };
    case 'dup': case 'unique': return { type: 'expression', formulae: [`COUNTIF(${abs(a[0])}:${abs(a[1] ?? a[0])},${a[0]})${rule.type === 'dup' ? '>' : '='}1`], style };
    case 'top': case 'bottom': return { type: 'top10', rank: rule.rank ?? 10, percent: !!rule.percent, bottom: rule.type === 'bottom', style };
    case 'aboveAvg': case 'belowAvg': return { type: 'aboveAverage', aboveAverage: rule.type === 'aboveAvg', style };
    case 'scale': {
      const cs = rule.colors ?? ['#f8696b', '#63be7b'];
      return { type: 'colorScale', cfvo: cs.length === 3 ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }], color: cs.map(argb) };
    }
    case 'bar': return { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: argb(rule.colors?.[0] ?? '#638ec6') };
  }
  return null;
}
function dvOut(rule: DVRule, ds: string[][]): any {
  if (rule.raw) return rule.raw;
  const base: any = { allowBlank: rule.allowBlank !== false, showErrorMessage: true, errorStyle: rule.errorStyle ?? 'stop', error: rule.error, errorTitle: rule.errorTitle, showInputMessage: !!rule.prompt, prompt: rule.prompt, promptTitle: rule.promptTitle };
  if (rule.type === 'list') {
    if (rule.listRef) return { ...base, type: 'list', formulae: [addPrefixes(toEn('=' + rule.listRef.replace(/^=/, ''))).slice(1)] };
    const items = rule.list ?? [];
    const inline = items.join(',');
    // Excel: danh sách gõ sẵn tối đa 255 ký tự, không chứa dấu phẩy/ngoặc kép → nếu vượt thì đưa ra sheet ẩn
    if (inline.length > 255 || items.some((v) => /[,"]/.test(v))) {
      ds.push(items);
      const col = colName(ds.length - 1);
      return { ...base, type: 'list', formulae: [`'${DS_SHEET}'!$${col}$1:$${col}$${Math.max(1, items.length)}`] };
    }
    return { ...base, type: 'list', formulae: [`"${inline}"`] };
  }
  let f: unknown[] = [operandOut(rule.v1), operandOut(rule.v2)].filter((v) => v !== undefined);
  // ExcelJS đổi ngày từ đối tượng Date → truyền Date cho loại "date"
  if (rule.type === 'date') f = f.map((v) => (typeof v === 'number' ? new Date(EPOCH + v * 86400000) : v));
  return { ...base, type: rule.type, operator: OP_OUT[rule.op ?? 'between'], formulae: f };
}


// =================== BIỂU ĐỒ trong tệp (DrawingML) ===================
const EMU = 9525;
interface RawChart { type: ChartType | null; title?: string; legend: Chart['legend']; labels: boolean; series: { name?: string; cats?: string; vals: string }[]; from: { c: number; co: number; r: number; ro: number }; to?: { c: number; co: number; r: number; ro: number }; ext?: { cx: number; cy: number } }

const dirOf = (p: string) => p.slice(0, p.lastIndexOf('/') + 1);
function resolve(base: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/');
  const out: string[] = [];
  for (const x of parts) { if (x === '..') out.pop(); else if (x !== '.') out.push(x); }
  return out.join('/');
}
const relsPath = (p: string) => dirOf(p) + '_rels/' + p.slice(p.lastIndexOf('/') + 1) + '.rels';
async function relsOf(zip: JSZip, part: string): Promise<Map<string, string>> {
  const xml = (await zip.file(relsPath(part))?.async('string')) ?? '';
  const m = new Map<string, string>();
  for (const x of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) { const a = attrs(x[1]); if (a.Id && a.Target) m.set(a.Id, resolve(part, a.Target)); }
  return m;
}
// bỏ tiền tố namespace (xdr:, c:, a:) để đọc được cả tệp dùng namespace mặc định
const strip = (xml: string) => xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');
const tagVal = (xml: string, tag: string) => { const m = new RegExp(`<${tag}\\b[^>]*\\bval="([^"]*)"`).exec(xml); return m ? m[1] : undefined; };
const between = (xml: string, tag: string) => { const m = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml); return m ? m[1] : undefined; };
function anchorPart(xml: string | undefined) {
  if (!xml) return undefined;
  const n = (t: string) => +(new RegExp(`<${t}>\\s*(-?\\d+)\\s*</${t}>`).exec(xml)?.[1] ?? 0);
  return { c: n('col'), co: n('colOff'), r: n('row'), ro: n('rowOff') };
}

async function readDrawing(zip: JSZip, dpath: string): Promise<{ charts: RawChart[]; pics: number; unsupported: number; shapes: number; parts: string[] }> {
  const xml = strip((await zip.file(dpath)?.async('string')) ?? '');
  const rels = await relsOf(zip, dpath);
  const charts: RawChart[] = [];
  let pics = 0, unsupported = 0, shapes = 0;
  for (const m of xml.matchAll(/<(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/\1>/g)) {
    const a = m[0];
    if (/<pic\b/.test(a)) { pics++; continue; }
    const rid = /<chart\b[^>]*\bid="([^"]+)"/.exec(a)?.[1];
    if (!rid) { shapes++; continue; }
    const cpath = rels.get(rid);
    const cxml = cpath ? await zip.file(cpath)?.async('string') : undefined;
    if (!cxml) continue;
    const rc = parseChartXml(strip(cxml));
    if (!rc.type) { unsupported++; continue; }
    rc.from = anchorPart(between(a, 'from')) ?? { c: 0, co: 0, r: 0, ro: 0 };
    rc.to = anchorPart(between(a, 'to'));
    const ext = /<ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(a);
    if (ext) rc.ext = { cx: +ext[1], cy: +ext[2] };
    charts.push(rc);
  }
  return { charts, pics, unsupported, shapes, parts: [dpath, relsPath(dpath), ...[...rels.values()].filter((p) => /charts\//.test(p)).flatMap((p) => [p, relsPath(p)])] };
}

function parseChartXml(xml: string): RawChart {
  const plot = between(xml, 'plotArea') ?? '';
  let type: ChartType | null = null;
  let body = '';
  const kinds: [RegExp, (b: string) => ChartType][] = [
    [/<(bar|bar3D)Chart>([\s\S]*?)<\/\1Chart>/, (b) => (tagVal(b, 'barDir') === 'bar' ? 'bar' : /stacked/i.test(tagVal(b, 'grouping') ?? '') ? 'colStack' : 'col')],
    [/<(line|line3D)Chart>([\s\S]*?)<\/\1Chart>/, () => 'line'],
    [/<(area|area3D)Chart>([\s\S]*?)<\/\1Chart>/, () => 'area'],
    [/<(pie|pie3D|ofPie)Chart>([\s\S]*?)<\/\1Chart>/, () => 'pie'],
    [/<(doughnut)Chart>([\s\S]*?)<\/\1Chart>/, () => 'doughnut'],
  ];
  for (const [re, f] of kinds) { const m = re.exec(plot); if (m) { body = m[2]; type = f(m[2]); break; } }
  const series: RawChart['series'] = [];
  for (const m of body.matchAll(/<ser>([\s\S]*?)<\/ser>/g)) {
    const sx = m[1];
    const fOf = (part?: string) => { const f = part ? /<f>([\s\S]*?)<\/f>/.exec(part)?.[1] : undefined; return f ? unesc(f) : undefined; };
    const vals = fOf(between(sx, 'val'));
    if (!vals) continue;
    series.push({ name: fOf(between(sx, 'tx')), cats: fOf(between(sx, 'cat')), vals });
  }
  const chartPart = between(xml, 'chart') ?? '';
  const titleXml = /<title>([\s\S]*?)<\/title>/.exec(chartPart.replace(/<plotArea>[\s\S]*<\/plotArea>/, ''))?.[1];
  const title = titleXml ? [...titleXml.matchAll(/<t>([\s\S]*?)<\/t>/g)].map((x) => unesc(x[1])).join('') || undefined : undefined;
  const legXml = /<legend>([\s\S]*?)<\/legend>/.exec(chartPart)?.[1];
  const lp = legXml !== undefined ? tagVal(legXml, 'legendPos') ?? 'r' : null;
  const legend: Chart['legend'] = lp === null ? 'none' : lp === 'b' ? 'b' : lp === 't' ? 't' : 'r';
  const labels = /<showVal val="1"\/>|<showPercent val="1"\/>/.test(body);
  return { type: series.length ? type : null, title, legend, labels, series, from: { c: 0, co: 0, r: 0, ro: 0 } };
}

function refIn(wb: Workbook, sheet: number, f?: string): Ref | undefined {
  if (!f) return undefined;
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(f.trim());
  let id = sheet, ref = f.trim();
  if (m) { const sid = wb.hf.getSheetId((m[1] ?? m[2]).replace(/''/g, "'")); if (sid === undefined) return undefined; id = sid; ref = m[3]; }
  const g = parseRef(ref.replace(/^\(|\)$/g, ''));
  return g ? { sheet: id, range: g } : undefined;
}
function chartIn(wb: Workbook, sheet: number, rc: RawChart, i: number): Chart | null {
  const series = rc.series.map((x) => ({ name: refIn(wb, sheet, x.name), cats: refIn(wb, sheet, x.cats), vals: refIn(wb, sheet, x.vals)! })).filter((x) => x.vals);
  if (!series.length || !rc.type) return null;
  const m = wb.meta(sheet);
  const x1 = m.colW.offset(rc.from.c) + rc.from.co / EMU, y1 = m.rowH.offset(rc.from.r) + rc.from.ro / EMU;
  let w = 480, h = 288;
  if (rc.to) { w = m.colW.offset(rc.to.c) + rc.to.co / EMU - x1; h = m.rowH.offset(rc.to.r) + rc.to.ro / EMU - y1; }
  else if (rc.ext) { w = rc.ext.cx / EMU; h = rc.ext.cy / EMU; }
  return { id: 'x' + i + Math.random().toString(36).slice(2, 8), type: rc.type, title: rc.title, series, legend: rc.legend, labels: rc.labels, anchor: { r: rc.from.r, c: rc.from.c, dx: Math.round(rc.from.co / EMU), dy: Math.round(rc.from.ro / EMU) }, w: Math.max(80, Math.round(w)), h: Math.max(60, Math.round(h)) };
}

const xesc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!));
const absRef = (g: Range) => `$${colName(g.c1)}$${g.r1 + 1}` + (g.r1 === g.r2 && g.c1 === g.c2 ? '' : `:$${colName(g.c2)}$${g.r2 + 1}`);
const fRef = (wb: Workbook, r: Ref) => `${quoteSheet(wb.sheetName(r.sheet))}!${absRef(r.range)}`;

function chartXml(wb: Workbook, ch0: Chart): string {
  // bỏ tham chiếu tới sheet không còn
  const live = new Set(wb.sheetIds());
  const ok = (r?: Ref) => (r && live.has(r.sheet) ? r : undefined);
  const ch: Chart = { ...ch0, series: ch0.series.filter((x) => ok(x.vals)).map((x) => ({ vals: x.vals, name: ok(x.name), cats: ok(x.cats) })) };
  // bộ đệm ghi đủ mọi ô (kể cả hàng ẩn) để chỉ số khớp với vùng <c:f>
  const d = {
    cats: ch.series[0]?.cats ? (cellsOf(wb, ch.series[0].cats, 'text', true) as string[]) : [],
    series: ch.series.map((x, i) => ({ name: (x.name ? String(cellsOf(wb, x.name, 'text', true)[0] ?? '') : '') || `Chuỗi ${i + 1}`, values: cellsOf(wb, x.vals, 'num', true) as (number | null)[], cats: x.cats ? (cellsOf(wb, x.cats, 'text', true) as string[]) : [] })),
  };
  const pie = ch.type === 'pie' || ch.type === 'doughnut';
  const colorOfI = (i: number) => ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300'][i % 10];
  const strCache = (vals: string[]) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${xesc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (vals: (number | null)[]) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache>`;
  const dl = ch.labels ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${pie ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${pie ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>` : '';
  const ser = ch.series.map((s, i) => {
    const sd = d.series[i];
    const fill = (c: string) => `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`;
    const line = ch.type === 'line';
    const sp = line ? `<c:spPr><a:ln w="28575" cap="rnd">${fill(colorOfI(i))}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>` : pie ? '' : `<c:spPr>${fill(colorOfI(i))}</c:spPr>`;
    const dpt = pie ? d.cats.map((_, j) => `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/><c:spPr>${fill(colorOfI(j))}<a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '';
    const inv = !line && !pie && ch.type !== 'area' ? '<c:invertIfNegative val="0"/>' : '';
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
      (s.name ? `<c:tx><c:strRef><c:f>${xesc(fRef(wb, s.name))}</c:f>${strCache([sd.name])}</c:strRef></c:tx>` : '') +
      sp + inv + dpt + dl +
      (s.cats ? `<c:cat><c:strRef><c:f>${xesc(fRef(wb, s.cats))}</c:f>${strCache(sd.cats)}</c:strRef></c:cat>` : '') +
      `<c:val><c:numRef><c:f>${xesc(fRef(wb, s.vals))}</c:f>${numCache(sd.values)}</c:numRef></c:val>` +
      (line ? '<c:smooth val="0"/>' : '') + '</c:ser>';
  }).join('');
  const ax = (horiz: boolean) => `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'b' : 'l'}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const ids = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  let plot = '';
  switch (ch.type) {
    case 'col': case 'colStack': case 'bar': {
      const stacked = ch.type === 'colStack';
      plot = `<c:barChart><c:barDir val="${ch.type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${stacked ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${ser}<c:gapWidth val="150"/>${stacked ? '<c:overlap val="100"/>' : ''}${ids}</c:barChart>${ax(ch.type === 'bar')}`;
      break;
    }
    case 'line': plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/>${ids}</c:lineChart>${ax(false)}`; break;
    case 'area': plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}${ids}</c:areaChart>${ax(false)}`; break;
    case 'pie': plot = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`; break;
    case 'doughnut': plot = `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`; break;
  }
  const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="vi-VN" sz="1400" b="1"/><a:t>${xesc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = ch.legend === 'none' ? '' : `<c:legend><c:legendPos val="${ch.legend ?? 'b'}"/><c:overlay val="0"/></c:legend>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`;
}

/** Tìm ô + lệch (EMU) tại khoảng cách px tính từ đầu trục */
function locate(offset: (i: number) => number, px: number, max: number) {
  let lo = 0, hi = max;
  while (lo < hi) { const m = (lo + hi + 1) >> 1; if (offset(m) <= px) lo = m; else hi = m - 1; }
  return { i: lo, off: Math.round((px - offset(lo)) * EMU) };
}

async function addChartsToZip(buf: Uint8Array, wb: Workbook): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buf);
  const wbXml = (await zip.file('xl/workbook.xml')!.async('string'));
  const rels = await relsOf(zip, 'xl/workbook.xml');
  let ct = await zip.file('[Content_Types].xml')!.async('string');
  let chartNo = 0, drawNo = 0;
  for (const m of wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const a = attrs(m[1]);
    const id = wb.hf.getSheetId(a.name);
    if (id === undefined) continue;
    const charts = wb.meta(id).charts.filter((ch) => ch.series.some((x) => wb.sheetIds().includes(x.vals.sheet)));
    if (!charts.length) continue;
    const spath = rels.get(a.id);
    if (!spath) continue;
    const meta = wb.meta(id);
    drawNo++;
    const dpath = `xl/drawings/drawing${drawNo}.xml`;
    let anchors = '';
    const drels: string[] = [];
    charts.forEach((ch, i) => {
      chartNo++;
      const cpath = `xl/charts/chart${chartNo}.xml`;
      zip.file(cpath, chartXml(wb, ch));
      ct = ct.replace('</Types>', `<Override PartName="/${cpath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
      drels.push(`<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNo}.xml"/>`);
      const x1 = meta.colW.offset(ch.anchor.c) + ch.anchor.dx, y1 = meta.rowH.offset(ch.anchor.r) + ch.anchor.dy;
      const to = { c: locate((k) => meta.colW.offset(k), x1 + ch.w, 16383), r: locate((k) => meta.rowH.offset(k), y1 + ch.h, 1048575) };
      anchors += `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${ch.anchor.c}</xdr:col><xdr:colOff>${Math.round(ch.anchor.dx * EMU)}</xdr:colOff><xdr:row>${ch.anchor.r}</xdr:row><xdr:rowOff>${Math.round(ch.anchor.dy * EMU)}</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>${to.c.i}</xdr:col><xdr:colOff>${to.c.off}</xdr:colOff><xdr:row>${to.r.i}</xdr:row><xdr:rowOff>${to.r.off}</xdr:rowOff></xdr:to>` +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="${xesc(ch.title || 'Biểu đồ ' + (i + 1))}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
    });
    zip.file(dpath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`);
    zip.file(relsPath(dpath), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drels.join('')}</Relationships>`);
    ct = ct.replace('</Types>', `<Override PartName="/${dpath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
    // liên kết sheet → bản vẽ
    const srp = relsPath(spath);
    let srel = (await zip.file(srp)?.async('string')) ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const rid = 'rIdXLDraw1';
    srel = srel.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawNo}.xml"/></Relationships>`);
    zip.file(srp, srel);
    let sx = await zip.file(spath)!.async('string');
    if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    const tag = `<drawing r:id="${rid}"/>`;
    const at = ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst', '</worksheet>'].map((t) => sx.indexOf(t)).filter((i) => i >= 0).sort((p, q) => p - q)[0];
    sx = sx.slice(0, at) + tag + sx.slice(at);
    zip.file(spath, sx);
  }
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// =================== thiết lập in ===================
function printIn(ws: ExcelJS.Worksheet): PrintSetup | null {
  const pgs: any = ws.pageSetup ?? {};
  const hf: any = (ws as any).headerFooter ?? {};
  const tr = /\$?(\d+):\$?(\d+)/.exec(String(pgs.printTitlesRow ?? ''));
  const tc = /\$?([A-Z]+):\$?([A-Z]+)/i.exec(String(pgs.printTitlesColumn ?? ''));
  const m = pgs.margins;
  const mg = m ? { t: +m.top, b: +m.bottom, l: +m.left, r: +m.right, h: +m.header, f: +m.footer } : MARGIN_PRESETS.normal;
  const preset = (Object.entries(MARGIN_PRESETS).find(([, v]) => JSON.stringify(v) === JSON.stringify(mg))?.[0] ?? 'custom') as NPrint['margins'];
  const fit = pgs.fitToPage ? { w: pgs.fitToWidth ?? 1, h: pgs.fitToHeight ?? 1 } : null;
  const p: NPrint = {
    ...normPrint(null),
    orientation: pgs.orientation === 'landscape' ? 'landscape' : 'portrait',
    paper: paperByCode(pgs.paperSize ?? 9),
    gridlines: !!pgs.showGridLines,
    headings: !!pgs.showRowColHeaders,
    margins: preset, mg,
    scale: pgs.scale && !pgs.fitToPage ? +pgs.scale : 100,
    fit,
    titleRowRange: tr ? [+tr[1] - 1, +tr[2] - 1] : null,
    titleColRange: tc ? [colIdx(tc[1]), colIdx(tc[2])] : null,
    center: { h: !!pgs.horizontalCentered, v: !!pgs.verticalCentered },
    firstPage: pgs.useFirstPageNumber || (pgs.firstPageNumber && pgs.firstPageNumber !== 1) ? +pgs.firstPageNumber : null,
    order: pgs.pageOrder === 'overThenDown' ? 'over' : 'down',
    header: hf.oddHeader ?? '',
    footer: hf.oddFooter ?? '',
    area: pgs.printArea ? parseRef(String(pgs.printArea).split(/[,&]/)[0]) : null,
  };
  const def = normPrint(null);
  const same = JSON.stringify({ ...p, margins: def.margins, mg: def.mg }) === JSON.stringify(def) && preset === 'normal';
  return same ? null : storePrint(p);
}
function printOut(ws: ExcelJS.Worksheet, ps0: PrintSetup) {
  const ps = normPrint(ps0);
  ws.pageSetup = {
    ...ws.pageSetup,
    orientation: ps.orientation, paperSize: PAPERS[ps.paper].code as any, showGridLines: ps.gridlines, showRowColHeaders: ps.headings,
    fitToPage: !!ps.fit, ...(ps.fit ? { fitToWidth: ps.fit.w, fitToHeight: ps.fit.h } : { scale: ps.scale }),
    margins: { left: ps.mg.l, right: ps.mg.r, top: ps.mg.t, bottom: ps.mg.b, header: ps.mg.h, footer: ps.mg.f },
    horizontalCentered: ps.center.h, verticalCentered: ps.center.v,
    pageOrder: ps.order === 'over' ? 'overThenDown' : 'downThenOver',
    ...(ps.firstPage ? { firstPageNumber: ps.firstPage } : {}),
    ...(ps.titleRowRange ? { printTitlesRow: `${ps.titleRowRange[0] + 1}:${ps.titleRowRange[1] + 1}` } : {}),
    ...(ps.titleColRange ? { printTitlesColumn: `${colName(ps.titleColRange[0])}:${colName(ps.titleColRange[1])}` } : {}),
    ...(ps.area ? { printArea: rangeToA1(ps.area) } : {}),
  } as any;
  if (ps.header || ps.footer) (ws as any).headerFooter = { ...(ws as any).headerFooter, oddHeader: ps.header || undefined, oddFooter: ps.footer || undefined };
}
function colIdx(s: string) { let c = 0; for (const ch of s.toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; return c - 1; }
/** Ngắt trang theo cột (ExcelJS chỉ ghi được ngắt theo hàng) */
async function addColBreaks(buf: Uint8Array, cols: number[][]): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buf);
  for (let i = 0; i < cols.length; i++) {
    if (!cols[i].length) continue;
    const path = `xl/worksheets/sheet${i + 1}.xml`;
    let xml = await zip.file(path)?.async('string');
    if (!xml) continue;
    const blk = `<colBreaks count="${cols[i].length}" manualBreakCount="${cols[i].length}">${cols[i].map((c) => `<brk id="${c}" max="1048575" man="1"/>`).join('')}</colBreaks>`;
    const after = /<\/rowBreaks>|<rowBreaks[^>]*\/>/.exec(xml) ?? /<\/headerFooter>|<headerFooter[^>]*\/>/.exec(xml) ?? /<pageSetup[^>]*\/>/.exec(xml) ?? /<pageMargins[^>]*\/>/.exec(xml);
    xml = after ? xml.slice(0, after.index + after[0].length) + blk + xml.slice(after.index + after[0].length) : xml.replace('</worksheet>', blk + '</worksheet>');
    zip.file(path, xml);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** Chỉ nhận liên kết an toàn: web, email, tệp, vị trí trong sổ (chặn javascript:, data:…) */
export function safeLink(u: string): boolean { return /^(https?:|mailto:|file:|#|\\\\|[A-Za-z]:\\)/i.test(u.trim()) || (!/^[a-z][\w+.-]*:/i.test(u.trim()) && !!u.trim()); }

/** ExcelJS ghi location="#Sheet!A1" — Excel cần không có dấu # */
async function fixLocations(buf: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buf);
  for (const f of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
    const xml = await zip.file(f)!.async('string');
    if (xml.includes('location="#')) zip.file(f, xml.replace(/location="#/g, 'location="'));
  }
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
