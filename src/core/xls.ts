// Mở tệp Excel 97–2003 (.xls, BIFF8): đọc giá trị + công thức bằng SheetJS, tự đọc định dạng
// (phông, màu, viền, nền, căn lề, định dạng số, cố định ô, khổ in…) từ bản ghi BIFF, rồi dựng
// một tệp .xlsx tương đương để đi qua đúng đường nhập .xlsx sẵn có (một bộ đọc duy nhất).
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { tintHex } from './color';
import { addRawMerges } from './xlsx';
import { colName, type Range } from './address';

/** bảng màu mặc định BIFF8 (chỉ số 8..63) */
const PAL8 = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
/** bộ màu chủ đề Office 2007 (tệp .xls do Excel 2007+ lưu mà không kèm chủ đề riêng), thứ tự lt1, dk1, lt2, dk2, accent1..6 */
const THEME2007 = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'];
const BORDER = ['', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
const PATTERN = ['', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
const HALIGN = [undefined, 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const VALIGN = ['top', 'middle', undefined, 'justify', 'distributed'];

interface Font { name: string; size: number; bold: boolean; italic: boolean; strike: boolean; underline?: string; icv: number; vert?: 'superscript' | 'subscript' }
interface Xf { font: number; fmt: number; locked: boolean; hidden: boolean; h: number; wrap: boolean; v: number; rot: number; indent: number; shrink: boolean; bd: number[]; bc: number[]; diag: number; fls: number; fore: number; back: number }
interface ExtBook { path: string; sheets: string[]; cells: Map<number, unknown>[] }
const ERRS: Record<number, string> = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A' };
/** đường dẫn mã hoá của SupBook (VirtualPath) → đường dẫn thường */
function virtPath(p: string): string {
  if (p.charCodeAt(0) !== 1) return p;
  let o = '';
  for (let i = 1; i < p.length; i++) {
    const c = p.charCodeAt(i);
    if (c === 1) { const v = p[++i]; o += v === '@' ? '\\\\' : v + ':\\'; }
    else if (c === 2 || c === 3) o += '\\';
    else if (c === 4) o += '..\\';
    else if (c === 5) { const n = p.charCodeAt(++i); o += p.slice(i + 1, i + 1 + n); i += n; }
    else if (c >= 6 && c <= 8) { /* thư mục khởi động / thư viện: bỏ */ }
    else o += p[i];
  }
  return o;
}
type XColor = { rgb?: string; theme?: number; tint?: number; icv?: number; auto?: boolean };
interface SheetFmt {
  name: string; ixfe: Map<number, number>; grid: boolean; zeros: boolean; rtl: boolean; frozen?: { x: number; y: number; top: number; left: number }; zoom?: number;
  defColW?: number; stdW?: number; defRowH?: number; setup?: { paper: number; scale: number; fitW: number; fitH: number; portrait?: boolean; valid: boolean; start?: number };
  zeroRows: number[];
  calc: { iterate?: boolean; count?: number; delta?: number; manual?: boolean };
  cols: { c1: number; c2: number; w: number; hidden: boolean; level: number }[];
  fit: boolean; hc: boolean; vc: boolean; header?: string; footer?: string; hBreaks: number[]; vBreaks: number[]; tab?: number; drawings: boolean; cf: boolean; dv: boolean;
}

/** chuỗi XLUnicodeString (cch 2 byte) hoặc ShortXLUnicodeString (cch 1 byte) */
function ustr(d: DataView, p: number, short: boolean): [string, number] {
  const cch = short ? d.getUint8(p) : d.getUint16(p, true);
  p += short ? 1 : 2;
  const flags = d.getUint8(p++);
  const hi = flags & 1;
  if (flags & 8) p += 2; // rich text runs count
  if (flags & 4) p += 4; // ext size
  const n = hi ? cch * 2 : cch;
  const bytes = new Uint8Array(d.buffer, d.byteOffset + p, Math.max(0, Math.min(n, d.byteLength - p)));
  return [new TextDecoder(hi ? 'utf-16le' : 'latin1').decode(bytes), p + n];
}

/** đọc các bản ghi định dạng từ luồng Workbook (BIFF8) */
export function parseBiffFormats(stream: Uint8Array) {
  const d = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const fonts: Font[] = [], xfs: Xf[] = [], fmts = new Map<number, string>(), xfext = new Map<number, Map<number, XColor>>();
  let palette = PAL8.slice();
  let theme: string[] | null = null, themeBytes: number[] | null = null;
  const bounds: { pos: number; name: string }[] = [];
  const sheets = new Map<number, SheetFmt>();
  const books: ExtBook[] = [];
  let curBook: ExtBook | null = null, curTab = -1;
  let cur: SheetFmt | null = null, depth = 0, biff8 = true;
  const visMemo = new Map<number, boolean>();
  // ô trống chỉ giữ khi kiểu khác kiểu ô mặc định (XF 15: phông, định dạng số, căn lề, nền, viền…); tối đa BLANK_CAP ô cho cả sổ
  const ctx = { blanks: 0, over: 0, visible: (ix: number) => { let v = visMemo.get(ix); if (v === undefined) { const x = xfs[ix], d = xfs[15]; v = !!x && (!d || keyXf(x) !== keyXf(d)); visMemo.set(ix, v); } return v; } };
  let p = 0, prevType = 0, damaged = false, sheetBofs = 0;
  const nonWs = new Set<number>();
  let sstCount = -1, sstChunks: Uint8Array[] | null = null;
  while (p + 4 <= d.byteLength) {
    const t = d.getUint16(p, true), len = d.getUint16(p + 2, true), s = p + 4;
    if (s + len > d.byteLength) { damaged = true; break; }
    const u8 = (o: number) => d.getUint8(s + o), u16 = (o: number) => d.getUint16(s + o, true), u32 = (o: number) => d.getUint32(s + o, true);
    switch (t) {
      case 0x809: { // BOF
        depth++;
        if (p === 0 && u16(0) !== 0x600) biff8 = false;
        if (depth === 1 && p > 0) { sheetBofs++; if (u16(2) !== 0x10) nonWs.add(p); }
        if (u16(2) === 0x10 && depth === 1) {
          const b = bounds.find((x) => x.pos === p);
          if (b) { cur = { name: b.name, ixfe: new Map(), grid: true, zeros: true, rtl: false, zeroRows: [], calc: {}, cols: [], fit: false, hc: false, vc: false, hBreaks: [], vBreaks: [], drawings: false, cf: false, dv: false }; sheets.set(p, cur); }
        }
        break;
      }
      case 0x0A: depth = Math.max(0, depth - 1); if (!depth) cur = null; break; // EOF
      case 0xFC: if (len >= 8) { sstCount = u32(4); sstChunks = [new Uint8Array(d.buffer, d.byteOffset + s + 8, len - 8)]; } break; // SST
      case 0x3C: if (prevType === 0xFC && sstChunks) sstChunks.push(new Uint8Array(d.buffer, d.byteOffset + s, len)); break; // Continue của SST
      case 0x85: { const [nm] = ustr(d, s + 6, true); bounds.push({ pos: u32(0), name: nm }); break; } // BoundSheet8
      case 0x31: { // Font
        const [nm] = ustr(d, s + 14, true);
        const g = u16(2), uls = u8(10), sss = u16(8);
        fonts.push({ name: nm, size: u16(0) / 20, italic: !!(g & 2), strike: !!(g & 8), bold: u16(6) >= 700, icv: u16(4), underline: uls === 1 ? 'single' : uls === 2 ? 'double' : uls === 0x21 ? 'singleAccounting' : uls === 0x22 ? 'doubleAccounting' : undefined, vert: sss === 1 ? 'superscript' : sss === 2 ? 'subscript' : undefined });
        break;
      }
      case 0x41E: { const [f] = ustr(d, s + 2, false); fmts.set(u16(0), f); break; } // Format
      case 0x1AE: { // SupBook: sổ ngoài (liên kết sang tệp khác)
        curBook = null;
        const ctab = u16(0), cch = u16(2);
        if (cch === 0x0401 || cch === 0x3A01 || !ctab || cch === 0) break;
        try {
          const hi = u8(4) & 1, n = hi ? cch * 2 : cch;
          const path = new TextDecoder(hi ? 'utf-16le' : 'latin1').decode(new Uint8Array(d.buffer, d.byteOffset + s + 5, n));
          let q = s + 5 + n; const sheetsX: string[] = [];
          for (let i = 0; i < ctab && q < s + len; i++) { const [nm, q2] = ustr(d, q, false); sheetsX.push(nm); q = q2; }
          curBook = { path: virtPath(path), sheets: sheetsX, cells: sheetsX.map(() => new Map()) };
          books.push(curBook);
        } catch { /* bản ghi lạ: bỏ */ }
        break;
      }
      case 0x59: curTab = curBook ? u16(2) : -1; break; // XCT
      case 0x5A: { // CRN: giá trị lưu sẵn của ô ở sổ ngoài
        if (!curBook || curTab < 0 || !curBook.cells[curTab]) break;
        const last = u8(0), first = u8(1), row = u16(2); let q = s + 4;
        for (let c = first; c <= last && q < s + len; c++) {
          const ty = d.getUint8(q++); let v: unknown = null;
          if (ty === 1) { v = d.getFloat64(q, true); q += 8; }
          else if (ty === 2) { const [str, q2] = ustr(d, q, false); v = str; q = q2; }
          else if (ty === 4) { v = !!d.getUint8(q); q += 8; }
          else if (ty === 16) { v = { e: ERRS[d.getUint8(q)] ?? '#N/A' }; q += 8; }
          else q += 8;
          if (v !== null) curBook.cells[curTab].set(row * 16384 + c, v);
        }
        break;
      }
      case 0xE0: { // XF
        if (len < 20) break;
        const b1 = u32(10), b2 = u32(14), c = u16(18), al = u8(6), ind = u8(8), fl = u16(4);
        xfs.push({
          font: u16(0), fmt: u16(2), locked: !!(fl & 1), hidden: !!(fl & 2), h: al & 7, wrap: !!(al & 8), v: (al >> 4) & 7, rot: u8(7), indent: ind & 15, shrink: !!(ind & 16),
          bd: [b1 & 15, (b1 >> 4) & 15, (b1 >> 8) & 15, (b1 >> 12) & 15, (b2 >> 21) & 15], // trái, phải, trên, dưới, chéo
          bc: [(b1 >> 16) & 127, (b1 >> 23) & 127, b2 & 127, (b2 >> 7) & 127, (b2 >> 14) & 127], diag: (b1 >>> 30) & 3,
          fls: (b2 >>> 26) & 63, fore: c & 127, back: (c >> 7) & 127,
        });
        break;
      }
      case 0x92: { const n = u16(0); const pal = PAL8.slice(); for (let i = 0; i < n && i < 56; i++) pal[i] = hex(u8(2 + i * 4), u8(3 + i * 4), u8(4 + i * 4)); palette = pal; break; } // Palette
      case 0x87D: { // XFExt: màu đầy đủ (RGB/chủ đề) của Excel 2007+
        const ix = u16(14), n = u16(18); let q = 20; const m = new Map<number, XColor>();
        for (let i = 0; i < n && q + 4 <= len; i++) {
          const et = u16(q), cb = u16(q + 2);
          if ([4, 5, 7, 8, 9, 10, 11, 13].includes(et) && cb >= 20) {
            const xt = u16(q + 4), tint = d.getInt16(s + q + 6, true) / 32767, v = u32(q + 8);
            const col: XColor = xt === 2 ? { rgb: hex(u8(q + 8), u8(q + 9), u8(q + 10)), tint } : xt === 3 ? { theme: v, tint } : xt === 1 ? { icv: v, tint } : xt === 0 ? { auto: true } : {};
            if (xt <= 3) m.set(et, col);
          }
          q += Math.max(cb, 4);
        }
        if (m.size) xfext.set(ix, m);
        break;
      }
      case 0x896: { // Theme (gói zip chứa theme1.xml; có thể nối tiếp ở ContinueFrt12)
        const ver = u32(12);
        if (ver !== 0 && len > 16) themeBytes = Array.from(new Uint8Array(d.buffer, d.byteOffset + s + 16, len - 16));
        break;
      }
      case 0x87F: if (prevType === 0x896 || (prevType === 0x87F && themeBytes)) { if (themeBytes && len > 12) themeBytes.push(...new Uint8Array(d.buffer, d.byteOffset + s + 12, len - 12)); } break;
      default:
        if (cur) sheetRec(cur, t, len, d, s, ctx);
    }
    if (t === 0x3C && prevType === 0xFC) { /* giữ: các Continue liên tiếp của SST */ } else if (t !== 0x87F) prevType = t; else prevType = prevType === 0x896 ? 0x87F : prevType;
    p = s + len;
  }
  void theme;
  // cấu trúc bản ghi phải liền mạch: mỗi trang khai báo trỏ đúng BOF, BOF/EOF cân, đọc hết luồng
  if (sstChunks && sstCount >= 0 && !sstOk(sstChunks, sstCount)) damaged = true;
  if (stream.length && (depth !== 0 || bounds.some((b) => !sheets.has(b.pos) && !nonWs.has(b.pos)) || sheetBofs < bounds.length)) damaged = true;
  return { fonts, xfs, fmts, palette, xfext, themeBytes, sheets: [...sheets.values()], biff8, books, blanksDropped: ctx.over, damaged, sstCount };
}

/** đi hết bảng chuỗi dùng chung (SST + Continue): phải đọc vừa khít đúng số chuỗi khai báo */
function sstOk(chunks: Uint8Array[], count: number): boolean {
  let ci = 0, pos = 0;
  const left = () => (chunks[ci]?.length ?? 0) - pos;
  const next = () => { ci++; pos = 0; return ci < chunks.length; };
  const byte = (): number => { while (left() <= 0) if (!next()) throw 0; return chunks[ci][pos++]; };
  const skip = (n: number) => { while (n > 0) { if (left() <= 0 && !next()) throw 0; const k = Math.min(n, left()); pos += k; n -= k; } };
  try {
    for (let i = 0; i < count; i++) {
      const cch = byte() | (byte() << 8);
      const fl = byte();
      let hi = fl & 1;
      const runs = fl & 8 ? byte() | (byte() << 8) : 0;
      const ext = fl & 4 ? (byte() | (byte() << 8) | (byte() << 16) | (byte() << 24)) >>> 0 : 0;
      let n = cch;
      while (n > 0) {
        if (left() <= 0) { if (!next()) return false; hi = chunks[ci][pos++] & 1; continue; } // chuỗi nối sang Continue: byte cờ mới
        const k = Math.min(n, Math.floor(left() / (hi ? 2 : 1)));
        if (k === 0) return false;
        pos += k * (hi ? 2 : 1); n -= k;
      }
      skip(runs * 4); skip(ext);
    }
  } catch { return false; }
  return ci === chunks.length - 1 ? left() === 0 : ci >= chunks.length;
}
export const BLANK_CAP = 300_000;
const keyXf = (x: Xf) => JSON.stringify([x.font, x.fmt, x.h, x.wrap, x.v, x.rot, x.indent, x.shrink, x.bd, x.bc, x.diag, x.fls, x.fore, x.back, x.locked, x.hidden]);
function sheetRec(sh: SheetFmt, t: number, len: number, d: DataView, s: number, ctx: { blanks: number; over: number; visible: (ix: number) => boolean }) {
  const u8 = (o: number) => d.getUint8(s + o), u16 = (o: number) => (o + 2 <= len ? d.getUint16(s + o, true) : 0);
  const set = (r: number, c: number, x: number) => { if (c <= 255) sh.ixfe.set(r * 16384 + c, x); };
  const blank = (r: number, c: number, x: number) => {
    if (c > 255 || !ctx.visible(x)) return;
    if (ctx.blanks >= BLANK_CAP) { ctx.over++; return; }
    ctx.blanks++; sh.ixfe.set(r * 16384 + c, x);
  };
  switch (t) {
    case 0x201: if (len >= 6) blank(u16(0), u16(2), u16(4)); break; // Blank
    case 0x06: case 0x203: case 0x204: case 0x205: case 0x27E: case 0xFD: if (len >= 6) set(u16(0), u16(2), u16(4)); break; // Formula, Blank, Number, Label, BoolErr, RK, LabelSst
    case 0xBD: { const r = u16(0), c0 = u16(2), n = Math.floor((len - 6) / 6); for (let i = 0; i < n; i++) set(r, c0 + i, u16(4 + i * 6)); break; } // MulRk
    case 0xBE: { const r = u16(0), c0 = u16(2), n = Math.floor((len - 6) / 2); for (let i = 0; i < n; i++) blank(r, c0 + i, u16(4 + i * 2)); break; } // MulBlank
    case 0x23E: { const g = u16(0); sh.grid = !!(g & 2); sh.zeros = !!(g & 0x10); sh.rtl = !!(g & 0x40); if (g & 8) sh.frozen = sh.frozen ?? { x: 0, y: 0, top: 0, left: 0 }; else sh.frozen = undefined; break; } // Window2
    case 0x41: { if (sh.frozen) sh.frozen = { x: u16(0), y: u16(2), top: u16(4), left: u16(6) }; else sh.frozen = { x: u16(0), y: u16(2), top: u16(4), left: u16(6), split: true } as any; break; } // Pane
    case 0xA0: { const a = u16(0), b = u16(2); if (b) sh.zoom = Math.round((a / b) * 100); break; } // Scl
    case 0x55: sh.defColW = u16(0); break;
    // thiết lập tính toán (Excel ghi lặp lại ở mỗi trang): tính lặp, số lần, sai số, chế độ tính
    case 0x11: sh.calc.iterate = !!u16(0); break; // ITERATION
    case 0x0C: sh.calc.count = u16(0); break; // CALCCOUNT
    case 0x10: if (len >= 8) sh.calc.delta = d.getFloat64(s, true); break; // DELTA
    case 0x0D: sh.calc.manual = u16(0) === 0; break; // CALCMODE: 0 thủ công
    case 0x208: if (len >= 16 && ((u16(6) & 0x7fff) === 0 || (u16(12) & 0x20))) sh.zeroRows.push(u16(0)); break; // Row: cao 0 / ẩn
    case 0x7D: { const g = u16(8); sh.cols.push({ c1: u16(0), c2: Math.min(u16(2), 16383), w: u16(4) / 256, hidden: !!(g & 1), level: (g >> 8) & 7 }); break; } // ColInfo
    case 0x99: sh.stdW = u16(0) / 256; break;
    case 0x225: sh.defRowH = u16(2) / 20; break;
    case 0xA1: { const g = u16(10); sh.setup = { paper: u16(0), scale: u16(2), start: (g & 0x80) ? u16(4) : undefined, fitW: u16(6), fitH: u16(8), portrait: g & 0x40 ? undefined : !!(g & 2), valid: !(g & 4) }; break; }
    case 0x81: sh.fit = !!(u16(0) & 0x100); break;
    case 0x83: sh.hc = !!u16(0); break;
    case 0x84: sh.vc = !!u16(0); break;
    case 0x14: if (len > 0) sh.header = ustr(d, s, false)[0]; break;
    case 0x15: if (len > 0) sh.footer = ustr(d, s, false)[0]; break;
    case 0x1B: { const n = Math.min(u16(0), Math.floor((len - 2) / 6)); for (let i = 0; i < n; i++) sh.hBreaks.push(u16(2 + i * 6)); break; }
    case 0x1A: { const n = Math.min(u16(0), Math.floor((len - 2) / 6)); for (let i = 0; i < n; i++) sh.vBreaks.push(u16(2 + i * 6)); break; }
    case 0x862: if (len >= 20) { const icv = u8(16) & 127; if (icv !== 127 && icv !== 64) sh.tab = icv; } break; // SheetExt
    case 0x5D: { if (len >= 6 && ![0x19].includes(u16(4))) sh.drawings = true; break; } // Obj (bỏ qua ghi chú)
    case 0x1B0: sh.cf = true; break;
    case 0x1B2: sh.dv = true; break;
  }
  void u8;
}

function hex(r: number, g: number, b: number) { return [r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase(); }

/** đổi .xls (byte) thành .xlsx (byte) + cảnh báo */
export async function xlsToXlsx(buf: ArrayBuffer | Uint8Array): Promise<{ data: Uint8Array; warnings: string[] }> {
  const XLSX: any = await import('@e965/xlsx');
  const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  const warnings: string[] = [];
  let cfb: any;
  try { cfb = XLSX.CFB.read(bytes, { type: 'array' }); } catch { throw new Error('Tệp .xls bị hỏng hoặc không phải tệp Excel 97–2003.'); }
  // .xlsx đặt mật khẩu cũng là tệp OLE2 (EncryptedPackage)
  if (XLSX.CFB.find(cfb, 'EncryptedPackage') || XLSX.CFB.find(cfb, 'EncryptionInfo')) throw new Error('Tệp có mật khẩu mở — ExcelLike chưa mở được. Hãy bỏ mật khẩu trong Excel rồi mở lại.');
  const ent = XLSX.CFB.find(cfb, 'Workbook') ?? XLSX.CFB.find(cfb, 'Book');
  if (!ent?.content) throw new Error('Không tìm thấy dữ liệu sổ tính Excel 97–2003 trong tệp.');
  const stream = ent.content instanceof Uint8Array ? ent.content : new Uint8Array(ent.content);
  // tệp đặt mật khẩu mở (FILEPASS) → không đọc được
  if (hasRecord(stream, 0x2F)) throw new Error('Tệp .xls có mật khẩu mở — ExcelLike chưa mở được. Hãy bỏ mật khẩu trong Excel rồi mở lại.');
  // bản đưa cho SheetJS bỏ (đổi sang kiểu lạ, SheetJS bỏ qua): Theme (gói zip — SheetJS giải nén không giới hạn,
  // đọc riêng có trần) và ô trống Blank/MulBlank (SheetJS tạo đối tượng cho từng ô → tràn bộ nhớ với tệp kẻ viền cả trang;
  // định dạng ô trống đã do parseBiffFormats đọc từ luồng gốc)
  let feed = bytes;
  {
    const s2 = new Uint8Array(stream);
    const dv = new DataView(s2.buffer);
    let n = 0;
    for (let p = 0; p + 4 <= s2.length; p += 4 + dv.getUint16(p + 2, true)) { const t = dv.getUint16(p, true); if (t === 0x896 || t === 0x87F || t === 0x201 || t === 0xBE) { dv.setUint16(p, 0x00EF, true); n++; } } // 0xEF PhoneticInfo: SheetJS bỏ qua không báo
    if (n) { ent.content = s2; ent.size = s2.length; feed = new Uint8Array(XLSX.CFB.write(cfb, { type: 'array' }) as ArrayLike<number>); }
  }
  let src: any;
  try { src = XLSX.read(feed, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, cellDates: false, sheetStubs: false, WTF: false }); }
  catch (e) { console.warn('xls:', e); throw new Error('Không đọc được tệp .xls (tệp bị hỏng hoặc không đúng định dạng Excel 97–2003).'); }
  let fmt: ReturnType<typeof parseBiffFormats>;
  let broken = false;
  try { fmt = parseBiffFormats(stream); } catch { broken = true; fmt = parseBiffFormats(new Uint8Array(0)); warnings.push('Không đọc được định dạng ô của tệp .xls (tệp lỗi một phần): đã mở giá trị và công thức.'); }
  if (!fmt.biff8) fmt = { ...parseBiffFormats(new Uint8Array(0)), biff8: false }; // Excel 5/95: vị trí trường khác BIFF8 → bỏ định dạng
  // số chuỗi SheetJS đọc được phải khớp số khai báo trong SST (rác trong vùng chuỗi làm lệch)
  const nStr = Array.isArray(src.Strings) ? src.Strings.length : -1;
  if (fmt.sstCount >= 0 && nStr >= 0 && nStr !== fmt.sstCount) fmt.damaged = true;
  if (fmt.damaged) warnings.unshift('CẢNH BÁO: cấu trúc tệp .xls bị hỏng (bản ghi không liền mạch) — số liệu có thể sai hoặc thiếu. Hãy mở bằng Excel để sửa tệp, hoặc đối chiếu kỹ trước khi dùng.');
  if (fmt.blanksDropped) warnings.push(`Tệp có rất nhiều ô trống được kẻ viền/tô nền: chỉ giữ định dạng của ${BLANK_CAP.toLocaleString('vi-VN')} ô trống đầu tiên.`);
  // tham chiếu sang sổ ngoài: SheetJS chỉ ghi tên trang → gắn lại [n] theo SupBook có trang đó
  const local = new Set<string>(src.SheetNames);
  const extOf = new Map<string, number>();
  fmt.books.forEach((b, i) => b.sheets.forEach((nm) => { if (!local.has(nm) && !extOf.has(nm)) extOf.set(nm, i + 1); }));
  const usedBooks = new Map<number, number>(); // số SupBook → số [n] trong tệp .xlsx (theo thứ tự gặp)
  const extFix = (f: string) => fixExtRefs(f, extOf, usedBooks);
  if (!fmt.biff8 && !broken) warnings.push('Tệp Excel đời cũ (trước Excel 97): đã mở giá trị và công thức, định dạng có thể không đầy đủ.');
  // màu chủ đề
  let theme = THEME2007;
  if (fmt.themeBytes) {
    try {
      const z = await JSZip.loadAsync(new Uint8Array(fmt.themeBytes));
      const f = Object.keys(z.files).find((k) => /theme\/theme\d*\.xml$/i.test(k));
      const size = f ? Number((z.files[f] as any)._data?.uncompressedSize ?? Infinity) : Infinity;
      if (f && fmt.themeBytes.length < 2_000_000 && size < 5_000_000) { const { parseTheme } = await import('./color'); theme = parseTheme(await z.file(f)!.async('string')) ?? THEME2007; }
    } catch { /* giữ bộ mặc định */ }
  }
  const pal = (icv: number): string | undefined => {
    if (icv >= 8 && icv < 64) return fmt.palette[icv - 8];
    if (icv < 8) return PAL8[icv];
    return undefined; // 64/65/0x7FFF: màu tự động
  };
  const xcol = (x: XColor | undefined, icv: number): string | undefined => {
    if (x) {
      if (x.auto) return undefined;
      const base = x.rgb ?? (x.theme !== undefined ? theme[x.theme] : x.icv !== undefined ? pal(x.icv) : undefined);
      if (base) return x.tint ? tintHex(base, x.tint) : base;
    }
    return pal(icv);
  };
  const argb = (h: string | undefined) => (h ? { argb: 'FF' + h.toUpperCase() } : undefined);
  const fontOf = (i: number) => fmt.fonts[i >= 4 ? i - 1 : i] ?? fmt.fonts[0];
  const SSF = XLSX.SSF;
  const fmtStr = (id: number): string | undefined => {
    if (id === 0) return undefined;
    const s = fmt.fmts.get(id) ?? (SSF?.get_table?.()[id] as string | undefined);
    return s && s !== 'General' ? s : undefined;
  };
  const styleCache = new Map<number, Partial<ExcelJS.Style>>();
  const styleOf = (ix: number): Partial<ExcelJS.Style> | undefined => {
    if (styleCache.has(ix)) return styleCache.get(ix);
    const x = fmt.xfs[ix];
    if (!x) return undefined;
    const ext = fmt.xfext.get(ix);
    const f = fontOf(x.font);
    const st: Partial<ExcelJS.Style> = {};
    if (f) {
      const col = xcol(ext?.get(13), f.icv);
      st.font = { name: f.name, size: f.size, ...(f.bold ? { bold: true } : {}), ...(f.italic ? { italic: true } : {}), ...(f.strike ? { strike: true } : {}), ...(f.underline ? { underline: f.underline as any } : {}), ...(f.vert ? { vertAlign: f.vert } : {}), ...(col && col !== '000000' ? { color: argb(col) } : {}) };
    }
    const nf = fmtStr(x.fmt);
    if (nf) st.numFmt = nf;
    if (x.fls) {
      const fg = xcol(ext?.get(4), x.fore) ?? (x.fls === 1 ? undefined : '000000');
      const bg = xcol(ext?.get(5), x.back);
      if (x.fls === 1) { if (fg) st.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(fg) }; }
      else if (PATTERN[x.fls]) st.fill = { type: 'pattern', pattern: PATTERN[x.fls] as any, fgColor: argb(fg), ...(bg ? { bgColor: argb(bg) } : {}) };
    }
    const b: Partial<ExcelJS.Borders> = {};
    const side = (k: number, et: number) => (x.bd[k] && BORDER[x.bd[k]] ? { style: BORDER[x.bd[k]] as any, ...(xcol(ext?.get(et), x.bc[k]) ? { color: argb(xcol(ext?.get(et), x.bc[k])) } : {}) } : undefined);
    const L = side(0, 9), R = side(1, 10), T = side(2, 7), B = side(3, 8);
    if (L) b.left = L; if (R) b.right = R; if (T) b.top = T; if (B) b.bottom = B;
    if (x.diag && x.bd[4]) { const D = side(4, 11)!; b.diagonal = { ...D, up: !!(x.diag & 2), down: !!(x.diag & 1) } as any; }
    if (Object.keys(b).length) st.border = b as ExcelJS.Borders;
    const al: Partial<ExcelJS.Alignment> = {};
    if (HALIGN[x.h]) al.horizontal = HALIGN[x.h] as any;
    if (VALIGN[x.v]) al.vertical = VALIGN[x.v] as any;
    if (x.wrap) al.wrapText = true;
    if (x.shrink) al.shrinkToFit = true;
    if (x.indent) al.indent = x.indent;
    if (x.rot) al.textRotation = x.rot === 255 ? 'vertical' : x.rot > 90 ? -(x.rot - 90) : x.rot;
    if (Object.keys(al).length) st.alignment = al;
    if (!x.locked || x.hidden) st.protection = { locked: x.locked, ...(x.hidden ? { hidden: true } : {}) };
    styleCache.set(ix, st);
    return st;
  };

  const out = new ExcelJS.Workbook();
  if (src.Workbook?.WBProps?.date1904) (out.properties as any).date1904 = true;
  const meta = src.Workbook?.Sheets ?? [];
  const fmtByName = new Map(fmt.sheets.map((s) => [s.name, s]));
  let pics = 0, cfs = 0, dvs = 0;
  const merges: Range[][] = [];
  src.SheetNames.forEach((name: string, si: number) => {
    const ws = src.Sheets[name] ?? {};
    const sf = fmtByName.get(name);
    const hid = meta[si]?.Hidden;
    const o = out.addWorksheet(name, { state: hid === 2 ? 'veryHidden' : hid === 1 ? 'hidden' : 'visible' } as any);
    if (sf?.drawings) pics++;
    if (sf?.cf) cfs++;
    if (sf?.dv) dvs++;
    // cột / hàng
    if (sf?.cols.length) {
      for (const ci of sf.cols) for (let c = ci.c1; c <= ci.c2 && c <= 255; c++) { // .xls tối đa 256 cột (A..IV)
        const col = o.getColumn(c + 1); // ExcelJS bỏ cột rộng đúng 9 (coi là mặc định)
        col.width = ci.w === 9 ? 9 + 1 / 256 : ci.w; if (ci.hidden || ci.w === 0) col.hidden = true; if (ci.level) col.outlineLevel = ci.level;
      }
    } else {
      const cols: any[] = ws['!cols'] ?? [];
      cols.forEach((c, i) => { if (!c) return; const col = o.getColumn(i + 1); if (c.width ?? c.wch) col.width = c.width ?? c.wch; if (c.hidden) col.hidden = true; if (c.level) col.outlineLevel = c.level; });
    }
    const rows: any[] = ws['!rows'] ?? [];
    rows.forEach((r, i) => { if (!r) return; const row = o.getRow(i + 1); if (r.hpt) row.height = r.hpt; if (r.hidden || r.hpt === 0) { row.hidden = true; if (!r.hpt) row.height = 15; } if (r.level) row.outlineLevel = r.level; });
    const props: any = o.properties;
    if (sf?.stdW) props.defaultColWidth = sf.stdW; else if (sf?.defColW) props.defaultColWidth = sf.defColW + 0.71;
    if (sf?.defRowH) props.defaultRowHeight = sf.defRowH;
    if (sf?.tab !== undefined) { const h = pal(sf.tab); if (h) props.tabColor = { argb: 'FF' + h }; }
    // ô
    for (const k of Object.keys(ws)) {
      if (k[0] === '!') continue;
      const c = ws[k];
      if (!c) continue;
      if (!/^[A-I]?[A-Z]\d{1,5}$/.test(k)) continue; // .xls tối đa cột IV, hàng 65536
      const cell = o.getCell(k);
      const val = c.t === 'e' ? { error: c.w ?? '#N/A' } : c.t === 'b' ? !!c.v : c.t === 'n' ? c.v : c.t === 's' || c.t === 'str' ? String(c.v ?? '') : c.t === 'd' ? c.v : null;
      if (c.f) {
        const cv: any = { formula: extFix(c.f), result: val ?? undefined };
        if (c.F && c.F.split(':')[0] === k && c.F.includes(':')) { cv.shareType = 'array'; cv.ref = c.F; }
        cell.value = cv;
      } else if (c.l?.Target && typeof val === 'string') {
        cell.value = { text: val, hyperlink: c.l.Target, ...(c.l.Tooltip ? { tooltip: c.l.Tooltip } : {}) } as any;
      } else if (val !== null && val !== undefined) cell.value = val as any;
      if (c.c?.length) { const tx = c.c.map((x: any) => String(x.t ?? '').split('\u0000')[0]).join('\n').trim(); if (tx) cell.note = tx; }
    }
    // định dạng từng ô (kể cả ô trống có viền/nền)
    if (sf) for (const [key, ix] of sf.ixfe) {
      const st = styleOf(ix);
      if (!st || !Object.keys(st).length) continue;
      const r = Math.floor(key / 16384), c = key % 16384;
      o.getCell(r + 1, c + 1).style = st as ExcelJS.Style;
    }
    // ô gộp ghi thẳng vào XML: ExcelJS xoá nội dung các ô không phải ô đầu (Excel vẫn giữ)
    merges[si] = (ws['!merges'] ?? []).map((m: any) => ({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }));
    // khung nhìn
    const view: any = {};
    if (sf?.frozen && !(sf.frozen as any).split && (sf.frozen.x || sf.frozen.y)) Object.assign(view, { state: 'frozen', xSplit: sf.frozen.x, ySplit: sf.frozen.y });
    if (sf && !sf.grid) view.showGridLines = false;
    if (sf?.zoom && sf.zoom !== 100) view.zoomScale = sf.zoom;
    if (sf?.rtl) view.rightToLeft = true;
    if (sf && !sf.zeros) view.showZeros = false;
    if (Object.keys(view).length) o.views = [view];
    // in
    const mg = ws['!margins'];
    const ps: any = {};
    if (mg) ps.margins = { left: mg.left ?? 0.7, right: mg.right ?? 0.7, top: mg.top ?? 0.75, bottom: mg.bottom ?? 0.75, header: mg.header ?? 0.3, footer: mg.footer ?? 0.3 };
    if (sf?.setup?.valid) {
      if (sf.setup.paper) ps.paperSize = sf.setup.paper;
      if (sf.setup.portrait !== undefined) ps.orientation = sf.setup.portrait ? 'portrait' : 'landscape';
      if (sf.fit) { ps.fitToPage = true; ps.fitToWidth = sf.setup.fitW; ps.fitToHeight = sf.setup.fitH; } else if (sf.setup.scale && sf.setup.scale !== 100) ps.scale = sf.setup.scale;
      if (sf.setup.start) ps.firstPageNumber = sf.setup.start;
    }
    if (sf?.hc) ps.horizontalCentered = true;
    if (sf?.vc) ps.verticalCentered = true;
    if (Object.keys(ps).length) o.pageSetup = { ...o.pageSetup, ...ps };
    if (sf?.header || sf?.footer) o.headerFooter = { ...(sf.header ? { oddHeader: sf.header } : {}), ...(sf.footer ? { oddFooter: sf.footer } : {}) } as any;
    for (const r of sf?.hBreaks ?? []) if (r > 0) o.getRow(r).addPageBreak();
    for (const r of sf?.zeroRows ?? []) { const row = o.getRow(r + 1); row.hidden = true; if (!row.height) row.height = sf?.defRowH ?? 15; }
  });
  if (pics) warnings.push(`Hình ảnh, biểu đồ, nút bấm trong tệp .xls (${pics} trang) chưa được nhập.`);
  if (cfs) warnings.push(`Định dạng có điều kiện trong tệp .xls (${cfs} trang) chưa được nhập.`);
  if (dvs) warnings.push(`Kiểm tra dữ liệu (danh sách chọn…) trong tệp .xls (${dvs} trang) chưa được nhập.`);
  if ((cfb.FullPaths ?? []).some((x: string) => /_VBA_PROJECT_CUR/i.test(x))) warnings.push('Tệp .xls có macro VBA — macro không được giữ khi lưu thành .xlsx.');
  let data = new Uint8Array(await out.xlsx.writeBuffer());
  if (merges.some((x) => x?.length)) data = await addRawMerges(data, merges);
  // tên vùng (kể cả vùng in, tiêu đề in) ghi thẳng vào workbook.xml
  const names: any[] = src.Workbook?.Names ?? [];
  const defs: string[] = [];
  for (const n of names) {
    const nm = String(n.Name ?? '');
    const ref = extFix(String(n.Ref ?? '').replace(/^=/, ''));
    if (!nm || !ref || /^_xlfn\./i.test(nm) || /^#(NAME|REF)/.test(ref)) continue;
    if (/^_xlnm\./i.test(nm) && (n.Sheet === undefined || !/^_xlnm\.(Print_Area|Print_Titles)$/i.test(nm))) continue;
    defs.push(`<definedName name="${xe(nm)}"${n.Sheet !== undefined ? ` localSheetId="${n.Sheet}"` : ''}${n.Hidden ? ' hidden="1"' : ''}>${xe(ref)}</definedName>`);
  }
  const ext = [...usedBooks.keys()]; // đã theo thứ tự [1], [2]…
  // Excel ghi cùng thiết lập ở mọi trang; lấy trang có bật tính lặp nếu có
  const calc = fmt.sheets.find((x) => x.calc.iterate)?.calc ?? fmt.sheets.find((x) => x.calc.manual !== undefined)?.calc ?? {};
  const calcAttr = [calc.iterate ? `iterate="1" iterateCount="${calc.count || 100}" iterateDelta="${calc.delta && isFinite(calc.delta) ? calc.delta : 0.001}"` : '', calc.manual ? 'calcMode="manual"' : ''].filter(Boolean).join(' ');
  if (defs.length || ext.length || calcAttr) {
    const z = await JSZip.loadAsync(data);
    let wx = await z.file('xl/workbook.xml')!.async('string');
    let rx = await z.file('xl/_rels/workbook.xml.rels')!.async('string');
    let ct = await z.file('[Content_Types].xml')!.async('string');
    wx = wx.replace(/<definedNames>[\s\S]*?<\/definedNames>|<definedNames\/>/, '');
    let refs = '';
    ext.forEach((k, i) => {
      const b = fmt.books[k - 1], id = `rIdXls${i + 1}`, path = `xl/externalLinks/externalLink${i + 1}.xml`;
      const sd = b.cells.map((m, si) => {
        if (!m.size) return `<sheetData sheetId="${si}"/>`;
        const rows = new Map<number, string[]>();
        for (const [key, v] of [...m.entries()].sort((a, c) => a[0] - c[0])) {
          const r = Math.floor(key / 16384), c = key % 16384, a = colName(c) + (r + 1);
          const cell = typeof v === 'number' ? `<cell r="${a}"><v>${v}</v></cell>` : typeof v === 'boolean' ? `<cell r="${a}" t="b"><v>${v ? 1 : 0}</v></cell>` : typeof v === 'string' ? `<cell r="${a}" t="str"><v>${xe(v)}</v></cell>` : `<cell r="${a}" t="e"><v>${xe((v as any).e)}</v></cell>`;
          (rows.get(r) ?? rows.set(r, []).get(r)!).push(cell);
        }
        return `<sheetData sheetId="${si}">${[...rows].map(([r, cs]) => `<row r="${r + 1}">${cs.join('')}</row>`).join('')}</sheetData>`;
      }).join('');
      z.file(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames>${b.sheets.map((x) => `<sheetName val="${xe(x)}"/>`).join('')}</sheetNames><sheetDataSet>${sd}</sheetDataSet></externalBook></externalLink>`);
      const target = b.path.replace(/\\/g, '/').split('/').map((x) => encodeURIComponent(x).replace(/%3A/gi, ':')).join('/');
      z.file(`xl/externalLinks/_rels/externalLink${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${xe(/^[A-Za-z]:\//.test(target) ? 'file:///' + target : target)}" TargetMode="External"/></Relationships>`);
      rx = rx.replace('</Relationships>', () => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink${i + 1}.xml"/></Relationships>`);
      ct = ct.replace('</Types>', () => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>`);
      refs += `<externalReference r:id="${id}"/>`;
    });
    const tail = (refs ? `<externalReferences>${refs}</externalReferences>` : '') + (defs.length ? `<definedNames>${defs.join('')}</definedNames>` : '');
    wx = wx.replace('</sheets>', () => `</sheets>${tail}`);
    if (calcAttr) wx = /<calcPr\b/.test(wx) ? wx.replace(/<calcPr\b([^>]*?)(\/?)>/, (_m, a, sl) => `<calcPr${a} ${calcAttr}${sl}>`) : wx.replace('</workbook>', () => `<calcPr calcId="191029" ${calcAttr}/></workbook>`);
    if (!/xmlns:r=/.test(wx)) wx = wx.replace('<workbook ', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    z.file('xl/workbook.xml', wx); z.file('xl/_rels/workbook.xml.rels', rx); z.file('[Content_Types].xml', ct);
    data = await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  return { data, warnings };
}

/** gắn số sổ ngoài vào tham chiếu tên trang không có trong sổ: 'TH-HEM'!R7 → '[1]TH-HEM'!R7 (bỏ qua chuỗi trong ngoặc kép) */
export function fixExtRefs(f: string, extOf: Map<string, number>, used: Map<number, number>): string {
  if (!extOf.size) return f;
  return f.replace(/"(?:[^"]|"")*"|'((?:[^']|'')+)'!|([^\s'"!(),;=+\-*/&^<>:%{}]+)!/g, (m0, q?: string, b?: string) => {
    if (m0[0] === '"') return m0;
    const nm = (q ?? b ?? '').replace(/''/g, "'");
    const k = extOf.get(nm);
    if (!k) return m0;
    if (!used.has(k)) used.set(k, used.size + 1);
    return `'[${used.get(k)}]${nm.replace(/'/g, "''")}'!`;
  });
}
function xe(s: string) { return s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function hasRecord(stream: Uint8Array, type: number) {
  const d = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  let p = 0, n = 0;
  while (p + 4 <= d.byteLength && n++ < 5_000_000) { if (d.getUint16(p, true) === type) return true; p += 4 + d.getUint16(p + 2, true); }
  return false;
}
/** nhận diện tệp .xls (OLE2/CFB) theo chữ ký đầu tệp */
export function isXlsBytes(b: Uint8Array) { return b.length > 8 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0 && b[4] === 0xA1 && b[5] === 0xB1 && b[6] === 0x1A && b[7] === 0xE1; }
