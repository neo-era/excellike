import { describe, it, expect } from 'vitest';
import fs from 'fs';
import * as XLSX from '@e965/xlsx';
import { xlsToXlsx, parseBiffFormats, fixExtRefs, isXlsBytes } from '../src/core/xls';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { Workbook, numText, xlRound } from '../src/core/workbook';
import { tintHex } from '../src/core/color';

const BASIC = () => new Uint8Array(fs.readFileSync('tests/fixtures/basic.xls'));
const open = async (b: Uint8Array) => { const r = await xlsToXlsx(b); const x = await readXlsx(r.data); return { ...x, conv: r.warnings }; };
const val = (wb: Workbook, s: number, a: string) => { const p = XLSX.utils.decode_cell(a); const v: any = wb.value(s, p.r, p.c); return v && typeof v === 'object' ? v.value : v; };
const st = (wb: Workbook, s: number, a: string) => { const p = XLSX.utils.decode_cell(a); return wb.style(s, p.r, p.c) ?? {}; };
/** luồng Workbook của tệp .xls → sửa → đóng gói lại */
function patchStream(b: Uint8Array, fn: (s: Uint8Array) => Uint8Array): Uint8Array {
  const cfb: any = XLSX.CFB.read(b, { type: 'array' });
  const e = XLSX.CFB.find(cfb, 'Workbook');
  e.content = fn(new Uint8Array(e.content));
  e.size = e.content.length;
  return new Uint8Array(XLSX.CFB.write(cfb, { type: 'array' }) as any);
}
const rec = (t: number, body: number[]) => [t & 255, t >> 8, body.length & 255, body.length >> 8, ...body];
/** chèn bản ghi ngay sau bản ghi cuối cùng có kiểu `after` trong phần đầu sổ */
function insertAfterLast(s: Uint8Array, after: number, bytes: number[]): Uint8Array {
  const d = new DataView(s.buffer, s.byteOffset, s.byteLength);
  let p = 0, at = -1;
  while (p + 4 <= s.length) { const t = d.getUint16(p, true), l = d.getUint16(p + 2, true); if (t === after) at = p + 4 + l; if (t === 0x0A) break; p += 4 + l; }
  const out = new Uint8Array(s.length + bytes.length);
  out.set(s.subarray(0, at)); out.set(bytes, at); out.set(s.subarray(at), at + bytes.length);
  // dời vị trí bắt đầu các trang (BoundSheet8.lbPlyPos) theo số byte vừa chèn
  const o = new DataView(out.buffer); p = 0;
  while (p + 4 <= out.length) { const t = o.getUint16(p, true), l = o.getUint16(p + 2, true); if (t === 0x85) o.setUint32(p + 4, o.getUint32(p + 4, true) + bytes.length, true); if (t === 0x0A) break; p += 4 + l; }
  return out;
}

describe('Mở tệp Excel 97–2003 (.xls)', () => {
  it('nhận diện chữ ký OLE2; tệp hỏng báo lỗi rõ', async () => {
    expect(isXlsBytes(BASIC())).toBe(true);
    expect(isXlsBytes(new Uint8Array([0x50, 0x4b, 3, 4, 0, 0, 0, 0, 0]))).toBe(false);
    await expect(xlsToXlsx(new Uint8Array(100))).rejects.toThrow(/\.xls/);
  });
  it('giá trị + công thức tính lại đúng như Excel (số → chữ dùng dấu phẩy, ROUND làm sạch 15 chữ số)', async () => {
    const { wb } = await open(BASIC());
    const s = wb.sheetIdByName('Dữ liệu')!;
    expect(val(wb, s, 'B3')).toBe(5.5);
    expect(val(wb, s, 'B4')).toBe(129.953);
    expect(val(wb, s, 'B5')).toBe('2 x 55');
    expect(val(wb, s, 'B6')).toBe('lớn');
    expect(val(wb, s, 'B7')).toBe(42);
    expect(val(wb, s, 'B8')).toBe('#DIV/0!');
    expect(val(wb, s, 'B9')).toBe('1,75');
    expect(val(wb, s, 'B10')).toBe(true);
    expect(val(wb, s, 'B11')).toBe(46343);
    expect(val(wb, s, 'B12')).toBe('chữ thường');
    expect(wb.rawText(s, 2, 1)).toBe('=SUM(A3:A4)');
  });
  it('định dạng: phông, cỡ, đậm/nghiêng, màu chữ, màu nền (cả bảng màu đã đổi), viền, căn lề, xuống dòng, định dạng số', async () => {
    const { wb } = await open(BASIC());
    const s = wb.sheetIdByName('Dữ liệu')!;
    const a1: any = st(wb, s, 'A1');
    expect(a1.b).toBe(true);
    expect(a1.fs).toBe(14);
    expect((a1.fn ?? 'Times New Roman')).toBe('Times New Roman');
    expect(a1.fg?.toLowerCase()).toBe('#ff0000');
    expect(a1.bg?.toLowerCase()).toBe('#ffff00');
    expect(a1.bd).toMatchObject({ l: { s: 'thin' }, r: { s: 'medium' }, t: { s: 'double' }, b: { s: 'dashed' } });
    expect([a1.al, a1.va, a1.wr]).toEqual(['c', 'm', true]);
    const a3: any = st(wb, s, 'A3');
    expect(a3.i).toBe(true);
    expect(a3.fn).toBe('Arial');
    expect(a3.fs).toBe(10);
    expect(a3.bg?.toLowerCase()).toBe('#123456');
    expect(a3.fmt).toBe('#,##0.00');
    expect(wb.display(s, 2, 0).text).toBe('2,00');
    expect(st(wb, s, 'B11').fmt).toBe('dd/mm/yyyy');
    // ô trống có định dạng sẵn (gõ số vào sẽ hiện theo định dạng)
    expect(st(wb, s, 'B14')).toMatchObject({ fmt: '#,##0', b: true, al: 'c' });
    expect(wb.isRowHidden(s, 20)).toBe(true);
  });
  it('ô gộp giữ cả nội dung ô không phải ô đầu; cột/hàng: độ rộng, chiều cao, ẩn; cố định ô; trang ẩn; khổ in', async () => {
    const { wb } = await open(BASIC());
    const s = wb.sheetIdByName('Dữ liệu')!;
    const m = wb.meta(s);
    expect(m.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 2 }]);
    expect(wb.rawText(s, 0, 1)).toBe('ẩn trong ô gộp');
    expect(m.colW.base(1)).toBe(Math.round(20 * 8 + 5));
    expect(wb.isColHidden(s, 3)).toBe(true);
    expect(m.rowH.base(1)).toBe(40);
    expect(wb.isRowHidden(s, 5)).toBe(true);
    expect(m.freeze).toEqual({ rows: 2, cols: 1 });
    expect(wb.meta(wb.sheetIdByName('Phụ')!).hidden).toBe(true);
    expect(m.print?.orientation).toBe('landscape');
    expect(m.print?.paper).toBe('A3');
  });
  it('lưu thành .xlsx rồi mở lại: giữ giá trị, công thức, ô gộp (kể cả nội dung ô ẩn trong vùng gộp), định dạng', async () => {
    const { wb } = await open(BASIC());
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s = w2.sheetIdByName('Dữ liệu')!;
    expect(w2.rawText(s, 0, 1)).toBe('ẩn trong ô gộp');
    expect(w2.meta(s).merges.length).toBe(1);
    expect(val(w2, s, 'B4')).toBe(129.953);
    expect(st(w2, s, 'A3').bg?.toLowerCase()).toBe('#123456');
    expect(w2.meta(w2.sheetIdByName('Phụ')!).hidden).toBe(true);
  });
  it('lưu .xlsx giữ cột rộng đúng 9, hàng ẩn không có ô, đủ 13 kiểu nét viền của Excel', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'x');
    wb.meta(s).colW.set(3, Math.round(9 * 8 + 5));
    wb.hideRows(s, 5, 5, true);
    const kinds = ['hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot', 'dotted', 'dashed', 'double'] as const;
    kinds.forEach((k, i) => wb.setBorder(s, { r1: 10 + i, c1: 1, r2: 10 + i, c2: 1 }, 'bottom', k));
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0];
    expect(w2.meta(s2).colW.base(3)).toBe(77);
    expect(w2.isRowHidden(s2, 5)).toBe(true);
    kinds.forEach((k, i) => expect(w2.style(s2, 10 + i, 1)?.bd?.b?.s).toBe(k));
  });
  it('màu đầy đủ của Excel 2007+ (XFExt): RGB cho nền, màu chủ đề + sắc độ cho chữ', async () => {
    const b = BASIC();
    const pf = parseBiffFormats(new Uint8Array((XLSX.CFB.find(XLSX.CFB.read(b, { type: 'array' }), 'Workbook') as any).content));
    const ix = pf.sheets[0].ixfe.get(11 * 16384 + 1)!; // B12
    const ext = (type: number, xclr: number, tint: number, v: number[]) => [type & 255, type >> 8, 20, 0, xclr, 0, tint & 255, (tint >> 8) & 255, ...v, 0, 0, 0, 0, 0, 0, 0, 0];
    const body = [0x7D, 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ix & 255, ix >> 8, 0, 0, 2, 0,
      ...ext(4, 2, 0, [0x11, 0x22, 0x33, 0]), // nền: RGB #112233
      ...ext(13, 3, -16384 & 0xffff, [4, 0, 0, 0])]; // chữ: Accent1 (Office 2007: 4F81BD), tối 50%
    const patched = patchStream(b, (s) => {
      // mẫu nền đặc cho XF đó để có màu nền: sửa trực tiếp bản ghi XF số ix
      const d = new DataView(s.buffer, s.byteOffset, s.byteLength);
      let p = 0, k = -1;
      while (p + 4 <= s.length) { const t = d.getUint16(p, true), l = d.getUint16(p + 2, true); if (t === 0xE0 && ++k === ix) { const b2 = d.getUint32(p + 4 + 14, true); d.setUint32(p + 4 + 14, (b2 & 0x03ffffff) | (1 << 26), true); } if (t === 0x0A) break; p += 4 + l; }
      return insertAfterLast(s, 0xE0, rec(0x87D, body));
    });
    const { wb } = await open(patched);
    const s = wb.sheetIdByName('Dữ liệu')!;
    const b12: any = st(wb, s, 'B12');
    expect(b12.bg?.toLowerCase()).toBe('#112233');
    expect(b12.fg?.toUpperCase()).toBe('#' + tintHex('4F81BD', -16384 / 32767).toUpperCase());
  });
  it('tệp có mật khẩu mở (FilePass): báo không mở được', async () => {
    const b = patchStream(BASIC(), (s) => { const d = new DataView(s.buffer, s.byteOffset); const l = d.getUint16(2, true); const at = 4 + l; const out = new Uint8Array(s.length + 10); out.set(s.subarray(0, at)); out.set(rec(0x2F, [0, 0, 0, 0, 0, 0]), at); out.set(s.subarray(at), at + 10); return out; });
    await expect(xlsToXlsx(b)).rejects.toThrow(/mật khẩu/);
  });
  it('rất nhiều ô trống kẻ viền: giới hạn số ô giữ định dạng, có cảnh báo, không tràn bộ nhớ', async () => {
    const b = patchStream(BASIC(), (s) => {
      const d = new DataView(s.buffer, s.byteOffset, s.byteLength);
      let p = 0, lastEof = -1;
      while (p + 4 <= s.length) { const t = d.getUint16(p, true), l = d.getUint16(p + 2, true); if (t === 0x0A) lastEof = p; p += 4 + l; }
      const ixA1 = parseBiffFormats(s).sheets[0].ixfe.get(0)!;
      const recs: number[] = [];
      const rows = 20000; // 5,1 triệu ô trống kẻ viền (~10 MB)
      for (let r = 0; r < rows; r++) { const body = [r & 255, r >> 8, 0, 0]; for (let c = 0; c < 256; c++) body.push(ixA1 & 255, ixA1 >> 8); body.push(255, 0); recs.push(...rec(0xBE, body)); }
      const out = new Uint8Array(s.length + recs.length);
      out.set(s.subarray(0, lastEof)); out.set(recs, lastEof); out.set(s.subarray(lastEof), lastEof + recs.length);
      return out;
    });
    const t = Date.now();
    const r = await xlsToXlsx(b);
    expect(r.warnings.join(' ')).toMatch(/ô trống/);
    expect(Date.now() - t).toBeLessThan(60000);
    expect(process.memoryUsage().heapUsed).toBeLessThan(1.5e9);
  }, 120000);
  it('giữ chế độ Tính lặp (ITERATION, CALCCOUNT, DELTA) của tệp .xls; lưu .xlsx ghi iterate="1"', async () => {
    const b = patchStream(BASIC(), (s) => {
      const d = new DataView(s.buffer, s.byteOffset, s.byteLength);
      let p = 0, lastEof = -1;
      while (p + 4 <= s.length) { const t = d.getUint16(p, true), l = d.getUint16(p + 2, true); if (t === 0x0A) lastEof = p; p += 4 + l; }
      const f = new Uint8Array(new Float64Array([0.0005]).buffer);
      const recs = [...rec(0x11, [1, 0]), ...rec(0x0C, [50, 0]), ...rec(0x10, [...f])];
      const out = new Uint8Array(s.length + recs.length);
      out.set(s.subarray(0, lastEof)); out.set(recs, lastEof); out.set(s.subarray(lastEof), lastEof + recs.length);
      return out;
    });
    const { wb } = await open(b);
    expect(wb.iterate).toMatchObject({ on: true, count: 50, delta: 0.0005 });
    const JSZip = (await import('jszip')).default;
    const wx = await (await JSZip.loadAsync(await writeXlsx(wb))).file('xl/workbook.xml')!.async('string');
    expect(wx).toMatch(/<calcPr[^>]*iterate="1"/);
  });
  it('.xlsx đặt mật khẩu (OLE2 EncryptedPackage) báo đúng là có mật khẩu', async () => {
    const cfb: any = XLSX.CFB.utils.cfb_new();
    XLSX.CFB.utils.cfb_add(cfb, 'EncryptionInfo', new Uint8Array(64));
    XLSX.CFB.utils.cfb_add(cfb, 'EncryptedPackage', new Uint8Array(64));
    await expect(xlsToXlsx(new Uint8Array(XLSX.CFB.write(cfb, { type: 'array' }) as any))).rejects.toThrow(/mật khẩu/);
  });
  it('tệp hỏng (rác ghi đè giữa luồng dữ liệu) không bao giờ mở im lặng: báo lỗi hoặc cảnh báo', async () => {
    for (const frac of [0.2, 0.4, 0.5, 0.7, 0.9]) {
      const b = patchStream(BASIC(), (s) => { const c = new Uint8Array(s); const at = Math.floor(c.length * frac), n = Math.floor(c.length / 10); for (let i = 0; i < n; i++) c[at + i] = (i * 7919 + 13) & 255; return c; });
      let msg = '';
      try { msg = (await xlsToXlsx(b)).warnings.join(' '); } catch (e) { msg = String(e); }
      expect(msg).toMatch(/hỏng|Không đọc được/);
    }
    expect((await xlsToXlsx(BASIC())).warnings.join(' ')).not.toMatch(/hỏng/);
  });
  it('tham chiếu sang sổ ngoài: gắn [n] theo thứ tự gặp; chuỗi trong ngoặc kép giữ nguyên', () => {
    const ext = new Map([['TH-HEM', 2], ["O'Brien", 1]]), used = new Map<number, number>();
    expect(fixExtRefs("+'TH-HEM'!$R$7+A1", ext, used)).toBe("+'[1]TH-HEM'!$R$7+A1");
    expect(fixExtRefs('"TH-HEM!A1"&\'O\'\'Brien\'!B2', ext, used)).toBe('"TH-HEM!A1"&\'[2]O\'\'Brien\'!B2');
    expect(fixExtRefs('Local!A1', ext, used)).toBe('Local!A1');
  });
});

describe('Đúng như Excel (phát hiện khi đối chiếu tệp .xls thật)', () => {
  it('số → chữ: 15 chữ số có nghĩa, dấu thập phân phẩy', () => {
    expect(numText(0.55 * 100)).toBe('55');
    expect(numText(4.072)).toBe('4,072');
    expect(numText(1 / 3)).toBe('0,333333333333333');
    expect(numText(-1.5)).toBe('-1,5');
    expect(numText(1e20)).toBe('1E+20');
    expect(numText(123456789012345)).toBe('123456789012345');
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '=0,302*100&"% x "&"Gtb"');
    wb.setInput(s, 0, 1, '=CONCATENATE(1,18;"")');
    wb.setInput(s, 0, 2, '=LEFT(2,5;3)');
    expect([wb.value(s, 0, 0), wb.value(s, 0, 1), wb.value(s, 0, 2)]).toEqual(['30,2% x Gtb', '1,18', '2,5']);
  });
  it('ROUND/ROUNDUP/ROUNDDOWN làm sạch sai số dấu phẩy động như Excel', () => {
    expect(xlRound(519.81 * 0.25, 3, 'near')).toBe(129.953);
    expect(xlRound(1.005, 2, 'near')).toBe(1.01);
    expect(xlRound(-2.5, 0, 'near')).toBe(-3);
    expect(xlRound(0.1 + 0.2, 1, 'up')).toBe(0.3);
    expect(xlRound(0.1 + 0.7, 1, 'down')).toBe(0.8);
    expect(xlRound(1234.5, -2, 'near')).toBe(1200);
    expect(xlRound(1250, -2, 'near')).toBe(1300);
    expect(Object.is(xlRound(-0.1, 0, 'near'), 0)).toBe(true);
    expect(xlRound(1.5, 400, 'near')).toBe(1.5);
    expect(xlRound(1.5, -400, 'near')).toBe(0);
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, '=ROUND(519,81*0,25;3)'); wb.setInput(s, 0, 1, '=ROUNDUP(0,1+0,2;1)'); wb.setInput(s, 0, 2, '=ROUNDDOWN(-1,999;2)'); wb.setInput(s, 0, 3, '=ROUND(2,5)');
    expect([wb.value(s, 0, 0), wb.value(s, 0, 1), wb.value(s, 0, 2), wb.value(s, 0, 3)]).toEqual([129.953, 0.3, -1.99, 3]);
  });
  it('tên vùng rất lớn ($A$1:$IU$65535 — vùng in cũ của .xls) mở nhanh, không tràn bộ nhớ', () => {
    const b: any = Workbook.create().toJSON();
    b.sheets[0].cells = [[1, 2], [3, 4]];
    b.names = [{ name: 'VungIn', expr: `='${b.sheets[0].name}'!$A$1:$IU$65535` }];
    const t = Date.now();
    const wb = Workbook.fromJSON(b);
    expect(Date.now() - t).toBeLessThan(3000);
    const s = wb.sheetIds()[0];
    wb.setInput(s, 3, 300, '=SUM(VungIn)');
    expect(wb.value(s, 3, 300)).toBe(10);
    // vùng thật dài hơn giới hạn .xls (70000 hàng) giữ nguyên, không đổi thành nguyên cột
    const b2: any = Workbook.create().toJSON();
    b2.names = [{ name: 'Dai', expr: `='${b2.sheets[0].name}'!$A$1:$A$70000` }];
    const w2 = Workbook.fromJSON(b2);
    expect(String(w2.hf.getNamedExpression('Dai')?.expression)).toMatch(/\$A\$1:\$A\$70000/);
  });
  it('ô gộp: công thức ở ô không phải ô đầu vẫn còn sau khi lưu/mở lại (lỗi mất dữ liệu ở 1.0.2)', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 1, 0, '5'); wb.setInput(s, 1, 1, '7');
    wb.merge(s, { r1: 0, c1: 0, r2: 0, c2: 5 }, false);
    wb.hf.setCellContents({ sheet: s, row: 0, col: 5 }, [['=SUM(A2:B2)']]);
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const s2 = w2.sheetIds()[0];
    expect(w2.rawText(s2, 0, 5)).toBe('=SUM(A2:B2)');
    expect(w2.value(s2, 0, 5)).toBe(12);
    expect(w2.meta(s2).merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 5 }]);
  });
  it('căn giữa theo chiều dọc của tệp Excel (vertical="center") được đọc đúng', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'x'); wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { va: 'm' });
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    let x = await z.file('xl/styles.xml')!.async('string');
    x = x.replace(/vertical="(middle|center)"/g, 'vertical="center"');
    z.file('xl/styles.xml', x);
    const { wb: w2 } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    expect(w2.style(w2.sheetIds()[0], 0, 0)?.va).toBe('m');
  });
});
