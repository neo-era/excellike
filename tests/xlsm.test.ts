import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

/** dựng tệp .xlsm giả lập: 3 trang (1 veryHidden), codeName, vbaProject.bin + chữ ký, khai báo kiểu `Default bin` hoặc `Override` */
async function makeXlsm(ctMode: 'default' | 'override', withCache = true) {
  const wb = Workbook.create(); const s = wb.sheetIds()[0];
  wb.renameSheet(s, 'Tổng hợp');
  wb.setInput(s, 0, 0, 'a'); wb.setInput(s, 0, 1, 'b');
  const s2 = wb.addSheet('Ẩn sâu'); wb.setInput(s2, 0, 0, '1');
  const s3 = wb.addSheet('In'); wb.setInput(s3, 0, 0, '2');
  const base = await writeXlsx(wb);
  const z = await JSZip.loadAsync(base);
  const vba = new Uint8Array(4096).map((_, i) => (i * 37 + 11) & 255);
  const sig = new Uint8Array(300).map((_, i) => (i * 13) & 255);
  z.file('xl/vbaProject.bin', vba); z.file('xl/vbaProjectSignature.bin', sig);
  z.file('xl/_rels/vbaProject.bin.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/></Relationships>');
  let ct = await z.file('[Content_Types].xml')!.async('string');
  ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
  ct = ct.replace('<Default ', ctMode === 'default' ? '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ' : '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ');
  ct = ct.replace('</Types>', '<Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/></Types>');
  z.file('[Content_Types].xml', ct);
  let rels = await z.file('xl/_rels/workbook.xml.rels')!.async('string');
  rels = rels.replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
  z.file('xl/_rels/workbook.xml.rels', rels);
  let wx = await z.file('xl/workbook.xml')!.async('string');
  wx = /<workbookPr\b/.test(wx) ? wx.replace(/<workbookPr\b/, '<workbookPr codeName="ThisWorkbook"') : wx.replace('<bookViews', '<workbookPr codeName="ThisWorkbook"/><bookViews');
  wx = wx.replace(/<sheet [^>]*name="Ẩn sâu"[^>]*?\/>/, (t) => (/state="/.test(t) ? t.replace(/state="[^"]*"/, 'state="veryHidden"') : t.replace('/>', ' state="veryHidden"/>')));
  z.file('xl/workbook.xml', wx);
  for (const [i, cn] of [[1, 'Sheet1'], [2, 'Sheet7'], [3, 'Sheet3']] as const) {
    const p = `xl/worksheets/sheet${i}.xml`;
    let x = await z.file(p)!.async('string');
    x = /<sheetPr\b/.test(x) ? x.replace(/<sheetPr\b/, `<sheetPr codeName="${cn}"`) : x.replace(/(<worksheet\b[^>]*>)/, `$1<sheetPr codeName="${cn}"/>`);
    if (i === 1) x = x.replace(/<c r="C1"[^>]*>[\s\S]*?<\/c>|(<\/row>)/, (m0, end) => (end ? `<c r="C1" t="str"><f>NoiChuoiDuyNhat(A1:B1)</f>${withCache ? '<v>a, b</v>' : ''}</c></row>` : m0));
    z.file(p, x);
  }
  return { bytes: await z.generateAsync({ type: 'uint8array' }), vba, sig };
}

describe('Giữ macro .xlsm', () => {
  it('ghi tệp mẫu cho kiểm thử giao diện (e2e/fixtures/macro.xlsm)', async () => {
    const { bytes } = await makeXlsm('default');
    mkdirSync(new URL('../e2e/fixtures/', import.meta.url), { recursive: true });
    writeFileSync(new URL('../e2e/fixtures/macro.xlsm', import.meta.url), bytes);
  });
  for (const mode of ['default', 'override'] as const) {
    it(`mở → sửa → lưu: vbaProject.bin giống từng byte, kiểu nội dung macro (${mode})`, async () => {
      const { bytes, vba, sig } = await makeXlsm(mode);
      const { wb, warnings } = await readXlsx(bytes);
      expect(wb.hasMacros()).toBe(true);
      expect(warnings.join(' ')).toMatch(/macro/i);
      wb.setInput(wb.sheetIds()[0], 5, 0, 'sửa');
      const out = await writeXlsx(wb);
      const z = await JSZip.loadAsync(out);
      expect(new Uint8Array(await z.file('xl/vbaProject.bin')!.async('uint8array'))).toEqual(vba);
      expect(new Uint8Array(await z.file('xl/vbaProjectSignature.bin')!.async('uint8array'))).toEqual(sig);
      expect(await z.file('xl/_rels/vbaProject.bin.rels')!.async('string')).toMatch(/vbaProjectSignature/);
      const ct = await z.file('[Content_Types].xml')!.async('string');
      expect(ct).toMatch(/sheet\.macroEnabled\.main\+xml/);
      expect(ct).not.toMatch(/spreadsheetml\.sheet\.main\+xml/);
      expect(ct).toMatch(mode === 'default' ? /<Default Extension="bin" ContentType="application\/vnd\.ms-office\.vbaProject"\/>/ : /<Override PartName="\/xl\/vbaProject\.bin" ContentType="application\/vnd\.ms-office\.vbaProject"\/>/);
      expect(await z.file('xl/_rels/workbook.xml.rels')!.async('string')).toMatch(/relationships\/vbaProject" Target="vbaProject\.bin"/);
      const wx = await z.file('xl/workbook.xml')!.async('string');
      expect(wx).toMatch(/<workbookPr[^>]*codeName="ThisWorkbook"/);
      expect(wx).toMatch(/name="Ẩn sâu"[^>]*state="veryHidden"|state="veryHidden"[^>]*name="Ẩn sâu"/);
      expect(await z.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(/<sheetPr[^>]*codeName="Sheet1"/);
      expect(await z.file('xl/worksheets/sheet2.xml')!.async('string')).toMatch(/<sheetPr[^>]*codeName="Sheet7"/);
      // mở lại vẫn còn
      const { wb: w2 } = await readXlsx(out);
      expect(w2.hasMacros()).toBe(true);
    });
  }
  it('codeName đi theo trang khi đổi tên / đổi thứ tự; trang sao chép không mang codeName', async () => {
    const { bytes } = await makeXlsm('override');
    const { wb } = await readXlsx(bytes);
    const [a, , c] = wb.sheetIds();
    wb.renameSheet(c, 'In mới');
    wb.moveSheet(c, 0);
    const cp = wb.copySheet(a);
    expect(wb.meta(cp).codeName).toBeUndefined();
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    const order = wb.sheetIds().map((id) => wb.sheetName(id));
    const xmlOf = async (n: string) => z.file(`xl/worksheets/sheet${order.indexOf(n) + 1}.xml`)!.async('string');
    expect(await xmlOf('In mới')).toMatch(/codeName="Sheet3"/);
    expect(await xmlOf('Tổng hợp')).toMatch(/codeName="Sheet1"/);
    expect(await xmlOf(wb.sheetName(cp))).not.toMatch(/codeName=/);
  });
  it('trang veryHidden: không hiện trong danh sách Bỏ ẩn', async () => {
    const { bytes } = await makeXlsm('override');
    const { wb } = await readXlsx(bytes);
    const id = wb.sheetIds().find((x) => wb.sheetName(x) === 'Ẩn sâu')!;
    expect(wb.meta(id).hidden).toBe(true);
    expect(wb.meta(id).veryHidden).toBe(true);
    expect(wb.unhideableSheetIds()).not.toContain(id);
  });
  it('ô dùng hàm VBA: hiện giá trị lưu sẵn, lưu lại giữ công thức + giá trị (không ghi #NAME?)', async () => {
    const { bytes } = await makeXlsm('override');
    const { wb } = await readXlsx(bytes);
    const s = wb.sheetIds()[0];
    expect(wb.value(s, 0, 2)).toBe('a, b');
    expect(wb.rawText(s, 0, 2)).toBe('=NoiChuoiDuyNhat(A1:B1)');
    const x = await (await JSZip.loadAsync(await writeXlsx(wb))).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(x).toMatch(/<c r="C1"[^>]*><f>NoiChuoiDuyNhat\(A1:B1\)<\/f><v>a, b<\/v><\/c>/i);
    // không có giá trị lưu sẵn → #NAME?
    const { wb: w2 } = await readXlsx((await makeXlsm('override', false)).bytes);
    expect((w2.value(w2.sheetIds()[0], 0, 2) as any)?.value).toBe('#NAME?');
  });
  it('chèn hàng: giá trị lưu sẵn của ô hàm VBA dời theo ô; hoàn tác trả về', async () => {
    const { bytes } = await makeXlsm('default');
    const { wb } = await readXlsx(bytes);
    const s = wb.sheetIds()[0];
    wb.insertRows(s, 0, 3);
    expect(wb.value(s, 3, 2)).toBe('a, b');
    wb.undo();
    expect(wb.value(s, 0, 2)).toBe('a, b');
  });
  it('lưu thành .xlsx (bỏ macro): không còn phần macro, kiểu nội dung thường', async () => {
    const { bytes } = await makeXlsm('default');
    const { wb } = await readXlsx(bytes);
    const z = await JSZip.loadAsync(await writeXlsx(wb, { macro: false }));
    expect(z.file('xl/vbaProject.bin')).toBeNull();
    const ct = await z.file('[Content_Types].xml')!.async('string');
    expect(ct).not.toMatch(/macroEnabled|vbaProject/);
    expect(await z.file('xl/_rels/workbook.xml.rels')!.async('string')).not.toMatch(/vbaProject/);
  });
  it('sổ thường (không macro) lưu ra không có gì về macro', async () => {
    const wb = Workbook.create();
    expect(wb.hasMacros()).toBe(false);
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    expect(await z.file('[Content_Types].xml')!.async('string')).not.toMatch(/macroEnabled|vbaProject/);
  });
  it('khai báo macro nhưng thiếu vbaProject.bin → coi như không có macro, có cảnh báo', async () => {
    const { bytes } = await makeXlsm('override');
    const z = await JSZip.loadAsync(bytes); z.remove('xl/vbaProject.bin');
    const { wb, warnings } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    expect(wb.hasMacros()).toBe(false);
    expect(warnings.join(' ')).toMatch(/macro/i);
  });
  const REAL = '/mnt/user-data/attachments/ccaa393e-e2fa-51aa-a3c2-ab7a8c80a09f/c0114642-TONG_HOP_DE_XUAT_DEN_LED_PA3_V2A_them4cot_DADATHANG_2139_1.xlsm';
  it.skipIf(!existsSync(REAL))('tệp thật đề xuất đèn LED: giữ macro, 7 trang veryHidden', async () => {
    const src = readFileSync(REAL);
    const { wb } = await readXlsx(src);
    expect(wb.hasMacros()).toBe(true);
    expect(wb.sheetIds().filter((id) => wb.meta(id).veryHidden).length).toBe(7);
    const out = await writeXlsx(wb);
    const a = await (await JSZip.loadAsync(src)).file('xl/vbaProject.bin')!.async('uint8array');
    const b = await (await JSZip.loadAsync(out)).file('xl/vbaProject.bin')!.async('uint8array');
    expect(Buffer.from(b).equals(Buffer.from(a))).toBe(true);
  }, 120000);
  it.skipIf(!existsSync(REAL2))('tệp thật: báo PivotTable sẽ thành bảng giá trị; tên hàm VBA giữ chữ khi lưu', async () => {
    const { wb, warnings } = await readXlsx(new Uint8Array(readFileSync(REAL2)));
    expect(warnings.join(' ')).toMatch(/2 PivotTable/);
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    const all = (await Promise.all(Object.keys(z.files).filter((f) => /worksheets\/sheet\d+\.xml$/.test(f)).map((f) => z.file(f)!.async('string')))).join('');
    expect(all).toContain('NoiChuoiDuyNhat(');
    expect(all).not.toContain('NOICHUOIDUYNHAT(');
  }, 120_000);
});
const REAL2 = '/mnt/user-data/attachments/ccaa393e-e2fa-51aa-a3c2-ab7a8c80a09f/c0114642-TONG_HOP_DE_XUAT_DEN_LED_PA3_V2A_them4cot_DADATHANG_2139_1.xlsm';
