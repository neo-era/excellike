// Tương thích tệp dự toán thật: tính lặp (tham chiếu vòng có chủ ý) và liên kết sang tệp khác (giá trị lưu sẵn)
import JSZip from 'jszip';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { toVi } from '../src/core/locale';

const isCycle = (v: unknown) => !!v && typeof v === 'object' && (v as any).type === 'CYCLE';

function cycleBook() {
  const wb = Workbook.create();
  wb.setInput(0, 0, 0, '1000');                 // A1 chi phí xây dựng
  wb.setInput(0, 0, 1, '=A1+C1');               // B1 tổng mức = xây dựng + QLDA
  wb.setInput(0, 0, 2, '=ROUND(B1*10%;0)');      // C1 QLDA = 10% tổng mức (vòng)
  wb.setInput(0, 0, 3, '=B1*2');                // D1 phụ thuộc vòng
  wb.setInput(0, 1, 1, '=SUM(B1:C1)');          // B2 vùng chứa ô vòng
  return wb;
}

describe('Tính lặp (Enable iterative calculation)', () => {
  it('tắt: báo #CYCLE! như Excel khi chưa bật tính lặp', () => {
    const wb = cycleBook();
    expect(isCycle(wb.value(0, 0, 1))).toBe(true);
  });
  it('bật: hội tụ như Excel, cả ô phụ thuộc và vùng chứa ô vòng', () => {
    const wb = cycleBook();
    wb.setIterate({ on: true });
    expect(wb.value(0, 0, 2)).toBe(111);
    expect(wb.value(0, 0, 1)).toBe(1111);
    expect(wb.value(0, 0, 3)).toBe(2222);
    expect(wb.value(0, 1, 1)).toBe(1222);
    expect(wb.display(0, 0, 1).text).toBe('1111');
    expect(wb.rawText(0, 0, 2)).toBe('=ROUND(B1*10%;0)'); // công thức giữ nguyên
  });
  it('sửa ô đầu vào → tính lặp lại; hoàn tác trả lại kết quả cũ', () => {
    const wb = cycleBook();
    wb.setIterate({ on: true });
    wb.setInput(0, 0, 0, '2000');
    expect(wb.value(0, 0, 2)).toBe(222);
    expect(wb.value(0, 0, 1)).toBe(2222);
    wb.undo();
    expect(wb.value(0, 0, 1)).toBe(1111);
  });
  it('chèn hàng phía trên vẫn tính đúng; lưu/mở lại sổ giữ chế độ tính lặp', () => {
    const wb = cycleBook();
    wb.setIterate({ on: true });
    wb.insertRows(0, 0, 2);
    expect(wb.value(0, 2, 1)).toBe(1111);
    const w2 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(w2.iterate.on).toBe(true);
    expect(w2.value(0, 2, 1)).toBe(1111);
  });
  it('đọc calcPr iterate="1" từ tệp Excel và ghi lại khi lưu', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'xl-it-'));
    const f = join(dir, 'vong.xlsx');
    execFileSync('python3', ['-c', `
import openpyxl,sys
wb=openpyxl.Workbook(); ws=wb.active
ws['A1']=1000; ws['B1']='=A1+C1'; ws['C1']='=ROUND(B1*10%,0)'
wb.calculation.iterate=True
wb.save(sys.argv[1])`, f]);
    const { wb, warnings } = await readXlsx(readFileSync(f));
    expect(wb.iterate.on).toBe(true);
    expect(wb.value(0, 0, 1)).toBe(1111);
    expect(warnings.join(' ')).not.toMatch(/chưa hỗ trợ/);
    const out = await writeXlsx(wb);
    const z = await JSZip.loadAsync(out);
    expect(await z.file('xl/workbook.xml')!.async('string')).toMatch(/<calcPr[^>]*iterate="1"/);
  });
});

/** Tạo tệp có liên kết ngoài [1] (giống Excel ghi): externalLink1.xml chứa giá trị lưu sẵn */
async function externalFixture(): Promise<Uint8Array> {
  const wb = Workbook.create();
  wb.setInput(0, 0, 0, 'x');
  const base = await JSZip.loadAsync(await writeXlsx(wb));
  const sheet = await base.file('xl/worksheets/sheet1.xml')!.async('string');
  const cells = `<c r="B1"><f>'[1]TMĐT-2'!J118/10</f><v>170</v></c><c r="C1"><f>[1]TMDT!$H$9*2</f><v>10</v></c><c r="D1"><f>SUM('[1]TMĐT-2'!J118:J119)</f><v>1702</v></c>`;
  base.file('xl/worksheets/sheet1.xml', sheet.replace(/(<row r="1"[^>]*>[\s\S]*?)(<\/row>)/, `$1${cells}$2`));
  let wbx = await base.file('xl/workbook.xml')!.async('string');
  wbx = wbx.replace('</sheets>', '</sheets><externalReferences><externalReference r:id="rIdExt1"/></externalReferences>');
  base.file('xl/workbook.xml', wbx);
  let rels = await base.file('xl/_rels/workbook.xml.rels')!.async('string');
  rels = rels.replace('</Relationships>', '<Relationship Id="rIdExt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>');
  base.file('xl/_rels/workbook.xml.rels', rels);
  let ct = await base.file('[Content_Types].xml')!.async('string');
  ct = ct.replace('</Types>', '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>');
  base.file('[Content_Types].xml', ct);
  base.file('xl/externalLinks/externalLink1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="TMĐT-2"/><sheetName val="TMDT"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="118"><cell r="J118"><v>1700</v></cell></row><row r="119"><cell r="J119"><v>2</v></cell></row></sheetData><sheetData sheetId="1"><row r="9"><cell r="H9"><v>5</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>`);
  base.file('xl/externalLinks/_rels/externalLink1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///D:/Du%20an/TMDT-2.xlsx" TargetMode="External"/></Relationships>`);
  return base.generateAsync({ type: 'uint8array' });
}

describe('Liên kết sang tệp khác (external links)', () => {
  it('chuyển [1]Sheet!A1 viết liền thành dạng có nháy', () => {
    expect(toVi('=[1]TMDT!$H$9*2')).toBe("='[1]TMDT'!$H$9*2");
    expect(toVi("='[1]TMĐT-2'!J118/10")).toBe("='[1]TMĐT-2'!J118/10");
    expect(toVi('="[1]TMDT!A1"')).toBe('="[1]TMDT!A1"');
  });
  it('dùng giá trị lưu sẵn như Excel (không báo #REF!), sheet đệm không hiện ra', async () => {
    const { wb, warnings } = await readXlsx(await externalFixture());
    expect(wb.value(0, 0, 1)).toBe(170);
    expect(wb.value(0, 0, 2)).toBe(10);
    expect(wb.value(0, 0, 3)).toBe(1702);
    expect(wb.sheetIds().map((i) => wb.sheetName(i))).toEqual(['Sheet1']);
    expect(wb.rawText(0, 0, 2)).toMatch(/\[1\]TMDT'?!\$H\$9\*2/);
    expect(warnings.join(' ')).toMatch(/liên kết/);
  });
  it('lưu .xlsx giữ nguyên liên kết ngoài (Excel vẫn cập nhật được), mở lại vẫn đúng', async () => {
    const { wb } = await readXlsx(await externalFixture());
    const w2 = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    expect(w2.value(0, 0, 1)).toBe(170);
    const out = await writeXlsx(w2);
    const z = await JSZip.loadAsync(out);
    expect(await z.file('xl/workbook.xml')!.async('string')).toMatch(/<externalReferences>\s*<externalReference r:id="[^"]+"\/>/);
    expect(z.file('xl/externalLinks/externalLink1.xml')).toBeTruthy();
    expect(Object.keys(z.files).some((f) => /worksheets\/sheet\d+\.xml/.test(f) && f !== 'xl/worksheets/sheet1.xml')).toBe(false);
    const dir = mkdtempSync(join(tmpdir(), 'xl-ext-'));
    writeFileSync(join(dir, 'o.xlsx'), out);
    const py = execFileSync('python3', ['-c', 'import openpyxl,sys;wb=openpyxl.load_workbook(sys.argv[1]);print(wb.sheetnames, wb.worksheets[0]["B1"].value, len(wb._external_links))', join(dir, 'o.xlsx')]).toString().trim();
    expect(py).toMatch(/\['Sheet1'\] ='\[1\]TMĐT-2'!J118\/10 1/);
    const { wb: w3 } = await readXlsx(out);
    expect(w3.value(0, 0, 3)).toBe(1702);
  });
});

describe('Giá trị như Excel', () => {
  it('công thức tham chiếu ô trống cho 0 (không để trống)', () => {
    const wb = Workbook.create();
    wb.setInput(0, 0, 1, '=+A1');
    wb.setInput(0, 0, 2, '=Sheet1!A5');
    expect(wb.value(0, 0, 1)).toBe(0);
    expect(wb.display(0, 0, 2).text).toBe('0');
    expect(wb.value(0, 3, 3)).toBe(null);
  });
});
