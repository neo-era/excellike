import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { readXlsx } from '../src/core/xlsx';

// Tệp xuất từ phần mềm khác: khoảng trắng quanh dấu "=" (r:id ="rId1"), nháy đơn, thẻ có khoảng trắng thừa
describe('Đọc .xlsx viết “lạ”', () => {
  it('r:id ="rId1", t = \'s\' vẫn đọc đủ dữ liệu', async () => {
    const z = new JSZip();
    z.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
    z.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    z.file('xl/workbook.xml', '<?xml version="1.0"?><workbook xmlns:r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns= "http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id ="rId1" /></sheets></workbook>');
    z.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" /><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" /></Relationships>');
    z.file('xl/sharedStrings.xml', '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Tên trụ</t></si></sst>');
    z.file('xl/worksheets/sheet1.xml', '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" ></sheetView></sheetViews ><sheetData><row r="1" ><c r="A1" t = \'s\'  ><v>0</v></c><c r="B1" t="n"  ><v>12</v></c></row></sheetData></worksheet>');
    const { wb } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    const s = wb.sheetIds()[0];
    expect(wb.value(s, 0, 0)).toBe('Tên trụ');
    expect(wb.value(s, 0, 1)).toBe(12);
  });
});
