import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { Workbook } from '../src/core/workbook';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { DEFAULT_THEME, paletteTints, parseTheme, tintHex } from '../src/core/color';

const near = (a: string, b: string, d = 2) => [0, 2, 4].every((i) => Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)) <= d);

describe('Bảng màu (C2)', () => {
  it('sổ mới: bộ màu Office 2013–2022 (Accent1 #4472C4), thứ tự lt1, dk1, lt2, dk2', () => {
    expect(DEFAULT_THEME.slice(0, 10)).toEqual(['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']);
  });
  it('sắc độ theo HLS: khớp bảng màu Excel (sai lệch ≤ 2/kênh)', () => {
    expect(near(tintHex('4472C4', 0.7999), 'D9E1F2')).toBe(true);
    expect(near(tintHex('4472C4', 0.3999), '8EA9DB')).toBe(true);
    expect(near(tintHex('4472C4', -0.2499), '305496')).toBe(true);
    expect(near(tintHex('ED7D31', 0.3999), 'F4B084')).toBe(true);
    expect(tintHex('000000', 0.5)).toBe('808080');
    expect(paletteTints('#FFFFFF')).toEqual([-0.05, -0.15, -0.25, -0.35, -0.5]);
    expect(paletteTints('#000000')).toEqual([0.5, 0.35, 0.25, 0.15, 0.05]);
    expect(paletteTints('#E7E6E6')).toEqual([-0.1, -0.25, -0.5, -0.75, -0.9]);
    expect(paletteTints('#4472C4')).toEqual([0.8, 0.6, 0.4, -0.25, -0.5]);
  });
  it('đọc theme1.xml: srgbClr và sysClr lastClr', () => {
    const xml = '<a:theme><a:themeElements><a:clrScheme name="X"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2><a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="123456"/></a:accent4><a:accent5><a:srgbClr val="654321"/></a:accent5><a:accent6><a:srgbClr val="ABCDEF"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>';
    expect(parseTheme(xml)!.slice(0, 5)).toEqual(['FEFEFE', '111111', 'EEEEEE', '222222', 'FF0000']);
  });
  it('ô tô màu theo chủ đề + sắc độ (theme="4" tint) hiện theo theme1.xml của tệp; lưu lại giữ bộ màu', async () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    wb.setInput(s, 0, 0, 'x'); wb.setStyle(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { bg: '#123456' });
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    let th = await z.file('xl/theme/theme1.xml')!.async('string');
    th = th.replace(/<a:accent1>[\s\S]*?<\/a:accent1>/, '<a:accent1><a:srgbClr val="C00000"/></a:accent1>');
    z.file('xl/theme/theme1.xml', th);
    let st = await z.file('xl/styles.xml')!.async('string');
    st = st.replace(/<fgColor rgb="FF123456"\/>/, '<fgColor theme="4" tint="0.39997558519241921"/>');
    st = st.replace('</styleSheet>', '<colors><mruColors><color rgb="FF00B050"/><color rgb="FF7030A0"/></mruColors></colors></styleSheet>');
    z.file('xl/styles.xml', st);
    const { wb: w2 } = await readXlsx(await z.generateAsync({ type: 'uint8array' }));
    const bg = w2.style(w2.sheetIds()[0], 0, 0).bg!;
    expect(near(bg.slice(1).toUpperCase(), tintHex('C00000', 0.39997558519241921), 0)).toBe(true);
    expect(w2.theme?.[4]).toBe('C00000');
    expect(w2.mru).toEqual(['#00B050', '#7030A0']);
    w2.addRecentColor('#ff00ff');
    const z2 = await JSZip.loadAsync(await writeXlsx(w2));
    expect(await z2.file('xl/theme/theme1.xml')!.async('string')).toContain('<a:accent1><a:srgbClr val="C00000"/></a:accent1>');
    expect(await z2.file('xl/styles.xml')!.async('string')).toContain('<mruColors><color rgb="FFFF00FF"/><color rgb="FF00B050"/><color rgb="FF7030A0"/></mruColors>');
    const w3 = Workbook.fromJSON(JSON.parse(JSON.stringify(w2.toJSON())));
    expect(w3.mru.length).toBe(3);
    expect(w3.theme?.[4]).toBe('C00000');
  });
  it('sổ mới lưu rồi mở lại vẫn là bộ màu Office 2013–2022 (không đổi sang bộ 2007 của ExcelJS)', async () => {
    const wb = Workbook.create();
    const z = await JSZip.loadAsync(await writeXlsx(wb));
    expect(await z.file('xl/theme/theme1.xml')!.async('string')).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>');
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    expect(w2.theme).toBe(null);
  });
  it('Màu Gần đây tối đa 10, không trùng, mới nhất đầu', () => {
    const wb = Workbook.create();
    for (let i = 0; i < 12; i++) wb.addRecentColor('#0000' + i.toString(16).padStart(2, '0'));
    wb.addRecentColor('#000005');
    expect(wb.mru.length).toBe(10);
    expect(wb.mru[0]).toBe('#000005');
    expect(new Set(wb.mru).size).toBe(10);
  });
});
