import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/core/workbook';
import { buildPrintHtml } from '../src/core/print';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

const big = () => { const wb = Workbook.create(); const s = wb.sheetIds()[0]; for (let r = 0; r < 160; r++) wb.setInput(s, r, 0, 'Dòng ' + (r + 1)); return { wb, s }; };
describe('0.8 đầu/chân trang', () => {
  it('trang đầu khác, trang chẵn khác, ảnh &G', () => {
    const { wb, s } = big();
    const png = 'data:image/png;base64,iVBORw0KGgo=';
    wb.setPrint(s, { orientation: 'portrait', paper: 'A4', gridlines: false, header: '&CThường &P', hfFirst: { header: '&CBÌA', footer: '' }, hfEven: { header: '&CChẵn &P', footer: '' }, footer: '&L&G', hfImg: { fl: png } });
    const html = buildPrintHtml(wb, s);
    const heads = [...html.matchAll(/<div class="hf h"[^>]*>([\s\S]*?)<\/div>/g)].map((m) => m[1].replace(/<[^>]+>/g, ''));
    expect(heads[0]).toBe('BÌA');
    expect(heads[1]).toBe('Chẵn 2');
    expect(heads[2]).toBe('Thường 3');
    expect(html).toContain('<img class="hfi" src="data:image/png;base64,iVBORw0KGgo="');
  });
  it('.xlsx giữ differentFirst/differentOddEven', async () => {
    const { wb, s } = big();
    wb.setPrint(s, { orientation: 'portrait', paper: 'A4', gridlines: false, header: '&CA', hfFirst: { header: '&CB', footer: '&CC' }, hfEven: { header: '&CD', footer: '' } });
    const { wb: w2 } = await readXlsx(await writeXlsx(wb));
    const p = w2.meta(w2.sheetIds()[0]).print!;
    expect(p.hfFirst).toEqual({ header: '&CB', footer: '&CC' });
    expect(p.hfEven?.header).toBe('&CD');
  });
  it('kéo ngắt trang → ngắt thủ công mới', () => {
    const { wb, s } = big();
    wb.insertPageBreak(s, 40, 0);
    wb.movePageBreak(s, 'r', 40, 30);
    expect(wb.meta(s).breaks.rows).toEqual([30]);
    wb.movePageBreak(s, 'r', 55, 50); // ngắt tự động → thủ công
    expect(wb.meta(s).breaks.rows).toEqual([30, 50]);
    wb.undo();
    expect(wb.meta(s).breaks.rows).toEqual([30]);
  });
});
