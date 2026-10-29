// Đo thời gian mở tệp .xlsx lớn (6 trang × 6.000 hàng × 12 cột, 108.000 công thức, viền + định dạng số)
// và mở lại app với sổ đó. Chạy: node e2e/perf-open.mjs   (cần python3 + openpyxl để tạo tệp mẫu)
import { createRequire } from 'node:module';
import { execFileSync, spawn } from 'node:child_process';
import { existsSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const { chromium } = createRequire(import.meta.url)('playwright');
const f = join(mkdtempSync(join(tmpdir(), 'xl-perf-')), 'lon.xlsx');
execFileSync('python3', ['-c', `
import openpyxl, random, sys
from openpyxl.styles import Font, Border, Side
wb=openpyxl.Workbook(); random.seed(1); t=Side(style='thin'); bd=Border(left=t,right=t,top=t,bottom=t)
for si in range(6):
    ws=wb.active if si==0 else wb.create_sheet(); ws.title=f'Sheet {si+1}'
    ws.append(['STT','Phường','Gói','Mã','SL','Loại','Thông số','Đơn giá','Thành tiền','Ghi chú','VAT','Tổng'])
    for r in range(2,6002):
        ws.append([r-1,'Phú Thạnh','Gói 1','CAN49',random.randint(1,20),'KIENG','D110',random.randint(100000,9000000),f'=E{r}*H{r}','ghi chú',f'=I{r}*0.08',f'=I{r}+K{r}'])
        for c in range(1,13):
            ce=ws.cell(r,c); ce.border=bd; ce.font=Font(name='Times New Roman',size=11)
            if c in (8,9,11,12): ce.number_format='#,##0'
wb.save(sys.argv[1])`, f]);
const PORT = 4197;
const srv = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'pipe', detached: true, cwd: new URL('..', import.meta.url).pathname });
await new Promise((r) => { srv.stdout.on('data', (d) => { if (String(d).includes('Local')) r(); }); setTimeout(r, 6000); });
const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 1366, height: 800 } });
await ctx.addInitScript(() => { delete window.showOpenFilePicker; });
const p = await ctx.newPage();
await p.goto(`http://localhost:${PORT}/`); await p.waitForFunction(() => window.__xl);
await p.click('#tabs [data-tab="file"]'); await p.click('.bs-nav [data-bs="open"]');
const [fc] = await Promise.all([p.waitForEvent('filechooser'), p.click('#backstage [data-a="browse"]')]);
const t0 = Date.now();
await Promise.all([p.waitForEvent('load', { timeout: 300000 }), fc.setFiles(f)]);
await p.waitForFunction(() => window.__xl && window.__xl.wb.sheetIds().length >= 6, null, { timeout: 300000 });
const open = Date.now() - t0;
const t1 = Date.now();
await p.reload(); await p.waitForFunction(() => window.__xl && window.__xl.wb.sheetIds().length >= 6, null, { timeout: 300000 });
const reopen = Date.now() - t1;
const ok = (await p.evaluate(() => window.__xl.wb.display(window.__xl.grid.sheet, 5, 8).text)) !== '';
console.log(`Mở tệp: ${(open / 1000).toFixed(1)} s; mở lại app với sổ đó: ${(reopen / 1000).toFixed(1)} s`);
const pass = open < 15000 && reopen < 8000 && ok;
console.log(pass ? '✓ đạt (mở < 15 s, mở lại < 8 s)' : '✗ chậm hơn ngưỡng');
await b.close(); try { process.kill(-srv.pid); } catch {}
process.exit(pass ? 0 : 1);
