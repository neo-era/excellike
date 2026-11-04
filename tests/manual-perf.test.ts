// B1: đo thời gian ghi một ô ở chế độ Thủ công trên tệp đề xuất đèn LED thật (bỏ qua nếu không có tệp).
import { describe, it, expect } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { hostname, cpus } from 'node:os';
import { readXlsx } from '../src/core/xlsx';

const REAL = '/mnt/user-data/attachments/ccaa393e-e2fa-51aa-a3c2-ab7a8c80a09f/c0114642-TONG_HOP_DE_XUAT_DEN_LED_PA3_V2A_them4cot_DADATHANG_2139_1.xlsm';
const q = (a: number[], p: number) => { const s = [...a].sort((x, y) => x - y); return s[Math.min(s.length - 1, Math.floor(p * s.length))]; };

describe('B1 hiệu năng Thủ công', () => {
  it.skipIf(!existsSync(REAL))('100 lần sửa ô trên tệp đèn LED: trung vị và p95 ≤ 50 ms', async () => {
    const { wb } = await readXlsx(new Uint8Array(readFileSync(REAL)));
    // các ô đầu vào nặng nhất (đo trước: sửa ở chế độ Tự động mất 0,7–1,5 s): bảng QCVN, QUATEST, cột hệ số
    const at = (name: string, r: number, c: number): [number, number, number] => [wb.hf.getSheetId(name)!, r, c];
    const cand = [at('QCVN 07-7-2023', 6, 3), at('QCVN 07-7-2023', 3, 3), at('QCVN 07-7-2023', 5, 3), at('QCVN 07-7-2023', 4, 3), at('QUATEST 2026', 7, 3), at('QUATEST 2026', 5, 3), at('TONG HOP DE XUAT', 69, 39), at('TONG HOP DE XUAT', 137, 39)];
    const orig = cand.map(([s, r, c]) => wb.rawText(s, r, c));
    const run = (n: number) => { const t: number[] = []; for (let i = 0; i < n; i++) { const k = i % cand.length, [s, r, c] = cand[k]; const v = Math.floor(i / cand.length) % 2 ? orig[k] : orig[k] + '0'; const t0 = performance.now(); wb.lazy(() => wb.setInput(s, r, c, v)); t.push(performance.now() - t0); } return t; };
    const auto = run(8);
    wb.setCalcMode('manual');
    const man = run(100);
    const t0 = performance.now(); wb.calcNow(); const f9 = performance.now() - t0;
    console.log(`[B1] máy ${hostname()} (${cpus()[0]?.model}, ${cpus().length} nhân): Tự động trung vị ${q(auto, 0.5).toFixed(1)} ms; Thủ công trung vị ${q(man, 0.5).toFixed(1)} ms, p95 ${q(man, 0.95).toFixed(1)} ms; F9 ${f9.toFixed(0)} ms`);
    // B3: sau F9 giá trị trùng chế độ Tự động (cùng các lần sửa cuối)
    const { wb: ref } = await readXlsx(new Uint8Array(readFileSync(REAL)));
    for (const [s, r, c] of cand) ref.setInput(s, r, c, wb.rawText(s, r, c));
    let diff = 0, n = 0;
    for (const s of wb.sheetIds()) { const d = wb.used(s); for (let r = 0; r < d.rows; r++) for (let c = 0; c < d.cols; c++) {
      const a = wb.value(s, r, c), b = ref.value(s, r, c); n++;
      const same = typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)) : JSON.stringify(a) === JSON.stringify(b);
      if (!same) diff++;
    } }
    console.log(`[B3] so ${n} ô sau F9 với Tự động: ${diff} ô khác`);
    expect(diff).toBe(0);
    expect(q(man, 0.5)).toBeLessThanOrEqual(50);
    expect(q(man, 0.95)).toBeLessThanOrEqual(50);
  }, 300_000);
});
