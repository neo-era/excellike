// Kế hoạch kéo điền theo quy tắc Excel
export interface SrcCell { raw: unknown; num?: number; isDate?: boolean; formula?: boolean }
export type FillItem = { value: unknown } | { copy: number };

export function fillPlan(src: SrcCell[], count: number, forward: boolean): FillItem[] {
  const n = src.length;
  const out: FillItem[] = [];
  const allNum = n >= 2 && src.every((s) => !s.formula && typeof s.num === 'number');
  if (allNum) {
    const first = src[0].num!, step = (src[n - 1].num! - first) / (n - 1);
    for (let k = 1; k <= count; k++) out.push({ value: +(forward ? first + step * (n - 1 + k) : first - step * k).toPrecision(15) });
    return out;
  }
  if (n === 1 && !src[0].formula) {
    const s = src[0];
    if (s.isDate && typeof s.num === 'number') {
      for (let k = 1; k <= count; k++) out.push({ value: s.num + (forward ? k : -k) });
      return out;
    }
    const m = typeof s.raw === 'string' ? /^(.*?)(\d+)$/.exec(s.raw) : null;
    if (m && m[1] !== '') {
      const base = +m[2];
      const width = m[2].length;
      for (let k = 1; k <= count; k++) out.push({ value: m[1] + String(Math.max(0, base + (forward ? k : -k))).padStart(width, '0') });
      return out;
    }
  }
  for (let k = 1; k <= count; k++) out.push({ copy: forward ? (k - 1) % n : (((n - k) % n) + n) % n });
  return out;
}

// ---------- Trang đầu ▸ Điền ▸ Chuỗi… ----------
export interface SeriesOpts { type: 'linear' | 'growth' | 'date' | 'auto'; step?: number; stop?: number; unit?: 'day' | 'weekday' | 'month' | 'year' }
const DAY = 86400000, EPOCH0 = Date.UTC(1899, 11, 30);
const toDate = (n: number) => new Date(EPOCH0 + Math.round(n * DAY));
const toSerial = (d: Date) => Math.round((d.getTime() - EPOCH0) / DAY);
/** Cộng tháng như EDATE: ngày 31 + 1 tháng → ngày cuối tháng sau */
function addMonths(n: number, k: number) {
  const d = toDate(n), y = d.getUTCFullYear(), m = d.getUTCMonth() + k, day = d.getUTCDate();
  const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return toSerial(new Date(Date.UTC(y, m, Math.min(day, last)))) + (n - Math.floor(n));
}
/** Dãy `count` giá trị bắt đầu từ start (dừng trước khi vượt `stop`) */
export function seriesValues(start: number, count: number, o: SeriesOpts): number[] {
  const step = o.step ?? 1, out: number[] = [];
  const over = (v: number) => o.stop !== undefined && (step >= 0 ? v > o.stop + 1e-9 : v < o.stop - 1e-9);
  let wd = start;
  for (let i = 0; i < count; i++) {
    let v: number;
    if (o.type === 'growth') v = start * Math.pow(step, i);
    else if (o.type === 'date') {
      const u = o.unit ?? 'day';
      if (u === 'month') v = addMonths(start, i * step);
      else if (u === 'year') v = addMonths(start, i * step * 12);
      else if (u === 'weekday') {
        if (i > 0) { let k = Math.abs(step); const dir = Math.sign(step) || 1; while (k > 0) { wd += dir; const dow = toDate(wd).getUTCDay(); if (dow !== 0 && dow !== 6) k--; } }
        v = wd;
      } else v = start + i * step;
    } else v = start + i * step;
    v = Math.round(v * 1e10) / 1e10;
    if (over(v)) break;
    out.push(v);
  }
  return out;
}
