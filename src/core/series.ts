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
