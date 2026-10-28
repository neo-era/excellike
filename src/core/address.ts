export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export interface Range { r1: number; c1: number; r2: number; c2: number }
export interface Cell { r: number; c: number }

export function colName(c: number): string {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colIndex(s: string): number {
  let c = 0;
  for (const ch of s.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
}

export const toA1 = (r: number, c: number) => colName(c) + (r + 1);

export function parseA1(s: string): Cell | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(s.trim());
  if (!m) return null;
  const c = colIndex(m[1]);
  const r = +m[2] - 1;
  if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return null;
  return { r, c };
}

export function normRange(a: Cell, b: Cell = a): Range {
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}

export function parseRange(s: string): Range | null {
  const [p, q] = s.split(':');
  const a = parseA1(p);
  const b = q === undefined ? a : parseA1(q);
  return a && b ? normRange(a, b) : null;
}

export function rangeToA1(g: Range): string {
  const a = toA1(g.r1, g.c1);
  return g.r1 === g.r2 && g.c1 === g.c2 ? a : a + ':' + toA1(g.r2, g.c2);
}

export const rangeSize = (g: Range) => (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
export const inRange = (g: Range, r: number, c: number) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;

/** Như parseRange nhưng nhận thêm cột nguyên (C:C, A:D), hàng nguyên (2:5) và dấu $ */
export function parseRef(s: string): Range | null {
  const t = s.replace(/\$/g, '').trim();
  let m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t);
  if (m) { const a = colIndex(m[1]), b = colIndex(m[2]); return a < MAX_COLS && b < MAX_COLS ? { r1: 0, r2: MAX_ROWS - 1, c1: Math.min(a, b), c2: Math.max(a, b) } : null; }
  m = /^(\d{1,7}):(\d{1,7})$/.exec(t);
  if (m) { const a = +m[1] - 1, b = +m[2] - 1; return a >= 0 && b >= 0 && a < MAX_ROWS && b < MAX_ROWS ? { r1: Math.min(a, b), r2: Math.max(a, b), c1: 0, c2: MAX_COLS - 1 } : null; }
  return parseRange(t);
}
