import { colIndex, colName, MAX_COLS, MAX_ROWS } from './address';

// Dời tham chiếu tương đối trong công thức (dùng khi kéo điền), giữ phần có dấu $
const REF = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d{1,7}):(\$?)(\d{1,7})/y;

export function shiftFormula(f: string, dr: number, dc: number): string {
  if (!f.startsWith('=')) return f;
  let out = '';
  let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') {
      const j = f.indexOf('"', i + 1) === -1 ? f.length : skipString(f, i);
      out += f.slice(i, j); i = j; continue;
    }
    if (ch === "'") { // tên sheet có dấu nháy
      let j = i + 1;
      while (j < f.length && !(f[j] === "'" && f[j + 1] !== "'")) j += f[j] === "'" ? 2 : 1;
      out += f.slice(i, j + 1); i = j + 1; continue;
    }
    const prev = i > 0 ? f[i - 1] : '';
    if (/[A-Za-z0-9_.$]/.test(prev) && !/[$]/.test(ch)) { out += ch; i++; continue; }
    REF.lastIndex = i;
    const m = REF.exec(f);
    if (m && f[REF.lastIndex] === '!') { out += m[0]; i = REF.lastIndex; continue; } // tên sheet giống địa chỉ ô (Q1!A1)
    if (m && !/[A-Za-z0-9_(.]/.test(f[REF.lastIndex] ?? '')) {
      out += rewrite(m, dr, dc);
      i = REF.lastIndex;
      continue;
    }
    out += ch; i++;
  }
  return out;
}

function skipString(f: string, i: number) {
  let j = i + 1;
  while (j < f.length) {
    if (f[j] === '"' && f[j + 1] === '"') { j += 2; continue; }
    if (f[j] === '"') return j + 1;
    j++;
  }
  return f.length;
}

function cell(ca: string, cn: string, ra: string, rn: string, dr: number, dc: number): string | null {
  const c = ca ? colIndex(cn) : colIndex(cn) + dc;
  const r = ra ? +rn - 1 : +rn - 1 + dr;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
  return `${ca}${colName(c)}${ra}${r + 1}`;
}

function rewrite(m: RegExpExecArray, dr: number, dc: number): string {
  if (m[2] !== undefined) {
    if (colIndex(m[2]) >= MAX_COLS) return m[0];
    const a = cell(m[1], m[2], m[3], m[4], dr, dc);
    if (m[6] === undefined) return a ?? '#REF!';
    const b = cell(m[5], m[6], m[7], m[8], dr, dc);
    return a && b ? `${a}:${b}` : '#REF!';
  }
  if (m[10] !== undefined) {
    const sh = (abs: string, n: string) => { const c = abs ? colIndex(n) : colIndex(n) + dc; return c < 0 || c >= MAX_COLS ? null : abs + colName(c); };
    const a = sh(m[9], m[10]), b = sh(m[11], m[12]);
    return a && b ? `${a}:${b}` : '#REF!';
  }
  const sh = (abs: string, n: string) => { const r = abs ? +n : +n + dr; return r < 1 || r > MAX_ROWS ? null : abs + r; };
  const a = sh(m[13], m[14]), b = sh(m[15], m[16]);
  return a && b ? `${a}:${b}` : '#REF!';
}

/** Tên sheet cần đặt trong dấu nháy khi viết công thức */
export function needsQuote(name: string): boolean {
  return !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^(TRUE|FALSE)$/i.test(name);
}
export function quoteSheet(name: string): string {
  return needsQuote(name) ? `'${name.replace(/'/g, "''")}'` : name;
}
const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** Thêm dấu nháy cho các tên sheet bị viết trần (vd =2024!A1 → ='2024'!A1), bỏ qua phần trong chuỗi "..." */
export function fixSheetQuotes(f: string, names: string[]): string {
  if (!f.startsWith('=') || !names.length) return f;
  const sorted = [...names].sort((a, b) => b.length - a.length);
  const re = new RegExp(`(^|[^'A-Za-z0-9_.\\u00C0-\\uFFFF])(${sorted.map(esc).join('|')})!`, 'g');
  return f.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(re, (_m, pre, n) => `${pre}'${n.replace(/'/g, "''")}'!`))).join('');
}
