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

// ---------- F4: đổi kiểu tham chiếu khi sửa công thức ----------
// Tham chiếu ô/vùng (có thể kèm tên sheet) trong chuỗi công thức; bỏ qua trong chuỗi "…" và tên hàm (theo sau là "(")
const CREF = /((?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?/y;
function refsIn(f: string): { start: number; end: number; sheet: string; a: string; b?: string }[] {
  const out: { start: number; end: number; sheet: string; a: string; b?: string }[] = [];
  let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') { i = skipString(f, i); continue; }
    const prev = i > 0 ? f[i - 1] : '';
    if (!/[A-Za-z0-9_.$À-￿]/.test(prev) || ch === "'") {
      CREF.lastIndex = i;
      const m = CREF.exec(f);
      if (m && !/[A-Za-z0-9_(.!]/.test(f[CREF.lastIndex] ?? '') && colIndex(m[2].replace(/\$/g, '').replace(/\d+$/, '')) < MAX_COLS) {
        out.push({ start: i, end: CREF.lastIndex, sheet: m[1] ?? '', a: m[2], b: m[3] });
        i = CREF.lastIndex; continue;
      }
      if (ch === "'") { let j = i + 1; while (j < f.length && !(f[j] === "'" && f[j + 1] !== "'")) j += f[j] === "'" ? 2 : 1; i = j + 1; continue; }
    }
    i++;
  }
  return out;
}
/** A1 → $A$1 → A$1 → $A1 → A1 */
function nextAbs(ref: string): string {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref)!;
  const st = (m[1] ? 1 : 0) + (m[3] ? 2 : 0); // 0: A1, 3: $A$1, 2: A$1, 1: $A1
  const next: Record<number, [string, string]> = { 0: ['$', '$'], 3: ['', '$'], 2: ['$', ''], 1: ['', ''] };
  const [c, r] = next[st];
  return c + m[2] + r + m[4];
}
/**
 * F4 khi đang sửa công thức: đổi kiểu tham chiếu tại con trỏ (hoặc mọi tham chiếu trong đoạn bôi đen).
 * Trả về chuỗi mới và đoạn chọn mới (bao tham chiếu vừa đổi) để bấm F4 tiếp.
 */
export function cycleRef(f: string, selStart: number, selEnd: number): { text: string; start: number; end: number } {
  const refs = refsIn(f);
  let hit = selStart === selEnd
    ? refs.filter((x) => selStart >= x.start && selStart <= x.end).slice(-1)
    : refs.filter((x) => x.end > selStart && x.start < selEnd);
  if (!hit.length && selStart === selEnd) { const before = refs.filter((x) => x.end <= selStart); hit = before.slice(-1); }
  if (!hit.length) return { text: f, start: selStart, end: selEnd };
  let out = '', last = 0, start = -1, end = -1;
  for (const x of hit) {
    const nt = x.sheet + nextAbs(x.a) + (x.b ? ':' + nextAbs(x.b) : '');
    out += f.slice(last, x.start);
    if (start < 0) start = out.length;
    out += nt; end = out.length; last = x.end;
  }
  out += f.slice(last);
  return { text: out, start, end };
}

/**
 * Dán chuyển vị: tham chiếu tương đối đổi hướng (lệch cột ↔ lệch hàng) như Excel;
 * tham chiếu tuyệt đối giữ nguyên, tham chiếu hỗn hợp dời theo độ lệch vị trí.
 */
export function transposeFormula(f: string, sr: number, sc: number, tr: number, tc: number): string {
  if (!f.startsWith('=')) return f;
  const refs = refsIn(f);
  let out = '', last = 0;
  const one = (ref: string): string | null => {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref)!;
    const c = colIndex(m[2]), r = +m[4] - 1;
    let nr: number, nc: number;
    if (!m[1] && !m[3]) { nr = tr + (c - sc); nc = tc + (r - sr); }
    else { nr = m[3] ? r : r + (tr - sr); nc = m[1] ? c : c + (tc - sc); }
    if (nr < 0 || nc < 0 || nr >= MAX_ROWS || nc >= MAX_COLS) return null;
    return m[1] + colName(nc) + m[3] + (nr + 1);
  };
  for (const x of refs) {
    const a = one(x.a), b = x.b ? one(x.b) : '';
    out += f.slice(last, x.start) + (a === null || b === null ? '#REF!' : x.sheet + a + (x.b ? ':' + b : ''));
    last = x.end;
  }
  return out + f.slice(last);
}
