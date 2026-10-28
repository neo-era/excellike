// Hiển thị giá trị theo quy ước Việt Nam: nhóm nghìn ".", thập phân ","
export type Align = 'l' | 'c' | 'r';

const viDec = (s: string) => s.replace('.', ',');

export function formatGeneral(n: number): string {
  if (!isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const a = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (Number.isInteger(n) && a < 1e11) return String(n);
  if (a >= 1e11 || a < 1e-9) {
    const [m, e] = a.toExponential(5).split('e');
    const mant = m.replace(/\.?0+$/, '');
    const exp = +e;
    return sign + viDec(mant) + 'E' + (exp < 0 ? '-' : '+') + String(Math.abs(exp)).padStart(2, '0');
  }
  for (let p = 15; p >= 1; p--) {
    let s = a.toPrecision(p);
    if (s.includes('e')) continue;
    if (s.includes('.')) s = s.replace(/\.?0+$/, '');
    if (s.length <= 11) return sign + viDec(s);
  }
  return sign + viDec(a.toPrecision(6));
}

const EPOCH = Date.UTC(1899, 11, 30);
export function serialToDate(n: number) {
  const d = new Date(EPOCH + Math.round(n * 86400000));
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), h: d.getUTCHours(), mi: d.getUTCMinutes(), s: d.getUTCSeconds() };
}

const p2 = (n: number) => String(n).padStart(2, '0');
const DOW = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];
const DOW_LONG = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];

/** Tách các phần của mã định dạng theo dấu ; (bỏ qua trong "..." và sau \) */
function sections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '', q = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"') q = !q;
    if (!q && ch === '\\') { cur += ch + (fmt[i + 1] ?? ''); i++; continue; }
    if (!q && ch === ';') { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}
type Tok = { lit: string } | { code: string };
/** Tách mã thành phần chữ cố định và phần mã; bỏ [Red], [$-409], _x, *x */
function tokens(sec: string): Tok[] {
  const out: Tok[] = [];
  const add = (t: Tok) => {
    const last = out[out.length - 1];
    if (last && 'lit' in t && 'lit' in last) last.lit += t.lit;
    else if (last && 'code' in t && 'code' in last) last.code += t.code;
    else out.push(t);
  };
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); add({ lit: sec.slice(i + 1, j < 0 ? sec.length : j) }); i = j < 0 ? sec.length : j; continue; }
    if (ch === '\\') { add({ lit: sec[i + 1] ?? '' }); i++; continue; }
    if (ch === '_') { add({ lit: ' ' }); i++; continue; }
    if (ch === '*') { i++; continue; }
    if (ch === '[') {
      const j = sec.indexOf(']', i);
      const inner = sec.slice(i + 1, j < 0 ? sec.length : j);
      if (/^(h+|m+|s+)$/i.test(inner)) add({ code: '[' + inner + ']' });
      else if (/^\$[^-]/.test(inner)) add({ lit: inner.slice(1).split('-')[0] });
      i = j < 0 ? sec.length : j;
      continue;
    }
    if (/[0#?.,%Ee+\-yYmMdDhHsSaApP/:@]/.test(ch) && !(/[+\-]/.test(ch) && !/[Ee]/.test(sec[i - 1] ?? ''))) {
      if (ch === '/' || ch === ':') { add({ code: ch }); continue; }
      add({ code: ch });
      continue;
    }
    add({ lit: ch });
  }
  return out;
}
const codeOf = (t: Tok[]) => t.map((x) => ('code' in x ? x.code : '')).join('');

export const isDateFmt = (fmt?: string) => {
  if (!fmt) return false;
  const c = codeOf(tokens(sections(fmt)[0])).replace(/E[+-]/gi, '');
  return /[yYdDhHsS]|\[[hms]+\]/.test(c) || (/m/i.test(c) && !/[0#?]/.test(c));
};

function formatDate(n: number, sec: string): string {
  const ms = Math.round(n * 86400000);
  const d = new Date(EPOCH + ms);
  const t = { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), h: d.getUTCHours(), mi: d.getUTCMinutes(), s: d.getUTCSeconds(), dow: d.getUTCDay() };
  const toks = tokens(sec);
  const ampm = /AM\/PM|A\/P/i.test(codeOf(toks));
  // danh sách mã ngày giờ để phân biệt m (tháng) với m (phút)
  type Part = { k: string; lit?: string };
  const parts: Part[] = [];
  for (const tk of toks) {
    if ('lit' in tk) { parts.push({ k: '', lit: tk.lit }); continue; }
    const re = /\[h+\]|\[m+\]|\[s+\]|yyyy|yy|y+|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P|\.0+|./gi;
    let mm: RegExpExecArray | null;
    while ((mm = re.exec(tk.code))) parts.push({ k: mm[0] });
  }
  const kinds = parts.map((p) => p.k.toLowerCase());
  const out: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i], k = kinds[i];
    if (p.lit !== undefined) { out.push(p.lit); continue; }
    const prev = kinds.slice(0, i).reverse().find((x) => /^[hs]|^\[|^[dmy]/.test(x));
    const next = kinds.slice(i + 1).find((x) => /^[hs]|^\[|^[dmy]/.test(x));
    const minute = (k === 'm' || k === 'mm') && ((prev && /^h|^\[h/.test(prev)) || (next && /^s|^\[s/.test(next)));
    const hh = ampm ? ((t.h + 11) % 12) + 1 : t.h;
    switch (k) {
      case 'yyyy': case 'yyy': out.push(String(t.y)); break;
      case 'yy': case 'y': out.push(p2(t.y % 100)); break;
      case 'mmmmm': out.push(String(t.m)); break;
      case 'mmmm': out.push('Tháng ' + t.m); break;
      case 'mmm': out.push('thg ' + t.m); break;
      case 'mm': out.push(minute ? p2(t.mi) : p2(t.m)); break;
      case 'm': out.push(minute ? String(t.mi) : String(t.m)); break;
      case 'dddd': out.push(DOW_LONG[t.dow]); break;
      case 'ddd': out.push(DOW[t.dow]); break;
      case 'dd': out.push(p2(t.d)); break;
      case 'd': out.push(String(t.d)); break;
      case 'hh': out.push(p2(hh)); break;
      case 'h': out.push(String(hh)); break;
      case 'ss': out.push(p2(t.s)); break;
      case 's': out.push(String(t.s)); break;
      case 'am/pm': out.push(t.h < 12 ? 'SA' : 'CH'); break;
      case 'a/p': out.push(t.h < 12 ? 'S' : 'C'); break;
      default:
        if (/^\[h+\]$/.test(k)) out.push(String(Math.floor(ms / 3600000)));
        else if (/^\[m+\]$/.test(k)) out.push(String(Math.floor(ms / 60000)));
        else if (/^\[s+\]$/.test(k)) out.push(String(Math.floor(ms / 1000)));
        else if (/^\.0+$/.test(k)) out.push(',' + String(Math.floor((ms % 1000) / Math.pow(10, 3 - (k.length - 1)))).padStart(k.length - 1, '0'));
        else out.push(p.k);
    }
  }
  return out.join('');
}

function formatNum(n: number, sec: string): string {
  const toks = tokens(sec);
  // chữ trước / sau phần số
  let firstCode = toks.findIndex((t) => 'code' in t && /[0#?]/.test(t.code));
  if (firstCode < 0) firstCode = toks.findIndex((t) => 'code' in t);
  let lastCode = -1;
  toks.forEach((t, i) => { if ('code' in t && /[0#?]/.test(t.code)) lastCode = i; });
  const pre = toks.slice(0, Math.max(0, firstCode)).map((t) => ('lit' in t ? t.lit : t.code.replace(/[^%]/g, ''))).join('');
  const post = lastCode < 0 ? '' : toks.slice(lastCode + 1).map((t) => ('lit' in t ? t.lit : t.code.replace(/[^%]/g, ''))).join('');
  const code = toks.slice(Math.max(0, firstCode), lastCode + 1).map((t) => ('code' in t ? t.code : '')).join('');
  if (!/[0#?]/.test(code)) return toks.map((t) => ('lit' in t ? t.lit : t.code.replace(/[#0?,.]/g, ''))).join('');
  const pct = (sec.replace(/"[^"]*"/g, '').match(/%/g) ?? []).length;
  let v = n * Math.pow(100, pct);
  const exp = /E[+-]/i.exec(code);
  let mantCode = exp ? code.slice(0, exp.index) : code;
  // dấu phẩy ngay sau chữ số cuối = chia 1000 mỗi dấu
  const sc = /([0#?])(,+)([^0#?]*)$/.exec(mantCode);
  if (sc) { v /= Math.pow(1000, sc[2].length); mantCode = mantCode.slice(0, sc.index + 1) + sc[3]; }
  const pctIn = (mantCode.match(/%/g) ?? []).length;
  const [ip, fp = ''] = mantCode.replace(/%/g, '').split('.');
  const decimals = fp.replace(/[^0#?]/g, '').length;
  const minDec = fp.replace(/[^0]/g, '').length;
  const group = /[0#?],[0#?]/.test(ip);
  const minInt = ip.replace(/[^0]/g, '').length;
  let body: string;
  if (exp) {
    const e = v === 0 ? 0 : Math.floor(Math.log10(Math.abs(v)));
    const mant = Math.abs(v) / Math.pow(10, e);
    let m = mant.toFixed(decimals);
    if (decimals > minDec) m = m.replace(/0+$/, '').replace(/\.$/, '');
    body = m.replace('.', ',') + 'E' + (e < 0 ? '-' : exp[0][1] === '+' ? '+' : '') + String(Math.abs(e)).padStart((code.slice(exp.index + 2).match(/0/g) ?? []).length || 1, '0');
  } else {
    let fixed = Math.abs(v).toFixed(decimals);
    if (decimals > minDec) { fixed = fixed.replace(new RegExp(`0{1,${decimals - minDec}}$`), ''); if (fixed.endsWith('.')) fixed = fixed.slice(0, -1); }
    let [int, frac] = fixed.split('.');
    if (int === '0' && minInt === 0) int = '';
    if (int.length < minInt) int = int.padStart(minInt, '0');
    if (group) int = int.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    body = int + (frac ? ',' + frac : decimals && minDec ? '' : '');
  }
  return pre + body + '%'.repeat(pctIn) + post;
}

export function formatNumber(n: number, fmt: string): string {
  if (!fmt || fmt === 'General' || fmt === '@') return formatGeneral(n);
  const secs = sections(fmt);
  let sec = secs[0], neg = n < 0;
  if (n < 0 && secs.length > 1 && secs[1] !== '') { sec = secs[1]; neg = false; }
  else if (n === 0 && secs.length > 2 && secs[2] !== '') { sec = secs[2]; }
  if (/^\s*general\s*$/i.test(codeOf(tokens(sec)))) return (neg ? '-' : '') + formatGeneral(Math.abs(n));
  const a = secs.length > 1 && n < 0 ? Math.abs(n) : n;
  if (isDateFmt(sec)) return formatDate(a, sec);
  const t = formatNum(Math.abs(a), sec);
  const isZero = !/[1-9]/.test(t);
  return (neg && !isZero ? '-' : '') + t;
}

export function formatValue(v: unknown, fmt?: string): { text: string; align: Align } {
  if (v === null || v === undefined || v === '') return { text: '', align: 'l' };
  if (typeof v === 'number') return { text: fmt ? formatNumber(v, fmt) : formatGeneral(v), align: 'r' };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', align: 'c' };
  if (typeof v === 'object' && v && 'value' in v) return { text: String((v as { value: unknown }).value), align: 'c' };
  return { text: String(v), align: 'l' };
}

// Chuỗi để sửa trong ô (giống Excel: số thập phân dùng dấu phẩy, ngày dd/mm/yyyy, % giữ %)
export function editString(v: unknown, fmt?: string): string {
  if (typeof v === 'number') {
    if (isDateFmt(fmt)) {
      const c = codeOf(tokens(sections(fmt!)[0])).toLowerCase();
      const hasD = /[dy]/.test(c) || (/m/.test(c) && !/h|s/.test(c)), hasT = /h|s/.test(c);
      return formatDate(v, hasD && hasT ? 'dd/mm/yyyy hh:mm:ss' : hasT ? 'hh:mm:ss' : 'dd/mm/yyyy');
    }
    if (fmt && fmt.endsWith('%')) return viDec(String(+(v * 100).toPrecision(15))) + '%';
    return viDec(String(+v.toPrecision(15)));
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v == null ? '' : String(v);
}
