// Định dạng có điều kiện (CF) và kiểm tra dữ liệu (DV)
import type { Range } from './address';
import { normalizeInput } from './input';

export interface CFFormat { fg?: string; bg?: string; b?: boolean; i?: boolean }
export type CFType = 'cellIs' | 'text' | 'dup' | 'unique' | 'top' | 'bottom' | 'aboveAvg' | 'belowAvg' | 'scale' | 'bar';
export interface CFRule {
  id?: string;
  range: Range;
  type: CFType;
  op?: 'gt' | 'lt' | 'ge' | 'le' | 'eq' | 'ne' | 'between' | 'notBetween';
  v1?: string; v2?: string;
  text?: string;
  rank?: number; percent?: boolean;
  colors?: string[];
  fmt?: CFFormat;
  raw?: unknown; // quy tắc .xlsx chưa hỗ trợ: giữ nguyên để ghi lại
}
export interface CFResult extends CFFormat { bar?: { frac: number; color: string } }

export type DVType = 'list' | 'whole' | 'decimal' | 'date' | 'textLength' | 'custom';
export interface DVRule {
  id?: string;
  range: Range;
  type: DVType;
  op?: CFRule['op'];
  v1?: string; v2?: string;
  list?: string[];        // danh sách cố định
  listRef?: string;       // hoặc vùng ô, vd "$A$1:$A$5" hoặc "'Danh mục'!A1:A9"
  allowBlank?: boolean;
  error?: string; errorTitle?: string; errorStyle?: 'stop' | 'warning' | 'information';
  prompt?: string; promptTitle?: string;
  raw?: unknown;
}

export const inR = (g: Range, r: number, c: number) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;

export function compare(op: CFRule['op'], x: number, a: number, b?: number): boolean {
  switch (op) {
    case 'gt': return x > a;
    case 'lt': return x < a;
    case 'ge': return x >= a;
    case 'le': return x <= a;
    case 'eq': return x === a;
    case 'ne': return x !== a;
    case 'between': return x >= Math.min(a, b ?? a) && x <= Math.max(a, b ?? a);
    case 'notBetween': return !(x >= Math.min(a, b ?? a) && x <= Math.max(a, b ?? a));
    default: return false;
  }
}

/** Đọc số từ chuỗi người dùng nhập theo kiểu Việt Nam (1.250,5 / 1,5 / 15% / ngày) */
export function numOf(s: string | undefined, parse?: (raw: string) => unknown): number | null {
  if (s === undefined || s === '') return null;
  const t = s.trim();
  if (/^-?\d+(,\d+)?$/.test(t)) return +t.replace(',', '.');
  if (/^-?\d+(,\d+)?[eE][+-]?\d+$/.test(t)) return +t.replace(',', '.');
  const n = normalizeInput(t);
  if (typeof n.content === 'number') return n.content;
  if (parse) { const v = parse(t); if (typeof v === 'number') return v; }
  return null;
}

// ---------- màu ----------
function hex(c: string) { const h = c.replace('#', ''); const f = h.length === 3 ? h.split('').map((x) => x + x).join('') : h; return [0, 2, 4].map((i) => parseInt(f.slice(i, i + 2), 16)); }
const toHex = (rgb: number[]) => '#' + rgb.map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('');
export function mix(a: string, b: string, t: number) { const x = hex(a), y = hex(b); return toHex(x.map((v, i) => v + (y[i] - v) * t)); }

/** Thống kê vùng cần cho một số loại quy tắc */
export interface CFStats { nums: number[]; min: number; max: number; avg: number; counts: Map<string, number>; cut?: number }

export function cfStats(rule: CFRule, values: unknown[]): CFStats {
  const nums = values.filter((v): v is number => typeof v === 'number').sort((a, b) => a - b);
  const counts = new Map<string, number>();
  for (const v of values) if (v !== null && v !== '' && v !== undefined) { const k = keyOf(v); counts.set(k, (counts.get(k) ?? 0) + 1); }
  const st: CFStats = { nums, min: nums[0] ?? 0, max: nums[nums.length - 1] ?? 0, avg: nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : 0, counts };
  if (rule.type === 'top' || rule.type === 'bottom') {
    const n = Math.max(1, rule.percent ? Math.floor((nums.length * (rule.rank ?? 10)) / 100) : rule.rank ?? 10);
    st.cut = rule.type === 'top' ? nums[Math.max(0, nums.length - n)] : nums[Math.min(nums.length - 1, n - 1)];
  }
  return st;
}
const keyOf = (v: unknown) => (typeof v === 'string' ? 's:' + v.toLowerCase() : typeof v + ':' + String(v));

export function cfEval(rule: CFRule, v: unknown, text: string, st: CFStats | null): CFResult | null {
  const f = rule.fmt ?? {};
  switch (rule.type) {
    case 'cellIs': {
      if (typeof v === 'number') {
        const a = numOf(rule.v1), b = numOf(rule.v2);
        return a !== null && compare(rule.op, v, a, b ?? undefined) ? f : null;
      }
      if (typeof v === 'string' && (rule.op === 'eq' || rule.op === 'ne')) {
        const same = v.toLowerCase() === String(rule.v1 ?? '').replace(/^"|"$/g, '').toLowerCase();
        return (rule.op === 'eq') === same ? f : null;
      }
      return null;
    }
    case 'text': return rule.text && text.toLowerCase().includes(rule.text.toLowerCase()) ? f : null;
    case 'dup': return v !== null && v !== '' && (st!.counts.get(keyOf(v)) ?? 0) > 1 ? f : null;
    case 'unique': return v !== null && v !== '' && st!.counts.get(keyOf(v)) === 1 ? f : null;
    case 'top': return typeof v === 'number' && st!.cut !== undefined && v >= st!.cut ? f : null;
    case 'bottom': return typeof v === 'number' && st!.cut !== undefined && v <= st!.cut ? f : null;
    case 'aboveAvg': return typeof v === 'number' && v > st!.avg ? f : null;
    case 'belowAvg': return typeof v === 'number' && v < st!.avg ? f : null;
    case 'scale': {
      if (typeof v !== 'number' || !rule.colors?.length) return null;
      const t = st!.max === st!.min ? 0.5 : (v - st!.min) / (st!.max - st!.min);
      const cs = rule.colors;
      if (cs.length === 2) return { bg: mix(cs[0], cs[1], t) };
      return { bg: t <= 0.5 ? mix(cs[0], cs[1], t * 2) : mix(cs[1], cs[2], (t - 0.5) * 2) };
    }
    case 'bar': {
      if (typeof v !== 'number') return null;
      const lo = Math.min(0, st!.min), hi = st!.max;
      return { bar: { frac: hi === lo ? 1 : (v - lo) / (hi - lo), color: rule.colors?.[0] ?? '#638ec6' } };
    }
  }
  return null;
}

export function dvDescribe(rule: DVRule): string {
  const op: Record<string, string> = { gt: 'lớn hơn', lt: 'nhỏ hơn', ge: 'từ', le: 'đến tối đa', eq: 'bằng', ne: 'khác', between: 'từ', notBetween: 'ngoài khoảng' };
  const kind = { whole: 'số nguyên', decimal: 'số', date: 'ngày', textLength: 'độ dài chữ', list: '', custom: '' }[rule.type];
  if (rule.type === 'list') return 'Giá trị phải chọn trong danh sách.';
  if (rule.op === 'between' || rule.op === 'notBetween') return `Giá trị phải là ${kind} ${op[rule.op]} ${rule.v1} ${rule.op === 'between' ? 'đến' : '–'} ${rule.v2}.`;
  return `Giá trị phải là ${kind} ${op[rule.op ?? 'eq']} ${rule.v1}.`;
}
