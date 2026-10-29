// Tìm & Thay thế (Ctrl+F / Ctrl+H) và Đi tới Đặc biệt (F5 ▸ Đặc biệt…, Ctrl+[ ], Alt+;, Ctrl+\ …) theo Excel.
import type { Range } from './address';
import { compileRefs, type Workbook } from './workbook';

export interface FindOpts {
  what: string;
  within: 'sheet' | 'book';
  by: 'rows' | 'cols';
  lookIn: 'formulas' | 'values' | 'notes';
  matchCase: boolean;
  whole: boolean;
}
export interface Hit { s: number; r: number; c: number; text: string }

/** Mẫu Excel → RegExp: * (nhiều ký tự), ? (một ký tự), ~ (thoát) */
export function patternRe(what: string, o: { matchCase: boolean; whole: boolean }, global = false): RegExp {
  let src = '';
  for (let i = 0; i < what.length; i++) {
    const ch = what[i];
    if (ch === '~' && i + 1 < what.length && '*?~'.includes(what[i + 1])) { src += '\\' + what[++i]; continue; }
    if (ch === '*') { src += '[\\s\\S]*'; continue; }
    if (ch === '?') { src += '[\\s\\S]'; continue; }
    src += ch.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }
  if (o.whole) src = '^' + src + '$';
  return new RegExp(src, (o.matchCase ? '' : 'i') + (global ? 'g' : '') + 'u');
}
function textOf(wb: Workbook, s: number, r: number, c: number, lookIn: FindOpts['lookIn']): string {
  if (lookIn === 'values') return wb.display(s, r, c).text;
  if (lookIn === 'notes') return '';
  return wb.rawText(s, r, c);
}
function sheetsOf(wb: Workbook, s: number, o: FindOpts) {
  if (o.within === 'sheet') return [s];
  const all = wb.visibleSheetIds();
  const i = all.indexOf(s);
  return i < 0 ? all : [...all.slice(i), ...all.slice(0, i)];
}

/** Mọi ô khớp (theo thứ tự tìm của Excel). scope = vùng chọn nhiều ô (chỉ áp cho trang hiện hành) */
export function findAll(wb: Workbook, s: number, o: FindOpts, scope?: Range[]): Hit[] {
  if (!o.what) return [];
  const re = patternRe(o.what, o);
  const out: Hit[] = [];
  for (const sid of sheetsOf(wb, s, o)) {
    const u = wb.used(sid);
    const areas = sid === s && scope?.length ? scope : [{ r1: 0, c1: 0, r2: u.rows - 1, c2: u.cols - 1 }];
    const seen = new Set<number>();
    const hits: Hit[] = [];
    for (const a of areas) {
      const r2 = Math.min(a.r2, u.rows - 1), c2 = Math.min(a.c2, u.cols - 1);
      for (let r = a.r1; r <= r2; r++) for (let c = a.c1; c <= c2; c++) {
        const k = r * 20000 + c;
        if (seen.has(k)) continue;
        seen.add(k);
        const t = textOf(wb, sid, r, c, o.lookIn);
        if (t && re.test(t)) hits.push({ s: sid, r, c, text: t });
      }
    }
    hits.sort((x, y) => (o.by === 'rows' ? x.r - y.r || x.c - y.c : x.c - y.c || x.r - y.r));
    out.push(...hits);
  }
  return out;
}

/** Tìm tiếp (dir 1) / ngược lại (dir −1) từ ô hiện hành, quay vòng */
export function findNext(wb: Workbook, s: number, from: { s: number; r: number; c: number }, o: FindOpts, dir: 1 | -1, scope?: Range[]): Hit | null {
  const all = findAll(wb, s, o, scope);
  if (!all.length) return null;
  const order = (h: { s: number; r: number; c: number }) => {
    const si = sheetsOf(wb, s, o).indexOf(h.s);
    return [si, o.by === 'rows' ? h.r : h.c, o.by === 'rows' ? h.c : h.r];
  };
  const cmp = (a: number[], b: number[]) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
  const cur = order(from);
  if (dir > 0) return all.find((h) => cmp(order(h), cur) > 0) ?? all[0];
  for (let i = all.length - 1; i >= 0; i--) if (cmp(order(all[i]), cur) < 0) return all[i];
  return all[all.length - 1];
}

/** Văn bản mới khi thay; null nếu không đổi hoặc tạo công thức lỗi */
export function replaceText(wb: Workbook, s: number, raw: string, o: FindOpts, repl: string): string | null {
  const re = patternRe(o.what, o, true);
  const next = o.whole ? (re.test(raw) ? repl : raw) : raw.replace(re, () => repl);
  if (next === raw) return null;
  if (next.startsWith('=') && !wb.isValidFormula(next, s)) return null;
  return next;
}
/** Thay thế tất cả (luôn tìm trong công thức như Excel); một bước hoàn tác. Trả về số ô đã thay. */
export function replaceAll(wb: Workbook, s: number, o: FindOpts, repl: string, scope?: Range[]): number {
  const hits = findAll(wb, s, { ...o, lookIn: 'formulas' }, scope);
  const bySheet = new Map<number, [number, number, string][]>();
  for (const h of hits) {
    const t = replaceText(wb, h.s, h.text, o, repl);
    if (t === null) continue;
    let l = bySheet.get(h.s);
    if (!l) bySheet.set(h.s, (l = []));
    l.push([h.r, h.c, t]);
  }
  let n = 0;
  wb.group(() => { for (const [sid, l] of bySheet) { wb.writeCells(sid, l); n += l.length; } });
  return n;
}

// ---------------- Đi tới Đặc biệt ----------------
export type SpecialKind = 'blanks' | 'constants' | 'formulas' | 'region' | 'lastcell' | 'visible' | 'rowDiff' | 'colDiff'
  | 'precedents' | 'dependents' | 'cf' | 'dv' | 'notes' | 'objects';
export interface SpecialOpts {
  kind: SpecialKind;
  types?: { num?: boolean; text?: boolean; bool?: boolean; err?: boolean };
  all?: boolean;                 // tiền lệ/phụ thuộc: mọi cấp
  active?: { r: number; c: number };
}

/** Danh sách ô [r, c] → các khối chữ nhật (gộp theo hàng rồi gộp dọc các khối cùng cột) */
export function toRuns(cells: [number, number][]): Range[] {
  const rows = new Map<number, number[]>();
  for (const [r, c] of cells) { let l = rows.get(r); if (!l) rows.set(r, (l = [])); l.push(c); }
  const open = new Map<string, Range>();
  const out: Range[] = [];
  const rs = [...rows.keys()].sort((a, b) => a - b);
  let prevR = -2;
  for (const r of rs) {
    const cs = [...new Set(rows.get(r)!)].sort((a, b) => a - b);
    const runs: [number, number][] = [];
    for (const c of cs) { const l = runs[runs.length - 1]; if (l && c === l[1] + 1) l[1] = c; else runs.push([c, c]); }
    const next = new Map<string, Range>();
    for (const [c1, c2] of runs) {
      const key = c1 + ':' + c2;
      const o = r === prevR + 1 ? open.get(key) : undefined;
      if (o) { o.r2 = r; next.set(key, o); }
      else { const g = { r1: r, c1, r2: r, c2 }; out.push(g); next.set(key, g); }
    }
    open.clear(); for (const [k, v] of next) open.set(k, v);
    prevR = r;
  }
  return out.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
}

const isBlank = (v: unknown) => v === null || v === '' || v === undefined;
function typeOf(v: unknown): 'num' | 'text' | 'bool' | 'err' {
  if (typeof v === 'number') return 'num';
  if (typeof v === 'boolean') return 'bool';
  if (v && typeof v === 'object') return 'err';
  return 'text';
}

/** Đi tới Đặc biệt. sel = các vùng đang chọn; một ô → xét cả vùng đã dùng (như Excel). Trả về [] nếu không có ô nào. */
export function goSpecial(wb: Workbook, s: number, sel: Range[], o: SpecialOpts): Range[] {
  const u = wb.used(s);
  const last = wb.lastCell(s);
  const single = sel.length === 1 && sel[0].r1 === sel[0].r2 && sel[0].c1 === sel[0].c2;
  const usedR = { r1: 0, c1: 0, r2: Math.max(last.r, u.rows - 1), c2: Math.max(last.c, u.cols - 1) };
  const areas = (single ? [usedR] : sel).map((g) => ({ r1: g.r1, c1: g.c1, r2: Math.min(g.r2, usedR.r2), c2: Math.min(g.c2, usedR.c2) }));
  const each = (f: (r: number, c: number) => boolean) => {
    const out: [number, number][] = [];
    for (const a of areas) for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) if (f(r, c)) out.push([r, c]);
    return toRuns(out);
  };
  const want = (t: ReturnType<typeof typeOf>) => !o.types || !Object.values(o.types).some(Boolean) || !!o.types[t];
  const act = o.active ?? { r: sel[sel.length - 1]?.r1 ?? 0, c: sel[sel.length - 1]?.c1 ?? 0 };
  switch (o.kind) {
    case 'blanks': return each((r, c) => !wb.isFormula(s, r, c) && isBlank(wb.value(s, r, c)) && !wb.mergeCovered(s, r, c));
    case 'constants': return each((r, c) => !wb.isFormula(s, r, c) && !isBlank(wb.value(s, r, c)) && want(typeOf(wb.value(s, r, c))));
    case 'formulas': return each((r, c) => wb.isFormula(s, r, c) && want(typeOf(wb.value(s, r, c))));
    case 'region': return [wb.currentRegion(s, act.r, act.c)];
    case 'lastcell': return [{ r1: last.r, c1: last.c, r2: last.r, c2: last.c }];
    case 'visible': {
      const out: [number, number][] = [];
      for (const g of sel) {
        const r2 = g.r2 >= 1048575 ? Math.max(g.r1, usedR.r2) : g.r2, c2 = g.c2 >= 16383 ? Math.max(g.c1, usedR.c2) : g.c2;
        for (let r = g.r1; r <= r2; r++) {
          if (wb.isRowHidden(s, r)) continue;
          for (let c = g.c1; c <= c2; c++) if (!wb.isColHidden(s, c)) out.push([r, c]);
        }
      }
      return toRuns(out);
    }
    case 'rowDiff': case 'colDiff': {
      const same = (a: unknown, b: unknown) => (a && typeof a === 'object' && b && typeof b === 'object' ? (a as any).value === (b as any).value : a === b || (isBlank(a) && isBlank(b)));
      return each((r, c) => {
        const ref = o.kind === 'rowDiff' ? wb.value(s, r, act.c) : wb.value(s, act.r, c);
        if (o.kind === 'rowDiff' ? c === act.c : r === act.r) return false;
        return !same(wb.value(s, r, c), ref);
      });
    }
    case 'precedents': {
      const found = new Set<string>(), out: [number, number][] = [];
      let frontier: [number, number][] = [];
      for (const a of sel) for (let r = a.r1; r <= Math.min(a.r2, usedR.r2); r++) for (let c = a.c1; c <= Math.min(a.c2, usedR.c2); c++) frontier.push([r, c]);
      while (frontier.length) {
        const next: [number, number][] = [];
        for (const [r, c] of frontier) {
          if (!wb.isFormula(s, r, c)) continue;
          for (const p of compileRefs(wb.rawText(s, r, c), (n) => (n === undefined ? s : wb.sheetIdByName(n)))) {
            if (typeof p === 'string' || p.s !== s) continue;
            for (let rr = p.r1; rr <= Math.min(p.r2, usedR.r2); rr++) for (let cc = p.c1; cc <= Math.min(p.c2, usedR.c2); cc++) {
              const k = rr + ':' + cc;
              if (found.has(k)) continue;
              found.add(k); out.push([rr, cc]); next.push([rr, cc]);
            }
          }
        }
        frontier = o.all ? next : [];
      }
      return toRuns(out);
    }
    case 'dependents': {
      const found = new Set<string>(), out: [number, number][] = [];
      let frontier: [number, number][] = [];
      for (const a of sel) for (let r = a.r1; r <= Math.min(a.r2, usedR.r2); r++) for (let c = a.c1; c <= Math.min(a.c2, usedR.c2); c++) frontier.push([r, c]);
      while (frontier.length) {
        const next: [number, number][] = [];
        for (const [r, c] of frontier) for (const d of wb.dependentsOf(s, r, c)) {
          if (d.s !== s) continue;
          const k = d.r + ':' + d.c;
          if (found.has(k)) continue;
          found.add(k); out.push([d.r, d.c]); next.push([d.r, d.c]);
        }
        frontier = o.all ? next : [];
      }
      return toRuns(out);
    }
    case 'cf': return wb.meta(s).cf.map((x) => ({ ...x.range }));
    case 'dv': return wb.meta(s).dv.map((x) => ({ ...x.range }));
    case 'notes': case 'objects': return [];
  }
  return [];
}
