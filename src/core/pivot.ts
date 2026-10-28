// PivotTable: tổng hợp dữ liệu theo hàng/cột (bố cục dạng bảng như Excel)
export type Agg = 'sum' | 'count' | 'avg' | 'max' | 'min';
export interface PivotSpec {
  rows: string[];
  cols: string[];
  values: { field: string; agg: Agg }[];
  filters: { field: string; values: string[] }[];
}
export type PKind = 'h' | 'r' | 'v' | 't' | 'g' | '';
export interface PivotResult { cells: (string | number | null)[][]; kinds: PKind[][]; warnings: string[] }

export const AGG_LABEL: Record<Agg, string> = { sum: 'Tổng', count: 'Đếm', avg: 'Trung bình', max: 'Lớn nhất', min: 'Nhỏ nhất' };
export const BLANK = '(trống)';
const COLL = new Intl.Collator('vi', { sensitivity: 'accent', numeric: true });

export const keyText = (v: unknown) => (v === null || v === undefined || v === '' ? BLANK : typeof v === 'number' ? String(v).replace('.', ',') : String(v));

function cmpKey(a: { t: string; v: unknown }, b: { t: string; v: unknown }) {
  if (a.t === BLANK || b.t === BLANK) return a.t === b.t ? 0 : a.t === BLANK ? 1 : -1;
  if (typeof a.v === 'number' && typeof b.v === 'number') return a.v - b.v;
  if (typeof a.v === 'number') return -1;
  if (typeof b.v === 'number') return 1;
  return COLL.compare(a.t, b.t);
}

class Acc {
  sum = 0; n = 0; cnt = 0; max = -Infinity; min = Infinity;
  add(v: unknown) {
    if (v === null || v === undefined || v === '') return;
    this.cnt++;
    if (typeof v === 'number' && isFinite(v)) { this.sum += v; this.n++; if (v > this.max) this.max = v; if (v < this.min) this.min = v; }
  }
  get(agg: Agg): number | null {
    switch (agg) {
      case 'count': return this.cnt;
      case 'sum': return this.n ? this.sum : this.cnt ? 0 : null;
      case 'avg': return this.n ? this.sum / this.n : null;
      case 'max': return this.n ? this.max : null;
      case 'min': return this.n ? this.min : null;
    }
  }
}

/** Bảng nguồn: text[r][c] (nếu có) là chữ hiển thị dùng làm nhãn, vd. ngày đã định dạng */
export interface PivotTableData { header: string[]; rows: unknown[][]; text?: (string | undefined)[][] }
export const labelOf = (t: PivotTableData, r: number, i: number) => t.text?.[r]?.[i] ?? keyText(t.rows[r][i]);
// ký tự ngăn khóa không thể có trong dữ liệu
const SEP = '\u0002', ALL = '\u0003', OUT = '\u0004';

export function computePivot(table: PivotTableData, spec: PivotSpec): PivotResult {
  const warnings: string[] = [];
  const idx = (f: string) => table.header.findIndex((h) => String(h ?? '').trim() === f);
  const pick = (names: string[]) => names.map((f) => { const i = idx(f); if (i < 0) warnings.push(`Không còn cột “${f}” trong vùng nguồn.`); return i; }).filter((i) => i >= 0);
  const rI = pick(spec.rows), cI = pick(spec.cols);
  const vals = spec.values.filter((v) => idx(v.field) >= 0).map((v) => ({ ...v, i: idx(v.field) }));
  const flt = spec.filters.filter((f) => idx(f.field) >= 0 && f.values?.length).map((f) => ({ i: idx(f.field), set: new Set(f.values) }));
  const rowIdx = table.rows.map((_, r) => r).filter((r) => table.rows[r].some((v) => v !== null && v !== '') && flt.every((f) => f.set.has(labelOf(table, r, f.i))));
  const rows = rowIdx.map((r) => table.rows[r]);
  const lab = (j: number, i: number) => labelOf(table, rowIdx[j], i);

  // khóa hàng, khóa cột
  type K = { t: string; v: unknown };
  const tupleKey = (j: number, ids: number[]) => ids.map((i) => lab(j, i)).join('\u0001');
  const distinct = (ids: number[]) => {
    const m = new Map<string, K[]>();
    rows.forEach((row, j) => { const k = tupleKey(j, ids); if (!m.has(k)) m.set(k, ids.map((i) => ({ t: lab(j, i), v: row[i] }))); });
    return [...m.entries()].sort((a, b) => { for (let j = 0; j < ids.length; j++) { const d = cmpKey(a[1][j], b[1][j]); if (d) return d; } return 0; });
  };
  const rowKeys = distinct(rI), colKeys = cI.length ? distinct(cI) : [];
  // gom
  const cell = new Map<string, Acc[]>();
  const acc = (k: string) => { let a = cell.get(k); if (!a) { a = vals.map(() => new Acc()); cell.set(k, a); } return a; };
  rows.forEach((row, jr) => {
    const rk = tupleKey(jr, rI), ck = tupleKey(jr, cI);
    const outer = rI.length > 1 ? lab(jr, rI[0]) : null;
    const targets = [rk + SEP + ck, rk + SEP + ALL, ALL + SEP + ck, ALL + SEP + ALL];
    if (outer !== null) targets.push(OUT + outer + SEP + ck, OUT + outer + SEP + ALL);
    for (const t of targets) acc(t).forEach((a, j) => a.add(row[vals[j].i]));
  });
  const get = (rk: string, ck: string, j: number) => cell.get(rk + SEP + ck)?.[j].get(vals[j].agg) ?? null;
  const vLabel = (v: { field: string; agg: Agg }) => `${AGG_LABEL[v.agg]} ${v.field}`;

  const cells: (string | number | null)[][] = [];
  const kinds: PKind[][] = [];
  const nR = Math.max(1, rI.length);
  const push = (row: (string | number | null)[], k: PKind[]) => { cells.push(row); kinds.push(k); };
  // các cột giá trị: (khóa cột hoặc '*', chỉ số giá trị)
  const valCols: { ck: string; j: number; label: string | null }[] = [];
  if (colKeys.length) {
    for (const [ck, parts] of colKeys) vals.forEach((_, j) => valCols.push({ ck, j, label: parts.map((p) => p.t).join(' – ') }));
    vals.forEach((_, j) => valCols.push({ ck: ALL, j, label: 'Tổng cộng' }));
  } else vals.forEach((_, j) => valCols.push({ ck: ALL, j, label: null }));

  // tiêu đề
  if (colKeys.length) {
    const top: (string | number | null)[] = Array(nR + valCols.length).fill(null);
    top[0] = vals.length === 1 ? vLabel(vals[0]) : 'Giá trị';
    top[nR] = spec.cols.filter((f) => idx(f) >= 0).join(' – ');
    push(top, top.map((v) => (v === null ? '' : 'h')));
    if (vals.length > 1) {
      const mid: (string | number | null)[] = Array(nR).fill(null);
      valCols.forEach((vc, i) => mid.push(i % vals.length === 0 ? vc.label : null));
      push(mid, mid.map((v) => (v === null ? '' : 'h')));
    }
  }
  const head: (string | number | null)[] = rI.length ? spec.rows.filter((f) => idx(f) >= 0) : [vals.length === 1 && colKeys.length ? '' : ''];
  if (!rI.length) head[0] = colKeys.length ? null : 'Giá trị';
  for (const vc of valCols) head.push(colKeys.length ? (vals.length > 1 ? vLabel(vals[vc.j]) : vc.label) : vLabel(vals[vc.j]));
  push(head, head.map(() => 'h'));

  // thân
  const line = (labels: (string | null)[], rk: string, kind: PKind) => {
    const row: (string | number | null)[] = [...labels];
    const k: PKind[] = labels.map(() => (kind === 'v' ? 'r' : kind));
    for (const vc of valCols) { row.push(get(rk, vc.ck, vc.j)); k.push(kind === 'v' ? 'v' : kind); }
    push(row, k);
  };
  if (!rI.length) line([vals.length === 1 ? vLabel(vals[0]) : 'Giá trị'], '', 'v');
  else {
    let prevOuter: string | null = null;
    rowKeys.forEach(([rk, parts], i) => {
      const labels = parts.map((p, j) => (j === 0 && rI.length > 1 && p.t === prevOuter ? null : p.t));
      line(labels, rk, 'v');
      prevOuter = parts[0].t;
      const next = rowKeys[i + 1];
      if (rI.length > 1 && (!next || next[1][0].t !== parts[0].t)) {
        const sub: (string | null)[] = Array(rI.length).fill(null);
        sub[0] = `Tổng ${parts[0].t}`;
        line(sub, OUT + parts[0].t, 't');
        prevOuter = null;
      }
    });
  }
  const g: (string | null)[] = Array(nR).fill(null);
  g[0] = 'Tổng cộng';
  if (rI.length) line(g, ALL, 'g');
  return { cells, kinds, warnings };
}
