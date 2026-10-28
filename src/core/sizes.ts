// Kích thước hàng/cột: mặc định + danh sách tùy chỉnh (sắp xếp) để tính offset nhanh
export class AxisSizes {
  private idx: number[] = [];
  private size = new Map<number, number>();
  /** Hàng/cột đang bị ẩn (do lọc) — giữ riêng để không mất kích thước gốc */
  private hid = new Set<number>();
  private pd: number[] = [0];

  constructor(public def: number, public count: number) {}

  get(i: number) { return this.hid.has(i) ? 0 : this.size.get(i) ?? this.def; }
  /** Kích thước gốc (bỏ qua trạng thái ẩn) */
  base(i: number) { return this.size.get(i) ?? this.def; }
  isHidden(i: number) { return this.hid.has(i); }
  setHidden(ids: Iterable<number>) { this.hid = new Set(ids); this.rebuild(); }
  hidden(): number[] { return [...this.hid]; }

  set(i: number, s: number) {
    s = Math.max(0, Math.round(s));
    if (s === this.def) this.size.delete(i);
    else this.size.set(i, s);
    this.rebuild();
  }

  private rebuild() {
    this.idx = [...new Set([...this.size.keys(), ...this.hid])].sort((a, b) => a - b);
    this.pd = [0];
    for (const i of this.idx) this.pd.push(this.pd[this.pd.length - 1] + this.get(i) - this.def);
  }

  private lowerBound(i: number) {
    let lo = 0, hi = this.idx.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (this.idx[m] < i) lo = m + 1; else hi = m; }
    return lo;
  }

  /** Tổng kích thước các phần tử 0..i-1 */
  offset(i: number) { return i * this.def + this.pd[this.lowerBound(i)]; }

  /** Chỉ số lớn nhất có offset ≤ px (bỏ qua phần tử ẩn cỡ 0) */
  indexAt(px: number) {
    if (px <= 0) return this.firstVisible(0);
    let lo = 0, hi = this.count - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (this.offset(m) <= px) lo = m; else hi = m - 1;
    }
    return lo;
  }

  private firstVisible(i: number) { while (i < this.count - 1 && this.get(i) === 0) i++; return i; }

  total() { return this.offset(this.count); }

  shiftInsert(at: number, n: number) { this.remap((i) => (i >= at ? i + n : i)); }

  shiftDelete(at: number, n: number) { this.remap((i) => (i < at ? i : i >= at + n ? i - n : -1)); }

  private remap(f: (i: number) => number) {
    const m = new Map<number, number>();
    for (const [i, s] of this.size) { const j = f(i); if (j >= 0 && j < this.count) m.set(j, s); }
    this.size = m;
    const h = new Set<number>();
    for (const i of this.hid) { const j = f(i); if (j >= 0 && j < this.count) h.add(j); }
    this.hid = h;
    this.rebuild();
  }

  entries(): [number, number][] { return [...this.size.entries()]; }
  load(e: [number, number][], hidden?: number[]) { this.size = new Map(e); if (hidden) this.hid = new Set(hidden); this.rebuild(); }
  clone() { const a = new AxisSizes(this.def, this.count); a.load(this.entries(), this.hidden()); return a; }
}
