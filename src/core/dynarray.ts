// Hàm mảng động kiểu Excel 365 mà HyperFormula chưa có hoặc làm khác Excel:
// FILTER(mảng; điều_kiện; [nếu_rỗng]) — lọc theo hàng (hoặc cột) với mảng 2 chiều,
// SORTBY, TAKE, DROP, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW.
// Kết quả "tràn" (spill) sang các ô bên cạnh; bộ tính giữ trước vùng bằng kích thước tối đa có thể (ô thừa để trống).
import { FunctionPlugin, FunctionArgumentType, CellError, ErrorType, SimpleRangeValue, ArraySize, EmptyValue } from 'hyperformula';

type Grid = unknown[][];
const R = FunctionArgumentType.RANGE, N = FunctionArgumentType.NUMBER, S = FunctionArgumentType.SCALAR;
const truthy = (v: unknown) => (typeof v === 'number' ? v !== 0 : typeof v === 'boolean' ? v : typeof v === 'string' ? /^true$/i.test(v) : false);
const isErr = (v: unknown): v is CellError => v instanceof CellError;
const out = (g: Grid) => (g.length && g[0].length ? SimpleRangeValue.onlyValues(g as any) : new CellError(ErrorType.VALUE, 'Mảng rỗng'));
const rank = (v: unknown) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === EmptyValue || v === null ? 4 : 3);
const COLL = new Intl.Collator('vi', { sensitivity: 'base' });
const cmp = (a: unknown, b: unknown) => {
  const d = rank(a) - rank(b);
  if (d) return d;
  if (typeof a === 'number') return a - (b as number);
  if (typeof a === 'string') return COLL.compare(a, b as string);
  if (typeof a === 'boolean') return Number(a) - Number(b);
  return 0;
};
const lit = (ast: any): number | undefined => (ast?.type === 'NUMBER' ? ast.value : ast?.type === 'MINUS_UNARY_OP' && ast.value?.type === 'NUMBER' ? -ast.value.value : undefined);

export class DynArrayPlugin extends FunctionPlugin {
  static implementedFunctions = {
    FILTER: { method: 'filterX', sizeOfResultArrayMethod: 'firstSize', enableArrayArithmeticForArguments: true, parameters: [{ argumentType: R }, { argumentType: R }, { argumentType: S, optionalArg: true }], vectorizationForbidden: true },
    SORTBY: { method: 'sortby', sizeOfResultArrayMethod: 'firstSize', enableArrayArithmeticForArguments: true, parameters: [{ argumentType: R }, { argumentType: R }, { argumentType: N, defaultValue: 1 }], repeatLastArgs: 2, vectorizationForbidden: true },
    TAKE: { method: 'take', sizeOfResultArrayMethod: 'takeSize', parameters: [{ argumentType: R }, { argumentType: N }, { argumentType: N, optionalArg: true }], vectorizationForbidden: true },
    DROP: { method: 'drop', sizeOfResultArrayMethod: 'dropSize', parameters: [{ argumentType: R }, { argumentType: N }, { argumentType: N, optionalArg: true }], vectorizationForbidden: true },
    CHOOSECOLS: { method: 'choosecols', sizeOfResultArrayMethod: 'chooseColsSize', parameters: [{ argumentType: R }, { argumentType: N }], repeatLastArgs: 1, vectorizationForbidden: true },
    CHOOSEROWS: { method: 'chooserows', sizeOfResultArrayMethod: 'chooseRowsSize', parameters: [{ argumentType: R }, { argumentType: N }], repeatLastArgs: 1, vectorizationForbidden: true },
    TOCOL: { method: 'tocol', sizeOfResultArrayMethod: 'toColSize', parameters: [{ argumentType: R }, { argumentType: N, defaultValue: 0 }, { argumentType: N, defaultValue: 0 }], vectorizationForbidden: true },
    TOROW: { method: 'torow', sizeOfResultArrayMethod: 'toRowSize', parameters: [{ argumentType: R }, { argumentType: N, defaultValue: 0 }, { argumentType: N, defaultValue: 0 }], vectorizationForbidden: true },
  };
  private sz(ast: any, state: any, i = 0) { return (this as any).arraySizeForAst(ast.args[i], state) as ArraySize; }
  firstSize(ast: any, state: any) { if (!ast.args.length) return ArraySize.error(); const a = this.sz(ast, state); return new ArraySize(a.width, a.height); }
  takeSize(ast: any, state: any) {
    const a = this.sz(ast, state), r = lit(ast.args[1]), c = lit(ast.args[2]);
    return new ArraySize(c === undefined ? (ast.args[2] ? a.width : a.width) : Math.min(a.width, Math.abs(c)) || a.width, r === undefined ? a.height : Math.min(a.height, Math.abs(r)) || a.height);
  }
  dropSize(ast: any, state: any) {
    const a = this.sz(ast, state), r = lit(ast.args[1]), c = lit(ast.args[2]);
    return new ArraySize(c === undefined ? a.width : Math.max(1, a.width - Math.abs(c)), r === undefined ? a.height : Math.max(1, a.height - Math.abs(r)));
  }
  chooseColsSize(ast: any, state: any) { const a = this.sz(ast, state); return new ArraySize(Math.max(1, ast.args.length - 1), a.height); }
  chooseRowsSize(ast: any, state: any) { const a = this.sz(ast, state); return new ArraySize(a.width, Math.max(1, ast.args.length - 1)); }
  toColSize(ast: any, state: any) { const a = this.sz(ast, state); return new ArraySize(1, Math.max(1, a.width * a.height)); }
  toRowSize(ast: any, state: any) { const a = this.sz(ast, state); return new ArraySize(Math.max(1, a.width * a.height), 1); }

  filterX(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('FILTER'), (arr: SimpleRangeValue, inc: SimpleRangeValue, empty: unknown) => {
      const d = arr.data as Grid, f = inc.data as Grid, h = arr.height(), w = arr.width();
      let rows: Grid;
      if (inc.width() === 1 && inc.height() === h) {
        rows = [];
        for (let i = 0; i < h; i++) { const v = f[i][0]; if (isErr(v)) return v; if (truthy(v)) rows.push(d[i]); }
      } else if (inc.height() === 1 && inc.width() === w) {
        const keep: number[] = [];
        for (let j = 0; j < w; j++) { const v = f[0][j]; if (isErr(v)) return v; if (truthy(v)) keep.push(j); }
        rows = keep.length ? d.map((row) => keep.map((j) => row[j])) : [];
      } else return new CellError(ErrorType.VALUE, 'Điều kiện lọc phải cùng số hàng (hoặc cột) với mảng');
      if (!rows.length || !rows[0].length) return empty !== undefined ? empty : new CellError(ErrorType.VALUE, 'Không có kết quả');
      return out(rows);
    });
  }
  sortby(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('SORTBY'), (arr: SimpleRangeValue, ...rest: unknown[]) => {
      const d = arr.data as Grid, h = arr.height(), w = arr.width();
      const keys: { v: unknown[]; o: number }[] = [];
      for (let k = 0; k < rest.length; k += 2) {
        const by = rest[k] as SimpleRangeValue, o = (rest[k + 1] as number) ?? 1;
        if (o !== 1 && o !== -1) return new CellError(ErrorType.VALUE, 'Thứ tự phải là 1 hoặc -1');
        if (by.width() === 1 && by.height() === h) keys.push({ v: (by.data as Grid).map((r) => r[0]), o });
        else if (by.height() === 1 && by.width() === w && h === 1) keys.push({ v: (by.data as Grid)[0], o });
        else return new CellError(ErrorType.VALUE, 'Mảng sắp xếp phải cùng số hàng với mảng');
      }
      const idx = [...Array(h).keys()].sort((a, b) => { for (const k of keys) { const c = cmp(k.v[a], k.v[b]); if (c) return c * k.o; } return a - b; });
      return out(idx.map((i) => d[i]));
    });
  }
  take(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('TAKE'), (arr: SimpleRangeValue, r: number, c?: number) => {
      let d = arr.data as Grid;
      if (r !== undefined && r !== null) { r = Math.trunc(r); if (!r) return new CellError(ErrorType.VALUE); d = r > 0 ? d.slice(0, r) : d.slice(r); }
      if (c !== undefined && c !== null) { c = Math.trunc(c); if (!c) return new CellError(ErrorType.VALUE); d = d.map((row) => (c! > 0 ? row.slice(0, c) : row.slice(c))); }
      return out(d);
    });
  }
  drop(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('DROP'), (arr: SimpleRangeValue, r: number, c?: number) => {
      let d = arr.data as Grid;
      if (r) { r = Math.trunc(r); d = r > 0 ? d.slice(r) : d.slice(0, r); }
      if (c) { c = Math.trunc(c); d = d.map((row) => (c! > 0 ? row.slice(c) : row.slice(0, c))); }
      return d.length && d[0].length ? out(d) : new CellError(ErrorType.VALUE, 'Mảng rỗng');
    });
  }
  choosecols(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('CHOOSECOLS'), (arr: SimpleRangeValue, ...ns: number[]) => {
      const w = arr.width(), d = arr.data as Grid, cols: number[] = [];
      for (const n of ns) { const k = Math.trunc(n); const j = k > 0 ? k - 1 : w + k; if (!k || j < 0 || j >= w) return new CellError(ErrorType.VALUE, 'Số cột ngoài phạm vi'); cols.push(j); }
      return out(d.map((row) => cols.map((j) => row[j])));
    });
  }
  chooserows(ast: any, state: any) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('CHOOSEROWS'), (arr: SimpleRangeValue, ...ns: number[]) => {
      const h = arr.height(), d = arr.data as Grid, rows: Grid = [];
      for (const n of ns) { const k = Math.trunc(n); const i = k > 0 ? k - 1 : h + k; if (!k || i < 0 || i >= h) return new CellError(ErrorType.VALUE, 'Số hàng ngoài phạm vi'); rows.push(d[i]); }
      return out(rows);
    });
  }
  private flat(arr: SimpleRangeValue, ignore: number, byCol: number) {
    const d = arr.data as Grid, o: unknown[] = [];
    const push = (v: unknown) => { if ((ignore === 1 || ignore === 3) && (v === EmptyValue || v === null)) return; if ((ignore === 2 || ignore === 3) && isErr(v)) return; o.push(v); };
    if (byCol) { for (let j = 0; j < arr.width(); j++) for (let i = 0; i < arr.height(); i++) push(d[i][j]); }
    else for (const row of d) for (const v of row) push(v);
    return o;
  }
  tocol(ast: any, state: any) { return (this as any).runFunction(ast.args, state, (this as any).metadata('TOCOL'), (arr: SimpleRangeValue, ig: number, bc: number) => { const f = this.flat(arr, ig, bc); return f.length ? out(f.map((v) => [v])) : new CellError(ErrorType.VALUE); }); }
  torow(ast: any, state: any) { return (this as any).runFunction(ast.args, state, (this as any).metadata('TOROW'), (arr: SimpleRangeValue, ig: number, bc: number) => { const f = this.flat(arr, ig, bc); return f.length ? out([f]) : new CellError(ErrorType.VALUE); }); }
}
export const DYN_NAMES = ['FILTER', 'SORTBY', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW'];
