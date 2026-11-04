// LET, LAMBDA và các hàm nhận LAMBDA (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY) kiểu Excel 365.
// HyperFormula chưa có các hàm này. Cách làm:
//  - Tên trong LET và tham số LAMBDA được bộ phân tích hiểu là "tên vùng" (NAMED_EXPRESSION); khi tính,
//    bộ thông dịch được vá để tra tên trong môi trường cục bộ trước (không tạo tên toàn cục).
//  - Gọi hàm qua tên (f(3), NhanDoi(4)) là lời gọi hàm lạ (FUNCTION_CALL) → tra trong môi trường / tên vùng LAMBDA.
//  - Gọi ngay LAMBDA(x;x+1)(5) không phân tích được → lớp viết lại đổi thành XLAMCALL(LAMBDA(x;x+1);5) và đổi ngược khi đọc ra.
import { FunctionPlugin, FunctionArgumentType, CellError, ErrorType, SimpleRangeValue, EmptyValue } from 'hyperformula';

/** #CALC! (HyperFormula không có loại lỗi này): lỗi ERROR mang thông điệp riêng, sổ tính hiện thành #CALC! */
export const CALC_MSG = '#CALC!';
const calcErr = () => new CellError(ErrorType.ERROR, CALC_MSG);
const valErr = (m?: string) => new CellError(ErrorType.VALUE, m);

interface Env { vars: Map<string, unknown>; parent: Env | null }
class Lam { constructor(public params: string[], public body: any, public env: Env | null) {} }
/** tên LET / tham số LAMBDA gắn với một tham chiếu ô/vùng: giữ cây tham chiếu để ROW, OFFSET, ISREF… thấy “tham chiếu” như Excel */
class Ref { constructor(public ast: any) {} }
const REF_T = new Set(['CELL_REFERENCE', 'CELL_RANGE', 'COLUMN_RANGE', 'ROW_RANGE']);
/** giá trị đưa vào tên LET / tham số LAMBDA: tham chiếu giữ dạng tham chiếu, còn lại tính ra giá trị */
function bindArg(interp: any, ast: any, state: any): unknown {
  let a = ast; while (a?.type === 'PARENTHESIS') a = a.expression;
  if (REF_T.has(a?.type)) return new Ref(a);
  if (a?.type === 'NAMED_EXPRESSION') { const b = look(a.expressionName); if (b instanceof Ref) return b; }
  return ev(interp, ast, state);
}
/** thay đối số là tên gắn tham chiếu bằng chính tham chiếu đó (cho hàm cần tham chiếu: ROW, COLUMN, OFFSET, ISREF, CELL…) */
function subst(ast: any): any {
  if (!CUR || !ast.args?.length) return ast;
  let hit = false;
  const args = ast.args.map((x: any) => { if (x?.type === 'NAMED_EXPRESSION') { const b = look(x.expressionName); if (b instanceof Ref) { hit = true; return b.ast; } } return x; });
  return hit ? { ...ast, args } : ast;
}
let CUR: Env | null = null;
let DEPTH = 0;
/** độ sâu gọi lồng tối đa (đệ quy); quá thì #NUM! như Excel */
export const MAX_DEPTH = 5000;

const look = (name: string): unknown => { const k = name.toLowerCase(); for (let e = CUR; e; e = e.parent) if (e.vars.has(k)) return e.vars.get(k); return undefined; };
/** tên hợp lệ cho LET/LAMBDA: không chứa dấu chấm (Excel dùng r, c làm tham số được — ví dụ MAKEARRAY của Microsoft) */
const okName = (n: string) => !n.includes('.');

/** giá trị cuối đưa ra ô: LAMBDA chưa gọi → #CALC! */
const fin = (v: unknown) => (v instanceof Lam ? calcErr() : v);
const truth = (v: unknown): boolean | CellError | undefined => {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === EmptyValue || v === null) return false;
  if (typeof v === 'string') return /^true$/i.test(v) ? true : /^false$/i.test(v) ? false : valErr();
  if (v instanceof CellError) return v;
  if (v && typeof v === 'object' && typeof (v as any).val === 'number') return (v as any).val !== 0;
  return undefined; // mảng: để bộ tính xử lý như cũ
};
const numErr = () => new CellError(ErrorType.NUM, 'Đệ quy quá sâu');
const BIN: Record<string, string> = { CONCATENATE_OP: 'concatOp', EQUALS_OP: 'equalOp', NOT_EQUAL_OP: 'notEqualOp', GREATER_THAN_OP: 'greaterThanOp', LESS_THAN_OP: 'lessThanOp', GREATER_THAN_OR_EQUAL_OP: 'greaterThanOrEqualOp', LESS_THAN_OR_EQUAL_OP: 'lessThanOrEqualOp', PLUS_OP: 'plusOp', MINUS_OP: 'minusOp', TIMES_OP: 'timesOp', POWER_OP: 'powerOp', DIV_OP: 'divOp' };
const UN: Record<string, string> = { PLUS_UNARY_OP: 'unaryPlusOp', MINUS_UNARY_OP: 'unaryMinusOp', PERCENT_OP: 'percentOp' };
const one = (v: unknown) => (v instanceof SimpleRangeValue && v.width() === 1 && v.height() === 1 ? (v.data as unknown[][])[0][0] : v);

/**
 * Tính một biểu thức trong LET/LAMBDA, giữ nguyên giá trị hàm (LAMBDA).
 * Ít khung ngăn xếp: nhánh IF và lời gọi LAMBDA ở vị trí cuối chạy bằng vòng lặp (đệ quy sâu ≥ 1.000 tầng vẫn tính được).
 */
function ev(interp: any, ast: any, state: any): unknown {
  const entry = CUR;
  let calls = 0;
  try {
    for (;;) {
      switch (ast?.type) {
        case 'NUMBER': case 'STRING': return ast.value;
        case 'PARENTHESIS': ast = ast.expression; continue;
        case 'NAMED_EXPRESSION': {
          const b = look(ast.expressionName);
          if (b instanceof Ref) return interp.evaluateAst(b.ast, state);
          if (b !== undefined) return b;
          return namedLambda(interp, ast.expressionName, state) ?? interp.evaluateAst(ast, state);
        }
        case 'FUNCTION_CALL': {
          const n: string = ast.procedureName, args: any[] = ast.args ?? [];
          if (n === 'LAMBDA') return makeLam(ast);
          if (n === 'IF' && args.length >= 2 && args.length <= 3) {
            const c = truth(one(ev(interp, args[0], state)));
            if (c instanceof CellError) return c;
            if (c === undefined) return interp.evaluateAst(ast, state);
            if (!c && !args[2]) return false;
            ast = c ? args[1] : args[2]; continue;
          }
          let lam = look(n);
          const builtin = !!interp.functionRegistry.getFunction(n);
          if (lam !== undefined && !(lam instanceof Lam)) { if (!builtin) return valErr('Không phải hàm'); lam = undefined; } // =LET(sum;5;SUM(sum;1)): SUM( là hàm SUM
          if (lam === undefined && !builtin) lam = namedLambda(interp, n, state);
          if (lam instanceof Lam) {
            const vals = args.map((a) => bindArg(interp, a, state));
            if (vals.length !== lam.params.length) return valErr('Sai số đối số');
            if (DEPTH >= MAX_DEPTH) return numErr();
            DEPTH++; calls++;
            CUR = { vars: new Map(lam.params.map((p, i) => [p, vals[i]])), parent: lam.env };
            ast = lam.body; continue;
          }
          return interp.evaluateAst(subst(ast), state);
        }
        default: {
          const op = BIN[ast?.type];
          if (op) {
            const l = ev(interp, ast.left, state), r = ev(interp, ast.right, state);
            if (l instanceof Lam || r instanceof Lam) return calcErr();
            return interp.binaryRangeWrapper(interp[op], l, r, state);
          }
          const u = UN[ast?.type];
          if (u) { const x = ev(interp, ast.value, state); return x instanceof Lam ? calcErr() : interp.unaryRangeWrapper(interp[u], x, state); }
          return interp.evaluateAst(ast, state);
        }
      }
    }
  } catch (e) { if (e instanceof RangeError) return numErr(); throw e; }
  finally { CUR = entry; DEPTH -= calls; }
}
const evalArg = ev;
function makeLam(ast: any): Lam | CellError {
  const args = ast.args ?? [];
  if (!args.length) return valErr();
  if (args.length - 1 > 253) return valErr('Quá 253 tham số');
  const ps: string[] = [];
  for (const a of args.slice(0, -1)) { if (a?.type !== 'NAMED_EXPRESSION' || !okName(a.expressionName)) return valErr('Tên tham số không hợp lệ'); ps.push(a.expressionName.toLowerCase()); }
  if (new Set(ps).size !== ps.length) return valErr('Trùng tên tham số');
  return new Lam(ps, args[args.length - 1], CUR);
}
const VOL = new WeakSet<object>();
/** tên vùng (Trình Quản lý Tên) có giá trị là LAMBDA */
function namedLambda(interp: any, name: string, state: any): Lam | undefined {
  try {
    const ne = interp.namedExpressions?.nearestNamedExpression(name, state.formulaAddress?.sheet);
    if (!ne) return undefined;
    const g = interp.dependencyGraph, v = g.addressMapping.getCell(ne.address);
    const ast = v?.getFormula?.(g.lazilyTransformingAstService);
    if (ast?.type !== 'FUNCTION_CALL' || ast.procedureName !== 'LAMBDA') return undefined;
    const prev = CUR; CUR = null;
    try {
      const l = makeLam(ast);
      // lời gọi tên LAMBDA không tạo quan hệ phụ thuộc trong bộ tính → ô gọi được coi là biến động (luôn tính lại),
      // để đổi ô mà LAMBDA đọc, đổi/xóa/hoàn tác định nghĩa tên thì kết quả vẫn đúng
      const fv = state.formulaVertex;
      if (l instanceof Lam && fv && !VOL.has(fv)) { VOL.add(fv); g.graph.markNodeAsVolatile(fv); }
      return l instanceof Lam ? l : undefined;
    } finally { CUR = prev; }
  } catch { return undefined; }
}
/** gọi LAMBDA với các giá trị đã tính (MAP, REDUCE…) */
function call(interp: any, lam: Lam, vals: unknown[], state: any): unknown {
  if (vals.length !== lam.params.length) return valErr('Sai số đối số');
  if (DEPTH >= MAX_DEPTH) return numErr();
  const prev = CUR;
  CUR = { vars: new Map(lam.params.map((p, i) => [p, vals[i]])), parent: lam.env };
  DEPTH++;
  try { return ev(interp, lam.body, state); }
  finally { DEPTH--; CUR = prev; }
}
const callAsts = (interp: any, lam: Lam, asts: any[], state: any) => call(interp, lam, asts.map((a) => bindArg(interp, a, state)), state);
/** vá bộ thông dịch: tên cục bộ, gọi hàm qua tên, IF lười khi đang trong LAMBDA (để đệ quy dừng được) */
export function patchInterpreter(proto: any) {
  if (proto.__lam) return;
  const orig = proto.evaluateAstWithoutPostprocessing;
  proto.evaluateAstWithoutPostprocessing = function (this: any, ast: any, state: any) {
    if (ast.type === 'NAMED_EXPRESSION') {
      if (CUR) { const b = look(ast.expressionName); if (b instanceof Ref) return this.evaluateAst(b.ast, state); if (b !== undefined) return fin(b); }
    } else if (ast.type === 'FUNCTION_CALL') {
      const n: string = ast.procedureName;
      // ISREF của HyperFormula chỉ nhận lỗi #REF!: làm như Excel — TRUE khi đối số là tham chiếu
      if (n === 'ISREF' && ast.args?.length === 1) {
        let x = subst(ast).args[0]; while (x?.type === 'PARENTHESIS') x = x.expression;
        if (REF_T.has(x?.type) || (x?.type === 'FUNCTION_CALL' && /^(OFFSET|INDIRECT|XOFS[345]|INDEX)$/.test(x.procedureName))) return true;
        if (x?.type === 'NAMED_EXPRESSION' && !look(x.expressionName)) { const ne = this.namedExpressions?.nearestNamedExpression(x.expressionName, state.formulaAddress?.sheet); if (ne) { const v = this.dependencyGraph.addressMapping.getCell(ne.address); const f = v?.getFormula?.(this.dependencyGraph.lazilyTransformingAstService); if (REF_T.has(f?.type)) return true; } }
        return false;
      }
      if (CUR) {
        const b = look(n);
        if (b instanceof Lam) return fin(callAsts(this, b, ast.args, state));
        if (b !== undefined && !this.functionRegistry.getFunction(n)) return valErr('Không phải hàm');
      }
      if (DEPTH > 0 && n === 'IF' && ast.args.length >= 2 && ast.args.length <= 3) {
        const c = truth(this.evaluateAst(ast.args[0], state));
        if (c instanceof CellError) return c;
        if (c !== undefined) return c ? evalArg(this, ast.args[1], state) : ast.args[2] ? evalArg(this, ast.args[2], state) : false;
      }
      if (!this.functionRegistry.getFunction(n)) { const g = namedLambda(this, n, state); if (g) return fin(callAsts(this, g, ast.args, state)); }
      if (CUR) return orig.call(this, subst(ast), state);
    }
    return orig.call(this, ast, state);
  };
  proto.__lam = true;
}

type Grid = unknown[][];
const A = FunctionArgumentType.ANY;
const gridOf = (v: unknown): Grid => (v instanceof SimpleRangeValue ? (v.data as Grid) : [[v]]);
const scalar = (v: unknown) => (v instanceof SimpleRangeValue ? (v.width() === 1 && v.height() === 1 ? (v.data as Grid)[0][0] : calcErr()) : fin(v));
const arr = (g: Grid) => SimpleRangeValue.onlyValues(g as any);

export class LambdaPlugin extends FunctionPlugin {
  static implementedFunctions = {
    LET: { method: 'let', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    LAMBDA: { method: 'lambda', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    XLAMCALL: { method: 'xlamcall', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    MAP: { method: 'map', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    REDUCE: { method: 'reduce', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    SCAN: { method: 'scan', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    BYROW: { method: 'byrow', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    BYCOL: { method: 'bycol', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    MAKEARRAY: { method: 'makearray', parameters: [{ argumentType: A }], repeatLastArgs: 1 },
    ISOMITTED: { method: 'isomitted', parameters: [{ argumentType: A }] },
  };
  private get I() { return (this as any).interpreter; }
  let(ast: any, state: any) {
    const args = ast.args ?? [];
    if (args.length < 3 || args.length % 2 === 0 || args.length > 253) return valErr('LET cần các cặp tên–giá trị và phép tính');
    const prev = CUR;
    const env: Env = { vars: new Map(), parent: CUR };
    CUR = env;
    try {
      for (let i = 0; i < args.length - 1; i += 2) {
        const nm = args[i];
        // Microsoft: tên LET không được là r/c (dành cho kiểu R1C1); tham số LAMBDA thì được (vd MAKEARRAY(…;LAMBDA(r;c;…)))
        if (nm?.type !== 'NAMED_EXPRESSION' || !okName(nm.expressionName) || /^[rc]$/i.test(nm.expressionName)) return valErr('Tên trong LET không hợp lệ');
        env.vars.set(nm.expressionName.toLowerCase(), bindArg(this.I, args[i + 1], state));
      }
      return fin(evalArg(this.I, args[args.length - 1], state));
    } finally { CUR = prev; }
  }
  /** LAMBDA đứng một mình (không gọi) → #CALC! */
  lambda(ast: any) { const l = makeLam(ast); return l instanceof Lam ? calcErr() : l; }
  xlamcall(ast: any, state: any) {
    const f = evalArg(this.I, ast.args[0], state);
    if (!(f instanceof Lam)) return f instanceof CellError ? f : valErr();
    return fin(callAsts(this.I, f, ast.args.slice(1), state));
  }
  private fn(ast: any, state: any): Lam | CellError {
    const f = evalArg(this.I, ast.args[ast.args.length - 1], state);
    return f instanceof Lam ? f : f instanceof CellError ? f : valErr('Cần một LAMBDA');
  }
  map(ast: any, state: any) {
    if (ast.args.length < 2) return valErr();
    const f = this.fn(ast, state); if (f instanceof CellError) return f;
    const gs = ast.args.slice(0, -1).map((a: any) => gridOf(this.I.evaluateAst(a, state)));
    const h = Math.max(...gs.map((g: Grid) => g.length)), w = Math.max(...gs.map((g: Grid) => g[0]?.length ?? 0));
    const at = (g: Grid, i: number, j: number) => (g.length === 1 && g[0].length === 1 ? g[0][0] : g[i]?.[j]);
    if (gs.some((g: Grid) => !(g.length === 1 && g[0].length === 1) && (g.length !== h || g[0].length !== w))) return valErr('Các mảng phải cùng kích thước');
    const out: Grid = [];
    for (let i = 0; i < h; i++) { const row: unknown[] = []; for (let j = 0; j < w; j++) row.push(scalar(call(this.I, f, gs.map((g: Grid) => at(g, i, j) ?? new CellError(ErrorType.NA)), state))); out.push(row); }
    return arr(out);
  }
  reduce(ast: any, state: any) {
    if (ast.args.length !== 3) return valErr();
    const f = this.fn(ast, state); if (f instanceof CellError) return f;
    let acc = evalArg(this.I, ast.args[0], state);
    for (const row of gridOf(this.I.evaluateAst(ast.args[1], state))) for (const v of row) { acc = call(this.I, f, [acc, v], state); if (acc instanceof Lam) return calcErr(); }
    return fin(acc);
  }
  scan(ast: any, state: any) {
    if (ast.args.length !== 3) return valErr();
    const f = this.fn(ast, state); if (f instanceof CellError) return f;
    let acc = evalArg(this.I, ast.args[0], state);
    const out: Grid = gridOf(this.I.evaluateAst(ast.args[1], state)).map((row) => row.map((v) => { acc = call(this.I, f, [acc, v], state); return scalar(acc); }));
    return arr(out);
  }
  byrow(ast: any, state: any) {
    if (ast.args.length !== 2) return valErr();
    const f = this.fn(ast, state); if (f instanceof CellError) return f;
    const g = gridOf(this.I.evaluateAst(ast.args[0], state));
    return arr(g.map((row) => [scalar(call(this.I, f, [arr([row])], state))]));
  }
  bycol(ast: any, state: any) {
    if (ast.args.length !== 2) return valErr();
    const f = this.fn(ast, state); if (f instanceof CellError) return f;
    const g = gridOf(this.I.evaluateAst(ast.args[0], state));
    return arr([Array.from({ length: g[0]?.length ?? 0 }, (_, j) => scalar(call(this.I, f, [arr(g.map((row) => [row[j]]))], state)))]);
  }
  makearray(ast: any, state: any) {
    if (ast.args.length !== 3) return valErr();
    const f = this.fn(ast, state); if (f instanceof CellError) return f;
    const num = (a: any) => { const v = scalar(this.I.evaluateAst(a, state)); return typeof v === 'number' ? Math.trunc(v) : v && typeof v === 'object' && typeof (v as any).val === 'number' ? Math.trunc((v as any).val) : NaN; };
    const h = num(ast.args[0]), w = num(ast.args[1]);
    if (!(h >= 1 && w >= 1) || h * w > 1_048_576) return valErr('Số hàng/cột không hợp lệ');
    const out: Grid = [];
    for (let i = 1; i <= h; i++) { const row: unknown[] = []; for (let j = 1; j <= w; j++) row.push(scalar(call(this.I, f, [i, j], state))); out.push(row); }
    return arr(out);
  }
  /** ExcelLike chưa hỗ trợ tham số tùy chọn [x] của LAMBDA: luôn FALSE */
  isomitted() { return false; }
}
export const LAMBDA_NAMES = ['LET', 'LAMBDA', 'XLAMCALL', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'ISOMITTED'];

// ---------- lớp viết lại chữ công thức ----------
/** vị trí ngoặc đóng khớp với ngoặc mở tại i (bỏ qua chuỗi "…", tên trang '…', [..]) */
export function matchParen(f: string, i: number): number {
  let d = 0;
  for (let j = i; j < f.length; j++) {
    const ch = f[j];
    if (ch === '"' || ch === "'") { j++; while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; } j++; } continue; }
    if (ch === '(') d++;
    else if (ch === ')') { d--; if (!d) return j; }
  }
  return -1;
}
/** LAMBDA(x;x+1)(5) → XLAMCALL(LAMBDA(x;x+1);5) (đệ quy, ngoài chuỗi) */
export function lamIn(f: string, sep = ';'): string {
  if (!/LAMBDA\s*\(/i.test(f)) return f;
  const re = /(^|[^A-Za-z0-9_.\\])(?:_xlfn\.)?LAMBDA\s*\(/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(f))) {
    const s = m.index + m[1].length;
    if (inString(f, s)) continue;
    const open = f.indexOf('(', s), close = matchParen(f, open);
    if (close < 0) return f;
    let k = close + 1; while (f[k] === ' ') k++;
    if (f[k] !== '(') continue;
    const c2 = matchParen(f, k);
    if (c2 < 0) return f;
    const lam = f.slice(s, close + 1), args = f.slice(k + 1, c2);
    const rep = `XLAMCALL(${lamIn(lam, sep)}${args.trim() ? sep + lamIn(args, sep) : ''})`;
    f = f.slice(0, s) + rep + f.slice(c2 + 1);
    re.lastIndex = s + rep.length;
  }
  return f;
}
/** XLAMCALL(LAMBDA(…);a;b) → LAMBDA(…)(a;b) */
export function lamOut(f: string, sep = ';'): string {
  if (!/XLAMCALL\(/i.test(f)) return f;
  const re = /XLAMCALL\(/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(f))) {
    if (inString(f, m.index)) continue;
    const open = m.index + m[0].length - 1, close = matchParen(f, open);
    if (close < 0) return f;
    const inner = f.slice(open + 1, close);
    const lo = inner.search(/\S/), lc = matchParen(inner, inner.indexOf('('));
    if (lo < 0 || lc < 0) continue;
    const lam = inner.slice(lo, lc + 1);
    let rest = inner.slice(lc + 1).trimStart();
    if (rest.startsWith(sep)) rest = rest.slice(1);
    const rep = `${lamOut(lam, sep)}(${lamOut(rest, sep)})`;
    f = f.slice(0, m.index) + rep + f.slice(close + 1);
    re.lastIndex = m.index + rep.length;
  }
  return f;
}
function inString(f: string, at: number): boolean {
  let q = '';
  for (let i = 0; i < at; i++) { const ch = f[i]; if (q) { if (ch === q) { if (f[i + 1] === q) i++; else q = ''; } } else if (ch === '"' || ch === "'") q = ch; }
  return !!q;
}
/** tách đối số cấp đầu của lời gọi có ngoặc mở tại open: trả [vị trí đầu, cuối) từng đối số */
function argSpans(f: string, open: number, sep: string): [number, number][] {
  const close = matchParen(f, open), out: [number, number][] = [];
  let d = 0, b = 0, st = open + 1;
  for (let j = open + 1; j < close; j++) {
    const ch = f[j];
    if (ch === '"' || ch === "'") { j++; while (j < close) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; } j++; } continue; }
    if (ch === '(') d++; else if (ch === ')') d--; else if (ch === '{') b++; else if (ch === '}') b--;
    else if (ch === sep && !d && !b) { out.push([st, j]); st = j + 1; }
  }
  out.push([st, close]);
  return out;
}
/** các tên khai báo trong LET và tham số LAMBDA của công thức (viết hoa) */
export function declaredNames(f: string, sep = ';'): Set<string> { return new Set(declaredMap(f, sep).keys()); }
/** tên khai báo: dạng viết hoa → cách viết gốc */
/** phạm vi của từng LET/LAMBDA: [mở ngoặc, đóng ngoặc] và các tên khai báo (viết hoa) — tên chỉ có hiệu lực trong ngoặc đó */
export function declScopes(f: string, sep = ';'): { open: number; close: number; names: Set<string> }[] {
  const out: { open: number; close: number; names: Set<string> }[] = [];
  const re = /(^|[^A-Za-z0-9_.\\])(?:_xlfn\.)?(LET|LAMBDA)\s*\(/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(f))) {
    if (inString(f, m.index + m[1].length)) continue;
    const open = f.indexOf('(', m.index + m[1].length), sp = argSpans(f, open, sep), let_ = m[2].toUpperCase() === 'LET';
    const names = new Set<string>();
    sp.slice(0, -1).forEach(([a, b], i) => { if (let_ && i % 2) return; const t = f.slice(a, b).trim().replace(/^_xlpm\./i, ''); if (/^[A-Za-z_\\À-￿][\w.\\À-￿]*$/.test(t)) names.add(t.toUpperCase()); });
    out.push({ open, close: matchParen(f, open), names });
  }
  return out;
}
export function declaredMap(f: string, sep = ';'): Map<string, string> {
  const names = new Map<string, string>();
  const re = /(^|[^A-Za-z0-9_.\\])(?:_xlfn\.)?(LET|LAMBDA)\s*\(/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(f))) {
    if (inString(f, m.index + m[1].length)) continue;
    const open = f.indexOf('(', m.index + m[1].length), sp = argSpans(f, open, sep), let_ = m[2].toUpperCase() === 'LET';
    sp.slice(0, -1).forEach(([a, b], i) => { if (let_ && i % 2) return; const t = f.slice(a, b).trim().replace(/^_xlpm\./i, ''); if (/^[A-Za-z_\\À-￿][\w.\\À-￿]*$/.test(t) && !names.has(t.toUpperCase())) names.set(t.toUpperCase(), t); });
  }
  return names;
}
/** bộ tính viết hoa tên hàm: f(3) → F(3); đưa lời gọi tên LET/LAMBDA về cách viết lúc khai báo */
export function callCase(f: string, mapIdents: (f: string, fn: (tok: string, next: string, prev: string) => string | undefined) => string, names?: Map<string, string>, builtin?: (n: string) => boolean): string {
  const m = /LET\s*\(|LAMBDA\s*\(/i.test(f) ? declaredMap(f) : new Map<string, string>();
  const g = names?.size ? names : null;
  if (!m.size && !g) return f;
  return mapIdents(f, (tok, next) => {
    if (next !== '(') return undefined;
    const u = tok.toUpperCase();
    if (builtin?.(u)) return undefined; // =LET(sum;5;SUM(sum;1)): SUM( là hàm SUM
    if (m.has(u)) return m.get(u);
    if (g?.has(u) && !builtin?.(u)) return g.get(u); // gọi tên LAMBDA trong Trình Quản lý Tên
    return undefined;
  });
}
