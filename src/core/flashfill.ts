// Điền Nhanh (Ctrl+E) — học mẫu từ ví dụ người dùng gõ, áp cho các hàng còn lại.
// Chương trình = dãy đoạn: chữ cố định | cả ô | một "từ"/"số" thứ i (đếm từ đầu hoặc từ cuối) | chữ cái đầu của từ đó,
// mỗi đoạn có thể đổi HOA / thường / Hoa Đầu Từ. Chọn chương trình khớp mọi ví dụ và áp được cho nhiều hàng nhất.

type Case = 'id' | 'up' | 'lo' | 'pr';
type Kind = 'w' | 'n' | 'a';
type Seg =
  | { t: 'lit'; s: string }
  | { t: 'all'; col: number; cs: Case }
  | { t: 'tok'; col: number; kind: Kind; idx: number; cs: Case; first: boolean };

interface Tok { s: string; kind: 'w' | 'n' }
const tokCache = new Map<string, Tok[]>();
function tokens(s: string): Tok[] {
  let t = tokCache.get(s);
  if (t) return t;
  t = [];
  for (const m of s.matchAll(/[\p{L}\p{M}]+|\d+/gu)) t.push({ s: m[0], kind: /^\d/.test(m[0]) ? 'n' : 'w' });
  if (tokCache.size > 50000) tokCache.clear();
  tokCache.set(s, t);
  return t;
}
const proper = (s: string) => s.toLocaleLowerCase('vi').replace(/(^|[^\p{L}\p{M}])(\p{L})/gu, (_, a, b) => a + b.toLocaleUpperCase('vi'));
function applyCase(s: string, cs: Case) {
  return cs === 'up' ? s.toLocaleUpperCase('vi') : cs === 'lo' ? s.toLocaleLowerCase('vi') : cs === 'pr' ? proper(s) : s;
}
function evalSeg(g: Seg, row: string[]): string | null {
  if (g.t === 'lit') return g.s;
  const src = row[g.col] ?? '';
  if (g.t === 'all') return src ? applyCase(src, g.cs) : null;
  const list = tokens(src).filter((x) => g.kind === 'a' || x.kind === g.kind);
  const k = g.idx >= 0 ? g.idx : list.length + g.idx;
  const tk = list[k];
  if (!tk) return null;
  const v = g.first ? tk.s[0] : tk.s;
  return applyCase(v, g.cs);
}
function run(p: Seg[], row: string[]): string | null {
  let out = '';
  for (const g of p) { const v = evalSeg(g, row); if (v === null) return null; out += v; }
  return out;
}
/** Mọi đoạn (không phải chữ cố định) có thể sinh ra từ một hàng */
function candidates(row: string[]): { g: Seg; v: string }[] {
  const out: { g: Seg; v: string }[] = [];
  const cases: Case[] = ['id', 'up', 'lo', 'pr'];
  row.forEach((src, col) => {
    if (!src) return;
    for (const cs of cases) out.push({ g: { t: 'all', col, cs }, v: applyCase(src, cs) });
    const toks = tokens(src);
    for (const kind of ['w', 'n', 'a'] as Kind[]) {
      const list = toks.filter((x) => kind === 'a' || x.kind === kind);
      list.forEach((tk, i) => {
        for (const idx of [i, i - list.length]) for (const cs of cases) for (const first of [false, true]) {
          if (first && kind === 'n') continue;
          const v = applyCase(first ? tk.s[0] : tk.s, cs);
          out.push({ g: { t: 'tok', col, kind, idx, cs, first }, v });
        }
      });
    }
  });
  // đoạn dài trước (ít mảnh hơn), cả ô trước từng từ
  return out.sort((a, b) => b.v.length - a.v.length);
}

/**
 * rows: các cột nguồn của từng hàng; examples: kết quả người dùng đã gõ cho các hàng đầu.
 * Trả về kết quả cho mọi hàng (hàng không áp được = null), hoặc null nếu không tìm ra mẫu.
 */
export function flashFill(rows: string[][], examples: string[]): (string | null)[] | null {
  if (!examples.length || !rows.length) return null;
  const n = examples.length;
  // gộp các đoạn cho cùng kết quả trên mọi ví dụ (giữ đoạn ưu tiên hơn) để cây tìm kiếm không nổ
  const seen = new Set<string>();
  const cand = candidates(rows[0]).filter(({ g }) => {
    // khoá = kết quả trên các ví dụ và vài hàng đầu: đoạn khác hành vi ở hàng khác vẫn được giữ
    const k = rows.slice(0, Math.max(examples.length, 24)).map((r) => evalSeg(g, r) ?? '\u0000').join('\u0001');
    if (seen.has(k)) return false;
    seen.add(k); return true;
  });
  let nodes = 0;
  const t0 = Date.now();
  const progs: Seg[][] = [];
  const LIMIT = 3000;
  // tìm đồng thời trên mọi ví dụ: mỗi đoạn phải khớp ở vị trí hiện tại của tất cả ví dụ
  const pos = new Array(n).fill(0);
  const dfs = (cur: Seg[]) => {
    if (progs.length >= LIMIT || ++nodes > 60000 || (nodes % 1024 === 0 && Date.now() - t0 > 800)) { nodes = Infinity; return; }
    if (pos.every((p, i) => p === examples[i].length)) { progs.push(cur.map((g) => ({ ...g }))); return; }
    for (const { g } of cand) {
      const vs: string[] = [];
      let good = true;
      for (let i = 0; i < n && good; i++) {
        const v = evalSeg(g, rows[i]);
        if (!v || !examples[i].startsWith(v, pos[i])) good = false; else vs.push(v);
      }
      if (!good) continue;
      vs.forEach((v, i) => (pos[i] += v.length));
      cur.push(g); dfs(cur); cur.pop();
      vs.forEach((v, i) => (pos[i] -= v.length));
      if (progs.length >= LIMIT) return;
    }
    // chữ cố định: cùng ký tự ở mọi ví dụ
    const ch = examples[0][pos[0]];
    if (ch === undefined || !examples.every((ex, i) => ex[pos[i]] === ch)) return;
    const last = cur[cur.length - 1];
    pos.forEach((_, i) => pos[i]++);
    if (last?.t === 'lit') { last.s += ch; dfs(cur); last.s = last.s.slice(0, -1); }
    else { cur.push({ t: 'lit', s: ch }); dfs(cur); cur.pop(); }
    pos.forEach((_, i) => pos[i]--);
  };
  dfs([]);
  const ok = progs.filter((p) => examples.every((ex, i) => run(p, rows[i]) === ex));
  if (!ok.length) return null;
  const litLen = (p: Seg[]) => p.reduce((n, g) => n + (g.t === 'lit' ? g.s.length : 0), 0);
  // chấm điểm trên mẫu tối đa 200 hàng (sổ lớn không bị treo)
  const step = Math.max(1, Math.floor(rows.length / 200));
  const sample = rows.filter((_, i) => i % step === 0).slice(0, 200);
  let best: Seg[] | null = null, bestScore: number[] = [];
  for (const p of ok) {
    const covered = sample.reduce((n, r) => n + (run(p, r) !== null ? 1 : 0), 0);
    // ít chữ cố định → nhiều hàng áp được nhất → ít mảnh → ưu tiên đếm từ đầu
    const fromEnd = p.reduce((n, g) => n + (g.t === 'tok' && g.idx < 0 ? 1 : 0), 0);
    const score = [litLen(p), -covered, p.length, fromEnd];
    if (!best || lexLess(score, bestScore)) { best = p; bestScore = score; }
  }
  return rows.map((r, i) => (i < examples.length ? examples[i] : run(best!, r)));
}

function lexLess(a: number[], b: number[]) { for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i] < b[i]; return false; }
