// Vẽ biểu đồ lên canvas (cột, cột chồng, thanh ngang, đường, vùng, tròn, vành khuyên)
import { compactNum, niceScale, PALETTE, type Chart, type ChartData } from '../core/chart';

export interface ChartTheme { bg: string; fg: string; muted: string; grid: string; font: string }
export const LIGHT: ChartTheme = { bg: '#ffffff', fg: '#1f2a24', muted: '#5b6a62', grid: '#e3e8e5', font: "'Be Vietnam Pro', Arial, sans-serif" };
export const DARK: ChartTheme = { bg: '#141b17', fg: '#e3ebe6', muted: '#93a49b', grid: '#29342e', font: LIGHT.font };

const fit = (ctx: CanvasRenderingContext2D, t: string, w: number) => {
  if (ctx.measureText(t).width <= w) return t;
  let s = t;
  while (s.length > 1 && ctx.measureText(s + '…').width > w) s = s.slice(0, -1);
  return s + '…';
};

export function drawChart(ctx: CanvasRenderingContext2D, W: number, H: number, ch: Pick<Chart, 'type' | 'title' | 'legend' | 'labels'>, d: ChartData, T: ChartTheme) {
  ctx.save();
  ctx.fillStyle = T.bg;
  ctx.fillRect(0, 0, W, H);
  const pad = 12;
  let top = pad, bottom = H - pad, left = pad, right = W - pad;
  const pie = ch.type === 'pie' || ch.type === 'doughnut';
  // tiêu đề
  if (ch.title) {
    ctx.font = `600 14px ${T.font}`;
    ctx.fillStyle = T.fg;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText(fit(ctx, ch.title, W - 2 * pad), W / 2, top);
    top += 24;
  }
  // chú giải
  const items = pie ? d.cats.map((c, i) => ({ name: c, color: PALETTE[i % PALETTE.length] })) : d.series.map((s, i) => ({ name: s.name, color: PALETTE[i % PALETTE.length] }));
  const legend = ch.legend ?? 'b';
  ctx.font = `12px ${T.font}`;
  ctx.textBaseline = 'middle';
  if (legend !== 'none' && items.length && (items.length > 1 || pie)) {
    if (legend === 'r') {
      const lw = Math.min(W * 0.32, Math.max(...items.map((it) => ctx.measureText(it.name).width)) + 26);
      let y = top + 6;
      for (const it of items) {
        if (y > bottom - 8) break;
        ctx.fillStyle = it.color; ctx.fillRect(right - lw, y - 5, 10, 10);
        ctx.fillStyle = T.fg; ctx.textAlign = 'left';
        ctx.fillText(fit(ctx, it.name, lw - 16), right - lw + 15, y);
        y += 18;
      }
      right -= lw + 8;
    } else {
      // xếp hàng, xuống dòng khi hết chỗ
      const rows: { name: string; color: string; w: number }[][] = [[]];
      let lineW = 0;
      for (const it of items) {
        const w = Math.min(ctx.measureText(it.name).width, 160) + 30;
        if (lineW + w > W - 2 * pad && rows[rows.length - 1].length) { rows.push([]); lineW = 0; }
        rows[rows.length - 1].push({ ...it, w }); lineW += w;
      }
      const shown = rows.slice(0, 3);
      const lh = 18;
      let y = legend === 't' ? top + 8 : bottom - shown.length * lh + 8;
      for (const row of shown) {
        let x = W / 2 - row.reduce((a, r) => a + r.w, 0) / 2;
        for (const it of row) {
          ctx.fillStyle = it.color; ctx.fillRect(x, y - 5, 10, 10);
          ctx.fillStyle = T.fg; ctx.textAlign = 'left';
          ctx.fillText(fit(ctx, it.name, it.w - 30), x + 15, y);
          x += it.w;
        }
        y += lh;
      }
      if (legend === 't') top += shown.length * lh + 6; else bottom -= shown.length * lh + 6;
    }
  }
  if (!d.series.length || !d.cats.length) {
    ctx.fillStyle = T.muted; ctx.textAlign = 'center';
    ctx.fillText('Chưa có dữ liệu số trong vùng đã chọn', W / 2, (top + bottom) / 2);
    ctx.restore();
    return;
  }
  if (pie) { drawPie(ctx, ch, d, T, left, top, right, bottom); ctx.restore(); return; }

  const horiz = ch.type === 'bar';
  const stacked = ch.type === 'colStack';
  const nCat = d.cats.length;
  let lo = 0, hi = 0;
  for (let i = 0; i < nCat; i++) {
    if (stacked) {
      let pos = 0, neg = 0;
      for (const s of d.series) { const v = s.values[i] ?? 0; if (v >= 0) pos += v; else neg += v; }
      hi = Math.max(hi, pos); lo = Math.min(lo, neg);
    } else for (const s of d.series) { const v = s.values[i]; if (v !== null && v !== undefined) { hi = Math.max(hi, v); lo = Math.min(lo, v); } }
  }
  const sc = niceScale(lo, hi, horiz ? 4 : 5);
  ctx.font = `11px ${T.font}`;
  const ticks: number[] = [];
  for (let v = sc.lo; v <= sc.hi + sc.step / 2; v += sc.step) ticks.push(+v.toPrecision(12));
  const tickW = Math.max(...ticks.map((t) => ctx.measureText(compactNum(t)).width));
  const catW = horiz ? Math.min(W * 0.3, Math.max(...d.cats.map((c) => ctx.measureText(c).width)) + 8) : 0;
  const pl = left + (horiz ? catW : tickW + 8), pr = right - 4, pt = top + 6;
  const pb = bottom - (horiz ? 18 : 30);
  const pw = pr - pl, ph = pb - pt;
  if (pw < 20 || ph < 20) { ctx.restore(); return; }
  const val2 = (v: number) => (horiz ? pl + ((v - sc.lo) / (sc.hi - sc.lo)) * pw : pb - ((v - sc.lo) / (sc.hi - sc.lo)) * ph);
  // lưới + nhãn trục giá trị
  ctx.strokeStyle = T.grid; ctx.lineWidth = 1; ctx.fillStyle = T.muted;
  for (const t of ticks) {
    const p = Math.round(val2(t)) + 0.5;
    ctx.beginPath();
    if (horiz) { ctx.moveTo(p, pt); ctx.lineTo(p, pb); } else { ctx.moveTo(pl, p); ctx.lineTo(pr, p); }
    ctx.stroke();
    ctx.textAlign = horiz ? 'center' : 'right';
    ctx.textBaseline = horiz ? 'top' : 'middle';
    if (horiz) ctx.fillText(compactNum(t), p, pb + 4); else ctx.fillText(compactNum(t), pl - 6, p);
  }
  // trục 0
  ctx.strokeStyle = T.muted;
  ctx.beginPath();
  const z = Math.round(val2(0)) + 0.5;
  if (horiz) { ctx.moveTo(z, pt); ctx.lineTo(z, pb); } else { ctx.moveTo(pl, z); ctx.lineTo(pr, z); }
  ctx.stroke();
  // nhãn danh mục
  const band = (horiz ? ph : pw) / nCat;
  ctx.fillStyle = T.muted;
  const every = Math.max(1, Math.ceil((horiz ? 14 : 48) / band));
  for (let i = 0; i < nCat; i += every) {
    const c = (i + 0.5) * band;
    if (horiz) { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(fit(ctx, d.cats[i], catW - 6), pl - 6, pt + c); }
    else { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(fit(ctx, d.cats[i], band * every - 4), pl + c, pb + 6); }
  }
  // dữ liệu
  const nS = d.series.length;
  ctx.font = `10.5px ${T.font}`;
  if (ch.type === 'line' || ch.type === 'area') {
    d.series.forEach((s, si) => {
      const color = PALETTE[si % PALETTE.length];
      const pts = s.values.map((v, i) => (v === null ? null : { x: pl + (i + 0.5) * band, y: val2(v) }));
      if (ch.type === 'area') {
        ctx.fillStyle = color; ctx.globalAlpha = 0.25;
        ctx.beginPath();
        let started = false;
        pts.forEach((p) => { if (!p) return; if (!started) { ctx.moveTo(p.x, z); started = true; } ctx.lineTo(p.x, p.y); });
        const last = [...pts].reverse().find(Boolean);
        if (last) { ctx.lineTo(last.x, z); ctx.closePath(); ctx.fill(); }
        ctx.globalAlpha = 1;
      }
      ctx.strokeStyle = color; ctx.lineWidth = 2.25; ctx.lineJoin = 'round';
      ctx.beginPath();
      let pen = false;
      for (const p of pts) { if (!p) { pen = false; continue; } if (!pen) { ctx.moveTo(p.x, p.y); pen = true; } else ctx.lineTo(p.x, p.y); }
      ctx.stroke();
      if (ch.type === 'line') for (const p of pts) if (p) { ctx.fillStyle = T.bg; ctx.beginPath(); ctx.arc(p.x, p.y, 3.5, 0, Math.PI * 2); ctx.fill(); ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.stroke(); }
      if (ch.labels) { ctx.fillStyle = T.fg; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; s.values.forEach((v, i) => { if (v !== null && pts[i]) ctx.fillText(compactNum(v), pts[i]!.x, pts[i]!.y - 5); }); }
    });
  } else {
    const inner = band * 0.7;
    const bw = stacked ? inner : inner / nS;
    for (let i = 0; i < nCat; i++) {
      let pos = 0, neg = 0;
      d.series.forEach((s, si) => {
        const v = s.values[i];
        if (v === null || v === undefined) return;
        const base = stacked ? (v >= 0 ? pos : neg) : 0;
        const a = val2(base), b = val2(base + v);
        if (stacked) { if (v >= 0) pos += v; else neg += v; }
        const off = (i + 0.15) * band + (stacked ? 0 : si * bw);
        ctx.fillStyle = PALETTE[si % PALETTE.length];
        if (horiz) ctx.fillRect(Math.min(a, b), pt + off, Math.abs(b - a), Math.max(1, bw - 1));
        else ctx.fillRect(pl + off, Math.min(a, b), Math.max(1, bw - 1), Math.abs(b - a));
        if (ch.labels && Math.abs(b - a) > 2) {
          ctx.fillStyle = stacked ? '#ffffff' : T.fg;
          if (horiz) { ctx.textAlign = stacked ? 'center' : 'left'; ctx.textBaseline = 'middle'; ctx.fillText(compactNum(v), stacked ? (a + b) / 2 : Math.max(a, b) + 4, pt + off + bw / 2); }
          else { ctx.textAlign = 'center'; ctx.textBaseline = stacked ? 'middle' : 'bottom'; ctx.fillText(compactNum(v), pl + off + bw / 2, stacked ? (a + b) / 2 : Math.min(a, b) - 3); }
        }
      });
    }
  }
  ctx.restore();
}

function drawPie(ctx: CanvasRenderingContext2D, ch: Pick<Chart, 'type' | 'labels'>, d: ChartData, T: ChartTheme, l: number, t: number, r: number, b: number) {
  const vals = (d.series[0]?.values ?? []).map((v) => (v && v > 0 ? v : 0));
  const total = vals.reduce((a, v) => a + v, 0);
  const cx = (l + r) / 2, cy = (t + b) / 2, R = Math.max(10, Math.min(r - l, b - t) / 2 - 6);
  if (!total) { ctx.fillStyle = T.muted; ctx.textAlign = 'center'; ctx.fillText('Không có giá trị dương để vẽ', cx, cy); return; }
  let a = -Math.PI / 2;
  vals.forEach((v, i) => {
    const da = (v / total) * Math.PI * 2;
    ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, a, a + da); ctx.closePath();
    ctx.fillStyle = PALETTE[i % PALETTE.length]; ctx.fill();
    ctx.strokeStyle = T.bg; ctx.lineWidth = 1.5; ctx.stroke();
    a += da;
  });
  if (ch.type === 'doughnut') { ctx.fillStyle = T.bg; ctx.beginPath(); ctx.arc(cx, cy, R * 0.55, 0, Math.PI * 2); ctx.fill(); }
  if (ch.labels !== false) {
    a = -Math.PI / 2;
    ctx.font = `600 11px ${T.font}`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    vals.forEach((v) => {
      const da = (v / total) * Math.PI * 2;
      const pct = (v / total) * 100;
      if (pct >= 4) {
        const m = a + da / 2, rr = ch.type === 'doughnut' ? R * 0.78 : R * 0.64;
        ctx.fillStyle = '#ffffff';
        ctx.fillText(pct.toFixed(pct < 10 ? 1 : 0).replace('.', ',') + '%', cx + Math.cos(m) * rr, cy + Math.sin(m) * rr);
      }
      a += da;
    });
  }
}

/** Ảnh PNG của biểu đồ (dùng khi in / lưu ảnh) */
export function chartPng(ch: Chart, d: ChartData, scale = 2): string {
  const c = document.createElement('canvas');
  c.width = ch.w * scale; c.height = ch.h * scale;
  const ctx = c.getContext('2d')!;
  ctx.scale(scale, scale);
  drawChart(ctx, ch.w, ch.h, ch, d, LIGHT);
  return c.toDataURL('image/png');
}
