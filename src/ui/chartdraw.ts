// Vẽ biểu đồ lên canvas (cột, cột chồng, thanh ngang, đường, vùng, tròn, vành khuyên)
import { compactNum, isObj, niceScale, paletteOf, type Chart, type ChartData } from '../core/chart';

export interface ChartTheme { bg: string; fg: string; muted: string; grid: string; font: string }
export const LIGHT: ChartTheme = { bg: '#ffffff', fg: '#1f2a24', muted: '#5b6a62', grid: '#e3e8e5', font: "'Be Vietnam Pro', Arial, sans-serif" };
export const DARK: ChartTheme = { bg: '#141b17', fg: '#e3ebe6', muted: '#93a49b', grid: '#29342e', font: LIGHT.font };

const fit = (ctx: CanvasRenderingContext2D, t: string, w: number) => {
  if (ctx.measureText(t).width <= w) return t;
  let s = t;
  while (s.length > 1 && ctx.measureText(s + '…').width > w) s = s.slice(0, -1);
  return s + '…';
};

export type DrawSpec = Pick<Chart, 'type' | 'title' | 'legend' | 'labels'> & Partial<Pick<Chart, 'xTitle' | 'yTitle' | 'y2Title' | 'grid' | 'colors' | 'fill' | 'border' | 'series' | 'img' | 'shape' | 'text' | 'fs' | 'fg'>>;
/** ảnh đã nạp (vẽ đồng bộ lên canvas); onLoad để lớp đối tượng vẽ lại khi ảnh nạp xong */
const IMGS = new Map<string, HTMLImageElement>();
export let onImgLoad: () => void = () => {};
export const setImgLoad = (f: () => void) => { onImgLoad = f; };
export function imgOf(src: string): HTMLImageElement | null {
  let im = IMGS.get(src);
  if (!im) { im = new Image(); im.onload = () => onImgLoad(); im.src = src; IMGS.set(src, im); if (IMGS.size > 200) IMGS.delete(IMGS.keys().next().value!); }
  return im.complete && im.naturalWidth ? im : null;
}
/** ảnh, hình dạng, hộp văn bản */
function drawObj(ctx: CanvasRenderingContext2D, W: number, H: number, ch: DrawSpec, T: ChartTheme) {
  if (ch.type === 'pic') {
    const im = ch.img ? imgOf(ch.img) : null;
    if (im) ctx.drawImage(im, 0, 0, W, H); else { ctx.fillStyle = '#f2f2f2'; ctx.fillRect(0, 0, W, H); }
    if (ch.border) { ctx.strokeStyle = ch.border; ctx.lineWidth = 1.5; ctx.strokeRect(0.75, 0.75, W - 1.5, H - 1.5); }
    return;
  }
  const fill = ch.type === 'text' ? ch.fill || '#ffffff' : ch.fill || '#4472c4';
  const line = ch.border ?? (ch.type === 'text' ? '#a6a6a6' : '#2f528f');
  const p = 1.5, w = W - 2 * p, h = H - 2 * p;
  ctx.save(); ctx.translate(p, p);
  ctx.beginPath();
  const k = ch.type === 'text' ? 'rect' : ch.shape ?? 'rect';
  if (k === 'rect') ctx.rect(0, 0, w, h);
  else if (k === 'roundRect') { const r = Math.min(w, h) * 0.16; ctx.moveTo(r, 0); ctx.arcTo(w, 0, w, h, r); ctx.arcTo(w, h, 0, h, r); ctx.arcTo(0, h, 0, 0, r); ctx.arcTo(0, 0, w, 0, r); ctx.closePath(); }
  else if (k === 'ellipse') ctx.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
  else if (k === 'triangle') { ctx.moveTo(w / 2, 0); ctx.lineTo(w, h); ctx.lineTo(0, h); ctx.closePath(); }
  else if (k === 'diamond') { ctx.moveTo(w / 2, 0); ctx.lineTo(w, h / 2); ctx.lineTo(w / 2, h); ctx.lineTo(0, h / 2); ctx.closePath(); }
  else if (k === 'rightArrow') { const t = h * 0.25, hx = w - Math.min(w * 0.4, h * 0.6); ctx.moveTo(0, t); ctx.lineTo(hx, t); ctx.lineTo(hx, 0); ctx.lineTo(w, h / 2); ctx.lineTo(hx, h); ctx.lineTo(hx, h - t); ctx.lineTo(0, h - t); ctx.closePath(); }
  else if (k === 'star5') { for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + (i * Math.PI) / 5, rr = i % 2 ? 0.4 : 1; const x = w / 2 + Math.cos(a) * (w / 2) * rr, y = h / 2 + Math.sin(a) * (h / 2) * rr; if (i) ctx.lineTo(x, y); else ctx.moveTo(x, y); } ctx.closePath(); }
  if (k === 'line') { ctx.moveTo(0, 0); ctx.lineTo(w, h); ctx.strokeStyle = ch.border || fill; ctx.lineWidth = 2; ctx.stroke(); ctx.restore(); return; }
  if (fill !== 'none') { ctx.fillStyle = fill; ctx.fill(); }
  if (line && line !== 'none') { ctx.strokeStyle = line; ctx.lineWidth = 1.25; ctx.stroke(); }
  if (ch.text) {
    const fs = (ch.fs ?? 11) * 4 / 3;
    ctx.fillStyle = ch.fg || (ch.type === 'text' ? '#000000' : '#ffffff');
    ctx.font = `${fs}px ${T.font}`; ctx.textBaseline = 'top';
    ctx.textAlign = ch.type === 'text' ? 'left' : 'center';
    const maxW = w - 12, lines: string[] = [];
    for (const para of ch.text.split('\n')) { let cur = ''; for (const word of para.split(/(\s+)/)) { const t = cur + word; if (ctx.measureText(t).width > maxW && cur.trim()) { lines.push(cur.trimEnd()); cur = word.trimStart(); } else cur = t; } lines.push(cur); }
    const lh = fs * 1.25, total = lines.length * lh;
    const y0 = ch.type === 'text' ? 6 : Math.max(4, (h - total) / 2);
    lines.forEach((ln, i) => { if (y0 + i * lh < h) ctx.fillText(ln, ch.type === 'text' ? 6 : w / 2, y0 + i * lh, maxW); });
  }
  ctx.restore();
}
export function drawChart(ctx: CanvasRenderingContext2D, W: number, H: number, ch: DrawSpec, d: ChartData, T: ChartTheme) {
  if (isObj(ch.type)) { ctx.save(); ctx.clearRect(0, 0, W, H); drawObj(ctx, W, H, ch, T); ctx.restore(); return; }
  ctx.save();
  const P = paletteOf(ch);
  ctx.fillStyle = ch.fill || T.bg;
  ctx.fillRect(0, 0, W, H);
  if (ch.border) { ctx.strokeStyle = ch.border; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, W - 1, H - 1); }
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
  const items = pie ? d.cats.map((c, i) => ({ name: c, color: P[i % P.length] })) : d.series.map((s, i) => ({ name: s.name, color: P[i % P.length] }));
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
  if (!d.series.length || (!d.cats.length && ch.type !== 'scatter')) {
    ctx.fillStyle = T.muted; ctx.textAlign = 'center';
    ctx.fillText('Chưa có dữ liệu số trong vùng đã chọn', W / 2, (top + bottom) / 2);
    ctx.restore();
    return;
  }
  if (pie) { drawPie(ctx, ch, d, T, left, top, right, bottom, P); ctx.restore(); return; }

  // tiêu đề trục
  ctx.font = `600 11px ${T.font}`; ctx.fillStyle = T.muted;
  if (ch.xTitle) { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(fit(ctx, ch.xTitle, W - 2 * pad), (left + right) / 2, bottom); bottom -= 16; }
  const rot = (t: string, x: number) => { ctx.save(); ctx.translate(x, (top + bottom) / 2); ctx.rotate(-Math.PI / 2); ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(fit(ctx, t, bottom - top), 0, 0); ctx.restore(); };
  if (ch.yTitle) { rot(ch.yTitle, left + 6); left += 16; }
  const kinds = d.series.map((_, i) => (ch.type === 'combo' ? (ch.series?.[i]?.kind ?? (i === 0 ? 'col' : 'line')) : null));
  const onY2 = d.series.map((_, i) => ch.type === 'combo' && !!ch.series?.[i]?.y2);
  const hasY2 = onY2.some(Boolean);
  if (hasY2 && ch.y2Title) { ctx.save(); ctx.translate(right - 6, (top + bottom) / 2); ctx.rotate(Math.PI / 2); ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(fit(ctx, ch.y2Title, bottom - top), 0, 0); ctx.restore(); right -= 16; }

  const scatter = ch.type === 'scatter';
  const horiz = ch.type === 'bar';
  const stacked = ch.type === 'colStack';
  const nCat = scatter ? 0 : d.cats.length;
  const range = (sel: (i: number) => boolean) => {
    let lo = 0, hi = 0;
    if (stacked) {
      for (let i = 0; i < nCat; i++) { let pos = 0, neg = 0; for (const s of d.series) { const v = s.values[i] ?? 0; if (v >= 0) pos += v; else neg += v; } hi = Math.max(hi, pos); lo = Math.min(lo, neg); }
    } else d.series.forEach((s, si) => { if (!sel(si)) return; for (const v of s.values) if (v !== null && v !== undefined) { hi = Math.max(hi, v); lo = Math.min(lo, v); } });
    return niceScale(lo, hi, horiz ? 4 : 5);
  };
  const sc = range((i) => !onY2[i]);
  const sc2 = hasY2 ? range((i) => onY2[i]) : null;
  // trục X số (phân tán)
  let xs: { lo: number; hi: number; step: number } | null = null;
  if (scatter) {
    let lo = Infinity, hi = -Infinity;
    d.series.forEach((s) => s.values.forEach((v, i) => { const x = s.xs ? s.xs[i] : i + 1; if (v !== null && x !== null && x !== undefined) { lo = Math.min(lo, x); hi = Math.max(hi, x); } }));
    if (!isFinite(lo)) { lo = 0; hi = 1; }
    xs = niceScale(lo >= 0 && hi > 0 && lo < hi * 0.25 ? 0 : lo, hi, 5);
    if (lo > 0 && xs.lo === 0 && lo > hi * 0.25) xs = { ...xs, lo: Math.floor(lo / xs.step) * xs.step };
  }
  ctx.font = `11px ${T.font}`;
  const ticksOf = (q: { lo: number; hi: number; step: number }) => { const t: number[] = []; for (let v = q.lo; v <= q.hi + q.step / 2; v += q.step) t.push(+v.toPrecision(12)); return t; };
  const ticks = ticksOf(sc), ticks2 = sc2 ? ticksOf(sc2) : [];
  const tickW = Math.max(...ticks.map((t) => ctx.measureText(compactNum(t)).width));
  const tickW2 = sc2 ? Math.max(...ticks2.map((t) => ctx.measureText(compactNum(t)).width)) : 0;
  const catW = horiz ? Math.min(W * 0.3, Math.max(...d.cats.map((c) => ctx.measureText(c).width)) + 8) : 0;
  const pl = left + (horiz ? catW : tickW + 8), pr = right - 4 - (sc2 ? tickW2 + 8 : 0), pt = top + 6;
  const pb = bottom - (horiz ? 18 : scatter ? 20 : 30);
  const pw = pr - pl, ph = pb - pt;
  if (pw < 20 || ph < 20) { ctx.restore(); return; }
  const valOf = (q: { lo: number; hi: number }) => (v: number) => (horiz ? pl + ((v - q.lo) / (q.hi - q.lo)) * pw : pb - ((v - q.lo) / (q.hi - q.lo)) * ph);
  const val2 = valOf(sc), valR = sc2 ? valOf(sc2) : val2;
  const xPos = (x: number) => pl + ((x - xs!.lo) / (xs!.hi - xs!.lo)) * pw;
  // lưới + nhãn trục giá trị
  ctx.strokeStyle = T.grid; ctx.lineWidth = 1; ctx.fillStyle = T.muted;
  for (const t of ticks) {
    const p = Math.round(val2(t)) + 0.5;
    if (ch.grid !== false) { ctx.beginPath(); if (horiz) { ctx.moveTo(p, pt); ctx.lineTo(p, pb); } else { ctx.moveTo(pl, p); ctx.lineTo(pr, p); } ctx.stroke(); }
    ctx.textAlign = horiz ? 'center' : 'right';
    ctx.textBaseline = horiz ? 'top' : 'middle';
    if (horiz) ctx.fillText(compactNum(t), p, pb + 4); else ctx.fillText(compactNum(t), pl - 6, p);
  }
  if (sc2) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; for (const t of ticks2) ctx.fillText(compactNum(t), pr + 6, Math.round(valR(t)) + 0.5); }
  if (scatter) {
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    for (const t of ticksOf(xs!)) { const p = Math.round(xPos(t)) + 0.5; if (ch.grid !== false) { ctx.strokeStyle = T.grid; ctx.beginPath(); ctx.moveTo(p, pt); ctx.lineTo(p, pb); ctx.stroke(); } ctx.fillStyle = T.muted; ctx.fillText(compactNum(t), p, pb + 4); }
  }
  // trục 0
  ctx.strokeStyle = T.muted;
  ctx.beginPath();
  const z = Math.round(val2(Math.max(sc.lo, Math.min(sc.hi, 0)))) + 0.5;
  if (horiz) { ctx.moveTo(z, pt); ctx.lineTo(z, pb); } else { ctx.moveTo(pl, z); ctx.lineTo(pr, z); }
  ctx.stroke();
  const nS = d.series.length;
  ctx.font = `10.5px ${T.font}`;
  if (scatter) {
    d.series.forEach((s, si) => {
      const color = P[si % P.length];
      s.values.forEach((v, i) => {
        const x = s.xs ? s.xs[i] : i + 1;
        if (v === null || x === null || x === undefined) return;
        const px = xPos(x), py = val2(v);
        ctx.fillStyle = color; ctx.beginPath(); ctx.arc(px, py, 3.8, 0, Math.PI * 2); ctx.fill();
        if (ch.labels) { ctx.fillStyle = T.fg; ctx.textAlign = 'left'; ctx.textBaseline = 'bottom'; ctx.fillText(compactNum(v), px + 5, py - 3); }
      });
    });
    ctx.restore();
    return;
  }
  // nhãn danh mục
  const band = (horiz ? ph : pw) / nCat;
  ctx.fillStyle = T.muted;
  ctx.font = `11px ${T.font}`;
  const every = Math.max(1, Math.ceil((horiz ? 14 : 48) / band));
  for (let i = 0; i < nCat; i += every) {
    const c = (i + 0.5) * band;
    if (horiz) { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(fit(ctx, d.cats[i], catW - 6), pl - 6, pt + c); }
    else { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(fit(ctx, d.cats[i], band * every - 4), pl + c, pb + 6); }
  }
  ctx.font = `10.5px ${T.font}`;
  const isLine = (si: number) => (ch.type === 'combo' ? kinds[si] === 'line' : ch.type === 'line' || ch.type === 'area');
  // cột/thanh
  const colIdx = d.series.map((_, si) => si).filter((si) => !isLine(si));
  if (colIdx.length) {
    const inner = band * 0.7;
    const bw = stacked ? inner : inner / colIdx.length;
    for (let i = 0; i < nCat; i++) {
      let pos = 0, neg = 0;
      colIdx.forEach((si, k) => {
        const s = d.series[si], v = s.values[i];
        if (v === null || v === undefined) return;
        const vf = onY2[si] ? valR : val2;
        const base = stacked ? (v >= 0 ? pos : neg) : 0;
        const a = vf(base), b = vf(base + v);
        if (stacked) { if (v >= 0) pos += v; else neg += v; }
        const off = (i + 0.15) * band + (stacked ? 0 : k * bw);
        ctx.fillStyle = P[si % P.length];
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
  // đường/vùng
  d.series.forEach((s, si) => {
    if (!isLine(si)) return;
    const color = P[si % P.length], vf = onY2[si] ? valR : val2;
    const pts = s.values.map((v, i) => (v === null ? null : { x: pl + (i + 0.5) * band, y: vf(v) }));
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
    if (ch.type !== 'area') for (const p of pts) if (p) { ctx.fillStyle = ch.fill || T.bg; ctx.beginPath(); ctx.arc(p.x, p.y, 3.5, 0, Math.PI * 2); ctx.fill(); ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.stroke(); }
    if (ch.labels) { ctx.fillStyle = T.fg; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; s.values.forEach((v, i) => { if (v !== null && pts[i]) ctx.fillText(compactNum(v), pts[i]!.x, pts[i]!.y - 5); }); }
  });
  void nS;
  ctx.restore();
}

function drawPie(ctx: CanvasRenderingContext2D, ch: Pick<Chart, 'type' | 'labels'>, d: ChartData, T: ChartTheme, l: number, t: number, r: number, b: number, P: string[]) {
  const vals = (d.series[0]?.values ?? []).map((v) => (v && v > 0 ? v : 0));
  const total = vals.reduce((a, v) => a + v, 0);
  const cx = (l + r) / 2, cy = (t + b) / 2, R = Math.max(10, Math.min(r - l, b - t) / 2 - 6);
  if (!total) { ctx.fillStyle = T.muted; ctx.textAlign = 'center'; ctx.fillText('Không có giá trị dương để vẽ', cx, cy); return; }
  let a = -Math.PI / 2;
  vals.forEach((v, i) => {
    const da = (v / total) * Math.PI * 2;
    ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, a, a + da); ctx.closePath();
    ctx.fillStyle = P[i % P.length]; ctx.fill();
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
