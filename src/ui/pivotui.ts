// PivotTable: hộp thoại tạo + ngăn trường (Bộ lọc / Cột / Hàng / Giá trị)
import { rangeToA1, parseA1, type Range } from '../core/address';
import { AGG_LABEL, labelOf, type Agg, type PivotSpec, type PivotTableData } from '../core/pivot';
import type { PivotDef, Workbook } from '../core/workbook';
import { parseSheetRef, refText } from './charts';
import { field, h, input, openDialog, select } from './dialogs';
import { ic } from './icons';

export function pivotCreateDialog(wb: Workbook, sheet: number, g: Range, cur: { r: number; c: number }, onDone: (res: { sheet: number; id: string } | null, err?: string) => void) {
  const body = h('div', { class: 'stack' });
  const src = input(refText(wb, { sheet, range: g }, sheet), 'ví dụ: A1:F200');
  const dest = select([['new', 'Sheet mới'], ['here', 'Sheet hiện tại']], 'new');
  const at = input(rangeToA1({ r1: cur.r, c1: cur.c, r2: cur.r, c2: cur.c }), 'ví dụ: H3');
  const atF = field('Đặt tại ô', at);
  atF.hidden = true;
  dest.addEventListener('change', () => { atF.hidden = (dest as HTMLSelectElement).value !== 'here'; });
  const err = h('p', { class: 'err-text', role: 'alert' });
  body.append(field('Vùng dữ liệu nguồn', src, 'Hàng đầu tiên phải là tên cột.'), field('Vị trí đặt PivotTable', dest), atF, err);
  openDialog('Tạo PivotTable', body, [
    { label: 'Hủy', act: () => onDone(null) },
    { label: 'Tạo', kind: 'primary', act: () => {
      err.textContent = '';
      const ref = parseSheetRef(wb, src.value, sheet);
      if (!ref || ref.range.r2 <= ref.range.r1) { err.textContent = 'Vùng nguồn không hợp lệ (cần ít nhất 1 hàng tiêu đề và 1 hàng dữ liệu).'; return false; }
      const u = wb.used(ref.sheet);
      if (ref.range.r1 >= u.rows || ref.range.c1 >= u.cols) { err.textContent = 'Vùng nguồn đang trống.'; return false; }
      let target: number | undefined, pos: { r: number; c: number } | undefined;
      if ((dest as HTMLSelectElement).value === 'here') {
        const a = parseA1(at.value.trim().replace(/\$/g, ''));
        if (!a) { err.textContent = 'Ô đặt PivotTable không hợp lệ.'; return false; }
        target = sheet; pos = { r: a.r, c: a.c };
      }
      try {
        const spec: PivotSpec = { rows: [], cols: [], values: [], filters: [] };
        const res = wb.addPivot({ ...spec, src: ref, ...(pos ? { at: pos } : {}) }, target);
        onDone(res);
      } catch (e) { err.textContent = (e as Error).message; return false; }
    } },
  ]);
}

type Zone = 'filters' | 'cols' | 'rows' | 'values';
const ZONES: [Zone, string][] = [['filters', 'Bộ lọc'], ['cols', 'Cột'], ['rows', 'Hàng'], ['values', 'Giá trị']];

export class PivotPanel {
  readonly el: HTMLElement;
  sheet = -1;
  id = '';
  closedFor = '';
  private openFilter = '';

  constructor(host: HTMLElement, private wb: Workbook, private onMsg: (m: string, kind?: 'err') => void, private onChanged: () => void) {
    this.el = h('aside', { class: 'pv-panel', 'aria-label': 'Trường PivotTable', hidden: '' });
    host.appendChild(this.el);
  }
  private get def(): PivotDef | null { return this.wb.sheetIds().includes(this.sheet) ? this.wb.meta(this.sheet).pivots.find((p) => p.id === this.id) ?? null : null; }

  /** Hiện khi ô hiện hành nằm trong một PivotTable */
  follow(sheet: number, r: number, c: number) {
    const p = this.wb.pivotAt(sheet, r, c);
    if (!p) { this.hide(); return; }
    if (this.closedFor === p.id) return;
    if (this.el.hidden || this.id !== p.id || this.sheet !== sheet) this.show(sheet, p.id!);
  }
  show(sheet: number, id: string) { this.sheet = sheet; this.id = id; this.closedFor = ''; this.el.hidden = false; this.render(); }
  hide() { this.el.hidden = true; }
  refreshView() { if (!this.el.hidden) { if (!this.def) this.hide(); else this.render(); } }

  private apply(patch: Partial<PivotSpec>) {
    try { const w = this.wb.refreshPivot(this.sheet, this.id, patch); this.onChanged(); if (w?.length) this.onMsg(w.join(' '), 'err'); }
    catch (e) { this.onMsg((e as Error).message, 'err'); }
    this.render();
  }

  private fields() {
    const p = this.def!;
    const t = this.wb.pivotTable(p.src);
    const numeric = t.header.map((_, i) => t.rows.some((r) => typeof r[i] === 'number'));
    return { names: t.header, numeric, table: t };
  }
  private used(p: PivotDef) {
    return new Set([...p.rows, ...p.cols, ...p.values.map((v) => v.field), ...p.filters.map((f) => f.field)]);
  }
  private list(p: PivotDef, z: Zone): string[] { return z === 'values' ? p.values.map((v) => v.field) : z === 'filters' ? p.filters.map((f) => f.field) : p[z]; }

  /** Đưa trường vào vùng (bỏ khỏi vùng cũ nếu là hàng/cột/lọc) */
  private moveTo(fieldName: string, from: Zone | null, fromIdx: number, to: Zone, toIdx = -1) {
    const p = this.def!;
    const spec: PivotSpec = { rows: [...p.rows], cols: [...p.cols], values: p.values.map((v) => ({ ...v })), filters: p.filters.map((f) => ({ ...f, values: [...f.values] })) };
    let val: { field: string; agg: Agg } | null = null;
    if (from) {
      if (from === 'values') val = spec.values.splice(fromIdx, 1)[0];
      else if (from === 'filters') spec.filters.splice(fromIdx, 1);
      else spec[from].splice(fromIdx, 1);
    }
    const ins = <T>(arr: T[], x: T) => { if (toIdx < 0 || toIdx > arr.length) arr.push(x); else arr.splice(toIdx, 0, x); };
    if (to === 'values') {
      const isNum = this.fields().numeric[this.fields().names.indexOf(fieldName)];
      ins(spec.values, val ?? { field: fieldName, agg: isNum ? 'sum' : 'count' });
    } else {
      // mỗi trường chỉ ở một trong hàng/cột/lọc
      spec.rows = spec.rows.filter((f) => f !== fieldName);
      spec.cols = spec.cols.filter((f) => f !== fieldName);
      spec.filters = spec.filters.filter((f) => f.field !== fieldName);
      if (to === 'filters') ins(spec.filters, { field: fieldName, values: [] });
      else ins(spec[to], fieldName);
    }
    this.apply(spec);
  }
  private removeAt(z: Zone, i: number) {
    const p = this.def!;
    if (z === 'values') this.apply({ values: p.values.filter((_, j) => j !== i) });
    else if (z === 'filters') this.apply({ filters: p.filters.filter((_, j) => j !== i) });
    else this.apply({ [z]: p[z].filter((_, j) => j !== i) } as Partial<PivotSpec>);
  }
  private toggleField(name: string, numeric: boolean) {
    const p = this.def!;
    if (this.used(p).has(name)) {
      this.apply({ rows: p.rows.filter((f) => f !== name), cols: p.cols.filter((f) => f !== name), values: p.values.filter((v) => v.field !== name), filters: p.filters.filter((f) => f.field !== name) });
    } else this.moveTo(name, null, -1, numeric ? 'values' : 'rows');
  }

  render() {
    const p = this.def;
    if (!p) { this.hide(); return; }
    const { names, numeric, table } = this.fields();
    const used = this.used(p);
    const el = this.el;
    el.replaceChildren();
    const head = h('header', {}, h('h2', {}, 'Trường PivotTable'));
    const x = h('button', { class: 'dlg-x', 'aria-label': 'Đóng ngăn PivotTable', title: 'Đóng' });
    x.innerHTML = ic('x', 18);
    x.onclick = () => { this.closedFor = this.id; this.hide(); };
    head.append(x);
    el.append(head);
    el.append(h('p', { class: 'muted pv-src' }, 'Nguồn: ' + refText(this.wb, p.src, -1)));
    // danh sách trường
    const fl = h('div', { class: 'pv-fields', role: 'group', 'aria-label': 'Chọn trường' });
    names.forEach((n, i) => {
      const cb = h('input', { type: 'checkbox' }) as HTMLInputElement;
      cb.checked = used.has(n);
      cb.onchange = () => this.toggleField(n, numeric[i]);
      const lab = h('label', { class: 'chk pv-f', draggable: 'true', title: 'Kéo vào một vùng bên dưới' }, cb, ' ' + n);
      if (numeric[i]) lab.append(h('small', {}, 'số'));
      lab.addEventListener('dragstart', (e) => e.dataTransfer!.setData('text/x-pv', JSON.stringify({ f: n, z: null, i: -1 })));
      fl.append(lab);
    });
    el.append(fl);
    // 4 vùng
    const zones = h('div', { class: 'pv-zones' });
    for (const [z, label] of ZONES) {
      const box = h('div', { class: 'pv-zone', 'data-z': z });
      box.append(h('h3', {}, label));
      const items = this.list(p, z);
      const ul = h('div', { class: 'pv-items' });
      items.forEach((f, i) => ul.append(this.chip(p, z, f, i, items.length, table)));
      if (!items.length) ul.append(h('p', { class: 'pv-empty' }, 'Kéo trường vào đây'));
      box.append(ul);
      box.addEventListener('dragover', (e) => { if (e.dataTransfer?.types.includes('text/x-pv')) { e.preventDefault(); box.classList.add('drop'); } });
      box.addEventListener('dragleave', () => box.classList.remove('drop'));
      box.addEventListener('drop', (e) => {
        e.preventDefault(); box.classList.remove('drop');
        const d = JSON.parse(e.dataTransfer!.getData('text/x-pv') || 'null');
        if (!d) return;
        const chips = [...ul.querySelectorAll<HTMLElement>('.pv-chip')];
        let idx = chips.findIndex((c) => e.clientY < c.getBoundingClientRect().top + c.offsetHeight / 2);
        if (idx < 0) idx = chips.length;
        if (d.z === z && d.i < idx) idx--;
        this.moveTo(d.f, d.z, d.i, z, idx);
      });
      zones.append(box);
    }
    el.append(zones);
    const foot = h('div', { class: 'pv-foot' });
    const rf = h('button', { class: 'btn' }); rf.innerHTML = ic('refresh', 16) + ' Làm mới';
    rf.onclick = () => { this.apply({}); this.onMsg('Đã làm mới PivotTable theo dữ liệu nguồn.'); };
    const del = h('button', { class: 'btn danger' }); del.innerHTML = ic('trash', 16) + ' Xóa PivotTable';
    del.onclick = () => { try { this.wb.removePivot(this.sheet, this.id); this.hide(); this.onChanged(); this.onMsg('Đã xóa PivotTable (Ctrl+Z để hoàn tác).'); } catch (e) { this.onMsg((e as Error).message, 'err'); } };
    foot.append(del, rf);
    el.append(foot);
  }

  private chip(p: PivotDef, z: Zone, f: string, i: number, n: number, table: PivotTableData) {
    const c = h('div', { class: 'pv-chip', draggable: 'true' });
    c.addEventListener('dragstart', (e) => e.dataTransfer!.setData('text/x-pv', JSON.stringify({ f, z, i })));
    const name = h('span', { class: 'pv-name' }, z === 'values' ? `${AGG_LABEL[p.values[i].agg]} ${f}` : f);
    c.append(name);
    if (z === 'values') {
      const agg = select((Object.keys(AGG_LABEL) as Agg[]).map((a) => [a, AGG_LABEL[a]] as [string, string]), p.values[i].agg) as HTMLSelectElement;
      agg.setAttribute('aria-label', 'Cách tính cho ' + f);
      agg.classList.add('pv-agg');
      agg.onchange = () => this.apply({ values: p.values.map((v, j) => (j === i ? { ...v, agg: agg.value as Agg } : v)) });
      c.append(agg);
    }
    // chuyển vùng (cho màn hình cảm ứng, không kéo-thả được)
    const mv = select([['', '⋯'], ...ZONES.filter(([k]) => k !== z).map(([k, l]) => [k, '→ ' + l] as [string, string]), ...(i > 0 ? [['up', '↑ Lên'] as [string, string]] : []), ...(i < n - 1 ? [['down', '↓ Xuống'] as [string, string]] : [])]) as HTMLSelectElement;
    mv.setAttribute('aria-label', 'Chuyển ' + f);
    mv.onchange = () => {
      const v = mv.value;
      if (v === 'up') this.moveTo(f, z, i, z, i - 1);
      else if (v === 'down') this.moveTo(f, z, i, z, i + 1);
      else if (v) this.moveTo(f, z, i, v as Zone);
    };
    c.append(mv);
    if (z === 'filters') {
      const flt = p.filters[i];
      const fb = h('button', { class: 'pv-btn', title: 'Chọn giá trị lọc', 'aria-expanded': String(this.openFilter === f) });
      fb.innerHTML = ic('filter', 14) + (flt.values.length ? ` ${flt.values.length}` : '');
      fb.onclick = () => { this.openFilter = this.openFilter === f ? '' : f; this.render(); };
      c.append(fb);
    }
    const rm = h('button', { class: 'pv-btn', 'aria-label': 'Bỏ ' + f, title: 'Bỏ trường' });
    rm.innerHTML = ic('x', 14);
    rm.onclick = () => this.removeAt(z, i);
    c.append(rm);
    if (z === 'filters' && this.openFilter === f) {
      const wrap = h('div', { class: 'pv-flist' });
      const col = table.header.indexOf(f);
      const vals = [...new Set(table.rows.map((_, r) => labelOf(table, r, col)))]
        .sort((a, b) => a.localeCompare(b, 'vi', { numeric: true })).slice(0, 500);
      const sel = new Set(p.filters[i].values.length ? p.filters[i].values : vals);
      const all = h('input', { type: 'checkbox' }) as HTMLInputElement;
      all.checked = sel.size === vals.length;
      wrap.append(h('label', { class: 'chk' }, all, ' (Tất cả)'));
      const boxes = vals.map((v) => { const b = h('input', { type: 'checkbox' }) as HTMLInputElement; b.checked = sel.has(v); b.value = v; wrap.append(h('label', { class: 'chk' }, b, ' ' + v)); return b; });
      all.onchange = () => boxes.forEach((b) => (b.checked = all.checked));
      const ok = h('button', { class: 'btn primary' }, 'Áp dụng lọc');
      ok.onclick = () => {
        const chosen = boxes.filter((b) => b.checked).map((b) => b.value);
        if (!chosen.length) { this.onMsg('Hãy chọn ít nhất một giá trị.', 'err'); return; }
        this.openFilter = '';
        this.apply({ filters: p.filters.map((x, j) => (j === i ? { ...x, values: chosen.length === vals.length ? [] : chosen } : x)) });
      };
      wrap.append(ok);
      const outer = h('div', { class: 'pv-chipwrap' }, c, wrap);
      return outer;
    }
    return c;
  }
}
