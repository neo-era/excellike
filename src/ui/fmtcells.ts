// Hộp thoại Định dạng Ô (Ctrl+1) như Excel 365: Số · Căn chỉnh · Phông chữ · Viền · Tô · Bảo vệ
import type { Range } from '../core/address';
import type { Border, BorderStyle, Style } from '../core/workbook';
import { formatValue } from '../core/numfmt';
import { openDialog, h } from './dialogs';
import { colorPalette, FONTS, DEFAULT_FONT, type App } from './app';
import { openPopover } from './menu';

export type FcTab = 'num' | 'align' | 'font' | 'border' | 'fill' | 'prot';
const TABS: [FcTab, string][] = [['num', 'Số'], ['align', 'Căn chỉnh'], ['font', 'Phông chữ'], ['border', 'Viền'], ['fill', 'Tô'], ['prot', 'Bảo vệ']];

// ---------- Số ----------
type Cat = 'gen' | 'num' | 'cur' | 'acc' | 'date' | 'time' | 'pct' | 'frac' | 'sci' | 'text' | 'spec' | 'custom';
const CATS: [Cat, string, string][] = [
  ['gen', 'Chung', 'Ô định dạng Chung không có định dạng số cụ thể.'],
  ['num', 'Số', 'Dùng cho số thông thường. Kế toán và Tiền tệ có định dạng riêng cho giá trị tiền.'],
  ['cur', 'Tiền tệ', 'Dùng cho giá trị tiền tệ thông thường.'],
  ['acc', 'Kế toán', 'Kế toán thẳng hàng ký hiệu tiền tệ và dấu thập phân trong cột.'],
  ['date', 'Ngày', 'Hiện số sê-ri ngày/giờ thành ngày.'],
  ['time', 'Giờ', 'Hiện số sê-ri ngày/giờ thành giờ.'],
  ['pct', 'Phần trăm', 'Nhân giá trị với 100 và hiện kèm ký hiệu %.'],
  ['frac', 'Phân số', ''],
  ['sci', 'Khoa học', ''],
  ['text', 'Văn bản', 'Ô định dạng Văn bản được coi là chữ kể cả khi gõ số.'],
  ['spec', 'Đặc biệt', 'Định dạng đặc biệt dùng cho mã, số điện thoại…'],
  ['custom', 'Tùy chỉnh', 'Gõ mã định dạng số, dùng một mã có sẵn làm điểm bắt đầu.'],
];
const SYMBOLS: [string, string][] = [['', 'Không'], ['"₫"', '₫'], ['$', '$'], ['"€"', '€'], ['"VND"', 'VND']];
const DATES = ['dd/mm/yyyy', 'd/m/yyyy', 'dd/mm/yy', 'd-mmm-yy', 'dd-mmm-yyyy', 'mmmm yyyy', 'mmm-yy', 'dddd, dd/mm/yyyy', 'dddd, dd mmmm yyyy', '"Ngày" dd "tháng" mm "năm" yyyy', 'dd/mm/yyyy h:mm', 'yyyy-mm-dd'];
const TIMES = ['h:mm', 'h:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'dd/mm/yyyy h:mm', 'mm:ss', '[h]:mm:ss', 'mm:ss.0'];
const FRACS: [string, string][] = [['# ?/?', 'Tối đa một chữ số (1/4)'], ['# ??/??', 'Tối đa hai chữ số (21/25)'], ['# ???/???', 'Tối đa ba chữ số (312/943)'], ['# ?/2', 'Phần hai (1/2)'], ['# ?/4', 'Phần tư (2/4)'], ['# ?/8', 'Phần tám (4/8)'], ['# ??/16', 'Phần mười sáu (8/16)'], ['# ?/10', 'Phần mười (3/10)'], ['# ??/100', 'Phần trăm (30/100)']];
const SPECIAL: [string, string][] = [['00000', 'Mã bưu chính (5 số)'], ['000000', 'Mã bưu chính Việt Nam (6 số)'], ['0### ### ###', 'Số điện thoại'], ['000-000-0000', 'Số điện thoại có gạch'], ['@', 'Mã số (giữ số 0 đầu)']];
const CUSTOM = ['General', '0', '0.00', '#,##0', '#,##0.00', '#,##0_);(#,##0)', '#,##0_);[Red](#,##0)', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)', '#,##0 "₫"', '#,##0.00 "₫"', '0%', '0.00%', '0.00E+00', '##0.0E+0', '# ?/?', '# ??/??', 'dd/mm/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'dd/mm/yyyy h:mm', 'mm:ss', '[h]:mm:ss', '@', '_-* #,##0 "₫"_-;-* #,##0 "₫"_-;_-* "-" "₫"_-;_-@_-'];

function numCode(dec: number, thou: boolean, neg: number, sym = '', pos: 'pre' | 'post' = 'post'): string {
  const body = (thou ? '#,##0' : '0') + (dec > 0 ? '.' + '0'.repeat(dec) : '');
  const s = sym ? (pos === 'pre' ? sym + body : body + ' ' + sym) : body;
  // kiểu số âm như Excel: -1234,10 · [Red]1234,10 · (1234,10) · [Red](1234,10)
  return [s + ';-' + s, s + ';[Red]' + s, s + '_);(' + s + ')', s + '_);[Red](' + s + ')'][neg] ?? s;
}
function accCode(dec: number, sym: string) {
  const d = dec > 0 ? '.' + '0'.repeat(dec) : '';
  const sy = sym ? ' ' + sym : '';
  return `_-* #,##0${d}${sy}_-;-* #,##0${d}${sy}_-;_-* "-"${'?'.repeat(dec)}${sy}_-;_-@_-`;
}
/** Đoán danh mục + tùy chọn từ mã định dạng hiện có */
export function catOf(f: string | undefined): Cat {
  if (!f || f === 'General') return 'gen';
  if (f === '@') return 'text';
  if (CUSTOM.includes(f) && /^\[h\]|^mm:ss/.test(f)) return 'time';
  if (/^_-\*/.test(f)) return 'acc';
  if (/%/.test(f)) return 'pct';
  if (/E\+/i.test(f)) return 'sci';
  if (/\?\/[?\d]/.test(f)) return 'frac';
  const plain = f.replace(/"[^"]*"/g, '');
  if (/[dy]/i.test(plain)) return 'date';
  if (/[hs]/i.test(plain) || /AM\/PM/.test(plain)) return 'time';
  if (/₫|\$|€|VND/.test(f)) return 'cur';
  if (/^[#0,.]+(;|_|$)/.test(f)) return 'num';
  return 'custom';
}

const STYLES: [BorderStyle | 'none', string][] = [['none', 'Không'], ['dotted', 'Chấm'], ['dashed', 'Gạch'], ['thin', 'Mảnh'], ['medium', 'Vừa'], ['thick', 'Dày'], ['double', 'Kép']];
type Edge = 't' | 'h' | 'b' | 'l' | 'v' | 'r' | 'd' | 'u';

export function formatCellsDialog(app: App, tab: FcTab = 'num') {
  const wb = app.wb, s = app.grid.sheet, areas = app.grid.ranges();
  const g: Range = app.grid.range, { ar, ac } = app.grid.sel;
  const cur: Style = { ...wb.style(s, ar, ac) };
  const val = wb.value(s, ar, ac);
  const patch: Style = {};
  const dirty = new Set<keyof Style>();
  const set = <K extends keyof Style>(k: K, v: Style[K]) => { (patch as any)[k] = v; dirty.add(k); };
  const eff = <K extends keyof Style>(k: K): Style[K] => (dirty.has(k) ? patch[k] : cur[k]);

  const bar = h('div', { class: 'ps-tabs', role: 'tablist' });
  const box = h('div', { class: 'ps fc' }, bar);
  const panes: Record<string, HTMLElement> = {};
  const show = (k: FcTab) => {
    bar.querySelectorAll('[role=tab]').forEach((x) => x.setAttribute('aria-selected', String((x as HTMLElement).dataset.t === k)));
    for (const [kk] of TABS) panes[kk].hidden = kk !== k;
  };
  for (const [k, l] of TABS) {
    const b = h('button', { role: 'tab', 'aria-selected': String(k === tab), 'data-t': k }, l);
    b.onclick = () => show(k);
    bar.append(b);
    panes[k] = h('div', { class: 'ps-pane fc-pane', role: 'tabpanel', 'data-p': k });
    panes[k].hidden = k !== tab;
    box.append(panes[k]);
  }
  // Ctrl+Tab / Ctrl+Shift+Tab chuyển thẻ
  box.addEventListener('keydown', (e) => {
    if (e.key === 'Tab' && e.ctrlKey) {
      e.preventDefault(); e.stopPropagation();
      const i = TABS.findIndex(([k]) => !panes[k].hidden);
      show(TABS[(i + (e.shiftKey ? TABS.length - 1 : 1)) % TABS.length][0]);
    }
  });

  // ===== Số =====
  {
    const P = panes.num;
    const list = h('div', { class: 'fc-cats', role: 'listbox', 'aria-label': 'Danh mục' });
    const sample = h('div', { class: 'fc-sample', id: 'fc-sample' });
    const opts = h('div', { class: 'fc-opts stack' });
    const desc = h('p', { class: 'muted fc-desc' });
    const code = h('input', { type: 'text', id: 'fc-code', spellcheck: 'false', 'aria-label': 'Kiểu' }) as HTMLInputElement;
    let cat: Cat = catOf(cur.fmt);
    let fmt = cur.fmt;
    const upd = (f: string | undefined) => {
      fmt = f;
      set('fmt', f === 'General' ? undefined : f);
      const v = typeof val === 'number' ? val : val === null || val === '' ? null : val;
      sample.textContent = 'Mẫu: ' + (v === null ? '' : formatValue(v, f && f !== 'General' ? f : undefined).text);
      if (cat === 'custom') code.value = f ?? 'General';
    };
    const numInp = (v: number, min: number, max: number) => { const i = h('input', { type: 'number', min: String(min), max: String(max), value: String(v), class: 'fc-n' }) as HTMLInputElement; return i; };
    const decOf = (f?: string) => { const m = /\.(0+)/.exec(f ?? ''); return m ? m[1].length : f && /[0#]/.test(f) ? 0 : 2; };
    const render = () => {
      list.querySelectorAll('button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.c === cat)));
      opts.innerHTML = '';
      desc.textContent = CATS.find((c) => c[0] === cat)?.[2] ?? '';
      const f0 = fmt;
      if (cat === 'gen') upd(undefined);
      else if (cat === 'text') upd('@');
      else if (cat === 'num' || cat === 'cur') {
        const dec = numInp(decOf(f0), 0, 30); dec.id = 'fc-dec';
        const thou = h('input', { type: 'checkbox', id: 'fc-thou' }) as HTMLInputElement; thou.checked = cat === 'cur' || /#,##/.test(f0 ?? '');
        const sym = h('select', { id: 'fc-sym' }) as HTMLSelectElement;
        for (const [v, l] of SYMBOLS) sym.append(h('option', { value: v }, l));
        sym.value = (/"₫"|₫/.test(f0 ?? '') ? '"₫"' : /\$/.test(f0 ?? '') ? '$' : /€/.test(f0 ?? '') ? '"€"' : /VND/.test(f0 ?? '') ? '"VND"' : cat === 'cur' ? '"₫"' : '');
        const neg = h('div', { class: 'fc-neg', role: 'listbox', 'aria-label': 'Số âm' });
        let ni = /\[Red\]\(/.test(f0 ?? '') ? 3 : /\(/.test(f0 ?? '') ? 2 : /\[Red\]/.test(f0 ?? '') ? 1 : 0;
        const code1 = () => numCode(Math.max(0, Math.min(30, +dec.value || 0)), cat === 'cur' ? true : thou.checked, ni, cat === 'cur' ? sym.value : '', sym.value === '$' ? 'pre' : 'post');
        const negList = () => {
          neg.innerHTML = '';
          [0, 1, 2, 3].forEach((k) => {
            const f = numCode(Math.max(0, +dec.value || 0), cat === 'cur' || thou.checked, k, cat === 'cur' ? sym.value : '', sym.value === '$' ? 'pre' : 'post');
            const t = formatValue(-1234.1, f).text;
            const b = h('button', { 'aria-selected': String(k === ni), 'data-k': String(k), class: k % 2 ? 'red' : '' }, t);
            b.onclick = () => { ni = k; negList(); upd(code1()); };
            neg.append(b);
          });
        };
        const go = () => { negList(); upd(code1()); };
        dec.oninput = go; thou.onchange = go; sym.onchange = go;
        opts.append(h('label', { class: 'fld row' }, h('span', {}, 'Số chữ số thập phân:'), dec));
        if (cat === 'num') opts.append(h('label', { class: 'chk' }, thou, ' Dùng dấu phân cách hàng nghìn (.)'));
        else opts.append(h('label', { class: 'fld row' }, h('span', {}, 'Ký hiệu:'), sym));
        opts.append(h('span', { class: 'muted' }, 'Số âm:'), neg);
        go();
      } else if (cat === 'acc') {
        const dec = numInp(/\.(0+)/.exec(f0 ?? '')?.[1].length ?? 0, 0, 30); dec.id = 'fc-dec';
        const sym = h('select', { id: 'fc-sym' }) as HTMLSelectElement;
        for (const [v, l] of SYMBOLS) sym.append(h('option', { value: v }, l));
        sym.value = /\$/.test(f0 ?? '') ? '$' : /€/.test(f0 ?? '') ? '"€"' : /^_-\*/.test(f0 ?? '') && !/₫/.test(f0 ?? '') ? '' : '"₫"';
        const go = () => upd(accCode(Math.max(0, +dec.value || 0), sym.value));
        dec.oninput = go; sym.onchange = go;
        opts.append(h('label', { class: 'fld row' }, h('span', {}, 'Số chữ số thập phân:'), dec), h('label', { class: 'fld row' }, h('span', {}, 'Ký hiệu:'), sym));
        go();
      } else if (cat === 'date' || cat === 'time' || cat === 'frac' || cat === 'spec') {
        const items: [string, string][] = cat === 'date' ? DATES.map((d) => [d, formatValue(typeof val === 'number' ? val : 46324.5, d).text])
          : cat === 'time' ? TIMES.map((d) => [d, formatValue(typeof val === 'number' ? val : 46324.5625, d).text]) : cat === 'frac' ? FRACS : SPECIAL;
        const lb = h('div', { class: 'fc-list', role: 'listbox', 'aria-label': 'Kiểu', id: 'fc-types' });
        const pick = (f: string) => { lb.querySelectorAll('button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.f === f))); upd(f); };
        for (const [f, l] of items) { const b = h('button', { 'data-f': f, 'aria-selected': 'false' }, l); b.onclick = () => pick(f); lb.append(b); }
        opts.append(h('span', { class: 'muted' }, 'Kiểu:'), lb);
        pick(items.find(([f]) => f === f0)?.[0] ?? items[0][0]);
      } else if (cat === 'pct' || cat === 'sci') {
        const dec = numInp(decOf(f0), 0, 30); dec.id = 'fc-dec';
        const go = () => { const d = Math.max(0, +dec.value || 0); upd((cat === 'pct' ? '0' : '0') + (d ? '.' + '0'.repeat(d) : '') + (cat === 'pct' ? '%' : 'E+00')); };
        dec.oninput = go;
        opts.append(h('label', { class: 'fld row' }, h('span', {}, 'Số chữ số thập phân:'), dec));
        go();
      } else if (cat === 'custom') {
        const lb = h('div', { class: 'fc-list', role: 'listbox', 'aria-label': 'Mã có sẵn' });
        const codes = [...new Set([...CUSTOM, ...app.usedFormats()])];
        for (const f of codes) { const b = h('button', { 'data-f': f }, f); b.onclick = () => { code.value = f; upd(f); lb.querySelectorAll('button').forEach((x) => x.setAttribute('aria-selected', String(x === b))); }; lb.append(b); }
        code.value = f0 ?? 'General';
        code.oninput = () => upd(code.value.trim() || 'General');
        opts.append(h('label', { class: 'fld' }, h('span', {}, 'Kiểu:'), code), lb);
        upd(code.value);
      }
    };
    for (const [c, l] of CATS) {
      const b = h('button', { 'data-c': c, role: 'option' }, l);
      b.onclick = () => { cat = c; render(); };
      list.append(b);
    }
    P.append(h('div', { class: 'fc-num' }, h('div', {}, h('span', { class: 'muted' }, 'Danh mục:'), list), h('div', { class: 'stack' }, sample, opts, desc)));
    render();
    // mở hộp thoại không được coi là đổi định dạng
    dirty.delete('fmt'); delete patch.fmt;
  }

  // ===== Căn chỉnh =====
  {
    const P = panes.align;
    const hz = h('select', { id: 'fc-h' }) as HTMLSelectElement;
    for (const [v, l] of [['', 'Chung'], ['l', 'Trái (Thụt lề)'], ['c', 'Giữa'], ['r', 'Phải (Thụt lề)'], ['fill', 'Điền'], ['just', 'Căn đều'], ['cas', 'Giữa qua Vùng chọn'], ['dist', 'Phân phối (Thụt lề)']]) hz.append(h('option', { value: v }, l));
    hz.value = cur.hx ?? cur.al ?? '';
    const vt = h('select', { id: 'fc-v' }) as HTMLSelectElement;
    for (const [v, l] of [['t', 'Trên'], ['m', 'Giữa'], ['b', 'Dưới']]) vt.append(h('option', { value: v }, l));
    vt.value = cur.va ?? 'b';
    const ind = h('input', { type: 'number', min: '0', max: '250', value: String(cur.ind ?? 0), id: 'fc-ind', class: 'fc-n' }) as HTMLInputElement;
    const wrap = h('input', { type: 'checkbox', id: 'fc-wrap' }) as HTMLInputElement; wrap.checked = !!cur.wr;
    const shk = h('input', { type: 'checkbox', id: 'fc-shk' }) as HTMLInputElement; shk.checked = !!cur.shk;
    const merged = !!wb.mergeAt(s, ar, ac);
    const mrg = h('input', { type: 'checkbox', id: 'fc-merge' }) as HTMLInputElement; mrg.checked = merged;
    const deg = h('input', { type: 'number', min: '-90', max: '90', value: String(cur.rot && cur.rot !== 255 ? cur.rot : 0), id: 'fc-deg', class: 'fc-n' }) as HTMLInputElement;
    const vert = h('input', { type: 'checkbox', id: 'fc-vert' }) as HTMLInputElement; vert.checked = cur.rot === 255;
    const dial = h('div', { class: 'fc-dial', 'aria-hidden': 'true' }, h('span', { class: 'fc-needle' }, 'Văn bản'));
    const needle = () => { (dial.firstChild as HTMLElement).style.transform = `rotate(${-(+deg.value || 0)}deg)`; dial.classList.toggle('off', vert.checked); };
    hz.onchange = () => {
      const v = hz.value;
      if (v === 'fill' || v === 'just' || v === 'cas' || v === 'dist') { set('hx', v); set('al', v === 'cas' || v === 'dist' ? 'c' : undefined); }
      else { set('hx', undefined); set('al', (v || undefined) as Style['al']); }
      ind.disabled = !(v === 'l' || v === 'r' || v === 'dist');
    };
    ind.disabled = !['l', 'r', 'dist'].includes(hz.value);
    vt.onchange = () => set('va', vt.value === 'b' ? undefined : (vt.value as Style['va']));
    ind.oninput = () => { const n = Math.max(0, Math.min(250, Math.round(+ind.value || 0))); set('ind', n || undefined); if (n && !hz.value) { hz.value = 'l'; hz.onchange!(new Event('change')); } };
    wrap.onchange = () => set('wr', wrap.checked || undefined);
    shk.onchange = () => set('shk', shk.checked || undefined);
    deg.oninput = () => { const n = Math.max(-90, Math.min(90, Math.round(+deg.value || 0))); vert.checked = false; set('rot', n || undefined); needle(); };
    vert.onchange = () => { set('rot', vert.checked ? 255 : (+deg.value || undefined)); needle(); };
    needle();
    P.append(
      h('h4', {}, 'Căn chỉnh văn bản'),
      h('div', { class: 'fc-grid2' }, h('label', { class: 'fld' }, h('span', {}, 'Ngang:'), hz), h('label', { class: 'fld' }, h('span', {}, 'Thụt lề:'), ind), h('label', { class: 'fld' }, h('span', {}, 'Dọc:'), vt)),
      h('h4', {}, 'Điều khiển văn bản'),
      h('label', { class: 'chk' }, wrap, ' Ngắt dòng'), h('label', { class: 'chk' }, shk, ' Co chữ cho vừa'), h('label', { class: 'chk' }, mrg, ' Gộp ô'),
      h('h4', {}, 'Hướng'),
      h('div', { class: 'fc-orient' }, h('label', { class: 'chk' }, vert, ' Chữ dọc'), dial, h('label', { class: 'fld row' }, h('span', {}, 'Độ:'), deg)),
    );
    (P as any).__merge = () => mrg.checked !== merged ? mrg.checked : null;
  }

  // ===== Phông chữ =====
  {
    const P = panes.font;
    const fam = h('input', { type: 'text', id: 'fc-font', list: 'fc-fonts', value: cur.fn ?? DEFAULT_FONT }) as HTMLInputElement;
    const dl = h('datalist', { id: 'fc-fonts' }); for (const f of FONTS) dl.append(h('option', { value: f }));
    const sty = h('select', { id: 'fc-style' }) as HTMLSelectElement;
    for (const [v, l] of [['', 'Thường'], ['i', 'Nghiêng'], ['b', 'Đậm'], ['bi', 'Đậm Nghiêng']]) sty.append(h('option', { value: v }, l));
    sty.value = (cur.b ? 'b' : '') + (cur.i ? 'i' : '');
    const size = h('input', { type: 'number', min: '1', max: '409', step: '0.5', id: 'fc-size', value: String(cur.fs ?? 11), class: 'fc-n' }) as HTMLInputElement;
    const und = h('select', { id: 'fc-under' }) as HTMLSelectElement;
    for (const [v, l] of [['', 'Không'], ['1', 'Đơn'], ['2', 'Kép'], ['a1', 'Kế toán Đơn'], ['a2', 'Kế toán Kép']]) und.append(h('option', { value: v }, l));
    und.value = cur.u ? (cur.ud ? '2' : '1') : '';
    const strike = h('input', { type: 'checkbox', id: 'fc-strike' }) as HTMLInputElement; strike.checked = !!cur.s;
    const sup = h('input', { type: 'checkbox', id: 'fc-sup' }) as HTMLInputElement; sup.checked = cur.vp === 'sup';
    const sub = h('input', { type: 'checkbox', id: 'fc-sub' }) as HTMLInputElement; sub.checked = cur.vp === 'sub';
    const colorBtn = h('button', { class: 'btn fc-color', id: 'fc-fg', type: 'button' }, 'Màu chữ');
    const prev = h('div', { class: 'fc-preview', id: 'fc-fprev' }, 'AaBbCc ÀẢÃÁẠ 123');
    const paint = () => {
      const fs = +size.value || 11;
      prev.style.fontFamily = `'${fam.value}', Calibri, sans-serif`;
      prev.style.fontSize = Math.min(28, fs * 4 / 3) + 'px';
      prev.style.fontWeight = sty.value.includes('b') ? '700' : '400';
      prev.style.fontStyle = sty.value.includes('i') ? 'italic' : 'normal';
      prev.style.textDecoration = [und.value ? 'underline' : '', strike.checked ? 'line-through' : ''].filter(Boolean).join(' ') + (und.value.endsWith('2') ? ' double' : '');
      prev.style.color = eff('fg') ?? '';
      prev.style.verticalAlign = sup.checked ? 'super' : sub.checked ? 'sub' : '';
      colorBtn.style.setProperty('--bar', eff('fg') ?? 'currentColor');
    };
    fam.onchange = () => { const v = fam.value.trim(); set('fn', !v || v === DEFAULT_FONT ? undefined : v); paint(); };
    sty.onchange = () => { set('b', sty.value.includes('b') || undefined); set('i', sty.value.includes('i') || undefined); paint(); };
    size.oninput = () => { const n = parseFloat(size.value); if (n >= 1 && n <= 409) set('fs', n === 11 ? undefined : n); paint(); };
    und.onchange = () => { set('u', !!und.value || undefined); set('ud', und.value.endsWith('2') || undefined); paint(); };
    strike.onchange = () => { set('s', strike.checked || undefined); paint(); };
    sup.onchange = () => { if (sup.checked) sub.checked = false; set('vp', sup.checked ? 'sup' : sub.checked ? 'sub' : undefined); paint(); };
    sub.onchange = () => { if (sub.checked) sup.checked = false; set('vp', sub.checked ? 'sub' : sup.checked ? 'sup' : undefined); paint(); };
    colorBtn.onclick = () => { const r = colorBtn.getBoundingClientRect(); openPopover(colorPalette((c) => { set('fg', c); paint(); }, 'Tự động'), r.left, r.bottom + 2); };
    P.append(
      h('div', { class: 'fc-grid3' }, h('label', { class: 'fld' }, h('span', {}, 'Phông:'), fam, dl), h('label', { class: 'fld' }, h('span', {}, 'Kiểu phông:'), sty), h('label', { class: 'fld' }, h('span', {}, 'Cỡ:'), size)),
      h('div', { class: 'fc-grid2' }, h('label', { class: 'fld' }, h('span', {}, 'Gạch chân:'), und), h('div', { class: 'fld' }, h('span', {}, 'Màu:'), colorBtn)),
      h('h4', {}, 'Hiệu ứng'),
      h('label', { class: 'chk' }, strike, ' Gạch ngang'), h('label', { class: 'chk' }, sup, ' Chỉ số trên'), h('label', { class: 'chk' }, sub, ' Chỉ số dưới'),
      h('h4', {}, 'Xem trước'), prev,
    );
    paint();
  }

  // ===== Viền =====
  const bspec: Partial<Record<Edge, Border | null>> = {};
  {
    const P = panes.border;
    let line: BorderStyle | 'none' = 'thin';
    let color: string | undefined;
    const multiR = g.r2 > g.r1, multiC = g.c2 > g.c1;
    // trạng thái hiện tại (theo ô góc) để vẽ xem trước
    const at = (r: number, c: number) => wb.style(s, r, c).bd ?? {};
    const state: Record<Edge, Border | null> = {
      t: at(g.r1, g.c1).t ?? null, b: at(Math.min(g.r2, g.r1 + 50), g.c1).b ?? null, l: at(g.r1, g.c1).l ?? null, r: at(g.r1, Math.min(g.c2, g.c1 + 50)).r ?? null,
      h: multiR ? at(g.r1, g.c1).b ?? null : null, v: multiC ? at(g.r1, g.c1).r ?? null : null, d: at(g.r1, g.c1).d ?? null, u: at(g.r1, g.c1).u ?? null,
    };
    const styles = h('div', { class: 'fc-lines', role: 'listbox', 'aria-label': 'Kiểu nét', id: 'fc-lines' });
    for (const [v, l] of STYLES) {
      const b = h('button', { 'data-s': v, 'aria-selected': String(v === line), title: l }, h('i', { class: 'ln ln-' + v }), h('span', {}, l));
      b.onclick = () => { line = v; styles.querySelectorAll('button').forEach((x) => x.setAttribute('aria-selected', String(x === b))); };
      styles.append(b);
    }
    const colBtn = h('button', { class: 'btn fc-color', id: 'fc-bcolor', type: 'button' }, 'Màu nét');
    colBtn.onclick = () => { const r = colBtn.getBoundingClientRect(); openPopover(colorPalette((c) => { color = c; colBtn.style.setProperty('--bar', c ?? 'currentColor'); }, 'Tự động'), r.left, r.bottom + 2); };
    const cur1 = (): Border | null => (line === 'none' ? null : color ? { s: line, c: color } : { s: line });
    const pv = h('div', { class: 'fc-bprev', id: 'fc-bprev' });
    const draw = () => {
      pv.innerHTML = '';
      const cls = (e: Edge) => 'e e-' + e + (state[e] ? ' on ln-' + state[e]!.s : '');
      for (const e of ['t', 'b', 'l', 'r', 'h', 'v', 'd', 'u'] as Edge[]) {
        if ((e === 'h' && !multiR) || (e === 'v' && !multiC)) continue;
        const d = h('div', { class: cls(e) });
        if (state[e]?.c) d.style.setProperty('--bc', state[e]!.c!);
        pv.append(d);
      }
      pv.append(h('span', { class: 'fc-txt' }, 'Văn bản'));
    };
    const toggle = (e: Edge) => { const want = state[e] && cur1() && state[e]!.s === cur1()!.s && state[e]!.c === cur1()!.c ? null : cur1(); state[e] = want; bspec[e] = want; draw(); };
    const edgeBtn = (e: Edge, label: string) => { const b = h('button', { class: 'btn ic fc-edge', 'data-e': e, title: label, 'aria-label': label, type: 'button' }, label); b.onclick = () => toggle(e); if ((e === 'h' && !multiR) || (e === 'v' && !multiC)) b.disabled = true; return b; };
    const preset = (k: 'none' | 'outline' | 'inside', label: string) => {
      const b = h('button', { class: 'btn fc-preset', 'data-p': k, type: 'button' }, label);
      b.onclick = () => {
        const c = cur1();
        if (k === 'none') for (const e of ['t', 'b', 'l', 'r', 'h', 'v', 'd', 'u'] as Edge[]) { state[e] = null; bspec[e] = null; }
        else for (const e of (k === 'outline' ? ['t', 'b', 'l', 'r'] : ['h', 'v']) as Edge[]) { if ((e === 'h' && !multiR) || (e === 'v' && !multiC)) continue; state[e] = c; bspec[e] = c; }
        draw();
      };
      if (k === 'inside' && !multiR && !multiC) b.disabled = true;
      return b;
    };
    P.append(
      h('div', { class: 'fc-border' },
        h('div', { class: 'stack' }, h('span', { class: 'muted' }, 'Nét — Kiểu:'), styles, colBtn),
        h('div', { class: 'stack' },
          h('span', { class: 'muted' }, 'Đặt sẵn:'), h('div', { class: 'row' }, preset('none', 'Không'), preset('outline', 'Viền ngoài'), preset('inside', 'Bên trong')),
          h('span', { class: 'muted' }, 'Viền:'),
          h('div', { class: 'fc-bwrap' },
            h('div', { class: 'fc-bcol' }, edgeBtn('t', 'Trên'), edgeBtn('h', 'Giữa ngang'), edgeBtn('b', 'Dưới'), edgeBtn('u', 'Chéo lên ↗')),
            pv,
            h('div', { class: 'fc-brow' }, edgeBtn('l', 'Trái'), edgeBtn('v', 'Giữa dọc'), edgeBtn('r', 'Phải'), edgeBtn('d', 'Chéo xuống ↘'))),
          h('p', { class: 'muted' }, 'Chọn kiểu nét và màu, rồi bấm nút đặt sẵn, sơ đồ hoặc các nút viền để áp.'))),
    );
    draw();
  }

  // ===== Tô =====
  {
    const P = panes.fill;
    const sample = h('div', { class: 'fc-fillprev', id: 'fc-fillprev' });
    const paint = () => { sample.style.background = eff('bg') ?? 'var(--cell, #fff)'; };
    const pal = colorPalette((c) => { set('bg', c); paint(); }, 'Không Màu');
    P.append(h('span', { class: 'muted' }, 'Màu Nền:'), pal, h('span', { class: 'muted' }, 'Mẫu:'), sample);
    paint();
  }

  // ===== Bảo vệ =====
  {
    const P = panes.prot;
    const lk = h('input', { type: 'checkbox', id: 'fc-lock' }) as HTMLInputElement; lk.checked = cur.lk !== false;
    const hd = h('input', { type: 'checkbox', id: 'fc-hidden' }) as HTMLInputElement; hd.checked = !!cur.hid;
    lk.onchange = () => set('lk', lk.checked ? undefined : false);
    hd.onchange = () => set('hid', hd.checked || undefined);
    P.append(h('label', { class: 'chk' }, lk, ' Khóa'), h('label', { class: 'chk' }, hd, ' Ẩn'),
      h('p', { class: 'muted' }, 'Khóa ô hoặc ẩn công thức chỉ có tác dụng sau khi bảo vệ trang tính (thẻ Xem lại ▸ Bảo vệ Trang tính, có ở bản 0.9). Thuộc tính được lưu vào tệp .xlsx.'));
  }

  openDialog('Định dạng Ô', box, [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      const mrg = (panes.align as any).__merge?.();
      app.applyFormatCells(areas, dirty.size ? Object.fromEntries([...dirty].map((k) => [k, patch[k]])) as Style : null, Object.keys(bspec).length ? bspec : null, mrg);
    } },
  ], () => app.grid.focusKeys()).querySelector('.dlg')?.classList.add('wide', 'fc-dlg');
}
