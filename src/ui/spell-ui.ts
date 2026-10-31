// Xem lại ▸ Chính tả (F7): duyệt các ô chữ trên trang tính, gợi ý sửa, bỏ qua, thêm vào từ điển riêng
import { openDialog, closeDialog, h } from './dialogs';
import { toA1 } from '../core/address';
import { Speller, words } from '../core/spell';
import type { App } from './app';

let speller: Speller | null = null;
async function load(): Promise<Speller> {
  if (speller) return speller;
  const [va, vd, ea, ed] = await Promise.all([import('../dict/vi.aff?raw'), import('../dict/vi.dic?raw'), import('../dict/en.aff?raw'), import('../dict/en.dic?raw')]);
  speller = new Speller({ aff: va.default, dic: vd.default }, { aff: ea.default, dic: ed.default });
  try { speller.addWords(JSON.parse(localStorage.getItem('xl-dict') ?? '[]')); } catch { /* bỏ qua */ }
  return speller;
}
const saveCustom = (w: string) => { try { const l: string[] = JSON.parse(localStorage.getItem('xl-dict') ?? '[]'); if (!l.includes(w)) { l.push(w); localStorage.setItem('xl-dict', JSON.stringify(l)); } } catch { /* bỏ qua */ } };

export async function spellCheck(app: App) {
  app.toast('Đang nạp từ điển…');
  const sp = await load();
  const s = app.grid.sheet, wb = app.wb, u = wb.used(s);
  if (!u.rows) { app.toast('Trang tính trống.'); return; }
  const ignore = new Set<string>(), changeAll = new Map<string, string>();
  const start = { r: app.grid.sel.ar, c: app.grid.sel.ac };
  // các ô theo thứ tự hàng → cột, bắt đầu từ ô hiện hành rồi vòng lại đầu trang
  // (tính vị trí theo chỉ số, không dựng mảng mọi ô — trang lớn không tốn bộ nhớ)
  const N = u.rows * u.cols;
  const cells = { at: (i: number) => ({ r: Math.floor(i / u.cols), c: i % u.cols }) };
  let k = start.r >= u.rows ? 0 : start.r * u.cols + Math.min(start.c, u.cols);
  let seen = 0, offset = 0, changes = 0;
  type Hit = { r: number; c: number; w: string; i: number; text: string };
  const next = (): Hit | null => {
    while (seen < N) {
      const p = cells.at(k % N);
      const raw = wb.rawText(s, p.r, p.c);
      if (raw && !raw.startsWith('=') && typeof wb.value(s, p.r, p.c) === 'string') {
        let text = raw.replace(/^'/, '');
        // Thay đổi Tất cả đã chọn trước đó
        if (changeAll.size) {
          let t2 = text;
          for (const x of words(text).reverse()) { const rep = changeAll.get(x.w); if (rep !== undefined) t2 = t2.slice(0, x.i) + rep + t2.slice(x.i + x.w.length); }
          if (t2 !== text) { wb.setInput(s, p.r, p.c, (raw.startsWith("'") ? "'" : '') + t2); text = t2; changes++; }
        }
        for (const x of words(text)) {
          if (x.i < offset) continue;
          if (ignore.has(x.w) || sp.correct(x.w)) continue;
          return { r: p.r, c: p.c, w: x.w, i: x.i, text };
        }
      }
      k++; seen++; offset = 0;
    }
    return null;
  };
  const word = h('div', { class: 'sp-ctx', 'aria-live': 'polite' });
  const sug = h('select', { size: '6', class: 'nm-list', 'aria-label': 'Gợi ý' }) as HTMLSelectElement;
  const where = h('div', { class: 'muted' });
  let cur: Hit | null = null;
  const show = () => {
    cur = next();
    if (!cur) {
      closeDialog();
      app.grid.invalidate(); app.syncSelection(); app.scheduleSave();
      app.toast(`Đã hoàn tất kiểm tra chính tả cho toàn bộ trang tính${changes ? ` (đã sửa ${changes} ô)` : ''}.`, 'ok');
      return;
    }
    app.grid.select(cur.r, cur.c); app.syncSelection();
    where.textContent = `Ô ${toA1(cur.r, cur.c)} — tiếng Việt / English`;
    word.replaceChildren(cur.text.slice(0, cur.i), h('mark', {}, cur.w), cur.text.slice(cur.i + cur.w.length));
    const list = sp.suggest(cur.w);
    sug.replaceChildren(...(list.length ? list.map((x) => h('option', { value: x }, x)) : [h('option', { value: '', disabled: '' }, '(Không có gợi ý)')]));
    sug.selectedIndex = list.length ? 0 : -1;
  };
  const advance = () => { if (cur) offset = cur.i + cur.w.length; show(); };
  const replace = (all: boolean) => {
    if (!cur) return;
    const rep = sug.value;
    if (!rep) return;
    const raw = wb.rawText(s, cur.r, cur.c);
    const t2 = cur.text.slice(0, cur.i) + rep + cur.text.slice(cur.i + cur.w.length);
    wb.setInput(s, cur.r, cur.c, (raw.startsWith("'") ? "'" : '') + t2); changes++;
    if (all) changeAll.set(cur.w, rep);
    offset = cur.i + rep.length;
    cur = null; show();
  };
  const btn = (t: string, f: () => void) => { const b = h('button', { class: 'btn' }, t); b.onclick = f; return b; };
  const body = h('div', { class: 'sp-dlg' },
    h('div', {}, h('b', {}, 'Không có trong Từ điển:'), word, where, h('b', {}, 'Gợi ý:'), sug),
    h('div', { class: 'an-side' },
      btn('Bỏ qua Một lần', advance),
      btn('Bỏ qua Tất cả', () => { if (cur) ignore.add(cur.w); advance(); }),
      btn('Thêm vào Từ điển', () => { if (cur) { sp.addWords([cur.w]); saveCustom(cur.w); } advance(); }),
      btn('Thay đổi', () => replace(false)),
      btn('Thay đổi Tất cả', () => replace(true))));
  sug.ondblclick = () => replace(false);
  openDialog('Chính tả: Tiếng Việt, English', body, [{ label: 'Đóng', act: () => { app.grid.invalidate(); app.syncSelection(); app.scheduleSave(); } }]);
  show();
}
