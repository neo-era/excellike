// Ngôn ngữ giao diện (Tệp ▸ Tùy chọn ▸ Ngôn ngữ): Tiếng Việt (mặc định) hoặc English như Excel 365 bản tiếng Anh.
// Mã viết chuỗi bằng tiếng Việt. Khi chọn English, một MutationObserver dịch chữ hiển thị và các thuộc tính
// title / aria-label / placeholder theo từ điển (khớp nguyên chuỗi, hoặc theo mẫu có {0}, {1}…).
// Nội dung ô, tên trang tính, tên tệp là dữ liệu của người dùng: không dịch (lưới vẽ bằng canvas nên không bị đụng).
import { EN as EN1, EN_PAT } from './i18n-en';
import { EN2, PAT2, STRICT, DATA, LABEL_Q } from './i18n-extra';
const EN: Record<string, string> = { ...EN1, ...EN2 };
const VI = /[À-ÖØ-öø-ỹĐđ]/;

export type Lang = 'vi' | 'en';
let cur: Lang = (() => { try { return (localStorage.getItem('xl-lang') as Lang) === 'en' ? 'en' : 'vi'; } catch { return 'vi'; } })();
export const lang = () => cur;
/** Đổi ngôn ngữ (app tải lại để dựng lại toàn bộ giao diện) */
export function setLang(l: Lang) { cur = l; try { localStorage.setItem('xl-lang', l); } catch { /* bỏ qua */ } }

const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
let pats: { re: RegExp; en: string; strict: boolean; data: Set<number> }[] | null = null;
/** chỉ số tham số là dữ liệu người dùng (không dịch) */
function dataArgs(vi: string) {
  const out = new Set<number>();
  for (const m of vi.matchAll(/\{(\d+)\}/g)) { const i = +m[1]; if (DATA.has(vi) || (!LABEL_Q.has(vi) && vi.includes(`“{${i}}”`))) out.add(i); }
  return out;
}
function patterns() {
  if (!pats) pats = [...EN_PAT, ...PAT2].map(([vi, en]) => ({ re: new RegExp('^' + esc(vi).replace(/\\\{(\d+)\\\}/g, '([\\s\\S]*?)') + '$'), en, strict: STRICT.has(vi), data: dataArgs(vi) }));
  return pats;
}
/** Dịch một chuỗi (giữ khoảng trắng đầu/cuối); không có trong từ điển thì giữ nguyên */
export function tr(s: string): string {
  if (cur !== 'en' || !s) return s;
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(s)!;
  const core = m[2];
  if (!core || !/[À-ỹĐđ]|[a-zA-Z]/.test(core)) return s;
  let out = EN[core];
  if (out === undefined) {
    for (const p of patterns()) {
      const x = p.re.exec(core);
      if (!x) continue;
      const args = x.slice(1).map((a) => a ?? ''), ta = args.map((a, i) => (p.data.has(i) ? a : tr(a)));
      // nhãn ghép: phần ghép phải dịch được, không thì để nguyên tiếng Việt (tránh “Remove khỏi …”)
      if (p.strict && args.some((a, i) => VI.test(a) && ta[i] === a)) continue;
      out = p.en.replace(/\{(\d+)\}/g, (_, i) => ta[+i] ?? '');
      break;
    }
  }
  // dòng ghép nhiều mục bằng “ · ” (cờ thanh trạng thái…): dịch từng mục
  if (out === undefined && core.includes(' · ')) { const parts = core.split(' · '), t = parts.map(tr); if (t.some((x, i) => x !== parts[i])) out = t.join(' · '); }
  return out === undefined ? s : m[1] + out + m[3];
}
/** Dịch chuỗi trong mã (khi cần trả về chữ trước khi gắn vào giao diện) */
export const T = (vi: string) => tr(vi);

const ATTRS = ['title', 'aria-label', 'placeholder', 'data-tip'];
/** vùng chứa dữ liệu người dùng: không dịch */
const SKIP = '.cell-editor, #fx-input, #namebox, .sheets .st, #doc-title, .fr-results td, .keys-row kbd, .clip-it, .fn-list b, .ac-list b, .bs-recent .rr .nm, .bs-recent .rr .path, .nv-name, [data-noi18n]';
function walk(n: Node) {
  if (n.nodeType === 3) {
    const p = (n as Text).parentElement;
    if (!p || p.closest(SKIP) || p.tagName === 'SCRIPT' || p.tagName === 'STYLE' || p.tagName === 'TEXTAREA') return;
    const v = n.nodeValue ?? '', t = tr(v);
    if (t !== v) n.nodeValue = t;
    return;
  }
  if (n.nodeType !== 1) return;
  const el = n as Element;
  if (el.closest?.(SKIP)) return;
  for (const a of ATTRS) { const v = el.getAttribute(a); if (v) { const t = tr(v); if (t !== v) el.setAttribute(a, t); } }
  if (el.tagName === 'INPUT' && (el as HTMLInputElement).type === 'button') { const i = el as HTMLInputElement; i.value = tr(i.value); }
  for (const c of Array.from(el.childNodes)) walk(c);
}
/** Bật dịch tự động cho toàn trang (gọi một lần khi khởi động nếu chọn English) */
export function startI18n() {
  if (cur !== 'en') return;
  document.documentElement.lang = 'en';
  walk(document.body);
  document.title = tr(document.title);
  new MutationObserver((ms) => {
    for (const m of ms) {
      if (m.type === 'characterData') walk(m.target);
      else if (m.type === 'attributes') { if (m.target.nodeType === 1 && ATTRS.includes(m.attributeName!)) { const el = m.target as Element, v = el.getAttribute(m.attributeName!) ?? ''; const t = tr(v); if (t !== v && !el.closest(SKIP)) el.setAttribute(m.attributeName!, t); } }
      else m.addedNodes.forEach(walk);
    }
  }).observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ATTRS });
}
