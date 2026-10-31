// Kiểm tra chính tả (F7): từ điển Hunspell tiếng Việt (hunspell-vi, GPLv3) + tiếng Anh (MIT/BSD), thư viện nspell.
// Tiếng Việt viết theo âm tiết nên kiểm từng âm tiết; một từ đúng nếu có trong từ điển Việt HOẶC Anh.
// @ts-expect-error: nspell không kèm khai báo kiểu
import NSpell from 'nspell';

interface NS { correct(w: string): boolean; suggest(w: string): string[]; add(w: string): void }
const VI_MARK = /[À-ỹĐđ]/;
export interface SpellOpts { ignoreUpper?: boolean; ignoreDigits?: boolean }
export class Speller {
  private vi: NS; private en: NS;
  private custom = new Set<string>();
  constructor(vi: { aff: string; dic: string }, en: { aff: string; dic: string }) {
    this.vi = NSpell(vi.aff, vi.dic); this.en = NSpell(en.aff, en.dic);
  }
  addWords(ws: Iterable<string>) { for (const w of ws) this.custom.add(w.normalize('NFC').toLowerCase()); }
  correct(w0: string, o: SpellOpts = {}): boolean {
    const w = w0.normalize('NFC');
    if (w.length < 2) return true;
    if ((o.ignoreDigits ?? true) && /\d/.test(w)) return true;
    if ((o.ignoreUpper ?? true) && w === w.toUpperCase() && /[A-ZÀ-ỸĐ]/.test(w)) return true;
    const lo = w.toLowerCase();
    if (this.custom.has(lo)) return true;
    if (this.vi.correct(w) || this.vi.correct(lo)) return true;
    if (!VI_MARK.test(w) && (this.en.correct(w) || this.en.correct(lo))) return true;
    return false;
  }
  suggest(w0: string): string[] {
    const w = w0.normalize('NFC');
    if (w.length > 24) return [];
    const a = VI_MARK.test(w) || w.length <= 4 ? this.vi.suggest(w.toLowerCase()) : [];
    const b = !VI_MARK.test(w) ? this.en.suggest(w) : [];
    const up = w[0] === w[0].toUpperCase();
    const out = [...a, ...b].map((x) => (up ? x[0].toUpperCase() + x.slice(1) : x));
    return [...new Set(out)].slice(0, 8);
  }
}
/** tách từ: chuỗi chữ cái (kể cả dấu tiếng Việt), bỏ qua địa chỉ web/email */
export function words(text: string): { w: string; i: number }[] {
  const out: { w: string; i: number }[] = [];
  const t = text.normalize('NFC');
  const skip = /(https?:\/\/\S+|www\.\S+|\S+@\S+\.\S+)/g;
  const masked = t.replace(skip, (m) => ' '.repeat(m.length));
  for (const m of masked.matchAll(/[\p{L}\p{M}][\p{L}\p{M}'’]*[\p{L}\p{M}]|[\p{L}\p{M}]/gu)) out.push({ w: m[0], i: m.index! });
  return out;
}
