import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { Speller, words } from '../src/core/spell';
const rd = (f: string) => readFileSync(new URL('../src/dict/' + f, import.meta.url), 'utf8');
describe('Chính tả', () => {
  const sp = new Speller({ aff: rd('vi.aff'), dic: rd('vi.dic') }, { aff: rd('en.aff'), dic: rd('en.dic') });
  it('nhận đúng tiếng Việt + Anh, bắt lỗi gõ sai', () => {
    for (const w of ['Thành', 'tiền', 'khối', 'lượng', 'chiếu', 'sáng', 'công', 'trình', 'cable', 'lighting', 'HCM', 'LED250'])
      expect(sp.correct(w), w).toBe(true);
    for (const w of ['tiềnn', 'khôis', 'chieeus', 'lightng']) expect(sp.correct(w), w).toBe(false);
    expect(sp.suggest('lightng')).toContain('lighting');
    sp.addWords(['Sapulico']);
    expect(sp.correct('Sapulico')).toBe(true);
  });
  it('tách từ bỏ qua web/email', () => {
    expect(words('Dự toán www.abc.vn gửi a@b.vn ok').map((x) => x.w)).toEqual(['Dự', 'toán', 'gửi', 'ok']);
  });
});
