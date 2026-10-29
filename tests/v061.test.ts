import { describe, it, expect } from 'vitest';
import { setLang, tr } from '../src/ui/i18n';
import { EN2 } from '../src/ui/i18n-extra';
import { EN } from '../src/ui/i18n-en';
import { relTime } from '../src/ui/backstage';
import { Workbook } from '../src/core/workbook';
import { TEMPLATES, applyTemplate } from '../src/ui/templates';

describe('0.6.1 ngôn ngữ', () => {
  it('dịch nhãn, mẹo có phím tắt và ghi chú phạm vi', () => {
    setLang('en');
    expect(tr('Tự động hóa')).toBe('Automate');
    expect(tr('Đậm (Ctrl+B)')).toBe('Bold (Ctrl+B)');
    expect(tr('Bảng (Ctrl+T) — có ở bản 0.7')).toBe('Table (Ctrl+T) — available in version 0.7');
    expect(tr('Bản đồ — ngoài phạm vi (Bing)')).toBe('Maps — out of scope (Bing)');
    expect(tr('  Kiểu Ô ')).toBe('  Cell Styles ');
    expect(tr('3 phút trước')).toBe('3 min ago');
    setLang('vi');
    expect(tr('Tự động hóa')).toBe('Tự động hóa');
  });
  it('không ghép nửa Anh nửa Việt khi phần ghép là nhãn chưa dịch', () => {
    setLang('en');
    expect(tr('Bỏ khỏi Ưa thích')).toBe('Remove from Favorites');
    expect(tr('Bỏ xyz ạ')).toBe('Bỏ xyz ạ');
    expect(tr('Ô A1')).toBe('Cell A1');
    // dữ liệu người dùng vẫn được giữ nguyên trong mẫu thường
    expect(tr('Mẫu: Bảng dự toán')).toBe('Sample: Bảng dự toán');
    // tên do người dùng đặt không bị dịch dù trùng chữ giao diện
    expect(tr('Tổng (đang mở)')).toBe('Tổng (open)');
    expect(tr('Đã xóa trang tính “Mới”.')).toBe('Deleted sheet "Mới".');
    expect(tr('Đặt tên rõ nghĩa cho trang tính “Sheet1”')).toBe('Give sheet "Sheet1" a meaningful name');
    expect(tr('“Bản đồ”: ngoài phạm vi (Bing).')).toBe('"Maps": out of scope (Bing).');
    setLang('vi');
  });
  it('mọi chuỗi bổ sung có bản dịch không rỗng, không còn dấu tiếng Việt', () => {
    for (const [k, v] of Object.entries({ ...EN, ...EN2 })) {
      expect(v.trim(), k).not.toBe('');
      if (k !== v && !/[ÀẢÃÁẠ]|“|Zalo/.test(k)) expect(/[À-ÖØ-öø-ỹĐđ]/.test(v), `${k} → ${v}`).toBe(false);
    }
  });
});

describe('0.6.1 Backstage và mẫu', () => {
  it('ngày sửa đổi tương đối', () => {
    const now = new Date(2026, 9, 29, 15, 0).getTime();
    expect(relTime(now - 20_000, now)).toBe('Vừa xong');
    expect(relTime(now - 5 * 60_000, now)).toBe('5 phút trước');
    expect(relTime(new Date(2026, 9, 29, 9, 5).getTime(), now)).toMatch(/^Hôm nay lúc 0?9:05$/);
    expect(relTime(new Date(2026, 9, 28, 14, 20).getTime(), now)).toMatch(/^Hôm qua lúc 14:20$/);
    expect(relTime(new Date(2026, 8, 1).getTime(), now)).toMatch(/01\/09\/2026/);
  });
  it('mỗi mẫu dựng được, không lỗi công thức', () => {
    for (const t of TEMPLATES) {
      const wb = Workbook.create();
      const s = wb.sheetIds()[0];
      expect(applyTemplate(wb, s, t.id)).toBe(t.name);
      const u = wb.used(s);
      for (let r = 0; r < u.rows; r++) for (let c = 0; c < u.cols; c++) {
        const v = wb.value(s, r, c);
        expect(typeof v === 'string' && v.startsWith('#') && /^#[A-Z/0!?]+/.test(v), `${t.id} ${r},${c}: ${v}`).toBe(false);
      }
    }
  });
  it('mẫu Dự toán: tổng = trực tiếp + chung + thu nhập + VAT', () => {
    const wb = Workbook.create(); const s = wb.sheetIds()[0];
    applyTemplate(wb, s, 'du-toan');
    const T = wb.value(s, 9, 8) as number;
    const C = T * 0.065, L = (T + C) * 0.055, V = (T + C + L) * 0.08;
    expect(wb.value(s, 13, 8)).toBeCloseTo(T + C + L + V, 3);
    expect(applyTemplate(wb, s, 'khong-co')).toBeNull();
  });
});
