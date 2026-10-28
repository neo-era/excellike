// Chuẩn hóa chuỗi người dùng gõ trước khi đưa vào bộ tính
export type Content = string | number | boolean | null;

export function normalizeInput(raw: string): { content: Content; fmt?: string } {
  if (raw === '') return { content: null };
  if (raw.startsWith('=') && raw.length > 1) return { content: raw };
  const t = raw.trim();
  // Số có dấu chấm nhóm nghìn kiểu Việt Nam: 1.250 hoặc -1.250.000,5
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) {
    return { content: +t.replace(/\./g, '').replace(',', '.') };
  }
  const pm = /^-?\d+(?:,(\d+))?\s*%$/.exec(t);
  if (pm) return { content: t.replace(/\s/g, ''), fmt: pm[1] ? '0.' + '0'.repeat(Math.min(pm[1].length, 4)) + '%' : '0%' };
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(t)) return { content: t, fmt: 'dd/mm/yyyy' };
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}$/.test(t)) return { content: t, fmt: 'dd/mm/yyyy hh:mm' };
  if (/^(true|false)$/i.test(t)) return { content: t.toUpperCase() === 'TRUE' };
  return { content: raw };
}
