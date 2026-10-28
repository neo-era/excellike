// Đọc/ghi CSV: tự nhận dấu phân cách (",", ";" hoặc tab), bỏ BOM, hỗ trợ ngoặc kép
export function detectDelim(text: string): string {
  const sample = text.slice(0, 20000);
  const count: Record<string, number> = { ',': 0, ';': 0, '\t': 0 };
  let q = false;
  for (const ch of sample) {
    if (ch === '"') q = !q;
    else if (!q && ch in count) count[ch]++;
    else if (!q && ch === '\n') break;
  }
  return Object.entries(count).sort((a, b) => b[1] - a[1])[0][1] > 0 ? Object.entries(count).sort((a, b) => b[1] - a[1])[0][0] : ',';
}

export function parseCSV(text: string, delim?: string): { rows: string[][]; delim: string } {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const d = delim ?? detectDelim(text);
  const rows: string[][] = [];
  let row: string[] = [], cell = '', i = 0, q = false;
  while (i < text.length) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i += 2; continue; }
      if (ch === '"') { q = false; i++; continue; }
      cell += ch; i++; continue;
    }
    if (ch === '"' && cell === '') { q = true; i++; continue; }
    if (ch === d) { row.push(cell); cell = ''; i++; continue; }
    if (ch === '\r' || ch === '\n') {
      row.push(cell); rows.push(row); row = []; cell = '';
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      continue;
    }
    cell += ch; i++;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return { rows, delim: d };
}

export function toCSV(rows: string[][], delim = ',', bom = true): string {
  const q = (s: string) => (s.includes(delim) || /["\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
  return (bom ? '﻿' : '') + rows.map((r) => r.map(q).join(delim)).join('\r\n');
}
