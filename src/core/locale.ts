// Đổi công thức giữa kiểu Excel tiếng Việt (dùng trong app: ";" ngăn tham số, "," thập phân)
// và kiểu gốc lưu trong tệp .xlsx (",", ".").
type Mode = 'en' | 'vi';

function convert(f: string, to: Mode): string {
  if (!f.startsWith('=')) return f;
  const argFrom = to === 'en' ? ';' : ',', argTo = to === 'en' ? ',' : ';';
  const decFrom = to === 'en' ? ',' : '.', decTo = to === 'en' ? '.' : ',';
  let out = '';
  let i = 0;
  let inArray = 0;
  while (i < f.length) {
    const ch = f[i];
    // hằng mảng {…}: cột luôn dùng ",", hàng là ";" trong tệp và "|" trong app
    if (ch === '{') { inArray++; out += ch; i++; continue; }
    if (ch === '}') { inArray = Math.max(0, inArray - 1); out += ch; i++; continue; }
    if (inArray) { out += ch === (to === 'vi' ? ';' : '|') ? (to === 'vi' ? '|' : ';') : ch; i++; continue; }
    if (ch === '"') { // chuỗi
      let j = i + 1;
      while (j < f.length) { if (f[j] === '"' && f[j + 1] === '"') { j += 2; continue; } if (f[j] === '"') break; j++; }
      out += f.slice(i, j + 1); i = j + 1; continue;
    }
    if (ch === "'") { // tên sheet trong nháy
      let j = i + 1;
      while (j < f.length) { if (f[j] === "'" && f[j + 1] === "'") { j += 2; continue; } if (f[j] === "'") break; j++; }
      // 'Sheet'!#REF! (tham chiếu đã hỏng trong Excel) → #REF! như Excel hiển thị
      if (to === 'vi' && f.startsWith('!#REF!', j + 1)) { out += '#REF!'; i = j + 7; continue; }
      out += f.slice(i, j + 1); i = j + 1; continue;
    }
    // liên kết ngoài viết liền [1]Sheet1!A1 → '[1]Sheet1'!A1 (bộ tính chỉ hiểu dạng có nháy)
    if (to === 'vi' && ch === '[') {
      const m = /^\[(\d+)\]([^'!\s()+\-*\/^&=<>,;:"{}[\]]+)!/.exec(f.slice(i));
      if (m && !/[\w.\u00C0-\uFFFF\]]/.test(i > 0 ? f[i - 1] : '')) { out += `'[${m[1]}]${m[2]}'!`; i += m[0].length; continue; }
    }
    if (to === 'vi' && ch === '!' && f.startsWith('!#REF!', i)) {
      out = out.replace(/(\[[^\]]*\])?[A-Za-z_À-￿][\w.À-￿]*$/, '') + '#REF!'; i += 6; continue;
    }
    const prev = i > 0 ? f[i - 1] : '';
    if (/[0-9]/.test(ch) && !/[A-Za-z_$.0-9À-￿]/.test(prev)) {
      // số: phần nguyên, phần thập phân, số mũ
      let j = i;
      while (j < f.length && /[0-9]/.test(f[j])) j++;
      let num = f.slice(i, j);
      if (f[j] === decFrom && /[0-9]/.test(f[j + 1] ?? '')) {
        let k = j + 1;
        while (k < f.length && /[0-9]/.test(f[k])) k++;
        num += decTo + f.slice(j + 1, k);
        j = k;
      }
      if (/[eE]/.test(f[j] ?? '') && /[+\-0-9]/.test(f[j + 1] ?? '')) {
        let k = j + 1;
        if (/[+-]/.test(f[k])) k++;
        while (k < f.length && /[0-9]/.test(f[k])) k++;
        num += f.slice(j, k);
        j = k;
      }
      // tham chiếu hàng kiểu 1:3 hoặc phần số của A1 đã bị loại bởi prev; ở đây chỉ là số
      out += num; i = j; continue;
    }
    if (ch === argFrom) { out += argTo; i++; continue; }
    out += ch; i++;
  }
  return out;
}

export const toEn = (f: string) => convert(f, 'en');
export const toVi = (f: string) => convert(f, 'vi');
