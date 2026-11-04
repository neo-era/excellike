// Màu kiểu Excel: màu chủ đề (theme1.xml), sắc độ (tint) theo HLS (ECMA-376 §18.8.19),
// bảng màu đánh số (indexedColors) và màu gần đây (mruColors).

/** Màu chủ đề Office 2013–2022 theo chỉ số Excel: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4–9 accent1–6, 10 hlink, 11 folHlink */
export const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
/** Bảng 64 màu đánh số mặc định của Excel (64 = màu chữ hệ thống, 65 = nền hệ thống) */
export const DEFAULT_INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  '000000', 'FFFFFF',
];

const hls = (hex: string) => {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (mx + mn) / 2;
  if (mx !== mn) {
    const d = mx - mn;
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  return { h, s, l };
};
/** Áp sắc độ (tint −1…1) kiểu Excel: đổi độ sáng L trong HLS. hex không có dấu #, trả về không có # */
export function tintHex(hex: string, t: number): string {
  const { h, s } = hls(hex);
  let { l } = hls(hex);
  l = t < 0 ? l * (1 + t) : l * (1 - t) + t;
  const f = (p: number, q: number, x: number) => { if (x < 0) x += 1; if (x > 1) x -= 1; return x < 1 / 6 ? p + (q - p) * 6 * x : x < 1 / 2 ? q : x < 2 / 3 ? p + (q - p) * (2 / 3 - x) * 6 : p; };
  let r: number, g: number, b: number;
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}
/**
 * Năm sắc độ của một màu chủ đề trong bảng chọn màu (như Office): theo độ sáng của màu gốc
 * (trắng → tối dần; đen → sáng dần; rất tối/rất sáng → dải rộng; còn lại 80%, 60%, 40% sáng hơn, 25%, 50% tối hơn).
 */
export function paletteTints(hex: string): number[] {
  const { l } = hls(hex.replace('#', ''));
  if (l >= 0.999) return [-0.05, -0.15, -0.25, -0.35, -0.5];
  if (l <= 0.001) return [0.5, 0.35, 0.25, 0.15, 0.05];
  if (l < 0.2) return [0.9, 0.75, 0.5, 0.25, 0.1];
  if (l > 0.8) return [-0.1, -0.25, -0.5, -0.75, -0.9];
  return [0.8, 0.6, 0.4, -0.25, -0.5];
}

const ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
/** đọc 12 màu chủ đề từ theme1.xml (srgbClr hoặc sysClr lastClr); thiếu thì null */
export function parseTheme(xml: string): string[] | null {
  const sch = /<(?:\w+:)?clrScheme\b[^>]*>([\s\S]*?)<\/(?:\w+:)?clrScheme>/.exec(xml)?.[1];
  if (!sch) return null;
  const out = ORDER.map((tag, i) => {
    const m = new RegExp(`<(?:\\w+:)?${tag}>([\\s\\S]*?)</(?:\\w+:)?${tag}>`).exec(sch)?.[1] ?? '';
    const srgb = /<(?:\w+:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(m)?.[1];
    const sys = /<(?:\w+:)?sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(m)?.[1];
    return (srgb ?? sys ?? DEFAULT_THEME[i]).toUpperCase();
  });
  return out;
}
/** khối clrScheme ghi vào theme1.xml */
export function themeSchemeXml(theme: string[], name = 'Office'): string {
  const x = (tag: string) => { const c = theme[ORDER.indexOf(tag)] ?? DEFAULT_THEME[ORDER.indexOf(tag)]; return `<a:${tag}><a:srgbClr val="${c}"/></a:${tag}>`; };
  return `<a:clrScheme name="${name}">${['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(x).join('')}</a:clrScheme>`;
}
const rgbList = (xml: string) => [...xml.matchAll(/<(?:\w+:)?(?:rgbColor|color)\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/g)].map((m) => m[1].slice(-6).toUpperCase());
/** bảng màu đánh số tùy biến trong styles.xml (nếu có) */
export function parseIndexed(styles: string): string[] | null {
  const m = /<(?:\w+:)?indexedColors>([\s\S]*?)<\/(?:\w+:)?indexedColors>/.exec(styles)?.[1];
  if (!m) return null;
  const list = rgbList(m);
  return list.length ? [...list, ...DEFAULT_INDEXED.slice(list.length)] : null;
}
/** màu gần đây (mruColors) trong styles.xml */
export function parseMru(styles: string): string[] {
  const m = /<(?:\w+:)?mruColors>([\s\S]*?)<\/(?:\w+:)?mruColors>/.exec(styles)?.[1];
  return m ? rgbList(m).slice(0, 10).map((c) => '#' + c) : [];
}
/** ghi mruColors vào styles.xml (thay khối colors cũ nếu có) */
export function withMru(styles: string, mru: string[]): string {
  const block = mru.length ? `<colors><mruColors>${mru.slice(0, 10).map((c) => `<color rgb="FF${c.replace('#', '').toUpperCase()}"/>`).join('')}</mruColors></colors>` : '';
  if (/<colors>[\s\S]*?<\/colors>/.test(styles)) return styles.replace(/<colors>[\s\S]*?<\/colors>/, block);
  if (!block) return styles;
  return /<extLst>/.test(styles) ? styles.replace('<extLst>', block + '<extLst>') : styles.replace('</styleSheet>', block + '</styleSheet>');
}
