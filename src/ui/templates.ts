// Mẫu sổ làm việc (Tệp ▸ Trang đầu / Mới): dựng sẵn bố cục và công thức cho công việc hằng ngày
import type { Workbook } from '../core/workbook';

type Cell = string | number | null;
export interface Template { id: string; name: string; desc: string; icon: string; build: (wb: Workbook, s: number) => void }

const HEAD = { b: true, bg: '#DDEBF7', al: 'c' as const, wr: true, va: 'm' as const };
const NUM = '#,##0';
const NUM2 = '#,##0.00';

/** ghi một bảng 2 chiều bắt đầu tại (r0, c0) */
function put(wb: Workbook, s: number, r0: number, c0: number, rows: Cell[][]) {
  const list: [number, number, Cell][] = [];
  rows.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== '') list.push([r0 + i, c0 + j, v]); }));
  wb.writeCells(s, list);
}
const R = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
function widths(wb: Workbook, s: number, w: number[]) { w.forEach((x, c) => wb.setColWidth(s, c, x)); }
function title(wb: Workbook, s: number, text: string, sub: string, cols: number) {
  put(wb, s, 0, 0, [[text], [sub]]);
  wb.merge(s, R(0, 0, 0, cols - 1)); wb.merge(s, R(1, 0, 1, cols - 1));
  wb.setStyle(s, R(0, 0), { b: true, fs: 14, al: 'c' });
  wb.setStyle(s, R(1, 0), { i: true, al: 'c' });
}

export const TEMPLATES: Template[] = [
  {
    id: 'du-toan', name: 'Dự toán xây dựng', icon: 'calc', desc: 'Bảng dự toán theo đơn giá: khối lượng × đơn giá vật liệu, nhân công, máy',
    build(wb, s) {
      title(wb, s, 'BẢNG DỰ TOÁN CHI PHÍ XÂY DỰNG', 'Công trình: ……………………   Hạng mục: ……………………', 9);
      put(wb, s, 3, 0, [
        ['STT', 'Mã hiệu', 'Nội dung công việc', 'Đơn vị', 'Khối lượng', 'Đơn giá vật liệu', 'Đơn giá nhân công', 'Đơn giá máy', 'Thành tiền'],
        [1, 'AB.11312', 'Đào móng bằng thủ công, đất cấp II', 'm3', 12.5, 0, 285000, 0, '=E5*(F5+G5+H5)'],
        [2, 'AF.11111', 'Bê tông lót móng đá 4x6, M100', 'm3', 2.4, 980000, 310000, 45000, '=E6*(F6+G6+H6)'],
        [3, 'AF.11213', 'Bê tông móng đá 1x2, M250', 'm3', 6.8, 1350000, 420000, 85000, '=E7*(F7+G7+H7)'],
        [4, 'AF.61110', 'Cốt thép móng, đường kính ≤ 10 mm', 'tấn', 0.42, 17500000, 2800000, 150000, '=E8*(F8+G8+H8)'],
        [5, 'AF.81111', 'Ván khuôn móng', '100m2', 0.35, 4200000, 3600000, 0, '=E9*(F9+G9+H9)'],
        ['', '', 'Cộng chi phí trực tiếp', '', '', '', '', '', '=SUM(I5:I9)'],
        ['', '', 'Chi phí chung (6,5%)', '', 0.065, '', '', '', '=I10*E11'],
        ['', '', 'Thu nhập chịu thuế tính trước (5,5%)', '', 0.055, '', '', '', '=(I10+I11)*E12'],
        ['', '', 'Thuế GTGT (8%)', '', 0.08, '', '', '', '=(I10+I11+I12)*E13'],
        ['', '', 'TỔNG CỘNG', '', '', '', '', '', '=SUM(I10:I13)'],
      ]);
      widths(wb, s, [5, 11, 40, 8, 11, 15, 15, 13, 17]);
      wb.setStyle(s, R(3, 0, 3, 8), HEAD);
      wb.setBorder(s, R(3, 0, 14, 8), 'all', 'thin');
      wb.setStyle(s, R(4, 4, 8, 4), { fmt: '#,##0.000' });
      wb.setStyle(s, R(4, 5, 14, 8), { fmt: NUM });
      wb.setStyle(s, R(10, 4, 12, 4), { fmt: '0.0%' });
      wb.setStyle(s, R(4, 0, 8, 0), { al: 'c' });
      wb.setStyle(s, R(9, 2, 14, 8), { b: true });
      wb.setStyle(s, R(13, 0, 13, 8), { bg: '#FFF2CC' });
      wb.setFreeze(s, 4, 0);
    },
  },
  {
    id: 'khoi-luong', name: 'Bảng tính khối lượng', icon: 'table', desc: 'Diễn giải kích thước, số lượng cấu kiện và khối lượng từng hạng mục',
    build(wb, s) {
      title(wb, s, 'BẢNG TÍNH KHỐI LƯỢNG', 'Công trình: ……………………', 9);
      put(wb, s, 3, 0, [
        ['STT', 'Tên công việc / Diễn giải', 'Đơn vị', 'Số cấu kiện', 'Dài (m)', 'Rộng (m)', 'Cao (m)', 'Khối lượng 1 CK', 'Tổng khối lượng'],
        [1, 'Đào đất móng M1', 'm3', 6, 1.5, 1.5, 1.2, '=E5*F5*G5', '=D5*H5'],
        [2, 'Bê tông lót móng M1', 'm3', 6, 1.5, 1.5, 0.1, '=E6*F6*G6', '=D6*H6'],
        [3, 'Bê tông móng M1', 'm3', 6, 1.3, 1.3, 0.35, '=E7*F7*G7', '=D7*H7'],
        [4, 'Bê tông giằng móng GM', 'm3', 1, 42.6, 0.22, 0.3, '=E8*F8*G8', '=D8*H8'],
        [5, 'Xây tường móng gạch đặc', 'm3', 1, 42.6, 0.22, 0.6, '=E9*F9*G9', '=D9*H9'],
      ]);
      widths(wb, s, [5, 36, 8, 11, 10, 10, 10, 15, 16]);
      wb.setStyle(s, R(3, 0, 3, 8), HEAD);
      wb.setBorder(s, R(3, 0, 8, 8), 'all', 'thin');
      wb.setStyle(s, R(4, 4, 8, 7), { fmt: NUM2 });
      wb.setStyle(s, R(4, 8, 8, 8), { fmt: '#,##0.000', b: true });
      wb.setStyle(s, R(4, 0, 8, 0), { al: 'c' });
      wb.setStyle(s, R(4, 2, 8, 3), { al: 'c' });
      wb.setFreeze(s, 4, 2);
    },
  },
  {
    id: 'cham-cong', name: 'Bảng chấm công', icon: 'dateFn', desc: 'Chấm công theo ngày trong tháng, tự đếm số công',
    build(wb, s) {
      const days = Array.from({ length: 31 }, (_, i) => i + 1);
      title(wb, s, 'BẢNG CHẤM CÔNG', 'Tháng …… năm ……   (x: đi làm, p: phép, ô: ốm)', 37);
      put(wb, s, 3, 0, [
        ['STT', 'Họ và tên', 'Chức vụ', ...days, 'Số công', 'Phép', 'Ốm'],
        ...[1, 2, 3, 4, 5].map((n) => {
          const r = 3 + n + 1;
          return [n, '', '', ...days.map(() => ''), `=COUNTIF(D${r}:AH${r};"x")`, `=COUNTIF(D${r}:AH${r};"p")`, `=COUNTIF(D${r}:AH${r};"ô")`] as Cell[];
        }),
      ]);
      widths(wb, s, [5, 24, 12, ...days.map(() => 3.6), 8, 6, 6]);
      wb.setStyle(s, R(3, 0, 3, 36), HEAD);
      wb.setStyle(s, R(4, 3, 8, 36), { al: 'c' });
      wb.setBorder(s, R(3, 0, 8, 36), 'all', 'thin');
      wb.setFreeze(s, 4, 3);
    },
  },
  {
    id: 'chi-phi', name: 'Theo dõi thu chi', icon: 'currency', desc: 'Sổ thu chi theo ngày, số dư lũy kế',
    build(wb, s) {
      title(wb, s, 'SỔ THEO DÕI THU CHI', 'Đơn vị tính: đồng', 6);
      put(wb, s, 3, 0, [
        ['Ngày', 'Nội dung', 'Loại', 'Thu', 'Chi', 'Số dư'],
        ['', 'Số dư đầu kỳ', '', '', '', 0],
        ['', '', '', '', '', '=F5+D6-E6'],
        ['', '', '', '', '', '=F6+D7-E7'],
        ['', '', '', '', '', '=F7+D8-E8'],
        ['', '', '', '', '', '=F8+D9-E9'],
        ['', 'Cộng', '', '=SUM(D6:D9)', '=SUM(E6:E9)', '=F9'],
      ]);
      widths(wb, s, [12, 36, 14, 15, 15, 16]);
      wb.setStyle(s, R(3, 0, 3, 5), HEAD);
      wb.setStyle(s, R(4, 0, 9, 0), { fmt: 'dd/mm/yyyy' });
      wb.setStyle(s, R(4, 3, 9, 5), { fmt: NUM });
      wb.setStyle(s, R(9, 0, 9, 5), { b: true, bg: '#F2F2F2' });
      wb.setBorder(s, R(3, 0, 9, 5), 'all', 'thin');
    },
  },
];

/** Áp mẫu cho sổ trống vừa tạo (gọi sau khi tải lại trang) */
export function applyTemplate(wb: Workbook, s: number, id: string): string | null {
  const t = TEMPLATES.find((x) => x.id === id);
  if (!t) return null;
  t.build(wb, s);
  wb.name = t.name;
  return t.name;
}
