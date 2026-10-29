// Bảng ánh xạ phím tắt — NGUỒN DUY NHẤT cho: bộ điều phối phím, bảng phím tắt F1, tài liệu rà soát và test e2e.
// Theo Microsoft Support “Keyboard shortcuts in Excel” (Windows), tra ngày 29/10/2026, cùng các phím Excel chuẩn
// có trong chú thích Ribbon (Ctrl+F, Ctrl+H, Ctrl+G, Ctrl+[…).
// Tổ hợp chuẩn hoá: Ctrl+Alt+Shift+<phím gốc>; phím gốc lấy theo vị trí (e.code) nên Ctrl+Shift+$ = 'Ctrl+Shift+4'.

export type KeyCtx = 'nav' | 'edit' | 'any';
export type KeyStatus = 'ok' | 'later' | 'no';
export interface KeyDef {
  id: string;
  /** cách viết hiển thị (như Microsoft) */
  label: string;
  /** tổ hợp chuẩn hoá; rỗng = phím đặc biệt (chuỗi phím, chuột, KeyTips…) xử lý chỗ khác */
  combos: string[];
  ctx: KeyCtx;
  desc: string;
  group: string;
  status: KeyStatus;
  /** later: bản sẽ có; no: lý do không làm */
  note?: string;
}

const G = {
  common: 'Thường dùng',
  ribbon: 'Ribbon & KeyTips',
  nav: 'Di chuyển trong ô',
  fmt: 'Định dạng ô',
  paste: 'Hộp Dán Đặc biệt',
  sel: 'Chọn & thao tác',
  data: 'Dữ liệu, hàm, thanh công thức',
  fn: 'Phím chức năng',
  other: 'Phím khác',
} as const;

const k = (id: string, label: string, combos: string[], ctx: KeyCtx, desc: string, group: string, status: KeyStatus = 'ok', note?: string): KeyDef =>
  ({ id, label, combos, ctx, desc, group, status, note });

export const KEYMAP: KeyDef[] = [
  // ---------- Thường dùng ----------
  k('closeBook', 'Ctrl+W', ['Ctrl+W'], 'any', 'Đóng sổ làm việc', G.common),
  k('open', 'Ctrl+O', ['Ctrl+O'], 'any', 'Mở sổ làm việc', G.common),
  k('new', 'Ctrl+N', ['Ctrl+N'], 'any', 'Sổ làm việc mới', G.common),
  k('save', 'Ctrl+S', ['Ctrl+S'], 'any', 'Lưu sổ làm việc', G.common),
  k('print', 'Ctrl+P', ['Ctrl+P'], 'any', 'In (xem trước trong Tệp ▸ In)', G.common),
  k('copy', 'Ctrl+C', ['Ctrl+C', 'Ctrl+Insert'], 'nav', 'Sao chép', G.common),
  k('paste', 'Ctrl+V', ['Ctrl+V', 'Shift+Insert'], 'nav', 'Dán', G.common),
  k('cut', 'Ctrl+X', ['Ctrl+X'], 'nav', 'Cắt', G.common),
  k('undo', 'Ctrl+Z', ['Ctrl+Z', 'Alt+Backspace'], 'any', 'Hoàn tác', G.common),
  k('redo', 'Ctrl+Y', ['Ctrl+Y', 'Ctrl+Shift+Z'], 'any', 'Làm lại; nếu không có gì để làm lại thì lặp lại lệnh cuối', G.common),
  k('del', 'Delete', ['Delete'], 'nav', 'Xóa nội dung ô (giữ định dạng)', G.common),
  k('fillColor', 'Alt+H, H', [], 'nav', 'Chọn màu tô (KeyTips)', G.common),
  k('centerAl', 'Alt+H, A, C', [], 'nav', 'Căn giữa (KeyTips)', G.common),
  k('borderKt', 'Alt+H, B', [], 'nav', 'Thêm viền (KeyTips)', G.common),
  k('delColKt', 'Alt+H, D, C', [], 'nav', 'Xóa cột (KeyTips)', G.common),
  k('bold', 'Ctrl+B', ['Ctrl+B', 'Ctrl+2'], 'nav', 'Đậm / bỏ đậm', G.common),
  k('ctxMenu', 'Shift+F10', ['Shift+F10', 'ContextMenu'], 'nav', 'Mở menu chuột phải', G.common),
  k('hideRows', 'Ctrl+9', ['Ctrl+9'], 'nav', 'Ẩn hàng đang chọn', G.common),
  k('hideCols', 'Ctrl+0', ['Ctrl+0'], 'nav', 'Ẩn cột đang chọn', G.common),
  k('unhideRows', 'Ctrl+Shift+(', ['Ctrl+Shift+9'], 'nav', 'Bỏ ẩn hàng trong vùng chọn', G.common),
  k('unhideCols', 'Ctrl+Shift+)', ['Ctrl+Shift+0'], 'nav', 'Bỏ ẩn cột trong vùng chọn', G.common),

  // ---------- Ribbon ----------
  k('tabFile', 'Alt+F', [], 'any', 'Mở thẻ Tệp', G.ribbon),
  k('tabHome', 'Alt+H', [], 'any', 'Mở thẻ Trang đầu', G.ribbon),
  k('tabInsert', 'Alt+N', [], 'any', 'Mở thẻ Chèn', G.ribbon),
  k('tabLayout', 'Alt+P', [], 'any', 'Mở thẻ Bố trí Trang', G.ribbon),
  k('tabFormulas', 'Alt+M', [], 'any', 'Mở thẻ Công thức', G.ribbon),
  k('tabData', 'Alt+A', [], 'any', 'Mở thẻ Dữ liệu', G.ribbon),
  k('tabReview', 'Alt+R', [], 'any', 'Mở thẻ Xem lại', G.ribbon),
  k('tabView', 'Alt+W', [], 'any', 'Mở thẻ Xem', G.ribbon),
  k('tell', 'Alt+Q', ['Alt+Q'], 'any', 'Tới ô Tìm kiếm lệnh', G.ribbon),
  k('keytips', 'Alt hoặc F10', ['F10'], 'any', 'Hiện/ẩn KeyTips', G.ribbon),
  k('ribbonTab', 'Tab / Shift+Tab', [], 'any', 'Chuyển giữa các lệnh trên Ribbon (khi Ribbon đang nhận phím)', G.ribbon),
  k('ribbonArrows', 'Phím mũi tên', [], 'any', 'Di chuyển giữa các mục trên Ribbon', G.ribbon),
  k('ribbonTip', 'Ctrl+Shift+F10', ['Ctrl+Shift+F10'], 'any', 'Hiện chú thích của lệnh đang chọn trên Ribbon', G.ribbon),
  k('ribbonGo', 'Space hoặc Enter', [], 'any', 'Thực hiện lệnh đang chọn trên Ribbon', G.ribbon),
  k('ribbonDrop', 'Alt+↓ / ↓', [], 'any', 'Mở danh sách/menu của lệnh đang chọn', G.ribbon),
  k('ribbonToggle', 'Ctrl+F1', ['Ctrl+F1'], 'any', 'Thu gọn / hiện Ribbon', G.ribbon),
  k('ribbonGroups', 'Ctrl+← / Ctrl+→', [], 'any', 'Nhảy giữa các nhóm lệnh trên Ribbon', G.ribbon),

  // ---------- Di chuyển ----------
  k('arrows', '↑ ↓ ← →', ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'], 'nav', 'Di chuyển một ô', G.nav),
  k('tab', 'Tab / Shift+Tab', ['Tab', 'Shift+Tab'], 'nav', 'Sang ô phải / trái (đi vòng trong vùng chọn)', G.nav),
  k('jump', 'Ctrl+mũi tên', ['Ctrl+ArrowUp', 'Ctrl+ArrowDown', 'Ctrl+ArrowLeft', 'Ctrl+ArrowRight'], 'nav', 'Tới mép vùng dữ liệu', G.nav),
  k('endMode', 'End, mũi tên', ['End'], 'nav', 'Chế độ End: End rồi mũi tên = tới ô có dữ liệu kế tiếp', G.nav),
  k('lastCell', 'Ctrl+End', ['Ctrl+End'], 'nav', 'Tới ô cuối cùng có dùng', G.nav),
  k('lastCellExt', 'Ctrl+Shift+End', ['Ctrl+Shift+End'], 'nav', 'Mở rộng vùng chọn tới ô cuối cùng có dùng', G.nav),
  k('home', 'Home', ['Home'], 'nav', 'Về đầu hàng', G.nav),
  k('top', 'Ctrl+Home', ['Ctrl+Home'], 'nav', 'Về đầu trang tính (sau ngăn cố định)', G.nav),
  k('pgDn', 'Page Down', ['PageDown'], 'nav', 'Xuống một màn hình', G.nav),
  k('pgUp', 'Page Up', ['PageUp'], 'nav', 'Lên một màn hình', G.nav),
  k('pgRight', 'Alt+Page Down', ['Alt+PageDown'], 'nav', 'Sang phải một màn hình', G.nav),
  k('pgLeft', 'Alt+Page Up', ['Alt+PageUp'], 'nav', 'Sang trái một màn hình', G.nav),
  k('nextSheet', 'Ctrl+Page Down', ['Ctrl+PageDown'], 'any', 'Sang trang tính sau', G.nav),
  k('prevSheet', 'Ctrl+Page Up', ['Ctrl+PageUp'], 'any', 'Về trang tính trước', G.nav),
  k('dvList', 'Alt+↓', ['Alt+ArrowDown'], 'nav', 'Mở danh sách chọn (kiểm tra dữ liệu, nút lọc, hoặc chọn từ các giá trị trong cột)', G.nav),
  k('shapes', 'Ctrl+Alt+5, Tab', ['Ctrl+Alt+5'], 'nav', 'Duyệt qua các biểu đồ/đối tượng nổi (Esc để thoát)', G.nav),
  k('hScroll', 'Ctrl+Shift+cuộn chuột', [], 'nav', 'Cuộn ngang', G.nav),
  k('zoomIn', 'Ctrl+Alt+=', ['Ctrl+Alt+='], 'any', 'Phóng to', G.nav),
  k('zoomOut', 'Ctrl+Alt+−', ['Ctrl+Alt+-'], 'any', 'Thu nhỏ', G.nav),

  // ---------- Định dạng ô ----------
  k('fmtCells', 'Ctrl+1', ['Ctrl+1'], 'nav', 'Hộp thoại Định dạng Ô', G.fmt),
  k('fmtFont', 'Ctrl+Shift+F', ['Ctrl+Shift+F', 'Ctrl+Shift+P'], 'nav', 'Định dạng Ô ▸ thẻ Phông', G.fmt),
  k('edit', 'F2', ['F2'], 'nav', 'Sửa ô hiện hành', G.fmt),
  k('note', 'Shift+F2', ['Shift+F2'], 'nav', 'Chèn/sửa ghi chú ô', G.fmt, 'later', '0.7 (Ghi chú ô: đọc/ghi .xlsx)'),
  k('comment', 'Ctrl+Shift+F2', ['Ctrl+Shift+F2'], 'nav', 'Chèn/trả lời chú thích dạng chuỗi', G.fmt, 'later', '0.7 (Chú thích)'),
  k('insertDlg', 'Ctrl+Shift++', ['Ctrl+Shift+=', 'Ctrl+NumpadAdd'], 'nav', 'Hộp thoại Chèn ô (chọn cả hàng/cột thì chèn luôn)', G.fmt),
  k('deleteDlg', 'Ctrl+−', ['Ctrl+-', 'Ctrl+NumpadSubtract'], 'nav', 'Hộp thoại Xóa ô (chọn cả hàng/cột thì xóa luôn)', G.fmt),
  k('nowTime', 'Ctrl+Shift+:', ['Ctrl+Shift+;'], 'any', 'Nhập giờ hiện tại', G.fmt),
  k('today', 'Ctrl+;', ['Ctrl+;'], 'any', 'Nhập ngày hiện tại', G.fmt),
  k('showF', 'Ctrl+`', ['Ctrl+`'], 'nav', 'Hiện công thức / giá trị', G.fmt),
  k('fAbove', "Ctrl+'", ["Ctrl+'"], 'any', 'Chép công thức ô phía trên vào ô hiện hành', G.fmt),
  k('pasteSp', 'Ctrl+Alt+V', ['Ctrl+Alt+V'], 'nav', 'Hộp thoại Dán Đặc biệt', G.fmt),
  k('pasteVals', 'Ctrl+Shift+V', ['Ctrl+Shift+V'], 'nav', 'Dán chỉ giá trị (Excel 365)', G.fmt),
  k('italic', 'Ctrl+I', ['Ctrl+I', 'Ctrl+3'], 'nav', 'Nghiêng / bỏ nghiêng', G.fmt),
  k('under', 'Ctrl+U', ['Ctrl+U', 'Ctrl+4'], 'nav', 'Gạch chân / bỏ gạch chân', G.fmt),
  k('strike', 'Ctrl+5', ['Ctrl+5'], 'nav', 'Gạch ngang / bỏ gạch ngang', G.fmt),
  k('objects', 'Ctrl+6', ['Ctrl+6'], 'nav', 'Hiện / ẩn / hiện chỗ giữ cho đối tượng (biểu đồ)', G.fmt),
  k('outlineBd', 'Ctrl+Shift+&', ['Ctrl+Shift+7'], 'nav', 'Kẻ viền ngoài', G.fmt),
  k('noBd', 'Ctrl+Shift+_', ['Ctrl+Shift+-'], 'nav', 'Bỏ viền ngoài', G.fmt),
  k('outline', 'Ctrl+8', ['Ctrl+8'], 'nav', 'Hiện/ẩn ký hiệu dàn bài (nhóm hàng/cột)', G.fmt, 'later', '0.7 (Nhóm hàng/cột)'),
  k('fillDown', 'Ctrl+D', ['Ctrl+D'], 'nav', 'Điền xuống', G.fmt),
  k('fillRight', 'Ctrl+R', ['Ctrl+R'], 'nav', 'Điền sang phải', G.fmt),
  k('nfGeneral', 'Ctrl+Shift+~', ['Ctrl+Shift+`'], 'nav', 'Định dạng Chung', G.fmt),
  k('nfCurrency', 'Ctrl+Shift+$', ['Ctrl+Shift+4'], 'nav', 'Định dạng Tiền tệ (2 số lẻ, số âm trong ngoặc)', G.fmt),
  k('nfPercent', 'Ctrl+Shift+%', ['Ctrl+Shift+5'], 'nav', 'Định dạng Phần trăm, không số lẻ', G.fmt),
  k('nfSci', 'Ctrl+Shift+^', ['Ctrl+Shift+6'], 'nav', 'Định dạng Khoa học, 2 số lẻ', G.fmt),
  k('nfDate', 'Ctrl+Shift+#', ['Ctrl+Shift+3'], 'nav', 'Định dạng Ngày (ngày, tháng, năm)', G.fmt),
  k('nfTime', 'Ctrl+Shift+@', ['Ctrl+Shift+2'], 'nav', 'Định dạng Giờ (giờ, phút, SA/CH)', G.fmt),
  k('nfNumber', 'Ctrl+Shift+!', ['Ctrl+Shift+1'], 'nav', 'Định dạng Số: 2 số lẻ, dấu phân cách hàng nghìn', G.fmt),
  k('link', 'Ctrl+K', ['Ctrl+K'], 'nav', 'Chèn / sửa liên kết', G.fmt),
  k('spell', 'F7', ['F7'], 'nav', 'Kiểm tra chính tả', G.fmt, 'later', '0.9 (cần từ điển tiếng Việt)'),
  k('quick', 'Ctrl+Q', ['Ctrl+Q'], 'nav', 'Phân tích Nhanh cho vùng chọn', G.fmt),
  k('table', 'Ctrl+L hoặc Ctrl+T', ['Ctrl+L', 'Ctrl+T'], 'nav', 'Tạo Bảng', G.fmt, 'later', '0.7 (Bảng + tham chiếu có cấu trúc)'),
  k('stats', 'Ctrl+Shift+G', ['Ctrl+Shift+G'], 'nav', 'Thống kê Sổ làm việc', G.fmt),

  // ---------- Dán đặc biệt (trong hộp thoại) ----------
  k('psA', 'A', [], 'any', 'Dán tất cả', G.paste),
  k('psF', 'F', [], 'any', 'Chỉ công thức', G.paste),
  k('psV', 'V', [], 'any', 'Chỉ giá trị', G.paste),
  k('psT', 'T', [], 'any', 'Chỉ định dạng', G.paste),
  k('psC', 'C', [], 'any', 'Chỉ chú thích & ghi chú', G.paste),
  k('psN', 'N', [], 'any', 'Chỉ kiểm tra dữ liệu', G.paste),
  k('psH', 'H', [], 'any', 'Tất cả, dùng chủ đề nguồn', G.paste),
  k('psX', 'X', [], 'any', 'Tất cả trừ viền', G.paste),
  k('psW', 'W', [], 'any', 'Độ rộng cột', G.paste),
  k('psR', 'R', [], 'any', 'Công thức và định dạng số', G.paste),
  k('psU', 'U', [], 'any', 'Giá trị và định dạng số', G.paste),

  // ---------- Chọn & thao tác ----------
  k('selAll', 'Ctrl+A / Ctrl+Shift+Space', ['Ctrl+A', 'Ctrl+Shift+Space'], 'nav', 'Chọn vùng dữ liệu hiện tại; bấm lần nữa chọn cả trang tính', G.sel),
  k('grpNext', 'Ctrl+Shift+Page Down', ['Ctrl+Shift+PageDown'], 'nav', 'Chọn thêm trang tính sau (nhóm trang tính)', G.sel),
  k('grpPrev', 'Ctrl+Shift+Page Up', ['Ctrl+Shift+PageUp'], 'nav', 'Chọn thêm trang tính trước (nhóm trang tính)', G.sel),
  k('extOne', 'Shift+mũi tên', ['Shift+ArrowUp', 'Shift+ArrowDown', 'Shift+ArrowLeft', 'Shift+ArrowRight'], 'nav', 'Mở rộng vùng chọn một ô', G.sel),
  k('extJump', 'Ctrl+Shift+mũi tên', ['Ctrl+Shift+ArrowUp', 'Ctrl+Shift+ArrowDown', 'Ctrl+Shift+ArrowLeft', 'Ctrl+Shift+ArrowRight'], 'nav', 'Mở rộng vùng chọn tới ô có dữ liệu cuối', G.sel),
  k('extMode', 'F8', ['F8'], 'nav', 'Bật/tắt chế độ Mở rộng Vùng chọn', G.sel),
  k('addMode', 'Shift+F8', ['Shift+F8'], 'nav', 'Thêm vùng không liền kề vào vùng chọn', G.sel),
  k('addClick', 'Ctrl+nhấp chuột', [], 'nav', 'Thêm ô/vùng không liền kề vào vùng chọn', G.sel),
  k('newLine', 'Alt+Enter', [], 'edit', 'Xuống dòng trong ô', G.sel),
  k('fillSel', 'Ctrl+Enter', [], 'edit', 'Điền nội dung đang nhập vào mọi ô đang chọn', G.sel),
  k('enterUp', 'Shift+Enter', ['Shift+Enter'], 'any', 'Hoàn tất và lên ô trên', G.sel),
  k('selCol', 'Ctrl+Space', ['Ctrl+Space'], 'nav', 'Chọn cả cột', G.sel),
  k('selRow', 'Shift+Space', ['Shift+Space'], 'nav', 'Chọn cả hàng', G.sel),
  k('topExt', 'Ctrl+Shift+Home', ['Ctrl+Shift+Home'], 'nav', 'Mở rộng vùng chọn tới đầu trang tính', G.sel),
  k('homeExt', 'Shift+Home', ['Shift+Home'], 'nav', 'Mở rộng vùng chọn về đầu hàng', G.sel),
  k('pgExt', 'Shift+Page Down / Up', ['Shift+PageDown', 'Shift+PageUp'], 'nav', 'Mở rộng vùng chọn một màn hình', G.sel),
  k('region', 'Ctrl+Shift+*', ['Ctrl+Shift+8', 'Ctrl+NumpadMultiply'], 'nav', 'Chọn vùng dữ liệu quanh ô hiện hành', G.sel),
  k('repeat', 'F4 (khi không sửa ô)', [], 'nav', 'Lặp lại lệnh/thao tác cuối', G.sel),
  k('visible', 'Alt+;', ['Alt+;'], 'nav', 'Chỉ chọn các ô đang hiện', G.sel),
  k('prec', 'Ctrl+[', ['Ctrl+['], 'nav', 'Chọn ô tiền lệ trực tiếp', G.sel),
  k('precAll', 'Ctrl+Shift+{', ['Ctrl+Shift+['], 'nav', 'Chọn mọi ô tiền lệ', G.sel),
  k('dep', 'Ctrl+]', ['Ctrl+]'], 'nav', 'Chọn ô phụ thuộc trực tiếp', G.sel),
  k('depAll', 'Ctrl+Shift+}', ['Ctrl+Shift+]'], 'nav', 'Chọn mọi ô phụ thuộc', G.sel),
  k('rowDiff', 'Ctrl+\\', ['Ctrl+\\'], 'nav', 'Chọn ô khác biệt trong hàng (so với cột của ô hiện hành)', G.sel),
  k('colDiff', 'Ctrl+Shift+|', ['Ctrl+Shift+\\'], 'nav', 'Chọn ô khác biệt trong cột (so với hàng của ô hiện hành)', G.sel),
  k('selNotes', 'Ctrl+Shift+O', ['Ctrl+Shift+O'], 'nav', 'Chọn các ô có ghi chú', G.sel),
  k('corner', 'Ctrl+.', ['Ctrl+.'], 'nav', 'Ô hiện hành chuyển tới góc kế tiếp của vùng chọn', G.sel),
  k('collapse', 'Shift+Backspace', ['Shift+Backspace'], 'nav', 'Thu vùng chọn về ô hiện hành', G.sel),
  k('showActive', 'Ctrl+Backspace', ['Ctrl+Backspace'], 'nav', 'Cuộn tới ô hiện hành', G.sel),

  // ---------- Dữ liệu, hàm, thanh công thức ----------
  k('fTips', 'Ctrl+Alt+P', ['Ctrl+Alt+P'], 'any', 'Bật/tắt chú thích hàm khi gõ công thức', G.data),
  k('fxTall', 'Ctrl+Shift+U', ['Ctrl+Shift+U'], 'any', 'Mở rộng / thu gọn thanh công thức', G.data),
  k('cancel', 'Esc', ['Escape'], 'any', 'Hủy nhập; đóng menu/hộp thoại; bỏ viền sao chép', G.data),
  k('enter', 'Enter', ['Enter'], 'any', 'Hoàn tất và xuống ô dưới (Tab rồi Enter: về cột bắt đầu)', G.data),
  k('cse', 'Ctrl+Shift+Enter', [], 'edit', 'Nhập công thức mảng (Excel 365: như Enter, mảng tự tràn)', G.data),
  k('calcAll', 'F9', ['F9'], 'any', 'Tính lại mọi trang tính (khi sửa: tính phần công thức đang bôi đen)', G.data),
  k('calcSheet', 'Shift+F9', ['Shift+F9'], 'nav', 'Tính lại trang tính hiện hành', G.data),
  k('calcFull', 'Ctrl+Alt+F9', ['Ctrl+Alt+F9', 'Ctrl+Alt+Shift+F9'], 'nav', 'Tính lại toàn bộ, kể cả ô không đổi', G.data),
  k('errMenu', 'Alt+Shift+F10', ['Alt+Shift+F10'], 'nav', 'Menu nút Kiểm tra Lỗi', G.data, 'later', '0.9 (Kiểm tra Lỗi)'),
  k('fArgs', 'Ctrl+A (sau tên hàm)', [], 'edit', 'Hộp thoại Đối số Hàm', G.data),
  k('fArgNames', 'Ctrl+Shift+A (sau tên hàm)', [], 'edit', 'Chèn tên đối số và ngoặc', G.data),
  k('autoSum', 'Alt+=', ['Alt+='], 'nav', 'Tổng Tự động', G.data),
  k('flash', 'Ctrl+E', ['Ctrl+E'], 'nav', 'Điền Nhanh (Flash Fill)', G.data),
  k('f4', 'F4 (khi sửa công thức)', [], 'edit', 'Đổi tham chiếu A1 → $A$1 → A$1 → $A1', G.data),
  k('insFn', 'Shift+F3', ['Shift+F3'], 'any', 'Chèn Hàm', G.data),
  k('vAbove', 'Ctrl+Shift+"', ["Ctrl+Shift+'"], 'any', 'Chép giá trị ô phía trên vào ô hiện hành', G.data),
  k('chartHere', 'Alt+F1', ['Alt+F1'], 'nav', 'Tạo biểu đồ nhúng từ vùng hiện tại', G.data),
  k('chartSheet', 'F11', ['F11'], 'nav', 'Tạo biểu đồ trên trang tính riêng', G.data),
  k('defName', 'Alt+M, M, D', [], 'nav', 'Xác định tên', G.data, 'later', '0.7 (Tên vùng)'),
  k('pasteName', 'F3', ['F3'], 'any', 'Dán tên đã xác định', G.data, 'later', '0.7 (Tên vùng)'),
  k('macros', 'Alt+F8 / Alt+F11 / Alt+F12', [], 'any', 'Macro, VBA, Power Query', G.data, 'no', 'Không chạy macro/VBA/Power Query'),
  k('filter', 'Ctrl+Shift+L', ['Ctrl+Shift+L'], 'nav', 'Bật/tắt Lọc', G.data),
  k('refresh', 'Ctrl+F5', ['Ctrl+F5'], 'nav', 'Làm mới dữ liệu trang tính (PivotTable trên trang)', G.data),
  k('refreshAll', 'Ctrl+Alt+F5', ['Ctrl+Alt+F5'], 'nav', 'Làm mới tất cả (mọi PivotTable)', G.data),

  // ---------- Phím chức năng ----------
  k('help', 'F1', ['F1'], 'any', 'Trợ giúp: bảng phím tắt', G.fn),
  k('newSheet', 'Shift+F11 / Alt+Shift+F1', ['Shift+F11', 'Alt+Shift+F1'], 'any', 'Chèn trang tính mới', G.fn),
  k('fullScreen', 'Ctrl+Shift+F1', ['Ctrl+Shift+F1'], 'any', 'Bật/tắt toàn màn hình (ẩn Ribbon, thanh công thức, thanh trạng thái)', G.fn),
  k('editBar', 'Ctrl+F2', ['Ctrl+F2'], 'any', 'Sửa ô trên thanh công thức; đang sửa thì chuyển qua lại ô ↔ thanh công thức', G.fn),
  k('goto', 'F5 / Ctrl+G', ['F5', 'Ctrl+G'], 'nav', 'Hộp thoại Đi tới', G.fn),
  k('find', 'Ctrl+F / Shift+F5', ['Ctrl+F', 'Shift+F5'], 'nav', 'Tìm', G.fn),
  k('replace', 'Ctrl+H', ['Ctrl+H'], 'nav', 'Thay thế', G.fn),
  k('findNext', 'Shift+F4', ['Shift+F4'], 'nav', 'Tìm tiếp (lặp lần tìm cuối)', G.fn),
  k('findPrev', 'Ctrl+Shift+F4', ['Ctrl+Shift+F4'], 'nav', 'Tìm ngược lại', G.fn),
  k('closeWin', 'Ctrl+F4', ['Ctrl+F4'], 'any', 'Đóng sổ làm việc', G.fn),
  k('quit', 'Alt+F4', [], 'any', 'Đóng ExcelLike', G.fn),
  k('pane', 'F6 / Shift+F6', ['F6', 'Shift+F6'], 'any', 'Chuyển giữa trang tính, Ribbon, thanh trạng thái/thu phóng', G.fn),
  k('winSwitch', 'Ctrl+F6 / Ctrl+Tab', [], 'any', 'Chuyển giữa các cửa sổ sổ làm việc', G.fn, 'no', 'Mỗi cửa sổ một sổ — chuyển bằng Windows (Alt+Tab)'),
  k('winMin', 'Ctrl+F9', ['Ctrl+F9'], 'any', 'Thu nhỏ cửa sổ', G.fn),
  k('winMax', 'Ctrl+F10', ['Ctrl+F10'], 'any', 'Phóng to / khôi phục cửa sổ', G.fn),
  k('winMoveSize', 'Ctrl+F7 / Ctrl+F8 / Ctrl+F5 (cửa sổ con)', [], 'any', 'Di chuyển/đổi cỡ cửa sổ con', G.fn, 'no', 'Không có cửa sổ con trong ExcelLike'),
  k('saveAs', 'F12', ['F12'], 'any', 'Lưu thành…', G.fn),
  k('openF12', 'Ctrl+F12', ['Ctrl+F12'], 'any', 'Mở…', G.fn),
  k('saveF12', 'Shift+F12', ['Shift+F12'], 'any', 'Lưu', G.fn),

  // ---------- Phím khác ----------
  k('backspace', 'Backspace', ['Backspace'], 'nav', 'Xóa ô hiện hành và bắt đầu sửa', G.other),
  k('spaceCtl', 'Alt+Space', [], 'any', 'Menu điều khiển cửa sổ', G.other, 'no', 'Do Windows xử lý'),
  k('shiftDrag', 'Shift+kéo ô', [], 'nav', 'Di chuyển và chèn ô', G.other, 'later', '0.9'),
  k('dlgTabs', 'Ctrl+Tab / Ctrl+Shift+Tab (trong hộp thoại)', [], 'any', 'Chuyển thẻ trong hộp thoại', G.other),
  k('painter', 'Chổi định dạng (Alt+H, F, P)', [], 'nav', 'Sao chép định dạng (nhấp đúp để dùng nhiều lần, Esc để thôi)', G.other),
  k('dateEdit', 'Ctrl+Delete (khi sửa)', [], 'edit', 'Xóa từ con trỏ tới cuối dòng', G.other),
];

/** Các phím trong phạm vi (đã làm) */
export const keysOk = () => KEYMAP.filter((x) => x.status === 'ok');

const CODE_BASE: Record<string, string> = {
  Semicolon: ';', Equal: '=', Minus: '-', Backquote: '`', Quote: "'", BracketLeft: '[', BracketRight: ']',
  Backslash: '\\', Period: '.', Comma: ',', Slash: '/', Space: 'Space', IntlBackslash: '\\',
};
/** KeyboardEvent → tổ hợp chuẩn hoá (vd 'Ctrl+Shift+4') */
export function comboOf(e: KeyboardEvent): string {
  let base: string;
  const code = e.code || '';
  if (/^Key[A-Z]$/.test(code)) base = code.slice(3);
  else if (/^Digit\d$/.test(code)) base = code.slice(5);
  else if (CODE_BASE[code]) base = CODE_BASE[code];
  else if (/^Numpad(Add|Subtract|Multiply|Divide)$/.test(code)) base = code;
  else if (e.key === ' ') base = 'Space';
  else if (e.key.length === 1) base = e.key.toUpperCase();
  else base = e.key;
  if (base === 'Control' || base === 'Shift' || base === 'Alt' || base === 'Meta') return base;
  // AltGr (bàn phím Đức, Pháp…) báo như Ctrl+Alt: đó là gõ ký tự, không phải phím tắt
  if (e.getModifierState?.('AltGraph') && e.key.length === 1) return e.key;
  const m: string[] = [];
  if (e.ctrlKey || e.metaKey) m.push('Ctrl');
  if (e.altKey) m.push('Alt');
  if (e.shiftKey) m.push('Shift');
  return [...m, base].join('+');
}
/** Tìm định nghĩa phím theo tổ hợp và ngữ cảnh */
export function lookup(combo: string, ctx: 'nav' | 'edit'): KeyDef | undefined {
  return KEYMAP.find((x) => x.status === 'ok' && x.combos.includes(combo) && (x.ctx === 'any' || x.ctx === ctx));
}

/** Phím 'any' do ô đang sửa xử lý (cần biết con trỏ chữ) */
export const EDITOR_ANY = new Set(['today', 'nowTime', 'fAbove', 'vAbove', 'insFn', 'calcAll', 'editBar', 'enterUp', 'enter', 'cancel']);
/** Phím 'any' không xử lý chung khi đang gõ trong ô nhập (để ô nhập tự hoàn tác) */
export const FIELD_KEYS = new Set(['undo', 'redo', 'zoomIn', 'zoomOut', 'nextSheet', 'prevSheet']);
/** Định nghĩa phím cho một pha xử lý */
export function keyFor(combo: string, phase: 'global' | 'nav' | 'edit', inField: boolean): KeyDef | undefined {
  return KEYMAP.find((x) => {
    if (x.status !== 'ok' || !x.combos.includes(combo)) return false;
    if (phase === 'global') return x.ctx === 'any' && !EDITOR_ANY.has(x.id) && !(inField && FIELD_KEYS.has(x.id));
    if (phase === 'nav') return x.ctx === 'nav' || (x.ctx === 'any' && (EDITOR_ANY.has(x.id) || FIELD_KEYS.has(x.id)));
    return x.ctx === 'edit' || (x.ctx === 'any' && EDITOR_ANY.has(x.id));
  });
}
