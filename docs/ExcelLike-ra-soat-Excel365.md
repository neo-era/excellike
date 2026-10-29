# ExcelLike – Rà soát tính năng & phím tắt so với Excel 365 (Bước 0 cho v2)

Ngày tra: 28/10/2026 (cập nhật 29/10/2026 cho bản 0.6). Đối chiếu **ExcelLike 0.4.1 → 0.6.1** với **Excel 365 cho Windows (Current Channel)**.

**Nguồn chính thức**
- Microsoft Support – “Keyboard shortcuts in Excel” (tab Windows): https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f — 298 dòng phím tắt, 13 mục.
- Microsoft OfficeDev – Office Fluent UI Command Identifiers, Microsoft 365 Current Channel (`excelcontrols.xlsx`): danh sách tab → nhóm → lệnh của Ribbon Excel.
- Microsoft Support: Excel status bar options; Available number formats; Change the alignment of text in a cell; Apply or remove cell borders; Apply or remove cell shading; Lock or unlock specific areas of a protected worksheet.

**Ký hiệu**
- Trạng thái: ✅ có · ◐ có một phần · ✗ chưa có · — không làm (cần máy chủ/đám mây Microsoft, ngôn ngữ lập trình nhúng, hoặc không hợp app một người dùng offline).
- Ưu tiên: **P1** = dùng hằng ngày cho dự toán, khối lượng, báo cáo · **P2** = hữu ích · **P3** = ít dùng.

---

## A. Giao diện (so với Excel 365)

| Thành phần Excel 365 | ExcelLike 0.5.0 | Ưu tiên |
|---|---|---|
| Thanh tiêu đề xanh Excel: Quick Access Toolbar (AutoSave, Lưu, Hoàn tác, Làm lại), tên tệp, ô Tìm kiếm (Alt+Q) | ✅ 0.5 (Tự động lưu ghi vào tệp .xlsx đang mở, QAT tùy biến 9 lệnh, tên tệp bấm để đổi, Tìm kiếm Alt+Q không dấu) | P1 |
| Thứ tự tab: Tệp, Trang đầu, Chèn, Bố trí Trang, Công thức, Dữ liệu, Xem lại, Xem, Trợ giúp | ✅ 0.5 | P1 |
| Ribbon: nút lớn/nhỏ xếp 3 tầng trong nhóm, tên nhóm ở đáy, nút mở hộp thoại (↘) góc nhóm, nút thả xuống tách đôi (split button) | ✅ 0.5 (nhóm tự co thành một nút khi cửa sổ hẹp; lệnh chưa làm ghi rõ “có ở bản 0.x”) | P1 |
| Backstage (Tệp) toàn màn hình: Trang đầu, Mới, Mở, Thông tin, Lưu, Lưu thành, In (xem trước bên phải), Xuất, Đóng, Tùy chọn | ✅ 0.5 (+ Tài khoản: Chủ đề Office) | P1 |
| Hộp Tên + nút ✕ ✓ fx + thanh công thức mở rộng (Ctrl+Shift+U) | ✅ 0.5 | P1 |
| Lưới: màu đường lưới, tiêu đề cột/hàng, ô chọn viền xanh đậm + chấm điền, vùng chọn tô mờ, tiêu đề tô đậm theo vùng chọn | ✅ 0.5 (ô mặc định 64×20 px như Excel; Xem trước Ngắt Trang) | P1 |
| Thẻ sheet: nút ◀ ▶, ⊕, màu thẻ, chuột phải (Chèn, Xóa, Đổi tên, Di chuyển/Sao chép, Màu thẻ, Ẩn/Hiện, Bảo vệ) | ✅ 0.5 (Bảo vệ → 0.9) | P1 |
| Thanh trạng thái: Sẵn sàng/Nhập/Sửa/Trỏ, Trung bình/Đếm/Tổng (+ tùy chọn Min/Max/Đếm số), nút Thường/Bố trí trang/Ngắt trang, thanh trượt thu phóng; chuột phải để chọn mục | ✅ 0.5 (Bố trí trang → 0.8) | P1 |
| Menu chuột phải ô: thanh mini định dạng + Cắt/Sao chép/Tùy chọn dán/Dán đặc biệt/Chèn…/Xóa…/Xóa nội dung/Lọc ▸/Sắp xếp ▸/Bình luận/Ghi chú/Định dạng ô…/Chọn từ danh sách/Đặt tên/Liên kết | ✅ 0.5 (Dán đặc biệt, Bình luận/Ghi chú, Định dạng ô, Đặt tên, Liên kết: mục có, lệnh ở 0.6/0.7/0.9) | P1 |
| KeyTips (Alt/F10 hiện chữ trên nút, Alt+H, Alt+N…) | ✅ 0.5 | P1 |
| Chủ đề Office: Nhiều màu, Xám đậm, Đen, Trắng; chế độ tối cho lưới | ✅ 0.5 (Nhiều màu, Xám Đậm, Đen, Trắng, Theo hệ thống; bản Windows đổi màu nút cửa sổ theo) | P2 |
| Ribbon thu gọn/đơn giản hóa (Ctrl+F1) | ✅ | – |
| Phông giao diện Segoe UI, mặc định ô Aptos/Calibri 11 | ✅ 0.5 | P1 |

## B. Tính năng theo tab

### Tệp (Backstage)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Mới, Mở, Lưu, Lưu thành, gần đây, tự lưu nháp | ✅ | – |
| Mẫu sổ tính (dự toán, khối lượng, thanh toán) | ✗ | P2 |
| In: xem trước ngay trong Backstage, in vùng chọn/cả sổ, số bản, khổ, lề, co giãn (vừa 1 trang / vừa cột / vừa hàng) | ✅ 0.5 (xem trước từng trang ◀ 1/N ▶, Hiện lề; bản Windows chọn máy in và in thẳng, không qua hộp thoại hệ thống) | P1 |
| Xuất PDF | ✅ 0.5 (bản Windows ghi tệp .pdf thật) | P1 |
| Thông tin: bảo vệ sổ bằng mật khẩu, mở chỉ đọc, kiểm tra sổ | ✗ | P3 |
| Đóng (Ctrl+W), nhiều sổ mở cùng lúc, chuyển cửa sổ (Ctrl+F6, Ctrl+Tab) | ◐ 0.6 (Ctrl+W/Ctrl+F4 đóng sổ, hỏi lưu; nhiều sổ cùng lúc ✗) | P2 |
| Mở .xls (97–2003), .xlsb, .ods; lưu .xlsm giữ macro | ✗ | P2 |
| Tùy chọn (Options): chế độ tính, tính lặp, hướng Enter, phông mặc định… | ◐ 0.5 (Tính lặp + số lần/sai số, hướng Enter) | P2 |
| Chia sẻ, OneDrive, Phiên bản, Power BI | — | — |

### Trang đầu (Home)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Cắt, Sao chép, Dán | ✅ | – |
| Dán đặc biệt (Ctrl+Alt+V): giá trị, công thức, định dạng, độ rộng cột, chuyển vị, phép tính, bỏ ô trống; nút Tùy chọn dán sau khi dán | ✅ 0.6 (12 kiểu dán + 5 phép tính + bỏ ô trống + chuyển vị + dán liên kết, phím chữ A F V T C N H X W R U; menu Dán và hàng Tùy chọn Dán; Ctrl+Shift+V dán giá trị) | P1 |
| Chổi định dạng (nhấp/nhấp đúp) | ✅ 0.6 (nhấp đúp = dùng nhiều lần, Esc để thôi) | P1 |
| Phông, cỡ, Đậm/Nghiêng/Gạch chân, gạch ngang (Ctrl+5), màu chữ, tô nền | ✅ | – |
| Tăng/Giảm cỡ chữ, gạch chân kép, chỉ số trên/dưới | ✅ 0.6 | P2 |
| Viền: các kiểu sẵn | ✅ | – |
| Viền: vẽ viền, màu/kiểu nét, Thêm viền… (hộp thoại) | ◐ 0.6 (hộp thoại Viền: 7 kiểu nét, màu, 8 cạnh gồm đường chéo; Vẽ viền bằng chuột ✗) | P2 |
| Bảng màu đầy đủ (màu chủ đề + sắc độ, Màu khác…) | ◐ | P1 |
| Căn trên/giữa/dưới, trái/giữa/phải, xuống dòng, gộp ô (3 kiểu), bỏ gộp | ✅ | – |
| Thụt lề (tăng/giảm), xoay chữ (góc, dọc), Căn giữa qua vùng chọn, Co chữ vừa ô | ✅ 0.6 (cả Điền, Căn đều, Phân phối) | P1 (thụt lề), P2 |
| Định dạng số: danh sách nhanh, Kế toán, %, dấu phân cách, tăng/giảm số lẻ | ✅ (Kế toán ◐) | – |
| Hộp thoại Định dạng ô (Ctrl+1): Số (Tùy chỉnh), Căn chỉnh, Phông, Viền, Tô, Bảo vệ | ✅ 0.6 (Tô: màu nền; mẫu/chuyển màu ✗) | P1 |
| Định dạng có điều kiện: tô ô theo quy tắc, trùng lặp, Top/Bottom, trên/dưới TB, thanh dữ liệu, thang màu, quản lý | ✅ | – |
| … Bộ biểu tượng, “Ngày xảy ra”, Quy tắc mới bằng công thức | ✗ | P2 |
| Định dạng như Bảng (Table: tên bảng, hàng tổng, dải màu, tham chiếu có cấu trúc `Bảng1[Cột]`) | ✗ | P1 |
| Kiểu ô (Cell Styles) | ✗ | P2 |
| Chèn/Xóa hàng, cột | ✅ | – |
| Chèn/Xóa ô (dồn xuống/phải, lên/trái) | ✅ 0.5 | P1 |
| Định dạng: chiều cao hàng/độ rộng cột theo số, Tự khớp, Ẩn/Hiện hàng-cột-sheet, Di chuyển/Sao chép sheet, Màu thẻ, Bảo vệ sheet, Khóa ô | ✅ 0.5 (trừ Bảo vệ sheet, Khóa ô → 0.9) | P1 |
| Tổng tự động ▸ Tổng/Trung bình/Đếm/Max/Min | ✅ 0.5 | P1 |
| Điền ▸ Xuống/Phải/Lên/Trái/Chuỗi…/Flash Fill | ✅ 0.6 (Chuỗi: tuyến tính, tăng trưởng, ngày/ngày làm việc/tháng/năm, bước, giá trị dừng; Điền Nhanh Ctrl+E) | P2 |
| Xóa ▸ Tất cả/Định dạng/Nội dung/Ghi chú/Liên kết | ◐ 0.6 (Tất cả, Định dạng, Nội dung, Siêu kết nối; Ghi chú → 0.7) | P1 |
| Sắp xếp & Lọc | ✅ | – |
| Lọc nâng cao: theo màu, theo số (lớn hơn, top 10), theo ngày (tháng này, quý…) | ◐ (giá trị, chữ chứa) | P2 |
| Tìm & Chọn: Tìm (Ctrl+F), Thay thế (Ctrl+H), Đi tới (F5/Ctrl+G), Đi tới Đặc biệt (ô trống, công thức, hằng số…) | ✅ 0.6 (hộp Tìm/Thay không chặn trang tính: ký tự đại diện * ? ~, khớp hoa/thường, toàn ô, trong trang/sổ, theo hàng/cột, công thức/giá trị, Tìm Tất cả có danh sách; Đi tới Đặc biệt 14 loại; Tìm theo Định dạng ✗) | **P1** |
| Phân tích dữ liệu, Copilot, Nhãn độ nhạy, Phần bổ trợ | — | — |

### Chèn (Insert)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| PivotTable, Bảng (Table) | ◐ (PivotTable ✅, Bảng ✗) | P1 |
| PivotTable đề xuất, PivotChart | ✗ | P3 |
| Biểu đồ: Cột/Thanh, Đường/Vùng, Tròn/Vành khuyên | ✅ | – |
| Biểu đồ: Phân tán (XY), Kết hợp (cột + đường, trục phụ), Thác nước, Radar, Histogram, Pareto, Hộp | ✗ | P2 (XY, Kết hợp), P3 |
| Biểu đồ đề xuất; tab Thiết kế/Định dạng biểu đồ (phần tử biểu đồ, nhãn, trục, đường xu hướng, kiểu, màu) | ◐ (hộp thoại riêng) | P2 |
| Hình ảnh (đặt trên ô / trong ô), Hình vẽ, Biểu tượng, SmartArt, Ảnh chụp màn hình | ✗ | P2 (ảnh), P3 |
| Biểu đồ thu nhỏ trong ô (Sparklines) | ✗ | P3 |
| Slicer, Timeline | ✗ | P3 |
| Liên kết (Ctrl+K) | ✅ 0.6 (web, vị trí trong sổ, email; nhấp để mở; đọc/ghi .xlsx) | P2 |
| Bình luận | ✗ (xem Xem lại) | P1 |
| Hộp kiểm trong ô (Checkbox) | ✗ | P3 |
| Hộp văn bản, Đầu trang & Chân trang, WordArt, Chữ ký, Đối tượng | ✗ (đầu/chân trang P1 – xem Bố trí trang) | P3 |
| Phương trình, Ký hiệu | ✗ | P3 |
| Bản đồ 3D, Biểu mẫu Microsoft | — | — |

### Bố trí Trang (Page Layout)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Lề, Hướng, Khổ giấy | ✅ 0.5 (8 khổ giấy, lề tùy chỉnh) | P1 |
| Vùng in (đặt/xóa/thêm), Ngắt trang (chèn/xóa/đặt lại), In tiêu đề (hàng **và** cột lặp lại) | ✅ 0.5 | P1 |
| Co giãn vừa trang: Rộng _ trang × Cao _ trang, Tỷ lệ % | ✅ 0.5 | P1 |
| Đầu trang/Chân trang: số trang “Trang 1/5”, tên sheet, ngày, chữ tự gõ 3 phần trái/giữa/phải | ✅ 0.5 (+ số trang đầu tiên, căn giữa trang, thứ tự trang) | **P1** (hồ sơ dự toán) |
| Đường lưới / Tiêu đề hàng-cột: Xem, In | ✅ 0.5 | P2 |
| Chủ đề (màu, phông), Ảnh nền, Sắp xếp đối tượng | ✗ | P3 |
| Dạng xem Bố trí trang, Xem trước ngắt trang (kéo đường ngắt trang) | ◐ 0.5 (Xem trước Ngắt Trang có số trang + đường ngắt; kéo đường ngắt và Bố trí trang → 0.8) | P1 |

### Công thức (Formulas)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Chèn hàm, Tổng tự động, thư viện hàm theo nhóm | ◐ (hộp chọn hàm có tìm kiếm; chưa có các nút nhóm Tài chính/Logic/…) | P2 |
| Gợi ý hàm khi gõ + chú thích tham số | ✅ | – |
| Hộp thoại Đối số hàm (Ctrl+A sau tên hàm) | ✅ 0.6 (kèm Ctrl+Shift+A chèn tên đối số) | P2 |
| ~420 hàm Excel; mảng động (FILTER, SORT, UNIQUE, SEQUENCE tràn ô), LET, LAMBDA, XLOOKUP | ◐ (XLOOKUP ✅; mảng động/LET/LAMBDA ✗) | P1 (FILTER, SORT, UNIQUE), P2 |
| Hằng mảng nhiều cột `{1,2;3,4}` | ✗ (đang sai khi đọc tệp) | P1 |
| Tên vùng: Đặt tên, Trình quản lý tên, Tạo từ vùng chọn, Dán tên (F3) | ✗ (tên trong tệp chưa đọc) | **P1** |
| Truy vết: Ô tiền lệ/phụ thuộc (mũi tên), Hiện công thức (Ctrl+`), Kiểm tra lỗi, Đánh giá công thức, Cửa sổ theo dõi | ✗ | P1 (tiền lệ/phụ thuộc, Ctrl+`), P2 |
| Tham chiếu vòng: bật tính lặp | ✅ v0.4.1 | – |
| Chế độ tính: Tự động / Thủ công (F9, Shift+F9) | ◐ (luôn tự động) | P2 |
| Liên kết ngoài: giữ giá trị lưu sẵn, Sửa liên kết, cập nhật từ tệp khác | ◐ (giữ giá trị ✅; cập nhật ✗) | P2 |
| Python trong Excel | — | — |

### Dữ liệu (Data)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Mở CSV/Text | ✅ | – |
| Sắp xếp, Lọc, Xóa lọc, Áp dụng lại | ✅ | – |
| Lọc nâng cao (vùng điều kiện) | ✗ | P3 |
| Văn bản thành cột, Xóa trùng lặp | ✗ | **P1** |
| Kiểm tra dữ liệu, khoanh dữ liệu không hợp lệ | ◐ (khoanh ✗) | P2 |
| Hợp nhất (Consolidate) | ✗ | P3 |
| What-If: Goal Seek, Scenario, Data Table; Dự báo | ✗ | P2 (Goal Seek), P3 |
| Nhóm/Bỏ nhóm hàng-cột (dấu +/−), Tổng phụ (Subtotal) | ✗ | **P1** (dự toán nhiều cấp) |
| Lấy dữ liệu (Power Query), Kết nối, Mô hình dữ liệu, Kiểu dữ liệu Cổ phiếu/Địa lý | — | — |

### Xem lại (Review) — tab chưa có
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Ghi chú ô (Shift+F2), tam giác đỏ góc ô, hiện khi rê chuột; đọc/ghi từ .xlsx | ✗ (hiện mất khi lưu đè) | **P1** |
| Bình luận dạng chuỗi (Ctrl+Shift+F2), trả lời, đánh dấu xong | ✗ | P2 |
| Kiểm tra chính tả (F7), Từ đồng nghĩa, Thống kê sổ (Ctrl+Shift+G) | ✗ | P3 (F7), P2 (Thống kê) |
| Bảo vệ sheet / sổ (khóa ô, ẩn công thức, mật khẩu), Cho phép sửa vùng | ✗ | P2 |
| Kiểm tra trợ năng, Dịch, Hiện thay đổi | — | — |

### Xem (View)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Thường / Xem trước ngắt trang / Bố trí trang | ◐ 0.5 (Thường, Xem trước Ngắt Trang) | P1 |
| Hiện: Thanh công thức, Đường lưới, Tiêu đề | ✅ 0.5 | P2 |
| Thu phóng, 100%, Thu phóng theo vùng chọn | ✅ 0.5 | P2 |
| Cố định ngăn: tại ô, hàng đầu, cột đầu | ✅ | – |
| Tách cửa sổ (Split), Cửa sổ mới, Xếp cạnh nhau | ✗ | P2 (Split) |
| Ẩn/hiện sổ, Dạng xem tùy chỉnh, Sheet View | ✗ | P3 |
| Macro | — | — |

### Tab ngữ cảnh
| Tab | ExcelLike | Ưu tiên |
|---|---|---|
| Thiết kế Bảng (Table Design) | ✗ | P1 (cùng tính năng Bảng) |
| PivotTable: Phân tích / Thiết kế (bố cục Gọn/Dàn ý/Bảng, tổng phụ, tổng cộng, kiểu, nhóm theo ngày, trường tính toán, % của tổng) | ◐ (ngăn trường riêng; bố cục Bảng) | P2 |
| Biểu đồ: Thiết kế / Định dạng | ◐ (hộp thoại riêng) | P2 |
| Định dạng Ảnh/Hình | ✗ | P3 |

### Hộp thoại Định dạng ô (Ctrl+1)
| Tab | ExcelLike | Ưu tiên |
|---|---|---|
| Số: Chung, Số, Tiền tệ, Kế toán, Ngày, Giờ, %, Phân số, Khoa học, Văn bản, Đặc biệt, **Tùy chỉnh** (gõ mã, xem mẫu) | ✅ 0.6 | **P1** |
| Căn chỉnh: ngang (gồm Điền, Căn đều, Giữa qua vùng chọn, Phân phối), dọc, thụt lề, xuống dòng, co vừa ô, gộp, góc xoay | ✅ 0.6 (dọc: Trên/Giữa/Dưới) | P1 |
| Phông: kiểu, gạch chân (đơn/kép/kế toán), gạch ngang, chỉ số trên/dưới | ✅ 0.6 (kế toán = đơn/kép) | P1 |
| Viền: kiểu nét, màu, bấm từng cạnh, đường chéo | ✅ 0.6 | P1 |
| Tô: màu, mẫu (pattern), chuyển màu | ◐ (màu) | P2 |
| Bảo vệ: Khóa, Ẩn công thức | ◐ 0.6 (thuộc tính lưu vào .xlsx; Bảo vệ trang tính → 0.9) | P2 |

### Tương thích tệp .xlsx (đọc → lưu không mất)
| Nội dung | ExcelLike | Ưu tiên |
|---|---|---|
| Giá trị, công thức, định dạng, gộp ô, cố định ngăn, lọc, CF, DV, biểu đồ cơ bản, thiết lập in, tính lặp, liên kết ngoài | ✅ | – |
| Tên vùng (definedNames), vùng in nhiều khối | ✗ | **P1** |
| Ghi chú/bình luận, hình ảnh, hình vẽ, hộp chữ | ✗ (báo trước, mất khi lưu đè) | P1 (ghi chú, ảnh), P2 |
| Bảng (Table) có tham chiếu có cấu trúc | ✗ | P1 |
| Nhóm hàng/cột (outline), hàng/cột ẩn | ◐ (ẩn ✅ 0.5 có Ẩn/Bỏ ẩn, outline ✗) | P1 |
| Đầu trang/chân trang, ngắt trang thủ công | ✅ 0.5 | P1 |
| Bảo vệ sheet, màu thẻ sheet, sheet ẩn | ◐ 0.5 (màu thẻ, sheet ẩn, cột ẩn giữ độ rộng; bảo vệ → 0.9) | P2 |
| PivotTable “sống” (pivotCache) | ✗ (ghi thành bảng giá trị) | P3 |
| Macro VBA (.xlsm): giữ nguyên khi lưu, không chạy | ✗ | P2 |

## C. Phím tắt (cập nhật 0.6 — tra lại trang Microsoft ngày 29/10/2026)

Nguồn duy nhất trong mã: `src/ui/keymap.ts` — dùng cho bộ điều phối phím, bảng phím tắt F1 trong app, bảng dưới đây (`node --experimental-strip-types tools/keytable.mjs`) và test e2e (`e2e/v06.mjs` bấm thật từng phím).

Không làm (giữ nguyên lý do): Power Pivot (27 dòng), Copilot (9), Alt+F8/F11/F12 (macro, VBA, Power Query), Ctrl+F6/Ctrl+Tab đổi cửa sổ sổ khác (mỗi cửa sổ một sổ), Ctrl+F5/F7/F8 cửa sổ con kiểu cũ, Alt+Space (Windows xử lý).
Có ở bản sau (cần tính năng lớn trước): Shift+F2 ghi chú, Ctrl+Shift+F2 chú thích, Ctrl+8 dàn bài, Ctrl+L/Ctrl+T Bảng, F3 / Alt+M,M,D tên vùng → 0.7; F7 chính tả, Alt+Shift+F10 kiểm tra lỗi, Shift+kéo chèn ô → 0.9.

Hành vi đã làm giống Excel: Ctrl+A chọn vùng dữ liệu rồi mới cả trang; Ctrl+Home về ô đầu vùng không cố định; F4 khi sửa đổi $ tham chiếu, khi không sửa thì lặp lệnh cuối (Ctrl+Y cũng lặp khi không có gì để làm lại); Tab…Tab rồi Enter về cột bắt đầu; Enter/Tab đi vòng trong vùng chọn và sang khối kế tiếp; Ctrl+Enter điền mọi ô đang chọn (công thức dời tham chiếu); Delete giữ định dạng; Backspace xóa rồi sửa; Ctrl+D/R chép cả định dạng; Ctrl+Shift++ / Ctrl+− chọn cả hàng/cột thì chèn/xóa luôn; Ctrl+; Ctrl+Shift+: Ctrl+' Ctrl+Shift+" chèn tại con trỏ khi đang sửa; nhóm trang tính áp nhập liệu/định dạng/xóa cho mọi trang trong nhóm.

Tổng 180 mục: ✅ 167 · ◐ 9 · — 4

**Thường dùng**

| Phím | Hành động | Trạng thái |
|---|---|---|
| Ctrl+W | Đóng sổ làm việc | ✅ |
| Ctrl+O | Mở sổ làm việc | ✅ |
| Ctrl+N | Sổ làm việc mới | ✅ |
| Ctrl+S | Lưu sổ làm việc | ✅ |
| Ctrl+P | In (xem trước trong Tệp ▸ In) | ✅ |
| Ctrl+C | Sao chép | ✅ |
| Ctrl+V | Dán | ✅ |
| Ctrl+X | Cắt | ✅ |
| Ctrl+Z | Hoàn tác | ✅ |
| Ctrl+Y | Làm lại; nếu không có gì để làm lại thì lặp lại lệnh cuối | ✅ |
| Delete | Xóa nội dung ô (giữ định dạng) | ✅ |
| Alt+H, H | Chọn màu tô (KeyTips) | ✅ |
| Alt+H, A, C | Căn giữa (KeyTips) | ✅ |
| Alt+H, B | Thêm viền (KeyTips) | ✅ |
| Alt+H, D, C | Xóa cột (KeyTips) | ✅ |
| Ctrl+B | Đậm / bỏ đậm | ✅ |
| Shift+F10 | Mở menu chuột phải | ✅ |
| Ctrl+9 | Ẩn hàng đang chọn | ✅ |
| Ctrl+0 | Ẩn cột đang chọn | ✅ |
| Ctrl+Shift+( | Bỏ ẩn hàng trong vùng chọn | ✅ |
| Ctrl+Shift+) | Bỏ ẩn cột trong vùng chọn | ✅ |

**Ribbon & KeyTips**

| Phím | Hành động | Trạng thái |
|---|---|---|
| Alt+F | Mở thẻ Tệp | ✅ |
| Alt+H | Mở thẻ Trang đầu | ✅ |
| Alt+N | Mở thẻ Chèn | ✅ |
| Alt+P | Mở thẻ Bố trí Trang | ✅ |
| Alt+M | Mở thẻ Công thức | ✅ |
| Alt+A | Mở thẻ Dữ liệu | ✅ |
| Alt+R | Mở thẻ Xem lại | ✅ |
| Alt+W | Mở thẻ Xem | ✅ |
| Alt+Q | Tới ô Tìm kiếm lệnh | ✅ |
| Alt hoặc F10 | Hiện/ẩn KeyTips | ✅ |
| Tab / Shift+Tab | Chuyển giữa các lệnh trên Ribbon (khi Ribbon đang nhận phím) | ✅ |
| Phím mũi tên | Di chuyển giữa các mục trên Ribbon | ✅ |
| Ctrl+Shift+F10 | Hiện chú thích của lệnh đang chọn trên Ribbon | ✅ |
| Space hoặc Enter | Thực hiện lệnh đang chọn trên Ribbon | ✅ |
| Alt+↓ / ↓ | Mở danh sách/menu của lệnh đang chọn | ✅ |
| Ctrl+F1 | Thu gọn / hiện Ribbon | ✅ |
| Ctrl+← / Ctrl+→ | Nhảy giữa các nhóm lệnh trên Ribbon | ✅ |

**Di chuyển trong ô**

| Phím | Hành động | Trạng thái |
|---|---|---|
| ↑ ↓ ← → | Di chuyển một ô | ✅ |
| Tab / Shift+Tab | Sang ô phải / trái (đi vòng trong vùng chọn) | ✅ |
| Ctrl+mũi tên | Tới mép vùng dữ liệu | ✅ |
| End, mũi tên | Chế độ End: End rồi mũi tên = tới ô có dữ liệu kế tiếp | ✅ |
| Ctrl+End | Tới ô cuối cùng có dùng | ✅ |
| Ctrl+Shift+End | Mở rộng vùng chọn tới ô cuối cùng có dùng | ✅ |
| Home | Về đầu hàng | ✅ |
| Ctrl+Home | Về đầu trang tính (sau ngăn cố định) | ✅ |
| Page Down | Xuống một màn hình | ✅ |
| Page Up | Lên một màn hình | ✅ |
| Alt+Page Down | Sang phải một màn hình | ✅ |
| Alt+Page Up | Sang trái một màn hình | ✅ |
| Ctrl+Page Down | Sang trang tính sau | ✅ |
| Ctrl+Page Up | Về trang tính trước | ✅ |
| Alt+↓ | Mở danh sách chọn (kiểm tra dữ liệu, nút lọc, hoặc chọn từ các giá trị trong cột) | ✅ |
| Ctrl+Alt+5, Tab | Duyệt qua các biểu đồ/đối tượng nổi (Esc để thoát) | ✅ |
| Ctrl+Shift+cuộn chuột | Cuộn ngang | ✅ |
| Ctrl+Alt+= | Phóng to | ✅ |
| Ctrl+Alt+− | Thu nhỏ | ✅ |

**Định dạng ô**

| Phím | Hành động | Trạng thái |
|---|---|---|
| Ctrl+1 | Hộp thoại Định dạng Ô | ✅ |
| Ctrl+Shift+F | Định dạng Ô ▸ thẻ Phông | ✅ |
| F2 | Sửa ô hiện hành | ✅ |
| Shift+F2 | Chèn/sửa ghi chú ô | ◐ sau (0.7 (Ghi chú ô: đọc/ghi .xlsx)) |
| Ctrl+Shift+F2 | Chèn/trả lời chú thích dạng chuỗi | ◐ sau (0.7 (Chú thích)) |
| Ctrl+Shift++ | Hộp thoại Chèn ô (chọn cả hàng/cột thì chèn luôn) | ✅ |
| Ctrl+− | Hộp thoại Xóa ô (chọn cả hàng/cột thì xóa luôn) | ✅ |
| Ctrl+Shift+: | Nhập giờ hiện tại | ✅ |
| Ctrl+; | Nhập ngày hiện tại | ✅ |
| Ctrl+` | Hiện công thức / giá trị | ✅ |
| Ctrl+' | Chép công thức ô phía trên vào ô hiện hành | ✅ |
| Ctrl+Alt+V | Hộp thoại Dán Đặc biệt | ✅ |
| Ctrl+Shift+V | Dán chỉ giá trị (Excel 365) | ✅ |
| Ctrl+I | Nghiêng / bỏ nghiêng | ✅ |
| Ctrl+U | Gạch chân / bỏ gạch chân | ✅ |
| Ctrl+5 | Gạch ngang / bỏ gạch ngang | ✅ |
| Ctrl+6 | Hiện / ẩn / hiện chỗ giữ cho đối tượng (biểu đồ) | ✅ |
| Ctrl+Shift+& | Kẻ viền ngoài | ✅ |
| Ctrl+Shift+_ | Bỏ viền ngoài | ✅ |
| Ctrl+8 | Hiện/ẩn ký hiệu dàn bài (nhóm hàng/cột) | ◐ sau (0.7 (Nhóm hàng/cột)) |
| Ctrl+D | Điền xuống | ✅ |
| Ctrl+R | Điền sang phải | ✅ |
| Ctrl+Shift+~ | Định dạng Chung | ✅ |
| Ctrl+Shift+$ | Định dạng Tiền tệ (2 số lẻ, số âm trong ngoặc) | ✅ |
| Ctrl+Shift+% | Định dạng Phần trăm, không số lẻ | ✅ |
| Ctrl+Shift+^ | Định dạng Khoa học, 2 số lẻ | ✅ |
| Ctrl+Shift+# | Định dạng Ngày (ngày, tháng, năm) | ✅ |
| Ctrl+Shift+@ | Định dạng Giờ (giờ, phút, SA/CH) | ✅ |
| Ctrl+Shift+! | Định dạng Số: 2 số lẻ, dấu phân cách hàng nghìn | ✅ |
| Ctrl+K | Chèn / sửa liên kết | ✅ |
| F7 | Kiểm tra chính tả | ◐ sau (0.9 (cần từ điển tiếng Việt)) |
| Ctrl+Q | Phân tích Nhanh cho vùng chọn | ✅ |
| Ctrl+L hoặc Ctrl+T | Tạo Bảng | ◐ sau (0.7 (Bảng + tham chiếu có cấu trúc)) |
| Ctrl+Shift+G | Thống kê Sổ làm việc | ✅ |

**Hộp Dán Đặc biệt**

| Phím | Hành động | Trạng thái |
|---|---|---|
| A | Dán tất cả | ✅ |
| F | Chỉ công thức | ✅ |
| V | Chỉ giá trị | ✅ |
| T | Chỉ định dạng | ✅ |
| C | Chỉ chú thích & ghi chú | ✅ |
| N | Chỉ kiểm tra dữ liệu | ✅ |
| H | Tất cả, dùng chủ đề nguồn | ✅ |
| X | Tất cả trừ viền | ✅ |
| W | Độ rộng cột | ✅ |
| R | Công thức và định dạng số | ✅ |
| U | Giá trị và định dạng số | ✅ |

**Chọn & thao tác**

| Phím | Hành động | Trạng thái |
|---|---|---|
| Ctrl+A / Ctrl+Shift+Space | Chọn vùng dữ liệu hiện tại; bấm lần nữa chọn cả trang tính | ✅ |
| Ctrl+Shift+Page Down | Chọn thêm trang tính sau (nhóm trang tính) | ✅ |
| Ctrl+Shift+Page Up | Chọn thêm trang tính trước (nhóm trang tính) | ✅ |
| Shift+mũi tên | Mở rộng vùng chọn một ô | ✅ |
| Ctrl+Shift+mũi tên | Mở rộng vùng chọn tới ô có dữ liệu cuối | ✅ |
| F8 | Bật/tắt chế độ Mở rộng Vùng chọn | ✅ |
| Shift+F8 | Thêm vùng không liền kề vào vùng chọn | ✅ |
| Ctrl+nhấp chuột | Thêm ô/vùng không liền kề vào vùng chọn | ✅ |
| Alt+Enter | Xuống dòng trong ô | ✅ |
| Ctrl+Enter | Điền nội dung đang nhập vào mọi ô đang chọn | ✅ |
| Shift+Enter | Hoàn tất và lên ô trên | ✅ |
| Ctrl+Space | Chọn cả cột | ✅ |
| Shift+Space | Chọn cả hàng | ✅ |
| Ctrl+Shift+Home | Mở rộng vùng chọn tới đầu trang tính | ✅ |
| Shift+Home | Mở rộng vùng chọn về đầu hàng | ✅ |
| Shift+Page Down / Up | Mở rộng vùng chọn một màn hình | ✅ |
| Ctrl+Shift+* | Chọn vùng dữ liệu quanh ô hiện hành | ✅ |
| F4 (khi không sửa ô) | Lặp lại lệnh/thao tác cuối | ✅ |
| Alt+; | Chỉ chọn các ô đang hiện | ✅ |
| Ctrl+[ | Chọn ô tiền lệ trực tiếp | ✅ |
| Ctrl+Shift+{ | Chọn mọi ô tiền lệ | ✅ |
| Ctrl+] | Chọn ô phụ thuộc trực tiếp | ✅ |
| Ctrl+Shift+} | Chọn mọi ô phụ thuộc | ✅ |
| Ctrl+\ | Chọn ô khác biệt trong hàng (so với cột của ô hiện hành) | ✅ |
| Ctrl+Shift+\| | Chọn ô khác biệt trong cột (so với hàng của ô hiện hành) | ✅ |
| Ctrl+Shift+O | Chọn các ô có ghi chú | ✅ |
| Ctrl+. | Ô hiện hành chuyển tới góc kế tiếp của vùng chọn | ✅ |
| Shift+Backspace | Thu vùng chọn về ô hiện hành | ✅ |
| Ctrl+Backspace | Cuộn tới ô hiện hành | ✅ |

**Dữ liệu, hàm, thanh công thức**

| Phím | Hành động | Trạng thái |
|---|---|---|
| Ctrl+Alt+P | Bật/tắt chú thích hàm khi gõ công thức | ✅ |
| Ctrl+Shift+U | Mở rộng / thu gọn thanh công thức | ✅ |
| Esc | Hủy nhập; đóng menu/hộp thoại; bỏ viền sao chép | ✅ |
| Enter | Hoàn tất và xuống ô dưới (Tab rồi Enter: về cột bắt đầu) | ✅ |
| Ctrl+Shift+Enter | Nhập công thức mảng (Excel 365: như Enter, mảng tự tràn) | ✅ |
| F9 | Tính lại mọi trang tính (khi sửa: tính phần công thức đang bôi đen) | ✅ |
| Shift+F9 | Tính lại trang tính hiện hành | ✅ |
| Ctrl+Alt+F9 | Tính lại toàn bộ, kể cả ô không đổi | ✅ |
| Alt+Shift+F10 | Menu nút Kiểm tra Lỗi | ◐ sau (0.9 (Kiểm tra Lỗi)) |
| Ctrl+A (sau tên hàm) | Hộp thoại Đối số Hàm | ✅ |
| Ctrl+Shift+A (sau tên hàm) | Chèn tên đối số và ngoặc | ✅ |
| Alt+= | Tổng Tự động | ✅ |
| Ctrl+E | Điền Nhanh (Flash Fill) | ✅ |
| F4 (khi sửa công thức) | Đổi tham chiếu A1 → $A$1 → A$1 → $A1 | ✅ |
| Shift+F3 | Chèn Hàm | ✅ |
| Ctrl+Shift+" | Chép giá trị ô phía trên vào ô hiện hành | ✅ |
| Alt+F1 | Tạo biểu đồ nhúng từ vùng hiện tại | ✅ |
| F11 | Tạo biểu đồ trên trang tính riêng | ✅ |
| Alt+M, M, D | Xác định tên | ◐ sau (0.7 (Tên vùng)) |
| F3 | Dán tên đã xác định | ◐ sau (0.7 (Tên vùng)) |
| Alt+F8 / Alt+F11 / Alt+F12 | Macro, VBA, Power Query | — không làm (Không chạy macro/VBA/Power Query) |
| Ctrl+Shift+L | Bật/tắt Lọc | ✅ |
| Ctrl+F5 | Làm mới dữ liệu trang tính (PivotTable trên trang) | ✅ |
| Ctrl+Alt+F5 | Làm mới tất cả (mọi PivotTable) | ✅ |

**Phím chức năng**

| Phím | Hành động | Trạng thái |
|---|---|---|
| F1 | Trợ giúp: bảng phím tắt | ✅ |
| Shift+F11 / Alt+Shift+F1 | Chèn trang tính mới | ✅ |
| Ctrl+Shift+F1 | Bật/tắt toàn màn hình (ẩn Ribbon, thanh công thức, thanh trạng thái) | ✅ |
| Ctrl+F2 | Sửa ô trên thanh công thức; đang sửa thì chuyển qua lại ô ↔ thanh công thức | ✅ |
| F5 / Ctrl+G | Hộp thoại Đi tới | ✅ |
| Ctrl+F / Shift+F5 | Tìm | ✅ |
| Ctrl+H | Thay thế | ✅ |
| Shift+F4 | Tìm tiếp (lặp lần tìm cuối) | ✅ |
| Ctrl+Shift+F4 | Tìm ngược lại | ✅ |
| Ctrl+F4 | Đóng sổ làm việc | ✅ |
| Alt+F4 | Đóng ExcelLike | ✅ |
| F6 / Shift+F6 | Chuyển giữa trang tính, Ribbon, thanh trạng thái/thu phóng | ✅ |
| Ctrl+F6 / Ctrl+Tab | Chuyển giữa các cửa sổ sổ làm việc | — không làm (Mỗi cửa sổ một sổ — chuyển bằng Windows (Alt+Tab)) |
| Ctrl+F9 | Thu nhỏ cửa sổ | ✅ |
| Ctrl+F10 | Phóng to / khôi phục cửa sổ | ✅ |
| Ctrl+F7 / Ctrl+F8 / Ctrl+F5 (cửa sổ con) | Di chuyển/đổi cỡ cửa sổ con | — không làm (Không có cửa sổ con trong ExcelLike) |
| F12 | Lưu thành… | ✅ |
| Ctrl+F12 | Mở… | ✅ |
| Shift+F12 | Lưu | ✅ |

**Phím khác**

| Phím | Hành động | Trạng thái |
|---|---|---|
| Backspace | Xóa ô hiện hành và bắt đầu sửa | ✅ |
| Alt+Space | Menu điều khiển cửa sổ | — không làm (Do Windows xử lý) |
| Shift+kéo ô | Di chuyển và chèn ô | ◐ sau (0.9) |
| Ctrl+Tab / Ctrl+Shift+Tab (trong hộp thoại) | Chuyển thẻ trong hộp thoại | ✅ |
| Chổi định dạng (Alt+H, F, P) | Sao chép định dạng (nhấp đúp để dùng nhiều lần, Esc để thôi) | ✅ |
| Ctrl+Delete (khi sửa) | Xóa từ con trỏ tới cuối dòng | ✅ |


## D. Lộ trình đề xuất (mỗi phiên bản có test + kiểm chứng + bộ cài)

| Phiên bản | Nội dung |
|---|---|
| **0.5 – Giao diện giống hệt Excel 365 + In ấn (đã làm)** | Thanh tiêu đề + QAT + ô Tìm lệnh; Ribbon 3 tầng có nút tách đôi và nút ↘ hộp thoại; đủ 9 tab (thêm Bố trí Trang, Xem lại, Trợ giúp); Backstage toàn màn hình; lưới/tiêu đề/vùng chọn/chấm điền đúng màu và nét Excel; thẻ sheet + thanh trạng thái đầy đủ (Chế độ, Trung bình/Đếm/Tổng tùy chọn, 3 nút dạng xem, thanh trượt thu phóng); menu chuột phải đầy đủ + thanh mini; KeyTips (Alt/F10); phông Segoe UI/Aptos (Calibri khi không có). Giữ chạy tốt trên điện thoại (dạng ribbon đơn giản hóa). |
| **0.6.1 – Giao diện theo ảnh Excel 365 + ngôn ngữ (đã làm)** | Ribbon đủ nhóm như ảnh (Kiểu Ô dạng bộ sưu tập, Phần bổ trợ, Kiểu Dữ liệu, Hiệu suất, Dạng xem Trang tính, Ô Tiêu điểm, Dẫn hướng, thẻ Tự động hóa), Tệp ▸ Trang đầu (mẫu, Ưa thích, Tìm tệp, ngày tương đối, Chia sẻ), Trợ năng ở thanh trạng thái, chọn Tiếng Việt/English trong Tùy chọn. |
| **0.6 – Bàn phím giống hệt (đã làm)** | 167/180 mục phím của trang Microsoft trong phạm vi (bảng ở mục C), Tìm & Thay thế, Đi tới/Đặc biệt, Dán Đặc biệt, Chổi định dạng, Định dạng Ô (Ctrl+1), F4, Ctrl+`, Điền Nhanh, Chuỗi, Liên kết, chọn nhiều vùng, nhóm trang tính, Phân tích Nhanh, Bảng tạm Office, bảng phím tắt F1. Test bấm thật từng phím: 164/164 + 3 phím bản Windows. |
| **0.7 – Dữ liệu dự toán** | Tên vùng (đọc/ghi .xlsx), Bảng (Table) + tham chiếu có cấu trúc, Nhóm hàng/cột + Tổng phụ, Văn bản thành cột, Xóa trùng lặp, Ghi chú ô (đọc/ghi), hằng mảng nhiều cột, mảng động FILTER/SORT/UNIQUE. |
| **0.8 – In ấn hồ sơ** | Phần lớn đã kéo lên 0.5 theo yêu cầu “in ấn giống Excel 365”. Còn lại: Dạng xem Bố trí trang (thước, sửa đầu/chân trang tại chỗ), kéo đường ngắt trang, trang đầu/chẵn lẻ khác, ảnh trong đầu trang. |
| **0.9 – Phân tích** | Truy vết tiền lệ/phụ thuộc, Đánh giá công thức, Goal Seek, biểu đồ XY + Kết hợp + tab Thiết kế/Định dạng biểu đồ, ảnh trong sổ, liên kết (Ctrl+K), Bảo vệ sheet. |
| **1.0** | Rà soát lại bảng này, kiểm chứng tổng, cập nhật trạng thái. |

**Không làm (giữ nguyên lý do):** Power Query/Power Pivot/Mô hình dữ liệu, Copilot, Python trong Excel, chạy macro VBA, Office Scripts, cộng tác thời gian thực/OneDrive/SharePoint, Kiểu dữ liệu Cổ phiếu/Địa lý, Bản đồ 3D, Microsoft Forms, Nhãn độ nhạy.
