# ExcelLike – Rà soát tính năng & phím tắt so với Excel 365 (Bước 0 cho v2)

Ngày tra: 28/10/2026. Đối chiếu **ExcelLike 0.4.1** với **Excel 365 cho Windows (Current Channel)**.

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
| Đóng (Ctrl+W), nhiều sổ mở cùng lúc, chuyển cửa sổ (Ctrl+F6, Ctrl+Tab) | ✗ (một sổ một lúc) | P2 |
| Mở .xls (97–2003), .xlsb, .ods; lưu .xlsm giữ macro | ✗ | P2 |
| Tùy chọn (Options): chế độ tính, tính lặp, hướng Enter, phông mặc định… | ◐ 0.5 (Tính lặp + số lần/sai số, hướng Enter) | P2 |
| Chia sẻ, OneDrive, Phiên bản, Power BI | — | — |

### Trang đầu (Home)
| Tính năng | ExcelLike | Ưu tiên |
|---|---|---|
| Cắt, Sao chép, Dán | ✅ | – |
| Dán đặc biệt (Ctrl+Alt+V): giá trị, công thức, định dạng, độ rộng cột, chuyển vị, phép tính, bỏ ô trống; nút Tùy chọn dán sau khi dán | ✗ | P1 |
| Chổi định dạng (nhấp/nhấp đúp) | ✗ | P1 |
| Phông, cỡ, Đậm/Nghiêng/Gạch chân, gạch ngang (Ctrl+5), màu chữ, tô nền | ✅ (gạch ngang chưa có nút) | – |
| Tăng/Giảm cỡ chữ, gạch chân kép, chỉ số trên/dưới | ✗ | P2 |
| Viền: các kiểu sẵn | ✅ | – |
| Viền: vẽ viền, màu/kiểu nét, Thêm viền… (hộp thoại) | ✗ | P2 |
| Bảng màu đầy đủ (màu chủ đề + sắc độ, Màu khác…) | ◐ | P1 |
| Căn trên/giữa/dưới, trái/giữa/phải, xuống dòng, gộp ô (3 kiểu), bỏ gộp | ✅ | – |
| Thụt lề (tăng/giảm), xoay chữ (góc, dọc), Căn giữa qua vùng chọn, Co chữ vừa ô | ✗ | P1 (thụt lề), P2 |
| Định dạng số: danh sách nhanh, Kế toán, %, dấu phân cách, tăng/giảm số lẻ | ✅ (Kế toán ◐) | – |
| Hộp thoại Định dạng ô (Ctrl+1): Số (Tùy chỉnh), Căn chỉnh, Phông, Viền, Tô, Bảo vệ | ✗ | P1 |
| Định dạng có điều kiện: tô ô theo quy tắc, trùng lặp, Top/Bottom, trên/dưới TB, thanh dữ liệu, thang màu, quản lý | ✅ | – |
| … Bộ biểu tượng, “Ngày xảy ra”, Quy tắc mới bằng công thức | ✗ | P2 |
| Định dạng như Bảng (Table: tên bảng, hàng tổng, dải màu, tham chiếu có cấu trúc `Bảng1[Cột]`) | ✗ | P1 |
| Kiểu ô (Cell Styles) | ✗ | P2 |
| Chèn/Xóa hàng, cột | ✅ | – |
| Chèn/Xóa ô (dồn xuống/phải, lên/trái) | ✅ 0.5 | P1 |
| Định dạng: chiều cao hàng/độ rộng cột theo số, Tự khớp, Ẩn/Hiện hàng-cột-sheet, Di chuyển/Sao chép sheet, Màu thẻ, Bảo vệ sheet, Khóa ô | ✅ 0.5 (trừ Bảo vệ sheet, Khóa ô → 0.9) | P1 |
| Tổng tự động ▸ Tổng/Trung bình/Đếm/Max/Min | ✅ 0.5 | P1 |
| Điền ▸ Xuống/Phải/Lên/Trái/Chuỗi…/Flash Fill | ◐ (Xuống, Phải, kéo điền chuỗi) | P2 |
| Xóa ▸ Tất cả/Định dạng/Nội dung/Ghi chú/Liên kết | ◐ 0.5 (Tất cả, Định dạng, Nội dung; Ghi chú → 0.7) | P1 |
| Sắp xếp & Lọc | ✅ | – |
| Lọc nâng cao: theo màu, theo số (lớn hơn, top 10), theo ngày (tháng này, quý…) | ◐ (giá trị, chữ chứa) | P2 |
| Tìm & Chọn: Tìm (Ctrl+F), Thay thế (Ctrl+H), Đi tới (F5/Ctrl+G), Đi tới Đặc biệt (ô trống, công thức, hằng số…) | ✗ | **P1** |
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
| Liên kết (Ctrl+K) | ✗ | P2 |
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
| Hộp thoại Đối số hàm (Ctrl+A sau tên hàm) | ✗ | P2 |
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
| Số: Chung, Số, Tiền tệ, Kế toán, Ngày, Giờ, %, Phân số, Khoa học, Văn bản, Đặc biệt, **Tùy chỉnh** (gõ mã, xem mẫu) | ◐ (bộ máy định dạng đủ, chưa có hộp thoại) | **P1** |
| Căn chỉnh: ngang (gồm Điền, Căn đều, Giữa qua vùng chọn, Phân phối), dọc, thụt lề, xuống dòng, co vừa ô, gộp, góc xoay | ◐ | P1 |
| Phông: kiểu, gạch chân (đơn/kép/kế toán), gạch ngang, chỉ số trên/dưới | ◐ | P1 |
| Viền: kiểu nét, màu, bấm từng cạnh, đường chéo | ◐ | P1 |
| Tô: màu, mẫu (pattern), chuyển màu | ◐ (màu) | P2 |
| Bảo vệ: Khóa, Ẩn công thức | ✗ | P2 |

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

## C. Phím tắt

**Tổng trang Microsoft: 298 dòng.** Không làm: Power Pivot (27), Copilot (9), làm mới dữ liệu ngoài (3), macro/VBA/Power Query (Alt+F8, Alt+F11, Alt+F12), lệnh cửa sổ con kiểu Windows cũ (Ctrl+F5/F7/F8/F9/F10), hình nổi Ctrl+Alt+5. **Trong phạm vi: ~245 dòng.**

| Mục (trang Microsoft) | Số dòng | ExcelLike đã có | Còn thiếu tiêu biểu |
|---|---|---|---|
| Thường dùng | 22 | 13 | Ctrl+W, Alt+H…, Shift+F10, Ctrl+9 / Ctrl+0 (ẩn hàng/cột), KeyTips Alt+H,H / Alt+H,B / Alt+H,D,C |
| Phím truy cập tab | 9 | 0 | Alt+F/H/N/P/M/A/R/W, Alt+Q |
| Làm việc với Ribbon | 12 | 1 (Ctrl+F1) | F10, Tab/mũi tên trên Ribbon, Ctrl+Shift+F10 |
| Di chuyển trong ô | 24 | 12 | End + mũi tên, Ctrl+End, Ctrl+Shift+End, Ctrl+PgUp/PgDn (đổi sheet), Alt+PgUp/PgDn, Alt+↓ (danh sách DV), Ctrl+Alt+= / − (thu phóng) |
| Định dạng ô | 38 | 6 | Ctrl+1, Ctrl+Shift+F, Shift+F2, Ctrl+Shift++, Ctrl+−, Ctrl+Shift+:, Ctrl+`, Ctrl+', Ctrl+Alt+V, Ctrl+2/3/4/5, Ctrl+Shift+& / _, Ctrl+Shift+~ $ % ^ # @ !, Ctrl+K, Ctrl+T/L, Ctrl+8 |
| Dán đặc biệt | 11 | 0 | A F V T C N H X W R U |
| Chọn & thao tác | 21 | 10 | Ctrl+A theo vùng hiện tại (lần 2 cả sheet), Ctrl+Space, Shift+Space, F8, Shift+F8, Ctrl+Enter, Ctrl+Shift+*, Ctrl+Shift+Home, Ctrl+Shift+PgUp/PgDn |
| Dữ liệu, hàm, thanh công thức | 27 (trừ 3 macro/PQ) | 7 | F4 (đổi $ tham chiếu và lặp lệnh), Ctrl+Shift+U, Shift+F9, Ctrl+Alt+F9, Ctrl+Shift+A, Ctrl+E, Ctrl+Shift+", Alt+F1, F11, F3, Alt+M,M,D |
| Phím chức năng F1–F12 | 39 (trừ ~8 cửa sổ/macro) | 6 | F1 (trợ giúp/bảng phím), Alt+Shift+F1 / Shift+F11 (sheet mới), Ctrl+Shift+F1 (toàn màn hình), Ctrl+F2, F5, F6, F7, F10, Shift+F10, F12 |
| Phím hữu ích khác | 58 | 30 | End mode, Alt+Space, Shift+kéo để chèn, Ctrl+Tab trong hộp thoại |
| **Tổng trong phạm vi** | **~245** | **~85 (≈35%)** | |

Ghi chú hành vi phải giống Excel: Ctrl+A chọn vùng dữ liệu trước rồi mới cả sheet; F4 khi đang sửa công thức đổi `A1 → $A$1 → A$1 → $A1`, khi không sửa thì lặp lệnh cuối; Ctrl+Y = lặp lại/làm lại; Enter sau Tab quay về cột bắt đầu; Ctrl+D/Ctrl+R chép cả định dạng; Delete giữ định dạng; Backspace xóa rồi vào sửa ô.

## D. Lộ trình đề xuất (mỗi phiên bản có test + kiểm chứng + bộ cài)

| Phiên bản | Nội dung |
|---|---|
| **0.5 – Giao diện giống hệt Excel 365 + In ấn (đã làm)** | Thanh tiêu đề + QAT + ô Tìm lệnh; Ribbon 3 tầng có nút tách đôi và nút ↘ hộp thoại; đủ 9 tab (thêm Bố trí Trang, Xem lại, Trợ giúp); Backstage toàn màn hình; lưới/tiêu đề/vùng chọn/chấm điền đúng màu và nét Excel; thẻ sheet + thanh trạng thái đầy đủ (Chế độ, Trung bình/Đếm/Tổng tùy chọn, 3 nút dạng xem, thanh trượt thu phóng); menu chuột phải đầy đủ + thanh mini; KeyTips (Alt/F10); phông Segoe UI/Aptos (Calibri khi không có). Giữ chạy tốt trên điện thoại (dạng ribbon đơn giản hóa). |
| **0.6 – Bàn phím giống hệt** | ~245 phím trong phạm vi + Find/Replace/Go To/Go To Special, Dán đặc biệt, Chổi định dạng, Định dạng ô (Ctrl+1), F4, Ctrl+`, bảng phím tắt F1. Test bấm thật từng phím. |
| **0.7 – Dữ liệu dự toán** | Tên vùng (đọc/ghi .xlsx), Bảng (Table) + tham chiếu có cấu trúc, Nhóm hàng/cột + Tổng phụ, Văn bản thành cột, Xóa trùng lặp, Ghi chú ô (đọc/ghi), hằng mảng nhiều cột, mảng động FILTER/SORT/UNIQUE. |
| **0.8 – In ấn hồ sơ** | Phần lớn đã kéo lên 0.5 theo yêu cầu “in ấn giống Excel 365”. Còn lại: Dạng xem Bố trí trang (thước, sửa đầu/chân trang tại chỗ), kéo đường ngắt trang, trang đầu/chẵn lẻ khác, ảnh trong đầu trang. |
| **0.9 – Phân tích** | Truy vết tiền lệ/phụ thuộc, Đánh giá công thức, Goal Seek, biểu đồ XY + Kết hợp + tab Thiết kế/Định dạng biểu đồ, ảnh trong sổ, liên kết (Ctrl+K), Bảo vệ sheet. |
| **1.0** | Rà soát lại bảng này, kiểm chứng tổng, cập nhật trạng thái. |

**Không làm (giữ nguyên lý do):** Power Query/Power Pivot/Mô hình dữ liệu, Copilot, Python trong Excel, chạy macro VBA, Office Scripts, cộng tác thời gian thực/OneDrive/SharePoint, Kiểu dữ liệu Cổ phiếu/Địa lý, Bản đồ 3D, Microsoft Forms, Nhãn độ nhạy.
