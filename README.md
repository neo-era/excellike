# ExcelLike

Phần mềm bảng tính chức năng tương tự Excel 365 (dùng cá nhân). Một mã nguồn chạy được dạng **app Windows (.exe)** (Electron) và **PWA** (cài trên máy tính, điện thoại).

**Dùng ngay trên web:** https://neo-era.github.io/excellike/ (cài được như ứng dụng từ Chrome/Edge, chạy cả khi mất mạng).

## Bản 0.5 — giao diện và in ấn như Excel 365

- **Giao diện:**
  - Thanh tiêu đề xanh có Tự động lưu, Thanh công cụ Truy nhập Nhanh tùy chỉnh được, tên tệp và ô **Tìm kiếm (Alt+Q)** (gõ không dấu vẫn tìm được).
  - 9 thẻ: Tệp, Trang đầu, Chèn, Bố trí Trang, Công thức, Dữ liệu, Xem lại, Xem, Trợ giúp.
  - Ribbon 3 tầng có nút tách đôi và nút ↘. **KeyTips** bằng Alt hoặc F10 (ví dụ Alt, H, 1 để in đậm).
  - Tệp mở trang **Backstage**.
  - Chủ đề Office: Nhiều màu, Xám Đậm, Đen, Trắng.
  - Thẻ trang tính có màu thẻ, ẩn/bỏ ẩn, di chuyển hoặc sao chép.
  - Thanh trạng thái có Trung bình, Số đếm, Tổng, Tối thiểu, Tối đa và thanh trượt thu phóng.
  - Menu chuột phải có thanh mini. Có Chèn/Xóa ô dồn ô, Chiều cao hàng / Độ rộng cột, Ẩn/Bỏ ẩn hàng, cột.
- **In ấn:**
  - Tệp ▸ In (Ctrl+P) có xem trước từng trang. Chọn được số bản, máy in, in trang tính / cả sổ / vùng chọn, trang từ–đến, khổ giấy, hướng, lề, co giãn (vừa 1 trang, vừa cột, vừa hàng, % hoặc X × Y trang).
  - **Thiết lập Trang** có đầu trang/chân trang (Trang 1 / N, tên trang tính, ngày, tên tệp), hàng và cột lặp lại, căn giữa trang, in đường lưới và đầu đề, thứ tự trang.
  - Có **ngắt trang thủ công** và dạng xem **Xem trước Ngắt Trang**.
  - Bản Windows in thẳng ra máy in đã chọn (không qua hộp thoại) và xuất **PDF** thật.

Lệnh Excel chưa làm vẫn hiện trên ribbon (mờ), ghi rõ “có ở bản 0.x”.

## Cài trên Windows

- **ExcelLike-Setup-0.5.0.exe**: bộ cài cho tài khoản đang dùng, không cần quyền quản trị. Có lối tắt Start/màn hình, mục “Mở bằng…” cho .xlsx và .csv; tùy chọn đặt ExcelLike làm ứng dụng mặc định. Gỡ trong Cài đặt → Ứng dụng (dữ liệu sổ tính được giữ lại).
- **ExcelLike-Portable-0.5.0.zip**: giải nén rồi chạy `ExcelLike.exe`, không cần cài.
- Nhấp đúp tệp .xlsx/.csv để mở (đang mở thì chuyển vào cửa sổ hiện có). **Ctrl+S** lưu đè đúng tệp đã mở; **Lưu thành…** chọn nơi lưu; đóng cửa sổ khi còn thay đổi chưa lưu vào tệp sẽ được hỏi.
- Windows SmartScreen có thể cảnh báo vì bộ cài chưa ký số: bấm “Thông tin thêm” → “Vẫn chạy”.

Điện thoại: ribbon đơn giản một hàng, chọn thẻ bằng nút “Trang đầu ▾”. Khi dọc có thanh lệnh dưới (Nhập, Hàm, Lọc, Biểu đồ, Thêm). Khi nằm ngang, ribbon chung hàng với thanh tiêu đề. Trên máy tính, Ctrl+F1 thu gọn hoặc ghim ribbon.

## Đợt 3 thêm gì

- **Biểu đồ**: Cột, Cột chồng, Thanh ngang, Đường, Vùng, Tròn, Vành khuyên. Chọn vùng (hoặc một ô trong bảng) → Chèn → Biểu đồ: hộp thoại có xem trước, tự nhận hàng tiêu đề (tên chuỗi) và cột chữ (nhãn). Kéo để di chuyển, kéo góc/cạnh để đổi cỡ, nhấp đúp hoặc nút **Sửa** để đổi loại/tiêu đề/vùng/chú giải/nhãn giá trị, **Delete** để xóa, Ctrl+Z hoàn tác. Tự vẽ lại khi dữ liệu đổi; dời theo khi chèn/xóa hàng, cột. Trục rút gọn kiểu Việt Nam (tr, tỷ).
- **Biểu đồ trong .xlsx**: mở được biểu đồ do Excel tạo (cột/thanh/đường/vùng/tròn/vành khuyên); lưu ra biểu đồ Excel thật (mở bằng Excel vẫn sửa được).
- **PivotTable**: Chèn → PivotTable (sheet mới hoặc ô chỉ định). Ngăn **Trường PivotTable**: tích trường (chữ → Hàng, số → Giá trị), kéo-thả giữa Bộ lọc / Cột / Hàng / Giá trị (hoặc dùng menu ⋯ trên điện thoại), cách tính Tổng / Đếm / Trung bình / Lớn nhất / Nhỏ nhất, lọc theo giá trị, tổng phụ khi có 2 cấp hàng, Tổng cộng. **Làm mới** khi dữ liệu nguồn đổi; **Xóa PivotTable**. Cột tự nới cho vừa số.
- **In / Lưu PDF** (Ctrl+P): toàn trang tính, vùng chọn hoặc vùng in; dọc/ngang; A4/A3/Letter; lề; thu vừa chiều ngang; lặp hàng tiêu đề mỗi trang; đường lưới; in kèm biểu đồ. Chọn "Lưu dưới dạng PDF" ở hộp thoại in. Thiết lập trang được ghi vào .xlsx (Excel đọc được).

Lưu ý: PivotTable lưu ra .xlsx dưới dạng bảng giá trị (chưa phải PivotTable "sống" của Excel). Hình ảnh, ghi chú, macro trong tệp vẫn chưa hỗ trợ (app báo trước khi mở).

## Đợt 2 thêm gì

- **Mở / lưu tệp Excel .xlsx**: giữ giá trị, công thức (tự đổi dấu `;`/`,` cho đúng Excel), định dạng số, phông và cỡ chữ, màu, tô nền, viền, căn lề, xuống dòng, ô gộp, độ rộng cột, chiều cao hàng, cố định ngăn, định dạng có điều kiện, kiểm tra dữ liệu, bộ lọc. Trên Chrome/Edge máy tính, "Lưu .xlsx" ghi thẳng vào tệp đã mở.
- **CSV**: nhập (tự nhận dấu `,` `;` tab, UTF-8 có BOM) thành sheet mới; xuất CSV dấu `;` (Excel tiếng Việt) hoặc dấu `,`.
- **Viền** (tất cả, ngoài, ngoài đậm, trong, trên/dưới/trái/phải, kép, nét đứt), **gộp ô** (gộp và căn giữa, gộp từng hàng, bỏ gộp), **xuống dòng** tự tăng chiều cao hàng, **căn dọc**, **phông/cỡ chữ**.
- **Sắp xếp** A→Z, Z→A, hoặc nhiều mức; tự nhận vùng dữ liệu và hàng tiêu đề; xếp chữ tiếng Việt đúng thứ tự (Đ sau D); ô trống luôn xuống cuối.
- **Lọc** (Ctrl+Shift+L): nút lọc ở hàng tiêu đề, chọn giá trị hoặc lọc theo chữ chứa; sắp xếp ngay trong menu lọc.
- **Định dạng có điều kiện**: lớn hơn / nhỏ hơn / trong khoảng / bằng / chứa chữ / trùng lặp / top-bottom N / trên trung bình / thanh dữ liệu / thang màu; quản lý và xóa quy tắc.
- **Kiểm tra dữ liệu**: danh sách (gõ hoặc lấy từ vùng ô, có mũi tên chọn), số nguyên, số thập phân, ngày, độ dài chữ; câu báo lỗi và gợi ý riêng.


## Đợt 1 có gì

- Lưới ô vẽ bằng Canvas, cuộn mượt với 100.000 hàng; cố định ngăn; đổi cỡ cột/hàng (kéo hoặc nhấp đúp để vừa chữ); thu phóng.
- Công thức theo kiểu Excel tiếng Việt: dấu `;` ngăn tham số, dấu `,` thập phân. Khoảng 420 hàm (SUM, IF, VLOOKUP, XLOOKUP, SUMIFS, TEXTJOIN, NETWORKDAYS, PMT…). Tự tính lại chỉ những ô phụ thuộc.
- Gõ công thức có gợi ý tên hàm (Tab để chọn), bấm chuột hoặc phím mũi tên để chèn địa chỉ ô.
- Nhập số kiểu Việt Nam: `1.250.000`, `1,5`, `15%`, `01/02/2026`.
- Hoàn tác / làm lại (Ctrl+Z / Ctrl+Y) cho cả nội dung, định dạng, chèn/xóa, sheet.
- Sao chép / cắt / dán (giữ công thức tương đối, dán được dữ liệu từ Excel thật), kéo điền (dãy số, ngày, "Tủ 1 → Tủ 2", công thức).
- Định dạng: đậm, nghiêng, gạch dưới, màu chữ, tô nền, căn lề, định dạng số (#,##0, tiền ₫, %, ngày), tăng/giảm số lẻ.
- Chèn/xóa hàng, cột; nhiều sheet (thêm, đổi tên, xóa, tham chiếu chéo `='Tổng hợp'!A1`).
- Tự lưu trong máy (IndexedDB), danh sách sổ đã lưu (Tệp → Mở…).
- Sáng / tối / theo máy; giao diện điện thoại dọc và ngang (sửa ô trên thanh công thức).
- PWA: cài đặt được, mở được khi mất mạng.

Chưa có (theo kế hoạch): mở/lưu .xlsx, CSV, viền, gộp ô, sắp xếp/lọc, định dạng có điều kiện, kiểm tra dữ liệu (Đợt 2); biểu đồ, PivotTable, PDF/in (Đợt 3); bản .exe (Đợt 4).

## Phím tắt chính

| Phím | Việc |
|---|---|
| F2 / nhấp đúp | Sửa ô |
| Enter / Tab (Shift để đi ngược) | Nhập và chuyển ô |
| Alt+Enter | Xuống dòng trong ô |
| Esc | Hủy sửa, bỏ viền sao chép |
| Ctrl+mũi tên, Ctrl+Shift+mũi tên | Nhảy tới mép dữ liệu, chọn tới mép |
| Ctrl+Home | Về A1 |
| Ctrl+C / X / V | Sao chép / cắt / dán |
| Ctrl+Z / Ctrl+Y | Hoàn tác / làm lại |
| Ctrl+B / I / U | Đậm / nghiêng / gạch dưới |
| Ctrl+D / Ctrl+R | Điền xuống / điền sang phải |
| Ctrl+; | Ngày hôm nay |
| Alt+= | Tổng tự động |
| Shift+F3 | Chèn hàm |
| Ctrl+S | Lưu |
| Ctrl+O | Mở tệp .xlsx / .csv |
| Ctrl+Shift+L | Bật/tắt lọc |
| Ctrl+P | In / Lưu PDF |
| Ctrl+F1 | Thu gọn / ghim ribbon |
| Delete / mũi tên / Enter (khi chọn biểu đồ) | Xóa / dịch chuyển / sửa biểu đồ |
| Ctrl+cuộn chuột | Thu phóng |

## Chạy và kiểm thử

```bash
npm install
npm run dev          # chạy bản phát triển tại http://localhost:5173
npm test             # unit test (Vitest): đối chiếu kết quả Excel, .xlsx khứ hồi, CSV, sắp xếp, lọc…
npm run build        # bản PWA trong thư mục dist/
npm run e2e          # kiểm thử trên trình duyệt (Đợt 1, 2, 3) + ảnh chụp mobile (cần Playwright)
npm run demo         # gộp thành 1 file HTML chạy thử (dist-demo/excellike-demo.html)
npm run desktop      # chạy bản máy tính (Electron)
npm run build:win    # đóng gói Windows: release/ExcelLike-Setup-x.y.z.exe + bản Portable .zip (cần makensis)
npm run e2e:desktop  # kiểm thử bản máy tính (Linux cần xvfb-run)
```

Mỗi lần đẩy mã lên nhánh `main`, GitHub Actions (`.github/workflows/pages.yml`) tự chạy unit test, build và cập nhật bản web trên GitHub Pages.

## Cấu trúc

```
src/core/   address · numfmt · input · tsv · csv · sizes · series · formula · locale · rules · pivot · chart · print · xlsx · workbook (lõi)
src/core/paginate (chia trang in, đầu/chân trang)
src/ui/     grid (lưới Canvas) · app (lệnh, ribbon, thanh tiêu đề/trạng thái, menu) · ribbon (ribbon, KeyTips) · menu · backstage · printui (trang In) · pagesetup · dialogs · charts + chartdraw (biểu đồ) · pivotui · printui · files · store (IndexedDB) · icons · funcs
tests/      unit test        e2e/   kiểm thử trình duyệt, điện thoại, bản máy tính       public/  manifest, service worker, icon
electron/   main.cjs (cửa sổ, mở/lưu tệp, một phiên) · preload.cjs (cầu nối an toàn)
build/      icon, installer.nsi (bộ cài NSIS), afterPack.cjs (gắn icon vào .exe)      tools/  build-win.mjs, artifact.mjs
```

## Giấy phép thư viện

Đọc/ghi .xlsx dùng ExcelJS (MIT). Bộ tính công thức HyperFormula dùng giấy phép GPLv3 (miễn phí khi dùng cá nhân). Nếu sau này giao cho đơn vị khác cần mua giấy phép thương mại hoặc công bố mã nguồn.
