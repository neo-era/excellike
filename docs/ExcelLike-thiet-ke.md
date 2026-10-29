# ExcelLike — Doc thiết kế (Bước 1)

Phần mềm bảng tính chức năng tương tự Excel 365. Người dùng: chỉ anh Lâm dùng cá nhân.
Trạng thái: Bước 0 đã duyệt 28/10/2026 · Chốt UI: chọn **phương án A – Ribbon** (28/10/2026) · Đợt 1 đã duyệt (28/10/2026) · Đợt 2 xong, đã review độc lập + kiểm chứng, chờ anh duyệt Bước 5.

## 1. Mục tiêu v1

Một mã nguồn web chạy được hai kiểu:
- **PWA**: cài được trên desktop (Chrome/Edge) và mobile (Android, iOS), mở lại khi offline.
- **App Windows (Electron)**: cài bằng file .exe, chạy offline, mở/lưu file trực tiếp, gắn đuôi .xlsx.

Phạm vi: Cốt lõi + Xlsx + Biểu đồ + PivotTable. Không có kênh góp ý (dùng cá nhân).

## 2. Tiêu chí đạt (đã duyệt ở Bước 0)

| # | Tiêu chí | Cách đo |
|---|---|---|
| T1 | Lưới ô: chọn ô/vùng, kéo điền, sao chép/dán, undo/redo, đóng băng ngăn, đổi cỡ cột/hàng, nhiều sheet | Test E2E Playwright |
| T2 | Công thức cho kết quả trùng Excel | Bộ file .xlsx mẫu có giá trị Excel đã tính sẵn → so khớp từng ô |
| T3 | Tính lại theo phụ thuộc; sửa 1 ô chỉ tính các ô liên quan; phát hiện tham chiếu vòng | Unit test + đếm số ô được tính lại |
| T4 | Định dạng ô: số, tiền, %, ngày, font, màu, viền, căn lề, gộp ô, định dạng có điều kiện | Unit + ảnh chụp |
| T5 | Sắp xếp, lọc, Data Validation | Unit + E2E |
| T6 | Biểu đồ cột, đường, tròn; tự cập nhật khi đổi dữ liệu | E2E |
| T7 | PivotTable: kéo-thả hàng/cột/giá trị/bộ lọc; Sum/Count/Average | Đối chiếu với Pivot của Excel trên cùng dữ liệu |
| T8 | Mở/lưu .xlsx giữ công thức, định dạng, nhiều sheet; CSV vào/ra; xuất PDF; in | Mở → lưu → mở lại, so sánh |
| T9 | 100.000 hàng × 20 cột: cuộn mượt, mở và tính lại trong ngưỡng | Đo thời gian (chốt ngưỡng ở Bước 5) |
| T10 | PWA cài được, offline vào được shell | Lighthouse + cài thật |
| T11 | UI hiện đại, sáng/tối, responsive (mobile dọc + ngang, 360×640), dùng được bằng bàn phím, WCAG AA | Ảnh chụp tự động + kiểm tra bàn phím |

Ca biên/ca xấu: tham chiếu vòng, chia 0, ô trống hoặc chữ lẫn số, công thức lồng sâu, file .xlsx hỏng hoặc có tính năng chưa hỗ trợ (báo rõ, không mất dữ liệu), dán dữ liệu rất lớn, hàm chưa có (#NAME?).

Chưa làm ở v1: macro/VBA, cộng tác thời gian thực, Power Query, LET/LAMBDA và mảng động đầy đủ, nguồn dữ liệu ngoài.

## 3. Kiến trúc

```
┌──────────────────────── Giao diện (React + TypeScript) ────────────────────────┐
│ Ribbon/Toolbar · Thanh công thức · Lưới Canvas · Tab sheet · Bảng Biểu đồ/Pivot │
└───────────────┬─────────────────────────────────────────────┬──────────────────┘
                │ lệnh (command) + undo/redo                   │ sự kiện thay đổi
┌───────────────▼──────────────┐   ┌──────────────────────────▼─────────────────┐
│ Mô hình Workbook (core)       │◄──┤ Engine công thức (HyperFormula, Web Worker)│
│ sheet, ô, style, merge, CF,   │   │ parse, đồ thị phụ thuộc, tính lại từng phần │
│ validation, chart, pivot      │   └────────────────────────────────────────────┘
└───────┬───────────────┬──────┘
        │               │
┌───────▼─────┐  ┌──────▼────────────────────────────────┐
│ IO: ExcelJS │  │ Lớp nền tảng (platform adapter)        │
│ .xlsx / CSV │  │ PWA: File System Access + IndexedDB    │
│ PDF / in    │  │ Electron: fs, hộp thoại Open/Save, .exe│
└─────────────┘  └───────────────────────────────────────┘
```

### Lựa chọn công nghệ

| Phần | Chọn | Lý do |
|---|---|---|
| Ngôn ngữ, build | TypeScript, Vite | Một mã nguồn cho PWA và Electron |
| UI | TypeScript thuần (không framework) | Đổi so với kế hoạch ban đầu (React): ribbon và hộp thoại đơn giản, lưới vẽ Canvas; bỏ React giúp nhẹ và ít phụ thuộc |
| Lưới ô | Tự viết, vẽ Canvas, ảo hóa hàng/cột | Đạt T9 (100.000 hàng); DOM không chịu nổi |
| Engine công thức | HyperFormula 3.4, chạy trên luồng chính (Đợt 1) | ~420 hàm, đồ thị phụ thuộc, tính lại từng phần. Sửa 1 ô ~5 ms nên chưa cần Worker; xem lại khi có sổ rất lớn |
| Đọc/ghi .xlsx | ExcelJS | Đọc/ghi được style, merge, CF, validation |
| Biểu đồ | ECharts | Cột/đường/tròn, vẽ Canvas, nhẹ |
| PivotTable | Tự viết (group-by trên vùng dữ liệu) | Đơn giản, kiểm soát được |
| PDF/in | Electron: `printToPDF`; PWA: CSS in + hộp thoại in của trình duyệt | |
| Lưu tạm/khôi phục | IndexedDB, tự lưu mỗi 30 giây | Không mất dữ liệu khi app bị đóng |
| Đóng gói Windows | Electron + electron-builder (NSIS .exe) | Gắn đuôi .xlsx, cập nhật thủ công |
| Test | Vitest (unit), Playwright (E2E + ảnh chụp mobile) | |

## 4. Module

| Module | Việc | Ghi chú |
|---|---|---|
| `core/model` | Workbook, Sheet, Cell, Style, Range, địa chỉ A1 ↔ (r,c) | Thuần TS, không phụ thuộc UI |
| `core/commands` | Mọi thay đổi đi qua lệnh có `do/undo` | Undo/redo, kéo điền, dán |
| `core/engine` | Bọc HyperFormula trong Worker; đồng bộ ô ↔ engine | |
| `core/format` | Định dạng số kiểu Excel (`#,##0`, `0.0%`, `dd/mm/yyyy`), locale vi-VN | |
| `core/cf`, `core/validation` | Định dạng có điều kiện, Data Validation | |
| `core/sortfilter` | Sắp xếp nhiều khóa, AutoFilter | |
| `core/pivot` | Tính bảng Pivot từ vùng nguồn | |
| `io/xlsx`, `io/csv`, `io/pdf` | Nhập/xuất | |
| `ui/grid` | Canvas: vẽ, cuộn ảo, chọn, sửa tại ô, đóng băng | |
| `ui/ribbon`, `ui/formula-bar`, `ui/sheets`, `ui/chart`, `ui/pivot` | Giao diện theo phương án được chọn | |
| `platform/pwa`, `platform/electron` | File, menu, cài đặt, offline | |

## 5. Rủi ro và cách xử lý

| Rủi ro | Xử lý |
|---|---|
| HyperFormula giấy phép GPLv3 (miễn phí khi dùng cá nhân); nếu sau này bán/giao cho đơn vị khác phải mua giấy phép hoặc mở mã nguồn | Ghi rõ; bọc qua lớp `core/engine` để thay engine được |
| ExcelJS **không ghi được biểu đồ và PivotTable gốc** vào .xlsx | v1: lưu biểu đồ/pivot trong phần dữ liệu riêng của app (vẫn mở lại được trong ExcelLike); khi mở bằng Excel thật, pivot hiện thành bảng giá trị, biểu đồ không hiện. Báo rõ khi lưu |
| File .xlsx dùng tính năng chưa hỗ trợ (macro, slicer…) | Giữ nguyên phần không hiểu khi có thể, cảnh báo trước khi lưu đè |
| Hiệu năng 100.000 hàng | Canvas + ảo hóa + engine trong Worker; đo sớm ở task 3 |
| Bàn phím ảo trên mobile che ô đang sửa | Sửa trong thanh công thức trên mobile, cuộn ô vào vùng nhìn thấy |
| iOS không có File System Access | Mở bằng chọn file, lưu bằng Chia sẻ/tải xuống |

## 5b. Quyết định kỹ thuật đã chốt khi làm Đợt 1

- **Quy ước tiếng Việt**: công thức dùng `;` ngăn tham số, `,` thập phân; hiển thị nhóm nghìn `.`. Hằng mảng viết `{1,2}` bị nhầm với số thập phân → chưa hỗ trợ hằng mảng (dùng vùng ô thay thế). Khi làm .xlsx (Đợt 2) phải đổi `;`↔`,` và `,`↔`.` lúc đọc/ghi.
- **Sai khác của HyperFormula so với Excel đã vá**: `INT` (số âm), `MOD` (dấu theo số chia), `TRUE`/`FALSE` viết không ngoặc.
- **Lỗi vòng lặp** hiển thị `#CYCLE!` (Excel hiện 0 kèm cảnh báo).
- **Hoàn tác**: một ngăn xếp chung; mỗi mục ghi "có bước của HyperFormula hay không" + hàm hoàn tác riêng cho định dạng/cỡ cột/cố định ngăn.
- **Kéo điền công thức**: tự dời tham chiếu bằng `core/formula.ts` (shiftFormula), không dùng copy/paste của HyperFormula (không chạy được trong batch).
- **Mở sổ lớn**: dùng `buildFromSheets` (nhanh hơn ~2,5 lần): 2 triệu ô mở trong ~4 giây.
- **Điện thoại**: sửa ô trên thanh công thức (không bị bàn phím ảo che); chạm 2 lần để sửa, giữ lâu để mở menu.
- **Định dạng cả cột/hàng** chỉ áp cho phần đã có dữ liệu (chưa có kiểu định dạng theo cột như Excel).

## 6. Kế hoạch task (thứ tự, mỗi task có test)

| # | Task | Phụ thuộc | Kiểm chứng |
|---|---|---|---|
| 0 | **Chốt UI: 3 phương án A/B/C** (desktop + mobile dọc/ngang, sáng/tối) | – | Anh chọn |
| 1 | Khung dự án Vite + TS + Vitest + Playwright; PWA (manifest, icon, service worker) | 0 | Lighthouse Installable |
| 2 | `core/model` + địa chỉ A1 + `core/commands` (undo/redo) | 1 | Unit |
| 3 | `ui/grid` Canvas: cuộn ảo, chọn, sửa ô, đóng băng, đổi cỡ | 2 | E2E + đo 100.000 hàng |
| 4 | `core/engine` HyperFormula trong Worker; thanh công thức | 2 | Bộ đối chiếu Excel |
| 5 | Sao chép/dán, kéo điền, chèn/xóa hàng cột, nhiều sheet | 3, 4 | E2E |
| 6 | `core/format` + định dạng ô + gộp ô | 3 | Unit + ảnh |
| 7 | `io/xlsx` + CSV (mở/lưu, round-trip) | 4, 6 | So sánh trước/sau |
| 8 | Sắp xếp, lọc, Data Validation, định dạng có điều kiện | 6 | Unit + E2E |
| 9 | Biểu đồ cột/đường/tròn | 4 | E2E |
| 10 | PivotTable | 4 | Đối chiếu Pivot Excel |
| 11 | Xuất PDF, in | 6 | Mở PDF kiểm tra |
| 12 | Electron: đóng gói .exe, hộp thoại file, gắn đuôi .xlsx | 7 | Cài thật trên Windows |
| 13 | Mobile: dọc/ngang/360px, sửa ô qua thanh công thức | 3 | Bộ ảnh Playwright |
| 14 | Review độc lập (subagent) → sửa → Bước 5 kiểm chứng | tất cả | Báo cáo đạt/chưa đạt |

Làm theo từng đợt, mỗi đợt anh nhận một bản chạy được:
- **Đợt 1** (task 1–5): lưới + công thức + nhiều sheet, chạy trên trình duyệt.
- **Đợt 2** (task 6–8): định dạng, .xlsx, dữ liệu.
- **Đợt 3** (task 9–11): biểu đồ, pivot, PDF.
- **Đợt 4** (task 12–14): bản .exe, hoàn thiện mobile, kiểm chứng.

## 6b. Kết quả kiểm chứng Đợt 1 (28/10/2026)

| Hạng mục | Kết quả |
|---|---|
| Unit test (Vitest) | 58/58 đạt (gồm bộ đối chiếu Excel và 12 ca từ review) |
| Kiểm thử trình duyệt (Playwright) | 38/38 đạt |
| Mở sổ 100.000 × 20 ô (2 triệu ô, 100.000 công thức) | ~4 giây |
| Sửa 1 ô → tính lại | chỉ ô phụ thuộc, ~5 ms |
| Vẽ khi cuộn (trình duyệt không GPU) | trung vị ~5 ms/khung |
| PWA | manifest + service worker đạt, mất mạng vẫn mở được |
| Mobile | dọc 390×844, ngang 844×390, máy nhỏ 360×640, sáng + tối: không tràn, nút ≥ 40px |

Review độc lập tìm 15 lỗi (3 nghiêm trọng: hoàn tác định dạng lệch sau chèn/dán; mở lỗi thì ghi đè sổ cũ; tên sheet dạng số/Q1 mất dấu nháy) — đã sửa hết, mỗi lỗi có test.

Chưa thử trên máy thật: gõ tiếng Việt bằng Unikey/EVKey, cài PWA thật trên Android/iPhone, cảm ứng thật.

## 6c. Đợt 2 (28/10/2026) — .xlsx, CSV, viền, gộp ô, sắp xếp, lọc, CF, DV

| Hạng mục | Kết quả |
|---|---|
| Unit test | 90/90 đạt (thêm: chuyển công thức Việt↔Excel, CSV, gộp ô, viền, sắp xếp, lọc, CF, DV, đọc .xlsx kiểu Excel, khứ hồi .xlsx, 17 ca từ review) |
| Kiểm thử trình duyệt | Đợt 2: 21/21 · Đợt 1: 38/38 (không hồi quy) |
| Tệp .xlsx app lưu ra | openpyxl đọc lại đúng công thức, ô gộp, màu, lọc, cố định ngăn, CF, DV |

Quyết định kỹ thuật Đợt 2:
- ExcelJS 4.4 (MIT) đọc/ghi; phần **Kiểm tra dữ liệu tự đọc từ XML** (JSZip) vì ExcelJS trải từng ô (cột C:C mất 26 giây, 641 MB) và bỏ qua dạng mới trong extLst.
- Công thức trong tệp dùng `,`/`.`; app dùng `;`/`,` → `core/locale.ts` đổi hai chiều (cả hằng mảng `{1,2;3,4}` ↔ `{1,2|3,4}`), thêm/bỏ tiền tố `_xlfn.` cho hàm mới.
- Chữ trông như số/công thức (`+84…`, `1E5`, `=abc`, `0123`) được giữ là chữ bằng dấu `'` như Excel.
- Danh sách Kiểm tra dữ liệu dài >255 ký tự hoặc có dấu phẩy/ngoặc kép → ghi ra sheet ẩn `_ExcelLike_DS` (app tự đọc lại).
- Độ rộng cột: 8 px/ký tự (phông của app rộng hơn Calibri) để số không thành ####.
- **Độ chính xác số: 15 chữ số có nghĩa** (HyperFormula mặc định chỉ 10 → đã sửa).
- Định dạng số Excel: nhiều phần (`;`), màu `[Red]`, chữ cố định, `_ *`, giờ `h:mm:ss`, `[h]:mm`, `mmm`, khoa học `0.00E+00`, chia nghìn `#,##0,,`.
- Khi đang lọc, sắp xếp chỉ đụng các hàng đang hiện (như Excel).

Chưa làm/còn hạn chế: hình ảnh, biểu đồ, ghi chú, macro trong tệp (báo trước, mất nếu lưu đè); điều kiện lọc không lưu vào tệp (hàng ẩn vẫn giữ); xóa CF/DV ở vùng chọn xóa cả quy tắc chồng lấn; hằng mảng có số thập phân.

## 6d. Đợt 3 (28/10/2026) — biểu đồ, PivotTable, In / PDF

| Hạng mục | Kết quả |
|---|---|
| Unit test | 117/117 đạt (thêm: Pivot đối chiếu kết quả Excel, dữ liệu & trục biểu đồ, biểu đồ trong .xlsx đọc/ghi, trang in, 13 ca từ review) |
| Kiểm thử trình duyệt | Đợt 3: 32/32 · Đợt 2: 21/21 · Đợt 1: 38/38 (không hồi quy) |
| Tệp .xlsx app lưu ra | openpyxl đọc được 7 loại biểu đồ, hướng giấy, lặp hàng tiêu đề |
| In | PDF tạo từ trang in: đúng khổ A4 ngang, có số liệu và ảnh biểu đồ (pdftotext/pdfinfo) |

Quyết định kỹ thuật Đợt 3:
- Biểu đồ vẽ bằng Canvas riêng (`ui/chartdraw.ts`), không thêm thư viện; nằm trên lớp nổi phía trên lưới, neo theo ô + lệch (như Excel), kéo/đổi cỡ/sửa/xóa đều hoàn tác được.
- Biểu đồ trong .xlsx: tự đọc `drawing` + `chart` XML (JSZip) và gỡ khỏi gói trước khi đưa ExcelJS (ExcelJS lỗi với biểu đồ); khi lưu tự sinh `chartN.xml`, `drawingN.xml`, liên kết và content type. Bộ đệm giá trị ghi đủ cả hàng ẩn để chỉ số khớp vùng.
- PivotTable tính trong lõi (`core/pivot.ts`), ghi kết quả ra ô như bảng giá trị; định nghĩa Pivot lưu trong sổ (IndexedDB). Lưu ra .xlsx là bảng giá trị, chưa phải PivotTable "sống" của Excel.
- Như Excel: không chèn/xóa hàng/cột cắt ngang Pivot; không đặt Pivot chồng lên vùng nguồn; tên cột trùng được đánh số; trường ngày hiện theo ngày đã định dạng.
- Xóa sheet: biểu đồ/Pivot trỏ tới sheet đó bị gỡ tham chiếu (mã sheet của HyperFormula được dùng lại) — hoàn tác trả lại.
- In: dựng HTML riêng (khổ, hướng, lề, vừa chiều ngang bằng zoom, lặp tiêu đề bằng `thead`, ô gộp bằng rowspan/colspan) rồi in qua khung ẩn; trang in có CSP chặn script, kiểu ô được thoát ký tự.
- Mở tệp openpyxl có đường dẫn tuyệt đối/ghi chú đặt chỗ không chuẩn: tự sửa đường dẫn, gỡ ghi chú và báo trước.

Review độc lập Đợt 3: 2 lỗi nặng (chèn mã qua tên phông trong trang in; biểu đồ/Pivot bám nhầm sheet mới sau khi xóa sheet), 7 lỗi vừa, 7 lỗi nhẹ — đã sửa, mỗi lỗi có test.

Chưa làm/còn hạn chế: PivotTable chưa ghi dạng Pivot "sống" vào .xlsx; ô lỗi (#DIV/0!) trong cột giá trị bị bỏ qua khi tổng hợp; biểu đồ tán xạ/radar/kết hợp chưa hỗ trợ; hình ảnh, hình vẽ, ghi chú trong tệp vẫn mất nếu lưu đè (có báo trước); điện thoại nằm ngang vùng lưới còn hẹp (xử lý ở Đợt 4).

## 6e. Đợt 4 (28/10/2026) — bản Windows .exe, điện thoại, kiểm chứng cuối

Quyết định kỹ thuật Đợt 4:
- Electron 33: nạp bản build web qua giao thức riêng `app://` (chặn thoát thư mục, gắn CSP); `contextIsolation` + `sandbox`, trang chỉ thấy cầu nối `xlDesktop` (mở, lưu, nhận tệp, đóng cửa sổ). Một phiên duy nhất: nhấp đúp tệp khi app đang mở → chuyển vào cửa sổ hiện có.
- Lưu tệp: ghi ra tệp tạm (tên riêng mỗi lần) rồi đổi tên; các lần lưu cùng một tệp chạy tuần tự; chỉ ghi thẳng vào tệp người dùng đã chọn (danh sách nhớ ở thư mục dữ liệu app), đường dẫn lạ → hỏi bằng hộp thoại Lưu thành.
- “Có thay đổi chưa lưu vào tệp” được lưu kèm sổ nháp → mở lại app vẫn nhớ. Đóng cửa sổ: hỏi Lưu / Không lưu / Hủy; trang treo hoặc chưa sẵn sàng thì vẫn đóng được (bản nháp đã tự lưu).
- Bộ cài: NSIS viết tay (makensis gốc, không cần Wine), cài cho người dùng hiện tại vào `%LOCALAPPDATA%\Programs\ExcelLike` (luôn thêm thư mục con ExcelLike), lối tắt, “Mở bằng…” cho .xlsx/.csv, tùy chọn đặt làm mặc định; gỡ cài đặt chỉ xóa đúng tệp đã cài, giữ dữ liệu sổ tính. Biểu tượng + thông tin phiên bản gắn vào .exe bằng resedit. Kèm bản Portable .zip.
- Phông Be Vietnam Pro đóng gói kèm (không gọi Google Fonts) → chạy hoàn toàn offline.
- Điện thoại: ribbon thu gọn được (Ctrl+F1, nút ⌃/⌄, nhấp đúp thẻ); nằm ngang thì mặc định thu gọn, chạm thẻ để hiện tạm; thanh tiêu đề + thẻ chung một hàng.

Review độc lập Đợt 4: 1 lỗi nặng (Esc ở hộp thoại hỏi lưu làm không đóng được cửa sổ), 4 lỗi vừa (lưu chồng nhau có thể hỏng tệp; nút Lưu trong hộp thoại đóng; trạng thái chưa lưu mất khi mở lại; trang có thể ghi vào đường dẫn bất kỳ), 7 lỗi nhẹ — đã sửa, có kiểm thử.

## 6f. Bản 0.5 (29/10/2026) — giao diện như Excel 365 + in ấn như Excel 365

**Yêu cầu:** “giao diện và chức năng hoàn toàn giống Excel 365” (lộ trình 0.5–1.0, đã duyệt) và “chức năng in ấn nâng cấp hoàn toàn giống Excel 365” (gộp phần in của 0.8 vào 0.5). Chốt giao diện: phương án A – Nhiều màu.

**Đã làm**
- Thanh tiêu đề tự vẽ:
  - Tự động lưu ghi thẳng vào tệp .xlsx đang mở.
  - QAT tùy biến được 9 lệnh.
  - Bấm tên tệp để đổi tên.
  - Tìm kiếm lệnh Alt+Q, tìm được cả khi gõ không dấu.
- Bản Windows là cửa sổ không viền. Nút cửa sổ của hệ thống đổi màu theo chủ đề.
- 9 thẻ đúng thứ tự Excel. Ribbon 3 tầng gồm nút lớn, nút nhỏ có chữ, nút biểu tượng, nút tách đôi, hộp chọn phông/cỡ/định dạng số và nút ↘ mở hộp thoại.
  - Nhóm tự co thành một nút khi cửa sổ hẹp, theo thứ tự ưu tiên như Excel.
  - Lệnh chưa làm vẫn hiện nhưng mờ, ghi rõ “có ở bản 0.x”.
- KeyTips: Alt hoặc F10, Alt+chữ, gõ nhiều ký tự (Alt, H, A, C), Esc để lùi một cấp. Mã phím dùng theo Excel.
- Backstage: Trang đầu, Mới, Mở, Thông tin, Lưu, Lưu thành, In, Xuất, Đóng, Tài khoản (Chủ đề Office: Nhiều màu / Xám Đậm / Đen / Trắng / Hệ thống), Tùy chọn (tính lặp, hướng Enter).
- **In ấn:**
  - App tự chia trang (`core/paginate.ts`), nên xem trước và bản in là cùng một HTML, mỗi trang một khối đúng khổ giấy.
  - Thiết lập in:
    - 8 khổ giấy.
    - Lề Thường/Rộng/Hẹp/Tùy chỉnh (inch như Excel).
    - Co giãn % hoặc vừa X × Y trang (tìm tỷ lệ nguyên lớn nhất như Excel).
    - Ngắt trang thủ công.
    - Hàng và cột lặp lại, không cắt ngang ô gộp.
    - Thứ tự trang xuống-rồi-ngang hoặc ngang-rồi-xuống.
    - Căn giữa trang, in đường lưới, in đầu đề hàng/cột.
    - Đầu trang và chân trang 3 phần với mã &P &N &D &T &A &F, số trang đầu tiên.
    - Số không vừa ô in thành ####.
  - Trang In trong Backstage có:
    - Số bản, Máy in, In trang tính / cả sổ / vùng chọn, Trang từ–đến, Đối chiếu, Hướng, Khổ, Lề, Co giãn.
    - Liên kết Thiết lập Trang.
    - Xem trước ◀ 1/N ▶, Hiện Lề, Thu phóng Trang.
  - Hộp Thiết lập Trang có 4 thẻ.
  - Thẻ Bố trí Trang gồm Lề, Hướng, Kích thước, Vùng In, Ngắt, In Tiêu đề, Co giãn và Tùy chọn Trang tính.
  - Dạng xem Xem trước Ngắt Trang.
  - **Bản Windows:**
    - Đọc danh sách máy in của Windows và in thẳng không qua hộp thoại hệ thống (số bản, đối chiếu, khổ giấy).
    - Nút “Thuộc tính Máy in” mở hộp thoại của Windows.
    - “Lưu thành PDF” ghi tệp .pdf thật bằng printToPDF.
  - **Bản web:** in qua hộp thoại của trình duyệt, có “Lưu dưới dạng PDF”.
  - Khứ hồi .xlsx toàn bộ thiết lập in. Ngắt trang theo cột được tự ghi vì ExcelJS không hỗ trợ.
- Thẻ sheet:
  - Có ◀ ▶ ⊕.
  - Menu chuột phải: Chèn…, Xóa, Đổi tên, Di chuyển hoặc Sao chép… (tạo bản sao), Màu Tab (bảng màu Excel), Ẩn, Bỏ ẩn…
  - Ctrl+PgUp/PgDn chuyển sheet, Shift+F11 tạo sheet mới.
- Thanh trạng thái:
  - Chế độ, Trung bình/Số đếm/Số đếm Số/Tối thiểu/Tối đa/Tổng (bấm chuột phải để chọn mục).
  - 3 nút dạng xem.
  - Thanh trượt thu phóng 10–100–400%, hộp Thu phóng.
- Menu chuột phải ô, tiêu đề hàng, tiêu đề cột theo Excel, kèm thanh mini.
- Hộp Chèn… / Xóa… dồn ô. Chiều cao hàng tính theo điểm, độ rộng cột tính theo ký tự. Ẩn/Bỏ ẩn hàng và cột (giữ kích thước gốc).
- Lưới:
  - Ô mặc định 64×20 px.
  - Phông giao diện Segoe UI, phông ô Aptos Narrow → Calibri, 11 pt = 14,67 px, đậm 700.
  - Màu tiêu đề/đường lưới/vùng chọn như Excel.
  - Hiện Công thức, ẩn/hiện Đầu đề và Thanh công thức.
- Điện thoại:
  - Ribbon đơn giản một hàng, chọn thẻ bằng nút chọn thẻ.
  - Khi dọc có thanh lệnh dưới (Nhập, Hàm, Lọc, Biểu đồ, Thêm). Khi nằm ngang, ribbon chung hàng với thanh tiêu đề.
- Định dạng Phân số (`# ?/?`, `# ?/4`).

**Quyết định kỹ thuật**
- Chia trang trong app thay vì để trình duyệt chia. Lý do: số trang trong đầu/chân trang, “Trang 1 / N”, xem trước từng trang và in thẳng ra máy in phải cho cùng một kết quả.
- `@page { size: <khổ>; margin: 0 }` và mỗi `.page` có kích thước mm cố định. Lề nằm bên trong trang (inch như Excel) nên Chromium/Electron không tự thêm lề.
- Lấy số trang `&N` theo số trang cuối, tính cả “Số trang đầu tiên”. Chưa đối chiếu với Excel thật, cần kiểm khi có Excel.
- Vẽ lưới nhanh hơn khoảng 2 lần, bù cho việc có nhiều ô hơn vì ô nhỏ như Excel. Cách làm: mỗi khung hình chỉ đọc bố cục DOM một lần, có bộ nhớ đệm phông và độ rộng chữ, `style()` không tạo đối tượng mới. Kết quả: trung vị 3 ms / khung khi cuộn sổ 2 triệu ô.
- Sửa lỗi tiềm ẩn ở bản Windows: khung xem trước in (iframe) làm tiến trình chính tưởng trang đang nạp lại. Hệ quả là tệp nhấp đúp lúc đó không được mở và lần đóng cửa sổ sau đó không hỏi lưu. Nay chỉ tính lần nạp lại của khung chính.

## 6g. Bản 0.5.1 (29/10/2026) — sửa #ERROR! khi mở tệp dự toán (OFFSET động, tính lặp)
- **Nguyên nhân:** tệp TMDT (THKP hạng mục R12/R17) dùng `FORECAST(x;OFFSET(K12;0;MATCH(...);1;2);…)`. HyperFormula chỉ nhận OFFSET khi số hàng/cột là hằng → lỗi phân tích → #ERROR! lan sang Chi phí chung, LT, tổng cộng (535 ô).
- **Cách sửa (src/core/offset.ts):** khi đưa vào bộ tính, OFFSET động đổi thành hàm nội bộ `XOFSn(…; khung phụ thuộc)` (volatile như Excel); khi hiển thị/lưu đổi ngược lại → người dùng và tệp .xlsx luôn thấy OFFSET. Khung phụ thuộc = ô gốc + các *vùng* trong đối số (bỏ ô đơn để không tạo vòng lặp giả), không chứa chính ô công thức.
- **Tính lặp:** tệp bật iterative calculation (vòng R11→H20→…→R12). Ô phụ thuộc vòng lặp nhưng bọc IFERROR (HyperFormula không báo #CYCLE!) nay cũng được đưa vào tập tính lặp (duyệt getCellDependents) → TT38!R18/S18… đúng như Excel.
- **Kiểm chứng:** 2.748/2.748 ô công thức trong tệp khớp giá trị Excel lưu sẵn (trước: 563 sai). tests/offset.test.ts (4 ca: đổi qua lại, tính & tính lại, chèn hàng/hoàn tác/lưu .xlsx, tính lặp qua IFERROR). vitest 173, e2e v05 120, run 38, dot2 21, dot3 32, mobile 32.

## 6h. Bản 0.6 (29/10/2026) — phím tắt giống hệt Excel 365 + tính năng đi kèm
- **Nguồn duy nhất `src/ui/keymap.ts`**: 180 mục theo trang Microsoft “Keyboard shortcuts in Excel” (tra 29/10/2026) + phím Excel chuẩn có trong chú thích Ribbon (Ctrl+F/H/G, Ctrl+[ ], Alt+;, Ctrl+\ …). Mỗi mục: phím hiển thị, tổ hợp chuẩn hoá theo vị trí phím (`e.code`, nên Ctrl+Shift+$ = Ctrl+Shift+4, không phụ thuộc bộ gõ), ngữ cảnh (nav/edit/any), trạng thái ok/later/no + lý do. Dùng chung cho bộ điều phối (`keyFor` 3 pha: global → nav → edit), bảng F1 trong app, bảng tài liệu (`tools/keytable.mjs`) và test e2e.
- **Lõi mới**: `core/find.ts` (Tìm/Thay ký tự đại diện * ? ~, Tìm tiếp/ngược, Thay tất cả một bước hoàn tác, bỏ qua thay thế tạo công thức lỗi; Đi tới Đặc biệt 14 loại; `toRuns` gộp ô rời thành khối), `core/flashfill.ts` (Điền Nhanh: tìm chương trình đồng thời trên mọi ví dụ — đoạn chữ cố định / cả ô / từ thứ i đếm từ đầu-cuối / chữ cái đầu × HOA/thường/Hoa Đầu; xếp hạng ít chữ cố định → áp được nhiều hàng), `formula.ts` `cycleRef` (F4) + `transposeFormula` (dán chuyển vị đổi hướng tham chiếu như Excel), `series.ts` `seriesValues` (tuyến tính, tăng trưởng, ngày/ngày làm việc/tháng như EDATE/năm, giá trị dừng).
- **Workbook**: `group()` gộp nhiều lệnh thành một bước hoàn tác; `writeCells`, `pasteSpecial` (12 kiểu + 4 phép tính + bỏ ô trống + chuyển vị + dán liên kết + lặp nguồn cho vùng đích bội số; chụp nguồn trước khi ghi; cấm với Cắt như Excel), `setBorders` (8 cạnh gồm chéo), `paintStyles` (Chổi định dạng, giữ liên kết đích), `fillSeries`, `lastCell`, `dependentsOf`, `evalFormula`, `setDefaultColWidth`. Kiểu ô thêm `ud vp rot shk hx lk hid ln bd.d bd.u` — vẽ trên lưới (xoay, chữ dọc, co vừa ô, Điền, Giữa qua vùng chọn, thụt lề, đường chéo), in, đọc/ghi .xlsx (gồm siêu kết nối, bảo vệ ô). Sửa kèm: hàng chỉ có định dạng (không có giá trị) trước đây bị mất khi mở .xlsx.
- **Lưới**: vùng chọn nhiều khối (`areas`, Ctrl+nhấp, Shift+F8, kết quả Đi tới Đặc biệt), Enter/Tab đi vòng và sang khối kế tiếp, `Selection.box` để ô hiện hành nằm trong vùng mà không làm co vùng (sửa lỗi cũ: Enter trong vùng chọn làm co vùng), F8 mở rộng, chế độ End, Ctrl+Shift+cuộn ngang.
- **Hộp thoại**: Định dạng Ô 6 thẻ (`ui/fmtcells.ts`, chỉ áp các trường đã đổi, Ctrl+Tab chuyển thẻ), Tìm & Thay thế không chặn trang tính (kéo di chuyển, Tìm Tất cả có danh sách), Đi tới / Đặc biệt, Dán Đặc biệt (phím chữ), Liên kết, Chuỗi, Đối số Hàm, Phân tích Nhanh, Bảng tạm Office (24 mục), bảng phím tắt F1 có ô tìm không dấu (`ui/dlg06.ts`). Ribbon: mọi mục “có ở bản 0.6” đã thành lệnh thật; bỏ chú thích phím sai (Ctrl+E ở Căn giữa, Ctrl+Shift+> <). Điều hướng Ribbon bằng bàn phím (Alt rồi mũi tên/Tab, Ctrl+←/→ giữa nhóm, Alt+↓ mở menu, Esc về lưới). Nhóm trang tính (Ctrl+Shift+PgDn/PgUp, tiêu đề “[Nhóm]”). Bản Windows: Ctrl+F9/Ctrl+F10 (IPC `xl:win`), mở liên kết bằng ứng dụng mặc định (IPC `xl:open-external`, chỉ http/https/mailto).
- **Kiểm chứng**: vitest 199 (tests/v06.test.ts 26 ca mới), e2e/v06.mjs bấm thật **164/164** phím trong phạm vi trên trình duyệt + 3 phím bản Windows ở desktop.mjs (35/35), v05 120, run 39, dot2 21, dot3 32, mobile 32.

## 6i. Bản 0.6.1 (29/10/2026) — giao diện theo ảnh Excel 365 của anh + chọn ngôn ngữ

- **Ribbon theo 8 ảnh chụp Excel 365:**
  - Trang đầu: bộ sưu tập Kiểu Ô ngay trên ribbon (6 kiểu + nút mở rộng 14 kiểu) và nhóm Phần bổ trợ.
  - Góc phải: nút Chú thích và Chia sẻ.
  - Chèn: Biểu mẫu, Mô hình 3D, SmartArt, Ảnh chụp màn hình, nhóm Điều khiển (Hộp kiểm), Bản đồ.
  - Bố trí Trang: Căn / Nhóm / Xoay.
  - Công thức: Dùng Gần đây (nhớ 10 hàm vừa chèn) và Hàm Khác. Bỏ nút Tính lặp (vẫn có trong Tùy chọn).
  - Dữ liệu: Từ Ảnh, Nguồn Gần đây, Kết nối Hiện có, Truy vấn & Kết nối, Thuộc tính, nhóm Kiểu Dữ liệu, Mô hình Dữ liệu, Phân tích Dữ liệu.
  - Xem lại: Hiệu suất (Kiểm tra Hiệu suất chạy thật), Ngôn ngữ, Thay đổi, đủ nút Chú thích, Hủy chia sẻ, Mực.
  - Xem: nhóm Dạng xem Trang tính, Dạng xem Tùy chỉnh, Dẫn hướng (ngăn liệt kê trang tính), Thước, Biểu tượng Kiểu Dữ liệu, **Ô Tiêu điểm** (tô dải hàng/cột ô hiện hành), Xem Cạnh nhau, Cuộn Đồng bộ, Đặt lại Vị trí, Chuyển Cửa sổ.
  - Thẻ mới **Tự động hóa** (Office Scripts, Power Automate).
  - Lệnh cần dịch vụ trực tuyến (OneDrive, Bing, Power Query, Office Store…) ghi rõ "ngoài phạm vi" thay vì hẹn phiên bản.
- **Tệp ▸ Trang đầu như Excel 365:**
  - Lời chào và hàng mẫu: Sổ trống, Dự toán xây dựng, Bảng tính khối lượng, Bảng chấm công, kèm "Thêm mẫu →". Trang Mới có thêm mẫu Theo dõi thu chi.
  - Thẻ Gần đây / Ưa thích (ghim ☆) / Được chia sẻ, ô Tìm tệp, cột Ngày sửa đổi dạng tương đối (Vừa xong, 5 phút trước, Hôm qua lúc…).
  - Liên kết Khôi phục sổ làm việc chưa lưu và Thêm sổ làm việc →.
  - Thanh trái thêm trang Chia sẻ (gửi .xlsx/PDF).
- **Thanh trạng thái:** mục "Trợ năng: Mọi thứ đều ổn", bấm vào để kiểm tra.
- **Ngôn ngữ (Tệp ▸ Tùy chọn ▸ Ngôn ngữ):**
  - Tiếng Việt (mặc định) hoặc English theo thuật ngữ Excel 365 bản Mỹ.
  - Từ điển `i18n-en.ts` (≈1.300 mục, 95 mẫu) + `i18n-extra.ts`.
  - Mẫu "nghiêm" (STRICT) chỉ dịch khi phần ghép cũng dịch được, để không ra chữ nửa Anh nửa Việt.
  - Nội dung ô, tên trang tính, tên tệp không dịch.
  - Công thức giữ dấu ; và dấu phẩy thập phân.
- **Kiểm chứng:**
  - `tests/v061.test.ts`: dịch, mẫu nghiêm, ngày tương đối, 4 mẫu không lỗi công thức, tổng dự toán đúng.
  - `e2e/i18n.mjs` (90 kiểm thử): có đủ nhóm và lệnh theo ảnh, bộ sưu tập Kiểu Ô, Ô Tiêu điểm, Dẫn hướng, Kiểm tra Hiệu suất, Backstage, tạo từ mẫu. Chế độ English quét mọi thẻ, 10 trang Backstage và 4 hộp thoại: **0 chữ tiếng Việt sót**.

## 6j. Bản 0.6.2 (29/10/2026) — key bản quyền khi cài đặt + app Quản lý Key + khởi động an toàn

**Mục tiêu:**
- Nhập key ngay trong bộ cài.
- Key có hai loại: gắn 1 máy, hoặc dùng mọi máy (có thể đặt hạn dùng).
- App riêng để sinh, lưu và tra key.
- Không cần mạng, không làm giả được key.

**Thiết kế key:**
- Gồm 19 byte nội dung và chữ ký Ed25519 (64 byte), viết bằng Base32 Crockford thành 27 nhóm 5 ký tự.
- Nội dung: phiên bản, sản phẩm, loại key, số thứ tự (uint32), ngày cấp, ngày hết hạn (số ngày tính từ 01/01/2020; 0 = vĩnh viễn), băm máy 8 byte.
- Khóa bí mật chỉ nằm trong app Quản lý Key (mã hóa DPAPI theo tài khoản Windows). ExcelLike và bộ cài chỉ chứa khóa công khai (`electron/license-config.cjs`).
- **Mã máy** = SHA-512("ExcelLike|" + MachineGuid viết hoa), lấy 8 byte đầu, hiển thị dạng `XXXX-XXXX-XXXXX`.
  - Bộ cài (lic.exe) và ExcelLike (Node) tính ra cùng một mã; đã đối chiếu dưới Wine.
- Lõi chung `electron/license.cjs` cung cấp: makeKey, decodeKey, checkKey, machineHash/machineCode.
- Chống chỉnh lùi đồng hồ: ghi `lastSeen` vào license-state.json. Key có hạn mà đồng hồ máy lùi quá 2 ngày so với lần dùng trước thì báo lỗi.
- Thu hồi: danh sách `REVOKED` trong license-config.cjs, xuất ra từ app Quản lý Key. Có hiệu lực từ bản cài kế tiếp, vì key được kiểm ngay trên máy, không cần mạng.

**Bộ cài (NSIS):**
- Trang "Kích hoạt bản quyền" đặt sau trang Chào mừng.
  - Hiện Mã máy và ô dán key; điền sẵn key nếu đã cài trước đó.
  - Khi bấm Tiến, `lic.exe` (C + TweetNaCl, biên dịch bằng mingw) kiểm chữ ký, máy và hạn dùng, báo đúng lý do nếu key không hợp lệ.
- Key hợp lệ được lưu vào `%APPDATA%\ExcelLike\license.key`. Gỡ cài đặt vẫn giữ tệp này.
- Cài im lặng: `ExcelLike-Setup.exe /S /KEY=…`. Key sai thì thoát với mã 3 và không cài.

**ExcelLike khi mở (bản Windows):**
- Kiểm key ngay khi mở. Chưa hợp lệ thì hiện màn Kích hoạt chặn mọi thao tác (Mã máy, nút Chép, ô dán key).
- Tiến trình chính từ chối lưu, in, xuất PDF khi chưa kích hoạt.
- Tệp ▸ Tài khoản ▸ Bản quyền: xem thông tin key và Đổi key.
- Key còn ≤ 14 ngày thì nhắc gia hạn.
- Bản web/PWA không cần key.

**Khởi động an toàn** (sửa lỗi mở ra cửa sổ trống trên máy anh Lâm; máy anh không có tệp nhật ký nên chưa rõ nguyên nhân):
- Mỗi lần mở đều ghi nhật ký `excellike-log.txt`, gồm cả lỗi trong trang và lỗi tiến trình đồ họa.
- Cờ `boot-pending`: lần mở trước không khởi động xong thì lần này vào chế độ an toàn (`?safe=1`): sổ trống, tiếng Việt, sổ cũ vẫn giữ.
- Sau 25 giây chưa khởi động xong: hỏi Chờ thêm / Chế độ an toàn / Tắt tăng tốc đồ họa / Đóng.
- Tiến trình đồ họa hỏng 2 lần: tự tắt tăng tốc đồ họa và mở lại.
- Lỗi khi dựng giao diện: hiện trang báo lỗi có nút Chế độ an toàn, không để màn hình trống.

**App Quản lý Key (keymgr/, ExcelLike Key Manager 1.0.0):**
- Lần đầu: nhập tệp khóa ký.
- Cấp key: khách hàng, đơn vị, liên hệ, loại key, Mã máy, thời hạn (vĩnh viễn / 30 ngày / 3–6 tháng / 1–2 năm / đến ngày cụ thể), ghi chú. Sau khi cấp: chép key, chép tin nhắn gửi khách, lưu .txt.
- Sổ key: thống kê theo trạng thái, tìm kiếm, lọc, chi tiết (sửa thông tin, thu hồi/bỏ thu hồi), xuất CSV, chép danh sách thu hồi.
- Kiểm tra key: đọc thông tin, kiểm chữ ký, đối chiếu với sổ.
- Cài đặt: sao lưu khóa ký, sao lưu/khôi phục (gộp) sổ key, mở thư mục dữ liệu. Tự sao lưu sổ hằng ngày, giữ 30 bản.

**Kiểm chứng:**
- `tests/license.test.ts` (11 test).
- `e2e/desktop-license.mjs` (18): màn kích hoạt, key sai / máy khác / hết hạn / đúng, tự kích hoạt từ key bộ cài, nhắc gia hạn, quá hạn khi mở, chế độ an toàn.
- `e2e/keymgr.mjs` (27).
- Bộ cài thật chạy dưới Wine: cài im lặng với key máy khác hoặc key rác thì thoát mã 3, không cài; key đúng thì cài xong và ghi license.key. Đã chụp màn trang Kích hoạt và hộp báo key sai.
- `lic.exe` dưới Wine: 6 loại key đều trả đúng mã lỗi.
- Khóa ký **không** nằm trong repo hay trong bộ cài nào. Bản sao lưu đặt ở tài liệu dự án `claude/ExcelLike-khoa-ky.md`.

## 7. Kiểm chứng cuối theo tiêu chí Bước 0

| # | Tiêu chí | Kết quả |
|---|---|---|
| T1 | Lưới ô, kéo điền, sao chép/dán, hoàn tác, cố định ngăn, đổi cỡ, nhiều sheet | Đạt — E2E Đợt 1 38/38 |
| T2 | Công thức trùng Excel | Đạt — bộ đối chiếu với giá trị Excel (unit test), ~420 hàm |
| T3 | Tính lại theo phụ thuộc, phát hiện vòng | Đạt — sửa 1 ô chỉ tính ô phụ thuộc (~5 ms) |
| T4 | Định dạng ô, viền, gộp, có điều kiện | Đạt — unit + E2E Đợt 2 21/21 |
| T5 | Sắp xếp, lọc, kiểm tra dữ liệu | Đạt |
| T6 | Biểu đồ cột/đường/tròn, tự cập nhật | Đạt — 7 loại, E2E Đợt 3 32/32 |
| T7 | PivotTable kéo-thả, Sum/Count/Average | Đạt — khớp kết quả Pivot của Excel |
| T8 | Mở/lưu .xlsx, CSV, PDF, in | Đạt — openpyxl đọc lại; PDF đúng khổ/hướng |
| T9 | 100.000 hàng × 20 cột | Sổ trong app: mở ~4 giây, cuộn trung vị ~5 ms/khung. Tệp .xlsx: 10.000 hàng ~2 giây; 30.000 hàng ~5 giây; 100.000 hàng mở ~24 giây, lưu ~29 giây (giới hạn của thư viện ExcelJS — ghi nhận là hạn chế) |
| T10 | PWA cài được, offline | Đạt — manifest + service worker, phông đóng gói kèm |
| T11 | UI hiện đại, sáng/tối, điện thoại dọc/ngang, bàn phím, WCAG AA | Đạt — Lighthouse Trợ năng 100, Chuẩn kỹ thuật 100; E2E điện thoại 32/32 (360×640, 390×844, 844×390, 640×360) |
| — | Bản Windows | E2E bản máy tính 21/21 (Electron); cài/gỡ thử không giao diện bằng Wine: tệp, lối tắt, registry đúng, gỡ sạch, giữ tệp người dùng |

Còn hạn chế v1: tệp .xlsx rất lớn mở/lưu chậm; PivotTable lưu ra .xlsx là bảng giá trị; hình ảnh, ghi chú, macro trong tệp mất nếu lưu đè (có báo trước); biểu đồ tán xạ/radar chưa có; bộ cài chưa ký số (SmartScreen sẽ cảnh báo); chưa thử trên máy Windows thật và điện thoại thật (gõ Unikey/EVKey, cài PWA).

## 7b. Trạng thái

- **28/10/2026: chốt bản v1 — ExcelLike 0.4.0** (bộ cài Windows `ExcelLike-Setup-0.4.0.exe`, bản Portable `.zip`, PWA). Việc tiếp theo: cài thử trên máy Windows thật và điện thoại thật; ghi lỗi gặp phải để làm bản 0.4.x.
- Hướng cải thiện sau v1 (chưa cam kết): mở/lưu .xlsx lớn nhanh hơn (tự ghi XML sheet thay ExcelJS), PivotTable “sống” trong .xlsx, biểu đồ tán xạ/kết hợp, giữ hình ảnh/ghi chú khi lưu, ký số bộ cài.

Bài học quy trình (rút ra khi làm 4 đợt):
- Review độc lập sau mỗi đợt đều tìm ra lỗi nặng mà test của chính mình bỏ sót (chèn mã qua tên phông, tham chiếu sheet bị dùng lại, Esc làm treo đóng cửa sổ) → giữ bước review độc lập, giao cho người review chạy thử thật (script, Wine, Electron) chứ không chỉ đọc mã.
- Đóng gói Windows từ Linux không cần Wine: electron-builder `--win dir` + resedit (icon/phiên bản) + makensis gốc (`Target amd64-unicode`); Wine64 chỉ dùng để chạy thử bộ cài 64-bit.
- Đo hiệu năng khi máy rảnh (đo lúc đang đóng gói song song cho số sai lệch).

## 8. Cách chạy (cập nhật khi có mã nguồn)

```
npm install
npm run dev          # chạy bản web
npm test             # unit test (Vitest)
npm run build        # bản PWA trong dist/
npm run e2e          # E2E + ảnh chụp (Playwright, cần bản build)
node tools/artifact.mjs   # gộp 1 file HTML chạy thử
npm run desktop      # chạy bản máy tính (Electron)
npm run build:win    # release/ExcelLike-Setup-x.y.z.exe + ExcelLike-Portable-x.y.z.zip (cần makensis)
npm run e2e:desktop  # kiểm thử bản máy tính (Linux: xvfb-run)
```
