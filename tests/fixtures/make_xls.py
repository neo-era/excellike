# Tạo tệp mẫu .xls (BIFF8) cho tests/xls.test.ts — chạy: python3 tests/fixtures/make_xls.py (cần: pip install xlwt)
import xlwt, datetime
wb = xlwt.Workbook(encoding='utf-8')
wb.set_colour_RGB(0x21, 0x12, 0x34, 0x56)  # đổi bảng màu: ô số 0x21 thành #123456
ws = wb.add_sheet('Dữ liệu', cell_overwrite_ok=True)
ph = wb.add_sheet('Phụ')
ph.visibility = 1  # trang ẩn
ph.write(0, 0, 21)

title = xlwt.easyxf('font: name Times New Roman, height 280, bold on, colour red; pattern: pattern solid, fore_colour yellow; align: horiz center, vert center, wrap on; borders: left thin, right medium, top double, bottom dashed')
ws.merge(0, 0, 0, 2)
ws.write(0, 0, 'Tiêu đề', title)
ws.write(0, 1, 'ẩn trong ô gộp', title)  # nội dung ô không phải ô đầu của vùng gộp (Excel vẫn giữ)
num = xlwt.easyxf('font: name Arial, height 200, italic on; pattern: pattern solid, fore_colour 0x21', num_format_str='#,##0.00')
date = xlwt.easyxf(num_format_str='dd/mm/yyyy')
ws.write(2, 0, 2, num)
ws.write(3, 0, 3.5, num)
ws.write(2, 1, xlwt.Formula('SUM(A3:A4)'))
ws.write(3, 1, xlwt.Formula('ROUND(519.81*0.25,3)'))
ws.write(4, 1, xlwt.Formula('A3&" x "&0.55*100'))
ws.write(5, 1, xlwt.Formula('IF(A3>1,"lớn","nhỏ")'))
ws.write(6, 1, xlwt.Formula("'Phụ'!A1*2"))
ws.write(7, 1, xlwt.Formula('1/0'))
ws.write(8, 1, xlwt.Formula('A4/2&""'))
ws.write(9, 1, True)
ws.write(10, 1, datetime.date(2026, 11, 17), date)
ws.write(11, 1, 'chữ thường')
ws.write(13, 1, None, xlwt.easyxf('font: bold on; align: horiz center', num_format_str='#,##0'))  # ô trống có định dạng sẵn
ws.row(20).height_mismatch = True
ws.row(20).height = 0  # hàng cao 0 → ẩn
ws.col(1).width = 20 * 256
ws.col(3).hidden = True
ws.row(1).height_mismatch = True
ws.row(1).height = 600  # 30pt
ws.row(5).hidden = True
ws.set_panes_frozen(True)
ws.set_horz_split_pos(2)
ws.set_vert_split_pos(1)
ws.portrait = False
ws.paper_size_code = 8  # A3
ws.fit_num_pages = 0
wb.save('tests/fixtures/basic.xls')
print('ok')
