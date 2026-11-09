"""Dựng lại QuanLyXe_75A40008.xlsx dạng gọn (3 sheet: TONG_HOP, THU, CHI), giữ kiểu trình bày của bản cũ
(Arial, tiêu đề + ghi chú + dòng TỔNG ở hàng 1–3, tiêu đề cột hàng 4, dữ liệu từ hàng 5; chữ xanh = nhập tay).
Chạy ngày 03/10/2026 để chuyển từ bản 7 sheet. Giữ lại để tham khảo cách dựng công thức TONG_HOP.

Dùng: python tao_file_gon.py <file_cu_7sheet.xlsx> <file_moi.xlsx>
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference

SRC, DST = sys.argv[1], sys.argv[2]
LOAI_CHI = ["Dầu", "Sửa chữa", "Lương, đi đường", "Thuế, lãi", "Khác"]
FIRST, LAST = 5, 1004  # dòng dữ liệu

# ---------- đọc dữ liệu cũ ----------
old = openpyxl.load_workbook(SRC)
o = old["CHI"]
chi = []
for r in range(5, 1005):
    if o[f"B{r}"].value is None:
        continue
    gc = []
    if o[f"O{r}"].value == "Cần kiểm tra":
        gc.append("Cần kiểm tra")
    if o[f"N{r}"].value:
        gc.append(f"Ảnh: {o[f'N{r}'].value}")
    note = o[f"P{r}"].value or ""
    if note and not note.startswith("Chủ xe nhập thô"):
        gc.append(note)
    chi.append((o[f"B{r}"].value, o[f"E{r}"].value, o[f"G{r}"].value, o[f"J{r}"].value, "; ".join(gc) or None))

o = old["THU"]
thu = []
for r in range(5, 1005):
    if o[f"B{r}"].value is None:
        continue
    nguon = o[f"K{r}"].value or ""
    if nguon.startswith("9 chuyến"):
        nguon = "9 chuyến không rõ ngày, tổng 22.400 kg"
    thu.append((o[f"B{r}"].value, o[f"C{r}"].value, o[f"E{r}"].value - (o[f"F{r}"].value or 0), o[f"H{r}"].value, nguon))

# ---------- kiểu (giống bản cũ) ----------
FONT = "Arial"
NAVY = "1F4E78"
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
F_HDR = Font(name=FONT, bold=True, color="FFFFFF")
F_IN = Font(name=FONT, color="0000FF")      # nhập tay
F_CALC = Font(name=FONT, color="000000")    # công thức
F_B = Font(name=FONT, bold=True)
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="DDEBF7")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FILL_LAI = PatternFill("solid", fgColor="C6E0B4")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_WARN = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
MONEY = "#,##0;\\(#,##0\\);\\-"
TON = "#,##0.00;\\(#,##0.00\\);\\-"
PCT = "0.0%;\\(0.0%\\);\\-"
DATE, MONTH = "dd/mm/yyyy", "mm/yyyy"

wb = openpyxl.Workbook()
normal = next(s for s in wb._named_styles if s.name == "Normal")
normal.font = Font(name=FONT, size=11)


def sheet_top(ws, title, note, headers, widths):
    ws["A1"].value, ws["A1"].font = title, F_TITLE
    ws["A2"].value, ws["A2"].font = note, F_NOTE
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[4].height = 31.5
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(4, i, h)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = f"A{FIRST}"


def fmt_col(ws, col, font, nf=None):
    for r in range(FIRST, LAST + 1):
        c = ws[f"{col}{r}"]
        c.font = font
        if nf:
            c.number_format = nf


# ---------- THU ----------
ws = wb.active
ws.title = "THU"
sheet_top(ws, "THU – SẢN LƯỢNG XUẤT BÃI",
          "Mỗi chuyến 1 dòng. Tấn = xuất bãi − hàng trừ. Cột Thành tiền tự tính.",
          ["Ngày", "Phiếu", "Tấn", "Đơn giá (đ/tấn)", "Thành tiền (đ)", "Nguồn / Ghi chú"],
          [12, 9, 11, 15, 16, 42])
for col, nf in (("A", DATE), ("B", None), ("C", TON), ("D", MONEY), ("F", None)):
    fmt_col(ws, col, F_IN, nf)
fmt_col(ws, "E", F_CALC, MONEY)
for r in range(FIRST, LAST + 1):
    ws[f"E{r}"].value = f'=IF(C{r}="","",C{r}*D{r})'
for i, row in enumerate(sorted(thu, key=lambda x: x[0]), FIRST):
    for col, v in zip("ABCDF", row):
        ws[f"{col}{i}"].value = v
ws["B3"].value, ws["B3"].font = "TỔNG", F_B
for col, nf in (("C", TON), ("E", MONEY)):
    c = ws[f"{col}3"]
    c.value, c.font, c.fill, c.number_format = f"=SUM({col}{FIRST}:{col}{LAST})", F_B, FILL_TOT, nf

# ---------- CHI ----------
ws = wb.create_sheet("CHI")
sheet_top(ws, "CHI – CHI PHÍ XE",
          "Mỗi khoản 1 dòng. Loại chi: Dầu / Sửa chữa / Khác. Ghi chú bắt đầu bằng \"Cần kiểm tra\" thì dòng tô cam.",
          ["Ngày", "Loại chi", "Nội dung", "Số tiền (đ)", "Ghi chú"],
          [12, 13, 32, 15, 45])
for col, nf in (("A", DATE), ("B", None), ("C", None), ("D", MONEY), ("E", None)):
    fmt_col(ws, col, F_IN, nf)
for i, row in enumerate(sorted(chi, key=lambda x: x[0]), FIRST):
    for col, v in zip("ABCDE", row):
        ws[f"{col}{i}"].value = v
ws["C3"].value, ws["C3"].font = "TỔNG", F_B
ws["C3"].alignment = Alignment(horizontal="right")
c = ws["D3"]
c.value, c.font, c.fill, c.number_format = f"=SUM(D{FIRST}:D{LAST})", F_B, FILL_TOT, MONEY
c = ws["E3"]
c.value = f'="Cần kiểm tra: "&COUNTIF(E{FIRST}:E{LAST},"Cần kiểm tra*")'
c.font = Font(name=FONT, bold=True, color="C00000")
dv = DataValidation(type="list", formula1='"' + ",".join(LOAI_CHI) + '"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"B{FIRST}:B{LAST}")
ws.conditional_formatting.add(f"A{FIRST}:E{LAST}",
                              FormulaRule(formula=[f'LEFT($E{FIRST},12)="Cần kiểm tra"'], fill=FILL_WARN))

# ---------- TONG_HOP ----------
ws = wb.create_sheet("TONG_HOP", 0)
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJ", (2, 26, 17, 9, 3, 11, 16, 16, 16, 11)):
    ws.column_dimensions[col].width = w
ws["B1"].value, ws["B1"].font = "TỔNG HỢP THU CHI – XE 75A 40008", Font(name=FONT, size=15, bold=True, color=NAVY)
ws["B2"].value, ws["B2"].font = "Chọn Năm và Tháng ở ô vàng. Mọi số liệu tự cập nhật.", F_NOTE
ws.row_dimensions[1].height = 19
for r, (lab, val) in enumerate((("Năm", 2026), ("Tháng", 9)), 4):
    ws[f"B{r}"].value, ws[f"B{r}"].font = lab, F_B
    c = ws[f"C{r}"]
    c.value, c.font, c.fill, c.border = val, Font(name=FONT, bold=True, color="0000FF"), FILL_IN, BOX
    c.alignment = Alignment(horizontal="center")
ws["B6"].value, ws["B6"].font = "Kỳ báo cáo", F_B
c = ws["C6"]
c.value, c.font, c.number_format, c.alignment = "=DATE(C4,C5,1)", F_B, MONTH, Alignment(horizontal="center")

K, KN = "$C$6", "EDATE($C$6,1)"
D_THU, D_CHI = f"THU!$A${FIRST}:$A${LAST}", f"CHI!$A${FIRST}:$A${LAST}"


def rng(sheet, col):
    return f"{sheet}!${col}${FIRST}:${col}${LAST}"


def in_month(dates, start=K, end=KN):
    return f'{dates},">="&{start},{dates},"<"&{end}'


def section(r, title, pct=True):
    for col in "BCD":
        ws[f"{col}{r}"].fill = FILL_SEC
    ws[f"B{r}"].value, ws[f"B{r}"].font = title, Font(name=FONT, bold=True, color=NAVY)
    if pct:
        ws[f"D{r}"].value, ws[f"D{r}"].font = "% DT", Font(name=FONT, size=9, bold=True, color=NAVY)


def line(r, label, formula, nf=MONEY, bold=False, fill=None, pct=False):
    ws[f"B{r}"].value = label
    ws[f"C{r}"].value, ws[f"C{r}"].number_format = formula, nf
    for col in "BC":
        ws[f"{col}{r}"].font = F_B if bold else Font(name=FONT)
    if pct:
        d = ws[f"D{r}"]
        d.value = f'=IF(N($C$11)=0,"",C{r}/$C$11)'
        d.number_format, d.font = PCT, Font(name=FONT, size=9, color="595959")
    if fill:
        for col in "BCD":
            ws[f"{col}{r}"].fill = fill


section(8, "THU", pct=False)
line(9, "Số chuyến (dòng)", f"=COUNTIFS({in_month(D_THU)})", "#,##0")
line(10, "Tổng tấn", f"=SUMIFS({rng('THU', 'C')},{in_month(D_THU)})", TON)
line(11, "DOANH THU", f"=SUMIFS({rng('THU', 'E')},{in_month(D_THU)})", bold=True, fill=FILL_TOT)

section(13, "CHI")
r = 14
for loai in LOAI_CHI:
    line(r, loai, f'=SUMIFS({rng("CHI", "D")},{rng("CHI", "B")},"{loai}",{in_month(D_CHI)})', pct=True)
    r += 1
R_TC = r
line(R_TC, "TỔNG CHI", f"=SUMIFS({rng('CHI', 'D')},{in_month(D_CHI)})", bold=True, fill=FILL_TOT, pct=True)

R_LN = R_TC + 2
line(R_LN, "LỢI NHUẬN", f"=C11-C{R_TC}", bold=True, fill=FILL_LAI, pct=True)
ws[f"C{R_LN}"].font = Font(name=FONT, size=12, bold=True)

R_CS = R_LN + 2
section(R_CS, "CHỈ SỐ", pct=False)
line(R_CS + 1, "Chi phí / tấn", f'=IF(N(C10)=0,"",C{R_TC}/C10)')
line(R_CS + 2, "Lãi / tấn", f'=IF(N(C10)=0,"",C{R_LN}/C10)')
line(R_CS + 3, "Khoản chi cần kiểm tra", f'=COUNTIFS({rng("CHI", "E")},"Cần kiểm tra*",{in_month(D_CHI)})', "0")
ws[f"C{R_CS + 3}"].font = Font(name=FONT, bold=True, color="C00000")

# bảng 12 tháng
ws["F8"].value, ws["F8"].font = "12 THÁNG NĂM ĐÃ CHỌN", Font(name=FONT, bold=True, color=NAVY)
for col, h in zip("FGHIJ", ("Tháng", "Doanh thu", "Tổng chi", "Lợi nhuận", "Tấn")):
    c = ws[f"{col}9"]
    c.value, c.font, c.fill, c.border, c.alignment = h, F_HDR, FILL_HDR, BOX, Alignment(horizontal="center")
for m in range(1, 13):
    rr = 9 + m
    s = f"F{rr}"
    ws[s].value = f"=DATE($C$4,{m},1)"
    ws[f"G{rr}"].value = f"=SUMIFS({rng('THU', 'E')},{in_month(D_THU, s, f'EDATE({s},1)')})"
    ws[f"H{rr}"].value = f"=SUMIFS({rng('CHI', 'D')},{in_month(D_CHI, s, f'EDATE({s},1)')})"
    ws[f"I{rr}"].value = f"=G{rr}-H{rr}"
    ws[f"J{rr}"].value = f"=SUMIFS({rng('THU', 'C')},{in_month(D_THU, s, f'EDATE({s},1)')})"
    for col, nf in zip("FGHIJ", (MONTH, MONEY, MONEY, MONEY, TON)):
        c = ws[f"{col}{rr}"]
        c.number_format, c.border, c.font = nf, BOX, Font(name=FONT)
ws["F22"].value, ws["F22"].font = "Cả năm", F_B
for col, nf in zip("GHIJ", (MONEY, MONEY, MONEY, TON)):
    c = ws[f"{col}22"]
    c.value, c.number_format, c.font, c.fill, c.border = f"=SUM({col}10:{col}21)", nf, F_B, FILL_TOT, BOX

ch = BarChart()
ch.type = "col"
ch.title = "Thu – Chi – Lợi nhuận theo tháng"
ch.y_axis.numFmt = "#,##0"
ch.y_axis.majorGridlines = None
ch.add_data(Reference(ws, min_col=7, max_col=9, min_row=9, max_row=21), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=6, min_row=10, max_row=21))
for s, color in zip(ch.series, ("2E75B6", "ED7D31", "70AD47")):
    s.graphicalProperties.solidFill = color
    s.graphicalProperties.line.solidFill = color
ch.legend.position = "b"
ch.height, ch.width = 7.5, 14.5
ws.add_chart(ch, "F24")

wb.active = 0
wb.save(DST)
print("THU", len(thu), "dong; CHI", len(chi), "dong; TONG CHI o dong", R_TC, "LOI NHUAN o dong", R_LN)
