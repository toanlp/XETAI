"""09/11/2026: tách loại chi Khác thành 'Lương, đi đường' / 'Thuế, lãi' / 'Khác'. Đã chạy, KHÔNG chạy lại."""
import copy, os
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "QuanLyXe_75A40008.xlsx")
LOAI = ["Dầu", "Sửa chữa", "Lương, đi đường", "Thuế, lãi", "Khác"]
M = 'CHI!$A$5:$A$1004,">="&$C$6,CHI!$A$5:$A$1004,"<"&EDATE($C$6,1)'

wb = openpyxl.load_workbook(SRC)
ws = wb["TONG_HOP"]
ws.move_range("B16:D24", rows=2)          # Khác..CHỈ SỐ xuống 2 dòng
for r in (16, 17):                        # 2 dòng loại mới, kiểu như dòng Sửa chữa
    for col in "BCD":
        ws[f"{col}{r}"]._style = copy.copy(ws[f"{col}15"]._style)
for r, loai in zip(range(14, 19), LOAI):
    ws[f"B{r}"] = loai
    ws[f"C{r}"] = f'=SUMIFS(CHI!$D$5:$D$1004,CHI!$B$5:$B$1004,"{loai}",{M})'
    ws[f"D{r}"] = f'=IF(N($C$11)=0,"",C{r}/$C$11)'
ws["D19"] = '=IF(N($C$11)=0,"",C19/$C$11)'   # TỔNG CHI
ws["C21"] = "=C11-C19"                       # LỢI NHUẬN
ws["D21"] = '=IF(N($C$11)=0,"",C21/$C$11)'
ws["C24"] = '=IF(N(C10)=0,"",C19/C10)'
ws["C25"] = '=IF(N(C10)=0,"",C21/C10)'

chi = wb["CHI"]
chi.data_validations.dataValidation = []
dv = DataValidation(type="list", formula1='"' + ",".join(LOAI) + '"', allow_blank=True)
chi.add_data_validation(dv); dv.add("B5:B1004")
chi["A2"] = 'Mỗi khoản 1 dòng. Loại chi: ' + " / ".join(LOAI) + '. Ghi chú bắt đầu bằng "Cần kiểm tra" thì dòng tô cam.'
for r in range(5, 1005):
    if chi[f"B{r}"].value != "Khác":
        continue
    nd = chi[f"C{r}"].value or ""
    if nd.startswith(("Lương", "Đi đường", "CA", "Phí qua trạm")):
        chi[f"B{r}"] = "Lương, đi đường"
    elif nd.startswith(("Thuế", "Lãi")):
        chi[f"B{r}"] = "Thuế, lãi"
# chủ xe trả lời 09/11: biết việc mua 2 lốp; lãi đúng theo dư nợ
for r in range(5, 1005):
    e = chi[f"E{r}"].value or ""
    if chi[f"C{r}"].value == "Mua 2 cái lốp":
        chi[f"E{r}"] = "Phát Huy trả (chưa trả ngày 05/11). Chủ xe biết việc mua"
    if chi[f"C{r}"].value == "Lãi tháng 10":
        chi[f"E{r}"] = "≈ 0,75% × dư nợ 531.263.000"
    if chi[f"C{r}"].value == "Lãi tháng 9":
        chi[f"E{r}"] = "≈ 0,75% × dư nợ 561.473.000"
wb.save(SRC)
