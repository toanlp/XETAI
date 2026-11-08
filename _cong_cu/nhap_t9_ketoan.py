"""Script đã dùng để nhập tháng 9/2026 theo bảng tổng kết của kế toán (ảnh trong Thu/). Giữ lại làm mẫu cho các tháng sau, KHÔNG chạy lại (sẽ ghi đè THU/CHI)."""
import datetime as dt
import openpyxl

import os
SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "QuanLyXe_75A40008.xlsx")
D = lambda d: dt.datetime(2026, 9, d)
FIRST, LAST = 5, 1004

wb = openpyxl.load_workbook(SRC)

# ---------- THU ----------
ws = wb["THU"]
old = []
for r in range(FIRST, LAST + 1):
    if ws[f"A{r}"].value is not None:
        old.append((ws[f"A{r}"].value, ws[f"B{r}"].value, ws[f"C{r}"].value, ws[f"D{r}"].value, ws[f"F{r}"].value))
# giữ 20 chuyến Hương Giang (= 20 chuyến "Nam Đông (Huy)" trong bảng kế toán), bỏ dòng gộp 9 chuyến không rõ ngày
huy = [x for x in old if x[4] == "Trạm Hương Giang"]
assert len(huy) == 20, len(huy)
huy = [(a, b, c, d, "Trạm Hương Giang – Nam Đông (Huy) → Hào Hưng") for a, b, c, d, _ in huy]

hiep_kg = [(1, 35770), (3, 33940), (5, 32650), (5, 33810), (6, 34140), (8, 33330),
           (9, 34950), (17, 33940), (18, 32350), (25, 35810), (26, 32540), (29, 31090)]
ymy_kg = [(3, 24700), (4, 23650), (8, 26700), (14, 32990), (20, 30170),
          (22, 26000), (24, 26330), (29, 27600), (30, 30800)]
new = [(D(d), None, kg / 1000, 73000, "Nam Đông (Hiệp) → Hào Hưng") for d, kg in hiep_kg]
new += [(D(d), None, kg / 1000, 90000, "Chở ngoài – Nam Đông → Ý Mỹ") for d, kg in ymy_kg]
new += [(D(15), None, 22.11, 130000, "Chở ngoài – Đà Nẵng → Năng Lượng Lson"),
        (D(24), None, 22.31, 130000, "Chở ngoài – Đà Nẵng → Ý Mỹ")]
rows = sorted(huy + new, key=lambda x: (x[0], x[4]))
assert len(rows) == 43
for r in range(FIRST, LAST + 1):
    for c in "ABCDF":
        ws[f"{c}{r}"].value = None
for i, (a, b, c, d, f) in enumerate(rows):
    r = FIRST + i
    ws[f"A{r}"].value, ws[f"B{r}"].value, ws[f"C{r}"].value, ws[f"D{r}"].value, ws[f"F{r}"].value = a, b, c, d, f

# ---------- CHI ----------
ws = wb["CHI"]
chi = [
    # Dầu – cty đổ (Phát Huy xd 25)
    (D(6), "Dầu", "Dầu (cty đổ, Phát Huy xd 25) 294,12 L", 8100000, None),
    (D(6), "Dầu", "Dầu (cty đổ, Phát Huy xd 25) 254,18 L", 7000000, None),
    (D(15), "Dầu", "Dầu (cty đổ, Phát Huy xd 25) 344,105 L", 9800000, None),
    (D(21), "Dầu", "Dầu (cty đổ, Phát Huy xd 25) 273,888 L", 8200000, None),
    (D(26), "Dầu", "Dầu (cty đổ, Phát Huy xd 25) 180,387 L", 5500000, None),
    # Sửa xe
    (D(1), "Sửa chữa", "Cân lốp", 100000, None),
    (D(4), "Sửa chữa", "Mua dầu ty lái", 70000, None),
    (D(7), "Sửa chữa", "Làm lốp thay ruột mới", 600000, None),
    (D(18), "Sửa chữa", "3 cái dè, 2 bù lon", 390000, None),
    (D(18), "Sửa chữa", "Lông đền ốc", 50000, "Bảng kế toán không ghi ngày, xếp theo dòng trên"),
    (D(23), "Sửa chữa", "Hàn ốp bửng", 250000, None),
    (D(25), "Sửa chữa", "Mua két nước", 250000, None),
    (D(27), "Sửa chữa", "1 lóc phanh, 1 má phanh cong tán", 1100000, "Phát Huy trả"),
    (D(29), "Sửa chữa", "Làm 2 dè sau", 200000, None),
    (D(30), "Sửa chữa", "Mua dầu côn", 100000, None),
    # Khác – theo bảng kế toán, ghi gộp cả tháng vào 30/9
    (D(30), "Khác", "Lương tài xế Lý 43 chuyến (khoán 32×350k + 11×450k + lương cứng 3tr)", 19150000, None),
    (D(30), "Khác", "Đi đường – bồi dưỡng tài xế 43 chuyến", 5000000, None),
    (D(30), "Khác", "Đi đường – CA (3 chuyến Ý Mỹ)", 1300000, None),
    (D(30), "Khác", "Đi đường – vé (4 chuyến Ý Mỹ)", 590000, None),
    (D(30), "Khác", "CA (dì Lộc)", 1000000, None),
    (D(30), "Khác", "Phí qua trạm", 1770000, None),
    (D(30), "Khác", "Thuế T8 của cty (4,5% × thu vận chuyển 80.367.000)", 3616000, None),
    (D(30), "Khác", "Lãi tháng 9", 4200000, None),
]
chi.sort(key=lambda x: x[0])  # sort ổn định: giữ thứ tự Dầu/Sửa chữa/Khác trong cùng ngày
for r in range(FIRST, LAST + 1):
    for c in "ABCDE":
        ws[f"{c}{r}"].value = None
for i, row in enumerate(chi):
    r = FIRST + i
    for c, v in zip("ABCDE", row):
        ws[f"{c}{r}"].value = v

wb.save(SRC)
print("THU", len(rows), "dòng, tổng tấn", round(sum(x[2] for x in rows), 3),
      "doanh thu", round(sum(x[2] * x[3] for x in rows)))
print("CHI", len(chi), "dòng, tổng", sum(x[3] for x in chi))
