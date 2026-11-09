"""Nhập tháng 10/2026 theo bảng tổng kết của kế toán (ảnh chủ xe gửi 09/11/2026).
Phiếu lấy từ file bãi Hiệp (THÁNG 10. 2026CL.xlsx) và trạm Hương Giang (bãi Huy).
Đã chạy 09/11/2026 – mẫu, KHÔNG chạy lại (sẽ nhập trùng).
"""
import datetime as dt, os
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "QuanLyXe_75A40008.xlsx")
D = lambda d, m=10: dt.datetime(2026, m, d)

HIEP = "Nam Đông (Hiệp) → Hào Hưng"
HUY = "Trạm Hương Giang – Nam Đông (Huy) → Hào Hưng"
# (ngày, phiếu, kg theo bảng kế toán, cước đ/kg, nguồn)
trips = [
    (3, 4, 32510, 73, HIEP), (5, 8, 33550, 73, HIEP), (6, 11, 33990, 73, HIEP),
    (14, 16, 33150, 73, HIEP), (14, 18, 30190, 73, HIEP), (15, 20, 34240, 73, HIEP),
    (16, 23, 31270, 73, HIEP), (17, 26, 31430, 73, HIEP), (21, 36, 32000, 73, HIEP),
    (22, 37, 30880, 73, HIEP), (23, 40, 35060, 73, HIEP),
    (26, 44, 29960, 73, HIEP + "; kế toán ghi 29,96 t, file bãi Hiệp 26,96 t"),
    (1, 1, 33250, 73, HUY), (1, 2, 33240, 73, HUY), (1, 3, 34970, 73, HUY),
    (2, 6, 34780, 73, HUY), (3, 8, 32060, 73, HUY), (3, 9, 33620, 73, HUY),
    (4, 11, 34040, 73, HUY), (5, 14, 34080, 73, HUY), (6, 16, 36030, 73, HUY),
    (8, 17, 37360, 73, HUY), (9, 19, 36050, 73, HUY), (10, 21, 34950, 73, HUY),
    (11, 22, 33200, 73, HUY), (12, 24, 34120, 73, HUY), (12, 26, 34350, 73, HUY),
    (13, 27, 35950, 73, HUY), (14, 30, 36620, 73, HUY), (15, 33, 34590, 73, HUY),
    (16, 34, 33750, 73, HUY), (18, 36, 34560, 73, HUY), (19, 38, 32950, 73, HUY),
    (19, 40, 34450, 73, HUY), (21, 44, 34050, 73, HUY), (22, 48, 31030, 73, HUY),
    (24, 52, 32420, 73, HUY), (24, 54, 34330, 73, HUY), (25, 56, 34020, 73, HUY),
    (25, 57, 35510, 73, HUY),
    (26, 58, 37420, 73, HUY + "; kế toán ghi 37,42 t, file trạm 37,43 t"),
    (2, None, 20970, 100, "Chở ngoài – Nam Đông → Ý Mỹ"),
    (7, None, 25700, 100, "Chở ngoài – Nam Đông → Ý Mỹ"),
    (13, None, 31980, 100, "Chở ngoài – Nam Đông → Ý Mỹ"),
    (17, None, 26600, 100, "Chở ngoài – Nam Đông → Ý Mỹ"),
    (20, None, 23390, 100, "Chở ngoài – Bến Ván → Ý Mỹ"),
    (23, None, 27560, 100, "Chở ngoài – Nam Đông → Ý Mỹ"),
]

A = "Ảnh: 2026-10/anh/"
CK = "Cần kiểm tra: "
dau = "Dầu (cty đổ, Phát Huy xd 25) "
# (ngày, loại, nội dung, số tiền, ghi chú)
chi = [
    (D(1), "Dầu", dau + "324,697 L – đổ ngày 30/9", 9900000,
     "Đổ 30/9, kế toán tính vào T10. Ảnh: 2026-09/anh/2026-09-30_bom_dau_9tr9.png"),
    (D(3), "Dầu", dau + "277,686 L", 8250000, A + "2026-10-03_bom_dau.png"),
    (D(9), "Dầu", dau + "267,862 L", 7800000, A + "2026-10-09_bom_dau.png"),
    (D(14), "Dầu", dau + "137,363 L", 4000000, A + "2026-10-14_bom_dau.png"),
    (D(17), "Dầu", dau + "321,18 L", 9250000, A + "2026-10-17_bom_dau.png"),
    (D(22), "Dầu", dau + "290,6 L", 8500000, A + "2026-10-22_bom_dau.png"),
    (D(9), "Sửa chữa", "2 bộ thụt (Định Loan 1.700.000) + tiền xe thồ 150.000", 1850000,
     A + "2026-10-09_dinh_loan_bo_thut.png"),
    (D(9), "Sửa chữa", "Vá lốp, thay ruột + yếm mới (An Tây)", 800000, A + "2026-10-09_an_tay_va_lop.png"),
    (D(9), "Sửa chữa", "Tiện lại láp sau + công tháo ráp (Gara Tây La Sơn)", 600000,
     "Phát Huy trả. " + A + "2026-10-09_gara_tay_lap_sau.png"),
    (D(11), "Sửa chữa", "2 nhông đầu cầu giữa + cước vận chuyển + công tháo ráp (Gara Tây La Sơn)", 5800000,
     CK + "trên 5tr, đã khớp hóa đơn gara. Phát Huy trả. " + A + "2026-10-11_gara_tay_nhong_cau.png"),
    (D(12), "Sửa chữa", "Vá lốp lưu động, thay ruột mới (An Tây)", 850000, A + "2026-10-12_an_tay_va_lop.png"),
    (D(16), "Sửa chữa", "Mua 5 lít dầu cầu lỏng", 280000, "Không có ảnh, theo bảng kế toán"),
    (D(20), "Sửa chữa", "Làm 1 lốp, thay 1 yếm", 300000, A + "2026-10-20_lam_lop.png"),
    (D(22), "Sửa chữa", "Công thay lốp mới (An Tây)", 200000,
     "Kế toán ghi 22/10, hóa đơn ngày 23/10. " + A + "2026-10-23_an_tay_thay_lop.png"),
    (D(27), "Sửa chữa", "Mua 2 cái lốp", 13600000,
     CK + "trên 5tr, kế toán ghi 'đã có hóa đơn' nhưng chưa thấy ảnh. Phát Huy trả (chưa trả ngày 05/11)"),
    (D(31), "Khác", "Lương tài xế Lý 47 chuyến (khoán 41×350k + 6×450k + lương cứng 3tr)", 20050000, None),
    (D(31), "Khác", "Đi đường – bồi dưỡng tài xế 47 chuyến", 4700000, None),
    (D(31), "Khác", "Đi đường – CA (4/10 2.000.000; 16/10 500.000 + 300.000; 25/10 300.000)", 3100000,
     CK + "CA 4/10 2.000.000 lớn bất thường, không có biên bản"),
    (D(31), "Khác", "CA (dì Lộc)", 1000000, None),
    (D(31), "Khác", "Phí qua trạm", 1560000, None),
    (D(31), "Khác", "Thuế của cty (4,5% × thu vận chuyển 101.176.000)", 4552000, None),
    (D(31), "Khác", "Lãi tháng 10", 4000000, CK + "chủ xe đối chiếu với số lãi vay thật"),
    (D(2, 11), "Dầu", "Dầu 72,751 L (sổ nợ cây xăng)", 2200000, "Ảnh: 2026-11/anh/2026-11-02_bom_dau.png"),
]


def last_row(ws):
    r = 5
    while ws[f"A{r}"].value is not None:
        r += 1
    return r


def main():
    wb = openpyxl.load_workbook(SRC)
    ws = wb["THU"]
    r = last_row(ws)
    for d, p, kg, cuoc, src in sorted(trips, key=lambda t: (t[0], t[4].startswith("Chở"))):
        ws[f"A{r}"] = D(d); ws[f"B{r}"] = p; ws[f"C{r}"] = kg / 1000
        ws[f"D{r}"] = cuoc * 1000; ws[f"F{r}"] = src
        r += 1
    ws = wb["CHI"]
    r = last_row(ws)
    for d, loai, nd, tien, gc in sorted(chi, key=lambda c: c[0]):
        ws[f"A{r}"] = d; ws[f"B{r}"] = loai; ws[f"C{r}"] = nd; ws[f"D{r}"] = tien; ws[f"E{r}"] = gc
        r += 1
    wb.save(SRC)


if __name__ == "__main__":
    main()
