"""Xuất các sheet dữ liệu của QuanLyXe_75A40008.xlsx ra _xem_thay_doi/*.csv
để xem thay đổi trong VS Code (Source Control). Chỉ xuất cột nhập tay, bỏ cột công thức.

Chạy: "%LOCALAPPDATA%\\Programs\\Python\\Python312\\python.exe" _cong_cu\\xuat_csv.py
"""
import csv, datetime, os
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "QuanLyXe_75A40008.xlsx")
OUT = os.path.join(ROOT, "_xem_thay_doi")

# sheet: (hàng tiêu đề, dòng dữ liệu cuối, các cột xuất)
SHEETS = {
    "THU":         (4, 1004, "ABCDEFHK"),
    "CHI":         (4, 1004, "ABDEGHIJKLMNOP"),
    "CHI_CO_DINH": (4, 34,   "ABCDEG"),
    "TAM_UNG":     (4, 504,  "ACDE"),
    "DANH_MUC":    (4, 19,   "ABCEFGH"),
}


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def main():
    wb = openpyxl.load_workbook(SRC)
    os.makedirs(OUT, exist_ok=True)
    for name, (hdr, last, cols) in SHEETS.items():
        ws = wb[name]
        rows = [["Dòng"] + [fmt(ws[f"{c}{hdr}"].value) for c in cols]]
        for r in range(hdr + 1, last + 1):
            vals = [fmt(ws[f"{c}{r}"].value) for c in cols]
            if any(vals):
                rows.append([str(r)] + vals)
        # utf-8-sig để Excel mở đúng tiếng Việt
        with open(os.path.join(OUT, f"{name}.csv"), "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rows)
        print(f"{name}: {len(rows) - 1} dòng")


if __name__ == "__main__":
    main()
