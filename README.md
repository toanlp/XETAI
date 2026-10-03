# README – Dự án quản lý thu chi xe 75A 40008

> **Dành cho Claude ở các phiên sau:** đọc file này trước khi làm bất cứ việc gì. Sau đó đọc tiếp `QUY_TAC_NHAP_LIEU.md`.
> Đến cuối mỗi phiên, cập nhật mục **6. Trạng thái hiện tại** và mục **7. Nhật ký thay đổi**.

---

## 1. Bối cảnh

- Chủ xe là **toanlp**, quản lý 1 xe tải biển số **75A 40008**.
- **Thu:** kế toán lập file Excel ghi sản lượng xuất bãi từng ngày, gồm Ngày, Phiếu, Số xe, Xuất bãi, Hàng trừ, Còn lại, Đơn giá, Thành tiền. Thu là cước chở hàng, nhưng **có thể có nhiều file** (ngoài trạm Hương Giang còn chuyến khác, chủ xe sẽ gửi bổ sung). Mỗi file thu nhập thêm vào THU, không thay thế file khác.
- **Chi:** tài xế gửi chi phí phát sinh vào **nhóm chat Zalo**, khi thì bằng tin nhắn chữ, khi thì bằng ảnh hóa đơn. Các khoản thường gặp: dầu, sửa xe, CSGT, BOT...
  - Tài xế **không** điền form và **không** làm theo cú pháp nào. Đừng đề xuất thay đổi cách tài xế gửi.
- **Mục tiêu:** mỗi tháng biết được thu bao nhiêu, chi bao nhiêu, lãi bao nhiêu.
- **Cách làm đã chọn:** chủ xe lưu ảnh và tin nhắn Zalo vào thư mục. Claude đọc, phân loại và ghi vào Excel. Chủ xe chỉ duyệt lại các dòng "Cần kiểm tra".

## 2. Cấu trúc thư mục

Thư mục nằm trên máy chủ xe (Windows), đường dẫn `C:\Users\Admin\Documents\Xe 75A-40008\`

```
Xe 75A-40008/
├── README.md                ← file này
├── QUY_TAC_NHAP_LIEU.md     ← "skill" nhập liệu, LUÔN cập nhật khi có kiểu dữ liệu mới
├── HUONG_DAN.md             ← hướng dẫn cho chủ xe (người đọc)
├── QuanLyXe_75A40008.xlsx   ← file dữ liệu chính
├── Thu/                     ← file thu kế toán gửi, giữ nguyên tên kế toán đặt (vd "trạm hương giang tháng 9.2026.xlsx")
├── _cong_cu/xuat_csv.py     ← xuất các sheet ra CSV để xem thay đổi
├── _xem_thay_doi/*.csv      ← bản CSV của THU, CHI, CHI_CO_DINH, TAM_UNG, DANH_MUC (chỉ cột nhập tay)
└── Chi/
    └── yyyy-mm/
        ├── anh/             ← ảnh hóa đơn lưu từ Zalo
        └── tin_nhan.txt     ← tin nhắn copy từ Zalo; dòng "=== ĐÃ NHẬP ĐẾN ĐÂY ===" đánh dấu chỗ đã xử lý
```

## 3. Cấu trúc file Excel `QuanLyXe_75A40008.xlsx`

**Quy ước màu chữ:** xanh dương là ô nhập tay, đen là công thức. **Không được ghi đè công thức.**

| Sheet | Hàng tiêu đề | Dữ liệu | Cột nhập tay | Cột công thức |
|---|---|---|---|---|
| TONG_HOP | – | – | C4 = Năm, C5 = Tháng | Phần còn lại; bảng 12 tháng ở F9:J22, biểu đồ đặt tại F25 |
| THU | 4 | 5–1004 | A STT, B Ngày, C Phiếu, D Số xe, E Xuất bãi, F Hàng trừ, H Đơn giá | G Còn lại, I Thành tiền, J Tháng |
| CHI | 4 | 5–1004 | A STT, B Ngày, D Số xe, E Loại chi, G Nội dung, H Số lít, I Đơn giá, J Số tiền, K Km, L Người chi, M Nguồn, N File chứng từ, O Trạng thái, P Ghi chú | C Tháng, F Nhóm |
| CHI_CO_DINH | 4 | 5–34 | A–E, G | F Tiền/tháng (khoản theo năm thì chia 12) |
| TAM_UNG | 4 | 5–504 | A Ngày, C Số tiền, D Hình thức, E Ghi chú | B Tháng |
| DANH_MUC | 4 | 5–19 | A Loại chi, B Nhóm, C Từ khóa; E–H là các danh sách chọn | – |
| HUONG_DAN | – | – | – | – |

**Cách thao tác kỹ thuật với file:**

- Sửa bằng `openpyxl`, mở theo chế độ mặc định để giữ nguyên công thức. **Không lưu file sau khi mở bằng `data_only=True`.**
- Python: `%LOCALAPPDATA%\Programs\Python\Python312\python.exe` (lệnh `python` trong PATH là lối tắt Store, không dùng được). Máy không có LibreOffice nên không chạy được `recalc.py`; công thức sẽ tự tính khi chủ xe mở bằng Excel.
- Trước khi ghi, kiểm tra file khóa `~$QuanLyXe_75A40008.xlsx`. Nếu có thì Excel đang mở file, nhắc chủ xe đóng lại. Sao lưu file vào scratchpad trước khi sửa.

**Quy trình xem thay đổi (chủ xe chọn ngày 03/10/2026):** chủ xe xem thay đổi bằng CSV + git trong VS Code (Source Control).

1. **Trước khi sửa Excel:** nếu git còn thay đổi chưa commit (tức là lần sửa trước chủ xe đã xem xong), commit lại với nội dung mô tả lần sửa đó. Như vậy thay đổi lần này sẽ hiện riêng.
2. Sửa Excel.
3. **Sau khi sửa:** chạy `_cong_cu/xuat_csv.py`, **không commit**, để chủ xe mở Source Control xem các dòng thêm/sửa trong `_xem_thay_doi/*.csv`.
4. Nếu sửa cấu trúc sheet (thêm cột, đổi dòng tiêu đề), cập nhật luôn bảng `SHEETS` trong `xuat_csv.py`.
- Thêm dòng mới ngay sau dòng cuối đang có dữ liệu ở cột B. Nếu dữ liệu sắp vượt dòng 1004, cần kéo dài công thức và vùng danh sách chọn xuống dưới.

## 4. Các yêu cầu chủ xe hay dùng

| Chủ xe nói | Claude làm |
|---|---|
| "cập nhật chi tuần này" / "cập nhật chi tháng X" | Làm theo `QUY_TAC_NHAP_LIEU.md` mục 2, rồi báo số khoản mới, tổng tiền và các dòng cần kiểm tra |
| "chốt tháng X" | Nhập thu từ `Thu/`, rà soát bất thường, soạn tin tóm tắt để gửi nhóm Zalo (mục 8 của QUY_TAC) |
| Gửi file chi cố định / tạm ứng cuối tháng | Cập nhật CHI_CO_DINH và TAM_UNG theo QUY_TAC mục 7 |
| "tạo tháng mới" | Tạo `Chi/yyyy-mm/anh/` và `tin_nhan.txt`, copy nội dung từ file mẫu của tháng trước |
| Hỏi số liệu | Đọc file Excel (stage lên rồi đọc bằng `data_only=True`, chỉ đọc, không lưu) và trả lời |

## 5. ⚠️ Luôn cập nhật "skill" nhập liệu

`QUY_TAC_NHAP_LIEU.md` là bộ quy tắc sống, **phải được cập nhật liên tục**. Bất cứ khi nào gặp điều gì khác với quy tắc hiện có, Claude phải **cập nhật file quy tắc ngay trong phiên đó**, không chỉ xử lý cho xong lần đó. Ví dụ:

- Tài xế đổi cách viết, chẳng hạn "500" nay nghĩa là 500k, dùng từ lóng mới, hoặc viết tắt mới.
- Xuất hiện loại ảnh mới: hóa đơn điện tử, ảnh chụp màn hình chuyển khoản, phiếu viết tay, biên bản.
- Có loại chi mới. Khi đó thêm vào DANH_MUC trong Excel và bổ sung từ khóa.
- Chủ xe sửa lại một dòng Claude đã nhập. Phải hỏi hoặc tự rút ra lý do, rồi ghi thành quy tắc.
- Đổi tài xế, có thêm xe, hoặc đổi cách lưu ảnh và tin nhắn.
- Kế toán đổi mẫu file thu (thứ tự cột, thêm cột).

**Cách cập nhật:**

1. Sửa đúng mục liên quan trong `QUY_TAC_NHAP_LIEU.md`, ví dụ bảng viết tắt hay điều kiện "Cần kiểm tra".
2. Ghi một dòng vào mục **7. Nhật ký thay đổi** bên dưới: ngày, thay đổi gì, vì sao.
3. Nói cho chủ xe biết trong câu trả lời: "Đã cập nhật quy tắc: …"
4. Nếu thay đổi làm đổi cấu trúc Excel, cập nhật luôn bảng ở mục 3 của README này.

Khi không chắc quy tắc mới có đúng không, **hỏi chủ xe trước**, chưa ghi thành quy tắc.

## 6. Trạng thái hiện tại

*Cập nhật lần cuối: 03/10/2026*

- [x] Dựng thư mục, file Excel, hướng dẫn và quy tắc nhập liệu.
- [x] Đã xóa dòng ví dụ. CHI có 11 dòng chi tháng 9 (dòng 5–15, tổng 43.090.000 đ) do chủ xe nhập thô, chưa có ảnh chứng từ.
- [x] Sheet THU có 2 dòng thật của tháng 9 (ngày 01/09 và 02/09) do chủ xe cung cấp. Khi có file thu đầy đủ của tháng 9 thì thay thế.
- [ ] CHI_CO_DINH và TAM_UNG đang trống. Cuối tháng chủ xe sẽ gửi file riêng để cập nhật, chưa biết mẫu file.
- [x] Đã cài Python 3.12 và openpyxl tại `%LOCALAPPDATA%\Programs\Python\Python312\python.exe`. Lệnh `python` trong PATH vẫn trỏ vào lối tắt Store, và máy không có `py`, nên luôn gọi Python theo đường dẫn đầy đủ.
- [ ] Chưa chạy thử với ảnh và tin nhắn Zalo thật của tháng 9.
- [ ] Chưa đặt tác vụ tự chạy hằng tuần. Dự định làm sau 2–3 tuần chạy ổn, vào tối Chủ nhật, và cần máy tính đang bật.

**Việc tiếp theo:** chạy thử tháng 9 với dữ liệu Zalo thật, sau đó tinh chỉnh `QUY_TAC_NHAP_LIEU.md` theo cách tài xế thực tế gửi.

## 7. Nhật ký thay đổi

| Ngày | Thay đổi | Lý do |
|---|---|---|
| 03/10/2026 | Khởi tạo dự án: cấu trúc thư mục, Excel 7 sheet, quy tắc nhập liệu v1 | Bắt đầu quản lý thu chi |
| 03/10/2026 | Sheet CHI chỉ ghi chi phí xe (Dầu, Sửa chữa, Lốp, Bảo dưỡng - nhớt, Rửa xe, Khác; BOT/CSGT/Ăn uống/Bốc xếp giữ trong DANH_MUC nhưng chưa dùng), khoản ngoài phạm vi thì báo chủ xe. Chi cố định và tạm ứng lấy từ file chủ xe gửi cuối tháng. QUY_TAC thêm mục 6–7, chốt tháng chuyển sang mục 8 | Chủ xe yêu cầu
| 03/10/2026 | Nhập 11 khoản chi tháng 9. QUY_TAC: thêm nguồn "chủ xe nhập thô", bỏ ngưỡng 5tr với dầu (thay bằng ngưỡng 12tr/lần). DANH_MUC: thêm từ khóa cân lốp, làm lốp, sửa xe, phanh, ốp, linh tinh | Chủ xe gửi danh sách chi tháng 9; dầu 5–10tr/lần là bình thường
