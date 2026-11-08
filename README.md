# README – Dự án quản lý thu chi xe 75A 40008

> **Dành cho Claude ở các phiên sau:** đọc file này trước khi làm bất cứ việc gì. Sau đó đọc tiếp `QUY_TAC_NHAP_LIEU.md`.
> Đến cuối mỗi phiên, cập nhật mục **6. Trạng thái hiện tại** và mục **7. Nhật ký thay đổi**.

---

## 1. Bối cảnh

- Chủ xe là **toanlp**, quản lý 1 xe tải biển số **75A 40008**.
- **Thu:** kế toán lập file Excel ghi sản lượng xuất bãi từng ngày, gồm Ngày, Phiếu, Số xe, Xuất bãi, Hàng trừ, Còn lại, Đơn giá, Thành tiền. Thu là cước chở hàng, nhưng **có thể có nhiều file** (ngoài trạm Hương Giang còn chuyến khác, chủ xe sẽ gửi bổ sung). Mỗi file thu nhập thêm vào THU, không thay thế file khác.
- **Chi:** tài xế gửi chi phí phát sinh vào **nhóm chat Zalo**, khi thì bằng tin nhắn chữ, khi thì bằng ảnh hóa đơn. Các khoản thường gặp: dầu, sửa xe, CSGT, BOT...
  - Tài xế **không** điền form và **không** làm theo cú pháp nào. Đừng đề xuất thay đổi cách tài xế gửi.
- **Bảng tổng kết của kế toán khác (từ tháng 9/2026):** một kế toán khác tổng kết cả tháng thành 1 bảng (ảnh): toàn bộ chuyến + dầu, sửa xe, đi đường, lương, thuế, lãi, CÒN LẠI. **Chủ xe nhận phần CÒN LẠI và đóng vai người xác nhận lại** bảng này: đối chiếu với file trạm/phiếu cân, tin nhắn + ảnh Zalo tài xế, số chủ xe tự ghi; đầu ra là **danh sách sai lệch gửi kế toán**. Bảng này là nguồn đầy đủ nhất của tháng; cách nhập ở QUY_TAC mục 9, ghi chú đối chiếu từng tháng ở `Doi_chieu/yyyy-mm.md`.
- **Các bên (chủ xe giải thích 08/11/2026):**
  - **Phát Huy** = công ty của **O Lý**; xe chạy dưới tên công ty này. O Lý **góp vốn chủ xe** cùng toanlp. Kế toán lập bảng tổng kết là kế toán bên O Lý. Trong bảng, "o lí chi" / "Phát Huy trả" / "cty đổ" đều là phía O Lý trả.
  - **Dì Lộc** = mẹ của toanlp. Tài xế ứng tiền qua dì Lộc.
  - **Tài xế:** anh Lý ("Lương A Lý 43ch"), khác O Lý.
  - Phần CÒN LẠI: tạm thời **không** cần tính chia giữa chủ xe và O Lý.
  - **Hiệp, Huy** = kế toán của 2 bãi nhập hàng (Hào Hưng). Bãi Huy = file trạm Hương Giang; bãi Hiệp có phiếu/file riêng.
  - **Chở ngoài** (Ý Mỹ, Đà Nẵng...) = chuyến không vào bãi Hiệp/Huy, bên mình vẫn thu.
  - **CA** = phí công an. **Vé** chưa rõ (có thể là vé cân xe), khác "Phí qua trạm".
  - **Lãi tháng** = lãi vay mua xe, không cố định (chủ xe gửi sau).
  - Bảng tổng kết thường có vào **ngày 5 tháng sau**. Chủ xe không có hạn chót, chỉ xác nhận kế toán làm đúng hay sai.
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
├── Thu/                     ← file thu, giữ nguyên tên kế toán đặt. Tháng 9 để thẳng trong Thu/; từ tháng 10 để trong Thu/yyyy-mm/ (bảng tổng kết kế toán, file bãi Hiệp, trạm Hương Giang, phiếu cân, lãi vay)
├── Doi_chieu/yyyy-mm.md     ← ghi chú đối chiếu với bảng tổng kết của kế toán, điểm cần kiểm tra
├── _cong_cu/xuat_csv.py     ← xuất các sheet ra CSV để xem thay đổi
├── _xem_thay_doi/*.csv      ← bản CSV của THU, CHI (chỉ cột nhập tay)
└── Chi/
    └── yyyy-mm/
        ├── anh/             ← ảnh hóa đơn lưu từ Zalo
        └── tin_nhan.txt     ← tin nhắn copy từ Zalo; dòng "=== ĐÃ NHẬP ĐẾN ĐÂY ===" đánh dấu chỗ đã xử lý
```

## 3. Cấu trúc file Excel `QuanLyXe_75A40008.xlsx`

**Chủ xe muốn file gọn (03/10/2026):** chỉ 3 sheet (TONG_HOP, THU, CHI), ít cột. Không thêm cột/sheet khi chưa hỏi chủ xe.
**Giữ kiểu trình bày** (chủ xe thích): font Arial; hàng 1 tiêu đề, hàng 2 ghi chú nghiêng, hàng 3 dòng TỔNG (nền xanh lá), hàng 4 tiêu đề cột (nền xanh đậm, chữ trắng), dữ liệu hàng 5–1004. **Chữ xanh dương = ô nhập tay, chữ đen = công thức, không ghi đè.** Khi sửa file phải giữ nguyên định dạng này; dựng lại thì dùng `_cong_cu/tao_file_gon.py`.

| Sheet | Cột | Ghi chú |
|---|---|---|
| TONG_HOP | C4 Năm, C5 Tháng (ô vàng), C6 kỳ báo cáo; THU B8:C11, CHI B13:D17 (Dầu 14, Sửa chữa 15, Khác 16, TỔNG CHI 17), LỢI NHUẬN 19, CHỈ SỐ 21–24; bảng 12 tháng F9:J22; biểu đồ F24 | Ẩn lưới. Toàn công thức, lọc theo khoảng ngày (`SUMIFS` với `>=đầu tháng`, `<đầu tháng sau`). Cột D = % doanh thu |
| THU | A Ngày, B Phiếu, C Tấn, D Đơn giá, **E Thành tiền (công thức)**, F Nguồn / Ghi chú | Tấn = xuất bãi − hàng trừ |
| CHI | A Ngày, B Loại chi (chọn Dầu / Sửa chữa / Khác), C Nội dung, D Số tiền, E Ghi chú | E3 đếm "Cần kiểm tra". Ghi chú bắt đầu bằng "Cần kiểm tra" thì dòng tô cam |

Từ khóa phân loại nằm trong `QUY_TAC_NHAP_LIEU.md` mục 6 (không còn sheet DANH_MUC). `_cong_cu/tao_file_gon.py` là script đã dựng file này, xem khi cần sửa công thức TONG_HOP.

**Cách thao tác kỹ thuật với file:**

- Sửa bằng `openpyxl`, mở theo chế độ mặc định để giữ nguyên công thức. **Không lưu file sau khi mở bằng `data_only=True`.**
- Python: `%LOCALAPPDATA%\Programs\Python\Python312\python.exe` (lệnh `python` trong PATH là lối tắt Store, không dùng được).
- **Kiểm tra công thức:** máy có Excel nên dùng COM qua PowerShell (`New-Object -ComObject Excel.Application`, ẩn) mở **bản sao trong scratchpad**, `CalculateFull()`, đọc `.Text` của ô kết quả. Không có LibreOffice nên không chạy được `recalc.py`.
- Trước khi ghi, kiểm tra file khóa `~$QuanLyXe_75A40008.xlsx`. Nếu có thì Excel đang mở file, nhắc chủ xe đóng lại. Sao lưu file vào scratchpad trước khi sửa.

**Quy trình xem thay đổi (chủ xe chọn ngày 03/10/2026):** chủ xe xem thay đổi bằng CSV + git trong VS Code (Source Control).

1. **Trước khi sửa Excel:** nếu git còn thay đổi chưa commit (tức là lần sửa trước chủ xe đã xem xong), commit lại với nội dung mô tả lần sửa đó. Như vậy thay đổi lần này sẽ hiện riêng.
2. Sửa Excel.
3. **Sau khi sửa:** chạy `_cong_cu/xuat_csv.py`, **không commit**, để chủ xe mở Source Control xem các dòng thêm/sửa trong `_xem_thay_doi/*.csv`.
4. Nếu sửa cấu trúc sheet (thêm cột, đổi dòng tiêu đề), cập nhật luôn bảng `SHEETS` trong `xuat_csv.py`.
- Thêm dòng mới ngay sau dòng cuối đang có Ngày ở cột A (từ hàng 5). Chỉ ghi giá trị, ô đã có sẵn định dạng. Nếu dữ liệu sắp vượt dòng 1004, kéo dài công thức THU!E, định dạng, danh sách chọn, định dạng có điều kiện của CHI và các vùng `$5:$1004` trong TONG_HOP.

## 4. Các yêu cầu chủ xe hay dùng

| Chủ xe nói | Claude làm |
|---|---|
| "cập nhật chi tuần này" / "cập nhật chi tháng X" | Làm theo `QUY_TAC_NHAP_LIEU.md` mục 2, rồi báo số khoản mới, tổng tiền và các dòng cần kiểm tra |
| "chốt tháng X" | Nhập thu từ `Thu/`, rà soát bất thường, soạn tin tóm tắt để gửi nhóm Zalo (mục 8 của QUY_TAC) |
| Gửi file đăng kiểm, bảo hiểm… cuối tháng | Nhập từng khoản vào CHI, loại Khác (QUY_TAC mục 7) |
| "tạo tháng mới" | Tạo `Chi/yyyy-mm/anh/` và `tin_nhan.txt`, copy nội dung từ file mẫu của tháng trước |
| Hỏi số liệu | Đọc file Excel (stage lên rồi đọc bằng `data_only=True`, chỉ đọc, không lưu) và trả lời |

## 5. ⚠️ Luôn cập nhật "skill" nhập liệu

`QUY_TAC_NHAP_LIEU.md` là bộ quy tắc sống, **phải được cập nhật liên tục**. Bất cứ khi nào gặp điều gì khác với quy tắc hiện có, Claude phải **cập nhật file quy tắc ngay trong phiên đó**, không chỉ xử lý cho xong lần đó. Ví dụ:

- Tài xế đổi cách viết, chẳng hạn "500" nay nghĩa là 500k, dùng từ lóng mới, hoặc viết tắt mới.
- Xuất hiện loại ảnh mới: hóa đơn điện tử, ảnh chụp màn hình chuyển khoản, phiếu viết tay, biên bản.
- Có loại chi mới. Làm theo QUY_TAC mục 6 (bảng từ khóa, danh sách chọn, dòng ở TONG_HOP).
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

*Cập nhật lần cuối: 08/11/2026*

- [x] Dựng thư mục, file Excel, hướng dẫn và quy tắc nhập liệu.
- [x] File Excel đã làm gọn (3 sheet).
- [x] **Tháng 9 đã nhập lại theo bảng tổng kết của kế toán (08/11/2026), CHƯA kiểm tra.** THU 43 chuyến (dòng 5–47), 1.394,28 tấn, 108.546.360 đ. CHI 23 dòng (5–27): Dầu 38.600.000, Sửa chữa 3.110.000, Khác 36.626.000, tổng 78.336.000. Lợi nhuận 30.210.360 = đúng "CÒN LẠI" của kế toán. Khác khớp bảng nhưng điểm cần kiểm tra ghi ở `Doi_chieu/2026-09.md`.
- [x] Không có chi cố định riêng: đăng kiểm, bảo hiểm... ghi vào CHI tháng phát sinh. Không theo dõi tạm ứng.
- [x] Đã cài Python 3.12 và openpyxl tại `%LOCALAPPDATA%\Programs\Python\Python312\python.exe`. Lệnh `python` trong PATH vẫn trỏ vào lối tắt Store, và máy không có `py`, nên luôn gọi Python theo đường dẫn đầy đủ.
- [ ] Chưa chạy thử với ảnh và tin nhắn Zalo thật của tháng 9.
- [ ] Chưa đặt tác vụ tự chạy hằng tuần. Dự định làm sau 2–3 tuần chạy ổn, vào tối Chủ nhật, và cần máy tính đang bật.

**Việc tiếp theo:** khi chủ xe muốn, kiểm tra tháng 9 theo `Doi_chieu/2026-09.md` mục 3. Tháng 10: chờ bảng tổng kết của kế toán và/hoặc chi thô của chủ xe. Câu hỏi đang mở: có tách loại Khác (36,6tr) thành Lương / Đi đường / Thuế / Lãi không.

## 7. Nhật ký thay đổi

| Ngày | Thay đổi | Lý do |
|---|---|---|
| 03/10/2026 | Khởi tạo dự án: cấu trúc thư mục, Excel 7 sheet, quy tắc nhập liệu v1 | Bắt đầu quản lý thu chi |
| 03/10/2026 | Sheet CHI chỉ ghi chi phí xe (Dầu, Sửa chữa, Lốp, Bảo dưỡng - nhớt, Rửa xe, Khác; BOT/CSGT/Ăn uống/Bốc xếp giữ trong DANH_MUC nhưng chưa dùng), khoản ngoài phạm vi thì báo chủ xe. Chi cố định và tạm ứng lấy từ file chủ xe gửi cuối tháng. QUY_TAC thêm mục 6–7, chốt tháng chuyển sang mục 8 | Chủ xe yêu cầu
| 03/10/2026 | Nhập 11 khoản chi tháng 9. QUY_TAC: thêm nguồn "chủ xe nhập thô", bỏ ngưỡng 5tr với dầu (thay bằng ngưỡng 12tr/lần). DANH_MUC: thêm từ khóa cân lốp, làm lốp, sửa xe, phanh, ốp, linh tinh | Chủ xe gửi danh sách chi tháng 9; dầu 5–10tr/lần là bình thường
| 03/10/2026 | Nhập thu tháng 9 (20 chuyến) từ file trạm Hương Giang. Thu có thể nhiều file/nguồn. Bỏ tính tiêu hao dầu; dầu tính theo ngày đổ. Thêm `_cong_cu/xuat_csv.py` + `_xem_thay_doi/` + git để chủ xe xem thay đổi trong VS Code | Chủ xe trả lời rà soát tháng 9 và muốn xem được thay đổi |
| 03/10/2026 | Gộp loại Lốp vào Sửa chữa (CHI + DANH_MUC). THU thêm cột K "Nguồn / Ghi chú"; thêm dòng 9 chuyến không rõ ngày (22,4 tấn, 13.000 đ/tấn) | Chủ xe: lốp thuộc sửa xe; có thu nguồn khác không rõ ngày |
| 03/10/2026 | Làm gọn Excel: 7 sheet → 4 (TONG_HOP, THU, CHI, CO_DINH). Bỏ DANH_MUC, TAM_UNG, HUONG_DAN; CHI còn 5 cột, THU còn 6 cột; bỏ cột Tháng (TONG_HOP lọc theo khoảng ngày). Loại chi: Dầu / Sửa chữa / Khác (lốp, nhớt, rửa xe vào Sửa chữa). "Cần kiểm tra" chuyển thành tiền tố ở Ghi chú | Chủ xe thấy file quá chi tiết, muốn gọn |
| 03/10/2026 | Bỏ luôn sheet CO_DINH (còn 3 sheet). Đăng kiểm, bảo hiểm... ghi vào CHI tháng phát sinh, loại Khác | Chủ xe: không có phí cố định, phát sinh tháng nào ghi tháng đó |
| 03/10/2026 | Dựng lại bản gọn với kiểu trình bày của bản cũ (Arial, tiêu đề/TỔNG/tiêu đề cột ở hàng 1–4, chữ xanh nhập tay, khối màu ở TONG_HOP, % DT, chỉ số /tấn) | Chủ xe chê bản gọn đầu tiên xấu |
| 08/11/2026 | Nhập lại toàn bộ tháng 9 theo bảng tổng kết của kế toán (THU 43 chuyến, CHI 23 dòng, thêm lương/đi đường/thuế/lãi vào Khác). Bỏ dòng gộp 9 chuyến không rõ ngày, bỏ dầu 30/9 9,9tr, thêm dầu 6/9 8,1tr. Lưu ảnh bảng vào `Thu/`, thêm `Doi_chieu/`, QUY_TAC mục 9 | Chủ xe gửi bảng của kế toán khác; chủ xe là người xác nhận, tháng 9 chỉ nhập chưa kiểm tra |
