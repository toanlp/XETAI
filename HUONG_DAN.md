# QUẢN LÝ THU CHI – XE 75A 40008

## Cấu trúc thư mục

```
Xe 75A-40008/
├── QuanLyXe_75A40008.xlsx   ← file chính (xem sheet TONG_HOP)
├── HUONG_DAN.md             ← file này
├── QUY_TAC_NHAP_LIEU.md     ← quy tắc Claude dùng khi nhập chi phí
├── Thu/                     ← file thu kế toán gửi hằng tháng
└── Chi/
    ├── 2026-09/
    │   ├── anh/             ← ảnh hóa đơn lưu từ Zalo
    │   └── tin_nhan.txt     ← tin nhắn chi phí copy từ Zalo
    └── 2026-10/ ...
```

Sang tháng mới, tạo thêm thư mục `Chi/<năm-tháng>/` có cấu trúc như trên, hoặc nhờ Claude tạo.

## Hằng tuần (khoảng 5 phút)

1. Trên Zalo PC, mở nhóm chat và chọn các ảnh hóa đơn mới, chọn **Lưu về máy** vào `Chi/<năm-tháng>/anh/`.
2. Copy các tin nhắn có nói đến chi phí, dán vào **cuối** file `Chi/<năm-tháng>/tin_nhan.txt`.
3. Mở Claude và nhắn: **"cập nhật chi tuần này"**.
4. Mở file Excel, sheet CHI, xem các dòng tô **cam** (Cần kiểm tra). Kiểm tra xong thì đổi thành *Đã xác nhận*.

## Đầu tháng (chốt tháng trước)

1. Bỏ file thu của kế toán vào thư mục `Thu/` và nhắn Claude: **"chốt tháng 10"**.
2. Xem sheet **TONG_HOP**: chọn năm và tháng ở ô vàng.
3. Claude soạn sẵn tin tóm tắt chi trong tháng, bạn gửi vào nhóm Zalo để tài xế xác nhận.

## Thiết lập một lần

- Sheet **CHI_CO_DINH**: điền số tiền lương cứng, trả góp, bảo hiểm, đăng kiểm, phí đường bộ, GPS.
- Sheet **TAM_UNG**: nếu có đưa tiền trước cho tài xế, ghi mỗi lần đưa.
- Sheet **CHI** dòng 5 là **dòng ví dụ**, xóa khi bắt đầu nhập thật.
