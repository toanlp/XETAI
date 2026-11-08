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

## Làm ở nhiều máy

- Mở dự án ở máy nào thì việc đầu tiên: **Source Control → Pull** (hoặc `git pull`).
- Làm xong: **Commit + Push**. Không sửa file Excel ở 2 máy cùng lúc (file Excel không gộp được).

## Hằng tuần (khoảng 5 phút)

1. Trên Zalo PC, mở nhóm chat và chọn các ảnh hóa đơn mới, chọn **Lưu về máy** vào `Chi/<năm-tháng>/anh/`.
2. Copy các tin nhắn có nói đến chi phí, dán vào **cuối** file `Chi/<năm-tháng>/tin_nhan.txt`.
3. Mở Claude và nhắn: **"cập nhật chi tuần này"**.
4. Mở file Excel, sheet CHI, xem các dòng tô **cam** (Ghi chú bắt đầu bằng "Cần kiểm tra"). Kiểm tra xong thì xóa chữ đó ở Ghi chú.
5. Muốn xem Claude vừa sửa gì: trong VS Code mở **Source Control** (`Ctrl+Shift+G`), bấm vào file trong `_xem_thay_doi/`.

## Đầu tháng (chốt tháng trước)

1. Bỏ file thu của kế toán vào thư mục `Thu/` và nhắn Claude: **"chốt tháng 10"**.
2. Xem sheet **TONG_HOP**: chọn năm và tháng ở ô vàng.
3. Claude soạn sẵn tin tóm tắt chi trong tháng, bạn gửi vào nhóm Zalo để tài xế xác nhận.

## File Excel (3 sheet)

- **TONG_HOP**: chọn năm, tháng ở ô vàng để xem thu, chi, lãi; bảng và biểu đồ 12 tháng.
- **THU**: mỗi chuyến 1 dòng (Ngày, Phiếu, Tấn, Đơn giá, Nguồn). Thành tiền tự tính.
- **CHI**: mỗi khoản 1 dòng (Ngày, Loại chi, Nội dung, Số tiền, Ghi chú). Loại chi gồm Dầu, Sửa chữa, Khác. Đăng kiểm, bảo hiểm... trả tháng nào thì ghi vào tháng đó, loại Khác.
