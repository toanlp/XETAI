# QUY TẮC NHẬP LIỆU CHI PHÍ (dành cho Claude)

Tài liệu này mô tả cách Claude đọc ảnh và tin nhắn Zalo rồi ghi vào sheet **CHI** của `QuanLyXe_75A40008.xlsx`.
Khi được nhờ "cập nhật chi", Claude làm đúng theo các bước dưới đây.

## 1. Nguồn dữ liệu

| Nguồn | Vị trí |
|---|---|
| Ảnh hóa đơn, phiếu dầu, biên bản | `Chi/<yyyy-mm>/anh/` |
| Tin nhắn chữ copy từ Zalo | `Chi/<yyyy-mm>/tin_nhan.txt` |
| Danh sách chủ xe gõ thô trong chat với Claude | Claude dán nguyên văn vào `Chi/<yyyy-mm>/tin_nhan.txt` (kèm dòng `# [ngày] Chủ xe nhập thô`) trước khi xử lý, để còn đối chiếu |
| File thu của kế toán | `Thu/` |

## 2. Các bước xử lý

1. **Đọc sheet CHI** để lấy danh sách file trong cột *File chứng từ* và các dòng đã có. Ảnh nào đã có tên trong cột này thì **bỏ qua**.
2. **Đọc từng ảnh mới** và lấy ra: ngày, số tiền, số lít, đơn giá, nơi phát sinh, km đồng hồ (nếu có).
3. **Đọc tin nhắn mới** trong `tin_nhan.txt` (phần nằm sau dòng `=== ĐÃ NHẬP ĐẾN ĐÂY ===`).
4. **Phân loại** theo cột *Từ khóa nhận diện* ở sheet DANH_MUC. Chỉ nhập các loại chi nằm trong phạm vi ở mục 6. Khoản nào có vẻ là chi cho xe nhưng không khớp từ khóa thì xếp vào loại `Khác` và đặt Trạng thái `Cần kiểm tra`. Khoản nào nằm ngoài phạm vi thì không nhập, chỉ báo lại cho chủ xe.
5. **Chống trùng:** khoản trong tin nhắn và khoản trong ảnh có cùng ngày (lệch tối đa ±1 ngày), cùng loại và cùng số tiền thì chỉ ghi **1 dòng**, cột Nguồn ghi `Cả hai`.
6. **Ghi vào CHI:** thêm dòng mới ngay sau dòng cuối đang có dữ liệu. Chỉ ghi vào các ô nhập tay (chữ xanh), không sửa công thức ở cột C và F.
7. **Đánh dấu đã nhập:** chèn lại dòng `=== ĐÃ NHẬP ĐẾN ĐÂY ===` vào cuối file `tin_nhan.txt`.
8. **Báo lại cho chủ xe:** số khoản mới, tổng tiền, các dòng `Cần kiểm tra` kèm lý do.

## 3. Cách hiểu số tiền viết tắt

| Tài xế viết | Hiểu là |
|---|---|
| 500k, 500 nghìn, 500 | 500.000 |
| 1tr2, 1.2tr, 1 triệu 2 | 1.200.000 |
| 2tr150 | 2.150.000 |
| 1,5 củ | 1.500.000 |

Số tiền dưới 1.000 mà không có đơn vị thì hiểu là **nghìn đồng**. Nếu không chắc, đặt Trạng thái `Cần kiểm tra`.

**Chủ xe nhập thô** theo dạng `<ngày> <nội dung> <số tiền>`, có thể đánh số hoặc không, có thể gõ sai chính tả (ví dụ "suawrxe" là "sửa xe"):
- Chỉ ghi ngày, ví dụ `26`, thì hiểu là ngày 26 của tháng đang nhập.
- Khoản do chủ xe tự liệt kê: Trạng thái `Đã xác nhận`, Nguồn `Tin nhắn`, Người chi `Tài xế`, Ghi chú `Chủ xe nhập thô <ngày nhập>`.
- Sắp xếp các dòng theo ngày trước khi ghi vào CHI.

## 4. Đặt Trạng thái "Cần kiểm tra" khi

- Ảnh mờ, không đọc được số tiền hoặc ngày.
- Số tiền trong tin nhắn khác số tiền trên ảnh của cùng một khoản.
- Một khoản **không phải dầu** lớn hơn 5.000.000 đ. Dầu thường 5–10 triệu mỗi lần đổ nên không áp ngưỡng này.
- Một lần đổ dầu lớn hơn 12.000.000 đ. **Không** tính mức tiêu hao (lít/km) hay đơn giá dầu, chủ xe không cần.
- Không xác định được loại chi.
- Không có ngày: lấy ngày của tin nhắn hoặc ngày chụp ảnh và ghi chú lại.

Các trường hợp còn lại để Trạng thái `Đã xác nhận`.

## 5. Giá trị mặc định

- Số xe: `75A 40008`
- Người chi: `Tài xế`, trừ khi tin nhắn ghi rõ chủ xe trả hoặc chuyển khoản trực tiếp.
- File chứng từ: đường dẫn tương đối, ví dụ `2026-10/anh/IMG_0012.jpg`.

## 6. Phạm vi sheet CHI – chỉ chi phí của xe

Hiện tại sheet CHI **chỉ ghi các khoản chi trực tiếp cho xe**:

| Loại chi đang dùng (đúng tên trong DANH_MUC) | Ví dụ |
|---|---|
| Dầu | đổ dầu, phiếu dầu, hóa đơn cây xăng |
| Sửa chữa | sửa, hàn, gara, phụ tùng, má phanh, **lốp** (cân lốp, làm lốp, thay/vá/đắp lốp). Chủ xe xác nhận lốp thuộc sửa xe, **không** tách loại Lốp riêng |
| Bảo dưỡng - nhớt | thay nhớt, lọc dầu, lọc gió |
| Rửa xe - vệ sinh | rửa xe |
| Khác | có vẻ là chi cho xe nhưng không rõ loại, đặt `Cần kiểm tra` |

- DANH_MUC còn các loại **chưa dùng**: Phí BOT - cầu đường, CSGT - phạt, Ăn uống - công tác, Bốc xếp. Không xóa các loại này, để dành cho lúc chủ xe muốn bật lên.
- Gặp khoản chi thuộc loại chưa dùng hoặc không liên quan đến xe: **không nhập** vào CHI. Liệt kê các khoản đó trong phần báo lại cho chủ xe để chủ xe quyết định có thêm loại chi mới hay không.
- Khi chủ xe bổ sung loại chi mới, Claude phải: thêm dòng vào bảng trên, thêm vào DANH_MUC trong Excel kèm từ khóa nhận diện, và ghi lại ở mục 7 Nhật ký của README.

## 7. Không nhập vào CHI: chi cố định và tạm ứng

- Chi cố định (lương, trả góp, bảo hiểm, đăng kiểm, phí đường bộ, GPS...) và tạm ứng: **cuối tháng chủ xe gửi một file riêng**. Claude dùng file đó để cập nhật sheet **CHI_CO_DINH** và **TAM_UNG**. Không suy ra các khoản này từ tin nhắn Zalo.
- Nếu tài xế gửi những khoản này trong Zalo thì không nhập vào CHI, chỉ báo lại cho chủ xe biết.
- Lần đầu nhận file này, ghi lại mẫu file (tên file, thư mục lưu, các cột) vào đây để những lần sau làm theo.

## 8. Chốt tháng

### Mẫu file thu của kế toán (từ tháng 9/2026)

- Tên file do kế toán đặt, ví dụ `trạm hương giang tháng 9.2026.xlsx`. **Không đổi tên file**, nhận diện theo cụm "tháng <m>.<yyyy>".
- Sheet đầu tiên tên `CƯỚC 40008`. Hàng 2 là tiêu đề, dữ liệu bắt đầu từ hàng 3. Thứ tự cột giống sheet THU: B Ngày, C Phiếu, D Số xe, E Xuất bãi, F Hàng trừ, G Còn lại (công thức), H Đơn giá, I Thành tiền (công thức), J Ghi chú.
- Cột Phiếu lúc là số, lúc là chữ, nên đổi hết sang số khi nhập.
- Có các dòng STT để trống (chỉ có công thức) và dòng tổng ở cuối, trong đó cột Ngày là công thức liên kết ra file khác. **Chỉ lấy dòng có Ngày là ngày thật và có số Xuất bãi.**
- Đơn giá tháng 9/2026: 73.000 đ/tấn.
- Có thể có nhiều file thu trong một tháng (nhiều trạm/nguồn). Nhập lại một file thì chỉ thay các dòng của **chính file đó** (đối chiếu theo Ngày + Phiếu), không xóa dòng của file khác.

### Thu chủ xe báo bằng lời (không có file)

- Ví dụ: "9 xe không rõ ngày, tổng 22400 kg, nhân giá 13". Hiểu là **tổng** số kg của cả 9 chuyến, đổi ra tấn (22,4), đơn giá 13.000 đ/tấn (tương đương 13 đ/kg).
- Ghi **1 dòng gộp**, không chia ra từng chuyến. Không có ngày thì ghi ngày cuối tháng, để trống Phiếu. Cột K ghi số chuyến, tổng kg, "chủ xe báo <ngày>" và lý do chọn ngày đó.
- Dòng từ file kế toán thì cột K ghi tên trạm, ví dụ `Trạm Hương Giang`.
- Lưu ý: ô "Số phiếu" ở TONG_HOP đếm theo dòng, nên dòng gộp chỉ được tính là 1.

### Các bước

1. Đưa file thu trong `Thu/` vào sheet THU. Nếu chủ xe đã gửi file chi cố định và tạm ứng thì cập nhật CHI_CO_DINH và TAM_UNG (mục 7).
2. Rà soát và báo các điểm bất thường: dầu chiếm bao nhiêu % doanh thu, chi phí/tấn so với tháng trước. Dầu tính theo ngày đổ, **không** chia lại theo tháng dù có đổ cuối tháng dùng sang tháng sau. Nhắc chủ xe nếu tháng đó mới có 1 file thu (có thể còn thu từ nguồn khác).
3. Soạn tin tóm tắt chi trong tháng để chủ xe gửi vào nhóm Zalo cho tài xế xác nhận.
