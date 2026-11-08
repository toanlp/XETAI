# QUY TẮC NHẬP LIỆU (dành cho Claude)

Tài liệu này mô tả cách Claude đọc ảnh, tin nhắn Zalo và file thu rồi ghi vào `QuanLyXe_75A40008.xlsx`.
Cấu trúc file Excel xem ở README mục 3. **Chủ xe muốn file gọn**: không thêm cột hay sheet mới khi chưa hỏi chủ xe.

## 1. Nguồn dữ liệu

| Nguồn | Vị trí |
|---|---|
| Ảnh hóa đơn, phiếu dầu, biên bản | `Chi/<yyyy-mm>/anh/` |
| Tin nhắn chữ copy từ Zalo | `Chi/<yyyy-mm>/tin_nhan.txt` |
| Danh sách chủ xe gõ thô trong chat với Claude | Claude dán nguyên văn vào `Chi/<yyyy-mm>/tin_nhan.txt` (kèm dòng `# [ngày] Chủ xe nhập thô`) trước khi xử lý, để còn đối chiếu |
| File thu của kế toán | `Thu/` |
| Bảng tổng kết tháng của kế toán khác (ảnh) | Lưu vào `Thu/` tên `doanh thu xe 75A 40008 tháng mm.yyyy (kế toán tổng kết).png`, cách nhập ở mục 9 |

## 2. Các bước "cập nhật chi"

1. **Đọc sheet CHI.** Ảnh nào đã có trong cột Ghi chú (dạng `Ảnh: 2026-10/anh/IMG_0012.jpg`) thì **bỏ qua**.
2. **Đọc từng ảnh mới** và lấy ra ngày, số tiền, nơi phát sinh.
3. **Đọc tin nhắn mới** trong `tin_nhan.txt` (phần nằm sau dòng `=== ĐÃ NHẬP ĐẾN ĐÂY ===`).
4. **Phân loại** theo bảng ở mục 6. Khoản không rõ là gì thì xếp vào `Khác` và đánh dấu Cần kiểm tra. Khoản không liên quan đến xe thì không nhập, chỉ báo lại cho chủ xe.
5. **Chống trùng:** khoản trong tin nhắn và khoản trong ảnh có cùng ngày (lệch tối đa ±1 ngày), cùng loại và cùng số tiền thì chỉ ghi **1 dòng**, kèm `Ảnh: …` ở Ghi chú.
6. **Ghi vào CHI:** thêm dòng mới ngay sau dòng cuối đang có Ngày, sắp theo ngày.
7. **Đánh dấu đã nhập:** chèn lại dòng `=== ĐÃ NHẬP ĐẾN ĐÂY ===` vào cuối file `tin_nhan.txt`.
8. **Báo lại cho chủ xe:** số khoản mới, tổng tiền theo loại, các dòng Cần kiểm tra kèm lý do.

## 3. Cách hiểu số tiền viết tắt

| Tài xế viết | Hiểu là |
|---|---|
| 500k, 500 nghìn, 500 | 500.000 |
| 1tr2, 1.2tr, 1 triệu 2 | 1.200.000 |
| 2tr150 | 2.150.000 |
| 1,5 củ | 1.500.000 |

Số tiền dưới 1.000 mà không có đơn vị thì hiểu là **nghìn đồng**. Nếu không chắc thì đánh dấu Cần kiểm tra.

**Chủ xe nhập thô** theo dạng `<ngày> <nội dung> <số tiền>`, có thể đánh số hoặc không, có thể gõ sai chính tả (ví dụ "suawrxe" là "sửa xe"):
- Chỉ ghi ngày, ví dụ `26`, thì hiểu là ngày 26 của tháng đang nhập.
- Khoản do chủ xe tự liệt kê thì coi là đã xác nhận, để trống Ghi chú.

## 4. Đánh dấu "Cần kiểm tra"

Cách đánh dấu: cột Ghi chú **bắt đầu bằng** `Cần kiểm tra: <lý do>`. Excel sẽ tự tô cam cả dòng. Chủ xe kiểm tra xong thì xóa chữ đó đi.

Đánh dấu khi:
- Ảnh mờ, không đọc được số tiền hoặc ngày.
- Số tiền trong tin nhắn khác số tiền trên ảnh của cùng một khoản.
- Một khoản **không phải dầu** lớn hơn 5.000.000 đ.
- Một lần đổ dầu lớn hơn 12.000.000 đ. **Không** tính mức tiêu hao (lít/km) hay đơn giá dầu, chủ xe không cần.
- Không xác định được loại chi.
- Không có ngày: lấy ngày của tin nhắn hoặc ngày chụp ảnh, ghi rõ ở Ghi chú.

## 5. Ghi chú (cột E của CHI)

- Để trống nếu không có gì đặc biệt. **Không** ghi số xe, người chi, nguồn hay ngày nhập.
- Có ảnh chứng từ thì ghi `Ảnh: <yyyy-mm>/anh/<tên file>`.
- Chủ xe trả trực tiếp (không phải tài xế) thì ghi `Chủ xe trả`.

## 6. Loại chi – chỉ chi phí của xe

Danh sách chọn ở cột B của CHI nằm trong data validation (`"Dầu,Sửa chữa,Khác"`).

| Loại chi | Từ khóa nhận diện |
|---|---|
| Dầu | dầu, DO, đổ dầu, xăng, cây xăng, petrolimex, lít |
| Sửa chữa | sửa, sửa xe, thay, hàn, ốp, gara, phụ tùng, phanh, má phanh, bơm, linh tinh; **lốp**: lốp, vỏ, ruột, cân lốp, làm lốp, vá lốp, đắp lốp; **bảo dưỡng**: nhớt, thay nhớt, bảo dưỡng, lọc dầu, lọc gió; rửa xe |
| Khác | các khoản chi cho xe không phải dầu hay sửa chữa: **đăng kiểm, bảo hiểm, phí đường bộ, GPS, lương...**; ghi rõ ở Nội dung. Khoản không rõ là gì thì đánh dấu Cần kiểm tra |

- Chủ xe đã xác nhận: lốp, nhớt, bảo dưỡng và rửa xe đều tính vào **Sửa chữa**.
- Các khoản như BOT, CSGT, ăn uống, bốc xếp: hiện **không nhập**. Liệt kê trong phần báo lại để chủ xe quyết định.
- **Khi chủ xe thêm loại chi mới:** (1) thêm dòng vào bảng trên; (2) thêm vào danh sách data validation cột B của CHI; (3) thêm 1 dòng loại đó trong khối CHI ở TONG_HOP, ngay trên dòng "Chi cố định", dùng công thức giống dòng Dầu; (4) ghi vào Nhật ký của README.

## 7. Không có chi cố định riêng

- Chủ xe xác nhận (03/10/2026): **không có sheet chi cố định**. Đăng kiểm, bảo hiểm, phí đường bộ... phát sinh tháng nào thì ghi vào **CHI** đúng ngày trả của tháng đó, loại `Khác` (mục 6). Không chia 12 hay phân bổ.
- Cuối tháng chủ xe có thể gửi file các khoản này. Nhập từng khoản thành 1 dòng CHI. Lần đầu nhận file thì ghi lại mẫu file vào đây.
- Tạm ứng không phải chi phí nên không nhập. Nếu chủ xe muốn theo dõi thì hỏi lại trước khi thêm gì vào file.

## 8. Thu và chốt tháng

### Mẫu file thu của kế toán (từ tháng 9/2026)

- Tên file do kế toán đặt, ví dụ `trạm hương giang tháng 9.2026.xlsx`. **Không đổi tên file**, nhận diện theo cụm "tháng <m>.<yyyy>".
- Sheet đầu tiên tên `CƯỚC 40008`. Hàng 2 là tiêu đề, dữ liệu bắt đầu từ hàng 3: B Ngày, C Phiếu, D Số xe, E Xuất bãi, F Hàng trừ, G Còn lại (công thức), H Đơn giá, I Thành tiền (công thức), J Ghi chú.
- Cột Phiếu lúc là số, lúc là chữ, nên đổi hết sang số khi nhập.
- Có các dòng STT để trống (chỉ có công thức) và dòng tổng ở cuối, trong đó cột Ngày là công thức liên kết ra file khác. **Chỉ lấy dòng có Ngày là ngày thật và có số Xuất bãi.**
- Đơn giá tháng 9/2026: 73.000 đ/tấn.

### Ghi vào sheet THU

- Cột: A Ngày, B Phiếu, C Tấn, D Đơn giá, E Thành tiền (công thức, không ghi đè), F Nguồn / Ghi chú.
- **Tấn = Xuất bãi − Hàng trừ.** Nếu có Hàng trừ thì ghi ở cột F, ví dụ `Trạm Hương Giang; trừ 0,5 tấn`.
- Cột F ghi tên trạm, ví dụ `Trạm Hương Giang`.
- Có thể có nhiều file thu trong một tháng (nhiều trạm hoặc nguồn). Nhập lại một file thì chỉ thay các dòng của **chính file đó** (đối chiếu theo Ngày + Phiếu), không xóa dòng của file khác.

### Thu chủ xe báo bằng lời (không có file)

- Ví dụ: "9 xe không rõ ngày, tổng 22400 kg, nhân giá 13". Hiểu là **tổng** số kg của cả 9 chuyến, đổi ra tấn (22,4), đơn giá 13.000 đ/tấn (tương đương 13 đ/kg).
- Ghi **1 dòng gộp**, không chia ra từng chuyến. Không có ngày thì ghi ngày cuối tháng, để trống Phiếu. Cột F ghi số chuyến và tổng kg.
- Lưu ý: "Số chuyến (dòng)" ở TONG_HOP đếm theo dòng, nên dòng gộp chỉ được tính là 1.

### Các bước chốt tháng

1. Đưa file thu trong `Thu/` vào sheet THU. Nếu chủ xe gửi file các khoản đăng kiểm, bảo hiểm... thì nhập vào CHI (mục 7).
2. Rà soát và báo các điểm bất thường: dầu chiếm bao nhiêu % doanh thu, chi phí/tấn so với tháng trước. Dầu tính theo ngày đổ, **không** chia lại theo tháng dù có đổ cuối tháng dùng sang tháng sau. Nhắc chủ xe nếu tháng đó mới có 1 file thu (có thể còn thu từ nguồn khác).
3. Soạn tin tóm tắt chi trong tháng để chủ xe gửi vào nhóm Zalo cho tài xế xác nhận.

## 9. Bảng tổng kết tháng của kế toán khác (từ tháng 9/2026)

Kế toán khác gửi 1 ảnh tổng kết cả tháng. Chủ xe là người **xác nhận lại** bảng này. Bảng là nguồn đầy đủ nhất, nên khi có bảng thì **dữ liệu tháng đó lấy theo bảng** (thay dữ liệu thô đã nhập trước), còn chỗ khác nhau ghi vào `Doi_chieu/yyyy-mm.md`. Mẫu chi tiết và ví dụ: `Doi_chieu/2026-09.md` mục 1.

**Bảng chuyến → THU** (1 chuyến 1 dòng, sắp theo ngày):
- Cột "Khối lượng HH" hoặc "chở ngoài" là **kg** → Tấn = kg / 1000. CƯỚC ghi theo đ/kg (73, 90, 130) → Đơn giá = cước × 1000 đ/tấn.
- Nguồn ghi `NƠI ĐI → NƠI Nhập`; chuyến cột "chở ngoài" thêm tiền tố `Chở ngoài – `. Chuyến "Nam Đông (Huy)" trùng file trạm Hương Giang: giữ Phiếu từ file trạm.
- Kiểm tra: tổng Thành tiền = DOANH THU của bảng.

**Khối tổng kết → CHI:**
- Dầu cty đổ → `Dầu`, mỗi lần đổ 1 dòng (bảng DẦU). Sửa xe (tài xế chi và cty/"o Lí" chi) → `Sửa chữa`, mỗi khoản 1 dòng; khoản cty trả ghi chú `Phát Huy trả`.
- Đi đường (bồi dưỡng, CA, vé, cân), lương tài xế, thuế, CA (dì Lộc), phí qua trạm, lãi → `Khác`, **gộp mỗi khoản 1 dòng ngày cuối tháng**, Nội dung ghi rõ (như tháng 9).
- Kiểm tra: LỢI NHUẬN ở TONG_HOP = "CÒN LẠI" của bảng.
- Quyết toán tài xế (lương, ứng, còn lại) và tạm ứng: không nhập Excel, chỉ ghi vào `Doi_chieu/yyyy-mm.md`.

**Quy tắc rút ra (08/11/2026, chủ xe xác nhận):**
- Cước cố định theo tuyến, ít đổi (trừ khi giá thị trường biến động): Nam Đông → Hào Hưng 73 đ/kg; Nam Đông → Ý Mỹ 90; Đà Nẵng → ... 130. Cước khác với tháng trước thì ghi vào điểm cần kiểm tra.
- Lương khoán: chuyến vào bãi 350.000, chuyến chở ngoài 450.000; lương cứng 3.000.000/tháng.
- Bồi dưỡng theo nơi nhập (cố định theo bãi): Hào Hưng 100.000; Đà Nẵng 200.000; Ý Mỹ chưa rõ (tháng 9 lúc 100k lúc 200k; chủ xe cũng không rõ → Claude tự tìm quy luật qua các tháng sau, chưa tìm ra thì đưa vào danh sách hỏi kế toán).
- Thuế 4,5%: tháng 9 chỉ tính trên thu chuyến cước 73 (80.367.000). Chưa rõ cách tính, chờ tháng 10 để rút ra quy tắc.
- Dầu đổ cuối tháng kế toán có thể tính sang tháng sau (dầu 30/9 9,9tr chuyển sang T10). Khi đối chiếu, khoản dầu thiếu trong tháng thì tìm ở tháng sau trước khi báo sai lệch.
- **Đầu ra khi kiểm tra:** danh sách sai lệch để chủ xe gửi kế toán bên O Lý (nêu dòng, số kế toán ghi, số đúng theo nguồn, nguồn).
