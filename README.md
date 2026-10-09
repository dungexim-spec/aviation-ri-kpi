# Bảng chỉ số TBH hàng không

Web app tĩnh để theo dõi danh mục nhận tái bảo hiểm hàng không theo năm nghiệp vụ (UY).
Tải lên các file theo dõi `UY_x_update_*.xlsx`, app trả về ngay:

- **Tổng quan**: GWP, NWP, net earning, phí đã hưởng (pro-rata theo ngày báo cáo), incurred (đã trả/dự phòng), loss ratio, LR bỏ 2 vụ lớn nhất, chi phí khai thác, capacity, kỳ phí quá hạn, cùng danh sách các điểm cần chú ý được tạo tự động.
- **Phân khúc**: theo broker, kênh, vùng, leader, quốc gia, cedant, loại hợp đồng, so sánh giữa các UY.
- **Bồi thường**: lọc theo UY, trạng thái, cảnh báo ngày tổn thất ngoài kỳ hợp đồng (kèm gợi ý UY đúng).
- **Thu phí**: tuổi nợ kỳ phí, độ phủ lịch thanh toán, danh sách kỳ quá hạn.
- **Cỡ line & tập trung**: line vượt trần, tập trung leader, quốc gia nhạy cảm cấm vận, bind sau ngày hiệu lực, phí dưới mức tối thiểu.
- **Kiểm tra dữ liệu**: danh sách lỗi dữ liệu theo mức Cao / Trung bình / Thấp, kèm hướng xử lý.
- **Xuất báo cáo Excel** gồm tất cả các bảng trên.

## Bảo mật dữ liệu

App **không có máy chủ**. File Excel được đọc và tính toán bằng JavaScript ngay trong trình duyệt của người dùng; không có dữ liệu nào được gửi đi. Đóng tab là dữ liệu mất. Trình duyệt chỉ lưu các ngưỡng cảnh báo ở tab Thiết lập.

**Không commit file dữ liệu thật vào repository.** `.gitignore` đã chặn `*.xlsx`, `*.xls`, `*.csv`.

## Triển khai lên GitHub Pages

1. Đẩy toàn bộ thư mục này lên một repository GitHub (nhánh `main`).
2. Vào **Settings → Pages**.
3. Mục **Source** chọn **Deploy from a branch**, nhánh `main`, thư mục `/ (root)`, bấm **Save**.
4. Sau 1–2 phút, app chạy tại `https://<tài-khoản>.github.io/<tên-repo>/`.

GitHub Pages với repository private cần gói GitHub trả phí (Pro/Team/Enterprise). Với tài khoản miễn phí, repository phải để public; điều này không ảnh hưởng dữ liệu vì repository chỉ chứa mã nguồn.

Chạy trên máy: mở `index.html` bằng trình duyệt là dùng được (cần Internet để tải thư viện SheetJS và font). Hoặc chạy `python3 -m http.server` trong thư mục rồi mở `http://localhost:8000`.

## Định dạng file đầu vào

Mỗi file là một năm nghiệp vụ. App nhận nhãn UY từ tên file (`UY_1` → `UY1`); nhãn có thể sửa trên giao diện. Tên cột không phân biệt hoa thường, thứ tự cột tuỳ ý. Sheet `Policy`, `Risk`, `Premium` phải cùng thứ tự dòng (cùng MRO trên cùng dòng). Bấm **Tải file mẫu** trong app để lấy khung trống.

| Sheet | Cột dùng đến |
|---|---|
| Policy | MRO, Broker, PVI ref number, Original Insured, Cedant Name, Region, Domicile, Risk offer date, Status, Risk bound date, Inception, Expiry, Leader, Risk type, Type, Category, Channel |
| Risk | MRO, PVI share (0,5 = 0,5%), Hull, CSL, A/C, AFV |
| Premium | MRO, Deposit GWP, PVI deposite GWP, PVI deposite NWP, Total deduction, Facility fee, Other Deduction, Fund, PVI actual earning, Premium status |
| Claim | MRO, Claim number, Date claim reported, Occurrence date, Coverage, Type of loss, Gross Incurred, PVI share, PVI Paid Claim, PVI Paid Fee, PVI reserved claim, PVI reserved Fee, Status, Paid Status |
| Installment | MRO, Broker, Original Insured, `Inst n Amount`, `Inst n Due Date`, `Inst n Paid` ("Paid" khi đã thu), `Inst n Actual NWP` (n = 1…12) |

Ngày nhập dạng chữ phải theo `yyyy-mm-dd` hoặc `dd/mm/yyyy`; ô định dạng ngày của Excel được đọc trực tiếp. Ngày không hợp lệ sẽ xuất hiện ở tab Kiểm tra dữ liệu.

## Cách tính

- **Phí đã hưởng** = net earning × số ngày đã qua đến ngày báo cáo / số ngày của hợp đồng (tối đa 100%). Nếu ngày hết hạn không sau ngày hiệu lực, tạm tính 12 tháng và ghi lỗi.
- **Incurred** = đã trả + phí đã trả + dự phòng + phí dự phòng. Bồi thường gắn với hợp đồng qua MRO trong cùng UY; MRO trùng (sửa đổi bổ sung) gắn vào dòng gốc đang Bound.
- **Loss ratio** = incurred / phí đã hưởng. Chưa gồm IBNR, hoa hồng MGA/profit commission và chi phí treaty.
- **Chi phí khai thác** = 1 − net earning / GWP.
- Chuẩn hoá: Region "Asia" → "Asia Pacific RIM"; "LATAM", "South America" → "Americas"; Leader "TMK" → "Tokio Marine Kiln" (sửa trong `DEFAULT_SETTINGS` của `js/engine.js`).

Ngưỡng mặc định (đổi ở tab Thiết lập): trần line 1% chung và 0,5% cho hãng dưới 10 tàu hoặc quốc gia rủi ro cao; một leader tối đa 30% GWP; phí tối thiểu 2.500 USD; cảnh báo kỳ phí quá hạn trên 90 ngày.

## Cấu trúc mã

```
index.html          giao diện
css/app.css         giao diện (sáng/tối)
js/engine.js        đọc file, tính chỉ số, kiểm tra dữ liệu (dùng được trên trình duyệt và Node)
js/app.js           điều khiển giao diện, xuất Excel
tests/engine.test.js kiểm thử engine bằng dữ liệu giả lập
```

Kiểm thử: `npm install && npm test`.

Thư viện: [SheetJS](https://sheetjs.com/) 0.18.5 (tải từ cdnjs).
