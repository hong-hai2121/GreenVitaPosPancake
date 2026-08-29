# Doanh Thu GreenVita - Pancake POS

Dự án lấy doanh thu từ Pancake POS để tính thưởng hàng ngày cho nhân viên.

## Cài đặt

```powershell
cd d:\Python\DoanhThuGreenVitaPosPancake
pip install -r requirements.txt
copy .env.example .env
```

## Lấy API key từ Pancake POS

1. Đăng nhập https://pos.pancake.vn
2. Vào **Cấu hình (Settings)** → **Ứng dụng khác / API**
3. Bấm **Tạo API key** và sao chép (key chỉ hiển thị 1 lần)
4. Dán vào file `.env`:

```
PANCAKE_API_KEY=key_vua_tao
PANCAKE_SHOP_ID=          # để trống, bước sau sẽ biết
```

## Kiểm tra kết nối

```powershell
python test_connection.py
```

Script sẽ liệt kê các shop kèm `shop_id`. Điền `shop_id` của cửa hàng vào `.env` rồi chạy lại lần nữa để xác nhận đọc được đơn hàng.

## Lấy doanh thu

```powershell
python doanh_thu.py                            # hôm nay, tất cả nhân viên
python doanh_thu.py 2026-08-28                 # một ngày cụ thể
python doanh_thu.py 2026-08-28 --bophan sale   # chỉ nhân viên bộ phận có chữ "sale"
python doanh_thu.py 2026-08-01 2026-08-28 --bophan sale  # từng ngày trong khoảng
```

Bộ lọc `--bophan` không phân biệt hoa/thường: `sale` khớp "SALE OCP", "Sale NT", "Sale Nghỉ"...
Gõ từ khóa không khớp bộ phận nào, script sẽ liệt kê toàn bộ bộ phận hiện có để chọn lại.

Mặc định chỉ hiện nhân viên CÓ đơn (giống báo cáo Doanh thu → Nhân viên trên POS).
Thêm `--full` để hiện cả nhân viên 0 đơn của bộ phận (hữu ích khi tính thưởng cho cả đội).

Kết quả:
- In doanh thu ngày, doanh thu theo từng nhân viên kèm bộ phận
- Xuất file Excel `output/doanhthu_<ngày>.xlsx` gồm 2 sheet:
  - **Tổng hợp NV**: mỗi nhân viên một dòng (bộ phận, số đơn, doanh thu) + dòng tổng
  - **Chi tiết đơn**: toàn bộ đơn trong ngày, có cột đánh dấu đơn nào được tính doanh thu

## Cách tính doanh thu (giống báo cáo Doanh thu → Nhân viên trên POS)

- Tính theo ngày **tạo đơn** (`inserted_at`), múi giờ Việt Nam
- **Đơn chốt** = đơn có trạng thái: Đã xác nhận, Đã gửi hàng, Đã nhận, Đang đóng hàng,
  Chờ chuyển hàng, Hoàn 1 phần, Đã thu tiền → sửa trong `CLOSED_STATUSES` (`config.py`)
- **Doanh thu** = tổng tiền các đơn chốt; **Đơn hoàn** = đơn Đang hoàn/Đã hoàn;
  **Doanh số** = Doanh thu + Doanh thu hoàn; **Tỷ lệ hoàn** = Đơn hoàn / (Đơn chốt + Đơn hoàn)
- **Tiền hàng** = Tổng tiền − Phí VC thu của khách (`shipping_fee`) − Phụ thu (`surcharge`)
- Sheet "Tổng hợp NV" có đủ cột: Nhân viên, Bộ phận, Đơn chốt, Doanh số, Doanh thu,
  Đơn hoàn, Tỷ lệ hoàn, Doanh thu hoàn, Tiền hàng, Phí VC thu của khách, Phụ thu

## Bảng "Đề Xuất chi thưởng GR" theo tháng

```powershell
python thuong_thang.py            # cập nhật 5 ngày gần nhất (chạy hàng ngày)
python thuong_thang.py --ngay 10  # đổi cửa sổ cập nhật thành 10 ngày
python thuong_thang.py --ca-thang # tính lại toàn bộ tháng (nghỉ chạy lâu / đối soát)
python thuong_thang.py 2026-07    # tháng cũ: luôn tính lại cả tháng
```

- Ghi vào tab `Thưởng GR T<tháng>.<năm>` trong cùng Google Sheet
- Cấu trúc: STT | Họ và tên | Bộ phận | từng ngày trong tháng | Tổng tháng; dòng cuối
  là Tổng theo ngày (công thức SUM nên sửa tay trên sheet vẫn tự cộng lại)
- Thưởng ngày tính từ **doanh số ngày** của từng nhân viên, theo mốc trong `config.py`:
  - Ngày thường (`BONUS_TIERS`): ≥ 5tr→50k, 10tr→100k, 14tr→150k, 18tr→200k,
    22tr→300k, 25tr→400k
  - **Chủ nhật** (`BONUS_TIERS_SUNDAY`): ≥ 5tr→150k, 10tr→200k, 14tr→250k,
    18tr→300k, 22tr→350k
- Chạy hàng ngày chỉ tính lại N ngày gần nhất (mặc định 5): số thưởng các ngày cũ
  được đọc lại từ chính tab trên sheet và giữ nguyên; đơn cũ đổi trạng thái ngoài
  cửa sổ 5 ngày sẽ không được cập nhật — thỉnh thoảng chạy `--ca-thang` để đối soát
- Google Sheet: xác thực bằng `service_account.json` (sheet phải chia sẻ Editor cho
  `client_email` trong file đó); ID sheet đặt ở `GOOGLE_SHEET_ID` trong `.env`.
  Google không cấp dung lượng Drive cho service account nên khi cần sheet mới, tự tạo
  trong Drive của bạn rồi chia sẻ cho service account và cập nhật `GOOGLE_SHEET_ID`
- Sheet hiện tại: https://docs.google.com/spreadsheets/d/1uli8IN4Ht3O1u5I8y-L-PibLfccx11Guy4ALZVljook

### Chi phí gọi API mỗi lần cập nhật bảng tháng

- Pancake không có API thống kê công khai (trang Thống kê trên web POS dùng API nội bộ
  theo phiên đăng nhập) → phải tự cộng dồn từ danh sách đơn hàng
- Chạy hàng ngày (cửa sổ 5 ngày): 1 call nhân viên + ~2 call đơn hàng (page_size=1000)
  + ~9 call Google Sheets (đọc bảng cũ, ghi, tô màu) ≈ **12 call, ~8 giây**
- Chạy `--ca-thang`: thêm ~8 call đơn hàng (~10 call cho cả tháng ~9-10k đơn)

## Thư mục api_data/ - dữ liệu thô từ Pancake

- Mỗi lần chạy script, toàn bộ phản hồi nhận về từ Pancake API được lưu vào `api_data/`
  (mỗi call 1 file JSON: `001_shops_..._users.json`, `002_..._orders_trang1.json`, ...)
- File chứa: url, tham số gọi (đã ẩn api_key), thời điểm gọi, và dữ liệu trả về nguyên bản
- **Thư mục tự xóa sạch khi bắt đầu lần chạy mới** - chỉ giữ dữ liệu của lần gần nhất

## Thông tin kỹ thuật

- API: `https://pos.pages.fm/api/v1` (Pancake POS Open API)
- Xác thực: `api_key` truyền qua query string
- Tài liệu: https://docs.pancake.biz/pos/api/
