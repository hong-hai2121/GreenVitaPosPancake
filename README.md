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

Kết quả:
- In doanh thu ngày, doanh thu theo từng nhân viên kèm bộ phận
- Xuất file Excel `output/doanhthu_<ngày>.xlsx` gồm 2 sheet:
  - **Tổng hợp NV**: mỗi nhân viên một dòng (bộ phận, số đơn, doanh thu) + dòng tổng
  - **Chi tiết đơn**: toàn bộ đơn trong ngày, có cột đánh dấu đơn nào được tính doanh thu

## Cách tính doanh thu

- Tính theo ngày **tạo đơn** (`inserted_at`), múi giờ Việt Nam
- Cộng `total_price` của các đơn, **loại trừ** đơn: Đang hoàn (4), Đã hoàn (5), Đã hủy (6), Đã xóa (7)
- Muốn đổi trạng thái loại trừ: sửa `EXCLUDED_STATUSES` trong `config.py`

## Thông tin kỹ thuật

- API: `https://pos.pages.fm/api/v1` (Pancake POS Open API)
- Xác thực: `api_key` truyền qua query string
- Tài liệu: https://docs.pancake.biz/pos/api/
