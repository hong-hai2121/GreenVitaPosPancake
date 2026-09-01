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
python thuong_thang.py            # chạy hàng ngày: cập nhật cả 3 tab + tự chốt sổ tháng trước
python thuong_thang.py 2026-08    # tháng đã qua: CHỐT SỔ ngay (tính lại cả tháng rồi khóa)
```

MỘT lệnh lấy dữ liệu MỘT lần rồi cập nhật cả 3 tab:
1. **Thưởng Sale GR T08.2026** - thưởng ngày bộ phận Sale
2. **Thưởng CSKH GR T08.2026** - thưởng ngày bộ phận CSKH
3. **BC02 Thưởng DS Sale- CSKH T08.2026** - thưởng doanh số tháng (giữ cột nhập tay)

Logic "chốt ngày, chốt sổ":
- **Trễ 2 ngày**: hôm nay 29 thì bảng chỉ hiển thị đến 27 (2 ngày cuối trạng thái đơn
  còn thay đổi nên chưa đưa vào)
- **Ngày đã lên bảng = đã chốt**: chạy hàng ngày chỉ gọi API lấy đơn của NGÀY MỚI
  (thường 1 call) rồi nối cột; số các ngày cũ giữ nguyên. Không có ngày mới thì
  không gọi API đơn hàng (~5 giây)
- **Chốt sổ cuối tháng**: từ mùng 2 tháng sau, chạy mặc định sẽ tự gọi lại API một lần
  trọn tháng trước để sửa thưởng lần cuối (bắt đơn hoàn/hủy muộn), đóng dấu
  **"ĐÃ CHỐT SỔ"** lên tiêu đề 3 tab - từ đó tab bị khóa, mọi lần chạy sau bỏ qua

- Ghi vào tab `Thưởng GR T<tháng>.<năm>` trong cùng Google Sheet
- Cấu trúc: STT | Họ và tên | Bộ phận | từng ngày trong tháng | Tổng tháng; dòng cuối
  là Tổng theo ngày (công thức SUM nên sửa tay trên sheet vẫn tự cộng lại)
- Thưởng ngày tính từ **doanh số ngày** của từng nhân viên, theo mốc riêng từng nhóm
  trong `BONUS_TIERS_BY_GROUP` (`config.py`):
  - **Sale** ngày thường: ≥ 5tr→50k, 10tr→100k, 14tr→150k, 18tr→200k, 22tr→300k, 25tr→400k
  - **CSKH** ngày thường: ≥ 5tr→50k, 10tr→100k, 14tr→150k, 18tr→200k, 22tr→250k
  - **Chủ nhật** (cả 2 nhóm): ≥ 5tr→150k, 10tr→200k, 14tr→250k, 18tr→300k, 22tr→350k
- Google Sheet: xác thực bằng `service_account.json` (sheet phải chia sẻ Editor cho
  `client_email` trong file đó); ID sheet đặt ở `GOOGLE_SHEET_ID` trong `.env`.
  Google không cấp dung lượng Drive cho service account nên khi cần sheet mới, tự tạo
  trong Drive của bạn rồi chia sẻ cho service account và cập nhật `GOOGLE_SHEET_ID`
- Sheet hiện tại: https://docs.google.com/spreadsheets/d/1uli8IN4Ht3O1u5I8y-L-PibLfccx11Guy4ALZVljook

### Chi phí gọi API mỗi lần chạy (cả 3 bảng)

- Pancake không có API thống kê công khai (trang Thống kê trên web POS dùng API nội bộ
  theo phiên đăng nhập) → phải tự cộng dồn từ danh sách đơn hàng
- Chạy hàng ngày: 1 call nhân viên + ~1 call đơn của ngày mới + ~15 call Google Sheets
- Chốt sổ cuối tháng (1 lần/tháng): thêm ~10 call đơn trọn tháng (page_size=1000)

## Bảng BC02 "Thưởng DS Sale- CSKH" theo tháng

```powershell
python bc02_thuong_ds.py            # tháng hiện tại -> tab "BC02 Thưởng DS Sale- CSKH T08.2026"
python bc02_thuong_ds.py 2026-07    # tháng cụ thể
```

- Gồm TẤT CẢ nhân viên các bộ phận có chữ sale hoặc cskh, dòng Tổng nằm ngay dưới header
- Cột từ API: Đơn chốt, Đơn hoàn tháng này, DS bán hàng (= tổng tiền đơn chốt)
- Cột **Đơn hoàn tháng trước**: TẠM THỜI để 0 theo yêu cầu (code tính thật đã có sẵn -
  hàm `dem_hoan_thang_truoc` đọc `status_history`, khi cần bật lại chỉ 1 dòng)
- **Tỷ lệ hoàn** = (hoàn tháng này + hoàn tháng trước) / đơn chốt
- Cột **Thưởng** = cột "Tổng tháng" trên 2 tab Thưởng Sale/CSKH GR, tự mang qua
  khớp theo tên nhân viên
- Cột công thức trên sheet: Tỷ lệ hoàn = hoàn/(chốt+hoàn); Thực nhận = Thưởng x % Thưởng
  (% trống hiểu là 100%). Công thức dùng dấu `;` vì sheet locale Việt Nam
- Cột **% Thưởng**: mặc định điền sẵn **100%**, chỉnh tay trên sheet thì
  chạy lại script vẫn giữ nguyên giá trị đã chỉnh; **Thực nhận = Thưởng x % Thưởng**

## Ứng dụng desktop + lịch chạy tự động 9h sáng

**Cách 1 - Ứng dụng GUI (nháy đúp `app_cap_nhat.pyw`):** mở cửa sổ desktop:
tự chạy cập nhật ngay, hiển thị tiến trình trong khung log, rồi hiện **đồng hồ
đếm ngược tới 9h sáng hôm sau** và tự chạy tiếp. Có nút **Cập nhật ngay**,
**Mở Google Sheet**, **Mở file log**.

**Cách 2 - Chạy ngầm dự phòng:** Task Scheduler của Windows có task
**GreenVita_CapNhatThuong** chạy `pythonw app_cap_nhat.pyw --ngam` lúc **9:00 sáng**
mỗi ngày (máy bật muộn hơn thì tự chạy bù). Hai cách chạy trùng nhau không sao -
script tự phát hiện "không có ngày mới" và bỏ qua.
- Kết quả mỗi lần chạy ghi vào `logs/cap_nhat.log` (mở bằng VS Code để xem)
- Quản lý: mở **Task Scheduler** (gõ vào Start menu) → tìm task theo tên; hoặc lệnh:
  - Chạy ngay: `Start-ScheduledTask -TaskName "GreenVita_CapNhatThuong"`
  - Tạm dừng: `Disable-ScheduledTask -TaskName "GreenVita_CapNhatThuong"`
  - Xóa lịch: `Unregister-ScheduledTask -TaskName "GreenVita_CapNhatThuong"`

## Thư mục api_data/ - kho dữ liệu thô theo NGÀY

- **Mỗi ngày 1 file** đặt tên theo ngày dữ liệu: `donhang_2026-08-27.json`
  (toàn bộ đơn tạo trong ngày đó, kèm thời điểm tải); nhân viên: `nhanvien.json`
- Ngày nào được gọi lại API (chốt sổ, tải lại) thì file ngày đó được **ghi đè** bằng bản mới
- **Giữ 2 tháng gần nhất** (`API_DATA_THANG_GIU` trong config.py) - file tháng cũ hơn
  tự xóa sau mỗi lần chạy
- Tải lại dữ liệu một khoảng ngày (không đụng Google Sheet):
  `python tai_lai_du_lieu.py` (đầu tháng đến nay) hoặc
  `python tai_lai_du_lieu.py 2026-08-01 2026-08-29`
- Giới hạn server Pancake: tối đa **1.000 đơn/call** (đã dùng mức trần,
  `ORDERS_PAGE_SIZE` trong config.py)

## Thông tin kỹ thuật

- API: `https://pos.pages.fm/api/v1` (Pancake POS Open API)
- Xác thực: `api_key` truyền qua query string
- Tài liệu: https://docs.pancake.biz/pos/api/
