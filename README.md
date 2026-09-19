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
python thuong_thang.py            # chạy hàng ngày: cập nhật cả 5 tab + tự chốt sổ tháng trước
                                  # (chỉ từ 11h00 mùng 2 tháng sau, sớm hơn thì chưa đụng)
python thuong_thang.py 2026-08    # tháng đã qua: CHỐT SỔ ngay (tính lại cả tháng rồi khóa)
python thuong_thang.py --tinh-lai # tính lại MỌI ngày của tháng hiện tại theo quy tắc hiện
                                  # hành (không khóa) - dùng sau khi đổi mốc thưởng/phụ cấp
```

MỘT lệnh lấy dữ liệu MỘT lần rồi cập nhật cả 5 tab trên HAI TRANG TÍNH riêng
(`.env`: `GOOGLE_SHEET_ID` = trang tính Sale, `GOOGLE_SHEET_ID_CSKH` = trang tính CSKH):

Trang tính **Sale** (không có thông tin CSKH):
1. **Thưởng Sale GR T09.2026** - thưởng ngày bộ phận Sale; NGAY DƯỚI bảng có khối
   **"ĐĂNG KÍ LÀM CHỦ NHẬT"** CHÉP từ trang tính *Lịch trực ngày nghỉ - Greenvita*:
   tab nào có chữ **sale** trong tên -> trang tính Sale, có chữ **cskh** -> trang
   tính CSKH (nhiều tab cùng chữ thì gộp; tab không có cột "Họ Tên" bị bỏ qua).
   Người = đúng danh sách trên lịch, cột = mọi Chủ nhật CỦA THÁNG (kể cả ngày
   chưa lên bảng thưởng). Ô Chủ nhật CHỈ ghi ngày có đăng kí làm, TÁCH RÕ dạng
   gọn `100k + 250k` = phụ cấp + thưởng (phụ cấp `config.PHU_CAP_CHU_NHAT`: Sale
   100.000, CSKH 200.000); ngày chưa lên bảng thưởng -> `Đăng kí`; không đăng kí
   -> để trống (không ghi "Nghỉ"); `-` = lịch chưa có cột ngày đó. Còn Ô CHỦ
   NHẬT TRÊN BẢNG THƯỞNG là số ĐÃ CỘNG (thưởng theo mốc + phụ cấp, kể cả khi
   không đạt mốc nào; 1 người nhận phụ cấp 1 lần dù có nhiều tài khoản Pancake)
   -> Tổng tháng và cột Thưởng của BC02 đã gồm phụ cấp. Hai cột tên CẠNH NHAU
   để soát: **Tên trên lịch trực** và **Tên thực tế trên bảng thưởng** (= tên lịch
   map sang tên nhân viên Pancake; 1 tên lịch ứng với nhiều tài khoản thì CHỈ hiện
   tài khoản có số ở ô Chủ nhật trên bảng thưởng, mỗi tài khoản 1 dòng trong ô -
   chưa tài khoản nào có số thì hiện cả; không khớp ai thì cảnh báo). Dòng cuối liệt kê nhân viên
   có trên bảng thưởng nhưng chưa có trong lịch, kèm link lịch trực
2. **Doanh số Sale T09.2026** - ma trận DOANH SỐ ngày nhân viên Sale (cùng cấu trúc
   bảng thưởng, để đối chiếu: doanh số ô nào -> thưởng ô đó theo mốc)
3. **Doanh số Sale Page T09.2026** - ma trận DOANH SỐ ngày theo PAGE NGUỒN:
   gộp theo nguồn đơn hàng (page quảng cáo dẫn đơn về, trường `account`),
   doanh số = đơn chốt + đơn hoàn, tính theo NGÀY TẠO đơn (cùng mốc với ma trận
   nhân viên); chỉ tính đơn do nhân viên bộ phận Sale phụ trách (dựng từ đơn thô
   `api_data/`, xếp theo tổng tháng giảm dần; trang tính CSKH có tab
   "Doanh số CSKH Page ..." lọc theo CSKH)
4. **BC02 Thưởng DS Sale T09.2026** - thưởng doanh số tháng bộ phận Sale
   (cột % Thưởng tự tính từ bảng CHẤM CÔNG tháng, kèm khối "CHẤM CÔNG" để soát)

Trang tính **CSKH** (không có thông tin Sale):
4. **Thưởng CSKH GR T09.2026** - thưởng ngày bộ phận CSKH (cũng kèm khối
   "ĐĂNG KÍ LÀM CHỦ NHẬT" của bộ phận CSKH)
5. **Doanh số CSKH T09.2026** - ma trận DOANH SỐ ngày nhân viên CSKH
6. **BC02 Thưởng DS CSKH T09.2026** - thưởng doanh số tháng bộ phận CSKH
   (cột % Thưởng tự tính từ bảng CHẤM CÔNG tháng, kèm khối "CHẤM CÔNG" như Sale)

Lần chạy đầu sau khi tách: dữ liệu CSKH của tháng chưa khóa được tự CHUYỂN từ trang
tính cũ sang trang tính CSKH (kế thừa số đã chốt); các tab gộp cũ ("Doanh số NV",
"BC02 Thưởng DS Sale- CSKH") được tách thành tab riêng từng bộ phận rồi xóa.
Trang tính CSKH phải chia sẻ quyền Editor cho email service account
(`client_email` trong `service_account.json`).

Logic "chốt ngày, chốt sổ":
- **Trễ 2 ngày**: hôm nay 29 thì bảng chỉ hiển thị đến 27 (2 ngày cuối trạng thái đơn
  còn thay đổi nên chưa đưa vào)
- **Mỗi lần chạy gọi API 7 NGÀY GẦN NHẤT** (đến hôm nay − 2, giới hạn trong tháng)
  và GHI ĐÈ kho `api_data/` - kho luôn tươi trong cửa sổ 7 ngày
- **Ngày đã lên bảng = đã chốt** (tab Thưởng GR / Doanh số NV / BC02): chỉ NỐI CỘT
  của ngày mới; số các ngày cũ giữ nguyên trên sheet, kể cả khi kho api_data
  quá khứ đã được ghi đè mới hơn
- **Tab Doanh số Page luôn dựng lại từ kho** - 7 ngày gần nhất của nó phản ánh
  trạng thái đơn mới nhất (có thể lệch nhẹ với ma trận NV ở 7 ngày cuối, vì ma
  trận NV đã chốt còn Page thì cập nhật tiếp)
- **Chốt sổ cuối tháng**: từ **11h00 mùng 2 tháng sau** (`CHOT_SO_NGAY` / `CHOT_SO_GIO`
  trong `config.py`), chạy mặc định sẽ tự gọi lại API một lần trọn tháng trước để sửa
  thưởng lần cuối (bắt đơn hoàn/hủy muộn), đóng dấu **"ĐÃ CHỐT SỔ"** lên tiêu đề các
  tab - từ đó tab bị khóa, mọi lần chạy sau bỏ qua, kho api_data của tháng đó cũng
  không bị ghi đè nữa. Lần chạy TRƯỚC mốc đó (9h sáng mùng 1, mùng 2) chỉ in dòng
  "sẽ tự CHỐT SỔ lúc 11:00 ngày 02/..." và không đụng tháng trước; máy tắt đúng giờ
  thì lần chạy đầu tiên sau mốc sẽ chốt bù. Muốn chốt ngay không chờ:
  `python thuong_thang.py 2026-09`

- Ghi vào tab `Thưởng GR T<tháng>.<năm>` trong cùng Google Sheet
- Cấu trúc: STT | Họ và tên | Bộ phận | từng ngày trong tháng | Tổng tháng; dòng cuối
  là Tổng theo ngày (công thức SUM nên sửa tay trên sheet vẫn tự cộng lại)
- Thưởng ngày tính từ **doanh số ngày** của từng nhân viên, theo mốc riêng từng nhóm
  trong `BONUS_TIERS_BY_GROUP` (`config.py`):
  - **Sale** ngày thường: ≥ 5tr→50k, 10tr→100k, 14tr→150k, 18tr→200k, 22tr→300k, 25tr→400k
  - **Sale** Chủ nhật: ≥ 5tr→150k, 10tr→200k, 14tr→250k, 18tr→300k, 22tr→400k, 25tr→500k
  - **CSKH** ngày thường: ≥ 5tr→50k, 10tr→100k, 14tr→150k, 18tr→200k, 22tr→250k
  - **CSKH** Chủ nhật: ≥ 5tr→150k, 10tr→200k, 14tr→250k, 18tr→300k, 22tr→350k
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
- Cột **Trừ tiền đơn hoàn** (`config.TRU_DON_HOAN_*`, CHỈ tab CSKH - tab Sale không có
  cột này): đơn hoàn
  vượt quá **3%** số đơn chốt thì mỗi đơn vượt trừ **50.000đ** - số đơn hoàn tối đa =
  INT(đơn chốt x 3%), vd 100 đơn chốt -> tối đa 3 đơn hoàn, hoàn 5 đơn -> vượt 2 -> trừ
  100.000đ (công thức trên sheet, sửa tay số đơn thì tự tính lại)
- Cột **Thưởng** = cột "Tổng tháng" trên 2 tab Thưởng Sale/CSKH GR, tự mang qua
  khớp theo tên nhân viên
- Cột công thức trên sheet: Tỷ lệ hoàn, Trừ tiền đơn hoàn (CSKH), Thực nhận = Thưởng x
  % Thưởng - Trừ tiền đơn hoàn (% trống hiểu là 100%). Công thức dùng dấu `;` vì sheet
  locale Việt Nam
- Cột **% Thưởng** (bộ phận trong `config.CHAM_CONG_AP_DUNG_NHOM`, mặc định cả Sale và CSKH):
  TỰ TÍNH từ **2 bảng CHẤM CÔNG** - xem `cham_cong.py`:
  - **Bảng 1 "CHẤM CÔNG NT/TK"** (file Excel trên Drive, `GOOGLE_SHEET_ID_CHAM_CONG`,
    tab "BCC.<tháng>") tìm trước; nhân viên **không có tên** ở bảng 1 thì tìm sang
    **bảng 2 "Chấm công OCP"** (Google Sheet, `GOOGLE_SHEET_ID_CHAM_CONG_2`, tab
    "BCC T<tháng>" / "Tháng <tháng>"), cùng 1 cách tính. Tên gọi 2 bảng đặt ở
    `config.CHAM_CONG_TEN_BANG`. Bảng nào không đọc được chỉ cảnh báo, vẫn dùng bảng kia.
  - Tab của tháng chọn theo tên tab + ô THÁNG/NĂM trên header; lệch nhau (tab "Tháng 9"
    của OCP quên sửa ô THÁNG nên dòng ngày vẫn là tháng 8) thì tin theo tên tab, các cột
    hiểu là ngày 1..31 của tháng đó và ghi dòng "Lưu ý" dưới khối chấm công.
  - **Số ngày nghỉ trong tháng** = nghỉ phép `P` + nghỉ không lương `KL` + nửa công
    (`P/2`, `KL/2`, `X/2`, `M/2` hoặc ô ghi số giờ từ 4 trở xuống = 0,5 ngày) - giống
    cột "Tổng số ngày nghỉ" của HR. KHÔNG tính nghỉ chế độ theo quy định `CĐ`, nghỉ
    lễ/Tết `NL`, đủ công `X`/`M`/ô trên 4h, học việc `HV`, đi làm Chủ nhật `CN`/`CN/2`;
    trống = không phải ngày làm.
  - **% Thưởng**: từ 0-2 ngày nghỉ -> **100%**; trên 2 ngày -> **85%**; trên 3 ngày
    HOẶC có ngày nghỉ **không giấy phép** (mã `KP`/`NKP`/`KGP`/`VKP`) -> **60%**.
    Mốc và mã ô đặt trong `config.CHAM_CONG_*`. Tên khớp theo tiền tố bỏ dấu (như Lịch trực).
  - Người KHÔNG có trên cả 2 bảng (đã nghỉ, tên viết khác ...) hoặc không đọc được
    bảng nào: giữ % đang có trên sheet (nhập tay được), mặc định 100%. Log ghi
    `[CHAM CONG][OK]` / `[CHAM CONG][LOI]` cho TỪNG bảng như lịch trực.
  - Người CÓ đơn / doanh số mà không có tên trên bảng chấm công nào thì **cả dòng tô
    VÀNG** ở bảng BC02 lẫn trong khối chấm công bên dưới - báo hiệu cần sửa tên trên
    Pancake hoặc trên chấm công cho khớp (tài khoản không có đơn thì không tô).
  - Dưới bảng có khối **CHẤM CÔNG THÁNG ...**: mỗi tài khoản Pancake 1 dòng - tên
    trên chấm công, cột **Bảng chấm công** ("CHẤM CÔNG NT/TK" hoặc "Chấm công OCP" =
    bảng đã lấy số liệu), ngày phép, KL, nửa công, không giấy phép, tổng ngày nghỉ, %
    và chi tiết ngày để soát. Khối
    CHỈ liệt kê tài khoản có Đơn chốt > 0 hoặc có doanh số trong tháng; ai cả hai
    = 0 (đã nghỉ, chưa có đơn) thì bỏ ra ngoài khối (tiêu đề khối ghi số bị bỏ qua).
- **Thực nhận = Thưởng x % Thưởng - Trừ tiền đơn hoàn** (khoản trừ chỉ có ở CSKH; Thưởng
  trống mà vẫn bị trừ thì ra số âm = khoản phải trừ tiếp vào lương). Sale: Thực nhận =
  Thưởng x % Thưởng

## Ứng dụng desktop + lịch chạy tự động (9h sáng hằng ngày, 11h mùng 2 chốt sổ)

**Cách 1 - Ứng dụng GUI (nháy đúp `app_cap_nhat.pyw`):** mở cửa sổ desktop hiện
tiến trình trong khung log và **đồng hồ đếm ngược tới lần chạy gần nhất** rồi tự chạy.
Có 2 lịch: **9h sáng hằng ngày** (chỉnh được trên giao diện) cập nhật tháng hiện tại,
và **11h00 mùng 2 hằng tháng** chạy **chốt sổ tháng trước** (dòng dưới đồng hồ ghi rõ
cả 2 mốc). Có nút **Cập nhật ngay**, **Sheet Sale**, **Sheet CSKH**, **Mở file log**;
ba ô *Link lịch trực CN*, *Link chấm công 1* (file Excel "CHẤM CÔNG NT/TK") và
*Link chấm công 2* (Google Sheet "Chấm công OCP 2026") - cả hai dùng tính % Thưởng,
bảng 1 tìm trước - đều có nút **Lưu link** (ghi vào `.env`, nhận cả link đầy đủ lẫn ID)
và nút **Mở ...** (mở trang tính đang dùng trên trình duyệt).
- **Chạy nền giống Unikey:** app có icon **GV** ở khay hệ thống (góc phải taskbar,
  cần `pystray` + `Pillow` trong `requirements.txt`). Bấm **X** chỉ **ẩn cửa sổ xuống
  khay** - app vẫn chạy, đồng hồ vẫn đếm và vẫn tự cập nhật đúng giờ (lần đầu ẩn có
  thông báo nhắc). Nháy trái icon khay (hoặc nháy đúp lối tắt trên Desktop lần nữa)
  thì cửa sổ hiện lại; chuột phải icon có menu **Mở cửa sổ / Cập nhật ngay / Mở file
  log / Thoát**. Chỉ **Thoát** ở menu này mới tắt hẳn (đang cập nhật dở thì hỏi lại).
  App chỉ chạy **1 phiên bản**: mở lần 2 chỉ gọi cửa sổ đang chạy lên, không mở trùng
  (nhận diện qua cổng localhost 47321).
- **Khởi động cùng Windows:** tick ô **Khởi động cùng Windows (chạy ẩn ở khay)** trên
  giao diện -> tạo lối tắt **GreenVita - Cập nhật thưởng (khay)** trong thư mục Startup
  (`shell:startup`) chạy `pythonw app_cap_nhat.pyw --khay`: đăng nhập máy là app tự chạy
  ẩn ở khay, không hiện cửa sổ. Bỏ tick để tắt (hoặc `python tao_loi_tat.py --bo-khoi-dong`).
- Tạo icon trên Desktop (chạy 1 lần, hoặc chạy lại khi đổi máy/đổi thư mục):
  `python tao_loi_tat.py` -> vẽ `icon_app.ico`, tạo lối tắt **GreenVita - Cập nhật thưởng**
  trên Desktop (chạy bằng `pythonw`, không hiện cửa sổ đen) VÀ lối tắt khởi động cùng
  Windows nói trên.

**Cách 2 - Chạy ngầm dự phòng:** Task Scheduler của Windows có task
**GreenVita_CapNhatThuong** chạy `pythonw app_cap_nhat.pyw --ngam` với 2 lịch kích hoạt:
**9:00 sáng mỗi ngày** và **11:00 ngày 2 hằng tháng** (chốt sổ tháng trước); máy bật
muộn hơn thì tự chạy bù. Hai cách chạy trùng nhau không sao - script tự phát hiện
"không có ngày mới" / "ĐÃ CHỐT SỔ" và bỏ qua.
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
