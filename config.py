# -*- coding: utf-8 -*-
"""Đọc cấu hình từ file .env."""
import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

PANCAKE_API_KEY = os.getenv("PANCAKE_API_KEY", "").strip()
PANCAKE_SHOP_ID = os.getenv("PANCAKE_SHOP_ID", "").strip()

# Múi giờ tính doanh thu theo ngày
TIMEZONE = "Asia/Ho_Chi_Minh"

# Số đơn tối đa mỗi call API (server Pancake giới hạn cứng <= 1000)
ORDERS_PAGE_SIZE = 1000

# Giữ dữ liệu thô trong api_data/ của bao nhiêu tháng gần nhất (24 = 2 năm)
API_DATA_THANG_GIU = 24

# CHỐT SỔ tháng trước: chạy mặc định (thuong_thang.py / app) CHỈ tự chốt từ CHOT_SO_GIO
# ngày mùng CHOT_SO_NGAY của tháng sau (mặc định 11h00 mùng 2). Lần chạy 9h sáng mùng
# 1, mùng 2 chỉ báo giờ chốt, KHÔNG đụng tháng trước; đúng 11h mùng 2 app / Task
# Scheduler chạy lại để chốt (máy tắt lúc đó thì lần chạy đầu tiên sau mốc chốt bù).
# Đã chốt sổ = tab đóng dấu "ĐÃ CHỐT SỔ", khóa vĩnh viễn, không sửa nữa.
# CHOT_SO_NGAY phải >= 2 để ngày cuối tháng đủ 2 ngày chờ chốt (SETTLE_DELAY_DAYS).
CHOT_SO_NGAY = 2
CHOT_SO_GIO = (11, 0)

# ---- Google Sheets (tùy chọn, dùng với cờ --gsheet) ----
# File service account (đã copy từ dự án ADS_facebook)
SERVICE_ACCOUNT_FILE = BASE_DIR / os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json")
# ID Google Sheet đích; để trống thì lần chạy --gsheet đầu tiên sẽ tự tạo sheet mới
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
# ID trang tính RIÊNG cho bộ phận CSKH (trang tính GOOGLE_SHEET_ID ở trên dành cho Sale)
GOOGLE_SHEET_ID_CSKH = os.getenv("GOOGLE_SHEET_ID_CSKH", "").strip()
# Trang tính "Lịch trực ngày nghỉ - Greenvita": đăng ký làm CHỦ NHẬT
# (Chủ nhật không đăng ký -> thưởng ngày đó = 0)
GOOGLE_SHEET_ID_LICH_TRUC = os.getenv(
    "GOOGLE_SHEET_ID_LICH_TRUC", "1QuUke2Lqms0pfYSRz8FdEExSXaSxAFyvUvaq0yuma8c").strip()
# Bảng chấm công 1: "CHẤM CÔNG NT/TK_2026_GreenVita" (file Excel .xlsx trên Drive, tab
# "BCC.<n>"; nút "Mở chấm công 1" trên GUI). Là file Office nên KHÔNG đọc được bằng
# Sheets API/gspread; cham_cong.py tải qua Drive API rồi mở bằng openpyxl.
GOOGLE_SHEET_ID_CHAM_CONG = os.getenv(
    "GOOGLE_SHEET_ID_CHAM_CONG", "12Ht00Fy5dvV1QMtmSjPx0zuUZRR6KIhM").strip()
# Bảng chấm công 2: "Chấm công OCP 2026" (Google Sheet gốc, tab "Tháng <n>" / "BCC T<n>";
# nút "Mở chấm công 2" trên GUI). Nhân viên KHÔNG có tên ở bảng 1 thì tìm sang bảng 2,
# cùng 1 cách tính % Thưởng (xem cham_cong.py).
GOOGLE_SHEET_ID_CHAM_CONG_2 = os.getenv(
    "GOOGLE_SHEET_ID_CHAM_CONG_2", "1oU7jgxYnOpIaaFGajSsuCy06JCVFs0ToVuHxIvzpRuw").strip()
# Tên gọi 2 bảng chấm công - ghi ở cột "Bảng chấm công" của khối CHẤM CÔNG dưới bảng BC02
# để biết % Thưởng của nhân viên lấy từ bảng nào (1 = file NT/TK, 2 = Chấm công OCP)
CHAM_CONG_TEN_BANG = {1: "CHẤM CÔNG NT/TK", 2: "Chấm công OCP"}
# % THƯỞNG trên BC02 tính từ bảng chấm công (tab của tháng, xem cham_cong.py):
#   quy đổi ngày nghỉ = ngày nghỉ KHÔNG LƯƠNG + (giờ làm thiếu so với 8h) / 8
#   cứ đủ CHAM_CONG_NGAY_NGHI_MOI_BAC ngày -> trừ CHAM_CONG_TRU_MOI_BAC % (không âm)
CHAM_CONG_GIO_CHUAN = 8               # 1 ngày công = 8 giờ; ô ghi < 8 là làm thiếu giờ
CHAM_CONG_NGAY_NGHI_MOI_BAC = 2       # nghỉ đủ 2 ngày ...
CHAM_CONG_TRU_MOI_BAC = 10            # ... trừ 10% thưởng (4 ngày -> 20%, ...)
# Mã ô -> số ngày nghỉ KHÔNG LƯƠNG (so sau khi bỏ dấu, viết HOA, bỏ khoảng trắng)
CHAM_CONG_MA_KHONG_LUONG = {"KL": 1.0, "KL/2": 0.5, "X/2": 0.5, "M/2": 0.5}
# Mã CÓ LƯƠNG / đủ công -> không trừ (P phép, NL nghỉ lễ, CĐ nghỉ chế độ, X/M đủ công,
# HV học việc, CN / CN/2 đi làm Chủ nhật cả / nửa ngày - bảng OCP hay ghi)
CHAM_CONG_MA_CO_LUONG = {"P", "P/2", "NL", "CD", "X", "M", "HV", "CN", "CN/2"}
# Bộ phận áp dụng (tab BC02 của nhóm nào lấy % từ chấm công): cả Sale lẫn CSKH
CHAM_CONG_AP_DUNG_NHOM = {"sale", "cskh"}
# Email được chia sẻ quyền chỉnh sửa khi tự tạo sheet mới
GOOGLE_SHARE_EMAIL = os.getenv("GOOGLE_SHARE_EMAIL", "").strip()
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ĐƠN CHỐT - các trạng thái được tính vào doanh thu (giống báo cáo POS):
# 1: Đã xác nhận, 2: Đã gửi hàng, 3: Đã nhận, 8: Đang đóng hàng,
# 9: Chờ chuyển hàng, 15: Hoàn 1 phần, 16: Đã thu tiền
CLOSED_STATUSES = {1, 2, 3, 8, 9, 15, 16}

# ĐƠN HOÀN - 4: Đang hoàn, 5: Đã hoàn
RETURN_STATUSES = {4, 5}

# CÔNG THỨC THƯỞNG NGÀY (bảng "Đề Xuất chi thưởng GR") - THEO TỪNG NHÓM BỘ PHẬN
# Mỗi dòng: (mốc doanh số ngày, tiền thưởng) - doanh số ĐẠT mốc (>=) nào cao nhất
# thì nhận thưởng mốc đó (vd: đúng 5.000.000 -> thưởng theo mốc 5tr).
# "weekday" = Thứ 2 - Thứ 7; "sunday" = riêng Chủ nhật (mức cao hơn).
BONUS_TIERS_BY_GROUP = {
    "sale": {
        "weekday": [
            (25_000_000, 400_000),
            (22_000_000, 300_000),
            (18_000_000, 200_000),
            (14_000_000, 150_000),
            (10_000_000, 100_000),
            (5_000_000, 50_000),
        ],
        # Theo Quyết định của công ty: Sale Chủ nhật có mốc 22tr -> 400k và 25tr -> 500k
        "sunday": [
            (25_000_000, 500_000),
            (22_000_000, 400_000),
            (18_000_000, 300_000),
            (14_000_000, 250_000),
            (10_000_000, 200_000),
            (5_000_000, 150_000),
        ],
    },
    "cskh": {
        "weekday": [
            (22_000_000, 250_000),
            (18_000_000, 200_000),
            (14_000_000, 150_000),
            (10_000_000, 100_000),
            (5_000_000, 50_000),
        ],
        "sunday": [
            (22_000_000, 350_000),
            (18_000_000, 300_000),
            (14_000_000, 250_000),
            (10_000_000, 200_000),
            (5_000_000, 150_000),
        ],
    },
}

# PHỤ CẤP đi làm CHỦ NHẬT (đồng / 1 Chủ nhật có "Đăng kí làm" trên Lịch trực).
# CỘNG THẲNG vào ô Chủ nhật trên bảng thưởng GR (= thưởng theo mốc + phụ cấp, kể cả
# khi không đạt mốc nào; 1 người nhận 1 lần dù có nhiều tài khoản Pancake) -> Tổng
# tháng / BC02 đã gồm phụ cấp. Khối "Đăng kí làm Chủ nhật" phía dưới tách riêng
# "<phụ cấp> lương + <thưởng> thưởng" để soát. Đổi số ở đây xong chạy:
#     python thuong_thang.py --tinh-lai      (áp lại cho các ngày đã lên bảng)
PHU_CAP_CHU_NHAT = {"sale": 100_000, "cskh": 200_000}

# Ý nghĩa mã trạng thái đơn hàng của Pancake POS
ORDER_STATUS = {
    0: "Mới",
    1: "Đã xác nhận",
    2: "Đã gửi hàng",
    3: "Đã nhận",
    4: "Đang hoàn",
    5: "Đã hoàn",
    6: "Đã hủy",
    7: "Đã xóa",
    8: "Đang đóng hàng",
    9: "Chờ chuyển hàng",
    11: "Chờ hàng",
    12: "Chờ in",
    13: "Đã in",
    15: "Hoàn một phần",
    16: "Đã thu tiền",
    20: "Đã đặt hàng",
}


def require_api_key() -> str:
    if not PANCAKE_API_KEY or PANCAKE_API_KEY.startswith("dien_api_key"):
        raise SystemExit(
            "Chưa cấu hình PANCAKE_API_KEY.\n"
            "1. Sao chép .env.example thành .env\n"
            "2. Đăng nhập https://pos.pancake.vn -> Cấu hình -> Ứng dụng khác / API -> Tạo API key\n"
            "3. Dán API key vào file .env"
        )
    return PANCAKE_API_KEY
