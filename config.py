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

# ---- Google Sheets (tùy chọn, dùng với cờ --gsheet) ----
# File service account (đã copy từ dự án ADS_facebook)
SERVICE_ACCOUNT_FILE = BASE_DIR / os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json")
# ID Google Sheet đích; để trống thì lần chạy --gsheet đầu tiên sẽ tự tạo sheet mới
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
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

# CÔNG THỨC THƯỞNG NGÀY (bảng "Đề Xuất chi thưởng GR")
# Mỗi dòng: (mốc doanh số ngày, tiền thưởng) - doanh số ĐẠT mốc (>=) nào cao nhất
# thì nhận thưởng mốc đó (vd: đúng 5.000.000 -> thưởng 50.000).

# Ngày thường (Thứ 2 - Thứ 7)
BONUS_TIERS = [
    (25_000_000, 400_000),
    (22_000_000, 300_000),
    (18_000_000, 200_000),
    (14_000_000, 150_000),
    (10_000_000, 100_000),
    (5_000_000, 50_000),
]

# Riêng CHỦ NHẬT thưởng theo mức cao hơn
BONUS_TIERS_SUNDAY = [
    (22_000_000, 350_000),
    (18_000_000, 300_000),
    (14_000_000, 250_000),
    (10_000_000, 200_000),
    (5_000_000, 150_000),
]

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
