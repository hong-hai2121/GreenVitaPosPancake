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

# Các trạng thái đơn KHÔNG tính vào doanh thu
# 4: Đang hoàn, 5: Đã hoàn, 6: Đã hủy, 7: Đã xóa
EXCLUDED_STATUSES = {4, 5, 6, 7}

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
