# -*- coding: utf-8 -*-
"""Lưu trữ dữ liệu thô từ Pancake vào api_data/ THEO NGÀY.

- Mỗi ngày 1 file: api_data/donhang_2026-08-27.json (toàn bộ đơn tạo trong ngày đó)
- Danh sách nhân viên: api_data/nhanvien.json (ghi đè mỗi lần chạy)
- Ngày nào được gọi lại API thì file của ngày đó được ghi đè bằng dữ liệu mới
- don_dep(): xóa file của các tháng cũ, chỉ giữ N tháng gần nhất (config.API_DATA_THANG_GIU)
"""
import json
import re
from datetime import date, datetime
from pathlib import Path

import config

API_DATA_DIR = Path(__file__).resolve().parent / "api_data"


def luu_nhan_vien(users: list) -> None:
    API_DATA_DIR.mkdir(exist_ok=True)
    payload = {
        "fetched_at": datetime.now().isoformat(timespec="seconds"),
        "so_nhan_vien": len(users),
        "data": users,
    }
    with open(API_DATA_DIR / "nhanvien.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)


def luu_don_theo_ngay(orders_by_day: dict[date, list]) -> list[Path]:
    """Ghi mỗi ngày 1 file donhang_YYYY-MM-DD.json (ghi đè nếu đã có)."""
    API_DATA_DIR.mkdir(exist_ok=True)
    written = []
    for day, orders in sorted(orders_by_day.items()):
        payload = {
            "ngay": day.isoformat(),
            "fetched_at": datetime.now().isoformat(timespec="seconds"),
            "so_don": len(orders),
            "orders": orders,
        }
        path = API_DATA_DIR / f"donhang_{day.isoformat()}.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=1)
        written.append(path)
    return written


def don_dep(today: date | None = None) -> int:
    """Xóa file donhang_ của tháng cũ, giữ config.API_DATA_THANG_GIU tháng gần nhất."""
    today = today or date.today()
    keep: set[tuple[int, int]] = set()
    y, m = today.year, today.month
    for _ in range(config.API_DATA_THANG_GIU):
        keep.add((y, m))
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    deleted = 0
    if not API_DATA_DIR.exists():
        return 0
    for f in API_DATA_DIR.glob("donhang_*.json"):
        match = re.match(r"donhang_(\d{4})-(\d{2})-\d{2}\.json$", f.name)
        if match and (int(match.group(1)), int(match.group(2))) not in keep:
            f.unlink()
            deleted += 1
    return deleted
