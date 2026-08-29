# -*- coding: utf-8 -*-
"""Tải lại dữ liệu đơn hàng thô vào api_data/ theo từng ngày (KHÔNG đụng Google Sheet).

Chạy:
    python tai_lai_du_lieu.py                          # từ đầu tháng hiện tại đến hôm nay
    python tai_lai_du_lieu.py 2026-08-01 2026-08-29    # khoảng ngày tùy chọn

Mỗi ngày 1 file api_data/donhang_YYYY-MM-DD.json; ngày đã có file sẽ được ghi đè
bằng dữ liệu mới nhất. Tự dọn file của tháng cũ (giữ 2 tháng gần nhất).
"""
import sys
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import config
import luu_tru
from doanh_thu import load_staff
from pancake_client import PancakeClient, PancakeError
from thuong_thang import fetch_range_data

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def main() -> None:
    api_key = config.require_api_key()
    shop_id = config.PANCAKE_SHOP_ID
    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()

    # Cùng nguyên tắc "trễ 2 ngày" với thuong_thang.py: ngày chưa đủ 2 ngày chờ
    # chốt thì CHƯA gọi (dữ liệu còn thay đổi) - sẽ được gọi đúng ngày của nó
    han_chot = today - timedelta(days=2)

    args = sys.argv[1:]
    try:
        if len(args) >= 2:
            start_day, end_day = date.fromisoformat(args[0]), date.fromisoformat(args[1])
        elif len(args) == 1:
            start_day, end_day = date.fromisoformat(args[0]), han_chot
        else:
            start_day, end_day = today.replace(day=1), han_chot
    except ValueError:
        raise SystemExit("Ngày không hợp lệ. Ví dụ: python tai_lai_du_lieu.py 2026-08-01 2026-08-27")

    if end_day > han_chot:
        print(f"CHÚ Ý: ngày kết thúc lùi về {han_chot.strftime('%d/%m/%Y')} "
              f"(các ngày sau đó chưa đủ 2 ngày chờ chốt, sẽ được gọi đúng lịch).")
        end_day = han_chot
    if start_day > end_day:
        raise SystemExit("Không có ngày nào đủ điều kiện để tải.")

    client = PancakeClient(api_key)
    load_staff(client, shop_id)      # đồng thời lưu api_data/nhanvien.json
    print(f"Đang tải đơn hàng {start_day.strftime('%d/%m')} - {end_day.strftime('%d/%m/%Y')} ...")
    try:
        fetch_range_data(client, shop_id, start_day, end_day, tz)
    except PancakeError as e:
        raise SystemExit(f"Lỗi khi lấy đơn: {e}")

    files = sorted(luu_tru.API_DATA_DIR.glob("donhang_*.json"))
    print(f"Hoàn tất. api_data/ hiện có {len(files)} file ngày "
          f"({files[0].name} ... {files[-1].name}) + nhanvien.json")


if __name__ == "__main__":
    main()
