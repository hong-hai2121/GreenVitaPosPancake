# -*- coding: utf-8 -*-
"""Kiểm tra kết nối Pancake POS.

Chạy:  python test_connection.py

- Liệt kê các shop mà API key truy cập được (kèm shop_id để điền vào .env)
- Nếu đã có PANCAKE_SHOP_ID trong .env: thử lấy vài đơn hàng hôm nay để xác nhận quyền đọc đơn.
"""
import sys
from datetime import datetime, time
from zoneinfo import ZoneInfo

import config
from pancake_client import PancakeClient, PancakeError

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def main() -> None:
    api_key = config.require_api_key()
    client = PancakeClient(api_key)

    print("Đang kết nối tới Pancake POS ...")
    try:
        shops = client.get_shops()
    except PancakeError as e:
        raise SystemExit(f"KẾT NỐI THẤT BẠI: {e}")

    if not shops:
        print("Kết nối OK nhưng API key này không thấy shop nào.")
        return

    print(f"KẾT NỐI THÀNH CÔNG. Tìm thấy {len(shops)} shop:\n")
    for shop in shops:
        print(f"  - shop_id = {shop.get('id')}  |  Tên: {shop.get('name')}")
        for page in shop.get("pages") or []:
            print(f"      trang bán hàng: {page.get('name')} ({page.get('platform', '')})")
    print()

    shop_id = config.PANCAKE_SHOP_ID
    if not shop_id:
        print("=> Chưa có PANCAKE_SHOP_ID trong .env. Hãy chọn shop_id ở trên và điền vào .env.")
        return

    # Thử đọc đơn hàng hôm nay của shop đã cấu hình
    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()
    start_ts = int(datetime.combine(today, time.min, tzinfo=tz).timestamp())
    end_ts = int(datetime.combine(today, time.max, tzinfo=tz).timestamp())

    print(f"Thử lấy đơn hàng hôm nay ({today}) của shop {shop_id} ...")
    try:
        data = client.get_orders_page(shop_id, start_ts, end_ts, page_size=5)
    except PancakeError as e:
        raise SystemExit(f"Đọc đơn hàng THẤT BẠI: {e}")

    orders = data.get("data") or data.get("orders") or []
    total = data.get("total_entries", "?")
    print(f"OK. Hôm nay có {total} đơn (hiển thị {len(orders)} đơn đầu):")
    for o in orders:
        status = config.ORDER_STATUS.get(o.get("status"), o.get("status"))
        print(f"  - Đơn #{o.get('id')} | trạng thái: {status} | tổng tiền: {o.get('total_price')}")
    print("\nMọi thứ sẵn sàng. Chạy tiếp:  python doanh_thu.py")


if __name__ == "__main__":
    main()
