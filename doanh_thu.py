# -*- coding: utf-8 -*-
"""Lấy doanh thu theo ngày từ Pancake POS, lọc theo bộ phận, xuất Excel.

Chạy:
    python doanh_thu.py                              # hôm nay, tất cả nhân viên
    python doanh_thu.py 2026-08-28                   # ngày cụ thể
    python doanh_thu.py 2026-08-28 --bophan sale     # chỉ nhân viên thuộc bộ phận có chữ "sale"
    python doanh_thu.py 2026-08-01 2026-08-28 --bophan sale   # khoảng ngày

Bộ lọc --bophan không phân biệt hoa/thường: "sale" khớp "SALE OCP", "Sale NT", "Sale Nghỉ"...
Kết quả xuất ra file Excel trong thư mục output/ (sheet Tổng hợp NV + sheet Chi tiết đơn).
"""
import argparse
import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config
from pancake_client import PancakeClient, PancakeError

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def vnd(amount) -> str:
    try:
        return f"{int(amount):,.0f} đ".replace(",", ".")
    except (TypeError, ValueError):
        return "0 đ"


# ----------------------------------------------------------------------
# Nhân viên & bộ phận
# ----------------------------------------------------------------------
def load_user_departments(client: PancakeClient, shop_id: str) -> dict[str, str]:
    """Trả về map: user_id -> tên bộ phận."""
    users = client.get_users(shop_id)
    mapping = {}
    for u in users:
        uid = u.get("user_id") or (u.get("user") or {}).get("id")
        dept = (u.get("department") or {}).get("name") or "(chưa gán bộ phận)"
        if uid:
            mapping[uid] = dept
    return mapping


def matched_departments(user_depts: dict[str, str], keyword: str) -> set[str]:
    """Các bộ phận có tên chứa keyword (không phân biệt hoa/thường)."""
    kw = keyword.lower()
    return {d for d in set(user_depts.values()) if kw in d.lower()}


def seller_of(order: dict) -> tuple[str, str]:
    """(user_id, tên) của nhân viên phụ trách đơn."""
    seller = order.get("assigning_seller") or {}
    if seller.get("id") or seller.get("name"):
        return seller.get("id") or "", seller.get("name") or "(không rõ)"
    creator = order.get("creator") or {}
    return creator.get("id") or "", creator.get("name") or "(không rõ)"


# ----------------------------------------------------------------------
# Lấy & tổng hợp dữ liệu 1 ngày
# ----------------------------------------------------------------------
def fetch_day(
    client: PancakeClient,
    shop_id: str,
    day: date,
    tz: ZoneInfo,
    user_depts: dict[str, str],
    dept_filter: set[str] | None,
) -> dict:
    start_ts = int(datetime.combine(day, time.min, tzinfo=tz).timestamp())
    end_ts = int(datetime.combine(day, time.max, tzinfo=tz).timestamp())

    rows = []
    revenue = 0            # doanh thu theo bộ lọc
    revenue_all = 0        # doanh thu toàn shop (tham khảo)
    counted = 0
    total_orders = 0
    by_status: dict[str, int] = {}
    by_seller: dict[str, dict] = {}

    for o in client.iter_orders(shop_id, start_ts, end_ts):
        total_orders += 1
        status_code = o.get("status")
        status = config.ORDER_STATUS.get(status_code, f"Mã {status_code}")
        by_status[status] = by_status.get(status, 0) + 1

        total_price = o.get("total_price") or 0
        excluded_status = status_code in config.EXCLUDED_STATUSES

        uid, name = seller_of(o)
        dept = user_depts.get(uid, "(không rõ)")
        in_filter = dept_filter is None or dept in dept_filter

        if not excluded_status:
            revenue_all += total_price
            if in_filter:
                revenue += total_price
                counted += 1
                s = by_seller.setdefault(name, {"orders": 0, "revenue": 0, "dept": dept})
                s["orders"] += 1
                s["revenue"] += total_price

        if excluded_status:
            tinh = "Không (hủy/hoàn/xóa)"
        elif not in_filter:
            tinh = "Không (ngoài bộ phận lọc)"
        else:
            tinh = "Có"

        rows.append({
            "ma_don": o.get("id"),
            "thoi_gian_tao": o.get("inserted_at") or "",
            "trang_thai": status,
            "tinh_doanh_thu": tinh,
            "tong_tien": total_price,
            "cod": o.get("cod") or 0,
            "phi_ship": o.get("shipping_fee") or 0,
            "giam_gia": o.get("total_discount") or 0,
            "nhan_vien": name,
            "bo_phan": dept,
            "khach_hang": (o.get("customer") or {}).get("name") or o.get("bill_full_name") or "",
        })

    return {
        "day": day,
        "orders": rows,
        "revenue": revenue,
        "revenue_all": revenue_all,
        "counted": counted,
        "total_orders": total_orders,
        "by_status": by_status,
        "by_seller": by_seller,
    }


# ----------------------------------------------------------------------
# In báo cáo ra màn hình
# ----------------------------------------------------------------------
def print_report(result: dict, dept_filter: set[str] | None) -> None:
    day = result["day"]
    print("=" * 62)
    print(f"DOANH THU NGÀY {day.strftime('%d/%m/%Y')}")
    if dept_filter is not None:
        print(f"Bộ phận lọc: {', '.join(sorted(dept_filter)) or '(không khớp bộ phận nào)'}")
    print("=" * 62)
    print(f"Tổng số đơn trong ngày:  {result['total_orders']}")
    print(f"Số đơn tính doanh thu:   {result['counted']}")
    print(f"DOANH THU:               {vnd(result['revenue'])}")
    if dept_filter is not None:
        print(f"(Doanh thu toàn shop:    {vnd(result['revenue_all'])})")

    if result["by_seller"]:
        print("\nTheo nhân viên:")
        for name, s in sorted(result["by_seller"].items(), key=lambda x: -x[1]["revenue"]):
            print(f"  {name:<32} {s['dept']:<15} {s['orders']:>3} đơn   {vnd(s['revenue'])}")
    print()


# ----------------------------------------------------------------------
# Xuất Excel
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F7A4D")
HEADER_FONT = Font(bold=True, color="FFFFFF")
MONEY_FMT = "#,##0"

DETAIL_COLS = [
    ("Mã đơn", "ma_don", 10),
    ("Thời gian tạo", "thoi_gian_tao", 20),
    ("Trạng thái", "trang_thai", 16),
    ("Tính doanh thu", "tinh_doanh_thu", 24),
    ("Tổng tiền", "tong_tien", 14),
    ("COD", "cod", 14),
    ("Phí ship", "phi_ship", 12),
    ("Giảm giá", "giam_gia", 12),
    ("Nhân viên", "nhan_vien", 32),
    ("Bộ phận", "bo_phan", 16),
    ("Khách hàng", "khach_hang", 24),
]


def _write_header(ws, titles):
    for col, title in enumerate(titles, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def export_excel(result: dict, dept_filter: set[str] | None) -> Path:
    OUTPUT_DIR.mkdir(exist_ok=True)
    suffix = "_locbophan" if dept_filter is not None else ""
    path = OUTPUT_DIR / f"doanhthu_{result['day'].isoformat()}{suffix}.xlsx"

    wb = Workbook()

    # --- Sheet 1: Tổng hợp theo nhân viên ---
    ws = wb.active
    ws.title = "Tổng hợp NV"
    _write_header(ws, ["Nhân viên", "Bộ phận", "Số đơn", "Doanh thu"])
    row = 2
    for name, s in sorted(result["by_seller"].items(), key=lambda x: -x[1]["revenue"]):
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=s["dept"])
        ws.cell(row=row, column=3, value=s["orders"])
        c = ws.cell(row=row, column=4, value=s["revenue"])
        c.number_format = MONEY_FMT
        row += 1
    # Dòng tổng
    total_cell = ws.cell(row=row, column=1, value="TỔNG")
    total_cell.font = Font(bold=True)
    ws.cell(row=row, column=3, value=result["counted"]).font = Font(bold=True)
    c = ws.cell(row=row, column=4, value=result["revenue"])
    c.font = Font(bold=True)
    c.number_format = MONEY_FMT
    for col, width in zip(range(1, 5), (35, 18, 10, 16)):
        ws.column_dimensions[get_column_letter(col)].width = width

    # --- Sheet 2: Chi tiết đơn ---
    ws2 = wb.create_sheet("Chi tiết đơn")
    _write_header(ws2, [t for t, _, _ in DETAIL_COLS])
    for r, order in enumerate(result["orders"], start=2):
        for col, (_, key, _) in enumerate(DETAIL_COLS, start=1):
            cell = ws2.cell(row=r, column=col, value=order[key])
            if key in ("tong_tien", "cod", "phi_ship", "giam_gia"):
                cell.number_format = MONEY_FMT
    for col, (_, _, width) in enumerate(DETAIL_COLS, start=1):
        ws2.column_dimensions[get_column_letter(col)].width = width
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(DETAIL_COLS))}{len(result['orders']) + 1}"

    wb.save(path)
    return path


# ----------------------------------------------------------------------
def parse_args() -> tuple[date, date, str | None]:
    parser = argparse.ArgumentParser(description="Doanh thu theo ngày từ Pancake POS")
    parser.add_argument("start", nargs="?", help="Ngày bắt đầu YYYY-MM-DD (mặc định: hôm nay)")
    parser.add_argument("end", nargs="?", help="Ngày kết thúc YYYY-MM-DD (mặc định: bằng ngày bắt đầu)")
    parser.add_argument("--bophan", metavar="TUKHOA", default=None,
                        help='Chỉ tính nhân viên thuộc bộ phận có tên chứa từ khóa này, ví dụ: --bophan sale')
    args = parser.parse_args()

    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()
    try:
        start_day = date.fromisoformat(args.start) if args.start else today
        end_day = date.fromisoformat(args.end) if args.end else start_day
    except ValueError:
        raise SystemExit("Ngày không hợp lệ. Dùng định dạng YYYY-MM-DD, ví dụ: python doanh_thu.py 2026-08-28")
    return start_day, end_day, args.bophan


def main() -> None:
    api_key = config.require_api_key()
    shop_id = config.PANCAKE_SHOP_ID
    if not shop_id:
        raise SystemExit("Chưa có PANCAKE_SHOP_ID trong .env. Chạy 'python test_connection.py' để xem danh sách shop_id.")

    start_day, end_day, dept_keyword = parse_args()
    tz = ZoneInfo(config.TIMEZONE)
    client = PancakeClient(api_key)

    user_depts = load_user_departments(client, shop_id)

    dept_filter: set[str] | None = None
    if dept_keyword:
        dept_filter = matched_departments(user_depts, dept_keyword)
        if not dept_filter:
            all_depts = ", ".join(sorted(set(user_depts.values())))
            raise SystemExit(
                f"Không có bộ phận nào chứa '{dept_keyword}'.\nCác bộ phận hiện có: {all_depts}"
            )
        print(f"Bộ phận khớp '{dept_keyword}': {', '.join(sorted(dept_filter))}\n")

    grand_total = 0
    day = start_day
    while day <= end_day:
        try:
            result = fetch_day(client, shop_id, day, tz, user_depts, dept_filter)
        except PancakeError as e:
            raise SystemExit(f"Lỗi khi lấy đơn ngày {day}: {e}")
        print_report(result, dept_filter)
        xlsx_path = export_excel(result, dept_filter)
        print(f"Đã xuất Excel: {xlsx_path}\n")
        grand_total += result["revenue"]
        day += timedelta(days=1)

    if start_day != end_day:
        print("=" * 62)
        print(f"TỔNG DOANH THU {start_day.strftime('%d/%m')} - {end_day.strftime('%d/%m/%Y')}: {vnd(grand_total)}")


if __name__ == "__main__":
    main()
