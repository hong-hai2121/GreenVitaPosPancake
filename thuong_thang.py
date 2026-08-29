# -*- coding: utf-8 -*-
"""Bảng "Đề Xuất chi thưởng GR" theo tháng, đẩy lên Google Sheet.

Chạy:
    python thuong_thang.py                    # tháng hiện tại: chỉ tính lại 5 ngày gần nhất,
                                              # các ngày cũ giữ nguyên số trên sheet
    python thuong_thang.py --ngay 10          # đổi cửa sổ cập nhật thành 10 ngày
    python thuong_thang.py --ca-thang         # tính lại toàn bộ tháng (khi nghỉ chạy lâu ngày)
    python thuong_thang.py 2026-07            # tháng cũ: luôn tính lại cả tháng
    python thuong_thang.py --bophan sale      # (mặc định đã là sale)

Cấu trúc bảng (1 tab / tháng trong Google Sheet):
    Đề Xuất chi thưởng GR Tháng 8.2026
    STT | Họ và tên | Bộ phận | 01/08/2026 | 02/08/2026 | ... | Tổng tháng
    ... mỗi nhân viên 1 dòng, ô = tiền thưởng ngày (trống nếu không đạt mốc)
    Dòng cuối: Tổng theo từng ngày (công thức SUM - tự cập nhật trên sheet)

Thưởng ngày tính từ DOANH SỐ ngày của từng nhân viên theo mốc BONUS_TIERS
trong config.py (đạt mốc >= nào cao nhất thì hưởng mốc đó).

Logic tiết kiệm API: chạy hàng ngày chỉ gọi Pancake lấy đơn của N ngày gần nhất
(mặc định 5, ví dụ hôm nay 29 thì lấy 24->29); số thưởng các ngày trước đó được
đọc lại từ chính tab trên Google Sheet và giữ nguyên.
"""
import argparse
import calendar
import sys
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

import config
import google_sheet
from doanh_thu import clean_name, load_staff, matched_departments, seller_of, vnd
from pancake_client import PancakeClient, PancakeError

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def bonus_for(doanh_so: int, day: date) -> int:
    """Tiền thưởng theo mốc doanh số ngày (đạt mốc >= nào cao nhất hưởng mốc đó).

    Chủ nhật dùng bảng mốc riêng BONUS_TIERS_SUNDAY (mức thưởng cao hơn).
    """
    tiers = config.BONUS_TIERS_SUNDAY if day.weekday() == 6 else config.BONUS_TIERS
    for nguong, thuong in sorted(tiers, reverse=True):
        if doanh_so >= nguong:
            return thuong
    return 0


def fetch_month_sales(
    client: PancakeClient, shop_id: str, first: date, last: date, tz: ZoneInfo
) -> dict[str, dict[date, int]]:
    """Doanh số theo (nhân viên, ngày): cộng total_price các đơn chốt + đơn hoàn."""
    start_ts = int(datetime.combine(first, time.min, tzinfo=tz).timestamp())
    end_ts = int(datetime.combine(last, time.max, tzinfo=tz).timestamp())
    sales: dict[str, dict[date, int]] = {}
    counted_statuses = config.CLOSED_STATUSES | config.RETURN_STATUSES
    # page_size=1000 (mức tối đa server chấp nhận) -> cả tháng ~10 call thay vì ~93
    for o in client.iter_orders(shop_id, start_ts, end_ts, page_size=1000):
        if o.get("status") not in counted_statuses:
            continue
        uid, _name = seller_of(o)
        if not uid:
            continue
        inserted = o.get("inserted_at") or ""
        try:
            day = datetime.fromisoformat(inserted).date()
        except ValueError:
            continue
        sales.setdefault(uid, {})
        sales[uid][day] = sales[uid].get(day, 0) + (o.get("total_price") or 0)
    return sales


def parse_old_bonus(old_values: list[list[str]] | None) -> dict[tuple[str, date], int]:
    """Đọc số thưởng cũ từ tab trên sheet: {(tên NV, ngày): tiền thưởng}."""
    result: dict[tuple[str, date], int] = {}
    if not old_values or len(old_values) < 3:
        return result
    header = old_values[1]
    col_dates: dict[int, date] = {}
    for i, h in enumerate(header):
        try:
            col_dates[i] = datetime.strptime(h.strip(), "%d/%m/%Y").date()
        except ValueError:
            continue
    for row in old_values[2:]:
        if len(row) < 2 or not row[1].strip() or row[1].strip() == "Tổng":
            continue
        name = row[1].strip()
        for i, d in col_dates.items():
            if i < len(row):
                digits = row[i].replace(".", "").replace(",", "").strip()
                if digits.isdigit() and int(digits) > 0:
                    result[(name, d)] = int(digits)
    return result


def build_table(
    month_label: str,
    days: list[date],
    roster: list[tuple[str, dict]],
    sales: dict[str, dict[date, int]],
    window_start: date,
    old_bonus: dict[tuple[str, date], int],
) -> list[list]:
    """Dựng ma trận giá trị cho tab Google Sheet.

    Ngày < window_start: lấy lại số cũ trên sheet; ngày >= window_start: tính mới.
    """
    n_day_cols = len(days)
    first_day_col = 4                      # cột D (sau STT, Họ và tên, Bộ phận)
    total_col = google_sheet._col_letter(first_day_col + n_day_cols)
    header = ["STT", "Họ và tên", "Bộ phận"] + [d.strftime("%d/%m/%Y") for d in days] + ["Tổng tháng"]

    values: list[list] = [
        [f"Đề Xuất chi thưởng GR Tháng {month_label}"],
        header,
    ]
    for idx, (uid, info) in enumerate(roster, start=1):
        row_num = len(values) + 1
        cells = []
        for d in days:
            if d < window_start:
                cells.append(old_bonus.get((info["name"], d), ""))
            else:
                b = bonus_for(sales.get(uid, {}).get(d, 0), d)
                cells.append(b if b else "")
        first_cell = f"{google_sheet._col_letter(first_day_col)}{row_num}"
        last_cell = f"{google_sheet._col_letter(first_day_col + n_day_cols - 1)}{row_num}"
        values.append([idx, info["name"], info["dept"]] + cells + [f"=SUM({first_cell}:{last_cell})"])

    # Dòng tổng theo ngày (công thức SUM để sheet tự cập nhật khi sửa tay)
    first_data_row, last_data_row = 3, len(values)
    total_row: list = ["", "Tổng", ""]
    for i in range(n_day_cols + 1):
        col = google_sheet._col_letter(first_day_col + i)
        total_row.append(f"=SUM({col}{first_data_row}:{col}{last_data_row})")
    values.append(total_row)
    return values


def parse_args() -> tuple[int, int, str, int, bool]:
    parser = argparse.ArgumentParser(description='Bảng "Đề Xuất chi thưởng GR" theo tháng')
    parser.add_argument("thang", nargs="?", help="Tháng dạng YYYY-MM (mặc định: tháng hiện tại)")
    parser.add_argument("--bophan", metavar="TUKHOA", default="sale",
                        help='Bộ phận có tên chứa từ khóa (mặc định: sale)')
    parser.add_argument("--ngay", type=int, default=5, metavar="N",
                        help="Chỉ tính lại N ngày gần nhất (mặc định 5), ngày cũ giữ số trên sheet")
    parser.add_argument("--ca-thang", action="store_true", dest="ca_thang",
                        help="Tính lại toàn bộ tháng (dùng khi nghỉ chạy nhiều ngày hoặc muốn đối soát)")
    args = parser.parse_args()

    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()
    if args.thang:
        try:
            year, month = map(int, args.thang.split("-"))
            date(year, month, 1)
        except ValueError:
            raise SystemExit("Tháng không hợp lệ. Dùng dạng YYYY-MM, ví dụ: python thuong_thang.py 2026-07")
    else:
        year, month = today.year, today.month
    return year, month, args.bophan, args.ngay, args.ca_thang


def main() -> None:
    api_key = config.require_api_key()
    shop_id = config.PANCAKE_SHOP_ID
    if not shop_id:
        raise SystemExit("Chưa có PANCAKE_SHOP_ID trong .env.")

    year, month, dept_keyword, window_days, ca_thang = parse_args()
    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()

    first = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])
    if first > today:
        raise SystemExit(f"Tháng {month}/{year} chưa bắt đầu.")
    last = min(month_end, today)          # tháng hiện tại: cập nhật đến hôm nay
    days = [date(year, month, d) for d in range(1, last.day + 1)]
    tab_title = f"Thưởng GR T{month:02d}.{year}"

    # Xác định cửa sổ cập nhật: mặc định chỉ N ngày gần nhất, số cũ đọc lại từ sheet.
    # Tính lại CẢ THÁNG khi: --ca-thang, tháng cũ, hoặc tab chưa có trên sheet.
    old_values = google_sheet.read_table(tab_title)
    old_bonus = parse_old_bonus(old_values)
    full_month = ca_thang or last < today or old_values is None or not old_bonus
    if full_month:
        window_start = first
    else:
        window_start = max(first, today - timedelta(days=window_days))

    client = PancakeClient(api_key)
    staff = load_staff(client, shop_id)
    dept_filter = matched_departments(staff, dept_keyword)
    if not dept_filter:
        raise SystemExit(f"Không có bộ phận nào chứa '{dept_keyword}'.")

    roster = sorted(
        ((uid, info) for uid, info in staff.items() if info["dept"] in dept_filter),
        key=lambda x: (x[1]["dept"], x[1]["name"]),
    )

    print(f"Bộ phận: {', '.join(sorted(dept_filter))} ({len(roster)} nhân viên)")
    if full_month:
        print(f"Tính lại CẢ THÁNG: lấy đơn {first.strftime('%d/%m')} - {last.strftime('%d/%m/%Y')} ...")
    else:
        print(f"Chỉ cập nhật {window_start.strftime('%d/%m')} - {last.strftime('%d/%m/%Y')} "
              f"({window_days} ngày gần nhất); các ngày trước giữ số trên sheet. "
              "(--ca-thang để tính lại cả tháng)")
    try:
        sales = fetch_month_sales(client, shop_id, window_start, last, tz)
    except PancakeError as e:
        raise SystemExit(f"Lỗi khi lấy đơn: {e}")

    month_label = f"{month}.{year}"
    values = build_table(month_label, days, roster, sales, window_start, old_bonus)

    # Vùng ô tiền: từ cột D dòng 3 tới cột Tổng tháng dòng Tổng
    end_col = google_sheet._col_letter(3 + len(days) + 1)
    money_range = f"D3:{end_col}{len(values)}"
    tab_title = f"Thưởng GR T{month:02d}.{year}"
    url = google_sheet.write_table(tab_title, values, money_range=money_range)

    # Tóm tắt ra màn hình
    print(f"\nTHƯỞNG GR THÁNG {month_label} (tính đến {last.strftime('%d/%m')}):")
    total_all = 0
    for row in values[2:-1]:
        bonus_total = sum(v for v in row[3:-1] if isinstance(v, int))
        total_all += bonus_total
        if bonus_total:
            print(f"  {row[1]:<32} {vnd(bonus_total)}")
    print(f"  {'TỔNG':<32} {vnd(total_all)}")
    print(f"\nĐã ghi tab '{tab_title}': {url}")


if __name__ == "__main__":
    main()
