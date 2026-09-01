# -*- coding: utf-8 -*-
"""Cập nhật CẢ 3 BẢNG thưởng trên Google Sheet - logic "chốt ngày, chốt sổ".

Chạy:
    python thuong_thang.py            # cập nhật tháng hiện tại + tự chốt sổ tháng trước
    python thuong_thang.py 2026-08    # tháng đã qua: CHỐT SỔ (tính lại cả tháng rồi khóa)

Nguyên tắc:
    1. TRỄ 2 NGÀY: hôm nay 29 thì bảng chỉ hiển thị đến ngày 27 - hai ngày cuối
       trạng thái đơn còn thay đổi nhiều nên chưa đưa vào.
    2. NGÀY ĐÃ LÊN BẢNG = ĐÃ CHỐT: mỗi lần chạy chỉ gọi API lấy đơn của các NGÀY MỚI
       rồi nối thêm cột; số của các ngày cũ giữ nguyên như trên sheet.
    3. CHỐT SỔ CUỐI THÁNG: sang tháng mới (từ mùng 2), script tự gọi lại API MỘT LẦN
       trọn tháng trước để sửa thưởng lần cuối (bắt đơn hoàn/hủy muộn), đóng dấu
       "ĐÃ CHỐT SỔ" lên tiêu đề cả 3 tab - từ đó không lần chạy nào sửa được nữa.

3 tab: "Thưởng Sale GR", "Thưởng CSKH GR" (ma trận thưởng ngày) và
"BC02 Thưởng DS Sale- CSKH" (thưởng doanh số tháng, 2 cột nhập tay được giữ nguyên).
"""
import argparse
import calendar
import sys
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

import config
import google_sheet
from bc02_thuong_ds import build_table as bc02_build_table
from bc02_thuong_ds import dem_hoan_thang_truoc
from bc02_thuong_ds import parse_old_manual as bc02_parse_old_manual
from bc02_thuong_ds import read_gr_bonus_totals
from doanh_thu import load_staff, matched_departments, seller_of, vnd
from pancake_client import PancakeClient, PancakeError

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

GROUPS = [("sale", "Sale"), ("cskh", "CSKH")]
LOCK_MARK = "ĐÃ CHỐT SỔ"
SETTLE_DELAY_DAYS = 2      # số ngày trễ trước khi một ngày được đưa vào bảng


def bonus_for(doanh_so: int, day: date, tiers_cfg: dict) -> int:
    """Tiền thưởng theo mốc doanh số ngày; Chủ nhật dùng bảng mốc riêng."""
    tiers = tiers_cfg["sunday"] if day.weekday() == 6 else tiers_cfg["weekday"]
    for nguong, thuong in sorted(tiers, reverse=True):
        if doanh_so >= nguong:
            return thuong
    return 0


# ----------------------------------------------------------------------
# Lấy dữ liệu Pancake (1 lần cho cả 3 bảng)
# ----------------------------------------------------------------------
def fetch_range_data(
    client: PancakeClient, shop_id: str, start_day: date, end_day: date, tz: ZoneInfo
) -> dict[str, dict[date, dict]]:
    """Quét đơn trong [start_day, end_day], gom theo (nhân viên, ngày):
    {"ds_all": doanh số (chốt+hoàn), "chot": số đơn chốt, "hoan": số đơn hoàn,
     "ds_chot": tổng tiền đơn chốt}

    Đồng thời LƯU TRỮ đơn thô của từng ngày vào api_data/donhang_<ngày>.json
    (ghi đè ngày được gọi lại; tự dọn file của tháng cũ, giữ 2 tháng gần nhất)."""
    import luu_tru

    start_ts = int(datetime.combine(start_day, time.min, tzinfo=tz).timestamp())
    end_ts = int(datetime.combine(end_day, time.max, tzinfo=tz).timestamp())
    data: dict[str, dict[date, dict]] = {}
    # Chuẩn bị đủ mọi ngày trong khoảng để ngày không có đơn vẫn có file (rỗng)
    orders_by_day: dict[date, list] = {
        start_day + timedelta(days=i): []
        for i in range((end_day - start_day).days + 1)
    }
    for o in client.iter_orders(shop_id, start_ts, end_ts, page_size=config.ORDERS_PAGE_SIZE):
        try:
            day = datetime.fromisoformat(o.get("inserted_at") or "").date()
        except ValueError:
            continue
        orders_by_day.setdefault(day, []).append(o)

        status = o.get("status")
        if status in config.CLOSED_STATUSES:
            kind = "chot"
        elif status in config.RETURN_STATUSES:
            kind = "hoan"
        else:
            continue
        uid, _ = seller_of(o)
        if not uid:
            continue
        price = o.get("total_price") or 0
        cell = data.setdefault(uid, {}).setdefault(
            day, {"ds_all": 0, "chot": 0, "hoan": 0, "ds_chot": 0})
        cell["ds_all"] += price
        if kind == "chot":
            cell["chot"] += 1
            cell["ds_chot"] += price
        else:
            cell["hoan"] += 1

    luu_tru.luu_don_theo_ngay(orders_by_day)
    luu_tru.don_dep()
    return data


# ----------------------------------------------------------------------
# Đọc lại dữ liệu cũ trên sheet
# ----------------------------------------------------------------------
def is_locked(old_values: list[list[str]] | None) -> bool:
    return bool(old_values and old_values[0] and LOCK_MARK in old_values[0][0])


def parse_header_days(old_values: list[list[str]] | None) -> set[date]:
    """Các ngày đã có cột trên tab (đã chốt)."""
    days: set[date] = set()
    if not old_values or len(old_values) < 2:
        return days
    for h in old_values[1]:
        try:
            days.add(datetime.strptime(h.strip(), "%d/%m/%Y").date())
        except ValueError:
            continue
    return days


def parse_old_bonus(old_values: list[list[str]] | None) -> dict[tuple[str, date], int]:
    """Số thưởng cũ trên tab ma trận: {(tên NV, ngày): tiền thưởng}."""
    result: dict[tuple[str, date], int] = {}
    if not old_values or len(old_values) < 3:
        return result
    col_dates: dict[int, date] = {}
    for i, h in enumerate(old_values[1]):
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


def parse_old_bc02_stats(old_values: list[list[str]] | None) -> dict[str, dict]:
    """Số liệu cũ trên BC02: {tên NV: {"chot": n, "hoan": n, "ds": tiền}}."""
    result: dict[str, dict] = {}
    if not old_values or len(old_values) < 4:
        return result
    header = old_values[1]
    try:
        i_name = header.index("Tên")
        i_chot = header.index("Đơn chốt")
        i_hoan = header.index("Đơn hoàn tháng này")
        i_ds = header.index("DS bán hàng")
    except ValueError:
        return result

    def num(row, i):
        digits = row[i].replace(".", "").replace(",", "").strip() if i < len(row) else ""
        return int(digits) if digits.isdigit() else 0

    for row in old_values[2:]:
        name = row[i_name].strip() if i_name < len(row) else ""
        if not name or name == "Tổng":
            continue
        result[name] = {"chot": num(row, i_chot), "hoan": num(row, i_hoan),
                        "ds": num(row, i_ds)}
    return result


# ----------------------------------------------------------------------
# Dựng và ghi bảng
# ----------------------------------------------------------------------
def build_matrix(
    month_label: str, days: list[date], roster: list[tuple[str, dict]],
    data: dict, tiers_cfg: dict, group_label: str,
    settled_days: set[date], old_bonus: dict, title_suffix: str = "",
) -> list[list]:
    """Ma trận thưởng ngày: ngày đã chốt lấy số cũ trên sheet, ngày mới tính từ data."""
    n_day_cols = len(days)
    first_day_col = 4
    header = ["STT", "Họ và tên", "Bộ phận"] + [d.strftime("%d/%m/%Y") for d in days] + ["Tổng tháng"]
    values: list[list] = [
        [f"Đề Xuất chi thưởng GR Tháng {month_label} - Bộ phận {group_label}{title_suffix}"],
        header,
    ]
    for idx, (uid, info) in enumerate(roster, start=1):
        row_num = len(values) + 1
        cells = []
        for d in days:
            if d in settled_days:
                cells.append(old_bonus.get((info["name"], d), ""))
            else:
                ds_all = data.get(uid, {}).get(d, {}).get("ds_all", 0)
                b = bonus_for(ds_all, d, tiers_cfg)
                cells.append(b if b else "")
        first_cell = f"{google_sheet._col_letter(first_day_col)}{row_num}"
        last_cell = f"{google_sheet._col_letter(first_day_col + n_day_cols - 1)}{row_num}"
        values.append([idx, info["name"], info["dept"]] + cells + [f"=SUM({first_cell}:{last_cell})"])

    first_data_row, last_data_row = 3, len(values)
    total_row: list = ["", "Tổng", ""]
    for i in range(n_day_cols + 1):
        col = google_sheet._col_letter(first_day_col + i)
        total_row.append(f"=SUM({col}{first_data_row}:{col}{last_data_row})")
    values.append(total_row)
    return values


def group_roster(staff: dict, keyword: str) -> list[tuple[str, dict]]:
    dept_filter = matched_departments(staff, keyword)
    return sorted(
        ((uid, info) for uid, info in staff.items() if info["dept"] in dept_filter),
        key=lambda x: (x[1]["dept"], x[1]["name"]),
    )


def run_month(client: PancakeClient, shop_id: str, staff: dict, tz: ZoneInfo,
              today: date, year: int, month: int, finalize: bool,
              only_if_exists: bool = False) -> None:
    """Cập nhật (hoặc chốt sổ) cả 3 tab của 1 tháng."""
    first = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])

    tabs = {kw: f"Thưởng {label} GR T{month:02d}.{year}" for kw, label in GROUPS}
    bc02_tab = f"BC02 Thưởng DS Sale- CSKH T{month:02d}.{year}"
    old = {kw: google_sheet.read_table(t) for kw, t in tabs.items()}
    old_bc02 = google_sheet.read_table(bc02_tab)

    if is_locked(old["sale"]) or is_locked(old_bc02):
        print(f"Tháng {month:02d}.{year} {LOCK_MARK} - giữ nguyên, không sửa.")
        return
    if only_if_exists and old["sale"] is None:
        return

    if finalize:
        cutoff = month_end
        settled_days: set[date] = set()          # tính lại toàn bộ
        title_suffix = f" ({LOCK_MARK} {today.strftime('%d/%m/%Y')})"
        print(f"CHỐT SỔ tháng {month:02d}.{year}: lấy lại đơn CẢ THÁNG để sửa thưởng lần cuối ...")
    else:
        cutoff = today - timedelta(days=SETTLE_DELAY_DAYS)
        if cutoff < first:
            print(f"Tháng {month:02d}.{year}: chưa có ngày nào đủ {SETTLE_DELAY_DAYS} ngày chờ chốt.")
            return
        cutoff = min(cutoff, month_end)
        settled_days = {d for d in parse_header_days(old["sale"]) if d <= cutoff}
        # Nếu tab cũ có cột ngày vượt cutoff (dữ liệu của logic cũ) -> làm lại từ đầu
        if any(d > cutoff for d in parse_header_days(old["sale"])):
            settled_days = set()
        title_suffix = ""

    days = [first + timedelta(days=i) for i in range((cutoff - first).days + 1)]
    new_days = [d for d in days if d not in settled_days]

    if not new_days:
        print(f"Tháng {month:02d}.{year}: đã chốt đến {cutoff.strftime('%d/%m')} - không có ngày mới.")
        return

    print(f"Tháng {month:02d}.{year}: hiển thị đến {cutoff.strftime('%d/%m')}, "
          f"lấy đơn {new_days[0].strftime('%d/%m')} - {new_days[-1].strftime('%d/%m')} ...")
    data = fetch_range_data(client, shop_id, new_days[0], cutoff, tz)

    # --- 2 bảng ma trận thưởng ngày ---
    month_label = f"{month}.{year}"
    for keyword, label in GROUPS:
        roster = group_roster(staff, keyword)
        if not roster:
            continue
        values = build_matrix(month_label, days, roster, data,
                              config.BONUS_TIERS_BY_GROUP[keyword], label,
                              settled_days, parse_old_bonus(old[keyword]), title_suffix)
        end_col = google_sheet._col_letter(3 + len(days) + 1)
        sunday_cols = [3 + i for i, d in enumerate(days) if d.weekday() == 6]
        google_sheet.write_table(tabs[keyword], values,
                                 money_range=f"D3:{end_col}{len(values)}",
                                 sunday_cols=sunday_cols)
        total = sum(sum(v for v in row[3:-1] if isinstance(v, int)) for row in values[2:-1])
        print(f"  [OK] {tabs[keyword]}: {len(roster)} NV, tổng thưởng {vnd(total)}")

    # --- BC02: cộng dồn ngày mới vào số cũ (chốt sổ thì tính lại từ data cả tháng) ---
    roster_all = sorted(
        ((uid, info) for uid, info in staff.items()
         if any(kw in info["dept"].lower() for kw, _ in GROUPS)),
        key=lambda x: (x[1]["dept"], x[1]["name"]),
    )
    old_stats = {} if (finalize or not settled_days) else parse_old_bc02_stats(old_bc02)
    stats: dict[str, dict] = {}
    for uid, info in roster_all:
        base = old_stats.get(info["name"], {"chot": 0, "hoan": 0, "ds": 0})
        s = {"chot": base["chot"], "hoan": base["hoan"], "ds": base["ds"]}
        for d in new_days:
            cell = data.get(uid, {}).get(d)
            if cell:
                s["chot"] += cell["chot"]
                s["hoan"] += cell["hoan"]
                s["ds"] += cell["ds_chot"]
        stats[uid] = s
    # Cột Thưởng của BC02 = Tổng tháng trên 2 tab GR vừa ghi (khớp theo tên NV)
    bonus_totals = read_gr_bonus_totals(month, year)
    # Đơn hoàn tháng trước: TẠM THỜI để 0 theo yêu cầu
    # (bật lại: hoan_truoc = dem_hoan_thang_truoc(client, shop_id, year, month, tz))
    hoan_truoc: dict[str, int] = {}
    values = bc02_build_table(month, year, roster_all, stats,
                              bc02_parse_old_manual(old_bc02), title_suffix,
                              bonus_totals=bonus_totals, hoan_truoc=hoan_truoc)
    google_sheet.write_bc02_table(bc02_tab, values)
    tong_chot = sum(s["chot"] for s in stats.values())
    tong_hoan = sum(s["hoan"] for s in stats.values())
    tong_ht = sum(hoan_truoc.get(uid, 0) for uid, _ in roster_all)
    tong_ds = sum(s["ds"] for s in stats.values())
    print(f"  [OK] {bc02_tab}: {tong_chot} đơn chốt, {tong_hoan} hoàn tháng này, "
          f"{tong_ht} hoàn tháng trước, DS {vnd(tong_ds)}")
    if finalize:
        print(f"  Đã đóng dấu '{LOCK_MARK}' - 3 tab tháng {month:02d}.{year} bị khóa vĩnh viễn.")


def parse_args() -> tuple[int, int, bool]:
    parser = argparse.ArgumentParser(
        description="Cập nhật 3 bảng thưởng theo logic chốt ngày (trễ 2 ngày) + chốt sổ cuối tháng")
    parser.add_argument("thang", nargs="?", help="Tháng YYYY-MM; tháng đã qua sẽ CHỐT SỔ luôn")
    args = parser.parse_args()
    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()
    if args.thang:
        try:
            year, month = map(int, args.thang.split("-"))
            date(year, month, 1)
        except ValueError:
            raise SystemExit("Tháng không hợp lệ. Dùng dạng YYYY-MM, ví dụ: python thuong_thang.py 2026-08")
        return year, month, True
    return today.year, today.month, False


def main() -> None:
    api_key = config.require_api_key()
    shop_id = config.PANCAKE_SHOP_ID
    if not shop_id:
        raise SystemExit("Chưa có PANCAKE_SHOP_ID trong .env.")

    year, month, explicit = parse_args()
    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()
    first = date(year, month, 1)
    if first > today:
        raise SystemExit(f"Tháng {month}/{year} chưa bắt đầu.")
    month_end = date(year, month, calendar.monthrange(year, month)[1])

    client = PancakeClient(api_key)
    staff = load_staff(client, shop_id)

    # Tháng đã qua (gọi tường minh) -> chốt sổ tháng đó
    run_month(client, shop_id, staff, tz, today, year, month,
              finalize=explicit and today > month_end)

    # Chạy mặc định đầu tháng mới -> tự chốt sổ tháng trước (nếu tab tồn tại và chưa khóa)
    if not explicit:
        prev_last = first - timedelta(days=1)
        if today >= prev_last + timedelta(days=SETTLE_DELAY_DAYS):
            run_month(client, shop_id, staff, tz, today,
                      prev_last.year, prev_last.month, finalize=True, only_if_exists=True)

    print(f"\nXong. Sheet: https://docs.google.com/spreadsheets/d/{config.GOOGLE_SHEET_ID}")


if __name__ == "__main__":
    main()
