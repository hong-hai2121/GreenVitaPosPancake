# -*- coding: utf-8 -*-
"""Bảng BC02 "Thưởng DS Sale- CSKH" theo THÁNG, đẩy lên Google Sheet.

Chạy:
    python bc02_thuong_ds.py            # tháng hiện tại
    python bc02_thuong_ds.py 2026-07    # tháng cụ thể

Cấu trúc tab "BC02 Thưởng DS Sale- CSKH T08.2026" (mỗi tháng 1 tab):
    STT | Tên | Bộ phận | Đơn chốt | Đơn hoàn tháng này | DS bán hàng
        | Tỷ lệ hoàn | Thưởng | % Thưởng | Thực nhận
    - Dòng Tổng nằm NGAY DƯỚI header (giống mẫu), nhân viên từ dòng 4
    - Gồm tất cả nhân viên các bộ phận có chữ "sale" hoặc "cskh"

Nguồn số liệu:
    - Đơn chốt / Đơn hoàn / DS bán hàng: tính từ đơn hàng của tháng trên Pancake
      (DS bán hàng = tổng tiền các ĐƠN CHỐT)
    - Tỷ lệ hoàn = Đơn hoàn / (Đơn chốt + Đơn hoàn)  (công thức trên sheet)
    - Thưởng, % Thưởng: NHẬP TAY trên sheet (ô nền vàng) - chạy lại script vẫn
      GIỮ NGUYÊN số đã nhập; Thực nhận = Thưởng x % Thưởng (công thức, % trống = 100%)
"""
import argparse
import calendar
import sys
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

import config
import google_sheet
from doanh_thu import load_staff, seller_of, vnd
from pancake_client import PancakeClient, PancakeError

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HEADER = ["STT", "Tên", "Bộ phận", "Đơn chốt", "Đơn hoàn tháng này",
          "Đơn hoàn tháng trước", "DS bán hàng", "Tỷ lệ hoàn",
          "Thưởng", "% Thưởng", "Thực nhận"]
GROUP_KEYWORDS = ("sale", "cskh")


def dem_hoan_thang_truoc(client: PancakeClient, shop_id: str,
                         year: int, month: int, tz: ZoneInfo) -> dict[str, int]:
    """Đếm theo nhân viên: đơn tạo THÁNG TRƯỚC bị chuyển sang hoàn TRONG tháng (year, month).

    Cách làm (1 call API): lấy các đơn tạo tháng trước hiện đang ở trạng thái hoàn
    (filter_status 4/5), rồi đọc status_history của từng đơn để biết thời điểm
    chuyển sang hoàn có rơi vào tháng này không.
    """
    first_this = date(year, month, 1)
    last_this = date(year, month, calendar.monthrange(year, month)[1])
    prev_last = first_this - timedelta(days=1)
    prev_first = prev_last.replace(day=1)
    start_ts = int(datetime.combine(prev_first, time.min, tzinfo=tz).timestamp())
    end_ts = int(datetime.combine(prev_last, time.max, tzinfo=tz).timestamp())

    counts: dict[str, int] = {}
    page = 1
    while True:
        data = client.get_orders_page(shop_id, start_ts, end_ts, page_number=page,
                                      page_size=config.ORDERS_PAGE_SIZE,
                                      filter_status=[4, 5])
        orders = data.get("data") or []
        for o in orders:
            # Thời điểm ĐẦU TIÊN đơn chuyển sang trạng thái hoàn
            t_hoan = None
            for h in (o.get("status_history") or []):
                if h.get("status") in (4, 5) and h.get("updated_at"):
                    if t_hoan is None or h["updated_at"] < t_hoan:
                        t_hoan = h["updated_at"]
            if not t_hoan:
                continue
            try:
                d = datetime.fromisoformat(t_hoan).date()
            except ValueError:
                continue
            if not (first_this <= d <= last_this):
                continue
            uid, _ = seller_of(o)
            if uid:
                counts[uid] = counts.get(uid, 0) + 1

        total_pages = data.get("total_pages")
        if not orders or total_pages is None or page >= int(total_pages):
            break
        page += 1
    return counts


def fetch_month_stats(
    client: PancakeClient, shop_id: str, first: date, last: date, tz: ZoneInfo
) -> dict[str, dict]:
    """Theo nhân viên: số đơn chốt, số đơn hoàn, DS bán hàng (tổng tiền đơn chốt)."""
    start_ts = int(datetime.combine(first, time.min, tzinfo=tz).timestamp())
    end_ts = int(datetime.combine(last, time.max, tzinfo=tz).timestamp())
    stats: dict[str, dict] = {}
    for o in client.iter_orders(shop_id, start_ts, end_ts, page_size=config.ORDERS_PAGE_SIZE):
        status = o.get("status")
        uid, _ = seller_of(o)
        if not uid:
            continue
        s = stats.setdefault(uid, {"chot": 0, "hoan": 0, "ds": 0})
        if status in config.CLOSED_STATUSES:
            s["chot"] += 1
            s["ds"] += o.get("total_price") or 0
        elif status in config.RETURN_STATUSES:
            s["hoan"] += 1
    return stats


def parse_old_manual(old_values: list[list[str]] | None) -> dict[str, tuple]:
    """Đọc lại 2 cột nhập tay (Thưởng, % Thưởng) từ tab cũ: {tên: (thưởng, %)}."""
    result: dict[str, tuple] = {}
    if not old_values or len(old_values) < 3:
        return result
    header = old_values[1]
    try:
        i_name = header.index("Tên")
        i_th = header.index("Thưởng")
        i_pct = header.index("% Thưởng")
    except ValueError:
        return result
    for row in old_values[2:]:
        if len(row) <= max(i_th, i_pct):
            continue
        name = row[i_name].strip()
        if not name or name == "Tổng":
            continue
        digits = row[i_th].replace(".", "").replace(",", "").strip()
        thuong = int(digits) if digits.isdigit() else ""
        pct = row[i_pct].strip()
        if thuong or pct:
            result[name] = (thuong, pct)
    return result


def read_gr_bonus_totals(month: int, year: int) -> dict[str, int]:
    """Đọc cột 'Tổng tháng' của 2 tab Thưởng Sale/CSKH GR: {tên NV: tổng thưởng tháng}."""
    result: dict[str, int] = {}
    for label in ("Sale", "CSKH"):
        vals = google_sheet.read_table(f"Thưởng {label} GR T{month:02d}.{year}")
        if not vals or len(vals) < 3:
            continue
        header = vals[1]
        try:
            i_name = header.index("Họ và tên")
            i_tong = header.index("Tổng tháng")
        except ValueError:
            continue
        for row in vals[2:]:
            name = row[i_name].strip() if i_name < len(row) else ""
            if not name or name == "Tổng":
                continue
            digits = row[i_tong].replace(".", "").replace(",", "").strip() if i_tong < len(row) else ""
            if digits.isdigit():
                result[name] = int(digits)
    return result


def build_table(month: int, year: int, roster: list[tuple[str, dict]],
                stats: dict[str, dict], old_manual: dict[str, tuple],
                title_suffix: str = "",
                bonus_totals: dict[str, int] | None = None,
                hoan_truoc: dict[str, int] | None = None) -> list[list]:
    n = len(roster)
    first_data_row, last_data_row = 4, 3 + n     # dòng sheet (1-based)

    values: list[list] = [
        [f"THƯỞNG THÁNG SALE - CSKH THÁNG {month:02d}.{year}{title_suffix}"],
        HEADER,
        # Dòng Tổng (dòng 3) - công thức SUM để tự cập nhật khi sửa tay
        # Công thức dùng dấu ; (locale Việt Nam dùng , làm dấu thập phân)
        # Tỷ lệ hoàn = (hoàn tháng này + hoàn tháng trước) / đơn chốt
        ["", "Tổng", "",
         f"=SUM(D{first_data_row}:D{last_data_row})",
         f"=SUM(E{first_data_row}:E{last_data_row})",
         f"=SUM(F{first_data_row}:F{last_data_row})",
         f"=SUM(G{first_data_row}:G{last_data_row})",
         '=IF(D3=0;"";(E3+F3)/D3)',
         f"=SUM(I{first_data_row}:I{last_data_row})",
         "",
         f"=SUM(K{first_data_row}:K{last_data_row})"],
    ]
    bonus_totals = bonus_totals or {}
    hoan_truoc = hoan_truoc or {}
    for idx, (uid, info) in enumerate(roster, start=1):
        r = 3 + idx
        s = stats.get(uid, {"chot": 0, "hoan": 0, "ds": 0})
        # Thưởng = Tổng tháng từ tab Thưởng Sale/CSKH GR (khớp theo tên);
        # % Thưởng mặc định 100%, chỉnh tay trên sheet thì giữ nguyên giá trị đã chỉnh
        thuong = bonus_totals.get(info["name"], "") or ""
        _, pct = old_manual.get(info["name"], ("", ""))
        pct = pct or "100%"
        values.append([
            idx, info["name"], info["dept"],
            s["chot"], s["hoan"], hoan_truoc.get(uid, 0), s["ds"],
            f'=IF(D{r}=0;"";(E{r}+F{r})/D{r})',
            thuong, pct,
            f'=IF(I{r}="";"";I{r}*IF(J{r}="";1;J{r}))',
        ])
    return values


def parse_args() -> tuple[int, int]:
    parser = argparse.ArgumentParser(description='Bảng BC02 "Thưởng DS Sale- CSKH" theo tháng')
    parser.add_argument("thang", nargs="?", help="Tháng dạng YYYY-MM (mặc định: tháng hiện tại)")
    args = parser.parse_args()
    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()
    if args.thang:
        try:
            year, month = map(int, args.thang.split("-"))
            date(year, month, 1)
        except ValueError:
            raise SystemExit("Tháng không hợp lệ. Dùng dạng YYYY-MM, ví dụ: python bc02_thuong_ds.py 2026-07")
    else:
        year, month = today.year, today.month
    return year, month


def main() -> None:
    api_key = config.require_api_key()
    shop_id = config.PANCAKE_SHOP_ID
    if not shop_id:
        raise SystemExit("Chưa có PANCAKE_SHOP_ID trong .env.")

    year, month = parse_args()
    tz = ZoneInfo(config.TIMEZONE)
    today = datetime.now(tz).date()
    first = date(year, month, 1)
    if first > today:
        raise SystemExit(f"Tháng {month}/{year} chưa bắt đầu.")
    # Cùng logic "trễ 2 ngày" với thuong_thang.py: chỉ tính đến (hôm nay - 2)
    last = min(date(year, month, calendar.monthrange(year, month)[1]),
               today - timedelta(days=2))
    if last < first:
        raise SystemExit(f"Tháng {month}/{year}: chưa có ngày nào đủ 2 ngày chờ chốt.")
    tab_title = f"BC02 Thưởng DS Sale- CSKH T{month:02d}.{year}"

    # Tab đã chốt sổ thì không sửa nữa
    old_values = google_sheet.read_table(tab_title)
    if old_values and old_values[0] and "ĐÃ CHỐT SỔ" in old_values[0][0]:
        raise SystemExit(f"Tab '{tab_title}' ĐÃ CHỐT SỔ - không sửa nữa.")

    # Giữ lại số Thưởng / % Thưởng đã nhập tay trên tab cũ (nếu có)
    old_manual = parse_old_manual(old_values)

    client = PancakeClient(api_key)
    staff = load_staff(client, shop_id)
    roster = sorted(
        ((uid, info) for uid, info in staff.items()
         if any(kw in info["dept"].lower() for kw in GROUP_KEYWORDS)),
        key=lambda x: (x[1]["dept"], x[1]["name"]),
    )
    print(f"Bảng gồm {len(roster)} nhân viên các bộ phận Sale + CSKH")
    print(f"Lấy đơn hàng {first.strftime('%d/%m')} - {last.strftime('%d/%m/%Y')} từ Pancake POS ...")
    try:
        stats = fetch_month_stats(client, shop_id, first, last, tz)
    except PancakeError as e:
        raise SystemExit(f"Lỗi khi lấy đơn: {e}")

    bonus_totals = read_gr_bonus_totals(month, year)
    if not bonus_totals:
        print("CHÚ Ý: chưa đọc được cột Tổng tháng từ 2 tab Thưởng GR - cột Thưởng sẽ trống.")
    # Đơn hoàn tháng trước: TẠM THỜI để 0 theo yêu cầu (bật lại bằng dem_hoan_thang_truoc)
    hoan_truoc: dict[str, int] = {}
    values = build_table(month, year, roster, stats, old_manual,
                         bonus_totals=bonus_totals, hoan_truoc=hoan_truoc)
    url = google_sheet.write_bc02_table(tab_title, values)

    tong_chot = sum(stats.get(uid, {}).get("chot", 0) for uid, _ in roster)
    tong_hoan = sum(stats.get(uid, {}).get("hoan", 0) for uid, _ in roster)
    tong_ds = sum(stats.get(uid, {}).get("ds", 0) for uid, _ in roster)
    tong_ht = sum(hoan_truoc.get(uid, 0) for uid, _ in roster)
    print(f"\nTổng đơn chốt: {tong_chot} | Hoàn tháng này: {tong_hoan} | "
          f"Hoàn tháng trước: {tong_ht} | DS bán hàng: {vnd(tong_ds)}")
    if old_manual:
        print(f"Đã giữ nguyên Thưởng/% Thưởng nhập tay của {len(old_manual)} nhân viên.")
    print("Cột Thưởng = Tổng tháng từ 2 tab Thưởng GR; % Thưởng (nền vàng) nhập tay; Thực nhận tự tính.")
    print(f"\nĐã ghi tab '{tab_title}': {url}")


if __name__ == "__main__":
    main()
