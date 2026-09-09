# -*- coding: utf-8 -*-
"""Cập nhật CẢ 3 BẢNG thưởng trên Google Sheet - logic "chốt ngày, chốt sổ".

Chạy:
    python thuong_thang.py            # cập nhật tháng hiện tại + tự chốt sổ tháng trước
    python thuong_thang.py 2026-08    # tháng đã qua: CHỐT SỔ (tính lại cả tháng rồi khóa)

Nguyên tắc:
    1. TRỄ 2 NGÀY: hôm nay 29 thì bảng chỉ hiển thị đến ngày 27 - hai ngày cuối
       trạng thái đơn còn thay đổi nhiều nên chưa đưa vào.
    2. MỖI LẦN CHẠY GỌI API 7 NGÀY GẦN NHẤT (đến hôm nay - 2) và GHI ĐÈ kho
       api_data - kho luôn tươi trong cửa sổ 7 ngày.
    3. NGÀY ĐÃ LÊN BẢNG = ĐÃ CHỐT (bảng Thưởng GR / Doanh số NV / BC02): chỉ NỐI
       THÊM cột của ngày mới (hôm nay - 2); số các ngày cũ giữ nguyên như trên
       sheet, KỂ CẢ khi kho api_data quá khứ đã được ghi đè mới hơn.
    4. Tab DOANH SỐ PAGE: luôn DỰNG LẠI từ kho api_data -> 7 ngày gần nhất của nó
       phản ánh trạng thái đơn mới nhất.
    5. CHỐT SỔ CUỐI THÁNG: sang tháng mới (từ mùng 2), script tự gọi lại API MỘT LẦN
       trọn tháng trước để sửa thưởng lần cuối (bắt đơn hoàn/hủy muộn), đóng dấu
       "ĐÃ CHỐT SỔ" lên tiêu đề các tab - từ đó không đụng đến tháng đó nữa
       (kể cả kho api_data của tháng đó).

HAI TRANG TÍNH riêng, mỗi bộ phận 1 file, KHÔNG lẫn thông tin của nhau
(.env: GOOGLE_SHEET_ID = Sale, GOOGLE_SHEET_ID_CSKH = CSKH):
- Trang tính Sale: "Thưởng Sale GR", "Doanh số Sale" (ma trận thưởng/doanh số ngày)
  và "BC02 Thưởng DS Sale" (thưởng doanh số tháng, cột % nhập tay giữ nguyên).
- Trang tính CSKH: "Thưởng CSKH GR", "Doanh số CSKH", "BC02 Thưởng DS CSKH".
Lần chạy đầu sau khi tách: tự CHUYỂN dữ liệu CSKH của tháng chưa khóa từ trang tính cũ
sang trang tính mới, tách các tab gộp cũ ("Doanh số NV", "BC02 ... Sale- CSKH")
thành tab riêng từng bộ phận (kế thừa số đã chốt) rồi xóa tab gộp đi.
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
from doanh_thu import (clean_name, load_staff, matched_departments,
                       roster_sort_key, seller_of, vnd)
from pancake_client import PancakeClient, PancakeError

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

GROUPS = [("sale", "Sale"), ("cskh", "CSKH")]
LOCK_MARK = "ĐÃ CHỐT SỔ"
SETTLE_DELAY_DAYS = 2      # số ngày trễ trước khi một ngày được đưa vào bảng
FETCH_WINDOW_DAYS = 7      # mỗi lần chạy gọi API ghi đè kho bấy nhiêu ngày gần nhất


def bonus_for(doanh_so: int, day: date, tiers_cfg: dict) -> int:
    """Tiền thưởng theo mốc doanh số ngày; Chủ nhật dùng bảng mốc riêng."""
    tiers = tiers_cfg["sunday"] if day.weekday() == 6 else tiers_cfg["weekday"]
    for nguong, thuong in sorted(tiers, reverse=True):
        if doanh_so >= nguong:
            return thuong
    return 0


def dang_ky_chu_nhat(roster: list[tuple[str, dict]], keyword: str,
                     days: list[date]) -> tuple[set[tuple[str, date]], list[str]]:
    """Đối chiếu "Lịch trực ngày nghỉ": trả về (set (uid, Chủ nhật ĐÃ đăng ký làm),
    danh sách cảnh báo). Chủ nhật không có trong set -> thưởng ngày đó = 0."""
    import lich_truc

    sundays = [d for d in days if d.weekday() == 6]
    if not sundays:
        return set(), []
    lich = lich_truc.doc_dang_ky().get(keyword) or {"cols": set(), "reg": {}}
    reg, cols = lich["reg"], lich["cols"]
    ok: set[tuple[str, date]] = set()
    canh_bao: list[str] = []
    for uid, info in roster:
        ten = lich_truc.chuan_hoa_ten(info["name"])
        # Tên trên Pancake thường = tên lịch trực + hậu tố bộ phận -> khớp theo tiền tố
        khop = [ln for ln in reg if ten.startswith(ln) or ln.startswith(ten)]
        if not khop:
            canh_bao.append(f"'{info['name']}' không có trong lịch trực -> thưởng CN = 0")
            continue
        ln = max(khop, key=len)
        for d in sundays:
            if (d.day, d.month) in reg[ln]:
                ok.add((uid, d))
    for d in sundays:
        if (d.day, d.month) not in cols:
            canh_bao.append(f"Lịch trực CHƯA có cột ngày {d.strftime('%d/%m')} "
                            f"-> cả bộ phận 0 thưởng CN này")
    return ok, canh_bao


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

    # CHỈ lưu các ngày trong khoảng yêu cầu: API có thể trả kèm vài đơn "rìa" của
    # ngày lân cận - nếu lưu cả sẽ GHI ĐÈ file ngày đó chỉ với mấy đơn rìa (mất dữ liệu)
    luu_tru.luu_don_theo_ngay({d: lst for d, lst in orders_by_day.items()
                               if start_day <= d <= end_day})
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
                # giữ cả số 0 (CN không đăng ký làm bị ghi 0 rõ ràng); ô rỗng thì bỏ qua
                if digits.isdigit():
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
    dang_ky_cn: set[tuple[str, date]] | None = None,
) -> list[list]:
    """Ma trận thưởng ngày: ngày đã chốt lấy số cũ trên sheet, ngày mới tính từ data.

    dang_ky_cn: set (uid, ngày CN đã đăng ký làm theo Lịch trực) - Chủ nhật không
    nằm trong set thì thưởng = 0; None = không áp quy tắc lịch trực."""
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
            # Ngày đã chốt: lấy số cũ trên sheet; KHÔNG THẤY số cũ (NV đổi tên trên
            # Pancake, dòng mới...) thì tính lại từ data thay vì bỏ trắng
            cu = old_bonus.get((info["name"], d)) if d in settled_days else None
            if cu is not None:
                cells.append(cu)
                continue
            ds_all = data.get(uid, {}).get(d, {}).get("ds_all", 0)
            b = bonus_for(ds_all, d, tiers_cfg)
            if (d.weekday() == 6 and dang_ky_cn is not None
                    and (uid, d) not in dang_ky_cn):
                # CN không đăng ký làm trên Lịch trực: lẽ ra có thưởng -> ghi 0 rõ ràng,
                # không đạt mốc nào -> để rỗng như bình thường
                cells.append(0 if b else "")
            else:
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


def build_ds_matrix(
    month_label: str, days: list[date], roster: list[tuple[str, dict]],
    data: dict, settled_days: set[date], old_values: dict, group_label: str,
    title_suffix: str = "",
) -> list[list]:
    """Ma trận DOANH SỐ ngày theo nhân viên (căn cứ tính thưởng) - để đối chiếu."""
    n_day_cols = len(days)
    first_day_col = 4
    header = ["STT", "Họ và tên", "Bộ phận"] + [d.strftime("%d/%m/%Y") for d in days] + ["Tổng tháng"]
    values: list[list] = [
        [f"Doanh số ngày Bộ phận {group_label} Tháng {month_label} (căn cứ tính thưởng GR){title_suffix}"],
        header,
    ]
    for idx, (uid, info) in enumerate(roster, start=1):
        row_num = len(values) + 1
        cells = []
        for d in days:
            # Ngày đã chốt lấy số cũ; không thấy số cũ thì tính lại từ data (xem build_matrix)
            cu = old_values.get((info["name"], d)) if d in settled_days else None
            if cu is not None:
                cells.append(cu)
            else:
                ds = data.get(uid, {}).get(d, {}).get("ds_all", 0)
                cells.append(ds if ds else "")
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


def build_page_matrix(month_label: str, days: list[date], staff: dict[str, dict],
                      keyword: str, group_label: str, title_suffix: str = "") -> list[list]:
    """Ma trận DOANH SỐ ngày theo PAGE NGUỒN:
      - DOANH SỐ = tổng tiền ĐƠN CHỐT + ĐƠN HOÀN (Đang hoàn/Đã hoàn)
      - gộp theo NGUỒN ĐƠN HÀNG (`account` = page quảng cáo dẫn đơn về),
        không phải page hội thoại nơi tạo đơn
      - tính theo NGÀY TẠO đơn (cùng mốc với ma trận Doanh số nhân viên,
        nên tổng ngày của 2 bảng đối chiếu được với nhau)
    CHỈ tính đơn do nhân viên bộ phận chứa `keyword` phụ trách.

    Dựng lại toàn bộ từ đơn thô đã lưu trong api_data/ (không gọi thêm API);
    các page xếp theo tổng doanh số tháng giảm dần."""
    import luu_tru

    dept_kw = keyword.lower()
    ds: dict[str, dict[date, int]] = {}      # nguồn -> {ngày tạo: doanh số}
    names: dict[str, str] = {}
    plats: dict[str, str] = {}
    for day, orders in luu_tru.doc_don_theo_ngay(days).items():
        for o in orders:
            status = o.get("status")
            if (status not in config.CLOSED_STATUSES
                    and status not in config.RETURN_STATUSES):
                continue                     # doanh số = đơn chốt + đơn hoàn
            uid, _ = seller_of(o)
            info = staff.get(uid)
            if not info or dept_kw not in info["dept"].lower():
                continue                     # đơn của bộ phận khác -> không tính
            acc = str(o.get("account") or "")
            name = clean_name(o.get("account_name") or "") or "(Không có nguồn)"
            key = acc or name
            names[key] = name
            plats[key] = ("" if not acc
                          else "Zalo" if acc.startswith("pzl_") else "Facebook")
            cell = ds.setdefault(key, {})
            cell[day] = cell.get(day, 0) + (o.get("total_price") or 0)

    # Page khác id nhưng trùng tên -> thêm đuôi id để không lẫn nhau
    dem_ten: dict[str, int] = {}
    for n in names.values():
        dem_ten[n] = dem_ten.get(n, 0) + 1

    def ten_hien_thi(key: str) -> str:
        n = names[key]
        return f"{n} ({key[-4:]})" if dem_ten[n] > 1 and key != n else n

    tong = {k: sum(v.values()) for k, v in ds.items()}
    keys = sorted(ds, key=lambda k: (-tong[k], names[k]))

    n_day_cols = len(days)
    first_day_col = 4
    header = ["STT", "Page nguồn", "Nền tảng"] + [d.strftime("%d/%m/%Y") for d in days] + ["Tổng tháng"]
    values: list[list] = [
        [f"Doanh số ngày theo PAGE NGUỒN - Bộ phận {group_label} Tháng {month_label} "
         f"(doanh số = đơn chốt + đơn hoàn; gộp theo nguồn đơn, tính theo ngày tạo đơn)"
         f"{title_suffix}"],
        header,
    ]
    for idx, k in enumerate(keys, start=1):
        row_num = len(values) + 1
        cells = [ds[k].get(d) or "" for d in days]
        first_cell = f"{google_sheet._col_letter(first_day_col)}{row_num}"
        last_cell = f"{google_sheet._col_letter(first_day_col + n_day_cols - 1)}{row_num}"
        values.append([idx, ten_hien_thi(k), plats[k]] + cells
                      + [f"=SUM({first_cell}:{last_cell})"])

    first_data_row, last_data_row = 3, len(values)
    total_row: list = ["", "Tổng", ""]
    for i in range(n_day_cols + 1):
        col = google_sheet._col_letter(first_day_col + i)
        total_row.append(f"=SUM({col}{first_data_row}:{col}{last_data_row})")
    values.append(total_row)
    return values


def build_rules_block(only: str | None = None) -> list[list]:
    """Bảng quy tắc thưởng (ngày thường + Chủ nhật) - tự sinh từ config.

    only: chỉ lấy quy tắc của 1 nhóm ("sale"/"cskh"); None = cả Sale lẫn CSKH."""
    tiers = config.BONUS_TIERS_BY_GROUP
    groups = [(kw, label) for kw, label in GROUPS if only is None or kw == only]
    thresholds = sorted({nguong for kw, _ in groups
                         for arr in tiers[kw].values() for nguong, _ in arr})

    def muc(group: str, kind: str, nguong: int):
        for n, thuong in tiers[group][kind]:
            if n == nguong:
                return thuong
        return ""

    header = ["Doanh số ngày từ"]
    for _, label in groups:
        header += [f"{label} - Ngày thường", f"{label} - Chủ nhật"]
    rows: list[list] = [
        ["QUY TẮC TÍNH THƯỞNG THEO DOANH SỐ NGÀY (đạt mốc >= nào cao nhất thì hưởng mốc đó)"],
        header,
    ]
    for t in thresholds:
        row: list = [t]
        for kw, _ in groups:
            row += [muc(kw, "weekday", t), muc(kw, "sunday", t)]
        rows.append(row)
    rows.append(["Chủ nhật KHÔNG 'Đăng kí làm' trên Lịch trực ngày nghỉ "
                 "-> thưởng Chủ nhật = 0 (ô ghi số 0)"])
    return rows


def group_roster(staff: dict, keyword: str) -> list[tuple[str, dict]]:
    dept_filter = matched_departments(staff, keyword)
    return sorted(
        ((uid, info) for uid, info in staff.items() if info["dept"] in dept_filter),
        key=roster_sort_key,
    )


def run_month(client: PancakeClient, shop_id: str, staff: dict, tz: ZoneInfo,
              today: date, year: int, month: int, finalize: bool,
              only_if_exists: bool = False) -> None:
    """Cập nhật (hoặc chốt sổ) cả 3 tab của 1 tháng."""
    first = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])

    tabs = {kw: f"Thưởng {label} GR T{month:02d}.{year}" for kw, label in GROUPS}
    bc02_tabs = {kw: f"BC02 Thưởng DS {label} T{month:02d}.{year}" for kw, label in GROUPS}
    ds_tabs = {kw: f"Doanh số {label} T{month:02d}.{year}" for kw, label in GROUPS}
    # Tab doanh số theo page: mỗi trang tính 1 tab, lọc theo bộ phận tương ứng
    page_tabs = {kw: f"Doanh số {label} Page T{month:02d}.{year}" for kw, label in GROUPS}
    # Tab gộp Sale+CSKH của bản cũ (đều nằm ở trang tính Sale)
    bc02_tab_gop_cu = f"BC02 Thưởng DS Sale- CSKH T{month:02d}.{year}"
    ds_tab_gop_cu = f"Doanh số NV T{month:02d}.{year}"

    # Mỗi bộ phận 1 TRANG TÍNH riêng: tab của nhóm nào đọc/ghi ở trang tính nhóm đó.
    # Bản cũ để mọi tab ở trang tính Sale -> tab CSKH chưa có ở trang tính mới thì
    # đọc từ trang tính Sale để kế thừa số đã chốt, ghi xong sẽ dọn tab cũ đi.
    old: dict[str, list | None] = {}
    tab_cskh_o_sheet_cu = False
    for kw, _label in GROUPS:
        vals = google_sheet.read_table(tabs[kw], nhom=kw)
        if kw == "cskh" and vals is None:
            vals = google_sheet.read_table(tabs[kw])
            tab_cskh_o_sheet_cu = vals is not None
        old[kw] = vals
    # Tab riêng chưa có (lần đầu chạy bản tách) -> kế thừa số cũ từ tab gộp
    old_bc02_gop = google_sheet.read_table(bc02_tab_gop_cu)
    old_bc02_raw = {kw: google_sheet.read_table(t, nhom=kw) for kw, t in bc02_tabs.items()}
    old_bc02 = {kw: v if v else old_bc02_gop for kw, v in old_bc02_raw.items()}
    old_ds_gop = google_sheet.read_table(ds_tab_gop_cu)
    old_ds_raw = {kw: google_sheet.read_table(t, nhom=kw) for kw, t in ds_tabs.items()}
    old_ds = {kw: v if v else old_ds_gop for kw, v in old_ds_raw.items()}
    if is_locked(old["sale"]) or is_locked(old_bc02["sale"]):
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
        # Ngày "đã chốt" = ngày đã có cột trên CẢ tab thưởng lẫn 2 tab doanh số
        # (tab nào thiếu ngày thì ngày đó được lấy lại dữ liệu để bổ sung)
        settled_days = parse_header_days(old["sale"])
        for vals in old_ds.values():
            settled_days &= parse_header_days(vals)
        settled_days = {d for d in settled_days if d <= cutoff}
        # Nếu tab cũ có cột ngày vượt cutoff (dữ liệu của logic cũ) -> làm lại từ đầu
        if any(d > cutoff for d in parse_header_days(old["sale"])):
            settled_days = set()
        title_suffix = ""

    days = [first + timedelta(days=i) for i in range((cutoff - first).days + 1)]
    new_days = [d for d in days if d not in settled_days]

    if finalize:
        data = fetch_range_data(client, shop_id, first, cutoff, tz)
    else:
        # LUÔN gọi API 7 NGÀY GẦN NHẤT (giới hạn trong tháng) ghi đè kho api_data:
        # - bảng Thưởng/Doanh số NV/BC02 chỉ NỐI CỘT ngày mới, số cũ giữ nguyên
        # - tab Doanh số Page dựng lại từ kho -> luôn theo trạng thái đơn mới nhất
        fetch_start = max(first, cutoff - timedelta(days=FETCH_WINDOW_DAYS - 1))
        if new_days:
            fetch_start = min(fetch_start, new_days[0])   # chạy bù ngày cũ chưa lên bảng
        thong_bao_moi = (", ngày mới lên bảng: "
                         + ", ".join(d.strftime("%d/%m") for d in new_days)
                         if new_days else ", không có ngày mới lên bảng")
        print(f"Tháng {month:02d}.{year}: hiển thị đến {cutoff.strftime('%d/%m')}, "
              f"lấy API {fetch_start.strftime('%d/%m')} - {cutoff.strftime('%d/%m')} "
              f"ghi đè kho{thong_bao_moi} ...")
        data = fetch_range_data(client, shop_id, fetch_start, cutoff, tz)

    # --- 2 bảng ma trận thưởng ngày (mỗi nhóm ghi vào trang tính riêng) ---
    # Lịch trực Chủ nhật: đọc 1 lần; LỖI -> bỏ qua quy tắc, KHÔNG chặn cập nhật
    lich_truc_loi = None
    try:
        import lich_truc
        lich_truc.doc_dang_ky()
    except Exception as e:                        # noqa: BLE001 - mọi lỗi đều chỉ cảnh báo
        lich_truc_loi = f"{type(e).__name__}: {e}"

    month_label = f"{month}.{year}"
    da_ghi: set[str] = set()
    so_nguoi_tru_cn = 0
    sunday_idx = [3 + i for i, d in enumerate(days) if d.weekday() == 6]
    for keyword, label in GROUPS:
        roster = group_roster(staff, keyword)
        if not roster:
            continue
        if lich_truc_loi is None:
            dang_ky, canh_bao = dang_ky_chu_nhat(roster, keyword, days)
        else:
            dang_ky, canh_bao = None, []
        for cb in canh_bao:
            print(f"  [Lịch trực {label}] {cb}")
        values = build_matrix(month_label, days, roster, data,
                              config.BONUS_TIERS_BY_GROUP[keyword], label,
                              settled_days, parse_old_bonus(old[keyword]), title_suffix,
                              dang_ky_cn=dang_ky)
        # Đếm người có ô CN = 0 (đạt mốc thưởng nhưng không đăng ký làm)
        so_nguoi_tru_cn += sum(
            1 for row in values[2:-1]
            if any(i < len(row) and row[i] == 0 for i in sunday_idx))
        end_col = google_sheet._col_letter(3 + len(days) + 1)
        sunday_cols = [3 + i for i, d in enumerate(days) if d.weekday() == 6]
        google_sheet.write_table(tabs[keyword], values,
                                 money_range=f"D3:{end_col}{len(values)}",
                                 sunday_cols=sunday_cols, nhom=keyword)
        da_ghi.add(keyword)
        total = sum(sum(v for v in row[3:-1] if isinstance(v, int)) for row in values[2:-1])
        print(f"  [OK] {tabs[keyword]}: {len(roster)} NV, tổng thưởng {vnd(total)}")
    if lich_truc_loi is not None:
        print(f"  [LICH TRUC][LOI] Không đọc được Lịch trực ngày nghỉ - BỎ QUA trừ thưởng "
              f"Chủ nhật lần chạy này ({lich_truc_loi[:100]})")
    elif sunday_idx:
        print(f"  [LICH TRUC][OK] Lịch trực kết nối OK - tháng {month:02d}.{year} có "
              f"{so_nguoi_tru_cn} người bị trừ thưởng Chủ nhật về 0 "
              f"(ô ghi 0 = có thưởng mà bị trừ; ô rỗng = bình thường)")

    # --- 2 tab DOANH SỐ theo bộ phận: ô = doanh số ngày (đối chiếu), mỗi nhóm 1 trang tính ---
    end_col = google_sheet._col_letter(3 + len(days) + 1)
    sunday_cols = [3 + i for i, d in enumerate(days) if d.weekday() == 6]
    da_ghi_ds: set[str] = set()
    for keyword, label in GROUPS:
        roster = group_roster(staff, keyword)
        if not roster:
            continue
        ds_values = build_ds_matrix(month_label, days, roster, data, settled_days,
                                    parse_old_bonus(old_ds[keyword]), label, title_suffix)
        google_sheet.write_table(ds_tabs[keyword], ds_values,
                                 money_range=f"D3:{end_col}{len(ds_values)}",
                                 sunday_cols=sunday_cols,
                                 rules_block=build_rules_block(only=keyword),
                                 nhom=keyword)
        da_ghi_ds.add(keyword)
        tong_ds_ngay = sum(
            sum(v for v in row[3:-1] if isinstance(v, int)) for row in ds_values[2:-1]
        )
        print(f"  [OK] {ds_tabs[keyword]}: {len(roster)} NV, tổng doanh số {vnd(tong_ds_ngay)}")

    # --- Tab DOANH SỐ THEO PAGE (mỗi bộ phận 1 tab ở trang tính riêng, lọc đơn theo
    #     nhân viên bộ phận đó): dựng lại từ đơn thô đã lưu ---
    for keyword, label in GROUPS:
        page_values = build_page_matrix(month_label, days, staff, keyword, label, title_suffix)
        google_sheet.write_table(page_tabs[keyword], page_values,
                                 money_range=f"D3:{end_col}{len(page_values)}",
                                 sunday_cols=sunday_cols, nhom=keyword)
        print(f"  [OK] {page_tabs[keyword]}: {len(page_values) - 3} page")

    # --- Dọn trang tính Sale (cũ) sau khi đã chuyển/tách xong ---
    if tab_cskh_o_sheet_cu and "cskh" in da_ghi and google_sheet.delete_tab(tabs["cskh"]):
        print(f"  Đã chuyển tab '{tabs['cskh']}' sang trang tính CSKH, xóa khỏi trang tính Sale.")
    if (old_ds_gop is not None and da_ghi_ds >= {kw for kw, _ in GROUPS}
            and google_sheet.delete_tab(ds_tab_gop_cu)):
        print(f"  Đã xóa tab gộp cũ '{ds_tab_gop_cu}' (Sale / CSKH giờ mỗi nhóm 1 trang tính riêng).")

    # --- BC02 (mỗi bộ phận 1 tab ở trang tính riêng): cộng dồn ngày mới vào số cũ,
    #     chốt sổ thì tính lại từ data cả tháng ---
    # Cột Thưởng của BC02 = Tổng tháng trên 2 tab GR vừa ghi (khớp theo tên NV)
    bonus_totals = read_gr_bonus_totals(month, year)
    # Đơn hoàn tháng trước: TẠM THỜI để 0 theo yêu cầu
    # (bật lại: hoan_truoc = dem_hoan_thang_truoc(client, shop_id, year, month, tz))
    hoan_truoc: dict[str, int] = {}
    da_ghi_bc02: set[str] = set()
    for keyword, label in GROUPS:
        roster = group_roster(staff, keyword)
        if not roster:
            continue
        old_stats = ({} if (finalize or not settled_days)
                     else parse_old_bc02_stats(old_bc02[keyword]))
        stats: dict[str, dict] = {}
        for uid, info in roster:
            base = old_stats.get(info["name"], {"chot": 0, "hoan": 0, "ds": 0})
            s = {"chot": base["chot"], "hoan": base["hoan"], "ds": base["ds"]}
            for d in new_days:
                cell = data.get(uid, {}).get(d)
                if cell:
                    s["chot"] += cell["chot"]
                    s["hoan"] += cell["hoan"]
                    s["ds"] += cell["ds_all"]   # DS bán hàng = DOANH SỐ (chốt + hoàn)
            stats[uid] = s
        values = bc02_build_table(month, year, roster, stats,
                                  bc02_parse_old_manual(old_bc02[keyword]), title_suffix,
                                  bonus_totals=bonus_totals, hoan_truoc=hoan_truoc,
                                  group_label=label)
        google_sheet.write_bc02_table(bc02_tabs[keyword], values, nhom=keyword)
        da_ghi_bc02.add(keyword)
        tong_chot = sum(s["chot"] for s in stats.values())
        tong_hoan = sum(s["hoan"] for s in stats.values())
        tong_ht = sum(hoan_truoc.get(uid, 0) for uid, _ in roster)
        tong_ds = sum(s["ds"] for s in stats.values())
        print(f"  [OK] {bc02_tabs[keyword]}: {tong_chot} đơn chốt, {tong_hoan} hoàn tháng này, "
              f"{tong_ht} hoàn tháng trước, DS {vnd(tong_ds)}")
    if (old_bc02_gop is not None and da_ghi_bc02 >= {kw for kw, _ in GROUPS}
            and google_sheet.delete_tab(bc02_tab_gop_cu)):
        print(f"  Đã xóa tab BC02 gộp cũ '{bc02_tab_gop_cu}' (đã tách mỗi bộ phận 1 tab riêng).")
    if finalize:
        print(f"  Đã đóng dấu '{LOCK_MARK}' - các tab tháng {month:02d}.{year} bị khóa vĩnh viễn.")


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

    # Tháng mới nhất xếp bên TRÁI ở cả 2 trang tính
    google_sheet.sap_xep_tab("sale")
    if config.GOOGLE_SHEET_ID_CSKH:
        google_sheet.sap_xep_tab("cskh")

    print(f"\nXong."
          f"\n  Trang tính Sale: https://docs.google.com/spreadsheets/d/{config.GOOGLE_SHEET_ID}")
    if config.GOOGLE_SHEET_ID_CSKH:
        print(f"  Trang tính CSKH: https://docs.google.com/spreadsheets/d/{config.GOOGLE_SHEET_ID_CSKH}")


if __name__ == "__main__":
    main()
