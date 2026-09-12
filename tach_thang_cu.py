# -*- coding: utf-8 -*-
"""Tách các tab GỘP của 1 tháng ĐÃ CHỐT SỔ (lưu trữ) sang 2 trang tính riêng.

Dùng 1 lần cho mỗi tháng cũ khi chuyển sang mô hình mỗi bộ phận 1 trang tính:
    python tach_thang_cu.py 2026-08

Số liệu giữ NGUYÊN dạng tĩnh (không gọi API), kể cả dấu "ĐÃ CHỐT SỔ" trên tiêu đề:
    - Chuyển "Thưởng CSKH GR Txx" từ trang tính Sale sang trang tính CSKH
    - Tách "Doanh số NV Txx"            -> "Doanh số Sale Txx" (trang tính Sale)
                                          + "Doanh số CSKH Txx" (trang tính CSKH)
    - Tách "BC02 Thưởng DS Sale- CSKH Txx" -> "BC02 Thưởng DS Sale Txx" (Sale)
                                             + "BC02 Thưởng DS CSKH Txx" (CSKH)
    - Xóa tab gộp / tab CSKH khỏi trang tính Sale sau khi chuyển xong
"""
import argparse
import re
import sys
from datetime import date, datetime

import google_sheet
from thuong_thang import GROUPS, build_rules_block

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


# ----------------------------------------------------------------------
# Đọc lại giá trị đã hiển thị trên sheet
# ----------------------------------------------------------------------
def to_num_cell(s):
    """'1.234.567' -> 1234567; '' giữ nguyên; chữ giữ nguyên."""
    digits = (s or "").replace(".", "").replace(",", "").strip()
    if digits and digits.lstrip("-").isdigit():
        return int(digits)
    return s or ""


def to_pct(s):
    """'100%' -> 1.0; '20,0%' -> 0.2; giá trị khác giữ nguyên."""
    t = (s or "").strip()
    if not t.endswith("%"):
        return t
    try:
        return float(t[:-1].replace(".", "").replace(",", ".")) / 100
    except ValueError:
        return t


def lock_suffix(title: str) -> str:
    m = re.search(r"\s*\(ĐÃ CHỐT SỔ[^)]*\)", title or "")
    return m.group(0) if m else ""


def pad(row: list, n: int) -> list:
    return list(row) + [""] * (n - len(row))


def sunday_cols_of(header: list) -> list[int]:
    cols = []
    for i, h in enumerate(header):
        try:
            if datetime.strptime((h or "").strip(), "%d/%m/%Y").weekday() == 6:
                cols.append(i)
        except ValueError:
            continue
    return cols


def table_rows(values: list[list]) -> tuple[list, list[list], list | None]:
    """(header, các dòng nhân viên, dòng Tổng) của tab ma trận thưởng/doanh số."""
    header = values[1]
    n = len(header)
    for i in range(2, len(values)):
        row = pad(values[i], n)
        if row[1].strip() == "Tổng":
            return header, [pad(r, n) for r in values[2:i]], row
    return header, [pad(r, n) for r in values[2:]], None


# ----------------------------------------------------------------------
# Dựng bảng mới từ số liệu cũ
# ----------------------------------------------------------------------
def convert_matrix(values: list[list]) -> list[list]:
    """Tab ma trận giữ nguyên danh sách (dùng khi CHUYỂN cả tab sang trang tính khác)."""
    header, emps, total = table_rows(values)
    out = [[values[0][0] if values[0] else ""], list(header)]
    for r in emps:
        out.append([to_num_cell(r[0]), r[1], r[2]] + [to_num_cell(v) for v in r[3:]])
    if total:
        out.append(["", "Tổng", ""] + [to_num_cell(v) for v in total[3:]])
    return out


def build_group_matrix(title: str, header: list, emp_rows: list[list], kw: str) -> list[list]:
    """Tab ma trận chỉ gồm nhân viên có Bộ phận chứa `kw`; dòng Tổng cộng tĩnh lại."""
    rows = [r for r in emp_rows if kw in (r[2] or "").lower()]
    n = len(header)
    out = [[title], list(header)]
    sums = [0] * (n - 3)
    for idx, r in enumerate(rows, start=1):
        cells = []
        for j, v in enumerate(r[3:]):
            x = to_num_cell(v)
            cells.append(x)
            if isinstance(x, int):
                sums[j] += x
        out.append([idx, r[1], r[2]] + cells)
    out.append(["", "Tổng", ""] + sums)
    return out


def build_group_bc02(label: str, month: int, year: int,
                     values: list[list], kw: str) -> list[list]:
    """Tab BC02 chỉ gồm nhân viên có Bộ phận chứa `kw`; dòng Tổng (dòng 3) cộng tĩnh."""
    header = values[1]
    n = len(header)
    num_cols = (3, 4, 5, 6, 8, 10)   # Đơn chốt, Hoàn tháng này/trước, DS, Thưởng, Thực nhận
    pct_cols = (7, 9)                # Tỷ lệ hoàn, % Thưởng
    emps = [pad(r, n) for r in values[3:]
            if len(r) > 2 and kw in (r[2] or "").lower()]
    suffix = lock_suffix(values[0][0] if values[0] else "")

    conv = []
    for idx, r in enumerate(emps, start=1):
        row = [idx] + r[1:]
        for j in num_cols:
            row[j] = to_num_cell(r[j])
        for j in pct_cols:
            row[j] = to_pct(r[j])
        conv.append(row)

    def s(j, kinds=(int,)):
        return sum(r[j] for r in conv if isinstance(r[j], kinds))

    d, e, f, g = s(3), s(4), s(5), s(6)
    ty_le = (e + f) / d if d else ""
    total = ["", "Tổng", "", d, e, f, g, ty_le, s(8), "", s(10, (int, float))]
    return [[f"THƯỞNG THÁNG {label.upper()} THÁNG {month:02d}.{year}{suffix}"],
            list(header), total] + conv


# ----------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Tách tab gộp của 1 tháng đã chốt sổ sang 2 trang tính riêng (Sale/CSKH)")
    parser.add_argument("thang", help="Tháng dạng YYYY-MM, ví dụ: 2026-08")
    args = parser.parse_args()
    try:
        year, month = map(int, args.thang.split("-"))
        date(year, month, 1)
    except ValueError:
        raise SystemExit("Tháng không hợp lệ. Dùng dạng YYYY-MM, ví dụ: python tach_thang_cu.py 2026-08")
    t = f"T{month:02d}.{year}"

    # 1) CHUYỂN "Thưởng CSKH GR" sang trang tính CSKH
    tab = f"Thưởng CSKH GR {t}"
    src = google_sheet.read_table(tab)
    if src:
        out = convert_matrix(src)
        google_sheet.write_table(tab, out, sunday_cols=sunday_cols_of(out[1]), nhom="cskh")
        google_sheet.delete_tab(tab)
        print(f"[OK] Đã chuyển '{tab}' sang trang tính CSKH ({len(out) - 3} NV).")
    else:
        print(f"(Bỏ qua: không thấy '{tab}' trên trang tính Sale.)")

    # 2) TÁCH "Doanh số NV" thành 2 tab
    tab_gop = f"Doanh số NV {t}"
    src = google_sheet.read_table(tab_gop)
    if src:
        header, emps, _ = table_rows(src)
        suffix = lock_suffix(src[0][0] if src[0] else "")
        sunday = sunday_cols_of(header)
        for kw, label in GROUPS:
            title = (f"Doanh số ngày Bộ phận {label} Tháng {month}.{year} "
                     f"(căn cứ tính thưởng GR){suffix}")
            out = build_group_matrix(title, header, emps, kw)
            google_sheet.write_table(f"Doanh số {label} {t}", out, sunday_cols=sunday,
                                     block_rows=build_rules_block(only=kw), nhom=kw)
            print(f"[OK] 'Doanh số {label} {t}' ({len(out) - 3} NV).")
        google_sheet.delete_tab(tab_gop)
        print(f"[OK] Đã xóa tab gộp '{tab_gop}'.")
    else:
        print(f"(Bỏ qua: không thấy '{tab_gop}' trên trang tính Sale.)")

    # 3) TÁCH BC02 thành 2 tab
    tab_gop = f"BC02 Thưởng DS Sale- CSKH {t}"
    src = google_sheet.read_table(tab_gop)
    if src:
        for kw, label in GROUPS:
            out = build_group_bc02(label, month, year, src, kw)
            google_sheet.write_bc02_table(f"BC02 Thưởng DS {label} {t}", out, nhom=kw)
            print(f"[OK] 'BC02 Thưởng DS {label} {t}' ({len(out) - 3} NV).")
        google_sheet.delete_tab(tab_gop)
        print(f"[OK] Đã xóa tab gộp '{tab_gop}'.")
    else:
        print(f"(Bỏ qua: không thấy '{tab_gop}' trên trang tính Sale.)")

    # Tháng mới nhất xếp bên TRÁI ở cả 2 trang tính
    for kw, _label in GROUPS:
        google_sheet.sap_xep_tab(kw)

    print("\nXong. Số liệu tháng cũ giữ nguyên (dạng tĩnh), mỗi bộ phận 1 trang tính riêng.")


if __name__ == "__main__":
    main()
