# -*- coding: utf-8 -*-
"""Đọc "Lịch trực ngày nghỉ - Greenvita" (Google Sheet riêng, GOOGLE_SHEET_ID_LICH_TRUC).

Tab chọn theo CHỮ trong tên tab: có "sale" -> bộ phận Sale, có "cskh" -> bộ phận
CSKH (vd "Bộ phận SALE", "Bộ phận CSKH"; nhiều tab cùng chữ thì gộp lại). Tab phải
có dòng header STT | Họ Tên | Cơ sở | các cột ngày CHỦ NHẬT dạng d/m; ô ghi
"Đăng kí làm" = có đăng ký trực hôm đó. Tab không có cột "Họ Tên" (lịch mẫu cũ) bị bỏ qua.

Quy tắc áp ở thuong_thang.py: Chủ nhật KHÔNG đăng ký làm -> thưởng ngày đó = 0.
"""
from __future__ import annotations

import re
import unicodedata

import config
import google_sheet

# nhóm -> chữ phải có trong tên tab lịch trực (so sau khi chuẩn hóa, không phân biệt hoa/thường)
TU_KHOA_TAB = {"sale": "sale", "cskh": "cskh"}
_NGAY_RE = re.compile(r"^\s*(\d{1,2})\s*/\s*(\d{1,2})\s*$")


def duong_dan() -> str:
    """Link mở trang tính Lịch trực (dùng cho nút trên app và ô nguồn trên sheet)."""
    return (f"https://docs.google.com/spreadsheets/d/{config.GOOGLE_SHEET_ID_LICH_TRUC}"
            if config.GOOGLE_SHEET_ID_LICH_TRUC else "")


def khop_ten(reg: dict, ten_pancake: str) -> str | None:
    """Tên tương ứng trên lịch trực của 1 nhân viên Pancake (None nếu không có).

    Tên trên Pancake thường = tên lịch trực + hậu tố bộ phận -> khớp theo tiền tố."""
    ten = chuan_hoa_ten(ten_pancake)
    khop = [ln for ln in reg if ten and (ten.startswith(ln) or ln.startswith(ten))]
    return max(khop, key=len) if khop else None


def chuan_hoa_ten(s: str) -> str:
    """Chuẩn hóa để so khớp tên: thường hóa, BỎ DẤU, gọn khoảng trắng
    (lịch trực viết tay có thể thiếu dấu so với tên trên Pancake)."""
    s = (s or "").replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.lower().split())


_cache: dict | None = None


def doc_dang_ky() -> dict[str, dict]:
    """{nhóm: {"tabs": [tên tab đã đọc], "cols": {(ngày, tháng) có cột},
    "nguoi": [{"ten", "co_so", "ngay": {(ngày, tháng) đã đăng ký}}] theo thứ tự trên lịch,
    "reg": {tên chuẩn hóa: {(ngày, tháng) đã đăng ký}}, "co_so": {tên chuẩn hóa: cơ sở}}}.

    Đọc 1 lần mỗi lần chạy (cache). Google trả 429/5xx tạm thời -> tự thử lại
    (cùng cơ chế với các tab thưởng/doanh số) thay vì bỏ qua trừ thưởng Chủ nhật.
    """
    global _cache
    if _cache is not None:
        return _cache
    try:
        _cache = _doc_tu_google()
    except SystemExit as e:
        # _with_retry đổi 403 thành SystemExit (dành cho trang tính ghi). Lịch trực
        # chỉ là dữ liệu phụ: KHÔNG được dừng cả lần cập nhật, chỉ báo lỗi để cảnh báo.
        raise PermissionError(
            "Service account chưa được chia sẻ Lịch trực ngày nghỉ (403)") from e
    return _cache


@google_sheet._with_retry
def _doc_tu_google() -> dict[str, dict]:
    ss = google_sheet._connect().open_by_key(config.GOOGLE_SHEET_ID_LICH_TRUC)
    worksheets = ss.worksheets()
    out: dict[str, dict] = {}
    for nhom, tu_khoa in TU_KHOA_TAB.items():
        tabs: list[str] = []
        cols: set[tuple[int, int]] = set()
        nguoi: dict[str, dict] = {}          # tên chuẩn hóa -> người (giữ thứ tự trên lịch)
        for ws in worksheets:
            if tu_khoa not in chuan_hoa_ten(ws.title):
                continue
            vals = ws.get_all_values()
            h = next((i for i, r in enumerate(vals)
                      if any(chuan_hoa_ten(c) == "ho ten" for c in r)), None)
            if h is None:
                continue                     # tab khác cấu trúc (lịch mẫu cũ) -> bỏ qua
            tabs.append(ws.title)
            date_cols: dict[int, tuple[int, int]] = {}
            for j, c in enumerate(vals[h]):
                m = _NGAY_RE.match(c or "")
                if m:
                    date_cols[j] = (int(m.group(1)), int(m.group(2)))
            cols |= set(date_cols.values())
            for row in vals[h + 1:]:
                ten_goc = " ".join((row[1] if len(row) > 1 else "").split())
                name = chuan_hoa_ten(ten_goc)
                if not name:
                    continue
                days = {dm for j, dm in date_cols.items()
                        if j < len(row) and "dang k" in chuan_hoa_ten(row[j])}
                p = nguoi.setdefault(name, {"ten": ten_goc, "co_so": "", "ngay": set()})
                p["ngay"] |= days
                p["co_so"] = p["co_so"] or (row[2].strip() if len(row) > 2 else "")
        out[nhom] = {"tabs": tabs, "cols": cols, "nguoi": list(nguoi.values()),
                     "reg": {k: v["ngay"] for k, v in nguoi.items()},
                     "co_so": {k: v["co_so"] for k, v in nguoi.items()}}
    return out
