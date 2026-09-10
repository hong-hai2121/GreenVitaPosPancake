# -*- coding: utf-8 -*-
"""Đọc "Lịch trực ngày nghỉ - Greenvita" (Google Sheet riêng, GOOGLE_SHEET_ID_LICH_TRUC).

Tab "Bộ phận SALE" / "Bộ phận CSKH": dòng header có STT | Họ Tên | Cơ sở |
các cột ngày CHỦ NHẬT dạng d/m; ô ghi "Đăng kí làm" = có đăng ký trực hôm đó.

Quy tắc áp ở thuong_thang.py: Chủ nhật KHÔNG đăng ký làm -> thưởng ngày đó = 0.
"""
from __future__ import annotations

import re
import unicodedata

import config
import google_sheet

TABS = {"sale": "Bộ phận SALE", "cskh": "Bộ phận CSKH"}
_NGAY_RE = re.compile(r"^\s*(\d{1,2})\s*/\s*(\d{1,2})\s*$")


def chuan_hoa_ten(s: str) -> str:
    """Chuẩn hóa để so khớp tên: thường hóa, BỎ DẤU, gọn khoảng trắng
    (lịch trực viết tay có thể thiếu dấu so với tên trên Pancake)."""
    s = (s or "").replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.lower().split())


_cache: dict | None = None


def doc_dang_ky() -> dict[str, dict]:
    """{nhóm: {"cols": {(ngày, tháng) có cột}, "reg": {tên chuẩn hóa: {(ngày, tháng) đã đăng ký}}}}.

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
    worksheets = {chuan_hoa_ten(w.title): w for w in ss.worksheets()}
    out: dict[str, dict] = {}
    for nhom, tab in TABS.items():
        cols: set[tuple[int, int]] = set()
        reg: dict[str, set[tuple[int, int]]] = {}
        ws = worksheets.get(chuan_hoa_ten(tab))
        if ws is not None:
            vals = ws.get_all_values()
            h = next((i for i, r in enumerate(vals)
                      if any(chuan_hoa_ten(c) == "ho ten" for c in r)), None)
            if h is not None:
                date_cols: dict[int, tuple[int, int]] = {}
                for j, c in enumerate(vals[h]):
                    m = _NGAY_RE.match(c or "")
                    if m:
                        date_cols[j] = (int(m.group(1)), int(m.group(2)))
                cols = set(date_cols.values())
                for row in vals[h + 1:]:
                    name = chuan_hoa_ten(row[1] if len(row) > 1 else "")
                    if not name:
                        continue
                    days = {dm for j, dm in date_cols.items()
                            if j < len(row) and "dang k" in chuan_hoa_ten(row[j])}
                    reg[name] = reg.get(name, set()) | days
        out[nhom] = {"cols": cols, "reg": reg}
    return out
