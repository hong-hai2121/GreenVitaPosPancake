# -*- coding: utf-8 -*-
"""Ghi bảng thưởng tháng lên Google Sheets bằng service account.

- Mở sheet theo GOOGLE_SHEET_ID trong .env; nếu chưa có thì tự tạo sheet mới,
  chia sẻ cho GOOGLE_SHARE_EMAIL và tự ghi ID vào .env.
- write_table(): ghi đè + tô màu 1 tab bảng thưởng (dùng bởi thuong_thang.py).
"""
from __future__ import annotations

import re

import gspread
from google.oauth2.service_account import Credentials

import config

SPREADSHEET_NAME = "Doanh Thu GreenVita POS"


def _col_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _connect() -> gspread.Client:
    if not config.SERVICE_ACCOUNT_FILE.exists():
        raise SystemExit(f"Không tìm thấy file service account: {config.SERVICE_ACCOUNT_FILE}")
    creds = Credentials.from_service_account_file(
        str(config.SERVICE_ACCOUNT_FILE), scopes=config.GOOGLE_SCOPES
    )
    return gspread.authorize(creds)


def _save_sheet_id_to_env(sheet_id: str) -> None:
    """Ghi GOOGLE_SHEET_ID vào .env để lần sau dùng lại đúng sheet."""
    env_path = config.BASE_DIR / ".env"
    text = env_path.read_text(encoding="utf-8") if env_path.exists() else ""
    if re.search(r"^GOOGLE_SHEET_ID=.*$", text, flags=re.M):
        text = re.sub(r"^GOOGLE_SHEET_ID=.*$", f"GOOGLE_SHEET_ID={sheet_id}", text, flags=re.M)
    else:
        text += f"\nGOOGLE_SHEET_ID={sheet_id}\n"
    env_path.write_text(text, encoding="utf-8")


_cached_ss = None


def _spreadsheet() -> "gspread.Spreadsheet":
    """Kết nối + mở spreadsheet, cache lại để nhiều thao tác trong 1 lần chạy dùng chung."""
    global _cached_ss
    if _cached_ss is None:
        client = _connect()
        _cached_ss, _ = _open_or_create(client)
    return _cached_ss


def read_table(tab_title: str) -> list[list[str]] | None:
    """Đọc toàn bộ giá trị (đã định dạng) của 1 tab; None nếu tab chưa tồn tại."""
    ss = _spreadsheet()
    for ws in ss.worksheets():
        if ws.title == tab_title:
            return ws.get_all_values()
    return None


def _open_or_create(client: gspread.Client) -> tuple[gspread.Spreadsheet, bool]:
    if config.GOOGLE_SHEET_ID:
        return client.open_by_key(config.GOOGLE_SHEET_ID), False

    try:
        ss = client.create(SPREADSHEET_NAME)
    except gspread.exceptions.APIError as e:
        # Google không cấp dung lượng Drive cho service account -> không tự tạo được file
        raise SystemExit(
            "Service account không tự tạo được Google Sheet (Drive không có dung lượng).\n"
            "Cách khắc phục:\n"
            "1. Bạn tự tạo 1 Google Sheet mới trong Drive của mình\n"
            "2. Chia sẻ quyền Editor cho email service account (xem client_email trong service_account.json)\n"
            "3. Dán ID của sheet (đoạn giữa /d/ và /edit trên URL) vào GOOGLE_SHEET_ID trong .env\n"
            f"Chi tiết lỗi: {e}"
        ) from e
    if config.GOOGLE_SHARE_EMAIL:
        ss.share(config.GOOGLE_SHARE_EMAIL, perm_type="user", role="writer", notify=True)
    _save_sheet_id_to_env(ss.id)
    return ss, True


# ---- Màu sắc bảng thưởng tháng ----
def _rgb(hex_str: str) -> dict:
    h = hex_str.lstrip("#")
    return {"red": int(h[0:2], 16) / 255,
            "green": int(h[2:4], 16) / 255,
            "blue": int(h[4:6], 16) / 255}


C_HEADER_BG = _rgb("1F7A4D")     # xanh lá đậm
C_TITLE_TEXT = _rgb("145A32")
C_STRIPE_BG = _rgb("F1F7F3")     # sọc xen kẽ xanh rất nhạt
C_BONUS_BG = _rgb("D5EEDD")      # ô đạt thưởng
C_BONUS_TEXT = _rgb("145A32")
C_TOTALCOL_BG = _rgb("FFF3D6")   # cột Tổng tháng - vàng nhạt
C_TOTALROW_BG = _rgb("DCEDE3")   # dòng Tổng - xanh nhạt
C_WHITE = _rgb("FFFFFF")
C_BORDER = _rgb("B7C7BD")
# Cột CHỦ NHẬT - tông cam để nhận biết ngay
C_SUNDAY_HEADER = _rgb("D97706")   # header cam đậm
C_SUNDAY_BG = _rgb("FDF1DC")       # thân cột cam rất nhạt
C_SUNDAY_BONUS_BG = _rgb("FADFAE") # ô đạt thưởng Chủ nhật
C_SUNDAY_BONUS_TEXT = _rgb("8A4B08")


def _delete_conditional_rules(ss, sheet_id: int) -> list[dict]:
    """Yêu cầu xóa các rule định dạng có điều kiện cũ của tab (tránh trùng khi chạy lại)."""
    meta = ss.fetch_sheet_metadata({"fields": "sheets(properties.sheetId,conditionalFormats)"})
    count = 0
    for sh in meta.get("sheets", []):
        if sh.get("properties", {}).get("sheetId") == sheet_id:
            count = len(sh.get("conditionalFormats", []))
    return [{"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": 0}}
            for _ in range(count)]


def _style_month_table(ss, ws, n_rows: int, n_cols: int, n_fixed_cols: int = 3,
                       sunday_cols: list[int] | None = None) -> None:
    """Tô màu bảng thưởng: title, header, sọc xen kẽ, ô đạt thưởng, tổng, khung, độ rộng cột.

    sunday_cols: chỉ số cột (0-based) của các ngày Chủ nhật -> tô tông cam nhận biết.
    """
    sunday_cols = sunday_cols or []
    sid = ws.id
    day_c0, day_c1 = n_fixed_cols, n_cols - 1     # các cột ngày (0-based, half-open)
    data_r0, data_r1 = 2, n_rows - 1              # các dòng nhân viên
    total_row = n_rows - 1

    def grid(r0, r1, c0, c1):
        return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    def repeat(rng, fmt, fields):
        return {"repeatCell": {"range": rng, "cell": {"userEnteredFormat": fmt},
                               "fields": fields}}

    req = []
    # 0. Xóa định dạng cũ toàn vùng bảng (chạy lại không bị lem màu cũ)
    req.append(repeat(grid(0, n_rows + 5, 0, n_cols + 2), {}, "userEnteredFormat"))
    req += _delete_conditional_rules(ss, sid)

    # 1. Tiêu đề
    req.append(repeat(grid(0, 1, 0, n_cols),
                      {"textFormat": {"bold": True, "fontSize": 13,
                                      "foregroundColor": C_TITLE_TEXT}},
                      "userEnteredFormat.textFormat"))
    # 2. Header: nền xanh đậm, chữ trắng đậm, căn giữa
    req.append(repeat(grid(1, 2, 0, n_cols),
                      {"backgroundColor": C_HEADER_BG,
                       "textFormat": {"bold": True, "foregroundColor": C_WHITE},
                       "horizontalAlignment": "CENTER",
                       "verticalAlignment": "MIDDLE",
                       "wrapStrategy": "WRAP"},
                      "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,"
                      "verticalAlignment,wrapStrategy)"))
    # 2b. Header các cột ngày: Sheets lưu "01/08/2026" thành số ngày (serial),
    #     phải gán định dạng DATE để hiển thị đúng dd/mm/yyyy (nếu không sẽ ra 46235...)
    req.append(repeat(grid(1, 2, day_c0, day_c1),
                      {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"}},
                      "userEnteredFormat.numberFormat"))
    # 3. Sọc xen kẽ các dòng nhân viên (dòng lẻ tính từ dòng đầu dữ liệu)
    for r in range(data_r0 + 1, data_r1, 2):
        req.append(repeat(grid(r, r + 1, 0, n_cols),
                          {"backgroundColor": C_STRIPE_BG},
                          "userEnteredFormat.backgroundColor"))
    # 3b. Cột CHỦ NHẬT: header cam đậm + thân cột nền cam nhạt (đè lên sọc)
    for c in sunday_cols:
        req.append(repeat(grid(1, 2, c, c + 1),
                          {"backgroundColor": C_SUNDAY_HEADER},
                          "userEnteredFormat.backgroundColor"))
        req.append(repeat(grid(data_r0, data_r1, c, c + 1),
                          {"backgroundColor": C_SUNDAY_BG},
                          "userEnteredFormat.backgroundColor"))
    # 4. Cột STT căn giữa
    req.append(repeat(grid(data_r0, data_r1, 0, 1),
                      {"horizontalAlignment": "CENTER"},
                      "userEnteredFormat.horizontalAlignment"))
    # 5. Định dạng số #,##0 cho vùng tiền (cột ngày + Tổng tháng, cả dòng Tổng)
    req.append(repeat(grid(data_r0, n_rows, day_c0, n_cols),
                      {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}},
                      "userEnteredFormat.numberFormat"))
    # 6. Cột Tổng tháng: vàng nhạt + đậm
    req.append(repeat(grid(data_r0, data_r1, n_cols - 1, n_cols),
                      {"backgroundColor": C_TOTALCOL_BG, "textFormat": {"bold": True}},
                      "userEnteredFormat(backgroundColor,textFormat.bold)"))
    # 7. Dòng Tổng: xanh nhạt + đậm
    req.append(repeat(grid(total_row, n_rows, 0, n_cols),
                      {"backgroundColor": C_TOTALROW_BG, "textFormat": {"bold": True}},
                      "userEnteredFormat(backgroundColor,textFormat.bold)"))
    # 8. Ô đạt thưởng (>0) trong vùng cột ngày: nền xanh, chữ xanh đậm (conditional
    #    formatting nên sửa số trực tiếp trên sheet màu vẫn tự cập nhật)
    req.append({"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [grid(data_r0, data_r1, day_c0, day_c1)],
        "booleanRule": {
            "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
            "format": {"backgroundColor": C_BONUS_BG,
                       "textFormat": {"bold": True, "foregroundColor": C_BONUS_TEXT}},
        }}}})
    # 8b. Ô đạt thưởng CHỦ NHẬT: tô cam thay vì xanh (index 0 -> ưu tiên hơn rule xanh)
    if sunday_cols:
        req.append({"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [grid(data_r0, data_r1, c, c + 1) for c in sunday_cols],
            "booleanRule": {
                "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                "format": {"backgroundColor": C_SUNDAY_BONUS_BG,
                           "textFormat": {"bold": True, "foregroundColor": C_SUNDAY_BONUS_TEXT}},
            }}}})
    # 9. Kẻ khung cả bảng
    border = {"style": "SOLID", "color": C_BORDER}
    req.append({"updateBorders": {"range": grid(1, n_rows, 0, n_cols),
                                  "top": border, "bottom": border, "left": border,
                                  "right": border, "innerHorizontal": border,
                                  "innerVertical": border}})
    # 10. Độ rộng cột: STT / Họ tên / Bộ phận / các ngày / Tổng tháng
    for c0, c1, px in [(0, 1, 40), (1, 2, 230), (2, 3, 110),
                       (day_c0, day_c1, 88), (n_cols - 1, n_cols, 115)]:
        req.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": c0, "endIndex": c1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    # 11. Cố định 2 dòng đầu + 3 cột đầu
    req.append({"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": 2, "frozenColumnCount": 3}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}})

    ss.batch_update({"requests": req})


C_MANUAL_BG = _rgb("FFF9E0")     # ô nhập tay (Thưởng, % Thưởng) - vàng nhạt


def _get_or_create_ws(ss, tab_title: str, n_rows: int, n_cols: int):
    for ws in ss.worksheets():
        if ws.title == tab_title:
            break
    else:
        ws = ss.add_worksheet(title=tab_title, rows=n_rows, cols=n_cols)
    if ws.col_count < n_cols or ws.row_count < n_rows:
        ws.resize(rows=max(ws.row_count, n_rows), cols=max(ws.col_count, n_cols))
    return ws


def _style_bc02(ss, ws, n_rows: int, n_cols: int) -> None:
    """Tô màu bảng BC02 (thưởng doanh số tháng): dòng Tổng ở TRÊN (dòng 3),
    nhân viên từ dòng 4; cột Thưởng / %%Thưởng nền vàng = nhập tay."""
    sid = ws.id
    total_r = 2                 # dòng Tổng (0-based)
    data_r0, data_r1 = 3, n_rows

    def grid(r0, r1, c0, c1):
        return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    def repeat(rng, fmt, fields):
        return {"repeatCell": {"range": rng, "cell": {"userEnteredFormat": fmt},
                               "fields": fields}}

    req = [repeat(grid(0, n_rows + 5, 0, n_cols + 2), {}, "userEnteredFormat")]
    req += _delete_conditional_rules(ss, sid)
    req.append(repeat(grid(0, 1, 0, n_cols),
                      {"textFormat": {"bold": True, "fontSize": 13,
                                      "foregroundColor": C_TITLE_TEXT}},
                      "userEnteredFormat.textFormat"))
    req.append(repeat(grid(1, 2, 0, n_cols),
                      {"backgroundColor": C_HEADER_BG,
                       "textFormat": {"bold": True, "foregroundColor": C_WHITE},
                       "horizontalAlignment": "CENTER",
                       "verticalAlignment": "MIDDLE",
                       "wrapStrategy": "WRAP"},
                      "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,"
                      "verticalAlignment,wrapStrategy)"))
    # Dòng Tổng (ngay dưới header)
    req.append(repeat(grid(total_r, total_r + 1, 0, n_cols),
                      {"backgroundColor": C_TOTALROW_BG, "textFormat": {"bold": True}},
                      "userEnteredFormat(backgroundColor,textFormat.bold)"))
    # Sọc xen kẽ nhân viên
    for r in range(data_r0 + 1, data_r1, 2):
        req.append(repeat(grid(r, r + 1, 0, n_cols),
                          {"backgroundColor": C_STRIPE_BG},
                          "userEnteredFormat.backgroundColor"))
    # STT căn giữa
    req.append(repeat(grid(data_r0, data_r1, 0, 1),
                      {"horizontalAlignment": "CENTER"},
                      "userEnteredFormat.horizontalAlignment"))
    # Số #,##0 cho vùng D..K, sau đó đè % cho cột Tỷ lệ hoàn (H=7) và % Thưởng (J=9)
    req.append(repeat(grid(total_r, n_rows, 3, n_cols),
                      {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}},
                      "userEnteredFormat.numberFormat"))
    for c in (7, 9):
        req.append(repeat(grid(total_r, n_rows, c, c + 1),
                          {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}},
                          "userEnteredFormat.numberFormat"))
    # Cột nhập tay: chỉ % Thưởng (J=9) - nền vàng nhạt
    # (cột Thưởng I tự lấy từ Tổng tháng của 2 tab Thưởng GR, không nhập tay)
    req.append(repeat(grid(data_r0, data_r1, 9, 10),
                      {"backgroundColor": C_MANUAL_BG},
                      "userEnteredFormat.backgroundColor"))
    # Cột Thực nhận (J=9): vàng đậm hơn + in đậm
    req.append(repeat(grid(total_r, n_rows, n_cols - 1, n_cols),
                      {"backgroundColor": C_TOTALCOL_BG, "textFormat": {"bold": True}},
                      "userEnteredFormat(backgroundColor,textFormat.bold)"))
    border = {"style": "SOLID", "color": C_BORDER}
    req.append({"updateBorders": {"range": grid(1, n_rows, 0, n_cols),
                                  "top": border, "bottom": border, "left": border,
                                  "right": border, "innerHorizontal": border,
                                  "innerVertical": border}})
    widths = [(0, 1, 40), (1, 2, 230), (2, 3, 110), (3, 4, 90), (4, 5, 105),
              (5, 6, 105), (6, 7, 135), (7, 8, 90), (8, 9, 115), (9, 10, 90),
              (10, 11, 125)]
    for c0, c1, px in widths:
        req.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": c0, "endIndex": c1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    req.append({"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": 2, "frozenColumnCount": 3}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}})
    ss.batch_update({"requests": req})


def write_bc02_table(tab_title: str, values: list[list]) -> str:
    """Ghi đè + tô màu tab BC02 (thưởng doanh số tháng). Trả về URL spreadsheet."""
    ss = _spreadsheet()
    ws = _get_or_create_ws(ss, tab_title, max(len(values) + 5, 50),
                           max(len(values[1]) + 2, 12))
    ws.clear()
    end = _col_letter(len(values[1]))
    ws.update(values=values, range_name=f"A1:{end}{len(values)}",
              value_input_option="USER_ENTERED")
    _style_bc02(ss, ws, n_rows=len(values), n_cols=len(values[1]))
    return ss.url


def write_table(tab_title: str, values: list[list], money_range: str | None = None,
                sunday_cols: list[int] | None = None) -> str:
    """Ghi đè toàn bộ 1 tab bằng ma trận `values` (bảng thưởng tháng) rồi tô màu.

    - Dòng 1: tiêu đề; dòng 2: header; dòng cuối: Tổng; 3 cột đầu cố định.
    Trả về URL của spreadsheet.
    """
    ss = _spreadsheet()

    n_rows = max(len(values) + 5, 50)
    n_cols = max(len(values[1]) + 2, 10)
    for ws in ss.worksheets():
        if ws.title == tab_title:
            break
    else:
        ws = ss.add_worksheet(title=tab_title, rows=n_rows, cols=n_cols)

    if ws.col_count < n_cols or ws.row_count < n_rows:
        ws.resize(rows=max(ws.row_count, n_rows), cols=max(ws.col_count, n_cols))

    ws.clear()
    end = _col_letter(len(values[1]))
    ws.update(values=values, range_name=f"A1:{end}{len(values)}",
              value_input_option="USER_ENTERED")
    _style_month_table(ss, ws, n_rows=len(values), n_cols=len(values[1]),
                       sunday_cols=sunday_cols)
    return ss.url
