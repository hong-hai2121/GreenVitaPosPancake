# -*- coding: utf-8 -*-
"""Tạo ICON + LỐI TẮT Desktop + LỐI TẮT KHỞI ĐỘNG CÙNG WINDOWS cho ứng dụng CẬP NHẬT THƯỞNG.

Chạy 1 lần (hoặc chạy lại khi đổi máy / đổi thư mục):   python tao_loi_tat.py

- Vẽ file icon_app.ico (chữ GV trên nền xanh lá) bằng Pillow - nếu đã có thì giữ nguyên.
- Tạo "GreenVita - Cập nhật thưởng.lnk" trên Desktop: chạy pythonw.exe app_cap_nhat.pyw
  (không hiện cửa sổ đen), thư mục làm việc = thư mục dự án, icon = icon_app.ico.
- Tạo "GreenVita - Cập nhật thưởng (khay).lnk" trong thư mục Startup của Windows
  (shell:startup): mỗi lần đăng nhập máy, app tự chạy ẨN dưới khay hệ thống
  (app_cap_nhat.pyw --khay) giống Unikey. Bỏ tự khởi động: chạy
  `python tao_loi_tat.py --bo-khoi-dong` hoặc bỏ tick "Khởi động cùng Windows" trên app.
- Tạo lại thì ghi đè lối tắt cũ, không tạo trùng.
"""
from __future__ import annotations

import base64
import os
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
APP_FILE = BASE_DIR / "app_cap_nhat.pyw"
ICO_FILE = BASE_DIR / "icon_app.ico"
TEN_LOI_TAT = "GreenVita - Cập nhật thưởng.lnk"
TEN_LOI_TAT_KHOI_DONG = "GreenVita - Cập nhật thưởng (khay).lnk"
# Thư mục Startup của người dùng: mọi .lnk trong đây được Windows chạy khi đăng nhập
THU_MUC_KHOI_DONG = (Path(os.environ.get("APPDATA", Path.home() / "AppData/Roaming"))
                     / "Microsoft/Windows/Start Menu/Programs/Startup")
LOI_TAT_KHOI_DONG = THU_MUC_KHOI_DONG / TEN_LOI_TAT_KHOI_DONG
# Mô tả (tooltip) phải KHÔNG DẤU: WScript.Shell ghi tiếng Việt có dấu thành dấu "?"
MO_TA = "Cap nhat thuong GreenVita (Pancake POS -> Google Sheet)"

XANH_LA = (31, 122, 58)
TRANG = (255, 255, 255)


def ve_icon(path: Path) -> None:
    """Vẽ icon vuông bo góc, nền xanh lá, chữ GV trắng; lưu nhiều cỡ trong 1 file .ico."""
    from PIL import Image, ImageDraw, ImageFont

    size = 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((8, 8, size - 9, size - 9), radius=56, fill=XANH_LA)

    try:
        font = ImageFont.truetype("arialbd.ttf", 124)      # có sẵn trên Windows
    except OSError:
        font = ImageFont.load_default()
    text = "GV"
    x0, y0, x1, y1 = d.textbbox((0, 0), text, font=font)
    d.text(((size - (x1 - x0)) / 2 - x0, (size - (y1 - y0)) / 2 - y0),
           text, font=font, fill=TRANG)

    img.save(path, format="ICO",
             sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])


def duong_dan_pythonw() -> Path:
    """pythonw.exe cùng thư mục với python đang chạy (không hiện cửa sổ console)."""
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    return pyw if pyw.exists() else exe


def _ps(s: str) -> str:
    """Bọc chuỗi cho PowerShell (nháy đơn, nhân đôi nháy đơn bên trong)."""
    return "'" + s.replace("'", "''") + "'"


def tao_loi_tat(thu_muc: Path | None = None, ten: str = TEN_LOI_TAT,
                tham_so: str = "") -> Path:
    """Tạo .lnk qua WScript.Shell (PowerShell). Trả về đường dẫn lối tắt đã tạo.

    thu_muc: thư mục đặt lối tắt (None = Desktop, để PowerShell tự tìm - đúng cả khi
    Desktop bị chuyển vào OneDrive). tham_so: tham số thêm sau tên file .pyw (vd --khay).

    WScript.Shell KHÔNG lưu được .lnk có tên chứa dấu tiếng Việt (báo "Unable to save
    shortcut ...C?p nh?t..."), nên tạo với tên tạm không dấu rồi đổi tên bằng Python.
    """
    if thu_muc is None:
        dong_thu_muc = "$thumuc = [Environment]::GetFolderPath('Desktop')"
    else:
        thu_muc.mkdir(parents=True, exist_ok=True)
        dong_thu_muc = f"$thumuc = {_ps(str(thu_muc))}"
    doi_so = '"' + str(APP_FILE) + '"' + (f" {tham_so}" if tham_so else "")
    script = f"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
{dong_thu_muc}
$tam = Join-Path $thumuc 'GreenVita_tmp_shortcut.lnk'
$s = (New-Object -ComObject WScript.Shell).CreateShortcut($tam)
$s.TargetPath = {_ps(str(duong_dan_pythonw()))}
$s.Arguments = {_ps(doi_so)}
$s.WorkingDirectory = {_ps(str(BASE_DIR))}
$s.IconLocation = {_ps(str(ICO_FILE) + ",0")}
$s.Description = {_ps(MO_TA)}
$s.WindowStyle = 1
$s.Save()
if (-not (Test-Path -LiteralPath $tam)) {{ throw "Save() không tạo được file $tam" }}
Write-Output $tam
"""
    # -EncodedCommand: truyền script dạng UTF-16LE base64 -> không lỗi tiếng Việt / dấu nháy
    enc = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-EncodedCommand", enc],
                       capture_output=True, text=True, encoding="utf-8",
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if r.returncode != 0:
        raise RuntimeError(f"Tạo lối tắt thất bại: {r.stderr or r.stdout}")
    tam = Path(r.stdout.strip().splitlines()[-1])
    dich = tam.with_name(ten)
    tam.replace(dich)                       # ghi đè lối tắt cũ nếu đã có
    if not dich.exists():
        raise RuntimeError(f"Không đổi tên được lối tắt sang: {dich}")
    return dich


# ----------------------------------------------------------------------
# Khởi động cùng Windows (lối tắt trong thư mục Startup, chạy ẩn ở khay)
# ----------------------------------------------------------------------
def co_khoi_dong_cung_windows() -> bool:
    return LOI_TAT_KHOI_DONG.exists()


def bat_khoi_dong_cung_windows() -> Path:
    """Tạo lối tắt trong Startup: đăng nhập Windows là app tự chạy ẩn ở khay (--khay)."""
    return tao_loi_tat(THU_MUC_KHOI_DONG, TEN_LOI_TAT_KHOI_DONG, "--khay")


def tat_khoi_dong_cung_windows() -> bool:
    """Xóa lối tắt trong Startup. Trả về True nếu có file để xóa."""
    if LOI_TAT_KHOI_DONG.exists():
        LOI_TAT_KHOI_DONG.unlink()
        return True
    return False


def main() -> None:
    if not APP_FILE.exists():
        raise SystemExit(f"Không thấy {APP_FILE}")
    if "--bo-khoi-dong" in sys.argv:
        if tat_khoi_dong_cung_windows():
            print(f"Đã bỏ khởi động cùng Windows (xóa {LOI_TAT_KHOI_DONG})")
        else:
            print("App chưa được đặt khởi động cùng Windows - không có gì để bỏ.")
        return
    if ICO_FILE.exists():
        print(f"Icon đã có, giữ nguyên: {ICO_FILE}")
    else:
        ve_icon(ICO_FILE)
        print(f"Đã vẽ icon: {ICO_FILE}")
    try:
        lnk = tao_loi_tat()
    except RuntimeError as e:
        raise SystemExit(str(e))
    print(f"Đã tạo lối tắt Desktop: {lnk}")
    print(f"  -> {duong_dan_pythonw()} \"{APP_FILE}\"")
    try:
        lnk2 = bat_khoi_dong_cung_windows()
    except RuntimeError as e:
        raise SystemExit(str(e))
    print(f"Đã tạo lối tắt khởi động cùng Windows: {lnk2}")
    print(f"  -> {duong_dan_pythonw()} \"{APP_FILE}\" --khay  (chạy ẩn ở khay khi đăng nhập)")
    print("Nháy đúp icon 'GreenVita - Cập nhật thưởng' trên Desktop để mở cửa sổ;")
    print("đóng cửa sổ thì app vẫn chạy ngầm ở khay (góc phải thanh taskbar).")


if __name__ == "__main__":
    main()
