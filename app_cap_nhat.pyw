# -*- coding: utf-8 -*-
"""Ứng dụng desktop CẬP NHẬT THƯỞNG GREENVITA (Pancake POS -> Google Sheet).

- Nháy đúp file này -> mở cửa sổ: tự chạy cập nhật ngay, hiển thị tiến trình,
  rồi ĐẾM NGƯỢC tới 9h sáng hôm sau và tự chạy tiếp.
- Nút bấm: Cập nhật ngay / Mở Google Sheet / Mở file log.
- Chế độ chạy ngầm cho Task Scheduler:  pythonw app_cap_nhat.pyw --ngam
  (chạy 1 lần, ghi log rồi thoát - không mở cửa sổ).

Mọi lần chạy đều ghi thêm vào logs/cap_nhat.log.
"""
import json
import os
import queue
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
LOG_FILE = LOG_DIR / "cap_nhat.log"
CAI_DAT_FILE = BASE_DIR / "cai_dat_app.json"
GIO_CHAY_MAC_DINH = (9, 0)       # giờ:phút tự chạy hằng ngày (mặc định 9h sáng)


def doc_gio_chay() -> tuple[int, int]:
    try:
        d = json.loads(CAI_DAT_FILE.read_text(encoding="utf-8"))
        return int(d["gio"]) % 24, int(d["phut"]) % 60
    except Exception:
        return GIO_CHAY_MAC_DINH


def luu_gio_chay(gio: int, phut: int) -> None:
    CAI_DAT_FILE.write_text(json.dumps({"gio": gio, "phut": phut}), encoding="utf-8")

XANH_DAM = "#1F7A4D"
XANH_NHAT = "#E8F3EC"
CAM = "#D97706"
DO = "#C0392B"
XAM = "#5F6B66"


def lan_chay_tiep_theo(now: datetime, gio: int, phut: int) -> datetime:
    target = now.replace(hour=gio, minute=phut, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return target


def chay_cap_nhat(on_line) -> int:
    """Chạy thuong_thang.py, gọi on_line(dòng) cho từng dòng, ghi log. Trả về mã lỗi."""
    LOG_DIR.mkdir(exist_ok=True)
    proc = subprocess.Popen(
        [sys.executable, str(BASE_DIR / "thuong_thang.py")],
        cwd=str(BASE_DIR), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"\n==================== {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} ====================\n")
        for line in proc.stdout:
            on_line(line.rstrip("\n"))
            f.write(line)
        code = proc.wait()
        if code != 0:
            f.write(f"[LỖI] Mã thoát: {code}\n")
    return code


def chay_ngam() -> None:
    """Chế độ Task Scheduler: chạy 1 lần, ghi log, thoát (không cửa sổ, không print)."""
    chay_cap_nhat(lambda line: None)


# ======================================================================
# Giao diện
# ======================================================================
def chay_giao_dien() -> None:
    import tkinter as tk
    from tkinter import scrolledtext, ttk

    class App:
        def __init__(self, root: tk.Tk):
            self.root = root
            self.dang_chay = False
            self.hang_doi: queue.Queue = queue.Queue()
            self.gio_chay, self.phut_chay = doc_gio_chay()
            self.lan_tiep_theo = lan_chay_tiep_theo(datetime.now(),
                                                    self.gio_chay, self.phut_chay)

            root.title("Cập Nhật Thưởng GreenVita")
            # Hiện cửa sổ ở CHÍNH GIỮA màn hình
            w, h = 720, 520
            x = (root.winfo_screenwidth() - w) // 2
            y = (root.winfo_screenheight() - h) // 2
            root.geometry(f"{w}x{h}+{x}+{y}")
            root.minsize(600, 420)
            root.configure(bg=XANH_NHAT)

            # --- Tiêu đề ---
            tk.Label(root, text="CẬP NHẬT THƯỞNG GREENVITA",
                     font=("Segoe UI", 16, "bold"), fg="white", bg=XANH_DAM,
                     pady=10).pack(fill="x")

            # --- Trạng thái + đồng hồ ---
            self.lb_trang_thai = tk.Label(root, text="Sẵn sàng",
                                          font=("Segoe UI", 12, "bold"),
                                          fg=XANH_DAM, bg=XANH_NHAT, pady=6)
            self.lb_trang_thai.pack()
            self.lb_dong_ho = tk.Label(root, text="--:--:--",
                                       font=("Consolas", 30, "bold"),
                                       fg=XANH_DAM, bg=XANH_NHAT)
            self.lb_dong_ho.pack()
            self.lb_lan_sau = tk.Label(root, text="", font=("Segoe UI", 10),
                                       fg=XAM, bg=XANH_NHAT)
            self.lb_lan_sau.pack(pady=(0, 8))

            # --- Chỉnh giờ tự chạy hằng ngày ---
            khung_gio = tk.Frame(root, bg=XANH_NHAT)
            khung_gio.pack(pady=(0, 2))
            tk.Label(khung_gio, text="Giờ tự cập nhật hằng ngày:",
                     font=("Segoe UI", 10), fg=XAM, bg=XANH_NHAT).grid(row=0, column=0, padx=(0, 6))
            self.bien_gio = tk.StringVar(value=f"{self.gio_chay:02d}")
            self.bien_phut = tk.StringVar(value=f"{self.phut_chay:02d}")
            tk.Spinbox(khung_gio, from_=0, to=23, width=3, format="%02.0f",
                       font=("Segoe UI", 11), textvariable=self.bien_gio, wrap=True,
                       command=self.doi_gio, justify="center").grid(row=0, column=1)
            tk.Label(khung_gio, text=":", font=("Segoe UI", 11, "bold"),
                     bg=XANH_NHAT).grid(row=0, column=2)
            tk.Spinbox(khung_gio, from_=0, to=59, width=3, format="%02.0f",
                       font=("Segoe UI", 11), textvariable=self.bien_phut, wrap=True,
                       command=self.doi_gio, justify="center").grid(row=0, column=3)

            # --- Nút ---
            khung_nut = tk.Frame(root, bg=XANH_NHAT)
            khung_nut.pack(pady=4)
            style = ttk.Style()
            style.configure("TButton", font=("Segoe UI", 10), padding=6)
            self.nut_chay = ttk.Button(khung_nut, text="  Cập nhật ngay  ",
                                       command=self.bam_cap_nhat)
            self.nut_chay.grid(row=0, column=0, padx=6)
            ttk.Button(khung_nut, text="  Mở Google Sheet  ",
                       command=self.mo_sheet).grid(row=0, column=1, padx=6)
            ttk.Button(khung_nut, text="  Mở file log  ",
                       command=self.mo_log).grid(row=0, column=2, padx=6)

            # --- Khung tiến trình ---
            self.khung_log = scrolledtext.ScrolledText(
                root, font=("Consolas", 10), state="disabled", wrap="word",
                bg="white", relief="flat", borderwidth=6)
            self.khung_log.pack(fill="both", expand=True, padx=12, pady=(6, 12))

            self.ghi("Chào mừng! Bấm 'Cập nhật ngay' để chạy thủ công,")
            self.ghi("hoặc chờ đồng hồ đếm ngược - đến giờ hẹn ứng dụng tự cập nhật.\n")
            self.cap_nhat_nhan_lan_sau()

            self.root.after(200, self.vong_lap)

        # ------------------------------------------------------------------
        def doi_gio(self) -> None:
            """Người dùng chỉnh giờ tự chạy trên giao diện."""
            try:
                self.gio_chay = int(self.bien_gio.get()) % 24
                self.phut_chay = int(self.bien_phut.get()) % 60
            except ValueError:
                return
            luu_gio_chay(self.gio_chay, self.phut_chay)
            self.lan_tiep_theo = lan_chay_tiep_theo(datetime.now(),
                                                    self.gio_chay, self.phut_chay)
            self.cap_nhat_nhan_lan_sau()

        def cap_nhat_nhan_lan_sau(self) -> None:
            self.lb_lan_sau.config(
                text=f"Lần cập nhật tự động tiếp theo: "
                     f"{self.lan_tiep_theo.strftime('%H:%M ngày %d/%m/%Y')}")

        # ------------------------------------------------------------------
        def ghi(self, text: str) -> None:
            self.khung_log.configure(state="normal")
            self.khung_log.insert("end", text + "\n")
            self.khung_log.see("end")
            self.khung_log.configure(state="disabled")

        def bam_cap_nhat(self) -> None:
            if not self.dang_chay:
                self.bat_dau_cap_nhat()

        def bat_dau_cap_nhat(self) -> None:
            if self.dang_chay:
                return
            self.dang_chay = True
            self.nut_chay.state(["disabled"])
            self.lb_trang_thai.config(text="ĐANG CHẠY CẬP NHẬT ...", fg=CAM)
            self.ghi("=" * 66)
            self.ghi(f"ĐANG CHẠY CẬP NHẬT ... ({datetime.now().strftime('%d/%m/%Y %H:%M:%S')})")
            threading.Thread(target=self._worker, daemon=True).start()

        def _worker(self) -> None:
            try:
                code = chay_cap_nhat(lambda line: self.hang_doi.put(("line", line)))
            except Exception as e:
                self.hang_doi.put(("line", f"LỖI: {e}"))
                code = 1
            self.hang_doi.put(("done", code))

        def vong_lap(self) -> None:
            # Nhận output từ tiến trình con
            try:
                while True:
                    loai, gia_tri = self.hang_doi.get_nowait()
                    if loai == "line":
                        self.ghi("  " + str(gia_tri))
                    else:
                        self._xong(int(gia_tri))
            except queue.Empty:
                pass

            # Đồng hồ đếm ngược + tự chạy đúng giờ
            now = datetime.now()
            if not self.dang_chay:
                if now >= self.lan_tiep_theo:
                    self.bat_dau_cap_nhat()
                else:
                    con_lai = self.lan_tiep_theo - now
                    h, du = divmod(int(con_lai.total_seconds()), 3600)
                    m, s = divmod(du, 60)
                    self.lb_dong_ho.config(text=f"{h:02d}:{m:02d}:{s:02d}")
            self.root.after(250, self.vong_lap)

        def _xong(self, code: int) -> None:
            self.dang_chay = False
            self.nut_chay.state(["!disabled"])
            self.lan_tiep_theo = lan_chay_tiep_theo(datetime.now(),
                                                    self.gio_chay, self.phut_chay)
            gio = datetime.now().strftime("%H:%M:%S")
            if code == 0:
                self.ghi(f"HOÀN TẤT lúc {gio}\n")
                self.lb_trang_thai.config(
                    text=f"Cập nhật thành công lúc {gio} - đang đếm ngược lần tiếp theo",
                    fg=XANH_DAM)
            else:
                self.ghi(f"CÓ LỖI (mã {code}) lúc {gio} - xem chi tiết phía trên\n")
                self.lb_trang_thai.config(
                    text=f"Lần chạy {gio} bị lỗi - sẽ thử lại theo lịch", fg=DO)
            self.cap_nhat_nhan_lan_sau()

        def mo_sheet(self) -> None:
            import config
            if config.GOOGLE_SHEET_ID:
                webbrowser.open(f"https://docs.google.com/spreadsheets/d/{config.GOOGLE_SHEET_ID}")

        def mo_log(self) -> None:
            if LOG_FILE.exists():
                os.startfile(str(LOG_FILE))

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    if "--ngam" in sys.argv:
        chay_ngam()
    else:
        chay_giao_dien()
