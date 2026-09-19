# -*- coding: utf-8 -*-
"""Ứng dụng desktop CẬP NHẬT THƯỞNG GREENVITA (Pancake POS -> Google Sheet).

- Nháy đúp file này -> mở cửa sổ hiển thị tiến trình và ĐẾM NGƯỢC tới lần chạy
  gần nhất rồi tự chạy. Có 2 lịch:
    + 9h sáng hằng ngày (chỉnh được trên giao diện): cập nhật tháng hiện tại;
    + 11h00 mùng 2 hằng tháng (config.CHOT_SO_NGAY / CHOT_SO_GIO): CHỐT SỔ tháng
      trước - thuong_thang.py chỉ chốt từ mốc này, chốt xong tab bị khóa, không sửa nữa.
- CHẠY NỀN GIỐNG UNIKEY: app có icon ở khay hệ thống (góc phải taskbar). Bấm X đóng
  cửa sổ thì app chỉ ẨN xuống khay, đồng hồ đếm ngược vẫn chạy và vẫn tự cập nhật
  đúng giờ. Nháy vào icon khay (hoặc nháy đúp lối tắt trên Desktop lần nữa) thì cửa sổ
  hiện lại; chuột phải icon khay có menu Mở cửa sổ / Cập nhật ngay / Mở log / Thoát.
  Chỉ chạy được 1 phiên bản: mở lần 2 chỉ hiện lại cửa sổ đang chạy, không mở trùng.
- Khởi động cùng Windows: tick "Khởi động cùng Windows" trên giao diện (hoặc chạy
  python tao_loi_tat.py) -> tạo lối tắt trong thư mục Startup chạy
  pythonw app_cap_nhat.pyw --khay : đăng nhập máy là app tự chạy ẩn ở khay.
- Nút bấm: Cập nhật ngay / Sheet Sale / Sheet CSKH / Mở file log;
  ô "Link lịch trực CN", "Link chấm công 1" (file NT/TK) và "Link chấm công 2"
  (file OCP - không có tên ở bảng 1 thì tìm ở đây) đều có nút Lưu link (ghi .env)
  và nút Mở (mở trang tính trên trình duyệt).
- Chế độ chạy ngầm cho Task Scheduler:  pythonw app_cap_nhat.pyw --ngam
  (chạy 1 lần, ghi log rồi thoát - không mở cửa sổ, không dính tới icon khay).

Cần thư viện pystray + Pillow cho icon khay (pip install -r requirements.txt); thiếu thì
app vẫn chạy bình thường nhưng bấm X là thoát hẳn.
Mọi lần chạy đều ghi thêm vào logs/cap_nhat.log.
"""
import json
import os
import queue
import re
import socket
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
ENV_FILE = BASE_DIR / ".env"
ICO_FILE = BASE_DIR / "icon_app.ico"
GIO_CHAY_MAC_DINH = (9, 0)       # giờ:phút tự chạy hằng ngày (mặc định 9h sáng)
CHOT_SO_MAC_DINH = (2, 11, 0)    # (mùng, giờ, phút) chốt sổ tháng trước - dự phòng khi
                                 # không đọc được config.py
TEN_APP = "Cập Nhật Thưởng GreenVita"
# Cổng localhost để bảo đảm CHỈ 1 PHIÊN BẢN app chạy: phiên bản đầu giữ cổng này; mở
# lần 2 thấy cổng bận thì gửi chữ "hien" tới rồi thoát -> phiên bản đầu hiện cửa sổ lên.
CONG_MOT_PHIEN = 47321


def doc_env(key: str) -> str:
    """Đọc 1 biến từ file .env (không cần load_dotenv)."""
    try:
        for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith(key + "="):
                return line.split("=", 1)[1].strip()
    except OSError:
        pass
    return ""


def luu_env(key: str, value: str) -> None:
    """Ghi/cập nhật 1 biến trong file .env."""
    text = ENV_FILE.read_text(encoding="utf-8") if ENV_FILE.exists() else ""
    if re.search(rf"^{key}=.*$", text, flags=re.M):
        text = re.sub(rf"^{key}=.*$", f"{key}={value}", text, flags=re.M)
    else:
        text += f"\n{key}={value}\n"
    ENV_FILE.write_text(text, encoding="utf-8")


def rut_id_sheet(s: str) -> str:
    """Nhận cả link đầy đủ lẫn ID trần: trả về ID trang tính."""
    m = re.search(r"/d/([A-Za-z0-9_-]{20,})", s or "")
    return m.group(1) if m else (s or "").strip()


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


def moc_chot_so() -> tuple[int, int, int]:
    """(mùng, giờ, phút) tự chốt sổ tháng trước - đọc config.py (CHOT_SO_NGAY / CHOT_SO_GIO),
    cùng mốc thuong_thang.py dùng để quyết định có chốt hay không."""
    try:
        import config
        return int(config.CHOT_SO_NGAY), int(config.CHOT_SO_GIO[0]), int(config.CHOT_SO_GIO[1])
    except Exception:
        return CHOT_SO_MAC_DINH


def lan_chot_so_tiep_theo(now: datetime) -> datetime:
    """Lần CHỐT SỔ tự động tiếp theo: 11h00 mùng 2 gần nhất còn ở sau `now`."""
    ngay, gio, phut = moc_chot_so()
    target = now.replace(day=ngay, hour=gio, minute=phut, second=0, microsecond=0)
    if target <= now:
        target = (now.replace(day=1) + timedelta(days=32)).replace(
            day=ngay, hour=gio, minute=phut, second=0, microsecond=0)
    return target


def thang_chot_so(lan_chot: datetime) -> str:
    """Tháng sẽ được chốt ở lần chạy `lan_chot` (= tháng liền trước), dạng 09.2026."""
    truoc = lan_chot.replace(day=1) - timedelta(days=1)
    return truoc.strftime("%m.%Y")


def chay_cap_nhat(on_line, on_start=None) -> int:
    """Chạy thuong_thang.py, gọi on_line(dòng) cho từng dòng, ghi log. Trả về mã lỗi.

    on_start(proc): được gọi ngay khi tiến trình con bắt đầu (để GUI giữ tham chiếu,
    dừng được khi người dùng Thoát giữa chừng)."""
    LOG_DIR.mkdir(exist_ok=True)
    proc = subprocess.Popen(
        [sys.executable, str(BASE_DIR / "thuong_thang.py")],
        cwd=str(BASE_DIR), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if on_start:
        on_start(proc)
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
# Chỉ 1 phiên bản app (single instance) qua cổng localhost
# ======================================================================
def giu_cong_mot_phien() -> "socket.socket | None":
    """Thử giữ cổng CONG_MOT_PHIEN. Được -> trả về socket đang lắng nghe (app là phiên
    bản duy nhất). Không được (app khác đang giữ) -> trả về None."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", CONG_MOT_PHIEN))
        s.listen(5)
        return s
    except OSError:
        s.close()
        return None


def bao_phien_dang_chay_hien_len() -> bool:
    """Gửi "hien" tới phiên bản đang chạy để nó hiện cửa sổ. Trả về True nếu gửi được."""
    try:
        with socket.create_connection(("127.0.0.1", CONG_MOT_PHIEN), timeout=2) as c:
            c.sendall(b"hien")
        return True
    except OSError:
        return False


# ======================================================================
# Giao diện
# ======================================================================
def chay_giao_dien(an_o_khay: bool = False) -> None:
    """an_o_khay=True (tham số --khay, dùng khi khởi động cùng Windows): mở app nhưng
    KHÔNG hiện cửa sổ, chỉ có icon ở khay; nháy icon thì cửa sổ mới hiện."""
    import tkinter as tk
    from tkinter import messagebox, scrolledtext, ttk

    class App:
        def __init__(self, root: tk.Tk, cong: "socket.socket | None", an_o_khay: bool):
            self.root = root
            self.dang_chay = False
            self.tien_trinh: subprocess.Popen | None = None   # thuong_thang.py đang chạy
            self.hang_doi: queue.Queue = queue.Queue()
            self.gio_chay, self.phut_chay = doc_gio_chay()
            self.bat_dau_luc = datetime.now()      # lúc bắt đầu lần chạy gần nhất
            self.tinh_lan_tiep_theo(datetime.now())
            self.khay = None                       # pystray.Icon (None = không có khay)
            self.da_bao_an_khay = False            # đã hiện thông báo "app ẩn ở khay" chưa
            self.cong = cong

            root.title(TEN_APP)
            try:
                root.iconbitmap(str(ICO_FILE))
            except Exception:
                pass
            # Hiện cửa sổ ở CHÍNH GIỮA màn hình
            w, h = 720, 595
            x = (root.winfo_screenwidth() - w) // 2
            y = (root.winfo_screenheight() - h) // 2
            root.geometry(f"{w}x{h}+{x}+{y}")
            root.minsize(600, 495)
            root.configure(bg=XANH_NHAT)
            # Bấm X: ẩn xuống khay (giống Unikey) thay vì thoát
            root.protocol("WM_DELETE_WINDOW", self.an_xuong_khay)

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
            # Tick "Khởi động cùng Windows" = có lối tắt --khay trong thư mục Startup
            self.bien_khoi_dong = tk.BooleanVar(value=self.co_khoi_dong_cung_windows())
            tk.Checkbutton(khung_gio, text="Khởi động cùng Windows (chạy ẩn ở khay)",
                           variable=self.bien_khoi_dong, command=self.doi_khoi_dong,
                           font=("Segoe UI", 10), fg=XAM, bg=XANH_NHAT,
                           activebackground=XANH_NHAT, selectcolor="white"
                           ).grid(row=0, column=4, padx=(18, 0))

            # --- Link lịch trực Chủ nhật + 2 link chấm công (sửa được, lưu vào .env) ---
            self.bien_lich_truc = self.tao_hang_link(
                root, "Link lịch trực CN:", "GOOGLE_SHEET_ID_LICH_TRUC",
                " Mở lịch trực ", "lịch trực")
            self.bien_cham_cong = self.tao_hang_link(
                root, "Link chấm công 1:", "GOOGLE_SHEET_ID_CHAM_CONG",
                " Mở chấm công 1 ", "chấm công 1")
            self.bien_cham_cong_2 = self.tao_hang_link(
                root, "Link chấm công 2:", "GOOGLE_SHEET_ID_CHAM_CONG_2",
                " Mở chấm công 2 ", "chấm công 2")

            # --- Nút ---
            khung_nut = tk.Frame(root, bg=XANH_NHAT)
            khung_nut.pack(pady=4)
            style = ttk.Style()
            style.configure("TButton", font=("Segoe UI", 10), padding=6)
            self.nut_chay = ttk.Button(khung_nut, text="  Cập nhật ngay  ",
                                       command=self.bam_cap_nhat)
            self.nut_chay.grid(row=0, column=0, padx=6)
            ttk.Button(khung_nut, text="  Sheet Sale  ",
                       command=lambda: self.mo_sheet("sale")).grid(row=0, column=1, padx=6)
            ttk.Button(khung_nut, text="  Sheet CSKH  ",
                       command=lambda: self.mo_sheet("cskh")).grid(row=0, column=2, padx=6)
            ttk.Button(khung_nut, text="  Mở file log  ",
                       command=self.mo_log).grid(row=0, column=3, padx=6)

            # --- Khung tiến trình ---
            self.khung_log = scrolledtext.ScrolledText(
                root, font=("Consolas", 10), state="disabled", wrap="word",
                bg="white", relief="flat", borderwidth=6)
            self.khung_log.pack(fill="both", expand=True, padx=12, pady=(6, 12))
            # Màu dòng thông báo lịch trực / chấm công: lỗi đỏ / kết nối OK xanh
            self.khung_log.tag_configure("do", foreground=DO)
            self.khung_log.tag_configure("xanh", foreground="#1E8E3E")

            self.ghi("Chào mừng! Bấm 'Cập nhật ngay' để chạy thủ công,")
            self.ghi("hoặc chờ đồng hồ đếm ngược - đến giờ hẹn ứng dụng tự cập nhật.")
            ngay, gio, phut = moc_chot_so()
            self.ghi(f"Chốt sổ tháng trước tự chạy lúc {gio:02d}:{phut:02d} mùng {ngay} "
                     f"hằng tháng - chốt xong tab tháng đó bị khóa, không sửa nữa.")
            self.cap_nhat_nhan_lan_sau()

            self.tao_khay()
            if self.khay is not None:
                self.ghi("Bấm X chỉ ẨN cửa sổ xuống khay hệ thống (góc phải taskbar) - app "
                         "vẫn chạy nền và tự cập nhật đúng giờ. Nháy icon khay để mở lại; "
                         "chuột phải icon -> Thoát để tắt hẳn.\n")
            else:
                self.ghi("(Chưa cài pystray/Pillow nên không có icon khay - bấm X là thoát "
                         "hẳn. Cài: pip install pystray pillow)\n")
            if self.cong is not None:
                threading.Thread(target=self._lang_nghe_cong, daemon=True).start()
            if an_o_khay and self.khay is None:
                # Không có khay mà lại ẩn thì không cách nào mở lại -> cứ hiện cửa sổ
                self.root.after(0, self.hien_cua_so)

            self.root.after(200, self.vong_lap)

        # ------------------------------------------------------------------
        # Khay hệ thống (system tray) + ẩn/hiện cửa sổ + thoát
        # ------------------------------------------------------------------
        def tao_khay(self) -> None:
            """Tạo icon ở khay hệ thống bằng pystray, chạy ở luồng riêng.

            Menu của pystray chạy ở luồng khác, KHÔNG được đụng Tk trực tiếp -> mọi
            hành động chỉ đẩy vào hang_doi, vong_lap (luồng Tk) sẽ xử lý."""
            try:
                import pystray
                from PIL import Image
            except ImportError:
                return
            try:
                anh = Image.open(ICO_FILE)
                anh.load()
            except Exception:
                from PIL import Image as _Im, ImageDraw
                anh = _Im.new("RGBA", (64, 64), (0, 0, 0, 0))
                ImageDraw.Draw(anh).rounded_rectangle((2, 2, 61, 61), radius=14,
                                                      fill=(31, 122, 58))

            def gui(loai):
                return lambda: self.hang_doi.put((loai, None))

            menu = pystray.Menu(
                pystray.MenuItem("Mở cửa sổ", gui("hien"), default=True),
                pystray.MenuItem("Cập nhật ngay", gui("chay")),
                pystray.MenuItem("Mở file log", gui("log")),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Thoát", gui("thoat")),
            )
            self.khay = pystray.Icon("GreenVita", anh, TEN_APP, menu)
            threading.Thread(target=self.khay.run, daemon=True).start()

        def dat_ten_khay(self, text: str) -> None:
            """Tooltip khi rê chuột lên icon khay."""
            if self.khay is not None:
                try:
                    self.khay.title = f"{TEN_APP}\n{text}"[:127]
                except Exception:
                    pass

        def an_xuong_khay(self) -> None:
            """Bấm X: ẩn cửa sổ, app vẫn chạy ở khay. Không có khay thì thoát hẳn."""
            if self.khay is None:
                self.thoat()
                return
            self.root.withdraw()
            if not self.da_bao_an_khay:
                self.da_bao_an_khay = True
                try:
                    self.khay.notify("Ứng dụng vẫn chạy ngầm ở khay hệ thống và sẽ tự cập "
                                     "nhật đúng giờ. Nháy icon để mở lại cửa sổ.", TEN_APP)
                except Exception:
                    pass

        def hien_cua_so(self) -> None:
            """Hiện lại cửa sổ (từ khay, hoặc khi mở app lần 2) và đưa lên trên cùng."""
            self.root.deiconify()
            self.root.state("normal")
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(150, lambda: self.root.attributes("-topmost", False))
            self.root.focus_force()

        def thoat(self) -> bool:
            """Thoát hẳn (menu khay). Đang cập nhật dở thì hỏi lại rồi dừng tiến trình con.
            Trả về True nếu đã thoát thật, False nếu người dùng đổi ý."""
            if self.dang_chay:
                self.hien_cua_so()
                if not messagebox.askyesno(
                        "Thoát", "Đang cập nhật dở, thoát bây giờ sẽ dừng ngang lần cập "
                        "nhật này (lần sau chạy lại từ đầu, không mất dữ liệu).\n\nVẫn thoát?",
                        parent=self.root):
                    return False
                if self.tien_trinh is not None and self.tien_trinh.poll() is None:
                    try:
                        self.tien_trinh.kill()
                    except OSError:
                        pass
            if self.khay is not None:
                try:
                    self.khay.stop()
                except Exception:
                    pass
            if self.cong is not None:
                try:
                    self.cong.close()
                except OSError:
                    pass
            self.root.destroy()
            return True

        def _lang_nghe_cong(self) -> None:
            """Luồng nền: ai kết nối tới cổng (mở app lần 2) -> yêu cầu hiện cửa sổ."""
            while True:
                try:
                    conn, _ = self.cong.accept()
                except OSError:
                    return                       # cổng đã đóng (app thoát)
                with conn:
                    self.hang_doi.put(("hien", None))

        # ------------------------------------------------------------------
        # Khởi động cùng Windows
        # ------------------------------------------------------------------
        @staticmethod
        def co_khoi_dong_cung_windows() -> bool:
            try:
                import tao_loi_tat
                return tao_loi_tat.co_khoi_dong_cung_windows()
            except Exception:
                return False

        def doi_khoi_dong(self) -> None:
            """Tick / bỏ tick 'Khởi động cùng Windows': tạo / xóa lối tắt trong Startup."""
            try:
                import tao_loi_tat
                if self.bien_khoi_dong.get():
                    lnk = tao_loi_tat.bat_khoi_dong_cung_windows()
                    self.ghi(f"Đã bật khởi động cùng Windows: đăng nhập máy là app tự chạy "
                             f"ẩn ở khay. (Lối tắt: {lnk})")
                else:
                    tao_loi_tat.tat_khoi_dong_cung_windows()
                    self.ghi("Đã tắt khởi động cùng Windows.")
            except Exception as e:
                self.ghi(f"[LOI] Không đổi được chế độ khởi động cùng Windows: {e}")
                self.bien_khoi_dong.set(self.co_khoi_dong_cung_windows())

        # ------------------------------------------------------------------
        def doi_gio(self) -> None:
            """Người dùng chỉnh giờ tự chạy trên giao diện."""
            try:
                self.gio_chay = int(self.bien_gio.get()) % 24
                self.phut_chay = int(self.bien_phut.get()) % 60
            except ValueError:
                return
            luu_gio_chay(self.gio_chay, self.phut_chay)
            self.tinh_lan_tiep_theo(datetime.now())
            self.cap_nhat_nhan_lan_sau()

        def tinh_lan_tiep_theo(self, now: datetime, bat_dau: datetime | None = None) -> None:
            """2 lịch tự chạy; đồng hồ đếm ngược tới lịch nào gần hơn.

            Lịch chốt sổ tính từ `bat_dau` (lúc BẮT ĐẦU lần chạy vừa rồi): lần chạy tay
            bắt đầu 10h58 kéo dài qua 11h00 mùng 2 thì vẫn còn lượt chốt sổ 11h00 (lần
            chạy đó chưa đủ giờ nên thuong_thang.py chưa chốt)."""
            self.lan_hang_ngay = lan_chay_tiep_theo(now, self.gio_chay, self.phut_chay)
            self.lan_chot_so = lan_chot_so_tiep_theo(bat_dau or now)
            self.lan_tiep_theo = min(self.lan_hang_ngay, self.lan_chot_so)

        def cap_nhat_nhan_lan_sau(self) -> None:
            self.lb_lan_sau.config(
                text=f"Cập nhật hằng ngày tiếp theo: "
                     f"{self.lan_hang_ngay.strftime('%H:%M ngày %d/%m/%Y')}   |   "
                     f"Chốt sổ tháng {thang_chot_so(self.lan_chot_so)}: "
                     f"{self.lan_chot_so.strftime('%H:%M ngày %d/%m/%Y')}")

        # ------------------------------------------------------------------
        def ghi(self, text: str) -> None:
            tag = ()
            if "[LICH TRUC][LOI]" in text or "[CHAM CONG][LOI]" in text:
                tag = ("do",)
            elif ("[LICH TRUC][OK]" in text or "[CHAM CONG][OK]" in text
                  or "Đã lưu link" in text):
                tag = ("xanh",)
            self.khung_log.configure(state="normal")
            self.khung_log.insert("end", text + "\n", tag)
            self.khung_log.see("end")
            self.khung_log.configure(state="disabled")

        def tao_hang_link(self, root, nhan: str, env_key: str,
                          ten_nut_mo: str, ten: str):
            """Tạo 1 hàng: nhãn + ô nhập link/ID + nút Lưu link + nút Mở.

            Giá trị đọc từ .env (env_key); trả về StringVar gắn với ô nhập.
            """
            khung = tk.Frame(root, bg=XANH_NHAT)
            khung.pack(pady=(2, 2))
            tk.Label(khung, text=nhan, font=("Segoe UI", 10), fg=XAM, bg=XANH_NHAT,
                     width=17, anchor="e").grid(row=0, column=0, padx=(0, 6))
            bien = tk.StringVar(value=doc_env(env_key))
            tk.Entry(khung, textvariable=bien, width=46,
                     font=("Segoe UI", 9)).grid(row=0, column=1, padx=(0, 6))
            ttk.Button(khung, text=" Lưu link ", width=10,
                       command=lambda: self.luu_link(bien, env_key, ten)
                       ).grid(row=0, column=2)
            ttk.Button(khung, text=ten_nut_mo, width=15,
                       command=lambda: self.mo_link(bien, env_key, ten)
                       ).grid(row=0, column=3, padx=(6, 0))
            return bien

        def luu_link(self, bien, env_key: str, ten: str) -> None:
            """Lưu link/ID trong ô nhập vào .env (nhận cả link đầy đủ lẫn ID trần)."""
            sheet_id = rut_id_sheet(bien.get())
            if not sheet_id:
                self.ghi(f"Chưa nhập link/ID {ten}.")
                return
            luu_env(env_key, sheet_id)
            bien.set(sheet_id)
            self.ghi(f"Đã lưu link {ten} (ID: {sheet_id}) - áp dụng từ lần cập nhật sau.")

        def mo_link(self, bien, env_key: str, ten: str) -> None:
            """Mở trang tính đang dùng trên trình duyệt (link trong ô bên cạnh)."""
            sheet_id = rut_id_sheet(bien.get()) or doc_env(env_key)
            if not sheet_id:
                self.ghi(f"Chưa có link/ID {ten} để mở.")
                return
            webbrowser.open(f"https://docs.google.com/spreadsheets/d/{sheet_id}")
            self.ghi(f"Đã mở {ten} (ID: {sheet_id}) trên trình duyệt.")

        def bam_cap_nhat(self) -> None:
            if not self.dang_chay:
                self.bat_dau_cap_nhat()

        def bat_dau_cap_nhat(self, chot_so: bool = False) -> None:
            """chot_so: lượt chạy theo lịch 11h mùng 2 (thuong_thang.py sẽ chốt sổ tháng trước)."""
            if self.dang_chay:
                return
            self.dang_chay = True
            self.bat_dau_luc = datetime.now()
            self.nut_chay.state(["disabled"])
            viec = (f"CHỐT SỔ THÁNG {thang_chot_so(self.lan_chot_so)}" if chot_so
                    else "CẬP NHẬT")
            self.lb_trang_thai.config(text=f"ĐANG CHẠY {viec} ...", fg=CAM)
            self.dat_ten_khay(f"Đang chạy {viec.lower()} ...")
            self.ghi("=" * 66)
            self.ghi(f"ĐANG CHẠY {viec} ... ({self.bat_dau_luc.strftime('%d/%m/%Y %H:%M:%S')})")
            threading.Thread(target=self._worker, daemon=True).start()

        def _worker(self) -> None:
            def _giu(proc):
                self.tien_trinh = proc
            try:
                code = chay_cap_nhat(lambda line: self.hang_doi.put(("line", line)), _giu)
            except Exception as e:
                self.hang_doi.put(("line", f"LỖI: {e}"))
                code = 1
            self.tien_trinh = None
            self.hang_doi.put(("done", code))

        def vong_lap(self) -> None:
            # Nhận output từ tiến trình con + lệnh từ icon khay / phiên bản mở lần 2
            try:
                while True:
                    loai, gia_tri = self.hang_doi.get_nowait()
                    if loai == "line":
                        self.ghi("  " + str(gia_tri))
                    elif loai == "done":
                        self._xong(int(gia_tri))
                    elif loai == "hien":
                        self.hien_cua_so()
                    elif loai == "chay":
                        self.bam_cap_nhat()
                    elif loai == "log":
                        self.mo_log()
                    elif loai == "thoat":
                        if self.thoat():
                            return           # root đã destroy, không đặt after nữa
            except queue.Empty:
                pass

            # Đồng hồ đếm ngược + tự chạy đúng giờ
            now = datetime.now()
            if not self.dang_chay:
                if now >= self.lan_tiep_theo:
                    self.bat_dau_cap_nhat(chot_so=now >= self.lan_chot_so)
                else:
                    con_lai = self.lan_tiep_theo - now
                    h, du = divmod(int(con_lai.total_seconds()), 3600)
                    m, s = divmod(du, 60)
                    self.lb_dong_ho.config(text=f"{h:02d}:{m:02d}:{s:02d}")
            self.root.after(250, self.vong_lap)

        def _xong(self, code: int) -> None:
            self.dang_chay = False
            self.nut_chay.state(["!disabled"])
            self.tinh_lan_tiep_theo(datetime.now(), bat_dau=self.bat_dau_luc)
            gio = datetime.now().strftime("%H:%M:%S")
            if code == 0:
                self.ghi(f"HOÀN TẤT lúc {gio}\n")
                self.lb_trang_thai.config(
                    text=f"Cập nhật thành công lúc {gio} - đang đếm ngược lần tiếp theo",
                    fg=XANH_DAM)
                self.dat_ten_khay(f"Cập nhật OK lúc {gio} - lần tiếp: "
                                  f"{self.lan_tiep_theo.strftime('%H:%M %d/%m')}")
            else:
                self.ghi(f"CÓ LỖI (mã {code}) lúc {gio} - xem chi tiết phía trên\n")
                self.lb_trang_thai.config(
                    text=f"Lần chạy {gio} bị lỗi - sẽ thử lại theo lịch", fg=DO)
                self.dat_ten_khay(f"LỖI lúc {gio} - lần tiếp: "
                                  f"{self.lan_tiep_theo.strftime('%H:%M %d/%m')}")
            self.cap_nhat_nhan_lan_sau()

        def mo_sheet(self, nhom: str = "sale") -> None:
            import config
            sheet_id = (config.GOOGLE_SHEET_ID_CSKH if nhom == "cskh"
                        else config.GOOGLE_SHEET_ID)
            if sheet_id:
                webbrowser.open(f"https://docs.google.com/spreadsheets/d/{sheet_id}")

        def mo_log(self) -> None:
            if LOG_FILE.exists():
                os.startfile(str(LOG_FILE))

    # Chỉ 1 phiên bản: app khác đang chạy (ở khay) -> bảo nó hiện cửa sổ rồi thoát
    cong = giu_cong_mot_phien()
    if cong is None:
        if bao_phien_dang_chay_hien_len():
            return
        # Cổng bận vì chương trình khác - vẫn chạy, chỉ mất tính năng "1 phiên bản"

    root = tk.Tk()
    if an_o_khay:
        root.withdraw()                # ẩn ngay từ đầu, không lóe cửa sổ
    app = App(root, cong, an_o_khay)
    root.mainloop()
    # Thoát mainloop nhưng icon khay còn (vd Tk bị đóng bất thường) -> dọn cho sạch
    if app.khay is not None:
        try:
            app.khay.stop()
        except Exception:
            pass


if __name__ == "__main__":
    if "--ngam" in sys.argv:
        chay_ngam()
    else:
        chay_giao_dien(an_o_khay="--khay" in sys.argv)
