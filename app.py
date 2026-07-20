# -*- coding: utf-8 -*-
"""
نرم‌افزار نوبت‌دهی قابل تنظیم
================================
- سربرگ (متن بالای فیش) قابل تغییر
- شماره‌ی شروع و پایان قابل تنظیم
- اندازه‌ی فونت سربرگ و شماره قابل تنظیم
- تنظیمات در فایل config.json ذخیره و در اجراهای بعدی بازیابی می‌شود
- چاپ روی فیش‌پرینترهای حرارتی (ESC/POS) از طریق win32print (ویندوز)

اجرا:
    pip install -r requirements.txt
    python app.py
"""

import json
import os
from datetime import date
import tkinter as tk
from tkinter import messagebox, ttk

try:
    import win32print
    WINDOWS_PRINTING_AVAILABLE = True
except ImportError:
    WINDOWS_PRINTING_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
STATE_FILE = os.path.join(BASE_DIR, "state.json")

DEFAULT_CONFIG = {
    "header_text": "سیستم نوبت‌دهی",
    "printer_name": "Meva TP1000",
    "start_number": 1,
    "end_number": 999,
    "header_font_size": 1,     # 0=عادی، 1=دو برابر، 2=سه برابر ...
    "number_font_size": 3,     # اندازه‌ی چاپ شماره (بزرگ‌تر = درشت‌تر)
    "show_date": True,
}


def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                merged = dict(default)
                merged.update(data)
                return merged
        except Exception:
            pass
    return dict(default)


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def list_printers():
    if not WINDOWS_PRINTING_AVAILABLE:
        return []
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    return [p[2] for p in printers]


# نگاشت اندازه‌ی فونت به کد ESC/POS (GS ! n)
# n: بیت‌های 0-3 = عرض (0..7)، بیت‌های 4-7 = ارتفاع (0..7)
def font_size_byte(size_level: int) -> bytes:
    size_level = max(0, min(7, size_level))
    n = (size_level << 4) | size_level  # عرض و ارتفاع برابر
    return bytes([n])


def build_escpos_bytes(number: int, cfg: dict) -> bytes:
    ESC = b"\x1b"
    GS = b"\x1d"

    data = b""
    data += ESC + b"@"                      # ریست پرینتر
    data += ESC + b"a" + b"\x01"            # وسط‌چین

    # ---- سربرگ ----
    data += GS + b"!" + font_size_byte(cfg["header_font_size"])
    for line in cfg["header_text"].splitlines():
        data += (line + "\n").encode("cp1256", errors="ignore")
    data += GS + b"!" + b"\x00"             # برگشت به فونت عادی
    data += b"--------------------------------\n"

    # ---- شماره ----
    data += GS + b"!" + font_size_byte(cfg["number_font_size"])
    data += (str(number) + "\n").encode("cp1256", errors="ignore")
    data += GS + b"!" + b"\x00"

    data += b"--------------------------------\n"
    if cfg.get("show_date", True):
        data += (date.today().strftime("%Y-%m-%d") + "\n").encode("cp1256", errors="ignore")

    data += b"\n\n\n"
    data += GS + b"V" + b"\x41" + b"\x10"   # برش کاغذ
    return data


def print_ticket(number: int, cfg: dict) -> bool:
    if not WINDOWS_PRINTING_AVAILABLE:
        messagebox.showwarning(
            "چاپ در دسترس نیست",
            "چاپ فقط روی ویندوز و با نصب pywin32 کار می‌کند.\n"
            "دستور نصب: pip install pywin32",
        )
        return False
    try:
        payload = build_escpos_bytes(number, cfg)
        hprinter = win32print.OpenPrinter(cfg["printer_name"])
        try:
            win32print.StartDocPrinter(hprinter, 1, ("Nobat Ticket", None, "RAW"))
            win32print.StartPagePrinter(hprinter)
            win32print.WritePrinter(hprinter, payload)
            win32print.EndPagePrinter(hprinter)
            win32print.EndDocPrinter(hprinter)
        finally:
            win32print.ClosePrinter(hprinter)
        return True
    except Exception as e:
        messagebox.showerror(
            "خطا در چاپ",
            f"چاپ روی پرینتر «{cfg['printer_name']}» ناموفق بود:\n{e}\n\n"
            "پرینترهای موجود:\n" + "\n".join(list_printers()),
        )
        return False


class SettingsDialog(tk.Toplevel):
    """پنجره‌ی تنظیمات: سربرگ، بازه‌ی شماره، اندازه‌ی فونت‌ها، نام پرینتر"""

    def __init__(self, master, cfg: dict, on_save):
        super().__init__(master)
        self.title("تنظیمات")
        self.geometry("420x520")
        self.cfg = dict(cfg)
        self.on_save = on_save
        self.resizable(False, False)

        pad = {"padx": 15, "pady": 6}

        tk.Label(self, text="متن سربرگ (چند خط مجاز است):", font=("Tahoma", 10, "bold")).pack(anchor="e", **pad)
        self.header_text = tk.Text(self, height=4, font=("Tahoma", 11))
        self.header_text.insert("1.0", self.cfg.get("header_text", ""))
        self.header_text.pack(fill="x", padx=15)

        frame_range = tk.Frame(self)
        frame_range.pack(fill="x", **pad)
        tk.Label(frame_range, text="شماره شروع:", font=("Tahoma", 10)).grid(row=0, column=1, sticky="e", padx=5)
        self.start_var = tk.IntVar(value=self.cfg.get("start_number", 1))
        tk.Entry(frame_range, textvariable=self.start_var, width=10, justify="center").grid(row=0, column=0, padx=5)

        tk.Label(frame_range, text="شماره پایان:", font=("Tahoma", 10)).grid(row=1, column=1, sticky="e", padx=5, pady=5)
        self.end_var = tk.IntVar(value=self.cfg.get("end_number", 999))
        tk.Entry(frame_range, textvariable=self.end_var, width=10, justify="center").grid(row=1, column=0, padx=5)

        tk.Label(self, text="اندازه‌ی فونت سربرگ (۰ تا ۷):", font=("Tahoma", 10)).pack(anchor="e", **pad)
        self.header_size_var = tk.IntVar(value=self.cfg.get("header_font_size", 1))
        tk.Spinbox(self, from_=0, to=7, textvariable=self.header_size_var, width=8, justify="center").pack(anchor="e", padx=15)

        tk.Label(self, text="اندازه‌ی فونت شماره (۰ تا ۷):", font=("Tahoma", 10)).pack(anchor="e", **pad)
        self.number_size_var = tk.IntVar(value=self.cfg.get("number_font_size", 3))
        tk.Spinbox(self, from_=0, to=7, textvariable=self.number_size_var, width=8, justify="center").pack(anchor="e", padx=15)

        self.show_date_var = tk.BooleanVar(value=self.cfg.get("show_date", True))
        tk.Checkbutton(self, text="نمایش تاریخ روی فیش", variable=self.show_date_var,
                        font=("Tahoma", 10)).pack(anchor="e", padx=15, pady=6)

        tk.Label(self, text="نام پرینتر (دقیقاً مطابق ویندوز):", font=("Tahoma", 10, "bold")).pack(anchor="e", **pad)
        self.printer_var = tk.StringVar(value=self.cfg.get("printer_name", "Meva TP1000"))
        printers = list_printers()
        if printers:
            combo = ttk.Combobox(self, textvariable=self.printer_var, values=printers, justify="center")
            combo.pack(fill="x", padx=15)
        else:
            tk.Entry(self, textvariable=self.printer_var, justify="center").pack(fill="x", padx=15)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="ذخیره", font=("Tahoma", 11, "bold"), bg="#2e7d32", fg="white",
                  padx=20, pady=8, command=self.save).pack(side="right", padx=10)
        tk.Button(btn_frame, text="انصراف", font=("Tahoma", 11), padx=20, pady=8,
                  command=self.destroy).pack(side="right")

    def save(self):
        try:
            start = int(self.start_var.get())
            end = int(self.end_var.get())
        except Exception:
            messagebox.showerror("خطا", "شماره شروع و پایان باید عدد صحیح باشند.")
            return
        if end < start:
            messagebox.showerror("خطا", "شماره‌ی پایان باید بزرگ‌تر یا مساوی شماره‌ی شروع باشد.")
            return

        self.cfg["header_text"] = self.header_text.get("1.0", "end").strip()
        self.cfg["start_number"] = start
        self.cfg["end_number"] = end
        self.cfg["header_font_size"] = int(self.header_size_var.get())
        self.cfg["number_font_size"] = int(self.number_size_var.get())
        self.cfg["show_date"] = bool(self.show_date_var.get())
        self.cfg["printer_name"] = self.printer_var.get().strip()

        save_json(CONFIG_FILE, self.cfg)
        self.on_save(self.cfg)
        self.destroy()


class NobatApp:
    def __init__(self, root):
        self.root = root
        self.root.title("نرم‌افزار نوبت‌دهی")
        self.root.geometry("440x420")
        self.root.configure(bg="#f4f4f4")

        self.cfg = load_json(CONFIG_FILE, DEFAULT_CONFIG)
        self.state = load_json(STATE_FILE, {"date": "", "current_number": self.cfg["start_number"] - 1})

        today_str = date.today().isoformat()
        if self.state.get("date") != today_str:
            self.state = {"date": today_str, "current_number": self.cfg["start_number"] - 1}
            save_json(STATE_FILE, self.state)

        self.build_ui()

    def build_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        top_bar = tk.Frame(self.root, bg="#f4f4f4")
        top_bar.pack(fill="x", pady=(10, 0), padx=10)
        tk.Button(top_bar, text="⚙ تنظیمات", font=("Tahoma", 10), command=self.open_settings).pack(side="left")

        self.header_label = tk.Label(
            self.root, text=self.cfg["header_text"], font=("Tahoma", 16, "bold"),
            bg="#f4f4f4", justify="center"
        )
        self.header_label.pack(pady=(15, 5))

        self.number_label = tk.Label(
            self.root, text=str(max(self.state["current_number"], self.cfg["start_number"] - 1)),
            font=("Tahoma", 60, "bold"), bg="#f4f4f4", fg="#1a1a1a"
        )
        self.number_label.pack(pady=10)

        self.range_label = tk.Label(
            self.root, text=f"بازه: {self.cfg['start_number']} تا {self.cfg['end_number']}",
            font=("Tahoma", 9), bg="#f4f4f4", fg="#777"
        )
        self.range_label.pack()

        self.status_label = tk.Label(self.root, text="آماده", font=("Tahoma", 10), bg="#f4f4f4", fg="#555")
        self.status_label.pack(pady=(5, 0))

        tk.Button(
            self.root, text="نوبت بعدی و چاپ", font=("Tahoma", 14, "bold"),
            bg="#2e7d32", fg="white", padx=20, pady=12,
            command=self.next_ticket
        ).pack(pady=20)

        tk.Button(
            self.root, text="شروع مجدد از عدد اول بازه", font=("Tahoma", 9),
            fg="#a33", command=self.reset_sequence
        ).pack()

        if not WINDOWS_PRINTING_AVAILABLE:
            self.status_label.config(text="⚠ pywin32 نصب نیست - چاپ غیرفعال است", fg="red")

    def open_settings(self):
        SettingsDialog(self.root, self.cfg, self.on_settings_saved)

    def on_settings_saved(self, new_cfg):
        self.cfg = new_cfg
        self.build_ui()

    def reset_sequence(self):
        if messagebox.askyesno("تایید", "شماره‌ی نوبت از ابتدای بازه شروع شود؟"):
            self.state["current_number"] = self.cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)
            self.build_ui()

    def next_ticket(self):
        next_num = self.state["current_number"] + 1
        if next_num > self.cfg["end_number"]:
            if messagebox.askyesno(
                "پایان بازه",
                "به آخرین شماره‌ی بازه رسیدید. می‌خواهید دوباره از ابتدای بازه شروع شود؟"
            ):
                next_num = self.cfg["start_number"]
            else:
                return

        self.state["current_number"] = next_num
        save_json(STATE_FILE, self.state)
        self.number_label.config(text=str(next_num))

        ok = print_ticket(next_num, self.cfg)
        if ok:
            self.status_label.config(text=f"نوبت {next_num} چاپ شد", fg="#2e7d32")
        else:
            self.status_label.config(text=f"نوبت {next_num} ثبت شد ولی چاپ نشد", fg="red")


if __name__ == "__main__":
    root = tk.Tk()
    app = NobatApp(root)
    root.mainloop()
