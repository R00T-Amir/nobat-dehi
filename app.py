# -*- coding: utf-8 -*-
"""
نرم‌افزار نوبت‌دهی پیشرفته
================================
- رابط کاربری مدرن با پیش‌نمایش زنده فیش
- چاپ دستی (تک‌تک) و چاپ خودکار در بازه مشخص
- تنظیم سایز کاغذ (58mm / 80mm)
- مدیریت تاریخ (تاریخ سیستم یا تاریخ دستی / حذف تاریخ)
- ذخیره تنظیمات و状态 روزانه
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
    "start_number": 100,
    "end_number": 999,
    "header_font_size": 1,
    "number_font_size": 3,
    "show_date": True,
    "custom_date_text": "",  # خالی باشد = تاریخ امروز
    "paper_width": 32,       # 32 برای 58mm ، 48 برای 80mm
    "auto_start": 1,
    "auto_end": 20,
}

# رنگ‌های مدرن برای رابط کاربری
COLORS = {
    "bg": "#f0f4f8",
    "card_bg": "#ffffff",
    "primary": "#3b82f6",
    "primary_hover": "#2563eb",
    "success": "#10b981",
    "danger": "#ef4444",
    "text_dark": "#1e293b",
    "text_gray": "#64748b",
    "border": "#e2e8f0"
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

def font_size_byte(size_level: int) -> bytes:
    size_level = max(0, min(7, size_level))
    n = (size_level << 4) | size_level
    return bytes([n])

def build_escpos_bytes(number: int, cfg: dict) -> bytes:
    ESC = b"\x1b"
    GS = b"\x1d"
    width = cfg.get("paper_width", 32)
    dashes = "-" * width + "\n"

    data = b""
    data += ESC + b"@"                      # ریست
    data += ESC + b"a" + b"\x01"            # وسط‌چین

    # سربرگ
    data += GS + b"!" + font_size_byte(cfg["header_font_size"])
    for line in cfg["header_text"].splitlines():
        data += (line + "\n").encode("cp1256", errors="ignore")
    data += GS + b"!" + b"\x00"             # فونت عادی
    
    # خط تیره
    data += ESC + b"a" + b"\x00"            # چپ‌چین برای خط تیره
    data += dashes.encode("cp1256", errors="ignore")
    data += ESC + b"a" + b"\x01"            # بازگشت به وسط‌چین

    # شماره
    data += GS + b"!" + font_size_byte(cfg["number_font_size"])
    data += (str(number) + "\n").encode("cp1256", errors="ignore")
    data += GS + b"!" + b"\x00"

    # خط تیره و تاریخ
    data += ESC + b"a" + b"\x00"
    data += dashes.encode("cp1256", errors="ignore")
    
    if cfg.get("show_date", True):
        if cfg.get("custom_date_text", "").strip():
            date_str = cfg["custom_date_text"]
        else:
            date_str = date.today().strftime("%Y-%m-%d")
        data += (date_str + "\n").encode("cp1256", errors="ignore")

    data += b"\n\n\n"
    data += GS + b"V" + b"\x41" + b"\x10"   # برش کاغذ
    return data

def print_ticket(number: int, cfg: dict) -> bool:
    if not WINDOWS_PRINTING_AVAILABLE:
        messagebox.showwarning("چاپ در دسترس نیست", "چاپ فقط روی ویندوز و با نصب pywin32 کار می‌کند.")
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
        messagebox.showerror("خطا در چاپ", f"چاپ ناموفق بود:\n{e}\n\nپرینترهای موجود:\n" + "\n".join(list_printers()))
        return False

class SettingsDialog(tk.Toplevel):
    def __init__(self, master, cfg: dict, on_save):
        super().__init__(master)
        self.title("تنظیمات سیستم")
        self.geometry("500x650")
        self.configure(bg=COLORS["bg"])
        self.cfg = dict(cfg)
        self.on_save = on_save
        self.resizable(False, False)
        self.grab_set()

        # اسکرول فریم ساده
        canvas = tk.Canvas(self, bg=COLORS["bg"], highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)
        scroll_y = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scroll_y.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scroll_y.set)
        
        frame = tk.Frame(canvas, bg=COLORS["bg"])
        canvas.create_window((0,0), window=frame, anchor="nw")
        
        # ساخت ویجت‌ها
        self.create_widgets(frame)
        
        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    def create_widgets(self, frame):
        pad = {"padx": 20, "pady": 5}
        font_label = ("Tahoma", 10, "bold")
        font_entry = ("Tahoma", 10)

        # --- بخش سربرگ ---
        tk.Label(frame, text="متن سربرگ (چند خط مجاز است):", bg=COLORS["bg"], font=font_label, fg=COLORS["text_dark"]).pack(anchor="e", **pad)
        self.header_text = tk.Text(frame, height=4, font=font_entry, relief="solid", borderwidth=1)
        self.header_text.insert("1.0", self.cfg.get("header_text", ""))
        self.header_text.pack(fill="x", **pad)

        # --- بازه اعداد ---
        range_frame = tk.LabelFrame(frame, text="بازه شماره‌گذاری", bg=COLORS["bg"], font=font_label, fg=COLORS["text_gray"], bd=1, relief="solid")
        range_frame.pack(fill="x", **pad, pady=10)
        
        tk.Label(range_frame, text="شماره شروع:", bg=COLORS["bg"], font=font_entry).grid(row=0, column=1, sticky="e", padx=5, pady=5)
        self.start_var = tk.IntVar(value=self.cfg.get("start_number", 1))
        tk.Entry(range_frame, textvariable=self.start_var, width=10, justify="center", font=font_entry).grid(row=0, column=0, padx=5)

        tk.Label(range_frame, text="شماره پایان:", bg=COLORS["bg"], font=font_entry).grid(row=1, column=1, sticky="e", padx=5, pady=5)
        self.end_var = tk.IntVar(value=self.cfg.get("end_number", 999))
        tk.Entry(range_frame, textvariable=self.end_var, width=10, justify="center", font=font_entry).grid(row=1, column=0, padx=5)

        # --- چاپ خودکار ---
        auto_frame = tk.LabelFrame(frame, text="بازه چاپ خودکار", bg=COLORS["bg"], font=font_label, fg=COLORS["text_gray"], bd=1, relief="solid")
        auto_frame.pack(fill="x", **pad, pady=10)

        tk.Label(auto_frame, text="از شماره:", bg=COLORS["bg"], font=font_entry).grid(row=0, column=1, sticky="e", padx=5, pady=5)
        self.auto_start_var = tk.IntVar(value=self.cfg.get("auto_start", 1))
        tk.Entry(auto_frame, textvariable=self.auto_start_var, width=10, justify="center", font=font_entry).grid(row=0, column=0, padx=5)

        tk.Label(auto_frame, text="تا شماره:", bg=COLORS["bg"], font=font_entry).grid(row=1, column=1, sticky="e", padx=5, pady=5)
        self.auto_end_var = tk.IntVar(value=self.cfg.get("auto_end", 10))
        tk.Entry(auto_frame, textvariable=self.auto_end_var, width=10, justify="center", font=font_entry).grid(row=1, column=0, padx=5)

        # --- سایز فونت ---
        font_frame = tk.LabelFrame(frame, text="تنظیمات فونت و کاغذ", bg=COLORS["bg"], font=font_label, fg=COLORS["text_gray"], bd=1, relief="solid")
        font_frame.pack(fill="x", **pad, pady=10)

        tk.Label(font_frame, text="سایز فونت سربرگ (۰-۷):", bg=COLORS["bg"], font=font_entry).grid(row=0, column=1, sticky="e", padx=5, pady=5)
        self.header_size_var = tk.IntVar(value=self.cfg.get("header_font_size", 1))
        ttk.Spinbox(font_frame, from_=0, to=7, textvariable=self.header_size_var, width=8, justify="center").grid(row=0, column=0, padx=5)

        tk.Label(font_frame, text="سایز فونت شماره (۰-۷):", bg=COLORS["bg"], font=font_entry).grid(row=1, column=1, sticky="e", padx=5, pady=5)
        self.number_size_var = tk.IntVar(value=self.cfg.get("number_font_size", 3))
        ttk.Spinbox(font_frame, from_=0, to=7, textvariable=self.number_size_var, width=8, justify="center").grid(row=1, column=0, padx=5)

        tk.Label(font_frame, text="عرض کاغذ پرینتر:", bg=COLORS["bg"], font=font_entry).grid(row=2, column=1, sticky="e", padx=5, pady=5)
        self.paper_var = tk.IntVar(value=self.cfg.get("paper_width", 32))
        paper_combo = ttk.Combobox(font_frame, textvariable=self.paper_var, values=[32, 48], state="readonly", justify="center", width=6)
        paper_combo.grid(row=2, column=0, padx=5)
        tk.Label(font_frame, text="(۳۲=58mm | ۴۸=80mm)", bg=COLORS["bg"], font=("Tahoma", 8), fg=COLORS["text_gray"]).grid(row=2, column=2, sticky="w", padx=5)

        # --- تنظیمات تاریخ ---
        date_frame = tk.LabelFrame(frame, text="تنظیمات تاریخ", bg=COLORS["bg"], font=font_label, fg=COLORS["text_gray"], bd=1, relief="solid")
        date_frame.pack(fill="x", **pad, pady=10)

        self.show_date_var = tk.BooleanVar(value=self.cfg.get("show_date", True))
        tk.Checkbutton(date_frame, text="نمایش تاریخ در فیش", variable=self.show_date_var, bg=COLORS["bg"], font=font_entry).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)

        tk.Label(date_frame, text="متن تاریخ (خالی=تاریخ امروز):", bg=COLORS["bg"], font=font_entry).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.custom_date_var = tk.StringVar(value=self.cfg.get("custom_date_text", ""))
        tk.Entry(date_frame, textvariable=self.custom_date_var, width=20, justify="center", font=font_entry).grid(row=1, column=1, padx=5)

        # --- پرینتر ---
        tk.Label(frame, text="انتخاب پرینتر:", bg=COLORS["bg"], font=font_label, fg=COLORS["text_dark"]).pack(anchor="e", **pad)
        self.printer_var = tk.StringVar(value=self.cfg.get("printer_name", "Meva TP1000"))
        printers = list_printers()
        if printers:
            ttk.Combobox(frame, textvariable=self.printer_var, values=printers, justify="center").pack(fill="x", **pad)
        else:
            tk.Entry(frame, textvariable=self.printer_var, justify="center").pack(fill="x", **pad)

        # --- دکمه‌ها ---
        btn_frame = tk.Frame(frame, bg=COLORS["bg"])
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="ذخیره تنظیمات", font=("Tahoma", 11, "bold"), bg=COLORS["success"], fg="white", padx=30, pady=8, bd=0, command=self.save).pack(side="right", padx=10)
        tk.Button(btn_frame, text="انصراف", font=("Tahoma", 11), bg=COLORS["text_gray"], fg="white", padx=30, pady=8, bd=0, command=self.destroy).pack(side="right")

    def save(self):
        try:
            start = int(self.start_var.get())
            end = int(self.end_var.get())
            auto_start = int(self.auto_start_var.get())
            auto_end = int(self.auto_end_var.get())
        except:
            messagebox.showerror("خطا", "اعداد بازه باید عدد صحیح باشند.", parent=self)
            return
        if end < start:
            messagebox.showerror("خطا", "شماره پایان باید بزرگ‌تر از شروع باشد.", parent=self)
            return

        self.cfg["header_text"] = self.header_text.get("1.0", "end").strip()
        self.cfg["start_number"] = start
        self.cfg["end_number"] = end
        self.cfg["auto_start"] = auto_start
        self.cfg["auto_end"] = auto_end
        self.cfg["header_font_size"] = int(self.header_size_var.get())
        self.cfg["number_font_size"] = int(self.number_size_var.get())
        self.cfg["paper_width"] = int(self.paper_var.get())
        self.cfg["show_date"] = bool(self.show_date_var.get())
        self.cfg["custom_date_text"] = self.custom_date_var.get().strip()
        self.cfg["printer_name"] = self.printer_var.get().strip()

        save_json(CONFIG_FILE, self.cfg)
        self.on_save(self.cfg)
        self.destroy()

class NobatApp:
    def __init__(self, root):
        self.root = root
        self.root.title("نرم‌افزار نوبت‌دهی هوشمند")
        self.root.geometry("500x720")
        self.root.configure(bg=COLORS["bg"])
        
        self.is_auto_printing = False

        self.cfg = load_json(CONFIG_FILE, DEFAULT_CONFIG)
        self.state = load_json(STATE_FILE, {"date": "", "current_number": self.cfg["start_number"] - 1})

        today_str = date.today().isoformat()
        if self.state.get("date") != today_str:
            self.state = {"date": today_str, "current_number": self.cfg["start_number"] - 1}
            save_json(STATE_FILE, self.state)

        # رفع باگ: اگر شماره فعلی از شروع جدید کمتر بود، اصلاح کن
        if self.state["current_number"] < self.cfg["start_number"] - 1:
            self.state["current_number"] = self.cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)

        self.build_ui()

    def build_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        # هدر بالا
        top_bar = tk.Frame(self.root, bg=COLORS["bg"])
        top_bar.pack(fill="x", pady=(10,0), padx=15)
        tk.Button(top_bar, text="⚙ تنظیمات", font=("Tahoma", 10, "bold"), bg=COLORS["text_gray"], fg="white", bd=0, padx=15, pady=5, command=self.open_settings).pack(side="right")
        
        self.header_label = tk.Label(top_bar, text=self.cfg["header_text"], font=("Tahoma", 14, "bold"), bg=COLORS["bg"], fg=COLORS["text_dark"])
        self.header_label.pack(side="right", padx=20)

        # --- پیش‌نمایش زنده فیش ---
        preview_frame = tk.Frame(self.root, bg=COLORS["card_bg"], bd=1, relief="solid")
        preview_frame.pack(pady=15, padx=40, fill="x")
        
        tk.Label(preview_frame, text="پیش‌نمایش فیش", font=("Tahoma", 9, "bold"), bg=COLORS["card_bg"], fg=COLORS["text_gray"]).pack(pady=(10,0))

        # محاسبه سایز فونت برای پیش‌نمایش (تقریبی بر اساس ESC/POS)
        prev_header_size = 10 + (self.cfg["header_font_size"] * 4)
        prev_num_size = 24 + (self.cfg["number_font_size"] * 12)
        
        self.prev_header = tk.Label(preview_frame, text=self.cfg["header_text"], font=("Tahoma", prev_header_size, "bold"), bg=COLORS["card_bg"], fg=COLORS["text_dark"], justify="center")
        self.prev_header.pack(pady=(10,5))

        width = self.cfg.get("paper_width", 32)
        self.prev_dashes1 = tk.Label(preview_frame, text="-"*width, font=("Courier", 8), bg=COLORS["card_bg"], fg=COLORS["border"])
        self.prev_dashes1.pack()

        self.prev_number = tk.Label(preview_frame, text=str(max(self.state["current_number"], self.cfg["start_number"])), font=("Tahoma", prev_num_size, "bold"), bg=COLORS["card_bg"], fg=COLORS["primary"])
        self.prev_number.pack(pady=10)

        self.prev_dashes2 = tk.Label(preview_frame, text="-"*width, font=("Courier", 8), bg=COLORS["card_bg"], fg=COLORS["border"])
        self.prev_dashes2.pack()
        
        if self.cfg.get("show_date", True):
            date_txt = self.cfg.get("custom_date_text", "").strip() or date.today().strftime("%Y-%m-%d")
            self.prev_date = tk.Label(preview_frame, text=date_txt, font=("Tahoma", 10), bg=COLORS["card_bg"], fg=COLORS["text_gray"])
            self.prev_date.pack(pady=(5, 15))
        else:
            self.prev_date = tk.Label(preview_frame, text="", bg=COLORS["card_bg"])
            self.prev_date.pack(pady=(0, 15))

        # --- بخش چاپ خودکار ---
        auto_frame = tk.Frame(self.root, bg=COLORS["bg"])
        auto_frame.pack(fill="x", padx=20, pady=5)
        
        auto_info = f"چاپ خودکار: از {self.cfg['auto_start']} تا {self.cfg['auto_end']}"
        tk.Label(auto_frame, text=auto_info, font=("Tahoma", 10), bg=COLORS["bg"], fg=COLORS["text_dark"]).pack(side="right")
        
        self.btn_auto = tk.Button(auto_frame, text="▶ شروع چاپ خودکار", font=("Tahoma", 9, "bold"), bg=COLORS["primary"], fg="white", bd=0, padx=10, pady=5, command=self.toggle_auto_print)
        self.btn_auto.pack(side="left")

        # --- دکمه‌های اصلی ---
        btn_frame = tk.Frame(self.root, bg=COLORS["bg"])
        btn_frame.pack(pady=15)
        
        self.btn_next = tk.Button(btn_frame, text="نوبت بعدی و چاپ", font=("Tahoma", 13, "bold"), bg=COLORS["success"], fg="white", padx=40, pady=12, bd=0, command=self.next_ticket)
        self.btn_next.pack()

        self.status_label = tk.Label(self.root, text="سیستم آماده کار", font=("Tahoma", 10), bg=COLORS["bg"], fg=COLORS["text_gray"])
        self.status_label.pack(pady=(5, 10))
        
        tk.Button(self.root, text="شروع مجدد از عدد اول بازه", font=("Tahoma", 9), bg=COLORS["bg"], fg=COLORS["danger"], bd=0, command=self.reset_sequence).pack()

    def update_preview_number(self, num):
        self.prev_number.config(text=str(num))

    def open_settings(self):
        SettingsDialog(self.root, self.cfg, self.on_settings_saved)

    def on_settings_saved(self, new_cfg):
        self.cfg = new_cfg
        if self.state["current_number"] < self.cfg["start_number"] - 1:
            self.state["current_number"] = self.cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)
        self.build_ui()

    def reset_sequence(self):
        if messagebox.askyesno("تایید", "شماره نوبت از ابتدای بازه شروع شود؟"):
            self.state["current_number"] = self.cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)
            self.update_preview_number(self.cfg["start_number"])
            self.status_label.config(text="بازه ریست شد", fg=COLORS["primary"])

    def toggle_auto_print(self):
        if self.is_auto_printing:
            self.is_auto_printing = False
            self.btn_auto.config(text="▶ شروع چاپ خودکار", bg=COLORS["primary"])
            self.btn_next.config(state="normal")
            self.status_label.config(text="چاپ خودکار متوقف شد", fg=COLORS["danger"])
        else:
            curr = self.state["current_number"]
            if curr < self.cfg["auto_start"] - 1:
                self.state["current_number"] = self.cfg["auto_start"] - 1
                save_json(STATE_FILE, self.state)
            
            self.is_auto_printing = True
            self.btn_auto.config(text="⏸ توقف چاپ خودکار", bg=COLORS["danger"])
            self.btn_next.config(state="disabled")
            self.status_label.config(text="در حال چاپ خودکار...", fg=COLORS["primary"])
            self.root.after(500, self.auto_print_step)

    def auto_print_step(self):
        if not self.is_auto_printing:
            return
            
        curr = self.state["current_number"]
        if curr >= self.cfg["auto_end"]:
            self.is_auto_printing = False
            self.btn_auto.config(text="▶ شروع چاپ خودکار", bg=COLORS["primary"])
            self.btn_next.config(state="normal")
            self.status_label.config(text="چاپ خودکار به پایان رسید", fg=COLORS["success"])
            return

        next_num = curr + 1
        self.state["current_number"] = next_num
        save_json(STATE_FILE, self.state)
        self.update_preview_number(next_num)
        
        ok = print_ticket(next_num, self.cfg)
        if ok:
            self.status_label.config(text=f"چاپ شد: {next_num}", fg=COLORS["success"])
            self.root.after(1000, self.auto_print_step) # 1 ثانیه فاصله تا چاپ بعدی
        else:
            self.is_auto_printing = False
            self.btn_auto.config(text="▶ شروع چاپ خودکار", bg=COLORS["primary"])
            self.btn_next.config(state="normal")
            self.status_label.config(text="خطا در چاپ خودکار!", fg=COLORS["danger"])

    def next_ticket(self):
        next_num = self.state["current_number"] + 1
        if next_num > self.cfg["end_number"]:
            if messagebox.askyesno("پایان بازه", "به آخرین شماره رسیدید. از ابتدا شروع شود؟"):
                next_num = self.cfg["start_number"]
            else:
                return

        self.state["current_number"] = next_num
        save_json(STATE_FILE, self.state)
        self.update_preview_number(next_num)

        ok = print_ticket(next_num, self.cfg)
        if ok:
            self.status_label.config(text=f"نوبت {next_num} چاپ شد", fg=COLORS["success"])
        else:
            self.status_label.config(text=f"نوبت {next_num} ثبت شد ولی چاپ نشد", fg=COLORS["danger"])

if __name__ == "__main__":
    root = tk.Tk()
    app = NobatApp(root)
    root.mainloop()
