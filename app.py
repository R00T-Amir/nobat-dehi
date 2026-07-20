# -*- coding: utf-8 -*-
import json
import os
from datetime import date
import tkinter as tk
from tkinter import messagebox
import winsound  # برای پخش صدای بوق

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
    "header_font_size": 1,
    "number_font_size": 3,
    "show_date": True,
    "custom_date_text": "",
    "paper_width": 32,
    "ticket_mode": "normal",  # normal یا compact
    "auto_start": 1,
    "auto_end": 20,
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
    data += ESC + b"@"                      
    data += ESC + b"a" + b"\x01"            

    # سربرگ
    data += GS + b"!" + font_size_byte(cfg["header_font_size"])
    for line in cfg["header_text"].splitlines():
        data += (line + "\n").encode("cp1256", errors="ignore")
    data += GS + b"!" + b"\x00"             
    
    # خط تیره اول
    data += ESC + b"a" + b"\x00"            # چپ‌چین برای خط تیره
    data += dashes.encode("cp1256", errors="ignore")
    data += ESC + b"a" + b"\x01"            # بازگشت به وسط‌چین برای شماره

    # شماره نوبت
    data += GS + b"!" + font_size_byte(cfg["number_font_size"])
    data += (str(number) + "\n").encode("cp1256", errors="ignore")
    data += GS + b"!" + b"\x00"

    # خط تیره دوم
    data += ESC + b"a" + b"\x00"            # چپ‌چین برای خط تیره
    data += dashes.encode("cp1256", errors="ignore")
    
    # تاریخ (اصلاح وسط‌چین بودن)
    if cfg.get("show_date", True):
        data += ESC + b"a" + b"\x01"  # ----> بازگشت به وسط‌چین برای تاریخ
        if cfg.get("custom_date_text", "").strip():
            date_str = cfg["custom_date_text"]
        else:
            date_str = date.today().strftime("%Y-%m-%d")
        data += (date_str + "\n").encode("cp1256", errors="ignore")

    # مدیریت طول کاغذ (فاصله پایین و برش)
    if cfg.get("ticket_mode", "normal") == "compact":
        # حذف کامل خطوط خالی و کاهش فاصله برش به ۵ واحد (حداقل ممکن)
        data += b"\n" 
        data += GS + b"V" + b"\x41" + b"\x05"   # برش با حداقل فاصله
    else:
        # حالت عادی با فاصله استاندارد
        data += b"\n\n\n"
        data += GS + b"V" + b"\x41" + b"\x10"   # برش با فاصله معمولی
        
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
        self.geometry("450x720")
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

        frame_auto = tk.Frame(self)
        frame_auto.pack(fill="x", **pad)
        tk.Label(frame_auto, text="شروع چاپ خودکار:", font=("Tahoma", 10)).grid(row=0, column=1, sticky="e", padx=5)
        self.auto_start_var = tk.IntVar(value=self.cfg.get("auto_start", 1))
        tk.Entry(frame_auto, textvariable=self.auto_start_var, width=10, justify="center").grid(row=0, column=0, padx=5)

        tk.Label(frame_auto, text="پایان چاپ خودکار:", font=("Tahoma", 10)).grid(row=1, column=1, sticky="e", padx=5, pady=5)
        self.auto_end_var = tk.IntVar(value=self.cfg.get("auto_end", 10))
        tk.Entry(frame_auto, textvariable=self.auto_end_var, width=10, justify="center").grid(row=1, column=0, padx=5)

        frame_font = tk.Frame(self)
        frame_font.pack(fill="x", **pad)
        tk.Label(frame_font, text="سایز فونت سربرگ (۰ تا ۷):", font=("Tahoma", 10)).grid(row=0, column=1, sticky="e", padx=5)
        self.header_size_var = tk.IntVar(value=self.cfg.get("header_font_size", 1))
        tk.Spinbox(frame_font, from_=0, to=7, textvariable=self.header_size_var, width=8, justify="center").grid(row=0, column=0, padx=5)

        tk.Label(frame_font, text="سایز فونت شماره (۰ تا ۷):", font=("Tahoma", 10)).grid(row=1, column=1, sticky="e", padx=5, pady=5)
        self.number_size_var = tk.IntVar(value=self.cfg.get("number_font_size", 3))
        tk.Spinbox(frame_font, from_=0, to=7, textvariable=self.number_size_var, width=8, justify="center").grid(row=1, column=0, padx=5)

        tk.Label(frame_font, text="عرض کاغذ (۳۲ یا ۴۸):", font=("Tahoma", 10)).grid(row=2, column=1, sticky="e", padx=5, pady=5)
        self.paper_var = tk.IntVar(value=self.cfg.get("paper_width", 32))
        tk.Spinbox(frame_font, from_=32, to=48, increment=16, textvariable=self.paper_var, width=8, justify="center").grid(row=2, column=0, padx=5)

        tk.Label(frame_font, text="حالت چاپ (طول):", font=("Tahoma", 10)).grid(row=3, column=1, sticky="e", padx=5, pady=5)
        self.ticket_mode_var = tk.StringVar(value=self.cfg.get("ticket_mode", "normal"))
        mode_menu = tk.OptionMenu(frame_font, self.ticket_mode_var, "normal", "compact")
        mode_menu.config(font=("Tahoma", 9), width=8)
        mode_menu.grid(row=3, column=0, padx=5, sticky="w")
        tk.Label(frame_font, text="(کوچک=کمترین فاصله)", font=("Tahoma", 8), fg="#777").grid(row=3, column=2, sticky="w", padx=5)

        frame_date = tk.Frame(self)
        frame_date.pack(fill="x", **pad)
        self.show_date_var = tk.BooleanVar(value=self.cfg.get("show_date", True))
        tk.Checkbutton(frame_date, text="نمایش تاریخ روی فیش", variable=self.show_date_var, font=("Tahoma", 10)).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)

        tk.Label(frame_date, text="تاریخ دلخواه (خالی=امروز):", font=("Tahoma", 10)).grid(row=1, column=1, sticky="e", padx=5)
        self.custom_date_var = tk.StringVar(value=self.cfg.get("custom_date_text", ""))
        tk.Entry(frame_date, textvariable=self.custom_date_var, width=15, justify="center").grid(row=1, column=0, padx=5)

        tk.Label(self, text="نام پرینتر (دقیقاً مطابق ویندوز):", font=("Tahoma", 10, "bold")).pack(anchor="e", **pad)
        self.printer_var = tk.StringVar(value=self.cfg.get("printer_name", "Meva TP1000"))
        printers = list_printers()
        if printers:
            printer_menu = tk.OptionMenu(self, self.printer_var, self.printer_var.get(), *printers)
            printer_menu.config(font=("Tahoma", 10), width=40)
            printer_menu.pack(fill="x", padx=15)
        else:
            tk.Entry(self, textvariable=self.printer_var, justify="center").pack(fill="x", padx=15)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="ذخیره", font=("Tahoma", 11, "bold"), bg="#2e7d32", fg="white", padx=20, pady=8, command=self.save).pack(side="right", padx=10)
        tk.Button(btn_frame, text="انصراف", font=("Tahoma", 11), padx=20, pady=8, command=self.destroy).pack(side="right")

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
        self.cfg["ticket_mode"] = self.ticket_mode_var.get()
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
        self.root.geometry("440x600")
        self.root.configure(bg="#f4f4f4")
        
        self.is_auto_printing = False

        self.cfg = load_json(CONFIG_FILE, DEFAULT_CONFIG)
        self.state = load_json(STATE_FILE, {"date": "", "current_number": self.cfg["start_number"] - 1})

        today_str = date.today().isoformat()
        if self.state.get("date") != today_str:
            self.state = {"date": today_str, "current_number": self.cfg["start_number"] - 1}
            save_json(STATE_FILE, self.state)

        if self.state["current_number"] < self.cfg["start_number"] - 1:
            self.state["current_number"] = self.cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)

        self.build_ui()
        
        # میانبر صفحه‌کلید (Enter برای چاپ)
        self.root.bind("<Return>", lambda e: self.next_ticket())
        self.root.bind("<space>", lambda e: self.next_ticket())

    def build_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        top_bar = tk.Frame(self.root, bg="#f4f4f4")
        top_bar.pack(fill="x", pady=(10,0), padx=10)
        tk.Button(top_bar, text="⚙ تنظیمات", font=("Tahoma", 10), command=self.open_settings).pack(side="left")
        
        self.header_label = tk.Label(top_bar, text=self.cfg["header_text"], font=("Tahoma", 14, "bold"), bg="#f4f4f4", justify="center")
        self.header_label.pack(side="left", padx=20)

        # --- پیش‌نمایش زنده فیش ---
        preview_frame = tk.Frame(self.root, bg="white", bd=1, relief="solid")
        preview_frame.pack(pady=10, padx=30, fill="x")
        
        prev_header_size = 10 + (self.cfg["header_font_size"] * 3)
        prev_num_size = 24 + (self.cfg["number_font_size"] * 10)
        
        self.prev_header = tk.Label(preview_frame, text=self.cfg["header_text"], font=("Tahoma", prev_header_size, "bold"), bg="white", fg="#1a1a1a", justify="center")
        self.prev_header.pack(pady=(10,5))

        width = self.cfg.get("paper_width", 32)
        self.prev_dashes1 = tk.Label(preview_frame, text="-"*width, font=("Courier", 8), bg="white", fg="#ccc")
        self.prev_dashes1.pack()

        self.prev_number = tk.Label(preview_frame, text=str(max(self.state["current_number"], self.cfg["start_number"])), font=("Tahoma", prev_num_size, "bold"), bg="white", fg="#3b82f6")
        self.prev_number.pack(pady=10)

        self.prev_dashes2 = tk.Label(preview_frame, text="-"*width, font=("Courier", 8), bg="white", fg="#ccc")
        self.prev_dashes2.pack()
        
        if self.cfg.get("show_date", True):
            date_txt = self.cfg.get("custom_date_text", "").strip() or date.today().strftime("%Y-%m-%d")
            self.prev_date = tk.Label(preview_frame, text=date_txt, font=("Tahoma", 10), bg="white", fg="#777")
            self.prev_date.pack(pady=(5, 10))
        else:
            self.prev_date = tk.Label(preview_frame, text="", bg="white")
            self.prev_date.pack(pady=(0, 10))

        self.range_label = tk.Label(
            self.root, text=f"دستی: {self.cfg['start_number']} تا {self.cfg['end_number']} | خودکار: {self.cfg['auto_start']} تا {self.cfg['auto_end']}",
            font=("Tahoma", 9), bg="#f4f4f4", fg="#777"
        )
        self.range_label.pack()

        self.status_label = tk.Label(self.root, text="آماده کار (Enter = نوبت بعدی)", font=("Tahoma", 10), bg="#f4f4f4", fg="#555")
        self.status_label.pack(pady=(5, 0))

        btn_frame = tk.Frame(self.root, bg="#f4f4f4")
        btn_frame.pack(pady=10)
        
        self.btn_next = tk.Button(btn_frame, text="نوبت بعدی و چاپ", font=("Tahoma", 14, "bold"), bg="#2e7d32", fg="white", padx=20, pady=12, command=self.next_ticket)
        self.btn_next.pack(side="left", padx=10)
        
        self.btn_auto = tk.Button(btn_frame, text="▶ چاپ خودکار", font=("Tahoma", 12, "bold"), bg="#3b82f6", fg="white", padx=15, pady=12, command=self.toggle_auto_print)
        self.btn_auto.pack(side="left")

        tk.Button(self.root, text="شروع مجدد از عدد اول بازه", font=("Tahoma", 9), fg="#a33", bg="#f4f4f4", command=self.reset_sequence).pack()

    def update_preview_number(self, num):
        self.prev_number.config(text=str(num))

    def open_settings(self):
        SettingsDialog(self.root, self.cfg, self.on_settings_saved)

    def on_settings_saved(self, new_cfg):
        if new_cfg["start_number"] != self.cfg.get("start_number"):
            self.state["current_number"] = new_cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)
        elif self.state["current_number"] < new_cfg["start_number"] - 1 or self.state["current_number"] > new_cfg["end_number"]:
            self.state["current_number"] = new_cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)
            
        self.cfg = new_cfg
        self.build_ui()

    def reset_sequence(self):
        if messagebox.askyesno("تایید", "شماره نوبت از ابتدای بازه شروع شود؟"):
            self.state["current_number"] = self.cfg["start_number"] - 1
            save_json(STATE_FILE, self.state)
            self.update_preview_number(self.cfg["start_number"])
            self.status_label.config(text="بازه ریست شد", fg="#3b82f6")

    def toggle_auto_print(self):
        if self.is_auto_printing:
            self.is_auto_printing = False
            self.btn_auto.config(text="▶ چاپ خودکار", bg="#3b82f6")
            self.btn_next.config(state="normal")
            self.status_label.config(text="چاپ خودکار متوقف شد", fg="#ef4444")
        else:
            curr = self.state["current_number"]
            if curr < self.cfg["auto_start"] - 1 or curr > self.cfg["auto_end"]:
                self.state["current_number"] = self.cfg["auto_start"] - 1
                save_json(STATE_FILE, self.state)
                self.update_preview_number(self.cfg["auto_start"])
            
            self.is_auto_printing = True
            self.btn_auto.config(text="⏸ توقف", bg="#ef4444")
            self.btn_next.config(state="disabled")
            self.status_label.config(text="در حال چاپ خودکار...", fg="#3b82f6")
            self.root.after(500, self.auto_print_step)

    def auto_print_step(self):
        if not self.is_auto_printing:
            return
            
        curr = self.state["current_number"]
        if curr >= self.cfg["auto_end"]:
            self.is_auto_printing = False
            self.btn_auto.config(text="▶ چاپ خودکار", bg="#3b82f6")
            self.btn_next.config(state="normal")
            self.status_label.config(text="چاپ خودکار به پایان رسید", fg="#2e7d32")
            return

        next_num = curr + 1
        self.state["current_number"] = next_num
        save_json(STATE_FILE, self.state)
        self.update_preview_number(next_num)
        
        ok = print_ticket(next_num, self.cfg)
        if ok:
            winsound.Beep(1000, 150) # صدای بوق
            self.status_label.config(text=f"چاپ شد: {next_num}", fg="#2e7d32")
            self.root.after(1000, self.auto_print_step)
        else:
            self.is_auto_printing = False
            self.btn_auto.config(text="▶ چاپ خودکار", bg="#3b82f6")
            self.btn_next.config(state="normal")
            self.status_label.config(text="خطا در چاپ خودکار!", fg="#ef4444")

    def next_ticket(self):
        # جلوگیری از تداخل با دکمه غیرفعال در زمان چاپ خودکار
        if self.is_auto_printing:
            return
            
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
            winsound.Beep(1000, 150) # صدای بوق
            self.status_label.config(text=f"نوبت {next_num} چاپ شد", fg="#2e7d32")
        else:
            self.status_label.config(text=f"نوبت {next_num} ثبت شد ولی چاپ نشد", fg="#ef4444")

if __name__ == "__main__":
    root = tk.Tk()
    app = NobatApp(root)
    root.mainloop()
