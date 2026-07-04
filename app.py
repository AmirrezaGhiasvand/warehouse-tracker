# -*- coding: utf-8 -*-
"""
app.py
Warehouse Tracker (سامانه پیگیری انبار) - main GUI application.

Single-page desktop app for a storage department:
  - Drag & drop (or browse) the daily Excel export to import it
  - Search products by name or code (case-insensitive)
  - View current location + full location history
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import db
import importer
import fonts

# Optional drag-and-drop support. The app still works without it
# (the browse button always works) - it just won't accept drag & drop
# if tkinterdnd2 isn't installed.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
    BaseTk = TkinterDnD.Tk
except Exception:
    DND_AVAILABLE = False
    BaseTk = tk.Tk

APP_TITLE = "سامانه پیگیری انبار"

# ---------- Color palette ----------
NAVY = "#0F2C4C"
NAVY_DARK = "#0A1E36"
BLUE = "#2F6FED"
BLUE_HOVER = "#2559C4"
WHITE = "#FFFFFF"
BG = "#F4F7FB"
CARD_BG = "#FFFFFF"
BORDER = "#E1E8F0"
TEXT_DARK = "#1B2430"
TEXT_MUTED = "#6B7785"
ROW_ALT = "#F7FAFD"
SUCCESS = "#1E9E64"


class WarehouseApp(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1150x780")
        self.minsize(950, 650)
        self.configure(bg=BG)

        db.init_db()

        self.font_family = fonts.resolve_family(self)
        self._setup_styles()
        self._build_ui()
        self._refresh_upload_log()
        self._search()

    # ---------- Styling ----------

    def _setup_styles(self):
        f = self.font_family
        self.f_normal = (f, 10)
        self.f_bold = (f, 10, "bold")
        self.f_title = (f, 16, "bold")
        self.f_section = (f, 12, "bold")
        self.f_small = (f, 9)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(".", font=self.f_normal, background=BG)
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD_BG)
        style.configure("Header.TFrame", background=NAVY)

        style.configure("Title.TLabel", background=NAVY, foreground=WHITE, font=self.f_title)
        style.configure("Subtitle.TLabel", background=NAVY, foreground="#B9CBEA", font=self.f_small)
        style.configure("Card.TLabel", background=CARD_BG, foreground=TEXT_DARK, font=self.f_normal)
        style.configure("Section.TLabel", background=CARD_BG, foreground=NAVY, font=self.f_section)
        style.configure("Muted.TLabel", background=CARD_BG, foreground=TEXT_MUTED, font=self.f_small)
        style.configure("Status.TLabel", background=NAVY_DARK, foreground=WHITE, font=self.f_small)

        style.configure(
            "Accent.TButton",
            background=BLUE, foreground=WHITE, font=self.f_bold,
            padding=(14, 8), borderwidth=0,
        )
        style.map("Accent.TButton", background=[("active", BLUE_HOVER)])

        style.configure(
            "Secondary.TButton",
            background=WHITE, foreground=NAVY, font=self.f_normal,
            padding=(14, 8), borderwidth=1,
        )
        style.map("Secondary.TButton", background=[("active", "#EAF0FB")])

        style.configure(
            "TEntry", fieldbackground=WHITE, padding=8, font=self.f_normal
        )

        style.configure(
            "Treeview",
            background=WHITE, fieldbackground=WHITE, foreground=TEXT_DARK,
            rowheight=28, font=self.f_normal, borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background=NAVY, foreground=WHITE, font=self.f_bold, padding=(8, 6),
        )
        style.map("Treeview.Heading", background=[("active", NAVY)])
        style.map("Treeview", background=[("selected", BLUE)], foreground=[("selected", WHITE)])

        style.configure("TPanedwindow", background=BG)
        style.configure("Horizontal.TProgressbar", background=BLUE, troughcolor=BORDER)

    # ---------- UI construction ----------

    def _build_ui(self):
        # Header
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x")
        inner = ttk.Frame(header, style="Header.TFrame")
        inner.pack(fill="x", padx=24, pady=16)
        ttk.Label(inner, text=APP_TITLE, style="Title.TLabel").pack(anchor="e")
        ttk.Label(
            inner,
            text="مدیریت و پیگیری موقعیت کالاها در انبار",
            style="Subtitle.TLabel",
        ).pack(anchor="e", pady=(2, 0))

        body = ttk.Frame(self, style="TFrame")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # ----- Upload card -----
        upload_card = ttk.Frame(body, style="Card.TFrame")
        upload_card.pack(fill="x", pady=(0, 14))
        self._build_upload_section(upload_card)

        # ----- Search card -----
        search_card = ttk.Frame(body, style="Card.TFrame")
        search_card.pack(fill="x", pady=(0, 14))
        self._build_search_section(search_card)

        # ----- Results / History / Log (resizable panes) -----
        paned = ttk.PanedWindow(body, orient="vertical")
        paned.pack(fill="both", expand=True)

        results_card = ttk.Frame(paned, style="Card.TFrame")
        history_card = ttk.Frame(paned, style="Card.TFrame")
        log_card = ttk.Frame(paned, style="Card.TFrame")

        paned.add(results_card, weight=3)
        paned.add(history_card, weight=2)
        paned.add(log_card, weight=2)

        self._build_results_section(results_card)
        self._build_history_section(history_card)
        self._build_log_section(log_card)

        # Status bar
        self.status_var = tk.StringVar(value="آماده")
        status_bar = ttk.Label(self, textvariable=self.status_var, style="Status.TLabel", anchor="e", padding=(16, 6))
        status_bar.pack(fill="x", side="bottom")

    def _card_padding(self, card):
        wrapper = ttk.Frame(card, style="Card.TFrame")
        wrapper.pack(fill="both", expand=True, padx=18, pady=14)
        return wrapper

    # --- Upload section ---
    def _build_upload_section(self, card):
        wrapper = self._card_padding(card)

        ttk.Label(wrapper, text="بارگذاری فایل روزانه", style="Section.TLabel").pack(anchor="e")
        ttk.Label(
            wrapper,
            text="فایل اکسل امروز را بکشید و رها کنید، یا با دکمه زیر انتخاب کنید",
            style="Muted.TLabel",
        ).pack(anchor="e", pady=(2, 10))

        content = ttk.Frame(wrapper, style="Card.TFrame")
        content.pack(fill="x")

        self.dropzone = tk.Label(
            content,
            text="📂  فایل اکسل را اینجا رها کنید\nیا کلیک کنید تا انتخاب شود",
            font=self.f_normal,
            bg="#EAF0FB",
            fg=NAVY,
            relief="flat",
            justify="center",
            height=4,
            cursor="hand2",
        )
        self.dropzone.pack(fill="x", ipady=10)
        self.dropzone.bind("<Button-1>", lambda e: self._choose_and_upload())

        if DND_AVAILABLE:
            self.dropzone.drop_target_register(DND_FILES)
            self.dropzone.dnd_bind("<<Drop>>", self._on_drop)

        btn_row = ttk.Frame(wrapper, style="Card.TFrame")
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="انتخاب فایل", style="Accent.TButton", command=self._choose_and_upload).pack(side="left")

        self.upload_progress = ttk.Progressbar(wrapper, mode="determinate")
        self.upload_progress.pack(fill="x", pady=(12, 0))

    def _on_drop(self, event):
        # tkinterdnd2 wraps paths in {} if they contain spaces
        raw = event.data
        paths = self.tk.splitlist(raw)
        if not paths:
            return
        filepath = paths[0]
        if not filepath.lower().endswith((".xlsx", ".xls")):
            messagebox.showerror("خطا", "لطفاً یک فایل اکسل (.xlsx) انتخاب کنید")
            return
        self._start_import(filepath)

    # --- Search section ---
    def _build_search_section(self, card):
        wrapper = self._card_padding(card)

        ttk.Label(wrapper, text="جستجوی کالا", style="Section.TLabel").pack(anchor="e")

        row = ttk.Frame(wrapper, style="Card.TFrame")
        row.pack(fill="x", pady=(10, 0))

        ttk.Button(row, text="پاک کردن", style="Secondary.TButton", command=self._clear_search).pack(side="left", padx=(0, 8))
        ttk.Button(row, text="جستجو", style="Accent.TButton", command=self._search).pack(side="left")

        self.search_var = tk.StringVar()
        entry = ttk.Entry(row, textvariable=self.search_var, font=self.f_normal, justify="right")
        entry.pack(side="left", fill="x", expand=True, padx=8)
        entry.bind("<Return>", lambda e: self._search())
        entry.focus()

        ttk.Label(row, text="نام یا کد کالا:", style="Card.TLabel").pack(side="left", padx=(0, 8))

    # --- Results section ---
    def _build_results_section(self, card):
        wrapper = self._card_padding(card)
        ttk.Label(wrapper, text="نتایج", style="Section.TLabel").pack(anchor="e")

        table_frame = ttk.Frame(wrapper, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True, pady=(8, 0))

        columns = ("code", "name", "model", "location", "last_updated")
        self.results_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        headings = {
            "code": "کد فنی",
            "name": "نام کالا",
            "model": "مدل",
            "location": "موقعیت فعلی",
            "last_updated": "آخرین به‌روزرسانی",
        }
        widths = {"code": 110, "name": 320, "model": 200, "location": 150, "last_updated": 150}
        for col in columns:
            self.results_tree.heading(col, text=headings[col])
            self.results_tree.column(col, width=widths[col], anchor="center")
        self.results_tree.tag_configure("odd", background=ROW_ALT)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=vsb.set)
        self.results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.results_tree.bind("<<TreeviewSelect>>", self._on_select_result)

    # --- History section ---
    def _build_history_section(self, card):
        wrapper = self._card_padding(card)
        self.history_label_var = tk.StringVar(value="تاریخچه موقعیت: (یک کالا را از بالا انتخاب کنید)")
        ttk.Label(wrapper, textvariable=self.history_label_var, style="Section.TLabel").pack(anchor="e")

        table_frame = ttk.Frame(wrapper, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True, pady=(8, 0))

        hist_columns = ("changed_at", "location", "source_file")
        self.history_tree = ttk.Treeview(table_frame, columns=hist_columns, show="headings")
        hist_headings = {
            "changed_at": "تاریخ تغییر",
            "location": "موقعیت",
            "source_file": "فایل منبع",
        }
        for col in hist_columns:
            self.history_tree.heading(col, text=hist_headings[col])
            self.history_tree.column(col, width=200, anchor="center")

        hvsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hvsb.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        hvsb.pack(side="right", fill="y")

    # --- Upload log section ---
    def _build_log_section(self, card):
        wrapper = self._card_padding(card)
        ttk.Label(wrapper, text="تاریخچه بارگذاری‌ها", style="Section.TLabel").pack(anchor="e")

        table_frame = ttk.Frame(wrapper, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True, pady=(8, 0))

        log_columns = ("uploaded_at", "filename", "row_count", "new_parts", "location_changes")
        self.log_tree = ttk.Treeview(table_frame, columns=log_columns, show="headings")
        log_headings = {
            "uploaded_at": "تاریخ/ساعت",
            "filename": "نام فایل",
            "row_count": "تعداد ردیف",
            "new_parts": "کالای جدید",
            "location_changes": "تغییر موقعیت",
        }
        for col in log_columns:
            self.log_tree.heading(col, text=log_headings[col])
            self.log_tree.column(col, width=160, anchor="center")

        lvsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.log_tree.yview)
        self.log_tree.configure(yscrollcommand=lvsb.set)
        self.log_tree.pack(side="left", fill="both", expand=True)
        lvsb.pack(side="right", fill="y")

    # ---------- Behaviors ----------

    def _search(self):
        query = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        for row in self.results_tree.get_children():
            self.results_tree.delete(row)

        results = db.search_parts(query)
        for i, r in enumerate(results):
            tag = "odd" if i % 2 else ""
            self.results_tree.insert(
                "", "end", iid=r["code"], tags=(tag,),
                values=(r["code"], r["name"], r["model"], r["current_location"], r["last_updated"]),
            )
        self.status_var.set(f"{len(results)} کالا یافت شد")

        for row in self.history_tree.get_children():
            self.history_tree.delete(row)
        self.history_label_var.set("تاریخچه موقعیت: (یک کالا را از بالا انتخاب کنید)")

    def _clear_search(self):
        self.search_var.set("")
        self._search()

    def _on_select_result(self, event):
        selection = self.results_tree.selection()
        if not selection:
            return
        code = selection[0]
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)

        history = db.get_history(code)
        for h in history:
            self.history_tree.insert(
                "", "end",
                values=(h["changed_at"], h["location"] or "(خالی)", h["source_file"]),
            )
        name = self.results_tree.set(code, "name")
        self.history_label_var.set(f"تاریخچه موقعیت: {name} ({code}) — {len(history)} رکورد")

    def _choose_and_upload(self):
        filepath = filedialog.askopenfilename(
            title="انتخاب فایل اکسل روزانه",
            filetypes=[("فایل اکسل", "*.xlsx *.xls")],
        )
        if not filepath:
            return
        self._start_import(filepath)

    def _start_import(self, filepath):
        self.upload_progress["value"] = 0
        self.status_var.set("در حال بارگذاری...")

        def progress_callback(current, total):
            if total:
                pct = min(100, int(current / total * 100))
                self.upload_progress.after(0, lambda: self.upload_progress.configure(value=pct))

        def run_import():
            try:
                summary = importer.import_excel_file(filepath, progress_callback=progress_callback)
            except importer.ImportError_ as e:
                self.after(0, lambda: messagebox.showerror("خطا در بارگذاری", str(e)))
                self.after(0, lambda: self.status_var.set("بارگذاری ناموفق بود"))
                return
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("خطای غیرمنتظره", str(e)))
                self.after(0, lambda: self.status_var.set("بارگذاری ناموفق بود"))
                return

            def on_done():
                self.upload_progress["value"] = 100
                msg = (
                    f"تعداد ردیف پردازش‌شده: {summary['row_count']}\n"
                    f"کالای جدید: {summary['new_parts']}\n"
                    f"تغییر موقعیت: {summary['location_changes']}\n"
                    f"ردیف نادیده‌گرفته‌شده (بدون کد): {summary['skipped']}"
                )
                messagebox.showinfo("بارگذاری با موفقیت انجام شد", msg)
                self.status_var.set("بارگذاری با موفقیت انجام شد")
                self._refresh_upload_log()
                self._search()

            self.after(0, on_done)

        threading.Thread(target=run_import, daemon=True).start()

    def _refresh_upload_log(self):
        for row in self.log_tree.get_children():
            self.log_tree.delete(row)
        for entry in db.get_upload_log():
            self.log_tree.insert(
                "", "end",
                values=(
                    entry["uploaded_at"], entry["filename"], entry["row_count"],
                    entry["new_parts"], entry["location_changes"],
                ),
            )


if __name__ == "__main__":
    app = WarehouseApp()
    app.mainloop()
