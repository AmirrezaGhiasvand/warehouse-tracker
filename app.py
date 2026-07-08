# -*- coding: utf-8 -*-
"""
app.py
Warehouse Tracker (سامانه پیگیری انبار) - main GUI application.

Single-page desktop app for a storage department:
  - Drag & drop (or browse) the daily Excel export to import it
  - Type a product name or code and pick it from live suggestions
  - View its full location history (empty/no-location entries hidden)
"""

import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import persian_date as pdate
import db
import importer
import fonts
import openpyxl

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

MAX_SUGGESTIONS = 8


class WarehouseApp(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1150x880")
        self.minsize(1100, 850)
        self.configure(bg=BG)

        db.init_db()

        self.font_family = fonts.resolve_family(self)
        self._setup_styles()

        self.suggestion_win = None
        self.suggestion_tree = None
        self._suggestion_data = []
        self._current_code = None

        self._build_ui()

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
            padding=(14, 6), borderwidth=0,
        )
        style.map("Accent.TButton", background=[("active", BLUE_HOVER)])

        style.configure(
            "Secondary.TButton",
            background=WHITE, foreground=NAVY, font=self.f_normal,
            padding=(14, 6), borderwidth=1,
        )
        style.map("Secondary.TButton", background=[("active", "#EAF0FB")])

        style.configure("TEntry", fieldbackground=WHITE, padding=8, font=self.f_normal)

        style.configure(
            "Treeview",
            background=WHITE, fieldbackground=WHITE, foreground=TEXT_DARK,
            rowheight=30, font=self.f_normal, borderwidth=0,
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
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x")
        inner = ttk.Frame(header, style="Header.TFrame")
        inner.pack(fill="x", padx=24, pady=12)
        ttk.Label(inner, text=APP_TITLE, style="Title.TLabel").pack(anchor="e")
        ttk.Label(
            inner, text="مدیریت و پیگیری موقعیت کالاها در انبار", style="Subtitle.TLabel"
        ).pack(anchor="e", pady=(2, 0))

        body = ttk.Frame(self, style="TFrame")
        body.pack(fill="both", expand=True, padx=20, pady=14)

        # ----- Upload card (compact) -----
        upload_card = ttk.Frame(body, style="Card.TFrame")
        upload_card.pack(fill="x", pady=(0, 12))
        self._build_upload_section(upload_card)

        
        # ----- Search card -----
        search_card = ttk.Frame(body, style="Card.TFrame")
        search_card.pack(fill="x", pady=(0, 12))
        self._build_search_section(search_card)

       # ----- Bottom area (50/50) -----
        bottom = ttk.Frame(body)
        bottom.pack(fill="both", expand=True)

        bottom.grid_rowconfigure(0, weight=1, minsize=180)  # History
        bottom.grid_rowconfigure(1, weight=2, minsize=320)  # Date Search
        bottom.grid_columnconfigure(0, weight=1)
        # ----- Date search card -----
        date_card = ttk.Frame(bottom, style="Card.TFrame")
        date_card.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        date_card.grid_rowconfigure(0, weight=2)
        date_card.grid_columnconfigure(0, weight=2)

        self._build_date_search_section(date_card)

        self.status_var = tk.StringVar(value="آماده")
        status_bar = ttk.Label(self, textvariable=self.status_var, style="Status.TLabel", anchor="e", padding=(16, 6))
        status_bar.pack(fill="x", side="bottom")
        
        # ----- History card -----
        history_card = ttk.Frame(bottom, style="Card.TFrame")
        history_card.grid(row=0, column=0, sticky="nsew", pady=(0, 6))
        history_card.grid_rowconfigure(0, weight=1)
        history_card.grid_columnconfigure(0, weight=1)

        self._build_history_section(history_card)

        

    def _card_padding(self, card, pady=12):
        wrapper = ttk.Frame(card, style="Card.TFrame")
        wrapper.pack(fill="both", expand=True, padx=18, pady=pady)
        return wrapper

    # --- Upload section (compact) ---
    def _build_upload_section(self, card):
        wrapper = self._card_padding(card, pady=10)

        row = ttk.Frame(wrapper, style="Card.TFrame")
        row.pack(fill="x")

        self.dropzone = tk.Label(
            row,
            text="📂  فایل اکسل روزانه را اینجا رها کنید یا کلیک کنید",
            font=self.f_normal,
            bg="#EAF0FB",
            fg=NAVY,
            relief="flat",
            justify="center",
            cursor="hand2",
        )
        self.dropzone.pack(side="left", fill="x", expand=True, ipady=10)
        self.dropzone.bind("<Button-1>", lambda e: self._choose_and_upload())

        if DND_AVAILABLE:
            self.dropzone.drop_target_register(DND_FILES)
            self.dropzone.dnd_bind("<<Drop>>", self._on_drop)

        ttk.Button(row, text="انتخاب فایل", style="Accent.TButton", command=self._choose_and_upload).pack(
            side="left", padx=(10, 0)
        )

        self.upload_progress = ttk.Progressbar(wrapper, mode="determinate")
        self.upload_progress.pack(fill="x", pady=(8, 0))

    def _on_drop(self, event):
        raw = event.data
        paths = self.tk.splitlist(raw)
        if not paths:
            return
        filepath = paths[0]
        if not filepath.lower().endswith((".xlsx", ".xls")):
            messagebox.showerror("خطا", "لطفاً یک فایل اکسل (.xlsx) انتخاب کنید")
            return
        self._start_import(filepath)

    # --- Search section with live suggestions ---
    def _build_search_section(self, card):
        wrapper = self._card_padding(card, pady=12)

        ttk.Label(wrapper, text="جستجوی کالا", style="Section.TLabel").pack(anchor="e")

        row = ttk.Frame(wrapper, style="Card.TFrame")
        row.pack(fill="x", pady=(10, 0))

        ttk.Button(row, text="پاک کردن", style="Secondary.TButton", command=self._clear_search).pack(
            side="left", padx=(0, 8)
        )

        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(row, textvariable=self.search_var, font=self.f_normal, justify="right")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=8)
        self.search_entry.bind("<KeyRelease>", self._on_search_key)
        self.search_entry.bind("<Down>", self._focus_suggestions)
        self.search_entry.bind("<Return>", self._select_first_suggestion)
        self.search_entry.bind("<Escape>", lambda e: self._hide_suggestions())
        self.search_entry.bind("<FocusOut>", lambda e: self.after(150, self._hide_suggestions))

        # Ctrl+V sometimes doesn't fire on non-English (e.g. Persian) keyboard
        # layouts because Tk's default binding listens for a specific keysym
        # that the layout may not produce. Bind paste explicitly as a fallback.
        self.search_entry.bind("<<Paste>>", self._paste_into_search)
        self.search_entry.bind("<Control-v>", self._paste_into_search)
        self.search_entry.bind("<Control-V>", self._paste_into_search)

        # Right-click context menu as an easy, layout-independent way to paste
        self.search_context_menu = tk.Menu(self, tearoff=0, font=self.f_normal)
        self.search_context_menu.add_command(label="Paste", command=lambda: self._paste_into_search(None))
        self.search_context_menu.add_command(label="Copy", command=lambda: self.search_entry.event_generate("<<Copy>>"))
        self.search_context_menu.add_command(label="Cut", command=lambda: self.search_entry.event_generate("<<Cut>>"))
        self.search_context_menu.add_separator()
        self.search_context_menu.add_command(label="Clear", command=self._clear_search)
        self.search_entry.bind("<Button-3>", self._show_search_context_menu)
        self.search_entry.focus()

        ttk.Label(row, text="نام یا کد کالا:", style="Card.TLabel").pack(side="left", padx=(0, 8))

        self.selected_product_var = tk.StringVar(value="")
        ttk.Label(wrapper, textvariable=self.selected_product_var, style="Muted.TLabel").pack(anchor="e", pady=(6, 0))

    def _paste_into_search(self, event):
        try:
            clip = self.clipboard_get()
        except tk.TclError:
            return "break"
        try:
            if self.search_entry.selection_present():
                self.search_entry.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        self.search_entry.insert("insert", clip)
        self._on_search_key(type("E", (), {"keysym": ""})())
        return "break"

    def _show_search_context_menu(self, event):
        self.search_context_menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def _on_search_key(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        query = self.search_var.get().strip()
        if not query:
            self._hide_suggestions()
            self._clear_history()
            return
        results = db.search_parts(query)[:MAX_SUGGESTIONS]
        self._show_suggestions(results)

    def _show_suggestions(self, results):
        if not results:
            self._hide_suggestions()
            return

        row_height = 34

        if self.suggestion_win is None:
            self.suggestion_win = tk.Toplevel(self)
            self.suggestion_win.overrideredirect(True)
            self.suggestion_win.attributes("-topmost", True)
            # A thin navy frame gives the popup a crisp card-like border
            # instead of the default flat system border.
            border = tk.Frame(self.suggestion_win, bg=NAVY)
            border.pack(fill="both", expand=True)

            # Using two real columns (rather than one string with the name
            # and code mashed together) avoids any Persian/Latin bidi text
            # mixing issues entirely, and reuses the app's existing table style.
            self.suggestion_tree = ttk.Treeview(
                border, columns=("name", "code"), show="headings", height=len(results)
            )
            self.suggestion_tree.heading("name", text="نام کالا")
            self.suggestion_tree.heading("code", text="کد فنی")
            self.suggestion_tree.column("name", anchor="center", width=260)
            self.suggestion_tree.column("code", anchor="center", width=160)
            self.suggestion_tree.tag_configure("odd", background=ROW_ALT)
            self.suggestion_tree.pack(fill="both", expand=True, padx=1, pady=1)

            self.suggestion_tree.bind("<ButtonRelease-1>", self._on_suggestion_pick)
            self.suggestion_tree.bind("<Return>", self._on_suggestion_pick)
        else:
            for iid in self.suggestion_tree.get_children():
                self.suggestion_tree.delete(iid)
            self.suggestion_tree.configure(height=len(results))

        self._suggestion_data = results
        for i, r in enumerate(results):
            tag = "odd" if i % 2 else ""
            self.suggestion_tree.insert("", "end", iid=str(i), tags=(tag,), values=(r["name"], r["code"]))

        self.update_idletasks()
        x = self.search_entry.winfo_rootx()
        y = self.search_entry.winfo_rooty() + self.search_entry.winfo_height()
        w = self.search_entry.winfo_width()
        header_height = 30
        self.suggestion_win.geometry(f"{w}x{len(results) * row_height + header_height + 2}+{x}+{y}")
        self.suggestion_win.deiconify()
        self.suggestion_win.lift()

    def _hide_suggestions(self):
        if self.suggestion_win is not None:
            self.suggestion_win.withdraw()

    def _focus_suggestions(self, event):
        if self.suggestion_win is not None and self.suggestion_tree.get_children():
            first = self.suggestion_tree.get_children()[0]
            self.suggestion_tree.focus_set()
            self.suggestion_tree.selection_set(first)
            self.suggestion_tree.focus(first)

    def _select_first_suggestion(self, event):
        if self.suggestion_win is not None and self.suggestion_tree.get_children():
            self._load_product(self._suggestion_data[0])
            self._hide_suggestions()

    def _on_suggestion_pick(self, event):
        selection = self.suggestion_tree.selection()
        if not selection:
            return
        idx = int(selection[0])
        self._load_product(self._suggestion_data[idx])
        self._hide_suggestions()

    def _load_product(self, part_row):
        self._current_code = part_row["code"]
        self.search_var.set(part_row["name"])
        self.selected_product_var.set(
            f"کد: {part_row['code']}    |    مدل: {part_row['model'] or '—'}    |    موقعیت فعلی: {part_row['current_location'] or 'بدون موقعیت'}"
        )
        self._load_history(part_row["code"], part_row["name"])

    def _clear_search(self):
        self.search_var.set("")
        self.selected_product_var.set("")
        self._current_code = None
        self._hide_suggestions()
        self._clear_history()

    
    # --- History section (now the main focus of the app) ---
    def _build_history_section(self, card):
        wrapper = self._card_padding(card)
        self.history_label_var = tk.StringVar(value="تاریخچه موقعیت: (یک کالا را جستجو کنید)")
        ttk.Label(wrapper, textvariable=self.history_label_var, style="Section.TLabel").pack(anchor="e")

        table_frame = ttk.Frame(wrapper, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True, pady=(10, 0))

        hist_columns = ("changed_at", "location")
        self.history_tree = ttk.Treeview(table_frame, columns=hist_columns, show="headings", height=12)
        hist_headings = {
            "changed_at": "تاریخ تغییر",
            "location": "موقعیت",
        }
        widths = {"changed_at": 280, "location": 400}
        for col in hist_columns:
            self.history_tree.heading(col, text=hist_headings[col])
            self.history_tree.column(col, width=widths[col], anchor="center")
        self.history_tree.tag_configure("odd", background=ROW_ALT)

        hvsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hvsb.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        hvsb.pack(side="right", fill="y")

    def _load_history(self, code, name):
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)

        history = db.get_history(code)
        for i, h in enumerate(history):
            tag = "odd" if i % 2 else ""
            self.history_tree.insert(
                "", "end", tags=(tag,),
                values=(h["changed_at"], h["location"]),
            )

        if history:
            self.history_label_var.set(f"تاریخچه موقعیت: {name} ({code}) — {len(history)} رکورد")
        else:
            self.history_label_var.set(f"تاریخچه موقعیت: {name} ({code}) — هیچ سابقه‌ای ثبت نشده است")

    def _clear_history(self):
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)
        self.history_label_var.set("تاریخچه موقعیت: (یک کالا را جستجو کنید)")


    # --- Date search section (Persian calendar) ---
  
    def _build_date_search_section(self, card):
        wrapper = self._card_padding(card, pady=2)

        ttk.Label(wrapper, text="جستجو بر اساس تاریخ", style="Section.TLabel").pack(anchor="e")
        ttk.Label(
            wrapper,
            text="یک تاریخ را انتخاب کنید تا کالاهای جدید و تغییرات موقعیت آن روز را ببینید",
            style="Muted.TLabel",
        ).pack(anchor="e", pady=(2, 8))

        today = pdate.today_jalali()
        self.selected_jalali_year = today.year
        self.selected_jalali_month = today.month
        self.selected_jalali_day = today.day
        self._picker_year = today.year
        self._picker_month = today.month
        self._calendar_win = None

        row = ttk.Frame(wrapper, style="Card.TFrame")
        row.pack(fill="x")

        ttk.Button(row, text="امروز", style="Secondary.TButton", command=self._jump_to_today).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(row, text="خروجی اکسل", style="Secondary.TButton", command=self._export_date_results).pack(
            side="left", padx=(8, 0)
        )

        self.date_display_var = tk.StringVar()
        self._refresh_date_display()
        self.date_picker_btn = tk.Button(
            row, textvariable=self.date_display_var, font=self.f_bold,
            bg="#EAF0FB", fg=NAVY, relief="flat", bd=0, padx=14, pady=6,
            cursor="hand2", activebackground="#DCE7FA", activeforeground=NAVY,
            command=self._toggle_calendar,
        )
        self.date_picker_btn.pack(side="left", padx=8)

        ttk.Label(row, text="تاریخ:", style="Card.TLabel").pack(side="left", padx=(0, 8))

        self.date_results_label_var = tk.StringVar(value="")
        ttk.Label(wrapper, textvariable=self.date_results_label_var, style="Muted.TLabel").pack(anchor="e", pady=(8, 4))

        table_frame = ttk.Frame(wrapper, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True)

        columns = ("type", "name", "code", "previous", "location")
        self.date_results_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=12)
        headings = {
            "type": "نوع", "name": "نام کالا", "code": "کد فنی",
            "previous": "موقعیت قبلی", "location": "موقعیت جدید",
        }
        widths = {"type": 110, "name": 260, "code": 120, "previous": 150, "location": 150}
        for col in columns:
            self.date_results_tree.heading(col, text=headings[col])
            self.date_results_tree.column(col, width=widths[col], anchor="center")
        self.date_results_tree.tag_configure("odd", background=ROW_ALT)

        dvsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.date_results_tree.yview)
        self.date_results_tree.configure(yscrollcommand=dvsb.set)
        self.date_results_tree.pack(side="left", fill="both", expand=True)
        dvsb.pack(side="right", fill="y")

        self._search_by_date()

    def _refresh_date_display(self):
        self.date_display_var.set(
            f"\U0001F4C5  {self.selected_jalali_day} {pdate.PERSIAN_MONTHS[self.selected_jalali_month - 1]} {self.selected_jalali_year}"
        )

    def _jump_to_today(self):
        today = pdate.today_jalali()
        self.selected_jalali_year = today.year
        self.selected_jalali_month = today.month
        self.selected_jalali_day = today.day
        self._refresh_date_display()
        self._search_by_date()

    def _toggle_calendar(self):
        if self._calendar_win is not None:
            self._close_calendar()
        else:
            self._open_calendar()

    def _open_calendar(self):
        self._picker_year = self.selected_jalali_year
        self._picker_month = self.selected_jalali_month

        self._calendar_win = tk.Toplevel(self)
        self._calendar_win.overrideredirect(True)
        self._calendar_win.attributes("-topmost", True)
        border = tk.Frame(self._calendar_win, bg=NAVY)
        border.pack(fill="both", expand=True)
        self._calendar_inner = tk.Frame(border, bg=WHITE)
        self._calendar_inner.pack(fill="both", expand=True, padx=1, pady=1)

        self._render_calendar()

        self.update_idletasks()
        x = self.date_picker_btn.winfo_rootx()
        y = self.date_picker_btn.winfo_rooty() + self.date_picker_btn.winfo_height() + 2
        self._calendar_win.geometry(f"+{x}+{y}")

        self._calendar_win.bind("<FocusOut>", lambda e: self.after(150, self._close_calendar))
        self._calendar_win.focus_set()

    def _close_calendar(self):
        if self._calendar_win is not None:
            self._calendar_win.destroy()
            self._calendar_win = None

    def _render_calendar(self):
        for widget in self._calendar_inner.winfo_children():
            widget.destroy()

        header = tk.Frame(self._calendar_inner, bg=NAVY)
        header.pack(fill="x")
        tk.Button(
            header, text="ماه بعد \u25B6", font=self.f_small, bg=NAVY, fg=WHITE, relief="flat",
            activebackground=BLUE_HOVER, activeforeground=WHITE, bd=0, cursor="hand2",
            command=self._picker_next_month,
        ).pack(side="left", padx=6, pady=6)
        tk.Label(
            header, text=f"{pdate.PERSIAN_MONTHS[self._picker_month - 1]} {self._picker_year}",
            font=self.f_bold, bg=NAVY, fg=WHITE,
        ).pack(side="left", expand=True, pady=6)
        tk.Button(
            header, text="\u25C0 ماه قبل", font=self.f_small, bg=NAVY, fg=WHITE, relief="flat",
            activebackground=BLUE_HOVER, activeforeground=WHITE, bd=0, cursor="hand2",
            command=self._picker_prev_month,
        ).pack(side="right", padx=6, pady=6)

        grid = tk.Frame(self._calendar_inner, bg=WHITE)
        grid.pack(padx=8, pady=8)

        weekday_labels_by_display_col = ["ج", "پ", "چ", "س", "د", "ی", "ش"]
        for col, wd in enumerate(weekday_labels_by_display_col):
            tk.Label(grid, text=wd, font=self.f_bold, bg=WHITE, fg=TEXT_MUTED, width=4).grid(
                row=0, column=col, pady=(0, 2)
            )

        first_weekday = pdate.first_weekday_of_month(self._picker_year, self._picker_month)
        days_count = pdate.days_in_month(self._picker_year, self._picker_month)
        today = pdate.today_jalali()

        row_i = 1
        col_i = 6 - first_weekday
        for day in range(1, days_count + 1):
            is_today = (
                self._picker_year == today.year and self._picker_month == today.month and day == today.day
            )
            is_selected = (
                self._picker_year == self.selected_jalali_year
                and self._picker_month == self.selected_jalali_month
                and day == self.selected_jalali_day
            )

            if is_selected:
                bg, fg = BLUE, WHITE
            elif is_today:
                bg, fg = "#EAF0FB", NAVY
            else:
                bg, fg = WHITE, TEXT_DARK

            cell = tk.Label(grid, text=str(day), font=self.f_normal, bg=bg, fg=fg, width=4, height=1, cursor="hand2")
            cell.grid(row=row_i, column=col_i, padx=1, pady=1)
            cell.bind("<Button-1>", lambda e, d=day: self._pick_day(d))

            col_i -= 1
            if col_i < 0:
                col_i = 6
                row_i += 1

    def _picker_next_month(self):
        self._picker_month += 1
        if self._picker_month > 12:
            self._picker_month = 1
            self._picker_year += 1
        self._render_calendar()

    def _picker_prev_month(self):
        self._picker_month -= 1
        if self._picker_month < 1:
            self._picker_month = 12
            self._picker_year -= 1
        self._render_calendar()

    def _pick_day(self, day):
        self.selected_jalali_year = self._picker_year
        self.selected_jalali_month = self._picker_month
        self.selected_jalali_day = day
        self._refresh_date_display()
        self._close_calendar()
        self._search_by_date()

    def _search_by_date(self):
        year = self.selected_jalali_year
        month = self.selected_jalali_month
        day = self.selected_jalali_day
        gregorian_str = pdate.jalali_to_gregorian_str(year, month, day)

        for row in self.date_results_tree.get_children():
            self.date_results_tree.delete(row)

        new_parts = db.get_new_parts_by_date(gregorian_str)
        changes = db.get_location_changes_by_date(gregorian_str)

        i = 0
        for p in new_parts:
            tag = "odd" if i % 2 else ""
            self.date_results_tree.insert(
                "", "end", tags=(tag,),
                values=(
                "کالای جدید",
                p["name"],
                p["code"],
                "—",
                p["location"] or "—",
            ),
            )
            i += 1
        for c in changes:
            tag = "odd" if i % 2 else ""
            self.date_results_tree.insert(
                "", "end", tags=(tag,),
                values=(
                "تغییر موقعیت",
                c["name"],
                c["code"],
                c["previous_location"] or "—",
                c["location"],
            ),
            )
            i += 1

        total = len(new_parts) + len(changes)
        if total == 0:
            self.date_results_label_var.set("هیچ موردی در این تاریخ ثبت نشده است")
        else:
            self.date_results_label_var.set(
                f"کالای جدید: {len(new_parts)}        تغییر موقعیت: {len(changes)}"
            )
    def _export_date_results(self):
        rows = self.date_results_tree.get_children()
        if not rows:
            messagebox.showinfo("خروجی اکسل", "نتیجه‌ای برای خروجی گرفتن وجود ندارد. ابتدا جستجو کنید.")
            return

        year = self.selected_jalali_year
        month_name = pdate.PERSIAN_MONTHS[self.selected_jalali_month - 1]
        day = self.selected_jalali_day
        default_name = f"گزارش {day} {month_name} {year}.xlsx"

        filepath = filedialog.asksaveasfilename(
            title="ذخیره خروجی اکسل",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("فایل اکسل", "*.xlsx")],
        )
        if not filepath:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "گزارش"
        ws.append(["نوع", "نام کالا", "کد فنی", "موقعیت قبلی", "موقعیت جدید"])
        for iid in rows:
            values = self.date_results_tree.item(iid)["values"]
            ws.append(list(values))

        column_letters = ["A", "B", "C", "D", "E"]
        widths = [14, 32, 16, 18, 18]
        for letter, w in zip(column_letters, widths):
            ws.column_dimensions[letter].width = w

        try:
            wb.save(filepath)
        except Exception as e:
            messagebox.showerror("خطا", f"ذخیره فایل با خطا مواجه شد:\n{e}")
            return

        messagebox.showinfo("خروجی اکسل", f"فایل با موفقیت ذخیره شد:\n{filepath}")



    # ---------- Upload behavior ----------

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
                # refresh currently viewed product's history, if any
                if self._current_code:
                    part = db.get_part(self._current_code)
                    if part:
                        self._load_product(part)

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