"""
app.py
Warehouse Tracker - main GUI application.

A simple desktop app for a storage department:
  - Upload the daily Excel export
  - Search products by name or code
  - View current location + full location history
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

import db
import importer

APP_TITLE = "Warehouse Tracker"


class WarehouseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x650")
        self.minsize(800, 500)

        db.init_db()

        self._build_ui()
        self._refresh_upload_log()
        self._search()  # show everything initially

    # ---------- UI construction ----------

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.search_tab = ttk.Frame(notebook)
        self.upload_tab = ttk.Frame(notebook)

        notebook.add(self.search_tab, text="Search")
        notebook.add(self.upload_tab, text="Upload Daily File")

        self._build_search_tab()
        self._build_upload_tab()

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken")
        status_bar.pack(fill="x", side="bottom")

    def _build_search_tab(self):
        frame = self.search_tab

        top = ttk.Frame(frame)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Label(top, text="Search by name or code:").pack(side="left")
        self.search_var = tk.StringVar()
        entry = ttk.Entry(top, textvariable=self.search_var, width=50)
        entry.pack(side="left", padx=8)
        entry.bind("<Return>", lambda e: self._search())
        entry.focus()

        ttk.Button(top, text="Search", command=self._search).pack(side="left")
        ttk.Button(top, text="Clear", command=self._clear_search).pack(side="left", padx=5)

        # Split view: results on top, history on bottom
        paned = ttk.PanedWindow(frame, orient="vertical")
        paned.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Results table
        results_frame = ttk.Frame(paned)
        paned.add(results_frame, weight=2)

        ttk.Label(results_frame, text="Results (click a row to see its history):").pack(anchor="w")

        columns = ("code", "name", "model", "location", "last_updated")
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=12)
        headings = {
            "code": "Code",
            "name": "Product Name",
            "model": "Model",
            "location": "Current Location",
            "last_updated": "Last Updated",
        }
        widths = {"code": 120, "name": 300, "model": 200, "location": 150, "last_updated": 140}
        for col in columns:
            self.results_tree.heading(col, text=headings[col])
            self.results_tree.column(col, width=widths[col], anchor="w")

        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=vsb.set)
        self.results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.results_tree.bind("<<TreeviewSelect>>", self._on_select_result)

        # History table
        history_frame = ttk.Frame(paned)
        paned.add(history_frame, weight=1)

        self.history_label_var = tk.StringVar(value="Location history: (select a product above)")
        ttk.Label(history_frame, textvariable=self.history_label_var).pack(anchor="w")

        hist_columns = ("changed_at", "location", "source_file")
        self.history_tree = ttk.Treeview(history_frame, columns=hist_columns, show="headings", height=8)
        hist_headings = {
            "changed_at": "Date Changed",
            "location": "Location",
            "source_file": "Source File",
        }
        hist_widths = {"changed_at": 160, "location": 200, "source_file": 200}
        for col in hist_columns:
            self.history_tree.heading(col, text=hist_headings[col])
            self.history_tree.column(col, width=hist_widths[col], anchor="w")

        hvsb = ttk.Scrollbar(history_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hvsb.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        hvsb.pack(side="right", fill="y")

    def _build_upload_tab(self):
        frame = self.upload_tab

        top = ttk.Frame(frame)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Label(
            top,
            text="Upload today's Excel export. New products are added, and\n"
                 "any changed locations are recorded automatically in the history.",
            justify="left",
        ).pack(anchor="w")

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(btn_frame, text="Choose Excel File & Upload", command=self._choose_and_upload).pack(side="left")

        self.upload_progress = ttk.Progressbar(frame, mode="determinate")
        self.upload_progress.pack(fill="x", padx=10, pady=10)

        ttk.Label(frame, text="Upload history:").pack(anchor="w", padx=10)

        log_columns = ("uploaded_at", "filename", "row_count", "new_parts", "location_changes")
        self.log_tree = ttk.Treeview(frame, columns=log_columns, show="headings", height=12)
        log_headings = {
            "uploaded_at": "Date/Time",
            "filename": "File",
            "row_count": "Rows Processed",
            "new_parts": "New Products",
            "location_changes": "Location Changes",
        }
        for col in log_columns:
            self.log_tree.heading(col, text=log_headings[col])
            self.log_tree.column(col, width=160, anchor="w")
        self.log_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # ---------- Behaviors ----------

    def _search(self):
        query = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        for row in self.results_tree.get_children():
            self.results_tree.delete(row)

        results = db.search_parts(query)
        for r in results:
            self.results_tree.insert(
                "", "end", iid=r["code"],
                values=(r["code"], r["name"], r["model"], r["current_location"], r["last_updated"]),
            )
        self.status_var.set(f"{len(results)} product(s) found")

        # clear history panel
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)
        self.history_label_var.set("Location history: (select a product above)")

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
                values=(h["changed_at"], h["location"] or "(empty)", h["source_file"]),
            )
        name = self.results_tree.set(code, "name")
        self.history_label_var.set(f"Location history for: {name}  ({code})  \u2014 {len(history)} record(s)")

    def _choose_and_upload(self):
        filepath = filedialog.askopenfilename(
            title="Select the daily Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if not filepath:
            return

        self.upload_progress["value"] = 0
        self.status_var.set("Importing...")

        def progress_callback(current, total):
            if total:
                pct = min(100, int(current / total * 100))
                self.upload_progress.after(0, lambda: self.upload_progress.configure(value=pct))

        def run_import():
            try:
                summary = importer.import_excel_file(filepath, progress_callback=progress_callback)
            except importer.ImportError_ as e:
                self.after(0, lambda: messagebox.showerror("Import Error", str(e)))
                self.after(0, lambda: self.status_var.set("Import failed"))
                return
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Unexpected Error", str(e)))
                self.after(0, lambda: self.status_var.set("Import failed"))
                return

            def on_done():
                self.upload_progress["value"] = 100
                msg = (
                    f"Imported {summary['row_count']} rows.\n"
                    f"New products: {summary['new_parts']}\n"
                    f"Location changes: {summary['location_changes']}\n"
                    f"Skipped (no code): {summary['skipped']}"
                )
                messagebox.showinfo("Import Complete", msg)
                self.status_var.set("Import complete")
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
