import os
import re
import sys
import time
import queue
import datetime
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
from dbfread import DBF
from tkcalendar import DateEntry

try:
    sys.setswitchinterval(0.001)
except (AttributeError, ValueError):
    pass

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

DBF_COLS = ["ALVNBR", "ALVDATE", "ACCTNBR", "CUSTCODE", "CUSTNAME", "ALVAMT"]
ACCENT_BLUE = "#1E90FF"
ACCENT_GREEN = "#2FA572"
ACCENT_ORANGE = "#E59400"
DEFAULT_ROW_BG = "#2b2b2b"
MATCH_ROW_BG = "#1a3a1a"
ROW_TEXT = "#d0d0d0"
MATCH_ROW_TEXT = "#aaffaa"


def clean_illegal_chars(val):
    if isinstance(val, str):
        return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', val).strip()
    return val


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


PAGE_SIZE = 250
MIN_LOADING_TIME = 1.0


def compute_positions(table, q):
    """Return the row indices of table.source_df matching the search query.

    Runs off the UI thread; builds (and caches) a lower-cased concatenated
    key string per row using vectorized pandas ops instead of row-wise lambdas.
    """
    src = table.source_df
    n = len(src)
    if n == 0:
        return []
    q = (q or "").strip().lower()
    if not q:
        return list(range(n))
    keys = table._keys
    if keys is None:
        parts = [src[c].astype(str).str.lower() for c in src.columns]
        keys = parts[0]
        for p in parts[1:]:
            keys = keys + " " + p
        table._keys = keys
    mask = keys.str.contains(q, na=False, regex=False).to_numpy()
    return [i for i in range(n) if mask[i]]


class ScrollableTable(ctk.CTkFrame):
    """Scrollable, paged table with header, debounced async search, and row highlighting.

    Only a bounded window of rows (PAGE_SIZE) is rendered at a time. All heavy work
    (DataFrame prep, search-key building, filtering) runs off the UI thread via the
    app's background queue so the window stays responsive even with very large data.
    """

    def __init__(self, master, columns=None, app=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app
        self.columns = columns or []
        self.source_df = pd.DataFrame()
        self.highlight_col = None
        self._keys = None
        self._filtered = []
        self.page = 0
        self.page_size = PAGE_SIZE
        self.row_widgets = []
        self._filter_after = None
        self._filter_seq = 0
        self._computing = False

        search_frame = ctk.CTkFrame(self, fg_color="transparent")
        search_frame.pack(fill="x", padx=5, pady=(5, 2))
        ctk.CTkLabel(search_frame, text="Search:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_changed)
        self.ent_search = ctk.CTkEntry(search_frame, textvariable=self.search_var,
                                       placeholder_text="Filter rows...")
        self.ent_search.pack(side="left", fill="x", expand=True)

        self.head_frame = ctk.CTkFrame(self, fg_color="#3a3a3a", corner_radius=4)
        self.head_frame.pack(fill="x", padx=5, pady=(0, 2))

        self.header_labels = []
        self._build_header()

        self.table_frame = ctk.CTkScrollableFrame(self, fg_color="#2b2b2b", corner_radius=6)
        self.table_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=5, pady=(0, 4))
        self.btn_prev = ctk.CTkButton(footer, text="\u25C0 Prev", width=70, height=26,
                                      font=ctk.CTkFont(size=11), fg_color="#333", hover_color="#555",
                                      command=self._prev_page, state="disabled")
        self.btn_prev.pack(side="left", padx=(0, 5))
        self.result_count_label = ctk.CTkLabel(footer, text="0 of 0 rows",
                                               font=ctk.CTkFont(size=10), text_color="gray")
        self.result_count_label.pack(side="left", expand=True)
        self.btn_next = ctk.CTkButton(footer, text="Next \u25B6", width=70, height=26,
                                      font=ctk.CTkFont(size=11), fg_color="#333", hover_color="#555",
                                      command=self._next_page, state="disabled")
        self.btn_next.pack(side="right", padx=(5, 0))

    def _build_header(self):
        for lbl in self.header_labels:
            lbl.destroy()
        self.header_labels.clear()
        for col in self.columns:
            lbl = ctk.CTkLabel(self.head_frame, text=str(col), font=ctk.CTkFont(size=11, weight="bold"),
                               anchor="w", text_color="#e0e0e0")
            lbl.pack(side="left", padx=5, pady=4, expand=True, fill="x")
            self.header_labels.append(lbl)

    def set_columns(self, columns):
        self.columns = list(columns)
        self._build_header()
        self.clear()

    def set_data(self, df, highlight_col=None):
        self._computing = True
        self._filtered = []
        self.page = 0
        self.highlight_col = highlight_col
        self.source_df = df if (df is not None and not df.empty) else pd.DataFrame()
        self._keys = None
        self._set_loading_text()
        if self.source_df.empty:
            self._computing = False
            self._render_page()
            return
        self._filter_seq += 1
        seq = self._filter_seq
        q = self.search_var.get().strip().lower()
        if self.app is not None:
            self.app.start_table_compute(self, q, seq)
        else:
            self._filtered = compute_positions(self, q)
            self._computing = False
            self._render_page()

    def _set_loading_text(self):
        self.result_count_label.configure(text="Preparing data\u2026")
        self.btn_prev.configure(state="disabled")
        self.btn_next.configure(state="disabled")

    def _on_search_changed(self, *_):
        if self._filter_after:
            self.after_cancel(self._filter_after)
        self._filter_after = self.after(250, self._dispatch_search)

    def _dispatch_search(self):
        if self.source_df.empty:
            return
        self._computing = True
        self._set_loading_text()
        self._filter_seq += 1
        seq = self._filter_seq
        q = self.search_var.get().strip().lower()
        if self.app is not None:
            self.app.start_table_compute(self, q, seq)
        else:
            self._filtered = compute_positions(self, q)
            self._computing = False
            self._render_page()

    def apply_compute(self, positions, seq, err):
        if not self.winfo_exists():
            return
        if seq != self._filter_seq:
            return
        self._computing = False
        self.page = 0
        if err:
            self.result_count_label.configure(text=f"Search error: {err}")
            return
        self._filtered = positions
        self._render_page()

    def _invalidate(self):
        if self._filter_after:
            try:
                self.after_cancel(self._filter_after)
            except Exception:
                pass
            self._filter_after = None
        self._filter_seq += 1
        self._computing = False

    def _render_page(self):
        for w in self.row_widgets:
            w.destroy()
        self.row_widgets.clear()

        total = len(self._filtered)
        if total == 0:
            self.result_count_label.configure(text="0 of 0 rows")
            self.btn_prev.configure(state="disabled")
            self.btn_next.configure(state="disabled")
            return

        page_count = (total + self.page_size - 1) // self.page_size
        if self.page >= page_count:
            self.page = page_count - 1
        lo = self.page * self.page_size
        hi = min(lo + self.page_size, total)

        for pos in self._filtered[lo:hi]:
            self._append_row(self._row_to_dict(pos))

        count_text = f"Showing {lo + 1:,}\u2013{hi:,} of {total:,} row" + ("s" if total != 1 else "")
        if self.search_var.get().strip():
            count_text += " (filtered)"
        self.result_count_label.configure(text=count_text)
        self.btn_prev.configure(state="disabled" if self.page <= 0 else "normal")
        self.btn_next.configure(state="disabled" if self.page >= page_count - 1 else "normal")

    def _row_to_dict(self, pos):
        row = self.source_df.iloc[pos]
        row_data = {}
        for col in self.columns:
            val = row.get(col, "")
            if pd.notna(val):
                if isinstance(val, pd.Timestamp):
                    val = val.strftime("%Y-%m-%d")
                else:
                    val = str(val)
            else:
                val = ""
            row_data[col] = val
        if self.highlight_col and self.highlight_col in row.index:
            v = row[self.highlight_col]
            row_data["_highlight"] = bool(v) if (isinstance(v, bool) or pd.notna(v)) else False
        else:
            row_data["_highlight"] = False
        return row_data

    def _append_row(self, row_data):
        is_match = row_data.get("_highlight", False)
        bg = MATCH_ROW_BG if is_match else DEFAULT_ROW_BG
        fg = MATCH_ROW_TEXT if is_match else ROW_TEXT
        row_frame = tk.Frame(self.table_frame, bg=bg)
        row_frame.pack(fill="x", padx=2, pady=1)
        for col in self.columns:
            val = row_data.get(col, "")
            lbl = tk.Label(row_frame, text=val, bd=0, bg=bg, fg=fg,
                           font=("Segoe UI", 11), anchor="w")
            lbl.pack(side="left", padx=5, pady=2, expand=True, fill="x")
        self.row_widgets.append(row_frame)

    def _prev_page(self):
        if not self._computing and self.page > 0:
            self.page -= 1
            self._render_page()

    def _next_page(self):
        if self._computing:
            return
        page_count = max(1, (len(self._filtered) + self.page_size - 1) // self.page_size)
        if self.page < page_count - 1:
            self.page += 1
            self._render_page()

    def clear(self):
        self.source_df = pd.DataFrame()
        self.highlight_col = None
        self._keys = None
        self._filtered = []
        self.page = 0
        self._invalidate()
        for w in self.row_widgets:
            w.destroy()
        self.row_widgets.clear()
        self.result_count_label.configure(text="0 of 0 rows")
        self.btn_prev.configure(state="disabled")
        self.btn_next.configure(state="disabled")
        self.search_var.set("")


class StepIndicator(ctk.CTkFrame):
    def __init__(self, master, number, title, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.number = number
        self.badge = ctk.CTkLabel(self, text=str(number), width=30, height=30,
                                  corner_radius=15, fg_color="#555555", text_color="gray",
                                  font=ctk.CTkFont(size=13, weight="bold"))
        self.badge.pack(side="left", padx=(0, 8))
        self.title_lbl = ctk.CTkLabel(self, text=title, font=ctk.CTkFont(size=12, weight="bold"),
                                      text_color="gray")
        self.title_lbl.pack(side="left")
        self.set_state("pending")

    def set_state(self, state):
        if state == "done":
            self.badge.configure(text="\u2713", fg_color=ACCENT_GREEN, text_color="white")
            self.title_lbl.configure(text_color="white")
        elif state == "active":
            self.badge.configure(text=str(self.number), fg_color=ACCENT_BLUE, text_color="white")
            self.title_lbl.configure(text_color="white")
        else:
            self.badge.configure(text=str(self.number), fg_color="#555555", text_color="gray")
            self.title_lbl.configure(text_color="gray")


class StepProgressBar(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, corner_radius=10, **kwargs)
        self.steps = [
            StepIndicator(self, 1, "Load DBF"),
            StepIndicator(self, 2, "Upload Excel"),
            StepIndicator(self, 3, "Results"),
        ]
        for i, ind in enumerate(self.steps):
            ind.pack(side="left", padx=10, pady=8)
            if i < len(self.steps) - 1:
                ctk.CTkLabel(self, text="\u2192", font=ctk.CTkFont(size=16),
                             text_color="gray").pack(side="left")

    def set_status(self, current):
        for idx in range(3):
            if idx + 1 < current:
                self.steps[idx].set_state("done")
            elif idx + 1 == current:
                self.steps[idx].set_state("active")
            else:
                self.steps[idx].set_state("pending")


class WelcomeStep(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color="transparent")
        self.app = app

        hero = ctk.CTkFrame(self, fg_color="transparent")
        hero.pack(fill="x", pady=(15, 5))
        ctk.CTkLabel(hero, text="DBF Data Extractor & Excel Validator By Kuya Daks",
                     font=ctk.CTkFont(size=28, weight="bold")).pack(pady=(10, 2))
        ctk.CTkLabel(hero, text="Validate DBF records against your Excel file in three simple steps.",
                     font=ctk.CTkFont(size=14), text_color="gray").pack()

        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.pack(fill="x", padx=30, pady=10)
        howto = [
            ("1", "Select DBF File", "Choose the .dbf source file and set the date range."),
            ("2", "Upload Excel File", "Pick the .xlsx / .xls target file to validate against."),
            ("3", "Review Results", "See matched records side-by-side with a summary dashboard."),
        ]
        for num, t, d in howto:
            c = ctk.CTkFrame(cards, corner_radius=10, width=280)
            c.pack(side="left", expand=True, fill="both", padx=8)
            c.pack_propagate(False)
            ctk.CTkLabel(c, text=num, font=ctk.CTkFont(size=22, weight="bold"),
                         text_color=ACCENT_BLUE).pack(pady=(14, 0))
            ctk.CTkLabel(c, text=t, font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(2, 0))
            ctk.CTkLabel(c, text=d, font=ctk.CTkFont(size=11), text_color="gray",
                         wraplength=230, justify="center").pack(pady=(0, 12))

        cfg = ctk.CTkFrame(self, corner_radius=12)
        cfg.pack(fill="x", padx=60, pady=10)

        ctk.CTkLabel(cfg, text="Load DBF File & Set Date Range",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=20, pady=(12, 4))

        file_row = ctk.CTkFrame(cfg, fg_color="transparent")
        file_row.pack(fill="x", padx=20, pady=6)
        self.btn_browse = ctk.CTkButton(file_row, text="Select DBF File",
                                        command=self.browse_dbf, fg_color="#1f538d")
        self.btn_browse.pack(side="left", padx=(0, 10))
        self.lbl_dbf = ctk.CTkLabel(file_row, text="No DBF file selected", text_color="gray")
        self.lbl_dbf.pack(side="left")

        date_row = ctk.CTkFrame(cfg, fg_color="transparent")
        date_row.pack(fill="x", padx=20, pady=6)
        ctk.CTkLabel(date_row, text="Start Date:").pack(side="left", padx=(0, 5))
        today = datetime.date.today()
        one_year_ago = today.replace(year=today.year - 1)
        self.cal_start = DateEntry(date_row, width=12, background="#1f538d", foreground="white",
                                   borderwidth=2, date_pattern="yyyy-mm-dd",
                                   year=one_year_ago.year, month=one_year_ago.month,
                                   day=one_year_ago.day, font=ctk.CTkFont(size=11))
        self.cal_start.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(date_row, text="End Date:").pack(side="left", padx=(0, 5))
        self.cal_end = DateEntry(date_row, width=12, background="#1f538d", foreground="white",
                                 borderwidth=2, date_pattern="yyyy-mm-dd",
                                 year=today.year, month=today.month, day=today.day,
                                 font=ctk.CTkFont(size=11))
        self.cal_end.pack(side="left")

        self.btn_continue = ctk.CTkButton(cfg, text="Continue to Step 2  \u2192", height=34,
                                          fg_color=ACCENT_GREEN, hover_color="darkgreen",
                                          command=self.on_continue)
        self.btn_continue.pack(pady=(10, 6))

        self.lbl_status = ctk.CTkLabel(cfg, text="", font=ctk.CTkFont(size=11), text_color="gray")
        self.lbl_status.pack()

        self.progress = ctk.CTkProgressBar(cfg, height=14, mode="indeterminate")
        self.progress.set(0)

        self.buttons = [self.btn_browse, self.btn_continue]

    def browse_dbf(self):
        path = filedialog.askopenfilename(filetypes=[("DBF Files", "*.dbf"), ("All Files", "*.*")])
        if path:
            self.app.dbf_path = path
            self.lbl_dbf.configure(text=os.path.basename(path), text_color="white")
            self.app.log_message(f"Selected DBF File: {path}")

    def on_continue(self):
        if not self.app.dbf_path:
            messagebox.showwarning("Warning", "Please select a DBF file first.")
            return
        if self.cal_start.get_date() > self.cal_end.get_date():
            messagebox.showwarning("Warning", "Start date cannot be after the end date.")
            return
        self.app.proceed_to_excel(self.cal_start.get_date(), self.cal_end.get_date())

    def show_loading(self, message):
        self.lbl_status.configure(text=message)
        self.progress.pack(fill="x", padx=20, pady=(4, 10))
        self.progress.start()
        self.update_idletasks()
        for b in self.buttons:
            b.configure(state="disabled")

    def hide_loading(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.progress.set(0)
        self.lbl_status.configure(text="")
        for b in self.buttons:
            b.configure(state="normal")


class ExcelStep(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color="transparent")
        self.app = app

        ctk.CTkLabel(self, text="Upload Excel File",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(15, 2))
        ctk.CTkLabel(self,
                     text="Step 2 of 3 \u2014 choose the Excel file to validate the DBF records against.",
                     font=ctk.CTkFont(size=13), text_color="gray").pack()

        recap = ctk.CTkFrame(self, corner_radius=12)
        recap.pack(fill="x", padx=60, pady=8)
        ctk.CTkLabel(recap, text="Step 1 Recap",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=20, pady=(12, 4))
        self.lbl_recap = ctk.CTkLabel(recap, text="", font=ctk.CTkFont(size=12), text_color="gray",
                                      justify="left", anchor="w")
        self.lbl_recap.pack(anchor="w", padx=20, pady=(0, 12))

        cfg = ctk.CTkFrame(self, corner_radius=12)
        cfg.pack(fill="x", padx=60, pady=8)
        ctk.CTkLabel(cfg, text="Excel File",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=20, pady=(12, 4))
        row = ctk.CTkFrame(cfg, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=6)
        self.btn_browse = ctk.CTkButton(row, text="Select Excel File",
                                        command=self.browse_excel, fg_color="#1f538d")
        self.btn_browse.pack(side="left", padx=(0, 10))
        self.lbl_excel = ctk.CTkLabel(row, text="No Excel file selected", text_color="gray")
        self.lbl_excel.pack(side="left")

        self.lbl_account = ctk.CTkLabel(cfg, text="", font=ctk.CTkFont(size=12),
                                        text_color=ACCENT_GREEN)
        self.lbl_account.pack(anchor="w", padx=20, pady=(0, 4))

        self.lbl_status = ctk.CTkLabel(cfg, text="", font=ctk.CTkFont(size=11), text_color="gray")
        self.lbl_status.pack()
        self.progress = ctk.CTkProgressBar(cfg, height=14, mode="indeterminate")
        self.progress.set(0)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=60, pady=8)
        self.btn_back = ctk.CTkButton(btn_row, text="\u2190 Back", width=120,
                                      fg_color="#333", hover_color="#555", command=self.on_back)
        self.btn_back.pack(side="left")
        self.btn_run = ctk.CTkButton(btn_row, text="Run Validation & View Results  \u2192",
                                     height=34, fg_color=ACCENT_GREEN, hover_color="darkgreen",
                                     command=self.on_run)
        self.btn_run.pack(side="right")

        self.buttons = [self.btn_browse, self.btn_back, self.btn_run]
        self.refresh_recap()

    def refresh_recap(self):
        app = self.app
        start = getattr(app, "start_date", None)
        end = getattr(app, "end_date", None)
        dbf = os.path.basename(app.dbf_path) if app.dbf_path else "-"
        records = f"{app.total_records:,}" if app.total_records else "0"
        self.lbl_recap.configure(text=f"DBF File : {dbf}\nDate Range : {start}  to  {end}\n"
                                      f"Records Extracted : {records}")

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")])
        if path:
            self.app.excel_path = path
            self.lbl_excel.configure(text=os.path.basename(path), text_color="white")
            self.lbl_account.configure(text="")
            self.app.probe_excel(self)

    def on_excel_ready(self, acct_col, num_rows):
        if acct_col:
            self.lbl_account.configure(text=f"\u2713 Account column detected: {acct_col} ({num_rows:,} rows)")
        else:
            self.lbl_account.configure(text="Account column not found \u2014 matcher may not work.",
                                       text_color=ACCENT_ORANGE)

    def on_back(self):
        self.app.show_step(1)

    def on_run(self):
        if not self.app.excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        self.app.proceed_to_results()

    def show_loading(self, message):
        self.lbl_status.configure(text=message)
        self.progress.pack(fill="x", padx=20, pady=(4, 10))
        self.progress.start()
        self.update_idletasks()
        for b in self.buttons:
            b.configure(state="disabled")

    def hide_loading(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.progress.set(0)
        self.lbl_status.configure(text="")
        for b in self.buttons:
            b.configure(state="normal")


class ResultsStep(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color="transparent")
        self.app = app

        title_row = ctk.CTkFrame(self, fg_color="transparent")
        title_row.pack(fill="x", padx=10, pady=(8, 2))
        ctk.CTkLabel(title_row, text="Validation Results",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")
        self.btn_export = ctk.CTkButton(title_row, text="Export Results to Excel",
                                        fg_color="#1f538d", command=self.app.export_results)
        self.btn_export.pack(side="right", padx=5)
        self.btn_restart = ctk.CTkButton(title_row, text="Restart", fg_color="#333",
                                         hover_color="#555", command=self.restart)
        self.btn_restart.pack(side="right", padx=5)
        self.btn_back = ctk.CTkButton(title_row, text="\u2190 Back", fg_color="#333",
                                      hover_color="#555",
                                      command=lambda: self.app.show_step(2))
        self.btn_back.pack(side="right", padx=5)

        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.pack(fill="x", padx=10, pady=6)
        self.card_records = self._make_card(cards, "Total Records (Filtered)")
        self.card_amount = self._make_card(cards, "Total DBF Amount", ACCENT_GREEN)
        self.card_matches = self._make_card(cards, "Number of Matches", ACCENT_BLUE)
        self.card_validated = self._make_card(cards, "Total Validated Amount", ACCENT_ORANGE)

        tables = ctk.CTkFrame(self, fg_color="transparent")
        tables.pack(fill="both", expand=True, padx=5, pady=(4, 4))
        tables.columnconfigure(0, weight=1)
        tables.columnconfigure(1, weight=1)
        tables.rowconfigure(0, weight=1)

        left = ctk.CTkFrame(tables)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
        ctk.CTkLabel(left, text="DBF Extracted Data", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=ACCENT_BLUE).pack(anchor="w", padx=10, pady=(4, 0))
        self.table_dbf = ScrollableTable(left, columns=DBF_COLS, app=self.app)
        self.table_dbf.pack(fill="both", expand=True, padx=4, pady=4)

        right = ctk.CTkFrame(tables)
        right.grid(row=0, column=1, sticky="nsew", padx=(3, 0))
        ctk.CTkLabel(right, text="Excel Data", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=ACCENT_GREEN).pack(anchor="w", padx=10, pady=(4, 0))
        self.table_excel = ScrollableTable(right, columns=DBF_COLS, app=self.app)
        self.table_excel.pack(fill="both", expand=True, padx=4, pady=4)

        self.lbl_loading = ctk.CTkLabel(tables, text="Loading result data\u2026",
                                        font=ctk.CTkFont(size=14), text_color="gray")
        self.lbl_loading.grid(row=1, column=0, columnspan=2, pady=30)

        self.refresh()

    def _make_card(self, parent, label, color=None):
        card = ctk.CTkFrame(parent, corner_radius=8)
        card.pack(side="left", expand=True, fill="both", padx=4)
        ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=11)).pack(pady=(8, 2))
        lbl = ctk.CTkLabel(card, text="0", font=ctk.CTkFont(size=18, weight="bold"), text_color=color)
        lbl.pack(pady=(0, 8))
        return lbl

    def refresh(self):
        app = self.app
        self.card_records.configure(text=f"{app.total_records:,}")
        self.card_amount.configure(text=f"{app.total_amount:,.2f}")
        self.card_matches.configure(text=f"{app.matches:,}")
        self.card_validated.configure(text=f"{app.validated_amount:,.2f}")

    def load_frames(self, dbf_disp, excel_disp, excel_cols):
        if self.lbl_loading is not None:
            self.lbl_loading.destroy()
            self.lbl_loading = None
        self.table_dbf.set_data(dbf_disp, highlight_col="MATCHED_IN_EXCEL")
        self.table_excel.set_columns(excel_cols)
        self.table_excel.set_data(excel_disp, highlight_col="_IS_MATCH")

    def restart(self):
        if messagebox.askyesno("Restart", "Start over from step 1?"):
            self.app.restart()


class DBFExtractorValidatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DBF Extractor & Excel Validator By Kuya Daks")
        self.geometry("1280x800")
        self.minsize(900, 600)
        self.state("zoomed")

        try:
            from PIL import Image, ImageTk
            img = ImageTk.PhotoImage(Image.open(resource_path("icon.png")))
            self.iconphoto(True, img)
            self._icon_ref = img
        except Exception:
            pass

        self.dbf_path = ""
        self.excel_path = ""
        self.extracted_df = pd.DataFrame()
        self.excel_df = pd.DataFrame()
        self.validated_df = pd.DataFrame()
        self.acct_col = None
        self.start_date = None
        self.end_date = None
        self.total_records = 0
        self.total_amount = 0.0
        self.matches = 0
        self.validated_amount = 0.0
        self.current_step = None
        self.log_visible = False
        self._busy = False
        self._bg_queue = queue.Queue()
        self._results_seq = 0
        self._loading_start = time.time()

        self.create_widgets()
        self._poll_bg_queue()
        self.log_message("System ready. Load a .dbf file to start.")
        self.show_step(1)

    def create_widgets(self):
        self.progress_bar = StepProgressBar(self)
        self.progress_bar.pack(side="top", fill="x", padx=15, pady=(10, 4))

        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(side="top", fill="both", expand=True, padx=15, pady=(0, 4))

        self.log_toggle_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.log_toggle_frame.pack(side="bottom", fill="x", padx=15, pady=(0, 4))

        self.btn_toggle_log = ctk.CTkButton(
            self.log_toggle_frame, text="Show Log", width=90, height=26,
            font=ctk.CTkFont(size=11), fg_color="#333", hover_color="#555",
            command=self.toggle_log)
        self.btn_toggle_log.pack(side="left")

        self.log_bar = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=8)
        self.txt_log = ctk.CTkTextbox(self.log_bar, height=90)

    def show_step(self, num):
        if self.current_step is not None:
            self.current_step.destroy()
        self.progress_bar.set_status(num)
        if num == 1:
            self.current_step = WelcomeStep(self.container, self)
        elif num == 2:
            self.current_step = ExcelStep(self.container, self)
        else:
            self.current_step = ResultsStep(self.container, self)
        self.current_step.pack(fill="both", expand=True)

    def toggle_log(self):
        if self.log_visible:
            self.log_bar.pack_forget()
            self.btn_toggle_log.configure(text="Show Log")
            self.log_visible = False
        else:
            self.log_bar.pack(side="bottom", fill="x", padx=15, pady=(0, 4))
            self.txt_log.pack(fill="both", expand=True, padx=5, pady=5)
            self.btn_toggle_log.configure(text="Hide Log")
            self.log_visible = True

    def log_message(self, msg):
        self.txt_log.insert("end", f"> {msg}\n")
        self.txt_log.see("end")

    def _run_bg(self, step, message, work, on_done):
        if self._busy:
            return
        self._loading_start = time.time()
        if hasattr(step, "show_loading"):
            step.show_loading(message)
        self._busy = True

        def target():
            try:
                result = work()
            except Exception as e:
                result = ("ERROR", str(e))
            self.post(("bg_done", (step, result, on_done)))

        threading.Thread(target=target, daemon=True).start()

    def post(self, item):
        self._bg_queue.put(item)

    def _poll_bg_queue(self):
        try:
            while True:
                item = self._bg_queue.get_nowait()
                self._dispatch_bg(item)
        except queue.Empty:
            pass
        self.after(50, self._poll_bg_queue)

    def _dispatch_bg(self, item):
        kind = item[0]
        if kind == "bg_done":
            _step, result, on_done = item[1]
            self._finish_bg(_step, result, on_done)
        elif kind == "table_ready":
            table, seq, positions, err = item[1]
            table.apply_compute(positions, seq, err)
        elif kind == "results_frames":
            dbf_disp, excel_disp, excel_cols, seq = item[1]
            if (dbf_disp is not None
                    and isinstance(self.current_step, ResultsStep)
                    and seq == getattr(self, "_results_seq", -1)
                    and callable(getattr(self.current_step, "load_frames", None))):
                self.current_step.load_frames(dbf_disp, excel_disp, excel_cols)

    def start_table_compute(self, table, q, seq):
        def work():
            try:
                positions = compute_positions(table, q)
                return ("table_ready", (table, seq, positions, None))
            except Exception as e:
                return ("table_ready", (table, seq, [], str(e)))

        threading.Thread(target=lambda: self.post(work()), daemon=True).start()


    def _finish_bg(self, step, result, on_done):
        self._busy = False
        elapsed = time.time() - getattr(self, "_loading_start", time.time())
        remaining_ms = max(0, int((MIN_LOADING_TIME - elapsed) * 1000))

        def finish():
            if hasattr(step, "hide_loading"):
                step.hide_loading()
            if isinstance(result, tuple) and result and result[0] == "ERROR":
                messagebox.showerror("Error", result[1])
                self.log_message(f"Error: {result[1]}")
                return
            on_done(result)

        if remaining_ms:
            self.after(remaining_ms, finish)
        else:
            finish()

    def proceed_to_excel(self, start_date, end_date):
        if self._busy:
            return
        self.start_date = start_date
        self.end_date = end_date
        self._run_bg(self.current_step, "Extracting DBF data...",
                     self._extract_work, self._after_extract)

    def _extract_work(self):
        start = self.start_date
        end = self.end_date
        table = DBF(self.dbf_path, encoding="latin-1", ignore_missing_memofile=True)
        data = []
        for record in table:
            data.append({k: record.get(k) for k in DBF_COLS})
            if len(data) % 10000 == 0:
                time.sleep(0.002)
        df = pd.DataFrame(data)
        if not df.empty:
            df["ALVDATE"] = pd.to_datetime(df["ALVDATE"], errors="coerce")
            df = df[df["ALVDATE"].notna()]
            if start:
                df = df[df["ALVDATE"].dt.date >= start]
            if end:
                df = df[df["ALVDATE"].dt.date <= end]
            df["CLEAN_ACCT"] = df["ACCTNBR"].astype(str).apply(
                lambda x: re.sub(r"[^0-9a-zA-Z]", "", x))
        self.extracted_df = df
        self.total_records = len(df)
        self.total_amount = float(df["ALVAMT"].fillna(0).sum()) if "ALVAMT" in df else 0.0
        return {"empty": df.empty}

    def _after_extract(self, result):
        if result["empty"]:
            messagebox.showinfo("Info", "No records found in the selected date range.")
            self.log_message("Extraction returned 0 records.")
            return
        self.log_message(f"Extracted {self.total_records:,} records "
                         f"(Total ALVAMT {self.total_amount:,.2f}).")
        self.show_step(2)

    def probe_excel(self, step):
        if self._busy:
            return

        def work():
            excel_df = pd.read_excel(self.excel_path)
            acct = None
            for col in excel_df.columns:
                if "account" in str(col).lower():
                    acct = col
                    break
            return {"df": excel_df, "acct_col": acct}

        def on_done(r):
            self.excel_df = r["df"]
            self.acct_col = r["acct_col"]
            self.log_message(f"Excel file loaded: {len(r['df']):,} rows; "
                             f"account column: {r['acct_col']}")
            if callable(getattr(step, "on_excel_ready", None)):
                step.on_excel_ready(r["acct_col"], len(r["df"]))

        self._run_bg(step, "Reading Excel file...", work, on_done)

    def proceed_to_results(self):
        if self._busy:
            return
        self._run_bg(self.current_step, "Validating records...",
                     self._validate_work, self._after_validate)

    def _validate_work(self):
        if self.excel_df is None or self.excel_df.empty:
            if not self.excel_path:
                return {"error": "No Excel file selected."}
            self.excel_df = pd.read_excel(self.excel_path)
        excel_df = self.excel_df
        if self.acct_col is None:
            for col in excel_df.columns:
                if "account" in str(col).lower():
                    self.acct_col = col
                    break
        if self.acct_col is None:
            return {"error": "Could not find an 'Account No.' column in the Excel file."}

        excel_df["CLEAN_EXCEL_ACCT"] = excel_df[self.acct_col].astype(str).apply(
            lambda x: re.sub(r"[^0-9a-zA-Z]", "", x))
        excel_acct_set = set(excel_df["CLEAN_EXCEL_ACCT"])

        if not self.extracted_df.empty:
            self.extracted_df["MATCHED_IN_EXCEL"] = self.extracted_df["CLEAN_ACCT"].isin(
                excel_acct_set)
            matched = self.extracted_df[self.extracted_df["MATCHED_IN_EXCEL"]]
            self.matches = len(matched)
            self.validated_amount = float(matched["ALVAMT"].fillna(0).sum()) if "ALVAMT" in matched else 0.0
        else:
            self.matches = 0
            self.validated_amount = 0.0
            self.extracted_df["MATCHED_IN_EXCEL"] = False
        return {"error": None}

    def _after_validate(self, result):
        if result.get("error"):
            messagebox.showerror("Error", result["error"])
            self.log_message(f"Validation error: {result['error']}")
            return
        self.log_message(f"Validation complete! Matches: {self.matches:,} "
                         f"| Validated Amount: {self.validated_amount:,.2f}")
        self.show_step(3)
        self.populate_results()

    def populate_results(self):
        if not isinstance(self.current_step, ResultsStep):
            return
        self._results_seq = getattr(self, "_results_seq", 0) + 1
        seq = self._results_seq

        def work():
            try:
                dbf_disp = self.get_dbf_display_df()
                excel_disp = self.get_excel_display_df()
                excel_cols = [c for c in excel_disp.columns if c != "_IS_MATCH"]
                return ("results_frames", (dbf_disp, excel_disp, excel_cols, seq))
            except Exception:
                return ("results_frames", (None, None, [], seq))

        threading.Thread(target=lambda: self.post(work()), daemon=True).start()

    def get_dbf_display_df(self):
        if self.extracted_df.empty:
            return pd.DataFrame()
        cols = DBF_COLS + (["MATCHED_IN_EXCEL"] if "MATCHED_IN_EXCEL" in self.extracted_df.columns else [])
        return self.extracted_df[cols].copy()

    def get_excel_display_df(self):
        if self.excel_df is None or self.excel_df.empty:
            return pd.DataFrame()
        excel_display_cols = [c for c in DBF_COLS if c in self.excel_df.columns]
        if not excel_display_cols:
            excel_display_cols = list(self.excel_df.columns[:6])
        disp = self.excel_df[excel_display_cols].copy()
        if "CLEAN_EXCEL_ACCT" in self.excel_df.columns and not self.extracted_df.empty:
            dbf_set = set(self.extracted_df["CLEAN_ACCT"])
            disp["_IS_MATCH"] = self.excel_df["CLEAN_EXCEL_ACCT"].isin(dbf_set)
        else:
            disp["_IS_MATCH"] = False
        return disp

    def export_results(self):
        if self.extracted_df.empty:
            messagebox.showwarning("Warning", "No extracted data available to export.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if save_path:
            try:
                cols_to_export = list(DBF_COLS)
                if "MATCHED_IN_EXCEL" in self.extracted_df.columns:
                    cols_to_export.append("MATCHED_IN_EXCEL")

                export_df = self.extracted_df[cols_to_export].copy()

                for col in export_df.columns:
                    if export_df[col].dtype == "object":
                        export_df[col] = export_df[col].apply(clean_illegal_chars)

                if "ALVDATE" in export_df.columns:
                    export_df["ALVDATE"] = export_df["ALVDATE"].dt.strftime("%Y-%m-%d")

                export_df.to_excel(save_path, index=False, engine="openpyxl")
                messagebox.showinfo("Success", f"Data exported successfully to:\n{save_path}")
                self.log_message(f"Exported data to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save export file:\n{str(e)}")

    def restart(self):
        self.dbf_path = ""
        self.excel_path = ""
        self.extracted_df = pd.DataFrame()
        self.excel_df = pd.DataFrame()
        self.validated_df = pd.DataFrame()
        self.acct_col = None
        self.start_date = None
        self.end_date = None
        self.total_records = 0
        self.total_amount = 0.0
        self.matches = 0
        self.validated_amount = 0.0
        self.show_step(1)


if __name__ == "__main__":
    app = DBFExtractorValidatorApp()
    app.mainloop()