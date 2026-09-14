import os
import re
import sys
import time
import queue
import traceback
import datetime
import threading
from html import unescape as html_unescape
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont
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
DBF_VALIDATOR_NAME = "DBF Extractor"
SQL_VALIDATOR_NAME = "SQL Summary Receipt (tblsummaryreceipt)"
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
HEADER_HEIGHT = 32
TABLE_ROW_HEIGHT = 28
MIN_COL_WIDTH = 60
MAX_COL_WIDTH = 420
COL_WIDTH_PADDING = 26
COL_WIDTH_SAMPLE = 200
RESIZE_TOLERANCE = 6


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


def parse_money(val):
    """Convert a money value that may contain commas / currency symbols to float."""
    if val is None:
        return 0.0
    if isinstance(val, str):
        s = val.strip().replace(",", "")
        if not s:
            return 0.0
        s = re.sub(r"[^0-9.\-]", "", s)
    else:
        s = str(val)
    try:
        return float(s)
    except (TypeError, ValueError):
        return 0.0


def name_tokens(s):
    """Return lowercased alphanumeric word tokens (length >= 2) of a name.

    Drops single-letter tokens (e.g. middle initials) so a missing/extra
    initial on either side does not break the match.
    """
    if s is None:
        return frozenset()
    try:
        if pd.isna(s):
            return frozenset()
    except (TypeError, ValueError):
        pass
    text = re.sub(r"[^0-9a-zA-Z ]", " ", str(s)).lower()
    return frozenset(w for w in text.split() if len(w) >= 2)


def find_account_column(df):
    """Pick the Excel column holding the account number (account/acct, not name)."""
    for col in df.columns:
        s = str(col).lower()
        if any(k in s for k in ("account", "acct")) and "name" not in s:
            return col
    for col in df.columns:
        s = str(col).lower()
        if any(k in s for k in ("account", "acct")):
            return col
    return None


def find_name_column(df):
    """Pick the Excel column holding the customer name (prefers account/customer name)."""
    for col in df.columns:
        s = str(col).lower()
        if "name" in s and any(k in s for k in ("account", "customer", "cust")):
            return col
    for col in df.columns:
        if "name" in str(col).lower():
            return col
    return None


def probe_xml_columns(path):
    """Return lowercased column names from an XML export (header-only probe)."""
    import xml.etree.ElementTree as ET

    try:
        for _event, elem in ET.iterparse(path):
            if not list(elem):
                continue
            children = list(elem)
            leaves = [c for c in children if not list(c)]
            if leaves and len(leaves) == len(children):
                names = set()
                for c in leaves:
                    tag = str(c.tag).split('}')[-1].strip().lower()
                    if tag == "field" and c.attrib.get("name"):
                        name = str(c.attrib["name"]).strip().lower()
                        if name:
                            names.add(name)
                            continue
                    if tag:
                        names.add(tag)
                if names:
                    return sorted(names)
    except Exception:
        pass
    try:
        df = pd.read_xml(path)
        if not df.empty:
            return [str(c).strip().lower() for c in df.columns]
        return []
    except Exception:
        return []


def probe_data_columns(path):
    """Return (column names, error) for a data file without reading all rows.

    DBF columns keep their original case; CSV/XML columns are lowercased to
    match how they are later read by the SQL parser.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".dbf":
        try:
            tbl = DBF(path, encoding="latin-1", ignore_missing_memofile=True)
            return list(tbl.field_names), None
        except Exception as e:
            return None, f"Failed to read DBF columns: {e}"
    if ext == ".csv":
        for enc in ("utf-8-sig", "latin-1"):
            for kw in ({"sep": ","}, {"sep": "\t"}, {"sep": ";"}):
                try:
                    df = pd.read_csv(path, nrows=0, dtype=str, keep_default_na=False,
                                     encoding=enc, on_bad_lines="skip", **kw)
                    if isinstance(df.index, pd.MultiIndex):
                        raise ValueError("misaligned columns")
                    return [str(c).strip().lower() for c in df.columns], None
                except Exception:
                    continue
        return None, "Could not read CSV columns."
    if ext == ".xml":
        cols = probe_xml_columns(path)
        if cols:
            return cols, None
        return None, "Could not read XML columns."
    return None, "Unsupported file type. Use .dbf, .csv or .xml."


def detect_column_roles(cols):
    """Suggest default role selections for the given source columns."""
    roles = {"match": "", "sums": [], "date": "", "type": ""}
    if not cols:
        return roles
    for c in cols:
        s = str(c).strip().lower()
        if not roles["match"] and any(k in s for k in ("account", "acct")) \
                and "name" not in s:
            roles["match"] = c
        if not roles["date"] and "date" in s:
            roles["date"] = c
        if not roles["type"] and (s == "type" or s.endswith("type")
                                  or s.endswith("_type")):
            roles["type"] = c
        if s in ("alvamt", "cashout", "loanamount", "loanamt", "amount", "amt",
                 "totalamount", "totalamt", "principal", "interest", "amountdue"):
            roles["sums"].append(c)
        elif any(k in s for k in ("amount", "amt", "cash", "cashout", "loan",
                                  "alvamt", "total")):
            roles["sums"].append(c)
    seen = set()
    uniq = []
    for c in roles["sums"]:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    roles["sums"] = uniq
    return roles


def _read_csv_any(path):
    attempts = [{"sep": ","}, {"sep": "\t"}, {"sep": ";"}]
    for enc in ("utf-8-sig", "latin-1"):
        for kw in attempts:
            try:
                df = pd.read_csv(path, dtype=str, keep_default_na=False,
                                 encoding=enc, on_bad_lines="skip", **kw)
                if isinstance(df.index, pd.MultiIndex):
                    raise ValueError("misaligned columns (unquoted commas in field?)")
                return df
            except Exception:
                continue
    return None


def _read_xml_any(path):
    import xml.etree.ElementTree as ET

    try:
        tree = ET.parse(path)
        root = tree.getroot()
        records = []
        for elem in root.iter():
            if elem is root or not list(elem):
                continue
            fields = {}
            for child in elem:
                if child.tag.lower() == "field" and child.attrib.get("name"):
                    key = child.attrib["name"].lower()
                else:
                    key = child.tag.lower()
                val = (child.text or "").strip()
                if val:
                    try:
                        val = html_unescape(val)
                    except Exception:
                        pass
                fields[key] = val
            if fields:
                records.append(fields)
            if len(records) % 5000 == 0:
                time.sleep(0.002)
        if records:
            cols = sorted(set().union(*(set(r) for r in records)))
            return pd.DataFrame([{c: r.get(c, "") for c in cols} for r in records])
    except Exception:
        pass
    try:
        return pd.read_xml(path)
    except Exception:
        return None


def read_sql_table_file(path):
    """Read a tblsummaryreceipt export (.csv or .xml) into a DataFrame."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".csv":
            df = _read_csv_any(path)
        elif ext == ".xml":
            df = _read_xml_any(path)
        else:
            return None, "Unsupported file type. Use .csv or .xml."
        if df is None:
            return None, "Could not read the file as a table."
        df.columns = [str(c).strip().lower() for c in df.columns]
        return df, None
    except Exception as e:
        return None, f"Failed to parse file: {e}"


class ScrollableTable(ctk.CTkFrame):
    """Scrollable, paged table with header, debounced async search, and row highlighting.

    Only a bounded window of rows (PAGE_SIZE) is rendered at a time. All heavy work
    (DataFrame prep, search-key building, filtering) runs off the UI thread via the
    app's background queue so the window stays responsive even with very large data.

    Column widths are draggable: grab the divider in the header to resize that column.
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

        self._col_widths = {}
        self._widths_dirty = False
        self._header_texts = []
        self._header_dividers = []
        self._row_cell_sets = []
        self._drag_col = None
        self._drag_x0 = 0
        self._drag_w0 = 0

        search_frame = ctk.CTkFrame(self, fg_color="transparent")
        search_frame.pack(fill="x", padx=5, pady=(5, 2))
        ctk.CTkLabel(search_frame, text="Search:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_changed)
        self.ent_search = ctk.CTkEntry(search_frame, textvariable=self.search_var,
                                       placeholder_text="Filter rows...")
        self.ent_search.pack(side="left", fill="x", expand=True)

        self.head_canvas = tk.Canvas(self, height=HEADER_HEIGHT, bg="#3a3a3a", highlightthickness=0)
        self.head_canvas.pack(fill="x", padx=5, pady=(0, 2))
        self._build_header()
        self.head_canvas.bind("<ButtonPress-1>", self._on_head_press)
        self.head_canvas.bind("<B1-Motion>", self._on_head_drag)
        self.head_canvas.bind("<ButtonRelease-1>", self._on_head_release)
        self.head_canvas.bind("<Motion>", self._on_head_motion)

        body = tk.Frame(self, bg="#2b2b2b")
        body.pack(fill="both", expand=True, padx=5, pady=(0, 0))
        self.scroll_y = tk.Scrollbar(body, orient="vertical")
        self.scroll_y.pack(side="right", fill="y")
        self.canvas = tk.Canvas(body, bg="#2b2b2b", highlightthickness=0,
                                yscrollcommand=self.scroll_y.set)
        self.scroll_y.configure(command=self.canvas.yview)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body_inner = tk.Frame(self.canvas, bg="#2b2b2b")
        self._body_window = self.canvas.create_window((0, 0), window=self.body_inner, anchor="nw")

        self.scroll_x = tk.Scrollbar(self, orient="horizontal", command=self._xview)
        self.canvas.configure(xscrollcommand=self.scroll_x.set)
        self.head_canvas.configure(xscrollcommand=self.scroll_x.set)
        self.scroll_x.pack(fill="x", padx=5, pady=(3, 2))

        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.head_canvas.bind("<Configure>", self._on_canvas_configure)
        self.body_inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.body_inner.bind("<MouseWheel>", self._on_wheel)
        self.head_canvas.bind("<MouseWheel>", self._on_wheel)

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

    # ------------------------------------------------------------------ header
    def _build_header(self):
        for item in self._header_texts + self._header_dividers:
            self.head_canvas.delete(item)
        self._header_texts.clear()
        self._header_dividers.clear()
        for col in self.columns:
            item = self.head_canvas.create_text(0, HEADER_HEIGHT // 2, anchor="w",
                                                text=str(col), fill="#e0e0e0",
                                                font=("Segoe UI", 11, "bold"))
            self._header_texts.append(item)
        self._place_header()

    def _place_header(self):
        x = 0
        for i, item in enumerate(self._header_texts):
            w = self._col_widths.get(self.columns[i], 100)
            self.head_canvas.coords(item, x + 6, HEADER_HEIGHT // 2)
            x += w
        while len(self._header_dividers) < max(0, len(self.columns) - 1):
            div = self.head_canvas.create_line(0, 3, 0, HEADER_HEIGHT - 3, fill="#5a5a5a")
            self._header_dividers.append(div)
        x = 0
        for i, div in enumerate(self._header_dividers):
            x += self._col_widths.get(self.columns[i], 100)
            self.head_canvas.coords(div, x, 3, x, HEADER_HEIGHT - 3)

    def _place_row_cells(self, container, cells):
        x = 0
        for i, lbl in enumerate(cells):
            w = self._col_widths.get(self.columns[i], 100)
            lbl.place(x=x, y=0, width=w, height=TABLE_ROW_HEIGHT)
            x += w

    def _boundaries(self):
        edges = []
        x = 0
        for c in self.columns:
            x += self._col_widths.get(c, 100)
            edges.append(x)
        return edges

    # ------------------------------------------------------------ drag resizing
    def _on_head_press(self, event):
        if self._drag_col is not None or not self.columns:
            return
        x = self.head_canvas.canvasx(event.x)
        for i, bx in enumerate(self._boundaries()):
            if abs(x - bx) <= RESIZE_TOLERANCE:
                self._drag_col = i
                self._drag_x0 = event.x
                self._drag_w0 = self._col_widths.get(self.columns[i], 100)
                return

    def _on_head_drag(self, event):
        if self._drag_col is None:
            return
        delta = event.x - self._drag_x0
        new_w = max(MIN_COL_WIDTH, min(int(self._drag_w0 + delta), MAX_COL_WIDTH))
        self._col_widths[self.columns[self._drag_col]] = new_w
        self._place_header()
        for row_cells in self._row_cell_sets:
            if row_cells:
                self._place_row_cells(row_cells[0].master, row_cells)
        self.head_canvas.configure(cursor="sb_h_double_arrow")
        self._sync_widths()

    def _on_head_release(self, _event):
        self._drag_col = None
        self.head_canvas.configure(cursor="")

    def _on_head_motion(self, event):
        if self._drag_col is not None:
            self.head_canvas.configure(cursor="sb_h_double_arrow")
            return
        x = self.head_canvas.canvasx(event.x)
        near = any(abs(x - bx) <= RESIZE_TOLERANCE for bx in self._boundaries())
        self.head_canvas.configure(cursor="sb_h_double_arrow" if near else "")

    # ------------------------------------------------------------ width sizing
    def _measure_width(self, col):
        head_font = tkfont.Font(family="Segoe UI", size=11, weight="bold")
        body_font = tkfont.Font(family="Segoe UI", size=11)
        w = head_font.measure(str(col))
        if col in self.source_df.columns:
            ser = self.source_df[col]
            for v in ser.iloc[:COL_WIDTH_SAMPLE]:
                if pd.notna(v):
                    if isinstance(v, pd.Timestamp):
                        text = v.strftime("%Y-%m-%d")
                    else:
                        text = str(v)
                    m = body_font.measure(text)
                    if m > w:
                        w = m
        return int(min(max(w + COL_WIDTH_PADDING, MIN_COL_WIDTH), MAX_COL_WIDTH))

    def _ensure_widths(self):
        if not self._widths_dirty:
            return
        self._widths_dirty = False
        for c in self.columns:
            if c not in self._col_widths:
                self._col_widths[c] = self._measure_width(c)
        self._place_header()

    def _content_width(self):
        if not self.columns:
            return 0
        return sum(self._col_widths.get(c, 100) for c in self.columns)

    # ------------------------------------------------------------ scroll/view
    def _xview(self, *args):
        self.canvas.xview(*args)
        self.head_canvas.xview(*args)

    def _sync_widths(self):
        view_ok = self.canvas.winfo_width() > 1 and self.head_canvas.winfo_width() > 1
        if not view_ok:
            return
        total_w = max(self._content_width(), self.canvas.winfo_width(),
                      self.head_canvas.winfo_width())
        self.canvas.itemconfigure(self._body_window, width=total_w)
        b_h = self.body_inner.winfo_reqheight()
        self.canvas.configure(scrollregion=(0, 0, total_w, max(b_h, self.canvas.winfo_height())))
        self.head_canvas.configure(scrollregion=(0, 0, total_w, HEADER_HEIGHT))

    def _on_canvas_configure(self, _event):
        self._sync_widths()

    def _on_inner_configure(self, _event):
        self._sync_widths()

    def _on_wheel(self, event):
        try:
            self.canvas.yview_scroll(int(-event.delta / 120), "units")
        except Exception:
            pass

    def set_columns(self, columns):
        self.columns = list(columns)
        for c in list(self._col_widths):
            if c not in self.columns:
                del self._col_widths[c]
        self._widths_dirty = True
        self._build_header()
        self.clear()

    def set_data(self, df, highlight_col=None):
        self._computing = True
        self._filtered = []
        self.page = 0
        self.highlight_col = highlight_col
        self.source_df = df if (df is not None and not df.empty) else pd.DataFrame()
        self._keys = None
        self._widths_dirty = True
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
            if self.app is not None and hasattr(self.app, "log_message"):
                self.app.log_message(f"Table search error: {err}")
            return
        self._filtered = positions
        try:
            self._render_page()
        except Exception:
            if self.app is not None and hasattr(self.app, "log_message"):
                self.app.log_message(f"Table render error:\n{traceback.format_exc()}")
            self.result_count_label.configure(text="Render error \u2014 check the Log.")

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
        self._row_cell_sets.clear()
        self._ensure_widths()

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
        row_frame = tk.Frame(self.body_inner, bg=bg)
        row_frame.pack(fill="x", padx=2, pady=1)
        cells = []
        for col in self.columns:
            val = row_data.get(col, "")
            lbl = tk.Label(row_frame, text=val, bd=0, bg=bg, fg=fg,
                           font=("Segoe UI", 11), anchor="w")
            cells.append(lbl)
        self._place_row_cells(row_frame, cells)
        self.row_widgets.append(row_frame)
        self._row_cell_sets.append(cells)

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
        self._row_cell_sets.clear()
        self.result_count_label.configure(text="0 of 0 rows")
        self.btn_prev.configure(state="disabled")
        self.btn_next.configure(state="disabled")
        self.search_var.set("")

    def to_tsv(self):
        """Tab-separated copy of the displayed columns (header + all rows).

        Built with pandas vectorized to_csv so large tables copy fast without
        blocking the UI thread.
        """
        if self.source_df is None or self.source_df.empty or not self.columns:
            return ""
        want = [c for c in self.columns if c in self.source_df.columns]
        if not want:
            return ""
        txt = self.source_df[want].to_csv(sep="\t", index=False,
                                          lineterminator="\n", na_rep="")
        return txt.rstrip("\n")


class RoleCheckMatrix(ctk.CTkScrollableFrame):
    """Row-per-column matrix of role checkboxes.

    Each source column is a row; each column of the UI is a role checkbox
    (e.g. Match / Sum / Date / Type). Roles listed in ``single_select`` allow
    at most one column chosen at a time (checking one unchecks the others);
    the rest are multi-select.
    """

    def __init__(self, master, roles, single_select=(), **kwargs):
        super().__init__(master, **kwargs)
        self.roles = list(roles)
        self.single = set(single_select)
        self.on_change = None
        self._rows = []  # list of (col, {role: tk.BooleanVar})
        self.lbl_header = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=11, weight="bold"), anchor="w")
        self.lbl_empty = ctk.CTkLabel(self, text="", text_color="gray", font=ctk.CTkFont(size=11),
                                      anchor="w")

    def set_columns(self, columns, preselect=None):
        """Rebuild the matrix for the given columns.

        preselect maps a role to a column name (single) or list of names (multi).
        """
        preselect = preselect or {}
        for row in self._rows:
            for v in row[1].values():
                v.set(False)
        for w, _v in self._rows:
            try:
                w.destroy()
            except Exception:
                pass
        self._rows = []
        self.lbl_header.pack_forget()
        self.lbl_empty.pack_forget()
        columns = list(columns)
        if not columns:
            self.lbl_empty.configure(text="No columns available.")
            self.lbl_empty.pack(anchor="w", fill="x", padx=8, pady=4)
            return
        header = "Column"
        for role in self.roles:
            header += "   " + role.title()
        self.lbl_header.configure(text=header)
        self.lbl_header.pack(anchor="w", fill="x", padx=8, pady=(2, 4))
        for c in columns:
            row = ctk.CTkFrame(self, fg_color="transparent")
            row.pack(fill="x", padx=4, pady=1)
            row.grid_columnconfigure(0, weight=1)
            vars_ = {}
            for i, role in enumerate(self.roles):
                default = self._default_value(preselect, role, c)
                v = tk.BooleanVar(value=bool(default))
                vars_[role] = v
                cb = ctk.CTkCheckBox(row, text=role.title(), width=96,
                                     variable=v, font=ctk.CTkFont(size=11),
                                     command=lambda cc=c, rr=role: self._on_change(cc, rr))
                cb.grid(row=0, column=i + 1, padx=(2, 2), pady=1)
            lbl = ctk.CTkLabel(row, text=str(c), anchor="w",
                               font=ctk.CTkFont(size=11))
            lbl.grid(row=0, column=0, sticky="w", padx=(6, 6))
            self._rows.append((c, vars_))
        self._on_change()

    def _default_value(self, preselect, role, col):
        val = preselect.get(role)
        if val is None:
            return False
        if isinstance(val, (list, tuple, set, frozenset)):
            return col in val
        return str(val) == str(col)

    def _on_change(self, col=None, role=None, *_):
        if col is not None and role is not None and role in self.single:
            was = None
            for c, v in self._rows:
                if c == col:
                    was = v[role].get()
                    break
            if was:  # this checkbox was just turned ON -> clear the others
                for oc, ov in self._rows:
                    if oc != col and ov[role].get():
                        ov[role].set(False)
        if callable(self.on_change):
            self.on_change(self.get_selected())

    def get_selected(self):
        selected = {}
        for role in self.roles:
            if role in self.single:
                selected[role] = next((c for c, v in self._rows if v[role].get()), None)
            else:
                selected[role] = [c for c, v in self._rows if v[role].get()]
        return selected


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


class WelcomeStep(ctk.CTkScrollableFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color="transparent")
        self.app = app

        hero = ctk.CTkFrame(self, fg_color="transparent")
        hero.pack(fill="x", pady=(15, 5))
        ctk.CTkLabel(hero, text="DBF / SQL Data Extractor & Excel Validator By Kuya Daks",
                     font=ctk.CTkFont(size=28, weight="bold")).pack(pady=(10, 2))
        ctk.CTkLabel(hero, text="Validate DBF records or SQL table exports (CSV/XML) against your Excel file in three simple steps.",
                     font=ctk.CTkFont(size=14), text_color="gray").pack()

        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.pack(fill="x", padx=30, pady=10)
        howto = [
            ("1", "Select Data File + Columns", "Pick a .dbf file or a .csv / .xml export, then tick which columns to match, sum and group by."),
            ("2", "Upload Excel File", "Pick the .xlsx / .xls target file and choose the column that holds the account numbers."),
            ("3", "Review Results", "See matched records side-by-side with totals and highlighting."),
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

        ctk.CTkLabel(cfg, text="Load Data File, Pick Columns & Set Date Range",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=20, pady=(12, 4))

        file_row = ctk.CTkFrame(cfg, fg_color="transparent")
        file_row.pack(fill="x", padx=20, pady=6)
        self.btn_browse = ctk.CTkButton(file_row, text="Select Data File",
                                        command=self.browse_dbf, fg_color="#1f538d")
        self.btn_browse.pack(side="left", padx=(0, 10))
        self.lbl_dbf = ctk.CTkLabel(file_row, text="No data file selected", text_color="gray")
        self.lbl_dbf.pack(side="left")

        self.lbl_mode = ctk.CTkLabel(cfg, text="", font=ctk.CTkFont(size=11, weight="bold"),
                                     text_color=ACCENT_BLUE, anchor="w", wraplength=1000,
                                     justify="left")
        self.lbl_mode.pack(anchor="w", padx=20, pady=(0, 2))

        ctk.CTkLabel(cfg, text="Date Range Filter",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=20, pady=(6, 0))
        date_row = ctk.CTkFrame(cfg, fg_color="transparent")
        date_row.pack(fill="x", padx=20, pady=6)
        ctk.CTkLabel(date_row, text="From Date:").pack(side="left", padx=(0, 5))
        today = datetime.date.today()
        one_year_ago = today.replace(year=today.year - 1)
        self.cal_start = DateEntry(date_row, width=12, background="#1f538d", foreground="white",
                                   borderwidth=2, date_pattern="yyyy-mm-dd",
                                   year=one_year_ago.year, month=one_year_ago.month,
                                   day=one_year_ago.day, font=ctk.CTkFont(size=11))
        self.cal_start.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(date_row, text="To Date:").pack(side="left", padx=(0, 5))
        self.cal_end = DateEntry(date_row, width=12, background="#1f538d", foreground="white",
                                 borderwidth=2, date_pattern="yyyy-mm-dd",
                                 year=today.year, month=today.month, day=today.day,
                                 font=ctk.CTkFont(size=11))
        self.cal_end.pack(side="left")

        roles_box = ctk.CTkFrame(cfg, fg_color="transparent")
        roles_box.pack(fill="x", padx=20, pady=(2, 0))
        ctk.CTkLabel(roles_box, text="Column Roles",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(roles_box, text="Match = key compared against the Excel account column \u2022 Sum = totals you want \u2022 Date = the column filtered by the date range \u2022 Type = optional group for the totals table (e.g. transactiontype)",
                     font=ctk.CTkFont(size=11), text_color="gray", wraplength=1000,
                     justify="left", anchor="w").pack(anchor="w")
        self.roles = RoleCheckMatrix(roles_box, roles=("match", "sum", "date", "type"),
                                     single_select=("match", "date", "type"),
                                     height=120, corner_radius=6)
        self.roles.pack(fill="x", pady=(4, 0))
        self.roles.on_change = self._on_roles_changed
        self.roles.lbl_empty.configure(text="Select a data file to see its columns.")

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
        path = filedialog.askopenfilename(
            filetypes=[("Data Files", "*.dbf *.csv *.xml"),
                       ("DBF Files", "*.dbf"),
                       ("CSV Files", "*.csv"),
                       ("XML Files", "*.xml"),
                       ("All Files", "*.*")])
        if path:
            self.app.dbf_path = path
            self.lbl_dbf.configure(text=os.path.basename(path), text_color="white")
            self.app.log_message(f"Selected Data File: {path}")
            self._detect_mode(path)

    def _detect_mode(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".dbf":
            self.app.validator_type = "dbf"
        else:
            self.app.validator_type = "sqlsummary"
        self.lbl_mode.configure(
            text="Detecting columns...",
            text_color=ACCENT_GREEN if self.app.validator_type == "sqlsummary"
            else ACCENT_BLUE)
        self.app.log_message(f"Detected validator type: {self.app.validator_type}")
        self.app.probe_source_columns(self)

    def on_columns_probed(self, cols, err):
        if err:
            self.lbl_mode.configure(text=f"Could not read columns: {err}", text_color=ACCENT_ORANGE)
            return
        self.app.src_columns = list(cols)
        roles = detect_column_roles(cols)
        self.app.role_match_col = roles["match"] or None
        self.app.role_sum_cols = list(roles["sums"])
        self.app.role_date_col = roles["date"] or None
        self.app.role_type_col = roles["type"] or None
        self.roles.set_columns(cols, preselect={
            "match": roles["match"],
            "sum": roles["sums"],
            "date": roles["date"],
            "type": roles["type"],
        })
        self.app.log_message(
            f"Auto-detected roles | Match: {roles['match'] or '-'} | "
            f"Sum: {', '.join(roles['sums']) or '-'} | Date: {roles['date'] or '-'} | "
            f"Type: {roles['type'] or '-'}")
        self._update_mode_badge()

    def _on_roles_changed(self, selected):
        self.app.role_match_col = selected.get("match")
        self.app.role_sum_cols = list(selected.get("sum") or [])
        self.app.role_date_col = selected.get("date")
        self.app.role_type_col = selected.get("type")
        self._update_mode_badge()
        self.app.log_message(
            f"Roles set | Match: {self.app.role_match_col or '-'} | "
            f"Sum: {', '.join(self.app.role_sum_cols) or '-'} | "
            f"Date: {self.app.role_date_col or '-'} | "
            f"Type: {self.app.role_type_col or '-'}")

    def _update_mode_badge(self):
        if not self.app.validator_type:
            return
        name = (SQL_VALIDATOR_NAME if self.app.validator_type == "sqlsummary"
                else DBF_VALIDATOR_NAME)
        parts = [f"Validator: {name}"]
        if self.app.role_date_col:
            parts.append(f"Date filter -> {self.app.role_date_col}")
        else:
            parts.append("No Date column selected \u2014 records NOT date-filtered")
        if self.app.role_match_col:
            parts.append(f"Match key -> {self.app.role_match_col}")
        if self.app.role_sum_cols:
            parts.append(f"Sums -> {', '.join(self.app.role_sum_cols)}")
        if self.app.role_type_col:
            parts.append(f"Group by -> {self.app.role_type_col}")
        self.lbl_mode.configure(
            text="   \u2022   ".join(parts),
            text_color=ACCENT_GREEN if self.app.validator_type == "sqlsummary"
            else ACCENT_BLUE)

    def on_continue(self):
        if not self.app.dbf_path:
            messagebox.showwarning("Warning", "Please select a data file first.")
            return
        if not self.app.validator_type:
            messagebox.showwarning("Warning", "Please select a data file first.")
            return
        if not self.app.role_match_col:
            messagebox.showwarning(
                "Warning", "Tick a Match column first \u2014 that is the key "
                            "compared against the Excel account numbers.")
            return
        if not self.app.role_date_col:
            ok = messagebox.askyesno(
                "No Date Column Selected",
                "No Date column is ticked, so the From/To date range will not "
                "filter any records.\n\nTick the Date column that should be "
                "filtered (e.g. reportdate) to avoid that.\n\n"
                "Continue without a Date column?")
            if not ok:
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

        roles_box = ctk.CTkFrame(cfg, fg_color="transparent")
        roles_box.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkLabel(roles_box, text="Excel Columns",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(roles_box, text="Match = the account-number column checked against your data file \u2022 Name = optional customer-name column used as a fallback matcher",
                     font=ctk.CTkFont(size=11), text_color="gray", wraplength=1000,
                     justify="left", anchor="w").pack(anchor="w")
        self.excel_roles = RoleCheckMatrix(roles_box, roles=("match", "name"),
                                           single_select=("match", "name"),
                                           height=130, corner_radius=6)
        self.excel_roles.pack(fill="x", pady=(4, 0))
        self.excel_roles.on_change = self._on_excel_roles_changed
        self.excel_roles.lbl_empty.configure(text="Select an Excel file to see its columns.")

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
        mode = "SQL (" + SQL_VALIDATOR_NAME + ")" if app.validator_type == "sqlsummary" else "DBF"
        roles = [f"Match: {app.role_match_col or '-'}"]
        if app.role_sum_cols:
            roles.append(f"Sum: {', '.join(app.role_sum_cols)}")
        if app.role_date_col:
            roles.append(f"Date: {app.role_date_col}")
        if app.role_type_col:
            roles.append(f"Group by: {app.role_type_col}")
        self.lbl_recap.configure(text=f"Source File : {dbf}\nValidator : {mode}\n"
                                      f"{'  |  '.join(roles)}\n"
                                      f"Date Range : {start}  to  {end}\n"
                                      f"Records Extracted : {records}")

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")])
        if path:
            self.app.excel_path = path
            self.lbl_excel.configure(text=os.path.basename(path), text_color="white")
            self.lbl_account.configure(text="")
            self.app.probe_excel(self)

    def on_excel_ready(self, acct_col, name_col, num_rows, columns=None):
        if columns is not None:
            self.app.excel_columns = list(columns)
            self.excel_roles.set_columns(columns, preselect={
                "match": acct_col or "",
                "name": name_col or "",
            })
        if acct_col:
            self.lbl_account.configure(text=f"\u2713 Suggested Match column: {acct_col}"
                                            f"{' | Name: ' + name_col if name_col else ''}"
                                            f" ({num_rows:,} rows)")
        else:
            self.lbl_account.configure(text="Pick a Match column from the list below.",
                                       text_color=ACCENT_ORANGE)

    def _on_excel_roles_changed(self, selected):
        self.app.acct_col = selected.get("match")
        self.app.name_col = selected.get("name")
        self.app.log_message(f"Excel roles | Match: {self.app.acct_col or '-'} | "
                             f"Name: {self.app.name_col or '-'}")
        if self.app.acct_col:
            rows = len(self.app.excel_df) if self.app.excel_df is not None else 0
            self.lbl_account.configure(
                text=f"\u2713 Match column: {self.app.acct_col}"
                     f"{' | Name: ' + self.app.name_col if self.app.name_col else ''}"
                     f" ({rows:,} rows)", text_color=ACCENT_GREEN)

    def on_back(self):
        self.app.show_step(1)

    def on_run(self):
        if not self.app.excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        if not self.app.acct_col:
            messagebox.showwarning("Warning", "Tick a Match column on the Excel file "
                                              "before running the validation.")
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
        self.is_sql = self.app.validator_type == "sqlsummary"
        self.sum_cols = list(self.app.role_sum_cols or [])
        self.has_type_col = bool(self.app.role_type_col
                                and self.app.role_type_col in self.app.extracted_df.columns)
        self.card_records = self._make_card(cards, "Total Records (Filtered)")
        self.card_matches = self._make_card(cards, "Number of Matches", ACCENT_BLUE)
        self.card_amounts = {}
        for sc in self.sum_cols:
            if self.app._has_sum_col(sc):
                self.card_amounts[sc] = {
                    "total": self._make_card(cards, f"Total {sc}", ACCENT_GREEN),
                    "validated": self._make_card(cards, f"Matched {sc}", ACCENT_ORANGE),
                }

        self.totals_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.totals_frame.pack(fill="x", padx=10, pady=(0, 4))
        totals_row = ctk.CTkFrame(self.totals_frame, fg_color="transparent")
        totals_row.pack(anchor="w", fill="x", padx=2, pady=(0, 2))
        self.lbl_totals = ctk.CTkLabel(totals_row, text="Totals by Type",
                                       font=ctk.CTkFont(size=12, weight="bold"),
                                       text_color=ACCENT_ORANGE, anchor="w")
        self.lbl_totals.pack(side="left")
        self.btn_copy_totals = ctk.CTkButton(totals_row, text="Copy", width=64, height=24,
                                             font=ctk.CTkFont(size=10),
                                             fg_color="#333", hover_color="#555",
                                             command=self._copy_totals)
        self.btn_copy_totals.pack(side="right")
        self.table_totals = ScrollableTable(self.totals_frame, height=165, app=self.app,
                                            columns=["Group", "Count"])
        self.table_totals.pack(fill="x", padx=4, pady=(0, 2))
        if not self.has_type_col:
            self.totals_frame.pack_forget()

        self.tables = ctk.CTkFrame(self, fg_color="transparent")
        self.tables.pack(fill="both", expand=True, padx=5, pady=(4, 4))
        self.tables.columnconfigure(0, weight=1, uniform="resultpane")
        self.tables.columnconfigure(1, weight=1, uniform="resultpane")
        self.tables.rowconfigure(0, weight=1)

        self.pane_left = ctk.CTkFrame(self.tables)
        self.pane_left.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
        left_row = ctk.CTkFrame(self.pane_left, fg_color="transparent")
        left_row.pack(anchor="w", fill="x", padx=10, pady=(4, 0))
        self.lbl_left_header = ctk.CTkLabel(
            left_row,
            text=SQL_VALIDATOR_NAME if self.is_sql else "DBF Extracted Data",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=ACCENT_BLUE)
        self.lbl_left_header.pack(side="left")
        self._show_all = False
        self.btn_toggle_cols = ctk.CTkButton(left_row, text="All Columns",
                                             width=100, height=24,
                                             font=ctk.CTkFont(size=10),
                                             fg_color="#333", hover_color="#555",
                                             command=self.toggle_cols)
        self.btn_toggle_cols.pack(side="right", padx=3)
        self.btn_left_hide = ctk.CTkButton(left_row, text="Hide", width=56, height=24,
                                           font=ctk.CTkFont(size=10),
                                           fg_color="#5a2d2d", hover_color="#7a3d3d",
                                           command=lambda: self._set_pane_mode("right"))
        self.btn_left_hide.pack(side="right", padx=3)
        self.btn_left_mode = ctk.CTkButton(left_row, text="Expand", width=70, height=24,
                                           font=ctk.CTkFont(size=10),
                                           fg_color="#1f538d", hover_color="#2a6ab3",
                                           command=lambda: self._set_pane_mode("left"))
        self.btn_left_mode.pack(side="right", padx=3)
        self.btn_copy_left = ctk.CTkButton(left_row, text="Copy", width=64, height=24,
                                           font=ctk.CTkFont(size=10),
                                           fg_color="#333", hover_color="#555",
                                           command=lambda: self._copy_table(self.table_dbf,
                                                                            "source results"))
        self.btn_copy_left.pack(side="right", padx=3)
        self.table_dbf = ScrollableTable(self.pane_left, columns=DBF_COLS, app=self.app)
        self.table_dbf.pack(fill="both", expand=True, padx=4, pady=4)

        self.pane_right = ctk.CTkFrame(self.tables)
        self.pane_right.grid(row=0, column=1, sticky="nsew", padx=(3, 0))
        right_row = ctk.CTkFrame(self.pane_right, fg_color="transparent")
        right_row.pack(anchor="w", fill="x", padx=10, pady=(4, 0))
        lbl_excel = ctk.CTkLabel(right_row, text="Excel Data", font=ctk.CTkFont(size=12, weight="bold"),
                                 text_color=ACCENT_GREEN)
        lbl_excel.pack(side="left")
        self.btn_right_hide = ctk.CTkButton(right_row, text="Hide", width=56, height=24,
                                            font=ctk.CTkFont(size=10),
                                            fg_color="#5a2d2d", hover_color="#7a3d3d",
                                            command=lambda: self._set_pane_mode("left"))
        self.btn_right_hide.pack(side="right", padx=3)
        self.btn_right_mode = ctk.CTkButton(right_row, text="Expand", width=70, height=24,
                                            font=ctk.CTkFont(size=10),
                                            fg_color="#1f538d", hover_color="#2a6ab3",
                                            command=lambda: self._set_pane_mode("right"))
        self.btn_right_mode.pack(side="right", padx=3)
        self.btn_copy_right = ctk.CTkButton(right_row, text="Copy", width=64, height=24,
                                            font=ctk.CTkFont(size=10),
                                            fg_color="#333", hover_color="#555",
                                            command=lambda: self._copy_table(self.table_excel,
                                                                             "Excel results"))
        self.btn_copy_right.pack(side="right", padx=3)
        self.table_excel = ScrollableTable(self.pane_right, columns=DBF_COLS, app=self.app)
        self.table_excel.pack(fill="both", expand=True, padx=4, pady=4)

        self.lbl_loading = ctk.CTkLabel(self.tables, text="Loading result data\u2026",
                                        font=ctk.CTkFont(size=14), text_color="gray")
        self.lbl_loading.grid(row=1, column=0, columnspan=2, pady=30)

        self.refresh()

    def _set_pane_mode(self, mode):
        """Control result-pane layout: "split" (both), "left", or "right"."""
        self.app.log_message(f"Results view: {mode} pane mode.")
        if mode == "split":
            self.tables.columnconfigure(0, weight=1, uniform="resultpane")
            self.tables.columnconfigure(1, weight=1, uniform="resultpane")
            self.pane_left.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
            self.pane_right.grid(row=0, column=1, sticky="nsew", padx=(3, 0))
            self.btn_left_mode.configure(text="Expand", width=70,
                                         command=lambda: self._set_pane_mode("left"))
            self.btn_right_mode.configure(text="Expand", width=70,
                                          command=lambda: self._set_pane_mode("right"))
            self.btn_left_hide.pack(side="right", padx=3)
            self.btn_right_hide.pack(side="right", padx=3)
        elif mode == "left":
            self.tables.columnconfigure(0, weight=1)
            self.tables.columnconfigure(1, weight=0)
            self.pane_left.grid(row=0, column=0, sticky="nsew", padx=0)
            self.pane_right.grid_remove()
            self.btn_left_hide.pack_forget()
            self.btn_left_mode.configure(text="Split View", width=90,
                                         command=lambda: self._set_pane_mode("split"))
        else:  # right
            self.tables.columnconfigure(0, weight=0)
            self.tables.columnconfigure(1, weight=1)
            self.pane_left.grid_remove()
            self.pane_right.grid(row=0, column=1, sticky="nsew", padx=0)
            self.btn_right_hide.pack_forget()
            self.btn_right_mode.configure(text="Split View", width=90,
                                          command=lambda: self._set_pane_mode("split"))

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
        self.card_matches.configure(text=f"{app.matches:,}")
        for sc, cards_ in self.card_amounts.items():
            cards_["total"].configure(text=f"{app.total_amounts.get(sc, 0.0):,.2f}")
            cards_["validated"].configure(text=f"{app.validated_amounts.get(sc, 0.0):,.2f}")

    def load_frames(self, dbf_disp, excel_disp, excel_cols, totals_df=None):
        if self.lbl_loading is not None:
            self.lbl_loading.destroy()
            self.lbl_loading = None
        dbf_cols = list(dbf_disp.columns) if not dbf_disp.empty else []
        left_hc = next((c for c in ("Matched in Excel", "MATCHED_IN_EXCEL")
                        if c in dbf_cols), None)
        self.table_dbf.set_columns(dbf_cols)
        self.table_dbf.set_data(dbf_disp, highlight_col=left_hc)
        self.table_excel.set_columns(excel_cols)
        self.table_excel.set_data(excel_disp, highlight_col="_IS_MATCH")
        if self.has_type_col and totals_df is not None and not totals_df.empty:
            self.table_totals.set_columns(list(totals_df.columns))
            self.table_totals.set_data(totals_df)

    def toggle_cols(self):
        self._show_all = not self._show_all
        self.btn_toggle_cols.configure(text="Key Columns" if self._show_all else "All Columns")
        self.app.populate_results(show_all=self._show_all)

    def _copy_table(self, table, name):
        rows = 0 if table.source_df is None else len(table.source_df)
        if table.source_df is None or table.source_df.empty or not table.columns:
            messagebox.showinfo("Nothing to Copy",
                                f"The {name} table is empty \u2014 nothing to copy yet.")
            return
        text = table.to_tsv()
        if not text.strip():
            messagebox.showinfo("Nothing to Copy",
                                f"The {name} table is empty \u2014 nothing to copy yet.")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.app.log_message(f"Copied {rows:,} rows ({name}) to clipboard.")

    def _copy_totals(self):
        totals = self.app._build_totals_df()
        if totals is None or totals.empty:
            messagebox.showinfo("Nothing to Copy",
                                "The Totals by Type table is empty \u2014 nothing to copy yet.")
            return
        lines = ["\t".join(str(c) for c in totals.columns)]
        for row in totals.itertuples(index=False):
            lines.append("\t".join(str(v) for v in row))
        text = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.app.log_message(f"Copied totals by type ({len(lines) - 1} groups) to clipboard.")

    def restart(self):
        if messagebox.askyesno("Restart", "Start over from step 1?"):
            self.app.restart()


class DBFExtractorValidatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DBF / SQL Data Extractor & Excel Validator By Kuya Daks")
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
        self.validator_type = None
        self.extracted_df = pd.DataFrame()
        self.excel_df = pd.DataFrame()
        self.validated_df = pd.DataFrame()
        self.acct_col = None
        self.name_col = None
        self.start_date = None
        self.end_date = None
        self.total_records = 0
        self.total_amount = 0.0
        self.total_loan_amount = 0.0
        self.matches = 0
        self.validated_amount = 0.0
        self.validated_loan_amount = 0.0
        self.transaction_totals = {}
        self.src_columns = []
        self.excel_columns = []
        self.role_match_col = None
        self.role_sum_cols = []
        self.role_date_col = None
        self.role_type_col = None
        self.total_amounts = {}
        self.validated_amounts = {}
        self.current_step = None
        self.log_visible = False
        self._busy = False
        self._bg_queue = queue.Queue()
        self._results_seq = 0
        self._loading_start = time.time()

        self.create_widgets()
        self._poll_bg_queue()
        self.log_message("System ready. Load a .dbf, .csv or .xml file to start.")
        self.show_step(1)

    def create_widgets(self):
        self.progress_bar = StepProgressBar(self)
        self.progress_bar.pack(side="top", fill="x", padx=15, pady=(10, 4))

        self.log_toggle_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.log_toggle_frame.pack(side="bottom", fill="x", padx=15, pady=(0, 4))

        self.btn_toggle_log = ctk.CTkButton(
            self.log_toggle_frame, text="Show Log", width=90, height=26,
            font=ctk.CTkFont(size=11), fg_color="#333", hover_color="#555",
            command=self.toggle_log)
        self.btn_toggle_log.pack(side="left")

        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(side="top", fill="both", expand=True, padx=15, pady=(0, 4))

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
            dbf_disp, excel_disp, excel_cols, totals_df, seq = item[1]
            if not isinstance(self.current_step, ResultsStep):
                return
            if dbf_disp is None:
                err = getattr(self, "_results_error", None) or "unknown error"
                self.log_message(f"Failed to build result tables:\n{err}")
                if getattr(self.current_step, "lbl_loading", None) is not None:
                    self.current_step.lbl_loading.configure(
                        text="Failed to build result tables \u2014 check the Log.",
                        text_color="#e59400")
                return
            if seq != getattr(self, "_results_seq", -1):
                return
            try:
                self.current_step.load_frames(dbf_disp, excel_disp, excel_cols, totals_df)
            except Exception:
                self.log_message(f"Error rendering result tables:\n{traceback.format_exc()}")

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
        self.matches = 0
        self.validated_amount = 0.0
        self.validated_loan_amount = 0.0
        self.validated_amounts = {}
        self.transaction_totals = {}
        self._run_bg(self.current_step, "Extracting data...",
                     self._extract_work, self._after_extract)

    def probe_source_columns(self, step):
        if self._busy:
            return

        def work():
            cols, err = probe_data_columns(self.dbf_path)
            return {"cols": cols, "err": err}

        def on_done(r):
            self.log_message(f"Detected columns: {', '.join(r['cols']) if r['cols'] else '(none)'}")
            if callable(getattr(step, "on_columns_probed", None)):
                step.on_columns_probed(r["cols"], r["err"])

        self._run_bg(step, "Detecting columns...", work, on_done)

    def _extract_work(self):
        if self.validator_type == "sqlsummary":
            df, err = read_sql_table_file(self.dbf_path)
            if err:
                return {"error": err}
        else:
            try:
                table = DBF(self.dbf_path, encoding="latin-1",
                            ignore_missing_memofile=True)
                data = []
                for record in table:
                    data.append(dict(record))
                    if len(data) % 10000 == 0:
                        time.sleep(0.002)
                df = pd.DataFrame(data)
            except Exception as e:
                return {"error": f"Failed to read DBF file: {e}"}
        if df is None or df.empty:
            return {"empty": True}
        return self._apply_roles_to_df(df)

    def _apply_roles_to_df(self, df):
        """Apply the selected column roles: date filter + numeric sum columns."""
        date_col = self.role_date_col
        if date_col and date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
            df = df[df[date_col].notna()]
            if self.start_date:
                df = df[df[date_col].dt.date >= self.start_date]
            if self.end_date:
                df = df[df[date_col].dt.date <= self.end_date]
        for sc in self.role_sum_cols or []:
            if sc in df.columns:
                df[f"_SUM_NUM_{sc}"] = df[sc].apply(parse_money)
        self.extracted_df = df
        self.total_records = len(df)
        self.total_amounts = {}
        self.total_amount = 0.0
        self.total_loan_amount = 0.0
        for sc in self.role_sum_cols or []:
            c = f"_SUM_NUM_{sc}"
            if c in df.columns:
                amt = float(df[c].sum())
                self.total_amounts[sc] = amt
                self.total_amount += amt
        return {"empty": df.empty}

    def _after_extract(self, result):
        if result.get("error"):
            messagebox.showerror("Error", result["error"])
            self.log_message(f"Extraction error: {result['error']}")
            return
        if result["empty"]:
            messagebox.showinfo("Info", "No records found in the selected date range.")
            self.log_message("Extraction returned 0 records.")
            return
        sums = ", ".join(f"{k} {v:,.2f}" for k, v in self.total_amounts.items()) or "no sum columns"
        self.log_message(f"Extracted {self.total_records:,} records (Totals: {sums}).")
        self.show_step(2)

    def probe_excel(self, step):
        if self._busy:
            return

        def work():
            excel_df = pd.read_excel(self.excel_path)
            return {"df": excel_df, "acct_col": find_account_column(excel_df),
                    "name_col": find_name_column(excel_df),
                    "columns": [str(c) for c in excel_df.columns]}

        def on_done(r):
            self.excel_df = r["df"]
            self.acct_col = r["acct_col"]
            self.name_col = r["name_col"]
            self.log_message(f"Excel file loaded: {len(r['df']):,} rows; "
                             f"match column: {r['acct_col']} | name column: {r['name_col']}")
            if callable(getattr(step, "on_excel_ready", None)):
                step.on_excel_ready(r["acct_col"], r["name_col"], len(r["df"]),
                                    r.get("columns"))

        self._run_bg(step, "Reading Excel file...", work, on_done)

    def proceed_to_results(self):
        if self._busy:
            return
        self._run_bg(self.current_step, "Validating records...",
                     self._validate_work, self._after_validate)

    def _has_sum_col(self, sc):
        return f"_SUM_NUM_{sc}" in getattr(self, "extracted_df", pd.DataFrame()).columns

    def _src_name_cols(self):
        cols = list(getattr(self, "extracted_df", pd.DataFrame()).columns)
        if self.validator_type == "sqlsummary":
            picked = [c for c in ("lname", "fname", "middle") if c in cols]
            if not picked and "accountname" in cols:
                picked = ["accountname"]
            if picked:
                return picked
        return [c for c in cols if "name" in str(c).lower()]

    def _validate_work(self):
        if self.excel_df is None or self.excel_df.empty:
            if not self.excel_path:
                return {"error": "No Excel file selected."}
            self.excel_df = pd.read_excel(self.excel_path)
        excel_df = self.excel_df
        if self.acct_col is None:
            self.acct_col = find_account_column(excel_df)
        if self.name_col is None:
            self.name_col = find_name_column(excel_df)
        if self.acct_col is None:
            return {"error": "Could not find a Match column in the Excel file. "
                             "Tick one on Step 2."}
        if not self.role_match_col or self.role_match_col not in self.extracted_df.columns:
            return {"error": "The Match column selected on Step 1 is not in the data file. "
                             "Go back and pick the correct column."}

        excel_df["_CLEAN_EXCEL_MATCH"] = excel_df[self.acct_col].astype(str).apply(
            lambda x: re.sub(r"[^0-9a-zA-Z]", "", x))
        excel_match_set = set(excel_df["_CLEAN_EXCEL_MATCH"])

        if not self.extracted_df.empty:
            self.extracted_df["_CLEAN_MATCH"] = self.extracted_df[self.role_match_col].astype(str).apply(
                lambda x: re.sub(r"[^0-9a-zA-Z]", "", x))
            self.extracted_df["MATCHED_IN_EXCEL"] = self.extracted_df["_CLEAN_MATCH"].isin(
                excel_match_set)
            self._apply_name_fallback(excel_df, *self._src_name_cols())
            matched = self.extracted_df[self.extracted_df["MATCHED_IN_EXCEL"]]
            self.matches = len(matched)
            self.validated_amounts = {}
            self.validated_amount = 0.0
            self.validated_loan_amount = 0.0
            for sc in self.role_sum_cols or []:
                c = f"_SUM_NUM_{sc}"
                if c in self.extracted_df.columns:
                    amt = float(matched[c].sum()) if not matched.empty else 0.0
                    self.validated_amounts[sc] = amt
                    self.validated_amount += amt
            self.transaction_totals = self._compute_transaction_totals()
        else:
            self.matches = 0
            self.validated_amount = 0.0
            self.validated_loan_amount = 0.0
            self.validated_amounts = {}
            self.extracted_df["MATCHED_IN_EXCEL"] = False
            self.transaction_totals = {}
        return {"error": None}

    def _apply_name_fallback(self, excel_df, *src_name_cols):
        """Name-based fallback for records whose account number did not match.

        Builds the source-side name from src_name_cols (e.g. CUSTNAME for DBF,
        lname/fname/middle for SQL)) and matches it against the Excel name column.
        Marks extra extracted rows as matched and sets excel_df['_NAME_MATCHED'].
        """
        self.name_fallback_matches = 0
        if "_NAME_MATCHED" in excel_df.columns:
            excel_df.drop(columns=["_NAME_MATCHED"], inplace=True)
        if not src_name_cols or not self.name_col or self.name_col not in excel_df.columns:
            return
        present = [c for c in src_name_cols if c in self.extracted_df.columns]
        if not present:
            return
        parts = [self.extracted_df[c].fillna("").astype(str) for c in present]
        src_full = parts[0]
        for p in parts[1:]:
            src_full = src_full + " " + p
        name_hits = self._match_names(
            src_full.map(name_tokens), excel_df[self.name_col].map(name_tokens))
        self.name_fallback_matches = len(name_hits["src_positions"])
        matched_rows = pd.Series(False, index=self.extracted_df.index)
        matched_rows.iloc[list(name_hits["src_positions"])] = True
        self.extracted_df.loc[matched_rows, "MATCHED_IN_EXCEL"] = True
        em = pd.Series(False, index=excel_df.index)
        em.iloc[list(name_hits["excel_positions"])] = True
        excel_df["_NAME_MATCHED"] = em

    def _match_names(self, src_names, excel_names):
        """Return dict of positional sets for name matches.

        A source record matches an Excel row when the smaller token set is fully
        contained in the larger one and they share at least 2 name tokens (a
        missing middle name / initial on either side still matches).
        """
        src_positions = set()
        excel_positions = set()
        if src_names.empty or excel_names.empty:
            return {"src_positions": src_positions, "excel_positions": excel_positions}
        index = {}
        for pos, toks in enumerate(src_names):
            for tok in toks:
                index.setdefault(tok, set()).add(pos)
        if not index:
            return {"src_positions": src_positions, "excel_positions": excel_positions}
        for epos, excel_toks in enumerate(excel_names):
            if not excel_toks:
                continue
            cand = None
            for tok in excel_toks:
                s = index.get(tok)
                if not s:
                    cand = None
                    break
                cand = s if cand is None else (cand & s)
                if not cand:
                    break
            if not cand:
                continue
            for spos in cand:
                src_toks = src_names.iloc[spos]
                common = src_toks & excel_toks
                if (len(common) >= 2 and
                        (src_toks <= excel_toks or excel_toks <= src_toks)):
                    src_positions.add(spos)
                    excel_positions.add(epos)
            if epos % 10000 == 0:
                time.sleep(0.002)
        return {"src_positions": src_positions, "excel_positions": excel_positions}

    def _compute_transaction_totals(self):
        totals = {}
        if self.extracted_df.empty:
            return totals
        type_col = self.role_type_col
        if not type_col or type_col not in self.extracted_df.columns:
            return totals
        sum_cols = [sc for sc in (self.role_sum_cols or []) if self._has_sum_col(sc)]
        grouped = self.extracted_df.groupby(type_col)
        for gval, grp in grouped:
            matched = grp[grp["MATCHED_IN_EXCEL"]]
            entry = {
                "group": str(gval),
                "count": int(grp.shape[0]),
                "matched_count": int(matched.shape[0]),
                "totals": {},
                "matched_totals": {},
            }
            for sc in sum_cols:
                c = f"_SUM_NUM_{sc}"
                entry["totals"][sc] = float(grp[c].sum())
                entry["matched_totals"][sc] = float(matched[c].sum()) if not matched.empty else 0.0
            totals[str(gval)] = entry
        return dict(sorted(totals.items(),
                           key=lambda kv: sum(kv[1]["totals"].values()), reverse=True))

    def _build_totals_df(self):
        sum_cols = [sc for sc in (self.role_sum_cols or []) if self._has_sum_col(sc)]
        cols = ["Group", "Count"]
        cols += [f"Total {sc}" for sc in sum_cols]
        cols += ["Matched Count"] + [f"Matched {sc}" for sc in sum_cols]
        if not getattr(self, "transaction_totals", None):
            return pd.DataFrame(columns=cols)
        rows = []
        for g, d in self.transaction_totals.items():
            row = {"Group": d["group"], "Count": d["count"]}
            for sc in sum_cols:
                row[f"Total {sc}"] = f"{d['totals'].get(sc, 0.0):,.2f}"
            row["Matched Count"] = d["matched_count"]
            for sc in sum_cols:
                row[f"Matched {sc}"] = f"{d['matched_totals'].get(sc, 0.0):,.2f}"
            rows.append(row)
        return pd.DataFrame(rows, columns=cols)

    def _after_validate(self, result):
        if result.get("error"):
            messagebox.showerror("Error", result["error"])
            self.log_message(f"Validation error: {result['error']}")
            return
        name_note = ""
        if getattr(self, "name_fallback_matches", 0):
            name_note = f" | Name fallback matches: {self.name_fallback_matches:,}"
        sums = ", ".join(f"{k} {v:,.2f}" for k, v in self.validated_amounts.items()) or "no sum columns"
        self.log_message(f"Validation complete! Matches: {self.matches:,} "
                         f"| Matched totals: {sums}{name_note}")
        self.show_step(3)
        self.populate_results()

    def populate_results(self, show_all=False):
        if not isinstance(self.current_step, ResultsStep):
            return
        self._results_seq = getattr(self, "_results_seq", 0) + 1
        seq = self._results_seq
        self._results_error = None

        def work():
            try:
                dbf_disp = self.get_dbf_display_df(full=show_all)
                excel_disp = self.get_excel_display_df(full=show_all)
                excel_cols = [c for c in excel_disp.columns if c != "_IS_MATCH"]
                totals_df = self._build_totals_df() if getattr(
                    self, "role_type_col", None) else None
                return ("results_frames", (dbf_disp, excel_disp, excel_cols, totals_df, seq))
            except Exception:
                err = traceback.format_exc()
                self._results_error = err
                return ("results_frames", (None, None, [], None, seq))

        threading.Thread(target=lambda: self.post(work()), daemon=True).start()

    def _internal_cols(self):
        """Columns that are internal to the matcher and hidden from display/export."""
        return set(c for c in self.extracted_df.columns if c.startswith("_"))

    def get_dbf_display_df(self, full=False):
        if self.extracted_df.empty:
            return pd.DataFrame()
        if full:
            cols = [c for c in self.extracted_df.columns
                    if c not in self._internal_cols()]
            return self.extracted_df[cols].copy()
        disp = pd.DataFrame()
        seen = set()
        for c in (self.role_match_col, self.role_date_col, self.role_type_col):
            if c and c in self.extracted_df.columns and c not in seen:
                seen.add(c)
                if c in (self.role_sum_cols or []):
                    disp[c] = self._fmt_display_amount(self.extracted_df[c])
                else:
                    disp[c] = self.extracted_df[c]
        for sc in self.role_sum_cols or []:
            if sc in self.extracted_df.columns and sc not in seen:
                seen.add(sc)
                disp[sc] = self._fmt_display_amount(self.extracted_df[sc])
        if "MATCHED_IN_EXCEL" in self.extracted_df.columns:
            disp["Matched in Excel"] = self.extracted_df["MATCHED_IN_EXCEL"]
        return disp

    def _fmt_display_amount(self, series):
        return series.astype(str).apply(
            lambda s: "" if not s.strip() else f"{parse_money(s):,.2f}")

    def get_excel_display_df(self, full=False):
        if self.excel_df is None or self.excel_df.empty:
            return pd.DataFrame()
        internal = {"_CLEAN_EXCEL_MATCH", "_NAME_MATCHED"}
        if full:
            excel_display_cols = [c for c in self.excel_df.columns if c not in internal]
        else:
            picked = []
            for c in (self.acct_col, self.name_col):
                if c and c in self.excel_df.columns and c not in picked:
                    picked.append(c)
            dbf_like = [c for c in DBF_COLS if c in self.excel_df.columns]
            if picked:
                excel_display_cols = picked
            elif dbf_like:
                excel_display_cols = dbf_like
            else:
                excel_display_cols = [c for c in self.excel_df.columns if c not in internal][:6]
        disp = self.excel_df[excel_display_cols].copy()
        acct_match = pd.Series(False, index=disp.index)
        if "_CLEAN_EXCEL_MATCH" in self.excel_df.columns and not self.extracted_df.empty:
            dbf_set = set(self.extracted_df["_CLEAN_MATCH"]) \
                if "_CLEAN_MATCH" in self.extracted_df.columns else set()
            acct_match = self.excel_df["_CLEAN_EXCEL_MATCH"].isin(dbf_set)
        name_match = (self.excel_df["_NAME_MATCHED"]
                      if "_NAME_MATCHED" in self.excel_df.columns
                      else pd.Series(False, index=self.excel_df.index))
        disp["_IS_MATCH"] = (acct_match | name_match).reindex(disp.index).fillna(False)
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
                internal = self._internal_cols()
                cols_to_export = [c for c in self.extracted_df.columns if c not in internal]
                export_df = self.extracted_df[cols_to_export].copy()

                for col in export_df.columns:
                    if export_df[col].dtype == "object":
                        export_df[col] = export_df[col].apply(clean_illegal_chars)

                if (self.role_date_col and self.role_date_col in export_df.columns
                        and pd.api.types.is_datetime64_any_dtype(export_df[self.role_date_col])):
                    export_df[self.role_date_col] = export_df[self.role_date_col].dt.strftime("%Y-%m-%d")

                if getattr(self, "transaction_totals", None):
                    totals_df = self._build_totals_df()
                    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                        export_df.to_excel(writer, index=False, sheet_name="Records")
                        totals_df.to_excel(writer, index=False, sheet_name="Totals by Type")
                else:
                    export_df.to_excel(save_path, index=False, engine="openpyxl")
                messagebox.showinfo("Success", f"Data exported successfully to:\n{save_path}")
                self.log_message(f"Exported data to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save export file:\n{str(e)}")

    def restart(self):
        self.dbf_path = ""
        self.excel_path = ""
        self.validator_type = None
        self.extracted_df = pd.DataFrame()
        self.excel_df = pd.DataFrame()
        self.validated_df = pd.DataFrame()
        self.acct_col = None
        self.name_col = None
        self.start_date = None
        self.end_date = None
        self.total_records = 0
        self.total_amount = 0.0
        self.total_loan_amount = 0.0
        self.matches = 0
        self.name_fallback_matches = 0
        self.validated_amount = 0.0
        self.validated_loan_amount = 0.0
        self.transaction_totals = {}
        self.src_columns = []
        self.excel_columns = []
        self.role_match_col = None
        self.role_sum_cols = []
        self.role_date_col = None
        self.role_type_col = None
        self.total_amounts = {}
        self.validated_amounts = {}
        self.show_step(1)


if __name__ == "__main__":
    app = DBFExtractorValidatorApp()
    app.mainloop()