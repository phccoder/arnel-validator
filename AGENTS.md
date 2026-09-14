# AGENTS.md

## Project

Single-file desktop app (`dbf_extractor_validator.py`) built with Python **3.14** + CustomTkinter + pandas + dbfread + tkcalendar. It reads a data file (a DBF file, or a SQL `tblsummaryreceipt` table exported as CSV/XML), filters records by date range, then validates them against an uploaded Excel file (account-number matching), showing side-by-side results with match highlighting.

Run: `python dbf_extractor_validator.py`

## Architecture

- **3-step full-screen wizard** controlled by `DBFExtractorValidatorApp`:
  1. `WelcomeStep` data file pick + **column-role checkboxes** + calendar date range. Smart auto-detect: `.dbf` → `validator_type = "dbf"`; `.csv`/`.xml` → `validator_type = "sqlsummary"`. `probe_data_columns` reads the file's columns without loading data (DBF via `field_names`; CSV via `pd.read_csv(nrows=0)`; XML via `ET.iterparse` first record) and `detect_column_roles` pre-checks known names (`accountno`/`ACCTNBR` → match; `ALVAMT`/`cashout`/`loanamount`/…"amount"/"cash"/"loan" → sums; *date* → date; *type* → group-by).
  2. `ExcelStep` Excel upload; the match (account) and optional name columns are also chosen via checkboxes, pre-checked with `find_account_column`/`find_name_column`.
  3. `ResultsStep` dashboard cards + "Totals by Type" table + side-by-side `ScrollableTable`s + Export/Restart/Back.
- Controller owns all app state (`extracted_df`, `excel_df`, `validator_type`, counters, `transaction_totals`) and steps are thin views. Role selections live on the controller: `role_match_col` (str), `role_sum_cols` (list), `role_date_col`, `role_type_col` (optional), plus `acct_col`/`name_col` for the Excel side.
- **Column roles are user-selectable** via `RoleCheckMatrix` (row per column, checkbox per role) for **Match / Sum / Type** only. Match and Type are single-select (enforced with the just-clicked column winning); Sum is multi-select. The **Date column is NOT user-selectable** — it is auto-detected by `detect_column_roles` (`*date*` in the name) and `role_date_col` is set automatically to the detected column; the calendar date range always filters that column.
- Matching is generic: the selected `role_match_col` is cleaned (regex `[^0-9a-zA-Z]` removed) vs the cleaned Excel `acct_col` → `MATCHED_IN_EXCEL` (source side) and `_IS_MATCH` (Excel side). Optional name fallback uses `name_col` and the source's name-ish column(s) (`_src_name_cols()`).
- Extraction/validation are one generic pipeline (`_extract_work`/`_validate_work`) driven by the roles. The auto-detected `role_date_col` is parsed then date-filtered, if a date-like column exists (records are NOT silently mis-filtered; the mode badge warns "No date column detected" otherwise). Each sum col gets `_SUM_NUM_<col>` via `parse_money`; per-column totals live in `total_amounts`/`validated_amounts` dicts.
- SQL reads: `read_sql_table_file` → `_read_csv_any` (utf-8-sig → latin-1; comma/tab/semicolon; rejects MultiIndex = unquoted commas) or `_read_xml_any` (xml.etree for both `<row><field name=.../>` mysqldump and `<DATA_RECORD><tag>...</tag>` phpMyAdmin/paw forms + `pd.read_xml` fallback; strips whitespace + HTML-unescapes field values). DBF reads via dbfread record→dict loop with yields.
- Totals by type: `_compute_transaction_totals` groups the filtered extract by `role_type_col` (if set) into `{group, count, matched_count, totals:{sumcol:amt}, matched_totals:{sumcol:amt}}` sorted by combined totals desc; `_build_totals_df` renders it as `Group | Count | Total <sum1> | … | Matched Count | Matched <sum1> | …`. Empty if no type role.
- Dashboard cards: Total Records + Number of Matches, plus one "Total <col>" and "Matched <col>" card per selected sum column — built dynamically in `ResultsStep.__init__`.
- Internal columns (`_CLEAN_MATCH`, `_SUM_NUM_*`, `_CLEAN_EXCEL_MATCH`, `_NAME_MATCHED`, …) all start with `_` and are hidden via `_internal_cols()` (prefix filter); `MATCHED_IN_EXCEL` IS shown.
- Export: "Records" sheet with all non-internal columns + `MATCHED_IN_EXCEL`; when a type role produced totals, an extra "Totals by Type" sheet. Date column strftime'd before export.

## Critical rules (do not break)

- **NEVER call Tk from a worker thread.** On Python 3.14, `root.after()` from a thread raises `RuntimeError: main thread is not in main loop`. All background work must post to `self._bg_queue` (a `queue.Queue`) and the main thread polls via `_poll_bg_queue` (50ms) → `_dispatch_bg`.
- Queue message kinds in `_dispatch_bg`: `("bg_done", (step, result, on_done))`, `("table_ready", (table, seq, positions, err))`, `("results_frames", (dbf_disp, excel_disp, excel_cols, totals_df, seq))` (5-tuple since SQL mode).
- **Row rendering must stay bounded.** `ScrollableTable` renders only `PAGE_SIZE` (250) rows at a time + Prev/Next. CTk widgets are ~5x slower than plain `tk.Frame`/`tk.Label` — table body rows use tk widgets; headers/Cards/chrome use CTk. The table body is a `tk.Canvas` with horizontal + vertical scrollbars (header lives in a second canvas whose xview is synced via `_xview`; both scrollregions sized by `_sync_widths`). Result panes are always exactly half width via `uniform="resultpane"` grid columns.
- **All data prep/search off the UI thread:** `compute_positions()` builds/filters a cached lower-cased concatenated key Series with vectorized pandas (no `agg(" ".join, axis=1)`, no row lambdas). `set_data` stores the df reference (no copy) and kicks off async compute; search is debounced 250ms.
- **Staleness guards:** `ScrollableTable._filter_seq` — `apply_compute` ignores results whose seq != current. `App._results_seq` guards `results_frames`. `clear()`/`set_columns()` must bump the seq and cancel pending `after`.
- **Loading feedback:** `_run_bg` shows an indeterminate bar; `MIN_LOADING_TIME` (1.0s) is enforced in `_finish_bg` so fast ops still show the bar. `update_idletasks` + indeterminate mode keep it animating.
- **GIL vs. the Tk loop:** pure-Python worker loops (e.g. dbfread record→dict conversion in `_extract_work`, `re.sub` per row) deny the main thread the GIL and freeze the UI. Must `sys.setswitchinterval(0.001)` at module import, AND yield inside hot loops (`time.sleep(0.002)` every ~10k records in `_extract_work`). pandas/openpyxl vectorized ops release the GIL and need no yielding.
- Type issues: convert `pd.Timestamp` to `strftime` before display; strings must be cleaned of illegal Excel chars via `clean_illegal_chars` before exporting.
- Everything lives in the ONE file. Keep all changes there; no extra project files needed.

## Verification

- Syntax: `python -m py_compile dbf_extractor_validator.py`
- Functional smoke tests instantiate the real `DBFExtractorValidatorApp` in a desktop session, `withdraw()` it, and pump `app.update()` until async results land (wait for `t._computing`/row counts, not fixed sleeps).
- Edge cases to keep passing: empty extract, empty search query, stale async results dropped, 100k-row load without blocking, min-loading-time, header count matches columns.