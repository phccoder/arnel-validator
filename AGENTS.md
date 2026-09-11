# AGENTS.md

## Project

Single-file desktop app (`dbf_extractor_validator.py`) built with Python **3.14** + CustomTkinter + pandas + dbfread + tkcalendar. It reads a DBF file, filters records by date range, then validates them against an uploaded Excel file (account-number matching), showing side-by-side results with match highlighting.

Run: `python dbf_extractor_validator.py`

## Architecture

- **3-step full-screen wizard** controlled by `DBFExtractorValidatorApp`:
  1. `WelcomeStep` DBF file pick + calendar date range
  2. `ExcelStep` Excel upload (auto-detects account column: contains "account")
  3. `ResultsStep` summary dashboard cards + side-by-side `ScrollableTable`s + Export/Restart/Back
- Controller owns all app state (`extracted_df`, `excel_df`, counters) and steps are thin views.
- DBF columns: `ALVNBR ALVDATE ACCTNBR CUSTCODE CUSTNAME ALVAMT`. Matching uses cleaned account numbers (regex `[^0-9a-zA-Z]` removed) via `MATCHED_IN_EXCEL` (DBF side) and `_IS_MATCH` (Excel side).

## Critical rules (do not break)

- **NEVER call Tk from a worker thread.** On Python 3.14, `root.after()` from a thread raises `RuntimeError: main thread is not in main loop`. All background work must post to `self._bg_queue` (a `queue.Queue`) and the main thread polls via `_poll_bg_queue` (50ms) → `_dispatch_bg`.
- Queue message kinds in `_dispatch_bg`: `("bg_done", (step, result, on_done))`, `("table_ready", (table, seq, positions, err))`, `("results_frames", (dbf_disp, excel_disp, excel_cols, seq))`.
- **Row rendering must stay bounded.** `ScrollableTable` renders only `PAGE_SIZE` (250) rows at a time + Prev/Next. CTk widgets are ~5x slower than plain `tk.Frame`/`tk.Label` — table body rows use tk widgets; headers/Cards/chrome use CTk.
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