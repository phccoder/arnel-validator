![DBF Extractor & Excel Validator](icon.png)

# DBF Extractor & Excel Validator

A single-file desktop app that **extracts records from a DBF file** (filtered by date range) and **validates them against an Excel file** by matching account numbers, then shows the results side-by-side with match highlighting.

Built with Python 3.14 + CustomTkinter + pandas + dbfread + tkcalendar. Distributable as a **portable single-file `.exe`** (no installation required on the target machine).

---

## Features

- **3-step full-screen wizard** — pick DBF file + date range → upload Excel → review results
- **Date-range filtering** with calendar pickers (start / end)
- **Automatic account-column detection** in the Excel file (column name containing *"account"*)
- **Side-by-side results** — DBF records on the left, Excel data on the right
- **Match highlighting** — matched rows are highlighted green on both sides
- **Live search** — debounced, filters thousands of rows without blocking the UI
- **Summary dashboard** — total records, total amount, matches found, validated amount
- **Export results to Excel** — cleaned of illegal Excel characters, dates formatted
- **Background processing** — heavy work (DBF extraction, validation, table building) runs off the UI thread; the window never freezes, even with 100k+ records
- **Portable build** — one `.exe`, no Python required on end-user machines

---

## Requirements

- Windows 10/11
- Python **3.14+**
- A DBF file with the columns used by the app (see [Data Format](#data-format))
- An Excel file (`.xlsx` / `.xls`) containing an account-number column

Python dependencies (`pip install -r requirements.txt`):

| Package        | Version |
| -------------- | ------- |
| customtkinter  | 6.0.0   |
| pandas         | 2.3.3   |
| dbfread        | 2.0.7   |
| tkcalendar     | 1.6.1   |
| Pillow         | 12.3.0  |
| openpyxl       | 3.1.5   |

---

## Getting Started

### Run from source

```bash
# 1. Create a virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
python dbf_extractor_validator.py
```

### Use the portable .exe

Download `DBFExtractorValidator.exe` from the [Releases](../../releases) section (or build it yourself, see below) and double-click it. No Python or dependencies are needed. First launch takes a few seconds while the single-file exe unpacks itself.

---

## How to Use (3 Steps)

1. **Step 1 — DBF & Date Range**
   Click **Select DBF File**, choose your `.dbf`, then pick the **Start Date** and **End Date**, and click **Continue to Step 2**.

2. **Step 2 — Excel Upload**
   Click **Select Excel File** and choose your `.xlsx` / `.xls`. The app auto-detects the account column (column name containing *"account"*). Click **Run Validation & View Results**.

3. **Step 3 — Results**
   Review the summary dashboard and the side-by-side tables:
   - **DBF Extracted Data** (left) — filtered records with a `Matched in Excel` highlight
   - **Excel Data** (right) — uploaded rows with a `Matched` highlight
   - Search boxes filter each table live
   - **Export Results to Excel** saves the validated DBF records
   - **Restart** starts over from Step 1

---

## Data Format

### DBF file

The app reads the following columns from the DBF (other columns are ignored):

| Column      | Description                 |
| ----------- | --------------------------- |
| `ALVNBR`    | ALV number                  |
| `ALVDATE`   | Date of the record          |
| `ACCTNBR`   | Account number (matched)    |
| `CUSTCODE`  | Customer code               |
| `CUSTNAME`  | Customer name               |
| `ALVAMT`    | Amount                       |

### Matching logic

Account numbers on **both** sides are cleaned by removing every non-alphanumeric character (regex `[^0-9a-zA-Z]`) before comparing, so formatting differences (−, spaces, dashes) don't break matches.

---

## Building the Portable .exe

[PyInstaller](https://pyinstaller.org/) is required. Everything (dependencies + icon) is bundled into a single `dist\DBFExtractorValidator.exe`.

```bash
pip install pyinstaller

python -m PyInstaller --noconfirm --onefile --windowed ^
  --name "DBFExtractorValidator" ^
  --icon "icon.png" ^
  --add-data "icon.png;." ^
  --collect-data customtkinter --collect-data tkcalendar ^
  --exclude-module IPython --exclude-module jedi --exclude-module pytest ^
  --exclude-module pygments --exclude-module parso --exclude-module sqlalchemy ^
  --exclude-module lxml --exclude-module matplotlib --exclude-module scipy ^
  --exclude-module imageio --exclude-module tests ^
  dbf_extractor_validator.py
```

> The `--exclude-module` flags remove large unused packages that pandas' build hooks pull in (IPython, jedi, pytest, pygments, lxml, sqlalchemy, matplotlib, ...), cutting the exe from ~78 MB to ~58 MB and speeding up startup.

---

## Project Structure

```
arnel-validator/
├── dbf_extractor_validator.py   # the entire application (single file)
├── requirements.txt             # Python dependencies
├── icon.png                     # app / exe icon
├── AGENTS.md                    # developer guidance for AI agents / contributors
└── DBFExtractorValidator.spec   # PyInstaller spec (generated during build)
```

---

## License

Distributed under the MIT License.