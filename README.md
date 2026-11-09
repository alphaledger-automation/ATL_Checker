# ATL Checker

Bulk Active Taxpayer List (ATL) status checker for Pakistan. Feed an NTN/CNIC list, get filer / non-filer status in one pass.

Reads registration numbers from an Excel file, checks each one against the FBR IRIS Active Taxpayer List (Income Tax) as of a date you choose, saves the verification slip as a PDF, and writes an Excel summary with name, filing status and checking date.

Browser automation runs headless in the background — it does not take over your mouse or keyboard.

---

## Requirements

- Windows 10/11
- Python 3.10+

Everything else is installed on first run. `run.bat` will install Python itself via `winget` if it is missing.

## Install and run

```bash
git clone https://github.com/alphaledger-automation/ATL_Checker.git
cd ATL_Checker
pip install -r requirements.txt
python main.py
```

Or, on Windows, just double-click **`run.bat`** — it finds Python, and `main.py` then installs any missing dependency itself:

- Python packages (`customtkinter`, `playwright`, `openpyxl`, `pillow`, `pytesseract`, `pdfplumber`)
- Playwright's Chromium build (~150 MB, once)
- Tesseract-OCR (Windows may prompt to allow the installer)

Prepare dependencies without launching the GUI:

```bash
python main.py --setup-only
```

## Input file

- First row must be the column headers.
- The registration column is auto-detected (`Registration Number`, `Registration No`, `NTN`, `CNIC`, …); pick it manually if detection misses.
- 7 characters = NTN (e.g. `1234567`, `E123456`); 13 digits = CNIC (dashes ignored).
- Duplicate numbers are checked once.
- Your input file is never modified — results go to a separate summary workbook.

## Output

```
<PDF folder>/<RegNo>_<DD-MM-YYYY>.pdf          one slip per number
<PDF folder>/ATL_Summary_<DD-MM-YYYY>.xlsx     summary workbook
```

Summary columns: `Length | Registration Number | file name | Status | Name | Filing Status | Filing Status Checking Date | ATL_PDFs Status`

The workbook is created when the run starts and rewritten after every number, so it is always current. Keep it closed in Excel during a run.

Re-running with the same date and folder resumes: numbers that already have a PDF are reported `Already downloaded` and skipped, unless you tick **Re-download existing PDFs**.

## Features

- Dashboard page: file picker, registration-column selector, checking date (DD/MM/YYYY), PDF output folder
- Current Status page: Done / Total counter and a live row per number with name, filing status and result
- Pause / Resume and Cancel, applied after the number in flight finishes
- Light / Dark theme, remembered between launches
- Optional visible browser window for diagnosing a failing run
- Captcha retried up to 5 times per number; a failed number stays pending so a second Start retries only what did not succeed
- Incremental Excel writing — a crash or cancel never loses completed rows

## Project layout

| File | Purpose |
|---|---|
| `run.bat` | Windows launcher (finds or installs Python) |
| `main.py` | Entry point; `--setup-only` flag |
| `bootstrap.py` | Dependency check and installer |
| `gui.py` | Dashboard / Current Status window |
| `browser_engine.py` | IRIS automation |
| `excel_handler.py` | Reads the input workbook |
| `report_writer.py` | Writes the Excel summary |
| `constants.py` | URLs, selectors, timings, palette |
| `Assets/` | Application icons |

`USER_GUIDE.txt` holds the end-user walkthrough and troubleshooting notes.

## Notes

- No credentials are stored or required — the ATL check is a public IRIS lookup.
- No taxpayer data, sample workbooks or output PDFs are included in this repository.
- Setup log: `%LocalAppData%\ATL_Checker\setup.log`

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 Alpha Edge Advisory.
