"""
ATL Checker — Constants & Configuration
=======================================
All URLs, selectors, timing values and column-detection aliases live here
so the rest of the code stays clean.
"""

import pathlib

# ── Target URL ────────────────────────────────────────────────────────────────
IRIS_URL = "https://iris.fbr.gov.pk/iris/home.jsf"

# ── DOM Selectors (IRIS 2.0 "Active Taxpayer List (Income Tax)" form) ────────
SEL_ATL_LINK       = "a:has-text('Active Taxpayer List (Income Tax)')"
SEL_PARAM_BOX      = "#parameterType"
SEL_PARAM_CURRENT  = "#parameterType .select-box__current"
SEL_PARAM_OPTION   = "#parameterType label.select-box__option"
SEL_REG_INPUT      = "#regNo"
SEL_DATE_INPUT     = "#date"
SEL_CAPTCHA_CANVAS = "canvas.captcha-canvas"
SEL_CAPTCHA_INPUT  = "input[placeholder='Enter '], input[placeholder='Enter']"
SEL_CAPTCHA_REFRESH = "button:has(i.fa-refresh)"
SEL_VERIFY_BUTTON  = "button.active-btn-green:has-text('Verify'), button:has-text('Verify')"

# Text that only appears once a verification result is on screen
RESULT_MARKER    = "Filing Status Checking Date"
NO_RECORD_MARKER = "no record exists"

# ── Timing (milliseconds for Playwright) ─────────────────────────────────────
PAGE_LOAD_TIMEOUT_MS = 60_000
ELEMENT_TIMEOUT_MS   = 20_000

# ── Timing (seconds for Python sleeps) ───────────────────────────────────────
VERIFY_WAIT_SECS     = 10       # how long to wait for a result after Verify
RETRY_DELAY_SECS     = 3
INTER_RECORD_DELAY_SECS = 1     # brief pause between records
MAX_RETRIES          = 2        # whole-record retries (page reload in between)
MAX_CAPTCHA_RETRIES  = 5        # captcha attempts per record try

# ── Registration-number column auto-detection ────────────────────────────────
REG_COLUMN_ALIASES = [
    "registration number", "registration no", "registration no.",
    "registration #", "registration", "reg no", "reg no.", "reg number",
    "reg_no", "regno", "ntn", "cnic", "ntn/cnic", "cnic/ntn",
    "ntn / cnic", "cnic / ntn", "ntn no", "ntn no.", "ntn number",
    "cnic no", "cnic no.", "cnic number",
]

# ── Output ───────────────────────────────────────────────────────────────────
PDF_SUBFOLDER = "ATL_PDFs"


def default_output_dir(excel_path: str = "") -> pathlib.Path:
    """ATL_PDFs next to the chosen Excel file (or in Downloads if none yet)."""
    if excel_path:
        return pathlib.Path(excel_path).resolve().parent / PDF_SUBFOLDER
    return pathlib.Path.home() / "Downloads" / PDF_SUBFOLDER


# ── Record result values ─────────────────────────────────────────────────────
RES_PENDING    = "Pending"
RES_PROCESSING = "Processing"
RES_DOWNLOADED = "Downloaded"
RES_EXISTING   = "Already downloaded"
RES_INVALID    = "Invalid number"
RES_FAILED     = "Failed"            # always followed by " - <reason>"

DONE_RESULTS = (RES_DOWNLOADED, RES_EXISTING)

# ── GUI Appearance ────────────────────────────────────────────────────────────
APP_NAME    = "ATL Checker"
APP_VERSION = "4.1"

# Colour palette. Every entry is (light, dark).
# The dark values are identical to ATL Checker v3.1 (dark navy + cyan accent);
# the light values are the same navy / cyan family on a pale background, with
# text and status colours darkened so they stay readable.
C = {
    # surfaces
    "bg800": ("#EEF3F8", "#132238"),      # window
    "side":  ("#E1E9F2", "#0E1A2E"),      # sidebar
    "bg700": ("#FFFFFF", "#182B45"),      # cards
    "bg900": ("#F7FAFC", "#0E1A2E"),      # inputs, log, table
    "bg600": ("#DCE6F0", "#1E3452"),      # table heading
    "bg500": ("#C9D7E6", "#243D60"),
    "div":   ("#CBD5E1", "#243D60"),      # card borders
    # cyan accent
    "a600": ("#BDE8F3", "#006B84"),       # selection background
    "a500": ("#008EAD", "#008EAD"),
    "a400": ("#0096B8", "#00B4D8"),       # progress bar, checkbox
    "a300": ("#00758F", "#22D3EE"),       # accent text / primary button
    "a200": ("#005F75", "#67E8F9"),
    "acc":  ("#00758F", "#22D3EE"),
    "on_acc": ("#FFFFFF", "#080F1C"),     # text on an accent-filled button
    "sel_fg": ("#0F172A", "#FFFFFF"),     # text on a selected table row
    "seg_sel": ("#9ADBEA", "#008EAD"),    # selected Light/Dark switch segment
    "seg_hover": ("#7CCFE2", "#00A0C4"),
    # slate
    "sl700": ("#DCE6F0", "#1C2E44"),      # hover background
    "sl600": ("#B6C4D4", "#2A4060"),      # input borders, scrollbars
    "sl500": ("#94A3B8", "#3D5A7A"),
    "sl400": ("#8496AB", "#5A7A9A"),      # disabled text
    "sl300": ("#64748B", "#7AA0BC"),
    "sl200": ("#475569", "#A8C4D8"),      # secondary text
    "sl100": ("#334155", "#D8EAF4"),
    # text + status
    "white": ("#0F172A", "#FFFFFF"),      # primary text
    "ok":    ("#0B8A5F", "#2EE8A8"),
    "warn":  ("#B45309", "#FBBC41"),
    "err":   ("#D12E43", "#F95C6B"),
    "warn_hover": ("#FDF3DC", "#3A3320"),
    "err_hover":  ("#FDE7EA", "#3A2030"),
    "log_info": ("#1E4E6B", "#90D5F0"), "log_ok": ("#0B8A5F", "#2EE8A8"),
    "log_warn": ("#B45309", "#FBBC41"), "log_err": ("#D12E43", "#F95C6B"),
}

# The header bar keeps the v3.1 navy in both modes.
HDR = {"bg": "#080F1C", "title": "#FFFFFF", "sub": "#A8C4D8", "count": "#22D3EE"}

# Segoe UI / Consolas are hinted for Windows ClearType, so they stay sharp
# at small sizes (customtkinter's default Roboto does not).
FONT_UI   = "Segoe UI"
FONT_MONO = "Consolas"
