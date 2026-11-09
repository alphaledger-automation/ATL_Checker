"""
ATL Checker — Excel Input Handler
=================================
Reads registration numbers (NTN / CNIC) from an .xlsx file.
Auto-detects the registration column or lets the user pick one manually.
The input workbook is only ever read — results go to a separate report.
"""

import re
from dataclasses import dataclass
from typing import List, Optional, Tuple

import openpyxl

from constants import REG_COLUMN_ALIASES, RES_PENDING, RES_INVALID


@dataclass
class Record:
    """One registration number and everything we learn about it."""
    idx: int                       # 1-based position in the run
    reg_raw: str                   # value exactly as found in Excel
    reg_no: str = ""               # cleaned digits typed into IRIS
    param_type: str = ""           # "NTN" or "CNIC"
    name: str = ""
    business_name: str = ""
    filing_status: str = ""
    checking_date: str = ""
    result: str = RES_PENDING
    pdf_file: str = ""
    timestamp: str = ""


def _open_sheet(filepath: str):
    ext = filepath.rsplit(".", 1)[-1].lower()
    if ext not in ("xlsx", "xlsm"):
        raise ValueError(
            f"Unsupported file format: .{ext}\n"
            "Please save the file as .xlsx and try again."
        )
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    return wb, wb.active


def _header_row(ws) -> List[str]:
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return ["" if c is None else str(c).strip() for c in row]
    return []


def get_column_names(filepath: str) -> List[str]:
    """Return every non-blank column header from the active sheet."""
    wb, ws = _open_sheet(filepath)
    try:
        return [h for h in _header_row(ws) if h]
    finally:
        wb.close()


def detect_reg_column(columns: List[str]) -> Optional[str]:
    """Match headers against known aliases. Returns the header, or None."""
    for col in columns:
        if col.strip().lower() in REG_COLUMN_ALIASES:
            return col
    return None


def classify(value) -> Tuple[str, str]:
    """
    Turn an Excel cell value into (reg_no, param_type).

      7 characters        -> NTN   (digits, or with a letter: E123456)
      8 characters        -> NTN   (trailing check digit dropped)
      5-6 digits          -> NTN   (Excel dropped the leading zeros)
      13 digits           -> CNIC
      anything else       -> ("", "")  = invalid
    Dashes and spaces are ignored.
    """
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value).strip()
    s = re.sub(r"\.0+$", "", s)              # "1234567.0" read as text
    clean = re.sub(r"[^A-Za-z0-9]", "", s).upper()

    if clean.isdigit():
        if len(clean) == 13:
            return clean, "CNIC"
        if len(clean) == 8:
            return clean[:7], "NTN"
        if 5 <= len(clean) <= 7:
            return clean.zfill(7), "NTN"
        return "", ""
    if len(clean) in (7, 8) and sum(ch.isalpha() for ch in clean) <= 2:
        return clean[:7], "NTN"
    return "", ""


def read_records(filepath: str, column_name: str) -> List[Record]:
    """
    Read the chosen column and return one Record per UNIQUE registration
    number, in the order they first appear. Blank cells are skipped.
    """
    wb, ws = _open_sheet(filepath)
    try:
        headers = _header_row(ws)
        if column_name not in headers:
            raise ValueError(
                f"Column '{column_name}' not found.\n"
                f"Available: {', '.join(h for h in headers if h)}"
            )
        col = headers.index(column_name)

        records: List[Record] = []
        seen = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if col >= len(row) or row[col] is None:
                continue
            raw = str(row[col]).strip()
            if not raw:
                continue
            reg_no, ptype = classify(row[col])
            key = reg_no or raw
            if key in seen:
                continue
            seen.add(key)
            rec = Record(idx=len(records) + 1, reg_raw=raw,
                         reg_no=reg_no, param_type=ptype)
            if not reg_no:
                rec.result = RES_INVALID
            records.append(rec)
        return records
    finally:
        wb.close()
