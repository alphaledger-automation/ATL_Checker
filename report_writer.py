"""
ATL Checker — Summary File Writer
=================================
Maintains ATL_Summary_<date>.xlsx in the PDF folder, in the same layout
as the classic input sheet:

  Length | Registration Number | file name | Status | Name |
  Filing Status | Filing Status Checking Date | ATL_PDFs Status

The file is created when a run starts (every row "Not-Downloaded") and
re-saved after every record, so it is always up to date.
"""

import pathlib
from typing import List

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from constants import DONE_RESULTS, RES_PENDING, RES_PROCESSING, RES_FAILED

# ── Styles ────────────────────────────────────────────────────────────────────
_HEADER_FILL = PatternFill("solid", fgColor="1A5530")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center")

_HEADERS = ["Length", "Registration Number", "file name", "Status", "Name",
            "Filing Status", "Filing Status Checking Date", "ATL_PDFs Status"]
_COL_WIDTHS = [11, 23, 28, 26, 36, 22, 28, 20]

STATUS_DOWNLOADED     = "Downloaded"
STATUS_NOT_DOWNLOADED = "Not-Downloaded"
PDF_PRESENT = "Saved in ATL_PDFs"
PDF_ABSENT  = "Not in ATL_PDFs"


def report_path(output_dir: pathlib.Path, date_for_file: str) -> pathlib.Path:
    return output_dir / f"ATL_Summary_{date_for_file}.xlsx"


def _status(result: str) -> str:
    if result in DONE_RESULTS:
        return STATUS_DOWNLOADED
    if result in (RES_PENDING, RES_PROCESSING):
        return STATUS_NOT_DOWNLOADED
    if result.startswith(RES_FAILED):
        return "Error" + result[len(RES_FAILED):]        # "Error - <reason>"
    return result                                        # "Invalid number"


def write_report(records: List, output_dir: pathlib.Path,
                 date_for_file: str) -> str:
    """Write the summary workbook and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for col_idx, header in enumerate(_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = _HEADER_FONT
        cell.fill      = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
    for col_idx, width in enumerate(_COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for row, rec in enumerate(records, start=2):
        reg = rec.reg_no or rec.reg_raw
        file_name = f"{rec.reg_no}_{date_for_file}.pdf" if rec.reg_no else ""
        on_disk = bool(file_name) and (output_dir / file_name).exists()

        ws.cell(row=row, column=1, value=f"=LEN(B{row})")
        # Registration numbers stay text: keeps leading zeros (0123456),
        # letters (E123456) and stops CNICs turning into 3.52E+12.
        reg_cell = ws.cell(row=row, column=2, value=str(reg))
        reg_cell.number_format = "@"
        ws.cell(row=row, column=3, value=file_name)
        ws.cell(row=row, column=4, value=_status(rec.result))
        ws.cell(row=row, column=5, value=rec.name or None)
        ws.cell(row=row, column=6, value=rec.filing_status or None)
        ws.cell(row=row, column=7, value=rec.checking_date or None)
        ws.cell(row=row, column=8, value=PDF_PRESENT if on_disk else PDF_ABSENT)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(len(records) + 1, 2)}"

    path = report_path(output_dir, date_for_file)
    wb.save(str(path))
    return str(path)
