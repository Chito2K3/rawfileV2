"""Styled Excel output writer for HIS File Cleaner."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .utils import coerce_int_or_float, safe_save_workbook

HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="000000")
BODY_FONT = Font(name="Calibri", size=11, bold=False, color="000000")
TOTAL_FONT = Font(name="Calibri", size=11, bold=True, color="000000")
TOTAL_FILL = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")

THIN_BORDER_SIDE = Side(style="thin", color="BFBFBF")
CELL_BORDER = Border(
    left=THIN_BORDER_SIDE,
    right=THIN_BORDER_SIDE,
    top=THIN_BORDER_SIDE,
    bottom=THIN_BORDER_SIDE,
)

NUMERIC_FORMAT = "#,##0.##"


def _is_numeric_value(val: Any) -> bool:
    if isinstance(val, bool) or val is None:
        return False
    if isinstance(val, (int, float)):
        return not (isinstance(val, float) and (math.isnan(val) or math.isinf(val)))
    return False


def _clean_cell_value(val: Any) -> Any:
    if val is None:
        return ""
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return ""
        return coerce_int_or_float(val)
    if pd.isna(val):
        return ""
    return val


def write_styled_excel(
    df: pd.DataFrame,
    output_path: Path,
    sheet_name: str = "Cleaned",
    total_row_indices: set[int] | None = None,
) -> tuple[bool, str | None]:
    """Write a DataFrame as a cleanly styled, A4 print-ready Excel workbook."""
    wb = Workbook()
    ws = wb.active
    safe_title = (sheet_name or "Cleaned").strip()[:31] or "Cleaned"
    for ch in r"[]:*?\/":
        safe_title = safe_title.replace(ch, "_")
    ws.title = safe_title

    columns = [str(c) for c in df.columns]
    if not columns:
        ws.cell(row=1, column=1, value="No Data")
        return safe_save_workbook(wb, output_path)

    # Detect numeric columns by checking non-empty values
    numeric_cols: set[int] = set()
    for col_idx, col_name in enumerate(columns, start=1):
        series = df.iloc[:, col_idx - 1].dropna()
        non_empty = [v for v in series.tolist() if str(v).strip() != ""]
        if non_empty and all(_is_numeric_value(v) for v in non_empty):
            numeric_cols.add(col_idx)

    # Write header row
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = CELL_BORDER
        cell.alignment = Alignment(
            horizontal="right" if col_idx in numeric_cols else "left",
            vertical="center",
            wrap_text=True,
        )
    ws.row_dimensions[1].height = 24

    total_rows = total_row_indices or set()

    # Write body rows
    for row_offset, (_, row) in enumerate(df.iterrows()):
        excel_row = row_offset + 2
        is_total_row = row_offset in total_rows
        max_lines = 1

        for col_idx in range(1, len(columns) + 1):
            raw_val = row.iloc[col_idx - 1]
            val = _clean_cell_value(raw_val)
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.border = CELL_BORDER
            cell.font = TOTAL_FONT if is_total_row else BODY_FONT
            if is_total_row:
                cell.fill = TOTAL_FILL

            if col_idx in numeric_cols and _is_numeric_value(val):
                cell.number_format = NUMERIC_FORMAT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                text_val = str(val) if val != "" else ""
                wrap = len(text_val) > 36 or "\n" in text_val
                if wrap and text_val:
                    approx_lines = max(1, math.ceil(len(text_val) / 42))
                    max_lines = max(max_lines, approx_lines)
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=wrap,
                )

        if max_lines > 1:
            ws.row_dimensions[excel_row].height = min(16 * max_lines + 4, 84)
        else:
            ws.row_dimensions[excel_row].height = 19

    # Auto-fit column widths neatly with sensible caps so descriptions wrap cleanly
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(col_name)
        for row_idx in range(2, len(df) + 2):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None or val == "":
                continue
            if col_idx in numeric_cols and _is_numeric_value(val):
                formatted = f"{val:,.2f}" if isinstance(val, float) else f"{val:,}"
                max_len = max(max_len, len(formatted))
            else:
                line_len = max((len(line) for line in str(val).splitlines()), default=0)
                max_len = max(max_len, line_len)

        min_w = 10
        max_w = 46 if col_idx not in numeric_cols else 20
        ws.column_dimensions[col_letter].width = max(min_w, min(max_len + 3, max_w))

    # Freeze header row and configure A4 print fit
    ws.freeze_panes = "A2"
    ws.views.sheetView[0].showGridLines = True

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = (
        ws.ORIENTATION_LANDSCAPE if len(columns) > 5 else ws.ORIENTATION_PORTRAIT
    )
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    return safe_save_workbook(wb, output_path)

