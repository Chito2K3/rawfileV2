"""Format auto-detection and cleaning pipeline for raw HIS Excel and CSV exports."""

from __future__ import annotations

import csv
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
import xlrd

from .excel_styler import write_styled_excel
from .utils import (
    clean_str,
    coerce_int_or_float,
    format_render_date,
    is_empty_cell,
    is_valid_item_code,
    normalize_item_code,
    safe_save_csv,
    sanitize_segment,
    to_number,
)

FORMAT_LABELS = {
    "detailed_ancillary": "Detailed Consumption",
    "consumption_summary": "Summary Consumption",
    "issuance_matrix": "Issuance Matrix",
    "generic_table": "Generic Table",
}

_NON_WARD_HEADERS = {
    "ITEM CODE",
    "ITEMCODE",
    "ITEM_CODE",
    "CODE",
    "DESCRIPTION",
    "ITEM DESCRIPTION",
    "ITEM NAME",
    "ARTICLE",
    "PARTICULARS",
    "GENERIC NAME",
    "UNIT",
    "UOM",
    "UNIT OF MEASURE",
    "CATEGORY",
    "CLASS",
    "PRICE",
    "UNIT COST",
    "COST",
    "AMOUNT",
    "TOTAL",
    "TOTAL ISSUANCE",
    "TOTAL ISSUED",
    "GRAND TOTAL",
    "SUB TOTAL",
    "SUBTOTAL",
    "REMARKS",
    "BALANCE",
}

# Internal pharmacy transfer columns excluded from Issuance Matrix columns and Total Issuance
EXCLUDED_ISSUANCE_PHARMACY_HEADERS = {
    "PHARMACY",
    "PHARMACY CONSIGNMENT",
    "PHARMACY STORAGE",
}

_FOOTER_KEYWORDS = (
    "prepared by",
    "approved by",
    "checked by",
    "verified by",
    "noted by",
    "certified by",
    "received by",
    "reviewed by",
    "run date",
    "date printed",
    "printed by",
    "printed on",
    "end of report",
    "*** end",
)


class _HTMLTableParser(HTMLParser):
    """Pure-stdlib HTML table parser for disguised HTML .xls exports."""

    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._in_table = False
        self._in_row = False
        self._in_cell = False
        self._current_table: list[list[str]] = []
        self._current_row: list[str] = []
        self._cell_buf: list[str] = []
        self._colspan = 1

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t == "table":
            self._in_table = True
            self._current_table = []
        elif t == "tr" and self._in_table:
            self._in_row = True
            self._current_row = []
        elif t in {"td", "th"} and self._in_row:
            self._in_cell = True
            self._cell_buf = []
            attr_dict = dict(attrs)
            try:
                self._colspan = max(1, int(attr_dict.get("colspan") or 1))
            except ValueError:
                self._colspan = 1

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in {"td", "th"} and self._in_cell:
            text = " ".join("".join(self._cell_buf).split())
            self._current_row.append(text)
            for _ in range(self._colspan - 1):
                self._current_row.append("")
            self._in_cell = False
        elif t == "tr" and self._in_row:
            if any(c != "" for c in self._current_row):
                self._current_table.append(self._current_row)
            self._in_row = False
        elif t == "table" and self._in_table:
            if self._current_table:
                self.tables.append(self._current_table)
            self._in_table = False

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._cell_buf.append(data)


@dataclass
class CleanedArtifact:
    """Represents one generated cleaned file (.xlsx or .csv)."""

    filename: str
    path: Path
    kind: str  # "excel" | "csv" | "patient_csv"
    size_bytes: int = 0


@dataclass
class FileCleanResult:
    """Structured result of cleaning a single raw HIS file."""

    source_filename: str
    detected_format: str
    format_label: str
    row_count: int
    item_count: int
    total_quantity: float | int | None
    outputs: list[CleanedArtifact] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    summary_df: pd.DataFrame | None = None
    patient_df: pd.DataFrame | None = None


def _trim_empty_rows(rows: list[list[Any]]) -> list[list[Any]]:
    return [list(r) for r in rows if any(not is_empty_cell(c) for c in r)]


def _normalize_row_widths(rows: list[list[Any]]) -> list[list[Any]]:
    if not rows:
        return []
    max_cols = max(len(r) for r in rows)
    return [list(r) + [None] * (max_cols - len(r)) for r in rows]


def _load_csv_rows(file_path: Path) -> list[tuple[str, list[list[Any]], int]]:
    raw_bytes = file_path.read_bytes()
    text = None
    for enc in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = raw_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw_bytes.decode("latin1", errors="replace")

    sample = text[:4096]
    delimiter = ","
    tab_count = sample.count("\t")
    comma_count = sample.count(",")
    semi_count = sample.count(";")
    if tab_count > comma_count and tab_count >= semi_count:
        delimiter = "\t"
    elif semi_count > comma_count and semi_count > tab_count:
        delimiter = ";"

    reader = csv.reader(text.splitlines(), delimiter=delimiter)
    rows = _trim_empty_rows([list(row) for row in reader])
    if not rows:
        return []
    return [(file_path.stem, _normalize_row_widths(rows), 0)]


def _load_html_rows(text: str, default_name: str) -> list[tuple[str, list[list[Any]], int]]:
    parser = _HTMLTableParser()
    parser.feed(text)
    parser.close()
    if not parser.tables:
        return []
    # Pick the largest table if multiple layout tables exist, or combine if needed
    largest = max(parser.tables, key=lambda t: sum(len(r) for r in t))
    rows = _trim_empty_rows(largest)
    if not rows:
        return []
    return [(default_name, _normalize_row_widths(rows), 0)]


def _load_xml_spreadsheet_rows(text: str, default_name: str) -> list[tuple[str, list[list[Any]], int]]:
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        return []

    sheets: list[tuple[str, list[list[Any]], int]] = []
    for elem in root.iter():
        if elem.tag.endswith("Worksheet"):
            sheet_name = elem.attrib.get("{urn:schemas-microsoft-com:office:spreadsheet}Name") or elem.attrib.get("Name") or default_name
            rows: list[list[Any]] = []
            for row_el in elem.iter():
                if row_el.tag.endswith("Row"):
                    row_vals: list[Any] = []
                    for cell_el in row_el:
                        if not cell_el.tag.endswith("Cell"):
                            continue
                        idx_attr = cell_el.attrib.get("{urn:schemas-microsoft-com:office:spreadsheet}Index") or cell_el.attrib.get("Index")
                        if idx_attr:
                            try:
                                target_idx = int(idx_attr) - 1
                                while len(row_vals) < target_idx:
                                    row_vals.append(None)
                            except ValueError:
                                pass
                        val = ""
                        for child in cell_el:
                            if child.tag.endswith("Data") and child.text is not None:
                                val = child.text
                                break
                        row_vals.append(val)
                    rows.append(row_vals)
            cleaned = _trim_empty_rows(rows)
            if cleaned:
                sheets.append((sheet_name, _normalize_row_widths(cleaned), 0))
    return sheets


def load_raw_sheets(file_path: Path) -> list[tuple[str, list[list[Any]], int]]:
    """Load and unmerge raw rows from .xlsx, .xlsm, .xls (including HTML/XML/CSV fallback), or .csv."""
    ext = file_path.suffix.lower()

    if ext in {".xlsx", ".xlsm"}:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheets: list[tuple[str, list[list[Any]], int]] = []
        try:
            for ws in wb.worksheets:
                for merged_range in list(ws.merged_cells.ranges):
                    ws.unmerge_cells(str(merged_range))
                raw_rows = [[cell.value for cell in row] for row in ws.iter_rows()]
                cleaned = _trim_empty_rows(raw_rows)
                if cleaned:
                    sheets.append((ws.title, _normalize_row_widths(cleaned), 0))
        finally:
            wb.close()
        return sheets

    if ext == ".xls":
        try:
            wb = xlrd.open_workbook(file_path, formatting_info=False)
            sheets = []
            for sh in wb.sheets():
                raw_rows: list[list[Any]] = []
                for r in range(sh.nrows):
                    row_vals: list[Any] = []
                    for c in range(sh.ncols):
                        cell = sh.cell(r, c)
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            try:
                                row_vals.append(xlrd.xldate_as_datetime(cell.value, wb.datemode))
                            except Exception:
                                row_vals.append(cell.value)
                        else:
                            row_vals.append(cell.value)
                    raw_rows.append(row_vals)
                cleaned = _trim_empty_rows(raw_rows)
                if cleaned:
                    sheets.append((sh.name, _normalize_row_widths(cleaned), wb.datemode))
            return sheets
        except Exception:
            # Disguised HTML, SpreadsheetML XML, or CSV/TSV saved with .xls extension
            raw_bytes = file_path.read_bytes()
            head = raw_bytes[:4096].decode("latin1", errors="replace").lower()
            if any(tag in head for tag in ("<html", "<table", "<!doctype", "<tr")):
                text = raw_bytes.decode("utf-8-sig", errors="replace")
                html_sheets = _load_html_rows(text, file_path.stem)
                if html_sheets:
                    return html_sheets
            if "<?xml" in head and ("<workbook" in head or "<worksheet" in head):
                text = raw_bytes.decode("utf-8-sig", errors="replace")
                xml_sheets = _load_xml_spreadsheet_rows(text, file_path.stem)
                if xml_sheets:
                    return xml_sheets
            return _load_csv_rows(file_path)

    # .csv / .tsv / fallback
    return _load_csv_rows(file_path)


def _norm_header_token(val: Any) -> str:
    return re.sub(r"[\s_]+", " ", clean_str(val)).strip().upper()


_CONSUMPTION_HEADER_RE = re.compile(
    r"\b(?:RENDER\w*|RETURN\w*|CONSUM\w*|CONSIGN\w*|QTY|QUANTITY|UNIT\s+COST|UNIT\s+PRICE|BEGINNING|ENDING|PATIENT\s+NAME|SLIP\w*|DELIVERED)\b",
    re.IGNORECASE,
)


def _find_issuance_header_row(rows: list[list[Any]], strict: bool = True) -> int | None:
    for idx, row in enumerate(rows[:40]):
        normed = [_norm_header_token(c) for c in row]
        if any(tok in {"ITEM CODE", "ITEMCODE", "ITEM_CODE"} for tok in normed):
            non_empty = [tok for tok in normed if tok]
            if len(non_empty) < 3:
                continue
            if strict and any(_CONSUMPTION_HEADER_RE.search(tok) for tok in non_empty):
                # Allow matrices with many ward columns even if one column says e.g. "TOTAL QTY"
                non_cons_wards = [
                    tok
                    for tok in non_empty
                    if tok not in _NON_WARD_HEADERS and not _CONSUMPTION_HEADER_RE.search(tok)
                ]
                has_strong_cons = any(
                    re.search(r"\b(?:RENDER\w*|RETURN\w*|CONSIGN\w*|CONSUM\w*|PATIENT\s+NAME|SLIP\w*)\b", tok, re.IGNORECASE)
                    for tok in non_empty
                )
                if has_strong_cons or len(non_cons_wards) < 2:
                    continue
            ward_candidates = [
                tok for tok in non_empty if tok not in _NON_WARD_HEADERS
            ]
            if ward_candidates:
                return idx
    return None


def detect_format(rows: list[list[Any]], format_hint: str | None = None) -> str:
    """Auto-detect which of the 4 HIS formats the loaded sheet rows belong to."""
    if not rows:
        return "generic_table"

    # Format 1: Detailed Ancillary Consumption
    for row in rows[:120]:
        for cell in row:
            text = clean_str(cell).lower()
            if not text:
                continue
            if (
                "detailed report" in text
                or "item code:" in text
                or re.search(r"\bslip\s*no\b", text)
            ):
                return "detailed_ancillary"

    # Format 2: Issuance Matrix (skip when caller explicitly indicates a consumption slot)
    if format_hint == "issuance":
        if _find_issuance_header_row(rows, strict=False) is not None:
            return "issuance_matrix"
    elif format_hint != "consumption" and _find_issuance_header_row(rows, strict=True) is not None:
        return "issuance_matrix"

    # Format 3: Consumption / Consignment Summary
    item_code_rows = 0
    numeric_item_rows = 0
    for row in rows:
        first_cells = row[:4]
        matched = any(is_valid_item_code(c, strict=True) for c in first_cells)
        if matched:
            item_code_rows += 1
            if any(to_number(c) is not None for c in row[1:]):
                numeric_item_rows += 1
    if item_code_rows >= 2 and numeric_item_rows >= 2:
        return "consumption_summary"

    # Also allow 1 item row if explicit consumption headers or consumption hint exist
    if item_code_rows >= 1 and numeric_item_rows >= 1:
        if format_hint == "consumption":
            return "consumption_summary"
        for row in rows[:30]:
            joined = " ".join(_norm_header_token(c) for c in row)
            if "RENDERED" in joined or "RETURNED" in joined or "CONSIGN" in joined:
                return "consumption_summary"

    return "generic_table"


def _extract_label_value(row: list[Any], label_prefix: str) -> str | None:
    """Extract value right after a label like 'Item Code:' or inline 'Item Code: DMC001'."""
    prefix_lower = label_prefix.lower()
    for i, cell in enumerate(row):
        text = clean_str(cell)
        if not text:
            continue
        low = text.lower()
        if prefix_lower in low:
            # Check inline after colon
            if ":" in text:
                after = text.split(":", 1)[1].strip()
                if after:
                    return after
            # Otherwise look at next non-empty cell(s) before another label
            for j in range(i + 1, len(row)):
                nxt = clean_str(row[j])
                if nxt:
                    if nxt.endswith(":"):
                        return None
                    return nxt
    return None


def _map_detailed_columns(row: list[Any]) -> dict[str, int] | None:
    normed = [_norm_header_token(c) for c in row]
    mapping: dict[str, int] = {}
    for idx, tok in enumerate(normed):
        if not tok:
            continue
        if any(k in tok for k in ("SLIP", "CHARGE NO", "REF NO", "REF #", "CS NO", "DOC NO", "TRANS NO")) and "slip" not in mapping:
            mapping["slip"] = idx
        elif "PATIENT" in tok and "patient" not in mapping:
            mapping["patient"] = idx
        elif (
            "DATE" in tok
            and not any(x in tok for x in ("RUN DATE", "DATE PRINTED", "PRINT DATE", "EXPIR"))
            and "date" not in mapping
        ):
            mapping["date"] = idx
        elif (
            (
                "RETURNED" in tok
                or "RETURN" in tok
                or re.search(r"\bRET\.?(?:\s+QTY)?\b", tok)
                or "QTY RET" in tok
            )
            and "returned" not in mapping
        ):
            mapping["returned"] = idx
        elif (
            ("RENDERED" in tok or "QTY RENDERED" in tok or "DISPENSED" in tok or "ISSUED" in tok or tok in {"QTY", "QUANTITY"})
            and "rendered" not in mapping
        ):
            mapping["rendered"] = idx
        elif (
            ("TOTAL QTY" in tok or "NET QTY" in tok or tok in {"TOTAL", "NET"})
            and not any(x in tok for x in ("COST", "AMOUNT", "PRICE"))
            and "total_qty" not in mapping
        ):
            mapping["total_qty"] = idx
        elif "AMOUNT" in tok and "amount" not in mapping:
            mapping["amount"] = idx
        elif ("ITEM CODE" in tok or tok == "CODE") and "item_code" not in mapping:
            mapping["item_code"] = idx
        elif ("ITEM NAME" in tok or "DESCRIPTION" in tok) and "description" not in mapping:
            mapping["description"] = idx

    if "slip" in mapping or ("patient" in mapping and ("rendered" in mapping or "date" in mapping or "returned" in mapping)):
        return mapping
    return None


def _get_adjacent_number(
    row: list[Any], idx: int | None, blocked: set[int]
) -> float | int | None:
    if idx is None:
        return None
    if idx < len(row):
        val = to_number(row[idx], None)
        if val is not None:
            return val
    for offset in (1, -1):
        cand = idx + offset
        if 0 <= cand < len(row) and cand not in blocked:
            val = to_number(row[cand], None)
            if val is not None:
                return val
    return None


def _get_adjacent_text(row: list[Any], idx: int | None, blocked: set[int]) -> str:
    if idx is None:
        return ""
    if idx < len(row):
        txt = clean_str(row[idx])
        if txt:
            return txt
    for offset in (1, -1):
        cand = idx + offset
        if 0 <= cand < len(row) and cand not in blocked:
            txt = clean_str(row[cand])
            if txt:
                return txt
    return ""


def _clean_detailed_ancillary(
    rows: list[list[Any]], datemode: int = 0
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Clean Format 1: Detailed Ancillary Consumption -> (summary_df, patient_df)."""
    col_map: dict[str, int] | None = None
    current_code = ""
    current_desc = ""

    items_order: list[str] = []
    item_meta: dict[str, str] = {}
    item_patient_totals: dict[str, list[float]] = {}  # [rendered, returned]
    item_subtotal_rows: dict[str, list[float]] = {}
    patient_records: list[dict[str, Any]] = []

    def ensure_item(code: str, desc: str) -> None:
        if code not in item_meta:
            items_order.append(code)
            item_meta[code] = desc
            item_patient_totals[code] = [0.0, 0.0]
        elif desc and not item_meta[code]:
            item_meta[code] = desc

    for row in rows:
        joined_low = " ".join(clean_str(c).lower() for c in row if clean_str(c))
        if not joined_low:
            continue

        # Check for Item Code / Item Name block headers
        found_code = _extract_label_value(row, "item code:")
        found_name = _extract_label_value(row, "item name:") or _extract_label_value(
            row, "description:"
        )
        if found_code:
            current_code = normalize_item_code(found_code)
            if found_name:
                current_desc = found_name
            elif current_code not in item_meta:
                current_desc = ""
            ensure_item(current_code, current_desc)
            continue

        if found_name and current_code:
            current_desc = found_name
            ensure_item(current_code, current_desc)
            continue

        # Check for column header row
        maybe_map = _map_detailed_columns(row)
        if maybe_map is not None:
            col_map = maybe_map
            continue

        # Check for SUB TOTAL row
        if "sub total" in joined_low or "subtotal" in joined_low:
            if current_code and col_map:
                rend_idx = col_map.get("rendered")
                ret_idx = col_map.get("returned")
                blocked_num = {
                    v for k, v in col_map.items() if k not in {"rendered", "returned"}
                }
                rend_val = _get_adjacent_number(
                    row, rend_idx, blocked_num | ({ret_idx} if ret_idx is not None else set())
                )
                ret_val = _get_adjacent_number(
                    row, ret_idx, blocked_num | ({rend_idx} if rend_idx is not None else set())
                )
                if rend_val is None:
                    nums = [to_number(c) for c in row if to_number(c) is not None]
                    if nums:
                        rend_val = nums[0]
                item_subtotal_rows[current_code] = [
                    float(rend_val or 0),
                    abs(float(ret_val or 0)),
                ]
            continue

        # Skip report grand totals / headers / footers
        if any(k in joined_low for k in _FOOTER_KEYWORDS) or joined_low.startswith("grand total"):
            continue

        # Try parsing as patient transaction row
        if col_map is None:
            continue

        mapped_indices = set(col_map.values())
        row_code = (
            normalize_item_code(row[col_map["item_code"]])
            if "item_code" in col_map and col_map["item_code"] < len(row)
            else current_code
        )
        row_desc = (
            clean_str(row[col_map["description"]])
            if "description" in col_map and col_map["description"] < len(row)
            else current_desc
        )

        slip_idx = col_map.get("slip")
        slip_val = _get_adjacent_text(
            row, slip_idx, mapped_indices - ({slip_idx} if slip_idx is not None else set())
        )
        patient_val = (
            clean_str(row[col_map["patient"]])
            if "patient" in col_map and col_map["patient"] < len(row)
            else ""
        )
        date_idx = col_map.get("date")
        raw_date = (
            row[date_idx]
            if date_idx is not None and date_idx < len(row)
            else ""
        )
        date_val = format_render_date(raw_date, datemode=datemode)
        if not date_val and date_idx is not None:
            for offset in (1, -1):
                cand = date_idx + offset
                if 0 <= cand < len(row) and cand not in mapped_indices:
                    cand_val = format_render_date(row[cand], datemode=datemode)
                    if cand_val:
                        date_val = cand_val
                        break

        rend_idx = col_map.get("rendered")
        ret_idx = col_map.get("returned")
        rend_num = _get_adjacent_number(
            row, rend_idx, mapped_indices - ({rend_idx} if rend_idx is not None else set())
        ) or 0
        ret_num = _get_adjacent_number(
            row, ret_idx, mapped_indices - ({ret_idx} if ret_idx is not None else set())
        ) or 0

        # Must look like a genuine patient transaction row
        has_identity = bool(slip_val or patient_val or date_val)
        has_qty = (rend_num or 0) != 0 or (ret_num or 0) != 0
        if not row_code or not (has_identity and (has_qty or slip_val or patient_val)):
            continue
        if slip_val.lower() in {"sub total", "subtotal", "grand total", "total"}:
            continue
        if patient_val.lower() in {"sub total", "subtotal", "grand total", "total"}:
            continue

        ensure_item(row_code, row_desc)
        rend_f = float(rend_num or 0)
        ret_f = abs(float(ret_num or 0))
        net_f = rend_f - ret_f

        item_patient_totals[row_code][0] += rend_f
        item_patient_totals[row_code][1] += ret_f

        patient_records.append(
            {
                "Item Code": row_code,
                "Description": item_meta.get(row_code, row_desc),
                "Slip No": slip_val,
                "Patient Name": patient_val,
                "Render Date": date_val,
                "Rendered Quantity": coerce_int_or_float(rend_f),
                "Returned Quantity": coerce_int_or_float(ret_f),
                "Total Quantity": coerce_int_or_float(net_f),
            }
        )

    summary_rows: list[dict[str, Any]] = []
    for code in items_order:
        p_rend, p_ret = item_patient_totals[code]
        if p_rend == 0 and p_ret == 0 and code in item_subtotal_rows:
            p_rend, p_ret = item_subtotal_rows[code]
        total_q = p_rend - p_ret
        summary_rows.append(
            {
                "Item Code": code,
                "Description": item_meta[code],
                "Rendered Quantity": coerce_int_or_float(p_rend),
                "Returned Quantity": coerce_int_or_float(p_ret),
                "Total Quantity": coerce_int_or_float(total_q),
            }
        )

    summary_cols = [
        "Item Code",
        "Description",
        "Rendered Quantity",
        "Returned Quantity",
        "Total Quantity",
    ]
    patient_cols = [
        "Item Code",
        "Description",
        "Slip No",
        "Patient Name",
        "Render Date",
        "Rendered Quantity",
        "Returned Quantity",
        "Total Quantity",
    ]
    summary_df = pd.DataFrame(summary_rows, columns=summary_cols)
    patient_df = pd.DataFrame(patient_records, columns=patient_cols)
    return summary_df, patient_df


def _clean_issuance_matrix(rows: list[list[Any]], strict: bool = True) -> pd.DataFrame:
    """Clean Format 2: Issuance Matrix -> valid items, active non-zero wards + Total Issuance."""
    header_idx = _find_issuance_header_row(rows, strict=strict)
    if header_idx is None:
        return pd.DataFrame(columns=["Item Code", "Description", "Total Issuance"])

    header_row = rows[header_idx]
    code_col = None
    desc_col = None
    ward_cols: list[tuple[int, str]] = []

    for idx, cell in enumerate(header_row):
        raw_name = clean_str(cell)
        tok = _norm_header_token(cell)
        if not tok:
            continue
        if code_col is None and tok in {"ITEM CODE", "ITEMCODE", "ITEM_CODE"}:
            code_col = idx
            continue
        if desc_col is None and tok in {
            "DESCRIPTION",
            "ITEM DESCRIPTION",
            "ITEM NAME",
            "ARTICLE",
            "PARTICULARS",
            "GENERIC NAME",
        }:
            desc_col = idx
            continue
        if tok in _NON_WARD_HEADERS or tok in EXCLUDED_ISSUANCE_PHARMACY_HEADERS:
            continue
        ward_cols.append((idx, raw_name))

    if code_col is None:
        return pd.DataFrame(columns=["Item Code", "Description", "Total Issuance"])

    # Aggregate by Item Code in case duplicates occur
    items_order: list[str] = []
    item_desc: dict[str, str] = {}
    item_wards: dict[str, dict[str, float]] = {}

    for row in rows[header_idx + 1 :]:
        raw_code = row[code_col] if code_col < len(row) else ""
        if not is_valid_item_code(raw_code):
            continue
        code = normalize_item_code(raw_code)
        desc = (
            clean_str(row[desc_col])
            if desc_col is not None and desc_col < len(row)
            else ""
        )
        if code not in item_wards:
            items_order.append(code)
            item_desc[code] = desc
            item_wards[code] = {w_name: 0.0 for _, w_name in ward_cols}
        elif desc and not item_desc[code]:
            item_desc[code] = desc

        for col_idx, w_name in ward_cols:
            if col_idx < len(row):
                num = to_number(row[col_idx], 0)
                if num is not None:
                    item_wards[code][w_name] += float(num)

    # Keep only active non-zero ward columns
    active_wards: list[str] = []
    for _, w_name in ward_cols:
        if w_name in active_wards:
            continue
        col_sum = sum(item_wards[code][w_name] for code in items_order)
        if col_sum != 0:
            active_wards.append(w_name)

    records: list[dict[str, Any]] = []
    for code in items_order:
        row_total = sum(item_wards[code][w_name] for w_name in active_wards)
        if row_total == 0:
            continue
        rec: dict[str, Any] = {
            "Item Code": code,
            "Description": item_desc[code],
        }
        for w_name in active_wards:
            val = item_wards[code][w_name]
            rec[w_name] = coerce_int_or_float(val)
        rec["Total Issuance"] = coerce_int_or_float(row_total)
        records.append(rec)

    cols = ["Item Code", "Description", *active_wards, "Total Issuance"]
    return pd.DataFrame(records, columns=cols)


def _clean_consumption_summary(rows: list[list[Any]]) -> pd.DataFrame:
    """Clean Format 3: Consumption / Consignment Summary."""
    first_item_idx = None
    code_col_counts: dict[int, int] = {}

    for r_idx, row in enumerate(rows):
        for c_idx in range(min(4, len(row))):
            if is_valid_item_code(row[c_idx], strict=True):
                if first_item_idx is None:
                    first_item_idx = r_idx
                code_col_counts[c_idx] = code_col_counts.get(c_idx, 0) + 1

    if first_item_idx is None or not code_col_counts:
        return pd.DataFrame(
            columns=[
                "Item Code",
                "Description",
                "Rendered Quantity",
                "Returned Quantity",
                "Total Quantity",
            ]
        )

    code_col = max(code_col_counts, key=lambda k: code_col_counts[k])

    # Inspect candidate header rows just above first_item_idx
    header_row: list[Any] | None = None
    for idx in range(max(0, first_item_idx - 8), first_item_idx):
        non_empty = [clean_str(c) for c in rows[idx] if clean_str(c)]
        if len(non_empty) >= 3:
            header_row = rows[idx]

    desc_col: int | None = None
    rendered_col: int | None = None
    returned_col: int | None = None
    total_col: int | None = None

    if header_row is not None:
        # Prefer explicit Description headers before generic "ITEM"
        for idx, cell in enumerate(header_row):
            if idx == code_col:
                continue
            tok = _norm_header_token(cell)
            if any(k in tok for k in ("DESCRIPTION", "ITEM NAME", "PARTICULAR", "ARTICLE")):
                desc_col = idx
                break

        for idx, cell in enumerate(header_row):
            if idx == code_col:
                continue
            tok = _norm_header_token(cell)
            if not tok:
                continue
            if desc_col is None and "ITEM" in tok and "UNIT" not in tok:
                desc_col = idx
                continue
            if any(x in tok for x in ("COST", "PRICE", "AMOUNT", "UNIT", "UOM", "BEGINNING", "ENDING", "BALANCE", "STOCK")):
                continue
            if rendered_col is None and any(
                k in tok for k in ("RENDER", "ISSUED", "SOLD", "DISPENS")
            ):
                rendered_col = idx
            elif returned_col is None and (
                "RETURN" in tok or re.search(r"\bRET\.?(?:\s+QTY)?\b", tok)
            ):
                returned_col = idx
            elif total_col is None and any(
                k in tok for k in ("TOTAL QTY", "NET QTY", "CONSUMED", "CONSUMPTION", "TOTAL", "QTY", "QUANTITY")
            ):
                total_col = idx

    # Infer description and numeric columns from item rows if headers were merged/ambiguous
    item_rows = [
        r
        for r in rows[first_item_idx:]
        if code_col < len(r) and is_valid_item_code(r[code_col], strict=True)
    ]

    # Snap header-mapped numeric columns if merged header cell shifted data by 1 column
    def _snap_numeric_col(col: int | None, blocked: set[int]) -> int | None:
        if col is None or not item_rows:
            return col
        hits = sum(1 for r in item_rows if col < len(r) and to_number(r[col]) is not None)
        if hits > 0:
            return col
        for offset in (1, -1):
            cand = col + offset
            if cand >= 0 and cand not in blocked:
                cand_hits = sum(
                    1 for r in item_rows if cand < len(r) and to_number(r[cand]) is not None
                )
                if cand_hits > 0:
                    return cand
        return col

    base_blocked = {c for c in (code_col, desc_col) if c is not None}
    rendered_col = _snap_numeric_col(
        rendered_col, base_blocked | {c for c in (returned_col, total_col) if c is not None}
    )
    returned_col = _snap_numeric_col(
        returned_col, base_blocked | {c for c in (rendered_col, total_col) if c is not None}
    )
    total_col = _snap_numeric_col(
        total_col, base_blocked | {c for c in (rendered_col, returned_col) if c is not None}
    )

    if desc_col is None and item_rows:
        num_cols = max(len(r) for r in item_rows)
        best_text_col = None
        best_text_score = 0
        for c_idx in range(code_col + 1, num_cols):
            text_hits = sum(
                1
                for r in item_rows
                if c_idx < len(r)
                and clean_str(r[c_idx])
                and to_number(r[c_idx]) is None
            )
            if text_hits > best_text_score:
                best_text_score = text_hits
                best_text_col = c_idx
        desc_col = best_text_col

    if rendered_col is None and returned_col is None and total_col is None and item_rows:
        num_cols = max(len(r) for r in item_rows)
        numeric_candidates: list[int] = []
        for c_idx in range(code_col + 1, num_cols):
            if c_idx == desc_col:
                continue
            hits = sum(1 for r in item_rows if c_idx < len(r) and to_number(r[c_idx]) is not None)
            if hits >= max(1, len(item_rows) // 2):
                numeric_candidates.append(c_idx)
        if len(numeric_candidates) >= 2:
            rendered_col = numeric_candidates[0]
            returned_col = numeric_candidates[1]
            if len(numeric_candidates) >= 3:
                total_col = numeric_candidates[2]
        elif len(numeric_candidates) == 1:
            total_col = numeric_candidates[0]

    items_order: list[str] = []
    item_desc: dict[str, str] = {}
    item_vals: dict[str, list[float]] = {}  # [rendered, returned, total]

    for row in item_rows:
        code = normalize_item_code(row[code_col])
        desc = clean_str(row[desc_col]) if desc_col is not None and desc_col < len(row) else ""

        rend_v = (
            to_number(row[rendered_col], 0)
            if rendered_col is not None and rendered_col < len(row)
            else None
        )
        ret_v = (
            to_number(row[returned_col], 0)
            if returned_col is not None and returned_col < len(row)
            else None
        )
        tot_v = (
            to_number(row[total_col], 0)
            if total_col is not None and total_col < len(row)
            else None
        )

        if rend_v is None and ret_v is None:
            rend_f = float(tot_v or 0)
            ret_f = 0.0
            tot_f = rend_f
        else:
            rend_f = float(rend_v or 0)
            ret_f = abs(float(ret_v or 0))
            tot_f = float(tot_v) if tot_v is not None else (rend_f - ret_f)

        if code not in item_vals:
            items_order.append(code)
            item_desc[code] = desc
            item_vals[code] = [rend_f, ret_f, tot_f]
        else:
            if desc and not item_desc[code]:
                item_desc[code] = desc
            item_vals[code][0] += rend_f
            item_vals[code][1] += ret_f
            item_vals[code][2] += tot_f

    records = [
        {
            "Item Code": code,
            "Description": item_desc[code],
            "Rendered Quantity": coerce_int_or_float(item_vals[code][0]),
            "Returned Quantity": coerce_int_or_float(item_vals[code][1]),
            "Total Quantity": coerce_int_or_float(item_vals[code][2]),
        }
        for code in items_order
    ]

    return pd.DataFrame(
        records,
        columns=[
            "Item Code",
            "Description",
            "Rendered Quantity",
            "Returned Quantity",
            "Total Quantity",
        ],
    )


def _is_footer_row(row: list[Any]) -> bool:
    non_empty = [clean_str(c) for c in row if clean_str(c)]
    if not non_empty:
        return True
    joined = " ".join(non_empty).lower()
    if any(kw in joined for kw in _FOOTER_KEYWORDS):
        return True
    if re.match(r"^page\s+\d+(\s+of\s+\d+)?$", joined):
        return True
    first = non_empty[0].lower().rstrip(":")
    if first in {"total", "grand total", "sub total", "subtotal", "totals"}:
        return True
    if len(non_empty) >= 2 and non_empty[1].lower().rstrip(":") in {
        "total",
        "grand total",
        "sub total",
        "subtotal",
    } and to_number(non_empty[0]) is None:
        return True
    return False


def _looks_like_code_or_id(text: str) -> bool:
    """Keep leading-zero IDs (e.g. '00123') as text rather than stripping leading zeros."""
    stripped = text.strip()
    if len(stripped) > 1 and stripped.startswith("0") and stripped[1].isdigit():
        return True
    return False


def _clean_generic_table(rows: list[list[Any]]) -> pd.DataFrame:
    """Clean Format 4: Generic Merged HIS Table."""
    if not rows:
        return pd.DataFrame()

    best_idx = 0
    best_score = -1

    for idx, row in enumerate(rows[:35]):
        non_empty = [clean_str(c) for c in row if clean_str(c)]
        if len(non_empty) < 3:
            continue
        text_headers = [
            c
            for c in non_empty
            if to_number(c) is None
            and len(c) <= 60
            and not any(kw in c.lower() for kw in ("run date", "printed", "page "))
        ]
        if len(text_headers) >= 3:
            score = len(text_headers) * 10 - idx * 0.1
            if score > best_score:
                best_score = score
                best_idx = idx

    if best_score < 0:
        # Fallback to first row with at least 2 non-empty cells
        for idx, row in enumerate(rows[:20]):
            if sum(1 for c in row if clean_str(c)) >= 2:
                best_idx = idx
                break

    header_row = rows[best_idx]
    body_rows = [r for r in rows[best_idx + 1 :] if not _is_footer_row(r)]

    max_cols = max([len(header_row), *(len(r) for r in body_rows)], default=len(header_row))
    # Keep columns that have at least one non-empty value in header or body
    active_col_indices: list[int] = []
    for c_idx in range(max_cols):
        h_val = clean_str(header_row[c_idx]) if c_idx < len(header_row) else ""
        has_body = any(c_idx < len(r) and clean_str(r[c_idx]) for r in body_rows)
        if h_val or has_body:
            active_col_indices.append(c_idx)

    if not active_col_indices:
        return pd.DataFrame()

    raw_headers: list[str] = []
    seen_headers: dict[str, int] = {}
    for pos, c_idx in enumerate(active_col_indices, start=1):
        name = clean_str(header_row[c_idx]) if c_idx < len(header_row) else ""
        if not name:
            name = f"Column_{pos}"
        count = seen_headers.get(name, 0) + 1
        seen_headers[name] = count
        if count > 1:
            name = f"{name}_{count}"
        raw_headers.append(name)

    extracted_rows: list[list[Any]] = []
    for r in body_rows:
        extracted_rows.append(
            [r[c_idx] if c_idx < len(r) else "" for c_idx in active_col_indices]
        )

    # Normalize column types cleanly
    normalized_cols: dict[str, list[Any]] = {}
    for col_pos, col_name in enumerate(raw_headers):
        col_vals = [row[col_pos] for row in extracted_rows]
        non_empty = [v for v in col_vals if not is_empty_cell(v)]

        all_numeric = bool(non_empty) and all(
            to_number(v) is not None and not _looks_like_code_or_id(clean_str(v))
            for v in non_empty
        )
        if all_numeric:
            normalized_cols[col_name] = [
                coerce_int_or_float(to_number(v, 0) or 0) if not is_empty_cell(v) else ""
                for v in col_vals
            ]
        else:
            normalized_cols[col_name] = [
                format_render_date(v) if isinstance(v, (pd.Timestamp,)) or hasattr(v, "strftime") else clean_str(v)
                for v in col_vals
            ]

    return pd.DataFrame(normalized_cols)


def _compute_total_qty(df: pd.DataFrame, fmt: str) -> float | int | None:
    if df.empty:
        return 0
    if fmt in {"detailed_ancillary", "consumption_summary"} and "Total Quantity" in df.columns:
        total = pd.to_numeric(df["Total Quantity"], errors="coerce").fillna(0).sum()
        return coerce_int_or_float(float(total))
    if fmt == "issuance_matrix" and "Total Issuance" in df.columns:
        total = pd.to_numeric(df["Total Issuance"], errors="coerce").fillna(0).sum()
        return coerce_int_or_float(float(total))
    return None


def clean_file(
    raw_path: Path,
    output_dir: Path,
    original_filename: str | None = None,
    write_outputs: bool = True,
    date_suffix: str | None = None,
    output_base_name: str | None = None,
    format_hint: str | None = None,
) -> FileCleanResult:
    """Clean a single raw HIS file and optionally write styled .xlsx + .csv outputs to output_dir."""
    source_name = original_filename or raw_path.name
    stem = sanitize_segment(Path(source_name).stem, fallback="HIS_FILE")
    if output_base_name:
        base_name = sanitize_segment(output_base_name)
    elif date_suffix:
        base_name = f"{stem}_CLEANED_{sanitize_segment(date_suffix)}"
    else:
        base_name = f"{stem}_CLEANED"

    sheets = load_raw_sheets(raw_path)
    if not sheets:
        return FileCleanResult(
            source_filename=source_name,
            detected_format="generic_table",
            format_label=FORMAT_LABELS["generic_table"],
            row_count=0,
            item_count=0,
            total_quantity=0,
            warnings=[f"No readable rows found in '{source_name}'."],
        )

    # Pick the primary sheet (largest non-empty sheet) or handle multi-sheet workbooks cleanly
    primary_name, primary_rows, datemode = max(
        sheets, key=lambda item: sum(len(r) for r in item[1])
    )
    fmt = detect_format(primary_rows, format_hint=format_hint)
    warnings: list[str] = []
    outputs: list[CleanedArtifact] = []

    patient_df: pd.DataFrame | None = None
    if fmt == "detailed_ancillary":
        summary_df, patient_df = _clean_detailed_ancillary(primary_rows, datemode=datemode)
    elif fmt == "issuance_matrix":
        summary_df = _clean_issuance_matrix(primary_rows, strict=(format_hint != "issuance"))
    elif fmt == "consumption_summary":
        summary_df = _clean_consumption_summary(primary_rows)
    else:
        summary_df = _clean_generic_table(primary_rows)

    if write_outputs:
        output_dir.mkdir(parents=True, exist_ok=True)

        xlsx_path = output_dir / f"{base_name}.xlsx"
        csv_path = output_dir / f"{base_name}.csv"

        ok_xl, err_xl = write_styled_excel(
            summary_df, xlsx_path, sheet_name=primary_name or "Cleaned"
        )
        if ok_xl and xlsx_path.exists():
            outputs.append(
                CleanedArtifact(
                    filename=xlsx_path.name,
                    path=xlsx_path,
                    kind="excel",
                    size_bytes=xlsx_path.stat().st_size,
                )
            )
        elif err_xl:
            warnings.append(err_xl)

        ok_csv, err_csv = safe_save_csv(summary_df, csv_path)
        if ok_csv and csv_path.exists():
            outputs.append(
                CleanedArtifact(
                    filename=csv_path.name,
                    path=csv_path,
                    kind="csv",
                    size_bytes=csv_path.stat().st_size,
                )
            )
        elif err_csv:
            warnings.append(err_csv)

        if patient_df is not None:
            patient_base = f"{stem}_PATIENT_ROWS_CLEANED_{sanitize_segment(date_suffix)}" if date_suffix else f"{stem}_PATIENT_ROWS_CLEANED"
            patient_csv_path = output_dir / f"{patient_base}.csv"
            ok_p, err_p = safe_save_csv(patient_df, patient_csv_path)
            if ok_p and patient_csv_path.exists():
                outputs.append(
                    CleanedArtifact(
                        filename=patient_csv_path.name,
                        path=patient_csv_path,
                        kind="patient_csv",
                        size_bytes=patient_csv_path.stat().st_size,
                    )
                )
            elif err_p:
                warnings.append(err_p)

    row_count = int(len(summary_df))
    item_count = (
        int(summary_df["Item Code"].nunique())
        if "Item Code" in summary_df.columns and not summary_df.empty
        else row_count
    )
    total_qty = _compute_total_qty(summary_df, fmt)

    return FileCleanResult(
        source_filename=source_name,
        detected_format=fmt,
        format_label=FORMAT_LABELS[fmt],
        row_count=row_count,
        item_count=item_count,
        total_quantity=total_qty,
        outputs=outputs,
        warnings=warnings,
        summary_df=summary_df,
        patient_df=patient_df,
    )

