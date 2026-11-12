"""Shared utilities for HIS File Cleaner."""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

# Matches standard HIS item code prefixes (longest prefixes first so DMCSP/DMDON/PRODMR match cleanly).
ITEM_CODE_RE = re.compile(
    r"^(?:DMCSP|DMDON|PRODMR|AFS|DMC|DMD|DMEP|DMNP|DMR)[A-Z0-9._-]*$",
    re.IGNORECASE,
)

# Stricter check requiring at least one digit or code suffix after prefix for general text scanning.
ITEM_CODE_STRICT_RE = re.compile(
    r"^(?:DMCSP|DMDON|PRODMR|AFS|DMC|DMD|DMEP|DMNP|DMR)[A-Z0-9._-]+$",
    re.IGNORECASE,
)

_INVALID_PATH_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_NUM_CLEAN_RE = re.compile(r"^\s*([+-]?\d{1,3}(?:,\d{3})*(?:\.\d+)?|[+-]?\d+(?:\.\d+)?)\s*$")
_PAREN_NEG_RE = re.compile(r"^\s*\(\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)\s*\)\s*$")


def clean_str(value: Any) -> str:
    """Convert any cell value into a trimmed single-line/clean string."""
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null", "nat"}:
        return ""
    return text


def is_empty_cell(value: Any) -> bool:
    """Return True if cell value is empty or whitespace."""
    return clean_str(value) == ""


def is_valid_item_code(value: Any, strict: bool = False) -> bool:
    """Check whether a value matches HIS ITEM_CODE_RE."""
    code = clean_str(value).upper()
    if not code:
        return False
    pattern = ITEM_CODE_STRICT_RE if strict else ITEM_CODE_RE
    return bool(pattern.match(code))


def normalize_item_code(value: Any) -> str:
    """Normalize an item code to uppercase trimmed string."""
    return clean_str(value).upper()


def to_number(value: Any, default: float | None = None) -> float | int | None:
    """Parse numeric values from floats, ints, comma-formatted strings, or parenthesized negatives."""
    if value is None or isinstance(value, bool):
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return default
        return int(value) if value.is_integer() else round(value, 6)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "-", "--", "n/a"}:
        return default

    # Strip leading currency symbols if present
    text_no_curr = re.sub(r"^[₱$€£]\s*", "", text)

    paren_match = _PAREN_NEG_RE.match(text_no_curr)
    if paren_match:
        num = -float(paren_match.group(1).replace(",", ""))
        return int(num) if num.is_integer() else round(num, 6)

    num_match = _NUM_CLEAN_RE.match(text_no_curr)
    if num_match:
        num = float(num_match.group(1).replace(",", ""))
        return int(num) if num.is_integer() else round(num, 6)

    return default


def coerce_int_or_float(value: float | int) -> int | float:
    """Return int if value has no fractional part, otherwise float."""
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return 0
        if value.is_integer():
            return int(value)
        return round(value, 6)
    return value


def parse_date_value(value: Any, datemode: int = 0) -> datetime | None:
    """Parse datetime from Excel serial number, datetime/date, or common HIS timestamp strings."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if not math.isnan(float(value)) and 1 <= float(value) <= 100000:
            try:
                import xlrd

                return xlrd.xldate_as_datetime(float(value), datemode)
            except Exception:
                pass

    text = clean_str(value)
    if not text:
        return None

    formats = (
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y %I:%M %p",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%m-%d-%Y %I:%M %p",
        "%m-%d-%Y",
        "%b %d, %Y %I:%M %p",
        "%b %d, %Y",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def format_render_date(value: Any, datemode: int = 0) -> str:
    """Format a transaction date cleanly as YYYY-MM-DD HH:MM (or YYYY-MM-DD if midnight)."""
    dt = parse_date_value(value, datemode=datemode)
    if dt is None:
        return clean_str(value)
    if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d %H:%M")


def sanitize_segment(name: str, fallback: str = "folder") -> str:
    """Sanitize a single path segment (folder or file stem) safely for Windows."""
    cleaned = _INVALID_PATH_CHARS_RE.sub("_", (name or "").strip())
    cleaned = cleaned.strip(" .")
    if not cleaned or cleaned in {".", ".."}:
        return fallback
    return cleaned


def resolve_safe_path(base_dir: Path, relative_path: str) -> Path:
    """Resolve a user-supplied relative path under base_dir, rejecting path traversal."""
    base_resolved = base_dir.resolve()
    rel = (relative_path or "").replace("\\", "/").lstrip("/")
    target = (base_resolved / rel).resolve()
    if target != base_resolved and not target.is_relative_to(base_resolved):
        raise ValueError(f"Unsafe path outside output root: {relative_path}")
    return target


def safe_save_workbook(wb: Any, path: Path) -> tuple[bool, str | None]:
    """Save an openpyxl Workbook, returning a friendly warning on PermissionError."""
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path)
        return True, None
    except PermissionError:
        return (
            False,
            f"Permission denied saving '{path.name}' — it may be open in Excel. Close the file and re-run.",
        )
    except OSError as exc:
        return False, f"Failed saving '{path.name}': {exc}"


def safe_save_csv(df: pd.DataFrame, path: Path) -> tuple[bool, str | None]:
    """Save a DataFrame to UTF-8-SIG CSV, returning a friendly warning on PermissionError."""
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
        return True, None
    except PermissionError:
        return (
            False,
            f"Permission denied saving '{path.name}' — it may be open in another program. Close it and re-run.",
        )
    except OSError as exc:
        return False, f"Failed saving '{path.name}': {exc}"

