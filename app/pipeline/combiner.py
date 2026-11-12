"""Batch and monthly rollup combiners for cleaned HIS files."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from .cleaner import EXCLUDED_ISSUANCE_PHARMACY_HEADERS, CleanedArtifact, FileCleanResult
from .excel_styler import write_styled_excel
from .utils import (
    clean_str,
    coerce_int_or_float,
    normalize_item_code,
    safe_save_csv,
    sanitize_segment,
    to_number,
)

DATED_DAILY_CSV_RE = re.compile(
    r"^.+_CLEANED_\d{4}-\d{2}-\d{2}\.csv$",
    re.IGNORECASE,
)

_CONSUMPTION_REQUIRED_COLS = {
    "Item Code",
    "Rendered Quantity",
    "Returned Quantity",
    "Total Quantity",
}
_ISSUANCE_EXCLUDED_COLS = {
    "Item Code",
    "Description",
    "Total Issuance",
    "TOTAL",
    "GRAND TOTAL",
}


@dataclass
class CombineSummaryItem:
    kind: str  # "consumption" | "issuance"
    item_count: int
    total_quantity: float | int
    source_files: int
    artifacts: list[CleanedArtifact] = field(default_factory=list)


@dataclass
class CombineOutcome:
    summaries: list[CombineSummaryItem] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    scanned_folders: int = 0

    @property
    def all_artifacts(self) -> list[CleanedArtifact]:
        items: list[CleanedArtifact] = []
        for s in self.summaries:
            items.extend(s.artifacts)
        return items


def _classify_cleaned_csv(df: pd.DataFrame) -> str | None:
    cols = set(str(c).strip() for c in df.columns)
    if _CONSUMPTION_REQUIRED_COLS.issubset(cols):
        return "consumption"
    if "Item Code" in cols and ("Total Issuance" in cols or "TOTAL" in cols):
        return "issuance"
    return None


def is_daily_dated_csv(path: Path) -> bool:
    """Return True if path is a daily date-stamped cleaned CSV file (excluding patient rows)."""
    name = path.name
    if "_PATIENT_ROWS_CLEANED" in name.upper():
        return False
    return bool(DATED_DAILY_CSV_RE.match(name))


def is_daily_dated_combined_csv(path: Path) -> bool:
    """Backwards-compatible alias for daily date-stamped CSV check."""
    return is_daily_dated_csv(path)


def _is_per_file_cleaned_csv(path: Path) -> bool:
    name = path.name
    if not name.endswith("_CLEANED.csv"):
        return False
    if name.endswith("_PATIENT_ROWS_CLEANED.csv"):
        return False
    if "_COMBINED_CLEANED" in name:
        return False
    return True


def combine_consumption_dfs(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """Deduplicate and sum cleaned consumption tables by Item Code."""
    item_desc: dict[str, str] = {}
    item_totals: dict[str, list[float]] = {}  # [rendered, returned, total]

    for df in dfs:
        if df.empty or "Item Code" not in df.columns:
            continue
        for _, row in df.iterrows():
            code = normalize_item_code(row.get("Item Code"))
            if not code or code in {"TOTAL", "GRAND TOTAL", "SUB TOTAL", "SUBTOTAL"}:
                continue
            desc = clean_str(row.get("Description"))
            rend = float(to_number(row.get("Rendered Quantity"), 0) or 0)
            ret = abs(float(to_number(row.get("Returned Quantity"), 0) or 0))
            tot_raw = to_number(row.get("Total Quantity"), None)
            tot = float(tot_raw) if tot_raw is not None else (rend - ret)

            if code not in item_totals:
                item_desc[code] = desc
                item_totals[code] = [rend, ret, tot]
            else:
                if desc and not item_desc[code]:
                    item_desc[code] = desc
                item_totals[code][0] += rend
                item_totals[code][1] += ret
                item_totals[code][2] += tot

    records: list[dict[str, Any]] = []
    for code in sorted(item_totals):
        rend, ret, tot = item_totals[code]
        records.append(
            {
                "Item Code": code,
                "Description": item_desc[code],
                "Rendered Quantity": coerce_int_or_float(rend),
                "Returned Quantity": coerce_int_or_float(ret),
                "Total Quantity": coerce_int_or_float(tot),
            }
        )

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


def combine_issuance_dfs(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """Unify issuance matrices across active ward columns + TOTAL."""
    ward_order: list[str] = []
    item_desc: dict[str, str] = {}
    item_wards: dict[str, dict[str, float]] = {}

    for df in dfs:
        if df.empty or "Item Code" not in df.columns:
            continue
        df_wards = [
            str(c)
            for c in df.columns
            if str(c).strip() not in _ISSUANCE_EXCLUDED_COLS
            and " ".join(str(c).split()).strip().upper() not in EXCLUDED_ISSUANCE_PHARMACY_HEADERS
        ]
        for w in df_wards:
            if w not in ward_order:
                ward_order.append(w)

        for _, row in df.iterrows():
            code = normalize_item_code(row.get("Item Code"))
            if not code or code in {"TOTAL", "GRAND TOTAL", "SUB TOTAL", "SUBTOTAL"}:
                continue
            desc = clean_str(row.get("Description"))
            if code not in item_wards:
                item_desc[code] = desc
                item_wards[code] = {}
            elif desc and not item_desc[code]:
                item_desc[code] = desc

            for w in df_wards:
                val = float(to_number(row.get(w), 0) or 0)
                item_wards[code][w] = item_wards[code].get(w, 0.0) + val

    active_wards = [
        w
        for w in ward_order
        if sum(item_wards[code].get(w, 0.0) for code in item_wards) != 0
    ]

    records: list[dict[str, Any]] = []
    for code in sorted(item_wards):
        row_total = sum(item_wards[code].get(w, 0.0) for w in active_wards)
        if row_total == 0:
            continue
        rec: dict[str, Any] = {
            "Item Code": code,
            "Description": item_desc[code],
        }
        for w in active_wards:
            val = item_wards[code].get(w, 0.0)
            rec[w] = coerce_int_or_float(val)
        rec["TOTAL"] = coerce_int_or_float(row_total)
        records.append(rec)

    return pd.DataFrame(records, columns=["Item Code", "Description", *active_wards, "TOTAL"])


def _write_combined_pair(
    df: pd.DataFrame,
    output_dir: Path,
    base_name: str,
    sheet_name: str,
    warnings: list[str],
) -> list[CleanedArtifact]:
    artifacts: list[CleanedArtifact] = []
    xlsx_path = output_dir / f"{base_name}.xlsx"
    csv_path = output_dir / f"{base_name}.csv"

    ok_xl, err_xl = write_styled_excel(df, xlsx_path, sheet_name=sheet_name)
    if ok_xl and xlsx_path.exists():
        artifacts.append(
            CleanedArtifact(
                filename=xlsx_path.name,
                path=xlsx_path,
                kind="excel",
                size_bytes=xlsx_path.stat().st_size,
            )
        )
    elif err_xl:
        warnings.append(err_xl)

    ok_csv, err_csv = safe_save_csv(df, csv_path)
    if ok_csv and csv_path.exists():
        artifacts.append(
            CleanedArtifact(
                filename=csv_path.name,
                path=csv_path,
                kind="csv",
                size_bytes=csv_path.stat().st_size,
            )
        )
    elif err_csv:
        warnings.append(err_csv)

    return artifacts


def _build_and_write_combined(
    consumption_dfs: list[pd.DataFrame],
    issuance_dfs: list[pd.DataFrame],
    output_dir: Path,
    date_str: str | None = None,
    scanned_folders: int = 1,
    warnings: list[str] | None = None,
) -> CombineOutcome:
    out_warnings = list(warnings or [])
    summaries: list[CombineSummaryItem] = []
    suffix = f"_{sanitize_segment(date_str)}" if date_str else ""

    if consumption_dfs:
        combined_cons = combine_consumption_dfs(consumption_dfs)
        artifacts = _write_combined_pair(
            combined_cons,
            output_dir,
            f"CONSUMPTION_COMBINED_CLEANED{suffix}",
            "Consumption Combined",
            out_warnings,
        )
        total_qty = (
            coerce_int_or_float(
                float(pd.to_numeric(combined_cons["Total Quantity"], errors="coerce").fillna(0).sum())
            )
            if not combined_cons.empty
            else 0
        )
        summaries.append(
            CombineSummaryItem(
                kind="consumption",
                item_count=int(len(combined_cons)),
                total_quantity=total_qty,
                source_files=len(consumption_dfs),
                artifacts=artifacts,
            )
        )

    if issuance_dfs:
        combined_iss = combine_issuance_dfs(issuance_dfs)
        artifacts = _write_combined_pair(
            combined_iss,
            output_dir,
            f"ISSUANCE_COMBINED_CLEANED{suffix}",
            "Issuance Combined",
            out_warnings,
        )
        total_qty = (
            coerce_int_or_float(
                float(pd.to_numeric(combined_iss["TOTAL"], errors="coerce").fillna(0).sum())
            )
            if not combined_iss.empty
            else 0
        )
        summaries.append(
            CombineSummaryItem(
                kind="issuance",
                item_count=int(len(combined_iss)),
                total_quantity=total_qty,
                source_files=len(issuance_dfs),
                artifacts=artifacts,
            )
        )

    return CombineOutcome(
        summaries=summaries,
        warnings=out_warnings,
        scanned_folders=scanned_folders,
    )


def combine_clean_results(
    results: list[FileCleanResult],
    output_dir: Path,
    date_str: str | None = None,
    combine_issuance: bool = False,
) -> CombineOutcome:
    """Combine in-memory Consumption (and optionally Issuance) results into date-stamped outputs."""
    consumption_dfs: list[pd.DataFrame] = []
    issuance_dfs: list[pd.DataFrame] = []

    for res in results:
        if res.summary_df is None or res.summary_df.empty:
            continue
        if res.detected_format in {"detailed_ancillary", "consumption_summary"}:
            consumption_dfs.append(res.summary_df)
        elif combine_issuance and res.detected_format == "issuance_matrix":
            issuance_dfs.append(res.summary_df)

    return _build_and_write_combined(
        consumption_dfs=consumption_dfs,
        issuance_dfs=issuance_dfs,
        output_dir=output_dir,
        date_str=date_str,
        scanned_folders=1,
    )


def _issuance_bucket(path: Path) -> str:
    low = path.name.lower()
    if "storage" in low:
        return "issuance_storage"
    if "disp" in low:
        return "issuance_dispensing"
    return "issuance"


def _combine_dated_month_csvs(
    csv_paths: list[Path],
    output_dir: Path,
    scanned_folders: int,
) -> CombineOutcome:
    """Roll up monthly Consumption and separated Issuance (Storage vs Dispensing) files."""
    consumption_dfs: list[pd.DataFrame] = []
    issuance_groups: dict[str, list[pd.DataFrame]] = {
        "issuance_storage": [],
        "issuance_dispensing": [],
        "issuance": [],
    }
    warnings: list[str] = []

    for csv_path in sorted(csv_paths):
        try:
            df = pd.read_csv(csv_path, dtype={"Item Code": str})
        except Exception as exc:
            warnings.append(f"Could not read '{csv_path.name}' for combining: {exc}")
            continue
        kind = _classify_cleaned_csv(df)
        if kind == "consumption":
            consumption_dfs.append(df)
        elif kind == "issuance":
            issuance_groups[_issuance_bucket(csv_path)].append(df)

    outcome = _build_and_write_combined(
        consumption_dfs=consumption_dfs,
        issuance_dfs=issuance_groups["issuance"],
        output_dir=output_dir,
        date_str=None,
        scanned_folders=scanned_folders,
        warnings=warnings,
    )

    bucket_meta = (
        ("issuance_storage", "ISSUANCE_STORAGE_COMBINED_CLEANED", "Storage Issuance Combined"),
        ("issuance_dispensing", "ISSUANCE_DISPENSING_COMBINED_CLEANED", "Dispensing Issuance Combined"),
    )
    for bucket_key, base_name, sheet_name in bucket_meta:
        dfs = issuance_groups[bucket_key]
        if not dfs:
            continue
        combined_iss = combine_issuance_dfs(dfs)
        artifacts = _write_combined_pair(
            combined_iss,
            output_dir,
            base_name,
            sheet_name,
            outcome.warnings,
        )
        total_qty = (
            coerce_int_or_float(
                float(pd.to_numeric(combined_iss["TOTAL"], errors="coerce").fillna(0).sum())
            )
            if not combined_iss.empty
            else 0
        )
        outcome.summaries.append(
            CombineSummaryItem(
                kind=bucket_key,
                item_count=int(len(combined_iss)),
                total_quantity=total_qty,
                source_files=len(dfs),
                artifacts=artifacts,
            )
        )

    return outcome


def _combine_csv_paths(
    csv_paths: list[Path],
    output_dir: Path,
    scanned_folders: int,
    date_str: str | None = None,
) -> CombineOutcome:
    consumption_dfs: list[pd.DataFrame] = []
    issuance_dfs: list[pd.DataFrame] = []
    warnings: list[str] = []

    for csv_path in sorted(csv_paths):
        try:
            df = pd.read_csv(csv_path, dtype={"Item Code": str})
        except Exception as exc:
            warnings.append(f"Could not read '{csv_path.name}' for combining: {exc}")
            continue
        kind = _classify_cleaned_csv(df)
        if kind == "consumption":
            consumption_dfs.append(df)
        elif kind == "issuance":
            issuance_dfs.append(df)

    return _build_and_write_combined(
        consumption_dfs=consumption_dfs,
        issuance_dfs=issuance_dfs,
        output_dir=output_dir,
        date_str=date_str,
        scanned_folders=scanned_folders,
        warnings=warnings,
    )


def combine_batch_dir(target_dir: Path, date_str: str | None = None) -> CombineOutcome:
    """Combine all per-file *_CLEANED.csv files inside a single batch/day folder."""
    if not target_dir.exists():
        return CombineOutcome()
    csv_paths = [p for p in target_dir.iterdir() if p.is_file() and _is_per_file_cleaned_csv(p)]
    return _combine_csv_paths(csv_paths, output_dir=target_dir, scanned_folders=1, date_str=date_str)


def combine_month_dir(month_dir: Path) -> CombineOutcome:
    """Combine all daily dated CSVs (in DAILY_FINAL_RESULTS/ or month_dir) into month_dir."""
    if not month_dir.exists() or not month_dir.is_dir():
        return CombineOutcome(warnings=[f"Month folder '{month_dir.name}' does not exist."])

    # 1. Preferred layout: DAILY_FINAL_RESULTS/ or flat month_dir with *_CLEANED_YYYY-MM-DD.csv
    dated_csvs: list[Path] = []
    daily_final_dir = month_dir / "DAILY_FINAL_RESULTS"
    if daily_final_dir.exists() and daily_final_dir.is_dir():
        dated_csvs.extend(
            f for f in daily_final_dir.iterdir() if f.is_file() and is_daily_dated_csv(f)
        )
    dated_csvs.extend(
        f for f in month_dir.iterdir() if f.is_file() and is_daily_dated_csv(f)
    )
    dated_csvs = sorted(dated_csvs)

    if dated_csvs:
        unique_dates = {f.stem.rsplit("_", 1)[-1] for f in dated_csvs}
        return _combine_dated_month_csvs(
            dated_csvs,
            output_dir=month_dir,
            scanned_folders=len(unique_dates),
        )

    # 2. Backwards compatibility for legacy daily subfolders or flat per-file *_CLEANED.csv
    daily_dirs = sorted(
        p for p in month_dir.iterdir() if p.is_dir() and p.name not in {"OTHER_TABLES", "DAILY_FINAL_RESULTS"}
    )
    csv_paths: list[Path] = []
    for d in daily_dirs:
        for f in d.iterdir():
            if f.is_file() and _is_per_file_cleaned_csv(f):
                csv_paths.append(f)

    if not daily_dirs:
        csv_paths = [
            f for f in month_dir.iterdir() if f.is_file() and _is_per_file_cleaned_csv(f)
        ]

    if not csv_paths:
        return CombineOutcome(
            warnings=[f"No daily cleaned consumption or issuance files found inside '{month_dir.name}'."],
            scanned_folders=len(daily_dirs),
        )

    return _combine_csv_paths(
        csv_paths,
        output_dir=month_dir,
        scanned_folders=len(daily_dirs) if daily_dirs else 1,
    )

