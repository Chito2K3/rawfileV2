"""Standalone local HTTP server for HIS File Cleaner."""

from __future__ import annotations

import argparse
import email.policy
import io
import json
import mimetypes
import os
import subprocess
import sys
import tempfile
import threading
import urllib.parse
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import datetime
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

# Support running either as `python app/server.py` or `python -m app.server`
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.pipeline.cleaner import FileCleanResult, clean_file
from app.pipeline.combiner import (
    CombineOutcome,
    combine_clean_results,
    combine_month_dir,
    is_daily_dated_csv,
)
from app.pipeline.utils import resolve_safe_path, sanitize_segment

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "index.html"
DEFAULT_OUTPUT_ROOT = PROJECT_ROOT / "OUTPUT_CLEANED"
ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".csv", ".tsv"}
MAX_UPLOAD_BYTES = 200 * 1024 * 1024  # 200 MB
BACKEND_VERSION = "2026-11-12-v7"

SLOT_LABELS = {
    "overall_consumption": "Overall Consumption",
    "consignment_consumption": "Consignment Consumption",
    "storage_issuance": "Storage Issuance",
    "dispensing_issuance": "Dispensing Issuance",
    "other_files": "Other HIS Table",
    "files": "Auto-Detected",
}


@dataclass
class UploadedFile:
    filename: str
    data: bytes
    field_name: str = "files"


def parse_multipart_form(
    content_type: str, body: bytes
) -> tuple[dict[str, str], list[UploadedFile]]:
    """Parse multipart/form-data using Python stdlib email.parser.BytesParser (3.13+ safe)."""
    if "multipart/form-data" not in (content_type or "").lower():
        raise ValueError("Expected multipart/form-data request.")

    raw_msg = (
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("latin-1")
        + body
    )
    msg = BytesParser(policy=email.policy.default).parsebytes(raw_msg)
    if not msg.is_multipart():
        raise ValueError("Malformed multipart payload.")

    fields: dict[str, str] = {}
    files: list[UploadedFile] = []

    for part in msg.iter_parts():
        field_name = part.get_param("name", header="content-disposition")
        filename = part.get_param("filename", header="content-disposition")
        payload = part.get_payload(decode=True)
        if payload is None:
            raw_payload = part.get_payload()
            payload = (
                raw_payload.encode("utf-8")
                if isinstance(raw_payload, str)
                else b""
            )

        if filename is not None:
            if filename.strip():
                files.append(
                    UploadedFile(
                        filename=Path(filename).name,
                        data=payload,
                        field_name=field_name or "files",
                    )
                )
        elif field_name:
            fields[field_name] = payload.decode("utf-8", errors="replace").strip()

    return fields, files


def _file_info_dict(file_path: Path, output_root: Path) -> dict[str, Any]:
    stat = file_path.stat()
    rel = file_path.relative_to(output_root).as_posix()
    return {
        "name": file_path.name,
        "relative_path": rel,
        "size_bytes": stat.st_size,
        "modified_at": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
    }


def list_outputs_tree(output_root: Path) -> list[dict[str, Any]]:
    """List month folders, their root combined files, and daily subfolders."""
    output_root.mkdir(parents=True, exist_ok=True)
    months: list[dict[str, Any]] = []

    month_dirs = sorted(
        (p for p in output_root.iterdir() if p.is_dir()),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    for m_dir in month_dirs:
        root_files = sorted(
            (f for f in m_dir.iterdir() if f.is_file()),
            key=lambda f: f.name,
        )
        subfolders: list[dict[str, Any]] = []
        for sub in sorted((p for p in m_dir.iterdir() if p.is_dir()), reverse=True):
            sub_files = sorted(
                (f for f in sub.iterdir() if f.is_file()),
                key=lambda f: f.name,
            )
            subfolders.append(
                {
                    "name": sub.name,
                    "relative_folder": sub.relative_to(output_root).as_posix(),
                    "files": [_file_info_dict(f, output_root) for f in sub_files],
                }
            )

        daily_final = m_dir / "DAILY_FINAL_RESULTS"
        has_dated = any(is_daily_dated_csv(f) for f in root_files) or (
            daily_final.is_dir() and any(is_daily_dated_csv(f) for f in daily_final.iterdir() if f.is_file())
        )
        months.append(
            {
                "name": m_dir.name,
                "relative_folder": m_dir.relative_to(output_root).as_posix(),
                "can_combine_month": has_dated or len(subfolders) > 0,
                "has_daily_subfolders": len(subfolders) > 0,
                "files": [_file_info_dict(f, output_root) for f in root_files],
                "subfolders": subfolders,
            }
        )

    return months


def _serialize_clean_result(
    res: FileCleanResult, output_root: Path, slot: str = "files"
) -> dict[str, Any]:
    return {
        "source_filename": res.source_filename,
        "slot": slot,
        "slot_label": SLOT_LABELS.get(slot, "Auto-Detected"),
        "detected_format": res.detected_format,
        "format_label": res.format_label,
        "row_count": res.row_count,
        "item_count": res.item_count,
        "total_quantity": res.total_quantity,
        "warnings": res.warnings,
        "outputs": [
            {
                "filename": art.filename,
                "relative_path": art.path.relative_to(output_root).as_posix(),
                "kind": art.kind,
                "size_bytes": art.size_bytes,
            }
            for art in res.outputs
        ],
    }


def _serialize_combine_outcome(outcome: CombineOutcome, output_root: Path) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for s in outcome.summaries:
        items.append(
            {
                "kind": s.kind,
                "item_count": s.item_count,
                "total_quantity": s.total_quantity,
                "source_files": s.source_files,
                "outputs": [
                    {
                        "filename": art.filename,
                        "relative_path": art.path.relative_to(output_root).as_posix(),
                        "kind": art.kind,
                        "size_bytes": art.size_bytes,
                    }
                    for art in s.artifacts
                ],
            }
        )
    return items


class CleanerHTTPServer(ThreadingHTTPServer):
    """ThreadingHTTPServer carrying configurable output_root."""

    allow_reuse_address = True
    daemon_threads = True

    def __init__(
        self,
        server_address: tuple[str, int],
        RequestHandlerClass: type[BaseHTTPRequestHandler],
        output_root: Path = DEFAULT_OUTPUT_ROOT,
    ) -> None:
        self.output_root = output_root.resolve()
        self.output_root.mkdir(parents=True, exist_ok=True)
        super().__init__(server_address, RequestHandlerClass)


class CleanerRequestHandler(BaseHTTPRequestHandler):
    """HTTP API and UI handler for HIS File Cleaner."""

    server: CleanerHTTPServer

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        # Keep console output clean
        sys.stderr.write(f"[HIS Cleaner] {self.address_string()} - {format % args}\n")

    def _send_json(self, status: int, payload: dict[str, Any]) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def _send_error_json(self, status: int, message: str) -> None:
        self._send_json(status, {"ok": False, "error": message})

    def _read_request_body(self) -> bytes:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            raise ValueError("Invalid Content-Length header.")
        if length < 0 or length > MAX_UPLOAD_BYTES:
            raise ValueError("Request payload size out of bounds.")
        if length == 0:
            return b""
        return self.rfile.read(length)

    def _read_json_or_form_params(self) -> dict[str, str]:
        content_type = self.headers.get("Content-Type", "")
        body = self._read_request_body()
        if not body:
            return {}
        if "application/json" in content_type.lower():
            data = json.loads(body.decode("utf-8"))
            if not isinstance(data, dict):
                return {}
            return {str(k): str(v) for k, v in data.items() if v is not None}
        if "multipart/form-data" in content_type.lower():
            fields, _ = parse_multipart_form(content_type, body)
            return fields
        parsed = urllib.parse.parse_qs(body.decode("utf-8", errors="replace"))
        return {k: v[0] for k, v in parsed.items() if v}

    def do_GET(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        route = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if route == "/" or route == "/index.html":
            self._handle_index()
        elif route == "/api/outputs":
            self._handle_list_outputs()
        elif route == "/api/download_file":
            rel_path = (query.get("path") or [""])[0]
            self._handle_download_file(rel_path)
        elif route == "/api/download_zip":
            rel_folder = (query.get("folder") or [""])[0]
            self._handle_download_zip(rel_folder)
        else:
            self._send_error_json(HTTPStatus.NOT_FOUND, f"Unknown endpoint: {route}")

    def do_POST(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        route = parsed.path

        if route == "/api/clean":
            self._handle_clean()
        elif route == "/api/combine_month":
            self._handle_combine_month()
        elif route == "/api/open_folder":
            self._handle_open_folder()
        else:
            self._send_error_json(HTTPStatus.NOT_FOUND, f"Unknown endpoint: {route}")

    def _handle_index(self) -> None:
        if not TEMPLATE_PATH.exists():
            self._send_error_json(HTTPStatus.INTERNAL_SERVER_ERROR, "index.html template missing.")
            return
        html = TEMPLATE_PATH.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(html)))
        self.end_headers()
        self.wfile.write(html)

    def _handle_list_outputs(self) -> None:
        tree = list_outputs_tree(self.server.output_root)
        self._send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "backend_version": BACKEND_VERSION,
                "output_root": str(self.server.output_root),
                "months": tree,
            },
        )

    def _handle_clean(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        try:
            body = self._read_request_body()
            fields, uploaded_files = parse_multipart_form(content_type, body)
        except ValueError as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return

        if not uploaded_files:
            self._send_error_json(HTTPStatus.BAD_REQUEST, "No files were uploaded.")
            return

        now = datetime.now()
        default_month = f"CLEANED_{now.strftime('%Y_%m')}"
        month_raw = fields.get("month_folder") or fields.get("folder_label") or default_month
        month_folder = sanitize_segment(month_raw, fallback=default_month)

        date_raw = fields.get("date_str") or fields.get("date") or now.strftime("%Y-%m-%d")
        date_str = sanitize_segment(date_raw, fallback=now.strftime("%Y-%m-%d"))

        try:
            daily_dir = resolve_safe_path(
                self.server.output_root, f"{month_folder}/DAILY_FINAL_RESULTS"
            )
            other_dir = resolve_safe_path(
                self.server.output_root, f"{month_folder}/OTHER_TABLES"
            )
        except ValueError as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return

        daily_dir.mkdir(parents=True, exist_ok=True)

        slot_counts: dict[str, int] = {}
        for up_file in uploaded_files:
            slot_counts[up_file.field_name] = slot_counts.get(up_file.field_name, 0) + 1

        results: list[tuple[FileCleanResult, str]] = []
        warnings: list[str] = []

        with tempfile.TemporaryDirectory() as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            for idx, up_file in enumerate(uploaded_files):
                ext = Path(up_file.filename).suffix.lower()
                if ext not in ALLOWED_EXTENSIONS:
                    warnings.append(
                        f"Skipped '{up_file.filename}': unsupported extension '{ext or '(none)'}'."
                    )
                    continue

                tmp_path = tmp_dir / f"{idx}_{sanitize_segment(Path(up_file.filename).stem)}{ext}"
                tmp_path.write_bytes(up_file.data)
                try:
                    slot_hint = None
                    if up_file.field_name in {"overall_consumption", "consignment_consumption"}:
                        slot_hint = "consumption"
                    elif up_file.field_name in {"storage_issuance", "dispensing_issuance"}:
                        slot_hint = "issuance"

                    # Clean in memory first; only combine consumption, keep issuance & generic separated
                    res = clean_file(
                        tmp_path,
                        daily_dir,
                        original_filename=up_file.filename,
                        write_outputs=False,
                        date_suffix=date_str,
                        format_hint=slot_hint,
                    )
                    if res.detected_format == "issuance_matrix":
                        out_name = None
                        if up_file.field_name == "storage_issuance" and slot_counts.get("storage_issuance") == 1:
                            out_name = f"ISSUANCE_STORAGE_CLEANED_{date_str}"
                        elif up_file.field_name == "dispensing_issuance" and slot_counts.get("dispensing_issuance") == 1:
                            out_name = f"ISSUANCE_DISPENSING_CLEANED_{date_str}"
                        res = clean_file(
                            tmp_path,
                            daily_dir,
                            original_filename=up_file.filename,
                            write_outputs=True,
                            date_suffix=date_str,
                            output_base_name=out_name,
                            format_hint=slot_hint,
                        )
                    elif res.detected_format == "generic_table":
                        res = clean_file(
                            tmp_path,
                            other_dir,
                            original_filename=up_file.filename,
                            write_outputs=True,
                            date_suffix=date_str,
                            format_hint=slot_hint,
                        )
                    results.append((res, up_file.field_name))
                    warnings.extend(res.warnings)
                except Exception as exc:
                    warnings.append(f"Failed cleaning '{up_file.filename}': {exc}")

        if not results and warnings:
            self._send_json(
                HTTPStatus.BAD_REQUEST,
                {"ok": False, "error": warnings[0], "warnings": warnings},
            )
            return

        combine_outcome = combine_clean_results(
            [r for r, _ in results],
            daily_dir,
            date_str=date_str,
            combine_issuance=False,
        )
        warnings.extend(combine_outcome.warnings)

        self._send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "month_folder": month_folder,
                "date_str": date_str,
                "target_folder": daily_dir.relative_to(self.server.output_root).as_posix(),
                "target_abs_path": str(daily_dir),
                "files": [
                    _serialize_clean_result(r, self.server.output_root, slot=slot)
                    for r, slot in results
                ],
                "combined": _serialize_combine_outcome(combine_outcome, self.server.output_root),
                "warnings": warnings,
            },
        )

    def _handle_combine_month(self) -> None:
        try:
            params = self._read_json_or_form_params()
        except Exception as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, f"Invalid request body: {exc}")
            return

        month_raw = params.get("month_folder", "").strip()
        if not month_raw:
            self._send_error_json(HTTPStatus.BAD_REQUEST, "Missing 'month_folder'.")
            return

        try:
            month_dir = resolve_safe_path(self.server.output_root, month_raw)
        except ValueError as exc:
            self._send_error_json(HTTPStatus.FORBIDDEN, str(exc))
            return

        if not month_dir.exists() or not month_dir.is_dir():
            self._send_error_json(
                HTTPStatus.NOT_FOUND, f"Month folder '{month_raw}' not found."
            )
            return

        outcome = combine_month_dir(month_dir)
        self._send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "month_folder": month_dir.name,
                "target_abs_path": str(month_dir),
                "scanned_folders": outcome.scanned_folders,
                "combined": _serialize_combine_outcome(outcome, self.server.output_root),
                "warnings": outcome.warnings,
            },
        )

    def _handle_download_file(self, rel_path: str) -> None:
        if not rel_path:
            self._send_error_json(HTTPStatus.BAD_REQUEST, "Missing 'path' parameter.")
            return
        try:
            file_path = resolve_safe_path(self.server.output_root, rel_path)
        except ValueError as exc:
            self._send_error_json(HTTPStatus.FORBIDDEN, str(exc))
            return

        if not file_path.exists() or not file_path.is_file():
            self._send_error_json(HTTPStatus.NOT_FOUND, "File not found.")
            return

        data = file_path.read_bytes()
        mime_type, _ = mimetypes.guess_type(file_path.name)
        if file_path.suffix.lower() == ".xlsx":
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif file_path.suffix.lower() == ".csv":
            mime_type = "text/csv; charset=utf-8"

        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mime_type or "application/octet-stream")
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{file_path.name}"',
        )
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _handle_download_zip(self, rel_folder: str) -> None:
        if not rel_folder:
            self._send_error_json(HTTPStatus.BAD_REQUEST, "Missing 'folder' parameter.")
            return
        try:
            folder_path = resolve_safe_path(self.server.output_root, rel_folder)
        except ValueError as exc:
            self._send_error_json(HTTPStatus.FORBIDDEN, str(exc))
            return

        if not folder_path.exists() or not folder_path.is_dir():
            self._send_error_json(HTTPStatus.NOT_FOUND, "Folder not found.")
            return

        buf = io.BytesIO()
        files_added = 0
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for item in sorted(folder_path.rglob("*")):
                if item.is_file():
                    arcname = item.relative_to(folder_path).as_posix()
                    zf.write(item, arcname=arcname)
                    files_added += 1

        if files_added == 0:
            self._send_error_json(HTTPStatus.NOT_FOUND, "Folder has no cleaned files to zip.")
            return

        zip_bytes = buf.getvalue()
        zip_name = f"{sanitize_segment(folder_path.name, fallback='cleaned')}.zip"

        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Disposition", f'attachment; filename="{zip_name}"')
        self.send_header("Content-Length", str(len(zip_bytes)))
        self.end_headers()
        self.wfile.write(zip_bytes)

    def _handle_open_folder(self) -> None:
        try:
            params = self._read_json_or_form_params()
        except Exception as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, f"Invalid request body: {exc}")
            return

        rel_folder = params.get("folder", "").strip()
        try:
            folder_path = (
                resolve_safe_path(self.server.output_root, rel_folder)
                if rel_folder
                else self.server.output_root
            )
        except ValueError as exc:
            self._send_error_json(HTTPStatus.FORBIDDEN, str(exc))
            return

        if not folder_path.exists():
            folder_path.mkdir(parents=True, exist_ok=True)

        try:
            if sys.platform.startswith("win"):
                os.startfile(str(folder_path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder_path)])
            else:
                subprocess.Popen(["xdg-open", str(folder_path)])
            self._send_json(HTTPStatus.OK, {"ok": True, "path": str(folder_path)})
        except Exception as exc:
            self._send_error_json(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                f"Could not open folder in file explorer: {exc}",
            )


def run_server(
    host: str = "127.0.0.1",
    port: int = 5055,
    open_browser: bool = False,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
) -> None:
    try:
        server = CleanerHTTPServer((host, port), CleanerRequestHandler, output_root=output_root)
    except OSError as exc:
        print(
            f"\nERROR: Could not start HIS File Cleaner on {host}:{port} ({exc}).\n"
            "An older HIS File Cleaner window is already running on this port.\n"
            "Please close the other server window first, then run launch_cleaner.bat again.",
            file=sys.stderr,
        )
        sys.exit(1)

    url = f"http://{host}:{port}"
    print(f"HIS File Cleaner ({BACKEND_VERSION}) running at {url}")
    print(f"Output directory: {server.output_root}")

    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down HIS File Cleaner...")
    finally:
        server.server_close()


def main() -> None:
    parser = argparse.ArgumentParser(description="HIS File Cleaner local web app")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host (default 127.0.0.1)")
    parser.add_argument("--port", type=int, default=5055, help="Bind port (default 5055)")
    parser.add_argument("--open", action="store_true", help="Open browser automatically")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_ROOT,
        help="Cleaned output root folder",
    )
    args = parser.parse_args()
    run_server(
        host=args.host,
        port=args.port,
        open_browser=args.open,
        output_root=args.output_dir,
    )


if __name__ == "__main__":
    main()

