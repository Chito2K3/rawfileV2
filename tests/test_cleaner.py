"""Comprehensive tests for HIS File Cleaner pipeline and HTTP server."""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.pipeline.cleaner import clean_file
from app.pipeline.combiner import combine_batch_dir, combine_month_dir
from app.pipeline.excel_styler import write_styled_excel
from app.pipeline.utils import resolve_safe_path
from app.server import CleanerRequestHandler


def _build_multipart_body(
    fields: dict[str, str], files: list[tuple[str, str, bytes]]
) -> tuple[str, bytes]:
    boundary = "----HISCleanerTestBoundary987654"
    chunks: list[bytes] = []
    for key, val in fields.items():
        chunks.append(f"--{boundary}\r\n".encode("utf-8"))
        chunks.append(
            f'Content-Disposition: form-data; name="{key}"\r\n\r\n{val}\r\n'.encode("utf-8")
        )
    for field_name, filename, content in files:
        chunks.append(f"--{boundary}\r\n".encode("utf-8"))
        chunks.append(
            (
                f'Content-Disposition: form-data; name="{field_name}"; '
                f'filename="{filename}"\r\n'
                f"Content-Type: application/octet-stream\r\n\r\n"
            ).encode("utf-8")
        )
        chunks.append(content)
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return f"multipart/form-data; boundary={boundary}", b"".join(chunks)


class TestCleanerPipeline(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="his_cleaner_test_"))
        self.raw_dir = self.tmp_dir / "raw"
        self.out_dir = self.tmp_dir / "out"
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.out_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_format1_detailed_ancillary_disguised_html_xls(self) -> None:
        """Detailed Ancillary Consumption (disguised HTML .xls) produces summary + patient CSV + styled Excel."""
        html_content = """<!DOCTYPE html>
<html><body>
<table>
  <tr><td colspan="6">CITY GENERAL HOSPITAL</td></tr>
  <tr><td colspan="6">Detailed Report - Ancillary Consumption</td></tr>
  <tr><td>SlipNo</td><td>Patient Name</td><td>Render Date</td><td>Rendered Qty</td><td>Returned Qty</td><td>Amount</td></tr>
  <tr><td>Item Code:</td><td>DMC001</td><td>Item Name:</td><td colspan="3">SYRINGE 5ML DISPOSABLE</td></tr>
  <tr><td>CS-1001</td><td>Dela Cruz, Juan</td><td>11/11/2026 08:30 AM</td><td>10</td><td>2</td><td>120.00</td></tr>
  <tr><td>CS-1002</td><td>Santos, Maria</td><td>11/11/2026 09:15 AM</td><td>5</td><td>0</td><td>75.00</td></tr>
  <tr><td colspan="3">SUB TOTAL</td><td>15</td><td>2</td><td>195.00</td></tr>
  <tr><td>Item Code:</td><td>DMD002</td><td>Item Name:</td><td colspan="3">PARACETAMOL 500MG TAB</td></tr>
  <tr><td>CS-1003</td><td>Reyes, Pedro</td><td>11/11/2026 10:05 AM</td><td>20</td><td>4</td><td>80.00</td></tr>
  <tr><td colspan="3">SUB TOTAL</td><td>20</td><td>4</td><td>80.00</td></tr>
</table>
</body></html>"""
        raw_file = self.raw_dir / "CONSUMP_ANCILLIARY_20261111.xls"
        raw_file.write_text(html_content, encoding="utf-8")

        res = clean_file(raw_file, self.out_dir)
        self.assertEqual(res.detected_format, "detailed_ancillary")
        self.assertEqual(res.item_count, 2)
        self.assertEqual(res.total_quantity, 29)  # (15-2) + (20-4) = 13 + 16 = 29

        cleaned_csv = self.out_dir / "CONSUMP_ANCILLIARY_20261111_CLEANED.csv"
        cleaned_xlsx = self.out_dir / "CONSUMP_ANCILLIARY_20261111_CLEANED.xlsx"
        patient_csv = self.out_dir / "CONSUMP_ANCILLIARY_20261111_PATIENT_ROWS_CLEANED.csv"

        self.assertTrue(cleaned_csv.exists())
        self.assertTrue(cleaned_xlsx.exists())
        self.assertTrue(patient_csv.exists())

        df_sum = pd.read_csv(cleaned_csv)
        self.assertEqual(
            list(df_sum.columns),
            ["Item Code", "Description", "Rendered Quantity", "Returned Quantity", "Total Quantity"],
        )
        self.assertEqual(df_sum.loc[0, "Item Code"], "DMC001")
        self.assertEqual(int(df_sum.loc[0, "Total Quantity"]), 13)
        self.assertEqual(int(df_sum.loc[1, "Total Quantity"]), 16)

        df_pat = pd.read_csv(patient_csv)
        self.assertEqual(len(df_pat), 3)
        self.assertEqual(df_pat.loc[0, "Render Date"], "2026-11-11 08:30")

        # Verify Excel styling
        wb = openpyxl.load_workbook(cleaned_xlsx)
        ws = wb.active
        self.assertEqual(ws.freeze_panes, "A2")
        self.assertEqual(ws["A1"].fill.start_color.rgb[-6:], "D9E1F2")
        self.assertTrue(ws["A1"].font.bold)
        self.assertEqual(ws["A1"].font.name, "Calibri")
        self.assertEqual(ws["C2"].number_format, "#,##0.##")
        self.assertEqual(str(ws.page_setup.paperSize), str(ws.PAPERSIZE_A4))
        self.assertEqual(ws.page_setup.fitToWidth, 1)
        wb.close()

    def test_format2_issuance_matrix_csv_and_xlsx(self) -> None:
        """Issuance Matrix keeps valid ITEM_CODE_RE rows, drops zero wards, adds Total Issuance."""
        csv_text = "\n".join(
            [
                "CITY GENERAL HOSPITAL,,,,",
                "ISSUANCE STORAGE REPORT,,,,",
                "ITEM CODE,DESCRIPTION,ER,OPD,DEAD_WARD,TOTAL",
                "CATEGORY: SYRINGES,,,,,",
                "DMC001,SYRINGE 5ML,10,5,0,15",
                "DMDON10,GLOVES STERILE,0,8,0,8",
                "INVALID_NOTE,SHOULD BE DROPPED,99,99,99,297",
                "TOTAL,,10,13,0,23",
            ]
        )
        raw_file = self.raw_dir / "issuance_storage_20261111.csv"
        raw_file.write_text(csv_text, encoding="utf-8")

        res = clean_file(raw_file, self.out_dir)
        self.assertEqual(res.detected_format, "issuance_matrix")
        self.assertEqual(res.item_count, 2)
        self.assertEqual(res.total_quantity, 23)

        df = pd.read_csv(self.out_dir / "issuance_storage_20261111_CLEANED.csv")
        self.assertEqual(list(df.columns), ["Item Code", "Description", "ER", "OPD", "Total Issuance"])
        self.assertNotIn("DEAD_WARD", df.columns)
        self.assertEqual(list(df["Item Code"]), ["DMC001", "DMDON10"])
        self.assertEqual(list(df["Total Issuance"]), [15, 8])

    def test_format3_consumption_summary_disguised_tsv_xls(self) -> None:
        """Consumption / Consignment Summary (.xls TSV fallback) detects ITEM_CODE_RE rows."""
        tsv_text = "\n".join(
            [
                "HOSPITAL PHARMACY CONSIGNMENT REPORT",
                "Run Date: 2026-11-11",
                "Code\tItem Description\tRendered\tReturned\tNet Qty",
                "AFS001\tALCOHOL 70% 500ML\t30\t2\t28",
                "PRODMR09\tGAUZE PAD 4X4\t50\t5\t45",
                "GRAND TOTAL\t\t80\t7\t73",
            ]
        )
        raw_file = self.raw_dir / "CONS_CONSUMP_20261111.xls"
        raw_file.write_text(tsv_text, encoding="utf-8")

        res = clean_file(raw_file, self.out_dir)
        self.assertEqual(res.detected_format, "consumption_summary")
        self.assertEqual(res.item_count, 2)
        self.assertEqual(res.total_quantity, 73)

        df = pd.read_csv(self.out_dir / "CONS_CONSUMP_20261111_CLEANED.csv")
        self.assertEqual(list(df["Item Code"]), ["AFS001", "PRODMR09"])
        self.assertEqual(list(df["Total Quantity"]), [28, 45])

    def test_format4_generic_merged_his_table(self) -> None:
        """Generic Merged HIS Table unmerges cells, strips banners/blank cols/footers, normalizes numbers."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Stock Balance"
        ws["A1"] = "REPUBLIC OF THE PHILIPPINES - MEDICAL CENTER"
        ws.merge_cells("A1:F1")
        ws["A2"] = "MONTHLY STOCK CARD REPORT"
        ws.merge_cells("A2:F2")
        # Row 3 blank spacer
        ws["A4"] = "Stock No"
        ws["B4"] = "Item Name"
        ws["D4"] = "Beginning"
        ws["E4"] = "Ending"
        # Column C intentionally completely empty spacer!
        ws["A5"] = "00101"
        ws["B5"] = "IV CATHETER G20"
        ws["D5"] = "1,200"
        ws["E5"] = "850"

        ws["A6"] = "00102"
        ws["B6"] = "SURGICAL MASK"
        ws["D6"] = "500"
        ws["E6"] = "300"

        # Footer rows
        ws["A8"] = "GRAND TOTAL"
        ws["D8"] = "1,700"
        ws["E8"] = "1,150"
        ws["A10"] = "Prepared by: Jane Doe"
        ws["A11"] = "Approved by: John Smith"
        ws["A12"] = "Run Date: 11/11/2026"

        raw_file = self.raw_dir / "STOCK_CARD_NOV.xlsx"
        wb.save(raw_file)
        wb.close()

        res = clean_file(raw_file, self.out_dir)
        self.assertEqual(res.detected_format, "generic_table")
        self.assertEqual(res.row_count, 2)

        df = pd.read_csv(self.out_dir / "STOCK_CARD_NOV_CLEANED.csv", dtype={"Stock No": str})
        self.assertEqual(list(df.columns), ["Stock No", "Item Name", "Beginning", "Ending"])
        self.assertEqual(list(df["Stock No"]), ["00101", "00102"])
        self.assertEqual(list(df["Beginning"]), [1200, 500])
        self.assertEqual(list(df["Ending"]), [850, 300])

    def test_batch_and_month_rollup_no_double_counting(self) -> None:
        """Batch combine + multi-day month rollup avoids double-counting on re-upload of same date."""
        month_dir = self.out_dir / "CLEANED_2026_11"
        day1_dir = month_dir / "2026-11-10"
        day2_dir = month_dir / "2026-11-11"

        cons1 = self.raw_dir / "CONS_A.csv"
        cons1.write_text(
            "Item Code,Description,Rendered Quantity,Returned Quantity,Total Quantity\n"
            "DMC001,SYRINGE 5ML,10,2,8\n"
            "DMD002,PARACETAMOL,5,0,5\n",
            encoding="utf-8",
        )
        cons2 = self.raw_dir / "CONS_B.csv"
        cons2.write_text(
            "Item Code,Description,Rendered Quantity,Returned Quantity,Total Quantity\n"
            "DMC001,SYRINGE 5ML,4,1,3\n",
            encoding="utf-8",
        )
        iss1 = self.raw_dir / "issuance_1.csv"
        iss1.write_text(
            "ITEM CODE,DESCRIPTION,ER,ICU\nDMC001,SYRINGE 5ML,12,4\n",
            encoding="utf-8",
        )
        iss2 = self.raw_dir / "issuance_2.csv"
        iss2.write_text(
            "ITEM CODE,DESCRIPTION,ER,OPD\nDMC001,SYRINGE 5ML,3,7\n",
            encoding="utf-8",
        )

        # Day 1 upload
        clean_file(cons1, day1_dir)
        clean_file(iss1, day1_dir)
        b1 = combine_batch_dir(day1_dir)
        self.assertEqual(len(b1.summaries), 2)

        # Re-upload Day 1 (same file overwrites cleanly)
        clean_file(cons1, day1_dir)
        clean_file(iss1, day1_dir)
        combine_batch_dir(day1_dir)

        # Day 2 upload
        clean_file(cons2, day2_dir)
        clean_file(iss2, day2_dir)
        combine_batch_dir(day2_dir)

        # Month rollup
        m_res = combine_month_dir(month_dir)
        self.assertEqual(m_res.scanned_folders, 2)
        self.assertEqual(len(m_res.summaries), 2)

        df_m_cons = pd.read_csv(month_dir / "CONSUMPTION_COMBINED_CLEANED.csv")
        dmc_row = df_m_cons[df_m_cons["Item Code"] == "DMC001"].iloc[0]
        self.assertEqual(int(dmc_row["Rendered Quantity"]), 14)  # 10 + 4 (not 24)
        self.assertEqual(int(dmc_row["Returned Quantity"]), 3)   # 2 + 1
        self.assertEqual(int(dmc_row["Total Quantity"]), 11)     # 8 + 3

        df_m_iss = pd.read_csv(month_dir / "ISSUANCE_COMBINED_CLEANED.csv")
        self.assertEqual(list(df_m_iss.columns), ["Item Code", "Description", "ER", "ICU", "OPD", "TOTAL"])
        self.assertEqual(int(df_m_iss.loc[0, "ER"]), 15)   # 12 + 3
        self.assertEqual(int(df_m_iss.loc[0, "ICU"]), 4)
        self.assertEqual(int(df_m_iss.loc[0, "OPD"]), 7)
        self.assertEqual(int(df_m_iss.loc[0, "TOTAL"]), 26)

    def test_locked_excel_permission_warning(self) -> None:
        """Locked Excel output raises a clear warning instead of crashing."""
        df = pd.DataFrame([{"Item Code": "DMC001", "Total Quantity": 5}])
        target = self.out_dir / "LOCKED.xlsx"
        with patch("openpyxl.Workbook.save", side_effect=PermissionError("File locked")):
            ok, msg = write_styled_excel(df, target)
        self.assertFalse(ok)
        self.assertIsNotNone(msg)
        self.assertIn("Permission denied", msg or "")

    def test_path_traversal_guard(self) -> None:
        """Path traversal outside output root is rejected."""
        with self.assertRaises(ValueError):
            resolve_safe_path(self.out_dir, "../../Windows/System32")

    def test_server_endpoints_multipart_and_downloads(self) -> None:
        """HTTP handler handles multipart upload, outputs listing, month combine, ZIP, and path guard."""

        class _MockServer:
            def __init__(self, output_root: Path) -> None:
                self.output_root = output_root.resolve()
                self.output_root.mkdir(parents=True, exist_ok=True)

        class _FakeSocket:
            def __init__(self, raw_request: bytes) -> None:
                self._rfile = BytesIO(raw_request)
                self._wfile = BytesIO()

            def makefile(self, mode: str, *args: object, **kwargs: object) -> BytesIO:
                if "r" in mode:
                    return self._rfile
                return self._wfile

            def sendall(self, data: bytes) -> None:
                self._wfile.write(data)

            def close(self) -> None:
                pass

        def dispatch(
            mock_srv: _MockServer,
            method: str,
            path: str,
            headers: dict[str, str] | None = None,
            body: bytes = b"",
        ) -> tuple[int, dict[str, str], bytes]:
            hdr_lines = [f"{method} {path} HTTP/1.1", "Host: 127.0.0.1:5055"]
            for k, v in (headers or {}).items():
                hdr_lines.append(f"{k}: {v}")
            if body and not any(k.lower() == "content-length" for k in (headers or {})):
                hdr_lines.append(f"Content-Length: {len(body)}")
            raw_req = ("\r\n".join(hdr_lines) + "\r\n\r\n").encode("latin-1") + body
            sock = _FakeSocket(raw_req)
            CleanerRequestHandler(sock, ("127.0.0.1", 5055), mock_srv)  # type: ignore[arg-type]
            raw_resp = sock._wfile.getvalue()
            head_raw, _, resp_body = raw_resp.partition(b"\r\n\r\n")
            head_lines = head_raw.decode("latin-1").split("\r\n")
            status_code = int(head_lines[0].split(" ", 2)[1])
            resp_headers: dict[str, str] = {}
            for line in head_lines[1:]:
                if ":" in line:
                    k, v = line.split(":", 1)
                    resp_headers[k.strip().lower()] = v.strip()
            return status_code, resp_headers, resp_body

        mock_srv = _MockServer(self.out_dir)

        # 1. GET /
        status, _, resp_body = dispatch(mock_srv, "GET", "/")
        self.assertEqual(status, 200)
        self.assertIn(b"HIS File Cleaner", resp_body)

        # 2. POST /api/clean with 2 files (consumption + issuance)
        cons_bytes = (
            "Code,Description,Rendered,Returned,Total\n"
            "DME001,BANDAGE,10,1,9\n"
            "DMEP01,IV SET,6,0,6\n"
        ).encode("utf-8")
        iss_bytes = (
            "ITEM CODE,DESCRIPTION,ER,WARD1\n"
            "DMEP01,IV SET,4,2\n"
        ).encode("utf-8")

        ctype, body = _build_multipart_body(
            {
                "month_folder": "CLEANED_2026_11",
                "date_str": "2026-11-11",
            },
            [
                ("files", "CONS_TEST.csv", cons_bytes),
                ("files", "issuance_test.csv", iss_bytes),
            ],
        )
        status, _, resp_body = dispatch(
            mock_srv,
            "POST",
            "/api/clean",
            headers={"Content-Type": ctype, "Content-Length": str(len(body))},
            body=body,
        )
        self.assertEqual(status, 200)
        payload = json.loads(resp_body.decode("utf-8"))
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["target_folder"], "CLEANED_2026_11/DAILY_FINAL_RESULTS")
        self.assertEqual(len(payload["files"]), 2)
        # Only Consumption is combined; Issuance stays separate
        self.assertEqual(len(payload["combined"]), 1)
        self.assertEqual(payload["combined"][0]["kind"], "consumption")

        # Verify ONLY combined consumption + separated issuance date-stamped files exist in DAILY_FINAL_RESULTS
        month_dir = self.out_dir / "CLEANED_2026_11"
        daily_dir = month_dir / "DAILY_FINAL_RESULTS"
        self.assertEqual([p.name for p in month_dir.iterdir() if p.is_dir()], ["DAILY_FINAL_RESULTS"])
        self.assertEqual(
            sorted(p.name for p in daily_dir.iterdir() if p.is_file()),
            [
                "CONSUMPTION_COMBINED_CLEANED_2026-11-11.csv",
                "CONSUMPTION_COMBINED_CLEANED_2026-11-11.xlsx",
                "issuance_test_CLEANED_2026-11-11.csv",
                "issuance_test_CLEANED_2026-11-11.xlsx",
            ],
        )

        # 3. GET /api/outputs
        status, _, resp_body = dispatch(mock_srv, "GET", "/api/outputs")
        self.assertEqual(status, 200)
        out_data = json.loads(resp_body.decode("utf-8"))
        self.assertTrue(out_data["ok"])
        self.assertEqual(out_data["months"][0]["name"], "CLEANED_2026_11")
        self.assertTrue(out_data["months"][0]["can_combine_month"])

        # 4. POST /api/combine_month
        cmb_body = json.dumps({"month_folder": "CLEANED_2026_11"}).encode("utf-8")
        status, _, resp_body = dispatch(
            mock_srv,
            "POST",
            "/api/combine_month",
            headers={"Content-Type": "application/json", "Content-Length": str(len(cmb_body))},
            body=cmb_body,
        )
        self.assertEqual(status, 200)
        cmb_data = json.loads(resp_body.decode("utf-8"))
        self.assertTrue(cmb_data["ok"])
        self.assertEqual(cmb_data["scanned_folders"], 1)

        # 5. GET /api/download_file
        status, _, resp_body = dispatch(
            mock_srv,
            "GET",
            "/api/download_file?path=CLEANED_2026_11/CONSUMPTION_COMBINED_CLEANED.csv",
        )
        self.assertEqual(status, 200)
        dl_csv = resp_body.decode("utf-8-sig")
        self.assertIn("DMEP01", dl_csv)

        # 6. GET /api/download_zip
        status, _, resp_body = dispatch(
            mock_srv,
            "GET",
            "/api/download_zip?folder=CLEANED_2026_11/DAILY_FINAL_RESULTS",
        )
        self.assertEqual(status, 200)
        with zipfile.ZipFile(BytesIO(resp_body), "r") as zf:
            names = zf.namelist()
        self.assertIn("CONSUMPTION_COMBINED_CLEANED_2026-11-11.xlsx", names)
        self.assertIn("issuance_test_CLEANED_2026-11-11.xlsx", names)

        # 7. Path traversal guard check
        status, _, _ = dispatch(
            mock_srv,
            "GET",
            "/api/download_file?path=../../secret.txt",
        )
        self.assertEqual(status, 403)

        # 8. Upload via the 4 categorized slot fields for 2026-11-12 and verify separated issuance + combined consumption
        status, _, html_body = dispatch(mock_srv, "GET", "/")
        self.assertEqual(status, 200)
        html_text = html_body.decode("utf-8")
        for expected_card in (
            "Overall Consumption",
            "Consignment Consumption",
            "Storage Issuance",
            "Dispensing Issuance",
            "cleanProgressWrap",
        ):
            self.assertIn(expected_card, html_text)

        slot_ctype, slot_body = _build_multipart_body(
            {
                "month_folder": "CLEANED_2026_11",
                "date_str": "2026-11-12",
            },
            [
                ("overall_consumption", "CONSUMP_ANCILLIARY.csv", cons_bytes),
                ("consignment_consumption", "CONS_CONSUMP.csv", cons_bytes),
                ("storage_issuance", "issuance_storage.csv", iss_bytes),
                ("dispensing_issuance", "issuance_disp.csv", iss_bytes),
            ],
        )
        status, _, resp_body = dispatch(
            mock_srv,
            "POST",
            "/api/clean",
            headers={"Content-Type": slot_ctype, "Content-Length": str(len(slot_body))},
            body=slot_body,
        )
        self.assertEqual(status, 200)
        slot_payload = json.loads(resp_body.decode("utf-8"))
        self.assertTrue(slot_payload["ok"])
        self.assertEqual(slot_payload["target_folder"], "CLEANED_2026_11/DAILY_FINAL_RESULTS")
        self.assertEqual(len(slot_payload["files"]), 4)
        self.assertEqual(
            [f["slot"] for f in slot_payload["files"]],
            [
                "overall_consumption",
                "consignment_consumption",
                "storage_issuance",
                "dispensing_issuance",
            ],
        )
        self.assertEqual(
            [f["slot_label"] for f in slot_payload["files"]],
            [
                "Overall Consumption",
                "Consignment Consumption",
                "Storage Issuance",
                "Dispensing Issuance",
            ],
        )
        # Consumption has no per-file outputs (combined only); Storage & Dispensing Issuance have separate outputs
        cons_files = [f for f in slot_payload["files"] if f["detected_format"] != "issuance_matrix"]
        iss_files = [f for f in slot_payload["files"] if f["detected_format"] == "issuance_matrix"]
        self.assertTrue(all(f["outputs"] == [] for f in cons_files))
        self.assertEqual(
            [o["filename"] for o in iss_files[0]["outputs"]],
            ["ISSUANCE_STORAGE_CLEANED_2026-11-12.xlsx", "ISSUANCE_STORAGE_CLEANED_2026-11-12.csv"],
        )
        self.assertEqual(
            [o["filename"] for o in iss_files[1]["outputs"]],
            ["ISSUANCE_DISPENSING_CLEANED_2026-11-12.xlsx", "ISSUANCE_DISPENSING_CLEANED_2026-11-12.csv"],
        )
        self.assertEqual(len(slot_payload["combined"]), 1)
        cons_comb = slot_payload["combined"][0]
        self.assertEqual(cons_comb["kind"], "consumption")
        self.assertEqual(cons_comb["source_files"], 2)
        self.assertTrue(all("2026-11-12" in out["filename"] for out in cons_comb["outputs"]))
        self.assertFalse((daily_dir / "ISSUANCE_COMBINED_CLEANED_2026-11-12.xlsx").exists())
        self.assertTrue((daily_dir / "ISSUANCE_STORAGE_CLEANED_2026-11-12.xlsx").exists())
        self.assertTrue((daily_dir / "ISSUANCE_DISPENSING_CLEANED_2026-11-12.xlsx").exists())

        # Re-upload 2026-11-12 (overwrites cleanly) then combine whole month across 2 dates (11-11 + 11-12)
        status, _, _ = dispatch(
            mock_srv,
            "POST",
            "/api/clean",
            headers={"Content-Type": slot_ctype, "Content-Length": str(len(slot_body))},
            body=slot_body,
        )
        self.assertEqual(status, 200)

        status, _, resp_body = dispatch(
            mock_srv,
            "POST",
            "/api/combine_month",
            headers={"Content-Type": "application/json", "Content-Length": str(len(cmb_body))},
            body=cmb_body,
        )
        self.assertEqual(status, 200)
        cmb_data = json.loads(resp_body.decode("utf-8"))
        self.assertEqual(cmb_data["scanned_folders"], 2)
        df_m_cons = pd.read_csv(month_dir / "CONSUMPTION_COMBINED_CLEANED.csv")
        dmep = df_m_cons[df_m_cons["Item Code"] == "DMEP01"].iloc[0]
        # 2026-11-11 had 6, 2026-11-12 had 6+6=12 => total 18 (no double-counting)
        self.assertEqual(int(dmep["Total Quantity"]), 18)
        # Monthly rollup also keeps Storage & Dispensing Issuance separated
        self.assertTrue((month_dir / "ISSUANCE_STORAGE_COMBINED_CLEANED.csv").exists())
        self.assertTrue((month_dir / "ISSUANCE_DISPENSING_COMBINED_CLEANED.csv").exists())

    def test_format1_xlsx_merged_datetime_cells(self) -> None:
        """Detailed Ancillary .xlsx with merged banners and Python datetime values cleans cleanly."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Detailed Report - Pharmacy"
        ws.merge_cells("A1:E1")
        ws["A2"] = "SlipNo"
        ws["B2"] = "Patient Name"
        ws["C2"] = "Render Date"
        ws["D2"] = "Rendered"
        ws["E2"] = "Returned"
        ws["A3"] = "Item Code:"
        ws["B3"] = "DMCSP05"
        ws["C3"] = "Item Name:"
        ws["D3"] = "CATHETER FOLEY FR16"
        ws.merge_cells("D3:E3")
        ws["A4"] = "S-900"
        ws["B4"] = "Cruz, Ana"
        ws["C4"] = datetime(2026, 11, 11, 14, 45)
        ws["D4"] = 3
        ws["E4"] = 1

        raw_file = self.raw_dir / "CONSUMP_ANCILLIARY_XLSX.xlsx"
        wb.save(raw_file)
        wb.close()

        res = clean_file(raw_file, self.out_dir)
        self.assertEqual(res.detected_format, "detailed_ancillary")
        self.assertEqual(res.total_quantity, 2)
        df_pat = pd.read_csv(self.out_dir / "CONSUMP_ANCILLIARY_XLSX_PATIENT_ROWS_CLEANED.csv")
        self.assertEqual(df_pat.loc[0, "Render Date"], "2026-11-11 14:45")
        self.assertEqual(df_pat.loc[0, "Item Code"], "DMCSP05")

    def test_consignment_summary_with_item_code_and_qty_not_misdetected_as_issuance(self) -> None:
        """Consignment summary with ITEM CODE/DESCRIPTION/UNIT/QTY/AMOUNT is not misdetected as issuance."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "CONSIGNMENT CONSUMPTION REPORT"
        ws["A2"] = "ITEM CODE"
        ws["B2"] = "DESCRIPTION"
        ws["C2"] = "UNIT"
        ws["D2"] = "QTY"
        ws["E2"] = "AMOUNT"
        ws["A3"] = "DMC000010"
        ws["B3"] = "D5NSS 1L"
        ws["C3"] = "BOTTLE"
        ws["D3"] = 120
        ws["E3"] = 9600
        ws["A4"] = "DMC000023"
        ws["B4"] = "ACETYLCYSTEINE 600MG"
        ws["C4"] = "SACHET"
        ws["D4"] = 45
        ws["E4"] = 900

        raw_file = self.raw_dir / "CONS_CONSUMP_TEST.xlsx"
        wb.save(raw_file)
        wb.close()

        res = clean_file(raw_file, self.out_dir)
        self.assertEqual(res.detected_format, "consumption_summary")
        self.assertEqual(res.item_count, 2)
        self.assertEqual(res.total_quantity, 165)

    def test_detailed_ancillary_shifted_merged_headers(self) -> None:
        """Detailed Ancillary with Crystal-style merged/shifted columns extracts slip, date, rendered, returned."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Detailed Report - Pharmacy"
        # Merged header cells where data cells sit in the right-hand column of the merge
        ws["A2"] = "Charge Slip"
        ws["B2"] = "Patient Name"
        ws["C2"] = "Rendered Date"
        ws.merge_cells("C2:D2")
        ws["E2"] = "Qty Rendered"
        ws.merge_cells("E2:F2")
        ws["G2"] = "Ret. Qty"
        ws.merge_cells("G2:H2")

        ws["A3"] = "Item Code:"
        ws["B3"] = "DMC000518"
        ws["C3"] = "Item Name:"
        ws["D3"] = "GLICLAZIDE 60MG"

        ws["A4"] = "CS-777"
        ws["B4"] = "Santos, Maria"
        # Date and numbers shifted into col D, F, H due to Crystal Reports merged layout
        ws["D4"] = "11/11/2026 09:30 AM"
        ws["F4"] = 30
        ws["H4"] = 5

        ws["A5"] = "SUB TOTAL"
        ws["F5"] = 30
        ws["H5"] = 5

        raw_file = self.raw_dir / "CONSUMP_SHIFTED.xlsx"
        wb.save(raw_file)
        wb.close()

        res = clean_file(raw_file, self.out_dir)
        self.assertEqual(res.detected_format, "detailed_ancillary")
        self.assertEqual(res.total_quantity, 25)
        df_pat = pd.read_csv(self.out_dir / "CONSUMP_SHIFTED_PATIENT_ROWS_CLEANED.csv")
        self.assertEqual(df_pat.loc[0, "Slip No"], "CS-777")
        self.assertEqual(df_pat.loc[0, "Render Date"], "2026-11-11 09:30")
        self.assertEqual(int(df_pat.loc[0, "Rendered Quantity"]), 30)
        self.assertEqual(int(df_pat.loc[0, "Returned Quantity"]), 5)
        self.assertEqual(int(df_pat.loc[0, "Total Quantity"]), 25)

    def test_issuance_matrix_with_outpatient_ward_not_blocked(self) -> None:
        """Issuance matrix containing OUTPATIENT ward column must still detect as issuance_matrix."""
        csv_text = "\n".join(
            [
                "ITEM CODE,DESCRIPTION,OUTPATIENT DEPARTMENT,ER,TOTAL",
                "DMC000010,D5NSS 1L,12,8,20",
                "DMD000023,PARACETAMOL,5,0,5",
            ]
        )
        raw_file = self.raw_dir / "issuance_disp_outpatient.csv"
        raw_file.write_text(csv_text, encoding="utf-8")

        res_auto = clean_file(raw_file, self.out_dir)
        self.assertEqual(res_auto.detected_format, "issuance_matrix")
        self.assertEqual(res_auto.total_quantity, 25)

        res_hint = clean_file(raw_file, self.out_dir, format_hint="issuance")
        self.assertEqual(res_hint.detected_format, "issuance_matrix")
        self.assertEqual(res_hint.total_quantity, 25)

    def test_issuance_matrix_with_format_hint_never_returns_empty_headers_only(self) -> None:
        """Even if a ward column contains consumption-like words, format_hint='issuance' populates rows."""
        csv_text = "\n".join(
            [
                "ITEM CODE,DESCRIPTION,OUT PATIENT,DELIVERED WARD,TOTAL",
                "DMC000010,D5NSS 1L,10,5,15",
                "DMR000001,PNSS 1L,0,7,7",
            ]
        )
        raw_file = self.raw_dir / "issuance_storage_edge.csv"
        raw_file.write_text(csv_text, encoding="utf-8")

        res = clean_file(raw_file, self.out_dir, format_hint="issuance")
        self.assertEqual(res.detected_format, "issuance_matrix")
        self.assertEqual(res.item_count, 2)
        self.assertEqual(res.total_quantity, 22)
        df = pd.read_csv(self.out_dir / "issuance_storage_edge_CLEANED.csv")
        self.assertEqual(len(df), 2)
        self.assertIn("OUT PATIENT", df.columns)

    def test_issuance_excludes_internal_pharmacy_columns_and_computations(self) -> None:
        """PHARMACY, PHARMACY CONSIGNMENT, PHARMACY STORAGE are removed before computing Total Issuance."""
        from app.pipeline.combiner import combine_issuance_dfs

        storage_csv = self.raw_dir / "issuance_storage_pharm.csv"
        storage_csv.write_text(
            "\n".join(
                [
                    "ITEM CODE,DESCRIPTION,MEDICINE WARD,PHARMACY,PHARMACY CONSIGNMENT,PHARMACY DONATION/BORROWED,Total Issuance",
                    "DMC000010,D5NSS 1L,0,5400,0,0,5400",
                    "DMC000023,ACETYLCYSTEINE,10,0,300,5,315",
                ]
            ),
            encoding="utf-8",
        )
        disp_csv = self.raw_dir / "issuance_disp_pharm.csv"
        disp_csv.write_text(
            "\n".join(
                [
                    "ITEM CODE,DESCRIPTION,PEDIA WARD,PHARMACY CONSIGNMENT,PHARMACY STORAGE,PSYCHIATRY,Total Issuance",
                    "DMC000010,D5NSS 1L,0,100,800,0,900",
                    "DMD000023,PARACETAMOL,3,10,0,7,20",
                ]
            ),
            encoding="utf-8",
        )

        res_st = clean_file(storage_csv, self.out_dir, format_hint="issuance")
        self.assertEqual(res_st.item_count, 1)  # DMC000010 dropped because remaining wards == 0
        self.assertEqual(res_st.total_quantity, 15)  # 10 (MEDICINE WARD) + 5 (DONATION/BORROWED)
        df_st = pd.read_csv(self.out_dir / "issuance_storage_pharm_CLEANED.csv")
        self.assertNotIn("PHARMACY", df_st.columns)
        self.assertNotIn("PHARMACY CONSIGNMENT", df_st.columns)
        self.assertIn("PHARMACY DONATION/BORROWED", df_st.columns)
        self.assertEqual(int(df_st.loc[0, "Total Issuance"]), 15)

        res_dp = clean_file(disp_csv, self.out_dir, format_hint="issuance")
        self.assertEqual(res_dp.item_count, 1)
        self.assertEqual(res_dp.total_quantity, 10)  # 3 (PEDIA WARD) + 7 (PSYCHIATRY)
        df_dp = pd.read_csv(self.out_dir / "issuance_disp_pharm_CLEANED.csv")
        self.assertNotIn("PHARMACY CONSIGNMENT", df_dp.columns)
        self.assertNotIn("PHARMACY STORAGE", df_dp.columns)
        self.assertEqual(int(df_dp.loc[0, "Total Issuance"]), 10)

        # Combiner also strips those columns if present
        raw_df = pd.DataFrame(
            [
                {
                    "Item Code": "DMC000023",
                    "Description": "ACETYLCYSTEINE",
                    "PHARMACY": 500,
                    "MEDICINE WARD": 12,
                    "TOTAL": 512,
                }
            ]
        )
        comb = combine_issuance_dfs([raw_df])
        self.assertNotIn("PHARMACY", comb.columns)
        self.assertEqual(int(comb.loc[0, "TOTAL"]), 12)


if __name__ == "__main__":
    unittest.main(verbosity=2)
