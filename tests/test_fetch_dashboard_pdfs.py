import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from fetch_dashboard_pdfs import run_ingestion
from period_scopes import period_scope_filename
from reporting_periods import ReportingSchedule, build_quarter_period


class _FakeGetRequest:
    def __init__(self, payload: dict):
        self._payload = payload

    def execute(self):
        return self._payload


class _FakeMessages:
    def __init__(self, messages: dict[str, dict]):
        self._messages = messages

    def get(self, userId: str, id: str, format: str = "full"):
        return _FakeGetRequest(self._messages[id])


class _FakeUsers:
    def __init__(self, messages: dict[str, dict]):
        self._messages = messages

    def messages(self):
        return _FakeMessages(self._messages)


class _FakeService:
    def __init__(self, messages: dict[str, dict]):
        self._messages = messages

    def users(self):
        return _FakeUsers(self._messages)


class FetchDashboardPdfsTests(unittest.TestCase):
    def test_period_scope_filename(self):
        self.assertEqual(period_scope_filename("quarter_2026_Q1", "argentina"), "quarter_2026_Q1_ARG.pdf")
        self.assertEqual(period_scope_filename("quarter_2026_Q1", "holding"), "quarter_2026_Q1_HOLDING.pdf")
        self.assertEqual(period_scope_filename("quarter_2026_Q1", "combined"), "quarter_2026_Q1_ARG_HOLDING.pdf")

    def test_run_ingestion_creates_manifest_and_period_scope_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            data_dir = tmp_dir / "data"
            pdf_dir = tmp_dir / "local_data" / "inbox_pdfs"
            schedule = ReportingSchedule(
                timezone="America/Argentina/Buenos_Aires",
                reference_date="2026-04-01",
                periods=[build_quarter_period(2026, 1)],
            )
            message_payload = {
                "id": "m1",
                "internalDate": str(1711929600000),
                "payload": {
                    "headers": [
                        {"name": "Subject", "value": "Dashboard CI 2026-Q1 ARG"},
                        {"name": "From", "value": "sender@example.com"},
                    ],
                    "parts": [
                        {
                            "filename": "Dashboard_CI_2026-Q1_ARG.pdf",
                            "body": {"attachmentId": "a1"},
                        }
                    ],
                },
            }
            fake_service = _FakeService({"m1": message_payload})
            with patch("fetch_dashboard_pdfs.DATA_DIR", data_dir), \
                    patch("fetch_dashboard_pdfs.resolve_schedule_from_env", return_value=schedule), \
                    patch("fetch_dashboard_pdfs.save_schedule"), \
                    patch("fetch_dashboard_pdfs.build_gmail_service", return_value=fake_service), \
                    patch("fetch_dashboard_pdfs.iter_message_ids", return_value=["m1"]), \
                    patch("fetch_dashboard_pdfs.download_attachment", return_value=b"%PDF-1.4 fake pdf"), \
                    patch.dict(os.environ, {
                        "ALLOW_PARTIAL_PERIOD": "false",
                        "REPORT_REQUIRED_SCOPES": "argentina",
                    }, clear=False):
                manifest = run_ingestion(pdf_dir=pdf_dir)
                saved_pdf = pdf_dir / "quarter_2026_Q1_ARG.pdf"
                self.assertTrue(saved_pdf.exists())
                persisted_manifest = json.loads((pdf_dir / "manifest.json").read_text(encoding="utf-8"))

        self.assertEqual(manifest["status"], "ok")
        self.assertEqual(manifest["files"][0]["saved_filename"], "quarter_2026_Q1_ARG.pdf")
        self.assertEqual(manifest["files"][0]["period"], "quarter_2026_Q1")
        self.assertEqual(manifest["files"][0]["scope"], "argentina")
        self.assertIn("checksum_sha256", manifest["files"][0])
        self.assertEqual(persisted_manifest["files"][0]["saved_filename"], "quarter_2026_Q1_ARG.pdf")


if __name__ == "__main__":
    unittest.main()
