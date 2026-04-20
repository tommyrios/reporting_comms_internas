import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from pdf_processor import summarize_month


class PdfProcessorTests(unittest.TestCase):
    def test_uses_primary_cache_when_available_and_not_forced(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            primary_cache = tmp_dir / "data_cache"
            legacy_cache = tmp_dir / "legacy_cache"
            summary = {"month": "2026-03", "data": {"push_volume": 9}}
            primary_cache.mkdir(parents=True, exist_ok=True)
            (primary_cache / "2026-03.json").write_text(json.dumps(summary), encoding="utf-8")
            with patch("pdf_processor.SUMMARIES_DIR", primary_cache), \
                    patch("pdf_processor.LEGACY_SUMMARIES_DIR", legacy_cache):
                result = summarize_month(Mock(), "2026-03", force_regenerate=False)
        self.assertEqual(result, summary)

    def test_uses_legacy_cache_when_primary_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            primary_cache = tmp_dir / "data_cache"
            legacy_cache = tmp_dir / "legacy_cache"
            summary = {"month": "2026-03", "data": {"push_volume": 7}}
            legacy_cache.mkdir(parents=True, exist_ok=True)
            (legacy_cache / "2026-03.json").write_text(json.dumps(summary), encoding="utf-8")
            with patch("pdf_processor.SUMMARIES_DIR", primary_cache), \
                    patch("pdf_processor.LEGACY_SUMMARIES_DIR", legacy_cache):
                result = summarize_month(Mock(), "2026-03", force_regenerate=False)
        self.assertEqual(result, summary)

    def test_force_regenerate_falls_back_to_existing_cache_when_llm_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            primary_cache = tmp_dir / "data_cache"
            legacy_cache = tmp_dir / "legacy_cache"
            pdf_dir = tmp_dir / "pdfs"
            pdf_dir.mkdir(parents=True, exist_ok=True)
            (pdf_dir / "2026-03.pdf").write_bytes(b"%PDF-1.4 mock content")
            summary = {"month": "2026-03", "data": {"push_volume": 5}}
            primary_cache.mkdir(parents=True, exist_ok=True)
            (primary_cache / "2026-03.json").write_text(json.dumps(summary), encoding="utf-8")
            uploaded = SimpleNamespace(state=SimpleNamespace(name="DONE"), name="files/123")
            client = Mock()
            client.files.delete = Mock()
            with patch("pdf_processor.SUMMARIES_DIR", primary_cache), \
                    patch("pdf_processor.LEGACY_SUMMARIES_DIR", legacy_cache), \
                    patch("pdf_processor.PDF_DIR", pdf_dir), \
                    patch("pdf_processor.load_prompt", return_value="prompt"), \
                    patch("pdf_processor._upload_with_retries", return_value=uploaded), \
                    patch("pdf_processor._wait_until_processed", return_value=uploaded), \
                    patch("pdf_processor.call_gemini_for_json", side_effect=RuntimeError("503 UNAVAILABLE")):
                result = summarize_month(client, "2026-03", force_regenerate=True)
        self.assertEqual(result, summary)

    def test_generates_local_fallback_and_persists_when_llm_fails_without_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            primary_cache = tmp_dir / "data_cache"
            legacy_cache = tmp_dir / "legacy_cache"
            pdf_dir = tmp_dir / "pdfs"
            pdf_dir.mkdir(parents=True, exist_ok=True)
            (pdf_dir / "2026-03.pdf").write_bytes(b"%PDF-1.4 mock content")
            uploaded = SimpleNamespace(state=SimpleNamespace(name="DONE"), name="files/123")
            client = Mock()
            client.files.delete = Mock()
            with patch("pdf_processor.SUMMARIES_DIR", primary_cache), \
                    patch("pdf_processor.LEGACY_SUMMARIES_DIR", legacy_cache), \
                    patch("pdf_processor.PDF_DIR", pdf_dir), \
                    patch("pdf_processor.load_prompt", return_value="prompt"), \
                    patch("pdf_processor._upload_with_retries", return_value=uploaded), \
                    patch("pdf_processor._wait_until_processed", return_value=uploaded), \
                    patch("pdf_processor.call_gemini_for_json", side_effect=RuntimeError("503 UNAVAILABLE")):
                result = summarize_month(client, "2026-03", force_regenerate=False)
            persisted = json.loads((primary_cache / "2026-03.json").read_text(encoding="utf-8"))
        self.assertEqual(result["month"], "2026-03")
        self.assertEqual(result["generation_mode"], "local_fallback")
        self.assertIn("No se pudo generar resumen mensual con Gemini", result["warning"])
        self.assertEqual(persisted["generation_mode"], "local_fallback")


if __name__ == "__main__":
    unittest.main()
