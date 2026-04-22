import json
import sys
import tempfile
import unittest
from pathlib import Path
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

    def test_force_regenerate_falls_back_to_existing_cache_when_deterministic_extraction_fails(self):
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
            with patch("pdf_processor.SUMMARIES_DIR", primary_cache), \
                    patch("pdf_processor.LEGACY_SUMMARIES_DIR", legacy_cache), \
                    patch("pdf_processor.CANONICAL_MONTHLY_DIR", primary_cache), \
                    patch("pdf_processor.PDF_DIR", pdf_dir), \
                    patch("pdf_processor.extract_raw_monthly_pdf", side_effect=RuntimeError("parser failed")):
                result = summarize_month(Mock(), "2026-03", force_regenerate=True)
        self.assertEqual(result, summary)

    def test_generates_local_fallback_and_persists_when_deterministic_extraction_fails_without_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            primary_cache = tmp_dir / "data_cache"
            legacy_cache = tmp_dir / "legacy_cache"
            pdf_dir = tmp_dir / "pdfs"
            pdf_dir.mkdir(parents=True, exist_ok=True)
            (pdf_dir / "2026-03.pdf").write_bytes(b"%PDF-1.4 mock content")
            with patch("pdf_processor.SUMMARIES_DIR", primary_cache), \
                    patch("pdf_processor.LEGACY_SUMMARIES_DIR", legacy_cache), \
                    patch("pdf_processor.CANONICAL_MONTHLY_DIR", primary_cache), \
                    patch("pdf_processor.PDF_DIR", pdf_dir), \
                    patch("pdf_processor.extract_raw_monthly_pdf", side_effect=RuntimeError("parser failed")):
                result = summarize_month(Mock(), "2026-03", force_regenerate=False)
            persisted = json.loads((primary_cache / "2026-03.json").read_text(encoding="utf-8"))
        self.assertEqual(result["month"], "2026-03")
        self.assertEqual(result["generation_mode"], "local_fallback")
        self.assertIn("No se pudo generar resumen mensual determinístico", result["warning"])
        self.assertEqual(persisted["generation_mode"], "local_fallback")

    def test_persists_artifacts_when_deterministic_extraction_succeeds(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            cache_dir = tmp_dir / "cache"
            legacy_cache = tmp_dir / "legacy"
            raw_dir = tmp_dir / "raw"
            validation_dir = tmp_dir / "validation"
            pdf_dir = tmp_dir / "pdfs"
            pdf_dir.mkdir(parents=True, exist_ok=True)
            (pdf_dir / "2026-03.pdf").write_bytes(b"%PDF-1.4 mock content")
            raw = {"month": "2026-03", "metrics": {}}
            canonical = {
                "month": "2026-03",
                "generation_mode": "deterministic_pdf",
                "plan_total": 1,
                "site_notes_total": 2,
                "site_total_views": 3,
                "mail_total": 4,
                "mail_open_rate": 70.0,
                "mail_interaction_rate": 8.0,
                "strategic_axes": [],
                "internal_clients": [],
                "channel_mix": [],
                "format_mix": [],
                "top_push_by_interaction": [],
                "top_push_by_open_rate": [],
                "top_pull_notes": [],
                "hitos": [],
                "events": [],
                "quality_flags": {
                    "scope_country": "AR",
                    "scope_mixed": False,
                    "site_has_no_data_sections": False,
                    "events_summary_available": False,
                    "push_ranking_available": False,
                    "pull_ranking_available": False,
                    "historical_comparison_allowed": True,
                },
            }
            validation = {"is_valid": True, "warnings": [], "errors": []}
            with patch("pdf_processor.SUMMARIES_DIR", cache_dir), \
                    patch("pdf_processor.CANONICAL_MONTHLY_DIR", cache_dir), \
                    patch("pdf_processor.LEGACY_SUMMARIES_DIR", legacy_cache), \
                    patch("pdf_processor.PDF_DIR", pdf_dir), \
                    patch("pdf_processor.extract_raw_monthly_pdf", return_value=raw), \
                    patch("pdf_processor.canonicalize_monthly", return_value=canonical), \
                    patch("pdf_processor.validate_canonical_monthly", return_value=validation), \
                    patch("pdf_processor.persist_monthly_artifacts") as persist_artifacts:
                result = summarize_month(Mock(), "2026-03", force_regenerate=True)
                self.assertEqual(result["generation_mode"], "deterministic_pdf")
                self.assertTrue((cache_dir / "2026-03.json").exists())
                persist_artifacts.assert_called_once()


if __name__ == "__main__":
    unittest.main()
