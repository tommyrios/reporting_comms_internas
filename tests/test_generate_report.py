import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from generate_report import _deep_merge, generate_period_report


class GenerateReportTests(unittest.TestCase):
    def test_deep_merge_nested_dict_and_list_override(self):
        base = {
            "slide": {"title": "A", "nested": {"x": 1, "y": 2}},
            "items": [1, 2],
            "keep": "yes",
        }
        override = {
            "slide": {"nested": {"y": 9, "z": 3}},
            "items": [7],
            "keep": None,
        }
        merged = _deep_merge(base, override)
        self.assertEqual(merged["slide"]["nested"], {"x": 1, "y": 9, "z": 3})
        self.assertEqual(merged["items"], [7])
        self.assertEqual(merged["keep"], "yes")

    def test_generate_period_report_reuses_cached_summary_when_month_generation_fails(self):
        period = {"slug": "month_2026_03", "months": ["2026-03"], "label": "Marzo 2026", "email_subject": "Subj"}
        cached_summary = {"month": "2026-03", "kpi": 1}
        kpis = {
            "calculated_totals": {},
            "timelines": {},
            "aggregated_distributions": {},
            "consolidated_rankings": {},
            "hitos_crudos": [],
        }
        with tempfile.TemporaryDirectory() as tmp:
            summaries_dir = Path(tmp)
            (summaries_dir / "2026-03.json").write_text(json.dumps(cached_summary), encoding="utf-8")
            with patch("generate_report.get_period_definition", return_value=period), \
                    patch("generate_report.build_genai_client", return_value=object()), \
                    patch("generate_report.summarize_month", side_effect=RuntimeError("503 UNAVAILABLE")), \
                    patch("generate_report.SUMMARIES_DIR", summaries_dir), \
                    patch("generate_report.compute_kpis", return_value=kpis) as compute_kpis_mock, \
                    patch("generate_report.load_prompt", return_value="prompt"), \
                    patch("generate_report.call_gemini_for_json", side_effect=RuntimeError("LLM final down")), \
                    patch("generate_report.build_fallback_report", return_value={"ok": True}), \
                    patch("generate_report.load_manual_context", return_value={}), \
                    patch("generate_report.validate_report_json", side_effect=lambda x: x), \
                    patch("generate_report.write_report_artifacts", return_value="/tmp/report"):
                result = generate_period_report("month_2026_03")

        self.assertEqual(compute_kpis_mock.call_args.args[0], [cached_summary])
        self.assertEqual(result["generation_mode"], "fallback")
        self.assertIn("summary cacheado de 2026-03", result["warning"])
        self.assertIn("sin redacción del LLM", result["warning"])

    def test_generate_period_report_raises_when_month_generation_fails_without_cache(self):
        period = {"slug": "month_2026_03", "months": ["2026-03"], "label": "Marzo 2026", "email_subject": "Subj"}
        with tempfile.TemporaryDirectory() as tmp:
            summaries_dir = Path(tmp)
            with patch("generate_report.get_period_definition", return_value=period), \
                    patch("generate_report.build_genai_client", return_value=object()), \
                    patch("generate_report.summarize_month", side_effect=RuntimeError("503 UNAVAILABLE")), \
                    patch("generate_report.SUMMARIES_DIR", summaries_dir), \
                    patch("generate_report.compute_kpis", new=Mock()) as compute_kpis_mock:
                with self.assertRaises(RuntimeError) as ctx:
                    generate_period_report("month_2026_03")
        self.assertIn("no existe caché previa", str(ctx.exception))
        compute_kpis_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
