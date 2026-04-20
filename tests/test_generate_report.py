import sys
import unittest
from pathlib import Path
from unittest.mock import patch

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

    def test_generate_period_report_collects_monthly_local_fallback_warning(self):
        period = {"slug": "month_2026_03", "months": ["2026-03"], "label": "Marzo 2026", "email_subject": "Subj"}
        summary_local_fallback = {
            "month": "2026-03",
            "generation_mode": "local_fallback",
            "warning": "No se pudo generar resumen mensual con Gemini",
            "data": {},
            "insights": {},
        }
        kpis = {
            "calculated_totals": {},
            "timelines": {},
            "aggregated_distributions": {},
            "consolidated_rankings": {},
            "hitos_crudos": [],
        }
        with patch("generate_report.get_period_definition", return_value=period), \
                patch("generate_report.build_genai_client", return_value=object()), \
                patch("generate_report.summarize_month", return_value=summary_local_fallback), \
                patch("generate_report.compute_kpis", return_value=kpis) as compute_kpis_mock, \
                patch("generate_report.apply_historical_comparison", side_effect=lambda period, payload: payload), \
                patch("generate_report.load_prompt", return_value="prompt"), \
                patch("generate_report.call_gemini_for_json", side_effect=RuntimeError("LLM final down")), \
                patch("generate_report.build_fallback_report", return_value={"ok": True}), \
                patch("generate_report.load_manual_context", return_value={}), \
                patch("generate_report.validate_report_json", side_effect=lambda x: x), \
                patch("generate_report.persist_calculated_totals"), \
                patch("generate_report.write_report_artifacts", return_value="/tmp/report"):
            result = generate_period_report("month_2026_03")

        self.assertEqual(compute_kpis_mock.call_args.args[0], [summary_local_fallback])
        self.assertEqual(result["generation_mode"], "fallback")
        self.assertIn("No se pudo generar resumen mensual con Gemini", result["warning"])
        self.assertIn("sin redacción del LLM", result["warning"])


if __name__ == "__main__":
    unittest.main()
