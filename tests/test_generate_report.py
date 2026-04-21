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

    def test_generate_period_report_omits_events_module_without_data(self):
        period = {"slug": "month_2026_03", "months": ["2026-03"], "label": "Marzo 2026", "email_subject": "Subj", "kind": "month"}
        summary = {
            "month": "2026-03",
            "plan_total": 66,
            "site_notes_total": 17,
            "site_total_views": 6205,
            "mail_total": 36,
            "mail_open_rate": 77.53,
            "mail_interaction_rate": 8.86,
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
        narrative = {
            "executive_summary": "Resumen",
            "executive_takeaways": ["a", "b", "c"],
            "channel_management": "x",
            "mix_thematic_clients": "x",
            "ranking_push": "x",
            "ranking_pull": "x",
            "milestones": "x",
            "events": "x",
        }

        with patch("generate_report.get_period_definition", return_value=period), \
                patch("generate_report.build_genai_client", side_effect=[object(), object()]), \
                patch("generate_report.summarize_month", return_value=summary), \
                patch("generate_report.apply_historical_comparison", side_effect=lambda period, payload: payload), \
                patch("generate_report.load_prompt", return_value="prompt"), \
                patch("generate_report.call_gemini_for_json", return_value=narrative), \
                patch("generate_report.load_manual_context", return_value={}), \
                patch("generate_report.persist_calculated_totals"), \
                patch("generate_report.write_report_artifacts", return_value="/tmp/report"):
            result = generate_period_report("month_2026_03")

        self.assertEqual(result["generation_mode"], "llm")
        self.assertIn("Módulo de eventos omitido", result["warning"])

    def test_generate_period_report_fail_fast_invalid_monthly_contract(self):
        period = {"slug": "month_2026_03", "months": ["2026-03"], "label": "Marzo 2026", "email_subject": "Subj", "kind": "month"}
        invalid_summary = {"plan_total": 1}

        with patch("generate_report.get_period_definition", return_value=period), \
                patch("generate_report.build_genai_client", return_value=object()), \
                patch("generate_report.summarize_month", return_value=invalid_summary):
            with self.assertRaises(ValueError):
                generate_period_report("month_2026_03")


if __name__ == "__main__":
    unittest.main()
