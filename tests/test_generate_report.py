import os
import sys
import unittest
from copy import deepcopy
from pathlib import Path
from unittest.mock import patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from generate_report import _deep_merge, generate_period_report


def _valid_summary(scope="combined", *, validation_valid=True, events=None):
    return {
        "period": "quarter_2026_Q1",
        "month": "quarter_2026_Q1",
        "scope": scope,
        "scope_label": {"argentina": "Argentina", "holding": "Holding", "combined": "Argentina + Holding"}.get(scope, scope),
        "generation_mode": "deterministic_pdf",
        "validation": {"is_valid": validation_valid, "errors": [] if validation_valid else ["error de prueba"], "warnings": []},
        "plan_total": 123,
        "site_notes_total": 34,
        "site_total_views": 38297,
        "site_average_views": 1126,
        "mail_total": 70,
        "mail_open_rate": 79.44,
        "mail_interaction_rate": 12.6,
        "mail_interaction_rate_over_opened": 15.88,
        "strategic_axes": [],
        "internal_clients": [],
        "channel_mix": [],
        "format_mix": [],
        "top_push_by_interaction": [],
        "top_push_by_open_rate": [],
        "top_pull_notes": [],
        "top_pull_notes_tgm": [],
        "hitos": [],
        "events": events if events is not None else [],
        "quality_flags": {
            "scope_country": scope,
            "scope_mixed": scope == "combined",
            "site_has_no_data_sections": False,
            "events_summary_available": bool(events),
            "push_ranking_available": False,
            "pull_ranking_available": False,
            "historical_comparison_allowed": False,
        },
    }


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

    def test_generate_period_report_uses_deterministic_mode(self):
        period = {
            "slug": "quarter_2026_Q1",
            "months": ["2026-01", "2026-02", "2026-03"],
            "label": "Q1 2026",
            "email_subject": "Subj",
            "kind": "quarter",
        }
        summaries = {
            "argentina": _valid_summary("argentina"),
            "holding": _valid_summary("holding"),
            "combined": _valid_summary("combined"),
        }

        with patch.dict(os.environ, {"REPORT_REQUIRED_SCOPES": "argentina,holding,combined"}, clear=False), \
                patch("generate_report.get_period_definition", return_value=period), \
                patch("generate_report.resolve_period_scope_pdfs", return_value={}), \
                patch("generate_report.summarize_period_scope", side_effect=lambda period, scope, **kwargs: deepcopy(summaries[scope])), \
                patch("generate_report.apply_historical_comparison", side_effect=lambda period, payload: payload), \
                patch("generate_report._request_narrative", side_effect=AssertionError("no debe invocarse")), \
                patch("generate_report.load_manual_context", return_value={}), \
                patch("generate_report.persist_calculated_totals"), \
                patch("generate_report.write_report_artifacts", return_value="/tmp/report"):
            result = generate_period_report("quarter_2026_Q1")

        self.assertEqual(result["generation_mode"], "deterministic")
        self.assertIsNone(result["warning"])

    def test_generate_period_report_fail_fast_invalid_period_contract(self):
        period = {
            "slug": "quarter_2026_Q1",
            "months": ["2026-01", "2026-02", "2026-03"],
            "label": "Q1 2026",
            "email_subject": "Subj",
            "kind": "quarter",
        }
        invalid_summary = {"plan_total": 1}

        with patch.dict(os.environ, {"REPORT_REQUIRED_SCOPES": "argentina,holding,combined"}, clear=False), \
                patch("generate_report.get_period_definition", return_value=period), \
                patch("generate_report.resolve_period_scope_pdfs", return_value={}), \
                patch("generate_report.summarize_period_scope", return_value=invalid_summary):
            with self.assertRaises(ValueError) as ctx:
                generate_period_report("quarter_2026_Q1")
        self.assertIn("Contrato mensual incompleto", str(ctx.exception))

    def test_generate_period_report_requires_all_scopes(self):
        period = {"slug": "quarter_2026_Q1", "label": "Q1 2026", "email_subject": "Subj", "kind": "quarter"}
        with patch.dict(os.environ, {"REPORT_REQUIRED_SCOPES": "argentina,combined"}, clear=False), \
                patch("generate_report.get_period_definition", return_value=period):
            with self.assertRaises(ValueError) as ctx:
                generate_period_report("quarter_2026_Q1")
        self.assertIn("Se requieren scopes argentina, holding y combined", str(ctx.exception))

    def test_generate_period_report_does_not_continue_invalid_scope_validation(self):
        period = {
            "slug": "quarter_2026_Q1",
            "months": ["2026-01", "2026-02", "2026-03"],
            "label": "Q1 2026",
            "email_subject": "Subj",
            "kind": "quarter",
        }
        invalid_scope = _valid_summary("argentina", validation_valid=False)

        with patch.dict(os.environ, {"REPORT_REQUIRED_SCOPES": "argentina,holding,combined"}, clear=False), \
                patch("generate_report.get_period_definition", return_value=period), \
                patch("generate_report.resolve_period_scope_pdfs", return_value={}), \
                patch("generate_report.summarize_period_scope", return_value=invalid_scope) as summarize_scope_mock:
            with self.assertRaises(ValueError) as ctx:
                generate_period_report("quarter_2026_Q1")
        self.assertIn("Resumen inválido para quarter_2026_Q1 [argentina]", str(ctx.exception))
        self.assertEqual(summarize_scope_mock.call_count, 1)


if __name__ == "__main__":
    unittest.main()
