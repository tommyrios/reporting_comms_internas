import sys
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from analyzer import (
    build_render_plan,
    compute_kpis,
    validate_monthly_summary_contract,
)


class AnalyzerTests(unittest.TestCase):
    def test_validate_monthly_summary_contract_normalizes_legacy(self):
        legacy = {
            "month": "2026-03",
            "data": {
                "push_volume": 66,
                "pull_notes": 17,
                "pull_reads": 6205,
                "push_opens_pct": 77.53,
                "push_interaction_pct": 8.86,
            },
            "insights": {
                "strategic_axes": [{"theme": "Negocio", "weight": 40}],
                "internal_clients": [{"label": "Retail", "value": 30}],
                "top_push_comm": {"name": "mail_importante", "clicks": 123, "interaction": 7.1, "open_rate": 70},
                "top_pull_note": {"title": "nota_relevante", "unique_reads": 300, "total_reads": 450},
            },
        }
        out = validate_monthly_summary_contract(legacy)
        self.assertEqual(out["plan_total"], 66)
        self.assertEqual(out["site_notes_total"], 17)
        self.assertEqual(out["site_total_views"], 6205)
        self.assertEqual(out["mail_open_rate"], 77.53)
        self.assertIn("quality_flags", out)

    def test_compute_kpis_builds_flags_and_rankings(self):
        summaries = [
            {
                "month": "2026-02",
                "plan_total": 60,
                "site_notes_total": 15,
                "site_total_views": 5000,
                "mail_total": 30,
                "mail_open_rate": 70,
                "mail_interaction_rate": 7,
                "strategic_axes": [{"theme": "Negocio", "weight": 50}],
                "internal_clients": [{"label": "Retail", "value": 40}],
                "channel_mix": [{"label": "Mail", "value": 70}],
                "format_mix": [{"label": "Newsletter", "value": 60}],
                "top_push_by_interaction": [{"name": "push uno", "clicks": 100, "interaction": 9, "open_rate": 70}],
                "top_push_by_open_rate": [{"name": "push dos", "clicks": 80, "interaction": 6, "open_rate": 80}],
                "top_pull_notes": [{"title": "nota 1", "unique_reads": 500, "total_reads": 700}],
                "hitos": [{"title": "hito 1", "description": "desc"}],
                "events": [],
                "quality_flags": {
                    "scope_country": "AR",
                    "scope_mixed": False,
                    "site_has_no_data_sections": False,
                    "events_summary_available": False,
                    "push_ranking_available": True,
                    "pull_ranking_available": True,
                    "historical_comparison_allowed": True,
                },
            },
            {
                "month": "2026-03",
                "plan_total": 66,
                "site_notes_total": 17,
                "site_total_views": 6205,
                "mail_total": 36,
                "mail_open_rate": 77.53,
                "mail_interaction_rate": 8.86,
                "strategic_axes": [{"theme": "Personas", "weight": 30}],
                "internal_clients": [{"label": "Corporate", "value": 35}],
                "channel_mix": [{"label": "Site", "value": 30}],
                "format_mix": [{"label": "Nota", "value": 40}],
                "top_push_by_interaction": [{"name": "push tres", "clicks": 140, "interaction": 10, "open_rate": 76}],
                "top_push_by_open_rate": [{"name": "push cuatro", "clicks": 90, "interaction": 5, "open_rate": 85}],
                "top_pull_notes": [{"title": "nota 2", "unique_reads": 650, "total_reads": 900}],
                "hitos": [{"title": "hito 2", "description": "desc"}],
                "events": [{"name": "Townhall", "participants": 150, "date": "2026-03-12"}],
                "quality_flags": {
                    "scope_country": "AR",
                    "scope_mixed": False,
                    "site_has_no_data_sections": False,
                    "events_summary_available": True,
                    "push_ranking_available": True,
                    "pull_ranking_available": True,
                    "historical_comparison_allowed": True,
                },
            },
        ]
        kpis = compute_kpis(summaries)
        self.assertEqual(kpis["calculated_totals"]["plan_total"], 126)
        self.assertEqual(kpis["calculated_totals"]["site_notes_total"], 32)
        self.assertEqual(kpis["quality_flags"]["events_summary_available"], True)
        self.assertEqual(len(kpis["consolidated_rankings"]["top_push_by_interaction"]), 2)
        self.assertEqual(len(kpis["consolidated_rankings"]["top_pull_notes"]), 2)

    def test_build_render_plan_omits_events_when_not_available(self):
        kpis = {
            "calculated_totals": {"plan_total": 10, "site_notes_total": 2, "site_total_views": 200, "mail_total": 5, "mail_open_rate": 40, "mail_interaction_rate": 5},
            "mixes": {"strategic_axes": [], "internal_clients": [], "channel_mix": [], "format_mix": []},
            "consolidated_rankings": {"top_push_by_interaction": [], "top_push_by_open_rate": [], "top_pull_notes": []},
            "timelines": {"mail_total": [], "site_notes_total": []},
            "hitos": [],
            "events": [],
            "quality_flags": {
                "events_summary_available": False,
                "push_ranking_available": False,
                "pull_ranking_available": False,
                "historical_comparison_allowed": False,
                "site_has_no_data_sections": True,
            },
        }
        plan = build_render_plan({"slug": "month_2026_03", "label": "Marzo 2026"}, kpis, {})
        keys = [m["key"] for m in plan["modules"]]
        self.assertNotIn("events", keys)
        self.assertEqual(keys[0], "executive_summary")

    def test_validate_contract_error_mentions_missing_fields(self):
        with self.assertRaises(ValueError) as ctx:
            validate_monthly_summary_contract({"month": "2026-03"})
        self.assertIn("Contrato mensual incompleto", str(ctx.exception))
        self.assertIn("plan_total", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
