import sys
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from data_quality import normalize_push_row, sanitize_push_ranking, validate_canonical_quality, validate_report_quality


class DataQualityTests(unittest.TestCase):
    def test_normalize_push_row_flags_high_interaction_without_clicks(self):
        row = normalize_push_row({"name": "Mail líder", "clicks": 0, "open_rate": 93.2, "interaction": 73.0})
        self.assertFalse(row["data_complete"])
        self.assertEqual(row["data_quality_issue"], "interaccion_alta_sin_clicks")

    def test_sanitize_push_ranking_sorts_and_normalizes_percentages(self):
        rows = sanitize_push_ranking([
            {"name": "B", "clicks": 10, "open_rate": "90%", "interaction": "20%"},
            {"name": "A", "clicks": 20, "open_rate": "95%", "interaction": "0,35"},
        ], metric_key="interaction")
        self.assertEqual(rows[0]["name"], "A")
        self.assertEqual(rows[0]["interaction"], 35.0)

    def test_validate_canonical_quality_detects_pull_inconsistency(self):
        canonical = {
            "month": "2026-01",
            "plan_total": 54,
            "site_notes_total": 10,
            "site_total_views": 4071,
            "mail_total": 23,
            "mail_open_rate": 80.05,
            "mail_interaction_rate": 12.54,
            "strategic_axes": [],
            "internal_clients": [],
            "channel_mix": [],
            "format_mix": [],
            "top_push_by_interaction": [],
            "top_push_by_open_rate": [],
            "top_pull_notes": [{"title": "Nota", "unique_reads": 200, "total_reads": 100}],
            "quality_flags": {},
        }
        result = validate_canonical_quality(canonical)
        self.assertTrue(result["is_valid"])
        self.assertIn("usuarios únicos mayores", " | ".join(result["warnings"]))

    def test_validate_report_quality_requires_modules(self):
        result = validate_report_quality({"period": {}, "kpis": {}, "narrative": {}, "quality_flags": {}, "render_plan": {"modules": []}})
        self.assertFalse(result["is_valid"])
        self.assertIn("render_plan.modules vacío", result["errors"])


if __name__ == "__main__":
    unittest.main()
