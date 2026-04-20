import sys
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from analyzer import BASE_STRUCTURE, compute_kpis, validate_report_json


class AnalyzerTests(unittest.TestCase):
    def test_compute_kpis_uses_weighted_rates(self):
        monthly = [
            {
                "month": "2026-01",
                "data": {
                    "push_volume": 10,
                    "pull_notes": 2,
                    "pull_reads": 50,
                    "push_opens_pct": 90,
                    "push_interaction_pct": 20,
                },
                "insights": {},
            },
            {
                "month": "2026-02",
                "data": {
                    "push_volume": 100,
                    "pull_notes": 3,
                    "pull_reads": 150,
                    "push_opens_pct": 10,
                    "push_interaction_pct": 5,
                },
                "insights": {},
            },
        ]
        kpis = compute_kpis(monthly)
        totals = kpis["calculated_totals"]
        self.assertEqual(totals["average_open_rate"], 17.3)
        self.assertEqual(totals["average_interaction_rate"], 6.4)

    def test_validate_report_json_fills_defaults_and_maps_legacy(self):
        payload = {
            "slide_1_cover": {"period": "Q1 2026"},
            "slide_6_hitos": [{"description": "hito"}],
        }
        validated = validate_report_json(payload)
        self.assertEqual(validated["slide_1_cover"]["period"], "Q1 2026")
        self.assertEqual(validated["slide_7_hitos"], [{"description": "hito"}])
        self.assertEqual(set(validated.keys()), set(BASE_STRUCTURE.keys()))


if __name__ == "__main__":
    unittest.main()
