import json
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from history_manager import apply_historical_comparison, persist_calculated_totals


class HistoryManagerTests(unittest.TestCase):
    def test_apply_historical_comparison_without_previous_data(self):
        period = {"kind": "month", "months": ["2026-03"], "slug": "month_2026_03"}
        kpis = {"calculated_totals": {"push_volume_period": 27}, "quality_flags": {"scope_country": "AR", "historical_comparison_allowed": True}}

        with tempfile.TemporaryDirectory() as tmpdir:
            history_path = Path(tmpdir) / "historico_kpis.json"
            enriched = apply_historical_comparison(period, kpis, history_path=history_path)

        self.assertEqual(enriched["calculated_totals"]["volume_previous"], "Sin datos previos")
        self.assertEqual(enriched["calculated_totals"]["volume_change"], "Sin datos previos")

    def test_apply_historical_comparison_with_previous_month(self):
        period_feb = {"kind": "month", "months": ["2026-02"], "slug": "month_2026_02"}
        period_mar = {"kind": "month", "months": ["2026-03"], "slug": "month_2026_03"}
        kpis_feb = {"calculated_totals": {"push_volume_period": 20}, "quality_flags": {"scope_country": "AR", "historical_comparison_allowed": True}}
        kpis_mar = {"calculated_totals": {"push_volume_period": 30}, "quality_flags": {"scope_country": "AR", "historical_comparison_allowed": True}}

        with tempfile.TemporaryDirectory() as tmpdir:
            history_path = Path(tmpdir) / "historico_kpis.json"
            persist_calculated_totals(period_feb, kpis_feb, history_path=history_path)
            enriched = apply_historical_comparison(period_mar, kpis_mar, history_path=history_path)

            payload = json.loads(history_path.read_text(encoding="utf-8"))

        self.assertIn("month:2026-02", payload["records"])
        self.assertEqual(enriched["calculated_totals"]["volume_previous"], 20)
        self.assertEqual(enriched["calculated_totals"]["volume_change"], "50.0%")
        self.assertEqual(enriched["calculated_totals"]["previous_push_volume"], 20)
        self.assertEqual(enriched["calculated_totals"]["latest_push_variation"], "50.0%")

    def test_disable_comparison_when_scope_changes(self):
        period_feb = {"kind": "month", "months": ["2026-02"], "slug": "month_2026_02"}
        period_mar = {"kind": "month", "months": ["2026-03"], "slug": "month_2026_03"}
        kpis_feb = {"calculated_totals": {"push_volume_period": 20}, "quality_flags": {"scope_country": "AR", "historical_comparison_allowed": True}}
        kpis_mar = {"calculated_totals": {"push_volume_period": 30}, "quality_flags": {"scope_country": "AR,UY", "historical_comparison_allowed": False}}

        with tempfile.TemporaryDirectory() as tmpdir:
            history_path = Path(tmpdir) / "historico_kpis.json"
            persist_calculated_totals(period_feb, kpis_feb, history_path=history_path)
            enriched = apply_historical_comparison(period_mar, kpis_mar, history_path=history_path)

        self.assertEqual(enriched["calculated_totals"]["volume_change"], "No comparable por alcance de fuente")
        self.assertFalse(enriched["quality_flags"]["historical_comparison_allowed"])

    def test_keep_comparison_when_previous_scope_missing(self):
        period_feb = {"kind": "month", "months": ["2026-02"], "slug": "month_2026_02"}
        period_mar = {"kind": "month", "months": ["2026-03"], "slug": "month_2026_03"}
        kpis_feb = {"calculated_totals": {"push_volume_period": 20}}
        kpis_mar = {"calculated_totals": {"push_volume_period": 30}, "quality_flags": {"scope_country": "AR", "historical_comparison_allowed": True}}

        with tempfile.TemporaryDirectory() as tmpdir:
            history_path = Path(tmpdir) / "historico_kpis.json"
            persist_calculated_totals(period_feb, kpis_feb, history_path=history_path)
            enriched = apply_historical_comparison(period_mar, kpis_mar, history_path=history_path)

        self.assertEqual(enriched["calculated_totals"]["volume_change"], "50.0%")


if __name__ == "__main__":
    unittest.main()
