import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from reporting_periods import resolve_schedule_from_env, unique_months_from_periods


class ReportingPeriodsTests(unittest.TestCase):
    def test_resolve_month_and_quarter(self):
        env = {
            "REPORT_MODE": "month_and_quarter",
            "REPORT_YEAR": "2026",
            "REPORT_QUARTER": "2",
            "REPORT_TIMEZONE": "America/Argentina/Buenos_Aires",
        }
        with patch.dict(os.environ, env, clear=False):
            schedule = resolve_schedule_from_env()

        self.assertEqual(len(schedule.periods), 2)
        self.assertEqual(schedule.periods[0].slug, "month_2026_06")
        self.assertEqual(schedule.periods[1].slug, "quarter_2026_Q2")
        self.assertEqual(unique_months_from_periods(schedule.periods), ["2026-04", "2026-05", "2026-06"])

    def test_invalid_month_raises(self):
        env = {
            "REPORT_MODE": "month",
            "REPORT_YEAR": "2026",
            "REPORT_MONTH": "13",
        }
        with patch.dict(os.environ, env, clear=False):
            with self.assertRaises(ValueError):
                resolve_schedule_from_env()


if __name__ == "__main__":
    unittest.main()
