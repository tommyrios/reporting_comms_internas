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
    def test_resolve_quarter_only(self):
        env = {
            "REPORT_MODE": "quarter",
            "REPORT_YEAR": "2026",
            "REPORT_QUARTER": "2",
            "REPORT_TIMEZONE": "America/Argentina/Buenos_Aires",
        }
        with patch.dict(os.environ, env, clear=False):
            schedule = resolve_schedule_from_env()

        self.assertEqual(len(schedule.periods), 1)
        self.assertEqual(schedule.periods[0].slug, "quarter_2026_Q2")
        self.assertEqual(schedule.periods[0].kind, "quarter")
        self.assertEqual(unique_months_from_periods(schedule.periods), ["2026-04", "2026-05", "2026-06"])

    def test_resolve_year_only(self):
        env = {
            "REPORT_MODE": "year",
            "REPORT_YEAR": "2026",
            "REPORT_TIMEZONE": "America/Argentina/Buenos_Aires",
        }
        with patch.dict(os.environ, env, clear=False):
            schedule = resolve_schedule_from_env()

        self.assertEqual(len(schedule.periods), 1)
        self.assertEqual(schedule.periods[0].slug, "year_2026")
        self.assertEqual(schedule.periods[0].kind, "year")
        self.assertEqual(schedule.periods[0].months[0], "2026-01")
        self.assertEqual(schedule.periods[0].months[-1], "2026-12")

    def test_invalid_quarter_raises(self):
        env = {
            "REPORT_MODE": "quarter",
            "REPORT_YEAR": "2026",
            "REPORT_QUARTER": "5",
        }
        with patch.dict(os.environ, env, clear=False):
            with self.assertRaises(ValueError):
                resolve_schedule_from_env()


if __name__ == "__main__":
    unittest.main()
