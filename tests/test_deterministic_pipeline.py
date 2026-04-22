import sys
import unittest
from pathlib import Path
from unittest.mock import patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from deterministic_pipeline import canonicalize_monthly, extract_raw_monthly_pdf, validate_canonical_monthly


class DeterministicPipelineTests(unittest.TestCase):
    def test_extract_raw_monthly_pdf_parses_percent_and_ratio(self):
        pages = [
            "Mail enviados 120\nTasa apertura 73,07%\nTasa interacción 0.0825\nPlan 120\nNotas publicadas 10\nPáginas vistas 5000",
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-03", Path("/tmp/fake.pdf"))
        self.assertEqual(raw["metrics"]["mail_open_rate"]["value"], 73.07)
        self.assertEqual(raw["metrics"]["mail_interaction_rate"]["value"], 8.25)

    def test_validate_canonical_monthly_rejects_out_of_range_percent(self):
        canonical = canonicalize_monthly(
            {
                "month": "2026-03",
                "metrics": {
                    "plan_total": {"value": 1},
                    "site_notes_total": {"value": 1},
                    "site_total_views": {"value": 1},
                    "mail_total": {"value": 1},
                    "mail_open_rate": {"value": 170},
                    "mail_interaction_rate": {"value": 2},
                },
                "warnings": [],
                "parser": "deterministic_pdf_v1",
            }
        )
        validation = validate_canonical_monthly(canonical)
        self.assertFalse(validation["is_valid"])
        self.assertIn("mail_open_rate fuera de rango 0-100", validation["errors"])


if __name__ == "__main__":
    unittest.main()
