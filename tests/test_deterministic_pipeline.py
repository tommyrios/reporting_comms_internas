import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from deterministic_pipeline import (
    canonicalize_monthly,
    extract_single_pdf_to_raw,
    extract_raw_monthly_pdf,
    parse_integer_value,
    parse_percent_value,
    validate_canonical_monthly,
)


class DeterministicPipelineTests(unittest.TestCase):
    def test_extract_single_pdf_to_raw_writes_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            input_pdf = tmp_dir / "2026-01_dashboard.pdf"
            output_json = tmp_dir / "debug" / "2026-01_raw.json"
            input_pdf.write_bytes(b"%PDF-1.4 fake")
            fake_raw = {"month": "2026-01", "metrics": {"plan_total": {"value": 58}}}
            with patch("deterministic_pipeline.extract_raw_monthly_pdf", return_value=fake_raw) as extract_mock:
                result = extract_single_pdf_to_raw(input_pdf, output_json)
            self.assertTrue(output_json.exists())
        self.assertEqual(result["month"], "2026-01")
        self.assertEqual(extract_mock.call_args[0][0], "2026-01")

    def _valid_sample_pages(self) -> list[str]:
        return [
            "Resumen planificación\nNº total de comunicaciones 66\nMedia comunicaciones diarias 3",
            (
                "Resumen site\nNoticias Publicadas 17\nTotal Páginas Vistas 6.205\nPromedio Vistas 365\n"
                "Top five noticias\n1 Nota A 520\n2 Nota B 490\n3 Nota C 410\n4 Nota D 390\n5 Nota E 360"
            ),
            (
                "Resumen mailing\nMails enviados 36\nTasa de apertura promedio 77,53%\n"
                "Tasa de interacción sobre mails enviados 8,86%\n"
                "Tasa de interacción sobre mails abiertos 11,43%"
            ),
        ]

    def _jan_2026_shifted_fixture_pages(self) -> list[str]:
        return [
            "Carátula del dashboard enero 2026",
            "Panel planificación\nN° total de comunicaciones 58 Media comunicaciones diarias 3",
            "Panel site\nNoticias Publicadas 16\nTotal Páginas Vistas 5,580\nPromedio Vistas 349",
            (
                "Panel mail\nMails enviados 42\n"
                "Tasa de apertura promedio 78,68% Tasa de interacción sobre mails enviados 9,20%\n"
                "Tasa de interacción sobre mails abiertos 11,70%"
            ),
        ]

    def test_extract_raw_monthly_pdf_extracts_exact_anchors(self):
        with patch("deterministic_pipeline._extract_pages_text", return_value=self._valid_sample_pages()):
            raw = extract_raw_monthly_pdf("2026-03", Path("/tmp/fake.pdf"))
        self.assertEqual(raw["metrics"]["plan_total"]["value"], 66.0)
        self.assertEqual(raw["metrics"]["site_notes_total"]["value"], 17.0)
        self.assertEqual(raw["metrics"]["site_total_views"]["value"], 6205.0)
        self.assertEqual(raw["metrics"]["mail_total"]["value"], 36.0)
        self.assertEqual(raw["metrics"]["mail_open_rate"]["value"], 77.53)
        self.assertEqual(raw["metrics"]["mail_interaction_rate"]["value"], 8.86)

    def test_extract_raw_monthly_pdf_marks_anchor_when_found_on_unexpected_page(self):
        pages = self._jan_2026_shifted_fixture_pages()
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-01", Path("/tmp/fake.pdf"))
        self.assertEqual(raw["metrics"]["mail_open_rate"]["value"], 78.68)
        self.assertTrue(any("anchor_out_of_expected_page:Tasa de apertura promedio" in w for w in raw["warnings"]))

    def test_extract_raw_monthly_pdf_missing_anchor_returns_null_and_warning(self):
        pages = [
            "Nº total de comunicaciones 66\nMedia comunicaciones diarias 3",
            "Noticias Publicadas 17\nPromedio Vistas 365",
            "Mails enviados 36\nTasa de apertura promedio 77,53%\nTasa de interacción sobre mails enviados 8,86%",
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-03", Path("/tmp/fake.pdf"))
        self.assertIsNone(raw["metrics"]["site_total_views"]["value"])
        self.assertTrue(any("missing_anchor:site_total_views" in w for w in raw["warnings"]))

    def test_parse_percent_value_works_with_comma(self):
        self.assertEqual(parse_percent_value("80,75%"), 80.75)

    def test_parse_integer_value_dashboard_thousands(self):
        self.assertEqual(parse_integer_value("11.785"), 11785)
        self.assertEqual(parse_integer_value("5,580"), 5580)

    def test_top_five_does_not_set_site_notes_to_five(self):
        pages = [
            "Nº total de comunicaciones 70\nMedia comunicaciones diarias 3",
            "Top five noticias\n1 A\n2 B\n3 C\n4 D\n5 E\nNoticias Publicadas 17\nTotal Páginas Vistas 6205\nPromedio Vistas 365",
            "Mails enviados 36\nTasa de apertura promedio 77,53%\nTasa de interacción sobre mails enviados 8,86%\nTasa de interacción sobre mails abiertos 11,43%",
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-03", Path("/tmp/fake.pdf"))
        self.assertEqual(raw["metrics"]["site_notes_total"]["value"], 17.0)

    def test_regression_jan_feb_mar_like_structure_does_not_produce_nonsense(self):
        pages = [
            "Panel planificación\nNº total de comunicaciones 58\nMedia comunicaciones diarias 3",
            "Panel site\nNoticias Publicadas 16\nTotal Páginas Vistas 5,580\nPromedio Vistas 349\nTop five",
            "Panel mail\nMails enviados 42\nTasa de apertura promedio 78,68%\nTasa de interacción sobre mails enviados 9,20%\nTasa de interacción sobre mails abiertos 11,70%",
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-01", Path("/tmp/fake.pdf"))
        canonical = canonicalize_monthly(raw)
        validation = validate_canonical_monthly(canonical)
        self.assertTrue(validation["is_valid"])
        self.assertNotEqual(canonical["plan_total"], 0)
        self.assertGreater(canonical["site_total_views"], canonical["site_notes_total"])
        self.assertNotEqual(canonical["mail_total"], 5)


    def test_extract_raw_monthly_pdf_prefers_number_after_anchor_not_last_number_in_line(self):
        pages = self._jan_2026_shifted_fixture_pages()
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-01", Path("/tmp/fake.pdf"))
        self.assertEqual(raw["metrics"]["plan_total"]["value"], 58.0)
        self.assertEqual(raw["metrics"]["mail_open_rate"]["value"], 78.68)
        self.assertEqual(raw["metrics"]["mail_interaction_rate"]["value"], 9.2)

    def test_extract_raw_monthly_pdf_falls_back_to_other_page_when_layout_is_shifted(self):
        pages = [
            "Carátula del dashboard",
            "Panel planificación\nNº total de comunicaciones 58\nMedia comunicaciones diarias 3",
            "Panel site\nNoticias Publicadas 16\nTotal Páginas Vistas 5,580\nPromedio Vistas 349",
            "Panel mail\nMails enviados 42\nTasa de apertura promedio 78,68%\nTasa de interacción sobre mails enviados 9,20%\nTasa de interacción sobre mails abiertos 11,70%",
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-01", Path("/tmp/fake.pdf"))
        canonical = canonicalize_monthly(raw)
        validation = validate_canonical_monthly(canonical)
        self.assertEqual(canonical["plan_total"], 58)
        self.assertEqual(canonical["site_total_views"], 5580)
        self.assertEqual(canonical["mail_open_rate"], 78.68)
        self.assertTrue(validation["is_valid"])
        self.assertTrue(any(str(w).startswith("anchor_out_of_expected_page:") for w in raw["warnings"]))

    def test_jan_2026_fixture_does_not_collapse_mail_metrics(self):
        pages = self._jan_2026_shifted_fixture_pages()
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-01", Path("/tmp/fake.pdf"))
        canonical = canonicalize_monthly(raw)
        validation = validate_canonical_monthly(canonical)
        self.assertNotEqual(canonical["mail_open_rate"], canonical["mail_interaction_rate"])
        self.assertTrue(validation["is_valid"])

    def test_validate_canonical_monthly_rejects_missing_anchor(self):
        raw = {
            "month": "2026-03",
            "metrics": {
                "plan_total": {"value": 66},
                "site_notes_total": {"value": 17},
                "site_total_views": {"value": None},
                "mail_total": {"value": 36},
                "mail_open_rate": {"value": 77.53},
                "mail_interaction_rate": {"value": 8.86},
            },
            "warnings": ["missing_anchor:site_total_views:No se encontró ancla exacta"],
            "parser": "deterministic_pdf_v2",
        }
        canonical = canonicalize_monthly(raw)
        validation = validate_canonical_monthly(canonical)
        self.assertFalse(validation["is_valid"])
        self.assertIn("Faltan KPIs primarios por ancla exacta", validation["errors"])

    def test_validate_canonical_monthly_rejects_suspicious_metrics(self):
        canonical = {
            "month": "2026-03",
            "plan_total": 66,
            "site_notes_total": 16,
            "site_total_views": 64,
            "mail_total": 5,
            "mail_open_rate": 78.68,
            "mail_interaction_rate": 78.68,
            "extraction_warnings": [],
        }
        validation = validate_canonical_monthly(canonical)
        self.assertFalse(validation["is_valid"])
        self.assertIn("site_total_views sospechosamente bajo respecto a site_notes_total", validation["errors"])
        self.assertIn("mail_total sospechosamente bajo respecto a plan_total", validation["errors"])
        self.assertIn("mail_open_rate y mail_interaction_rate no deberían colapsar al mismo valor", validation["errors"])


if __name__ == "__main__":
    unittest.main()
