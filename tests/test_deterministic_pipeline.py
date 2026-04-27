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
    infer_month_key_from_pdf_path,
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
            "Resumen planificación\nNº total de comunicaciones\n66\nMedia comunicaciones diarias\n3",
            (
                "Resumen site\nNoticias Publicadas\n17\nTotal Páginas Vistas\n6.205\nPromedio Vistas\n365\n"
                "Top five noticias\n1 Nota A 520\n2 Nota B 490\n3 Nota C 410\n4 Nota D 390\n5 Nota E 360"
            ),
            (
                "Resumen mailing\nMails enviados\n36\nTasa de apertura promedio\n77,53%\n"
                "Tasa de interacción sobre mails enviados\n8,86%\n"
                "Tasa de interacción sobre mails abiertos\n11,43%"
            ),
        ]

    def _jan_2026_shifted_fixture_pages(self) -> list[str]:
        return [
            "Carátula del dashboard enero 2026",
            "Panel planificación\nN° total de comunicaciones\n58\nMedia comunicaciones diarias\n3",
            "Panel site\nNoticias Publicadas\n16\nTotal Páginas Vistas\n5,580\nPromedio Vistas\n349",
            (
                "Panel mail\nMails enviados\n42\n"
                "Tasa de apertura promedio\n78,68%\nTasa de interacción sobre mails enviados\n9,20%\n"
                "Tasa de interacción sobre mails abiertos\n11,70%"
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
            "Nº total de comunicaciones\n66\nMedia comunicaciones diarias\n3",
            "Noticias Publicadas\n17\nPromedio Vistas\n365",
            "Mails enviados\n36\nTasa de apertura promedio\n77,53%\nTasa de interacción sobre mails enviados\n8,86%",
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
            "Nº total de comunicaciones\n70\nMedia comunicaciones diarias\n3",
            "Top five noticias\n1 A\n2 B\n3 C\n4 D\n5 E\nNoticias Publicadas\n17\nTotal Páginas Vistas\n6205\nPromedio Vistas\n365",
            "Mails enviados\n36\nTasa de apertura promedio\n77,53%\nTasa de interacción sobre mails enviados\n8,86%\nTasa de interacción sobre mails abiertos\n11,43%",
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-03", Path("/tmp/fake.pdf"))
        self.assertEqual(raw["metrics"]["site_notes_total"]["value"], 17.0)

    def test_regression_jan_feb_mar_like_structure_does_not_produce_nonsense(self):
        pages = [
            "Panel planificación\nNº total de comunicaciones\n58\nMedia comunicaciones diarias\n3",
            "Panel site\nNoticias Publicadas\n16\nTotal Páginas Vistas\n5,580\nPromedio Vistas\n349\nTop five",
            "Panel mail\nMails enviados\n42\nTasa de apertura promedio\n78,68%\nTasa de interacción sobre mails enviados\n9,20%\nTasa de interacción sobre mails abiertos\n11,70%",
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
            "Panel planificación\nNº total de comunicaciones\n58\nMedia comunicaciones diarias\n3",
            "Panel site\nNoticias Publicadas\n16\nTotal Páginas Vistas\n5,580\nPromedio Vistas\n349",
            "Panel mail\nMails enviados\n42\nTasa de apertura promedio\n78,68%\nTasa de interacción sobre mails enviados\n9,20%\nTasa de interacción sobre mails abiertos\n11,70%",
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

    def test_extract_raw_monthly_pdf_realistic_extract_text_layout_regression(self):
        pages = [
            (
                "Página 1\nMedia comunicaciones diarias\n0.06\nN total de comunicaciones por mes\n100\n"
                "Nº total de comunicaciones\n54"
            ),
            "Página 2\nTotal Páginas Vistas\n4,071\nNoticias Publicadas\n10\nPromedio Vistas\n407",
            (
                "Página 3\nTasa de apertura promedio\n80.05%\nTasa de interacción sobre mails enviados\n12.54%\n"
                "Tasa de interacción sobre mails abiertos\n15.67%\nMails enviados\n23"
            ),
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-02", Path("/tmp/fake.pdf"))
        canonical = canonicalize_monthly(raw)
        validation = validate_canonical_monthly(canonical)
        self.assertEqual(canonical["plan_daily_average"], 0.06)
        self.assertEqual(canonical["plan_total"], 54)
        self.assertEqual(canonical["site_total_views"], 4071)
        self.assertEqual(canonical["site_notes_total"], 10)
        self.assertEqual(canonical["site_average_views"], 407)
        self.assertEqual(canonical["mail_open_rate"], 80.05)
        self.assertEqual(canonical["mail_interaction_rate"], 12.54)
        self.assertEqual(canonical["mail_interaction_rate_over_opened"], 15.67)
        self.assertEqual(canonical["mail_total"], 23)
        self.assertNotEqual(canonical["mail_open_rate"], canonical["mail_interaction_rate"])
        self.assertTrue(validation["is_valid"])

    def test_extract_raw_monthly_pdf_real_layout_with_glued_anchors_does_not_cross_kpis(self):
        pages = [
            "Página 1\nMedia comunicaciones diarias0.06Nº total de comunicaciones54",
            (
                "Página 2\nJan 22, 2026 Nota de ejemplo ARGENTINA 96 123Total Páginas Vistas\n"
                "Noticias Publicadas10Total Páginas Vistas4071Promedio Vistas407"
            ),
            (
                "Página 3\nTasa de apertura promedio80.05%Tasa de interacción sobre mails enviados12.54%\n"
                "Tasa de interacción sobre mails abiertos15.67%Mails enviados23"
            ),
        ]
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-01", Path("/tmp/fake.pdf"))
        canonical = canonicalize_monthly(raw)
        validation = validate_canonical_monthly(canonical)
        self.assertEqual(canonical["plan_daily_average"], 0.06)
        self.assertEqual(canonical["plan_total"], 54)
        self.assertNotEqual(canonical["plan_total"], 0)
        self.assertEqual(canonical["site_total_views"], 4071)
        self.assertNotEqual(canonical["site_total_views"], 123)
        self.assertEqual(canonical["mail_open_rate"], 80.05)
        self.assertEqual(canonical["mail_interaction_rate"], 12.54)
        self.assertNotEqual(canonical["mail_open_rate"], canonical["mail_interaction_rate"])
        self.assertTrue(validation["is_valid"])

    def test_jan_2026_fixture_does_not_collapse_mail_metrics(self):
        pages = self._jan_2026_shifted_fixture_pages()
        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-01", Path("/tmp/fake.pdf"))
        canonical = canonicalize_monthly(raw)
        validation = validate_canonical_monthly(canonical)
        self.assertNotEqual(canonical["mail_open_rate"], canonical["mail_interaction_rate"])
        self.assertTrue(validation["is_valid"])

    def test_validate_canonical_fails_missing_required_anchors(self):
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
        self.assertIn(
            "Faltan KPIs primarios por ancla exacta: site_total_views",
            validation["errors"],
        )

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

    def test_canonicalize_monthly_normalizes_contract_aliases(self):
        raw = {
            "month": "2026-04",
            "parser": "deterministic_pdf_v6_kpi_push_pull",
            "metrics": {
                "plan_daily_average": {"value": 2.4},
                "plan_total": {"value": 48},
                "site_notes_total": {"value": 15},
                "site_total_views": {"value": 5200},
                "site_average_views": {"value": 347},
                "mail_total": {"value": 40},
                "mail_open_rate": {"value": 79.12},
                "mail_interaction_rate": {"value": 9.42},
                "mail_interaction_rate_over_opened": {"value": 11.9},
            },
            "strategic_axes": [{"axis": "RCP", "count": "20"}],
            "channel_mix": [{"channel": "Mail", "pct": "70%"}],
            "format_mix": [{"format": "Noticia propia", "pct": "55,5%"}],
            "top_push_interaction": [{"title": "Comms 1", "clicks": "123", "ctr": "10,5%", "open_rate": "80%"}],
            "top_push_open": [{"name": "Comms 2", "clicks": "98", "interaction_rate": "8,2%", "open_rate": "78,4%"}],
            "top_pull_notes": [{"name": "Nota A", "users": "250", "views": "340"}],
            "warnings": [],
        }

        canonical = canonicalize_monthly(raw)

        self.assertEqual(canonical["strategic_axes"], [{"label": "RCP", "value": 20.0}])
        self.assertEqual(canonical["channel_mix"], [{"label": "Mail", "value": 70.0}])
        self.assertEqual(canonical["format_mix"], [{"label": "Noticia propia", "value": 55.5}])
        self.assertEqual(canonical["top_push_by_interaction"][0]["name"], "Comms 1")
        self.assertEqual(canonical["top_push_by_interaction"][0]["interaction"], 10.5)
        self.assertEqual(canonical["top_push_by_open_rate"][0]["interaction"], 8.2)
        self.assertEqual(canonical["top_pull_notes"][0]["title"], "Nota A")
        self.assertEqual(canonical["top_pull_notes"][0]["unique_reads"], 250)
        self.assertEqual(canonical["top_pull_notes"][0]["total_reads"], 340)

    def test_canonicalize_normalizes_raw_extracted_shapes_for_analyzer_contract(self):
        raw = {
            "month": "2026-03",
            "parser": "test",
            "metrics": {
                "plan_daily_average": {"value": 2.1},
                "plan_total": {"value": 66},
                "site_notes_total": {"value": 17},
                "site_total_views": {"value": 5580},
                "site_average_views": {"value": 328},
                "mail_total": {"value": 61},
                "mail_open_rate": {"value": 77.53},
                "mail_interaction_rate": {"value": 9.17},
                "mail_interaction_rate_over_opened": {"value": 11.83},
            },
            "strategic_axes": [{"axis": "RCP", "count": 20}],
            "channel_mix": [{"channel": "Mail", "pct": 60}],
            "format_mix": [{"format": "Video", "pct": 40}],
            "top_push_interaction": [{"title": "Mail A", "clicks": 100, "ctr": 12.5, "open_rate": 80}],
            "top_push_open": [{"title": "Mail B", "clicks": 80, "ctr": 8.5, "open_rate": 90}],
            "top_pull_notes": [{"title": "Nota A", "users": 300, "views": 500}],
            "warnings": [],
        }

        canonical = canonicalize_monthly(raw)

        self.assertEqual(canonical["strategic_axes"], [{"label": "RCP", "value": 20}])
        self.assertEqual(canonical["channel_mix"], [{"label": "Mail", "value": 60}])
        self.assertEqual(canonical["format_mix"], [{"label": "Video", "value": 40}])

        self.assertEqual(canonical["top_push_by_interaction"][0]["name"], "Mail A")
        self.assertEqual(canonical["top_push_by_interaction"][0]["interaction"], 12.5)
        self.assertEqual(canonical["top_push_by_interaction"][0]["open_rate"], 80)

        self.assertEqual(canonical["top_pull_notes"][0]["unique_reads"], 300)
        self.assertEqual(canonical["top_pull_notes"][0]["total_reads"], 500)

    def test_validate_canonical_allows_missing_optional_anchors(self):
        canonical = {
            "month": "2026-03",
            "plan_total": 66,
            "site_notes_total": 17,
            "site_total_views": 5580,
            "mail_total": 61,
            "mail_open_rate": 77.53,
            "mail_interaction_rate": 9.17,
            "mail_interaction_rate_over_opened": 0,
            "extraction_warnings": [
                "missing_anchor:site_average_views:Promedio Vistas",
                "missing_anchor:mail_interaction_rate_over_opened:Tasa de interacción sobre mails abiertos",
            ],
        }

        validation = validate_canonical_monthly(canonical)

        self.assertTrue(validation["is_valid"])
        self.assertEqual(validation["errors"], [])
        self.assertTrue(any("KPIs secundarios" in warning for warning in validation["warnings"]))
        self.assertFalse(
            any(
                "mail_interaction_rate_over_opened es menor a 1%" in warning
                for warning in validation["warnings"]
            )
        )

    def test_extract_raw_monthly_pdf_accepts_percent_with_space_before_symbol(self):
        pages = [
            "Página 1\nMedia comunicaciones diarias\n2.1\nNº total de comunicaciones\n66",
            "Página 2\nTotal Páginas Vistas\n5,580\nNoticias Publicadas\n17\nPromedio Vistas\n328",
            (
                "Página 3\nMails enviados\n61\nTasa de apertura promedio\n77,53 %\n"
                "Tasa de interacción sobre mails enviados\n9,17 %\n"
                "Tasa de interacción sobre mails abiertos\n11,83 %"
            ),
        ]

        with patch("deterministic_pipeline._extract_pages_text", return_value=pages):
            raw = extract_raw_monthly_pdf("2026-03", Path("/tmp/fake.pdf"))

        self.assertEqual(raw["metrics"]["mail_open_rate"]["value"], 77.53)
        self.assertEqual(raw["metrics"]["mail_interaction_rate"]["value"], 9.17)
        self.assertEqual(raw["metrics"]["mail_interaction_rate_over_opened"]["value"], 11.83)

    def test_infer_month_key_requires_explicit_month_when_filename_has_no_yyyy_mm(self):
        with self.assertRaisesRegex(ValueError, "No pude inferir month_key.*Pasa month_key explícitamente"):
            infer_month_key_from_pdf_path(Path("/tmp/Dashboard_Marzo.pdf"))


if __name__ == "__main__":
    unittest.main()
