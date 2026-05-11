import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx import Presentation

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from pptx_renderer import _render_body_with_js, create_pptx


def _slide_texts(slide) -> str:
    chunks = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            chunks.append(shape.text_frame.text)
    return "\n".join(chunks)


class PptxRendererFrameTemplateTests(unittest.TestCase):
    def _sample_report(self, include_events: bool = False):
        def _scope(label: str):
            return {
                "scope": label.lower(),
                "scope_label": label,
                "plan_total": 66,
                "strategic_axes": [{"label": "Negocio", "value": 55}, {"label": "Personas", "value": 45}],
                "channel_mix": [{"label": "Mail", "value": 60}, {"label": "Intranet", "value": 40}],
                "internal_clients": [{"label": "Talento", "value": 40}, {"label": "Negocio", "value": 60}],
                "mail_total": 36,
                "mail_send_total": 36,
                "mail_unique_total": 20,
                "mail_open_rate": 77.53,
                "mail_interaction_rate": 8.86,
                "mail_interaction_rate_over_opened": 11.2,
                "top_push_by_open_rate": [{"title": "Campaña A", "open_rate": 87.2}],
                "top_push_by_interaction": [{"title": "Campaña B", "interaction": 12.4}],
                "site_notes_total": 17,
                "site_total_views": 6205,
                "site_average_views": 365,
                "top_pull_notes": [{"title": "Nota 1", "unique_reads": 220, "total_reads": 320}],
                "top_pull_notes_tgm": [{"title": "Nota TGM", "unique_reads": 180, "total_reads": 250}],
            }

        return {
            "period": {"slug": "month_2026_03", "label": "Marzo 2026"},
            "kpis": {
                "scopes": {
                    "argentina": _scope("Argentina"),
                    "holding": _scope("Holding"),
                    "combined": _scope("Argentina + Holding"),
                }
            },
            "narrative": {},
            "quality_flags": {},
            "render_plan": {
                "period": {"slug": "month_2026_03", "label": "Marzo 2026"},
                "modules": [{"key": "planning_comparison", "title": "Planificación | Argentina vs Holding", "payload": {}}],
            },
        }

    def test_frame_mode_uses_template_cover_body_and_closing(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            template_path = tmp_dir / "plantilla-bbva.pptx"
            output_path = tmp_dir / "report.pptx"

            template = Presentation()
            cover = template.slides.add_slide(template.slide_layouts[0])
            cover.shapes.title.text = "PORTADA_TEMPLATE"
            cover.placeholders[1].text = "FECHA"
            closing = template.slides.add_slide(template.slide_layouts[1])
            closing.shapes.title.text = "CLOSING_TEMPLATE"
            template.save(str(template_path))

            create_pptx(self._sample_report(include_events=False), output_path, template_mode="frame", template_path=template_path)

            rendered = Presentation(str(output_path))
            all_text = "\n".join(_slide_texts(slide) for slide in rendered.slides)
            self.assertEqual(len(rendered.slides), 8)  # cover + 6 body + closing
            self.assertIn("PORTADA_TEMPLATE", _slide_texts(rendered.slides[0]))
            self.assertIn("Marzo 2026", _slide_texts(rendered.slides[0]))
            self.assertNotIn("FECHA", _slide_texts(rendered.slides[0]))
            self.assertIn("CLOSING_TEMPLATE", _slide_texts(rendered.slides[-1]))
            self.assertNotIn("slide_", all_text.lower())

            for idx in range(1, len(rendered.slides) - 1):
                self.assertTrue(_slide_texts(rendered.slides[idx]).strip())

    def test_default_mode_uses_js_full_cover_and_closing(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = Path(tmp) / "report.pptx"
            create_pptx(self._sample_report(include_events=False), output_path)

            rendered = Presentation(str(output_path))
            self.assertEqual(len(rendered.slides), 8)
            self.assertIn("Marzo 2026", _slide_texts(rendered.slides[0]))
            self.assertIn("Comunicaciones Internas", _slide_texts(rendered.slides[0]))
            self.assertIn("Gestión", _slide_texts(rendered.slides[0]))
            self.assertIn("Comunicaciones Internas", _slide_texts(rendered.slides[-1]))

    def test_slide_count_constant_when_tables_empty(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            template_path = tmp_dir / "plantilla-bbva.pptx"
            out_without_events = tmp_dir / "without_events.pptx"
            out_with_events = tmp_dir / "with_events.pptx"

            template = Presentation()
            template.slides.add_slide(template.slide_layouts[0]).placeholders[1].text = "FECHA"
            template.slides.add_slide(template.slide_layouts[1]).shapes.title.text = "CLOSING"
            template.save(str(template_path))

            empty_report = self._sample_report(include_events=False)
            empty_report["kpis"]["scopes"]["argentina"]["top_push_by_open_rate"] = []
            empty_report["kpis"]["scopes"]["argentina"]["top_push_by_interaction"] = []
            empty_report["kpis"]["scopes"]["argentina"]["top_pull_notes"] = []
            empty_report["kpis"]["scopes"]["argentina"]["top_pull_notes_tgm"] = []

            create_pptx(empty_report, out_without_events, template_mode="full", template_path=template_path)
            create_pptx(self._sample_report(include_events=False), out_with_events, template_mode="full", template_path=template_path)

            self.assertEqual(len(Presentation(str(out_without_events)).slides), 8)
            self.assertEqual(len(Presentation(str(out_with_events)).slides), 8)

    def test_observations_placeholder_present(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = Path(tmp) / "report.pptx"
            create_pptx(self._sample_report(include_events=False), output_path)
            rendered = Presentation(str(output_path))
            for idx in range(1, 7):
                self.assertIn("Observaciones del manager", _slide_texts(rendered.slides[idx]))

    def test_no_executive_phrases_rendered(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = Path(tmp) / "report.pptx"
            create_pptx(self._sample_report(include_events=False), output_path)
            rendered = Presentation(str(output_path))
            all_text = "\n".join(_slide_texts(slide) for slide in rendered.slides)
            for phrase in [
                "Resumen ejecutivo del período",
                "Lectura ejecutiva",
                "Plan de mejora",
                "Quick wins",
                "Experimentos",
            ]:
                self.assertNotIn(phrase, all_text)

    def test_missing_chart_uses_placeholder(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = Path(tmp) / "report.pptx"
            report = self._sample_report(include_events=False)
            report["kpis"]["scopes"]["argentina"]["strategic_axes"] = []
            report["kpis"]["scopes"]["argentina"]["channel_mix"] = []
            report["kpis"]["scopes"]["argentina"]["internal_clients"] = []
            create_pptx(report, output_path)
            rendered = Presentation(str(output_path))
            all_text = "\n".join(_slide_texts(slide) for slide in rendered.slides)
            self.assertIn("Gráfico no disponible", all_text)

    def test_js_renderer_is_invoked_in_body_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = Path(tmp) / "body.pptx"
            report = self._sample_report()

            def _fake_run(*args, **kwargs):
                Presentation().save(str(output_path))
                return None

            with patch("pptx_renderer.subprocess.run", side_effect=_fake_run) as run_mock:
                _render_body_with_js(report, output_path)

            cmd = run_mock.call_args.args[0]
            self.assertIn("--mode=body", cmd)


if __name__ == "__main__":
    unittest.main()
