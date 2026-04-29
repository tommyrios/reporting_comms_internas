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
        modules = [
            {
                "key": "executive_summary",
                "title": "Resumen ejecutivo del período",
                "payload": {
                    "plan_total": 66,
                    "site_notes_total": 17,
                    "site_total_views": 6205,
                    "mail_total": 36,
                    "mail_open_rate": 77.53,
                    "mail_interaction_rate": 8.86,
                    "historical_note": "No comparable por alcance de fuente",
                    "takeaways": ["Takeaway 1", "Takeaway 2", "Takeaway 3"],
                },
            },
            {
                "key": "ranking_push",
                "title": "Ranking push",
                "payload": {
                    "available": False,
                    "by_interaction": [],
                    "by_open_rate": [],
                    "message": "Sin ranking",
                },
            },
        ]
        if include_events:
            modules.append(
                {
                    "key": "events",
                    "title": "Eventos del mes",
                    "payload": {
                        "events": [{"name": "Townhall", "participants": 100, "date": "2026-03-12"}],
                        "total_events": 1,
                        "total_participants": 100,
                        "message": "Detalle de eventos",
                    },
                }
            )

        return {
            "period": {"slug": "month_2026_03", "label": "Marzo 2026"},
            "kpis": {},
            "narrative": {},
            "quality_flags": {},
            "render_plan": {
                "period": {"slug": "month_2026_03", "label": "Marzo 2026"},
                "modules": modules,
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
            self.assertEqual(len(rendered.slides), 4)  # cover + 2 body + closing
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
            self.assertEqual(len(rendered.slides), 4)
            self.assertIn("Marzo 2026", _slide_texts(rendered.slides[0]))
            self.assertIn("Informe gestión", _slide_texts(rendered.slides[0]))
            self.assertIn("Gracias", _slide_texts(rendered.slides[-1]))

    def test_conditional_module_changes_slide_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            template_path = tmp_dir / "plantilla-bbva.pptx"
            out_without_events = tmp_dir / "without_events.pptx"
            out_with_events = tmp_dir / "with_events.pptx"

            template = Presentation()
            template.slides.add_slide(template.slide_layouts[0]).placeholders[1].text = "FECHA"
            template.slides.add_slide(template.slide_layouts[1]).shapes.title.text = "CLOSING"
            template.save(str(template_path))

            create_pptx(self._sample_report(include_events=False), out_without_events, template_mode="full", template_path=template_path)
            create_pptx(self._sample_report(include_events=True), out_with_events, template_mode="full", template_path=template_path)

            self.assertEqual(len(Presentation(str(out_without_events)).slides), 4)
            self.assertEqual(len(Presentation(str(out_with_events)).slides), 5)

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
