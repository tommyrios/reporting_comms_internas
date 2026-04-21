import os
import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from analyzer import BASE_STRUCTURE
from pptx_renderer import create_pptx


def _slide_texts(slide) -> str:
    chunks = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            chunks.append(shape.text_frame.text)
    return "\n".join(chunks)


class PptxRendererFrameTemplateTests(unittest.TestCase):
    def test_frame_mode_keeps_only_cover_and_closing_and_replaces_fecha(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            template_path = tmp_dir / "plantilla-bbva.pptx"
            output_path = tmp_dir / "report.pptx"

            template = Presentation()
            cover = template.slides.add_slide(template.slide_layouts[0])
            cover.shapes.title.text = "PORTADA_TEMPLATE"
            cover.placeholders[1].text = "FECHA"
            middle = template.slides.add_slide(template.slide_layouts[1])
            middle.shapes.title.text = "MIDDLE_TEMPLATE"
            closing = template.slides.add_slide(template.slide_layouts[1])
            closing.shapes.title.text = "CLOSING_TEMPLATE"
            template.save(str(template_path))

            report = dict(BASE_STRUCTURE)
            report["slide_1_cover"] = dict(BASE_STRUCTURE["slide_1_cover"])
            report["slide_1_cover"]["period"] = "Marzo 2026"

            original_mode = os.environ.get("PPTX_TEMPLATE_MODE")
            os.environ["PPTX_TEMPLATE_MODE"] = "frame"
            try:
                create_pptx(report, output_path, template_path=template_path)
            finally:
                if original_mode is None:
                    os.environ.pop("PPTX_TEMPLATE_MODE", None)
                else:
                    os.environ["PPTX_TEMPLATE_MODE"] = original_mode

            rendered = Presentation(str(output_path))
            all_text = "\n".join(_slide_texts(slide) for slide in rendered.slides)

            self.assertEqual(len(rendered.slides), 10)
            self.assertIn("PORTADA_TEMPLATE", _slide_texts(rendered.slides[0]))
            self.assertIn("Marzo 2026", _slide_texts(rendered.slides[0]))
            self.assertNotIn("FECHA", _slide_texts(rendered.slides[0]))
            self.assertIn("CLOSING_TEMPLATE", _slide_texts(rendered.slides[-1]))
            self.assertNotIn("MIDDLE_TEMPLATE", all_text)
            self.assertNotIn("PORTADA_TEMPLATE\nMIDDLE_TEMPLATE\nCLOSING_TEMPLATE", all_text)


if __name__ == "__main__":
    unittest.main()
