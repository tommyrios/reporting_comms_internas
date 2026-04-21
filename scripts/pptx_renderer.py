from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

from analyzer import validate_report_json
from config import ASSETS_DIR


DEFAULT_TEMPLATE_PATH = ASSETS_DIR / "plantilla-bbva.pptx"
DEFAULT_TEMPLATE_MODE = "frame"


def _normalize(value: str) -> str:
    normalized_text = "".join(ch for ch in unicodedata.normalize("NFKD", value) if not unicodedata.combining(ch))
    return normalized_text.strip().lower()


def _find_layout(prs: Presentation, aliases: list[str], fallback_index: int = 0):
    normalized_aliases = [_normalize(alias) for alias in aliases]
    for layout in prs.slide_layouts:
        name = _normalize(layout.name or "")
        if any(alias == name or alias in name for alias in normalized_aliases):
            return layout
    return prs.slide_layouts[fallback_index]


def _title_placeholder(slide):
    if slide.shapes.title is not None:
        return slide.shapes.title
    for ph in slide.placeholders:
        if ph.placeholder_format.type in {PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE}:
            return ph
    return None


def _body_placeholder(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.type in {
            PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT,
            PP_PLACEHOLDER.SUBTITLE,
        }:
            return ph
    return None


def _render_cover(slide, cover: dict[str, Any]) -> None:
    title = _title_placeholder(slide)
    if title is not None:
        title.text = str(cover.get("area") or "Comunicaciones Internas")

    subtitle = _body_placeholder(slide)
    if subtitle is not None:
        subtitle.text = f"{cover.get('period') or '-'}\n{cover.get('subtitle') or 'Informe de gestión'}"


def _replace_token_in_slide(slide, token: str, replacement: str) -> None:
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text_frame = shape.text_frame
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if token in run.text:
                    run.text = run.text.replace(token, replacement)


def _remove_slide(prs: Presentation, slide_index: int) -> None:
    del prs.slides._sldIdLst[slide_index]


def _move_slide_to_end(prs: Presentation, slide_id_element) -> None:
    prs.slides._sldIdLst.remove(slide_id_element)
    prs.slides._sldIdLst.append(slide_id_element)


def _as_bullets(value: Any) -> list[str]:
    if isinstance(value, list):
        lines = []
        for item in value:
            if isinstance(item, dict):
                text = ", ".join(f"{k}: {v}" for k, v in item.items()).strip()
            else:
                text = str(item).strip()
            if text:
                lines.append(text)
        return lines
    if isinstance(value, dict):
        return [f"{k}: {v}" for k, v in value.items()]
    if value in (None, ""):
        return []
    return [str(value)]


def _render_content_slide(slide, title: str, body_sections: list[tuple[str, Any]]) -> None:
    title_box = _title_placeholder(slide)
    if title_box is not None:
        title_box.text = title

    body = _body_placeholder(slide)
    if body is None:
        return

    lines: list[str] = []
    for section_name, section_value in body_sections:
        bullets = _as_bullets(section_value)
        if not bullets:
            continue
        lines.append(f"{section_name}:")
        lines.extend([f"• {bullet}" for bullet in bullets])
        lines.append("")
    body.text = "\n".join(lines).strip() or "-"


def _build_content_map(report: dict[str, Any]) -> list[tuple[str, Any]]:
    return [
        ("slide_2_overview", report.get("slide_2_overview", {})),
        ("slide_3_plan", report.get("slide_3_plan", {})),
        ("slide_4_strategy", report.get("slide_4_strategy", {})),
        ("slide_5_push_ranking", report.get("slide_5_push_ranking", {})),
        ("slide_6_pull_performance", report.get("slide_6_pull_performance", {})),
        ("slide_7_hitos", report.get("slide_7_hitos", [])),
        ("slide_8_events", report.get("slide_8_events", {})),
        ("slide_9_closure", report.get("slide_9_closure", {})),
    ]


def _render_template_append(report: dict[str, Any], output_path: Path, template_path: Path) -> None:
    prs = Presentation(str(template_path))
    portada_layout = _find_layout(prs, ["Portada", "Cover", "Title Slide"], fallback_index=0)
    contenido_layout = _find_layout(
        prs,
        ["Título y Contenido", "Titulo y Contenido", "Title and Content"],
        fallback_index=1 if len(prs.slide_layouts) > 1 else 0,
    )

    cover_slide = prs.slides.add_slide(portada_layout)
    _render_cover(cover_slide, report.get("slide_1_cover", {}))

    for slide_key, payload in _build_content_map(report):
        slide = prs.slides.add_slide(contenido_layout)
        if isinstance(payload, dict):
            title = str(payload.get("title") or payload.get("headline") or slide_key)
            body_sections = [(key, value) for key, value in payload.items() if key not in {"title", "headline"}]
        else:
            title = slide_key
            body_sections = [("contenido", payload)]
        _render_content_slide(slide, title, body_sections)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))


def _render_template_frame(report: dict[str, Any], output_path: Path, template_path: Path) -> None:
    prs = Presentation(str(template_path))
    initial_slide_count = len(prs.slides)
    closing_slide_id = prs.slides._sldIdLst[initial_slide_count - 1] if initial_slide_count else None
    contenido_layout = _find_layout(
        prs,
        ["Título y Contenido", "Titulo y Contenido", "Title and Content"],
        fallback_index=1 if len(prs.slide_layouts) > 1 else 0,
    )

    cover_slide = prs.slides[0]
    period = str(report.get("slide_1_cover", {}).get("period") or "-")
    _replace_token_in_slide(cover_slide, "FECHA", period)

    for slide_key, payload in _build_content_map(report):
        slide = prs.slides.add_slide(contenido_layout)
        if isinstance(payload, dict):
            title = str(payload.get("title") or payload.get("headline") or slide_key)
            body_sections = [(key, value) for key, value in payload.items() if key not in {"title", "headline"}]
        else:
            title = slide_key
            body_sections = [("contenido", payload)]
        _render_content_slide(slide, title, body_sections)

    for idx in range(max(initial_slide_count - 2, 0), 0, -1):
        _remove_slide(prs, idx)

    if closing_slide_id is not None and prs.slides._sldIdLst[-1] is not closing_slide_id:
        _move_slide_to_end(prs, closing_slide_id)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))


def _render_with_node_fallback(report: dict[str, Any], output_path: Path) -> None:
    renderer_path = Path(__file__).with_suffix(".js")
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as tmp:
        tmp.write(json.dumps(report, ensure_ascii=False, indent=2))
        tmp_path = Path(tmp.name)

    try:
        subprocess.run(
            ["node", str(renderer_path), str(tmp_path), str(output_path)],
            check=True,
            cwd=str(Path(__file__).resolve().parent.parent),
        )
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass


def create_pptx(report: dict[str, Any], output_path: Path, template_path: Path | None = None) -> None:
    output_path = Path(output_path)
    safe_report = validate_report_json(report)
    template = Path(template_path) if template_path else DEFAULT_TEMPLATE_PATH

    if template.exists():
        template_mode = (os.getenv("PPTX_TEMPLATE_MODE") or DEFAULT_TEMPLATE_MODE).strip().lower()
        if template_mode == "frame":
            _render_template_frame(safe_report, output_path, template)
        else:
            _render_template_append(safe_report, output_path, template)
        return

    _render_with_node_fallback(safe_report, output_path)


def main() -> None:
    repo_root = Path(__file__).resolve().parent.parent
    if len(sys.argv) == 3:
        input_json = Path(sys.argv[1])
        output_pptx = Path(sys.argv[2])
    else:
        input_json = repo_root / "sample_report_definitive.json"
        output_pptx = repo_root / "sample_bbva_report_definitive.pptx"

    report = json.loads(input_json.read_text(encoding="utf-8"))
    create_pptx(report, output_pptx)
    print(f"PPTX generado: {output_pptx}")


if __name__ == "__main__":
    main()
