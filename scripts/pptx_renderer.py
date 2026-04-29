from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from analyzer import validate_report_json
from config import ASSETS_DIR
from deck_assembler import assemble_deck

DEFAULT_TEMPLATE_PATH = ASSETS_DIR / "plantilla-bbva.pptx"
DEFAULT_TEMPLATE_MODE = "full"


def _period_label(report: dict[str, Any]) -> str:
    render_plan = report.get("render_plan") if isinstance(report.get("render_plan"), dict) else {}
    period = render_plan.get("period") if isinstance(render_plan.get("period"), dict) else {}
    if period.get("label"):
        return str(period["label"])
    if isinstance(report.get("period"), dict) and report["period"].get("label"):
        return str(report["period"]["label"])
    return "-"


def _render_body_with_js(report: dict[str, Any], output_path: Path, mode: str = "body") -> None:
    renderer_path = Path(__file__).with_suffix(".js")
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as tmp:
        tmp.write(json.dumps(report, ensure_ascii=False, indent=2))
        input_json_path = Path(tmp.name)

    try:
        subprocess.run(
            ["node", str(renderer_path), str(input_json_path), str(output_path), f"--mode={mode}"],
            check=True,
            cwd=str(Path(__file__).resolve().parent.parent),
        )
    finally:
        input_json_path.unlink(missing_ok=True)


def create_pptx(
    report: dict[str, Any],
    output_path: Path,
    template_mode: str = DEFAULT_TEMPLATE_MODE,
    template_path: Path | None = None,
) -> None:
    output_path = Path(output_path)
    safe_report = validate_report_json(report)
    template = Path(template_path) if template_path else DEFAULT_TEMPLATE_PATH
    mode = (template_mode or DEFAULT_TEMPLATE_MODE).strip().lower()

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        body_path = Path(tmp) / "body.pptx"
        _render_body_with_js(safe_report, body_path, mode="body" if mode == "frame" else "full")

        if mode == "frame" and template.exists():
            assemble_deck(template, body_path, output_path, _period_label(safe_report))
            return

        body_bytes = body_path.read_bytes()
        output_path.write_bytes(body_bytes)


def main() -> None:
    repo_root = Path(__file__).resolve().parent.parent
    if len(sys.argv) == 3:
        input_json = Path(sys.argv[1])
        output_pptx = Path(sys.argv[2])
    else:
        input_json = repo_root / "templates" / "sample_report_definitive.json"
        output_pptx = repo_root / "sample_bbva_report_definitive.pptx"

    report = json.loads(input_json.read_text(encoding="utf-8"))
    create_pptx(report, output_pptx)
    print(f"PPTX generado: {output_pptx}")


if __name__ == "__main__":
    main()
