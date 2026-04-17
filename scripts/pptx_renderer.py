from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from analyzer import validate_report_json


def create_pptx(report: dict[str, Any], output_path: Path) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    safe_report = validate_report_json(report)

    renderer_path = Path(__file__).with_suffix('.js')
    with tempfile.NamedTemporaryFile('w', suffix='.json', delete=False, encoding='utf-8') as tmp:
        tmp.write(json.dumps(safe_report, ensure_ascii=False, indent=2))
        tmp_path = Path(tmp.name)

    try:
        subprocess.run(['node', str(renderer_path), str(tmp_path), str(output_path)], check=True, cwd=str(Path(__file__).resolve().parent.parent))
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass


def main() -> None:
    repo_root = Path(__file__).resolve().parent.parent
    if len(sys.argv) == 3:
        input_json = Path(sys.argv[1])
        output_pptx = Path(sys.argv[2])
    else:
        input_json = repo_root / 'sample_report_definitive.json'
        output_pptx = repo_root / 'sample_bbva_report_definitive.pptx'

    report = json.loads(input_json.read_text(encoding='utf-8'))
    create_pptx(report, output_pptx)
    print(f"PPTX generado: {output_pptx}")


if __name__ == '__main__':
    main()
