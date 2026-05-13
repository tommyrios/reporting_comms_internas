from __future__ import annotations

from pathlib import Path
from typing import Any

import fitz

PAGE_ANCHORS = {
    "planning": ["Herramienta de planificación", "Nº total de comunicaciones", "Acciones de comunicación"],
    "mailing": ["Herramienta de mailing", "Mails enviados", "Tasa de apertura promedio"],
    "contents": ["Contenidos publicados en site", "Noticias publicadas", "Top five - Notas más leídas"],
}

# crops más quirúrgicos que la versión anterior: solo asset visual, no tablas completas
CROP_CONFIG = {
    "planning": {
        "strategic_axes": {"x": 0.03, "y": 0.67, "w": 0.30, "h": 0.16},
        "channel_mix": {"x": 0.60, "y": 0.35, "w": 0.22, "h": 0.18},
        "internal_clients": {"x": 0.02, "y": 0.06, "w": 0.58, "h": 0.09},
    },
    "mailing": {
        "monthly_trend": {"x": 0.04, "y": 0.55, "w": 0.48, "h": 0.18},
    },
}


def find_page(doc: fitz.Document, anchors: list[str]) -> int | None:
    for i, page in enumerate(doc):
        text = " ".join((page.get_text("text") or "").split()).lower()
        if any(anchor.lower() in text for anchor in anchors):
            return i
    return None


def pct_rect(page: fitz.Page, box: dict[str, float]) -> fitz.Rect:
    rect = page.rect
    x0 = rect.x0 + rect.width * box["x"]
    y0 = rect.y0 + rect.height * box["y"]
    x1 = x0 + rect.width * box["w"]
    y1 = y0 + rect.height * box["h"]
    return fitz.Rect(x0, y0, x1, y1)


def render_crop(page: fitz.Page, crop_box: dict[str, float], out_path: Path, zoom: float = 3.0) -> str:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=pct_rect(page, crop_box), alpha=False)
    pix.save(str(out_path))
    return str(out_path)


def build_dashboard_crops(period_slug: str, scope_pdf_paths: dict[str, Path], output_dir: Path) -> dict[str, dict[str, dict[str, str]]]:
    root = output_dir / "dashboard_crops" / period_slug
    result: dict[str, dict[str, dict[str, str]]] = {}
    for scope, pdf in scope_pdf_paths.items():
        scope_result: dict[str, dict[str, str]] = {}
        with fitz.open(str(pdf)) as doc:
            for module, crops in CROP_CONFIG.items():
                idx = find_page(doc, PAGE_ANCHORS[module])
                scope_result[module] = {}
                if idx is None:
                    continue
                page = doc[idx]
                for name, box in crops.items():
                    try:
                        out = root / scope / module / f"{name}.png"
                        scope_result[module][name] = render_crop(page, box, out)
                    except Exception:
                        scope_result[module][name] = ""
        result[scope] = scope_result
    return result
