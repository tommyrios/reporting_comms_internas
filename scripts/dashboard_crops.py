from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import fitz

logger = logging.getLogger(__name__)

PAGE_ANCHORS = {
    "planning": [
        "Herramienta de planificación",
        "Nº total de comunicaciones",
        "N° total de comunicaciones",
        "¿Qué canales y formatos se han utilizado?",
    ],
    "mailing": [
        "Herramienta de mailing",
        "Mails enviados",
        "Tasa de apertura promedio",
    ],
    "contents": [
        "Contenidos publicados en site",
        "Noticias publicadas",
        "Top five - Notas más leídas",
    ],
}

# Coordenadas relativas al tamaño de página (x/y/w/h entre 0 y 1).
# Estos recortes están pensados para los PDFs exportados desde Looker Studio:
# - planificación: bloque de áreas superior, mix de canales y ejes de la parte inferior
# - mailing: gráfico de tendencia mensual
CROP_CONFIG: dict[str, dict[str, dict[str, float]]] = {
    "planning": {
        "internal_clients": {"x": 0.02, "y": 0.02, "w": 0.58, "h": 0.19},
        "channel_mix": {"x": 0.50, "y": 0.45, "w": 0.47, "h": 0.22},
        "strategic_axes": {"x": 0.02, "y": 0.72, "w": 0.52, "h": 0.25},
    },
    "mailing": {
        "monthly_trend": {"x": 0.03, "y": 0.57, "w": 0.52, "h": 0.22},
    },
}


def _normalize_text(text: str) -> str:
    return " ".join((text or "").replace("\xa0", " ").split()).lower()


def _find_page_index(doc: fitz.Document, anchors: list[str]) -> int | None:
    normalized_anchors = [_normalize_text(anchor) for anchor in anchors]
    for index, page in enumerate(doc):
        text = _normalize_text(page.get_text("text") or "")
        if any(anchor in text for anchor in normalized_anchors):
            return index
    return None


def _pct_rect(page: fitz.Page, crop: dict[str, float]) -> fitz.Rect:
    rect = page.rect
    x0 = rect.x0 + rect.width * crop["x"]
    y0 = rect.y0 + rect.height * crop["y"]
    x1 = x0 + rect.width * crop["w"]
    y1 = y0 + rect.height * crop["h"]
    return fitz.Rect(x0, y0, x1, y1)


def _render_crop(page: fitz.Page, crop: dict[str, float], output_path: Path, zoom: float = 3.0) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=_pct_rect(page, crop), alpha=False)
    pix.save(str(output_path))


def build_dashboard_crops(
    period_slug: str,
    scope_pdf_paths: dict[str, Path],
    output_dir: Path,
) -> dict[str, dict[str, dict[str, str]]]:
    """Genera PNGs recortados de gráficos del dashboard para cada scope.

    La función es tolerante: si falla un crop, deja string vacío para que el
    renderer inserte placeholder, sin romper el reporte completo.
    """
    crops_root = Path(output_dir) / "dashboard_crops" / period_slug
    result: dict[str, dict[str, dict[str, str]]] = {}

    for scope, pdf_path in scope_pdf_paths.items():
        scope_result: dict[str, dict[str, str]] = {}
        try:
            with fitz.open(str(pdf_path)) as doc:
                for module, module_crops in CROP_CONFIG.items():
                    page_index = _find_page_index(doc, PAGE_ANCHORS.get(module, []))
                    scope_result[module] = {}
                    if page_index is None:
                        logger.warning("event=dashboard_crop_page_missing scope=%s module=%s pdf=%s", scope, module, pdf_path)
                        continue
                    page = doc[page_index]
                    for crop_name, crop_box in module_crops.items():
                        output_path = crops_root / scope / module / f"{crop_name}.png"
                        try:
                            _render_crop(page, crop_box, output_path)
                            scope_result[module][crop_name] = str(output_path)
                        except Exception as exc:  # noqa: BLE001
                            logger.warning(
                                "event=dashboard_crop_failed scope=%s module=%s crop=%s pdf=%s reason=%s",
                                scope,
                                module,
                                crop_name,
                                pdf_path,
                                exc,
                            )
                            scope_result[module][crop_name] = ""
        except Exception as exc:  # noqa: BLE001
            logger.warning("event=dashboard_crop_pdf_failed scope=%s pdf=%s reason=%s", scope, pdf_path, exc)
        result[scope] = scope_result

    return result
