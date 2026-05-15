from __future__ import annotations

import logging
from pathlib import Path

import fitz

logger = logging.getLogger(__name__)

# Convención real del PDF exportado desde el dashboard:
#   página 0 = Herramienta de planificación
#   página 1 = Contenidos publicados en el site
#   página 2 = Herramienta de mailing
#
# Las cajas están en coordenadas absolutas del PDF (no porcentajes), tomadas
# sobre el export estándar del dashboard. Esto es intencional: si el layout es
# estable, las coordenadas fijas son más confiables que buscar por texto/heurística.
PAGE_INDEX_BY_MODULE: dict[str, int] = {
    "planning": 0,
    "contents": 1,
    "mailing": 2,
}

PAGE_ANCHORS: dict[str, list[str]] = {
    "planning": ["Nº total de comunicaciones", "Listado completo de comunicaciones"],
    "contents": ["Contenidos publicados en site", "Noticias publicadas", "Top five - Notas más leídas"],
    "mailing": ["Mails enviados", "Tasa de apertura promedio", "Tasa de interacción sobre mails enviados"],
}

# Coordenadas provistas para el dashboard Q1.
# Formato: x, y, w, h en unidades PDF absolutas.
# Mantenemos nombres compatibles con pptx_renderer.py:
#   planning.strategic_axes / channel_mix / internal_clients
#   mailing.monthly_trend
# Y agregamos crops nuevos para futuros usos/debug.
CROP_CONFIG: dict[str, dict[str, dict[str, float]]] = {
    "contents": {
        "site_notes_total_card": {"x": 100, "y": 300, "w": 205, "h": 80},
        "site_total_views_card": {"x": 335, "y": 300, "w": 205, "h": 80},
        "site_average_views_card": {"x": 570, "y": 300, "w": 205, "h": 80},
        "top_notes_uu": {"x": 25, "y": 2875, "w": 845, "h": 160},
        "top_notes_tgm": {"x": 25, "y": 3045, "w": 845, "h": 160},
    },
    "mailing": {
        "mail_total_card": {"x": 35, "y": 270, "w": 130, "h": 80},
        "mail_open_rate_card": {"x": 180, "y": 270, "w": 205, "h": 80},
        "mail_interaction_sent_card": {"x": 395, "y": 270, "w": 235, "h": 80},
        "mail_interaction_opened_card": {"x": 640, "y": 270, "w": 235, "h": 80},
        "monthly_trend": {"x": 20, "y": 1190, "w": 860, "h": 270},
        "top_open_rate": {"x": 20, "y": 1475, "w": 420, "h": 160},
        "top_interaction": {"x": 465, "y": 1475, "w": 420, "h": 160},
    },
    "planning": {
        "plan_total_card": {"x": 690, "y": 300, "w": 150, "h": 50},
        "strategic_axes": {"x": 15, "y": 770, "w": 400, "h": 320},
        "internal_clients": {"x": 15, "y": 1110, "w": 875, "h": 240},
        "channel_mix": {"x": 15, "y": 1375, "w": 400, "h": 365},
    },
}


def _normalized_text(page: fitz.Page) -> str:
    return " ".join((page.get_text("text") or "").split()).lower()


def _page_matches_module(page: fitz.Page, module: str) -> bool:
    text = _normalized_text(page)
    anchors = PAGE_ANCHORS.get(module, [])
    return not anchors or any(anchor.lower() in text for anchor in anchors)


def find_page(doc: fitz.Document, module: str) -> int | None:
    """Devuelve la página del módulo.

    Primero intenta la posición fija esperada. Si el export cambiara de orden,
    cae a búsqueda por anclas para evitar capturar una página equivocada.
    """
    fixed_idx = PAGE_INDEX_BY_MODULE.get(module)
    if fixed_idx is not None and fixed_idx < len(doc) and _page_matches_module(doc[fixed_idx], module):
        return fixed_idx

    anchors = PAGE_ANCHORS.get(module, [])
    for i, page in enumerate(doc):
        text = _normalized_text(page)
        if any(anchor.lower() in text for anchor in anchors):
            logger.warning(
                "event=dashboard_crop_page_fallback module=%s expected_page=%s resolved_page=%s",
                module,
                fixed_idx,
                i,
            )
            return i
    return fixed_idx if fixed_idx is not None and fixed_idx < len(doc) else None


def abs_rect(page: fitz.Page, box: dict[str, float]) -> fitz.Rect:
    """Convierte una caja absoluta x/y/w/h en fitz.Rect, recortándola al tamaño de página."""
    page_rect = page.rect
    x0 = max(page_rect.x0, float(box["x"]))
    y0 = max(page_rect.y0, float(box["y"]))
    x1 = min(page_rect.x1, float(box["x"]) + float(box["w"]))
    y1 = min(page_rect.y1, float(box["y"]) + float(box["h"]))
    rect = fitz.Rect(x0, y0, x1, y1)
    if rect.is_empty or rect.width <= 1 or rect.height <= 1:
        raise ValueError(f"Crop inválido fuera de página: box={box}, page={page.rect}")
    return rect


def render_crop(page: fitz.Page, crop_box: dict[str, float], out_path: Path, zoom: float = 5.0) -> str:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=abs_rect(page, crop_box), alpha=False)
    pix.set_dpi(300, 300)
    pix.save(str(out_path))
    return str(out_path)


def render_debug_page(page: fitz.Page, module: str, crops: dict[str, dict[str, float]], out_path: Path) -> str:
    """Genera una página completa con rectángulos rojos para auditar coordenadas."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open()
    copied = doc.new_page(width=page.rect.width, height=page.rect.height)
    copied.show_pdf_page(copied.rect, page.parent, page.number)
    for name, box in crops.items():
        rect = abs_rect(copied, box)
        copied.draw_rect(rect, color=(1, 0, 0), width=3)
        copied.insert_text((rect.x0 + 4, rect.y0 + 14), name, fontsize=12, color=(1, 0, 0))
    pix = copied.get_pixmap(matrix=fitz.Matrix(0.35, 0.35), alpha=False)
    pix.save(str(out_path))
    doc.close()
    return str(out_path)


def build_dashboard_crops(
    period_slug: str,
    scope_pdf_paths: dict[str, Path],
    output_dir: Path,
    *,
    debug: bool = True,
) -> dict[str, dict[str, dict[str, str]]]:
    root = output_dir / "dashboard_crops" / period_slug
    result: dict[str, dict[str, dict[str, str]]] = {}

    for scope, pdf in scope_pdf_paths.items():
        scope_result: dict[str, dict[str, str]] = {}
        with fitz.open(str(pdf)) as doc:
            for module, crops in CROP_CONFIG.items():
                idx = find_page(doc, module)
                scope_result[module] = {}
                if idx is None:
                    logger.warning("event=dashboard_crop_page_missing scope=%s module=%s pdf=%s", scope, module, pdf)
                    continue

                page = doc[idx]
                if not _page_matches_module(page, module):
                    logger.warning(
                        "event=dashboard_crop_anchor_mismatch scope=%s module=%s page=%s pdf=%s",
                        scope,
                        module,
                        idx,
                        pdf,
                    )
                    continue

                if debug:
                    try:
                        render_debug_page(page, module, crops, root / scope / module / "__debug_boxes.png")
                    except Exception as exc:
                        logger.warning("event=dashboard_crop_debug_failed scope=%s module=%s reason=%s", scope, module, exc)

                for name, box in crops.items():
                    try:
                        out = root / scope / module / f"{name}.png"
                        scope_result[module][name] = render_crop(page, box, out)
                    except Exception as exc:
                        logger.warning(
                            "event=dashboard_crop_failed scope=%s module=%s crop=%s reason=%s",
                            scope,
                            module,
                            name,
                            exc,
                        )
                        scope_result[module][name] = ""
        result[scope] = scope_result
    return result
