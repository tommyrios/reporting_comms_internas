from __future__ import annotations

from pathlib import Path

import fitz

# Los dashboards trimestrales exportados desde Looker/Data Studio tienen layout fijo:
#   página 0 = planificación
#   página 1 = contenidos/site
#   página 2 = mailing
# Por eso evitamos detectar páginas por texto: esa heurística mezclaba módulos cuando
# los textos aparecían concatenados o repetidos.
MODULE_PAGES = {
    "planning": 0,
    "contents": 1,
    "mailing": 2,
}

# Coordenadas relativas a cada página del PDF: x, y, w, h en rango 0..1.
# Mantener acá el contrato visual dashboard -> PPTX. Si cambia el layout del dashboard,
# solo se ajustan estos boxes.
CROP_CONFIG = {
    "planning": {
        # Gráfico inferior: Distribución por eje estratégico.
        "strategic_axes": {"x": 0.02, "y": 0.705, "w": 0.62, "h": 0.245},
        # Donut de canales.
        "channel_mix": {"x": 0.58, "y": 0.295, "w": 0.36, "h": 0.265},
        # Barras horizontales de áreas solicitantes.
        "internal_clients": {"x": 0.02, "y": 0.055, "w": 0.78, "h": 0.175},
    },
    "contents": {
        # Tráfico generado por noticias sobre cada prioridad.
        "priority_traffic": {"x": 0.47, "y": 0.245, "w": 0.47, "h": 0.175},
        # Volumen de contenidos publicados por prioridad.
        "priority_volume": {"x": 0.02, "y": 0.565, "w": 0.56, "h": 0.155},
        # Top five - Notas más leídas (uu).
        "top_notes_uu": {"x": 0.02, "y": 0.720, "w": 0.90, "h": 0.115},
        # Top five - Notas más leídas (colectivo TGM).
        "top_notes_tgm": {"x": 0.02, "y": 0.845, "w": 0.90, "h": 0.115},
    },
    "mailing": {
        # Tendencia mensual de envíos y aperturas.
        "monthly_trend": {"x": 0.04, "y": 0.545, "w": 0.54, "h": 0.185},
        # Ranking de apertura.
        "top_open_rate": {"x": 0.60, "y": 0.565, "w": 0.36, "h": 0.135},
        # Ranking de interacción.
        "top_interaction": {"x": 0.60, "y": 0.735, "w": 0.36, "h": 0.135},
    },
}


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


def build_dashboard_crops(
    period_slug: str,
    scope_pdf_paths: dict[str, Path],
    output_dir: Path,
) -> dict[str, dict[str, dict[str, str]]]:
    root = output_dir / "dashboard_crops" / period_slug
    result: dict[str, dict[str, dict[str, str]]] = {}

    for scope, pdf in scope_pdf_paths.items():
        scope_result: dict[str, dict[str, str]] = {}
        with fitz.open(str(pdf)) as doc:
            for module, crops in CROP_CONFIG.items():
                page_idx = MODULE_PAGES[module]
                scope_result[module] = {}
                if page_idx >= len(doc):
                    continue
                page = doc[page_idx]
                for name, box in crops.items():
                    out = root / scope / module / f"{name}.png"
                    try:
                        scope_result[module][name] = render_crop(page, box, out)
                    except Exception:
                        scope_result[module][name] = ""
        result[scope] = scope_result

    return result
