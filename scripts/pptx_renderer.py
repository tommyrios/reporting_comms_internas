from __future__ import annotations

import json
import math
import sys
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.util import Inches, Pt

from analyzer import validate_report_json
from config import ASSETS_DIR

SLIDE_W = 13.333
SLIDE_H = 7.5

COLORS = {
    "bg": RGBColor(245, 246, 248),
    "white": RGBColor(255, 255, 255),
    "bbva_blue": RGBColor(0, 45, 156),
    "bbva_dark": RGBColor(8, 28, 87),
    "bbva_mid": RGBColor(83, 102, 141),
    "text": RGBColor(28, 34, 45),
    "muted": RGBColor(107, 116, 131),
    "border": RGBColor(224, 228, 235),
    "placeholder": RGBColor(238, 240, 244),
    "obs": RGBColor(234, 236, 241),
}

COVER_PATH = ASSETS_DIR / "reference" / "boceto_cover.png"
BBVA_LOGO_BLUE = ASSETS_DIR / "brand" / "bbva_logo_blue.png"
BBVA_LOGO_WHITE = ASSETS_DIR / "brand" / "bbva_logo_white.png"


def _in(v: float):
    return Inches(v)


def _safe_text(v: Any, default: str = "-") -> str:
    if v is None:
        return default
    text = str(v)
    if not text.strip():
        return default
    replacements = {
        "�": "",
        "\u00f3": "ó",
        "\u00e1": "á",
        "\u00e9": "é",
        "\u00ed": "í",
        "\u00fa": "ú",
        "\u00f1": "ñ",
        "\u2026": "…",
    }
    for k, val in replacements.items():
        text = text.replace(k, val)
    text = text.replace("\n", " ")
    text = " ".join(text.split())
    return text or default


def _parse_num(v: Any) -> float:
    if v in (None, "", "-"):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    text = str(v).strip()
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except Exception:
        digits = "".join(ch for ch in text if ch.isdigit() or ch in ".-")
        try:
            return float(digits)
        except Exception:
            return 0.0


def _fmt_int(v: Any) -> str:
    return f"{int(round(_parse_num(v))):,}".replace(",", ".")


def _fmt_pct(v: Any) -> str:
    return f"{_parse_num(v):.2f}%".replace(".", ",")


def _clip(text: Any, max_len: int) -> str:
    t = _safe_text(text, "")
    if len(t) <= max_len:
        return t
    cut = t[: max_len - 1].rsplit(" ", 1)[0].strip()
    return (cut or t[: max_len - 1].strip()) + "…"


def _scope_bundle(report: dict[str, Any]) -> dict[str, dict[str, Any]]:
    scopes = {}
    if isinstance(report.get("kpis"), dict) and isinstance(report["kpis"].get("scopes"), dict):
        scopes = report["kpis"]["scopes"]
    elif isinstance(report.get("scopes"), dict):
        scopes = report["scopes"]
    # normalize nested objects
    normalized = {}
    for key in ("argentina", "holding", "combined"):
        data = scopes.get(key) if isinstance(scopes, dict) else None
        normalized[key] = data if isinstance(data, dict) else {}
    return normalized


def _period_label(report: dict[str, Any]) -> str:
    period = report.get("period") if isinstance(report.get("period"), dict) else {}
    return _safe_text(period.get("label"), "Q1 2026 (ene-mar)")


def _period_title(report: dict[str, Any]) -> str:
    return f"Gestión CI - {_period_label(report)}"


def _assets_crop(report: dict[str, Any], scope: str, module: str, name: str) -> Path | None:
    crops = report.get("dashboard_crops") if isinstance(report.get("dashboard_crops"), dict) else {}
    raw = (((crops.get(scope) or {}).get(module) or {}).get(name)) if isinstance(crops, dict) else None
    if not raw:
        return None
    p = Path(raw)
    if not p.is_absolute():
        p = (Path.cwd() / p).resolve()
    return p if p.exists() else None


def _prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = _in(SLIDE_W)
    prs.slide_height = _in(SLIDE_H)
    return prs


def _solid_bg(slide, color=COLORS["bg"]):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_text(slide, x, y, w, h, text, size=12, color=None, bold=False, font="Arial", align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(_in(x), _in(y), _in(w), _in(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = _in(0.02)
    tf.margin_right = _in(0.02)
    tf.margin_top = _in(0.01)
    tf.margin_bottom = _in(0.01)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _safe_text(text, "")
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color or COLORS["text"]
    return tb


def _add_rect(slide, x, y, w, h, fill, line=None, radius=False):
    shp = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE, _in(x), _in(y), _in(w), _in(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line or fill
    shp.line.width = Pt(0.6)
    return shp


def _add_logo(slide, white=False):
    path = BBVA_LOGO_WHITE if white else BBVA_LOGO_BLUE
    if path.exists():
        slide.shapes.add_picture(str(path), _in(11.92), _in(0.18), width=_in(0.95))
    else:
        _add_text(slide, 11.8, 0.16, 1.1, 0.35, "BBVA", size=16, color=COLORS["white"] if white else COLORS["bbva_blue"], bold=True, align=PP_ALIGN.RIGHT)


def _add_section_header(slide, subtitle: str):
    _add_text(slide, 0.48, 0.18, 8.0, 0.35, _period_title(report_context), size=19, color=COLORS["bbva_dark"], bold=True, font="Georgia")
    _add_logo(slide, white=False)
    _add_text(slide, 0.48, 0.58, 5.5, 0.22, subtitle, size=8, color=COLORS["muted"], bold=False)
    line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, _in(0.48), _in(0.92), _in(12.35), _in(0.02))
    line.fill.solid(); line.fill.fore_color.rgb = COLORS["border"]; line.line.color.rgb = COLORS["border"]


def _image_or_placeholder(slide, path: Path | None, x, y, w, h):
    if path and path.exists():
        slide.shapes.add_picture(str(path), _in(x), _in(y), width=_in(w), height=_in(h))
    else:
        _add_rect(slide, x, y, w, h, COLORS["white"], line=COLORS["border"])


def _kpi_card(slide, x, y, w, h, title, value, dark=False):
    fill = COLORS["bbva_blue"] if dark else COLORS["bbva_mid"]
    _add_rect(slide, x, y, w, h, fill, line=fill, radius=True)
    _add_text(slide, x + 0.14, y + 0.08, w - 0.28, 0.18, title, size=7, color=COLORS["white"], bold=True, align=PP_ALIGN.CENTER)
    _add_text(slide, x + 0.14, y + 0.25, w - 0.28, 0.28, value, size=17, color=COLORS["white"], bold=True, align=PP_ALIGN.CENTER)


def _obs_box(slide, x, y, w, h):
    _add_rect(slide, x, y, w, h, COLORS["obs"], line=COLORS["obs"])


def _rows(rows: Any) -> list[dict[str, Any]]:
    return rows if isinstance(rows, list) else []


def _top_mail_rows(scope_data: dict[str, Any], key: str, scope_label: str, max_rows: int = 2) -> list[list[str]]:
    out = []
    for row in _rows(scope_data.get(key))[:max_rows]:
        title = _clip(row.get("name") or row.get("title"), 46)
        metric = row.get("open_rate") if key == "top_push_by_open_rate" else row.get("interaction") or row.get("ctr")
        out.append([scope_label, title, _fmt_pct(metric)])
    return out


def _top_pull_rows(scope_data: dict[str, Any], key: str, scope_label: str, max_rows: int = 2) -> list[list[str]]:
    out = []
    for row in _rows(scope_data.get(key))[:max_rows]:
        title = _clip(row.get("title") or row.get("name"), 52)
        views = row.get("total_reads") or row.get("views") or row.get("page_views") or row.get("reads") or row.get("total_views")
        out.append([scope_label, title, _fmt_int(views)])
    return out


def _add_table(slide, x, y, w, h, title, headers, rows, col_widths=None, title_fill=COLORS["bbva_blue"], max_rows=None):
    _add_text(slide, x, y - 0.22, w, 0.18, title, size=8, color=COLORS["bbva_dark"], bold=True)
    use_rows = rows[:max_rows] if max_rows else rows
    rows_n = max(2, len(use_rows) + 1)
    table = slide.shapes.add_table(rows_n, len(headers), _in(x), _in(y), _in(w), _in(h)).table
    if col_widths:
        for idx, cw in enumerate(col_widths):
            table.columns[idx].width = _in(cw)
    header_h = h / rows_n
    for i, head in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = head
        cell.fill.solid(); cell.fill.fore_color.rgb = title_fill
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.runs[0]
        r.font.name = "Arial"; r.font.size = Pt(7); r.font.bold = True; r.font.color.rgb = COLORS["white"]
        cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    for r_idx in range(1, rows_n):
        for c_idx in range(len(headers)):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid(); cell.fill.fore_color.rgb = COLORS["white"]
            val = use_rows[r_idx - 1][c_idx] if r_idx - 1 < len(use_rows) else ""
            text_value = _safe_text(val, "")
            cell.text = text_value
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c_idx != len(headers) - 1 else PP_ALIGN.CENTER
            p.space_after = 0
            p.space_before = 0
            # python-pptx can leave paragraphs without runs when the value is empty.
            # Avoid tuple index errors in real dashboards with missing Top Five rows.
            r = p.runs[0] if p.runs else p.add_run()
            if not r.text:
                r.text = text_value
            r.font.name = "Arial"; r.font.size = Pt(6.2 if len(headers) > 3 else 6.8); r.font.color.rgb = COLORS["text"]
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    for row in table.rows:
        row.height = _in(h / rows_n)


def _cover(slide):
    if COVER_PATH.exists():
        slide.shapes.add_picture(str(COVER_PATH), 0, 0, width=_in(SLIDE_W), height=_in(SLIDE_H))
    else:
        _solid_bg(slide, COLORS["bbva_dark"])
        _add_logo(slide, white=True)
        _add_logo(slide, white=False)
        _add_text(slide, 1.0, 3.0, 6.0, 1.0, "Comunicaciones Internas", size=27, color=COLORS["white"], bold=True, font="Georgia")
        _add_text(slide, 1.0, 4.0, 2.0, 0.5, "Gestión Q1", size=18, color=COLORS["white"], font="Georgia")


def _planning_compare(slide, scopes, report):
    _solid_bg(slide)
    _add_section_header(slide, "Planificación | Argentina vs Holding")
    arg = scopes["argentina"]; hol = scopes["holding"]
    _kpi_card(slide, 1.30, 1.04, 3.05, 0.62, "ARGENTINA · Acciones de Comunicación", _fmt_int(arg.get("plan_total")), dark=True)
    _kpi_card(slide, 7.20, 1.04, 3.05, 0.62, "HOLDING · Acciones de Comunicación", _fmt_int(hol.get("plan_total")), dark=False)

    _add_text(slide, 0.55, 1.92, 2.6, 0.16, "Distribución por Eje Estratégico", size=7, color=COLORS["bbva_blue"], bold=True)
    _add_text(slide, 3.92, 1.92, 2.6, 0.16, "Distribución por Canales", size=7, color=COLORS["bbva_blue"], bold=True)
    _add_text(slide, 6.85, 1.92, 2.6, 0.16, "Distribución por Eje Estratégico", size=7, color=COLORS["bbva_blue"], bold=True)
    _add_text(slide, 10.22, 1.92, 2.6, 0.16, "Distribución por Canales", size=7, color=COLORS["bbva_blue"], bold=True)

    _image_or_placeholder(slide, _assets_crop(report, "argentina", "planning", "strategic_axes"), 0.55, 2.08, 3.0, 1.18)
    _image_or_placeholder(slide, _assets_crop(report, "argentina", "planning", "channel_mix"), 3.92, 2.08, 2.55, 1.18)
    _image_or_placeholder(slide, _assets_crop(report, "holding", "planning", "strategic_axes"), 6.85, 2.08, 3.0, 1.18)
    _image_or_placeholder(slide, _assets_crop(report, "holding", "planning", "channel_mix"), 10.22, 2.08, 2.55, 1.18)

    _add_text(slide, 0.55, 3.42, 2.3, 0.16, "Área solicitante · Argentina", size=7, color=COLORS["bbva_blue"], bold=True)
    _add_text(slide, 6.85, 3.42, 2.3, 0.16, "Área solicitante · Holding", size=7, color=COLORS["bbva_blue"], bold=True)
    _image_or_placeholder(slide, _assets_crop(report, "argentina", "planning", "internal_clients"), 0.55, 3.58, 5.90, 1.12)
    _image_or_placeholder(slide, _assets_crop(report, "holding", "planning", "internal_clients"), 6.85, 3.58, 5.90, 1.12)

    _obs_box(slide, 0.55, 5.10, 12.20, 1.15)


def _planning_combined(slide, scopes, report):
    _solid_bg(slide)
    _add_section_header(slide, "Planificación | Argentina + Holding")
    cmb = scopes["combined"]
    _kpi_card(slide, 4.62, 1.03, 4.05, 0.64, "Acciones de Comunicación", _fmt_int(cmb.get("plan_total")), dark=True)
    _add_text(slide, 0.55, 1.95, 2.8, 0.16, "Distribución por Eje Estratégico", size=7, color=COLORS["bbva_blue"], bold=True)
    _add_text(slide, 4.18, 1.95, 2.3, 0.16, "Distribución por Canales", size=7, color=COLORS["bbva_blue"], bold=True)
    _add_text(slide, 8.02, 1.95, 2.0, 0.16, "Área solicitante", size=7, color=COLORS["bbva_blue"], bold=True)
    _image_or_placeholder(slide, _assets_crop(report, "combined", "planning", "strategic_axes"), 0.55, 2.08, 3.10, 1.45)
    _image_or_placeholder(slide, _assets_crop(report, "combined", "planning", "channel_mix"), 4.18, 2.08, 3.10, 1.45)
    _image_or_placeholder(slide, _assets_crop(report, "combined", "planning", "internal_clients"), 8.02, 2.08, 4.73, 1.45)
    _obs_box(slide, 0.55, 4.55, 12.20, 1.55)


def _mail_compare(slide, scopes, report):
    _solid_bg(slide)
    _add_section_header(slide, "Canal Mail | Argentina vs Holding")
    arg = scopes["argentina"]; hol = scopes["holding"]
    headers = ["Scope", "Mail únicos / envíos", "Apertura", "Interacción s/enviados", "Interacción s/abiertos"]
    rows = [
        ["Argentina", f"{_fmt_int(arg.get('mail_unique_total') or arg.get('mail_total'))} / {_fmt_int(arg.get('mail_send_total') or arg.get('mail_total'))}", _fmt_pct(arg.get('mail_open_rate')), _fmt_pct(arg.get('mail_interaction_rate')), _fmt_pct(arg.get('mail_interaction_rate_over_opened'))],
        ["Holding", f"{_fmt_int(hol.get('mail_unique_total') or hol.get('mail_total'))} / {_fmt_int(hol.get('mail_send_total') or hol.get('mail_total'))}", _fmt_pct(hol.get('mail_open_rate')), _fmt_pct(hol.get('mail_interaction_rate')), _fmt_pct(hol.get('mail_interaction_rate_over_opened'))],
    ]
    _add_table(slide, 0.55, 1.15, 12.20, 0.88, "", headers, rows, col_widths=[1.0, 3.0, 1.2, 1.6, 1.6])
    _add_text(slide, 0.55, 2.18, 2.4, 0.16, "Tendencia mensual", size=7, color=COLORS['bbva_blue'], bold=True)
    _image_or_placeholder(slide, _assets_crop(report, 'argentina', 'mailing', 'monthly_trend'), 0.55, 2.32, 5.55, 1.55)
    _image_or_placeholder(slide, _assets_crop(report, 'holding', 'mailing', 'monthly_trend'), 0.55, 3.98, 5.55, 1.55)
    open_rows = _top_mail_rows(arg, 'top_push_by_open_rate', 'Argentina') + _top_mail_rows(hol, 'top_push_by_open_rate', 'Holding')
    int_rows = _top_mail_rows(arg, 'top_push_by_interaction', 'Argentina') + _top_mail_rows(hol, 'top_push_by_interaction', 'Holding')
    _add_table(slide, 6.45, 2.32, 6.30, 1.50, 'Top five - Mayor Tasa de Apertura', ['Scope', 'Título', 'Tasa'], open_rows, col_widths=[1.0, 4.3, 1.0], max_rows=4)
    _add_table(slide, 6.45, 4.02, 6.30, 1.50, 'Top five - Mayor Tasa de Interacción', ['Scope', 'Título', 'Tasa'], int_rows, col_widths=[1.0, 4.3, 1.0], max_rows=4)
    _obs_box(slide, 0.55, 5.95, 12.20, 0.92)


def _mail_combined(slide, scopes, report):
    _solid_bg(slide)
    _add_section_header(slide, "Canal Mail | Argentina + Holding")
    cmb = scopes['combined']
    headers = ['Mails únicos / envíos', 'Apertura', 'Interacción s/enviados', 'Interacción s/abiertos']
    rows = [[f"{_fmt_int(cmb.get('mail_unique_total') or cmb.get('mail_total'))} / {_fmt_int(cmb.get('mail_send_total') or cmb.get('mail_total'))}", _fmt_pct(cmb.get('mail_open_rate')), _fmt_pct(cmb.get('mail_interaction_rate')), _fmt_pct(cmb.get('mail_interaction_rate_over_opened'))]]
    _add_table(slide, 1.55, 1.20, 10.25, 0.72, '', headers, rows, col_widths=[3.1, 1.7, 2.1, 2.1])
    _add_text(slide, 0.55, 2.16, 2.4, 0.16, 'Tendencia mensual', size=7, color=COLORS['bbva_blue'], bold=True)
    _image_or_placeholder(slide, _assets_crop(report, 'combined', 'mailing', 'monthly_trend'), 0.55, 2.30, 6.0, 2.0)
    open_rows = _top_mail_rows(cmb, 'top_push_by_open_rate', 'Combined', max_rows=3)
    int_rows = _top_mail_rows(cmb, 'top_push_by_interaction', 'Combined', max_rows=3)
    _add_table(slide, 6.80, 2.30, 5.95, 1.65, 'Top five - Mayor Tasa de Apertura', ['Scope', 'Título', 'Tasa'], open_rows, col_widths=[1.0, 3.9, 1.0], max_rows=3)
    _add_table(slide, 6.80, 4.20, 5.95, 1.65, 'Top five - Mayor Tasa de Interacción', ['Scope', 'Título', 'Tasa'], int_rows, col_widths=[1.0, 3.9, 1.0], max_rows=3)
    _obs_box(slide, 0.55, 6.05, 12.20, 0.82)


def _content_compare(slide, scopes):
    _solid_bg(slide)
    _add_section_header(slide, 'Canal Intranet / Contenidos | Argentina vs Holding')
    arg = scopes['argentina']; hol = scopes['holding']
    headers = ['Scope', 'Noticias publicadas', 'Total páginas vistas', 'Promedio vistas']
    rows = [
        ['Argentina', _fmt_int(arg.get('site_notes_total')), _fmt_int(arg.get('site_total_views')), _fmt_int(arg.get('site_average_views'))],
        ['Holding', _fmt_int(hol.get('site_notes_total')), _fmt_int(hol.get('site_total_views')), _fmt_int(hol.get('site_average_views'))],
    ]
    _add_table(slide, 0.95, 1.18, 11.40, 0.86, '', headers, rows, col_widths=[1.2, 2.2, 2.3, 2.0])
    uu_rows = _top_pull_rows(arg, 'top_pull_notes', 'Argentina') + _top_pull_rows(hol, 'top_pull_notes', 'Holding')
    tgm_rows = _top_pull_rows(arg, 'top_pull_notes_tgm', 'Argentina') + _top_pull_rows(hol, 'top_pull_notes_tgm', 'Holding')
    _add_table(slide, 0.55, 2.45, 12.20, 1.45, 'Top five - Notas más leídas (uu)', ['Scope', 'Titular', 'Vistas'], uu_rows, col_widths=[1.0, 9.6, 1.2], max_rows=4)
    _add_table(slide, 0.55, 4.20, 12.20, 1.45, 'Top five - Notas más leídas (TGM)', ['Scope', 'Titular', 'Vistas'], tgm_rows, col_widths=[1.0, 9.6, 1.2], max_rows=4)
    _obs_box(slide, 0.55, 6.05, 12.20, 0.82)


def _content_combined(slide, scopes):
    _solid_bg(slide)
    _add_section_header(slide, 'Canal Intranet / Contenidos | Argentina + Holding')
    cmb = scopes['combined']
    headers = ['Noticias publicadas', 'Total páginas vistas', 'Promedio vistas']
    rows = [[_fmt_int(cmb.get('site_notes_total')), _fmt_int(cmb.get('site_total_views')), _fmt_int(cmb.get('site_average_views'))]]
    _add_table(slide, 2.05, 1.18, 9.20, 0.74, '', headers, rows, col_widths=[2.4, 2.6, 2.2])
    uu_rows = _top_pull_rows(cmb, 'top_pull_notes', 'Combined', max_rows=3)
    tgm_rows = _top_pull_rows(cmb, 'top_pull_notes_tgm', 'Combined', max_rows=3)
    _add_table(slide, 0.55, 2.35, 12.20, 1.55, 'Top five - Notas más leídas (uu)', ['Scope', 'Titular', 'Vistas'], uu_rows, col_widths=[1.0, 9.6, 1.2], max_rows=3)
    _add_table(slide, 0.55, 4.18, 12.20, 1.55, 'Top five - Notas más leídas (TGM)', ['Scope', 'Titular', 'Vistas'], tgm_rows, col_widths=[1.0, 9.6, 1.2], max_rows=3)
    _obs_box(slide, 0.55, 6.08, 12.20, 0.80)


def _closing(slide):
    _solid_bg(slide, COLORS['bg'])
    _add_logo(slide, white=False)
    _add_text(slide, 0.80, 3.15, 6.0, 0.6, 'Comunicaciones Internas', size=23, color=COLORS['bbva_dark'], bold=True, font='Georgia')
    _add_text(slide, 0.80, 3.78, 3.0, 0.35, _safe_text(_period_label(report_context)), size=12, color=COLORS['muted'])


report_context: dict[str, Any] = {}


def render_management_deck(report: dict[str, Any], output_path: Path) -> None:
    global report_context
    report_context = report
    prs = _prs()
    scopes = _scope_bundle(report)
    missing = [key for key in ("argentina", "holding", "combined") if not scopes.get(key)]
    if missing:
        raise ValueError(f"Faltan scopes requeridos para renderizar el informe: {', '.join(missing)}")
    slides = [
        _cover,
        lambda s: _planning_compare(s, scopes, report),
        lambda s: _planning_combined(s, scopes, report),
        lambda s: _mail_compare(s, scopes, report),
        lambda s: _mail_combined(s, scopes, report),
        lambda s: _content_compare(s, scopes),
        lambda s: _content_combined(s, scopes),
        _closing,
    ]
    blank = prs.slide_layouts[6]
    for builder in slides:
        slide = prs.slides.add_slide(blank)
        builder(slide)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))


def create_pptx(report: dict[str, Any], output_path: Path, template_mode: str = 'full', template_path: Path | None = None) -> None:
    output_path = Path(output_path)
    safe_report = validate_report_json(report)
    render_management_deck(safe_report, output_path)


def main() -> None:
    repo_root = Path(__file__).resolve().parent.parent
    if len(sys.argv) == 3:
        input_json = Path(sys.argv[1])
        output_pptx = Path(sys.argv[2])
    else:
        input_json = repo_root / 'data' / 'report_boceto_ci_sample.json'
        output_pptx = repo_root / 'sample_bbva_report_definitive.pptx'
    report = json.loads(input_json.read_text(encoding='utf-8'))
    create_pptx(report, output_pptx)
    print(f'PPTX generado: {output_pptx}')


if __name__ == '__main__':
    main()
