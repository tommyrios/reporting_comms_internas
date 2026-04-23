import json
import logging
import os
from copy import deepcopy
from datetime import datetime
from typing import Any

from analyzer import (
    BASE_STRUCTURE,
    build_render_plan,
    compute_kpis,
    validate_monthly_summary_contract,
    validate_report_json,
)
from config import DATA_DIR, MANUAL_CONTEXT_DIR, REPORTS_DIR, ensure_dir
from history_manager import apply_historical_comparison, persist_calculated_totals
from llm_client import build_genai_client, call_gemini_for_json, load_prompt
from pdf_processor import summarize_month
from pptx_renderer import create_pptx

logger = logging.getLogger(__name__)


def get_period_definition(period_slug: str) -> dict[str, Any]:
    schedule_path = DATA_DIR / "reporting_periods.json"
    legacy_path = DATA_DIR / "selected_periods.json"
    source_path = schedule_path if schedule_path.exists() else legacy_path
    payload = json.loads(source_path.read_text(encoding="utf-8"))
    for period in payload.get("periods", []):
        if period.get("slug") == period_slug:
            return period
    raise KeyError(period_slug)


def _period_manual_context_path(period_slug: str):
    return MANUAL_CONTEXT_DIR / f"{period_slug}.json"


def _deep_merge(base: Any, override: Any) -> Any:
    if isinstance(base, dict) and isinstance(override, dict):
        merged = dict(base)
        for key, value in override.items():
            merged[key] = _deep_merge(merged.get(key), value)
        return merged
    if isinstance(base, list) and isinstance(override, list):
        return override
    return override if override is not None else base


def load_manual_context(period_slug: str) -> dict[str, Any]:
    path = _period_manual_context_path(period_slug)
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def _build_fallback_narrative(kpis: dict[str, Any]) -> dict[str, Any]:
    totals = kpis.get("calculated_totals", {})
    flags = kpis.get("quality_flags", {})
    historical_note = (
        "No comparable por alcance de fuente"
        if not flags.get("historical_comparison_allowed", True)
        else "Comparación histórica habilitada para el mismo alcance de fuentes."
    )
    return {
        "executive_summary": historical_note,
        "executive_takeaways": [
            f"{totals.get('plan_total', 0)} comunicaciones planificadas en el período.",
            f"{totals.get('site_notes_total', 0)} noticias publicadas con {totals.get('site_total_views', 0)} páginas vistas.",
            f"Apertura promedio {totals.get('mail_open_rate', 0)}% e interacción {totals.get('mail_interaction_rate', 0)}%.",
        ],
        "channel_management": "El desempeño de canales se consolidó con métricas verificables y sin inferencias.",
        "mix_thematic_clients": "El mix temático y de áreas solicitantes resume la demanda efectiva del período.",
        "ranking_push": "El ranking push se construyó solo con datos observables en la fuente.",
        "ranking_pull": "El ranking pull prioriza lecturas reales y evita proyecciones no sustentadas.",
        "milestones": "Los hitos incluidos reflejan actividades detectadas en el mes.",
        "events": "La sección de eventos se muestra únicamente cuando existe detalle suficiente.",
    }


def _sanitize_narrative(raw: Any) -> dict[str, Any]:
    if not isinstance(raw, dict):
        return {}
    clean: dict[str, Any] = {}
    for key, value in raw.items():
        if isinstance(value, str):
            text = " ".join(value.replace("_", " ").split())
            if "slide_" in text.lower():
                text = text.replace("slide_", "")
            clean[key] = text
        elif isinstance(value, list):
            clean[key] = [" ".join(str(item).replace("_", " ").split()) for item in value if str(item).strip()][:4]
    return clean


def write_report_artifacts(period_slug: str, report: dict[str, Any], metadata_extra: dict[str, Any] | None = None) -> str:
    report_dir = ensure_dir(REPORTS_DIR / period_slug)
    period_label = report.get("period", {}).get("label", "-")
    metadata = {
        "title": "Comunicaciones Internas",
        "subtitle": "Informe ejecutivo mensual",
        "period": period_label,
        "period_slug": period_slug,
        "generated_at": datetime.utcnow().isoformat() + "Z",
    }
    if metadata_extra:
        metadata.update(metadata_extra)

    (report_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    (report_dir / "report_raw.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    create_pptx(report, report_dir / "report.pptx", template_mode="frame")
    html_content = (
        "<html><body>"
        f"<h2>{metadata['title']}</h2>"
        f"<p>Período: {metadata['period']}</p>"
        "<p>El reporte fue generado en formato PowerPoint (.pptx).</p>"
        "</body></html>"
    )
    (report_dir / "report.html").write_text(html_content, encoding="utf-8")
    return str(report_dir)


def _request_narrative(period: dict[str, Any], kpis: dict[str, Any], monthly_summaries: list[dict[str, Any]]) -> tuple[dict[str, Any], str, str | None]:
    try:
        client = build_genai_client()
        prompt_base = load_prompt("period_report.txt")
        prompt_final = (
            f"{prompt_base}\n\n"
            f"INPUT (PERIODO):\n{json.dumps(period, ensure_ascii=False)}\n\n"
            f"INPUT (KPI_CALCULADOS):\n{json.dumps(kpis, ensure_ascii=False)}\n\n"
            f"INPUT (RESUMENES_MENSUALES):\n{json.dumps(monthly_summaries, ensure_ascii=False)}"
        )
        narrative_raw = call_gemini_for_json(client, [prompt_final])
        return _sanitize_narrative(narrative_raw), "llm", None
    except Exception as exc:
        return _build_fallback_narrative(kpis), "fallback", f"Se generó narrativa fallback: {str(exc)}"


def generate_period_report(period_slug: str, force_regenerate: bool = False) -> dict[str, Any]:
    period = get_period_definition(period_slug)

    monthly_summaries_raw: list[dict[str, Any]] = []
    monthly_summaries_validated: list[dict[str, Any]] = []
    warnings: list[str] = []

    for month_key in period.get("months", []):
        summary = summarize_month(None, month_key, force_regenerate)
        monthly_summaries_raw.append(summary)
        validation = summary.get("validation") if isinstance(summary.get("validation"), dict) else {}
        if summary.get("generation_mode") == "local_fallback":
            raise ValueError(summary.get("warning") or f"Resumen mensual en fallback para {month_key}")
        if summary.get("generation_mode") == "deterministic_pdf" and not validation.get("is_valid", False):
            raise ValueError(f"Resumen mensual inválido para {month_key}: {validation.get('errors', [])}")
        if validation.get("warnings"):
            warnings.extend([f"{month_key}: {warning}" for warning in validation.get("warnings", [])])

        validated = validate_monthly_summary_contract(summary)
        monthly_summaries_validated.append(validated)

    kpis_calculados = compute_kpis(monthly_summaries_validated)
    kpis_calculados = apply_historical_comparison(period, kpis_calculados)

    if not kpis_calculados.get("quality_flags", {}).get("historical_comparison_allowed", True):
        totals = kpis_calculados.setdefault("calculated_totals", {})
        totals["volume_change"] = "No comparable por alcance de fuente"
        totals["latest_push_variation"] = "No comparable por alcance de fuente"

    narrative, narrative_mode, narrative_warning = _request_narrative(period, kpis_calculados, monthly_summaries_raw)
    if narrative_warning:
        warnings.append(narrative_warning)

    render_plan = build_render_plan(period, kpis_calculados, narrative)

    report = validate_report_json(
        {
            **deepcopy(BASE_STRUCTURE),
            "period": {
                "slug": period.get("slug"),
                "label": period.get("label"),
            },
            "kpis": kpis_calculados,
            "narrative": narrative,
            "quality_flags": kpis_calculados.get("quality_flags", {}),
            "render_plan": render_plan,
        }
    )

    manual_context = load_manual_context(period_slug)
    if manual_context:
        report = _deep_merge(report, {k: v for k, v in manual_context.items() if k != "metadata"})

    modules = [module.get("key") for module in report.get("render_plan", {}).get("modules", []) if isinstance(module, dict)]
    is_events_omitted = "events" not in modules
    if is_events_omitted:
        warnings.append("Módulo de eventos omitido por falta de datos suficientes")

    metadata_extra = {
        "warning": " | ".join(warnings) if warnings else manual_context.get("metadata", {}).get("warning"),
        "email_subject": manual_context.get("metadata", {}).get("email_subject") or period.get("email_subject"),
        "generation_mode": narrative_mode,
        "rendered_modules": modules,
        "omitted_modules": ["events"] if is_events_omitted else [],
    }

    persist_calculated_totals(period, kpis_calculados)
    report_dir = write_report_artifacts(period_slug, report, metadata_extra=metadata_extra)

    logger.info("render_modules=%s", modules)
    if is_events_omitted:
        logger.info("omit_module=events reason=insufficient_data")

    return {
        "status": "ok",
        "period_slug": period_slug,
        "report_dir": report_dir,
        "generation_mode": narrative_mode,
        "warning": metadata_extra["warning"],
    }


def main() -> None:
    period_slug = os.environ.get("REPORT_SLUG", "").strip()
    if period_slug:
        print(json.dumps(generate_period_report(period_slug), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
