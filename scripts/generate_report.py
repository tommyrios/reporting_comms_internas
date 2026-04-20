import json
import os
from copy import deepcopy
from datetime import datetime
from typing import Any

from analyzer import BASE_STRUCTURE, compute_kpis, validate_report_json
from config import DATA_DIR, MANUAL_CONTEXT_DIR, REPORTS_DIR, SUMMARIES_DIR, ensure_dir
from llm_client import build_genai_client, call_gemini_for_json, load_prompt
from pdf_processor import summarize_month
from pptx_renderer import create_pptx


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


def _fmt_pct(value: Any) -> str:
    if value in (None, "", "-"):
        return "-"
    try:
        n = float(str(value).replace("%", "").replace(",", "."))
    except Exception:
        return str(value)
    if n.is_integer():
        return f"{int(n)}%"
    return f"{str(round(n, 1)).replace('.', ',')}%"


def build_fallback_report(period: dict[str, Any], kpis: dict[str, Any]) -> dict[str, Any]:
    report = validate_report_json(deepcopy(BASE_STRUCTURE))
    totals = kpis.get("calculated_totals", {})
    timelines = kpis.get("timelines", {})
    distributions = kpis.get("aggregated_distributions", {})
    rankings = kpis.get("consolidated_rankings", {})

    report["slide_1_cover"].update({
        "area": "Comunicaciones Internas",
        "period": period.get("label", "-"),
        "subtitle": "Informe de gestión",
    })

    report["slide_2_overview"].update({
        "headline": "¿Cómo nos fue? CI",
        "volume_current": totals.get("push_volume_period", "-"),
        "volume_previous": totals.get("previous_push_volume", "-"),
        "volume_change": totals.get("latest_push_variation", "-"),
        "push_open_rate": _fmt_pct(totals.get("average_open_rate", "-")),
        "push_interaction_rate": _fmt_pct(totals.get("average_interaction_rate", "-")),
        "pull_notes_current": totals.get("pull_notes_period", "-"),
        "average_reads": totals.get("average_reads_per_note", "-"),
        "audience_segments": distributions.get("audience_segments", []),
        "comparison_timeline": timelines.get("push_volume", []),
        "comparative_note": "La comparación muestra el comportamiento del volumen principal del período.",
        "conclusion_message": "El deck se generó en modo fallback a partir de KPIs consolidados del período.",
        "highlights": [
            f"{totals.get('push_volume_period', 0)} comunicaciones relevadas.",
            f"{totals.get('pull_notes_period', 0)} contenidos pull publicados.",
            f"{_fmt_pct(totals.get('average_open_rate', '-'))} de apertura promedio.",
        ],
    })

    report["slide_3_plan"].update({
        "mail_total": totals.get("push_volume_period", "-"),
        "open_rate": _fmt_pct(totals.get("average_open_rate", "-")),
        "pull_total": totals.get("pull_notes_period", "-"),
        "mail_timeline": timelines.get("push_volume", []),
        "pull_timeline": timelines.get("pull_notes", []),
        "mail_message": "El canal principal mantuvo el volumen operativo del período.",
        "pull_message": "El canal pull aportó profundidad de lectura y continuidad editorial.",
        "footer": "La gestión del período combinó volumen, lectura y continuidad de agenda.",
    })

    report["slide_4_strategy"].update({
        "content_distribution": distributions.get("strategic_axes", []),
        "internal_clients": distributions.get("internal_clients", []),
        "theme_message": "La agenda se concentró en los ejes con mayor peso relativo.",
        "balance_message": "El balance editorial debe leerse junto al mix institucional y de servicio.",
        "conclusion": "La estrategia del período puede explicarse a partir del mix temático consolidado.",
    })

    report["slide_5_push_ranking"].update({
        "top_communications": rankings.get("top_push", []),
        "key_learning": "Las comunicaciones top ayudan a detectar formatos y temáticas con mejor respuesta.",
    })

    report["slide_6_pull_performance"].update({
        "pub_current": totals.get("pull_notes_period", "-"),
        "top_notes": rankings.get("top_pull", []),
        "avg_reads": totals.get("average_reads_per_note", "-"),
        "total_views": totals.get("pull_reads_period", "-"),
        "secondary_message": "El desempeño pull permite leer profundidad y tracción de contenidos.",
        "conclusion": "Las notas top marcan qué contenidos sostuvieron lectura en el período.",
    })

    report["slide_7_hitos"] = kpis.get("hitos_crudos", [])
    report["slide_8_events"].update({
        "conclusion": "No se detectó un consolidado de eventos en la fuente automática de esta corrida.",
    })
    report["slide_9_closure"].update({
        "bullets": [
            report["slide_2_overview"]["conclusion_message"],
            report["slide_5_push_ranking"]["key_learning"],
            report["slide_6_pull_performance"]["conclusion"],
        ]
    })
    return report


def _decorate_report(report_raw: dict[str, Any], period: dict[str, Any], kpis: dict[str, Any]) -> dict[str, Any]:
    report = validate_report_json(report_raw)
    totals = kpis.get("calculated_totals", {})
    timelines = kpis.get("timelines", {})
    distributions = kpis.get("aggregated_distributions", {})
    rankings = kpis.get("consolidated_rankings", {})

    report["slide_1_cover"]["area"] = report["slide_1_cover"].get("area") or "Comunicaciones Internas"
    report["slide_1_cover"]["period"] = period.get("label") or report["slide_1_cover"].get("period")
    report["slide_1_cover"]["subtitle"] = report["slide_1_cover"].get("subtitle") or "Informe de gestión"

    overview = report["slide_2_overview"]
    overview["volume_current"] = overview.get("volume_current") if overview.get("volume_current") != "-" else totals.get("push_volume_period", "-")
    overview["volume_previous"] = overview.get("volume_previous") if overview.get("volume_previous") != "-" else totals.get("previous_push_volume", "-")
    overview["volume_change"] = overview.get("volume_change") if overview.get("volume_change") != "-" else totals.get("latest_push_variation", "-")
    overview["push_open_rate"] = overview.get("push_open_rate") if overview.get("push_open_rate") != "-" else _fmt_pct(totals.get("average_open_rate", "-"))
    overview["push_interaction_rate"] = overview.get("push_interaction_rate") if overview.get("push_interaction_rate") != "-" else _fmt_pct(totals.get("average_interaction_rate", "-"))
    overview["pull_notes_current"] = overview.get("pull_notes_current") if overview.get("pull_notes_current") != "-" else totals.get("pull_notes_period", "-")
    overview["average_reads"] = overview.get("average_reads") if overview.get("average_reads") != "-" else totals.get("average_reads_per_note", "-")
    if not overview.get("audience_segments"):
        overview["audience_segments"] = distributions.get("audience_segments", [])
    if not overview.get("comparison_timeline"):
        overview["comparison_timeline"] = timelines.get("push_volume", [])
    if not overview.get("highlights"):
        overview["highlights"] = [
            f"{totals.get('push_volume_period', 0)} comunicaciones relevadas.",
            f"{totals.get('pull_notes_period', 0)} contenidos pull.",
            f"{_fmt_pct(totals.get('average_open_rate', '-'))} de apertura promedio.",
        ]

    plan = report["slide_3_plan"]
    plan["mail_total"] = plan.get("mail_total") if plan.get("mail_total") != "-" else totals.get("push_volume_period", "-")
    plan["pull_total"] = plan.get("pull_total") if plan.get("pull_total") != "-" else totals.get("pull_notes_period", "-")
    plan["open_rate"] = plan.get("open_rate") if plan.get("open_rate") != "-" else _fmt_pct(totals.get("average_open_rate", "-"))
    if not plan.get("mail_timeline"):
        plan["mail_timeline"] = timelines.get("push_volume", [])
    if not plan.get("pull_timeline"):
        plan["pull_timeline"] = timelines.get("pull_notes", [])

    strategy = report["slide_4_strategy"]
    if not strategy.get("content_distribution"):
        strategy["content_distribution"] = distributions.get("strategic_axes", [])
    if not strategy.get("internal_clients"):
        strategy["internal_clients"] = distributions.get("internal_clients", [])

    if not report["slide_5_push_ranking"].get("top_communications"):
        report["slide_5_push_ranking"]["top_communications"] = rankings.get("top_push", [])

    pull = report["slide_6_pull_performance"]
    if not pull.get("top_notes"):
        pull["top_notes"] = rankings.get("top_pull", [])
    pull["pub_current"] = pull.get("pub_current") if pull.get("pub_current") != "-" else totals.get("pull_notes_period", "-")
    pull["avg_reads"] = pull.get("avg_reads") if pull.get("avg_reads") != "-" else totals.get("average_reads_per_note", "-")
    pull["total_views"] = pull.get("total_views") if pull.get("total_views") != "-" else totals.get("pull_reads_period", "-")

    if not report.get("slide_7_hitos"):
        report["slide_7_hitos"] = kpis.get("hitos_crudos", [])

    closure = report["slide_9_closure"]
    if not closure.get("bullets"):
        closure["bullets"] = [
            overview.get("conclusion_message", "-"),
            report["slide_5_push_ranking"].get("key_learning", "-"),
            pull.get("conclusion", "-"),
        ]

    return report


def write_report_artifacts(period_slug: str, report: dict[str, Any], metadata_extra: dict[str, Any] | None = None) -> str:
    report_dir = ensure_dir(REPORTS_DIR / period_slug)
    metadata = {
        "title": report.get("slide_1_cover", {}).get("area"),
        "subtitle": report.get("slide_1_cover", {}).get("subtitle"),
        "period": report.get("slide_1_cover", {}).get("period"),
        "period_slug": period_slug,
        "generated_at": datetime.utcnow().isoformat() + "Z",
    }
    if metadata_extra:
        metadata.update(metadata_extra)

    (report_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    (report_dir / "report_raw.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    create_pptx(report, report_dir / "report.pptx")
    html_content = (
        "<html><body>"
        f"<h2>{metadata['title']}</h2>"
        f"<p>Período: {metadata['period']}</p>"
        "<p>El reporte fue generado en formato PowerPoint (.pptx).</p>"
        "</body></html>"
    )
    (report_dir / "report.html").write_text(html_content, encoding="utf-8")
    return str(report_dir)


def generate_period_report(period_slug: str, force_regenerate: bool = False) -> dict[str, Any]:
    period = get_period_definition(period_slug)
    client = build_genai_client()
    summaries = []
    summary_warnings = []
    for month_key in period.get("months", []):
        try:
            summaries.append(summarize_month(client, month_key, force_regenerate))
        except Exception as exc:
            cached_summary_path = SUMMARIES_DIR / f"{month_key}.json"
            if cached_summary_path.exists():
                cached_summary = json.loads(cached_summary_path.read_text(encoding="utf-8"))
                summaries.append(cached_summary)
                summary_warnings.append(
                    f"Se reutilizó el summary cacheado de {month_key} porque Gemini falló: {exc}"
                )
            else:
                raise RuntimeError(
                    f"No se pudo generar el resumen mensual de {month_key} y no existe caché previa. Error: {exc}"
                ) from exc

    kpis_calculados = compute_kpis(summaries)

    prompt_base = load_prompt("period_report.txt")
    prompt_final = (
        f"{prompt_base}\n\n"
        f"INPUT (PERIODO):\n{json.dumps(period, ensure_ascii=False)}\n\n"
        f"INPUT (KPI_CALCULADOS):\n{json.dumps(kpis_calculados, ensure_ascii=False)}\n\n"
        f"CONTEXTO (RESUMENES_MENSUALES):\n{json.dumps(summaries, ensure_ascii=False)}"
    )

    generation_mode = "llm"
    warning = None
    try:
        report_raw = call_gemini_for_json(client, [prompt_final])
        report = _decorate_report(report_raw, period, kpis_calculados)
    except Exception as exc:
        generation_mode = "fallback"
        warning = f"Se generó el reporte sin redacción del LLM: {exc}"
        report = build_fallback_report(period, kpis_calculados)

    all_warnings = []
    if warning:
        all_warnings.append(warning)
    all_warnings.extend(summary_warnings)
    combined_warning = " | ".join(all_warnings) if all_warnings else None

    manual_context = load_manual_context(period_slug)
    if manual_context:
        report = _deep_merge(report, {k: v for k, v in manual_context.items() if k != "metadata"})

    report = validate_report_json(report)
    metadata_extra = {
        "warning": combined_warning or manual_context.get("metadata", {}).get("warning"),
        "email_subject": manual_context.get("metadata", {}).get("email_subject") or period.get("email_subject"),
        "generation_mode": generation_mode,
    }
    report_dir = write_report_artifacts(period_slug, report, metadata_extra=metadata_extra)
    return {
        "status": "ok",
        "period_slug": period_slug,
        "report_dir": report_dir,
        "generation_mode": generation_mode,
        "warning": combined_warning,
    }


def main() -> None:
    period_slug = os.environ.get("REPORT_SLUG", "").strip()
    if period_slug:
        print(json.dumps(generate_period_report(period_slug), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
