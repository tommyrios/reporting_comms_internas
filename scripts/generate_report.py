import argparse
from cProfile import label
import json
import logging
import os
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

from analyzer import (
    BASE_STRUCTURE,
    compute_kpis,
    validate_monthly_summary_contract,
    validate_report_json,
)
from config import DATA_DIR, INBOX_PDF_DIR, MANUAL_CONTEXT_DIR, REPORTS_DIR, ensure_dir
from dashboard_crops import build_dashboard_crops
from history_manager import apply_historical_comparison, persist_calculated_totals
from period_pdf_processor import resolve_period_scope_pdfs, summarize_period_scope
from period_scopes import required_scopes_from_env
from pptx_renderer import create_pptx

logger = logging.getLogger(__name__)


def get_period_definition(period_slug: str) -> dict[str, Any]:
    schedule_path = DATA_DIR / "reporting_periods.json"
    if schedule_path.exists():
        payload = json.loads(schedule_path.read_text(encoding="utf-8"))
    else:
        from reporting_periods import load_schedule
        payload = load_schedule().to_dict()

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


def write_report_artifacts(period_slug: str, report: dict[str, Any], metadata_extra: dict[str, Any] | None = None) -> str:
    report_dir = ensure_dir(REPORTS_DIR / period_slug)
    period_info = report.get("period", {})
    period_label = str(period_info.get("label", "")).strip()
    metadata = {
        "title": "Comunicaciones Internas",
        "subtitle": "Informe de gestión",
        "period": period_label,
        "period_slug": period_slug,
        "generated_at": datetime.utcnow().isoformat() + "Z",
    }
    if metadata_extra:
        metadata.update(metadata_extra)

    if label:
        pptx_filename = f"Informe de Gestión CI - {label}.pptx"
    else:
        pptx_filename = f"Informe de Gestión CI - {period_slug}.pptx"

    metadata_path = report_dir / "metadata.json"
    raw_path = report_dir / "report_raw.json"
    pptx_path = report_dir / pptx_filename
    html_path = report_dir / "report.html"

    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    raw_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    create_pptx(report, pptx_path, template_mode="full")

    html_content = (
        "<html><body>"
        f"<h2>{metadata['title']}</h2>"
        f"<p>Período: {metadata['period']}</p>"
        "<p>El reporte fue generado en formato PowerPoint (.pptx).</p>"
        "</body></html>"
    )
    html_path.write_text(html_content, encoding="utf-8")

    required = [metadata_path, html_path, pptx_path]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        existing = sorted(path.name for path in report_dir.glob("*"))
        raise RuntimeError(
            "La generación terminó sin artefactos requeridos. "
            f"Faltantes: {missing}. Archivos existentes en {report_dir}: {existing}"
        )

    return str(report_dir)


def _build_scope_comparison(period_summaries: dict[str, dict[str, Any]]) -> dict[str, Any]:
    def pick(scope: str) -> dict[str, Any]:
        summary = period_summaries.get(scope, {})
        return {
            "scope": scope,
            "scope_label": summary.get("scope_label", scope),
            "plan_total": summary.get("plan_total", 0),
            "plan_daily_average": summary.get("plan_daily_average", 0),
            "mail_total": summary.get("mail_total", 0),
            "mail_send_total": summary.get("mail_send_total", summary.get("mail_total", 0)),
            "mail_unique_total": summary.get("mail_unique_total", 0),
            "mail_planning_pct": summary.get("mail_planning_pct", 0),
            "mail_open_rate": summary.get("mail_open_rate", 0),
            "mail_interaction_rate": summary.get("mail_interaction_rate", 0),
            "mail_interaction_rate_over_opened": summary.get("mail_interaction_rate_over_opened", 0),
            "site_notes_total": summary.get("site_notes_total", 0),
            "site_total_views": summary.get("site_total_views", 0),
            "site_average_views": summary.get("site_average_views", 0),
            "channel_mix": summary.get("channel_mix", []),
            "strategic_axes": summary.get("strategic_axes", []),
            "internal_clients": summary.get("internal_clients", []),
            "format_mix": summary.get("format_mix", []),
            "monthly_trend": summary.get("monthly_trend", summary.get("mail_monthly_trend", [])),
            "top_push_by_interaction": summary.get("top_push_by_interaction", []),
            "top_push_by_open_rate": summary.get("top_push_by_open_rate", []),
            "top_pull_notes": summary.get("top_pull_notes", []),
            "top_pull_notes_tgm": summary.get("top_pull_notes_tgm", summary.get("top_pull_tgm", [])),
        }

    return {scope: pick(scope) for scope in period_summaries.keys()}


def generate_period_report(period_slug: str, force_regenerate: bool = False, pdf_dir: Path | None = None) -> dict[str, Any]:
    period = get_period_definition(period_slug)
    allow_partial = (os.environ.get("ALLOW_PARTIAL_PERIOD") or "false").lower() == "true"
    required_scopes = required_scopes_from_env(os.environ.get("REPORT_REQUIRED_SCOPES"))
    scope_pdf_paths = resolve_period_scope_pdfs(period_slug, scopes=required_scopes, pdf_dir=pdf_dir, allow_partial=allow_partial)

    period_summaries_raw: dict[str, dict[str, Any]] = {}
    warnings: list[str] = []

    for scope in required_scopes:
        summary = summarize_period_scope(period, scope, force_regenerate=force_regenerate, pdf_dir=pdf_dir)
        period_summaries_raw[scope] = summary
        validation = summary.get("validation") if isinstance(summary.get("validation"), dict) else {}
        if summary.get("generation_mode") == "local_fallback":
            raise ValueError(
                f"No se pudo generar resumen determinístico para {period_slug} [{scope}]: "
                f"{summary.get('warning', 'sin detalles')}"
            )
        if summary.get("generation_mode") == "deterministic_pdf" and not validation.get("is_valid", False):
            raise ValueError(f"Resumen inválido para {period_slug} [{scope}]: {validation.get('errors', [])}")
        if validation.get("warnings"):
            warnings.extend([f"{scope}: {warning}" for warning in validation.get("warnings", [])])

    # El consolidado Argentina + Holding es la fuente del total ejecutivo. No se suma a mano.
    main_scope = "combined" if "combined" in period_summaries_raw else required_scopes[0]
    main_summary = validate_monthly_summary_contract(period_summaries_raw[main_scope])
    kpis_calculados = compute_kpis([main_summary])
    kpis_calculados["scopes"] = _build_scope_comparison(period_summaries_raw)
    kpis_calculados["main_scope"] = main_scope
    kpis_calculados.setdefault("quality_flags", {})["historical_comparison_allowed"] = False
    kpis_calculados = apply_historical_comparison(period, kpis_calculados)

    totals = kpis_calculados.setdefault("calculated_totals", {})
    totals["volume_change"] = "No comparable: fuente trimestral filtrada"
    totals["latest_push_variation"] = "No comparable: fuente trimestral filtrada"

    dashboard_crops = build_dashboard_crops(
        period_slug=period_slug,
        scope_pdf_paths=scope_pdf_paths,
        output_dir=REPORTS_DIR / period_slug,
    )

    # El nuevo informe de gestión es determinístico y editable.
    # No se genera narrativa ni plan de mejora con GenAI para el PPTX.
    narrative: dict[str, Any] = {}
    narrative_mode = "deterministic"
    render_plan = {
        "template_mode": "full",
        "period": {"slug": period.get("slug"), "label": period.get("label")},
        "modules": [],
    }

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
            "dashboard_crops": dashboard_crops,
        }
    )

    manual_context = load_manual_context(period_slug)
    if manual_context:
        report = _deep_merge(report, {k: v for k, v in manual_context.items() if k != "metadata"})

    modules = [module.get("key") for module in report.get("render_plan", {}).get("modules", []) if isinstance(module, dict)]
    is_events_omitted = "events" not in modules

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


def _normalize_period_slug(period_arg: str | None) -> str | None:
    if not period_arg:
        return None
    period_arg = period_arg.strip()
    if period_arg.startswith(("quarter_", "year_")):
        return period_arg
    if len(period_arg) == 7 and period_arg[4] == "-" and period_arg[5].upper() == "Q":
        return f"quarter_{period_arg[:4]}_{period_arg[5:].upper()}"
    return period_arg


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera reporte de período desde PDFs locales (opcionalmente fetch email).")
    parser.add_argument("--period", default=os.environ.get("REPORT_SLUG", "").strip(), help="Slug de período (ej. 2026-Q1 / quarter_2026_Q1 / year_2026).")
    parser.add_argument("--pdf-dir", type=Path, default=INBOX_PDF_DIR, help="Directorio local de PDFs trimestrales/anuales por scope.")
    parser.add_argument("--force-regenerate", action="store_true", help="Ignora cache de período/scope y reextrae.")
    fetch_group = parser.add_mutually_exclusive_group()
    fetch_group.add_argument("--fetch-email", action="store_true", help="Primero ingiere PDFs desde email.")
    fetch_group.add_argument("--skip-email-fetch", action="store_true", help="No usa email; procesa solo PDFs locales.")
    args = parser.parse_args()

    if args.fetch_email:
        from fetch_dashboard_pdfs import run_ingestion
        run_ingestion(pdf_dir=args.pdf_dir)

    period_slug = _normalize_period_slug(args.period)
    if period_slug:
        print(
            json.dumps(
                generate_period_report(period_slug, force_regenerate=args.force_regenerate, pdf_dir=args.pdf_dir),
                ensure_ascii=False,
                indent=2,
            )
        )


if __name__ == "__main__":
    main()
