import json
import logging
from copy import deepcopy
from pathlib import Path
from typing import Any

from config import (
    CANONICAL_MONTHLY_DIR,
    INBOX_PDF_DIR,
    LEGACY_SUMMARIES_DIR,
    PDF_DIR,
    RAW_EXTRACTED_DIR,
    SUMMARIES_DIR,
    VALIDATION_DIR,
    ensure_dir,
)
from deterministic_pipeline import (
    canonicalize_monthly,
    extract_raw_monthly_pdf,
    persist_monthly_artifacts,
    validate_canonical_monthly,
)

logger = logging.getLogger(__name__)


def _build_local_fallback_summary(month_key: str, warning: str) -> dict:
    return {
        "month": month_key,
        "generation_mode": "local_fallback",
        "warning": warning,
        "plan_total": 0,
        "site_notes_total": 0,
        "site_total_views": 0,
        "mail_total": 0,
        "mail_open_rate": 0,
        "mail_interaction_rate": 0,
        "strategic_axes": [],
        "internal_clients": [],
        "channel_mix": [],
        "format_mix": [],
        "top_push_by_interaction": [],
        "top_push_by_open_rate": [],
        "top_pull_notes": [],
        "hitos": [],
        "events": [],
        "quality_flags": {
            "scope_country": "AR",
            "scope_mixed": False,
            "site_has_no_data_sections": True,
            "events_summary_available": False,
            "push_ranking_available": False,
            "pull_ranking_available": False,
            "historical_comparison_allowed": True,
        },
    }


def _cache_candidates(month_key: str) -> list[Path]:
    return [
        ensure_dir(CANONICAL_MONTHLY_DIR) / f"{month_key}.json",
        ensure_dir(SUMMARIES_DIR) / f"{month_key}.json",
        ensure_dir(LEGACY_SUMMARIES_DIR) / f"{month_key}.json",
    ]


def _read_cached_summary(month_key: str):
    for cache_path in _cache_candidates(month_key):
        if cache_path.exists():
            return json.loads(cache_path.read_text(encoding="utf-8")), cache_path
    return None, None


def _persist_summary(summary: dict, month_key: str) -> None:
    summary_copy = deepcopy(summary)
    summary_copy["month"] = month_key
    canonical_path = ensure_dir(CANONICAL_MONTHLY_DIR) / f"{month_key}.json"
    canonical_path.write_text(json.dumps(summary_copy, ensure_ascii=False, indent=2), encoding="utf-8")
    primary_path = ensure_dir(SUMMARIES_DIR) / f"{month_key}.json"
    primary_path.write_text(json.dumps(summary_copy, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("event=summary_saved month=%s path=%s", month_key, primary_path)


def month_pdf_candidates(pdf_dir: Path, month_key: str) -> list[Path]:
    return [
        pdf_dir / f"{month_key}_dashboard.pdf",
        pdf_dir / f"{month_key}.pdf",
    ]


def resolve_month_pdf_path(month_key: str, pdf_dir: Path | None = None) -> Path:
    active_pdf_dir = Path(pdf_dir) if pdf_dir else INBOX_PDF_DIR
    for candidate in month_pdf_candidates(active_pdf_dir, month_key):
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        f"No existe PDF mensual para {month_key} en {active_pdf_dir}. "
        f"Se esperaba alguno de: {[str(path.name) for path in month_pdf_candidates(active_pdf_dir, month_key)]}"
    )


def resolve_period_month_pdfs(month_keys: list[str], pdf_dir: Path | None = None, allow_partial: bool = False) -> dict[str, Path]:
    active_pdf_dir = Path(pdf_dir) if pdf_dir else INBOX_PDF_DIR
    resolved: dict[str, Path] = {}
    missing: list[str] = []
    for month_key in month_keys:
        try:
            resolved[month_key] = resolve_month_pdf_path(month_key, active_pdf_dir)
        except FileNotFoundError:
            missing.append(month_key)
    if missing and not allow_partial:
        logger.error("event=missing_month_for_period pdf_dir=%s missing_months=%s", active_pdf_dir, missing)
        raise FileNotFoundError(f"Faltan PDFs para el período solicitado en {active_pdf_dir}: {', '.join(missing)}")
    return resolved


def summarize_month(_client: Any, month_key: str, force_regenerate: bool = False, pdf_dir: Path | None = None) -> dict:
    cached_summary, cached_path = _read_cached_summary(month_key)
    if cached_summary and not force_regenerate:
        logger.info("event=summary_cache_hit month=%s path=%s", month_key, cached_path)
        return cached_summary

    active_pdf_dir = Path(pdf_dir) if pdf_dir else INBOX_PDF_DIR
    if not active_pdf_dir.exists() and PDF_DIR.exists():
        active_pdf_dir = PDF_DIR

    pdf_path = resolve_month_pdf_path(month_key, active_pdf_dir)

    logger.info("event=processing_started month=%s pdf=%s mode=deterministic_pdf", month_key, pdf_path.name)
    try:
        raw_extracted = extract_raw_monthly_pdf(month_key, pdf_path)
        canonical = canonicalize_monthly(raw_extracted)
        validation = validate_canonical_monthly(canonical)
        canonical["validation"] = validation
        if not validation.get("is_valid", False):
            logger.error(
                "event=validation_failed month=%s errors=%s raw_metrics=%s",
                month_key,
                validation.get("errors", []),
                raw_extracted.get("metrics", {}),
            )
            raise ValueError(f"Resumen mensual inválido para {month_key}: {validation.get('errors', [])}")
        persist_monthly_artifacts(month_key, raw_extracted, canonical, validation)
        logger.info("event=raw_json_written month=%s path=%s", month_key, ensure_dir(RAW_EXTRACTED_DIR) / f"{month_key}.json")
        logger.info(
            "event=canonical_json_written month=%s path=%s",
            month_key,
            ensure_dir(CANONICAL_MONTHLY_DIR) / f"{month_key}.json",
        )
        logger.info(
            "event=validation_json_written month=%s path=%s",
            month_key,
            ensure_dir(VALIDATION_DIR) / f"{month_key}.json",
        )
        _persist_summary(canonical, month_key)
        return canonical
    except Exception as exc:
        if cached_summary:
            logger.warning(
                "event=summary_fallback_to_cache month=%s reason=%s cache_path=%s",
                month_key,
                exc,
                cached_path,
            )
            return cached_summary
        logger.warning("event=summary_local_fallback month=%s reason=%s", month_key, exc)
        fallback_summary = _build_local_fallback_summary(
            month_key,
            f"No se pudo generar resumen mensual determinístico. Se usó modo local_fallback. Error: {str(exc)}",
        )
        _persist_summary(fallback_summary, month_key)
        return fallback_summary
