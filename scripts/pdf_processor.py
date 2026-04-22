import json
import logging
from copy import deepcopy
from pathlib import Path
from typing import Any

from config import CANONICAL_MONTHLY_DIR, LEGACY_SUMMARIES_DIR, PDF_DIR, SUMMARIES_DIR, ensure_dir
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


def summarize_month(_client: Any, month_key: str, force_regenerate: bool = False) -> dict:
    cached_summary, cached_path = _read_cached_summary(month_key)
    if cached_summary and not force_regenerate:
        logger.info("event=summary_cache_hit month=%s path=%s", month_key, cached_path)
        return cached_summary

    pdf_path = PDF_DIR / f"{month_key}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"No existe el PDF mensual esperado: {pdf_path}")

    logger.info("event=summary_start month=%s pdf=%s mode=deterministic_pdf", month_key, pdf_path.name)
    try:
        raw_extracted = extract_raw_monthly_pdf(month_key, pdf_path)
        canonical = canonicalize_monthly(raw_extracted)
        validation = validate_canonical_monthly(canonical)
        canonical["validation"] = validation
        if not validation.get("is_valid", False):
            raise ValueError(f"Resumen mensual inválido para {month_key}: {validation.get('errors', [])}")
        persist_monthly_artifacts(month_key, raw_extracted, canonical, validation)
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
