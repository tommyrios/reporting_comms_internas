from __future__ import annotations

import json
import logging
from copy import deepcopy
from pathlib import Path
from typing import Any

from config import (
    CANONICAL_PERIOD_DIR,
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
from period_scopes import (
    REQUIRED_PERIOD_SCOPES,
    SCOPE_FILE_TOKENS,
    SCOPE_LABELS,
    period_scope_filename,
)

logger = logging.getLogger(__name__)


def _build_local_fallback_summary(period_slug: str, scope: str, warning: str) -> dict:
    return {
        "period": period_slug,
        "month": period_slug,  # compat con analyzer existente
        "scope": scope,
        "scope_label": SCOPE_LABELS.get(scope, scope),
        "generation_mode": "local_fallback",
        "warning": warning,
        "plan_daily_average": 0,
        "plan_total": 0,
        "site_notes_total": 0,
        "site_total_views": 0,
        "site_average_views": 0,
        "mail_total": 0,
        "mail_open_rate": 0,
        "mail_interaction_rate": 0,
        "mail_interaction_rate_over_opened": 0,
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
            "scope_country": scope,
            "scope_mixed": scope == "combined",
            "site_has_no_data_sections": True,
            "events_summary_available": False,
            "push_ranking_available": False,
            "pull_ranking_available": False,
            "historical_comparison_allowed": False,
        },
    }


def _cache_candidates(period_slug: str, scope: str) -> list[Path]:
    filename = f"{period_slug}_{scope}.json"
    return [
        ensure_dir(CANONICAL_PERIOD_DIR) / filename,
        ensure_dir(SUMMARIES_DIR) / filename,
        ensure_dir(LEGACY_SUMMARIES_DIR) / filename,
    ]


def _read_cached_summary(period_slug: str, scope: str):
    for cache_path in _cache_candidates(period_slug, scope):
        if cache_path.exists():
            return json.loads(cache_path.read_text(encoding="utf-8")), cache_path
    return None, None


def _persist_summary(summary: dict, period_slug: str, scope: str, source_pdf: Path | None = None) -> None:
    summary_copy = deepcopy(summary)
    summary_copy["period"] = period_slug
    summary_copy["month"] = period_slug  # compat con analyzer existente
    summary_copy["scope"] = scope
    summary_copy["scope_label"] = SCOPE_LABELS.get(scope, scope)
    if source_pdf is not None:
        summary_copy["source_pdf"] = str(Path(source_pdf).resolve())
        summary_copy["source_pdf_name"] = Path(source_pdf).name
    filename = f"{period_slug}_{scope}.json"
    canonical_path = ensure_dir(CANONICAL_PERIOD_DIR) / filename
    canonical_path.write_text(json.dumps(summary_copy, ensure_ascii=False, indent=2), encoding="utf-8")
    primary_path = ensure_dir(SUMMARIES_DIR) / filename
    primary_path.write_text(json.dumps(summary_copy, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("event=period_summary_saved period=%s scope=%s path=%s", period_slug, scope, primary_path)


def period_scope_pdf_candidates(pdf_dir: Path, period_slug: str, scope: str) -> list[Path]:
    token = SCOPE_FILE_TOKENS.get(scope, scope.upper())
    return [
        pdf_dir / period_scope_filename(period_slug, scope),
        pdf_dir / f"{period_slug}_{scope}.pdf",
        pdf_dir / f"{period_slug}_{token}.pdf",
        pdf_dir / f"{period_slug}_dashboard_{token}.pdf",
        pdf_dir / f"dashboard_{period_slug}_{token}.pdf",
    ]


def resolve_period_scope_pdf_path(period_slug: str, scope: str, pdf_dir: Path | None = None) -> Path:
    active_pdf_dir = Path(pdf_dir) if pdf_dir else INBOX_PDF_DIR
    for candidate in period_scope_pdf_candidates(active_pdf_dir, period_slug, scope):
        if candidate.exists():
            return candidate
    expected = [str(path.name) for path in period_scope_pdf_candidates(active_pdf_dir, period_slug, scope)]
    raise FileNotFoundError(
        f"No existe PDF para {period_slug} [{scope}] en {active_pdf_dir}. "
        f"Se esperaba alguno de: {expected}"
    )


def resolve_period_scope_pdfs(
    period_slug: str,
    scopes: list[str] | None = None,
    pdf_dir: Path | None = None,
    allow_partial: bool = False,
) -> dict[str, Path]:
    active_pdf_dir = Path(pdf_dir) if pdf_dir else INBOX_PDF_DIR
    scopes = scopes or list(REQUIRED_PERIOD_SCOPES)
    resolved: dict[str, Path] = {}
    missing: list[str] = []
    for scope in scopes:
        try:
            resolved[scope] = resolve_period_scope_pdf_path(period_slug, scope, active_pdf_dir)
        except FileNotFoundError:
            missing.append(scope)
    if missing and not allow_partial:
        logger.error("event=missing_scope_for_period pdf_dir=%s period=%s missing_scopes=%s", active_pdf_dir, period_slug, missing)
        raise FileNotFoundError(f"Faltan PDFs para {period_slug} en {active_pdf_dir}: {', '.join(missing)}")
    return resolved


def summarize_period_scope(
    period: dict[str, Any],
    scope: str,
    force_regenerate: bool = False,
    pdf_dir: Path | None = None,
) -> dict:
    period_slug = str(period.get("slug") or period.get("period") or "").strip()
    if not period_slug:
        raise ValueError("El período no tiene slug")

    active_pdf_dir = Path(pdf_dir) if pdf_dir else INBOX_PDF_DIR
    if not active_pdf_dir.exists() and PDF_DIR.exists():
        active_pdf_dir = PDF_DIR

    pdf_path = resolve_period_scope_pdf_path(period_slug, scope, active_pdf_dir)

    cached_summary, cached_path = _read_cached_summary(period_slug, scope)
    expected_source = str(pdf_path.resolve())
    cached_source = str(cached_summary.get("source_pdf", "")) if isinstance(cached_summary, dict) else ""
    cached_name = str(cached_summary.get("source_pdf_name", "")) if isinstance(cached_summary, dict) else ""
    cache_matches_pdf = cached_source == expected_source or (not cached_source and cached_name == pdf_path.name)
    if cached_summary and not force_regenerate and cache_matches_pdf:
        logger.info("event=period_summary_cache_hit period=%s scope=%s path=%s source_pdf=%s", period_slug, scope, cached_path, pdf_path.name)
        return cached_summary
    if cached_summary and not force_regenerate and not cache_matches_pdf:
        logger.warning(
            "event=period_summary_cache_ignored period=%s scope=%s path=%s cached_source=%s expected_source=%s",
            period_slug,
            scope,
            cached_path,
            cached_source or cached_name or "<missing>",
            expected_source,
        )
    extraction_key = f"{period_slug}_{scope}"

    logger.info(
        "event=processing_started period=%s scope=%s pdf=%s mode=deterministic_pdf",
        period_slug,
        scope,
        pdf_path.name,
    )
    try:
        raw_extracted = extract_raw_monthly_pdf(extraction_key, pdf_path)
        raw_extracted["period"] = period_slug
        raw_extracted["scope"] = scope
        raw_extracted["scope_label"] = SCOPE_LABELS.get(scope, scope)
        canonical = canonicalize_monthly(raw_extracted)
        canonical["period"] = period_slug
        canonical["month"] = period_slug  # compat con analyzer existente
        canonical["scope"] = scope
        canonical["scope_label"] = SCOPE_LABELS.get(scope, scope)
        canonical.setdefault("quality_flags", {})["scope_country"] = scope
        canonical.setdefault("quality_flags", {})["scope_mixed"] = scope == "combined"
        canonical.setdefault("quality_flags", {})["historical_comparison_allowed"] = False

        validation = validate_canonical_monthly(canonical)
        canonical["validation"] = validation
        if not validation.get("is_valid", False):
            logger.error(
                "event=validation_failed period=%s scope=%s errors=%s raw_metrics=%s",
                period_slug,
                scope,
                validation.get("errors", []),
                raw_extracted.get("metrics", {}),
            )
            raise ValueError(f"Resumen inválido para {period_slug} [{scope}]: {validation.get('errors', [])}")

        # Persistimos artefactos crudos con key period_scope para no mezclar scopes.
        persist_monthly_artifacts(extraction_key, raw_extracted, canonical, validation)
        _persist_summary(canonical, period_slug, scope, source_pdf=pdf_path)
        return canonical
    except Exception as exc:
        if cached_summary and cache_matches_pdf:
            logger.warning(
                "event=period_summary_fallback_to_cache period=%s scope=%s reason=%s cache_path=%s",
                period_slug,
                scope,
                exc,
                cached_path,
            )
            return cached_summary
        logger.warning("event=period_summary_local_fallback period=%s scope=%s reason=%s", period_slug, scope, exc)
        fallback_summary = _build_local_fallback_summary(
            period_slug,
            scope,
            f"No se pudo generar resumen determinístico. Se usó modo local_fallback. Error: {str(exc)}",
        )
        _persist_summary(fallback_summary, period_slug, scope, source_pdf=pdf_path)
        return fallback_summary
