from __future__ import annotations

import json
import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pypdf import PdfReader

from config import CANONICAL_MONTHLY_DIR, RAW_EXTRACTED_DIR, VALIDATION_DIR, ensure_dir
from metric_utils import normalize_percentage, to_float_locale

NUMBER_PATTERN = re.compile(r"-?\d+(?:[.,]\d+)?%?")
MAX_MAIL_TO_PLAN_RATIO = 10

METRIC_SPECS = {
    "plan_total": {"keywords": ("plan", "planificación", "comunicaciones"), "kind": "count"},
    "site_notes_total": {"keywords": ("notas", "noticias", "publicadas"), "kind": "count"},
    "site_total_views": {"keywords": ("views", "lecturas", "páginas vistas"), "kind": "count"},
    "mail_total": {"keywords": ("mail", "envíos", "push"), "kind": "count"},
    "mail_open_rate": {"keywords": ("apertura", "open rate", "tasa apertura"), "kind": "percent"},
    "mail_interaction_rate": {"keywords": ("interacción", "ctr", "click rate"), "kind": "percent"},
}


def _to_float_locale(raw: str) -> float:
    return to_float_locale(raw, 0.0)


def _to_int_locale(raw: str) -> int:
    return int(round(_to_float_locale(raw)))


def _extract_pages_text(pdf_path: Path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    pages: list[str] = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return pages


def _find_metric(pages: list[str], keywords: tuple[str, ...], kind: str) -> dict[str, Any]:
    best: dict[str, Any] | None = None
    for page_idx, page_text in enumerate(pages, start=1):
        for line in page_text.splitlines():
            line_norm = line.lower()
            if not any(keyword in line_norm for keyword in keywords):
                continue
            candidates = NUMBER_PATTERN.findall(line)
            if not candidates:
                continue
            for raw_number in candidates:
                has_pct = "%" in raw_number
                if kind == "percent":
                    value = normalize_percentage(raw_number)
                    unit = "percent"
                else:
                    if has_pct:
                        continue
                    value = float(_to_int_locale(raw_number))
                    unit = "count"
                confidence = 0
                confidence += 2 if any(keyword in line_norm for keyword in keywords) else 0
                confidence += 1 if has_pct and kind == "percent" else 0
                row = {
                    "raw_value": raw_number,
                    "value": round(value, 2),
                    "unit": unit,
                    "page": page_idx,
                    "line": " ".join(line.split())[:220],
                    "confidence": confidence,
                }
                if best is None or row["confidence"] > best["confidence"]:
                    best = row
    if best is None:
        return {
            "raw_value": None,
            "value": 0.0 if kind == "percent" else 0,
            "unit": "percent" if kind == "percent" else "count",
            "page": None,
            "line": "",
            "confidence": 0,
            "missing": True,
        }
    return best


def extract_raw_monthly_pdf(month_key: str, pdf_path: Path) -> dict[str, Any]:
    pages = _extract_pages_text(pdf_path)
    metrics: dict[str, Any] = {}
    warnings: list[str] = []
    for metric, spec in METRIC_SPECS.items():
        extracted = _find_metric(pages, spec["keywords"], spec["kind"])
        metrics[metric] = extracted
        if extracted.get("missing"):
            warnings.append(f"No se pudo extraer {metric} de forma determinística")

    return {
        "month": month_key,
        "source_pdf": str(pdf_path),
        "extracted_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "parser": "deterministic_pdf_v1",
        "page_count": len(pages),
        "metrics": metrics,
        "warnings": warnings,
    }


def canonicalize_monthly(raw_extracted: dict[str, Any]) -> dict[str, Any]:
    metrics = raw_extracted.get("metrics", {})
    open_rate = float(metrics.get("mail_open_rate", {}).get("value", 0.0))
    interaction_rate = float(metrics.get("mail_interaction_rate", {}).get("value", 0.0))

    return {
        "month": raw_extracted.get("month"),
        "generation_mode": "deterministic_pdf",
        "extraction_method": raw_extracted.get("parser"),
        "plan_total": int(round(metrics.get("plan_total", {}).get("value", 0))),
        "site_notes_total": int(round(metrics.get("site_notes_total", {}).get("value", 0))),
        "site_total_views": int(round(metrics.get("site_total_views", {}).get("value", 0))),
        "mail_total": int(round(metrics.get("mail_total", {}).get("value", 0))),
        "mail_open_rate": round(open_rate, 2),
        "mail_interaction_rate": round(interaction_rate, 2),
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
            "site_has_no_data_sections": False,
            "events_summary_available": False,
            "push_ranking_available": False,
            "pull_ranking_available": False,
            "historical_comparison_allowed": True,
        },
        "extraction_warnings": raw_extracted.get("warnings", []),
    }


def validate_canonical_monthly(canonical: dict[str, Any]) -> dict[str, Any]:
    warnings: list[str] = []
    errors: list[str] = []
    for metric in ("plan_total", "site_notes_total", "site_total_views", "mail_total"):
        if canonical.get(metric, 0) < 0:
            errors.append(f"{metric} no puede ser negativo")

    for metric in ("mail_open_rate", "mail_interaction_rate"):
        value = float(canonical.get(metric, 0))
        if value < 0 or value > 100:
            errors.append(f"{metric} fuera de rango 0-100")
        elif value < 1:
            warnings.append(f"{metric} es menor a 1%; revisar escala")

    if canonical.get("mail_total", 0) > canonical.get("plan_total", 0) * MAX_MAIL_TO_PLAN_RATIO and canonical.get("plan_total", 0) > 0:
        warnings.append("mail_total luce desproporcionado respecto a plan_total")

    return {
        "month": canonical.get("month"),
        "validated_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "is_valid": not errors,
        "errors": errors,
        "warnings": warnings + list(canonical.get("extraction_warnings", [])),
    }


def persist_monthly_artifacts(month_key: str, raw_extracted: dict[str, Any], canonical: dict[str, Any], validation: dict[str, Any]) -> None:
    ensure_dir(RAW_EXTRACTED_DIR).joinpath(f"{month_key}.json").write_text(
        json.dumps(raw_extracted, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    ensure_dir(CANONICAL_MONTHLY_DIR).joinpath(f"{month_key}.json").write_text(
        json.dumps(canonical, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    ensure_dir(VALIDATION_DIR).joinpath(f"{month_key}.json").write_text(
        json.dumps(validation, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
