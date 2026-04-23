from __future__ import annotations

import argparse
import json
import logging
import re
import unicodedata
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pypdf import PdfReader

from config import CANONICAL_MONTHLY_DIR, RAW_EXTRACTED_DIR, VALIDATION_DIR, ensure_dir
from metric_utils import normalize_percentage, to_float_locale

NUMBER_PATTERN = re.compile(r"-?\d+(?:[.,]\d{3})*(?:[.,]\d+)?%?")
MAX_MAIL_TO_PLAN_RATIO = 10
MIN_SITE_VIEWS_PER_NOTE = 10
MIN_MAIL_TO_PLAN_RELATION = 0.2
MIN_MAIL_ABSOLUTE = 10
MIN_RATE_DIFFERENCE = 0.01
logger = logging.getLogger(__name__)


ANCHOR_VARIANTS: dict[str, tuple[str, ...]] = {
    "Nº total de comunicaciones": (
        "Nº total de comunicaciones",
        "N° total de comunicaciones",
        "No total de comunicaciones",
        "N total de comunicaciones",
    ),
    "Media comunicaciones diarias": (
        "Media comunicaciones diarias",
    ),
    "Noticias Publicadas": (
        "Noticias Publicadas",
    ),
    "Total Páginas Vistas": (
        "Total Páginas Vistas",
        "Total Paginas Vistas",
    ),
    "Promedio Vistas": (
        "Promedio Vistas",
    ),
    "Mails enviados": (
        "Mails enviados",
        "Mails Enviados",
    ),
    "Tasa de apertura promedio": (
        "Tasa de apertura promedio",
    ),
    "Tasa de interacción sobre mails enviados": (
        "Tasa de interacción sobre mails enviados",
        "Tasa de interaccion sobre mails enviados",
    ),
    "Tasa de interacción sobre mails abiertos": (
        "Tasa de interacción sobre mails abiertos",
        "Tasa de interaccion sobre mails abiertos",
    ),
}


def _normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    simplified = without_accents.replace("º", "o").replace("°", "o")
    return " ".join(simplified.lower().split())


def _to_float_locale(raw: str) -> float:
    return to_float_locale(raw, 0.0)


def parse_integer_value(raw: str | None) -> int | None:
    if raw in (None, "", "-"):
        return None

    text = str(raw).strip()
    sign = -1 if text.startswith("-") else 1
    clean = re.sub(r"[^\d.,]", "", text)
    if not clean:
        return None

    def _is_thousands_format(candidate: str, separator: str) -> bool:
        parts = candidate.split(separator)
        if len(parts) <= 1:
            return False
        if not all(part.isdigit() for part in parts):
            return False
        if not (1 <= len(parts[0]) <= 3):
            return False
        return all(len(part) == 3 for part in parts[1:])

    if "." in clean and "," in clean:
        return sign * int(round(_to_float_locale(clean)))

    if "." in clean:
        if _is_thousands_format(clean, "."):
            return sign * int(clean.replace(".", ""))
        return sign * int(round(_to_float_locale(clean)))

    if "," in clean:
        if _is_thousands_format(clean, ","):
            return sign * int(clean.replace(",", ""))
        return sign * int(round(_to_float_locale(clean)))

    digits_only = re.sub(r"[^\d]", "", clean)
    if clean != digits_only:
        logger.debug("event=integer_parse_fallback raw=%s clean=%s", raw, clean)
    return sign * int(digits_only) if digits_only else None


def parse_percent_value(raw: str | None) -> float | None:
    if raw in (None, "", "-"):
        return None
    text = str(raw).strip()
    value = normalize_percentage(text)
    return round(value, 2)


def _extract_pages_text(pdf_path: Path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    pages: list[str] = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return pages


def _anchor_variants(anchor: str) -> tuple[str, ...]:
    variants = ANCHOR_VARIANTS.get(anchor)
    if variants:
        return variants
    return (anchor,)


def _extract_number_after_anchor(line: str, anchor: str) -> str | None:
    if not line:
        return None
    anchor_norm = _normalize_text(anchor)
    line_norm = _normalize_text(line)
    anchor_pos = line_norm.find(anchor_norm)
    if anchor_pos < 0:
        return None

    original_start = 0
    if anchor_pos > 0:
        prefix_norm = line_norm[:anchor_pos]
        prefix_tokens = prefix_norm.split()
        if prefix_tokens:
            consumed = 0
            tokens_seen = 0
            for token in line.split():
                token_norm = _normalize_text(token)
                if tokens_seen < len(prefix_tokens) and token_norm == prefix_tokens[tokens_seen]:
                    tokens_seen += 1
                consumed += len(token) + 1
                if tokens_seen >= len(prefix_tokens):
                    break
            original_start = min(consumed, len(line))

    anchor_match = re.search(re.escape(anchor), line[original_start:], flags=re.IGNORECASE)
    if anchor_match:
        search_start = original_start + anchor_match.end()
    else:
        anchor_tokens = anchor.split()
        search_start = original_start
        if anchor_tokens:
            matches_seen = 0
            for token in line.split():
                token_norm = _normalize_text(token)
                if matches_seen < len(anchor_tokens) and token_norm == _normalize_text(anchor_tokens[matches_seen]):
                    matches_seen += 1
                search_start += len(token) + 1
                if matches_seen >= len(anchor_tokens):
                    break

    search_segment = line[search_start:]
    match = NUMBER_PATTERN.search(search_segment)
    return match.group(0) if match else None


def find_anchor_value(page_text: str, anchor: str, kind: str, page_number: int) -> dict[str, Any]:
    lines = page_text.splitlines()
    anchor_variants = _anchor_variants(anchor)
    for idx, line in enumerate(lines):
        line_norm = _normalize_text(line)
        matched_anchor = next((variant for variant in anchor_variants if _normalize_text(variant) in line_norm), None)
        if not matched_anchor:
            continue

        raw_value = _extract_number_after_anchor(line, matched_anchor)
        if raw_value is None:
            for offset in (1, 2):
                if idx + offset >= len(lines):
                    break
                next_line = lines[idx + offset]
                next_candidates = NUMBER_PATTERN.findall(next_line)
                if next_candidates:
                    raw_value = next_candidates[0]
                    break
        if raw_value is None:
            continue

        if kind == "percent":
            value = parse_percent_value(raw_value)
            unit = "percent"
        else:
            if "%" in raw_value:
                continue
            parsed_int = parse_integer_value(raw_value)
            value = float(parsed_int) if parsed_int is not None else None
            unit = "count"
        if value is None:
            continue

        return {
            "anchor": anchor,
            "raw_value": raw_value,
            "value": round(value, 2),
            "unit": unit,
            "page": page_number,
            "line": " ".join(line.split())[:220],
            "missing": False,
        }

    return {
        "anchor": anchor,
        "raw_value": None,
        "value": None,
        "unit": "percent" if kind == "percent" else "count",
        "page": page_number,
        "line": "",
        "missing": True,
    }


def extract_planning_page_metrics(page_text: str, page_number: int = 1) -> dict[str, dict[str, Any]]:
    return {
        "plan_total": find_anchor_value(page_text, "Nº total de comunicaciones", "count", page_number),
        "plan_daily_average": find_anchor_value(page_text, "Media comunicaciones diarias", "count", page_number),
    }


def extract_site_page_metrics(page_text: str, page_number: int = 2) -> dict[str, dict[str, Any]]:
    return {
        "site_notes_total": find_anchor_value(page_text, "Noticias Publicadas", "count", page_number),
        "site_total_views": find_anchor_value(page_text, "Total Páginas Vistas", "count", page_number),
        "site_average_views": find_anchor_value(page_text, "Promedio Vistas", "count", page_number),
    }


def extract_mail_page_metrics(page_text: str, page_number: int = 3) -> dict[str, dict[str, Any]]:
    return {
        "mail_total": find_anchor_value(page_text, "Mails enviados", "count", page_number),
        "mail_open_rate": find_anchor_value(page_text, "Tasa de apertura promedio", "percent", page_number),
        "mail_interaction_rate": find_anchor_value(
            page_text,
            "Tasa de interacción sobre mails enviados",
            "percent",
            page_number,
        ),
        "mail_interaction_rate_over_opened": find_anchor_value(
            page_text,
            "Tasa de interacción sobre mails abiertos",
            "percent",
            page_number,
        ),
    }


def _extract_metric_from_pages(
    pages: list[str],
    anchor: str,
    kind: str,
    expected_page_number: int,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    if 1 <= expected_page_number <= len(pages):
        preferred = find_anchor_value(pages[expected_page_number - 1], anchor, kind, expected_page_number)
        if not preferred.get("missing", False):
            return preferred, warnings

    for actual_page_number, page_text in enumerate(pages, start=1):
        if actual_page_number == expected_page_number:
            continue
        candidate = find_anchor_value(page_text, anchor, kind, actual_page_number)
        if candidate.get("missing", False):
            continue
        warnings.append(
            f"anchor_out_of_expected_page:{anchor}:Se encontró en página {actual_page_number} y no en la página esperada {expected_page_number}"
        )
        return candidate, warnings

    missing = {
        "anchor": anchor,
        "raw_value": None,
        "value": None,
        "unit": "percent" if kind == "percent" else "count",
        "page": expected_page_number,
        "line": "",
        "missing": True,
    }
    warnings.append(
        f"missing_anchor:No se encontró ancla exacta '{anchor}' en página esperada {expected_page_number} ni en el resto del documento"
    )
    return missing, warnings


def extract_raw_monthly_pdf(month_key: str, pdf_path: Path) -> dict[str, Any]:
    pages = _extract_pages_text(pdf_path)
    metric_specs = [
        ("plan_total", "Nº total de comunicaciones", "count", 1),
        ("plan_daily_average", "Media comunicaciones diarias", "count", 1),
        ("site_notes_total", "Noticias Publicadas", "count", 2),
        ("site_total_views", "Total Páginas Vistas", "count", 2),
        ("site_average_views", "Promedio Vistas", "count", 2),
        ("mail_total", "Mails enviados", "count", 3),
        ("mail_open_rate", "Tasa de apertura promedio", "percent", 3),
        ("mail_interaction_rate", "Tasa de interacción sobre mails enviados", "percent", 3),
        ("mail_interaction_rate_over_opened", "Tasa de interacción sobre mails abiertos", "percent", 3),
    ]

    metrics: dict[str, Any] = {}
    warnings: list[str] = []
    for metric_key, anchor, kind, expected_page in metric_specs:
        extracted, metric_warnings = _extract_metric_from_pages(pages, anchor, kind, expected_page)
        metrics[metric_key] = extracted
        if extracted.get("missing", False):
            warnings.append(f"missing_anchor:{metric_key}:{metric_warnings[-1]}")
        else:
            warnings.extend(metric_warnings)

    return {
        "month": month_key,
        "source_pdf": str(pdf_path),
        "extracted_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "parser": "deterministic_pdf_v3",
        "page_count": len(pages),
        "metrics": metrics,
        "warnings": warnings,
    }


def canonicalize_monthly(raw_extracted: dict[str, Any]) -> dict[str, Any]:
    metrics = raw_extracted.get("metrics", {})
    open_rate = float(metrics.get("mail_open_rate", {}).get("value") or 0.0)
    interaction_rate = float(metrics.get("mail_interaction_rate", {}).get("value") or 0.0)

    return {
        "month": raw_extracted.get("month"),
        "generation_mode": "deterministic_pdf",
        "extraction_method": raw_extracted.get("parser"),
        "plan_total": int(round(metrics.get("plan_total", {}).get("value") or 0)),
        "site_notes_total": int(round(metrics.get("site_notes_total", {}).get("value") or 0)),
        "site_total_views": int(round(metrics.get("site_total_views", {}).get("value") or 0)),
        "mail_total": int(round(metrics.get("mail_total", {}).get("value") or 0)),
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
    extraction_warnings = list(canonical.get("extraction_warnings", []))
    if any(str(w).startswith("missing_anchor:") for w in extraction_warnings):
        errors.append("Faltan KPIs primarios por ancla exacta")

    for metric in ("mail_open_rate", "mail_interaction_rate"):
        value = float(canonical.get(metric, 0))
        if value < 0 or value > 100:
            errors.append(f"{metric} fuera de rango 0-100")
        elif value < 1:
            warnings.append(f"{metric} es menor a 1%; revisar escala")

    if canonical.get("mail_total", 0) > canonical.get("plan_total", 0) * MAX_MAIL_TO_PLAN_RATIO and canonical.get("plan_total", 0) > 0:
        warnings.append("mail_total luce desproporcionado respecto a plan_total")

    plan_total = int(canonical.get("plan_total", 0) or 0)
    site_notes_total = int(canonical.get("site_notes_total", 0) or 0)
    site_total_views = int(canonical.get("site_total_views", 0) or 0)
    mail_total = int(canonical.get("mail_total", 0) or 0)
    open_rate = float(canonical.get("mail_open_rate", 0) or 0)
    interaction_rate = float(canonical.get("mail_interaction_rate", 0) or 0)

    if plan_total <= 0:
        errors.append("plan_total inválido: debe ser mayor a 0")
    if site_notes_total < 0:
        errors.append("site_notes_total no puede ser negativo")
    if site_total_views < 0:
        errors.append("site_total_views no puede ser negativo")
    if mail_total < 0:
        errors.append("mail_total no puede ser negativo")
    if mail_total >= 0 and plan_total > 0 and mail_total < max(MIN_MAIL_ABSOLUTE, int(round(plan_total * MIN_MAIL_TO_PLAN_RELATION))):
        errors.append("mail_total sospechosamente bajo respecto a plan_total")
    if site_notes_total > 0 and site_total_views < site_notes_total * MIN_SITE_VIEWS_PER_NOTE:
        errors.append("site_total_views sospechosamente bajo respecto a site_notes_total")
    if open_rate > 0 and interaction_rate > 0 and abs(open_rate - interaction_rate) < MIN_RATE_DIFFERENCE:
        errors.append("mail_open_rate y mail_interaction_rate no deberían colapsar al mismo valor")

    return {
        "month": canonical.get("month"),
        "validated_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "is_valid": not errors,
        "errors": errors,
        "warnings": warnings + extraction_warnings,
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


def infer_month_key_from_pdf_path(pdf_path: Path) -> str:
    match = re.search(r"(20\d{2}-\d{2})", pdf_path.name)
    if match:
        return match.group(1)
    return datetime.now(UTC).strftime("%Y-%m")


def extract_single_pdf_to_raw(input_pdf: Path, output_json: Path, month_key: str | None = None) -> dict[str, Any]:
    resolved_month = month_key or infer_month_key_from_pdf_path(input_pdf)
    raw_extracted = extract_raw_monthly_pdf(resolved_month, input_pdf)
    ensure_dir(output_json.parent)
    output_json.write_text(json.dumps(raw_extracted, ensure_ascii=False, indent=2), encoding="utf-8")
    return raw_extracted


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Extracción determinística de dashboard PDF.")
    parser.add_argument("--input", type=Path, required=True, help="Ruta local al PDF de entrada.")
    parser.add_argument("--output", type=Path, required=True, help="Ruta local del JSON raw de salida.")
    parser.add_argument("--month", default=None, help="Mes explícito YYYY-MM para el JSON raw.")
    args = parser.parse_args(argv)
    raw = extract_single_pdf_to_raw(args.input, args.output, args.month)
    print(json.dumps({"status": "ok", "month": raw.get("month"), "output": str(args.output)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
