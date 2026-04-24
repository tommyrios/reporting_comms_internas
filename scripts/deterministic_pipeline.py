from __future__ import annotations

import json
import logging
import re
import unicodedata
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pypdf import PdfReader

logger = logging.getLogger(__name__)

NUMBER_PATTERN = re.compile(r"-?\d+(?:[.,]\d{3})*(?:[.,]\d+)?%?")

MAX_MAIL_TO_PLAN_RATIO = 10
MIN_SITE_VIEWS_PER_NOTE = 10
MIN_MAIL_TO_PLAN_RELATION = 0.2
MIN_MAIL_ABSOLUTE = 10
MIN_RATE_DIFFERENCE = 0.01


def to_float_locale(raw: str | None, default: float = 0.0) -> float:
    if raw is None:
        return default

    text = str(raw).strip()
    text = text.replace("%", "")
    text = re.sub(r"[^\d,.\-]", "", text)

    if not text:
        return default

    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
    elif "." in text:
        parts = text.split(".")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            text = text.replace(".", "")

    try:
        return float(text)
    except Exception:
        return default


def parse_integer_value(raw: str | None) -> int | None:
    if raw in (None, "", "-"):
        return None

    value = to_float_locale(raw, 0.0)
    return int(round(value))


def parse_percent_value(raw: str | None) -> float | None:
    if raw in (None, "", "-"):
        return None

    text = str(raw).strip()
    value = to_float_locale(text, 0.0)

    if "%" not in text and 0 < value <= 1:
        value *= 100

    return round(value, 2)


# -------------------------
# PDF helpers
# -------------------------

def _normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    simplified = without_accents.replace("º", "o").replace("°", "o")
    return " ".join(simplified.lower().split())


def _extract_pages_text(pdf_path: Path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    return [page.extract_text() or "" for page in reader.pages]


def _value_immediately_after_label(page_text: str, label: str, kind: str) -> str | None:
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]
    label_norm = _normalize_text(label)

    matched_indexes = [
        i for i, line in enumerate(lines)
        if _normalize_text(line) == label_norm
    ]

    for i in reversed(matched_indexes):
        if i + 1 >= len(lines):
            continue

        next_line = lines[i + 1]
        nums = NUMBER_PATTERN.findall(next_line)

        if kind == "percent":
            nums = [n for n in nums if "%" in n]
        else:
            nums = [n for n in nums if "%" not in n]

        if nums:
            return nums[0]

    return None


def _metric(anchor: str, raw_value: str | None, kind: str, page: int) -> dict[str, Any]:
    if raw_value is None:
        return {
            "anchor": anchor,
            "raw_value": None,
            "value": None,
            "unit": "percent" if kind == "percent" else "number" if kind == "float" else "count",
            "page": page,
            "line": "",
            "missing": True,
        }

    if kind == "percent":
        value = parse_percent_value(raw_value)
        unit = "percent"
    elif kind == "float":
        value = round(to_float_locale(raw_value), 2)
        unit = "number"
    else:
        value = parse_integer_value(raw_value)
        unit = "count"

    return {
        "anchor": anchor,
        "raw_value": raw_value,
        "value": value,
        "unit": unit,
        "page": page,
        "line": "",
        "missing": value is None,
    }

def _extract_mail_table(page_text: str) -> list[dict[str, Any]]:
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]
    rows: list[dict[str, Any]] = []

    for line in lines:
        if not re.match(r"^[A-Z][a-z]{2}\s\d{1,2},\s20\d{2}", line):
            continue

        percents = re.findall(r"\d+(?:[.,]\d+)?%", line)

        if len(percents) < 3:
            continue

        open_rate = parse_percent_value(percents[-3])
        ctr = parse_percent_value(percents[-2])
        ctor = parse_percent_value(percents[-1])

        date_match = re.match(r"^([A-Z][a-z]{2}\s\d{1,2},\s20\d{2})\s+(.*)$", line)
        date = date_match.group(1) if date_match else None
        rest = date_match.group(2) if date_match else line

        body = rest
        for pct in percents[-3:]:
            body = body.replace(pct, "")

        metric_match = re.search(
            r"\s+Argentina\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s*$",
            body,
            flags=re.IGNORECASE,
        )

        sent = opens = clicks = None
        title = body.strip()

        if metric_match:
            sent = parse_integer_value(metric_match.group(1))
            opens = parse_integer_value(metric_match.group(2))
            clicks = parse_integer_value(metric_match.group(3))
            title = body[:metric_match.start()].strip()

        title = re.sub(r"\s+", " ", title).strip()

        rows.append({
            "date": date,
            "title": title[:180],
            "sent": sent,
            "opens": opens,
            "clicks": clicks,
            "open_rate": open_rate,
            "ctr": ctr,
            "ctor": ctor,
            "raw": line,
        })

    return rows


def _build_push_rankings(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    clean_rows = [
        r for r in rows
        if r.get("open_rate") is not None
        and r.get("ctr") is not None
        and (r.get("sent") or 0) >= 1000
    ]

    top_open = sorted(clean_rows, key=lambda x: x["open_rate"], reverse=True)[:5]
    top_interaction = sorted(clean_rows, key=lambda x: x["ctr"], reverse=True)[:5]

    return top_open, top_interaction

def _extract_top_pull_notes(page_text: str) -> list[dict[str, Any]]:
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]
    rows: list[dict[str, Any]] = []

    capture = False

    for line in lines:
        norm = _normalize_text(line)

        if "top five - notas mas leidas (uu)" in norm:
            capture = True
            continue

        if capture and "top five - notas mas leidas (colectivo tgm)" in norm:
            break

        if not capture:
            continue

        date_match = re.match(
            r"^([A-Z][a-z]{2}\s\d{1,2},\s(?:20\d{2}|20…))\s+(.*)$",
            line,
        )

        if not date_match:
            continue

        date = date_match.group(1)
        body = date_match.group(2)

        nums = NUMBER_PATTERN.findall(line)
        nums_no_percent = [n for n in nums if "%" not in n]

        if len(nums_no_percent) < 3:
            continue

        users = parse_integer_value(nums_no_percent[-2])
        views = parse_integer_value(nums_no_percent[-1])

        title = body
        title = re.sub(r"\s+ARGENTINA\s+[\d,]+\s+[\d,]+\s*$", "", title)
        title = re.sub(r"\s+", " ", title).strip()

        rows.append({
            "date": date.replace("20…", "2026"),
            "title": title[:180],
            "users": users,
            "views": views,
            "raw": line,
        })

    return rows[:5]

# -------------------------
# Extracción principal
# -------------------------

def extract_raw_monthly_pdf(month_key: str, pdf_path: Path) -> dict[str, Any]:
    pages = _extract_pages_text(pdf_path)

    if len(pages) < 3:
        raise ValueError(f"El PDF debería tener al menos 3 páginas. Tiene {len(pages)}.")

    p1 = pages[0]
    p2 = pages[1]
    p3 = pages[2]

    metrics = {
        "plan_daily_average": _metric(
            "Media comunicaciones diarias",
            _value_immediately_after_label(p1, "Media comunicaciones diarias", "float"),
            "float",
            1,
        ),
        "plan_total": _metric(
            "Nº total de comunicaciones",
            _value_immediately_after_label(p1, "Nº total de comunicaciones", "count"),
            "count",
            1,
        ),
        "site_total_views": _metric(
            "Total Páginas Vistas",
            _value_immediately_after_label(p2, "Total Páginas Vistas", "count"),
            "count",
            2,
        ),
        "site_notes_total": _metric(
            "Noticias Publicadas",
            _value_immediately_after_label(p2, "Noticias Publicadas", "count"),
            "count",
            2,
        ),
        "site_average_views": _metric(
            "Promedio Vistas",
            _value_immediately_after_label(p2, "Promedio Vistas", "count"),
            "count",
            2,
        ),
        "mail_open_rate": _metric(
            "Tasa de apertura promedio",
            _value_immediately_after_label(p3, "Tasa de apertura promedio", "percent"),
            "percent",
            3,
        ),
        "mail_interaction_rate": _metric(
            "Tasa de interacción sobre mails enviados",
            _value_immediately_after_label(p3, "Tasa de interacción sobre mails enviados", "percent"),
            "percent",
            3,
        ),
        "mail_interaction_rate_over_opened": _metric(
            "Tasa de interacción sobre mails abiertos",
            _value_immediately_after_label(p3, "Tasa de interacción sobre mails abiertos", "percent"),
            "percent",
            3,
        ),
        "mail_total": _metric(
            "Mails enviados",
            _value_immediately_after_label(p3, "Mails enviados", "count"),
            "count",
            3,
        ),
    }

    mail_rows = _extract_mail_table(p3)
    top_push_open, top_push_interaction = _build_push_rankings(mail_rows)
    top_pull_notes = _extract_top_pull_notes(p2)

    warnings = [
        f"missing_anchor:{k}:{v.get('anchor')}"
        for k, v in metrics.items()
        if v.get("missing")
    ]

    return {
        "month": month_key,
        "source_pdf": str(pdf_path),
        "extracted_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "parser": "deterministic_pdf_v4_exact_next_line",
        "page_count": len(pages),
        "metrics": metrics,
        "mail_table": mail_rows,
        "top_push_open": top_push_open,
        "top_push_interaction": top_push_interaction,
        "top_pull_notes": top_pull_notes,
        "warnings": warnings,
    }


# -------------------------
# Canonical
# -------------------------

def canonicalize_monthly(raw_extracted: dict[str, Any]) -> dict[str, Any]:
    metrics = raw_extracted.get("metrics", {})

    open_rate = float(metrics.get("mail_open_rate", {}).get("value") or 0.0)
    interaction_rate = float(metrics.get("mail_interaction_rate", {}).get("value") or 0.0)

    return {
        "month": raw_extracted.get("month"),
        "generation_mode": "deterministic_pdf",
        "extraction_method": raw_extracted.get("parser"),

        "plan_daily_average": float(metrics.get("plan_daily_average", {}).get("value") or 0.0),
        "plan_total": int(round(metrics.get("plan_total", {}).get("value") or 0)),
        "site_notes_total": int(round(metrics.get("site_notes_total", {}).get("value") or 0)),
        "site_total_views": int(round(metrics.get("site_total_views", {}).get("value") or 0)),
        "site_average_views": int(round(metrics.get("site_average_views", {}).get("value") or 0)),
        "mail_total": int(round(metrics.get("mail_total", {}).get("value") or 0)),
        "mail_open_rate": round(open_rate, 2),
        "mail_interaction_rate": round(interaction_rate, 2),
        "mail_interaction_rate_over_opened": round(
            float(metrics.get("mail_interaction_rate_over_opened", {}).get("value") or 0.0),
            2,
        ),

        "strategic_axes": [],
        "internal_clients": [],
        "channel_mix": [],
        "format_mix": [],
        "top_push_by_interaction": raw_extracted.get("top_push_interaction", []),
        "top_push_by_open_rate": raw_extracted.get("top_push_open", []),
        "top_pull_notes": raw_extracted.get("top_pull_notes", []),
        "hitos": [],
        "events": [],

        "quality_flags": {
            "scope_country": "AR",
            "scope_mixed": False,
            "site_has_no_data_sections": False,
            "events_summary_available": False,
            "push_ranking_available": bool(raw_extracted.get("top_push_interaction") or raw_extracted.get("top_push_open")),
            "pull_ranking_available": bool(raw_extracted.get("top_pull_notes")),
            "historical_comparison_allowed": True,
        },

        "extraction_warnings": raw_extracted.get("warnings", []),
    }


# -------------------------
# Validación
# -------------------------

def validate_canonical_monthly(canonical: dict[str, Any]) -> dict[str, Any]:
    warnings: list[str] = []
    errors: list[str] = []

    extraction_warnings = list(canonical.get("extraction_warnings", []))

    if any(str(w).startswith("missing_anchor:") for w in extraction_warnings):
        errors.append("Faltan KPIs primarios por ancla exacta")

    for metric in ("mail_open_rate", "mail_interaction_rate", "mail_interaction_rate_over_opened"):
        value = float(canonical.get(metric, 0))

        if value < 0 or value > 100:
            errors.append(f"{metric} fuera de rango 0-100")
        elif value < 1:
            warnings.append(f"{metric} es menor a 1%; revisar escala")

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

    if mail_total >= 0 and plan_total > 0:
        min_mail = max(MIN_MAIL_ABSOLUTE, int(round(plan_total * MIN_MAIL_TO_PLAN_RELATION)))

        if mail_total < min_mail:
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