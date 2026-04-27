from __future__ import annotations

import json
import logging
import re
import unicodedata
from datetime import UTC, datetime
from itertools import islice
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

# Con 2 números toleramos casos "label + valor + referencia", y evitamos filas de tablas más densas.
MAX_NUMBERS_IN_LABEL_LINE = 2
LOOKAHEAD_LINES_AFTER_LABEL = 8
ANCHOR_PREFIX_LENGTH = 8
MAX_LOGGED_CANDIDATE_LINES = 20
CANON_TITLE_MAX_LENGTH = 180

KNOWN_KPI_LABELS = [
    "Media comunicaciones diarias",
    "Nº total de comunicaciones",
    "N° total de comunicaciones",
    "No total de comunicaciones",
    "N total de comunicaciones",
    "N total de comunicaciones por mes",
    "Total Páginas Vistas",
    "Noticias Publicadas",
    "Promedio Vistas",
    "Tasa de apertura promedio",
    "Tasa de interacción sobre mails enviados",
    "Tasa de interacción sobre mails abiertos",
    "Mails enviados",
]

PLAN_TOTAL_LABEL_VARIANTS = [
    "Nº total de comunicaciones",
    "N° total de comunicaciones",
    "No total de comunicaciones",
    "N total de comunicaciones",
]

REQUIRED_METRIC_KEYS = {
    "plan_total",
    "site_notes_total",
    "site_total_views",
    "mail_total",
    "mail_open_rate",
    "mail_interaction_rate",
}

OPTIONAL_METRIC_KEYS = {
    "plan_daily_average",
    "site_average_views",
    "mail_interaction_rate_over_opened",
}


# -------------------------
# Utils reemplazo metric_utils.py
# TODO: consolidar estos helpers en metric_utils.py y reutilizarlos desde analyzer.py.
# -------------------------

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


def _compact_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    simplified = without_accents.replace("º", "o").replace("°", "o")
    return re.sub(r"[^a-z0-9]", "", simplified.lower())


def _normalize_number_spacing(line: str) -> str:
    line = re.sub(r"(?<=\d)\s+(?=%)", "", line)
    line = re.sub(r"(?<=\d)\s+(?=[.,]\d)", "", line)
    return line


def _filtered_numbers(text: str, kind: str) -> list[str]:
    line_for_numbers = _normalize_number_spacing(text)
    nums = NUMBER_PATTERN.findall(line_for_numbers)

    if kind == "percent":
        return [n for n in nums if "%" in n]

    return [n for n in nums if "%" not in n]


def _label_pattern(label: str) -> str:
    tokens = _normalize_text(label).split()
    return r"\s*".join(re.escape(token) for token in tokens)


def _line_contains_label(line: str, label: str) -> bool:
    """Detecta si una línea contiene un label, sin descartar líneas ruidosas.

    A diferencia de _line_matches_label, esta función no rechaza la línea por tener
    muchos números. Esto es importante para casos reales de pypdf como:
    "ARGENTINA 30 40Total Páginas Vistas".

    La especificidad evita que "N total de comunicaciones" matchee contra
    "N total de comunicaciones por mes".
    """
    line_compact = _compact_text(line)
    label_compact = _compact_text(label)

    if label_compact not in line_compact:
        return False

    label_pos = line_compact.find(label_compact)
    same_start_matches = [
        known
        for known in KNOWN_KPI_LABELS
        if line_compact.find(_compact_text(known)) == label_pos
    ]

    if same_start_matches:
        longest = max(same_start_matches, key=lambda known: len(_compact_text(known)))
        return _compact_text(longest) == label_compact

    return True


def _line_matches_label(line: str, label: str) -> bool:
    """Match estricto de label para casos no ruidosos.

    Se conserva para usos auxiliares. La extracción principal usa _line_contains_label
    para poder detectar anchors pegadas al final de filas/tablas.
    """
    if not _line_contains_label(line, label):
        return False

    line_norm = _normalize_text(line)
    label_norm = _normalize_text(label)

    if line_norm == label_norm:
        return True

    if _compact_text(line) == _compact_text(label):
        return True

    numbers = NUMBER_PATTERN.findall(_normalize_number_spacing(line))
    return len(numbers) <= MAX_NUMBERS_IN_LABEL_LINE


def _line_contains_any_known_anchor(line: str, current_labels: list[str]) -> bool:
    line_compact = _compact_text(line)
    current_compacts = {_compact_text(label) for label in current_labels}

    for known_label in KNOWN_KPI_LABELS:
        known_compact = _compact_text(known_label)
        if known_compact in current_compacts:
            continue
        if known_compact in line_compact:
            return True

    return False


def _numbers_after_label_in_line(line: str, label: str, kind: str) -> list[str]:
    line_for_numbers = _normalize_number_spacing(line)
    line_norm = _normalize_text(line_for_numbers)
    pattern = _label_pattern(label)

    if not pattern:
        return []

    for match in re.finditer(pattern, line_norm):
        after_label = line_norm[match.end():]
        nums = _filtered_numbers(after_label, kind)
        if nums:
            return nums

    return []


def _numbers_before_label_in_line(line: str, label: str, kind: str) -> list[str]:
    line_for_numbers = _normalize_number_spacing(line)
    line_norm = _normalize_text(line_for_numbers)
    pattern = _label_pattern(label)

    if not pattern:
        return []

    for match in re.finditer(pattern, line_norm):
        before_label = line_norm[:match.start()]
        nums = _filtered_numbers(before_label, kind)
        if nums:
            return nums

    return []


def _prefix_before_label(line: str, label: str) -> str:
    line_norm = _normalize_text(_normalize_number_spacing(line))
    pattern = _label_pattern(label)

    for match in re.finditer(pattern, line_norm):
        return line_norm[:match.start()]

    return ""


def _numbers_before_expected_next_label(
    line: str,
    expected_next_labels: list[str],
    kind: str,
) -> list[str]:
    """Extrae el valor de una métrica cuando pypdf devuelve 'valor + próximo label'.

    Ejemplos reales:
    - "77.53%Tasa de interacción sobre mails enviados"
      => 77.53% corresponde a "Tasa de apertura promedio".
    - "8.86%Tasa de interacción sobre mails abiertos"
      => 8.86% corresponde a "Tasa de interacción sobre mails enviados".
    - "4,071Noticias Publicadas"
      => 4,071 corresponde a "Total Páginas Vistas".

    No se usa para tomar números antes del label actual, porque en estos PDFs ese
    número suele pertenecer a la métrica anterior.
    """
    if not expected_next_labels:
        return []

    line_for_numbers = _normalize_number_spacing(line)

    for next_label in expected_next_labels:
        if not _line_contains_label(line_for_numbers, next_label):
            continue

        nums = _numbers_before_label_in_line(line_for_numbers, next_label, kind)
        if not nums:
            continue

        # Seguridad: si antes del próximo label hay muchos números/texto tabular,
        # probablemente no es un KPI sino una fila de ranking.
        prefix = _prefix_before_label(line_for_numbers, next_label)

        if kind == "percent":
            # Para porcentajes toleramos poco ruido, pero exigimos que haya pocos números.
            all_nums = NUMBER_PATTERN.findall(prefix)
            if len(all_nums) <= 2:
                return nums
            continue

        # Para counts/floats exigimos que el valor antes del próximo label sea simple.
        # Acepta "4,071Noticias Publicadas"; rechaza líneas con fecha/texto/varios números.
        prefix_without_number = prefix
        for num in nums:
            prefix_without_number = prefix_without_number.replace(num, "", 1)
        if len(NUMBER_PATTERN.findall(prefix)) == 1 and not re.search(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", prefix_without_number):
            return nums

    return []


def _extract_pages_text(pdf_path: Path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    return [page.extract_text() or "" for page in reader.pages]


def _value_immediately_after_label(
    page_text: str,
    labels: str | list[str],
    kind: str,
    expected_next_labels: list[str] | None = None,
) -> str | None:
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]
    label_variants = labels if isinstance(labels, list) else [labels]
    expected_next_labels = expected_next_labels or []

    matched_indexes = [
        i
        for i, line in enumerate(lines)
        if any(_line_contains_label(line, label) for label in label_variants)
    ]

    for i in reversed(matched_indexes):
        line = lines[i]

        # Caso 1: valor en la misma línea, después del ancla actual.
        for label in label_variants:
            if _line_contains_label(line, label):
                nums = _numbers_after_label_in_line(line, label, kind)
                if nums:
                    return nums[0]

        # Caso 2: valor en líneas siguientes.
        for j in range(i + 1, min(i + 1 + LOOKAHEAD_LINES_AFTER_LABEL, len(lines))):
            next_line = lines[j]

            # Caso 2a: pypdf devuelve "valor + próximo label".
            nums_before_expected_next = _numbers_before_expected_next_label(
                next_line,
                expected_next_labels,
                kind,
            )
            if nums_before_expected_next:
                return nums_before_expected_next[-1]

            # Si aparece otra ancla conocida y no pudimos extraer un valor antes de ella,
            # cortamos para no cruzar al bloque de otra métrica.
            if _line_contains_any_known_anchor(next_line, label_variants):
                break

            nums = _filtered_numbers(next_line, kind)
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


def _extract_metric_with_page_fallback(
    pages: list[str],
    anchor: str,
    kind: str,
    expected_page: int,
    anchor_variants: list[str] | None = None,
    expected_next_labels: list[str] | None = None,
) -> tuple[dict[str, Any], str | None]:
    """Extrae una métrica desde su página esperada y, si falla, busca en el resto."""
    labels = anchor_variants or [anchor]
    expected_index = expected_page - 1

    if 0 <= expected_index < len(pages):
        expected_raw = _value_immediately_after_label(
            pages[expected_index],
            labels,
            kind,
            expected_next_labels=expected_next_labels,
        )
        if expected_raw is not None:
            return _metric(anchor, expected_raw, kind, expected_page), None

    for idx, page_text in enumerate(pages, start=1):
        if idx == expected_page:
            continue

        raw_value = _value_immediately_after_label(
            page_text,
            labels,
            kind,
            expected_next_labels=expected_next_labels,
        )
        if raw_value is not None:
            warning = f"anchor_out_of_expected_page:{anchor}:expected={expected_page}:found={idx}"
            return _metric(anchor, raw_value, kind, idx), warning

    return _metric(anchor, None, kind, expected_page), None


def _resolve_metric_page_index(metrics: dict[str, dict[str, Any]], metric_key: str, default_page: int, page_count: int) -> int:
    """Resolve a bounded 0-based page index from metric metadata.

    Uses the metric page when present; if metric_key is missing (or has no page),
    falls back to default_page. Then clamps the result to [0, page_count - 1].
    """
    page_number = int(metrics.get(metric_key, {}).get("page", default_page))
    return max(0, min(page_count - 1, page_number - 1))


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
            "title": title[:CANON_TITLE_MAX_LENGTH],
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
            "title": title[:CANON_TITLE_MAX_LENGTH],
            "users": users,
            "views": views,
            "raw": line,
        })

    return rows[:5]


def _extract_percent_values_from_items(items: list[str]) -> list[float]:
    values = []

    for item in items:
        cleaned = re.sub(r"(?<=\d)\s+(?=[.,]\d)", "", item)  # 1 .9 -> 1.9
        cleaned = re.sub(r"(?<=[.,]\d)\s+(?=%)", "", cleaned)  # 1.9 % -> 1.9%
        nums = re.findall(r"\d+(?:[.,]\d+)?%", cleaned)

        for n in nums:
            parsed = parse_percent_value(n)
            if parsed is not None:
                values.append(parsed)

    return values


def _extract_channel_mix(page_text: str) -> list[dict[str, Any]]:
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]

    for i, line in enumerate(lines):
        if "que canales y formatos se han utilizado" not in _normalize_text(line):
            continue

        window = lines[i + 1:i + 14]

        pct_values = []
        for item in window:
            nums = re.findall(r"\d+(?:[.,]\d+)?\s*%", item)
            for n in nums:
                parsed = parse_percent_value(n.replace(" ", ""))
                if parsed is not None:
                    pct_values.append(parsed)

        labels = [
            "Mail",
            "Intranet",
            "SITE",
            "Cartelería / Pantallas",
            "Widget #notelopierdas",
        ]

        return [
            {"channel": label, "pct": pct}
            for label, pct in zip(labels, pct_values[:len(labels)])
            if pct is not None
        ]

    return []


def _extract_format_mix(page_text: str) -> list[dict[str, Any]]:
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]

    for i, line in enumerate(lines):
        if "que canales y formatos se han utilizado" not in _normalize_text(line):
            continue

        window = lines[i + 1:i + 18]

        pct_values = _extract_percent_values_from_items(window)

        # Los primeros 5 porcentajes son canales.
        # Los siguientes corresponden a formatos.
        format_pcts = pct_values[5:]

        labels = [
            "Postal/Carta",
            "Noticia propia",
            "Noticia bbva.com",
            "Video",
        ]

        return [
            {"format": label, "pct": pct}
            for label, pct in zip(labels, format_pcts[:len(labels)])
            if pct is not None
        ]

    return []


def _extract_strategic_axes(page_text: str) -> list[dict[str, Any]]:
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]

    labels = [
        "RCP",
        "Sostenibilidad",
        "Empresas",
        "Creación de valor",
        "Innovación",
        "Equipo",
        "Otros",
    ]

    for i, line in enumerate(lines):
        if "distribucion por eje estrategico" not in _normalize_text(line):
            continue

        window = lines[i:i + 70]

        # Normalizar números partidos por extracción PDF: "2 0" -> "20", "1 2 0" -> "120".
        normalized_items = []
        for item in window:
            stripped_item = item.strip()
            if re.fullmatch(r"\d(?:\s+\d)+", stripped_item):
                item = re.sub(r"\s+", "", stripped_item)
            normalized_items.append(item)

        values = []
        started = False

        for item in normalized_items:
            norm = _normalize_text(item)

            if norm == "impactos":
                started = True
                continue

            if not started:
                continue

            if norm in {"rcp", "sostenibilidad", "empresas", "creacion de valor", "innovacion", "equipo", "otros"}:
                break

            if "%" in item:
                continue

            if re.fullmatch(r"\d+", item):
                value = int(item)

                values.append(value)

                if len(values) >= len(labels):
                    break

        # La secuencia se repite muchas veces. Tomamos la primera tanda completa.
        return [
            {"axis": label, "count": count}
            for label, count in zip(labels, values[:len(labels)])
        ]

    return []


# -------------------------
# Extracción principal
# -------------------------

def extract_raw_monthly_pdf(month_key: str, pdf_path: Path) -> dict[str, Any]:
    pages = _extract_pages_text(pdf_path)

    if len(pages) < 3:
        raise ValueError(f"El PDF debería tener al menos 3 páginas. Tiene {len(pages)}.")

    metric_specs = [
        (
            "plan_daily_average",
            "Media comunicaciones diarias",
            "float",
            1,
            None,
            PLAN_TOTAL_LABEL_VARIANTS,
        ),
        (
            "plan_total",
            "Nº total de comunicaciones",
            "count",
            1,
            PLAN_TOTAL_LABEL_VARIANTS,
            None,
        ),
        (
            "site_total_views",
            "Total Páginas Vistas",
            "count",
            2,
            None,
            ["Noticias Publicadas"],
        ),
        (
            "site_notes_total",
            "Noticias Publicadas",
            "count",
            2,
            None,
            ["Promedio Vistas"],
        ),
        (
            "site_average_views",
            "Promedio Vistas",
            "count",
            2,
            None,
            None,
        ),
        (
            "mail_open_rate",
            "Tasa de apertura promedio",
            "percent",
            3,
            None,
            ["Tasa de interacción sobre mails enviados"],
        ),
        (
            "mail_interaction_rate",
            "Tasa de interacción sobre mails enviados",
            "percent",
            3,
            None,
            ["Tasa de interacción sobre mails abiertos"],
        ),
        (
            "mail_interaction_rate_over_opened",
            "Tasa de interacción sobre mails abiertos",
            "percent",
            3,
            None,
            ["Mails enviados"],
        ),
        (
            "mail_total",
            "Mails enviados",
            "count",
            3,
            None,
            None,
        ),
    ]

    metrics: dict[str, dict[str, Any]] = {}
    fallback_warnings: list[str] = []

    for key, anchor, kind, expected_page, anchor_variants, expected_next_labels in metric_specs:
        metric, warning = _extract_metric_with_page_fallback(
            pages,
            anchor,
            kind,
            expected_page,
            anchor_variants,
            expected_next_labels,
        )
        metrics[key] = metric
        if warning:
            fallback_warnings.append(warning)

    for key, anchor, _, expected_page, _, _ in metric_specs:
        if metrics.get(key, {}).get("missing"):
            logger.error(
                "event=metric_missing anchor=%s page=%s candidate_lines=%s",
                anchor,
                expected_page,
                list(islice((
                    line for page in pages for line in page.splitlines()
                    if _compact_text(anchor)[:ANCHOR_PREFIX_LENGTH] in _compact_text(line)
                ), MAX_LOGGED_CANDIDATE_LINES)),
            )

    page_for_mail_idx = _resolve_metric_page_index(metrics, "mail_total", 3, len(pages))
    page_for_site_idx = _resolve_metric_page_index(metrics, "site_total_views", 2, len(pages))
    page_for_plan_idx = _resolve_metric_page_index(metrics, "plan_total", 1, len(pages))

    page_for_mail = pages[page_for_mail_idx]
    page_for_site = pages[page_for_site_idx]
    page_for_plan = pages[page_for_plan_idx]

    mail_rows = _extract_mail_table(page_for_mail)
    top_push_open, top_push_interaction = _build_push_rankings(mail_rows)
    top_pull_notes = _extract_top_pull_notes(page_for_site)
    channel_mix = _extract_channel_mix(page_for_plan)
    format_mix = _extract_format_mix(page_for_plan)
    strategic_axes = _extract_strategic_axes(page_for_plan)

    warnings = fallback_warnings + [
        f"missing_anchor:{k}:{v.get('anchor')}"
        for k, v in metrics.items()
        if v.get("missing")
    ]

    return {
        "month": month_key,
        "source_pdf": str(pdf_path),
        "extracted_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "parser": "deterministic_pdf_v7_kpi_sequence_safe",
        "page_count": len(pages),
        "metrics": metrics,
        "mail_table": mail_rows,
        "top_push_open": top_push_open,
        "top_push_interaction": top_push_interaction,
        "top_pull_notes": top_pull_notes,
        "channel_mix": channel_mix,
        "format_mix": format_mix,
        "strategic_axes": strategic_axes,
        "warnings": warnings,
    }


# -------------------------
# Canonical
# -------------------------

def _first_present(item: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        value = item.get(key)
        if value not in (None, ""):
            return value
    return None


def _canon_distribution(
    items: list[dict[str, Any]],
    label_keys: tuple[str, ...],
    value_keys: tuple[str, ...],
) -> list[dict[str, Any]]:
    rows = []

    for item in items or []:
        if not isinstance(item, dict):
            continue

        label = _first_present(item, *label_keys)
        value = _first_present(item, *value_keys)

        if label in (None, "") or value in (None, ""):
            continue

        rows.append({
            "label": str(label).strip(),
            "value": round(to_float_locale(str(value), 0.0), 2),
        })

    return rows


def _canon_push_rows(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []

    for item in items or []:
        if not isinstance(item, dict):
            continue

        name = _first_present(item, "name", "title")
        if not name:
            continue

        clicks_raw = _first_present(item, "clicks")
        open_rate_raw = _first_present(item, "open_rate")
        interaction_raw = _first_present(item, "interaction", "interaction_rate", "ctr")

        rows.append({
            "name": str(name).strip()[:CANON_TITLE_MAX_LENGTH],
            "clicks": parse_integer_value(str(clicks_raw or 0)) or 0,
            "open_rate": parse_percent_value(str(open_rate_raw or 0)) or 0.0,
            "interaction": parse_percent_value(str(interaction_raw or 0)) or 0.0,
            "date": item.get("date"),
        })

    return rows


def _canon_pull_rows(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []

    for item in items or []:
        if not isinstance(item, dict):
            continue

        title = _first_present(item, "title", "name")
        if not title:
            continue

        unique_reads_raw = _first_present(item, "unique_reads", "users", "reads")
        total_reads_raw = _first_present(item, "total_reads", "views")

        rows.append({
            "title": str(title).strip()[:CANON_TITLE_MAX_LENGTH],
            "unique_reads": parse_integer_value(str(unique_reads_raw or 0)) or 0,
            "total_reads": parse_integer_value(str(total_reads_raw or 0)) or 0,
            "date": item.get("date"),
        })

    return rows


def canonicalize_monthly(raw_extracted: dict[str, Any]) -> dict[str, Any]:
    metrics = raw_extracted.get("metrics", {})

    plan_daily_average = float(metrics.get("plan_daily_average", {}).get("value") or 0.0)
    plan_total = int(round(metrics.get("plan_total", {}).get("value") or 0))
    site_notes_total = int(round(metrics.get("site_notes_total", {}).get("value") or 0))
    site_total_views = int(round(metrics.get("site_total_views", {}).get("value") or 0))
    site_average_views = int(round(metrics.get("site_average_views", {}).get("value") or 0))
    mail_total = int(round(metrics.get("mail_total", {}).get("value") or 0))
    open_rate = float(metrics.get("mail_open_rate", {}).get("value") or 0.0)
    interaction_rate = float(metrics.get("mail_interaction_rate", {}).get("value") or 0.0)
    interaction_rate_over_opened = float(metrics.get("mail_interaction_rate_over_opened", {}).get("value") or 0.0)

    derived_warnings: list[str] = []

    if site_total_views <= 0 and site_notes_total > 0 and site_average_views > 0:
        site_total_views = site_notes_total * site_average_views
        derived_warnings.append("derived_metric:site_total_views:site_notes_total*site_average_views")

    return {
        "month": raw_extracted.get("month"),
        "generation_mode": "deterministic_pdf",
        "extraction_method": raw_extracted.get("parser"),

        "plan_daily_average": plan_daily_average,
        "plan_total": plan_total,
        "site_notes_total": site_notes_total,
        "site_total_views": site_total_views,
        "site_average_views": site_average_views,
        "mail_total": mail_total,
        "mail_open_rate": round(open_rate, 2),
        "mail_interaction_rate": round(interaction_rate, 2),
        "mail_interaction_rate_over_opened": round(interaction_rate_over_opened, 2),

        "strategic_axes": _canon_distribution(
            raw_extracted.get("strategic_axes", []),
            label_keys=("label", "theme", "axis", "name"),
            value_keys=("value", "weight", "count"),
        ),
        "internal_clients": [],
        "channel_mix": _canon_distribution(
            raw_extracted.get("channel_mix", []),
            label_keys=("label", "channel", "name"),
            value_keys=("value", "pct", "weight"),
        ),
        "format_mix": _canon_distribution(
            raw_extracted.get("format_mix", []),
            label_keys=("label", "format", "name"),
            value_keys=("value", "pct", "weight"),
        ),
        "top_push_by_interaction": _canon_push_rows(raw_extracted.get("top_push_interaction", [])),
        "top_push_by_open_rate": _canon_push_rows(raw_extracted.get("top_push_open", [])),
        "top_pull_notes": _canon_pull_rows(raw_extracted.get("top_pull_notes", [])),
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

        "extraction_warnings": raw_extracted.get("warnings", []) + derived_warnings,
    }


# -------------------------
# Validación
# -------------------------

def _missing_anchor_key(warning: str) -> str | None:
    if not warning.startswith("missing_anchor:"):
        return None

    parts = warning.split(":")
    if len(parts) < 2:
        return None

    return parts[1]


def _derived_metric_key(warning: str) -> str | None:
    if not warning.startswith("derived_metric:"):
        return None

    parts = warning.split(":")
    if len(parts) < 2:
        return None

    return parts[1]


def validate_canonical_monthly(canonical: dict[str, Any]) -> dict[str, Any]:
    warnings: list[str] = []
    errors: list[str] = []

    extraction_warnings = list(canonical.get("extraction_warnings", []))

    derived_metrics = {
        key
        for warning in extraction_warnings
        if (key := _derived_metric_key(str(warning))) is not None
    }

    missing_required: list[str] = []
    missing_optional: list[str] = []

    for warning in extraction_warnings:
        key = _missing_anchor_key(str(warning))

        if key in REQUIRED_METRIC_KEYS and key not in derived_metrics:
            missing_required.append(key)
        elif key in OPTIONAL_METRIC_KEYS:
            missing_optional.append(key)

    if missing_required:
        errors.append(
            "Faltan KPIs primarios por ancla exacta: "
            + ", ".join(sorted(set(missing_required)))
        )

    if missing_optional:
        warnings.append(
            "Faltan KPIs secundarios por ancla exacta: "
            + ", ".join(sorted(set(missing_optional)))
        )

    missing_optional_set = set(missing_optional)

    for metric in ("mail_open_rate", "mail_interaction_rate", "mail_interaction_rate_over_opened"):
        value = float(canonical.get(metric, 0))

        if value < 0 or value > 100:
            errors.append(f"{metric} fuera de rango 0-100")
        elif value < 1 and metric not in missing_optional_set:
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

        if mail_total > plan_total * MAX_MAIL_TO_PLAN_RATIO:
            warnings.append("mail_total muy alto respecto a plan_total; revisar escala o extracción")

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


def persist_monthly_artifacts(
    month_key: str,
    raw_extracted: dict[str, Any],
    canonical: dict[str, Any],
    validation: dict[str, Any],
) -> None:
    from config import CANONICAL_MONTHLY_DIR, RAW_EXTRACTED_DIR, VALIDATION_DIR, ensure_dir

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

    raise ValueError(
        f"No pude inferir month_key desde el nombre del PDF: {pdf_path.name}. "
        "Pasa month_key explícitamente."
    )


def extract_single_pdf_to_raw(
    input_pdf: Path,
    output_json: Path,
    month_key: str | None = None,
) -> dict[str, Any]:
    resolved_month = month_key or infer_month_key_from_pdf_path(input_pdf)
    raw_extracted = extract_raw_monthly_pdf(resolved_month, input_pdf)

    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_json.write_text(
        json.dumps(raw_extracted, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return raw_extracted
