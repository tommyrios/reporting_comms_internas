from __future__ import annotations

import argparse
import json
import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

REQUIRED_CANONICAL_FIELDS = {
    "month",
    "plan_total",
    "site_notes_total",
    "site_total_views",
    "mail_total",
    "mail_open_rate",
    "mail_interaction_rate",
    "strategic_axes",
    "internal_clients",
    "channel_mix",
    "format_mix",
    "top_push_by_interaction",
    "top_push_by_open_rate",
    "top_pull_notes",
    "quality_flags",
}

REQUIRED_REPORT_FIELDS = {"period", "kpis", "narrative", "quality_flags", "render_plan"}

MIN_SITE_VIEWS_PER_NOTE = 10
MIN_MAIL_ABSOLUTE = 10
MIN_MAIL_TO_PLAN_RELATION = 0.2
MAX_MAIL_TO_PLAN_RATIO = 10


def to_float(value: Any, default: float = 0.0) -> float:
    if value in (None, "", "-"):
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("%", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        return float(text)
    except Exception:
        return default


def to_int(value: Any, default: int = 0) -> int:
    return int(round(to_float(value, float(default))))


def normalize_percentage(value: Any) -> float:
    number = to_float(value, 0.0)
    if 0 < number <= 1:
        return round(number * 100, 2)
    return round(number, 2)


def clean_text(value: Any, max_len: int = 180) -> str:
    text = str(value or "").replace("_", " ").replace("|", " ").strip()
    text = re.sub(r"\s+", " ", text).replace("...", "…")
    text = text.rstrip("…").strip(" -")
    if len(text) <= max_len:
        return text
    cut = text[:max_len].rsplit(" ", 1)[0].strip()
    return cut or text[:max_len].strip()


def normalize_push_row(row: dict[str, Any]) -> dict[str, Any]:
    interaction = normalize_percentage(row.get("interaction", row.get("interaction_rate", row.get("ctr", 0))))
    open_rate = normalize_percentage(row.get("open_rate", row.get("opens", 0)))
    clicks = to_int(row.get("clicks", 0))
    normalized = {
        "name": clean_text(row.get("name") or row.get("title") or "Sin título", 120),
        "clicks": clicks,
        "open_rate": open_rate,
        "interaction": interaction,
    }
    if row.get("month"):
        normalized["month"] = row.get("month")
    if row.get("date"):
        normalized["date"] = row.get("date")

    complete = True
    issue = None
    if interaction <= 0:
        complete = False
        issue = "sin_interaccion"
    elif clicks <= 0 and interaction > 20:
        complete = False
        issue = "interaccion_alta_sin_clicks"
    elif open_rate > 0 and interaction > open_rate + 10:
        complete = False
        issue = "interaccion_mayor_a_apertura"

    normalized["data_complete"] = complete
    if issue:
        normalized["data_quality_issue"] = issue
    return normalized


def sanitize_push_ranking(rows: Any, *, metric_key: str = "interaction", limit: int | None = 5) -> list[dict[str, Any]]:
    if not isinstance(rows, list):
        return []
    normalized = [normalize_push_row(row) for row in rows if isinstance(row, dict)]
    normalized.sort(key=lambda item: (to_float(item.get(metric_key), 0.0), to_int(item.get("clicks"), 0)), reverse=True)
    return normalized[:limit] if limit else normalized


def _validate_push_rows(rows: Any, label: str, *, metric_key: str = "interaction") -> list[str]:
    warnings: list[str] = []
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        item = normalize_push_row(row)
        name = item.get("name") or "sin título"
        interaction = normalize_percentage(item.get("interaction"))
        open_rate = normalize_percentage(item.get("open_rate"))
        clicks = to_int(item.get("clicks"))

        if metric_key == "open_rate":
            if open_rate <= 0:
                warnings.append(f"{label}: fila sin apertura medible en '{name}'")
            if open_rate > 0 and interaction > open_rate + 10:
                warnings.append(f"{label}: interacción mayor a apertura en '{name}'")
            continue

        if interaction <= 0:
            warnings.append(f"{label}: fila sin interacción medible en '{name}'")
        elif clicks <= 0 and interaction > 20:
            warnings.append(f"{label}: interacción alta y 0 clics en '{name}'")
        elif open_rate > 0 and interaction > open_rate + 10:
            warnings.append(f"{label}: interacción mayor a apertura en '{name}'")
    return warnings


def validate_canonical_quality(canonical: dict[str, Any]) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []

    if not isinstance(canonical, dict):
        return _result(False, ["canonical debe ser objeto JSON"], [])

    missing = sorted(REQUIRED_CANONICAL_FIELDS - set(canonical.keys()))
    if missing:
        errors.append("Faltan campos del contrato mensual: " + ", ".join(missing))

    month = clean_text(canonical.get("month"), 32)
    if not month or month == "-":
        errors.append("month faltante")

    plan_total = to_int(canonical.get("plan_total"))
    site_notes_total = to_int(canonical.get("site_notes_total"))
    site_total_views = to_int(canonical.get("site_total_views"))
    mail_total = to_int(canonical.get("mail_total"))
    open_rate = normalize_percentage(canonical.get("mail_open_rate"))
    interaction_rate = normalize_percentage(canonical.get("mail_interaction_rate"))

    if plan_total <= 0:
        errors.append("plan_total debe ser mayor a 0")
    if site_notes_total < 0:
        errors.append("site_notes_total no puede ser negativo")
    if site_total_views < 0:
        errors.append("site_total_views no puede ser negativo")
    if mail_total < 0:
        errors.append("mail_total no puede ser negativo")
    if not 0 <= open_rate <= 100:
        errors.append("mail_open_rate fuera de rango 0-100")
    if not 0 <= interaction_rate <= 100:
        errors.append("mail_interaction_rate fuera de rango 0-100")

    if plan_total > 0:
        min_mail = max(MIN_MAIL_ABSOLUTE, int(round(plan_total * MIN_MAIL_TO_PLAN_RELATION)))
        if mail_total < min_mail:
            errors.append("mail_total sospechosamente bajo respecto a plan_total")
        if mail_total > plan_total * MAX_MAIL_TO_PLAN_RATIO:
            warnings.append("mail_total muy alto respecto a plan_total; revisar escala o extracción")

    if site_notes_total > 0 and site_total_views < site_notes_total * MIN_SITE_VIEWS_PER_NOTE:
        errors.append("site_total_views sospechosamente bajo respecto a site_notes_total")

    if open_rate > 0 and interaction_rate > open_rate + 10:
        warnings.append("mail_interaction_rate supera mail_open_rate por más de 10 pp; revisar definición")

    warnings.extend(_validate_push_rows(canonical.get("top_push_by_interaction"), "top_push_by_interaction", metric_key="interaction"))
    warnings.extend(_validate_push_rows(canonical.get("top_push_by_open_rate"), "top_push_by_open_rate", metric_key="open_rate"))

    for row in canonical.get("top_pull_notes", []) or []:
        if not isinstance(row, dict):
            continue
        title = clean_text(row.get("title") or row.get("name") or "sin título", 96)
        unique = to_int(row.get("unique_reads", row.get("users", 0)))
        total = to_int(row.get("total_reads", row.get("views", 0)))
        if unique < 0 or total < 0:
            errors.append(f"top_pull_notes: lecturas negativas en '{title}'")
        if unique > 0 and total > 0 and unique > total:
            warnings.append(f"top_pull_notes: usuarios únicos mayores que vistas totales en '{title}'")

    return _result(not errors, errors, warnings)


def validate_report_quality(report: dict[str, Any]) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []
    if not isinstance(report, dict):
        return _result(False, ["report debe ser objeto JSON"], [])
    missing = sorted(REQUIRED_REPORT_FIELDS - set(report.keys()))
    if missing:
        errors.append("Faltan campos del reporte: " + ", ".join(missing))
    render_plan = report.get("render_plan") if isinstance(report.get("render_plan"), dict) else {}
    modules = render_plan.get("modules") if isinstance(render_plan.get("modules"), list) else []
    if not modules:
        errors.append("render_plan.modules vacío")
    seen = set()
    for module in modules:
        if not isinstance(module, dict):
            errors.append("render_plan.modules contiene un elemento no objeto")
            continue
        key = module.get("key")
        if not key:
            errors.append("módulo sin key")
        elif key in seen:
            warnings.append(f"módulo duplicado: {key}")
        seen.add(key)
        if module.get("payload") is None:
            errors.append(f"módulo sin payload: {key or 'sin key'}")
    return _result(not errors, errors, warnings)


def _result(is_valid: bool, errors: list[str], warnings: list[str]) -> dict[str, Any]:
    return {
        "validated_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "is_valid": is_valid,
        "errors": errors,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Valida artefactos JSON del pipeline de Comunicaciones Internas.")
    parser.add_argument("input", type=Path, help="Ruta al JSON canonical_monthly o report.json")
    parser.add_argument("--kind", choices=("canonical", "report"), default="canonical")
    parser.add_argument("--warn-only", action="store_true", help="Devuelve exit code 0 aunque haya errores")
    args = parser.parse_args()

    payload = json.loads(args.input.read_text(encoding="utf-8"))
    result = validate_report_quality(payload) if args.kind == "report" else validate_canonical_quality(payload)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if (result["is_valid"] or args.warn_only) else 1


if __name__ == "__main__":
    raise SystemExit(main())
