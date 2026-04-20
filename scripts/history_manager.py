from __future__ import annotations

import json
import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from config import DATA_DIR, ensure_dir

HISTORY_PATH = DATA_DIR / "historico_kpis.json"


def _parse_month_slug(month_slug: str) -> tuple[int, int] | None:
    try:
        year_str, month_str = str(month_slug).split("-")
        year = int(year_str)
        month = int(month_str)
    except Exception:
        return None
    if month < 1 or month > 12:
        return None
    return year, month


def _previous_month_slug(month_slug: str) -> str | None:
    parsed = _parse_month_slug(month_slug)
    if not parsed:
        return None
    year, month = parsed
    if month == 1:
        return f"{year - 1}-12"
    return f"{year:04d}-{month - 1:02d}"


def _quarter_key(year: int, quarter: int) -> str:
    return f"{year:04d}-Q{quarter}"


def _infer_period_identity(period: dict[str, Any]) -> tuple[str, str] | None:
    kind = str(period.get("kind") or "")
    slug = str(period.get("slug") or "")
    months = period.get("months") or []

    if kind == "month" and isinstance(months, list) and months:
        month_ref = str(months[-1])
        if _parse_month_slug(month_ref):
            return "month", month_ref

    if kind == "quarter":
        year = period.get("year")
        quarter = period.get("quarter")
        if isinstance(year, int) and isinstance(quarter, int) and quarter in {1, 2, 3, 4}:
            return "quarter", _quarter_key(year, quarter)

    if kind == "year":
        year = period.get("year")
        if isinstance(year, int):
            return "year", str(year)

    if slug.startswith("month_") and isinstance(months, list) and months:
        month_ref = str(months[-1])
        if _parse_month_slug(month_ref):
            return "month", month_ref
    if slug.startswith("quarter_"):
        return "quarter", slug
    if slug.startswith("year_"):
        return "year", slug.replace("year_", "")
    return None


def _previous_period_key(period_kind: str, period_ref: str) -> str | None:
    if period_kind == "month":
        return _previous_month_slug(period_ref)

    if period_kind == "quarter":
        if period_ref.startswith("quarter_"):
            return None
        try:
            year_str, quarter_str = period_ref.split("-Q")
            year = int(year_str)
            quarter = int(quarter_str)
        except Exception:
            return None
        if quarter == 1:
            return _quarter_key(year - 1, 4)
        return _quarter_key(year, quarter - 1)

    if period_kind == "year":
        try:
            return str(int(period_ref) - 1)
        except Exception:
            return None

    return None


def _to_float(value: Any) -> float | None:
    if value in (None, "", "-", "Sin datos previos"):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("%", "").strip()
    if "." in text and "," in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    normalized = "".join(ch for ch in text if ch.isdigit() or ch in ".-")
    if not normalized:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", normalized)
    if not match:
        return None
    try:
        return float(match.group(0))
    except Exception:
        return None


def _safe_pct_change(current: Any, previous: Any) -> str:
    current_num = _to_float(current)
    previous_num = _to_float(previous)
    if current_num is None or previous_num is None:
        return "Sin datos previos"
    if previous_num == 0:
        return "No comparable (previo=0)"
    return f"{round(((current_num - previous_num) / previous_num) * 100, 1)}%"


def load_history(history_path: Path = HISTORY_PATH) -> dict[str, Any]:
    if not history_path.exists():
        return {"records": {}}
    try:
        payload = json.loads(history_path.read_text(encoding="utf-8"))
    except Exception:
        return {"records": {}}
    records = payload.get("records")
    if not isinstance(records, dict):
        return {"records": {}}
    return {"records": records}


def persist_calculated_totals(
    period: dict[str, Any],
    kpis_calculados: dict[str, Any],
    history_path: Path = HISTORY_PATH,
) -> None:
    identity = _infer_period_identity(period)
    if not identity:
        return
    period_kind, period_ref = identity
    payload = load_history(history_path)
    records = payload["records"]
    records[f"{period_kind}:{period_ref}"] = {
        "kind": period_kind,
        "period_ref": period_ref,
        "period_slug": period.get("slug"),
        "calculated_totals": kpis_calculados.get("calculated_totals", {}),
        "updated_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
    }
    ensure_dir(history_path.parent)
    history_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def apply_historical_comparison(
    period: dict[str, Any],
    kpis_calculados: dict[str, Any],
    history_path: Path = HISTORY_PATH,
) -> dict[str, Any]:
    totals = kpis_calculados.setdefault("calculated_totals", {})
    identity = _infer_period_identity(period)
    if not identity:
        totals["volume_previous"] = "Sin datos previos"
        totals["volume_change"] = "Sin datos previos"
        totals["previous_push_volume"] = "Sin datos previos"
        totals["latest_push_variation"] = "Sin datos previos"
        return kpis_calculados

    period_kind, period_ref = identity
    previous_key = _previous_period_key(period_kind, period_ref)
    history = load_history(history_path)
    previous_record = history.get("records", {}).get(f"{period_kind}:{previous_key}") if previous_key else None

    if previous_record:
        previous_volume = previous_record.get("calculated_totals", {}).get("push_volume_period", "Sin datos previos")
    else:
        previous_volume = "Sin datos previos"

    current_volume = totals.get("push_volume_period", "Sin datos previos")
    volume_change = _safe_pct_change(current_volume, previous_volume)

    totals["volume_previous"] = previous_volume
    totals["volume_change"] = volume_change
    totals["previous_push_volume"] = previous_volume
    totals["latest_push_variation"] = volume_change
    return kpis_calculados
