from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from typing import Any

from metric_utils import normalize_percentage, to_float_locale


REQUIRED_MONTHLY_FIELDS = {
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
    "hitos",
    "events",
    "quality_flags",
}

REQUIRED_QUALITY_FLAGS = {
    "scope_country",
    "scope_mixed",
    "site_has_no_data_sections",
    "events_summary_available",
    "push_ranking_available",
    "pull_ranking_available",
    "historical_comparison_allowed",
}

BASE_STRUCTURE: dict[str, Any] = {
    "period": {"slug": "-", "label": "-"},
    "kpis": {},
    "narrative": {},
    "quality_flags": {},
    "render_plan": {"modules": []},
}

PLAN_MAIL_ABS_DELTA_THRESHOLD = 5
PLAN_MAIL_REL_DELTA_THRESHOLD = 0.5


def _to_int(value: Any, default: int = 0) -> int:
    if value in (None, "", "-"):
        return default
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(round(value))
    text = str(value).strip()
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return int(float(text))
    except Exception:
        digits = "".join(ch for ch in str(value) if ch.isdigit())
        return int(digits) if digits else default


def _to_float(value: Any, default: float = 0.0) -> float:
    if value in (None, "", "-"):
        return default
    return to_float_locale(value, default)


def _clean_title(value: Any, max_len: int = 90) -> str:
    text = str(value or "").replace("_", " ").replace("|", " ").strip()
    text = " ".join(text.split())
    if not text:
        return "Sin título"
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def _normalize_pct(value: Any) -> float:
    return normalize_percentage(value)


def _normalize_weighted_list(items: Any, label_keys: list[str], value_keys: list[str]) -> list[dict[str, Any]]:
    if not isinstance(items, list):
        return []
    rows: list[dict[str, Any]] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        label = ""
        for key in label_keys:
            label = str(item.get(key) or "").strip()
            if label:
                break
        if not label:
            continue
        value = 0.0
        for key in value_keys:
            if key in item:
                value = _to_float(item.get(key), 0.0)
                break
        rows.append({"label": _clean_title(label, 48), "value": round(value, 2)})
    return rows


def _infer_quality_flags(summary: dict[str, Any]) -> dict[str, Any]:
    source = summary.get("quality_flags") if isinstance(summary.get("quality_flags"), dict) else {}
    events = summary.get("events") if isinstance(summary.get("events"), list) else []
    top_push_i = summary.get("top_push_by_interaction") if isinstance(summary.get("top_push_by_interaction"), list) else []
    top_push_o = summary.get("top_push_by_open_rate") if isinstance(summary.get("top_push_by_open_rate"), list) else []
    top_pull = summary.get("top_pull_notes") if isinstance(summary.get("top_pull_notes"), list) else []

    flags = {
        "scope_country": str(source.get("scope_country") or "AR").strip() or "AR",
        "scope_mixed": bool(source.get("scope_mixed", False)),
        "site_has_no_data_sections": bool(source.get("site_has_no_data_sections", False)),
        "events_summary_available": bool(source.get("events_summary_available", bool(events))),
        "push_ranking_available": bool(source.get("push_ranking_available", bool(top_push_i or top_push_o))),
        "pull_ranking_available": bool(source.get("pull_ranking_available", bool(top_pull))),
        "historical_comparison_allowed": bool(source.get("historical_comparison_allowed", True)),
    }
    return flags


def normalize_monthly_summary(summary: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(summary, dict):
        raise ValueError("monthly summary debe ser objeto JSON")

    data = summary.get("data") if isinstance(summary.get("data"), dict) else {}
    insights = summary.get("insights") if isinstance(summary.get("insights"), dict) else {}

    top_push_legacy = insights.get("top_push_comm") if isinstance(insights.get("top_push_comm"), dict) else {}
    top_pull_legacy = insights.get("top_pull_note") if isinstance(insights.get("top_pull_note"), dict) else {}

    normalized = {
        "month": str(summary.get("month") or "-").strip(),
        "plan_total": _to_int(summary.get("plan_total", data.get("push_volume", summary.get("mail_total", 0)))),
        "site_notes_total": _to_int(summary.get("site_notes_total", data.get("pull_notes", 0))),
        "site_total_views": _to_int(summary.get("site_total_views", data.get("pull_reads", 0))),
        "mail_total": _to_int(summary.get("mail_total", data.get("push_volume", summary.get("plan_total", 0)))),
        "mail_open_rate": _normalize_pct(summary.get("mail_open_rate", data.get("push_opens_pct", 0))),
        "mail_interaction_rate": _normalize_pct(summary.get("mail_interaction_rate", data.get("push_interaction_pct", 0))),
        "strategic_axes": summary.get("strategic_axes") if isinstance(summary.get("strategic_axes"), list) else insights.get("strategic_axes", []),
        "internal_clients": summary.get("internal_clients") if isinstance(summary.get("internal_clients"), list) else insights.get("internal_clients", []),
        "channel_mix": summary.get("channel_mix") if isinstance(summary.get("channel_mix"), list) else [],
        "format_mix": summary.get("format_mix") if isinstance(summary.get("format_mix"), list) else [],
        "top_push_by_interaction": summary.get("top_push_by_interaction") if isinstance(summary.get("top_push_by_interaction"), list) else ([] if not top_push_legacy else [top_push_legacy]),
        "top_push_by_open_rate": summary.get("top_push_by_open_rate") if isinstance(summary.get("top_push_by_open_rate"), list) else ([] if not top_push_legacy else [top_push_legacy]),
        "top_pull_notes": summary.get("top_pull_notes") if isinstance(summary.get("top_pull_notes"), list) else ([] if not top_pull_legacy else [top_pull_legacy]),
        "hitos": summary.get("hitos") if isinstance(summary.get("hitos"), list) else ([{"title": _clean_title(insights.get("hitos_mes"), 72)}] if insights.get("hitos_mes") else []),
        "events": summary.get("events") if isinstance(summary.get("events"), list) else [],
        "quality_flags": {},
    }
    normalized["quality_flags"] = _infer_quality_flags({**summary, **normalized})
    return normalized


def validate_monthly_summary_contract(summary: dict[str, Any]) -> dict[str, Any]:
    source_keys = set(summary.keys()) if isinstance(summary, dict) else set()
    has_full_contract = REQUIRED_MONTHLY_FIELDS.issubset(source_keys)
    has_legacy_contract = isinstance(summary.get("data"), dict) and isinstance(summary.get("insights"), dict)
    if not has_full_contract and not has_legacy_contract:
        missing = sorted(REQUIRED_MONTHLY_FIELDS - source_keys)
        raise ValueError(f"Contrato mensual incompleto, faltan campos: {', '.join(missing)}")

    normalized = normalize_monthly_summary(summary)

    if not normalized.get("month") or normalized["month"] == "-":
        raise ValueError("Contrato mensual inválido: falta campo month")

    quality_flags = normalized.get("quality_flags", {})
    missing_flags = [key for key in REQUIRED_QUALITY_FLAGS if key not in quality_flags]
    if missing_flags:
        raise ValueError(f"quality_flags incompleto, faltan: {', '.join(sorted(missing_flags))}")

    return normalized


def _aggregate_weighted(items_by_month: list[list[dict[str, Any]]], label_key: str = "label") -> list[dict[str, Any]]:
    totals: dict[str, float] = defaultdict(float)
    for items in items_by_month:
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            label = str(item.get(label_key) or item.get("theme") or item.get("name") or "").strip()
            if not label:
                continue
            value = _to_float(item.get("value", item.get("weight", item.get("participants", 0))), 0.0)
            totals[_clean_title(label, 52)] += value
    rows = [{"label": key, "value": round(val, 2)} for key, val in totals.items()]
    rows.sort(key=lambda row: row["value"], reverse=True)
    return rows


def _looks_like_distribution(items: Any) -> bool:
    if not isinstance(items, list) or not items:
        return False
    values = [_to_float(item.get("value", item.get("weight", item.get("participants", 0))), 0.0) for item in items if isinstance(item, dict)]
    if not values:
        return False
    if all(0 <= value <= 1 for value in values):
        total = sum(values)
        return 0.95 <= total <= 1.05
    if any(value < 0 or value > 100 for value in values):
        return False
    total = sum(values)
    return 95 <= total <= 105


def _aggregate_distribution(items_by_month: list[list[dict[str, Any]]], label_key: str = "label") -> tuple[list[dict[str, Any]], bool]:
    monthly_distributions = [items for items in items_by_month if _looks_like_distribution(items)]
    if not monthly_distributions:
        return _aggregate_weighted(items_by_month, label_key=label_key), False

    values_by_label: dict[str, list[float]] = defaultdict(list)
    for month_items in monthly_distributions:
        total = sum(_to_float(item.get("value", item.get("weight", 0)), 0.0) for item in month_items if isinstance(item, dict)) or 1.0
        for item in month_items:
            if not isinstance(item, dict):
                continue
            label = str(item.get(label_key) or item.get("theme") or item.get("name") or "").strip()
            if not label:
                continue
            value = _to_float(item.get("value", item.get("weight", 0)), 0.0)
            values_by_label[_clean_title(label, 52)].append((value / total) * 100)

    rows = [{"label": label, "value": round(sum(values) / len(values), 2)} for label, values in values_by_label.items() if values]
    rows.sort(key=lambda row: row["value"], reverse=True)
    return rows, True


def _exceeds_plan_mail_threshold(plan_total: int, mail_total: int) -> bool:
    return abs(plan_total - mail_total) > max(
        PLAN_MAIL_ABS_DELTA_THRESHOLD,
        int(mail_total * PLAN_MAIL_REL_DELTA_THRESHOLD),
    )


def _top_push(summary_rows: list[dict[str, Any]], source_key: str, value_key: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in summary_rows:
        for item in row.get(source_key, []) or []:
            if not isinstance(item, dict):
                continue
            name = _clean_title(item.get("name") or item.get("title"), 90)
            if not name:
                continue
            rows.append({
                "name": name,
                "clicks": _to_int(item.get("clicks", 0)),
                "open_rate": _normalize_pct(item.get("open_rate", item.get("opens", 0))),
                "interaction": _normalize_pct(item.get("interaction", item.get("interaction_rate", 0))),
                "month": row.get("month"),
                "_sort": _to_float(item.get(value_key, item.get("interaction", item.get("open_rate", 0))), 0.0),
            })
    rows.sort(key=lambda item: (item.get("_sort", 0), item.get("clicks", 0)), reverse=True)
    for item in rows:
        item.pop("_sort", None)
    return rows[:5]


def _top_pull(summary_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in summary_rows:
        for item in row.get("top_pull_notes", []) or []:
            if not isinstance(item, dict):
                continue
            rows.append({
                "title": _clean_title(item.get("title") or item.get("name"), 96),
                "unique_reads": _to_int(item.get("unique_reads", item.get("reads", 0))),
                "total_reads": _to_int(item.get("total_reads", item.get("views", 0))),
                "month": row.get("month"),
            })
    rows.sort(key=lambda item: (item.get("unique_reads", 0), item.get("total_reads", 0)), reverse=True)
    return rows[:5]


def _safe_change(current: float, previous: float, comparable: bool) -> str:
    if not comparable:
        return "No comparable por alcance de fuente"
    if not previous:
        return "Sin datos previos"
    return f"{round(((current - previous) / previous) * 100, 1)}%"


def compute_kpis(monthly_summaries: list[dict]) -> dict[str, Any]:
    normalized_rows = [validate_monthly_summary_contract(summary) for summary in monthly_summaries]

    plan_total = sum(_to_int(row.get("plan_total")) for row in normalized_rows)
    site_notes_total = sum(_to_int(row.get("site_notes_total")) for row in normalized_rows)
    site_total_views = sum(_to_int(row.get("site_total_views")) for row in normalized_rows)
    mail_total = sum(_to_int(row.get("mail_total")) for row in normalized_rows)

    weighted_open_num = 0.0
    weighted_inter_num = 0.0
    weighted_den = 0
    for row in normalized_rows:
        row_mail_total = _to_int(row.get("mail_total"))
        if row_mail_total <= 0:
            continue
        weighted_open_num += _normalize_pct(row.get("mail_open_rate")) * row_mail_total
        weighted_inter_num += _normalize_pct(row.get("mail_interaction_rate")) * row_mail_total
        weighted_den += row_mail_total

    mail_open_rate = round(weighted_open_num / weighted_den, 2) if weighted_den else 0.0
    mail_interaction_rate = round(weighted_inter_num / weighted_den, 2) if weighted_den else 0.0

    strategic_axes, strategic_axes_is_distribution = _aggregate_distribution([row.get("strategic_axes", []) for row in normalized_rows])
    internal_clients, internal_clients_is_distribution = _aggregate_distribution([row.get("internal_clients", []) for row in normalized_rows])
    channel_mix, channel_mix_is_distribution = _aggregate_distribution([row.get("channel_mix", []) for row in normalized_rows])
    format_mix, format_mix_is_distribution = _aggregate_distribution([row.get("format_mix", []) for row in normalized_rows])
    strategic_axes = strategic_axes[:6]
    internal_clients = internal_clients[:6]
    channel_mix = channel_mix[:6]
    format_mix = format_mix[:6]

    top_push_by_interaction = _top_push(normalized_rows, "top_push_by_interaction", "interaction")
    top_push_by_open_rate = _top_push(normalized_rows, "top_push_by_open_rate", "open_rate")
    top_pull_notes = _top_pull(normalized_rows)

    hitos: list[dict[str, Any]] = []
    events: list[dict[str, Any]] = []
    quality_flags = deepcopy(normalized_rows[-1].get("quality_flags", {})) if normalized_rows else {}
    scope_country_values = set()
    scope_mixed = False
    site_has_no_data = False
    push_ranking_available = False
    pull_ranking_available = False
    events_available = False
    historical_allowed = True

    for row in normalized_rows:
        for hito in row.get("hitos", []) or []:
            if not isinstance(hito, dict):
                continue
            title = _clean_title(hito.get("title") or hito.get("description") or hito.get("name"), 88)
            if not title:
                continue
            hitos.append({
                "period": row.get("month"),
                "title": title,
                "description": _clean_title(hito.get("description") or title, 160),
                "bullets": [
                    _clean_title(b, 92)
                    for b in (hito.get("bullets") or [])
                    if str(b).strip()
                ][:3],
                "thumbnail_path": str(hito.get("thumbnail_path") or ""),
            })
        for event in row.get("events", []) or []:
            if not isinstance(event, dict):
                continue
            name = _clean_title(event.get("name") or event.get("title"), 88)
            if not name:
                continue
            events.append({
                "name": name,
                "participants": _to_int(event.get("participants", event.get("attendees", 0))),
                "date": str(event.get("date") or row.get("month") or ""),
                "description": _clean_title(event.get("description"), 140),
            })

        row_flags = row.get("quality_flags", {})
        scope_country_values.add(str(row_flags.get("scope_country") or "AR"))
        scope_mixed = scope_mixed or bool(row_flags.get("scope_mixed", False))
        site_has_no_data = site_has_no_data or bool(row_flags.get("site_has_no_data_sections", False))
        push_ranking_available = push_ranking_available or bool(row_flags.get("push_ranking_available", False))
        pull_ranking_available = pull_ranking_available or bool(row_flags.get("pull_ranking_available", False))
        events_available = events_available or bool(row_flags.get("events_summary_available", False))
        historical_allowed = historical_allowed and bool(row_flags.get("historical_comparison_allowed", True))

    if len(scope_country_values) > 1:
        scope_country = ",".join(sorted(scope_country_values))
    else:
        scope_country = next(iter(scope_country_values), "AR")

    quality_flags.update({
        "scope_country": scope_country,
        "scope_mixed": scope_mixed or len(scope_country_values) > 1,
        "site_has_no_data_sections": site_has_no_data,
        "events_summary_available": events_available and bool(events),
        "push_ranking_available": push_ranking_available and bool(top_push_by_interaction or top_push_by_open_rate),
        "pull_ranking_available": pull_ranking_available and bool(top_pull_notes),
        "historical_comparison_allowed": historical_allowed and not (scope_mixed or len(scope_country_values) > 1),
    })

    events.sort(key=lambda item: item.get("participants", 0), reverse=True)
    total_events = len(events)
    total_event_participants = sum(item.get("participants", 0) for item in events)

    push_timeline = [{"label": row.get("month"), "value": row.get("mail_total", 0)} for row in normalized_rows]
    pull_timeline = [{"label": row.get("month"), "value": row.get("site_notes_total", 0)} for row in normalized_rows]
    latest_push = _to_int(push_timeline[-1]["value"]) if push_timeline else 0
    previous_push = _to_int(push_timeline[-2]["value"]) if len(push_timeline) > 1 else 0
    validation_warnings: list[str] = []
    if mail_total and _exceeds_plan_mail_threshold(plan_total, mail_total):
        validation_warnings.append(
            f"Inconsistencia potencial entre plan_total ({plan_total}) y mail_total ({mail_total})"
        )
    if mail_open_rate < 0 or mail_open_rate > 100:
        validation_warnings.append(f"mail_open_rate fuera de rango: {mail_open_rate}")
    if mail_interaction_rate < 0 or mail_interaction_rate > 100:
        validation_warnings.append(f"mail_interaction_rate fuera de rango: {mail_interaction_rate}")
    if strategic_axes_is_distribution or internal_clients_is_distribution or channel_mix_is_distribution or format_mix_is_distribution:
        validation_warnings.append("Mixes consolidados como promedio de distribución mensual (no suma directa)")

    return {
        "monthly_contract": normalized_rows,
        "quality_flags": quality_flags,
        "calculated_totals": {
            "plan_total": plan_total,
            "site_notes_total": site_notes_total,
            "site_total_views": site_total_views,
            "mail_total": mail_total,
            "mail_open_rate": mail_open_rate,
            "mail_interaction_rate": mail_interaction_rate,
            "total_events": total_events,
            "total_event_participants": total_event_participants,
            "latest_push_volume": latest_push,
            "previous_push_volume": previous_push,
            "latest_push_variation": _safe_change(latest_push, previous_push, quality_flags.get("historical_comparison_allowed", True)),
            "average_reads_per_note": round(site_total_views / site_notes_total, 2) if site_notes_total else 0.0,
            # compat aliases
            "push_volume_period": plan_total,
            "pull_notes_period": site_notes_total,
            "pull_reads_period": site_total_views,
            "average_open_rate": mail_open_rate,
            "average_interaction_rate": mail_interaction_rate,
        },
        "mixes": {
            "strategic_axes": strategic_axes,
            "internal_clients": internal_clients,
            "channel_mix": channel_mix,
            "format_mix": format_mix,
        },
        "consolidated_rankings": {
            "top_push_by_interaction": top_push_by_interaction,
            "top_push_by_open_rate": top_push_by_open_rate,
            "top_pull_notes": top_pull_notes,
            # compat aliases
            "top_push": top_push_by_interaction,
            "top_pull": top_pull_notes,
        },
        "timelines": {
            "mail_total": push_timeline,
            "site_notes_total": pull_timeline,
            # compat aliases
            "push_volume": push_timeline,
            "pull_notes": pull_timeline,
        },
        "hitos": hitos[:6],
        "events": events[:10],
        "hitos_crudos": hitos[:6],
        "aggregated_distributions": {
            "audience_segments": channel_mix[:5],
            "strategic_axes": [{"theme": item["label"], "weight": item["value"]} for item in strategic_axes],
            "internal_clients": internal_clients,
        },
        "validation": {
            "is_valid": True,
            "warnings": validation_warnings,
            "mix_aggregation": {
                "strategic_axes": "distribution_average" if strategic_axes_is_distribution else "weighted_sum",
                "internal_clients": "distribution_average" if internal_clients_is_distribution else "weighted_sum",
                "channel_mix": "distribution_average" if channel_mix_is_distribution else "weighted_sum",
                "format_mix": "distribution_average" if format_mix_is_distribution else "weighted_sum",
            },
        },
    }


def build_render_plan(period: dict[str, Any], kpis: dict[str, Any], narrative: dict[str, Any]) -> dict[str, Any]:
    totals = kpis.get("calculated_totals", {})
    mixes = kpis.get("mixes", {})
    rankings = kpis.get("consolidated_rankings", {})
    flags = kpis.get("quality_flags", {})
    hitos = kpis.get("hitos", [])
    events = kpis.get("events", [])

    overview_module = {
        "key": "executive_summary",
        "title": "Resumen ejecutivo del período",
        "payload": {
            "headline": "Resumen ejecutivo del período",
            "plan_total": totals.get("plan_total", 0),
            "site_notes_total": totals.get("site_notes_total", 0),
            "site_total_views": totals.get("site_total_views", 0),
            "mail_total": totals.get("mail_total", 0),
            "mail_open_rate": totals.get("mail_open_rate", 0),
            "mail_interaction_rate": totals.get("mail_interaction_rate", 0),
            "historical_note": narrative.get("executive_summary")
            or ("No comparable por alcance de fuente" if not flags.get("historical_comparison_allowed", True) else "Comparación histórica disponible para el alcance actual."),
            "takeaways": (narrative.get("executive_takeaways") if isinstance(narrative.get("executive_takeaways"), list) else [])[:3],
        },
    }

    modules = [
        overview_module,
        {
            "key": "channel_management",
            "title": "Gestión de canales",
            "payload": {
                "mail_total": totals.get("mail_total", 0),
                "mail_open_rate": totals.get("mail_open_rate", 0),
                "mail_interaction_rate": totals.get("mail_interaction_rate", 0),
                "site_notes_total": totals.get("site_notes_total", 0),
                "site_total_views": totals.get("site_total_views", 0),
                "channel_mix": mixes.get("channel_mix", []),
                "timeline_mail": kpis.get("timelines", {}).get("mail_total", []),
                "timeline_site": kpis.get("timelines", {}).get("site_notes_total", []),
                "message": narrative.get("channel_management", "Gestión consolidada de canales del período."),
                "site_has_no_data_sections": flags.get("site_has_no_data_sections", False),
            },
        },
        {
            "key": "mix_thematic_clients",
            "title": "Mix temático y áreas solicitantes",
            "payload": {
                "strategic_axes": mixes.get("strategic_axes", []),
                "internal_clients": mixes.get("internal_clients", []),
                "format_mix": mixes.get("format_mix", []),
                "message": narrative.get("mix_thematic_clients", "La agenda combina ejes estratégicos y demanda interna."),
            },
        },
        {
            "key": "ranking_push",
            "title": "Ranking push",
            "payload": {
                "by_interaction": rankings.get("top_push_by_interaction", []),
                "by_open_rate": rankings.get("top_push_by_open_rate", []),
                "available": flags.get("push_ranking_available", False),
                "message": narrative.get("ranking_push", "El ranking push resume piezas con mejor respuesta."),
            },
        },
        {
            "key": "ranking_pull",
            "title": "Ranking pull",
            "payload": {
                "top_pull_notes": rankings.get("top_pull_notes", []),
                "available": flags.get("pull_ranking_available", False),
                "average_reads_per_note": totals.get("average_reads_per_note", 0),
                "site_total_views": totals.get("site_total_views", 0),
                "message": narrative.get("ranking_pull", "El ranking pull identifica contenidos con mayor lectura."),
            },
        },
        {
            "key": "milestones",
            "title": "Hitos del mes",
            "payload": {
                "items": hitos,
                "message": narrative.get("milestones", "Hitos relevantes de gestión del período."),
            },
        },
    ]

    if flags.get("events_summary_available") and events:
        modules.append(
            {
                "key": "events",
                "title": "Eventos del mes",
                "payload": {
                    "events": events,
                    "total_events": totals.get("total_events", len(events)),
                    "total_participants": totals.get("total_event_participants", 0),
                    "message": narrative.get("events", "Los eventos del período tuvieron alcance medible."),
                },
            }
        )

    modules = [module for module in modules if isinstance(module, dict) and module.get("payload") is not None]

    return {
        "template_mode": "frame",
        "period": {
            "slug": period.get("slug"),
            "label": period.get("label"),
        },
        "quality_flags": flags,
        "modules": modules,
    }


def validate_report_json(report: Any) -> dict[str, Any]:
    if not isinstance(report, dict):
        return deepcopy(BASE_STRUCTURE)
    merged = deepcopy(BASE_STRUCTURE)
    for key in BASE_STRUCTURE:
        if key in report:
            merged[key] = report[key]
    if not isinstance(merged.get("render_plan"), dict):
        merged["render_plan"] = {"modules": []}
    if not isinstance(merged["render_plan"].get("modules"), list):
        merged["render_plan"]["modules"] = []
    return merged
