from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from typing import Any


def _to_int(value: Any, default: int = 0) -> int:
    if value in (None, "", "-"):
        return default
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(round(value))
    text = str(value).strip().replace(".", "").replace(",", ".")
    try:
        return int(float(text))
    except Exception:
        digits = "".join(ch for ch in str(value) if ch.isdigit())
        return int(digits) if digits else default


def _to_float(value: Any, default: float = 0.0) -> float:
    if value in (None, "", "-"):
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("%", "").replace(".", "").replace(",", ".")
    try:
        return float(text)
    except Exception:
        filtered = "".join(ch for ch in str(value) if ch.isdigit() or ch in ".,-")
        if not filtered:
            return default
        filtered = filtered.replace(",", ".")
        try:
            return float(filtered)
        except Exception:
            return default


def _average_distribution(monthly_summaries: list[dict], key: str, label_key: str) -> list[dict[str, Any]]:
    accum: dict[str, float] = defaultdict(float)
    counts: dict[str, int] = defaultdict(int)
    for summary in monthly_summaries:
        items = summary.get("insights", {}).get(key, []) or []
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            label = str(item.get(label_key) or "").strip()
            if not label:
                continue
            value = _to_float(item.get("value", item.get("weight", 0)))
            accum[label] += value
            counts[label] += 1

    rows = []
    for label, total in accum.items():
        avg = total / max(counts[label], 1)
        rows.append({label_key: label, "value": round(avg, 1)})
    rows.sort(key=lambda item: item["value"], reverse=True)
    return rows


def _timeline(monthly_summaries: list[dict], metric_key: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for summary in monthly_summaries:
        rows.append({
            "label": str(summary.get("month") or ""),
            "value": _to_int(summary.get("data", {}).get(metric_key)),
        })
    return rows


def _safe_change(current: float, previous: float) -> str:
    if not previous:
        return "-"
    return f"{round(((current - previous) / previous) * 100, 1)}%"


def compute_kpis(monthly_summaries: list[dict]) -> dict[str, Any]:
    total_push_volume = 0
    total_pull_notes = 0
    total_pull_reads = 0
    open_rates: list[float] = []
    interaction_rates: list[float] = []
    all_push_comms: list[dict[str, Any]] = []
    all_pull_notes: list[dict[str, Any]] = []
    hitos_consolidados: list[dict[str, Any]] = []

    for summary in monthly_summaries:
        data = summary.get("data", {}) or {}
        insights = summary.get("insights", {}) or {}

        total_push_volume += _to_int(data.get("push_volume"))
        total_pull_notes += _to_int(data.get("pull_notes"))
        total_pull_reads += _to_int(data.get("pull_reads"))

        open_rate = _to_float(data.get("push_opens_pct"))
        if open_rate:
            open_rates.append(open_rate)
        interaction_rate = _to_float(data.get("push_interaction_pct"))
        if interaction_rate:
            interaction_rates.append(interaction_rate)

        top_push = insights.get("top_push_comm")
        if isinstance(top_push, dict):
            all_push_comms.append({
                "name": str(top_push.get("name") or "-").strip(),
                "clicks": _to_int(top_push.get("clicks")),
                "interaction": _to_float(top_push.get("interaction")),
                "month": summary.get("month"),
            })

        top_pull = insights.get("top_pull_note")
        if isinstance(top_pull, dict):
            all_pull_notes.append({
                "title": str(top_pull.get("title") or "-").strip(),
                "unique_reads": _to_int(top_pull.get("unique_reads")),
                "total_reads": _to_int(top_pull.get("total_reads")),
                "month": summary.get("month"),
            })

        hito = insights.get("hitos_mes")
        if hito:
            hitos_consolidados.append({
                "period": summary.get("month"),
                "description": str(hito).strip(),
            })

    all_push_comms.sort(key=lambda item: (item.get("clicks", 0), item.get("interaction", 0.0)), reverse=True)
    all_pull_notes.sort(key=lambda item: (item.get("unique_reads", 0), item.get("total_reads", 0)), reverse=True)

    avg_open_rate = round(sum(open_rates) / len(open_rates), 1) if open_rates else 0.0
    avg_interaction_rate = round(sum(interaction_rates) / len(interaction_rates), 1) if interaction_rates else 0.0
    avg_reads = round(total_pull_reads / total_pull_notes, 1) if total_pull_notes else 0.0

    push_timeline = _timeline(monthly_summaries, "push_volume")
    pull_timeline = _timeline(monthly_summaries, "pull_notes")
    latest_push = push_timeline[-1]["value"] if push_timeline else 0
    previous_push = push_timeline[-2]["value"] if len(push_timeline) > 1 else 0

    audience_segments_raw = _average_distribution(monthly_summaries, "audience_segmentation", "label")
    strategic_axes_raw = _average_distribution(monthly_summaries, "strategic_axes", "theme")
    internal_clients_raw = _average_distribution(monthly_summaries, "internal_clients", "label")

    audience_segments = [{"label": item["label"], "value": item["value"]} for item in audience_segments_raw[:6]]
    strategic_axes = [{"theme": item["theme"], "weight": item["value"]} for item in strategic_axes_raw[:6]]
    internal_clients = [{"label": item["label"], "value": item["value"]} for item in internal_clients_raw[:6]]

    return {
        "calculated_totals": {
            "push_volume_period": total_push_volume,
            "pull_notes_period": total_pull_notes,
            "pull_reads_period": total_pull_reads,
            "average_reads_per_note": avg_reads,
            "average_open_rate": avg_open_rate,
            "average_interaction_rate": avg_interaction_rate,
            "latest_push_volume": latest_push,
            "previous_push_volume": previous_push,
            "latest_push_variation": _safe_change(latest_push, previous_push),
        },
        "consolidated_rankings": {
            "top_push": all_push_comms[:3],
            "top_pull": all_pull_notes[:5],
        },
        "aggregated_distributions": {
            "audience_segments": audience_segments,
            "strategic_axes": strategic_axes,
            "internal_clients": internal_clients,
        },
        "timelines": {
            "push_volume": push_timeline,
            "pull_notes": pull_timeline,
        },
        "hitos_crudos": hitos_consolidados,
    }


BASE_STRUCTURE: dict[str, Any] = {
    "slide_1_cover": {
        "area": "Comunicaciones Internas",
        "period": "-",
        "subtitle": "Informe de gestión",
    },
    "slide_2_overview": {
        "headline": "¿Cómo nos fue? CI",
        "volume_current": "-",
        "volume_previous": "-",
        "volume_change": "-",
        "push_open_rate": "-",
        "push_interaction_rate": "-",
        "pull_notes_current": "-",
        "average_reads": "-",
        "audience_segments": [],
        "comparison_timeline": [],
        "comparative_note": "-",
        "conclusion_message": "-",
        "highlights": [],
    },
    "slide_3_plan": {
        "title": "Gestión del plan CI",
        "mail_total": "-",
        "segmented_share": "-",
        "open_rate": "-",
        "pull_total": "-",
        "mail_timeline": [],
        "pull_timeline": [],
        "mail_message": "-",
        "pull_message": "-",
        "footer": "-",
    },
    "slide_4_strategy": {
        "content_distribution": [],
        "internal_clients": [],
        "canal_balance": {"institutional": 0, "transactional_talent": 0},
        "theme_message": "-",
        "balance_message": "-",
        "conclusion": "-",
    },
    "slide_5_push_ranking": {
        "top_communications": [],
        "key_learning": "-",
    },
    "slide_6_pull_performance": {
        "pub_current": "-",
        "pub_previous": "-",
        "top_notes": [],
        "avg_reads": "-",
        "total_views": "-",
        "secondary_message": "-",
        "conclusion": "-",
    },
    "slide_7_hitos": [],
    "slide_8_events": {
        "total_events": "-",
        "total_participants": "-",
        "secondary_message": "-",
        "conclusion": "-",
        "event_breakdown": [],
    },
    "slide_9_closure": {
        "title": "Claves del período",
        "bullets": [],
    },
}


def validate_report_json(report: Any) -> dict[str, Any]:
    if not isinstance(report, dict):
        return deepcopy(BASE_STRUCTURE)

    validated = {
        key: value.copy() if isinstance(value, dict) else list(value)
        for key, value in BASE_STRUCTURE.items()
    }
    for key, default_value in BASE_STRUCTURE.items():
        candidate = report.get(key)
        if isinstance(default_value, dict) and isinstance(candidate, dict):
            merged = default_value.copy()
            merged.update(candidate)
            validated[key] = merged
        elif isinstance(default_value, list) and isinstance(candidate, list):
            validated[key] = candidate

    legacy_map = {
        "slide_3_strategy": "slide_4_strategy",
        "slide_4_push_ranking": "slide_5_push_ranking",
        "slide_5_pull_performance": "slide_6_pull_performance",
        "slide_6_hitos": "slide_7_hitos",
        "slide_7_events": "slide_8_events",
        "slide_8_closure": "slide_9_closure",
    }
    for old_key, new_key in legacy_map.items():
        if old_key in report and new_key in BASE_STRUCTURE:
            if isinstance(BASE_STRUCTURE[new_key], dict) and isinstance(report[old_key], dict):
                merged = validated[new_key].copy()
                merged.update(report[old_key])
                validated[new_key] = merged
            elif isinstance(BASE_STRUCTURE[new_key], list) and isinstance(report[old_key], list):
                validated[new_key] = report[old_key]

    return validated
