from typing import Any

def compute_kpis(monthly_summaries: list[dict]) -> dict:
    total_push_volume = 0
    total_pull_notes = 0
    total_pull_reads = 0
    nps_list = []
    all_push_comms = []
    all_pull_notes = []
    hitos_consolidados = []
    
    for s in monthly_summaries:
        data = s.get("data", {})
        insights = s.get("insights", {})
        
        total_push_volume += int(data.get("push_volume") or 0)
        total_pull_notes += int(data.get("pull_notes") or 0)
        total_pull_reads += int(data.get("pull_reads") or 0)
        
        if data.get("nps"):
            nps_list.append(int(data.get("nps")))
            
        if insights.get("top_push_comm") and isinstance(insights.get("top_push_comm"), dict):
            all_push_comms.append(insights["top_push_comm"])
        if insights.get("top_pull_note") and isinstance(insights.get("top_pull_note"), dict):
            all_pull_notes.append(insights["top_pull_note"])
            
        if insights.get("hitos_mes"):
            hitos_consolidados.append({"month": s.get("month"), "hito": insights.get("hitos_mes")})

    avg_nps = sum(nps_list) // len(nps_list) if nps_list else 0
    avg_reads = total_pull_reads // total_pull_notes if total_pull_notes > 0 else 0
    
    return {
        "calculated_totals": {
            "push_volume_period": total_push_volume,
            "pull_notes_period": total_pull_notes,
            "pull_reads_period": total_pull_reads,
            "average_nps": avg_nps,
            "average_reads_per_note": avg_reads
        },
        "consolidated_rankings": {
            "top_push": all_push_comms[:3],
            "top_pull": all_pull_notes[:3]
        },
        "hitos_crudos": hitos_consolidados
    }

def validate_report_json(report: Any) -> dict[str, Any]:
    base_structure = {
        "slide_1_cover": {"area": "", "period": "-"},
        "slide_2_overview": {
            "volume_current": "-", "volume_previous": "-", "volume_change": "-",
            "audience_segments": [], "conclusion_message": "-"
        },
        "slide_3_strategy": {
            "content_distribution": [], "internal_clients": [],
            "canal_balance": {"institutional": 0, "transactional_talent": 0}
        },
        "slide_4_push_ranking": {
            "top_communications": [], "key_learning": "-"
        },
        "slide_5_pull_performance": {
            "pub_current": "-", "pub_previous": "-",
            "top_notes": [], "avg_reads": "-", "total_views": "-"
        },
        "slide_6_hitos": [],
        "slide_7_events": {
            "total_events": "-", "total_participants": "-", "conclusion": "-"
        }
    }

    if not isinstance(report, dict):
        return base_structure

    validated = base_structure.copy()
    for key, default_val in base_structure.items():
        if key in report and isinstance(report[key], type(default_val)):
            validated[key] = report[key]
            
    return validated