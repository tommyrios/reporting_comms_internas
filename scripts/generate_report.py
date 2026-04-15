import json
import os
from datetime import datetime
from typing import Any

from config import DATA_DIR, REPORTS_DIR, ensure_dir
from llm_client import build_genai_client, call_gemini_for_json, load_prompt
from analyzer import compute_kpis
from pdf_processor import summarize_month
from pptx_renderer import create_pptx

def get_period_definition(period_slug: str) -> dict[str, Any]:
    payload = json.loads((DATA_DIR / "selected_periods.json").read_text(encoding="utf-8"))
    for p in payload.get("periods", []):
        if p.get("slug") == period_slug:
            return p
    raise KeyError(period_slug)

def write_report_artifacts(period_slug: str, report: dict[str, Any]) -> str:
    report_dir = ensure_dir(REPORTS_DIR / period_slug)
    
    metadata = {
        "title": report.get("slide_1_cover", {}).get("area"),
        "period_slug": period_slug,
        "generated_at": datetime.utcnow().isoformat() + "Z",
    }
    (report_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    (report_dir / "report_raw.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    
    create_pptx(report, report_dir / "report.pptx")
    
    html_content = f"<html><body><h2>Reporte Generado</h2><p>El reporte {period_slug} está en PowerPoint (.pptx).</p></body></html>"
    (report_dir / "report.html").write_text(html_content, encoding="utf-8")
    
    return str(report_dir)

def generate_period_report(period_slug: str, force_regenerate: bool = False) -> dict[str, Any]:
    period = get_period_definition(period_slug)
    client = build_genai_client()
    
    summaries = [summarize_month(client, m, force_regenerate) for m in period.get("months", [])]
    kpis_calculados = compute_kpis(summaries)
    
    prompt_base = load_prompt("period_report.txt")
    prompt_final = f"{prompt_base}\n\nINPUT (KPI_CALCULADOS):\n{json.dumps(kpis_calculados, ensure_ascii=False)}\n\nCONTEXTO (RESUMENES_MENSUALES):\n{json.dumps(summaries, ensure_ascii=False)}"
    
    report_raw = call_gemini_for_json(client, [prompt_final])
    report_dir = write_report_artifacts(period_slug, report_raw)
    
    return {"status": "ok", "period_slug": period_slug, "report_dir": report_dir}

def main():
    period_slug = os.environ.get("REPORT_SLUG", "").strip()
    if period_slug:
        print(json.dumps(generate_period_report(period_slug), ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()