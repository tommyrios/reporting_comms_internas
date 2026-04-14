from __future__ import annotations

import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.append(str(SCRIPT_DIR))

import extract_pdf_text as extract_step
import fetch_dashboard_pdfs as fetch_step
from generate_report import generate_period_report
from reporting_periods import resolve_schedule_from_env, save_schedule
from send_email import send_period_report

DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
RUN_SUMMARY_PATH = DATA_DIR / "run_summary.json"
REPORTS_DIR = OUTPUT_DIR / "reports"


def report_exists(period_slug: str) -> bool:
    period_dir = REPORTS_DIR / period_slug
    return (period_dir / "metadata.json").exists() and (period_dir / "report.html").exists()


def main() -> None:
    schedule = resolve_schedule_from_env()
    save_schedule(schedule)

    if not schedule.periods:
        payload = {
            "status": "skipped",
            "reason": "No hay períodos para generar en esta corrida.",
            "periods": [],
        }
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        RUN_SUMMARY_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    fetch_result = fetch_step.main()
    extract_step.main()

    results = []
    for period in schedule.periods:
        slug = period.slug
        print(f"Procesando período: {slug}")

        try:
            generate_result = generate_period_report(slug)
            results.append(
                {
                    "period": slug,
                    "generate_status": "ok",
                    "generation_mode": generate_result.get("generation_mode"),
                    "report_dir": generate_result.get("report_dir"),
                    "warning": generate_result.get("warning"),
                }
            )
        except Exception as exc:  # noqa: BLE001
            print(f"ERROR generando reporte para {slug}: {exc}", file=sys.stderr)
            results.append({"period": slug, "generate_status": "error", "error": str(exc)})
            continue

        if not report_exists(slug):
            print(
                f"Reporte no generado para {slug}: faltan metadata.json o report.html. Se omite envío.",
                file=sys.stderr,
            )
            results.append({"period": slug, "send_status": "skipped_missing_artifacts"})
            continue

        try:
            send_period_report(slug)
            results.append({"period": slug, "send_status": "sent"})
        except Exception as exc:  # noqa: BLE001
            print(f"ERROR enviando reporte para {slug}: {exc}", file=sys.stderr)
            results.append({"period": slug, "send_status": "error", "error": str(exc)})

    payload = {
        "status": "done",
        "fetch_result_type": type(fetch_result).__name__,
        "results": results,
    }
    RUN_SUMMARY_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
