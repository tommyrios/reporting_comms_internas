from __future__ import annotations

import json
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.append(str(SCRIPT_DIR))

import extract_pdf_text as extract_step
import fetch_dashboard_pdfs as fetch_step
from generate_report import generate_period_report
from reporting_periods import resolve_schedule_from_env, save_schedule
from send_email import send_period_report

DATA_DIR = Path("data")
RUN_SUMMARY_PATH = DATA_DIR / "run_summary.json"


def main() -> None:
    schedule = resolve_schedule_from_env()
    save_schedule(schedule)

    if not schedule.periods:
        payload = {
            "status": "skipped",
            "reason": "No hay cierres trimestrales o anuales para este run.",
            "periods": [],
        }
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        RUN_SUMMARY_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    fetch_step.main()
    extract_step.main()

    sent_periods = []
    for period in schedule.periods:
        result = generate_period_report(period.slug)
        send_period_report(period.slug)
        sent_periods.append(
            {
                "period_slug": period.slug,
                "email_subject": result["report"].get("email_subject", period.email_subject),
                "report_dir": result["report_dir"],
            }
        )

    payload = {
        "status": "ok",
        "periods_sent": sent_periods,
    }
    RUN_SUMMARY_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
