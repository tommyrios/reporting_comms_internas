import json
import sys
from pathlib import Path

import extract_pdf_text as extract_step
import fetch_dashboard_pdfs as fetch_step
from generate_report import generate_period_report
from send_email import send_period_report


BASE_DIR = Path(__file__).resolve().parent.parent
REPORTS_DIR = BASE_DIR / "output" / "reports"


def report_exists(period_slug: str) -> bool:
    period_dir = REPORTS_DIR / period_slug
    return (period_dir / "metadata.json").exists() and (period_dir / "report.html").exists()


def main() -> None:
    fetch_result = fetch_step.main()

    if fetch_result is None:
        periods_path = BASE_DIR / "data" / "selected_periods.json"
        if not periods_path.exists():
            raise RuntimeError("fetch_dashboard_pdfs no devolvió resultado y no existe selected_periods.json")
        fetch_result = json.loads(periods_path.read_text(encoding="utf-8"))
    elif isinstance(fetch_result, str):
        fetch_result = json.loads(fetch_result)

    periods = fetch_result.get("periods", [])
    if not periods:
        raise RuntimeError("No se detectaron períodos a procesar.")

    extract_step.main()

    results = []

    for period in periods:
        slug = period["slug"]
        print(f"Procesando período: {slug}")

        try:
            generation = generate_period_report(slug)
            results.append(
                {
                    "period": slug,
                    "generate_status": "ok",
                    "generation_mode": generation.get("generation_mode"),
                    "report_dir": generation.get("report_dir"),
                    "warning": generation.get("warning"),
                }
            )
        except Exception as e:
            print(f"ERROR generando reporte para {slug}: {e}", file=sys.stderr)
            results.append(
                {
                    "period": slug,
                    "generate_status": "error",
                    "error": str(e),
                }
            )
            continue

        if not report_exists(slug):
            print(
                f"Reporte no generado para {slug}: faltan metadata.json o report.html. Se omite envío.",
                file=sys.stderr,
            )
            results.append(
                {
                    "period": slug,
                    "send_status": "skipped_missing_artifacts",
                }
            )
            continue

        try:
            send_period_report(slug)
            results.append(
                {
                    "period": slug,
                    "send_status": "sent",
                }
            )
        except Exception as e:
            print(f"ERROR enviando reporte para {slug}: {e}", file=sys.stderr)
            results.append(
                {
                    "period": slug,
                    "send_status": "error",
                    "error": str(e),
                }
            )

    print(
        json.dumps(
            {
                "status": "done",
                "fetch_result_type": type(fetch_result).__name__,
                "results": results,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()