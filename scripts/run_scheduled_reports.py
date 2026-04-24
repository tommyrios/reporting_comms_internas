import json
import sys
from pathlib import Path

import fetch_dashboard_pdfs as fetch_step
from config import INBOX_PDF_DIR
from generate_report import generate_period_report
from send_email import send_period_report


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
REPORTS_DIR = BASE_DIR / "output" / "reports"


def report_exists(period_slug: str) -> bool:
    period_dir = REPORTS_DIR / period_slug
    return (period_dir / "metadata.json").exists() and (period_dir / "report.html").exists()


def _load_fetch_payload(fetch_result):
    if isinstance(fetch_result, dict):
        return fetch_result

    if isinstance(fetch_result, str):
        try:
            return json.loads(fetch_result)
        except json.JSONDecodeError:
            pass

    candidates = [
        DATA_DIR / "fetch_result.json",
        DATA_DIR / "selected_periods.json",
    ]

    for path in candidates:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))

    raise RuntimeError(
        "fetch_dashboard_pdfs no devolvió resultado y no existe ni data/fetch_result.json ni data/selected_periods.json"
    )


def main() -> None:
    fetch_result = fetch_step.run_ingestion()
    fetch_payload = _load_fetch_payload(fetch_result)
    pdf_dir = Path(fetch_payload.get("pdf_dir") or INBOX_PDF_DIR)

    periods = fetch_payload.get("periods", [])
    if not periods:
        raise RuntimeError("No se detectaron períodos a procesar.")

    results = []

    for period in periods:
        slug = period["slug"]
        print(f"Procesando período: {slug}")

        try:
            generation = generate_period_report(slug, pdf_dir=pdf_dir)
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

    has_errors = any(
    r.get("generate_status") == "error"
    or r.get("send_status") == "error"
    or r.get("send_status") == "skipped_missing_artifacts"
    for r in results
    )

    if has_errors:
        print(json.dumps({"status": "error", "results": results}, ensure_ascii=False, indent=2))
        raise SystemExit(1)

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
