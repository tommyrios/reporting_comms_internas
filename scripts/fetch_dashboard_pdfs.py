from __future__ import annotations

import base64
import json
import os
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional
from zoneinfo import ZoneInfo

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

import sys

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.append(str(SCRIPT_DIR))

from reporting_periods import (
    resolve_schedule_from_env,
    save_schedule,
    unique_months_from_periods,
)

DATA_DIR = Path("data")
PDF_DIR = DATA_DIR / "monthly_pdfs"
MANIFEST_PATH = DATA_DIR / "monthly_pdf_manifest.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
DEFAULT_SUBJECT_CONTAINS = "dashboard communications"
DEFAULT_QUERY = 'has:attachment filename:pdf newer_than:450d'


def build_gmail_service():
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        token_uri=os.environ.get("GOOGLE_TOKEN_URI", "https://oauth2.googleapis.com/token"),
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        scopes=SCOPES,
    )
    return build("gmail", "v1", credentials=creds)


def find_pdf_parts(payload: dict) -> List[dict]:
    found: List[dict] = []
    filename = (payload.get("filename") or "").strip()
    body = payload.get("body", {}) or {}

    if filename.lower().endswith(".pdf") and body.get("attachmentId"):
        found.append(payload)

    for part in payload.get("parts", []) or []:
        found.extend(find_pdf_parts(part))

    return found


def extract_headers(payload: dict) -> dict:
    headers = payload.get("headers", []) or []
    return {item["name"]: item["value"] for item in headers}


def month_slug_from_internal_date(internal_date_ms: str, tz_name: str) -> str:
    dt = datetime.fromtimestamp(int(internal_date_ms) / 1000, tz=timezone.utc)
    local_dt = dt.astimezone(ZoneInfo(tz_name))
    return f"{local_dt.year:04d}-{local_dt.month:02d}"


def iter_message_ids(service, query: str, max_pages: int = 10) -> List[str]:
    collected: List[str] = []
    request = service.users().messages().list(userId="me", q=query, maxResults=100)
    pages = 0

    while request is not None and pages < max_pages:
        response = request.execute()
        collected.extend(item["id"] for item in response.get("messages", []))
        request = service.users().messages().list_next(request, response)
        pages += 1

    return collected


def download_attachment(service, message_id: str, attachment_id: str) -> bytes:
    attachment = service.users().messages().attachments().get(
        userId="me",
        messageId=message_id,
        id=attachment_id,
    ).execute()
    return base64.urlsafe_b64decode(attachment["data"].encode("utf-8"))


def main() -> dict:
    schedule = resolve_schedule_from_env()
    save_schedule(schedule)

    if not schedule.periods:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        manifest = {
            "status": "skipped",
            "reason": "No hay períodos para generar en esta corrida.",
            "periods": [],
            "months_requested": [],
            "files": [],
        }
        MANIFEST_PATH.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print("No hay períodos para generar en esta corrida.")
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
        return manifest

    subject_contains = (os.environ.get("GMAIL_SUBJECT_CONTAINS") or DEFAULT_SUBJECT_CONTAINS).lower()
    query = os.environ.get("GMAIL_SEARCH_QUERY", DEFAULT_QUERY)
    allow_partial = (os.environ.get("ALLOW_PARTIAL_PERIOD") or "false").lower() == "true"

    months_needed = unique_months_from_periods(schedule.periods)

    service = build_gmail_service()
    message_ids = iter_message_ids(service, query=query)
    print(f"Mensajes candidatos encontrados: {len(message_ids)}")

    candidates_by_month: Dict[str, List[dict]] = defaultdict(list)

    for message_id in message_ids:
        full_message = service.users().messages().get(
            userId="me",
            id=message_id,
            format="full",
        ).execute()

        payload = full_message.get("payload", {}) or {}
        headers = extract_headers(payload)
        subject = (headers.get("Subject") or "").strip()
        subject_normalized = subject.lower()

        if subject_contains not in subject_normalized:
            continue

        pdf_parts = find_pdf_parts(payload)
        if not pdf_parts:
            continue

        month_slug = month_slug_from_internal_date(full_message.get("internalDate", "0"), schedule.timezone)
        if month_slug not in months_needed:
            continue

        for pdf_part in pdf_parts:
            filename = pdf_part.get("filename") or f"dashboard_{month_slug}.pdf"
            candidates_by_month[month_slug].append(
                {
                    "message": full_message,
                    "message_id": full_message.get("id"),
                    "internal_date": int(full_message.get("internalDate", "0")),
                    "headers": headers,
                    "subject": subject,
                    "filename": filename,
                    "attachment_id": pdf_part["body"]["attachmentId"],
                }
            )

    selected_files = []
    missing_months = []
    PDF_DIR.mkdir(parents=True, exist_ok=True)

    for month_slug in months_needed:
        month_candidates = candidates_by_month.get(month_slug, [])
        if not month_candidates:
            missing_months.append(month_slug)
            continue

        selected = max(month_candidates, key=lambda item: item["internal_date"])
        file_bytes = download_attachment(
            service,
            message_id=selected["message_id"],
            attachment_id=selected["attachment_id"],
        )

        output_path = PDF_DIR / f"{month_slug}.pdf"
        output_path.write_bytes(file_bytes)

        selected_files.append(
            {
                "month": month_slug,
                "pdf_path": str(output_path),
                "message_id": selected["message_id"],
                "subject": selected["subject"],
                "date": selected["headers"].get("Date"),
                "from": selected["headers"].get("From"),
                "filename": selected["filename"],
            }
        )

        print(f"Descargado {month_slug}: {selected['filename']}")

    manifest = {
        "status": "ok" if not missing_months or allow_partial else "error",
        "months_requested": months_needed,
        "missing_months": missing_months,
        "files": selected_files,
        "periods": [period.to_dict() for period in schedule.periods],
    }

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # Opcional: dejar también un alias más explícito
    (DATA_DIR / "fetch_result.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(json.dumps(manifest, ensure_ascii=False, indent=2))

    if missing_months and not allow_partial:
        raise RuntimeError(
            "Faltan PDFs para completar el período: " + ", ".join(missing_months)
        )

    return manifest


if __name__ == "__main__":
    main()