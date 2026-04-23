from __future__ import annotations

import argparse
import base64
import hashlib
import json
import logging
import os
import re
import unicodedata
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

from config import DATA_DIR, INBOX_PDF_DIR, ensure_dir
from reporting_periods import (
    resolve_schedule_from_env,
    save_schedule,
    unique_months_from_periods,
)

logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
DEFAULT_SUBJECT_CONTAINS = "dashboard" # Relajamos un poco el filtro
DEFAULT_QUERY = 'has:attachment filename:pdf newer_than:450d'

# Diccionario para mapear nombres de meses a su número
MONTHS_MAP = {
    "enero": "01", "january": "01", "jan": "01", "ene": "01",
    "febrero": "02", "february": "02", "feb": "02",
    "marzo": "03", "march": "03", "mar": "03",
    "abril": "04", "april": "04", "abr": "04", "apr": "04",
    "mayo": "05", "may": "05",
    "junio": "06", "june": "06", "jun": "06",
    "julio": "07", "july": "07", "jul": "07",
    "agosto": "08", "august": "08", "ago": "08", "aug": "08",
    "septiembre": "09", "september": "09", "sep": "09",
    "octubre": "10", "october": "10", "oct": "10",
    "noviembre": "11", "november": "11", "nov": "11",
    "diciembre": "12", "december": "12", "dic": "12", "dec": "12"
}

MONTH_NAME_PATTERN = re.compile(
    r"(?<![a-z0-9])(" + "|".join(sorted(MONTHS_MAP.keys(), key=len, reverse=True)) + r")(?![a-z0-9])"
)


def normalize_text(value: str) -> str:
    value = (value or "").strip().lower()
    normalized = unicodedata.normalize("NFD", value)
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def _keyword_to_pattern(keyword: str) -> re.Pattern[str]:
    escaped = re.escape(keyword.strip())
    escaped = escaped.replace(r"\ ", r"\s+")
    return re.compile(rf"(?<![a-z0-9]){escaped}(?![a-z0-9])")


def _has_expected_keywords(subject: str, filename: str, keywords: List[str]) -> bool:
    haystack = normalize_text(f"{subject} {filename}")
    if not keywords:
        return True
    return any(_keyword_to_pattern(normalize_text(keyword)).search(haystack) for keyword in keywords if keyword.strip())

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


def extract_month_from_text(text: str, default_year: int) -> str | None:
    """Busca menciones explícitas de un mes en el texto (ej. 'Enero', '01-2026')"""
    text = normalize_text(text)

    # 1. Busca formato explícito numérico (ej. 2026-01, 01/2026)
    match = re.search(r'(?<!\d)(20\d{2})[\-_/](0[1-9]|1[0-2])(?!\d)', text)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    match = re.search(r'(?<!\d)(0[1-9]|1[0-2])[\-_/](20\d{2})(?!\d)', text)
    if match:
        return f"{match.group(2)}-{match.group(1)}"

    # 2. Busca nombre del mes en texto
    month_match = MONTH_NAME_PATTERN.search(text)
    if month_match:
        word = month_match.group(1)
        num = MONTHS_MAP[word]
        context_start = max(0, month_match.start() - 24)
        context_end = min(len(text), month_match.end() + 24)
        local_context = text[context_start:context_end]
        year_match = re.search(r'(20\d{2})', local_context) or re.search(r'(20\d{2})', text)
        year = year_match.group(1) if year_match else str(default_year)
        return f"{year}-{num}"

    return None

def month_slug_from_internal_date(internal_date_ms: str, tz_name: str) -> str:
    dt = datetime.fromtimestamp(int(internal_date_ms) / 1000, tz=timezone.utc)
    local_dt = dt.astimezone(ZoneInfo(tz_name))
    return f"{local_dt.year:04d}-{local_dt.month:02d}"


def deterministic_pdf_filename(month_slug: str) -> str:
    return f"{month_slug}_dashboard.pdf"


def _manifest_path_for(pdf_dir: Path, manifest_path: Path | None = None) -> Path:
    if manifest_path:
        return manifest_path
    return pdf_dir / "manifest.json"


def _sha256_hex(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


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


def run_ingestion(pdf_dir: Path | None = None, manifest_path: Path | None = None) -> dict:
    schedule = resolve_schedule_from_env()
    save_schedule(schedule)
    pdf_dir = ensure_dir(pdf_dir or INBOX_PDF_DIR)
    manifest_path = _manifest_path_for(pdf_dir, manifest_path)

    if not schedule.periods:
        ensure_dir(DATA_DIR)
        manifest = {
            "status": "skipped",
            "reason": "No hay períodos para generar en esta corrida.",
            "periods": [],
            "months_requested": [],
            "files": [],
            "pdf_dir": str(pdf_dir),
        }
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        print("No hay períodos para generar en esta corrida.")
        return manifest

    subject_contains = normalize_text(os.environ.get("GMAIL_SUBJECT_CONTAINS") or DEFAULT_SUBJECT_CONTAINS)
    expected_sender = normalize_text(os.environ.get("GMAIL_EXPECTED_SENDER") or "")
    expected_keywords = [
        item.strip()
        for item in (os.environ.get("GMAIL_EXPECTED_KEYWORDS") or subject_contains).split(",")
        if item.strip()
    ]
    query = os.environ.get("GMAIL_SEARCH_QUERY", DEFAULT_QUERY)
    allow_partial = (os.environ.get("ALLOW_PARTIAL_PERIOD") or "false").lower() == "true"

    months_needed = unique_months_from_periods(schedule.periods)
    current_year = schedule.periods[0].year if schedule.periods else 2026

    logger.info("event=email_fetch_started query=%s months=%s", query, months_needed)
    service = build_gmail_service()
    message_ids = iter_message_ids(service, query=query)
    logger.info("event=email_candidates_found count=%s", len(message_ids))
    print(f"Mensajes candidatos encontrados: {len(message_ids)}")

    candidates_by_month: Dict[str, List[dict]] = defaultdict(list)

    for message_id in message_ids:
        full_message = service.users().messages().get(userId="me", id=message_id, format="full").execute()

        payload = full_message.get("payload", {}) or {}
        headers = extract_headers(payload)
        subject = (headers.get("Subject") or "").strip()
        subject_normalized = normalize_text(subject)
        sender_normalized = normalize_text(headers.get("From") or "")

        if subject_contains and not _has_expected_keywords(subject, "", [subject_contains]):
            continue

        if expected_sender and expected_sender not in sender_normalized:
            continue

        pdf_parts = find_pdf_parts(payload)
        if not pdf_parts:
            continue

        for pdf_part in pdf_parts:
            filename = pdf_part.get("filename") or "dashboard.pdf"
            if not _has_expected_keywords(subject, filename, expected_keywords):
                continue
            
            # MAGIA: Intentamos sacar el mes del Asunto o del Filename ANTES de usar la fecha del correo
            month_slug = extract_month_from_text(subject_normalized, current_year)
            if not month_slug:
                month_slug = extract_month_from_text(filename, current_year)
            if not month_slug:
                month_slug = month_slug_from_internal_date(full_message.get("internalDate", "0"), schedule.timezone)

            if month_slug not in months_needed:
                continue

            candidates_by_month[month_slug].append({
                "message": full_message,
                "message_id": full_message.get("id"),
                "internal_date": int(full_message.get("internalDate", "0")),
                "headers": headers,
                "subject": subject,
                "filename": filename,
                "attachment_id": pdf_part["body"]["attachmentId"],
            })

    selected_files = []
    missing_months = []
    ensure_dir(pdf_dir)
    candidate_total = sum(len(items) for items in candidates_by_month.values())
    logger.info("event=attachment_candidates_found total=%s by_month=%s", candidate_total, {k: len(v) for k, v in candidates_by_month.items()})

    for month_slug in months_needed:
        month_candidates = candidates_by_month.get(month_slug, [])
        if not month_candidates:
            missing_months.append(month_slug)
            continue

        selected = max(
            month_candidates,
            key=lambda item: (item["internal_date"], item["message_id"], item["filename"]),
        )
        file_bytes = download_attachment(
            service,
            message_id=selected["message_id"],
            attachment_id=selected["attachment_id"],
        )

        output_filename = deterministic_pdf_filename(month_slug)
        output_path = pdf_dir / output_filename
        output_path.write_bytes(file_bytes)
        logger.info("event=email_attachment_saved month=%s path=%s", month_slug, output_path)
        logger.info("event=month_assigned_to_file month=%s filename=%s", month_slug, output_filename)

        selected_files.append({
            "month": month_slug,
            "pdf_path": str(output_path),
            "saved_filename": output_filename,
            "message_id": selected["message_id"],
            "internal_date_ms": selected["internal_date"],
            "email_date": datetime.fromtimestamp(selected["internal_date"] / 1000, tz=timezone.utc).isoformat(),
            "subject": selected["subject"],
            "original_attachment_filename": selected["filename"],
            "downloaded_at": datetime.now(timezone.utc).isoformat(),
            "checksum_sha256": _sha256_hex(file_bytes),
            "selection_rule": "max(internal_date,message_id,filename)",
            "candidate_count_for_month": len(month_candidates),
            "status": "downloaded",
        })
        print(f"Descargado {month_slug}: {selected['filename']} (Extraído de: {selected['subject']})")

    manifest = {
        "status": "ok" if not missing_months or allow_partial else "error",
        "months_requested": months_needed,
        "missing_months": missing_months,
        "files": selected_files,
        "periods": [period.to_dict() for period in schedule.periods],
        "pdf_dir": str(pdf_dir),
        "manifest_path": str(manifest_path),
    }

    ensure_dir(DATA_DIR)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (DATA_DIR / "fetch_result.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (DATA_DIR / "selected_periods.json").write_text(json.dumps({"periods": manifest["periods"]}, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(manifest, ensure_ascii=False, indent=2))

    if missing_months and not allow_partial:
        raise RuntimeError("Faltan PDFs para completar el período: " + ", ".join(missing_months))

    return manifest


def main(argv: list[str] | None = None) -> dict:
    parser = argparse.ArgumentParser(description="Descarga dashboards PDF desde Gmail hacia un cache local.")
    parser.add_argument("--pdf-dir", type=Path, default=INBOX_PDF_DIR, help="Directorio local de PDFs descargados.")
    parser.add_argument("--manifest", type=Path, default=None, help="Ruta del manifest de descarga.")
    args = parser.parse_args(argv)
    return run_ingestion(pdf_dir=args.pdf_dir, manifest_path=args.manifest)


if __name__ == "__main__":
    main()
