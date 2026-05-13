from __future__ import annotations

import argparse
import base64
import hashlib
import json
import logging
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List
from zoneinfo import ZoneInfo

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.append(str(SCRIPT_DIR))

from config import DATA_DIR, INBOX_PDF_DIR, ensure_dir
from period_scopes import (
    SCOPE_FILE_TOKENS,
    SCOPE_LABELS,
    infer_scope_from_text,
    period_scope_filename,
    required_scopes_from_env,
    normalize_text,
)
from reporting_periods import ReportingPeriod, resolve_schedule_from_env, save_schedule

logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
DEFAULT_SUBJECT_CONTAINS = "dashboard"
DEFAULT_QUERY = 'has:attachment filename:pdf newer_than:450d'

QUARTER_WORDS = {
    "q1": 1, "1q": 1, "1t": 1, "t1": 1, "primer trimestre": 1,
    "q2": 2, "2q": 2, "2t": 2, "t2": 2, "segundo trimestre": 2,
    "q3": 3, "3q": 3, "3t": 3, "t3": 3, "tercer trimestre": 3,
    "q4": 4, "4q": 4, "4t": 4, "t4": 4, "cuarto trimestre": 4,
}


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
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

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


def _period_by_slug(periods: list[ReportingPeriod]) -> dict[str, ReportingPeriod]:
    return {period.slug: period for period in periods}


def _infer_year(text: str, fallback_year: int) -> int:
    match = re.search(r"20\d{2}", text)
    return int(match.group(0)) if match else fallback_year


def infer_period_slug_from_text(text: str, periods: list[ReportingPeriod], fallback_year: int) -> str | None:
    """Detecta quarter/year desde asunto o filename.

    Si no detecta nada y hay un único período planificado, se usa ese período.
    Esto permite asuntos simples tipo "Dashboard CI" cuando la corrida ya viene
    parametrizada con REPORT_MODE=quarter y REPORT_YEAR/REPORT_QUARTER.
    """
    normalized = normalize_text(text)
    year = _infer_year(normalized, fallback_year)

    if re.search(r"\b(anual|year|ano|año)\b", normalized):
        slug = f"year_{year}"
        if any(period.slug == slug for period in periods):
            return slug

    # Patrones compactos: 2026-Q1, Q1 2026, 1Q 2026, 1T 2026.
    match = re.search(r"\b(20\d{2})\s*[-_/ ]?\s*q([1-4])\b", normalized)
    if match:
        slug = f"quarter_{match.group(1)}_Q{match.group(2)}"
        if any(period.slug == slug for period in periods):
            return slug
    match = re.search(r"\b(?:q|t)?([1-4])\s*(?:q|t)?\s*[-_/ ]?\s*(20\d{2})\b", normalized)
    if match:
        slug = f"quarter_{match.group(2)}_Q{match.group(1)}"
        if any(period.slug == slug for period in periods):
            return slug

    for marker, quarter in QUARTER_WORDS.items():
        if marker in normalized:
            slug = f"quarter_{year}_Q{quarter}"
            if any(period.slug == slug for period in periods):
                return slug

    if len(periods) == 1:
        return periods[0].slug

    return None


def period_slug_from_internal_date(internal_date_ms: str, tz_name: str, periods: list[ReportingPeriod]) -> str | None:
    dt = datetime.fromtimestamp(int(internal_date_ms) / 1000, tz=timezone.utc)
    local_dt = dt.astimezone(ZoneInfo(tz_name))
    for period in periods:
        start = datetime.fromisoformat(period.start_date).date()
        end = datetime.fromisoformat(period.end_date_exclusive).date()
        if start <= local_dt.date() < end:
            return period.slug
    return periods[0].slug if len(periods) == 1 else None


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
    required_scopes = required_scopes_from_env(os.environ.get("REPORT_REQUIRED_SCOPES"))

    if not schedule.periods:
        manifest = {
            "status": "skipped",
            "reason": "No hay períodos para generar en esta corrida.",
            "periods": [],
            "required_scopes": required_scopes,
            "files": [],
            "pdf_dir": str(pdf_dir),
        }
        ensure_dir(DATA_DIR)
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
    current_year = schedule.periods[0].year if schedule.periods else 2026

    logger.info(
        "event=email_fetch_started query=%s periods=%s scopes=%s",
        query,
        [period.slug for period in schedule.periods],
        required_scopes,
    )
    service = build_gmail_service()
    message_ids = iter_message_ids(service, query=query)
    print(f"Mensajes candidatos encontrados: {len(message_ids)}")

    candidates_by_period_scope: Dict[tuple[str, str], List[dict]] = defaultdict(list)

    for message_id in message_ids:
        full_message = service.users().messages().get(userId="me", id=message_id, format="full").execute()
        payload = full_message.get("payload", {}) or {}
        headers = extract_headers(payload)
        subject = (headers.get("Subject") or "").strip()
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

            period_slug = infer_period_slug_from_text(f"{subject} {filename}", schedule.periods, current_year)
            if not period_slug:
                period_slug = period_slug_from_internal_date(full_message.get("internalDate", "0"), schedule.timezone, schedule.periods)
            if not period_slug or period_slug not in {period.slug for period in schedule.periods}:
                continue

            scope = infer_scope_from_text(subject, filename)
            if not scope:
                # Si se pide un solo scope, toleramos filename sin marca explícita.
                scope = required_scopes[0] if len(required_scopes) == 1 else None
            if not scope or scope not in required_scopes:
                continue

            candidates_by_period_scope[(period_slug, scope)].append({
                "message": full_message,
                "message_id": full_message.get("id"),
                "internal_date": int(full_message.get("internalDate", "0")),
                "headers": headers,
                "subject": subject,
                "filename": filename,
                "attachment_id": pdf_part["body"]["attachmentId"],
            })

    selected_files = []
    missing: list[dict[str, str]] = []
    ensure_dir(pdf_dir)

    for period in schedule.periods:
        for scope in required_scopes:
            key = (period.slug, scope)
            candidates = candidates_by_period_scope.get(key, [])
            if not candidates:
                missing.append({"period": period.slug, "scope": scope})
                continue

            selected = max(candidates, key=lambda item: (item["internal_date"], item["message_id"], item["filename"]))
            file_bytes = download_attachment(
                service,
                message_id=selected["message_id"],
                attachment_id=selected["attachment_id"],
            )

            output_filename = period_scope_filename(period.slug, scope)
            output_path = pdf_dir / output_filename
            output_path.write_bytes(file_bytes)

            selected_files.append({
                "period": period.slug,
                "scope": scope,
                "scope_label": SCOPE_LABELS.get(scope, scope),
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
                "candidate_count_for_period_scope": len(candidates),
                "status": "downloaded",
            })
            print(f"Descargado {period.slug} [{scope}]: {selected['filename']} (Extraído de: {selected['subject']})")

    manifest = {
        "status": "ok" if not missing or allow_partial else "error",
        "periods_requested": [period.slug for period in schedule.periods],
        "required_scopes": required_scopes,
        "missing": missing,
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

    if missing and not allow_partial:
        readable = ", ".join(f"{item['period']}[{item['scope']}]" for item in missing)
        raise RuntimeError("Faltan PDFs para completar el período: " + readable)

    return manifest


def main(argv: list[str] | None = None) -> dict:
    parser = argparse.ArgumentParser(description="Descarga dashboards PDF trimestrales/anuales desde Gmail hacia un cache local.")
    parser.add_argument("--pdf-dir", type=Path, default=INBOX_PDF_DIR, help="Directorio local de PDFs descargados.")
    parser.add_argument("--manifest", type=Path, default=None, help="Ruta del manifest de descarga.")
    args = parser.parse_args(argv)
    return run_ingestion(pdf_dir=args.pdf_dir, manifest_path=args.manifest)


if __name__ == "__main__":
    main()
