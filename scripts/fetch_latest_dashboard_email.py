import base64
import json
import os
from pathlib import Path
from typing import Optional

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

DATA_DIR = Path("data")
PDF_PATH = DATA_DIR / "latest_dashboard.pdf"
META_PATH = DATA_DIR / "metadata.json"

EMAIL_SUBJECT = 'Dashboard Communications | Comunicación interna'
PDF_FILENAME_CONTAINS = 'Dashboard Communications | Comunicación interna'

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


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


def find_attachment_part(payload: dict) -> Optional[dict]:
    parts = payload.get("parts", [])
    for part in parts:
        filename = part.get("filename", "") or ""
        body = part.get("body", {}) or {}

        if filename.lower().endswith(".pdf") and PDF_FILENAME_CONTAINS.lower() in filename.lower():
            if body.get("attachmentId"):
                return part

        nested_parts = part.get("parts", [])
        for nested in nested_parts:
            nested_filename = nested.get("filename", "") or ""
            nested_body = nested.get("body", {}) or {}
            if nested_filename.lower().endswith(".pdf") and PDF_FILENAME_CONTAINS.lower() in nested_filename.lower():
                if nested_body.get("attachmentId"):
                    return nested

    return None


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    service = build_gmail_service()

    query = f'subject:"{EMAIL_SUBJECT}" has:attachment filename:pdf'
    result = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=10
    ).execute()

    messages = result.get("messages", [])
    if not messages:
        raise RuntimeError("No se encontraron mails que coincidan con la búsqueda.")

    selected_message = None
    selected_attachment_part = None

    for msg in messages:
        full_msg = service.users().messages().get(
            userId="me",
            id=msg["id"],
            format="full"
        ).execute()

        payload = full_msg.get("payload", {})
        attachment_part = find_attachment_part(payload)

        if attachment_part:
            selected_message = full_msg
            selected_attachment_part = attachment_part
            break

    if not selected_message or not selected_attachment_part:
        raise RuntimeError("No se encontró un adjunto PDF válido en los mails encontrados.")

    attachment_id = selected_attachment_part["body"]["attachmentId"]
    filename = selected_attachment_part.get("filename", "latest_dashboard.pdf")

    attachment = service.users().messages().attachments().get(
        userId="me",
        messageId=selected_message["id"],
        id=attachment_id
    ).execute()

    file_data = base64.urlsafe_b64decode(attachment["data"].encode("UTF-8"))
    PDF_PATH.write_bytes(file_data)

    headers = selected_message.get("payload", {}).get("headers", [])
    header_map = {h["name"]: h["value"] for h in headers}

    metadata = {
        "message_id": selected_message["id"],
        "thread_id": selected_message.get("threadId"),
        "internal_date": selected_message.get("internalDate"),
        "subject": header_map.get("Subject"),
        "from": header_map.get("From"),
        "date": header_map.get("Date"),
        "attachment_filename": filename,
        "pdf_path": str(PDF_PATH),
    }

    META_PATH.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(metadata, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()