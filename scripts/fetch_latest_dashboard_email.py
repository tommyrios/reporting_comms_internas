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

TARGET_PHRASE = "dashboard communications"

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


def find_pdf_part(payload: dict) -> Optional[dict]:
    """Busca recursivamente un adjunto PDF"""
    filename = payload.get("filename", "") or ""
    body = payload.get("body", {}) or ""

    if filename.lower().endswith(".pdf") and body.get("attachmentId"):
        return payload

    for part in payload.get("parts", []) or []:
        found = find_pdf_part(part)
        if found:
            return found

    return None


def extract_headers(payload: dict) -> dict:
    headers = payload.get("headers", [])
    return {h["name"]: h["value"] for h in headers}


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    service = build_gmail_service()

    # 🔥 Query amplia (no restrictiva)
    query = 'has:attachment filename:pdf'

    result = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=20
    ).execute()

    messages = result.get("messages", [])

    if not messages:
        raise RuntimeError("No se encontraron mails con PDFs.")

    print(f"Se encontraron {len(messages)} mails con adjuntos PDF")

    candidates = []

    for msg in messages:
        full_msg = service.users().messages().get(
            userId="me",
            id=msg["id"],
            format="full"
        ).execute()

        payload = full_msg.get("payload", {})
        headers = extract_headers(payload)

        subject = (headers.get("Subject") or "").lower()

        print(f"Mail evaluado: {subject}")

        # 🔥 filtro flexible por contenido
        if TARGET_PHRASE in subject:
            pdf_part = find_pdf_part(payload)

            if pdf_part:
                candidates.append({
                    "msg": full_msg,
                    "pdf_part": pdf_part,
                    "subject": subject,
                })

    if not candidates:
        raise RuntimeError("No se encontró ningún mail con el subject esperado.")

    # 🔥 tomar el más reciente
    selected = max(
        candidates,
        key=lambda x: int(x["msg"].get("internalDate", 0))
    )

    message = selected["msg"]
    pdf_part = selected["pdf_part"]

    attachment_id = pdf_part["body"]["attachmentId"]
    filename = pdf_part.get("filename", "latest_dashboard.pdf")

    print(f"Seleccionado: {selected['subject']}")

    attachment = service.users().messages().attachments().get(
        userId="me",
        messageId=message["id"],
        id=attachment_id
    ).execute()

    file_data = base64.urlsafe_b64decode(attachment["data"].encode("utf-8"))
    PDF_PATH.write_bytes(file_data)

    headers = extract_headers(message["payload"])

    metadata = {
        "message_id": message["id"],
        "subject": headers.get("Subject"),
        "from": headers.get("From"),
        "date": headers.get("Date"),
        "filename": filename,
        "pdf_path": str(PDF_PATH),
    }

    META_PATH.write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    print("PDF descargado correctamente")
    print(json.dumps(metadata, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
