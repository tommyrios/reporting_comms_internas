import base64
import os
from email.message import EmailMessage
from pathlib import Path

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

HTML_PATH = Path("output/report.html")
TEXT_PATH = Path("output/report.txt")


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


def load_body():
    html = HTML_PATH.read_text(encoding="utf-8") if HTML_PATH.exists() else None
    text = TEXT_PATH.read_text(encoding="utf-8") if TEXT_PATH.exists() else "No se generó versión texto."
    return text, html


def main():
    service = build_gmail_service()

    report_to = os.environ["REPORT_TO"]
    report_cc = os.environ.get("REPORT_CC", "")
    subject = os.environ.get("REPORT_SUBJECT", "Reporte automático | Dashboard Communications")
    from_name = os.environ.get("REPORT_FROM_NAME", "Automatización Comms Internas")
    gmail_user = os.environ["GMAIL_USER"]

    text_body, html_body = load_body()

    msg = EmailMessage()
    msg["To"] = report_to
    if report_cc.strip():
        msg["Cc"] = report_cc
    msg["From"] = f"{from_name} <{gmail_user}>"
    msg["Subject"] = subject
    msg.set_content(text_body)

    if html_body:
        msg.add_alternative(html_body, subtype="html")

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")

    service.users().messages().send(
        userId="me",
        body={"raw": raw_message}
    ).execute()

    print("Reporte enviado correctamente por Gmail API.")


if __name__ == "__main__":
    main()