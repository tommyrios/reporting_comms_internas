import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

SMTP_HOST = os.environ["SMTP_HOST"]
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))
SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASSWORD = os.environ["SMTP_PASSWORD"]

REPORT_TO = os.environ["REPORT_TO"]
REPORT_CC = os.environ.get("REPORT_CC", "")
REPORT_BCC = os.environ.get("REPORT_BCC", "")

SUBJECT = os.environ.get("REPORT_SUBJECT", "Reporte automático | Dashboard Communications")
FROM_NAME = os.environ.get("REPORT_FROM_NAME", "Automatización Comms Internas")

HTML_PATH = Path("output/report.html")
TEXT_PATH = Path("output/report.txt")


def load_body():
    html = HTML_PATH.read_text(encoding="utf-8") if HTML_PATH.exists() else None
    text = TEXT_PATH.read_text(encoding="utf-8") if TEXT_PATH.exists() else "No se generó versión texto del reporte."
    return text, html


def build_message():
    text_body, html_body = load_body()

    msg = EmailMessage()
    msg["Subject"] = SUBJECT
    msg["From"] = f"{FROM_NAME} <{SMTP_USER}>"
    msg["To"] = REPORT_TO

    if REPORT_CC.strip():
        msg["Cc"] = REPORT_CC
    if REPORT_BCC.strip():
        msg["Bcc"] = REPORT_BCC

    msg.set_content(text_body)

    if html_body:
        msg.add_alternative(html_body, subtype="html")

    return msg


def main():
    msg = build_message()

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)

    print("Reporte enviado correctamente por Gmail.")


if __name__ == "__main__":
    main()