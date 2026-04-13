from __future__ import annotations

import json
import mimetypes
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import List, Optional

OUTPUT_DIR = Path("output")
DEFAULT_REPORTS_DIR = OUTPUT_DIR / "reports"


def _split_recipients(raw: str) -> List[str]:
    return [item.strip() for item in (raw or "").split(",") if item.strip()]


class EmailSender:
    def __init__(self) -> None:
        self.smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
        self.smtp_port = int(os.environ.get("SMTP_PORT", "587"))
        self.smtp_username = os.environ.get("SMTP_USERNAME") or os.environ.get("EMAIL_USER")
        self.smtp_password = os.environ.get("SMTP_PASSWORD") or os.environ.get("EMAIL_PASSWORD")
        self.report_from = os.environ.get("REPORT_FROM") or self.smtp_username
        self.report_to = _split_recipients(
            os.environ.get("REPORT_TO") or os.environ.get("EMAIL_DESTINATARIO", "")
        )
        self.report_cc = _split_recipients(
            os.environ.get("REPORT_CC") or os.environ.get("EMAIL_CC", "")
        )

        if not self.smtp_username or not self.smtp_password:
            raise ValueError("Faltan credenciales SMTP")
        if not self.report_to:
            raise ValueError("No hay destinatarios definidos en REPORT_TO o EMAIL_DESTINATARIO")

    def send_message(
        self,
        subject: str,
        html_body: str,
        text_body: str,
        attachments: Optional[List[Path]] = None,
    ) -> None:
        mixed_root = MIMEMultipart("mixed")
        mixed_root["Subject"] = subject
        mixed_root["From"] = self.report_from or self.smtp_username
        mixed_root["To"] = ", ".join(self.report_to)
        if self.report_cc:
            mixed_root["Cc"] = ", ".join(self.report_cc)

        alternative_part = MIMEMultipart("alternative")
        alternative_part.attach(MIMEText(text_body, "plain", "utf-8"))
        alternative_part.attach(MIMEText(html_body, "html", "utf-8"))
        mixed_root.attach(alternative_part)

        for attachment_path in attachments or []:
            if not attachment_path.exists():
                continue
            ctype, _ = mimetypes.guess_type(str(attachment_path))
            maintype, subtype = (ctype.split("/", 1) if ctype else ("application", "octet-stream"))
            with attachment_path.open("rb") as handler:
                part = MIMEApplication(handler.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=attachment_path.name)
            mixed_root.attach(part)

        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            server.starttls()
            server.login(self.smtp_username, self.smtp_password)
            server.sendmail(
                self.report_from or self.smtp_username,
                self.report_to + self.report_cc,
                mixed_root.as_string(),
            )

        print(f"Correo enviado: {subject}")


def send_period_report(period_slug: str) -> None:
    report_dir = DEFAULT_REPORTS_DIR / period_slug
    html_path = report_dir / "report.html"
    text_path = report_dir / "report.txt"
    json_path = report_dir / "report.json"

    if not html_path.exists() or not text_path.exists() or not json_path.exists():
        raise FileNotFoundError(f"No existe el reporte completo para {period_slug}")

    report_json = json.loads(json_path.read_text(encoding="utf-8"))
    subject = os.environ.get("EMAIL_SUBJECT_OVERRIDE") or report_json.get("email_subject") or f"Reporte CI | {period_slug}"

    attachments: List[Path] = []
    if (os.environ.get("ATTACH_HTML_REPORT") or "false").lower() == "true":
        attachments.append(html_path)
    if (os.environ.get("ATTACH_JSON_REPORT") or "false").lower() == "true":
        attachments.append(json_path)

    sender = EmailSender()
    sender.send_message(
        subject=subject,
        html_body=html_path.read_text(encoding="utf-8"),
        text_body=text_path.read_text(encoding="utf-8"),
        attachments=attachments,
    )


def main() -> None:
    period_slug = os.environ.get("TARGET_PERIOD_SLUG")
    if not period_slug:
        raise ValueError("Falta TARGET_PERIOD_SLUG")
    send_period_report(period_slug)


if __name__ == "__main__":
    main()
