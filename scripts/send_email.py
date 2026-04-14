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

SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent
OUTPUT_DIR = BASE_DIR / "output"
REPORTS_DIR = OUTPUT_DIR / "reports"


def _split_recipients(raw: str) -> List[str]:
    return [item.strip() for item in (raw or "").split(",") if item.strip()]


class EmailSender:
    def __init__(self) -> None:
        self.email_user = (os.environ.get("EMAIL_USER") or "").strip()
        self.email_password = (os.environ.get("EMAIL_PASSWORD") or "").strip()
        self.report_to = _split_recipients(os.environ.get("EMAIL_DESTINATARIO") or os.environ.get("REPORT_TO") or "")
        self.report_cc = _split_recipients(os.environ.get("EMAIL_CC") or os.environ.get("REPORT_CC") or "")
        self.report_from = (os.environ.get("REPORT_FROM") or self.email_user).strip()

        self.smtp_host = "smtp.gmail.com"
        self.smtp_port = 587

        missing = []
        if not self.email_user:
            missing.append("EMAIL_USER")
        if not self.email_password:
            missing.append("EMAIL_PASSWORD")
        if not self.report_to:
            missing.append("EMAIL_DESTINATARIO")
        if missing:
            raise RuntimeError(
                f"Faltan variables de entorno requeridas para envío de mail: {', '.join(missing)}"
            )

    def send_message(
        self,
        subject: str,
        html_body: str,
        text_body: str,
        attachments: Optional[List[Path]] = None,
    ) -> None:
        mixed_root = MIMEMultipart("mixed")
        mixed_root["Subject"] = subject
        mixed_root["From"] = self.report_from
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
            server.login(self.email_user, self.email_password)
            server.sendmail(
                self.report_from,
                self.report_to + self.report_cc,
                mixed_root.as_string(),
            )

        print(f"Correo enviado: {subject}")


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def _resolve_report_paths(period_slug: str) -> tuple[Path, Path, Path, Path]:
    period_dir = REPORTS_DIR / period_slug
    metadata_path = period_dir / "metadata.json"
    html_path = period_dir / "report.html"
    text_path = period_dir / "report.txt"
    json_path = period_dir / "report.json"

    missing = [
        str(path.name)
        for path in [metadata_path, html_path, text_path, json_path]
        if not path.exists()
    ]
    if missing:
        raise FileNotFoundError(
            f"No existe el reporte completo para {period_slug}. Faltan: {', '.join(missing)}"
        )

    return metadata_path, html_path, text_path, json_path


def send_period_report(period_slug: str) -> None:
    metadata_path, html_path, text_path, json_path = _resolve_report_paths(period_slug)

    metadata = _load_json(metadata_path)
    report_json = _load_json(json_path)
    html_body = html_path.read_text(encoding="utf-8")
    text_body = text_path.read_text(encoding="utf-8")

    subject = (
        os.environ.get("EMAIL_SUBJECT_OVERRIDE")
        or metadata.get("email_subject")
        or report_json.get("email_subject")
        or f"Informe CI | {period_slug}"
    )

    attachments: List[Path] = []
    if (os.environ.get("ATTACH_HTML_REPORT") or "false").lower() == "true":
        attachments.append(html_path)
    if (os.environ.get("ATTACH_JSON_REPORT") or "false").lower() == "true":
        attachments.append(json_path)

    sender = EmailSender()
    sender.send_message(
        subject=subject,
        html_body=html_body,
        text_body=text_body,
        attachments=attachments,
    )

    print(
        json.dumps(
            {
                "status": "sent",
                "period_slug": period_slug,
                "subject": subject,
                "to": sender.report_to,
                "cc": sender.report_cc,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


def main() -> None:
    period_slug = (os.environ.get("TARGET_PERIOD_SLUG") or os.environ.get("REPORT_SLUG") or "").strip()
    if not period_slug:
        raise RuntimeError("Falta TARGET_PERIOD_SLUG o REPORT_SLUG para enviar el reporte.")
    send_period_report(period_slug)


if __name__ == "__main__":
    main()
