import html
import json
import mimetypes
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
REPORTS_DIR = OUTPUT_DIR / "reports"


class EmailSender:
    def __init__(self):
        self.email_user = (os.environ.get("EMAIL_USER") or "").strip()
        self.email_password = (os.environ.get("EMAIL_PASSWORD") or "").strip()
        self.email_to = (os.environ.get("EMAIL_DESTINATARIO") or "").strip()
        self.email_cc = (os.environ.get("EMAIL_CC") or "").strip()
        self.email_bcc = (os.environ.get("EMAIL_BCC") or "").strip()
        self.email_from = (os.environ.get("EMAIL_FROM") or self.email_user).strip()
        self.smtp_host = (os.environ.get("SMTP_HOST") or "smtp.gmail.com").strip()
        self.smtp_port = int((os.environ.get("SMTP_PORT") or "587").strip())

        missing = []
        if not self.email_user:
            missing.append("EMAIL_USER")
        if not self.email_password:
            missing.append("EMAIL_PASSWORD")
        if not self.email_to:
            missing.append("EMAIL_DESTINATARIO")
        if missing:
            raise RuntimeError(f"Faltan variables de entorno requeridas para envío de mail: {', '.join(missing)}")

    def send_email(self, subject: str, html_content: str, plain_text: str, attachments: list[Path] | None = None) -> None:
        msg = MIMEMultipart("mixed")
        msg["From"] = self.email_from
        msg["To"] = self.email_to
        msg["Subject"] = subject
        if self.email_cc:
            msg["Cc"] = self.email_cc

        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain_text, "plain", "utf-8"))
        alt.attach(MIMEText(html_content, "html", "utf-8"))
        msg.attach(alt)

        for attachment_path in attachments or []:
            if not attachment_path.exists():
                continue
            content_type, _ = mimetypes.guess_type(str(attachment_path))
            subtype = "octet-stream"
            if content_type and "/" in content_type:
                _, subtype = content_type.split("/", 1)
            with open(attachment_path, "rb") as file_obj:
                part = MIMEApplication(file_obj.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=attachment_path.name)
            msg.attach(part)

        recipients = [addr.strip() for addr in self.email_to.split(",") if addr.strip()]
        recipients += [addr.strip() for addr in self.email_cc.split(",") if addr.strip()]
        recipients += [addr.strip() for addr in self.email_bcc.split(",") if addr.strip()]

        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            server.starttls()
            server.login(self.email_user, self.email_password)
            server.sendmail(self.email_from, recipients, msg.as_string())


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def _resolve_report_paths(period_slug: str) -> tuple[Path, Path, Path | None]:
    period_dir = REPORTS_DIR / period_slug
    metadata_path = period_dir / "metadata.json"
    html_path = period_dir / "report.html"
    pptx_path = period_dir / "report.pptx"
    if not metadata_path.exists():
        raise FileNotFoundError(f"No existe metadata.json para el período {period_slug}: {metadata_path}")
    if not html_path.exists():
        raise FileNotFoundError(f"No existe report.html para el período {period_slug}: {html_path}")
    return metadata_path, html_path, (pptx_path if pptx_path.exists() else None)


def _build_email_bodies(metadata: dict, attachment_label: str) -> tuple[str, str]:
    period_label = metadata.get("period") or metadata.get("period_label") or metadata.get("subtitle") or "el período"
    period_label = str(period_label).strip() or "el período"
    period_label_html = html.escape(period_label)

    plain_text = f"""Hola equipo!

Les comparto el informe de gestión de {period_label} 🚀:
Una mirada rápida sobre qué funcionó mejor, qué contenidos destacaron y qué oportunidades aparecen para seguir mejorando.

Ojalá les sirva para leer resultados y pensar próximos pasos.

Saludos!!""".strip()

    html_body = f"""
    <html>
      <body style="font-family:Arial,Helvetica,sans-serif;color:#111827;line-height:1.45;">
        <p>Hola equipo!</p>

        <p>Les comparto el informe de gestión de <strong>{period_label_html}</strong> 🚀:</p>
        <p>Una mirada rápida sobre qué funcionó mejor, qué contenidos destacaron y qué oportunidades aparecen para seguir mejorando.</p>

        <p>Ojalá les sirva para leer resultados y pensar próximos pasos.</p>

        <p>Saludos!!</p>
      </body>
    </html>
    """.strip()
    return plain_text, html_body


def send_period_report(period_slug: str) -> None:
    metadata_path, html_path, pptx_path = _resolve_report_paths(period_slug)
    metadata = _load_json(metadata_path)
    subject = metadata.get("email_subject") or metadata.get("title") or f"Informe CI | {period_slug}"

    attachments: list[Path] = []
    attachment_label = "PPTX"
    if not pptx_path or not pptx_path.exists():
        raise FileNotFoundError(f"No existe report.pptx para el período {period_slug}")

    attachments.append(pptx_path)

    plain_text, html_body = _build_email_bodies(metadata, attachment_label)
    sender = EmailSender()
    sender.send_email(subject=subject, html_content=html_body, plain_text=plain_text, attachments=attachments)

    print(json.dumps({
        "status": "sent",
        "period_slug": period_slug,
        "subject": subject,
        "to": sender.email_to,
        "attachments": [str(p) for p in attachments],
    }, ensure_ascii=False, indent=2))


def main() -> None:
    period_slug = (os.environ.get("REPORT_SLUG") or "").strip()
    if not period_slug:
        raise RuntimeError("Falta REPORT_SLUG para enviar el reporte.")
    send_period_report(period_slug)


if __name__ == "__main__":
    main()
