import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
REPORTS_DIR = DATA_DIR / "reports"


class EmailSender:
    def __init__(self):
        self.email_user = (os.environ.get("EMAIL_USER") or "").strip()
        self.email_password = (os.environ.get("EMAIL_PASSWORD") or "").strip()
        self.email_to = (os.environ.get("EMAIL_DESTINATARIO") or "").strip()

        self.smtp_host = "smtp.gmail.com"
        self.smtp_port = 587

        missing = []
        if not self.email_user:
            missing.append("EMAIL_USER")
        if not self.email_password:
            missing.append("EMAIL_PASSWORD")
        if not self.email_to:
            missing.append("EMAIL_DESTINATARIO")

        if missing:
            raise RuntimeError(
                f"Faltan variables de entorno requeridas para envío de mail: {', '.join(missing)}"
            )

    def send_email(self, subject: str, html_content: str) -> None:
        msg = MIMEMultipart("alternative")
        msg["From"] = self.email_user
        msg["To"] = self.email_to
        msg["Subject"] = subject

        msg.attach(MIMEText(html_content, "html", "utf-8"))

        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            server.starttls()
            server.login(self.email_user, self.email_password)
            server.sendmail(
                self.email_user,
                [addr.strip() for addr in self.email_to.split(",") if addr.strip()],
                msg.as_string(),
            )


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def _resolve_report_paths(period_slug: str) -> tuple[Path, Path]:
    period_dir = REPORTS_DIR / period_slug
    metadata_path = period_dir / "metadata.json"
    html_path = period_dir / "report.html"

    if not metadata_path.exists():
        raise FileNotFoundError(f"No existe metadata.json para el período {period_slug}: {metadata_path}")

    if not html_path.exists():
        raise FileNotFoundError(f"No existe report.html para el período {period_slug}: {html_path}")

    return metadata_path, html_path


def send_period_report(period_slug: str) -> None:
    metadata_path, html_path = _resolve_report_paths(period_slug)

    metadata = _load_json(metadata_path)
    html_content = html_path.read_text(encoding="utf-8")

    subject = (
        metadata.get("email_subject")
        or metadata.get("subject")
        or metadata.get("title")
        or f"Informe CI | {period_slug}"
    )

    sender = EmailSender()
    sender.send_email(subject=subject, html_content=html_content)

    print(
        json.dumps(
            {
                "status": "sent",
                "period_slug": period_slug,
                "subject": subject,
                "to": sender.email_to,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


def main() -> None:
    period_slug = (os.environ.get("REPORT_SLUG") or "").strip()
    if not period_slug:
        raise RuntimeError("Falta REPORT_SLUG para enviar el reporte.")

    send_period_report(period_slug)


if __name__ == "__main__":
    main()