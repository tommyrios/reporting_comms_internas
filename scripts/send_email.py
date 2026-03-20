import os
import smtplib
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

OUTPUT_HTML = Path("output/report.html")
OUTPUT_TEXT = Path("output/report.txt")


class MensajeSender:
    def __init__(self):
        self.email_user = os.environ.get("EMAIL_USER")
        self.email_pass = os.environ.get("EMAIL_PASSWORD")

        destinatarios_str = os.environ.get("EMAIL_DESTINATARIO", "")
        self.destinatarios = [d.strip() for d in destinatarios_str.split(",") if d.strip()]

        cc_str = os.environ.get("EMAIL_CC", "")
        self.cc = [d.strip() for d in cc_str.split(",") if d.strip()]

    def enviar_difusion(self, contenido_html, asunto="Reporte | Dashboard Communications"):
        if not self.email_user or not self.email_pass:
            raise ValueError("Faltan EMAIL_USER o EMAIL_PASSWORD.")

        if not self.destinatarios:
            raise ValueError("No hay destinatarios definidos en EMAIL_DESTINATARIO.")

        msg = MIMEMultipart("alternative")
        msg["Subject"] = asunto
        msg["From"] = self.email_user
        msg["To"] = ", ".join(self.destinatarios)

        if self.cc:
            msg["Cc"] = ", ".join(self.cc)

        texto_plano = (
            OUTPUT_TEXT.read_text(encoding="utf-8")
            if OUTPUT_TEXT.exists()
            else "Por favor, habilite la visualización HTML para ver este reporte."
        )

        msg.attach(MIMEText(texto_plano, "plain", "utf-8"))
        msg.attach(MIMEText(contenido_html, "html", "utf-8"))

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(self.email_user, self.email_pass)
            server.sendmail(
                self.email_user,
                self.destinatarios + self.cc,
                msg.as_string()
            )

        print(f"Correo enviado exitosamente a {len(self.destinatarios)} destinatarios.")


def main():
    if not OUTPUT_HTML.exists():
        raise FileNotFoundError("No existe output/report.html")

    contenido_html = OUTPUT_HTML.read_text(encoding="utf-8")
    asunto = os.environ.get("EMAIL_SUBJECT", "Reporte automático | Dashboard Communications")

    sender = MensajeSender()
    sender.enviar_difusion(contenido_html, asunto=asunto)


if __name__ == "__main__":
    main()