import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

DATA_DIR = Path('data')
REPORT_JSON = DATA_DIR / 'report.json'
REPORT_HTML = DATA_DIR / 'report.html'


def main() -> None:
    smtp_host = os.environ['SMTP_HOST']
    smtp_port = int(os.environ.get('SMTP_PORT', '587'))
    smtp_user = os.environ['SMTP_USERNAME']
    smtp_password = os.environ['SMTP_PASSWORD']
    report_to = os.environ['REPORT_TO']
    report_cc = os.environ.get('REPORT_CC', '')
    report_from = os.environ.get('REPORT_FROM') or smtp_user

    payload = json.loads(REPORT_JSON.read_text(encoding='utf-8'))
    html = REPORT_HTML.read_text(encoding='utf-8')
    subject = payload.get('subject') or 'Informe automatico | Dashboard Communications'

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = report_from
    msg['To'] = report_to
    if report_cc:
        msg['Cc'] = report_cc

    msg.attach(MIMEText(html, 'html', 'utf-8'))

    recipients = [x.strip() for x in (report_to + ',' + report_cc).split(',') if x.strip()]

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(report_from, recipients, msg.as_string())

    print(f'Reporte enviado a: {", ".join(recipients)}')


if __name__ == '__main__':
    main()
