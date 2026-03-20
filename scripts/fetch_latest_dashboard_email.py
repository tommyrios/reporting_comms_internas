import email
import imaplib
import json
import os
from email.header import decode_header
from email.message import Message
from pathlib import Path

DATA_DIR = Path('data')
DATA_DIR.mkdir(exist_ok=True)

SUBJECT_NEEDLE = 'Dashboard Communications | Comunicación interna'
ATTACHMENT_NEEDLE = 'Dashboard Communications | Comunicación interna'


def decode_mime(value: str | None) -> str:
    if not value:
        return ''
    parts = decode_header(value)
    out = []
    for part, enc in parts:
        if isinstance(part, bytes):
            out.append(part.decode(enc or 'utf-8', errors='replace'))
        else:
            out.append(part)
    return ''.join(out)


def attachment_matches(part: Message) -> tuple[bool, str]:
    filename = decode_mime(part.get_filename())
    if not filename:
        return False, ''
    lowered = filename.lower()
    return ATTACHMENT_NEEDLE.lower() in lowered and lowered.endswith('.pdf'), filename


def main() -> None:
    host = os.environ['IMAP_HOST']
    port = int(os.environ.get('IMAP_PORT', '993'))
    username = os.environ['EMAIL_USERNAME']
    password = os.environ['EMAIL_PASSWORD']

    mail = imaplib.IMAP4_SSL(host, port)
    mail.login(username, password)
    mail.select('INBOX')

    status, data = mail.search(None, 'ALL')
    if status != 'OK':
        raise RuntimeError('No se pudo buscar mails')

    ids = data[0].split()
    latest_match = None

    for msg_id in reversed(ids):
        status, msg_data = mail.fetch(msg_id, '(RFC822)')
        if status != 'OK':
            continue
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        subject = decode_mime(msg.get('Subject', ''))
        if SUBJECT_NEEDLE.lower() not in subject.lower():
            continue

        for part in msg.walk():
            if part.get_content_disposition() != 'attachment':
                continue
            matches, filename = attachment_matches(part)
            if not matches:
                continue
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            pdf_path = DATA_DIR / 'latest_dashboard.pdf'
            pdf_path.write_bytes(payload)
            latest_match = {
                'message_id': msg.get('Message-ID', ''),
                'subject': subject,
                'date': msg.get('Date', ''),
                'from': msg.get('From', ''),
                'attachment_name': filename,
                'pdf_path': str(pdf_path),
            }
            break
        if latest_match:
            break

    mail.logout()

    if not latest_match:
        raise FileNotFoundError('No encontre un mail con el asunto y PDF esperados')

    (DATA_DIR / 'metadata.json').write_text(
        json.dumps(latest_match, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )
    print(json.dumps(latest_match, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
