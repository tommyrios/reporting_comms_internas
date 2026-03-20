# Pipeline de informe de Comunicaciones Internas

Este repo implementa un flujo con GitHub Actions para:

1. Buscar el ultimo mail cuyo asunto contenga `Dashboard Communications | Comunicación interna`.
2. Descargar el adjunto PDF cuyo nombre contenga `Dashboard Communications | Comunicación interna`.
3. Extraer el texto del PDF.
4. Pedir a OpenAI un resumen ejecutivo con insights y alertas.
5. Enviar el informe por mail.

## Variables de entorno requeridas

Configuralas como **GitHub Secrets**:

- `IMAP_HOST`
- `IMAP_PORT` (normalmente `993`)
- `EMAIL_USERNAME`
- `EMAIL_PASSWORD`
- `OPENAI_API_KEY`
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`
- `REPORT_TO`
- `REPORT_CC` (opcional)
- `REPORT_FROM` (opcional; si no, usa `SMTP_USERNAME`)
- `OPENAI_MODEL` (opcional; default `gpt-5-mini`)

## Como funciona

El workflow corre por cron y tambien manualmente.

- `fetch_latest_dashboard_email.py`: busca el ultimo mail y descarga el PDF.
- `extract_pdf_text.py`: convierte el PDF a texto plano y genera metadatos.
- `generate_report.py`: usa IA para redactar el informe ejecutivo.
- `send_email.py`: envia el informe por correo en HTML.

## Recomendacion operativa

Como GitHub Actions no se dispara “por recepcion de mail” de forma nativa, la forma practica es:

- programar este workflow para correr cada hora o una vez por dia,
- tomar siempre el ultimo mail coincidente,
- y ejecutar el envio solo cuando haya un PDF nuevo.

Para evitar reenvios duplicados, el script genera una huella (`message_id` + `attachment_name`) y la compara con un archivo de estado si decidis persistirlo en el repo o en un bucket. En esta version base el flujo procesa siempre el ultimo PDF encontrado.
