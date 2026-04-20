# Pipeline de informe de Comunicaciones Internas · BBVA

Este repo genera un informe ejecutivo en PowerPoint a partir de los dashboards PDF que llegan por mail.

El flujo completo hace esto:

1. Busca en Gmail los PDFs del dashboard del período.
2. Descarga y ordena los PDFs por mes.
3. Le pide a Gemini una extracción estructurada de KPIs e insights.
4. Consolida el período con reglas determinísticas.
5. Renderiza un deck `.pptx` con estética BBVA usando PptxGenJS.
6. Envía el reporte por mail con el PowerPoint adjunto.

## Qué cambió en esta versión

- Renderer nuevo en `PptxGenJS` con narrativa fija de 9 slides.
- Lenguaje visual BBVA: paleta, jerarquías, cards KPI, charts y cierre ejecutivo.
- Soporte para miniaturas reales en top comunicaciones e hitos.
- Soporte para contexto manual opcional (`data/manual_context/<period_slug>.json`).
- Fallback automático: si falla Gemini, igual genera un deck básico con los KPIs calculados.
- Workflow listo para GitHub Actions con Python + Node.

## Estructura

- `scripts/fetch_dashboard_pdfs.py`: descarga los PDFs desde Gmail.
- `scripts/pdf_processor.py`: genera el resumen mensual estructurado.
- `scripts/analyzer.py`: consolida KPIs, rankings, timelines y distribuciones.
- `scripts/generate_report.py`: arma el JSON final del deck y genera los artefactos.
- `scripts/pptx_renderer.js`: renderer principal del PowerPoint.
- `scripts/pptx_renderer.py`: wrapper Python del renderer.
- `data/manual_context/`: overrides opcionales por período.
- `templates/manual_context.example.json`: ejemplo de override manual.
- `assets/demo/`: miniaturas demo para probar el renderer.
- `assets/brand/`: logos oficiales usados por el renderer (`bbva_logo_blue.png` y `bbva_logo_white.png`).

## Requisitos locales

- Python 3.11+
- Node 20+
- acceso a Gmail API
- `GEMINI_API_KEY`

## Instalación local

```bash
python -m venv .venv
source .venv/bin/activate  # en Windows: .venv\Scripts\activate
pip install -r requirements.txt
npm install
```

## Variables de entorno

### Gmail API
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN`
- `GOOGLE_TOKEN_URI` (opcional, default Google OAuth)
- `GMAIL_EXPECTED_SENDER` (opcional, para filtrar remitente esperado)
- `GMAIL_EXPECTED_KEYWORDS` (opcional, lista separada por coma para validar subject/filename)

### Gemini
- `GEMINI_API_KEY`
- `GEMINI_MODEL` (opcional, default `gemini-2.5-flash`)
- `GEMINI_FALLBACK_MODELS` (opcional, separados por coma)
- `GEMINI_ENABLE_MODEL_FALLBACK` (opcional, default `true`)
- `GEMINI_MAX_RETRIES_PER_MODEL` (opcional, default `3`)
- `GEMINI_INITIAL_BACKOFF_SECONDS` (opcional, default `3`)
- `GEMINI_MAX_BACKOFF_SECONDS` (opcional, default `30`)
- `GEMINI_UPLOAD_PROCESS_TIMEOUT_SECONDS` (opcional, default `300`)
- `GEMINI_UPLOAD_RETRIES` (opcional, default `3`)

### Mail de salida
- `EMAIL_USER`
- `EMAIL_PASSWORD`
- `EMAIL_DESTINATARIO`
- `EMAIL_CC` (opcional)
- `EMAIL_BCC` (opcional)
- `EMAIL_FROM` (opcional)
- `SMTP_HOST` (opcional, default `smtp.gmail.com`)
- `SMTP_PORT` (opcional, default `587`)

### Reporte
- `REPORT_TIMEZONE` (default `America/Argentina/Buenos_Aires`)
- `REPORT_MODE` (`auto`, `month`, `month_and_quarter`, `quarter`, `year`, `quarter_and_year`)
- `REPORT_YEAR`
- `REPORT_MONTH`
- `REPORT_QUARTER`
- `ALLOW_PARTIAL_PERIOD` (`true`/`false`)
- `REPORT_SLUG` (para generación/envío puntual)

## Uso local

### 1) Descargar PDFs del período

```bash
python scripts/fetch_dashboard_pdfs.py
```

### 2) Generar deck para un período puntual

```bash
REPORT_SLUG=month_2026_03 python scripts/generate_report.py
```

### 3) Enviar el reporte generado

```bash
REPORT_SLUG=month_2026_03 python scripts/send_email.py
```

### 4) Probar el renderer con la demo incluida

```bash
node scripts/pptx_renderer.js templates/sample_report_definitive.json sample_bbva_report_definitive.pptx
```

## Contexto manual opcional

Para completar hitos, eventos o miniaturas reales, podés crear:

`data/manual_context/<period_slug>.json`

Ejemplo en `templates/manual_context.example.json`.

Se usa para:
- `slide_5_push_ranking.top_communications[].thumbnail_path`
- `slide_7_hitos`
- `slide_8_events`
- `slide_9_closure`
- metadata adicional del envío

## Artefactos generados

En `output/reports/<period_slug>/` se guardan:

- `metadata.json`
- `report_raw.json`
- `report.html`
- `report.pptx`

Los summaries mensuales se cachean en `data/monthly_summaries/` para reutilización consistente entre corridas.

Si Gemini falla y no existe cache previa del mes, el pipeline genera automáticamente un summary mínimo en modo `local_fallback` para no frenar la generación del reporte.

## Nota operativa

El dashboard es la fuente principal del reporte, pero los módulos de `hitos`, `eventos` y miniaturas pueden requerir contexto manual adicional para quedar al nivel del deck editorial del equipo.


## Presentación

El renderer incorpora ajuste automático de texto y límites por caja para evitar superposición entre textos, gráficos y paneles cuando cambian los datos del período.
