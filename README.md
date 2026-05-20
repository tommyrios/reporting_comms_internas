# Reporting de Comunicaciones Internas · BBVA

Pipeline Python para generar y enviar reportes **trimestrales o anuales** de Comunicaciones Internas a partir de dashboards PDF exportados desde Looker Studio.

## Flujo productivo

1. **Resolución de período** (`scripts/reporting_periods.py`)
   - Modo automático: genera trimestre cerrado; en enero genera Q4 + anual del año anterior.
   - Modo manual: `REPORT_MODE=quarter`, `REPORT_MODE=year` o `REPORT_MODE=quarter_and_year`.
2. **Ingesta desde Gmail** (`scripts/fetch_dashboard_pdfs.py`)
   - Descarga los PDFs requeridos por scope: `argentina`, `holding`, `combined`.
   - Persiste `data/reporting_periods.json`, `data/fetch_result.json` y `local_data/inbox_pdfs/manifest.json`.
3. **Extracción determinística de PDF** (`scripts/deterministic_pipeline.py` + `scripts/period_pdf_processor.py`)
   - Genera artefactos por período/scope en `data/raw_extracted/`, `data/canonical_periods/`, `data/period_summaries/` y `data/validation/`.
4. **KPIs y guardrails** (`scripts/analyzer.py`, `scripts/data_quality.py`)
   - Consolida métricas, rankings, mixes y banderas de calidad.
   - Evita comparabilidad histórica cuando el alcance de fuente no es equivalente.
5. **Crops del dashboard** (`scripts/dashboard_crops.py`)
   - Recorta visuales relevantes del PDF para insertarlos en la presentación.
6. **Render PPTX Python** (`scripts/pptx_renderer.py`)
   - Genera `output/reports/<period_slug>/report.pptx`.
7. **Envío por email** (`scripts/send_email.py`)
   - Envía el PPTX y deja metadata auditable.

## Reglas de negocio vigentes

- No se generan reportes mensuales. El dashboard de entrada ya debe venir filtrado por trimestre o año.
- El reporte corporativo es 100% determinístico: no usa GenAI para narrativa ni para tomar decisiones de datos.
- Los scopes productivos esperados son `argentina`, `holding` y `combined`.
- El consolidado ejecutivo sale del scope `combined`; no se suma manualmente Argentina + Holding.
- Si no hay comparabilidad histórica, se informa como no comparable por alcance de fuente.
- Las slides vacías se omiten salvo que exista contexto manual.
- Los títulos se preservan en el contrato canónico hasta 180 caracteres; el renderer decide el recorte visual por layout.

## Variables principales

```bash
REPORT_MODE=auto                  # auto | quarter | year | quarter_and_year
REPORT_YEAR=2026                  # requerido para quarter/year manual
REPORT_QUARTER=1                  # requerido para quarter manual
REPORT_TIMEZONE=America/Argentina/Buenos_Aires
REPORT_REQUIRED_SCOPES=argentina,holding,combined
ALLOW_PARTIAL_PERIOD=false
ALLOW_PARTIAL_REPORT=false
```

## Instalación

```bash
pip install -r requirements.txt
```

## Tests

```bash
python -m unittest discover -s tests -p 'test*.py'
```

## QA local

```bash
python scripts/validate_report.py data/canonical_periods/quarter_2026_Q1_combined.json --kind canonical
python scripts/pptx_renderer.py data/report_boceto_ci_sample.json output/smoke/report_boceto_ci_sample.pptx
```

## Generar reporte puntual con PDFs locales

Los PDFs deben estar nombrados por período y scope, por ejemplo:

```text
local_data/inbox_pdfs/quarter_2026_Q1_ARG.pdf
local_data/inbox_pdfs/quarter_2026_Q1_HOLDING.pdf
local_data/inbox_pdfs/quarter_2026_Q1_ARG_HOLDING.pdf
```

Comando:

```bash
REPORT_MODE=quarter REPORT_YEAR=2026 REPORT_QUARTER=1 \
python scripts/reporting_periods.py

python scripts/generate_report.py --period quarter_2026_Q1 --skip-email-fetch --pdf-dir local_data/inbox_pdfs
```

## Pipeline completo

```bash
python scripts/run_scheduled_reports.py
```

Ese comando ejecuta:

```text
Gmail -> PDFs por scope -> extracción determinística -> KPIs -> crops -> PPTX -> email
```

## Contexto manual opcional

Para agregar overrides controlados, crear:

```text
data/manual_context/<period_slug>.json
```

Usar `templates/manual_context.example.json` como referencia.

## CI

Workflow: `.github/workflows/comms-report.yml`

Orden de ejecución:

1. instala dependencias Python
2. corre tests unitarios
3. valida muestra canónica trimestral
4. genera PPTX smoke-test con renderer Python
5. ejecuta pipeline real
6. sube artefactos
