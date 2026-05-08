# Reporting de Comunicaciones Internas · BBVA

Pipeline end-to-end para generar un **reporte mensual ejecutivo** a partir de dashboards PDF.

## Arquitectura final (MVP)

1. **Ingesta PDF** desde Gmail (`scripts/fetch_dashboard_pdfs.py`).
2. **Extracción mensual determinística** desde PDF (`scripts/deterministic_pipeline.py` + `scripts/pdf_processor.py`):
   - artefacto `data/raw_extracted/<mes>.json`
   - artefacto `data/canonical_monthly/<mes>.json`
   - artefacto `data/validation/<mes>.json`
3. **KPIs determinísticos y guardrails** (`scripts/analyzer.py`):
   - consolida métricas, mixes y rankings
   - aplica `quality_flags`
   - controla comparabilidad histórica real
4. **Data Quality explícito** (`scripts/data_quality.py` + `scripts/validate_report.py`):
   - valida contrato mensual y contrato de reporte
   - marca rankings push incompletos antes de renderizar
   - corta inconsistencias fuertes y deja warnings auditables
5. **Narrativa ejecutiva (LLM opcional)** (`prompts/period_report.txt`):
   - el LLM redacta copy
   - no decide números
6. **Render de body en JS** (`scripts/pptx_renderer.js`):
   - renderer principal
   - módulos dinámicos + empty states corporativos
7. **Ensamblado final de deck** (`scripts/deck_assembler.py`):
   - portada = primera slide de `assets/plantilla-bbva.pptx`
   - reemplazo de `FECHA` por período real
   - body generado en JS en el medio
   - cierre = última slide de plantilla

## Reglas de presentación

- No se anexa la plantilla completa.
- No se muestran títulos técnicos internos (`slide_*`).
- No se renderizan slides vacías.
- Cada slide prioriza: 1 insight central, 1 visual dominante y bullets cerrados sin puntos suspensivos.
- Los rankings se muestran como cards + gráfico; las distribuciones se muestran como donut o barras horizontales.
- La portada identifica automáticamente si el informe es mensual, trimestral o anual.
- La card de top campaña puede mostrar uplift vs. interacción promedio cuando hay base suficiente.
- Módulo de **eventos** es condicional:
  - si no hay data suficiente, se omite.
- Si no hay comparabilidad histórica, se informa: **“No comparable por alcance de fuente”**.

## Guardrails de calidad de datos

- Si un mail rankea con interacción alta pero trae `0 clics`, el pipeline agrega warning de inconsistencia y lo marca como `data_complete=false`.
- Si un mail trae interacción alta sin apertura, se considera dato incompleto para evitar publicar KPIs inválidos.
- Si una nota pull tiene usuarios únicos mayores que vistas totales, se genera warning de calidad.
- Las áreas solicitantes se extraen como categorías separadas: por ejemplo, `Client Solutions` y `Engineering & Data` no se fusionan.
- Los títulos de mails se normalizan reemplazando guiones bajos, reconstruyendo nombres truncados y evitando títulos con puntos suspensivos.
- Las slides vacías se omiten automáticamente salvo que exista contexto manual.

## Contrato mensual esperado

Cada resumen mensual debe incluir (mínimo):

- `plan_total`
- `site_notes_total`
- `site_total_views`
- `mail_total`
- `mail_open_rate`
- `mail_interaction_rate`
- `strategic_axes[]`
- `internal_clients[]`
- `channel_mix[]`
- `format_mix[]`
- `top_push_by_interaction[]`
- `top_push_by_open_rate[]`
- `top_pull_notes[]`
- `hitos[]`
- `events[]`
- `quality_flags`

`quality_flags` obligatorios:

- `scope_country`
- `scope_mixed`
- `site_has_no_data_sections`
- `events_summary_available`
- `push_ranking_available`
- `pull_ranking_available`
- `historical_comparison_allowed`

## Estructura del deck de salida

1. Portada template
2. Resumen ejecutivo del período
3. Gestión de canales
4. Mix temático y áreas solicitantes
5. Ranking push
6. Ranking pull
7. Hitos del mes
8. Eventos del mes (condicional)
9. Cierre template

## Comandos

### Instalación

```bash
pip install -r requirements.txt
npm install
```

### Tests

```bash
python -m unittest discover -s tests -p 'test*.py'
```

### QA local de muestra

Valida contratos y genera un PPTX smoke-test en `output/smoke/`:

```bash
npm run qa:sample
```

### Validar artefactos manualmente

```bash
python scripts/validate_report.py data/canonical_jan_2026_ejecutivo_v3.json --kind canonical
python scripts/validate_report.py data/report_jan_2026_ejecutivo_v3.json --kind report
```

### Generar reporte puntual (solo procesamiento local)

```bash
python scripts/generate_report.py --period 2026-03 --skip-email-fetch --pdf-dir local_data/inbox_pdfs
```

### Fetch + process para un período

```bash
python scripts/generate_report.py --period 2026-Q1 --fetch-email --pdf-dir local_data/inbox_pdfs
```

### Pipeline completo (fetch + generate + send)

```bash
python scripts/run_scheduled_reports.py
```

### Extracción directa de un PDF (debug raw)

```bash
python scripts/deterministic_pipeline.py --input local_data/inbox_pdfs/2026-01_dashboard.pdf --output local_data/debug/2026-01_raw.json
```

### Render de muestra del renderer JS (modo full demo)

```bash
npm run render:sample
```

## CI

Workflow: `.github/workflows/comms-report.yml`

Orden de ejecución:
1. instala dependencias Python + Node
2. corre tests unitarios
3. valida contratos de muestra
4. genera un deck smoke-test con el renderer JS
5. ejecuta pipeline real
6. sube artefactos

Si los tests fallan, no avanza a generación/envío.
