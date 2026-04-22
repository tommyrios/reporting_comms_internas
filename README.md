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
4. **Narrativa ejecutiva (LLM opcional)** (`prompts/period_report.txt`):
   - el LLM redacta copy
   - no decide números
5. **Render de body en JS** (`scripts/pptx_renderer.js`):
   - renderer principal
   - módulos dinámicos + empty states corporativos
6. **Ensamblado final de deck** (`scripts/deck_assembler.py`):
   - portada = primera slide de `assets/plantilla-bbva.pptx`
   - reemplazo de `FECHA` por período real
   - body generado en JS en el medio
   - cierre = última slide de plantilla

## Reglas de presentación

- No se anexa la plantilla completa.
- No se muestran títulos técnicos internos (`slide_*`).
- No se renderizan slides vacías.
- Módulo de **eventos** es condicional:
  - si no hay data suficiente, se omite.
- Si no hay comparabilidad histórica, se informa: **“No comparable por alcance de fuente”**.

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

### Generar reporte puntual

```bash
REPORT_SLUG=month_2026_03 python scripts/generate_report.py
```

### Pipeline completo (fetch + generate + send)

```bash
python scripts/run_scheduled_reports.py
```

### Render de muestra del renderer JS (modo full demo)

```bash
npm run render:sample
```

## CI

Workflow: `.github/workflows/comms-report.yml`

Orden de ejecución:
1. instala dependencias Python + Node
2. corre tests
3. ejecuta pipeline
4. sube artefactos

Si los tests fallan, no avanza a generación/envío.
