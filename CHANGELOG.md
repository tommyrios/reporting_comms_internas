# Changelog

## v1.2.0 - flujo productivo trimestral/anual

- Deja el repo en modo Python-only para producción.
- Elimina renderer JS, deck assembler, prompts/GenAI, processor mensual legacy y dependencias Node.
- Elimina muestras mensuales y conserva muestras trimestrales/anuales por período/scope.
- Consolida parsers numéricos en `metric_utils.py`.
- Quita fallback `selected_periods.json`; el cronograma canónico es `data/reporting_periods.json`.
- Actualiza CI para validar muestra canónica y renderizar PPTX con `scripts/pptx_renderer.py`.

## v1.1.1 - hotfix tests y smoke render

- Corrige validación canónica para no invalidar payloads parciales por campos estructurales faltantes.
- Mejora smoke render del renderer Python.

## v1.1.0 - ejecutivo v7

- Activa tests en GitHub Actions.
- Agrega validación explícita de contratos con `scripts/data_quality.py` y `scripts/validate_report.py`.
- Marca rankings push incompletos con `data_complete=false` y `data_quality_issue`.
- Integra validaciones de calidad en `deterministic_pipeline.py` y `analyzer.py`.
- Mejora renderer PPTX con subtítulo dinámico, uplift vs. promedio y empty states.
