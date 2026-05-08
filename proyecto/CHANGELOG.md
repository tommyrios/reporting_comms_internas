# Changelog

## v1.1.1 - hotfix tests y smoke render

- Corrige validate_canonical_monthly para no invalidar payloads parciales por campos estructurales faltantes.
- Restaura el modo frame por defecto en el wrapper Python del renderer para conservar portada y cierre de plantilla.
- Hace que pptx_renderer.js cree el directorio de salida antes de escribir el PPTX.

## v1.1.0 - ejecutivo v7

- Activa el paso de tests en GitHub Actions.
- Agrega validacion explicita de contratos con scripts/data_quality.py y scripts/validate_report.py.
- Marca rankings push incompletos con data_complete=false y data_quality_issue.
- Integra validaciones de calidad en deterministic_pipeline.py y analyzer.py.
- Mejora renderer PPTX: subtitulo dinamico, uplift vs. promedio y empty states.
- Agrega scripts NPM de QA local: validate:sample, render:ejecutivo, qa:sample.
- Limpia caches y artefactos temporales del paquete final.
