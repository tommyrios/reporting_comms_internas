# Changelog

## v1.1.0 - ejecutivo v7

- Activa el paso de tests en GitHub Actions.
- Agrega validacion explicita de contratos con scripts/data_quality.py y scripts/validate_report.py.
- Marca rankings push incompletos con data_complete=false y data_quality_issue.
- Integra validaciones de calidad en deterministic_pipeline.py y analyzer.py.
- Mejora renderer PPTX: subtitulo dinamico, uplift vs. promedio y empty states.
- Agrega scripts NPM de QA local: validate:sample, render:ejecutivo, qa:sample.
- Limpia caches y artefactos temporales del paquete final.
