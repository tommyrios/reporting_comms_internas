import html
import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from google import genai


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
PDF_DIR = DATA_DIR / "monthly_pdfs"
REPORTS_DIR = OUTPUT_DIR / "reports"


MONTH_NAMES_SHORT = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr",
    5: "may", 6: "jun", 7: "jul", 8: "ago",
    9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


MONTHLY_SUMMARY_PROMPT = """
Sos analista senior de comunicaciones internas en BBVA Argentina.

Contexto:
- BBVA prioriza claridad ejecutiva, foco en negocio y lectura rápida.
- Estás analizando VISUALMENTE el Dashboard de métricas adjunto en formato PDF.
- Presta especial atención a gráficos de tendencias, variaciones marcadas en colores (ej. verde/rojo para subidas/bajadas), y distribución en gráficos circulares.
- Las conclusiones deben ser accionables y comparables.
- Evitar texto largo, pensar en slides.

Objetivo:
- Sintetizar el mes con mirada gerencial interpretando los datos del dashboard.
- Identificar qué funcionó, qué no y por qué.
- Conectar métricas con comportamiento.

Formato JSON estricto requerido:
{
  "period_label": "Mar 2026",
  "executive_summary": "2 frases máximo con lectura de negocio basada en tendencias.",
  "headline_metrics": [
    {"label": "Apertura", "value": "75%", "insight": "Alta vs benchmark interno"},
    {"label": "CTR", "value": "8.9%", "insight": "Disparidad entre piezas detectada en gráficos"},
    {"label": "Envíos", "value": "94k", "insight": "Volumen alto"}
  ],
  "highlights": [
    "Insight claro basado en datos y gráficos visuales",
    "Patrón de comportamiento",
    "Contenido que performó mejor"
  ],
  "opportunities": [
    "Mejora concreta accionable",
    "Optimización de canal o contenido"
  ],
  "content_focus": [
    {"theme": "Beneficios", "detail": "Mayor engagement según barras"},
    {"theme": "Institucional", "detail": "Menor interacción"}
  ]
}

Reglas:
- estilo BBVA: corto, directo, accionable.
- no más de 12-15 palabras por bullet.
- priorizar insights analíticos, no descripción.
- No inventar métricas: extrae los números directamente de las tablas y gráficos del PDF.
""".strip()


PERIOD_REPORT_PROMPT = """
Sos analista senior de comunicaciones internas en BBVA Argentina.

Contexto:
- Reporte tipo comité / dirección
- Debe leerse en 2 minutos
- Foco en decisiones

Objetivo:
- consolidar insights del período
- mostrar evolución y patrones
- marcar qué escalar y qué corregir

Formato JSON estricto requerido:
{
  "title": "Informe Comunicaciones Internas BBVA",
  "subtitle": "Período ...",
  "executive_summary": "2-3 frases con lectura global del período.",
  "headline_metrics": [
    {"label": "Apertura promedio", "value": "xx%", "insight": "Evolución vs período anterior"},
    {"label": "CTR promedio", "value": "xx%", "insight": "Nivel de engagement"},
    {"label": "Volumen envíos", "value": "xx", "insight": "Impacto en saturación"}
  ],
  "key_wins": [
    "Qué funcionó y por qué",
    "Contenido con mayor impacto",
    "Mejora respecto período anterior"
  ],
  "watchouts": [
    "Riesgo o caída detectada",
    "Tema con bajo rendimiento"
  ],
  "monthly_snapshots": [
    {"month": "Mes", "summary": "1 línea insight"}
  ],
  "next_steps": [
    "Acción concreta 1",
    "Acción concreta 2",
    "Acción concreta 3"
  ]
}

Reglas:
- estilo ejecutivo BBVA
- bullets cortos
- foco en decisión, no descripción
""".strip()


def _safe_load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def _env_bool(name: str, default: bool = False) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return default
    return raw in {"1", "true", "yes", "y", "si", "sí"}


def _env_int(name: str, default: int) -> int:
    raw = (os.environ.get(name) or "").strip()
    if not raw:
        return default
    return int(raw)


def _get_period_definition(period_slug: str) -> dict[str, Any]:
    path = DATA_DIR / "selected_periods.json"
    if not path.exists():
        raise FileNotFoundError(f"No existe {path}")
    payload = _safe_load_json(path)
    for period in payload.get("periods", []):
        if period.get("slug") == period_slug:
            return period
    raise KeyError(f"No se encontró el período {period_slug} en selected_periods.json")


def _clean_json_response(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```json\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^```\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def _build_genai_client() -> genai.Client:
    api_key = (os.environ.get("GEMINI_API_KEY") or "").strip()
    if not api_key:
        raise RuntimeError("Falta GEMINI_API_KEY")
    return genai.Client(api_key=api_key)


def _candidate_models() -> list[str]:
    primary = (os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash").strip()
    fallback_raw = (os.environ.get("GEMINI_FALLBACK_MODELS") or "gemini-2.5-flash-lite,gemini-2.0-flash").strip()
    fallbacks = [m.strip() for m in fallback_raw.split(",") if m.strip()]
    seen = []
    for model in [primary, *fallbacks]:
        if model and model not in seen:
            seen.append(model)
    return seen


def _call_gemini_for_json(client: genai.Client, contents: list) -> dict[str, Any]:
    max_retries = _env_int("GEMINI_MAX_RETRIES", 4)
    initial_backoff = float((os.environ.get("GEMINI_INITIAL_BACKOFF_SECONDS") or "3").strip())
    models = _candidate_models()

    last_error: Exception | None = None

    for model_name in models:
        for attempt in range(1, max_retries + 1):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                )
                text = getattr(response, "text", "") or ""
                parsed = json.loads(_clean_json_response(text))
                return parsed
            except Exception as e:
                last_error = e
                error_text = str(e)
                is_503 = "503" in error_text or "UNAVAILABLE" in error_text.upper()
                if is_503 and attempt < max_retries:
                    wait_seconds = initial_backoff * (2 ** (attempt - 1))
                    print(f"Gemini saturado con {model_name}. Reintento {attempt}/{max_retries} en {wait_seconds:.1f}s")
                    time.sleep(wait_seconds)
                    continue
                print(f"Fallo Gemini con modelo {model_name} en intento {attempt}: {e}")
                break

    if last_error:
        raise last_error
    raise RuntimeError("No se pudo obtener respuesta de Gemini")


def _summaries_dir() -> Path:
    return _ensure_dir(OUTPUT_DIR / "monthly_summaries")


def _report_dir(period_slug: str) -> Path:
    return _ensure_dir(REPORTS_DIR / period_slug)


def summarize_month(client: genai.Client, month_key: str, force_regenerate: bool = False) -> dict[str, Any]:
    summaries_dir = _summaries_dir()
    summary_path = summaries_dir / f"{month_key}.json"

    if summary_path.exists() and not force_regenerate:
        return _safe_load_json(summary_path)

    pdf_path = PDF_DIR / f"{month_key}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"No existe el archivo PDF para {month_key}: {pdf_path}")

    uploaded_file = None
    try:
        print(f"Subiendo PDF multimodal a Gemini: {month_key}")
        uploaded_file = client.files.upload(
            file=str(pdf_path),
            config={'display_name': f"Dashboard_CI_{month_key}"}
        )

        prompt_text = f"{MONTHLY_SUMMARY_PROMPT}\n\nPor favor, analiza el dashboard adjunto correspondiente al período {month_key}."
        contents = [uploaded_file, prompt_text]

        summary = _call_gemini_for_json(client, contents)
        summary["month"] = month_key

        summary_path.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"Resumen IA generado exitosamente: {month_key}")
        return summary
    finally:
        if uploaded_file:
            try:
                client.files.delete(name=uploaded_file.name)
                print(f"Archivo temporal {uploaded_file.name} eliminado de Gemini.")
            except Exception as e:
                print(f"Aviso: No se pudo eliminar el archivo temporal de Gemini: {e}")


def _fallback_month_summary(month_key: str) -> dict[str, Any]:
    label = month_key
    try:
        year, month = month_key.split("-")
        label = f"{MONTH_NAMES_SHORT[int(month)].capitalize()} {year}"
    except Exception:
        pass

    return {
        "month": month_key,
        "period_label": label,
        "executive_summary": "Resumen generado en modo contingencia por indisponibilidad temporal del modelo.",
        "headline_metrics": [
            {"label": "Estado del archivo", "value": "OK", "insight": "El PDF fue descargado correctamente."},
            {"label": "Fuente", "value": "Dashboard PDF", "insight": "Se recomienda revisión manual visual."},
            {"label": "Modo", "value": "Fallback", "insight": "No se pudo realizar el análisis multimodal."},
        ],
        "highlights": [
            "Se logró descargar el dashboard del período.",
            "Conviene reintentar la generación IA para mejorar el insight ejecutivo."
        ],
        "opportunities": [
            "Reprocesar el período cuando Gemini normalice disponibilidad.",
            "Validar manualmente los principales indicadores visuales del dashboard."
        ],
        "content_focus": [
            {"theme": "Dashboard", "detail": "Análisis IA pendiente."}
        ],
    }


def _fallback_period_report(period: dict[str, Any], monthly_summaries: list[dict[str, Any]], error_message: str | None = None) -> dict[str, Any]:
    snapshots = []
    for summary in monthly_summaries:
        month = summary.get("period_label") or summary.get("month") or "Mes"
        executive = summary.get("executive_summary") or "Resumen no disponible."
        snapshots.append({"month": month, "summary": executive})

    warning = "Reporte generado en modo contingencia por falla del modelo."
    if error_message:
        warning = f"{warning} Error original: {error_message}"

    return {
        "title": period.get("title") or f"Informe CI | {period.get('slug')}",
        "subtitle": period.get("subtitle") or period.get("label") or period.get("slug"),
        "email_subject": period.get("email_subject") or f"Informe CI | {period.get('slug')}",
        "executive_summary": "Se generó una versión mínima del reporte para no interrumpir el pipeline de envío.",
        "headline_metrics": [
            {"label": "Período", "value": period.get("label", "-"), "insight": "Período procesado correctamente."},
            {"label": "Meses incluidos", "value": str(len(period.get('months', []))), "insight": "Cantidad de PDFs consolidados."},
            {"label": "Modo", "value": "Fallback", "insight": "Conviene regenerar con IA visual."},
        ],
        "key_wins": [
            "Se descargaron los PDFs del período.",
            "El envío no se interrumpe ante una caída de la API."
        ],
        "watchouts": [
            "El contenido ejecutivo carece del análisis visual del dashboard.",
        ],
        "monthly_snapshots": snapshots,
        "next_steps": [
            "Reintentar generación IA más tarde.",
            "Validar los gráficos del dashboard fuente manualmente.",
        ],
        "warning": warning,
    }


def _escape(text: Any) -> str:
    return html.escape(str(text or ""))


def _html_list(items: list[str]) -> str:
    rows = "".join(f"<li>{_escape(item)}</li>" for item in items if str(item).strip())
    return f"<ul>{rows}</ul>" if rows else "<p class='muted'>Sin datos.</p>"


def _render_report_html(report: dict[str, Any]) -> str:
    title = _escape(report.get("title"))
    subtitle = _escape(report.get("subtitle"))
    executive_summary = _escape(report.get("executive_summary"))
    warning = report.get("warning")

    headline_cards = ""
    for item in report.get("headline_metrics", []):
        headline_cards += f"""
        <div class="metric-card">
          <div class="metric-label">{_escape(item.get("label"))}</div>
          <div class="metric-value">{_escape(item.get("value"))}</div>
          <div class="metric-insight">{_escape(item.get("insight"))}</div>
        </div>
        """

    snapshots = ""
    for item in report.get("monthly_snapshots", []):
        snapshots += f"""
        <div class="snapshot-card">
          <div class="snapshot-month">{_escape(item.get("month"))}</div>
          <div class="snapshot-summary">{_escape(item.get("summary"))}</div>
        </div>
        """

    warning_block = ""
    if warning:
        warning_block = f"""
        <div class="warning-box">
          {_escape(warning)}
        </div>
        """

    return f"""
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>{title}</title>
  <style>
    @page {{
      size: A4;
      margin: 18mm;
    }}
    body {{
      font-family: Arial, Helvetica, sans-serif;
      color: #111827;
      margin: 0;
      background: #ffffff;
      font-size: 12px;
      line-height: 1.45;
    }}
    h1, h2, h3, p {{
      margin-top: 0;
    }}
    .cover {{
      padding: 28px 24px;
      border: 1px solid #dbe3f0;
      border-radius: 16px;
      margin-bottom: 18px;
      background: linear-gradient(135deg, #072146 0%, #0a3a8c 100%);
      color: white;
    }}
    .title {{
      font-size: 28px;
      font-weight: 700;
      margin-bottom: 8px;
      color: #072146;
    }}
    .subtitle {{
      font-size: 14px;
      color: #4b5563;
      margin-bottom: 14px;
    }}
    .summary {{
      font-size: 14px;
    }}
    .warning-box {{
      margin-top: 12px;
      padding: 10px 12px;
      border-radius: 10px;
      background: #fff7ed;
      border: 1px solid #fdba74;
      color: #9a3412;
      font-size: 11px;
    }}
    .section {{
      margin-bottom: 20px;
      page-break-inside: avoid;
    }}
    .section-title {{
      font-size: 18px;
      font-weight: 700;
      color: #072146;
      margin-bottom: 10px;
      border-bottom: 2px solid #dbe3f0;
      padding-bottom: 6px;
    }}
    .metrics-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr 1fr;
      gap: 10px;
    }}
    .metric-card {{
      border: 1px solid #dbe3f0;
      border-radius: 14px;
      padding: 14px;
      background: #ffffff;
    }}
    .metric-label {{
      font-size: 11px;
      color: #6b7280;
      margin-bottom: 8px;
      text-transform: uppercase;
      letter-spacing: 0.4px;
    }}
    .metric-value {{
      font-size: 22px;
      font-weight: 700;
      color: #072146;
      margin-bottom: 6px;
    }}
    .metric-insight {{
      font-size: 12px;
      color: #374151;
    }}
    .two-col {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }}
    .box {{
      border: 1px solid #dbe3f0;
      border-radius: 14px;
      padding: 14px;
      background: #ffffff;
    }}
    .box h3 {{
      font-size: 14px;
      color: #072146;
      margin-bottom: 8px;
    }}
    ul {{
      margin: 0;
      padding-left: 18px;
    }}
    li {{
      margin-bottom: 6px;
    }}
    .snapshots {{
      display: grid;
      grid-template-columns: 1fr;
      gap: 10px;
    }}
    .snapshot-card {{
      border: 1px solid #dbe3f0;
      border-radius: 12px;
      padding: 12px;
      background: #ffffff;
    }}
    .snapshot-month {{
      font-weight: 700;
      color: #072146;
      margin-bottom: 4px;
    }}
    .snapshot-summary {{
      color: #374151;
    }}
    .footer {{
      margin-top: 18px;
      font-size: 10px;
      color: #6b7280;
    }}
    .muted {{
      color: #6b7280;
    }}
  </style>
</head>
<body>
  <div class="cover">
    <div class="title">{title}</div>
    <div class="subtitle">{subtitle}</div>
    <div class="summary">{executive_summary}</div>
    {warning_block}
  </div>

  <div class="section">
    <div class="section-title">KPIs destacados</div>
    <div class="metrics-grid">
      {headline_cards}
    </div>
  </div>

  <div class="section">
    <div class="section-title">Lectura ejecutiva</div>
    <div class="two-col">
      <div class="box">
        <h3>Fortalezas</h3>
        {_html_list(report.get("key_wins", []))}
      </div>
      <div class="box">
        <h3>Alertas y oportunidades</h3>
        {_html_list(report.get("watchouts", []))}
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Evolución del período</div>
    <div class="snapshots">
      {snapshots}
    </div>
  </div>

  <div class="section">
    <div class="section-title">Próximos pasos</div>
    <div class="box">
      {_html_list(report.get("next_steps", []))}
    </div>
  </div>

  <div class="footer">
    Generado automáticamente el {datetime.now().strftime("%d/%m/%Y %H:%M")}
  </div>
</body>
</html>
""".strip()


def _write_report_artifacts(period_slug: str, report: dict[str, Any]) -> Path:
    report_dir = _report_dir(period_slug)

    metadata = {
        "title": report.get("title"),
        "subtitle": report.get("subtitle"),
        "email_subject": report.get("email_subject"),
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "period_slug": period_slug,
        "warning": report.get("warning"),
    }

    html_content = _render_report_html(report)

    (report_dir / "metadata.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (report_dir / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (report_dir / "report.html").write_text(html_content, encoding="utf-8")

    _render_pdf_if_possible(report_dir / "report.html", report_dir / "report.pdf")
    return report_dir


def _render_pdf_if_possible(html_path: Path, pdf_path: Path) -> None:
    try:
        from weasyprint import HTML  # type: ignore
    except Exception as e:
        print(f"WeasyPrint no disponible. Se omite PDF: {e}")
        return

    try:
        HTML(filename=str(html_path)).write_pdf(str(pdf_path))
        print(f"PDF generado: {pdf_path}")
    except Exception as e:
        print(f"No se pudo generar PDF desde HTML: {e}")


def generate_period_report(period_slug: str, force_regenerate: bool = False) -> dict[str, Any]:
    period = _get_period_definition(period_slug)
    months = period.get("months", [])

    if not months:
        raise RuntimeError(f"El período {period_slug} no tiene meses definidos.")

    client = _build_genai_client()
    monthly_summaries: list[dict[str, Any]] = []

    allow_heuristic_fallback = _env_bool("ALLOW_HEURISTIC_FALLBACK", True)

    for month_key in months:
        try:
            summary = summarize_month(client, month_key, force_regenerate=force_regenerate)
        except Exception as e:
            if not allow_heuristic_fallback:
                raise
            print(f"Fallo resumen visual IA para {month_key}. Se usa fallback local: {e}")
            summary = _fallback_month_summary(month_key)
        monthly_summaries.append(summary)

    payload = {
        "period": period,
        "monthly_summaries": monthly_summaries,
    }

    generation_mode = "ai_multimodal"
    warning = None

    try:
        # El reporte maestro que unifica consolida texto, así que este prompt sí va sin archivo adjunto
        contents = [PERIOD_REPORT_PROMPT + "\n\nINPUT:\n" + json.dumps(payload, ensure_ascii=False)]
        report = _call_gemini_for_json(client, contents)
        report["title"] = report.get("title") or period.get("title")
        report["subtitle"] = report.get("subtitle") or period.get("subtitle")
        report["email_subject"] = report.get("email_subject") or period.get("email_subject")
    except Exception as e:
        if not allow_heuristic_fallback:
            raise
        print(f"Fallo generación consolidada IA para {period_slug}. Se usa fallback local: {e}")
        report = _fallback_period_report(period, monthly_summaries, str(e))
        generation_mode = "fallback"
        warning = str(e)

    report_dir = _write_report_artifacts(period_slug, report)

    result = {
        "status": "ok",
        "period_slug": period_slug,
        "generation_mode": generation_mode,
        "report_dir": str(report_dir),
        "warning": warning,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def main() -> None:
    period_slug = (os.environ.get("REPORT_SLUG") or "").strip()
    if not period_slug:
        raise RuntimeError("Falta REPORT_SLUG para generar el reporte.")
    generate_period_report(period_slug)


if __name__ == "__main__":
    main()