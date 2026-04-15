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
- Estás analizando VISUALMENTE el Dashboard de métricas adjunto en formato PDF.
- El equipo de CI consume la información en formato "Slides" ejecutivas.
- Presta especial atención a gráficos de tendencias, variaciones marcadas en colores (ej. verde/rojo para subidas/bajadas), NPS, y distribución en gráficos circulares.

Objetivo:
- Sintetizar el mes con mirada gerencial interpretando los datos visuales.
- Identificar qué funcionó, qué no y por qué.

Formato JSON estricto requerido:
{
  "period_label": "Mar 2026",
  "executive_summary": "2 frases máximo con lectura de negocio basada en tendencias.",
  "headline_metrics": [
    {"label": "Apertura Mail", "value": "81%", "insight": "Alta vs benchmark interno"},
    {"label": "Interacción", "value": "14.1%", "insight": "Disparidad detectada en gráficos"},
    {"label": "NPS", "value": "61", "insight": "Subida significativa"}
  ],
  "highlights": [
    "Insight claro basado en datos y gráficos visuales",
    "Patrón de comportamiento"
  ],
  "opportunities": [
    "Mejora concreta accionable"
  ],
  "content_focus": [
    {"theme": "Beneficios", "detail": "Mayor engagement según barras"}
  ]
}

Reglas:
- estilo BBVA: corto, directo, accionable.
- no más de 12-15 palabras por bullet.
- No inventar métricas: extrae los números directamente de las tablas y gráficos del PDF.
""".strip()


PERIOD_REPORT_PROMPT = """
Sos analista senior de comunicaciones internas en BBVA Argentina.

Contexto:
- Estás creando el contenido para una PRESENTACIÓN EJECUTIVA (Slides) de Dirección.
- Debe poder leerse en 2 minutos.
- Estilo BBVA: Azul marino, números grandes, textos hiper-cortos, foco en la acción.

Objetivo:
- Consolidar los insights de los meses procesados.
- Armar la estructura exacta para las diapositivas de reporte.

Formato JSON estricto requerido:
{
  "cover": {
    "title": "Informe Gestión",
    "subtitle": "Comunicaciones Internas",
    "period": "Período..."
  },
  "executive_summary": "2 frases contundentes con la lectura global del período para abrir la presentación.",
  "slide_kpis": [
    {"label": "Tasa Apertura", "value": "xx%", "insight": "Evolución positiva vs anterior"},
    {"label": "Interacción", "value": "xx%", "insight": "Canales digitales lideran"},
    {"label": "Volumen", "value": "xx", "insight": "Foco en reducción de saturación"}
  ],
  "slide_fortalezas": [
    "Qué funcionó muy bien y por qué (max 10 palabras).",
    "Contenido estrella del período."
  ],
  "slide_alertas": [
    "Riesgo, saturación o caída detectada.",
    "Canal o tema con bajo rendimiento."
  ],
  "slide_next_steps": [
    "Acción concreta a implementar 1",
    "Ajuste táctico 2"
  ]
}

Reglas:
- Máxima capacidad de síntesis.
- Sin verbosidad, directo al grano.
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

        prompt_text = f"{MONTHLY_SUMMARY_PROMPT}\n\nAnaliza visualmente el dashboard adjunto del período {month_key}."
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
            except Exception as e:
                pass


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
        "executive_summary": "Modo contingencia activo. Análisis visual no disponible.",
        "headline_metrics": [
            {"label": "Modo", "value": "Fallback", "insight": "Regenerar cuando IA esté lista."}
        ],
        "highlights": ["Se recuperaron los PDFs."],
        "opportunities": ["Revisar dashboard original."],
        "content_focus": []
    }


def _fallback_period_report(period: dict[str, Any], monthly_summaries: list[dict[str, Any]], error_message: str | None = None) -> dict[str, Any]:
    warning = "Reporte en modo contingencia."
    if error_message:
        warning = f"{warning} Error: {error_message}"

    return {
        "cover": {
            "title": "Informe Gestión Fallback",
            "subtitle": "Comunicaciones Internas",
            "period": period.get("label", period.get("slug"))
        },
        "executive_summary": "Se generó una versión base por caída temporal del servicio de IA.",
        "slide_kpis": [
            {"label": "Meses", "value": str(len(period.get('months', []))), "insight": "Consolidados correctamente."},
            {"label": "Modo", "value": "Manual", "insight": "Revisar tableros fuente."}
        ],
        "slide_fortalezas": ["Continuidad operativa garantizada."],
        "slide_alertas": ["Faltan insights cualitativos."],
        "slide_next_steps": ["Reintentar generación IA más tarde."],
        "warning": warning,
    }


def _escape(text: Any) -> str:
    return html.escape(str(text or ""))


def _html_list(items: list[str], icon: str = "▪") -> str:
    rows = "".join(f"<li style='margin-bottom: 12px;'><span style='color: #1464A5; margin-right: 10px;'>{icon}</span>{_escape(item)}</li>" for item in items if str(item).strip())
    return f"<ul style='list-style-type: none; padding-left: 0; margin-top: 15px; font-size: 16px; color: #374151;'>{rows}</ul>" if rows else "<p>Sin datos.</p>"


def _render_report_html(report: dict[str, Any]) -> str:
    cover = report.get("cover", {})
    title = _escape(cover.get("title", "Informe de Gestión"))
    subtitle = _escape(cover.get("subtitle", "Comunicaciones Internas"))
    period = _escape(cover.get("period", ""))
    
    executive_summary = _escape(report.get("executive_summary"))
    warning = report.get("warning")

    # Slide 2: KPIs
    kpi_cards = ""
    for item in report.get("slide_kpis", []):
        kpi_cards += f"""
        <div class="kpi-card">
          <div class="kpi-label">{_escape(item.get("label"))}</div>
          <div class="kpi-value">{_escape(item.get("value"))}</div>
          <div class="kpi-insight">{_escape(item.get("insight"))}</div>
        </div>
        """

    warning_block = f'<div class="warning-box">{_escape(warning)}</div>' if warning else ""

    return f"""
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>{title} - {period}</title>
  <style>
    @page {{
      size: A4 landscape;
      margin: 0;
    }}
    body {{
      font-family: "Segoe UI", Arial, Helvetica, sans-serif;
      margin: 0;
      padding: 0;
      background: #f4f4f4;
      color: #111827;
      -webkit-print-color-adjust: exact;
    }}
    .slide {{
      width: 100vw;
      height: 100vh;
      box-sizing: border-box;
      page-break-after: always;
      position: relative;
      overflow: hidden;
      background: white;
    }}
    /* SLIDE 1: COVER */
    .bg-navy {{
      background-color: #072146;
      color: white;
      display: flex;
      flex-direction: column;
      justify-content: center;
      padding: 60px 80px;
    }}
    .brand-logo {{
      position: absolute;
      top: 50px;
      left: 80px;
      font-size: 36px;
      font-weight: 800;
      letter-spacing: 2px;
      color: white;
    }}
    .cover-title {{ font-size: 64px; font-weight: 300; margin: 0 0 10px 0; line-height: 1.1; }}
    .cover-subtitle {{ font-size: 64px; font-weight: 700; margin: 0 0 40px 0; line-height: 1.1; color: #4bd4ff; }}
    .cover-period {{ font-size: 24px; color: #dbe3f0; border-left: 4px solid #1464A5; padding-left: 15px; }}
    
    /* SLIDES INTERNAS */
    .slide-content {{
      padding: 50px 80px;
      height: 100%;
      display: flex;
      flex-direction: column;
    }}
    .header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      border-bottom: 2px solid #e5e7eb;
      padding-bottom: 20px;
      margin-bottom: 30px;
    }}
    .slide-title {{
      font-size: 32px;
      font-weight: 700;
      color: #072146;
      margin: 0;
    }}
    .bbva-tag {{ font-weight: bold; color: #1464A5; font-size: 20px; }}
    
    .executive-box {{
      background: #f0f7fb;
      border-left: 6px solid #1464A5;
      padding: 25px;
      border-radius: 0 8px 8px 0;
      font-size: 18px;
      color: #072146;
      margin-bottom: 40px;
      line-height: 1.5;
    }}

    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 30px;
      margin-top: 10px;
    }}
    .kpi-card {{
      background: white;
      border: 1px solid #e5e7eb;
      border-top: 6px solid #4bd4ff;
      border-radius: 12px;
      padding: 30px 20px;
      text-align: center;
      box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    }}
    .kpi-value {{ font-size: 56px; font-weight: 800; color: #072146; margin: 10px 0; }}
    .kpi-label {{ font-size: 14px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 1px; }}
    .kpi-insight {{ font-size: 14px; color: #1464A5; font-weight: 600; margin-top: 10px; background: #f0f7fb; padding: 6px; border-radius: 4px; }}

    .two-col {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 40px;
      flex-grow: 1;
    }}
    .content-box {{
      background: white;
      border: 1px solid #e5e7eb;
      border-radius: 12px;
      padding: 30px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    }}
    .content-box h3 {{ font-size: 20px; color: #072146; margin-top: 0; border-bottom: 1px solid #e5e7eb; padding-bottom: 10px; }}
    
    .warning-box {{
      position: absolute; bottom: 20px; right: 20px;
      background: #fff7ed; border: 1px solid #fdba74; color: #9a3412;
      padding: 8px 15px; border-radius: 6px; font-size: 12px;
    }}
    .footer-bar {{
      position: absolute; bottom: 0; left: 0; width: 100%; height: 12px;
      background: linear-gradient(90deg, #072146 0%, #1464A5 50%, #4bd4ff 100%);
    }}
  </style>
</head>
<body>

  <div class="slide bg-navy">
    <div class="brand-logo">BBVA</div>
    <div style="margin-top: 80px;">
      <h1 class="cover-title">{title}</h1>
      <h2 class="cover-subtitle">{subtitle}</h2>
      <div class="cover-period">{period}</div>
    </div>
    {warning_block}
    <div class="footer-bar"></div>
  </div>

  <div class="slide">
    <div class="slide-content">
      <div class="header">
        <h2 class="slide-title">Lectura Ejecutiva</h2>
        <span class="bbva-tag">BBVA</span>
      </div>
      <div class="executive-box">
        <strong>Síntesis del período:</strong><br>
        {executive_summary}
      </div>
      <h3 style="color: #6b7280; font-size: 16px; text-transform: uppercase; margin-bottom: 10px;">KPIs Destacados</h3>
      <div class="kpi-grid">
        {kpi_cards}
      </div>
    </div>
    <div class="footer-bar"></div>
  </div>

  <div class="slide">
    <div class="slide-content">
      <div class="header">
        <h2 class="slide-title">Análisis de Gestión</h2>
        <span class="bbva-tag">BBVA</span>
      </div>
      <div class="two-col">
        <div class="content-box" style="border-top: 4px solid #10b981;">
          <h3 style="color: #10b981;">Fortalezas e Hitos</h3>
          {_html_list(report.get("slide_fortalezas", []), "✓")}
        </div>
        <div class="content-box" style="border-top: 4px solid #f59e0b;">
          <h3 style="color: #f59e0b;">Alertas y Puntos de Atención</h3>
          {_html_list(report.get("slide_alertas", []), "⚠")}
        </div>
      </div>
    </div>
    <div class="footer-bar"></div>
  </div>

  <div class="slide">
    <div class="slide-content">
      <div class="header">
        <h2 class="slide-title">Próximos Pasos</h2>
        <span class="bbva-tag">BBVA</span>
      </div>
      <div class="content-box" style="flex-grow: 1; border-top: 4px solid #1464A5;">
        <h3>Acciones a implementar</h3>
        {_html_list(report.get("slide_next_steps", []), "➔")}
      </div>
    </div>
    <div class="footer-bar"></div>
  </div>

</body>
</html>
""".strip()


def _write_report_artifacts(period_slug: str, report: dict[str, Any]) -> Path:
    report_dir = _report_dir(period_slug)

    metadata = {
        "title": report.get("cover", {}).get("title"),
        "subtitle": report.get("cover", {}).get("subtitle"),
        "email_subject": f"Informe Ejecutivo CI | {period_slug}",
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
        print(f"PDF generado exitosamente (Modo Slide): {pdf_path}")
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
            print(f"Fallo resumen visual IA para {month_key}. Fallback local: {e}")
            summary = _fallback_month_summary(month_key)
        monthly_summaries.append(summary)

    payload = {
        "period": period,
        "monthly_summaries": monthly_summaries,
    }

    generation_mode = "ai_multimodal"
    warning = None

    try:
        contents = [PERIOD_REPORT_PROMPT + "\n\nINPUT PARA LA PRESENTACIÓN:\n" + json.dumps(payload, ensure_ascii=False)]
        report = _call_gemini_for_json(client, contents)
    except Exception as e:
        if not allow_heuristic_fallback:
            raise
        print(f"Fallo generación consolidada IA para {period_slug}. Fallback local: {e}")
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
