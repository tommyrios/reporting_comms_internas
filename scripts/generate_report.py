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

# URL del logo oficial proporcionado
BBVA_LOGO_URL = "https://upload.wikimedia.org/wikipedia/commons/9/98/BBVA_logo_2025.svg"


MONTH_NAMES_SHORT = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr",
    5: "may", 6: "jun", 7: "jul", 8: "ago",
    9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


MONTHLY_SUMMARY_PROMPT = """
Sos analista senior de comunicaciones internas en BBVA Argentina.

Contexto:
- Estás analizando VISUALMENTE el Dashboard de métricas adjunto en formato PDF.
- Presta especial atención a gráficos de tendencias, variaciones marcadas en colores, NPS, segmentación de audiencia, y rankings de notas (tanto Push como Pull/Intranet).

Objetivo:
- Sintetizar el mes interpretando los datos visuales para alimentar un reporte consolidado.

Formato JSON estricto requerido:
{
  "period_label": "Mar 2026",
  "data": {
    "push_volume": 12, 
    "push_opens": "81%",
    "push_interaction": "14.1%",
    "pull_notes": 45,
    "pull_reads": 12500,
    "nps": 61
  },
  "insights": {
    "audience_segmentation": "Texto breve sobre a quién se envió (ej. Todo el banco 80%)",
    "strategic_axes": "Temas principales detectados visualmente",
    "internal_clients": "Áreas que más solicitaron",
    "top_push_comm": "Nombre de la nota Push más exitosa",
    "top_pull_note": "Nombre de la nota Intranet más leída"
  }
}
""".strip()


PERIOD_REPORT_PROMPT = """
Sos analista senior de comunicaciones internas en BBVA Argentina.

Contexto:
- Estás creando el contenido exacto para una PRESENTACIÓN DE GESTIÓN consolidada (Slides apaisadas).
- Basándote en los resúmenes mensuales provistos, debes estructurar la información siguiendo ESTRICTAMENTE el modelo de 7 puntos definido por la dirección.

Objetivo:
- Generar un JSON que contenga los datos y textos sintéticos para cada una de las 7 diapositivas.

Formato JSON estricto requerido:
{
  "slide_1_cover": {
    "area": "Área/Departamento de...",
    "period": "Periodo (ej. Año 2024)"
  },
  "slide_2_overview": {
    "volume_current": "xx",
    "volume_previous": "xx",
    "volume_change": "+x%",
    "audience_segments": [
      {"label": "Todo el Banco", "value": 75},
      {"label": "Red Sucursales", "value": 15},
      {"label": "Líderes/Áreas Centrales", "value": 10}
    ],
    "conclusion_message": "Breve mensaje sobre logro de objetivos de impacto."
  },
  "slide_3_strategy": {
    "content_distribution": [
      {"theme": "Negocio", "weight": 40},
      {"theme": "Equipo", "weight": 30},
      {"theme": "Innovación", "weight": 20},
      {"theme": "Sostenibilidad", "weight": 10}
    ],
    "internal_clients": [
      {"label": "RRHH", "value": 35},
      {"label": "Negocio", "value": 25},
      {"label": "Riesgos", "value": 20},
      {"label": "Otros", "value": 20}
    ],
    "canal_balance": {"institutional": 60, "transactional_talent": 40}
  },
  "slide_4_push_ranking": {
    "top_communications": [
      {"name": "Nombre Comm 1", "clicks": "x.xxx", "interaction": "xx%"},
      {"name": "Nombre Comm 2", "clicks": "x.xxx", "interaction": "xx%"}
    ],
    "key_learning": "Una frase destacada que explique qué funcionó mejor."
  },
  "slide_5_pull_performance": {
    "pub_current": "xx",
    "pub_previous": "xx",
    "top_notes": [
      {"title": "Nota Intranet 1", "unique_reads": "x.xxx", "total_reads": "x.xxx"},
      {"title": "Nota Intranet 2", "unique_reads": "x.xxx", "total_reads": "x.xxx"}
    ],
    "avg_reads": "x.xxx",
    "total_views": "xxx.xxx"
  },
  "slide_6_hitos": [
    {"quarter": "Q1", "description": "Resumen sintético hitos Q1"},
    {"quarter": "Q2", "description": "Resumen sintético hitos Q2"}
  ],
  "slide_7_events": {
    "total_events": "xx",
    "total_participants": "xx.xxx",
    "conclusion": "Mensaje de cierre anual de eventos."
  }
}

Reglas:
- Máxima capacidad de síntesis: textos cortos pensados para slides.
- Extrae o estima los porcentajes de segmentación y estrategia basándote en los insights cualitativos de los resúmenes mensuales si no hay números exactos.
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
                    print(f"Gemini saturado. Reintento en {wait_seconds:.1f}s")
                    time.sleep(wait_seconds)
                    continue
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
        uploaded_file = client.files.upload(file=str(pdf_path))
        contents = [uploaded_file, MONTHLY_SUMMARY_PROMPT]
        summary = _call_gemini_for_json(client, contents)
        summary["month"] = month_key
        summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        return summary
    finally:
        if uploaded_file:
            try: client.files.delete(name=uploaded_file.name)
            except: pass


def _fallback_period_report(period: dict[str, Any]) -> dict[str, Any]:
    return {
        "slide_1_cover": {"area": "Comunicaciones Internas", "period": period.get("label", "")},
        "warning": "Reporte generado en modo fallback por indisponibilidad de IA."
    }


def _escape(text: Any) -> str:
    return html.escape(str(text or ""))

# Ayudante para gráficos de barras CSS simples
def _css_bar(value_percent: int, color: str = "#1464A5") -> str:
    return f"""
    <div style="width: 100%; background: #e5e7eb; border-radius: 4px; height: 10px; margin-top: 5px;">
      <div style="width: {value_percent}%; background: {color}; height: 10px; border-radius: 4px;"></div>
    </div>
    """

# Ayudante para gráficos circulares CSS simples (falsos)
def _css_pie(value_percent: int, color: str = "#1464A5") -> str:
    return f"""
    <div style="width: 50px; height: 50px; border-radius: 50%; background: conic-gradient({color} {value_percent}%, #e5e7eb 0%); margin: auto;"></div>
    """

def _render_report_html(report: dict[str, Any]) -> str:
    warning = report.get("warning")
    warning_block = f'<div class="warning-box">{_escape(warning)}</div>' if warning else ""
    
    # Comunes
    header_html = f'<div class="header"><img src="{BBVA_LOGO_URL}" class="logo-top"></div>'
    footer_html = '<div class="footer-bar"></div>'

    # --- 1. Portada ---
    s1 = report.get("slide_1_cover", {})
    slide_1 = f"""
    <div class="slide bg-navy">
      <div class="slide-content-cover">
        <h1 class="cover-title">Informe de Gestión</h1>
        <h1 class="cover-title-area">{_escape(s1.get("area"))}</h1>
        <div class="spacer"></div>
        <h2 class="cover-subtitle">{_escape(s1.get("period"))}</h2>
        <div class="eslogan">Creando oportunidades</div>
      </div>
      {warning_block}
      <img src="{BBVA_LOGO_URL}" class="logo-cover">
    </div>
    """

    # --- 2. Visión General ---
    s2 = report.get("slide_2_overview", {})
    audience_rows = ""
    for seg in s2.get("audience_segments", []):
        audience_rows += f"""
        <div class="segment-item">
          {_css_pie(seg['value'])}
          <div class="segment-label">{_escape(seg['label'])} ({seg['value']}%)</div>
        </div>
        """
    
    slide_2 = f"""
    <div class="slide">
      {header_html}
      <div class="slide-content">
        <h2 class="slide-title">2. Visión General y Audiencia (Canal Push)</h2>
        <div class="two-col">
          <div class="box">
            <h3>Evolución del Volumen</h3>
            <table class="data-table">
              <thead><tr><th>Periodo</th><th>Envíos</th></tr></thead>
              <tbody>
                <tr><td>Actual</td><td class="val-big">{_escape(s2.get("volume_current"))}</td></tr>
                <tr><td>Anterior</td><td>{_escape(s2.get("volume_previous"))}</td></tr>
                <tr><td colspan="2" class="change-cell">Cambio: {_escape(s2.get("volume_change"))}</td></tr>
              </tbody>
            </table>
          </div>
          <div class="box">
            <h3>Segmentación de Audiencia</h3>
            <div class="audience-grid">{audience_rows}</div>
          </div>
        </div>
        <div class="conclusion-box">{_escape(s2.get("conclusion_message"))}</div>
      </div>
      {footer_html}
    </div>
    """

    # --- 3. Ejes Estratégicos ---
    s3 = report.get("slide_3_strategy", {})
    dist_rows = ""
    for axis in s3.get("content_distribution", []):
        dist_rows += f"<li>{_escape(axis['theme'])} ({axis['weight']}%) {_css_bar(axis['weight'], '#4bd4ff')}</li>"

    client_rows = ""
    for cl in s3.get("internal_clients", []):
        client_rows += f"<li>{_escape(cl['label'])} ({cl['value']}%) {_css_bar(cl['value'])}</li>"
        
    balance = s3.get("canal_balance", {})

    slide_3 = f"""
    <div class="slide">
      {header_html}
      <div class="slide-content">
        <h2 class="slide-title">3. Ejes Estratégicos y Origen del Contenido</h2>
        <div class="three-col">
          <div class="box">
            <h3>Distribución del Contenido</h3>
            <ul class="bar-list">{dist_rows}</ul>
          </div>
          <div class="box">
            <h3>Clientes Internos (Top)</h3>
            <ul class="bar-list">{client_rows}</ul>
          </div>
          <div class="box balance-box">
            <h3>Balance del Canal</h3>
            <div class="balance-item">
              <div class="percent navy">{balance.get('institutional')}%</div>
              <div>Institucional</div>
            </div>
            <div class="balance-item">
              <div class="percent blue">{balance.get('transactional_talent')}%</div>
              <div>Transaccional / Talento</div>
            </div>
          </div>
        </div>
      </div>
      {footer_html}
    </div>
    """

    # --- 4. Ranking Impacto ---
    s4 = report.get("slide_4_push_ranking", {})
    ranking_rows = ""
    for i, comm in enumerate(s4.get("top_communications", [])):
        ranking_rows += f"""
        <div class="rank-itembox">
          <div class="rank-number">#{i+1}</div>
          <div class="rank-visual">[INSERTAR CAPTURA PIEZA GRÁFICA]</div>
          <div class="rank-data">
            <strong>{_escape(comm['name'])}</strong><br>
            Clicks: {comm['clicks']} | Interacción: {comm['interaction']}
          </div>
        </div>
        """

    slide_4 = f"""
    <div class="slide">
      {header_html}
      <div class="slide-content">
        <h2 class="slide-title">4. Ranking de Mayor Impacto (Canal Push)</h2>
        <div class="rank-grid">{ranking_rows}</div>
        <div class="learning-quote"><strong>Aprendizaje:</strong> {_escape(s4.get("key_learning"))}</div>
      </div>
      {footer_html}
    </div>
    """

    # --- 5. Desempeño Pull ---
    s5 = report.get("slide_5_pull_performance", {})
    notes_rows = ""
    for note in s5.get("top_notes", []):
        notes_rows += f"<tr><td>{_escape(note['title'])}</td><td>{note['unique_reads']}</td><td>{note['total_reads']}</td></tr>"

    slide_5 = f"""
    <div class="slide">
      {header_html}
      <div class="slide-content">
        <h2 class="slide-title">5. Desempeño del Canal Pull (Intranet / App)</h2>
        <div class="two-col-uneven">
          <div class="box">
            <h3>Volumen de Publicación (vs Año Anterior)</h3>
            <table class="data-table small">
              <thead><tr><th>Periodo</th><th>Notas Publicadas</th></tr></thead>
              <tbody>
                <tr><td>Actual</td><td class="val-big blue">{_escape(s5.get("pub_current"))}</td></tr>
                <tr><td>Anterior</td><td>{_escape(s5.get("pub_previous"))}</td></tr>
              </tbody>
            </table>
            <div class="kpi-pull-row">
              <div class="kpi-pull-card">Promedio Lecturas: <strong>{s5.get('avg_reads')}</strong></div>
              <div class="kpi-pull-card">Total Visualizaciones: <strong>{s5.get('total_views')}</strong></div>
            </div>
          </div>
          <div class="box">
            <h3>Ranking Top Notas más leídas</h3>
            <table class="data-table small ranking-table">
              <thead><tr><th>Título Nota</th><th>Lecturas Únicas</th><th>Totales</th></tr></thead>
              <tbody>{notes_rows}</tbody>
            </table>
          </div>
        </div>
      </div>
      {footer_html}
    </div>
    """

    # --- 6. Hitos Showcase ---
    hitos = report.get("slide_6_hitos", [])
    hitos_blocks = ""
    for hit in hitos:
        hitos_blocks += f"""
        <div class="hito-block box">
          <h3>{_escape(hit['quarter'])}</h3>
          <p>{_escape(hit['description'])}</p>
          <div class="visual-placeholder">[COLLAGE / CAPTURAS DEL TRIMESTRE]</div>
        </div>
        """

    slide_6 = f"""
    <div class="slide bg-navy hitos-slide">
      {header_html}
      <div class="slide-content">
        <h2 class="slide-title light">6. Hitos Destacados y Visual Showcase</h2>
        <div class="hito-grid">{hitos_blocks}</div>
      </div>
      {footer_html}
    </div>
    """

    # --- 7. Métricas Eventos ---
    s7 = report.get("slide_7_events", {})
    slide_7 = f"""
    <div class="slide">
      {header_html}
      <div class="slide-content">
        <h2 class="slide-title">7. Métricas de Eventos y Transmisiones</h2>
        <div class="box events-main">
          <h3>Tabla de Asistencia (Muestra - *Insertar desglose detallado*)</h3>
          <table class="data-table events-table small">
            <thead><tr><th>Nombre Evento</th><th>Presencial</th><th>Virtual</th><th>Total</th></tr></thead>
            <tbody>
              <tr><td>Ejemplo Evento 1</td><td>xxx</td><td>xxx</td><td><strong>x.xxx</strong></td></tr>
              <tr><td>Ejemplo Evento 2</td><td>xxx</td><td>xxx</td><td><strong>x.xxx</strong></td></tr>
            </tbody>
          </table>
        </div>
        <div class="conclusion-box footer-summarynavy">
          <strong>Resumen Anual:</strong> Cerramos el período con <strong>{s7.get('total_events')}</strong> eventos gestionados y un acumulado de <strong>{s7.get('total_participants')}</strong> participaciones.
        </div>
      </div>
      {footer_html}
    </div>
    """

    # HTML Final consolidado
    return f"""
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <style>
    @page {{ size: A4 landscape; margin: 0; }}
    body {{ font-family: "Segoe UI", Arial, sans-serif; margin: 0; color: #111827; -webkit-print-color-adjust: exact; }}
    .slide {{ width: 100vw; height: 100vh; page-break-after: always; position: relative; overflow: hidden; background: white; box-sizing: border-box; }}
    .bg-navy {{ background-color: #072146; color: white; }}
    .slide-content {{ padding: 40px 60px; height: calc(100% - 100px); display: flex; flex-direction: column; }}
    
    /* Portada */
    .slide-content-cover {{ padding: 100px 80px; }}
    .cover-title {{ font-size: 60px; font-weight: 300; margin: 0; }}
    .cover-title-area {{ font-size: 60px; font-weight: 700; margin: 0 0 40px 0; }}
    .cover-subtitle {{ font-size: 28px; font-weight: 400; color: #4bd4ff; margin: 0; }}
    .eslogan {{ font-size: 18px; color: #dbe3f0; margin-top: 10px; }}
    .logo-cover {{ position: absolute; bottom: 50px; right: 80px; width: 150px; }}
    .spacer {{ height: 80px; }}

    /* Cabecera Interna */
    .header {{ height: 60px; padding: 15px 60px 0 60px; display: flex; align-items: center; border-bottom: 1px solid #e5e7eb; }}
    .logo-top {{ width: 100px; }}
    .slide-title {{ font-size: 28px; font-weight: 700; color: #072146; margin: 0; flex-grow: 1; }}
    .slide-title.light {{ color: white; }}

    /* Layouts */
    .box {{ background: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.03); }}
    .box h3 {{ font-size: 16px; color: #1464A5; margin: 0 0 15px 0; border-bottom: 1px solid #e5e7eb; padding-bottom: 8px; }}
    .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-top: 20px; }}
    .three-col {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; margin-top: 20px; }}
    .two-col-uneven {{ display: grid; grid-template-columns: 1fr 2fr; gap: 30px; margin-top: 20px; }}
    .conclusion-box {{ margin-top: auto; background: #f0f7fb; border-left: 4px solid #1464A5; padding: 15px; border-radius: 4px; color: #072146; }}
    .footer-bar {{ position: absolute; bottom: 0; left: 0; width: 100%; height: 8px; background: linear-gradient(90deg, #072146 0%, #1464A5 50%, #4bd4ff 100%); }}
    .warning-box {{ position: absolute; bottom: 20px; right: 20px; background: #fff7ed; border: 1px solid #fdba74; padding: 8px 12px; border-radius: 4px; font-size: 11px; color: #9a3412; }}

    /* Tablas */
    .data-table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    .data-table th {{ text-align: left; color: #6b7280; padding-bottom: 10px; border-bottom: 1px solid #e5e7eb; }}
    .data-table td {{ padding: 12px 0; border-bottom: 1px solid #f3f4f6; }}
    .val-big {{ font-size: 36px; font-weight: 700; color: #072146; }}
    .val-big.blue {{ color: #1464A5; }}
    .change-cell {{ text-align: center; font-weight: bold; color: #1464A5; background: #f0f7fb; padding: 8px; border-radius: 4px; }}
    .data-table.small td {{ padding: 6px 0; font-size: 12px; }}
    .ranking-table td {{ vertical-align: top; }}

    /* Audiencia (CSS Pies) */
    .audience-grid {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; text-align: center; }}
    .segment-item {{ padding: 10px; }}
    .segment-label {{ font-size: 12px; margin-top: 8px; color: #374151; }}

    /* Ejes y Bar lists */
    .bar-list {{ list-style: none; padding: 0; margin: 0; font-size: 13px; }}
    .bar-list li {{ margin-bottom: 12px; }}
    .balance-box {{ display: flex; flex-direction: column; justify-content: center; text-align: center; }}
    .balance-item {{ margin-bottom: 20px; }}
    .percent {{ font-size: 40px; font-weight: 800; }}
    .percent.navy {{ color: #072146; }}
    .percent.blue {{ color: #1464A5; }}

    /* Ranking */
    .rank-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; flex-grow: 1; }}
    .rank-itembox {{ background: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 15px; display: flex; align-items: center; }}
    .rank-number {{ font-size: 24px; font-weight: bold; color: #1464A5; margin-right: 15px; }}
    .rank-visual {{ width: 100px; height: 60px; background: #f3f4f6; border: 1px dashed #d1d5db; color: #9ca3af; font-size: 9px; text-align: center; display: flex; align-items: center; justify-content: center; border-radius: 4px; margin-right: 15px; padding: 5px; box-sizing: border-box; }}
    .rank-data {{ font-size: 12px; color: #374151; flex-grow: 1; }}
    .learning-quote {{ background: #072146; color: white; padding: 15px; border-radius: 4px; font-size: 14px; margin-top: auto; }}

    /* Pull */
    .kpi-pull-row {{ display: flex; gap: 10px; margin-top: 15px; }}
    .kpi-pull-card {{ flex: 1; background: #f0f7fb; padding: 10px; border-radius: 4px; font-size: 11px; }}

    /* Hitos */
    .hitos-slide .slide-content {{ flex-direction: row; flex-wrap: wrap; gap: 20px; padding-top: 20px; }}
    .hitos-slide .box {{ background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2); color: white; }}
    .hitos-slide .box h3 {{ color: #4bd4ff; border-bottom: 1px solid rgba(255,255,255,0.1); }}
    .hito-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; width: 100%; height: 100%; }}
    .hito-block {{ font-size: 12px; display: flex; flex-direction: column; }}
    .hito-block p {{ margin: 0 0 10px 0; flex-grow: 1; }}
    .visual-placeholder {{ height: 100px; background: rgba(0,0,0,0.2); border: 1px dashed rgba(255,255,255,0.2); border-radius: 4px; color: rgba(255,255,255,0.5); display: flex; align-items: center; justify-content: center; font-size: 10px; text-align: center; padding: 10px; }}

    /* Eventos */
    .events-main {{ flex-grow: 1; margin-bottom: 20px; }}
    .footer-summarynavy {{ background: #072146; color: white; }}
    .footer-summarynavy strong {{ color: #4bd4ff; }}
  </style>
</head>
<body>
  {slide_1}
  {slide_2}
  {slide_3}
  {slide_4}
  {slide_5}
  {slide_6}
  {slide_7}
</body>
</html>
""".strip()


def _write_report_artifacts(period_slug: str, report: dict[str, Any]) -> Path:
    report_dir = _report_dir(period_slug)

    metadata = {
        "title": report.get("slide_1_cover", {}).get("area"),
        "subtitle": report.get("slide_1_cover", {}).get("period"),
        "email_subject": f"Informe Gestión CI | {period_slug}",
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "period_slug": period_slug,
        "warning": report.get("warning"),
    }

    html_content = _render_report_html(report)

    (report_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    (report_dir / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    (report_dir / "report.html").write_text(html_content, encoding="utf-8")

    _render_pdf_if_possible(report_dir / "report.html", report_dir / "report.pdf")
    return report_dir


def _render_pdf_if_possible(html_path: Path, pdf_path: Path) -> None:
    try:
        from weasyprint import HTML # type: ignore
        HTML(filename=str(html_path)).write_pdf(str(pdf_path))
        print(f"PDF generado (7 Puntos Narrativos): {pdf_path}")
    except Exception as e:
        print(f"No se pudo generar PDF: {e}")


def generate_period_report(period_slug: str, force_regenerate: bool = False) -> dict[str, Any]:
    period = _get_period_definition(period_slug)
    months = period.get("months", [])
    if not months: raise RuntimeError(f"No hay meses definidos.")

    client = _build_genai_client()
    monthly_summaries: list[dict[str, Any]] = []
    allow_heuristic_fallback = _env_bool("ALLOW_HEURISTIC_FALLBACK", True)

    for month_key in months:
        try:
            summary = summarize_month(client, month_key, force_regenerate=force_regenerate)
        except Exception as e:
            if not allow_heuristic_fallback: raise
            print(f"Fallback mensual para {month_key}: {e}")
            summary = _fallback_month_summary(month_key)
        monthly_summaries.append(summary)

    payload = {"period": period, "monthly_summaries": monthly_summaries}

    generation_mode = "ai_multimodal"
    warning = None

    try:
        prompt_final = f"{PERIOD_REPORT_PROMPT}\n\nINPUT RESÚMENES MENSUALES:\n{json.dumps(payload, ensure_ascii=False)}"
        report = _call_gemini_for_json(client, [prompt_final])
    except Exception as e:
        if not allow_heuristic_fallback: raise
        print(f"Fallback final para {period_slug}: {e}")
        report = _fallback_period_report(period)
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
    if not period_slug: raise RuntimeError("Falta REPORT_SLUG.")
    generate_period_report(period_slug)

if __name__ == "__main__":
    main()
