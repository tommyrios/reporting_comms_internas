import json
import os
import time
from pathlib import Path
from typing import Any

from google import genai

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
EXTRACTED_DIR = DATA_DIR / "extracted_text"
REPORTS_DIR = OUTPUT_DIR / "reports"


# =========================
# 🔹 PROMPT
# =========================

PERIOD_REPORT_PROMPT = """
Generá un reporte de Comunicaciones Internas estilo presentación ejecutiva BBVA.

Formato: SOLO JSON válido.

Debe incluir:
- cover
- overview (kpis + insights)
- plan_management
- content
- top_impacts
- milestones
- events
- closing

Estilo:
- corto
- visual
- ejecutivo
- no prosa larga
"""

# =========================
# 🔹 HELPERS
# =========================

def _safe_json(text: str):
    text = text.replace("```json", "").replace("```", "")
    return json.loads(text)


def _get(d, *keys, default=""):
    for k in keys:
        if isinstance(d, dict):
            d = d.get(k, {})
        else:
            return default
    return d if d else default


def _client():
    return genai.Client(api_key=os.environ["GEMINI_API_KEY"])


# =========================
# 🔹 DATA
# =========================

def _get_period(period_slug):
    data = json.loads((DATA_DIR / "fetch_result.json").read_text())
    for p in data["periods"]:
        if p["slug"] == period_slug:
            return p
    raise RuntimeError("Periodo no encontrado")


def _load_text(month):
    return (EXTRACTED_DIR / f"{month}.txt").read_text()


# =========================
# 🔹 AI
# =========================

def _call_ai(payload):
    client = _client()

    for i in range(4):
        try:
            res = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=PERIOD_REPORT_PROMPT + json.dumps(payload)
            )
            return _safe_json(res.text)
        except Exception as e:
            print("retry", i, e)
            time.sleep(2 * (i + 1))

    raise RuntimeError("Gemini failed")


# =========================
# 🔹 HTML RENDER PRO
# =========================

def _list(items):
    if not items:
        return "<li>Sin datos</li>"
    return "".join(f"<li>{x}</li>" for x in items)


def _kpis(kpis):
    html = ""
    for k in kpis or []:
        html += f"""
        <div class="kpi">
            <div class="kpi-value">{k.get('value','')}</div>
            <div class="kpi-label">{k.get('label','')}</div>
            <div class="kpi-insight">{k.get('insight','')}</div>
        </div>
        """
    return html


def _render(r):

    return f"""
<html>
<head>
<style>
body {{
 font-family: Arial;
 background:#f4f6f8;
}}

.slide {{
 width: 900px;
 margin:40px auto;
 background:white;
 padding:30px;
 border-radius:12px;
}}

.title {{
 font-size:28px;
 color:#072146;
 font-weight:bold;
}}

.subtitle {{
 color:#6b7280;
}}

.kpi {{
 display:inline-block;
 width:30%;
 text-align:center;
 padding:10px;
}}

.kpi-value {{
 font-size:26px;
 font-weight:bold;
 color:#072146;
}}

.section {{
 margin-top:20px;
}}

</style>
</head>

<body>

<div class="slide">
<div class="title">{_get(r,'cover','title')}</div>
<div class="subtitle">{_get(r,'cover','subtitle')}</div>
</div>

<div class="slide">
<h2>¿Cómo nos fue?</h2>
<p>{_get(r,'overview','summary')}</p>
<div>{_kpis(_get(r,'overview','kpis', default=[]))}</div>
<ul>{_list(_get(r,'overview','insights', default=[]))}</ul>
</div>

<div class="slide">
<h2>Gestión del Plan</h2>
<p><b>Mail:</b> {_get(r,'plan_management','mail','summary')}</p>
<p><b>Intranet:</b> {_get(r,'plan_management','intranet','summary')}</p>
</div>

<div class="slide">
<h2>Contenido</h2>
<p>{_get(r,'content','summary')}</p>
</div>

<div class="slide">
<h2>Top impactos</h2>
<ul>{_list([i.get('communication') for i in r.get('top_impacts',{}).get('items',[])])}</ul>
</div>

<div class="slide">
<h2>Hitos</h2>
<ul>{_list([i.get('headline') for i in r.get('milestones',{}).get('items',[])])}</ul>
</div>

<div class="slide">
<h2>Eventos</h2>
<ul>{_list([i.get('event') for i in r.get('events',{}).get('items',[])])}</ul>
</div>

<div class="slide">
<h2>Cierre</h2>
<h3>Fortalezas</h3>
<ul>{_list(_get(r,'closing','strengths', default=[]))}</ul>
<h3>Alertas</h3>
<ul>{_list(_get(r,'closing','watchouts', default=[]))}</ul>
</div>

</body>
</html>
"""


# =========================
# 🔹 PDF
# =========================

def _pdf(html_path, pdf_path):
    try:
        from weasyprint import HTML
        HTML(filename=str(html_path)).write_pdf(str(pdf_path))
        print("PDF OK")
    except Exception as e:
        print("PDF fail:", e)


# =========================
# 🔹 MAIN
# =========================

def generate_period_report(period_slug):

    period = _get_period(period_slug)
    texts = [_load_text(m) for m in period["months"]]

    payload = {"text": "\n".join(texts)}

    try:
        report = _call_ai(payload)
    except Exception:
        report = {"cover": {"title": period["title"], "subtitle": period["subtitle"]}}

    out = REPORTS_DIR / period_slug
    out.mkdir(parents=True, exist_ok=True)

    html = _render(report)

    (out / "report.html").write_text(html)
    (out / "report.json").write_text(json.dumps(report, indent=2))

    _pdf(out / "report.html", out / "report.pdf")

    return {"status": "ok"}