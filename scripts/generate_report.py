import json
import os
from pathlib import Path
from openai import OpenAI

DATA_DIR = Path('data')
TEXT_PATH = DATA_DIR / 'pdf_text.txt'
META_PATH = DATA_DIR / 'metadata.json'
REPORT_JSON = DATA_DIR / 'report.json'
REPORT_HTML = DATA_DIR / 'report.html'

SYSTEM_PROMPT = '''Sos un analista de datos senior de comunicaciones internas.
Tu tarea es leer el texto extraido de un dashboard PDF y redactar un informe ejecutivo en espanol para BBVA Argentina.

Objetivos:
- Resumir hallazgos del periodo.
- Destacar KPIs de mailing cuando esten disponibles.
- Destacar volumen de contenidos publicados y comunicaciones planificadas.
- Marcar patrones, anomalias, oportunidades y recomendaciones.
- No inventes datos. Si algo no esta claro, decilo.

Devolve JSON valido con esta estructura:
{
  "subject": "...",
  "headline": "...",
  "summary": "parrafo breve",
  "kpis": [
    {"label": "...", "value": "...", "comment": "..."}
  ],
  "insights": ["...", "...", "..."],
  "risks": ["...", "..."],
  "recommendations": ["...", "...", "..."],
  "closing": "..."
}
'''

HTML_TEMPLATE = '''<html><body style="font-family:Arial,Helvetica,sans-serif;line-height:1.45;color:#111;">
<h2>{headline}</h2>
<p>{summary}</p>
<h3>KPIs destacados</h3>
<ul>
{kpi_items}
</ul>
<h3>Insights</h3>
<ul>
{insight_items}
</ul>
<h3>Riesgos o alertas</h3>
<ul>
{risk_items}
</ul>
<h3>Recomendaciones</h3>
<ul>
{recommendation_items}
</ul>
<p>{closing}</p>
<hr>
<p style="font-size:12px;color:#666;">Reporte generado automaticamente a partir del PDF del dashboard recibido por mail.</p>
</body></html>'''


def li(items):
    return '\n'.join(f'<li>{x}</li>' for x in items)


def main() -> None:
    client = OpenAI(api_key=os.environ['OPENAI_API_KEY'])
    model = os.environ.get('OPENAI_MODEL', 'gpt-5-mini')
    text = TEXT_PATH.read_text(encoding='utf-8')
    meta = json.loads(META_PATH.read_text(encoding='utf-8'))

    user_prompt = f'''Metadatos del mail:
{json.dumps(meta, ensure_ascii=False, indent=2)}

Texto extraido del PDF:
{text}
'''

    response = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        text={"format": {"type": "json_object"}},
    )

    payload = json.loads(response.output_text)
    REPORT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')

    html = HTML_TEMPLATE.format(
        headline=payload['headline'],
        summary=payload['summary'],
        kpi_items='\n'.join(
            f"<li><strong>{k['label']}:</strong> {k['value']} — {k['comment']}</li>" for k in payload.get('kpis', [])
        ),
        insight_items=li(payload.get('insights', [])),
        risk_items=li(payload.get('risks', [])),
        recommendation_items=li(payload.get('recommendations', [])),
        closing=payload.get('closing', ''),
    )
    REPORT_HTML.write_text(html, encoding='utf-8')
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
