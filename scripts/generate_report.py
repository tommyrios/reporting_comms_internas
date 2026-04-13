from __future__ import annotations

import json
import os
import re
import time
from html import escape
from pathlib import Path
from typing import List

import sys

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.append(str(SCRIPT_DIR))

from reporting_periods import load_schedule

OUTPUT_DIR = Path("output")
CORPUS_PATH = OUTPUT_DIR / "monthly_corpus.json"
MONTHLY_AI_DIR = OUTPUT_DIR / "monthly_ai_summaries"
REPORTS_DIR = OUTPUT_DIR / "reports"

DEFAULT_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
FALLBACK_MODELS = [
    model.strip()
    for model in os.environ.get("GEMINI_FALLBACK_MODELS", "gemini-2.5-flash-lite,gemini-2.0-flash").split(",")
    if model.strip()
]
MAX_MODEL_RETRIES = int(os.environ.get("GEMINI_MAX_RETRIES", "4"))
INITIAL_BACKOFF_SECONDS = float(os.environ.get("GEMINI_INITIAL_BACKOFF_SECONDS", "3"))
ALLOW_HEURISTIC_FALLBACK = (os.environ.get("ALLOW_HEURISTIC_FALLBACK") or "true").lower() == "true"


MONTHLY_SUMMARY_PROMPT = """
Sos un analista de datos de Comunicaciones Internas de BBVA Argentina.

Recibirás el texto extraído de un dashboard mensual. Tu trabajo es convertirlo en un resumen estructurado, compacto y útil para luego armar un informe trimestral o anual.

Reglas:
- Español.
- No inventes datos ni completes huecos.
- Si un número no es visible, omitilo.
- Priorizá: envíos, aperturas, clics, tasas, lecturas, visitas, segmentación, contenidos destacados, canales y reputación/satisfacción si aparecen.
- Cada bullet debe ser corto y claro.
- Dejá afuera relleno narrativo.

Devolvé JSON válido con esta estructura exacta:
{
  "month": "YYYY-MM",
  "headline": "máximo 18 palabras",
  "kpis": [
    {"label": "", "value": "", "note": ""}
  ],
  "highlights": ["", ""],
  "alerts": ["", ""],
  "top_contents": [
    {"name": "", "metric": "", "reason": ""}
  ],
  "evidence_numbers": [""],
  "source_gaps": [""]
}

Límites:
- kpis: máximo 8
- highlights: máximo 5
- alerts: máximo 4
- top_contents: máximo 4
- evidence_numbers: máximo 10
- source_gaps: máximo 4
""".strip()


FINAL_DECK_PROMPT = """
Sos un analista senior de Comunicaciones Internas de BBVA Argentina.

Vas a recibir resúmenes mensuales ya estructurados. Tenés que transformarlos en un informe visual, corto y fácil de leer, con formato tipo slides para enviar por email.

Objetivo:
- Mostrar rápido qué pasó en el período.
- Resaltar métricas, hallazgos, alertas y recomendaciones.
- Evitar prosa extensa.

Reglas:
- Español.
- Tono ejecutivo, claro y directo.
- No inventar datos.
- Cada bullet: máximo 18 palabras.
- No escribir párrafos largos.
- Elegir solo lo más importante del período.
- Si hay comparación entre meses dentro del período, marcarla.
- Si el período es anual, reflejar evolución del año completo.

Devolvé JSON válido con esta estructura exacta:
{
  "email_subject": "",
  "preheader": "",
  "title": "",
  "subtitle": "",
  "hero_metrics": [
    {"label": "", "value": "", "note": ""}
  ],
  "slides": [
    {
      "title": "Resumen ejecutivo",
      "bullets": ["", ""],
      "callout": ""
    },
    {
      "title": "KPIs del período",
      "kpis": [
        {"label": "", "value": "", "note": ""}
      ],
      "callout": ""
    },
    {
      "title": "Qué funcionó",
      "bullets": ["", ""],
      "callout": ""
    },
    {
      "title": "Alertas y oportunidades",
      "bullets": ["", ""],
      "callout": ""
    },
    {
      "title": "Recomendaciones",
      "bullets": ["", ""],
      "callout": ""
    }
  ],
  "footer_note": ""
}

Límites:
- hero_metrics: entre 3 y 5
- slides: entre 5 y 6
- slide de KPIs: hasta 6 kpis
""".strip()


def get_client():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("Falta GEMINI_API_KEY")

    from google import genai

    return genai.Client(api_key=api_key)


def _clean_json_string(text: str) -> str:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()
    return cleaned


def _call_gemini_for_json(client, prompt: str, payload: dict) -> dict:
    from google.genai import errors, types

    models_to_try = []
    for model in [DEFAULT_MODEL, *FALLBACK_MODELS]:
        if model and model not in models_to_try:
            models_to_try.append(model)

    last_error = None
    for model_name in models_to_try:
        for attempt in range(1, MAX_MODEL_RETRIES + 1):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=f"{prompt}\n\nINPUT_JSON:\n{json.dumps(payload, ensure_ascii=False)}",
                    config=types.GenerateContentConfig(response_mime_type="application/json"),
                )
                raw_text = _clean_json_string(response.text or "")
                parsed = json.loads(raw_text)
                parsed["_meta_model"] = model_name
                parsed["_meta_attempt"] = attempt
                return parsed
            except (errors.ServerError, errors.APIError, json.JSONDecodeError, ValueError) as exc:
                last_error = exc
                is_retryable = False
                status_code = getattr(exc, "status_code", None)
                message = str(exc)
                if status_code in {429, 500, 502, 503, 504}:
                    is_retryable = True
                if "UNAVAILABLE" in message or "high demand" in message.lower():
                    is_retryable = True

                if attempt < MAX_MODEL_RETRIES and is_retryable:
                    sleep_seconds = INITIAL_BACKOFF_SECONDS * (2 ** (attempt - 1))
                    print(f"Gemini saturado con {model_name}. Reintento {attempt}/{MAX_MODEL_RETRIES} en {sleep_seconds:.1f}s")
                    time.sleep(sleep_seconds)
                    continue

                print(f"Fallo Gemini con modelo {model_name} en intento {attempt}: {exc}")
                break

    raise RuntimeError(f"Gemini no respondió después de probar {', '.join(models_to_try)}") from last_error


def load_corpus() -> dict:
    if not CORPUS_PATH.exists():
        raise FileNotFoundError(f"No existe {CORPUS_PATH}")
    return json.loads(CORPUS_PATH.read_text(encoding="utf-8"))


def load_period_by_slug(period_slug: str) -> dict:
    schedule = load_schedule()
    for period in schedule.periods:
        if period.slug == period_slug:
            return period.to_dict()
    raise ValueError(f"No se encontró el período {period_slug}")


def build_month_payload(item: dict) -> dict:
    max_chars = int(os.environ.get("MONTH_TEXT_MAX_CHARS", "70000"))
    return {
        "month": item["month"],
        "filename": item.get("filename"),
        "subject": item.get("subject"),
        "page_count": item.get("page_count"),
        "text": (item.get("text") or "")[:max_chars],
    }


def _extract_evidence_numbers(text: str, limit: int = 8) -> List[str]:
    matches = re.findall(r"\b\d+[\d\.,%]*\b", text or "")
    seen = []
    for value in matches:
        if value not in seen:
            seen.append(value)
        if len(seen) >= limit:
            break
    return seen


def build_month_fallback_summary(item: dict, error_message: str) -> dict:
    text = item.get("text") or ""
    evidence = _extract_evidence_numbers(text, limit=8)
    return {
        "month": item["month"],
        "headline": f"Resumen automático limitado para {item['month']}",
        "kpis": [
            {
                "label": "Páginas extraídas",
                "value": str(item.get("page_count") or "s/d"),
                "note": "Resumen heurístico por indisponibilidad temporal del modelo",
            },
            {
                "label": "Caracteres extraídos",
                "value": str(len(text)),
                "note": "Base textual disponible para el informe",
            },
        ],
        "highlights": [
            "Se extrajo correctamente el PDF mensual.",
            "La generación IA falló por saturación temporal del modelo.",
        ],
        "alerts": [
            "Revisar nuevamente más tarde para enriquecer insights automáticos.",
        ],
        "top_contents": [],
        "evidence_numbers": evidence,
        "source_gaps": [error_message[:180]],
        "_meta_fallback": True,
    }


def build_period_fallback_deck(period: dict, monthly_summaries: List[dict], error_message: str) -> dict:
    hero_metrics = []
    for summary in monthly_summaries[:3]:
        evidence = summary.get("evidence_numbers") or []
        hero_metrics.append({
            "label": summary.get("month", "Mes"),
            "value": evidence[0] if evidence else "s/d",
            "note": summary.get("headline", "Resumen disponible"),
        })

    bullets_resumen = [summary.get("headline", f"Mes {summary.get('month', '')}") for summary in monthly_summaries]
    bullets_alertas = []
    for summary in monthly_summaries:
        bullets_alertas.extend(summary.get("alerts") or [])
    bullets_alertas = bullets_alertas[:4] or ["El modelo IA estuvo temporalmente no disponible."]

    bullets_reco = [
        "Reintentar la generación automática para obtener insights enriquecidos.",
        "Mantener el almacenamiento mensual de PDFs para consolidaciones futuras.",
    ]

    return {
        "email_subject": period.get("email_subject", period.get("title", "Informe CI")),
        "preheader": "Informe generado con fallback por indisponibilidad temporal del modelo.",
        "title": period.get("title", "Informe CI"),
        "subtitle": period.get("subtitle", ""),
        "hero_metrics": hero_metrics[:5],
        "slides": [
            {
                "title": "Resumen del período",
                "bullets": bullets_resumen[:5],
                "callout": "Salida de contingencia: el texto fuente fue procesado, pero la capa generativa estuvo saturada.",
            },
            {
                "title": "KPIs disponibles",
                "kpis": hero_metrics[:5],
                "callout": "Los valores se muestran tal como fueron detectados en la extracción textual.",
            },
            {
                "title": "Alertas operativas",
                "bullets": bullets_alertas,
                "callout": error_message[:180],
            },
            {
                "title": "Siguiente acción",
                "bullets": bullets_reco,
                "callout": "El pipeline no se interrumpe: se entrega una versión mínima y se puede regenerar luego.",
            },
        ],
        "footer_note": "Fuente: dashboards mensuales de CI. Informe emitido en modo contingencia.",
        "_meta_fallback": True,
    }


def summarize_month(client, item: dict, force_regenerate: bool = False) -> dict:
    MONTHLY_AI_DIR.mkdir(parents=True, exist_ok=True)
    output_path = MONTHLY_AI_DIR / f"{item['month']}.json"

    if output_path.exists() and not force_regenerate:
        return json.loads(output_path.read_text(encoding="utf-8"))

    payload = build_month_payload(item)
    try:
        summary = _call_gemini_for_json(client, MONTHLY_SUMMARY_PROMPT, payload)
    except Exception as exc:
        if not ALLOW_HEURISTIC_FALLBACK:
            raise
        print(f"Usando fallback heurístico para {item['month']}: {exc}")
        summary = build_month_fallback_summary(item, str(exc))
    output_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def build_period_input(period: dict, corpus_files: List[dict], monthly_summaries: List[dict]) -> dict:
    available_months = {item["month"] for item in corpus_files}
    return {
        "period": {
            "kind": period["kind"],
            "year": period["year"],
            "quarter": period.get("quarter"),
            "label": period["label"],
            "title": period["title"],
            "subtitle": period["subtitle"],
            "months_requested": period["months"],
            "months_available": sorted(available_months),
        },
        "monthly_summaries": monthly_summaries,
    }


def generate_deck(client, period: dict, monthly_summaries: List[dict]) -> dict:
    payload = build_period_input(period, [{"month": item["month"]} for item in monthly_summaries], monthly_summaries)
    deck = _call_gemini_for_json(client, FINAL_DECK_PROMPT, payload)
    if not deck.get("title"):
        deck["title"] = period["title"]
    if not deck.get("subtitle"):
        deck["subtitle"] = period["subtitle"]
    if not deck.get("email_subject"):
        deck["email_subject"] = period["email_subject"]
    return deck


def render_kpi_cards(kpis: List[dict]) -> str:
    cards = []
    for kpi in kpis:
        label = escape(str(kpi.get("label", "")))
        value = escape(str(kpi.get("value", "")))
        note = escape(str(kpi.get("note", "")))
        cards.append(
            f"""
            <td style=\"padding:8px; vertical-align:top;\">
              <div style=\"background:#F6F8FB; border:1px solid #D9E2F2; border-radius:16px; padding:16px; min-height:120px;\">
                <div style=\"font-size:12px; text-transform:uppercase; letter-spacing:0.3px; color:#5B6F8C; margin-bottom:8px;\">{label}</div>
                <div style=\"font-size:26px; font-weight:700; color:#072146; margin-bottom:8px; line-height:1.1;\">{value}</div>
                <div style=\"font-size:13px; color:#4B5E7A; line-height:1.4;\">{note}</div>
              </div>
            </td>
            """.strip()
        )
    return "<tr>" + "".join(cards) + "</tr>"


def render_bullets(items: List[str]) -> str:
    if not items:
        return ""
    bullet_items = "".join(
        f'<li style="margin:0 0 10px 0;">{escape(str(item))}</li>' for item in items if str(item).strip()
    )
    return f'<ul style="padding-left:20px; margin:12px 0 0 0; color:#072146; font-size:15px; line-height:1.5;">{bullet_items}</ul>'


def render_slide(slide: dict, index: int) -> str:
    title = escape(str(slide.get("title", f"Slide {index}")))
    bullets_html = render_bullets(slide.get("bullets") or [])
    kpis = slide.get("kpis") or []
    callout = escape(str(slide.get("callout", "")))

    kpis_html = ""
    if kpis:
        kpis_html = f"""
        <table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"margin-top:14px;\">
          {render_kpi_cards(kpis)}
        </table>
        """

    callout_html = ""
    if callout:
        callout_html = f"""
        <div style=\"margin-top:16px; background:#EAF3FF; border-left:4px solid #1973B8; border-radius:12px; padding:14px 16px; color:#072146; font-size:14px; line-height:1.5;\">
          {callout}
        </div>
        """

    return f"""
    <section style=\"margin:0 0 24px 0; background:#FFFFFF; border:1px solid #DCE6F2; border-radius:22px; overflow:hidden;\">
      <div style=\"padding:18px 22px; background:#F3F7FB; border-bottom:1px solid #DCE6F2;\">
        <div style=\"font-size:12px; text-transform:uppercase; color:#5B6F8C; letter-spacing:0.4px;\">Slide {index}</div>
        <div style=\"font-size:24px; font-weight:700; color:#072146; margin-top:4px;\">{title}</div>
      </div>
      <div style=\"padding:22px;\">
        {bullets_html}
        {kpis_html}
        {callout_html}
      </div>
    </section>
    """.strip()


def render_html(period: dict, deck: dict) -> str:
    hero_metrics = deck.get("hero_metrics") or []
    slides = deck.get("slides") or []
    hero_html = ""
    if hero_metrics:
        hero_html = f"""
        <table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"margin-top:18px;\">
          {render_kpi_cards(hero_metrics)}
        </table>
        """

    slides_html = "\n".join(render_slide(slide, index + 1) for index, slide in enumerate(slides))
    footer_note = escape(str(deck.get("footer_note", "Fuente: dashboards mensuales de CI.")))
    preheader = escape(str(deck.get("preheader", "Resumen visual del período.")))
    title = escape(str(deck.get("title", period["title"])))
    subtitle = escape(str(deck.get("subtitle", period["subtitle"])))

    return f"""
<!doctype html>
<html lang=\"es\">
  <head>
    <meta charset=\"utf-8\" />
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />
    <title>{title}</title>
  </head>
  <body style=\"margin:0; padding:0; background:#EEF3F8; font-family:Arial, Helvetica, sans-serif;\">
    <div style=\"display:none; max-height:0; overflow:hidden; opacity:0;\">{preheader}</div>
    <table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"background:#EEF3F8;\">
      <tr>
        <td align=\"center\" style=\"padding:28px 14px;\">
          <table role=\"presentation\" width=\"960\" cellspacing=\"0\" cellpadding=\"0\" style=\"max-width:960px; width:100%;\">
            <tr>
              <td style=\"background:linear-gradient(135deg, #072146 0%, #0D4D8B 100%); padding:28px 30px; border-radius:24px; color:#FFFFFF;\">
                <div style=\"font-size:13px; letter-spacing:0.5px; text-transform:uppercase; opacity:0.9;\">Comunicaciones Internas</div>
                <div style=\"font-size:34px; line-height:1.15; font-weight:700; margin-top:8px;\">{title}</div>
                <div style=\"font-size:16px; line-height:1.4; margin-top:10px; opacity:0.95;\">{subtitle}</div>
                {hero_html}
              </td>
            </tr>
            <tr>
              <td style=\"padding-top:24px;\">{slides_html}</td>
            </tr>
            <tr>
              <td style=\"padding:12px 6px 0 6px; color:#5B6F8C; font-size:12px; line-height:1.5;\">{footer_note}</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
    """.strip()


def render_text(period: dict, deck: dict) -> str:
    lines = [
        deck.get("title", period["title"]),
        deck.get("subtitle", period["subtitle"]),
        "",
        "KPIs destacados:",
    ]

    for item in deck.get("hero_metrics") or []:
        lines.append(f"- {item.get('label')}: {item.get('value')} ({item.get('note')})")

    for slide in deck.get("slides") or []:
        lines.append("")
        lines.append(slide.get("title", "Slide"))
        for bullet in slide.get("bullets") or []:
            lines.append(f"- {bullet}")
        for kpi in slide.get("kpis") or []:
            lines.append(f"- {kpi.get('label')}: {kpi.get('value')} ({kpi.get('note')})")
        if slide.get("callout"):
            lines.append(f"Nota: {slide['callout']}")

    footer_note = deck.get("footer_note")
    if footer_note:
        lines.extend(["", footer_note])

    return "\n".join(str(line) for line in lines if line is not None)


def save_outputs(period: dict, deck: dict, html_report: str, text_report: str) -> dict:
    report_dir = REPORTS_DIR / period["slug"]
    report_dir.mkdir(parents=True, exist_ok=True)

    json_path = report_dir / "report.json"
    html_path = report_dir / "report.html"
    text_path = report_dir / "report.txt"

    json_path.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
    html_path.write_text(html_report, encoding="utf-8")
    text_path.write_text(text_report, encoding="utf-8")

    return {
        "report_dir": str(report_dir),
        "json_path": str(json_path),
        "html_path": str(html_path),
        "text_path": str(text_path),
    }


def generate_period_report(period_slug: str) -> dict:
    corpus = load_corpus()
    period = load_period_by_slug(period_slug)

    months_needed = set(period["months"])
    selected_files = [item for item in corpus.get("files", []) if item["month"] in months_needed]
    selected_files = sorted(selected_files, key=lambda item: item["month"])

    if not selected_files:
        raise RuntimeError(f"No hay textos disponibles para el período {period_slug}")

    client = get_client()
    force_regenerate = (os.environ.get("FORCE_REGENERATE_AI") or "false").lower() == "true"

    monthly_summaries = []
    for item in selected_files:
        summary = summarize_month(client, item, force_regenerate=force_regenerate)
        monthly_summaries.append(summary)
        print(f"Resumen IA generado: {item['month']}")

    try:
        deck = generate_deck(client, period, monthly_summaries)
    except Exception as exc:
        if not ALLOW_HEURISTIC_FALLBACK:
            raise
        print(f"Usando fallback de deck para {period_slug}: {exc}")
        deck = build_period_fallback_deck(period, monthly_summaries, str(exc))
    html_report = render_html(period, deck)
    text_report = render_text(period, deck)
    paths = save_outputs(period, deck, html_report, text_report)

    result = {
        "period_slug": period_slug,
        "period": period,
        "report": deck,
        "monthly_summaries": monthly_summaries,
        **paths,
    }
    return result


def main() -> None:
    period_slug = os.environ.get("TARGET_PERIOD_SLUG")
    if not period_slug:
        raise ValueError("Falta TARGET_PERIOD_SLUG")

    result = generate_period_report(period_slug)
    print(
        json.dumps(
            {
                "period_slug": result["period_slug"],
                "report_dir": result["report_dir"],
                "html_path": result["html_path"],
                "text_path": result["text_path"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
