from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime, timezone
from html import escape
from pathlib import Path
from typing import List

import sys

SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.append(str(SCRIPT_DIR))

from reporting_periods import load_schedule

OUTPUT_DIR = BASE_DIR / "output"
CORPUS_PATH = OUTPUT_DIR / "monthly_corpus.json"
MONTHLY_AI_DIR = OUTPUT_DIR / "monthly_ai_summaries"
REPORTS_DIR = OUTPUT_DIR / "reports"

DEFAULT_MODEL = (os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash").strip()
FALLBACK_MODELS = [
    item.strip()
    for item in (os.environ.get("GEMINI_FALLBACK_MODELS") or "gemini-2.5-flash-lite,gemini-2.0-flash").split(",")
    if item.strip()
]
MAX_RETRIES = int((os.environ.get("GEMINI_MAX_RETRIES") or "4").strip())
INITIAL_BACKOFF_SECONDS = float((os.environ.get("GEMINI_INITIAL_BACKOFF_SECONDS") or "3").strip())
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
    api_key = (os.environ.get("GEMINI_API_KEY") or "").strip()
    if not api_key:
        raise ValueError("Falta GEMINI_API_KEY")

    from google import genai

    return genai.Client(api_key=api_key)


def _clean_json_string(text: str) -> str:
    cleaned = (text or "").strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()
    return cleaned


def _is_retryable_model_error(exc: Exception) -> bool:
    message = str(exc).lower()
    retry_tokens = ["503", "unavailable", "deadline", "timeout", "temporarily", "resource exhausted", "429"]
    return any(token in message for token in retry_tokens)


def _call_gemini_for_json(client, prompt: str, payload: dict) -> dict:
    from google.genai import types

    models_to_try = [DEFAULT_MODEL] + [item for item in FALLBACK_MODELS if item != DEFAULT_MODEL]
    last_error: Exception | None = None

    for model_name in models_to_try:
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=f"{prompt}\n\nINPUT_JSON:\n{json.dumps(payload, ensure_ascii=False)}",
                    config=types.GenerateContentConfig(response_mime_type="application/json"),
                )
                raw_text = _clean_json_string(response.text or "")
                return json.loads(raw_text)
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                if _is_retryable_model_error(exc) and attempt < MAX_RETRIES:
                    backoff = INITIAL_BACKOFF_SECONDS * (2 ** (attempt - 1))
                    print(
                        f"Gemini saturado con {model_name}. Reintento {attempt}/{MAX_RETRIES} en {backoff:.1f}s"
                    )
                    time.sleep(backoff)
                    continue

                print(f"Fallo Gemini con modelo {model_name} en intento {attempt}: {exc}")
                break

    if last_error is None:
        raise RuntimeError("No se pudo invocar Gemini y no se recibió detalle del error")
    raise last_error


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
    max_chars = int((os.environ.get("MONTH_TEXT_MAX_CHARS") or "70000").strip())
    return {
        "month": item["month"],
        "filename": item.get("filename"),
        "subject": item.get("subject"),
        "page_count": item.get("page_count"),
        "text": (item.get("text") or "")[:max_chars],
    }


def summarize_month(client, item: dict, force_regenerate: bool = False) -> dict:
    MONTHLY_AI_DIR.mkdir(parents=True, exist_ok=True)
    output_path = MONTHLY_AI_DIR / f"{item['month']}.json"

    if output_path.exists() and not force_regenerate:
        return json.loads(output_path.read_text(encoding="utf-8"))

    payload = build_month_payload(item)
    summary = _call_gemini_for_json(client, MONTHLY_SUMMARY_PROMPT, payload)
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


def _extract_candidate_numbers(text: str, limit: int = 8) -> List[str]:
    seen: List[str] = []
    for match in re.findall(r"\b\d[\d\.,%]*\b", text or ""):
        value = match.strip()
        if len(value) < 2:
            continue
        if value not in seen:
            seen.append(value)
        if len(seen) >= limit:
            break
    return seen


def _fallback_month_summary(item: dict) -> dict:
    text = item.get("text") or ""
    numbers = _extract_candidate_numbers(text, limit=8)
    highlights = [
        f"Se procesó el dashboard del mes {item['month']}",
        f"Texto extraído: {item.get('text_chars', 0)} caracteres en {item.get('page_count', 0)} páginas",
    ]
    if numbers:
        highlights.append(f"Números visibles en fuente: {', '.join(numbers[:4])}")

    return {
        "month": item["month"],
        "headline": f"Resumen de contingencia {item['month']}",
        "kpis": [
            {"label": "Páginas", "value": str(item.get("page_count", "-")), "note": "PDF procesado"},
            {"label": "Texto extraído", "value": str(item.get("text_chars", "-")), "note": "Caracteres recuperados"},
        ],
        "highlights": highlights[:5],
        "alerts": ["Gemini no estuvo disponible; se usó resumen heurístico."],
        "top_contents": [],
        "evidence_numbers": numbers,
        "source_gaps": ["Resumen generado sin interpretación del modelo."] if numbers else ["No se detectaron suficientes números confiables en el texto."],
    }


def _fallback_deck(period: dict, monthly_summaries: List[dict], error: str) -> dict:
    hero_metrics = []
    months_list = ", ".join(period["months"])
    hero_metrics.append({"label": "Meses", "value": str(len(monthly_summaries)), "note": months_list})
    hero_metrics.append({"label": "Modo", "value": "Contingencia", "note": "Gemini no disponible"})

    evidence_pool: List[str] = []
    for item in monthly_summaries:
        for value in item.get("evidence_numbers") or []:
            if value not in evidence_pool:
                evidence_pool.append(value)
    if evidence_pool:
        hero_metrics.append({"label": "Datos visibles", "value": evidence_pool[0], "note": "Primer dato detectado en fuente"})

    highlights = []
    alerts = []
    for item in monthly_summaries:
        highlights.extend(item.get("highlights") or [])
        alerts.extend(item.get("alerts") or [])

    slides = [
        {
            "title": "Resumen ejecutivo",
            "bullets": (highlights or ["Se procesó el material del período sin análisis del modelo."])[:5],
            "callout": "Salida mínima para no cortar el pipeline.",
        },
        {
            "title": "KPIs del período",
            "kpis": hero_metrics[:5],
            "callout": "Validar manualmente los números críticos antes de reenviar.",
        },
        {
            "title": "Alertas",
            "bullets": (alerts or ["Gemini respondió con alta demanda o indisponibilidad."])[:5],
            "callout": str(error)[:220],
        },
        {
            "title": "Recomendación operativa",
            "bullets": [
                "Reintentar el workflow más tarde.",
                "Conservar el PDF mensual para histórico.",
                "Usar esta versión solo como contingencia interna.",
            ],
            "callout": "El pipeline quedó estable aun con caída del modelo.",
        },
    ]

    return {
        "email_subject": period.get("email_subject") or f"Informe CI | {period['label']}",
        "preheader": "Versión de contingencia generada sin IA por indisponibilidad temporal del modelo.",
        "title": period["title"],
        "subtitle": f"{period['subtitle']} | Modo contingencia",
        "hero_metrics": hero_metrics[:5],
        "slides": slides,
        "footer_note": "Fuente: dashboards mensuales de CI. Salida en modo contingencia.",
        "generation_mode": "fallback",
    }


def save_outputs(period: dict, deck: dict, html_report: str, text_report: str) -> dict:
    report_dir = REPORTS_DIR / period["slug"]
    report_dir.mkdir(parents=True, exist_ok=True)

    json_path = report_dir / "report.json"
    html_path = report_dir / "report.html"
    text_path = report_dir / "report.txt"
    metadata_path = report_dir / "metadata.json"

    metadata = {
        "period_slug": period["slug"],
        "title": deck.get("title", period["title"]),
        "subtitle": deck.get("subtitle", period["subtitle"]),
        "email_subject": deck.get("email_subject", period.get("email_subject", f"Informe CI | {period['slug']}")),
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "generation_mode": deck.get("generation_mode", "ai"),
        "months": period.get("months", []),
    }

    json_path.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
    html_path.write_text(html_report, encoding="utf-8")
    text_path.write_text(text_report, encoding="utf-8")
    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "report_dir": str(report_dir),
        "json_path": str(json_path),
        "html_path": str(html_path),
        "text_path": str(text_path),
        "metadata_path": str(metadata_path),
    }


def generate_period_report(period_slug: str) -> dict:
    corpus = load_corpus()
    period = load_period_by_slug(period_slug)

    months_needed = set(period["months"])
    selected_files = [item for item in corpus.get("files", []) if item["month"] in months_needed]
    selected_files = sorted(selected_files, key=lambda item: item["month"])

    if not selected_files:
        raise RuntimeError(f"No hay textos disponibles para el período {period_slug}")

    force_regenerate = (os.environ.get("FORCE_REGENERATE_AI") or "false").lower() == "true"
    monthly_summaries: List[dict] = []
    deck: dict | None = None
    client = None
    generation_mode = "ai"
    deck_error = None

    try:
        client = get_client()
        for item in selected_files:
            summary = summarize_month(client, item, force_regenerate=force_regenerate)
            monthly_summaries.append(summary)
            print(f"Resumen IA generado: {item['month']}")

        deck = generate_deck(client, period, monthly_summaries)
        deck["generation_mode"] = "ai"
    except Exception as exc:  # noqa: BLE001
        deck_error = exc
        if not ALLOW_HEURISTIC_FALLBACK:
            raise

        print(f"Fallo generación IA para {period_slug}. Se usa fallback local: {exc}")
        generation_mode = "fallback"
        if not monthly_summaries:
            monthly_summaries = [_fallback_month_summary(item) for item in selected_files]
        deck = _fallback_deck(period, monthly_summaries, str(exc))

    assert deck is not None
    deck["generation_mode"] = generation_mode
    html_report = render_html(period, deck)
    text_report = render_text(period, deck)
    paths = save_outputs(period, deck, html_report, text_report)

    result = {
        "period_slug": period_slug,
        "period": period,
        "report": deck,
        "monthly_summaries": monthly_summaries,
        "generation_mode": generation_mode,
        **paths,
    }
    if deck_error is not None:
        result["warning"] = str(deck_error)
    return result


def main() -> None:
    period_slug = (os.environ.get("TARGET_PERIOD_SLUG") or "").strip()
    if not period_slug:
        raise ValueError("Falta TARGET_PERIOD_SLUG")

    result = generate_period_report(period_slug)
    print(
        json.dumps(
            {
                "period_slug": result["period_slug"],
                "generation_mode": result.get("generation_mode"),
                "report_dir": result["report_dir"],
                "html_path": result["html_path"],
                "text_path": result["text_path"],
                "metadata_path": result["metadata_path"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
