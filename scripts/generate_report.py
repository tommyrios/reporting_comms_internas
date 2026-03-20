import os
from pathlib import Path
from google import genai
from google.genai import types

INPUT_PATH = Path("output/pdf_text.txt")
OUTPUT_DIR = Path("output")
TEXT_REPORT_PATH = OUTPUT_DIR / "report.txt"
HTML_REPORT_PATH = OUTPUT_DIR / "report.html"


PROMPT = """
Sos un analista de datos de comunicaciones internas de BBVA Argentina.

Tu tarea es redactar un reporte ejecutivo claro, breve y accionable a partir
del texto extraído de un dashboard PDF.

Quiero que generes:

1. Un resumen ejecutivo (máx 10 líneas)
2. KPIs clave (bullets)
3. 3 insights relevantes
4. 2 alertas
5. 2 recomendaciones accionables

Reglas:
- Español
- Tono ejecutivo
- No inventar datos
- Priorizar métricas: envíos, aperturas, clics, tasas
- Si hay tendencias, mencionarlas

Formato EXACTO de salida:

===TEXT===
[texto plano]

===HTML===
[html simple con h2, ul, li, p]
""".strip()


def load_pdf_text() -> str:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"No existe: {INPUT_PATH}")
    return INPUT_PATH.read_text(encoding="utf-8").strip()


def call_gemini(pdf_text: str) -> str:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("Falta GEMINI_API_KEY")

    client = genai.Client(api_key=api_key)

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=f"{PROMPT}\n\nTexto:\n{pdf_text[:120000]}",
        config=types.GenerateContentConfig(
            response_mime_type="text/plain"
        ),
    )

    return response.text.strip()


def split_outputs(content: str):
    if "===TEXT===" not in content or "===HTML===" not in content:
        raise ValueError("Formato incorrecto del modelo")

    text_part = content.split("===TEXT===", 1)[1].split("===HTML===", 1)[0].strip()
    html_part = content.split("===HTML===", 1)[1].strip()

    return text_part, html_part


def save_outputs(text_report: str, html_report: str):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEXT_REPORT_PATH.write_text(text_report, encoding="utf-8")
    HTML_REPORT_PATH.write_text(html_report, encoding="utf-8")


def main():
    pdf_text = load_pdf_text()
    model_output = call_gemini(pdf_text)
    text_report, html_report = split_outputs(model_output)
    save_outputs(text_report, html_report)
    print("Reporte generado con Gemini")


if __name__ == "__main__":
    main()
