import os
from pathlib import Path
from openai import OpenAI

INPUT_PATH = Path("output/pdf_text.txt")
OUTPUT_DIR = Path("output")
TEXT_REPORT_PATH = OUTPUT_DIR / "report.txt"
HTML_REPORT_PATH = OUTPUT_DIR / "report.html"

MODEL = os.environ.get("OPENAI_MODEL", "gpt-5-mini")

PROMPT = """
Sos un analista de datos de comunicaciones internas de BBVA Argentina.
Tu tarea es redactar un reporte ejecutivo claro, breve y accionable a partir
del texto extraído de un dashboard PDF.

Quiero que generes:
1. Un resumen ejecutivo de no más de 15 líneas.
2. Un bloque de KPIs clave en bullets.
3. 3 insights relevantes.
4. 2 alertas o limitaciones del análisis.
5. 2 recomendaciones accionables.
6. Un cierre breve.

Reglas:
- Escribí en español.
- Usá tono ejecutivo, claro y profesional.
- No inventes datos.
- Si hay datos confusos o ambiguos, aclaralo.
- Priorizá métricas como envíos, aperturas, clics, tasas y tendencias.
- Si detectás una evolución mensual, mencionarla.
- Si hay rankings o tablas de mails, destacá mejores y peores desempeños si se ven con claridad.

Devolvé la respuesta en dos secciones exactamente así:

===TEXT===
[versión texto plano]

===HTML===
[versión HTML simple, prolija, con h2, ul, li y p]
""".strip()


def load_pdf_text() -> str:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {INPUT_PATH}")
    return INPUT_PATH.read_text(encoding="utf-8").strip()


def call_openai(pdf_text: str) -> str:
    client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

    response = client.responses.create(
        model=MODEL,
        input=[
            {
                "role": "system",
                "content": [
                    {
                        "type": "input_text",
                        "text": PROMPT,
                    }
                ],
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": f"Texto extraído del dashboard:\n\n{pdf_text}",
                    }
                ],
            },
        ],
    )

    return response.output_text.strip()


def split_outputs(content: str) -> tuple[str, str]:
    if "===TEXT===" not in content or "===HTML===" not in content:
        raise ValueError("La respuesta del modelo no vino en el formato esperado.")

    text_part = content.split("===TEXT===", 1)[1].split("===HTML===", 1)[0].strip()
    html_part = content.split("===HTML===", 1)[1].strip()

    return text_part, html_part


def save_outputs(text_report: str, html_report: str) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEXT_REPORT_PATH.write_text(text_report, encoding="utf-8")
    HTML_REPORT_PATH.write_text(html_report, encoding="utf-8")


def main():
    pdf_text = load_pdf_text()
    model_output = call_openai(pdf_text)
    text_report, html_report = split_outputs(model_output)
    save_outputs(text_report, html_report)
    print("Reporte generado correctamente.")


if __name__ == "__main__":
    main()