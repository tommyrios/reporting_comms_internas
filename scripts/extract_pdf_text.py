import json
from pathlib import Path
from pypdf import PdfReader

DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")

PDF_PATH = DATA_DIR / "latest_dashboard.pdf"
TEXT_PATH = OUTPUT_DIR / "pdf_text.txt"
META_PATH = DATA_DIR / "metadata.json"


def main() -> None:
    if not PDF_PATH.exists():
        raise FileNotFoundError(f"No existe el PDF de entrada: {PDF_PATH}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(PDF_PATH))
    pages = []

    for i, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        pages.append(f"\n\n===== PAGINA {i} =====\n{text}")

    full_text = "".join(pages)
    TEXT_PATH.write_text(full_text, encoding="utf-8")

    meta = {}
    if META_PATH.exists():
        try:
            meta = json.loads(META_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            meta = {}

    meta["page_count"] = len(reader.pages)
    meta["text_chars"] = len(full_text)
    meta["pdf_path"] = str(PDF_PATH)
    meta["text_path"] = str(TEXT_PATH)

    META_PATH.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()