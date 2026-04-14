import json
from pathlib import Path

from pypdf import PdfReader


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
PDF_DIR = DATA_DIR / "monthly_pdfs"
EXTRACTED_DIR = DATA_DIR / "extracted_text"
MANIFEST_PATH = DATA_DIR / "fetch_result.json"


def extract_text_from_pdf(pdf_path: Path) -> tuple[str, int]:
    reader = PdfReader(str(pdf_path))
    texts = []
    page_count = 0

    for page in reader.pages:
        page_count += 1
        page_text = page.extract_text() or ""
        texts.append(page_text)

    full_text = "\n\n".join(texts)
    return full_text, page_count


def main() -> dict:
    if not MANIFEST_PATH.exists():
        raise FileNotFoundError(f"No existe el manifest de fetch: {MANIFEST_PATH}")

    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    files = manifest.get("files", [])

    if not files:
        result = {
            "status": "skipped",
            "files": [],
            "combined_text_chars": 0,
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return result

    EXTRACTED_DIR.mkdir(parents=True, exist_ok=True)

    extracted_files = []
    combined_parts = []

    for item in files:
        month = item["month"]
        pdf_path = Path(item["pdf_path"])

        if not pdf_path.is_absolute():
            pdf_path = BASE_DIR / pdf_path

        if not pdf_path.exists():
            raise FileNotFoundError(f"No existe el PDF para {month}: {pdf_path}")

        text, page_count = extract_text_from_pdf(pdf_path)

        txt_path = EXTRACTED_DIR / f"{month}.txt"
        txt_path.write_text(text, encoding="utf-8")

        combined_parts.append(f"\n\n===== {month} =====\n\n{text}")

        extracted_files.append(
            {
                "month": month,
                "page_count": page_count,
                "text_chars": len(text),
                "txt_path": str(txt_path),
            }
        )

        print(f"Texto extraído: {month} ({page_count} páginas)")

    combined_text = "".join(combined_parts)
    combined_path = EXTRACTED_DIR / "combined.txt"
    combined_path.write_text(combined_text, encoding="utf-8")

    result = {
        "status": "ok",
        "files": extracted_files,
        "combined_text_chars": len(combined_text),
        "combined_path": str(combined_path),
    }

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


if __name__ == "__main__":
    main()