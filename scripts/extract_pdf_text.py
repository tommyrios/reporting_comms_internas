from __future__ import annotations

import json
from pathlib import Path
from typing import List

from pypdf import PdfReader

DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
TEXTS_DIR = OUTPUT_DIR / "monthly_texts"
MANIFEST_PATH = DATA_DIR / "monthly_pdf_manifest.json"
CORPUS_PATH = OUTPUT_DIR / "monthly_corpus.json"
COMBINED_TEXT_PATH = OUTPUT_DIR / "monthly_corpus.txt"


def extract_text_from_pdf(pdf_path: Path) -> dict:
    reader = PdfReader(str(pdf_path))
    pages: List[str] = []

    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        pages.append(f"\n\n===== PAGINA {page_number} =====\n{text}")

    full_text = "".join(pages).strip()
    return {
        "page_count": len(reader.pages),
        "text_chars": len(full_text),
        "text": full_text,
    }


def main() -> None:
    if not MANIFEST_PATH.exists():
        raise FileNotFoundError(f"No existe {MANIFEST_PATH}")

    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    files = manifest.get("files", [])

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEXTS_DIR.mkdir(parents=True, exist_ok=True)

    if not files:
        empty_payload = {
            "status": manifest.get("status", "empty"),
            "files": [],
            "combined_text_path": str(COMBINED_TEXT_PATH),
        }
        CORPUS_PATH.write_text(json.dumps(empty_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        COMBINED_TEXT_PATH.write_text("", encoding="utf-8")
        print(json.dumps(empty_payload, ensure_ascii=False, indent=2))
        return

    corpus_files = []
    combined_chunks: List[str] = []

    for item in files:
        month = item["month"]
        pdf_path = Path(item["pdf_path"])
        if not pdf_path.exists():
            raise FileNotFoundError(f"No existe el PDF descargado: {pdf_path}")

        extracted = extract_text_from_pdf(pdf_path)
        text_path = TEXTS_DIR / f"{month}.txt"
        text_path.write_text(extracted["text"], encoding="utf-8")

        corpus_item = {
            **item,
            "text_path": str(text_path),
            "page_count": extracted["page_count"],
            "text_chars": extracted["text_chars"],
            "text": extracted["text"],
        }
        corpus_files.append(corpus_item)
        combined_chunks.append(
            f"\n\n##############################\nMES: {month}\nARCHIVO: {item['filename']}\nASUNTO: {item['subject']}\n##############################\n{extracted['text']}"
        )

        print(f"Texto extraído: {month} ({extracted['page_count']} páginas)")

    combined_text = "".join(combined_chunks).strip()
    COMBINED_TEXT_PATH.write_text(combined_text, encoding="utf-8")

    payload = {
        "status": manifest.get("status", "ok"),
        "periods": manifest.get("periods", []),
        "months_requested": manifest.get("months_requested", []),
        "missing_months": manifest.get("missing_months", []),
        "files": corpus_files,
        "combined_text_path": str(COMBINED_TEXT_PATH),
        "combined_text_chars": len(combined_text),
    }

    CORPUS_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(json.dumps(
        {
            "status": payload["status"],
            "files": [
                {
                    "month": item["month"],
                    "page_count": item["page_count"],
                    "text_chars": item["text_chars"],
                }
                for item in corpus_files
            ],
            "combined_text_chars": payload["combined_text_chars"],
        },
        ensure_ascii=False,
        indent=2,
    ))


if __name__ == "__main__":
    main()
