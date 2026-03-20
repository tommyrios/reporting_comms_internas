import json
from pathlib import Path
from pypdf import PdfReader

DATA_DIR = Path('data')
PDF_PATH = DATA_DIR / 'latest_dashboard.pdf'
TEXT_PATH = DATA_DIR / 'pdf_text.txt'
META_PATH = DATA_DIR / 'metadata.json'


def main() -> None:
    reader = PdfReader(str(PDF_PATH))
    pages = []
    for i, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ''
        pages.append(f'\n\n===== PAGINA {i} =====\n{text}')

    full_text = ''.join(pages)
    TEXT_PATH.write_text(full_text, encoding='utf-8')

    meta = json.loads(META_PATH.read_text(encoding='utf-8'))
    meta['page_count'] = len(reader.pages)
    meta['text_chars'] = len(full_text)
    META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
