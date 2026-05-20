from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
LOCAL_DATA_DIR = BASE_DIR / "local_data"
OUTPUT_DIR = BASE_DIR / "output"
PDF_DIR = DATA_DIR / "period_pdfs"
INBOX_PDF_DIR = LOCAL_DATA_DIR / "inbox_pdfs"
REPORTS_DIR = OUTPUT_DIR / "reports"
SUMMARIES_DIR = DATA_DIR / "period_summaries"
LEGACY_SUMMARIES_DIR = OUTPUT_DIR / "period_summaries"
RAW_EXTRACTED_DIR = DATA_DIR / "raw_extracted"
CANONICAL_PERIOD_DIR = DATA_DIR / "canonical_periods"
CANONICAL_MONTHLY_DIR = CANONICAL_PERIOD_DIR  # compat interna del extractor determinístico
VALIDATION_DIR = DATA_DIR / "validation"
MANUAL_CONTEXT_DIR = DATA_DIR / "manual_context"
ASSETS_DIR = BASE_DIR / "assets"


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path
