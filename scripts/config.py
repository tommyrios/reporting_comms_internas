from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
PDF_DIR = DATA_DIR / "monthly_pdfs"
REPORTS_DIR = OUTPUT_DIR / "reports"
PROMPTS_DIR = BASE_DIR / "prompts"
SUMMARIES_DIR = DATA_DIR / "monthly_summaries"
LEGACY_SUMMARIES_DIR = OUTPUT_DIR / "monthly_summaries"
RAW_EXTRACTED_DIR = DATA_DIR / "raw_extracted"
CANONICAL_MONTHLY_DIR = DATA_DIR / "canonical_monthly"
VALIDATION_DIR = DATA_DIR / "validation"
MANUAL_CONTEXT_DIR = DATA_DIR / "manual_context"
ASSETS_DIR = BASE_DIR / "assets"


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path
