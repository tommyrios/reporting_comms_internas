import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
PDF_DIR = DATA_DIR / "monthly_pdfs"
REPORTS_DIR = OUTPUT_DIR / "reports"
PROMPTS_DIR = BASE_DIR / "prompts"
SUMMARIES_DIR = OUTPUT_DIR / "monthly_summaries"

def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path