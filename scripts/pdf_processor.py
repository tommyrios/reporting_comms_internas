import json
import logging
import os
import time

from google import genai

from config import PDF_DIR, SUMMARIES_DIR, ensure_dir
from llm_client import call_gemini_for_json, load_prompt

logger = logging.getLogger(__name__)


def _is_retryable_error(exc: Exception) -> bool:
    text = str(exc).lower()
    retryable_markers = ("timeout", "temporarily", "rate limit", "429", "500", "502", "503", "504", "connection")
    return any(marker in text for marker in retryable_markers)


def _upload_with_retries(client: genai.Client, pdf_path: str, month_key: str):
    max_attempts = max(1, int((os.environ.get("GEMINI_UPLOAD_RETRIES") or "3").strip()))
    backoff = float((os.environ.get("GEMINI_UPLOAD_RETRY_BACKOFF_SECONDS") or "2").strip())

    for attempt in range(1, max_attempts + 1):
        try:
            return client.files.upload(file=pdf_path)
        except Exception as exc:
            if attempt >= max_attempts or not _is_retryable_error(exc):
                raise
            sleep_for = min(backoff * attempt, 30)
            logger.warning(
                "event=gemini_upload_retry month=%s attempt=%s sleep=%s reason=%s",
                month_key,
                attempt,
                sleep_for,
                exc,
            )
            time.sleep(sleep_for)


def _wait_until_processed(client: genai.Client, uploaded, month_key: str):
    timeout_seconds = max(1, int((os.environ.get("GEMINI_UPLOAD_PROCESS_TIMEOUT_SECONDS") or "300").strip()))
    poll_seconds = max(1, int((os.environ.get("GEMINI_UPLOAD_POLL_SECONDS") or "2").strip()))
    deadline = time.monotonic() + timeout_seconds

    while uploaded.state.name == "PROCESSING":
        if time.monotonic() >= deadline:
            raise TimeoutError(f"Timeout procesando PDF en Gemini para {month_key}")
        time.sleep(poll_seconds)
        uploaded = client.files.get(name=uploaded.name)
    return uploaded


def summarize_month(client: genai.Client, month_key: str, force_regenerate: bool = False) -> dict:
    summaries_dir = ensure_dir(SUMMARIES_DIR)
    path = summaries_dir / f"{month_key}.json"
    if path.exists() and not force_regenerate:
        return json.loads(path.read_text(encoding="utf-8"))

    pdf_path = PDF_DIR / f"{month_key}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"No existe el PDF mensual esperado: {pdf_path}")

    logger.info("event=summary_start month=%s pdf=%s", month_key, pdf_path.name)
    uploaded = _upload_with_retries(client, str(pdf_path), month_key)
    prompt_text = load_prompt("monthly_summary.txt")

    try:
        uploaded = _wait_until_processed(client, uploaded, month_key)

        if uploaded.state.name == "FAILED":
            raise RuntimeError(f"Gemini no pudo procesar el PDF {pdf_path.name}")

        summary = call_gemini_for_json(client, [uploaded, prompt_text])
        summary["month"] = month_key
        path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("event=summary_saved month=%s path=%s", month_key, path)
        return summary
    finally:
        try:
            if uploaded and getattr(uploaded, "name", None):
                client.files.delete(name=uploaded.name)
        except Exception:
            logger.warning("event=summary_cleanup_failed month=%s", month_key)
