import json
import logging
import os
import time
from copy import deepcopy

from google import genai

from config import LEGACY_SUMMARIES_DIR, PDF_DIR, SUMMARIES_DIR, ensure_dir
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


def _build_local_fallback_summary(month_key: str, warning: str) -> dict:
    return {
        "month": month_key,
        "headline": f"Resumen operativo disponible para {month_key}",
        "summary_text": "No se pudo generar resumen con LLM; se creó una salida mínima para continuar el pipeline.",
        "generation_mode": "local_fallback",
        "warning": warning,
        "data": {
            "push_volume": 0,
            "pull_notes": 0,
            "pull_reads": 0,
            "push_opens_pct": 0,
            "push_interaction_pct": 0,
        },
        "insights": {
            "audience_segmentation": [],
            "strategic_axes": [],
            "internal_clients": [],
            "top_push_comm": {},
            "top_pull_note": {},
            "hitos_mes": "",
        },
        "sections": {},
    }


def _cache_candidates(month_key: str) -> list:
    return [
        ensure_dir(SUMMARIES_DIR) / f"{month_key}.json",
        ensure_dir(LEGACY_SUMMARIES_DIR) / f"{month_key}.json",
    ]


def _read_cached_summary(month_key: str):
    for cache_path in _cache_candidates(month_key):
        if cache_path.exists():
            return json.loads(cache_path.read_text(encoding="utf-8")), cache_path
    return None, None


def _persist_summary(summary: dict, month_key: str) -> None:
    summary_copy = deepcopy(summary)
    summary_copy["month"] = month_key
    primary_path = ensure_dir(SUMMARIES_DIR) / f"{month_key}.json"
    primary_path.write_text(json.dumps(summary_copy, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("event=summary_saved month=%s path=%s", month_key, primary_path)


def summarize_month(client: genai.Client, month_key: str, force_regenerate: bool = False) -> dict:
    cached_summary, cached_path = _read_cached_summary(month_key)
    if cached_summary and not force_regenerate:
        logger.info("event=summary_cache_hit month=%s path=%s", month_key, cached_path)
        return cached_summary

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
        summary["generation_mode"] = "llm"
        _persist_summary(summary, month_key)
        return summary
    except Exception as exc:
        if cached_summary:
            logger.warning(
                "event=summary_fallback_to_cache month=%s reason=%s cache_path=%s",
                month_key,
                exc,
                cached_path,
            )
            return cached_summary
        logger.warning("event=summary_local_fallback month=%s reason=%s", month_key, exc)
        fallback_summary = _build_local_fallback_summary(
            month_key,
            f"No se pudo generar resumen mensual con Gemini. Se usó modo local_fallback. Error: {str(exc)}",
        )
        _persist_summary(fallback_summary, month_key)
        return fallback_summary
    finally:
        try:
            if uploaded and getattr(uploaded, "name", None):
                client.files.delete(name=uploaded.name)
        except Exception:
            logger.warning("event=summary_cleanup_failed month=%s", month_key)
