import json
import logging
import os
import re
import time
from typing import Any

from google import genai

from config import PROMPTS_DIR

logger = logging.getLogger(__name__)


def load_prompt(filename: str) -> str:
    path = PROMPTS_DIR / filename
    if not path.exists():
        raise FileNotFoundError(path)
    return path.read_text(encoding="utf-8").strip()


def clean_json_response(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```json\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^```\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def build_genai_client() -> genai.Client:
    api_key = (os.environ.get("GEMINI_API_KEY") or "").strip()
    if not api_key:
        raise RuntimeError("Falta GEMINI_API_KEY")
    return genai.Client(api_key=api_key)


def _candidate_models() -> list[str]:
    primary = (os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash").strip()
    fallback_enabled = (os.environ.get("GEMINI_ENABLE_MODEL_FALLBACK") or "true").strip().lower() not in ("0", "false", "no")
    models = [primary]
    if fallback_enabled:
        fallbacks = [m.strip() for m in (os.environ.get("GEMINI_FALLBACK_MODELS") or "").split(",") if m.strip()]
        models.extend(fallbacks)
    deduped = []
    seen = set()
    for model in models:
        if model and model not in seen:
            deduped.append(model)
            seen.add(model)
    return deduped


def _is_retryable_error(exc: Exception) -> bool:
    text = str(exc).lower()
    retryable_markers = (
        "429",
        "500",
        "502",
        "503",
        "504",
        "timeout",
        "temporarily",
        "unavailable",
        "connection reset",
        "deadline exceeded",
        "rate limit",
        "connection",
    )
    return any(marker in text for marker in retryable_markers)


def call_gemini_for_json(client: genai.Client, contents: list[Any]) -> dict[str, Any]:
    max_retries = max(1, int((os.environ.get("GEMINI_MAX_RETRIES_PER_MODEL") or "3").strip()))
    initial_backoff = max(0.0, float((os.environ.get("GEMINI_INITIAL_BACKOFF_SECONDS") or "3").strip()))
    max_backoff = max(initial_backoff, float((os.environ.get("GEMINI_MAX_BACKOFF_SECONDS") or "30").strip()))
    models = _candidate_models()
    last_error: Exception | None = None

    for model_index, model_name in enumerate(models, start=1):
        backoff = initial_backoff
        for attempt in range(1, max_retries + 1):
            try:
                logger.info(
                    "event=gemini_model_attempt model=%s model_index=%s/%s attempt=%s/%s",
                    model_name,
                    model_index,
                    len(models),
                    attempt,
                    max_retries,
                )
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config={"response_mime_type": "application/json"},
                )
                text = getattr(response, "text", "") or ""
                logger.info("event=gemini_model_success model=%s attempt=%s", model_name, attempt)
                return json.loads(clean_json_response(text))
            except Exception as exc:
                last_error = exc
                if not _is_retryable_error(exc):
                    logger.error(
                        "event=gemini_non_retryable_error model=%s attempt=%s reason=%s",
                        model_name,
                        attempt,
                        exc,
                    )
                    raise
                error_message = str(exc)
                logger.warning(
                    "event=gemini_retry model=%s attempt=%s/%s backoff_seconds=%s reason=%s",
                    model_name,
                    attempt,
                    max_retries,
                    backoff,
                    error_message,
                )
                if attempt < max_retries:
                    time.sleep(backoff)
                    backoff = min(backoff * 2, max_backoff)
                else:
                    logger.warning(
                        "event=gemini_model_exhausted model=%s retries=%s reason=%s",
                        model_name,
                        max_retries,
                        error_message,
                    )
                    if model_index < len(models):
                        logger.info(
                            "event=gemini_model_fallback from_model=%s to_model=%s",
                            model_name,
                            models[model_index],
                        )

    raise RuntimeError(
        f"Gemini agotó todos los modelos configurados ({', '.join(models)}). Último error: {str(last_error) if last_error else 'sin detalle'}"
    )
