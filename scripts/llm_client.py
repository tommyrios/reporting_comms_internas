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
    fallbacks = [m.strip() for m in (os.environ.get("GEMINI_FALLBACK_MODELS") or "").split(",") if m.strip()]
    models = [primary] + fallbacks
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
        "timeout",
        "temporarily",
        "rate limit",
        "429",
        "500",
        "502",
        "503",
        "504",
        "connection",
        "unavailable",
    )
    return any(marker in text for marker in retryable_markers)


def call_gemini_for_json(client: genai.Client, contents: list[Any]) -> dict[str, Any]:
    max_retries = int((os.environ.get("GEMINI_MAX_RETRIES") or "6").strip())
    initial_backoff = float((os.environ.get("GEMINI_INITIAL_BACKOFF_SECONDS") or "5").strip())
    last_error: Exception | None = None

    for model_name in _candidate_models():
        backoff = initial_backoff
        for attempt in range(1, max_retries + 1):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config={"response_mime_type": "application/json"},
                )
                text = getattr(response, "text", "") or ""
                return json.loads(clean_json_response(text))
            except Exception as exc:
                last_error = exc
                if not _is_retryable_error(exc):
                    raise
                logger.warning(
                    "event=gemini_retry model=%s attempt=%s/%s backoff_seconds=%s reason=%s",
                    model_name,
                    attempt,
                    max_retries,
                    backoff,
                    exc,
                )
                if attempt < max_retries:
                    time.sleep(backoff)
                    backoff = min(backoff * 2, 60)

    raise RuntimeError(str(last_error) if last_error else "No se pudo obtener respuesta JSON de Gemini")
