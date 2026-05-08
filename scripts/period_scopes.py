from __future__ import annotations

import re
import unicodedata
from typing import Iterable

REQUIRED_PERIOD_SCOPES = ["argentina", "holding", "combined"]

SCOPE_LABELS = {
    "argentina": "Argentina",
    "holding": "Holding",
    "combined": "Argentina + Holding",
}

SCOPE_FILE_TOKENS = {
    "argentina": "ARG",
    "holding": "HOLDING",
    "combined": "ARG_HOLDING",
}


def normalize_text(value: str | None) -> str:
    value = (value or "").strip().lower()
    normalized = unicodedata.normalize("NFD", value)
    without_accents = "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", without_accents).strip()


def compact_text(value: str | None) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize_text(value))


def infer_scope_from_text(*values: str | None) -> str | None:
    """Detecta el alcance del dashboard desde asunto/nombre de archivo.

    Orden intencional: primero combinado, luego Holding y por último Argentina.
    Si un texto dice "Argentina + Holding", no debe caer en Argentina por contener "arg".
    """
    haystack = normalize_text(" ".join(value or "" for value in values))
    compact = compact_text(haystack)

    combined_markers = [
        "argentinaholding",
        "argholding",
        "argxholding",
        "argyholding",
        "argmasholding",
        "argenthholding",
        "arg hol",
        "arg holding",
        "argentina holding",
        "argentina y holding",
        "argentina mas holding",
        "argentina holding combinado",
        "arg holding combinado",
        "combined",
        "consolidado",
    ]
    if any(marker.replace(" ", "") in compact for marker in combined_markers):
        return "combined"

    # HOL se usa en algunos filtros globales, pero se evita matchear palabras largas accidentales.
    if re.search(r"(?:^|\s)(holding|hol)(?:\s|$)", haystack):
        return "holding"

    if re.search(r"(?:^|\s)(argentina|arg)(?:\s|$)", haystack):
        return "argentina"

    return None


def period_scope_filename(period_slug: str, scope: str) -> str:
    token = SCOPE_FILE_TOKENS.get(scope, scope.upper())
    return f"{period_slug}_{token}.pdf"


def required_scopes_from_env(raw: str | None) -> list[str]:
    if not raw:
        return list(REQUIRED_PERIOD_SCOPES)
    values = [item.strip().lower() for item in raw.split(",") if item.strip()]
    invalid = [value for value in values if value not in REQUIRED_PERIOD_SCOPES]
    if invalid:
        raise ValueError(f"Scopes inválidos: {', '.join(invalid)}. Usá: {', '.join(REQUIRED_PERIOD_SCOPES)}")
    return values
