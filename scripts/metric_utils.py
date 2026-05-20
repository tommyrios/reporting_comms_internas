from __future__ import annotations

import re
from typing import Any


CANON_TITLE_MAX_LENGTH = 180


def to_float_locale(value: Any, default: float = 0.0) -> float:
    """Parsea números exportados por Looker en formato ES/EN.

    Soporta miles con punto o coma (`11.785`, `5,580`) y decimales con
    coma o punto (`80,75%`, `80.75%`).
    """
    if value in (None, "", "-"):
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("%", "")
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text:
        return default

    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
    elif "." in text:
        parts = text.split(".")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            text = text.replace(".", "")

    try:
        return float(text)
    except Exception:
        return default


def parse_integer_value(value: Any) -> int | None:
    if value in (None, "", "-"):
        return None
    return int(round(to_float_locale(value, 0.0)))


def parse_percent_value(value: Any) -> float | None:
    if value in (None, "", "-"):
        return None

    text = str(value).strip()
    number = to_float_locale(text, 0.0)
    if "%" not in text and 0 < number <= 1:
        number *= 100
    return round(number, 2)


def normalize_percentage(value: Any) -> float:
    numeric = to_float_locale(value, 0.0)
    raw = str(value or "")
    if "%" in raw:
        return round(numeric, 2)
    if 0 <= numeric <= 1:
        return round(numeric * 100, 2)
    return round(numeric, 2)
