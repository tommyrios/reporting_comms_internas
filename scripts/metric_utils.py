from typing import Any


def to_float_locale(value: Any, default: float = 0.0) -> float:
    if value in (None, "", "-"):
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("%", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except Exception:
        filtered = "".join(ch for ch in str(value) if ch.isdigit() or ch in ".,-")
        if not filtered:
            return default
        filtered = filtered.replace(",", ".")
        try:
            return float(filtered)
        except Exception:
            return default


def normalize_percentage(value: Any) -> float:
    numeric = to_float_locale(value, 0.0)
    raw = str(value or "")
    if "%" in raw:
        return round(numeric, 2)
    if 0 <= numeric <= 1:
        return round(numeric * 100, 2)
    return round(numeric, 2)
