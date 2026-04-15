import json
import os
import re
import time
from typing import Any
from google import genai
from config import PROMPTS_DIR

def load_prompt(filename: str) -> str:
    path = PROMPTS_DIR / filename
    if not path.exists():
        raise FileNotFoundError(path)
    return path.read_text(encoding="utf-8").strip()

def clean_json_response(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```json\s*", "", text, flags=re.IGNORECASE)
    return re.sub(r"\s*```$", "", re.sub(r"^```\s*", "", text)).strip()

def build_genai_client() -> genai.Client:
    return genai.Client(api_key=(os.environ.get("GEMINI_API_KEY") or "").strip())

def call_gemini_for_json(client: genai.Client, contents: list) -> dict[str, Any]:
    models = [(os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash").strip()]
    last_error = None
    
    for model_name in models:
        for _ in range(1, 6):
            try:
                res = client.models.generate_content(
                    model=model_name, 
                    contents=contents,
                    config={'response_mime_type': 'application/json'}
                )
                text = getattr(res, "text", "") or ""
                return json.loads(clean_json_response(text))
            except Exception as e:
                last_error = e
                error_str = str(e)
                if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str:
                    time.sleep(65)
                elif "503" in error_str:
                    time.sleep(15)
                else:
                    time.sleep(5)
                    
    raise RuntimeError(last_error)