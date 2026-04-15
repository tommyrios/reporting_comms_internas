import json
import time
from typing import Any
from google import genai
from config import PDF_DIR, SUMMARIES_DIR, ensure_dir
from llm_client import load_prompt, call_gemini_for_json

def summarize_month(client: genai.Client, month_key: str, force_regenerate: bool = False) -> dict[str, Any]:
    path = ensure_dir(SUMMARIES_DIR) / f"{month_key}.json"
    if path.exists() and not force_regenerate: 
        return json.loads(path.read_text(encoding="utf-8"))
    
    pdf_path = PDF_DIR / f"{month_key}.pdf"
    uploaded = client.files.upload(file=str(pdf_path))
    prompt_text = load_prompt("monthly_summary.txt")
    
    try:
        while uploaded.state.name == "PROCESSING":
            time.sleep(2)
            uploaded = client.files.get(name=uploaded.name)
        
        if uploaded.state.name == "FAILED":
            raise RuntimeError("FAILED")
            
        summary = call_gemini_for_json(client, [uploaded, prompt_text])
        summary["month"] = month_key
        path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        return summary
    finally:
        try: client.files.delete(name=uploaded.name)
        except: pass