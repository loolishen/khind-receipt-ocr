# vlm_receipt_parser.py
from pathlib import Path
from typing import Dict, Any, List
import os
import base64

from openai import OpenAI

VLM_BASE_URL = os.getenv("VLM_BASE_URL", "https://api.deepseek.com")
VLM_MODEL_NAME = os.getenv("VLM_MODEL_NAME", "deepseek-chat")

_client = None

def _get_client() -> OpenAI:
    global _client
    if _client is None:
        _client = OpenAI(
            api_key=os.environ["VLM_API_KEY"],
            base_url=VLM_BASE_URL,
        )
    return _client

def _encode_image(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")

def vlm_parse_receipt(image_path: Path, ocr_text: str) -> Dict[str, Any]:
    """
    Ask a VLM to parse the receipt into structured JSON.
    """
    client = _get_client()
    img_b64 = _encode_image(image_path)

    prompt = (
        "You are a strict receipt parser for Malaysian receipts.\n"
        "Given the image and OCR text, return ONLY valid JSON with keys:\n"
        "{\n"
        '  "store_name": string or null,\n'
        '  "store_location": string or null,\n'
        '  "amount_spent": string like "RM123.45" or null,\n'
        '  "items": [ {"name": string, "qty": int} ]\n'
        "}\n"
        "Do not add explanations. Try to be consistent with the OCR text:\n"
        f"OCR_TEXT:\n{ocr_text[:3000]}\n"
    )

    resp = client.chat.completions.create(
        model=VLM_MODEL_NAME,
        messages=[
            {"role": "system", "content": "You output only JSON."},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"},
                    },
                ],
            },
        ],
        temperature=0.0,
    )
    content = resp.choices[0].message.content
    # content should be JSON string; parse on caller side
    import json
    try:
        return json.loads(content)
    except Exception:
        return {}
