# deepseek_ocr_backend.py
from pathlib import Path
from typing import List, Optional
import os
import re

from transformers import AutoModel, AutoTokenizer
import torch

MODEL_NAME = "deepseek-ai/DeepSeek-OCR"

_tokenizer = None
_model = None

def _init_model():
    global _tokenizer, _model
    if _model is not None:
        return
    os.environ.setdefault("CUDA_VISIBLE_DEVICES", "0")
    _tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, trust_remote_code=True)
    _model = AutoModel.from_pretrained(
        MODEL_NAME,
        _attn_implementation="flash_attention_2",
        trust_remote_code=True,
        use_safetensors=True,
    )
    # BF16 on GPU as recommended on HF card :contentReference[oaicite:3]{index=3}
    if torch.cuda.is_available():
        _model = _model.eval().cuda().to(torch.bfloat16)
    else:
        _model = _model.eval()

def _markdown_to_lines(md: str) -> List[str]:
    """
    Convert DeepSeek OCR markdown/text output into clean receipt lines
    optimized for parsers.py heuristics.
    """
    lines: List[str] = []

    if not md:
        return lines

    for raw in md.splitlines():
        s = raw.strip()
        if not s:
            continue

        # 1) Remove markdown bullets / headings
        s = re.sub(r"^(\*+|-+|#+|\u2022|\>)\s*", "", s)

        # 2) Collapse whitespace
        s = " ".join(s.split())

        # 3) Ignore obvious non-receipt noise
        if s.lower() in {"thank you", "thanks", "please come again"}:
            continue

        # 4) Split combined TOTAL lines
        # Example: "TOTAL RM88.00 THANK YOU"
        m = _SPLIT_TOTAL_TAIL.search(s)
        if m:
            main = m.group(1).strip()
            tail = m.group(2).strip()
            if main:
                lines.append(main)
            if tail and len(tail) > 3:
                lines.append(tail)
            continue

        # 5) Split lines with multiple currency values
        # Example: "Subtotal RM10.00 Tax RM0.60 Total RM10.60"
        if s.upper().count("RM") >= 2:
            parts = re.split(r"(?=(?:RM|MYR|R\s*M)\s*\d)", s, flags=re.I)
            for p in parts:
                p = p.strip()
                if p:
                    lines.append(p)
            continue

        # 6) Normal case
        lines.append(s)

    return lines

def run_deepseek_ocr(image_path: Path, prompt: Optional[str] = None) -> List[str]:
    """
    Run DeepSeek-OCR on a single image and return a list of cleaned text lines.
    """
    _init_model()
    if prompt is None:
        # Keep it generic but "receipt aware"
        prompt = "<image>\nFree OCR. Extract all visible text from this receipt clearly, line by line."

    img_path_str = str(image_path)

    res = _model.infer(
        _tokenizer,
        prompt=prompt,
        image_file=img_path_str,
        output_path="",        # we don't need visualization
        base_size=1024,
        image_size=640,
        crop_mode=True,
        save_results=False,
        test_compress=True,
    )

    # 'res' is usually a string (markdown) or list; handle string safely
    if isinstance(res, str):
        text = res
    elif isinstance(res, dict) and "text" in res:
        text = str(res["text"])
    else:
        text = str(res)

    return _markdown_to_lines(text)
