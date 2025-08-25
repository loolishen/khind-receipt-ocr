from pathlib import Path
from typing import List, Optional, Any
from paddleocr import PaddleOCR
from PIL import Image, ImageOps
import numpy as np
import json
import math
from typing import Any, Iterable, Tuple
import numpy as np


# Config
MAX_OCR_SIDE = 2200
JPEG_QUALITY = 85

_ocr_instance: Optional[PaddleOCR] = None

def _get_ocr() -> PaddleOCR:
    global _ocr_instance
    if _ocr_instance is not None:
        return _ocr_instance
    # try a few configs (some builds don't like show_log/use_angle_cls combos)
    for kwargs in (
        {"lang": "en", "use_angle_cls": True, "show_log": False},
        {"lang": "en", "use_angle_cls": True},
        {"lang": "en"},
        {},
    ):
        try:
            _ocr_instance = PaddleOCR(**kwargs)
            break
        except Exception:
            _ocr_instance = None
            continue
    if _ocr_instance is None:
        _ocr_instance = PaddleOCR()
    return _ocr_instance

def _cached_resized_path(orig: Path) -> Path:
    cache_dir = orig.parent / "_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir / f"{orig.stem}_s{MAX_OCR_SIDE}.jpg"

def _prepare_image_for_ocr(image_path: Path) -> Path:
    try:
        cached = _cached_resized_path(image_path)
        if cached.exists():
            return cached
        with Image.open(image_path) as im:
            im = ImageOps.exif_transpose(im).convert("RGB")
            w, h = im.size
            m = max(w, h)
            if m > MAX_OCR_SIDE:
                scale = MAX_OCR_SIDE / float(m)
                im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.BILINEAR)
            im.save(cached, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        return cached
    except Exception:
        return image_path

def _group_into_lines(items: Iterable[Tuple[float, float, float, float, str]], y_tol=10.0) -> list[str]:
    """
    items: iterable of (cx, cy, x0, x1, text)
    Groups tokens by similar y (cy) and sorts by x0 to produce lines.
    y_tol: vertical tolerance in pixels (after resize); tune 8–14 if needed.
    """
    rows: list[list[Tuple[float, float, float, float, str]]] = []
    for cx, cy, x0, x1, txt in items:
        placed = False
        for row in rows:
            # same row if vertical distance small vs row's mean cy
            mean_cy = sum(r[1] for r in row) / max(1, len(row))
            if abs(cy - mean_cy) <= y_tol:
                row.append((cx, cy, x0, x1, txt))
                placed = True
                break
        if not placed:
            rows.append([(cx, cy, x0, x1, txt)])

    # sort rows by vertical position, then tokens by x
    rows.sort(key=lambda r: sum(t[1] for t in r) / len(r))
    lines: list[str] = []
    for row in rows:
        row.sort(key=lambda t: t[2])
        # join with single spaces
        parts = [t[4] for t in row if t[4]]
        s = " ".join(" ".join(parts).split())
        if s:
            lines.append(s)
    return lines

def _extract_items_from_classic(page: Any) -> list[Tuple[float, float, float, float, str]]:
    """
    Classic PaddleOCR 'page' is: [ [box, (text, prob)], ... ]
    box: [[x0,y0],[x1,y1],[x2,y2],[x3,y3]]
    """
    out = []
    if not isinstance(page, list):
        return out
    for item in page:
        try:
            box, pair = item[0], item[1]
            text = pair[0] if isinstance(pair, (list, tuple)) else None
            if not text:
                continue
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            x0, x1 = min(xs), max(xs)
            y0, y1 = min(ys), max(ys)
            cx = 0.5 * (x0 + x1)
            cy = 0.5 * (y0 + y1)
            out.append((cx, cy, x0, x1, str(text)))
        except Exception:
            continue
    return out

def _flatten_text_any(raw: Any) -> list[str]:
    """
    Robust text extractor that supports:
    1) PaddleOCR new format: [ { rec_texts: [...], ... }, ... ]
    2) Classic format: [ [ [box], ('text', prob) ], ... ] or list of such pages
    3) Generic nested dict/list with 'text'/'label'
    """
    lines: list[str] = []

    def add(s: str):
        s = " ".join(str(s).split())
        if s:
            lines.append(s)

    # ---- CASE 1: New Paddle dict format (rec_texts) ----
    if isinstance(raw, list) and raw and all(isinstance(p, dict) for p in raw):
        # Concatenate rec_texts page by page (reading order is already sensible)
        any_rec = False
        for page in raw:
            recs = page.get("rec_texts")
            if isinstance(recs, list):
                any_rec = True
                for t in recs:
                    if isinstance(t, str):
                        add(t)
        if any_rec:
            return lines

    # ---- CASE 2: Classic list format with boxes ----
    try:
        if isinstance(raw, list):
            # single page nested once (common)
            if len(raw) == 1 and isinstance(raw[0], list) and raw[0]:
                items = _extract_items_from_classic(raw[0])
                if items:
                    return _group_into_lines(items)
            # multi-page: collect all items
            any_items = []
            for page in raw:
                if isinstance(page, list):
                    items = _extract_items_from_classic(page)
                    if items:
                        any_items.extend(items)
            if any_items:
                return _group_into_lines(any_items)
    except Exception:
        pass

    # ---- CASE 3: Generic recursive walk for 'text'/'label' ----
    def walk(x: Any):
        if x is None:
            return
        if isinstance(x, dict):
            for k in ("text", "label"):
                if isinstance(x.get(k), str):
                    add(x[k])
            # common container keys
            for k in ("data", "res", "result", "ocr", "items"):
                if k in x:
                    walk(x[k])
            return
        if isinstance(x, (list, tuple)):
            # ('hello', 0.98)
            if len(x) == 2 and isinstance(x[1], (list, tuple)) and isinstance(x[1][0], str):
                add(x[1][0])
                return
            for y in x:
                walk(y)
            return

    walk(raw)
    return lines

    """
    Robust flattener:
    1) If we can find classic items with boxes, we line-group them.
    2) Else, collect 'text'/'label' fields or ('text',prob) pairs recursively.
    """
    # CASE A: list -> (maybe) list-of-pages -> classic items
    try:
        if isinstance(raw, list):
            # If it's a single page nested once (common)
            if len(raw) == 1 and isinstance(raw[0], list) and raw and raw[0]:
                items = _extract_items_from_classic(raw[0])
                if items:
                    return _group_into_lines(items)
            # If it's multi-page (list of lists)
            any_items = []
            for page in raw:
                if isinstance(page, list):
                    items = _extract_items_from_classic(page)
                    if items:
                        any_items.extend(items)
            if any_items:
                return _group_into_lines(any_items)
    except Exception:
        pass

    # CASE B: dict wrapper(s) with 'data'
    lines: list[str] = []
    def add(s: str):
        s = " ".join(str(s).split())
        if s:
            lines.append(s)

    def walk(x: Any):
        if x is None:
            return
        if isinstance(x, dict):
            # common keys
            for k in ("text", "label"):
                if k in x and isinstance(x[k], str):
                    add(x[k]);  # don't return; dict may contain nested pieces
            if "data" in x and isinstance(x["data"], list):
                for d in x["data"]:
                    walk(d)
            # sometimes wrapped under 'res' or 'result'
            for k in ("res", "result", "ocr", "items"):
                if k in x:
                    walk(x[k])
            return
        if isinstance(x, (list, tuple)):
            # ('hello', 0.98) or [ [box], ('hello', 0.98) ]
            if len(x) == 2 and isinstance(x[1], (list, tuple)) and isinstance(x[1][0], str):
                add(x[1][0])
                return
            for y in x:
                walk(y)
            return
        # ignore scalars that aren't strings

    walk(raw)
    return lines

def run_ocr(image_path: Path, debug_dump_to: Optional[Path] = None) -> List[str]:
    ocr = _get_ocr()
    prepped = _prepare_image_for_ocr(Path(image_path))
    with Image.open(prepped) as im:
        im = im.convert("RGB")
        arr = np.array(im)
    try:
        raw = ocr.ocr(arr, cls=True)
    except TypeError:
        raw = ocr.ocr(arr)

    # --- dump raw for inspection ---
    if debug_dump_to:
        try:
            debug_dump_to.parent.mkdir(parents=True, exist_ok=True)
            import json
            with open(debug_dump_to, "w", encoding="utf-8") as f:
                f.write(json.dumps(raw, ensure_ascii=False, indent=2, default=str))
        except Exception:
            pass

    # --- flatten to lines ---
    lines = _flatten_text_any(raw)

    # also write the flattened text beside the json (handy when debugging)
    try:
        if debug_dump_to:
            txt_path = debug_dump_to.with_suffix(".txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
    except Exception:
        pass

    return lines