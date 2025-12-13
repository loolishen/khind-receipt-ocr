from pathlib import Path
from typing import List, Optional, Any, Iterable, Tuple
import os
import json
import numpy as np
from PIL import Image, ImageOps

# IMPORTANT: Import paddle/paddleocr lazily so module import is fast
_PaddleOCR = None
_paddle_device = None

# -------- Tunables (trade accuracy vs speed) --------
MAX_OCR_SIDE = 1600   # down from 2200; big speedup with minimal loss for receipts
JPEG_QUALITY = 85

# Batch more text crops per forward pass
REC_BATCH_NUM = 32

# Threads for CPU math libs (helps on MKL/OpenBLAS)
os.environ.setdefault("OMP_NUM_THREADS", "4")
os.environ.setdefault("OPENBLAS_NUM_THREADS", "4")
os.environ.setdefault("MKL_NUM_THREADS", "4")
os.environ.setdefault("VECLIB_MAXIMUM_THREADS", "4")
os.environ.setdefault("NUMEXPR_NUM_THREADS", "4")

_ocr_instance = None

def _try_import_paddle():
    global _PaddleOCR, _paddle_device
    if _PaddleOCR is not None:
        return
    from paddleocr import PaddleOCR  # noqa: F401
    import paddle
    _PaddleOCR = PaddleOCR
    _paddle_device = paddle

def _get_ocr():
    """
    One-time initializer. Tries CUDA; falls back to CPU.
    Disables angle classifier (often not needed on receipts) and reduces det side len.
    Increases rec batch size. Hides logs.
    """
    global _ocr_instance
    if _ocr_instance is not None:
        return _ocr_instance

    _try_import_paddle()

    use_gpu = False
    try:
        use_gpu = bool(_paddle_device.device.is_compiled_with_cuda())  # type: ignore[attr-defined]
    except Exception:
        use_gpu = False

    # Some builds dislike certain flags; we try a few combos quickly
    trial_args = [
        dict(
            lang="en",
            use_angle_cls=False,     # disable for speed
            show_log=False,
            det_limit_type="max",
            det_limit_side_len=1280, # smaller detector input
            rec_batch_num=REC_BATCH_NUM,
            cpu_threads=4,
            use_gpu=use_gpu,
        ),
        dict(
            lang="en",
            use_angle_cls=False,
            show_log=False,
            rec_batch_num=REC_BATCH_NUM,
            cpu_threads=4,
            use_gpu=use_gpu,
        ),
        dict(lang="en", show_log=False, use_angle_cls=False),
        dict(lang="en"),
        {},
    ]

    for kwargs in trial_args:
        try:
            _ocr_instance = _PaddleOCR(**kwargs)
            break
        except Exception:
            _ocr_instance = None
            continue

    if _ocr_instance is None:
        # last resort
        _ocr_instance = _PaddleOCR()

    return _ocr_instance

def _cached_resized_path(orig: Path) -> Path:
    cache_dir = orig.parent / "_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir / f"{orig.stem}_s{MAX_OCR_SIDE}.jpg"

def _prepare_image_for_ocr(image_path: Path) -> Path:
    """
    Resize long side to MAX_OCR_SIDE (cached). This alone saves a lot of time.
    """
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
    rows: list[list[Tuple[float, float, float, float, str]]] = []
    for cx, cy, x0, x1, txt in items:
        placed = False
        for row in rows:
            mean_cy = sum(r[1] for r in row) / max(1, len(row))
            if abs(cy - mean_cy) <= y_tol:
                row.append((cx, cy, x0, x1, txt)); placed = True; break
        if not placed:
            rows.append([(cx, cy, x0, x1, txt)])

    rows.sort(key=lambda r: sum(t[1] for t in r) / len(r))
    lines: list[str] = []
    for row in rows:
        row.sort(key=lambda t: t[2])
        parts = [t[4] for t in row if t[4]]
        s = " ".join(" ".join(parts).split())
        if s: lines.append(s)
    return lines

def _extract_items_from_classic(page: Any) -> list[Tuple[float, float, float, float, str]]:
    out = []
    if not isinstance(page, list):
        return out
    for item in page:
        try:
            box, pair = item[0], item[1]
            text = pair[0] if isinstance(pair, (list, tuple)) else None
            if not text: continue
            xs = [p[0] for p in box]; ys = [p[1] for p in box]
            x0, x1 = min(xs), max(xs); y0, y1 = min(ys), max(ys)
            cx = 0.5 * (x0 + x1); cy = 0.5 * (y0 + y1)
            out.append((cx, cy, x0, x1, str(text)))
        except Exception:
            continue
    return out

def _flatten_text_any(raw: Any) -> list[str]:
    """
    Supports both new Paddle dict format and classic format.
    """
    lines: list[str] = []

    def add(s: str):
        s = " ".join(str(s).split())
        if s: lines.append(s)

    # New dict format
    if isinstance(raw, list) and raw and all(isinstance(p, dict) for p in raw):
        any_rec = False
        for page in raw:
            recs = page.get("rec_texts")
            if isinstance(recs, list):
                any_rec = True
                for t in recs:
                    if isinstance(t, str): add(t)
        if any_rec:
            return lines

    # Classic with boxes (single or multi page)
    try:
        if isinstance(raw, list):
            if len(raw) == 1 and isinstance(raw[0], list) and raw[0]:
                items = _extract_items_from_classic(raw[0])
                if items: return _group_into_lines(items)
            any_items = []
            for page in raw:
                if isinstance(page, list):
                    items = _extract_items_from_classic(page)
                    if items: any_items.extend(items)
            if any_items: return _group_into_lines(any_items)
    except Exception:
        pass

    # Generic recursive walk
    def walk(x: Any):
        if x is None: return
        if isinstance(x, dict):
            for k in ("text", "label"):
                v = x.get(k)
                if isinstance(v, str): add(v)
            for k in ("data", "res", "result", "ocr", "items"):
                if k in x: walk(x[k])
            return
        if isinstance(x, (list, tuple)):
            if len(x) == 2 and isinstance(x[1], (list, tuple)) and isinstance(x[1][0], str):
                add(x[1][0]); return
            for y in x: walk(y)
            return
    walk(raw)
    return lines

def run_ocr(image_path: Path, debug_dump_to: Optional[Path] = None) -> List[str]:
    ocr = _get_ocr()
    prepped = _prepare_image_for_ocr(Path(image_path))

    # Read once to numpy (fast path)
    with Image.open(prepped) as im:
        arr = np.array(im.convert("RGB"))

    # Call OCR (cls disabled in config; fewer passes)
    try:
        raw = ocr.ocr(arr)  # we configured it without angle cls; no need cls=True
    except TypeError:
        raw = ocr.ocr(arr)

    if debug_dump_to:
        try:
            debug_dump_to.parent.mkdir(parents=True, exist_ok=True)
            with open(debug_dump_to, "w", encoding="utf-8") as f:
                f.write(json.dumps(raw, ensure_ascii=False, indent=2, default=str))
            txt_path = debug_dump_to.with_suffix(".txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("\n".join(_flatten_text_any(raw)))
        except Exception:
            pass

    return _flatten_text_any(raw)
