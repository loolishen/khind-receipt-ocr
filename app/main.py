import io
import zipfile
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List
import re
import pandas as pd
from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .ocr_extractor import run_ocr, _get_ocr  # warmup
from . import parsers

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "app" / "templates"
STATIC_DIR = BASE_DIR / "app" / "static"
OUTPUTS_DIR = BASE_DIR / "outputs"
DATA_DIR = BASE_DIR / "data"

app = FastAPI(title="KHIND Receipt OCR")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
app.mount("/outputs", StaticFiles(directory=OUTPUTS_DIR), name="outputs")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

SUPPORTED_EXT = [".jpg", ".jpeg", ".png"]

# Toggle heavy debug dumps (JSON/TXT) to speed up
DEBUG_DUMP_MAX = 3  # dump only for first N rows; set 0 to disable

def _save_zip_to_dir(zf: UploadFile, dest: Path):
    dest.mkdir(parents=True, exist_ok=True)
    data = zf.file.read()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        z.extractall(dest)

def _scan_images(images_dir: Path, supported_ext: List[str]) -> List[Path]:
    images: List[Path] = []
    for p in images_dir.rglob("*"):
        if p.is_file() and p.suffix.lower() in supported_ext:
            images.append(p)
    # filename sort is fine when names are MLP_XXXX
    images.sort(key=lambda x: x.name.lower())
    return images

_RE_MLP = re.compile(r"(mlp)[_\-\s]*0*(\d+)", re.I)

def _normalize_id(s: str) -> str:
    s = (s or "").strip()
    m = _RE_MLP.search(s)
    if m:
        return f"{m.group(1).lower()}_{m.group(2)}"
    return re.sub(r"[^\w]+", "_", s.lower()).strip("_")

def _normalize_stem_from_path(p: Path) -> str:
    return _normalize_id(p.stem)

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/process", response_class=HTMLResponse)
def process(request: Request, raw_csv: UploadFile = File(...), images_zip: UploadFile = File(...)):
    # Save images zip
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = DATA_DIR / f"run_{ts}"
    img_dir = run_dir / "images"
    _save_zip_to_dir(images_zip, img_dir)

    # Pre-warm the OCR once (this also triggers model download only once)
    try:
        _get_ocr()
    except Exception as e:
        print("[WARN] OCR warmup failed:", e)

    # Build ordered image list
    images_list = _scan_images(img_dir, SUPPORTED_EXT)
    print(f"[IMAGES] Found {len(images_list)} images in {img_dir}")
    if images_list:
        print("[IMAGES] First few:", [p.name for p in images_list[:5]])

    # Build stem -> path
    img_by_norm: Dict[str, Path] = { _normalize_stem_from_path(p): p for p in images_list }

    # Load CSV
    df_raw = pd.read_csv(raw_csv.file)
    if "Submission No" not in df_raw.columns:
        return HTMLResponse("<h3>CSV missing 'Submission No' column.</h3>", status_code=400)

    # Optional: if you have thousands of rows but only 200 images, bound the loop
    n_rows = len(df_raw)

    # Detect fallback city/state columns (not used for location anymore in parsers)
    fallback_city_col: Optional[str] = None
    fallback_state_col: Optional[str] = None
    for c in df_raw.columns:
        cl = c.strip().lower()
        if cl == "city":
            fallback_city_col = c
        elif cl == "state":
            fallback_state_col = c

    new_cols = [
        "Amount spent", "Validity", "Reason for invalid",
        "Product purchased 1", "Amount purchased 1",
        "Product purchased 2", "Amount purchased 2",
        "Product purchased 3", "Amount purchased 3",
        "Store", "Store Location"
    ]
    extracted_rows: List[dict] = []

    used_images: set[Path] = set()

    def _pick_image_for_row(row_idx: int, submission_no: str) -> Optional[Path]:
        norm_sub = _normalize_id(submission_no)

        # 1) exact id match
        p = img_by_norm.get(norm_sub)
        if p is not None and p not in used_images:
            used_images.add(p)
            print(f"[MAP] row {row_idx+2} ({submission_no}) -> {p.name} [EXACT]")
            return p

        # 2) row-order candidate
        if row_idx < len(images_list):
            cand = images_list[row_idx]
            if cand not in used_images:
                used_images.add(cand)
                norm_cand = _normalize_stem_from_path(cand)
                tag = "[ROW-ORDER MATCH]" if norm_cand == norm_sub else "[ROW-ORDER]"
                if norm_cand != norm_sub:
                    print(f"[WARN] row {row_idx+2}: submission '{submission_no}' normalized '{norm_sub}' "
                          f"!= image stem '{norm_cand}'. Using row-order image.")
                print(f"[MAP] row {row_idx+2} ({submission_no}) -> {cand.name} {tag}")
                return cand

        # 3) first unused
        for cand in images_list:
            if cand not in used_images:
                used_images.add(cand)
                norm_cand = _normalize_stem_from_path(cand)
                if norm_cand != norm_sub:
                    print(f"[WARN] row {row_idx+2}: fallback picked '{cand.name}' "
                          f"(norm '{norm_cand}') for submission '{submission_no}' (norm '{norm_sub}').")
                print(f"[MAP] row {row_idx+2} ({submission_no}) -> {cand.name} [FALLBACK]")
                return cand

        print(f"[MAP] row {row_idx+2} ({submission_no}) -> None [NO IMAGE LEFT]")
        return None

    # Process rows
    for row_idx, row in df_raw.iterrows():
        submission_no = str(row.get("Submission No", ""))
        fb_city = str(row.get(fallback_city_col, "")) if fallback_city_col else None
        fb_state = str(row.get(fallback_state_col, "")) if fallback_state_col else None

        image_path = _pick_image_for_row(row_idx, submission_no)

        # --- OCR ---
        ocr_error = False
        image_missing = image_path is None
        lines: List[str] = []
        if not image_missing:
            try:
                # Only dump debug for the first few to keep IO low
                dbg_path = None
                if DEBUG_DUMP_MAX and row_idx < DEBUG_DUMP_MAX:
                    dbg_path = (run_dir / "debug_raw" / f"{_normalize_id(submission_no)}.json")
                lines = run_ocr(image_path, debug_dump_to=dbg_path)
            except Exception as e:
                ocr_error = True
                print(f"[OCR ERROR] row {row_idx+2} ({submission_no}) @ {image_path}: {e}")

        # --- Parse ---
        amount_spent = parsers.extract_amount_spent(lines) if lines else None
        store = parsers.extract_store_name(lines) if lines else None
        # NOTE: extract_store_location ignores participant fallback by design (your latest parser)
        store_loc = parsers.extract_store_location(lines, fb_city, fb_state) if lines else None
        products = parsers.extract_products(lines, max_items=3) if lines else []

        # --- Validity ---
        if image_missing:
            validity, reason = "INVALID", "Image missing"
        elif ocr_error:
            validity, reason = "INVALID", "OCR failed"
        else:
            validity, reason = parsers.decide_validity(amount_spent, products, image_missing=False)

        # --- Row dict ---
        data = {
            "Amount spent": amount_spent or "",
            "Validity": validity,
            "Reason for invalid": reason,
            "Product purchased 1": products[0][0] if len(products) >= 1 else "",
            "Amount purchased 1": products[0][1] if len(products) >= 1 else "",
            "Product purchased 2": products[1][0] if len(products) >= 2 else "",
            "Amount purchased 2": products[1][1] if len(products) >= 2 else "",
            "Product purchased 3": products[2][0] if len(products) >= 3 else "",
            "Amount purchased 3": products[2][1] if len(products) >= 3 else "",
            "Store": store or "",
            "Store Location": store_loc or "",
        }
        extracted_rows.append(data)

    # Assemble output
    df_new = pd.DataFrame(extracted_rows, columns=[
        "Amount spent", "Validity", "Reason for invalid",
        "Product purchased 1", "Amount purchased 1",
        "Product purchased 2", "Amount purchased 2",
        "Product purchased 3", "Amount purchased 3",
        "Store", "Store Location"
    ])
    df_out = pd.concat([df_new, df_raw], axis=1)

    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    csv_path = OUTPUTS_DIR / f"processed_{ts}.csv"
    xlsx_path = OUTPUTS_DIR / f"processed_{ts}.xlsx"
    df_out.to_csv(csv_path, index=False, encoding="utf-8")
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Processed")

    preview_rows = df_out.head(50).to_dict(orient="records")
    return templates.TemplateResponse(
        "results.html",
        {
            "request": request,
            "csv_url": f"/outputs/{csv_path.name}",
            "xlsx_url": f"/outputs/{xlsx_path.name}",
            "row_count": len(df_out),
            "columns": list(df_out.columns),
            "rows": preview_rows
        }
    )
