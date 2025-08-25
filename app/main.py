import io
import zipfile
from pathlib import Path
from datetime import datetime
from typing import Optional

import pandas as pd
from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .ocr_extractor import run_ocr
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

SUPPORTED_EXT = [".jpg", ".jpeg", ".png"]  # .svg not reliable for OCR


def _save_zip_to_dir(zf: UploadFile, dest: Path):
    dest.mkdir(parents=True, exist_ok=True)
    data = zf.file.read()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        z.extractall(dest)


def _scan_images(images_dir: Path, supported_ext: list[str]) -> list[Path]:
    """
    Return all images under images_dir in **sorted filename order**.
    We will map them to the CSV rows by index:
      row 0 (Excel row 2) -> images[0] (e.g., image_0.*)
      row 1 (Excel row 3) -> images[1], etc.
    """
    images: list[Path] = []
    for p in images_dir.rglob("*"):
        if p.is_file() and p.suffix.lower() in supported_ext:
            images.append(p)
    images.sort(key=lambda x: x.name.lower())
    return images


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

    # Build ordered image list
    images_list = _scan_images(img_dir, SUPPORTED_EXT)
    print(f"[IMAGES] Found {len(images_list)} images in {img_dir}")
    if images_list:
        sample = [p.name for p in images_list[:5]]
        print("[IMAGES] First few:", sample)

    # Load raw CSV
    df_raw = pd.read_csv(raw_csv.file)
    if "Submission No" not in df_raw.columns:
        return HTMLResponse("<h3>CSV missing 'Submission No' column.</h3>", status_code=400)

    # Detect fallback city/state columns (once)
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
    extracted_rows = []

    for row_idx, row in df_raw.iterrows():
        # --- values from row ---
        submission_no = str(row.get("Submission No", ""))
        fb_city = str(row.get(fallback_city_col, "")) if fallback_city_col else None
        fb_state = str(row.get(fallback_state_col, "")) if fallback_state_col else None

        # --- map to image by row index ---
        image_path = images_list[row_idx] if row_idx < len(images_list) else None
        print(f"[MAP] row {row_idx+2} ({submission_no}) -> {image_path}")

        # --- run OCR (with explicit error flag) ---
        ocr_error = False
        image_missing = image_path is None
        lines: list[str] = []
        if not image_missing:
            try:
                dbg_raw_path = run_dir / "debug_raw" / f"{submission_no}.json"
                lines = run_ocr(image_path, debug_dump_to=dbg_raw_path)

                dbg_dir = run_dir / "debug_txt"
                dbg_dir.mkdir(parents=True, exist_ok=True)
                with open(dbg_dir / f"{submission_no}.txt", "w", encoding="utf-8") as f:
                    f.write("\n".join(lines))
            except Exception as e:
                ocr_error = True
                print(f"[OCR ERROR] row {row_idx+2} ({submission_no}) @ {image_path}: {e}")

        # --- parse fields ---
        amount_spent = parsers.extract_amount_spent(lines) if lines else None
        store = parsers.extract_store_name(lines) if lines else None
        store_loc = (
            parsers.extract_store_location(lines, fb_city, fb_state) if lines
            else (f"{fb_city}, {fb_state}" if fb_city and fb_state else (fb_state or ""))
        )
        products = parsers.extract_products(lines, max_items=3) if lines else []

        # --- validity / reason (set ONCE) ---
        if image_missing:
            validity, reason = "INVALID", "Image missing"
        elif ocr_error:
            validity, reason = "INVALID", "OCR failed"
        else:
            validity, reason = parsers.decide_validity(amount_spent, products, image_missing=False)

        # --- assemble output row ---
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

    df_new = pd.DataFrame(extracted_rows, columns=new_cols)
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
