# yolo_receipt_crop.py
from pathlib import Path
from typing import Optional
import cv2
from ultralytics import YOLO

# You can start with a general model; later replace with your own receipt-only .pt
YOLO_MODEL_PATH = "yolov8n.pt"

_yolo_model: Optional[YOLO] = None

def _get_model() -> YOLO:
    global _yolo_model
    if _yolo_model is None:
        _yolo_model = YOLO(YOLO_MODEL_PATH)
    return _yolo_model

def crop_receipt(image_path: Path, out_dir: Optional[Path] = None) -> Path:
    """
    Run YOLO on the image and crop the largest detected rectangular region.
    If YOLO finds nothing, return the original image path.
    """
    img_path = str(image_path)
    model = _get_model()
    results = model.predict(img_path, conf=0.3, verbose=False)

    if not results or len(results[0].boxes) == 0:
        return image_path  # fallback

    # pick largest box by area
    boxes = results[0].boxes.xyxy.cpu().numpy()  # (N, 4) x1,y1,x2,y2
    if boxes.shape[0] == 0:
        return image_path

    areas = (boxes[:, 2] - boxes[:, 0]) * (boxes[:, 3] - boxes[:, 1])
    idx = areas.argmax()
    x1, y1, x2, y2 = boxes[idx].astype(int)

    img = cv2.imread(img_path)
    if img is None:
        return image_path

    h, w = img.shape[:2]
    x1 = max(0, min(x1, w-1))
    y1 = max(0, min(y1, h-1))
    x2 = max(1, min(x2, w))
    y2 = max(1, min(y2, h))

    crop = img[y1:y2, x1:x2]
    if crop.size == 0:
        return image_path

    if out_dir is None:
        out_dir = image_path.parent / "_cropped"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{image_path.stem}_crop.jpg"
    cv2.imwrite(str(out_path), crop)
    return out_path
