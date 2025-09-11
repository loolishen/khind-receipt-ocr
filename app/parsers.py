import re
from typing import List, Tuple, Optional

# ---------------------------------
# OPTIONAL curated lists (auto-load)
# ---------------------------------
# If these modules are present (generated from your CSV),
# their lists will be used as defaults unless the caller provides one.

# Stores (curated)
try:
    from .store_hints_w4 import STORE_HINTS_W4 as PREFERRED_STORE_HINTS  # type: ignore
except Exception:
    try:
        from store_hints_w4 import STORE_HINTS_W4 as PREFERRED_STORE_HINTS  # type: ignore
    except Exception:
        PREFERRED_STORE_HINTS = None

# Products / Items (curated)
PREFERRED_PRODUCT_HINTS: Optional[List[str]] = None
try:
    from .product_hints_w4 import PRODUCT_HINTS_W4 as PREFERRED_PRODUCT_HINTS  # type: ignore
except Exception:
    try:
        from product_hints_w4 import PRODUCT_HINTS_W4 as PREFERRED_PRODUCT_HINTS  # type: ignore
    except Exception:
        try:
            from .items_hints_w4 import ITEMS_HINTS_W4 as PREFERRED_PRODUCT_HINTS  # type: ignore
        except Exception:
            try:
                from items_hints_w4 import ITEMS_HINTS_W4 as PREFERRED_PRODUCT_HINTS  # type: ignore
            except Exception:
                PREFERRED_PRODUCT_HINTS = None

# -----------------------------
# constants / helpers
# -----------------------------

MALAYSIAN_STATES = [
    "Johor","Kedah","Kelantan","Malacca","Melaka","Negeri Sembilan","Pahang","Penang","Pulau Pinang",
    "Perak","Perlis","Sabah","Sarawak","Selangor","Terengganu","Kuala Lumpur","WP Kuala Lumpur",
    "Labuan","Putrajaya"
]

# chains / hints commonly seen in headers
STORE_HINTS = [
    # your preferred list first (stronger priority when passed via extract_store_name)
    "Khind Customer Service SDN BHD", "AEON BIG", "Jaya Superstore", "HomePro",
    "XingLee Electrical & HIFI Central SDN BHD", "AEON Big", "SK Hardware",
    "SK SERIES ENTERPRISE", "JAYA GROCER", "SNF ONLINE (M) SDN BHD", "Kemudi Timur",
    "ECONSAVE CASH & CARRY (FH) SDN BHD", "SENHENG", "HARVEY NORMAN",
    "Seruay Evergreen SDN BHD", "DYNAPOWER ENTERPRISE",
    "SENG HUAT ELECTRICAL & HOME APPLIANCES SDN BHD", "SURIA JERAI ELECTRICAL SDN BHD",
    "Diva Lighting", "Shopee Online", "LSS HYPER ELECTRICAL SDN BHD",
    "IKA Houseware Malaysia SDN BHD", "YING MIEW HOLDINGS",
    "SYARIKAT LETRIK LIM & ONG SDN BHD", "Jaya",
    # generic hints
    "KHIND", "SDN BHD", "ELECTRICAL", "SUPERMARKET", "HYPER", "STORE", "LIGHTING"
]

# numeric patterns
_NUM      = r"(?:\d{1,3}(?:,\d{3})*(?:\.\d{2})|\d+(?:\.\d{2})?)"
_CCY      = r"(?:RM|MYR|R\s*M)"
_SEP      = r"[:=\-\s]*"

# totals keywords (broad)
KW_TOTAL = (
    r"(?:grand\s*total|total\s*amount|total\s*payable|total\s*price|"
    r"net\s*amount|net\s*total|amount\s*due|amount\s*payable|balance\s*due|"
    r"total\s*after\s*discount|total\b)"
)

RE_KW_LEFT   = re.compile(rf"{KW_TOTAL}{_SEP}{_CCY}?\s*({_NUM})\b", re.I)
RE_KW_RIGHT  = re.compile(rf"{KW_TOTAL}{_SEP}({_NUM})\s*{_CCY}\b", re.I)
RE_ANY_CCY   = re.compile(rf"{_CCY}\s*({_NUM})\b", re.I)
RE_ANY_NUM   = re.compile(rf"({_NUM})")
RE_JUST_KW   = re.compile(KW_TOTAL, re.I)

# product-ish patterns
PRODUCT_CODE_RE = re.compile(r"(?<![A-Z0-9])([A-Z]{2,}[A-Z0-9\-]{1,})(?![A-Z0-9])")
QTY_RE = re.compile(
    r"\b(?:QTY|QTY:|QTY\.|QTY=|QTY\s+)\s*(\d+)\b|\b(\d+)x\b|\bx(\d+)\b|\b(\d+)\s*(?:pcs|unit|units|pcs\.)\b",
    re.IGNORECASE
)

# Recognize price tokens allowed on KHIND line only:
#  - Currency-tagged numbers
#  - Bare decimals with 2 d.p. (not glued to letters)
PRICE_TOKEN_RE = re.compile(
    r"""
    (?:
        (?P<ccy>RM|MYR|R\s*M)\s*
        (?P<ccyval>\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d+\.\d{2})
    )
    |
    (?:
        (?<![A-Za-z])(?P<dec>\d{1,3}(?:,\d{3})*\.\d{2})(?![A-Za-z])
    )
    """,
    re.IGNORECASE | re.VERBOSE
)

# qty context used to avoid picking qty instead of price
_QTY_WORDS_RE = re.compile(r"\b(qty|unit|units|pcs|pcs\.|piece|pieces|x\d+|\d+x)\b", re.I)

def _to_float(s: str) -> float:
    return float(s.replace(",", ""))

def _normalize_amount(val: str) -> Optional[str]:
    raw = re.sub(r"[^\d.,]", "", val or "")
    if not raw:
        return None
    if raw.count(".") >= 1 and raw.rsplit(".", 1)[-1].isdigit():
        num = raw.replace(",", "")
    elif raw.count(",") >= 1 and raw.rsplit(",", 1)[-1].isdigit():
        num = raw.replace(".", "").replace(",", ".")
    else:
        num = raw.replace(",", "")
    try:
        return f"RM{float(num):.2f}"
    except Exception:
        return None

def _contains_store_hint(text: str) -> bool:
    up = text.upper()
    return any(h in up for h in STORE_HINTS)

# -----------------------------
# store name / location
# -----------------------------

def extract_store_name(lines: List[str], preferred_stores: Optional[List[str]] = None) -> Optional[str]:
    """
    Prefer user-provided store list; then generic hints; then ALLCAPS; then first non-empty.
    If store_hints_w4.py is available, it is used by default as preferred_stores.
    """
    # Default to curated list when caller did not pass one
    if preferred_stores is None and PREFERRED_STORE_HINTS:
        preferred_stores = PREFERRED_STORE_HINTS

    top = lines[:10]

    if preferred_stores:
        pref_up = [s.upper() for s in preferred_stores]
        matches = []
        for ln in top:
            up = ln.upper()
            if any(s in up for s in pref_up):
                matches.append(ln.strip())
        if matches:
            return max(matches, key=len)

    hinted = [ln for ln in top if _contains_store_hint(ln)]
    if hinted:
        return max(hinted, key=len).strip()

    allcaps = [ln for ln in top if ln.strip() and ln.strip().upper() == ln.strip()]
    if allcaps:
        return max(allcaps, key=len).strip()

    for ln in top:
        if ln.strip():
            return ln.strip()
    return None

def extract_store_location(lines: List[str], fallback_city: Optional[str], fallback_state: Optional[str]) -> Optional[str]:
    states_ci = {s.lower(): s for s in MALAYSIAN_STATES}
    for line in reversed(lines[-14:]):
        low = line.lower()
        state_norm = next((states_ci[s] for s in states_ci if s in low), None)
        if not state_norm:
            continue
        parts = re.split(r"[,\-;|]", line)
        city = None
        for i, seg in enumerate(parts):
            if state_norm.lower() in seg.lower():
                if i > 0:
                    city = parts[i-1].strip()
                break
        if not city:
            left = low.split(state_norm.lower())[0]
            left = re.sub(r"[^a-zA-Z\s]", " ", left).strip()
            toks = [t for t in left.split() if len(t) >= 3]
            city = " ".join(toks[-2:]).strip() if toks else None
        return f"{city}, {state_norm}" if city else state_norm

    if fallback_city and fallback_state:
        return f"{fallback_city}, {fallback_state}"
    if fallback_state:
        return fallback_state
    return None

# -----------------------------
# KHIND row → price (same line only)
# -----------------------------

def _price_candidates(line: str) -> List[Tuple[int, float, int]]:
    """
    Return (pos, value, score). score 3 = currency-tagged, score 2 = bare decimal.
    Ignore plain integers (years/model numbers).
    """
    hits: List[Tuple[int, float, int]] = []
    for m in PRICE_TOKEN_RE.finditer(line):
        if m.group("ccy"):
            val = _to_float(m.group("ccyval"))
            hits.append((m.start("ccyval"), val, 3))
        elif m.group("dec"):
            val = _to_float(m.group("dec"))
            hits.append((m.start("dec"), val, 2))
    hits.sort(key=lambda x: x[0])  # left→right
    return hits

def _find_khind_rows(lines: List[str]) -> List[Tuple[int, str]]:
    return [(i, ln) for i, ln in enumerate(lines) if "KHIND" in ln.upper()]

def _looks_like_qty_context(line: str, span_start: int, span_end: int) -> bool:
    left = max(0, span_start - 10)
    right = min(len(line), span_end + 10)
    return bool(_QTY_WORDS_RE.search(line[left:right]))

def _choose_rightmost_best(cands: List[Tuple[int, float, int]], line: str) -> Optional[float]:
    """
    Prefer highest score; within same score choose the rightmost token.
    Reject values near qty context; keep 2..100000.
    """
    if not cands:
        return None
    best = max(s for _, _, s in cands)
    for pos, val, s in reversed([c for c in cands if c[2] == best]):
        if 2.0 <= val <= 100000.0 and not _looks_like_qty_context(line, pos, pos + 1):
            return val
    for pos, val, s in reversed(cands):
        if 2.0 <= val <= 100000.0 and not _looks_like_qty_context(line, pos, pos + 1):
            return val
    return None

# lines that should stop the look-ahead when scanning a KHIND row
_STOP_AFTER_RE = re.compile(r"\b(total|grand\s*total|balance|remarks?|thank|cash\s*rm?)\b", re.I)

def _khind_line_amount(lines: List[str]) -> Optional[str]:
    """
    Find price for a KHIND item where the row may span multiple OCR lines.
    Strategy:
      - Start at the line containing 'KHIND'
      - Look ahead up to 4 lines as long as we don't hit 'Total/Grand Total/Remarks/...'
      - On each line, collect price candidates and pick the rightmost, preferring
        later lines first (prices often appear a couple of lines below the desc).
    """
    LOOKAHEAD = 4

    for idx, ln in _find_khind_rows(lines):
        # scan forward window, skipping summary lines
        window_idxs = []
        for j in range(idx, min(len(lines), idx + LOOKAHEAD + 1)):
            if _STOP_AFTER_RE.search(lines[j]):
                break
            window_idxs.append(j)

        # search from the bottom of the window upward, so we prefer prices that
        # appear after the description (typically the actual unit/line total)
        for j in reversed(window_idxs):
            cands = _price_candidates(lines[j])
            val = _choose_rightmost_best(cands, lines[j])
            if val is not None:
                return f"RM{val:.2f}"

    return None


# -----------------------------
# products
# -----------------------------

def _match_preferred_items(lines: List[str], preferred_items: List[str], max_items: int) -> List[Tuple[str, int]]:
    """
    Greedy substring matching against curated item hints (case-insensitive).
    Returns unique items in order of first appearance with best-effort qty detection.
    """
    out: List[Tuple[str, int]] = []
    seen = set()
    pref_up = [(p, p.upper()) for p in preferred_items if isinstance(p, str) and p.strip()]
    if not pref_up:
        return out

    for ln in lines:
        up = ln.upper()
        for original, needle in pref_up:
            if needle in up:
                # normalize quantity if present nearby
                qty = 1
                q = QTY_RE.search(ln)
                if q:
                    for g in q.groups():
                        if g and g.isdigit():
                            qty = int(g); break
                key = (original.strip().casefold(), qty)
                if key in seen:
                    continue
                out.append((original.strip(), qty))
                seen.add(key)
                if len(out) >= max_items:
                    return out
    return out

def extract_products(lines: List[str], max_items: int = 3, preferred_items: Optional[List[str]] = None) -> List[Tuple[str, int]]:
    """
    Product extraction with priority:
      1) Curated item hints (if provided or auto-loaded)
      2) Lines containing 'KHIND' (as product names) + qty detection
      3) Generic product code heuristic (alphanumeric patterns)
    Returns a list of (name_or_code, qty).
    """
    # NEW: default to curated list when caller did not pass one
    if preferred_items is None and PREFERRED_PRODUCT_HINTS:
        preferred_items = PREFERRED_PRODUCT_HINTS

    items: List[Tuple[str, int]] = []

    # 1) Curated hints first (if available)
    if preferred_items:
        items.extend(_match_preferred_items(lines, preferred_items, max_items))
        if len(items) >= max_items:
            return items

    # 2) KHIND lines as product names
    for i, ln in _find_khind_rows(lines):
        name = ln.strip()
        qty = 1
        q = QTY_RE.search(ln)
        if q:
            for g in q.groups():
                if g and g.isdigit():
                    qty = int(g); break
        items.append((name, qty))
        if len(items) >= max_items:
            return items

    # 3) Generic alphanumeric product codes
    for ln in lines:
        if len(items) >= max_items:
            break
        m = PRODUCT_CODE_RE.search(ln)
        if not m:
            continue
        code = m.group(1).strip("-")
        qty = 1
        q = QTY_RE.search(ln)
        if q:
            for g in q.groups():
                if g and g.isdigit():
                    qty = int(g); break
        items.append((code, qty))
    return items

# -----------------------------
# amount spent (KHIND row first, totals fallback)
# -----------------------------

def extract_amount_spent(lines: List[str]) -> Optional[str]:
    if not lines:
        return None

    # 1) strictly the KHIND row price
    amt = _khind_line_amount(lines)
    if amt:
        return amt

    # 2) totals fallback only if KHIND row price not found
    for ln in lines:
        m = RE_KW_LEFT.search(ln) or RE_KW_RIGHT.search(ln)
        if m:
            try:
                val = _to_float(m.group(1))
                if 2.0 <= val <= 100000.0:
                    return f"RM{val:.2f}"
            except Exception:
                pass

    # 3) any currency-tagged amount (max)
    cands = []
    for ln in lines:
        for m in RE_ANY_CCY.finditer(ln):
            try:
                cands.append(_to_float(m.group(1)))
            except Exception:
                pass
    if cands:
        val = max(cands)
        if 2.0 <= val <= 100000.0:
            return f"RM{val:.2f}"

    # 4) last resort: largest naked number anywhere
    bare = []
    for ln in lines:
        for m in RE_ANY_NUM.finditer(ln):
            try:
                bare.append(_to_float(m.group(1)))
            except Exception:
                pass
    if bare:
        val = max(bare)
        if 2.0 <= val <= 100000.0:
            return f"RM{val:.2f}"

    return None

# -----------------------------
# validity decision
# -----------------------------

def decide_validity(amount_spent: Optional[str], products: List[Tuple[str, int]], image_missing: bool):
    if image_missing:
        return "INVALID", "Image missing"
    if amount_spent:
        return "VALID", ""
    if products:
        return "INVALID", "No total found"
    return "INVALID", "No total found"
