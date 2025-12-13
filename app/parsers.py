import re
from typing import List, Tuple, Optional

# ---------------------------------
# OPTIONAL curated lists (auto-load)
# ---------------------------------
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

# Store→Location curated map (optional)
STORE_LOC_MAP: Optional[dict] = None
try:
    from .store_loc_map_w4 import STORE_LOC_MAP as _SLM  # type: ignore
    STORE_LOC_MAP = _SLM
except Exception:
    try:
        from store_loc_map_w4 import STORE_LOC_MAP as _SLM  # type: ignore
        STORE_LOC_MAP = _SLM
    except Exception:
        STORE_LOC_MAP = None

# -----------------------------
# constants / helpers
# -----------------------------

MALAYSIAN_STATES = [
    "Johor","Kedah","Kelantan","Malacca","Melaka","Negeri Sembilan","Pahang","Penang","Pulau Pinang",
    "Perak","Perlis","Sabah","Sarawak","Selangor","Terengganu","Kuala Lumpur","WP Kuala Lumpur",
    "Labuan","Putrajaya"
]

STORE_HINTS = [
    "Khind Customer Service SDN BHD", "AEON BIG", "Jaya Superstore", "HomePro",
    "XingLee Electrical & HIFI Central SDN BHD", "AEON Big", "SK Hardware",
    "SK SERIES ENTERPRISE", "JAYA GROCER", "SNF ONLINE (M) SDN BHD", "Kemudi Timur",
    "ECONSAVE CASH & CARRY (FH) SDN BHD", "SENHENG", "HARVEY NORMAN",
    "Seruay Evergreen SDN BHD", "DYNAPOWER ENTERPRISE",
    "SENG HUAT ELECTRICAL & HOME APPLIANCES SDN BHD", "SURIA JERAI ELECTRICAL SDN BHD",
    "Diva Lighting", "Shopee Online", "LSS HYPER ELECTRICAL SDN BHD",
    "IKA Houseware Malaysia SDN BHD", "YING MIEW HOLDINGS",
    "SYARIKAT LETRIK LIM & ONG SDN BHD", "Jaya",
    "KHIND", "SDN BHD", "ELECTRICAL", "SUPERMARKET", "HYPER", "STORE", "LIGHTING"
]

# numeric patterns
_NUM      = r"(?:\d{1,3}(?:,\d{3})*(?:\.\d{2})|\d+(?:\.\d{2})?)"
_CCY      = r"(?:RM|MYR|R\s*M)"
_SEP      = r"[:=\-\s]*"

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

PRODUCT_CODE_RE = re.compile(r"(?<![A-Z0-9])([A-Z]{2,}[A-Z0-9\-]{1,})(?![A-Z0-9])")
QTY_RE = re.compile(
    r"\b(?:QTY|QTY:|QTY\.|QTY=|QTY\s+)\s*(\d+)\b|\b(\d+)x\b|\bx(\d+)\b|\b(\d+)\s*(?:pcs|unit|units|pcs\.)\b",
    re.IGNORECASE
)

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

def _norm(s: str) -> str:
    s = (s or "").upper()
    return re.sub(r"[\s\W_]+", " ", s).strip()

# -----------------------------
# store name / location
# -----------------------------

def _match_known_store(lines: List[str], preferred_stores: Optional[List[str]] = None) -> Optional[Tuple[str, str]]:
    """
    Try to find a store line using curated store hints (or generic hints),
    and return (store_line_text, normalized_key).
    """
    top = lines[:12]

    # 1) preferred curated list (if available)
    if preferred_stores is None and PREFERRED_STORE_HINTS:
        preferred_stores = PREFERRED_STORE_HINTS

    if preferred_stores:
        pref_up = [s.upper() for s in preferred_stores]
        for ln in top:
            up = ln.upper()
            if any(s in up for s in pref_up):
                return ln.strip(), _norm(ln)

    # 2) generic hints
    for ln in top:
        if _contains_store_hint(ln):
            return ln.strip(), _norm(ln)

    # 3) ALLCAPS fallback
    for ln in top:
        if ln.strip() and ln.strip().upper() == ln.strip():
            return ln.strip(), _norm(ln)

    # 4) first non-empty
    for ln in top:
        if ln.strip():
            return ln.strip(), _norm(ln)

    return None

def extract_store_name(lines: List[str], preferred_stores: Optional[List[str]] = None) -> Optional[str]:
    m = _match_known_store(lines, preferred_stores)
    return m[0] if m else None

# ---- Receipt-only location helpers + curated map check ----

_POSTCODE_RE = re.compile(r"\b(\d{5})\b")
_STATE_SET = {s.lower(): s for s in MALAYSIAN_STATES}

def _extract_city_state_from_line(line: str) -> Optional[str]:
    txt = " ".join(line.replace("|", ",").split())
    low = txt.lower()
    state_norm = None
    state_key = None
    for k, v in _STATE_SET.items():
        if k in low:
            state_norm = v; state_key = k; break
    if not state_norm:
        return None

    m_pc = _POSTCODE_RE.search(txt)
    if m_pc:
        left_right = low.split(state_key, 1)[0]
        after_pc = left_right.split(m_pc.group(1), 1)[-1]
        city_tokens = re.sub(r"[^A-Za-z\s]", " ", after_pc).split()
        city_tokens = [t for t in city_tokens if len(t) >= 2]
        if city_tokens:
            city = " ".join(city_tokens).strip().title()
            return f"{city}, {state_norm}"

    parts = [p.strip() for p in re.split(r"[,;|-]", txt) if p.strip()]
    for i, seg in enumerate(parts):
        if state_key and state_key in seg.lower():
            if i > 0:
                city = re.sub(r"[^A-Za-z\s]", " ", parts[i-1]).strip().title()
                if city:
                    return f"{city}, {state_norm}"
            return state_norm
    return state_norm

def extract_store_location(lines: List[str], fallback_city: Optional[str], fallback_state: Optional[str]) -> Optional[str]:
    """
    Priority:
      1) If we recognized a store and it matches a curated map entry -> return mapped location
      2) Else parse from receipt text (postcodes / 'City, State' / state-only)
      3) Never use participant's fallback city/state (by design)
    """
    # 1) curated map by store name
    store_match = _match_known_store(lines)
    if store_match and STORE_LOC_MAP:
        _, norm_key = store_match
        # try full-line key, else try each known key as a substring of the line
        # (normalize both sides)
        if norm_key in STORE_LOC_MAP:
            return STORE_LOC_MAP[norm_key]
        # substring match: if any curated key is contained in the line
        for k, loc in STORE_LOC_MAP.items():
            if k and k in norm_key:
                return loc

    # 2) receipt text heuristics
    for ln in lines[:20]:
        got = _extract_city_state_from_line(ln)
        if got:
            return got
    for ln in reversed(lines):
        got = _extract_city_state_from_line(ln)
        if got:
            return got

    # 3) no fallbacks to participant location
    return None

# -----------------------------
# KHIND row → price (multi-line aware)
# -----------------------------

def _price_candidates(line: str) -> List[Tuple[int, float, int]]:
    hits: List[Tuple[int, float, int]] = []
    for m in PRICE_TOKEN_RE.finditer(line):
        if m.group("ccy"):
            val = _to_float(m.group("ccyval"))
            hits.append((m.start("ccyval"), val, 3))
        elif m.group("dec"):
            val = _to_float(m.group("dec"))
            hits.append((m.start("dec"), val, 2))
    hits.sort(key=lambda x: x[0])
    return hits

def _find_khind_rows(lines: List[str]) -> List[Tuple[int, str]]:
    return [(i, ln) for i, ln in enumerate(lines) if "KHIND" in ln.upper()]

def _looks_like_qty_context(line: str, span_start: int, span_end: int) -> bool:
    left = max(0, span_start - 10); right = min(len(line), span_end + 10)
    return bool(_QTY_WORDS_RE.search(line[left:right]))

def _choose_rightmost_best(cands: List[Tuple[int, float, int]], line: str) -> Optional[float]:
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

_STOP_AFTER_RE = re.compile(r"\b(total|grand\s*total|balance|remarks?|thank|cash\s*rm?)\b", re.I)

def _khind_line_amount(lines: List[str]) -> Optional[str]:
    LOOKAHEAD = 4
    for idx, ln in _find_khind_rows(lines):
        window_idxs = []
        for j in range(idx, min(len(lines), idx + LOOKAHEAD + 1)):
            if _STOP_AFTER_RE.search(lines[j]):
                break
            window_idxs.append(j)
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
    out: List[Tuple[str, int]] = []
    seen = set()
    pref_up = [(p, p.upper()) for p in preferred_items if isinstance(p, str) and p.strip()]
    if not pref_up:
        return out
    for ln in lines:
        up = ln.upper()
        for original, needle in pref_up:
            if needle in up:
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
    if preferred_items is None and PREFERRED_PRODUCT_HINTS:
        preferred_items = PREFERRED_PRODUCT_HINTS

    items: List[Tuple[str, int]] = []

    if preferred_items:
        items.extend(_match_preferred_items(lines, preferred_items, max_items))
        if len(items) >= max_items:
            return items

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

    amt = _khind_line_amount(lines)
    if amt:
        return amt

    for ln in lines:
        m = RE_KW_LEFT.search(ln) or RE_KW_RIGHT.search(ln)
        if m:
            try:
                val = _to_float(m.group(1))
                if 2.0 <= val <= 100000.0:
                    return f"RM{val:.2f}"
            except Exception:
                pass

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
