# store_loc_map_w4.py
# Put one "STORE <tab> CITY, STATE" per line

RAW_STORE_LOCATIONS = r"""
SK SERIES ENTERPRISE	TAIPING, PERAK
T&L POWER AIRCOND & ELECTRICAL	MUADZAM SHAH, PAHANG
JAYA SUPERSTORE	TUARAN, SABAH
KB LAMPU SDN BHD 	KOTA BELUD, SABAH
PTA FIRST M SDN BHD	BESUT, TERENGGANU
ZBA JAYA SDN BHD	KUALA TERENGGANU, TERENGGANU
VICTORY FURNITURE & ELECTRICAL	JULAU, SARAWAK
HOMEPRO	JALAN TUN RAZAK, MELAKA
CHAI LI COMPANY	BINTULU, SARAWAK
AEON	CHERAS, KUALA LUMPUR
KEMUDI TIMUR	KOTA BHARU, KELANTAN
LIAN HONG TRADING	KOTA BHARU, KELANTAN
SERVEY EVERGREEN	TAWAU, SABAH
YING MIEW HOLDINGS	SRI AMAN, SARAWAK
KHIND MARKETING	SHAH ALAM, SELANGOR
HARVEY NORMAN	PETALING JAYA, SELANGOR
H GLOBAL GROUP INTEGRATED	PASIR MAS, KELANTAN
HLK SDN BHD	KLANG, SELANGOR
SK HARDWARE	KUCHING, SARAWAK
FRENSHI HOME	GELANG PATAH, JOHOR
TOP KINABALU TRADING	KOTA KINABALU, SABAH
IKA HOUSEWARE	BUTTERWORTH, PULAU PINANG
G-ORANGE HOMEMART	KUALA KRAI, KELANTAN
TF VALUEMART	KUALA KANGSAR, PERAK
SMART LIFESTYLE DISTRIBUTOR	SHAH ALAM, SELANGOR
KHIND.COM	SHAH ALAM, SELANGOR
METER TRADING	PUDU, KUALA LUMPUR
KUANG HUAT ELECTRICAL	KAMPAR, PERAK
SURIA JERAI ELECTRICAL	JOHOR BAHRU, JOHOR
SWEE THYE SALES & SERVICES	RAUB, PAHANG
CBH ELECTRICAL	KANGAR, PERLIS
SIMEE ELECTRICAL	IPOH, PERAK
SUN CHING SDN BHD	DURIAN TUNGGAL, MELAKA
BLT ELECTRONICS	KUANTAN, PAHANG
NSK TRADE CENTER	KUCHAI LAMA, KUALA LUMPUR
SYARIKAT WONG ELECTRIC	KEMAMAN, TERENGGANU
SOH ENG SENG TRADING	MUAR, JOHOR
SIN SENG ELECTRICAL	REMBAU, NEGERI SEMBILAN
KOKITAB	KUALA TERENGGANU, TERENGGANU
KP JAYA MARKETING	JOHOR BAHRU, JOHOR
SERVEY EVERGREEN	KENINGAU, SABAH
HOMEPRO	BANDAR JELUTONG, PULAU PINANG
"""

def _norm(s: str) -> str:
    # normalize for robust matching (remove punctuation/extra spaces, uppercase)
    import re
    s = (s or "").upper()
    s = re.sub(r"[\s\W_]+", " ", s).strip()
    return s

def build_store_loc_map(raw: str) -> dict:
    """
    Returns dict like:
      {
        "SK SERIES ENTERPRISE": "TAIPING, PERAK",
        "AEON": "CHERAS, KUALA LUMPUR",
        ...
      }
    Keys are normalized (via _norm); values are canonical as typed (not uppercased).
    If a store appears multiple times with different cities, we keep the first
    occurrence; you can customize as you like (e.g. prefer longer city).
    """
    out = {}
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        # split on tab or >=2 spaces
        parts = [p.strip() for p in (line.split("\t") if "\t" in line else re.split(r"\s{2,}", line)) if p.strip()]
        if len(parts) < 2:
            # try single space split (last token as location)
            parts = line.rsplit(" ", 1)
            if len(parts) < 2:
                continue
        store, loc = parts[0], parts[1]
        key = _norm(store)
        if key not in out and loc != "-":
            out[key] = loc
    return out

STORE_LOC_MAP = build_store_loc_map(RAW_STORE_LOCATIONS)
