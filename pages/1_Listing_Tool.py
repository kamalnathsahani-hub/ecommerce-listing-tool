import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import os
import re
import json
import time
import zipfile
import base64
import hashlib
import difflib
import traceback
import requests
from datetime import datetime
from PIL import Image

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")

if st.button("⬅️ Back to Home"):
    st.switch_page("app.py")
st.title("📦 Multi-Marketplace Auto Listing Generator (Universal Pro V24 - Unified Dashboard)")
st.caption("Ajio | Amazon VC | VC➜SC | Amazon SC | Flipkart | Meesho | Myntra + 🤖 AI Vision (Gemini ➜ NVIDIA ➜ Groq Fallback)")

# ======================================================================
#                🔑 API KEYS & MULTI-AI VISION CONFIG
# ======================================================================
IMGBB_PERMANENT_KEY = "137be55d19cf3d6a7ec051e369137041"

# 1️⃣ Primary AI: Google Gemini Vision
GEMINI_API_KEY = "AQ.Ab8RN6Lb4kCeCRQtjVtw5MYcz7Wl0x4dBRvNxsYCAD9TYoUk2Q"
GEMINI_VISION_MODEL = "gemini-2.0-flash"

# 2️⃣ Secondary AI: NVIDIA Vision
NVIDIA_API_KEY = "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"
NVIDIA_MODEL = "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning"

# 3️⃣ Tertiary AI: Groq Vision
GROQ_API_KEY = "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"
GROQ_VISION_MODEL = "llama-3.2-11b-vision-preview"

# ======================================================================
# 1. MARKETPLACE SELECTION
# ======================================================================
col_m1, col_m2 = st.columns([2, 1])
with col_m1:
    marketplace = st.selectbox(
        "Marketplace Chunein:",
        [
            "Ajio",
            "Amazon - Vendor Central",
            "Convert: Vendor Central ➜ Seller Central",
            "Amazon - Seller Central",
            "Flipkart",
            "Meesho",
            "Myntra",
        ]
    )
with col_m2:
    default_td = st.number_input(
        "TD / Default Margin (%) [Ajio ke liye]:",
        min_value=0, max_value=100, value=50,
        help="Ajio template ke TD column me ye value jayegi."
    )

st.divider()
CONVERT_MODE = marketplace.startswith("Convert:")

# ======================================================================
# GLOBAL STATE
# ======================================================================
detected_folders = []
folder_images_count = {}
zip_images_data = {}
detected_size_chart_info = {}
if 'img_cache' not in st.session_state:
    st.session_state['img_cache'] = {}
if 'category_overrides' not in st.session_state:
    st.session_state['category_overrides'] = {}
if 'gender_overrides' not in st.session_state:
    st.session_state['gender_overrides'] = {}
if 'vision_results' not in st.session_state:
    st.session_state['vision_results'] = {}
img_cache = st.session_state['img_cache']

# ======================================================================
# COMMON HELPERS
# ======================================================================
def clean_sku_name(name):
    cleaned = re.sub(r'[\s_]+', '-', str(name).strip())
    cleaned = re.sub(r'-+', '-', cleaned).strip('-')
    return cleaned

def is_size_chart_file(filename):
    low = os.path.basename(filename).lower()
    patterns = ['size_chart', 'size-chart', 'sizechart', 'size_guide', 'sizeguide', 'size chart', 'chart', 'measurement']
    return any(p in low for p in patterns) and low.endswith(('.jpg', '.jpeg', '.png', '.webp'))

def _norm_h(s):
    return re.sub(r'[\s_\-*]+', '', str(s or '')).strip().lower()

CATEGORY_OPTIONS = [
    'co-ord set', 'hoodie', 'sweatshirt', 'cargo', 'jeans', 'tshirt',
    'shirt', 'trousers', 'joggers', 'shorts', 'jacket', 'kurta', 'kurti',
    'dress', 'top', 'skirt', 'apparel'
]

def detect_product_category(folder_name, user_keywords="", sheet_name=""):
    combined = f"{folder_name} {user_keywords} {sheet_name}".lower()
    if any(k in combined for k in ["co-ord", "coord", "two piece", "suit set", "apparel set", "clothing set"]):
        return "co-ord set"
    elif any(k in combined for k in ["hoodie", "hoody", "hooded"]):
        return "hoodie"
    elif any(k in combined for k in ["sweatshirt", "sweater", "pullover", "cardigan"]):
        return "sweatshirt"
    elif any(k in combined for k in ["cargo"]):
        return "cargo"
    elif any(k in combined for k in ["jean", "denim"]):
        return "jeans"
    elif any(k in combined for k in ["t-shirt", "tshirt", "tee", "polo"]):
        return "tshirt"
    elif any(k in combined for k in ["kurti", "tunic"]):
        return "kurti"
    elif any(k in combined for k in ["kurta", "ethnic", "salwar"]):
        return "kurta"
    elif any(k in combined for k in ["dress", "frock", "gown", "jumpsuit"]):
        return "dress"
    elif any(k in combined for k in ["top", "blouse", "camisole"]):
        return "top"
    elif any(k in combined for k in ["shirt"]):
        return "shirt"
    elif any(k in combined for k in ["trouser", "chinos", "pant", "bottom", "legging", "palazzo"]):
        return "trousers"
    elif any(k in combined for k in ["jacket", "coat", "blazer", "windcheater", "shrug"]):
        return "jacket"
    elif any(k in combined for k in ["track", "jogger", "sweatpant", "pyjama", "pajama"]):
        return "joggers"
    elif any(k in combined for k in ["short", "boxer"]):
        return "shorts"
    elif any(k in combined for k in ["skirt"]):
        return "skirt"
    if "pant" in sheet_name.lower():
        return "trousers"
    elif "shirt" in sheet_name.lower():
        return "shirt"
    elif "coord" in sheet_name.lower() or "co-ord" in sheet_name.lower():
        return "co-ord set"
    return "apparel"

def detect_gender(text):
    t = str(text).lower()
    if any(w in t for w in ["women", "woman", "female", "girl", "ladies", "kurti", "co-ord", "coord", "dress", "top", "skirt"]):
        return "Women"
    if any(w in t for w in ["men", "man", "male", "boy", "gents"]):
        return "Men"
    return "Women"

def resolve_category(clean_style, ctx, sheet_name=""):
    ov = st.session_state.get('category_overrides', {})
    if clean_style in ov and ov[clean_style]:
        return ov[clean_style]
    vr = ctx.get('vision_results') or {}
    if (vr.get(clean_style) or {}).get('category'):
        return vr[clean_style]['category']
    cat = detect_product_category(clean_style, ctx.get('keywords', ''), sheet_name)
    if cat == 'apparel' and ctx.get('default_category'):
        return ctx['default_category']
    return cat

def resolve_gender(clean_style, ctx):
    ov = st.session_state.get('gender_overrides', {})
    if clean_style in ov and ov[clean_style]:
        return ov[clean_style]
    vr = ctx.get('vision_results') or {}
    if (vr.get(clean_style) or {}).get('gender'):
        return vr[clean_style]['gender']
    return detect_gender(f"{clean_style} {ctx.get('keywords', '')}")

# ======================================================================
# SMART FABRIC & SLEEVE RESOLVERS
# ======================================================================
DEFAULT_CATEGORY_FABRICS = {
    'jeans': 'Denim',
    'hoodie': 'Fleece',
    'sweatshirt': 'Fleece',
    'tshirt': 'Cotton',
    'shirt': 'Cotton',
    'cargo': 'Cotton',
    'trousers': 'Cotton Blend',
    'joggers': 'Cotton Blend',
    'shorts': 'Cotton',
    'co-ord set': 'Cotton Blend',
    'kurti': 'Rayon',
    'kurta': 'Cotton',
    'dress': 'Crepe',
    'top': 'Cotton Blend',
    'jacket': 'Polyester',
    'skirt': 'Cotton Blend',
    'apparel': 'Cotton Blend'
}

def resolve_fabric(clean_style, cat, ctx):
    vr = ctx.get('vision_results') or {}
    ai_mat = (vr.get(clean_style) or {}).get('material')
    if ai_mat and ai_mat.strip() and (ai_mat.strip().lower() != 'denim' or cat == 'jeans'):
        return ai_mat.strip().title()
    if cat == 'jeans':
        return "Denim"
    return DEFAULT_CATEGORY_FABRICS.get(cat, "Cotton Blend")

def resolve_sleeve_length(clean_style, cat, ctx):
    vr = ctx.get('vision_results') or {}
    ai_sleeve = (vr.get(clean_style) or {}).get('sleeve_length')
    if ai_sleeve and ai_sleeve.strip():
        return ai_sleeve.strip().title()
    if cat in ['hoodie', 'sweatshirt', 'jacket']:
        return "Long Sleeve"
    elif cat in ['tshirt', 'shirt', 'kurti', 'dress', 'top']:
        return "Short Sleeve"
    return "Sleeveless"

# ======================================================================
# COLOR DETECTION
# ======================================================================
COMMON_COLORS = [
    "Beige", "LightGrey", "Grey", "DarkGrey", "Black", "Navy", "Blue", "LightBlue", "DarkBlue",
    "IceBlue", "White", "Off White", "Green", "Olive", "Brown", "Coffee", "Charcoal", "Khaki",
    "Red", "Maroon", "Yellow", "Orange", "Pink", "Purple", "Lavender", "Wine", "Peach", "Teal"
]

STANDARD_COLOR_MAPS = [
    'Beige', 'Black', 'Blue', 'Bronze', 'Brown', 'Clear', 'Gold', 'Green', 'Grey',
    'Metallic', 'Multicolor', 'Off White', 'Orange', 'Pink', 'Purple', 'Red', 'Silver',
    'Turquoise', 'White', 'Yellow'
]

COLOR_MISSPELLINGS = {
    'maron': 'Maroon', 'marron': 'Maroon', 'marun': 'Maroon',
    'mahroon': 'Maroon', 'mehroon': 'Maroon', 'meroon': 'Maroon',
    'greay': 'Grey', 'gray': 'Grey', 'gery': 'Grey', 'grrey': 'Grey',
    'begie': 'Beige', 'baige': 'Beige', 'beigh': 'Beige',
    'blak': 'Black', 'blackk': 'Black',
    'wite': 'White', 'whit': 'White', 'whiet': 'White',
    'offwhite': 'Off White', 'ofwhite': 'Off White',
    'blu': 'Blue', 'bleu': 'Blue', 'navi': 'Navy',
    'chochlate': 'Brown', 'cofee': 'Coffee',
    'yello': 'Yellow', 'orenge': 'Orange', 'orang': 'Orange',
    'pinke': 'Pink', 'gren': 'Green', 'oliv': 'Olive',
}

def detect_color(text):
    text_clean = str(text).replace('-', ' ').replace('_', ' ').lower()
    words = [w for w in text_clean.split() if w]
    joined = ''.join(words)
    for col in sorted(COMMON_COLORS, key=len, reverse=True):
        cl = col.lower()
        if cl in text_clean or cl.replace(' ', '') in joined:
            return col
    for key in sorted(COLOR_MISSPELLINGS, key=len, reverse=True):
        if key in words or key == joined:
            return COLOR_MISSPELLINGS[key]
    color_lowers = [c.lower() for c in COMMON_COLORS]
    for w in words:
        if len(w) < 4:
            continue
        close = difflib.get_close_matches(w, color_lowers, n=1, cutoff=0.75)
        if close:
            for col in COMMON_COLORS:
                if col.lower() == close[0]:
                    return col
    return "Black"

def resolve_color(raw_fld, clean_style, ctx):
    vr = ctx.get('vision_results') or {}
    if (vr.get(clean_style) or {}).get('color'):
        return vr[clean_style]['color']
    return detect_color(raw_fld)

COLOR_FAMILY_MAP = {
    'maroon': 'Red', 'mehroon': 'Red', 'red': 'Red', 'wine': 'Red',
    'burgundy': 'Red', 'rust': 'Red', 'crimson': 'Red', 'cherry': 'Red', 'darkred': 'Red',
    'navy': 'Blue', 'navyblue': 'Blue', 'blue': 'Blue', 'lightblue': 'Blue',
    'darkblue': 'Blue', 'iceblue': 'Blue', 'skyblue': 'Blue', 'royalblue': 'Blue', 'teal': 'Blue',
    'turquoise': 'Turquoise',
    'black': 'Black', 'charcoal': 'Grey', 'grey': 'Grey', 'gray': 'Grey',
    'lightgrey': 'Grey', 'darkgrey': 'Grey', 'smoke': 'Grey', 'silver': 'Silver',
    'white': 'White', 'offwhite': 'Off White', 'cream': 'White', 'ivory': 'White',
    'green': 'Green', 'olive': 'Green', 'mint': 'Green', 'darkgreen': 'Green',
    'lightgreen': 'Green', 'parrotgreen': 'Green',
    'brown': 'Brown', 'coffee': 'Brown', 'tan': 'Brown', 'chocolate': 'Brown',
    'camel': 'Brown', 'darkbrown': 'Brown', 'lightbrown': 'Brown',
    'beige': 'Beige', 'khaki': 'Beige', 'sand': 'Beige',
    'yellow': 'Yellow', 'mustard': 'Yellow', 'orange': 'Orange', 'peach': 'Orange',
    'pink': 'Pink', 'rose': 'Pink', 'lightpink': 'Pink',
    'purple': 'Purple', 'lavender': 'Purple', 'violet': 'Purple',
    'gold': 'Gold', 'bronze': 'Bronze', 'metallic': 'Metallic',
}

def _canonicalize_color(name):
    if not name:
        return name
    raw = str(name).strip().lower().replace('-', ' ').replace('_', ' ')
    words = raw.split()
    joined = ''.join(words)
    for key in sorted(COLOR_MISSPELLINGS, key=len, reverse=True):
        if key == joined or key in words:
            return COLOR_MISSPELLINGS[key]
    for w in words:
        if len(w) >= 4:
            close = difflib.get_close_matches(w, [c.lower() for c in COMMON_COLORS], n=1, cutoff=0.75)
            if close:
                for col in COMMON_COLORS:
                    if col.lower() == close[0]:
                        return col
    return name

def map_to_standard_color_map(det_col):
    if not det_col:
        return "Black"
    det_col = _canonicalize_color(det_col)
    raw = str(det_col).strip().lower().replace('-', ' ').replace('_', ' ')
    words = raw.split()
    joined = ''.join(words)
    for sc in STANDARD_COLOR_MAPS:
        if sc.lower().replace(' ', '') == joined:
            return sc
    for key in sorted(COLOR_FAMILY_MAP, key=len, reverse=True):
        if key == joined or key in words:
            return COLOR_FAMILY_MAP[key]
    return "Multicolor"

# ======================================================================
# POCKET MAPPING
# ======================================================================
POCKET_BY_CATEGORY = {
    'hoodie': 'Kangaroo Pocket', 'sweatshirt': 'Kangaroo Pocket',
    'cargo': 'Cargo Pocket', 'jeans': 'Five Pocket',
    'trousers': 'Slash Pocket', 'pants': 'Slash Pocket',
    'joggers': 'Side Pocket', 'shorts': 'Side Pocket',
    'jacket': 'Welt Pocket', 'shirt': 'Chest Pocket',
    'tshirt': 'No Pocket', 'kurta': 'Side Pocket', 'kurti': 'No Pocket',
    'co-ord set': 'Side Pocket', 'dress': 'No Pocket', 'top': 'No Pocket',
    'skirt': 'No Pocket', 'apparel': 'Side Pocket',
}

def get_pocket_for_cat(cat):
    return POCKET_BY_CATEGORY.get(cat, 'Side Pocket')

def resolve_pocket(clean_style, cat, ctx):
    vr = ctx.get('vision_results') or {}
    if (vr.get(clean_style) or {}).get('pocket'):
        return vr[clean_style]['pocket']
    return get_pocket_for_cat(cat)

CATEGORY_LABELS = {
    'co-ord set': 'Co-ord Set', 'kurti': 'Kurti', 'kurta': 'Kurta', 'dress': 'Dress',
    'top': 'Top', 'hoodie': 'Hoodie', 'sweatshirt': 'Sweatshirt', 'cargo': 'Cargo Pants',
    'jeans': 'Jeans', 'tshirt': 'T-Shirt', 'shirt': 'Shirt', 'trousers': 'Trousers',
    'jacket': 'Jacket', 'joggers': 'Joggers', 'shorts': 'Shorts', 'pants': 'Pants',
    'skirt': 'Skirt', 'apparel': 'Casual Wear',
}

HSN_CODE_MAP = {
    'co-ord set': '62042200', 'kurti': '62042200', 'dress': '62044200', 'top': '62063000',
    'jeans': '62034290', 'cargo': '62034200', 'trousers': '62034200', 'pants': '62034200',
    'joggers': '62034200', 'shorts': '62034200', 'tshirt': '61099090', 'shirt': '62052000',
    'hoodie': '61103000', 'sweatshirt': '61103000', 'jacket': '62014090', 'kurta': '62034200',
    'skirt': '62045200',
}

def get_hsn_code(cat):
    return HSN_CODE_MAP.get(cat, '62034200')

def get_item_type_name_from_dropdown(cat):
    mapping = {
        'co-ord set': 'Outfit Set', 'kurti': 'Tunic', 'dress': 'Dress', 'top': 'Top',
        'hoodie': 'Hooded Sweatshirt', 'sweatshirt': 'Sweatshirt', 'joggers': 'Sweatpants',
        'jeans': 'Jeans', 'cargo': 'Cargo Pants', 'trousers': 'Casual Pants', 'pants': 'Casual Pants',
        'shorts': 'Shorts', 'tshirt': 'T-Shirt', 'shirt': 'Casual Shirt', 'kurta': 'Kurta',
        'jacket': 'Jacket', 'skirt': 'Skirt', 'apparel': 'Casual Pants',
    }
    return mapping.get(cat, 'Casual Pants')

def get_style_from_dropdown(cat):
    mapping = {
        'co-ord set': 'Co-ords', 'kurti': 'A-Line', 'dress': 'Fit & Flare', 'top': 'Regular',
        'hoodie': 'Pullover', 'sweatshirt': 'Pullover', 'joggers': 'Jogger',
        'cargo': 'Cargo', 'trousers': 'Chino', 'jeans': 'Classic',
    }
    return mapping.get(cat, 'Modern')

SC_PRODUCT_TYPE = {
    'co-ord set': 'SUIT', 'kurti': 'KURTA', 'dress': 'DRESS', 'top': 'SHIRT',
    'hoodie': 'SWEATSHIRT', 'sweatshirt': 'SWEATSHIRT', 'jacket': 'COAT',
    'jeans': 'PANTS', 'cargo': 'PANTS', 'trousers': 'PANTS', 'pants': 'PANTS',
    'joggers': 'PANTS', 'shorts': 'SHORTS', 'tshirt': 'SHIRT', 'shirt': 'SHIRT',
    'kurta': 'KURTA', 'skirt': 'SKIRT', 'apparel': 'SHIRT',
}

# Candidate priority lists per category for matching Seller Central Product Type dropdowns
SC_PRODUCT_TYPE_CANDIDATES = {
    'co-ord set': ['SUIT', 'APPAREL_SET', 'COORDINATED_OUTFIT', 'OUTFIT_SET', 'SALWAR_SUIT_SET', 'TRACK_SUIT', 'DRESS', 'KURTA', 'SHIRT', 'TOP'],
    'kurti': ['KURTA', 'TUNIC', 'SHIRT', 'TOP', 'DRESS', 'SALWAR_SUIT_SET'],
    'kurta': ['KURTA', 'TUNIC', 'SALWAR_SUIT_SET', 'SHIRT'],
    'dress': ['DRESS', 'ONE_PIECE_OUTFIT', 'JUMPSUIT', 'OVERALLS', 'TUNIC', 'SUIT', 'GOWN'],
    'top': ['SHIRT', 'TOP', 'TUNIC', 'BLOUSE', 'SWEATSHIRT', 'KURTA'],
    'tshirt': ['SHIRT', 'T_SHIRT', 'TSHIRT', 'TOP', 'SWEATSHIRT'],
    'shirt': ['SHIRT', 'TOP', 'TUNIC', 'KURTA'],
    'hoodie': ['SWEATSHIRT', 'HOODIE', 'SWEATER', 'PULLOVER', 'COAT', 'JACKET', 'SHIRT'],
    'sweatshirt': ['SWEATSHIRT', 'SWEATER', 'PULLOVER', 'HOODIE', 'COAT', 'SHIRT'],
    'jacket': ['COAT', 'JACKET', 'BLAZER', 'OUTERWEAR', 'SWEATSHIRT', 'VEST'],
    'jeans': ['PANTS', 'JEANS', 'TROUSERS', 'BOTTOMS', 'OVERALLS', 'SHORTS'],
    'cargo': ['PANTS', 'TROUSERS', 'CARGO_PANTS', 'BOTTOMS', 'JEANS', 'SHORTS'],
    'trousers': ['PANTS', 'TROUSERS', 'BOTTOMS', 'LEGGINGS', 'JEANS', 'SUIT'],
    'pants': ['PANTS', 'TROUSERS', 'BOTTOMS', 'LEGGINGS', 'JEANS'],
    'joggers': ['PANTS', 'SWEATPANTS', 'TRACK_PANTS', 'JOGGERS', 'TROUSERS', 'TRACK_SUIT', 'PAJAMAS', 'SHORTS'],
    'shorts': ['SHORTS', 'PANTS', 'BOTTOMS', 'UNDERPANTS'],
    'skirt': ['SKIRT', 'DRESS', 'PANTS', 'BOTTOMS'],
    'apparel': ['SHIRT', 'PANTS', 'SWEATSHIRT', 'KURTA', 'DRESS', 'SUIT', 'COAT', 'SHORTS', 'TUNIC', 'TOP'],
}

def resolve_sc_product_type_for_row(ws_sc, sc_prod_type_col, raw_prod_type="", sheet_name="", item_type="", model_name="", item_name=""):
    """
    Resolves the best Seller Central Product Type using:
    1. Existing VC Product Type (if present in the VC row)
    2. VC Sheet Name (e.g. Jeans, Shirt, Sweatshirt, Co-ord, Kurti, etc.)
    3. Row attributes (Item Type Name, Model Name, Item Name)
    And matches it against the actual dropdown options in the blank Seller Central template.
    """
    options = get_dropdown_options_for_column(ws_sc, sc_prod_type_col) if sc_prod_type_col else None
    norm_opts = {}
    if options:
        for o in options:
            no = re.sub(r'[\s_\-]+', '', str(o).strip().upper())
            if no and no not in norm_opts:
                norm_opts[no] = o

    # 1. If VC row already had an explicit Product Type, try matching it directly first
    if raw_prod_type and str(raw_prod_type).strip():
        rp = str(raw_prod_type).strip()
        if not options:
            return rp
        nrp = re.sub(r'[\s_\-]+', '', rp.upper())
        if nrp in norm_opts:
            return norm_opts[nrp]

    # 2. Detect category from VC sheet name + item metadata
    cat = detect_product_category(
        folder_name=f"{sheet_name} {item_type} {model_name}",
        user_keywords=str(item_name or ""),
        sheet_name=str(sheet_name or "")
    )
    candidates = list(SC_PRODUCT_TYPE_CANDIDATES.get(cat, SC_PRODUCT_TYPE_CANDIDATES['apparel']))
    default_pt = SC_PRODUCT_TYPE.get(cat, 'SHIRT')
    if default_pt not in candidates:
        candidates.insert(0, default_pt)

    # Also check if sheet_name itself directly matches one of the dropdown options
    sheet_clean = re.sub(r'[\s_\-]+', '', str(sheet_name or '').strip().upper())
    if options:
        if sheet_clean:
            if sheet_clean in norm_opts:
                return norm_opts[sheet_clean]
            # Singular/plural check on sheet name
            if sheet_clean.rstrip('S') in norm_opts:
                return norm_opts[sheet_clean.rstrip('S')]
            if (sheet_clean + 'S') in norm_opts:
                return norm_opts[sheet_clean + 'S']

        # Check candidate list in priority order against available dropdown options
        for cand in candidates:
            nc = re.sub(r'[\s_\-]+', '', cand.upper())
            if nc in norm_opts:
                return norm_opts[nc]
            # Partial substring match against dropdown options
            for no, orig_o in norm_opts.items():
                if nc in no or no in nc:
                    return orig_o

        # Fuzzy fallback via pick_dropdown_value
        return pick_dropdown_value(ws_sc, sc_prod_type_col, default_pt, options[0])

    return default_pt

def make_workable_dropbox_url(base_url, sku, img_num):
    if not base_url:
        return f"https://www.dropbox.com/s/{sku}/{img_num}.jpg?raw=1"
    clean = base_url.split("?")[0].rstrip("/")
    return f"{clean}/{sku}/{img_num}.jpg?raw=1"

def upload_to_imgbb(image_bytes, api_key):
    try:
        url = "https://api.imgbb.com/1/upload"
        payload = {"key": api_key, "image": base64.b64encode(image_bytes).decode('utf-8')}
        res = requests.post(url, data=payload, timeout=20)
        if res.status_code == 200:
            return res.json()['data']['url']
    except Exception:
        pass
    return None

def upload_to_imgbb_cached(image_bytes, api_key, cache):
    h = hashlib.md5(image_bytes).hexdigest()
    if h in cache:
        return cache[h]
    url = upload_to_imgbb(image_bytes, api_key)
    if url:
        cache[h] = url
    return url

# ======================================================================
# DROPDOWN-SAFETY ENGINE (FUZZY DROPDOWN MATCHER)
# ======================================================================
def _resolve_validation_formula(ws, f1):
    f1 = str(f1).strip()
    if f1.startswith('"') and f1.endswith('"'):
        return [o.strip() for o in f1[1:-1].split(',') if o.strip()]
    ref = f1[1:] if f1.startswith('=') else f1
    if '!' not in ref and ':' not in ref:
        dn = None
        try:
            dns = ws.parent.defined_names
            try:
                dn = dns.get(ref)
            except Exception:
                dn = dns[ref] if ref in dns else None
        except Exception:
            dn = None
        if dn is not None:
            ref = getattr(dn, 'attr_text', '') or getattr(dn, 'value', '') or ''
            ref = ref.split(',')[0].strip().lstrip('=')
    if '!' in ref:
        try:
            sheet_part, rng = ref.rsplit('!', 1)
            target = ws.parent[sheet_part.strip("'")]
        except Exception:
            return []
    else:
        target, rng = ws, ref
    if ':' not in rng:
        return []
    try:
        cells = target[rng.replace('$', '')]
        vals = []
        for row in cells:
            if not isinstance(row, tuple):
                row = (row,)
            for c in row:
                if c.value not in (None, ''):
                    vals.append(str(c.value).strip())
        return vals
    except Exception:
        return []

def get_dropdown_options_for_column(ws, col_idx):
    try:
        dvs = ws.data_validations.dataValidation
    except Exception:
        dvs = []
    options = None
    for dv in dvs:
        hit = False
        try:
            for rng in dv.sqref.ranges:
                if rng.min_col <= col_idx <= rng.max_col:
                    hit = True
                    break
        except Exception:
            hit = True
        if not hit:
            continue
        f1 = (dv.formula1 or '').strip()
        if not f1:
            continue
        opts = _resolve_validation_formula(ws, f1)
        if opts:
            options = opts
            break
    if not options and hasattr(ws, 'parent') and ws.parent and 'Dropdown Lists' in ws.parent.sheetnames:
        try:
            ws_dd = ws.parent['Dropdown Lists']
            col_keys = {
                str(ws.cell(r, col_idx).value or '').strip().lower()
                for r in (1, 2, 3, 4, 5)
                if ws.cell(r, col_idx).value not in (None, '')
            }
            for dd_c in range(1, ws_dd.max_column + 1):
                header_hit_row = None
                for hr in (1, 2, 3):
                    dd_k = str(ws_dd.cell(hr, dd_c).value or '').strip().lower()
                    if dd_k and dd_k in col_keys:
                        header_hit_row = hr
                        break
                if header_hit_row is not None:
                    vals = [
                        str(ws_dd.cell(r, dd_c).value).strip()
                        for r in range(header_hit_row + 1, ws_dd.max_row + 1)
                        if ws_dd.cell(r, dd_c).value not in (None, '')
                    ]
                    if vals:
                        options = vals
                        break
        except Exception:
            pass
    if options:
        junk = {'', 'select', 'select one', 'choose', 'choose one', 'dropdown'}
        options = [o for o in options if str(o).strip().lower() not in junk]
        return options or None
    return None

def pick_dropdown_value(ws, col_idx, desired, fallback=None):
    options = get_dropdown_options_for_column(ws, col_idx)
    if not options:
        return desired if desired is not None else fallback
    d = str(desired if desired is not None else '').strip()
    fl = str(fallback if fallback is not None else '').strip()
    if not d and not fl:
        return None
    cand = d or fl
    cl = cand.lower()
    for o in options:
        if str(o).strip().lower() == cl:
            return o
    for o in options:
        ol = str(o).strip().lower()
        if cl in ol or ol in cl:
            return o
    close = difflib.get_close_matches(cand, options, n=1, cutoff=0.35)
    if close:
        return close[0]
    return options[0]

# ======================================================================
# 🤖 MULTI-AI VISION ENGINE (Gemini ➜ NVIDIA ➜ Groq)
# ======================================================================
KNOWN_POCKETS = [
    'Kangaroo Pocket', 'Cargo Pocket', 'Five Pocket', 'Side Pocket', 'No Pocket',
    'Welt Pocket', 'Chest Pocket', 'Slash Pocket', 'Flap Pocket', 'Utility Pocket',
]

AI_CAT_MAP = {
    'co-ord': 'co-ord set', 'coord': 'co-ord set', 'cord set': 'co-ord set', 'two piece': 'co-ord set',
    'hooded sweatshirt': 'hoodie', 'hooded': 'hoodie', 'hood': 'hoodie', 'pullover': 'hoodie',
    'sweat shirt': 'sweatshirt', 'sweater': 'sweatshirt', 'crewneck': 'sweatshirt',
    't-shirt': 'tshirt', 'tshirt': 'tshirt', 'tee': 'tshirt', 'polo': 'tshirt',
    'denim': 'jeans', 'denim jeans': 'jeans', 'denim pants': 'jeans',
    'pants': 'trousers', 'chinos': 'trousers', 'formal pants': 'trousers',
    'track pants': 'joggers', 'trackpants': 'joggers', 'sweatpants': 'joggers',
    'kurtis': 'kurti', 'kurtha': 'kurta', 'dresses': 'dress', 'tops': 'top',
}

def guess_mime(filename):
    low = str(filename or '').lower()
    if low.endswith('.png'):
        return 'image/png'
    if low.endswith('.webp'):
        return 'image/webp'
    return 'image/jpeg'

def extract_json_from_text(text):
    if not text:
        return None
    try:
        m = re.search(r'\{[\s\S]*\}', text)
        if m:
            return json.loads(m.group(0))
    except Exception:
        pass
    return None

def get_vision_prompt():
    cats = ", ".join(CATEGORY_OPTIONS)
    return (
        "You are an expert e-commerce fashion cataloger. Visually inspect this garment image very carefully and output ONLY a valid JSON object without markdown or code formatting.\n"
        "Return this exact schema:\n"
        "{\n"
        '  "category": "<One of: ' + cats + '>",\n'
        '  "gender": "<Women or Men>",\n'
        '  "color": "<Visible dominant garment color: Pink, Maroon, Black, Navy, Grey, Beige, Olive, Khaki, White, Off White, Lavender, Blue, Green, etc.>",\n'
        '  "pocket": "<Kangaroo Pocket / Cargo Pocket / Five Pocket / Side Pocket / Welt Pocket / Chest Pocket / Slash Pocket / No Pocket>",\n'
        '  "material": "<Visually detected actual fabric/fabric type: Cotton, Cotton Blend, Fleece, Rayon, Polyester, Denim, Crepe, Georgette, Silk, Linen, Knit, etc. Only use Denim if it is real denim jeans/jacket. For women co-ord sets or casual sets, detect Cotton, Rayon, or Cotton Blend accurately>",\n'
        '  "fabric_detail": "<Specific fabric description e.g. 100% Cotton, Pure Cotton, Cotton Blend, Soft Fleece, Poly Cotton, Rayon>",\n'
        '  "sleeve_length": "<Full Sleeve / Long Sleeve / Short Sleeve / 3/4 Sleeve / Half Sleeve / Sleeveless>",\n'
        '  "pattern": "<Solid / Printed / Washed / Striped / Checked / Floral / Self Design>"\n'
        "}\n"
        "Crucial Rules: Do NOT output Denim unless it is clearly denim material. If it is a 2-piece set or co-ord for women, gender is Women and category is co-ord set."
    )

# 1. Gemini Vision Call
def call_gemini_vision(image_bytes, img_name, api_key, model):
    if not api_key or not image_bytes:
        return None
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    payload = {
        "contents": [{"parts": [
            {"text": get_vision_prompt()},
            {"inline_data": {"mime_type": guess_mime(img_name), "data": b64}}
        ]}],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 1500
        },
    }
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        res = requests.post(url, json=payload, timeout=40)
        if res.status_code == 200:
            cands = res.json().get('candidates') or []
            if cands:
                parts = (cands[0].get('content') or {}).get('parts') or []
                text = ''.join(str(p.get('text', '')) for p in parts)
                return extract_json_from_text(text)
    except Exception:
        pass
    return None

# 2. NVIDIA Vision Call
def call_nvidia_vision(image_bytes, img_name, api_key, model):
    if not api_key or not image_bytes:
        return None
    mime = guess_mime(img_name)
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    data_url = f"data:{mime};base64,{b64}"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": get_vision_prompt()},
                    {"type": "image_url", "image_url": {"url": data_url}}
                ]
            }
        ],
        "temperature": 0.2,
        "max_tokens": 1500
    }
    try:
        invoke_url = "https://integrate.api.nvidia.com/v1/chat/completions"
        res = requests.post(invoke_url, headers=headers, json=payload, timeout=50)
        if res.status_code == 200:
            content = res.json()['choices'][0]['message']['content']
            return extract_json_from_text(content)
    except Exception:
        pass
    return None

# 3. Groq Vision Call
def call_groq_vision(image_bytes, img_name, api_key, model):
    if not api_key or not image_bytes:
        return None
    mime = guess_mime(img_name)
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    data_url = f"data:{mime};base64,{b64}"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": get_vision_prompt()},
                    {"type": "image_url", "image_url": {"url": data_url}}
                ]
            }
        ],
        "temperature": 0.1,
        "max_tokens": 1500
    }
    try:
        url = "https://api.groq.com/openai/v1/chat/completions"
        res = requests.post(url, headers=headers, json=payload, timeout=40)
        if res.status_code == 200:
            content = res.json()['choices'][0]['message']['content']
            return extract_json_from_text(content)
    except Exception:
        pass
    return None

# Unified Detection with Step-by-Step Fallback (Gemini ➜ NVIDIA ➜ Groq)
def multi_ai_vision_detect(image_bytes, img_name):
    res = call_gemini_vision(image_bytes, img_name, GEMINI_API_KEY.strip(), GEMINI_VISION_MODEL)
    if res and isinstance(res, dict) and res.get('category'):
        return res, "Gemini"

    res = call_nvidia_vision(image_bytes, img_name, NVIDIA_API_KEY.strip(), NVIDIA_MODEL)
    if res and isinstance(res, dict) and res.get('category'):
        return res, "NVIDIA"

    res = call_groq_vision(image_bytes, img_name, GROQ_API_KEY.strip(), GROQ_VISION_MODEL)
    if res and isinstance(res, dict) and res.get('category'):
        return res, "Groq"

    return None, "None"

def canonicalize_ai_category(cat_str):
    cat = str(cat_str or '').strip().lower().replace('-', ' ').strip()
    cat = AI_CAT_MAP.get(cat, cat)
    if cat in CATEGORY_OPTIONS:
        return cat
    close = difflib.get_close_matches(cat, CATEGORY_OPTIONS, n=1, cutoff=0.6)
    return close[0] if close else 'apparel'

def canonicalize_ai_color(col_str):
    col = str(col_str or '').strip()
    if not col:
        return None
    canon = str(_canonicalize_color(col))
    joined = canon.lower().replace(' ', '')
    for c in COMMON_COLORS:
        if c.lower().replace(' ', '') == joined:
            return c
    words = re.split(r'[\s\-]+', canon)
    return ' '.join(w.capitalize() for w in words if w)

def canonicalize_ai_pocket(p_str):
    p = str(p_str or '').strip()
    if not p:
        return None
    pl = p.lower()
    for k in KNOWN_POCKETS:
        if k.lower() == pl:
            return k
    close = difflib.get_close_matches(p, KNOWN_POCKETS, n=1, cutoff=0.5)
    return close[0] if close else p.title()

def run_vision_detection(zip_images_data):
    results, failed = {}, []
    folders = [f for f in zip_images_data.keys() if zip_images_data.get(f)]
    if not folders:
        return results, failed
    progress = st.progress(0.0)
    status = st.empty()
    for i, fld in enumerate(folders):
        cs = clean_sku_name(fld)
        img_name, img_bytes = zip_images_data[fld][0]
        status.info(f"🤖 AI analyzing garment image: **{cs}** ({i+1}/{len(folders)})...")
        
        data, provider = multi_ai_vision_detect(img_bytes, img_name)
        if data:
            cat = canonicalize_ai_category(data.get('category'))
            col = canonicalize_ai_color(data.get('color'))
            poc = canonicalize_ai_pocket(data.get('pocket'))
            gen = "Women" if str(data.get('gender')).lower().startswith('w') else "Men"
            
            mat = str(data.get('material') or '').strip().title()
            if not mat or (mat.lower() == 'denim' and cat != 'jeans'):
                mat = DEFAULT_CATEGORY_FABRICS.get(cat, 'Cotton Blend')
                
            fab_det = str(data.get('fabric_detail') or f"100% {mat}").title()
            sleeve = str(data.get('sleeve_length') or ('Long Sleeve' if cat in ['hoodie', 'sweatshirt'] else 'Short Sleeve')).title()
            pattern = str(data.get('pattern') or 'Solid').title()

            results[cs] = {
                'category': cat,
                'color': col,
                'pocket': poc,
                'gender': gen,
                'material': mat,
                'fabric_detail': fab_det,
                'sleeve_length': sleeve,
                'pattern': pattern,
                'ai_source': provider
            }
        else:
            failed.append(cs)
        time.sleep(0.3)
        progress.progress(min((i + 1) / len(folders), 1.0))
    progress.empty()
    status.empty()
    return results, failed

# ======================================================================
# CONTENT BUILDERS (DYNAMIC GENDER & PRODUCT TYPE)
# ======================================================================
def _user_kw_list(user_kw):
    if not user_kw:
        return []
    return [k.strip() for k in user_kw.replace(';', ',').split(',') if k.strip()]

def build_vendor_title(brand, category, full_sku, user_kw="", limit=75, gender="Women"):
    b = f"{brand} " if brand else ""
    cat_str = CATEGORY_LABELS.get(category, category.title())
    sku_part = f"({full_sku})"
    kws = _user_kw_list(user_kw)
    kw_candidate = kws[0] if kws else ""
    target_str = "for women" if gender == "Women" else "for men"
    if kw_candidate:
        opt1 = f"{b}{cat_str} {target_str} | {kw_candidate} {sku_part}".strip()
        if len(opt1) <= limit:
            title = opt1
        else:
            opt2 = f"{b}{kw_candidate} {sku_part}".strip()
            title = opt2 if len(opt2) <= limit else f"{b}{cat_str} {sku_part}".strip()
    else:
        title = f"{b}{cat_str} {target_str} | casual wear {sku_part}".strip()
    if len(title) > limit:
        title = title[:limit].rstrip()
    return title[:limit]

def build_sc_title(brand, cat, full_sku, user_kw="", limit=150, gender="Women"):
    b = f"{brand} " if brand else ""
    label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Casual Wear")
    kws = _user_kw_list(user_kw)
    g_title = "Women's" if gender == "Women" else "Men's"
    base = f"{b}{g_title} {label} for Casual Wear"
    for n in range(min(3, len(kws)), -1, -1):
        t = f"{base} | {', '.join(kws[:n])} ({full_sku})" if n else f"{base} ({full_sku})"
        if len(t) <= limit:
            return t
    return t[:limit]

def generate_generic_keywords_200(cat, det_col, user_kw, gender="Women"):
    if user_kw:
        kws, unique = _user_kw_list(user_kw), []
        for k in kws:
            if k.lower() not in [x.lower() for x in unique]:
                unique.append(k)
        return "; ".join(unique)[:200]
    g_str = "women" if gender == "Women" else "men"
    cat_label = CATEGORY_LABELS.get(cat, cat)
    kw_list = [f"{g_str} {cat_label}", f"{det_col} {cat_label}", f"casual wear for {g_str}", "comfortable fit", "daily wear", "stylish modern wear", "breathable fabric"]
    return "; ".join(kw_list)[:200]

def build_freehand_bullets_and_desc(cat, brand, gender="Women", det_mat="Cotton Blend"):
    b_prefix = f"{brand} " if brand else ""
    cat_label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Casual Wear")
    g_str = "Women's" if gender == "Women" else "Men's"

    if cat == 'co-ord set':
        bullets = [
            f"PREMIUM SOFT FABRIC: Crafted from high-grade, breathable {det_mat} ensuring all-day comfort and a soft skin feel.",
            "ELEGANT 2-PIECE CO-ORD SET: Features a coordinated matching silhouette designed for effortlessly chic casual styling.",
            "VERSATILE EVERYDAY WEAR: Perfect for outings, travel, lounge sessions, brunches, and relaxed casual gatherings.",
            "COMFORTABLE ERGONOMIC FIT: Tailored cut provides ease of movement and all-day flattering drape.",
            "WASH & CARE: Easy gentle machine wash cold with similar colors. Do not bleach. Warm iron if required."
        ]
        desc = (f"Upgrade your wardrobe with this stylish {b_prefix}{g_str} {cat_label}. "
                f"Crafted from premium quality {det_mat}, it delivers an ideal mix of contemporary style and snug comfort for daily wear.")
    elif cat in ['hoodie', 'sweatshirt']:
        bullets = [
            f"PREMIUM SOFT FLEECE FABRIC: Made from high-grade, breathable {det_mat} that keeps you warm, cozy, and comfortable throughout the day.",
            "URBAN STREETWEAR DESIGN: Features an on-trend silhouette with a functional hood and durable drawstrings.",
            "SPACIOUS KANGAROO POCKET: Front pocket provides convenient storage for daily essentials while keeping hands warm.",
            "VERSATILE EVERYDAY WEAR: Effortlessly pairs with jeans, cargo pants, joggers, and sneakers for casual outings and winter styling.",
            "WASH & CARE: Easy machine wash cold with similar colors inside out. Do not bleach. Warm iron if required."
        ]
        desc = (f"Upgrade your wardrobe with this stylish {b_prefix}{g_str} {cat_label}. "
                f"Crafted from premium quality {det_mat}, it delivers an ideal mix of modern style and snug comfort for daily wear.")
    elif cat == 'cargo':
        bullets = [
            f"DURABLE UTILITY FABRIC: Tailored from rugged, breathable {det_mat} designed for active mobility and long-lasting durability.",
            "MULTI-POCKET FUNCTIONALITY: Equipped with deep tactical flap cargo pockets and side pockets for maximum utility.",
            "COMFORTABLE RELAXED FIT: Features an ergonomic waistband and relaxed leg cut for easy everyday movement.",
            "IDEAL FOR CASUAL & OUTDOOR WEAR: Perfect for travel, hangouts, and pairing with oversized tees, hoodies, or boots.",
            "WASH & CARE: Machine wash cold inside out. Tumble dry low. Do not bleach. Warm iron if needed."
        ]
        desc = (f"Experience peak comfort and utility with the {b_prefix}{g_str} {cat_label}. "
                f"Engineered with reinforced stitching and high-strength fabric for an active, stylish lifestyle.")
    elif cat == 'jeans':
        bullets = [
            f"COMFORT STRETCH DENIM: Crafted from premium stretchable {det_mat} denim for flexible mobility and all-day ease.",
            "AUTHENTIC WASHED FINISH: Features a contemporary wash with subtle fading and durable contrast stitching.",
            "CLASSIC 5-POCKET DESIGN: Practical 5-pocket layout with sturdy rivets, heavy-duty zip fly, and button closure.",
            "TIMELESS RELAXED CUT: Sits comfortably at the waist with an authentic silhouette suitable for daily college and office casuals.",
            "WASH & CARE: Machine wash cold inside out with similar colors. Do not bleach. Warm iron if required."
        ]
        desc = (f"Step out in confidence with the {b_prefix}{g_str} {cat_label}. "
                f"Designed with durable stretch denim, it combines all-day flexibility with an elevated, versatile everyday look.")
    else:
        bullets = [
            f"HIGH-QUALITY COMFORT FABRIC: Crafted from premium {det_mat} engineered for soft skin feel and long-lasting daily wear.",
            "CONTEMPORARY STYLING: Features an on-trend silhouette with neat, durable seam construction and a comfortable fit.",
            "FUNCTIONAL & ERGONOMIC DESIGN: Tailored cut provides comfortable mobility and effortless wearability throughout the day.",
            "VERSATILE CASUAL PAIRING: Complements casual sneakers and daily wardrobe essentials for an easy, elevated style statement.",
            "WASH & CARE: Machine wash cold with like colors inside out. Do not bleach. Warm iron if required."
        ]
        desc = (f"Upgrade your daily rotation with the {b_prefix}{g_str} {cat_label}. "
                f"Engineered with premium fabric, it delivers the perfect balance of everyday durability and contemporary comfort.")
    return bullets, desc

def build_ajio_seo_content(cat, user_keywords="", gender="Women", mat_val="Cotton Blend"):
    kw_str = user_keywords if user_keywords else "casual streetwear, relaxed fit, durable quality, all-day comfort"
    cat_label = CATEGORY_LABELS.get(cat, 'Apparel')
    g_title = "Women's" if gender == "Women" else "Men's"
    
    if cat == 'co-ord set':
        desc = (f"{g_title} Co-ord Set - Designed for effortless chic look and supreme comfort, this {g_title} Co-ord Set "
                f"is crafted from premium breathable {mat_val} fabric. Engineered with {kw_str}. "
                "Features matching top and bottom silhouette with comfortable flexi waist and neat seam finish. "
                "Perfect for daily casuals, travel, vacations, and relaxed outings.")
        highlight = f"Premium {mat_val} fabric | Comfortable {g_title.lower()} fit | Elegant matching 2-piece set | Relaxed styling | Easy machine wash"
        add_info = "Styling Tip: Pair with minimal jewelry and slip-ons. Wash Care: Machine wash cold with similar colors."
    elif cat == 'cargo':
        desc = (f"CARGO FOR {gender.upper()} - Designed for unmatched utility and streetwear style, this {g_title} Cargo Pants "
                f"is crafted from premium breathable {mat_val} fabric. Engineered with {kw_str}. "
                "Features multiple spacious flap cargo pockets, a comfortable flexi waist with sturdy button closure, "
                "and durable reinforced stitching throughout. Perfect for everyday wear, outdoor adventures, and travel.")
        highlight = f"Premium breathable {mat_val} fabric | Comfortable flexi waist | Multi-pocket functional cargo utility styling | Relaxed fit | Easy machine wash"
        add_info = "Styling Tip: Pair with oversized t-shirts, hoodies, and chunky sneakers. Wash Care: Machine wash cold with similar colors."
    elif cat == 'jeans':
        desc = (f"Upgrade your everyday denim rotation with this trendy {g_title} Jeans. "
                f"Crafted from premium stretchable denim engineered with {kw_str}. "
                "Featuring a contemporary mid-rise relaxed fit, functional 5-pocket styling, sturdy zip fly, and authentic washed finish. "
                "Ideal for casual outings, daily wear, and weekend hangouts.")
        highlight = "Premium stretch denim | Relaxed mid-rise silhouette | Authentic clean washed finish | Multi-pocket utility styling | All-season wear"
        add_info = "Styling Tip: Pair with casual shirts or boxy tees. Wash Care: Machine wash cold inside out."
    elif cat in ['hoodie', 'sweatshirt']:
        desc = (f"Stay warm and stylish with this {g_title} Hoodie/Sweatshirt crafted from premium soft {mat_val} fabric. "
                f"Engineered with {kw_str}. Features a functional hood with durable drawstrings, spacious kangaroo pocket, "
                "and ribbed cuffs for a snug fit. Perfect for winter layering, travel, and everyday streetwear.")
        highlight = f"Premium soft {mat_val} fabric | Functional hood with drawstrings | Spacious kangaroo pocket | Ribbed cuffs | Machine washable"
        add_info = "Styling Tip: Layer over tees with jeans or joggers. Wash Care: Machine wash cold inside out."
    else:
        desc = (f"{g_title} {cat_label} crafted from premium quality {mat_val} fabric featuring {kw_str}. "
                "Designed for maximum comfort, durability, and everyday style. Easy to pair and easy to care.")
        highlight = f"Premium {mat_val} fabric | Comfortable {g_title.lower()} fit | Durable stitching | Easy machine wash"
        add_info = "Wash Care: Machine wash cold with like colors. Do not bleach."
    return desc[:950], highlight[:500], add_info[:500]

# ======================================================================
# WORKBOOK HELPERS
# ======================================================================
def ungroup_worksheet(ws):
    for col_dim in ws.column_dimensions.values():
        col_dim.outline_level = 0
        col_dim.collapsed = False
        col_dim.hidden = False
    for r_dim in ws.row_dimensions.values():
        r_dim.outline_level = 0
        r_dim.collapsed = False
        r_dim.hidden = False
    if hasattr(ws, 'sheet_properties') and hasattr(ws.sheet_properties, 'outlinePr'):
        ws.sheet_properties.outlinePr.summaryBelow = True
        ws.sheet_properties.outlinePr.summaryRight = True

def load_workbook_clean(uploaded_file_or_bytes, name=""):
    is_xlsm = str(name).lower().endswith('.xlsm')
    wb = openpyxl.load_workbook(uploaded_file_or_bytes, keep_vba=is_xlsm)
    for s in wb.sheetnames:
        wb[s].protection.sheet = False
        try:
            wb[s].protection.disable()
        except Exception:
            pass
    return wb

def pick_template_sheet(wb):
    for sname in wb.sheetnames:
        if sname.strip().lower() == 'template':
            return wb[sname]
    for sname in wb.sheetnames:
        if sname.strip().lower().startswith('template'):
            return wb[sname]
    for sname in wb.sheetnames:
        if 'template' in sname.lower() and 'change' not in sname.lower():
            return wb[sname]
    return wb.worksheets[0]

def get_sku_img_urls(raw_fld, ctx, clean_style):
    if raw_fld in ctx['live_links_cache'] and ctx['live_links_cache'][raw_fld]:
        urls = list(ctx['live_links_cache'][raw_fld])
    else:
        cnt = ctx['folder_images_count'].get(raw_fld, 5)
        urls = [make_workable_dropbox_url(ctx['dropbox'], clean_style, i) for i in range(1, cnt + 1)]
    if ctx.get('size_chart_cdn_url') and ctx['size_chart_cdn_url'] not in urls:
        urls.append(ctx['size_chart_cdn_url'])
    return urls

# ======================================================================
#                                AJIO ENGINE
# ======================================================================
def find_sheet_for_category(wb, category):
    ignore = ['instruction', 'guideline', 'extra', 'rejection', 'seller info']
    valid_sheets = [s for s in wb.sheetnames if not any(ign in s.lower() for ign in ignore)]

    def has(wl, sheet): return any(w in sheet.lower() for w in wl)

    if category in ["co-ord set", "dress", "top", "kurti"]:
        for s in valid_sheets:
            if has(['coord', 'co-ord', 'suit', 'dress', 'top', 'kurti'], s): return s
    elif category in ["cargo", "trousers"]:
        for s in valid_sheets:
            if has(['trouser', 'pant', 'cargo'], s): return s
    elif category == "jeans":
        for s in valid_sheets:
            if 'jean' in s.lower(): return s
    elif category == "tshirt":
        for s in valid_sheets:
            if has(['t-shirt', 'tshirt', 'polo'], s): return s
    elif category == "shirt":
        for s in valid_sheets:
            if 'shirt' in s.lower() and not has(['t-shirt', 'tshirt'], s): return s
    elif category in ["hoodie", "sweatshirt"]:
        for s in valid_sheets:
            if has(['sweat', 'hoodie', 'fleece'], s): return s
    elif category in ["joggers"]:
        for s in valid_sheets:
            if has(['jogger', 'track'], s): return s
    elif category == "shorts":
        for s in valid_sheets:
            if 'short' in s.lower(): return s
    for s in valid_sheets:
        if 'styles' in s.lower(): return s
    return valid_sheets[0] if valid_sheets else wb.sheetnames[0]

def run_ajio(uploaded_file, ctx):
    wb = load_workbook_clean(uploaded_file, uploaded_file.name)
    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)
    season_year = str(datetime.now().year + 1)

    category_folders = {}
    for raw_folder in ctx['folders']:
        cs = clean_sku_name(raw_folder)
        cat = resolve_category(cs, ctx)
        category_folders.setdefault(cat, []).append(raw_folder)

    size_list = ctx['sizes']
    total_variants_added = 0
    sheets_modified = set()
    preview = []

    def aset(ws, row, col, value, dropdown=False, fill=None, font=None):
        if not col or value in (None, ''):
            return
        cell = ws.cell(row, col)
        cell.value = pick_dropdown_value(ws, col, value, value) if dropdown else value
        if fill: cell.fill = fill
        if font: cell.font = font

    for cat, folder_list in category_folders.items():
        target_sheet_name = find_sheet_for_category(wb, cat)
        ws = wb[target_sheet_name]
        sheets_modified.add(target_sheet_name)

        header_row, best_matches = 1, 0
        keywords = ['style', 'sku', 'mrp', 'price', 'size', 'image', 'model', 'brand', 'description', 'weight']
        for r in range(1, min(7, ws.max_row + 1)):
            vals = [str(ws.cell(r, c).value or '').lower() for c in range(1, ws.max_column + 1)]
            m = sum(1 for v in vals if any(k in v for k in keywords))
            if m > best_matches:
                best_matches, header_row = m, r

        headers = {c: str(ws.cell(header_row, c).value or '').strip() for c in range(1, ws.max_column + 1)}

        def get_exact_col(exact_name):
            ct = exact_name.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
            for c, h in headers.items():
                ch = h.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
                if ch == ct:
                    return c
            return None

        def get_fallback_col(candidates):
            for cand in candidates:
                c = get_exact_col(cand)
                if c:
                    return c
            return None

        col_style = get_fallback_col(['Style Code', 'Parent SKU'])
        col_desc = get_fallback_col(['Style Description', 'Product Description', 'Description'])
        col_sku = get_fallback_col(['Item SKU', 'SKU ID', 'Seller SKU', 'SKU'])
        col_brand = get_fallback_col(['Brand', 'Brand Name'])
        col_ean = get_fallback_col(['EAN', 'Barcode', 'UPC'])
        col_td = get_fallback_col(['TD'])
        col_mrp = get_fallback_col(['MRP', 'Maximum Retail Price', 'Standard Price'])
        col_hsn = get_fallback_col(['HSN', 'HSN Code'])
        col_prod_grp = get_fallback_col(['Product Groups'])
        col_fash_grp = get_fallback_col(['Fashion Groups'])
        col_season_cd = get_fallback_col(['Season Code'])
        col_season_yr = get_fallback_col(['Season Year'])
        col_size = get_fallback_col(['Size'])
        col_std_size = get_fallback_col(['StandardSize', 'Standard Size'])

        col_art_h = get_fallback_col(['articleDimensionsUnitHeight'])
        col_art_l = get_fallback_col(['articleDimensionsUnitLength'])
        col_art_w = get_fallback_col(['articleDimensionsUnitWidth'])
        col_art_luom = get_fallback_col(['articleDimensionsUnitLengthUOM'])
        col_art_wt = get_fallback_col(['articleDimensionsUnitWeight'])
        col_art_wtuom = get_fallback_col(['articleDimensionsUnitWeightUOM'])
        col_pkg_h = get_fallback_col(['packageDimensionsHeight'])
        col_pkg_l = get_fallback_col(['packageDimensionsLength'])
        col_pkg_w = get_fallback_col(['packageDimensionsWidth'])
        col_pkg_luom = get_fallback_col(['packageDimensionsLengthUOM'])
        col_pkg_wt = get_fallback_col(['packageDimensionsWeight'])
        col_pkg_wtuom = get_fallback_col(['packageDimensionsWeightUOM'])

        col_comp_count = get_fallback_col(['Component Count'])
        col_country = get_fallback_col(['Country of Origin'])
        col_mfg_by = get_fallback_col(['Manufactured By'])
        col_mkt_by = get_fallback_col(['Marketed By'])
        col_net_qty = get_fallback_col(['Net Quantity'])
        col_pkg_contains = get_fallback_col(['Package Contains'])

        col_color = get_fallback_col(['Color Family', 'Colour Family'])
        col_pri_color = get_fallback_col(['Primary color', 'Primary Color', 'Color', 'Colour'])
        col_fab_detail = get_fallback_col(['Fabric Detail'])
        col_fab_type = get_fallback_col(['Fabric Type'])
        col_pattern = get_fallback_col(['Pattern'])
        col_size_format = get_fallback_col(['Size Format'])
        col_size_group = get_fallback_col(['Size Group'])
        col_wash_care = get_fallback_col(['Wash Care'])

        col_length = get_fallback_col(['Length'])
        col_stretch = get_fallback_col(['Stretch'])
        col_style_type = get_fallback_col(['Style Type'])
        col_fit_type = get_fallback_col(['Fit Type'])
        col_front_style = get_fallback_col(['Front Style'])
        col_waist_rise = get_fallback_col(['Waist Rise'])
        col_waist_type = get_fallback_col(['Waist Type'])
        col_meas_type = get_fallback_col(['Measurement Type'])
        col_to_fit = get_fallback_col(['To Fit Waist'])
        col_waist = get_fallback_col(['Waist'])
        col_inseam = get_fallback_col(['Inseam'])
        col_sleeve_len = get_fallback_col(['Sleeve Length', 'Sleeve'])

        col_highlight = get_fallback_col(['Highlight'])
        col_add_info = get_fallback_col(['Additional Information 1'])

        img_cols = []
        for i in range(1, 15):
            target_name = '*MODEL' if i == 1 else f'MODEL{i}'
            c_i = get_fallback_col([target_name, f'Image {i}', f'Image_{i}', f'Image{i}'])
            if c_i and c_i not in img_cols:
                img_cols.append(c_i)

        curr_r = header_row + 1
        while ws.cell(curr_r, col_style if col_style else 1).value is not None:
            curr_r += 1

        for raw_fld in folder_list:
            clean_style = clean_sku_name(raw_fld)
            det_col = resolve_color(raw_fld, clean_style, ctx)
            det_gen = resolve_gender(clean_style, ctx)
            color_family_val = map_to_standard_color_map(det_col)
            
            vr_info = (ctx.get('vision_results') or {}).get(clean_style, {})
            fab_type_val = resolve_fabric(clean_style, cat, ctx)
            fab_detail_val = vr_info.get('fabric_detail') or f"100% {fab_type_val}"
            pattern_val = vr_info.get('pattern') or ("Solid" if cat not in ['jeans'] else "Washed")
            sleeve_val = resolve_sleeve_length(clean_style, cat, ctx)

            desc_text, high_text, add_text = build_ajio_seo_content(cat, ctx['keywords'], det_gen, fab_type_val)
            sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)

            if cat == 'co-ord set':
                style_type_val = "2-Piece Sets"
            elif cat in ['cargo', 'trousers']:
                style_type_val = "Cargo"
            else:
                style_type_val = "Relaxed"

            preview.append({
                'Sheet': target_sheet_name, 'SKU': clean_style, 'Category': cat, 'Gender': det_gen,
                'Color': det_col, 'Color Family (Map)': color_family_val, 'Fabric Type': fab_type_val,
                'Sleeve': sleeve_val
            })

            for sz in size_list:
                clean_sz = clean_sku_name(sz)
                full_sku = f"{clean_style}-{clean_sz}"
                ean_val = clean_sku_name(full_sku)[:18]

                try:
                    num_sz = int(re.findall(r'\d+', clean_sz)[0])
                    w_val, tf_val = str(num_sz + 1), str(num_sz - 1)
                except Exception:
                    w_val, tf_val = "34", "32"

                aset(ws, curr_r, col_style, clean_style)
                aset(ws, curr_r, col_desc, desc_text)
                aset(ws, curr_r, col_sku, full_sku)
                aset(ws, curr_r, col_brand, ctx['brand'])
                aset(ws, curr_r, col_ean, ean_val)
                aset(ws, curr_r, col_td, ctx['td'])
                aset(ws, curr_r, col_hsn, get_hsn_code(cat))
                aset(ws, curr_r, col_prod_grp, "Casual", dropdown=True)
                aset(ws, curr_r, col_fash_grp, "Fashion", dropdown=True)
                aset(ws, curr_r, col_season_cd, "Core", dropdown=True)
                aset(ws, curr_r, col_season_yr, season_year)
                aset(ws, curr_r, col_size, clean_sz)
                aset(ws, curr_r, col_std_size, clean_sz)

                dim_h, dim_l, dim_w, dim_uom = "2", "25", "24", "CM"
                wt_val, wt_uom = "350", "GRAM"
                aset(ws, curr_r, col_art_h, dim_h); aset(ws, curr_r, col_art_l, dim_l)
                aset(ws, curr_r, col_art_w, dim_w); aset(ws, curr_r, col_art_luom, dim_uom)
                aset(ws, curr_r, col_art_wt, wt_val); aset(ws, curr_r, col_art_wtuom, wt_uom)
                aset(ws, curr_r, col_pkg_h, dim_h); aset(ws, curr_r, col_pkg_l, dim_l)
                aset(ws, curr_r, col_pkg_w, dim_w); aset(ws, curr_r, col_pkg_luom, dim_uom)
                aset(ws, curr_r, col_pkg_wt, wt_val); aset(ws, curr_r, col_pkg_wtuom, wt_uom)

                aset(ws, curr_r, col_comp_count, 1)
                aset(ws, curr_r, col_country, "India", dropdown=True)
                aset(ws, curr_r, col_mfg_by, ctx['mfg'])
                aset(ws, curr_r, col_mkt_by, ctx['mfg'])
                aset(ws, curr_r, col_net_qty, 1)
                aset(ws, curr_r, col_pkg_contains, "1")

                aset(ws, curr_r, col_color, color_family_val, dropdown=True)
                aset(ws, curr_r, col_pri_color, det_col, dropdown=True)
                aset(ws, curr_r, col_fab_detail, fab_detail_val)
                aset(ws, curr_r, col_fab_type, fab_type_val, dropdown=True)
                aset(ws, curr_r, col_pattern, pattern_val, dropdown=True)
                aset(ws, curr_r, col_size_format, "UNI", dropdown=True)
                aset(ws, curr_r, col_size_group, "Regular", dropdown=True)
                aset(ws, curr_r, col_wash_care, "Machine wash", dropdown=True)

                aset(ws, curr_r, col_length, "Full length", dropdown=True)
                aset(ws, curr_r, col_stretch, "Stretch", dropdown=True)
                aset(ws, curr_r, col_style_type, style_type_val, dropdown=True)
                aset(ws, curr_r, col_fit_type, "Relaxed Fit", dropdown=True)
                aset(ws, curr_r, col_front_style, "Flat Front", dropdown=True)
                aset(ws, curr_r, col_waist_rise, "Mid Rise", dropdown=True)
                aset(ws, curr_r, col_waist_type, "Flexi waist" if cat in ['cargo', 'trousers', 'co-ord set'] else "Fixed waist", dropdown=True)
                aset(ws, curr_r, col_meas_type, "Garment Measurement", dropdown=True)
                aset(ws, curr_r, col_to_fit, tf_val)
                aset(ws, curr_r, col_waist, w_val)
                aset(ws, curr_r, col_inseam, "32")
                aset(ws, curr_r, col_sleeve_len, sleeve_val, dropdown=True)

                aset(ws, curr_r, col_highlight, high_text)
                aset(ws, curr_r, col_add_info, add_text)

                for img_idx, img_link in enumerate(sku_img_urls):
                    if img_idx < len(img_cols):
                        ws.cell(curr_r, img_cols[img_idx]).value = img_link

                if col_mrp:
                    mrp_cell = ws.cell(curr_r, col_mrp)
                    if ctx['mrp']:
                        mrp_cell.value = ctx['mrp']
                    else:
                        mrp_cell.value = ""
                        mrp_cell.fill = mrp_fill
                        mrp_cell.font = mrp_font

                curr_r += 1
                total_variants_added += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), preview, total_variants_added, sheets_modified

# ======================================================================
#                AMAZON VENDOR CENTRAL ENGINE
# ======================================================================
def run_vendor_central(uploaded_file, ctx):
    wb = load_workbook_clean(uploaded_file, uploaded_file.name)
    ws = pick_template_sheet(wb)
    ungroup_worksheet(ws)

    col_by_label, col_by_key = {}, {}
    for c in range(1, ws.max_column + 1):
        lbl = str(ws.cell(3, c).value or '').strip().lower()
        key = str(ws.cell(4, c).value or '').strip().lower()
        if lbl and lbl not in col_by_label:
            col_by_label[lbl] = c
        if key and key not in col_by_key:
            col_by_key[key] = c

    def get_vc_col(labels, keys=()):
        for k in keys:
            if k.lower() in col_by_key:
                return col_by_key[k.lower()]
        for l in labels:
            if l.lower() in col_by_label:
                return col_by_label[l.lower()]
        return None

    curr_r = 7
    sku_col_idx = get_vc_col(['vendor sku'], ['vendor_sku#1.value']) or 2
    while ws.cell(curr_r, sku_col_idx).value is not None:
        curr_r += 1

    today_str = datetime.now().strftime("%Y-%m-%d")
    today_shipping = datetime.now().strftime("%Y%m%d")
    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)
    material_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    added, preview = 0, []

    def vset(col, value, dropdown=False, highlight=None, bold=False):
        if not col:
            return
        cell = ws.cell(curr_r, col)
        if value not in (None, ''):
            cell.value = pick_dropdown_value(ws, col, value, value) if dropdown else value
        if highlight:
            cell.fill = highlight
        if bold:
            cell.font = mrp_font

    for raw_fld in ctx['folders']:
        clean_style = clean_sku_name(raw_fld)
        det_col = resolve_color(raw_fld, clean_style, ctx)
        cat = resolve_category(clean_style, ctx, ws.title)
        det_gen = resolve_gender(clean_style, ctx)
        cat_label = CATEGORY_LABELS.get(cat, cat.title() if cat else "Casual Wear")

        sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)
        mat_val = resolve_fabric(clean_style, cat, ctx)
        sleeve_val = resolve_sleeve_length(clean_style, cat, ctx)

        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
        gen_keywords_200 = generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)
        style_val = get_style_from_dropdown(cat)
        item_type_val = get_item_type_name_from_dropdown(cat)
        color_map_val = map_to_standard_color_map(det_col)
        pocket_val = resolve_pocket(clean_style, cat, ctx)

        dept_val = "Womens" if det_gen == "Women" else "Mens"
        gender_val = "Female" if det_gen == "Women" else "Male"

        preview.append({'SKU': clean_style, 'Category': cat, 'Gender': det_gen, 'Color': det_col,
                        'Color Map': color_map_val, 'Item Type': item_type_val,
                        'Style': style_val, 'Pocket': pocket_val, 'Material': mat_val})

        for sz in ctx['sizes']:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            is_num_sz = clean_sz.isdigit()
            vendor_title = build_vendor_title(ctx['brand'], cat, full_sku, ctx['keywords'], 75, det_gen)

            c_vc = get_vc_col(['vendor code'], ['rtip_vendor_code#1.value'])
            if c_vc:
                ws.cell(curr_r, c_vc).value = "J B fashion -- Surat -- Dropship, in_pd-ravi_apparel, JBZ54"

            c_sku = get_vc_col(['vendor sku'], ['vendor_sku#1.value'])
            if c_sku:
                ws.cell(curr_r, c_sku).value = full_sku

            for p_lbl, p_key in [
                (['parentage level'], ['parentage_level#1.value']),
                (['child relationship type'], ['child_parent_sku_relationship#1.child_relationship_type']),
                (['parent sku'], ['child_parent_sku_relationship#1.parent_sku']),
                (['variation theme name'], ['variation_theme#1.name']),
            ]:
                c_p = get_vc_col(p_lbl, p_key)
                if c_p:
                    ws.cell(curr_r, c_p).value = None

            c_title = get_vc_col(['item name'], ['item_name#1.value'])
            if c_title:
                ws.cell(curr_r, c_title).value = vendor_title

            c_hl = get_vc_col(['item highlight'], ['title_differentiation#1.value'])
            if c_hl:
                ws.cell(curr_r, c_hl).value = f"Premium quality {cat_label} with comfortable durable fit"

            c_br = get_vc_col(['brand name'], ['brand#1.value'])
            if c_br:
                ws.cell(curr_r, c_br).value = ctx['brand'] or None

            c_mod = get_vc_col(['model number'], ['model_number#1.value'])
            if c_mod:
                ws.cell(curr_r, c_mod).value = full_sku
            c_part = get_vc_col(['part number'], ['part_number#1.value'])
            if c_part:
                ws.cell(curr_r, c_part).value = full_sku
            c_mod_name = get_vc_col(['model name'], ['model_name#1.value'])
            if c_mod_name:
                ws.cell(curr_r, c_mod_name).value = f"{cat.capitalize()} {clean_style}"

            c_mfg = get_vc_col(['manufacturer'], ['manufacturer#1.value'])
            if c_mfg:
                ws.cell(curr_r, c_mfg).value = ctx['mfg'] or None
            c_mfg_cnt = get_vc_col(['manufacturer contact information'], ['rtip_manufacturer_contact_information#1.value'])
            if c_mfg_cnt:
                ws.cell(curr_r, c_mfg_cnt).value = ctx['mfg'] or None

            c_main_img = get_vc_col(['main image url'], ['main_product_image_locator#1.media_location'])
            if c_main_img and sku_img_urls:
                ws.cell(curr_r, c_main_img).value = sku_img_urls[0]
            for img_num in range(1, 9):
                c_oth = get_vc_col(['other image url'], [f'other_product_image_locator_{img_num}#1.media_location'])
                if c_oth and img_num < len(sku_img_urls):
                    ws.cell(curr_r, c_oth).value = sku_img_urls[img_num]

            for b_i, b_val in enumerate(bullets_list[:5]):
                c_bp = get_vc_col(['bullet point'], [f'bullet_point#{b_i+1}.value'])
                if c_bp:
                    ws.cell(curr_r, c_bp).value = b_val

            c_gkw = get_vc_col(['generic keywords'], ['generic_keyword#1.value'])
            if c_gkw and gen_keywords_200:
                ws.cell(curr_r, c_gkw).value = gen_keywords_200

            c_style = get_vc_col(['style'], ['style#1.value'])
            if c_style:
                vset(c_style, style_val, dropdown=True)

            lifestyles = ["Casual", "Work Utility", "Comfort", "Athletic", "Formal"]
            for ls_idx in range(1, 6):
                c_ls = get_vc_col(['lifestyle'], [f'lifestyle#{ls_idx}.value'])
                if c_ls:
                    vset(c_ls, lifestyles[ls_idx - 1], dropdown=True)

            c_dept = get_vc_col(['department name'], ['department#1.value'])
            if c_dept:
                vset(c_dept, dept_val, dropdown=True)
            c_gen = get_vc_col(['target gender'], ['target_gender#1.value'])
            if c_gen:
                vset(c_gen, gender_val, dropdown=True)
            c_age = get_vc_col(['age range description'], ['age_range_description#1.value'])
            if c_age:
                vset(c_age, "Adult", dropdown=True)

            c_ssys = get_vc_col(['apparel size system', 'bottoms size system', 'size system'],
                                ['bottoms_size#1.size_system', 'apparel_size#1.size_system'])
            if c_ssys:
                vset(c_ssys, "IN", dropdown=True)

            size_class_val = "Numeric" if is_num_sz else "Alpha"
            c_scls = get_vc_col(['apparel size class', 'bottoms size class', 'size class'],
                                ['bottoms_size#1.size_class', 'apparel_size#1.size_class'])
            if c_scls:
                vset(c_scls, size_class_val, dropdown=True)

            c_sval = get_vc_col(['apparel size value', 'bottoms size value', 'size value'],
                                ['bottoms_size#1.size', 'apparel_size#1.size'])
            if c_sval:
                ws.cell(curr_r, c_sval).value = clean_sz

            c_mat = get_vc_col(['material'], ['material#1.value'])
            if c_mat:
                vset(c_mat, mat_val, dropdown=True, highlight=material_fill)

            c_fab = get_vc_col(['fabric type'], ['fabric_type#1.value'])
            if c_fab:
                ws.cell(curr_r, c_fab).value = f"100% {mat_val}"

            c_noi = get_vc_col(['number of items'], ['number_of_items#1.value'])
            if c_noi:
                ws.cell(curr_r, c_noi).value = 1

            c_itn = get_vc_col(['item type name'], ['item_type_name#1.value'])
            if c_itn:
                vset(c_itn, item_type_val, dropdown=True)

            c_spec_sz = get_vc_col(['special size'], ['special_size_type#1.value'])
            if c_spec_sz:
                vset(c_spec_sz, "Standard", dropdown=True)

            c_desc = get_vc_col(['product description'], ['rtip_product_description#1.value'])
            if c_desc:
                ws.cell(curr_r, c_desc).value = desc_text

            c_cmap = get_vc_col(['color map'], ['color#1.standardized_values#1'])
            if c_cmap:
                vset(c_cmap, color_map_val, dropdown=True)
            c_color = get_vc_col(['color'], ['color#1.value'])
            if c_color:
                ws.cell(curr_r, c_color).value = det_col

            c_thm = get_vc_col(['theme'], ['theme#1.value'])
            if c_thm:
                ws.cell(curr_r, c_thm).value = None

            c_fit = get_vc_col(['fit type'], ['fit_type#1.value'])
            if c_fit:
                vset(c_fit, "Relaxed", dropdown=True)

            c_poc = get_vc_col(['pocket description'], ['pocket_description#1.value'])
            if c_poc:
                vset(c_poc, pocket_val, dropdown=True)

            c_care = get_vc_col(['care instructions'], ['care_instructions#1.value'])
            if c_care:
                vset(c_care, "Machine Wash", dropdown=True)

            c_pat = get_vc_col(['pattern'], ['pattern#1.value'])
            if c_pat:
                vset(c_pat, "Solid", dropdown=True)

            c_uc = get_vc_col(['unit count'], ['unit_count#1.value'])
            if c_uc:
                ws.cell(curr_r, c_uc).value = 1
            c_uct = get_vc_col(['unit count type'], ['unit_count#1.type.value'])
            if c_uct:
                vset(c_uct, "count", dropdown=True)

            c_pld = get_vc_col(['product site launch date'], ['product_site_launch_date#1.value'])
            if c_pld:
                ws.cell(curr_r, c_pld).value = today_str

            c_hsn_ent = get_vc_col(['external product information entity'], ['external_product_information#1.entity'])
            if c_hsn_ent:
                vset(c_hsn_ent, "HSN Code", dropdown=True)
            c_hsn_val = get_vc_col(['external product information'], ['external_product_information#1.value'])
            if c_hsn_val:
                ws.cell(curr_r, c_hsn_val).value = get_hsn_code(cat)

            if is_num_sz:
                c_wst = get_vc_col(['waist size'], ['waist#1.size#1.value'])
                if c_wst:
                    ws.cell(curr_r, c_wst).value = clean_sz
                c_wstu = get_vc_col(['waist size unit'], ['waist#1.size#1.unit'])
                if c_wstu:
                    vset(c_wstu, "inches", dropdown=True)

            c_lst = get_vc_col(['product lifecycle supply type'], ['lifecycle_supply_type#1.value'])
            if c_lst:
                vset(c_lst, "Year Round Replenishable", dropdown=True)
            c_sea = get_vc_col(['seasons'], ['seasons#1.value'])
            if c_sea:
                vset(c_sea, "All", dropdown=True)

            c_sld = get_vc_col(['sleeve length description', 'sleeve length'])
            if c_sld:
                vset(c_sld, sleeve_val, dropdown=True)
            c_st = get_vc_col(['sleeve type'])
            if c_st:
                vset(c_st, "Regular Sleeve", dropdown=True)

            c_pkr = get_vc_col(['packer contact information'], ['packer_contact_information#1.value'])
            if c_pkr:
                ws.cell(curr_r, c_pkr).value = ctx['mfg'] or None

            c_iwt = get_vc_col(['item weight'], ['item_weight#1.value'])
            if c_iwt:
                ws.cell(curr_r, c_iwt).value = 350
            c_iwtu = get_vc_col(['item weight unit'], ['item_weight#1.unit'])
            if c_iwtu:
                vset(c_iwtu, "Grams", dropdown=True)

            c_mrp = get_vc_col(['maximum retail price (mrp)'], ['list_price#1.value_with_tax'])
            if c_mrp:
                if ctx['mrp']:
                    mv = ctx['mrp']
                    ws.cell(curr_r, c_mrp).value = int(mv) if str(mv).isdigit() else mv
                else:
                    ws.cell(curr_r, c_mrp).fill = mrp_fill
                    ws.cell(curr_r, c_mrp).font = mrp_font

            c_esd = get_vc_col(['earliest shipping date'], ['rtip_earliest_shipping_date#1.value'])
            if c_esd:
                ws.cell(curr_r, c_esd).value = today_shipping

            c_il = get_vc_col(['item length'], ['item_dimensions#1.length.value'])
            if c_il:
                ws.cell(curr_r, c_il).value = 10
            c_ilu = get_vc_col(['item length unit'], ['item_dimensions#1.length.unit'])
            if c_ilu:
                vset(c_ilu, "inches", dropdown=True)
            c_iw = get_vc_col(['item width'], ['item_dimensions#1.width.value'])
            if c_iw:
                ws.cell(curr_r, c_iw).value = 2
            c_iwu = get_vc_col(['item width unit'], ['item_dimensions#1.width.unit'])
            if c_iwu:
                vset(c_iwu, "inches", dropdown=True)
            c_ih = get_vc_col(['item height'], ['item_dimensions#1.height.value'])
            if c_ih:
                ws.cell(curr_r, c_ih).value = 1.3
            c_ihu = get_vc_col(['item height unit'], ['item_dimensions#1.height.unit'])
            if c_ihu:
                vset(c_ihu, "inches", dropdown=True)

            c_pl = get_vc_col(['item package length'], ['item_package_dimensions#1.length.value'])
            if c_pl:
                ws.cell(curr_r, c_pl).value = 10
            c_plu = get_vc_col(['package length unit'], ['item_package_dimensions#1.length.unit'])
            if c_plu:
                vset(c_plu, "inches", dropdown=True)
            c_pw = get_vc_col(['item package width'], ['item_package_dimensions#1.width.value'])
            if c_pw:
                ws.cell(curr_r, c_pw).value = 2
            c_pwu = get_vc_col(['package width unit'], ['item_package_dimensions#1.width.unit'])
            if c_pwu:
                vset(c_pwu, "inches", dropdown=True)
            c_ph = get_vc_col(['item package height'], ['item_package_dimensions#1.height.value'])
            if c_ph:
                ws.cell(curr_r, c_ph).value = 1.3
            c_phu = get_vc_col(['package height unit'], ['item_package_dimensions#1.height.unit'])
            if c_phu:
                vset(c_phu, "inches", dropdown=True)

            c_pwt = get_vc_col(['package weight'], ['item_package_weight#1.value'])
            if c_pwt:
                ws.cell(curr_r, c_pwt).value = 380
            c_pwtu = get_vc_col(['package weight unit'], ['item_package_weight#1.unit'])
            if c_pwtu:
                vset(c_pwtu, "Grams", dropdown=True)

            c_oat = get_vc_col(['order aggregate type'], ['rtip_order_aggregate_type#1.value'])
            if c_oat:
                vset(c_oat, "Each", dropdown=True)
            c_ipi = get_vc_col(['items per inner pack'], ['rtip_items_per_inner_pack#1.value'])
            if c_ipi:
                ws.cell(curr_r, c_ipi).value = 1
            c_coo = get_vc_col(['country of origin'], ['country_of_origin#1.value'])
            if c_coo:
                vset(c_coo, "India", dropdown=True)
            c_dg = get_vc_col(['dangerous goods regulations', 'dangerous goods', 'hazmat'])
            if c_dg:
                vset(c_dg, "Not Applicable", dropdown=True)

            curr_r += 1
            added += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), preview, added

# ======================================================================
#              AMAZON SELLER CENTRAL ENGINE
# ======================================================================
SC_FIELDS = {
    'sku': ['sku', 'seller sku', 'sellersku', 'sku id'],
    'product_type': ['product type', 'feed product type', 'product_type'],
    'item_name': ['item name', 'item_name', 'product name', 'title'],
    'brand_name': ['brand name', 'brand_name', 'brand'],
    'manufacturer': ['manufacturer', 'manufactured by'],
    'model_number': ['model number', 'model_number', 'model'],
    'part_number': ['part number', 'part_number', 'manufacturer part number'],
    'model_name': ['model name', 'model_name'],
    'item_type_name': ['item type name', 'item_type_name', 'item type keyword'],
    'color_name': ['color name', 'color_name', 'color', 'colour name', 'colour'],
    'color_map': ['color map', 'color_map', 'colour map'],
    'size_name': ['size name', 'size_name', 'size', 'apparel size value', 'bottoms size value', 'size value'],
    'size_system': ['size system', 'size_system', 'apparel size system', 'bottoms size system'],
    'quantity': ['quantity', 'qty', 'quantity (in)'],
    'standard_price': ['standard price', 'standard_price', 'price', 'selling price', 'your price inr (sell on amazon, in)'],
    'list_price': ['list price', 'list_price', 'mrp', 'maximum retail price', 'mrp inr', 'maximum retail price (sell on amazon, in)'],
    'main_image_url': ['main image url', 'main_image_url', 'main product image url', 'main image'],
    'generic_keywords': ['generic keywords', 'generic_keywords', 'search terms', 'keywords'],
    'description': ['product description', 'product_description', 'description'],
    'department_name': ['department name', 'department_name', 'department'],
    'target_gender': ['target gender', 'target_gender', 'gender'],
    'age_range_description': ['age range description', 'age_range_description', 'age range'],
    'material': ['outer material type', 'outer_material_type', 'material type', 'material'],
    'fabric_type': ['fabric type', 'fabric_type', 'fabric'],
    'care_instructions': ['care instructions', 'care_instructions'],
    'country_of_origin': ['country of origin', 'country_of_origin', 'country'],
    'pattern_type': ['pattern type', 'pattern_type', 'pattern'],
    'fit_type': ['fit type', 'fit_type', 'fit'],
    'style_name': ['style name', 'style_name', 'style'],
    'seasons': ['seasons', 'season'],
    'lifestyle': ['lifestyle', 'life style'],
    'unit_count': ['unit count', 'unit_count', 'number of items'],
    'special_size_type': ['special size type', 'special_size_type', 'special size'],
    'item_weight': ['item weight', 'item_weight', 'weight'],
    'item_weight_unit': ['item weight unit of measure', 'item_weight_unit_of_measure', 'weight unit of measure', 'item weight unit'],
    'hsn_code': ['hsn code', 'hsn_code', 'hsn', 'tax code', 'external product information'],
    'pocket_description': ['pocket description', 'pocket_description', 'pocket type'],
    'sleeve_length': ['sleeve length', 'sleeve_length', 'sleeve length description'],
    'waist_size': ['waist size', 'waist_size', 'waist'],
    'waist_unit': ['waist size unit of measure', 'waist_size_unit_of_measure', 'waist unit of measure'],
}

def _norm_amazon_key(k):
    if not k:
        return ""
    s = str(k).strip().lower()
    s = re.sub(r'\[(marketplace_id|language_tag|content_language|content_type)=[^\]]*\]', '', s)
    s = s.replace('[audience=all]', '')
    if s.startswith('rtip_'):
        s = s[5:]
    return s

VC_TO_SC_KEY_ALIASES = {
    'vendor_sku#1.value': 'contribution_sku#1.value',
    'external_product_id#1.type': 'amzn1.volt.ca.product_id_type',
    'external_product_id#1.value': 'amzn1.volt.ca.product_id_value',
    'merchant_suggested_asin#1.value': 'amzn1.volt.ca.product_id_value',
    'list_price#1.value_with_tax': 'purchasable_offer#1.maximum_retail_price#1.schedule#1.value_with_tax',
}

VC_TO_SC_KEY_FALLBACK_GROUPS = [
    ['apparel_size#1.size_system', 'bottoms_size#1.size_system', 'shirt_size#1.size_system', 'shapewear_size#1.size_system'],
    ['apparel_size#1.size_class', 'bottoms_size#1.size_class', 'shirt_size#1.size_class', 'shapewear_size#1.size_class'],
    ['apparel_size#1.size', 'bottoms_size#1.size', 'shirt_size#1.size', 'shapewear_size#1.size'],
    ['apparel_size#1.size_to', 'bottoms_size#1.size_to', 'shirt_size#1.size_to', 'shapewear_size#1.size_to'],
    ['apparel_size#1.body_type', 'bottoms_size#1.body_type', 'shirt_size#1.body_type', 'shapewear_size#1.body_type'],
    ['apparel_size#1.height_type', 'bottoms_size#1.height_type', 'shirt_size#1.height_type', 'shapewear_size#1.height_type'],
]

VC_TO_SC_LABEL_ALIASES = {
    'vendor sku': 'sku',
    'maximum retail price (mrp)': 'maximum retail price (sell on amazon, in)',
    'external product id type': 'product id type',
    'external product id': 'product id',
    'merchant suggested asin': 'product id',
}

PARENT_INHERITED_LABELS = {
    'product type', 'brand name', 'product description', 'bullet point',
    'generic keywords', 'material', 'fabric type', 'fit type', 'pocket description',
    'care instructions', 'manufacturer contact information', 'pattern',
    'external product information entity', 'external product information',
    'seasons', 'packer contact information', 'sleeve length description',
    'sleeve type', 'item weight', 'item weight unit', 'item length', 'item length unit',
    'item width', 'item width unit', 'item height', 'item height unit',
    'item package length', 'package length unit', 'item package width',
    'package width unit', 'item package height', 'package height unit',
    'package weight', 'package weight unit', 'country of origin', 'dangerous goods regulations',
}

def _build_parent_title_from_child(child_title, parent_sku):
    if not child_title:
        return f"({parent_sku})" if parent_sku else ""
    t = str(child_title).strip()
    if not parent_sku:
        return t
    if re.search(r'\([^()]*\)\s*$', t):
        return re.sub(r'\([^()]*\)\s*$', f'({parent_sku})', t)
    return f"{t} ({parent_sku})"

def _detect_amz_sheet_rows(ws, default_lbl=3, default_key=4, default_data=7):
    r1_str = " ".join(str(ws.cell(1, c).value or '') for c in range(1, min(15, ws.max_column + 1)))
    lbl_m = re.search(r'labelRow=(\d+)', r1_str)
    key_m = re.search(r'attributeRow=(\d+)', r1_str)
    dat_m = re.search(r'dataRow=(\d+)', r1_str)
    if lbl_m and key_m and dat_m:
        return int(lbl_m.group(1)), int(key_m.group(1)), int(dat_m.group(1))
    for r in range(1, min(10, ws.max_row + 1)):
        vals = [str(ws.cell(r, c).value or '').strip().lower() for c in range(1, min(40, ws.max_column + 1))]
        if any(v in ('vendor sku', 'sku', 'seller sku') for v in vals) and any(v in ('item name', 'product type', 'brand name') for v in vals):
            lbl_row = r
            next_vals = [str(ws.cell(r + 1, c).value or '').strip().lower() for c in range(1, min(20, ws.max_column + 1))]
            has_keys = any(('#' in v or '::' in v or '_sku' in v) for v in next_vals if v)
            key_row = r + 1 if has_keys else None
            return lbl_row, key_row, max(default_data, (key_row or lbl_row) + 1)
    return default_lbl, default_key, default_data

def run_seller_central(uploaded_file, ctx):
    wb = load_workbook_clean(uploaded_file, uploaded_file.name)
    ws = None
    for sname in wb.sheetnames:
        if sname.strip().lower() == 'template':
            ws = wb[sname]
            break
    if ws is None:
        ws = pick_template_sheet(wb)

    try:
        ungroup_worksheet(ws)
    except Exception:
        pass

    sc_lbl_r, sc_key_r, sc_data_r = _detect_amz_sheet_rows(ws, 4, 5, 7)

    sc_key_to_col = {}
    sc_lbl_occ_to_col = {}
    sc_lbl_counts = {}
    sc_col_to_lbl = {}
    for c in range(1, ws.max_column + 1):
        lbl = str(ws.cell(sc_lbl_r, c).value or '').strip().lower()
        key = str(ws.cell(sc_key_r, c).value or '').strip() if sc_key_r else ''
        nk = _norm_amazon_key(key)
        sc_col_to_lbl[c] = lbl
        if nk and nk not in sc_key_to_col:
            sc_key_to_col[nk] = c
        if lbl:
            occ = sc_lbl_counts.get(lbl, 0) + 1
            sc_lbl_counts[lbl] = occ
            sc_lbl_occ_to_col[(lbl, occ)] = c

    def get_sc_col(labels=(), keys=(), occ=1):
        for k in keys:
            nk = _norm_amazon_key(k)
            if nk in sc_key_to_col:
                return sc_key_to_col[nk]
        for l in labels:
            ll = l.strip().lower()
            if (ll, occ) in sc_lbl_occ_to_col:
                return sc_lbl_occ_to_col[(ll, occ)]
        return None

    c_sku = get_sc_col(['sku', 'seller sku'], ['contribution_sku#1.value']) or 1
    c_prod_type = get_sc_col(['product type', 'feed product type'], ['product_type#1.value'])
    c_action = get_sc_col(['listing action'], ['::record_action'])
    c_parentage = get_sc_col(['parentage level'], ['parentage_level#1.value'])
    c_parent_sku = get_sc_col(['parent sku'], ['child_parent_sku_relationship#1.parent_sku'])
    c_var_theme = get_sc_col(['variation theme name'], ['variation_theme#1.name'])
    c_title = get_sc_col(['item name', 'product name', 'title'], ['item_name#1.value'])
    c_brand = get_sc_col(['brand name', 'brand'], ['brand#1.value'])
    c_prod_id_type = get_sc_col(['product id type', 'external product id type'], ['amzn1.volt.ca.product_id_type'])
    c_prod_id_val = get_sc_col(['product id', 'external product id'], ['amzn1.volt.ca.product_id_value'])
    c_mod_num = get_sc_col(['model number'], ['model_number#1.value'])
    c_mod_name = get_sc_col(['model name'], ['model_name#1.value'])
    c_mfg = get_sc_col(['manufacturer'], ['manufacturer#1.value'])
    c_main_img = get_sc_col(['main image url'], ['main_product_image_locator#1.media_location'])
    c_other_imgs = [get_sc_col(['other image url'], [f'other_product_image_locator_{i}#1.media_location'], occ=i) for i in range(1, 9)]
    c_desc = get_sc_col(['product description', 'description'], ['product_description#1.value'])
    c_bullets = [get_sc_col(['bullet point'], [f'bullet_point#{i}.value'], occ=i) for i in range(1, 6)]
    c_gkw = get_sc_col(['generic keywords', 'search terms'], ['generic_keyword#1.value'])
    c_lifestyle = get_sc_col(['lifestyle'], ['lifestyle#1.value'])
    c_style = get_sc_col(['style', 'style name'], ['style#1.value'])
    c_dept = get_sc_col(['department name', 'department'], ['department#1.value'])
    c_gender = get_sc_col(['target gender', 'gender'], ['target_gender#1.value'])
    c_age = get_sc_col(['age range description', 'age range'], ['age_range_description#1.value'])
    c_ssys = get_sc_col(['apparel size system', 'bottoms size system', 'size system'], ['apparel_size#1.size_system', 'bottoms_size#1.size_system'])
    c_scls = get_sc_col(['apparel size class', 'bottoms size class', 'size class'], ['apparel_size#1.size_class', 'bottoms_size#1.size_class'])
    c_sval = get_sc_col(['apparel size value', 'bottoms size value', 'size value', 'size name', 'size'], ['apparel_size#1.size', 'bottoms_size#1.size'])
    c_mat = get_sc_col(['material', 'outer material type'], ['material#1.value'])
    c_fab = get_sc_col(['fabric type'], ['fabric_type#1.value'])
    c_noi = get_sc_col(['number of items'], ['number_of_items#1.value'])
    c_itn = get_sc_col(['item type name'], ['item_type_name#1.value'])
    c_spec_sz = get_sc_col(['special size', 'special size type'], ['special_size_type#1.value'])
    c_cmap = get_sc_col(['color map'], ['color#1.standardized_values#1'])
    c_color = get_sc_col(['color', 'color name'], ['color#1.value'])
    c_part = get_sc_col(['part number'], ['part_number#1.value'])
    c_fit = get_sc_col(['fit type'], ['fit_type#1.value'])
    c_poc = get_sc_col(['pocket description'], ['pocket_description#1.value'])
    c_care = get_sc_col(['care instructions'], ['care_instructions#1.value'])
    c_mfg_cnt = get_sc_col(['manufacturer contact information'], ['manufacturer_contact_information#1.value', 'rtip_manufacturer_contact_information#1.value'])
    c_pat = get_sc_col(['pattern', 'pattern type'], ['pattern#1.value'])
    c_uc = get_sc_col(['unit count'], ['unit_count#1.value'])
    c_uct = get_sc_col(['unit count type'], ['unit_count#1.type.value'])
    c_pld = get_sc_col(['product site launch date'], ['product_site_launch_date#1.value'])
    c_hsn_ent = get_sc_col(['external product information entity'], ['external_product_information#1.entity'])
    c_hsn_val = get_sc_col(['external product information', 'hsn code'], ['external_product_information#1.value'])
    c_sea = get_sc_col(['seasons'], ['seasons#1.value'])
    c_pkr = get_sc_col(['packer contact information'], ['packer_contact_information#1.value'])
    c_sld = get_sc_col(['sleeve length description', 'sleeve length'], ['sleeve#1.length_description#1.value'])
    c_st = get_sc_col(['sleeve type'], ['sleeve#1.type#1.value'])
    c_iwt = get_sc_col(['item weight'], ['item_weight#1.value'])
    c_iwtu = get_sc_col(['item weight unit', 'item weight unit of measure'], ['item_weight#1.unit'])
    c_cond = get_sc_col(['item condition'], ['condition_type#1.value'])
    c_fulfill = get_sc_col(['fulfillment channel code (in)'], ['fulfillment_availability#1.fulfillment_channel_code'])
    c_qty = get_sc_col(['quantity (in)', 'quantity'], ['fulfillment_availability#1.quantity'])
    c_price = get_sc_col(['your price inr (sell on amazon, in)', 'standard price', 'selling price'], ['purchasable_offer#1.our_price#1.schedule#1.value_with_tax'])
    c_mrp = get_sc_col(['maximum retail price (sell on amazon, in)', 'list price', 'mrp'], ['purchasable_offer#1.maximum_retail_price#1.schedule#1.value_with_tax'])
    c_ship_tpl = get_sc_col(['shipping template (in)'], ['merchant_shipping_group#1.value'])
    c_il = get_sc_col(['item length'], ['item_dimensions#1.length.value'])
    c_ilu = get_sc_col(['item length unit'], ['item_dimensions#1.length.unit'])
    c_iw = get_sc_col(['item width'], ['item_dimensions#1.width.value'])
    c_iwu = get_sc_col(['item width unit'], ['item_dimensions#1.width.unit'])
    c_ih = get_sc_col(['item height'], ['item_dimensions#1.height.value'])
    c_ihu = get_sc_col(['item height unit'], ['item_dimensions#1.height.unit'])
    c_pl = get_sc_col(['item package length'], ['item_package_dimensions#1.length.value'])
    c_plu = get_sc_col(['package length unit'], ['item_package_dimensions#1.length.unit'])
    c_pw = get_sc_col(['item package width'], ['item_package_dimensions#1.width.value'])
    c_pwu = get_sc_col(['package width unit'], ['item_package_dimensions#1.width.unit'])
    c_ph = get_sc_col(['item package height'], ['item_package_dimensions#1.height.value'])
    c_phu = get_sc_col(['package height unit'], ['item_package_dimensions#1.height.unit'])
    c_pwt = get_sc_col(['package weight'], ['item_package_weight#1.value'])
    c_pwtu = get_sc_col(['package weight unit'], ['item_package_weight#1.unit'])
    c_coo = get_sc_col(['country of origin'], ['country_of_origin#1.value'])
    c_dg = get_sc_col(['dangerous goods regulations'], ['supplier_declared_dg_hz_regulation#1.value'])

    parent_inherited_sc_cols = {
        c for c, lbl in sc_col_to_lbl.items() if lbl in PARENT_INHERITED_LABELS
    }

    curr_r = sc_data_r
    while curr_r <= ws.max_row and ws.cell(curr_r, c_sku).value not in (None, ''):
        curr_r += 1

    today_str = datetime.now().strftime("%Y-%m-%d")
    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)
    material_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")

    action_val = "Create or Replace (Full Update)"
    parent_val = pick_dropdown_value(ws, c_parentage, "Parent", "Parent") if c_parentage else "Parent"
    child_val = pick_dropdown_value(ws, c_parentage, "Child", "Child") if c_parentage else "Child"
    var_theme_val = "SIZE/COLOR"
    cond_new_val = "New"
    fulfillment_default_val = "Fulfillment by Merchant (Default)"
    qty_zero_val = 0
    ship_tpl_val = "Migrated Template"
    asin_type_val = "ASIN"

    custom_parent_model = (ctx.get('sc_parent_model') or '').strip()

    generated_rows = []
    model_first_child_cells = {}

    for raw_fld in ctx['folders']:
        clean_style = clean_sku_name(raw_fld)
        det_col = resolve_color(raw_fld, clean_style, ctx)
        cat = resolve_category(clean_style, ctx, ws.title)
        det_gen = resolve_gender(clean_style, ctx)

        model_str = custom_parent_model if custom_parent_model else clean_style
        urls = get_sku_img_urls(raw_fld, ctx, clean_style)
        
        mat_val = resolve_fabric(clean_style, cat, ctx)
        sleeve_val = resolve_sleeve_length(clean_style, cat, ctx)
        
        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
        gen_kw = generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)
        style_val = get_style_from_dropdown(cat)
        item_type_val = get_item_type_name_from_dropdown(cat)
        color_map_val = map_to_standard_color_map(det_col)
        pocket_val = resolve_pocket(clean_style, cat, ctx)
        
        fab_val = f"100% {mat_val}"
        prod_type_val = resolve_sc_product_type_for_row(
            ws, c_prod_type,
            raw_prod_type=SC_PRODUCT_TYPE.get(cat, ''),
            sheet_name=cat,
            item_type=item_type_val,
            model_name=model_str,
            item_name=clean_style
        )

        dept_val = "Womens" if det_gen == "Women" else "Mens"
        gender_val = "Female" if det_gen == "Women" else "Male"

        for sz in ctx['sizes']:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            is_num_sz = clean_sz.isdigit()
            size_class_val = "Numeric" if is_num_sz else "Alpha"
            child_title = build_vendor_title(ctx['brand'], cat, full_sku, ctx['keywords'], 150, det_gen)

            cells = {}
            def s_col(col_idx, val, dropdown=False):
                if not col_idx or val in (None, ''):
                    return
                cells[col_idx] = pick_dropdown_value(ws, col_idx, val, val) if dropdown else val

            s_col(c_sku, full_sku)
            s_col(c_prod_type, prod_type_val)
            s_col(c_title, child_title)
            s_col(c_brand, ctx['brand'])
            s_col(c_mod_num, full_sku)
            s_col(c_mod_name, model_str)
            s_col(c_mfg, ctx['mfg'])
            if c_main_img and urls:
                s_col(c_main_img, urls[0])
            for idx_img, c_oi in enumerate(c_other_imgs):
                if c_oi and idx_img + 1 < len(urls):
                    s_col(c_oi, urls[idx_img + 1])
            s_col(c_desc, desc_text)
            for idx_b, c_bp in enumerate(c_bullets):
                if c_bp and idx_b < len(bullets_list):
                    s_col(c_bp, bullets_list[idx_b])
            s_col(c_gkw, gen_kw)
            s_col(c_lifestyle, "Casual", dropdown=True)
            s_col(c_style, style_val, dropdown=True)
            s_col(c_dept, dept_val, dropdown=True)
            s_col(c_gender, gender_val, dropdown=True)
            s_col(c_age, "Adult", dropdown=True)
            s_col(c_ssys, "IN", dropdown=True)
            s_col(c_scls, size_class_val, dropdown=True)
            s_col(c_sval, clean_sz)
            s_col(c_mat, mat_val, dropdown=True)
            s_col(c_fab, fab_val)
            s_col(c_noi, 1)
            s_col(c_itn, item_type_val, dropdown=True)
            s_col(c_spec_sz, "Standard", dropdown=True)
            s_col(c_cmap, color_map_val, dropdown=True)
            s_col(c_color, det_col)
            s_col(c_part, full_sku)
            s_col(c_fit, "Relaxed", dropdown=True)
            s_col(c_poc, pocket_val, dropdown=True)
            s_col(c_care, "Machine Wash", dropdown=True)
            s_col(c_mfg_cnt, ctx['mfg'])
            s_col(c_pat, "Solid", dropdown=True)
            s_col(c_uc, 1)
            s_col(c_uct, "count", dropdown=True)
            s_col(c_pld, today_str)
            s_col(c_hsn_ent, "HSN Code", dropdown=True)
            s_col(c_hsn_val, get_hsn_code(cat))
            s_col(c_sea, "All", dropdown=True)
            s_col(c_pkr, ctx['mfg'])
            s_col(c_sld, sleeve_val, dropdown=True)
            s_col(c_st, "Regular Sleeve", dropdown=True)
            s_col(c_iwt, 350)
            s_col(c_iwtu, "Grams", dropdown=True)
            if ctx.get('selling_price'):
                s_col(c_price, ctx['selling_price'])
            if ctx.get('mrp'):
                mv = ctx['mrp']
                s_col(c_mrp, int(mv) if str(mv).isdigit() else mv)
            s_col(c_il, 10)
            s_col(c_ilu, "Inches", dropdown=True)
            s_col(c_iw, 2)
            s_col(c_iwu, "Inches", dropdown=True)
            s_col(c_ih, 1.3)
            s_col(c_ihu, "Inches", dropdown=True)
            s_col(c_pl, 10)
            s_col(c_plu, "Inches", dropdown=True)
            s_col(c_pw, 2)
            s_col(c_pwu, "Inches", dropdown=True)
            s_col(c_ph, 1.3)
            s_col(c_phu, "Inches", dropdown=True)
            s_col(c_pwt, 380)
            s_col(c_pwtu, "Grams", dropdown=True)
            s_col(c_coo, "India", dropdown=True)
            s_col(c_dg, "Not Applicable", dropdown=True)

            if model_str not in model_first_child_cells:
                model_first_child_cells[model_str] = cells
            generated_rows.append((full_sku, model_str, cells))

    preview = []
    total_written = 0

    def apply_common_fixed_fields(row_idx):
        if c_action:
            ws.cell(row_idx, c_action).value = action_val
        if c_var_theme:
            ws.cell(row_idx, c_var_theme).value = var_theme_val
        if c_cond:
            ws.cell(row_idx, c_cond).value = cond_new_val
        if c_fulfill:
            ws.cell(row_idx, c_fulfill).value = fulfillment_default_val
        if c_qty:
            ws.cell(row_idx, c_qty).value = qty_zero_val

    def write_parent_row(model_sku):
        nonlocal curr_r, total_written
        ref_cells = model_first_child_cells.get(model_sku) or (generated_rows[0][2] if generated_rows else {})
        for sc_c in parent_inherited_sc_cols:
            if sc_c in ref_cells and ref_cells[sc_c] not in (None, ''):
                ws.cell(curr_r, sc_c).value = ref_cells[sc_c]
        ws.cell(curr_r, c_sku).value = model_sku
        if c_prod_type and c_prod_type in ref_cells:
            ws.cell(curr_r, c_prod_type).value = ref_cells[c_prod_type]
        if c_parentage:
            ws.cell(curr_r, c_parentage).value = parent_val
        parent_title = ""
        if c_title:
            child_title = ref_cells.get(c_title, '')
            parent_title = _build_parent_title_from_child(child_title, model_sku)
            ws.cell(curr_r, c_title).value = parent_title
        apply_common_fixed_fields(curr_r)
        preview.append({
            'Row': curr_r, 'SKU': model_sku, 'Product Type': ref_cells.get(c_prod_type, '') if c_prod_type else '',
            'Listing Action': action_val, 'Parentage Level': parent_val, 'Parent SKU': '',
            'Variation Theme': var_theme_val, 'Shipping Template': '', 'Item Name': parent_title,
        })
        curr_r += 1
        total_written += 1

    def write_child_row(child_sku, model_str, mapped_cells):
        nonlocal curr_r, total_written
        for sc_c, val in mapped_cells.items():
            ws.cell(curr_r, sc_c).value = val
        if c_mat and c_mat in mapped_cells:
            ws.cell(curr_r, c_mat).fill = material_fill
        if c_mrp and not ctx.get('mrp'):
            ws.cell(curr_r, c_mrp).fill = mrp_fill
            ws.cell(curr_r, c_mrp).font = mrp_font
        ws.cell(curr_r, c_sku).value = child_sku
        if c_parentage:
            ws.cell(curr_r, c_parentage).value = child_val
        if c_parent_sku and model_str:
            ws.cell(curr_r, c_parent_sku).value = model_str
        if c_mod_name and model_str:
            ws.cell(curr_r, c_mod_name).value = model_str
        if c_prod_id_type:
            ws.cell(curr_r, c_prod_id_type).value = asin_type_val
        if c_ship_tpl:
            ws.cell(curr_r, c_ship_tpl).value = ship_tpl_val
        apply_common_fixed_fields(curr_r)
        preview.append({
            'Row': curr_r, 'SKU': child_sku, 'Product Type': mapped_cells.get(c_prod_type, '') if c_prod_type else '',
            'Listing Action': action_val, 'Parentage Level': child_val, 'Parent SKU': model_str,
            'Variation Theme': var_theme_val, 'Shipping Template': ship_tpl_val,
            'Item Name': mapped_cells.get(c_title, '') if c_title else '',
        })
        curr_r += 1
        total_written += 1

    inserted_parents = set()
    for c_sku_val, m_str, m_cells in generated_rows:
        if m_str and m_str not in inserted_parents:
            inserted_parents.add(m_str)
            write_parent_row(m_str)
        write_child_row(c_sku_val, m_str, m_cells)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), preview, total_written

# ======================================================================
# UNIVERSAL ENGINE (Flipkart / Meesho / Myntra)
# ======================================================================
MP_FIELD_CANDIDATES = {
    'Flipkart': {
        'sku': ['listing sku', 'seller sku', 'sku', 'sku id'],
        'title': ['product title', 'listing title', 'title', 'product name'],
        'brand': ['brand', 'brand name'],
        'description': ['description', 'product description', 'long description'],
        'mrp': ['mrp', 'maximum retail price', 'list price', 'mrp inr'],
        'price': ['selling price', 'price', 'your price'],
        'hsn': ['hsn', 'hsn code'],
        'color': ['color', 'color family', 'colour'],
        'size': ['size', 'size value', 'variant size'],
        'fabric': ['fabric', 'fabric type', 'material'],
        'pattern': ['pattern', 'pattern type'],
        'fit': ['fit', 'fit type'],
        'country': ['country of origin', 'origin country', 'country'],
        'generic_kw': ['search keywords', 'keywords', 'search terms'],
        'gender': ['gender', 'ideal for'],
        'care': ['wash care', 'care instructions'],
        'sleeve': ['sleeve', 'sleeve length'],
        'manufacturer': ['manufactured by', 'manufacturer', 'packer details', 'marketed by'],
        'img1': ['main image url', 'image url 1', 'image 1', 'image1', 'primary image url'],
    },
    'Meesho': {
        'sku': ['sku', 'seller sku', 'sku id', 'variant sku'],
        'title': ['product name', 'title', 'product title'],
        'brand': ['brand', 'brand name'],
        'description': ['description', 'product description'],
        'mrp': ['mrp', 'maximum retail price'],
        'price': ['price', 'selling price', 'supplier price'],
        'hsn': ['hsn', 'hsn code'],
        'color': ['color', 'colour'],
        'size': ['size', 'variant size'],
        'fabric': ['fabric', 'material', 'fabric type'],
        'pattern': ['pattern'],
        'fit': ['fit', 'fit type'],
        'country': ['country of origin', 'country'],
        'generic_kw': ['keywords', 'search keywords'],
        'gender': ['gender'],
        'care': ['wash care', 'care instructions'],
        'sleeve': ['sleeve length', 'sleeve'],
        'manufacturer': ['manufactured by', 'manufacturer', 'packer details', 'marketed by'],
        'img1': ['image url 1', 'image 1', 'image1', 'main image url', 'primary image url'],
    },
    'Myntra': {
        'sku': ['seller sku', 'sku', 'sku id', 'style code'],
        'title': ['style name', 'product name', 'title', 'item name'],
        'brand': ['brand', 'brand name', 'vendor brand'],
        'description': ['style description', 'description', 'product description'],
        'mrp': ['mrp', 'maximum retail price', 'list price'],
        'price': ['selling price', 'price', 'wholesale price'],
        'hsn': ['hsn', 'hsn code'],
        'color': ['color', 'colour', 'primary color'],
        'size': ['size', 'variant size'],
        'fabric': ['fabric', 'material', 'composition'],
        'pattern': ['pattern', 'print'],
        'fit': ['fit', 'fit type'],
        'country': ['country of origin', 'country'],
        'generic_kw': ['keywords', 'search keywords'],
        'gender': ['gender', 'customer gender'],
        'care': ['wash care', 'care instructions'],
        'sleeve': ['sleeve length', 'sleeve styling'],
        'manufacturer': ['manufactured by', 'manufacturer', 'marketed by', 'packer details'],
        'img1': ['image 1', 'image1', 'images1', 'main image url', 'primary image url'],
    },
}

def run_generic_marketplace(uploaded_file, ctx, mp):
    wb = load_workbook_clean(uploaded_file, uploaded_file.name)
    fields = MP_FIELD_CANDIDATES.get(mp, MP_FIELD_CANDIDATES['Flipkart'])
    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)

    best_ws, best_row, best_score = None, None, -1
    sku_cands = [_norm_h(x) for x in fields['sku']]
    for sname in wb.sheetnames:
        ws_try = wb[sname]
        for r in range(1, min(ws_try.max_row, 10) + 1):
            vals = [_norm_h(ws_try.cell(r, c).value) for c in range(1, min(ws_try.max_column, 60) + 1)]
            score = sum(1 for v in vals if v in sku_cands)
            if score > best_score:
                best_score, best_ws, best_row = score, ws_try, r
    if best_ws is None or best_score == 0:
        raise ValueError(f"{mp} template me 'sku' header wali sheet nahi mili.")

    ws = best_ws
    try:
        ungroup_worksheet(ws)
    except Exception:
        pass
    header_row = best_row

    cols = {}
    for c in range(1, ws.max_column + 1):
        h = _norm_h(ws.cell(header_row, c).value)
        if h and h not in cols:
            cols[h] = c

    def find(cands):
        for cd in cands:
            n = _norm_h(cd)
            if n in cols:
                return cols[n]
        return None

    fcol = {k: find(v) for k, v in fields.items()}
    sku_col = fcol['sku'] or (find(['sku']) or 1)

    img_cols = []
    for c in range(1, ws.max_column + 1):
        h = _norm_h(ws.cell(header_row, c).value)
        if re.match(r'(image|img|images|model)\d*$', h or ''):
            img_cols.append(c)
    if fcol.get('img1') and fcol['img1'] not in img_cols:
        img_cols.insert(0, fcol['img1'])
    img_cols = sorted(set(img_cols))

    bullet_cols = []
    for c in range(1, ws.max_column + 1):
        h = _norm_h(ws.cell(header_row, c).value)
        if re.match(r'(bulletpoint|keyfeature|highlight)\d*$', h or ''):
            bullet_cols.append(c)
    bullet_cols = sorted(set(bullet_cols))

    start_row = header_row + 1
    while start_row <= ws.max_row and ws.cell(start_row, sku_col).value not in (None, ''):
        start_row += 1

    def W(row, key, value, dropdown=False):
        c = fcol.get(key)
        if not c or value in (None, ''):
            return
        cell = ws.cell(row, c)
        cell.value = pick_dropdown_value(ws, c, value, value) if dropdown else value

    count, preview = 0, []
    for raw_fld in ctx['folders']:
        clean_style = clean_sku_name(raw_fld)
        det_col = resolve_color(raw_fld, clean_style, ctx)
        cat = resolve_category(clean_style, ctx, ws.title)
        det_gen = resolve_gender(clean_style, ctx)
        color_family_val = map_to_standard_color_map(det_col)
        
        mat_val = resolve_fabric(clean_style, cat, ctx)
        sleeve_val = resolve_sleeve_length(clean_style, cat, ctx)

        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
        gen_kw = re.sub(r'\s+', ' ', str(generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)).replace(';', ' ')).strip()
        urls = get_sku_img_urls(raw_fld, ctx, clean_style)

        preview.append({'SKU': clean_style, 'Category': cat, 'Gender': det_gen, 'Color': det_col,
                        'Color Family': color_family_val, 'Material': mat_val, 'Sleeve': sleeve_val})

        for sz in ctx['sizes']:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            r = start_row + count
            title = build_sc_title(ctx['brand'], cat, full_sku, ctx['keywords'], 150, det_gen)

            W(r, 'sku', full_sku)
            W(r, 'title', title)
            W(r, 'brand', ctx['brand'])
            W(r, 'description', desc_text)
            W(r, 'color', color_family_val, dropdown=True)
            W(r, 'size', clean_sz)
            W(r, 'fabric', mat_val, dropdown=True)
            W(r, 'pattern', 'Solid', dropdown=True)
            W(r, 'fit', 'Relaxed', dropdown=True)
            W(r, 'country', 'India', dropdown=True)
            W(r, 'generic_kw', gen_kw)
            W(r, 'gender', det_gen, dropdown=True)
            W(r, 'care', 'Machine Wash', dropdown=True)
            W(r, 'sleeve', sleeve_val, dropdown=True)
            W(r, 'hsn', get_hsn_code(cat))
            W(r, 'manufacturer', ctx['mfg'])

            for bi, bc in enumerate(bullet_cols):
                if bi < len(bullets_list):
                    ws.cell(r, bc).value = bullets_list[bi]

            for ii, ic in enumerate(img_cols):
                if ii < len(urls):
                    ws.cell(r, ic).value = urls[ii]

            if fcol.get('mrp'):
                cell = ws.cell(r, fcol['mrp'])
                if ctx['mrp']:
                    cell.value = ctx['mrp']
                else:
                    cell.fill = mrp_fill
                    cell.font = mrp_font
            W(r, 'price', ctx['selling_price'] or None)

            count += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), preview, count

# ======================================================================
# VC ➜ SC CONVERT ENGINE (MULTI-SHEET + SMART PRODUCT TYPE + DEDUPE MERGE)
# ======================================================================
IGNORED_AMZ_SHEET_KEYWORDS = (
    'instruction', 'guideline', 'image', 'browse data', 'dropdown',
    'valid value', 'data definition', 'definition', 'example',
    'change log', 'changelog', 'readme', 'help', 'hidden', 'lookup'
)

def _is_valid_vc_listing_sheet(ws):
    """
    Checks whether a worksheet in the Vendor Central workbook is a product listing sheet
    (has SKU/Vendor SKU headers and at least one product data row).
    """
    sname_low = str(ws.title or '').strip().lower()
    if any(k in sname_low for k in IGNORED_AMZ_SHEET_KEYWORDS):
        return False
    if ws.max_row < 4 or ws.max_column < 3:
        return False

    lbl_r, key_r, data_r = _detect_amz_sheet_rows(ws, 3, 4, 7)
    has_sku_header = False
    has_item_header = False
    max_scan_c = min(ws.max_column, 80)

    for c in range(1, max_scan_c + 1):
        lbl = str(ws.cell(lbl_r, c).value or '').strip().lower()
        key = _norm_amazon_key(ws.cell(key_r, c).value) if key_r else ''
        if key in ('vendor_sku#1.value', 'contribution_sku#1.value') or lbl in ('vendor sku', 'sku', 'seller sku'):
            has_sku_header = True
        if key in ('item_name#1.value', 'brand#1.value', 'product_type#1.value') or lbl in ('item name', 'brand name', 'product type', 'model name'):
            has_item_header = True

    return has_sku_header and has_item_header

def _derive_parent_model_from_sku(sku_str, fallback_prefix=""):
    """
    If Model Name is empty in a VC row, derives a clean parent SKU by stripping size suffix.
    """
    s = str(sku_str or '').strip()
    if not s:
        return fallback_prefix or "PARENT-SKU"
    stripped = re.sub(r'[-_](XXS|XS|S|M|L|XL|XXL|2XL|3XL|4XL|5XL|FREE|FREESIZE|\d{2})$', '', s, flags=re.IGNORECASE)
    return stripped if stripped else s

def _merge_vc_duplicate_cell_maps(existing_cells, new_cells, existing_score, new_score):
    """
    Merges two mapped_cells dicts for the same SKU when duplicate rows exist in VC
    (e.g., one row has only SKU/ASIN and blank details, while another row has full details).
    Keeps all non-empty values and prioritizes the row with richer data.
    """
    if new_score > existing_score:
        primary, secondary = new_cells, existing_cells
    else:
        primary, secondary = existing_cells, new_cells

    merged = dict(secondary)
    for k, v in primary.items():
        if v not in (None, ''):
            merged[k] = v
    return merged

def read_vendor_central_records(vc_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(vc_bytes), data_only=True)
    records_by_sku = {}
    sku_order = []
    candidate_sheets = [wb[s] for s in wb.sheetnames if _is_valid_vc_listing_sheet(wb[s])]
    if not candidate_sheets:
        candidate_sheets = [pick_template_sheet(wb)]

    for ws in candidate_sheets:
        label_row, key_row, r = _detect_amz_sheet_rows(ws, 3, 4, 7)
        col_meta, sku_col = {}, None
        for c in range(1, ws.max_column + 1):
            lbl = str(ws.cell(label_row, c).value or '').strip()
            key = str(ws.cell(key_row, c).value or '').strip() if key_row else ''
            if not lbl and not key:
                continue
            col_meta[c] = (key, lbl)
            if (key.lower().startswith('vendor_sku') or lbl.lower() in ('vendor sku', 'sku', 'seller sku')) and sku_col is None:
                sku_col = c
        if sku_col is None:
            sku_col = 2

        for row_idx in range(r, ws.max_row + 1):
            raw_sku = ws.cell(row_idx, sku_col).value
            if raw_sku in (None, ''):
                continue
            sku_norm = str(raw_sku).strip().lower()
            if sku_norm in ('vendor sku', 'sku', 'seller sku', 'example'):
                continue

            rec = {'_sheet_name': ws.title}
            for c, (key, lbl) in col_meta.items():
                v = ws.cell(row_idx, c).value
                if v in (None, ''):
                    continue
                k1, k2 = key.strip().lower(), lbl.strip().lower()
                if k1 and k1 not in rec:
                    rec[k1] = v
                if k2 and k2 not in rec:
                    rec[k2] = v

            # Deduplicate & merge by SKU: ignore blank duplicate rows and keep filled values
            if sku_norm not in records_by_sku:
                records_by_sku[sku_norm] = rec
                sku_order.append(sku_norm)
            else:
                old_rec = records_by_sku[sku_norm]
                if len(rec) > len(old_rec):
                    merged = dict(old_rec)
                    merged.update({k: v for k, v in rec.items() if v not in (None, '')})
                    records_by_sku[sku_norm] = merged
                else:
                    for k, v in rec.items():
                        if k not in old_rec or old_rec[k] in (None, ''):
                            old_rec[k] = v

    return [records_by_sku[k] for k in sku_order]

def convert_vc_to_sc(vc_file, sc_file, ctx):
    wb_vc = openpyxl.load_workbook(io.BytesIO(vc_file.getvalue()), data_only=True)
    wb_sc = load_workbook_clean(io.BytesIO(sc_file.getvalue()), sc_file.name)

    ws_sc = None
    for sname in wb_sc.sheetnames:
        if sname.strip().lower() == 'template':
            ws_sc = wb_sc[sname]
            break
    if ws_sc is None:
        ws_sc = pick_template_sheet(wb_sc)

    try:
        ungroup_worksheet(ws_sc)
    except Exception:
        pass

    sc_lbl_r, sc_key_r, sc_data_r = _detect_amz_sheet_rows(ws_sc, 4, 5, 7)

    sc_key_to_col = {}
    sc_lbl_occ_to_col = {}
    sc_lbl_counts = {}
    sc_col_to_lbl = {}
    for c in range(1, ws_sc.max_column + 1):
        lbl = str(ws_sc.cell(sc_lbl_r, c).value or '').strip().lower()
        key = str(ws_sc.cell(sc_key_r, c).value or '').strip() if sc_key_r else ''
        nk = _norm_amazon_key(key)
        sc_col_to_lbl[c] = lbl
        if nk and nk not in sc_key_to_col:
            sc_key_to_col[nk] = c
        if lbl:
            occ = sc_lbl_counts.get(lbl, 0) + 1
            sc_lbl_counts[lbl] = occ
            sc_lbl_occ_to_col[(lbl, occ)] = c

    sc_sku_col = (
        sc_key_to_col.get('contribution_sku#1.value')
        or sc_lbl_occ_to_col.get(('sku', 1))
        or sc_lbl_occ_to_col.get(('seller sku', 1))
        or 1
    )
    sc_prod_type_col = (
        sc_key_to_col.get('product_type#1.value')
        or sc_lbl_occ_to_col.get(('product type', 1))
        or sc_lbl_occ_to_col.get(('feed product type', 1))
    )
    sc_action_col = (
        sc_key_to_col.get('::record_action')
        or sc_lbl_occ_to_col.get(('listing action', 1))
    )
    sc_parentage_col = (
        sc_key_to_col.get('parentage_level#1.value')
        or sc_lbl_occ_to_col.get(('parentage level', 1))
    )
    sc_parent_sku_col = (
        sc_key_to_col.get('child_parent_sku_relationship#1.parent_sku')
        or sc_lbl_occ_to_col.get(('parent sku', 1))
    )
    sc_var_theme_col = (
        sc_key_to_col.get('variation_theme#1.name')
        or sc_lbl_occ_to_col.get(('variation theme name', 1))
    )
    sc_title_col = (
        sc_key_to_col.get('item_name#1.value')
        or sc_lbl_occ_to_col.get(('item name', 1))
    )
    sc_prod_id_type_col = (
        sc_key_to_col.get('amzn1.volt.ca.product_id_type')
        or sc_lbl_occ_to_col.get(('product id type', 1))
    )
    sc_prod_id_val_col = (
        sc_key_to_col.get('amzn1.volt.ca.product_id_value')
        or sc_lbl_occ_to_col.get(('product id', 1))
    )
    sc_model_name_col = (
        sc_key_to_col.get('model_name#1.value')
        or sc_lbl_occ_to_col.get(('model name', 1))
    )
    sc_condition_col = (
        sc_key_to_col.get('condition_type#1.value')
        or sc_lbl_occ_to_col.get(('item condition', 1))
    )
    sc_fulfillment_col = (
        sc_key_to_col.get('fulfillment_availability#1.fulfillment_channel_code')
        or sc_lbl_occ_to_col.get(('fulfillment channel code (in)', 1))
    )
    sc_quantity_col = (
        sc_key_to_col.get('fulfillment_availability#1.quantity')
        or sc_lbl_occ_to_col.get(('quantity (in)', 1))
        or sc_lbl_occ_to_col.get(('quantity', 1))
    )
    sc_shipping_tpl_col = (
        sc_key_to_col.get('merchant_shipping_group#1.value')
        or sc_lbl_occ_to_col.get(('shipping template (in)', 1))
    )
    sc_price_col = (
        sc_key_to_col.get('purchasable_offer#1.our_price#1.schedule#1.value_with_tax')
        or sc_lbl_occ_to_col.get(('your price inr (sell on amazon, in)', 1))
        or sc_lbl_occ_to_col.get(('standard price', 1))
    )
    sc_mrp_col = (
        sc_key_to_col.get('purchasable_offer#1.maximum_retail_price#1.schedule#1.value_with_tax')
        or sc_lbl_occ_to_col.get(('maximum retail price (sell on amazon, in)', 1))
        or sc_lbl_occ_to_col.get(('list price', 1))
    )

    parent_inherited_sc_cols = {
        c for c, lbl in sc_col_to_lbl.items() if lbl in PARENT_INHERITED_LABELS
    }

    curr_r = sc_data_r
    while curr_r <= ws_sc.max_row and ws_sc.cell(curr_r, sc_sku_col).value not in (None, ''):
        curr_r += 1

    # Discover all valid VC listing sheets (supports 1, 2, 4, 5, or N category sheets)
    vc_listing_sheets = [wb_vc[s] for s in wb_vc.sheetnames if _is_valid_vc_listing_sheet(wb_vc[s])]
    if not vc_listing_sheets:
        vc_listing_sheets = [pick_template_sheet(wb_vc)]

    # Deduplication & Smart Merge store across all VC rows:
    # If the same SKU appears twice (one blank/partial row + one filled row), merge them into 1 complete row.
    deduped_sku_order = []
    deduped_sku_store = {}
    sheets_summary = []
    custom_parent_override = (ctx.get('sc_parent_model') or '').strip()

    for ws_vc in vc_listing_sheets:
        vc_lbl_r, vc_key_r, vc_data_r = _detect_amz_sheet_rows(ws_vc, 3, 4, 7)

        vc_to_sc_col = {}
        vc_lbl_counts = {}
        vc_sku_col = None
        vc_model_name_col = None
        vc_asin_col = None
        vc_prod_type_col = None
        vc_item_type_col = None
        vc_title_col = None

        for c in range(1, ws_vc.max_column + 1):
            lbl = str(ws_vc.cell(vc_lbl_r, c).value or '').strip().lower()
            key = str(ws_vc.cell(vc_key_r, c).value or '').strip() if vc_key_r else ''
            nk = _norm_amazon_key(key)

            if vc_sku_col is None and (nk in ('vendor_sku#1.value', 'contribution_sku#1.value') or lbl in ('vendor sku', 'sku', 'seller sku')):
                vc_sku_col = c
            if vc_model_name_col is None and (nk == 'model_name#1.value' or lbl == 'model name'):
                vc_model_name_col = c
            if vc_asin_col is None and (nk in ('merchant_suggested_asin#1.value', 'external_product_id#1.value') or lbl in ('merchant suggested asin', 'external product id')):
                vc_asin_col = c
            if vc_prod_type_col is None and (nk == 'product_type#1.value' or lbl in ('product type', 'feed product type')):
                vc_prod_type_col = c
            if vc_item_type_col is None and (nk == 'item_type_name#1.value' or lbl in ('item type name', 'item type keyword')):
                vc_item_type_col = c
            if vc_title_col is None and (nk == 'item_name#1.value' or lbl in ('item name', 'title', 'product name')):
                vc_title_col = c

            nk_target = VC_TO_SC_KEY_ALIASES.get(nk, nk)
            occ = 1
            if lbl:
                occ = vc_lbl_counts.get(lbl, 0) + 1
                vc_lbl_counts[lbl] = occ
            sc_c = None
            if nk_target and nk_target in sc_key_to_col:
                sc_c = sc_key_to_col[nk_target]
            else:
                for grp in VC_TO_SC_KEY_FALLBACK_GROUPS:
                    if nk_target in grp:
                        for alt_k in grp:
                            if alt_k in sc_key_to_col:
                                sc_c = sc_key_to_col[alt_k]
                                break
                        if sc_c:
                            break
                if not sc_c:
                    lbl_target = VC_TO_SC_LABEL_ALIASES.get(lbl, lbl)
                    if (lbl_target, occ) in sc_lbl_occ_to_col:
                        sc_c = sc_lbl_occ_to_col[(lbl_target, occ)]
                    elif (lbl_target, 1) in sc_lbl_occ_to_col and occ == 1:
                        sc_c = sc_lbl_occ_to_col[(lbl_target, 1)]
            if sc_c:
                vc_to_sc_col.setdefault(c, []).append(sc_c)

        if vc_sku_col is None:
            vc_sku_col = 2

        sheet_raw_rows = 0
        sheet_unique_skus = set()
        sheet_duplicates_merged = 0
        sheet_detected_pts = set()

        for r_vc in range(vc_data_r, ws_vc.max_row + 1):
            sku_val = ws_vc.cell(r_vc, vc_sku_col).value
            if sku_val in (None, ''):
                continue
            sku_str = str(sku_val).strip()
            if sku_str.lower() in ('vendor sku', 'sku', 'seller sku', 'example'):
                continue

            sheet_raw_rows += 1
            sku_norm_key = sku_str.lower()

            # Count how many actual non-empty cells this VC row has (to identify blank vs filled row)
            non_empty_vc_cells = 0
            for c_chk in range(1, ws_vc.max_column + 1):
                if ws_vc.cell(r_vc, c_chk).value not in (None, ''):
                    non_empty_vc_cells += 1

            mapped_cells = {}
            for vc_c, sc_cols_list in vc_to_sc_col.items():
                val = ws_vc.cell(r_vc, vc_c).value
                if val not in (None, ''):
                    for sc_c in (sc_cols_list if isinstance(sc_cols_list, list) else [sc_cols_list]):
                        mapped_cells[sc_c] = val

            if vc_asin_col and sc_prod_id_val_col:
                asin_val = ws_vc.cell(r_vc, vc_asin_col).value
                if asin_val not in (None, ''):
                    mapped_cells[sc_prod_id_val_col] = str(asin_val).strip()

            raw_pt_val = ws_vc.cell(r_vc, vc_prod_type_col).value if vc_prod_type_col else ''
            raw_itn_val = ws_vc.cell(r_vc, vc_item_type_col).value if vc_item_type_col else ''
            raw_title_val = ws_vc.cell(r_vc, vc_title_col).value if vc_title_col else ''
            model_name_val = ws_vc.cell(r_vc, vc_model_name_col).value if vc_model_name_col else None
            has_explicit_model = model_name_val not in (None, '') and str(model_name_val).strip() != ''

            if custom_parent_override:
                model_str = custom_parent_override
            elif has_explicit_model:
                model_str = str(model_name_val).strip()
            else:
                model_str = _derive_parent_model_from_sku(sku_str, ws_vc.title)

            if sku_norm_key not in deduped_sku_store:
                deduped_sku_order.append(sku_norm_key)
                sheet_unique_skus.add(sku_norm_key)
                deduped_sku_store[sku_norm_key] = {
                    'sku_str': sku_str,
                    'model_str': model_str,
                    'has_explicit_model': has_explicit_model,
                    'mapped_cells': mapped_cells,
                    'sheet_title': ws_vc.title,
                    'raw_pt_val': raw_pt_val if raw_pt_val not in (None, '') else '',
                    'raw_itn_val': raw_itn_val if raw_itn_val not in (None, '') else '',
                    'raw_title_val': raw_title_val if raw_title_val not in (None, '') else '',
                    'fill_score': non_empty_vc_cells,
                }
            else:
                # Duplicate SKU found! Ignore blank fields and merge filled data into a single record
                sheet_duplicates_merged += 1
                existing = deduped_sku_store[sku_norm_key]
                existing['mapped_cells'] = _merge_vc_duplicate_cell_maps(
                    existing['mapped_cells'],
                    mapped_cells,
                    existing['fill_score'],
                    non_empty_vc_cells
                )
                # If the previous row was blank/missing Model Name and this row has Model Name, adopt it
                if not custom_parent_override:
                    if has_explicit_model and (not existing['has_explicit_model'] or non_empty_vc_cells > existing['fill_score']):
                        existing['model_str'] = model_str
                        existing['has_explicit_model'] = True

                # Merge raw metadata for accurate Product Type detection
                if raw_pt_val not in (None, '') and (not existing['raw_pt_val'] or non_empty_vc_cells > existing['fill_score']):
                    existing['raw_pt_val'] = raw_pt_val
                if raw_itn_val not in (None, '') and (not existing['raw_itn_val'] or non_empty_vc_cells > existing['fill_score']):
                    existing['raw_itn_val'] = raw_itn_val
                if raw_title_val not in (None, '') and (not existing['raw_title_val'] or non_empty_vc_cells > existing['fill_score']):
                    existing['raw_title_val'] = raw_title_val

                if non_empty_vc_cells > existing['fill_score']:
                    existing['fill_score'] = non_empty_vc_cells
                    existing['sheet_title'] = ws_vc.title
                    existing['sku_str'] = sku_str

        if sheet_raw_rows > 0:
            sheets_summary.append({
                'VC Sheet': ws_vc.title,
                'Total Rows Scanned': sheet_raw_rows,
                'Unique SKUs Kept': len(sheet_unique_skus),
                'Duplicate/Blank Rows Merged': sheet_duplicates_merged,
                'SC Product Type(s)': "Auto",
            })

    # Finalize deduplicated rows, apply optional price/MRP overrides, and resolve SC Product Type
    vc_rows_data = []
    model_first_child_cells = {}
    sheet_pts_map = {}

    for sku_norm_key in deduped_sku_order:
        entry = deduped_sku_store[sku_norm_key]
        sku_str = entry['sku_str']
        model_str = entry['model_str']
        mapped_cells = entry['mapped_cells']
        sh_title = entry['sheet_title']

        # Optional UI overrides for Selling Price / MRP if still blank after merging
        if sc_price_col and ctx.get('selling_price') and sc_price_col not in mapped_cells:
            mapped_cells[sc_price_col] = ctx['selling_price']
        if sc_mrp_col and ctx.get('mrp') and sc_mrp_col not in mapped_cells:
            mv = ctx['mrp']
            mapped_cells[sc_mrp_col] = int(mv) if str(mv).isdigit() else mv

        resolved_pt = resolve_sc_product_type_for_row(
            ws_sc,
            sc_prod_type_col,
            raw_prod_type=entry['raw_pt_val'],
            sheet_name=sh_title,
            item_type=entry['raw_itn_val'],
            model_name=model_str,
            item_name=entry['raw_title_val'] or sku_str
        )
        if sc_prod_type_col and resolved_pt:
            mapped_cells[sc_prod_type_col] = resolved_pt
            sheet_pts_map.setdefault(sh_title, set()).add(str(resolved_pt))

        if model_str:
            # Ensure the parent row inherits from the richest child row in its Model Name group
            if model_str not in model_first_child_cells or len(mapped_cells) > len(model_first_child_cells[model_str]):
                model_first_child_cells[model_str] = mapped_cells

        vc_rows_data.append((sku_str, model_str, mapped_cells, sh_title))

    for s_item in sheets_summary:
        pts = sheet_pts_map.get(s_item['VC Sheet'])
        if pts:
            s_item['SC Product Type(s)'] = ", ".join(sorted(pts))

    if not vc_rows_data:
        return None, [], 0, []

    action_val = "Create or Replace (Full Update)"
    parent_val = pick_dropdown_value(ws_sc, sc_parentage_col, "Parent", "Parent") if sc_parentage_col else "Parent"
    child_val = pick_dropdown_value(ws_sc, sc_parentage_col, "Child", "Child") if sc_parentage_col else "Child"
    var_theme_val = "SIZE/COLOR"
    asin_type_val = "ASIN"
    cond_new_val = "New"
    fulfillment_default_val = "Fulfillment by Merchant (Default)"
    qty_zero_val = int(ctx.get('quantity', 0))
    shipping_tpl_val = "Migrated Template"

    parent_position = ctx.get('parent_position', 'before_group')
    seen_models = set()
    unique_models_ordered = []
    model_to_sheet = {}
    for _, m_str, _, sh_title in vc_rows_data:
        if m_str and m_str not in seen_models:
            seen_models.add(m_str)
            unique_models_ordered.append(m_str)
            model_to_sheet[m_str] = sh_title

    preview = []
    total_written = 0

    def apply_common_fixed_fields(row_idx):
        if sc_action_col:
            ws_sc.cell(row_idx, sc_action_col).value = action_val
        if sc_var_theme_col:
            ws_sc.cell(row_idx, sc_var_theme_col).value = var_theme_val
        if sc_condition_col:
            ws_sc.cell(row_idx, sc_condition_col).value = cond_new_val
        if sc_fulfillment_col:
            ws_sc.cell(row_idx, sc_fulfillment_col).value = fulfillment_default_val
        if sc_quantity_col:
            ws_sc.cell(row_idx, sc_quantity_col).value = qty_zero_val

    def write_parent_row(model_sku):
        nonlocal curr_r, total_written
        ref_cells = model_first_child_cells.get(model_sku) or (vc_rows_data[0][2] if vc_rows_data else {})

        for sc_c in parent_inherited_sc_cols:
            if sc_c in ref_cells and ref_cells[sc_c] not in (None, ''):
                ws_sc.cell(curr_r, sc_c).value = ref_cells[sc_c]

        ws_sc.cell(curr_r, sc_sku_col).value = model_sku
        if sc_prod_type_col and sc_prod_type_col in ref_cells:
            ws_sc.cell(curr_r, sc_prod_type_col).value = ref_cells[sc_prod_type_col]
        if sc_parentage_col:
            ws_sc.cell(curr_r, sc_parentage_col).value = parent_val

        parent_title = ""
        if sc_title_col:
            child_title = ref_cells.get(sc_title_col, '')
            parent_title = _build_parent_title_from_child(child_title, model_sku)
            ws_sc.cell(curr_r, sc_title_col).value = parent_title

        apply_common_fixed_fields(curr_r)

        preview.append({
            'Row': curr_r,
            'Source VC Sheet': model_to_sheet.get(model_sku, ''),
            'SKU': model_sku,
            'Product Type': ref_cells.get(sc_prod_type_col, '') if sc_prod_type_col else '',
            'Listing Action': action_val,
            'Parentage Level': parent_val,
            'Parent SKU': '',
            'Variation Theme': var_theme_val,
            'Product Id Type': '',
            'Product Id': '',
            'Shipping Template': '',
            'Model Name': '',
            'Item Name': parent_title,
        })
        curr_r += 1
        total_written += 1

    def write_child_row(child_sku, model_str, mapped_cells, src_sheet_name=""):
        nonlocal curr_r, total_written
        for sc_c, val in mapped_cells.items():
            ws_sc.cell(curr_r, sc_c).value = val

        ws_sc.cell(curr_r, sc_sku_col).value = child_sku
        if sc_parentage_col:
            ws_sc.cell(curr_r, sc_parentage_col).value = child_val
        if sc_parent_sku_col and model_str:
            ws_sc.cell(curr_r, sc_parent_sku_col).value = model_str
        if sc_model_name_col and model_str:
            ws_sc.cell(curr_r, sc_model_name_col).value = model_str
        if sc_prod_id_type_col:
            ws_sc.cell(curr_r, sc_prod_id_type_col).value = asin_type_val
        if sc_shipping_tpl_col:
            ws_sc.cell(curr_r, sc_shipping_tpl_col).value = shipping_tpl_val

        apply_common_fixed_fields(curr_r)

        preview.append({
            'Row': curr_r,
            'Source VC Sheet': src_sheet_name,
            'SKU': child_sku,
            'Product Type': mapped_cells.get(sc_prod_type_col, '') if sc_prod_type_col else '',
            'Listing Action': action_val,
            'Parentage Level': child_val,
            'Parent SKU': model_str,
            'Variation Theme': var_theme_val,
            'Product Id Type': asin_type_val,
            'Product Id': mapped_cells.get(sc_prod_id_val_col, '') if sc_prod_id_val_col else '',
            'Shipping Template': shipping_tpl_val,
            'Model Name': model_str,
            'Item Name': mapped_cells.get(sc_title_col, '') if sc_title_col else '',
        })
        curr_r += 1
        total_written += 1

    if parent_position == 'top':
        for m_sku in unique_models_ordered:
            write_parent_row(m_sku)
        for c_sku, m_str, m_cells, sh_title in vc_rows_data:
            write_child_row(c_sku, m_str, m_cells, sh_title)
    elif parent_position == 'bottom':
        for c_sku, m_str, m_cells, sh_title in vc_rows_data:
            write_child_row(c_sku, m_str, m_cells, sh_title)
        for m_sku in unique_models_ordered:
            write_parent_row(m_sku)
    else:
        inserted_parents = set()
        for c_sku, m_str, m_cells, sh_title in vc_rows_data:
            if m_str and m_str not in inserted_parents:
                inserted_parents.add(m_str)
                write_parent_row(m_str)
            write_child_row(c_sku, m_str, m_cells, sh_title)

    output = io.BytesIO()
    wb_sc.save(output)
    output.seek(0)
    return output.getvalue(), preview, total_written, sheets_summary

# ======================================================================
#   ✨ AJIO PERFECT LISTING ENGINE (merged from "3_Perfect_Listing.py")
#   Sab functions `pl_` / `PL_` prefix ke sath hain taaki Listing Tool ke
#   purane functions se naam na takraye.
# ======================================================================
# ======================================================================
# 4. COLOR ABBREVIATION, SKU & STYLE CODE ENGINE (100% DYNAMIC)
# ======================================================================
PL_COLOR_TOKENS = {
    'black', 'white', 'blue', 'lightblue', 'darkblue', 'skyblue', 'iceblue', 'navy', 'navyblue',
    'green', 'olive', 'pista', 'sea-green', 'purple', 'lavender', 'pink', 'red', 'maroon',
    'beige', 'brown', 'coffee', 'khaki', 'grey', 'gray', 'charcoal', 'yellow', 'orange',
    'peach', 'teal', 'wine', 'cream', 'rust', 'gold', 'silver'
}

PL_COLOR_ABBR = {
    'black': 'Blk', 'white': 'Wht', 'blue': 'Blu', 'lightblue': 'LBlu',
    'darkblue': 'DBlu', 'skyblue': 'SBlu', 'iceblue': 'IBlu', 'navy': 'Nvy',
    'green': 'Grn', 'olive': 'Olv', 'pista': 'Pst', 'sea-green': 'SGrn',
    'purple': 'Prp', 'lavender': 'Lav', 'pink': 'Pnk', 'red': 'Red',
    'maroon': 'Mrn', 'beige': 'Beg', 'brown': 'Brn', 'coffee': 'Cfe',
    'khaki': 'Khk', 'grey': 'Gry', 'gray': 'Gry', 'charcoal': 'Chr',
    'yellow': 'Ylw', 'orange': 'Orn', 'peach': 'Pch', 'teal': 'Tel',
    'wine': 'Win', 'cream': 'Crm', 'ribbin': 'Rbn', 'teddy': 'Tdy'
}

def pl_clean_sku_name(name):
    cleaned = re.sub(r'[\s_]+', '-', str(name).strip())
    return re.sub(r'-+', '-', cleaned).strip('-')

def pl_extract_style_code(folder_name):
    """
    Folder name se color suffix hata kar natural Style Code auto-extract hota hai.
    Zero hardcoded prefix - Universal for Saree, Kurta, Jeans, Men, Kids, Night suit.
    """
    clean = pl_clean_sku_name(folder_name)
    parts = clean.split('-')
    cut_idx = -1
    for i in range(len(parts) - 1, 0, -1):
        if parts[i].lower() in PL_COLOR_TOKENS:
            cut_idx = i
            break
    if cut_idx > 0:
        return '-'.join(parts[:cut_idx])
    if len(parts) >= 2 and not parts[-1].isdigit():
        return '-'.join(parts[:-1])
    return clean

def pl_generate_safe_ean(clean_style, clean_sz):
    """
    EAN hamesha <= 18 characters rahega aur size suffix (-{clean_sz}) kabhi cut nahi hoga.
    Prefix aur color name ko abbreviate karke fit karta hai.
    """
    sz_suffix = f"-{clean_sz}"
    candidate = f"{clean_style}{sz_suffix}"
    if len(candidate) <= 18:
        return candidate

    parts = clean_style.split('-')
    new_parts = []
    for p in parts:
        low = p.lower()
        if low in PL_COLOR_ABBR:
            new_parts.append(PL_COLOR_ABBR[low])
        elif p.isdigit():
            new_parts.append(str(int(p)))
        else:
            new_parts.append(p)

    candidate2 = "-".join(new_parts) + sz_suffix
    if len(candidate2) <= 18:
        return candidate2

    candidate3 = "".join(new_parts) + sz_suffix
    if len(candidate3) <= 18:
        return candidate3

    prefix_allowed = 18 - len(sz_suffix)
    return candidate3[:prefix_allowed] + sz_suffix

def pl_is_size_chart_file(filename):
    low = os.path.basename(filename).lower()
    patterns = ['size_chart', 'size-chart', 'sizechart', 'size_guide', 'sizeguide', 'size chart', 'chart', 'measurement']
    return any(p in low for p in patterns) and low.endswith(('.jpg', '.jpeg', '.png', '.webp'))

def pl_norm_h(s):
    return re.sub(r'[\s_\-*]+', '', str(s or '')).strip().lower()

PL_CATEGORY_OPTIONS = [
    'co-ord set', 'night suit', 'lounge wear', 'pajama set', 'saree', 'kurta', 'kurti',
    'jeans', 'cargo', 'trousers', 'joggers', 'shorts', 'tshirt', 'shirt', 'jacket',
    'hoodie', 'sweatshirt', 'dress', 'top', 'skirt', 'apparel'
]

PL_COMMON_COLORS = [
    "Beige", "LightGrey", "Grey", "DarkGrey", "Black", "Navy", "Blue", "LightBlue", "DarkBlue",
    "SkyBlue", "IceBlue", "White", "Off White", "Green", "Olive", "Pista", "Sea Green", "Brown",
    "Charcoal", "Khaki", "Red", "Maroon", "Yellow", "Orange", "Pink", "Purple", "Lavender", "Wine", "Peach", "Teal"
]

PL_COLOR_FAMILY_MAP = {
    'maroon': 'Red', 'red': 'Red', 'wine': 'Red', 'darkred': 'Red',
    'navy': 'Navy', 'navyblue': 'Navy', 'blue': 'Blue', 'lightblue': 'Blue',
    'darkblue': 'Blue', 'iceblue': 'Blue', 'skyblue': 'Blue', 'teal': 'Teal',
    'black': 'Black', 'charcoal': 'Grey', 'grey': 'Grey', 'gray': 'Grey',
    'lightgrey': 'Grey', 'darkgrey': 'Grey', 'white': 'White', 'offwhite': 'Off White',
    'green': 'Green', 'olive': 'Olive', 'pista': 'Sea Green', 'sea green': 'Sea Green',
    'brown': 'Brown', 'coffee': 'Brown', 'beige': 'Beige', 'khaki': 'Khaki',
    'yellow': 'Yellow', 'mustard': 'Yellow', 'orange': 'Orange', 'peach': 'Peach',
    'pink': 'Pink', 'purple': 'Purple', 'lavender': 'Purple'
}

def pl_detect_color_from_name(text):
    text_clean = str(text).replace('-', ' ').replace('_', ' ').lower()
    for col in sorted(PL_COMMON_COLORS, key=len, reverse=True):
        if col.lower() in text_clean:
            return col
    return "Pink"

def pl_map_color_to_family(col_name):
    c = str(col_name).strip().lower()
    return PL_COLOR_FAMILY_MAP.get(c, col_name)

# ======================================================================
# 5. DYNAMIC DIMENSIONS, WEIGHT & INTELLIGENT HSN
# ======================================================================
def pl_get_dimensions_and_weight(category):
    """
    Volumetric L*W*H / 5 aur actual weight dono strictly 300g se 450g ke beech rehte hain.
    """
    cat = str(category or '').lower()
    if any(k in cat for k in ['jean', 'denim', 'cargo', 'heavy', 'jacket']):
        l, w, h = 28, 25, 3.0
        art_wt, pkg_wt = 400, 430
    elif any(k in cat for k in ['saree', 'dress', 'kurta', 'kurti']):
        l, w, h = 26, 24, 2.5
        art_wt, pkg_wt = 330, 360
    elif any(k in cat for k in ['tshirt', 't-shirt', 'top', 'shirt', 'kids']):
        l, w, h = 25, 24, 2.5
        art_wt, pkg_wt = 300, 340
    else:
        l, w, h = 25, 24, 2.5
        art_wt, pkg_wt = 350, 380

    return {
        'art_l': l, 'art_w': w, 'art_h': h, 'art_luom': 'CM',
        'art_wt': art_wt, 'art_wtuom': 'GRAM',
        'pkg_l': l, 'pkg_w': w, 'pkg_h': h, 'pkg_luom': 'CM',
        'pkg_wt': pkg_wt, 'pkg_wtuom': 'GRAM'
    }

PL_HSN_INTELLIGENCE_MAP = {
    'saree': '54075200',
    'kurti': '62042200',
    'kurta': '62042200',
    'night suit': '61083100',
    'sleepwear': '61083100',
    'lounge wear': '61083100',
    'co-ord set': '61083100',
    'jeans': '62034290',
    'tshirt': '61099090',
    'shirt': '62052000',
    'trousers': '62034200',
    'cargo': '62034200',
    'joggers': '62034200',
    'shorts': '62034200',
    'dress': '62044200',
    'hoodie': '61103000',
    'sweatshirt': '61103000',
    'jacket': '62014090',
    'top': '62063000',
    'kids': '61112000',
    'apparel': '62042200'
}

def pl_resolve_hsn_code(user_hsn, category):
    if user_hsn and str(user_hsn).strip():
        return str(user_hsn).strip()
    cat = str(category or '').lower()
    for k, code in PL_HSN_INTELLIGENCE_MAP.items():
        if k in cat:
            return code
    return '62042200'

# ======================================================================
# 6. MULTI-AI VISION & SIZE CHART EXTRACTORS
# ======================================================================
def pl_guess_mime(filename):
    low = str(filename or '').lower()
    if low.endswith('.png'): return 'image/png'
    if low.endswith('.webp'): return 'image/webp'
    return 'image/jpeg'

def pl_extract_json_from_text(text):
    if not text: return None
    try:
        m = re.search(r'\{[\s\S]*\}', text)
        if m: return json.loads(m.group(0))
    except Exception:
        pass
    return None

def pl_get_garment_vision_prompt():
    cats = ", ".join(PL_CATEGORY_OPTIONS)
    return (
        "You are an expert e-commerce fashion cataloger. Visually inspect this garment image very carefully and output ONLY a valid JSON object without markdown or code formatting.\n"
        "Return this exact schema:\n"
        "{\n"
        '  "category": "<One of: ' + cats + '>",\n'
        '  "gender": "<Women, Men, Boys, Girls, or Infants>",\n'
        '  "color": "<Dominant color: Pink, Maroon, Black, Navy, Blue, LightBlue, SkyBlue, Grey, Beige, Olive, Pista, Sea Green, Lavender, White, Red, Yellow, etc.>",\n'
        '  "pocket": "<Kangaroo Pocket / Cargo Pocket / Five Pocket / Side Pocket / Welt Pocket / Chest Pocket / Slash Pocket / No Pocket>",\n'
        '  "material": "<Actual fabric: Cotton, Cotton Blend, Rayon, Fleece, Polyester, Denim, Crepe, Georgette, Silk, Knit, Velvet, Chiffon>",\n'
        '  "fabric_detail": "<Specific fabric description e.g. 100% Cotton, Pure Cotton, Cotton Blend, Soft Fleece, Rayon, Silk Blend>",\n'
        '  "sleeve_length": "<Short sleeve / Full-length sleeve / 3/4th sleeve / Sleeveless>",\n'
        '  "topwear_type": "<Top / Tunic / Shirt / Hoodie / Sweatshirt / Kurta / Kurti>",\n'
        '  "bottomwear_type": "<Capris / Pyjamas / Shorts / Trousers/Pants / Jeans / Skirts>",\n'
        '  "waist_type": "<Flexi waist / Elasticated waist / Fixed waist>",\n'
        '  "pattern": "<Solid / Printed / Washed / Striped / Checked / Floral / Quilted / Self-design>"\n'
        "}\n"
    )

def pl_get_size_chart_prompt():
    return (
        "You are an expert OCR & e-commerce technical fashion cataloger. Carefully read all text and table numbers in this size chart image.\n"
        "Extract all measurements for XS, S, M, L, XL, XXL (2XL), 3XL, 28, 30, 32, 34, 36, 38.\n"
        "Identify if bottomwear is Capris, Pyjamas, Pants, or Shorts, and if sleeve is Short sleeve or Full-length sleeve.\n"
        "Output ONLY a valid JSON object without markdown or code formatting with this schema:\n"
        "{\n"
        '  "bottomwear_type": "<Capris or Pyjamas or Shorts or Trousers/Pants>",\n'
        '  "sleeve_length": "<Short sleeve or Full-length sleeve>",\n'
        '  "measurements": {\n'
        '    "XS": {"bust": 37, "front_length": 25, "waist": 22, "hip": 38, "inseam": 28},\n'
        '    "S": {"bust": 39, "front_length": 25, "waist": 24, "hip": 40, "inseam": 28},\n'
        '    "M": {"bust": 41, "front_length": 26, "waist": 26, "hip": 42, "inseam": 28},\n'
        '    "L": {"bust": 43, "front_length": 26, "waist": 28, "hip": 45, "inseam": 29},\n'
        '    "XL": {"bust": 45, "front_length": 27, "waist": 30, "hip": 48, "inseam": 29},\n'
        '    "XXL": {"bust": 47, "front_length": 27, "waist": 32, "hip": 51, "inseam": 29}\n'
        '  }\n'
        "}\n"
        "Return numeric values for measurements in inches. If 2XL is written in table, map it to XXL."
    )

def pl_call_gemini_api(prompt, image_bytes, img_name):
    if not GEMINI_API_KEY or not image_bytes: return None
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    payload = {
        "contents": [{"parts": [
            {"text": prompt},
            {"inline_data": {"mime_type": pl_guess_mime(img_name), "data": b64}}
        ]}],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 1500},
    }
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_VISION_MODEL}:generateContent?key={GEMINI_API_KEY}"
        res = requests.post(url, json=payload, timeout=40)
        if res.status_code == 200:
            cands = res.json().get('candidates') or []
            if cands:
                parts = (cands[0].get('content') or {}).get('parts') or []
                text = ''.join(str(p.get('text', '')) for p in parts)
                return pl_extract_json_from_text(text)
    except Exception:
        pass
    return None

def pl_call_nvidia_api(prompt, image_bytes, img_name):
    if not NVIDIA_API_KEY or not image_bytes: return None
    mime = pl_guess_mime(img_name)
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    data_url = f"data:{mime};base64,{b64}"
    headers = {"Authorization": f"Bearer {NVIDIA_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": NVIDIA_MODEL,
        "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": data_url}}]}],
        "temperature": 0.1, "max_tokens": 1500
    }
    try:
        url = "https://integrate.api.nvidia.com/v1/chat/completions"
        res = requests.post(url, headers=headers, json=payload, timeout=50)
        if res.status_code == 200:
            return pl_extract_json_from_text(res.json()['choices'][0]['message']['content'])
    except Exception:
        pass
    return None

def pl_call_groq_api(prompt, image_bytes, img_name):
    if not GROQ_API_KEY or not image_bytes: return None
    mime = pl_guess_mime(img_name)
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    data_url = f"data:{mime};base64,{b64}"
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": GROQ_VISION_MODEL,
        "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": data_url}}]}],
        "temperature": 0.1, "max_tokens": 1500
    }
    try:
        url = "https://api.groq.com/openai/v1/chat/completions"
        res = requests.post(url, headers=headers, json=payload, timeout=40)
        if res.status_code == 200:
            return pl_extract_json_from_text(res.json()['choices'][0]['message']['content'])
    except Exception:
        pass
    return None

def pl_fallback_ocr_size_chart(image_bytes):
    try:
        import pytesseract
        img = Image.open(io.BytesIO(image_bytes))
        txt = pytesseract.image_to_string(img)
        lines = [l.strip() for l in txt.split('\n') if l.strip()]
        is_capri = any('capri' in l.lower() for l in lines)
        is_full_sleeve = any('full' in l.lower() for l in lines)
        bottom_type = 'Capris' if is_capri else 'Pyjamas'
        sleeve_len = 'Full-length sleeve' if is_full_sleeve else 'Short sleeve'
    except Exception:
        bottom_type = 'Pyjamas'
        sleeve_len = 'Short sleeve'

    measurements = {
        'XS': {'bust': 37, 'front_length': 25, 'waist': 22, 'hip': 38, 'inseam': 28},
        'S': {'bust': 39, 'front_length': 25, 'waist': 24, 'hip': 40, 'inseam': 28},
        'M': {'bust': 41, 'front_length': 26, 'waist': 26, 'hip': 42, 'inseam': 28},
        'L': {'bust': 43, 'front_length': 26, 'waist': 28, 'hip': 45, 'inseam': 29},
        'XL': {'bust': 45, 'front_length': 27, 'waist': 30, 'hip': 48, 'inseam': 29},
        'XXL': {'bust': 47, 'front_length': 27, 'waist': 32, 'hip': 51, 'inseam': 29},
        '2XL': {'bust': 47, 'front_length': 27, 'waist': 32, 'hip': 51, 'inseam': 29},
    }
    return {'bottomwear_type': bottom_type, 'sleeve_length': sleeve_len, 'measurements': measurements}

def pl_multi_ai_vision_detect(image_bytes, img_name):
    prompt = pl_get_garment_vision_prompt()
    res = pl_call_gemini_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('color'): return res, "Gemini"
    res = pl_call_nvidia_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('color'): return res, "NVIDIA"
    res = pl_call_groq_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('color'): return res, "Groq"
    return None, "None"

def pl_multi_ai_size_chart_detect(image_bytes, img_name):
    prompt = pl_get_size_chart_prompt()
    res = pl_call_gemini_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('measurements'): return res, "Gemini"
    res = pl_call_nvidia_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('measurements'): return res, "NVIDIA"
    res = pl_call_groq_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('measurements'): return res, "Groq"
    return pl_fallback_ocr_size_chart(image_bytes), "OCR-Fallback"

# ======================================================================
# 7. TEMPLATE & INSTRUCTIONS ANALYZER
# ======================================================================
def pl_read_and_analyze_template(wb):
    info = {
        'sheets': wb.sheetnames,
        'styles_sheet': None,
        'category_breadcrumb': '',
        'detected_gender': 'Women',
        'detected_category': 'Apparel',
        'col_to_dropdowns': {},
        'headers': {},
        'mandatory_cols': []
    }
    
    if 'Seller Information Sheet' in wb.sheetnames:
        target = wb['Seller Information Sheet'].cell(2, 1).value
        if target and str(target).strip() in wb.sheetnames:
            info['styles_sheet'] = str(target).strip()
            
    if not info['styles_sheet']:
        for s in wb.sheetnames:
            if 'styles' in s.lower() and 'instruction' not in s.lower():
                info['styles_sheet'] = s
                break
    if not info['styles_sheet']:
        info['styles_sheet'] = wb.sheetnames[0]

    ws_styles = wb[info['styles_sheet']]
    breadcrumb_val = ws_styles.cell(1, 1).value
    if breadcrumb_val:
        info['category_breadcrumb'] = str(breadcrumb_val).strip()
        parts = [p.strip() for p in info['category_breadcrumb'].split('->')]
        if len(parts) > 0: info['detected_gender'] = parts[0]
        if len(parts) > 1: info['detected_category'] = parts[-1]

    header_row = 3
    keywords = ['style', 'sku', 'mrp', 'price', 'size', 'image', 'brand', 'description']
    for r in range(1, min(7, ws_styles.max_row + 1)):
        vals = [str(ws_styles.cell(r, c).value or '').lower() for c in range(1, ws_styles.max_column + 1)]
        if sum(1 for v in vals if any(k in v for k in keywords)) > 3:
            header_row = r
            break
            
    info['header_row'] = header_row
    info['headers'] = {c: str(ws_styles.cell(header_row, c).value or '').strip() for c in range(1, ws_styles.max_column + 1)}
    info['mandatory_cols'] = [c for c, h in info['headers'].items() if h.startswith('*')]

    defined_ranges = {}
    if hasattr(wb, 'defined_names'):
        for name, dn in wb.defined_names.items():
            ref = getattr(dn, 'attr_text', '') or getattr(dn, 'value', '') or ''
            defined_ranges[name.lower()] = ref

    col_to_dropdowns = {}
    try:
        for dv in ws_styles.data_validations.dataValidation:
            if dv.type == 'list' and dv.formula1:
                f1 = str(dv.formula1).strip().replace('"', '')
                vals = []
                if ',' in f1:
                    vals = [x.strip() for x in f1.split(',') if x.strip()]
                elif f1.lower() in defined_ranges:
                    ref = defined_ranges[f1.lower()]
                    if '!' in ref:
                        s_name, cell_range = ref.split('!')
                        s_name = s_name.strip("'")
                        if s_name in wb.sheetnames:
                            for row in wb[s_name][cell_range.replace('$', '')]:
                                if not isinstance(row, tuple): row = (row,)
                                for cell in row:
                                    if cell.value is not None and str(cell.value).strip():
                                        vals.append(str(cell.value).strip())
                for rng in dv.sqref.ranges:
                    for col_idx in range(rng.min_col, rng.max_col + 1):
                        col_to_dropdowns[col_idx] = vals
    except Exception:
        pass

    info['col_to_dropdowns'] = col_to_dropdowns
    return info

def pl_pick_dropdown_value_safe(info, col_idx, desired, fallback=None):
    if not col_idx: return desired or fallback
    options = info.get('col_to_dropdowns', {}).get(col_idx)
    if not options: return desired if desired is not None else fallback
    d = str(desired if desired is not None else '').strip()
    fl = str(fallback if fallback is not None else '').strip()
    cand = d or fl
    if not cand: return None
    cl = cand.lower()
    for o in options:
        if str(o).strip().lower() == cl: return o
    for o in options:
        ol = str(o).strip().lower()
        if cl in ol or ol in cl: return o
    close = difflib.get_close_matches(cand, options, n=1, cutoff=0.4)
    return close[0] if close else options[0]

# ======================================================================
# 8. CDN UPLOADER & DYNAMIC SEO CONTENT
# ======================================================================
def pl_upload_to_imgbb(image_bytes, api_key):
    try:
        url = "https://api.imgbb.com/1/upload"
        payload = {"key": api_key, "image": base64.b64encode(image_bytes).decode('utf-8')}
        res = requests.post(url, data=payload, timeout=20)
        if res.status_code == 200:
            return res.json()['data']['url']
    except Exception:
        pass
    return None

def pl_upload_to_imgbb_cached(image_bytes, api_key, cache):
    h = hashlib.md5(image_bytes).hexdigest()
    if h in cache: return cache[h]
    url = pl_upload_to_imgbb(image_bytes, api_key)
    if url: cache[h] = url
    return url

def pl_build_dynamic_seo_content(category, gender, brand, fabric, pattern, keywords=""):
    g_title = "Women's" if str(gender).lower().startswith('w') else ("Men's" if str(gender).lower().startswith('m') else "Unisex")
    b_prefix = f"{brand} " if brand else ""
    cat_label = category.title() if category else "Apparel"
    kw_str = keywords if keywords else "comfortable relaxed fit, premium fabric, all-day comfort, durable quality"
    
    desc = (
        f"{b_prefix}{g_title} {cat_label} - Designed for modern style and maximum comfort, "
        f"this {g_title.lower()} {cat_label.lower()} is tailored from high-grade breathable {fabric} fabric. "
        f"Engineered with {kw_str}. "
        f"Featuring an elegant {pattern.lower()} pattern with refined seam construction. "
        f"Ideal for daily wear, casual outings, work from home, and relaxed styling."
    )
    highlight = f"Premium breathable {fabric} fabric | Comfortable {g_title.lower()} fit | Elegant {pattern.lower()} design | Durable stitching | Easy machine wash"
    add_info = "Wash Care: Machine wash cold with similar colors. Do not bleach. Warm iron if needed."
    return desc[:950], highlight[:500], add_info[:500]

# ======================================================================
# 9. UNIVERSAL AJIO LISTING ENGINE (100% PREFIX-FREE)
# ======================================================================
def pl_run_perfect_ajio_listing(template_file_or_bytes, zip_images_data, detected_size_charts, ctx):
    name = getattr(template_file_or_bytes, 'name', 'template.xlsx')
    wb = openpyxl.load_workbook(template_file_or_bytes, keep_vba=name.lower().endswith('.xlsm'))
    info = pl_read_and_analyze_template(wb)
    ws = wb[info['styles_sheet']]
    headers = info['headers']
    header_row = info['header_row']

    def find_col(candidates):
        for cand in candidates:
            c_norm = pl_norm_h(cand)
            for c, h in headers.items():
                if pl_norm_h(h) == c_norm: return c
        return None

    cat_detected = info['detected_category']
    gender_detected = info['detected_gender']

    col_style = find_col(['*Style Code', 'Style Code'])
    col_desc = find_col(['*Style Description', 'Style Description'])
    col_sku = find_col(['*Item SKU', 'Item SKU', 'SKU'])
    col_brand = find_col(['*Brand', 'Brand'])
    col_ean = find_col(['*EAN', 'EAN', 'Barcode'])
    col_td = find_col(['*TD', 'TD'])
    col_mrp = find_col(['*MRP', 'MRP', 'Maximum Retail Price'])
    col_hsn = find_col(['*HSN', 'HSN', 'HSN Code'])
    col_prod_grp = find_col(['*Product Groups', 'Product Groups'])
    col_fash_grp = find_col(['*Fashion Groups', 'Fashion Groups'])
    col_season_cd = find_col(['*Season Code', 'Season Code'])
    col_season_yr = find_col(['*Season Year', 'Season Year'])
    col_size = find_col(['*Size', 'Size'])
    col_std_size = find_col(['*StandardSize', 'StandardSize', 'Standard Size'])

    col_art_h = find_col(['*articleDimensionsUnitHeight'])
    col_art_l = find_col(['*articleDimensionsUnitLength'])
    col_art_w = find_col(['*articleDimensionsUnitWidth'])
    col_art_luom = find_col(['*articleDimensionsUnitLengthUOM'])
    col_art_wt = find_col(['*articleDimensionsUnitWeight'])
    col_art_wtuom = find_col(['*articleDimensionsUnitWeightUOM'])
    col_pkg_h = find_col(['*packageDimensionsHeight'])
    col_pkg_l = find_col(['*packageDimensionsLength'])
    col_pkg_w = find_col(['*packageDimensionsWidth'])
    col_pkg_luom = find_col(['*packageDimensionsLengthUOM'])
    col_pkg_wt = find_col(['*packageDimensionsWeight'])
    col_pkg_wtuom = find_col(['*packageDimensionsWeightUOM'])

    col_add_info1 = find_col(['Additional Information 1'])
    col_comp_count = find_col(['*Component Count'])
    col_country = find_col(['*Country of Origin'])
    col_highlight = find_col(['Highlight'])
    col_mfg_by = find_col(['Manufactured By'])
    col_mkt_by = find_col(['*Marketed By'])
    col_net_qty = find_col(['*Net Quantity'])
    col_pkg_contains = find_col(['*Package Contains'])

    col_color_fam = find_col(['*Color Family'])
    col_color_shade = find_col(['*Color Shade'])
    col_fab_detail = find_col(['Fabric Detail'])
    col_fab_type = find_col(['*Fabric Type'])
    col_pattern = find_col(['*Pattern'])
    col_pri_color = find_col(['*Primary Color'])
    col_size_format = find_col(['*Size Format'])
    col_size_group = find_col(['*Size Group'])
    col_wash_care = find_col(['*Wash Care'])
    col_plus_size = find_col(['*Plus Size Store'])
    col_prod_name = find_col(['*Product Name'])

    col_sleeve_len = find_col(['Sleeve Length'])
    col_waist_type = find_col(['Waist Type'])
    col_top_type = find_col(['*Topwear Type'])
    col_bottom_type = find_col(['Bottomwear Type'])
    col_style_type = find_col(['*Style Type'])

    col_meas_type = find_col(['*Measurement Type'])
    col_to_fit_bust = find_col(['*To Fit Bust'])
    col_bust = find_col(['*Bust'])
    col_front_len = find_col(['*Front Length'])
    col_to_fit_waist = find_col(['*To Fit Waist'])
    col_waist = find_col(['*Waist'])
    col_to_fit_hip = find_col(['*To Fit Hip'])
    col_hip = find_col(['*Hips'])
    col_inseam = find_col(['*Inseam'])

    img_cols = []
    for i in range(1, 15):
        c_i = find_col([f'*MODEL' if i == 1 else f'MODEL{i}'])
        if c_i and c_i not in img_cols: img_cols.append(c_i)

    # Dynamic Dimensions & Weights strictly between 300g and 450g
    dims_data = pl_get_dimensions_and_weight(cat_detected)
    hsn_val = pl_resolve_hsn_code(ctx.get('hsn'), cat_detected)

    curr_r = header_row + 1
    while ws.cell(curr_r, col_style or 1).value is not None:
        curr_r += 1

    total_added = 0
    preview = []
    sizes = ctx.get('sizes') or ['S', 'M', 'L', 'XL', '2XL']
    season_year = str(datetime.now().year + 1)
    
    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)

    def set_val(r, c, val, dropdown=False):
        if not c or val in (None, ''): return
        ws.cell(r, c).value = pl_pick_dropdown_value_safe(info, c, val, val) if dropdown else val

    # Fill data for each product folder in zip
    for fld, img_tuples in zip_images_data.items():
        # EXACT BASE SKU FROM ZIP FOLDER (NO HARDCODED PREFIX)
        base_sku = pl_clean_sku_name(fld)
        # NATURAL STYLE CODE EXTRACTED BY REMOVING COLOR SUFFIX
        style_code = pl_extract_style_code(fld)
        
        ai_data = ctx.get('vision_results', {}).get(base_sku, {})
        det_col = ai_data.get('color') or pl_detect_color_from_name(fld)
        col_fam = pl_map_color_to_family(det_col)
        mat_val = ai_data.get('material') or "Cotton"
        fab_detail = ai_data.get('fabric_detail') or f"100% {mat_val}"
        pattern_val = ai_data.get('pattern') or ("Quilted" if "quilt" in fld.lower() else "Solid")
        
        sc_data = detected_size_charts.get(fld) or detected_size_charts.get('default') or {}
        sleeve_val = sc_data.get('sleeve_length') or ai_data.get('sleeve_length') or "Short sleeve"
        bottom_val = sc_data.get('bottomwear_type') or ("Capris" if "capri" in fld.lower() else "Pyjamas")
        measurements = sc_data.get('measurements') or {}

        desc_text, high_text, add_text = pl_build_dynamic_seo_content(
            cat_detected, gender_detected, ctx.get('brand', ''), mat_val, pattern_val, ctx.get('keywords', '')
        )
        
        img_urls = list(ctx.get('live_links_cache', {}).get(fld, []))
        if ctx.get('size_chart_urls', {}).get(fld):
            chart_url = ctx['size_chart_urls'][fld]
            if chart_url not in img_urls: img_urls.append(chart_url)

        preview.append({
            'Style Code': style_code,
            'Base SKU': base_sku,
            'Category': cat_detected,
            'Color': det_col,
            'Color Family': col_fam,
            'Fabric': mat_val,
            'Pattern': pattern_val,
            'Images Count': len(img_urls)
        })

        for sz in sizes:
            clean_sz = pl_clean_sku_name(sz)
            # ITEM SKU = ZIP FOLDER NAME + '-' + SIZE
            full_sku = f"{base_sku}-{clean_sz}"
            # EAN = SAFE <= 18 CHARACTERS (SIZE SUFFIX ALWAYS PRESERVED)
            ean_val = pl_generate_safe_ean(base_sku, clean_sz)
            
            sz_meas = measurements.get(clean_sz) or measurements.get(clean_sz.upper()) or {}
            if not sz_meas:
                _alias = {'2XL': 'XXL', 'XXL': '2XL'}.get(clean_sz.upper())
                sz_meas = measurements.get(_alias) or {}
            bust_val = sz_meas.get('bust', 37)
            front_len_val = sz_meas.get('front_length', 25)
            waist_val = sz_meas.get('waist', 22)
            to_fit_waist_val = waist_val + 2
            hip_val = sz_meas.get('hip', 38)
            inseam_val = sz_meas.get('inseam', 28)

            set_val(curr_r, col_style, style_code)
            set_val(curr_r, col_desc, desc_text)
            set_val(curr_r, col_sku, full_sku)
            set_val(curr_r, col_brand, ctx.get('brand', ''))
            set_val(curr_r, col_ean, ean_val)
            set_val(curr_r, col_td, ctx.get('td', 50))
            
            if col_mrp:
                if ctx.get('mrp'):
                    ws.cell(curr_r, col_mrp).value = int(ctx['mrp']) if str(ctx['mrp']).isdigit() else ctx['mrp']
                else:
                    ws.cell(curr_r, col_mrp).fill = mrp_fill
                    ws.cell(curr_r, col_mrp).font = mrp_font

            set_val(curr_r, col_hsn, hsn_val)
            set_val(curr_r, col_prod_grp, "Evening", dropdown=True)
            set_val(curr_r, col_fash_grp, "Fashion", dropdown=True)
            set_val(curr_r, col_season_cd, "Core", dropdown=True)
            set_val(curr_r, col_season_yr, season_year)
            set_val(curr_r, col_size, clean_sz)
            set_val(curr_r, col_std_size, clean_sz)

            # Dimensions strictly 300g to 450g
            set_val(curr_r, col_art_h, dims_data['art_h']); set_val(curr_r, col_art_l, dims_data['art_l'])
            set_val(curr_r, col_art_w, dims_data['art_w']); set_val(curr_r, col_art_luom, dims_data['art_luom'])
            set_val(curr_r, col_art_wt, dims_data['art_wt']); set_val(curr_r, col_art_wtuom, dims_data['art_wtuom'])
            set_val(curr_r, col_pkg_h, dims_data['pkg_h']); set_val(curr_r, col_pkg_l, dims_data['pkg_l'])
            set_val(curr_r, col_pkg_w, dims_data['pkg_w']); set_val(curr_r, col_pkg_luom, dims_data['pkg_luom'])
            set_val(curr_r, col_pkg_wt, dims_data['pkg_wt']); set_val(curr_r, col_pkg_wtuom, dims_data['pkg_wtuom'])

            set_val(curr_r, col_add_info1, add_text)
            set_val(curr_r, col_comp_count, 1)
            set_val(curr_r, col_country, "India", dropdown=True)
            set_val(curr_r, col_highlight, high_text)
            set_val(curr_r, col_mfg_by, ctx.get('mfg', ''))
            set_val(curr_r, col_mkt_by, ctx.get('mfg', ''))
            set_val(curr_r, col_net_qty, 1)
            set_val(curr_r, col_pkg_contains, "1")

            set_val(curr_r, col_color_fam, col_fam, dropdown=True)
            set_val(curr_r, col_color_shade, "Light", dropdown=True)
            set_val(curr_r, col_fab_detail, fab_detail)
            set_val(curr_r, col_fab_type, mat_val, dropdown=True)
            set_val(curr_r, col_pattern, pattern_val, dropdown=True)
            set_val(curr_r, col_pri_color, det_col, dropdown=True)
            set_val(curr_r, col_size_format, "UNI", dropdown=True)
            set_val(curr_r, col_size_group, "Regular", dropdown=True)
            set_val(curr_r, col_wash_care, "Machine wash", dropdown=True)
            set_val(curr_r, col_plus_size, "No", dropdown=True)
            set_val(curr_r, col_prod_name, "Knitted Top", dropdown=True)

            set_val(curr_r, col_sleeve_len, sleeve_val, dropdown=True)
            set_val(curr_r, col_waist_type, "Flexi waist", dropdown=True)
            set_val(curr_r, col_top_type, "Top", dropdown=True)
            set_val(curr_r, col_bottom_type, bottom_val, dropdown=True)
            set_val(curr_r, col_style_type, "2-Piece Sets", dropdown=True)

            if col_meas_type:
                set_val(curr_r, col_meas_type, "Garment Measurement", dropdown=True)
                set_val(curr_r, col_to_fit_bust, bust_val)
                set_val(curr_r, col_bust, bust_val)
                set_val(curr_r, col_front_len, front_len_val)
                set_val(curr_r, col_to_fit_waist, to_fit_waist_val)
                set_val(curr_r, col_waist, waist_val)
                set_val(curr_r, col_to_fit_hip, hip_val)
                set_val(curr_r, col_hip, hip_val)
                set_val(curr_r, col_inseam, inseam_val)

            for idx_img, url in enumerate(img_urls):
                if idx_img < len(img_cols):
                    ws.cell(curr_r, img_cols[idx_img]).value = url

            curr_r += 1
            total_added += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), preview, total_added, info['styles_sheet']


# ======================================================================
#   ✨ UNIFIED DASHBOARD (Ajio jaisa dashboard - sabhi marketplaces ke liye)
# ======================================================================
def _mime_for(name):
    return ("application/vnd.ms-excel.sheet.macroEnabled.12"
            if str(name).lower().endswith('.xlsm')
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_listing_ui():
    is_ajio = (marketplace == "Ajio")
    if 'pl_vision_results' not in st.session_state:
        st.session_state['pl_vision_results'] = {}

    st.success(f"✨ **{marketplace} Listing Engine Active** — Multi-AI Vision (Gemini ➜ NVIDIA ➜ Groq) | "
               "Dynamic Size Chart Reader | Template Analyzer | 100% Dropdown Safety")

    pl_folders = []
    pl_zip_images = {}
    pl_size_charts = {}

    # ---------------- 1. ZIP UPLOAD ----------------
    st.subheader("1. Product Images, SKUs & Size Chart (ZIP Upload)")
    col_in1, col_in2 = st.columns(2)
    with col_in1:
        st.write("##### 📁 Product Folders Ka ZIP Upload")
        zip_file = st.file_uploader("Product Folders (.zip):", type=["zip"], key="pl_zip_in")
        if zip_file:
            try:
                with zipfile.ZipFile(zip_file, 'r') as z:
                    file_list = z.namelist()
                    folder_to_imgs = {}
                    for f in file_list:
                        clean = f.replace('\\', '/').strip('/')
                        parts = clean.split('/')
                        if len(parts) >= 2:
                            fld = parts[1] if parts[0].upper() == 'NIGHTSUIT' else parts[0]
                            filename = parts[-1]
                            if pl_is_size_chart_file(filename):
                                pl_size_charts[fld] = {'filename': filename, 'bytes': z.read(f)}
                            elif filename.lower().endswith(('.jpg', '.jpeg', '.png', '.webp')):
                                folder_to_imgs.setdefault(fld, []).append((filename, z.read(f)))

                    pl_folders = sorted(list(folder_to_imgs.keys()))
                    for fld in pl_folders:
                        imgs = folder_to_imgs[fld]
                        imgs.sort(key=lambda item: int(re.findall(r'\d+', item[0])[0]) if re.findall(r'\d+', item[0]) else 999)
                        pl_zip_images[fld] = imgs

                st.success(f"✅ ZIP Se Total **{len(pl_folders)}** SKUs Detect Hue:")
                st.code(", ".join([pl_clean_sku_name(f) for f in pl_folders]))
                if pl_size_charts:
                    st.info(f"📐 **{len(pl_size_charts)} Size Charts Detect Hue:** "
                            f"{', '.join(set(v['filename'] for v in pl_size_charts.values()))}")
            except Exception as e:
                st.error(f"ZIP error: {e}")

    with col_in2:
        st.write("##### 🌐 Live CDN & Cloud Hosting")
        auto_cdn = st.checkbox("Images ko auto CDN links (ImgBB) me convert karein", value=True, key="pl_auto_cdn")
        st.caption("⚡ Direct CDN Link Active: ImgBB se direct public URLs auto-generate honge.")

    dropbox_url, manual_skus = "", ""
    if not is_ajio:
        with st.expander("📎 Optional: Dropbox Link / Manual SKUs (ZIP na ho to)"):
            dropbox_url = st.text_input("Dropbox Shared Folder Link (Optional):", key="pl_dropbox",
                                        placeholder="https://www.dropbox.com/scl/fo/xyz/...")
            manual_skus = st.text_area("Manual SKUs List (Agar ZIP nahi hai):", key="pl_manual_skus",
                                       placeholder="SKU-1\nSKU-2\nSKU-4", height=90)

    st.divider()

    # ---------------- 2. TEMPLATE UPLOAD ----------------
    st.subheader("2. Marketplace Template Upload")
    tpl_file = st.file_uploader(f"{marketplace} Template (.xlsx / .xlsm):", type=["xlsx", "xlsm"], key="pl_tpl_in")

    if tpl_file and is_ajio:
        try:
            wb_check = openpyxl.load_workbook(tpl_file, data_only=True)
            t_info = pl_read_and_analyze_template(wb_check)
            c_t1, c_t2, c_t3 = st.columns(3)
            with c_t1: st.metric("Target Styles Sheet", t_info['styles_sheet'])
            with c_t2: st.metric("Template Columns", len(t_info['headers']))
            with c_t3: st.metric("Mandatory Columns (*)", len(t_info['mandatory_cols']))
            if t_info['category_breadcrumb']:
                st.info(f"📌 **Category Breadcrumb:** `{t_info['category_breadcrumb']}`")
        except Exception as e:
            st.warning(f"Template parsing warning: {e}")

    st.divider()

    # ---------------- 3. BRAND, MRP, LEGAL & SIZES ----------------
    st.subheader("3. Brand, MRP, Legal & Sizes")
    c_b1, c_b2 = st.columns([1, 2])
    with c_b1:
        size_preset = st.radio(
            "Sizes Preset Chunein (Choose Size Format):",
            ["Alpha Sizes (S, M, L, XL, 2XL)", "Numeric / Waist Sizes (28, 30, 32, 34, 36, 38)"],
            index=0, key="pl_size_preset"
        )
        preset_val = "S, M, L, XL, 2XL" if size_preset.startswith("Alpha") else "28, 30, 32, 34, 36, 38"
        sizes_input = st.text_input(
            "Sizes (comma-separated): Apne mutabiq add ya kam karein",
            value=preset_val, key=f"pl_sizes_{size_preset}"
        )
        brand_name = st.text_input("Brand Name:", value="", key="pl_brand",
                                   placeholder="e.g. PUMA (Khali chhodne par unbranded rahega)")
        mrp_value = st.text_input("MRP (optional):", value="", key="pl_mrp",
                                  placeholder="e.g. 1999 (Khali chhodne par Yellow Highlight hoga)")
        hsn_code = st.text_input("HSN Code (optional):", value="", key="pl_hsn",
                                 placeholder="e.g. 61083100 (Khali chhodne par category ke mutabiq auto-fill hoga)")

    with c_b2:
        mfg_address = st.text_area(
            "Manufactured By & Marketed By (Complete Address with Pincode):", value="", key="pl_mfg",
            placeholder="e.g. Surat, Gujarat\n(Pura address pincode ke sath daalein)", height=130
        )
        custom_keywords = st.text_area(
            "SEO & Trending Keywords (Optional):", value="", key="pl_kw",
            placeholder="e.g. casual streetwear, relaxed fit, durable quality, all-day comfort", height=70
        )

    sc_quantity, sc_selling_price, sc_parent_model = 0, "", ""
    if marketplace == "Amazon - Seller Central":
        with st.expander("🛒 Seller Central Extra Fields (Optional)", expanded=False):
            c_sc1, c_sc2, c_sc3 = st.columns(3)
            with c_sc1:
                sc_quantity = st.number_input("Default Stock Quantity (0 Fixed):", min_value=0, value=0, key="pl_sc_qty")
            with c_sc2:
                sc_selling_price = st.text_input("Selling Price (optional):", key="pl_sc_price", placeholder="e.g. 899")
            with c_sc3:
                sc_parent_model = st.text_input("Common Parent SKU / Model Name (Optional):", key="pl_sc_parent",
                                                placeholder="e.g. H-63-66 (Blank = Model Name from VC / folder)")

    st.divider()

    # ---------------- GENERATE ----------------
    if st.button("🚀 Perfect Listing File Generate Karein", type="primary", key="pl_generate"):
        manual_list = [s.strip() for s in manual_skus.split('\n') if s.strip()] if not is_ajio else []
        if not tpl_file:
            st.error("Kripya Marketplace Template Excel file upload karein!")
        elif not pl_folders and not manual_list:
            st.error("Kripya Product ZIP file upload karein!")
        else:
            sizes_list = [s.strip() for s in sizes_input.split(',') if s.strip()]
            try:
                if is_ajio:
                    # ============ AJIO: Perfect Listing engine ============
                    ctx = {
                        'sizes': sizes_list, 'brand': brand_name.strip(), 'mrp': mrp_value.strip(),
                        'mfg': mfg_address.strip(), 'hsn': hsn_code.strip(), 'keywords': custom_keywords.strip(),
                        'td': int(default_td), 'live_links_cache': {}, 'size_chart_urls': {},
                        'vision_results': st.session_state['pl_vision_results']
                    }
                    with st.spinner("🤖 1/3: AI Vision: Colors, Fabric & Sleeve length detect ho rahe hain..."):
                        for fld in pl_folders:
                            base_sku = pl_clean_sku_name(fld)
                            if base_sku not in ctx['vision_results'] and pl_zip_images.get(fld):
                                v_res, provider = pl_multi_ai_vision_detect(pl_zip_images[fld][0][1], pl_zip_images[fld][0][0])
                                if v_res:
                                    v_res['ai_source'] = provider
                                    ctx['vision_results'][base_sku] = v_res

                    size_chart_measurements = {}
                    with st.spinner("📐 2/3: AI Vision size charts ko read karke measurements extract kar raha hai..."):
                        for fld, sc_info in pl_size_charts.items():
                            sc_res, _ = pl_multi_ai_size_chart_detect(sc_info['bytes'], sc_info['filename'])
                            if sc_res:
                                size_chart_measurements[fld] = sc_res
                                if 'default' not in size_chart_measurements:
                                    size_chart_measurements['default'] = sc_res

                    with st.spinner("🌐 3/3: Images ko Live CDN par upload kiya ja raha hai..."):
                        for fld, sc_info in pl_size_charts.items():
                            url = pl_upload_to_imgbb_cached(sc_info['bytes'], IMGBB_PERMANENT_KEY, img_cache)
                            if url: ctx['size_chart_urls'][fld] = url
                        for fld, img_tuples in pl_zip_images.items():
                            ctx['live_links_cache'][fld] = []
                            for iname, ibytes in img_tuples:
                                url = pl_upload_to_imgbb_cached(ibytes, IMGBB_PERMANENT_KEY, img_cache)
                                if url: ctx['live_links_cache'][fld].append(url)

                    with st.spinner("📝 Template sheet me perfect listing data populate ho raha hai..."):
                        out_bytes, preview_data, total_added, target_sheet = pl_run_perfect_ajio_listing(
                            tpl_file, pl_zip_images, size_chart_measurements, ctx
                        )
                    st.success(f"🎉 Listing complete! Total **{total_added}** variants add hue (`{target_sheet}`).")
                    if preview_data:
                        st.dataframe(preview_data, use_container_width=True)
                    base_name, ext = os.path.splitext(tpl_file.name)
                    out_ext = ext if ext.lower() in ['.xlsm', '.xlsx'] else '.xlsx'
                    out_name = f"{base_name}_PERFECT_LISTED{out_ext}"
                    st.download_button("📥 Download Filled Marketplace File", data=out_bytes,
                                       file_name=out_name, mime=_mime_for(out_name), type="primary")

                else:
                    # ============ BAAKI MARKETPLACES: unke apne engines, Ajio jaisa dashboard ============
                    folders = pl_folders if pl_folders else manual_list
                    valid_keys = {clean_sku_name(f) for f in folders}
                    vr = {k: v for k, v in st.session_state.get('vision_results', {}).items() if k in valid_keys}
                    missing = {f: pl_zip_images[f] for f in pl_folders
                               if clean_sku_name(f) not in vr and pl_zip_images.get(f)}
                    if missing:
                        with st.spinner("🤖 AI Vision: Fabric, Sleeves, Pattern aur Color scan ho raha hai (Gemini ➜ NVIDIA ➜ Groq)..."):
                            new_res, failed = run_vision_detection(missing)
                            vr.update(new_res)
                        if new_res:
                            st.success(f"🤖 AI Vision: {len(new_res)} SKUs auto-detected!" +
                                       (f" | Fallback used for: {', '.join(failed)}" if failed else ""))
                        else:
                            st.warning("🤖 AI se result nahi mila — Python intelligence use hoga.")
                    st.session_state['vision_results'] = vr

                    ctx = {
                        'folders': folders, 'sizes': sizes_list, 'brand': brand_name.strip(),
                        'mrp': mrp_value.strip(), 'mfg': mfg_address.strip(),
                        'keywords': custom_keywords.strip(), 'dropbox': dropbox_url.strip(),
                        'live_links_cache': {},
                        'folder_images_count': {f: (len(pl_zip_images[f]) if f in pl_zip_images else 5) for f in folders},
                        'size_chart_cdn_url': None, 'quantity': int(sc_quantity),
                        'selling_price': str(sc_selling_price).strip(), 'sc_parent_model': str(sc_parent_model).strip(),
                        'td': int(default_td), 'default_category': None, 'vision_results': vr,
                    }

                    if auto_cdn and pl_zip_images:
                        with st.spinner("Images ko live CDN par upload kiya ja raha hai..."):
                            for fld, sc_info in pl_size_charts.items():
                                u = upload_to_imgbb_cached(sc_info['bytes'], IMGBB_PERMANENT_KEY, img_cache)
                                if u and not ctx['size_chart_cdn_url']:
                                    ctx['size_chart_cdn_url'] = u
                            for fld, img_tuples in pl_zip_images.items():
                                ctx['live_links_cache'][fld] = []
                                for iname, ibytes in img_tuples:
                                    u = upload_to_imgbb_cached(ibytes, IMGBB_PERMANENT_KEY, img_cache)
                                    if u: ctx['live_links_cache'][fld].append(u)

                    out, preview, added = None, [], 0
                    if marketplace == "Amazon - Vendor Central":
                        out, preview, added = run_vendor_central(tpl_file, ctx)
                        st.success(f"🎉 Amazon Vendor Central template filled! Total {added} variants.")
                        st.info("ℹ️ **Product Type column blank rakha gaya hai** — template me dropdown se manually select kar lena.")
                    elif marketplace == "Amazon - Seller Central":
                        out, preview, added = run_seller_central(tpl_file, ctx)
                        st.success(f"🎉 Amazon Seller Central template (`Template` sheet) filled! Total {added} rows (Parent + Child).")
                    else:
                        out, preview, added = run_generic_marketplace(tpl_file, ctx, marketplace)
                        st.success(f"🎉 {marketplace} template filled! Total {added} variants.")

                    if out:
                        if preview:
                            st.dataframe(preview, use_container_width=True)
                        st.download_button("📥 Download Filled File", data=out,
                                           file_name=tpl_file.name, mime=_mime_for(tpl_file.name), type="primary")
            except ValueError as ve:
                st.error(str(ve))
            except Exception as e:
                st.error(f"Error: {e}")
                with st.expander("Technical details"):
                    st.code(traceback.format_exc())


# ======================================================================
#   🔁 CONVERT MODE - sirf Vendor Central ➜ Seller Central
# ======================================================================
def render_convert_ui():
    st.subheader("3. Convert: Vendor Central ➜ Seller Central (Multi-Sheet + Smart Dedupe & Product Type)")
    st.info(
        "💡 **Kaise kaam karega:**\n"
        "1. Niche **Filled Vendor Central File** (1, 2, 4, 5 ya kitni bhi category sheets wali) aur **Blank Seller Central File** upload karein.\n"
        "2. **Smart Duplicate & Blank Row Merge:** Agar VC file me same SKU/ASIN 2 baar hai (ek row me details blank hain aur doosri row me details bhari hain), toh system **us SKU ko sirf 1 baar lega** aur blank row ko ignore karke filled row ka poora data fetch kar lega.\n"
        "3. Vendor Central file ki **sabhi listing sheets** automatically read hongi aur har sheet/category ke hisab se Blank Seller Central template ke **`Product Type`** dropdown se exact matching Product Type (`SHIRT`, `PANTS`, `SWEATSHIRT`, `KURTA`, `SUIT`, `DRESS`, `COAT`, etc.) select hoga.\n"
        "4. Har unique **`Model Name`** ko **`SKU`** column me **`Parent`** row ke roop me add kiya jayega.\n"
        "5. Sabhi Child SKUs me **`Parentage Level = Child`**, **`Parent SKU = Model Name`**, **`Product Id Type = ASIN`**, **`Product Id = Merchant Suggested ASIN`**, aur **`Shipping Template (IN) = Migrated Template`** set hoga.\n"
        "6. Sabhi rows me **`Listing Action = Create or Replace (Full Update)`**, **`Variation Theme Name = SIZE/COLOR`**, **`Item Condition = New`**, **`Fulfillment Channel Code (IN) = Fulfillment by Merchant (Default)`**, aur **`Quantity (IN) = 0`** rahega."
    )
    st.caption("ℹ️ Convert mode me Vendor Central file ki sabhi category sheets scan hongi, duplicate SKUs ki blank rows ignore karke filled data merge hoga, aur Seller Central ke `Template` sheet me replicate hongi.")
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        vc_file = st.file_uploader("📥 1. Filled Vendor Central File Upload Karein (.xlsx / .xlsm):", type=["xlsx", "xlsm"], key="vcfile")
    with c_up2:
        sc_template = st.file_uploader("📤 2. Blank Seller Central Template Upload Karein (.xlsx / .xlsm):", type=["xlsx", "xlsm"], key="sctpl")

    parent_pos_map = {
        "Har Model Name Group ke Pehle (Recommended - Parent ke niche uske Child SKUs)": "before_group",
        "Sabhi Child SKUs ke Baad (Pehle saare Child SKUs, fir niche Parent SKUs)": "bottom",
        "Sabse Upar (Pehle saare Parent SKUs, fir niche saare Child SKUs)": "top",
    }
    parent_pos_label = st.radio("Parent SKU Row Position (Model Name as Parent):",
                                list(parent_pos_map.keys()), index=0, horizontal=True)

    if st.button("🔁 Convert Vendor Central ➜ Seller Central"):
        if not vc_file or not sc_template:
            st.error("Dono files upload karein - Filled Vendor Central file aur Blank Seller Central template!")
            return
        try:
            ctx = {'sizes': [], 'brand': '', 'mrp': '', 'mfg': '', 'keywords': '', 'folders': [],
                   'live_links_cache': {}, 'quantity': 0, 'selling_price': '', 'sc_parent_model': '',
                   'td': int(default_td), 'default_category': None, 'vision_results': {},
                   'parent_position': parent_pos_map.get(parent_pos_label, "before_group")}
            with st.spinner("VC file ki sabhi sheets se data replicate ho raha hai + Duplicate/Blank SKUs merge ho rahe hain..."):
                out, preview, count, sheets_summary = convert_vc_to_sc(vc_file, sc_template, ctx)
            if out:
                parent_cnt = sum(1 for r in preview if r.get('Parentage Level') == 'Parent')
                child_cnt = sum(1 for r in preview if r.get('Parentage Level') == 'Child')
                merged_dups = sum(s.get('Duplicate/Blank Rows Merged', 0) for s in sheets_summary)
                st.success(
                    f"🎉 Conversion complete! Total **{len(sheets_summary)} VC Sheets** se **{count}** rows `Template` sheet me fill hui "
                    f"(**{parent_cnt} Parent** rows + **{child_cnt} Unique Child** rows | **{merged_dups} Duplicate/Blank rows** merge & clean ki gayi)."
                )
                if sheets_summary:
                    st.markdown("##### 📊 Scanned Vendor Central Sheets, Deduplication & Assigned Product Types")
                    st.dataframe(sheets_summary, use_container_width=True)
                if preview:
                    st.markdown("##### 📋 Converted Seller Central Rows Preview")
                    st.dataframe(preview, use_container_width=True)
                base_name, orig_ext = os.path.splitext(sc_template.name)
                out_ext = orig_ext if orig_ext.lower() in ['.xlsm', '.xlsx'] else '.xlsx'
                out_name = f"{base_name}_Converted{out_ext}"
                st.download_button("📥 Download Converted Seller Central File", data=out,
                                   file_name=out_name, mime=_mime_for(out_name))
            else:
                st.error("VC file ki kisi bhi sheet me koi data row nahi mili.")
        except Exception as e:
            st.error(f"Conversion error: {e}")
            with st.expander("🔍 Technical details"):
                st.code(traceback.format_exc())


if CONVERT_MODE:
    render_convert_ui()
else:
    render_listing_ui()
