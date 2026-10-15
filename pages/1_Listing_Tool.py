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

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")

if st.button("⬅️ Back to Home"):
    st.switch_page("app.py")
st.title("📦 Multi-Marketplace Auto Listing Generator (Universal Pro V20)")
st.caption("Ajio | Amazon VC | VC➜SC | Amazon SC | Flipkart | Meesho | Myntra + 🤖 Dynamic AI Vision (Auto Gender & Category Detection)")

# ======================================================================
#                🔑 API KEYS & AI CONFIG
# ======================================================================
IMGBB_PERMANENT_KEY = "137be55d19cf3d6a7ec051e369137041"

# 🤖 GEMINI AI VISION KEY (Updated)
GEMINI_API_KEY = "AQ.Ab8RN6Kg7bXQGlqes-iZQH1sDrPOftdY-wAM67lP6YCyRI-E9A"
VISION_MODEL = "gemini-2.0-flash"

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
    'co-ord set', 'kurti', 'kurta', 'dress', 'top', 'tshirt', 'shirt', 
    'jeans', 'cargo', 'trousers', 'joggers', 'shorts', 'hoodie', 
    'sweatshirt', 'jacket', 'skirt', 'apparel'
]

def detect_product_category(folder_name, user_keywords="", sheet_name=""):
    combined = f"{folder_name} {user_keywords} {sheet_name}".lower()
    if any(k in combined for k in ["co-ord", "coord", "set", "two piece", "suit set"]):
        return "co-ord set"
    elif any(k in combined for k in ["hoodie", "hoody", "hooded"]):
        return "hoodie"
    elif any(k in combined for k in ["sweatshirt", "sweater"]):
        return "sweatshirt"
    elif any(k in combined for k in ["cargo"]):
        return "cargo"
    elif any(k in combined for k in ["jean", "denim"]):
        return "jeans"
    elif any(k in combined for k in ["t-shirt", "tshirt", "tee", "polo"]):
        return "tshirt"
    elif any(k in combined for k in ["kurti"]):
        return "kurti"
    elif any(k in combined for k in ["kurta", "ethnic"]):
        return "kurta"
    elif any(k in combined for k in ["dress", "frock", "gown"]):
        return "dress"
    elif any(k in combined for k in ["top", "tunic"]):
        return "top"
    elif any(k in combined for k in ["shirt"]):
        return "shirt"
    elif any(k in combined for k in ["trouser", "chinos", "pant"]):
        return "trousers"
    elif any(k in combined for k in ["jacket", "coat", "blazer"]):
        return "jacket"
    elif any(k in combined for k in ["track", "jogger"]):
        return "joggers"
    elif any(k in combined for k in ["short"]):
        return "shorts"
    elif any(k in combined for k in ["skirt"]):
        return "skirt"
    return "apparel"

def detect_gender(text):
    t = str(text).lower()
    if any(w in t for w in ["women", "woman", "female", "girl", "ladies", "kurti"]):
        return "Women"
    if any(w in t for w in ["men", "man", "male", "boy", "gents"]):
        return "Men"
    return "Women"  # Neutral/Default fallback if ambiguous

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
    return "Multicolor"

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
# POCKET & ATTRIBUTES MAPPING
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
    return POCKET_BY_CATEGORY.get(cat, 'No Pocket')

def resolve_pocket(clean_style, cat, ctx):
    vr = ctx.get('vision_results') or {}
    if (vr.get(clean_style) or {}).get('pocket'):
        return vr[clean_style]['pocket']
    return get_pocket_for_cat(cat)

CATEGORY_LABELS = {
    'co-ord set': 'Co-ord Set', 'kurti': 'Kurti', 'kurta': 'Kurta', 'dress': 'Dress',
    'top': 'Top', 'hoodie': 'Hoodie', 'sweatshirt': 'Sweatshirt', 'cargo': 'Cargo Pants',
    'jeans': 'Jeans', 'tshirt': 'T-Shirt', 'shirt': 'Shirt', 'trousers': 'Trousers',
    'jacket': 'Jacket', 'joggers': 'Joggers', 'shorts': 'Shorts', 'skirt': 'Skirt',
    'apparel': 'Casual Wear',
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
        'jacket': 'Jacket', 'skirt': 'Skirt', 'apparel': 'Casual Wear',
    }
    return mapping.get(cat, 'Casual Wear')

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
# DROPDOWN SAFETY ENGINE
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
        return None
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
                str(ws.cell(r, col_idx).value or '').strip()
                for r in (4, 5)
                if ws.cell(r, col_idx).value not in (None, '')
            }
            for dd_c in range(1, ws_dd.max_column + 1):
                dd_k = str(ws_dd.cell(3, dd_c).value or '').strip()
                if dd_k and dd_k in col_keys:
                    vals = [
                        str(ws_dd.cell(r, dd_c).value).strip()
                        for r in range(4, ws_dd.max_row + 1)
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
# 🤖 ADVANCED AI VISION ENGINE (DYNAMIC GENDER + GARMENT DETECTION)
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

def gemini_vision_detect(image_bytes, img_name, api_key, model):
    if not api_key or not image_bytes:
        return None
    cats = ", ".join(CATEGORY_OPTIONS)
    prompt = (
        "You are an expert e-commerce cataloging specialist. Inspect the given clothing product photo carefully.\n"
        "Analyze the product VISUALLY from the image itself without any pre-assumptions.\n"
        "Return ONLY a valid single-line JSON without any markdown formatting or codeblocks:\n"
        '{"category": "<one of: ' + cats + '>", "gender": "<Women or Men>", "color": "<dominant visible garment color>", '
        '"pocket": "<pocket style: Kangaroo Pocket / Cargo Pocket / Five Pocket / Side Pocket / Welt Pocket / Chest Pocket / No Pocket>", '
        '"material": "<visible material e.g. Cotton, Fleece, Denim, Rayon, Silk, Polyester, Crepe>"}\n'
        "Strict Guidelines:\n"
        "1. If it is a matching top and bottom / two-piece set for women, set category to 'co-ord set' and gender to 'Women'.\n"
        "2. Do NOT default to Men or Jeans unless clearly visible as Men's jeans.\n"
        "3. Accurately detect whether it is intended for Women or Men based on styling, cuts, and presentation."
    )
    try:
        b64 = base64.b64encode(image_bytes).decode('utf-8')
        payload = {
            "contents": [{"parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": guess_mime(img_name), "data": b64}}
            ]}],
            "generationConfig": {"temperature": 0.0, "maxOutputTokens": 300},
        }
        res = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}",
            json=payload, timeout=60
        )
        if res.status_code != 200:
            return None
        cands = res.json().get('candidates') or []
        if not cands:
            return None
        parts = (cands[0].get('content') or {}).get('parts') or []
        text = ''.join(str(p.get('text', '')) for p in parts)
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if not m:
            return None
        return json.loads(m.group(0))
    except Exception:
        return None

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
        return "Multicolor"
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
        return 'No Pocket'
    pl = p.lower()
    for k in KNOWN_POCKETS:
        if k.lower() == pl:
            return k
    close = difflib.get_close_matches(p, KNOWN_POCKETS, n=1, cutoff=0.5)
    return close[0] if close else p.title()

def run_vision_detection(zip_images_data, api_key, model):
    results, failed = {}, []
    folders = [f for f in zip_images_data.keys() if zip_images_data.get(f)]
    if not folders:
        return results, failed
    progress = st.progress(0.0)
    status = st.empty()
    for i, fld in enumerate(folders):
        cs = clean_sku_name(fld)
        img_name, img_bytes = zip_images_data[fld][0]
        status.info(f"🤖 AI analyzing photo: **{cs}** ({i+1}/{len(folders)})...")
        data = gemini_vision_detect(img_bytes, img_name, api_key, model)
        if data:
            cat = canonicalize_ai_category(data.get('category'))
            col = canonicalize_ai_color(data.get('color'))
            poc = canonicalize_ai_pocket(data.get('pocket'))
            gen = "Women" if str(data.get('gender')).lower().startswith('w') else "Men"
            mat = str(data.get('material') or 'Cotton Blend').title()
            results[cs] = {'category': cat, 'color': col, 'pocket': poc, 'gender': gen, 'material': mat}
        else:
            failed.append(cs)
        time.sleep(0.5)
        progress.progress(min((i + 1) / len(folders), 1.0))
    progress.empty()
    status.empty()
    return results, failed

# ======================================================================
# DYNAMIC CONTENT BUILDERS (GENDER + CATEGORY AWARE)
# ======================================================================
def _user_kw_list(user_kw):
    if not user_kw:
        return []
    return [k.strip() for k in user_kw.replace(';', ',').split(',') if k.strip()]

def build_vendor_title(brand, category, full_sku, user_kw="", limit=75, gender="Women"):
    b = f"{brand} " if brand else ""
    cat_label = CATEGORY_LABELS.get(category, category.title())
    sku_part = f"({full_sku})"
    kws = _user_kw_list(user_kw)
    kw_cand = kws[0] if kws else "Casual Wear"
    target_str = "for Women" if gender == "Women" else "for Men"
    opt1 = f"{b}{cat_label} {target_str} | {kw_cand} {sku_part}".strip()
    if len(opt1) <= limit:
        title = opt1
    else:
        opt2 = f"{b}{cat_label} {sku_part}".strip()
        title = opt2 if len(opt2) <= limit else f"{cat_label} {sku_part}"[:limit]
    return title[:limit]

def build_sc_title(brand, cat, full_sku, user_kw="", limit=150, gender="Women"):
    b = f"{brand} " if brand else ""
    label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Outfit")
    kws = _user_kw_list(user_kw)
    prefix = "Women's" if gender == "Women" else "Men's"
    base = f"{b}{prefix} {label} for Casual & Daily Wear"
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
    g_prefix = "women" if gender == "Women" else "men"
    cat_label = CATEGORY_LABELS.get(cat, cat)
    kw_list = [
        f"{g_prefix} {cat_label}", f"{det_col} {cat_label}", f"casual {cat_label}",
        f"comfortable wear for {g_prefix}", "trendy stylish design", "breathable comfortable fabric",
        "western wear", "daily wear"
    ]
    return "; ".join(kw_list)[:200]

def build_freehand_bullets_and_desc(cat, brand, gender="Women", det_mat="Cotton Blend"):
    b_prefix = f"{brand} " if brand else ""
    cat_label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Outfit")
    g_text = "Women's" if gender == "Women" else "Men's"

    if cat == 'co-ord set':
        bullets = [
            f"PREMIUM COMFORT FABRIC: Tailored from high quality, breathable {det_mat} ensuring supreme comfort and all-day ease.",
            "STYLISH TWO-PIECE SET: Features a coordinated matching top and bottom silhouette designed for an effortlessly chic look.",
            "VERSATILE DAY-TO-NIGHT WEAR: Perfect for brunches, vacations, daily casual outings, travel, and festive get-togethers.",
            "PERFECT FIT & MOBILITY: Ergonomically crafted with a relaxed cut and comfort waistband for flawless movement.",
            "WASH & CARE: Gentle machine wash cold with similar colors inside out. Do not bleach. Warm iron if required."
        ]
        desc = (f"Elevate your fashion game with this trendy {b_prefix}{g_text} {cat_label}. "
                f"Expertly crafted from soft, premium {det_mat}, this matching set combines contemporary charm with effortless everyday comfort.")
    elif cat in ['hoodie', 'sweatshirt']:
        bullets = [
            f"PREMIUM FLEECE WARMTH: Made from high-grade, breathable {det_mat} fleece keeping you snug and warm.",
            "CONTEMPORARY STREETWEAR: On-trend relaxed silhouette with durable stitching and modern finish.",
            "SPACIOUS KANGAROO POCKET: Front pouch pocket delivers convenience for small essentials while keeping hands cozy.",
            "VERSATILE CASUAL LAYERING: Pairs effortlessly with jeans, joggers, and sneakers for a chic streetwear appeal.",
            "WASH & CARE: Machine wash cold with similar colors. Do not bleach. Warm iron if needed."
        ]
        desc = (f"Stay warm and chic with the {b_prefix}{g_text} {cat_label}. "
                f"Crafted from premium {det_mat}, it brings optimal comfort and elevated style to your casual rotation.")
    elif cat in ['jeans', 'cargo', 'trousers', 'pants']:
        bullets = [
            f"PREMIUM DURABLE FABRIC: Engineered from high quality {det_mat} providing flexibility, breathability, and long-lasting durability.",
            "PRACTICAL POCKET DETAILS: Functional multi-pocket styling offers effortless utility for phones, keys, and wallets.",
            "ERGONOMIC RELAXED FIT: Designed with an optimal rise and comfort waist for unrestricted all-day movement.",
            "SMART CASUAL PAIRING: Seamlessly matches with shirts, crop tops, t-shirts, and sneakers.",
            "WASH & CARE: Machine wash cold inside out. Tumble dry low. Do not bleach."
        ]
        desc = (f"Experience unmatched comfort and style with the {b_prefix}{g_text} {cat_label}. "
                f"Designed with high-performance {det_mat} for all-day mobility and effortless everyday charm.")
    else:
        bullets = [
            f"HIGH-GRADE SOFT FABRIC: Crafted from premium {det_mat} providing a featherlight, breathable touch against the skin.",
            "MODERN TAILORED CUT: Features an elegant and flattering silhouette with neat, reinforced seams.",
            "ALL-DAY COMFORT: Breathable weave ensures you stay cool and comfortable throughout the day.",
            "EASY TO STYLE: Versatile design that pairs effortlessly with your favorite footwear and accessories.",
            "WASH & CARE: Gentle machine wash with like colors. Line dry in shade. Warm iron if required."
        ]
        desc = (f"Upgrade your daily wardrobe with the {b_prefix}{g_text} {cat_label}. "
                f"Combining premium {det_mat} and fine craftsmanship, it delivers the ideal balance of elegance and comfort.")
    return bullets, desc

def build_ajio_seo_content(cat, user_keywords="", gender="Women"):
    kw_str = user_keywords if user_keywords else "casual streetwear, relaxed fit, durable quality, all-day comfort"
    cat_label = CATEGORY_LABELS.get(cat, 'Apparel')
    g_title = "Women's" if gender == "Women" else "Men's"
    desc = (f"{g_title} {cat_label} crafted from premium quality fabric featuring {kw_str}. "
            "Designed for maximum comfort, durability, and everyday style. Easy to pair and easy to care.")
    highlight = f"Premium fabric | Comfortable {g_title.lower()} fit | Durable stitching | Easy machine wash"
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
            if has(['dress', 'top', 'coord', 'suit', 'kurti', 'women'], s): return s
    elif category in ["cargo", "trousers"]:
        for s in valid_sheets:
            if has(['trouser', 'pant', 'cargo'], s): return s
    elif category == "jeans":
        for s in valid_sheets:
            if 'jean' in s.lower(): return s
    elif category in ["hoodie", "sweatshirt"]:
        for s in valid_sheets:
            if has(['sweat', 'hoodie', 'fleece'], s): return s
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
        keywords = ['style', 'sku', 'mrp', 'price', 'size', 'image', 'model', 'brand', 'description', 'gender']
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
        col_gender = get_fallback_col(['Gender', 'Department', 'Target Gender'])
        col_size = get_fallback_col(['Size'])
        col_std_size = get_fallback_col(['StandardSize', 'Standard Size'])
        col_color = get_fallback_col(['Color Family', 'Colour Family'])
        col_pri_color = get_fallback_col(['Primary color', 'Primary Color', 'Color', 'Colour'])
        col_fab_type = get_fallback_col(['Fabric Type', 'Fabric Detail'])
        col_pattern = get_fallback_col(['Pattern'])
        col_fit_type = get_fallback_col(['Fit Type', 'Fit'])

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
            desc_text, high_text, add_text = build_ajio_seo_content(cat, ctx['keywords'], det_gen)
            sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)

            preview.append({
                'Sheet': target_sheet_name, 'SKU': clean_style, 'Category': cat,
                'Gender': det_gen, 'Color': det_col, 'Color Family': color_family_val,
            })

            for sz in size_list:
                clean_sz = clean_sku_name(sz)
                full_sku = f"{clean_style}-{clean_sz}"

                aset(ws, curr_r, col_style, clean_style)
                aset(ws, curr_r, col_desc, desc_text)
                aset(ws, curr_r, col_sku, full_sku)
                aset(ws, curr_r, col_brand, ctx['brand'])
                aset(ws, curr_r, col_ean, full_sku)
                aset(ws, curr_r, col_td, ctx['td'])
                aset(ws, curr_r, col_hsn, get_hsn_code(cat))
                aset(ws, curr_r, col_gender, det_gen, dropdown=True)
                aset(ws, curr_r, col_size, clean_sz)
                aset(ws, curr_r, col_std_size, clean_sz)
                aset(ws, curr_r, col_color, color_family_val, dropdown=True)
                aset(ws, curr_r, col_pri_color, det_col, dropdown=True)
                aset(ws, curr_r, col_fab_type, "Cotton Blend", dropdown=True)
                aset(ws, curr_r, col_pattern, "Solid", dropdown=True)
                aset(ws, curr_r, col_fit_type, "Relaxed Fit", dropdown=True)

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
#                      AMAZON VENDOR CENTRAL ENGINE
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
        vr_info = (ctx.get('vision_results') or {}).get(clean_style, {})
        mat_val = vr_info.get('material') or ("Cotton" if cat in ['cargo', 'hoodie'] else "Cotton Blend")

        cat_label = CATEGORY_LABELS.get(cat, cat.title() if cat else "Casual Wear")
        sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)
        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
        gen_keywords_200 = generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)
        style_val = get_style_from_dropdown(cat)
        item_type_val = get_item_type_name_from_dropdown(cat)
        color_map_val = map_to_standard_color_map(det_col)
        pocket_val = resolve_pocket(clean_style, cat, ctx)

        dept_val = "Womens" if det_gen == "Women" else "Mens"
        gender_val = "Female" if det_gen == "Women" else "Male"

        preview.append({'SKU': clean_style, 'Category': cat, 'Gender': det_gen, 'Color': det_col,
                        'Color Map': color_map_val, 'Item Type': item_type_val, 'Pocket': pocket_val, 'Material': mat_val})

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

            c_title = get_vc_col(['item name'], ['item_name#1.value'])
            if c_title:
                ws.cell(curr_r, c_title).value = vendor_title

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
                ws.cell(curr_r, c_fab).value = mat_val

            c_itn = get_vc_col(['item type name'], ['item_type_name#1.value'])
            if c_itn:
                vset(c_itn, item_type_val, dropdown=True)

            c_desc = get_vc_col(['product description'], ['rtip_product_description#1.value'])
            if c_desc:
                ws.cell(curr_r, c_desc).value = desc_text

            c_cmap = get_vc_col(['color map'], ['color#1.standardized_values#1'])
            if c_cmap:
                vset(c_cmap, color_map_val, dropdown=True)
            c_color = get_vc_col(['color'], ['color#1.value'])
            if c_color:
                ws.cell(curr_r, c_color).value = det_col

            c_poc = get_vc_col(['pocket description'], ['pocket_description#1.value'])
            if c_poc:
                vset(c_poc, pocket_val, dropdown=True)

            c_care = get_vc_col(['care instructions'], ['care_instructions#1.value'])
            if c_care:
                vset(c_care, "Machine Wash", dropdown=True)

            c_hsn_ent = get_vc_col(['external product information entity'], ['external_product_information#1.entity'])
            if c_hsn_ent:
                vset(c_hsn_ent, "HSN Code", dropdown=True)
            c_hsn_val = get_vc_col(['external product information'], ['external_product_information#1.value'])
            if c_hsn_val:
                ws.cell(curr_r, c_hsn_val).value = get_hsn_code(cat)

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

            c_coo = get_vc_col(['country of origin'], ['country_of_origin#1.value'])
            if c_coo:
                vset(c_coo, "India", dropdown=True)

            curr_r += 1
            added += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), preview, added

# ======================================================================
#             AMAZON SELLER CENTRAL ENGINE (DYNAMIC GENDER)
# ======================================================================
def _norm_amazon_key(k):
    if not k:
        return ""
    s = str(k).strip().lower()
    s = re.sub(r'\[(marketplace_id|language_tag|content_language|content_type)=[^\]]*\]', '', s)
    s = s.replace('[audience=all]', '')
    if s.startswith('rtip_'):
        s = s[5:]
    return s

def _detect_amz_sheet_rows(ws, default_lbl=3, default_key=4, default_data=7):
    r1_str = " ".join(str(ws.cell(1, c).value or '') for c in range(1, min(15, ws.max_column + 1)))
    lbl_m = re.search(r'labelRow=(\d+)', r1_str)
    key_m = re.search(r'attributeRow=(\d+)', r1_str)
    dat_m = re.search(r'dataRow=(\d+)', r1_str)
    if lbl_m and key_m and dat_m:
        return int(lbl_m.group(1)), int(key_m.group(1)), int(dat_m.group(1))
    return default_lbl, default_key, default_data

def run_seller_central(uploaded_file, ctx):
    wb = load_workbook_clean(uploaded_file, uploaded_file.name)
    ws = pick_template_sheet(wb)
    try:
        ungroup_worksheet(ws)
    except Exception:
        pass

    sc_lbl_r, sc_key_r, sc_data_r = _detect_amz_sheet_rows(ws, 4, 5, 7)

    sc_key_to_col, sc_lbl_occ_to_col, sc_lbl_counts = {}, {}, {}
    for c in range(1, ws.max_column + 1):
        lbl = str(ws.cell(sc_lbl_r, c).value or '').strip().lower()
        key = str(ws.cell(sc_key_r, c).value or '').strip() if sc_key_r else ''
        nk = _norm_amazon_key(key)
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
    c_mfg = get_sc_col(['manufacturer'], ['manufacturer#1.value'])
    c_main_img = get_sc_col(['main image url'], ['main_product_image_locator#1.media_location'])
    c_other_imgs = [get_sc_col(['other image url'], [f'other_product_image_locator_{i}#1.media_location'], occ=i) for i in range(1, 9)]
    c_desc = get_sc_col(['product description', 'description'], ['product_description#1.value'])
    c_bullets = [get_sc_col(['bullet point'], [f'bullet_point#{i}.value'], occ=i) for i in range(1, 6)]
    c_gkw = get_sc_col(['generic keywords', 'search terms'], ['generic_keyword#1.value'])
    c_dept = get_sc_col(['department name', 'department'], ['department#1.value'])
    c_gender = get_sc_col(['target gender', 'gender'], ['target_gender#1.value'])
    c_ssys = get_sc_col(['apparel size system', 'bottoms size system', 'size system'], ['apparel_size#1.size_system'])
    c_scls = get_sc_col(['apparel size class', 'bottoms size class', 'size class'], ['apparel_size#1.size_class'])
    c_sval = get_sc_col(['apparel size value', 'bottoms size value', 'size value', 'size name', 'size'], ['apparel_size#1.size'])
    c_cmap = get_sc_col(['color map'], ['color#1.standardized_values#1'])
    c_color = get_sc_col(['color', 'color name'], ['color#1.value'])
    c_mrp = get_sc_col(['maximum retail price (sell on amazon, in)', 'list price', 'mrp'], ['purchasable_offer#1.maximum_retail_price#1.schedule#1.value_with_tax'])
    c_price = get_sc_col(['your price inr (sell on amazon, in)', 'standard price', 'selling price'], ['purchasable_offer#1.our_price#1.schedule#1.value_with_tax'])
    c_coo = get_sc_col(['country of origin'], ['country_of_origin#1.value'])

    curr_r = sc_data_r
    while curr_r <= ws.max_row and ws.cell(curr_r, c_sku).value not in (None, ''):
        curr_r += 1

    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)
    preview = []
    total_written = 0

    for raw_fld in ctx['folders']:
        clean_style = clean_sku_name(raw_fld)
        det_col = resolve_color(raw_fld, clean_style, ctx)
        cat = resolve_category(clean_style, ctx, ws.title)
        det_gen = resolve_gender(clean_style, ctx)
        vr_info = (ctx.get('vision_results') or {}).get(clean_style, {})
        mat_val = vr_info.get('material') or "Cotton Blend"

        dept_val = "Womens" if det_gen == "Women" else "Mens"
        gender_val = "Female" if det_gen == "Women" else "Male"

        urls = get_sku_img_urls(raw_fld, ctx, clean_style)
        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
        gen_kw = generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)
        prod_type_raw = SC_PRODUCT_TYPE.get(cat, 'SUIT' if det_gen == 'Women' else 'SHIRT')

        for sz in ctx['sizes']:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            is_num_sz = clean_sz.isdigit()
            child_title = build_sc_title(ctx['brand'], cat, full_sku, ctx['keywords'], 150, det_gen)

            def s_col(col_idx, val, dropdown=False):
                if col_idx and val not in (None, ''):
                    ws.cell(curr_r, col_idx).value = pick_dropdown_value(ws, col_idx, val, val) if dropdown else val

            s_col(c_sku, full_sku)
            s_col(c_prod_type, prod_type_raw, dropdown=True)
            s_col(c_title, child_title)
            s_col(c_brand, ctx['brand'])
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
            s_col(c_dept, dept_val, dropdown=True)
            s_col(c_gender, gender_val, dropdown=True)
            s_col(c_ssys, "IN", dropdown=True)
            s_col(c_scls, "Numeric" if is_num_sz else "Alpha", dropdown=True)
            s_col(c_sval, clean_sz)
            s_col(c_cmap, map_to_standard_color_map(det_col), dropdown=True)
            s_col(c_color, det_col)
            s_col(c_coo, "India", dropdown=True)

            if ctx.get('selling_price'):
                s_col(c_price, ctx['selling_price'])
            if c_mrp:
                if ctx.get('mrp'):
                    mv = ctx['mrp']
                    ws.cell(curr_r, c_mrp).value = int(mv) if str(mv).isdigit() else mv
                else:
                    ws.cell(curr_r, c_mrp).fill = mrp_fill
                    ws.cell(curr_r, c_mrp).font = mrp_font

            preview.append({
                'Row': curr_r, 'SKU': full_sku, 'Category': cat, 'Gender': det_gen,
                'Color': det_col, 'Title': child_title
            })
            curr_r += 1
            total_written += 1

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
        'description': ['description', 'product description'],
        'mrp': ['mrp', 'maximum retail price', 'list price'],
        'price': ['selling price', 'price', 'your price'],
        'hsn': ['hsn', 'hsn code'],
        'color': ['color', 'color family', 'colour'],
        'size': ['size', 'size value', 'variant size'],
        'gender': ['gender', 'ideal for'],
        'country': ['country of origin', 'country'],
        'img1': ['main image url', 'image url 1', 'image 1', 'primary image url'],
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
        'gender': ['gender'],
        'country': ['country of origin', 'country'],
        'img1': ['image url 1', 'image 1', 'main image url'],
    },
    'Myntra': {
        'sku': ['seller sku', 'sku', 'sku id', 'style code'],
        'title': ['style name', 'product name', 'title', 'item name'],
        'brand': ['brand', 'brand name'],
        'description': ['style description', 'description', 'product description'],
        'mrp': ['mrp', 'maximum retail price'],
        'price': ['selling price', 'price'],
        'hsn': ['hsn', 'hsn code'],
        'color': ['color', 'colour', 'primary color'],
        'size': ['size', 'variant size'],
        'gender': ['gender', 'customer gender'],
        'country': ['country of origin', 'country'],
        'img1': ['image 1', 'main image url', 'primary image url'],
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
    header_row = best_row
    cols = {_norm_h(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}

    def find(cands):
        for cd in cands:
            n = _norm_h(cd)
            if n in cols:
                return cols[n]
        return None

    fcol = {k: find(v) for k, v in fields.items()}
    sku_col = fcol['sku'] or 1
    start_row = header_row + 1
    while start_row <= ws.max_row and ws.cell(start_row, sku_col).value not in (None, ''):
        start_row += 1

    def W(row, key, value, dropdown=False):
        c = fcol.get(key)
        if c and value not in (None, ''):
            ws.cell(row, c).value = pick_dropdown_value(ws, c, value, value) if dropdown else value

    count, preview = 0, []
    for raw_fld in ctx['folders']:
        clean_style = clean_sku_name(raw_fld)
        det_col = resolve_color(raw_fld, clean_style, ctx)
        cat = resolve_category(clean_style, ctx, ws.title)
        det_gen = resolve_gender(clean_style, ctx)
        color_family_val = map_to_standard_color_map(det_col)
        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen)
        urls = get_sku_img_urls(raw_fld, ctx, clean_style)

        preview.append({'SKU': clean_style, 'Category': cat, 'Gender': det_gen, 'Color': det_col})

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
            W(r, 'gender', det_gen, dropdown=True)
            W(r, 'country', 'India', dropdown=True)
            W(r, 'hsn', get_hsn_code(cat))
            if urls:
                W(r, 'img1', urls[0])

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
# 2. PRODUCT IMAGES & SKUs INPUT
# ======================================================================
dropbox_base_url = ""
if not CONVERT_MODE:
    st.subheader("1. Product Images & SKUs Input (ZIP Folder ya Dropbox Link)")
    col_in1, col_in2 = st.columns(2)

    with col_in1:
        st.write("##### Option A: Product Folders Ka ZIP Upload")
        zip_file = st.file_uploader("Product Folders Ka ZIP (.zip):", type=["zip"])
        if zip_file:
            try:
                with zipfile.ZipFile(zip_file, 'r') as z:
                    file_list = z.namelist()
                    for f in file_list:
                        if is_size_chart_file(f):
                            detected_size_chart_info['filename'] = os.path.basename(f)
                            detected_size_chart_info['bytes'] = z.read(f)
                            break

                    img_exts = ('.jpg', '.jpeg', '.png', '.webp')
                    folder_to_imgs = {}
                    for f in file_list:
                        f_clean = f.replace('\\', '/')
                        if f_clean.lower().endswith(img_exts) and not is_size_chart_file(f_clean):
                            sub_folder_name = os.path.basename(os.path.dirname(f_clean))
                            if sub_folder_name and sub_folder_name not in ['.', '']:
                                folder_to_imgs.setdefault(sub_folder_name, []).append(f)

                    detected_folders = sorted(list(folder_to_imgs.keys()))
                    for fld, imgs in folder_to_imgs.items():
                        def sort_key(p):
                            nums = re.findall(r'\d+', os.path.basename(p))
                            return int(nums[0]) if nums else 999
                        imgs.sort(key=sort_key)
                        folder_images_count[fld] = len(imgs)
                        zip_images_data[fld] = [(os.path.basename(p), z.read(p)) for p in imgs]

                st.success(f"✅ Total {len(detected_folders)} Product Folders Detect Hue:")
                st.code(", ".join([clean_sku_name(f) for f in detected_folders]))
            except Exception as e:
                st.error(f"ZIP error: {e}")

        auto_cdn_upload = st.checkbox("🌐 ZIP images se automatic Live CDN Links banayein", value=True)

    with col_in2:
        st.write("##### Option B: Dropbox Link / Manual SKUs")
        dropbox_base_url = st.text_input("Dropbox Shared Link:", placeholder="https://www.dropbox.com/...")
        manual_skus_input = st.text_area("Manual SKUs (Agar ZIP nahi hai):", placeholder="SKU-1\nSKU-2", height=100)
        if not detected_folders and manual_skus_input.strip():
            detected_folders = [s.strip() for s in manual_skus_input.split('\n') if s.strip()]
            for fld in detected_folders:
                folder_images_count[fld] = 5

    st.divider()

# ======================================================================
# 1a. 🤖 AI VISION (DYNAMIC DETECTION)
# ======================================================================
default_category = None
if not CONVERT_MODE and detected_folders:
    st.subheader("1a. 🤖 Dynamic AI Vision (Auto Category, Gender & Color)")
    st.caption("AI image ko visually analyze karke Category, Gender (Women/Men), Color aur Pocket detect karega. Kuch bhi hardcoded Men/Women nahi rahega.")

    vision_enabled = st.checkbox(
        "🤖 AI Vision Enable Karein (Photo dekhkar Women/Men aur Garment auto-detect hoga)",
        value=True, key="vision_enabled"
    )

    with st.expander("⚙️ Manual Category / Gender Overrides (Agar manually adjust karna ho)"):
        ov_cols = st.columns(3)
        for i, fld in enumerate(detected_folders):
            cs = clean_sku_name(fld)
            with ov_cols[i % 3]:
                st.write(f"**{cs}**")
                g_choice = st.selectbox(f"Gender ({cs})", ["(auto)", "Women", "Men"], key=f"gen_{cs}")
                if g_choice != "(auto)":
                    st.session_state['gender_overrides'][cs] = g_choice
                c_choice = st.selectbox(f"Category ({cs})", ["(auto)"] + CATEGORY_OPTIONS, key=f"cat_{cs}")
                if c_choice != "(auto)":
                    st.session_state['category_overrides'][cs] = c_choice

    st.divider()

# ======================================================================
# 3. SIZES (UPGRADED PRESETS), BRAND & LEGAL
# ======================================================================
st.subheader("2. Sizes, Brand & Legal Information")
col_s1, col_s2 = st.columns([1, 2])

with col_s1:
    size_preset = st.radio(
        "Sizes Preset Chunein:",
        ["Alpha Sizes (S, M, L, XL, 2XL, 3XL)", "Numeric Sizes (30, 32, 34, 36, 38)", "Custom"],
        index=0,
        horizontal=True
    )
    if size_preset == "Alpha Sizes (S, M, L, XL, 2XL, 3XL)":
        default_sz = "S, M, L, XL, 2XL, 3XL"
    elif size_preset == "Numeric Sizes (30, 32, 34, 36, 38)":
        default_sz = "30, 32, 34, 36, 38"
    else:
        default_sz = "S, M, L, XL"

    sizes_input = st.text_input(
        "Sizes (comma-separated): Jo chahiye rakhein, baaki delete karein",
        value=default_sz
    )
    brand_name = st.text_input("Brand Name:", value="", placeholder="Enter Brand Name")
    mrp_value = st.text_input("MRP (Khali chhodne par Yellow Highlight hoga):", value="", placeholder="e.g. 1999")

with col_s2:
    mfg_address = st.text_area(
        "Manufactured By (Address with Pincode):", value="",
        placeholder="Enter complete address with Pincode...\n(Yahi Manufactured By aur Packer Details me use hoga)",
        height=120
    )

st.divider()

# ======================================================================
# 4. TEMPLATE & GENERATE
# ======================================================================
def build_ctx():
    return {
        'folders': detected_folders,
        'sizes': [sz.strip() for sz in sizes_input.split(",") if sz.strip()],
        'brand': brand_name.strip(),
        'mrp': mrp_value.strip(),
        'mfg': mfg_address,
        'keywords': "",
        'dropbox': dropbox_base_url,
        'live_links_cache': {},
        'folder_images_count': folder_images_count,
        'size_chart_cdn_url': None,
        'quantity': 0,
        'selling_price': "",
        'td': int(default_td),
        'default_category': None,
        'vision_results': st.session_state.get('vision_results', {}),
    }

st.subheader(f"3. {marketplace} Template Upload & Generate")
uploaded_file = st.file_uploader(f"Apni {marketplace} Template (.xlsx ya .xlsm) Upload Karein:", type=["xlsx", "xlsm"])

if uploaded_file and st.button("🚀 Generate Listing File"):
    if not detected_folders:
        st.error("Kripya ZIP upload karein ya SKUs enter karein!")
    else:
        ctx = build_ctx()
        v_enabled = st.session_state.get('vision_enabled', True)

        if v_enabled and GEMINI_API_KEY.strip() and zip_images_data and not ctx['vision_results']:
            with st.spinner("🤖 AI Vision photo dekhkar Category, Women/Men gender aur Details extract kar raha hai..."):
                auto_results, auto_failed = run_vision_detection(zip_images_data, GEMINI_API_KEY.strip(), VISION_MODEL)
                ctx['vision_results'] = auto_results
                st.session_state['vision_results'] = auto_results
            if auto_results:
                st.success(f"🤖 AI Vision ne {len(auto_results)} SKUs successfully analyze kar liye!")
                st.dataframe([
                    {'SKU': k, 'Category': v.get('category'), 'Gender': v.get('gender'), 'Color': v.get('color'), 'Material': v.get('material')}
                    for k, v in auto_results.items()
                ], use_container_width=True)

        if auto_cdn_upload and zip_images_data:
            with st.spinner("Images ko live CDN par upload kiya ja raha hai..."):
                for fld, img_tuples in zip_images_data.items():
                    ctx['live_links_cache'][fld] = []
                    for img_name, img_bytes in img_tuples:
                        cdn_url = upload_to_imgbb_cached(img_bytes, IMGBB_PERMANENT_KEY, img_cache)
                        if cdn_url:
                            ctx['live_links_cache'][fld].append(cdn_url)

        try:
            out, preview, added = None, [], 0
            if marketplace == "Ajio":
                out, preview, added, extra = run_ajio(uploaded_file, ctx)
            elif marketplace == "Amazon - Vendor Central":
                out, preview, added = run_vendor_central(uploaded_file, ctx)
            elif marketplace == "Amazon - Seller Central":
                out, preview, added = run_seller_central(uploaded_file, ctx)
            else:
                out, preview, added = run_generic_marketplace(uploaded_file, ctx, marketplace)

            if out:
                st.success(f"🎉 {marketplace} template successfully fill ho gayi! Total {added} rows added.")
                if preview:
                    st.dataframe(preview, use_container_width=True)
                st.download_button("📥 Download Filled File", data=out, file_name=f"LISTED_{uploaded_file.name}")
        except Exception as e:
            st.error(f"Error: {e}")
            with st.expander("Details"):
                st.code(traceback.format_exc())
