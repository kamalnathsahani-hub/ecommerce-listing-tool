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

st.title("📦 Multi-Marketplace Auto Listing Generator (Universal Pro V17 — AI Vision)")
st.caption("Ajio | Amazon VC | VC➜SC | Amazon SC | Flipkart | Meesho | Myntra + 🤖 FREE Gemini Vision (Optional Tick)")

# ======================================================================
#                🔑 API KEYS & AI CONFIG (YAHI EDIT KARNA HOTA HAI)
# ======================================================================
IMGBB_PERMANENT_KEY = "137be55d19cf3d6a7ec051e369137041"

# 🤖 GEMINI AI VISION (100% FREE) - Aapki key by default set hai.
# Nayi key banane ke liye: https://aistudio.google.com/apikey
# (Key badalni ho to sirf neeche quotes ke andar nayi key paste karo)
DEFAULT_GEMINI_API_KEY = "AQ.Ab8RN6Iir4BoBSujYsKLtoZZkKzzSMlxIWT9LakOoiwawdVlpg"
DEFAULT_VISION_MODEL = "gemini-2.0-flash"   # Free tier ka sabse fast model

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

def detect_product_category(folder_name, user_keywords="", sheet_name=""):
    """Python intelligence: folder-name se category detect karta hai."""
    combined = f"{folder_name} {user_keywords} {sheet_name}".lower()
    if any(k in combined for k in ["hoodie", "hoody", "hooded"]):
        return "hoodie"
    elif any(k in combined for k in ["sweatshirt", "sweater"]):
        return "sweatshirt"
    elif any(k in combined for k in ["cargo"]):
        return "cargo"
    elif any(k in combined for k in ["jean", "denim"]):
        return "jeans"
    elif any(k in combined for k in ["t-shirt", "tshirt", "tee", "polo"]):
        return "tshirt"
    elif any(k in combined for k in ["shirt"]):
        return "shirt"
    elif any(k in combined for k in ["trouser", "chinos", "pant"]):
        return "trousers"
    elif any(k in combined for k in ["kurta", "kurti", "ethnic"]):
        return "kurta"
    elif any(k in combined for k in ["jacket", "coat", "blazer"]):
        return "jacket"
    elif any(k in combined for k in ["track", "jogger"]):
        return "joggers"
    elif any(k in combined for k in ["short"]):
        return "shorts"
    if "pant" in sheet_name.lower():
        return "trousers"
    elif "shirt" in sheet_name.lower():
        return "shirt"
    return "apparel"

CATEGORY_OPTIONS = ['hoodie', 'sweatshirt', 'tshirt', 'shirt', 'jeans', 'cargo',
                    'trousers', 'joggers', 'shorts', 'jacket', 'kurta', 'apparel']

def resolve_category(clean_style, ctx, sheet_name=""):
    """Priority: User Override > AI Vision > Folder Name > Default Category > apparel"""
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

# ======================================================================
# COLOR DETECTION (PYTHON INTELLIGENCE - typo fixes ke saath)
# ======================================================================
COMMON_COLORS = [
    "Beige", "LightGrey", "Grey", "DarkGrey", "Black", "Navy", "Blue", "LightBlue", "DarkBlue",
    "IceBlue", "White", "Off White", "Green", "Olive", "Brown", "Coffee", "Charcoal", "Khaki",
    "Red", "Maroon", "Yellow", "Orange", "Pink"
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
    """Folder-name se color: exact -> joined -> misspelling -> fuzzy -> Black"""
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
    """Priority: AI Vision (image me jo dikh raha) > Folder Name"""
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
    """Relevant Color Map: Maroon -> Red | Navy -> Blue | Khaki -> Beige"""
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
# POCKET DETECTION
# ======================================================================
POCKET_BY_CATEGORY = {
    'hoodie': 'Kangaroo Pocket', 'sweatshirt': 'Kangaroo Pocket',
    'cargo': 'Cargo Pocket', 'jeans': 'Five Pocket',
    'trousers': 'Slash Pocket', 'pants': 'Slash Pocket',
    'joggers': 'Side Pocket', 'shorts': 'Side Pocket',
    'jacket': 'Welt Pocket', 'shirt': 'Chest Pocket',
    'tshirt': 'No Pocket', 'kurta': 'Side Pocket', 'apparel': 'Side Pocket',
}

def get_pocket_for_cat(cat):
    return POCKET_BY_CATEGORY.get(cat, 'Side Pocket')

def resolve_pocket(clean_style, cat, ctx):
    """Priority: AI Vision > Category-based"""
    vr = ctx.get('vision_results') or {}
    if (vr.get(clean_style) or {}).get('pocket'):
        return vr[clean_style]['pocket']
    return get_pocket_for_cat(cat)

CATEGORY_LABELS = {
    'hoodie': 'Hoodie', 'sweatshirt': 'Sweatshirt', 'cargo': 'Cargo Pants',
    'jeans': 'Jeans', 'tshirt': 'T-Shirt', 'shirt': 'Shirt', 'trousers': 'Trousers',
    'kurta': 'Kurta', 'jacket': 'Jacket', 'joggers': 'Joggers', 'shorts': 'Shorts',
    'pants': 'Pants', 'apparel': 'Casual Wear',
}

HSN_CODE_MAP = {
    'jeans': '62034290', 'cargo': '62034200', 'trousers': '62034200', 'pants': '62034200',
    'joggers': '62034200', 'shorts': '62034200', 'tshirt': '61099090', 'shirt': '62052000',
    'hoodie': '61103000', 'sweatshirt': '61103000', 'jacket': '62014090', 'kurta': '62034200',
}

def get_hsn_code(cat):
    return HSN_CODE_MAP.get(cat, '62034200')

def get_item_type_name_from_dropdown(cat):
    mapping = {
        'hoodie': 'Hooded Sweatshirt', 'sweatshirt': 'Sweatshirt', 'joggers': 'Sweatpants',
        'jeans': 'Jeans', 'cargo': 'Cargo Pants', 'trousers': 'Casual Pants', 'pants': 'Casual Pants',
        'shorts': 'Shorts', 'tshirt': 'T-Shirt', 'shirt': 'Casual Shirt', 'kurta': 'Kurta',
        'jacket': 'Jacket', 'apparel': 'Casual Pants',
    }
    return mapping.get(cat, 'Casual Pants')

def get_style_from_dropdown(cat):
    mapping = {
        'hoodie': 'Pullover', 'sweatshirt': 'Pullover', 'joggers': 'Jogger',
        'cargo': 'Cargo', 'trousers': 'Chino', 'jeans': 'Classic',
    }
    return mapping.get(cat, 'Modern')

SC_PRODUCT_TYPE = {
    'hoodie': 'sweatshirt', 'sweatshirt': 'sweatshirt', 'jacket': 'jacket',
    'jeans': 'jeans', 'cargo': 'pants', 'trousers': 'pants', 'pants': 'pants',
    'joggers': 'pants', 'shorts': 'shorts', 'tshirt': 'shirt', 'shirt': 'shirt',
    'kurta': 'kurta', 'apparel': 'shirt',
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
# DROPDOWN-SAFETY ENGINE
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
    if options:
        junk = {'', 'select', 'select one', 'choose', 'choose one', 'dropdown'}
        options = [o for o in options if str(o).strip().lower() not in junk]
        return options or None
    return None

def pick_dropdown_value(ws, col_idx, desired, fallback=None):
    """VALUE SIRF DROPDOWN SE: exact -> partial -> fuzzy -> dropdown ka pehla option.
    Dropdown na mile to aapki di hui value hi likhi jaati hai."""
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
# 🤖 AI VISION ENGINE (FREE - GOOGLE GEMINI)
# ======================================================================
KNOWN_POCKETS = [
    'Kangaroo Pocket', 'Cargo Pocket', 'Five Pocket', 'Side Pocket', 'No Pocket',
    'Welt Pocket', 'Chest Pocket', 'Slash Pocket', 'Flap Pocket', 'Utility Pocket',
]

AI_CAT_MAP = {
    'hooded sweatshirt': 'hoodie', 'hooded': 'hoodie', 'hood': 'hoodie', 'pullover': 'hoodie',
    'sweat shirt': 'sweatshirt', 'sweater': 'sweatshirt', 'crewneck': 'sweatshirt',
    't-shirt': 'tshirt', 'tshirt': 'tshirt', 'tee': 'tshirt', 'polo': 'tshirt',
    'denim': 'jeans', 'denim jeans': 'jeans', 'denim pants': 'jeans',
    'pants': 'trousers', 'chinos': 'trousers', 'formal pants': 'trousers',
    'track pants': 'joggers', 'trackpants': 'joggers', 'sweatpants': 'joggers',
    'co-ord': 'apparel',
}

def guess_mime(filename):
    low = str(filename or '').lower()
    if low.endswith('.png'):
        return 'image/png'
    if low.endswith('.webp'):
        return 'image/webp'
    return 'image/jpeg'

def gemini_vision_detect(image_bytes, img_name, api_key, model="gemini-2.0-flash"):
    """Folder ki pehli image Gemini ko bhej kar Category+Color+Pocket JSON me leta hai."""
    if not api_key or not image_bytes:
        return None
    cats = ", ".join([c for c in CATEGORY_OPTIONS if c != 'apparel'])
    prompt = (
        "You are a fashion product cataloging expert. Analyze this garment product photo and reply with ONLY valid JSON, no markdown, no extra text. Format:\n"
        '{"category": "<one of: ' + cats + '>", "color": "<dominant visible color of the garment in simple English e.g. Maroon, Black, Navy, Grey, Beige, Olive, Khaki, Mustard, Off White>", "pocket": "<pocket style visible: Kangaroo Pocket / Cargo Pocket / Five Pocket / Side Pocket / Welt Pocket / Chest Pocket / No Pocket>"}\n'
        "Rules: category = the exact garment type you SEE. color = dominant garment color (NOT background). "
        "For a hoodie/sweatshirt with front pouch pocket use Kangaroo Pocket. If no visible pockets use No Pocket."
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
    close = difflib.get_close_matches(cat, [c for c in CATEGORY_OPTIONS if c != 'apparel'], n=1, cutoff=0.6)
    return close[0] if close else None

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

def run_vision_detection(zip_images_data, api_key, model):
    """Har folder ki PEHLI image analyze karta hai. Returns (results, failed_folders)."""
    results, failed = {}, []
    folders = [f for f in zip_images_data.keys() if zip_images_data.get(f)]
    if not folders:
        return results, failed
    progress = st.progress(0.0)
    status = st.empty()
    for i, fld in enumerate(folders):
        cs = clean_sku_name(fld)
        img_name, img_bytes = zip_images_data[fld][0]
        status.info(f"🤖 AI analyzing: **{cs}** ({i+1}/{len(folders)})...")
        data = gemini_vision_detect(img_bytes, img_name, api_key, model)
        if data:
            cat = canonicalize_ai_category(data.get('category'))
            col = canonicalize_ai_color(data.get('color'))
            poc = canonicalize_ai_pocket(data.get('pocket'))
            if cat or col or poc:
                results[cs] = {'category': cat, 'color': col, 'pocket': poc}
            else:
                failed.append(cs)
        else:
            failed.append(cs)
        time.sleep(0.6)
        progress.progress(min((i + 1) / len(folders), 1.0))
    progress.empty()
    status.empty()
    return results, failed

# ======================================================================
# CONTENT BUILDERS (GENERIC - NO SKU/COLOR)
# ======================================================================
def _user_kw_list(user_kw):
    if not user_kw:
        return []
    return [k.strip() for k in user_kw.replace(';', ',').split(',') if k.strip()]

def build_vendor_title(brand, category, full_sku, user_kw="", limit=75):
    b = f"{brand} " if brand else ""
    cat_str = category.capitalize()
    sku_part = f"({full_sku})"
    kws = _user_kw_list(user_kw)
    kw_candidate = kws[0] if kws else ""
    if kw_candidate:
        opt1 = f"{b}{cat_str} for men | {kw_candidate} {sku_part}".strip()
        if len(opt1) <= limit:
            title = opt1
        else:
            opt2 = f"{b}{kw_candidate} {sku_part}".strip()
            title = opt2 if len(opt2) <= limit else f"{b}{cat_str} {sku_part}".strip()
    else:
        title = f"{b}{cat_str} for men | casual wear {sku_part}".strip()
    if len(title) > limit:
        title = title[:limit].rstrip()
    words, counts, filtered = title.split(), {}, []
    for w in words:
        wl = w.lower().strip("|,()")
        c = counts.get(wl, 0)
        if c < 2 or w in ["|", "(", ")", "-"]:
            filtered.append(w)
            if wl:
                counts[wl] = c + 1
    return " ".join(filtered)[:limit]

def build_sc_title(brand, cat, full_sku, user_kw="", limit=150):
    b = f"{brand} " if brand else ""
    label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Casual Wear")
    kws = _user_kw_list(user_kw)
    base = f"{b}Men's {label} for Casual Wear"
    for n in range(min(3, len(kws)), -1, -1):
        t = f"{base} | {', '.join(kws[:n])} ({full_sku})" if n else f"{base} ({full_sku})"
        if len(t) <= limit:
            return t
    return t[:limit]

def generate_generic_keywords_200(cat, det_col, user_kw):
    if user_kw:
        kws, unique = _user_kw_list(user_kw), []
        for k in kws:
            if k.lower() not in [x.lower() for x in unique]:
                unique.append(k)
        return "; ".join(unique)[:200]
    if cat in ['hoodie', 'sweatshirt']:
        kw_list = [f"men {cat}", f"{det_col} hoodie", "pullover sweatshirt", "winter wear for men", "casual fleece hoodie", "streetwear sweatshirt", "comfortable hoodie"]
    elif cat == 'cargo':
        kw_list = ["men cargo pants", f"{det_col} cargo", "utility pants", "multi pocket trousers", "streetwear cargo", "relaxed fit pants", "men cotton cargo"]
    elif cat == 'jeans':
        kw_list = ["men jeans", f"{det_col} denim", "stretchable jeans", "casual denim pants", "mid rise jeans", "regular fit denim", "washed jeans for men"]
    else:
        kw_list = [f"men {cat}", f"{det_col} casual wear", "comfortable fit", "daily wear for men", "stylish modern wear", "breathable fabric"]
    return "; ".join(kw_list)[:200]

def build_freehand_bullets_and_desc(cat, brand):
    b_prefix = f"{brand} " if brand else ""
    cat_label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Casual Wear")

    if cat in ['hoodie', 'sweatshirt']:
        bullets = [
            "PREMIUM SOFT FLEECE FABRIC: Made from high-grade, breathable cotton blend fleece that keeps you warm, cozy, and comfortable throughout the day.",
            "URBAN STREETWEAR DESIGN: Features an on-trend silhouette with a functional hood and durable drawstrings.",
            "SPACIOUS KANGAROO POCKET: Front pocket provides convenient storage for daily essentials like mobile, wallet, or keys while keeping hands warm.",
            "VERSATILE EVERYDAY WEAR: Effortlessly pairs with jeans, cargo pants, joggers, and sneakers for casual outings, college, travel, and winter streetwear.",
            "WASH & CARE: Easy machine wash cold with similar colors inside out. Do not bleach. Warm iron if required."
        ]
        desc = (f"Upgrade your wardrobe with this stylish {b_prefix}Men's {cat_label}. "
                f"Crafted from premium quality cotton blend fleece, it delivers an ideal mix of modern streetwear style and snug comfort for daily wear.")
    elif cat == 'cargo':
        bullets = [
            "DURABLE COTTON UTILITY FABRIC: Tailored from rugged, breathable cotton fabric designed for active mobility and long-lasting durability.",
            "MULTI-POCKET FUNCTIONALITY: Equipped with deep tactical flap cargo pockets, side slash pockets, and secure rear pockets for maximum utility.",
            "COMFORTABLE RELAXED FIT: Features an ergonomic mid-rise waistband and relaxed leg cut for easy everyday movement.",
            "IDEAL FOR CASUAL & OUTDOOR WEAR: Perfect for travel, road trips, casual hangouts, and pairing with oversized tees, hoodies, or boots.",
            "WASH & CARE: Machine wash cold inside out. Tumble dry low. Do not bleach. Warm iron if needed."
        ]
        desc = (f"Experience peak comfort and utility with the {b_prefix}Men's {cat_label}. "
                f"Engineered with reinforced stitching and high-strength cotton fabric for an active, stylish lifestyle.")
    elif cat == 'jeans':
        bullets = [
            "COMFORT STRETCH DENIM: Crafted from premium stretchable cotton blend denim for flexible mobility and all-day ease.",
            "AUTHENTIC WASHED FINISH: Features a contemporary wash with subtle fading and durable contrast stitching.",
            "CLASSIC 5-POCKET DESIGN: Practical 5-pocket layout with sturdy metal rivets, heavy-duty zip fly, and button closure.",
            "TIMELESS RELAXED CUT: Sits comfortably at the waist with an authentic silhouette suitable for daily college, office casuals, and weekends.",
            "WASH & CARE: Machine wash cold inside out with similar colors. Do not bleach. Warm iron if required."
        ]
        desc = (f"Step out in confidence with the {b_prefix}Men's {cat_label}. "
                f"Designed with durable stretch denim, it combines all-day flexibility with an elevated, versatile everyday look.")
    elif cat in ['trousers', 'pants', 'joggers', 'shorts']:
        bullets = [
            "PREMIUM COTTON BLEND: Engineered from lightweight, breathable fabric offering superior all-day breathability and wrinkle resistance.",
            "TAILORED CASUAL SILHOUETTE: Clean flat front cut suitable for both smart-casual and relaxed settings.",
            "FUNCTIONAL POCKET DETAILS: Deep front slash pockets and buttoned rear pockets provide seamless convenience for daily essentials.",
            "SMART CASUAL STYLING: Pairs cleanly with polo t-shirts, casual shirts, sneakers, or loafers for office, travel, and outings.",
            "WASH & CARE: Machine wash gentle cycle. Do not bleach. Line dry in shade. Warm iron if required."
        ]
        desc = (f"Elevate your everyday comfort with the {b_prefix}Men's {cat_label}. "
                f"Crafted from premium breathable fabric, these deliver effortless versatility from morning errands to evening dinners.")
    else:
        bullets = [
            "HIGH-QUALITY COMFORT FABRIC: Crafted from premium breathable material engineered for soft skin feel and long-lasting daily wear.",
            "CONTEMPORARY STYLING: Features an on-trend silhouette with neat, durable seam construction and a comfortable fit.",
            "FUNCTIONAL & ERGONOMIC DESIGN: Tailored cut provides comfortable mobility and effortless wearability throughout the day.",
            "VERSATILE CASUAL PAIRING: Complements casual sneakers and daily wardrobe essentials for an easy, elevated style statement.",
            "WASH & CARE: Machine wash cold with like colors inside out. Do not bleach. Warm iron if required."
        ]
        desc = (f"Upgrade your daily rotation with the {b_prefix}Men's {cat_label}. "
                f"Engineered with premium fabric, it delivers the perfect balance of everyday durability and contemporary comfort.")
    return bullets, desc

def build_ajio_seo_content(cat, user_keywords=""):
    kw_str = user_keywords if user_keywords else "casual streetwear, relaxed fit, durable quality, all-day comfort"
    if cat == 'cargo':
        desc = ("CARGO FOR MEN - Designed for unmatched utility and streetwear style, this Men's Cargo Pants "
                "is crafted from premium breathable cotton fabric. Engineered with " + kw_str + ". "
                "Features multiple spacious flap cargo pockets, a comfortable flexi waist with sturdy button closure, "
                "and durable reinforced stitching throughout. Perfect for everyday wear, outdoor adventures, college, and travel.")
        highlight = "Premium breathable cotton fabric | Comfortable flexi waist | Multi-pocket functional cargo utility styling | Relaxed fit | Easy machine wash"
        add_info = "Styling Tip: Pair with oversized t-shirts, hoodies, and chunky sneakers. Wash Care: Machine wash cold with similar colors."
    elif cat == 'jeans':
        desc = ("Upgrade your everyday denim rotation with this trendy Men's Jeans. "
                "Crafted from premium stretchable denim engineered with " + kw_str + ". "
                "Featuring a contemporary mid-rise relaxed fit, functional 5-pocket styling, sturdy zip fly, and authentic washed finish. "
                "Ideal for casual outings, daily wear, and weekend hangouts.")
        highlight = "Premium stretch denim | Relaxed mid-rise silhouette | Authentic clean washed finish | Multi-pocket utility styling | All-season wear"
        add_info = "Styling Tip: Pair with casual shirts or boxy tees. Wash Care: Machine wash cold inside out."
    elif cat in ['hoodie', 'sweatshirt']:
        desc = ("Stay warm and stylish with this Men's Hoodie/Sweatshirt crafted from premium soft fleece fabric. "
                "Engineered with " + kw_str + ". Features a functional hood with durable drawstrings, spacious kangaroo pocket, "
                "and ribbed cuffs for a snug fit. Perfect for winter layering, college, travel, and everyday streetwear.")
        highlight = "Premium soft fleece fabric | Functional hood with drawstrings | Spacious kangaroo pocket | Ribbed cuffs | Machine washable"
        add_info = "Styling Tip: Layer over tees with jeans or joggers. Wash Care: Machine wash cold inside out."
    else:
        label = CATEGORY_LABELS.get(cat, 'Apparel')
        desc = (f"Men's {label} crafted from premium quality fabric featuring " + kw_str + ". "
                "Designed for maximum comfort, durability, and everyday style. Easy to pair and easy to care.")
        highlight = "Premium quality fabric | Comfortable fit | Durable stitching | Easy machine wash"
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
        if 'template' in sname.lower():
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
#                    AJIO ENGINE
# ======================================================================
def find_sheet_for_category(wb, category):
    ignore = ['instruction', 'guideline', 'extra', 'rejection', 'seller info']
    valid_sheets = [s for s in wb.sheetnames if not any(ign in s.lower() for ign in ignore)]

    def has(wl, sheet): return any(w in sheet.lower() for w in wl)

    if category in ["cargo", "trousers"]:
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
            color_family_val = map_to_standard_color_map(det_col)
            desc_text, high_text, add_text = build_ajio_seo_content(cat, ctx['keywords'])
            sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)
            fab_type_val = "Cotton" if cat in ['cargo', 'trousers', 'hoodie', 'sweatshirt', 'shirt'] else "Denim"
            pattern_val = "Solid" if cat not in ['jeans'] else "Washed"

            preview.append({
                'Sheet': target_sheet_name, 'SKU': clean_style, 'Category': cat,
                'Color': det_col, 'Color Family (Map)': color_family_val,
            })

            for sz in size_list:
                clean_sz = clean_sku_name(sz)
                full_sku = f"{clean_style}-{clean_sz}"

                try:
                    num_sz = int(re.findall(r'\d+', clean_sz)[0])
                    w_val, tf_val = str(num_sz + 1), str(num_sz - 1)
                except Exception:
                    w_val, tf_val = "34", "32"

                aset(ws, curr_r, col_style, clean_style)
                aset(ws, curr_r, col_desc, desc_text)
                aset(ws, curr_r, col_sku, full_sku)
                aset(ws, curr_r, col_brand, ctx['brand'])
                aset(ws, curr_r, col_ean, full_sku)
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
                aset(ws, curr_r, col_fab_detail, fab_type_val)
                aset(ws, curr_r, col_fab_type, fab_type_val, dropdown=True)
                aset(ws, curr_r, col_pattern, pattern_val, dropdown=True)
                aset(ws, curr_r, col_size_format, "UNI", dropdown=True)
                aset(ws, curr_r, col_size_group, "Regular", dropdown=True)
                aset(ws, curr_r, col_wash_care, "Machine wash", dropdown=True)

                aset(ws, curr_r, col_length, "Full length", dropdown=True)
                aset(ws, curr_r, col_stretch, "Stretch", dropdown=True)
                aset(ws, curr_r, col_style_type, "Cargo" if cat in ['cargo', 'trousers'] else "Relaxed", dropdown=True)
                aset(ws, curr_r, col_fit_type, "Relaxed Fit", dropdown=True)
                aset(ws, curr_r, col_front_style, "Flat Front", dropdown=True)
                aset(ws, curr_r, col_waist_rise, "Mid Rise", dropdown=True)
                aset(ws, curr_r, col_waist_type, "Flexi waist" if cat in ['cargo', 'trousers'] else "Fixed waist", dropdown=True)
                aset(ws, curr_r, col_meas_type, "Garment Measurement", dropdown=True)
                aset(ws, curr_r, col_to_fit, tf_val)
                aset(ws, curr_r, col_waist, w_val)
                aset(ws, curr_r, col_inseam, "32")

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
#              AMAZON VENDOR CENTRAL ENGINE
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
        cat_label = CATEGORY_LABELS.get(cat, cat.title() if cat else "Casual Wear")

        sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)
        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'])
        gen_keywords_200 = generate_generic_keywords_200(cat, det_col, ctx['keywords'])
        style_val = get_style_from_dropdown(cat)
        item_type_val = get_item_type_name_from_dropdown(cat)
        color_map_val = map_to_standard_color_map(det_col)
        pocket_val = resolve_pocket(clean_style, cat, ctx)
        mat_val = "Cotton" if cat in ['cargo', 'hoodie', 'sweatshirt'] else "Cotton Blend"

        preview.append({'SKU': clean_style, 'Category': cat, 'Color': det_col,
                        'Color Map': color_map_val, 'Item Type': item_type_val,
                        'Style': style_val, 'Pocket': pocket_val, 'Material': mat_val})

        for sz in ctx['sizes']:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            is_num_sz = clean_sz.isdigit()
            vendor_title = build_vendor_title(ctx['brand'], cat, full_sku, ctx['keywords'], 75)

            c_vc = get_vc_col(['vendor code'], ['rtip_vendor_code#1.value'])
            if c_vc:
                ws.cell(curr_r, c_vc).value = "J B fashion -- Surat -- Dropship, in_pd-ravi_apparel, JBZ54"

            c_sku = get_vc_col(['vendor sku'], ['vendor_sku#1.value'])
            if c_sku:
                ws.cell(curr_r, c_sku).value = full_sku

            c_pt = get_vc_col(['product type'], ['product_type#1.value'])
            if c_pt:
                if "pant" in ws.title.lower():
                    vset(c_pt, "PANTS", dropdown=True)
                elif "shirt" in ws.title.lower():
                    vset(c_pt, "SHIRT", dropdown=True)
                else:
                    vset(c_pt, "PANTS" if cat in ['jeans', 'cargo', 'trousers', 'pants'] else "SWEATSHIRT", dropdown=True)

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
                vset(c_dept, "Mens", dropdown=True)
            c_gen = get_vc_col(['target gender'], ['target_gender#1.value'])
            if c_gen:
                vset(c_gen, "Male", dropdown=True)
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
                vset(c_sld, "Long Sleeve" if cat in ['hoodie', 'sweatshirt', 'shirt'] else "Short Sleeve", dropdown=True)
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
#        AMAZON SELLER CENTRAL ENGINE
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
    'size_name': ['size name', 'size_name', 'size'],
    'size_system': ['size system', 'size_system'],
    'quantity': ['quantity', 'qty'],
    'standard_price': ['standard price', 'standard_price', 'price', 'selling price'],
    'list_price': ['list price', 'list_price', 'mrp', 'maximum retail price', 'mrp inr'],
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
    'item_weight_unit': ['item weight unit of measure', 'item_weight_unit_of_measure', 'weight unit of measure'],
    'hsn_code': ['hsn code', 'hsn_code', 'hsn', 'tax code'],
    'pocket_description': ['pocket description', 'pocket_description', 'pocket type'],
    'sleeve_length': ['sleeve length', 'sleeve_length', 'sleeve length description'],
    'waist_size': ['waist size', 'waist_size', 'waist'],
    'waist_unit': ['waist size unit of measure', 'waist_size_unit_of_measure', 'waist unit of measure'],
}

def find_sc_header(ws):
    scan_cols = min(ws.max_column, 60)
    for r in range(1, min(ws.max_row, 12) + 1):
        vals = {_norm_h(ws.cell(r, c).value) for c in range(1, scan_cols + 1)}
        if 'sku' in vals or 'sellersku' in vals:
            if vals & {'itemname', 'productname', 'brandname', 'itemtypename', 'producttype'}:
                return r
    for r in range(1, min(ws.max_row, 12) + 1):
        vals = {_norm_h(ws.cell(r, c).value) for c in range(1, scan_cols + 1)}
        if 'sku' in vals or 'sellersku' in vals:
            return r
    return None

def locate_sc_grid(ws):
    header_row = find_sc_header(ws)
    if not header_row:
        raise ValueError("Header row (sku/item_name) detect nahi hui - pehli 12 rows check karein.")
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

    bullets_cols, others_cols = {}, {}
    for c in range(1, ws.max_column + 1):
        h = _norm_h(ws.cell(header_row, c).value)
        if not h:
            continue
        for pat, tgt in [(r'bulletpoint(\d+)$', bullets_cols), (r'keyproductfeatures(\d+)$', bullets_cols),
                         (r'otherimageurl(\d+)$', others_cols)]:
            m = re.match(pat, h)
            if m:
                tgt[int(m.group(1))] = c

    sku_col = find(SC_FIELDS['sku'])
    if not sku_col:
        raise ValueError("'sku' column nahi mila template me.")

    start_row = header_row + 1
    while start_row <= ws.max_row:
        v = ws.cell(start_row, sku_col).value
        if v in (None, ''):
            low = [str(ws.cell(start_row, c).value or '').strip().lower()
                   for c in range(1, min(ws.max_column, 15) + 1)]
            if any(('sku' in x or 'item name' in x or 'brand' in x) for x in low if x):
                start_row += 1
                continue
            break
        start_row += 1
    return sku_col, start_row, find, bullets_cols, others_cols

def fill_sc_rows(ws, find, bullets_cols, others_cols, start_row, records, ctx):
    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)
    material_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    count = 0

    def W(row, key, value, dropdown=False):
        c = find(SC_FIELDS.get(key, [key]))
        if not c or value in (None, ''):
            return
        cell = ws.cell(row, c)
        if dropdown:
            value = pick_dropdown_value(ws, c, value, value)
        cell.value = value

    for rec in records:
        r = start_row + count
        W(r, 'sku', rec.get('sku'))
        W(r, 'product_type', rec.get('product_type'), dropdown=True)
        W(r, 'item_name', rec.get('item_name'))
        W(r, 'brand_name', rec.get('brand_name'))
        W(r, 'manufacturer', rec.get('manufacturer'))
        W(r, 'model_number', rec.get('model_number') or rec.get('sku'))
        W(r, 'part_number', rec.get('part_number') or rec.get('sku'))
        W(r, 'model_name', rec.get('model_name'))
        W(r, 'item_type_name', rec.get('item_type_name'), dropdown=True)
        W(r, 'color_name', rec.get('color_name'))
        W(r, 'color_map', rec.get('color_map'), dropdown=True)
        W(r, 'size_name', rec.get('size_name'))
        W(r, 'size_system', rec.get('size_system'), dropdown=True)

        bl = rec.get('bullets') or []
        for idx in sorted(bullets_cols):
            if idx - 1 < len(bl):
                ws.cell(r, bullets_cols[idx]).value = bl[idx - 1]

        gk = rec.get('generic_keywords')
        if gk:
            gk = re.sub(r'\s+', ' ', str(gk).replace(';', ' ')).strip()
        W(r, 'generic_keywords', gk)
        W(r, 'description', rec.get('description'))

        mc = find(SC_FIELDS['main_image_url'])
        if mc and rec.get('main_image_url'):
            ws.cell(r, mc).value = rec['main_image_url']
        oi = rec.get('other_images') or []
        for idx, c in sorted(others_cols.items()):
            if idx - 1 < len(oi):
                ws.cell(r, c).value = oi[idx - 1]

        mrp = rec.get('list_price')
        if mrp in (None, '') and ctx.get('mrp'):
            mrp = ctx['mrp']
        lc = find(SC_FIELDS['list_price'])
        if lc:
            cell = ws.cell(r, lc)
            if mrp not in (None, ''):
                cell.value = mrp
            else:
                cell.fill = mrp_fill
                cell.font = mrp_font
        W(r, 'standard_price', rec.get('standard_price') or ctx.get('selling_price') or None)

        q = rec.get('quantity')
        if q in (None, ''):
            q = ctx.get('quantity', 1)
        W(r, 'quantity', q)

        mcol = find(SC_FIELDS['material'])
        if mcol and rec.get('material'):
            cell = ws.cell(r, mcol)
            cell.value = rec['material']
            cell.fill = material_fill
        W(r, 'fabric_type', rec.get('fabric_type'))
        W(r, 'care_instructions', rec.get('care_instructions'), dropdown=True)
        W(r, 'country_of_origin', rec.get('country_of_origin'), dropdown=True)
        W(r, 'pattern_type', rec.get('pattern_type'), dropdown=True)
        W(r, 'fit_type', rec.get('fit_type'), dropdown=True)
        W(r, 'style_name', rec.get('style_name'), dropdown=True)
        W(r, 'department_name', rec.get('department_name'), dropdown=True)
        W(r, 'target_gender', rec.get('target_gender'), dropdown=True)
        W(r, 'age_range_description', rec.get('age_range_description'), dropdown=True)
        W(r, 'seasons', rec.get('seasons'), dropdown=True)
        W(r, 'lifestyle', rec.get('lifestyle'), dropdown=True)
        W(r, 'unit_count', rec.get('unit_count'))
        W(r, 'special_size_type', rec.get('special_size_type'), dropdown=True)
        if rec.get('item_weight') is not None:
            W(r, 'item_weight', rec['item_weight'])
            W(r, 'item_weight_unit', rec.get('item_weight_unit') or 'grams')
        W(r, 'hsn_code', rec.get('hsn_code'))
        W(r, 'pocket_description', rec.get('pocket_description'), dropdown=True)
        W(r, 'sleeve_length', rec.get('sleeve_length'), dropdown=True)
        if rec.get('waist_size'):
            W(r, 'waist_size', rec['waist_size'])
            W(r, 'waist_unit', rec.get('waist_unit') or 'inches')
        count += 1
    return count

def run_seller_central(uploaded_file, ctx):
    wb = load_workbook_clean(uploaded_file, uploaded_file.name)
    ws = pick_template_sheet(wb)
    try:
        ungroup_worksheet(ws)
    except Exception:
        pass

    sku_col, start_row, find, bullets_cols, others_cols = locate_sc_grid(ws)

    records, preview = [], []
    for raw_fld in ctx['folders']:
        clean_style = clean_sku_name(raw_fld)
        det_col = resolve_color(raw_fld, clean_style, ctx)
        cat = resolve_category(clean_style, ctx, ws.title)

        urls = get_sku_img_urls(raw_fld, ctx, clean_style)
        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'])
        gen_kw = generate_generic_keywords_200(cat, det_col, ctx['keywords'])
        style_val = get_style_from_dropdown(cat)
        item_type_val = get_item_type_name_from_dropdown(cat)
        color_map_val = map_to_standard_color_map(det_col)
        pocket_val = resolve_pocket(clean_style, cat, ctx)
        mat_val = "Cotton" if cat in ['cargo', 'hoodie', 'sweatshirt'] else "Cotton Blend"

        preview.append({'SKU': clean_style, 'Category': cat, 'Color': det_col,
                        'Color Map': color_map_val, 'Item Type': item_type_val,
                        'Style': style_val, 'Pocket': pocket_val, 'Material': mat_val})

        for sz in ctx['sizes']:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            sc_title = build_sc_title(ctx['brand'], cat, full_sku, ctx['keywords'], 150)
            records.append({
                'sku': full_sku,
                'product_type': SC_PRODUCT_TYPE.get(cat, 'pants'),
                'item_name': sc_title,
                'brand_name': ctx['brand'] or None,
                'manufacturer': ctx['mfg'] or None,
                'model_number': full_sku, 'part_number': full_sku,
                'model_name': f"{cat.capitalize()} {clean_style}",
                'item_type_name': item_type_val,
                'color_name': det_col, 'color_map': color_map_val,
                'size_name': clean_sz, 'size_system': 'IN',
                'bullets': bullets_list[:5],
                'generic_keywords': gen_kw, 'description': desc_text,
                'main_image_url': urls[0] if urls else None,
                'other_images': urls[1:9],
                'list_price': ctx['mrp'] or None,
                'standard_price': ctx['selling_price'] or None,
                'quantity': ctx['quantity'],
                'material': mat_val, 'fabric_type': f"100% {mat_val}",
                'care_instructions': 'Machine Wash', 'country_of_origin': 'India',
                'pattern_type': 'Solid', 'fit_type': 'Relaxed', 'style_name': style_val,
                'department_name': 'Mens', 'target_gender': 'male', 'age_range_description': 'adult',
                'seasons': 'All-Season', 'lifestyle': 'Casual', 'unit_count': 1,
                'special_size_type': 'Standard',
                'item_weight': 350, 'item_weight_unit': 'grams',
                'hsn_code': get_hsn_code(cat),
                'pocket_description': pocket_val,
                'sleeve_length': 'Long Sleeve' if cat in ['hoodie', 'sweatshirt', 'shirt'] else 'Short Sleeve',
                'waist_size': clean_sz if clean_sz.isdigit() else None,
                'waist_unit': 'inches' if clean_sz.isdigit() else None,
            })

    count = fill_sc_rows(ws, find, bullets_cols, others_cols, start_row, records, ctx)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), preview, count

# ======================================================================
#        UNIVERSAL ENGINE (Flipkart / Meesho / Myntra)
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
        'gender': ['gender'],
        'care': ['wash care', 'care instructions'],
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
        color_family_val = map_to_standard_color_map(det_col)
        bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'])
        gen_kw = re.sub(r'\s+', ' ', str(generate_generic_keywords_200(cat, det_col, ctx['keywords'])).replace(';', ' ')).strip()
        urls = get_sku_img_urls(raw_fld, ctx, clean_style)
        mat_val = "Cotton" if cat in ['cargo', 'hoodie', 'sweatshirt'] else "Cotton Blend"

        preview.append({'SKU': clean_style, 'Category': cat, 'Color': det_col,
                        'Color Family': color_family_val, 'Material': mat_val})

        for sz in ctx['sizes']:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            r = start_row + count
            title = build_sc_title(ctx['brand'], cat, full_sku, ctx['keywords'], 150)

            W(r, 'sku', full_sku)
            W(r, 'title', title)
            W(r, 'brand', ctx['brand'])
            W(r, 'description', desc_text)
            W(r, 'color', color_family_val, dropdown=True)
            W(r, 'size', clean_sz)
            W(r, 'fabric', mat_val)
            W(r, 'pattern', 'Solid', dropdown=True)
            W(r, 'fit', 'Relaxed', dropdown=True)
            W(r, 'country', 'India', dropdown=True)
            W(r, 'generic_kw', gen_kw)
            W(r, 'gender', 'Men', dropdown=True)
            W(r, 'care', 'Machine Wash', dropdown=True)
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
#              VC ➜ SC CONVERT ENGINE
# ======================================================================
def read_vendor_central_records(vc_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(vc_bytes), data_only=True)
    ws = pick_template_sheet(wb)
    label_row = key_row = None
    for r in range(1, min(ws.max_row, 8) + 1):
        vals = [str(ws.cell(r, c).value or '').strip().lower() for c in range(1, min(ws.max_column, 80) + 1)]
        if 'vendor sku' in vals:
            label_row = r
            break
    if label_row:
        key_row = label_row + 1
    else:
        label_row, key_row = 3, 4

    col_meta, sku_col = {}, None
    for c in range(1, ws.max_column + 1):
        lbl = str(ws.cell(label_row, c).value or '').strip()
        key = str(ws.cell(key_row, c).value or '').strip() if key_row else ''
        if not lbl and not key:
            continue
        col_meta[c] = (key, lbl)
        if (key.lower().startswith('vendor_sku') or lbl.lower() == 'vendor sku') and sku_col is None:
            sku_col = c
    if sku_col is None:
        sku_col = 2

    r, guard = key_row + 1, 0
    while r <= ws.max_row and ws.cell(r, sku_col).value in (None, '') and guard < 10:
        r += 1
        guard += 1

    records = []
    while r <= ws.max_row and ws.cell(r, sku_col).value not in (None, ''):
        rec = {}
        for c, (key, lbl) in col_meta.items():
            v = ws.cell(r, c).value
            if v in (None, ''):
                continue
            k1, k2 = key.strip().lower(), lbl.strip().lower()
            if k1 and k1 not in rec:
                rec[k1] = v
            if k2 and k2 not in rec:
                rec[k2] = v
        records.append(rec)
        r += 1
    return records

def vc_record_to_sc(rec, ctx):
    def g(*keys):
        for k in keys:
            v = rec.get(k)
            if v not in (None, ''):
                return v
        return None

    r = {}
    sku = str(g('vendor_sku#1.value', 'vendor sku') or '')
    r['sku'] = sku
    size = g('bottoms_size#1.size', 'apparel_size#1.size', 'apparel size value', 'bottoms size value', 'size value')
    if size in (None, '') and sku:
        m = re.search(r'-([A-Za-z0-9]+)$', sku)
        if m:
            size = m.group(1)
    style_part = sku
    if size and sku.endswith('-' + str(size)):
        style_part = sku[:-(len(str(size)) + 1)]
    cs = clean_sku_name(style_part)
    cat = resolve_category(cs, ctx, '')

    r['size_name'] = size
    r['size_system'] = g('bottoms_size#1.size_system', 'apparel_size#1.size_system') or 'IN'
    r['item_name'] = g('item_name#1.value', 'item name')
    r['brand_name'] = g('brand#1.value', 'brand name') or (ctx.get('brand') or None)
    r['manufacturer'] = g('manufacturer#1.value', 'manufacturer') or (ctx.get('mfg') or None)
    r['model_number'] = g('model_number#1.value', 'model number') or sku
    r['part_number'] = g('part_number#1.value', 'part number') or sku
    r['model_name'] = g('model_name#1.value', 'model name')
    vc_color = _canonicalize_color(g('color#1.value', 'color'))
    r['color_name'] = vc_color or resolve_color(style_part, cs, ctx)
    r['color_map'] = g('color#1.standardized_values#1', 'color map') or map_to_standard_color_map(r['color_name'])
    r['item_type_name'] = g('item_type_name#1.value', 'item type name') or get_item_type_name_from_dropdown(cat)
    r['style_name'] = g('style#1.value', 'style') or get_style_from_dropdown(cat)
    r['product_type'] = SC_PRODUCT_TYPE.get(cat, 'pants')

    bullets = []
    for i in range(1, 6):
        v = g(f'bullet_point#{i}.value', f'bullet point {i}')
        if v:
            bullets.append(str(v))
    r['bullets'] = bullets
    r['generic_keywords'] = g('generic_keyword#1.value', 'generic keywords')
    r['description'] = g('rtip_product_description#1.value', 'product description')

    imgs = []
    mi = g('main_product_image_locator#1.media_location', 'main image url')
    if mi:
        imgs.append(str(mi))
    for i in range(1, 9):
        v = g(f'other_product_image_locator_{i}#1.media_location', f'other image url {i}')
        if v:
            imgs.append(str(v))
    r['main_image_url'] = imgs[0] if imgs else None
    r['other_images'] = imgs[1:9]

    r['list_price'] = g('list_price#1.value_with_tax', 'maximum retail price (mrp)')
    r['standard_price'] = ctx.get('selling_price') or None
    r['quantity'] = ctx.get('quantity', 1)
    r['material'] = g('material#1.value', 'material')
    r['fabric_type'] = g('fabric_type#1.value', 'fabric type')
    r['care_instructions'] = g('care_instructions#1.value', 'care instructions') or 'Machine Wash'
    r['country_of_origin'] = g('country_of_origin#1.value', 'country of origin') or 'India'
    r['pattern_type'] = g('pattern#1.value', 'pattern type', 'pattern') or 'Solid'
    r['fit_type'] = g('fit_type#1.value', 'fit type') or 'Relaxed'
    r['department_name'] = g('department#1.value', 'department name') or 'Mens'
    r['target_gender'] = g('target_gender#1.value', 'target gender') or 'Male'
    r['age_range_description'] = g('age_range_description#1.value', 'age range description') or 'Adult'
    r['seasons'] = g('seasons#1.value', 'seasons') or 'All-Season'
    r['lifestyle'] = g('lifestyle#1.value', 'lifestyle') or 'Casual'
    r['unit_count'] = g('unit_count#1.value', 'unit count') or 1
    r['special_size_type'] = g('special_size_type#1.value', 'special size') or 'Standard'
    r['item_weight'] = g('item_weight#1.value', 'item weight')
    r['item_weight_unit'] = g('item_weight#1.unit', 'item weight unit') or 'Grams'
    r['hsn_code'] = g('external_product_information#1.value', 'hsn code')
    r['pocket_description'] = g('pocket_description#1.value', 'pocket description') or resolve_pocket(cs, cat, ctx)
    r['sleeve_length'] = g('sleeve length description', 'sleeve length', 'sleeve_length#1.value')
    r['waist_size'] = g('waist#1.size#1.value', 'waist size')
    r['waist_unit'] = g('waist#1.size#1.unit') or ('inches' if r['waist_size'] else None)
    return r

def convert_vc_to_sc(vc_file, sc_file, ctx):
    raw_records = read_vendor_central_records(vc_file.getvalue())
    if not raw_records:
        return None, [], 0
    records = [vc_record_to_sc(r, ctx) for r in raw_records]

    wb = load_workbook_clean(io.BytesIO(sc_file.getvalue()), sc_file.name)
    ws = pick_template_sheet(wb)
    try:
        ungroup_worksheet(ws)
    except Exception:
        pass

    sku_col, start_row, find, bullets_cols, others_cols = locate_sc_grid(ws)
    count = fill_sc_rows(ws, find, bullets_cols, others_cols, start_row, records, ctx)

    preview = [{'SKU': r.get('sku'), 'Title': r.get('item_name'), 'Color': r.get('color_name'),
                'Color Map': r.get('color_map'), 'Size': r.get('size_name'),
                'Item Type': r.get('item_type_name'), 'Style': r.get('style_name')} for r in records]

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
    st.info(f"💡 **{marketplace}** engine active hai — template is marketplace ke format ke according fill hoga.")

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

                st.success(f"✅ ZIP Se Total {len(detected_folders)} Product Folders (SKUs) Detect Hue:")
                st.code(", ".join([clean_sku_name(f) for f in detected_folders]))

                if detected_size_chart_info:
                    st.info(f"📐 **Size Chart Detect Hua:** `{detected_size_chart_info['filename']}`")

                # Naya ZIP aane par stale AI results clear karo
                valid_keys = {clean_sku_name(f) for f in detected_folders}
                st.session_state['vision_results'] = {
                    k: v for k, v in st.session_state['vision_results'].items() if k in valid_keys
                }
            except Exception as e:
                st.error(f"ZIP read karne me error: {e}")

        auto_cdn_upload = st.checkbox("🌐 ZIP images se automatic Live CDN Links banayein", value=True)
        st.success("⚡ Live Workable CDN Active: 100% working direct links banenge.")

    with col_in2:
        st.write("##### Option B: Dropbox Link / Manual SKUs")
        dropbox_base_url = st.text_input(
            "Dropbox Shared Folder Link (Optional):",
            placeholder="https://www.dropbox.com/scl/fo/xyz/..."
        )
        manual_skus_input = st.text_area(
            "Manual SKUs List (Agar ZIP nahi hai):",
            placeholder="SKU-1\nSKU-2\nSKU-4", height=100
        )
        if not detected_folders and manual_skus_input.strip():
            detected_folders = [s.strip() for s in manual_skus_input.split('\n') if s.strip()]
            for fld in detected_folders:
                folder_images_count[fld] = 5

    st.divider()

# ======================================================================
# 1a. 🤖 AI VISION (OPTIONAL TICK) + PREVIEW + OVERRIDES
# ======================================================================
default_category = None
if not CONVERT_MODE and detected_folders:
    st.subheader("1a. 🤖 AI Vision (Optional) + SKU-wise Category / Color Check")
    st.caption(
        "✅ **Tick ON** = Generate dabane par har folder ki pehli image **FREE Gemini AI** se analyze hogi "
        "(Category + Color + Pocket image dekh kar set honge). "
        "⬜ **Tick OFF** = 100% Python intelligence (folder-name detection, typo-fixes ke saath) — koi API use nahi hogi. "
        "Agar AI fail ho jaye (limit khatam/Network), to automatic Python intelligence par fall back hoga. "
        "Priority: **Override > AI > Folder Name**."
    )

    vr_all = st.session_state.get('vision_results', {})
    kw_now = st.session_state.get('kw_input', '')

    # LIVE PREVIEW (abhi ke available sources ke saath)
    _ctx_tmp = {'vision_results': vr_all, 'keywords': kw_now, 'default_category': None}
    prev_rows = []
    for fld in detected_folders:
        cs = clean_sku_name(fld)
        p_cat = resolve_category(cs, _ctx_tmp, "")
        p_col = resolve_color(fld, cs, _ctx_tmp)
        if cs in vr_all:
            source = '🤖 AI'
        elif cs in st.session_state.get('category_overrides', {}):
            source = '✍️ Override'
        else:
            source = '📄 Python'
        prev_rows.append({
            'Folder/SKU': cs, 'Category': p_cat, 'Color': p_col,
            'Color Map/Family': map_to_standard_color_map(p_col),
            'Item Type': get_item_type_name_from_dropdown(p_cat),
            'Style': get_style_from_dropdown(p_cat),
            'Pocket': resolve_pocket(cs, p_cat, _ctx_tmp),
            'Source': source,
        })
    st.dataframe(prev_rows, use_container_width=True)

    # ⬜ OPTION TICK — uncheck karte hi 100% Python intelligence chalega
    vision_enabled = st.checkbox(
        "🤖 AI Vision Enable Karein (FREE Gemini) — OFF karo to sirf Python intelligence chalega",
        value=True, key="vision_enabled",
        help="Tick OFF karne par koi API call nahi hogi — chahe limit poori ho ya na ho, file banane me koi dikkat nahi aayegi."
    )

    vision_api_key, vision_model = "", DEFAULT_VISION_MODEL
    if vision_enabled:
        if zip_images_data:
            cv1, cv2 = st.columns(2)
            with cv1:
                vision_api_key = st.text_input(
                    "Gemini API Key (FREE):", type="password",
                    value=DEFAULT_GEMINI_API_KEY, placeholder="AIza... ya AQ...",
                    key="vision_key",
                    help="Free key: aistudio.google.com/apikey — key code me already set hai, yahan se badal bhi sakte hain."
                )
            with cv2:
                vision_model = st.selectbox(
                    "Model:", ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-1.5-flash"],
                    index=0, key="vision_model"
                )

            c_run1, c_run2 = st.columns(2)
            with c_run1:
                if st.button("🔍 Abhi AI Detect Karo (Manual)"):
                    if not vision_api_key.strip():
                        st.error("Gemini API key khaali hai — upar daalein.")
                    else:
                        results, failed = run_vision_detection(zip_images_data, vision_api_key.strip(), vision_model)
                        st.session_state['vision_results'] = results
                        if results:
                            st.success(f"✅ {len(results)} folders AI se detect ho gaye! Preview update ho gaya.")
                            st.rerun()
                        else:
                            st.warning("AI se result nahi mila — folder-name detection use hoga.")
            with c_run2:
                st.caption("ℹ️ **Generate dabane par bhi AI automatic chalega** (agar tick ON hai aur results abhi tak nahi bane).")
        else:
            st.caption("ℹ️ AI Vision sirf ZIP mode me chalta hai (images locally hone chahiye). Dropbox/Manual mode me Python intelligence use hogi.")

    # Default Category + Per-SKU Override
    with st.expander("⚙️ Default Category & Per-SKU Override (galat detect ho to yahan badlein)"):
        c_o1, c_o2 = st.columns([1, 2])
        with c_o1:
            defcat_sel = st.selectbox("Default Category (jinme kuch detect nahi hua):",
                                      ["(auto)"] + CATEGORY_OPTIONS, key="defcat_sel")
            default_category = None if defcat_sel == "(auto)" else defcat_sel

        valid_keys = {clean_sku_name(f) for f in detected_folders}
        st.session_state['category_overrides'] = {
            k: v for k, v in st.session_state['category_overrides'].items() if k in valid_keys
        }
        ov_cols = st.columns(3)
        for i, fld in enumerate(detected_folders):
            cs = clean_sku_name(fld)
            ai_cat = (st.session_state['vision_results'].get(cs) or {}).get('category')
            auto_cat = ai_cat or detect_product_category(cs, kw_now, "")
            tag = "🤖" if ai_cat else ""
            with ov_cols[i % 3]:
                choice = st.selectbox(f"{cs} (auto: {auto_cat} {tag})", ["(auto)"] + CATEGORY_OPTIONS, key=f"ov_{cs}")
                if choice != "(auto)":
                    st.session_state['category_overrides'][cs] = choice
                else:
                    st.session_state['category_overrides'].pop(cs, None)

    st.divider()

# ======================================================================
# 3. SIZES, BRAND & LEGAL (SINGLE ADDRESS)
# ======================================================================
st.subheader("2. Sizes, Brand & Legal Information")
col_s1, col_s2 = st.columns([1, 2])

with col_s1:
    sizes_input = st.text_input("Sizes (comma-separated):", value="30, 32, 34, 36, 38")
    brand_name = st.text_input(
        "Brand Name:", value="",
        placeholder="Enter Brand Name (blank = not applicable)"
    )
    mrp_value = st.text_input("MRP (Khali chhodne par Yellow Highlight hoga):", value="", placeholder="e.g. 2999")

with col_s2:
    mfg_address = st.text_area(
        "Manufactured By (Address with Pincode):", value="",
        placeholder="Enter complete address with Pincode...\n(Yahi address Manufactured By, Marketed By aur Packer Details — sab jagah use hoga)",
        height=120
    )

if CONVERT_MODE:
    st.caption("ℹ️ Convert mode me data VC file se aayega — ye fields sirf blank fallback ke liye hain.")
st.divider()

sc_quantity, sc_selling_price = 1, ""
if marketplace in ["Amazon - Seller Central", "Convert: Vendor Central ➜ Seller Central"]:
    st.subheader("Seller Central Extra Fields")
    c_sc1, c_sc2 = st.columns(2)
    with c_sc1:
        sc_quantity = st.number_input("Default Stock Quantity:", min_value=1, value=5)
    with c_sc2:
        sc_selling_price = st.text_input("Selling Price (optional):", placeholder="e.g. 899")
    st.divider()

# ======================================================================
# 4. SEO KEYWORDS
# ======================================================================
custom_keywords = ""
if not CONVERT_MODE:
    st.subheader("3. SEO & Trending Keywords")
    custom_keywords = st.text_area(
        "Custom Trending Keywords (Optional):",
        placeholder="e.g. printed hoodie, casual hoodie, fleece pullover, cotton blend winter wear",
        key="kw_input"
    )
    st.divider()

# ======================================================================
# 5. TEMPLATE & GENERATE
# ======================================================================
def build_ctx():
    return {
        'folders': detected_folders,
        'sizes': [sz.strip() for sz in sizes_input.split(",") if sz.strip()],
        'brand': brand_name.strip(),
        'mrp': mrp_value.strip(),
        'mfg': mfg_address,
        'keywords': custom_keywords,
        'dropbox': dropbox_base_url if not CONVERT_MODE else "",
        'live_links_cache': {},
        'folder_images_count': folder_images_count,
        'size_chart_cdn_url': None,
        'quantity': int(sc_quantity),
        'selling_price': sc_selling_price.strip(),
        'td': int(default_td),
        'default_category': default_category,
        'vision_results': st.session_state.get('vision_results', {}),
    }

if CONVERT_MODE:
    st.subheader("3. Convert: VC File + SC Template Upload")
    st.info("💡 Filled **Vendor Central file** + khali **Seller Central template** upload karein.")
    vc_file = st.file_uploader("📥 Filled Vendor Central File (.xlsx/.xlsm):", type=["xlsx", "xlsm"], key="vcfile")
    sc_template = st.file_uploader("📤 Seller Central Template (.xlsx/.xlsm):", type=["xlsx", "xlsm"], key="sctpl")

    if st.button("🔁 Convert Vendor Central ➜ Seller Central"):
        if not vc_file or not sc_template:
            st.error("Dono files upload karein - VC filled file aur SC template!")
        else:
            try:
                ctx = build_ctx()
                with st.spinner("VC file read + SC template fill ho raha hai..."):
                    out, preview, count = convert_vc_to_sc(vc_file, sc_template, ctx)
                if out:
                    st.success(f"🎉 Conversion complete! Total {count} rows fill hui.")
                    if preview:
                        st.dataframe(preview, use_container_width=True)
                    out_name = f"{os.path.splitext(sc_template.name)[0]}_Converted.xlsx"
                    st.download_button("📥 Download Converted Seller Central File", data=out,
                                       file_name=out_name,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.error("VC file me koi data row nahi mili.")
            except Exception as e:
                st.error(f"Conversion error: {e}")
                with st.expander("🔍 Technical details"):
                    st.code(traceback.format_exc())

else:
    st.subheader(f"4. {marketplace} Template Upload & Generate")
    uploaded_file = st.file_uploader(
        f"Apni {marketplace} Excel Template (.xlsx ya .xlsm) Upload Karein:",
        type=["xlsx", "xlsm"]
    )

    if uploaded_file and st.button("🚀 Generate Listing File"):
        if not detected_folders:
            st.error("Kripya ZIP file upload karein ya SKUs list enter karein!")
        else:
            ctx = build_ctx()

            # ==========================================================
            # 🤖 AUTO AI VISION — SIRF TAB JAB TICK ON HO
            # Tick OFF = 100% Python intelligence (koi API call nahi)
            # Tick ON but API fail = automatic Python fallback
            # ==========================================================
            v_enabled = st.session_state.get('vision_enabled', True)
            v_key = (st.session_state.get('vision_key', '') or DEFAULT_GEMINI_API_KEY).strip()
            v_model = st.session_state.get('vision_model', DEFAULT_VISION_MODEL)

            if v_enabled and v_key and zip_images_data and not ctx['vision_results']:
                with st.spinner("🤖 AI Vision: har folder ki pehli image analyze ho rahi hai (Category + Color + Pocket)..."):
                    auto_results, auto_failed = run_vision_detection(zip_images_data, v_key, v_model)
                    ctx['vision_results'] = auto_results
                    st.session_state['vision_results'] = auto_results
                if auto_results:
                    st.success(f"🤖 AI Vision: {len(auto_results)} SKUs auto-detect ho gaye!" +
                               (f" | Fail (Python fallback hoga): {', '.join(auto_failed)}" if auto_failed else ""))
                    st.dataframe([
                        {'SKU': k, 'Category': v.get('category'), 'Color': v.get('color'), 'Pocket': v.get('pocket')}
                        for k, v in auto_results.items()
                    ], use_container_width=True)
                else:
                    st.warning("🤖 AI se result nahi mila — Python intelligence (folder-name detection) use hoga. "
                               "(Tip: Tick OFF karke bhi file ban sakti hai.)")
            elif not v_enabled and zip_images_data:
                st.info("⬜ AI Vision OFF hai — 100% Python intelligence se file ban rahi hai.")

            # CDN Upload (cached + progress bar)
            if auto_cdn_upload:
                with st.spinner("Images ko live CDN par upload kiya ja raha hai..."):
                    progress = st.progress(0.0)
                    total = sum(len(v) for v in zip_images_data.values())
                    if 'bytes' in detected_size_chart_info:
                        total += 1
                    done = 0
                    if 'bytes' in detected_size_chart_info:
                        ctx['size_chart_cdn_url'] = upload_to_imgbb_cached(
                            detected_size_chart_info['bytes'], IMGBB_PERMANENT_KEY, img_cache)
                        done += 1
                        progress.progress(min(done / max(total, 1), 1.0))
                    for fld, img_tuples in zip_images_data.items():
                        ctx['live_links_cache'][fld] = []
                        for img_name, img_bytes in img_tuples:
                            cdn_url = upload_to_imgbb_cached(img_bytes, IMGBB_PERMANENT_KEY, img_cache)
                            if cdn_url:
                                ctx['live_links_cache'][fld].append(cdn_url)
                            done += 1
                            progress.progress(min(done / max(total, 1), 1.0))
                    progress.empty()

            try:
                out, preview, added, extra = None, [], 0, None

                if marketplace == "Ajio":
                    out, preview, added, extra = run_ajio(uploaded_file, ctx)
                    st.success(f"🎉 Ajio template filled! Total {added} variants | Sheets: {', '.join(extra)}")
                    out_name = f"AJIO_MULTI_CATEGORY_LISTED.{'xlsm' if uploaded_file.name.lower().endswith('.xlsm') else 'xlsx'}"

                elif marketplace == "Amazon - Vendor Central":
                    out, preview, added = run_vendor_central(uploaded_file, ctx)
                    st.success(f"🎉 Amazon Vendor Central template filled! Total {added} variants.")
                    out_name = uploaded_file.name

                elif marketplace == "Amazon - Seller Central":
                    out, preview, added = run_seller_central(uploaded_file, ctx)
                    st.success(f"🎉 Amazon Seller Central template filled! Total {added} variants.")
                    out_name = uploaded_file.name

                elif marketplace in ["Flipkart", "Meesho", "Myntra"]:
                    out, preview, added = run_generic_marketplace(uploaded_file, ctx, marketplace)
                    st.success(f"🎉 {marketplace} template filled! Total {added} variants.")
                    out_name = uploaded_file.name

                if out:
                    if preview:
                        st.dataframe(preview, use_container_width=True)
                    st.download_button("📥 Download Filled File", data=out, file_name=out_name,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except ValueError as ve:
                st.error(str(ve))
            except Exception as e:
                st.error(f"Processing error: {e}")
                with st.expander("🔍 Technical details"):
                    st.code(traceback.format_exc())
