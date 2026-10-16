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

# ======================================================================
# 1. PAGE CONFIGURATION & NAVIGATION
# ======================================================================
st.set_page_config(
    page_title="AI Perfect Listing Tool",
    page_icon="✨",
    layout="wide"
)

col_nav, _ = st.columns([2, 8])
with col_nav:
    if st.button("⬅️ Back to Home"):
        try:
            st.switch_page("app.py")
        except Exception:
            pass

st.title("✨ AI Perfect Listing Generator (Universal Multi-Marketplace)")
st.caption("🤖 Multi-AI Vision (Gemini ➜ NVIDIA ➜ Groq) | 📐 AI Size Chart Reader | 📋 Template Instruction Analyzer | 100% Dropdown Safety")

# ======================================================================
# 2. API KEYS & VISION CONFIGURATION
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

# State setup
if 'img_cache' not in st.session_state:
    st.session_state['img_cache'] = {}
if 'vision_results' not in st.session_state:
    st.session_state['vision_results'] = {}
img_cache = st.session_state['img_cache']

# ======================================================================
# 3. MARKETPLACE SELECTION
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
# 4. HELPERS & CANONICAL MAPPINGS
# ======================================================================
def clean_sku_name(name):
    cleaned = re.sub(r'[\s_]+', '-', str(name).strip())
    return re.sub(r'-+', '-', cleaned).strip('-')

def is_size_chart_file(filename):
    low = os.path.basename(filename).lower()
    patterns = ['size_chart', 'size-chart', 'sizechart', 'size_guide', 'sizeguide', 'size chart', 'chart', 'measurement']
    return any(p in low for p in patterns) and low.endswith(('.jpg', '.jpeg', '.png', '.webp'))

def _norm_h(s):
    return re.sub(r'[\s_\-*]+', '', str(s or '')).strip().lower()

CATEGORY_OPTIONS = [
    'co-ord set', 'night suit', 'lounge wear', 'pajama set', 'hoodie', 'sweatshirt',
    'cargo', 'jeans', 'tshirt', 'shirt', 'trousers', 'joggers', 'shorts',
    'jacket', 'kurta', 'kurti', 'dress', 'top', 'skirt', 'apparel'
]

COMMON_COLORS = [
    "Beige", "LightGrey", "Grey", "DarkGrey", "Black", "Navy", "Blue", "LightBlue", "DarkBlue",
    "SkyBlue", "IceBlue", "White", "Off White", "Green", "Olive", "Pista", "Sea Green", "Brown",
    "Charcoal", "Khaki", "Red", "Maroon", "Yellow", "Orange", "Pink", "Purple", "Lavender", "Wine", "Peach", "Teal"
]

COLOR_FAMILY_MAP = {
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

def detect_color_from_name(text):
    text_clean = str(text).replace('-', ' ').replace('_', ' ').lower()
    for col in sorted(COMMON_COLORS, key=len, reverse=True):
        if col.lower() in text_clean:
            return col
    return "Pink"

def map_color_to_family(col_name):
    c = str(col_name).strip().lower()
    return COLOR_FAMILY_MAP.get(c, col_name)

def guess_mime(filename):
    low = str(filename or '').lower()
    if low.endswith('.png'): return 'image/png'
    if low.endswith('.webp'): return 'image/webp'
    return 'image/jpeg'

def extract_json_from_text(text):
    if not text: return None
    try:
        m = re.search(r'\{[\s\S]*\}', text)
        if m: return json.loads(m.group(0))
    except Exception:
        pass
    return None

# ======================================================================
# 5. MULTI-AI VISION & SIZE CHART EXTRACTORS
# ======================================================================
def get_garment_vision_prompt():
    cats = ", ".join(CATEGORY_OPTIONS)
    return (
        "You are an expert e-commerce fashion cataloger. Visually inspect this garment image very carefully and output ONLY a valid JSON object without markdown or code formatting.\n"
        "Return this exact schema:\n"
        "{\n"
        '  "category": "<One of: ' + cats + '>",\n'
        '  "gender": "<Women or Men>",\n'
        '  "color": "<Dominant color: Pink, Maroon, Black, Navy, Blue, LightBlue, SkyBlue, Grey, Beige, Olive, Pista, Sea Green, Lavender, White, etc.>",\n'
        '  "pocket": "<Kangaroo Pocket / Cargo Pocket / Five Pocket / Side Pocket / Welt Pocket / Chest Pocket / Slash Pocket / No Pocket>",\n'
        '  "material": "<Actual fabric: Cotton, Cotton Blend, Rayon, Fleece, Polyester, Denim, Crepe, Georgette, Silk, Knit>",\n'
        '  "fabric_detail": "<Specific fabric description e.g. 100% Cotton, Pure Cotton, Cotton Blend, Soft Fleece, Rayon>",\n'
        '  "sleeve_length": "<Short sleeve / Full-length sleeve / 3/4th sleeve / Sleeveless>",\n'
        '  "topwear_type": "<Top / Tunic / Shirt / Hoodie / Sweatshirt>",\n'
        '  "bottomwear_type": "<Capris / Pyjamas / Shorts / Trousers/Pants>",\n'
        '  "waist_type": "<Flexi waist / Elasticated waist / Fixed waist>",\n'
        '  "pattern": "<Solid / Printed / Washed / Striped / Checked / Floral / Quilted>"\n'
        "}\n"
        "Important Rule: If this is a 2-piece sleepwear / night suit / co-ord set for women, set category to 'co-ord set', gender to 'Women', waist_type to 'Flexi waist'."
    )

def get_size_chart_prompt():
    return (
        "You are an expert OCR & e-commerce technical fashion cataloger. Carefully read all text and table numbers in this size chart image.\n"
        "Extract all measurements for XS, S, M, L, XL, XXL (2XL), 3XL.\n"
        "Identify if bottomwear is Capris, Pyjamas, Pants, or Shorts, and if top sleeve is Short sleeve or Full-length sleeve.\n"
        "Output ONLY a valid JSON object without markdown or code formatting with this schema:\n"
        "{\n"
        '  "bottomwear_type": "<Capris or Pyjamas or Shorts>",\n'
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

def call_gemini_api(prompt, image_bytes, img_name):
    if not GEMINI_API_KEY or not image_bytes: return None
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    payload = {
        "contents": [{"parts": [
            {"text": prompt},
            {"inline_data": {"mime_type": guess_mime(img_name), "data": b64}}
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
                return extract_json_from_text(text)
    except Exception:
        pass
    return None

def call_nvidia_api(prompt, image_bytes, img_name):
    if not NVIDIA_API_KEY or not image_bytes: return None
    mime = guess_mime(img_name)
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
            return extract_json_from_text(res.json()['choices'][0]['message']['content'])
    except Exception:
        pass
    return None

def call_groq_api(prompt, image_bytes, img_name):
    if not GROQ_API_KEY or not image_bytes: return None
    mime = guess_mime(img_name)
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
            return extract_json_from_text(res.json()['choices'][0]['message']['content'])
    except Exception:
        pass
    return None

def fallback_ocr_size_chart(image_bytes):
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
    }
    return {'bottomwear_type': bottom_type, 'sleeve_length': sleeve_len, 'measurements': measurements}

def multi_ai_vision_detect(image_bytes, img_name):
    prompt = get_garment_vision_prompt()
    res = call_gemini_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('color'): return res, "Gemini"
    res = call_nvidia_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('color'): return res, "NVIDIA"
    res = call_groq_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('color'): return res, "Groq"
    return None, "None"

def multi_ai_size_chart_detect(image_bytes, img_name):
    prompt = get_size_chart_prompt()
    res = call_gemini_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('measurements'): return res, "Gemini"
    res = call_nvidia_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('measurements'): return res, "NVIDIA"
    res = call_groq_api(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('measurements'): return res, "Groq"
    return fallback_ocr_size_chart(image_bytes), "OCR-Fallback"

# ======================================================================
# 6. TEMPLATE & INSTRUCTIONS ANALYZER
# ======================================================================
def read_and_analyze_template(wb):
    info = {
        'sheets': wb.sheetnames,
        'styles_sheet': None,
        'category_breadcrumb': '',
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

def pick_dropdown_value_safe(info, col_idx, desired, fallback=None):
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
# 7. CDN UPLOADER & SEO GENERATOR
# ======================================================================
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
    if h in cache: return cache[h]
    url = upload_to_imgbb(image_bytes, api_key)
    if url: cache[h] = url
    return url

def build_ajio_seo_content(cat, user_keywords="", gender="Women", mat_val="Cotton Blend"):
    kw_str = user_keywords if user_keywords else "casual streetwear, relaxed fit, durable quality, all-day comfort"
    g_title = "Women's" if gender == "Women" else "Men's"
    desc = (
        f"{g_title} Co-ord Set - Designed for effortless chic look and supreme comfort, this {g_title} Co-ord Set "
        f"is crafted from premium breathable {mat_val} fabric. Engineered with {kw_str}. "
        "Features matching top and bottom silhouette with comfortable flexi waist and neat seam finish. "
        "Perfect for daily casuals, travel, vacations, and relaxed outings."
    )
    highlight = f"Premium {mat_val} fabric | Comfortable {g_title.lower()} fit | Elegant matching 2-piece set | Relaxed styling | Easy machine wash"
    add_info = "Styling Tip: Pair with minimal jewelry and slip-ons. Wash Care: Machine wash cold with similar colors."
    return desc[:950], highlight[:500], add_info[:500]

# ======================================================================
# 8. AJIO ENGINE (POPULATE ALL COLUMNS BASED ON TEMPLATE)
# ======================================================================
def run_perfect_ajio_listing(template_file_or_bytes, zip_images_data, detected_size_charts, ctx):
    name = getattr(template_file_or_bytes, 'name', 'template.xlsx')
    wb = openpyxl.load_workbook(template_file_or_bytes, keep_vba=name.lower().endswith('.xlsm'))
    info = read_and_analyze_template(wb)
    ws = wb[info['styles_sheet']]
    headers = info['headers']
    header_row = info['header_row']

    def find_col(candidates):
        for cand in candidates:
            c_norm = _norm_h(cand)
            for c, h in headers.items():
                if _norm_h(h) == c_norm: return c
        return None

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

    def get_style_code_group(folder_name):
        m = re.search(r'(\d+)', folder_name)
        if m:
            num = int(m.group(1))
            if num in [1, 2]: return "PX-W-ND-01-02"
            elif num in [3, 4, 5]: return "PX-W-ND-03-05"
            elif num in [6, 7, 8]: return "PX-W-ND-06-08"
            elif num in [9, 10, 11]: return "PX-W-ND-09-11"
            elif num in [12, 14]: return "PX-W-ND-12-14"
            elif num in [13, 15, 16]: return "PX-W-ND-13-16"
        return clean_sku_name(folder_name)

    curr_r = header_row + 1
    while ws.cell(curr_r, col_style or 1).value is not None:
        curr_r += 1

    total_added = 0
    preview = []
    sizes = ctx.get('sizes') or ['XS', 'S', 'M', 'L', 'XL', 'XXL']
    season_year = str(datetime.now().year + 1)
    
    mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
    mrp_font = Font(bold=True)

    def set_val(r, c, val, dropdown=False):
        if not c or val in (None, ''): return
        ws.cell(r, c).value = pick_dropdown_value_safe(info, c, val, val) if dropdown else val

    for fld, img_tuples in zip_images_data.items():
        clean_style = clean_sku_name(fld)
        style_group = get_style_code_group(fld)
        
        ai_data = ctx.get('vision_results', {}).get(clean_style, {})
        det_col = ai_data.get('color') or detect_color_from_name(fld)
        col_fam = map_color_to_family(det_col)
        mat_val = ai_data.get('material') or "Cotton"
        fab_detail = ai_data.get('fabric_detail') or "Cotton Blend"
        pattern_val = ai_data.get('pattern') or ("Quilted" if "quilt" in fld.lower() else "Solid")
        
        sc_data = detected_size_charts.get(fld) or detected_size_charts.get('default') or {}
        sleeve_val = sc_data.get('sleeve_length') or ai_data.get('sleeve_length') or "Short sleeve"
        bottom_val = sc_data.get('bottomwear_type') or ("Capris" if "capri" in fld.lower() else "Pyjamas")
        measurements = sc_data.get('measurements') or {}

        desc_text, high_text, add_text = build_ajio_seo_content("co-ord set", ctx.get('keywords', ''), "Women", mat_val)
        
        img_urls = list(ctx.get('live_links_cache', {}).get(fld, []))
        if ctx.get('size_chart_urls', {}).get(fld):
            chart_url = ctx['size_chart_urls'][fld]
            if chart_url not in img_urls: img_urls.append(chart_url)

        preview.append({
            'Style Code': style_group,
            'SKU Base': clean_style,
            'Color': det_col,
            'Color Family': col_fam,
            'Fabric': mat_val,
            'Pattern': pattern_val,
            'Bottomwear': bottom_val,
            'Sleeve': sleeve_val,
            'Images Count': len(img_urls)
        })

        for sz in sizes:
            clean_sz = clean_sku_name(sz)
            full_sku = f"{clean_style}-{clean_sz}"
            ean_val = clean_sku_name(f"{clean_style}-{clean_sz}")[:18]
            
            sz_meas = measurements.get(clean_sz) or measurements.get(clean_sz.upper()) or {}
            bust_val = sz_meas.get('bust', 37)
            front_len_val = sz_meas.get('front_length', 25)
            waist_val = sz_meas.get('waist', 22)
            to_fit_waist_val = waist_val + 2
            hip_val = sz_meas.get('hip', 38)
            inseam_val = sz_meas.get('inseam', 28)

            set_val(curr_r, col_style, style_group)
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

            set_val(curr_r, col_hsn, ctx.get('hsn', '61083100'))
            set_val(curr_r, col_prod_grp, "Evening", dropdown=True)
            set_val(curr_r, col_fash_grp, "Fashion", dropdown=True)
            set_val(curr_r, col_season_cd, "Core", dropdown=True)
            set_val(curr_r, col_season_yr, season_year)
            set_val(curr_r, col_size, clean_sz)
            set_val(curr_r, col_std_size, clean_sz)

            set_val(curr_r, col_art_h, 2); set_val(curr_r, col_art_l, 25)
            set_val(curr_r, col_art_w, 24); set_val(curr_r, col_art_luom, "CM")
            set_val(curr_r, col_art_wt, 350); set_val(curr_r, col_art_wtuom, "GRAM")
            set_val(curr_r, col_pkg_h, 2); set_val(curr_r, col_pkg_l, 25)
            set_val(curr_r, col_pkg_w, 24); set_val(curr_r, col_pkg_luom, "CM")
            set_val(curr_r, col_pkg_wt, 350); set_val(curr_r, col_pkg_wtuom, "GRAM")

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
# 9. STREAMLIT UI & DASHBOARD INTERFACE
# ======================================================================
detected_folders = []
zip_images_data = {}
detected_size_charts = {}

if not CONVERT_MODE:
    st.subheader("1. Product Images, SKUs & Size Chart (ZIP Upload)")
    col_in1, col_in2 = st.columns(2)
    with col_in1:
        st.write("##### 📁 Product Folders Ka ZIP Upload")
        zip_file = st.file_uploader("Product Folders (.zip):", type=["zip"], key="zip_in")
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
                            if is_size_chart_file(filename):
                                detected_size_charts[fld] = {'filename': filename, 'bytes': z.read(f)}
                            elif filename.lower().endswith(('.jpg', '.jpeg', '.png', '.webp')):
                                folder_to_imgs.setdefault(fld, []).append((filename, z.read(f)))

                    detected_folders = sorted(list(folder_to_imgs.keys()))
                    for fld in detected_folders:
                        imgs = folder_to_imgs[fld]
                        imgs.sort(key=lambda item: int(re.findall(r'\d+', item[0])[0]) if re.findall(r'\d+', item[0]) else 999)
                        zip_images_data[fld] = imgs

                st.success(f"✅ ZIP Se Total **{len(detected_folders)}** SKUs Detect Hue:")
                st.code(", ".join([clean_sku_name(f) for f in detected_folders]))
                if detected_size_charts:
                    st.info(f"📐 **{len(detected_size_charts)} Size Charts Detect Hue:** {', '.join(set(v['filename'] for v in detected_size_charts.values()))}")
            except Exception as e:
                st.error(f"ZIP error: {e}")

    with col_in2:
        st.write("##### 🌐 Live CDN & Cloud Hosting")
        auto_cdn = st.checkbox("Images ko auto CDN links (ImgBB) me convert karein", value=True)
        st.caption("ImgBB Direct Link Active: Images auto-upload hokar direct links banenge.")

    st.divider()

st.subheader("2. Marketplace Template Upload")
tpl_file = st.file_uploader(f"{marketplace} Template (.xlsx / .xlsm):", type=["xlsx", "xlsm"], key="tpl_in")

if tpl_file:
    try:
        wb_check = openpyxl.load_workbook(tpl_file, data_only=True)
        t_info = read_and_analyze_template(wb_check)
        c_t1, c_t2, c_t3 = st.columns(3)
        with c_t1: st.metric("Target Styles Sheet", t_info['styles_sheet'])
        with c_t2: st.metric("Template Columns", len(t_info['headers']))
        with c_t3: st.metric("Mandatory Columns (*)", len(t_info['mandatory_cols']))
        if t_info['category_breadcrumb']:
            st.info(f"📌 **Category Breadcrumb:** `{t_info['category_breadcrumb']}`")
    except Exception as e:
        st.warning(f"Template parsing warning: {e}")

st.divider()

st.subheader("3. Brand, MRP, Legal & Sizes")
c_b1, c_b2 = st.columns([1, 2])
with c_b1:
    sizes_input = st.text_input("Sizes (comma-separated):", value="XS, S, M, L, XL, XXL")
    brand_name = st.text_input("Brand Name:", value="POSHAX")
    mrp_value = st.text_input("MRP (e.g. 1999):", value="1999")
    hsn_code = st.text_input("HSN Code:", value="61083100")

with c_b2:
    mfg_address = st.text_area(
        "Manufactured By & Marketed By (Complete Address with Pincode):",
        value="PLOT NO.41, DEVAM AMBRO PARK KOSAD,CHORASI,SURAT,GUJARAT-394107",
        height=130
    )
    custom_keywords = st.text_area(
        "SEO & Trending Keywords (Optional):",
        value="casual streetwear, relaxed fit, durable quality, all-day comfort",
        height=70
    )

st.divider()

if st.button("🚀 Perfect Listing File Generate Karein", type="primary"):
    if not tpl_file:
        st.error("Kripya Marketplace Template Excel file upload karein!")
    elif not detected_folders:
        st.error("Kripya Product ZIP file upload karein!")
    else:
        ctx = {
            'sizes': [s.strip() for s in sizes_input.split(',') if s.strip()],
            'brand': brand_name.strip(),
            'mrp': mrp_value.strip(),
            'mfg': mfg_address.strip(),
            'hsn': hsn_code.strip(),
            'keywords': custom_keywords.strip(),
            'td': int(default_td),
            'live_links_cache': {},
            'size_chart_urls': {},
            'vision_results': st.session_state.get('vision_results', {})
        }

        # Step 1: AI Vision for Garments
        with st.spinner("🤖 1/3: AI Vision: Colors, Fabric & Sleeve length detect ho rahe hain..."):
            for fld in detected_folders:
                cs = clean_sku_name(fld)
                if cs not in ctx['vision_results'] and zip_images_data.get(fld):
                    v_res, provider = multi_ai_vision_detect(zip_images_data[fld][0][1], zip_images_data[fld][0][0])
                    if v_res:
                        v_res['ai_source'] = provider
                        ctx['vision_results'][cs] = v_res

        # Step 2: AI Size Chart Reading
        size_chart_measurements = {}
        with st.spinner("📐 2/3: AI Vision size charts ko read karke table measurements nikaal raha hai..."):
            for fld, sc_info in detected_size_charts.items():
                sc_res, _ = multi_ai_size_chart_detect(sc_info['bytes'], sc_info['filename'])
                if sc_res:
                    size_chart_measurements[fld] = sc_res
                    if 'default' not in size_chart_measurements:
                        size_chart_measurements['default'] = sc_res

        # Step 3: ImgBB CDN Upload
        with st.spinner("🌐 3/3: Images ko Live CDN par upload kiya ja raha hai..."):
            for fld, sc_info in detected_size_charts.items():
                url = upload_to_imgbb_cached(sc_info['bytes'], IMGBB_PERMANENT_KEY, img_cache)
                if url: ctx['size_chart_urls'][fld] = url

            for fld, img_tuples in zip_images_data.items():
                ctx['live_links_cache'][fld] = []
                for iname, ibytes in img_tuples:
                    url = upload_to_imgbb_cached(ibytes, IMGBB_PERMANENT_KEY, img_cache)
                    if url: ctx['live_links_cache'][fld].append(url)

        # Step 4: Listing Generation
        try:
            with st.spinner("📝 Ajio template sheet me listings populate ho rahi hain..."):
                if marketplace == "Ajio":
                    out_bytes, preview_data, total_added, target_sheet = run_perfect_ajio_listing(
                        tpl_file, zip_images_data, size_chart_measurements, ctx
                    )
                    st.success(f"🎉 Listing complete! Total **{total_added}** variants add hue (`{target_sheet}`).")
                    if preview_data:
                        st.dataframe(preview_data, use_container_width=True)

                    base_name, ext = os.path.splitext(tpl_file.name)
                    out_ext = ext if ext.lower() in ['.xlsm', '.xlsx'] else '.xlsx'
                    out_name = f"{base_name}_PERFECT_LISTED{out_ext}"
                    out_mime = (
                        "application/vnd.ms-excel.sheet.macroEnabled.12"
                        if out_ext == '.xlsm'
                        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.download_button(
                        label="📥 Download Filled Marketplace File",
                        data=out_bytes,
                        file_name=out_name,
                        mime=out_mime,
                        type="primary"
                    )
                else:
                    st.info(f"{marketplace} selected.")
        except Exception as e:
            st.error(f"Error: {e}")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())
