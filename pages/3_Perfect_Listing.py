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
    page_title="AI Perfect Listing Tool (Universal)",
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
st.caption("🤖 Multi-AI Vision (Gemini ➜ NVIDIA ➜ Groq) | 📐 Dynamic Size Chart Reader | 📋 Template Instruction Analyzer | 100% Dropdown Safety")

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
# 4. COLOR ABBREVIATION, SKU & STYLE CODE ENGINE (100% DYNAMIC)
# ======================================================================
COLOR_TOKENS = {
    'black', 'white', 'blue', 'lightblue', 'darkblue', 'skyblue', 'iceblue', 'navy', 'navyblue',
    'green', 'olive', 'pista', 'sea-green', 'purple', 'lavender', 'pink', 'red', 'maroon',
    'beige', 'brown', 'coffee', 'khaki', 'grey', 'gray', 'charcoal', 'yellow', 'orange',
    'peach', 'teal', 'wine', 'cream', 'rust', 'gold', 'silver'
}

COLOR_ABBR = {
    'black': 'Blk', 'white': 'Wht', 'blue': 'Blu', 'lightblue': 'LBlu',
    'darkblue': 'DBlu', 'skyblue': 'SBlu', 'iceblue': 'IBlu', 'navy': 'Nvy',
    'green': 'Grn', 'olive': 'Olv', 'pista': 'Pst', 'sea-green': 'SGrn',
    'purple': 'Prp', 'lavender': 'Lav', 'pink': 'Pnk', 'red': 'Red',
    'maroon': 'Mrn', 'beige': 'Beg', 'brown': 'Brn', 'coffee': 'Cfe',
    'khaki': 'Khk', 'grey': 'Gry', 'gray': 'Gry', 'charcoal': 'Chr',
    'yellow': 'Ylw', 'orange': 'Orn', 'peach': 'Pch', 'teal': 'Tel',
    'wine': 'Win', 'cream': 'Crm', 'ribbin': 'Rbn', 'teddy': 'Tdy'
}

def clean_sku_name(name):
    cleaned = re.sub(r'[\s_]+', '-', str(name).strip())
    return re.sub(r'-+', '-', cleaned).strip('-')

def extract_style_code(folder_name):
    """
    Folder name se color suffix hata kar natural Style Code auto-extract hota hai.
    Zero hardcoded prefix - Universal for Saree, Kurta, Jeans, Men, Kids, Night suit.
    """
    clean = clean_sku_name(folder_name)
    parts = clean.split('-')
    cut_idx = -1
    for i in range(len(parts) - 1, 0, -1):
        if parts[i].lower() in COLOR_TOKENS:
            cut_idx = i
            break
    if cut_idx > 0:
        return '-'.join(parts[:cut_idx])
    if len(parts) >= 2 and not parts[-1].isdigit():
        return '-'.join(parts[:-1])
    return clean

def generate_safe_ean(clean_style, clean_sz):
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
        if low in COLOR_ABBR:
            new_parts.append(COLOR_ABBR[low])
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

def is_size_chart_file(filename):
    low = os.path.basename(filename).lower()
    patterns = ['size_chart', 'size-chart', 'sizechart', 'size_guide', 'sizeguide', 'size chart', 'chart', 'measurement']
    return any(p in low for p in patterns) and low.endswith(('.jpg', '.jpeg', '.png', '.webp'))

def _norm_h(s):
    return re.sub(r'[\s_\-*]+', '', str(s or '')).strip().lower()

CATEGORY_OPTIONS = [
    'co-ord set', 'night suit', 'lounge wear', 'pajama set', 'saree', 'kurta', 'kurti',
    'jeans', 'cargo', 'trousers', 'joggers', 'shorts', 'tshirt', 'shirt', 'jacket',
    'hoodie', 'sweatshirt', 'dress', 'top', 'skirt', 'apparel'
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

# ======================================================================
# 5. DYNAMIC DIMENSIONS, WEIGHT & INTELLIGENT HSN
# ======================================================================
def get_dimensions_and_weight(category):
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

HSN_INTELLIGENCE_MAP = {
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

def resolve_hsn_code(user_hsn, category):
    if user_hsn and str(user_hsn).strip():
        return str(user_hsn).strip()
    cat = str(category or '').lower()
    for k, code in HSN_INTELLIGENCE_MAP.items():
        if k in cat:
            return code
    return '62042200'

# ======================================================================
# 6. MULTI-AI VISION & SIZE CHART EXTRACTORS
# ======================================================================
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

def get_garment_vision_prompt():
    cats = ", ".join(CATEGORY_OPTIONS)
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

def get_size_chart_prompt():
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
        '2XL': {'bust': 47, 'front_length': 27, 'waist': 32, 'hip': 51, 'inseam': 29},
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
# 7. TEMPLATE & INSTRUCTIONS ANALYZER
# ======================================================================
def read_and_analyze_template(wb):
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
# 8. CDN UPLOADER & DYNAMIC SEO CONTENT
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

def build_dynamic_seo_content(category, gender, brand, fabric, pattern, keywords=""):
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
    dims_data = get_dimensions_and_weight(cat_detected)
    hsn_val = resolve_hsn_code(ctx.get('hsn'), cat_detected)

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
        ws.cell(r, c).value = pick_dropdown_value_safe(info, c, val, val) if dropdown else val

    # Fill data for each product folder in zip
    for fld, img_tuples in zip_images_data.items():
        # EXACT BASE SKU FROM ZIP FOLDER (NO HARDCODED PREFIX)
        base_sku = clean_sku_name(fld)
        # NATURAL STYLE CODE EXTRACTED BY REMOVING COLOR SUFFIX
        style_code = extract_style_code(fld)
        
        ai_data = ctx.get('vision_results', {}).get(base_sku, {})
        det_col = ai_data.get('color') or detect_color_from_name(fld)
        col_fam = map_color_to_family(det_col)
        mat_val = ai_data.get('material') or "Cotton"
        fab_detail = ai_data.get('fabric_detail') or f"100% {mat_val}"
        pattern_val = ai_data.get('pattern') or ("Quilted" if "quilt" in fld.lower() else "Solid")
        
        sc_data = detected_size_charts.get(fld) or detected_size_charts.get('default') or {}
        sleeve_val = sc_data.get('sleeve_length') or ai_data.get('sleeve_length') or "Short sleeve"
        bottom_val = sc_data.get('bottomwear_type') or ("Capris" if "capri" in fld.lower() else "Pyjamas")
        measurements = sc_data.get('measurements') or {}

        desc_text, high_text, add_text = build_dynamic_seo_content(
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
            clean_sz = clean_sku_name(sz)
            # ITEM SKU = ZIP FOLDER NAME + '-' + SIZE
            full_sku = f"{base_sku}-{clean_sz}"
            # EAN = SAFE <= 18 CHARACTERS (SIZE SUFFIX ALWAYS PRESERVED)
            ean_val = generate_safe_ean(base_sku, clean_sz)
            
            sz_meas = measurements.get(clean_sz) or measurements.get(clean_sz.upper()) or {}
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
# 10. STREAMLIT UI & INTERFACE
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
                            # Direct folder name as base SKU
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
        st.caption("⚡ Direct CDN Link Active: ImgBB se direct public URLs auto-generate honge.")

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
    size_preset = st.radio(
        "Sizes Preset Chunein (Choose Size Format):",
        [
            "Alpha Sizes (S, M, L, XL, 2XL)",
            "Extended Alpha (XS, S, M, L, XL, 2XL)",
            "Numeric / Waist Sizes (28, 30, 32, 34, 36, 38)",
            "Custom"
        ],
        index=0
    )
    if size_preset == "Alpha Sizes (S, M, L, XL, 2XL)":
        preset_val = "S, M, L, XL, 2XL"
    elif size_preset == "Extended Alpha (XS, S, M, L, XL, 2XL)":
        preset_val = "XS, S, M, L, XL, 2XL"
    elif size_preset == "Numeric / Waist Sizes (28, 30, 32, 34, 36, 38)":
        preset_val = "28, 30, 32, 34, 36, 38"
    else:
        preset_val = "S, M, L, XL"

    sizes_input = st.text_input(
        "Sizes (comma-separated): Apne mutabiq add ya kam karein",
        value=preset_val
    )
    brand_name = st.text_input(
        "Brand Name:",
        value="",
        placeholder="e.g. PUMA (Khali chhodne par unbranded rahega)"
    )
    mrp_value = st.text_input(
        "MRP (optional):",
        value="",
        placeholder="e.g. 1999 (Khali chhodne par Yellow Highlight hoga)"
    )
    hsn_code = st.text_input(
        "HSN Code (optional):",
        value="",
        placeholder="e.g. 61083100 (Khali chhodne par category ke mutabiq auto-fill hoga)"
    )

with c_b2:
    mfg_address = st.text_area(
        "Manufactured By & Marketed By (Complete Address with Pincode):",
        value="",
        placeholder="e.g. Surat, Gujarat\n(Pura address pincode ke sath daalein)",
        height=130
    )
    custom_keywords = st.text_area(
        "SEO & Trending Keywords (Optional):",
        value="",
        placeholder="e.g. casual streetwear, relaxed fit, durable quality, all-day comfort",
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
                base_sku = clean_sku_name(fld)
                if base_sku not in ctx['vision_results'] and zip_images_data.get(fld):
                    v_res, provider = multi_ai_vision_detect(zip_images_data[fld][0][1], zip_images_data[fld][0][0])
                    if v_res:
                        v_res['ai_source'] = provider
                        ctx['vision_results'][base_sku] = v_res

        # Step 2: AI Size Chart Reading
        size_chart_measurements = {}
        with st.spinner("📐 2/3: AI Vision size charts ko read karke table measurements extract kar raha hai..."):
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
            with st.spinner("📝 Template sheet me perfect listing data populate ho raha hai..."):
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
