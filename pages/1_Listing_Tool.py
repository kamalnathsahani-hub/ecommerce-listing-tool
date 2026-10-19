import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io, os, re, json, time, zipfile, base64, hashlib, difflib, traceback, requests, colorsys
from datetime import datetime
from PIL import Image

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")
if st.button("⬅️ Back to Home"): st.switch_page("app.py")
st.title("📦 Multi-Marketplace Auto Listing Generator (Universal Pro V25 - Unified Dashboard)")
st.caption("Ajio | Amazon VC | VC➜SC | Amazon SC | Flipkart | Meesho | Myntra + 🤖 AI Vision (Gemini ➜ NVIDIA ➜ Groq Fallback)")

def _get_secret(name, default=""):
  try: val = st.secrets[name]
  except Exception: val = os.environ.get(name, default)
  return str(val or default).strip()

IMGBB_PERMANENT_KEY = _get_secret("IMGBB_API_KEY", "137be55d19cf3d6a7ec051e369137041")
GEMINI_API_KEY = _get_secret("GEMINI_API_KEY", "AQ.Ab8RN6Lb4kCeCRQtjVtw5MYcz7Wl0x4dBRvNxsYCAD9TYoUk2Q")
GEMINI_VISION_MODEL = _get_secret("GEMINI_VISION_MODEL", "gemini-2.0-flash")
NVIDIA_API_KEY = _get_secret("NVIDIA_API_KEY", "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")
NVIDIA_MODEL = _get_secret("NVIDIA_MODEL", "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning")
GROQ_API_KEY = _get_secret("GROQ_API_KEY", "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")
GROQ_VISION_MODEL = _get_secret("GROQ_VISION_MODEL", "llama-3.2-11b-vision-preview")

_missing_keys = [n for n, v in [("IMGBB_API_KEY", IMGBB_PERMANENT_KEY), ("GEMINI_API_KEY", GEMINI_API_KEY), ("NVIDIA_API_KEY", NVIDIA_API_KEY), ("GROQ_API_KEY", GROQ_API_KEY)] if not v]
if _missing_keys:
  st.warning("⚠️ Ye keys `secrets.toml` me nahi mili: **" + ", ".join(_missing_keys) + "**. Missing AI key skip hogi, aur IMGBB_API_KEY na ho to CDN links nahi banenge.")

col_m1, col_m2 = st.columns([2, 1])
with col_m1:
  marketplace = st.selectbox("Marketplace Chunein:", ["Ajio", "Amazon - Vendor Central", "Convert: Vendor Central ➜ Seller Central", "Amazon - Seller Central", "Flipkart", "Meesho", "Myntra"])
with col_m2:
  default_td = st.number_input("TD / Default Margin (%) [Ajio ke liye]:", min_value=0, max_value=100, value=50, help="Ajio template ke TD column me ye value jayegi.")

st.divider()
CONVERT_MODE = marketplace.startswith("Convert:")

HIGHLIGHT_FILL = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
HIGHLIGHT_FONT = Font(bold=True)
DESC_MAX_LEN = 200

detected_folders, folder_images_count, zip_images_data, detected_size_chart_info = [], {}, {}, {}
for _k in ('img_cache', 'category_overrides', 'gender_overrides', 'vision_results'):
  if _k not in st.session_state: st.session_state[_k] = {}
img_cache = st.session_state['img_cache']

def clean_sku_name(name):
  cleaned = re.sub(r'[\s_]+', '-', str(name).strip())
  return re.sub(r'-+', '-', cleaned).strip('-')

def is_size_chart_file(filename):
  low = os.path.basename(filename).lower()
  patterns = ['size_chart', 'size-chart', 'sizechart', 'size_guide', 'sizeguide', 'size chart', 'chart', 'measurement']
  return any(p in low for p in patterns) and low.endswith(('.jpg', '.jpeg', '.png', '.webp'))

def _norm_h(s): return re.sub(r'[\s_\-*]+', '', str(s or '')).strip().lower()

def parse_product_zip(zip_file):
  IMG_EXT = ('.jpg', '.jpeg', '.png', '.webp')
  folder_to_imgs, size_charts, global_chart = {}, {}, None
  with zipfile.ZipFile(zip_file, 'r') as z:
    entries = []
    for info in z.infolist():
      if info.is_dir(): continue
      path = info.filename.replace('\\', '/').strip('/')
      parts = [p for p in path.split('/') if p]
      if len(parts) < 2 or parts[0] == '__MACOSX' or any(p.startswith('.') for p in parts): continue
      fname = parts[-1]
      if fname.lower() == 'thumbs.db' or not fname.lower().endswith(IMG_EXT): continue
      entries.append((parts[:-1], fname, info))
    while entries and len({d[0] for d, _, _ in entries if d}) == 1 and all(len(d) >= 1 for d, _, _ in entries) and any(len(d) >= 2 for d, _, _ in entries):
      entries = [(d[1:], f, i) for d, f, i in entries]
    for dir_parts, fname, info in entries:
      if not dir_parts:
        if is_size_chart_file(fname) and global_chart is None: global_chart = {'filename': fname, 'bytes': z.read(info)}
        continue
      fld = dir_parts[0]
      data = z.read(info)
      if is_size_chart_file(fname): size_charts[fld] = {'filename': fname, 'bytes': data}
      else: folder_to_imgs.setdefault(fld, []).append((fname, data))
  if global_chart:
    for fld in folder_to_imgs: size_charts.setdefault(fld, global_chart)
  size_charts = {k: v for k, v in size_charts.items() if k in folder_to_imgs}
  for fld in folder_to_imgs:
    folder_to_imgs[fld].sort(key=lambda item: int(re.findall(r'\d+', item[0])[0]) if re.findall(r'\d+', item[0]) else 999)
  return folder_to_imgs, size_charts

CATEGORY_OPTIONS = ['co-ord set', 'hoodie', 'sweatshirt', 'cargo', 'jeans', 'tshirt', 'shirt', 'trousers', 'joggers', 'shorts', 'jacket', 'kurta', 'kurti', 'dress', 'top', 'skirt', 'apparel']

def detect_product_category(folder_name, user_keywords="", sheet_name=""):
  combined = f"{folder_name} {user_keywords} {sheet_name}".lower()
  if any(k in combined for k in ["co-ord", "coord", "two piece", "suit set", "apparel set", "clothing set"]): return "co-ord set"
  elif any(k in combined for k in ["hoodie", "hoody", "hooded"]): return "hoodie"
  elif any(k in combined for k in ["sweatshirt", "sweater", "pullover", "cardigan"]): return "sweatshirt"
  elif any(k in combined for k in ["cargo"]): return "cargo"
  elif any(k in combined for k in ["jean", "denim"]): return "jeans"
  elif any(k in combined for k in ["t-shirt", "tshirt", "tee", "polo"]): return "tshirt"
  elif any(k in combined for k in ["kurti", "tunic"]): return "kurti"
  elif any(k in combined for k in ["kurta", "ethnic", "salwar"]): return "kurta"
  elif any(k in combined for k in ["dress", "frock", "gown", "jumpsuit"]): return "dress"
  elif any(k in combined for k in ["top", "blouse", "camisole"]): return "top"
  elif any(k in combined for k in ["shirt"]): return "shirt"
  elif any(k in combined for k in ["trouser", "chinos", "pant", "bottom", "legging", "palazzo"]): return "trousers"
  elif any(k in combined for k in ["jacket", "coat", "blazer", "windcheater", "shrug"]): return "jacket"
  elif any(k in combined for k in ["track", "jogger", "sweatpant", "pyjama", "pajama"]): return "joggers"
  elif any(k in combined for k in ["short", "boxer"]): return "shorts"
  elif any(k in combined for k in ["skirt"]): return "skirt"
  if "pant" in sheet_name.lower(): return "trousers"
  elif "shirt" in sheet_name.lower(): return "shirt"
  elif "coord" in sheet_name.lower() or "co-ord" in sheet_name.lower(): return "co-ord set"
  return "apparel"

def detect_gender(text):
  t = str(text).lower()
  if any(w in t for w in ["women", "woman", "female", "girl", "ladies", "kurti", "co-ord", "coord", "dress", "top", "skirt"]): return "Women"
  if any(w in t for w in ["men", "man", "male", "boy", "gents"]): return "Men"
  return "Women"

def resolve_category(clean_style, ctx, sheet_name=""):
  ov = st.session_state.get('category_overrides', {})
  if clean_style in ov and ov[clean_style]: return ov[clean_style]
  vr = ctx.get('vision_results') or {}
  if (vr.get(clean_style) or {}).get('category'): return vr[clean_style]['category']
  cat = detect_product_category(clean_style, ctx.get('keywords', ''), sheet_name)
  return ctx['default_category'] if (cat == 'apparel' and ctx.get('default_category')) else cat

def resolve_gender(clean_style, ctx):
  ov = st.session_state.get('gender_overrides', {})
  if clean_style in ov and ov[clean_style]: return ov[clean_style]
  vr = ctx.get('vision_results') or {}
  if (vr.get(clean_style) or {}).get('gender'): return vr[clean_style]['gender']
  return detect_gender(f"{clean_style} {ctx.get('keywords', '')}")

DEFAULT_CATEGORY_FABRICS = {'jeans': 'Denim', 'hoodie': 'Fleece', 'sweatshirt': 'Fleece', 'tshirt': 'Cotton', 'shirt': 'Cotton', 'cargo': 'Cotton', 'trousers': 'Cotton Blend', 'joggers': 'Cotton Blend', 'shorts': 'Cotton', 'co-ord set': 'Cotton Blend', 'kurti': 'Rayon', 'kurta': 'Cotton', 'dress': 'Crepe', 'top': 'Cotton Blend', 'jacket': 'Polyester', 'skirt': 'Cotton Blend', 'apparel': 'Cotton Blend'}

def resolve_fabric(clean_style, cat, ctx):
  vr = ctx.get('vision_results') or {}
  ai_mat = (vr.get(clean_style) or {}).get('material')
  if ai_mat and ai_mat.strip() and (ai_mat.strip().lower() != 'denim' or cat == 'jeans'): return ai_mat.strip().title()
  return "Denim" if cat == 'jeans' else DEFAULT_CATEGORY_FABRICS.get(cat, "Cotton Blend")

def resolve_sleeve_length(clean_style, cat, ctx):
  vr = ctx.get('vision_results') or {}
  ai_sleeve = (vr.get(clean_style) or {}).get('sleeve_length')
  if ai_sleeve and ai_sleeve.strip(): return ai_sleeve.strip().title()
  if cat in ['hoodie', 'sweatshirt', 'jacket']: return "Long Sleeve"
  elif cat in ['tshirt', 'shirt', 'kurti', 'dress', 'top']: return "Short Sleeve"
  return "Sleeveless"

COMMON_COLORS = ["Beige", "LightGrey", "Grey", "DarkGrey", "Black", "Navy", "Blue", "LightBlue", "DarkBlue", "SkyBlue", "IceBlue", "White", "Off White", "Green", "Olive", "Pista", "Sea Green", "Brown", "Coffee", "Charcoal", "Khaki", "Cream", "Tan", "Camel", "Mustard", "Rust", "Red", "Maroon", "Yellow", "Orange", "Pink", "Purple", "Lavender", "Wine", "Peach", "Teal"]
STANDARD_COLOR_MAPS = ['Beige', 'Black', 'Blue', 'Bronze', 'Brown', 'Clear', 'Gold', 'Green', 'Grey', 'Metallic', 'Multicolor', 'Off White', 'Orange', 'Pink', 'Purple', 'Red', 'Silver', 'Turquoise', 'White', 'Yellow']
COLOR_MISSPELLINGS = {'maron': 'Maroon', 'marron': 'Maroon', 'marun': 'Maroon', 'mahroon': 'Maroon', 'mehroon': 'Maroon', 'meroon': 'Maroon', 'mrn': 'Maroon', 'greay': 'Grey', 'gray': 'Grey', 'gery': 'Grey', 'grrey': 'Grey', 'gry': 'Grey', 'begie': 'Beige', 'baige': 'Beige', 'beigh': 'Beige', 'biege': 'Beige', 'bieg': 'Beige', 'beig': 'Beige', 'beg': 'Beige', 'badami': 'Beige', 'chiku': 'Beige', 'chikoo': 'Beige', 'blak': 'Black', 'blackk': 'Black', 'blk': 'Black', 'blck': 'Black', 'wite': 'White', 'whit': 'White', 'whiet': 'White', 'wht': 'White', 'offwhite': 'Off White', 'ofwhite': 'Off White', 'crm': 'Cream', 'blu': 'Blue', 'bleu': 'Blue', 'bule': 'Blue', 'bluu': 'Blue', 'lblu': 'LightBlue', 'dblu': 'DarkBlue', 'sblu': 'SkyBlue', 'iblu': 'IceBlue', 'navi': 'Navy', 'nvy': 'Navy', 'navyblue': 'Navy', 'chochlate': 'Brown', 'cofee': 'Coffee', 'brn': 'Brown', 'cfe': 'Coffee', 'yello': 'Yellow', 'ylw': 'Yellow', 'orenge': 'Orange', 'orang': 'Orange', 'orn': 'Orange', 'pinke': 'Pink', 'pnk': 'Pink', 'gren': 'Green', 'grn': 'Green', 'oliv': 'Olive', 'olv': 'Olive', 'khakhi': 'Khaki', 'khk': 'Khaki', 'lavendar': 'Lavender', 'lav': 'Lavender', 'prp': 'Purple', 'chr': 'Charcoal'}

def extract_explicit_color_from_text(text):
  if not text: return None
  text_clean = re.sub(r'[^a-zA-Z\s]', ' ', str(text)).lower()
  words = [w for w in text_clean.split() if w]
  joined = ''.join(words)
  for col in sorted(COMMON_COLORS, key=len, reverse=True):
    cl = col.lower()
    if (' ' in cl and cl in text_clean) or cl in words or cl.replace(' ', '') in words: return col
  for key in sorted(COLOR_MISSPELLINGS, key=len, reverse=True):
    if key in words: return COLOR_MISSPELLINGS[key]
  for col in sorted(COMMON_COLORS, key=len, reverse=True):
    cl_nospace = col.lower().replace(' ', '')
    if len(cl_nospace) >= 4 and cl_nospace in joined: return col
  for key in sorted(COLOR_MISSPELLINGS, key=len, reverse=True):
    if len(key) >= 4 and key in joined: return COLOR_MISSPELLINGS[key]
  color_lowers = [c.lower() for c in COMMON_COLORS]
  for w in words:
    if len(w) >= 4:
      close = difflib.get_close_matches(w, color_lowers, n=1, cutoff=0.82)
      if close:
        for col in COMMON_COLORS:
          if col.lower() == close[0]: return col
  return None

def analyze_image_dominant_color(image_bytes):
  if not image_bytes: return None
  try:
    img = Image.open(io.BytesIO(image_bytes)).convert('RGB')
    w, h = img.size
    cropped = img.crop((int(w * 0.25), int(h * 0.22), int(w * 0.75), int(h * 0.77))).resize((60, 60))
    pixels = list(cropped.getdata())
    corners = [img.getpixel((2, 2)), img.getpixel((w - 3, 2)), img.getpixel((2, h - 3)), img.getpixel((w - 3, h - 3))]
    bg_r, bg_g, bg_b = sum(c[0] for c in corners) / 4.0, sum(c[1] for c in corners) / 4.0, sum(c[2] for c in corners) / 4.0
    votes = {}
    for r, g, b in pixels:
      if abs(r - bg_r) + abs(g - bg_g) + abs(b - bg_b) < 35 and (r > 220 and g > 220 and b > 220): continue
      hue, sat, val = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)
      h_deg = hue * 360.0
      if 12 <= h_deg <= 38 and 0.22 <= sat <= 0.58 and 0.45 <= val <= 0.92 and r > g > b and (r - g) > 18: continue
      if val < 0.20: cname = "Black"
      elif sat < 0.12: cname = "White" if val > 0.85 else ("LightGrey" if val > 0.60 else ("Grey" if val > 0.35 else "Black"))
      elif 22 <= h_deg <= 52 and sat < 0.42 and val >= 0.62: cname = "Beige"
      elif 15 <= h_deg <= 45 and val < 0.62: cname = "Brown"
      elif 185 <= h_deg <= 255: cname = "Navy" if val < 0.38 else ("LightBlue" if (val > 0.78 and sat < 0.45) else "Blue")
      elif 68 <= h_deg < 165: cname = "Olive" if (h_deg < 105 and val < 0.65) else "Green"
      elif 165 <= h_deg < 185: cname = "Teal"
      elif 52 < h_deg < 68: cname = "Yellow"
      elif 15 <= h_deg < 22: cname = "Orange"
      elif 255 < h_deg <= 295: cname = "Lavender" if val > 0.75 else "Purple"
      elif h_deg > 295 or h_deg < 15: cname = "Maroon" if val < 0.48 else ("Red" if (sat > 0.62 and h_deg < 12) else "Pink")
      else: cname = "Black"
      votes[cname] = votes.get(cname, 0) + 1
    return max(votes.items(), key=lambda x: x[1])[0] if votes else None
  except Exception: return None

def detect_color(text, image_bytes=None):
  explicit = extract_explicit_color_from_text(text)
  if explicit: return explicit
  if image_bytes:
    px_col = analyze_image_dominant_color(image_bytes)
    if px_col: return px_col
  return "Black"

def _canonicalize_color(name):
  if not name: return name
  explicit = extract_explicit_color_from_text(name)
  if explicit: return explicit
  raw = str(name).strip().lower().replace('-', ' ').replace('_', ' ')
  words = raw.split()
  joined = ''.join(words)
  for key in sorted(COLOR_MISSPELLINGS, key=len, reverse=True):
    if key == joined or key in words: return COLOR_MISSPELLINGS[key]
  for w in words:
    if len(w) >= 4:
      close = difflib.get_close_matches(w, [c.lower() for c in COMMON_COLORS], n=1, cutoff=0.78)
      if close:
        for col in COMMON_COLORS:
          if col.lower() == close[0]: return col
  return name

def resolve_color(raw_fld, clean_style, ctx, image_bytes=None):
  explicit = extract_explicit_color_from_text(raw_fld) or extract_explicit_color_from_text(clean_style)
  if explicit: return explicit
  vr = ctx.get('vision_results') or {}
  ai_col = (vr.get(clean_style) or {}).get('color')
  if ai_col:
    ai_canon = _canonicalize_color(ai_col)
    if image_bytes and str(ai_canon).lower() in ('pink', 'lightpink', 'rose', 'peach'):
      px_col = analyze_image_dominant_color(image_bytes)
      if px_col and px_col.lower() not in ('pink', 'red', 'maroon', 'peach'): return px_col
    return ai_canon
  return detect_color(raw_fld, image_bytes=image_bytes)

COLOR_FAMILY_MAP = {'maroon': 'Red', 'mehroon': 'Red', 'red': 'Red', 'wine': 'Red', 'burgundy': 'Red', 'rust': 'Red', 'crimson': 'Red', 'cherry': 'Red', 'darkred': 'Red', 'navy': 'Blue', 'navyblue': 'Blue', 'blue': 'Blue', 'lightblue': 'Blue', 'darkblue': 'Blue', 'iceblue': 'Blue', 'skyblue': 'Blue', 'royalblue': 'Blue', 'teal': 'Blue', 'turquoise': 'Turquoise', 'black': 'Black', 'charcoal': 'Grey', 'grey': 'Grey', 'gray': 'Grey', 'lightgrey': 'Grey', 'darkgrey': 'Grey', 'smoke': 'Grey', 'silver': 'Silver', 'white': 'White', 'offwhite': 'Off White', 'cream': 'White', 'ivory': 'White', 'green': 'Green', 'olive': 'Green', 'mint': 'Green', 'darkgreen': 'Green', 'lightgreen': 'Green', 'parrotgreen': 'Green', 'pista': 'Green', 'seagreen': 'Green', 'brown': 'Brown', 'coffee': 'Brown', 'tan': 'Brown', 'chocolate': 'Brown', 'camel': 'Brown', 'darkbrown': 'Brown', 'lightbrown': 'Brown', 'beige': 'Beige', 'khaki': 'Beige', 'sand': 'Beige', 'yellow': 'Yellow', 'mustard': 'Yellow', 'orange': 'Orange', 'peach': 'Orange', 'pink': 'Pink', 'rose': 'Pink', 'lightpink': 'Pink', 'purple': 'Purple', 'lavender': 'Purple', 'violet': 'Purple', 'gold': 'Gold', 'bronze': 'Bronze', 'metallic': 'Metallic'}

def map_to_standard_color_map(det_col):
  if not det_col: return "Black"
  det_col = _canonicalize_color(det_col)
  raw = str(det_col).strip().lower().replace('-', ' ').replace('_', ' ')
  words, joined = raw.split(), ''.join(raw.split())
  for sc in STANDARD_COLOR_MAPS:
    if sc.lower().replace(' ', '') == joined: return sc
  for key in sorted(COLOR_FAMILY_MAP, key=len, reverse=True):
    if key == joined or key in words: return COLOR_FAMILY_MAP[key]
  return "Multicolor"

POCKET_BY_CATEGORY = {'hoodie': 'Kangaroo Pocket', 'sweatshirt': 'Kangaroo Pocket', 'cargo': 'Cargo Pocket', 'jeans': 'Five Pocket', 'trousers': 'Slash Pocket', 'pants': 'Slash Pocket', 'joggers': 'Side Pocket', 'shorts': 'Side Pocket', 'jacket': 'Welt Pocket', 'shirt': 'Chest Pocket', 'tshirt': 'No Pocket', 'kurta': 'Side Pocket', 'kurti': 'No Pocket', 'co-ord set': 'Side Pocket', 'dress': 'No Pocket', 'top': 'No Pocket', 'skirt': 'No Pocket', 'apparel': 'Side Pocket'}
def get_pocket_for_cat(cat): return POCKET_BY_CATEGORY.get(cat, 'Side Pocket')
def resolve_pocket(clean_style, cat, ctx):
  vr = ctx.get('vision_results') or {}
  return vr[clean_style]['pocket'] if (vr.get(clean_style) or {}).get('pocket') else get_pocket_for_cat(cat)

CATEGORY_LABELS = {'co-ord set': 'Co-ord Set', 'kurti': 'Kurti', 'kurta': 'Kurta', 'dress': 'Dress', 'top': 'Top', 'hoodie': 'Hoodie', 'sweatshirt': 'Sweatshirt', 'cargo': 'Cargo Pants', 'jeans': 'Jeans', 'tshirt': 'T-Shirt', 'shirt': 'Shirt', 'trousers': 'Trousers', 'jacket': 'Jacket', 'joggers': 'Joggers', 'shorts': 'Shorts', 'pants': 'Pants', 'skirt': 'Skirt', 'apparel': 'Casual Wear'}
HSN_CODE_MAP = {'co-ord set': '62042200', 'kurti': '62042200', 'dress': '62044200', 'top': '62063000', 'jeans': '62034290', 'cargo': '62034200', 'trousers': '62034200', 'pants': '62034200', 'joggers': '62034200', 'shorts': '62034200', 'tshirt': '61099090', 'shirt': '62052000', 'hoodie': '61103000', 'sweatshirt': '61103000', 'jacket': '62014090', 'kurta': '62034200', 'skirt': '62045200'}
def get_hsn_code(cat): return HSN_CODE_MAP.get(cat, '62034200')

def get_item_type_name_from_dropdown(cat):
  mapping = {'co-ord set': 'Outfit Set', 'kurti': 'Tunic', 'dress': 'Dress', 'top': 'Top', 'hoodie': 'Hooded Sweatshirt', 'sweatshirt': 'Sweatshirt', 'joggers': 'Sweatpants', 'jeans': 'Jeans', 'cargo': 'Cargo Pants', 'trousers': 'Casual Pants', 'pants': 'Casual Pants', 'shorts': 'Shorts', 'tshirt': 'T-Shirt', 'shirt': 'Casual Shirt', 'kurta': 'Kurta', 'jacket': 'Jacket', 'skirt': 'Skirt', 'apparel': 'Casual Pants'}
  return mapping.get(cat, 'Casual Pants')

def get_style_from_dropdown(cat):
  mapping = {'co-ord set': 'Co-ords', 'kurti': 'A-Line', 'dress': 'Fit & Flare', 'top': 'Regular', 'hoodie': 'Pullover', 'sweatshirt': 'Pullover', 'joggers': 'Jogger', 'cargo': 'Cargo', 'trousers': 'Chino', 'jeans': 'Classic'}
  return mapping.get(cat, 'Modern')

SC_PRODUCT_TYPE = {'co-ord set': 'SUIT', 'kurti': 'KURTA', 'dress': 'DRESS', 'top': 'SHIRT', 'hoodie': 'SWEATSHIRT', 'sweatshirt': 'SWEATSHIRT', 'jacket': 'COAT', 'jeans': 'PANTS', 'cargo': 'PANTS', 'trousers': 'PANTS', 'pants': 'PANTS', 'joggers': 'PANTS', 'shorts': 'SHORTS', 'tshirt': 'SHIRT', 'shirt': 'SHIRT', 'kurta': 'KURTA', 'skirt': 'SKIRT', 'apparel': 'SHIRT'}
SC_PRODUCT_TYPE_CANDIDATES = {'co-ord set': ['SUIT', 'APPAREL_SET', 'COORDINATED_OUTFIT', 'OUTFIT_SET', 'SALWAR_SUIT_SET', 'TRACK_SUIT', 'DRESS', 'KURTA', 'SHIRT', 'TOP'], 'kurti': ['KURTA', 'TUNIC', 'SHIRT', 'TOP', 'DRESS', 'SALWAR_SUIT_SET'], 'kurta': ['KURTA', 'TUNIC', 'SALWAR_SUIT_SET', 'SHIRT'], 'dress': ['DRESS', 'ONE_PIECE_OUTFIT', 'JUMPSUIT', 'OVERALLS', 'TUNIC', 'SUIT', 'GOWN'], 'top': ['SHIRT', 'TOP', 'TUNIC', 'BLOUSE', 'SWEATSHIRT', 'KURTA'], 'tshirt': ['SHIRT', 'T_SHIRT', 'TSHIRT', 'TOP', 'SWEATSHIRT'], 'shirt': ['SHIRT', 'TOP', 'TUNIC', 'KURTA'], 'hoodie': ['SWEATSHIRT', 'HOODIE', 'SWEATER', 'PULLOVER', 'COAT', 'JACKET', 'SHIRT'], 'sweatshirt': ['SWEATSHIRT', 'SWEATER', 'PULLOVER', 'HOODIE', 'COAT', 'SHIRT'], 'jacket': ['COAT', 'JACKET', 'BLAZER', 'OUTERWEAR', 'SWEATSHIRT', 'VEST'], 'jeans': ['PANTS', 'JEANS', 'TROUSERS', 'BOTTOMS', 'OVERALLS', 'SHORTS'], 'cargo': ['PANTS', 'TROUSERS', 'CARGO_PANTS', 'BOTTOMS', 'JEANS', 'SHORTS'], 'trousers': ['PANTS', 'TROUSERS', 'BOTTOMS', 'LEGGINGS', 'JEANS', 'SUIT'], 'pants': ['PANTS', 'TROUSERS', 'BOTTOMS', 'LEGGINGS', 'JEANS'], 'joggers': ['PANTS', 'SWEATPANTS', 'TRACK_PANTS', 'JOGGERS', 'TROUSERS', 'TRACK_SUIT', 'PAJAMAS', 'SHORTS'], 'shorts': ['SHORTS', 'PANTS', 'BOTTOMS', 'UNDERPANTS'], 'skirt': ['SKIRT', 'DRESS', 'PANTS', 'BOTTOMS'], 'apparel': ['SHIRT', 'PANTS', 'SWEATSHIRT', 'KURTA', 'DRESS', 'SUIT', 'COAT', 'SHORTS', 'TUNIC', 'TOP']}

def resolve_sc_product_type_for_row(ws_sc, sc_prod_type_col, raw_prod_type="", sheet_name="", item_type="", model_name="", item_name=""):
  options = get_dropdown_options_for_column(ws_sc, sc_prod_type_col) if sc_prod_type_col else None
  norm_opts = {}
  if options:
    for o in options:
      no = re.sub(r'[\s_\-]+', '', str(o).strip().upper())
      if no and no not in norm_opts: norm_opts[no] = o
  if raw_prod_type and str(raw_prod_type).strip():
    rp = str(raw_prod_type).strip()
    if not options: return rp
    nrp = re.sub(r'[\s_\-]+', '', rp.upper())
    if nrp in norm_opts: return norm_opts[nrp]
  cat = detect_product_category(f"{sheet_name} {item_type} {model_name}", str(item_name or ""), str(sheet_name or ""))
  candidates = list(SC_PRODUCT_TYPE_CANDIDATES.get(cat, SC_PRODUCT_TYPE_CANDIDATES['apparel']))
  default_pt = SC_PRODUCT_TYPE.get(cat, 'SHIRT')
  if default_pt not in candidates: candidates.insert(0, default_pt)
  sheet_clean = re.sub(r'[\s_\-]+', '', str(sheet_name or '').strip().upper())
  if options:
    if sheet_clean:
      if sheet_clean in norm_opts: return norm_opts[sheet_clean]
      if sheet_clean.rstrip('S') in norm_opts: return norm_opts[sheet_clean.rstrip('S')]
      if (sheet_clean + 'S') in norm_opts: return norm_opts[sheet_clean + 'S']
    for cand in candidates:
      nc = re.sub(r'[\s_\-]+', '', cand.upper())
      if nc in norm_opts: return norm_opts[nc]
      for no, orig_o in norm_opts.items():
        if nc in no or no in nc: return orig_o
    return pick_dropdown_value(ws_sc, sc_prod_type_col, default_pt, options[0])
  return default_pt

def make_workable_dropbox_url(base_url, sku, img_num):
  if not base_url: return f"https://www.dropbox.com/s/{sku}/{img_num}.jpg?raw=1"
  return f"{base_url.split('?')[0].rstrip('/')}/{sku}/{img_num}.jpg?raw=1"

def upload_to_imgbb(image_bytes, api_key):
  if not api_key: return None
  try:
    res = requests.post("https://api.imgbb.com/1/upload", data={"key": api_key, "image": base64.b64encode(image_bytes).decode('utf-8')}, timeout=20)
    if res.status_code == 200: return res.json()['data']['url']
  except Exception: pass
  return None

def upload_to_imgbb_cached(image_bytes, api_key, cache):
  h = hashlib.md5(image_bytes).hexdigest()
  if h in cache: return cache[h]
  url = upload_to_imgbb(image_bytes, api_key)
  if url: cache[h] = url
  return url

def _resolve_validation_formula(ws, f1):
  f1 = str(f1).strip()
  if f1.startswith('"') and f1.endswith('"'): return [o.strip() for o in f1[1:-1].split(',') if o.strip()]
  ref = f1[1:] if f1.startswith('=') else f1
  if '!' not in ref and ':' not in ref:
    dn = None
    try:
      dns = ws.parent.defined_names
      try: dn = dns.get(ref)
      except Exception: dn = dns[ref] if ref in dns else None
    except Exception: dn = None
    if dn is not None:
      ref = (getattr(dn, 'attr_text', '') or getattr(dn, 'value', '') or '').split(',')[0].strip().lstrip('=')
  if '!' in ref:
    try:
      sheet_part, rng = ref.rsplit('!', 1)
      target = ws.parent[sheet_part.strip("'")]
    except Exception: return []
  else: target, rng = ws, ref
  if ':' not in rng: return []
  try:
    vals = []
    for row in target[rng.replace('$', '')]:
      if not isinstance(row, tuple): row = (row,)
      for c in row:
        if c.value not in (None, ''): vals.append(str(c.value).strip())
    return vals
  except Exception: return []

def get_dropdown_options_for_column(ws, col_idx):
  try: dvs = ws.data_validations.dataValidation
  except Exception: dvs = []
  options = None
  for dv in dvs:
    hit = False
    try:
      for rng in dv.sqref.ranges:
        if rng.min_col <= col_idx <= rng.max_col: hit = True; break
    except Exception: hit = True
    if not hit: continue
    f1 = (dv.formula1 or '').strip()
    if not f1: continue
    opts = _resolve_validation_formula(ws, f1)
    if opts: options = opts; break
  if not options and hasattr(ws, 'parent') and ws.parent and 'Dropdown Lists' in ws.parent.sheetnames:
    try:
      ws_dd = ws.parent['Dropdown Lists']
      col_keys = {str(ws.cell(r, col_idx).value or '').strip().lower() for r in (1, 2, 3, 4, 5) if ws.cell(r, col_idx).value not in (None, '')}
      for dd_c in range(1, ws_dd.max_column + 1):
        header_hit_row = next((hr for hr in (1, 2, 3) if str(ws_dd.cell(hr, dd_c).value or '').strip().lower() in col_keys and str(ws_dd.cell(hr, dd_c).value or '').strip()), None)
        if header_hit_row is not None:
          vals = [str(ws_dd.cell(r, dd_c).value).strip() for r in range(header_hit_row + 1, ws_dd.max_row + 1) if ws_dd.cell(r, dd_c).value not in (None, '')]
          if vals: options = vals; break
    except Exception: pass
  if options:
    junk = {'', 'select', 'select one', 'choose', 'choose one', 'dropdown'}
    options = [o for o in options if str(o).strip().lower() not in junk]
    return options or None
  return None

def pick_dropdown_value(ws, col_idx, desired, fallback=None):
  options = get_dropdown_options_for_column(ws, col_idx)
  if not options: return desired if desired is not None else fallback
  cand = str(desired if desired is not None else (fallback or '')).strip()
  if not cand: return None
  cl = cand.lower()
  for o in options:
    if str(o).strip().lower() == cl: return o
  for o in options:
    ol = str(o).strip().lower()
    if cl in ol or ol in cl: return o
  close = difflib.get_close_matches(cand, options, n=1, cutoff=0.35)
  return close[0] if close else options[0]

KNOWN_POCKETS = ['Kangaroo Pocket', 'Cargo Pocket', 'Five Pocket', 'Side Pocket', 'No Pocket', 'Welt Pocket', 'Chest Pocket', 'Slash Pocket', 'Flap Pocket', 'Utility Pocket']
AI_CAT_MAP = {'co-ord': 'co-ord set', 'coord': 'co-ord set', 'cord set': 'co-ord set', 'two piece': 'co-ord set', 'hooded sweatshirt': 'hoodie', 'hooded': 'hoodie', 'hood': 'hoodie', 'pullover': 'hoodie', 'sweat shirt': 'sweatshirt', 'sweater': 'sweatshirt', 'crewneck': 'sweatshirt', 't-shirt': 'tshirt', 'tshirt': 'tshirt', 'tee': 'tshirt', 'polo': 'tshirt', 'denim': 'jeans', 'denim jeans': 'jeans', 'denim pants': 'jeans', 'pants': 'trousers', 'chinos': 'trousers', 'formal pants': 'trousers', 'track pants': 'joggers', 'trackpants': 'joggers', 'sweatpants': 'joggers', 'kurtis': 'kurti', 'kurtha': 'kurta', 'dresses': 'dress', 'tops': 'top'}

def guess_mime(filename):
  low = str(filename or '').lower()
  return 'image/png' if low.endswith('.png') else ('image/webp' if low.endswith('.webp') else 'image/jpeg')

def extract_json_from_text(text):
  if not text: return None
  try:
    m = re.search(r'\{[\s\S]*\}', text)
    if m: return json.loads(m.group(0))
  except Exception: pass
  return None

def get_vision_prompt(folder_hint=""):
  cats = ", ".join(CATEGORY_OPTIONS)
  hint_line = f'Folder/SKU name hint: "{folder_hint}". If the folder name mentions a color, prioritize that exact color.\n' if folder_hint else ""
  return (
    "You are an expert e-commerce fashion cataloger. Visually inspect the MAIN GARMENT worn by the model and output ONLY a valid JSON object:\n" + hint_line +
    '{"category": "<One of: ' + cats + '>", "gender": "<Women or Men>", "color": "<Black, Blue, Navy, LightBlue, DarkBlue, Beige, Grey, LightGrey, DarkGrey, White, Off White, Olive, Green, Pista, Brown, Coffee, Khaki, Cream, Maroon, Red, Yellow, Orange, Lavender, Purple, Teal, Wine, Peach, Pink>", "pocket": "<Kangaroo Pocket / Cargo Pocket / Five Pocket / Side Pocket / Welt Pocket / Chest Pocket / Slash Pocket / No Pocket>", "material": "<Cotton, Cotton Blend, Polycotton, Polyester, Fleece, Rayon, Denim, Crepe, Georgette, Silk, Linen, Knit>", "fabric_detail": "<100% Cotton, Cotton Blend, Polycotton, Soft Fleece, Polyester, Rayon>", "sleeve_length": "<Full Sleeve / Long Sleeve / Short Sleeve / 3/4 Sleeve / Sleeveless>", "pattern": "<Solid / Printed / Striped / Checked / Embroidered / Washed / Floral / Self Design>", "print_type": "<Solid / Botanical / Floral / Striped / Checked / Embroidered>", "neck_collar": "<Cuban Collar / Slim Collar / Mandarin / Spread Collar / Round Neck / Polo Neck / Hood>"}'
  )

def call_gemini_vision(image_bytes, img_name, api_key, model, folder_hint=""):
  if not api_key or not image_bytes: return None
  b64 = base64.b64encode(image_bytes).decode('utf-8')
  payload = {"contents": [{"parts": [{"text": get_vision_prompt(folder_hint)}, {"inline_data": {"mime_type": guess_mime(img_name), "data": b64}}]}], "generationConfig": {"temperature": 0.0, "maxOutputTokens": 1500}}
  try:
    res = requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}", json=payload, timeout=40)
    if res.status_code == 200:
      cands = res.json().get('candidates') or []
      if cands:
        parts = (cands[0].get('content') or {}).get('parts') or []
        return extract_json_from_text(''.join(str(p.get('text', '')) for p in parts))
  except Exception: pass
  return None

def call_nvidia_vision(image_bytes, img_name, api_key, model, folder_hint=""):
  if not api_key or not image_bytes: return None
  data_url = f"data:{guess_mime(img_name)};base64,{base64.b64encode(image_bytes).decode('utf-8')}"
  try:
    res = requests.post("https://integrate.api.nvidia.com/v1/chat/completions", headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}, json={"model": model, "messages": [{"role": "user", "content": [{"type": "text", "text": get_vision_prompt(folder_hint)}, {"type": "image_url", "image_url": {"url": data_url}}]}], "temperature": 0.1, "max_tokens": 1500}, timeout=50)
    if res.status_code == 200: return extract_json_from_text(res.json()['choices'][0]['message']['content'])
  except Exception: pass
  return None

def call_groq_vision(image_bytes, img_name, api_key, model, folder_hint=""):
  if not api_key or not image_bytes: return None
  data_url = f"data:{guess_mime(img_name)};base64,{base64.b64encode(image_bytes).decode('utf-8')}"
  try:
    res = requests.post("https://api.groq.com/openai/v1/chat/completions", headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}, json={"model": model, "messages": [{"role": "user", "content": [{"type": "text", "text": get_vision_prompt(folder_hint)}, {"type": "image_url", "image_url": {"url": data_url}}]}], "temperature": 0.1, "max_tokens": 1500}, timeout=40)
    if res.status_code == 200: return extract_json_from_text(res.json()['choices'][0]['message']['content'])
  except Exception: pass
  return None

def multi_ai_vision_detect(image_bytes, img_name, folder_hint=""):
  for fn, k, m, p in [(call_gemini_vision, GEMINI_API_KEY, GEMINI_VISION_MODEL, "Gemini"), (call_nvidia_vision, NVIDIA_API_KEY, NVIDIA_MODEL, "NVIDIA"), (call_groq_vision, GROQ_API_KEY, GROQ_VISION_MODEL, "Groq")]:
    res = fn(image_bytes, img_name, k, m, folder_hint=folder_hint)
    if res and isinstance(res, dict) and res.get('category'): return res, p
  return None, "None"

def canonicalize_ai_category(cat_str):
  cat = AI_CAT_MAP.get(str(cat_str or '').strip().lower().replace('-', ' ').strip(), str(cat_str or '').strip().lower().replace('-', ' ').strip())
  if cat in CATEGORY_OPTIONS: return cat
  close = difflib.get_close_matches(cat, CATEGORY_OPTIONS, n=1, cutoff=0.6)
  return close[0] if close else 'apparel'

def canonicalize_ai_color(col_str, folder_name="", image_bytes=None):
  explicit_fld = extract_explicit_color_from_text(folder_name)
  if explicit_fld: return explicit_fld
  col = str(col_str or '').strip()
  if col:
    canon = str(_canonicalize_color(col))
    joined = canon.lower().replace(' ', '')
    matched_col = next((c for c in COMMON_COLORS if c.lower().replace(' ', '') == joined), None) or ' '.join(w.capitalize() for w in re.split(r'[\s\-]+', canon) if w)
    if image_bytes and matched_col.lower() in ('pink', 'lightpink', 'rose', 'peach'):
      px_col = analyze_image_dominant_color(image_bytes)
      if px_col and px_col.lower() not in ('pink', 'red', 'maroon', 'peach'): return px_col
    return matched_col
  if image_bytes:
    px_col = analyze_image_dominant_color(image_bytes)
    if px_col: return px_col
  return "Black"

def canonicalize_ai_pocket(p_str):
  p = str(p_str or '').strip()
  if not p: return None
  for k in KNOWN_POCKETS:
    if k.lower() == p.lower(): return k
  close = difflib.get_close_matches(p, KNOWN_POCKETS, n=1, cutoff=0.5)
  return close[0] if close else p.title()

def run_vision_detection(zip_images_data):
  results, failed = {}, []
  folders = [f for f in zip_images_data.keys() if zip_images_data.get(f)]
  if not folders: return results, failed
  progress, status = st.progress(0.0), st.empty()
  for i, fld in enumerate(folders):
    cs = clean_sku_name(fld)
    img_name, img_bytes = zip_images_data[fld][0]
    status.info(f"🤖 AI analyzing garment image: **{cs}** ({i+1}/{len(folders)})...")
    data, provider = multi_ai_vision_detect(img_bytes, img_name, folder_hint=fld)
    if data:
      cat = canonicalize_ai_category(data.get('category'))
      col = canonicalize_ai_color(data.get('color'), folder_name=fld, image_bytes=img_bytes)
      poc = canonicalize_ai_pocket(data.get('pocket'))
      gen = "Women" if str(data.get('gender')).lower().startswith('w') else "Men"
      mat = str(data.get('material') or '').strip().title()
      if not mat or (mat.lower() == 'denim' and cat != 'jeans'): mat = DEFAULT_CATEGORY_FABRICS.get(cat, 'Cotton Blend')
      fab_det = str(data.get('fabric_detail') or f"100% {mat}").title()
      sleeve = str(data.get('sleeve_length') or ('Long Sleeve' if cat in ['hoodie', 'sweatshirt'] else 'Short Sleeve')).title()
      pattern = str(data.get('pattern') or 'Solid').title()
      results[cs] = {'category': cat, 'color': col, 'pocket': poc, 'gender': gen, 'material': mat, 'fabric_detail': fab_det, 'sleeve_length': sleeve, 'pattern': pattern, 'print_type': str(data.get('print_type') or pattern).title(), 'neck_collar': str(data.get('neck_collar') or '').title(), 'ai_source': provider}
    else:
      results[cs] = {'color': detect_color(fld, image_bytes=img_bytes), 'ai_source': 'Pixel/Folder-Fallback'}
      failed.append(cs)
    time.sleep(0.3)
    progress.progress(min((i + 1) / len(folders), 1.0))
  progress.empty(); status.empty()
  return results, failed

def _limit_desc_200(text):
  s = re.sub(r'\s+', ' ', str(text or '')).strip()
  if len(s) <= DESC_MAX_LEN: return s
  cut = s[:DESC_MAX_LEN].rsplit(' ', 1)[0].rstrip('.,;:- ')
  return (cut + '.') if len(cut) < DESC_MAX_LEN else s[:DESC_MAX_LEN]

def _user_kw_list(user_kw): return [k.strip() for k in user_kw.replace(';', ',').split(',') if k.strip()] if user_kw else []

def build_vendor_title(brand, category, full_sku, user_kw="", limit=75, gender="Women"):
  b = f"{brand} " if brand else ""
  cat_str, sku_part = CATEGORY_LABELS.get(category, category.title()), f"({full_sku})"
  kws = _user_kw_list(user_kw)
  kw_candidate = kws[0] if kws else ""
  target_str = "for women" if gender == "Women" else "for men"
  if kw_candidate:
    opt1 = f"{b}{cat_str} {target_str} | {kw_candidate} {sku_part}".strip()
    if len(opt1) <= limit: title = opt1
    else:
      opt2 = f"{b}{kw_candidate} {sku_part}".strip()
      title = opt2 if len(opt2) <= limit else f"{b}{cat_str} {sku_part}".strip()
  else: title = f"{b}{cat_str} {target_str} | casual wear {sku_part}".strip()
  return title[:limit].rstrip()

def build_sc_title(brand, cat, full_sku, user_kw="", limit=150, gender="Women"):
  b = f"{brand} " if brand else ""
  label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Casual Wear")
  kws = _user_kw_list(user_kw)
  base = f"{b}{'Women\'s' if gender == 'Women' else 'Men\'s'} {label} for Casual Wear"
  for n in range(min(3, len(kws)), -1, -1):
    t = f"{base} | {', '.join(kws[:n])} ({full_sku})" if n else f"{base} ({full_sku})"
    if len(t) <= limit: return t
  return t[:limit]

def generate_generic_keywords_200(cat, det_col, user_kw, gender="Women"):
  if user_kw:
    kws, unique = _user_kw_list(user_kw), []
    for k in kws:
      if k.lower() not in [x.lower() for x in unique]: unique.append(k)
    return "; ".join(unique)[:200]
  g_str = "women" if gender == "Women" else "men"
  cat_label = CATEGORY_LABELS.get(cat, cat)
  return "; ".join([f"{g_str} {cat_label}", f"{det_col} {cat_label}", f"casual wear for {g_str}", "comfortable fit", "daily wear", "stylish modern wear", "breathable fabric"])[:200]

def build_freehand_bullets_and_desc(cat, brand, gender="Women", det_mat="Cotton Blend"):
  b_prefix = f"{brand} " if brand else ""
  cat_label = CATEGORY_LABELS.get(cat, str(cat).title() if cat else "Casual Wear")
  g_str = "Women's" if gender == "Women" else "Men's"
  if cat == 'co-ord set':
    bullets = [f"PREMIUM SOFT FABRIC: Crafted from high-grade, breathable {det_mat} ensuring all-day comfort and a soft skin feel.", "ELEGANT 2-PIECE CO-ORD SET: Features a coordinated matching silhouette designed for effortlessly chic casual styling.", "VERSATILE EVERYDAY WEAR: Perfect for outings, travel, lounge sessions, brunches, and relaxed casual gatherings.", "COMFORTABLE ERGONOMIC FIT: Tailored cut provides ease of movement and all-day flattering drape.", "WASH & CARE: Easy gentle machine wash cold with similar colors. Do not bleach. Warm iron if required."]
    desc = f"Upgrade your wardrobe with this stylish {b_prefix}{g_str} {cat_label}. Crafted from breathable {det_mat} for modern style and all-day comfort."
  elif cat in ['hoodie', 'sweatshirt']:
    bullets = [f"PREMIUM SOFT FLEECE FABRIC: Made from high-grade, breathable {det_mat} that keeps you warm, cozy, and comfortable throughout the day.", "URBAN STREETWEAR DESIGN: Features an on-trend silhouette with a functional hood and durable drawstrings.", "SPACIOUS KANGAROO POCKET: Front pocket provides convenient storage for daily essentials while keeping hands warm.", "VERSATILE EVERYDAY WEAR: Effortlessly pairs with jeans, cargo pants, joggers, and sneakers for casual outings and winter styling.", "WASH & CARE: Easy machine wash cold with similar colors inside out. Do not bleach. Warm iron if required."]
    desc = f"Upgrade your wardrobe with this stylish {b_prefix}{g_str} {cat_label}. Crafted from soft {det_mat} for modern style and cozy daily comfort."
  elif cat == 'cargo':
    bullets = [f"DURABLE UTILITY FABRIC: Tailored from rugged, breathable {det_mat} designed for active mobility and long-lasting durability.", "MULTI-POCKET FUNCTIONALITY: Equipped with deep tactical flap cargo pockets and side pockets for maximum utility.", "COMFORTABLE RELAXED FIT: Features an ergonomic waistband and relaxed leg cut for easy everyday movement.", "IDEAL FOR CASUAL & OUTDOOR WEAR: Perfect for travel, hangouts, and pairing with oversized tees, hoodies, or boots.", "WASH & CARE: Machine wash cold inside out. Tumble dry low. Do not bleach. Warm iron if needed."]
    desc = f"Experience comfort and utility with {b_prefix}{g_str} {cat_label}. Made from durable {det_mat} with multi-pocket styling for everyday wear."
  elif cat == 'jeans':
    bullets = [f"COMFORT STRETCH DENIM: Crafted from premium stretchable {det_mat} denim for flexible mobility and all-day ease.", "AUTHENTIC WASHED FINISH: Features a contemporary wash with subtle fading and durable contrast stitching.", "CLASSIC 5-POCKET DESIGN: Practical 5-pocket layout with sturdy rivets, heavy-duty zip fly, and button closure.", "TIMELESS RELAXED CUT: Sits comfortably at the waist with an authentic silhouette suitable for daily college and office casuals.", "WASH & CARE: Machine wash cold inside out with similar colors. Do not bleach. Warm iron if required."]
    desc = f"Step out in style with {b_prefix}{g_str} {cat_label}. Crafted from stretchable {det_mat} for all-day flexibility and a classic casual look."
  else:
    bullets = [f"HIGH-QUALITY COMFORT FABRIC: Crafted from premium {det_mat} engineered for soft skin feel and long-lasting daily wear.", "CONTEMPORARY STYLING: Features an on-trend silhouette with neat, durable seam construction and a comfortable fit.", "FUNCTIONAL & ERGONOMIC DESIGN: Tailored cut provides comfortable mobility and effortless wearability throughout the day.", "VERSATILE CASUAL PAIRING: Complements casual sneakers and daily wardrobe essentials for an easy, elevated style statement.", "WASH & CARE: Machine wash cold with like colors inside out. Do not bleach. Warm iron if required."]
    desc = f"Upgrade your daily rotation with {b_prefix}{g_str} {cat_label}. Crafted from premium {det_mat} for lasting durability and everyday comfort."
  return bullets, _limit_desc_200(desc)

def build_ajio_seo_content(cat, user_keywords="", gender="Women", mat_val="Cotton Blend"):
  kw_str = user_keywords if user_keywords else "relaxed fit, all-day comfort"
  cat_label, g_title = CATEGORY_LABELS.get(cat, 'Apparel'), "Women's" if gender == "Women" else "Men's"
  if cat == 'co-ord set':
    desc = f"{g_title} Co-ord Set in breathable {mat_val} fabric. Features a matching top and bottom with {kw_str} for chic daily and travel wear."
    highlight = f"Premium {mat_val} fabric | Comfortable {g_title.lower()} fit | Elegant matching 2-piece set | Relaxed styling | Easy machine wash"
    add_info = "Styling Tip: Pair with minimal jewelry and slip-ons. Wash Care: Machine wash cold with similar colors."
  elif cat == 'cargo':
    desc = f"{g_title} Cargo Pants in durable {mat_val} fabric. Features utility flap pockets, flexi waist, and {kw_str} for casual outdoor wear."
    highlight = f"Premium breathable {mat_val} fabric | Comfortable flexi waist | Multi-pocket functional cargo utility styling | Relaxed fit | Easy machine wash"
    add_info = "Styling Tip: Pair with oversized t-shirts, hoodies, and chunky sneakers. Wash Care: Machine wash cold with similar colors."
  elif cat == 'jeans':
    desc = f"Trendy {g_title} Jeans in stretch denim with {kw_str}. Features a mid-rise relaxed fit, 5-pocket styling, and clean washed finish."
    highlight = "Premium stretch denim | Relaxed mid-rise silhouette | Authentic clean washed finish | Multi-pocket utility styling | All-season wear"
    add_info = "Styling Tip: Pair with casual shirts or boxy tees. Wash Care: Machine wash cold inside out."
  elif cat in ['hoodie', 'sweatshirt']:
    desc = f"Warm {g_title} {cat_label} in soft {mat_val} fabric with {kw_str}. Features hood, kangaroo pocket, and ribbed cuffs for streetwear style."
    highlight = f"Premium soft {mat_val} fabric | Functional hood with drawstrings | Spacious kangaroo pocket | Ribbed cuffs | Machine washable"
    add_info = "Styling Tip: Layer over tees with jeans or joggers. Wash Care: Machine wash cold inside out."
  else:
    desc = f"{g_title} {cat_label} crafted from breathable {mat_val} fabric featuring {kw_str} for lasting comfort, durability, and everyday style."
    highlight = f"Premium {mat_val} fabric | Comfortable {g_title.lower()} fit | Durable stitching | Easy machine wash"
    add_info = "Wash Care: Machine wash cold with like colors. Do not bleach."
  return _limit_desc_200(desc), highlight[:500], add_info[:500]

def ungroup_worksheet(ws):
  for col_dim in ws.column_dimensions.values(): col_dim.outline_level = 0; col_dim.collapsed = False; col_dim.hidden = False
  for r_dim in ws.row_dimensions.values(): r_dim.outline_level = 0; r_dim.collapsed = False; r_dim.hidden = False
  if hasattr(ws, 'sheet_properties') and hasattr(ws.sheet_properties, 'outlinePr'):
    ws.sheet_properties.outlinePr.summaryBelow = True; ws.sheet_properties.outlinePr.summaryRight = True

def load_workbook_clean(uploaded_file_or_bytes, name=""):
  wb = openpyxl.load_workbook(uploaded_file_or_bytes, keep_vba=str(name).lower().endswith('.xlsm'))
  for s in wb.sheetnames:
    wb[s].protection.sheet = False
    try: wb[s].protection.disable()
    except Exception: pass
  return wb

def pick_template_sheet(wb):
  for sname in wb.sheetnames:
    if sname.strip().lower() == 'template': return wb[sname]
  for sname in wb.sheetnames:
    if sname.strip().lower().startswith('template'): return wb[sname]
  for sname in wb.sheetnames:
    if 'template' in sname.lower() and 'change' not in sname.lower(): return wb[sname]
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
# AMAZON VENDOR CENTRAL ENGINE
# ======================================================================
def run_vendor_central(uploaded_file, ctx):
  wb = load_workbook_clean(uploaded_file, uploaded_file.name)
  ws = pick_template_sheet(wb)
  ungroup_worksheet(ws)
  col_by_label, col_by_key = {}, {}
  for c in range(1, ws.max_column + 1):
    lbl = str(ws.cell(3, c).value or '').strip().lower()
    key = str(ws.cell(4, c).value or '').strip().lower()
    if lbl and lbl not in col_by_label: col_by_label[lbl] = c
    if key and key not in col_by_key: col_by_key[key] = c

  def get_vc_col(labels, keys=()):
    for k in keys:
      if k.lower() in col_by_key: return col_by_key[k.lower()]
    for l in labels:
      if l.lower() in col_by_label: return col_by_label[l.lower()]
    return None

  curr_r = 7
  sku_col_idx = get_vc_col(['vendor sku'], ['vendor_sku#1.value']) or 2
  while ws.cell(curr_r, sku_col_idx).value is not None: curr_r += 1

  today_str, today_shipping = datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%Y%m%d")
  mrp_fill, mrp_font = HIGHLIGHT_FILL, HIGHLIGHT_FONT
  material_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
  added, preview = 0, []

  def vset(col, value, dropdown=False, highlight=None, bold=False):
    if not col: return
    cell = ws.cell(curr_r, col)
    if value not in (None, ''): cell.value = pick_dropdown_value(ws, col, value, value) if dropdown else value
    if highlight: cell.fill = highlight
    if bold: cell.font = mrp_font

  for raw_fld in ctx['folders']:
    clean_style = clean_sku_name(raw_fld)
    det_col = resolve_color(raw_fld, clean_style, ctx)
    cat = resolve_category(clean_style, ctx, ws.title)
    det_gen = resolve_gender(clean_style, ctx)
    cat_label = CATEGORY_LABELS.get(cat, cat.title() if cat else "Casual Wear")
    sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)
    mat_val, sleeve_val = resolve_fabric(clean_style, cat, ctx), resolve_sleeve_length(clean_style, cat, ctx)
    bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
    gen_keywords_200 = generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)
    style_val, item_type_val = get_style_from_dropdown(cat), get_item_type_name_from_dropdown(cat)
    color_map_val, pocket_val = map_to_standard_color_map(det_col), resolve_pocket(clean_style, cat, ctx)
    dept_val, gender_val = ("Womens", "Female") if det_gen == "Women" else ("Mens", "Male")

    preview.append({'SKU': clean_style, 'Category': cat, 'Gender': det_gen, 'Color': det_col, 'Color Map': color_map_val, 'Item Type': item_type_val, 'Style': style_val, 'Pocket': pocket_val, 'Material': mat_val})

    for sz in ctx['sizes']:
      clean_sz = clean_sku_name(sz)
      full_sku = f"{clean_style}-{clean_sz}"
      is_num_sz = clean_sz.isdigit()
      vendor_title = build_vendor_title(ctx['brand'], cat, full_sku, ctx['keywords'], 75, det_gen)

      vset(get_vc_col(['vendor code'], ['rtip_vendor_code#1.value']), "J B fashion -- Surat -- Dropship, in_pd-ravi_apparel, JBZ54")
      vset(get_vc_col(['vendor sku'], ['vendor_sku#1.value']), full_sku)
      for p_lbl, p_key in [(['parentage level'], ['parentage_level#1.value']), (['child relationship type'], ['child_parent_sku_relationship#1.child_relationship_type']), (['parent sku'], ['child_parent_sku_relationship#1.parent_sku']), (['variation theme name'], ['variation_theme#1.name'])]:
        cp = get_vc_col(p_lbl, p_key)
        if cp: ws.cell(curr_r, cp).value = None
      vset(get_vc_col(['item name'], ['item_name#1.value']), vendor_title)
      vset(get_vc_col(['item highlight'], ['title_differentiation#1.value']), f"Premium quality {cat_label} with comfortable durable fit")
      vset(get_vc_col(['brand name'], ['brand#1.value']), ctx['brand'] or None)
      vset(get_vc_col(['model number'], ['model_number#1.value']), full_sku)
      vset(get_vc_col(['part number'], ['part_number#1.value']), full_sku)
      vset(get_vc_col(['model name'], ['model_name#1.value']), f"{cat.capitalize()} {clean_style}", highlight=HIGHLIGHT_FILL, bold=True)
      vset(get_vc_col(['manufacturer'], ['manufacturer#1.value']), ctx['mfg'] or None)
      vset(get_vc_col(['manufacturer contact information'], ['rtip_manufacturer_contact_information#1.value']), ctx['mfg'] or None)

      c_main_img = get_vc_col(['main image url'], ['main_product_image_locator#1.media_location'])
      if c_main_img and sku_img_urls: ws.cell(curr_r, c_main_img).value = sku_img_urls[0]
      for img_num in range(1, 9):
        c_oth = get_vc_col(['other image url'], [f'other_product_image_locator_{img_num}#1.media_location'])
        if c_oth and img_num < len(sku_img_urls): ws.cell(curr_r, c_oth).value = sku_img_urls[img_num]
      for b_i, b_val in enumerate(bullets_list[:5]):
        vset(get_vc_col(['bullet point'], [f'bullet_point#{b_i+1}.value']), b_val)

      vset(get_vc_col(['generic keywords'], ['generic_keyword#1.value']), gen_keywords_200)
      vset(get_vc_col(['style'], ['style#1.value']), style_val, dropdown=True)
      for ls_idx, ls_val in enumerate(["Casual", "Work Utility", "Comfort", "Athletic", "Formal"], 1):
        vset(get_vc_col(['lifestyle'], [f'lifestyle#{ls_idx}.value']), ls_val, dropdown=True)
      vset(get_vc_col(['department name'], ['department#1.value']), dept_val, dropdown=True)
      vset(get_vc_col(['target gender'], ['target_gender#1.value']), gender_val, dropdown=True)
      vset(get_vc_col(['age range description'], ['age_range_description#1.value']), "Adult", dropdown=True)
      vset(get_vc_col(['apparel size system', 'bottoms size system', 'size system'], ['bottoms_size#1.size_system', 'apparel_size#1.size_system']), "IN", dropdown=True)
      vset(get_vc_col(['apparel size class', 'bottoms size class', 'size class'], ['bottoms_size#1.size_class', 'apparel_size#1.size_class']), "Numeric" if is_num_sz else "Alpha", dropdown=True)
      vset(get_vc_col(['apparel size value', 'bottoms size value', 'size value'], ['bottoms_size#1.size', 'apparel_size#1.size']), clean_sz)
      vset(get_vc_col(['material'], ['material#1.value']), mat_val, dropdown=True, highlight=material_fill)
      vset(get_vc_col(['fabric type'], ['fabric_type#1.value']), f"100% {mat_val}")
      vset(get_vc_col(['number of items'], ['number_of_items#1.value']), 1)
      vset(get_vc_col(['item type name'], ['item_type_name#1.value']), item_type_val, dropdown=True)
      vset(get_vc_col(['special size'], ['special_size_type#1.value']), "Standard", dropdown=True)
      vset(get_vc_col(['product description'], ['rtip_product_description#1.value']), _limit_desc_200(desc_text))
      vset(get_vc_col(['color map'], ['color#1.standardized_values#1']), color_map_val, dropdown=True)
      vset(get_vc_col(['color'], ['color#1.value']), det_col)
      c_thm = get_vc_col(['theme'], ['theme#1.value'])
      if c_thm: ws.cell(curr_r, c_thm).value = None
      vset(get_vc_col(['fit type'], ['fit_type#1.value']), "Relaxed", dropdown=True)
      vset(get_vc_col(['pocket description'], ['pocket_description#1.value']), pocket_val, dropdown=True)
      vset(get_vc_col(['care instructions'], ['care_instructions#1.value']), "Machine Wash", dropdown=True)
      vset(get_vc_col(['pattern'], ['pattern#1.value']), "Solid", dropdown=True)
      vset(get_vc_col(['unit count'], ['unit_count#1.value']), 1)
      vset(get_vc_col(['unit count type'], ['unit_count#1.type.value']), "count", dropdown=True)
      vset(get_vc_col(['product site launch date'], ['product_site_launch_date#1.value']), today_str)
      vset(get_vc_col(['external product information entity'], ['external_product_information#1.entity']), "HSN Code", dropdown=True)
      vset(get_vc_col(['external product information'], ['external_product_information#1.value']), get_hsn_code(cat))
      if is_num_sz:
        vset(get_vc_col(['waist size'], ['waist#1.size#1.value']), clean_sz)
        vset(get_vc_col(['waist size unit'], ['waist#1.size#1.unit']), "inches", dropdown=True)
      vset(get_vc_col(['product lifecycle supply type'], ['lifecycle_supply_type#1.value']), "Year Round Replenishable", dropdown=True)
      vset(get_vc_col(['seasons'], ['seasons#1.value']), "All", dropdown=True)
      vset(get_vc_col(['sleeve length description', 'sleeve length']), sleeve_val, dropdown=True)
      vset(get_vc_col(['sleeve type']), "Regular Sleeve", dropdown=True)
      vset(get_vc_col(['packer contact information'], ['packer_contact_information#1.value']), ctx['mfg'] or None)
      vset(get_vc_col(['item weight'], ['item_weight#1.value']), 350)
      vset(get_vc_col(['item weight unit'], ['item_weight#1.unit']), "Grams", dropdown=True)
      c_mrp = get_vc_col(['maximum retail price (mrp)'], ['list_price#1.value_with_tax'])
      if c_mrp:
        if ctx['mrp']: ws.cell(curr_r, c_mrp).value = int(ctx['mrp']) if str(ctx['mrp']).isdigit() else ctx['mrp']
        else: ws.cell(curr_r, c_mrp).fill = mrp_fill; ws.cell(curr_r, c_mrp).font = mrp_font
      vset(get_vc_col(['earliest shipping date'], ['rtip_earliest_shipping_date#1.value']), today_shipping)
      for ld, kd, vd in [(['item length'], ['item_dimensions#1.length.value'], 10), (['item width'], ['item_dimensions#1.width.value'], 2), (['item height'], ['item_dimensions#1.height.value'], 1.3), (['item package length'], ['item_package_dimensions#1.length.value'], 10), (['item package width'], ['item_package_dimensions#1.width.value'], 2), (['item package height'], ['item_package_dimensions#1.height.value'], 1.3)]:
        vset(get_vc_col(ld, kd), vd)
      for lu, ku in [(['item length unit'], ['item_dimensions#1.length.unit']), (['item width unit'], ['item_dimensions#1.width.unit']), (['item height unit'], ['item_dimensions#1.height.unit']), (['package length unit'], ['item_package_dimensions#1.length.unit']), (['package width unit'], ['item_package_dimensions#1.width.unit']), (['package height unit'], ['item_package_dimensions#1.height.unit'])]:
        vset(get_vc_col(lu, ku), "inches", dropdown=True)
      vset(get_vc_col(['package weight'], ['item_package_weight#1.value']), 380)
      vset(get_vc_col(['package weight unit'], ['item_package_weight#1.unit']), "Grams", dropdown=True)
      vset(get_vc_col(['order aggregate type'], ['rtip_order_aggregate_type#1.value']), "Each", dropdown=True)
      vset(get_vc_col(['items per inner pack'], ['rtip_items_per_inner_pack#1.value']), 1)
      vset(get_vc_col(['country of origin'], ['country_of_origin#1.value']), "India", dropdown=True)
      vset(get_vc_col(['dangerous goods regulations', 'dangerous goods', 'hazmat']), "Not Applicable", dropdown=True)
      curr_r += 1; added += 1

  output = io.BytesIO()
  wb.save(output); output.seek(0)
  return output.getvalue(), preview, added

# ======================================================================
# AMAZON SELLER CENTRAL ENGINE
# ======================================================================
def _norm_amazon_key(k):
  if not k: return ""
  s = re.sub(r'\[(marketplace_id|language_tag|content_language|content_type)=[^\]]*\]', '', str(k).strip().lower()).replace('[audience=all]', '')
  return s[5:] if s.startswith('rtip_') else s

VC_TO_SC_KEY_ALIASES = {'vendor_sku#1.value': 'contribution_sku#1.value', 'external_product_id#1.type': 'amzn1.volt.ca.product_id_type', 'external_product_id#1.value': 'amzn1.volt.ca.product_id_value', 'merchant_suggested_asin#1.value': 'amzn1.volt.ca.product_id_value', 'list_price#1.value_with_tax': 'purchasable_offer#1.maximum_retail_price#1.schedule#1.value_with_tax'}
VC_TO_SC_KEY_FALLBACK_GROUPS = [['apparel_size#1.size_system', 'bottoms_size#1.size_system', 'shirt_size#1.size_system', 'shapewear_size#1.size_system'], ['apparel_size#1.size_class', 'bottoms_size#1.size_class', 'shirt_size#1.size_class', 'shapewear_size#1.size_class'], ['apparel_size#1.size', 'bottoms_size#1.size', 'shirt_size#1.size', 'shapewear_size#1.size'], ['apparel_size#1.size_to', 'bottoms_size#1.size_to', 'shirt_size#1.size_to', 'shapewear_size#1.size_to'], ['apparel_size#1.body_type', 'bottoms_size#1.body_type', 'shirt_size#1.body_type', 'shapewear_size#1.body_type'], ['apparel_size#1.height_type', 'bottoms_size#1.height_type', 'shirt_size#1.height_type', 'shapewear_size#1.height_type']]
VC_TO_SC_LABEL_ALIASES = {'vendor sku': 'sku', 'maximum retail price (mrp)': 'maximum retail price (sell on amazon, in)', 'external product id type': 'product id type', 'external product id': 'product id', 'merchant suggested asin': 'product id'}
PARENT_INHERITED_LABELS = {'product type', 'brand name', 'product description', 'bullet point', 'generic keywords', 'material', 'fabric type', 'fit type', 'pocket description', 'care instructions', 'manufacturer contact information', 'pattern', 'external product information entity', 'external product information', 'seasons', 'packer contact information', 'sleeve length description', 'sleeve type', 'item weight', 'item weight unit', 'item length', 'item length unit', 'item width', 'item width unit', 'item height', 'item height unit', 'item package length', 'package length unit', 'item package width', 'package width unit', 'item package height', 'package height unit', 'package weight', 'package weight unit', 'country of origin', 'dangerous goods regulations'}

def _build_parent_title_from_child(child_title, parent_sku):
  if not child_title: return f"({parent_sku})" if parent_sku else ""
  t = str(child_title).strip()
  if not parent_sku: return t
  return re.sub(r'\([^()]*\)\s*$', f'({parent_sku})', t) if re.search(r'\([^()]*\)\s*$', t) else f"{t} ({parent_sku})"

def _detect_amz_sheet_rows(ws, default_lbl=3, default_key=4, default_data=7):
  r1_str = " ".join(str(ws.cell(1, c).value or '') for c in range(1, min(15, ws.max_column + 1)))
  lbl_m, key_m, dat_m = re.search(r'labelRow=(\d+)', r1_str), re.search(r'attributeRow=(\d+)', r1_str), re.search(r'dataRow=(\d+)', r1_str)
  if lbl_m and key_m and dat_m: return int(lbl_m.group(1)), int(key_m.group(1)), int(dat_m.group(1))
  for r in range(1, min(10, ws.max_row + 1)):
    vals = [str(ws.cell(r, c).value or '').strip().lower() for c in range(1, min(40, ws.max_column + 1))]
    if any(v in ('vendor sku', 'sku', 'seller sku') for v in vals) and any(v in ('item name', 'product type', 'brand name') for v in vals):
      next_vals = [str(ws.cell(r + 1, c).value or '').strip().lower() for c in range(1, min(20, ws.max_column + 1))]
      key_row = r + 1 if any(('#' in v or '::' in v or '_sku' in v) for v in next_vals if v) else None
      return r, key_row, max(default_data, (key_row or r) + 1)
  return default_lbl, default_key, default_data

def run_seller_central(uploaded_file, ctx):
  wb = load_workbook_clean(uploaded_file, uploaded_file.name)
  ws = next((wb[s] for s in wb.sheetnames if s.strip().lower() == 'template'), None) or pick_template_sheet(wb)
  try: ungroup_worksheet(ws)
  except Exception: pass

  sc_lbl_r, sc_key_r, sc_data_r = _detect_amz_sheet_rows(ws, 4, 5, 7)
  sc_key_to_col, sc_lbl_occ_to_col, sc_lbl_counts, sc_col_to_lbl = {}, {}, {}, {}
  for c in range(1, ws.max_column + 1):
    lbl = str(ws.cell(sc_lbl_r, c).value or '').strip().lower()
    key = str(ws.cell(sc_key_r, c).value or '').strip() if sc_key_r else ''
    nk = _norm_amazon_key(key)
    sc_col_to_lbl[c] = lbl
    if nk and nk not in sc_key_to_col: sc_key_to_col[nk] = c
    if lbl:
      occ = sc_lbl_counts.get(lbl, 0) + 1
      sc_lbl_counts[lbl] = occ
      sc_lbl_occ_to_col[(lbl, occ)] = c

  def get_sc_col(labels=(), keys=(), occ=1):
    for k in keys:
      nk = _norm_amazon_key(k)
      if nk in sc_key_to_col: return sc_key_to_col[nk]
    for l in labels:
      ll = l.strip().lower()
      if (ll, occ) in sc_lbl_occ_to_col: return sc_lbl_occ_to_col[(ll, occ)]
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
  c_il, c_ilu = get_sc_col(['item length'], ['item_dimensions#1.length.value']), get_sc_col(['item length unit'], ['item_dimensions#1.length.unit'])
  c_iw, c_iwu = get_sc_col(['item width'], ['item_dimensions#1.width.value']), get_sc_col(['item width unit'], ['item_dimensions#1.width.unit'])
  c_ih, c_ihu = get_sc_col(['item height'], ['item_dimensions#1.height.value']), get_sc_col(['item height unit'], ['item_dimensions#1.height.unit'])
  c_pl, c_plu = get_sc_col(['item package length'], ['item_package_dimensions#1.length.value']), get_sc_col(['package length unit'], ['item_package_dimensions#1.length.unit'])
  c_pw, c_pwu = get_sc_col(['item package width'], ['item_package_dimensions#1.width.value']), get_sc_col(['package width unit'], ['item_package_dimensions#1.width.unit'])
  c_ph, c_phu = get_sc_col(['item package height'], ['item_package_dimensions#1.height.value']), get_sc_col(['package height unit'], ['item_package_dimensions#1.height.unit'])
  c_pwt, c_pwtu = get_sc_col(['package weight'], ['item_package_weight#1.value']), get_sc_col(['package weight unit'], ['item_package_weight#1.unit'])
  c_coo, c_dg = get_sc_col(['country of origin'], ['country_of_origin#1.value']), get_sc_col(['dangerous goods regulations'], ['supplier_declared_dg_hz_regulation#1.value'])

  parent_inherited_sc_cols = {c for c, lbl in sc_col_to_lbl.items() if lbl in PARENT_INHERITED_LABELS}
  curr_r = sc_data_r
  while curr_r <= ws.max_row and ws.cell(curr_r, c_sku).value not in (None, ''): curr_r += 1

  today_str = datetime.now().strftime("%Y-%m-%d")
  mrp_fill, mrp_font = HIGHLIGHT_FILL, HIGHLIGHT_FONT
  material_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
  action_val = "Create or Replace (Full Update)"
  parent_val = pick_dropdown_value(ws, c_parentage, "Parent", "Parent") if c_parentage else "Parent"
  child_val = pick_dropdown_value(ws, c_parentage, "Child", "Child") if c_parentage else "Child"
  var_theme_val, cond_new_val, fulfillment_default_val, qty_zero_val, ship_tpl_val, asin_type_val = "SIZE/COLOR", "New", "Fulfillment by Merchant (Default)", 0, "Migrated Template", "ASIN"
  custom_parent_model = (ctx.get('sc_parent_model') or '').strip()
  generated_rows, model_first_child_cells = [], {}

  for raw_fld in ctx['folders']:
    clean_style = clean_sku_name(raw_fld)
    det_col = resolve_color(raw_fld, clean_style, ctx)
    cat = resolve_category(clean_style, ctx, ws.title)
    det_gen = resolve_gender(clean_style, ctx)
    model_str = custom_parent_model if custom_parent_model else clean_style
    urls = get_sku_img_urls(raw_fld, ctx, clean_style)
    mat_val, sleeve_val = resolve_fabric(clean_style, cat, ctx), resolve_sleeve_length(clean_style, cat, ctx)
    bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
    gen_kw = generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)
    style_val, item_type_val = get_style_from_dropdown(cat), get_item_type_name_from_dropdown(cat)
    color_map_val, pocket_val = map_to_standard_color_map(det_col), resolve_pocket(clean_style, cat, ctx)
    fab_val = f"100% {mat_val}"
    prod_type_val = resolve_sc_product_type_for_row(ws, c_prod_type, raw_prod_type=SC_PRODUCT_TYPE.get(cat, ''), sheet_name=cat, item_type=item_type_val, model_name=model_str, item_name=clean_style)
    dept_val, gender_val = ("Womens", "Female") if det_gen == "Women" else ("Mens", "Male")

    for sz in ctx['sizes']:
      clean_sz = clean_sku_name(sz)
      full_sku = f"{clean_style}-{clean_sz}"
      size_class_val = "Numeric" if clean_sz.isdigit() else "Alpha"
      child_title = build_vendor_title(ctx['brand'], cat, full_sku, ctx['keywords'], 150, det_gen)
      cells = {}
      def s_col(col_idx, val, dropdown=False):
        if not col_idx or val in (None, ''): return
        cells[col_idx] = pick_dropdown_value(ws, col_idx, val, val) if dropdown else val

      s_col(c_sku, full_sku); s_col(c_prod_type, prod_type_val); s_col(c_title, child_title)
      s_col(c_brand, ctx['brand']); s_col(c_mod_num, full_sku); s_col(c_mod_name, model_str); s_col(c_mfg, ctx['mfg'])
      if c_main_img and urls: s_col(c_main_img, urls[0])
      for idx_img, c_oi in enumerate(c_other_imgs):
        if c_oi and idx_img + 1 < len(urls): s_col(c_oi, urls[idx_img + 1])
      s_col(c_desc, _limit_desc_200(desc_text))
      for idx_b, c_bp in enumerate(c_bullets):
        if c_bp and idx_b < len(bullets_list): s_col(c_bp, bullets_list[idx_b])
      s_col(c_gkw, gen_kw); s_col(c_lifestyle, "Casual", dropdown=True); s_col(c_style, style_val, dropdown=True)
      s_col(c_dept, dept_val, dropdown=True); s_col(c_gender, gender_val, dropdown=True); s_col(c_age, "Adult", dropdown=True)
      s_col(c_ssys, "IN", dropdown=True); s_col(c_scls, size_class_val, dropdown=True); s_col(c_sval, clean_sz)
      s_col(c_mat, mat_val, dropdown=True); s_col(c_fab, fab_val); s_col(c_noi, 1)
      s_col(c_itn, item_type_val, dropdown=True); s_col(c_spec_sz, "Standard", dropdown=True)
      s_col(c_cmap, color_map_val, dropdown=True); s_col(c_color, det_col); s_col(c_part, full_sku)
      s_col(c_fit, "Relaxed", dropdown=True); s_col(c_poc, pocket_val, dropdown=True); s_col(c_care, "Machine Wash", dropdown=True)
      s_col(c_mfg_cnt, ctx['mfg']); s_col(c_pat, "Solid", dropdown=True); s_col(c_uc, 1); s_col(c_uct, "count", dropdown=True)
      s_col(c_pld, today_str); s_col(c_hsn_ent, "HSN Code", dropdown=True); s_col(c_hsn_val, get_hsn_code(cat))
      s_col(c_sea, "All", dropdown=True); s_col(c_pkr, ctx['mfg']); s_col(c_sld, sleeve_val, dropdown=True)
      s_col(c_st, "Regular Sleeve", dropdown=True); s_col(c_iwt, 350); s_col(c_iwtu, "Grams", dropdown=True)
      if ctx.get('selling_price'): s_col(c_price, ctx['selling_price'])
      if ctx.get('mrp'): s_col(c_mrp, int(ctx['mrp']) if str(ctx['mrp']).isdigit() else ctx['mrp'])
      s_col(c_il, 10); s_col(c_ilu, "Inches", dropdown=True); s_col(c_iw, 2); s_col(c_iwu, "Inches", dropdown=True)
      s_col(c_ih, 1.3); s_col(c_ihu, "Inches", dropdown=True); s_col(c_pl, 10); s_col(c_plu, "Inches", dropdown=True)
      s_col(c_pw, 2); s_col(c_pwu, "Inches", dropdown=True); s_col(c_ph, 1.3); s_col(c_phu, "Inches", dropdown=True)
      s_col(c_pwt, 380); s_col(c_pwtu, "Grams", dropdown=True); s_col(c_coo, "India", dropdown=True); s_col(c_dg, "Not Applicable", dropdown=True)
      if model_str not in model_first_child_cells: model_first_child_cells[model_str] = cells
      generated_rows.append((full_sku, model_str, cells))

  preview, total_written = [], 0

  def apply_common_fixed_fields(row_idx):
    if c_action: ws.cell(row_idx, c_action).value = action_val
    if c_var_theme: ws.cell(row_idx, c_var_theme).value = var_theme_val
    if c_cond: ws.cell(row_idx, c_cond).value = cond_new_val
    if c_fulfill: ws.cell(row_idx, c_fulfill).value = fulfillment_default_val
    if c_qty: ws.cell(row_idx, c_qty).value = qty_zero_val

  def write_parent_row(model_sku):
    nonlocal curr_r, total_written
    ref_cells = model_first_child_cells.get(model_sku) or (generated_rows[0][2] if generated_rows else {})
    for sc_c in parent_inherited_sc_cols:
      if sc_c in ref_cells and ref_cells[sc_c] not in (None, ''): ws.cell(curr_r, sc_c).value = ref_cells[sc_c]
    ws.cell(curr_r, c_sku).value = model_sku
    if c_prod_type and c_prod_type in ref_cells: ws.cell(curr_r, c_prod_type).value = ref_cells[c_prod_type]
    if c_parentage: ws.cell(curr_r, c_parentage).value = parent_val
    if c_mod_name:
      ws.cell(curr_r, c_mod_name).value = model_sku
      ws.cell(curr_r, c_mod_name).fill = HIGHLIGHT_FILL; ws.cell(curr_r, c_mod_name).font = HIGHLIGHT_FONT
    parent_title = ""
    if c_title:
      parent_title = _build_parent_title_from_child(ref_cells.get(c_title, ''), model_sku)
      ws.cell(curr_r, c_title).value = parent_title
    apply_common_fixed_fields(curr_r)
    preview.append({'Row': curr_r, 'SKU': model_sku, 'Product Type': ref_cells.get(c_prod_type, '') if c_prod_type else '', 'Listing Action': action_val, 'Parentage Level': parent_val, 'Parent SKU': '', 'Variation Theme': var_theme_val, 'Shipping Template': '', 'Item Name': parent_title})
    curr_r += 1; total_written += 1

  def write_child_row(child_sku, model_str, mapped_cells):
    nonlocal curr_r, total_written
    for sc_c, val in mapped_cells.items(): ws.cell(curr_r, sc_c).value = val
    if c_mat and c_mat in mapped_cells: ws.cell(curr_r, c_mat).fill = material_fill
    if c_mrp and not ctx.get('mrp'): ws.cell(curr_r, c_mrp).fill = mrp_fill; ws.cell(curr_r, c_mrp).font = mrp_font
    ws.cell(curr_r, c_sku).value = child_sku
    if c_parentage: ws.cell(curr_r, c_parentage).value = child_val
    if c_parent_sku and model_str: ws.cell(curr_r, c_parent_sku).value = model_str
    if c_mod_name:
      if model_str: ws.cell(curr_r, c_mod_name).value = model_str
      ws.cell(curr_r, c_mod_name).fill = HIGHLIGHT_FILL; ws.cell(curr_r, c_mod_name).font = HIGHLIGHT_FONT
    if c_prod_id_type: ws.cell(curr_r, c_prod_id_type).value = asin_type_val
    if c_ship_tpl: ws.cell(curr_r, c_ship_tpl).value = ship_tpl_val
    apply_common_fixed_fields(curr_r)
    preview.append({'Row': curr_r, 'SKU': child_sku, 'Product Type': mapped_cells.get(c_prod_type, '') if c_prod_type else '', 'Listing Action': action_val, 'Parentage Level': child_val, 'Parent SKU': model_str, 'Variation Theme': var_theme_val, 'Shipping Template': ship_tpl_val, 'Item Name': mapped_cells.get(c_title, '') if c_title else ''})
    curr_r += 1; total_written += 1

  inserted_parents = set()
  for c_sku_val, m_str, m_cells in generated_rows:
    if m_str and m_str not in inserted_parents:
      inserted_parents.add(m_str); write_parent_row(m_str)
    write_child_row(c_sku_val, m_str, m_cells)

  output = io.BytesIO()
  wb.save(output); output.seek(0)
  return output.getvalue(), preview, total_written

# ======================================================================
# 🛍️ DEDICATED MEESHO ENGINE (MeeshoTemplate2PricesGSTIN - Shirts/Trousers/All)
# ======================================================================
MEESHO_COLOR_SYNONYMS = {'pista': 'Mint Green', 'pistachio': 'Mint Green', 'sea green': 'Mint Green', 'seagreen': 'Mint Green', 'lightgreen': 'Mint Green', 'light green': 'Mint Green', 'lightgrey': 'Grey Melange', 'light grey': 'Grey Melange', 'lightgray': 'Grey Melange', 'light gray': 'Grey Melange', 'darkgrey': 'Charcoal', 'dark grey': 'Charcoal', 'navy': 'Navy Blue', 'navyblue': 'Navy Blue', 'navy blue': 'Navy Blue', 'darkblue': 'Navy Blue', 'dark blue': 'Navy Blue', 'skyblue': 'Blue', 'sky blue': 'Blue', 'lightblue': 'Blue', 'light blue': 'Blue', 'iceblue': 'Aqua Blue', 'ice blue': 'Aqua Blue', 'chiku': 'Cream', 'chikoo': 'Cream', 'badami': 'Beige', 'khakhi': 'Khaki', 'biege': 'Beige', 'coffee': 'Brown', 'chocolate': 'Brown', 'camel': 'Brown', 'tan': 'Brown', 'wine': 'Maroon', 'burgundy': 'Maroon', 'off white': 'Off-White', 'offwhite': 'Off-White', 'lavender': 'Lavendar', 'lavendar': 'Lavendar'}
MEESHO_VARIATION_SYNONYMS = {'2xl': 'XXL', 'xxl': 'XXL', '3xl': 'XXXL', 'xxxl': 'XXXL', '4xl': '4XL', 'xxxxl': 'XXXXL', '5xl': '5XL', 'xs': 'XS', 'xxs': 'XXS', 's': 'S', 'm': 'M', 'l': 'L', 'xl': 'XL', 'free': 'Free Size', 'freesize': 'Free Size', 'free size': 'Free Size'}
MEESHO_SHIRT_MEASUREMENTS = {'XS': {'chest': 38, 'length': 27.5, 'shoulder': 16.5}, 'S': {'chest': 40, 'length': 28, 'shoulder': 17}, 'M': {'chest': 42, 'length': 29, 'shoulder': 17.5}, 'L': {'chest': 44, 'length': 29.5, 'shoulder': 18}, 'XL': {'chest': 46, 'length': 30, 'shoulder': 19}, 'XXL': {'chest': 48, 'length': 30.5, 'shoulder': 20}, '2XL': {'chest': 48, 'length': 30.5, 'shoulder': 20}, 'XXXL': {'chest': 50, 'length': 31, 'shoulder': 21}, '3XL': {'chest': 50, 'length': 31, 'shoulder': 21}, '4XL': {'chest': 52, 'length': 31.5, 'shoulder': 22}}
MEESHO_TROUSER_MEASUREMENTS = {'24': {'waist': 24, 'hip': 32, 'length': 38}, '26': {'waist': 26, 'hip': 34, 'length': 38}, '28': {'waist': 28, 'hip': 36, 'length': 39}, '30': {'waist': 30, 'hip': 38, 'length': 39}, '32': {'waist': 32, 'hip': 40, 'length': 40}, '34': {'waist': 34, 'hip': 42, 'length': 40}, '36': {'waist': 36, 'hip': 44, 'length': 41}, '38': {'waist': 38, 'hip': 46, 'length': 41}, '40': {'waist': 40, 'hip': 48, 'length': 42}, '42': {'waist': 42, 'hip': 50, 'length': 42}, 'XS': {'waist': 26, 'hip': 36, 'length': 38}, 'S': {'waist': 28, 'hip': 38, 'length': 39}, 'M': {'waist': 30, 'hip': 40, 'length': 39}, 'L': {'waist': 32, 'hip': 42, 'length': 40}, 'XL': {'waist': 34, 'hip': 44, 'length': 40}, 'XXL': {'waist': 36, 'hip': 46, 'length': 41}, '2XL': {'waist': 36, 'hip': 46, 'length': 41}, 'XXXL': {'waist': 38, 'hip': 48, 'length': 41}, '3XL': {'waist': 38, 'hip': 48, 'length': 41}}

def _extract_meesho_col_header(cell_val):
  if cell_val in (None, ''): return ""
  lines = [ln.strip() for ln in str(cell_val).splitlines() if ln.strip()]
  return lines[0] if lines else ""

def analyze_meesho_template(wb):
  main_ws, val_ws = None, None
  for sname in wb.sheetnames:
    ws_try = wb[sname]
    r1_c1 = str(ws_try.cell(1, 1).value or '').strip().lower()
    r2_c1 = str(ws_try.cell(2, 1).value or '').strip().lower()
    r3_c1 = str(ws_try.cell(3, 1).value or '').strip().lower()
    if 'field type' in r1_c1 or 'validation' in sname.lower():
      if val_ws is None: val_ws = ws_try
      continue
    if ('field names' in r2_c1 or 'fields + description' in r3_c1) and main_ws is None:
      if not any(ign in sname.lower() for ign in ('sample', 'example', 'reference', 'instruction', 'return')):
        main_ws = ws_try
  if main_ws is None:
    main_ws = next((wb[s] for s in wb.sheetnames if 'fields + description' in str(wb[s].cell(3, 1).value or '').strip().lower()), None)
  if main_ws is None: return None

  banner = str(main_ws.cell(1, 1).value or main_ws.cell(1, 5).value or main_ws.title).strip()
  headers, compulsory_cols, rec_return_diff = {}, [], 28
  for c in range(1, main_ws.max_column + 1):
    r2_val = str(main_ws.cell(2, c).value or '').strip()
    r3_raw = str(main_ws.cell(3, c).value or '')
    h_name = _extract_meesho_col_header(r3_raw)
    if h_name: headers[c] = h_name
    if 'compulsory' in r2_val.lower(): compulsory_cols.append(c)
    if h_name.lower() == 'wrong/defective returns price':
      m_diff = re.search(r'Rs\.?\s*(\d+)\s*lower', r3_raw, flags=re.IGNORECASE)
      if m_diff: rec_return_diff = int(m_diff.group(1))

  dropdown_map = {}
  if val_ws is not None:
    val_headers = {_norm_h(val_ws.cell(1, vc).value): vc for vc in range(1, val_ws.max_column + 1) if str(val_ws.cell(1, vc).value or '').strip()}
    for c, h_name in headers.items():
      nh = _norm_h(h_name)
      vc = val_headers.get(nh) or (c if (c <= val_ws.max_column and _norm_h(val_ws.cell(1, c).value) == nh) else None)
      if vc is not None:
        opts = [val_ws.cell(r, vc).value for r in range(3, val_ws.max_row + 1) if val_ws.cell(r, vc).value not in (None, '') and str(val_ws.cell(r, vc).value).strip() and not str(val_ws.cell(r, vc).value).strip().lower().startswith('not available]')]
        if opts: dropdown_map[c] = opts
  for c in headers:
    if c not in dropdown_map:
      dv_opts = get_dropdown_options_for_column(main_ws, c)
      if dv_opts: dropdown_map[c] = dv_opts

  det_gender = "Men" if any(w in banner.lower() for w in ['men ', 'mens', '/men']) and 'women' not in banner.lower() else detect_gender(banner)
  det_cat = detect_product_category(banner, "", main_ws.title)
  return {'main_ws': main_ws, 'val_ws': val_ws, 'sheet_name': main_ws.title, 'banner': banner, 'headers': headers, 'compulsory_cols': compulsory_cols, 'dropdown_map': dropdown_map, 'rec_return_diff': rec_return_diff, 'detected_gender': det_gender, 'detected_category': det_cat}

def _meesho_pick_option(options, desired, fallback=None):
  if not options: return desired if desired is not None else fallback
  str_opts = [(o, str(o).strip()) for o in options if o not in (None, '')]
  if not str_opts: return desired if desired is not None else fallback
  for candidate in (desired, fallback):
    if candidate in (None, ''): continue
    cand_str = str(candidate).strip()
    cand_low, cand_norm = cand_str.lower(), _norm_h(cand_str)
    for orig, s_opt in str_opts:
      if s_opt.lower() == cand_low: return orig
    try:
      cand_num = float(cand_str)
      for orig, s_opt in str_opts:
        try:
          if abs(float(s_opt) - cand_num) < 1e-6: return orig
        except Exception: pass
    except Exception: pass
    for orig, s_opt in str_opts:
      if _norm_h(s_opt) == cand_norm: return orig
    for orig, s_opt in str_opts:
      if cand_low in s_opt.lower() or s_opt.lower() in cand_low: return orig
    close = difflib.get_close_matches(cand_str, [s for _, s in str_opts], n=1, cutoff=0.45)
    if close:
      for orig, s_opt in str_opts:
        if s_opt == close[0]: return orig
  return str_opts[0][0]

def _parse_meesho_mfg_details(mfg_raw, mfg_name_input="", mfg_pin_input="", brand_fallback=""):
  raw = str(mfg_raw or '').strip()
  pin = str(mfg_pin_input or '').strip()
  if not pin:
    m_pin = re.search(r'\b(\d{6})\b', raw)
    pin = m_pin.group(1) if m_pin else "394107"
  name, addr = str(mfg_name_input or '').strip(), raw
  if not name:
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if len(lines) >= 2: name, addr = lines[0].rstrip(',:-'), ", ".join(lines[1:])
    else: name = "JB Fashion"
  addr_clean = re.sub(r'[\s,\-]*\b' + re.escape(pin) + r'\b[\s,\-]*$', '', addr).strip().strip(',') or raw or "PLOT NO 41, DEVAM AMB, RING ROAD -300FTRO PARK, NEAR VEDANT EMBRO PARK, Kosad, Surat"
  return name, addr_clean, (int(pin) if str(pin).isdigit() else pin)

def _extract_meesho_group_prefix(folder_name):
  clean = clean_sku_name(folder_name)
  base = pl_extract_style_code(clean)
  base_no_num = re.sub(r'-\d+$', '', base)
  return base_no_num.lower() if base_no_num else base.lower()

def _infer_meesho_garment_attributes(raw_fld, clean_style, cat, det_gen, det_col, ctx):
  fld_low = f"{raw_fld} {clean_style}".lower()
  vr = (ctx.get('vision_results') or {}).get(clean_style) or {}
  col_key = str(extract_explicit_color_from_text(raw_fld) or det_col or 'Black').strip().lower()
  for tok in ('pista', 'chiku', 'chikoo', 'light grey', 'lightgrey', 'sky blue', 'sky-blue', 'skyblue', 'navyblue', 'navy blue', 'navy', 'khakhi', 'off white', 'offwhite', 'lavender'):
    if tok in fld_low: col_key = tok; break
  meesho_color_desired = MEESHO_COLOR_SYNONYMS.get(col_key, det_col)

  ai_sleeve = str(vr.get('sleeve_length') or '').lower()
  if any(w in fld_low for w in ['full-shirt', 'full_shirt', 'full sleeve', 'fullsleeve', 'crush-lining', 'rib-shirt', 'utility-shirt']): meesho_sleeve, title_sleeve = "Long Sleeves", "Full Sleeve"
  elif any(w in fld_low for w in ['half-shirt', 'half_shirt', 'half sleeve', 'short-kurta', 'aloha', 'crush-', 'embro-', 'insta-', 'lilen-strip', 'linen-e-', 'oxford-lining', 'resort-', 'sunrise-lining', 'turkys']): meesho_sleeve, title_sleeve = "Short Sleeves", "Short Sleeve"
  elif 'long' in ai_sleeve or 'full' in ai_sleeve: meesho_sleeve, title_sleeve = "Long Sleeves", "Full Sleeve"
  elif 'three' in ai_sleeve or '3/4' in ai_sleeve: meesho_sleeve, title_sleeve = "Three-Quarter Sleeves", "3/4 Sleeve"
  elif 'sleeveless' in ai_sleeve: meesho_sleeve, title_sleeve = "Sleeveless", "Sleeveless"
  else: meesho_sleeve, title_sleeve = "Short Sleeves", "Short Sleeve"

  if any(w in fld_low for w in ['mandarin', 'short-kurta', 'crazy-half-shirt', 'kurta']): meesho_neck = "Mandarin"
  elif any(w in fld_low for w in ['cuban', 'aloha']) or re.match(r'^shirt-\d+', clean_style.lower()): meesho_neck = "Cuban Collar"
  elif any(w in fld_low for w in ['linen-e-', 'spread']): meesho_neck = "Spread Collar"
  else: meesho_neck = "Slim Collar"

  ai_pat, ai_mat = str(vr.get('pattern') or '').strip(), str(vr.get('material') or '').strip()
  if any(w in fld_low for w in ['aloha', 'leaf', 'botanical']): meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Printed", "Botanical", "Polycotton", "Leaf Print Casual"
  elif 'embro-flower' in fld_low or ('embro' in fld_low and 'flower' in fld_low): meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Embroidered", "Embroidered", "Cotton", "Floral Embriodered Casual"
  elif 'linen-e-' in fld_low: meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Embroidered", "Embroidered", "Cotton", "Embriodered Print Double-Breasted Casual"
  elif 'embro' in fld_low or 'embroidered' in ai_pat.lower(): meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Embroidered", "Embroidered", "Cotton", "Embriodered Casual"
  elif any(w in fld_low for w in ['check', 'chacks', 'turkys']) or 'check' in ai_pat.lower(): meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Checked", "Checked", "Cotton Blend", "Chacks print Casual"
  elif any(w in fld_low for w in ['strip', 'lining', 'crazy-full', 'crazy-half']) or 'strip' in ai_pat.lower():
    meesho_pattern, meesho_print_type = "Striped", "Striped"
    meesho_fabric = "Cotton" if 'oxford' in fld_low else ("Polycotton" if any(w in fld_low for w in ['crush-lining', 'lilen-strip', 'sunrise-lining']) else (ai_mat if ai_mat in ('Cotton', 'Polycotton', 'Cotton Blend', 'Polyester') else "Polyester"))
    title_desc = "Striped Mandarin Collar" if meesho_neck == "Mandarin" else "Striped Casual"
  elif 'rib-shirt' in fld_low or 'corduroy' in fld_low: meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Solid", "Solid", "Polyester", "Corduroy Casual"
  elif 'short-kurta' in fld_low: meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Solid", "Solid", "Polyester", "Mandarin Collar"
  elif 'utility' in fld_low: meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Solid", "Solid", "Polycotton", "Solid Casual"
  elif 'resort' in fld_low: meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Solid", "Solid", "Cotton Blend", "Solid Casual"
  elif 'crush' in fld_low: meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Solid", "Solid", "Polyester", "Solid Casual"
  elif 'insta' in fld_low or re.match(r'^shirt-\d+', clean_style.lower()): meesho_pattern, meesho_print_type, meesho_fabric, title_desc = "Solid", "Solid", "Polyester", "Textured Casual"
  else:
    meesho_pattern = ai_pat if ai_pat else "Solid"
    meesho_print_type = meesho_pattern
    meesho_fabric = ai_mat if ai_mat else DEFAULT_CATEGORY_FABRICS.get(cat, "Cotton Blend")
    title_desc = "Solid Casual" if meesho_pattern.lower() == 'solid' else f"{meesho_pattern} Casual"

  brand_str = f"{ctx['brand']} " if ctx.get('brand') else ""
  g_title, cat_label = ("Women's" if det_gen == "Women" else "Men's"), CATEGORY_LABELS.get(cat, "Casual Wear")
  if cat == 'shirt':
    product_name = f"{brand_str}{g_title} {title_desc} Shirt".strip() if 'Double-Breasted' in title_desc else f"{brand_str}{g_title} {title_desc} {title_sleeve} Shirt".strip()
  elif cat in ('trousers', 'cargo', 'joggers', 'jeans', 'shorts'):
    fit_word = "Cargo" if cat == 'cargo' else ("Jogger" if cat == 'joggers' else "Regular Fit")
    product_name = f"{brand_str}{g_title} {meesho_pattern} {fit_word} {cat_label}".strip()
  else: product_name = f"{brand_str}{g_title} {title_desc} {cat_label}".strip()
  return {'product_name': re.sub(r'\s+', ' ', product_name).strip(), 'color': meesho_color_desired, 'fabric': meesho_fabric, 'pattern': meesho_pattern, 'print_type': meesho_print_type, 'neck': meesho_neck, 'sleeve_length': meesho_sleeve}

def run_meesho(uploaded_file, ctx, size_charts_data=None):
  wb = load_workbook_clean(uploaded_file, uploaded_file.name)
  m_info = analyze_meesho_template(wb)
  if not m_info: raise ValueError("Meesho template sheet ('Fields + Description:') nahi mili. Kripya valid Meesho template upload karein.")
  ws = m_info['main_ws']
  try: ungroup_worksheet(ws)
  except Exception: pass

  headers, dropdown_map, rec_diff, tpl_cat, tpl_gen = m_info['headers'], m_info['dropdown_map'], m_info['rec_return_diff'], m_info['detected_category'], m_info['detected_gender']
  norm_to_col = {}
  for c, h_name in headers.items():
    nh = _norm_h(h_name)
    if nh and nh not in norm_to_col: norm_to_col[nh] = c

  def get_mcol(candidates):
    for cand in candidates:
      nc = _norm_h(cand)
      if nc in norm_to_col: return norm_to_col[nc]
    return None

  c_prod_name, c_variation, c_meesho_price, c_return_price, c_mrp = get_mcol(['Product Name']), get_mcol(['Variation']), get_mcol(['Meesho Price']), get_mcol(['Wrong/Defective Returns Price']), get_mcol(['MRP'])
  c_gst, c_hsn, c_weight, c_inventory, c_country = get_mcol(['GST %', 'GST']), get_mcol(['HSN ID', 'HSN', 'HSN Code']), get_mcol(['Net Weight (gms)', 'Net Weight']), get_mcol(['Inventory']), get_mcol(['Country of Origin'])
  c_mfg_name, c_mfg_addr, c_mfg_pin = get_mcol(['Manufacturer Name']), get_mcol(['Manufacturer Address']), get_mcol(['Manufacturer Pincode'])
  c_pkr_name, c_pkr_addr, c_pkr_pin = get_mcol(['Packer Name']), get_mcol(['Packer Address']), get_mcol(['Packer Pincode'])
  c_imp_name, c_imp_addr, c_imp_pin = get_mcol(['Importer Name']), get_mcol(['Importer Address']), get_mcol(['Importer Pincode'])
  c_color, c_fabric, c_fit, c_generic = get_mcol(['Color', 'Colour']), get_mcol(['Fabric']), get_mcol(['Fit/Shape', 'Fit/ Shape', 'Fit']), get_mcol(['Generic Name'])
  c_neck, c_net_qty, c_occasion, c_pattern, c_print_type, c_sleeve_len = get_mcol(['Neck', 'Collar']), get_mcol(['Net Quantity (N)', 'Net Quantity']), get_mcol(['Occasion']), get_mcol(['Pattern']), get_mcol(['Print Or Pattern Type', 'Print or Pattern Type']), get_mcol(['Sleeve Length'])
  c_chest, c_length_sz, c_shoulder = get_mcol(['Chest Size', 'Bust Size']), get_mcol(['Length Size']), get_mcol(['Shoulder Size'])
  c_stretch, c_type, c_waist_rise, c_hip, c_waist_sz = get_mcol(['Stretchability']), get_mcol(['Type']), get_mcol(['Waist Rise']), get_mcol(['Hip Size']), get_mcol(['Waist Size'])
  c_imgs = [get_mcol(['Image 1 (Front)', 'Image 1', 'Image1']), get_mcol(['Image 2', 'Image2']), get_mcol(['Image 3', 'Image3']), get_mcol(['Image 4', 'Image4'])]
  c_style_id, c_sku_id, c_brand_name, c_group_id = get_mcol(['Product ID / Style ID', 'Style ID', 'Product ID']), get_mcol(['SKU ID', 'Seller SKU', 'SKU']), get_mcol(['Brand Name']), get_mcol(['Group ID'])
  c_distress, c_hemline, c_length_opt, c_pockets, c_sleeve_style, c_waist_closure = get_mcol(['Distress']), get_mcol(['Hemline']), get_mcol(['Length']), get_mcol(['Number of Pockets']), get_mcol(['Sleeve Styling']), get_mcol(['Waist closure', 'Waist Closure'])

  def mset(row_idx, col_idx, val, dropdown=False, fill=None, font=None):
    if not col_idx or val in (None, ''): return
    cell = ws.cell(row_idx, col_idx)
    cell.value = _meesho_pick_option(dropdown_map[col_idx], val, val) if (dropdown and col_idx in dropdown_map) else val
    if fill: cell.fill = fill
    if font: cell.font = font

  curr_r = 5
  check_col = c_prod_name or c_sku_id or 4
  while curr_r <= ws.max_row and ws.cell(curr_r, check_col).value not in (None, ''): curr_r += 1

  mfg_name, mfg_addr, mfg_pin = _parse_meesho_mfg_details(ctx.get('mfg', ''), ctx.get('meesho_mfg_name', ''), ctx.get('meesho_mfg_pin', ''), ctx.get('brand', ''))
  raw_meesho_price, raw_return_price, raw_mrp = str(ctx.get('selling_price') or '').strip(), str(ctx.get('meesho_return_price') or '').strip(), str(ctx.get('mrp') or '').strip()
  meesho_price_val = int(raw_meesho_price) if raw_meesho_price.isdigit() else (raw_meesho_price or None)
  mrp_val = int(raw_mrp) if raw_mrp.isdigit() else (raw_mrp or None)
  return_price_val = (int(raw_return_price) if raw_return_price.isdigit() else raw_return_price) if raw_return_price else ((meesho_price_val - rec_diff) if (isinstance(meesho_price_val, int) and meesho_price_val > rec_diff) else None)
  gst_val, inv_val = int(ctx.get('meesho_gst', 5)), int(ctx.get('quantity', 0))

  group_options = dropdown_map.get(c_group_id, [f"Group  {i}" for i in range(1, 201)])
  family_to_group, group_counter, total_added, preview = {}, 0, 0, []

  for raw_fld in ctx['folders']:
    clean_style = clean_sku_name(raw_fld)
    cat = resolve_category(clean_style, ctx, ws.title)
    if cat == 'apparel' and tpl_cat and tpl_cat != 'apparel': cat = tpl_cat
    det_gen = tpl_gen if tpl_gen in ('Men', 'Women') else resolve_gender(clean_style, ctx)
    det_col = resolve_color(raw_fld, clean_style, ctx)

    fam_key = _extract_meesho_group_prefix(raw_fld)
    if fam_key not in family_to_group:
      group_counter += 1
      family_to_group[fam_key] = _meesho_pick_option(group_options, f"Group  {group_counter}", f"Group {group_counter:02d}")
    group_id_val = family_to_group[fam_key]

    attrs = _infer_meesho_garment_attributes(raw_fld, clean_style, cat, det_gen, det_col, ctx)
    sku_img_urls = get_sku_img_urls(raw_fld, ctx, clean_style)
    custom_wt = str(ctx.get('meesho_weight') or '').strip()
    weight_val = int(custom_wt) if custom_wt.isdigit() else (300 if cat in ('shirt', 'tshirt', 'top', 'kurti', 'kurta') else 350)
    hsn_desired = str(ctx.get('hsn') or '').strip() or get_hsn_code(cat)
    generic_desired = "Trousers" if cat in ('trousers', 'cargo', 'joggers', 'pants') else ("Shirts" if cat == 'shirt' else CATEGORY_LABELS.get(cat, "Shirts"))

    preview.append({'Sheet': ws.title, 'Product ID / Style ID': clean_style, 'Group ID': group_id_val, 'Product Name': attrs['product_name'], 'Color': _meesho_pick_option(dropdown_map.get(c_color), attrs['color'], "Black"), 'Fabric': _meesho_pick_option(dropdown_map.get(c_fabric), attrs['fabric'], "Cotton"), 'Pattern': _meesho_pick_option(dropdown_map.get(c_pattern), attrs['pattern'], "Solid"), 'Meesho Price': meesho_price_val or 'Highlight', 'Return Price': return_price_val or 'Highlight', 'MRP': mrp_val or 'Highlight'})

    for sz in ctx['sizes']:
      clean_sz = clean_sku_name(sz)
      full_sku = f"{clean_style}-{clean_sz}"
      var_desired = MEESHO_VARIATION_SYNONYMS.get(clean_sz.lower(), clean_sz)
      mset(curr_r, c_prod_name, attrs['product_name'])
      mset(curr_r, c_variation, var_desired, dropdown=True)
      if c_meesho_price:
        if meesho_price_val is not None: ws.cell(curr_r, c_meesho_price).value = meesho_price_val
        else: ws.cell(curr_r, c_meesho_price).fill = HIGHLIGHT_FILL; ws.cell(curr_r, c_meesho_price).font = HIGHLIGHT_FONT
      if c_return_price:
        if return_price_val is not None: ws.cell(curr_r, c_return_price).value = return_price_val
        elif meesho_price_val is None: ws.cell(curr_r, c_return_price).fill = HIGHLIGHT_FILL; ws.cell(curr_r, c_return_price).font = HIGHLIGHT_FONT
      if c_mrp:
        if mrp_val is not None: ws.cell(curr_r, c_mrp).value = mrp_val
        else: ws.cell(curr_r, c_mrp).fill = HIGHLIGHT_FILL; ws.cell(curr_r, c_mrp).font = HIGHLIGHT_FONT

      mset(curr_r, c_gst, gst_val, dropdown=True); mset(curr_r, c_hsn, hsn_desired, dropdown=True)
      mset(curr_r, c_weight, weight_val); mset(curr_r, c_inventory, inv_val); mset(curr_r, c_country, "India", dropdown=True)
      mset(curr_r, c_mfg_name, mfg_name); mset(curr_r, c_mfg_addr, mfg_addr); mset(curr_r, c_mfg_pin, mfg_pin)
      mset(curr_r, c_pkr_name, mfg_name); mset(curr_r, c_pkr_addr, mfg_addr); mset(curr_r, c_pkr_pin, mfg_pin)
      mset(curr_r, c_imp_name, "Not Required"); mset(curr_r, c_imp_addr, "Not Required"); mset(curr_r, c_imp_pin, "Not Required")
      mset(curr_r, c_color, attrs['color'], dropdown=True); mset(curr_r, c_fabric, attrs['fabric'], dropdown=True)
      mset(curr_r, c_fit, "Cargo" if cat == 'cargo' else ("Jogger Pants" if cat == 'joggers' else "Regular"), dropdown=True)
      mset(curr_r, c_generic, generic_desired, dropdown=True); mset(curr_r, c_net_qty, 1, dropdown=True)
      mset(curr_r, c_occasion, "Casual", dropdown=True); mset(curr_r, c_pattern, attrs['pattern'], dropdown=True); mset(curr_r, c_print_type, attrs['print_type'], dropdown=True)
      if c_neck: mset(curr_r, c_neck, attrs['neck'], dropdown=True)
      if c_sleeve_len: mset(curr_r, c_sleeve_len, attrs['sleeve_length'], dropdown=True)

      sz_key = clean_sz.upper()
      sh_m = MEESHO_SHIRT_MEASUREMENTS.get(sz_key) or MEESHO_SHIRT_MEASUREMENTS.get(str(var_desired).upper()) or {'chest': 42, 'length': 29, 'shoulder': 17.5}
      if c_chest: mset(curr_r, c_chest, sh_m['chest'], dropdown=True)
      if c_shoulder: mset(curr_r, c_shoulder, sh_m['shoulder'], dropdown=True)

      tr_m = MEESHO_TROUSER_MEASUREMENTS.get(sz_key) or MEESHO_TROUSER_MEASUREMENTS.get(str(var_desired).upper())
      if not tr_m:
        try:
          w_num = int(re.findall(r'\d+', clean_sz)[0])
          tr_m = {'waist': w_num, 'hip': w_num + 8, 'length': 40}
        except Exception: tr_m = {'waist': 32, 'hip': 40, 'length': 40}
      if c_waist_sz: mset(curr_r, c_waist_sz, tr_m['waist'], dropdown=True)
      if c_hip: mset(curr_r, c_hip, tr_m['hip'], dropdown=True)
      if c_length_sz:
        mset(curr_r, c_length_sz, tr_m['length'] if (c_waist_sz or cat in ('trousers', 'cargo', 'joggers', 'jeans')) else sh_m['length'], dropdown=True)
      if c_type: mset(curr_r, c_type, "Cargos" if cat == 'cargo' else ("Joggers" if cat == 'joggers' else "Regular"), dropdown=True)
      if c_waist_rise: mset(curr_r, c_waist_rise, "Mid Rise", dropdown=True)

      for img_i, c_im in enumerate(c_imgs):
        if c_im and img_i < len(sku_img_urls): ws.cell(curr_r, c_im).value = sku_img_urls[img_i]
      mset(curr_r, c_style_id, clean_style); mset(curr_r, c_sku_id, full_sku); mset(curr_r, c_brand_name, ctx.get('brand', '')); mset(curr_r, c_group_id, group_id_val, dropdown=True)

      if c_hemline: mset(curr_r, c_hemline, "Straight", dropdown=True)
      if c_length_opt: mset(curr_r, c_length_opt, "Full Length" if (c_waist_sz or cat in ('trousers', 'cargo', 'joggers', 'jeans')) else "Regular", dropdown=True)
      if c_pockets: mset(curr_r, c_pockets, 2 if (c_waist_sz or cat in ('trousers', 'cargo', 'joggers', 'jeans')) else 1, dropdown=True)
      if c_sleeve_style: mset(curr_r, c_sleeve_style, "Regular", dropdown=True)
      if c_stretch: mset(curr_r, c_stretch, "No", dropdown=True)
      if c_distress: mset(curr_r, c_distress, "Not Applicable", dropdown=True)
      if c_waist_closure: mset(curr_r, c_waist_closure, "Drawstring" if cat == 'joggers' else "Button", dropdown=True)
      curr_r += 1; total_added += 1

  output = io.BytesIO()
  wb.save(output); output.seek(0)
  return output.getvalue(), preview, total_added, ws.title

# ======================================================================
# UNIVERSAL ENGINE (Flipkart / Myntra)
# ======================================================================
MP_FIELD_CANDIDATES = {
  'Flipkart': {'sku': ['listing sku', 'seller sku', 'sku', 'sku id'], 'title': ['product title', 'listing title', 'title', 'product name'], 'brand': ['brand', 'brand name'], 'description': ['description', 'product description', 'long description'], 'mrp': ['mrp', 'maximum retail price', 'list price', 'mrp inr'], 'price': ['selling price', 'price', 'your price'], 'hsn': ['hsn', 'hsn code'], 'color': ['color', 'color family', 'colour'], 'size': ['size', 'size value', 'variant size'], 'fabric': ['fabric', 'fabric type', 'material'], 'pattern': ['pattern', 'pattern type'], 'fit': ['fit', 'fit type'], 'country': ['country of origin', 'origin country', 'country'], 'generic_kw': ['search keywords', 'keywords', 'search terms'], 'gender': ['gender', 'ideal for'], 'care': ['wash care', 'care instructions'], 'sleeve': ['sleeve', 'sleeve length'], 'manufacturer': ['manufactured by', 'manufacturer', 'packer details', 'marketed by'], 'img1': ['main image url', 'image url 1', 'image 1', 'image1', 'primary image url']},
  'Myntra': {'sku': ['seller sku', 'sku', 'sku id', 'style code'], 'title': ['style name', 'product name', 'title', 'item name'], 'brand': ['brand', 'brand name', 'vendor brand'], 'description': ['style description', 'description', 'product description'], 'mrp': ['mrp', 'maximum retail price', 'list price'], 'price': ['selling price', 'price', 'wholesale price'], 'hsn': ['hsn', 'hsn code'], 'color': ['color', 'colour', 'primary color'], 'size': ['size', 'variant size'], 'fabric': ['fabric', 'material', 'composition'], 'pattern': ['pattern', 'print'], 'fit': ['fit', 'fit type'], 'country': ['country of origin', 'country'], 'generic_kw': ['keywords', 'search keywords'], 'gender': ['gender', 'customer gender'], 'care': ['wash care', 'care instructions'], 'sleeve': ['sleeve length', 'sleeve styling'], 'manufacturer': ['manufactured by', 'manufacturer', 'marketed by', 'packer details'], 'img1': ['image 1', 'image1', 'images1', 'main image url', 'primary image url']},
}

def run_generic_marketplace(uploaded_file, ctx, mp):
  wb = load_workbook_clean(uploaded_file, uploaded_file.name)
  fields = MP_FIELD_CANDIDATES.get(mp, MP_FIELD_CANDIDATES['Flipkart'])
  mrp_fill, mrp_font = HIGHLIGHT_FILL, HIGHLIGHT_FONT
  best_ws, best_row, best_score = None, None, -1
  sku_cands = [_norm_h(x) for x in fields['sku']]
  for sname in wb.sheetnames:
    ws_try = wb[sname]
    for r in range(1, min(ws_try.max_row, 10) + 1):
      vals = [_norm_h(ws_try.cell(r, c).value) for c in range(1, min(ws_try.max_column, 60) + 1)]
      score = sum(1 for v in vals if v in sku_cands)
      if score > best_score: best_score, best_ws, best_row = score, ws_try, r
  if best_ws is None or best_score == 0: raise ValueError(f"{mp} template me 'sku' header wali sheet nahi mili.")
  ws, header_row = best_ws, best_row
  try: ungroup_worksheet(ws)
  except Exception: pass

  cols = {}
  for c in range(1, ws.max_column + 1):
    h = _norm_h(ws.cell(header_row, c).value)
    if h and h not in cols: cols[h] = c

  def find(cands):
    for cd in cands:
      n = _norm_h(cd)
      if n in cols: return cols[n]
    return None

  fcol = {k: find(v) for k, v in fields.items()}
  sku_col = fcol['sku'] or (find(['sku']) or 1)
  img_cols = [c for c in range(1, ws.max_column + 1) if re.match(r'(image|img|images|model)\d*$', _norm_h(ws.cell(header_row, c).value) or '')]
  if fcol.get('img1') and fcol['img1'] not in img_cols: img_cols.insert(0, fcol['img1'])
  img_cols = sorted(set(img_cols))
  bullet_cols = sorted({c for c in range(1, ws.max_column + 1) if re.match(r'(bulletpoint|keyfeature|highlight)\d*$', _norm_h(ws.cell(header_row, c).value) or '')})

  start_row = header_row + 1
  while start_row <= ws.max_row and ws.cell(start_row, sku_col).value not in (None, ''): start_row += 1

  def W(row, key, value, dropdown=False):
    c = fcol.get(key)
    if not c or value in (None, ''): return
    ws.cell(row, c).value = pick_dropdown_value(ws, c, value, value) if dropdown else value

  count, preview = 0, []
  for raw_fld in ctx['folders']:
    clean_style = clean_sku_name(raw_fld)
    det_col = resolve_color(raw_fld, clean_style, ctx)
    cat, det_gen = resolve_category(clean_style, ctx, ws.title), resolve_gender(clean_style, ctx)
    color_family_val = map_to_standard_color_map(det_col)
    mat_val, sleeve_val = resolve_fabric(clean_style, cat, ctx), resolve_sleeve_length(clean_style, cat, ctx)
    bullets_list, desc_text = build_freehand_bullets_and_desc(cat, ctx['brand'], det_gen, mat_val)
    gen_kw = re.sub(r'\s+', ' ', str(generate_generic_keywords_200(cat, det_col, ctx['keywords'], det_gen)).replace(';', ' ')).strip()
    urls = get_sku_img_urls(raw_fld, ctx, clean_style)
    preview.append({'SKU': clean_style, 'Category': cat, 'Gender': det_gen, 'Color': det_col, 'Color Family': color_family_val, 'Material': mat_val, 'Sleeve': sleeve_val})

    for sz in ctx['sizes']:
      clean_sz = clean_sku_name(sz)
      full_sku = f"{clean_style}-{clean_sz}"
      r = start_row + count
      W(r, 'sku', full_sku); W(r, 'title', build_sc_title(ctx['brand'], cat, full_sku, ctx['keywords'], 150, det_gen)); W(r, 'brand', ctx['brand'])
      W(r, 'description', _limit_desc_200(desc_text)); W(r, 'color', color_family_val, dropdown=True)
      W(r, 'size', clean_sz); W(r, 'fabric', mat_val, dropdown=True); W(r, 'pattern', 'Solid', dropdown=True)
      W(r, 'fit', 'Relaxed', dropdown=True); W(r, 'country', 'India', dropdown=True)
      W(r, 'generic_kw', gen_kw); W(r, 'gender', det_gen, dropdown=True); W(r, 'care', 'Machine Wash', dropdown=True)
      W(r, 'sleeve', sleeve_val, dropdown=True); W(r, 'hsn', get_hsn_code(cat)); W(r, 'manufacturer', ctx['mfg'])
      for bi, bc in enumerate(bullet_cols):
        if bi < len(bullets_list): ws.cell(r, bc).value = bullets_list[bi]
      for ii, ic in enumerate(img_cols):
        if ii < len(urls): ws.cell(r, ic).value = urls[ii]
      if fcol.get('mrp'):
        cell = ws.cell(r, fcol['mrp'])
        if ctx['mrp']: cell.value = ctx['mrp']
        else: cell.fill = mrp_fill; cell.font = mrp_font
      W(r, 'price', ctx['selling_price'] or None)
      count += 1

  output = io.BytesIO()
  wb.save(output); output.seek(0)
  return output.getvalue(), preview, count

# ======================================================================
# VC ➜ SC CONVERT ENGINE
# ======================================================================
IGNORED_AMZ_SHEET_KEYWORDS = ('instruction', 'guideline', 'image', 'browse data', 'dropdown', 'valid value', 'data definition', 'definition', 'example', 'change log', 'changelog', 'readme', 'help', 'hidden', 'lookup')

def _is_valid_vc_listing_sheet(ws):
  sname_low = str(ws.title or '').strip().lower()
  if any(k in sname_low for k in IGNORED_AMZ_SHEET_KEYWORDS) or ws.max_row < 4 or ws.max_column < 3: return False
  lbl_r, key_r, _ = _detect_amz_sheet_rows(ws, 3, 4, 7)
  has_sku_header, has_item_header = False, False
  for c in range(1, min(ws.max_column, 80) + 1):
    lbl = str(ws.cell(lbl_r, c).value or '').strip().lower()
    key = _norm_amazon_key(ws.cell(key_r, c).value) if key_r else ''
    if key in ('vendor_sku#1.value', 'contribution_sku#1.value') or lbl in ('vendor sku', 'sku', 'seller sku'): has_sku_header = True
    if key in ('item_name#1.value', 'brand#1.value', 'product_type#1.value') or lbl in ('item name', 'brand name', 'product type', 'model name'): has_item_header = True
  return has_sku_header and has_item_header

def _derive_parent_model_from_sku(sku_str, fallback_prefix=""):
  s = str(sku_str or '').strip()
  if not s: return fallback_prefix or "PARENT-SKU"
  stripped = re.sub(r'[-_](XXS|XS|S|M|L|XL|XXL|2XL|3XL|4XL|5XL|FREE|FREESIZE|\d{2})$', '', s, flags=re.IGNORECASE)
  return stripped if stripped else s

def _merge_vc_duplicate_cell_maps(existing_cells, new_cells, existing_score, new_score):
  primary, secondary = (new_cells, existing_cells) if new_score > existing_score else (existing_cells, new_cells)
  merged = dict(secondary)
  for k, v in primary.items():
    if v not in (None, ''): merged[k] = v
  return merged

def convert_vc_to_sc(vc_file, sc_file, ctx):
  wb_vc = openpyxl.load_workbook(io.BytesIO(vc_file.getvalue()), data_only=True)
  wb_sc = load_workbook_clean(io.BytesIO(sc_file.getvalue()), sc_file.name)
  ws_sc = next((wb_sc[s] for s in wb_sc.sheetnames if s.strip().lower() == 'template'), None) or pick_template_sheet(wb_sc)
  try: ungroup_worksheet(ws_sc)
  except Exception: pass

  sc_lbl_r, sc_key_r, sc_data_r = _detect_amz_sheet_rows(ws_sc, 4, 5, 7)
  sc_key_to_col, sc_lbl_occ_to_col, sc_lbl_counts, sc_col_to_lbl = {}, {}, {}, {}
  for c in range(1, ws_sc.max_column + 1):
    lbl = str(ws_sc.cell(sc_lbl_r, c).value or '').strip().lower()
    key = str(ws_sc.cell(sc_key_r, c).value or '').strip() if sc_key_r else ''
    nk = _norm_amazon_key(key)
    sc_col_to_lbl[c] = lbl
    if nk and nk not in sc_key_to_col: sc_key_to_col[nk] = c
    if lbl:
      occ = sc_lbl_counts.get(lbl, 0) + 1
      sc_lbl_counts[lbl] = occ
      sc_lbl_occ_to_col[(lbl, occ)] = c

  sc_sku_col = sc_key_to_col.get('contribution_sku#1.value') or sc_lbl_occ_to_col.get(('sku', 1)) or sc_lbl_occ_to_col.get(('seller sku', 1)) or 1
  sc_prod_type_col = sc_key_to_col.get('product_type#1.value') or sc_lbl_occ_to_col.get(('product type', 1)) or sc_lbl_occ_to_col.get(('feed product type', 1))
  sc_action_col = sc_key_to_col.get('::record_action') or sc_lbl_occ_to_col.get(('listing action', 1))
  sc_parentage_col = sc_key_to_col.get('parentage_level#1.value') or sc_lbl_occ_to_col.get(('parentage level', 1))
  sc_parent_sku_col = sc_key_to_col.get('child_parent_sku_relationship#1.parent_sku') or sc_lbl_occ_to_col.get(('parent sku', 1))
  sc_var_theme_col = sc_key_to_col.get('variation_theme#1.name') or sc_lbl_occ_to_col.get(('variation theme name', 1))
  sc_title_col = sc_key_to_col.get('item_name#1.value') or sc_lbl_occ_to_col.get(('item name', 1))
  sc_desc_col = sc_key_to_col.get('product_description#1.value') or sc_lbl_occ_to_col.get(('product description', 1)) or sc_lbl_occ_to_col.get(('description', 1))
  sc_prod_id_type_col = sc_key_to_col.get('amzn1.volt.ca.product_id_type') or sc_lbl_occ_to_col.get(('product id type', 1))
  sc_prod_id_val_col = sc_key_to_col.get('amzn1.volt.ca.product_id_value') or sc_lbl_occ_to_col.get(('product id', 1))
  sc_model_name_col = sc_key_to_col.get('model_name#1.value') or sc_lbl_occ_to_col.get(('model name', 1))
  sc_condition_col = sc_key_to_col.get('condition_type#1.value') or sc_lbl_occ_to_col.get(('item condition', 1))
  sc_fulfillment_col = sc_key_to_col.get('fulfillment_availability#1.fulfillment_channel_code') or sc_lbl_occ_to_col.get(('fulfillment channel code (in)', 1))
  sc_quantity_col = sc_key_to_col.get('fulfillment_availability#1.quantity') or sc_lbl_occ_to_col.get(('quantity (in)', 1)) or sc_lbl_occ_to_col.get(('quantity', 1))
  sc_shipping_tpl_col = sc_key_to_col.get('merchant_shipping_group#1.value') or sc_lbl_occ_to_col.get(('shipping template (in)', 1))
  sc_price_col = sc_key_to_col.get('purchasable_offer#1.our_price#1.schedule#1.value_with_tax') or sc_lbl_occ_to_col.get(('your price inr (sell on amazon, in)', 1)) or sc_lbl_occ_to_col.get(('standard price', 1))
  sc_mrp_col = sc_key_to_col.get('purchasable_offer#1.maximum_retail_price#1.schedule#1.value_with_tax') or sc_lbl_occ_to_col.get(('maximum retail price (sell on amazon, in)', 1)) or sc_lbl_occ_to_col.get(('list price', 1))

  parent_inherited_sc_cols = {c for c, lbl in sc_col_to_lbl.items() if lbl in PARENT_INHERITED_LABELS}
  curr_r = sc_data_r
  while curr_r <= ws_sc.max_row and ws_sc.cell(curr_r, sc_sku_col).value not in (None, ''): curr_r += 1

  vc_listing_sheets = [wb_vc[s] for s in wb_vc.sheetnames if _is_valid_vc_listing_sheet(wb_vc[s])] or [pick_template_sheet(wb_vc)]
  deduped_sku_order, deduped_sku_store, sheets_summary = [], {}, []
  custom_parent_override = (ctx.get('sc_parent_model') or '').strip()

  for ws_vc in vc_listing_sheets:
    vc_lbl_r, vc_key_r, vc_data_r = _detect_amz_sheet_rows(ws_vc, 3, 4, 7)
    vc_to_sc_col, vc_lbl_counts = {}, {}
    vc_sku_col = vc_model_name_col = vc_asin_col = vc_prod_type_col = vc_item_type_col = vc_title_col = None

    for c in range(1, ws_vc.max_column + 1):
      lbl = str(ws_vc.cell(vc_lbl_r, c).value or '').strip().lower()
      key = str(ws_vc.cell(vc_key_r, c).value or '').strip() if vc_key_r else ''
      nk = _norm_amazon_key(key)
      if vc_sku_col is None and (nk in ('vendor_sku#1.value', 'contribution_sku#1.value') or lbl in ('vendor sku', 'sku', 'seller sku')): vc_sku_col = c
      if vc_model_name_col is None and (nk == 'model_name#1.value' or lbl == 'model name'): vc_model_name_col = c
      if vc_asin_col is None and (nk in ('merchant_suggested_asin#1.value', 'external_product_id#1.value') or lbl in ('merchant suggested asin', 'external product id')): vc_asin_col = c
      if vc_prod_type_col is None and (nk == 'product_type#1.value' or lbl in ('product type', 'feed product type')): vc_prod_type_col = c
      if vc_item_type_col is None and (nk == 'item_type_name#1.value' or lbl in ('item type name', 'item type keyword')): vc_item_type_col = c
      if vc_title_col is None and (nk == 'item_name#1.value' or lbl in ('item name', 'title', 'product name')): vc_title_col = c
      nk_target = VC_TO_SC_KEY_ALIASES.get(nk, nk)
      occ = 1
      if lbl: occ = vc_lbl_counts.get(lbl, 0) + 1; vc_lbl_counts[lbl] = occ
      sc_c = sc_key_to_col.get(nk_target) if nk_target else None
      if not sc_c:
        for grp in VC_TO_SC_KEY_FALLBACK_GROUPS:
          if nk_target in grp:
            sc_c = next((sc_key_to_col[alt_k] for alt_k in grp if alt_k in sc_key_to_col), None)
            if sc_c: break
        if not sc_c:
          lbl_target = VC_TO_SC_LABEL_ALIASES.get(lbl, lbl)
          sc_c = sc_lbl_occ_to_col.get((lbl_target, occ)) or (sc_lbl_occ_to_col.get((lbl_target, 1)) if occ == 1 else None)
      if sc_c: vc_to_sc_col.setdefault(c, []).append(sc_c)

    if vc_sku_col is None: vc_sku_col = 2
    sheet_raw_rows, sheet_unique_skus, sheet_duplicates_merged = 0, set(), 0

    for r_vc in range(vc_data_r, ws_vc.max_row + 1):
      sku_val = ws_vc.cell(r_vc, vc_sku_col).value
      if sku_val in (None, ''): continue
      sku_str = str(sku_val).strip()
      if sku_str.lower() in ('vendor sku', 'sku', 'seller sku', 'example'): continue
      sheet_raw_rows += 1
      sku_norm_key = sku_str.lower()
      non_empty_vc_cells = sum(1 for c_chk in range(1, ws_vc.max_column + 1) if ws_vc.cell(r_vc, c_chk).value not in (None, ''))
      mapped_cells = {}
      for vc_c, sc_cols_list in vc_to_sc_col.items():
        val = ws_vc.cell(r_vc, vc_c).value
        if val not in (None, ''):
          for sc_c in sc_cols_list: mapped_cells[sc_c] = _limit_desc_200(val) if sc_c == sc_desc_col else val
      if vc_asin_col and sc_prod_id_val_col:
        asin_val = ws_vc.cell(r_vc, vc_asin_col).value
        if asin_val not in (None, ''): mapped_cells[sc_prod_id_val_col] = str(asin_val).strip()

      raw_pt_val = ws_vc.cell(r_vc, vc_prod_type_col).value if vc_prod_type_col else ''
      raw_itn_val = ws_vc.cell(r_vc, vc_item_type_col).value if vc_item_type_col else ''
      raw_title_val = ws_vc.cell(r_vc, vc_title_col).value if vc_title_col else ''
      model_name_val = ws_vc.cell(r_vc, vc_model_name_col).value if vc_model_name_col else None
      has_explicit_model = model_name_val not in (None, '') and str(model_name_val).strip() != ''
      model_str = custom_parent_override if custom_parent_override else (str(model_name_val).strip() if has_explicit_model else _derive_parent_model_from_sku(sku_str, ws_vc.title))

      if sku_norm_key not in deduped_sku_store:
        deduped_sku_order.append(sku_norm_key); sheet_unique_skus.add(sku_norm_key)
        deduped_sku_store[sku_norm_key] = {'sku_str': sku_str, 'model_str': model_str, 'has_explicit_model': has_explicit_model, 'mapped_cells': mapped_cells, 'sheet_title': ws_vc.title, 'raw_pt_val': raw_pt_val or '', 'raw_itn_val': raw_itn_val or '', 'raw_title_val': raw_title_val or '', 'fill_score': non_empty_vc_cells}
      else:
        sheet_duplicates_merged += 1
        existing = deduped_sku_store[sku_norm_key]
        existing['mapped_cells'] = _merge_vc_duplicate_cell_maps(existing['mapped_cells'], mapped_cells, existing['fill_score'], non_empty_vc_cells)
        if not custom_parent_override and has_explicit_model and (not existing['has_explicit_model'] or non_empty_vc_cells > existing['fill_score']): existing['model_str'] = model_str; existing['has_explicit_model'] = True
        if raw_pt_val not in (None, '') and (not existing['raw_pt_val'] or non_empty_vc_cells > existing['fill_score']): existing['raw_pt_val'] = raw_pt_val
        if raw_itn_val not in (None, '') and (not existing['raw_itn_val'] or non_empty_vc_cells > existing['fill_score']): existing['raw_itn_val'] = raw_itn_val
        if raw_title_val not in (None, '') and (not existing['raw_title_val'] or non_empty_vc_cells > existing['fill_score']): existing['raw_title_val'] = raw_title_val
        if non_empty_vc_cells > existing['fill_score']: existing['fill_score'] = non_empty_vc_cells; existing['sheet_title'] = ws_vc.title; existing['sku_str'] = sku_str

    if sheet_raw_rows > 0:
      sheets_summary.append({'VC Sheet': ws_vc.title, 'Total Rows Scanned': sheet_raw_rows, 'Unique SKUs Kept': len(sheet_unique_skus), 'Duplicate/Blank Rows Merged': sheet_duplicates_merged, 'SC Product Type(s)': "Auto"})

  vc_rows_data, model_first_child_cells, sheet_pts_map = [], {}, {}
  for sku_norm_key in deduped_sku_order:
    entry = deduped_sku_store[sku_norm_key]
    sku_str, model_str, mapped_cells, sh_title = entry['sku_str'], entry['model_str'], entry['mapped_cells'], entry['sheet_title']
    if sc_desc_col and sc_desc_col in mapped_cells: mapped_cells[sc_desc_col] = _limit_desc_200(mapped_cells[sc_desc_col])
    if sc_price_col and ctx.get('selling_price') and sc_price_col not in mapped_cells: mapped_cells[sc_price_col] = ctx['selling_price']
    if sc_mrp_col and ctx.get('mrp') and sc_mrp_col not in mapped_cells: mapped_cells[sc_mrp_col] = int(ctx['mrp']) if str(ctx['mrp']).isdigit() else ctx['mrp']
    resolved_pt = resolve_sc_product_type_for_row(ws_sc, sc_prod_type_col, raw_prod_type=entry['raw_pt_val'], sheet_name=sh_title, item_type=entry['raw_itn_val'], model_name=model_str, item_name=entry['raw_title_val'] or sku_str)
    if sc_prod_type_col and resolved_pt:
      mapped_cells[sc_prod_type_col] = resolved_pt
      sheet_pts_map.setdefault(sh_title, set()).add(str(resolved_pt))
    if model_str and (model_str not in model_first_child_cells or len(mapped_cells) > len(model_first_child_cells[model_str])): model_first_child_cells[model_str] = mapped_cells
    vc_rows_data.append((sku_str, model_str, mapped_cells, sh_title))

  for s_item in sheets_summary:
    pts = sheet_pts_map.get(s_item['VC Sheet'])
    if pts: s_item['SC Product Type(s)'] = ", ".join(sorted(pts))
  if not vc_rows_data: return None, [], 0, []

  action_val = "Create or Replace (Full Update)"
  parent_val = pick_dropdown_value(ws_sc, sc_parentage_col, "Parent", "Parent") if sc_parentage_col else "Parent"
  child_val = pick_dropdown_value(ws_sc, sc_parentage_col, "Child", "Child") if sc_parentage_col else "Child"
  var_theme_val, asin_type_val, cond_new_val, fulfillment_default_val, qty_zero_val, shipping_tpl_val = "SIZE/COLOR", "ASIN", "New", "Fulfillment by Merchant (Default)", int(ctx.get('quantity', 0)), "Migrated Template"
  parent_position = ctx.get('parent_position', 'before_group')
  seen_models, unique_models_ordered, model_to_sheet = set(), [], {}
  for _, m_str, _, sh_title in vc_rows_data:
    if m_str and m_str not in seen_models: seen_models.add(m_str); unique_models_ordered.append(m_str); model_to_sheet[m_str] = sh_title

  preview, total_written = [], 0

  def apply_common_fixed_fields(row_idx):
    if sc_action_col: ws_sc.cell(row_idx, sc_action_col).value = action_val
    if sc_var_theme_col: ws_sc.cell(row_idx, sc_var_theme_col).value = var_theme_val
    if sc_condition_col: ws_sc.cell(row_idx, sc_condition_col).value = cond_new_val
    if sc_fulfillment_col: ws_sc.cell(row_idx, sc_fulfillment_col).value = fulfillment_default_val
    if sc_quantity_col: ws_sc.cell(row_idx, sc_quantity_col).value = qty_zero_val

  def write_parent_row(model_sku):
    nonlocal curr_r, total_written
    ref_cells = model_first_child_cells.get(model_sku) or (vc_rows_data[0][2] if vc_rows_data else {})
    for sc_c in parent_inherited_sc_cols:
      if sc_c in ref_cells and ref_cells[sc_c] not in (None, ''): ws_sc.cell(curr_r, sc_c).value = ref_cells[sc_c]
    ws_sc.cell(curr_r, sc_sku_col).value = model_sku
    if sc_prod_type_col and sc_prod_type_col in ref_cells: ws_sc.cell(curr_r, sc_prod_type_col).value = ref_cells[sc_prod_type_col]
    if sc_parentage_col: ws_sc.cell(curr_r, sc_parentage_col).value = parent_val
    if sc_model_name_col:
      ws_sc.cell(curr_r, sc_model_name_col).value = model_sku
      ws_sc.cell(curr_r, sc_model_name_col).fill = HIGHLIGHT_FILL; ws_sc.cell(curr_r, sc_model_name_col).font = HIGHLIGHT_FONT
    parent_title = ""
    if sc_title_col:
      parent_title = _build_parent_title_from_child(ref_cells.get(sc_title_col, ''), model_sku)
      ws_sc.cell(curr_r, sc_title_col).value = parent_title
    apply_common_fixed_fields(curr_r)
    preview.append({'Row': curr_r, 'Source VC Sheet': model_to_sheet.get(model_sku, ''), 'SKU': model_sku, 'Product Type': ref_cells.get(sc_prod_type_col, '') if sc_prod_type_col else '', 'Listing Action': action_val, 'Parentage Level': parent_val, 'Parent SKU': '', 'Variation Theme': var_theme_val, 'Product Id Type': '', 'Product Id': '', 'Shipping Template': '', 'Model Name': model_sku, 'Item Name': parent_title})
    curr_r += 1; total_written += 1

  def write_child_row(child_sku, model_str, mapped_cells, src_sheet_name=""):
    nonlocal curr_r, total_written
    for sc_c, val in mapped_cells.items(): ws_sc.cell(curr_r, sc_c).value = val
    ws_sc.cell(curr_r, sc_sku_col).value = child_sku
    if sc_parentage_col: ws_sc.cell(curr_r, sc_parentage_col).value = child_val
    if sc_parent_sku_col and model_str: ws_sc.cell(curr_r, sc_parent_sku_col).value = model_str
    if sc_model_name_col:
      if model_str: ws_sc.cell(curr_r, sc_model_name_col).value = model_str
      ws_sc.cell(curr_r, sc_model_name_col).fill = HIGHLIGHT_FILL; ws_sc.cell(curr_r, sc_model_name_col).font = HIGHLIGHT_FONT
    if sc_prod_id_type_col: ws_sc.cell(curr_r, sc_prod_id_type_col).value = asin_type_val
    if sc_shipping_tpl_col: ws_sc.cell(curr_r, sc_shipping_tpl_col).value = shipping_tpl_val
    apply_common_fixed_fields(curr_r)
    preview.append({'Row': curr_r, 'Source VC Sheet': src_sheet_name, 'SKU': child_sku, 'Product Type': mapped_cells.get(sc_prod_type_col, '') if sc_prod_type_col else '', 'Listing Action': action_val, 'Parentage Level': child_val, 'Parent SKU': model_str, 'Variation Theme': var_theme_val, 'Product Id Type': asin_type_val, 'Product Id': mapped_cells.get(sc_prod_id_val_col, '') if sc_prod_id_val_col else '', 'Shipping Template': shipping_tpl_val, 'Model Name': model_str, 'Item Name': mapped_cells.get(sc_title_col, '') if sc_title_col else ''})
    curr_r += 1; total_written += 1

  if parent_position == 'top':
    for m_sku in unique_models_ordered: write_parent_row(m_sku)
    for c_sku, m_str, m_cells, sh_title in vc_rows_data: write_child_row(c_sku, m_str, m_cells, sh_title)
  elif parent_position == 'bottom':
    for c_sku, m_str, m_cells, sh_title in vc_rows_data: write_child_row(c_sku, m_str, m_cells, sh_title)
    for m_sku in unique_models_ordered: write_parent_row(m_sku)
  else:
    inserted_parents = set()
    for c_sku, m_str, m_cells, sh_title in vc_rows_data:
      if m_str and m_str not in inserted_parents: inserted_parents.add(m_str); write_parent_row(m_str)
      write_child_row(c_sku, m_str, m_cells, sh_title)

  output = io.BytesIO()
  wb_sc.save(output); output.seek(0)
  return output.getvalue(), preview, total_written, sheets_summary

# ======================================================================
# ✨ AJIO PERFECT LISTING ENGINE
# ======================================================================
PL_COLOR_TOKENS = {'black', 'white', 'blue', 'lightblue', 'darkblue', 'skyblue', 'iceblue', 'navy', 'navyblue', 'green', 'olive', 'pista', 'sea-green', 'purple', 'lavender', 'pink', 'red', 'maroon', 'beige', 'biege', 'brown', 'coffee', 'khaki', 'grey', 'gray', 'charcoal', 'yellow', 'orange', 'peach', 'teal', 'wine', 'cream', 'rust', 'gold', 'silver'}
PL_COLOR_ABBR = {'black': 'Blk', 'white': 'Wht', 'blue': 'Blu', 'lightblue': 'LBlu', 'darkblue': 'DBlu', 'skyblue': 'SBlu', 'iceblue': 'IBlu', 'navy': 'Nvy', 'green': 'Grn', 'olive': 'Olv', 'pista': 'Pst', 'sea-green': 'SGrn', 'purple': 'Prp', 'lavender': 'Lav', 'pink': 'Pnk', 'red': 'Red', 'maroon': 'Mrn', 'beige': 'Beg', 'biege': 'Beg', 'brown': 'Brn', 'coffee': 'Cfe', 'khaki': 'Khk', 'grey': 'Gry', 'gray': 'Gry', 'charcoal': 'Chr', 'yellow': 'Ylw', 'orange': 'Orn', 'peach': 'Pch', 'teal': 'Tel', 'wine': 'Win', 'cream': 'Crm', 'ribbin': 'Rbn', 'teddy': 'Tdy'}

def pl_clean_sku_name(name): return clean_sku_name(name)
def pl_norm_h(s): return _norm_h(s)

def pl_extract_style_code(folder_name):
  clean = pl_clean_sku_name(folder_name)
  parts = clean.split('-')
  for i in range(len(parts) - 1, 0, -1):
    low_p = parts[i].lower()
    if low_p in PL_COLOR_TOKENS or low_p in COLOR_MISSPELLINGS: return '-'.join(parts[:i])
  return '-'.join(parts[:-1]) if (len(parts) >= 2 and not parts[-1].isdigit()) else clean

def pl_generate_safe_ean(clean_style, clean_sz):
  sz_suffix = f"-{clean_sz}"
  candidate = f"{clean_style}{sz_suffix}"
  if len(candidate) <= 18: return candidate
  new_parts = [PL_COLOR_ABBR[p.lower()] if p.lower() in PL_COLOR_ABBR else (str(int(p)) if p.isdigit() else p) for p in clean_style.split('-')]
  for sep in ("-", ""):
    cand = sep.join(new_parts) + sz_suffix
    if len(cand) <= 18: return cand
  return ("".join(new_parts))[:18 - len(sz_suffix)] + sz_suffix

PL_CATEGORY_OPTIONS = ['co-ord set', 'night suit', 'lounge wear', 'pajama set', 'saree', 'kurta', 'kurti', 'jeans', 'cargo', 'trousers', 'joggers', 'shorts', 'tshirt', 'shirt', 'jacket', 'hoodie', 'sweatshirt', 'dress', 'top', 'skirt', 'apparel']
PL_COLOR_FAMILY_MAP = {'maroon': 'Red', 'red': 'Red', 'wine': 'Red', 'darkred': 'Red', 'navy': 'Navy', 'navyblue': 'Navy', 'blue': 'Blue', 'lightblue': 'Blue', 'darkblue': 'Blue', 'iceblue': 'Blue', 'skyblue': 'Blue', 'teal': 'Teal', 'black': 'Black', 'charcoal': 'Grey', 'grey': 'Grey', 'gray': 'Grey', 'lightgrey': 'Grey', 'darkgrey': 'Grey', 'white': 'White', 'offwhite': 'Off White', 'green': 'Green', 'olive': 'Olive', 'pista': 'Sea Green', 'sea green': 'Sea Green', 'brown': 'Brown', 'coffee': 'Brown', 'beige': 'Beige', 'biege': 'Beige', 'khaki': 'Khaki', 'yellow': 'Yellow', 'mustard': 'Yellow', 'orange': 'Orange', 'peach': 'Peach', 'pink': 'Pink', 'purple': 'Purple', 'lavender': 'Purple'}

def pl_detect_color_from_name(text, image_bytes=None): return detect_color(text, image_bytes=image_bytes)
def pl_map_color_to_family(col_name):
  canon = _canonicalize_color(col_name)
  return PL_COLOR_FAMILY_MAP.get(str(canon or '').strip().lower(), canon or "Black")

def pl_get_dimensions_and_weight(category):
  cat = str(category or '').lower()
  if any(k in cat for k in ['jean', 'denim', 'cargo', 'heavy', 'jacket']): l, w, h, art_wt, pkg_wt = 28, 25, 3.0, 400, 430
  elif any(k in cat for k in ['saree', 'dress', 'kurta', 'kurti']): l, w, h, art_wt, pkg_wt = 26, 24, 2.5, 330, 360
  elif any(k in cat for k in ['tshirt', 't-shirt', 'top', 'shirt', 'kids']): l, w, h, art_wt, pkg_wt = 25, 24, 2.5, 300, 340
  else: l, w, h, art_wt, pkg_wt = 25, 24, 2.5, 350, 380
  return {'art_l': l, 'art_w': w, 'art_h': h, 'art_luom': 'CM', 'art_wt': art_wt, 'art_wtuom': 'GRAM', 'pkg_l': l, 'pkg_w': w, 'pkg_h': h, 'pkg_luom': 'CM', 'pkg_wt': pkg_wt, 'pkg_wtuom': 'GRAM'}

PL_HSN_INTELLIGENCE_MAP = {'saree': '54075200', 'kurti': '62042200', 'kurta': '62042200', 'night suit': '61083100', 'sleepwear': '61083100', 'lounge wear': '61083100', 'co-ord set': '61083100', 'jeans': '62034290', 'tshirt': '61099090', 'shirt': '62052000', 'trousers': '62034200', 'cargo': '62034200', 'joggers': '62034200', 'shorts': '62034200', 'dress': '62044200', 'hoodie': '61103000', 'sweatshirt': '61103000', 'jacket': '62014090', 'top': '62063000', 'kids': '61112000', 'apparel': '62042200'}
def pl_resolve_hsn_code(user_hsn, category):
  if user_hsn and str(user_hsn).strip(): return str(user_hsn).strip()
  cat = str(category or '').lower()
  return next((code for k, code in PL_HSN_INTELLIGENCE_MAP.items() if k in cat), '62042200')

def pl_get_garment_vision_prompt(folder_hint=""):
  cats = ", ".join(PL_CATEGORY_OPTIONS)
  hint_line = f'Folder/SKU name hint: "{folder_hint}".\n' if folder_hint else ""
  return "You are an expert e-commerce fashion cataloger. Inspect the MAIN GARMENT and output ONLY valid JSON:\n" + hint_line + '{"category": "<One of: ' + cats + '>", "gender": "<Women, Men, Boys, Girls>", "color": "<Black, Blue, Navy, LightBlue, DarkBlue, SkyBlue, Beige, Grey, LightGrey, DarkGrey, White, Off White, Olive, Green, Pista, Sea Green, Brown, Coffee, Khaki, Maroon, Red, Yellow, Orange, Lavender, Purple, Teal, Wine, Pink>", "pocket": "<Side Pocket / No Pocket / Cargo Pocket / Kangaroo Pocket>", "material": "<Cotton, Cotton Blend, Rayon, Fleece, Polyester, Denim, Crepe, Georgette, Silk, Knit>", "fabric_detail": "<100% Cotton, Cotton Blend, Soft Fleece, Rayon>", "sleeve_length": "<Short sleeve / Full-length sleeve / 3/4th sleeve / Sleeveless>", "topwear_type": "<Top / Tunic / Shirt / Hoodie / Sweatshirt / Kurta / Kurti>", "bottomwear_type": "<Capris / Pyjamas / Shorts / Trousers/Pants / Jeans>", "waist_type": "<Flexi waist / Elasticated waist / Fixed waist>", "pattern": "<Solid / Printed / Washed / Striped / Checked / Floral / Quilted>"}'

def pl_get_size_chart_prompt():
  return 'Extract garment size chart measurements in inches as JSON: {"bottomwear_type": "<Capris or Pyjamas>", "sleeve_length": "<Short sleeve or Full-length sleeve>", "measurements": {"XS": {"bust": 37, "front_length": 25, "waist": 22, "hip": 38, "inseam": 28}, "S": {"bust": 39, "front_length": 25, "waist": 24, "hip": 40, "inseam": 28}, "M": {"bust": 41, "front_length": 26, "waist": 26, "hip": 42, "inseam": 28}, "L": {"bust": 43, "front_length": 26, "waist": 28, "hip": 45, "inseam": 29}, "XL": {"bust": 45, "front_length": 27, "waist": 30, "hip": 48, "inseam": 29}, "XXL": {"bust": 47, "front_length": 27, "waist": 32, "hip": 51, "inseam": 29}}}'

def pl_call_gemini_api(prompt, image_bytes, img_name):
  if not GEMINI_API_KEY or not image_bytes: return None
  b64 = base64.b64encode(image_bytes).decode('utf-8')
  try:
    res = requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_VISION_MODEL}:generateContent?key={GEMINI_API_KEY}", json={"contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": guess_mime(img_name), "data": b64}}]}], "generationConfig": {"temperature": 0.0, "maxOutputTokens": 1500}}, timeout=40)
    if res.status_code == 200:
      cands = res.json().get('candidates') or []
      if cands: return extract_json_from_text(''.join(str(p.get('text', '')) for p in (cands[0].get('content') or {}).get('parts') or []))
  except Exception: pass
  return None

def pl_call_nvidia_api(prompt, image_bytes, img_name):
  if not NVIDIA_API_KEY or not image_bytes: return None
  data_url = f"data:{guess_mime(img_name)};base64,{base64.b64encode(image_bytes).decode('utf-8')}"
  try:
    res = requests.post("https://integrate.api.nvidia.com/v1/chat/completions", headers={"Authorization": f"Bearer {NVIDIA_API_KEY}", "Content-Type": "application/json"}, json={"model": NVIDIA_MODEL, "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": data_url}}]}], "temperature": 0.1, "max_tokens": 1500}, timeout=50)
    if res.status_code == 200: return extract_json_from_text(res.json()['choices'][0]['message']['content'])
  except Exception: pass
  return None

def pl_call_groq_api(prompt, image_bytes, img_name):
  if not GROQ_API_KEY or not image_bytes: return None
  data_url = f"data:{guess_mime(img_name)};base64,{base64.b64encode(image_bytes).decode('utf-8')}"
  try:
    res = requests.post("https://api.groq.com/openai/v1/chat/completions", headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}, json={"model": GROQ_VISION_MODEL, "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": data_url}}]}], "temperature": 0.1, "max_tokens": 1500}, timeout=40)
    if res.status_code == 200: return extract_json_from_text(res.json()['choices'][0]['message']['content'])
  except Exception: pass
  return None

def pl_fallback_ocr_size_chart(image_bytes):
  return {'bottomwear_type': 'Pyjamas', 'sleeve_length': 'Short sleeve', 'measurements': {'XS': {'bust': 37, 'front_length': 25, 'waist': 22, 'hip': 38, 'inseam': 28}, 'S': {'bust': 39, 'front_length': 25, 'waist': 24, 'hip': 40, 'inseam': 28}, 'M': {'bust': 41, 'front_length': 26, 'waist': 26, 'hip': 42, 'inseam': 28}, 'L': {'bust': 43, 'front_length': 26, 'waist': 28, 'hip': 45, 'inseam': 29}, 'XL': {'bust': 45, 'front_length': 27, 'waist': 30, 'hip': 48, 'inseam': 29}, 'XXL': {'bust': 47, 'front_length': 27, 'waist': 32, 'hip': 51, 'inseam': 29}, '2XL': {'bust': 47, 'front_length': 27, 'waist': 32, 'hip': 51, 'inseam': 29}}}

def pl_multi_ai_vision_detect(image_bytes, img_name, folder_hint=""):
  prompt = pl_get_garment_vision_prompt(folder_hint=folder_hint)
  for fn, provider in [(pl_call_gemini_api, "Gemini"), (pl_call_nvidia_api, "NVIDIA"), (pl_call_groq_api, "Groq")]:
    res = fn(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('color'):
      res['color'] = canonicalize_ai_color(res.get('color'), folder_name=folder_hint, image_bytes=image_bytes)
      return res, provider
  return {'color': detect_color(folder_hint, image_bytes=image_bytes)}, "Pixel/Folder-Fallback"

def pl_multi_ai_size_chart_detect(image_bytes, img_name):
  prompt = pl_get_size_chart_prompt()
  for fn, provider in [(pl_call_gemini_api, "Gemini"), (pl_call_nvidia_api, "NVIDIA"), (pl_call_groq_api, "Groq")]:
    res = fn(prompt, image_bytes, img_name)
    if res and isinstance(res, dict) and res.get('measurements'): return res, provider
  return pl_fallback_ocr_size_chart(image_bytes), "OCR-Fallback"

def pl_read_and_analyze_template(wb):
  info = {'sheets': wb.sheetnames, 'styles_sheet': None, 'category_breadcrumb': '', 'detected_gender': 'Women', 'detected_category': 'Apparel', 'col_to_dropdowns': {}, 'headers': {}, 'mandatory_cols': []}
  if 'Seller Information Sheet' in wb.sheetnames:
    target = wb['Seller Information Sheet'].cell(2, 1).value
    if target and str(target).strip() in wb.sheetnames: info['styles_sheet'] = str(target).strip()
  if not info['styles_sheet']:
    info['styles_sheet'] = next((s for s in wb.sheetnames if 'styles' in s.lower() and 'instruction' not in s.lower()), wb.sheetnames[0])
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
    if sum(1 for v in vals if any(k in v for k in keywords)) > 3: header_row = r; break
  info['header_row'] = header_row
  info['headers'] = {c: str(ws_styles.cell(header_row, c).value or '').strip() for c in range(1, ws_styles.max_column + 1)}
  info['mandatory_cols'] = [c for c, h in info['headers'].items() if h.startswith('*')]
  defined_ranges = {name.lower(): (getattr(dn, 'attr_text', '') or getattr(dn, 'value', '') or '') for name, dn in wb.defined_names.items()} if hasattr(wb, 'defined_names') else {}
  col_to_dropdowns = {}
  try:
    for dv in ws_styles.data_validations.dataValidation:
      if dv.type == 'list' and dv.formula1:
        f1 = str(dv.formula1).strip().replace('"', '')
        vals = []
        if ',' in f1: vals = [x.strip() for x in f1.split(',') if x.strip()]
        elif f1.lower() in defined_ranges and '!' in defined_ranges[f1.lower()]:
          s_name, cell_range = defined_ranges[f1.lower()].split('!')
          s_name = s_name.strip("'")
          if s_name in wb.sheetnames:
            for row in wb[s_name][cell_range.replace('$', '')]:
              if not isinstance(row, tuple): row = (row,)
              for cell in row:
                if cell.value is not None and str(cell.value).strip(): vals.append(str(cell.value).strip())
        for rng in dv.sqref.ranges:
          for col_idx in range(rng.min_col, rng.max_col + 1): col_to_dropdowns[col_idx] = vals
  except Exception: pass
  info['col_to_dropdowns'] = col_to_dropdowns
  return info

def pl_pick_dropdown_value_safe(info, col_idx, desired, fallback=None):
  if not col_idx: return desired or fallback
  options = info.get('col_to_dropdowns', {}).get(col_idx)
  if not options: return desired if desired is not None else fallback
  cand = str(desired if desired is not None else (fallback or '')).strip()
  if not cand: return None
  cl = cand.lower()
  for o in options:
    if str(o).strip().lower() == cl: return o
  for o in options:
    if cl in str(o).strip().lower() or str(o).strip().lower() in cl: return o
  close = difflib.get_close_matches(cand, options, n=1, cutoff=0.4)
  return close[0] if close else options[0]

def pl_upload_to_imgbb_cached(image_bytes, api_key, cache): return upload_to_imgbb_cached(image_bytes, api_key, cache)

def pl_build_dynamic_seo_content(category, gender, brand, fabric, pattern, keywords=""):
  g_title = "Women's" if str(gender).lower().startswith('w') else ("Men's" if str(gender).lower().startswith('m') else "Unisex")
  b_prefix = f"{brand} " if brand else ""
  cat_label = category.title() if category else "Apparel"
  kw_str = keywords if keywords else "relaxed fit and all-day comfort"
  desc = f"{b_prefix}{g_title} {cat_label} in breathable {fabric} fabric with an elegant {pattern.lower()} pattern. Designed for {kw_str} and effortless daily wear."
  highlight = f"Premium breathable {fabric} fabric | Comfortable {g_title.lower()} fit | Elegant {pattern.lower()} design | Durable stitching | Easy machine wash"
  add_info = "Wash Care: Machine wash cold with similar colors. Do not bleach. Warm iron if needed."
  return _limit_desc_200(desc), highlight[:500], add_info[:500]

def pl_run_perfect_ajio_listing(template_file_or_bytes, zip_images_data, detected_size_charts, ctx):
  name = getattr(template_file_or_bytes, 'name', 'template.xlsx')
  wb = openpyxl.load_workbook(template_file_or_bytes, keep_vba=name.lower().endswith('.xlsm'))
  info = pl_read_and_analyze_template(wb)
  ws = wb[info['styles_sheet']]
  headers, header_row = info['headers'], info['header_row']

  def find_col(candidates):
    for cand in candidates:
      c_norm = pl_norm_h(cand)
      for c, h in headers.items():
        if pl_norm_h(h) == c_norm: return c
    return None

  cat_detected, gender_detected = info['detected_category'], info['detected_gender']
  col_style, col_desc, col_sku, col_brand, col_ean, col_td, col_mrp, col_hsn = find_col(['*Style Code', 'Style Code']), find_col(['*Style Description', 'Style Description']), find_col(['*Item SKU', 'Item SKU', 'SKU']), find_col(['*Brand', 'Brand']), find_col(['*EAN', 'EAN', 'Barcode']), find_col(['*TD', 'TD']), find_col(['*MRP', 'MRP', 'Maximum Retail Price']), find_col(['*HSN', 'HSN', 'HSN Code'])
  col_prod_grp, col_fash_grp, col_season_cd, col_season_yr, col_size, col_std_size = find_col(['*Product Groups', 'Product Groups']), find_col(['*Fashion Groups', 'Fashion Groups']), find_col(['*Season Code', 'Season Code']), find_col(['*Season Year', 'Season Year']), find_col(['*Size', 'Size']), find_col(['*StandardSize', 'StandardSize', 'Standard Size'])
  col_art_h, col_art_l, col_art_w, col_art_luom, col_art_wt, col_art_wtuom = find_col(['*articleDimensionsUnitHeight']), find_col(['*articleDimensionsUnitLength']), find_col(['*articleDimensionsUnitWidth']), find_col(['*articleDimensionsUnitLengthUOM']), find_col(['*articleDimensionsUnitWeight']), find_col(['*articleDimensionsUnitWeightUOM'])
  col_pkg_h, col_pkg_l, col_pkg_w, col_pkg_luom, col_pkg_wt, col_pkg_wtuom = find_col(['*packageDimensionsHeight']), find_col(['*packageDimensionsLength']), find_col(['*packageDimensionsWidth']), find_col(['*packageDimensionsLengthUOM']), find_col(['*packageDimensionsWeight']), find_col(['*packageDimensionsWeightUOM'])
  col_add_info1, col_comp_count, col_country, col_highlight, col_mfg_by, col_mkt_by, col_net_qty, col_pkg_contains = find_col(['Additional Information 1']), find_col(['*Component Count']), find_col(['*Country of Origin']), find_col(['Highlight']), find_col(['Manufactured By']), find_col(['*Marketed By']), find_col(['*Net Quantity']), find_col(['*Package Contains'])
  col_color_fam, col_color_shade, col_fab_detail, col_fab_type, col_pattern, col_pri_color, col_size_format, col_size_group, col_wash_care, col_plus_size, col_prod_name, col_sleeve_len, col_waist_type, col_top_type, col_bottom_type, col_style_type = find_col(['*Color Family']), find_col(['*Color Shade']), find_col(['Fabric Detail']), find_col(['*Fabric Type']), find_col(['*Pattern']), find_col(['*Primary Color']), find_col(['*Size Format']), find_col(['*Size Group']), find_col(['*Wash Care']), find_col(['*Plus Size Store']), find_col(['*Product Name']), find_col(['Sleeve Length']), find_col(['Waist Type']), find_col(['*Topwear Type']), find_col(['Bottomwear Type']), find_col(['*Style Type'])
  col_meas_type, col_to_fit_bust, col_bust, col_front_len, col_to_fit_waist, col_waist, col_to_fit_hip, col_hip, col_inseam = find_col(['*Measurement Type']), find_col(['*To Fit Bust']), find_col(['*Bust']), find_col(['*Front Length']), find_col(['*To Fit Waist']), find_col(['*Waist']), find_col(['*To Fit Hip']), find_col(['*Hips']), find_col(['*Inseam'])

  img_cols = [find_col(['*MODEL' if i == 1 else f'MODEL{i}']) for i in range(1, 15)]
  img_cols = [c for c in img_cols if c]
  dims_data, hsn_val = pl_get_dimensions_and_weight(cat_detected), pl_resolve_hsn_code(ctx.get('hsn'), cat_detected)

  curr_r = header_row + 1
  while ws.cell(curr_r, col_style or 1).value is not None: curr_r += 1
  total_added, preview, sizes, season_year = 0, [], (ctx.get('sizes') or ['S', 'M', 'L', 'XL', '2XL']), str(datetime.now().year)

  def set_val(r, c, val, dropdown=False, fill=None, font=None):
    if not c or val in (None, ''): return
    cell = ws.cell(r, c)
    cell.value = pl_pick_dropdown_value_safe(info, c, val, val) if dropdown else val
    if fill: cell.fill = fill
    if font: cell.font = font

  for fld, img_tuples in zip_images_data.items():
    base_sku, style_code = pl_clean_sku_name(fld), pl_extract_style_code(fld)
    first_img_bytes = img_tuples[0][1] if img_tuples else None
    ai_data = ctx.get('vision_results', {}).get(base_sku, {})
    det_col = canonicalize_ai_color(extract_explicit_color_from_text(fld) or ai_data.get('color') or pl_detect_color_from_name(fld, image_bytes=first_img_bytes), folder_name=fld, image_bytes=first_img_bytes)
    col_fam = pl_map_color_to_family(det_col)
    mat_val = ai_data.get('material') or "Cotton"
    fab_detail = ai_data.get('fabric_detail') or f"100% {mat_val}"
    pattern_val = ai_data.get('pattern') or ("Quilted" if "quilt" in fld.lower() else "Solid")
    sc_data = detected_size_charts.get(fld) or detected_size_charts.get('default') or {}
    sleeve_val = sc_data.get('sleeve_length') or ai_data.get('sleeve_length') or "Short sleeve"
    bottom_val = sc_data.get('bottomwear_type') or ("Capris" if "capri" in fld.lower() else "Pyjamas")
    measurements = sc_data.get('measurements') or {}
    desc_text, high_text, add_text = pl_build_dynamic_seo_content(cat_detected, gender_detected, ctx.get('brand', ''), mat_val, pattern_val, ctx.get('keywords', ''))
    img_urls = list(ctx.get('live_links_cache', {}).get(fld, []))
    if ctx.get('size_chart_urls', {}).get(fld) and ctx['size_chart_urls'][fld] not in img_urls: img_urls.append(ctx['size_chart_urls'][fld])

    preview.append({'Style Code': style_code, 'Base SKU': base_sku, 'Category': cat_detected, 'Color': det_col, 'Color Family': col_fam, 'Fabric': mat_val, 'Pattern': pattern_val, 'Images Count': len(img_urls)})

    for sz in sizes:
      clean_sz = pl_clean_sku_name(sz)
      full_sku, ean_val = f"{base_sku}-{clean_sz}", pl_generate_safe_ean(base_sku, clean_sz)
      sz_meas = measurements.get(clean_sz) or measurements.get(clean_sz.upper()) or measurements.get({'2XL': 'XXL', 'XXL': '2XL'}.get(clean_sz.upper())) or {}
      bust_val, front_len_val, waist_val, hip_val, inseam_val = sz_meas.get('bust', 37), sz_meas.get('front_length', 25), sz_meas.get('waist', 22), sz_meas.get('hip', 38), sz_meas.get('inseam', 28)

      set_val(curr_r, col_style, style_code, fill=HIGHLIGHT_FILL, font=HIGHLIGHT_FONT)
      set_val(curr_r, col_desc, desc_text); set_val(curr_r, col_sku, full_sku)
      set_val(curr_r, col_brand, ctx.get('brand', '')); set_val(curr_r, col_ean, ean_val); set_val(curr_r, col_td, ctx.get('td', 50))
      if col_mrp:
        if ctx.get('mrp'): ws.cell(curr_r, col_mrp).value = int(ctx['mrp']) if str(ctx['mrp']).isdigit() else ctx['mrp']
        else: ws.cell(curr_r, col_mrp).fill = HIGHLIGHT_FILL; ws.cell(curr_r, col_mrp).font = HIGHLIGHT_FONT
      set_val(curr_r, col_hsn, hsn_val)
      set_val(curr_r, col_prod_grp, "Evening", dropdown=True); set_val(curr_r, col_fash_grp, "Fashion", dropdown=True)
      set_val(curr_r, col_season_cd, "Core", dropdown=True); set_val(curr_r, col_season_yr, season_year)
      set_val(curr_r, col_size, clean_sz); set_val(curr_r, col_std_size, clean_sz)
      set_val(curr_r, col_art_h, dims_data['art_h']); set_val(curr_r, col_art_l, dims_data['art_l']); set_val(curr_r, col_art_w, dims_data['art_w']); set_val(curr_r, col_art_luom, dims_data['art_luom']); set_val(curr_r, col_art_wt, dims_data['art_wt']); set_val(curr_r, col_art_wtuom, dims_data['art_wtuom'])
      set_val(curr_r, col_pkg_h, dims_data['pkg_h']); set_val(curr_r, col_pkg_l, dims_data['pkg_l']); set_val(curr_r, col_pkg_w, dims_data['pkg_w']); set_val(curr_r, col_pkg_luom, dims_data['pkg_luom']); set_val(curr_r, col_pkg_wt, dims_data['pkg_wt']); set_val(curr_r, col_pkg_wtuom, dims_data['pkg_wtuom'])
      set_val(curr_r, col_add_info1, add_text); set_val(curr_r, col_comp_count, 1); set_val(curr_r, col_country, "India", dropdown=True); set_val(curr_r, col_highlight, high_text)
      set_val(curr_r, col_mfg_by, ctx.get('mfg', '')); set_val(curr_r, col_mkt_by, ctx.get('mfg', '')); set_val(curr_r, col_net_qty, 1); set_val(curr_r, col_pkg_contains, "1")
      set_val(curr_r, col_color_fam, col_fam, dropdown=True); set_val(curr_r, col_color_shade, "Light", dropdown=True); set_val(curr_r, col_fab_detail, fab_detail); set_val(curr_r, col_fab_type, mat_val, dropdown=True)
      set_val(curr_r, col_pattern, pattern_val, dropdown=True); set_val(curr_r, col_pri_color, det_col, dropdown=True); set_val(curr_r, col_size_format, "UNI", dropdown=True); set_val(curr_r, col_size_group, "Regular", dropdown=True)
      set_val(curr_r, col_wash_care, "Machine wash", dropdown=True); set_val(curr_r, col_plus_size, "No", dropdown=True); set_val(curr_r, col_prod_name, "Knitted Top", dropdown=True)
      set_val(curr_r, col_sleeve_len, sleeve_val, dropdown=True); set_val(curr_r, col_waist_type, "Flexi waist", dropdown=True); set_val(curr_r, col_top_type, "Top", dropdown=True); set_val(curr_r, col_bottom_type, bottom_val, dropdown=True); set_val(curr_r, col_style_type, "2-Piece Sets", dropdown=True)
      if col_meas_type:
        set_val(curr_r, col_meas_type, "Garment Measurement", dropdown=True); set_val(curr_r, col_to_fit_bust, bust_val); set_val(curr_r, col_bust, bust_val); set_val(curr_r, col_front_len, front_len_val)
        set_val(curr_r, col_to_fit_waist, waist_val + 2); set_val(curr_r, col_waist, waist_val); set_val(curr_r, col_to_fit_hip, hip_val); set_val(curr_r, col_hip, hip_val); set_val(curr_r, col_inseam, inseam_val)
      for idx_img, url in enumerate(img_urls):
        if idx_img < len(img_cols): ws.cell(curr_r, img_cols[idx_img]).value = url
      curr_r += 1; total_added += 1

  output = io.BytesIO()
  wb.save(output); output.seek(0)
  return output.getvalue(), preview, total_added, info['styles_sheet']

# ======================================================================
# ✨ UNIFIED DASHBOARD & VC ➜ SC UI
# ======================================================================
def _mime_for(name):
  return "application/vnd.ms-excel.sheet.macroEnabled.12" if str(name).lower().endswith('.xlsm') else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def render_convert_ui():
  st.subheader("🔄 Convert: Amazon Vendor Central ➜ Amazon Seller Central")
  c1, c2 = st.columns(2)
  with c1: vc_file = st.file_uploader("1️⃣ Filled Vendor Central File (.xlsx / .xlsm):", type=["xlsx", "xlsm"], key="vc_src")
  with c2: sc_file = st.file_uploader("2️⃣ Blank Seller Central Template (.xlsx / .xlsm):", type=["xlsx", "xlsm"], key="sc_dst")
  c3, c4, c5, c6 = st.columns(4)
  with c3: default_qty = st.number_input("Quantity (IN) [Fixed = 0]:", min_value=0, value=0)
  with c4: selling_price = st.text_input("Your Price INR (optional):", value="", placeholder="e.g. 799")
  with c5: mrp_override = st.text_input("Fallback MRP (optional):", value="", placeholder="VC me MRP na ho to")
  with c6: custom_parent_model = st.text_input("Common Parent SKU / Model Name (Optional):", value="", placeholder="Blank = Auto per sheet/SKU")
  parent_pos_choice = st.radio("Parent Row Position:", options=["Har Model Group ke Pehle (Recommended)", "Sabse Upar (Top Rows me)", "Sabse Niche (Bottom Rows me)"], index=0, horizontal=True)
  pos_map = {"Har Model Group ke Pehle (Recommended)": "before_group", "Sabse Upar (Top Rows me)": "top", "Sabse Niche (Bottom Rows me)": "bottom"}

  if st.button("🔄 Convert VC ➜ SC & Generate File", type="primary"):
    if not vc_file or not sc_file: st.error("Dono files upload karein!")
    else:
      try:
        ctx = {'quantity': int(default_qty), 'selling_price': selling_price.strip(), 'mrp': mrp_override.strip(), 'sc_parent_model': custom_parent_model.strip(), 'parent_position': pos_map[parent_pos_choice]}
        with st.spinner("Converting Vendor Central rows to Seller Central `Template` sheet..."):
          out, preview, count, sheets_summary = convert_vc_to_sc(vc_file, sc_file, ctx)
        if not count or out is None: st.warning("⚠️ VC file me koi data row nahi mili.")
        else:
          st.success(f"🎉 Converted **{count}** total rows into Seller Central `Template` sheet!")
          if sheets_summary: st.dataframe(sheets_summary, use_container_width=True)
          if preview: st.dataframe(preview, use_container_width=True)
          base, ext = os.path.splitext(sc_file.name)
          dl_name = f"{base}_from_VC{ext}"
          st.download_button(f"📥 Download Converted Seller Central File ({dl_name})", data=out, file_name=dl_name, mime=_mime_for(dl_name))
      except Exception as e:
        st.error(f"Error: {e}")
        with st.expander("Technical details"): st.code(traceback.format_exc())

def render_listing_ui():
  is_ajio, is_meesho = (marketplace == "Ajio"), (marketplace == "Meesho")
  if 'pl_vision_results' not in st.session_state: st.session_state['pl_vision_results'] = {}
  st.success(f"✨ **{marketplace} Listing Engine Active** — Multi-AI Vision (Gemini ➜ NVIDIA ➜ Groq + Pixel Color Verification) | Descriptions <= 200 Chars | Highlighted Style Code / Model Name")

  pl_folders, pl_zip_images, pl_size_charts = [], {}, {}
  st.subheader("1. Product Images, SKUs & Size Chart (ZIP Upload)")
  col_in1, col_in2 = st.columns(2)
  with col_in1:
    st.write("##### 📁 Product Folders Ka ZIP Upload")
    zip_file = st.file_uploader("Product Folders (.zip):", type=["zip"], key="pl_zip_in")
    if zip_file:
      try:
        folder_to_imgs, found_charts = parse_product_zip(zip_file)
        pl_size_charts.update(found_charts)
        pl_folders = sorted(folder_to_imgs.keys())
        for fld in pl_folders: pl_zip_images[fld] = folder_to_imgs[fld]
        if not pl_folders: st.warning("⚠️ ZIP me koi product folder nahi mila.")
        else:
          st.success(f"✅ ZIP ke andar Total **{len(pl_folders)}** Product Folders (SKUs) Detect Hue:")
          st.code(", ".join([clean_sku_name(f) for f in pl_folders]))
        if pl_size_charts: st.info(f"📐 **{len(pl_size_charts)} Size Charts Detect Hue:** {', '.join(set(v['filename'] for v in pl_size_charts.values()))}")
      except Exception as e: st.error(f"ZIP error: {e}")
  with col_in2:
    st.write("##### 🌐 Live CDN & Cloud Hosting")
    auto_cdn = st.checkbox("Images ko auto CDN links (ImgBB) me convert karein", value=True, key="pl_auto_cdn")
    st.caption("⚡ Direct CDN Link Active: ImgBB se direct public URLs auto-generate honge.")

  dropbox_url, manual_skus = "", ""
  if not is_ajio:
    with st.expander("📎 Optional: Dropbox Link / Manual SKUs (ZIP na ho to)"):
      dropbox_url = st.text_input("Dropbox Shared Folder Link (Optional):", key="pl_dropbox", placeholder="https://www.dropbox.com/scl/fo/xyz/...")
      manual_skus = st.text_area("Manual SKUs List (Agar ZIP nahi hai):", key="pl_manual_skus", placeholder="SKU-1\nSKU-2", height=90)

  st.divider()
  st.subheader("2. Marketplace Template Upload")
  tpl_file = st.file_uploader(f"{marketplace} Template (.xlsx / .xlsm):", type=["xlsx", "xlsm"], key="pl_tpl_in")

  if tpl_file and is_ajio:
    try:
      wb_check = openpyxl.load_workbook(tpl_file, data_only=True)
      t_info = pl_read_and_analyze_template(wb_check)
      tpl_file.seek(0)
      c_t1, c_t2, c_t3 = st.columns(3)
      with c_t1: st.metric("Target Styles Sheet", t_info['styles_sheet'])
      with c_t2: st.metric("Template Columns", len(t_info['headers']))
      with c_t3: st.metric("Mandatory Columns (*)", len(t_info['mandatory_cols']))
      if t_info['category_breadcrumb']: st.info(f"📌 **Category Breadcrumb:** `{t_info['category_breadcrumb']}`")
    except Exception as e: st.warning(f"Template parsing warning: {e}")

  if tpl_file and is_meesho:
    try:
      wb_mcheck = openpyxl.load_workbook(tpl_file, data_only=True)
      m_info = analyze_meesho_template(wb_mcheck)
      tpl_file.seek(0)
      if m_info:
        c_m1, c_m2, c_m3 = st.columns(3)
        with c_m1: st.metric("Meesho Template Sheet", m_info['sheet_name'])
        with c_m2: st.metric("Total Columns", len(m_info['headers']))
        with c_m3: st.metric("Compulsory Fields (*)", len(m_info['compulsory_cols']))
        st.info(f"📌 **Detected Template:** `{m_info['banner']}` | **Recommended Wrong/Defective Return Price:** `Meesho Price - Rs.{m_info['rec_return_diff']}`")
    except Exception as e: st.warning(f"Meesho template check warning: {e}")

  st.divider()
  st.subheader("3. Brand, MRP, Legal & Sizes")
  c_b1, c_b2 = st.columns([1, 2])
  with c_b1:
    size_preset = st.radio("Sizes Preset Chunein (Choose Size Format):", ["Alpha Sizes (S, M, L, XL, 2XL)", "Numeric / Waist Sizes (28, 30, 32, 34, 36, 38)"], index=0, key="pl_size_preset")
    preset_val = "S, M, L, XL, 2XL" if size_preset.startswith("Alpha") else "28, 30, 32, 34, 36, 38"
    sizes_input = st.text_input("Sizes (comma-separated): Apne mutabiq add ya kam karein", value=preset_val, key=f"pl_sizes_{size_preset}")
    brand_name = st.text_input("Brand Name:", value="", key="pl_brand", placeholder="e.g. Poshax / PUMA")
    mrp_value = st.text_input("MRP (optional):", value="", key="pl_mrp", placeholder="e.g. 999 (Khali chhodne par Yellow Highlight hoga)")
    hsn_code = st.text_input("HSN Code (optional):", value="", key="pl_hsn", placeholder="e.g. 62052000")
  with c_b2:
    mfg_address = st.text_area("Manufactured By & Marketed By (Complete Address with Pincode):", value="", key="pl_mfg", placeholder="e.g. PLOT NO 41, DEVAM AMB, RING ROAD -300FTRO PARK, NEAR VEDANT EMBRO PARK, Kosad, Surat - 394107", height=130)
    custom_keywords = st.text_area("SEO & Trending Keywords (Optional):", value="", key="pl_kw", placeholder="e.g. casual streetwear, relaxed fit, durable quality", height=70)

  sc_quantity, sc_selling_price, sc_parent_model = 0, "", ""
  meesho_return_price, meesho_mfg_name, meesho_mfg_pin, meesho_gst, meesho_weight = "", "JB Fashion", "394107", 5, ""

  if marketplace == "Amazon - Seller Central":
    with st.expander("🛒 Seller Central Extra Fields (Optional)", expanded=False):
      c_sc1, c_sc2, c_sc3 = st.columns(3)
      with c_sc1: sc_quantity = st.number_input("Default Stock Quantity (0 Fixed):", min_value=0, value=0, key="pl_sc_qty")
      with c_sc2: sc_selling_price = st.text_input("Selling Price (optional):", key="pl_sc_price", placeholder="e.g. 899")
      with c_sc3: sc_parent_model = st.text_input("Common Parent SKU / Model Name (Optional):", key="pl_sc_parent", placeholder="e.g. H-63-66")

  if is_meesho:
    with st.expander("🛍️ Meesho Pricing, GST & Manufacturer/Packer Fields", expanded=True):
      cm1, cm2, cm3, cm4 = st.columns(4)
      with cm1: sc_selling_price = st.text_input("Meesho Price (Selling Price):", value="499", key="pl_meesho_price", placeholder="e.g. 499")
      with cm2: meesho_return_price = st.text_input("Wrong/Defective Return Price (Blank = Auto):", value="", key="pl_meesho_ret_price", placeholder="Auto (Meesho Price - 28/34)")
      with cm3: meesho_gst = st.selectbox("GST %:", [5, 12, 18], index=0, key="pl_meesho_gst")
      with cm4: sc_quantity = st.number_input("Inventory (Stock):", min_value=0, value=0, key="pl_meesho_inv")
      cm5, cm6, cm7 = st.columns(3)
      with cm5: meesho_mfg_name = st.text_input("Manufacturer & Packer Name:", value="JB Fashion", key="pl_meesho_mfg_name")
      with cm6: meesho_mfg_pin = st.text_input("Manufacturer & Packer Pincode:", value="394107", key="pl_meesho_mfg_pin")
      with cm7: meesho_weight = st.text_input("Net Weight (gms) [Blank = Auto 300/350]:", value="", key="pl_meesho_wt", placeholder="e.g. 300")

  st.divider()

  if st.button("🚀 Perfect Listing File Generate Karein", type="primary", key="pl_generate"):
    manual_list = [s.strip() for s in manual_skus.split('\n') if s.strip()] if not is_ajio else []
    if not tpl_file: st.error("Kripya Marketplace Template Excel file upload karein!")
    elif not pl_folders and not manual_list: st.error("Kripya Product ZIP file upload karein!")
    else:
      sizes_list = [s.strip() for s in sizes_input.split(',') if s.strip()]
      try:
        tpl_file.seek(0)
        if is_ajio:
          ctx = {'sizes': sizes_list, 'brand': brand_name.strip(), 'mrp': mrp_value.strip(), 'mfg': mfg_address.strip(), 'hsn': hsn_code.strip(), 'keywords': custom_keywords.strip(), 'td': int(default_td), 'live_links_cache': {}, 'size_chart_urls': {}, 'vision_results': st.session_state['pl_vision_results']}
          with st.spinner("🤖 1/3: AI Vision: Colors, Fabric & Sleeve length detect ho rahe hain..."):
            for fld in pl_folders:
              base_sku = pl_clean_sku_name(fld)
              if base_sku not in ctx['vision_results'] and pl_zip_images.get(fld):
                v_res, provider = pl_multi_ai_vision_detect(pl_zip_images[fld][0][1], pl_zip_images[fld][0][0], folder_hint=fld)
                if v_res: v_res['ai_source'] = provider; ctx['vision_results'][base_sku] = v_res
          size_chart_measurements = {}
          with st.spinner("📐 2/3: AI Vision size charts ko read karke measurements extract kar raha hai..."):
            for fld, sc_info in pl_size_charts.items():
              sc_res, _ = pl_multi_ai_size_chart_detect(sc_info['bytes'], sc_info['filename'])
              if sc_res:
                size_chart_measurements[fld] = sc_res
                if 'default' not in size_chart_measurements: size_chart_measurements['default'] = sc_res
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
            out_bytes, preview_data, total_added, target_sheet = pl_run_perfect_ajio_listing(tpl_file, pl_zip_images, size_chart_measurements, ctx)
          st.success(f"🎉 Listing complete! Total **{total_added}** variants add hue (`{target_sheet}`).")
          if preview_data: st.dataframe(preview_data, use_container_width=True)
          base_name, ext = os.path.splitext(tpl_file.name)
          out_ext = ext if ext.lower() in ['.xlsm', '.xlsx'] else '.xlsx'
          out_name = f"{base_name}_PERFECT_LISTED{out_ext}"
          st.download_button("📥 Download Filled Marketplace File", data=out_bytes, file_name=out_name, mime=_mime_for(out_name), type="primary")
        else:
          folders = pl_folders if pl_folders else manual_list
          valid_keys = {clean_sku_name(f) for f in folders}
          vr = {k: v for k, v in st.session_state.get('vision_results', {}).items() if k in valid_keys}
          missing = {f: pl_zip_images[f] for f in pl_folders if clean_sku_name(f) not in vr and pl_zip_images.get(f)}
          if missing:
            with st.spinner("🤖 AI Vision: Fabric, Sleeves, Pattern aur Color scan ho raha hai (Gemini ➜ NVIDIA ➜ Groq)..."):
              new_res, failed = run_vision_detection(missing)
              vr.update(new_res)
            if new_res: st.success(f"🤖 AI Vision: {len(new_res)} SKUs auto-detected!" + (f" | Fallback used for: {', '.join(failed)}" if failed else ""))
            else: st.warning("🤖 AI se result nahi mila — Python intelligence use hoga.")
          st.session_state['vision_results'] = vr

          ctx = {'folders': folders, 'sizes': sizes_list, 'brand': brand_name.strip(), 'mrp': mrp_value.strip(), 'mfg': mfg_address.strip(), 'hsn': hsn_code.strip(), 'keywords': custom_keywords.strip(), 'dropbox': dropbox_url.strip(), 'live_links_cache': {}, 'folder_images_count': {f: (len(pl_zip_images[f]) if f in pl_zip_images else 4) for f in folders}, 'size_chart_cdn_url': None, 'quantity': int(sc_quantity), 'selling_price': str(sc_selling_price).strip(), 'sc_parent_model': str(sc_parent_model).strip(), 'meesho_return_price': str(meesho_return_price).strip(), 'meesho_mfg_name': str(meesho_mfg_name).strip(), 'meesho_mfg_pin': str(meesho_mfg_pin).strip(), 'meesho_gst': int(meesho_gst), 'meesho_weight': str(meesho_weight).strip(), 'td': int(default_td), 'default_category': None, 'vision_results': vr}

          if auto_cdn and pl_zip_images:
            with st.spinner("Images ko live CDN par upload kiya ja raha hai..."):
              for fld, sc_info in pl_size_charts.items():
                u = upload_to_imgbb_cached(sc_info['bytes'], IMGBB_PERMANENT_KEY, img_cache)
                if u and not ctx['size_chart_cdn_url']: ctx['size_chart_cdn_url'] = u
              for fld, img_tuples in pl_zip_images.items():
                ctx['live_links_cache'][fld] = []
                for iname, ibytes in img_tuples:
                  u = upload_to_imgbb_cached(ibytes, IMGBB_PERMANENT_KEY, img_cache)
                  if u: ctx['live_links_cache'][fld].append(u)

          out, preview, added = None, [], 0
          if marketplace == "Amazon - Vendor Central":
            out, preview, added = run_vendor_central(tpl_file, ctx)
            st.success(f"🎉 Amazon Vendor Central template filled! Total {added} variants.")
          elif marketplace == "Amazon - Seller Central":
            out, preview, added = run_seller_central(tpl_file, ctx)
            st.success(f"🎉 Amazon Seller Central template (`Template` sheet) filled! Total {added} rows.")
          elif marketplace == "Meesho":
            out, preview, added, m_sheet = run_meesho(tpl_file, ctx, pl_size_charts)
            st.success(f"🎉 Meesho template (`{m_sheet}`) filled! Total **{added}** variants add hue.")
          else:
            out, preview, added = run_generic_marketplace(tpl_file, ctx, marketplace)
            st.success(f"🎉 {marketplace} template filled! Total {added} variants.")

          if out:
            if preview: st.dataframe(preview, use_container_width=True)
            base_name, ext = os.path.splitext(tpl_file.name)
            out_ext = ext if ext.lower() in ['.xlsm', '.xlsx'] else '.xlsx'
            dl_name = f"{base_name}_FILLED{out_ext}" if is_meesho else tpl_file.name
            st.download_button("📥 Download Filled File", data=out, file_name=dl_name, mime=_mime_for(dl_name), type="primary")
      except ValueError as ve: st.error(str(ve))
      except Exception as e:
        st.error(f"Error: {e}")
        with st.expander("Technical details"): st.code(traceback.format_exc())

if CONVERT_MODE:
  render_convert_ui()
else:
  render_listing_ui()
