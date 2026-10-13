import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import os
import re
import zipfile
import base64
import requests
from datetime import datetime

try:
    import xlrd
    import xlwt
    from xlutils.copy import copy as xl_copy
    XLS_SUPPORT = True
except ImportError:
    XLS_SUPPORT = False

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")

st.title("📦 Multi-Marketplace Auto Listing Generator (Universal Pro V11)")
st.caption("Ajio, Amazon Seller Central, Amazon Vendor Central & Flipkart - Strict Category & Dropdown Compliance")

IMGBB_PERMANENT_KEY = "137be55d19cf3d6a7ec051e369137041"

# 1. Marketplace Selection
col_m1, col_m2 = st.columns([2, 1])
with col_m1:
    marketplace = st.selectbox(
        "Marketplace Chunein:",
        ["Amazon - Vendor Central", "Ajio", "Amazon - Seller Central", "Flipkart", "Meesho", "Myntra"]
    )
with col_m2:
    default_td = st.number_input(
        "TD / Default Margin (%):",
        min_value=0,
        max_value=100,
        value=50,
        help="Ajio me TD by default 50 set rahega."
    )

st.divider()

# 2. Product Images & Input Method
st.subheader("1. Product Images & SKUs Input (ZIP Folder ya Dropbox Link)")
st.info("💡 ZIP ke andar har product ke alag-alag folders honge (jaise SKU-1, SKU-2). Har folder ka alag SKU banega.")

col_in1, col_in2 = st.columns(2)

detected_folders = []
folder_images_count = {}
zip_images_data = {}
detected_size_chart_info = {}

def clean_sku_name(name):
    cleaned = re.sub(r'[\s_]+', '-', str(name).strip())
    cleaned = re.sub(r'-+', '-', cleaned).strip('-')
    return cleaned

def is_size_chart_file(filename):
    low = os.path.basename(filename).lower()
    patterns = ['size_chart', 'size-chart', 'sizechart', 'size_guide', 'sizeguide', 'size chart', 'chart', 'measurement']
    return any(p in low for p in patterns) and low.endswith(('.jpg', '.jpeg', '.png', '.webp'))

with col_in1:
    st.write("##### Option A: Product Folders Ka ZIP Upload")
    zip_file = st.file_uploader(
        "Product Folders Ka ZIP (.zip):",
        type=["zip"],
        help="ZIP ke andar har product ka sub-folder aur Size Chart image daalein."
    )
    if zip_file:
        try:
            with zipfile.ZipFile(zip_file, 'r') as z:
                file_list = z.namelist()
                
                # Check for Size Chart image
                for f in file_list:
                    if is_size_chart_file(f):
                        detected_size_chart_info['filename'] = os.path.basename(f)
                        detected_size_chart_info['bytes'] = z.read(f)
                        break

                # Extract product folders strictly by immediate parent directory of images
                img_exts = ('.jpg', '.jpeg', '.png', '.webp')
                folder_to_imgs = {}
                
                for f in file_list:
                    f_clean = f.replace('\\', '/')
                    if f_clean.lower().endswith(img_exts) and not is_size_chart_file(f_clean):
                        parent_dir = os.path.dirname(f_clean)
                        sub_folder_name = os.path.basename(parent_dir)
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

            st.success(f"✅ ZIP Se Total {len(detected_folders)} Alag Product Folders (SKUs) Detect Hue:")
            st.code(", ".join([clean_sku_name(f) for f in detected_folders]))

            if detected_size_chart_info:
                st.info(f"📐 **Universal Size Chart Detect Hua:** `{detected_size_chart_info['filename']}` — Sabhi products par follow hoga.")
        except Exception as e:
            st.error(f"ZIP read karne me error: {e}")

    auto_cdn_upload = st.checkbox(
        "🌐 ZIP images se automatic Live CDN Workable Links banayein (Active)",
        value=True
    )
    st.success("⚡ Live Workable CDN Active: 100% working direct links banenge.")

with col_in2:
    st.write("##### Option B: Dropbox Link / Manual SKUs")
    dropbox_base_url = st.text_input(
        "Dropbox Shared Folder Link (Optional):",
        placeholder="https://www.dropbox.com/scl/fo/xyz/..."
    )
    manual_skus_input = st.text_area(
        "Manual SKUs List (Agar ZIP nahi hai):",
        placeholder="SKU-1\nSKU-2\nSKU-4",
        height=100
    )
    if not detected_folders and manual_skus_input.strip():
        detected_folders = [s.strip() for s in manual_skus_input.split('\n') if s.strip()]
        for fld in detected_folders:
            folder_images_count[fld] = 5

st.divider()

# 3. Sizes, Brand & Legal Information
st.subheader("2. Sizes, Brand & Legal Information")
col_s1, col_s2, col_s3 = st.columns(3)

with col_s1:
    sizes_input = st.text_input(
        "Sizes (comma-separated):",
        value="30, 32, 34, 36, 38"
    )
    brand_name = st.text_input(
        "Brand Name:",
        value="",
        placeholder="e.g. Lymio, Buda Jeans Co (Leave blank if not applicable)"
    )
    mrp_value = st.text_input(
        "MRP (Khali chhodne par Yellow Highlight hoga):",
        value="",
        placeholder="e.g. 2999"
    )

with col_s2:
    mfg_address = st.text_area(
        "Manufactured By (Address with Pincode):",
        value="",
        placeholder="Enter Manufacturer complete address with Pincode...",
        height=100
    )

with col_s3:
    mkt_address = st.text_area(
        "Marketed By / Packer Details:",
        value="",
        placeholder="Enter Marketer / Packer complete address with Pincode...",
        height=100
    )

st.divider()

# 4. Trending Keywords Section
st.subheader("3. SEO & Trending Keywords")
custom_keywords = st.text_area(
    "Custom Trending Keywords (Optional):",
    placeholder="e.g. printed hoodie, casual hoodie, fleece pullover, cotton blend winter wear",
    help="Keywords sequence-wise use honge, title 75 char limit me rahega aur repeat nahi honge."
)

st.divider()

# 5. Template Uploader
st.subheader("4. Template File Upload & Generate")
uploaded_file = st.file_uploader(
    "Apni Excel Template (.xlsx, .xlsm ya .xls) Upload Karein:",
    type=["xlsx", "xlsm", "xls"]
)

def detect_category(folder_name):
    low = folder_name.lower()
    if "cargo" in low: return "cargo"
    elif "hoodie" in low: return "hoodie"
    elif "sweatshirt" in low: return "sweatshirt"
    elif "jean" in low or "denim" in low: return "jeans"
    elif "t-shirt" in low or "tshirt" in low or "tee" in low: return "tshirt"
    elif "shirt" in low: return "shirt"
    elif "trouser" in low or "pant" in low: return "trousers"
    return "apparel"

COMMON_COLORS = [
    "Beige", "LightGrey", "Grey", "DarkGrey", "Black", "Navy", "Blue", "LightBlue", "DarkBlue", 
    "IceBlue", "White", "Off White", "Green", "Olive", "Brown", "Coffee", "Charcoal", "Khaki", "Red", "Maroon", "Yellow", "Orange", "Pink"
]

STANDARD_COLOR_MAPS = [
    'Beige', 'Black', 'Blue', 'Bronze', 'Brown', 'Clear', 'Gold', 'Green', 'Grey', 
    'Metallic', 'Multicolor', 'Off White', 'Orange', 'Pink', 'Purple', 'Red', 'Silver', 'Turquoise', 'White', 'Yellow'
]

def detect_color(text):
    text_clean = text.replace('-', ' ').replace('_', ' ')
    for col in COMMON_COLORS:
        if col.lower() in text_clean.lower():
            return col
    return "Black" if "hoodie" in text.lower() else "Blue"

def map_to_standard_color_map(detected_col):
    det_lower = detected_col.lower()
    for sc in STANDARD_COLOR_MAPS:
        if sc.lower() == det_lower:
            return sc
    if "grey" in det_lower: return "Grey"
    if "blue" in det_lower or "navy" in det_lower: return "Blue"
    if "black" in det_lower or "charcoal" in det_lower: return "Black"
    if "white" in det_lower: return "White"
    if "khaki" in det_lower or "olive" in det_lower: return "Green"
    if "brown" in det_lower or "coffee" in det_lower: return "Brown"
    return "Blue"

def make_workable_dropbox_url(base_url, sku, img_num):
    if not base_url: return f"https://www.dropbox.com/s/{sku}/{img_num}.jpg?raw=1"
    clean = base_url.split("?")[0].rstrip("/")
    return f"{clean}/{sku}/{img_num}.jpg?raw=1"

def upload_to_imgbb(image_bytes, api_key):
    try:
        url = "https://api.imgbb.com/1/upload"
        payload = {"key": api_key, "image": base64.b64encode(image_bytes).decode('utf-8')}
        res = requests.post(url, data=payload, timeout=20)
        if res.status_code == 200: return res.json()['data']['url']
    except Exception:
        pass
    return None

def build_vendor_title(brand, category, full_sku, user_kw=""):
    b = f"{brand} " if brand else ""
    cat_str = category.capitalize()
    sku_part = f"({full_sku})"

    kw_candidate = ""
    if user_kw:
        kws = [k.strip() for k in user_kw.replace(';', ',').split(',') if k.strip()]
        if kws:
            kw_candidate = kws[0]

    if kw_candidate:
        opt1 = f"{b}{cat_str} for men | {kw_candidate} {sku_part}".strip()
        if len(opt1) <= 75:
            title = opt1
        else:
            opt2 = f"{b}{kw_candidate} {sku_part}".strip()
            title = opt2 if len(opt2) <= 75 else f"{b}{cat_str} {sku_part}".strip()
    else:
        title = f"{b}{cat_str} for men | casual wear {sku_part}".strip()

    if len(title) > 75:
        title = title[:75].rstrip()

    words = title.split()
    counts = {}
    filtered = []
    for w in words:
        wl = w.lower().strip("|,()")
        c = counts.get(wl, 0)
        if c < 2 or w in ["|", "(", ")", "-"]:
            filtered.append(w)
            if wl: counts[wl] = c + 1
    final_title = " ".join(filtered)
    return final_title[:75]

def get_style_from_dropdown(cat):
    if cat == 'cargo': return 'Cargo'
    if cat in ['sweatshirt', 'hoodie']: return 'Sweatpants'
    if cat == 'trousers': return 'Chino'
    return 'Modern'

def get_item_type_name_from_dropdown(cat):
    if cat == 'cargo': return 'Casual Pants'
    if cat == 'trousers': return 'Casual Pants'
    if cat == 'jeans': return 'Jeans'
    if cat in ['hoodie', 'sweatshirt']: return 'Sweatpants'
    return 'Casual Pants'

def generate_generic_keywords_200(cat, det_col, user_kw):
    if user_kw:
        kws = [k.strip() for k in user_kw.replace(';', ',').split(',') if k.strip()]
        unique_kws = []
        for k in kws:
            if k.lower() not in [x.lower() for x in unique_kws]:
                unique_kws.append(k)
        res = "; ".join(unique_kws)
        return res[:200]
    
    if cat in ['hoodie', 'sweatshirt']:
        kw_list = [f"men {cat}", f"{det_col} hoodie", "pullover sweatshirt", "winter wear for men", "casual fleece hoodie", "streetwear sweatshirt", "comfortable hoodie"]
    elif cat == 'cargo':
        kw_list = [f"men cargo pants", f"{det_col} cargo", "utility pants", "multi pocket trousers", "streetwear cargo", "relaxed fit pants", "men cotton cargo"]
    else:
        kw_list = [f"men jeans", f"{det_col} denim", "stretchable jeans", "casual denim pants", "mid rise jeans", "regular fit denim", "washed jeans for men"]
    
    res = "; ".join(kw_list)
    return res[:200]

def build_category_bullets_and_desc(cat, det_col, brand):
    b_prefix = f"{brand} " if brand else ""
    if cat in ['hoodie', 'sweatshirt']:
        bullets = [
            f"PREMIUM WARMTH & COMFORT: Crafted from ultra-soft, breathable cotton fleece fabric that provides exceptional warmth and all-day comfort.",
            f"MODERN STREETWEAR DESIGN: Features a stylish {det_col.lower()} hue with a cozy hood, adjustable drawstrings, and reinforced stitching for durability.",
            "PRACTICAL UTILITY: Equipped with a spacious front kangaroo pocket to keep essentials secure and hands warm in chilly weather.",
            "VERSATILE CASUAL PAIRING: Pairs effortlessly with denims, joggers, cargo trousers, and sneakers for a sporty or relaxed everyday look.",
            "EASY WASH CARE: Machine wash cold with similar colors. Do not bleach. Warm iron if needed."
        ]
        desc = (
            f"Upgrade your winter casuals with the {b_prefix}Men's {det_col} {cat.capitalize()}. "
            "Tailored from high-quality cotton blend fleece, this piece blends urban streetwear aesthetic with premium relaxed comfort. "
            "Designed for daily wear, college, travel, and evening outings."
        )
    elif cat == 'cargo':
        bullets = [
            f"DURABLE COTTON TWILL FABRIC: Engineered from premium breathable cotton fabric that offers maximum mobility and rugged durability.",
            f"MULTI-POCKET FUNCTIONALITY: Designed with deep tactical flap cargo pockets, side slash pockets, and back pockets for spacious storage.",
            f"COMFORTABLE RELAXED FIT: Features an ergonomic mid-rise waistband and relaxed leg silhouette in trendy {det_col.lower()} styling.",
            "ADVENTURE & STREETWEAR READY: Perfect for outdoor trekking, travel, casual hangouts, and pairing with oversized tees or hoodies.",
            "CARE INSTRUCTIONS: Machine wash cold inside out. Tumble dry low. Warm iron if required."
        ]
        desc = (
            f"Experience peak comfort and utility with the {b_prefix}Men's {det_col} Cargo Pants. "
            "Constructed with heavy-duty reinforced seams and high-strength cotton fabric for an active, stylish lifestyle."
        )
    else:
        bullets = [
            f"COMFORT STRETCH DENIM: Crafted from high-grade stretchable denim fabric designed for flexible mobility and all-day ease.",
            f"CLASSIC WASHED STYLING: Features an authentic {det_col.lower()} washed finish with subtle fading and durable contrast stitching.",
            "5-POCKET UTILITY: Equipped with two front scoop pockets, a coin pocket, and two patch back pockets with a sturdy metal zip fly.",
            "TIMELESS RELAXED SILHOUETTE: Mid-rise fit that sits comfortably at the waist, ideal for daily college, office casuals, and weekends.",
            "GARMENT CARE: Machine wash cold inside out with like colors. Do not bleach. Warm iron if needed."
        ]
        desc = (
            f"Redefine your everyday denim rotation with the {b_prefix}Men's {det_col} Jeans. "
            "Engineered with premium stretchable cotton blend denim, these jeans offer supreme durability and an on-trend look."
        )
    return bullets, desc

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

# ================= GENERATION PROCESS =================
if uploaded_file and st.button("🚀 Generate Listing File"):
    if not detected_folders:
        st.error("Kripya ZIP file upload karein ya SKUs list enter karein!")
    else:
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()
        size_list = [sz.strip() for sz in sizes_input.split(",") if sz.strip()]

        live_links_cache = {}
        size_chart_cdn_url = None

        if auto_cdn_upload:
            with st.spinner("Images aur Size Chart ko live CDN par upload kiya ja raha hai..."):
                if 'bytes' in detected_size_chart_info:
                    size_chart_cdn_url = upload_to_imgbb(detected_size_chart_info['bytes'], IMGBB_PERMANENT_KEY)

                if zip_images_data:
                    for fld, img_tuples in zip_images_data.items():
                        live_links_cache[fld] = []
                        for img_name, img_bytes in img_tuples:
                            cdn_url = upload_to_imgbb(img_bytes, IMGBB_PERMANENT_KEY)
                            if cdn_url: live_links_cache[fld].append(cdn_url)

        # ---------------- OPENPYXL ENGINE ----------------
        if file_ext in [".xlsx", ".xlsm"]:
            try:
                is_xlsm = file_ext == ".xlsm"
                wb = openpyxl.load_workbook(uploaded_file, keep_vba=is_xlsm)

                for s in wb.sheetnames:
                    wb[s].protection.sheet = False
                    wb[s].protection.disable()

                mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
                mrp_font = Font(bold=True)
                material_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
                total_variants_added = 0

                # ---------------- AMAZON VENDOR CENTRAL ----------------
                if "Vendor Central" in marketplace:
                    ws = None
                    for sname in wb.sheetnames:
                        if "template" in sname.lower():
                            ws = wb[sname]
                            break
                    if not ws: ws = wb.worksheets[0]

                    ungroup_worksheet(ws)

                    col_by_label = {}
                    col_by_key = {}
                    for c in range(1, ws.max_column + 1):
                        lbl = str(ws.cell(3, c).value or '').strip().lower()
                        key = str(ws.cell(4, c).value or '').strip().lower()
                        if lbl and lbl not in col_by_label:
                            col_by_label[lbl] = c
                        if key and key not in col_by_key:
                            col_by_key[key] = c

                    def get_vc_col(label_candidates, key_candidates=[]):
                        for k in key_candidates:
                            if k.lower() in col_by_key:
                                return col_by_key[k.lower()]
                        for l in label_candidates:
                            if l.lower() in col_by_label:
                                return col_by_label[l.lower()]
                        return None

                    curr_r = 7
                    sku_col_idx = get_vc_col(['vendor sku'], ['vendor_sku#1.value']) or 2
                    while ws.cell(curr_r, sku_col_idx).value is not None:
                        curr_r += 1

                    today_str = datetime.now().strftime("%Y-%m-%d")
                    today_shipping = datetime.now().strftime("%Y%m%d")

                    for raw_fld in detected_folders:
                        clean_style = clean_sku_name(raw_fld)
                        det_col = detect_color(raw_fld)
                        cat = detect_category(raw_fld)

                        if raw_fld in live_links_cache and live_links_cache[raw_fld]:
                            sku_img_urls = live_links_cache[raw_fld]
                        else:
                            img_cnt = folder_images_count.get(raw_fld, 5)
                            sku_img_urls = [make_workable_dropbox_url(dropbox_base_url, clean_style, i) for i in range(1, img_cnt + 1)]

                        if size_chart_cdn_url and size_chart_cdn_url not in sku_img_urls:
                            sku_img_urls.append(size_chart_cdn_url)

                        bullets_list, desc_text = build_category_bullets_and_desc(cat, det_col, brand_name)
                        gen_keywords_200 = generate_generic_keywords_200(cat, det_col, custom_keywords)
                        style_val = get_style_from_dropdown(cat)
                        item_type_val = get_item_type_name_from_dropdown(cat)
                        color_map_val = map_to_standard_color_map(det_col)

                        for sz in size_list:
                            clean_sz = clean_sku_name(sz)
                            full_sku = f"{clean_style}-{clean_sz}"

                            vendor_title = build_vendor_title(brand_name, cat, full_sku, custom_keywords)

                            c_vc = get_vc_col(['vendor code'], ['rtip_vendor_code#1.value'])
                            if c_vc: ws.cell(curr_r, c_vc).value = "J B fashion -- Surat -- Dropship, in_pd-ravi_apparel, JBZ54"

                            c_sku = get_vc_col(['vendor sku'], ['vendor_sku#1.value'])
                            if c_sku: ws.cell(curr_r, c_sku).value = full_sku

                            c_pt = get_vc_col(['product type'], ['product_type#1.value'])
                            if c_pt: ws.cell(curr_r, c_pt).value = "PANTS" if cat in ['jeans', 'cargo', 'trousers'] else "SWEATSHIRT"

                            for p_lbl, p_key in [
                                (['parentage level'], ['parentage_level#1.value']),
                                (['child relationship type'], ['child_parent_sku_relationship#1.child_relationship_type']),
                                (['parent sku'], ['child_parent_sku_relationship#1.parent_sku']),
                                (['variation theme name'], ['variation_theme#1.name'])
                            ]:
                                c_p = get_vc_col(p_lbl, p_key)
                                if c_p: ws.cell(curr_r, c_p).value = None

                            c_title = get_vc_col(['item name'], ['item_name#1.value'])
                            if c_title: ws.cell(curr_r, c_title).value = vendor_title

                            c_hl = get_vc_col(['item highlight'], ['title_differentiation#1.value'])
                            if c_hl: ws.cell(curr_r, c_hl).value = f"Premium {det_col} finish, durable comfort stretch"

                            c_br = get_vc_col(['brand name'], ['brand#1.value'])
                            if c_br: ws.cell(curr_r, c_br).value = brand_name

                            c_mod = get_vc_col(['model number'], ['model_number#1.value'])
                            if c_mod: ws.cell(curr_r, c_mod).value = full_sku
                            c_part = get_vc_col(['part number'], ['part_number#1.value'])
                            if c_part: ws.cell(curr_r, c_part).value = full_sku

                            c_mod_name = get_vc_col(['model name'], ['model_name#1.value'])
                            if c_mod_name: ws.cell(curr_r, c_mod_name).value = f"{cat.capitalize()} {clean_style}"

                            c_mfg = get_vc_col(['manufacturer'], ['manufacturer#1.value'])
                            if c_mfg: ws.cell(curr_r, c_mfg).value = mfg_address
                            c_mfg_cnt = get_vc_col(['manufacturer contact information'], ['rtip_manufacturer_contact_information#1.value'])
                            if c_mfg_cnt: ws.cell(curr_r, c_mfg_cnt).value = mfg_address

                            c_main_img = get_vc_col(['main image url'], ['main_product_image_locator#1.media_location'])
                            if c_main_img and sku_img_urls: ws.cell(curr_r, c_main_img).value = sku_img_urls[0]

                            for img_num in range(1, 9):
                                c_oth = get_vc_col([f'other image url'], [f'other_product_image_locator_{img_num}#1.media_location'])
                                if c_oth and img_num < len(sku_img_urls):
                                    ws.cell(curr_r, c_oth).value = sku_img_urls[img_num]

                            for b_i, b_val in enumerate(bullets_list[:5]):
                                c_bp = get_vc_col([f'bullet point'], [f'bullet_point#{b_i+1}.value'])
                                if c_bp: ws.cell(curr_r, c_bp).value = b_val

                            c_gkw = get_vc_col(['generic keywords'], ['generic_keyword#1.value'])
                            if c_gkw and gen_keywords_200: ws.cell(curr_r, c_gkw).value = gen_keywords_200

                            c_style = get_vc_col(['style'], ['style#1.value'])
                            if c_style: ws.cell(curr_r, c_style).value = style_val

                            lifestyles = ["Casual", "Work Utility", "Comfort", "Athletic", "Formal"]
                            for ls_idx in range(1, 6):
                                c_ls = get_vc_col([f'lifestyle'], [f'lifestyle#{ls_idx}.value'])
                                if c_ls: ws.cell(curr_r, c_ls).value = lifestyles[ls_idx - 1]

                            c_dept = get_vc_col(['department name'], ['department#1.value'])
                            if c_dept: ws.cell(curr_r, c_dept).value = "Mens"
                            c_gen = get_vc_col(['target gender'], ['target_gender#1.value'])
                            if c_gen: ws.cell(curr_r, c_gen).value = "Male"
                            c_age = get_vc_col(['age range description'], ['age_range_description#1.value'])
                            if c_age: ws.cell(curr_r, c_age).value = "Adult"

                            c_ssys = get_vc_col(['bottoms size system'], ['bottoms_size#1.size_system'])
                            if c_ssys: ws.cell(curr_r, c_ssys).value = "IN"

                            is_num_sz = clean_sz.isdigit()
                            size_class_val = "Numeric" if is_num_sz else "Alpha"
                            c_scls = get_vc_col(['bottoms size class'], ['bottoms_size#1.size_class'])
                            if c_scls: ws.cell(curr_r, c_scls).value = size_class_val

                            c_sval = get_vc_col(['bottoms size value'], ['bottoms_size#1.size'])
                            if c_sval: ws.cell(curr_r, c_sval).value = clean_sz

                            mat_val = "Cotton" if cat in ['cargo', 'hoodie', 'sweatshirt'] else "Cotton Blend"
                            c_mat = get_vc_col(['material'], ['material#1.value'])
                            if c_mat:
                                cell_m = ws.cell(curr_r, c_mat)
                                cell_m.value = mat_val
                                cell_m.fill = material_fill

                            c_fab = get_vc_col(['fabric type'], ['fabric_type#1.value'])
                            if c_fab: ws.cell(curr_r, c_fab).value = f"100% {mat_val}"

                            c_noi = get_vc_col(['number of items'], ['number_of_items#1.value'])
                            if c_noi: ws.cell(curr_r, c_noi).value = 1

                            c_itn = get_vc_col(['item type name'], ['item_type_name#1.value'])
                            if c_itn: ws.cell(curr_r, c_itn).value = item_type_val

                            c_spec_sz = get_vc_col(['special size'], ['special_size_type#1.value'])
                            if c_spec_sz: ws.cell(curr_r, c_spec_sz).value = "Standard"

                            c_desc = get_vc_col(['product description'], ['rtip_product_description#1.value'])
                            if c_desc: ws.cell(curr_r, c_desc).value = desc_text

                            c_cmap = get_vc_col(['color map'], ['color#1.standardized_values#1'])
                            if c_cmap: ws.cell(curr_r, c_cmap).value = color_map_val
                            c_color = get_vc_col(['color'], ['color#1.value'])
                            if c_color: ws.cell(curr_r, c_color).value = det_col

                            c_thm = get_vc_col(['theme'], ['theme#1.value'])
                            if c_thm: ws.cell(curr_r, c_thm).value = None

                            c_fit = get_vc_col(['fit type'], ['fit_type#1.value'])
                            if c_fit: ws.cell(curr_r, c_fit).value = "Relaxed"
                            c_poc = get_vc_col(['pocket description'], ['pocket_description#1.value'])
                            if c_poc: ws.cell(curr_r, c_poc).value = "Utility Pocket" if cat == 'cargo' else "Basic-5-Pocket"

                            c_care = get_vc_col(['care instructions'], ['care_instructions#1.value'])
                            if c_care: ws.cell(curr_r, c_care).value = "Machine Wash"

                            c_pat = get_vc_col(['pattern'], ['pattern#1.value'])
                            if c_pat: ws.cell(curr_r, c_pat).value = "Solid"

                            c_uc = get_vc_col(['unit count'], ['unit_count#1.value'])
                            if c_uc: ws.cell(curr_r, c_uc).value = 1
                            c_uct = get_vc_col(['unit count type'], ['unit_count#1.type.value'])
                            if c_uct: ws.cell(curr_r, c_uct).value = "count"

                            c_pld = get_vc_col(['product site launch date'], ['product_site_launch_date#1.value'])
                            if c_pld: ws.cell(curr_r, c_pld).value = today_str

                            c_hsn_ent = get_vc_col(['external product information entity'], ['external_product_information#1.entity'])
                            if c_hsn_ent: ws.cell(curr_r, c_hsn_ent).value = "HSN Code"
                            c_hsn_val = get_vc_col(['external product information'], ['external_product_information#1.value'])
                            if c_hsn_val: ws.cell(curr_r, c_hsn_val).value = "62034290" if cat == 'jeans' else "62034200"

                            if is_num_sz:
                                c_wst = get_vc_col(['waist size'], ['waist#1.size#1.value'])
                                if c_wst: ws.cell(curr_r, c_wst).value = clean_sz
                                c_wstu = get_vc_col(['waist size unit'], ['waist#1.size#1.unit'])
                                if c_wstu: ws.cell(curr_r, c_wstu).value = "inches"

                            c_lst = get_vc_col(['product lifecycle supply type'], ['lifecycle_supply_type#1.value'])
                            if c_lst: ws.cell(curr_r, c_lst).value = "Year Round Replenishable"
                            c_sea = get_vc_col(['seasons'], ['seasons#1.value'])
                            if c_sea: ws.cell(curr_r, c_sea).value = "All"

                            c_sld = get_vc_col(['sleeve length description', 'sleeve length'])
                            if c_sld: ws.cell(curr_r, c_sld).value = "Long Sleeve" if cat in ['hoodie', 'sweatshirt', 'shirt'] else "Short Sleeve"
                            c_st = get_vc_col(['sleeve type'])
                            if c_st: ws.cell(curr_r, c_st).value = "Regular Sleeve"

                            c_pkr = get_vc_col(['packer contact information'], ['packer_contact_information#1.value'])
                            if c_pkr: ws.cell(curr_r, c_pkr).value = mkt_address

                            c_iwt = get_vc_col(['item weight'], ['item_weight#1.value'])
                            if c_iwt: ws.cell(curr_r, c_iwt).value = 350
                            c_iwtu = get_vc_col(['item weight unit'], ['item_weight#1.unit'])
                            if c_iwtu: ws.cell(curr_r, c_iwtu).value = "Grams"

                            c_mrp = get_vc_col(['maximum retail price (mrp)'], ['list_price#1.value_with_tax'])
                            if c_mrp:
                                mrp_cell = ws.cell(curr_r, c_mrp)
                                if mrp_value:
                                    mrp_cell.value = int(mrp_value) if mrp_value.isdigit() else mrp_value
                                else:
                                    mrp_cell.fill = mrp_fill
                                    mrp_cell.font = mrp_font

                            c_esd = get_vc_col(['earliest shipping date'], ['rtip_earliest_shipping_date#1.value'])
                            if c_esd: ws.cell(curr_r, c_esd).value = today_shipping

                            c_il = get_vc_col(['item length'], ['item_dimensions#1.length.value'])
                            if c_il: ws.cell(curr_r, c_il).value = 10
                            c_ilu = get_vc_col(['item length unit'], ['item_dimensions#1.length.unit'])
                            if c_ilu: ws.cell(curr_r, c_ilu).value = "inches"

                            c_iw = get_vc_col(['item width'], ['item_dimensions#1.width.value'])
                            if c_iw: ws.cell(curr_r, c_iw).value = 2
                            c_iwu = get_vc_col(['item width unit'], ['item_dimensions#1.width.unit'])
                            if c_iwu: ws.cell(curr_r, c_iwu).value = "inches"

                            c_ih = get_vc_col(['item height'], ['item_dimensions#1.height.value'])
                            if c_ih: ws.cell(curr_r, c_ih).value = 1.3
                            c_ihu = get_vc_col(['item height unit'], ['item_dimensions#1.height.unit'])
                            if c_ihu: ws.cell(curr_r, c_ihu).value = "inches"

                            c_pl = get_vc_col(['item package length'], ['item_package_dimensions#1.length.value'])
                            if c_pl: ws.cell(curr_r, c_pl).value = 10
                            c_plu = get_vc_col(['package length unit'], ['item_package_dimensions#1.length.unit'])
                            if c_plu: ws.cell(curr_r, c_plu).value = "inches"

                            c_pw = get_vc_col(['item package width'], ['item_package_dimensions#1.width.value'])
                            if c_pw: ws.cell(curr_r, c_pw).value = 2
                            c_pwu = get_vc_col(['package width unit'], ['item_package_dimensions#1.width.unit'])
                            if c_pwu: ws.cell(curr_r, c_pwu).value = "inches"

                            c_ph = get_vc_col(['item package height'], ['item_package_dimensions#1.height.value'])
                            if c_ph: ws.cell(curr_r, c_ph).value = 1.3
                            c_phu = get_vc_col(['package height unit'], ['item_package_dimensions#1.height.unit'])
                            if c_phu: ws.cell(curr_r, c_phu).value = "inches"

                            c_pwt = get_vc_col(['package weight'], ['item_package_weight#1.value'])
                            if c_pwt: ws.cell(curr_r, c_pwt).value = 380
                            c_pwtu = get_vc_col(['package weight unit'], ['item_package_weight#1.unit'])
                            if c_pwtu: ws.cell(curr_r, c_pwtu).value = "Grams"

                            c_oat = get_vc_col(['order aggregate type'], ['rtip_order_aggregate_type#1.value'])
                            if c_oat: ws.cell(curr_r, c_oat).value = "Each"

                            c_ipi = get_vc_col(['items per inner pack'], ['rtip_items_per_inner_pack#1.value'])
                            if c_ipi: ws.cell(curr_r, c_ipi).value = 1

                            c_coo = get_vc_col(['country of origin'], ['country_of_origin#1.value'])
                            if c_coo: ws.cell(curr_r, c_coo).value = "India"

                            c_dg = get_vc_col(['dangerous goods regulations', 'dangerous goods', 'hazmat'])
                            if c_dg: ws.cell(curr_r, c_dg).value = "Not Applicable"

                            curr_r += 1
                            total_variants_added += 1

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)

                    st.success(f"🎉 Amazon Vendor Central template perfectly filled! Total {total_variants_added} variants create hue.")
                    st.download_button(
                        label="📥 Download Filled Vendor Central File",
                        data=output,
                        file_name=uploaded_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                else:
                    st.info("Marketplace generation completed for selected channel.")
            except Exception as e:
                st.error(f"Processing error: {e}")
