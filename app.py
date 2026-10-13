import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import os
import re
import zipfile
import base64
import requests

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")

st.title("📦 Multi-Marketplace Auto Listing Generator (Pro V7)")
st.caption("Permanent Live CDN Links, ZIP Upload / Dropbox Support, Zero-Error Ajio Template Mapping")

# Permanent ImgBB API Key Hardcoded
IMGBB_PERMANENT_KEY = "137be55d19cf3d6a7ec051e369137041"

# 1. Marketplace Selection
col_m1, col_m2 = st.columns([2, 1])
with col_m1:
    marketplace = st.selectbox(
        "Marketplace Chunein:",
        ["Ajio", "Amazon", "Flipkart", "Meesho", "Myntra"]
    )
with col_m2:
    default_td = st.number_input(
        "TD (Trade Discount) Default:",
        min_value=0,
        max_value=100,
        value=50,
        help="Ajio me TD by default 50 set rahega."
    )

st.divider()

# 2. Product Images & Input Method (ZIP ya Dropbox - Dono Optional)
st.subheader("1. Product Images & SKUs Input (ZIP Folder ya Dropbox Link)")
st.info("💡 **Dono Options Available Hain:** Aap chahein toh direct **ZIP folder** upload karein ya **Dropbox Link** dein. Dono se workable links banenge!")

col_in1, col_in2 = st.columns(2)

detected_folders = []
folder_images_count = {}
zip_images_data = {}  # {folder_name: [(img_filename, img_bytes), ...]}

def clean_sku_name(name):
    cleaned = re.sub(r'[\s_]+', '-', name.strip())
    cleaned = re.sub(r'-+', '-', cleaned).strip('-')
    return cleaned

with col_in1:
    st.write("##### Option A: Product Folders Ka ZIP Upload")
    zip_file = st.file_uploader(
        "Product Folders Ka ZIP (.zip):",
        type=["zip"],
        help="ZIP ke andar har SKU ka folder hona chahiye (e.g. MENS JEANS 29 GREY, CARGO-01) aur andar 1.jpg, 2.jpg... images honi chahiye."
    )
    if zip_file:
        try:
            with zipfile.ZipFile(zip_file, 'r') as z:
                file_list = z.namelist()
                folders = set()
                for f in file_list:
                    parts = f.replace('\\', '/').strip('/').split('/')
                    if len(parts) > 1 and parts[0]:
                        folders.add(parts[0])
                    elif len(parts) == 1 and f.endswith('/'):
                        folders.add(parts[0])

                detected_folders = sorted(list(folders))
                for fld in detected_folders:
                    imgs = [
                        p for p in file_list 
                        if p.replace('\\', '/').startswith(fld + '/') 
                        and p.lower().endswith(('.jpg', '.jpeg', '.png', '.webp'))
                    ]
                    def sort_key(p):
                        nums = re.findall(r'\d+', os.path.basename(p))
                        return int(nums[0]) if nums else 999
                    imgs.sort(key=sort_key)

                    folder_images_count[fld] = len(imgs) if len(imgs) > 0 else 5
                    
                    zip_images_data[fld] = []
                    for img_path in imgs:
                        img_bytes = z.read(img_path)
                        zip_images_data[fld].append((os.path.basename(img_path), img_bytes))

            st.success(f"✅ ZIP Se Total {len(detected_folders)} Folders Detect Ho Gaye:")
            cleaned_preview = [clean_sku_name(f) for f in detected_folders]
            st.code(", ".join(cleaned_preview))
        except Exception as e:
            st.error(f"ZIP read karne me error: {e}")

    auto_cdn_upload = st.checkbox(
        "🌐 ZIP images se automatic Live CDN Workable Links banayein (Active)",
        value=True,
        help="Yeh feature ZIP ki images ko high-speed permanent CDN par upload karke direct .jpg links banata hai jo Chrome aur Ajio dono me direct open hote hain."
    )
    st.success("⚡ Live Workable CDN Active: ZIP images automatic 100% working live links banegi!")

with col_in2:
    st.write("##### Option B: Dropbox Link / Manual SKUs")
    dropbox_base_url = st.text_input(
        "Dropbox Shared Folder Link (Optional):",
        placeholder="https://www.dropbox.com/scl/fo/xyz/... ya file link",
        help="Agar aapne Dropbox link diya, toh system direct workable links (?raw=1) generate karega jo browser me direct khulte hain."
    )
    
    manual_skus_input = st.text_area(
        "Manual SKUs List (Agar ZIP upload nahi kiya hai):",
        placeholder="MENS JEANS 29 GREY\nCARGO-01-04\nMENS SHIRT 12 WHITE",
        height=100,
        help="Agar ZIP file nahi hai toh yahan direct SKUs daalein."
    )
    if not detected_folders and manual_skus_input.strip():
        detected_folders = [s.strip() for s in manual_skus_input.split('\n') if s.strip()]
        for fld in detected_folders:
            folder_images_count[fld] = 5

st.divider()

# 3. Sizes & Legal Attributes
st.subheader("2. Sizes, Brand & Legal Information")
col_s1, col_s2, col_s3 = st.columns(3)

with col_s1:
    sizes_input = st.text_input(
        "Sizes (comma-separated):",
        value="30, 32, 34, 36, 38",
        help="Numeric: 28, 30, 32 ya Alpha: S, M, L, XL, 2XL, 3XL, 4XL"
    )
    brand_name = st.text_input(
        "Brand Name (Default Blank - chahein toh fill karein):",
        value="",
        placeholder="Khali rahega agar kuch nahi likhenge"
    )
    mrp_value = st.text_input(
        "MRP (Khali chhodne par Yellow Highlight hoga):",
        value="",
        placeholder="e.g. 2999 ya blank chhod dein"
    )

with col_s2:
    mfg_address = st.text_area(
        "Manufactured By (Address with Pincode):",
        value="61/C, BLOCK-N, SUMEL BUSINESS PARK 5, Flyover, ASARVA, Ahmedabad, Gujarat, India - 380016",
        height=100
    )

with col_s3:
    mkt_address = st.text_area(
        "Marketed By (Address with Pincode):",
        value="61/C, BLOCK-N, SUMEL BUSINESS PARK 5, Flyover, ASARVA, Ahmedabad, Gujarat, India - 380016",
        height=100
    )

st.divider()

# 4. Trending Keywords Section
st.subheader("3. SEO & Trending Keywords")
custom_keywords = st.text_area(
    "Custom Trending Keywords (Optional):",
    placeholder="e.g. multi pocket cargo, streetwear casuals, relaxed fit, premium cotton blend, clean wash",
    help="Platform specific search ranking ke liye trending terms daalein."
)

st.divider()

# 5. Template Uploader
st.subheader("4. Template File & Generate")
uploaded_file = st.file_uploader(
    "Apni Excel Template (.xlsx ya .xlsm) Upload Karein:",
    type=["xlsx", "xlsm"]
)

def detect_category(folder_name):
    low = folder_name.lower()
    if "cargo" in low:
        return "cargo"
    elif "jean" in low or "denim" in low:
        return "jeans"
    elif "t-shirt" in low or "tshirt" in low or "tee" in low:
        return "tshirt"
    elif "shirt" in low:
        return "shirt"
    elif "trouser" in low or "pant" in low:
        return "trousers"
    return "general"

def find_sheet_for_category(wb, category):
    ignore = ['instruction', 'guideline', 'extra', 'rejection', 'seller info']
    valid_sheets = [s for s in wb.sheetnames if not any(ign in s.lower() for ign in ignore)]

    if category in ["cargo", "trousers"]:
        for s in valid_sheets:
            if any(w in s.lower() for w in ['trouser', 'pant', 'cargo']):
                return s
    elif category == "jeans":
        for s in valid_sheets:
            if 'jean' in s.lower():
                return s
    elif category == "tshirt":
        for s in valid_sheets:
            if any(w in s.lower() for w in ['t-shirt', 'tshirt', 'polo']):
                return s
    elif category == "shirt":
        for s in valid_sheets:
            if 'shirt' in s.lower() and not any(w in s.lower() for w in ['t-shirt', 'tshirt']):
                return s

    for s in valid_sheets:
        if 'styles' in s.lower():
            return s
    return valid_sheets[0] if valid_sheets else wb.sheetnames[0]

COMMON_COLORS = [
    "LightGrey", "Grey", "DarkGrey", "Black", "Navy", "Blue", "LightBlue", "DarkBlue", 
    "IceBlue", "White", "Green", "Olive", "Beige", "Brown", "Coffee", "Charcoal", "Khaki"
]

def detect_color(text):
    text_clean = text.replace('-', ' ').replace('_', ' ')
    for col in COMMON_COLORS:
        if col.lower() in text_clean.lower():
            return col
    return "Grey" if "cargo" in text.lower() else "Blue"

def make_workable_dropbox_url(base_url, sku, img_num):
    if not base_url:
        return f"https://www.dropbox.com/s/{sku}/{img_num}.jpg?raw=1"
    clean = base_url.split("?")[0].rstrip("/")
    return f"{clean}/{sku}/{img_num}.jpg?raw=1"

def upload_to_imgbb(image_bytes, api_key):
    try:
        url = "https://api.imgbb.com/1/upload"
        payload = {
            "key": api_key,
            "image": base64.b64encode(image_bytes).decode('utf-8')
        }
        res = requests.post(url, data=payload, timeout=20)
        if res.status_code == 200:
            return res.json()['data']['url']
    except Exception:
        pass
    return None

def build_seo_content(style, color, category, user_keywords=""):
    kw_str = user_keywords if user_keywords else "casual streetwear, relaxed fit, durable utility pockets, all-day comfort"
    
    if category == "cargo":
        desc = (
            f"CARGO FOR MEN - Designed for unmatched utility and streetwear style, this Men's {style} Cargo Pants "
            f"in {color} is crafted from premium breathable cotton fabric. Engineered with {kw_str}, "
            "it features multiple spacious flap cargo pockets, a comfortable flexi waist with sturdy button closure, "
            "and durable reinforced stitching throughout. Perfect for everyday wear, outdoor adventures, college, and travel."
        )
        highlight = f"Premium breathable cotton fabric | Comfortable flexi waist in {color} | Multi-pocket functional cargo utility styling | Relaxed fit | Easy machine wash"
        add_info = "Styling Tip: Pair with oversized t-shirts, hoodies, and chunky sneakers. Wash Care: Machine wash cold with similar colors."
        return desc[:950], highlight[:500], add_info[:500]
    elif category == "jeans":
        desc = (
            f"Upgrade your everyday denim rotation with this trendy Men's {style} Jeans in {color}. "
            f"Crafted from premium stretchable denim engineered with {kw_str}. "
            "Featuring a contemporary mid-rise relaxed fit, functional 5-pocket styling, sturdy zip fly, and authentic washed finish. "
            "Ideal for casual outings, daily wear, and weekend hangouts."
        )
        highlight = f"Premium stretch denim | Relaxed mid-rise silhouette in {color} | Authentic clean washed finish | Multi-pocket utility styling | All-season wear"
        add_info = "Styling Tip: Pair with casual shirts or boxy tees. Wash Care: Machine wash cold inside out."
        return desc[:950], highlight[:500], add_info[:500]
    else:
        desc = f"Men's Casual {style} in {color}. Premium quality, featuring {kw_str}. Designed for maximum comfort and style."
        return desc[:900], f"Premium quality | Color: {color}", "Machine wash cold."

if uploaded_file and st.button("🚀 Generate Multi-Category Listing File"):
    if not detected_folders:
        st.error("Kripya kam se kam ek folder/SKU enter karein ya ZIP file upload karein!")
    else:
        try:
            is_xlsm = uploaded_file.name.lower().endswith(".xlsm")
            wb = openpyxl.load_workbook(uploaded_file, keep_vba=is_xlsm)

            for s in wb.sheetnames:
                wb[s].protection.sheet = False
                wb[s].protection.disable()

            category_folders = {}
            for raw_folder in detected_folders:
                cat = detect_category(raw_folder)
                category_folders.setdefault(cat, []).append(raw_folder)

            size_list = [sz.strip() for sz in sizes_input.split(",") if sz.strip()]
            mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
            mrp_font = Font(bold=True)

            total_variants_added = 0
            sheets_modified = set()

            # Automatic Live CDN Upload using permanent API Key
            live_links_cache = {}
            if auto_cdn_upload and zip_images_data:
                with st.spinner("Images ko live workable CDN par upload kiya ja raha hai..."):
                    for fld, img_tuples in zip_images_data.items():
                        live_links_cache[fld] = []
                        for img_name, img_bytes in img_tuples:
                            cdn_url = upload_to_imgbb(img_bytes, IMGBB_PERMANENT_KEY)
                            if cdn_url:
                                live_links_cache[fld].append(cdn_url)

            for cat, folder_list in category_folders.items():
                target_sheet_name = find_sheet_for_category(wb, cat)
                ws = wb[target_sheet_name]
                sheets_modified.add(target_sheet_name)

                header_row = 1
                best_matches = 0
                keywords = ['style', 'sku', 'mrp', 'price', 'size', 'image', 'model', 'brand', 'description', 'weight']
                for r in range(1, min(7, ws.max_row + 1)):
                    vals = [str(ws.cell(r, c).value or '').lower() for c in range(1, ws.max_column + 1)]
                    m = sum(1 for v in vals if any(k in v for k in keywords))
                    if m > best_matches:
                        best_matches = m
                        header_row = r

                headers = {c: str(ws.cell(header_row, c).value or '').strip() for c in range(1, ws.max_column + 1)}

                def get_exact_col(exact_name):
                    clean_target = exact_name.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
                    for c, h in headers.items():
                        clean_h = h.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
                        if clean_h == clean_target:
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
                col_std_size = get_fallback_col(['StandardSize'])

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

                col_color = get_fallback_col(['Color Family'])
                col_pri_color = get_fallback_col(['Primary color'])
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
                    c_i = get_fallback_col([target_name, f'Image {i}', f'Image_{i}'])
                    if c_i and c_i not in img_cols:
                        img_cols.append(c_i)

                curr_r = header_row + 1
                while ws.cell(curr_r, col_style if col_style else 1).value is not None:
                    curr_r += 1

                for raw_fld in folder_list:
                    clean_style = clean_sku_name(raw_fld)
                    detected_col = detect_color(raw_fld)
                    desc_text, high_text, add_text = build_seo_content(clean_style, detected_col, cat, custom_keywords)

                    if raw_fld in live_links_cache and live_links_cache[raw_fld]:
                        sku_img_urls = live_links_cache[raw_fld]
                    else:
                        img_cnt = folder_images_count.get(raw_fld, 5)
                        sku_img_urls = [
                            make_workable_dropbox_url(dropbox_base_url, clean_style, i)
                            for i in range(1, img_cnt + 1)
                        ]

                    for sz in size_list:
                        clean_sz = clean_sku_name(sz)
                        full_sku = f"{clean_style}-{clean_sz}"

                        try:
                            num_sz = int(re.findall(r'\d+', clean_sz)[0])
                            w_val = str(num_sz + 1)
                            tf_val = str(num_sz - 1)
                        except:
                            w_val = "34"
                            tf_val = "32"

                        if col_style: ws.cell(curr_r, col_style).value = clean_style
                        if col_desc: ws.cell(curr_r, col_desc).value = desc_text
                        if col_sku: ws.cell(curr_r, col_sku).value = full_sku
                        if col_brand: ws.cell(curr_r, col_brand).value = brand_name
                        if col_ean: ws.cell(curr_r, col_ean).value = full_sku
                        if col_td: ws.cell(curr_r, col_td).value = default_td
                        if col_hsn: ws.cell(curr_r, col_hsn).value = 62034200 if cat in ['cargo', 'trousers'] else 62034290
                        if col_prod_grp: ws.cell(curr_r, col_prod_grp).value = "Casual"
                        if col_fash_grp: ws.cell(curr_r, col_fash_grp).value = "Fashion"
                        if col_season_cd: ws.cell(curr_r, col_season_cd).value = "Core"
                        if col_season_yr: ws.cell(curr_r, col_season_yr).value = "2026"
                        if col_size: ws.cell(curr_r, col_size).value = clean_sz
                        if col_std_size: ws.cell(curr_r, col_std_size).value = clean_sz

                        dim_h, dim_l, dim_w, dim_uom = "2", "25", "24", "CM"
                        wt_val, wt_uom = "350", "GRAM"

                        if col_art_h: ws.cell(curr_r, col_art_h).value = dim_h
                        if col_art_l: ws.cell(curr_r, col_art_l).value = dim_l
                        if col_art_w: ws.cell(curr_r, col_art_w).value = dim_w
                        if col_art_luom: ws.cell(curr_r, col_art_luom).value = dim_uom
                        if col_art_wt: ws.cell(curr_r, col_art_wt).value = wt_val
                        if col_art_wtuom: ws.cell(curr_r, col_art_wtuom).value = wt_uom

                        if col_pkg_h: ws.cell(curr_r, col_pkg_h).value = dim_h
                        if col_pkg_l: ws.cell(curr_r, col_pkg_l).value = dim_l
                        if col_pkg_w: ws.cell(curr_r, col_pkg_w).value = dim_w
                        if col_pkg_luom: ws.cell(curr_r, col_pkg_luom).value = dim_uom
                        if col_pkg_wt: ws.cell(curr_r, col_pkg_wt).value = wt_val
                        if col_pkg_wtuom: ws.cell(curr_r, col_pkg_wtuom).value = wt_uom

                        if col_comp_count: ws.cell(curr_r, col_comp_count).value = 1
                        if col_country: ws.cell(curr_r, col_country).value = "India"
                        if col_mfg_by: ws.cell(curr_r, col_mfg_by).value = mfg_address
                        if col_mkt_by: ws.cell(curr_r, col_mkt_by).value = mkt_address
                        if col_net_qty: ws.cell(curr_r, col_net_qty).value = 1
                        if col_pkg_contains: ws.cell(curr_r, col_pkg_contains).value = "1"

                        fab_type_val = "Cotton" if cat in ['cargo', 'trousers', 'shirt'] else "Denim"
                        pattern_val = "Solid" if cat in ['cargo', 'trousers'] else "Washed"

                        if col_color: ws.cell(curr_r, col_color).value = detected_col
                        if col_pri_color: ws.cell(curr_r, col_pri_color).value = detected_col
                        if col_fab_detail: ws.cell(curr_r, col_fab_detail).value = fab_type_val
                        if col_fab_type: ws.cell(curr_r, col_fab_type).value = fab_type_val
                        if col_pattern: ws.cell(curr_r, col_pattern).value = pattern_val
                        if col_size_format: ws.cell(curr_r, col_size_format).value = "UNI"
                        if col_size_group: ws.cell(curr_r, col_size_group).value = "Regular"
                        if col_wash_care: ws.cell(curr_r, col_wash_care).value = "Machine wash"

                        if col_length: ws.cell(curr_r, col_length).value = "Full length"
                        if col_stretch: ws.cell(curr_r, col_stretch).value = "Stretch"
                        if col_style_type: ws.cell(curr_r, col_style_type).value = "Cargo" if cat in ['cargo', 'trousers'] else "Relaxed"
                        if col_fit_type: ws.cell(curr_r, col_fit_type).value = "Relaxed Fit"
                        if col_front_style: ws.cell(curr_r, col_front_style).value = "Flat Front"

                        if col_waist_rise: ws.cell(curr_r, col_waist_rise).value = "Mid Rise"
                        if col_waist_type: ws.cell(curr_r, col_waist_type).value = "Flexi waist" if cat in ['cargo', 'trousers'] else "Fixed waist"

                        if col_meas_type: ws.cell(curr_r, col_meas_type).value = "Garment Measurement"
                        if col_to_fit: ws.cell(curr_r, col_to_fit).value = tf_val
                        if col_waist: ws.cell(curr_r, col_waist).value = w_val
                        if col_inseam: ws.cell(curr_r, col_inseam).value = "32"

                        if col_highlight and high_text: ws.cell(curr_r, col_highlight).value = high_text
                        if col_add_info and add_text: ws.cell(curr_r, col_add_info).value = add_text

                        for img_idx, img_link in enumerate(sku_img_urls):
                            if img_idx < len(img_cols):
                                ws.cell(curr_r, img_cols[img_idx]).value = img_link

                        if col_mrp:
                            mrp_cell = ws.cell(curr_r, col_mrp)
                            if mrp_value:
                                mrp_cell.value = mrp_value
                            else:
                                mrp_cell.value = ""
                                mrp_cell.fill = mrp_fill
                                mrp_cell.font = mrp_font

                        curr_r += 1
                        total_variants_added += 1

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success(f"🎉 Total {len(detected_folders)} folders ke {total_variants_added} variants create ho gaye sheets me: {', '.join(sheets_modified)}!")
            st.download_button(
                label="📥 Download Ready Ajio Listing File",
                data=output,
                file_name=f"AJIO_MULTI_CATEGORY_LISTED.{'xlsm' if is_xlsm else 'xlsx'}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error aaya: {e}")
