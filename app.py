import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import os
import re
import zipfile

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")

st.title("📦 Multi-Marketplace Auto Listing Generator (Pro V4)")
st.caption("Ajio Specialized Intelligence: Category Sheet Mapping, Auto-Hyphen SKUs, CARGO Standards & Strict Formatting")

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

# 2. Input Mode: ZIP Upload vs Manual / Dropbox
st.subheader("1. Product Images & Folders Input")
input_mode = st.radio(
    "Kaise Data Dena Chahte Hain?",
    ["📁 Option A: Upload Product Folders ZIP (Mix Categories - Jeans, Cargo, Shirts, etc.)", "🔗 Option B: Enter SKUs List & Dropbox Link Manually"],
    horizontal=True
)

detected_folders = []
folder_images_count = {}

col_left, col_right = st.columns(2)

def clean_sku_name(name):
    """Space ko '-' me badalna taaki koi space na rahe: e.g. MENS JEANS 29 GREY -> MENS-JEANS-29-GREY"""
    cleaned = re.sub(r'[\s_]+', '-', name.strip())
    # remove duplicate hyphens
    cleaned = re.sub(r'-+', '-', cleaned).strip('-')
    return cleaned

if "Option A" in input_mode:
    with col_left:
        st.write("##### Step A: Product Folders Ka ZIP Upload Karein")
        zip_file = st.file_uploader(
            "Product Folders Ka ZIP Upload Karein (.zip):",
            type=["zip"],
            help="ZIP me chahe Jeans, Cargo, Shirt, T-shirt sab mix folder hon. Har folder me 1.jpg, 2.jpg... images honi chahiye."
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
                        folder_images_count[fld] = len(imgs) if len(imgs) > 0 else 5

                st.success(f"✅ ZIP Se Total {len(detected_folders)} Folders Detect Ho Gaye:")
                cleaned_preview = [clean_sku_name(f) for f in detected_folders]
                st.code(", ".join(cleaned_preview))
            except Exception as e:
                st.error(f"ZIP read karne me error: {e}")

        dropbox_base_url = st.text_input(
            "Dropbox Shared Folder Link (Image Links Ke Liye):",
            placeholder="https://www.dropbox.com/scl/fo/xyz/...",
            help="Dropbox main link jisme folder ke andar 1.jpg, 2.jpg... images hain."
        )

    with col_right:
        st.write("##### Step B: Sizes & Attributes")
        sizes_input = st.text_input(
            "Sizes (comma-separated):",
            value="30, 32, 34, 36, 38",
            help="Numeric: 28, 30, 32 ya Alpha: S, M, L, XL, 2XL, 3XL, 4XL (CARGO format support)"
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

else:
    with col_left:
        skus_input = st.text_area(
            "Style Codes / Folder Names (Har line me ek):",
            value="MENS JEANS 29 GREY\nCARGO-01-04\nMENS SHIRT 12 WHITE",
            height=120,
            help="Space apne aap '-' me convert ho jayenge."
        )
        raw_list = [s.strip() for s in skus_input.split('\n') if s.strip()]
        detected_folders = raw_list
        dropbox_base_url = st.text_input(
            "Dropbox Main Shared Folder Link:",
            placeholder="https://www.dropbox.com/scl/fo/xyz/...",
            help="Dropbox main folder link."
        )
        images_count = st.number_input(
            "Har SKU ke folder me kitni images hain? (1, 2, 3...):",
            min_value=1,
            max_value=12,
            value=5
        )
        for fld in detected_folders:
            folder_images_count[fld] = images_count

    with col_right:
        sizes_input = st.text_input(
            "Sizes (comma-separated):",
            value="30, 32, 34, 36, 38"
        )
        brand_name = st.text_input(
            "Brand Name (Default Blank):",
            value="",
            placeholder="Khali rahega agar kuch nahi likhenge"
        )
        mrp_value = st.text_input(
            "MRP (Khali chhodne par Yellow Highlight hoga):",
            value="",
            placeholder="e.g. 2999 ya blank chhod dein"
        )

st.divider()

# 3. Trending Keywords Section
st.subheader("2. SEO & Trending Keywords")
custom_keywords = st.text_area(
    "Custom Trending Keywords (Optional):",
    placeholder="e.g. multi pocket cargo, streetwear casuals, relaxed fit, premium cotton blend, clean wash",
    help="Platform specific search ranking ke liye trending terms daalein."
)

st.divider()

# 4. Template Uploader
st.subheader("3. Template File & Generate")
uploaded_file = st.file_uploader(
    "Apni Excel Template (.xlsx ya .xlsm) Upload Karein:",
    type=["xlsx", "xlsm"]
)

# Helper: Detect Category from folder name (Jeans, Cargo, Shirt, T-Shirt)
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

# Helper: Find the exact sheet in workbook matching product category
def find_sheet_for_category(wb, category):
    ignore = ['instruction', 'guideline', 'extra', 'rejection', 'seller info']
    valid_sheets = [s for s in wb.sheetnames if not any(ign in s.lower() for ign in ignore)]

    if category == "cargo" or category == "trousers":
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

    # Fallback: Styles sheet
    for s in valid_sheets:
        if 'styles' in s.lower():
            return s
    return valid_sheets[0] if valid_sheets else wb.sheetnames[0]

# Helper: Color detection
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

def get_clean_dropbox_url(base_url, sku, img_num):
    if not base_url:
        return f"https://www.dropbox.com/s/{sku}/{img_num}.jpg?raw=1"
    clean = base_url.split("?")[0].rstrip("/")
    return f"{clean}/{sku}/{img_num}.jpg?raw=1"

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

if uploaded_file and st.button("🚀 Generate Ajio Multi-Category Listing File"):
    if not detected_folders:
        st.error("Kripya kam se kam ek folder/SKU enter karein ya ZIP file upload karein!")
    else:
        try:
            is_xlsm = uploaded_file.name.lower().endswith(".xlsm")
            wb = openpyxl.load_workbook(uploaded_file, keep_vba=is_xlsm)

            # Unprotect all sheets
            for s in wb.sheetnames:
                wb[s].protection.sheet = False
                wb[s].protection.disable()

            # Group folders by category
            category_folders = {}
            for raw_folder in detected_folders:
                cat = detect_category(raw_folder)
                category_folders.setdefault(cat, []).append(raw_folder)

            size_list = [sz.strip() for sz in sizes_input.split(",") if sz.strip()]
            mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
            mrp_font = Font(bold=True)

            total_variants_added = 0
            sheets_modified = set()

            for cat, folder_list in category_folders.items():
                target_sheet_name = find_sheet_for_category(wb, cat)
                ws = wb[target_sheet_name]
                sheets_modified.add(target_sheet_name)

                # Find header row
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

                def match_c(keys):
                    for c, h in headers.items():
                        cln = h.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
                        for k in keys:
                            if k.lower() in cln:
                                return c
                    return None

                col_style = match_c(['style code', 'parent sku'])
                col_sku = match_c(['item sku', 'sku id', 'seller sku', 'sku'])
                col_ean = match_c(['ean', 'barcode', 'upc'])
                col_td = match_c(['td'])
                col_size = match_c(['size', 'standard size', 'size format'])
                col_std_size = match_c(['standardsize'])
                col_brand = match_c(['brand', 'brand name'])
                col_mrp = match_c(['mrp', 'maximum retail price', 'standard price'])
                col_desc = match_c(['style description', 'product description', 'description', 'title'])
                col_hsn = match_c(['hsn', 'hsn code', 'tax code'])
                col_color = match_c(['color family', 'primary color', 'color'])
                col_pri_color = match_c(['primary color'])
                col_fabric = match_c(['fabric type', 'fabric', 'material'])
                col_fabric_det = match_c(['fabric detail'])
                col_wt_art = match_c(['articledimensionsunitweight', 'item weight', 'weight'])
                col_wt_pkg = match_c(['packagedimensionsweight', 'package weight', 'shipping weight'])
                col_wt_uom = match_c(['weightuom', 'weight unit'])
                col_dim_h = match_c(['height', 'articledimensionsunitheight'])
                col_dim_l = match_c(['length', 'articledimensionsunitlength'])
                col_dim_w = match_c(['width', 'articledimensionsunitwidth'])
                col_dim_uom = match_c(['lengthuom', 'dimension uom'])
                col_highlight = match_c(['highlight', 'bullet point 1', 'key features'])
                col_add_info = match_c(['additional information 1', 'bullet point 2'])
                col_waist = match_c(['waist'])
                col_to_fit = match_c(['to fit waist'])
                col_inseam = match_c(['inseam'])
                col_country = match_c(['country of origin'])
                col_prod_grp = match_c(['product groups'])
                col_fash_grp = match_c(['fashion groups'])
                col_season_cd = match_c(['season code'])
                col_season_yr = match_c(['season year'])
                col_fit_type = match_c(['fit type', 'style type'])
                col_front_style = match_c(['front style'])
                col_length = match_c(['length'])
                col_stretch = match_c(['stretch'])
                col_waist_rise = match_c(['waist rise'])
                col_waist_type = match_c(['waist type'])
                col_pkg_contains = match_c(['package contains'])
                col_net_qty = match_c(['net quantity'])

                img_cols = []
                for i in range(1, 15):
                    c_img = match_c([f'*model' if i==1 else f'model{i}', f'image {i}', f'image_{i}', f'main image' if i==1 else f'other image {i-1}', f'front image' if i==1 else f'back image'])
                    if c_img and c_img not in img_cols:
                        img_cols.append(c_img)

                curr_r = header_row + 1
                while ws.cell(curr_r, col_style if col_style else 1).value is not None:
                    curr_r += 1

                for raw_fld in folder_list:
                    # Clean Style Code (spaces to hyphens: e.g. MENS JEANS 29 GREY -> MENS-JEANS-29-GREY)
                    clean_style = clean_sku_name(raw_fld)
                    detected_col = detect_color(raw_fld)
                    desc_text, high_text, add_text = build_seo_content(clean_style, detected_col, cat, custom_keywords)

                    img_cnt = folder_images_count.get(raw_fld, 5)
                    sku_img_urls = [
                        get_clean_dropbox_url(dropbox_base_url, clean_style, i)
                        for i in range(1, img_cnt + 1)
                    ]

                    for sz in size_list:
                        clean_sz = clean_sku_name(sz)
                        full_sku = f"{clean_style}-{clean_sz}"

                        # Waist calculations
                        try:
                            num_sz = int(re.findall(r'\d+', clean_sz)[0])
                            w_val = str(num_sz + 1)
                            tf_val = str(num_sz - 1)
                        except:
                            w_val = "34"
                            tf_val = "32"

                        # Values mapping based on CARGO/Jeans rules
                        if col_style: ws.cell(curr_r, col_style).value = clean_style
                        if col_sku: ws.cell(curr_r, col_sku).value = full_sku
                        if col_ean: ws.cell(curr_r, col_ean).value = full_sku # Exact SKU match
                        if col_td: ws.cell(curr_r, col_td).value = default_td # By default 50
                        if col_size: ws.cell(curr_r, col_size).value = clean_sz
                        if col_std_size: ws.cell(curr_r, col_std_size).value = clean_sz
                        if col_brand: ws.cell(curr_r, col_brand).value = brand_name
                        if col_desc: ws.cell(curr_r, col_desc).value = desc_text
                        if col_hsn: ws.cell(curr_r, col_hsn).value = 62034200 if cat in ['cargo', 'trousers'] else 62034290
                        if col_prod_grp: ws.cell(curr_r, col_prod_grp).value = "Casual"
                        if col_fash_grp: ws.cell(curr_r, col_fash_grp).value = "Fashion"
                        if col_season_cd: ws.cell(curr_r, col_season_cd).value = "Core"
                        if col_season_yr: ws.cell(curr_r, col_season_yr).value = "2026"
                        if col_country: ws.cell(curr_r, col_country).value = "India"
                        if col_color: ws.cell(curr_r, col_color).value = detected_col
                        if col_pri_color: ws.cell(curr_r, col_pri_color).value = detected_col

                        # Fabric & Attributes
                        fab_val = "Cotton" if cat in ['cargo', 'trousers', 'shirt'] else "Denim"
                        if col_fabric: ws.cell(curr_r, col_fabric).value = fab_val
                        if col_fabric_det: ws.cell(curr_r, col_fabric_det).value = fab_val
                        if col_fit_type: ws.cell(curr_r, col_fit_type).value = "Relaxed Fit" if cat in ['cargo', 'trousers'] else "Relaxed"
                        if col_front_style: ws.cell(curr_r, col_front_style).value = "Flat Front"
                        if col_length: ws.cell(curr_r, col_length).value = "Full length"
                        if col_stretch: ws.cell(curr_r, col_stretch).value = "Stretch"
                        if col_waist_rise: ws.cell(curr_r, col_waist_rise).value = "Mid Rise"
                        if col_waist_type: ws.cell(curr_r, col_waist_type).value = "Flexi waist" if cat in ['cargo', 'trousers'] else "Fixed waist"
                        if col_pkg_contains: ws.cell(curr_r, col_pkg_contains).value = "1"
                        if col_net_qty: ws.cell(curr_r, col_net_qty).value = "1"

                        # Dimensions & Weight (CARGO standard 200g-350g, 25x24x2)
                        if col_wt_art: ws.cell(curr_r, col_wt_art).value = "350"
                        if col_wt_pkg: ws.cell(curr_r, col_wt_pkg).value = "380"
                        if col_wt_uom: ws.cell(curr_r, col_wt_uom).value = "GRAM"
                        if col_dim_h: ws.cell(curr_r, col_dim_h).value = "2"
                        if col_dim_l: ws.cell(curr_r, col_dim_l).value = "25"
                        if col_dim_w: ws.cell(curr_r, col_dim_w).value = "24"
                        if col_dim_uom: ws.cell(curr_r, col_dim_uom).value = "CM"

                        # Highlights & Add info
                        if col_highlight and high_text: ws.cell(curr_r, col_highlight).value = high_text
                        if col_add_info and add_text: ws.cell(curr_r, col_add_info).value = add_text
                        if col_waist: ws.cell(curr_r, col_waist).value = w_val
                        if col_to_fit: ws.cell(curr_r, col_to_fit).value = tf_val
                        if col_inseam: ws.cell(curr_r, col_inseam).value = "32"

                        # Image sequence
                        for img_idx, img_link in enumerate(sku_img_urls):
                            if img_idx < len(img_cols):
                                ws.cell(curr_r, img_cols[img_idx]).value = img_link

                        # Yellow MRP
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

            st.success(f"🎉 Zabardast! Total {len(detected_folders)} folders ke {total_variants_added} variants create ho gaye sheets me: {', '.join(sheets_modified)}!")
            st.download_button(
                label="📥 Download Ready Ajio Listing File",
                data=output,
                file_name=f"AJIO_MULTI_CATEGORY_LISTED.{'xlsm' if is_xlsm else 'xlsx'}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error aaya: {e}")
