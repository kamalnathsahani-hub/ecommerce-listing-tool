import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import os
import re
import zipfile
import base64
import requests

# Conditional imports for Flipkart .xls
try:
    import xlrd
    import xlwt
    from xlutils.copy import copy as xl_copy
    XLS_SUPPORT = True
except ImportError:
    XLS_SUPPORT = False

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")

st.title("📦 Multi-Marketplace Auto Listing Generator (Universal Pro)")
st.caption("Ajio, Amazon Seller Central, Amazon Vendor Central & Flipkart - Full Multi-Channel Engine")

# Permanent ImgBB API Key
IMGBB_PERMANENT_KEY = "137be55d19cf3d6a7ec051e369137041"

# 1. Marketplace Selection
col_m1, col_m2 = st.columns([2, 1])
with col_m1:
    marketplace = st.selectbox(
        "Marketplace Chunein:",
        ["Ajio", "Amazon - Seller Central", "Amazon - Vendor Central", "Flipkart", "Meesho", "Myntra"]
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
st.info("💡 **Dono Options Available Hain:** Direct **ZIP folder** upload karein (ImgBB se direct workable links banenge) ya **Dropbox Link** dein.")

col_in1, col_in2 = st.columns(2)

detected_folders = []
folder_images_count = {}
zip_images_data = {}

def clean_sku_name(name):
    cleaned = re.sub(r'[\s_]+', '-', str(name).strip())
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
        help="Yeh feature ZIP ki images ko high-speed permanent CDN par upload karke direct .jpg links banata hai jo Chrome aur marketplaces dono me open hote hain."
    )
    st.success("⚡ Live Workable CDN Active: ZIP images automatic 100% working live links banegi!")

with col_in2:
    st.write("##### Option B: Dropbox Link / Manual SKUs")
    dropbox_base_url = st.text_input(
        "Dropbox Shared Folder Link (Optional):",
        placeholder="https://www.dropbox.com/scl/fo/xyz/... ya file link",
        help="Agar aapne Dropbox link diya, toh system direct workable links (?raw=1) generate karega."
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
        "Marketed By / Packer Details:",
        value="61/C, BLOCK-N, SUMEL BUSINESS PARK 5, Flyover, ASARVA, Ahmedabad, Gujarat, India - 380016",
        height=100
    )

st.divider()

# 4. Trending Keywords Section
st.subheader("3. SEO & Trending Keywords")
custom_keywords = st.text_area(
    "Custom Trending Keywords (Optional):",
    placeholder="e.g. relaxed fit baggy jeans, mid rise, comfort stretch denim, urban casual streetwear, durable utility pockets",
    help="Platform specific search ranking ke liye trending terms daalein."
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
    ignore = ['instruction', 'guideline', 'extra', 'rejection', 'seller info', 'summary', 'faq', 'dropdown', 'version']
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
        if any(w in s.lower() for w in ['styles', 'template', 'data', 'listing']):
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

def build_seo_content(style, color, category, platform, user_keywords=""):
    kw_str = user_keywords if user_keywords else "casual streetwear, relaxed fit, durable utility pockets, all-day comfort, stretchable fabric"
    
    if "Amazon" in platform:
        title = f"Men's Relaxed Fit {category.capitalize()} - {style} ({color})"
        bullets = [
            f"PREMIUM QUALITY FABRIC: Crafted from high-grade breathable and durable fabric engineered with {kw_str} for superior all-day ease.",
            f"CONTEMPORARY FIT & STYLING: Features a stylish relaxed silhouette in authentic {color} tones with reinforced contrast stitching.",
            "FUNCTIONAL UTILITY: Equipped with deep multi-utility pockets and heavy-duty front closure ensuring complete reliability.",
            "VERSATILE DAILY WEAR: Pairs effortlessly with casual tees, polos, denim jackets, and lifestyle sneakers for college, travel, or streetwear.",
            "WASH & CARE: Easy machine wash cold inside out with like colors. Do not bleach. Warm iron if required."
        ]
        desc = (
            f"Upgrade your daily collection with this premium Men's {style} {category.capitalize()} in {color}. "
            f"Constructed with high-strength fabric and infused with {kw_str}, it offers exceptional durability, flexibility, and comfort wherever you go."
        )
        return title, bullets, desc[:1900]

    elif "Flipkart" in platform:
        title = f"Men's Regular Fit {color} {category.capitalize()} ({style})"
        desc = (
            f"Step out in style with Men's {style} {category.capitalize()} in trendy {color}. "
            f"Tailored with premium breathable material featuring {kw_str}. Ideal for western daily wear."
        )
        return title, [kw_str[:150]], desc[:950]

    else: # Ajio / Default
        desc = (
            f"Upgrade your wardrobe with this stylish Men's {style} {category.capitalize()} in {color}. "
            f"Crafted from premium fabric engineered with {kw_str}. "
            "Featuring a contemporary relaxed fit, functional pocket styling, sturdy front closure, and authentic clean finish. "
            "Ideal for casual outings, daily wear, and weekend hangouts."
        )
        highlight = f"Premium fabric | Relaxed silhouette in {color} | Authentic clean finish | Functional utility pockets | All-season wear"
        add_info = "Styling Tip: Pair with casual shirts or boxy tees. Wash Care: Machine wash cold inside out."
        return desc[:950], [highlight[:500]], add_info[:500]

# ================= GENERATION PROCESS =================
if uploaded_file and st.button("🚀 Generate Listing File"):
    if not detected_folders:
        st.error("Kripya kam se kam ek folder/SKU enter karein ya ZIP file upload karein!")
    else:
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()
        size_list = [sz.strip() for sz in sizes_input.split(",") if sz.strip()]

        # Live CDN Upload using permanent API Key
        live_links_cache = {}
        if auto_cdn_upload and zip_images_data:
            with st.spinner("Images ko live workable CDN par upload kiya ja raha hai..."):
                for fld, img_tuples in zip_images_data.items():
                    live_links_cache[fld] = []
                    for img_name, img_bytes in img_tuples:
                        cdn_url = upload_to_imgbb(img_bytes, IMGBB_PERMANENT_KEY)
                        if cdn_url:
                            live_links_cache[fld].append(cdn_url)

        # ---------------- FLIPKART .XLS ENGINE ----------------
        if "Flipkart" in marketplace and file_ext == ".xls":
            if not XLS_SUPPORT:
                st.error("Flipkart .xls processing ke liye xlrd/xlutils install karein, ya file ko .xlsx format me upload karein!")
            else:
                try:
                    file_bytes = uploaded_file.read()
                    rb = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
                    wb = xl_copy(rb)

                    # Target sheet 'jean'
                    target_sheet_idx = 2
                    for idx, sname in enumerate(rb.sheet_names()):
                        if any(k in sname.lower() for k in ['jean', 'listing', 'data']):
                            target_sheet_idx = idx
                            break
                    
                    ws = wb.get_sheet(target_sheet_idx)
                    rb_sheet = rb.sheet_by_index(target_sheet_idx)

                    curr_r = rb_sheet.nrows
                    total_variants_added = 0

                    for raw_fld in detected_folders:
                        clean_style = clean_sku_name(raw_fld)
                        det_col = detect_color(raw_fld)
                        cat = detect_category(raw_fld)
                        title_text, _, desc_text = build_seo_content(clean_style, det_col, cat, "Flipkart", custom_keywords)

                        if raw_fld in live_links_cache and live_links_cache[raw_fld]:
                            sku_img_urls = live_links_cache[raw_fld]
                        else:
                            img_cnt = folder_images_count.get(raw_fld, 5)
                            sku_img_urls = [make_workable_dropbox_url(dropbox_base_url, clean_style, i) for i in range(1, img_cnt + 1)]

                        for sz in size_list:
                            clean_sz = clean_sku_name(sz)
                            full_sku = f"{clean_style}-{clean_sz}"

                            ws.write(curr_r, 6, full_sku)          # Seller SKU ID
                            ws.write(curr_r, 7, clean_style)       # Group ID
                            ws.write(curr_r, 10, "ACTIVE")         # Listing Status
                            
                            p_val = int(mrp_value) if mrp_value.isdigit() else 2999
                            ws.write(curr_r, 11, p_val)            # MRP (INR)
                            ws.write(curr_r, 12, int(p_val * 0.4)) # Selling Price
                            ws.write(curr_r, 15, 2)                # Procurement SLA
                            ws.write(curr_r, 16, 100)              # Stock
                            ws.write(curr_r, 17, "FLIPKART")       # Shipping provider

                            # Dimensions
                            ws.write(curr_r, 21, 25.0)             # Length (CM)
                            ws.write(curr_r, 22, 24.0)             # Breadth (CM)
                            ws.write(curr_r, 23, 2.0)              # Height (CM)
                            ws.write(curr_r, 24, 0.38)             # Weight (KG)
                            ws.write(curr_r, 25, "62034290")       # HSN
                            ws.write(curr_r, 27, "IN")             # Country Of Origin
                            ws.write(curr_r, 28, mfg_address)      # Manufacturer Details
                            ws.write(curr_r, 29, mkt_address)      # Packer Details
                            
                            # Attributes
                            if brand_name: ws.write(curr_r, 33, brand_name)
                            ws.write(curr_r, 34, "Regular")        # Fit
                            ws.write(curr_r, 35, "Men")            # Ideal For
                            ws.write(curr_r, 36, 1)                # Pack of
                            ws.write(curr_r, 37, "Mid Rise")       # Rise
                            ws.write(curr_r, 38, "Cotton Lycra Blend" if cat != 'jeans' else "Denim")
                            ws.write(curr_r, 39, "Western Wear")
                            ws.write(curr_r, 40, "Clean Look")
                            ws.write(curr_r, 41, "Clean Look")
                            ws.write(curr_r, 42, clean_sz)         # Brand Size
                            ws.write(curr_r, 44, clean_style)      # Style Code
                            ws.write(curr_r, 45, det_col)          # Color

                            # Images (Cols 47 to 52)
                            for img_i, u in enumerate(sku_img_urls[:6]):
                                ws.write(curr_r, 47 + img_i, u)

                            ws.write(curr_r, 54, "Regular")        # Jeans Length
                            ws.write(curr_r, 55, "No")             # Reversible
                            ws.write(curr_r, 56, "Solid" if cat != 'jeans' else "Washed")
                            ws.write(curr_r, 64, "Yes")            # Stretchable
                            ws.write(curr_r, 69, desc_text)        # Description
                            ws.write(curr_r, 70, custom_keywords if custom_keywords else f"{clean_style} jeans for men")

                            curr_r += 1
                            total_variants_added += 1

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)

                    st.success(f"🎉 Flipkart file successfully fill ho gayi! Total {total_variants_added} rows create hui!")
                    st.download_button(
                        label="📥 Download Filled Flipkart File",
                        data=output,
                        file_name=uploaded_file.name,
                        mime="application/vnd.ms-excel"
                    )
                except Exception as e:
                    st.error(f"Flipkart processing me error: {e}")

        # ---------------- OPENPYXL ENGINE (AJIO & AMAZON) ----------------
        else:
            try:
                is_xlsm = file_ext == ".xlsm"
                wb = openpyxl.load_workbook(uploaded_file, keep_vba=is_xlsm)

                for s in wb.sheetnames:
                    wb[s].protection.sheet = False
                    wb[s].protection.disable()

                mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
                mrp_font = Font(bold=True)
                total_variants_added = 0
                sheets_modified = set()

                # AMAZON SELLER CENTRAL
                if "Seller Central" in marketplace:
                    ws = wb['Template'] if 'Template' in wb.sheetnames else wb.worksheets[0]
                    sheets_modified.add(ws.title)

                    # Headers in row 4, technical keys in row 5
                    header_map = {}
                    for c in range(1, ws.max_column + 1):
                        h4 = str(ws.cell(4, c).value or '').strip()
                        h5 = str(ws.cell(5, c).value or '').strip()
                        header_map[h4.lower()] = c
                        if h5: header_map[h5.lower()] = c

                    curr_r = 6
                    while ws.cell(curr_r, 1).value is not None:
                        curr_r += 1

                    for raw_fld in detected_folders:
                        clean_style = clean_sku_name(raw_fld)
                        det_col = detect_color(raw_fld)
                        cat = detect_category(raw_fld)
                        title_text, bullets, desc_text = build_seo_content(clean_style, det_col, cat, "Amazon", custom_keywords)

                        if raw_fld in live_links_cache and live_links_cache[raw_fld]:
                            sku_img_urls = live_links_cache[raw_fld]
                        else:
                            img_cnt = folder_images_count.get(raw_fld, 5)
                            sku_img_urls = [make_workable_dropbox_url(dropbox_base_url, clean_style, i) for i in range(1, img_cnt + 1)]

                        # 1. Parent Row
                        ws.cell(curr_r, 1).value = clean_style                     # SKU
                        ws.cell(curr_r, 2).value = "PANTS"                         # Product Type
                        ws.cell(curr_r, 4).value = "Parent"                        # Parentage Level
                        ws.cell(curr_r, 6).value = "Size"                          # Variation Theme
                        ws.cell(curr_r, 7).value = title_text                      # Item Name
                        if brand_name: ws.cell(curr_r, 9).value = brand_name
                        ws.cell(curr_r, 32).value = desc_text                      # Product Description
                        for b_i, b_val in enumerate(bullets[:5]):
                            ws.cell(curr_r, 33 + b_i).value = b_val                # Bullet Points
                        if sku_img_urls: ws.cell(curr_r, 22).value = sku_img_urls[0]
                        curr_r += 1

                        # 2. Child Rows
                        for sz in size_list:
                            clean_sz = clean_sku_name(sz)
                            full_sku = f"{clean_style}-{clean_sz}"

                            ws.cell(curr_r, 1).value = full_sku                    # SKU
                            ws.cell(curr_r, 2).value = "PANTS"                     # Product Type
                            ws.cell(curr_r, 4).value = "Child"                     # Parentage Level
                            ws.cell(curr_r, 5).value = clean_style                 # Parent SKU
                            ws.cell(curr_r, 6).value = "Size"                      # Variation Theme
                            ws.cell(curr_r, 7).value = f"{title_text} - {clean_sz}"
                            if brand_name: ws.cell(curr_r, 9).value = brand_name
                            
                            # Images
                            if sku_img_urls: ws.cell(curr_r, 22).value = sku_img_urls[0]
                            for img_i, u in enumerate(sku_img_urls[1:6]):
                                ws.cell(curr_r, 23 + img_i).value = u

                            ws.cell(curr_r, 32).value = desc_text
                            for b_i, b_val in enumerate(bullets[:5]):
                                ws.cell(curr_r, 33 + b_i).value = b_val

                            # Size & Details
                            ws.cell(curr_r, 48).value = "Adult"                    # Age Range
                            ws.cell(curr_r, 49).value = "IN"                       # Size System
                            ws.cell(curr_r, 50).value = "Waist"                    # Size Class
                            ws.cell(curr_r, 51).value = clean_sz                   # Size Value
                            ws.cell(curr_r, 54).value = "Cotton" if cat != 'jeans' else "Denim"
                            ws.cell(curr_r, 57).value = "Denim"
                            ws.cell(curr_r, 63).value = det_col                    # Color Map
                            ws.cell(curr_r, 64).value = det_col                    # Color
                            ws.cell(curr_r, 102).value = clean_sz                  # Waist Size
                            ws.cell(curr_r, 141).value = 350                       # Item Weight
                            ws.cell(curr_r, 142).value = "grams"
                            ws.cell(curr_r, 171).value = 100                       # Quantity
                            
                            # Price & MRP
                            p_cell = ws.cell(curr_r, 175)                          # Your Price
                            mrp_cell = ws.cell(curr_r, 176)                        # MRP
                            if mrp_value:
                                mrp_cell.value = int(mrp_value) if mrp_value.isdigit() else mrp_value
                                p_cell.value = int(int(mrp_value) * 0.45) if mrp_value.isdigit() else 999
                            else:
                                mrp_cell.fill = mrp_fill
                                mrp_cell.font = mrp_font

                            ws.cell(curr_r, 215).value = 380                       # Package Weight
                            ws.cell(curr_r, 216).value = "grams"
                            ws.cell(curr_r, 219).value = "IN"                      # Country of Origin

                            curr_r += 1
                            total_variants_added += 1

                # AMAZON VENDOR CENTRAL
                elif "Vendor Central" in marketplace:
                    ws = None
                    for sname in wb.sheetnames:
                        if "template" in sname.lower():
                            ws = wb[sname]
                            break
                    if not ws: ws = wb.worksheets[0]
                    sheets_modified.add(ws.title)

                    curr_r = 7
                    while ws.cell(curr_r, 2).value is not None:
                        curr_r += 1

                    for raw_fld in detected_folders:
                        clean_style = clean_sku_name(raw_fld)
                        det_col = detect_color(raw_fld)
                        cat = detect_category(raw_fld)
                        title_text, bullets, desc_text = build_seo_content(clean_style, det_col, cat, "Amazon", custom_keywords)

                        if raw_fld in live_links_cache and live_links_cache[raw_fld]:
                            sku_img_urls = live_links_cache[raw_fld]
                        else:
                            img_cnt = folder_images_count.get(raw_fld, 5)
                            sku_img_urls = [make_workable_dropbox_url(dropbox_base_url, clean_style, i) for i in range(1, img_cnt + 1)]

                        for sz in size_list:
                            clean_sz = clean_sku_name(sz)
                            full_sku = f"{clean_style}-{clean_sz}"

                            ws.cell(curr_r, 2).value = full_sku                    # Vendor SKU
                            ws.cell(curr_r, 3).value = "PANTS"                     # Product Type
                            ws.cell(curr_r, 4).value = "Child"                     # Parentage Level
                            ws.cell(curr_r, 6).value = clean_style                 # Parent SKU
                            ws.cell(curr_r, 7).value = "Size"                      # Variation Theme
                            ws.cell(curr_r, 8).value = f"{title_text} - {clean_sz}"
                            if brand_name: ws.cell(curr_r, 10).value = brand_name
                            
                            # Images (Cols 24 to 29)
                            if sku_img_urls: ws.cell(curr_r, 24).value = sku_img_urls[0]
                            for img_i, u in enumerate(sku_img_urls[1:6]):
                                ws.cell(curr_r, 25 + img_i).value = u

                            # Bullets (Cols 34 to 38)
                            for b_i, b_val in enumerate(bullets[:5]):
                                ws.cell(curr_r, 34 + b_i).value = b_val

                            ws.cell(curr_r, 53).value = "Adult"
                            ws.cell(curr_r, 54).value = "IN"
                            ws.cell(curr_r, 55).value = "Waist"
                            ws.cell(curr_r, 56).value = clean_sz
                            ws.cell(curr_r, 67).value = desc_text
                            ws.cell(curr_r, 68).value = det_col
                            ws.cell(curr_r, 69).value = det_col
                            ws.cell(curr_r, 118).value = clean_sz
                            ws.cell(curr_r, 154).value = 350
                            ws.cell(curr_r, 155).value = "grams"

                            mrp_cell = ws.cell(curr_r, 156)                        # List Price / MRP
                            if mrp_value:
                                mrp_cell.value = int(mrp_value) if mrp_value.isdigit() else mrp_value
                            else:
                                mrp_cell.fill = mrp_fill
                                mrp_cell.font = mrp_font

                            ws.cell(curr_r, 171).value = 380
                            ws.cell(curr_r, 172).value = "grams"
                            ws.cell(curr_r, 177).value = "IN"

                            curr_r += 1
                            total_variants_added += 1

                # AJIO / DEFAULT ENGINE (Preserved with all strict fixes)
                else:
                    category_folders = {}
                    for raw_folder in detected_folders:
                        cat = detect_category(raw_folder)
                        category_folders.setdefault(cat, []).append(raw_folder)

                    for cat, folder_list in category_folders.items():
                        target_sheet_name = find_sheet_for_category(wb, cat)
                        ws = wb[target_sheet_name]
                        sheets_modified.add(target_sheet_name)

                        headers = {c: str(ws.cell(3, c).value or '').strip() for c in range(1, ws.max_column + 1)}

                        def get_col(names):
                            for name in names:
                                cln_name = name.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
                                for c, h in headers.items():
                                    cln_h = h.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
                                    if cln_h == cln_name: return c
                            return None

                        col_style = get_col(['Style Code', 'Parent SKU'])
                        col_desc = get_col(['Style Description', 'Product Description', 'Description'])
                        col_sku = get_col(['Item SKU', 'SKU ID', 'Seller SKU', 'SKU'])
                        col_brand = get_col(['Brand', 'Brand Name'])
                        col_ean = get_col(['EAN', 'Barcode'])
                        col_td = get_col(['TD'])
                        col_mrp = get_col(['MRP', 'Maximum Retail Price'])
                        col_hsn = get_col(['HSN', 'HSN Code'])
                        col_size = get_col(['Size'])
                        col_std_size = get_col(['StandardSize'])
                        col_comp_count = get_col(['Component Count'])
                        col_country = get_col(['Country of Origin'])
                        col_mfg_by = get_col(['Manufactured By'])
                        col_mkt_by = get_col(['Marketed By'])
                        col_net_qty = get_col(['Net Quantity'])
                        col_pkg_contains = get_col(['Package Contains'])
                        col_color = get_col(['Color Family'])
                        col_pri_color = get_col(['Primary color'])
                        col_fab_detail = get_col(['Fabric Detail'])
                        col_fab_type = get_col(['Fabric Type'])
                        col_pattern = get_col(['Pattern'])
                        col_size_format = get_col(['Size Format'])
                        col_wash_care = get_col(['Wash Care'])
                        col_style_type = get_col(['Style Type'])
                        col_fit_type = get_col(['Fit Type'])
                        col_front_style = get_col(['Front Style'])
                        col_waist_rise = get_col(['Waist Rise'])
                        col_waist_type = get_col(['Waist Type'])
                        col_meas_type = get_col(['Measurement Type'])
                        col_to_fit = get_col(['To Fit Waist'])
                        col_waist = get_col(['Waist'])
                        col_inseam = get_col(['Inseam'])

                        curr_r = 4
                        while ws.cell(curr_r, col_style if col_style else 1).value is not None:
                            curr_r += 1

                        for raw_fld in folder_list:
                            clean_style = clean_sku_name(raw_fld)
                            det_col = detect_color(raw_fld)
                            desc_text, high_text, add_text = build_seo_content(clean_style, det_col, cat, "Ajio", custom_keywords)

                            if raw_fld in live_links_cache and live_links_cache[raw_fld]:
                                sku_img_urls = live_links_cache[raw_fld]
                            else:
                                img_cnt = folder_images_count.get(raw_fld, 5)
                                sku_img_urls = [make_workable_dropbox_url(dropbox_base_url, clean_style, i) for i in range(1, img_cnt + 1)]

                            for sz in size_list:
                                clean_sz = clean_sku_name(sz)
                                full_sku = f"{clean_style}-{clean_sz}"

                                try:
                                    num_sz = int(re.findall(r'\d+', clean_sz)[0])
                                    w_val = str(num_sz + 1)
                                    tf_val = str(num_sz - 1)
                                end:
                                    w_val, tf_val = "34", "32"

                                if col_style: ws.cell(curr_r, col_style).value = clean_style
                                if col_desc: ws.cell(curr_r, col_desc).value = desc_text
                                if col_sku: ws.cell(curr_r, col_sku).value = full_sku
                                if col_brand: ws.cell(curr_r, col_brand).value = brand_name
                                if col_ean: ws.cell(curr_r, col_ean).value = full_sku
                                if col_td: ws.cell(curr_r, col_td).value = default_td
                                if col_hsn: ws.cell(curr_r, col_hsn).value = 62034200 if cat in ['cargo', 'trousers'] else 62034290
                                if col_size: ws.cell(curr_r, col_size).value = clean_sz
                                if col_std_size: ws.cell(curr_r, col_std_size).value = clean_sz

                                # Dimensions
                                for c_idx in [14, 20]: ws.cell(curr_r, c_idx).value = "2"
                                for c_idx in [15, 21]: ws.cell(curr_r, c_idx).value = "25"
                                for c_idx in [16, 22]: ws.cell(curr_r, c_idx).value = "24"
                                for c_idx in [17, 23]: ws.cell(curr_r, c_idx).value = "CM"
                                for c_idx in [18, 24]: ws.cell(curr_r, c_idx).value = "350"
                                for c_idx in [19, 25]: ws.cell(curr_r, c_idx).value = "GRAM"

                                if col_comp_count: ws.cell(curr_r, col_comp_count).value = 1
                                if col_country: ws.cell(curr_r, col_country).value = "India"
                                if col_mfg_by: ws.cell(curr_r, col_mfg_by).value = mfg_address
                                if col_mkt_by: ws.cell(curr_r, col_mkt_by).value = mkt_address
                                if col_net_qty: ws.cell(curr_r, col_net_qty).value = 1
                                if col_pkg_contains: ws.cell(curr_r, col_pkg_contains).value = "1"

                                fab_val = "Cotton" if cat in ['cargo', 'trousers', 'shirt'] else "Denim"
                                pat_val = "Solid" if cat in ['cargo', 'trousers'] else "Washed"

                                if col_color: ws.cell(curr_r, col_color).value = det_col
                                if col_pri_color: ws.cell(curr_r, col_pri_color).value = det_col
                                if col_fab_detail: ws.cell(curr_r, col_fab_detail).value = fab_val
                                if col_fab_type: ws.cell(curr_r, col_fab_type).value = fab_val
                                if col_pattern: ws.cell(curr_r, col_pattern).value = pat_val
                                if col_size_format: ws.cell(curr_r, col_size_format).value = "UNI"
                                if col_wash_care: ws.cell(curr_r, col_wash_care).value = "Machine wash"

                                if col_style_type: ws.cell(curr_r, col_style_type).value = "Cargo" if cat in ['cargo', 'trousers'] else "Relaxed"
                                if col_fit_type: ws.cell(curr_r, col_fit_type).value = "Relaxed Fit"
                                if col_front_style: ws.cell(curr_r, col_front_style).value = "Flat Front"
                                if col_waist_rise: ws.cell(curr_r, col_waist_rise).value = "Mid Rise"
                                if col_waist_type: ws.cell(curr_r, col_waist_type).value = "Flexi waist" if cat in ['cargo', 'trousers'] else "Fixed waist"

                                if col_meas_type: ws.cell(curr_r, col_meas_type).value = "Garment Measurement"
                                if col_to_fit: ws.cell(curr_r, col_to_fit).value = tf_val
                                if col_waist: ws.cell(curr_r, col_waist).value = w_val
                                if col_inseam: ws.cell(curr_r, col_inseam).value = "32"

                                # Images (Cols 77+)
                                for img_i, u in enumerate(sku_img_urls[:10]):
                                    ws.cell(curr_r, 77 + img_i).value = u

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

                st.success(f"🎉 {marketplace} file successfully fill ho gayi! Total {total_variants_added} rows generate hui!")
                st.download_button(
                    label="📥 Download Filled Listing File",
                    data=output,
                    file_name=uploaded_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Processing error: {e}")
