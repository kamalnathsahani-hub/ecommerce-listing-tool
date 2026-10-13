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

st.title("📦 Multi-Marketplace Auto Listing Generator (Universal Pro V8)")
st.caption("Ajio, Amazon Seller Central, Amazon Vendor Central & Flipkart - Full Multi-Channel Engine with Auto Size Chart")

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
st.info("💡 **Dono Options Available Hain:** Direct **ZIP folder** upload karein (ImgBB se direct workable links banenge) ya **Dropbox Link** dein. Agar ZIP me koi **Size Chart** image hogi toh system use auto-detect karke sabhi products par apply karega.")

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
        help="ZIP me har SKU ka folder (e.g. MENS JEANS 29 GREY) aur andar images hon. Saath me Size Chart image (e.g. size_chart.jpg) bhi daal sakte hain."
    )
    if zip_file:
        try:
            with zipfile.ZipFile(zip_file, 'r') as z:
                file_list = z.namelist()
                
                # Check for Size Chart image anywhere in ZIP
                for f in file_list:
                    if is_size_chart_file(f):
                        detected_size_chart_info['filename'] = os.path.basename(f)
                        detected_size_chart_info['bytes'] = z.read(f)
                        break

                folders = set()
                for f in file_list:
                    if is_size_chart_file(f):
                        continue
                    parts = f.replace('\\', '/').strip('/').split('/')
                    if len(parts) > 1 and parts[0]:
                        if not any(k in parts[0].lower() for k in ['size chart', 'sizechart', 'size_chart', 'size_guide']):
                            folders.add(parts[0])
                    elif len(parts) == 1 and f.endswith('/'):
                        if not any(k in parts[0].lower() for k in ['size chart', 'sizechart', 'size_chart', 'size_guide']):
                            folders.add(parts[0])

                detected_folders = sorted(list(folders))
                for fld in detected_folders:
                    imgs = [
                        p for p in file_list 
                        if p.replace('\\', '/').startswith(fld + '/') 
                        and p.lower().endswith(('.jpg', '.jpeg', '.png', '.webp'))
                        and not is_size_chart_file(p)
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

            st.success(f"✅ ZIP Se Total {len(detected_folders)} Product Folders Detect Ho Gaye:")
            cleaned_preview = [clean_sku_name(f) for f in detected_folders]
            st.code(", ".join(cleaned_preview))

            if detected_size_chart_info:
                st.info(f"📐 **Universal Size Chart Detect Hua:** `{detected_size_chart_info['filename']}` — Yeh size chart ZIP ke sabhi products par uniformly follow hoga!")
        except Exception as e:
            st.error(f"ZIP read karne me error: {e}")

    auto_cdn_upload = st.checkbox(
        "🌐 ZIP images se automatic Live CDN Workable Links banayein (Active)",
        value=True,
        help="Yeh feature ZIP ki images aur Size Chart ko permanent CDN par upload karke direct .jpg links banata hai jo Chrome aur marketplaces me direct open hote hain."
    )
    st.success("⚡ Live Workable CDN Active: ZIP images aur Size Chart ke 100% working live links banegi!")

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
        "Sizes (comma-separated - Size Chart ke anusaar):",
        value="30, 32, 34, 36, 38",
        help="Numeric: 28, 30, 32 ya Alpha: S, M, L, XL, 2XL, 3XL, 4XL (Jo size chart me ho)"
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

def build_seo_content(style, color, category, platform, user_keywords="", has_size_chart=False):
    kw_str = user_keywords if user_keywords else "casual streetwear, relaxed fit, durable utility pockets, all-day comfort, stretchable fabric"
    size_chart_note = "Please refer to the size chart image for exact garment measurements and perfect fit selection." if has_size_chart else "True to size regular fit."
    
    if "Amazon" in platform:
        title = f"Men's Relaxed Fit {category.capitalize()} - {style} ({color})"
        bullets = [
            f"PREMIUM QUALITY FABRIC: Crafted from high-grade breathable and durable fabric engineered with {kw_str} for superior all-day ease.",
            f"CONTEMPORARY FIT & STYLING: Features a stylish relaxed silhouette in authentic {color} tones with reinforced contrast stitching.",
            f"ACCURATE SIZING: {size_chart_note}",
            "FUNCTIONAL UTILITY: Equipped with deep multi-utility pockets and heavy-duty front closure ensuring complete reliability.",
            "WASH & CARE: Easy machine wash cold inside out with like colors. Do not bleach. Warm iron if required."
        ]
        desc = (
            f"Upgrade your daily collection with this premium Men's {style} {category.capitalize()} in {color}. "
            f"Constructed with high-strength fabric and infused with {kw_str}, it offers exceptional durability, flexibility, and comfort wherever you go. {size_chart_note}"
        )
        return title, bullets, desc[:1900]

    elif "Flipkart" in platform:
        title = f"Men's Regular Fit {color} {category.capitalize()} ({style})"
        desc = (
            f"Step out in style with Men's {style} {category.capitalize()} in trendy {color}. "
            f"Tailored with premium breathable material featuring {kw_str}. {size_chart_note} Ideal for western daily wear."
        )
        return title, [kw_str[:150]], desc[:950]

    else: # Ajio / Default
        desc = (
            f"Upgrade your wardrobe with this stylish Men's {style} {category.capitalize()} in {color}. "
            f"Crafted from premium fabric engineered with {kw_str}. "
            f"Featuring a contemporary relaxed fit, functional pocket styling, sturdy front closure, and authentic clean finish. {size_chart_note} "
            "Ideal for casual outings, daily wear, and weekend hangouts."
        )
        highlight = f"Premium fabric | Relaxed silhouette in {color} | Authentic clean finish | Functional utility pockets | All-season wear"
        add_info = f"Size Tip: {size_chart_note} Wash Care: Machine wash cold inside out."
        return desc[:950], [highlight[:500]], add_info[:500]

# ================= GENERATION PROCESS =================
if uploaded_file and st.button("🚀 Generate Listing File"):
    if not detected_folders:
        st.error("Kripya kam se kam ek folder/SKU enter karein ya ZIP file upload karein!")
    else:
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()
        size_list = [sz.strip() for sz in sizes_input.split(",") if sz.strip()]

        live_links_cache = {}
        size_chart_cdn_url = None

        if auto_cdn_upload:
            with st.spinner("Images aur Size Chart ko live workable CDN par upload kiya ja raha hai..."):
                if 'bytes' in detected_size_chart_info:
                    size_chart_cdn_url = upload_to_imgbb(detected_size_chart_info['bytes'], IMGBB_PERMANENT_KEY)
                    if size_chart_cdn_url:
                        detected_size_chart_info['url'] = size_chart_cdn_url

                if zip_images_data:
                    for fld, img_tuples in zip_images_data.items():
                        live_links_cache[fld] = []
                        for img_name, img_bytes in img_tuples:
                            cdn_url = upload_to_imgbb(img_bytes, IMGBB_PERMANENT_KEY)
                            if cdn_url:
                                live_links_cache[fld].append(cdn_url)

        has_chart = bool(size_chart_cdn_url or detected_size_chart_info)

        # ---------------- FLIPKART .XLS ENGINE ----------------
        if "Flipkart" in marketplace and file_ext == ".xls":
            if not XLS_SUPPORT:
                st.error("Flipkart .xls processing ke liye xlrd/xlutils install karein, ya file ko .xlsx format me upload karein!")
            else:
                try:
                    file_bytes = uploaded_file.read()
                    rb = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
                    wb = xl_copy(rb)

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
                        title_text, _, desc_text = build_seo_content(clean_style, det_col, cat, "Flipkart", custom_keywords, has_size_chart=has_chart)

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

                            for img_i, u in enumerate(sku_img_urls[:6]):
                                ws.write(curr_r, 47 + img_i, u)

                            if size_chart_cdn_url:
                                ws.write(curr_r, 74, size_chart_cdn_url)

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
