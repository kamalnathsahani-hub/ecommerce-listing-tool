import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import re

st.set_page_config(page_title="Multi-SKU Marketplace Listing Generator", layout="wide")

st.title("📦 Multi-Marketplace Auto Listing Generator (Pro)")
st.caption("Multiple SKUs, Dropbox Image Automation, Smart SEO Keywords & Strict Template Limits")

# 1. Marketplace Selection
col_m1, col_m2 = st.columns([2, 1])
with col_m1:
    marketplace = st.selectbox(
        "Marketplace Chunein:",
        ["Ajio", "Meesho", "Flipkart", "Amazon", "Myntra"]
    )
with col_m2:
    product_category = st.selectbox(
        "Product Category:",
        ["Jeans / Denim", "Shirts", "T-Shirts", "Trousers / Chinos", "Kurta / Ethnic Wear"]
    )

st.divider()

# 2. Multi-SKU & Variant Details
st.subheader("1. Product & Multiple SKUs Configuration")

col_left, col_right = st.columns(2)

with col_left:
    skus_input = st.text_area(
        "Style Codes / SKUs List (Har line me ek SKU daalein):",
        value="SKU-1\nSKU-2\nSKU-4",
        height=100,
        help="Ek se zyada SKUs daalne ke liye nayi line (Enter) dabayein."
    )
    sizes_input = st.text_input(
        "Sizes (comma-separated - e.g. 30, 32, 34 ya XS, S, M, L, XL):",
        value="30, 32, 34, 36, 38",
        help="Jo sizes yahan likhenge, unke hisaab se SKU banega: SKU-1-30 ya SKU-1-XS"
    )
    mrp_value = st.text_input(
        "MRP (Khali chhodne par Yellow Highlight hoga):",
        value="",
        placeholder="e.g. 1999 ya blank chhod dein"
    )

with col_right:
    brand_name = st.text_input(
        "Brand Name (Default Blank rakha hai, chahein toh fill karein):",
        value="",
        placeholder="Blank rahega agar kuch nahi likhenge"
    )
    fabric_default = "Denim" if "Jeans" in product_category else "Cotton Blend"
    fabric_input = st.text_input(
        "Fabric (Default Auto-Selected - manual change kar sakte hain):",
        value=fabric_default
    )
    images_count = st.number_input(
        "Har SKU ke folder me kitni images hain? (1, 2, 3...):",
        min_value=1,
        max_value=12,
        value=5
    )

st.divider()

# 3. Dropbox Images & Trending Keywords Section
st.subheader("2. Dropbox Images & SEO Keywords")

col_drop, col_seo = st.columns(2)

with col_drop:
    dropbox_base_url = st.text_input(
        "Dropbox Main Shared Folder Link:",
        placeholder="https://www.dropbox.com/scl/fo/xyz/...",
        help="Dropbox folder link jisme har SKU ka subfolder ho."
    )
    st.info("💡 Dropbox link se image URLs sequence-wise banenge: Folder ke andar 1.jpg -> Main Image, 2.jpg -> Image 2, etc.")

with col_seo:
    custom_keywords = st.text_area(
        "Custom Trending Keywords (Optional):",
        placeholder="e.g. baggy jeans, cargo utility pockets, comfort stretch, streetwear, clean washed denim",
        help="Agar aapke paas trending search terms hain toh yahan daalein. System inko Title, Description aur Bullet points me mix kar dega."
    )

st.divider()

# 4. Template Uploader
st.subheader("3. Template File & Generate")
uploaded_file = st.file_uploader(
    "Apni Excel Template (.xlsx ya .xlsm) Upload Karein:",
    type=["xlsx", "xlsm"]
)

# Helper: Color detection from SKU name
COMMON_COLORS = [
    "Black", "Blue", "Navy", "Light Blue", "Dark Blue", "Ice Blue", 
    "Grey", "White", "Green", "Olive", "Beige", "Brown", "Coffee", "Charcoal"
]

def detect_color_from_sku(sku_name):
    sku_upper = sku_name.upper()
    for col in COMMON_COLORS:
        if col.upper() in sku_upper:
            return col
    return "Blue" if "JEANS" in sku_upper or "DENIM" in sku_upper else "Multicolor"

def get_clean_dropbox_url(base_url, sku, img_num):
    if not base_url:
        return f"https://www.dropbox.com/s/{sku}/{img_num}.jpg?raw=1"
    clean = base_url.split("?")[0].rstrip("/")
    return f"{clean}/{sku}/{img_num}.jpg?raw=1"

def build_seo_content(style, color, fabric, platform, user_keywords=""):
    base_kw = "baggy fit, streetwear, relaxed mid-rise, durable utility pockets, all-day stretch"
    kw_str = f"{user_keywords}, {base_kw}" if user_keywords else base_kw

    if platform == "Ajio":
        desc = (
            f"Upgrade your wardrobe with this stylish Men's {style} crafted from premium {fabric}. "
            f"Featuring an on-trend contemporary silhouette in {color} with modern detailing, "
            f"it is engineered with {kw_str} for peak durability and unmatched breathability. "
            "Equipped with functional pockets, a heavy-duty front closure, and reinforced stitching throughout. "
            "Ideal for casual outings, college wear, weekend get-togethers, and urban streetwear looks. "
            "Pair effortlessly with oversized t-shirts, casual shirts, and chunky sneakers for an elevated style statement."
        )
        highlight = f"Premium {fabric} fabric | Modern relaxed fit in {color} | Authentic clean finish | Multi-pocket utility styling | Easy machine wash"
        add_info = "Styling Tip: Pair with relaxed boxy tees and lifestyle sneakers. Wash Care: Machine wash cold with similar colors."
        return desc[:950], highlight[:500], add_info[:500]

    elif platform == "Meesho":
        desc = f"Trending Men's {style} in {color}. Fabric: Premium {fabric}. Fit: Relaxed Casual. Pattern: Solid/Washed. Features: {kw_str}."
        return desc[:450], "", ""

    elif platform == "Amazon":
        desc = f"Men's Casual {fabric} - {style} ({color}). Designed with {kw_str} for exceptional everyday comfort and longevity."
        highlight = f"High-quality {fabric} material | Breathable and flexible mobility | Classic {color} styling"
        return desc[:1900], highlight[:500], ""

    elif platform == "Flipkart":
        desc = f"Men's {style} in vibrant {color}. Made from premium {fabric} featuring {kw_str}. Perfect for casual and daily wear."
        return desc[:950], "", ""

    else: # Myntra
        desc = f"Men's {color} {fabric} {style}. Premium finish, tailored fit, featuring {kw_str}."
        return desc[:900], f"100% {fabric} | Classic {color}", ""

if uploaded_file and st.button("🚀 Generate Multi-SKU Listing File"):
    try:
        is_xlsm = uploaded_file.name.lower().endswith(".xlsm")
        wb = openpyxl.load_workbook(uploaded_file, keep_vba=is_xlsm)

        # 1. Unprotect all sheets
        for s in wb.sheetnames:
            wb[s].protection.sheet = False
            wb[s].protection.disable()

        # 2. Find target product sheet
        target_sheet = None
        for s in wb.sheetnames:
            s_low = s.lower()
            if any(k in s_low for k in ['styles', 'jeans', 'catalog', 'listing', 'template', 'data']) and not any(ign in s_low for ign in ['instruction', 'guideline', 'extra', 'rejection', 'seller info']):
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = wb.sheetnames[0]

        ws = wb[target_sheet]

        # 3. Find Header Row
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

        # Map core columns
        col_style = match_c(['style code', 'parent sku', 'product code'])
        col_sku = match_c(['item sku', 'sku id', 'seller sku', 'sku'])
        col_ean = match_c(['ean', 'barcode', 'upc'])
        col_size = match_c(['size', 'standard size', 'size format'])
        col_std_size = match_c(['standardsize'])
        col_brand = match_c(['brand', 'brand name'])
        col_mrp = match_c(['mrp', 'maximum retail price', 'standard price', 'listing price'])
        col_desc = match_c(['style description', 'product description', 'description', 'title', 'product name'])
        col_hsn = match_c(['hsn', 'hsn code', 'tax code'])
        col_color = match_c(['color family', 'primary color', 'color'])
        col_fabric = match_c(['fabric type', 'fabric', 'material'])
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

        # Image columns (1 to 12)
        img_cols = []
        for i in range(1, 15):
            c_img = match_c([f'*model' if i==1 else f'model{i}', f'image {i}', f'image_{i}', f'main image' if i==1 else f'other image {i-1}', f'front image' if i==1 else f'back image'])
            if c_img and c_img not in img_cols:
                img_cols.append(c_img)

        # Style Rows
        mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
        mrp_font = Font(bold=True)

        curr_r = header_row + 1
        while ws.cell(curr_r, col_style if col_style else 1).value is not None:
            curr_r += 1

        sku_list = [s.strip() for s in skus_input.split("\n") if s.strip()]
        size_list = [sz.strip() for sz in sizes_input.split(",") if sz.strip()]

        total_rows_created = 0

        for sku_item in sku_list:
            detected_color = detect_color_from_sku(sku_item)
            desc_text, high_text, add_text = build_seo_content(
                sku_item, detected_color, fabric_input, marketplace, custom_keywords
            )

            # Generate Dropbox URLs for this SKU (1.jpg, 2.jpg ...)
            sku_img_urls = [
                get_clean_dropbox_url(dropbox_base_url, sku_item, i) 
                for i in range(1, images_count + 1)
            ]

            for sz in size_list:
                full_sku = f"{sku_item}-{sz}"
                
                # Waist calculation if numeric
                try:
                    num_sz = int(sz)
                    w_val = str(num_sz + 1)
                    tf_val = str(num_sz - 1)
                except:
                    w_val = "33"
                    tf_val = "31"

                # Core data
                if col_style: ws.cell(curr_r, col_style).value = sku_item
                if col_sku: ws.cell(curr_r, col_sku).value = full_sku
                if col_ean: ws.cell(curr_r, col_ean).value = full_sku
                if col_size: ws.cell(curr_r, col_size).value = sz
                if col_std_size: ws.cell(curr_r, col_std_size).value = sz
                if col_brand: ws.cell(curr_r, col_brand).value = brand_name # Blank if left empty
                if col_desc: ws.cell(curr_r, col_desc).value = desc_text
                if col_hsn: ws.cell(curr_r, col_hsn).value = "62034290"
                if col_color: ws.cell(curr_r, col_color).value = detected_color
                if col_fabric: ws.cell(curr_r, col_fabric).value = fabric_input
                if col_country: ws.cell(curr_r, col_country).value = "India"

                # Highlights & Add info
                if col_highlight and high_text: ws.cell(curr_r, col_highlight).value = high_text
                if col_add_info and add_text: ws.cell(curr_r, col_add_info).value = add_text
                if col_waist: ws.cell(curr_r, col_waist).value = w_val
                if col_to_fit: ws.cell(curr_r, col_to_fit).value = tf_val
                if col_inseam: ws.cell(curr_r, col_inseam).value = "32"

                # Weight strictly 300g to 400g
                if col_wt_art: ws.cell(curr_r, col_wt_art).value = "350"
                if col_wt_pkg: ws.cell(curr_r, col_wt_pkg).value = "380"
                if col_wt_uom: ws.cell(curr_r, col_wt_uom).value = "GRAM"

                # Images Sequence
                for img_idx, img_link in enumerate(sku_img_urls):
                    if img_idx < len(img_cols):
                        ws.cell(curr_r, img_cols[img_idx]).value = img_link

                # MRP logic
                if col_mrp:
                    mrp_cell = ws.cell(curr_r, col_mrp)
                    if mrp_value:
                        mrp_cell.value = mrp_value
                    else:
                        mrp_cell.value = ""
                        mrp_cell.fill = mrp_fill
                        mrp_cell.font = mrp_font

                curr_r += 1
                total_rows_created += 1

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success(f"🎉 Zabardast! Total {len(sku_list)} SKUs ke {total_rows_created} variants successfully generate ho gaye!")
        st.download_button(
            label="📥 Download Filled Listing File",
            data=output,
            file_name=f"{marketplace}_MULTI_SKU_LISTING.{'xlsm' if is_xlsm else 'xlsx'}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error aaya: {e}")
