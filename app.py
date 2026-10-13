import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import os
import re
import zipfile

st.set_page_config(page_title="Multi-Marketplace Auto Listing Generator", layout="wide")

st.title("📦 Multi-Marketplace Auto Listing Generator (Pro V3)")
st.caption("ZIP Folder Upload, Multiple SKUs, Dropbox/Direct Image Links & Strict Marketplace Guidelines")

# 1. Marketplace & Category Selection
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

# 2. Input Mode: ZIP Upload vs Manual / Dropbox
st.subheader("1. Product Images & SKUs Input Mode")
input_mode = st.radio(
    "Kaise Data Dena Chahte Hain?",
    ["📁 Option A: Upload Product Folders ZIP (Recommended - Fast & Offline)", "🔗 Option B: Enter SKUs List & Dropbox Link Manually"],
    horizontal=True
)

detected_skus = []
sku_image_map = {}

col_left, col_right = st.columns(2)

if "Option A" in input_mode:
    with col_left:
        st.write("##### Step A: Images Ka ZIP Upload Karein")
        zip_file = st.file_uploader(
            "Product Folders Ka ZIP Upload Karein (.zip):",
            type=["zip"],
            help="ZIP me har SKU ka folder hona chahiye (e.g. JEANS-U25, JEANS-U26) aur andar 1.jpg, 2.jpg..."
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

                    detected_skus = sorted(list(folders))
                    
                    for sk in detected_skus:
                        imgs = [
                            p for p in file_list 
                            if p.replace('\\', '/').startswith(sk + '/') 
                            and p.lower().endswith(('.jpg', '.jpeg', '.png', '.webp'))
                        ]
                        sku_image_map[sk] = len(imgs) if len(imgs) > 0 else 5

                st.success(f"✅ ZIP Se Total {len(detected_skus)} Product Folders (SKUs) Mil Gaye:")
                st.code(", ".join(detected_skus))
            except Exception as e:
                st.error(f"ZIP read karne me error: {e}")

        dropbox_base_url = st.text_input(
            "Dropbox Shared Folder Link (Image Links Ke Liye):",
            placeholder="https://www.dropbox.com/scl/fo/xyz/...",
            help="Agar images Dropbox par hain toh link daalein taaki Excel me direct ?raw=1 links ban sakein."
        )

    with col_right:
        st.write("##### Step B: Sizes & Attributes")
        sizes_input = st.text_input(
            "Sizes (comma-separated - e.g. 30, 32, 34 ya XS, S, M, L):",
            value="30, 32, 34, 36, 38",
            help="Har SKU ke sath yeh sizes judenge: SKU-30, SKU-32 ya SKU-XS"
        )
        brand_name = st.text_input(
            "Brand Name (Default Blank rakha hai, chahein toh fill karein):",
            value="",
            placeholder="Khali rahega agar kuch nahi likhenge"
        )
        fabric_default = "Denim" if "Jeans" in product_category else "Cotton Blend"
        fabric_input = st.text_input(
            "Fabric (Default Auto-Selected - manual badal sakte hain):",
            value=fabric_default
        )
        mrp_value = st.text_input(
            "MRP (Khali chhodne par Yellow Highlight hoga):",
            value="",
            placeholder="e.g. 1999 ya blank chhod dein"
        )

else:
    with col_left:
        skus_input = st.text_area(
            "Style Codes / SKUs List (Har line me ek SKU):",
            value="SKU-1\nSKU-2\nSKU-4",
            height=120,
            help="Ek se zyada SKUs ke liye Enter dabakar nayi line me likhein."
        )
        detected_skus = [s.strip() for s in skus_input.split('\n') if s.strip()]
        dropbox_base_url = st.text_input(
            "Dropbox Main Shared Folder Link:",
            placeholder="https://www.dropbox.com/scl/fo/xyz/...",
            help="Dropbox main folder link jisme har SKU ka subfolder ho."
        )
        images_count = st.number_input(
            "Har SKU ke folder me kitni images hain? (1, 2, 3...):",
            min_value=1,
            max_value=12,
            value=5
        )
        for sk in detected_skus:
            sku_image_map[sk] = images_count

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
        fabric_default = "Denim" if "Jeans" in product_category else "Cotton Blend"
        fabric_input = st.text_input(
            "Fabric:",
            value=fabric_default
        )
        mrp_value = st.text_input(
            "MRP (Khali chhodne par Yellow Highlight hoga):",
            value="",
            placeholder="e.g. 1999 ya blank chhod dein"
        )

st.divider()

# 3. Trending Keywords Section
st.subheader("2. SEO & Trending Keywords")
custom_keywords = st.text_area(
    "Custom Trending Keywords (Optional):",
    placeholder="e.g. baggy jeans, cargo utility pockets, comfort stretch, streetwear, clean washed denim",
    help="Agar aapke paas trending search terms hain toh yahan daalein. System inko Title, Description aur Highlights me include kar dega."
)

st.divider()

# 4. Template Uploader
st.subheader("3. Template File & Generate")
uploaded_file = st.file_uploader(
    "Apni Excel Template (.xlsx ya .xlsm) Upload Karein:",
    type=["xlsx", "xlsm"]
)

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
    if not detected_skus:
        st.error("Kripya kam se kam ek SKU enter karein ya ZIP file upload karein!")
    else:
        try:
            is_xlsm = uploaded_file.name.lower().endswith(".xlsm")
            wb = openpyxl.load_workbook(uploaded_file, keep_vba=is_xlsm)

            for s in wb.sheetnames:
                wb[s].protection.sheet = False
                wb[s].protection.disable()

            target_sheet = None
            for s in wb.sheetnames:
                s_low = s.lower()
                if any(k in s_low for k in ['styles', 'jeans', 'catalog', 'listing', 'template', 'data']) and not any(ign in s_low for ign in ['instruction', 'guideline', 'extra', 'rejection', 'seller info']):
                    target_sheet = s
                    break
            if not target_sheet:
                target_sheet = wb.sheetnames[0]

            ws = wb[target_sheet]

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
            col_highlight = match_c(['highlight', 'bullet point 1', 'key features'])
            col_add_info = match_c(['additional information 1', 'bullet point 2'])
            col_waist = match_c(['waist'])
            col_to_fit = match_c(['to fit waist'])
            col_inseam = match_c(['inseam'])
            col_country = match_c(['country of origin'])

            img_cols = []
            for i in range(1, 15):
                c_img = match_c([f'*model' if i==1 else f'model{i}', f'image {i}', f'image_{i}', f'main image' if i==1 else f'other image {i-1}', f'front image' if i==1 else f'back image'])
                if c_img and c_img not in img_cols:
                    img_cols.append(c_img)

            mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
            mrp_font = Font(bold=True)

            curr_r = header_row + 1
            while ws.cell(curr_r, col_style if col_style else 1).value is not None:
                curr_r += 1

            size_list = [sz.strip() for sz in sizes_input.split(",") if sz.strip()]
            total_rows_created = 0

            for sku_item in detected_skus:
                detected_color = detect_color_from_sku(sku_item)
                desc_text, high_text, add_text = build_seo_content(
                    sku_item, detected_color, fabric_input, marketplace, custom_keywords
                )

                img_cnt = sku_image_map.get(sku_item, 5)
                sku_img_urls = [
                    get_clean_dropbox_url(dropbox_base_url, sku_item, i) 
                    for i in range(1, img_cnt + 1)
                ]

                for sz in size_list:
                    full_sku = f"{sku_item}-{sz}"
                    
                    try:
                        num_sz = int(sz)
                        w_val = str(num_sz + 1)
                        tf_val = str(num_sz - 1)
                    except:
                        w_val = "33"
                        tf_val = "31"

                    if col_style: ws.cell(curr_r, col_style).value = sku_item
                    if col_sku: ws.cell(curr_r, col_sku).value = full_sku
                    if col_ean: ws.cell(curr_r, col_ean).value = full_sku
                    if col_size: ws.cell(curr_r, col_size).value = sz
                    if col_std_size: ws.cell(curr_r, col_std_size).value = sz
                    if col_brand: ws.cell(curr_r, col_brand).value = brand_name
                    if col_desc: ws.cell(curr_r, col_desc).value = desc_text
                    if col_hsn: ws.cell(curr_r, col_hsn).value = "62034290"
                    if col_color: ws.cell(curr_r, col_color).value = detected_color
                    if col_fabric: ws.cell(curr_r, col_fabric).value = fabric_input
                    if col_country: ws.cell(curr_r, col_country).value = "India"

                    if col_highlight and high_text: ws.cell(curr_r, col_highlight).value = high_text
                    if col_add_info and add_text: ws.cell(curr_r, col_add_info).value = add_text
                    if col_waist: ws.cell(curr_r, col_waist).value = w_val
                    if col_to_fit: ws.cell(curr_r, col_to_fit).value = tf_val
                    if col_inseam: ws.cell(curr_r, col_inseam).value = "32"

                    if col_wt_art: ws.cell(curr_r, col_wt_art).value = "350"
                    if col_wt_pkg: ws.cell(curr_r, col_wt_pkg).value = "380"
                    if col_wt_uom: ws.cell(curr_r, col_wt_uom).value = "GRAM"

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
                    total_rows_created += 1

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success(f"🎉 Zabardast! Total {len(detected_skus)} SKUs ke {total_rows_created} variants successfully generate ho gaye!")
            st.download_button(
                label="📥 Download Filled Listing File",
                data=output,
                file_name=f"{marketplace}_MULTI_SKU_LISTING.{'xlsm' if is_xlsm else 'xlsx'}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error aaya: {e}")
