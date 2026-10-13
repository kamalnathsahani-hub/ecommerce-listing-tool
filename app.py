import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import re

st.set_page_config(page_title="Multi-Marketplace Auto Listing Tool", layout="wide")

st.title("📦 Multi-Marketplace Auto Listing Generator")
st.write("Apna marketplace chunein, template upload karein aur 1-click me listing file ready karein.")

# 1. Marketplace Selection
marketplace = st.selectbox(
    "Marketplace Chunein:",
    ["Ajio", "Meesho", "Flipkart", "Amazon", "Myntra"]
)

# 2. Input Parameters
col1, col2 = st.columns(2)
with col1:
    style_code = st.text_input("Style Code / Parent SKU:", value="JEANS-U25")
    brand_name = st.text_input("Brand Name:", value="Buda Jeans Co")
    sizes_input = st.text_input("Sizes (comma-separated):", value="30, 32, 34, 36, 38")
    mrp_value = st.text_input("MRP (Optional - Khali chhodne par Yellow Highlight hoga):", value="")

with col2:
    color_val = st.text_input("Color:", value="Blue")
    fabric_val = st.text_input("Fabric:", value="Denim")
    image_links_input = st.text_area("Image Links (Har line me ek link - 1.jpg, 2.jpg...):", height=130)

# 3. File Uploader
uploaded_file = st.file_uploader("Apni Excel Template Upload Karein (.xlsx ya .xlsm)", type=["xlsx", "xlsm"])

def generate_seo_desc(style, m_place):
    if m_place == "Meesho":
        return f"Trendy Men's {style} Denim Jeans. Fabric: Premium Cotton Blend Stretchable Denim. Fit: Relaxed Mid-Rise. Pattern: Clean Washed."
    elif m_place == "Flipkart":
        return f"Upgrade your style statement with Men's {style} Jeans. Crafted with premium stretchable breathable denim, mid rise waist, clean washed style."
    elif m_place == "Amazon":
        return f"Men's Casual Relaxed Fit Stretchable Denim Jeans - {style}. Designed with premium washed denim fabric, 5-pocket utility design."
    else:
        return (
            f"Upgrade your everyday denim rotation with this trendy Men's {style} Jeans crafted from premium, "
            "high-grade stretchable denim fabric designed for maximum comfort and all-day ease. "
            "Featuring a contemporary relaxed mid-rise fit that sits comfortably at the natural waist, this pair blends "
            "urban streetwear aesthetic with classic tailoring. Engineered with authentic washed tones, clean subtle fading, "
            "and reinforced contrast stitching, it ensures unmatched durability."
        )

# 4. Generate Button
if uploaded_file and st.button("🚀 Generate Listing File"):
    try:
        is_xlsm = uploaded_file.name.lower().endswith(".xlsm")
        wb = openpyxl.load_workbook(uploaded_file, keep_vba=is_xlsm)
        
        # Unprotect sheets
        for s in wb.sheetnames:
            wb[s].protection.sheet = False
            wb[s].protection.disable()
            
        # Target Sheet identification
        target_sheet = None
        for s in wb.sheetnames:
            if any(k in s.lower() for k in ['styles', 'jeans', 'catalog', 'listing', 'template']):
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = wb.sheetnames[0]
            
        ws = wb[target_sheet]
        
        # Find Header Row
        header_row = 1
        for r in range(1, min(6, ws.max_row + 1)):
            vals = [str(ws.cell(r, c).value or '').lower() for c in range(1, ws.max_column + 1)]
            if any(k in " ".join(vals) for k in ['style', 'sku', 'mrp', 'size']):
                header_row = r
                break
                
        headers = {c: str(ws.cell(header_row, c).value or '').strip() for c in range(1, ws.max_column + 1)}
        
        def match_c(keys):
            for c, h in headers.items():
                cln = h.lower().replace('*', '').replace('_', ' ').replace('-', ' ').strip()
                for k in keys:
                    if k.lower() in cln:
                        return c
            return None

        col_style = match_c(['style code', 'parent sku'])
        col_sku = match_c(['item sku', 'sku'])
        col_mrp = match_c(['mrp', 'maximum retail price', 'standard price'])
        col_size = match_c(['size', 'standardsize'])
        col_brand = match_c(['brand'])
        col_desc = match_c(['description', 'title', 'product name'])
        col_wt_art = match_c(['articledimensionsunitweight', 'weight'])
        col_wt_pkg = match_c(['packagedimensionsweight'])
        col_wt_uom = match_c(['weightuom', 'weight unit'])
        
        img_cols = []
        for i in range(1, 10):
            c_img = match_c([f'model{i if i>1 else ""}', f'image {i}', f'image_{i}', f'main image' if i==1 else f'other image {i-1}'])
            if c_img and c_img not in img_cols:
                img_cols.append(c_img)

        sizes = [s.strip() for s in sizes_input.split(",") if s.strip()]
        img_urls = [u.strip() for u in image_links_input.split("\n") if u.strip()]
        seo_text = generate_seo_desc(style_code, marketplace)
        
        mrp_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
        mrp_font = Font(bold=True)

        curr_r = header_row + 1
        while ws.cell(curr_r, col_style if col_style else 1).value is not None:
            curr_r += 1

        for sz in sizes:
            sku = f"{style_code}-{sz}"
            if col_style: ws.cell(curr_r, col_style).value = style_code
            if col_sku: ws.cell(curr_r, col_sku).value = sku
            if col_size: ws.cell(curr_r, col_size).value = sz
            if col_brand: ws.cell(curr_r, col_brand).value = brand_name
            if col_desc: ws.cell(curr_r, col_desc).value = seo_text
            
            # Weight rule 300g-400g
            if col_wt_art: ws.cell(curr_r, col_wt_art).value = "350"
            if col_wt_pkg: ws.cell(curr_r, col_wt_pkg).value = "380"
            if col_wt_uom: ws.cell(curr_r, col_wt_uom).value = "GRAM"

            # Images mapping
            for idx, url in enumerate(img_urls):
                if idx < len(img_cols):
                    ws.cell(curr_r, img_cols[idx]).value = url

            # MRP highlight ya value
            if col_mrp:
                cell_mrp = ws.cell(curr_r, col_mrp)
                if mrp_value:
                    cell_mrp.value = mrp_value
                else:
                    cell_mrp.fill = mrp_fill
                    cell_mrp.font = mrp_font

            curr_r += 1

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success(f"File successfully create ho gayi ({len(sizes)} rows add hui)!")
        st.download_button(
            label="📥 Download Filled Listing File",
            data=output,
            file_name=f"{marketplace}_{style_code}_LISTED.{'xlsm' if is_xlsm else 'xlsx'}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error aaya: {e}")