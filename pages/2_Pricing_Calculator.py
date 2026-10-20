import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 1. Page Config hamesha sabse upar hona chahiye
st.set_page_config(page_title="Pricing & TP Calculator", page_icon="💰", layout="wide")

# ==========================================
# 🔒 SECURITY CHECK START
# ==========================================
SECRET_PASSWORD = "MYSECRET123"

if not st.session_state.get("pricing_access", False):
    st.warning("🔒 Ye tool private hai. Isey use karne ke liye Access Code dalein.")
    
    pwd = st.text_input("Enter Access Code:", type="password")
    col_btn1, col_btn2 = st.columns([1, 5])
    with col_btn1:
        if st.button("Unlock Tool 🔓", use_container_width=True):
            if pwd == SECRET_PASSWORD:
                st.session_state.pricing_access = True
                st.rerun()
            else:
                st.error("❌ Access Denied! Sahi code enter karein.")
    with col_btn2:
        if st.button("⬅️ Back to Home"):
            st.switch_page("app.py")
            
    st.stop()
# ==========================================
# 🔒 SECURITY CHECK END
# ==========================================

# Top Navigation & Logout Button
top_col1, top_col2 = st.columns([8, 1])
with top_col1:
    if st.button("⬅️ Back to Home"):
        st.switch_page("app.py")
with top_col2:
    if st.button("🔒 Lock Tool"):
        st.session_state.pricing_access = False
        st.rerun()

st.title("💰 Multi-Brand Pricing & Amazon TP Calculator")
st.caption("Auto-calculate ASP for ~10% Goal Margin | Live Formula Calculator Sheet + Vendor Cost & DF Attributes Master")

# ==========================================
# 1. BRAND PRESET CONFIGURATIONS
# ==========================================
BRAND_CONFIGS = {
    "FABNEX": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 26.0,
        "ams_pct": 4.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "FP9UO", "sc_code": "FOHK"
    },
    "QNIM": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 30.0,
        "ams_pct": 4.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "KR8S7", "sc_code": "GCYB"
    },
    "WEREFLEX": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 30.0,
        "ams_pct": 5.0, "office_exp": 45.0, "consult_pct": 5.0,
        "deduct_consult": False, "vendor_code": "YHUBT", "sc_code": "GGHB"
    },
    "ZURITY": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 31.0,
        "ams_pct": 7.0, "office_exp": 50.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "PC9NG", "sc_code": "GZEZ"
    },
    "URBAN": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 31.0,
        "ams_pct": 6.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "URHY6", "sc_code": "GPGJ"
    },
    "LITZO": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 31.0,
        "ams_pct": 6.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "", "sc_code": "LITZO"
    },
    "ASMERIYO": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 31.0,
        "ams_pct": 6.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "RCCLU", "sc_code": "GSTK"
    },
    "LERIYA": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 31.0,
        "ams_pct": 6.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "9S98W", "sc_code": "HGQE"
    },
    "SS GLOBAL": {
        "tax_pct": 5.0, "margin_type": "fixed", "margin_pct": 31.0,
        "ams_pct": 6.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "SSGNJ", "sc_code": "HCLK"
    },
    "LYMIO": {
        "tax_pct": 5.0, "margin_type": "slab_3",
        "slab1_max": 499, "slab1_margin": 32.0,
        "slab2_min": 499, "slab2_max": 749, "slab2_margin": 31.0,
        "slab3_min": 749, "slab3_margin": 28.0,
        "ams_pct": 7.0, "office_exp": 40.0, "consult_pct": 5.0,
        "deduct_consult": True, "vendor_code": "JBZ54", "sc_code": "GFUK"
    },
    "LYMIO PREMIUM": {
        "tax_pct": 5.0, "margin_type": "slab_2",
        "slab1_max": 749, "slab1_margin": 31.0,
        "slab2_min": 749, "slab2_margin": 28.0,
        "ams_pct": 8.0, "office_exp": 60.0, "consult_pct": 5.0,
        "deduct_consult": False, "vendor_code": "JBZ54", "sc_code": "GFUK"
    }
}

# ==========================================
# 2. TOP PANEL: BRAND & RATIO SELECTION
# ==========================================
st.subheader("1️⃣ Step 1: Select Brand & Customize Ratios")

col_b1, col_b2, col_b3 = st.columns([1.5, 1.5, 1])
with col_b1:
    selected_brand = st.selectbox(
        "🏷️ Brand Name Select Karein:",
        options=list(BRAND_CONFIGS.keys()),
        index=3  # Default ZURITY
    )

cfg = BRAND_CONFIGS[selected_brand]
target_vendor_code = cfg.get("vendor_code", "")
target_sc_vendor_str = f"AmazonIn/{cfg.get('sc_code', '')}" if cfg.get("sc_code") else "AmazonIn/"

with col_b2:
    st.info(f"🏢 **Codes Auto-Selected:**\n- `vendor_code`: **{target_vendor_code or 'N/A'}**\n- `sc_vendor_name`: **{target_sc_vendor_str}**")

with col_b3:
    target_goal_pct = st.number_input("🎯 Target Goal Margin %", value=10.0, step=0.5, format="%.2f")

st.markdown("##### ⚙️ Brand Parameters:")
p_col1, p_col2, p_col3, p_col4, p_col5, p_col6 = st.columns(6)

with p_col1:
    tax_pct = st.number_input("Tax %", value=float(cfg["tax_pct"]), step=0.5, format="%.2f") / 100.0

with p_col2:
    ams_pct = st.number_input("AMS %", value=float(cfg["ams_pct"]), step=0.5, format="%.2f") / 100.0

with p_col3:
    office_exp = st.number_input("Office Expense (₹)", value=float(cfg["office_exp"]), step=5.0, format="%.2f")

with p_col4:
    consult_pct = st.number_input("Consulting Charge %", value=float(cfg["consult_pct"]), step=0.5, format="%.2f") / 100.0

with p_col5:
    deduct_consult = st.checkbox(
        "Deduct Consulting in Final Margin?",
        value=cfg["deduct_consult"],
        help="Agar tick hai to Margin % = (Final Profit - Consulting Charge) / Amazon TP hoga."
    )

# Margin % UI (Fixed vs Slab)
slab_params = {}
with p_col6:
    if cfg["margin_type"] == "fixed":
        fixed_margin_pct = st.number_input("Brand Margin %", value=float(cfg["margin_pct"]), step=0.5, format="%.2f") / 100.0
    else:
        st.info("📊 Slab-based Margin Active")
        fixed_margin_pct = 0.30

if cfg["margin_type"] == "slab_3":
    s1, s2, s3 = st.columns(3)
    with s1:
        slab1_m = st.number_input("Margin % (ASP < 499)", value=float(cfg["slab1_margin"]), step=0.5) / 100.0
    with s2:
        slab2_m = st.number_input("Margin % (ASP 499 to 748)", value=float(cfg["slab2_margin"]), step=0.5) / 100.0
    with s3:
        slab3_m = st.number_input("Margin % (ASP >= 749)", value=float(cfg["slab3_margin"]), step=0.5) / 100.0
    slab_params = {"type": "slab_3", "s1": slab1_m, "s2": slab2_m, "s3": slab3_m}
elif cfg["margin_type"] == "slab_2":
    s1, s2 = st.columns(2)
    with s1:
        slab1_m = st.number_input("Margin % (ASP < 749)", value=float(cfg["slab1_margin"]), step=0.5) / 100.0
    with s2:
        slab2_m = st.number_input("Margin % (ASP >= 749)", value=float(cfg["slab2_margin"]), step=0.5) / 100.0
    slab_params = {"type": "slab_2", "s1": slab1_m, "s2": slab2_m}
else:
    slab_params = {"type": "fixed", "margin": fixed_margin_pct}

# Allowed ASP Endings Info
if selected_brand == "WEREFLEX":
    allowed_endings = [19, 29, 39, 49, 59, 69, 79, 89, 99]
    st.caption("✅ **WEREFLEX Rule Active:** ASP endings allowed: `19, 29, 39, 49, 59, 69, 79, 89, 99`")
else:
    allowed_endings = [49, 59, 69, 79, 89, 99]
    st.caption(f"✅ **{selected_brand} Rule Active:** `09, 19, 29, 39` allowed nahi hain. Sirf `49, 59, 69, 79, 89, 99` endings use hongi.")

st.divider()

# ==========================================
# 3. CORE CALCULATION & ASP SOLVER FUNCTIONS
# ==========================================
def get_margin_pct_for_asp(asp, slab_info):
    if slab_info["type"] == "fixed":
        return slab_info["margin"]
    elif slab_info["type"] == "slab_3":
        if asp >= 749:
            return slab_info["s3"]
        elif asp >= 499:
            return slab_info["s2"]
        else:
            return slab_info["s1"]
    elif slab_info["type"] == "slab_2":
        if asp >= 749:
            return slab_info["s2"]
        else:
            return slab_info["s1"]
    return 0.30

def calculate_row_metrics(asp, cost, tax_r, slab_info, ams_r, off_exp, cons_r, deduct_c):
    asp_wo_tax = asp / (1.0 + tax_r)
    m_pct = get_margin_pct_for_asp(asp, slab_info)
    margin_val = asp_wo_tax * m_pct
    ams_val = asp_wo_tax * ams_r
    tp_with_ams = asp_wo_tax - margin_val - ams_val
    final_tp = round(tp_with_ams + ams_val, 2)
    profit_gross = round(final_tp - cost, 2)
    profit_net = round(tp_with_ams - cost, 2)
    final_profit = round(profit_net - off_exp, 2)
    consulting_charge = round(final_tp * cons_r, 2)
    profit_after_consult = round(final_profit - consulting_charge, 2)
    
    if deduct_c:
        margin_pct_final = round((profit_after_consult / final_tp) * 100.0, 2) if final_tp > 0 else 0.0
    else:
        margin_pct_final = round((final_profit / final_tp) * 100.0, 2) if final_tp > 0 else 0.0
        
    status = "OK" if margin_pct_final >= 10.0 else "LOW"
    return {
        "ASP": asp,
        "Product Cost": cost,
        "ASP without Tax": round(asp_wo_tax, 2),
        "Margin": round(margin_val, 2),
        "AMS": round(ams_val, 2),
        "Tp With AMS": round(tp_with_ams, 2),
        "Final Cost (Amazon TP)": final_tp,
        "Profit Gross": profit_gross,
        "Profit Net": profit_net,
        "Office Expense": off_exp,
        "Final Profit": final_profit,
        "Consulting Charge": consulting_charge,
        "Final Profit after Consulting": profit_after_consult,
        "Margin %": margin_pct_final,
        "Status": status,
        "Margin % Applied": round(m_pct * 100, 2)
    }

def find_best_asp(cost, tax_r, slab_info, ams_r, off_exp, cons_r, deduct_c, endings, goal_pct=10.0):
    candidates = []
    for base in range(100, 9900, 100):
        for end in endings:
            cand_asp = base + end
            res = calculate_row_metrics(cand_asp, cost, tax_r, slab_info, ams_r, off_exp, cons_r, deduct_c)
            m_pct = res["Margin %"]
            diff = abs(m_pct - goal_pct)
            candidates.append((diff, m_pct >= goal_pct, cand_asp, res))
    
    candidates.sort(key=lambda x: x[0])
    return candidates[0][3]

# ==========================================
# 4. EXCEL WORKBOOK GENERATORS
# ==========================================
def generate_calculator_workbook(df_results, brand_name, tax_r, slab_info, ams_r, off_exp, cons_r, deduct_c):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = brand_name[:31]

    title_font = Font(name="Calibri", size=13, bold=True, color="1F4E79")
    note_font = Font(name="Calibri", size=10, italic=True, color="595959")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    param_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    bold_font = Font(name="Calibri", size=10, bold=True)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws["A1"] = f"{brand_name} - Brand TP Calculator - Bulk (column format)"
    ws["A1"].font = title_font
    ws["A2"] = "Sirf peeli (yellow) cells me ASP aur Product Cost daalo. Baaki sab automatic calculate hoga. Har row = 1 article."
    ws["A2"].font = note_font

    ws["A4"] = "Tax %"
    ws["B4"] = tax_r
    ws["B4"].number_format = "0.0%"
    ws["C4"] = "Fixed values - change karna ho to yahin B4:B8 me karo"

    ws["A5"] = "Margin %"
    if slab_info["type"] == "fixed":
        ws["B5"] = slab_info["margin"]
        ws["B5"].number_format = "0.0%"
    else:
        ws["B5"] = "Auto"

    ws["A6"] = "AMS %"
    ws["B6"] = ams_r
    ws["B6"].number_format = "0.0%"

    ws["A7"] = "Office Expense"
    ws["B7"] = off_exp
    ws["B7"].number_format = "0.00"

    ws["A8"] = "Consulting Charge %"
    ws["B8"] = cons_r
    ws["B8"].number_format = "0.0%"

    for r in range(4, 9):
        ws[f"A{r}"].font = bold_font
        ws[f"B{r}"].fill = param_fill
        ws[f"B{r}"].font = bold_font

    if slab_info["type"] == "slab_3":
        ws["E4"], ws["F4"], ws["G4"] = "ASP slab", "ASP from", "Margin %"
        ws["E5"], ws["F5"], ws["G5"] = "Below 499", 0, slab_info["s1"]
        ws["E6"], ws["F6"], ws["G6"] = "499 to 748", 499, slab_info["s2"]
        ws["E7"], ws["F7"], ws["G7"] = "749 & above", 749, slab_info["s3"]
        for r in range(5, 8):
            ws[f"G{r}"].number_format = "0.0%"
    elif slab_info["type"] == "slab_2":
        ws["E4"], ws["F4"], ws["G4"] = "ASP slab", "ASP from", "Margin %"
        ws["E5"], ws["F5"], ws["G5"] = "Below 749", 0, slab_info["s1"]
        ws["E6"], ws["F6"], ws["G6"] = "749 & above", 749, slab_info["s2"]
        for r in range(5, 7):
            ws[f"G{r}"].number_format = "0.0%"

    ws["A9"] = "Note: Final Cost (Amazon TP) 2 decimal places tak calculate hota hai. Formulas 500 rows tak bhare hain."
    ws["A9"].font = note_font

    headers = [
        "Article", "ASP (INPUT)", "Product Cost (INPUT)", "ASP without Tax",
        "Margin", "AMS", "Tp With AMS", "Final Cost (Amazon TP)",
        "Profit Gross", "Profit Net", "Office Expense", "Final Profit",
        "Consulting Charge", "Final Profit after Consulting",
        "Margin % (Final Profit after Consulting / Amazon TP)", "Margin >= 10%?"
    ]
    if slab_info["type"] != "fixed":
        headers.append("Margin % Applied")

    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(11, col_idx, h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    total_rows = max(500, len(df_results) + 50)
    for idx in range(total_rows):
        r = 12 + idx
        if idx < len(df_results):
            row_data = df_results.iloc[idx]
            ws.cell(r, 1, str(row_data["ASIN"]))
            ws.cell(r, 2, float(row_data["ASP (INPUT)"]))
            ws.cell(r, 3, float(row_data["Product Cost (INPUT)"]))

        ws.cell(r, 2).fill = input_fill
        ws.cell(r, 3).fill = input_fill

        ws.cell(r, 4, f'=IF(OR($B{r}="",$C{r}=""),"",$B{r}/(1+$B$4))')
        if slab_info["type"] == "fixed":
            ws.cell(r, 5, f'=IF(OR($B{r}="",$C{r}=""),"",D{r}*$B$5)')
        elif slab_info["type"] == "slab_3":
            ws.cell(r, 5, f'=IF(OR($B{r}="",$C{r}=""),"",D{r}*IF($B{r}>=$F$7,$G$7,IF($B{r}>=$F$6,$G$6,$G$5)))')
        elif slab_info["type"] == "slab_2":
            ws.cell(r, 5, f'=IF(OR($B{r}="",$C{r}=""),"",D{r}*IF($B{r}>=$F$6,$G$6,$G$5))')

        ws.cell(r, 6, f'=IF(OR($B{r}="",$C{r}=""),"",D{r}*$B$6)')
        ws.cell(r, 7, f'=IF(OR($B{r}="",$C{r}=""),"",D{r}-E{r}-F{r})')
        ws.cell(r, 8, f'=IF(OR($B{r}="",$C{r}=""),"",ROUND(G{r}+F{r},2))')
        ws.cell(r, 9, f'=IF(OR($B{r}="",$C{r}=""),"",H{r}-$C{r})')
        ws.cell(r, 10, f'=IF(OR($B{r}="",$C{r}=""),"",G{r}-$C{r})')
        ws.cell(r, 11, f'=IF(OR($B{r}="",$C{r}=""),"",$B$7)')
        ws.cell(r, 12, f'=IF(OR($B{r}="",$C{r}=""),"",J{r}-K{r})')
        ws.cell(r, 13, f'=IF(OR($B{r}="",$C{r}=""),"",ROUND(H{r}*$B$8,2))')
        ws.cell(r, 14, f'=IF(OR($B{r}="",$C{r}=""),"",L{r}-M{r})')
        
        if deduct_consult:
            ws.cell(r, 15, f'=IF(OR($B{r}="",$C{r}=""),"",N{r}/H{r}*100)')
        else:
            ws.cell(r, 15, f'=IF(OR($B{r}="",$C{r}=""),"",L{r}/H{r}*100)')
            
        ws.cell(r, 16, f'=IF(OR($B{r}="",$C{r}=""),"",IF(O{r}>=10,"OK","LOW"))')

        if slab_info["type"] == "slab_3":
            ws.cell(r, 17, f'=IF(OR($B{r}="",$C{r}=""),"",IF($B{r}>=$F$7,$G$7,IF($B{r}>=$F$6,$G$6,$G$5)))')
            ws.cell(r, 17).number_format = "0.0%"
        elif slab_info["type"] == "slab_2":
            ws.cell(r, 17, f'=IF(OR($B{r}="",$C{r}=""),"",IF($B{r}>=$F$6,$G$6,$G$5))')
            ws.cell(r, 17).number_format = "0.0%"

        for c_i in [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]:
            ws.cell(r, c_i).number_format = "0.00"
            ws.cell(r, c_i).border = thin_border

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.row_dimensions[11].height = 32
    ws.freeze_panes = "A12"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def generate_vendor_master_workbook(df_results, vendor_code_val, sc_vendor_str):
    wb = openpyxl.Workbook()
    
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    grey_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True)

    # ==========================================
    # SHEET 1: Cost
    # ==========================================
    ws_cost = wb.active
    ws_cost.title = "Cost"
    cost_headers = [
        "gl_category", "asin", "vendor_code", "cost_price_currency",
        "cost_price_value_without_tax", "MRP", "DF Cost", "FC Cost", "Selling Price [SP for DF]"
    ]
    for c_idx, h in enumerate(cost_headers, start=1):
        cell = ws_cost.cell(1, c_idx, h)
        cell.fill = yellow_fill
        cell.font = header_font

    for r_idx, row in df_results.iterrows():
        r_num = r_idx + 2
        tp_val = round(float(row["Final Cost (Amazon TP)"]), 2)
        asp_val = float(row["ASP (INPUT)"])
        mrp_val = float(row["MRP"])

        c1 = ws_cost.cell(r_num, 1, "Apparel")
        c2 = ws_cost.cell(r_num, 2, str(row["ASIN"]))
        c1.fill = grey_fill
        c2.fill = grey_fill

        # vendor_code column (Column 3)
        ws_cost.cell(r_num, 3, vendor_code_val if vendor_code_val else None)
        ws_cost.cell(r_num, 4, "INR")
        
        c5 = ws_cost.cell(r_num, 5, tp_val)
        c6 = ws_cost.cell(r_num, 6, mrp_val)
        c7 = ws_cost.cell(r_num, 7, tp_val)
        c8 = ws_cost.cell(r_num, 8, tp_val)
        c9 = ws_cost.cell(r_num, 9, asp_val)
        
        c5.number_format = "0.00"
        c7.number_format = "0.00"
        c8.number_format = "0.00"

    for col in range(1, 10):
        ws_cost.column_dimensions[get_column_letter(col)].width = 22

    # ==========================================
    # SHEET 2: DF Attributes
    # ==========================================
    ws_df = wb.create_sheet(title="DF Attributes")
    ws_df["A1"] = "version=1.0.0"

    df_headers = [
        "asin", "sc_vendor_name", "rtip_dropship_status.value",
        "rtip_is_shipped_from_vendor.value", "replenishment_category.value",
        "deprecated_can_be_giftwrapped.value", "deprecated_can_be_giftmessaged.value",
        "item_dimensions.height.unit", "item_dimensions.height.value",
        "item_dimensions.length.unit", "item_dimensions.length.value",
        "item_dimensions.width.unit", "item_dimensions.width.value",
        "item_weight.unit", "item_weight.value",
        "item_package_weight.unit", "item_package_weight.value",
        "item_package_dimensions.height.unit", "item_package_dimensions.height.value",
        "item_package_dimensions.length.unit", "item_package_dimensions.length.value",
        "item_package_dimensions.width.unit", "item_package_dimensions.width.value",
        "list_price.value_with_tax", "list_price.currency", "vendor_sku.value"
    ]
    for c_idx, h in enumerate(df_headers, start=1):
        cell = ws_df.cell(2, c_idx, h)
        cell.fill = yellow_fill
        cell.font = header_font
    ws_df.cell(2, 29, "*All fields shoud be updated by Vendor")

    for r_idx, row in df_results.iterrows():
        r_num = r_idx + 3
        c_asin = ws_df.cell(r_num, 1, str(row["ASIN"]))
        c_asin.fill = grey_fill
        
        # sc_vendor_name column (Column 2) - format AmazonIn/{code}
        ws_df.cell(r_num, 2, sc_vendor_str)
        ws_df.cell(r_num, 3, "enabled_no_auto_pricing")
        ws_df.cell(r_num, 4, True)
        ws_df.cell(r_num, 5, "non_stockupable")
        ws_df.cell(r_num, 6, False)
        ws_df.cell(r_num, 7, False)
        ws_df.cell(r_num, 8, "centimeters")
        ws_df.cell(r_num, 9, 2)
        ws_df.cell(r_num, 10, "centimeters")
        ws_df.cell(r_num, 11, 30)
        ws_df.cell(r_num, 12, "centimeters")
        ws_df.cell(r_num, 13, 25)
        ws_df.cell(r_num, 14, "grams")
        ws_df.cell(r_num, 15, 400)
        ws_df.cell(r_num, 16, "grams")
        ws_df.cell(r_num, 17, 400)
        ws_df.cell(r_num, 18, "centimeters")
        ws_df.cell(r_num, 19, 2)
        ws_df.cell(r_num, 20, "centimeters")
        ws_df.cell(r_num, 21, 30)
        ws_df.cell(r_num, 22, "centimeters")
        ws_df.cell(r_num, 23, 25)
        ws_df.cell(r_num, 24, None)
        ws_df.cell(r_num, 25, "INR")
        ws_df.cell(r_num, 26, str(row["SKU"]))
        if r_idx == 0:
            ws_df.cell(r_num, 29, "* Sample data provided in Grey colour only for reference -PLEASE DELETE before sharing the file")

    for col in range(1, 27):
        ws_df.column_dimensions[get_column_letter(col)].width = 24

    # ==========================================
    # SHEET 3: Deals - to be filled by VM
    # ==========================================
    ws_deals = wb.create_sheet(title="Deals - to be filled by VM")
    ws_deals["A1"] = "Action"
    ws_deals["B1"] = "ASIN"
    ws_deals["C1"] = "Deal Price"
    ws_deals["D1"] = "PAWS"
    ws_deals["F1"] = "*Sheet to be filled by VM"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ==========================================
# 5. WORKSPACE TABS: BULK UPLOAD & RANDOM CHECKER
# ==========================================
tab1, tab2 = st.tabs(["📂 Bulk Excel Processor (Auto ASP + 2 Workbooks)", "🧮 Quick Single/Random Value Checker"])

with tab1:
    st.subheader("2️⃣ Step 2: Upload Input Excel (ASIN, SKU, Product Cost, MRP)")
    
    sample_df = pd.DataFrame({
        "ASIN": ["B0HMH6MRG1", "B0HMH1YKVP", "B0HMH52D6L"],
        "SKU": ["56-S-01-GREEN-2XL", "56-S-01-GREEN-L", "56-S-01-GREEN-M"],
        "Product Cost": [205, 220, 360],
        "MRP": [4999, 4999, 4999]
    })
    sample_buf = io.BytesIO()
    with pd.ExcelWriter(sample_buf, engine="openpyxl") as writer:
        sample_df.to_excel(writer, index=False, sheet_name="Input_Data")
    sample_buf.seek(0)

    c_up1, c_up2 = st.columns([2, 1])
    with c_up1:
        uploaded_file = st.file_uploader(
            "📤 Apni Input Excel File Upload Karein (.xlsx / .csv)",
            type=["xlsx", "xls", "csv"]
        )
    with c_up2:
        st.write("📄 **Sample Input Format:**")
        st.download_button(
            "⬇️ Download Sample Input Excel",
            data=sample_buf,
            file_name="Sample_Pricing_Input.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                raw_df = pd.read_csv(uploaded_file)
            else:
                raw_df = pd.read_excel(uploaded_file)

            st.success(f"✅ File uploaded successfully! Total Rows: {len(raw_df)}")

            cols = list(raw_df.columns)
            def guess_col(keywords, default_idx=0):
                for idx, c in enumerate(cols):
                    for k in keywords:
                        if k in str(c).lower():
                            return idx
                return min(default_idx, len(cols) - 1)

            st.markdown("##### 🔗 Column Mapping Confirm Karein:")
            m_c1, m_c2, m_c3, m_c4 = st.columns(4)
            with m_c1:
                asin_col = st.selectbox("ASIN Column", cols, index=guess_col(["asin", "article"], 0))
            with m_c2:
                sku_col = st.selectbox("SKU Column", cols, index=guess_col(["sku", "vendor_sku"], min(1, len(cols)-1)))
            with m_c3:
                cost_col = st.selectbox("Product Cost Column", cols, index=guess_col(["cost", "price", "tp"], min(2, len(cols)-1)))
            with m_c4:
                mrp_options = ["(Fixed Default 4999)"] + cols
                mrp_guess = guess_col(["mrp", "list_price"], 0)
                mrp_default_idx = (mrp_guess + 1) if any(k in str(cols[mrp_guess]).lower() for k in ["mrp", "list"]) else 0
                mrp_col = st.selectbox("MRP Column", mrp_options, index=mrp_default_idx)

            processed_rows = []
            for idx, r in raw_df.iterrows():
                asin_v = str(r[asin_col]).strip() if pd.notna(r[asin_col]) else ""
                sku_v = str(r[sku_col]).strip() if pd.notna(r[sku_col]) else ""
                if not asin_v and not sku_v:
                    continue
                try:
                    cost_v = float(r[cost_col])
                except (ValueError, TypeError):
                    continue

                if mrp_col == "(Fixed Default 4999)":
                    mrp_v = 4999.0
                else:
                    try:
                        mrp_v = float(r[mrp_col]) if pd.notna(r[mrp_col]) else 4999.0
                    except (ValueError, TypeError):
                        mrp_v = 4999.0

                best_res = find_best_asp(
                    cost=cost_v,
                    tax_r=tax_pct,
                    slab_info=slab_params,
                    ams_r=ams_pct,
                    off_exp=office_exp,
                    cons_r=consult_pct,
                    deduct_c=deduct_consult,
                    endings=allowed_endings,
                    goal_pct=target_goal_pct
                )

                processed_rows.append({
                    "ASIN": asin_v,
                    "SKU": sku_v,
                    "MRP": mrp_v,
                    "ASP (INPUT)": best_res["ASP"],
                    "Product Cost (INPUT)": cost_v,
                    "ASP without Tax": best_res["ASP without Tax"],
                    "Margin": best_res["Margin"],
                    "AMS": best_res["AMS"],
                    "Tp With AMS": best_res["Tp With AMS"],
                    "Final Cost (Amazon TP)": best_res["Final Cost (Amazon TP)"],
                    "Profit Gross": best_res["Profit Gross"],
                    "Profit Net": best_res["Profit Net"],
                    "Office Expense": best_res["Office Expense"],
                    "Final Profit": best_res["Final Profit"],
                    "Consulting Charge": best_res["Consulting Charge"],
                    "Final Profit after Consulting": best_res["Final Profit after Consulting"],
                    "Margin %": best_res["Margin %"],
                    "Margin >= 10%?": best_res["Status"]
                })

            if processed_rows:
                res_df = pd.DataFrame(processed_rows)
                st.subheader("📊 Preview Calculated Pricing & Amazon TP (Editable ASP)")
                st.caption("💡 Tip: Agar aap kisi row ka `ASP (INPUT)` manually change karna chahein, to table me double-click karke badal sakte hain—sari calculations aur download files apne aap update ho jayengi!")

                edited_df = st.data_editor(
                    res_df,
                    column_config={
                        "ASP (INPUT)": st.column_config.NumberColumn("ASP (INPUT) ✏️", min_value=1, max_value=50000, step=1),
                        "Product Cost (INPUT)": st.column_config.NumberColumn("Product Cost (INPUT) ✏️", format="₹%.2f"),
                        "Final Cost (Amazon TP)": st.column_config.NumberColumn("Final Cost (Amazon TP)", format="₹%.2f"),
                        "Consulting Charge": st.column_config.NumberColumn("Consulting Charge", format="₹%.2f"),
                        "Margin %": st.column_config.NumberColumn("Margin %", format="%.2f%%"),
                    },
                    disabled=[
                        "ASIN", "SKU", "ASP without Tax", "Margin", "AMS", "Tp With AMS",
                        "Final Cost (Amazon TP)", "Profit Gross", "Profit Net", "Office Expense",
                        "Final Profit", "Consulting Charge", "Final Profit after Consulting",
                        "Margin %", "Margin >= 10%?"
                    ],
                    use_container_width=True,
                    hide_index=True
                )

                final_rows = []
                for idx, r in edited_df.iterrows():
                    recalc = calculate_row_metrics(
                        asp=float(r["ASP (INPUT)"]),
                        cost=float(r["Product Cost (INPUT)"]),
                        tax_r=tax_pct,
                        slab_info=slab_params,
                        ams_r=ams_pct,
                        off_exp=office_exp,
                        cons_r=consult_pct,
                        deduct_c=deduct_consult
                    )
                    final_rows.append({
                        "ASIN": r["ASIN"],
                        "SKU": r["SKU"],
                        "MRP": r["MRP"],
                        "ASP (INPUT)": recalc["ASP"],
                        "Product Cost (INPUT)": recalc["Product Cost"],
                        "ASP without Tax": recalc["ASP without Tax"],
                        "Margin": recalc["Margin"],
                        "AMS": recalc["AMS"],
                        "Tp With AMS": recalc["Tp With AMS"],
                        "Final Cost (Amazon TP)": recalc["Final Cost (Amazon TP)"],
                        "Profit Gross": recalc["Profit Gross"],
                        "Profit Net": recalc["Profit Net"],
                        "Office Expense": recalc["Office Expense"],
                        "Final Profit": recalc["Final Profit"],
                        "Consulting Charge": recalc["Consulting Charge"],
                        "Final Profit after Consulting": recalc["Final Profit after Consulting"],
                        "Margin %": recalc["Margin %"],
                        "Margin >= 10%?": recalc["Status"]
                    })
                final_df = pd.DataFrame(final_rows)

                st.divider()
                st.subheader("📥 Step 3: Download Both Workbooks")
                dl_col1, dl_col2 = st.columns(2)

                calc_excel_bytes = generate_calculator_workbook(
                    final_df, selected_brand, tax_pct, slab_params, ams_pct, office_exp, consult_pct, deduct_consult
                )
                
                vendor_excel_bytes = generate_vendor_master_workbook(final_df, target_vendor_code, target_sc_vendor_str)

                with dl_col1:
                    st.download_button(
                        label=f"📊 Download {selected_brand} TP Calculator Workbook (With Live Formulas)",
                        data=calc_excel_bytes,
                        file_name=f"{selected_brand}_TP_Calculator_Bulk.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with dl_col2:
                    st.download_button(
                        label=f"📦 Download {selected_brand} Vendor Cost & DF Attributes Template",
                        data=vendor_excel_bytes,
                        file_name=f"{selected_brand}_Vendor_Cost_Attributes_Master.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                st.warning("⚠️ File me valid Product Cost ya ASIN/SKU nahi mila. Kripya column mapping check karein.")

        except Exception as e:
            st.error(f"Error reading file: {e}")

with tab2:
    st.subheader("🧮 Interactive Single / Random Value Calculator")
    st.write("Yahan aap koi bhi Product Cost daal kar auto-suggested ASP aur poori calculation turant check kar sakte hain:")

    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        test_cost = st.number_input("Enter Product Cost (₹)", value=205.0, step=5.0)
    with rc2:
        auto_best = find_best_asp(
            test_cost, tax_pct, slab_params, ams_pct, office_exp, consult_pct, deduct_consult, allowed_endings, target_goal_pct
        )
        test_asp = st.number_input(
            f"ASP (Auto-picked for ~{target_goal_pct}% Margin: ₹{auto_best['ASP']})",
            value=int(auto_best["ASP"]),
            step=10
        )
    with rc3:
        test_res = calculate_row_metrics(
            test_asp, test_cost, tax_pct, slab_params, ams_pct, office_exp, consult_pct, deduct_consult
        )
        st.metric(
            label="Final Cost (Amazon TP) [2 Decimals]",
            value=f"₹{test_res['Final Cost (Amazon TP)']:.2f}",
            delta=f"Margin: {test_res['Margin %']:.2f}% ({test_res['Status']})"
        )

    st.dataframe(pd.DataFrame([test_res]), use_container_width=True, hide_index=True)
