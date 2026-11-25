import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import io
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment

st.set_page_config(page_title="Payment Reco v3.0 — Universal Multi-Format Reco", layout="wide")
st.title("📊 Universal Marketplace Payment Reconciliation")
st.info("✅ Active Version: v3.0 — Supports All 3 Formats Automatically: (1) Order Info CSV/ZIP, (2) Forward Return Payment Reco CSV, (3) Snapdeal Settlement Excel/CSV")

uploaded_files = st.file_uploader(
    "Upload Order / Payment Reco File(s) — CSV, XLSX, XLS, or ZIP (Multiple files allowed)",
    type=["csv", "xlsx", "xls", "zip"],
    accept_multiple_files=True
)

CHANNEL_DISPLAY_MAP = {
    "JBFASHION_MYNTRA": "Myntra (JB Fashion)",
    "FLIPKART_POSHAX": "Flipkart (Poshax)",
    "AJIO_VRATI_FASHION_PX": "Ajio PX (Vrati)",
    "AJIO_LERIYA_FASHION_AT": "Ajio AT (Leriya)"
}

def clean_order_info_mp(ch):
    s = str(ch)
    if "Myntra" in s:
        return "Myntra"
    if "Amazon" in s:
        return "Amazon"
    if "Ajio" in s:
        return "Ajio"
    return s.strip() if s.strip() else "Other"

def get_parent_style(sku):
    return re.sub(
        r"-(28|30|32|34|36|38|40|42|XS|S|M|L|XL|XXL|3XL)$",
        "",
        str(sku),
        flags=re.IGNORECASE
    )

def get_return_action_plan(reason_str, ret_rate_pct):
    r = str(reason_str).lower()
    if any(k in r for k in ["small", "tight"]):
        return "📏 Size Too Small: Size chart me +1 inch chest/waist update karein & bullet points me 'Order 1 Size Up' likhein"
    elif any(k in r for k in ["large", "loose", "big"]):
        return "📏 Size Too Large: Garment measurements verify karein & listing me Relaxed/Loose Fit mention karein"
    elif any(k in r for k in ["different", "wrong", "incorrect"]):
        return "📦 Wrong Product Sent: Warehouse barcode/SKU sticker mapping check karein & pack-scan mandatory karein"
    elif any(k in r for k in ["quality", "fabric", "material", "stitch", "tear", "damage", "defect"]):
        return "🧵 Quality / Defect Issue: Supplier lot ka GSM/stitching QC check karein & damaged returns par SPF claim dalein"
    elif any(k in r for k in ["color", "colour", "shade", "image", "look", "description"]):
        return "📸 Color / Image Mismatch: Studio lighting me true-color photos upload karein & fabric finish clear likhein"
    elif any(k in r for k in ["mind", "longer needed", "better price", "late"]):
        return "🚚 Impulse / Late Delivery: Same-day dispatch (SLA) maintain karein taaki buyer cancellation/return kam ho"
    elif "dispute" in r or "overdue" in r:
        return "⚖️ Payment Disputed/Overdue: Marketplace seller portal par ticket raise karke POD/return proof attach karein"
    else:
        if ret_rate_pct >= 40:
            return "🔍 High Return Audit: Top return comments/remarks check karke listing title, size chart aur packaging improve karein"
        return "✅ Monitor: Return comments/settlement status track karein aur size chart accurate rakhein"

def read_raw_tables(file_obj):
    fname = file_obj.name
    lower = fname.lower()
    tables = []
    if lower.endswith(".zip"):
        with zipfile.ZipFile(file_obj, "r") as z:
            for inner in z.namelist():
                if inner.startswith("__MACOSX") or inner.endswith("/"):
                    continue
                ilower = inner.lower()
                with z.open(inner) as f:
                    if ilower.endswith(".csv"):
                        tables.append((inner, pd.read_csv(f, low_memory=False)))
                    elif ilower.endswith((".xlsx", ".xls")):
                        tables.append((inner, pd.read_excel(f)))
    elif lower.endswith(".csv"):
        tables.append((fname, pd.read_csv(file_obj, low_memory=False)))
    elif lower.endswith((".xlsx", ".xls")):
        tables.append((fname, pd.read_excel(file_obj)))
    return tables

def detect_and_normalize(raw_df, source_name):
    cols_lower = {str(c).strip().lower(): c for c in raw_df.columns}

    # --- 1. NEW FORWARD RETURN PAYMENT RECONCILIATION CSV ---
    if "channel code" in cols_lower and "actual net settlement" in cols_lower:
        df = raw_df.copy()
        ch_col = cols_lower["channel code"]
        df["Marketplace"] = df[ch_col].astype(str).str.strip().str.upper().map(
            lambda x: CHANNEL_DISPLAY_MAP.get(x, x)
        )

        st_col = cols_lower.get("order status")
        rt_col = cols_lower.get("return type")
        fr_col = cols_lower.get("forward reco status")

        order_st = df[st_col].astype(str).str.strip().str.upper() if st_col else pd.Series("", index=df.index)
        ret_type = df[rt_col].astype(str).str.strip().str.upper() if rt_col else pd.Series("NONE", index=df.index)
        fwd_reco = df[fr_col].astype(str).str.strip().str.upper() if fr_col else pd.Series("", index=df.index)

        act_settle = pd.to_numeric(df[cols_lower["actual net settlement"]], errors="coerce").fillna(0.0)
        exp_settle = pd.to_numeric(df[cols_lower["expected net settlement"]], errors="coerce").fillna(0.0) if "expected net settlement" in cols_lower else pd.Series(0.0, index=df.index)
        cost_price = pd.to_numeric(df[cols_lower["cost price"]], errors="coerce").fillna(0.0) if "cost price" in cols_lower else pd.Series(0.0, index=df.index)
        pf_val = pd.to_numeric(df[cols_lower["protection fund"]], errors="coerce").fillna(0.0) if "protection fund" in cols_lower else pd.Series(0.0, index=df.index)
        selling_price = pd.to_numeric(df[cols_lower["net selling price"]], errors="coerce").fillna(0.0) if "net selling price" in cols_lower else pd.Series(0.0, index=df.index)
        sale_qty = pd.to_numeric(df[cols_lower["uc sale quantity"]], errors="coerce").fillna(1.0) if "uc sale quantity" in cols_lower else pd.Series(1.0, index=df.index)
        ret_qty = pd.to_numeric(df[cols_lower["uc return quantity"]], errors="coerce").fillna(0.0) if "uc return quantity" in cols_lower else pd.Series(0.0, index=df.index)

        is_cancelled = (order_st == "CANCELLED") | ((ret_type == "NONE") & (act_settle == 0) & (fwd_reco == "CANCELLED"))
        is_rto = (~is_cancelled) & (ret_type == "RTO")
        is_cir = (~is_cancelled) & (ret_type == "CIR")
        is_in_transit = (~is_cancelled) & (~is_rto) & (~is_cir) & (order_st.isin(["SHIPPED", "DISPATCHED", "READY_TO_SHIP", "COURIER_ALLOCATED", "CREATED"]))

        norm = pd.DataFrame(index=df.index)
        norm["Source_Format"] = "NFR_RECO"
        norm["Marketplace"] = df["Marketplace"]
        norm["Order_ID"] = df[cols_lower.get("channel order code", cols_lower.get("order code"))].astype(str)
        norm["Sub_Order_ID"] = df[cols_lower.get("channel item code", cols_lower.get("order code"))].astype(str)
        norm["Invoice_No"] = df[cols_lower["sale invoice display code"]].fillna("").astype(str) if "sale invoice display code" in cols_lower else ""
        norm["Order_Date"] = df[cols_lower["channel created time"]] if "channel created time" in cols_lower else ""
        norm["Delivery_Date"] = df[cols_lower["channel delivered time"]] if "channel delivered time" in cols_lower else ""
        norm["Return_Date"] = df[cols_lower["return initiated date"]] if "return initiated date" in cols_lower else ""
        norm["SKU_Code"] = df[cols_lower["sku code"]].fillna("UNKNOWN_SKU").astype(str).str.strip() if "sku code" in cols_lower else "UNKNOWN_SKU"
        norm["Product_Name"] = df[cols_lower["sku code"]].fillna("Product").astype(str).str.strip() if "sku code" in cols_lower else "Product"
        norm["Parent_Style"] = norm["SKU_Code"].apply(get_parent_style)
        norm["Qty"] = sale_qty
        norm["Gross_Sale"] = selling_price
        norm["Actual_Settlement"] = act_settle
        norm["Expected_Settlement"] = exp_settle
        norm["Unit_COGS"] = np.where(sale_qty > 0, cost_price / sale_qty, cost_price)
        norm["Effective_COGS"] = cost_price
        norm["Protection_Fund"] = pf_val
        norm["Loss_Initiated_Flag"] = False

        norm["Status_Bucket"] = np.where(
            is_cancelled, "Cancelled",
            np.where(is_rto, "RTO",
            np.where(is_cir, "Customer Return",
            np.where(is_in_transit, "In-Transit / Active", "Delivered (Success)")))
        )
        norm["Raw_Status"] = order_st + " (" + ret_type + ")"
        norm["Return_Charge"] = np.where(
            act_settle < 0, -act_settle,
            np.where(is_cir & (act_settle > 0), np.abs(exp_settle), 0.0)
        )
        norm["RTO_Charge"] = np.where(is_rto, np.abs(act_settle), 0.0)
        norm["Good_Return_Qty"] = np.where(is_cir | is_rto, ret_qty, 0.0)
        norm["Bad_Return_Qty"] = 0.0
        norm["Bad_Return_Loss"] = 0.0

        rem_col = cols_lower.get("return remarks")
        rev_col = cols_lower.get("reverse reco status")
        rem_s = df[rem_col].fillna("").astype(str).str.strip() if rem_col else pd.Series("", index=df.index)
        rev_s = df[rev_col].fillna("").astype(str).str.strip() if rev_col else pd.Series("", index=df.index)
        norm["Return_Reason"] = np.where(
            rem_s != "", rem_s,
            np.where(rev_s != "", "CIR (" + rev_s + ")", "Customer Return (CIR)")
        )
        norm["Tracking_ID"] = df[cols_lower["order code"]].astype(str) if "order code" in cols_lower else ""

        norm["Is_Unpaid_Delivered"] = (
            (norm["Status_Bucket"] == "Delivered (Success)") &
            ((act_settle == 0) | (fwd_reco.isin(["PAYMENT_DISPUTED", "PAYMENT_OVERDUE", "AWAITING_PAYMENT"])))
        )
        norm["Pending_Payout_Amt"] = np.where(act_settle == 0, exp_settle, np.maximum(exp_settle - act_settle, 0.0))
        norm["Delivered_Profit"] = np.where(
            norm["Status_Bucket"] == "Delivered (Success)",
            np.where(act_settle != 0, act_settle, exp_settle) - cost_price,
            0.0
        )
        return norm

    # --- 2. SNAPDEAL SETTLEMENT REPORT (XLSX / CSV) ---
    elif "actual tp" in cols_lower or ("seller code" in cols_lower and "settled" in cols_lower):
        df = raw_df.copy()
        if "seller code" in cols_lower:
            sc_s = df[cols_lower["seller code"]].astype(str).str.strip().str.upper()
            if (sc_s == "S08163").any():
                df = df[sc_s == "S08163"].copy()

        st_col = cols_lower["order status"]
        df = df[df[st_col].notna() & (df[st_col].astype(str).str.strip() != "")].copy()

        order_st = df[st_col].astype(str).str.strip().str.upper()
        pay_st = df[cols_lower["payment status"]].astype(str).str.strip().str.upper() if "payment status" in cols_lower else pd.Series("", index=df.index)

        act_col = cols_lower.get("actual tp", cols_lower.get("settled"))
        exp_col = cols_lower.get("our tp", cols_lower.get("unsettled"))
        act_val = pd.to_numeric(df[act_col], errors="coerce").fillna(0.0) if act_col else pd.Series(0.0, index=df.index)
        exp_val = pd.to_numeric(df[exp_col], errors="coerce").fillna(0.0) if exp_col else pd.Series(0.0, index=df.index)
        cost_val = pd.to_numeric(df[cols_lower["cost price"]], errors="coerce").fillna(0.0) if "cost price" in cols_lower else pd.Series(0.0, index=df.index)
        pf_val = pd.to_numeric(df[cols_lower["claims & compensation"]], errors="coerce").fillna(0.0) if "claims & compensation" in cols_lower else pd.Series(0.0, index=df.index)
        sale_amt = pd.to_numeric(df[cols_lower["order amount"]], errors="coerce").fillna(0.0) if "order amount" in cols_lower else pd.Series(0.0, index=df.index)
        rpr_chg = pd.to_numeric(df[cols_lower["rpr charges (h)"]], errors="coerce").fillna(0.0) if "rpr charges (h)" in cols_lower else pd.Series(0.0, index=df.index)
        rto_chg = pd.to_numeric(df[cols_lower["rto charges (i)"]], errors="coerce").fillna(0.0) if "rto charges (i)" in cols_lower else pd.Series(0.0, index=df.index)

        is_cancelled = order_st.str.contains("CANCEL") | (pay_st == "CANCELLED ORDER")
        is_rto = (~is_cancelled) & order_st.isin(["COURIER RETURN", "RTO"])
        is_cir = (~is_cancelled) & order_st.isin(["CUSTOMER RETURN", "CIR"])
        is_loss = (~is_cancelled) & (~is_rto) & (~is_cir) & (order_st == "LOSS INITIATED")

        sku_col = cols_lower.get("uni sku", cols_lower.get("sku"))
        pname_col = cols_lower.get("product name", sku_col)

        norm = pd.DataFrame(index=df.index)
        norm["Source_Format"] = "SNAPDEAL_RECO"
        norm["Marketplace"] = "Snapdeal"
        norm["Order_ID"] = df[cols_lower["suborder code"]].astype(str) if "suborder code" in cols_lower else ""
        norm["Sub_Order_ID"] = norm["Order_ID"]
        norm["Invoice_No"] = df[cols_lower["supc"]].astype(str) if "supc" in cols_lower else ""
        norm["Order_Date"] = df[cols_lower["order_date"]] if "order_date" in cols_lower else ""
        norm["Delivery_Date"] = df[cols_lower["del_date"]] if "del_date" in cols_lower else ""
        norm["Return_Date"] = df[cols_lower["rpr_date"]] if "rpr_date" in cols_lower else ""
        norm["SKU_Code"] = df[sku_col].fillna("UNKNOWN_SKU").astype(str).str.strip() if sku_col else "UNKNOWN_SKU"
        norm["Product_Name"] = df[pname_col].fillna("Snapdeal Product").astype(str).str.strip() if pname_col else "Snapdeal Product"
        norm["Parent_Style"] = norm["SKU_Code"].apply(get_parent_style)
        norm["Qty"] = 1.0
        norm["Gross_Sale"] = sale_amt
        norm["Actual_Settlement"] = act_val
        norm["Expected_Settlement"] = exp_val
        norm["Unit_COGS"] = cost_val
        norm["Effective_COGS"] = cost_val
        norm["Protection_Fund"] = pf_val
        norm["Loss_Initiated_Flag"] = is_loss

        norm["Status_Bucket"] = np.where(
            is_cancelled, "Cancelled",
            np.where(is_rto, "RTO",
            np.where(is_cir, "Customer Return",
            np.where(is_loss, "Loss Initiated", "Delivered (Success)")))
        )
        norm["Raw_Status"] = order_st
        norm["Return_Charge"] = np.where(
            act_val < 0, -act_val,
            np.where(is_cir & (act_val > 0), np.abs(exp_val), np.abs(rpr_chg))
        )
        norm["RTO_Charge"] = np.where(is_rto, np.abs(act_val) + np.abs(rto_chg), 0.0)
        norm["Good_Return_Qty"] = np.where(is_cir | is_rto, 1.0, 0.0)
        norm["Bad_Return_Qty"] = np.where(is_loss, 1.0, 0.0)
        norm["Bad_Return_Loss"] = np.where(is_loss, cost_val, 0.0)
        norm["Return_Reason"] = np.where(
            is_loss, "Courier Loss Initiated (Claimable)",
            np.where(is_cir, "Customer Return (RPR)", "Courier Return (RTO)")
        )
        norm["Tracking_ID"] = norm["Order_ID"]
        norm["Is_Unpaid_Delivered"] = (norm["Status_Bucket"] == "Delivered (Success)") & (act_val == 0)
        norm["Pending_Payout_Amt"] = np.where(act_val == 0, exp_val, 0.0)
        norm["Delivered_Profit"] = np.where(
            norm["Status_Bucket"] == "Delivered (Success)",
            np.where(act_val != 0, act_val, exp_val) - cost_val,
            0.0
        )
        return norm

    # --- 3. STANDARD ORDER INFO CSV / ZIP ---
    else:
        df = raw_df.copy()
        num_cols = [
            "Total", "Expected Payment", "Cost of Goods", "Calc Last Purchase Price",
            "Qty", "Expected Profit", "Return Good Qty", "Return Bad Qty",
            "Reverse Shipping Fee", "Settlement Amount"
        ]
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            else:
                df[col] = 0.0

        st_s = df["Order Status"].fillna("").astype(str).str.strip()
        is_cancelled = st_s == "Cancelled"
        is_rto = st_s.isin(["Cancelled Return Received", "Cancel Init"])
        is_cir = st_s.isin(["Return Init", "Return Received"])
        is_in_transit = st_s.isin(["In Transit", "Ready to ship", "Packed", "New"])

        eff_cogs = np.where(df["Cost of Goods"] > 0, df["Cost of Goods"], df["Calc Last Purchase Price"] * df["Qty"])
        ret_chg = np.where(df["Expected Payment"] < 0, -df["Expected Payment"], df["Reverse Shipping Fee"])
        bad_loss = df["Return Bad Qty"] * df["Calc Last Purchase Price"]

        norm = pd.DataFrame(index=df.index)
        norm["Source_Format"] = "ORDER_INFO"
        norm["Marketplace"] = df["Channel Name"].apply(clean_order_info_mp) if "Channel Name" in df.columns else "Marketplace"
        norm["Order_ID"] = df["Channel Order Id"].astype(str) if "Channel Order Id" in df.columns else ""
        norm["Sub_Order_ID"] = df["Channel Sub Order Id"].astype(str) if "Channel Sub Order Id" in df.columns else ""
        norm["Invoice_No"] = df["Invoice Number"].fillna("").astype(str) if "Invoice Number" in df.columns else ""
        norm["Order_Date"] = df["Order Date"] if "Order Date" in df.columns else ""
        norm["Delivery_Date"] = df["Delivery Date"] if "Delivery Date" in df.columns else ""
        norm["Return_Date"] = df["Return Date"] if "Return Date" in df.columns else ""
        norm["SKU_Code"] = df["Product Sku Code"].fillna("UNKNOWN_SKU").astype(str).str.strip() if "Product Sku Code" in df.columns else "UNKNOWN_SKU"
        norm["Product_Name"] = df["Product Name"].fillna("Unknown Product").astype(str).str.strip() if "Product Name" in df.columns else "Product"
        norm["Parent_Style"] = norm["SKU_Code"].apply(get_parent_style)
        norm["Qty"] = df["Qty"]
        norm["Gross_Sale"] = df["Total"]
        norm["Actual_Settlement"] = df["Settlement Amount"]
        norm["Expected_Settlement"] = df["Expected Payment"]
        norm["Unit_COGS"] = df["Calc Last Purchase Price"]
        norm["Effective_COGS"] = eff_cogs
        norm["Protection_Fund"] = bad_loss
        norm["Loss_Initiated_Flag"] = False

        norm["Status_Bucket"] = np.where(
            is_cancelled, "Cancelled",
            np.where(is_rto, "RTO",
            np.where(is_cir, "Customer Return",
            np.where(is_in_transit, "In-Transit / Active", "Delivered (Success)")))
        )
        norm["Raw_Status"] = st_s
        norm["Return_Charge"] = np.where(is_cir, ret_chg, 0.0)
        norm["RTO_Charge"] = np.where(is_rto, ret_chg, 0.0)
        norm["Good_Return_Qty"] = df["Return Good Qty"]
        norm["Bad_Return_Qty"] = df["Return Bad Qty"]
        norm["Bad_Return_Loss"] = bad_loss
        norm["Return_Reason"] = df["Return Reason"].fillna("Not Specified").astype(str).str.strip() if "Return Reason" in df.columns else "Not Specified"
        norm["Tracking_ID"] = df["Shipment Tracker"].fillna("").astype(str) if "Shipment Tracker" in df.columns else ""

        pay_rec = df["Payment Received"].fillna("No").astype(str).str.strip().str.lower() if "Payment Received" in df.columns else pd.Series("no", index=df.index)
        norm["Is_Unpaid_Delivered"] = (norm["Status_Bucket"] == "Delivered (Success)") & (pay_rec != "yes")
        norm["Pending_Payout_Amt"] = np.where(norm["Is_Unpaid_Delivered"], df["Expected Payment"], 0.0)
        norm["Delivered_Profit"] = np.where(norm["Status_Bucket"] == "Delivered (Success)", df["Expected Profit"], 0.0)
        return norm

def detect_month_tag(df):
    if "Order_Date" not in df.columns:
        return "Report"
    parsed = pd.to_datetime(df["Order_Date"], errors="coerce").dropna()
    if parsed.empty:
        return "Report"
    return parsed.max().strftime("%b-%y")

def auto_fit_columns(ws, skip_cols=None, max_width=42, min_width=10, header_row=4):
    if skip_cols is None:
        skip_cols = set()
    for col in ws.columns:
        col_idx = col[0].column
        col_letter = get_column_letter(col_idx)
        if col_letter in skip_cols:
            continue
        max_len = 0
        for cell in col:
            if cell.row < header_row and cell.row in [1, 2]:
                continue
            val = str(cell.value or "")
            if val.startswith("="):
                max_len = max(max_len, 11)
            else:
                lines = val.split("\n")
                longest_line = max((len(line) for line in lines), default=0)
                max_len = max(max_len, longest_line)
        ws.column_dimensions[col_letter].width = min(max(max_len + 2.5, min_width), max_width)

def build_master_workbook(df, month_tag):
    wb = openpyxl.Workbook()
    default_ws = wb.active

    FONT_SHEET_TITLE = Font(name="Calibri", size=14, bold=True, color="1B365D")
    FONT_TBL_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_LABEL = Font(name="Calibri", size=10, color="1A202C")
    FONT_LABEL_BOLD = Font(name="Calibri", size=10, bold=True, color="1A202C")
    FONT_VAL = Font(name="Calibri", size=10, color="1A202C")
    FONT_VAL_BOLD = Font(name="Calibri", size=10, bold=True, color="1A202C")
    FONT_NOTE_HINT = Font(name="Calibri", size=9, italic=True, color="4A5568")
    FONT_RED = Font(name="Calibri", size=10, bold=True, color="9B2C2C")
    FONT_AMBER = Font(name="Calibri", size=10, bold=True, color="975A16")

    FILL_NAVY = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    FILL_SLATE = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
    FILL_PEACH = PatternFill(start_color="FBD38D", end_color="FBD38D", fill_type="solid")
    FILL_MINT = PatternFill(start_color="C6F6D5", end_color="C6F6D5", fill_type="solid")
    FILL_RED_SOFT = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid")
    FILL_TOTAL = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")

    BORDER_CELL = Border(
        left=Side(style="thin", color="CBD5E0"),
        right=Side(style="thin", color="CBD5E0"),
        top=Side(style="thin", color="CBD5E0"),
        bottom=Side(style="thin", color="CBD5E0")
    )
    BORDER_BOX = Border(
        left=Side(style="thin", color="718096"),
        right=Side(style="thin", color="718096"),
        top=Side(style="thin", color="718096"),
        bottom=Side(style="thin", color="718096")
    )

    ordered_mps = list(df["Marketplace"].unique())

    # 1. CHANNEL-WISE RECO TABS
    for mp in ordered_mps:
        sub = df[df["Marketplace"] == mp].copy()
        tab_title = f"{mp} {month_tag}"[:31]
        ws = wb.create_sheet(title=tab_title)
        ws.views.sheetView[0].showGridLines = True

        tot_order = len(sub)
        cancelled = int((sub["Status_Bucket"] == "Cancelled").sum())
        rto_rows = sub[sub["Status_Bucket"] == "RTO"]
        cret_rows = sub[sub["Status_Bucket"] == "Customer Return"]
        loss_rows = sub[sub["Status_Bucket"] == "Loss Initiated"]

        rto = len(rto_rows)
        c_return = len(cret_rows)
        loss_cnt = len(loss_rows)

        rto_charge = round(rto_rows["RTO_Charge"].sum())
        pf_sum = round(abs(sub["Protection_Fund"].sum()))
        pf_cnt = int((sub["Protection_Fund"] != 0).sum())

        neg_settle_sum = round(abs(sub.loc[sub["Actual_Settlement"] < 0, "Actual_Settlement"].sum()))
        pos_cir_rows = cret_rows[cret_rows["Actual_Settlement"] > 0]
        exp_pos_cir_sum = round(abs(pos_cir_rows["Expected_Settlement"].sum()))
        if neg_settle_sum == 0 and exp_pos_cir_sum == 0:
            exp_pos_cir_sum = round(cret_rows["Return_Charge"].sum())

        succ_rows = sub[sub["Status_Bucket"].isin(["Delivered (Success)", "In-Transit / Active"])]
        del_rows = sub[sub["Status_Bucket"] == "Delivered (Success)"]

        b14_count = len(succ_rows)
        del_only_count = len(del_rows)
        in_transit_count = b14_count - del_only_count

        b15_actual_pay = round(abs(succ_rows.loc[succ_rows["Actual_Settlement"] != 0, "Actual_Settlement"].sum()))
        zero_act_rows = succ_rows[succ_rows["Actual_Settlement"] == 0]
        c15_exp_pay = round(abs(zero_act_rows["Expected_Settlement"].sum()))

        b16_cogs = round(abs(succ_rows["Effective_COGS"].sum()))
        zero_cogs_cnt = int((succ_rows["Effective_COGS"] == 0).sum())

        del_only_exp_pay = round(del_rows["Expected_Settlement"].sum())
        del_only_act_pay = round(del_rows["Actual_Settlement"].sum())
        del_only_payout = del_only_act_pay if del_only_act_pay != 0 else del_only_exp_pay
        del_only_cogs = round(del_rows["Effective_COGS"].sum())
        del_ret_pen = neg_settle_sum + exp_pos_cir_sum

        ws["A1"] = tab_title
        ws["A1"].font = FONT_SHEET_TITLE
        ws["B1"] = "RECO SUMMARY"
        ws["B1"].font = Font(name="Calibri", size=10, bold=True, color="1B365D")
        ws["B1"].fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["B1"].border = BORDER_BOX

        ws["E1"] = "Other charges"
        ws["E1"].font = FONT_LABEL_BOLD
        ws["F1"] = "Amount (₹)"
        ws["F1"].font = FONT_LABEL_BOLD
        ws["F1"].alignment = Alignment(horizontal="right")

        ws["H1"] = "Other CN / SPF"
        ws["H1"].font = FONT_LABEL_BOLD
        ws["I1"] = "Amount (₹)"
        ws["I1"].font = FONT_LABEL_BOLD
        ws["I1"].alignment = Alignment(horizontal="right")

        ws["E2"] = "VFH / VFS Charges (Manual Entry)"
        ws["F2"] = 0
        ws["E3"] = "MFB Marketing Charges (Manual Entry)"
        ws["F3"] = 0
        ws["E4"] = "Ad Services / Promotion Fee (Manual)"
        ws["F4"] = 0

        ws["H2"] = f"Protection Fund / SPF Claim ({pf_cnt} orders)"
        ws["I2"] = pf_sum
        ws["H3"] = "Other Credit Notes (Manual Entry)"
        ws["I3"] = 0

        for r_idx in range(2, 5):
            ws.cell(row=r_idx, column=5).font = FONT_NOTE_HINT
            ws.cell(row=r_idx, column=6).number_format = "#,##0"
            ws.cell(row=r_idx, column=6).border = BORDER_CELL
            if r_idx <= 3:
                ws.cell(row=r_idx, column=8).font = FONT_NOTE_HINT
                ws.cell(row=r_idx, column=9).number_format = "#,##0"
                ws.cell(row=r_idx, column=9).border = BORDER_CELL

        rows_spec = [
            (2, "Total Order", tot_order, None, None, "#,##0", None),
            (3, "Cancelled", cancelled, None, None, "#,##0", None),
            (4, "RTO", rto, None, None, "#,##0", None),
            (5, "SPF Claim Amount", pf_sum, None, None, "#,##0",
             f"{pf_cnt} order(s) me Protection Fund / Claim ki non-zero value mili hai."),
            (6, "Loss Initiated", loss_cnt, None, None, "#,##0",
             "Courier lost parcels (Loss Initiated). Success orders se alag rakha gaya hai."),
            (8, "C. Return", c_return, None, None, "#,##0", None),
            (9, "RTO Charge", rto_charge, None, None, "#,##0", None),
            (10, "C. Return Charge", neg_settle_sum, exp_pos_cir_sum, None, "#,##0",
             f"B10 = Negative Actual Settlement (₹{neg_settle_sum:,}). C10 = Expected Settlement on {len(pos_cir_rows)} positive CIR / unsettled returns (₹{exp_pos_cir_sum:,})."),
            (11, "Per order return charge", "=IFERROR(ROUND((B10+C10)/B8, 2), 0)", None, None, "#,##0.00", None),
            (12, "Return %", "=IFERROR((B4+B8)/(B2-B3), 0)", None, None, "0%", None),
            (14, "Success order", "=B2-B3-B4-B6-B8", del_only_count, None, "#,##0",
             f"B14 = {b14_count} ({del_only_count} Delivered + {in_transit_count} In-Transit). C14 = {del_only_count} (Strictly Delivered only)."),
            (15, "Success order payout", b15_actual_pay, c15_exp_pay, None, "#,##0",
             f"B15 = Settled Actual Payout (₹{b15_actual_pay:,}). C15 = Expected Payout for {len(zero_act_rows)} unsettled/0-actual orders (₹{c15_exp_pay:,})."),
            (16, "Success Order product cost", b16_cogs, del_only_cogs, None, "#,##0",
             f"B16 = ₹{b16_cogs:,} COGS ({zero_cogs_cnt} orders me Cost Price 0 mili). C16 = ₹{del_only_cogs:,} COGS for Delivered-only orders."),
            (17, "Per order return charge on success order", "=IFERROR(ROUND((B10+C10)/B14, 0), 0)", None, None, "#,##0", None),
            (19, "Gross Profit", "=ROUND(ABS(B5+B15+C15-B16), 0)", None, None, "#,##0", None),
            (20, "Less Return Charge", "=ROUND(ABS(B10+C10), 0)", None, FILL_PEACH, "#,##0", None),
            (21, "Less Other Charges", "=ROUND(SUM(F2:F21), 0)", None, FILL_PEACH, "#,##0", None),
            (23, "Add other CN", "=ROUND(SUM(I2:I21), 0)", None, FILL_MINT, "#,##0", None),
            (25, "Net Profit", "=ROUND(B19-B20-B21+B23, 0)", None, None, "#,##0", None),
            (27, "Profit per success Order", "=IFERROR(ROUND(B25/B14, 2), 0)", None, None, "#,##0.00", None),
        ]

        for r_num, lbl, val_b, val_c, fill_b, num_fmt, note_txt in rows_spec:
            c_a = ws.cell(row=r_num, column=1, value=lbl)
            c_b = ws.cell(row=r_num, column=2, value=val_b)

            is_key_row = r_num in [19, 25, 27]
            c_a.font = FONT_LABEL_BOLD if is_key_row else FONT_LABEL
            c_b.font = FONT_VAL_BOLD if is_key_row else FONT_VAL
            c_b.number_format = num_fmt
            c_b.alignment = Alignment(horizontal="right", vertical="center")
            c_a.border = BORDER_CELL
            c_b.border = BORDER_CELL

            if fill_b:
                c_b.fill = fill_b
            elif r_num == 25:
                c_b.fill = FILL_MINT

            if val_c is not None:
                c_c = ws.cell(row=r_num, column=3, value=val_c)
                c_c.font = FONT_VAL
                c_c.number_format = "#,##0"
                c_c.alignment = Alignment(horizontal="right", vertical="center")
                c_c.border = BORDER_CELL

            if note_txt:
                c_b.comment = Comment(note_txt, "Payment Reco")

        ws["K1"] = "DELIVERED-ONLY ACTUALS"
        ws["K1"].font = FONT_TBL_HDR
        ws["K1"].fill = FILL_SLATE
        ws["K1"].alignment = Alignment(horizontal="left", vertical="center")
        ws["L1"] = "Amount / Qty"
        ws["L1"].font = FONT_TBL_HDR
        ws["L1"].fill = FILL_SLATE
        ws["L1"].alignment = Alignment(horizontal="right", vertical="center")

        del_box = [
            ("Strictly Delivered Orders", del_only_count, "#,##0"),
            ("In-Transit / Active Orders", in_transit_count, "#,##0"),
            ("Delivered Gross Sale (₹)", round(del_rows["Gross_Sale"].sum()), "₹#,##0"),
            ("Delivered Net Payout (₹)", round(del_only_payout), "₹#,##0"),
            ("Delivered Product Cost (₹)", round(del_only_cogs), "₹#,##0"),
            ("Return Penalty on Delivered (₹)", round(del_ret_pen), "₹#,##0"),
            ("Return Charge / Delivered (₹)", "=IFERROR(ROUND(L7/L2, 2), 0)", "₹#,##0.00"),
            ("True Delivered Net Profit (₹)", "=L5-L6-L7", "₹#,##0"),
            ("Profit / Delivered Order (₹)", "=IFERROR(ROUND(L9/L2, 0), 0)", "₹#,##0"),
        ]
        for idx_b, (b_lbl, b_val, b_fmt) in enumerate(del_box, start=2):
            ca = ws.cell(row=idx_b, column=11, value=b_lbl)
            cb = ws.cell(row=idx_b, column=12, value=b_val)
            ca.font = FONT_LABEL_BOLD if idx_b >= 9 else FONT_LABEL
            cb.font = FONT_VAL_BOLD if idx_b >= 9 else FONT_VAL
            cb.number_format = b_fmt
            cb.alignment = Alignment(horizontal="right")
            ca.border = BORDER_CELL
            cb.border = BORDER_CELL
            if idx_b == 9:
                cb.fill = FILL_MINT

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            if col_letter in {"D", "G", "J"}:
                ws.column_dimensions[col_letter].width = 2.5
                continue
            max_len = 0
            for cell in col:
                val = str(cell.value or "")
                if val.startswith("="):
                    max_len = max(max_len, 10)
                else:
                    max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max(max_len + 2.5, 10)

    # 2. ALL CHANNELS SUMMARY TAB
    ws_exec = wb.create_sheet(title="All Channels Summary")
    ws_exec.views.sheetView[0].showGridLines = True
    ws_exec["A1"] = f"CONSOLIDATED MARKETPLACE PAYMENT RECO & PROFITABILITY ({month_tag.upper()})"
    ws_exec["A1"].font = FONT_SHEET_TITLE
    ws_exec["A2"] = "Ek Nazar Mein: All Channels Comparison — Order Status, Payouts, COGS, Return Burden & Unpaid Settlements"
    ws_exec["A2"].font = FONT_NOTE_HINT

    exec_headers = [
        "Marketplace", "Total Orders", "Delivered (Success)", "In-Transit / Active",
        "Cancelled", "RTO", "C. Return", "Return %",
        "Delivered Gross Sale (₹)", "Delivered Net Payout (₹)", "Delivered Product Cost (₹)",
        "C. Return Charges (₹)", "Damaged / Lost Stock (₹)", "Return Charge / Delivered (₹)",
        "True Delivered Net Profit (₹)", "Net Profit / Delivered (₹)",
        "Unpaid / Disputed Orders", "Pending Payout Amount (₹)", "In-Transit Pipeline Payout (₹)"
    ]
    ws_exec.row_dimensions[4].height = 28
    for c, h in enumerate(exec_headers, start=1):
        cell = ws_exec.cell(row=4, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_CELL

    for r_i, mp in enumerate(ordered_mps, start=5):
        sub = df[df["Marketplace"] == mp]
        del_s = sub[sub["Status_Bucket"] == "Delivered (Success)"]
        act_s = sub[sub["Status_Bucket"] == "In-Transit / Active"]
        can_cnt = int((sub["Status_Bucket"] == "Cancelled").sum())
        rto_cnt = int((sub["Status_Bucket"] == "RTO").sum())
        ret_cnt = int((sub["Status_Bucket"] == "Customer Return").sum())

        del_payout = round(np.where(del_s["Actual_Settlement"] != 0, del_s["Actual_Settlement"], del_s["Expected_Settlement"]).sum())
        act_payout = round(np.where(act_s["Actual_Settlement"] != 0, act_s["Actual_Settlement"], act_s["Expected_Settlement"]).sum())
        ret_chg = round(sub["Return_Charge"].sum())
        bad_loss = round(sub["Bad_Return_Loss"].sum())

        vals = [
            mp, len(sub), len(del_s), len(act_s), can_cnt, rto_cnt, ret_cnt,
            f"=IFERROR((F{r_i}+G{r_i})/(B{r_i}-E{r_i}), 0)",
            round(del_s["Gross_Sale"].sum()), del_payout, round(del_s["Effective_COGS"].sum()),
            ret_chg, bad_loss,
            f"=IFERROR(ROUND(L{r_i}/C{r_i}, 2), 0)",
            f"=J{r_i}-K{r_i}-L{r_i}-M{r_i}",
            f"=IFERROR(ROUND(O{r_i}/C{r_i}, 0), 0)",
            int(del_s["Is_Unpaid_Delivered"].sum()),
            round(del_s.loc[del_s["Is_Unpaid_Delivered"], "Pending_Payout_Amt"].sum()),
            act_payout
        ]
        for c_i, v in enumerate(vals, start=1):
            cell = ws_exec.cell(row=r_i, column=c_i, value=v)
            cell.font = FONT_LABEL_BOLD if c_i == 1 else FONT_VAL
            cell.border = BORDER_CELL
            if c_i in [2, 3, 4, 5, 6, 7, 17]:
                cell.number_format = "#,##0"
            elif c_i == 8:
                cell.number_format = "0.0%"
            elif c_i == 14:
                cell.number_format = "₹#,##0.00"
            elif c_i >= 9:
                cell.number_format = "₹#,##0"

    last_mp_row = 4 + len(ordered_mps)
    tot_row = last_mp_row + 1
    ws_exec.cell(row=tot_row, column=1, value="TOTAL").font = FONT_LABEL_BOLD
    ws_exec.cell(row=tot_row, column=1).fill = FILL_TOTAL
    ws_exec.cell(row=tot_row, column=1).border = BORDER_CELL
    for c_i in range(2, 20):
        col_l = get_column_letter(c_i)
        cell = ws_exec.cell(row=tot_row, column=c_i)
        if c_i == 8:
            cell.value = f"=IFERROR((F{tot_row}+G{tot_row})/(B{tot_row}-E{tot_row}), 0)"
            cell.number_format = "0.0%"
        elif c_i == 14:
            cell.value = f"=IFERROR(ROUND(L{tot_row}/C{tot_row}, 2), 0)"
            cell.number_format = "₹#,##0.00"
        elif c_i == 15:
            cell.value = f"=J{tot_row}-K{tot_row}-L{tot_row}-M{tot_row}"
            cell.number_format = "₹#,##0"
        elif c_i == 16:
            cell.value = f"=IFERROR(ROUND(O{tot_row}/C{tot_row}, 0), 0)"
            cell.number_format = "₹#,##0"
        else:
            cell.value = f"=SUM({col_l}5:{col_l}{last_mp_row})"
            cell.number_format = "#,##0" if c_i in [2, 3, 4, 5, 6, 7, 17] else "₹#,##0"
        cell.font = FONT_VAL_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER_CELL

    if len(ordered_mps) > 0:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Order Breakdown by Marketplace"
        chart.y_axis.title = "Orders"
        data_ref = Reference(ws_exec, min_col=3, max_col=7, min_row=4, max_row=last_mp_row)
        cats_ref = Reference(ws_exec, min_col=1, min_row=5, max_row=last_mp_row)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.width = 18
        chart.height = 10
        ws_exec.add_chart(chart, f"A{tot_row + 2}")

    auto_fit_columns(ws_exec, max_width=28, min_width=13, header_row=4)

    # 3. SKU PRICING & RETURN ACTIONS TAB
    ws_sku = wb.create_sheet(title="SKU Pricing & Return Actions")
    ws_sku.views.sheetView[0].showGridLines = True
    ws_sku["A1"] = "SKU PROFITABILITY, HIGH-RETURN & PRICE HIKE ACTION REPORT"
    ws_sku["A1"].font = FONT_SHEET_TITLE
    ws_sku["A2"] = "Loss-making & High-return (40%+) SKUs sorted at the top with Suggested Selling Price (₹)"
    ws_sku["A2"].font = FONT_NOTE_HINT

    sku_grp = df.groupby(["Marketplace", "Parent_Style", "SKU_Code", "Product_Name"]).agg(
        Total_Orders=("Order_ID", "count"),
        Delivered=("Status_Bucket", lambda x: (x == "Delivered (Success)").sum()),
        Returns=("Status_Bucket", lambda x: (x == "Customer Return").sum()),
        Cancelled=("Status_Bucket", lambda x: x.isin(["Cancelled", "RTO"]).sum()),
        Unit_COGS=("Unit_COGS", "max"),
        Avg_Price=("Gross_Sale", "mean"),
        Del_Profit=("Delivered_Profit", "sum"),
        Ret_Charges=("Return_Charge", "sum"),
        Bad_Loss=("Bad_Return_Loss", "sum")
    ).reset_index()

    sku_grp["Closed"] = sku_grp["Delivered"] + sku_grp["Returns"]
    sku_grp["Return_Rate"] = np.where(sku_grp["Closed"] > 0, sku_grp["Returns"] / sku_grp["Closed"], 0.0)
    sku_grp["Net_Profit"] = sku_grp["Del_Profit"] - sku_grp["Ret_Charges"] - sku_grp["Bad_Loss"]

    def sku_action(row):
        if row["Net_Profit"] < 0:
            return "🔴 LOSS — Increase Price & Audit", round(max(row["Avg_Price"] + 70, row["Unit_COGS"] * 1.65), -1) - 1
        elif row["Closed"] >= 5 and row["Return_Rate"] >= 0.40:
            return "🟠 HIGH RETURN (40%+) — Size/Quality Check", round(row["Avg_Price"] + 40, -1) - 1
        elif row["Delivered"] > 0 and (row["Net_Profit"] / row["Delivered"]) < 140:
            return "🟡 LOW MARGIN (<₹140) — Increase Price", round(row["Avg_Price"] + 50, -1) - 1
        return "🟢 HEALTHY", round(row["Avg_Price"], 0)

    act_res = sku_grp.apply(sku_action, axis=1)
    sku_grp["Action"] = [r[0] for r in act_res]
    sku_grp["Suggested_Price"] = [r[1] for r in act_res]
    sku_grp["Rank"] = sku_grp["Action"].apply(
        lambda x: 1 if "🔴" in x else (2 if "🟠" in x else (3 if "🟡" in x else 4))
    )
    sku_grp = sku_grp.sort_values(["Rank", "Net_Profit", "Returns"], ascending=[True, True, False])

    sku_headers = [
        "Recommended Action", "Marketplace", "Parent Style", "SKU Code", "Product Name",
        "Total Orders", "Delivered", "Returns", "Cancelled/RTO", "Return Rate (%)",
        "Unit COGS (₹)", "Current Avg Price (₹)", "Suggested Price (₹)",
        "Delivered Profit (₹)", "Return Charges (₹)", "Damaged/Lost (₹)", "True Net Profit (₹)"
    ]
    ws_sku.row_dimensions[4].height = 26
    for c, h in enumerate(sku_headers, start=1):
        cell = ws_sku.cell(row=4, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_CELL

    cols_sku = [
        "Action", "Marketplace", "Parent_Style", "SKU_Code", "Product_Name",
        "Total_Orders", "Delivered", "Returns", "Cancelled", "Return_Rate",
        "Unit_COGS", "Avg_Price", "Suggested_Price", "Del_Profit", "Ret_Charges", "Bad_Loss", "Net_Profit"
    ]

    for r_i, row in enumerate(sku_grp[cols_sku].itertuples(index=False), start=5):
        for c_i, val in enumerate(row, start=1):
            cell = ws_sku.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_VAL
            cell.border = BORDER_CELL
            if c_i in [6, 7, 8, 9]:
                cell.number_format = "#,##0"
            elif c_i == 10:
                cell.number_format = "0.0%"
            elif c_i >= 11:
                cell.number_format = "₹#,##0"
            if c_i == 1:
                if "🔴" in str(val):
                    cell.fill = FILL_RED_SOFT
                    cell.font = FONT_RED
                elif "🟠" in str(val) or "🟡" in str(val):
                    cell.fill = FILL_PEACH
                    cell.font = FONT_AMBER

    auto_fit_columns(ws_sku, max_width=42, min_width=11, header_row=4)
    ws_sku.freeze_panes = "A5"
    if len(sku_grp) > 0:
        ws_sku.auto_filter.ref = f"A4:Q{4 + len(sku_grp)}"

    # 4. PRODUCT-WISE RETURN REASON & ACTION PLAN TAB
    ws_prod_ret = wb.create_sheet(title="Product Return Action Plan")
    ws_prod_ret.views.sheetView[0].showGridLines = True
    ws_prod_ret["A1"] = "PRODUCT-WISE RETURN REASON ROOT-CAUSE & ACTION PLAN"
    ws_prod_ret["A1"].font = FONT_SHEET_TITLE
    ws_prod_ret["A2"] = "Kaunsa product kis reason se return aa raha hai + har product ke saamne exact Action Plan"
    ws_prod_ret["A2"].font = FONT_NOTE_HINT

    ret_all = df[df["Status_Bucket"].isin(["Customer Return", "Loss Initiated"])].copy()

    sku_closed = df.groupby(["Marketplace", "SKU_Code"]).agg(
        Total_Sold_Del=("Status_Bucket", lambda x: (x == "Delivered (Success)").sum()),
        Total_SKU_Returns=("Status_Bucket", lambda x: (x == "Customer Return").sum())
    ).reset_index()
    sku_closed["SKU_Return_Rate"] = np.where(
        (sku_closed["Total_Sold_Del"] + sku_closed["Total_SKU_Returns"]) > 0,
        sku_closed["Total_SKU_Returns"] / (sku_closed["Total_Sold_Del"] + sku_closed["Total_SKU_Returns"]),
        0.0
    )

    if not ret_all.empty:
        prod_reason_grp = ret_all.groupby(
            ["Marketplace", "Parent_Style", "SKU_Code", "Product_Name", "Return_Reason"]
        ).agg(
            Reason_Return_Count=("Order_ID", "count"),
            Good_Return_Qty=("Good_Return_Qty", "sum"),
            Bad_Return_Qty=("Bad_Return_Qty", "sum"),
            Return_Penalty_Cost=("Return_Charge", "sum"),
            Damaged_Stock_Loss=("Bad_Return_Loss", "sum")
        ).reset_index()

        prod_reason_grp = prod_reason_grp.merge(
            sku_closed, on=["Marketplace", "SKU_Code"], how="left"
        )
        prod_reason_grp["Total_Loss_On_Reason"] = (
            prod_reason_grp["Return_Penalty_Cost"] + prod_reason_grp["Damaged_Stock_Loss"]
        )
        prod_reason_grp["Action_Plan"] = prod_reason_grp.apply(
            lambda r: get_return_action_plan(r["Return_Reason"], r["SKU_Return_Rate"] * 100),
            axis=1
        )
        prod_reason_grp = prod_reason_grp.sort_values(
            ["Reason_Return_Count", "Total_Loss_On_Reason"], ascending=[False, False]
        )
    else:
        prod_reason_grp = pd.DataFrame(columns=[
            "Marketplace", "Parent_Style", "SKU_Code", "Product_Name",
            "Return_Reason", "Reason_Return_Count", "Total_Sold_Del", "Total_SKU_Returns",
            "SKU_Return_Rate", "Good_Return_Qty", "Bad_Return_Qty", "Total_Loss_On_Reason", "Action_Plan"
        ])

    pr_headers = [
        "Marketplace", "Parent Style", "SKU Code", "Product Name",
        "Primary Return Reason", "Returns (For This Reason)",
        "Total SKU Delivered", "Total SKU Returns", "Overall SKU Return %",
        "Good Qty", "Bad / Lost Qty", "Total Direct Loss (₹)",
        "Recommended Action Plan (Kya Karna Chahiye)"
    ]
    ws_prod_ret.row_dimensions[4].height = 28
    for c, h in enumerate(pr_headers, start=1):
        cell = ws_prod_ret.cell(row=4, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_CELL

    cols_pr = [
        "Marketplace", "Parent_Style", "SKU_Code", "Product_Name",
        "Return_Reason", "Reason_Return_Count",
        "Total_Sold_Del", "Total_SKU_Returns", "SKU_Return_Rate",
        "Good_Return_Qty", "Bad_Return_Qty", "Total_Loss_On_Reason",
        "Action_Plan"
    ]
    for r_i, row in enumerate(prod_reason_grp[cols_pr].itertuples(index=False), start=5):
        for c_i, val in enumerate(row, start=1):
            cell = ws_prod_ret.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_VAL
            cell.border = BORDER_CELL
            if c_i in [6, 7, 8, 10, 11]:
                cell.number_format = "#,##0"
            elif c_i == 9:
                cell.number_format = "0.0%"
                if float(val) >= 0.50:
                    cell.fill = FILL_RED_SOFT
                    cell.font = FONT_RED
                elif float(val) >= 0.35:
                    cell.fill = FILL_PEACH
                    cell.font = FONT_AMBER
            elif c_i == 12:
                cell.number_format = "₹#,##0"
            elif c_i == 13:
                cell.font = FONT_LABEL_BOLD

    auto_fit_columns(ws_prod_ret, max_width=56, min_width=11, header_row=4)
    ws_prod_ret.freeze_panes = "A5"
    if len(prod_reason_grp) > 0:
        ws_prod_ret.auto_filter.ref = f"A4:M{4 + len(prod_reason_grp)}"

    # 5. UNPAID / DISPUTED DELIVERED ORDERS TAB
    del_pending = df[df["Is_Unpaid_Delivered"]].copy()
    ws_unpaid = wb.create_sheet(title=f"Unpaid Delivered ({len(del_pending)})"[:31])
    ws_unpaid.views.sheetView[0].showGridLines = True
    ws_unpaid["A1"] = f"UNPAID / DISPUTED DELIVERED ORDERS — IMMEDIATE PAYMENT RECONCILIATION ({len(del_pending)} ORDERS)"
    ws_unpaid["A1"].font = FONT_SHEET_TITLE
    ws_unpaid["A2"] = f"Delivered orders with 0 Actual Settlement or Disputed/Overdue Status | Pending Amount: ₹{del_pending['Pending_Payout_Amt'].sum():,.0f}"
    ws_unpaid["A2"].font = FONT_NOTE_HINT

    if not del_pending.empty:
        del_pending["Delivery_Dt_Parsed"] = pd.to_datetime(del_pending["Delivery_Date"], errors="coerce")
        ref_dt = del_pending["Delivery_Dt_Parsed"].max() if del_pending["Delivery_Dt_Parsed"].notna().any() else pd.Timestamp.now()
        del_pending["Days_Since_Delivery"] = (ref_dt - del_pending["Delivery_Dt_Parsed"]).dt.days.fillna(0).astype(int)
        del_pending["Aging"] = np.where(
            del_pending["Days_Since_Delivery"] > 15, "Overdue (>15 Days)",
            np.where(del_pending["Days_Since_Delivery"] >= 7, "Due Soon (7-15 Days)", "Recent (<7 Days)")
        )
        del_pending = del_pending.sort_values(["Marketplace", "Days_Since_Delivery"], ascending=[True, False])

    unp_headers = [
        "Marketplace", "Order ID", "Invoice / SUPC", "Order Date", "Delivery Date",
        "Days Since Delivery", "Aging Status", "SKU Code", "Product Name", "Qty",
        "Order Value (₹)", "Pending Payout (₹)", "Product COGS (₹)", "Expected Profit (₹)", "Tracking / Code"
    ]
    ws_unpaid.row_dimensions[4].height = 26
    for c, h in enumerate(unp_headers, start=1):
        cell = ws_unpaid.cell(row=4, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_CELL

    cols_unp = [
        "Marketplace", "Order_ID", "Invoice_No", "Order_Date", "Delivery_Date",
        "Days_Since_Delivery", "Aging", "SKU_Code", "Product_Name", "Qty",
        "Gross_Sale", "Pending_Payout_Amt", "Effective_COGS", "Delivered_Profit", "Tracking_ID"
    ]
    if not del_pending.empty:
        for r_i, row in enumerate(del_pending[cols_unp].itertuples(index=False), start=5):
            for c_i, val in enumerate(row, start=1):
                clean_v = str(val).lstrip("`") if isinstance(val, str) else val
                cell = ws_unpaid.cell(row=r_i, column=c_i, value=clean_v)
                cell.font = FONT_VAL
                cell.border = BORDER_CELL
                if c_i in [6, 10]:
                    cell.number_format = "#,##0"
                elif c_i in [11, 12, 13, 14]:
                    cell.number_format = "₹#,##0.00"
                if c_i == 7:
                    if "Overdue" in str(clean_v):
                        cell.fill = FILL_RED_SOFT
                        cell.font = FONT_RED
                    elif "Due Soon" in str(clean_v):
                        cell.fill = FILL_PEACH
                        cell.font = FONT_AMBER

    auto_fit_columns(ws_unpaid, max_width=32, min_width=11, header_row=4)
    ws_unpaid.freeze_panes = "A5"
    if not del_pending.empty:
        ws_unpaid.auto_filter.ref = f"A4:O{4 + len(del_pending)}"

    # 6. RETURN REASON & SPF / LOSS CLAIM AUDIT TAB
    ws_ret = wb.create_sheet(title="Return & SPF Claim Audit")
    ws_ret.views.sheetView[0].showGridLines = True
    ws_ret["A1"] = "RETURN REASON BREAKDOWN & DAMAGED / LOST RETURN CLAIM LIST"
    ws_ret["A1"].font = FONT_SHEET_TITLE
    ws_ret["A2"] = "Top Table: Return Reasons Summary | Bottom Table: Damaged / Courier-Lost Orders Ready for SPF/Claim"
    ws_ret["A2"].font = FONT_NOTE_HINT

    if not ret_all.empty:
        reason_grp = ret_all.groupby(["Marketplace", "Return_Reason"]).agg(
            Return_Orders=("Order_ID", "count"),
            Good_Qty=("Good_Return_Qty", "sum"),
            Bad_Qty=("Bad_Return_Qty", "sum"),
            Return_Penalties=("Return_Charge", "sum"),
            Damaged_Loss=("Bad_Return_Loss", "sum")
        ).reset_index().sort_values(["Marketplace", "Return_Orders"], ascending=[True, False])
    else:
        reason_grp = pd.DataFrame(columns=["Marketplace", "Return_Reason", "Return_Orders", "Good_Qty", "Bad_Qty", "Return_Penalties", "Damaged_Loss"])

    r_headers = [
        "Marketplace", "Return Reason", "Return Orders", "Good Return Qty",
        "Bad / Lost Qty", "Return Penalties (₹)", "Damaged / Lost Stock (₹)"
    ]
    for c, h in enumerate(r_headers, start=1):
        cell = ws_ret.cell(row=4, column=c, value=h)
        cell.fill = FILL_SLATE
        cell.font = FONT_TBL_HDR
        cell.border = BORDER_CELL

    for r_i, row in enumerate(reason_grp.itertuples(index=False), start=5):
        for c_i, val in enumerate(row, start=1):
            cell = ws_ret.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_VAL
            cell.border = BORDER_CELL
            if c_i in [3, 4, 5]:
                cell.number_format = "#,##0"
            elif c_i in [6, 7]:
                cell.number_format = "₹#,##0"

    spf_start = 5 + len(reason_grp) + 3
    ws_ret.cell(row=spf_start - 1, column=1, value="DAMAGED / COURIER-LOST RETURN ORDERS (RAISE SPF / LOSS CLAIM)").font = FONT_SHEET_TITLE
    spf_headers = [
        "Marketplace", "Status", "Order ID", "Sub Order ID",
        "Return Date", "SKU Code", "Product Name", "Bad / Lost Qty",
        "Claimable Cost (₹)", "Return Reason", "Tracking ID"
    ]
    for c, h in enumerate(spf_headers, start=1):
        cell = ws_ret.cell(row=spf_start, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.border = BORDER_CELL

    bad_df = df[(df["Bad_Return_Qty"] > 0) | (df["Loss_Initiated_Flag"])].copy()
    cols_bad = [
        "Marketplace", "Raw_Status", "Order_ID", "Sub_Order_ID",
        "Return_Date", "SKU_Code", "Product_Name", "Bad_Return_Qty",
        "Bad_Return_Loss", "Return_Reason", "Tracking_ID"
    ]
    if not bad_df.empty:
        for idx_s, row in enumerate(bad_df[cols_bad].itertuples(index=False), start=spf_start + 1):
            for c_i, val in enumerate(row, start=1):
                clean_v = str(val).lstrip("`") if isinstance(val, str) else val
                cell = ws_ret.cell(row=idx_s, column=c_i, value=clean_v)
                cell.font = FONT_VAL
                cell.border = BORDER_CELL
                if c_i == 8:
                    cell.number_format = "#,##0"
                elif c_i == 9:
                    cell.number_format = "₹#,##0"

    auto_fit_columns(ws_ret, max_width=40, min_width=12, header_row=4)

    wb.remove(default_ws)
    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue(), ordered_mps

if uploaded_files:
    all_norms = []
    detected_types = []
    for f_obj in uploaded_files:
        for tbl_name, raw_tbl in read_raw_tables(f_obj):
            norm_tbl = detect_and_normalize(raw_tbl, tbl_name)
            all_norms.append(norm_tbl)
            detected_types.append(f"{tbl_name} → **{norm_tbl['Source_Format'].iloc[0]}** ({len(norm_tbl):,} rows)")

    df = pd.concat(all_norms, ignore_index=True)
    st.caption("Detected Files: " + " | ".join(detected_types))

    month_tag = detect_month_tag(df)
    excel_bytes, ordered_mps = build_master_workbook(df, month_tag)

    st.success(f"✅ Universal Master Workbook Ready! ({len(ordered_mps)} Channel RECO Tabs + 5 Action Reports in 1 Excel)")
    st.download_button(
        label=f"⬇️ Download Master Payment Reco Excel (Marketplace_Payment_Reco_Master_{month_tag}.xlsx)",
        data=excel_bytes,
        file_name=f"Marketplace_Payment_Reco_Master_{month_tag}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    summary_rows = []
    for mp in ordered_mps:
        grp = df[df["Marketplace"] == mp]
        d = grp[grp["Status_Bucket"] == "Delivered (Success)"]
        a = grp[grp["Status_Bucket"] == "In-Transit / Active"]
        c = grp[grp["Status_Bucket"] == "Cancelled"]
        rto = grp[grp["Status_Bucket"] == "RTO"]
        ret = grp[grp["Status_Bucket"] == "Customer Return"]

        denom = len(grp) - len(c)
        del_pay = np.where(d["Actual_Settlement"] != 0, d["Actual_Settlement"], d["Expected_Settlement"]).sum()
        ret_chg = grp["Return_Charge"].sum()
        bad_loss = grp["Bad_Return_Loss"].sum()
        net_prof = del_pay - d["Effective_COGS"].sum() - ret_chg - bad_loss

        summary_rows.append({
            "Marketplace": mp,
            "Total Orders": len(grp),
            "Delivered": len(d),
            "In-Transit": len(a),
            "Cancelled": len(c),
            "RTO": len(rto),
            "C. Return": len(ret),
            "Return %": round(((len(rto) + len(ret)) / denom * 100) if denom > 0 else 0, 1),
            "Delivered Payout (₹)": round(del_pay, 2),
            "Delivered Cost (₹)": round(d["Effective_COGS"].sum(), 2),
            "Return Charges (₹)": round(ret_chg, 2),
            "Return Chg / Delivered (₹)": round(ret_chg / len(d), 2) if len(d) > 0 else 0,
            "True Net Profit (₹)": round(net_prof, 2),
            "Unpaid / Disputed": int(d["Is_Unpaid_Delivered"].sum())
        })

    st.subheader("1. All Channels Consolidated Summary")
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

    st.subheader("2. Loss-Making & High-Return SKUs (Price Action Required)")
    sku_prev = df.groupby(["Marketplace", "SKU_Code", "Product_Name"]).agg(
        Delivered=("Status_Bucket", lambda x: (x == "Delivered (Success)").sum()),
        Returns=("Status_Bucket", lambda x: (x == "Customer Return").sum()),
        Unit_COGS=("Unit_COGS", "max"),
        Avg_Price=("Gross_Sale", "mean"),
        Del_Profit=("Delivered_Profit", "sum"),
        Ret_Charges=("Return_Charge", "sum"),
        Bad_Loss=("Bad_Return_Loss", "sum")
    ).reset_index()
    sku_prev["Return %"] = np.where(
        (sku_prev["Delivered"] + sku_prev["Returns"]) > 0,
        (sku_prev["Returns"] / (sku_prev["Delivered"] + sku_prev["Returns"])) * 100, 0
    ).round(1)
    sku_prev["Net Profit (₹)"] = (sku_prev["Del_Profit"] - sku_prev["Ret_Charges"] - sku_prev["Bad_Loss"]).round(2)
    sku_prev["Suggested Price (₹)"] = np.where(
        sku_prev["Net Profit (₹)"] < 0,
        np.maximum(sku_prev["Avg_Price"] + 70, sku_prev["Unit_COGS"] * 1.65).round(-1) - 1,
        np.where(sku_prev["Return %"] >= 40, (sku_prev["Avg_Price"] + 40).round(-1) - 1, sku_prev["Avg_Price"].round(0))
    )
    sku_prev["Action"] = np.where(
        sku_prev["Net Profit (₹)"] < 0, "🔴 LOSS: Increase Price",
        np.where(sku_prev["Return %"] >= 40, "🟠 HIGH RETURN: Audit Size/Quality", "🟢 HEALTHY")
    )
    st.dataframe(sku_prev.sort_values("Net Profit (₹)"), use_container_width=True)
