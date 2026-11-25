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

st.set_page_config(page_title="Payment Reco v2.2 — Single Master Workbook", layout="wide")
st.title("📊 Marketplace Payment Reconciliation (v2.2 — Auto-Fit + Product Return Action Plan)")
st.info("✅ Active Version: v2.2 (Auto-Adjusted Columns + L:M Delivered Box + Product-Level Return Reason & Action Plan Tab)")

uploaded_file = st.file_uploader("Upload Order / Payment Reco CSV or ZIP", type=["csv", "zip"])

def load_uploaded_data(file_obj):
    file_name = file_obj.name.lower()
    if file_name.endswith(".zip"):
        with zipfile.ZipFile(file_obj, "r") as z:
            csv_files = [
                f for f in z.namelist()
                if f.lower().endswith(".csv") and not f.startswith("__MACOSX")
            ]
            if not csv_files:
                st.error("ZIP folder ke andar koi .csv file nahi mili.")
                st.stop()
            with z.open(csv_files[0]) as f:
                return pd.read_csv(f, low_memory=False)
    else:
        return pd.read_csv(file_obj, low_memory=False)

def clean_mp(ch):
    s = str(ch)
    if "Myntra" in s:
        return "Myntra"
    if "Amazon" in s:
        return "Amazon"
    if "Ajio" in s:
        return "Ajio"
    return s.strip() if s.strip() else "Other"

def bucket_status(s):
    if s == "Delivered":
        return "Delivered (Success)"
    if s in ["In Transit", "Ready to ship", "Packed", "New"]:
        return "In-Transit / Active"
    if s in ["Cancelled", "Cancel Init", "Cancelled Return Received"]:
        return "Cancelled / RTO"
    if s in ["Return Init", "Return Received"]:
        return "Customer Return"
    return str(s)

def get_parent_style(sku):
    return re.sub(
        r"-(28|30|32|34|36|38|40|42|XS|S|M|L|XL|XXL|3XL)$",
        "",
        str(sku),
        flags=re.IGNORECASE
    )

def detect_month_tag(df):
    if "Order Date" not in df.columns:
        return "Report"
    parsed = pd.to_datetime(df["Order Date"], errors="coerce").dropna()
    if parsed.empty:
        return "Report"
    return parsed.max().strftime("%b-%y")

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
    else:
        if ret_rate_pct >= 40:
            return "🔍 High Return Audit: Top return comments check karke listing title, size chart aur packaging improve karein"
        return "✅ Monitor: Return comments track karein aur size chart accurate rakhein"

def auto_fit_columns(ws, skip_cols=None, max_width=45, min_width=12, header_row=4):
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
                max_len = max(max_len, 14)
            else:
                lines = val.split("\n")
                longest_line = max((len(line) for line in lines), default=0)
                max_len = max(max_len, longest_line)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, min_width), max_width)

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

    RTO_STATUSES = ["Cancelled Return Received", "Cancel Init"]
    CRETURN_STATUSES = ["Return Init", "Return Received"]

    ordered_mps = []
    for pref in ["Myntra", "Amazon", "Ajio"]:
        if pref in df["Marketplace"].unique():
            ordered_mps.append(pref)
    for other_mp in df["Marketplace"].unique():
        if other_mp not in ordered_mps:
            ordered_mps.append(other_mp)

    # --------------------------------------------------------------------------
    # 1. CHANNEL-WISE RECO TABS (Spacers at D, G, J | Delivered Box at L7:M16)
    # --------------------------------------------------------------------------
    for mp in ordered_mps:
        sub = df[df["Marketplace"] == mp].copy()
        tab_title = f"{mp} {month_tag}"[:31]
        ws = wb.create_sheet(title=tab_title)
        ws.views.sheetView[0].showGridLines = True

        tot_order = len(sub)
        cancelled = int((sub["Order Status"] == "Cancelled").sum())
        rto_rows = sub[sub["Order Status"].isin(RTO_STATUSES)]
        cret_rows = sub[sub["Order Status"].isin(CRETURN_STATUSES)]
        rto = len(rto_rows)
        c_return = len(cret_rows)

        rto_settle = rto_rows["Settlement Amount"].sum()
        rto_exp = rto_rows["Expected Payment"].sum()
        rto_charge = round(abs(rto_settle if rto_settle != 0 else (rto_exp if rto_exp < 0 else 0)))

        cr_settle = cret_rows["Settlement Amount"].sum()
        cr_exp = cret_rows["Expected Payment"].sum()
        cr_charge_settle = round(abs(cr_settle))
        cr_charge_exp = round(abs(cr_exp)) if cr_exp < 0 else 0

        cr_neg = int((cret_rows["Settlement Amount"] < 0).sum())
        cr_pos = int((cret_rows["Settlement Amount"] > 0).sum())
        cr_zero = int((cret_rows["Settlement Amount"] == 0).sum())
        cr_exp_neg = int((cret_rows["Expected Payment"] < 0).sum())

        succ_rows = sub[~sub["Order Status"].isin(["Cancelled"] + RTO_STATUSES + CRETURN_STATUSES)]
        del_rows = sub[sub["Order Status"] == "Delivered"]

        b13_count = len(succ_rows)
        del_only_count = len(del_rows)
        in_transit_count = b13_count - del_only_count

        b14_settle = round(succ_rows["Settlement Amount"].sum())
        s14_pos = int((succ_rows["Settlement Amount"] > 0).sum())
        s14_neg = int((succ_rows["Settlement Amount"] < 0).sum())
        s14_zero = int((succ_rows["Settlement Amount"] == 0).sum())

        c14_exp_pay = round(succ_rows.loc[succ_rows["Settlement Amount"] == 0, "Expected Payment"].sum())
        del_only_exp_pay = round(del_rows["Expected Payment"].sum())

        b15_cogs_raw = round(succ_rows["Cost of Goods"].sum())
        b15_cogs_eff = round(succ_rows["Effective_COGS"].sum())
        del_only_cogs = round(del_rows["Effective_COGS"].sum())

        cogs_pos = int((succ_rows["Cost of Goods"] > 0).sum())
        cogs_neg = int((succ_rows["Cost of Goods"] < 0).sum())
        cogs_zero = int((succ_rows["Cost of Goods"] == 0).sum())

        bad_ret_qty = int(sub["Return Bad Qty"].sum())
        bad_ret_loss = round(sub["Bad_Return_Loss"].sum())

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
        ws["E4"] = "Promotion Charges (Manual Entry)"
        ws["F4"] = 0

        ws["H2"] = f"Damaged Return SPF Claim ({bad_ret_qty} Bad Qty)"
        ws["I2"] = bad_ret_loss
        ws["I2"].comment = Comment(
            f"Auto-detected {bad_ret_qty} units with Return Bad Qty > 0 (Purchase Cost = ₹{bad_ret_loss}). Raise SPF claim on portal.",
            "Payment Reco"
        )
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
            (5, "RTO Charge", rto_charge, None, None, "#,##0", None),
            (6, "SPF Claim Amount", bad_ret_loss, None, None, "#,##0", None),
            (7, "C. Return", c_return, None, None, "#,##0", None),
            (8, "Per order return charge", "=IFERROR(ROUND(B9/B7, 2), 0)", None, None, "#,##0.00", None),
            (9, "C. Return Charge", cr_charge_settle, cr_charge_exp, None, "#,##0",
             f"{cr_neg} negative Settlement, {cr_pos} positive Settlement, {cr_zero} zero Settlement (out of {c_return} returns). C9 shows Expected Payment penalty (₹{cr_charge_exp}) from {cr_exp_neg} orders."),
            (11, "Return %", "=IFERROR((B4+B7)/(B2-B3), 0)", None, None, "0%", None),
            (13, "Success order", b13_count, del_only_count, None, "#,##0",
             f"B13 = {b13_count} (All non-Cancelled/RTO/Return orders: {del_only_count} Delivered + {in_transit_count} In-Transit). C13 = {del_only_count} (Strictly Delivered only)."),
            (14, "Success order payout", b14_settle, c14_exp_pay, None, "#,##0",
             f"{s14_pos} positive, {s14_neg} negative, {s14_zero} zero Settlement (out of {b13_count}). C14 = ₹{c14_exp_pay:,} Expected Payment across unsettled orders (Delivered-only = ₹{del_only_exp_pay:,})."),
            (15, "Success Order product cost", b15_cogs_eff, del_only_cogs, None, "#,##0",
             f"B15 = ₹{b15_cogs_eff:,} COGS for {b13_count} orders (Raw COGS = ₹{b15_cogs_raw:,}; {cogs_pos} positive, {cogs_zero} zero). C15 = ₹{del_only_cogs:,} COGS for {del_only_count} Delivered-only orders."),
            (16, "Per order return charge on success order", "=IFERROR(ROUND(MAX(B9,C9)/B13, 0), 0)", None, None, "#,##0", None),
            (18, "Gross Profit", "=B14+C14-B15", None, None, "#,##0", None),
            (19, "Less Return Charge", "=MAX(B9,C9)", None, FILL_PEACH, "#,##0", None),
            (20, "Less Other Charges", "=ROUND(SUM(F2:F20), 0)", None, FILL_PEACH, "#,##0", None),
            (22, "Add other CN", "=ROUND(SUM(I2:I20), 0)", None, FILL_MINT, "#,##0", None),
            (24, "Net Profit", "=B18-B19-B20+B22", None, None, "#,##0", None),
            (26, "Profit per success Order", "=IFERROR(ROUND(B24/B13, 0), 0)", None, None, "#,##0", None),
        ]

        for r_num, lbl, val_b, val_c, fill_b, num_fmt, note_txt in rows_spec:
            c_a = ws.cell(row=r_num, column=1, value=lbl)
            c_b = ws.cell(row=r_num, column=2, value=val_b)

            is_key_row = r_num in [18, 24, 26]
            c_a.font = FONT_LABEL_BOLD if is_key_row else FONT_LABEL
            c_b.font = FONT_VAL_BOLD if is_key_row else FONT_VAL
            c_b.number_format = num_fmt
            c_b.alignment = Alignment(horizontal="right", vertical="center")
            c_a.border = BORDER_CELL
            c_b.border = BORDER_CELL

            if fill_b:
                c_b.fill = fill_b
            elif r_num == 24:
                c_b.fill = FILL_MINT

            if val_c is not None:
                c_c = ws.cell(row=r_num, column=3, value=val_c)
                c_c.font = FONT_VAL
                c_c.number_format = "#,##0"
                c_c.alignment = Alignment(horizontal="right", vertical="center")
                c_c.border = BORDER_CELL

            if note_txt:
                c_b.comment = Comment(note_txt, "Payment Reco")

        # Side Box at L7:M16 (with Column J as spacer just like D & G!)
        ws["L7"] = "DELIVERED-ONLY ACTUALS (EXCLUDING IN-TRANSIT)"
        ws["L7"].font = FONT_TBL_HDR
        ws["L7"].fill = FILL_SLATE
        ws.merge_cells("L7:M7")

        del_box = [
            ("Strictly Delivered Orders", del_only_count, "#,##0"),
            ("In-Transit / Active Orders", in_transit_count, "#,##0"),
            ("Delivered Gross Sale (₹)", round(del_rows["Total"].sum()), "₹#,##0"),
            ("Delivered Expected Payout (₹)", del_only_exp_pay, "₹#,##0"),
            ("Delivered Product Cost (₹)", del_only_cogs, "₹#,##0"),
            ("Return Penalty on Delivered (₹)", cr_charge_exp, "₹#,##0"),
            ("Return Charge / Delivered Order (₹)", "=IFERROR(ROUND(M13/M8, 2), 0)", "₹#,##0.00"),
            ("True Delivered Net Profit (₹)", "=M11-M12-M13", "₹#,##0"),
            ("Profit / Delivered Order (₹)", "=IFERROR(ROUND(M15/M8, 0), 0)", "₹#,##0"),
        ]
        for idx_b, (b_lbl, b_val, b_fmt) in enumerate(del_box, start=8):
            ca = ws.cell(row=idx_b, column=12, value=b_lbl)
            cb = ws.cell(row=idx_b, column=13, value=b_val)
            ca.font = FONT_LABEL_BOLD if idx_b >= 15 else FONT_LABEL
            cb.font = FONT_VAL_BOLD if idx_b >= 15 else FONT_VAL
            cb.number_format = b_fmt
            cb.alignment = Alignment(horizontal="right")
            ca.border = BORDER_CELL
            cb.border = BORDER_CELL
            if idx_b == 15:
                cb.fill = FILL_MINT

        # Auto-adjust all data columns + narrow spacer columns D, G, J
        auto_fit_columns(ws, skip_cols={"D", "G", "J"}, max_width=44, min_width=14, header_row=1)
        ws.column_dimensions["D"].width = 4
        ws.column_dimensions["G"].width = 4
        ws.column_dimensions["J"].width = 4

    # --------------------------------------------------------------------------
    # 2. ALL CHANNELS SUMMARY TAB
    # --------------------------------------------------------------------------
    ws_exec = wb.create_sheet(title="All Channels Summary")
    ws_exec.views.sheetView[0].showGridLines = True
    ws_exec["A1"] = f"CONSOLIDATED MARKETPLACE PAYMENT RECO & PROFITABILITY ({month_tag.upper()})"
    ws_exec["A1"].font = FONT_SHEET_TITLE
    ws_exec["A2"] = "Ek Nazar Mein: All Channels Comparison — Order Status, Payouts, COGS, Return Burden & Unpaid Settlements"
    ws_exec["A2"].font = FONT_NOTE_HINT

    exec_headers = [
        "Marketplace", "Total Orders", "Delivered (Success)", "In-Transit / Active",
        "Cancelled", "RTO", "C. Return", "Return %",
        "Delivered Gross Sale (₹)", "Delivered Expected Payout (₹)", "Delivered Product Cost (₹)",
        "C. Return Charges (₹)", "Damaged Return Loss (₹)", "Return Charge / Delivered (₹)",
        "True Delivered Net Profit (₹)", "Net Profit / Delivered (₹)",
        "Unpaid Delivered Orders", "Unpaid Delivered Payout (₹)", "In-Transit Pipeline Payout (₹)"
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
        can_cnt = int((sub["Order Status"] == "Cancelled").sum())
        rto_cnt = int(sub["Order Status"].isin(RTO_STATUSES).sum())
        ret_cnt = int(sub["Order Status"].isin(CRETURN_STATUSES).sum())

        ret_chg = round(sub["Return_Charge"].sum())
        bad_loss = round(sub["Bad_Return_Loss"].sum())

        vals = [
            mp, len(sub), len(del_s), len(act_s), can_cnt, rto_cnt, ret_cnt,
            f"=IFERROR((F{r_i}+G{r_i})/(B{r_i}-E{r_i}), 0)",
            round(del_s["Total"].sum()), round(del_s["Expected Payment"].sum()), round(del_s["Effective_COGS"].sum()),
            ret_chg, bad_loss,
            f"=IFERROR(ROUND(L{r_i}/C{r_i}, 2), 0)",
            f"=J{r_i}-K{r_i}-L{r_i}-M{r_i}",
            f"=IFERROR(ROUND(O{r_i}/C{r_i}, 0), 0)",
            int(del_s["Is_Unpaid_Delivered"].sum()),
            round(del_s.loc[del_s["Is_Unpaid_Delivered"], "Expected Payment"].sum()),
            round(act_s["Expected Payment"].sum())
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

    auto_fit_columns(ws_exec, max_width=28, min_width=15, header_row=4)

    # --------------------------------------------------------------------------
    # 3. SKU PRICING & RETURN ACTIONS TAB
    # --------------------------------------------------------------------------
    ws_sku = wb.create_sheet(title="SKU Pricing & Return Actions")
    ws_sku.views.sheetView[0].showGridLines = True
    ws_sku["A1"] = "SKU PROFITABILITY, HIGH-RETURN & PRICE HIKE ACTION REPORT"
    ws_sku["A1"].font = FONT_SHEET_TITLE
    ws_sku["A2"] = "Loss-making & High-return (40%+) SKUs sorted at the top with Suggested Selling Price (₹)"
    ws_sku["A2"].font = FONT_NOTE_HINT

    sku_grp = df.groupby(["Marketplace", "Parent_Style", "Product Sku Code", "Product Name"]).agg(
        Total_Orders=("Channel Order Id", "count"),
        Delivered=("Status_Bucket", lambda x: (x == "Delivered (Success)").sum()),
        Returns=("Status_Bucket", lambda x: (x == "Customer Return").sum()),
        Cancelled=("Status_Bucket", lambda x: (x == "Cancelled / RTO").sum()),
        Unit_COGS=("Calc Last Purchase Price", "max"),
        Avg_Price=("Total", "mean"),
        Del_Profit=("Expected Profit", lambda x: x[df.loc[x.index, "Status_Bucket"] == "Delivered (Success)"].sum()),
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
        "Delivered Profit (₹)", "Return Charges (₹)", "Damaged Loss (₹)", "True Net Profit (₹)"
    ]
    ws_sku.row_dimensions[4].height = 26
    for c, h in enumerate(sku_headers, start=1):
        cell = ws_sku.cell(row=4, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_CELL

    cols_sku = [
        "Action", "Marketplace", "Parent_Style", "Product Sku Code", "Product Name",
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

    auto_fit_columns(ws_sku, max_width=44, min_width=13, header_row=4)
    ws_sku.freeze_panes = "A5"
    if len(sku_grp) > 0:
        ws_sku.auto_filter.ref = f"A4:Q{4 + len(sku_grp)}"

    # --------------------------------------------------------------------------
    # 4. PRODUCT-WISE RETURN REASON & ACTION PLAN TAB (NEW!)
    # --------------------------------------------------------------------------
    ws_prod_ret = wb.create_sheet(title="Product Return Action Plan")
    ws_prod_ret.views.sheetView[0].showGridLines = True
    ws_prod_ret["A1"] = "PRODUCT-WISE RETURN REASON ROOT-CAUSE & ACTION PLAN"
    ws_prod_ret["A1"].font = FONT_SHEET_TITLE
    ws_prod_ret["A2"] = "Kaunsa product kis reason se return aa raha hai + har product ke saamne exact Action Plan"
    ws_prod_ret["A2"].font = FONT_NOTE_HINT

    ret_all = df[df["Status_Bucket"] == "Customer Return"].copy()

    sku_closed = df.groupby(["Marketplace", "Product Sku Code"]).agg(
        Total_Sold_Del=("Status_Bucket", lambda x: (x == "Delivered (Success)").sum()),
        Total_SKU_Returns=("Status_Bucket", lambda x: (x == "Customer Return").sum())
    ).reset_index()
    sku_closed["SKU_Return_Rate"] = np.where(
        (sku_closed["Total_Sold_Del"] + sku_closed["Total_SKU_Returns"]) > 0,
        sku_closed["Total_SKU_Returns"] / (sku_closed["Total_Sold_Del"] + sku_closed["Total_SKU_Returns"]),
        0.0
    )

    prod_reason_grp = ret_all.groupby(
        ["Marketplace", "Parent_Style", "Product Sku Code", "Product Name", "Return Reason"]
    ).agg(
        Reason_Return_Count=("Channel Order Id", "count"),
        Good_Return_Qty=("Return Good Qty", "sum"),
        Bad_Return_Qty=("Return Bad Qty", "sum"),
        Return_Penalty_Cost=("Return_Charge", "sum"),
        Damaged_Stock_Loss=("Bad_Return_Loss", "sum")
    ).reset_index()

    prod_reason_grp = prod_reason_grp.merge(
        sku_closed, on=["Marketplace", "Product Sku Code"], how="left"
    )
    prod_reason_grp["Total_Loss_On_Reason"] = (
        prod_reason_grp["Return_Penalty_Cost"] + prod_reason_grp["Damaged_Stock_Loss"]
    )
    prod_reason_grp["Action_Plan"] = prod_reason_grp.apply(
        lambda r: get_return_action_plan(r["Return Reason"], r["SKU_Return_Rate"] * 100),
        axis=1
    )
    prod_reason_grp = prod_reason_grp.sort_values(
        ["Reason_Return_Count", "Total_Loss_On_Reason"], ascending=[False, False]
    )

    pr_headers = [
        "Marketplace", "Parent Style", "SKU Code", "Product Name",
        "Primary Return Reason", "Returns (For This Reason)",
        "Total SKU Delivered", "Total SKU Returns", "Overall SKU Return %",
        "Good Qty", "Bad (Damaged) Qty", "Total Direct Loss (₹)",
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
        "Marketplace", "Parent_Style", "Product Sku Code", "Product Name",
        "Return Reason", "Reason_Return_Count",
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

    auto_fit_columns(ws_prod_ret, max_width=58, min_width=13, header_row=4)
    ws_prod_ret.freeze_panes = "A5"
    if len(prod_reason_grp) > 0:
        ws_prod_ret.auto_filter.ref = f"A4:M{4 + len(prod_reason_grp)}"

    # --------------------------------------------------------------------------
    # 5. UNPAID DELIVERED ORDERS TAB
    # --------------------------------------------------------------------------
    del_pending = df[df["Is_Unpaid_Delivered"]].copy()
    ws_unpaid = wb.create_sheet(title=f"Unpaid Delivered ({len(del_pending)})"[:31])
    ws_unpaid.views.sheetView[0].showGridLines = True
    ws_unpaid["A1"] = f"UNPAID DELIVERED ORDERS — IMMEDIATE PAYMENT RECONCILIATION LIST ({len(del_pending)} ORDERS)"
    ws_unpaid["A1"].font = FONT_SHEET_TITLE
    ws_unpaid["A2"] = f"All Delivered orders where Payment Received = 'No' | Total Unpaid Expected Payout: ₹{del_pending['Expected Payment'].sum():,.0f}"
    ws_unpaid["A2"].font = FONT_NOTE_HINT

    if not del_pending.empty:
        del_pending["Delivery_Dt_Parsed"] = pd.to_datetime(del_pending["Delivery Date"], errors="coerce")
        ref_dt = del_pending["Delivery_Dt_Parsed"].max() if del_pending["Delivery_Dt_Parsed"].notna().any() else pd.Timestamp.now()
        del_pending["Days_Since_Delivery"] = (ref_dt - del_pending["Delivery_Dt_Parsed"]).dt.days.fillna(0).astype(int)
        del_pending["Aging"] = np.where(
            del_pending["Days_Since_Delivery"] > 15, "Overdue (>15 Days)",
            np.where(del_pending["Days_Since_Delivery"] >= 7, "Due Soon (7-15 Days)", "Recent (<7 Days)")
        )
        del_pending = del_pending.sort_values(["Marketplace", "Days_Since_Delivery"], ascending=[True, False])

    unp_headers = [
        "Marketplace", "Channel Order ID", "Invoice Number", "Order Date", "Delivery Date",
        "Days Since Delivery", "Aging Status", "SKU Code", "Product Name", "Qty",
        "Order Total (₹)", "Expected Payout (₹)", "Product COGS (₹)", "Expected Profit (₹)", "Tracking ID"
    ]
    ws_unpaid.row_dimensions[4].height = 26
    for c, h in enumerate(unp_headers, start=1):
        cell = ws_unpaid.cell(row=4, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_CELL

    cols_unp = [
        "Marketplace", "Channel Order Id", "Invoice Number", "Order Date", "Delivery Date",
        "Days_Since_Delivery", "Aging", "Product Sku Code", "Product Name", "Qty",
        "Total", "Expected Payment", "Effective_COGS", "Expected Profit", "Shipment Tracker"
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

    auto_fit_columns(ws_unpaid, max_width=32, min_width=13, header_row=4)
    ws_unpaid.freeze_panes = "A5"
    if not del_pending.empty:
        ws_unpaid.auto_filter.ref = f"A4:O{4 + len(del_pending)}"

    # --------------------------------------------------------------------------
    # 6. RETURN REASON & SPF CLAIM AUDIT TAB
    # --------------------------------------------------------------------------
    ws_ret = wb.create_sheet(title="Return & SPF Claim Audit")
    ws_ret.views.sheetView[0].showGridLines = True
    ws_ret["A1"] = "RETURN REASON BREAKDOWN & DAMAGED RETURN SPF CLAIM LIST"
    ws_ret["A1"].font = FONT_SHEET_TITLE
    ws_ret["A2"] = "Top Table: Overall Return Reasons Summary | Bottom Table: Damaged/Wrong-Item Returns Ready for SPF Claim"
    ws_ret["A2"].font = FONT_NOTE_HINT

    reason_grp = ret_all.groupby(["Marketplace", "Return Reason"]).agg(
        Return_Orders=("Channel Order Id", "count"),
        Good_Qty=("Return Good Qty", "sum"),
        Bad_Qty=("Return Bad Qty", "sum"),
        Return_Penalties=("Return_Charge", "sum"),
        Damaged_Loss=("Bad_Return_Loss", "sum")
    ).reset_index().sort_values(["Marketplace", "Return_Orders"], ascending=[True, False])

    r_headers = [
        "Marketplace", "Return Reason", "Return Orders", "Good Return Qty",
        "Bad Return Qty", "Return Penalties (₹)", "Damaged Stock Loss (₹)"
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
    ws_ret.cell(row=spf_start - 1, column=1, value="DAMAGED / WRONG ITEM RETURN ORDERS (RAISE SPF CLAIM)").font = FONT_SHEET_TITLE
    spf_headers = [
        "Marketplace", "Order Status", "Channel Order ID", "Sub Order ID",
        "Return Date", "SKU Code", "Product Name", "Bad Qty",
        "Claimable Cost (₹)", "Return Reason", "Return Tracking ID"
    ]
    for c, h in enumerate(spf_headers, start=1):
        cell = ws_ret.cell(row=spf_start, column=c, value=h)
        cell.fill = FILL_NAVY
        cell.font = FONT_TBL_HDR
        cell.border = BORDER_CELL

    bad_df = df[df["Return Bad Qty"] > 0].copy().sort_values(["Marketplace", "Return Date"])
    cols_bad = [
        "Marketplace", "Order Status", "Channel Order Id", "Channel Sub Order Id",
        "Return Date", "Product Sku Code", "Product Name", "Return Bad Qty",
        "Bad_Return_Loss", "Return Reason", "Return Tracking Id"
    ]
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

    auto_fit_columns(ws_ret, max_width=40, min_width=14, header_row=4)

    wb.remove(default_ws)
    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue(), ordered_mps

if uploaded_file is not None:
    df = load_uploaded_data(uploaded_file)

    df["Marketplace"] = df["Channel Name"].apply(clean_mp)
    df["Status_Bucket"] = df["Order Status"].apply(bucket_status)
    df["Product Sku Code"] = df["Product Sku Code"].fillna("UNKNOWN_SKU").astype(str).str.strip()
    df["Product Name"] = df["Product Name"].fillna("Unknown Product").astype(str).str.strip()
    df["Return Reason"] = df["Return Reason"].fillna("Not Specified").astype(str).str.strip()
    df["Parent_Style"] = df["Product Sku Code"].apply(get_parent_style)

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

    df["Effective_COGS"] = np.where(
        df["Cost of Goods"] > 0,
        df["Cost of Goods"],
        df["Calc Last Purchase Price"] * df["Qty"]
    )
    df["Return_Charge"] = np.where(
        df["Expected Payment"] < 0,
        -df["Expected Payment"],
        df["Reverse Shipping Fee"]
    )
    df["Bad_Return_Loss"] = df["Return Bad Qty"] * df["Calc Last Purchase Price"]
    df["Is_Unpaid_Delivered"] = (
        (df["Status_Bucket"] == "Delivered (Success)") &
        (df["Payment Received"].fillna("No").str.lower() != "yes")
    )

    month_tag = detect_month_tag(df)
    excel_bytes, ordered_mps = build_master_workbook(df, month_tag)

    st.success(f"✅ Single Master Workbook Ready! ({len(ordered_mps) + 5} tabs in 1 Excel file)")
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
        c = grp[grp["Status_Bucket"] == "Cancelled / RTO"]
        r = grp[grp["Status_Bucket"] == "Customer Return"]

        closed_cnt = len(d) + len(r)
        ret_chg = r["Return_Charge"].sum() + c["Return_Charge"].sum()
        bad_loss = grp["Bad_Return_Loss"].sum()
        net_prof = d["Expected Profit"].sum() - ret_chg - bad_loss

        summary_rows.append({
            "Marketplace": mp,
            "Total Orders": len(grp),
            "Delivered": len(d),
            "In-Transit": len(a),
            "Cancelled/RTO": len(c),
            "Returns": len(r),
            "Return %": round((len(r) / closed_cnt * 100) if closed_cnt > 0 else 0, 1),
            "Delivered Payout (₹)": round(d["Expected Payment"].sum(), 2),
            "Delivered Cost (₹)": round(d["Effective_COGS"].sum(), 2),
            "Return Charges (₹)": round(ret_chg, 2),
            "Damaged Loss (₹)": round(bad_loss, 2),
            "Return Chg / Delivered (₹)": round(ret_chg / len(d), 2) if len(d) > 0 else 0,
            "True Net Profit (₹)": round(net_prof, 2),
            "Unpaid Delivered": int(d["Is_Unpaid_Delivered"].sum()),
            "Unpaid Amount (₹)": round(d.loc[d["Is_Unpaid_Delivered"], "Expected Payment"].sum(), 2)
        })

    st.subheader("1. All Channels Consolidated Summary")
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

    st.subheader("2. Product-Wise Return Reason & Action Plan (Top Return Issues)")
    ret_only = df[df["Status_Bucket"] == "Customer Return"].copy()
    pr_prev = ret_only.groupby(["Marketplace", "Product Sku Code", "Product Name", "Return Reason"]).agg(
        Returns_For_Reason=("Channel Order Id", "count"),
        Bad_Qty=("Return Bad Qty", "sum"),
        Return_Penalty=("Return_Charge", "sum"),
        Damaged_Loss=("Bad_Return_Loss", "sum")
    ).reset_index()
    pr_prev["Action Plan"] = pr_prev["Return Reason"].apply(lambda x: get_return_action_plan(x, 45))
    st.dataframe(pr_prev.sort_values("Returns_For_Reason", ascending=False), use_container_width=True)
