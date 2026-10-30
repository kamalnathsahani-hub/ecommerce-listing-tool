import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Multi-Channel Payment Reco & Profit Analyzer", layout="wide")

st.title("💰 Multi-Channel Payment & Profit Reconciliation")
st.markdown("Settlement Excel ya Multi-Channel CSV/Excel file upload karein. File format detect karke system automatically channel-wise aur SKU-wise live dynamic formulas ke saath report generate karega.")

# File Uploader
uploaded_file = st.file_uploader("📂 Upload Settlement / Reco File (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, low_memory=False)
        else:
            xls = pd.ExcelFile(uploaded_file)
            if 'Orders' in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name='Orders', header=1)
                df = df.dropna(subset=['Order ID']).copy()
            else:
                df = pd.read_excel(uploaded_file)
                df = df.dropna(how='all')

        def find_col(possible_names):
            for col in df.columns:
                for name in possible_names:
                    if str(name).strip().lower() == str(col).strip().lower() or name.lower() in str(col).lower():
                        return col
            return None

        # Detect columns
        channel_col = find_col(['Channel Code', 'Channel'])
        order_id_col = find_col(['Channel Order Code', 'Order ID', 'Order Code'])
        sku_col = find_col(['Sku Code', 'Seller SKU', 'SKU'])
        cost_price_col = find_col(['Cost Price', 'Product Cost', 'Unit Cost', 'COGS'])
        sale_col = find_col(['Channel Forward Sale Amount', 'Channel Net Sale Amount', 'Sale Amount (Rs.)', 'Sale Amount', 'UC Selling Price'])
        refund_col = find_col(['Channel Refund Amount', 'Refund (Rs.)', 'Refund Amount', 'UC Refund'])
        bank_col = find_col(['Actual Net Settlement', 'Bank Settlement Value', 'Net Settlement'])
        return_type_col = find_col(['Return Type'])
        rev_ship_col = find_col(['Actual Reverse Shipping Fee', 'Reverse Shipping Fee (Rs.)', 'Reverse Shipping'])
        comm_col = find_col(['Actual Commission Fee', 'Commission (Rs.)'])
        gst_mp_col = find_col(['Actual Taxed Fee', 'GST on MP Fees (Rs.)', 'GST on Marketplace'])
        order_status_col = find_col(['Order Status', 'Item Return Status'])

        is_multi_channel = bool(channel_col and df[channel_col].dropna().nunique() > 1)
        has_cost_price = bool(cost_price_col)

        # Style Definitions
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        metric_font = Font(name="Calibri", size=10, bold=True)
        regular_font = Font(name="Calibri", size=10)
        border_thin = Border(left=Side(style='thin', color='DDDDDD'),
                             right=Side(style='thin', color='DDDDDD'),
                             top=Side(style='thin', color='DDDDDD'),
                             bottom=Side(style='thin', color='DDDDDD'))
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        alert_red_fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
        profit_fill = PatternFill(start_color="D9EAD3", fill_type="solid")

        wb = openpyxl.Workbook()
        wb.remove(wb.active) # Remove default sheet

        # Helper: Extract summary metrics
        def extract_metrics(sub_df):
            tot = len(sub_df)
            if return_type_col:
                cr = len(sub_df[sub_df[return_type_col].astype(str).str.upper().isin(['CUSTOMER RETURN', 'CIR'])])
                rto = len(sub_df[sub_df[return_type_col].astype(str).str.upper().isin(['LOGISTICS RETURN', 'RTO'])])
            else:
                cr, rto = 0, 0

            deliv = tot - (cr + rto)
            sales = float(pd.to_numeric(sub_df[sale_col], errors='coerce').fillna(0).sum()) if sale_col else 0.0
            refunds = abs(float(pd.to_numeric(sub_df[refund_col], errors='coerce').fillna(0).sum())) if refund_col else 0.0
            payout = float(pd.to_numeric(sub_df[bank_col], errors='coerce').fillna(0).sum()) if bank_col else 0.0
            rev_ship = abs(float(pd.to_numeric(sub_df[rev_ship_col], errors='coerce').fillna(0).sum())) if rev_ship_col else 0.0
            comm = abs(float(pd.to_numeric(sub_df[comm_col], errors='coerce').fillna(0).sum())) if comm_col else 0.0
            gst = abs(float(pd.to_numeric(sub_df[gst_mp_col], errors='coerce').fillna(0).sum())) if gst_mp_col else 0.0
            avg_cost = round(float(pd.to_numeric(sub_df[cost_price_col], errors='coerce').fillna(0).mean()), 2) if has_cost_price else 0.0

            return {
                "total": tot, "delivered": deliv, "cr": cr, "rto": rto,
                "sales": sales, "refunds": refunds, "payout": payout,
                "rev_ship": rev_ship, "comm": comm, "gst": gst, "avg_cost": avg_cost
            }

        # Helper: Create Summary Sheet with dynamic Excel formulas
        def make_summary_sheet(ws, title, m):
            ws.title = title[:31]
            ws.views.sheetView[0].showGridLines = True
            ws.append(["Metric", "Value"])
            ws["A1"].fill = header_fill
            ws["A1"].font = header_font
            ws["B1"].fill = header_fill
            ws["B1"].font = header_font

            rows = [
                ("Settlement File Total Transactions", m["total"], "0"),
                ("Delivered / Successful Orders", "=B2-B4-B6", "0"),
                ("Customer Return (CIR) Count", m["cr"], "0"),
                ("Customer Return %", "=B4/B2", "0.00%"),
                ("RTO (Logistics Return) Count", m["rto"], "0"),
                ("RTO (Logistics Return) %", "=B6/B2", "0.00%"),
                ("Gross Customer Sales Amount (₹)", m["sales"], "₹#,##0.00"),
                ("Customer Refund Amount (₹)", m["refunds"], "₹#,##0.00"),
                ("Reverse Shipping Fees Deducted (₹)", m["rev_ship"], "₹#,##0.00"),
                ("Marketplace Commission Fees (₹)", m["comm"], "₹#,##0.00"),
                ("GST Deducted on MP Fees (₹)", m["gst"], "₹#,##0.00"),
                ("Net Bank Settlement Received (₹)", m["payout"], "₹#,##0.00"),
                ("Average Unit Cost (₹) [FILL IF 0]", m["avg_cost"], "₹#,##0.00"),
                ("Total Product Cost (Sold Units) (₹)", "=B3*B14", "₹#,##0.00"),
                ("Net Profit in Hand (₹)", "=B13-B15", "₹#,##0.00"),
                ("Average Return Charge per Return Order (₹)", "=IF((B4+B6)>0, B10/(B4+B6), 0)", "₹#,##0.00"),
                ("Return Charge Burden per 1 Delivered Order (₹)", "=IF(B3>0, (B10+B12)/B3, 0)", "₹#,##0.00")
            ]

            for row_idx, (label, val, num_fmt) in enumerate(rows, start=2):
                ca = ws[f"A{row_idx}"]
                cb = ws[f"B{row_idx}"]
                ca.value = label
                cb.value = val
                cb.number_format = num_fmt
                ca.font = metric_font if row_idx in [3, 5, 7, 13, 16] else regular_font
                cb.font = metric_font if row_idx in [3, 5, 7, 13, 16] else regular_font
                ca.border = border_thin
                cb.border = border_thin
                if row_idx == 14 and (not has_cost_price or m["avg_cost"] == 0):
                    ca.fill = yellow_fill
                    cb.fill = yellow_fill
                    cb.comment = Comment("Please enter your average product purchase/manufacturing cost here.", "Payment Reco")
                if row_idx == 16:
                    ca.fill = profit_fill
                    cb.fill = profit_fill

            ws.column_dimensions["A"].width = 50
            ws.column_dimensions["B"].width = 24

        # Helper: Create SKU Sheet with '=B2-D2-E2' Delivered Formula
        def make_sku_sheet(ws, title, sub_df, summary_sheet_name):
            ws.title = title[:31]
            ws.views.sheetView[0].showGridLines = True
            headers = [
                "SKU Code", "Total Orders", "Delivered Orders", 
                "Customer Returns", "RTO Returns", "Return %", 
                "Unit Cost (₹)", "Total Payout (₹)", "Avg Payout / Delivered (₹)", 
                "Estimated Profit (₹)"
            ]
            ws.append(headers)
            for c_i in range(1, len(headers) + 1):
                c = ws.cell(row=1, column=c_i)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center", vertical="center")

            sku_list = []
            for sku_name, grp in sub_df.groupby(sku_col):
                sku_tot = len(grp)
                if return_type_col:
                    sku_cr = len(grp[grp[return_type_col].astype(str).str.upper().isin(['CUSTOMER RETURN', 'CIR'])])
                    sku_rto = len(grp[grp[return_type_col].astype(str).str.upper().isin(['LOGISTICS RETURN', 'RTO'])])
                else:
                    sku_cr, sku_rto = 0, 0

                sku_bank = float(pd.to_numeric(grp[bank_col], errors='coerce').fillna(0).sum()) if bank_col else 0.0
                sku_cost = float(pd.to_numeric(grp[cost_price_col], errors='coerce').fillna(0).mean()) if has_cost_price else 0.0

                sku_list.append({
                    "sku": str(sku_name),
                    "total": int(sku_tot),
                    "cr": int(sku_cr),
                    "rto": int(sku_rto),
                    "cost": round(sku_cost, 2),
                    "payout": sku_bank
                })

            sku_list.sort(key=lambda x: x["total"], reverse=True)

            for r_idx, item in enumerate(sku_list, start=2):
                cost_cell_val = item["cost"] if has_cost_price and item["cost"] > 0 else f"='{summary_sheet_name}'!$B$14"
                ws.append([
                    item["sku"],
                    item["total"],
                    f"=B{r_idx}-D{r_idx}-E{r_idx}",               # Dynamic Delivered Orders formula: B2-D2-E2
                    item["cr"],
                    item["rto"],
                    f"=(D{r_idx}+E{r_idx})/B{r_idx}",             # Return % formula
                    cost_cell_val,
                    round(item["payout"], 2),                    # Total Payout
                    f"=IF(C{r_idx}>0, H{r_idx}/C{r_idx}, 0)",     # Avg Payout / Delivered
                    f"=H{r_idx}-(C{r_idx}*G{r_idx})"              # Estimated Profit
                ])
                ws[f"B{r_idx}"].number_format = "0"
                ws[f"C{r_idx}"].number_format = "0"
                ws[f"D{r_idx}"].number_format = "0"
                ws[f"E{r_idx}"].number_format = "0"
                ws[f"F{r_idx}"].number_format = "0.0%"
                ws[f"G{r_idx}"].number_format = "₹#,##0.00"
                ws[f"H{r_idx}"].number_format = "₹#,##0.00"
                ws[f"I{r_idx}"].number_format = "₹#,##0.00"
                ws[f"J{r_idx}"].number_format = "₹#,##0.00"

                if item["cost"] == 0 and item["payout"] > 0 and has_cost_price:
                    ws[f"G{r_idx}"].fill = alert_red_fill
                    ws[f"G{r_idx}"].comment = Comment("Payout received but Cost is 0! Please update cost price.", "Payment Reco")

                for col_num in range(1, 11):
                    ws.cell(row=r_idx, column=col_num).border = border_thin

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col[:100])
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

        # GENERATE WORKBOOK SHEETS ACCORDING TO TEMPLATE TYPE
        if is_multi_channel:
            # Multi-channel template (e.g. SAmple.xlsx)
            channels = df[channel_col].dropna().unique()
            for ch in channels:
                ch_df = df[df[channel_col] == ch].copy()
                ch_clean = str(ch).replace("/", "_").replace("\\", "_")
                
                # Excel 31 characters limit safe titles
                sum_sheet_title = f"{ch_clean}_Summary"[:31]
                sku_sheet_title = f"{ch_clean}_SKU"[:31]

                # Channel Summary Sheet
                ws_ch_sum = wb.create_sheet()
                ch_m = extract_metrics(ch_df)
                make_summary_sheet(ws_ch_sum, sum_sheet_title, ch_m)

                # Channel SKU Sheet
                if sku_col:
                    ws_ch_sku = wb.create_sheet()
                    make_sku_sheet(ws_ch_sku, sku_sheet_title, ch_df, sum_sheet_title)
        else:
            # Single Marketplace template (Standard Flipkart Settlement)
            ws_single_sum = wb.create_sheet()
            single_m = extract_metrics(df)
            make_summary_sheet(ws_single_sum, "Executive_Summary", single_m)

            if sku_col:
                ws_single_sku = wb.create_sheet()
                make_sku_sheet(ws_single_sku, "SKU_Wise_Performance", df, "Executive_Summary")

        # UI Summary Display
        st.success(f"✅ Processed successfully! Mode: {'Multi-Marketplace Channel Mode' if is_multi_channel else 'Single Marketplace Mode'}")
        
        all_metrics = extract_metrics(df)
        st.subheader("📌 File Overview")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Transactions", f"{all_metrics['total']:,}")
        m2.metric("Delivered Orders", f"{all_metrics['delivered']:,}", f"{(all_metrics['delivered']/all_metrics['total']*100):.1f}% Delivery Rate")
        m3.metric("Customer Return (CIR)", f"{all_metrics['cr']:,}", f"{(all_metrics['cr']/all_metrics['total']*100):.1f}%")
        m4.metric("Courier Return (RTO)", f"{all_metrics['rto']:,}", f"{(all_metrics['rto']/all_metrics['total']*100):.1f}%")

        if is_multi_channel:
            st.markdown("---")
            st.subheader("🏪 Marketplace Channels Detected (Each has its own dedicated tabs)")
            ch_list_ui = []
            for ch in df[channel_col].dropna().unique():
                m_ch = extract_metrics(df[df[channel_col] == ch])
                ch_list_ui.append({
                    "Channel Code": ch,
                    "Total Orders": m_ch["total"],
                    "Delivered (Calculated)": m_ch["delivered"],
                    "CIR": m_ch["cr"],
                    "RTO": m_ch["rto"],
                    "Net Payout (₹)": round(m_ch["payout"], 2)
                })
            st.dataframe(pd.DataFrame(ch_list_ui), use_container_width=True)

        # Download Button
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Channel-Wise Reconciliation Excel File",
            data=output.getvalue(),
            file_name="Multi_Marketplace_Channel_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")

else:
    st.info("👆 Please upload your marketplace payment settlement Excel or CSV file above to generate the report.")
