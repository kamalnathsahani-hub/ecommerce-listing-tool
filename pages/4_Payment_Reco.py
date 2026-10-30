import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Multi-Channel Payment Reco & Profit Analyzer", layout="wide")

st.title("💰 Multi-Channel Payment & Profit Reconciliation")
st.markdown("Settlement Excel ya Multi-Channel CSV file upload karein. System dynamic formulas, auto cost price detection aur SKU-wise report banakar dega.")

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
                    if str(name).lower() == str(col).lower() or name.lower() in str(col).lower():
                        return col
            return None

        # Detect columns
        channel_col = find_col(['Channel Code', 'Channel'])
        order_id_col = find_col(['Channel Order Code', 'Order ID', 'Order Code'])
        sku_col = find_col(['Sku Code', 'Seller SKU', 'SKU'])
        cost_price_col = find_col(['Cost Price', 'Product Cost', 'Unit Cost', 'COGS'])
        sale_col = find_col(['Channel Forward Sale Amount', 'Sale Amount (Rs.)', 'Sale Amount', 'UC Selling Price'])
        refund_col = find_col(['Channel Refund Amount', 'Refund (Rs.)', 'Refund Amount', 'UC Refund'])
        bank_col = find_col(['Actual Net Settlement', 'Bank Settlement Value', 'Net Settlement'])
        return_type_col = find_col(['Return Type'])
        rev_ship_col = find_col(['Actual Reverse Shipping Fee', 'Reverse Shipping Fee (Rs.)', 'Reverse Shipping'])
        comm_col = find_col(['Actual Commission Fee', 'Commission (Rs.)'])
        gst_mp_col = find_col(['Actual Taxed Fee', 'GST on MP Fees (Rs.)', 'GST on Marketplace'])
        order_status_col = find_col(['Order Status', 'Item Return Status'])

        is_multi_channel = bool(channel_col)
        has_cost_price = bool(cost_price_col)

        total_transactions = len(df)

        if return_type_col:
            cust_returns = len(df[df[return_type_col].astype(str).str.upper().isin(['CUSTOMER RETURN', 'CIR'])])
            rto_returns = len(df[df[return_type_col].astype(str).str.upper().isin(['LOGISTICS RETURN', 'RTO'])])
        else:
            cust_returns, rto_returns = 0, 0

        if order_status_col:
            delivered_df = df[df[order_status_col].astype(str).str.upper().isin(['DELIVERED', 'COMPLETE'])].copy()
            if len(delivered_df) == 0:
                delivered_df = df[df[refund_col] == 0].copy() if refund_col else df.copy()
        elif refund_col:
            delivered_df = df[df[refund_col] == 0].copy()
        else:
            delivered_df = df.copy()

        delivered_count = len(delivered_df)
        total_returns = cust_returns + rto_returns

        gross_sales = float(pd.to_numeric(df[sale_col], errors='coerce').fillna(0).sum()) if sale_col else 0.0
        customer_refunds = abs(float(pd.to_numeric(df[refund_col], errors='coerce').fillna(0).sum())) if refund_col else 0.0
        bank_payout = float(pd.to_numeric(df[bank_col], errors='coerce').fillna(0).sum()) if bank_col else 0.0
        total_rev_shipping = abs(float(pd.to_numeric(df[rev_ship_col], errors='coerce').fillna(0).sum())) if rev_ship_col else 0.0
        total_gst_mp = abs(float(pd.to_numeric(df[gst_mp_col], errors='coerce').fillna(0).sum())) if gst_mp_col else 0.0
        total_comm = abs(float(pd.to_numeric(df[comm_col], errors='coerce').fillna(0).sum())) if comm_col else 0.0

        # Check for zero cost items with positive settlement
        zero_cost_count = 0
        if has_cost_price and bank_col:
            settled_mask = pd.to_numeric(df[bank_col], errors='coerce').fillna(0) > 0
            cost_series = pd.to_numeric(df[cost_price_col], errors='coerce').fillna(0)
            zero_cost_items = df[settled_mask & (cost_series == 0)]
            zero_cost_count = len(zero_cost_items)

        # UI Overview Metrics
        st.success(f"✅ File Processed! Mode: {'Multi-Marketplace Channel File' if is_multi_channel else 'Single Marketplace File'}")
        
        if zero_cost_count > 0:
            st.warning(f"⚠️ **Alert:** {zero_cost_count} orders ka payment settle hua hai lekin unka 'Cost Price' 0 ya blank hai! Excel sheet me yeh rows red highlight aur comment ke saath di gayi hain.")

        st.subheader("📌 Order & Return Breakdown")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Transactions", f"{total_transactions:,}")
        m2.metric("Delivered Orders", f"{delivered_count:,}", f"{(delivered_count/total_transactions*100):.1f}% Delivery Rate")
        m3.metric("Customer Return (CIR)", f"{cust_returns:,}", f"{(cust_returns/total_transactions*100):.1f}%")
        m4.metric("Courier Return (RTO)", f"{rto_returns:,}", f"{(rto_returns/total_transactions*100):.1f}%")

        f1, f2, f3 = st.columns(3)
        f1.metric("Gross Sales Value", f"₹{gross_sales:,.2f}")
        f2.metric("Customer Refunds", f"-₹{customer_refunds:,.2f}")
        f3.metric("Net Bank Settlement", f"₹{bank_payout:,.2f}")

        st.markdown("---")

        # --- CHANNEL WISE BREAKDOWN ---
        channel_data_list = []
        if is_multi_channel:
            for ch_name, grp in df.groupby(channel_col):
                ch_tot = len(grp)
                if order_status_col:
                    ch_del = len(grp[grp[order_status_col].astype(str).str.upper().isin(['DELIVERED', 'COMPLETE'])])
                elif refund_col:
                    ch_del = len(grp[grp[refund_col] == 0])
                else:
                    ch_del = ch_tot

                if return_type_col:
                    ch_cr = len(grp[grp[return_type_col].astype(str).str.upper().isin(['CUSTOMER RETURN', 'CIR'])])
                    ch_rto = len(grp[grp[return_type_col].astype(str).str.upper().isin(['LOGISTICS RETURN', 'RTO'])])
                else:
                    ch_cr, ch_rto = 0, 0

                ch_payout = float(pd.to_numeric(grp[bank_col], errors='coerce').fillna(0).sum()) if bank_col else 0.0
                
                channel_data_list.append({
                    "channel": str(ch_name),
                    "total_orders": ch_tot,
                    "delivered": ch_del,
                    "customer_return": ch_cr,
                    "rto_return": ch_rto,
                    "payout": ch_payout
                })
            
            df_channels_preview = pd.DataFrame(channel_data_list).sort_values(by="total_orders", ascending=False)
            st.subheader("🏪 Marketplace Channel-Wise Performance")
            st.dataframe(df_channels_preview.rename(columns={
                "channel": "Channel Name", "total_orders": "Total Orders", "delivered": "Delivered",
                "customer_return": "CIR (Customer Return)", "rto_return": "RTO", "payout": "Net Payout (₹)"
            }), use_container_width=True)
            st.markdown("---")

        # --- SKU WISE PIVOT CALCULATION ---
        sku_df = pd.DataFrame()
        if sku_col:
            sku_summary = []
            for sku_name, group in df.groupby(sku_col):
                sku_tot = len(group)
                if order_status_col:
                    sku_del = len(group[group[order_status_col].astype(str).str.upper().isin(['DELIVERED', 'COMPLETE'])])
                elif refund_col:
                    sku_del = len(group[group[refund_col] == 0])
                else:
                    sku_del = sku_tot

                if return_type_col:
                    sku_cr = len(group[group[return_type_col].astype(str).str.upper().isin(['CUSTOMER RETURN', 'CIR'])])
                    sku_rto = len(group[group[return_type_col].astype(str).str.upper().isin(['LOGISTICS RETURN', 'RTO'])])
                else:
                    sku_cr, sku_rto = 0, 0

                sku_bank = float(pd.to_numeric(group[bank_col], errors='coerce').fillna(0).sum()) if bank_col else 0.0
                
                # Check SKU level avg cost if available
                sku_cost = 0.0
                if has_cost_price:
                    sku_cost = float(pd.to_numeric(group[cost_price_col], errors='coerce').fillna(0).mean())

                sku_summary.append({
                    "sku": str(sku_name),
                    "total_orders": int(sku_tot),
                    "delivered_orders": int(sku_del),
                    "customer_returns": int(sku_cr),
                    "rto_returns": int(sku_rto),
                    "cost_price": round(sku_cost, 2),
                    "payout": float(sku_bank)
                })
            sku_df = pd.DataFrame(sku_summary).sort_values(by="total_orders", ascending=False)
            
            st.subheader("📦 Top 10 SKU-Wise Performance Preview")
            st.dataframe(sku_df.head(10).rename(columns={
                "sku": "SKU", "total_orders": "Total Orders", "delivered_orders": "Delivered Orders",
                "customer_returns": "Customer Returns", "rto_returns": "RTO Returns", "cost_price": "Avg Cost Price (₹)", "payout": "Total Bank Payout (₹)"
            }), use_container_width=True)

        # ====================================================
        # BUILD EXCEL WORKBOOK WITH DYNAMIC FORMULAS
        # ====================================================
        wb = openpyxl.Workbook()
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        metric_font = Font(name="Calibri", size=10, bold=True)
        regular_font = Font(name="Calibri", size=10)
        border_thin = Border(left=Side(style='thin', color='DDDDDD'),
                             right=Side(style='thin', color='DDDDDD'),
                             top=Side(style='thin', color='DDDDDD'),
                             bottom=Side(style='thin', color='DDDDDD'))

        # Alert styling
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Manual fill reminder
        alert_red_fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid") # Zero cost warning

        # ----------------------------------------------------
        # SHEET 1: Executive_Summary
        # ----------------------------------------------------
        ws1 = wb.active
        ws1.title = "Executive_Summary"
        ws1.views.sheetView[0].showGridLines = True

        ws1.append(["Metric", "Value"])
        ws1["A1"].fill = header_fill
        ws1["A1"].font = header_font
        ws1["B1"].fill = header_fill
        ws1["B1"].font = header_font

        # Calculate average cost from file if present, else 0
        default_cost_val = round(float(pd.to_numeric(df[cost_price_col], errors='coerce').fillna(0).mean()), 2) if has_cost_price else 0.0

        summary_rows = [
            ("Settlement File Total Transactions", total_transactions, "0"),
            ("Delivered / Successful Orders", delivered_count, "0"),
            ("Customer Return (CIR) Count", cust_returns, "0"),
            ("Customer Return %", "=B4/B2", "0.00%"),
            ("RTO (Logistics Return) Count", rto_returns, "0"),
            ("RTO (Logistics Return) %", "=B6/B2", "0.00%"),
            ("Gross Customer Sales Amount (₹)", gross_sales, "₹#,##0.00"),
            ("Customer Refund Amount (₹)", customer_refunds, "₹#,##0.00"),
            ("Reverse Shipping Fees Deducted (₹)", total_rev_shipping, "₹#,##0.00"),
            ("Marketplace Commission Fees (₹)", total_comm, "₹#,##0.00"),
            ("GST Deducted on MP Fees (₹)", total_gst_mp, "₹#,##0.00"),
            ("Net Bank Settlement Received (₹)", bank_payout, "₹#,##0.00"),
            ("Average Unit Cost (COGS + Pack) (₹) [FILL IF 0]", default_cost_val, "₹#,##0.00"),
            ("Total Product & Packaging Cost (Sold Units) (₹)", "=B3*B14", "₹#,##0.00"),
            ("Net Profit in Hand (₹)", "=B13-B15", "₹#,##0.00"),
            ("Average Return Charge per Return Order (₹)", "=IF((B4+B6)>0, B10/(B4+B6), 0)", "₹#,##0.00"),
            ("Return Charge Burden per 1 Delivered Order (₹)", "=IF(B3>0, (B10+B12)/B3, 0)", "₹#,##0.00")
        ]

        for row_idx, (label, val, num_fmt) in enumerate(summary_rows, start=2):
            cell_a = ws1[f"A{row_idx}"]
            cell_b = ws1[f"B{row_idx}"]
            cell_a.value = label
            cell_b.value = val
            cell_b.number_format = num_fmt
            cell_a.font = metric_font if row_idx in [5, 7, 13, 16] else regular_font
            cell_b.font = metric_font if row_idx in [5, 7, 13, 16] else regular_font
            cell_a.border = border_thin
            cell_b.border = border_thin

            # Highlight Unit Cost cell in yellow if not detected from file
            if row_idx == 14:
                if not has_cost_price or default_cost_val == 0:
                    cell_a.fill = yellow_fill
                    cell_b.fill = yellow_fill
                    cell_b.comment = Comment("Please enter your average product purchase/manufacturing cost here.", "Payment Reco")

            if row_idx == 16:
                cell_a.fill = PatternFill(start_color="D9EAD3", fill_type="solid")
                cell_b.fill = PatternFill(start_color="D9EAD3", fill_type="solid")

        ws1.column_dimensions["A"].width = 50
        ws1.column_dimensions["B"].width = 24

        # ----------------------------------------------------
        # SHEET 2: Channel_Wise_Summary (IF MULTI-CHANNEL)
        # ----------------------------------------------------
        if channel_data_list:
            ws_ch = wb.create_sheet(title="Channel_Wise_Summary")
            ws_ch.views.sheetView[0].showGridLines = True
            
            ch_headers = [
                "Channel Name", "Total Orders", "Delivered Orders", 
                "CIR Returns", "RTO Returns", "Total Return %", 
                "Net Payout (₹)", "Avg Payout / Delivered (₹)", 
                "Estimated Profit (₹)"
            ]
            ws_ch.append(ch_headers)
            for c_i in range(1, len(ch_headers) + 1):
                c = ws_ch.cell(row=1, column=c_i)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center", vertical="center")

            for r_idx, item in enumerate(channel_data_list, start=2):
                ws_ch.append([
                    item["channel"],
                    item["total_orders"],
                    item["delivered"],
                    item["customer_return"],
                    item["rto_return"],
                    f"=(D{r_idx}+E{r_idx})/B{r_idx}",
                    round(item["payout"], 2),
                    f"=IF(C{r_idx}>0, G{r_idx}/C{r_idx}, 0)",
                    f"=G{r_idx}-(C{r_idx}*Executive_Summary!$B$14)"
                ])
                ws_ch[f"B{r_idx}"].number_format = "0"
                ws_ch[f"C{r_idx}"].number_format = "0"
                ws_ch[f"D{r_idx}"].number_format = "0"
                ws_ch[f"E{r_idx}"].number_format = "0"
                ws_ch[f"F{r_idx}"].number_format = "0.0%"
                ws_ch[f"G{r_idx}"].number_format = "₹#,##0.00"
                ws_ch[f"H{r_idx}"].number_format = "₹#,##0.00"
                ws_ch[f"I{r_idx}"].number_format = "₹#,##0.00"
                for col_num in range(1, 10):
                    ws_ch.cell(row=r_idx, column=col_num).border = border_thin

            for col in ws_ch.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_ch.column_dimensions[col_letter].width = max(max_len + 3, 16)

        # ----------------------------------------------------
        # SHEET 3: SKU_Wise_Performance
        # ----------------------------------------------------
        if not sku_df.empty:
            ws2 = wb.create_sheet(title="SKU_Wise_Performance")
            ws2.views.sheetView[0].showGridLines = True
            
            sku_headers = [
                "SKU Code", "Total Orders", "Delivered Orders", 
                "Customer Returns", "RTO Returns", "Return %", 
                "Unit Cost (₹)", "Total Payout (₹)", "Avg Payout / Delivered (₹)", 
                "Estimated Profit (₹)"
            ]
            ws2.append(sku_headers)
            for col_num in range(1, len(sku_headers) + 1):
                c = ws2.cell(row=1, column=col_num)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center", vertical="center")

            for r_idx, item in enumerate(sku_df.to_dict('records'), start=2):
                cost_cell_val = item["cost_price"] if has_cost_price and item["cost_price"] > 0 else "=Executive_Summary!$B$14"
                ws2.append([
                    item["sku"],
                    item["total_orders"],
                    item["delivered_orders"],
                    item["customer_returns"],
                    item["rto_returns"],
                    f"=(D{r_idx}+E{r_idx})/B{r_idx}",
                    cost_cell_val,
                    round(item["payout"], 2),
                    f"=IF(C{r_idx}>0, H{r_idx}/C{r_idx}, 0)",
                    f"=H{r_idx}-(C{r_idx}*G{r_idx})"
                ])
                ws2[f"B{r_idx}"].number_format = "0"
                ws2[f"C{r_idx}"].number_format = "0"
                ws2[f"D{r_idx}"].number_format = "0"
                ws2[f"E{r_idx}"].number_format = "0"
                ws2[f"F{r_idx}"].number_format = "0.0%"
                ws2[f"G{r_idx}"].number_format = "₹#,##0.00"
                ws2[f"H{r_idx}"].number_format = "₹#,##0.00"
                ws2[f"I{r_idx}"].number_format = "₹#,##0.00"
                ws2[f"J{r_idx}"].number_format = "₹#,##0.00"
                
                # Highlight if cost is 0 but payout was received
                if item["cost_price"] == 0 and item["payout"] > 0 and has_cost_price:
                    ws2[f"G{r_idx}"].fill = alert_red_fill
                    ws2[f"G{r_idx}"].comment = Comment("Payout received but Cost is 0! Please update cost price.", "Payment Reco")

                for col_num in range(1, 11):
                    ws2.cell(row=r_idx, column=col_num).border = border_thin

            for col in ws2.columns:
                max_len = max(len(str(cell.value or '')) for cell in col[:100])
                col_letter = get_column_letter(col[0].column)
                ws2.column_dimensions[col_letter].width = max(max_len + 3, 14)

        # ----------------------------------------------------
        # SHEET 4: Simplified_Orders_Reco (ROW DATA WITH HIGHLIGHTS)
        # ----------------------------------------------------
        ws3 = wb.create_sheet(title="Simplified_Orders_Reco")
        ws3.views.sheetView[0].showGridLines = True
        
        order_headers = [
            "Order ID", "Channel", "SKU", "Order Status",
            "Sale Amount (₹)", "Refund Amount (₹)", "Reverse Shipping (₹)",
            "Bank Settlement (₹)", "Unit Cost (₹)", "Estimated Order Profit (₹)"
        ]
        ws3.append(order_headers)
        for col_num in range(1, len(order_headers) + 1):
            c = ws3.cell(row=1, column=col_num)
            c.fill = header_fill
            c.font = header_font

        clean_df = pd.DataFrame()
        clean_df['order_id'] = df[order_id_col] if order_id_col else ""
        clean_df['channel'] = df[channel_col] if channel_col else "Default"
        clean_df['sku'] = df[sku_col] if sku_col else ""
        clean_df['status'] = df[order_status_col] if order_status_col else (df[return_type_col].fillna('Delivered') if return_type_col else "Delivered")
        clean_df['sale'] = pd.to_numeric(df[sale_col], errors='coerce').fillna(0) if sale_col else 0
        clean_df['refund'] = pd.to_numeric(df[refund_col], errors='coerce').fillna(0) if refund_col else 0
        clean_df['rev_ship'] = pd.to_numeric(df[rev_ship_col], errors='coerce').fillna(0) if rev_ship_col else 0
        clean_df['bank'] = pd.to_numeric(df[bank_col], errors='coerce').fillna(0) if bank_col else 0
        
        # Pull cost price if exists, else None
        if has_cost_price:
            clean_df['cost'] = pd.to_numeric(df[cost_price_col], errors='coerce').fillna(0)
        else:
            clean_df['cost'] = None

        # Write clean orders
        for row_i, r in enumerate(clean_df.to_dict('records'), start=2):
            if r['cost'] is not None:
                cost_formula = r['cost']
            else:
                cost_formula = f"=IF(F{row_i}=0, Executive_Summary!$B$14, 0)"

            ws3.append([
                r['order_id'], r['channel'], r['sku'], r['status'],
                r['sale'], r['refund'], r['rev_ship'], r['bank'],
                cost_formula,
                f"=H{row_i}-I{row_i}"
            ])
            ws3[f"E{row_i}"].number_format = "0.00"
            ws3[f"F{row_i}"].number_format = "0.00"
            ws3[f"G{row_i}"].number_format = "0.00"
            ws3[f"H{row_i}"].number_format = "0.00"
            ws3[f"I{row_i}"].number_format = "0.00"
            ws3[f"J{row_i}"].number_format = "0.00"

            # Check if Cost is 0 while bank settlement > 0
            if r['cost'] is not None and r['cost'] == 0 and r['bank'] > 0:
                ws3[f"I{row_i}"].fill = alert_red_fill
                ws3[f"I{row_i}"].comment = Comment("Payout settled but Cost is 0! Update cost price.", "Payment Reco")
            elif not has_cost_price:
                # If no cost column in original file, highlight yellow for user entry
                ws3[f"I{row_i}"].fill = yellow_fill

        for col in ws3.columns:
            max_len = max(len(str(cell.value or '')) for cell in col[:100])
            col_letter = get_column_letter(col[0].column)
            ws3.column_dimensions[col_letter].width = max(max_len + 3, 14)

        # Output Download Button
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Simplified Multi-Marketplace Excel Report",
            data=output.getvalue(),
            file_name="Multi_Marketplace_Payment_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")

else:
    st.info("👆 Please upload your marketplace payment settlement Excel or CSV file above to generate the report.")
