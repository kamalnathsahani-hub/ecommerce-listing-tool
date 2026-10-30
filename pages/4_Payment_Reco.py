import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(page_title="Payment Reco & Profit Analyzer", layout="wide")

st.title("💰 Marketplace Payment & Profit Reconciliation")
st.markdown("Settlement Excel file upload karke sales, returns, deductions, live formulas aur **SKU-Wise Pivot** report download karein.")

# Sidebar Cost Inputs
st.sidebar.header("⚙️ Costing & Expenses")
cogs_input = st.sidebar.number_input("Average Product Cost / COGS (₹ per item)", min_value=0.0, value=120.0, step=5.0)
pkg_cost = st.sidebar.number_input("Packaging & Label Cost (₹ per item)", min_value=0.0, value=10.0, step=1.0)
total_unit_cost = float(cogs_input + pkg_cost)

# File Uploader
uploaded_file = st.file_uploader("📂 Upload Settlement Excel File (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # Check sheet structure
        if 'Orders' in xls.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name='Orders', header=1)
            df = df.dropna(subset=['Order ID']).copy()
        else:
            df = pd.read_excel(uploaded_file)
            df = df.dropna(how='all')

        # Helper column matchers
        def find_col(possible_names):
            for col in df.columns:
                for name in possible_names:
                    if name.lower() in str(col).lower():
                        return col
            return None

        order_id_col = find_col(['Order ID'])
        sale_col = find_col(['Sale Amount (Rs.)', 'Sale Amount'])
        refund_col = find_col(['Refund (Rs.)', 'Refund Amount', 'Refund'])
        bank_col = [c for c in df.columns if 'Bank Settlement' in str(c)][0] if any('Bank Settlement' in str(c) for c in df.columns) else find_col(['Settlement Value', 'Net Settlement'])
        return_type_col = find_col(['Return Type'])
        rev_ship_col = find_col(['Reverse Shipping Fee', 'Reverse Shipping'])
        gst_mp_col = find_col(['GST on MP Fees', 'GST on Marketplace'])
        mp_fee_col = find_col(['Marketplace Fee (Rs.)', 'Marketplace Fee', 'Total Marketplace Fee'])
        sku_col = find_col(['Seller SKU', 'SKU'])
        order_date_col = find_col(['Order Date'])

        # Data processing
        total_transactions = len(df)
        df_delivered = df[df[refund_col] == 0].copy() if refund_col else df.copy()
        df_returns = df[df[refund_col] < 0].copy() if refund_col else pd.DataFrame()

        delivered_count = len(df_delivered)
        return_count = len(df_returns)

        cust_return_count = len(df[df[return_type_col] == 'Customer Return']) if return_type_col else 0
        rto_count = len(df[df[return_type_col] == 'Logistics Return']) if return_type_col else 0

        cust_return_pct = (cust_return_count / total_transactions * 100) if total_transactions > 0 else 0
        rto_pct = (rto_count / total_transactions * 100) if total_transactions > 0 else 0

        gross_sales = float(df[sale_col].sum()) if sale_col else 0.0
        customer_refunds = abs(float(df[refund_col].sum())) if refund_col else 0.0
        bank_payout = float(df[bank_col].sum()) if bank_col else 0.0
        total_rev_shipping = abs(float(df[rev_ship_col].sum())) if rev_ship_col else 0.0
        total_gst_mp = abs(float(df[gst_mp_col].sum())) if gst_mp_col else 0.0
        total_mp_fee = abs(float(df[mp_fee_col].sum())) if mp_fee_col else 0.0

        total_cogs_delivered = delivered_count * total_unit_cost
        estimated_profit = bank_payout - total_cogs_delivered
        profit_margin = (estimated_profit / bank_payout * 100) if bank_payout > 0 else 0

        avg_charge_per_return = (total_rev_shipping / return_count) if return_count > 0 else 0.0
        return_burden_per_delivered = (total_rev_shipping + total_gst_mp) / delivered_count if delivered_count > 0 else 0.0

        # UI Metrics Display
        st.success("✅ Payment File Analyzed Successfully!")
        
        st.subheader("📌 Key Order & Return Breakdown")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Transactions", f"{total_transactions:,}")
        m2.metric("Delivered Orders", f"{delivered_count:,}", f"{(delivered_count/total_transactions*100):.1f}% Delivery Rate")
        m3.metric("Customer Return %", f"{cust_return_pct:.1f}%", f"{cust_return_count} Orders")
        m4.metric("RTO (Courier Return) %", f"{rto_pct:.1f}%", f"{rto_count} Orders")

        st.subheader("💵 Financials & Profitability")
        f1, f2, f3, f4 = st.columns(4)
        f1.metric("Gross Sales", f"₹{gross_sales:,.2f}")
        f2.metric("Customer Refunds", f"-₹{customer_refunds:,.2f}")
        f3.metric("Bank Payout (Received)", f"₹{bank_payout:,.2f}")
        f4.metric("Estimated Net Profit", f"₹{estimated_profit:,.2f}", f"{profit_margin:.1f}% Margin")

        st.markdown("---")

        # --- SKU WISE PIVOT CALCULATION ---
        sku_df = pd.DataFrame()
        if sku_col:
            sku_summary = []
            for sku_name, group in df.groupby(sku_col):
                sku_total_orders = len(group)
                sku_deliv = len(group[group[refund_col] == 0]) if refund_col else len(group)
                sku_cr = len(group[group[return_type_col] == 'Customer Return']) if return_type_col else 0
                sku_rto = len(group[group[return_type_col] == 'Logistics Return']) if return_type_col else 0
                sku_bank = group[bank_col].sum() if bank_col else 0.0
                sku_summary.append({
                    "sku": str(sku_name),
                    "total_orders": int(sku_total_orders),
                    "delivered_orders": int(sku_deliv),
                    "customer_returns": int(sku_cr),
                    "rto_returns": int(sku_rto),
                    "payout": float(sku_bank)
                })
            sku_df = pd.DataFrame(sku_summary).sort_values(by="total_orders", ascending=False)
            
            # Preview in UI
            preview_display = sku_df.rename(columns={
                "sku": "SKU",
                "total_orders": "Total Orders",
                "delivered_orders": "Delivered Orders",
                "customer_returns": "Customer Returns",
                "rto_returns": "RTO Returns",
                "payout": "Total Bank Payout (₹)"
            })
            st.subheader("📦 SKU-Wise Performance Preview (Top 10)")
            st.dataframe(preview_display.head(10), use_container_width=True)

        # --- CREATE EXCEL WORKBOOK VIA OPENPYXL ---
        wb = openpyxl.Workbook()
        
        # Styles
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        metric_font = Font(name="Calibri", size=10, bold=True)
        regular_font = Font(name="Calibri", size=10)
        border_thin = Border(left=Side(style='thin', color='DDDDDD'),
                             right=Side(style='thin', color='DDDDDD'),
                             top=Side(style='thin', color='DDDDDD'),
                             bottom=Side(style='thin', color='DDDDDD'))

        # ==========================================
        # SHEET 1: Executive_Summary (DYNAMIC FORMULAS)
        # ==========================================
        ws1 = wb.active
        ws1.title = "Executive_Summary"
        ws1.views.sheetView[0].showGridLines = True

        ws1.append(["Metric", "Value"])
        ws1["A1"].fill = header_fill
        ws1["A1"].font = header_font
        ws1["B1"].fill = header_fill
        ws1["B1"].font = header_font

        summary_rows = [
            ("Settlement File Total Transactions", total_transactions, "0"),
            ("Delivered / Successful Orders", delivered_count, "0"),
            ("Customer Return Count", cust_return_count, "0"),
            ("Customer Return %", "=B4/B2", "0.00%"),
            ("RTO (Logistics Return) Count", rto_count, "0"),
            ("RTO (Logistics Return) %", "=B6/B2", "0.00%"),
            ("Gross Customer Sales Amount (₹)", gross_sales, "₹#,##0.00"),
            ("Customer Refund Amount (₹)", customer_refunds, "₹#,##0.00"),
            ("Reverse Shipping Fees Deducted (₹)", total_rev_shipping, "₹#,##0.00"),
            ("Marketplace Commission & Other Fees (₹)", max(0.0, total_mp_fee - total_rev_shipping), "₹#,##0.00"),
            ("GST Deducted on MP Fees (₹)", total_gst_mp, "₹#,##0.00"),
            ("Net Bank Settlement Received (₹)", bank_payout, "₹#,##0.00"),
            ("Average Unit Cost (COGS + Pack) (₹)", total_unit_cost, "₹#,##0.00"),
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
            
            if row_idx == 16:
                cell_a.fill = PatternFill(start_color="D9EAD3", fill_type="solid")
                cell_b.fill = PatternFill(start_color="D9EAD3", fill_type="solid")

        ws1.column_dimensions["A"].width = 46
        ws1.column_dimensions["B"].width = 24

        # ==========================================
        # SHEET 2: SKU_Wise_Performance (WITH FORMULAS)
        # ==========================================
        if not sku_df.empty:
            ws2 = wb.create_sheet(title="SKU_Wise_Performance")
            ws2.views.sheetView[0].showGridLines = True
            
            sku_headers = [
                "SKU Code", "Total Orders", "Delivered Orders", 
                "Customer Returns", "RTO Returns", "Return %", 
                "Total Payout (₹)", "Avg Payout / Delivered (₹)", 
                "Estimated Profit (₹)"
            ]
            ws2.append(sku_headers)
            for col_num in range(1, len(sku_headers) + 1):
                c = ws2.cell(row=1, column=col_num)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center", vertical="center")

            # Safely iterating dictionary records to avoid itertuples name errors
            for r_idx, item in enumerate(sku_df.to_dict('records'), start=2):
                ws2.append([
                    item["sku"],
                    item["total_orders"],
                    item["delivered_orders"],
                    item["customer_returns"],
                    item["rto_returns"],
                    f"=(D{r_idx}+E{r_idx})/B{r_idx}",            # Return % formula
                    round(item["payout"], 2),                   # Total Payout
                    f"=IF(C{r_idx}>0, G{r_idx}/C{r_idx}, 0)",    # Avg Payout / Delivered
                    f"=G{r_idx}-(C{r_idx}*Executive_Summary!$B$14)" # Estimated Profit
                ])
                
                ws2[f"B{r_idx}"].number_format = "0"
                ws2[f"C{r_idx}"].number_format = "0"
                ws2[f"D{r_idx}"].number_format = "0"
                ws2[f"E{r_idx}"].number_format = "0"
                ws2[f"F{r_idx}"].number_format = "0.0%"
                ws2[f"G{r_idx}"].number_format = "₹#,##0.00"
                ws2[f"H{r_idx}"].number_format = "₹#,##0.00"
                ws2[f"I{r_idx}"].number_format = "₹#,##0.00"
                
                for col_num in range(1, 10):
                    ws2.cell(row=r_idx, column=col_num).border = border_thin

            for col in ws2.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws2.column_dimensions[col_letter].width = max(max_len + 3, 14)

        # ==========================================
        # SHEET 3: Simplified_Orders_Reco (ROW DATA)
        # ==========================================
        ws3 = wb.create_sheet(title="Simplified_Orders_Reco")
        ws3.views.sheetView[0].showGridLines = True
        
        order_headers = [
            "Order ID", "Order Date", "SKU", "Order Status",
            "Sale Amount (₹)", "Refund Amount (₹)", "Reverse Shipping (₹)",
            "Bank Settlement (₹)", "Unit Cost (₹)", "Estimated Order Profit (₹)"
        ]
        ws3.append(order_headers)
        for col_num in range(1, len(order_headers) + 1):
            c = ws3.cell(row=1, column=col_num)
            c.fill = header_fill
            c.font = header_font

        # Clean Order Level Data
        clean_df = pd.DataFrame()
        clean_df['order_id'] = df[order_id_col] if order_id_col else ""
        clean_df['order_date'] = df[order_date_col].astype(str) if order_date_col else ""
        clean_df['sku'] = df[sku_col] if sku_col else ""
        clean_df['status'] = df[return_type_col].fillna('Delivered') if return_type_col else "Delivered"
        clean_df['sale'] = pd.to_numeric(df[sale_col], errors='coerce').fillna(0) if sale_col else 0
        clean_df['refund'] = pd.to_numeric(df[refund_col], errors='coerce').fillna(0) if refund_col else 0
        clean_df['rev_ship'] = pd.to_numeric(df[rev_ship_col], errors='coerce').fillna(0) if rev_ship_col else 0
        clean_df['bank'] = pd.to_numeric(df[bank_col], errors='coerce').fillna(0) if bank_col else 0

        for row_i, r in enumerate(clean_df.to_dict('records'), start=2):
            ws3.append([
                r['order_id'], r['order_date'], r['sku'], r['status'],
                r['sale'], r['refund'], r['rev_ship'], r['bank'],
                f"=IF(F{row_i}=0, Executive_Summary!$B$14, 0)",
                f"=H{row_i}-I{row_i}"
            ])
            ws3[f"E{row_i}"].number_format = "0.00"
            ws3[f"F{row_i}"].number_format = "0.00"
            ws3[f"G{row_i}"].number_format = "0.00"
            ws3[f"H{row_i}"].number_format = "0.00"
            ws3[f"I{row_i}"].number_format = "0.00"
            ws3[f"J{row_i}"].number_format = "0.00"

        for col in ws3.columns:
            max_len = max(len(str(cell.value or '')) for cell in col[:100])
            col_letter = get_column_letter(col[0].column)
            ws3.column_dimensions[col_letter].width = max(max_len + 3, 14)

        # Output to Streamlit Download Button
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Simplified Excel Report (With Live Formulas & SKU Pivot)",
            data=output.getvalue(),
            file_name="Business_Payment_Profit_Report_Formulas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")

else:
    st.info("👆 Please upload your marketplace payment settlement Excel file above to generate the report.")
