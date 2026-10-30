import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Multi-Channel Payment Reco & Profit Analyzer", layout="wide")

st.title("💰 Multi-Channel Payment & Profit Reconciliation")
st.markdown("Settlement Excel ya Multi-Channel CSV file upload karein. File me agar `Channel Code` (Col E) hai, toh har channel ke liye **exact Apps Script accounting layout** aur SKU pivot report generate hogi.")

uploaded_file = st.file_uploader("📂 Upload Settlement / Reco File (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        file_name = uploaded_file.name
        base_name = file_name.rsplit('.', 1)[0]
        
        if file_name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file, low_memory=False)
        else:
            xls = pd.ExcelFile(uploaded_file)
            if 'Orders' in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name='Orders', header=1)
                df = df.dropna(subset=['Order ID']).copy()
            else:
                df = pd.read_excel(uploaded_file)
                df = df.dropna(how='all')

        # Strict exact & controlled matching
        def find_col(possible_names, exclude_keywords=None):
            for name in possible_names:
                for col in df.columns:
                    if str(name).strip().lower() == str(col).strip().lower():
                        return col
            for name in possible_names:
                for col in df.columns:
                    c_clean = str(col).strip().lower()
                    n_clean = str(name).strip().lower()
                    if n_clean in c_clean:
                        if exclude_keywords and any(ex.lower() in c_clean for ex in exclude_keywords):
                            continue
                        return col
            return None

        # Column identification
        channel_col = find_col(['Channel Code'], exclude_keywords=['order', 'item', 'product', 'quantity', 'sale', 'refund', 'date', 'time'])
        order_id_col = find_col(['Channel Order Code', 'Order ID', 'Order Code'])
        sku_col = find_col(['Sku Code', 'Seller SKU', 'SKU'])
        cost_price_col = find_col(['Cost Price', 'Product Cost', 'Unit Cost', 'COGS'])
        actual_net_col = find_col(['Actual Net Settlement', 'Bank Settlement Value', 'Net Settlement'])
        expected_net_col = find_col(['Expected Net Settlement'])
        return_type_col = find_col(['Return Type'])
        order_status_col = find_col(['Order Status', 'Item Return Status'])
        protection_fund_col = find_col(['Protection Fund', 'SPF'])
        forward_reco_col = find_col(['Forward Reco Status'])
        sale_col = find_col(['Channel Forward Sale Amount', 'Channel Net Sale Amount', 'Sale Amount (Rs.)', 'Sale Amount', 'UC Selling Price'])
        refund_col = find_col(['Channel Refund Amount', 'Refund (Rs.)', 'Refund Amount', 'UC Refund'])
        rev_ship_col = find_col(['Actual Reverse Shipping Fee', 'Reverse Shipping Fee (Rs.)', 'Reverse Shipping'])
        comm_col = find_col(['Actual Commission Fee', 'Commission (Rs.)'])
        gst_mp_col = find_col(['Actual Taxed Fee', 'GST on MP Fees (Rs.)', 'GST on Marketplace'])
        ad_fee_col = find_col(['Actual Marketing Fee', 'Ads Fees', 'Ad Services Fee'])

        is_multi_channel = bool(channel_col and df[channel_col].dropna().nunique() > 1)
        has_cost_price = bool(cost_price_col)

        # Style Definitions
        bold_header_font = Font(name="Calibri", size=11, bold=True)
        bold_metric_font = Font(name="Calibri", size=10, bold=True)
        regular_font = Font(name="Calibri", size=10)
        table_head_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        table_head_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        
        peach_fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
        green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        profit_fill = PatternFill(start_color="D9EAD3", fill_type="solid")
        border_thin = Border(left=Side(style='thin', color='E0E0E0'),
                             right=Side(style='thin', color='E0E0E0'),
                             top=Side(style='thin', color='E0E0E0'),
                             bottom=Side(style='thin', color='E0E0E0'))

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        def to_num(val):
            if pd.isna(val): return 0.0
            try: return float(val)
            except: return 0.0

        # Apps Script exact calculation logic
        def calculate_apps_script_metrics(sub_df):
            total_rows = len(sub_df)
            cancelled_count = 0
            os_cancelled_count = 0
            fr_cancelled_count = 0
            rto_count = 0
            cir_count = 0
            rto_net_settlement_sum = 0.0
            negative_net_settlement_sum = 0.0
            expected_net_settlement_sum = 0.0
            positive_case_count = 0

            none_actual_sum14 = 0.0
            none_expected_sum14 = 0.0
            zero_actual_count14 = 0

            none_cost_price_sum15 = 0.0
            zero_cost_price_count15 = 0

            pf_sum = 0.0
            pf_count = 0
            ad_fee_sum = 0.0

            for _, row in sub_df.iterrows():
                os_val = str(row.get(order_status_col, '')).strip().upper() if order_status_col else ''
                if os_val == 'CANCELLED':
                    cancelled_count += 1
                    os_cancelled_count += 1

                rtype = str(row.get(return_type_col, '')).strip().upper() if return_type_col else 'NONE'
                if rtype in ['RTO', 'LOGISTICS RETURN']: rto_count += 1
                if rtype in ['CIR', 'CUSTOMER RETURN']: cir_count += 1

                act_val = to_num(row.get(actual_net_col, 0)) if actual_net_col else 0.0
                exp_val = to_num(row.get(expected_net_col, 0)) if expected_net_col else 0.0
                cost_val = to_num(row.get(cost_price_col, 0)) if cost_price_col else 0.0
                pf_val = to_num(row.get(protection_fund_col, 0)) if protection_fund_col else 0.0
                fr_val = str(row.get(forward_reco_col, '')).strip().upper() if forward_reco_col else ''
                if ad_fee_col:
                    ad_fee_sum += to_num(row.get(ad_fee_col, 0))

                if rtype in ['RTO', 'LOGISTICS RETURN']:
                    rto_net_settlement_sum += act_val

                if act_val < 0:
                    negative_net_settlement_sum += act_val
                elif act_val > 0 and rtype in ['CIR', 'CUSTOMER RETURN']:
                    positive_case_count += 1
                    expected_net_settlement_sum += exp_val

                if rtype == 'NONE' and os_val != 'CANCELLED':
                    if act_val == 0:
                        if fr_val == 'CANCELLED':
                            cancelled_count += 1
                            fr_cancelled_count += 1
                        else:
                            zero_actual_count14 += 1
                            none_expected_sum14 += exp_val
                    else:
                        none_actual_sum14 += act_val

                    if cost_val == 0:
                        zero_cost_price_count15 += 1
                    none_cost_price_sum15 += cost_val

                if pf_val != 0:
                    pf_count += 1
                    pf_sum += pf_val

            return {
                "rowCount": total_rows,
                "cancelledCount": cancelled_count,
                "osCancelledCount": os_cancelled_count,
                "frCancelledCount": fr_cancelled_count,
                "rtoCount": rto_count,
                "cirCount": cir_count,
                "protectionFundSum": round(abs(pf_sum)),
                "protectionFundCount": pf_count,
                "rtoNetSettlementSum": round(abs(rto_net_settlement_sum)),
                "negativeNetSettlementSum": round(abs(negative_net_settlement_sum)),
                "expectedNetSettlementSum": round(abs(expected_net_settlement_sum)),
                "positiveCaseCount": positive_case_count,
                "noneActualSum14": round(abs(none_actual_sum14)),
                "noneExpectedSum14": round(abs(none_expected_sum14)),
                "zeroActualCount14": zero_actual_count14,
                "noneCostPriceSum15": round(abs(none_cost_price_sum15)),
                "zeroCostPriceCount15": zero_cost_price_count15,
                "adFeeSum": round(abs(ad_fee_sum))
            }

        # Build Sheet matching image layout
        def build_image_matched_summary_sheet(ws, title, m, channel_display_name):
            ws.title = title[:31]
            ws.views.sheetView[0].showGridLines = True

            ws["A1"] = f"{channel_display_name} {base_name}"
            ws["A1"].font = Font(name="Calibri", size=11, bold=True)
            ws["E1"] = "Other charges"
            ws["E1"].font = Font(name="Calibri", size=10, bold=True)
            ws["H1"] = "Other CN"
            ws["H1"].font = Font(name="Calibri", size=10, bold=True)

            rows = [
                (2, "Total Order", m["rowCount"], None, "0", None),
                (3, "Cancelled", m["cancelledCount"], None, "0", None),
                (4, "RTO", m["rtoCount"], None, "0", None),
                (5, "SPF Claim Amount", m["protectionFundSum"], None, "0", f"{m['protectionFundCount']} order(s) me Protection Fund ki 0 ke alava value mili hai."),
                (7, "C. Return", m["cirCount"], None, "0", None),
                (8, "RTO Charge", m["rtoNetSettlementSum"], None, "0", None),
                (9, "C. Return Charge", m["negativeNetSettlementSum"], m["expectedNetSettlementSum"], "0", f"{m['positiveCaseCount']} CIR order(s) ka value Expected Net Settlement se liya gaya hai, kyunki unme Actual Net Settlement positive tha."),
                (10, "Per order return charge", "=IFERROR(ROUND((B9+C9)/B7,0),0)", None, "0", None),
                (11, "Return %", '=ROUND(((B4+B7)/(B2-B3))*100,0)&"%"', None, "0%", None),
                (13, "Success order", "=B2-B3-B4-B7", None, "0", None),
                (14, "Success order payout", m["noneActualSum14"], m["noneExpectedSum14"], "0", f"{m['zeroActualCount14']} order(s) ka Actual Net Settlement 0 show kar raha tha, isliye unka Expected Net Settlement C14 me liya gaya hai."),
                (15, "Success Order product cost", m["noneCostPriceSum15"], None, "0", f"{m['zeroCostPriceCount15']} order ki Cost Price 0 mili hai."),
                (16, "Per order return charge on success order", "=IFERROR(ROUND((B9+C9)/B13,0),0)", None, "0", None),
                (18, "Gross Profit", "=ROUND(ABS(B5+B14+C14-B15),0)", None, "0", None),
                (19, "Less Return Charge", "=ROUND(ABS(B9+C9),0)", None, "0", None),
                (20, "Less Other Charges", "=SUM(F2:F20)", None, "0", None),
                (22, "Add other CN", "=SUM(I2:I20)", None, "0", None),
                (24, "Net Profit", "=ROUND(B18-B19-B20+B22,0)", None, "0", None),
                (26, "Profit per success Order", "=IFERROR(ROUND(B24/B13,2),0)", None, "0", None)
            ]

            ws["E3"] = "Ad Services Fee"
            ws["E3"].font = regular_font
            ws["F3"] = m["adFeeSum"]
            ws["F3"].font = regular_font
            ws["F3"].number_format = "0"
            ws["F3"].border = border_thin

            for r_idx, label, b_val, c_val, num_fmt, note_text in rows:
                ca = ws[f"A{r_idx}"]; cb = ws[f"B{r_idx}"]; cc = ws[f"C{r_idx}"]
                ca.value = label
                cb.value = b_val
                if c_val is not None:
                    cc.value = c_val
                    cc.number_format = num_fmt
                    cc.font = regular_font
                    cc.border = border_thin

                if num_fmt and b_val is not None:
                    cb.number_format = num_fmt

                ca.font = bold_metric_font if r_idx in [2, 13, 18, 24, 26] else regular_font
                cb.font = bold_metric_font if r_idx in [2, 13, 18, 24, 26] else regular_font

                ca.border = border_thin
                cb.border = border_thin

                if note_text:
                    target_cell = cc if c_val is not None else cb
                    target_cell.comment = Comment(note_text, "Payment Reco")

                if r_idx in [19, 20]:
                    cb.fill = peach_fill
                elif r_idx == 22:
                    cb.fill = green_fill

            ws.column_dimensions["A"].width = 38
            ws.column_dimensions["B"].width = 18
            ws.column_dimensions["C"].width = 18
            ws.column_dimensions["D"].width = 6
            ws.column_dimensions["E"].width = 24
            ws.column_dimensions["F"].width = 18
            ws.column_dimensions["G"].width = 6
            ws.column_dimensions["H"].width = 20
            ws.column_dimensions["I"].width = 16

        # Build SKU Sheet with dynamic delivered orders formula: B2-D2-E2
        def build_sku_sheet(ws, title, sub_df, summary_sheet_name):
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
                c.fill = table_head_fill; c.font = table_head_font
                c.alignment = Alignment(horizontal="center", vertical="center")

            sku_list = []
            for sku_name, grp in sub_df.groupby(sku_col):
                sku_tot = len(grp)
                if return_type_col:
                    sku_cr = len(grp[grp[return_type_col].astype(str).str.upper().isin(['CUSTOMER RETURN', 'CIR'])])
                    sku_rto = len(grp[grp[return_type_col].astype(str).str.upper().isin(['LOGISTICS RETURN', 'RTO'])])
                else:
                    sku_cr, sku_rto = 0, 0

                sku_bank = float(pd.to_numeric(grp[actual_net_col], errors='coerce').fillna(0).sum()) if actual_net_col else 0.0
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
                cost_cell_val = item["cost"] if has_cost_price and item["cost"] > 0 else f"='{summary_sheet_name}'!$B$15/'{summary_sheet_name}'!$B$13"
                ws.append([
                    item["sku"],
                    item["total"],
                    f"=B{r_idx}-D{r_idx}-E{r_idx}",               # Delivered Orders dynamic formula
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

                for col_num in range(1, 11):
                    ws.cell(row=r_idx, column=col_num).border = border_thin

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col[:100])
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

        # Standard Single Marketplace Summary
        def build_standard_single_summary(ws, title, sub_df):
            ws.title = title[:31]
            ws.views.sheetView[0].showGridLines = True
            ws.append(["Metric", "Value"])
            ws["A1"].fill = table_head_fill; ws["A1"].font = table_head_font
            ws["B1"].fill = table_head_fill; ws["B1"].font = table_head_font

            tot = len(sub_df)
            cr = len(sub_df[sub_df[return_type_col].astype(str).str.upper().isin(['CUSTOMER RETURN', 'CIR'])]) if return_type_col else 0
            rto = len(sub_df[sub_df[return_type_col].astype(str).str.upper().isin(['LOGISTICS RETURN', 'RTO'])]) if return_type_col else 0
            sales = float(pd.to_numeric(sub_df[sale_col], errors='coerce').fillna(0).sum()) if sale_col else 0.0
            refunds = abs(float(pd.to_numeric(sub_df[refund_col], errors='coerce').fillna(0).sum())) if refund_col else 0.0
            payout = float(pd.to_numeric(sub_df[actual_net_col], errors='coerce').fillna(0).sum()) if actual_net_col else 0.0
            rev_ship = abs(float(pd.to_numeric(sub_df[rev_ship_col], errors='coerce').fillna(0).sum())) if rev_ship_col else 0.0
            comm = abs(float(pd.to_numeric(sub_df[comm_col], errors='coerce').fillna(0).sum())) if comm_col else 0.0
            gst = abs(float(pd.to_numeric(sub_df[gst_mp_col], errors='coerce').fillna(0).sum())) if gst_mp_col else 0.0

            rows = [
                ("Settlement File Total Transactions", tot, "0"),
                ("Delivered / Successful Orders", "=B2-B4-B6", "0"),
                ("Customer Return (CIR) Count", cr, "0"),
                ("Customer Return %", "=B4/B2", "0.00%"),
                ("RTO (Logistics Return) Count", rto, "0"),
                ("RTO (Logistics Return) %", "=B6/B2", "0.00%"),
                ("Gross Customer Sales Amount (₹)", sales, "₹#,##0.00"),
                ("Customer Refund Amount (₹)", refunds, "₹#,##0.00"),
                ("Reverse Shipping Fees Deducted (₹)", rev_ship, "₹#,##0.00"),
                ("Marketplace Commission Fees (₹)", comm, "₹#,##0.00"),
                ("GST Deducted on MP Fees (₹)", gst, "₹#,##0.00"),
                ("Net Bank Settlement Received (₹)", payout, "₹#,##0.00"),
                ("Average Return Charge per Return Order (₹)", "=IF((B4+B6)>0, B10/(B4+B6), 0)", "₹#,##0.00"),
                ("Return Charge Burden per 1 Delivered Order (₹)", "=IF(B3>0, (B10+B12)/B3, 0)", "₹#,##0.00")
            ]

            for row_idx, (label, val, num_fmt) in enumerate(rows, start=2):
                ca = ws[f"A{row_idx}"]; cb = ws[f"B{row_idx}"]
                ca.value = label; cb.value = val; cb.number_format = num_fmt
                ca.font = bold_metric_font if row_idx in [3, 5, 7, 13] else regular_font
                cb.font = bold_metric_font if row_idx in [3, 5, 7, 13] else regular_font
                ca.border = border_thin; cb.border = border_thin

            ws.column_dimensions["A"].width = 50
            ws.column_dimensions["B"].width = 24

        # EXECUTION: CHECK TEMPLATE TYPE
        if is_multi_channel:
            channels = df[channel_col].dropna().unique()
            for ch in channels:
                ch_df = df[df[channel_col] == ch].copy()
                ch_clean = str(ch).replace("/", "_").replace("\\", "_").strip()
                
                sum_sheet_title = f"{ch_clean}_Summary"[:31]
                sku_sheet_title = f"{ch_clean}_SKU"[:31]

                ws_sum = wb.create_sheet()
                ch_m = calculate_apps_script_metrics(ch_df)
                build_image_matched_summary_sheet(ws_sum, sum_sheet_title, ch_m, ch_clean)

                if sku_col:
                    ws_sku = wb.create_sheet()
                    build_sku_sheet(ws_sku, sku_sheet_title, ch_df, sum_sheet_title)
        else:
            ws_single_sum = wb.create_sheet()
            build_standard_single_summary(ws_single_sum, "Executive_Summary", df)

            if sku_col:
                ws_single_sku = wb.create_sheet()
                build_sku_sheet(ws_single_sku, "SKU_Wise_Performance", df, "Executive_Summary")

        st.success(f"✅ Processed successfully! Mode: {'Multi-Marketplace Channel Mode' if is_multi_channel else 'Single Marketplace Mode'}")
        
        all_metrics = calculate_apps_script_metrics(df)
        st.subheader("📌 File Overview")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Transactions", f"{all_metrics['rowCount']:,}")
        m2.metric("Cancelled Orders", f"{all_metrics['cancelledCount']:,}")
        m3.metric("RTO Returns", f"{all_metrics['rtoCount']:,}")
        m4.metric("CIR Returns", f"{all_metrics['cirCount']:,}")

        if is_multi_channel:
            st.markdown("---")
            st.subheader("🏪 Marketplace Channels Detected (Column E)")
            ch_list_ui = []
            for ch in df[channel_col].dropna().unique():
                m_ch = calculate_apps_script_metrics(df[df[channel_col] == ch])
                ch_success = m_ch["rowCount"] - m_ch["cancelledCount"] - m_ch["rtoCount"] - m_ch["cirCount"]
                ch_list_ui.append({
                    "Channel Code": ch,
                    "Total Order (B2)": m_ch["rowCount"],
                    "Cancelled (B3)": m_ch["cancelledCount"],
                    "RTO (B4)": m_ch["rtoCount"],
                    "CIR (B7)": m_ch["cirCount"],
                    "Success Order (B13)": ch_success,
                    "Success Payout (B14)": f"₹{m_ch['noneActualSum14']:,}",
                    "Product Cost (B15)": f"₹{m_ch['noneCostPriceSum15']:,}"
                })
            st.dataframe(pd.DataFrame(ch_list_ui), use_container_width=True)

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
