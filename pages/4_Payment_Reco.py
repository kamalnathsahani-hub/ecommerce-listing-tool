import streamlit as st
import pandas as pd
import io

# Page Config
st.set_page_config(page_title="Payment Reconciliation & Profit Analyzer", layout="wide")

st.title("💰 Marketplace Payment & Profit Reconciliation")
st.markdown("Flipkart ya kisi bhi marketplace ki settlement Excel file upload karke payment breakup, deductions, return charges aur net profit analyze karein.")

# Sidebar inputs for costing
st.sidebar.header("⚙️ Costing & Expenses")
cogs_input = st.sidebar.number_input("Average Product Cost / COGS (₹ per item)", min_value=0.0, value=120.0, step=5.0)
pkg_cost = st.sidebar.number_input("Packaging & Label Cost (₹ per item)", min_value=0.0, value=10.0, step=1.0)
total_unit_cost = cogs_input + pkg_cost

# File Uploader
uploaded_file = st.file_uploader("📂 Upload Settlement Excel File (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # Check sheet structure
        if 'Orders' in xls.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name='Orders', header=1)
            # Remove any sub-header row where Order ID is empty
            df = df.dropna(subset=['Order ID']).copy()
        else:
            df = pd.read_excel(uploaded_file)
            df = df.dropna(how='all')

        # Helper column finders (Safe matching for different marketplace naming)
        def find_col(possible_names):
            for col in df.columns:
                for name in possible_names:
                    if name.lower() in str(col).lower():
                        return col
            return None

        order_id_col = find_col(['Order ID'])
        sale_col = find_col(['Sale Amount (Rs.)', 'Sale Amount'])
        refund_col = find_col(['Refund (Rs.)', 'Refund Amount', 'Refund'])
        bank_col = find_col(['Bank Settlement Value', 'Settlement Value', 'Net Settlement'])
        return_type_col = find_col(['Return Type'])
        rev_ship_col = find_col(['Reverse Shipping Fee', 'Reverse Shipping'])
        gst_mp_col = find_col(['GST on MP Fees', 'GST on Marketplace'])
        mp_fee_col = find_col(['Marketplace Fee (Rs.)', 'Marketplace Fee', 'Total Marketplace Fee'])
        sku_col = find_col(['Seller SKU', 'SKU'])
        order_date_col = find_col(['Order Date'])

        # Basic counts & splits
        total_transactions = len(df)
        
        if refund_col:
            df_delivered = df[df[refund_col] == 0].copy()
            df_returns = df[df[refund_col] < 0].copy()
        else:
            df_delivered = df.copy()
            df_returns = pd.DataFrame()

        delivered_count = len(df_delivered)
        return_count = len(df_returns)

        # Return types breakdown
        cust_return_count = len(df[df[return_type_col] == 'Customer Return']) if return_type_col else 0
        rto_count = len(df[df[return_type_col] == 'Logistics Return']) if return_type_col else 0

        cust_return_pct = (cust_return_count / total_transactions * 100) if total_transactions > 0 else 0
        rto_pct = (rto_count / total_transactions * 100) if total_transactions > 0 else 0
        total_return_pct = (return_count / total_transactions * 100) if total_transactions > 0 else 0

        # Financial totals
        gross_sales = float(df[sale_col].sum()) if sale_col else 0.0
        customer_refunds = abs(float(df[refund_col].sum())) if refund_col else 0.0
        bank_payout = float(df[bank_col].sum()) if bank_col else 0.0
        total_rev_shipping = abs(float(df[rev_ship_col].sum())) if rev_ship_col else 0.0
        total_gst_mp = abs(float(df[gst_mp_col].sum())) if gst_mp_col else 0.0
        total_mp_fee = abs(float(df[mp_fee_col].sum())) if mp_fee_col else 0.0

        # Calculations
        total_cogs_delivered = delivered_count * total_unit_cost
        estimated_profit = bank_payout - total_cogs_delivered
        profit_margin = (estimated_profit / bank_payout * 100) if bank_payout > 0 else 0

        # Average return penalties
        avg_charge_per_return = (total_rev_shipping / return_count) if return_count > 0 else 0.0
        return_burden_per_delivered = (total_rev_shipping + total_gst_mp) / delivered_count if delivered_count > 0 else 0.0

        # --- Dashboard Metrics Display ---
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

        st.subheader("⚠️ Return Cost Impact (Reverse Shipping)")
        r1, r2, r3 = st.columns(3)
        r1.metric("Total Reverse Shipping Deducted", f"₹{total_rev_shipping:,.2f}")
        r2.metric("Avg Reverse Charge / Return", f"₹{avg_charge_per_return:.2f}")
        r3.metric("Return Charge Burden / Delivered Order", f"₹{return_burden_per_delivered:.2f}")

        st.markdown("---")

        # --- Create Simplified Excel Workbook ---
        summary_records = [
            {"Metric": "Settlement File Total Transactions", "Value": total_transactions},
            {"Metric": "Delivered / Successful Orders", "Value": delivered_count},
            {"Metric": "Customer Return Count", "Value": cust_return_count},
            {"Metric": "Customer Return %", "Value": f"{cust_return_pct:.2f}%"},
            {"Metric": "RTO (Logistics Return) Count", "Value": rto_count},
            {"Metric": "RTO (Logistics Return) %", "Value": f"{rto_pct:.2f}%"},
            {"Metric": "Gross Customer Sales Amount (₹)", "Value": round(gross_sales, 2)},
            {"Metric": "Customer Refund Amount (₹)", "Value": round(customer_refunds, 2)},
            {"Metric": "Reverse Shipping Fees Deducted (₹)", "Value": round(total_rev_shipping, 2)},
            {"Metric": "Marketplace Commission & Other Fees (₹)", "Value": round(total_mp_fee - total_rev_shipping, 2)},
            {"Metric": "GST Deducted on MP Fees (₹)", "Value": round(total_gst_mp, 2)},
            {"Metric": "Net Bank Settlement Received (₹)", "Value": round(bank_payout, 2)},
            {"Metric": f"Total Product & Packaging Cost ({delivered_count} Sold Units) (₹)", "Value": round(total_cogs_delivered, 2)},
            {"Metric": "Net Profit in Hand (₹)", "Value": round(estimated_profit, 2)},
            {"Metric": "Average Return Charge per Return Order (₹)", "Value": round(avg_charge_per_return, 2)},
            {"Metric": "Return Charge Burden per 1 Delivered Order (₹)", "Value": round(return_burden_per_delivered, 2)}
        ]
        df_summary_report = pd.DataFrame(summary_records)

        # Simplified Clean Orders Dataframe
        clean_df = pd.DataFrame()
        if order_id_col: clean_df['Order ID'] = df[order_id_col]
        if order_date_col: clean_df['Order Date'] = df[order_date_col]
        if sku_col: clean_df['SKU'] = df[sku_col]
        if return_type_col: clean_df['Order Status'] = df[return_type_col].fillna('Delivered')
        if sale_col: clean_df['Sale Amount (₹)'] = df[sale_col]
        if refund_col: clean_df['Refund Amount (₹)'] = df[refund_col]
        if rev_ship_col: clean_df['Reverse Shipping (₹)'] = df[rev_ship_col]
        if bank_col: clean_df['Bank Settlement (₹)'] = df[bank_col]
        
        # Add Unit Cost & Profit column for delivered orders
        if refund_col and bank_col:
            clean_df['Unit Cost (₹)'] = clean_df['Refund Amount (₹)'].apply(lambda x: total_unit_cost if x == 0 else 0)
            clean_df['Estimated Order Profit (₹)'] = clean_df['Bank Settlement (₹)'] - clean_df['Unit Cost (₹)']

        # Preview Table
        st.subheader("📋 Executive Summary Table")
        st.dataframe(df_summary_report, use_container_width=True)

        # Excel Writer in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_summary_report.to_excel(writer, sheet_name='Executive_Summary', index=False)
            clean_df.to_excel(writer, sheet_name='Simplified_Orders_Reco', index=False)
        
        st.download_button(
            label="📥 Download Simplified Excel Report",
            data=output.getvalue(),
            file_name="Business_Payment_Profit_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")

else:
    st.info("👆 Please upload your marketplace payment settlement Excel file above to generate the report.")
