import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Payment Reconciliation", layout="wide")

st.title("💰 Marketplace Payment & Profit Reconciliation")
st.write("Upload your Flipkart / Marketplace Payment Settlement Excel file to auto-summarize sales, returns, deductions & profits.")

uploaded_file = st.file_uploader("Upload Settlement Excel File", type=["xlsx", "xls"])

cogs_input = st.sidebar.number_input("Average Product Cost (₹ per item)", min_value=0.0, value=120.0, step=5.0)
pkg_cost = st.sidebar.number_input("Packaging Cost (₹ per item)", min_value=0.0, value=10.0, step=1.0)

if uploaded_file is not None:
    try:
        # Load sheets
        xls = pd.ExcelFile(uploaded_file)
        
        if 'Orders' in xls.sheet_names:
            df_orders = pd.read_excel(uploaded_file, sheet_name='Orders', header=1)
            # Clean duplicate header rows if any
            df_orders = df_orders.dropna(subset=['Order ID']).copy()
        else:
            df_orders = pd.read_excel(uploaded_file)

        # Metrics Extraction
        total_rows = len(df_orders)
        
        # Refunds & Delivered
        refund_col = 'Refund (Rs.)' if 'Refund (Rs.)' in df_orders.columns else None
        sale_col = 'Sale Amount (Rs.)' if 'Sale Amount (Rs.)' in df_orders.columns else None
        bank_col = [c for c in df_orders.columns if 'Bank Settlement' in str(c)][0]
        rev_ship_col = 'Reverse Shipping Fee (Rs.)' if 'Reverse Shipping Fee (Rs.)' in df_orders.columns else None
        gst_mp_col = 'GST on MP Fees (Rs.)' if 'GST on MP Fees (Rs.)' in df_orders.columns else None
        return_type_col = 'Return Type' if 'Return Type' in df_orders.columns else None
        
        delivered_df = df_orders[df_orders[refund_col] == 0] if refund_col else df_orders
        returns_df = df_orders[df_orders[refund_col] < 0] if refund_col else pd.DataFrame()
        
        delivered_count = len(delivered_df)
        returns_count = len(returns_df)
        
        cust_returns = len(df_orders[df_orders[return_type_col] == 'Customer Return']) if return_type_col else 0
        rto_returns = len(df_orders[df_orders[return_type_col] == 'Logistics Return']) if return_type_col else 0
        
        total_gross_sales = df_orders[sale_col].sum() if sale_col else 0
        total_refunds = df_orders[refund_col].sum() if refund_col else 0
        total_bank_settled = df_orders[bank_col].sum() if bank_col else 0
        total_rev_shipping = abs(df_orders[rev_ship_col].sum()) if rev_ship_col else 0
        total_gst_mp = abs(df_orders[gst_mp_col].sum()) if gst_mp_col else 0
        
        # Financial Calculations
        total_product_cogs = delivered_count * (cogs_input + pkg_cost)
        estimated_net_profit = total_bank_settled - total_product_cogs
        return_cost_per_delivered = (total_rev_shipping + total_gst_mp) / delivered_count if delivered_count > 0 else 0
        
        # Display KPIs
        st.subheader("📊 High-Level Performance Metrics")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Delivered Orders", f"{delivered_count:,}")
        kpi2.metric("Total Returns", f"{returns_count:,}", f"{(returns_count/total_rows*100):.1f}%")
        kpi3.metric("Customer Return vs RTO", f"{cust_returns} CR | {rto_returns} RTO")
        kpi4.metric("Bank Settlement", f"₹{total_bank_settled:,.2f}")
        
        kpi5, kpi6, kpi7, kpi8 = st.columns(4)
        kpi5.metric("Gross Customer Sales", f"₹{total_gross_sales:,.2f}")
        kpi6.metric("Return Charge / Sold Item", f"₹{return_cost_per_delivered:.2f}")
        kpi7.metric("Total COGS + Packaging", f"₹{total_product_cogs:,.2f}")
        kpi8.metric("Estimated Net Profit", f"₹{estimated_net_profit:,.2f}", 
                    delta=f"{(estimated_net_profit/total_bank_settled*100):.1f}% Margin")

        st.markdown("---")
        
        # Summary Table for Download
        summary_data = {
            "Metric": [
                "Delivered Orders (Success)",
                "Customer Returns",
                "Logistics Returns (RTO)",
                "Gross Sales Value (₹)",
                "Customer Refunds (₹)",
                "Reverse Shipping Charges (₹)",
                "GST on Marketplace Fees (₹)",
                "Net Bank Settlement Received (₹)",
                "Estimated Product Cost (₹)",
                "Net Estimated Profit (₹)",
                "Return Penalty Burden per Delivered Order (₹)"
            ],
            "Value": [
                delivered_count,
                cust_returns,
                rto_returns,
                round(total_gross_sales, 2),
                round(total_refunds, 2),
                round(total_rev_shipping, 2),
                round(total_gst_mp, 2),
                round(total_bank_settled, 2),
                round(total_product_cogs, 2),
                round(estimated_net_profit, 2),
                round(return_cost_per_delivered, 2)
            ]
        }
        df_summary_out = pd.DataFrame(summary_data)
        st.subheader("📋 Executive Settlement Summary")
        st.dataframe(df_summary_out, use_container_width=True)

        # Export to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_summary_out.to_excel(writer, sheet_name='Executive_Summary', index=False)
            df_orders.to_excel(writer, sheet_name='Raw_Orders_Data', index=False)
        
        st.download_button(
            label="📥 Download Simplified Excel Report",
            data=output.getvalue(),
            file_name="Simplified_Payment_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error reading settlement file: {str(e)}")
