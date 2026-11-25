import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Daily Sale Report Generator", layout="wide")

st.title("📊 Daily Sale Report & Inventory Generator")
st.write("Sale Orders aur Inventory Database file upload karein aur channel-wise SKU sale + stock report download karein.")

# Channel name ko aapke report format ke hisaab se clean/map karne ke liye
CHANNEL_MAPPING = {
    "JBFASHION": "Lymio",
    "AJIO_Vrati_Fashion_PX": "Ajio PX",
    "JBFASHION_MYNTRA": "Myntra JB",
    "FLIPKART_POSHAX": "Flipkart",
    "AJIO_Leriya_fashion_AT": "Ajio AT",
    "SNAPDEAL_LF": "Snapdeal",
    "LERIYAFASHION": "Leriya Fashion",
    "YHUB": "Yhub",
    "MEESHOYHUB": "Meesho Yhub",
    "MEESHOLERIYA": "Meesho Leriya"
}

col1, col2 = st.columns(2)
with col1:
    orders_file = st.file_uploader("1. Sale Orders File Upload Karein (CSV / Excel)", type=["csv", "xlsx"])
with col2:
    inventory_file = st.file_uploader("2. Inventory Database File Upload Karein (CSV / Excel)", type=["csv", "xlsx"])

exclude_cancelled = st.checkbox("Cancelled orders ko report se hatayein (Recommended)", value=True)

def load_file(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file, low_memory=False)
    else:
        return pd.read_excel(uploaded_file)

if orders_file and inventory_file:
    try:
        df_orders = load_file(orders_file)
        df_inv = load_file(inventory_file)

        # Required columns check (K = Category, AQ = Item Type Name, AU = Channel Name)
        required_order_cols = ['Category', 'Item Type Name', 'Channel Name']
        missing_cols = [c for c in required_order_cols if c not in df_orders.columns]

        if missing_cols:
            st.error(f"Sale Orders file mein ye columns missing hain: {missing_cols}")
            st.stop()

        # Filter out cancelled orders if checked
        if exclude_cancelled and 'Sale Order Item Status' in df_orders.columns:
            df_orders = df_orders[df_orders['Sale Order Item Status'].str.upper() != 'CANCELLED'].copy()

        # Clean text columns
        df_orders['Category'] = df_orders['Category'].fillna('Uncategorized').astype(str).str.strip()
        df_orders['Item Type Name'] = df_orders['Item Type Name'].fillna('Unknown').astype(str).str.strip()
        df_orders['Channel Name'] = df_orders['Channel Name'].fillna('Other').astype(str).str.strip()

        # Map Channel Names to clean display names (agar naya channel aaye to original naam hi rakhega)
        df_orders['Clean_Channel'] = df_orders['Channel Name'].replace(CHANNEL_MAPPING)

        # Date suffix nikalne ke liye (e.g., 25-11)
        report_date_str = datetime.now().strftime("%d-%m")
        if 'Invoice Created' in df_orders.columns and df_orders['Invoice Created'].dropna().shape[0] > 0:
            try:
                sample_date = pd.to_datetime(df_orders['Invoice Created'].dropna().iloc[0])
                report_date_str = sample_date.strftime("%d-%m")
            except Exception:
                pass

        # Har row = 1 unit sale
        df_orders['Units'] = 1

        # Pivot Table: Product Name (AQ) & Parent/Category (K) vs Channel (AU)
        pivot_sku = pd.pivot_table(
            df_orders,
            values='Units',
            index=['Item Type Name', 'Category'],
            columns='Clean_Channel',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot_sku.rename(columns={
            'Item Type Name': 'Product Name',
            'Category': 'Parent'
        }, inplace=True)

        # Channel columns list
        channel_cols = [c for c in pivot_sku.columns if c not in ['Product Name', 'Parent']]
        
        # Preferred column order (baaki naye channels end mein add ho jayenge)
        preferred_order = ["Lymio", "Ajio PX", "Myntra JB", "Flipkart", "Ajio AT", "Snapdeal"]
        ordered_channels = [c for c in preferred_order if c in channel_cols] + [c for c in channel_cols if c not in preferred_order]
        
        pivot_sku = pivot_sku[['Product Name', 'Parent'] + ordered_channels]

        # Total Sold Units column add karein
        pivot_sku['Total Sale'] = pivot_sku[ordered_channels].sum(axis=1)

        # Inventory Database mapping (Item Name -> Available (ATP) & Total (Stock on hand))
        if 'Item Name' in df_inv.columns:
            df_inv['Item Name'] = df_inv['Item Name'].astype(str).str.strip()
            
            # Stock columns detect karein
            atp_col = 'Available (ATP)' if 'Available (ATP)' in df_inv.columns else None
            total_stock_col = 'Total (Stock on hand)' if 'Total (Stock on hand)' in df_inv.columns else None

            agg_dict = {}
            if atp_col:
                df_inv[atp_col] = pd.to_numeric(df_inv[atp_col], errors='coerce').fillna(0)
                agg_dict[atp_col] = 'sum'
            if total_stock_col:
                df_inv[total_stock_col] = pd.to_numeric(df_inv[total_stock_col], errors='coerce').fillna(0)
                agg_dict[total_stock_col] = 'sum'

            if agg_dict:
                inv_grouped = df_inv.groupby('Item Name', as_index=False).agg(agg_dict)
                final_df = pd.merge(
                    pivot_sku,
                    inv_grouped,
                    left_on='Product Name',
                    right_on='Item Name',
                    how='left'
                ).drop(columns=['Item Name'])

                # Jiska data database mein nahi hai use 0 consider karein
                if atp_col:
                    final_df[atp_col] = final_df[atp_col].fillna(0).astype(int)
                    final_df.rename(columns={atp_col: 'Available Stock'}, inplace=True)
                if total_stock_col:
                    final_df[total_stock_col] = final_df[total_stock_col].fillna(0).astype(int)
                    final_df.rename(columns={total_stock_col: 'Total Stock'}, inplace=True)
            else:
                st.warning("Inventory file mein 'Available (ATP)' ya 'Total (Stock on hand)' column nahi mila.")
                final_df = pivot_sku
        else:
            st.error("Inventory file mein 'Item Name' column nahi mila!")
            st.stop()

        # Sort by Parent (Category) and Product Name
        final_df = final_df.sort_values(by=['Parent', 'Product Name']).reset_index(drop=True)

        # Sheet 2: Category-wise Summary (Aapki Sheet2 ki tarah)
        numeric_cols = [c for c in final_df.columns if c not in ['Product Name', 'Parent']]
        category_summary = final_df.groupby('Parent', as_index=False)[numeric_cols].sum()
        category_summary.rename(columns={'Parent': 'Row Labels'}, inplace=True)

        # Grand Total row for Category Summary
        grand_total_cat = {'Row Labels': 'Grand Total'}
        for col in numeric_cols:
            grand_total_cat[col] = category_summary[col].sum()
        category_summary = pd.concat([category_summary, pd.DataFrame([grand_total_cat])], ignore_index=True)

        # Channel columns mein date add karein display ke liye (e.g., Lymio 25-11)
        rename_date_cols = {c: f"{c}\n{report_date_str}" for c in ordered_channels}
        final_df_export = final_df.rename(columns=rename_date_cols)
        category_summary_export = category_summary.rename(columns=rename_date_cols)

        # Top summary row for Sheet 1 (jaise aapki original file mein row 1 par totals hain)
        top_totals = {col: (final_df_export[col].sum() if col not in ['Product Name', 'Parent'] else '') for col in final_df_export.columns}
        top_totals['Product Name'] = 'GRAND TOTAL'

        # Metrics Show Karein
        st.success("Report successfully generate ho gayi hai!")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Unique SKUs Sold", len(final_df))
        m2.metric("Total Units Sold", int(final_df['Total Sale'].sum()))
        if 'Available Stock' in final_df.columns:
            m3.metric("Total Available Stock (Sold SKUs)", int(final_df['Available Stock'].sum()))

        st.subheader("Preview: SKU Wise Daily Sale & Stock Report")
        st.dataframe(final_df_export, use_container_width=True)

        # Excel File Create Karein
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet 1: SKU Wise Report (Top Total Row ke saath)
            pd.DataFrame([top_totals]).to_excel(writer, sheet_name='SKU_Wise_Report', index=False, header=False, startrow=0)
            final_df_export.to_excel(writer, sheet_name='SKU_Wise_Report', index=False, startrow=1)
            
            # Sheet 2: Category Summary
            category_summary_export.to_excel(writer, sheet_name='Category_Summary', index=False)

        output.seek(0)
        file_name = f"Daily Sale Report {report_date_str} With Inventory.xlsx"

        st.download_button(
            label="📥 Download Daily Sale Report (Excel)",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
