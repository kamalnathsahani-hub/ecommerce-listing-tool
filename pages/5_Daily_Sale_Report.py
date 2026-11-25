import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Daily Sale Report Generator", layout="wide")

st.title("📊 Daily Sale Report & Inventory Generator")
st.write("Sale Orders aur Inventory Database file upload karein aur channel-wise SKU sale + stock report download karein.")

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

col1, col2, col3 = st.columns([2, 2, 1.5])
with col1:
    orders_file = st.file_uploader("1. Sale Orders File Upload Karein (CSV / Excel)", type=["csv", "xlsx"])
with col2:
    inventory_file = st.file_uploader("2. Inventory Database File Upload Karein (CSV / Excel)", type=["csv", "xlsx"])
with col3:
    min_days_cover = st.number_input("Safe Stock Days (Green ke liye)", min_value=1, max_value=30, value=3, help="Agar stock (Daily Sale × Days) se kam hai to Red/Low dikhega")

exclude_cancelled = st.checkbox("Cancelled orders ko report se hatayein (Recommended)", value=True)

def load_file(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file, low_memory=False)
    else:
        return pd.read_excel(uploaded_file)

def format_excel_sheet(ws, header_row_idx, status_col_name="Stock Status"):
    # Colors & Styles
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006", bold=True)
    
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    yellow_font = Font(color="9C6500", bold=True)
    
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(color="006100", bold=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Header Row Formatting (C1 jaisa Wrap Text + Center Alignment sabhi headers par)
    ws.row_dimensions[header_row_idx].height = 34
    status_col_idx = None

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = thin_border
        if str(cell.value) == status_col_name:
            status_col_idx = col_idx

    # Top Total Row formatting (agar header row 2 hai)
    if header_row_idx == 2:
        ws.row_dimensions[1].height = 22
        for col_idx in range(1, ws.max_column + 1):
            c = ws.cell(row=1, column=col_idx)
            c.font = Font(bold=True)
            c.fill = total_fill
            c.border = thin_border
            if col_idx > 2:
                c.alignment = Alignment(horizontal="center", vertical="center")

    # Data Rows Formatting + Red/Green Highlighting
    for row_idx in range(header_row_idx + 1, ws.max_row + 1):
        is_grand_total = str(ws.cell(row=row_idx, column=1).value) == "Grand Total"
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            if is_grand_total:
                cell.font = Font(bold=True)
                cell.fill = total_fill
            
            if col_idx > 2:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Highlight Stock Status Column
            if status_col_idx and col_idx == status_col_idx and not is_grand_total:
                val = str(cell.value)
                if "Low Stock" in val or "Out of Stock" in val:
                    cell.fill = red_fill
                    cell.font = red_font
                elif "Reorder Soon" in val:
                    cell.fill = yellow_fill
                    cell.font = yellow_font
                elif "Sufficient" in val:
                    cell.fill = green_fill
                    cell.font = green_font

    # Auto-adjust column widths cleanly
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(ws.cell(row=header_row_idx, column=col_idx).value or "")
        if col_idx == 1:
            ws.column_dimensions[col_letter].width = 28
        elif col_idx == 2:
            ws.column_dimensions[col_letter].width = 16
        elif "\n" in header_val:
            ws.column_dimensions[col_letter].width = 13
        elif status_col_idx and col_idx == status_col_idx:
            ws.column_dimensions[col_letter].width = 24
        else:
            ws.column_dimensions[col_letter].width = 15

if orders_file and inventory_file:
    try:
        df_orders = load_file(orders_file)
        df_inv = load_file(inventory_file)

        required_order_cols = ['Category', 'Item Type Name', 'Channel Name']
        missing_cols = [c for c in required_order_cols if c not in df_orders.columns]

        if missing_cols:
            st.error(f"Sale Orders file mein ye columns missing hain: {missing_cols}")
            st.stop()

        if exclude_cancelled and 'Sale Order Item Status' in df_orders.columns:
            df_orders = df_orders[df_orders['Sale Order Item Status'].str.upper() != 'CANCELLED'].copy()

        df_orders['Category'] = df_orders['Category'].fillna('Uncategorized').astype(str).str.strip()
        df_orders['Item Type Name'] = df_orders['Item Type Name'].fillna('Unknown').astype(str).str.strip()
        df_orders['Channel Name'] = df_orders['Channel Name'].fillna('Other').astype(str).str.strip()
        df_orders['Clean_Channel'] = df_orders['Channel Name'].replace(CHANNEL_MAPPING)

        report_date_str = datetime.now().strftime("%d-%m")
        if 'Invoice Created' in df_orders.columns and df_orders['Invoice Created'].dropna().shape[0] > 0:
            try:
                sample_date = pd.to_datetime(df_orders['Invoice Created'].dropna().iloc[0])
                report_date_str = sample_date.strftime("%d-%m")
            except Exception:
                pass

        df_orders['Units'] = 1

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

        channel_cols = [c for c in pivot_sku.columns if c not in ['Product Name', 'Parent']]
        preferred_order = ["Lymio", "Ajio PX", "Myntra JB", "Flipkart", "Ajio AT", "Snapdeal"]
        ordered_channels = [c for c in preferred_order if c in channel_cols] + [c for c in channel_cols if c not in preferred_order]
        
        pivot_sku = pivot_sku[['Product Name', 'Parent'] + ordered_channels]
        pivot_sku['Total Sale'] = pivot_sku[ordered_channels].sum(axis=1)

        if 'Item Name' in df_inv.columns:
            df_inv['Item Name'] = df_inv['Item Name'].astype(str).str.strip()
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

        # Stock vs Sale Status Column Add Karein
        stock_ref_col = 'Available Stock' if 'Available Stock' in final_df.columns else ('Total Stock' if 'Total Stock' in final_df.columns else None)
        if stock_ref_col:
            def get_stock_status(row):
                stock = row[stock_ref_col]
                sale = row['Total Sale']
                if stock <= 0:
                    return "Out of Stock (0)"
                elif stock < sale:
                    return f"Low Stock (<1 Day)"
                elif stock < (sale * min_days_cover):
                    return f"Reorder Soon ({round(stock/sale, 1)}d)"
                else:
                    return f"Sufficient ({round(stock/sale, 1)}d)"

            final_df['Stock Status'] = final_df.apply(get_stock_status, axis=1)

        final_df = final_df.sort_values(by=['Parent', 'Product Name']).reset_index(drop=True)

        # Sheet 2: Category-wise Summary
        numeric_cols = [c for c in final_df.columns if c not in ['Product Name', 'Parent', 'Stock Status']]
        category_summary = final_df.groupby('Parent', as_index=False)[numeric_cols].sum()
        category_summary.rename(columns={'Parent': 'Row Labels'}, inplace=True)

        if stock_ref_col:
            def get_cat_status(row):
                stock = row[stock_ref_col]
                sale = row['Total Sale']
                if sale == 0:
                    return "Sufficient"
                if stock <= 0:
                    return "Out of Stock (0)"
                elif stock < sale:
                    return "Low Stock (<1 Day)"
                elif stock < (sale * min_days_cover):
                    return f"Reorder Soon ({round(stock/sale, 1)}d)"
                else:
                    return f"Sufficient ({round(stock/sale, 1)}d)"

            category_summary['Stock Status'] = category_summary.apply(get_cat_status, axis=1)

        grand_total_cat = {'Row Labels': 'Grand Total'}
        for col in numeric_cols:
            grand_total_cat[col] = category_summary[col].sum()
        if 'Stock Status' in category_summary.columns:
            grand_total_cat['Stock Status'] = ''
        category_summary = pd.concat([category_summary, pd.DataFrame([grand_total_cat])], ignore_index=True)

        # Channel columns mein 2-line format (C1 jaisa: Channel Name \n Date)
        rename_date_cols = {c: f"{c}\n{report_date_str}" for c in ordered_channels}
        final_df_export = final_df.rename(columns=rename_date_cols)
        category_summary_export = category_summary.rename(columns=rename_date_cols)

        top_totals = {col: (final_df_export[col].sum() if col not in ['Product Name', 'Parent', 'Stock Status'] else '') for col in final_df_export.columns}
        top_totals['Product Name'] = 'GRAND TOTAL'

        st.success("Report successfully generate ho gayi hai!")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Unique SKUs Sold", len(final_df))
        m2.metric("Total Units Sold", int(final_df['Total Sale'].sum()))
        if 'Stock Status' in final_df.columns:
            low_count = final_df['Stock Status'].str.contains('Low Stock|Out of Stock').sum()
            good_count = final_df['Stock Status'].str.contains('Sufficient').sum()
            m3.metric("🔴 Low / Out of Stock SKUs", int(low_count))
            m4.metric("🟢 Sufficient Stock SKUs", int(good_count))

        # Streamlit Web Preview mein bhi Red/Green highlight
        def highlight_status_web(val):
            val_str = str(val)
            if "Low Stock" in val_str or "Out of Stock" in val_str:
                return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;'
            elif "Reorder Soon" in val_str:
                return 'background-color: #FFEB9C; color: #9C6500; font-weight: bold;'
            elif "Sufficient" in val_str:
                return 'background-color: #C6EFCE; color: #006100; font-weight: bold;'
            return ''

        st.subheader("Preview: SKU Wise Daily Sale & Stock Report")
        if 'Stock Status' in final_df_export.columns:
            st.dataframe(final_df_export.style.map(highlight_status_web, subset=['Stock Status']), use_container_width=True)
        else:
            st.dataframe(final_df_export, use_container_width=True)

        # Excel File Export with Proper Wrap Text & Color Formatting
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame([top_totals]).to_excel(writer, sheet_name='SKU_Wise_Report', index=False, header=False, startrow=0)
            final_df_export.to_excel(writer, sheet_name='SKU_Wise_Report', index=False, startrow=1)
            category_summary_export.to_excel(writer, sheet_name='Category_Summary', index=False)

            # Apply formatting to both sheets
            format_excel_sheet(writer.sheets['SKU_Wise_Report'], header_row_idx=2, status_col_name="Stock Status")
            format_excel_sheet(writer.sheets['Category_Summary'], header_row_idx=1, status_col_name="Stock Status")

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
