import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Daily Sale Report Generator", layout="wide")

st.title("📊 Daily Sale Report & Inventory Generator")
st.write("Unicommerce ya OMS ki Sale Orders aur Inventory files (CSV, Excel, ya ZIP) upload karein aur channel-wise SKU sale + available stock report download karein.")

# Unicommerce Channel Mapping
UNICOM_CHANNEL_MAPPING = {
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

# OMS Channel Mapping
OMS_CHANNEL_MAPPING = {
    "URBAN POCKETS LIFESTYLE - Myntra PPMP": "Myntra PPMP",
    "URBAN POCKETS LIFESTYLE - Amazon Direct Fulfillment": "Amazon DF",
    "URBAN POCKETS LIFESTYLE - Ajio Api": "Ajio API"
}

SIZE_ORDER = {
    "XXS": 1, "2XS": 1,
    "XS": 2,
    "S": 3,
    "M": 4,
    "L": 5,
    "XL": 6,
    "XXL": 7, "2XL": 7,
    "XXXL": 8, "3XL": 8,
    "4XL": 9,
    "5XL": 10,
    "6XL": 11,
    "7XL": 12,
    "FREE": 90, "FREESIZE": 90, "FS": 90
}

def get_sort_keys(sku_name):
    s = str(sku_name).strip()
    parts = s.rsplit('-', 1)
    if len(parts) == 2:
        base_name, size_part = parts[0], parts[1].upper().strip()
        if size_part in SIZE_ORDER:
            return (base_name.lower(), SIZE_ORDER[size_part], size_part)
        if size_part.isdigit():
            return (base_name.lower(), 20 + int(size_part), size_part)
    return (s.lower(), 999, "")

def load_any_file(uploaded_file):
    """CSV, XLSX, XLS, aur ZIP (jisme andar CSV/Excel ho) sabhi ko read karta hai."""
    fname = uploaded_file.name.lower()
    if fname.endswith('.zip'):
        with zipfile.ZipFile(uploaded_file, 'r') as z:
            valid_files = [f for f in z.namelist() if not f.startswith('__MACOSX') and (f.lower().endswith('.csv') or f.lower().endswith('.xlsx') or f.lower().endswith('.xls'))]
            if not valid_files:
                raise ValueError(f"ZIP file '{uploaded_file.name}' ke andar koi CSV ya Excel file nahi mili.")
            inner_name = valid_files[0]
            with z.open(inner_name) as f:
                if inner_name.lower().endswith('.csv'):
                    return pd.read_csv(f, low_memory=False)
                else:
                    return pd.read_excel(f)
    elif fname.endswith('.csv'):
        return pd.read_csv(uploaded_file, low_memory=False)
    else:
        return pd.read_excel(uploaded_file)

def clean_oms_channel(val):
    s = str(val).strip()
    if s in OMS_CHANNEL_MAPPING:
        return OMS_CHANNEL_MAPPING[s]
    if " - " in s:
        return s.split(" - ", 1)[1].strip()
    return s

def format_excel_sheet(ws, header_row_idx, status_col_name="Stock Status"):
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

    if header_row_idx == 2:
        ws.row_dimensions[1].height = 22
        for col_idx in range(1, ws.max_column + 1):
            c = ws.cell(row=1, column=col_idx)
            c.font = Font(bold=True)
            c.fill = total_fill
            c.border = thin_border
            if col_idx > 2:
                c.alignment = Alignment(horizontal="center", vertical="center")

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

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(ws.cell(row=header_row_idx, column=col_idx).value or "")
        if col_idx == 1:
            ws.column_dimensions[col_letter].width = 30
        elif col_idx == 2:
            ws.column_dimensions[col_letter].width = 18
        elif "\n" in header_val:
            ws.column_dimensions[col_letter].width = 15
        elif status_col_idx and col_idx == status_col_idx:
            ws.column_dimensions[col_letter].width = 24
        else:
            ws.column_dimensions[col_letter].width = 16

col1, col2, col3 = st.columns([2, 2, 1.5])
with col1:
    orders_file = st.file_uploader("1. Sale Orders File Upload Karein (CSV / XLSX / ZIP)", type=["csv", "xlsx", "xls", "zip"])
with col2:
    inventory_file = st.file_uploader("2. Inventory Database File Upload Karein (CSV / XLSX / ZIP)", type=["csv", "xlsx", "xls", "zip"])
with col3:
    min_days_cover = st.number_input(
        "Safe Stock Days (Green ke liye)",
        min_value=1,
        max_value=30,
        value=3,
        help="Agar Available Stock (Daily Sale × Days) se kam hai to Red/Yellow dikhega"
    )

opt_col1, opt_col2 = st.columns([2, 2])
with opt_col1:
    sku_mode = st.radio(
        "Report kis SKU column ke hisab se banani hai?",
        options=["Primary Product SKU (Item Type Name / Product Sku Code)", "Channel / Seller SKU (Seller SKU Code / Listing Sku Code)"],
        horizontal=True
    )
with opt_col2:
    exclude_cancelled = st.checkbox("Cancelled orders ko report se hatayein (Recommended)", value=True)

if orders_file and inventory_file:
    try:
        df_orders = load_any_file(orders_file)
        df_inv = load_any_file(inventory_file)

        use_alt_sku = "Channel / Seller SKU" in sku_mode

        # Auto-Detect Format: OMS vs Unicommerce
        is_oms = ('Product Sku Code' in df_orders.columns and 'Category Name' in df_orders.columns)

        if is_oms:
            st.info("Detected Format: **OMS Order & Inventory Report**")
            target_col = "Listing Sku Code" if use_alt_sku else "Product Sku Code"
            display_col_name = "Listing Sku Code" if use_alt_sku else "Product Sku Code"
            cat_col = "Category Name"

            # Filter Cancelled in OMS
            if exclude_cancelled and 'Order Status' in df_orders.columns:
                cancel_statuses = ['CANCELLED', 'CANCEL INIT', 'CANCELLED RETURN RECEIVED']
                df_orders = df_orders[~df_orders['Order Status'].astype(str).str.upper().str.strip().isin(cancel_statuses)].copy()

            df_orders[cat_col] = df_orders[cat_col].fillna('Uncategorized').astype(str).str.strip()
            df_orders['Product Sku Code'] = df_orders['Product Sku Code'].fillna('Unknown').astype(str).str.strip()
            if use_alt_sku:
                df_orders['Listing Sku Code'] = df_orders['Listing Sku Code'].fillna('Unknown_Listing_SKU').astype(str).str.strip()

            df_orders['Clean_Channel'] = df_orders['Channel Name'].fillna('Other').apply(clean_oms_channel)

            # Quantity column in OMS
            if 'Qty' in df_orders.columns:
                df_orders['Units'] = pd.to_numeric(df_orders['Qty'], errors='coerce').fillna(1)
            else:
                df_orders['Units'] = 1

            # Extract Report Date
            report_date_str = datetime.now().strftime("%d-%m")
            for date_col in ['Invoice Date', 'Order Date']:
                if date_col in df_orders.columns and df_orders[date_col].dropna().shape[0] > 0:
                    try:
                        # Latest date lein kyunki OMS file mein multiple dates ho sakti hain
                        parsed_dates = pd.to_datetime(df_orders[date_col].dropna(), errors='coerce').dropna()
                        if not parsed_dates.empty:
                            report_date_str = parsed_dates.max().strftime("%d-%m")
                            break
                    except Exception:
                        pass

            # Pivot Table
            pivot_sku = pd.pivot_table(
                df_orders,
                values='Units',
                index=[target_col, cat_col],
                columns='Clean_Channel',
                aggfunc='sum',
                fill_value=0
            ).reset_index()

            pivot_sku.rename(columns={target_col: display_col_name, cat_col: 'Parent'}, inplace=True)

            channel_cols = [c for c in pivot_sku.columns if c not in [display_col_name, 'Parent']]
            preferred_order = ["Myntra PPMP", "Amazon DF", "Ajio API"]
            ordered_channels = [c for c in preferred_order if c in channel_cols] + [c for c in channel_cols if c not in preferred_order]

            pivot_sku = pivot_sku[[display_col_name, 'Parent'] + ordered_channels]
            pivot_sku['Total Sale'] = pivot_sku[ordered_channels].sum(axis=1)

            # Inventory Mapping for OMS
            if 'Item SkuCode' in df_inv.columns and 'Qty' in df_inv.columns:
                df_inv['Item SkuCode'] = df_inv['Item SkuCode'].astype(str).str.strip()
                df_inv['Qty'] = pd.to_numeric(df_inv['Qty'], errors='coerce').fillna(0)
                inv_grouped = df_inv.groupby('Item SkuCode', as_index=False)['Qty'].sum()

                if not use_alt_sku:
                    final_df = pd.merge(
                        pivot_sku,
                        inv_grouped,
                        left_on=display_col_name,
                        right_on='Item SkuCode',
                        how='left'
                    ).drop(columns=['Item SkuCode'])
                else:
                    sku_map = df_orders[['Listing Sku Code', 'Product Sku Code']].drop_duplicates()
                    sku_with_stock = pd.merge(
                        sku_map,
                        inv_grouped,
                        left_on='Product Sku Code',
                        right_on='Item SkuCode',
                        how='left'
                    )
                    sku_with_stock['Qty'] = sku_with_stock['Qty'].fillna(0)
                    listing_stock_map = sku_with_stock.groupby('Listing Sku Code', as_index=False)['Qty'].min()

                    final_df = pd.merge(
                        pivot_sku,
                        listing_stock_map,
                        left_on=display_col_name,
                        right_on='Listing Sku Code',
                        how='left'
                    )

                final_df['Available Stock'] = final_df['Qty'].fillna(0).astype(int)
                final_df.drop(columns=['Qty'], inplace=True)
            else:
                st.error("OMS Inventory file mein 'Item SkuCode' ya 'Qty' column nahi mila!")
                st.stop()

        else:
            # Unicommerce Format
            st.info("Detected Format: **Unicommerce Sale Orders & Inventory Report**")
            target_col = "Seller SKU Code" if use_alt_sku else "Item Type Name"
            display_col_name = "Seller SKU Code" if use_alt_sku else "Product Name"

            required_order_cols = ['Category', 'Item Type Name', 'Channel Name']
            if use_alt_sku:
                required_order_cols.append('Seller SKU Code')

            missing_cols = [c for c in required_order_cols if c not in df_orders.columns]
            if missing_cols:
                st.error(f"Sale Orders file mein ye columns missing hain: {missing_cols}")
                st.stop()

            if exclude_cancelled and 'Sale Order Item Status' in df_orders.columns:
                df_orders = df_orders[df_orders['Sale Order Item Status'].astype(str).str.upper().str.strip() != 'CANCELLED'].copy()

            df_orders['Category'] = df_orders['Category'].fillna('Uncategorized').astype(str).str.strip()
            df_orders['Item Type Name'] = df_orders['Item Type Name'].fillna('Unknown').astype(str).str.strip()
            df_orders['Channel Name'] = df_orders['Channel Name'].fillna('Other').astype(str).str.strip()
            if use_alt_sku:
                df_orders['Seller SKU Code'] = df_orders['Seller SKU Code'].fillna('Unknown_SKU').astype(str).str.strip()

            df_orders['Clean_Channel'] = df_orders['Channel Name'].replace(UNICOM_CHANNEL_MAPPING)

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
                index=[target_col, 'Category'],
                columns='Clean_Channel',
                aggfunc='sum',
                fill_value=0
            ).reset_index()

            pivot_sku.rename(columns={target_col: display_col_name, 'Category': 'Parent'}, inplace=True)

            channel_cols = [c for c in pivot_sku.columns if c not in [display_col_name, 'Parent']]
            preferred_order = ["Lymio", "Ajio PX", "Myntra JB", "Flipkart", "Ajio AT", "Snapdeal"]
            ordered_channels = [c for c in preferred_order if c in channel_cols] + [c for c in channel_cols if c not in preferred_order]
            
            pivot_sku = pivot_sku[[display_col_name, 'Parent'] + ordered_channels]
            pivot_sku['Total Sale'] = pivot_sku[ordered_channels].sum(axis=1)

            if 'Item Name' in df_inv.columns and 'Available (ATP)' in df_inv.columns:
                df_inv['Item Name'] = df_inv['Item Name'].astype(str).str.strip()
                df_inv['Available (ATP)'] = pd.to_numeric(df_inv['Available (ATP)'], errors='coerce').fillna(0)
                inv_grouped = df_inv.groupby('Item Name', as_index=False)['Available (ATP)'].sum()

                if not use_alt_sku:
                    final_df = pd.merge(
                        pivot_sku,
                        inv_grouped,
                        left_on=display_col_name,
                        right_on='Item Name',
                        how='left'
                    ).drop(columns=['Item Name'])
                else:
                    sku_to_item = df_orders[['Seller SKU Code', 'Item Type Name']].drop_duplicates()
                    sku_with_stock = pd.merge(
                        sku_to_item,
                        inv_grouped,
                        left_on='Item Type Name',
                        right_on='Item Name',
                        how='left'
                    )
                    sku_with_stock['Available (ATP)'] = sku_with_stock['Available (ATP)'].fillna(0)
                    seller_stock_map = sku_with_stock.groupby('Seller SKU Code', as_index=False)['Available (ATP)'].min()

                    final_df = pd.merge(
                        pivot_sku,
                        seller_stock_map,
                        left_on=display_col_name,
                        right_on='Seller SKU Code',
                        how='left'
                    )

                final_df['Available Stock'] = final_df['Available (ATP)'].fillna(0).astype(int)
                final_df.drop(columns=['Available (ATP)'], inplace=True)
            else:
                st.error("Inventory file mein 'Item Name' ya 'Available (ATP)' column nahi mila!")
                st.stop()

        # Common Processing: Stock Status + Size Sorting + Export
        final_df['Total Sale'] = final_df['Total Sale'].astype(int)
        for ch in ordered_channels:
            final_df[ch] = final_df[ch].astype(int)

        def get_stock_status(row):
            stock = row['Available Stock']
            sale = row['Total Sale']
            if stock <= 0:
                return "Out of Stock (0)"
            elif stock < sale:
                return "Low Stock (<1 Day)"
            elif stock < (sale * min_days_cover):
                return f"Reorder Soon ({round(stock/sale, 1)}d)"
            else:
                return f"Sufficient ({round(stock/sale, 1)}d)"

        final_df['Stock Status'] = final_df.apply(get_stock_status, axis=1)

        # Smart Size Sorting (XS, S, M, L, XL, 2XL... ya 30, 32, 34, 36, 38)
        sort_tuples = final_df[display_col_name].apply(get_sort_keys)
        final_df['_base_sort'] = [t[0] for t in sort_tuples]
        final_df['_size_rank'] = [t[1] for t in sort_tuples]
        final_df['_raw_size'] = [t[2] for t in sort_tuples]

        final_df = final_df.sort_values(
            by=['Parent', '_base_sort', '_size_rank', '_raw_size', display_col_name]
        ).drop(columns=['_base_sort', '_size_rank', '_raw_size']).reset_index(drop=True)

        # Sheet 2: Category-wise Summary
        numeric_cols = [c for c in final_df.columns if c not in [display_col_name, 'Parent', 'Stock Status']]
        category_summary = final_df.groupby('Parent', as_index=False)[numeric_cols].sum()
        category_summary.rename(columns={'Parent': 'Row Labels'}, inplace=True)

        def get_cat_status(row):
            stock = row['Available Stock']
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
        grand_total_cat['Stock Status'] = ''
        category_summary = pd.concat([category_summary, pd.DataFrame([grand_total_cat])], ignore_index=True)

        # 2-Line Header Format (Channel Name \n Date)
        rename_date_cols = {c: f"{c}\n{report_date_str}" for c in ordered_channels}
        final_df_export = final_df.rename(columns=rename_date_cols)
        category_summary_export = category_summary.rename(columns=rename_date_cols)

        top_totals = {col: (final_df_export[col].sum() if col not in [display_col_name, 'Parent', 'Stock Status'] else '') for col in final_df_export.columns}
        top_totals[display_col_name] = 'GRAND TOTAL'

        st.success(f"Report ({display_col_name} wise) successfully generate ho gayi hai!")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric(f"Total Unique SKUs", len(final_df))
        m2.metric("Total Units Sold", int(final_df['Total Sale'].sum()))
        low_count = final_df['Stock Status'].str.contains('Low Stock|Out of Stock').sum()
        good_count = final_df['Stock Status'].str.contains('Sufficient').sum()
        m3.metric("🔴 Low / Out of Stock", int(low_count))
        m4.metric("🟢 Sufficient Stock", int(good_count))

        def highlight_status_web(val):
            val_str = str(val)
            if "Low Stock" in val_str or "Out of Stock" in val_str:
                return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;'
            elif "Reorder Soon" in val_str:
                return 'background-color: #FFEB9C; color: #9C6500; font-weight: bold;'
            elif "Sufficient" in val_str:
                return 'background-color: #C6EFCE; color: #006100; font-weight: bold;'
            return ''

        st.subheader(f"Preview: {display_col_name} Wise Daily Sale & Available Stock")
        st.dataframe(final_df_export.style.map(highlight_status_web, subset=['Stock Status']), use_container_width=True)

        # Excel File Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame([top_totals]).to_excel(writer, sheet_name='SKU_Wise_Report', index=False, header=False, startrow=0)
            final_df_export.to_excel(writer, sheet_name='SKU_Wise_Report', index=False, startrow=1)
            category_summary_export.to_excel(writer, sheet_name='Category_Summary', index=False)

            format_excel_sheet(writer.sheets['SKU_Wise_Report'], header_row_idx=2, status_col_name="Stock Status")
            format_excel_sheet(writer.sheets['Category_Summary'], header_row_idx=1, status_col_name="Stock Status")

        output.seek(0)
        system_tag = "OMS" if is_oms else "Unicom"
        file_name = f"Daily Sale Report {report_date_str} ({system_tag}).xlsx"

        st.download_button(
            label="📥 Download Daily Sale Report (Excel)",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
