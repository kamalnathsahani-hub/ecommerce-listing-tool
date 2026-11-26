import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Daily Sale Report Generator", layout="wide")

st.title("📊 Daily Sale Report, Ads & Liquidation Generator")
st.write(
    "Unicommerce ya OMS ki Sale Orders aur Inventory files (CSV, Excel, ya ZIP) upload karein "
    "aur channel-wise SKU sale, Channel Product ID pivot, Ads priority, Marketplace Top 10/20, Stock health, aur Liquidation alerts ki report download karein."
)

# Unicommerce Channel Mapping (AU -> AV Actual Name)
UNICOM_CHANNEL_MAPPING = {
    "AJIO_Leriya_fashion_AT": "Leriya_Ajio",
    "AJIO_Vrati_Fashion_PX": "Vrati_Ajio",
    "FLIPKART_POSHAX": "Leriya_Flipkart",
    "JBFASHION": "JB Fashion_VC",
    "JBFASHION_MYNTRA": "JB Fashion_Myntra",
    "LERIYAFASHION": "Leriya_VC",
    "MEESHOLERIYA": "Leriya_Meesho",
    "MEESHOYHUB": "Y Hub_Meesho",
    "SNAPDEAL_LF": "Leriya_Snapdeal",
    "YHUB": "Y Hub_VC",
}

# OMS Channel Mapping (Unchanged)
OMS_CHANNEL_MAPPING = {
    "URBAN POCKETS LIFESTYLE - Myntra PPMP": "Myntra PPMP",
    "URBAN POCKETS LIFESTYLE - Amazon Direct Fulfillment": "Amazon DF",
    "URBAN POCKETS LIFESTYLE - Ajio Api": "Ajio API",
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
    "FREE": 90, "FREESIZE": 90, "FS": 90,
}


def get_sort_keys(sku_name):
    s = str(sku_name).strip()
    parts = s.rsplit("-", 1)
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
    if fname.endswith(".zip"):
        with zipfile.ZipFile(uploaded_file, "r") as z:
            valid_files = [
                f for f in z.namelist()
                if not f.startswith("__MACOSX")
                and (f.lower().endswith(".csv") or f.lower().endswith(".xlsx") or f.lower().endswith(".xls"))
            ]
            if not valid_files:
                raise ValueError(f"ZIP file '{uploaded_file.name}' ke andar koi CSV ya Excel file nahi mili.")
            inner_name = valid_files[0]
            with z.open(inner_name) as f:
                if inner_name.lower().endswith(".csv"):
                    return pd.read_csv(f, low_memory=False)
                return pd.read_excel(f)
    elif fname.endswith(".csv"):
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


def classify_sales_tier(rank_val, sale_val):
    if sale_val <= 0 or pd.isna(rank_val):
        return "Zero Sale"
    rank_int = int(rank_val)
    if rank_int <= 10:
        return "Top 10 (Ads Priority)"
    elif rank_int <= 20:
        return "Top 11-20 (Ads Scale)"
    elif rank_int <= 50:
        return "Top 21-50 (Test Ads)"
    return "Normal Mover"


def get_liquidation_status(stock, sale, min_liq_stock, max_liq_sale, high_cover_days):
    if stock >= min_liq_stock and sale <= max_liq_sale:
        return "Liquidate (High Stock / Low Sale)"
    if sale > 0 and (stock / sale) >= high_cover_days and stock >= min_liq_stock:
        return f"Watch Liquidation ({round(stock / sale, 1)}d cover)"
    return "Normal"


def auto_fit_columns(ws, start_row=1, min_width=14, max_width=60, padding=4):
    """Sabhi columns ki width header + data content ke hisab se auto-adjust karta hai."""
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(start_row, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None:
                continue
            val_str = str(val)
            # Row 1 ke =SUBTOTAL(...) formula ki length ko ignore karein
            if row_idx == 1 and val_str.startswith("="):
                continue
            lines = val_str.split("\n")
            longest_line = max(len(line.strip()) for line in lines)
            if longest_line > max_len:
                max_len = longest_line
        calculated_width = min(max(int(max_len * 1.15) + padding, min_width), max_width)
        ws.column_dimensions[col_letter].width = calculated_width


def format_excel_sheet(
    ws,
    numeric_col_indices,
    freeze_cell,
    status_col_name="Stock Status",
    tier_col_name="Sales Tier",
    liq_col_name="Liquidation Alert",
):
    """
    Sheet formatting:
    - Row 1: GRAND TOTAL with dynamic =SUBTOTAL(9, ...) formula on numeric columns
    - Row 2: Headers with Wrap Text + AutoFilter enabled
    - Row 3+: Data rows with conditional colors + auto-adjusted column widths
    """
    header_row_idx = 2
    data_start_row = 3
    data_end_row = ws.max_row

    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006", bold=True)

    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    yellow_font = Font(color="9C6500", bold=True)

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(color="006100", bold=True)

    ads_top10_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    ads_top10_font = Font(color="1B5E20", bold=True)

    ads_top20_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    ads_top20_font = Font(color="2E7D32", bold=True)

    liq_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    liq_font = Font(color="C00000", bold=True)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    ws.freeze_panes = freeze_cell

    last_col_letter = get_column_letter(ws.max_column)
    if data_end_row >= header_row_idx:
        ws.auto_filter.ref = f"A{header_row_idx}:{last_col_letter}{data_end_row}"

    ws.row_dimensions[header_row_idx].height = 36

    status_col_idx = None
    tier_col_idx = None
    liq_col_idx = None

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = thin_border

        col_title = str(cell.value)
        if status_col_name and col_title == status_col_name:
            status_col_idx = col_idx
        elif tier_col_name and col_title == tier_col_name:
            tier_col_idx = col_idx
        elif liq_col_name and col_title == liq_col_name:
            liq_col_idx = col_idx

    # Row 1: Dynamic =SUBTOTAL(9, ...) only on actual numeric measure columns
    ws.row_dimensions[1].height = 24
    numeric_set = set(numeric_col_indices)

    for col_idx in range(1, ws.max_column + 1):
        c = ws.cell(row=1, column=col_idx)
        c.font = Font(bold=True)
        c.fill = total_fill
        c.border = thin_border
        col_letter = get_column_letter(col_idx)

        if col_idx == 1:
            c.value = "GRAND TOTAL"
            c.alignment = Alignment(horizontal="left", vertical="center")
        elif col_idx in numeric_set:
            if data_end_row >= data_start_row:
                c.value = f"=SUBTOTAL(9,{col_letter}{data_start_row}:{col_letter}{data_end_row})"
            else:
                c.value = 0
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.value = ""
            c.alignment = Alignment(horizontal="center", vertical="center")

    # Data Rows Formatting
    for row_idx in range(data_start_row, data_end_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border

            if col_idx in numeric_set:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if status_col_idx and col_idx == status_col_idx:
                val = str(cell.value)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if "Low Stock" in val or "Out of Stock" in val:
                    cell.fill = red_fill
                    cell.font = red_font
                elif "Reorder Soon" in val:
                    cell.fill = yellow_fill
                    cell.font = yellow_font
                elif "Sufficient" in val:
                    cell.fill = green_fill
                    cell.font = green_font

            if tier_col_idx and col_idx == tier_col_idx:
                val = str(cell.value)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if "Top 10" in val:
                    cell.fill = ads_top10_fill
                    cell.font = ads_top10_font
                elif "Top 11-20" in val:
                    cell.fill = ads_top20_fill
                    cell.font = ads_top20_font

            if liq_col_idx and col_idx == liq_col_idx:
                val = str(cell.value)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if "Liquidate (" in val:
                    cell.fill = liq_fill
                    cell.font = liq_font
                elif "Watch Liquidation" in val:
                    cell.fill = yellow_fill
                    cell.font = yellow_font

    auto_fit_columns(ws, start_row=1, min_width=14, max_width=60, padding=4)


# Upload & Controls
col1, col2, col3 = st.columns([2, 2, 1.5])
with col1:
    orders_file = st.file_uploader("1. Sale Orders File Upload Karein (CSV / XLSX / ZIP)", type=["csv", "xlsx", "xls", "zip"])
with col2:
    inventory_file = st.file_uploader("2. Inventory Database File Upload Karein (CSV / XLSX / ZIP)", type=["csv", "xlsx", "xls", "zip"])
with col3:
    min_days_cover = st.number_input(
        "Safe Stock Days (Green ke liye)",
        min_value=1,
        max_value=180,
        value=30,
        help="Agar Available Stock (Sale × Days) se kam hai to Red/Yellow dikhega",
    )

opt_col1, opt_col2 = st.columns([2, 2])
with opt_col1:
    sku_mode = st.radio(
        "Report kis SKU column ke hisab se banani hai?",
        options=[
            "Primary Product SKU (Item Type Name / Product Sku Code)",
            "Channel / Seller SKU (Seller SKU Code / Listing Sku Code)",
        ],
        horizontal=True,
    )
with opt_col2:
    exclude_cancelled = st.checkbox("Cancelled orders ko report se hatayein (Recommended)", value=True)

with st.expander("⚙️ Ads, Marketplace Top N & Liquidation Rules Customize Karein (Optional)"):
    r1, r2, r3, r4 = st.columns(4)
    with r1:
        liq_min_stock = st.number_input(
            "Liquidation Min Stock",
            min_value=1,
            max_value=100000,
            value=20,
            help="Isse zyada stock hone par slow-moving SKU ko Liquidation ke liye flag kiya jayega",
        )
    with r2:
        liq_max_sale = st.number_input(
            "Liquidation Max Sale",
            min_value=0,
            max_value=1000,
            value=1,
            help="Isse kam ya barabar sale hone par high-stock SKU ko Liquidation alert milega",
        )
    with r3:
        liq_cover_days = st.number_input(
            "High Cover Days Alert",
            min_value=7,
            max_value=365,
            value=30,
            help="Agar stock itne din ki sale ke barabar ya zyada hai (with min stock), to Watch Liquidation dikhega",
        )
    with r4:
        mp_top_n = st.selectbox(
            "Marketplace-wise Top SKUs",
            options=[10, 20],
            index=1,
            help="Har marketplace ke Top 10 ya Top 20 SKUs alag report mein dikhenge",
        )

if orders_file and inventory_file:
    try:
        df_orders = load_any_file(orders_file)
        df_inv = load_any_file(inventory_file)

        use_alt_sku = "Channel / Seller SKU" in sku_mode
        is_oms = ("Product Sku Code" in df_orders.columns and "Category Name" in df_orders.columns)

        cpid_pivot_df = None  # Sirf Unicommerce format ke liye generate hoga

        if is_oms:
            st.info("Detected Format: **OMS Order & Inventory Report**")
            target_col = "Listing Sku Code" if use_alt_sku else "Product Sku Code"
            display_col_name = "Listing Sku Code" if use_alt_sku else "Product Sku Code"
            cat_col = "Category Name"

            if exclude_cancelled and "Order Status" in df_orders.columns:
                cancel_statuses = ["CANCELLED", "CANCEL INIT", "CANCELLED RETURN RECEIVED"]
                df_orders = df_orders[~df_orders["Order Status"].astype(str).str.upper().str.strip().isin(cancel_statuses)].copy()

            df_orders[cat_col] = df_orders[cat_col].fillna("Uncategorized").astype(str).str.strip()
            df_orders["Product Sku Code"] = df_orders["Product Sku Code"].fillna("Unknown").astype(str).str.strip()
            if use_alt_sku:
                df_orders["Listing Sku Code"] = df_orders["Listing Sku Code"].fillna("Unknown_Listing_SKU").astype(str).str.strip()

            df_orders["Clean_Channel"] = df_orders["Channel Name"].fillna("Other").apply(clean_oms_channel)

            if "Qty" in df_orders.columns:
                df_orders["Units"] = pd.to_numeric(df_orders["Qty"], errors="coerce").fillna(1)
            else:
                df_orders["Units"] = 1

            report_date_str = datetime.now().strftime("%d-%m")
            for date_col in ["Invoice Date", "Order Date"]:
                if date_col in df_orders.columns and df_orders[date_col].dropna().shape[0] > 0:
                    try:
                        parsed_dates = pd.to_datetime(df_orders[date_col].dropna(), errors="coerce").dropna()
                        if not parsed_dates.empty:
                            report_date_str = parsed_dates.max().strftime("%d-%m")
                            break
                    except Exception:
                        pass

            pivot_sku = pd.pivot_table(
                df_orders,
                values="Units",
                index=[target_col, cat_col],
                columns="Clean_Channel",
                aggfunc="sum",
                fill_value=0,
            ).reset_index()

            pivot_sku.rename(columns={target_col: display_col_name, cat_col: "Parent"}, inplace=True)

            channel_cols = [c for c in pivot_sku.columns if c not in [display_col_name, "Parent"]]
            preferred_order = ["Myntra PPMP", "Amazon DF", "Ajio API"]
            ordered_channels = [c for c in preferred_order if c in channel_cols] + [c for c in channel_cols if c not in preferred_order]

            pivot_sku = pivot_sku[[display_col_name, "Parent"] + ordered_channels]
            pivot_sku["Total Sale"] = pivot_sku[ordered_channels].sum(axis=1)

            if "Item SkuCode" in df_inv.columns and "Qty" in df_inv.columns:
                df_inv["Item SkuCode"] = df_inv["Item SkuCode"].astype(str).str.strip()
                df_inv["Qty"] = pd.to_numeric(df_inv["Qty"], errors="coerce").fillna(0)
                inv_grouped = df_inv.groupby("Item SkuCode", as_index=False)["Qty"].sum()

                if not use_alt_sku:
                    final_df = pd.merge(
                        pivot_sku,
                        inv_grouped,
                        left_on=display_col_name,
                        right_on="Item SkuCode",
                        how="left",
                    ).drop(columns=["Item SkuCode"])
                else:
                    sku_map = df_orders[["Listing Sku Code", "Product Sku Code"]].drop_duplicates()
                    sku_with_stock = pd.merge(
                        sku_map,
                        inv_grouped,
                        left_on="Product Sku Code",
                        right_on="Item SkuCode",
                        how="left",
                    )
                    sku_with_stock["Qty"] = sku_with_stock["Qty"].fillna(0)
                    listing_stock_map = sku_with_stock.groupby("Listing Sku Code", as_index=False)["Qty"].min()

                    final_df = pd.merge(
                        pivot_sku,
                        listing_stock_map,
                        left_on=display_col_name,
                        right_on="Listing Sku Code",
                        how="left",
                    )

                final_df["Available Stock"] = final_df["Qty"].fillna(0).astype(int)
                final_df.drop(columns=["Qty"], inplace=True)
            else:
                st.error("OMS Inventory file mein 'Item SkuCode' ya 'Qty' column nahi mila!")
                st.stop()

        else:
            st.info("Detected Format: **Unicommerce Sale Orders & Inventory Report**")
            target_col = "Seller SKU Code" if use_alt_sku else "Item Type Name"
            display_col_name = "Seller SKU Code" if use_alt_sku else "Product Name"

            required_order_cols = ["Category", "Item Type Name", "Channel Name"]
            if use_alt_sku:
                required_order_cols.append("Seller SKU Code")

            missing_cols = [c for c in required_order_cols if c not in df_orders.columns]
            if missing_cols:
                st.error(f"Sale Orders file mein ye columns missing hain: {missing_cols}")
                st.stop()

            if exclude_cancelled and "Sale Order Item Status" in df_orders.columns:
                df_orders = df_orders[df_orders["Sale Order Item Status"].astype(str).str.upper().str.strip() != "CANCELLED"].copy()

            df_orders["Category"] = df_orders["Category"].fillna("Uncategorized").astype(str).str.strip()
            df_orders["Item Type Name"] = df_orders["Item Type Name"].fillna("Unknown").astype(str).str.strip()
            df_orders["Channel Name"] = df_orders["Channel Name"].fillna("Other").astype(str).str.strip()
            if use_alt_sku:
                df_orders["Seller SKU Code"] = df_orders["Seller SKU Code"].fillna("Unknown_SKU").astype(str).str.strip()

            # Map AU (Channel Name) -> AV (Actual Name)
            df_orders["Clean_Channel"] = df_orders["Channel Name"].replace(UNICOM_CHANNEL_MAPPING)

            # Start Date to End Date from 'Order Date as dd/mm/yyyy hh:MM:ss'
            report_date_str = datetime.now().strftime("%d-%m")
            order_date_col = "Order Date as dd/mm/yyyy hh:MM:ss"
            if order_date_col in df_orders.columns and df_orders[order_date_col].dropna().shape[0] > 0:
                try:
                    parsed_order_dates = pd.to_datetime(
                        df_orders[order_date_col].dropna(), dayfirst=True, errors="coerce"
                    ).dropna()
                    if not parsed_order_dates.empty:
                        start_dt = parsed_order_dates.min().strftime("%d-%m")
                        end_dt = parsed_order_dates.max().strftime("%d-%m")
                        report_date_str = start_dt if start_dt == end_dt else f"{start_dt} to {end_dt}"
                except Exception:
                    pass
            elif "Invoice Created" in df_orders.columns and df_orders["Invoice Created"].dropna().shape[0] > 0:
                try:
                    sample_date = pd.to_datetime(df_orders["Invoice Created"].dropna().iloc[0], dayfirst=True, errors="coerce")
                    if pd.notna(sample_date):
                        report_date_str = sample_date.strftime("%d-%m")
                except Exception:
                    pass

            df_orders["Units"] = 1

            pivot_sku = pd.pivot_table(
                df_orders,
                values="Units",
                index=[target_col, "Category"],
                columns="Clean_Channel",
                aggfunc="sum",
                fill_value=0,
            ).reset_index()

            pivot_sku.rename(columns={target_col: display_col_name, "Category": "Parent"}, inplace=True)

            channel_cols = [c for c in pivot_sku.columns if c not in [display_col_name, "Parent"]]
            preferred_order = [
                "JB Fashion_VC",
                "Vrati_Ajio",
                "JB Fashion_Myntra",
                "Leriya_Flipkart",
                "Leriya_Ajio",
                "Leriya_Snapdeal",
                "Leriya_VC",
                "Leriya_Meesho",
                "Y Hub_Meesho",
                "Y Hub_VC",
            ]
            ordered_channels = [c for c in preferred_order if c in channel_cols] + [c for c in channel_cols if c not in preferred_order]

            pivot_sku = pivot_sku[[display_col_name, "Parent"] + ordered_channels]
            pivot_sku["Total Sale"] = pivot_sku[ordered_channels].sum(axis=1)

            # Channel Product Id Wise Pivot (Sirf Unicommerce format ke liye)
            if "Channel Product Id" in df_orders.columns:
                df_orders["Channel Product Id"] = (
                    df_orders["Channel Product Id"].fillna("Unknown_ID").astype(str).str.strip()
                )
                cpid_pivot = pd.pivot_table(
                    df_orders,
                    values="Units",
                    index=["Channel Product Id", "Category"],
                    columns="Clean_Channel",
                    aggfunc="sum",
                    fill_value=0,
                ).reset_index()
                cpid_pivot.rename(columns={"Category": "Parent"}, inplace=True)

                for ch in ordered_channels:
                    if ch not in cpid_pivot.columns:
                        cpid_pivot[ch] = 0
                    cpid_pivot[ch] = cpid_pivot[ch].astype(int)

                # Har Channel Product Id ke saath mapped SKU(s) bhi dikhayein
                mapped_skus = (
                    df_orders.groupby(["Channel Product Id", "Category"])[target_col]
                    .apply(lambda s: ", ".join(sorted(set(s.astype(str)))))
                    .reset_index()
                    .rename(columns={"Category": "Parent", target_col: f"Mapped {display_col_name}"})
                )
                cpid_pivot = pd.merge(cpid_pivot, mapped_skus, on=["Channel Product Id", "Parent"], how="left")

                cpid_pivot["Total Sale"] = cpid_pivot[ordered_channels].sum(axis=1).astype(int)
                cpid_pivot = cpid_pivot.sort_values(
                    by=["Total Sale", "Channel Product Id"], ascending=[False, True]
                ).reset_index(drop=True)

                cpid_pivot["Sale Rank"] = 0
                pos_cpid = cpid_pivot["Total Sale"] > 0
                cpid_pivot.loc[pos_cpid, "Sale Rank"] = (
                    cpid_pivot.loc[pos_cpid, "Total Sale"]
                    .rank(ascending=False, method="min")
                    .astype(int)
                )

                cpid_col_order = (
                    ["Channel Product Id", "Parent", f"Mapped {display_col_name}", "Sale Rank"]
                    + ordered_channels
                    + ["Total Sale"]
                )
                cpid_pivot_df = cpid_pivot[cpid_col_order]

            if "Item Name" in df_inv.columns and "Available (ATP)" in df_inv.columns:
                df_inv["Item Name"] = df_inv["Item Name"].astype(str).str.strip()
                df_inv["Available (ATP)"] = pd.to_numeric(df_inv["Available (ATP)"], errors="coerce").fillna(0)
                inv_grouped = df_inv.groupby("Item Name", as_index=False)["Available (ATP)"].sum()

                if not use_alt_sku:
                    final_df = pd.merge(
                        pivot_sku,
                        inv_grouped,
                        left_on=display_col_name,
                        right_on="Item Name",
                        how="left",
                    ).drop(columns=["Item Name"])
                else:
                    sku_to_item = df_orders[["Seller SKU Code", "Item Type Name"]].drop_duplicates()
                    sku_with_stock = pd.merge(
                        sku_to_item,
                        inv_grouped,
                        left_on="Item Type Name",
                        right_on="Item Name",
                        how="left",
                    )
                    sku_with_stock["Available (ATP)"] = sku_with_stock["Available (ATP)"].fillna(0)
                    seller_stock_map = sku_with_stock.groupby("Seller SKU Code", as_index=False)["Available (ATP)"].min()

                    final_df = pd.merge(
                        pivot_sku,
                        seller_stock_map,
                        left_on=display_col_name,
                        right_on="Seller SKU Code",
                        how="left",
                    )

                final_df["Available Stock"] = final_df["Available (ATP)"].fillna(0).astype(int)
                final_df.drop(columns=["Available (ATP)"], inplace=True)
            else:
                st.error("Inventory file mein 'Item Name' ya 'Available (ATP)' column nahi mila!")
                st.stop()

        # Common Processing & Business Intelligence Columns
        final_df["Total Sale"] = final_df["Total Sale"].astype(int)
        final_df["Available Stock"] = final_df["Available Stock"].astype(int)
        for ch in ordered_channels:
            final_df[ch] = final_df[ch].astype(int)

        # Active Channels count (kitne channels par sale hui)
        final_df["Active Channels"] = (final_df[ordered_channels] > 0).sum(axis=1).astype(int)

        # Sale Rank & Ads Tier (Top 10, Top 11-20, Top 21-50, Normal Mover, Zero Sale)
        positive_sales_mask = final_df["Total Sale"] > 0
        final_df["Sale Rank"] = 0
        final_df.loc[positive_sales_mask, "Sale Rank"] = (
            final_df.loc[positive_sales_mask, "Total Sale"]
            .rank(ascending=False, method="min")
            .astype(int)
        )
        final_df["Sales Tier"] = [
            classify_sales_tier(r, s) for r, s in zip(final_df["Sale Rank"], final_df["Total Sale"])
        ]

        def get_stock_status(row):
            stock = row["Available Stock"]
            sale = row["Total Sale"]
            if stock <= 0:
                return "Out of Stock (0)"
            elif sale > 0 and stock < sale:
                return "Low Stock (<1 Day)"
            elif sale > 0 and stock < (sale * min_days_cover):
                return f"Reorder Soon ({round(stock / sale, 1)}d)"
            elif sale == 0 and stock > 0:
                return "Sufficient (No Sale)"
            else:
                return f"Sufficient ({round(stock / sale, 1)}d)"

        final_df["Stock Status"] = final_df.apply(get_stock_status, axis=1)

        final_df["Liquidation Alert"] = [
            get_liquidation_status(stk, sl, liq_min_stock, liq_max_sale, liq_cover_days)
            for stk, sl in zip(final_df["Available Stock"], final_df["Total Sale"])
        ]

        # Smart Size Sorting for main display
        sort_tuples = final_df[display_col_name].apply(get_sort_keys)
        final_df["_base_sort"] = [t[0] for t in sort_tuples]
        final_df["_size_rank"] = [t[1] for t in sort_tuples]
        final_df["_raw_size"] = [t[2] for t in sort_tuples]

        final_df = (
            final_df.sort_values(
                by=["Parent", "_base_sort", "_size_rank", "_raw_size", display_col_name]
            )
            .drop(columns=["_base_sort", "_size_rank", "_raw_size"])
            .reset_index(drop=True)
        )

        # Reorder columns cleanly so operational insights are easy to scan
        sku_col_order = (
            [display_col_name, "Parent", "Sale Rank", "Sales Tier", "Active Channels"]
            + ordered_channels
            + ["Total Sale", "Available Stock", "Stock Status", "Liquidation Alert"]
        )
        final_df = final_df[sku_col_order]

        # Sheet 2: Category-wise Summary
        cat_sum_cols = ordered_channels + ["Total Sale", "Available Stock"]
        category_summary = final_df.groupby("Parent", as_index=False)[cat_sum_cols].sum()
        category_summary.rename(columns={"Parent": "Row Labels"}, inplace=True)

        def get_cat_status(row):
            stock = row["Available Stock"]
            sale = row["Total Sale"]
            if sale == 0:
                return "Sufficient" if stock > 0 else "Out of Stock (0)"
            if stock <= 0:
                return "Out of Stock (0)"
            elif stock < sale:
                return "Low Stock (<1 Day)"
            elif stock < (sale * min_days_cover):
                return f"Reorder Soon ({round(stock / sale, 1)}d)"
            else:
                return f"Sufficient ({round(stock / sale, 1)}d)"

        category_summary["Stock Status"] = category_summary.apply(get_cat_status, axis=1)

        # Sheet 3: Overall Ads Priority (Top 20) & Liquidation Action List
        ads_top_df = (
            final_df[final_df["Total Sale"] > 0]
            .sort_values(by=["Total Sale", "Available Stock"], ascending=[False, False])
            .head(20)
            .copy()
        )
        ads_top_df.insert(0, "Action Focus", [f"Ads Target #{i+1}" for i in range(len(ads_top_df))])

        liq_df = (
            final_df[final_df["Liquidation Alert"] != "Normal"]
            .sort_values(by=["Available Stock", "Total Sale"], ascending=[False, True])
            .copy()
        )
        liq_df.insert(0, "Action Focus", "Liquidate / Clearance")

        ads_liq_combined = pd.concat([ads_top_df, liq_df], ignore_index=True)

        # Sheet 4: Marketplace-wise Top 10 / Top 20 Report
        mp_rows = []
        for ch in ordered_channels:
            ch_df = (
                final_df[final_df[ch] > 0]
                .sort_values(by=[ch, "Available Stock"], ascending=[False, False])
                .head(mp_top_n)
                .copy()
            )
            if ch_df.empty:
                continue
            ch_df = ch_df.reset_index(drop=True)
            for idx, row in ch_df.iterrows():
                mp_rank = idx + 1
                mp_tier = "Top 10 (Ads Priority)" if mp_rank <= 10 else "Top 11-20 (Ads Scale)"
                mp_rows.append(
                    {
                        "Marketplace": ch,
                        "MP Rank": mp_rank,
                        "MP Sales Tier": mp_tier,
                        display_col_name: row[display_col_name],
                        "Parent": row["Parent"],
                        "Marketplace Sale": int(row[ch]),
                        "Overall Total Sale": int(row["Total Sale"]),
                        "Available Stock": int(row["Available Stock"]),
                        "Stock Status": row["Stock Status"],
                        "Liquidation Alert": row["Liquidation Alert"],
                    }
                )

        if mp_rows:
            mp_top_df = pd.DataFrame(mp_rows)
        else:
            mp_top_df = pd.DataFrame(
                columns=[
                    "Marketplace",
                    "MP Rank",
                    "MP Sales Tier",
                    display_col_name,
                    "Parent",
                    "Marketplace Sale",
                    "Overall Total Sale",
                    "Available Stock",
                    "Stock Status",
                    "Liquidation Alert",
                ]
            )

        # 2-Line Header Format (Channel Name \n Date or Date Range)
        rename_date_cols = {c: f"{c}\n{report_date_str}" for c in ordered_channels}
        final_df_export = final_df.rename(columns=rename_date_cols)
        category_summary_export = category_summary.rename(columns=rename_date_cols)
        ads_liq_export = ads_liq_combined.rename(columns=rename_date_cols)
        cpid_pivot_export = cpid_pivot_df.rename(columns=rename_date_cols) if cpid_pivot_df is not None else None

        st.success(f"Report ({display_col_name} wise | {report_date_str}) successfully generate ho gayi hai!")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Unique SKUs", len(final_df))
        m2.metric("Total Units Sold", int(final_df["Total Sale"].sum()))
        top10_units = int(final_df[final_df["Sale Rank"].between(1, 10)]["Total Sale"].sum())
        m3.metric("🎯 Top 10 Ads Sale", top10_units)
        low_count = int(final_df["Stock Status"].str.contains("Low Stock|Out of Stock").sum())
        m4.metric("🔴 Low / Out of Stock", low_count)
        liq_count = int(final_df["Liquidation Alert"].str.startswith("Liquidate (").sum())
        m5.metric("🏷️ Liquidation Alerts", liq_count)

        def highlight_status_web(val):
            val_str = str(val)
            if "Low Stock" in val_str or "Out of Stock" in val_str:
                return "background-color: #FFC7CE; color: #9C0006; font-weight: bold;"
            elif "Reorder Soon" in val_str:
                return "background-color: #FFEB9C; color: #9C6500; font-weight: bold;"
            elif "Sufficient" in val_str:
                return "background-color: #C6EFCE; color: #006100; font-weight: bold;"
            return ""

        def highlight_tier_web(val):
            val_str = str(val)
            if "Top 10" in val_str:
                return "background-color: #D9EAD3; color: #1B5E20; font-weight: bold;"
            elif "Top 11-20" in val_str:
                return "background-color: #E8F5E9; color: #2E7D32; font-weight: bold;"
            return ""

        def highlight_liq_web(val):
            val_str = str(val)
            if "Liquidate (" in val_str:
                return "background-color: #FCE4D6; color: #C00000; font-weight: bold;"
            elif "Watch Liquidation" in val_str:
                return "background-color: #FFEB9C; color: #9C6500; font-weight: bold;"
            return ""

        tab_labels = [
            "📋 All SKUs Report",
            "🎯 Top 20 Ads Winners",
            "🏪 Marketplace Top 10/20",
            "🏷️ Liquidation Candidates",
        ]
        if cpid_pivot_export is not None:
            tab_labels.append("🆔 Channel Product ID Pivot")

        tabs = st.tabs(tab_labels)

        with tabs[0]:
            styled_main = (
                final_df_export.style
                .map(highlight_status_web, subset=["Stock Status"])
                .map(highlight_tier_web, subset=["Sales Tier"])
                .map(highlight_liq_web, subset=["Liquidation Alert"])
            )
            st.dataframe(styled_main, use_container_width=True)

        with tabs[1]:
            st.caption("Overall sabse zyada sale wale Top 20 SKUs — in par ads chalane ya budget scale karne par dhyan dein.")
            st.dataframe(ads_top_df, use_container_width=True)

        with tabs[2]:
            st.caption(f"Har marketplace ke Top {mp_top_n} SKUs — marketplace-wise ads aur inventory planning ke liye.")
            selected_mp = st.selectbox(
                "Marketplace Filter Karein:",
                options=["All Marketplaces"] + ordered_channels,
                key="mp_filter_select",
            )
            mp_view_df = mp_top_df if selected_mp == "All Marketplaces" else mp_top_df[mp_top_df["Marketplace"] == selected_mp]
            styled_mp = (
                mp_view_df.style
                .map(highlight_status_web, subset=["Stock Status"])
                .map(highlight_tier_web, subset=["MP Sales Tier"])
                .map(highlight_liq_web, subset=["Liquidation Alert"])
            )
            st.dataframe(styled_mp, use_container_width=True)

        with tabs[3]:
            st.caption(f"High Stock (>= {liq_min_stock}) aur Low Sale (<= {liq_max_sale}) wale SKUs — liquidation / offer / bundle ke liye.")
            st.dataframe(liq_df, use_container_width=True)

        if cpid_pivot_export is not None:
            with tabs[4]:
                st.caption("Sabhi marketplaces ka Channel Product ID-wise sale pivot report.")
                st.dataframe(cpid_pivot_export, use_container_width=True)

        # Excel File Export with Top =SUBTOTAL(9,...) Row, AutoFilter, Freeze Panes & Auto Column Width
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df_export.to_excel(writer, sheet_name="SKU_Wise_Report", index=False, startrow=1)
            category_summary_export.to_excel(writer, sheet_name="Category_Summary", index=False, startrow=1)
            ads_liq_export.to_excel(writer, sheet_name="Ads_&_Liquidation", index=False, startrow=1)
            mp_top_df.to_excel(writer, sheet_name="Marketplace_Top_SKUs", index=False, startrow=1)

            if cpid_pivot_export is not None:
                cpid_pivot_export.to_excel(writer, sheet_name="Channel_Product_ID_Report", index=False, startrow=1)

            # Sheet 1: SKU_Wise_Report
            sku_numeric_cols = [5] + list(range(6, 6 + len(ordered_channels))) + [
                6 + len(ordered_channels),
                7 + len(ordered_channels),
            ]
            format_excel_sheet(
                writer.sheets["SKU_Wise_Report"],
                numeric_col_indices=sku_numeric_cols,
                freeze_cell="C3",
                status_col_name="Stock Status",
                tier_col_name="Sales Tier",
                liq_col_name="Liquidation Alert",
            )

            # Sheet 2: Category_Summary
            cat_numeric_cols = list(range(2, 2 + len(ordered_channels) + 2))
            format_excel_sheet(
                writer.sheets["Category_Summary"],
                numeric_col_indices=cat_numeric_cols,
                freeze_cell="B3",
                status_col_name="Stock Status",
                tier_col_name="",
                liq_col_name="",
            )

            # Sheet 3: Ads_&_Liquidation
            ads_liq_numeric_cols = [6] + list(range(7, 7 + len(ordered_channels))) + [
                7 + len(ordered_channels),
                8 + len(ordered_channels),
            ]
            format_excel_sheet(
                writer.sheets["Ads_&_Liquidation"],
                numeric_col_indices=ads_liq_numeric_cols,
                freeze_cell="D3",
                status_col_name="Stock Status",
                tier_col_name="Sales Tier",
                liq_col_name="Liquidation Alert",
            )

            # Sheet 4: Marketplace_Top_SKUs
            mp_numeric_cols = [6, 7, 8]
            format_excel_sheet(
                writer.sheets["Marketplace_Top_SKUs"],
                numeric_col_indices=mp_numeric_cols,
                freeze_cell="E3",
                status_col_name="Stock Status",
                tier_col_name="MP Sales Tier",
                liq_col_name="Liquidation Alert",
            )

            # Sheet 5: Channel_Product_ID_Report (Only for Unicommerce format)
            if cpid_pivot_export is not None:
                # Cols: 1:Channel Product Id, 2:Parent, 3:Mapped SKUs, 4:Sale Rank, 5..N:Channels, N+1:Total Sale
                cpid_numeric_cols = list(range(5, 5 + len(ordered_channels) + 1))
                format_excel_sheet(
                    writer.sheets["Channel_Product_ID_Report"],
                    numeric_col_indices=cpid_numeric_cols,
                    freeze_cell="E3",
                    status_col_name="",
                    tier_col_name="",
                    liq_col_name="",
                )

        output.seek(0)
        system_tag = "OMS" if is_oms else "Unicom"
        file_name = f"Daily Sale Report {report_date_str} ({system_tag}).xlsx"

        st.download_button(
            label="📥 Download Upgraded Daily Sale Report (Excel)",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
