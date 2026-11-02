import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Inventory Updater", page_icon="📦", layout="wide")

st.title("📦 Inventory Stock (ATP) Updater")
st.markdown("Shopify **inventory_export** file me Datatable ke **Available (ATP)** se **On hand (new)** update karein. Unmapped items me **0** fill hoga.")

# --- SIDEBAR YA OPTIONAL MASTER FILE ---
st.sidebar.header("Settings / Mapping File")
default_master_path = "inventory_export - Copy.csv"

uploaded_master = st.sidebar.file_uploader(
    "Upload Mapping File (inventory_export - Copy.csv)", 
    type=["csv"]
)

mapping_df = None
if uploaded_master is not None:
    mapping_df = pd.read_csv(uploaded_master)
    st.sidebar.success("Custom mapping file loaded!")
elif os.path.exists(default_master_path):
    mapping_df = pd.read_csv(default_master_path)
    st.sidebar.info(f"Default mapping file loaded: `{default_master_path}`")
else:
    st.sidebar.warning("Mapping file (inventory_export - Copy.csv) nahi mili. Kripya upload karein.")

# --- MAIN INPUT SECTION ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Main Inventory File")
    inv_file = st.file_uploader("Upload `inventory_export.csv`", type=["csv"], key="inv_file")

with col2:
    st.subheader("2. Datatable Search Inventory")
    datatable_file = st.file_uploader("Upload `DATATABLE SEARCH INVENTORY.csv`", type=["csv"], key="dt_file")

if inv_file and datatable_file:
    if mapping_df is None:
        st.error("Error: Mapping file (inventory_export - Copy.csv) available nahi hai. Sidebar se upload karein.")
    else:
        if st.button("🚀 Process & Update Inventory", type="primary"):
            with st.spinner("Processing files and matching SKUs..."):
                try:
                    # Files load karein
                    inv_df = pd.read_csv(inv_file)
                    dt_df = pd.read_csv(datatable_file)

                    # Composite keys for exact variant matching
                    key_cols = ['Handle', 'Option1 Value', 'Option2 Value', 'SKU']
                    available_keys = [c for c in key_cols if c in inv_df.columns and c in mapping_df.columns]
                    if not available_keys:
                        available_keys = ['SKU']

                    # Whitespace clean karein
                    for col in available_keys:
                        inv_df[col] = inv_df[col].astype(str).str.strip()
                        mapping_df[col] = mapping_df[col].astype(str).str.strip()

                    # Mapping DF se lookup merge karein
                    mapping_subset = mapping_df[available_keys + ['uni sku']].drop_duplicates(subset=available_keys)
                    merged = inv_df.merge(mapping_subset, on=available_keys, how='left')

                    # Datatable lookup
                    dt_df['Sku Code'] = dt_df['Sku Code'].astype(str).str.strip()
                    dt_lookup = dt_df.drop_duplicates(subset=['Sku Code']).set_index('Sku Code')['Available (ATP)'].to_dict()

                    # Match check karein: unmapped rows ki pehchan ke liye flag banayein
                    merged['matched_atp'] = merged['uni sku'].map(dt_lookup)
                    is_unmapped = merged['matched_atp'].isna()

                    # Unmapped par 0 fill karein aur integer format rakhein
                    inv_df['On hand (new)'] = merged['matched_atp'].fillna(0).astype(int)

                    matched_count = (~is_unmapped).sum()
                    unmatched_count = is_unmapped.sum()
                    total_count = len(inv_df)

                    st.success(f"Processing complete! **{matched_count}** matched, **{unmatched_count}** unmapped (0 filled) out of {total_count} rows.")

                    # --- PREVIEW WITH HIGHLIGHTING ---
                    st.write("### Preview (Unmapped items Yellow color me highlight hain):")
                    
                    preview_df = inv_df[['Handle', 'Option1 Value', 'Option2 Value', 'SKU', 'On hand (current)', 'On hand (new)']].copy()
                    
                    # Row highlight styling function
                    def highlight_unmapped_rows(row):
                        if is_unmapped.iloc[row.name]:
                            return ['background-color: #fff3cd; color: #856404; font-weight: bold;'] * len(row)
                        return [''] * len(row)

                    styled_preview = preview_df.head(50).style.apply(highlight_unmapped_rows, axis=1)
                    st.dataframe(styled_preview, use_container_width=True)

                    # --- DOWNLOAD BUTTONS ---
                    st.write("### Download Options:")
                    d_col1, d_col2 = st.columns(2)

                    with d_col1:
                        # 1. Standard CSV (Shopify upload ke liye jisme 0 fill hai)
                        csv_buffer = io.StringIO()
                        inv_df.to_csv(csv_buffer, index=False)
                        st.download_button(
                            label="📥 Download Updated CSV (For Shopify)",
                            data=csv_buffer.getvalue().encode('utf-8'),
                            file_name="inventory_export_updated.csv",
                            mime="text/csv",
                            use_container_width=True
                        )

                    with d_col2:
                        # 2. Highlighted Excel File (Review karne ke liye jisme unmapped rows colored hongi)
                        excel_buffer = io.BytesIO()
                        
                        def style_excel_unmapped(row):
                            if is_unmapped.iloc[row.name]:
                                return ['background-color: #FFF2CC'] * len(row)
                            return [''] * len(row)

                        styled_full = inv_df.style.apply(style_excel_unmapped, axis=1)
                        styled_full.to_excel(excel_buffer, index=False, engine='openpyxl')
                        
                        st.download_button(
                            label="📊 Download Highlighted Excel (.xlsx)",
                            data=excel_buffer.getvalue(),
                            file_name="inventory_export_highlighted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                except Exception as e:
                    st.error(f"Error during processing: {str(e)}")
