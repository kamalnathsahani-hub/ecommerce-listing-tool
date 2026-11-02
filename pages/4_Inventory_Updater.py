import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Inventory Updater", page_icon="📦", layout="wide")

st.title("📦 Inventory Stock (ATP) Updater")
st.markdown("Shopify **inventory_export** file me Datatable ke **Available (ATP)** se **On hand (new)** update karein.")

# --- SIDEBAR YA OPTIONAL MASTER FILE ---
st.sidebar.header("Settings / Mapping File")
# Agar project directory me inventory_export - Copy.csv rakhi hai to auto-load hogi
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

                    # Data cleaning & normalization
                    # Mapping DataFrame se uni sku mapping create karna
                    # Accurate mapping ke liye composite key use karenge (Handle + Option1 + Option2 + SKU)
                    key_cols = ['Handle', 'Option1 Value', 'Option2 Value', 'SKU']
                    
                    # Ensure columns exist
                    available_keys = [c for c in key_cols if c in inv_df.columns and c in mapping_df.columns]
                    
                    if not available_keys:
                        available_keys = ['SKU']

                    # Clean whitespace
                    for col in available_keys:
                        inv_df[col] = inv_df[col].astype(str).str.strip()
                        mapping_df[col] = mapping_df[col].astype(str).str.strip()

                    # Mapping DF se duplicate drop karke composite lookup create karte hain
                    mapping_subset = mapping_df[available_keys + ['uni sku']].drop_duplicates(subset=available_keys)

                    # Merge to get 'uni sku'
                    merged = inv_df.merge(mapping_subset, on=available_keys, how='left')

                    # Datatable me clean Sku Code aur Available (ATP)
                    dt_df['Sku Code'] = dt_df['Sku Code'].astype(str).str.strip()
                    dt_lookup = dt_df.drop_duplicates(subset=['Sku Code']).set_index('Sku Code')['Available (ATP)'].to_dict()

                    # uni sku ke match par Available (ATP) lana
                    merged['matched_atp'] = merged['uni sku'].map(dt_lookup)

                    # inventory_export me 'On hand (new)' column update karna
                    # Jahan match mile wahan ATP value, baaki jagah blank ya 0
                    if 'On hand (new)' not in inv_df.columns:
                        inv_df['On hand (new)'] = None

                    inv_df['On hand (new)'] = merged['matched_atp']

                    # Summary stats
                    matched_count = inv_df['On hand (new)'].notna().sum()
                    total_count = len(inv_df)

                    st.success(f"Processing complete! **{matched_count}** out of **{total_count}** rows successfully matched.")

                    # Preview table
                    st.write("### Preview (Updated Data):")
                    preview_cols = ['Handle', 'Option1 Value', 'Option2 Value', 'SKU', 'On hand (current)', 'On hand (new)']
                    show_cols = [c for c in preview_cols if c in inv_df.columns]
                    st.dataframe(inv_df[show_cols].head(20), use_container_width=True)

                    # Download button
                    csv_buffer = io.StringIO()
                    inv_df.to_csv(csv_buffer, index=False)
                    updated_csv = csv_buffer.getvalue().encode('utf-8')

                    st.download_button(
                        label="📥 Download Updated inventory_export.csv",
                        data=updated_csv,
                        file_name="inventory_export_updated.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error during processing: {str(e)}")
