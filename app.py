import streamlit as st

st.set_page_config(page_title="E-Commerce Business Hub", page_icon="🛍️", layout="wide")

st.title("🛍️ E-Commerce Seller Dashboard")
st.write("Apne kaam ke hisaab se niche diye gaye tools mein se choose karein:")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("📦 Product Listing & SEO Tool")
    st.write("Product titles, descriptions, catalog automation aur trending SEO keywords generate karne ke liye yahan jayein.")
    if st.button("Go to Listing Tool ➡️", use_container_width=True):
        st.switch_page("pages/1_Listing_Tool.py")

with col2:
    st.subheader("💰 Pricing & Margin Calculator")
    st.write("Selling price, platform charges, shipping, GST aur apna final net profit calculate karne ke liye yahan jayein.")
    if st.button("Go to Pricing Calculator ➡️", use_container_width=True):
        st.switch_page("pages/2_Pricing_Calculator.py")
