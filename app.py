import streamlit as st

st.set_page_config(page_title="E-commerce Seller Suite", page_icon="🛍️", layout="wide")

st.title("🛍️ E-commerce Multi-Tool Dashboard")
st.write("Apna manpasand tool select karein:")

# 3-Column Tool Cards Layout on Home Page
col1, col2, col3 = st.columns(3)

with col1:
    st.container(border=True)
    st.markdown("### 📋 Listing Tool")
    st.write("Standard multi-marketplace listing generator.")
    st.page_link("pages/1_Listing_Tool.py", label="Open Listing Tool", icon="📋")

with col2:
    st.container(border=True)
    st.markdown("### 💰 Pricing Calculator")
    st.write("Calculate marketplace fees, margins, and selling prices.")
    st.page_link("pages/2_Pricing_Calculator.py", label="Open Calculator", icon="💰")
