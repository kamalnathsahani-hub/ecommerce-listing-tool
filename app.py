import streamlit as st

st.set_page_config(page_title="E-Commerce Business Hub", page_icon="🛍️", layout="wide")

# Session state mein access track karne ke liye
if "pricing_access" not in st.session_state:
    st.session_state.pricing_access = False

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
    st.subheader("💰 Pricing & Margin Calculator 🔒")
    st.write("Selling price, platform charges, shipping, GST aur apna final net profit calculate karne ke liye yahan jayein.")
    
    # Agar user ne abhi tak password nahi dala hai
    if not st.session_state.pricing_access:
        access_code = st.text_input("Access Code / Password dalein:", type="password", key="calc_pass")
        if st.button("Unlock & Go to Calculator 🔓", use_container_width=True):
            # "MYSECRET123" ki jagah apna manpasand password set karein (ya st.secrets use karein)
            if access_code == "MYSECRET123":
                st.session_state.pricing_access = True
                st.switch_page("pages/2_Pricing_Calculator.py")
            else:
                st.error("❌ Galat password! Sirf authorized log hi isey dekh sakte hain.")
    else:
        # Agar password pehle hi sahi daal diya hai
        if st.button("Go to Pricing Calculator ➡️", use_container_width=True):
            st.switch_page("pages/2_Pricing_Calculator.py")
