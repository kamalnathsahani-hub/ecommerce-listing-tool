import streamlit as st

st.set_page_config(page_title="Pricing Calculator", page_icon="💰", layout="wide")

if st.button("⬅️ Back to Home"):
    st.switch_page("app.py")

st.title("💰 Pricing & Profit Margin Calculator")
st.write("Yahan hum aapke products ki pricing aur profit calculate karne ka tool banayenge.")
