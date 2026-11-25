import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Payment Reco & Profit Analyzer", layout="wide")
st.title("📊 Marketplace Payment Reconciliation & SKU Action System")

uploaded_file = st.file_uploader("Upload Order / Payment Reco CSV or ZIP", type=["csv", "zip"])

if uploaded_file:
    df = pd.read_csv(uploaded_file, low_memory=False)
    
    def clean_mp(ch):
        s = str(ch)
        if "Myntra" in s: return "Myntra"
        if "Amazon" in s: return "Amazon"
        if "Ajio" in s: return "Ajio"
        return s

    def bucket_status(s):
        if s == "Delivered": return "Delivered (Success)"
        if s in ["In Transit", "Ready to ship", "Packed", "New"]: return "In-Transit / Active"
        if s in ["Cancelled", "Cancel Init", "Cancelled Return Received"]: return "Cancelled / RTO"
        if s in ["Return Init", "Return Received"]: return "Customer Return"
        return str(s)

    df["Marketplace"] = df["Channel Name"].apply(clean_mp)
    df["Status_Bucket"] = df["Order Status"].apply(bucket_status)
    
    for col in ["Total", "Expected Payment", "Cost of Goods", "Calc Last Purchase Price", "Qty", "Expected Profit", "Return Bad Qty", "Reverse Shipping Fee"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["Effective_COGS"] = np.where(df["Cost of Goods"] > 0, df["Cost of Goods"], df["Calc Last Purchase Price"] * df["Qty"])
    df["Return_Charge"] = np.where(df["Expected Payment"] < 0, -df["Expected Payment"], df["Reverse Shipping Fee"])
    df["Bad_Return_Loss"] = df["Return Bad Qty"] * df["Calc Last Purchase Price"]

    # 1. Marketplace Summary
    summary = []
    for mp, grp in df.groupby("Marketplace"):
        d = grp[grp["Status_Bucket"] == "Delivered (Success)"]
        r = grp[grp["Status_Bucket"] == "Customer Return"]
        c = grp[grp["Status_Bucket"] == "Cancelled / RTO"]
        a = grp[grp["Status_Bucket"] == "In-Transit / Active"]
        
        ret_chg = r["Return_Charge"].sum() + c["Return_Charge"].sum()
        bad_loss = grp["Bad_Return_Loss"].sum()
        net_prof = d["Expected Profit"].sum() - ret_chg - bad_loss
        
        summary.append({
            "Marketplace": mp,
            "Total Orders": len(grp),
            "Delivered": len(d),
            "In-Transit": len(a),
            "Cancelled": len(c),
            "Returns": len(r),
            "Return %": round((len(r) / (len(d) + len(r)) * 100) if (len(d) + len(r)) > 0 else 0, 1),
            "Success Payout (₹)": round(d["Expected Payment"].sum(), 2),
            "Success Cost (₹)": round(d["Effective_COGS"].sum(), 2),
            "Return Charges (₹)": round(ret_chg, 2),
            "Return Chg / Success (₹)": round(ret_chg / len(d), 2) if len(d) > 0 else 0,
            "True Net Profit (₹)": round(net_prof, 2),
            "Unpaid Delivered": len(d[d["Payment Received"].str.lower() != "yes"])
        })

    sum_df = pd.DataFrame(summary)
    st.subheader("1. Marketplace Summary (Ek Nazar Mein)")
    st.dataframe(sum_df, use_container_width=True)

    # 2. SKU Action Recommendations
    sku_df = df.groupby(["Marketplace", "Product Sku Code", "Product Name"]).agg(
        Delivered=("Status_Bucket", lambda x: (x == "Delivered (Success)").sum()),
        Returns=("Status_Bucket", lambda x: (x == "Customer Return").sum()),
        Unit_COGS=("Calc Last Purchase Price", "max"),
        Avg_Price=("Total", "mean"),
        Del_Profit=("Expected Profit", lambda x: x[df.loc[x.index, "Status_Bucket"] == "Delivered (Success)"].sum()),
        Ret_Charges=("Return_Charge", "sum"),
        Bad_Loss=("Bad_Return_Loss", "sum")
    ).reset_index()

    sku_df["Return_Rate_%"] = np.where((sku_df["Delivered"] + sku_df["Returns"]) > 0,
                                       (sku_df["Returns"] / (sku_df["Delivered"] + sku_df["Returns"])) * 100, 0).round(1)
    sku_df["Net_Profit"] = (sku_df["Del_Profit"] - sku_df["Ret_Charges"] - sku_df["Bad_Loss"]).round(2)
    sku_df["Action"] = np.where(sku_df["Net_Profit"] < 0, "🔴 LOSS: Increase Price / Audit",
                       np.where(sku_df["Return_Rate_%"] >= 40, "🟠 HIGH RETURN: Check Size/Quality", "🟢 HEALTHY"))

    st.subheader("2. SKU Profitability & Price Action Suggestions")
    st.dataframe(sku_df.sort_values("Net_Profit"), use_container_width=True)
