import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import io
import re

st.set_page_config(page_title="Payment Reco & SKU Action System", layout="wide")
st.title("📊 Marketplace Payment Reconciliation & SKU Action System")

uploaded_file = st.file_uploader("Upload Order / Payment Reco CSV or ZIP", type=["csv", "zip"])

def load_uploaded_data(file_obj):
    file_name = file_obj.name.lower()
    if file_name.endswith(".zip"):
        with zipfile.ZipFile(file_obj, "r") as z:
            csv_files = [f for f in z.namelist() if f.lower().endswith(".csv") and not f.startswith("__MACOSX")]
            if not csv_files:
                st.error("ZIP folder ke andar koi .csv file nahi mili.")
                st.stop()
            with z.open(csv_files[0]) as f:
                return pd.read_csv(f, low_memory=False)
    else:
        return pd.read_csv(file_obj, low_memory=False)

if uploaded_file:
    df = load_uploaded_data(uploaded_file)

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

    def get_parent_style(sku):
        return re.sub(r"-(28|30|32|34|36|38|40|42|XS|S|M|L|XL|XXL|3XL)$", "", str(sku), flags=re.IGNORECASE)

    df["Marketplace"] = df["Channel Name"].apply(clean_mp)
    df["Status_Bucket"] = df["Order Status"].apply(bucket_status)
    df["Parent_Style"] = df["Product Sku Code"].fillna("UNKNOWN").apply(get_parent_style)

    num_cols = [
        "Total", "Expected Payment", "Cost of Goods", "Calc Last Purchase Price",
        "Qty", "Expected Profit", "Return Good Qty", "Return Bad Qty",
        "Reverse Shipping Fee", "Settlement Amount"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        else:
            df[col] = 0.0

    df["Effective_COGS"] = np.where(df["Cost of Goods"] > 0, df["Cost of Goods"], df["Calc Last Purchase Price"] * df["Qty"])
    df["Return_Charge"] = np.where(df["Expected Payment"] < 0, -df["Expected Payment"], df["Reverse Shipping Fee"])
    df["Bad_Return_Loss"] = df["Return Bad Qty"] * df["Calc Last Purchase Price"]
    df["Is_Unpaid_Delivered"] = (df["Status_Bucket"] == "Delivered (Success)") & (df["Payment Received"].fillna("No").str.lower() != "yes")

    # 1. Marketplace Summary
    summary = []
    for mp, grp in df.groupby("Marketplace"):
        d = grp[grp["Status_Bucket"] == "Delivered (Success)"]
        a = grp[grp["Status_Bucket"] == "In-Transit / Active"]
        c = grp[grp["Status_Bucket"] == "Cancelled / RTO"]
        r = grp[grp["Status_Bucket"] == "Customer Return"]

        closed_cnt = len(d) + len(r)
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
            "Return %": round((len(r) / closed_cnt * 100) if closed_cnt > 0 else 0, 1),
            "Success Gross Sale (₹)": round(d["Total"].sum(), 2),
            "Success Payout (₹)": round(d["Expected Payment"].sum(), 2),
            "Success Cost (₹)": round(d["Effective_COGS"].sum(), 2),
            "Return Charges (₹)": round(ret_chg, 2),
            "Damaged Loss (₹)": round(bad_loss, 2),
            "Return Chg / Success (₹)": round(ret_chg / len(d), 2) if len(d) > 0 else 0,
            "True Net Profit (₹)": round(net_prof, 2),
            "Unpaid Delivered": int(d["Is_Unpaid_Delivered"].sum()),
            "Unpaid Amount (₹)": round(d.loc[d["Is_Unpaid_Delivered"], "Expected Payment"].sum(), 2)
        })

    sum_df = pd.DataFrame(summary)
    st.subheader("1. Marketplace Summary (Ek Nazar Mein)")
    st.dataframe(sum_df, use_container_width=True)

    # 2. SKU Action Recommendations
    sku_df = df.groupby(["Marketplace", "Parent_Style", "Product Sku Code", "Product Name"]).agg(
        Total_Orders=("Channel Order Id", "count"),
        Delivered=("Status_Bucket", lambda x: (x == "Delivered (Success)").sum()),
        Returns=("Status_Bucket", lambda x: (x == "Customer Return").sum()),
        Cancelled=("Status_Bucket", lambda x: (x == "Cancelled / RTO").sum()),
        Unit_COGS=("Calc Last Purchase Price", "max"),
        Avg_Price=("Total", "mean"),
        Del_Profit=("Expected Profit", lambda x: x[df.loc[x.index, "Status_Bucket"] == "Delivered (Success)"].sum()),
        Ret_Charges=("Return_Charge", "sum"),
        Bad_Loss=("Bad_Return_Loss", "sum")
    ).reset_index()

    sku_df["Return_Rate_%"] = np.where(
        (sku_df["Delivered"] + sku_df["Returns"]) > 0,
        (sku_df["Returns"] / (sku_df["Delivered"] + sku_df["Returns"])) * 100, 0
    ).round(1)
    sku_df["Net_Profit"] = (sku_df["Del_Profit"] - sku_df["Ret_Charges"] - sku_df["Bad_Loss"]).round(2)
    sku_df["Suggested_Price"] = np.where(
        sku_df["Net_Profit"] < 0, np.maximum(sku_df["Avg_Price"] + 70, sku_df["Unit_COGS"] * 1.65).round(-1) - 1,
        np.where(sku_df["Return_Rate_%"] >= 40, (sku_df["Avg_Price"] + 40).round(-1) - 1, sku_df["Avg_Price"].round(0))
    )
    sku_df["Action"] = np.where(
        sku_df["Net_Profit"] < 0, "🔴 LOSS: Increase Price / Audit",
        np.where(sku_df["Return_Rate_%"] >= 40, "🟠 HIGH RETURN: Check Size/Quality", "🟢 HEALTHY")
    )

    st.subheader("2. SKU Profitability & Price Action Suggestions")
    st.dataframe(sku_df.sort_values("Net_Profit"), use_container_width=True)

    # 3. Unpaid Delivered Orders & SPF Claim Lists
    unpaid_df = df[df["Is_Unpaid_Delivered"]].copy()
    spf_df = df[df["Return Bad Qty"] > 0].copy()

    # Create All-in-One Downloadable ZIP of Excel Reports
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, frame in [
            ("1_Marketplace_Summary.xlsx", sum_df),
            ("2_Unpaid_Delivered_Action_List.xlsx", unpaid_df),
            ("3_SKU_Pricing_Actions.xlsx", sku_df.sort_values("Net_Profit")),
            ("4_SPF_Damaged_Return_Claims.xlsx", spf_df)
        ]:
            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                frame.to_excel(writer, index=False)
            zf.writestr(fname, excel_buf.getvalue())

    st.download_button(
        label="⬇️ Download All 4 Reco Reports (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Marketplace_Payment_Reco_Reports.zip",
        mime="application/zip"
    )
