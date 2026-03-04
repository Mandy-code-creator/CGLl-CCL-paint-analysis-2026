import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Process Yield Dashboard", layout="wide")
st.title("Process Yield & Paint Efficiency Dashboard")

uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx','csv'])

if uploaded_file:
    if uploaded_file.name.endswith('.xlsx'):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file)

    df.columns = df.columns.str.replace(r'\s+','',regex=True)

    # Column mapping
    order_col = "訂單號碼"
    cgl_len_col = "镀锌測長度"
    ccl_len_col = "實測長度"
    width_col = "實測寬度"
    coating_col = "塗層厚度_mm"

    # Convert mm → m
    df[cgl_len_col] = df[cgl_len_col] / 1000
    df[ccl_len_col] = df[ccl_len_col] / 1000
    df[width_col] = df[width_col] / 1000

    # Aggregate per order
    summary = df.groupby(order_col).agg(
        CGL_Total = (cgl_len_col,'sum'),
        CCL_Total = (ccl_len_col,'sum'),
        Avg_Width = (width_col,'mean')
    ).reset_index()

    # 1️⃣ Length Yield
    summary["Length_Yield_%"] = (summary["CCL_Total"] / summary["CGL_Total"]) * 100

    # 2️⃣ Mechanical Loss
    summary["Mechanical_Loss_m"] = summary["CGL_Total"] - summary["CCL_Total"]

    # 3️⃣ Paint theoretical usage (based on OUTPUT area)
    summary["CCL_Area_m2"] = summary["CCL_Total"] * summary["Avg_Width"]

    rho_paint = 1200  # kg/m3
    if coating_col in df.columns:
        avg_coating_m = df[coating_col].mean() / 1000
        summary["Paint_Theoretical_kg"] = summary["CCL_Area_m2"] * avg_coating_m * rho_paint
    else:
        summary["Paint_Theoretical_kg"] = None

    st.subheader("Order Summary")
    st.dataframe(summary, use_container_width=True)

    # Visualization
    fig = px.bar(summary, x=order_col, y="Length_Yield_%",
                 title="Length Yield per Order (%)")
    st.plotly_chart(fig, use_container_width=True)

    # Export
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary.to_excel(writer, index=False)

    st.download_button(
        "Download Excel Report",
        buffer.getvalue(),
        "Process_Yield_Report.xlsx"
    )
