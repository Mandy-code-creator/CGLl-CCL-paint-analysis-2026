import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Length Variance Dashboard", layout="wide")

st.title("Length Input vs Output Analysis")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Preview Data")
    st.dataframe(df.head())

    # ===== SELECT COLUMNS =====
    cgl_col = st.selectbox("Select Input Length Column (CGL)", df.columns)
    ccl_col = st.selectbox("Select Output Length Column (CCL)", df.columns)

    # ===== CLEAN DATA =====
    df[cgl_col] = pd.to_numeric(df[cgl_col], errors='coerce')
    df[ccl_col] = pd.to_numeric(df[ccl_col], errors='coerce')

    total_input = df[cgl_col].sum()
    total_output = df[ccl_col].sum()

    variance = total_output - total_input
    variance_percent = (variance / total_input * 100) if total_input != 0 else 0
    yield_percent = (total_output / total_input * 100) if total_input != 0 else 0

    st.divider()

    col1, col2, col3 = st.columns(3)

    col1.metric("Total Input Length (m)", f"{total_input:,.2f}")
    col2.metric("Total Output Length (m)", f"{total_output:,.2f}")
    col3.metric("Length Variance (m)", f"{variance:,.2f}")

    st.metric("Variance (%)", f"{variance_percent:.2f}%")
    st.metric("Length Yield (%)", f"{yield_percent:.2f}%")

    st.divider()

    # ===== BAR CHART =====
    st.subheader("Total Length Comparison")

    fig, ax = plt.subplots()
    ax.bar(["Total Input", "Total Output"], [total_input, total_output])
    ax.set_ylabel("Length (m)")
    ax.set_title("Input vs Output Length")

    st.pyplot(fig)
