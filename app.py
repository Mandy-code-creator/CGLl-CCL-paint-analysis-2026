import streamlit as st
import pandas as pd
import plotly.express as px
import io

# ----------------------------
# PAGE CONFIG
# ----------------------------
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")
st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between **Galvanizing (CGL)** mother coils 
and **Color Coating (CCL)** baby coils to estimate hidden paint loss.
""")

# ----------------------------
# FILE UPLOAD
# ----------------------------
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx','csv'])

if uploaded_file:
    try:
        if isinstance(uploaded_file, str):
            df = pd.read_excel(uploaded_file) if uploaded_file.endswith('.xlsx') else pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)

        # Clean columns
        df.columns = df.columns.str.replace(r'\s+','',regex=True)
        df = df.loc[:, ~df.columns.duplicated()]
        st.session_state['saved_data'] = df
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# ----------------------------
# DATA PROCESSING
# ----------------------------
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()

    # Columns mapping
    order_col = "訂單號碼"
    mother_col = "投入鋼捲號碼"
    baby_col = "產出鋼捲號碼"
    cgl_len_col = "镀锌測長度"
    ccl_len_col = "實測長度"
    ccl_width_col = "實測寬度"
    coating_col = "塗層厚度_mm"  # optional

    # ----------------------------
    # Convert units if necessary (assume input mm)
    df[cgl_len_col] = df[cgl_len_col] / 1000
    df[ccl_len_col] = df[ccl_len_col] / 1000
    df[ccl_width_col] = df[ccl_width_col] / 1000

    # ----------------------------
    # Step1: Aggregate per mother coil
    step1_agg = {
        ccl_len_col: 'sum',      # sum all baby coils
        ccl_width_col: 'mean',
        cgl_len_col: 'first'     # take actual mother coil length, not sum
    }
    df_step1 = df.groupby([order_col, mother_col]).agg(step1_agg).reset_index()

    # Step2: Aggregate per order
    step2_agg = {
        ccl_len_col: 'sum',
        ccl_width_col: 'mean',
        cgl_len_col: 'sum'  # sum mother coils per order
    }
    df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()
    df_summary.rename(columns={cgl_len_col:'CGL_Total_Length', ccl_len_col:'CCL_Total_Length'}, inplace=True)

    # ----------------------------
    # Delta Length
    df_summary['Delta_Length'] = df_summary['CCL_Total_Length'] - df_summary['CGL_Total_Length']

    # Extra Area
    df_summary['Extra_Area_m2'] = (df_summary['Delta_Length'] * df_summary[ccl_width_col]).clip(lower=0)

    # Paint Mass if coating thickness exists
    rho_paint = 1200  # kg/m3
    if coating_col in df.columns:
        df_summary['Paint_Volume_m3'] = df_summary['Extra_Area_m2'] * (df[coating_col].mean()/1000)
        df_summary['Paint_Mass_kg'] = df_summary['Paint_Volume_m3'] * rho_paint
    else:
        df_summary['Paint_Volume_m3'] = None
        df_summary['Paint_Mass_kg'] = None

    # ----------------------------
    # ORDER SUMMARY DISPLAY
    st.subheader("1. Order Summary")
    display_cols = ['CGL_Total_Length','CCL_Total_Length','Delta_Length','Extra_Area_m2','Paint_Mass_kg']
    df_display = df_summary[display_cols].copy()
    df_display.columns = ['CGL Total Length (m)','CCL Total Length (m)','Delta Length (m)','Extra Area (m2)','Paint (kg)']
    st.dataframe(df_display, use_container_width=True)
    st.divider()

    # ----------------------------
    # BABY COIL DETAILS
    st.subheader("2. Baby Coil Details")
    orders = df[order_col].dropna().unique().tolist()
    selected_order = st.selectbox("Select Order:", orders)

    if selected_order:
        df_detail = df[df[order_col]==selected_order].copy()
        df_detail['Delta_Length'] = df_detail[ccl_len_col] - df_detail[cgl_len_col]
        df_detail['Extra_Area_m2'] = (df_detail['Delta_Length'] * df_detail[ccl_width_col]).clip(lower=0)
        if coating_col in df.columns:
            df_detail['Paint_Mass_kg'] = df_detail['Extra_Area_m2'] * (df[coating_col].mean()/1000) * rho_paint
        else:
            df_detail['Paint_Mass_kg'] = None
        st.dataframe(df_detail[[mother_col,baby_col,cgl_len_col,ccl_len_col,'Delta_Length','Extra_Area_m2','Paint_Mass_kg']], use_container_width=True)

    # ----------------------------
    # VISUALS
    st.subheader("3. Visual Analysis")
    fig1 = px.bar(df_display, x=df_display.index, y='Extra Area (m2)', text='Extra Area (m2)', title="Extra Painted Area per Order")
    st.plotly_chart(fig1, use_container_width=True)

    fig2 = px.histogram(df_display, x='Delta Length (m)', nbins=20, title="Distribution of Delta Length")
    st.plotly_chart(fig2, use_container_width=True)

    # ----------------------------
    # EXPORT
    st.subheader("4. Export Reports")
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df_display.to_excel(writer, sheet_name='Order Summary', index=False)
        if selected_order:
            df_detail.to_excel(writer, sheet_name='Baby Coil Details', index=False)

    st.download_button(
        label="📥 Download Excel Report",
        data=excel_buffer.getvalue(),
        file_name=f"Paint_Yield_Report_{selected_order if selected_order else 'All'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.info("💡 To save PDF, press Ctrl+P (or Cmd+P) and select 'Save as PDF'.")
else:
    st.info("👆 Please upload your master data file (.xlsx or .csv).")
