import streamlit as st
import pandas as pd
import plotly.express as px
import io

# ----------------------------
# PAGE CONFIGURATION
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
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        if isinstance(uploaded_file, str):
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.endswith('.xlsx') else pd.read_csv(uploaded_file)
        else:
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)

        # Clean column names
        df_temp.columns = df_temp.columns.str.replace(r'\s+', '', regex=True)
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]

        st.session_state['saved_data'] = df_temp
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# ----------------------------
# DATA PROCESSING
# ----------------------------
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()

    # ERP Column Mapping
    order_col = "訂單號碼"
    mother_coil_col = "投入鋼捲號碼"
    baby_coil_col = "產出鋼捲號碼"
    cgl_len = "镀锌測長度"
    ccl_len = "實測長度"
    ccl_width = "實測寬度"

    # Optional: nếu có lớp sơn phủ
    coating_thick_col = "塗層厚度_mm"  # thêm cột này nếu có

    try:
        with st.spinner('Processing and aggregating data...'):
            # Step 1: Aggregate per mother coil
            step1_agg = {
                ccl_len: 'sum',
                ccl_width: 'mean',
                cgl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_coil_col]).agg(step1_agg).reset_index()

            # Step 2: Aggregate per order
            step2_agg = {
                ccl_len: 'sum',
                ccl_width: 'mean',
                cgl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()
            df_summary.rename(columns={cgl_len: 'CGL_Total_Length', ccl_len: 'CCL_Total_Length'}, inplace=True)

            # Delta Length
            df_summary['Delta_Length'] = df_summary['CCL_Total_Length'] - df_summary['CGL_Total_Length']

            # Extra Area based on Delta Length
            df_summary['Extra_Area_m2'] = (df_summary['Delta_Length'] * df_summary[ccl_width] / 1000).clip(lower=0)  # diện tích dương

            # Paint usage only if coating thickness exists
            rho_paint = 1200  # kg/m3
            if coating_thick_col in df.columns:
                df_summary['Paint_Volume_m3'] = df_summary['Extra_Area_m2'] * (df_summary[coating_thick_col] / 1000)
                df_summary['Paint_Mass_kg'] = df_summary['Paint_Volume_m3'] * rho_paint
            else:
                df_summary['Paint_Volume_m3'] = None
                df_summary['Paint_Mass_kg'] = None

        # ----------------------------
        # ORDER SUMMARY
        # ----------------------------
        st.subheader("1. Order Summary")
        summary_display_cols = ['CGL_Total_Length', 'CCL_Total_Length', 'Delta_Length', 'Extra_Area_m2', 'Paint_Mass_kg']
        df_summary_display = df_summary.copy()
        df_summary_display = df_summary_display.rename(columns={
            'CGL_Total_Length': 'CGL Total Length (m)',
            'CCL_Total_Length': 'CCL Total Length (m)',
            'Delta_Length': 'Delta Length (m)',
            'Extra_Area_m2': 'Extra Area (m2)',
            'Paint_Mass_kg': 'Paint (kg)'
        })
        st.dataframe(df_summary_display, use_container_width=True)
        st.divider()

        # ----------------------------
        # BABY COIL DETAILS
        # ----------------------------
        st.subheader("2. Baby Coil Details")
        st.markdown("Select an order to view its breakdown.")
        order_list = df[order_col].dropna().unique().tolist()
        selected_order = st.selectbox("Select Order Number:", options=order_list)

        if selected_order:
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Delta_Length'] = df_detail[ccl_len] - df_detail[cgl_len]
            df_detail['Extra_Area_m2'] = (df_detail['Delta_Length'] * df_detail[ccl_width] / 1000).clip(lower=0)
            if coating_thick_col in df_detail.columns:
                df_detail['Paint_Mass_kg'] = df_detail['Extra_Area_m2'] * (df_detail[coating_thick_col]/1000) * rho_paint
            else:
                df_detail['Paint_Mass_kg'] = None

            st.dataframe(df_detail[[mother_coil_col, baby_coil_col, cgl_len, ccl_len, 'Delta_Length', 'Extra_Area_m2', 'Paint_Mass_kg']], use_container_width=True)

        # ----------------------------
        # VISUAL ANALYSIS
        # ----------------------------
        st.subheader("3. Visual Analysis")
        fig1 = px.bar(df_summary_display, x=df_summary.index, y='Extra Area (m2)', text='Extra Area (m2)', title="Extra Painted Area per Order")
        st.plotly_chart(fig1, use_container_width=True)

        fig2 = px.histogram(df_summary_display, x='Delta Length (m)', nbins=20, title="Distribution of Delta Length (CCL - CGL)")
        st.plotly_chart(fig2, use_container_width=True)

        # ----------------------------
        # EXPORT REPORT
        # ----------------------------
        st.subheader("4. Export Reports")
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df_summary_display.to_excel(writer, sheet_name='Order Summary', index=False)
            if selected_order:
                df_detail.to_excel(writer, sheet_name='Baby Coil Details', index=False)

        st.download_button(
            label="📥 Download Excel Report",
            data=excel_buffer.getvalue(),
            file_name=f"Paint_Yield_Report_{selected_order if selected_order else 'All'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("💡 To save the dashboard as PDF, press Ctrl+P (or Cmd+P) and select 'Save as PDF'.")

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")

else:
    st.info("👆 Please upload your master data file (.xlsx or .csv) to begin.")
