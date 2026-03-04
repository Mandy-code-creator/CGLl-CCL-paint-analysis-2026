import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")
st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between **Galvanizing (CGL)** mother coils 
and **Color Coating (CCL)** baby coils to estimate hidden paint loss.
""")

# =============================
# 1. FILE UPLOAD & PRE-PROCESSING
# =============================
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        if isinstance(uploaded_file, str):
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.endswith('.xlsx') else pd.read_csv(uploaded_file)
        else:
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)

        # CLEANING: Remove whitespaces and handle duplicates in column names
        df_temp.columns = df_temp.columns.str.replace(r'\s+', '', regex=True)
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]

        st.session_state['saved_data'] = df_temp
        st.success("Data loaded and cleaned successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# =============================
# 2. DATA ANALYSIS ENGINE
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()

    # Column Mapping
    order_col = "訂單號碼"
    mother_coil_col = "投入鋼捲號碼"
    baby_coil_col = "產出鋼捲號碼"
    cgl_thick = "镀锌實測厚度"
    cgl_width = "镀锌測寬度"
    cgl_len = "镀锌測長度"
    ccl_thick = "實測厚度"
    ccl_width = "實測寬度"
    ccl_len = "實測長度"

    try:
        with st.spinner('Calculating yields...'):
            # Step 1: Aggregate by Mother Coil
            step1_agg = {
                cgl_thick: 'first', cgl_width: 'first', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_coil_col]).agg(step1_agg).reset_index()

            # Step 2: Aggregate by Order Number
            step2_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            df_summary.rename(columns={cgl_len: 'CGL_Total_Length', ccl_len: 'CCL_Total_Length'}, inplace=True)
            df_summary['Delta_Length'] = df_summary['CCL_Total_Length'] - df_summary['CGL_Total_Length']
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_Length']

        # --- SECTION 1: ORDER SUMMARY ---
        st.subheader("1. Order Summary")
        summary_display_cols = [order_col, 'CGL_Total_Length', 'CCL_Total_Length', 'Delta_Length', 'Thickness_Variance', 'Extra_Area_m2']
        df_summary_display = df_summary[summary_display_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_display.columns = ['Order Number', 'CGL Length (m)', 'CCL Length (m)', 'Delta Length (m)', 'Thickness Var (mm)', 'Extra Area (m2)']
        st.dataframe(df_summary_display, use_container_width=True)

        # --- SECTION 2: DETAILS ---
        st.divider()
        st.subheader("2. Detailed Analysis per Order")
        order_list = df[order_col].dropna().unique().tolist()
        selected_order = st.selectbox("Search for an Order:", options=order_list)

        df_detail_final = None
        if selected_order:
            order_totals = df_summary[df_summary[order_col] == selected_order].iloc[0]
            m1, m2, m3 = st.columns(3)
            m1.metric("CGL Total", f"{order_totals['CGL_Total_Length']:,.0f} m")
            m2.metric("CCL Total", f"{order_totals['CCL_Total_Length']:,.0f} m")
            m3.metric("Elongation", f"{order_totals['Delta_Length']:,.0f} m", delta=f"{order_totals['Delta_Length']:,.0f} m", delta_color="inverse")
            
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            detail_cols = [mother_coil_col, baby_coil_col, cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len]
            df_detail_final = df_detail[detail_cols].copy()
            df_detail_final.columns = ['Mother Coil ID', 'Baby Coil ID', 'CGL Thick', 'CCL Thick', 'Diff (mm)', 'Length (m)']
            st.dataframe(df_detail_final, use_container_width=True)

        # --- SECTION 3: CHARTS ---
        st.divider()
        st.subheader("3. Executive Dashboard")
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.bar(df_summary_display.head(10), x='Order Number', y='Extra Area (m2)', title="Top 10 High Waste Orders"), use_container_width=True)
        with c2:
            st.plotly_chart(px.scatter(df_summary, x='Thickness_Variance', y='Delta_Length', color='Extra_Area_m2', title="Thickness vs Elongation Correlation"), use_container_width=True)

        # =============================
        # 4. EXPORT (SỬA LỖI TẠI ĐÂY)
        # =============================
        st.sidebar.header("Export Report")
        
        # EXCEL EXPORT (Cách an toàn nhất)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_summary_display.to_excel(writer, sheet_name='Summary', index=False)
            if df_detail_final is not None:
                df_detail_final.to_excel(writer, sheet_name='Order_Details', index=False)
        
        st.sidebar.download_button(
            label="📥 Download Excel Report",
            data=buffer.getvalue(),
            file_name=f"Paint_Yield_{selected_order}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.sidebar.info("To save as PDF: Use your browser's Print (Ctrl+P) and select 'Save as PDF'.")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("👆 Please upload your data file.")
