import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

# --- CSS FOR PRINTING ---
st.markdown("""
    <style>
    @media print {
        .stActionButton, .stSidebar, [data-testid="stHeader"], [data-testid="stFileUploadDropzone"] { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0 !important; margin: 0 !important; }
        table { width: 100% !important; table-layout: fixed !important; border-collapse: collapse !important; }
        th, td { font-size: 9px !important; word-wrap: break-word !important; border: 1px solid #ccc !important; padding: 4px !important; }
        .js-plotly-plot { width: 100% !important; }
        .element-container { page-break-inside: avoid; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("Professional analysis of steel elongation and paint usage efficiency.")

# =============================
# 1. FILE UPLOAD
# =============================
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        df_temp.columns = df_temp.columns.str.replace(r'\s+', '', regex=True)
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]
        st.session_state['saved_data'] = df_temp
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error: {e}")

# =============================
# 2. DATA PROCESSING (CORRECTED LOGIC)
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    # Column Mapping
    order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
    cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
    ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

    try:
        with st.spinner('Calculating with high-precision logic...'):
            # STEP 1: Aggregate by Mother Coil
            # We take 'first' for CGL length because it's the same for all baby coils from one mother.
            step1_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_col]).agg(step1_agg).reset_index()

            # STEP 2: Aggregate by Order
            # Now we can safely SUM the mother lengths and baby lengths.
            step2_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            # Final Calculations
            df_summary.rename(columns={cgl_len: 'CGL_Total', ccl_len: 'CCL_Total'}, inplace=True)
            df_summary['Delta_m'] = df_summary['CCL_Total'] - df_summary['CGL_Total']
            df_summary['Thick_Var_mm'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_m']

        # =============================
        # 3. TABLES (Print-friendly)
        # =============================
        st.subheader("1. Order Summary")
        disp_cols = [order_col, 'CGL_Total', 'CCL_Total', 'Delta_m', 'Thick_Var_mm', 'Extra_Area_m2']
        df_summary_disp = df_summary[disp_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_disp.columns = ['Order', 'CGL Length (m)', 'CCL Length (m)', 'Delta (m)', 'Thick Var (mm)', 'Extra Area (m2)']
        st.table(df_summary_disp.head(30)) 

        st.divider()
        st.subheader("2. Baby Coil Details")
        selected_order = st.selectbox("Select Order Number:", options=df[order_col].unique())

        df_detail_final = None
        if selected_order:
            row = df_summary[df_summary[order_col] == selected_order].iloc[0]
            c1, c2, c3 = st.columns(3)
            c1.metric("CGL Total", f"{row['CGL_Total']:,.0f} m")
            c2.metric("CCL Total", f"{row['CCL_Total']:,.0f} m")
            c3.metric("Delta", f"{row['Delta_m']:,.0f} m", delta=f"{row['Delta_m']:,.0f} m", delta_color="inverse")
            
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            d_cols = [mother_col, baby_col, cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len]
            df_detail_final = df_detail[d_cols].sort_values(by=mother_col).copy()
            df_detail_final.columns = ['Mother Coil', 'Baby Coil', 'CGL Thick', 'CCL Thick', 'Var (mm)', 'CCL Len (m)']
            st.table(df_detail_final)

        # =============================
        # 4. VISUAL ANALYSIS (Your Original Designs)
        # =============================
        st.divider()
        st.subheader("3. Visual Analysis: Length & Extra Area")
        fig1 = px.bar(df_summary_disp, x='Order', y='Extra Area (m2)', color='Delta (m)', text='Extra Area (m2)', title="Extra Painted Area per Order")
        st.plotly_chart(fig1, use_container_width=True)

        st.subheader("4. Distribution of Length Variance")
        fig2 = px.histogram(df_summary_disp, x='Delta (m)', nbins=20, title="Distribution of Delta Length (CCL - CGL)")
        st.plotly_chart(fig2, use_container_width=True)

        st.subheader("5. Thickness vs Length Variance")
        fig3 = px.scatter(df_summary, x='Thick_Var_mm', y='Delta_m', color='Extra_Area_m2', hover_data=[order_col], title="Thickness Variance vs Length Delta")
        st.plotly_chart(fig3, use_container_width=True)

        # =============================
        # 5. EXCEL EXPORT
        # =============================
        excel_data = io.BytesIO()
        with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
            df_summary_disp.to_excel(writer, sheet_name='Summary', index=False)
            if df_detail_final is not None:
                df_detail_final.to_excel(writer, sheet_name='Details', index=False)
        st.sidebar.download_button("📥 Download Excel Report", data=excel_data.getvalue(), file_name="Paint_Yield_Report.xlsx")

    except Exception as e:
        st.error(f"Logic Error: {e}")
else:
    st.info("👆 Please upload your master data file.")
