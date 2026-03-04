import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

# --- CSS FOR PRINTING (Fixes the "Hidden Columns" and "Cut Charts" issues) ---
st.markdown("""
    <style>
    @media print {
        /* Hide UI elements for a clean report */
        .stActionButton, .stSidebar, .stHeader, [data-testid="stFileUploadDropzone"], [data-testid="stHeader"] { display: none !important; }
        
        /* Force container to use full page width */
        .main .block-container { max-width: 100% !important; padding: 0 !important; }
        
        /* Make tables print fully without scrollbars */
        table { width: 100% !important; font-size: 10px !important; table-layout: fixed !important; border-collapse: collapse; }
        th, td { word-wrap: break-word !important; border: 1px solid #ccc !important; padding: 4px !important; }
        div[data-testid="stTable"], .element-container { overflow: visible !important; height: auto !important; }
        
        /* Ensure Plotly charts resize for the page */
        .js-plotly-plot { max-width: 100% !important; height: auto !important; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Paint Yield Analysis Report")

# =============================
# 1. FILE UPLOAD
# =============================
uploaded_file = st.file_uploader("Upload Data File", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        
        # Data Cleaning
        df_temp.columns = df_temp.columns.str.replace(r'\s+', '', regex=True)
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]
        st.session_state['saved_data'] = df_temp
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error: {e}")

# =============================
# 2. PROCESSING & DISPLAY (Strictly Your Logic)
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
    cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
    ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

    try:
        # Two-step aggregation logic
        step1 = df.groupby([order_col, mother_col]).agg({
            cgl_thick: 'first', cgl_width: 'first', cgl_len: 'first',
            ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
        }).reset_index()

        df_summary = step1.groupby(order_col).agg({
            cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
            ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
        }).reset_index()

        df_summary.rename(columns={cgl_len: 'CGL_Len', ccl_len: 'CCL_Len'}, inplace=True)
        df_summary['Delta'] = df_summary['CCL_Len'] - df_summary['CGL_Len']
        df_summary['Thick_Var'] = df_summary[ccl_thick] - df_summary[cgl_thick]
        df_summary['Extra_Area'] = (df_summary[ccl_width] / 1000) * df_summary['Delta']

        # --- SUMMARY TABLE (st.table for full visibility when printing) ---
        st.subheader("1. Summary Table")
        disp_summary = df_summary.sort_values(by='Extra_Area', ascending=False).copy()
        disp_summary = disp_summary[[order_col, 'CGL_Len', 'CCL_Len', 'Delta', 'Thick_Var', 'Extra_Area']]
        disp_summary.columns = ['Order', 'CGL(m)', 'CCL(m)', 'Delta(m)', 'Thick Var', 'Area(m2)']
        
        # Using st.table to prevent scrollbars in print
        st.table(disp_summary.head(30))

        # --- DETAIL DRILL DOWN ---
        st.divider()
        st.subheader("2. Detailed Breakdown")
        selected_order = st.selectbox("Select Order Number:", options=df[order_col].unique())
        
        df_detail_final = None
        if selected_order:
            row = df_summary[df_summary[order_col] == selected_order].iloc[0]
            c1, c2, c3 = st.columns(3)
            c1.metric("CGL Total", f"{row['CGL_Len']:,.0f} m")
            c2.metric("CCL Total", f"{row['CCL_Len']:,.0f} m")
            c3.metric("Delta", f"{row['Delta']:,.0f} m", delta=f"{row['Delta']:,.0f} m", delta_color="inverse")
            
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Var'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            df_detail_final = df_detail[[mother_col, baby_col, cgl_thick, ccl_thick, 'Var', ccl_len]].copy()
            df_detail_final.columns = ['Mother', 'Baby', 'CGL-T', 'CCL-T', 'Diff', 'Len(m)']
            st.table(df_detail_final)

        # --- VISUAL ANALYSIS (Your Original Chart Designs) ---
        st.divider()
        st.subheader("3. Visual Analysis")
        
        # Original Bar Chart
        fig1 = px.bar(disp_summary.head(10), x='Order', y='Area(m2)', color='Delta(m)', text='Area(m2)', title="Extra Painted Area per Order")
        st.plotly_chart(fig1, use_container_width=True)

        # Original Histogram
        fig2 = px.histogram(disp_summary, x='Delta(m)', nbins=20, title="Distribution of Elongation (CCL - CGL)")
        st.plotly_chart(fig2, use_container_width=True)

        # Original Scatter Plot
        fig3 = px.scatter(df_summary, x='Thick_Var', y='Delta', color='Extra_Area', hover_data=[order_col], title="Correlation: Thickness vs Length")
        st.plotly_chart(fig3, use_container_width=True)

        # =============================
        # 3. EXCEL EXPORT
        # =============================
        excel_data = io.BytesIO()
        with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
            disp_summary.to_excel(writer, sheet_name='Summary', index=False)
            if df_detail_final is not None:
                df_detail_final.to_excel(writer, sheet_name='Details', index=False)
        st.sidebar.download_button("📥 Download Excel Report", data=excel_data.getvalue(), file_name="Paint_Yield_Report.xlsx")

    except Exception as e:
        st.error(f"Processing Error: {e}")

else:
    st.info("👆 Please upload your master data file to start.")
