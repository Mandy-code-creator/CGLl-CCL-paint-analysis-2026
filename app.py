import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- PAGE CONFIGURATION (Your Design) ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

# --- CSS FOR PRINTING (Strictly to prevent data being hidden/cut off) ---
st.markdown("""
    <style>
    @media print {
        /* Expand the main container to full width */
        .main .block-container { max-width: 100% !important; padding: 0 !important; }
        
        /* Hide interactive UI elements that don't belong in a PDF */
        .stActionButton, .stSidebar, .stHeader, [data-testid="stFileUploadDropzone"], .stSelectbox { display: none !important; }
        
        /* Force tables to show all columns and avoid internal scrollbars */
        table { width: 100% !important; font-size: 10px !important; table-layout: fixed !important; border-collapse: collapse; }
        th, td { word-wrap: break-word !important; border: 1px solid #ccc !important; padding: 4px !important; }
        
        /* Ensure charts resize to fit the printed page width */
        .js-plotly-plot { max-width: 100% !important; height: auto !important; }
        
        /* Prevent page breaks inside charts or metrics */
        .element-container { page-break-inside: avoid; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between **Galvanizing (CGL)** mother coils 
and **Color Coating (CCL)** baby coils to estimate hidden paint loss.
""")

# =============================
# 1. FILE UPLOAD (Your Logic)
# =============================
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        if isinstance(uploaded_file, str):
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.endswith('.xlsx') else pd.read_csv(uploaded_file)
        else:
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)

        # Cleaning logic to prevent "Column Not Found" errors
        df_temp.columns = df_temp.columns.str.replace(r'\s+', '', regex=True)
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]

        st.session_state['saved_data'] = df_temp
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# =============================
# 2. DATA PROCESSING (Strictly Your Logic)
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()

    # Your ERP Column Mapping
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
        with st.spinner('Processing and aggregating data...'):
            # Two-step aggregation to handle multiple mother coils
            step1_agg = {
                cgl_thick: 'first', cgl_width: 'first', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_coil_col]).agg(step1_agg).reset_index()

            step2_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            # Variance Calculations
            df_summary.rename(columns={cgl_len: 'CGL_Total_Length', ccl_len: 'CCL_Total_Length'}, inplace=True)
            df_summary['Delta_Length'] = df_summary['CCL_Total_Length'] - df_summary['CGL_Total_Length']
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_Length']

        # =============================
        # 3. ORDER SUMMARY (Your Design)
        # =============
        st.subheader("1. Order Summary")
        summary_display_cols = [order_col, 'CGL_Total_Length', 'CCL_Total_Length', 'Delta_Length', 'Thickness_Variance', 'Extra_Area_m2']
        df_summary_display = df_summary[summary_display_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_display.columns = ['Order Number', 'CGL Total Length (m)', 'CCL Total Length (m)', 'Delta Length (m)', 'Thickness Variance (mm)', 'Extra Area (m2)']
        
        # CHANGED: Using st.table for printability (no scrollbars)
        st.table(df_summary_display.head(50)) 
        st.divider()

        # =============================
        # 4. BABY COIL DETAILS (Your Design)
        # =============================
        st.subheader("2. Baby Coil Details")
        st.markdown("Select an order to view its specific breakdown and length totals.")
        order_list = df[order_col].dropna().unique().tolist()
        selected_order = st.selectbox("Select Order Number:", options=order_list)

        df_detail_display = None
        if selected_order:
            order_totals = df_summary[df_summary[order_col] == selected_order].iloc[0]
            st.markdown(f"**Length Summary for Order: {selected_order}**")
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Mother Coils Length (CGL)", f"{order_totals['CGL_Total_Length']:,.0f} m")
            col2.metric("Total Baby Coils Length (CCL)", f"{order_totals['CCL_Total_Length']:,.0f} m")
            col3.metric("Length Variance (Delta)", f"{order_totals['Delta_Length']:,.0f} m", 
                        delta=f"{order_totals['Delta_Length']:,.0f} m", delta_color="inverse")
            
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            detail_display_cols = [mother_coil_col, baby_coil_col, cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len]
            df_detail_display = df_detail[detail_display_cols].sort_values(by=mother_coil_col).copy()
            df_detail_display.columns = ['Mother Coil', 'Baby Coil', 'CGL Thickness', 'CCL Thickness', 'Thickness Variance (mm)', 'CCL Length (m)']
            
            # CHANGED: Using st.table for printability (no scrollbars)
            st.table(df_detail_display)

        # =============================
        # 5. VISUAL ANALYSIS (Strictly Your Visuals)
        # =============================
        st.divider()
        st.subheader("3. Visual Analysis: Length & Extra Area")
        fig1 = px.bar(df_summary_display, x='Order Number', y='Extra Area (m2)', color='Delta Length (m)', text='Extra Area (m2)', title="Extra Painted Area per Order")
        st.plotly_chart(fig1, use_container_width=True)

        st.subheader("4. Distribution of Length Variance")
        fig2 = px.histogram(df_summary_display, x='Delta Length (m)', nbins=20, title="Distribution of Delta Length (CCL - CGL)")
        st.plotly_chart(fig2, use_container_width=True)

        st.subheader("5. Thickness vs Length Variance")
        fig3 = px.scatter(df_summary, x='Thickness_Variance', y='Delta_Length', color='Extra_Area_m2', hover_data=[order_col], title="Thickness Variance vs Length Delta")
        st.plotly_chart(fig3, use_container_width=True)

        # =============================
        # 6. EXPORT REPORTS
        # =============================
        st.subheader("6. Export Reports")
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df_summary_display.to_excel(writer, sheet_name='Order Summary', index=False)
            if df_detail_display is not None:
                df_detail_display.to_excel(writer, sheet_name='Details', index=False)
        
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_buffer.getvalue(),
            file_name=f"Paint_Yield_Report_{selected_order if selected_order else 'All'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("💡 To save the dashboard as PDF: Press **Ctrl+P**, set Layout to **Landscape**, and choose **Save as PDF**.")

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")

else:
    st.info("👆 Please upload your master data file (.xlsx or .csv) to begin.")
