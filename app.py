import streamlit as st
import pandas as pd
import plotly.express as px

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between **Galvanizing (CGL)** mother coils 
and **Color Coating (CCL)** baby coils to estimate hidden paint loss due to steel elongation.
""")

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

# Development Mode: Uncomment below if you want to load a local file automatically
# dev_file_path = "data.xlsx" 
# uploaded_file = dev_file_path         

if uploaded_file is not None:
    try:
        # Load Data
        if isinstance(uploaded_file, str):
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.endswith('.xlsx') else pd.read_csv(uploaded_file)
        else:
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
            
        # 1. CLEANING: Remove whitespaces from column names
        df_temp.columns = df_temp.columns.str.replace(r'\s+', '', regex=True)
        
        # 2. CLEANING: Remove duplicate columns from ERP export
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]
        
        # Store in session state to persist data during UI interactions
        st.session_state['saved_data'] = df_temp 
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# --- MAIN ANALYSIS LOGIC ---
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    # Define Column Names (Matching your ERP headers)
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
            # Step 1: Aggregate by Mother Coil to avoid duplicating mother length
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

            # Variance Calculations
            df_summary.rename(columns={cgl_len: 'CGL_Total_Length', ccl_len: 'CCL_Total_Length'}, inplace=True)
            df_summary['Delta_Length'] = df_summary['CCL_Total_Length'] - df_summary['CGL_Total_Length']
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_Length']

        # --- SECTION 1: ORDER SUMMARY ---
        st.subheader("1. Order Summary")
        
        summary_display_cols = [
            order_col, 'CGL_Total_Length', 'CCL_Total_Length', 
            'Delta_Length', 'Thickness_Variance', 'Extra_Area_m2'
        ]
        
        df_summary_display = df_summary[summary_display_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        
        # Safe Rename using Dictionary
        df_summary_display.rename(columns={
            order_col: 'Order Number',
            'CGL_Total_Length': 'CGL Total Length (m)',
            'CCL_Total_Length': 'CCL Total Length (m)',
            'Delta_Length': 'Delta Length (m)',
            'Thickness_Variance': 'Thickness Change (mm)',
            'Extra_Area_m2': 'Extra Paint Area (m2)'
        }, inplace=True)
        
        st.dataframe(df_summary_display, use_container_width=True)
        st.divider()

        # --- SECTION 2: BABY COIL DETAILS ---
        st.subheader("2. Baby Coil Details")
        st.markdown("Select an order to view the breakdown and total length analysis.")
        
        order_list = df[order_col].dropna().unique().tolist()
        selected_order = st.selectbox("Select Order Number:", options=order_list)
        
        if selected_order:
            # Metric Cards for selected order
            order_totals = df_summary[df_summary[order_col] == selected_order].iloc[0]
            
            st.markdown(f"**Performance Metrics for: {selected_order}**")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Mother Coils (CGL)", f"{order_totals['CGL_Total_Length']:,.0f} m")
            m2.metric("Total Baby Coils (CCL)", f"{order_totals['CCL_Total_Length']:,.0f} m")
            m3.metric("Length Variance", f"{order_totals['Delta_Length']:,.0f} m", 
                      delta=f"{order_totals['Delta_Length']:,.0f} m", delta_color="inverse")
            
            st.write("") 
            
            # Detail Table Filter
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            
            try:
                detail_display_cols = [
                    mother_coil_col, baby_coil_col, 
                    cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len
                ]
                df_detail_display = df_detail[detail_display_cols].sort_values(by=mother_coil_col).copy()
                
                # Professional Rename for Details
                df_detail_display.rename(columns={
                    mother_coil_col: 'Mother Coil ID',
                    baby_coil_col: 'Baby Coil ID',
                    cgl_thick: 'CGL Thick (mm)',
                    ccl_thick: 'CCL Thick (mm)',
                    'Thickness_Variance': 'Diff (mm)',
                    ccl_len: 'Baby Length (m)'
                }, inplace=True)
                
                st.dataframe(df_detail_display, use_container_width=True)
                
            except KeyError as e:
                st.warning(f"Note: Column {e} not found for selective display.")

    except Exception as e:
        st.error(f"An error occurred during calculation: {e}")

else:
    st.info("👆 Please upload your master data file (.xlsx or .csv) to start the analysis.")
