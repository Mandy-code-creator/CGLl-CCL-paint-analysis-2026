import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between Galvanizing (CGL) mother coils 
and Color Coating (CCL) baby coils to estimate hidden paint loss.
""")

# ==========================================
# FILE UPLOAD & DEVELOPMENT MODE
# ==========================================
# Trick: If you are tired of re-uploading during coding, comment out the st.file_uploader line 
# and uncomment the dev_file_path line with your actual file name.

uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

# dev_file_path = "your_data_file.xlsx" # <-- Uncomment this line and put your file name here
# uploaded_file = dev_file_path         # <-- Uncomment this line

if uploaded_file is not None:
    try:
        if isinstance(uploaded_file, str):
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.endswith('.xlsx') else pd.read_csv(uploaded_file)
        else:
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
            
        st.session_state['saved_data'] = df_temp 
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    # --- DATA COLUMNS (KEEPING ORIGINAL CHINESE HEADERS FOR PARSING) ---
    order_col = "訂單號碼"
    mother_coil_col = "投入鋼捲號碼"
    
    # PLEASE UPDATE THIS TO MATCH YOUR BABY COIL COLUMN NAME
    baby_coil_col = "子鋼捲號碼" 
    
    cgl_thick = "镀锌實測厚度"
    cgl_width = "镀锌測寬度"
    cgl_len = "镀锌測長度"
    
    ccl_thick = "實測厚度"
    ccl_width = "實測寬度"
    ccl_len = "實測長度"

    try:
        with st.spinner('Processing and aggregating data...'):
            # Step 1: Group by Mother Coil
            step1_agg = {
                cgl_thick: 'first', cgl_width: 'first', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_coil_col]).agg(step1_agg).reset_index()

            # Step 2: Group by Order
            step2_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            # Rename and calculate variances
            df_summary.rename(columns={cgl_len: 'CGL_Total_Length', ccl_len: 'CCL_Total_Length'}, inplace=True)
            df_summary['Delta_Length'] = df_summary['CCL_Total_Length'] - df_summary['CGL_Total_Length']
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_Length']

        # ==========================================
        # SECTION 1: ORDER SUMMARY (SIMPLIFIED)
        # ==========================================
        st.subheader("1. Order Summary")
        
        summary_display_cols = [
            order_col, 
            'CGL_Total_Length', 
            'CCL_Total_Length', 
            'Delta_Length', 
            'Thickness_Variance',
            'Extra_Area_m2'
        ]
        
        # Translate headers to English for the UI display
        df_summary_display = df_summary[summary_display_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_display.columns = [
            'Order Number', 'CGL Total Length (m)', 'CCL Total Length (m)', 
            'Delta Length (m)', 'Thickness Variance (mm)', 'Extra Area (m2)'
        ]
        
        st.dataframe(df_summary_display, use_container_width=True)

        st.divider()

        # ==========================================
        # SECTION 2: BABY COIL DETAILS (SIMPLIFIED)
        # ==========================================
        st.subheader("2. Baby Coil Details")
        st.markdown("Select an order to view the thickness variance analysis for individual baby coils.")
        
        order_list = df[order_col].dropna().unique().tolist()
        selected_order = st.selectbox("Select Order Number:", options=order_list)
        
        if selected_order:
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            
            try:
                # Select only analysis-relevant columns
                detail_display_cols = [
                    mother_coil_col, baby_coil_col, 
                    cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len
                ]
                df_detail_display = df_detail[detail_display_cols].sort_values(by=mother_coil_col).copy()
                
                # Rename columns to clean English for the UI
                df_detail_display.columns = [
                    'Mother Coil', 'Baby Coil', 
                    'CGL Thickness', 'CCL Thickness', 'Thickness Variance', 'CCL Length (m)'
                ]
                
                st.dataframe(df_detail_display, use_container_width=True)
                
            except KeyError:
                st.warning("Could not find the baby coil column. Please update the `baby_coil_col` variable in the code with the exact column name from your Excel file.")
                st.dataframe(df_detail)

    except KeyError as e:
        st.error(f"Missing column in your file: {e}")
        st.info("Please ensure the uploaded file contains the correct Chinese column names.")
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")

else:
    st.info("👆 Please upload your master data file (.xlsx or .csv) to begin.")
