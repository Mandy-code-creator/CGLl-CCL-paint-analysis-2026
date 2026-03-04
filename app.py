import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

# --- CSS FOR PRINTING (Fixing the "Hidden Columns" issue) ---
st.markdown("""
    <style>
    @media print {
        /* Hide UI elements that are not needed in report */
        .stActionButton, .stSidebar, .stHeader, [data-testid="stFileUploadDropzone"] { display: none !important; }
        
        /* Force tables to show all columns and fit page width */
        .main .block-container { max-width: 100% !important; padding: 0 !important; }
        table { width: 100% !important; font-size: 10px !important; table-layout: fixed !important; }
        th, td { word-wrap: break-word !important; }
        
        /* Remove scrolling for printing */
        div[data-testid="stTable"] { overflow: visible !important; }
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
    except Exception as e:
        st.error(f"Error: {e}")

# =============================
# 2. PROCESSING & DISPLAY
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
    cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
    ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

    try:
        # Aggregation Logic
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
        df_summary['Extra_Area'] = (df_summary[ccl_width] / 1000) * df_summary['Delta']

        # --- HIỂN THỊ BẢNG (Dùng st.table để in không bị mất cột) ---
        st.subheader("1. Summary Table")
        disp_summary = df_summary[[order_col, 'CGL_Len', 'CCL_Len', 'Delta', 'Extra_Area']].copy()
        disp_summary.columns = ['Order', 'CGL(m)', 'CCL(m)', 'Delta(m)', 'Area(m2)']
        
        # Sử dụng st.table thay cho st.dataframe để hiển thị toàn bộ cột khi in
        st.table(disp_summary.sort_values(by='Area', ascending=False).head(20))

        st.subheader("2. Detailed Breakdown")
        selected_order = st.selectbox("Select Order:", options=df[order_col].unique())
        
        if selected_order:
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Var'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            df_detail_final = df_detail[[mother_col, baby_col, cgl_thick, ccl_thick, 'Var', ccl_len]].copy()
            df_detail_final.columns = ['Mother', 'Baby', 'CGL-T', 'CCL-T', 'Diff', 'Len']
            
            # Hiển thị bảng chi tiết dạng tĩnh để in ấn trọn vẹn
            st.table(df_detail_final)

        # =============================
        # 3. EXPORT
        # =============================
        excel_data = io.BytesIO()
        with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
            disp_summary.to_excel(writer, sheet_name='Summary', index=False)
        st.sidebar.download_button("📥 Download Excel (Full Data)", data=excel_data.getvalue(), file_name="Report.xlsx")

    except Exception as e:
        st.error(f"Processing Error: {e}")
