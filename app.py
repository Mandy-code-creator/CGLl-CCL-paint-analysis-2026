import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

# --- CSS FOR PRINTING (TỐI ƯU CHIỀU RỘNG BẢNG) ---
st.markdown("""
    <style>
    @media print {
        header, .stActionButton, .stSidebar, [data-testid="stHeader"], [data-testid="stFileUploadDropzone"] { 
            display: none !important; 
        }
        
        html, body, .stApp, .main {
            height: auto !important;
            overflow: visible !important;
        }
        
        .main .block-container { 
            max-width: 100% !important; 
            padding: 0.5cm !important; 
            margin: 0 !important; 
        }

        /* Thu hẹp khoảng cách cột trong bảng khi in */
        table { 
            width: auto !important; /* Không ép bảng rộng 100% */
            margin-left: 0 !important;
            border-collapse: collapse !important; 
        }
        th, td { 
            font-size: 10px !important; 
            padding: 4px 8px !important; /* Giảm padding ngang để các cột gần nhau hơn */
            border: 1px solid #ccc !important; 
            text-align: center !important;
        }

        .stPlotlyChart { width: 100% !important; }
        .stPlotlyChart, .stDataFrame, .stTable, .stMarkdown { 
            page-break-inside: avoid !important; 
        }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Paint Yield Analysis: CGL vs CCL Elongation")

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
# 2. DATA PROCESSING
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
    cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
    ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

    try:
        with st.spinner('Calculating...'):
            step1_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_col]).agg(step1_agg).reset_index()

            step2_agg = {
                mother_col: 'count', 
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            df_summary.rename(columns={
                mother_col: 'Mothers',
                cgl_len: 'CGL_Total', 
                ccl_len: 'CCL_Total'
            }, inplace=True)
            df_summary['Delta_m'] = df_summary['CCL_Total'] - df_summary['CGL_Total']
            df_summary['Thick_Var_mm'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_m']

        # =============================
        # =============================
        # 3. TABLES (ÉP CHIỀU RỘNG CỘT CỐ ĐỊNH)
        # =============================
        st.subheader("1. Order Summary")
        disp_cols = [order_col, 'Mothers', 'CGL_Total', 'CCL_Total', 'Delta_m', 'Thick_Var_mm', 'Extra_Area_m2']
        df_summary_disp = df_summary[disp_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_disp.columns = ['Order', 'Mothers', 'CGL (m)', 'CCL (m)', 'Delta (m)', 'Thick Var (mm)', 'Extra Area (m2)']
        
        df_summary_disp['Mothers'] = df_summary_disp['Mothers'].astype(int)
        df_summary_disp['CGL (m)'] = df_summary_disp['CGL (m)'].round(0).astype(int)
        df_summary_disp['CCL (m)'] = df_summary_disp['CCL (m)'].round(0).astype(int)
        
        df_summary_disp.insert(0, 'STT', range(1, len(df_summary_disp) + 1))
        df_summary_disp = df_summary_disp.set_index('STT')

        # Dùng HTML để kiểm soát tuyệt đối chiều rộng (Width Control)
        styled_summary = df_summary_disp.style.format({
            "Delta (m)": "{:.2f}",
            "Thick Var (mm)": "{:.3f}", 
            "Extra Area (m2)": "{:.2f}"
        }).set_table_styles([
            {'selector': 'th', 'props': [('width', '80px'), ('text-align', 'center'), ('background-color', '#f0f2f6')]},
            {'selector': 'td', 'props': [('width', '80px'), ('text-align', 'center')]},
            # Ép riêng cột Order rộng hơn một chút để không bị nhảy dòng
            {'selector': '.col0', 'props': [('width', '150px')]}, 
        ])

        # Sử dụng st.write kết hợp với CSS ẩn để bảng trông gọn nhất
        st.write(styled_summary, unsafe_allow_html=True)

        st.divider()
        st.subheader("2. Baby Coil Details")
        # (Giữ nguyên phần selectbox và bảng chi tiết bên dưới)

        # =============================
        # 4. VISUAL ANALYSIS (Giữ nguyên các biểu đồ của bạn)
        # =============================
        st.divider()
        st.subheader("3. 視覺化分析與結論 (Visual Analysis & Insights)")
        
        # Biểu đồ 1
        fig1 = px.bar(df_summary_disp.reset_index(), x='Order', y='Extra Area (m2)', color='Delta (m)', title="Extra Painted Area per Order")
        st.plotly_chart(fig1, use_container_width=True)
        
        # Biểu đồ 2
        fig2 = px.histogram(df_summary_disp, x='Delta (m)', nbins=20, title="Distribution of Elongation")
        st.plotly_chart(fig2, use_container_width=True)

        # Biểu đồ 3
        fig3 = px.scatter(df_summary, x='Thick_Var_mm', y='Delta_m', color='Extra_Area_m2', hover_data=[order_col], title="Thickness Variance vs Length Delta")
        st.plotly_chart(fig3, use_container_width=True)

        # =============================
        # 5. CONCLUSION (Giữ nguyên nội dung tiếng Trung)
        # =============================
        st.divider()
        st.subheader("💡 4. 執行摘要與產出分析 (Executive Summary & Yield Insights)")
        
        total_input = df_summary_disp['CGL (m)'].sum()
        total_output = df_summary_disp['CCL (m)'].sum()
        elongation_df = df_summary_disp[df_summary_disp['Delta (m)'] > 0]
        shortage_df = df_summary_disp[df_summary_disp['Delta (m)'] < 0]
        
        st.markdown(f"""
        **整體生產指標 (Overall Production Metrics):**
        * **投入與產出 (Input vs Output):** 投入 **{total_input:,.0f} m**，產出 **{total_output:,.0f} m**。
        * 📉 **長度短缺 (Length Shortfall):** 產出長 độ ngắn hơn đầu vào tương đương **{abs(shortage_df['Extra Area (m2)'].sum()):,.2f} m²** diện tích thiếu hụt.
        """)

        # =============================
        # 6. EXPORT BUTTONS
        # =============================
        st.divider()
        st.subheader("📥 5. 匯出資料 (Export Reports)")
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            excel_data = io.BytesIO()
            with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                df_summary_disp.to_excel(writer, sheet_name='Summary')
            st.download_button("📊 Tải xuống Excel", data=excel_data.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)

        with col_btn2:
            components.html("""
                <script>function printPage() { window.parent.print(); }</script>
                <button onclick="printPage()" style="background-color: white; color: #0066cc; border: 1px solid #d3d3d3; border-radius: 6px; padding: 8px 16px; font-size: 15px; cursor: pointer; width: 100%; height: 40px; font-weight: 500;">
                    🖨️ Save as PDF Report
                </button>
            """, height=50)

    except Exception as e:
        st.error(f"Logic Error: {e}")
else:
    st.info("👆 Please upload your data file to start.")
