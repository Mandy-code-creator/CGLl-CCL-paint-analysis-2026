import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

# --- CSS FOR UI OPTIMIZATION & PRINTING ---
st.markdown("""
    <style>
    /* 1. Giao diện Web: Thu gọn khoảng cách */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 1200px !important; /* Giới hạn độ rộng để không bị loãng trên màn hình lớn */
    }
    
    /* 2. Ép bảng Order Summary gọn lại */
    .stTable {
        width: fit-content !important;
        margin: 0 auto !important;
    }
    
    table {
        border: 1px solid #e6e9ef !important;
        border-radius: 5px !important;
    }

    th {
        background-color: #f8f9fb !important;
        text-align: center !important;
        font-size: 13px !important;
        padding: 8px !important;
    }

    td {
        text-align: center !important;
        font-size: 13px !important;
        padding: 5px !important;
    }

    /* 3. CSS dành riêng cho IN ẤN (PDF) */
    @media print {
        header, .stActionButton, .stSidebar, [data-testid="stHeader"], [data-testid="stFileUploadDropzone"], .stDivider { 
            display: none !important; 
        }
        html, body, .stApp, .main {
            height: auto !important;
            overflow: visible !important;
        }
        .main .block-container { 
            max-width: 100% !important; 
            padding: 0.5cm !important; 
        }
        table { 
            width: 100% !important; 
            table-layout: fixed !important; 
        }
        th, td { 
            font-size: 10px !important; 
            border: 1px solid #ccc !important; 
        }
        .stPlotlyChart { width: 100% !important; }
        .stInfo, .stWarning, .stSuccess { 
            padding: 10px !important;
            font-size: 11px !important;
            page-break-inside: avoid !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🎨 Paint Yield Analysis: CGL vs CCL")

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

            df_summary.rename(columns={mother_col: 'Mothers', cgl_len: 'CGL_Total', ccl_len: 'CCL_Total'}, inplace=True)
            df_summary['Delta_m'] = df_summary['CCL_Total'] - df_summary['CGL_Total']
            df_summary['Thick_Var_mm'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_m']

        # =============================
        # 3. TABLES (TRÌNH BÀY GỌN)
        # =============================
        st.subheader("1. Order Summary")
        disp_cols = [order_col, 'Mothers', 'CGL_Total', 'CCL_Total', 'Delta_m', 'Thick_Var_mm', 'Extra_Area_m2']
        df_summary_disp = df_summary[disp_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_disp.columns = ['Order', 'Mothers', 'CGL(m)', 'CCL(m)', 'Δ(m)', 'Var(mm)', 'Area(m2)']
        
        df_summary_disp['Mothers'] = df_summary_disp['Mothers'].astype(int)
        df_summary_disp['CGL(m)'] = df_summary_disp['CGL(m)'].round(0).astype(int)
        df_summary_disp['CCL(m)'] = df_summary_disp['CCL(m)'].round(0).astype(int)
        df_summary_disp.insert(0, 'STT', range(1, len(df_summary_disp) + 1))
        
        styled_summary = df_summary_disp.set_index('STT').style.format({
            "Δ(m)": "{:.2f}", "Var(mm)": "{:.3f}", "Area(m2)": "{:.2f}"
        })
        st.table(styled_summary)

        st.divider()
        st.subheader("2. Baby Coil Details")
        selected_order = st.selectbox("🔍 Select Order to view details:", options=df[order_col].unique())
        
        df_detail_final = None
        if selected_order:
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            d_cols = [mother_col, baby_col, cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len]
            df_detail_final = df_detail[d_cols].sort_values(by=mother_col).copy()
            df_detail_final.columns = ['Mother', 'Baby', 'CGL_T', 'CCL_T', 'Var', 'CCL_L']
            st.table(df_detail_final.style.format({"CGL_T": "{:.3f}", "CCL_T": "{:.3f}", "Var": "{:.3f}"}))

        # =============================
        # 4. VISUAL ANALYSIS
        # =============================
        st.divider()
        st.subheader("3. 視覺化分析與結論 (Insights)")
        
        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.bar(df_summary_disp, x='Order', y='Area(m2)', color='Δ(m)', title="Extra Area per Order")
            fig1.update_layout(margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig1, use_container_width=True)
            st.info("**📊 圖表意義:** 此圖顯示各訂單產生之額外面積。負值代表產出短缺，正值代表延展。")

        with c2:
            fig2 = px.histogram(df_summary_disp, x='Δ(m)', nbins=15, title="Length Variance Distribution")
            fig2.update_layout(margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig2, use_container_width=True)
            st.warning("**📊 圖表意義:** 顯示整體趨勢。主區塊代表常態損耗，離群值代表異常生產。")

        fig3 = px.scatter(df_summary, x='Thick_Var_mm', y='Delta_m', color='Extra_Area_m2', hover_data=[order_col], title="Thickness vs Length Variance")
        st.plotly_chart(fig3, use_container_width=True)
        st.success("**📊 圖表意義:** 分析厚度與長度的物理關聯。若無趨勢，則異常源於操作或記錄誤差。")

        # =============================
        # 5. CONCLUSION
        # =============================
        st.divider()
        st.subheader("💡 4. 執行摘要 (Executive Summary)")
        total_input, total_output = df_summary_disp['CGL(m)'].sum(), df_summary_disp['CCL(m)'].sum()
        short_area = abs(df_summary_disp[df_summary_disp['Δ(m)'] < 0]['Area(m2)'].sum())
        
        st.markdown(f"""
        **整體生產指標 (Overall Production Metrics):**
        * **投入與產出:** 投入 **{total_input:,.0f} m**，產出 **{total_output:,.0f} m**。
        * 📉 **長度短缺 (Shortfall):** 相當於 **{short_area:,.2f} m²** 的不明面積差異，需進一步調查原因。
        """)

        # =============================
        # 6. EXPORT BUTTONS
        # =============================
        st.divider()
        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            excel_data = io.BytesIO()
            with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                df_summary_disp.to_excel(writer, sheet_name='Summary')
            st.download_button("📊 Tải xuống Excel", data=excel_data.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
        with c_btn2:
            components.html("""
                <script>function printPage() { window.parent.print(); }</script>
                <button onclick="printPage()" style="background-color: white; color: #0066cc; border: 1px solid #d3d3d3; border-radius: 6px; padding: 10px; font-size: 15px; cursor: pointer; width: 100%; font-weight: 500;">
                    🖨️ Save as PDF Report
                </button>
            """, height=60)

    except Exception as e:
        st.error(f"Logic Error: {e}")
else:
    st.info("👆 Please upload your data file to start.")
