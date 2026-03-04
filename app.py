import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

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
# 2. DATA PROCESSING (TRUE LOGIC)
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
    cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
    ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

    try:
        with st.spinner('Calculating yields...'):
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
                mother_col: 'Mother_Count',
                cgl_len: 'CGL_Total', 
                ccl_len: 'CCL_Total'
            }, inplace=True)
            df_summary['Delta_m'] = df_summary['CCL_Total'] - df_summary['CGL_Total']
            df_summary['Thick_Var_mm'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_m']

        # =============================
        # 3. TABLES (ĐÃ SỬA FORMAT SỐ THẬP PHÂN)
        # =============================
        st.subheader("1. Order Summary")
        disp_cols = [order_col, 'Mother_Count', 'CGL_Total', 'CCL_Total', 'Delta_m', 'Thick_Var_mm', 'Extra_Area_m2']
        df_summary_disp = df_summary[disp_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_disp.columns = ['Order', 'Mothers', 'CGL (m)', 'CCL (m)', 'Delta (m)', 'Thick Var (mm)', 'Extra Area (m2)']
        
        # Làm tròn và ép kiểu số nguyên cho số lượng và chiều dài
        df_summary_disp['Mothers'] = df_summary_disp['Mothers'].astype(int)
        df_summary_disp['CGL (m)'] = df_summary_disp['CGL (m)'].round(0).astype(int)
        df_summary_disp['CCL (m)'] = df_summary_disp['CCL (m)'].round(0).astype(int)
        
        # Thêm cột STT
        df_summary_disp.insert(0, 'STT', range(1, len(df_summary_disp) + 1))
        df_summary_disp = df_summary_disp.set_index('STT')

        # ÉP ĐỊNH DẠNG: Chỉ hiện đúng 2 số thập phân (riêng độ dày giữ 3 số vì nó rất nhỏ)
        styled_summary = df_summary_disp.style.format({
            "Delta (m)": "{:.2f}",
            "Thick Var (mm)": "{:.3f}", 
            "Extra Area (m2)": "{:.2f}"
        })
        st.table(styled_summary) 

        st.divider()
        st.subheader("2. Baby Coil Details")
        selected_order = st.selectbox("Select Order Number:", options=df[order_col].unique())

        df_detail_final = None
        if selected_order:
            row = df_summary[df_summary[order_col] == selected_order].iloc[0]
            st.markdown(f"**Performance for Order: {selected_order} ({int(row['Mother_Count'])} Mother Coils)**")
            c1, c2, c3 = st.columns(3)
            c1.metric("CGL Total (Input)", f"{row['CGL_Total']:,.0f} m")
            c2.metric("CCL Total (Output)", f"{row['CCL_Total']:,.0f} m")
            c3.metric("Elongation (Delta)", f"{row['Delta_m']:,.0f} m", delta=f"{row['Delta_m']:,.0f} m", delta_color="inverse")
            
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            d_cols = [mother_col, baby_col, cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len]
            df_detail_final = df_detail[d_cols].sort_values(by=mother_col).copy()
            df_detail_final.columns = ['Mother Coil', 'Baby Coil', 'CGL Thick', 'CCL Thick', 'Var (mm)', 'CCL Len (m)']
            
            # Ép định dạng 3 số thập phân cho bảng chi tiết độ dày
            styled_detail = df_detail_final.style.format({
                "CGL Thick": "{:.3f}",
                "CCL Thick": "{:.3f}",
                "Var (mm)": "{:.3f}",
                "CCL Len (m)": "{:.0f}"
            })
            st.table(styled_detail)

        # =============================
        # 4. VISUAL ANALYSIS
        # =============================
        st.divider()
        st.subheader("3. 視覺化分析與結論 (Visual Analysis & Insights)")
        
        st.markdown("#### 3.1 各訂單額外塗漆面積 (Extra Painted Area per Order)")
        fig1 = px.bar(df_summary_disp, x='Order', y='Extra Area (m2)', color='Delta (m)', text='Extra Area (m2)', title="Extra Painted Area per Order")
        st.plotly_chart(fig1, use_container_width=True)
        st.info("""
        **📊 圖表意義 (Chart Insight):**
        * 此圖表顯示每個訂單因長度差異而產生的「額外塗漆面積」(Extra Area)。
        * **柱狀圖朝下 (Negative Values):** 代表產出長度小於投入長度（短缺），這可能是未記錄的廢料 (Scrap) 或感測器誤差。
        * **柱狀圖朝上 (Positive Values):** 代表鋼帶被拉伸（延展），導致油漆消耗量增加。
        * **結論 (Conclusion):** 管理層應優先調查具有最深藍色/最長柱體的訂單，以找出材料流失的根本原因。
        """)

        st.markdown("#### 3.2 長度差異分佈圖 (Distribution of Length Variance)")
        fig2 = px.histogram(df_summary_disp, x='Delta (m)', nbins=20, title="Distribution of Elongation (CCL - CGL)")
        st.plotly_chart(fig2, use_container_width=True)
        st.warning("""
        **📊 圖表意義 (Chart Insight):**
        * 此分佈圖顯示工廠整體的長度差異趨勢。
        * **集中區域 (The Main Cluster):** 高聳的柱體群代表工廠的「常態短缺/延展」範圍（例如正常的切邊廢料）。
        * **離群值 (Outliers):** 遠離主群體的矮柱（無論是在極左還是極右）代表異常的生產批次。
        * **結論 (Conclusion):** 若發現有訂單落在常態分佈之外（例如短缺超過 -1500m），QC 與生產單位必須立即介入，檢查機台張力設定或廢料記錄流程。
        """)

        st.markdown("#### 3.3 厚度差異 vs. 長度差異 (Thickness vs Length Variance)")
        fig3 = px.scatter(df_summary, x='Thick_Var_mm', y='Delta_m', color='Extra_Area_m2', hover_data=[order_col], title="Thickness Variance vs Length Delta")
        st.plotly_chart(fig3, use_container_width=True)
        st.success("""
        **📊 圖表意義 (Chart Insight):**
        * 此散佈圖旨在分析「厚度變化」是否與「長度流失」有直接物理關聯。
        * **X軸 (Thick_Var_mm):** 塗漆前後的厚度差（約 0.02mm - 0.09mm，通常為正常的漆膜增加厚度）。
        * **Y軸 (Delta_m):** 長度短缺的嚴重程度（向下延伸至 -4500m）。
        * **結論 (Conclusion):** 數據點呈現水平隨機散佈，缺乏明顯的線性趨勢。這強烈暗示那些數千米的嚴重長度短缺（圖表底部的點）**並非源於鋼帶物理變形**。管理層應將調查重點轉向操作面，例如：大量切邊廢料未記錄、CGL/CCL 設備計數器誤差，或資料輸入遺漏。
        """)

        # =============================
        # 5. CONCLUSION
        # =============================
        st.divider()
        st.subheader("💡 4. 執行摘要與產出分析 (Executive Summary & Yield Insights)")
        
        total_input = df_summary_disp['CGL (m)'].sum()
        total_output = df_summary_disp['CCL (m)'].sum()
        
        elongation_df = df_summary_disp[df_summary_disp['Delta (m)'] > 0]
        shortage_df = df_summary_disp[df_summary_disp['Delta (m)'] < 0]
        
        total_elong_area = elongation_df['Extra Area (m2)'].sum() if not elongation_df.empty else 0
        total_shortage_area = abs(shortage_df['Extra Area (m2)'].sum()) if not shortage_df.empty else 0

        st.markdown(f"""
        **整體生產指標 (Overall Production Metrics):**
        * **投入與產出 (Input vs Output):** 投入 **{total_input:,.0f} m** (CGL) 的鋼捲，產出 **{total_output:,.0f} m** (CCL) 的成品。
        * 📈 **正向延展 (Positive Elongation):** 鋼帶延展產生了 **{total_elong_area:,.2f} m²** 的額外表面積。這直接代表了額外的油漆消耗。
        * 📉 **長度短缺 (Length Shortfall):** 產出長度小於投入長度，相當於 **{total_shortage_area:,.2f} m²** 的面積短缺。確切原因（感測器誤差、未記錄的廢料或抽樣裁切）需要進一步調查。
        """)

        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("🔴 **最大正向延展 (Top Positive Elongation):**")
            if not elongation_df.empty:
                worst_elong = elongation_df.sort_values(by='Extra Area (m2)', ascending=False).iloc[0]
                st.info(f"訂單 **`{worst_elong['Order']}`** 延展了 {worst_elong['Delta (m)']:,.0f} m，產生了 **{worst_elong['Extra Area (m2)']:,.2f} m²** 的額外塗漆面積。")
            else:
                st.success("未檢測到明顯的延展現象 (No significant elongation detected).")
                
        with col2:
            st.markdown("🟠 **最大長度短缺 (Top Length Shortfall):**")
            if not shortage_df.empty:
                worst_short = shortage_df.sort_values(by='Extra Area (m2)', ascending=True).iloc[0]
                st.warning(f"訂單 **`{worst_short['Order']}`** 短缺了 {abs(worst_short['Delta (m)']):,.0f} m，代表有 **{abs(worst_short['Extra Area (m2)']):,.2f} m²** 的不明面積差異。")
            else:
                st.success("未檢測到明顯的長度短缺 (No significant length shortage detected).")

        # =============================
        # 6. EXCEL EXPORT 
        # =============================
        excel_data = io.BytesIO()
        with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
            # Lưu ý: file excel tải về cũng sẽ được làm tròn số cho đẹp
            df_summary_export = df_summary_disp.copy()
            df_summary_export['Delta (m)'] = df_summary_export['Delta (m)'].round(2)
            df_summary_export['Extra Area (m2)'] = df_summary_export['Extra Area (m2)'].round(2)
            df_summary_export.to_excel(writer, sheet_name='Summary')
            
            if df_detail_final is not None:
                df_detail_final.to_excel(writer, sheet_name='Details', index=False)
        
        st.divider()
        st.subheader("📥 5. 匯出資料 (Export Report to Excel)")
        st.markdown("Bấm vào nút bên dưới để tải toàn bộ bảng tổng hợp và chi tiết về máy:")
        
        st.download_button(
            label="👉 TẢI XUỐNG FILE EXCEL 👈", 
            data=excel_data.getvalue(), 
            file_name="Paint_Yield_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        # =============================
        # 7. PDF EXPORT (NÚT IN BÁO CÁO)
        # =============================
        st.divider()
        st.subheader("🖨️ 6. 匯出 PDF (Export to PDF)")
        st.markdown("Bấm nút bên dưới để mở giao diện in báo cáo. Vui lòng chọn **Save as PDF (Lưu dưới dạng PDF)** trong cửa sổ hiện ra.")
        
        components.html(
            """
            <script>
            function printPage() {
                window.parent.print();
            }
            </script>
            <button onclick="printPage()" style="
                padding: 10px 20px; 
                font-size: 16px; 
                font-weight: bold; 
                background-color: #ff4b4b; 
                color: white; 
                border: none; 
                border-radius: 5px; 
                cursor: pointer;
                width: 100%;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            ">
                🖨️ XUẤT FILE PDF / IN BÁO CÁO
            </button>
            """,
            height=60
        )

    except Exception as e:
        st.error(f"Logic Error: {e}")
else:
    st.info("👆 Please upload your data file to start.")
