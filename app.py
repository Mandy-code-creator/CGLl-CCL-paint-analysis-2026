import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Steel Yield Insight", layout="wide")

# --- GIỮ NGUYÊN THIẾT KẾ LUXURY CỦA BẠN ---
st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
        background-color: white; padding: 25px; border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); margin-bottom: 25px; border: 1px solid #eef2f6;
    }
    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
    table { width: 100% !important; border-radius: 10px !important; border: 1px solid #e2e8f0 !important; }
    th { background-color: #1e3a8a !important; color: white !important; font-weight: 600 !important; text-align: center !important; font-size: 12px; }
    td { text-align: center !important; font-size: 13px !important; border-bottom: 1px solid #f1f5f9 !important; }
    tr:nth-child(even) {background-color: #f8fafc;}
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; background-color: white !important; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Steel Production Yield Analytics")

# =============================
# 1. ĐẦU VÀO: LINK GOOGLE SHEETS
# =============================
st.subheader("🔗 數據來源 (Data Source)")
gsheet_url = st.text_input(
    "Dán Link Google Sheets của bạn vào đây (Phải để chế độ 'Anyone with the link can view'):",
    placeholder="https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"
)

def get_google_sheet_data(url):
    try:
        if "docs.google.com/spreadsheets" in url:
            base_url = url.split('/edit')[0]
            gid = "0"
            if "gid=" in url:
                gid = url.split("gid=")[1]
            csv_url = f"{base_url}/export?format=csv&gid={gid}"
            df = pd.read_csv(csv_url)
            # Làm sạch tên cột
            df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"❌ Lỗi kết nối dữ liệu: {e}")
        return None

if gsheet_url:
    df_raw = get_google_sheet_data(gsheet_url)
    if df_raw is not None:
        st.session_state['saved_data'] = df_raw
        st.success("✅ Đã kết nối dữ liệu trực tuyến thành công!")

# =============================
# 2. PHẦN LOGIC CHI TIẾT (KHÔNG THAY ĐỔI)
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
    cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
    ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

    try:
        # Bước 1: Tính toán theo cuộn mẹ
        step1_agg = {
            cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'first',
            ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
        }
        df_step1 = df.groupby([order_col, mother_col]).agg(step1_agg).reset_index()

        # Bước 2: Tổng hợp theo Đơn hàng
        step2_agg = {
            mother_col: 'count', 
            cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
            ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
        }
        df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()
        df_summary.rename(columns={mother_col: 'Qty', cgl_len: 'Input_m', ccl_len: 'Output_m'}, inplace=True)
        
        # Tính toán chênh lệch
        df_summary['Delta'] = df_summary['Output_m'] - df_summary['Input_m']
        df_summary['Thick_Var'] = df_summary[ccl_thick] - df_summary[cgl_thick]
        df_summary['Area_m2'] = (df_summary[cgl_width] / 1000) * df_summary['Delta']

        # --- BẢNG 1: ORDER SUMMARY ---
        st.subheader("📋 1. 訂單匯總表 (Order Summary)")
        sum_disp = df_summary[[order_col, 'Qty', 'Input_m', 'Output_m', 'Delta', 'Thick_Var', 'Area_m2']].copy()
        sum_disp.columns = ['Order ID', 'Mothers', 'Input (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
        
        sum_disp['Mothers'] = sum_disp['Mothers'].astype(int)
        sum_disp['Input (m)'] = sum_disp['Input (m)'].round(0).astype(int)
        sum_disp['Output (m)'] = sum_disp['Output (m)'].round(0).astype(int)
        sum_disp.insert(0, 'No.', range(1, len(sum_disp) + 1))
        
        st.table(sum_disp.set_index('No.').style.format({
            "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:.2f}"
        }))

        # --- BẢNG 2: BABY COIL DETAILS (PHẦN BẠN CẦN GIỮ) ---
        st.divider()
        st.subheader("🔍 2. 產出鋼捲明細 (Baby Coil Details)")
        selected_order = st.selectbox("Chọn mã đơn hàng để xem chi tiết:", options=df[order_col].unique())

        if selected_order:
            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thick_Diff'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            d_cols = [mother_col, baby_col, cgl_thick, ccl_thick, 'Thick_Diff', ccl_len]
            df_detail_final = df_detail[d_cols].sort_values(by=mother_col).copy()
            df_detail_final.columns = ['Mother Coil', 'Baby Coil', 'CGL Thick', 'CCL Thick', 'Var (mm)', 'CCL Len (m)']
            
            st.table(df_detail_final.style.format({
                "CGL Thick": "{:.3f}", "CCL Thick": "{:.3f}", "Var (mm)": "{:.3f}", "CCL Len (m)": "{:.0f}"
            }))

        # --- PHẦN 3: BIỂU ĐỒ (VISUAL INSIGHTS) ---
        st.divider()
        st.subheader("📈 3. 數據可視化與分析 (Visual Insights)")
        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.bar(sum_disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', color_continuous_scale='RdBu', title="Extra Area per Order")
            st.plotly_chart(fig1, use_container_width=True)
            st.info("**📊 圖表意義:** 此圖顯示各訂單面積流失/延展情況。")
        with c2:
            fig2 = px.histogram(sum_disp, x='Diff (m)', nbins=15, title="Variance Distribution")
            st.plotly_chart(fig2, use_container_width=True)
            st.warning("**📊 圖表意義:** 顯示生產常態與異常點。")

        fig3 = px.scatter(df_summary, x='Thick_Var', y='Delta', color='Area_m2', hover_data=[order_col], title="Thickness vs Length Variance")
        st.plotly_chart(fig3, use_container_width=True)
        st.success("**📊 圖表意義:** 分析厚度變化與長度增減的物理關聯。")

        # --- PHẦN 4: EXECUTIVE SUMMARY (KẾT LUẬN TIẾNG TRUNG) ---
        st.divider()
        st.subheader("💡 4. 執行摘要 (Executive Summary & Yield Insights)")
        total_input, total_output = sum_disp['Input (m)'].sum(), sum_disp['Output (m)'].sum()
        shortage_area = abs(sum_disp[sum_disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())
        
        st.markdown(f"""
        **整體生產指標 (Overall Production Metrics):**
        * **投入與產出 (Input vs Output):** 投入 **{total_input:,.0f} m**，產出 **{total_output:,.0f} m**。
        * 📉 **長度短缺 (Length Shortfall):** 相當於 **{shortage_area:,.2f} m²** 的不明面積差異，需進一步調查。
        """)

        # --- PHẦN 5: XUẤT FILE ---
        st.subheader("💾 5. 數據導出 (Data Export)")
        c_ex, c_pdf = st.columns(2)
        with c_ex:
            excel_bio = io.BytesIO()
            with pd.ExcelWriter(excel_bio, engine='xlsxwriter') as writer:
                sum_disp.to_excel(writer, sheet_name='Summary', index=False)
                df_detail_final.to_excel(writer, sheet_name='Details', index=False)
            st.download_button("📊 DOWNLOAD EXCEL", data=excel_bio.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
        with c_pdf:
            components.html("""
                <script>function printPage() { window.parent.print(); }</script>
                <button onclick="printPage()" style="background-color: white; color: #1e3a8a; border: 2px solid #1e3a8a; 
                border-radius: 8px; padding: 10px; font-size: 15px; cursor: pointer; width: 100%; font-weight: bold; width: 100%;"> 🖨️ SAVE AS PDF REPORT </button>
            """, height=70)

    except Exception as e:
        st.error(f"⚠️ Logic Error: {e}")
