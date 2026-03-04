import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Steel Yield Insight", layout="wide")

# ==========================================================
# 1. CẤU HÌNH LINK TỰ ĐỘNG (THAY LINK CỦA BẠN TẠI ĐÂY)
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

# --- THIẾT KẾ LUXURY & LẰN KẺ NGANG SIÊU MẢNH ---
st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; }
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
        background-color: white; padding: 25px; border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); margin-bottom: 25px; border: 1px solid #eef2f6;
    }
    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
    table { width: 100% !important; border-collapse: collapse !important; font-family: 'Segoe UI', sans-serif; }
    th { border-bottom: 2px solid #1e3a8a !important; color: #1e3a8a !important; text-align: center !important; padding: 12px 8px !important; font-size: 13px !important; }
    td { text-align: center !important; padding: 10px 8px !important; border-bottom: 0.5px solid #e2e8f0 !important; font-size: 13px !important; }
    tr:hover { background-color: #f1f5f9; }
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; background-color: white !important; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Steel Production Yield Analytics (Auto-Sync)")

@st.cache_data(ttl=300)
def load_auto_data(url):
    try:
        if "docs.google.com/spreadsheets" in url:
            base_url = url.split('/edit')[0]
            gid = "0"
            if "gid=" in url:
                gid = url.split("gid=")[1].split("&")[0]
            csv_url = f"{base_url}/export?format=csv&gid={gid}"
            df = pd.read_csv(csv_url)
            # Làm sạch tên cột tự động
            df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"⚠️ Lỗi kết nối dữ liệu: Hãy đảm bảo Link ở chế độ Public. Lỗi: {e}")
        return None

# =============================
# 2. LOGIC XỬ LÝ (FIX LỖI MISMATCH)
# =============================
if GSHEET_URL and GSHEET_URL != "CHÈN_LINK_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
        cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

        try:
            # Step 1: Aggregate by Mother
            df_step1 = df.groupby([order_col, mother_col]).agg({
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }).reset_index()

            # Step 2: Summary by Order
            df_summary = df_step1.groupby(order_col).agg({
                mother_col: 'count', cgl_len: 'sum', ccl_len: 'sum', 
                cgl_width: 'mean', cgl_thick: 'mean', ccl_thick: 'mean'
            }).reset_index()

            # Rename an toàn (không bao giờ lỗi Length Mismatch)
            df_summary = df_summary.rename(columns={
                mother_col: 'Qty', 
                cgl_len: 'Input_m', 
                ccl_len: 'Output_m'
            })

            # Tính toán các chỉ số
            df_summary['Delta'] = df_summary['Output_m'] - df_summary['Input_m']
            df_summary['Thick_Var'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Area_m2'] = (df_summary[cgl_width] / 1000) * df_summary['Delta']

            # --- 📋 1. ORDER SUMMARY ---
            st.subheader("📋 1. 訂單匯總表 (Order Summary)")
            sum_disp = df_summary[[order_col, 'Qty', 'Input_m', 'Output_m', 'Delta', 'Thick_Var', 'Area_m2']].copy()
            
            # Đổi tên hiển thị cho người dùng
            disp_names = {
                order_col: 'Order ID', 'Qty': 'Mothers', 'Input_m': 'Input (m)', 
                'Output_m': 'Output (m)', 'Delta': 'Diff (m)', 
                'Thick_Var': 'Thick Var', 'Area_m2': 'Diff Area (m²)'
            }
            sum_disp = sum_disp.rename(columns=disp_names)
            
            sum_disp.insert(0, 'No.', range(1, len(sum_disp) + 1))
            st.table(sum_disp.set_index('No.').style.format({
                "Input (m)": "{:,.0f}", "Output (m)": "{:,.0f}",
                "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:.2f}"
            }))

            # --- 🔍 2. BABY COIL DETAILS ---
            st.divider()
            st.subheader("🔍 2. 產出鋼捲明細 (Baby Coil Details)")
            selected_order = st.selectbox("Select Order Number:", options=df[order_col].unique())
            if selected_order:
                df_det = df[df[order_col] == selected_order].copy()
                df_det['Var'] = df_det[ccl_thick] - df_det[cgl_thick]
                df_det_final = df_det[[mother_col, baby_col, cgl_thick, ccl_thick, 'Var', ccl_len]].copy()
                df_det_final.columns = ['Mother Coil', 'Baby Coil', 'CGL Thick', 'CCL Thick', 'Var (mm)', 'CCL Len (m)']
                st.table(df_det_final.style.format({
                    "CGL Thick": "{:.3f}", "CCL Thick": "{:.3f}", "Var (mm)": "{:.3f}", "CCL Len (m)": "{:.0f}"
                }))

            # --- 📈 3. VISUAL INSIGHTS ---
            st.divider()
            st.subheader("📈 3. 數據可視化與分析 (Visual Insights)")
            fig1 = px.bar(sum_disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                          color_continuous_scale='RdBu', title="Extra Area per Order")
            fig1.update_layout(plot_bgcolor='white', margin=dict(l=10,r=10,t=40,b=10))
            st.plotly_chart(fig1, use_container_width=True)
            st.info("**💡 結論:** 深色柱體代表長 độ dài chênh lệch lớn, cần kiểm tra lại quy trình.")

            fig2 = px.histogram(sum_disp, x='Diff (m)', nbins=15, title="Variance Distribution")
            fig2.update_layout(plot_bgcolor='white')
            st.plotly_chart(fig2, use_container_width=True)

            fig3 = px.scatter(df_summary, x='Thick_Var', y='Delta', color='Area_m2', hover_data=[order_col], title="Thickness vs Length Variance")
            st.plotly_chart(fig3, use_container_width=True)

            # --- 💡 4. EXECUTIVE SUMMARY ---
            st.divider()
            st.subheader("💡 4. 執行摘要 (Executive Summary)")
            total_in, total_out = sum_disp['Input (m)'].sum(), sum_disp['Output (m)'].sum()
            short_area = abs(sum_disp[sum_disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())
            st.markdown(f"""
            * **投入與產出:** 投入 **{total_in:,.0f} m**，產出 **{total_out:,.0f} m**。
            * 📉 **長度短缺:** Tương đương **{short_area:,.2f} m²** diện tích thiếu hụt không rõ nguyên nhân.
            """)

            # --- 💾 5. DATA EXPORT ---
            st.subheader("💾 5. 數據導出 (Data Export)")
            c1, c2 = st.columns(2)
            with c1:
                excel_bio = io.BytesIO()
                with pd.ExcelWriter(excel_bio, engine='xlsxwriter') as writer:
                    sum_disp.to_excel(writer, sheet_name='Summary', index=False)
                st.download_button("📊 DOWNLOAD EXCEL", data=excel_bio.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
            with c2:
                components.html("""
                    <script>function printPage() { window.parent.print(); }</script>
                    <button onclick="printPage()" style="background-color: white; color: #1e3a8a; border: 2px solid #1e3a8a; 
                    border-radius: 8px; padding: 10px; font-size: 15px; cursor: pointer; width: 100%; font-weight: bold; width: 100%;"> 🖨️ SAVE AS PDF REPORT </button>
                """, height=70)

        except Exception as e:
            st.error(f"⚠️ Logic Error: {e}")
