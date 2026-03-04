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
    
    /* Style cho các Card nội dung */
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
        background-color: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        margin-bottom: 25px;
        border: 1px solid #eef2f6;
    }

    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }

    /* Tối ưu hóa bảng hiển thị */
    table { width: 100% !important; border-radius: 10px !important; border: 1px solid #e2e8f0 !important; }
    th { background-color: #1e3a8a !important; color: white !important; font-weight: 600 !important; text-align: center !important; font-size: 12px; }
    td { text-align: center !important; font-size: 13px !important; border-bottom: 1px solid #f1f5f9 !important; }
    tr:nth-child(even) {background-color: #f8fafc;}

    /* Ẩn các thành phần khi in */
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; background-color: white !important; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Steel Production Yield Analytics")

# =============================
# 1. THAY ĐỔI ĐẦU VÀO THÀNH LINK (GIỮ NGUYÊN LOGIC)
# =============================
st.subheader("🔗 數據來源 (Data Source)")
data_link = st.text_input("Dán Link Google Sheets của bạn vào đây:", 
                         placeholder="https://docs.google.com/spreadsheets/d/...")

def load_data(link):
    try:
        # Tự động xử lý link Google Sheets để đọc dữ liệu
        if "docs.google.com/spreadsheets" in link:
            link = link.replace("/edit#gid=", "/export?format=csv&gid=")
            if "/edit" in link and "gid=" not in link:
                link = link.replace("/edit", "/export?format=csv")
        
        df = pd.read_csv(link) if ("csv" in link.lower() or "docs.google.com" in link) else pd.read_excel(link)
        df.columns = df.columns.str.replace(r'\s+', '', regex=True)
        return df.loc[:, ~df.columns.duplicated()]
    except Exception as e:
        st.error(f"❌ Không thể nạp dữ liệu. Hãy đảm bảo bạn đã 'Chia sẻ cho bất kỳ ai có liên kết'.")
        return None

if data_link:
    df_raw = load_data(data_link)
    if df_raw is not None:
        st.session_state['saved_data'] = df_raw
        st.success("✅ Dữ liệu đã được kết nối!")

# =============================
# 2. GIỮ NGUYÊN TOÀN BỘ THIẾT KẾ PHẦN DƯỚI
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    order_col, mother_col = "訂單號碼", "投入鋼捲號碼"
    cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
    ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

    try:
        # Core Logic
        step1 = df.groupby([order_col, mother_col]).agg({
            cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'first',
            ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
        }).reset_index()

        df_summary = step1.groupby(order_col).agg({
            mother_col: 'count', cgl_len: 'sum', ccl_len: 'sum', cgl_width: 'mean'
        }).reset_index()

        df_summary.rename(columns={mother_col: 'Qty', cgl_len: 'Input_m', ccl_len: 'Output_m'}, inplace=True)
        df_summary['Delta'] = df_summary['Output_m'] - df_summary['Input_m']
        df_summary['Area_m2'] = (df_summary[cgl_width] / 1000) * df_summary['Delta']

        # --- PHẦN 1: BẢNG TỔNG HỢP ---
        st.subheader("📋 1. 訂單匯總表 (Order Summary)")
        sum_disp = df_summary[[order_col, 'Qty', 'Input_m', 'Output_m', 'Delta', 'Area_m2']].copy()
        sum_disp.columns = ['Order ID', 'Mothers', 'Input (m)', 'Output (m)', 'Diff (m)', 'Diff Area (m²)']
        
        sum_disp['Mothers'] = sum_disp['Mothers'].astype(int)
        sum_disp['Input (m)'] = sum_disp['Input (m)'].round(0).astype(int)
        sum_disp['Output (m)'] = sum_disp['Output (m)'].round(0).astype(int)
        sum_disp.insert(0, 'No.', range(1, len(sum_disp) + 1))
        
        st.table(sum_disp.set_index('No.').style.format({
            "Diff (m)": "{:.2f}", "Diff Area (m²)": "{:.2f}"
        }))

        # --- PHẦN 2: BIỂU ĐỒ & KẾT LUẬN ---
        st.subheader("📈 2. 數據可視化與分析 (Visual Insights)")
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            fig1 = px.bar(sum_disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                          color_continuous_scale='RdBu', title="Extra Surface Area per Order")
            fig1.update_layout(plot_bgcolor='white', margin=dict(l=10,r=10,t=40,b=10))
            st.plotly_chart(fig1, use_container_width=True)
            st.info("**💡 結論:** 負值代表面積流失。深色柱體為重點調查對象。")

        with col_chart2:
            fig2 = px.histogram(sum_disp, x='Diff (m)', nbins=15, title="Production Variance Distribution")
            fig2.update_traces(marker_color='#1e3a8a', marker_line_color='white', marker_line_width=1)
            fig2.update_layout(plot_bgcolor='white', margin=dict(l=10,r=10,t=40,b=10))
            st.plotly_chart(fig2, use_container_width=True)
            st.warning("**💡 結論:** 集中區代表常態損耗，孤立點代表生產異常。")

        # --- PHẦN 3: XUẤT BÁO CÁO ---
        st.subheader("💾 3. 數據導出 (Data Export)")
        c_ex, c_pdf = st.columns(2)
        
        with c_ex:
            excel_bio = io.BytesIO()
            with pd.ExcelWriter(excel_bio, engine='xlsxwriter') as writer:
                sum_disp.to_excel(writer, sheet_name='Summary', index=False)
            st.download_button("📊 DOWNLOAD EXCEL", data=excel_bio.getvalue(), 
                               file_name="Production_Report.xlsx", type="primary", use_container_width=True)
            
        with c_pdf:
            components.html("""
                <script>function printPage() { window.parent.print(); }</script>
                <button onclick="printPage()" style="background-color: white; color: #1e3a8a; border: 2px solid #1e3a8a; 
                border-radius: 8px; padding: 10px; font-size: 15px; cursor: pointer; width: 100%; font-weight: bold;
                transition: 0.3s; width: 100%;"> 🖨️ SAVE AS PDF REPORT </button>
            """, height=70)

    except Exception as e:
        st.error(f"⚠️ Logic Error: {e}")
