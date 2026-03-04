import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Steel Yield Insight", layout="wide")

# ==========================================================
# 1. AUTO-SYNC CONFIGURATION (INSERT YOUR LINK HERE)
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

# --- MINIMALIST DESIGN: UNIFORM GRID LINES ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
        background-color: #ffffff; padding: 20px; border-radius: 0px;
        margin-bottom: 20px; border: none;
    }
    h1, h2, h3 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
    table { 
        width: 100% !important; 
        border-collapse: collapse !important; 
        font-family: 'Segoe UI', sans-serif;
        color: #334155;
        border: 1px solid #e2e8f0 !important;
    }
    th { 
        border: 1px solid #e2e8f0 !important; 
        color: #1e3a8a !important; 
        text-align: center !important; 
        padding: 12px 8px !important;
        font-size: 13px !important;
        background-color: #f8fafc !important;
    }
    td { 
        text-align: center !important; 
        padding: 10px 8px !important; 
        border: 1px solid #e2e8f0 !important; 
        font-size: 13px !important;
    }
    tr:hover { background-color: #f1f5f9; }
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; }
        table { border: 1px solid #000 !important; }
        th, td { border: 0.5pt solid #ccc !important; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Steel Production Yield Analytics")

# --- DATA FETCHING ---
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
            df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# =============================
# 2. CORE LOGIC
# =============================
if GSHEET_URL and GSHEET_URL != "CHÈN_LINK_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        order_c, mother_c, baby_c = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
        cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"

        try:
            # Aggregate Logic
            s1 = df.groupby([order_c, mother_c]).agg({
                cgl_t: 'mean', cgl_w: 'mean', cgl_l: 'first',
                ccl_t: 'mean', ccl_w: 'mean', ccl_l: 'sum'
            }).reset_index()

            summary = s1.groupby(order_c).agg({
                mother_c: 'count', cgl_l: 'sum', ccl_l: 'sum', 
                ccl_t: 'mean', cgl_t: 'mean', cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty', cgl_l: 'In_m', ccl_l: 'Out_m'})
            summary['Diff'] = summary['Out_m'] - summary['In_m']
            summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- 1. ORDER SUMMARY ---
            st.subheader("1. Order Summary")
            
            # Ghi chú thuật ngữ (Technical Term Note)
            st.markdown("""
            > **💡 術語說明 (Technical Note):** > **Diff Area (m²)** = **塗層面積差異** (Coating Area Variance)  
            > * **正值 (+):** 鋼帶延展 (Elongation)，導致塗漆消耗量增加。  
            > * **負值 (-):** 長度短缺 (Shortage)，可能源於剪切廢料 (Scrap) 或感測器誤差。
            """)

            disp = summary[[order_c, 'Qty', 'In_m', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
            disp.columns = ['Order ID', 'Input Coil Number', 'Input (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
            disp['Mothers'] = disp['Mothers'].astype(int)
            disp.insert(0, 'No.', range(1, len(disp) + 1))
            
            st.table(disp.set_index('No.').style.format({
                "Input (m)": "{:,.0f}", "Output (m)": "{:,.0f}",
                "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:.2f}"
            }))

           # --- 2. PRODUCTION COIL DETAILS ---
            st.divider()
            st.subheader("2. Production Coil Details") # Đổi tên subheader cho chuyên nghiệp
            
            order_list = df[order_c].unique()
            sel_order = st.selectbox("Select Order ID to view details:", options=order_list)
            
            if sel_order:
                det = df[df[order_c] == sel_order].copy()
                det['Var'] = det[ccl_t] - det[cgl_t]
                
                # Chọn các cột hiển thị
                det_f = det[[
                    mother_c, baby_c, 
                    cgl_t, cgl_w, cgl_l, 
                    ccl_t, ccl_w, 'Var', ccl_l 
                ]].copy()
                
                # Đổi tên cột: Bỏ "Mother/Baby", dùng "Input/Output" hoặc "CGL/CCL"
                det_f.columns = [
                    'Input Coil ID (CGL)', 
                    'Output Coil ID (CCL)', 
                    'Input Thick (mm)', 
                    'Input Width (mm)', 
                    'Input Length (m)', 
                    'Output Thick (mm)', 
                    'Output Width (mm)', 
                    'Thick Deviation (mm)', 
                    'Output Length (m)'
                ]
                
                st.table(det_f.style.format({
                    "Input Thick (mm)": "{:.3f}", 
                    "Input Width (mm)": "{:,.0f}", 
                    "Input Length (m)": "{:,.0f}",
                    "Output Thick (mm)": "{:.3f}", 
                    "Output Width (mm)": "{:,.0f}",
                    "Thick Deviation (mm)": "{:.3f}", 
                    "Output Length (m)": "{:,.0f}"
                }))
            # --- 3. VISUAL INSIGHTS (CONCLUSIONS IN CHINESE) ---
            st.divider()
            st.subheader("3. Visual Insights & Analysis")
            
            f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                        color_continuous_scale='RdBu', title="Extra Area per Order")
            st.plotly_chart(f1, use_container_width=True)
            st.info("**分析結論:** 監控各訂單的塗層面積偏差。偏離中心值的數據代表生產投入與產出不一致，建議優先核對該批次的生產日誌。")

            f2 = px.histogram(disp, x='Diff (m)', nbins=15, title="Production Variance Distribution")
            st.plotly_chart(f2, use_container_width=True)
            st.warning("**分析結論:** 數據分布反映生產穩定性。離群值標示該訂單存在異常長度變化，需確認是物理延展、裁切損耗或是計量誤差。")

            # --- 4. EXECUTIVE SUMMARY ---
            st.divider()
            st.subheader("4. Executive Summary")
            t_in, t_out = disp['Input (m)'].sum(), disp['Output (m)'].sum()
            area_s = abs(disp[disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())
            st.markdown(f"""
            **生產產出綜合分析:**
            * **總投入 (Total Input):** {t_in:,.0f} m
            * **總產出 (Total Output):** {t_out:,.0f} m
            * **不明面積差異 (Area Shortfall):** {area_s:,.2f} m² (需進一步核實廢料申報準確性)
            """)

            # --- 5. EXPORT ---
            st.subheader("5. Export Data")
            c1, c2 = st.columns(2)
            with c1:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    disp.to_excel(writer, sheet_name='Summary', index=False)
                st.download_button("📊 Download Excel", data=buf.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
            with c2:
                components.html("""
                    <script>function printPage() { window.parent.print(); }</script>
                    <button onclick="printPage()" style="background-color: white; color: #1e3a8a; border: 1.5px solid #1e3a8a; 
                    border-radius: 4px; padding: 10px; font-size: 14px; cursor: pointer; width: 100%; font-weight: 600;"> 
                    Save as PDF Report </button>
                """, height=70)

        except Exception as e:
            st.error(f"Logic Error: {e}")
else:
    st.info("Please insert the Google Sheet Link in the source code (GSHEET_URL).")
