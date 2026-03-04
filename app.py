import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Steel Yield Insight", layout="wide")

# ==========================================================
# 1. AUTO-SYNC CONFIGURATION
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

# --- THIẾT KẾ ĐỒNG NHẤT & KHUNG CUỘN BẢNG ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    h1, h2, h3 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
    
    /* KHUNG CUỘN CHO BẢNG SUMMARY */
    .scrollable-table {
        max-height: 400px;
        overflow-y: auto;
        border: 1px solid #e2e8f0;
        margin-bottom: 25px;
    }

    table { 
        width: 100% !important; 
        border-collapse: collapse !important; 
        font-family: 'Segoe UI', sans-serif;
        color: #334155;
    }
    th { 
        position: sticky; top: 0; 
        border: 1px solid #e2e8f0 !important; 
        color: #1e3a8a !important; 
        text-align: center !important; 
        padding: 12px 8px !important;
        font-size: 13px !important;
        background-color: #f8fafc !important;
        z-index: 10;
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
        .scrollable-table { max-height: none !important; overflow: visible !important; }
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
# 2. CORE LOGIC & FULL DISPLAY
# =============================
if GSHEET_URL:
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        order_c, mother_c, baby_c = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
        cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"

        try:
            # --- XỬ LÝ DỮ LIỆU ---
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

            # --- PHẦN 1: ORDER SUMMARY (KHUNG CUỘN) ---
            st.subheader("1. Order Summary")
            st.markdown("> **Technical Note:** Diff Area (m²) = Coating Area Variance (+ Elongation / - Shortage)")
            
            disp = summary[[order_c, 'Qty', 'In_m', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
            disp.columns = ['Order ID', 'Qty', 'In (m)', 'Out (m)', 'Diff (m)', 'Thick Var', 'Area (m²)']
            disp['Qty'] = disp['Qty'].astype(int)
            disp.insert(0, 'No.', range(1, len(disp) + 1))
            
            st.markdown('<div class="scrollable-table">', unsafe_allow_html=True)
            st.table(disp.set_index('No.').style.format({
                "In (m)": "{:,.0f}", "Out (m)": "{:,.0f}",
                "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Area (m²)": "{:.2f}"
            }))
            st.markdown('</div>', unsafe_allow_html=True)

            # --- PHẦN 2: BABY COIL DETAILS ---
            st.divider()
            st.subheader("2. Baby Coil Details")
            sel_order = st.selectbox("Select Order ID to view details:", options=df[order_c].unique())
            if sel_order:
                det = df[df[order_c] == sel_order].copy()
                det['Var'] = det[ccl_t] - det[cgl_t]
                det_f = det[[mother_c, baby_c, cgl_t, ccl_t, 'Var', ccl_l]].copy()
                det_f.columns = ['Mother Coil', 'Baby Coil', 'CGL Thick', 'CCL Thick', 'Var (mm)', 'CCL Len (m)']
                st.table(det_f.style.format({
                    "CGL Thick": "{:.3f}", "CCL Thick": "{:.3f}", "Var (mm)": "{:.3f}", "CCL Len (m)": "{:.0f}"
                }))

            # --- PHẦN 3: VISUAL INSIGHTS (3 BIỂU ĐỒ) ---
            st.divider()
            st.subheader("3. Visual Insights & Analysis")
            
            # Biểu đồ 1: Bar Chart
            f1 = px.bar(disp, x='Order ID', y='Area (m²)', color='Diff (m)', 
                        color_continuous_scale='RdBu', title="Extra Area per Order")
            st.plotly_chart(f1, use_container_width=True)

            # Biểu đồ 2: Histogram
            f2 = px.histogram(disp, x='Diff (m)', nbins=15, title="Production Variance Distribution")
            st.plotly_chart(f2, use_container_width=True)

            # Biểu đồ 3: Scatter Plot (Dày vs Dài)
            f3 = px.scatter(summary, x='Thick_Var', y='Diff', color='Area_m2', hover_data=[order_c], 
                            title="Thickness Variance vs Length Variance")
            st.plotly_chart(f3, use_container_width=True)

            # --- PHẦN 4: EXECUTIVE SUMMARY (TIẾNG TRUNG) ---
            st.divider()
            st.subheader("4. Executive Summary")
            t_in, t_out = disp['In (m)'].sum(), disp['Out (m)'].sum()
            area_s = abs(disp[disp['Diff (m)'] < 0]['Area (m²)'].sum())
            st.markdown(f"""
            **整體生產指標分析:**
            * **總投入 (Total Input):** {t_in:,.0f} m
            * **總產出 (Total Output):** {t_out:,.0f} m
            * **面積差異 (Area Shortfall):** {area_s:,.2f} m² (此部分需進一步核實廢料申報準確性)
            """)

            # --- PHẦN 5: EXPORT ---
            st.subheader("5. Export Data")
            c1, c2 = st.columns(2)
            with c1:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    disp.to_excel(writer, sheet_name='Summary', index=False)
                st.download_button("📊 Download Excel Report", data=buf.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
            with c2:
                components.html("""
                    <script>function printPage() { window.parent.print(); }</script>
                    <button onclick="printPage()" style="background-color: white; color: #1e3a8a; border: 1.5px solid #1e3a8a; 
                    border-radius: 4px; padding: 10px; font-size: 14px; cursor: pointer; width: 100%; font-weight: 600;"> 
                    Save as PDF Report </button>
                """, height=70)

        except Exception as e:
            st.error(f"Logic Error: {e}")
