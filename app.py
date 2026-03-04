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

# --- MINIMALIST DESIGN & UNIFORM GRID LINES ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    
    /* Headers */
    h1, h2, h3 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }

    /* UNIFORM GRID DESIGN: ALL CELLS HAVE THIN GREY LINES */
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
        padding: 10px;
        font-size: 13px !important;
        background-color: #f8fafc !important;
    }
    td { 
        text-align: center !important; 
        padding: 8px; 
        border: 1px solid #e2e8f0 !important; 
        font-size: 13px !important;
    }
    tr:hover { background-color: #f1f5f9; }

    /* Custom Expander Header */
    .streamlit-expanderHeader { 
        background-color: #f8fafc !important; 
        border: 1px solid #e2e8f0 !important;
        font-weight: 600 !important;
    }

    /* Print Optimization */
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Steel Production Yield Analytics")

# --- DATA FETCHING FUNCTION ---
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
        st.error(f"Connection Error: Please check Google Sheet link permissions. Error: {e}")
        return None

# =============================
# 2. DATA PROCESSING & UI
# =============================
if GSHEET_URL and GSHEET_URL != "CHÈN_LINK_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        # Internal Column Mapping
        order_c, mother_c, baby_c = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
        cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"

        try:
            # Step 1: Aggregate by Mother Coil
            s1 = df.groupby([order_c, mother_c]).agg({
                cgl_t: 'mean', cgl_w: 'mean', cgl_l: 'first',
                ccl_t: 'mean', ccl_w: 'mean', ccl_l: 'sum'
            }).reset_index()

            # Step 2: Summary by Order ID
            summary = s1.groupby(order_c).agg({
                mother_c: 'count', cgl_l: 'sum', ccl_l: 'sum', 
                ccl_t: 'mean', cgl_t: 'mean', cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty', cgl_l: 'In_m', ccl_l: 'Out_m'})
            summary['Diff'] = summary['Out_m'] - summary['In_m']
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- SECTION 1: COLLAPSIBLE ORDER SUMMARY ---
            with st.expander("📋 1. Order Summary Table (Click to Expand/Collapse)", expanded=True):
                st.markdown("""
                > **💡 術語說明 (Technical Note):**
                > * **Diff Area (m²)** = **塗層面積差異** (Coating Area Variance)
                > * **正值 (+):** 鋼帶延展 | **負值 (-):** 長度短缺
                """)
                
                disp = summary[[order_c, 'Qty', 'In_m', 'Out_m', 'Diff', 'Area_m2']].copy()
                disp.columns = ['Order ID', 'Qty', 'In (m)', 'Out (m)', 'Diff (m)', 'Diff Area (m²)']
                disp['Qty'] = disp['Qty'].astype(int)
                disp.insert(0, 'No.', range(1, len(disp) + 1))
                st.table(disp.set_index('No.').style.format({"Diff (m)": "{:.2f}", "Diff Area (m²)": "{:.2f}"}))

            # --- SECTION 2: BABY COIL DETAILS ---
            st.subheader("🔍 2. Baby Coil Details")
            sel_order = st.selectbox("Select Order ID to view details:", options=df[order_c].unique())
            if sel_order:
                det = df[df[order_c] == sel_order].copy()
                det['Var'] = det[ccl_t] - det[cgl_t]
                det_f = det[[mother_c, baby_c, cgl_t, ccl_t, 'Var', ccl_l]].copy()
                det_f.columns = ['Mother Coil', 'Baby Coil', 'CGL Thick', 'CCL Thick', 'Var (mm)', 'CCL Len (m)']
                st.table(det_f.style.format({"CGL Thick": "{:.3f}", "CCL Thick": "{:.3f}", "Var (mm)": "{:.3f}", "CCL Len (m)": "{:.0f}"}))

            # --- SECTION 3: VISUAL INSIGHTS ---
            st.divider()
            st.subheader("📈 3. Visual Insights & Analysis")
            
            f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', color_continuous_scale='RdBu', title="Area Variance per Order")
            f1.update_layout(plot_bgcolor='white')
            st.plotly_chart(f1, use_container_width=True)
            st.info("**分析結論:** 監控各訂單的塗層面積偏差。異常數據代表生產投入與產出不一致，建議核對該批次的生產日誌。")

            f2 = px.histogram(disp, x='Diff (m)', nbins=15, title="Variance Distribution")
            f2.update_layout(plot_bgcolor='white')
            st.plotly_chart(f2, use_container_width=True)
            st.warning("**分析結論:** 數據分布反映生產穩定性。離群值標示該訂單有顯著長度變化，需核實生產記錄或設備計量。")

            # --- SECTION 4: EXECUTIVE SUMMARY ---
            st.divider()
            st.subheader("💡 4. Executive Summary")
            t_in, t_out = disp['In (m)'].sum(), disp['Out (m)'].sum()
            st.markdown(f"""
            **生產產出綜合分析 (Production Summary):**
            * **總投入 (Total Input):** {t_in:,.0f} m
            * **總產出 (Total Output):** {t_out:,.0f} m
            * **淨長度差異 (Net Variance):** {t_out - t_in:,.2f} m
            """)

            # --- SECTION 5: DATA EXPORT ---
            st.subheader("💾 5. Export Data")
            c1, c2 = st.columns(2)
            with c1:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    disp.to_excel(writer, sheet_name='Summary', index=False)
                st.download_button("📊 Download Excel Report", data=buf.getvalue(), file_name="Steel_Report.xlsx", type="primary", use_container_width=True)
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
    st.info("👋 Chào bạn! Hãy chèn Link Google Sheet của bạn vào biến GSHEET_URL ở dòng 13 để bắt đầu.")
