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

# --- MINIMALIST ENGLISH UI & THIN HORIZONTAL LINES ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    
    /* Content Card */
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
        background-color: #ffffff; padding: 20px; border-radius: 0px;
        margin-bottom: 20px; border: none;
    }

    h1, h2, h3 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }

    /* MINIMALIST TABLE DESIGN: ONLY THIN GREY HORIZONTAL LINES */
    table { 
        width: 100% !important; 
        border-collapse: collapse !important; 
        font-family: 'Segoe UI', sans-serif;
        color: #334155;
    }
    th { 
        border-bottom: 1.5px solid #1e3a8a !important; 
        color: #1e3a8a !important; 
        text-align: center !important; 
        padding: 12px 8px !important;
        font-size: 13px !important;
        background-color: transparent !important;
    }
    td { 
        text-align: center !important; 
        padding: 10px 8px !important; 
        border-bottom: 0.5px solid #e2e8f0 !important; /* Minimalist thin line */
        font-size: 13px !important;
    }
    tr:hover { background-color: #f8fafc; }

    /* Hide elements during PDF print */
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; }
        div[data-testid="stVerticalBlock"] > div { border: none !important; }
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
        st.error(f"Connection Error: Please check link permissions. Error: {e}")
        return None

# =============================
# 2. CORE LOGIC (STRICTLY ENGLISH)
# =============================
if GSHEET_URL and GSHEET_URL != "CHÈN_LINK_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        order_col, mother_col, baby_col = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
        cgl_thick, cgl_width, cgl_len = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_thick, ccl_width, ccl_len = "實測厚度", "實測寬度", "實測長度"

        try:
            # Processing Summary
            step1 = df.groupby([order_col, mother_col]).agg({
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }).reset_index()

            df_summary = step1.groupby(order_col).agg({
                mother_col: 'count', cgl_len: 'sum', ccl_len: 'sum', 
                cgl_width: 'mean', cgl_thick: 'mean', ccl_thick: 'mean'
            }).reset_index()

            df_summary = df_summary.rename(columns={
                mother_col: 'Qty', cgl_len: 'In_m', ccl_len: 'Out_m'
            })

            df_summary['Diff'] = df_summary['Out_m'] - df_summary['In_m']
            df_summary['Thick_Var'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Area_m2'] = (df_summary[cgl_width] / 1000) * df_summary['Diff']

            # --- 1. ORDER SUMMARY ---
            st.subheader("1. Order Summary")
            sum_disp = df_summary[[order_col, 'Qty', 'In_m', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
            
            sum_disp.columns = ['Order ID', 'Mothers', 'Input (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
            sum_disp['Mothers'] = sum_disp['Mothers'].astype(int)
            sum_disp['Input (m)'] = sum_disp['Input (m)'].round(0).astype(int)
            sum_disp['Output (m)'] = sum_disp['Output (m)'].round(0).astype(int)
            sum_disp.insert(0, 'No.', range(1, len(sum_disp) + 1))
            
            st.table(sum_disp.set_index('No.').style.format({
                "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:.2f}"
            }))

            # --- 2. BABY COIL DETAILS ---
            st.divider()
            st.subheader("2. Baby Coil Details")
            selected_order = st.selectbox("Select Order ID:", options=df[order_col].unique())
            if selected_order:
                df_det = df[df[order_col] == selected_order].copy()
                df_det['Var'] = df_det[ccl_thick] - df_det[cgl_thick]
                df_det_final = df_det[[mother_col, baby_col, cgl_thick, ccl_thick, 'Var', ccl_len]].copy()
                df_det_final.columns = ['Mother', 'Baby', 'CGL Thick', 'CCL Thick', 'Var (mm)', 'CCL Len (m)']
                st.table(df_det_final.style.format({
                    "CGL Thick": "{:.3f}", "CCL Thick": "{:.3f}", "Var (mm)": "{:.3f}", "CCL Len (m)": "{:.0f}"
                }))

            # --- 3. VISUAL INSIGHTS ---
            st.divider()
            st.subheader("3. Visual Insights")
            fig1 = px.bar(sum_disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', color_continuous_scale='RdBu', title="Extra Area per Order")
            fig1.update_layout(plot_bgcolor='white')
            st.plotly_chart(fig1, use_container_width=True)

            fig2 = px.histogram(sum_disp, x='Diff (m)', nbins=15, title="Variance Distribution")
            fig2.update_layout(plot_bgcolor='white')
            st.plotly_chart(fig2, use_container_width=True)

            # --- 4. EXECUTIVE SUMMARY ---
            st.divider()
            st.subheader("4. Executive Summary")
            total_in, total_out = sum_disp['Input (m)'].sum(), sum_disp['Output (m)'].sum()
            short_area = abs(sum_disp[sum_disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())
            st.markdown(f"""
            * **Total Input:** {total_in:,.0f} m
            * **Total Output:** {total_out:,.0f} m
            * **Area Shortfall:** {short_area:,.2f} m²
            """)

            # --- 5. DATA EXPORT ---
            st.subheader("5. Export Data")
            c1, c2 = st.columns(2)
            with c1:
                excel_bio = io.BytesIO()
                with pd.ExcelWriter(excel_bio, engine='xlsxwriter') as writer:
                    sum_disp.to_excel(writer, sheet_name='Summary', index=False)
                st.download_button("📊 Download Excel", data=excel_bio.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
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
    st.info("Please insert Google Sheet Link in the source code (GSHEET_URL).")
