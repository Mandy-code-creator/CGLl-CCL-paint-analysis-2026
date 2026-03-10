import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis: Total CGL vs CCL per Order", layout="wide")

# --- STYLING ---
st.markdown("""
<style>
.stApp { background-color: #ffffff; }
div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
    background-color: #ffffff; padding: 20px; border-radius: 0px;
    margin-bottom: 20px; border: none;
}
h1, h2, h3 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
table { width: 100% !important; border-collapse: collapse !important; font-family: 'Segoe UI', sans-serif; color: #334155; border: 1px solid #e2e8f0 !important;}
th { border: 1px solid #e2e8f0 !important; color: #1e3a8a !important; text-align: center !important; padding: 12px 8px !important; font-size: 13px !important; background-color: #f8fafc !important; }
td { text-align: center !important; padding: 10px 8px !important; border: 1px solid #e2e8f0 !important; font-size: 13px !important;}
tr:hover { background-color: #f1f5f9; }
@media print {
    header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
    .main .block-container { max-width: 100% !important; padding: 0.5cm !important; }
    table { border: 1px solid #000 !important; }
    th, td { border: 0.5pt solid #ccc !important; }
}
</style>
""", unsafe_allow_html=True)

st.title("Length Variance Analysis: Total CGL vs CCL per Order")

# --- DATA FETCHING ---
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

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
            df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# =============================
# CORE LOGIC
# =============================
if GSHEET_URL and GSHEET_URL != "CHÈN_LINK_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        # Column mapping
        order_c, mother_c, baby_c = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
        cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"
        outer_cut, inner_cut = "outercutlength", "innercutlength"

        for col in [outer_cut, inner_cut]:
            if col not in df.columns:
                df[col] = 0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        for col in [cgl_t, cgl_w, cgl_l, ccl_t, ccl_w, ccl_l]:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # --- BASE COIL ---
        df['base_coil'] = df[mother_c].astype(str).str.replace(r'[A-Za-z]\d{2}$', '', regex=True)

        # --- CHECK MAIN COIL ---
        is_main_coil = df[mother_c].astype(str).str.endswith('X00')
        main_coil_present = df[is_main_coil].groupby([order_c, 'base_coil']).size().reset_index(name='has_main')
        df = df.merge(main_coil_present, on=[order_c, 'base_coil'], how='left')
        df['has_main'] = df['has_main'].fillna(0) > 0

        # --- HANDLE ORPHAN & P0 coils ---
        # Step 1: create a map from base_coil to main CGL_L
        main_length_map = df[is_main_coil].set_index(['訂單號碼','base_coil'])[cgl_l].to_dict()

        # Step 2: assign Input Length
        def assign_input(row):
            key = (row[order_c], row['base_coil'])
            if row[mother_c].endswith('X00'):  # main coil
                return row[cgl_l]
            elif row['has_main']:  # child coil with main in group (P0)
                return main_length_map.get(key, row[ccl_l])  # take mother gốc length
            else:  # orphan coil
                return row[ccl_l]  # take own output length

        df['CGL_L_fixed'] = df.apply(assign_input, axis=1)

        # Step 3: fill Thickness & Width
        df[cgl_t] = df.groupby([order_c, 'base_coil'])[cgl_t].transform(lambda x: x.ffill().bfill()).fillna(0)
        df[cgl_w] = df.groupby([order_c, 'base_coil'])[cgl_w].transform(lambda x: x.ffill().bfill()).fillna(0)

        # fill empty CCL
        df[ccl_t] = df[ccl_t].fillna(0)
        df[ccl_w] = df[ccl_w].fillna(0)
        df[ccl_l] = df[ccl_l].fillna(0)

        # --- GROUP BY COIL ---
        s1 = df.groupby([order_c, mother_c]).agg({
            cgl_t: 'mean', cgl_w: 'mean', 'CGL_L_fixed': 'first',
            ccl_t: 'mean', ccl_w: 'mean', ccl_l: 'sum',
            outer_cut: 'max', inner_cut: 'max' 
        }).reset_index()

        # --- ORDER SUMMARY ---
        summary = s1.groupby(order_c).agg({
            mother_c: 'count', 'CGL_L_fixed': 'sum', ccl_l: 'sum',
            outer_cut: 'sum', inner_cut: 'sum',
            ccl_t: 'mean', cgl_t: 'mean', cgl_w: 'mean'
        }).reset_index()
        summary = summary.rename(columns={mother_c: 'Qty (Coils)', 'CGL_L_fixed': 'In_m', ccl_l: 'Out_m'})

        summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
        summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
        summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]
        summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

        # --- DISPLAY ORDER SUMMARY ---
        st.subheader("1. Order Summary")
        disp = summary[[order_c, 'Qty (Coils)', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
        disp.columns = ['Order ID', 'Qty (Coils)', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
        disp = disp.sort_values(by='Cut Scrap (m)', ascending=False).reset_index(drop=True)
        disp.insert(0, 'No.', range(1, len(disp)+1))
        st.table(disp.set_index('No.').style.format({
            "Input (m)": "{:,.0f}", "Cut Scrap (m)": "{:,.0f}", "Output (m)": "{:,.0f}",
            "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:.2f}"
        }))

        # --- PRODUCTION COIL DETAILS ---
        st.divider()
        st.subheader("2. Production Coil Details") 
        order_list = df[order_c].unique()
        sel_order = st.selectbox("🔍 Select Order ID:", options=order_list)
        if sel_order:
            det = df[df[order_c]==sel_order].copy()
            det['Var'] = det[ccl_t] - det[cgl_t]
            det_f = det[[mother_c, baby_c, cgl_t, cgl_w, 'CGL_L_fixed', outer_cut, inner_cut, ccl_t, ccl_w, 'Var', ccl_l]]
            det_f.columns = ['Input Coil ID (CGL)','Output Coil ID (CCL)','Input Thick (mm)','Input Width (mm)','Input Length (m)','Outer Cut (m)','Inner Cut (m)','Output Thick (mm)','Output Width (mm)','Thick Deviation (mm)','Output Length (m)']
            st.table(det_f.style.format({
                "Input Thick (mm)":"{:.3f}", "Input Width (mm)":"{:,.0f}", "Input Length (m)":"{:,.0f}",
                "Outer Cut (m)":"{:,.0f}", "Inner Cut (m)":"{:,.0f}",
                "Output Thick (mm)":"{:.3f}", "Output Width (mm)":"{:,.0f}", "Thick Deviation (mm)":"{:.3f}", "Output Length (m)":"{:,.0f}"
            }))

        # --- VISUALS ---
        st.divider()
        st.subheader("3. Visual Insights & Analysis")
        f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', color_continuous_scale='RdBu', title="Extra Area per Order")
        st.plotly_chart(f1, use_container_width=True)

        f2 = px.histogram(disp, x='Diff (m)', nbins=15, title="Production Variance Distribution")
        st.plotly_chart(f2, use_container_width=True)

        f3 = px.bar(disp.sort_values(by='Cut Scrap (m)', ascending=False), x='Order ID', y='Cut Scrap (m)', color='Cut Scrap (m)', color_continuous_scale='Reds', title="Total Cut Scrap per Order")
        f3.update_layout(yaxis_title="Scrap Length (m)")
        st.plotly_chart(f3, use_container_width=True)

        # --- EXECUTIVE SUMMARY ---
        st.divider()
        st.subheader("4. Executive Summary")
        t_in, t_out = disp['Input (m)'].sum(), disp['Output (m)'].sum()
        area_s = abs(disp[disp['Diff (m)']<0]['Diff Area (m²)'].sum())
        st.markdown(f"**Total Input:** {t_in:,.0f} m, **Total Output:** {t_out:,.0f} m, **Area Shortfall:** {area_s:,.2f} m²")

        # --- EXPORT ---
        st.subheader("5. Export Data")
        c1, c2 = st.columns(2)
        with c1:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                disp.to_excel(writer, sheet_name='Summary', index=False)
            st.download_button("📊 Download Excel", data=buf.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
        with c2:
            components.html("""<script>function printPage(){window.parent.print();}</script>
            <button onclick="printPage()" style="background-color:white;color:#1e3a8a;border:1.5px solid #1e3a8a;border-radius:4px;padding:10px;width:100%;font-weight:600;">Save as PDF</button>""", height=70)
else:
    st.info("Please insert the Google Sheet Link in the source code (GSHEET_URL).")
