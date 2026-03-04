import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Steel Yield Insight", layout="wide")

st.title("Production Yield Analysis & Variance Control System: Paint code 4890")

# ==========================================================
# 1. GOOGLE SHEET LINK
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0"

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
            df.columns = df.columns.str.normalize('NFKC')  # Normalize Unicode
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# =============================
# 2. LOAD DATA
# =============================
df = load_auto_data(GSHEET_URL)
if df is None:
    st.info("Cannot load data. Check your Google Sheet URL.")
    st.stop()

# --- COLUMN MAPPING ---
order_c, mother_c, baby_c = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"

# =============================
# 3. SUMMARY CALCULATION
# =============================
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
summary['Outlier'] = summary['Diff'].apply(lambda x: 'Yes' if abs(x) > 10 else 'No')  # Threshold

# =============================
# 4. ORDER SUMMARY TABLE
# =============================
st.subheader("1. Order Summary")
disp = summary[[order_c, 'Qty', 'In_m', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2', 'Outlier']].copy()
disp.columns = ['Order ID', 'Mothers', 'Input (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)', 'Outlier']
disp.insert(0, 'No.', range(1, len(disp)+1))

st.table(disp.style.applymap(
    lambda x: 'color: red; font-weight: bold;' if x=='Yes' else '', subset=['Outlier']
).format({
    "Input (m)":"{:,.0f}", "Output (m)":"{:,.0f}", "Diff (m)":"{:.2f}",
    "Thick Var":"{:.3f}", "Diff Area (m²)":"{:.2f}"
}))

# =============================
# 5. BABY COIL DETAILS
# =============================
st.divider()
st.subheader("2. Baby Coil Details")
sel_order = st.selectbox("Select Order ID:", options=df[order_c].unique())
if sel_order:
    det = df[df[order_c]==sel_order].copy()
    det['Var'] = det[ccl_t] - det[cgl_t]
    det_f = det[[mother_c, baby_c, cgl_t, ccl_t, 'Var', ccl_l]].copy()
    det_f.columns = ['Mother Coil','Baby Coil','CGL Thick','CCL Thick','Var (mm)','CCL Len (m)']
    st.table(det_f.style.format({
        "CGL Thick":"{:.3f}","CCL Thick":"{:.3f}","Var (mm)":"{:.3f}","CCL Len (m)":"{:.0f}"
    }))

# =============================
# 6. VISUAL INSIGHTS
# =============================
st.divider()
st.subheader("3. Visual Insights & Analysis")

f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)',
            color_continuous_scale='RdBu', title="Extra Area per Order")
st.plotly_chart(f1, use_container_width=True)

f2 = px.histogram(disp, x='Diff (m)', nbins=15, title="Production Variance Distribution")
st.plotly_chart(f2, use_container_width=True)

f3 = px.scatter(disp, x='Order ID', y='Diff (m)', color='Outlier', size='Diff Area (m²)',
                title="Outlier Detection")
st.plotly_chart(f3, use_container_width=True)

# =============================
# 7. EXECUTIVE SUMMARY
# =============================
st.divider()
st.subheader("4. Executive Summary")
t_in, t_out = disp['Input (m)'].sum(), disp['Output (m)'].sum()
area_s = abs(disp[disp['Diff (m)']<0]['Diff Area (m²)'].sum())

st.markdown(f"""
**生產產出綜合分析:**
* **總投入 (Total Input):** {t_in:,.0f} m
* **總產出 (Total Output):** {t_out:,.0f} m
* **不明面積差異 (Area Shortfall):** {area_s:,.2f} m²
""")

# =============================
# 8. EXPORT OPTIONS
# =============================
st.divider()
st.subheader("5. Export Data")
c1, c2 = st.columns(2)
with c1:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        disp.to_excel(writer, sheet_name='Summary', index=False)
    st.download_button("📊 Download Excel", data=buf.getvalue(), file_name="Report.xlsx")

with c2:
    components.html("""
        <script>function printPage(){window.parent.print();}</script>
        <button onclick="printPage()" style="background-color:white;color:#1e3a8a;
        border:1.5px solid #1e3a8a;border-radius:4px;padding:10px;font-size:14px;
        cursor:pointer;width:100%;font-weight:600;">
        Save as PDF Report</button>
    """, height=70)
