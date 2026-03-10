import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis: Total CGL vs CCL per Order", layout="wide")

# ==========================================================
# 1. THEME SELECTION WITH ANIMATED GRADIENT
# ==========================================================
theme_choice = st.radio("🎨 Select App Theme:", ["Light Mode (Standard)", "Dark Mode (Professional)"], horizontal=True)

if theme_choice == "Dark Mode (Professional)":
    # Dải màu tối: Deep Blue -> Purple -> Teal
    grad_colors = "#0f172a, #1e3a8a, #5b21b6, #0d9488"
    card_bg = "rgba(30, 41, 59, 0.7)"
    text_color = "#f8fafc"
    sub_text = "#cbd5e1"
    plotly_template = "plotly_dark"
else:
    # Dải màu sáng: Soft Blue -> Pink -> Lavender
    grad_colors = "#f0f9ff, #e0f2fe, #fbcfe8, #ede9fe"
    card_bg = "rgba(255, 255, 255, 0.8)"
    text_color = "#1e3a8a"
    sub_text = "#334155"
    plotly_template = "plotly_white"

# CSS hiệu ứng chuyển động Gradient
st.markdown(f"""
    <style>
    .stApp {{
        background: linear-gradient(-45deg, {grad_colors});
        background-size: 400% 400%;
        animation: gradient_move 15s ease infinite;
    }}

    @keyframes gradient_move {{
        0% {{ background-position: 0% 50%; }}
        50% {{ background-position: 100% 50%; }}
        100% {{ background-position: 0% 50%; }}
    }}

    /* Hiệu ứng Glassmorphism cho các khung dữ liệu */
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stDataFrame) {{
        background-color: {card_bg}; 
        backdrop-filter: blur(10px);
        padding: 20px; border-radius: 12px;
        margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.1);
    }}

    h1, h2, h3 {{ color: {text_color}; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }}
    .stMarkdown p {{ color: {sub_text} !important; }}
    .stSelectbox label, .stRadio label {{ color: {text_color} !important; font-weight: bold; }}
    </style>
    """, unsafe_allow_html=True)

st.title("Length Variance Analysis: Total CGL vs CCL per Order")

# ==========================================================
# 2. DATA PROCESSING (YOUR ORIGINAL LOGIC)
# ==========================================================
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

if GSHEET_URL:
    df = load_auto_data(GSHEET_URL)
    if df is not None:
        def get_col(default, possible_names):
            for name in possible_names:
                if name in df.columns: return name
            return default

        order_c = get_col("訂單號碼", ["訂單號碼", "订单号码"])
        mother_c = get_col("投入鋼捲號碼", ["投入鋼捲號碼", "投入钢卷号码"])
        baby_c = get_col("產出鋼捲號碼", ["產出鋼捲號碼", "产出钢卷号码"])
        cgl_l = get_col("镀锌測長度", ["镀锌測長度", "镀锌實測長度", "镀锌长度", "鍍鋅測長度"])
        ccl_l = get_col("實測長度", ["實測長度", "实测长度"])
        cgl_w = get_col("镀锌測寬度", ["镀锌測寬度", "镀锌測寬", "镀锌宽度", "鍍鋅測寬度", "镀锌实测宽度"])
        cgl_t = get_col("镀锌實測厚度", ["镀锌實測厚度", "镀锌測厚", "镀锌厚度", "鍍鋅實測厚度"])
        ccl_w = get_col("實測寬度", ["實測寬度", "实测宽度"])
        ccl_t = get_col("實測厚度", ["實測厚度", "实测厚度"])
        outer_cut = get_col("outercutlength", ["outercutlength", "outercut"])
        inner_cut = get_col("innercutlength", ["innercutlength", "innercut"])
        line_c = get_col("線別", ["線別", "线别"])
        out_grade_c = get_col("產出等級", ["產出等級", "产出等级"])
        next_proc_c = get_col("下製程", ["下製程", "下制程"])

        try:
            for col in [line_c, out_grade_c, next_proc_c]:
                if col not in df.columns: df[col] = "-"
                df[col] = df[col].fillna("-").astype(str)

            for col in [outer_cut, inner_cut, cgl_t, cgl_w, cgl_l, ccl_t, ccl_w, ccl_l]:
                if col not in df.columns: df[col] = 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # --- ROUTING LOGIC (RESTORED) ---
            df['is_first_baby'] = ~df.duplicated(subset=[order_c, baby_c], keep='first')
            df[ccl_l] = df.apply(lambda r: r[ccl_l] if r['is_first_baby'] else 0, axis=1)
            df['base_coil'] = df[mother_c].astype(str).str[:-3]
            df['is_x00'] = df[mother_c].astype(str).str.endswith('X00', na=False)
            df['family_has_x00'] = df.groupby([order_c, 'base_coil'])['is_x00'].transform('any')
            df['family_has_cgl'] = df.groupby([order_c, 'base_coil'])[cgl_l].transform(lambda x: (x > 0).any())
            df['is_first_mother'] = ~df.duplicated(subset=[order_c, mother_c])
            df['sum_ccl_by_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')

            def resolve_input(row):
                if not row['is_first_mother']: return 0
                if row[cgl_l] > 0:
                    if row['family_has_x00'] and not row['is_x00']: return 0
                    return row[cgl_l]
                if row[cgl_l] == 0:
                    if row['family_has_cgl']: return 0
                    return row['sum_ccl_by_mother']
                return 0
            
            df[cgl_l] = df.apply(resolve_input, axis=1)
            df[outer_cut] = df.apply(lambda r: r[outer_cut] if r['is_first_mother'] else 0, axis=1)
            df[inner_cut] = df.apply(lambda r: r[inner_cut] if r['is_first_mother'] else 0, axis=1)
            df[cgl_t] = df.groupby([order_c, 'base_coil'])[cgl_t].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)
            df[cgl_w] = df.groupby([order_c, 'base_coil'])[cgl_w].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            s1 = df.groupby([order_c, mother_c]).agg({
                cgl_t: 'mean', cgl_w: 'mean', cgl_l: 'first',
                ccl_t: 'mean', ccl_w: 'mean', ccl_l: 'sum',
                outer_cut: 'max', inner_cut: 'max' 
            }).reset_index()

            summary = s1.groupby(order_c).agg({
                mother_c: 'count', cgl_l: 'sum', ccl_l: 'sum', 
                outer_cut: 'sum', inner_cut: 'sum',
                ccl_t: 'mean', cgl_t: 'mean', cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty (Coils)', cgl_l: 'In_m', ccl_l: 'Out_m'})
            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
            summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- UI: ORDER SUMMARY ---
            st.subheader("1. Order Summary")
            disp = summary[[order_c, 'Qty (Coils)', cgl_w, 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
            disp.columns = ['Order ID', 'Qty (Coils)', 'Input Width (mm)', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
            disp = disp.sort_values(by='Cut Scrap (m)', ascending=False).reset_index(drop=True)
            disp.insert(0, 'No.', range(1, len(disp) + 1))
            st.dataframe(disp.set_index('No.').style.format({
                "Input Width (mm)": "{:,.0f}", "Input (m)": "{:,.0f}", "Cut Scrap (m)": "{:,.0f}", 
                "Output (m)": "{:,.0f}", "Diff (m)": "{:,.0f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:,.0f}"
            }), height=500, use_container_width=True)

           # --- UI: PRODUCTION COIL DETAILS ---
            st.divider()
            st.subheader("2. Production Coil Details") 
            sel_order = st.selectbox("🔍 Select Order ID:", options=df[order_c].unique(), index=None)
            if sel_order:
                det = df[df[order_c] == sel_order].copy()
                det['Var'] = det[ccl_t] - det[cgl_t]
                det_f = det[[mother_c, baby_c, line_c, out_grade_c, next_proc_c, cgl_t, cgl_w, cgl_l, outer_cut, inner_cut, ccl_t, ccl_w, 'Var', ccl_l]].copy()
                det_f.columns = ['Input ID', 'Output ID', 'Line', 'Grade', 'Next Proc', 'In Thick', 'In Width', 'In Len', 'Outer Cut', 'Inner Cut', 'Out Thick', 'Out Width', 'Thick Dev', 'Out Len']
                st.dataframe(det_f.style.format({
                    "In Thick": "{:.3f}", "In Width": "{:,.0f}", "In Len": "{:,.0f}",
                    "Outer Cut": "{:,.0f}", "Inner Cut": "{:,.0f}", "Out Thick": "{:.3f}", 
                    "Out Width": "{:,.0f}", "Thick Dev": "{:.3f}", "Out Len": "{:,.0f}"
                }), height=500, use_container_width=True)
                
            # --- UI: VISUAL INSIGHTS ---
            st.divider()
            st.subheader("3. Visual Insights & Analysis")
            st.plotly_chart(px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', color_continuous_scale='Blues_r', template=plotly_template), use_container_width=True)
            st.info("**分析結論:** 監控各訂單的塗層面積偏差。")

            st.plotly_chart(px.bar(disp.sort_values(by='Cut Scrap (m)', ascending=False), x='Order ID', y='Cut Scrap (m)', color='Cut Scrap (m)', color_continuous_scale='Reds', template=plotly_template), use_container_width=True)
            st.error("**分析結論:** 各訂單的剪切廢料總量。")

            # --- UI: EXECUTIVE SUMMARY ---
            st.divider()
            st.subheader("4. Executive Summary")
            t_in, t_out = disp['Input (m)'].sum(), disp['Output (m)'].sum()
            area_s = abs(disp[disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())
            st.markdown(f"**生產產出綜合分析:** \n* **總投入 (Total Input):** {t_in:,.0f} m  \n* **總產出 (Total Output):** {t_out:,.0f} m  \n* **不明面積差異 (Area Shortfall):** {area_s:,.2f} m²")

            # --- UI: EXPORT ---
            st.subheader("5. Export Data")
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                disp.to_excel(writer, sheet_name='Summary', index=False)
            st.download_button("📊 Download Excel Report", data=buf.getvalue(), file_name="Report.xlsx", type="primary")

        except Exception as e:
            st.error(f"Logic Error: {e}")
else:
    st.info("Please insert the Google Sheet Link.")
