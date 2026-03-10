import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis: Total CGL vs CCL per Order", layout="wide")

# ==========================================================
# 1. THEME SELECTION & CLEAN ANIMATED BACKGROUND
# ==========================================================
theme_choice = st.radio("🎨 Select App Theme:", ["Light Mode (Standard)", "Dark Mode (Professional)"], horizontal=True)

if theme_choice == "Dark Mode (Professional)":
    grad_colors = "#0f172a, #1e3a8a, #312e81, #0f172a"
    card_bg = "rgba(30, 41, 59, 0.4)" 
    text_color = "#f8fafc"
    sub_text = "#cbd5e1"
    plotly_template = "plotly_dark"
else:
    grad_colors = "#f8fafc, #f1f5f9, #e2e8f0, #f8fafc"
    card_bg = "rgba(255, 255, 255, 0.6)"
    text_color = "#1e3a8a"
    sub_text = "#475569"
    plotly_template = "plotly_white"

st.markdown(f"""
<style>
    /* Nền chuyển động mượt mà */
    .stApp {{
        background: linear-gradient(-45deg, {grad_colors});
        background-size: 400% 400%;
        animation: gradient_move 20s ease infinite;
    }}
    @keyframes gradient_move {{
        0% {{ background-position: 0% 50%; }}
        50% {{ background-position: 100% 50%; }}
        100% {{ background-position: 0% 50%; }}
    }}

    /* XÓA BỎ HOÀN TOÀN CÁC KHUNG CHỒNG BẢNG */
    /* Chỉ áp dụng style cho chính vùng chứa biểu đồ và bảng, không áp dụng cho div bọc ngoài */
    [data-testid="stVerticalBlock"] > div {{
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }}

    /* Tạo nền mờ duy nhất cho Bảng và Biểu đồ */
    .stDataFrame, .js-plotly-plot {{
        background-color: {card_bg} !important;
        backdrop-filter: blur(10px);
        border-radius: 12px;
        border: 1px solid rgba(255,255,255,0.1) !important;
        padding: 10px;
    }}

    /* Chỉnh chữ */
    h1, h2, h3 {{ color: {text_color} !important; font-family: 'Segoe UI', sans-serif; }}
    .stMarkdown p {{ color: {sub_text} !important; }}
    
    /* Gọn hóa thanh cuộn */
    ::-webkit-scrollbar {{ width: 5px; height: 5px; }}
    ::-webkit-scrollbar-thumb {{ background: rgba(148, 163, 184, 0.3); border-radius: 10px; }}
</style>
""", unsafe_allow_html=True)

st.title("Length Variance Analysis: Total CGL vs CCL per Order")

# ==========================================================
# 2. CORE LOGIC (GIỮ NGUYÊN LOGIC CHUẨN CỦA BẠN)
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

@st.cache_data(ttl=300)
def load_auto_data(url):
    try:
        base_url = url.split('/edit')[0]
        gid = url.split("gid=")[1].split("&")[0] if "gid=" in url else "0"
        csv_url = f"{base_url}/export?format=csv&gid={gid}"
        df = pd.read_csv(csv_url)
        df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(r'\s+', '', regex=True)
        return df
    except: return None

if GSHEET_URL:
    df = load_auto_data(GSHEET_URL)
    if df is not None:
        def get_col(d, names):
            for n in names:
                if n in df.columns: return n
            return d

        order_c, mother_c, baby_c = get_col("訂單號碼", ["訂單號碼", "订单号码"]), get_col("投入鋼捲號碼", ["投入鋼捲號碼", "投入钢卷号码"]), get_col("產出鋼捲號碼", ["產出鋼捲號碼", "产出钢卷号码"])
        cgl_l, ccl_l, cgl_w = get_col("镀锌測長度", ["镀锌測長度", "镀锌长度"]), get_col("實測長度", ["實測長度"]), get_col("镀锌測寬度", ["镀锌測寬度", "镀锌宽度"])
        cgl_t, ccl_t = get_col("镀锌實測厚度", ["镀锌厚度"]), get_col("實測厚度", ["实测厚度"])
        outer_cut, inner_cut = get_col("outercutlength", ["outercut"]), get_col("innercutlength", ["innercut"])
        line_c, out_grade_c, next_proc_c = get_col("線別", ["線別"]), get_col("產出等級", ["產出等級"]), get_col("下製程", ["下製程"])

        try:
            for col in [line_c, out_grade_c, next_proc_c]: df[col] = df[col].fillna("-").astype(str)
            for col in [outer_cut, inner_cut, cgl_t, cgl_w, cgl_l, ccl_t, ccl_l]: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # --- ROUTING LOGIC ---
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

            # AGGREGATION
            s1 = df.groupby([order_c, mother_c]).agg({cgl_t: 'mean', cgl_w: 'mean', cgl_l: 'first', ccl_t: 'mean', ccl_l: 'sum', outer_cut: 'max', inner_cut: 'max'}).reset_index()
            summary = s1.groupby(order_c).agg({mother_c: 'count', cgl_l: 'sum', ccl_l: 'sum', outer_cut: 'sum', inner_cut: 'sum', ccl_t: 'mean', cgl_t: 'mean', cgl_w: 'mean'}).reset_index()
            summary = summary.rename(columns={mother_c: 'Qty', cgl_l: 'In_m', ccl_l: 'Out_m'})
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - (summary[outer_cut] + summary[inner_cut]))
            summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- DISPLAY ---
            st.subheader("1. Order Summary")
            st.dataframe(summary[[order_c, 'Qty', cgl_w, 'In_m', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].style.format({
                "In_m": "{:,.0f}", "Out_m": "{:,.0f}", "Diff": "{:,.0f}", "Area_m2": "{:,.0f}", "cgl_w": "{:,.0f}", "Thick_Var": "{:.3f}"
            }), height=450, use_container_width=True)

            st.divider()
            st.subheader("2. Production Coil Details")
            sel_order = st.selectbox("Select Order ID:", options=df[order_c].unique(), index=None)
            if sel_order:
                det = df[df[order_c] == sel_order].copy()
                det['Var'] = det[ccl_t] - det[cgl_t]
                det_f = det[[mother_c, baby_c, line_c, out_grade_c, next_proc_c, cgl_t, cgl_w, cgl_l, outer_cut, inner_cut, ccl_t, 'Var', ccl_l]]
                det_f.columns = ['Input ID', 'Output ID', 'Line', 'Grade', 'Next Proc', 'In Thick', 'In Width', 'In Len', 'Outer Cut', 'Inner Cut', 'Out Thick', 'Thick Dev', 'Out Len']
                st.dataframe(det_f.style.format({
                    "In Thick": "{:.3f}", "Out Thick": "{:.3f}", "Thick Dev": "{:.3f}",
                    "In Width": "{:,.0f}", "In Len": "{:,.0f}", "Out Len": "{:,.0f}"
                }), height=450, use_container_width=True)

            st.divider()
            st.subheader("3. Visual Insights & Analysis")
            st.plotly_chart(px.bar(summary, x=order_c, y='Area_m2', color='Diff', color_continuous_scale='Blues_r', template=plotly_template, title="Extra Area per Order"), use_container_width=True)
            st.info("**分析結論:** 監控各訂單的塗層面積偏差。")

        except Exception as e:
            st.error(f"Logic Error: {e}")
