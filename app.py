import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Dynamic Gradient Analytics", layout="wide")

# ==========================================================
# 1. SIÊU CẤU HÌNH NỀN CHUYỂN ĐỘNG (ANIMATED GRADIENT)
# ==========================================================
# Tạo hiệu ứng dải màu trôi bồng bềnh 15 giây/chu kỳ
st.markdown("""
    <style>
    /* Ép màu nền vào lớp phủ toàn diện của Streamlit */
    .stApp {
        background: linear-gradient(-45deg, #0f172a, #1e3a8a, #5b21b6, #0d9488);
        background-size: 400% 400%;
        animation: gradient_move 15s ease infinite;
        color: white;
    }

    @keyframes gradient_move {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Làm các bảng dữ liệu trở nên trong suốt kiểu thủy tinh (Glassmorphism) */
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stDataFrame) {
        background-color: rgba(255, 255, 255, 0.05); 
        backdrop-filter: blur(15px);
        padding: 25px; 
        border-radius: 15px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin-bottom: 20px;
    }

    /* Chỉnh màu chữ cho khớp với nền tối chuyển động */
    h1, h2, h3, p, span, label {
        color: #ffffff !important;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
    }
    
    /* Chỉnh màu cho thanh selectbox */
    .stSelectbox div[data-baseweb="select"] {
        background-color: rgba(255,255,255,0.1);
        border-radius: 8px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🚀 Dynamic Length Variance Analytics")

# ==========================================================
# 2. CORE LOGIC (GIỮ NGUYÊN LOGIC CHUẨN CỦA BẠN)
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

@st.cache_data(ttl=300)
def load_auto_data(url):
    try:
        if "docs.google.com/spreadsheets" in url:
            base_url = url.split('/edit')[0]
            gid = url.split("gid=")[1].split("&")[0] if "gid=" in url else "0"
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

        # MAPPING COLUMNS
        order_c = get_col("訂單號碼", ["訂單號碼", "订单号码"])
        mother_c = get_col("投入鋼捲號碼", ["投入鋼捲號碼", "投入钢卷号码"])
        baby_c = get_col("產出鋼捲號碼", ["產出鋼捲號碼", "产出钢卷号码"])
        cgl_l = get_col("镀锌測長度", ["镀锌測長度", "镀锌實測長度", "镀锌长度"])
        ccl_l = get_col("實測長度", ["實測長度", "实测长度"])
        cgl_w = get_col("镀锌測寬度", ["镀锌測寬度", "镀锌測寬", "镀锌宽度"])
        cgl_t = get_col("镀锌實測厚度", ["镀锌實測厚度", "镀锌測厚", "镀锌厚度"])
        ccl_w = get_col("實測寬度", ["實測寬度", "实测宽度"])
        ccl_t = get_col("實測厚度", ["實測厚度", "实测厚度"])
        outer_cut = get_col("outercutlength", ["outercutlength", "outercut"])
        inner_cut = get_col("innercutlength", ["innercutlength", "innercut"])
        line_c = get_col("線別", ["線別", "线别"])
        out_grade_c = get_col("產出等級", ["產出等級", "产出等级"])
        next_proc_c = get_col("下製程", ["下製程", "下制程"])

        try:
            # Data Cleaning
            for col in [line_c, out_grade_c, next_proc_c]:
                df[col] = df[col].fillna("-").astype(str)
            for col in [outer_cut, inner_cut, cgl_t, cgl_w, cgl_l, ccl_t, ccl_w, ccl_l]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # --- YOUR ORIGINAL ROUTING LOGIC ---
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

            # --- AGGREGATION ---
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

            summary = summary.rename(columns={mother_c: 'Qty', cgl_l: 'In_m', ccl_l: 'Out_m'})
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - (summary[outer_cut] + summary[inner_cut]))
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- DISPLAY ---
            st.subheader("📊 1. Order Summary")
            st.dataframe(summary[[order_c, 'Qty', cgl_w, 'In_m', 'Out_m', 'Diff', 'Area_m2']].style.format({
                "In_m": "{:,.0f}", "Out_m": "{:,.0f}", "Diff": "{:,.0f}", "Area_m2": "{:,.0f}"
            }), height=500, use_container_width=True)

            st.divider()
            st.subheader("🔍 2. Production Coil Details")
            sel_order = st.selectbox("Select Order ID:", options=df[order_c].unique(), index=None)
            if sel_order:
                det_f = df[df[order_c] == sel_order][[mother_c, baby_c, line_c, out_grade_c, next_proc_c, cgl_t, cgl_w, ccl_t, ccl_l]]
                st.dataframe(det_f.style.format({"cgl_t": "{:.3f}", "ccl_t": "{:.3f}"}), height=500, use_container_width=True)

            st.divider()
            st.subheader("📈 3. Visual Insights")
            st.plotly_chart(px.bar(summary, x=order_c, y='Area_m2', color='Diff', color_continuous_scale='Turbo', template="plotly_dark"), use_container_width=True)

        except Exception as e:
            st.error(f"Logic Error: {e}")
