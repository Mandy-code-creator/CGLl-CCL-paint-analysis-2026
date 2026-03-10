import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis: Total CGL vs CCL per Order", layout="wide")

# ==========================================================
# 1. AUTO-SYNC CONFIGURATION (INSERT YOUR LINK HERE)
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

# --- MINIMALIST DESIGN ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
        background-color: #ffffff; padding: 20px; border-radius: 0px;
        margin-bottom: 20px; border: none;
    }
    h1, h2, h3 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
    table { width: 100% !important; border-collapse: collapse !important; font-family: 'Segoe UI', sans-serif; color: #334155; border: 1px solid #e2e8f0 !important; }
    th { border: 1px solid #e2e8f0 !important; color: #1e3a8a !important; text-align: center !important; padding: 12px 8px !important; font-size: 13px !important; background-color: #f8fafc !important; }
    td { text-align: center !important; padding: 10px 8px !important; border: 1px solid #e2e8f0 !important; font-size: 13px !important; }
    tr:hover { background-color: #f1f5f9; }
    </style>
    """, unsafe_allow_html=True)

st.title("Length Variance Analysis: Total CGL vs CCL per Order")

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
            # Giữ nguyên tên gốc để hiển thị nhưng tạo bản copy chuẩn hóa
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# =============================
# 2. CORE LOGIC
# =============================
if GSHEET_URL:
    df_raw = load_auto_data(GSHEET_URL)
    
    if df_raw is not None:
        df = df_raw.copy()
        # Chuẩn hóa tên cột để tìm kiếm (xóa khoảng trắng, lowercase)
        cols = {c: c.strip().lower().replace(" ", "") for c in df.columns}
        
        # Hàm tìm cột thông minh dựa trên từ khóa
        def find_col(keywords, default_name):
            for original, normalized in cols.items():
                if all(k in normalized for k in keywords):
                    return original
            return default_name

        # Định nghĩa các cột chính bằng từ khóa để tránh lỗi 'Logic Error'
        order_c = find_col(["订单", "号码"], "訂單號碼")
        mother_c = find_col(["投入", "钢卷"], "投入鋼捲號碼")
        cgl_l = find_col(["镀锌", "长度"], "镀锌實測長度")
        ccl_l = find_col(["实测", "长度"], "實測長度")
        cgl_t = find_col(["镀锌", "厚度"], "镀锌實測厚度")
        ccl_t = find_col(["实测", "厚度"], "實測厚度")
        cgl_w = find_col(["镀锌", "宽度"], "镀锌測寬度")
        outer_cut = find_col(["outer", "cut"], "outercutlength")
        inner_cut = find_col(["inner", "cut"], "innercutlength")

        try:
            # 1. Chuyển đổi số liệu
            for col in [cgl_l, ccl_l, outer_cut, inner_cut, cgl_t, cgl_w, ccl_t]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                else:
                    df[col] = 0

            # 2. XỬ LÝ GỐC PHÔI (ROOT ID) - Chống cộng dồn
            df['root_id'] = df[mother_c].astype(str).str[:-3]
            
            # 3. LOGIC TÍNH INPUT THÔNG MINH
            df['has_cgl_val'] = df[cgl_l] > 0
            root_has_cgl = df.groupby([order_c, 'root_id'])['has_cgl_val'].transform('any')
            df['is_first_of_root'] = ~df.duplicated(subset=[order_c, 'root_id'])
            df['sum_ccl_by_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')
            df['root_has_cgl_any'] = root_has_cgl

            def resolve_input_length(row):
                if row['is_first_of_root']:
                    if row[cgl_l] > 0:
                        return row[cgl_l]
                    elif not row['root_has_cgl_any']:
                        return row['sum_ccl_by_mother'] # Trường hợp mồ mồ côi trống số
                return 0

            df['final_input'] = df.apply(resolve_input_length, axis=1)

            # --- GOM NHÓM KẾT QUẢ ---
            summary = df.groupby(order_c).agg({
                mother_c: 'count',
                'final_input': 'sum',
                ccl_l: 'sum',
                outer_cut: 'sum',
                inner_cut: 'sum',
                cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty', 'final_input': 'In_m', ccl_l: 'Out_m'})
            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
            summary['Diff_Area'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- HIỂN THỊ ---
            st.subheader("1. Order Summary")
            disp = summary[[order_c, 'Qty', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Diff_Area']].copy()
            disp.columns = ['Order ID', 'Coils', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Diff Area (m²)']
            
            st.table(disp.style.format({
                "Input (m)": "{:,.0f}", "Output (m)": "{:,.0f}", "Diff (m)": "{:.2f}", "Diff Area (m²)": "{:.2f}"
            }))

            st.divider()
            st.subheader("2. Visual Insights")
            f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                        color_continuous_scale='RdBu', title="Extra Area per Order")
            st.plotly_chart(f1, use_container_width=True)

        except Exception as e:
            st.error(f"Logic Error: {e}. Vui lòng kiểm tra lại tiêu đề các cột trên Google Sheet.")
