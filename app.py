import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis", layout="wide")

# ==========================================================
# 1. AUTO-SYNC CONFIGURATION
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

# --- CSS THIẾT KẾ NGUYÊN BẢN ---
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
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# =============================
# 2. CORE LOGIC (VẠN NĂNG & CHÍNH XÁC)
# =============================
if GSHEET_URL:
    df_raw = load_auto_data(GSHEET_URL)
    
    if df_raw is not None:
        df = df_raw.copy()
        
        # Tìm cột thông minh
        raw_cols = {c: str(c).strip().lower().replace(" ", "") for c in df.columns}
        def find_col(keywords, default):
            for orig, norm in raw_cols.items():
                if all(k in norm for k in keywords): return orig
            return default

        order_c = find_col(["订单", "号码"], "訂單號碼")
        mother_c = find_col(["投入", "钢卷"], "投入鋼捲號碼")
        baby_c   = find_col(["产出", "钢卷"], "產出鋼捲號碼")
        cgl_l    = find_col(["镀锌", "长度"], "镀锌實測長度")
        ccl_l    = find_col(["实测", "长度"], "實測長度")
        cgl_w    = find_col(["镀锌", "宽度"], "镀锌測寬度")
        outer_cut = find_col(["outer", "cut"], "outercutlength")
        inner_cut = find_col(["inner", "cut"], "innercutlength")

        try:
            # 1. Ép kiểu số
            for c in [cgl_l, ccl_l, outer_cut, inner_cut, cgl_w]:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

            # 2. XÁC ĐỊNH MÃ GỐC (ROOT ID) - Dùng Regex để cắt bỏ hậu tố linh hoạt
            # Ví dụ: 4BC133X00 -> 4BC133, 61A983AP0 -> 61A983
            def get_root(s):
                return re.sub(r'[A-Za-z]\d{2}$', '', str(s))
            df['root_id'] = df[mother_c].apply(get_root)
            
            # 3. KIỂM TRA MỒ CÔI (ORPHAN)
            # Một nhóm phôi là mồ côi nếu KHÔNG có cuộn nào kết thúc bằng X00 hoặc A00
            df['is_main_id'] = df[mother_c].astype(str).str.contains(r'X00|A00', case=False, regex=True)
            df['group_has_main'] = df.groupby([order_c, 'root_id'])['is_main_id'].transform('any')

            # 4. TÍNH INPUT (FINAL_INPUT)
            # Chỉ lấy giá trị ở dòng đầu tiên xuất hiện của mỗi cụm (Order + Mother ID)
            df['is_first_entry'] = ~df.duplicated(subset=[order_c, mother_c])
            # Tính tổng Output của từng Mother ID để dùng cho trường hợp mồ côi
            df['sum_ccl_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')

            def resolve_logic(row):
                # Nếu không phải dòng đầu tiên của cuộn đó -> Input = 0 (Chống cộng dồn)
                if not row['is_first_entry']:
                    return 0
                
                # Nếu có số CGL thực tế (>0)
                if row[cgl_l] > 0:
                    # Nếu là cuộn con nhưng đơn hàng có cuộn mẹ X00/A00 gánh -> 0
                    if not row['is_main_id'] and row['group_has_main']:
                        return 0
                    return row[cgl_l]
                
                # Nếu trống số CGL và là cuộn mồ côi (không có X00/A00 trong đơn hàng)
                if row[cgl_l] == 0 and not row['group_has_main']:
                    # Lấy tổng CCL của cuộn đó làm Input
                    return row['sum_ccl_mother']
                
                return 0

            df['final_input'] = df.apply(resolve_logic, axis=1)

            # --- TỔNG HỢP ---
            summary = df.groupby(order_c).agg({
                mother_c: 'count', 'final_input': 'sum', ccl_l: 'sum',
                outer_cut: 'sum', inner_cut: 'sum', cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty', 'final_input': 'In_m', ccl_l: 'Out_m'})
            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
            summary['Diff_Area'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- HIỂN THỊ ---
            st.subheader("1. Order Summary")
            disp = summary[[order_c, 'Qty', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Diff_Area']]
            disp.columns = ['Order ID', 'Coils', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Diff Area (m²)']
            st.table(disp.style.format({
                "Input (m)": "{:,.0f}", "Output (m)": "{:,.0f}", "Diff (m)": "{:.2f}", "Diff Area (m²)": "{:.2f}"
            }))

            st.divider()
            st.subheader("2. Production Coil Details")
            sel_order = st.selectbox("🔍 Search Order ID:", options=df[order_c].unique(), index=None)
            if sel_order:
                check = df[df[order_c] == sel_order][[mother_c, baby_c, cgl_l, 'final_input', ccl_l]]
                check.columns = ['Input ID', 'Output ID', 'CGL Orig', 'Final Input', 'CCL Output']
                st.table(check.style.format({"CGL Orig": "{:,.0f}", "Final Input": "{:,.0f}", "CCL Output": "{:,.0f}"}))

            st.divider()
            f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                        color_continuous_scale='RdBu', title="Extra Area per Order")
            st.plotly_chart(f1, use_container_width=True)

        except Exception as e:
            st.error(f"Logic Error: {e}")
