import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- PAGE CONFIGURATION (Thiết kế nguyên bản) ---
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
# 2. CORE LOGIC (SỬA LỖI INPUT = 0)
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
            # 1. Chuyển đổi số liệu
            numeric_cols = [cgl_l, ccl_l, outer_cut, inner_cut, cgl_w]
            for c in numeric_cols:
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
                else:
                    df[c] = 0

            # 2. LOGIC TÍNH INPUT MỚI (KHẮC PHỤC LỖI TOÀN SỐ 0)
            # Bước A: Xác định Root ID (Dùng 7 ký tự đầu thay vì cắt 3 để chính xác hơn)
            df['root_id'] = df[mother_c].astype(str).str[:7]
            
            # Bước B: Đánh dấu dòng đầu tiên của mỗi Mother ID thực tế để lấy CGL
            df['is_first_mother'] = ~df.duplicated(subset=[order_c, mother_c])
            
            # Bước C: Kiểm tra xem nhóm phôi này có cuộn mẹ gốc (X00/A00) mang số CGL hay không
            df['is_main'] = df[mother_c].astype(str).str.contains('X00|A00', case=False)
            df['has_main_with_val'] = (df['is_main']) & (df[cgl_l] > 0)
            df['main_present'] = df.groupby([order_c, 'root_id'])['has_main_with_val'].transform('any')

            def resolve_input(row):
                # Ưu tiên 1: Nếu dòng có số CGL thực tế (>0)
                if row['is_first_mother'] and row[cgl_l] > 0:
                    # Nếu là cuộn con nhưng cùng nhóm đã có cuộn mẹ X00/A00 gánh số -> trả về 0
                    if not row['is_main'] and row['main_present']:
                        return 0
                    return row[cgl_l]
                
                # Ưu tiên 2: Nếu dòng trống số CGL nhưng là cuộn mồ côi (không có main trong nhóm)
                if row['is_first_mother'] and row[cgl_l] == 0:
                    if not row['main_present']:
                        # Lấy CCL đắp qua cho dòng đầu tiên xuất hiện
                        return row[ccl_l]
                return 0

            df['final_input'] = df.apply(resolve_input, axis=1)

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
                check.columns = ['Input ID', 'Output ID', 'CGL Original', 'Final Input', 'CCL Output']
                st.table(check.style.format({"CGL Original": "{:,.0f}", "Final Input": "{:,.0f}", "CCL Output": "{:,.0f}"}))

        except Exception as e:
            st.error(f"Error: {e}")
