import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis", layout="wide")

# ==========================================================
# 1. AUTO-SYNC CONFIGURATION
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
        # Chuẩn hóa tên cột để tìm kiếm tự động
        raw_cols = {c: str(c).strip().lower().replace(" ", "") for c in df.columns}
        
        def find_col(keywords, default):
            for orig, norm in raw_cols.items():
                if all(k in norm for k in keywords): return orig
            return default

        # Tự động nhận diện cột (Tránh lỗi Logic Error)
        order_c = find_col(["订单", "号码"], "訂單號碼")
        mother_c = find_col(["投入", "钢卷"], "投入鋼捲號碼")
        baby_c   = find_col(["产出", "钢卷"], "產出鋼捲號碼")
        cgl_l    = find_col(["镀锌", "长度"], "镀锌實測長度")
        ccl_l    = find_col(["实测", "长度"], "實測長度")
        cgl_t    = find_col(["镀锌", "厚度"], "镀锌實測厚度")
        cgl_w    = find_col(["镀锌", "宽度"], "镀锌測寬度")
        outer_cut = find_col(["outer", "cut"], "outercutlength")
        inner_cut = find_col(["inner", "cut"], "innercutlength")

        try:
            # Ép kiểu dữ liệu số
            numeric_cols = [cgl_l, ccl_l, outer_cut, inner_cut, cgl_t, cgl_w]
            for c in numeric_cols:
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
                else:
                    df[c] = 0

            # --- THUẬT TOÁN LOGIC TỔNG HỢP ---
            
            # 1. Tách "Huyết thống" (Root ID) - Cắt bỏ 3 ký tự cuối
            # Ví dụ: 61A983A00 và 61A983AP0 -> chung gốc 61A983
            df['root_id'] = df[mother_c].astype(str).str[:-3]
            
            # 2. Kiểm tra xem nhóm root_id này có cuộn mẹ gốc (X00/A00) trong đơn hàng không
            df['is_main_coil'] = df[mother_c].astype(str).str.contains('X00|A00', case=False)
            df['has_main_in_order'] = df.groupby([order_c, 'root_id'])['is_main_coil'].transform('any')
            
            # 3. Loại bỏ cộng dồn trên cùng một Mother ID thực tế
            df['is_first_mother'] = ~df.duplicated(subset=[order_c, mother_c])
            df['input_step1'] = df.apply(lambda r: r[cgl_l] if r['is_first_mother'] else 0, axis=1)

            # 4. Quyết định Input cuối cùng (final_input)
            def final_logic(row):
                # Nếu có số CGL thực tế > 0
                if row['input_step1'] > 0:
                    # Nếu là cuộn con nhưng đơn hàng đã có cuộn mẹ (X00/A00) gánh team -> trả về 0
                    if not row['is_main_coil'] and row['has_main_in_order']:
                        return 0
                    return row['input_step1']
                # Nếu trống số CGL
                else:
                    # Nếu nhóm này "mồ côi" hoàn toàn (không có X00/A00)
                    if not row['has_main_in_order']:
                        # Lấy số CCL đắp qua cho dòng đầu tiên xuất hiện
                        return row[ccl_l] if row['is_first_mother'] else 0
                    return 0

            df['final_input'] = df.apply(final_logic, axis=1)

            # --- SAO CHÉP THÔNG SỐ RỘNG/DÀY ---
            df[cgl_w] = df.groupby([order_c, 'root_id'])[cgl_w].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            # --- TỔNG HỢP ---
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

            # --- GIAO DIỆN ---
            st.subheader("1. Order Summary (Logic Đa Tầng Hoàn Chỉnh)")
            disp = summary[[order_c, 'Qty', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Diff_Area']].copy()
            disp.columns = ['Order ID', 'Coils', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Diff Area (m²)']
            
            st.table(disp.sort_values('Diff (m)').style.format({
                "Input (m)": "{:,.0f}", "Output (m)": "{:,.0f}", "Diff (m)": "{:.2f}", "Diff Area (m²)": "{:.2f}"
            }))

            st.divider()
            st.subheader("2. Kiểm tra chi tiết để đối chiếu")
            sel_order = st.selectbox("Chọn Order ID để xem chi tiết từng cuộn:", options=df[order_c].unique())
            if sel_order:
                check = df[df[order_c] == sel_order][[mother_c, baby_c, cgl_l, "final_input", ccl_l]]
                st.dataframe(check)

            # Biểu đồ phân tích
            st.divider()
            f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                        color_continuous_scale='RdBu', title="Extra Area per Order")
            st.plotly_chart(f1, use_container_width=True)

        except Exception as e:
            st.error(f"Error: {e}")
