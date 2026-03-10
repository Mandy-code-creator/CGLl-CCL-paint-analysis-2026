import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis", layout="wide")

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
            df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

if GSHEET_URL and GSHEET_URL != "CHÈN_LINK_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        # Định nghĩa cột dựa trên tiêu đề tiếng Trung đã lowercase
        order_c, mother_c = "訂單號碼", "投入鋼捲號碼"
        cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"
        outer_cut, inner_cut = "outercutlength", "innercutlength"

        try:
            # Chuyển đổi số liệu
            for col in [cgl_l, ccl_l, outer_cut, inner_cut, cgl_t, cgl_w, ccl_t, ccl_w]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # --- THUẬT TOÁN LOGIC TỔNG HỢP ---
            # A. Tách gốc phôi (Base ID)
            df['base_id'] = df[mother_c].astype(str).str.replace(r'[A-Za-z]\d{2}$', '', regex=True)
            
            # B. Kiểm tra sự hiện diện của cuộn mẹ (X00 hoặc A00) trong mỗi Base ID của đơn hàng
            df['is_main'] = df[mother_c].astype(str).str.contains('X00|A00', case=False)
            main_present = df.groupby([order_c, 'base_id'])['is_main'].transform('any')
            
            # C. Xử lý Input Length (cgl_l) - QUAN TRỌNG NHẤT
            def calculate_input(row):
                # 1. Nếu là dòng đầu tiên xuất hiện của cuộn mẹ/cuộn xẻ có số CGL > 0 -> Lấy số đó
                # Dùng thuộc tính trùng lặp để chỉ lấy số CGL ở dòng đầu tiên của mỗi Mother ID
                return row[cgl_l]

            # Loại bỏ cộng dồn trên cùng một Mother ID
            df['input_step1'] = df.groupby([order_c, mother_c])[cgl_l].transform('first')
            mask_duplicate_mother = df.duplicated(subset=[order_c, mother_c])
            df.loc[mask_duplicate_mother, 'input_step1'] = 0

            # D. Xử lý logic cuộn con/mồ côi
            def final_input_logic(row):
                # Nếu cuộn có số CGL thực tế -> Lấy số đó (Case 1, 3)
                if row['input_step1'] > 0:
                    # Nhưng nếu nó là cuộn con (không phải X00/A00) và đơn hàng ĐÃ CÓ cuộn mẹ gánh
                    # thì phải trả về 0 để không cộng dồn kép
                    if not row['is_main'] and row['has_main_in_order']:
                        return 0
                    return row['input_step1']
                
                # Nếu cuộn trống số CGL (input_step1 == 0) (Case 4)
                else:
                    # Nếu đơn hàng KHÔNG có cuộn mẹ nào gánh cho Base ID này -> Mồ côi tuyệt đối
                    if not row['has_main_in_order']:
                        # Chỉ lấy CCL đắp qua cho dòng đầu tiên của cuộn mồ côi này
                        return row[ccl_l] if not row['is_duplicated_mother'] else 0
                    return 0

            # Tiền xử lý các cờ kiểm tra
            df['has_main_in_order'] = main_present
            df['is_duplicated_mother'] = df.duplicated(subset=[order_c, mother_c])
            
            df['final_input'] = df.apply(final_input_logic, axis=1)

            # --- SAO CHÉP ĐỘ DÀY/RỘNG ---
            df[cgl_t] = df.groupby([order_c, 'base_id'])[cgl_t].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)
            df[cgl_w] = df.groupby([order_c, 'base_id'])[cgl_w].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            # --- GOM NHÓM KẾT QUẢ ---
            summary = df.groupby(order_c).agg({
                mother_c: 'count',
                'final_input': 'sum',
                ccl_l: 'sum',
                outer_cut: 'sum',
                inner_cut: 'sum',
                cgl_t: 'mean', ccl_t: 'mean', cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty', 'final_input': 'In_m', ccl_l: 'Out_m'})
            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
            summary['Diff_Area'] = (summary[cgl_w] / 1000) * summary['Diff']

            # Hiển thị bảng
            st.subheader("1. Order Summary (Logic Đa Tầng)")
            disp = summary[[order_c, 'Qty', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Diff_Area']].copy()
            disp.columns = ['Order ID', 'Coils', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Diff Area (m²)']
            
            st.table(disp.sort_values('Diff (m)').style.format({
                "Input (m)": "{:,.0f}", "Output (m)": "{:,.0f}", "Diff (m)": "{:.2f}", "Diff Area (m²)": "{:.2f}"
            }))

            # --- CHI TIẾT ĐỂ KIỂM CHỨNG ---
            st.divider()
            st.subheader("2. Kiểm tra chi tiết từng cuộn")
            sel_order = st.selectbox("Chọn đơn hàng:", options=df[order_c].unique())
            if sel_order:
                check = df[df[order_c] == sel_order][[mother_c, "產出鋼捲號碼", cgl_l, "final_input", ccl_l]]
                st.dataframe(check)

        except Exception as e:
            st.error(f"Logic Error: {e}")
