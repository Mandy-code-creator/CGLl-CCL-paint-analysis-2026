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
            df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# =============================
# 2. CORE LOGIC
# =============================
if GSHEET_URL:
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        # Định nghĩa tên cột dựa trên tiêu đề tiếng Trung (đã lowercase)
        order_c, mother_c = "訂單號碼", "投入鋼捲號碼"
        cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌實測長度"
        ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"
        outer_cut, inner_cut = "outercutlength", "innercutlength"

        try:
            # 1. Chuyển đổi số liệu & Xử lý giá trị rỗng
            for col in [cgl_l, ccl_l, outer_cut, inner_cut, cgl_t, cgl_w, ccl_t, ccl_w]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 2. XỬ LÝ GỐC PHÔI (ROOT ID) - Chống cộng dồn A00/AP0/X00
            # Cắt bỏ 3 ký tự cuối để lấy ID gốc
            df['root_id'] = df[mother_c].astype(str).str[:-3]
            
            # 3. LOGIC TÍNH INPUT THÔNG MINH
            # Bước A: Kiểm tra xem trong cụm root_id này có dòng nào có dữ liệu CGL thực tế không?
            df['has_cgl_val'] = df[cgl_l] > 0
            root_has_cgl = df.groupby([order_c, 'root_id'])['has_cgl_val'].transform('any')
            
            # Bước B: Xác định dòng đầu tiên của cụm root_id để gán Input
            df['is_first_of_root'] = ~df.duplicated(subset=[order_c, 'root_id'])

            def resolve_input_length(row):
                # Nếu là dòng đầu tiên của cụm root_id
                if row['is_first_of_root']:
                    # Nếu có dữ liệu CGL thực tế -> Lấy CGL
                    if row[cgl_l] > 0:
                        return row[cgl_l]
                    # Nếu cả cụm root_id này không có dòng nào có CGL (Mồ côi & Trống số)
                    # -> Lấy Tổng Output của chính cuộn mẹ này làm Input
                    elif not row['root_has_cgl_any']:
                        return row['sum_ccl_by_mother']
                return 0 # Các dòng sau của cùng root_id luôn bằng 0 để chống cộng dồn

            # Hỗ trợ tính tổng CCL theo từng cuộn mẹ để đắp vào dòng mồ côi
            df['root_has_cgl_any'] = root_has_cgl
            df['sum_ccl_by_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')
            
            df['final_input'] = df.apply(resolve_input_length, axis=1)

            # 4. SAO CHÉP THÔNG SỐ RỘNG/DÀY (Cho các dòng mồ côi)
            df[cgl_t] = df.groupby([order_c, 'root_id'])[cgl_t].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)
            df[cgl_w] = df.groupby([order_c, 'root_id'])[cgl_w].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            # --- GOM NHÓM KẾT QUẢ ---
            summary = df.groupby(order_c).agg({
                mother_c: 'count',
                'final_input': 'sum',
                ccl_l: 'sum',
                outer_cut: 'sum',
                inner_cut: 'sum',
                cgl_t: 'mean', cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty', 'final_input': 'In_m', ccl_l: 'Out_m'})
            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
            # Công thức: Diff = Output - (Input - Cut)
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
            summary['Diff_Area'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- HIỂN THỊ ---
            st.subheader("1. Order Summary (Logic Root ID & Orphan Protection)")
            
            disp = summary[[order_c, 'Qty', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Diff_Area']].copy()
            disp.columns = ['Order ID', 'Coils', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Diff Area (m²)']
            disp = disp.sort_values(by='Diff (m)', ascending=True).reset_index(drop=True)
            disp.insert(0, 'No.', range(1, len(disp) + 1))

            st.table(disp.set_index('No.').style.format({
                "Input (m)": "{:,.0f}", "Cut Scrap (m)": "{:,.0f}", 
                "Output (m)": "{:,.0f}", "Diff (m)": "{:.2f}", "Diff Area (m²)": "{:.2f}"
            }))

            # Biểu đồ
            st.divider()
            f1 = px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                        color_continuous_scale='RdBu', title="Extra Area per Order")
            st.plotly_chart(f1, use_container_width=True)

            # Export
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                disp.to_excel(writer, sheet_name='Summary', index=False)
            st.download_button("📊 Download Report Excel", data=buf.getvalue(), file_name="Length_Report.xlsx", type="primary")

        except Exception as e:
            st.error(f"Logic Error: {e}")
else:
    st.info("Please check the GSHEET_URL in the code.")
