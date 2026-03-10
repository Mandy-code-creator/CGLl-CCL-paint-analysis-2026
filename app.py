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
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; }
        table { border: 1px solid #000 !important; }
        th, td { border: 0.5pt solid #ccc !important; }
    }
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
            # Khử khoảng trắng để dễ xử lý, nhưng KHÔNG viết thường toàn bộ để hiển thị cho đẹp ở Sidebar
            df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

# =============================
# 2. CORE LOGIC
# =============================
if GSHEET_URL and GSHEET_URL != "CHÈN_LINK_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
    df = load_auto_data(GSHEET_URL)
    
    if df is not None:
        # --- THÊM SIDEBAR SETTINGS ĐỂ CHỌN CỘT AL ---
        st.sidebar.header("⚙️ Data Settings")
        st.sidebar.markdown("Vui lòng chọn cột tương đương với **Cột AL** trong Excel (Dùng để đắp số cho cuộn mồ côi):")
        al_col_sel = st.sidebar.selectbox(
            "Select Fallback Column (AL):", 
            options=["-- Bỏ qua (Set = 0) --"] + list(df.columns)
        )

        order_c, mother_c, baby_c = "訂單號碼", "投入鋼捲號碼", "產出鋼捲號碼"
        cgl_t, cgl_w, cgl_l = "镀锌實測厚度", "镀锌測寬度", "镀锌測長度"
        ccl_t, ccl_w, ccl_l = "實測厚度", "實測寬度", "實測長度"

        # Vì ta khử khoảng trắng nên check lại tên cột
        outer_cut = "outercutlength"
        inner_cut = "innercutlength"

        # Đổi tên các cột cut thành chữ thường trong df nội bộ để dễ tìm
        df.columns = [c.lower() if c.lower() in [outer_cut, inner_cut] else c for c in df.columns]

        try:
            for col in [outer_cut, inner_cut]:
                if col not in df.columns:
                    df[col] = 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            for col in [cgl_t, cgl_w, cgl_l, ccl_t, ccl_w, ccl_l]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                else:
                    df[col] = 0
                
            # --- TÁCH GỐC CUỘN ---
            df['base_coil'] = df[mother_c].astype(str).str.replace(r'[A-Za-z]\d{2}$', '', regex=True)
            
            is_main_coil = df[mother_c].astype(str).str.endswith('X00')
            main_coil_present = df[is_main_coil].groupby([order_c, 'base_coil']).size().reset_index(name='has_main')
            
            df = df.merge(main_coil_present, on=[order_c, 'base_coil'], how='left')
            df['has_main'] = df['has_main'].fillna(0) > 0

            # --- XỬ LÝ CỘT AL (TỪ DROPDOWN) ---
            has_al_col = al_col_sel != "-- Bỏ qua (Set = 0) --"
            if has_al_col:
                df[al_col_sel] = pd.to_numeric(df[al_col_sel], errors='coerce').fillna(0)

            def resolve_cgl_l(row):
                val = row[cgl_l]
                if pd.isna(val) or val == 0:
                    if row['has_main']:
                        return 0  # Có X00 gánh -> Trả về 0
                    else:
                        return row[al_col_sel] if has_al_col else 0 # Mồ côi -> Lấy cột AL
                return val

            df[cgl_l] = df.apply(resolve_cgl_l, axis=1)

            df[cgl_t] = df.groupby([order_c, 'base_coil'])[cgl_t].transform(lambda x: x.ffill().bfill()).fillna(0)
            df[cgl_w] = df.groupby([order_c, 'base_coil'])[cgl_w].transform(lambda x: x.ffill().bfill()).fillna(0)

            df[ccl_t] = df[ccl_t].fillna(0)
            df[ccl_w] = df[ccl_w].fillna(0)
            df[ccl_l] = df[ccl_l].fillna(0)

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

            # Hiển thị
            st.subheader("1. Order Summary")
            st.markdown("""
            > **💡 術語說明:** **Diff Area (m²)** = **塗層面積差異**
            > * **正值 (+):** 鋼帶延展 (Elongation)。  
            > * **負值 (-):** 長度短缺 (Shortage)，已扣除頭尾廢料 (Scrap Deducted)。
            """)

            disp = summary[[order_c, 'Qty (Coils)', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
            disp.columns = ['Order ID', 'Qty (Coils)', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
            
            disp = disp.sort_values(by='Cut Scrap (m)', ascending=False).reset_index(drop=True)
            disp['Qty (Coils)'] = disp['Qty (Coils)'].astype(int)
            disp.insert(0, 'No.', range(1, len(disp) + 1))
            
            st.table(disp.set_index('No.').style.format({
                "Input (m)": "{:,.0f}", "Cut Scrap (m)": "{:,.0f}", "Output (m)": "{:,.0f}",
                "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:.2f}"
            }))

            if not has_al_col:
                st.error("⚠️ Bạn chưa chọn cột AL ở thanh Sidebar bên trái! Các cuộn mồ côi (như 4BC134) đang bị tính đầu vào = 0, dẫn đến sai lệch Diff.")

        except Exception as e:
            st.error(f"Logic Error: {e}")
else:
    st.info("Please insert the Google Sheet Link.")
