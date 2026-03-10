import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Steel Coil Production Variance Dashboard", layout="wide")

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

st.title("Steel Coil Production Variance Dashboard")

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
        
        # Smart column finder to prevent Logic Errors
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
        cgl_t    = find_col(["镀锌", "厚度"], "镀锌實測厚度")
        ccl_w    = find_col(["实测", "宽度"], "實測寬度")
        ccl_t    = find_col(["实测", "厚度"], "實測厚度")
        outer_cut = find_col(["outer", "cut"], "outercutlength")
        inner_cut = find_col(["inner", "cut"], "innercutlength")

        try:
            # 1. Convert to numeric
            for c in [cgl_l, ccl_l, outer_cut, inner_cut, cgl_w, cgl_t, ccl_w, ccl_t]:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

            # 2. IDENTIFY ROOT FAMILY
            def get_root(s):
                return re.sub(r'[A-Za-z]\d{2}$', '', str(s))
            df['root_id'] = df[mother_c].apply(get_root)
            
            # Check if order has the main mother coil
            df['is_main_id'] = df[mother_c].astype(str).str.contains(r'X00|A00', case=False, regex=True)
            df['group_has_main_cgl'] = df.groupby([order_c, 'root_id'])[cgl_l].transform(lambda x: (x > 0).any())

            # 3. PREVENT DOUBLE COUNTING ON INPUT & SCRAP
            df['is_first_mother'] = ~df.duplicated(subset=[order_c, mother_c])
            df['sum_ccl_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')

            def resolve_input_logic(row):
                if not row['is_first_mother']: return 0
                if row[cgl_l] > 0:
                    # If it's a baby coil but group has a mother, set input to 0
                    if not row['is_main_id'] and row['group_has_main_cgl']: return 0
                    return row[cgl_l]
                # If orphan with no CGL data, use CCL sum
                if row[cgl_l] == 0 and not row['group_has_main_cgl']:
                    return row['sum_ccl_mother']
                return 0

            df['final_input'] = df.apply(resolve_input_logic, axis=1)
            
            # Prevent Scrap double count
            df['final_outer_cut'] = df.apply(lambda r: r[outer_cut] if r['is_first_mother'] else 0, axis=1)
            df['final_inner_cut'] = df.apply(lambda r: r[inner_cut] if r['is_first_mother'] else 0, axis=1)

            # 4. PREVENT DOUBLE COUNTING ON OUTPUT
            # Keep the FIRST occurrence of the baby coil
            df['is_first_baby'] = ~df.duplicated(subset=[order_c, baby_c], keep='first')
            df['final_output'] = df.apply(lambda r: r[ccl_l] if r['is_first_baby'] else 0, axis=1)

            # Copy Width/Thick for orphans
            df[cgl_w] = df.apply(lambda r: r[ccl_w] if r[cgl_w] == 0 else r[cgl_w], axis=1)
            df[cgl_t] = df.apply(lambda r: r[ccl_t] if r[cgl_t] == 0 else r[cgl_t], axis=1)

            # --- 5. AGGREGATE SUMMARY ---
            summary = df.groupby(order_c).agg({
                mother_c: 'count', 
                'final_input': 'sum', 
                'final_output': 'sum', 
                'final_outer_cut': 'sum', 
                'final_inner_cut': 'sum', 
                cgl_w: 'mean', cgl_t: 'mean', ccl_t: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty', 'final_input': 'In_m', 'final_output': 'Out_m'})
            summary['Total_Cut'] = summary['final_outer_cut'] + summary['final_inner_cut']
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
            summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # --- UI: ORDER SUMMARY ---
            st.subheader("1. Order Summary")
            disp = summary[[order_c, 'Qty', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
            disp.columns = ['Order ID', 'Coils', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
            
            # Sort and display only 20 items for clarity
            disp = disp.sort_values(by='Cut Scrap (m)', ascending=False).reset_index(drop=True)
            disp_top = disp.head(20)
            
            st.table(disp_top.style.format({
                "Input (m)": "{:,.0f}", "Cut Scrap (m)": "{:,.0f}", "Output (m)": "{:,.0f}", 
                "Diff (m)": "{:.2f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:.2f}"
            }))

            # --- UI: COIL DETAILS ---
            st.divider()
            st.subheader("2. Production Coil Details")
            sel_order = st.selectbox("🔍 Search Order ID:", options=df[order_c].unique(), index=None)
            if sel_order:
                check = df[df[order_c] == sel_order][[mother_c, baby_c, cgl_l, 'final_input', ccl_l, 'final_output']]
                check.columns = ['Input ID', 'Output ID', 'CGL Orig', 'Final Input', 'CCL Orig', 'Final Output']
                st.table(check.head(20).style.format({
                    "CGL Orig": "{:,.0f}", "Final Input": "{:,.0f}", "CCL Orig": "{:,.0f}", "Final Output": "{:,.0f}"
                }))

            # --- UI: VISUAL INSIGHTS ---
            st.divider()
            f1 = px.bar(disp_top, x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                        color_continuous_scale='Blues_r', title="Extra Area per Order")
            st.plotly_chart(f1, use_container_width=True)

        except Exception as e:
            st.error(f"Logic Error: {e}")
