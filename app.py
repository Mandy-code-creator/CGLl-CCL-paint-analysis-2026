import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Yield & Variance Analytics: Galvanized Steel Coils", layout="wide")

# ==========================================================
# 1. AUTO-SYNC CONFIGURATION
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

# --- DARK MODE DESIGN ---
st.markdown("""
    <style>
    .stApp { background-color: #0f172a; }
    div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart), 
    div[data-testid="stVerticalBlock"] > div:has(div.stDataFrame),
    div[data-testid="stVerticalBlock"] > div:has(div.stTable) {
        background-color: #1e293b; padding: 20px; border-radius: 8px;
        margin-bottom: 20px; border: none;
    }
    h1, h2, h3 { color: #f8fafc; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
    table { 
        width: 100% !important; 
        border-collapse: collapse !important; 
        font-family: 'Segoe UI', sans-serif;
        color: #e2e8f0;
        border: 1px solid #334155 !important;
    }
    th { 
        border: 1px solid #334155 !important; 
        color: #38bdf8 !important; 
        text-align: center !important; 
        padding: 12px 8px !important;
        font-size: 13px !important;
        background-color: #0f172a !important;
    }
    td { 
        text-align: center !important; 
        padding: 10px 8px !important; 
        border: 1px solid #334155 !important; 
        font-size: 13px !important;
    }
    tr:hover { background-color: #334155; }
    .stSelectbox label { color: #f8fafc !important; font-weight: bold; }
    .stMarkdown p { color: #cbd5e1 !important; }
    @media print {
        header, .stSidebar, .stButton, [data-testid="stHeader"], .stDivider, .stTextInput { display: none !important; }
        .main .block-container { max-width: 100% !important; padding: 0.5cm !important; }
        table { border: 1px solid #000 !important; color: #000 !important; }
        th, td { border: 0.5pt solid #ccc !important; }
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Yield & Variance Analytics: Galvanized Steel Coils")

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
        # Smart column filter to prevent mapping errors
        def get_col(default, possible_names):
            for name in possible_names:
                if name in df.columns: return name
            return default

        order_c = get_col("訂單號碼", ["訂單號碼", "订单号码"])
        mother_c = get_col("投入鋼捲號碼", ["投入鋼捲號碼", "投入钢卷号码"])
        baby_c = get_col("產出鋼捲號碼", ["產出鋼捲號碼", "产出钢卷号码"])
        cgl_l = get_col("镀锌測長度", ["镀锌測長度", "镀锌實測長度", "镀锌长度", "鍍鋅測長度"])
        ccl_l = get_col("實測長度", ["實測長度", "实测长度"])
        cgl_w = get_col("镀锌測寬度", ["镀锌測寬度", "镀锌測寬", "镀锌宽度", "鍍鋅測寬度", "镀锌实测宽度", "鍍鋅測寬"])
        cgl_t = get_col("镀锌實測厚度", ["镀锌實測厚度", "镀锌測厚", "镀锌厚度", "鍍鋅實測厚度"])
        ccl_w = get_col("實測寬度", ["實測寬度", "实测宽度"])
        ccl_t = get_col("實測厚度", ["實測厚度", "实测厚度"])
        outer_cut = get_col("outercutlength", ["outercutlength", "outercut"])
        inner_cut = get_col("innercutlength", ["innercutlength", "innercut"])

        try:
            for col in [outer_cut, inner_cut]:
                if col not in df.columns:
                    df[col] = 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # Convert all dimension columns to numeric
            for col in [cgl_t, cgl_w, cgl_l, ccl_t, ccl_w, ccl_l]:
                if col not in df.columns: df[col] = 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
            # --- AUTOMATIC ROUTING ALGORITHM ---
            
            # Step 1: Prevent double counting for OUTPUT coils FIRST
            df['is_first_baby'] = ~df.duplicated(subset=[order_c, baby_c], keep='first')
            df[ccl_l] = df.apply(lambda r: r[ccl_l] if r['is_first_baby'] else 0, axis=1)

            # Step 2: Identify base family ID (Remove X00, A00, etc.)
            def get_root(s):
                return re.sub(r'[A-Za-z]\d{2}$', '', str(s))
            df['base_coil'] = df[mother_c].apply(get_root)
            
            # Step 3: Check characteristics of the family group
            df['is_x00'] = df[mother_c].astype(str).str.endswith('X00', na=False)
            df['family_has_x00'] = df.groupby([order_c, 'base_coil'])['is_x00'].transform('any')
            df['family_has_cgl'] = df.groupby([order_c, 'base_coil'])[cgl_l].transform(lambda x: (x > 0).any())
            
            # Step 4: Apply logic to calculate precise input length
            df['is_first_mother'] = ~df.duplicated(subset=[order_c, mother_c])
            df['sum_ccl_by_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')

            def resolve_input(row):
                # Only keep the first occurrence of the mother coil
                if not row['is_first_mother']: return 0
                
                # If CGL length exists
                if row[cgl_l] > 0:
                    # Rule 1: If family has X00, only X00 keeps the value. Children get 0.
                    if row['family_has_x00'] and not row['is_x00']: return 0
                    # Rule 2: If no X00, siblings (A00, B00) keep their own values.
                    return row[cgl_l]
                
                # If CGL length is empty
                if row[cgl_l] == 0:
                    # Rule 3: If other siblings have CGL, this empty one gets 0.
                    if row['family_has_cgl']: return 0
                    # Rule 4: Total orphan. Use the sum of CCL.
                    return row['sum_ccl_by_mother']
                return 0
            
            # Apply the filtered data back to the columns
            df[cgl_l] = df.apply(resolve_input, axis=1)
            df[outer_cut] = df.apply(lambda r: r[outer_cut] if r['is_first_mother'] else 0, axis=1)
            df[inner_cut] = df.apply(lambda r: r[inner_cut] if r['is_first_mother'] else 0, axis=1)
            # ------------------------------------------------------------------

            # Copy Thickness & Width from root coils
            df[cgl_t] = df.groupby([order_c, 'base_coil'])[cgl_t].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)
            df[cgl_w] = df.groupby([order_c, 'base_coil'])[cgl_w].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            # Handle missing data for Output
            df[ccl_t] = df[ccl_t].fillna(0)
            df[ccl_w] = df[ccl_w].fillna(0)
            df[ccl_l] = df[ccl_l].fillna(0)

            # --- GROUP BY MOTHER COIL ---
            s1 = df.groupby([order_c, mother_c]).agg({
                cgl_t: 'mean', cgl_w: 'mean', cgl_l: 'first',
                ccl_t: 'mean', ccl_w: 'mean', ccl_l: 'sum',
                outer_cut: 'max', inner_cut: 'max' 
            }).reset_index()

            # AGGREGATE ENTIRE ORDER
            summary = s1.groupby(order_c).agg({
                mother_c: 'count', cgl_l: 'sum', ccl_l: 'sum', 
                outer_cut: 'sum', inner_cut: 'sum',
                ccl_t: 'mean', cgl_t: 'mean', cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty (Coils)', cgl_l: 'In_m', ccl_l: 'Out_m'})
            
            # CALCULATE VARIANCE
            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])
            summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # ==========================================================
            # UI: INTERACTIVE DATA GRIDS (WITH EXCEL-LIKE SCROLLBARS)
            # ==========================================================

            # --- 1. ORDER SUMMARY ---
            st.subheader("1. Order Summary")
            
            st.markdown("""
            > **💡 術語說明 (Technical Note):** > **Diff Area (m²)** = **塗層面積差異** (Coating Area Variance)  
            > * **正值 (+):** 鋼帶延展 (Elongation)，導致塗漆消耗量增加。  
            > * **負值 (-):** 長度短缺 (Shortage)，已扣除頭尾廢料 (Scrap Deducted)，可能源於感測器誤差。
            """)

            disp = summary[[order_c, 'Qty (Coils)', 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()
            disp.columns = ['Order ID', 'Qty (Coils)', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']
            
            disp = disp.sort_values(by='Cut Scrap (m)', ascending=False).reset_index(drop=True)
            disp['Qty (Coils)'] = disp['Qty (Coils)'].astype(int)
            disp.insert(0, 'No.', range(1, len(disp) + 1))
            
            # Display only the top 20 items using dataframe
            st.dataframe(
                disp.head(20).set_index('No.').style.format({
                    "Input (m)": "{:,.0f}", "Cut Scrap (m)": "{:,.0f}", "Output (m)": "{:,.0f}",
                    "Diff (m)": "{:,.0f}", "Thick Var": "{:.3f}", "Diff Area (m²)": "{:,.0f}"
                }),
                height=600, 
                use_container_width=True
            )

           # --- 2. DATA INSPECTION ---
            st.divider()
            st.subheader("2. Production Coil Details (Data Inspection)") 
            
            order_list = df[order_c].unique()
            
            sel_order = st.selectbox(
                "🔍 Type or Select Order ID to view details:", 
                options=order_list,
                index=None,
                placeholder="Ex: P240801... (Click and type here to search)"
            )
            
            if sel_order:
                det = df[df[order_c] == sel_order].copy()
                det['Var'] = det[ccl_t] - det[cgl_t]
                
                det_f = det[[
                    mother_c, baby_c, 
                    cgl_t, cgl_w, cgl_l, 
                    outer_cut, inner_cut,
                    ccl_t, ccl_w, 'Var', ccl_l 
                ]].copy()
                
                det_f.columns = [
                    'Input Coil ID (CGL)', 
                    'Output Coil ID (CCL)', 
                    'Input Thick (mm)', 
                    'Input Width (mm)', 
                    'Input Length (m)', 
                    'Outer Cut (m)',
                    'Inner Cut (m)',
                    'Output Thick (mm)', 
                    'Output Width (mm)', 
                    'Thick Deviation (mm)', 
                    'Output Length (m)'
                ]
                
                # Interactive data grid for inspection details (top 20, integers only)
                st.dataframe(
                    det_f.head(20).style.format({
                        "Input Thick (mm)": "{:.0f}", 
                        "Input Width (mm)": "{:.0f}", 
                        "Input Length (m)": "{:.0f}",
                        "Outer Cut (m)": "{:.0f}",
                        "Inner Cut (m)": "{:.0f}",
                        "Output Thick (mm)": "{:.0f}", 
                        "Output Width (mm)": "{:.0f}",
                        "Thick Deviation (mm)": "{:.0f}", 
                        "Output Length (m)": "{:.0f}"
                    }),
                    height=600,
                    use_container_width=True
                )
                
            # --- 3. VISUAL INSIGHTS & ANALYSIS ---
            st.divider()
            st.subheader("3. Visual Insights & Analysis")
            
            # Use darker color scales
            f1 = px.bar(disp.head(20), x='Order ID', y='Diff Area (m²)', color='Diff (m)', 
                        color_continuous_scale='Inferno', title="Extra Area per Order", template="plotly_dark")
            st.plotly_chart(f1, use_container_width=True)
            st.info("**分析結論:** 監控各訂單的塗層面積偏差。偏離中心值的數據代表生產投入與產出不一致，建議優先核對該批次的生產日誌。")

            f2 = px.histogram(disp, x='Diff (m)', nbins=15, title="Production Variance Distribution", template="plotly_dark")
            f2.update_traces(marker_color='#38bdf8')
            st.plotly_chart(f2, use_container_width=True)
            st.warning("**分析結論:** 數據分布反映生產穩定性。離群值標示該訂單存在異常長度變化，需確認是物理延展、裁切損耗或是計量誤差。")

            disp_chart = disp.sort_values(by='Cut Scrap (m)', ascending=False).head(20)
            f3 = px.bar(disp_chart, x='Order ID', y='Cut Scrap (m)', 
                        title="Total Cut Scrap per Order (Outer + Inner)",
                        color='Cut Scrap (m)', color_continuous_scale='Magma', template="plotly_dark")
            f3.update_layout(yaxis_title="Scrap Length (m)")
            st.plotly_chart(f3, use_container_width=True)
            st.error("**分析結論:** 各訂單的剪切廢料總量。監控此數據有助於評估來料質量與生產初期的裁切損耗。若數值異常偏高，需檢查鋼捲頭尾品質狀況。")

            # --- 4. EXECUTIVE SUMMARY ---
            st.divider()
            st.subheader("4. Executive Summary")
            t_in, t_out = disp['Input (m)'].sum(), disp['Output (m)'].sum()
            area_s = abs(disp[disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())
            st.markdown(f"""
            **生產產出綜合分析:**
            * **總投入 (Total Input):** {t_in:,.0f} m
            * **總產出 (Total Output):** {t_out:,.0f} m
            * **不明面積差異 (Area Shortfall):** {area_s:,.2f} m² (需進一步核實廢料申報準確性)
            """)

            # --- 5. EXPORT ---
            st.subheader("5. Export Data")
            c1, c2 = st.columns(2)
            with c1:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    disp.to_excel(writer, sheet_name='Summary', index=False)
                st.download_button("📊 Download Excel", data=buf.getvalue(), file_name="Report.xlsx", type="primary", use_container_width=True)
            with c2:
                components.html("""
                    <script>function printPage() { window.parent.print(); }</script>
                    <button onclick="printPage()" style="background-color: transparent; color: #38bdf8; border: 1.5px solid #38bdf8; 
                    border-radius: 4px; padding: 10px; font-size: 14px; cursor: pointer; width: 100%; font-weight: 600;"> 
                    Save as PDF Report </button>
                """, height=70)

        except Exception as e:
            st.error(f"Logic Error: {e}")
