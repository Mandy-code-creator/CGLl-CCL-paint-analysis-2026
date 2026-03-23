import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis", layout="wide")

# ==========================================================
# UI IMPROVEMENT (CLEAN WHITE THEME)
# ==========================================================
st.markdown("""
<style>
html, body, [class*="css"]  {
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
}
.block-container{
    padding-top:2rem;
    padding-bottom:2rem;
}
/* Card style ONLY for Charts */
div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart) {
    border-radius:12px;
    box-shadow:0 6px 25px rgba(0,0,0,0.05);
    padding:22px;
}
/* Remove default background/padding from DataFrame containers */
div[data-testid="stVerticalBlock"] > div:has(div.stDataFrame) {
    background-color: transparent !important;
    padding: 0 !important;
    border: none !important;
    box-shadow: none !important;
}
/* Titles & Dividers */
h1 { font-size:36px; letter-spacing:0.3px; }
h2, h3 { margin-top:10px; margin-bottom:10px; }
hr { border:1px solid rgba(0,0,0,0.05); }
/* Table elements */
thead tr th { font-weight:600 !important; }
tbody tr:hover { background-color:rgba(0,0,0,0.03); }
div[data-baseweb="select"], button[kind="primary"] { border-radius:8px; }
button[kind="primary"] { font-weight:600; }
div[data-testid="stAlert"], .js-plotly-plot { border-radius:10px; }
/* Scrollbar */
::-webkit-scrollbar { width:8px; }
::-webkit-scrollbar-thumb { background:#94a3b8; border-radius:10px; }
</style>
""", unsafe_allow_html=True)

grad_colors = "#ffffff, #f8fafc, #f1f5f9, #ffffff"
card_bg = "rgba(255, 255, 255, 0.95)"
text_color = "#0f172a" 
sub_text = "#334155"
plotly_template = "plotly_white"
border_color = "#e2e8f0"

st.markdown(f"""
<style>
.stApp {{
    background: linear-gradient(-45deg, {grad_colors});
    background-size: 400% 400%;
    animation: gradient_move 20s ease infinite;
}}
@keyframes gradient_move {{
    0% {{ background-position: 0% 50%; }}
    50% {{ background-position: 100% 50%; }}
    100% {{ background-position: 0% 50%; }}
}}
/* Apply theme colors ONLY to Charts */
div[data-testid="stVerticalBlock"] > div:has(div.stPlotlyChart) {{
    background-color: {card_bg};
    padding: 20px;
    border-radius: 8px;
    margin-bottom: 20px;
    border: 1px solid {border_color};
}}
h1, h2, h3 {{ color: {text_color}; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }}
.stMarkdown p {{ color: {sub_text} !important; }}
.stSelectbox label, .stRadio label {{ color: {text_color} !important; font-weight: bold; }}
</style>
""", unsafe_allow_html=True)

st.title("Length Variance Analysis: Total CGL vs CCL per Order")

# ==========================================================
# 2. DATA PROCESSING
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"
# LINK DỮ LIỆU SỐ MÉT TRẢ VỀ
RETURN_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1w1ob2a4lLa6vo-eJZVt4wtPUqQ1gECjNshkOtsrk51Y/edit?gid=0#gid=0"

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
            # Chuẩn hóa tên cột: bỏ khoảng trắng, in thường
            df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(r'\s+', '', regex=True)
            return df
        return None
    except Exception as e:
        st.error(f"Connection Error: {e}")
        return None

def get_col(default, possible_names, dataframe):
    for name in possible_names:
        if name in dataframe.columns:
            return name
    return default

if GSHEET_URL:
    df = load_auto_data(GSHEET_URL)
    df_return = load_auto_data(RETURN_GSHEET_URL) # Load dữ liệu trả về

    if df is not None:
        # --- Cột cho df chính ---
        order_c = get_col("訂單號碼", ["訂單號碼", "订单号码"], df)
        mother_c = get_col("投入鋼捲號碼", ["投入鋼捲號碼", "投入钢卷号码"], df)
        baby_c = get_col("產出鋼捲號碼", ["產出鋼捲號碼", "产出钢卷号码"], df)
        cgl_l = get_col("镀锌測長度", ["镀锌測長度", "镀锌實測長度", "镀锌长度", "鍍鋅測長度"], df)
        ccl_l = get_col("實測長度", ["實測長度", "实测长度"], df)
        cgl_w = get_col("镀锌測寬度", ["镀锌測寬度", "镀锌測寬", "镀锌宽度", "鍍鋅測寬度", "镀锌实测宽度"], df)
        cgl_t = get_col("镀锌實測厚度", ["镀锌實測厚度", "镀锌測厚", "镀锌厚度", "鍍鋅實測厚度"], df)
        ccl_w = get_col("實測寬度", ["實測寬度", "实测宽度"], df)
        ccl_t = get_col("實測厚度", ["實測厚度", "实测厚度"], df)
        outer_cut = get_col("outercutlength", ["outercutlength", "outercut"], df)
        inner_cut = get_col("innercutlength", ["innercutlength", "innercut"], df)
        line_c = get_col("線別", ["線別", "线别"], df)
        out_grade_c = get_col("產出等級", ["產出等級", "产出等级"], df)
        next_proc_c = get_col("下製程", ["下製程", "下制程"], df)
        theo_paint_c = get_col("合計理論耗用", ["合計理論耗用", "合计理论耗用", "理論耗用", "理论耗用"], df)
        act_paint_c = get_col("合計實際耗用", ["合計實際耗用", "合计实际耗用", "實際耗用", "实际耗用"], df)

        # --- Cột cho df_return (File trả về) ---
        return_data_dict = {}
        if df_return is not None:
            # Code tìm cột đã được tự động làm sạch (viết thường, không khoảng trắng)
            ret_order_c = get_col("ordernumber", ["ordernumber"], df_return)
            
            # Khai báo chính xác tên cột chứa số mét trả về sau khi làm sạch
            ret_len_c = get_col("qualityclass7mlength", ["qualityclass7mlength"], df_return)
            
            if ret_len_c in df_return.columns and ret_order_c in df_return.columns:
                df_return[ret_len_c] = pd.to_numeric(df_return[ret_len_c], errors='coerce').fillna(0)
                # Gom nhóm số mét trả về theo từng đơn hàng (Order ID)
                return_grouped = df_return.groupby(ret_order_c)[ret_len_c].sum().to_dict()
                return_data_dict = return_grouped

        try:
            # Data cleaning
            for col in [line_c, out_grade_c, next_proc_c]:
                if col not in df.columns:
                    df[col] = "-"
                df[col] = df[col].fillna("-").astype(str)

            for col in [outer_cut, inner_cut, cgl_t, cgl_w, cgl_l, ccl_t, ccl_w, ccl_l, theo_paint_c, act_paint_c]:
                if col not in df.columns:
                    df[col] = 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # Deduplicate for first baby coil
            df['is_first_baby'] = ~df.duplicated(subset=[order_c, baby_c], keep='first')
            df[ccl_l] = df.apply(lambda r: r[ccl_l] if r['is_first_baby'] else 0, axis=1)

            # Family mapping
            df['base_coil'] = df[mother_c].astype(str).str[:-3]
            df['is_x00'] = df[mother_c].astype(str).str.endswith('X00', na=False)
            df['family_has_x00'] = df.groupby([order_c, 'base_coil'])['is_x00'].transform('any')
            df['family_has_cgl'] = df.groupby([order_c, 'base_coil'])[cgl_l].transform(lambda x: (x > 0).any())
            df['is_first_mother'] = ~df.duplicated(subset=[order_c, mother_c])
            df['sum_ccl_by_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')

            # Input resolution logic
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

            # Forward/Backward fill for parameters
            df[cgl_t] = df.groupby([order_c, 'base_coil'])[cgl_t].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)
            df[cgl_w] = df.groupby([order_c, 'base_coil'])[cgl_w].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            # Level 1 Aggregation: Mother Coil
            s1 = df.groupby([order_c, mother_c]).agg({
                line_c: 'first',
                cgl_t: 'mean', cgl_w: 'mean', cgl_l: 'first',
                ccl_t: 'mean', ccl_w: 'mean', ccl_l: 'sum',
                outer_cut: 'max', inner_cut: 'max',
                theo_paint_c: 'max', act_paint_c: 'max'   
            }).reset_index()

            # Level 2 Aggregation: Order Summary
            summary = s1.groupby(order_c).agg({
                line_c: 'first',
                mother_c: 'count', cgl_l: 'sum', ccl_l: 'sum',
                outer_cut: 'sum', inner_cut: 'sum',
                ccl_t: 'mean', cgl_t: 'mean', cgl_w: 'mean',
                theo_paint_c: 'max', act_paint_c: 'max'   
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty (Coils)', cgl_l: 'In_m', ccl_l: 'Out_m'})
            
            # --- MAP SỐ MÉT TRẢ VỀ VÀO BẢNG TỔNG HỢP ---
            # Gắn số mét trả về từ file 2 vào từng đơn hàng. Nếu không có thì gán = 0
            summary['Return_m'] = summary[order_c].map(return_data_dict).fillna(0)
            
            # --- CẬP NHẬT CÔNG THỨC DIFF ---
            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]
            # Diff mới = (Đầu ra + Trả về) - (Đầu vào - Cắt phế)
            summary['Diff'] = (summary['Out_m'] + summary['Return_m']) - (summary['In_m'] - summary['Total_Cut'])
            
            summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]
            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # Variance Breakdown Calculation (Lưu ý: paint tính theo Out_m, bỏ qua hàng trả về)
            def calc_variance_breakdown(row):
                act_paint = row[act_paint_c]
                if act_paint <= 0: return pd.Series([0, 0, 0, 0])
                
                theo_paint = row[theo_paint_c]
                out_m = row['Out_m'] # Dùng Out_m thật (không gồm trả về) để tính Yield
                cut_scrap = row['Total_Cut']
                diff = row['Diff'] # Diff đã bù trừ hàng trả về
                
                yield_pct = (theo_paint / act_paint) * 100
                dinh_muc = (theo_paint / out_m) if out_m > 0 else 0
                scrap_loss = (cut_scrap * dinh_muc) / act_paint * 100
                len_loss = (abs(diff) * dinh_muc) / act_paint * 100 if diff < 0 else 0
                
                other_loss = 100 - yield_pct - scrap_loss - len_loss
                if other_loss < 0: other_loss = 0
                
                return pd.Series([yield_pct, scrap_loss, len_loss, other_loss])

            summary[['Yield (%)', 'Scrap Loss (%)', 'Len Var Loss (%)', 'Other Causes (%)']] = summary.apply(calc_variance_breakdown, axis=1)

            # --- UI: ORDER SUMMARY ---
            col1, col2 = st.columns([8, 2])
            with col1:
                st.subheader("1. Order Summary & Variance Breakdown")
            with col2:
                row_limit_summary = st.selectbox("Show rows:", options=[20, 50, 100, "All"], index=0, key="summary_rows")

            # Thêm 'Return_m' vào bảng hiển thị
            disp = summary[[order_c, line_c, 'Qty (Coils)', cgl_w, 'In_m', 'Total_Cut', 'Return_m', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2', theo_paint_c, act_paint_c, 'Yield (%)', 'Scrap Loss (%)', 'Len Var Loss (%)', 'Other Causes (%)']].copy()
            disp.columns = ['Order ID', 'Line', 'Qty (Coils)', 'Input Width', 'Input (m)', 'Cut Scrap (m)', 'Return (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)', 'Theo Paint (kg)', 'Real Paint (kg)', 'Yield (%)', 'Scrap Loss (%)', 'Len Var Loss (%)', 'Other Causes (%)']
            
            disp = disp.sort_values(by='Cut Scrap (m)', ascending=False).reset_index(drop=True)
            disp.insert(0, 'No.', range(1, len(disp) + 1))

            if row_limit_summary == "All":
                disp_view = disp
            else:
                disp_view = disp.head(row_limit_summary)

            st.dataframe(
                disp_view.set_index('No.').style.format({
                    "Input Width": "{:,.0f}", "Input (m)": "{:,.0f}", "Cut Scrap (m)": "{:,.0f}", 
                    "Return (m)": "{:,.0f}", "Output (m)": "{:,.0f}", "Diff (m)": "{:,.0f}", 
                    "Thick Var": "{:.3f}", "Diff Area (m²)": "{:,.0f}",
                    "Theo Paint (kg)": "{:,.2f}", "Real Paint (kg)": "{:,.2f}",
                    "Yield (%)": "{:.2f}%", "Scrap Loss (%)": "{:.2f}%",
                    "Len Var Loss (%)": "{:.2f}%", "Other Causes (%)": "{:.2f}%"
                }), 
                use_container_width=True
            )

            st.divider()

            # --- UI: PRODUCTION COIL DETAILS ---
            col3, col4 = st.columns([8, 2])
            with col3:
                st.subheader("2. Production Coil Details")
            with col4:
                row_limit_details = st.selectbox("Show rows:", options=[20, 50, 100, "All"], index=0, key="details_rows")
            
            sel_order = st.selectbox("🔍 Select Order ID:", options=df[order_c].unique(), index=None)

            if sel_order:
                det = df[df[order_c] == sel_order].copy()
                det['Var'] = det[ccl_t] - det[cgl_t]
                det_f = det[[mother_c, baby_c, line_c, out_grade_c, next_proc_c, cgl_t, cgl_w, cgl_l, outer_cut, inner_cut, ccl_t, ccl_w, 'Var', ccl_l]].copy()
                det_f.columns = ['Input ID', 'Output ID', 'Line', 'Grade', 'Next Proc', 'In Thick', 'In Width', 'In Len', 'Outer Cut', 'Inner Cut', 'Out Thick', 'Out Width', 'Thick Dev', 'Out Len']

                if row_limit_details == "All":
                    det_view = det_f
                else:
                    det_view = det_f.head(row_limit_details)

                st.dataframe(
                    det_view.style.format({
                        "In Thick": "{:.3f}", "In Width": "{:,.0f}", "In Len": "{:,.0f}",
                        "Outer Cut": "{:,.0f}", "Inner Cut": "{:,.0f}", "Out Thick": "{:.3f}", 
                        "Out Width": "{:,.0f}", "Thick Dev": "{:.3f}", "Out Len": "{:,.0f}"
                    }), 
                    use_container_width=True
                )

            st.divider()

            # --- UI: VISUAL INSIGHTS ---
            st.subheader("3. Visual Insights & Analysis")

            fig_breakdown = px.bar(
                disp_view, 
                x='Order ID', 
                y=['Yield (%)', 'Scrap Loss (%)', 'Len Var Loss (%)', 'Other Causes (%)'],
                title="Paint Consumption Variance Breakdown / 塗料耗用差異分析",
                template=plotly_template,
                labels={'value': 'Percentage (%)', 'variable': 'Category'},
                color_discrete_map={
                    'Yield (%)': '#059669',        
                    'Scrap Loss (%)': '#dc2626',   
                    'Len Var Loss (%)': '#d97706', 
                    'Other Causes (%)': '#475569'  
                }
            )
            fig_breakdown.update_layout(barmode='stack')
            st.plotly_chart(fig_breakdown, use_container_width=True)
            st.info("**分析結論:** 塗料耗用差異分析 (Variance Breakdown)。顯示各訂單中，實際塗料耗用的結構比例。")

            st.plotly_chart(px.bar(disp_view, x='Order ID', y='Diff Area (m²)', color='Diff (m)', color_continuous_scale='Tealgrn', title="Extra Area per Order", template=plotly_template), use_container_width=True)
            st.info("**分析結論:** 監控各訂單的塗層面積偏差。偏離中心值的數據代表生產投入與產出不一致，建議優先核對該批次的生產日誌。")

            st.plotly_chart(px.bar(disp_view.sort_values(by='Cut Scrap (m)', ascending=False), x='Order ID', y='Cut Scrap (m)', 
                           title="Total Cut Scrap per Order (Outer + Inner)", color='Cut Scrap (m)', color_continuous_scale='Reds', template=plotly_template), use_container_width=True)

            st.divider()

            # --- UI: EXECUTIVE SUMMARY ---
            st.subheader("4. Executive Summary")

            # CALCULATE METRICS FOR EXECUTIVE SUMMARY
            t_in = disp['Input (m)'].sum()
            t_out = disp['Output (m)'].sum()
            t_cut = disp['Cut Scrap (m)'].sum()
            t_return = disp['Return (m)'].sum()
            t_diff = disp['Diff (m)'].sum()
            avg_thick_var = disp['Thick Var'].mean()
            area_s = abs(disp[disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())
            avg_yield = (disp['Theo Paint (kg)'].sum() / disp['Real Paint (kg)'].sum() * 100) if disp['Real Paint (kg)'].sum() > 0 else 0

            # UPDATED MARKDOWN WITH RETURN METERS
            st.markdown(f"""
            **生產產出綜合分析 (Comprehensive Output Analysis):**
            * **總投入長度 (Total Input):** {t_in:,.0f} m  
            * **總產出長度 (Total Output):** {t_out:,.0f} m  
            * **總退回長度 (Total Returned):** <span style="color:blue; font-weight:bold;">{t_return:,.0f} m</span>
            * **總切廢長度 (Total Cut Scrap):** {t_cut:,.0f} m
            * **總長度差異 (Total Length Diff):** <span style="color:red; font-weight:bold;">{t_diff:,.0f} m</span> *(Đã bù trừ hàng trả về)*
            * **平均厚度差異 (Avg Thickness Var):** {avg_thick_var:.3f} mm
            * **不明面積差異 (Area Shortfall):** {area_s:,.2f} m² 
            * **平均績效 (Avg Yield):** {avg_yield:.2f}%
            """, unsafe_allow_html=True)

            st.subheader("5. Export Data")

            buf = io.BytesIO()

            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                disp.to_excel(writer, sheet_name='Summary', index=False)

            st.download_button(
                "📊 Download Excel Report",
                data=buf.getvalue(),
                file_name="Report.xlsx",
                type="primary"
            )

        except Exception as e:
            st.error(f"Logic Error: {e}")
else:
    st.info("Please insert the Google Sheet Link.")
