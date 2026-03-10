```python
import streamlit as st
import pandas as pd
import plotly.express as px
import io
import streamlit.components.v1 as components
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Length Variance Analysis: Total CGL vs CCL per Order", layout="wide")

# ==========================================================
# UI STYLE (VISUAL ONLY)
# ==========================================================
st.markdown("""
<style>

html, body, [class*="css"]  {
    font-family: 'Segoe UI', system-ui, sans-serif;
}

.block-container{
    padding-top:2rem;
}

thead tr th{
    font-weight:600 !important;
}

tbody tr:hover{
    background-color:rgba(120,120,120,0.08);
}

div[data-testid="stAlert"]{
    border-radius:10px;
}

.js-plotly-plot{
    border-radius:10px;
}

</style>
""", unsafe_allow_html=True)

# ==========================================================
# THEME SELECTION
# ==========================================================
theme_choice = st.radio("🎨 Select App Theme:", ["Light Mode (Standard)", "Dark Mode (Professional)"], horizontal=True)

if theme_choice == "Dark Mode (Professional)":
    plotly_template = "plotly_dark"
else:
    plotly_template = "plotly_white"

st.title("Length Variance Analysis: Total CGL vs CCL per Order")

# ==========================================================
# DATA SOURCE
# ==========================================================
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1-kayrLVYwOO66Xxc7Vk7dbTNZ5Aph4MVd9DMTz6RJS0/edit?gid=0#gid=0"

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


if GSHEET_URL:

    df = load_auto_data(GSHEET_URL)

    if df is not None:

        def get_col(default, possible_names):
            for name in possible_names:
                if name in df.columns:
                    return name
            return default

        order_c = get_col("訂單號碼", ["訂單號碼", "订单号码"])
        mother_c = get_col("投入鋼捲號碼", ["投入鋼捲號碼", "投入钢卷号码"])
        baby_c = get_col("產出鋼捲號碼", ["產出鋼捲號碼", "产出钢卷号码"])

        cgl_l = get_col("镀锌測長度", ["镀锌測長度", "镀锌實測長度", "镀锌长度", "鍍鋅測長度"])
        ccl_l = get_col("實測長度", ["實測長度", "实测长度"])

        cgl_w = get_col("镀锌測寬度", ["镀锌測寬度", "镀锌測寬", "镀锌宽度", "鍍鋅測寬度"])
        cgl_t = get_col("镀锌實測厚度", ["镀锌實測厚度", "镀锌測厚", "镀锌厚度"])

        ccl_w = get_col("實測寬度", ["實測寬度", "实测宽度"])
        ccl_t = get_col("實測厚度", ["實測厚度", "实测厚度"])

        outer_cut = get_col("outercutlength", ["outercutlength", "outercut"])
        inner_cut = get_col("innercutlength", ["innercutlength", "innercut"])

        line_c = get_col("線別", ["線別", "线别"])
        out_grade_c = get_col("產出等級", ["產出等級", "产出等级"])
        next_proc_c = get_col("下製程", ["下製程", "下制程"])

        try:

            for col in [line_c, out_grade_c, next_proc_c]:

                if col not in df.columns:
                    df[col] = "-"

                df[col] = df[col].fillna("-").astype(str)

            for col in [outer_cut, inner_cut, cgl_t, cgl_w, cgl_l, ccl_t, ccl_w, ccl_l]:

                if col not in df.columns:
                    df[col] = 0

                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # ==========================================================
            # ROUTING LOGIC (UNCHANGED)
            # ==========================================================

            df['is_first_baby'] = ~df.duplicated(subset=[order_c, baby_c], keep='first')

            df[ccl_l] = df.apply(lambda r: r[ccl_l] if r['is_first_baby'] else 0, axis=1)

            df['base_coil'] = df[mother_c].astype(str).str[:-3]

            df['is_x00'] = df[mother_c].astype(str).str.endswith('X00', na=False)

            df['family_has_x00'] = df.groupby([order_c, 'base_coil'])['is_x00'].transform('any')

            df['family_has_cgl'] = df.groupby([order_c, 'base_coil'])[cgl_l].transform(lambda x: (x > 0).any())

            df['is_first_mother'] = ~df.duplicated(subset=[order_c, mother_c])

            df['sum_ccl_by_mother'] = df.groupby([order_c, mother_c])[ccl_l].transform('sum')

            def resolve_input(row):

                if not row['is_first_mother']:
                    return 0

                if row[cgl_l] > 0:

                    if row['family_has_x00'] and not row['is_x00']:
                        return 0

                    return row[cgl_l]

                if row[cgl_l] == 0:

                    if row['family_has_cgl']:
                        return 0

                    return row['sum_ccl_by_mother']

                return 0

            df[cgl_l] = df.apply(resolve_input, axis=1)

            df[outer_cut] = df.apply(lambda r: r[outer_cut] if r['is_first_mother'] else 0, axis=1)

            df[inner_cut] = df.apply(lambda r: r[inner_cut] if r['is_first_mother'] else 0, axis=1)

            df[cgl_t] = df.groupby([order_c, 'base_coil'])[cgl_t].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            df[cgl_w] = df.groupby([order_c, 'base_coil'])[cgl_w].transform(lambda x: x.replace(0, pd.NA).ffill().bfill()).fillna(0)

            s1 = df.groupby([order_c, mother_c]).agg({
                cgl_t: 'mean',
                cgl_w: 'mean',
                cgl_l: 'first',
                ccl_t: 'mean',
                ccl_w: 'mean',
                ccl_l: 'sum',
                outer_cut: 'max',
                inner_cut: 'max'
            }).reset_index()

            summary = s1.groupby(order_c).agg({
                mother_c: 'count',
                cgl_l: 'sum',
                ccl_l: 'sum',
                outer_cut: 'sum',
                inner_cut: 'sum',
                ccl_t: 'mean',
                cgl_t: 'mean',
                cgl_w: 'mean'
            }).reset_index()

            summary = summary.rename(columns={mother_c: 'Qty (Coils)', cgl_l: 'In_m', ccl_l: 'Out_m'})

            summary['Total_Cut'] = summary[outer_cut] + summary[inner_cut]

            summary['Diff'] = summary['Out_m'] - (summary['In_m'] - summary['Total_Cut'])

            summary['Thick_Var'] = summary[ccl_t] - summary[cgl_t]

            summary['Area_m2'] = (summary[cgl_w] / 1000) * summary['Diff']

            # ==========================================================
            # DISPLAY TABLE
            # ==========================================================

            st.subheader("1. Order Summary")

            disp = summary[[order_c, 'Qty (Coils)', cgl_w, 'In_m', 'Total_Cut', 'Out_m', 'Diff', 'Thick_Var', 'Area_m2']].copy()

            disp.columns = ['Order ID', 'Qty (Coils)', 'Input Width (mm)', 'Input (m)', 'Cut Scrap (m)', 'Output (m)', 'Diff (m)', 'Thick Var', 'Diff Area (m²)']

            disp = disp.sort_values(by='Cut Scrap (m)', ascending=False).reset_index(drop=True)

            disp.insert(0, 'No.', range(1, len(disp) + 1))

            st.dataframe(disp.set_index('No.'), height=500, use_container_width=True)

            # ==========================================================
            # KPI DASHBOARD
            # ==========================================================

            st.divider()

            k1,k2,k3,k4,k5 = st.columns(5)

            k1.metric("Orders", f"{disp.shape[0]}")
            k2.metric("Total Input", f"{disp['Input (m)'].sum():,.0f} m")
            k3.metric("Total Output", f"{disp['Output (m)'].sum():,.0f} m")
            k4.metric("Total Scrap", f"{disp['Cut Scrap (m)'].sum():,.0f} m")
            k5.metric("Area Diff", f"{disp['Diff Area (m²)'].sum():,.0f} m²")

            # ==========================================================
            # COIL DETAILS
            # ==========================================================

            st.divider()

            st.subheader("2. Production Coil Details")

            sel_order = st.selectbox("🔍 Select Order ID:", options=df[order_c].unique(), index=None)

            if sel_order:

                det = df[df[order_c] == sel_order].copy()

                det['Var'] = det[ccl_t] - det[cgl_t]

                det_f = det[[mother_c, baby_c, line_c, out_grade_c, next_proc_c, cgl_t, cgl_w, cgl_l, outer_cut, inner_cut, ccl_t, ccl_w, 'Var', ccl_l]].copy()

                det_f.columns = ['Input ID', 'Output ID', 'Line', 'Grade', 'Next Proc', 'In Thick', 'In Width', 'In Len', 'Outer Cut', 'Inner Cut', 'Out Thick', 'Out Width', 'Thick Dev', 'Out Len']

                st.dataframe(det_f, height=500, use_container_width=True)

            # ==========================================================
            # VISUAL INSIGHTS
            # ==========================================================

            st.divider()

            st.subheader("3. Visual Insights & Analysis")

            st.plotly_chart(px.bar(disp, x='Order ID', y='Diff Area (m²)', color='Diff (m)', template=plotly_template), use_container_width=True)

            st.plotly_chart(px.bar(disp.sort_values(by='Cut Scrap (m)', ascending=False), x='Order ID', y='Cut Scrap (m)', template=plotly_template), use_container_width=True)

            # ==========================================================
            # EXECUTIVE SUMMARY
            # ==========================================================

            st.divider()

            st.subheader("4. Executive Summary")

            t_in = disp['Input (m)'].sum()

            t_out = disp['Output (m)'].sum()

            area_s = abs(disp[disp['Diff (m)'] < 0]['Diff Area (m²)'].sum())

            st.markdown(f"""
**Production Summary**

• **Total Input:** {t_in:,.0f} m  
• **Total Output:** {t_out:,.0f} m  
• **Area Shortfall:** {area_s:,.2f} m²
""")

            # ==========================================================
            # EXPORT
            # ==========================================================

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
```
