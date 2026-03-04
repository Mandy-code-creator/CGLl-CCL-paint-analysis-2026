import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")
st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between Galvanizing (CGL) mother coils 
and Color Coating (CCL) baby coils to estimate hidden paint loss.
""")

# =============================
# FILE UPLOAD
# =============================
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        if isinstance(uploaded_file, str):
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.endswith('.xlsx') else pd.read_csv(uploaded_file)
        else:
            df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)

        df_temp.columns = df_temp.columns.str.replace(r'\s+', '', regex=True)
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]

        st.session_state['saved_data'] = df_temp
        st.success("Data loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# =============================
# DATA PROCESSING
# =============================
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()

    order_col = "訂單號碼"
    mother_coil_col = "投入鋼捲號碼"
    baby_coil_col = "產出鋼捲號碼"

    cgl_thick = "镀锌實測厚度"
    cgl_width = "镀锌測寬度"
    cgl_len = "镀锌測長度"

    ccl_thick = "實測厚度"
    ccl_width = "實測寬度"
    ccl_len = "實測長度"

    try:
        with st.spinner('Processing and aggregating data...'):
            step1_agg = {
                cgl_thick: 'first', cgl_width: 'first', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_coil_col]).agg(step1_agg).reset_index()

            step2_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            df_summary.rename(columns={cgl_len: 'CGL_Total_Length', ccl_len: 'CCL_Total_Length'}, inplace=True)
            df_summary['Delta_Length'] = df_summary['CCL_Total_Length'] - df_summary['CGL_Total_Length']
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta_Length']

        # =============================
        # ORDER SUMMARY
        # =============================
        st.subheader("1. Order Summary")
        summary_display_cols = [
            order_col, 'CGL_Total_Length', 'CCL_Total_Length', 
            'Delta_Length', 'Thickness_Variance', 'Extra_Area_m2'
        ]
        df_summary_display = df_summary[summary_display_cols].sort_values(by='Extra_Area_m2', ascending=False).copy()
        df_summary_display.rename(columns={
            order_col: 'Order Number',
            'CGL_Total_Length': 'CGL Total Length (m)',
            'CCL_Total_Length': 'CCL Total Length (m)',
            'Delta_Length': 'Delta Length (m)',
            'Thickness_Variance': 'Thickness Variance (mm)',
            'Extra_Area_m2': 'Extra Area (m2)'
        }, inplace=True)
        st.dataframe(df_summary_display, use_container_width=True)
        st.divider()

        # =============================
        # BABY COIL DETAILS
        # =============================
        st.subheader("2. Baby Coil Details")
        st.markdown("Select an order to view its specific breakdown and length totals.")
        order_list = df[order_col].dropna().unique().tolist()
        selected_order = st.selectbox("Select Order Number:", options=order_list)

        df_detail_display = None
        if selected_order:
            order_totals = df_summary[df_summary[order_col] == selected_order].iloc[0]
            st.markdown(f"**Length Summary for Order: {selected_order}**")
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Mother Coils Length (CGL)", f"{order_totals['CCL_Total_Length'] - order_totals['Delta_Length']:,.0f} m")
            col2.metric("Total Baby Coils Length (CCL)", f"{order_totals['CCL_Total_Length']:,.0f} m")
            col3.metric("Length Variance (Delta)", f"{order_totals['Delta_Length']:,.0f} m", 
                        delta=f"{order_totals['Delta_Length']:,.0f} m", delta_color="inverse")
            st.write("")

            df_detail = df[df[order_col] == selected_order].copy()
            df_detail['Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            detail_display_cols = [mother_coil_col, baby_coil_col, cgl_thick, ccl_thick, 'Thickness_Variance', ccl_len]
            df_detail_display = df_detail[detail_display_cols].sort_values(by=mother_coil_col).copy()
            df_detail_display.rename(columns={
                mother_coil_col: 'Mother Coil',
                baby_coil_col: 'Baby Coil',
                cgl_thick: 'CGL Thickness',
                ccl_thick: 'CCL Thickness',
                'Thickness_Variance': 'Thickness Variance (mm)',
                ccl_len: 'CCL Length (m)'
            }, inplace=True)
            st.dataframe(df_detail_display, use_container_width=True)

        # =============================
        # VISUAL ANALYSIS
        # =============================
        st.subheader("3. Visual Analysis: Length & Extra Area")
        fig1 = px.bar(df_summary_display,
                      x='Order Number',
                      y='Extra Area (m2)',
                      color='Delta Length (m)',
                      text='Extra Area (m2)',
                      title="Extra Painted Area per Order",
                      labels={'Delta Length (m)': 'Delta Length (m)'})
        fig1.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig1, use_container_width=True)

        st.subheader("4. Download Plotly Chart PNG")
        try:
            chart_png = fig1.to_image(format="png", width=700, height=400, scale=2)
            st.download_button(
                label="📥 Download Chart as PNG",
                data=chart_png,
                file_name="Paint_Yield_Chart.png",
                mime="image/png"
            )
        except Exception:
            st.warning("Cannot export chart PNG: Kaleido / Chrome not available on web environment.")

        # =============================
        # OUTLIER ALERTS
        # =============================
        threshold = df_summary['Delta_Length'].mean() + 2*df_summary['Delta_Length'].std()
        outliers = df_summary[df_summary['Delta_Length'] > threshold]

        if not outliers.empty:
            st.warning("⚠️ Orders with unusually high Delta Length detected:")
            st.dataframe(outliers[[order_col, 'Delta_Length', 'Extra_Area_m2']])

        # =============================
        # EXPORT EXCEL & PDF
        # =============================
        st.subheader("5. Export Reports")

        # --- EXCEL ---
        excel_buffer = io.BytesIO()
        try:
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df_summary_display.to_excel(writer, sheet_name='Order Summary', index=False)
                if df_detail_display is not None:
                    df_detail_display.to_excel(writer, sheet_name=f'Details_{selected_order}', index=False)
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_buffer.getvalue(),
                file_name="Paint_Yield_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except ModuleNotFoundError:
            st.warning("Excel export unavailable: Please install 'XlsxWriter'.")

        # --- PDF (web-safe) ---
        try:
            from fpdf import FPDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, "Paint Yield Report", ln=True, align='C')

            pdf.set_font("Arial", "", 12)
            pdf.ln(5)
            pdf.multi_cell(0, 5, df_summary_display.to_string(index=False))

            if df_detail_display is not None:
                pdf.ln(5)
                pdf.set_font("Arial", "B", 12)
                pdf.cell(0, 10, f"Details for Order {selected_order}", ln=True)
                pdf.set_font("Arial", "", 12)
                pdf.multi_cell(0, 5, df_detail_display.to_string(index=False))

            pdf_buffer = io.BytesIO()
            pdf.output(pdf_buffer)
            pdf_buffer.seek(0)

            st.download_button(
                label="📥 Download PDF Report (Table Only)",
                data=pdf_buffer,
                file_name="Paint_Yield_Report.pdf",
                mime="application/pdf"
            )
            st.info("PDF contains tables only. Charts cannot be embedded on web environment without Chrome/Kaleido.")
        except ModuleNotFoundError:
            st.warning("PDF export unavailable: Please install 'fpdf'.")

    except KeyError as e:
        st.error(f"Missing column in your file: {e}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")

else:
    st.info("👆 Please upload your master data file (.xlsx or .csv) to begin.")
