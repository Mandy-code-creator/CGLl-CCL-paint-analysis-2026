import streamlit as st
import pandas as pd
import plotly.express as px

# Set page configuration
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

st.title("塗料耗損分析：鍍鋅(CGL)與彩塗(CCL)長度延伸差異")
st.markdown("""
本應用程式透過比對鍍鋅母捲與彩塗子捲的實際長度，來分析因鋼帶延伸造成的隱藏塗料耗損。
""")

# Single file uploader for the master file
uploaded_file = st.file_uploader("上傳資料總表 (.xlsx 或 .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        # Load the master file
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        
        # --- Column mapping based on actual data ---
        order_col = "訂單號碼"
        mother_coil_col = "投入鋼捲號碼"
        
        # CGL (Input/Mother) columns
        cgl_thick = "镀锌實測厚度"
        cgl_width = "镀锌測寬度"
        cgl_len = "镀锌測長度"
        
        # CCL (Output/Baby) columns
        ccl_thick = "實測厚度"
        ccl_width = "實測寬度"
        ccl_len = "實測長度"

        with st.spinner('正在處理與聚合資料...'):
            # Step 1: First groupby [Order Number, Mother Coil Number]
            # This extracts the true length of each mother coil without duplication, 
            # and sums up the baby coils' lengths for that specific mother coil.
            step1_agg = {
                cgl_thick: 'first',
                cgl_width: 'first',
                cgl_len: 'first', # Original length of each individual mother coil
                ccl_thick: 'mean',
                ccl_width: 'mean',
                ccl_len: 'sum'    # Sum of baby coils for this specific mother coil
            }
            df_step1 = df.groupby([order_col, mother_coil_col]).agg(step1_agg).reset_index()

            # Step 2: Second groupby [Order Number] to aggregate total order data
            step2_agg = {
                cgl_thick: 'mean',
                cgl_width: 'mean',
                cgl_len: 'sum',   # Sum of all unique mother coils' lengths in the order (SUM 镀锌測長度)
                ccl_thick: 'mean',
                ccl_width: 'mean',
                ccl_len: 'sum'    # Sum of all baby coils' lengths in the order (SUM 子鋼捲)
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            # --- Rename columns to match requirements ---
            df_summary.rename(columns={
                cgl_len: 'SUM 镀锌測長度',
                ccl_len: 'SUM 子鋼捲'
            }, inplace=True)

            # --- Calculate variances ---
            # Delta Length CGL-CCL = SUM Baby Coils - SUM Mother Coils
            df_summary['Delta Length CGL-CCL'] = df_summary['SUM 子鋼捲'] - df_summary['SUM 镀锌測長度']
            
            # Thickness Variance = CCL Thickness - CGL Thickness
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            
            # Calculate Extra Surface Area (m2) due to elongation
            # Extra Area = Average CCL Width (m) * Delta Length (m)
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta Length CGL-CCL']

        st.success("資料處理成功！")

        # Display Summary Metrics
        st.subheader("數據總覽")
        m1, m2, m3 = st.columns(3)
        m1.metric("分析訂單總數", len(df_summary))
        m2.metric("總延伸長度 (m)", f"{df_summary['Delta Length CGL-CCL'].sum():,.2f}")
        m3.metric("總額外塗層面積 (m²)", f"{df_summary['Extra_Area_m2'].sum():,.2f}")

        # Visualization: Scatter plot (Thickness loss vs Length gain)
        st.subheader("關聯分析：厚度減少與長度增加之關係")
        fig = px.scatter(
            df_summary, 
            x='Thickness_Variance', 
            y='Delta Length CGL-CCL', 
            hover_data=[order_col, 'SUM 镀锌測長度', 'SUM 子鋼捲'],
            color='Extra_Area_m2',
            color_continuous_scale='Viridis',
            title="彩塗製程中鋼帶延伸分析",
            labels={
                'Thickness_Variance': '厚度變化 (彩塗 - 鍍鋅)',
                'Delta Length CGL-CCL': '延伸長度 (m)',
                'Extra_Area_m2': '額外面積 (m²)'
            }
        )
        # Add reference lines
        fig.add_hline(y=0, line_dash="dash", line_color="red")
        fig.add_vline(x=0, line_dash="dash", line_color="red")
        st.plotly_chart(fig, use_container_width=True)

        # Display Detailed Table
        st.subheader("訂單詳細資料 (依延伸長度排序)")
        
        # Select columns to display in a clean order
        display_cols = [
            order_col, 
            cgl_thick, ccl_thick, 'Thickness_Variance',
            cgl_width, ccl_width,
            'SUM 镀锌測長度', 'SUM 子鋼捲', 'Delta Length CGL-CCL', 
            'Extra_Area_m2'
        ]
        
        # Format the dataframe for display
        st.dataframe(df_summary[display_cols].sort_values(by='Delta Length CGL-CCL', ascending=False))

    except KeyError as e:
        st.error(f"檔案中找不到欄位: {e}")
        st.info("請確保上傳的 Excel 檔案包含正確的欄位名稱 (例如 '訂單號碼', '投入鋼捲號碼' 等)。")
    except Exception as e:
        st.error(f"發生未預期的錯誤: {e}")
