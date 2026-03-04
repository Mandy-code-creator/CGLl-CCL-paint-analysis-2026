import streamlit as st
import pandas as pd
import plotly.express as px


# Set page configuration
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between Galvanizing (CGL) mother coils 
and Color Coating (CCL) baby coils to estimate hidden paint loss.
""")

# Single file uploader for the master file
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        # Load the master file
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        
        # --- COLUMN MAPPING BASED ON USER'S ACTUAL DATA ---
        order_col = "訂單號碼"
        
        # CGL (Input/Mother) columns
        cgl_thick = "镀锌實測厚度"
        cgl_width = "镀锌測寬度"
        cgl_len = "镀锌測長度"
        
        # CCL (Output/Baby) columns
        ccl_thick = "實測厚度"
        ccl_width = "實測寬度"
        ccl_len = "實測長度"

        with st.spinner('Aggregating data by Order Number...'):
            # Define aggregation logic
            # Assuming 1 Order Number has duplicated CGL data for multiple CCL baby rows.
            # We take 'first' for CGL data to avoid over-summing duplicated mother coils, 
            # and 'sum' for CCL length to get the total baby coils length.
            aggregation_logic = {
                cgl_thick: 'mean',
                cgl_width: 'mean',
                cgl_len: 'first', # Get the original mother coil length
                ccl_thick: 'mean',
                ccl_width: 'mean',
                ccl_len: 'sum'    # Sum all baby coils lengths
            }
            
            # Group by Order Number (訂單號碼)
            df_summary = df.groupby(order_col).agg(aggregation_logic).reset_index()

            # --- RENAME COLUMNS TO MATCH USER'S REQUEST ---
            df_summary.rename(columns={
                cgl_len: 'SUM 镀锌測長度',
                ccl_len: 'SUM 子鋼捲'
            }, inplace=True)

            # --- CALCULATE VARIANCES ---
            # Delta Length CGL-CCL = SUM Baby Coils - SUM Mother Coils
            df_summary['Delta Length CGL-CCL'] = df_summary['SUM 子鋼捲'] - df_summary['SUM 镀锌測長度']
            
            # Thickness Variance = CCL Thickness - CGL Thickness
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            
            # Calculate Extra Surface Area (m2) due to elongation
            # Extra Area = Average CCL Width (m) * Delta Length (m)
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta Length CGL-CCL']

        st.success("Data processed successfully!")

        # Display Summary Metrics
        st.subheader("Overview")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Orders Analyzed", len(df_summary))
        m2.metric("Total Elongation / Delta Length (m)", f"{df_summary['Delta Length CGL-CCL'].sum():,.2f}")
        m3.metric("Total Extra Paint Area (m²)", f"{df_summary['Extra_Area_m2'].sum():,.2f}")

        # Visualization: Scatter plot (Thickness loss vs Length gain)
        st.subheader("Correlation: Thickness Reduction vs Length Elongation")
        fig = px.scatter(
            df_summary, 
            x='Thickness_Variance', 
            y='Delta Length CGL-CCL', 
            hover_data=[order_col, 'SUM 镀锌測長度', 'SUM 子鋼捲'],
            color='Extra_Area_m2',
            color_continuous_scale='Viridis',
            title="Analysis of Coil Elongation during CCL Process",
            labels={
                'Thickness_Variance': 'Thickness Change (CCL - CGL)',
                'Delta Length CGL-CCL': 'Length Gain (m)',
                'Extra_Area_m2': 'Extra Area (m²)'
            }
        )
        # Add reference lines
        fig.add_hline(y=0, line_dash="dash", line_color="red")
        fig.add_vline(x=0, line_dash="dash", line_color="red")
        st.plotly_chart(fig, use_container_width=True)

        # Display Detailed Table
        st.subheader("Detailed Order Data")
        
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
        st.error(f"Missing column in your file: {e}")
        st.info("Please ensure your uploaded Excel file has exact column names like '訂單號碼', '镀锌實測厚度', etc.")
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
