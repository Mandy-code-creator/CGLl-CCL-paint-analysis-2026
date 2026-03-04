import streamlit as st
import pandas as pd
import plotly.express as px

# Cấu hình trang
st.set_page_config(page_title="Paint Yield Analyzer", layout="wide")

st.title("Paint Yield Analysis: CGL vs CCL Elongation")
st.markdown("""
This application analyzes the length variance between Galvanizing (CGL) mother coils 
and Color Coating (CCL) baby coils to estimate hidden paint loss.
""")

# 1. Nút tải file
uploaded_file = st.file_uploader("Upload Master Data File (.xlsx or .csv)", type=['xlsx', 'csv'])

# 2. Lưu file vào bộ nhớ tạm (Session State) để không bị mất khi thao tác
if uploaded_file is not None:
    df_temp = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
    st.session_state['saved_data'] = df_temp 
    st.success("File uploaded and saved to memory!")

# 3. Chạy ứng dụng nếu đã có dữ liệu trong bộ nhớ
if 'saved_data' in st.session_state:
    df = st.session_state['saved_data'].copy()
    
    # --- KHAI BÁO CÁC CỘT DỮ LIỆU ---
    order_col = "訂單號碼"
    mother_coil_col = "投入鋼捲號碼"
    
    # BẠN HÃY SỬA TÊN CỘT NÀY THÀNH TÊN CỘT CHỨA MÃ CUỘN CON TRONG FILE CỦA BẠN NHÉ
    # (Ví dụ: "子鋼捲號碼" hoặc "產出鋼捲號")
    baby_coil_col = "子鋼捲號碼" 
    
    # CGL (Input/Mother) columns
    cgl_thick = "镀锌實測厚度"
    cgl_width = "镀锌測寬度"
    cgl_len = "镀锌測長度"
    
    # CCL (Output/Baby) columns
    ccl_thick = "實測厚度"
    ccl_width = "實測寬度"
    ccl_len = "實測長度"

    try:
        # ==========================================
        # PHẦN 1: BẢNG TỔNG HỢP THEO ĐƠN HÀNG
        # ==========================================
        with st.spinner('Processing and aggregating data...'):
            # Gom nhóm lần 1 (Theo cuộn mẹ)
            step1_agg = {
                cgl_thick: 'first', cgl_width: 'first', cgl_len: 'first',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_step1 = df.groupby([order_col, mother_coil_col]).agg(step1_agg).reset_index()

            # Gom nhóm lần 2 (Theo Đơn hàng)
            step2_agg = {
                cgl_thick: 'mean', cgl_width: 'mean', cgl_len: 'sum',
                ccl_thick: 'mean', ccl_width: 'mean', ccl_len: 'sum'
            }
            df_summary = df_step1.groupby(order_col).agg(step2_agg).reset_index()

            # Đổi tên và tính toán chênh lệch
            df_summary.rename(columns={cgl_len: 'SUM 镀锌測長度', ccl_len: 'SUM 子鋼捲'}, inplace=True)
            df_summary['Delta Length CGL-CCL'] = df_summary['SUM 子鋼捲'] - df_summary['SUM 镀锌測長度']
            df_summary['Thickness_Variance'] = df_summary[ccl_thick] - df_summary[cgl_thick]
            df_summary['Extra_Area_m2'] = (df_summary[ccl_width] / 1000) * df_summary['Delta Length CGL-CCL']

        st.subheader("1. Order Summary (Tổng hợp theo Đơn hàng)")
        
        # Format bảng tổng hợp
        display_cols = [
            order_col, cgl_thick, ccl_thick, 'Thickness_Variance',
            'SUM 镀锌測長度', 'SUM 子鋼捲', 'Delta Length CGL-CCL', 'Extra_Area_m2'
        ]
        st.dataframe(df_summary[display_cols].sort_values(by='Delta Length CGL-CCL', ascending=False), use_container_width=True)

        st.divider() # Thêm một đường gạch ngang phân cách

        # ==========================================
        # PHẦN 2: TRA CỨU CHI TIẾT TỪNG CUỘN CON
        # ==========================================
        st.subheader("2. Baby Coil Details (Tra cứu chi tiết cuộn con)")
        st.markdown("Select an order to view the specific thickness variance for each baby coil.")
        
        # Tạo thanh chọn Đơn hàng (lấy danh sách các đơn hàng duy nhất)
        order_list = df[order_col].dropna().unique().tolist()
        selected_order = st.selectbox("Select Order Number (Chọn Đơn hàng):", options=order_list)
        
        if selected_order:
            # Lọc dữ liệu gốc chỉ lấy đơn hàng được chọn
            df_detail = df[df[order_col] == selected_order].copy()
            
            # Tính độ chênh lệch độ dày cho TỪNG CUỘN CON
            df_detail['Baby_Thickness_Variance'] = df_detail[ccl_thick] - df_detail[cgl_thick]
            
            # Sắp xếp các cột cho dễ nhìn (ưu tiên các cột quan trọng lên trước)
            # Dùng try-except để phòng trường hợp bạn chưa điền đúng tên cột Baby Coil
            try:
                detail_display_cols = [
                    mother_coil_col, baby_coil_col, 
                    cgl_thick, ccl_thick, 'Baby_Thickness_Variance',
                    cgl_width, ccl_width,
                    ccl_len
                ]
                st.dataframe(df_detail[detail_display_cols].sort_values(by=mother_coil_col), use_container_width=True)
            except KeyError:
                # Nếu không tìm thấy tên cột baby_coil, in toàn bộ bảng ra
                st.warning("Hiển thị toàn bộ cột (Hãy cập nhật biến `baby_coil_col` trong code thành tên cột mã cuộn con của bạn)")
                st.dataframe(df_detail)

    except KeyError as e:
        st.error(f"Missing column in your file: {e}")
        st.info("Vui lòng kiểm tra lại tên cột trong file Excel.")
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")

else:
    # Lời nhắc khi chưa tải file
    st.info("👆 Vui lòng tải file Excel/CSV lên để bắt đầu phân tích / Please upload your data file to begin.")
