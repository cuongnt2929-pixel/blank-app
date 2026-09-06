import pandas as pd
import streamlit as st

# Cấu hình giao diện trang web
st.set_page_config(
    page_title="Truy Xuất Dữ Liệu Khách Hàng", page_icon="📊", layout="wide"
)

st.title("📱 Ứng Dụng Truy Xuất & Quản Lý Dữ Liệu Khách Hàng")
st.markdown(
    "Tải lên file Excel chứa dữ liệu khách hàng để tra cứu, lọc và xuất file mới."
)

# 1. Tải file Excel lên
uploaded_file = st.file_uploader(
    "Chọn file Excel khách hàng (.xlsx, .xls)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    # Đọc file Excel
    df = pd.read_excel(uploaded_file)

    st.success(
        f"Đã tải lên thành công! Tổng số dòng dữ liệu: {len(df)} khách hàng."
    )

    # Hiển thị bảng dữ liệu gốc (thu gọn)
    with st.expander("🔍 Xem trước toàn bộ dữ liệu gốc"):
      st.dataframe(df, use_container_width=True)

    st.markdown("---")
    st.subheader("🔎 Tìm kiếm và Lọc dữ liệu khách hàng")

    # Tạo các tùy chọn tìm kiếm/lọc linh hoạt
    # Chọn cột để tìm kiếm
    search_columns = st.multiselect(
        "Chọn các cột muốn tìm kiếm:",
        options=df.columns.tolist(),
        default=df.columns.tolist()[:2]
        if len(df.columns) >= 2
        else df.columns.tolist(),
    )

    search_query = st.text_input(
        "Nhập từ khóa tìm kiếm (tên, số điện thoại, mã KH, địa chỉ,...):"
    )

    # Lọc dữ liệu dựa trên từ khóa
    filtered_df = df.copy()
    if search_query and search_columns:
      # Lọc trên các cột được chọn
      mask = False
      for col in search_columns:
        mask = mask | filtered_df[col].astype(str).str.contains(
            search_query, case=False, na=False
        )
      filtered_df = filtered_df[mask]

    st.markdown(
        f"**Kết quả tìm kiếm:** Tìm thấy **{len(filtered_df)}** bản ghi phù hợp."
    )

    # Hiển thị bảng kết quả đã lọc
    st.dataframe(filtered_df, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Xuất dữ liệu ra File Excel mới")

    # Tùy chọn xuất tất cả hay chỉ xuất kết quả đang lọc
    export_option = st.radio(
        "Chọn phạm vi dữ liệu cần xuất:",
        ["Xuất toàn bộ dữ liệu", "Chỉ xuất dữ liệu đang lọc/tìm kiếm"],
    )

    data_to_export = (
        filtered_df
        if export_option == "Chỉ xuất dữ liệu đang lọc/tìm kiếm"
        else df
    )

    # Đặt tên file xuất
    output_filename = st.text_input(
        "Tên file Excel mới xuất ra:", value="Danh_sach_khach_hang_moi.xlsx"
    )
    if not output_filename.endswith(".xlsx"):
      output_filename += ".xlsx"


    # Hàm chuyển đổi dataframe thành file excel trong bộ nhớ để tải về
    def convert_df_to_excel(df_to_save):
      from io import BytesIO
output = BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_save.to_excel(writer, index=False, sheet_name="KhachHang")
      processed_data = output.getvalue()
      return processed_data


    excel_data = convert_df_to_excel(data_to_export)

    # Nút bấm tải file Excel mới
    st.download_button(
        label="📥 Tải xuống File Excel Mới",
        data=excel_data,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  except Exception as e:
    st.error(
        f"Đã xảy ra lỗi khi đọc file Excel. Vui lòng kiểm tra lại định dạng file!"
    )
    st.exception(e)
else:
  st.info("Vui lòng tải lên một file Excel để bắt đầu sử dụng ứng dụng.")