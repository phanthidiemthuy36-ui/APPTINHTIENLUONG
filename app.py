
import streamlit as st
import pandas as pd
from io import BytesIO

# =========================
# CẤU HÌNH TRANG WEB
# =========================
st.set_page_config(
    page_title="Tính thuế TNCN và lương thực lãnh",
    page_icon="💰",
    layout="wide"
)

st.title("💰 PHẦN MỀM TÍNH TIỀN LƯƠNG - DIỄM THÚY")
st.caption(
    "Công cụ ước tính thuế thu nhập cá nhân và tiền lương thực nhận "
    "hằng tháng của người lao động."
)

# =========================
# HÀM ĐỊNH DẠNG TIỀN
# =========================
def dinh_dang_tien(so_tien):
    return f"{so_tien:,.0f}".replace(",", ".") + " VNĐ"


def tinh_thue(thu_nhap_tinh_thue, bieu_thue):
    """Tính thuế theo biểu thuế được lựa chọn."""
    thu_nhap_tinh_thue = max(0.0, thu_nhap_tinh_thue)

    if bieu_thue == "Biểu thuế năm 2026":
        if thu_nhap_tinh_thue <= 10:
            return thu_nhap_tinh_thue * 0.05
        elif thu_nhap_tinh_thue <= 30:
            return thu_nhap_tinh_thue * 0.15 - 1
        elif thu_nhap_tinh_thue <= 60:
            return thu_nhap_tinh_thue * 0.25 - 4
        elif thu_nhap_tinh_thue <= 100:
            return thu_nhap_tinh_thue * 0.30 - 7
        else:
            return thu_nhap_tinh_thue * 0.35 - 12

    # Biểu thuế giữ nguyên theo code ban đầu của người dùng
    if thu_nhap_tinh_thue <= 10:
        return thu_nhap_tinh_thue * 0.05
    elif thu_nhap_tinh_thue <= 30:
        return thu_nhap_tinh_thue * 0.10 - 0.5
    elif thu_nhap_tinh_thue <= 60:
        return thu_nhap_tinh_thue * 0.20 - 3.5
    elif thu_nhap_tinh_thue <= 100:
        return thu_nhap_tinh_thue * 0.30 - 9.5
    else:
        return thu_nhap_tinh_thue * 0.35 - 14.5


# =========================
# NHẬP THÔNG TIN LƯƠNG
# =========================
st.subheader("1. Nhập thông tin thu nhập")

with st.form("form_tinh_luong"):
    col1, col2 = st.columns(2)

    with col1:
        luong_co_ban = st.number_input(
            "Lương cơ bản (triệu đồng)",
            min_value=0.0,
            value=15.0,
            step=0.5,
            format="%.2f"
        )

        tien_thuong = st.number_input(
            "Tiền thưởng (triệu đồng)",
            min_value=0.0,
            value=0.0,
            step=0.5,
            format="%.2f"
        )

        phu_cap_co_bhxh = st.number_input(
            "Phụ cấp tính đóng BHXH (triệu đồng)",
            min_value=0.0,
            value=0.0,
            step=0.5,
            format="%.2f"
        )

    with col2:
        phu_cap_khong_bhxh = st.number_input(
            "Phụ cấp không tính đóng BHXH (triệu đồng)",
            min_value=0.0,
            value=0.0,
            step=0.5,
            format="%.2f"
        )

        so_nguoi_phu_thuoc = st.number_input(
            "Số người phụ thuộc",
            min_value=0,
            max_value=100,
            value=0,
            step=1
        )

        ty_le_bao_hiem = st.number_input(
            "Tỷ lệ bảo hiểm người lao động đóng (%)",
            min_value=0.0,
            max_value=100.0,
            value=10.5,
            step=0.1,
            format="%.1f"
        )

    st.markdown("#### 2. Thiết lập tính thuế")

    bieu_thue = st.selectbox(
        "Biểu thuế TNCN áp dụng",
        options=[
            "Biểu thuế năm 2026",
            "Biểu thuế theo code gốc"
        ],
        index=0
    )

    giam_tru_ban_than = st.number_input(
        "Giảm trừ bản thân (triệu đồng/tháng)",
        min_value=0.0,
        value=15.5,
        step=0.5,
        format="%.2f"
    )

    giam_tru_nguoi_phu_thuoc = st.number_input(
        "Giảm trừ mỗi người phụ thuộc (triệu đồng/tháng)",
        min_value=0.0,
        value=6.2,
        step=0.1,
        format="%.2f"
    )

    da_tinh = st.form_submit_button(
        "🧮 TÍNH THUẾ VÀ LƯƠNG THỰC LÃNH",
        use_container_width=True
    )


# =========================
# TÍNH TOÁN VÀ HIỂN THỊ
# =========================
if da_tinh:
    # Đơn vị tính trong chương trình là triệu đồng
    tong_thu_nhap = (
        luong_co_ban
        + tien_thuong
        + phu_cap_co_bhxh
        + phu_cap_khong_bhxh
    )

    # Giả định toàn bộ lương cơ bản và phụ cấp BHXH
    # được tính vào căn cứ đóng bảo hiểm
    tien_bao_hiem = (
        luong_co_ban + phu_cap_co_bhxh
    ) * ty_le_bao_hiem / 100

    tong_giam_tru = (
        giam_tru_ban_than
        + giam_tru_nguoi_phu_thuoc * so_nguoi_phu_thuoc
    )

    thu_nhap_tinh_thue = max(
        0.0,
        tong_thu_nhap - tien_bao_hiem - tong_giam_tru
    )

    thue_tncn = tinh_thue(
        thu_nhap_tinh_thue,
        bieu_thue
    )

    tien_thuc_lanh = (
        tong_thu_nhap - tien_bao_hiem - thue_tncn
    )

    # Đổi triệu đồng sang VNĐ để trình bày
    st.divider()
    st.subheader("3. Kết quả tính lương")

    k1, k2, k3 = st.columns(3)

    k1.metric(
        "Tổng thu nhập",
        dinh_dang_tien(tong_thu_nhap * 1_000_000)
    )

    k2.metric(
        "Thuế TNCN",
        dinh_dang_tien(thue_tncn * 1_000_000)
    )

    k3.metric(
        "💵 Tiền thực lãnh",
        dinh_dang_tien(tien_thuc_lanh * 1_000_000)
    )

    du_lieu = [
        ["Tổng thu nhập", tong_thu_nhap],
        ["Tiền đóng bảo hiểm", tien_bao_hiem],
        ["Giảm trừ bản thân", giam_tru_ban_than],
        [
            "Giảm trừ người phụ thuộc",
            giam_tru_nguoi_phu_thuoc * so_nguoi_phu_thuoc
        ],
        ["Thu nhập tính thuế", thu_nhap_tinh_thue],
        ["Thuế TNCN phải nộp", thue_tncn],
        ["Tiền thực lãnh", tien_thuc_lanh],
    ]

    bang_ket_qua = pd.DataFrame(
        du_lieu,
        columns=["Nội dung", "Số tiền (triệu đồng)"]
    )

    bang_ket_qua["Số tiền (triệu đồng)"] = (
        bang_ket_qua["Số tiền (triệu đồng)"].round(4)
    )

    st.dataframe(
        bang_ket_qua,
        use_container_width=True,
        hide_index=True
    )

    # Tạo dữ liệu CSV để tải về
    csv_data = bang_ket_qua.to_csv(
        index=False,
        encoding="utf-8-sig"
    )

    st.download_button(
        label="📥 TẢI BẢNG KẾT QUẢ CSV",
        data=csv_data,
        file_name="ket_qua_tinh_luong.csv",
        mime="text/csv",
        use_container_width=True
    )

    st.info(
        "Kết quả là số tiền ước tính theo các tham số đã nhập. "
        "Cần kiểm tra căn cứ đóng BHXH, khoản thu nhập chịu thuế, "
        "các khoản miễn thuế và điều kiện giảm trừ trước khi sử dụng "
        "để quyết toán hoặc trả lương thực tế."
    )

else:
    st.info(
        "Nhập thông tin thu nhập, chọn biểu thuế và nhấn "
        "TÍNH THUẾ VÀ LƯƠNG THỰC LÃNH để xem kết quả."
    )

st.divider()
st.caption(
    "Đơn vị nhập liệu: triệu đồng/tháng. "
    "Đơn vị hiển thị kết quả chính: VNĐ."
)
