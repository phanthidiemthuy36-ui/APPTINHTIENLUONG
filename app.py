
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

# =====================================================
# 1. CẤU HÌNH TRANG
# =====================================================
st.set_page_config(
    page_title="Tính tiền lương - Diễm Thúy",
    page_icon="💰",
    layout="wide"
)

MAX_NHAN_VIEN = 500

COT_NHAP = [
    "Mã nhân viên",
    "Họ và tên",
    "Lương cơ bản (triệu đồng)",
    "Tiền thưởng (triệu đồng)",
    "Phụ cấp có BHXH (triệu đồng)",
    "Phụ cấp không BHXH (triệu đồng)",
    "Số người phụ thuộc"
]

COT_SO = [
    "Lương cơ bản (triệu đồng)",
    "Tiền thưởng (triệu đồng)",
    "Phụ cấp có BHXH (triệu đồng)",
    "Phụ cấp không BHXH (triệu đồng)",
    "Số người phụ thuộc"
]

# =====================================================
# 2. HÀM HỖ TRỢ
# =====================================================
def dinh_dang_tien(so_tien):
    return f"{so_tien:,.0f}".replace(",", ".") + " VNĐ"


def tao_du_lieu_mau(so_luong):
    """Tạo bảng trống theo số lượng người lao động."""
    df = pd.DataFrame({
        "Mã nhân viên": [f"NV{i:03d}" for i in range(1, so_luong + 1)],
        "Họ và tên": ["" for _ in range(so_luong)],
        "Lương cơ bản (triệu đồng)": [0.0] * so_luong,
        "Tiền thưởng (triệu đồng)": [0.0] * so_luong,
        "Phụ cấp có BHXH (triệu đồng)": [0.0] * so_luong,
        "Phụ cấp không BHXH (triệu đồng)": [0.0] * so_luong,
        "Số người phụ thuộc": [0] * so_luong,
    })
    return df


def chuan_hoa_du_lieu(df):
    """Đưa dữ liệu Excel về đúng cấu trúc và giới hạn 500 dòng."""
    df = df.copy()

    for cot in COT_NHAP:
        if cot not in df.columns:
            if cot in COT_SO:
                df[cot] = 0
            else:
                df[cot] = ""

    df = df[COT_NHAP].head(MAX_NHAN_VIEN).copy()

    for cot in COT_SO:
        df[cot] = pd.to_numeric(df[cot], errors="coerce").fillna(0)

    df["Mã nhân viên"] = df["Mã nhân viên"].fillna("").astype(str)
    df["Họ và tên"] = df["Họ và tên"].fillna("").astype(str)

    return df.reset_index(drop=True)


def tinh_thue(thu_nhap_tinh_thue, bieu_thue):
    """
    Đơn vị: triệu đồng/tháng.
    Biểu thuế được tính theo các mức trong ứng dụng.
    """
    thu_nhap = max(0.0, float(thu_nhap_tinh_thue))

    if bieu_thue == "Biểu thuế 5 bậc năm 2026":
        if thu_nhap <= 10:
            thue = thu_nhap * 0.05
        elif thu_nhap <= 30:
            thue = thu_nhap * 0.15 - 1
        elif thu_nhap <= 60:
            thue = thu_nhap * 0.25 - 4
        elif thu_nhap <= 100:
            thue = thu_nhap * 0.30 - 7
        else:
            thue = thu_nhap * 0.35 - 12
    else:
        # Biểu thuế theo code gốc người dùng
        if thu_nhap <= 10:
            thue = thu_nhap * 0.05
        elif thu_nhap <= 30:
            thue = thu_nhap * 0.10 - 0.5
        elif thu_nhap <= 60:
            thue = thu_nhap * 0.20 - 3.5
        elif thu_nhap <= 100:
            thue = thu_nhap * 0.30 - 9.5
        else:
            thue = thu_nhap * 0.35 - 14.5

    return max(0.0, thue)


def tinh_bang_luong(
    df,
    ty_le_bao_hiem,
    giam_tru_ban_than,
    giam_tru_nguoi_phu_thuoc,
    bieu_thue
):
    """Tính thuế và tiền thực lãnh cho toàn bộ người lao động."""

    ket_qua = []

    for _, dong in df.iterrows():
        luong = float(dong["Lương cơ bản (triệu đồng)"])
        thuong = float(dong["Tiền thưởng (triệu đồng)"])
        pc_co_bhxh = float(dong["Phụ cấp có BHXH (triệu đồng)"])
        pc_khong_bhxh = float(dong["Phụ cấp không BHXH (triệu đồng)"])
        so_nguoi_pt = int(dong["Số người phụ thuộc"])

        tong_thu_nhap = luong + thuong + pc_co_bhxh + pc_khong_bhxh

        tien_bao_hiem = (
            luong + pc_co_bhxh
        ) * ty_le_bao_hiem / 100

        giam_tru_pt = so_nguoi_pt * giam_tru_nguoi_phu_thuoc
        tong_giam_tru = giam_tru_ban_than + giam_tru_pt

        thu_nhap_tinh_thue = max(
            0.0,
            tong_thu_nhap - tien_bao_hiem - tong_giam_tru
        )

        thue_tncn = tinh_thue(thu_nhap_tinh_thue, bieu_thue)

        tien_thuc_lanh = tong_thu_nhap - tien_bao_hiem - thue_tncn

        ket_qua.append({
            "Mã nhân viên": dong["Mã nhân viên"],
            "Họ và tên": dong["Họ và tên"],
            "Lương cơ bản (triệu đồng)": luong,
            "Tiền thưởng (triệu đồng)": thuong,
            "Phụ cấp có BHXH (triệu đồng)": pc_co_bhxh,
            "Phụ cấp không BHXH (triệu đồng)": pc_khong_bhxh,
            "Số người phụ thuộc": so_nguoi_pt,
            "Tổng thu nhập (triệu đồng)": tong_thu_nhap,
            "Tiền bảo hiểm (triệu đồng)": tien_bao_hiem,
            "Giảm trừ bản thân (triệu đồng)": giam_tru_ban_than,
            "Giảm trừ người phụ thuộc (triệu đồng)": giam_tru_pt,
            "Thu nhập tính thuế (triệu đồng)": thu_nhap_tinh_thue,
            "Thuế TNCN (triệu đồng)": thue_tncn,
            "Tiền thực lãnh (triệu đồng)": tien_thuc_lanh,
        })

    return pd.DataFrame(ket_qua)


def tao_file_excel(df_ket_qua, bang_tham_so):
    """Xuất file Excel gồm bảng lương, chi tiết và tham số."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_ket_qua.to_excel(
            writer, sheet_name="Bang_luong", index=False
        )

        cot_chi_tiet = [
            "Mã nhân viên",
            "Họ và tên",
            "Tổng thu nhập (triệu đồng)",
            "Tiền bảo hiểm (triệu đồng)",
            "Giảm trừ bản thân (triệu đồng)",
            "Giảm trừ người phụ thuộc (triệu đồng)",
            "Thu nhập tính thuế (triệu đồng)",
            "Thuế TNCN (triệu đồng)",
            "Tiền thực lãnh (triệu đồng)"
        ]

        df_ket_qua[cot_chi_tiet].to_excel(
            writer, sheet_name="Chi_tiet", index=False
        )

        bang_tham_so.to_excel(
            writer, sheet_name="Tham_so", index=False
        )

        # Định dạng độ rộng cột và số liệu
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cells in ws.columns:
                max_length = max(
                    len(str(cell.value or "")) for cell in cells
                )
                ws.column_dimensions[cells[0].column_letter].width = min(
                    max(max_length + 2, 12), 35
                )

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

    output.seek(0)
    return output.getvalue()


# =====================================================
# 3. TIÊU ĐỀ VÀ LOGO
# =====================================================
duong_dan_logo = Path(__file__).parent / "LOGO.jpg"

if duong_dan_logo.exists():
    st.image(str(duong_dan_logo), width=180)

st.title("💰 PHẦN MỀM TÍNH TIỀN LƯƠNG - DIỄM THÚY")
st.caption(
    "Tính thuế thu nhập cá nhân và tiền thực lãnh "
    "cho tối đa 500 người lao động."
)

st.divider()

# =====================================================
# 4. THIẾT LẬP CHUNG
# =====================================================
st.subheader("1. Thiết lập số lượng và tham số tính lương")

col1, col2, col3 = st.columns(3)

with col1:
    so_luong = st.number_input(
        "Số người lao động cần tính",
        min_value=1,
        max_value=MAX_NHAN_VIEN,
        value=10,
        step=1
    )

with col2:
    ty_le_bao_hiem = st.number_input(
        "Tỷ lệ bảo hiểm người lao động đóng (%)",
        min_value=0.0,
        max_value=100.0,
        value=10.5,
        step=0.1
    )

with col3:
    bieu_thue = st.selectbox(
        "Biểu thuế áp dụng",
        [
            "Biểu thuế 5 bậc năm 2026",
            "Biểu thuế theo code gốc"
        ]
    )

col4, col5 = st.columns(2)

with col4:
    giam_tru_ban_than = st.number_input(
        "Giảm trừ bản thân (triệu đồng/tháng)",
        min_value=0.0,
        value=15.5,
        step=0.5
    )

with col5:
    giam_tru_nguoi_phu_thuoc = st.number_input(
        "Giảm trừ mỗi người phụ thuộc (triệu đồng/tháng)",
        min_value=0.0,
        value=6.2,
        step=0.1
    )

st.caption(
    "Các tham số thuế có thể chỉnh sửa. Hãy xác nhận quy định áp dụng "
    "cho kỳ tính thuế trước khi sử dụng kết quả chính thức."
)

st.divider()

# =====================================================
# 5. NHẬP DỮ LIỆU
# =====================================================
st.subheader("2. Nhập danh sách người lao động")

st.write(
    "Bạn có thể nhập trực tiếp vào bảng hoặc tải file Excel đã điền dữ liệu."
)

# Tạo file Excel mẫu
file_mau = BytesIO()

with pd.ExcelWriter(file_mau, engine="openpyxl") as writer:
    tao_du_lieu_mau(so_luong).to_excel(
        writer, sheet_name="Danh_sach", index=False
    )

file_mau.seek(0)

st.download_button(
    "📥 Tải file Excel mẫu",
    data=file_mau.getvalue(),
    file_name="mau_nhap_danh_sach_luong.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

file_excel = st.file_uploader(
    "Tải file Excel danh sách người lao động (không bắt buộc)",
    type=["xlsx"]
)

if file_excel is not None:
    try:
        df_upload = pd.read_excel(file_excel)

        if len(df_upload) > MAX_NHAN_VIEN:
            st.warning(
                f"File có hơn {MAX_NHAN_VIEN} dòng. "
                f"Chỉ sử dụng {MAX_NHAN_VIEN} dòng đầu tiên."
            )

        df_nguon = chuan_hoa_du_lieu(df_upload)

        if len(df_nguon) > so_luong:
            st.info(
                f"Bạn chọn tính {so_luong} người; "
                "chỉ lấy số dòng tương ứng đầu tiên."
            )
            df_nguon = df_nguon.head(so_luong).copy()

        elif len(df_nguon) < so_luong:
            bo_sung = tao_du_lieu_mau(
                so_luong - len(df_nguon)
            )
            df_nguon = pd.concat(
                [df_nguon, bo_sung], ignore_index=True
            )

        df_nguon = chuan_hoa_du_lieu(df_nguon)

    except Exception as e:
        st.error(f"Không đọc được file Excel: {e}")
        df_nguon = tao_du_lieu_mau(so_luong)
else:
    df_nguon = tao_du_lieu_mau(so_luong)

st.info(
    f"Số dòng hiện tại: {len(df_nguon)}. "
    "Nhập họ tên và các khoản thu nhập theo đơn vị triệu đồng."
)

# Key thay đổi khi số lượng người thay đổi để bảng khớp số dòng
df_da_nhap = st.data_editor(
    df_nguon,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    key=f"bang_nhap_luong_{so_luong}_{file_excel.name if file_excel else 'truc_tiep'}",
    column_config={
        "Mã nhân viên": st.column_config.TextColumn(
            "Mã nhân viên"
        ),
        "Họ và tên": st.column_config.TextColumn(
            "Họ và tên"
        ),
        "Lương cơ bản (triệu đồng)": st.column_config.NumberColumn(
            "Lương cơ bản (triệu đồng)", min_value=0.0, format="%.2f"
        ),
        "Tiền thưởng (triệu đồng)": st.column_config.NumberColumn(
            "Tiền thưởng (triệu đồng)", min_value=0.0, format="%.2f"
        ),
        "Phụ cấp có BHXH (triệu đồng)": st.column_config.NumberColumn(
            "Phụ cấp có BHXH (triệu đồng)", min_value=0.0, format="%.2f"
        ),
        "Phụ cấp không BHXH (triệu đồng)": st.column_config.NumberColumn(
            "Phụ cấp không BHXH (triệu đồng)", min_value=0.0, format="%.2f"
        ),
        "Số người phụ thuộc": st.column_config.NumberColumn(
            "Số người phụ thuộc", min_value=0, step=1, format="%d"
        ),
    }
)

st.divider()

# =====================================================
# 6. TÍNH TOÁN
# =====================================================
if st.button(
    "🧮 TÍNH THUẾ VÀ LƯƠNG THỰC LÃNH",
    type="primary",
    use_container_width=True
):
    df_tinh = df_da_nhap.copy()

    # Kiểm tra dữ liệu
    for cot in COT_SO:
        df_tinh[cot] = pd.to_numeric(
            df_tinh[cot], errors="coerce"
        )

    if df_tinh[COT_SO].isna().any().any():
        st.error(
            "Có ô số liệu bị bỏ trống hoặc không hợp lệ. "
            "Vui lòng kiểm tra lại các cột thu nhập và người phụ thuộc."
        )
        st.stop()

    if (df_tinh[COT_SO] < 0).any().any():
        st.error("Các khoản thu nhập và số người phụ thuộc không được âm.")
        st.stop()

    if (
        df_tinh["Số người phụ thuộc"]
        != df_tinh["Số người phụ thuộc"].astype(int)
    ).any():
        st.error("Số người phụ thuộc phải là số nguyên.")
        st.stop()

    if len(df_tinh) > MAX_NHAN_VIEN:
        st.error(f"Chỉ hỗ trợ tối đa {MAX_NHAN_VIEN} người.")
        st.stop()

    # Chuẩn hóa mã và tên
    df_tinh["Mã nhân viên"] = (
        df_tinh["Mã nhân viên"].fillna("").astype(str).str.strip()
    )
    df_tinh["Họ và tên"] = (
        df_tinh["Họ và tên"].fillna("").astype(str).str.strip()
    )

    if (df_tinh["Họ và tên"] == "").any():
        st.error("Vui lòng nhập đầy đủ họ và tên cho tất cả các dòng.")
        st.stop()

    # Tính toán cho toàn bộ danh sách
    df_ket_qua = tinh_bang_luong(
        df_tinh,
        ty_le_bao_hiem,
        giam_tru_ban_than,
        giam_tru_nguoi_phu_thuoc,
        bieu_thue
    )

    st.session_state["df_ket_qua_luong"] = df_ket_qua
    st.session_state["tham_so_luong"] = {
        "Tỷ lệ bảo hiểm (%)": ty_le_bao_hiem,
        "Giảm trừ bản thân (triệu đồng/tháng)": giam_tru_ban_than,
        "Giảm trừ người phụ thuộc (triệu đồng/tháng)": (
            giam_tru_nguoi_phu_thuoc
        ),
        "Biểu thuế": bieu_thue,
        "Số người lao động": len(df_ket_qua)
    }

# =====================================================
# 7. HIỂN THỊ KẾT QUẢ
# =====================================================
if "df_ket_qua_luong" in st.session_state:
    df_ket_qua = st.session_state["df_ket_qua_luong"]
    tham_so = st.session_state["tham_so_luong"]

    st.divider()
    st.subheader("3. Bảng tổng hợp kết quả tính lương")

    tong_thu_nhap = df_ket_qua[
        "Tổng thu nhập (triệu đồng)"
    ].sum()

    tong_bao_hiem = df_ket_qua[
        "Tiền bảo hiểm (triệu đồng)"
    ].sum()

    tong_thue = df_ket_qua[
        "Thuế TNCN (triệu đồng)"
    ].sum()

    tong_thuc_lanh = df_ket_qua[
        "Tiền thực lãnh (triệu đồng)"
    ].sum()

    a, b, c, d = st.columns(4)

    a.metric(
        "Số người lao động",
        f"{len(df_ket_qua):,}"
    )
    b.metric(
        "Tổng thu nhập",
        dinh_dang_tien(tong_thu_nhap * 1_000_000)
    )
    c.metric(
        "Tổng thuế TNCN",
        dinh_dang_tien(tong_thue * 1_000_000)
    )
    d.metric(
        "Tổng thực lãnh",
        dinh_dang_tien(tong_thuc_lanh * 1_000_000)
    )

    st.dataframe(
        df_ket_qua.round(4),
        use_container_width=True,
        hide_index=True
    )

    # Tạo bảng tham số
    bang_tham_so = pd.DataFrame(
        list(tham_so.items()),
        columns=["Tham số", "Giá trị"]
    )

    # Xuất Excel
    file_ket_qua = tao_file_excel(
        df_ket_qua,
        bang_tham_so
    )

    st.download_button(
        label="📊 XUẤT FILE EXCEL BẢNG LƯƠNG",
        data=file_ket_qua,
        file_name="bang_luong_va_thue_TNCN.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    # Xuất CSV nếu cần
    csv_data = df_ket_qua.to_csv(
        index=False,
        encoding="utf-8-sig"
    )

    st.download_button(
        label="📥 Xuất bảng lương CSV",
        data=csv_data,
        file_name="bang_luong_va_thue_TNCN.csv",
        mime="text/csv",
        use_container_width=True
    )

    st.caption(
        "File Excel gồm 3 sheet: Bang_luong, Chi_tiet và Tham_so. "
        "Số liệu thu nhập trong file tính theo triệu đồng."
    )

st.divider()
st.caption(
    "PHẦN MỀM TÍNH TIỀN LƯƠNG - DIỄM THÚY | "
    "Kết quả phụ thuộc dữ liệu nhập và tham số tính thuế."
)
