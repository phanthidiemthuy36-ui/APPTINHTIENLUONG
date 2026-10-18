import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
from datetime import datetime
import os
import smtplib
import ssl
from email.message import EmailMessage
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# =====================================================
# 1. CẤU HÌNH TRANG
# =====================================================
st.set_page_config(
    page_title="Phần mềm tính tiền lương - Diễm Thúy",
    page_icon="💰",
    layout="wide"
)

THU_MUC_GOC = Path(__file__).resolve().parent
DUONG_DAN_LOGO = THU_MUC_GOC / "LOGO.jpg"
MAX_NHAN_VIEN = 500

COT_NHAP = [
    "Mã nhân viên",
    "Họ và tên",
    "Email",
    "Lương cơ bản",
    "Tiền thưởng",
    "Phụ cấp tính BHXH",
    "Phụ cấp không tính BHXH",
    "Số người phụ thuộc",
]
COT_SO = [
    "Lương cơ bản",
    "Tiền thưởng",
    "Phụ cấp tính BHXH",
    "Phụ cấp không tính BHXH",
    "Số người phụ thuộc",
]
COT_TIEN = [
    "Lương cơ bản", "Tiền thưởng", "Phụ cấp tính BHXH",
    "Phụ cấp không tính BHXH", "Tổng thu nhập",
    "Bảo hiểm người lao động đóng", "Giảm trừ bản thân",
    "Giảm trừ người phụ thuộc", "Thu nhập tính thuế",
    "Thuế TNCN", "Tiền thực lãnh"
]

# =====================================================
# 2. HÀM HỖ TRỢ
# =====================================================
def dinh_dang_tien(so_tien):
    return f"{float(so_tien):,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")

def tao_du_lieu_mau(so_luong):
    return pd.DataFrame({
        "Mã nhân viên": [f"NV{i:03d}" for i in range(1, so_luong + 1)],
        "Họ và tên": ["" for _ in range(so_luong)],
        "Email": ["" for _ in range(so_luong)],
        "Lương cơ bản": [0.0] * so_luong,
        "Tiền thưởng": [0.0] * so_luong,
        "Phụ cấp tính BHXH": [0.0] * so_luong,
        "Phụ cấp không tính BHXH": [0.0] * so_luong,
        "Số người phụ thuộc": [0] * so_luong,
    })

def chuan_hoa_du_lieu(df, so_luong=None):
    df = df.copy()
    for cot in COT_NHAP:
        if cot not in df.columns:
            df[cot] = 0 if cot in COT_SO else ""
    df = df[COT_NHAP].head(MAX_NHAN_VIEN).copy()
    for cot in COT_SO:
        df[cot] = pd.to_numeric(df[cot], errors="coerce").fillna(0)
    for cot in ["Mã nhân viên", "Họ và tên", "Email"]:
        df[cot] = df[cot].fillna("").astype(str).str.strip()
    if so_luong is not None:
        df = df.head(int(so_luong)).copy()
        if len(df) < int(so_luong):
            df = pd.concat(
                [df, tao_du_lieu_mau(int(so_luong) - len(df))],
                ignore_index=True
            )
    return df.reset_index(drop=True)

def tinh_thue_tncn(thu_nhap_tinh_thue, bieu_thue):
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
        # Giữ biểu thuế trong code gốc để đối chiếu
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

def tinh_luong(nhan_vien, ty_le_bao_hiem, giam_tru_ban_than,
               giam_tru_nguoi_phu_thuoc, bieu_thue):
    luong = float(nhan_vien["Lương cơ bản"])
    thuong = float(nhan_vien["Tiền thưởng"])
    pc_co = float(nhan_vien["Phụ cấp tính BHXH"])
    pc_khong = float(nhan_vien["Phụ cấp không tính BHXH"])
    so_pt = int(nhan_vien["Số người phụ thuộc"])

    tong_thu_nhap = luong + thuong + pc_co + pc_khong
    tien_bao_hiem = (luong + pc_co) * ty_le_bao_hiem / 100
    giam_tru_pt = giam_tru_nguoi_phu_thuoc * so_pt
    thu_nhap_tinh_thue = max(
        0.0, tong_thu_nhap - tien_bao_hiem
        - giam_tru_ban_than - giam_tru_pt
    )
    thue = tinh_thue_tncn(thu_nhap_tinh_thue, bieu_thue)
    thuc_lanh = tong_thu_nhap - tien_bao_hiem - thue

    return {
        "Tổng thu nhập": tong_thu_nhap,
        "Bảo hiểm người lao động đóng": tien_bao_hiem,
        "Giảm trừ bản thân": giam_tru_ban_than,
        "Giảm trừ người phụ thuộc": giam_tru_pt,
        "Thu nhập tính thuế": thu_nhap_tinh_thue,
        "Thuế TNCN": thue,
        "Tiền thực lãnh": thuc_lanh,
    }

def tinh_bang_luong(df, ty_le_bao_hiem, giam_tru_ban_than,
                    giam_tru_nguoi_phu_thuoc, bieu_thue):
    rows = []
    for _, nv in df.iterrows():
        row = {
            "Mã nhân viên": nv["Mã nhân viên"],
            "Họ và tên": nv["Họ và tên"],
            "Email": nv["Email"],
            "Lương cơ bản": float(nv["Lương cơ bản"]),
            "Tiền thưởng": float(nv["Tiền thưởng"]),
            "Phụ cấp tính BHXH": float(nv["Phụ cấp tính BHXH"]),
            "Phụ cấp không tính BHXH": float(nv["Phụ cấp không tính BHXH"]),
            "Số người phụ thuộc": int(nv["Số người phụ thuộc"]),
        }
        row.update(tinh_luong(
            nv, ty_le_bao_hiem, giam_tru_ban_than,
            giam_tru_nguoi_phu_thuoc, bieu_thue
        ))
        rows.append(row)
    return pd.DataFrame(rows)

# =====================================================
# 3. PDF: HỖ TRỢ FONT TIẾNG VIỆT
# =====================================================
def dang_ky_font_pdf():
    candidates = [
        THU_MUC_GOC / "DejaVuSans.ttf",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/calibri.ttf"),
    ]
    for font_path in candidates:
        if font_path.exists():
            try:
                if "DejaVuSansPayroll" not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont("DejaVuSansPayroll", str(font_path)))
                bold_candidates = [
                    THU_MUC_GOC / "DejaVuSans-Bold.ttf",
                    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
                    Path("C:/Windows/Fonts/arialbd.ttf"),
                ]
                for bold_path in bold_candidates:
                    if bold_path.exists():
                        try:
                            if "DejaVuSansPayroll-Bold" not in pdfmetrics.getRegisteredFontNames():
                                pdfmetrics.registerFont(TTFont("DejaVuSansPayroll-Bold", str(bold_path)))
                            return "DejaVuSansPayroll", "DejaVuSansPayroll-Bold"
                        except Exception:
                            pass
                return "DejaVuSansPayroll", "DejaVuSansPayroll"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"

def tao_pdf_bang_luong(df_ket_qua, tieu_de="BẢNG TỔNG HỢP LƯƠNG VÀ THUẾ TNCN"):
    output = BytesIO()
    font_regular, font_bold = dang_ky_font_pdf()
    doc = SimpleDocTemplate(
        output, pagesize=landscape(A4),
        rightMargin=10*mm, leftMargin=10*mm,
        topMargin=12*mm, bottomMargin=12*mm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="PayrollTitle", fontName=font_bold, fontSize=14,
        leading=18, alignment=TA_CENTER, spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        name="PayrollSmall", fontName=font_regular, fontSize=7,
        leading=9, alignment=TA_LEFT
    ))
    styles.add(ParagraphStyle(
        name="PayrollHeader", fontName=font_bold, fontSize=7,
        leading=8, textColor=colors.white, alignment=TA_CENTER
    ))
    elements = [
        Paragraph(escape(tieu_de), styles["PayrollTitle"]),
        Paragraph(
            "Đơn vị tiền: triệu đồng/tháng | Ngày xuất: "
            + datetime.now().strftime("%d/%m/%Y %H:%M"),
            styles["PayrollSmall"]
        ),
        Spacer(1, 5*mm)
    ]
    columns = [
        "Mã nhân viên", "Họ và tên", "Tổng thu nhập",
        "Bảo hiểm người lao động đóng", "Thu nhập tính thuế",
        "Thuế TNCN", "Tiền thực lãnh"
    ]
    header = [Paragraph(escape(c), styles["PayrollHeader"]) for c in columns]
    data = [header]
    for _, r in df_ket_qua.iterrows():
        data.append([
            Paragraph(escape(str(r["Mã nhân viên"])), styles["PayrollSmall"]),
            Paragraph(escape(str(r["Họ và tên"])), styles["PayrollSmall"]),
            f'{r["Tổng thu nhập"]:,.3f}',
            f'{r["Bảo hiểm người lao động đóng"]:,.3f}',
            f'{r["Thu nhập tính thuế"]:,.3f}',
            f'{r["Thuế TNCN"]:,.3f}',
            f'{r["Tiền thực lãnh"]:,.3f}',
        ])
    table = Table(
        data,
        repeatRows=1,
        colWidths=[25*mm, 42*mm, 29*mm, 36*mm, 30*mm, 25*mm, 30*mm]
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), font_regular),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F6FA")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(table)
    doc.build(elements)
    output.seek(0)
    return output.getvalue()

def tao_pdf_phieu_luong(nhan_vien, thang_luong, font_regular=None, font_bold=None):
    output = BytesIO()
    if font_regular is None or font_bold is None:
        font_regular, font_bold = dang_ky_font_pdf()

    doc = SimpleDocTemplate(
        output, pagesize=A4,
        rightMargin=20*mm, leftMargin=20*mm,
        topMargin=18*mm, bottomMargin=18*mm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="SlipTitle", fontName=font_bold, fontSize=16,
        leading=20, alignment=TA_CENTER, spaceAfter=8
    ))
    styles.add(ParagraphStyle(
        name="SlipText", fontName=font_regular, fontSize=10,
        leading=15
    ))
    styles.add(ParagraphStyle(
        name="SlipBold", fontName=font_bold, fontSize=10,
        leading=15
    ))

    elements = [
        Paragraph("PHIẾU LƯƠNG CÁ NHÂN", styles["SlipTitle"]),
        Paragraph("Kỳ lương: " + escape(thang_luong), styles["SlipText"]),
        Spacer(1, 5*mm),
        Paragraph("Mã nhân viên: " + escape(str(nhan_vien["Mã nhân viên"])), styles["SlipText"]),
        Paragraph("Họ và tên: " + escape(str(nhan_vien["Họ và tên"])), styles["SlipText"]),
        Spacer(1, 4*mm)
    ]
    labels = [
        ("Lương cơ bản", "Lương cơ bản"),
        ("Tiền thưởng", "Tiền thưởng"),
        ("Phụ cấp tính BHXH", "Phụ cấp tính BHXH"),
        ("Phụ cấp không tính BHXH", "Phụ cấp không tính BHXH"),
        ("Tổng thu nhập", "Tổng thu nhập"),
        ("Bảo hiểm người lao động đóng", "Các khoản bảo hiểm người lao động đóng"),
        ("Giảm trừ bản thân", "Giảm trừ bản thân"),
        ("Giảm trừ người phụ thuộc", "Giảm trừ người phụ thuộc"),
        ("Thu nhập tính thuế", "Thu nhập tính thuế"),
        ("Thuế TNCN", "Thuế thu nhập cá nhân"),
        ("Tiền thực lãnh", "TIỀN THỰC LÃNH"),
    ]
    data = [[
        Paragraph("Khoản mục", styles["SlipBold"]),
        Paragraph("Số tiền (triệu đồng)", styles["SlipBold"])
    ]]
    for key, label in labels:
        data.append([
            Paragraph(escape(label), styles["SlipText"]),
            Paragraph(f'{float(nhan_vien[key]):,.3f}', styles["SlipText"])
        ])
    table = Table(data, colWidths=[105*mm, 55*mm], hAlign="LEFT")
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
        ("FONTNAME", (0, 0), (-1, -1), font_regular),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2F0D9")),
        ("FONTNAME", (0, -1), (-1, -1), font_bold),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.extend([
        table, Spacer(1, 8*mm),
        Paragraph(
            "Phiếu lương này chứa thông tin thu nhập cá nhân. "
            "Vui lòng bảo mật và liên hệ bộ phận nhân sự nếu cần đối chiếu.",
            styles["SlipText"]
        )
    ])
    doc.build(elements)
    output.seek(0)
    return output.getvalue()

# =====================================================
# 4. EMAIL: CẤU HÌNH VÀ GỬI RIÊNG TỪNG NGƯỜI
# =====================================================
def doc_cau_hinh_email():
    try:
        cfg = st.secrets.get("email", {})
        return {
            "smtp_server": str(cfg.get("smtp_server", "")),
            "smtp_port": int(cfg.get("smtp_port", 587)),
            "sender_email": str(cfg.get("sender_email", "")),
            "sender_password": str(cfg.get("sender_password", "")),
            "use_tls": bool(cfg.get("use_tls", True)),
        }
    except Exception:
        return {
            "smtp_server": "", "smtp_port": 587,
            "sender_email": "", "sender_password": "", "use_tls": True
        }

def gui_email_phieu_luong(cau_hinh, email_nhan, ten_nhan, thang_luong, pdf_data):
    msg = EmailMessage()
    msg["Subject"] = f"Phiếu lương cá nhân - {thang_luong}"
    msg["From"] = cau_hinh["sender_email"]
    msg["To"] = email_nhan
    msg.set_content(
        f"Kính gửi {ten_nhan},\n\n"
        f"Phòng phụ trách gửi bạn phiếu lương kỳ {thang_luong} trong file PDF đính kèm.\n"
        "Vui lòng kiểm tra thông tin và phản hồi bộ phận phụ trách nếu có sai sót.\n\n"
        "Trân trọng."
    )
    msg.add_attachment(
        pdf_data,
        maintype="application",
        subtype="pdf",
        filename=f"Phieu_luong_{thang_luong.replace('/', '-')}.pdf"
    )

    context = ssl.create_default_context()
    with smtplib.SMTP(cau_hinh["smtp_server"], cau_hinh["smtp_port"], timeout=30) as server:
        if cau_hinh["use_tls"]:
            server.starttls(context=context)
        server.login(cau_hinh["sender_email"], cau_hinh["sender_password"])
        server.send_message(msg)

# =====================================================
# 5. LOGO VÀ TIÊU ĐỀ
# =====================================================
if DUONG_DAN_LOGO.exists():
    st.image(str(DUONG_DAN_LOGO), width=190)

st.title("💰 PHẦN MỀM TÍNH TIỀN LƯƠNG - DIỄM THÚY")
st.caption("Tính lương, thuế TNCN, xuất Excel/PDF và gửi phiếu lương riêng cho người lao động.")
st.warning(
    "Bảng lương là dữ liệu nhạy cảm. Email sẽ được gửi riêng từng người, "
    "mỗi người chỉ nhận phiếu lương PDF của chính mình. Hãy kiểm tra địa chỉ email trước khi gửi."
)

# =====================================================
# 6. THIẾT LẬP CHUNG
# =====================================================
st.subheader("1. Thiết lập tính lương")
c1, c2, c3 = st.columns(3)
with c1:
    so_nhan_vien = int(st.number_input(
        "Số người lao động cần tính", min_value=1,
        max_value=MAX_NHAN_VIEN, value=10, step=1
    ))
with c2:
    ty_le_bao_hiem = st.number_input(
        "Tỷ lệ bảo hiểm người lao động đóng (%)",
        min_value=0.0, max_value=100.0, value=10.5, step=0.1
    )
with c3:
    bieu_thue = st.selectbox(
        "Biểu thuế áp dụng",
        ["Biểu thuế theo code gốc", "Biểu thuế 5 bậc năm 2026"],
        index=0
    )

c4, c5 = st.columns(2)
with c4:
    giam_tru_ban_than = st.number_input(
        "Giảm trừ bản thân (triệu đồng/tháng)",
        min_value=0.0, value=15.5, step=0.5
    )
with c5:
    giam_tru_nguoi_phu_thuoc = st.number_input(
        "Giảm trừ mỗi người phụ thuộc (triệu đồng/tháng)",
        min_value=0.0, value=6.2, step=0.1
    )

st.caption(
    "Đơn vị tiền nhập: triệu đồng/tháng. Biểu thuế và mức giảm trừ cần được xác minh "
    "theo kỳ tính thuế trước khi dùng để trả lương chính thức."
)

# =====================================================
# 7. NHẬP DỮ LIỆU
# =====================================================
st.subheader("2. Nhập danh sách người lao động")
mau = BytesIO()
with pd.ExcelWriter(mau, engine="openpyxl") as writer:
    tao_du_lieu_mau(so_nhan_vien).to_excel(writer, sheet_name="Danh_sach", index=False)
mau.seek(0)

st.download_button(
    "📥 Tải file Excel mẫu nhập lương",
    data=mau.getvalue(),
    file_name="mau_nhap_luong_va_email.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded = st.file_uploader("Tải file Excel danh sách (không bắt buộc)", type=["xlsx"])

if uploaded is not None:
    try:
        df_nguon = chuan_hoa_du_lieu(pd.read_excel(uploaded), so_nhan_vien)
    except Exception as e:
        st.error(f"Không đọc được file Excel: {e}")
        df_nguon = tao_du_lieu_mau(so_nhan_vien)
else:
    if "du_lieu_nhan_vien" not in st.session_state:
        st.session_state.du_lieu_nhan_vien = tao_du_lieu_mau(MAX_NHAN_VIEN)
    df_nguon = st.session_state.du_lieu_nhan_vien.iloc[:so_nhan_vien].copy()

st.info(
    f"Đang nhập {so_nhan_vien} dòng. Cột Email dùng để gửi phiếu lương PDF riêng cho từng người."
)

df_nhap = st.data_editor(
    df_nguon,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    key=f"editor_{so_nhan_vien}_{uploaded.name if uploaded else 'manual'}",
    column_config={
        "Mã nhân viên": st.column_config.TextColumn("Mã nhân viên"),
        "Họ và tên": st.column_config.TextColumn("Họ và tên", required=True),
        "Email": st.column_config.TextColumn("Email nhận phiếu lương"),
        "Lương cơ bản": st.column_config.NumberColumn("Lương cơ bản (triệu)", min_value=0.0, format="%.3f"),
        "Tiền thưởng": st.column_config.NumberColumn("Tiền thưởng (triệu)", min_value=0.0, format="%.3f"),
        "Phụ cấp tính BHXH": st.column_config.NumberColumn("Phụ cấp tính BHXH (triệu)", min_value=0.0, format="%.3f"),
        "Phụ cấp không tính BHXH": st.column_config.NumberColumn("Phụ cấp không tính BHXH (triệu)", min_value=0.0, format="%.3f"),
        "Số người phụ thuộc": st.column_config.NumberColumn("Số người phụ thuộc", min_value=0, step=1, format="%d"),
    }
)

if uploaded is None:
    du_lieu_luu = st.session_state.du_lieu_nhan_vien.copy()
    du_lieu_luu.loc[:so_nhan_vien-1, COT_NHAP] = df_nhap[COT_NHAP].reset_index(drop=True).values
    st.session_state.du_lieu_nhan_vien = du_lieu_luu

# =====================================================
# 8. TÍNH LƯƠNG
# =====================================================
if st.button("🧮 TÍNH LƯƠNG CHO DANH SÁCH", type="primary", use_container_width=True):
    df_tinh = chuan_hoa_du_lieu(df_nhap, so_nhan_vien)
    if (df_tinh["Họ và tên"].str.strip() == "").any():
        st.error("Vui lòng nhập đầy đủ họ và tên cho các dòng đang tính.")
        st.stop()
    if df_tinh[COT_SO].isna().any().any() or (df_tinh[COT_SO] < 0).any().any():
        st.error("Các khoản lương và số người phụ thuộc phải là số không âm.")
        st.stop()
    if (df_tinh["Số người phụ thuộc"] % 1 != 0).any():
        st.error("Số người phụ thuộc phải là số nguyên.")
        st.stop()

    df_kq = tinh_bang_luong(
        df_tinh, ty_le_bao_hiem, giam_tru_ban_than,
        giam_tru_nguoi_phu_thuoc, bieu_thue
    )
    st.session_state["ket_qua_luong"] = df_kq
    st.session_state["tham_so_luong"] = {
        "Tỷ lệ bảo hiểm (%)": ty_le_bao_hiem,
        "Giảm trừ bản thân (triệu đồng/tháng)": giam_tru_ban_than,
        "Giảm trừ mỗi người phụ thuộc (triệu đồng/tháng)": giam_tru_nguoi_phu_thuoc,
        "Biểu thuế": bieu_thue,
        "Số người lao động": len(df_kq),
    }

# =====================================================
# 9. KẾT QUẢ + EXCEL + PDF
# =====================================================
if "ket_qua_luong" in st.session_state:
    df_kq = st.session_state["ket_qua_luong"]
    st.divider()
    st.subheader("3. Kết quả tính lương")
    a, b, c, d = st.columns(4)
    a.metric("Số người", f"{len(df_kq):,}")
    b.metric("Tổng thu nhập", dinh_dang_tien(df_kq["Tổng thu nhập"].sum()) + " triệu")
    c.metric("Tổng thuế TNCN", dinh_dang_tien(df_kq["Thuế TNCN"].sum()) + " triệu")
    d.metric("Tổng thực lãnh", dinh_dang_tien(df_kq["Tiền thực lãnh"].sum()) + " triệu")
    st.dataframe(df_kq, use_container_width=True, hide_index=True)

    excel_output = BytesIO()
    with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
        df_kq.to_excel(writer, sheet_name="Bang_luong", index=False)
        chi_tiet_cols = [
            "Mã nhân viên", "Họ và tên", "Email", "Tổng thu nhập",
            "Bảo hiểm người lao động đóng", "Giảm trừ bản thân",
            "Giảm trừ người phụ thuộc", "Thu nhập tính thuế",
            "Thuế TNCN", "Tiền thực lãnh"
        ]
        df_kq[chi_tiet_cols].to_excel(writer, sheet_name="Chi_tiet", index=False)
        pd.DataFrame(list(st.session_state["tham_so_luong"].items()),
                     columns=["Tham số", "Giá trị"]).to_excel(
            writer, sheet_name="Tham_so", index=False
        )
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                width = min(max(max(len(str(x.value or "")) for x in cells) + 2, 12), 32)
                ws.column_dimensions[cells[0].column_letter].width = width
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.000"
    excel_output.seek(0)

    st.download_button(
        "📊 TẢI FILE EXCEL BẢNG LƯƠNG",
        data=excel_output.getvalue(),
        file_name="bang_luong_va_thue_TNCN.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    pdf_all = tao_pdf_bang_luong(df_kq)
    st.download_button(
        "📄 TẢI FILE PDF BẢNG LƯƠNG TỔNG HỢP",
        data=pdf_all,
        file_name="bang_luong_tong_hop.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    # =================================================
    # 10. GỬI EMAIL PHIẾU LƯƠNG CÁ NHÂN
    # =================================================
    st.divider()
    st.subheader("4. Gửi email phiếu lương cho người lao động")
    st.write(
        "Chọn các dòng muốn gửi. Mỗi người sẽ nhận một email riêng, "
        "đính kèm PDF chỉ chứa phiếu lương của chính họ."
    )

    df_email = df_kq.copy()
    df_email.insert(0, "Gửi email?", False)

    df_chon = st.data_editor(
        df_email[["Gửi email?", "Mã nhân viên", "Họ và tên", "Email", "Tiền thực lãnh"]],
        use_container_width=True,
        hide_index=True,
        key="chon_email_editor",
        disabled=["Mã nhân viên", "Họ và tên", "Email", "Tiền thực lãnh"],
        column_config={
            "Gửi email?": st.column_config.CheckboxColumn("Chọn gửi", default=False),
            "Tiền thực lãnh": st.column_config.NumberColumn("Tiền thực lãnh (triệu)", format="%.3f")
        }
    )

    email_co_chon = df_chon[df_chon["Gửi email?"] == True].copy()
    st.caption(f"Đã chọn {len(email_co_chon)} người để gửi email.")

    thang_luong = st.text_input(
        "Kỳ lương ghi trên email và PDF",
        value=datetime.now().strftime("%m/%Y"),
        help="Ví dụ: 10/2026"
    )

    st.markdown("##### Cấu hình email gửi đi")
    cau_hinh = doc_cau_hinh_email()
    if not cau_hinh["smtp_server"] or not cau_hinh["sender_email"] or not cau_hinh["sender_password"]:
        st.warning(
            "Chưa có cấu hình SMTP. Hãy cấu hình trong file .streamlit/secrets.toml "
            "theo hướng dẫn bên dưới. Không nhập mật khẩu email vào mã nguồn."
        )
    else:
        st.success(f"Đã đọc cấu hình SMTP cho tài khoản gửi: {cau_hinh['sender_email']}")

    xac_nhan_gui = st.checkbox(
        "Tôi đã kiểm tra danh sách, địa chỉ email và xác nhận gửi phiếu lương cá nhân."
    )

    if st.button("✉️ GỬI EMAIL CHO NGƯỜI ĐÃ CHỌN", type="primary", use_container_width=True):
        if not email_co_chon.shape[0]:
            st.error("Bạn chưa chọn người lao động nào.")
        elif not xac_nhan_gui:
            st.error("Vui lòng đánh dấu xác nhận trước khi gửi.")
        elif not cau_hinh["smtp_server"] or not cau_hinh["sender_email"] or not cau_hinh["sender_password"]:
            st.error("Thiếu cấu hình SMTP. Hãy cấu hình secrets.toml rồi chạy lại ứng dụng.")
        else:
            # Lập bảng tra cứu kết quả theo mã nhân viên và tên.
            gui_rows = []
            loi_rows = []
            tien_trinh = st.progress(0)
            thong_bao = st.empty()

            for idx, (_, selected) in enumerate(email_co_chon.iterrows(), start=1):
                email_nhan = str(selected["Email"]).strip()
                ma_nv = str(selected["Mã nhân viên"]).strip()
                ten_nv = str(selected["Họ và tên"]).strip()

                if not email_nhan or "@" not in email_nhan or "." not in email_nhan.split("@")[-1]:
                    loi_rows.append({
                        "Mã nhân viên": ma_nv,
                        "Họ và tên": ten_nv,
                        "Email": email_nhan,
                        "Lỗi": "Email trống hoặc không hợp lệ"
                    })
                    tien_trinh.progress(idx / len(email_co_chon))
                    continue

                matches = df_kq[
                    (df_kq["Mã nhân viên"].astype(str) == ma_nv)
                    & (df_kq["Họ và tên"].astype(str) == ten_nv)
                ]
                if matches.empty:
                    loi_rows.append({
                        "Mã nhân viên": ma_nv,
                        "Họ và tên": ten_nv,
                        "Email": email_nhan,
                        "Lỗi": "Không tìm thấy dữ liệu lương tương ứng"
                    })
                    tien_trinh.progress(idx / len(email_co_chon))
                    continue

                nhan_vien = matches.iloc[0]
                try:
                    pdf_ca_nhan = tao_pdf_phieu_luong(nhan_vien, thang_luong)
                    gui_email_phieu_luong(
                        cau_hinh, email_nhan, ten_nv, thang_luong, pdf_ca_nhan
                    )
                    gui_rows.append({
                        "Mã nhân viên": ma_nv,
                        "Họ và tên": ten_nv,
                        "Email": email_nhan,
                        "Trạng thái": "Đã gửi"
                    })
                except Exception as e:
                    loi_rows.append({
                        "Mã nhân viên": ma_nv,
                        "Họ và tên": ten_nv,
                        "Email": email_nhan,
                        "Lỗi": str(e)
                    })

                tien_trinh.progress(idx / len(email_co_chon))
                thong_bao.info(f"Đang xử lý {idx}/{len(email_co_chon)} email...")

            st.session_state["ket_qua_gui_email"] = {
                "thanh_cong": pd.DataFrame(gui_rows),
                "loi": pd.DataFrame(loi_rows)
            }
            if gui_rows:
                st.success(f"Đã gửi thành công {len(gui_rows)} email.")
            if loi_rows:
                st.warning(f"Có {len(loi_rows)} email chưa gửi được. Xem chi tiết bên dưới.")

    if "ket_qua_gui_email" in st.session_state:
        kq_gui = st.session_state["ket_qua_gui_email"]
        if not kq_gui["thanh_cong"].empty:
            st.markdown("**Email đã gửi**")
            st.dataframe(kq_gui["thanh_cong"], use_container_width=True, hide_index=True)
        if not kq_gui["loi"].empty:
            st.markdown("**Email chưa gửi được**")
            st.dataframe(kq_gui["loi"], use_container_width=True, hide_index=True)

# =====================================================
# 11. HƯỚNG DẪN SMTP
# =====================================================
with st.expander("⚙️ Hướng dẫn cấu hình gửi email an toàn"):
    st.markdown("""
    Ứng dụng không lưu mật khẩu email trong `app.py`. Hãy tạo thư mục `.streamlit`
    cùng cấp với `app.py`, sau đó tạo file `secrets.toml` bên trong.

    Nội dung mẫu:

    ```toml
    [email]
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = "dia_chi_email_cua_ban@gmail.com"
    sender_password = "MAT_KHAU_UNG_DUNG"
    use_tls = true
    ```

    Với Gmail, hãy bật xác minh 2 bước và dùng **mật khẩu ứng dụng** nếu tài khoản
    của bạn yêu cầu. Không dùng mật khẩu Gmail thông thường. Nếu dùng Microsoft 365
    hoặc máy chủ email của cơ quan, hãy thay SMTP server/port theo hướng dẫn của quản trị viên.

    Không chia sẻ file `secrets.toml` hoặc đưa file này lên GitHub. Email lương là thông tin
    nhạy cảm; hãy kiểm tra kỹ địa chỉ email trước khi gửi.
    """)

st.divider()
st.caption(
    "PHẦN MỀM TÍNH TIỀN LƯƠNG - DIỄM THÚY | "
    "Cần xác minh quy định thuế và bảo hiểm áp dụng trước khi dùng để trả lương thực tế."
)
