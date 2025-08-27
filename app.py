# app.py
import streamlit as st
from docxtpl import DocxTemplate

st.title("📄 Tạo công văn tự động")

# Form nhập liệu
with st.form("van_ban_form"):
    ten_co_quan = st.text_input("Tên cơ quan, tổ chức")
    so_ky_hieu = st.text_input("Số, ký hiệu của văn bản")
    dia_danh_thoi_gian = st.text_input("Địa danh và thời gian ban hành văn bản")
    ten_loai_trich_yeu = st.text_input("Tên loại và trích yếu nội dung văn bản")
    trich_yeu_cong_van = st.text_area("Trích yếu nội dung công văn")
    to_chuc_nhan = st.text_input("Tổ chức nhận văn bản")
    noi_dung = st.text_area("Nội dung văn bản")
    nguoi_tham_quyen = st.text_input("Người có thẩm quyền")
    noi_nhan = st.text_area("Nơi nhận")
    
    submit = st.form_submit_button("Tạo văn bản")

if submit:
    # Load template
    template_path = "data/mau_cong_van.docx"  # mẫu word của bạn
    doc = DocxTemplate(template_path)

    context = {
        "ten_co_quan": ten_co_quan,
        "so_ky_hieu": so_ky_hieu,
        "dia_danh_thoi_gian": dia_danh_thoi_gian,
        "ten_loai_trich_yeu": ten_loai_trich_yeu,
        "trich_yeu_cong_van": trich_yeu_cong_van,
        "to_chuc_nhan": to_chuc_nhan,
        "noi_dung": noi_dung,
        "nguoi_tham_quyen": nguoi_tham_quyen,
        "noi_nhan": noi_nhan,
    }

    # Render nội dung vào mẫu
    doc.render(context)
    output_path = "van_ban_hoan_thanh.docx"
    doc.save(output_path)

    with open(output_path, "rb") as file:
        st.download_button(
            label="⬇️ Tải văn bản hoàn chỉnh",
            data=file,
            file_name="van_ban.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )