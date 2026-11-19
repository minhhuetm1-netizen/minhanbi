"""
Module Xuất Hồ sơ Đề xuất & Báo giá B2B chuyên nghiệp (DOCX & Excel)
Tự động sinh tài liệu hoàn chỉnh sẵn sàng gửi cho khách hàng doanh nghiệp.
"""

import os
import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def generate_docx_proposal(data: dict, output_path: str) -> str:
    """Tạo file Word .docx Báo giá & Hồ sơ Đề xuất B2B hoàn chỉnh."""
    doc = Document()

    # Thiết lập lề trang tiêu chuẩn
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # 1. Header Công ty
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    
    col_left = header_table.cell(0, 0)
    col_right = header_table.cell(0, 1)
    col_left.width = Inches(4.5)
    col_right.width = Inches(2.5)

    p_logo = col_left.paragraphs[0]
    r_logo = p_logo.add_run("CÔNG TY CỔ PHẦN CÔNG NGHỆ MINH AN\n")
    r_logo.bold = True
    r_logo.font.size = Pt(11)
    r_logo.font.color.rgb = RGBColor(15, 76, 129) # Classic Navy
    
    r_sub = p_logo.add_run("Địa chỉ: Tầng 8, Tòa nhà Công nghệ Minh An, Cầu Giấy, Hà Nội\nHotline: (024) 3888 9999 | Email: kd@minhan.com.vn | Web: minhan.com.vn")
    r_sub.font.size = Pt(8.5)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    p_meta = col_right.paragraphs[0]
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_code = p_meta.add_run(f"Số HS: {data.get('ho_so_co_hoi', {}).get('ma_co_hoi', 'OPP-MINHAN')}\n")
    r_code.bold = True
    r_code.font.size = Pt(9.5)
    r_date = p_meta.add_run(f"Ngày lập: {datetime.datetime.now().strftime('%d/%m/%Y')}\nHiệu lực: 15 ngày")
    r_date.font.size = Pt(8.5)
    r_date.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 2. Tiêu đề Hồ sơ
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("THƯ CHÀO GIÁ & ĐỀ XUẤT GIẢI PHÁP CÔNG NGHỆ")
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(15, 76, 129)
    p_title.paragraph_format.space_after = Pt(12)

    # 3. Thông tin Khách hàng & Phụ trách
    dossier = data.get('ho_so_co_hoi', {})
    info_table = doc.add_table(rows=4, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_background(info_table.cell(0, 0), "F1F5F9")
    set_cell_background(info_table.cell(0, 1), "F1F5F9")

    info_table.cell(0, 0).paragraphs[0].add_run("THÔNG TIN KHÁCH HÀNG").bold = True
    info_table.cell(0, 1).paragraphs[0].add_run("ĐẠI DIỆN PHỤ TRÁCH MINH AN").bold = True

    info_table.cell(1, 0).paragraphs[0].add_run(f"Đơn vị: {dossier.get('khach_hang', '')}")
    info_table.cell(1, 1).paragraphs[0].add_run(f"Chuyên viên: {dossier.get('nvkd_phu_trach', '')}")

    info_table.cell(2, 0).paragraphs[0].add_run(f"Người liên hệ: {dossier.get('nguoi_lien_he', '')} - {dossier.get('sdt_kh', '')}")
    info_table.cell(2, 1).paragraphs[0].add_run(f"Bộ phận: {dossier.get('chuyen_mon', '')}")

    info_table.cell(3, 0).paragraphs[0].add_run(f"Email: {dossier.get('email_kh', '')}")
    info_table.cell(3, 1).paragraphs[0].add_run(f"Email: {dossier.get('email_nvkd', '')} | Hotline: 1900-6868")

    for row in info_table.rows:
        for cell in row.cells:
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 4. Bảng Báo giá Chi tiết
    p_sec1 = doc.add_paragraph()
    r_sec1 = p_sec1.add_run("I. BẢNG DỰ TOÁN THIẾT BỊ & DỊCH VỤ")
    r_sec1.bold = True
    r_sec1.font.size = Pt(11)
    r_sec1.font.color.rgb = RGBColor(15, 76, 129)

    quote = data.get('bang_bao_gia', {})
    items = quote.get('chi_tiet_hang_hoa', [])

    table = doc.add_table(rows=1 + len(items) + 5, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["STT", "Danh mục Hàng hóa / Dịch vụ", "SL", "Đơn giá niêm yết (VNĐ)", "Thành tiền (VNĐ)"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.paragraphs[0].add_run(h).bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_background(cell, "0F4C81")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        if i in [0, 2]:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif i in [3, 4]:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for idx, it in enumerate(items):
        row_idx = idx + 1
        table.cell(row_idx, 0).paragraphs[0].add_run(str(idx + 1))
        table.cell(row_idx, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.cell(row_idx, 1).paragraphs[0].add_run(it.get('ten_san_pham', ''))
        table.cell(row_idx, 2).paragraphs[0].add_run(str(it.get('so_luong', 1)))
        table.cell(row_idx, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.cell(row_idx, 3).paragraphs[0].add_run(f"{it.get('don_gia_niem_yet', 0):,.0f}")
        table.cell(row_idx, 3).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        table.cell(row_idx, 4).paragraphs[0].add_run(f"{it.get('thanh_tien_niem_yet', 0):,.0f}")
        table.cell(row_idx, 4).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Các dòng tổng hợp
    r_sub = 1 + len(items)
    summary_rows = [
        ("Tổng giá trị niêm yết:", f"{quote.get('tong_gia_niem_yet', 0):,.0f} VNĐ", False),
        (f"Chiết khấu thương mại ({quote.get('chiet_khau_pct', 0):.1f}%):", f"-{quote.get('tien_chiet_khau', 0):,.0f} VNĐ", False),
        ("Tổng giá trị trước thuế (Net):", f"{quote.get('tong_truoc_vat', 0):,.0f} VNĐ", False),
        ("Thuế giá trị gia tăng (VAT 10%):", f"{quote.get('thue_vat_10pct', 0):,.0f} VNĐ", False),
        ("TỔNG GIÁ TRỊ THANH TOÁN (SAU THUẾ):", f"{quote.get('tong_sau_vat', 0):,.0f} VNĐ", True)
    ]

    for offset, (lbl, val, is_bold) in enumerate(summary_rows):
        cur_row = r_sub + offset
        c_lbl = table.cell(cur_row, 0)
        c_val = table.cell(cur_row, 4)
        c_lbl.merge(table.cell(cur_row, 3))
        
        p1 = c_lbl.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r1 = p1.add_run(lbl)
        r1.bold = is_bold
        r1.font.size = Pt(9.5 if not is_bold else 10)

        p2 = c_val.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r2 = p2.add_run(val)
        r2.bold = is_bold
        r2.font.size = Pt(9.5 if not is_bold else 10)

        if is_bold:
            set_cell_background(c_lbl, "EFF6FF")
            set_cell_background(c_val, "EFF6FF")
            r1.font.color.rgb = RGBColor(15, 76, 129)
            r2.font.color.rgb = RGBColor(15, 76, 129)

    for row in table.rows:
        for cell in row.cells:
            if cell.paragraphs and cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].font.name = "Arial"
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 5. Cam kết Tiến độ & Dịch vụ
    stock = data.get('tham_dinh_ton_kho', {})
    p_sec2 = doc.add_paragraph()
    r_sec2 = p_sec2.add_run("II. CAM KẾT TIẾN ĐỘ & BẢO HÀNH CHÍNH HÃNG")
    r_sec2.bold = True
    r_sec2.font.size = Pt(11)
    r_sec2.font.color.rgb = RGBColor(15, 76, 129)

    doc.add_paragraph(f"• Thời hạn cung ứng: {stock.get('ghi_chu_tien_do', 'Giao hàng tiêu chuẩn theo cam kết.')}", style='List Bullet')
    doc.add_paragraph("• Tiêu chuẩn chất lượng: Hàng mới 100%, nguyên đai nguyên kiện, đầy đủ chứng nhận CO/CQ từ Hãng.", style='List Bullet')
    doc.add_paragraph("• Chế độ bảo hành: Bảo hành chính hãng tận nơi 12 - 36 tháng theo tiêu chuẩn ProSupport của Dell, Cisco, Fortinet.", style='List Bullet')
    doc.add_paragraph("• Hỗ trợ kỹ thuật: Kỹ sư Minh An trực tiếp bàn giao, cấu hình và hỗ trợ tiếp nhận sự cố SLA 24/7.", style='List Bullet')

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 6. Phần Ký duyệt & Xác nhận
    sign_table = doc.add_table(rows=3, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_table.cell(0, 0).paragraphs[0].add_run("ĐẠI DIỆN KHÁCH HÀNG\n(Ký và ghi rõ họ tên)").bold = True
    sign_table.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    appr = data.get('trinh_duyet', {})
    approver_title = appr.get('cap_phe_duyet', 'Trưởng phòng Kinh doanh')
    sign_table.cell(0, 1).paragraphs[0].add_run(f"ĐẠI DIỆN MINH AN COMPANY\n({approver_title} Phê duyệt)").bold = True
    sign_table.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    sign_table.cell(2, 0).paragraphs[0].add_run("\n\n__________________________").alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_app = sign_table.cell(2, 1).paragraphs[0]
    p_app.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_stamp = p_app.add_run(f"[ĐÃ KIỂM SOÁT THẨM QUYỀN - {appr.get('trang_thai', 'ĐÃ_DUYỆT')}]\n\n{dossier.get('nvkd_phu_trach', '')}")
    r_stamp.font.size = Pt(9)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    return output_path

