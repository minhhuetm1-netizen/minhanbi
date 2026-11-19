"""
exporter.py
Xuất trọn bộ kết quả tại Bước 5:
1. File Excel tổng hợp (.xlsx) gồm nhiều tab: TongHop, ChiTietCoHoi, DuLieuLamSach
2. File Word (.docx) & Text (.txt) chứa Hồ sơ cơ hội, Email nháp và Phiếu chuyển duyệt
3. File Nhật ký truy vết (audit_log.json & audit_log.txt)
"""

import io
import json
import datetime
from typing import List, Dict, Any
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from doc_formatter import (
    setup_document_styles,
    add_markdown_to_docx,
    apply_table_styles,
    format_compact_vnd,
    format_detailed_vnd
)

def build_summary_metrics(analyzed: List[Dict[str, Any]]) -> Dict[str, Any]:
    counts = {
        "Có thể tiếp tục": 0,
        "Cần bổ sung thông tin": 0,
        "Chưa đáp ứng": 0,
        "Chờ duyệt": 0
    }
    eligible_total = 0.0
    continue_total = 0.0
    pending_total = 0.0

    for item in analyzed:
        cat = item["phan_loai"]
        if cat in counts:
            counts[cat] += 1
        val = float(item.get("gia_tri_sau_vat", 0) or 0)
        if item.get("du_dieu_kien_tinh_tong"):
            eligible_total += val
            if cat == "Có thể tiếp tục":
                continue_total += val
            elif cat == "Chờ duyệt":
                pending_total += val

    return {
        "counts": counts,
        "total_count": len(analyzed),
        "eligible_total_value": eligible_total,
        "continue_total_value": continue_total,
        "pending_total_value": pending_total
    }

def export_excel_bytes(analyzed: List[Dict[str, Any]], cleaned_df: pd.DataFrame) -> bytes:
    """Tạo file Excel nhiều tab (.xlsx): TongHop, ChiTietCoHoi, DuLieuLamSach."""
    metrics = build_summary_metrics(analyzed)
    counts = metrics["counts"]

    df_summary = pd.DataFrame([
        {"Nhóm Trạng thái Cơ hội": "1. Có thể tiếp tục", "Số lượng": counts["Có thể tiếp tục"], "Tổng giá trị sau VAT (VNĐ)": metrics["continue_total_value"]},
        {"Nhóm Trạng thái Cơ hội": "2. Chờ duyệt", "Số lượng": counts["Chờ duyệt"], "Tổng giá trị sau VAT (VNĐ)": metrics["pending_total_value"]},
        {"Nhóm Trạng thái Cơ hội": "3. Cần bổ sung thông tin", "Số lượng": counts["Cần bổ sung thông tin"], "Tổng giá trị sau VAT (VNĐ)": 0},
        {"Nhóm Trạng thái Cơ hội": "4. Chưa đáp ứng", "Số lượng": counts["Chưa đáp ứng"], "Tổng giá trị sau VAT (VNĐ)": 0},
        {"Nhóm Trạng thái Cơ hội": "TỔNG ĐỦ ĐIỀU KIỆN (Tiếp tục + Chờ duyệt)", "Số lượng": counts["Có thể tiếp tục"] + counts["Chờ duyệt"], "Tổng giá trị sau VAT (VNĐ)": metrics["eligible_total_value"]}
    ])

    detail_rows = []
    for it in analyzed:
        detail_rows.append({
            "Mã yêu cầu": it["id_yeu_cau"],
            "Mã cơ hội": it["ma_co_hoi"],
            "Khách hàng": it["khach_hang"],
            "NVKD phụ trách": it["nvkd_phu_trach"],
            "Phân loại": it["phan_loai"],
            "Đủ ĐK tính tổng GT": "CÓ" if it["du_dieu_kien_tinh_tong"] else "KHÔNG",
            "Tự xác nhận đơn": "KHÔNG",
            "Giá niêm yết (VNĐ)": it["gia_tri_niem_yet"],
            "Chiết khấu (%)": it["chiet_khau_pct"],
            "Giá trị sau VAT (VNĐ)": it["gia_tri_sau_vat"],
            "Lợi nhuận gộp (VNĐ)": it["loi_nhuan_gop"],
            "Gross Margin (%)": it["margin_pct"],
            "Cấp phê duyệt": it["cap_phe_duyet"],
            "Căn cứ đối chiếu": it["can_cu"],
            "Bước xử lý tiếp theo": it["buoc_tiep_theo"]
        })
    df_detail = pd.DataFrame(detail_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="TongHop", index=False)
        df_detail.to_excel(writer, sheet_name="ChiTietCoHoi", index=False)
        cleaned_df.to_excel(writer, sheet_name="DuLieuLamSach", index=False)

    return output.getvalue()

def export_documents_docx_bytes(analyzed: List[Dict[str, Any]], docs_map: Dict[str, Dict[str, str]], approver_name: str = "Ban Giám Đốc / Đại diện Phê duyệt", approval_notes: str = "") -> bytes:
    """Tạo file Word (.docx) chứa Báo cáo Kinh doanh Tổng hợp và trọn bộ Hồ sơ cơ hội, Email nháp, Phiếu duyệt chuẩn đẹp."""
    doc = Document()
    setup_document_styles(doc, header_text="CÔNG TY CỔ PHẦN CÔNG NGHỆ MINH AN • HỒ SƠ BÁN HÀNG B2B")

    title = doc.add_heading("CÔNG TY CỔ PHẦN CÔNG NGHỆ MINH AN", level=0)
    title.runs[0].font.color.rgb = RGBColor(12, 14, 8)
    sub = doc.add_heading("BÁO CÁO KINH DOANH & TỔNG KẾT CƠ HỘI BÁN HÀNG B2B", level=1)
    sub.runs[0].font.color.rgb = RGBColor(27, 54, 93)

    metrics = build_summary_metrics(analyzed)
    counts = metrics["counts"]

    p_meta = doc.add_paragraph(f"Thời gian lập: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')} • Người phê duyệt: {approver_name}")
    p_meta.runs[0].font.size = Pt(9)
    if approval_notes:
        p_note = doc.add_paragraph(f"Ý kiến phê duyệt: {approval_notes}")
        p_note.runs[0].font.italic = True
        p_note.runs[0].font.size = Pt(8.5)

    # 1. BẢNG TỔNG KẾT KINH DOANH
    doc.add_heading("I. TỔNG HỢP CHỈ SỐ KINH DOANH B2B", level=2)
    tbl_kpi = doc.add_table(rows=6, cols=3)
    kpi_headers = ["Nhóm Trạng thái Cơ hội", "Số lượng cơ hội", "Tổng giá trị sau VAT (VNĐ)"]
    for ci, h in enumerate(kpi_headers):
        tbl_kpi.cell(0, ci).paragraphs[0].add_run(h).bold = True

    kpi_rows = [
        ("1. Có thể tiếp tục (NVKD Tự quyết)", str(counts["Có thể tiếp tục"]), format_compact_vnd(metrics['continue_total_value'])),
        ("2. Chờ duyệt (Trình TPKD ngoại lệ)", str(counts["Chờ duyệt"]), format_compact_vnd(metrics['pending_total_value'])),
        ("3. Cần bổ sung thông tin (Hỏi lại KH)", str(counts["Cần bổ sung thông tin"]), "0 đ (Chưa rõ)"),
        ("4. Chưa đáp ứng (Rủi ro SLA / Tồn kho)", str(counts["Chưa đáp ứng"]), "0 đ (Từ chối khéo)"),
        ("TỔNG ĐỦ ĐIỀU KIỆN (Tiếp tục + Chờ duyệt)", str(counts["Có thể tiếp tục"] + counts["Chờ duyệt"]), format_compact_vnd(metrics['eligible_total_value']))
    ]
    for ri, rdata in enumerate(kpi_rows, start=1):
        tbl_kpi.cell(ri, 0).paragraphs[0].add_run(rdata[0])
        tbl_kpi.cell(ri, 1).paragraphs[0].add_run(rdata[1])
        r_val = tbl_kpi.cell(ri, 2).paragraphs[0].add_run(rdata[2])
        if ri == 5:
            r_val.bold = True
    apply_table_styles(tbl_kpi)

    # 2. BẢNG CHI TIẾT CÁC CƠ HỘI
    doc.add_heading("II. DANH MỤC CÁC CƠ HỘI BÁN HÀNG B2B", level=2)
    tbl_opp = doc.add_table(rows=1 + len(analyzed), cols=6)
    opp_hdrs = ["Mã YC", "Khách hàng", "NVKD", "Sau VAT", "CK / Margin", "Phân loại & Cấp duyệt"]
    for ci, h in enumerate(opp_hdrs):
        tbl_opp.cell(0, ci).paragraphs[0].add_run(h).bold = True
    for ri, it in enumerate(analyzed, start=1):
        tbl_opp.cell(ri, 0).paragraphs[0].add_run(str(it.get("id_yeu_cau", "")))
        tbl_opp.cell(ri, 1).paragraphs[0].add_run(str(it.get("khach_hang", "")))
        tbl_opp.cell(ri, 2).paragraphs[0].add_run(str(it.get("nvkd_phu_trach", "")))
        val_str = format_compact_vnd(it.get('gia_tri_sau_vat', 0)) if it.get("gia_tri_sau_vat") else "-"
        tbl_opp.cell(ri, 3).paragraphs[0].add_run(val_str)
        tbl_opp.cell(ri, 4).paragraphs[0].add_run(f"{it.get('chiet_khau_pct', 0)}% / {it.get('margin_pct', 0)}%")
        tbl_opp.cell(ri, 5).paragraphs[0].add_run(f"{it.get('phan_loai', '')}\n({it.get('cap_phe_duyet', '')})")
    apply_table_styles(tbl_opp)

    doc.add_page_break()

    # 3. CHI TIẾT TỪNG BỘ HỒ SƠ
    doc.add_heading("III. BỘ HỒ SƠ CHI TIẾT TỪNG CƠ HỘI & EMAIL NHÁP", level=2)
    for it in analyzed:
        req_id = it["id_yeu_cau"]
        d = docs_map.get(req_id, {})
        h = doc.add_heading(f"{req_id} — {it['khach_hang']} [{it['phan_loai']}]", level=3)
        h.runs[0].font.color.rgb = RGBColor(15, 76, 129)

        doc.add_heading("1. Hồ sơ Cơ hội", level=4)
        add_markdown_to_docx(doc, d.get("ho_so_co_hoi", ""))

        doc.add_heading("2. Bản nháp Email Phản hồi Khách hàng", level=4)
        add_markdown_to_docx(doc, d.get("email_nhap", ""))

        if d.get("phieu_chuyen_duyet"):
            doc.add_heading("3. Phiếu Chuyển Duyệt Ngoại lệ", level=4)
            add_markdown_to_docx(doc, d.get("phieu_chuyen_duyet", ""))

        doc.add_page_break()

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def export_documents_text(analyzed: List[Dict[str, Any]], docs_map: Dict[str, Dict[str, str]]) -> str:
    blocks = []
    for it in analyzed:
        req_id = it["id_yeu_cau"]
        d = docs_map.get(req_id, {})
        block = f"################################################################\n{req_id} - {it['khach_hang']} [{it['phan_loai']}]\n################################################################\n\n"
        block += d.get("ho_so_co_hoi", "") + "\n\n--- EMAIL NHÁP ---\n" + d.get("email_nhap", "") + "\n"
        if d.get("phieu_chuyen_duyet"):
            block += "\n--- PHIẾU CHUYỂN DUYỆT ---\n" + d.get("phieu_chuyen_duyet", "") + "\n"
        blocks.append(block)
    return "\n\n".join(blocks)

def export_audit_log_json(audit_log: List[Dict[str, Any]]) -> str:
    return json.dumps(audit_log, ensure_ascii=False, indent=2)

def export_audit_log_txt(audit_log: List[Dict[str, Any]]) -> str:
    lines = ["MINH AN B2B SALES AI AGENT — NHẬT KÝ TRUY VẾT (AUDIT LOG)", "=" * 65]
    for entry in audit_log:
        lines.append(f"[{entry.get('timestamp', '')}] [{entry.get('step', '')}] ({entry.get('tool', '')}) -> {entry.get('detail', '')}")
    return "\n".join(lines)

