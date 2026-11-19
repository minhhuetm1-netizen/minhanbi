"""
doc_formatter.py - Bộ công cụ chuẩn hóa định dạng Word (.docx) chuyên nghiệp:
1. Parser Markdown -> python-docx: Heading, Bold, Italic, List Bullet, Table, Blockquote.
2. Tuyệt đối không để sót ký tự thô: **, ###, ---, | :--- |.
3. Định dạng số tiền gọn gàng theo Tỷ VND / Triệu VND.
4. Tạo biểu đồ đường (Line Chart) xu hướng theo tháng nhúng trực tiếp vào Word.
5. Thiết lập bảng tinh gọn, viền nhẹ, màu nền header, lề trang chuẩn, header/footer Segoe UI.
"""

import io
import re
import datetime
from typing import List, Dict, Any, Optional

import pandas as pd
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ---------------------------------------------------------------------------
# 1. HÀM ĐỊNH DẠNG SỐ TIỀN & SỐ HỌC
# ---------------------------------------------------------------------------

def format_compact_vnd(val: Any, is_currency: bool = True) -> str:
    """
    Làm tròn và định dạng số tiền gọn gàng:
    Ví dụ: 449.597.290.000 -> 449,60 Tỷ VND
           57.860.102.000  -> 57,86 Tỷ VND
           29.973.152      -> 29,97 Triệu VND
    """
    if val is None or pd.isna(val):
        return "-"
    try:
        num = float(val)
    except (ValueError, TypeError):
        return str(val)

    unit = " VND" if is_currency else ""
    abs_num = abs(num)
    if abs_num >= 1_000_000_000:
        s = f"{num / 1_000_000_000:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{s} Tỷ{unit}"
    elif abs_num >= 1_000_000:
        s = f"{num / 1_000_000:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{s} Triệu{unit}"
    elif abs_num >= 1_000:
        s = f"{num:,.0f}".replace(",", ".")
        return f"{s}{unit}"
    return f"{num:,.0f}{unit}"

def format_detailed_vnd(val: Any, is_currency: bool = True) -> str:
    """Định dạng số chi tiết có dấu phân cách hàng nghìn cho bảng dữ liệu."""
    if val is None or pd.isna(val):
        return "-"
    try:
        num = float(val)
    except (ValueError, TypeError):
        return str(val)
    unit = " đ" if is_currency else ""
    return f"{num:,.0f}".replace(",", ".") + unit

# ---------------------------------------------------------------------------
# 2. HÀM THIẾT LẬP STYLE, LỀ & HEADER/FOOTER CHO DOCX
# ---------------------------------------------------------------------------

def setup_document_styles(doc: docx.Document, header_text: str = "BÁO CÁO ĐIỀU HÀNH KINH DOANH • HỆ THỐNG SALES BI AGENT"):
    """Cấu hình font thống nhất (Segoe UI), khoảng cách dòng 1.15, lề trang chuẩn A4 và header/footer."""
    # Style Normal
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Segoe UI"
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = RGBColor(12, 14, 8)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)

    # Set margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.8)

        # Header
        header = section.header
        p_head = header.paragraphs[0]
        p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_head.paragraph_format.space_after = Pt(0)
        p_head.text = ""
        r_head = p_head.add_run(header_text)
        r_head.font.name = "Segoe UI"
        r_head.font.size = Pt(8.5)
        r_head.font.color.rgb = RGBColor(110, 119, 129)

        # Footer
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.text = ""
        p_foot.paragraph_format.space_after = Pt(0)
        r_foot = p_foot.add_run("Tài liệu nội bộ — Bảo mật | Trang ")
        r_foot.font.name = "Segoe UI"
        r_foot.font.size = Pt(8.5)
        r_foot.font.color.rgb = RGBColor(110, 119, 129)

        # Page Number Field
        fldSimple = OxmlElement("w:fldSimple")
        fldSimple.set(qn("w:instr"), "PAGE")
        p_foot._p.append(fldSimple)

        r_div = p_foot.add_run(" / ")
        r_div.font.name = "Segoe UI"
        r_div.font.size = Pt(8.5)
        r_div.font.color.rgb = RGBColor(110, 119, 129)

        fldPages = OxmlElement("w:fldSimple")
        fldPages.set(qn("w:instr"), "NUMPAGES")
        p_foot._p.append(fldPages)

def set_cell_background(cell, hex_color: str):
    """Đặt màu nền cho cell trong table (hex_color không có dấu #)."""
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove existing shd if any
    for child in list(tcPr):
        if child.tag.endswith("shd"):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    """Đặt padding cho cell (dxa: 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def apply_table_styles(table: docx.table.Table, header_bg: str = "1B365D", alt_row_bg: str = "F8F9FA", border_color: str = "D0D7DE"):
    """Định dạng bảng Word hiện đại: Header navy/đậm, viền ngang nhẹ, không viền dọc, padding thoáng."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{border_color}"/>'
        f'<w:left w:val="none"/><w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    for ri, row in enumerate(table.rows):
        is_header = (ri == 0)
        bg = header_bg if is_header else (alt_row_bg if ri % 2 == 1 else "FFFFFF")
        for cell in row.cells:
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_after = Pt(1)
                for run in p.runs:
                    run.font.name = "Segoe UI"
                    if is_header:
                        run.font.size = Pt(9.5)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                    else:
                        run.font.size = Pt(9)

# ---------------------------------------------------------------------------
# 3. PARSER NỘI DUNG MARKDOWN SANG WORD CHUẨN
# ---------------------------------------------------------------------------

def _add_inline_runs(paragraph: docx.text.paragraph.Paragraph, text: str, default_size=Pt(10), default_color=None, font_name="Segoe UI"):
    """
    Phân tích inline markdown (***bold italic***, **bold**, *italic*, _italic_)
    và chèn vào paragraph dưới dạng Run chuẩn. Tuyệt đối không để sót dấu sao ** hay gạch dưới __.
    """
    # Regex bắt token markdown
    tokens = re.split(r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*|___.*?___|__.*?__|(?<!\w)_.*?_(?!\w))', text)
    for token in tokens:
        if not token:
            continue
        run = paragraph.add_run()
        run.font.name = font_name
        run.font.size = default_size
        if default_color:
            run.font.color.rgb = default_color

        if (token.startswith("***") and token.endswith("***")) or (token.startswith("___") and token.endswith("___")):
            run.text = token[3:-3]
            run.bold = True
            run.italic = True
        elif (token.startswith("**") and token.endswith("**")) or (token.startswith("__") and token.endswith("__")):
            run.text = token[2:-2]
            run.bold = True
        elif (token.startswith("*") and token.endswith("*")) or (token.startswith("_") and token.endswith("_")):
            run.text = token[1:-1]
            run.italic = True
        else:
            # Loại bỏ sạch sẽ bất kỳ ký tự markdown thô nào bị lọt lưới
            clean = token.replace("**", "").replace("###", "").replace("---", "").replace("|", "")
            run.text = clean

def add_markdown_to_docx(doc: docx.Document, md_text: str):
    """
    Parser chuyển đổi toàn bộ chuỗi Markdown sang định dạng Word (.docx) chuyên nghiệp:
    - #, ##, ###, #### -> Heading có màu Navy (#1B365D), cỡ chữ phân cấp rõ ràng
    - * bullet, - bullet -> Word List Bullet chuẩn có parse inline bold/italic
    - 1. numbered list -> Numbered paragraph
    - | ... | -> Table Word có viền nhẹ và tô màu header
    - ---, === -> Bỏ qua đường kẻ thô, thay bằng khoảng đệm
    - **bold** -> run.bold = True
    """
    if not md_text or not md_text.strip():
        return

    lines = md_text.splitlines()
    i = 0
    total = len(lines)

    while i < total:
        raw_line = lines[i]
        line = raw_line.strip()

        # 1. Dòng trống
        if not line:
            i += 1
            continue

        # 2. Dòng kẻ ngang (---, ***, ===, ___)
        if re.match(r'^[-*_=\s]{3,}$', line):
            i += 1
            continue

        # 3. Bảng Markdown (| col 1 | col 2 | ...)
        if line.startswith("|") and line.endswith("|"):
            table_lines = []
            while i < total and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                row_str = lines[i].strip()
                # Bỏ qua dòng phân cách | :--- | :--- |
                if not re.match(r'^\|[\s:\-+|]+\|$', row_str):
                    # Tách các cell
                    cells = [c.strip() for c in row_str.strip("|").split("|")]
                    table_lines.append(cells)
                i += 1

            if table_lines:
                num_cols = max(len(r) for r in table_lines)
                table = doc.add_table(rows=len(table_lines), cols=num_cols)
                for r_idx, row_cells in enumerate(table_lines):
                    for c_idx in range(num_cols):
                        cell_val = row_cells[c_idx] if c_idx < len(row_cells) else ""
                        p = table.cell(r_idx, c_idx).paragraphs[0]
                        p.text = ""
                        _add_inline_runs(p, cell_val, default_size=Pt(9) if r_idx > 0 else Pt(9.5))
                apply_table_styles(table)
                # Thêm khoảng đệm sau bảng
                p_spacer = doc.add_paragraph()
                p_spacer.paragraph_format.space_after = Pt(4)
            continue

        # 4. Heading (#, ##, ###, ####)
        m_head = re.match(r'^(#{1,6})\s+(.*)$', line)
        if m_head:
            level = len(m_head.group(1))
            h_text = m_head.group(2).strip()
            # Bỏ ký tự thô trong heading nếu có
            h_text = h_text.replace("**", "").replace("`", "")

            # Tạo paragraph styled as heading
            p_h = doc.add_paragraph()
            p_h.paragraph_format.keep_with_next = True
            r_h = p_h.add_run(h_text)
            r_h.font.name = "Segoe UI"
            r_h.font.bold = True

            if level == 1:
                p_h.paragraph_format.space_before = Pt(12)
                p_h.paragraph_format.space_after = Pt(4)
                r_h.font.size = Pt(14)
                r_h.font.color.rgb = RGBColor(27, 54, 93)  # Navy
            elif level == 2:
                p_h.paragraph_format.space_before = Pt(10)
                p_h.paragraph_format.space_after = Pt(3)
                r_h.font.size = Pt(12)
                r_h.font.color.rgb = RGBColor(27, 54, 93)  # Navy
            elif level == 3:
                p_h.paragraph_format.space_before = Pt(8)
                p_h.paragraph_format.space_after = Pt(2)
                r_h.font.size = Pt(11)
                r_h.font.color.rgb = RGBColor(12, 14, 8)
            else:
                p_h.paragraph_format.space_before = Pt(6)
                p_h.paragraph_format.space_after = Pt(2)
                r_h.font.size = Pt(10)
                r_h.font.color.rgb = RGBColor(40, 40, 40)

            i += 1
            continue

        # 5. Gạch đầu dòng Bullet (*, -, +)
        m_bullet = re.match(r'^(\s*)[*+-]\s+(.*)$', raw_line)
        if m_bullet:
            indent_spaces = len(m_bullet.group(1))
            b_text = m_bullet.group(2).strip()

            p_bullet = doc.add_paragraph(style="List Bullet")
            p_bullet.paragraph_format.space_after = Pt(3)
            p_bullet.paragraph_format.line_spacing = 1.15
            if indent_spaces >= 2:
                p_bullet.paragraph_format.left_indent = Inches(0.4)

            _add_inline_runs(p_bullet, b_text)
            i += 1
            continue

        # 6. Danh sách có số thứ tự (1. 2. 3.)
        m_num = re.match(r'^(\s*(\d+)\.)\s+(.*)$', raw_line)
        if m_num:
            prefix = m_num.group(1)
            n_text = m_num.group(3).strip()

            p_num = doc.add_paragraph()
            p_num.paragraph_format.space_after = Pt(3)
            p_num.paragraph_format.line_spacing = 1.15
            p_num.paragraph_format.left_indent = Inches(0.25)

            r_pre = p_num.add_run(prefix + " ")
            r_pre.font.name = "Segoe UI"
            r_pre.font.bold = True
            r_pre.font.size = Pt(10)

            _add_inline_runs(p_num, n_text)
            i += 1
            continue

        # 7. Blockquote (> quote)
        if line.startswith(">"):
            q_text = line.lstrip("> ").strip()
            p_q = doc.add_paragraph()
            p_q.paragraph_format.left_indent = Inches(0.35)
            p_q.paragraph_format.space_after = Pt(4)
            _add_inline_runs(p_q, q_text, default_color=RGBColor(80, 80, 80))
            for run in p_q.runs:
                run.italic = True
            i += 1
            continue

        # 8. Đoạn văn bản thông thường
        p_text = doc.add_paragraph()
        p_text.paragraph_format.line_spacing = 1.15
        p_text.paragraph_format.space_after = Pt(4)
        _add_inline_runs(p_text, line)
        i += 1

# ---------------------------------------------------------------------------
# 4. HÀM TẠO BIỂU ĐỒ ĐƯỜNG (LINE CHART) XU HƯỚNG THÁNG
# ---------------------------------------------------------------------------

def generate_monthly_line_chart(monthly_data: List[Dict[str, Any]]) -> Optional[bytes]:
    """
    Vẽ biểu đồ đường (Line Chart) xu hướng Doanh thu & Lợi nhuận theo Tháng (sắp xếp tăng dần thời gian).
    Tự động gắn cờ Đỉnh (Peak) và Đáy (Lowest) doanh thu.
    """
    if not monthly_data:
        return None
    try:
        # Sắp xếp thời gian tăng dần (vd: 2024-01 đến 2025-12)
        sorted_m = sorted(monthly_data, key=lambda x: str(x.get("group_value", "")))
        months = [str(x.get("group_value", "")) for x in sorted_m]
        if not months:
            return None

        # Quy đổi đơn vị: Nếu >= 1 tỷ thì quy về Tỷ VND, ngược lại Triệu VND
        max_rev = max([float(x.get("revenue", 0)) for x in sorted_m] or [0])
        scale = 1_000_000_000.0 if max_rev >= 1_000_000_000.0 else 1_000_000.0
        unit_lbl = "Tỷ VND" if scale >= 1_000_000_000.0 else "Triệu VND"

        revs = [float(x.get("revenue", 0)) / scale for x in sorted_m]
        profs = [float(x.get("profit", 0)) / scale if x.get("profit") is not None else 0.0 for x in sorted_m]

        plt.figure(figsize=(7.4, 2.9), dpi=180)
        ax = plt.gca()
        ax.set_facecolor("#FFFFFF")

        # Vẽ đường Doanh thu
        plt.plot(months, revs, marker="o", markersize=4.5, color="#0F62FE", linewidth=2.2, label=f"Doanh thu ({unit_lbl})")
        plt.fill_between(months, revs, color="#0F62FE", alpha=0.08)

        # Vẽ đường Lợi nhuận nếu có số liệu
        has_prof = any(p > 0 for p in profs)
        if has_prof:
            plt.plot(months, profs, marker="s", markersize=3.5, color="#0E8A16", linewidth=1.6, linestyle="--", label=f"Lợi nhuận ({unit_lbl})")

        # Điểm Đỉnh (Peak) và Điểm Đáy (Lowest)
        max_idx = revs.index(max(revs))
        min_idx = revs.index(min(revs))

        plt.annotate(
            f"Đỉnh: {revs[max_idx]:.2f} {unit_lbl}",
            xy=(months[max_idx], revs[max_idx]),
            xytext=(0, 9),
            textcoords="offset points",
            ha="center",
            fontsize=7.5,
            fontweight="bold",
            color="#D9381E",
            bbox=dict(boxstyle="round,pad=0.25", facecolor="#FFEEEE", edgecolor="#D9381E", alpha=0.95, lw=0.8)
        )

        plt.annotate(
            f"Đáy: {revs[min_idx]:.2f} {unit_lbl}",
            xy=(months[min_idx], revs[min_idx]),
            xytext=(0, -16),
            textcoords="offset points",
            ha="center",
            fontsize=7.5,
            fontweight="bold",
            color="#555555",
            bbox=dict(boxstyle="round,pad=0.25", facecolor="#F0F0F0", edgecolor="#888888", alpha=0.95, lw=0.8)
        )

        plt.title("XU HƯỚNG TĂNG TRƯỞNG DOANH THU & LỢI NHUẬN THEO THỜI GIAN", fontsize=10, fontweight="bold", pad=10, color="#1B365D")
        plt.xlabel("Kỳ báo cáo (Tháng)", fontsize=8, labelpad=5, color="#555555")
        plt.ylabel(f"Giá trị ({unit_lbl})", fontsize=8, labelpad=5, color="#555555")
        plt.xticks(rotation=45, ha="right", fontsize=7)
        plt.yticks(fontsize=7.5)
        plt.grid(True, linestyle=":", alpha=0.5, color="#D0D7DE")
        plt.legend(loc="upper left", fontsize=7.5, framealpha=0.9)
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=180, bbox_inches="tight")
        plt.close()
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        print("[Chart Error]", e)
        return None

