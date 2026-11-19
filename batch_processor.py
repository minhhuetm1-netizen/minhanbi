"""
Batch Processor: Xử lý hàng loạt 180 yêu cầu/tuần & Lập Báo cáo Điều hành Ban Giám đốc
Mô phỏng và xử lý toàn bộ dung lượng 1 tuần làm việc của 6 NVKD Công ty Minh An.
"""

import os
import random
import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from agent_engine import MinhAnAgentEngine

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

CUSTOMER_TEMPLATES = [
    ("Tập đoàn Công nghệ NextGen", "Anh Nguyễn Văn Tuấn - Giám đốc CNTT", "tuan.nv@nextgen.vn"),
    ("Công ty CP Vận tải Quốc tế Biển Đông", "Chị Phạm Thu Trang - P. Mua sắm", "trang.pt@biendong-logistics.com"),
    ("Bệnh viện Đa khoa Tâm Anh", "Anh Lê Hoàng Nam - Trưởng phòng IT", "nam.lh@tamanhhospital.vn"),
    ("Chuỗi Bán lẻ Thời trang Nova", "Chị Vũ Mai Lan - Giám đốc Vận hành", "lan.vm@novafashion.vn"),
    ("Trường Quốc tế Song ngữ Horizon", "Thầy Đỗ Minh Trí - Ban Quản trị", "tri.dm@horizon-edu.vn"),
    ("Công ty CP Xây dựng Hòa Bình Land", "Anh Hoàng Quốc Việt - P. Đấu thầu", "viet.hq@hoabinhland.vn"),
    ("Ngân hàng Số FinTech V-Pay", "Anh Trần Đức Thắng - Tech Lead", "thang.td@vpay.vn"),
    ("Nhà máy Sản xuất Cơ khí Chính xác Sakura", "Ông Sato Kenji / Chị Vy - Mua hàng", "vy.do@sakura-mfg.com.vn"),
    ("Khách sạn & Nghỉ dưỡng Sun Grand Resort", "Chị Đặng Thảo My - Trưởng bộ phận IT", "my.dt@sungrandresort.vn"),
    ("Công ty Dược phẩm BioPharm Việt Nam", "Anh Bùi Tấn Phát - P. Thiết bị", "phat.bt@biopharm.com.vn")
]

PRODUCT_POOLS = [
    {"sku": "LAP-DELL-5540", "name": "Laptop Dell Latitude 5540", "typical_qty": (3, 15)},
    {"sku": "LAP-LEN-T14", "name": "Laptop Lenovo ThinkPad T14", "typical_qty": (2, 10)},
    {"sku": "SRV-DELL-R750", "name": "Máy chủ Dell PowerEdge R750", "typical_qty": (1, 3)},
    {"sku": "SW-CISCO-C9200", "name": "Switch Cisco Catalyst C9200L", "typical_qty": (1, 6)},
    {"sku": "FW-FORTI-60F", "name": "Firewall Fortinet FortiGate 60F", "typical_qty": (1, 4)},
    {"sku": "SVC-SEC-1Y", "name": "Dịch vụ Cấu hình An toàn Mạng", "typical_qty": (1, 2)},
    {"sku": "SVC-MAINT-247", "name": "Gói Bảo trì Hệ thống IT SLA 24/7", "typical_qty": (1, 2)},
    {"sku": "LIC-M365-BUS", "name": "Bản quyền Microsoft 365 Business", "typical_qty": (10, 80)}
]

def generate_weekly_180_inquiries():
    """Tạo bộ dữ liệu mô phỏng đúng 180 yêu cầu trong 1 tuần (từ Thứ 2 đến Thứ 7)."""
    random.seed(42)
    inquiries = []
    base_date = datetime.date(2026, 11, 2) # Tuần đầu tháng 11/2026

    for i in range(1, 181):
        day_offset = (i - 1) // 30 # 30 yêu cầu/ngày trong 6 ngày
        req_date = base_date + datetime.timedelta(days=day_offset)
        cust_tmpl = random.choice(CUSTOMER_TEMPLATES)
        cust_name = f"{cust_tmpl[0]} (Chi nhánh {i % 12 + 1})"

        # Chọn 1-3 sản phẩm
        num_items = random.choices([1, 2, 3], weights=[0.5, 0.35, 0.15])[0]
        selected_prods = random.sample(PRODUCT_POOLS, num_items)
        items_list = []
        text_parts = []

        for p in selected_prods:
            qty = random.randint(p["typical_qty"][0], p["typical_qty"][1])
            items_list.append({"sku": p["sku"], "qty": qty})
            text_parts.append(f"{qty} {p['name']}")

        # Quyết định xem yêu cầu này có thiếu thông tin không (khoảng 15% thiếu thông tin)
        is_missing_info = (random.random() < 0.15)
        
        if is_missing_info:
            missing_type = random.choice(["budget", "deadline", "both"])
            budget = None if missing_type in ["budget", "both"] else random.randint(80, 500) * 1_000_000
            deadline = None if missing_type in ["deadline", "both"] else random.randint(3, 15)
            discount = None
            text = f"Chào Minh An, bên mình đang cần tham khảo gấp {', '.join(text_parts)}. Xin vui lòng gửi báo giá và tiến độ sớm nhất có thể."
        else:
            budget = random.randint(50, 600) * 1_000_000
            deadline = random.choice([3, 5, 7, 10, 14, 20])
            discount = random.choices([0.0, 3.0, 5.0, 8.0, 10.0, 15.0], weights=[0.3, 0.25, 0.25, 0.1, 0.06, 0.04])[0]
            disc_str = f" Đề nghị chiết khấu {discount}%." if discount > 0 else ""
            text = f"Kính gửi Minh An, chúng tôi cần triển khai {', '.join(text_parts)}. Ngân sách dự kiến khoảng {budget//1_000_000} triệu. Yêu cầu bàn giao trong {deadline} ngày.{disc_str}"

        inquiries.append({
            "id_yeu_cau": f"REQ-W45-{i:03d}",
            "ngay_nhan": req_date.strftime("%Y-%m-%d"),
            "khach_hang": cust_name,
            "nguoi_lien_he": cust_tmpl[1],
            "email_kh": f"contact{i}@" + cust_tmpl[2].split('@')[1],
            "sdt_kh": f"09{random.randint(10000000, 99999999)}",
            "noi_dung_yeu_cau": text,
            "ngan_sach_cong_bo": budget,
            "thoi_han_ngay": deadline,
            "chiet_khau_de_xuat_pct": discount
        })

    return pd.DataFrame(inquiries)

def process_batch_and_generate_report():
    """Chạy toàn bộ 180 yêu cầu qua AI Engine và tạo Báo cáo Ban Giám đốc."""
    engine = MinhAnAgentEngine()
    df_raw = generate_weekly_180_inquiries()
    
    # Lưu file raw 180 yêu cầu
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(STATIC_DIR, exist_ok=True)
    df_raw.to_csv(os.path.join(DATA_DIR, "180_yeu_cau_tuan.csv"), index=False, encoding="utf-8-sig")

    processed_rows = []
    
    total_inquiries = len(df_raw)
    valid_count = 0
    missing_count = 0
    total_pipeline_value = 0
    total_cogs_value = 0
    approval_counts = {"NVKD Tự quyết": 0, "Trưởng phòng Kinh doanh (TPKD)": 0, "Ban Giám đốc (BGĐ)": 0}
    rep_stats = {}
    stockout_risks = 0
    lead_time_risks = 0

    for rep in engine.sales_reps:
        rep_stats[rep["ho_ten"]] = {
            "ma_nv": rep["ma_nv"],
            "chuyen_mon": rep["chuyen_mon"],
            "so_yeu_cau": 0,
            "tong_gia_tri": 0,
            "tong_loi_nhuan": 0,
            "so_don_can_duyet": 0
        }

    for idx, row in df_raw.iterrows():
        row_dict = row.to_dict()
        parsed = engine.parse_raw_text(row_dict["noi_dung_yeu_cau"], row_dict)
        res = engine.run_full_pipeline(parsed)

        is_valid = res.get("quy_trinh_hoan_tat", False)
        if is_valid:
            valid_count += 1
            dossier = res["ho_so_co_hoi"]
            quote = res["bang_bao_gia"]
            appr = res["trinh_duyet"]
            stock = res["tham_dinh_ton_kho"]

            rep_name = dossier["nvkd_phu_trach"]
            grand_total = quote["tong_sau_vat"]
            profit = quote["loi_nhuan_gop"]

            total_pipeline_value += grand_total
            total_cogs_value += quote["tong_gia_von"]

            cap = appr["cap_phe_duyet"]
            if "Ban Giám đốc" in cap:
                approval_counts["Ban Giám đốc (BGĐ)"] += 1
            elif "Trưởng phòng" in cap:
                approval_counts["Trưởng phòng Kinh doanh (TPKD)"] += 1
            else:
                approval_counts["NVKD Tự quyết"] += 1

            if rep_name in rep_stats:
                rep_stats[rep_name]["so_yeu_cau"] += 1
                rep_stats[rep_name]["tong_gia_tri"] += grand_total
                rep_stats[rep_name]["tong_loi_nhuan"] += profit
                if appr["can_phe_duyet"]:
                    rep_stats[rep_name]["so_don_can_duyet"] += 1

            if stock.get("rui_ro_tien_do"):
                lead_time_risks += 1
            if not stock.get("chi_tiet", [{}])[0].get("du_hang", True):
                stockout_risks += 1

            processed_rows.append({
                "Mã yêu cầu": row_dict["id_yeu_cau"],
                "Ngày nhận": row_dict["ngay_nhan"],
                "Khách hàng": row_dict["khach_hang"],
                "Tình trạng thông tin": "ĐẦY ĐỦ",
                "Mã cơ hội": dossier["ma_co_hoi"],
                "NVKD phụ trách": rep_name,
                "Chuyên môn NVKD": dossier["chuyen_mon"],
                "Tổng giá niêm yết": quote["tong_gia_niem_yet"],
                "Chiết khấu %": quote["chiet_khau_pct"],
                "Tổng sau thuế (VAT)": grand_total,
                "Lợi nhuận gộp": profit,
                "Biên lãi gộp %": quote["ty_suat_loi_nhuan_gop_pct"],
                "Tồn kho & SLA": stock["tinh_trang_kho"],
                "Rủi ro tiến độ": "CÓ" if stock.get("rui_ro_tien_do") else "KHÔNG",
                "Cấp phê duyệt": appr["cap_phe_duyet"],
                "Trạng thái CRM": res["trang_thai_crm"]
            })
        else:
            missing_count += 1
            processed_rows.append({
                "Mã yêu cầu": row_dict["id_yeu_cau"],
                "Ngày nhận": row_dict["ngay_nhan"],
                "Khách hàng": row_dict["khach_hang"],
                "Tình trạng thông tin": "THIẾU THÔNG TIN",
                "Mã cơ hội": "-",
                "NVKD phụ trách": "Chờ xác nhận",
                "Chuyên môn NVKD": "-",
                "Tổng giá niêm yết": 0,
                "Chiết khấu %": 0,
                "Tổng sau thuế (VAT)": 0,
                "Lợi nhuận gộp": 0,
                "Biên lãi gộp %": 0,
                "Tồn kho & SLA": "-",
                "Rủi ro tiến độ": "-",
                "Cấp phê duyệt": "Chưa lập báo giá",
                "Trạng thái CRM": "CHO_LAM_RO"
            })

    # Tạo Workbook Báo cáo Ban Giám đốc chuyên nghiệp bằng openpyxl
    wb = Workbook()
    
    # ---------------------------------------------------------
    # SHEET 1: TỔNG QUAN ĐIỀU HÀNH (KPI TUẦN)
    # ---------------------------------------------------------
    ws_kpi = wb.active
    ws_kpi.title = "BaoCao_DieuHanh_Tuan"
    ws_kpi.views.sheetView[0].showGridLines = True

    # Styling helpers
    navy_fill = PatternFill(start_color="0F4C81", end_color="0F4C81", fill_type="solid")
    blue_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    gray_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    white_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    bold_title = Font(name="Arial", size=14, bold=True, color="0F4C81")
    section_font = Font(name="Arial", size=11, bold=True, color="1E3A8A")
    regular_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", size=10, bold=True)
    thin_border = Border(left=Side(style='thin', color='CBD5E1'),
                         right=Side(style='thin', color='CBD5E1'),
                         top=Side(style='thin', color='CBD5E1'),
                         bottom=Side(style='thin', color='CBD5E1'))

    # Tiêu đề
    ws_kpi["A1"] = "CÔNG TY CỔ PHẦN CÔNG NGHỆ MINH AN - BÁO CÁO ĐIỀU HÀNH BÁN HÀNG B2B TUẦN"
    ws_kpi["A1"].font = bold_title
    ws_kpi["A2"] = f"Dung lượng xử lý: 180 Yêu cầu tiếp nhận | Khung thời gian: Tuần 45/2026 | Ngày lập: {datetime.datetime.now().strftime('%d/%m/%Y')}"
    ws_kpi["A2"].font = Font(name="Arial", size=9, italic=True, color="64748B")

    # Bảng 1: Chỉ số KPI Tổng quan
    ws_kpi["A4"] = "I. CHỈ SỐ HOẠT ĐỘNG TOÀN DIỆN (EXECUTIVE KPIS)"
    ws_kpi["A4"].font = section_font

    kpi_data = [
        ("Tổng số yêu cầu tiếp nhận", f"{total_inquiries} yêu cầu", "100% qua luồng AI Agent"),
        ("Tỷ lệ hồ sơ đủ điều kiện lập báo giá ngay", f"{valid_count} ({valid_count/total_inquiries*100:.1f}%)", "Tự động phân bổ NVKD & kích hoạt SLA 4h"),
        ("Số yêu cầu thiếu thông tin (Đã tự động gửi email hỏi lại)", f"{missing_count} ({missing_count/total_inquiries*100:.1f}%)", "Gắn trạng thái [CHO_LAM_RO], không bị bỏ sót"),
        ("Tổng giá trị phễu cơ hội (Pipeline Value)", f"{total_pipeline_value:,.0f} VNĐ", "Tính theo bảng giá Master V4.2 đã gồm VAT"),
        ("Ước tính Lợi nhuận gộp (Estimated Gross Margin)", f"{total_pipeline_value - total_cogs_value:,.0f} VNĐ", f"Tỷ suất: {(total_pipeline_value - total_cogs_value)/total_pipeline_value*100:.1f}%"),
        ("Cảnh báo rủi ro xung đột tiến độ giao hàng (Lead-time conflict)", f"{lead_time_risks} vụ việc", "Đã tự động đề xuất phân kỳ giao hàng 2 đợt"),
        ("Số đơn hàng cần Trưởng phòng / Ban Giám đốc duyệt", f"{approval_counts['Trưởng phòng Kinh doanh (TPKD)'] + approval_counts['Ban Giám đốc (BGĐ)']} đơn", "Kiểm soát chặt thẩm quyền chiết khấu & giá trị")
    ]

    ws_kpi["A5"] = "Chỉ tiêu Quản trị & Hiệu suất"
    ws_kpi["B5"] = "Kết quả Đạt được"
    ws_kpi["C5"] = "Ghi chú Giám sát của AI Agent"
    for col in ["A5", "B5", "C5"]:
        ws_kpi[col].fill = navy_fill
        ws_kpi[col].font = white_font
        ws_kpi[col].alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, (k, v, note) in enumerate(kpi_data, start=6):
        ws_kpi[f"A{r_idx}"] = k
        ws_kpi[f"B{r_idx}"] = v
        ws_kpi[f"C{r_idx}"] = note
        ws_kpi[f"A{r_idx}"].font = bold_font
        ws_kpi[f"B{r_idx}"].font = bold_font
        ws_kpi[f"B{r_idx}"].alignment = Alignment(horizontal="right")
        ws_kpi[f"C{r_idx}"].font = regular_font
        for c in ["A", "B", "C"]:
            ws_kpi[f"{c}{r_idx}"].border = thin_border

    # Bảng 2: Phân bổ Tải công việc 6 NVKD
    start_r = 15
    ws_kpi[f"A{start_r}"] = "II. PHÂN BỔ TẢI CÔNG VIỆC & DOANH SỐ PHỄU GIỮA 6 NHÂN VIÊN KINH DOANH"
    ws_kpi[f"A{start_r}"].font = section_font

    headers_rep = ["Mã NV", "Họ và tên NVKD", "Mảng Chuyên môn", "Số Yêu cầu Gắn trách nhiệm", "Tổng Doanh số Phễu (VNĐ)", "Lợi nhuận gộp (VNĐ)", "Đơn cần Duyệt vượt quyền"]
    for c_idx, h in enumerate(headers_rep, start=1):
        cell = ws_kpi.cell(row=start_r+1, column=c_idx, value=h)
        cell.fill = blue_header_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, (name, s) in enumerate(rep_stats.items(), start=start_r+2):
        ws_kpi[f"A{r_idx}"] = s["ma_nv"]
        ws_kpi[f"B{r_idx}"] = name
        ws_kpi[f"C{r_idx}"] = s["chuyen_mon"]
        ws_kpi[f"D{r_idx}"] = s["so_yeu_cau"]
        ws_kpi[f"E{r_idx}"] = s["tong_gia_tri"]
        ws_kpi[f"F{r_idx}"] = s["tong_loi_nhuan"]
        ws_kpi[f"G{r_idx}"] = s["so_don_can_duyet"]

        ws_kpi[f"A{r_idx}"].alignment = Alignment(horizontal="center")
        ws_kpi[f"D{r_idx}"].alignment = Alignment(horizontal="center")
        ws_kpi[f"G{r_idx}"].alignment = Alignment(horizontal="center")
        ws_kpi[f"E{r_idx}"].number_format = '#,##0'
        ws_kpi[f"F{r_idx}"].number_format = '#,##0'

        for c_idx in range(1, 8):
            ws_kpi.cell(row=r_idx, column=c_idx).border = thin_border
            ws_kpi.cell(row=r_idx, column=c_idx).font = regular_font

    # Bảng 3: Ma trận Phân quyền Duyệt
    start_appr = start_r + len(rep_stats) + 3
    ws_kpi[f"A{start_appr}"] = "III. BÁO CÁO KIỂM SOÁT THẨM QUYỀN PHÊ DUYỆT (APPROVAL MATRIX)"
    ws_kpi[f"A{start_appr}"].font = section_font

    appr_headers = ["Cấp Phê duyệt", "Số lượng Đơn", "Tỷ lệ %", "Quy định Phân quyền Thẩm quyền"]
    for c_idx, h in enumerate(appr_headers, start=1):
        cell = ws_kpi.cell(row=start_appr+1, column=c_idx, value=h)
        cell.fill = navy_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    appr_matrix = [
        ("NVKD Tự quyết", approval_counts["NVKD Tự quyết"], f"{approval_counts['NVKD Tự quyết']/valid_count*100:.1f}%", "Chiết khấu <= 5.0%, Giá trị <= 200M, Hàng có sẵn"),
        ("Trưởng phòng Kinh doanh (TPKD)", approval_counts["Trưởng phòng Kinh doanh (TPKD)"], f"{approval_counts['Trưởng phòng Kinh doanh (TPKD)']/valid_count*100:.1f}%", "Chiết khấu 5.1% - 12%, Giá trị <= 800M hoặc Bảo lãnh rút ngắn tiến độ"),
        ("Ban Giám đốc (BGĐ)", approval_counts["Ban Giám đốc (BGĐ)"], f"{approval_counts['Ban Giám đốc (BGĐ)']/valid_count*100:.1f}%", "Chiết khấu > 12%, Giá trị > 800M hoặc Gross Margin < 8%")
    ]

    for r_idx, (role, cnt, pct, rule) in enumerate(appr_matrix, start=start_appr+2):
        ws_kpi[f"A{r_idx}"] = role
        ws_kpi[f"B{r_idx}"] = cnt
        ws_kpi[f"C{r_idx}"] = pct
        ws_kpi[f"D{r_idx}"] = rule
        ws_kpi[f"B{r_idx}"].alignment = Alignment(horizontal="center")
        ws_kpi[f"C{r_idx}"].alignment = Alignment(horizontal="center")
        for c_idx in range(1, 5):
            ws_kpi.cell(row=r_idx, column=c_idx).border = thin_border
            ws_kpi.cell(row=r_idx, column=c_idx).font = regular_font

    # Căn chỉnh độ rộng cột
    for col in ws_kpi.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_kpi.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # ---------------------------------------------------------
    # SHEET 2: CHI TIẾT TOÀN BỘ 180 CƠ HỘI
    # ---------------------------------------------------------
    ws_det = wb.create_sheet(title="ChiTiet_180_YeuCau")
    ws_det.views.sheetView[0].showGridLines = True

    df_processed = pd.DataFrame(processed_rows)
    headers_det = list(df_processed.columns)

    for c_idx, h in enumerate(headers_det, start=1):
        cell = ws_det.cell(row=1, column=c_idx, value=h)
        cell.fill = navy_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row in df_processed.iterrows():
        excel_row = r_idx + 2
        for c_idx, col_name in enumerate(headers_det, start=1):
            val = row[col_name]
            cell = ws_det.cell(row=excel_row, column=c_idx, value=val)
            cell.border = thin_border
            cell.font = regular_font
            if isinstance(val, (int, float)) and val > 1000:
                cell.number_format = '#,##0'

    for col in ws_det.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_det.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    # Lưu file báo cáo
    excel_report_path = os.path.join(STATIC_DIR, "Bao_Cao_Tuan_Ban_Giam_Doc.xlsx")
    wb.save(excel_report_path)
    wb.save(os.path.join(DATA_DIR, "Bao_Cao_Tuan_Ban_Giam_Doc.xlsx"))

    print(f"✅ Đã xử lý thành công 180 yêu cầu/tuần và tạo báo cáo: {excel_report_path}")
    return {
        "total": total_inquiries,
        "valid": valid_count,
        "missing": missing_count,
        "pipeline_value": total_pipeline_value,
        "gross_profit": total_pipeline_value - total_cogs_value,
        "lead_time_risks": lead_time_risks,
        "approval_counts": approval_counts,
        "rep_stats": rep_stats,
        "report_file": "/Bao_Cao_Tuan_Ban_Giam_Doc.xlsx"
    }

if __name__ == "__main__":
    process_batch_and_generate_report()

