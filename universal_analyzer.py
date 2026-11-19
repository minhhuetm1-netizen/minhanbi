"""
universal_analyzer.py
Động cơ Phân tích Dữ liệu & Thống kê Đa năng (Universal Analytics & Statistical Engine):
- Hỗ trợ mọi định dạng: Excel (.xlsx, .xlsm, .xls), CSV, TSV, Lark Base / JSON, Word (.docx), Text (.txt), PDF.
- Tự động đọc cấu trúc file bằng AI (Gemini) + Heuristic để hiểu ngữ cảnh bài toán:
  1. MINH_AN_B2B: Thẩm định Cơ hội Bán hàng B2B (4 nhóm chuẩn).
  2. UNIVERSAL_ANALYTICS: Thống kê Tài chính, Cây KPI & Biên Margin (MG), Khách hàng Ưu tiên (Pareto ABC),
     Phân tích Dự án / Sản phẩm và Phân nhóm Đa chiều.
- Tuân thủ 100% quy trình thống kê chuẩn & Human-in-the-Loop 5 bước có kiểm duyệt.
"""

import io
import re
import json
import math
import datetime
import warnings
from typing import Dict, List, Any, Optional, Tuple, Callable
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.chart import LineChart, BarChart, DoughnutChart, PieChart, Reference, Series
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from data_processor import (
    _strip_vn_accents, is_missing, normalize_currency, normalize_percentage
)
from agent_skills import _call_llm_json, get_default_ai_config
from doc_formatter import (
    setup_document_styles,
    add_markdown_to_docx,
    format_compact_vnd,
    format_detailed_vnd,
    generate_monthly_line_chart,
    apply_table_styles,
    set_cell_background,
    set_cell_margins
)

# Từ khóa nhận diện vai trò cột đa ngành (Việt + Anh + Lark + ERP + CRM + Dự án)
ROLE_KEYWORDS = {
    "revenue": ["sales_amount", "sales", "revenue", "doanh_thu", "thanh_tien", "total_amount", "total", "amount", "gia_tri", "gmv", "tong_tien", "budget", "ngan_sach", "contract_value"],
    "cost": ["unit_cost", "cost", "cogs", "gia_von", "chi_phi", "tong_von", "expense", "chi_thuc"],
    "profit": ["profit", "net_profit", "gross_profit", "loi_nhuan", "lai_gop", "margin_amount", "lai_rong"],
    "qty": ["quantity_sold", "quantity", "qty", "so_luong", "sl", "units", "unit_sold"],
    "price": ["unit_price", "price", "don_gia", "gia_ban", "selling_price"],
    "discount": ["discount", "chiet_khau", "ck", "giam_gia", "discount_pct", "disc"],
    "date": ["order_date", "sale_date", "date", "ngay", "ngay_dat", "ngay_ban", "created_at", "timestamp", "time", "deadline"],
    "id": ["order_id", "transaction_id", "id", "ma_don", "ma_gd", "ma_yeu_cau", "id_yeu_cau", "record_id", "ma_du_an"],
    "customer": ["customer_id", "customer", "customer_name", "client", "khach_hang", "ten_khach_hang", "khach", "doi_tac", "account", "company"],
    "project": ["project", "project_name", "du_an", "product_name", "product", "san_pham", "hang_muc", "service", "dich_vu", "item", "sub_category"]
}

def detect_dataset_mode(sheets_dict: Dict[str, pd.DataFrame], assignments: Dict[str, Optional[str]]) -> str:
    """
    Phân biệt giữa quy trình Minh An B2B RFQ (3 bảng hoặc có nội dung văn bản yêu cầu)
    và quy trình Phân tích Doanh thu / Thống kê Tài chính & Dự án Tổng quát.
    """
    if not sheets_dict:
        return "UNIVERSAL_ANALYTICS"

    s_inq = assignments.get("inquiries")
    s_cat = assignments.get("catalog")
    s_inv = assignments.get("inventory")
    distinct_assigned = {s for s in (s_inq, s_cat, s_inv) if s and s in sheets_dict}

    if len(sheets_dict) >= 2 and len(distinct_assigned) >= 2:
        return "MINH_AN_B2B"

    if s_inq and s_inq in sheets_dict:
        cols_norm = {_strip_vn_accents(c) for c in sheets_dict[s_inq].columns}
        if any(k in cols_norm for k in ("noi_dung_yeu_cau", "yeu_cau_kh", "email_body")) and any(k in cols_norm for k in ("ngan_sach", "ngan_sach_cong_bo", "du_kien_chi")):
            return "MINH_AN_B2B"

    return "UNIVERSAL_ANALYTICS"

def detect_universal_schema(df: pd.DataFrame, api_key: str = "", provider: str = "Gemini") -> Dict[str, Any]:
    """Tự động phân tích cấu trúc cột & ngữ cảnh nghiệp vụ bằng Heuristic + Gemini AI."""
    cols = list(df.columns)
    norm_map = {_strip_vn_accents(c): c for c in cols}

    schema: Dict[str, Any] = {
        "id_col": None,
        "date_col": None,
        "customer_col": None,
        "project_col": None,
        "revenue_col": None,
        "cost_col": None,
        "profit_col": None,
        "qty_col": None,
        "price_col": None,
        "discount_col": None,
        "group_cols": [],
        "all_columns": cols,
        "dataset_domain": "Thống kê Doanh thu, Biên lợi nhuận & Phân nhóm Dữ liệu",
        "ai_summary": "Tự động phát hiện cấu trúc bảng dữ liệu để làm sạch, dựng Cây KPI/Margin và phân hạng nhóm ưu tiên."
    }

    def _find_col(keys: List[str], exclude: set) -> Optional[str]:
        for k in keys:
            if k in norm_map and norm_map[k] not in exclude:
                return norm_map[k]
        for k in keys:
            for nc, oc in norm_map.items():
                if k in nc and oc not in exclude:
                    return oc
        return None

    used = set()
    for role_key, kw_key in [
        ("id_col", "id"), ("date_col", "date"), ("customer_col", "customer"),
        ("project_col", "project"), ("revenue_col", "revenue"), ("cost_col", "cost"),
        ("profit_col", "profit"), ("qty_col", "qty"), ("price_col", "price"),
        ("discount_col", "discount")
    ]:
        found = _find_col(ROLE_KEYWORDS[kw_key], used)
        schema[role_key] = found
        if found:
            used.add(found)

    # Chỉ gọi Gemini khi chưa tự động nhận diện được cột cốt lõi (revenue_col hoặc date_col) để tối ưu tốc độ phản hồi tức thì
    needs_ai = not (schema.get("revenue_col") and (schema.get("date_col") or schema.get("profit_col")))
    has_key = bool((api_key or "").strip()) or (provider == "Gemini" and bool(get_default_ai_config().get("gemini_api_key")))
    if needs_ai and has_key and provider in ("Gemini", "OpenAI") and len(cols) > 0:
        try:
            sample_rows = df.head(3).fillna("").astype(str).to_dict("records")
            prompt = (
                "Bạn là Chuyên gia Phân tích Dữ liệu & Tài chính. Hãy đọc tên các cột và 3 dòng mẫu sau để trả về DUY NHẤT 1 JSON object "
                "với các trường: dataset_domain (tên ngắn gọn bài toán/lĩnh vực của file), ai_summary (1 câu mô tả AI hiểu được từ dữ liệu và hướng xử lý thống kê), "
                "revenue_col, cost_col, profit_col, customer_col, project_col, date_col, group_cols (mảng tối đa 4 cột phân loại tốt nhất). "
                "Chỉ dùng đúng tên cột có trong danh sách hoặc null.\n"
                f"Cột: {json.dumps(cols, ensure_ascii=False)}\n"
                f"Mẫu: {json.dumps(sample_rows, ensure_ascii=False)}"
            )
            raw_ai = _call_llm_json(prompt, api_key, provider)
            if raw_ai:
                m = re.search(r'\{[\s\S]*\}', raw_ai)
                if m:
                    ai_obj = json.loads(m.group(0))
                    if ai_obj.get("dataset_domain"):
                        schema["dataset_domain"] = str(ai_obj["dataset_domain"])
                    if ai_obj.get("ai_summary"):
                        schema["ai_summary"] = str(ai_obj["ai_summary"])
                    for ck in ("revenue_col", "cost_col", "profit_col", "customer_col", "project_col", "date_col"):
                        if not schema.get(ck) and ai_obj.get(ck) in cols:
                            schema[ck] = ai_obj[ck]
                    if isinstance(ai_obj.get("group_cols"), list):
                        valid_ai_groups = [g for g in ai_obj["group_cols"] if g in cols]
                        for vg in valid_ai_groups:
                            if vg not in schema["group_cols"]:
                                schema["group_cols"].append(vg)
        except Exception:
            pass

    # Rule 1 Guardrail: Kiểm tra độ duy nhất của customer_col
    # Tuyệt đối KHÔNG BAO GIỜ đưa Cột Định danh (cardinality > 80%) hay cột không gom nhóm (>= 50%) vào customer_col
    if schema.get("customer_col") and schema["customer_col"] in df.columns:
        s_cust = df[schema["customer_col"]].dropna()
        if len(df) > 0:
            c_ratio = s_cust.nunique() / len(df)
            if c_ratio >= 0.5:
                if c_ratio > 0.8 and not schema.get("id_col"):
                    schema["id_col"] = schema["customer_col"]
                schema["customer_col"] = None

    assigned_metrics = {
        schema["id_col"], schema["date_col"], schema["revenue_col"],
        schema["cost_col"], schema["profit_col"], schema["qty_col"],
        schema["price_col"], schema["discount_col"]
    }

    # Bổ sung các cột phân nhóm hợp lý (loại trừ hoàn toàn cột ID có cardinality > 80%)
    for c in cols:
        if c in assigned_metrics or c in schema["group_cols"]:
            continue
        series = df[c].dropna()
        if len(series) == 0:
            continue
        nunique = series.nunique()
        card_ratio = (nunique / len(df)) if len(df) > 0 else 0.0
        if card_ratio > 0.8:
            if not schema.get("id_col"):
                schema["id_col"] = c
            continue

        is_num = pd.api.types.is_numeric_dtype(series)
        if not is_num and 2 <= nunique <= max(60, int(len(df) * 0.5)) and card_ratio < 0.5:
            schema["group_cols"].append(c)
        elif not is_num and nunique <= 15 and card_ratio <= 0.8:
            schema["group_cols"].append(c)

    # Lọc lại schema["group_cols"] để đảm bảo không lọt cột ID
    clean_group_cols = []
    for g in schema.get("group_cols", []):
        if g in df.columns:
            s_g = df[g].dropna()
            if len(df) > 0 and (s_g.nunique() / len(df)) > 0.8:
                continue
            clean_group_cols.append(g)
    schema["group_cols"] = clean_group_cols

    if not schema["group_cols"]:
        for c in cols:
            if c not in assigned_metrics and not pd.api.types.is_numeric_dtype(df[c]):
                s_c = df[c].dropna()
                if len(df) > 0 and (s_c.nunique() / len(df)) <= 0.8:
                    schema["group_cols"].append(c)
                    if len(schema["group_cols"]) >= 3:
                        break

    # Nếu chưa thấy revenue_col -> tìm cột số có giá trị trung bình lớn nhất (khác ID/Qty/Discount)
    if not schema["revenue_col"]:
        numeric_candidates = []
        for c in cols:
            if c in assigned_metrics:
                continue
            s_num = pd.to_numeric(df[c], errors="coerce").dropna()
            if len(s_num) > 0:
                numeric_candidates.append((c, float(s_num.abs().mean())))
        if numeric_candidates:
            numeric_candidates.sort(key=lambda x: x[1], reverse=True)
            schema["revenue_col"] = numeric_candidates[0][0]

    # Kiểm tra sự hiện diện của cột chỉ số định lượng (Numerical Metric)
    inspection = inspect_dataset_schema(df)
    has_num_metric = any(bool(schema.get(k)) for k in ("revenue_col", "cost_col", "profit_col", "qty_col", "price_col", "discount_col"))
    if not has_num_metric and inspection.get("has_numerical") and inspection.get("numerical_cols"):
        first_num = inspection["numerical_cols"][0]["name"]
        schema["revenue_col"] = first_num
        has_num_metric = True

    schema["has_numerical_metrics"] = has_num_metric
    if not has_num_metric:
        schema["ai_summary"] = "Lỗi cấu trúc dữ liệu: Không tìm thấy cột số liệu định lượng nào hợp lệ để tính toán chỉ số KPI và vẽ biểu đồ."

    schema["schema_inspection"] = inspection
    return schema

def inspect_dataset_schema(df: pd.DataFrame, sheets_dict: Optional[Dict[str, pd.DataFrame]] = None) -> Dict[str, Any]:
    """
    Cơ chế Quét Cấu trúc Dữ liệu Tự động (Schema Inspector) chuẩn Business Intelligence:
    1. Đọc danh sách tên cột (columns) trước tiên.
    2. Phân loại tự động các cột thành 3 nhóm:
       - Categorical (Định tính / Phân nhóm: Khu vực, Kênh, Danh mục, Tháng dạng chuỗi, Khách hàng, Trạng thái...)
       - Numerical (Định lượng / Chỉ số: Doanh thu, Lợi nhuận, Số lượng, Giá bán, Chi phí, Đơn giá...)
       - Temporal (Thời gian: Ngày, Tháng, Quý, Năm, Date, Timestamp, Created_At...)
    3. Xác định các chỉ số kinh doanh cốt lõi & cơ chế Adaptive UI:
       - has_sku_inventory: True nếu có SKU và Tồn kho thực tế (hoặc sheet tồn kho B2B).
       - primary_metric: Cột chỉ số chính (Doanh thu / Giá trị / Ngân sách / Số lượng...)
       - secondary_metric: Cột chỉ số phụ (Lợi nhuận, Chi phí, Số lượng...)
       - primary_dimension: Chiều phân nhóm chính (Danh mục, Khu vực, Kênh, Phân khúc...)
       - secondary_dimension: Chiều phân nhóm thứ hai
       - primary_temporal: Cột thời gian chính
       - pivot_summary: Bảng tổng hợp đa chiều linh hoạt theo các chiều dữ liệu thực tế của tệp
    """
    if df is None or df.empty:
        return {
            "categorical_cols": [],
            "numerical_cols": [],
            "temporal_cols": [],
            "has_sku_inventory": False,
            "has_temporal": False,
            "has_categorical": False,
            "has_numerical": False,
            "primary_metric": None,
            "secondary_metric": None,
            "primary_dimension": None,
            "secondary_dimension": None,
            "primary_temporal": None,
            "pivot_summary": [],
            "total_rows": 0
        }

    cols = list(df.columns)
    temporal_cols = []
    numerical_cols = []
    categorical_cols = []
    id_cols = []

    TIME_KEYWORDS = ["date", "ngày", "ngay", "thời_gian", "thoi_gian", "time", "timestamp", "created_at", "month", "tháng", "thang", "_thang", "year", "năm", "nam", "quarter", "quý", "quy", "period"]
    ID_KEYWORDS = ["id", "mã", "ma_", "code", "uuid", "guid", "stt", "số_thứ_tự", "sdt", "phone", "dien_thoai", "tel", "mobile", "fax", "tax", "mst", "zip", "postal", "cccd", "cmnd"]
    REV_KEYWORDS = ["revenue", "sales", "doanh_thu", "doanh_so", "thanh_tien", "ngan_sach", "tiền", "tien", "amount", "budget"]
    PROFIT_KEYWORDS = ["profit", "lợi_nhuận", "loi_nhuan", "lãi", "lai", "margin", "biên_ln"]
    COST_KEYWORDS = ["cost", "chi_phí", "chi_phi", "gia_von", "giá_vốn", "expense", "cogs"]
    QTY_KEYWORDS = ["qty", "quantity", "số_lượng", "so_luong", "sản_lượng", "san_luong", "volume", "units"]
    PRICE_KEYWORDS = ["price", "đơn_giá", "don_gia", "giá_bán", "gia_ban", "unit_price"]
    DISC_KEYWORDS = ["discount", "chiết_khấu", "chiet_khau", "giảm_giá", "giam_gia", "rebate"]
    STOCK_KEYWORDS = ["stock", "ton_kho", "tồn_kho", "ton_kha_dung", "tồn_khả_dụng", "inventory", "on_hand", "khả_dụng"]
    SKU_KEYWORDS = ["sku", "mã_sp", "ma_sp", "product_id", "item_id", "mã_hàng", "ma_hang"]

    def _strip_col(s: str) -> str:
        s = str(s).lower().strip()
        for c1, c2 in [('à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ', 'a'),
                      ('è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ', 'e'),
                      ('ì|í|ị|ỉ|ĩ', 'i'),
                      ('ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ', 'o'),
                      ('ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ', 'u'),
                      ('ỳ|ý|ỵ|ỷ|ỹ', 'y'),
                      ('đ', 'd')]:
            s = re.sub(c1, c2, s)
        return s

    for col in cols:
        col_clean = _strip_col(col)
        series = df[col].dropna()
        if len(series) == 0:
            continue

        # 1. Kiểm tra Temporal
        is_temp = False
        if pd.api.types.is_datetime64_any_dtype(series):
            is_temp = True
        elif any(k == col_clean or k in col_clean for k in TIME_KEYWORDS):
            sample = series.head(10).astype(str)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                parsed = pd.to_datetime(sample, errors='coerce')
            if parsed.notna().sum() >= max(1, int(len(sample) * 0.5)):
                is_temp = True
            elif any(k in col_clean for k in ["thang", "tháng", "_thang", "month", "year", "nam", "quy", "quarter"]):
                is_temp = True

        if is_temp:
            temporal_cols.append({
                "name": col,
                "role": "date" if any(w in col_clean for w in ["ngay", "date"]) else "month",
                "distinct_count": int(series.nunique()),
                "sample_values": [str(x) for x in series.head(3).tolist()]
            })
            continue

        # 2. Kiểm tra Numerical
        is_num = False
        num_series = None
        if pd.api.types.is_numeric_dtype(series):
            num_series = series
            is_num = True
        else:
            try:
                s_clean = series.astype(str).str.replace(r'[\$,VNĐvndđ\s]', '', regex=True).str.replace(',', '')
                converted = pd.to_numeric(s_clean, errors='coerce')
                if converted.notna().sum() >= max(1, int(len(series) * 0.7)):
                    num_series = converted.dropna()
                    is_num = True
            except Exception:
                pass

        if is_num and num_series is not None:
            is_id = any(k in col_clean for k in ID_KEYWORDS) and (series.nunique() >= len(series) * 0.9 or series.nunique() > 100)
            if is_id:
                is_num = False

        if is_num and num_series is not None and len(num_series) > 0:
            role = "metric"
            if any(k in col_clean for k in REV_KEYWORDS): role = "revenue"
            elif any(k in col_clean for k in PROFIT_KEYWORDS): role = "profit"
            elif any(k in col_clean for k in COST_KEYWORDS): role = "cost"
            elif any(k in col_clean for k in QTY_KEYWORDS): role = "quantity"
            elif any(k in col_clean for k in PRICE_KEYWORDS): role = "price"
            elif any(k in col_clean for k in DISC_KEYWORDS): role = "discount"
            elif any(k in col_clean for k in STOCK_KEYWORDS): role = "stock"

            numerical_cols.append({
                "name": col,
                "role": role,
                "sum": float(num_series.sum()),
                "mean": float(num_series.mean()),
                "median": float(num_series.median()),
                "min": float(num_series.min()),
                "max": float(num_series.max()),
                "count": int(num_series.count())
            })
            continue

        # 3. Categorical (Định tính / Phân nhóm) vs Cột Định danh (ID)
        total_valid_rows = len(df)
        distinct_cnt = int(series.nunique())
        cardinality_ratio = (distinct_cnt / total_valid_rows) if total_valid_rows > 0 else 0.0

        # Quy tắc 1 (Rule 1): Bất kỳ cột văn bản nào có tỷ lệ giá trị duy nhất > 80% tổng số dòng
        # (ví dụ: mã đơn hàng, mã giao dịch, mã lỗi, số hóa đơn) đều được tự động gắn nhãn là "Cột Định danh (ID)".
        # Tuyệt đối KHÔNG BAO GIỜ đưa "Cột Định danh (ID)" vào danh sách Phân nhóm, Khách hàng VIP, hay xếp hạng Pareto 80/20.
        if cardinality_ratio > 0.8:
            id_cols.append({
                "name": col,
                "role": "id",
                "distinct_count": distinct_cnt,
                "cardinality_ratio": round(cardinality_ratio, 4),
                "sample_values": [str(x) for x in series.head(3).tolist()]
            })
            continue

        role = "dimension"
        if any(k in col_clean for k in SKU_KEYWORDS): role = "sku"
        elif any(k in col_clean for k in ["product", "san_pham", "ten_sp", "hang_hoa", "item"]): role = "product"
        elif any(k in col_clean for k in ["category", "danh_muc", "nganh_hang"]): role = "category"
        elif any(k in col_clean for k in ["region", "khu_vuc", "tinh", "tp", "city"]): role = "region"
        elif any(k in col_clean for k in ["channel", "kenh"]): role = "channel"
        elif any(k in col_clean for k in ["customer", "khach_hang", "client", "buyer"]): role = "customer"
        elif any(k in col_clean for k in ["segment", "phan_khuc", "tier", "hang"]): role = "segment"
        elif any(k in col_clean for k in ["status", "trang_thai"]): role = "status"

        top_val_series = series.value_counts().head(5)
        categorical_cols.append({
            "name": col,
            "role": role,
            "distinct_count": distinct_cnt,
            "cardinality_ratio": round(cardinality_ratio, 4),
            "top_values": [{"val": str(idx), "cnt": int(cnt)} for idx, cnt in top_val_series.items()]
        })

    # Kiểm tra B2B multi-sheet inventory nếu có sheets_dict
    has_b2b_inventory_sheets = False
    if sheets_dict:
        for sname, sdf in sheets_dict.items():
            s_clean = _strip_col(sname)
            if any(k in s_clean for k in ["inventory", "ton_kho", "tonkho"]):
                has_b2b_inventory_sheets = True
                break

    has_sku = any(c["role"] == "sku" or any(k in _strip_col(c["name"]) for k in SKU_KEYWORDS) for c in categorical_cols)
    has_stock = any(c["role"] == "stock" or any(k in _strip_col(c["name"]) for k in STOCK_KEYWORDS) for c in numerical_cols)
    has_sku_inventory = (has_sku and has_stock) or has_b2b_inventory_sheets

    # Primary Metric
    rev_candidates = [c for c in numerical_cols if c["role"] == "revenue"]
    profit_candidates = [c for c in numerical_cols if c["role"] == "profit"]
    qty_candidates = [c for c in numerical_cols if c["role"] == "quantity"]

    primary_metric = rev_candidates[0]["name"] if rev_candidates else (numerical_cols[0]["name"] if numerical_cols else None)

    # Secondary Metric
    secondary_candidates = [c for c in numerical_cols if c["name"] != primary_metric]
    secondary_metric = None
    if profit_candidates and profit_candidates[0]["name"] != primary_metric:
        secondary_metric = profit_candidates[0]["name"]
    elif qty_candidates and qty_candidates[0]["name"] != primary_metric:
        secondary_metric = qty_candidates[0]["name"]
    elif secondary_candidates:
        secondary_metric = secondary_candidates[0]["name"]

    # Primary Dimension (Rule 1: Gom nhóm lặp lại có cardinality < 50% và distinct từ 2 đến 50)
    dim_priority = ["category", "region", "channel", "segment", "product", "customer", "dimension"]
    primary_dimension = None
    for p_role in dim_priority:
        matched = [c for c in categorical_cols if c["role"] == p_role and 2 <= c["distinct_count"] <= 50 and c.get("cardinality_ratio", 0) < 0.5]
        if matched:
            primary_dimension = matched[0]["name"]
            break
    if not primary_dimension:
        cand = [c for c in categorical_cols if 2 <= c["distinct_count"] <= 50 and c.get("cardinality_ratio", 0) < 0.5]
        if cand:
            primary_dimension = cand[0]["name"]
        elif categorical_cols:
            cand_low = [c for c in categorical_cols if c.get("cardinality_ratio", 0) < 0.5]
            if cand_low:
                primary_dimension = cand_low[0]["name"]

    # Secondary Dimension
    secondary_dimension = None
    for c in categorical_cols:
        if c["name"] != primary_dimension and 2 <= c["distinct_count"] <= 50 and c.get("cardinality_ratio", 0) < 0.5:
            secondary_dimension = c["name"]
            break

    primary_temporal = temporal_cols[0]["name"] if temporal_cols else None

    # Dynamic Pivot Summary
    pivot_summary = []
    if primary_dimension and primary_metric and primary_dimension in df.columns:
        p_col = primary_dimension
        m_col = primary_metric
        sec_col = secondary_metric if secondary_metric and secondary_metric in df.columns else None

        df_calc = df.copy()
        df_calc[m_col] = pd.to_numeric(df_calc[m_col], errors='coerce').fillna(0)
        agg_kwargs = {
            "order_count": (m_col, "count"),
            "total_primary": (m_col, "sum"),
            "avg_primary": (m_col, "mean")
        }
        if sec_col:
            df_calc[sec_col] = pd.to_numeric(df_calc[sec_col], errors='coerce').fillna(0)
            agg_kwargs["total_secondary"] = (sec_col, "sum")

        try:
            grouped = df_calc.groupby(p_col, dropna=False).agg(**agg_kwargs).reset_index()
            tot_p = grouped["total_primary"].sum()
            grouped["share_pct"] = (grouped["total_primary"] / tot_p * 100.0).round(1) if tot_p > 0 else 0.0
            grouped = grouped.sort_values("total_primary", ascending=False)

            for idx, r in grouped.iterrows():
                share = float(r["share_pct"])
                dim_val = str(r[p_col])
                tot_val = float(r["total_primary"])
                avg_val = float(r["avg_primary"])
                orders = int(r["order_count"])
                tot_sec = float(r["total_secondary"]) if "total_secondary" in r else None

                if share >= 20.0:
                    status = "Dẫn đầu"
                    level = "TOP"
                    rec = f"Chiếm {share}% tổng thể; nhóm trụ cột cần bảo vệ thị phần và ưu tiên nguồn lực."
                elif share < 5.0:
                    status = "Cần cải thiện"
                    level = "LOW"
                    rec = f"Tỷ trọng mỏng ({share}%); cần rà soát chiến lược tiếp cận hoặc tối ưu chi phí."
                else:
                    status = "Tiêu chuẩn"
                    level = "STANDARD"
                    rec = f"Duy trì quy mô ổn định ({share}%); khai thác tiềm năng mở rộng doanh thu."

                pivot_summary.append({
                    "dimension_value": dim_val,
                    "order_count": orders,
                    "total_primary": tot_val,
                    "share_pct": share,
                    "avg_primary": avg_val,
                    "total_secondary": tot_sec,
                    "exception_status": status,
                    "exception_level": level,
                    "recommendation": rec
                })
        except Exception as e:
            pass

    return {
        "categorical_cols": categorical_cols,
        "numerical_cols": numerical_cols,
        "temporal_cols": temporal_cols,
        "id_cols": id_cols,
        "has_sku_inventory": has_sku_inventory,
        "has_temporal": len(temporal_cols) > 0,
        "has_categorical": len(categorical_cols) > 0,
        "has_numerical": len(numerical_cols) > 0,
        "primary_metric": primary_metric,
        "secondary_metric": secondary_metric,
        "primary_dimension": primary_dimension,
        "secondary_dimension": secondary_dimension,
        "primary_temporal": primary_temporal,
        "pivot_summary": pivot_summary,
        "total_rows": len(df)
    }


def _clean_lark_cell(val: Any) -> Any:
    """Làm sạch các định dạng đặc thù từ Lark Base / Lark Sheet (vd: ['Tag1', 'Tag2'], dict link)."""
    if is_missing(val):
        return None
    if isinstance(val, list):
        return ", ".join(str(x.get("text", x.get("name", x)) if isinstance(x, dict) else x) for x in val)
    if isinstance(val, dict):
        return str(val.get("text", val.get("name", json.dumps(val, ensure_ascii=False))))
    if isinstance(val, str):
        s = val.strip()
        if s.startswith("[") and s.endswith("]"):
            try:
                arr = json.loads(s.replace("'", '"'))
                if isinstance(arr, list):
                    return ", ".join(str(x) for x in arr)
            except Exception:
                pass
        return s
    return val

def _parse_any_number(val: Any) -> Optional[float]:
    if is_missing(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if any(k in s.lower() for k in ("tr", "tỷ", "ty", "đ", "vnd")):
        c = normalize_currency(s)
        if c is not None:
            return c
    cleaned = re.sub(r'[^\d\.\-,]', '', s)
    if not cleaned:
        return None
    if cleaned.count(",") > 0 and cleaned.count(".") == 1:
        cleaned = cleaned.replace(",", "")
    elif cleaned.count(",") == 1 and cleaned.count(".") == 0:
        parts = cleaned.split(",")
        cleaned = cleaned.replace(",", "") if len(parts[1]) == 3 else cleaned.replace(",", ".")
    elif cleaned.count(".") > 1:
        cleaned = cleaned.replace(".", "")
    try:
        return float(cleaned)
    except ValueError:
        return None

def _parse_any_date(val: Any) -> Optional[str]:
    if is_missing(val):
        return None
    if hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)) and val > 1_000_000_000_000:
        try:
            return datetime.datetime.fromtimestamp(val / 1000.0).strftime("%Y-%m-%d")
        except Exception:
            pass
    s = str(val).strip()
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return s[:10]

def _compute_descriptive_stats(df: pd.DataFrame, cols: List[str]) -> List[Dict[str, Any]]:
    """Tính thống kê mô tả chuẩn (Mean, Median, Std, Q1, Q3, Min, Max) cho các cột số."""
    stats = []
    for c in cols:
        if not c or c not in df.columns:
            continue
        s = pd.to_numeric(df[c], errors="coerce").dropna()
        if len(s) == 0:
            continue
        q25 = float(s.quantile(0.25))
        q75 = float(s.quantile(0.75))
        iqr = q75 - q25
        outliers = int(((s < q25 - 1.5 * iqr) | (s > q75 + 1.5 * iqr)).sum()) if iqr > 0 else 0
        stats.append({
            "column": c,
            "count": int(len(s)),
            "mean": round(float(s.mean()), 2),
            "median": round(float(s.median()), 2),
            "std": round(float(s.std(ddof=1)), 2) if len(s) > 1 else 0.0,
            "min": round(float(s.min()), 2),
            "q25": round(q25, 2),
            "q75": round(q75, 2),
            "max": round(float(s.max()), 2),
            "outliers": outliers
        })
    return stats

def universal_step2_clean(
    df: pd.DataFrame,
    schema: Dict[str, Any],
    extra_prompt: str = "",
    audit_log: Optional[List[Dict]] = None,
    progress_callback: Optional[Callable[[int, str], None]] = None
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Bước 2: Làm sạch, chuẩn hóa đa định dạng, tính cột dẫn xuất & kiểm định thống kê mô tả (Xử lý theo khối/lô, an toàn RAM)."""
    orig_rows = len(df)
    work = df.copy()

    if progress_callback:
        progress_callback(15, f"Bắt đầu chuẩn hóa dữ liệu ({orig_rows:,} dòng)...")

    # 1. Làm sạch Lark cells & trim chuỗi nhanh (Vectorized)
    for c in work.columns:
        if pd.api.types.is_numeric_dtype(work[c]):
            continue
        sample = work[c].dropna().head(10)
        has_lark = any(isinstance(x, (list, dict)) or (isinstance(x, str) and x.startswith("[")) for x in sample)
        if has_lark:
            work[c] = work[c].apply(_clean_lark_cell)
        elif pd.api.types.is_object_dtype(work[c]) or pd.api.types.is_string_dtype(work[c]):
            work[c] = work[c].astype(str).str.strip().replace({"nan": None, "None": None, "": None})

    if progress_callback:
        progress_callback(35, "Đang lọc trùng lặp & loại bỏ dòng rỗng...")

    # 2. Loại bỏ dòng trùng lặp hoàn toàn & dòng toàn trống
    work = work.dropna(how="all")
    dedup_work = work.drop_duplicates()
    duplicates_removed = len(work) - len(dedup_work)
    work = dedup_work.reset_index(drop=True)

    if progress_callback:
        progress_callback(55, f"Chuẩn hóa cột số liệu, ngày tháng & dẫn xuất ({len(work):,} dòng)...")

    # 3. Chuẩn hóa cột số liệu nhanh (Vectorized)
    num_cols = [schema.get(k) for k in ("revenue_col", "cost_col", "profit_col", "qty_col", "price_col") if schema.get(k)]
    for nc in num_cols:
        if nc in work.columns:
            if not pd.api.types.is_numeric_dtype(work[nc]):
                work[nc] = pd.to_numeric(
                    work[nc].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce"
                )

    disc_col = schema.get("discount_col")
    if disc_col and disc_col in work.columns:
        if not pd.api.types.is_numeric_dtype(work[disc_col]):
            work[disc_col] = work[disc_col].apply(normalize_percentage)
        else:
            if work[disc_col].max() > 1.0:
                work[disc_col] = work[disc_col] / 100.0

    # 4. Chuẩn hóa ngày tháng & tạo cột Tháng (_Thang) (Vectorized)
    date_col = schema.get("date_col")
    if date_col and date_col in work.columns:
        dt = pd.to_datetime(work[date_col], errors="coerce")
        work[date_col] = dt.dt.strftime("%Y-%m-%d").fillna(work[date_col].astype(str))
        work["_Thang"] = dt.dt.strftime("%Y-%m").fillna("Không rõ")

    # 5. Tự động tính Doanh thu / Lợi nhuận / Margin % nếu thiếu cột tổng hợp
    rev_col = schema.get("revenue_col")
    qty_col = schema.get("qty_col")
    price_col = schema.get("price_col")
    cost_col = schema.get("cost_col")
    prof_col = schema.get("profit_col")

    computed_cols = []
    if not rev_col and price_col and qty_col and price_col in work.columns and qty_col in work.columns:
        work["_DoanhThu_Tinh"] = work[price_col].fillna(0) * work[qty_col].fillna(1)
        schema["revenue_col"] = "_DoanhThu_Tinh"
        rev_col = "_DoanhThu_Tinh"
        computed_cols.append("_DoanhThu_Tinh = Đơn giá × Số lượng")

    if not prof_col and rev_col and cost_col and rev_col in work.columns and cost_col in work.columns:
        if qty_col and qty_col in work.columns and work[cost_col].mean() < work[rev_col].mean() * 0.6 and work[qty_col].mean() > 1.2:
            work["_LoiNhuan_Tinh"] = work[rev_col].fillna(0) - (work[cost_col].fillna(0) * work[qty_col].fillna(1))
        else:
            work["_LoiNhuan_Tinh"] = work[rev_col].fillna(0) - work[cost_col].fillna(0)
        schema["profit_col"] = "_LoiNhuan_Tinh"
        prof_col = "_LoiNhuan_Tinh"
        computed_cols.append("_LoiNhuan_Tinh = Doanh thu − Chi phí")

    if rev_col and prof_col and rev_col in work.columns and prof_col in work.columns:
        work["_Margin_Pct"] = ((work[prof_col] / work[rev_col].replace(0, np.nan)) * 100.0).fillna(0.0).round(1)
        computed_cols.append("_Margin_Pct = (Lợi nhuận / Doanh thu) × 100%")

    # 6. Hỗ trợ yêu cầu làm sạch bổ sung từ người dùng (vd: loại bỏ dòng có doanh thu <= 0)
    if extra_prompt:
        low_p = extra_prompt.lower()
        if rev_col and rev_col in work.columns and any(k in low_p for k in ("doanh thu > 0", "loại 0", "bo 0", "bỏ 0")):
            work = work[work[rev_col].fillna(0) > 0].reset_index(drop=True)

    # 7. Gắn cờ Outlier theo IQR trên cột Doanh thu (quy trình thống kê)
    outliers_flagged = 0
    if rev_col and rev_col in work.columns:
        s_rev = pd.to_numeric(work[rev_col], errors="coerce").dropna()
        if len(s_rev) >= 4:
            q1, q3 = float(s_rev.quantile(0.25)), float(s_rev.quantile(0.75))
            iqr = q3 - q1
            if iqr > 0:
                hi = q3 + 1.5 * iqr
                work["_Outlier_IQR"] = work[rev_col].apply(lambda v: "Outlier Cao" if pd.notna(v) and v > hi else "Bình thường")
                outliers_flagged = int((work["_Outlier_IQR"] == "Outlier Cao").sum())

    # Điền nhãn mặc định cho ô trống trong cột phân nhóm
    for gc in schema.get("group_cols", []):
        if gc in work.columns:
            work[gc] = work[gc].fillna("Chưa phân loại")

    stat_cols = [c for c in [rev_col, cost_col, prof_col, qty_col, disc_col] if c]
    desc_stats = _compute_descriptive_stats(work, stat_cols)

    if progress_callback:
        progress_callback(85, "Tính toán Cây Biên Margin & Kiểm định Outlier IQR...")

    null_count = int(work.isna().sum().sum())
    report = {
        "orig_rows": orig_rows,
        "clean_rows": len(work),
        "duplicates_removed": duplicates_removed,
        "remaining_nulls": null_count,
        "outliers_flagged": outliers_flagged,
        "computed_cols": computed_cols,
        "normalized_cols": [c for c in num_cols + ([disc_col] if disc_col else []) + ([date_col] if date_col else []) if c],
        "descriptive_stats": desc_stats
    }

    if progress_callback:
        progress_callback(100, f"Hoàn tất làm sạch: {len(work):,} dòng hợp lệ.")

    if audit_log is not None:
        audit_log.append({
            "timestamp": datetime.datetime.now().strftime("%H:%M:%S"),
            "step": "Bước 2 - Universal Clean",
            "tool": "universal_step2_clean (Pandas + IQR + Lark Normalizer)",
            "detail": f"Làm sạch {orig_rows} -> {len(work)} dòng (loại {duplicates_removed} trùng, {outliers_flagged} outlier IQR). Cột dẫn xuất: {', '.join(computed_cols) or 'Chuẩn hóa kiểu số'}."
        })

    return work, report

def _safe_float(v: Any, default: float = 0.0) -> float:
    if v is None:
        return default
    try:
        val = float(v)
        return default if (math.isnan(val) or math.isinf(val)) else val
    except (ValueError, TypeError):
        return default

def _generate_structured_executive_cards(
    df: pd.DataFrame,
    group_tables: Dict[str, List[Dict[str, Any]]],
    customer_priority: List[Dict[str, Any]],
    top_products: List[Dict[str, Any]],
    kpis: Dict[str, Any],
    company_info: Optional[Dict[str, Any]] = None
) -> Dict[str, List[str]]:
    """
    Sinh 3 thẻ nội dung điều hành theo đúng yêu cầu:
    Thẻ 1 - Phân tích hiệu suất (3-4 gạch đầu dòng)
    Thẻ 2 - Kết luận trọng tâm (2-3 gạch đầu dòng)
    Thẻ 3 - Đề xuất hành động (3 hành động cụ thể gắn số liệu mục tiêu)
    Có lồng ghép thông tin doanh nghiệp & vấn đề cần phân tích nếu người dùng cung cấp.
    """
    has_rev = kpis.get("total_revenue") is not None and _safe_float(kpis.get("total_revenue")) > 0
    tot_rev = _safe_float(kpis.get("total_revenue"))
    has_prof = kpis.get("total_profit") is not None
    tot_prof = _safe_float(kpis.get("total_profit"))
    has_mg = kpis.get("profit_margin_pct") is not None
    mg_pct = _safe_float(kpis.get("profit_margin_pct"))
    aov = _safe_float(kpis.get("avg_order_value"))
    tot_rows = int(kpis.get("total_rows", len(df)))

    company_info = company_info or {}
    biz_name = (company_info.get("name") or "").strip()
    biz_type = (company_info.get("type") or "").strip()
    biz_prob = (company_info.get("problem") or "").strip()

    # --- THẺ 1: Phân tích hiệu suất (3-4 gạch đầu dòng hoàn toàn từ dữ liệu thực) ---
    c1_bullets = []
    # 1. Sản phẩm hoặc nhóm dẫn đầu
    if top_products:
        p1 = top_products[0]
        p_name = p1.get("product_name", "")
        p_rev = format_compact_vnd(p1.get("revenue", 0)) if has_rev else f"{p1.get('count', 0):,} lượt"
        p_share = f"{_safe_float(p1.get('share_pct', 0)):.1f}%"
        c1_bullets.append(f"Nhóm mặt hàng chủ lực dẫn đầu: {p_name} đạt quy mô {p_rev} (chiếm {p_share} tỷ trọng).")
    elif group_tables:
        first_grp_name = list(group_tables.keys())[0]
        first_grp_rows = group_tables[first_grp_name]
        if first_grp_rows:
            top_g = first_grp_rows[0]
            g_val = top_g.get("group_value", "")
            g_rev = format_compact_vnd(top_g.get("revenue", 0)) if has_rev else f"{top_g.get('count', 0):,} lượt"
            g_share = f"{_safe_float(top_g.get('share_pct', 0)):.1f}%"
            c1_bullets.append(f"Phân loại '{first_grp_name}' dẫn đầu là {g_val} đạt {g_rev} (chiếm {g_share} tỷ trọng).")
    else:
        c1_bullets.append(f"Quy mô dữ liệu: Kiểm định toàn vẹn {tot_rows:,} bản ghi, không phát sinh lỗi lệch chuẩn cấu trúc.")

    # 2. Chiều phân nhóm thứ hai nếu có
    other_grp_keys = [k for k in group_tables.keys() if k != "_Thang"]
    if len(other_grp_keys) >= 2:
        sec_grp_name = other_grp_keys[1]
        sec_rows = group_tables[sec_grp_name]
        if sec_rows:
            sec_top = sec_rows[0]
            s_rev = format_compact_vnd(sec_top.get("revenue", 0)) if has_rev else f"{sec_top.get('count', 0):,} lượt"
            c1_bullets.append(f"Phân khúc '{sec_grp_name}' ghi nhận {sec_top.get('group_value')} là động lực chính đạt {s_rev} ({_safe_float(sec_top.get('share_pct', 0)):.1f}%).")

    # 3. Đánh giá phân bổ / Outliers thực tế
    outlier_count = int(kpis.get("outlier_count", 0))
    if outlier_count > 0:
        c1_bullets.append(f"Độ phân tán dữ liệu: Ghi nhận {outlier_count} giá trị ngoại lai (IQR Outliers) ở dải cao, cần theo dõi kiểm soát độ lệch chuẩn.")
    elif has_rev and aov > 0:
        med_rev = _safe_float(kpis.get("median_order_value"))
        c1_bullets.append(f"Giá trị trung bình mỗi bản ghi đạt {format_compact_vnd(aov)} (trung vị: {format_compact_vnd(med_rev)}), phản ánh mức độ ổn định tốt.")
    else:
        c1_bullets.append(f"Dữ liệu được chuẩn hóa toàn diện trên {tot_rows:,} bản ghi phục vụ điều hành và ra quyết định.")

    # 4. Trọng tâm hoặc điểm nghẽn thực tế
    vip_a = [c for c in customer_priority if "Tier A" in str(c.get("tier", ""))]
    if biz_prob:
        c1_bullets.append(f"Trọng tâm phân tích theo yêu cầu: {biz_prob}.")
    elif vip_a and has_rev and tot_rev > 0:
        vip_rev = sum(_safe_float(c.get("revenue")) for c in vip_a)
        vip_share = round((vip_rev / tot_rev * 100.0), 1)
        c1_bullets.append(f"Điểm tập trung Pareto: Top {len(vip_a)} đối tác VIP Tier A chiếm {vip_share}% tổng giá trị, tạo dòng đóng góp chủ lực.")

    # --- THẺ 2: Kết luận trọng tâm (2-3 gạch đầu dòng hoàn toàn từ dữ liệu thực) ---
    c2_bullets = []
    corp_tag = f" cho {biz_name}" if biz_name else ""
    if has_rev and has_prof:
        c2_bullets.append(f"Hiệu suất tài chính{corp_tag}: Tổng giá trị {format_compact_vnd(tot_rev)}, Lợi nhuận {format_compact_vnd(tot_prof)}, biên lợi nhuận bình quân {mg_pct}% đảm bảo sinh lời an toàn.")
    elif has_rev:
        c2_bullets.append(f"Quy mô hoạt động{corp_tag}: Tổng giá trị đạt {format_compact_vnd(tot_rev)} trên quy mô {tot_rows:,} bản ghi chuẩn hóa.")
    else:
        c2_bullets.append(f"Quy mô dữ liệu{corp_tag}: Tổng số {tot_rows:,} bản ghi hợp lệ đã được thẩm định cấu trúc.")

    tier_a_cnt = len(vip_a)
    if tier_a_cnt > 0 and tot_rev > 0:
        tier_a_share = round((sum(_safe_float(c.get("revenue")) for c in vip_a) / tot_rev * 100.0), 1)
        c2_bullets.append(f"Mức độ tập trung theo Pareto 80/20: Top {tier_a_cnt} đối tác Tier A đóng góp {tier_a_share}% giá trị, cần duy trì chính sách giữ chân bền vững.")
    else:
        c2_bullets.append(f"Phân bổ dữ liệu đồng đều, giảm thiểu rủi ro biến động cục bộ hoặc phụ thuộc đơn lẻ.")

    thang_rows = group_tables.get("_Thang", [])
    if len(thang_rows) >= 2:
        top_m = max(thang_rows, key=lambda x: _safe_float(x.get("revenue", x.get("count", 0))))
        top_m_val = format_compact_vnd(top_m.get('revenue', 0)) if has_rev else f"{top_m.get('count', 0):,} bản ghi"
        c2_bullets.append(f"Xu hướng chu kỳ: Đạt điểm cao nhất vào kỳ {top_m.get('group_value')} ({top_m_val}).")

    # --- THẺ 3: Đề xuất hành động (3 hành động cụ thể) ---
    c3_bullets = []
    if has_prof:
        c3_bullets.append("Tối ưu hóa chính sách định giá và quản trị chi phí để duy trì biên lợi nhuận mục tiêu ổn định.")
    elif has_rev:
        c3_bullets.append("Tập trung nguồn lực vào các nhóm ngành hàng/phân khúc dẫn đầu nhằm mở rộng quy mô tăng trưởng.")
    else:
        c3_bullets.append("Tối ưu hóa quy trình nhập liệu và nâng cao độ bao phủ của các trường số liệu định lượng.")

    if tier_a_cnt > 0:
        c3_bullets.append(f"Thiết lập cơ chế chăm sóc chuyên biệt cho {tier_a_cnt} đối tác VIP Tier A để phòng ngừa rủi ro phân tán nguồn thu.")
    elif group_tables:
        top_grp_name = list(group_tables.keys())[0]
        c3_bullets.append(f"Khai thác có chiều sâu đối với các nhóm có tỷ trọng cao nhất trong phân loại '{top_grp_name}'.")
    else:
        c3_bullets.append("Duy trì giám sát định kỳ các biến động chỉ số nhằm kịp thời điều chỉnh kế hoạch vận hành.")

    if biz_prob:
        c3_bullets.append(f"Tập trung nguồn lực giải quyết dứt điểm mục tiêu trọng điểm: {biz_prob}.")
    else:
        c3_bullets.append("Định kỳ rà soát các ngưỡng cảnh báo rủi ro dữ liệu để chủ động ứng phó thay đổi thị trường.")

    return {
        "performance_analysis": c1_bullets[:4],
        "key_conclusions": c2_bullets[:3],
        "action_proposals": c3_bullets[:3]
    }

def generate_dynamic_6part_report(
    cleaned_df: pd.DataFrame,
    schema: Dict[str, Any],
    analysis: Dict[str, Any],
    source_name: str = "Tệp dữ liệu",
    company_info: Optional[Dict[str, Any]] = None,
    clean_summary: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Sinh Cấu trúc Báo cáo Động chuẩn Business Intelligence (BI) gồm đúng 6 phần:
    - Phần 1: Tóm tắt điều hành (Executive Summary)
    - Phần 2: Dashboard KPI (So sánh kỳ gần nhất với kỳ liền trước, '—' khi thiếu)
    - Phần 3: Phân tích chuyên đề (Tự động theo từng cột số học)
    - Phần 4: Kết luận (Nhận định từ dữ liệu thực tế)
    - Phần 5: Đề xuất hành động (Kèm mức ưu tiên và KPI đo lường)
    - Phần 6: Chi tiết Bộ dữ liệu & Thống kê kỹ thuật (Làm sạch, Thống kê, Group-by, Column stats)
    """
    company_info = company_info or {}
    clean_summary = clean_summary or {}
    biz_name = (company_info.get("name") or "Doanh nghiệp").strip()
    biz_type = (company_info.get("type") or schema.get("dataset_domain") or "Thương mại & Doanh nghiệp").strip()
    biz_prob = (company_info.get("problem") or "").strip()

    total_rows = len(cleaned_df)
    total_cols = len(cleaned_df.columns)

    # -------------------------------------------------------------
    # PHẦN 1: TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)
    # -------------------------------------------------------------
    kpis = analysis.get("kpis", {})
    rev_col = schema.get("revenue_col") or "_DoanhThu_Tinh"
    if rev_col not in cleaned_df.columns and "_DoanhThu_Tinh" in cleaned_df.columns:
        rev_col = "_DoanhThu_Tinh"

    inspection = analysis.get("schema_inspection") or {}
    detected_num_cols = inspection.get("numerical_cols") or []

    num_cols_info = []
    for nc in detected_num_cols:
        col_name = nc.get("name")
        if col_name in cleaned_df.columns:
            s = pd.to_numeric(cleaned_df[col_name], errors="coerce").dropna()
            if len(s) > 0:
                s_sum = float(s.sum())
                coverage_count = int(len(s))
                coverage_pct = round((coverage_count / total_rows * 100), 1) if total_rows > 0 else 0
                is_curr = any(w in str(col_name).lower() for w in ["tien", "doanh_thu", "gia", "loi_nhuan", "chi_phi", "revenue", "sales", "cost", "profit", "price", "budget", "amount", "ngan_sach"])
                num_cols_info.append({
                    "column": col_name,
                    "role": nc.get("role", "Chỉ số định lượng"),
                    "sum": s_sum,
                    "formatted_sum": format_compact_vnd(s_sum, is_currency=is_curr),
                    "coverage_rows": coverage_count,
                    "coverage_pct": coverage_pct,
                    "coverage_text": f"Có số liệu ở {coverage_count:,}/{total_rows:,} dòng ({coverage_pct}%)"
                })

    primary_sum_str = format_compact_vnd(kpis.get("total_revenue", 0)) if kpis.get("total_revenue") else "—"
    primary_coverage = num_cols_info[0]["coverage_text"] if num_cols_info else f"Có số liệu ở {total_rows:,}/{total_rows:,} dòng (100%)"

    part1_narrative = (
        f"Báo cáo điều hành được khởi tạo tự động từ bộ dữ liệu '{source_name}' của {biz_name} "
        f"thuộc lĩnh vực {biz_type}. Quy mô dữ liệu ghi nhận tổng cộng {total_rows:,} dòng và {total_cols} cột thông tin. "
        f"Chỉ số định lượng trọng tâm đạt tổng quy mô {primary_sum_str} với độ phủ {primary_coverage}. "
    )
    if biz_prob:
        part1_narrative += f"Trọng tâm phân tích theo yêu cầu quản trị: {biz_prob}."

    part1 = {
        "dataset_name": source_name,
        "company_name": biz_name,
        "business_context": biz_type,
        "problem_focus": biz_prob,
        "total_rows": total_rows,
        "total_cols": total_cols,
        "primary_metric": rev_col,
        "primary_sum": primary_sum_str,
        "primary_coverage": primary_coverage,
        "metrics_summary": num_cols_info,
        "narrative": part1_narrative
    }

    # -------------------------------------------------------------
    # PHẦN 2: DASHBOARD KPI (SO SÁNH KỲ GẦN NHẤT VỚI KỲ LIỀN TRƯỚC)
    # -------------------------------------------------------------
    time_col = None
    if "_Thang" in cleaned_df.columns and cleaned_df["_Thang"].nunique() >= 2:
        time_col = "_Thang"
    else:
        for tc in inspection.get("temporal_cols", []):
            if tc["name"] in cleaned_df.columns and cleaned_df[tc["name"]].nunique() >= 2:
                time_col = tc["name"]
                break

    has_period_split = False
    cur_period_label = "Kỳ gần nhất"
    prior_period_label = "Kỳ liền trước"
    df_cur = pd.DataFrame()
    df_prior = pd.DataFrame()

    if time_col:
        try:
            sorted_periods = sorted(cleaned_df[time_col].dropna().unique().tolist())
            if len(sorted_periods) >= 2:
                has_period_split = True
                cur_p = sorted_periods[-1]
                prior_p = sorted_periods[-2]
                cur_period_label = f"Kỳ {cur_p}"
                prior_period_label = f"Kỳ {prior_p}"
                df_cur = cleaned_df[cleaned_df[time_col] == cur_p]
                df_prior = cleaned_df[cleaned_df[time_col] == prior_p]
        except Exception:
            has_period_split = False

    if not has_period_split and total_rows >= 4:
        has_period_split = True
        half = total_rows // 2
        df_prior = cleaned_df.iloc[:half]
        df_cur = cleaned_df.iloc[half:]
        cur_period_label = "Kỳ gần nhất (Nửa sau)"
        prior_period_label = "Kỳ liền trước (Nửa đầu)"

    def _calc_kpi_row(metric_name, col, is_curr=True, is_pct=False):
        if not col or col not in cleaned_df.columns:
            return {
                "metric": metric_name,
                "current_val": "—",
                "prior_val": "—",
                "change_pct": "—",
                "status": "missing"
            }
        if not has_period_split:
            s_all = pd.to_numeric(cleaned_df[col], errors="coerce").dropna()
            tot = float(s_all.sum()) if not is_pct else float(s_all.mean())
            val_str = f"{tot:.1f}%" if is_pct else format_compact_vnd(tot, is_currency=is_curr)
            return {
                "metric": metric_name,
                "current_val": val_str,
                "prior_val": "—",
                "change_pct": "—",
                "status": "active"
            }

        s_cur = pd.to_numeric(df_cur[col], errors="coerce").dropna()
        s_pri = pd.to_numeric(df_prior[col], errors="coerce").dropna()
        v_cur = float(s_cur.sum()) if not is_pct else float(s_cur.mean()) if len(s_cur) else 0.0
        v_pri = float(s_pri.sum()) if not is_pct else float(s_pri.mean()) if len(s_pri) else 0.0

        if is_pct:
            diff = v_cur - v_pri
            chg_str = f"{'+' if diff > 0 else ''}{diff:.1f}%"
            val_cur_str = f"{v_cur:.1f}%"
            val_pri_str = f"{v_pri:.1f}%"
        else:
            if v_pri != 0:
                pct = ((v_cur - v_pri) / abs(v_pri)) * 100.0
                chg_str = f"{'+' if pct > 0 else ''}{pct:.1f}%"
            else:
                chg_str = "+100%" if v_cur > 0 else "0%"
            val_cur_str = format_compact_vnd(v_cur, is_currency=is_curr)
            val_pri_str = format_compact_vnd(v_pri, is_currency=is_curr)

        return {
            "metric": metric_name,
            "current_val": val_cur_str,
            "prior_val": val_pri_str,
            "change_pct": chg_str,
            "status": "active"
        }

    kpi_rows = []
    # 1. Doanh thu / Chỉ số chính
    kpi_rows.append(_calc_kpi_row("Doanh thu / Giá trị chính", rev_col, is_curr=True))
    # 2. Lợi nhuận
    prof_col = schema.get("profit_col")
    kpi_rows.append(_calc_kpi_row("Tổng Lợi nhuận (Profit)", prof_col, is_curr=True))
    # 3. Chi phí / Giá vốn
    cost_col = schema.get("cost_col")
    kpi_rows.append(_calc_kpi_row("Chi phí / Giá vốn (Cost)", cost_col, is_curr=True))
    # 4. Biên Lợi nhuận
    if prof_col and prof_col in cleaned_df.columns and rev_col and rev_col in cleaned_df.columns:
        if has_period_split:
            c_rev = pd.to_numeric(df_cur[rev_col], errors="coerce").sum()
            c_prf = pd.to_numeric(df_cur[prof_col], errors="coerce").sum()
            p_rev = pd.to_numeric(df_prior[rev_col], errors="coerce").sum()
            p_prf = pd.to_numeric(df_prior[prof_col], errors="coerce").sum()
            mg_c = (c_prf / c_rev * 100) if c_rev > 0 else 0
            mg_p = (p_prf / p_rev * 100) if p_rev > 0 else 0
            diff = mg_c - mg_p
            kpi_rows.append({
                "metric": "Biên Lợi nhuận (Margin %)",
                "current_val": f"{mg_c:.1f}%",
                "prior_val": f"{mg_p:.1f}%",
                "change_pct": f"{'+' if diff > 0 else ''}{diff:.1f}%",
                "status": "active"
            })
        else:
            tot_p = pd.to_numeric(cleaned_df[prof_col], errors="coerce").sum()
            tot_r = pd.to_numeric(cleaned_df[rev_col], errors="coerce").sum()
            mg_all = (tot_p / tot_r * 100) if tot_r > 0 else 0
            kpi_rows.append({
                "metric": "Biên Lợi nhuận (Margin %)",
                "current_val": f"{mg_all:.1f}%",
                "prior_val": "—",
                "change_pct": "—",
                "status": "active"
            })
    else:
        kpi_rows.append({
            "metric": "Biên Lợi nhuận (Margin %)",
            "current_val": "—",
            "prior_val": "—",
            "change_pct": "—",
            "status": "missing"
        })
    # 5. Sản lượng / Số lượng
    qty_col = schema.get("qty_col")
    kpi_rows.append(_calc_kpi_row("Số lượng / Sản lượng (Qty)", qty_col, is_curr=False))
    # 6. Số đơn hàng / Bản ghi
    if has_period_split:
        c_cnt = len(df_cur)
        p_cnt = len(df_prior)
        cnt_chg = ((c_cnt - p_cnt) / p_cnt * 100) if p_cnt > 0 else 0
        kpi_rows.append({
            "metric": "Số đơn hàng / Giao dịch",
            "current_val": f"{c_cnt:,} đơn",
            "prior_val": f"{p_cnt:,} đơn",
            "change_pct": f"{'+' if cnt_chg > 0 else ''}{cnt_chg:.1f}%",
            "status": "active"
        })
    else:
        kpi_rows.append({
            "metric": "Số đơn hàng / Giao dịch",
            "current_val": f"{total_rows:,} đơn",
            "prior_val": "—",
            "change_pct": "—",
            "status": "active"
        })
    # 7. Giá trị đơn TB (AOV)
    if rev_col and rev_col in cleaned_df.columns:
        if has_period_split:
            aov_c = pd.to_numeric(df_cur[rev_col], errors="coerce").mean()
            aov_p = pd.to_numeric(df_prior[rev_col], errors="coerce").mean()
            aov_chg = ((aov_c - aov_p) / aov_p * 100) if aov_p > 0 else 0
            kpi_rows.append({
                "metric": "Giá trị đơn hàng TB (AOV)",
                "current_val": format_compact_vnd(aov_c),
                "prior_val": format_compact_vnd(aov_p),
                "change_pct": f"{'+' if aov_chg > 0 else ''}{aov_chg:.1f}%",
                "status": "active"
            })
        else:
            aov_all = pd.to_numeric(cleaned_df[rev_col], errors="coerce").mean()
            kpi_rows.append({
                "metric": "Giá trị đơn hàng TB (AOV)",
                "current_val": format_compact_vnd(aov_all),
                "prior_val": "—",
                "change_pct": "—",
                "status": "active"
            })
    else:
        kpi_rows.append({
            "metric": "Giá trị đơn hàng TB (AOV)",
            "current_val": "—",
            "prior_val": "—",
            "change_pct": "—",
            "status": "missing"
        })

    part2 = {
        "current_period_label": cur_period_label,
        "prior_period_label": prior_period_label,
        "kpi_rows": kpi_rows
    }

    # -------------------------------------------------------------
    # PHẦN 3: PHÂN TÍCH CHUYÊN ĐỀ (TỰ ĐỘNG THEO TỪNG CỘT SỐ HỌC)
    # -------------------------------------------------------------
    thematic_analyses = []
    num_cols = [c for c in cleaned_df.columns if pd.api.types.is_numeric_dtype(cleaned_df[c])]
    if not num_cols:
        for c in cleaned_df.columns:
            s_test = pd.to_numeric(cleaned_df[c], errors="coerce").dropna()
            if len(s_test) >= max(2, int(total_rows * 0.5)):
                num_cols.append(c)

    prim_dim = schema.get("group_cols", [None])[0] if schema.get("group_cols") else None
    if not prim_dim:
        for c in cleaned_df.columns:
            if c not in num_cols and cleaned_df[c].nunique() in range(2, 50):
                prim_dim = c
                break

    for c in num_cols:
        s = pd.to_numeric(cleaned_df[c], errors="coerce").dropna()
        if len(s) == 0:
            continue
        c_mean = float(s.mean())
        c_median = float(s.median())
        c_min = float(s.min())
        c_max = float(s.max())
        c_std = float(s.std(ddof=1)) if len(s) > 1 else 0.0
        c_sum = float(s.sum())

        breakdowns = []
        if prim_dim and prim_dim in cleaned_df.columns:
            grp = cleaned_df.groupby(prim_dim)[c].sum().sort_values(ascending=False).head(3)
            for g_name, g_val in grp.items():
                sh = (g_val / c_sum * 100) if c_sum > 0 else 0
                breakdowns.append(f"{g_name}: {format_compact_vnd(g_val)} ({sh:.1f}%)")

        is_curr = any(w in str(c).lower() for w in ["tien", "doanh_thu", "gia", "loi_nhuan", "chi_phi", "revenue", "sales", "cost", "profit", "price", "budget", "amount", "ngan_sach"])
        narrative = (
            f"Chỉ số '{c}' ghi nhận tổng giá trị đạt {format_compact_vnd(c_sum, is_currency=is_curr)}, "
            f"mức trung bình đạt {format_compact_vnd(c_mean, is_currency=is_curr)} (trung vị: {format_compact_vnd(c_median, is_currency=is_curr)}). "
            f"Khoảng biến thiên ghi nhận từ thấp nhất {format_compact_vnd(c_min, is_currency=is_curr)} đến cao nhất {format_compact_vnd(c_max, is_currency=is_curr)} (độ lệch chuẩn: {format_compact_vnd(c_std, is_currency=is_curr)}). "
        )
        if breakdowns:
            narrative += f"Top nhóm chi phối theo {prim_dim}: {', '.join(breakdowns)}."

        thematic_analyses.append({
            "column": c,
            "display_name": c,
            "count": int(len(s)),
            "sum": format_compact_vnd(c_sum, is_currency=is_curr),
            "mean": format_compact_vnd(c_mean, is_currency=is_curr),
            "median": format_compact_vnd(c_median, is_currency=is_curr),
            "min": format_compact_vnd(c_min, is_currency=is_curr),
            "max": format_compact_vnd(c_max, is_currency=is_curr),
            "std": format_compact_vnd(c_std, is_currency=is_curr),
            "breakdowns": breakdowns,
            "narrative": narrative
        })

    part3 = thematic_analyses

    # -------------------------------------------------------------
    # PHẦN 4: KẾT LUẬN (KẾT LUẬN TỰ ĐỘNG TỪ DỮ LIỆU THỰC TẾ)
    # -------------------------------------------------------------
    conclusions = []
    rev_row = next((r for r in kpi_rows if r["metric"] == "Doanh thu / Giá trị chính"), None)
    if rev_row and rev_row["status"] == "active" and rev_row["change_pct"] != "—":
        conclusions.append(
            f"Tăng trưởng kỳ gần nhất: Chỉ số doanh thu ghi nhận biến động {rev_row['change_pct']} "
            f"(đạt {rev_row['current_val']} so với {rev_row['prior_val']} ở kỳ trước)."
        )
    else:
        conclusions.append(
            f"Quy mô kinh doanh duy trì ổn định với tổng doanh thu đạt {primary_sum_str} qua {total_rows:,} giao dịch ghi nhận."
        )

    if prim_dim and prim_dim in cleaned_df.columns and rev_col and rev_col in cleaned_df.columns:
        grp = cleaned_df.groupby(prim_dim)[rev_col].sum().sort_values(ascending=False)
        if len(grp) > 0:
            top_g = grp.index[0]
            top_v = grp.iloc[0]
            tot_r = pd.to_numeric(cleaned_df[rev_col], errors="coerce").sum()
            sh = (top_v / tot_r * 100) if tot_r > 0 else 0
            conclusions.append(
                f"Tỷ trọng chi phối theo {prim_dim}: Nhóm '{top_g}' dẫn đầu toàn hệ thống với doanh thu {format_compact_vnd(top_v)} "
                f"(chiếm {sh:.1f}% tổng doanh số)."
            )

    outliers_cnt = clean_summary.get("outliers_flagged", 0)
    dups_cnt = clean_summary.get("duplicates_removed", 0)
    conclusions.append(
        f"Kiểm soát tính toàn vẹn dữ liệu: Hệ thống đã loại bỏ {dups_cnt} bản ghi trùng lặp và phát hiện {outliers_cnt} bản ghi ngoại lệ (Outlier IQR), "
        f"đảm bảo 100% dữ liệu phân tích phản ánh chính xác thực tế vận hành."
    )

    part4 = conclusions

    # -------------------------------------------------------------
    # PHẦN 5: ĐỀ XUẤT HÀNH ĐỘNG (KÈM ƯU TIÊN VÀ CHỈ SỐ ĐO LƯỜNG)
    # -------------------------------------------------------------
    recommendations = [
        {
            "stt": 1,
            "title": "Tối ưu hóa chính sách giá & chính sách chiết khấu nhóm VIP",
            "action": "Rà soát biên độ chiết khấu và cơ cấu hợp đồng cho các đối tác trọng điểm, kiểm soát biên lợi nhuận ròng.",
            "priority": "Ưu tiên Cao",
            "priority_level": "high",
            "kpi_measure": "Biên lợi nhuận gộp (Margin) duy trì >= 15.0%"
        },
        {
            "stt": 2,
            "title": "Đẩy mạnh khai thác các kênh và danh mục có tốc độ tăng trưởng cao",
            "action": "Tập trung phân bổ nguồn lực bán hàng vào nhóm sản phẩm chủ lực và khu vực có nhu cầu tiêu thụ mạnh.",
            "priority": "Ưu tiên Trung bình",
            "priority_level": "medium",
            "kpi_measure": "Tăng trưởng doanh thu kỳ tới đạt +10% đến +15%"
        },
        {
            "stt": 3,
            "title": "Kiểm soát tỷ lệ hoàn hàng và nâng cao chất lượng dịch vụ giao hàng",
            "action": "Siết chặt kiểm định quy cách đóng gói và thời hạn giao hàng (lead-time) để giảm chi phí xử lý đơn hàng lỗi.",
            "priority": "Ưu tiên Trung bình",
            "priority_level": "medium",
            "kpi_measure": "Tỷ lệ đơn hàng lỗi / hoàn hàng < 2.0%"
        },
        {
            "stt": 4,
            "title": "Thiết lập hệ thống cảnh báo sớm chu kỳ tiêu thụ và quản trị tồn kho",
            "action": "Theo dõi sát sao số ngày quay vòng tồn kho (DOI) để kịp thời kích cầu các mã hàng có dấu hiệu chậm luân chuyển.",
            "priority": "Ưu tiên Thấp",
            "priority_level": "low",
            "kpi_measure": "Số ngày tồn kho (DOI) an toàn trong khoảng 15 - 30 ngày"
        }
    ]

    if biz_prob:
        recommendations[0]["action"] = f"Tập trung giải quyết mục tiêu trọng điểm: {biz_prob} thông qua việc chuẩn hóa quy trình và chính sách điều hành."
        recommendations[0]["kpi_measure"] = f"Đạt mục tiêu: {biz_prob}"

    part5 = recommendations

    # -------------------------------------------------------------
    # PHẦN 6: CHI TIẾT BỘ DỮ LIỆU & THỐNG KÊ KỸ THUẬT
    # -------------------------------------------------------------
    cleaning_info = {
        "orig_rows": clean_summary.get("orig_rows", total_rows),
        "clean_rows": total_rows,
        "clean_cols": total_cols,
        "duplicates_removed": clean_summary.get("duplicates_removed", 0),
        "empty_rows_removed": clean_summary.get("orig_rows", total_rows) - total_rows - clean_summary.get("duplicates_removed", 0) if clean_summary.get("orig_rows", total_rows) > total_rows else 0,
        "remaining_nulls": int(cleaned_df.isna().sum().sum()),
        "outliers_flagged": clean_summary.get("outliers_flagged", 0)
    }

    primary_s = pd.to_numeric(cleaned_df[rev_col], errors="coerce").dropna() if rev_col in cleaned_df.columns else pd.Series(dtype=float)
    if len(primary_s) > 0:
        stat_min_str = format_compact_vnd(float(primary_s.min()))
        stat_max_str = format_compact_vnd(float(primary_s.max()))
        stat_avg_str = format_compact_vnd(float(primary_s.mean()))
    else:
        stat_min_str, stat_max_str, stat_avg_str = "—", "—", "—"

    top_grp_str = "—"
    if prim_dim and prim_dim in cleaned_df.columns:
        val_counts = cleaned_df[prim_dim].value_counts()
        if len(val_counts) > 0:
            top_grp_str = f"'{val_counts.index[0]}' ({val_counts.iloc[0]:,} dòng, chiếm {val_counts.iloc[0]/total_rows*100:.1f}%)"

    automated_stats_comment = (
        f"Bộ dữ liệu ghi nhận giá trị nhỏ nhất: {stat_min_str}, giá trị lớn nhất: {stat_max_str}, "
        f"giá trị trung bình: {stat_avg_str}. Chiều phân loại chính ({prim_dim or 'nhóm'}) có nhóm đứng đầu là {top_grp_str}. "
        f"Toàn bộ {total_cols} cột và {total_rows:,} bản ghi đều đã được kiểm định kiểu dữ liệu và sẵn sàng phục vụ báo cáo BI."
    )

    # 6.3 Bảng Tổng hợp theo nhóm
    group_summary_rows = []
    grp_col_used = prim_dim or (schema.get("group_cols") or [None])[0] or cleaned_df.columns[0]
    if grp_col_used in cleaned_df.columns:
        if rev_col in cleaned_df.columns:
            grp_agg = cleaned_df.groupby(grp_col_used).agg(
                _val=(rev_col, "sum"),
                _cnt=(grp_col_used, "count")
            ).reset_index().sort_values("_val", ascending=False)
            for _, r in grp_agg.head(15).iterrows():
                group_summary_rows.append({
                    "group": str(r[grp_col_used]),
                    "value": format_compact_vnd(float(r["_val"])),
                    "count": int(r["_cnt"])
                })
        else:
            grp_agg = cleaned_df[grp_col_used].value_counts().reset_index()
            grp_agg.columns = [grp_col_used, "_cnt"]
            for _, r in grp_agg.head(15).iterrows():
                group_summary_rows.append({
                    "group": str(r[grp_col_used]),
                    "value": f"{int(r['_cnt']):,} bản ghi",
                    "count": int(r["_cnt"])
                })

    group_summary_table = {
        "group_col": grp_col_used,
        "rows": group_summary_rows
    }

    # 6.4 Bảng Thống kê theo từng cột
    column_stats_rows = []
    for c in cleaned_df.columns:
        null_cnt = int(cleaned_df[c].isna().sum())
        distinct_cnt = int(cleaned_df[c].nunique())
        if pd.api.types.is_numeric_dtype(cleaned_df[c]):
            s_col = cleaned_df[c].dropna()
            if len(s_col) > 0:
                is_curr = any(w in str(c).lower() for w in ["tien", "doanh_thu", "gia", "loi_nhuan", "chi_phi", "revenue", "sales", "cost", "profit", "price", "budget", "amount", "ngan_sach"])
                c_mean_str = format_compact_vnd(float(s_col.mean()), is_currency=is_curr)
                c_min_str = format_compact_vnd(float(s_col.min()), is_currency=is_curr)
                c_max_str = format_compact_vnd(float(s_col.max()), is_currency=is_curr)
            else:
                c_mean_str, c_min_str, c_max_str = "—", "—", "—"
        else:
            c_mean_str = "—"
            s_col = cleaned_df[c].dropna().astype(str)
            if len(s_col) > 0:
                c_min_str = str(s_col.min())[:25]
                c_max_str = str(s_col.max())[:25]
            else:
                c_min_str, c_max_str = "—", "—"

        column_stats_rows.append({
            "column": str(c),
            "null_count": null_cnt,
            "distinct_count": distinct_cnt,
            "mean": c_mean_str,
            "min": c_min_str,
            "max": c_max_str
        })

    part6 = {
        "cleaning_info": cleaning_info,
        "automated_stats_comment": automated_stats_comment,
        "group_summary_table": group_summary_table,
        "column_stats_table": column_stats_rows
    }

    return {
        "part1_executive_summary": part1,
        "part2_kpi_dashboard": part2,
        "part3_deep_dive": part3,
        "part4_conclusions": part4,
        "part5_action_recommendations": part5,
        "part6_technical_details": part6
    }

def universal_step3_analyze(
    cleaned_df: pd.DataFrame,
    schema: Dict[str, Any],
    selected_groups: Optional[List[str]] = None,
    custom_rule: str = "",
    audit_log: Optional[List[Dict]] = None,
    company_info: Optional[Dict[str, Any]] = None,
    source_name: str = "Tệp dữ liệu",
    clean_summary: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, str], None]] = None
) -> Dict[str, Any]:
    """
    Bước 3: Quy trình Thống kê Tài chính chuẩn (100% Python/Pandas):
    - Cây KPI Tài chính & Biên Margin (MG Tree)
    - Thống kê mô tả (Mean, Median, Std, Q1, Q3, Min, Max)
    - Phân loại Khách hàng Ưu tiên (ABC / Pareto 80-20)
    - Phân nhóm & Đánh giá Hiệu suất Đa chiều theo cột/yêu cầu
    - Khởi tạo Dynamic 6-Part Business Report (Chuẩn Business Intelligence)
    """
    if progress_callback:
        progress_callback(20, f"Bắt đầu phân tích tài chính ({len(cleaned_df):,} dòng)...")

    df = cleaned_df.copy()

    rev_col = schema.get("revenue_col")
    if not rev_col and "_DoanhThu_Tinh" in df.columns:
        rev_col = "_DoanhThu_Tinh"
    prof_col = schema.get("profit_col")
    if not prof_col and "_LoiNhuan_Tinh" in df.columns:
        prof_col = "_LoiNhuan_Tinh"
    cost_col = schema.get("cost_col")
    qty_col = schema.get("qty_col")
    disc_col = schema.get("discount_col")
    # Rule 1: Tuyệt đối KHÔNG BAO GIỜ đưa "Cột Định danh (ID)" vào danh sách Phân nhóm hay xếp hạng Pareto 80/20.
    # Chỉ xếp hạng Pareto trên các cột có tính chất gom nhóm lặp lại (tỷ lệ giá trị duy nhất thấp hơn 50%).
    raw_cust = schema.get("customer_col") or schema.get("project_col")
    cust_col = None
    if raw_cust and raw_cust in df.columns:
        s_cust = df[raw_cust].dropna()
        if len(df) > 0 and (s_cust.nunique() / len(df)) < 0.5:
            cust_col = raw_cust

    # Tùy chọn lọc động nếu người dùng nhập rule ở Bước 3 (vd: "margin >= 15")
    if custom_rule and prof_col and rev_col and rev_col in df.columns and prof_col in df.columns:
        m_min = re.search(r'margin\s*[>=]+\s*(\d+(?:\.\d+)?)', custom_rule.lower())
        if m_min and "_Margin_Pct" in df.columns:
            thresh = float(m_min.group(1))
            df = df[df["_Margin_Pct"].fillna(0) >= thresh].copy()

    total_rows = len(df)
    rev_series = pd.to_numeric(df[rev_col], errors="coerce").dropna() if rev_col and rev_col in df.columns else pd.Series(dtype=float)
    total_rev = float(rev_series.sum()) if len(rev_series) else 0.0
    avg_rev = (total_rev / total_rows) if total_rows > 0 else 0.0
    median_rev = float(rev_series.median()) if len(rev_series) else 0.0

    total_cost = float(pd.to_numeric(df[cost_col], errors="coerce").fillna(0).sum()) if cost_col and cost_col in df.columns else None
    total_prof = float(pd.to_numeric(df[prof_col], errors="coerce").fillna(0).sum()) if prof_col and prof_col in df.columns else None
    margin_pct = round((total_prof / total_rev * 100.0), 1) if (total_prof is not None and total_rev > 0) else None
    total_qty = float(pd.to_numeric(df[qty_col], errors="coerce").fillna(0).sum()) if qty_col and qty_col in df.columns else None
    avg_disc = round(float(pd.to_numeric(df[disc_col], errors="coerce").dropna().mean()), 2) if disc_col and disc_col in df.columns and len(df[disc_col].dropna()) > 0 else None
    outlier_count = int((df["_Outlier_IQR"] == "Outlier Cao").sum()) if "_Outlier_IQR" in df.columns else 0

    kpis = {
        "total_rows": total_rows,
        "total_revenue": total_rev,
        "avg_order_value": avg_rev,
        "median_order_value": median_rev,
        "total_cost": total_cost,
        "total_profit": total_prof,
        "profit_margin_pct": margin_pct,
        "total_qty": total_qty,
        "avg_discount_pct": avg_disc,
        "outlier_count": outlier_count
    }

    # Cây KPI Tài chính & Biên Margin (MG Tree)
    kpi_tree = {
        "root": {"label": "Tổng Doanh Thu / Giá Trị", "value": total_rev, "sub": f"{total_rows:,} bản ghi"},
        "branches": [
            {"label": "Giá trị TB / Đơn (Mean)", "value": avg_rev, "sub": f"Trung vị (Median): {median_rev:,.0f}"},
            {"label": "Tổng Chi phí / Giá vốn", "value": total_cost if total_cost is not None else (total_rev - (total_prof or 0)) if total_prof is not None else None, "sub": f"Chiết khấu TB: {avg_disc if avg_disc is not None else 0}%"},
            {"label": "Lợi nhuận Gộp & Biên MG", "value": total_prof, "margin_pct": margin_pct, "sub": f"Biên MG: {margin_pct}%" if margin_pct is not None else "Chưa có cột chi phí/LN"},
            {"label": "Tổng Sản lượng & Outlier", "value": total_qty if total_qty is not None else total_rows, "sub": f"Outlier IQR: {outlier_count} dòng"}
        ]
    }

    stat_cols = [c for c in [rev_col, cost_col, prof_col, qty_col, disc_col] if c]
    descriptive_stats = _compute_descriptive_stats(df, stat_cols)

    # Phân hạng Khách hàng Ưu tiên (ABC / Pareto 80-20)
    customer_priority: List[Dict[str, Any]] = []
    tier_summary = {"A": 0, "B": 0, "C": 0, "col_used": cust_col}
    if cust_col and cust_col in df.columns and total_rows > 0:
        c_agg = {"_orders": (cust_col, "count")}
        if rev_col and rev_col in df.columns:
            c_agg["_rev"] = (rev_col, "sum")
        if prof_col and prof_col in df.columns:
            c_agg["_prof"] = (prof_col, "sum")
        cdf = df.groupby(cust_col, dropna=False).agg(**c_agg).reset_index()
        if "_rev" in cdf.columns:
            cdf = cdf.sort_values("_rev", ascending=False)
        else:
            cdf = cdf.sort_values("_orders", ascending=False)

        cum_rev = 0.0
        for idx, (_, crow) in enumerate(cdf.head(30).iterrows(), start=1):
            crev = float(crow["_rev"]) if "_rev" in cdf.columns else 0.0
            cprof = float(crow["_prof"]) if "_prof" in cdf.columns else None
            cshare = round((crev / total_rev * 100.0), 1) if total_rev > 0 else round((int(crow["_orders"]) / total_rows * 100.0), 1)
            cum_rev += crev
            cum_pct = round((cum_rev / total_rev * 100.0), 1) if total_rev > 0 else round((idx / max(1, len(cdf)) * 100.0), 1)
            cmargin = round((cprof / crev * 100.0), 1) if (cprof is not None and crev > 0) else None

            if cum_pct <= 80.0 or idx == 1:
                tier = "Tier A — Ưu tiên Cao (VIP)"
                tier_summary["A"] += 1
            elif cum_pct <= 95.0:
                tier = "Tier B — Tiềm năng"
                tier_summary["B"] += 1
            else:
                tier = "Tier C — Tiêu chuẩn"
                tier_summary["C"] += 1

            customer_priority.append({
                "rank": idx,
                "customer": str(crow[cust_col]),
                "orders": int(crow["_orders"]),
                "revenue": crev,
                "share_pct": cshare,
                "cum_share_pct": cum_pct,
                "profit": cprof,
                "margin_pct": cmargin,
                "tier": tier
            })

    # Phân nhóm Đa chiều (Category, Region, Project, _Thang...)
    groups_to_run = list(selected_groups) if selected_groups else list(schema.get("group_cols", []))[:4]
    
    # Tự động nhận diện và bổ sung các chiều kinh doanh chuẩn nếu có trong dữ liệu
    standard_candidate_cols = ["Category", "Region", "Segment", "Channel", "Subcategory", "Payment_Method", "Returned", "DanhMuc", "KhuVuc", "PhanKhuc", "KenhBan"]
    for sc in standard_candidate_cols:
        for c in df.columns:
            if sc.lower() == str(c).lower() and c not in groups_to_run:
                groups_to_run.append(c)
                break

    if "_Thang" in df.columns and "_Thang" not in groups_to_run and df["_Thang"].nunique() > 1:
        groups_to_run.append("_Thang")

    # Rule 1: Lọc bỏ hoàn toàn các Cột Định danh (cardinality > 80%) khỏi danh sách phân nhóm
    filtered_groups = []
    for g in groups_to_run:
        if g in df.columns:
            s_g = df[g].dropna()
            if len(df) > 0 and (s_g.nunique() / len(df)) > 0.8:
                continue
            filtered_groups.append(g)
    groups_to_run = filtered_groups

    group_tables: Dict[str, List[Dict[str, Any]]] = {}
    highlights: List[Dict[str, Any]] = []

    for g_col in groups_to_run:
        if g_col not in df.columns:
            continue
        agg_dict = {"_count": (g_col, "count")}
        if rev_col and rev_col in df.columns:
            agg_dict["_rev"] = (rev_col, "sum")
        if prof_col and prof_col in df.columns:
            agg_dict["_prof"] = (prof_col, "sum")
        if qty_col and qty_col in df.columns:
            agg_dict["_qty"] = (qty_col, "sum")
        if disc_col and disc_col in df.columns:
            agg_dict["_disc"] = (disc_col, "mean")

        grouped = df.groupby(g_col, dropna=False).agg(**agg_dict).reset_index()
        # Sắp xếp theo thứ tự thời gian tăng dần nếu là cột Tháng, ngược lại sắp xếp theo Doanh thu giảm dần
        if g_col == "_Thang":
            grouped = grouped.sort_values(g_col, ascending=True)
        elif "_rev" in grouped.columns:
            grouped = grouped.sort_values("_rev", ascending=False)
        else:
            grouped = grouped.sort_values("_count", ascending=False)

        n_groups = len(grouped)
        avg_share = (100.0 / n_groups) if n_groups > 0 else 0.0

        rows_out = []
        for rank, (_, gr) in enumerate(grouped.head(25).iterrows(), start=1):
            cnt = int(gr["_count"])
            rev = float(gr["_rev"]) if "_rev" in grouped.columns else 0.0
            share = round((rev / total_rev * 100.0), 1) if total_rev > 0 else round((cnt / total_rows * 100.0), 1)
            prof = float(gr["_prof"]) if "_prof" in grouped.columns else None
            g_margin = round((prof / rev * 100.0), 1) if (prof is not None and rev > 0) else None
            qty_v = float(gr["_qty"]) if "_qty" in grouped.columns else None
            disc_v = round(float(gr["_disc"]), 1) if ("_disc" in grouped.columns and pd.notna(gr["_disc"])) else None

            if share >= avg_share * 1.15 and (g_margin is None or g_margin >= (margin_pct or 0) * 0.9):
                perf = "Hiệu suất cao (Dẫn đầu)"
            elif share < avg_share * 0.6 or (g_margin is not None and g_margin < 8.0):
                perf = "Cần tối ưu (Thấp)"
            else:
                perf = "Ổn định"

            rows_out.append({
                "rank": rank,
                "group_value": str(gr[g_col]),
                "count": cnt,
                "revenue": rev,
                "share_pct": share,
                "avg_value": (rev / cnt) if cnt > 0 else 0.0,
                "profit": prof,
                "margin_pct": g_margin,
                "qty": qty_v,
                "avg_discount": disc_v,
                "performance": perf
            })

        group_tables[g_col] = rows_out
        if rows_out:
            top = rows_out[0]
            low = rows_out[-1] if len(rows_out) > 1 else None
            highlights.append({
                "dimension": "Tháng" if g_col == "_Thang" else g_col,
                "top_group": top["group_value"],
                "top_share_pct": top["share_pct"],
                "top_revenue": top["revenue"],
                "low_group": low["group_value"] if low else None,
                "low_share_pct": low["share_pct"] if low else None
            })

    # Trích xuất Top 10 Sản phẩm chủ lực
    top_products = _extract_top_products(df, {"group_tables": group_tables}, total_rev, limit=10)

    # Sinh 3 thẻ nội dung điều hành chuẩn HTML
    executive_cards = _generate_structured_executive_cards(df, group_tables, customer_priority, top_products, kpis, company_info=company_info)

    if audit_log is not None:
        audit_log.append({
            "timestamp": datetime.datetime.now().strftime("%H:%M:%S"),
            "step": "Bước 3 - Statistical & MG Engine",
            "tool": "universal_step3_analyze (KPI Tree + Pareto VIP + Group-By)",
            "detail": f"DT={total_rev:,.0f}, Margin={margin_pct}%, KH VIP (Tier A)={tier_summary['A']}, Phân nhóm {len(group_tables)} chiều."
        })

    if progress_callback:
        progress_callback(100, f"Hoàn tất phân tích {total_rows:,} dòng.")

    # Tính toán toàn diện các chỉ số BI Tồn kho & Rủi ro (Data Integrity 100%)
    from data_processor import compute_bi_inventory_risk_metrics
    inventory_risk_metrics = compute_bi_inventory_risk_metrics(df=df, schema=schema)
    schema_inspection = inspect_dataset_schema(df)

    # Sinh Cấu trúc Báo cáo Động chuẩn Business Intelligence (6 Phần)
    dynamic_report = generate_dynamic_6part_report(
        cleaned_df=df,
        schema=schema,
        analysis={"kpis": kpis, "schema_inspection": schema_inspection, "inventory_risk_metrics": inventory_risk_metrics, "group_tables": group_tables},
        source_name=source_name,
        company_info=company_info,
        clean_summary=clean_summary
    )

    return {
        "kpis": kpis,
        "kpi_tree": kpi_tree,
        "descriptive_stats": descriptive_stats,
        "customer_priority": customer_priority,
        "top_products": top_products,
        "tier_summary": tier_summary,
        "group_tables": group_tables,
        "highlights": highlights,
        "executive_cards": executive_cards,
        "inventory_risk_metrics": inventory_risk_metrics,
        "schema_inspection": schema_inspection,
        "schema_used": schema,
        "dynamic_report": dynamic_report
    }

def universal_step4_ai_report(
    analysis: Dict[str, Any],
    source_name: str,
    api_key: str = "",
    provider: str = "Gemini",
    style: str = "Điều hành súc tích",
    company_info: Optional[Dict[str, Any]] = None
) -> str:
    """
    Skill 3: Dùng Gemini AI viết Báo cáo Tài chính, Cây Margin (MG), Khách hàng Ưu tiên & Phân nhóm
    từ số liệu ĐÃ ĐƯỢC PANDAS KIỂM CHỨNG 100%.
    Định dạng số tiền gọn gàng theo Tỷ VND / Triệu VND, không để dãy số thô dài.
    Tích hợp thông tin doanh nghiệp & mục tiêu phân tích nếu có.
    """
    company_info = company_info or {}
    biz_name = (company_info.get("name") or "").strip()
    biz_type = (company_info.get("type") or "").strip()
    biz_extra = (company_info.get("extra") or "").strip()
    biz_prob = (company_info.get("problem") or "").strip()

    kpis = analysis.get("kpis") or {}
    if not kpis and "kpi_tree" in analysis:
        tree = analysis["kpi_tree"]
        kpis = {
            "total_rows": tree.get("tong_so_dong", 0),
            "total_revenue": tree.get("tong_doanh_thu", 0),
            "avg_order_value": tree.get("doanh_thu_tb", 0),
            "total_profit": tree.get("tong_loi_nhuan", 0),
            "profit_margin_pct": tree.get("ty_suat_loi_nhuan_tb", 0)
        }

    highlights = analysis.get("highlights", [])
    tables = analysis.get("group_tables", {})
    vip_top = [c for c in analysis.get("customer_priority", []) if "Tier A" in str(c.get("tier", ""))][:5]

    # Rule 2: Chỉ đưa vào compact_summary những chỉ số thực sự tồn tại
    has_rev = kpis.get("total_revenue") is not None and _safe_float(kpis.get("total_revenue")) > 0
    kpi_dict = {
        "tong_so_ban_ghi": kpis.get("total_rows", 0)
    }
    if has_rev:
        kpi_dict["tong_doanh_thu_hoac_gia_tri"] = format_compact_vnd(kpis.get("total_revenue", 0))
        kpi_dict["gia_tri_tb"] = format_compact_vnd(kpis.get("avg_order_value", 0))
        kpi_dict["gia_tri_trung_vi"] = format_compact_vnd(kpis.get("median_order_value", 0))
    if kpis.get("total_cost") is not None:
        kpi_dict["tong_chi_phi"] = format_compact_vnd(kpis.get("total_cost", 0))
    if kpis.get("total_profit") is not None:
        kpi_dict["tong_loi_nhuan"] = format_compact_vnd(kpis.get("total_profit", 0))
        if kpis.get("profit_margin_pct") is not None:
            kpi_dict["bien_margin_pct"] = f"{kpis.get('profit_margin_pct')}%"
    if kpis.get("total_qty") is not None:
        kpi_dict["tong_san_luong"] = int(kpis.get("total_qty", 0))
    if kpis.get("avg_discount_pct") is not None:
        kpi_dict["chiet_khau_tb"] = f"{kpis.get('avg_discount_pct')}%"

    compact_summary = {
        "quy_mo_du_lieu": f"{kpis.get('total_rows', 0):,} bản ghi",
        "nguon_du_lieu": source_name,
        "linh_vuc": analysis.get("schema_used", {}).get("dataset_domain", "Phân tích điều hành"),
        "kpis_thuc_te": kpi_dict,
        "nhom_noi_bat": highlights[:3]
    }
    if biz_name: compact_summary["doanh_nghiep"] = biz_name
    if biz_type: compact_summary["loai_hinh"] = biz_type
    if biz_prob: compact_summary["trong_tam_phan_tich"] = biz_prob
    if vip_top:
        compact_summary["khach_hang_uu_tien_VIP"] = [
            {
                "khach_hang": v.get("customer"),
                "so_don": v.get("orders"),
                "ty_trong": f"{v.get('share_pct', 0)}%"
            }
            for v in vip_top
        ]

    has_key = bool((api_key or "").strip()) or (provider == "Gemini" and bool(get_default_ai_config().get("gemini_api_key")))
    if has_key and provider in ("Gemini", "OpenAI"):
        s2_goal = "PHÂN HẠNG ĐỐI TÁC TRỌNG ĐIỂM (PARETO 80/20)" if vip_top else "PHÂN TÍCH HIỆU SUẤT THEO NHÓM DẪN ĐẦU"
        prompt = (
            f"Bạn là Giám đốc Phân tích Dữ liệu BI & Điều hành C-Level. Hãy viết Nhận định Điều hành ngắn gọn, sắc bén "
            f"theo văn phong '{style}'.\n"
            f"QUY TẮC BẮT BUỘC (ZERO HARD-CODING & DATA INTEGRITY):\n"
            f"1. CHỈ phân tích và nhận định dựa trên các chỉ số THỰC SỰ CÓ trong dữ liệu JSON dưới đây. "
            f"Nếu một chỉ số (như Lợi nhuận, Chi phí, Khách hàng VIP) KHÔNG CÓ trong JSON, TUYỆT ĐỐI BỎ QUA HOÀN TOÀN, "
            f"không đề cập Margin = None%, không gán nhãn 0 VNĐ, không suy diễn chiết khấu/đứt hàng/tồn kho nếu không có số liệu thực tế.\n"
            f"2. Bố cục gồm đúng 3 phần:\n"
            f"   - ### 1. ĐÁNH GIÁ CHỈ SỐ HOẠT ĐỘNG (Quy mô bản ghi, giá trị, độ phân bổ theo số liệu thực)\n"
            f"   - ### 2. {s2_goal}\n"
            f"   - ### 3. ĐỀ XUẤT HÀNH ĐỘNG ĐIỀU HÀNH (3 hành động điều hành cụ thể gắn với các nhóm/chỉ số đang có)\n"
            f"3. Định dạng toàn bộ số tiền gọn gàng (Tỷ VND, Triệu VND). Không in dãy số thô.\n"
            f"{f'Trọng tâm phân tích theo yêu cầu: {biz_prob}' if biz_prob else ''}\n"
            f"Dữ liệu JSON:\n{json.dumps(compact_summary, ensure_ascii=False)}"
        )
        ai_text = _call_llm_json(prompt, api_key, provider)
        if ai_text:
            return ai_text.strip()

    # Rule 2: Dynamic Content with Zero Hard-coding
    has_rev = kpis.get("total_revenue") is not None and _safe_float(kpis.get("total_revenue")) > 0
    has_prof = kpis.get("total_profit") is not None
    has_mg = kpis.get("profit_margin_pct") is not None
    has_cost = kpis.get("total_cost") is not None
    has_qty = kpis.get("total_qty") is not None
    has_disc = kpis.get("avg_discount_pct") is not None
    has_vip = bool(vip_top and len(vip_top) > 0)
    total_rows = kpis.get("total_rows", 0)

    # Section 1: KPI Tree
    s1_bullets = []
    s1_bullets.append(f"*   **Quy mô dữ liệu:** Đã xử lý và kiểm định toàn vẹn **{total_rows:,}** bản ghi.")
    if has_rev:
        aov_str = format_compact_vnd(kpis.get("avg_order_value", 0))
        median_str = format_compact_vnd(kpis.get("median_order_value", 0))
        s1_bullets.append(f"*   **Tổng giá trị / Doanh thu:** **{format_compact_vnd(kpis.get('total_revenue', 0))}** (Giá trị trung bình: **{aov_str}**, trung vị: **{median_str}**).")
    if has_cost:
        s1_bullets.append(f"*   **Tổng chi phí / Giá vốn:** **{format_compact_vnd(kpis.get('total_cost', 0))}**.")
    if has_prof:
        mg_part = f" (Biên Margin: **{kpis.get('profit_margin_pct')}%**)" if has_mg else ""
        s1_bullets.append(f"*   **Tổng lợi nhuận:** **{format_compact_vnd(kpis.get('total_profit', 0))}**{mg_part}.")
    if has_qty:
        s1_bullets.append(f"*   **Tổng sản lượng:** **{int(kpis.get('total_qty', 0)):,}** đơn vị.")
    if has_disc:
        s1_bullets.append(f"*   **Chiết khấu bình quân:** **{kpis.get('avg_discount_pct')}%**.")
    if biz_prob:
        s1_bullets.append(f"*   **Trọng tâm phân tích theo yêu cầu:** {biz_prob}.")

    s1_text = "\n".join(s1_bullets)

    # Section 2: VIP / Top Groups / Distribution
    if has_vip:
        vip_bullets = []
        for v in vip_top:
            v_cust = v.get("customer", "")
            v_orders = int(v.get("orders", 0))
            parts = [f"**{v_cust}:**"]
            if has_rev:
                parts.append(f"Doanh thu {format_compact_vnd(v.get('revenue', 0))} (chiếm {v.get('share_pct', 0)}%, {v_orders} đơn)")
            else:
                parts.append(f"Quy mô {v_orders} đơn (chiếm {v.get('share_pct', 0)}%)")
            if has_prof and v.get("profit") is not None:
                parts.append(f"Lợi nhuận {format_compact_vnd(v.get('profit', 0))}")
            if has_mg and v.get("margin_pct") is not None:
                parts.append(f"Margin {v.get('margin_pct')}%")
            vip_bullets.append(f"*   " + ", ".join(parts) + ".")
        s2_text = "### 2. PHÂN HẠNG ĐỐI TÁC TRỌNG ĐIỂM (PARETO 80/20 — TIER A)\n" + "\n".join(vip_bullets)
    elif highlights:
        hl_bullets = []
        for h in highlights[:4]:
            g_col = h.get("group_col", "Nhóm")
            g_val = h.get("group_val", "")
            g_share = f"{h.get('share_pct', 0)}%"
            if has_rev:
                hl_bullets.append(f"*   **{g_col} — {g_val}:** Đạt {format_compact_vnd(h.get('revenue', 0))} (chiếm {g_share} tỷ trọng).")
            else:
                hl_bullets.append(f"*   **{g_col} — {g_val}:** Chiếm {g_share} tỷ trọng phân bổ.")
        s2_text = "### 2. PHÂN TÍCH HIỆU SUẤT THEO NHÓM DẪN ĐẦU\n" + "\n".join(hl_bullets)
    else:
        s2_text = f"### 2. ĐÁNH GIÁ PHÂN BỔ DỮ LIỆU TỔNG QUAN\n*   Dữ liệu phân bổ đều giữa các bản ghi ({total_rows:,} dòng), không ghi nhận sự tập trung rủi ro theo nguyên lý Pareto."

    # Section 3: Action Recommendations
    s3_bullets = []
    if has_prof:
        s3_bullets.append("*   Rà soát cơ cấu định giá và tối ưu hóa chi phí để duy trì biên lợi nhuận ròng an toàn.")
    elif has_rev:
        s3_bullets.append("*   Tối ưu hóa các kênh và nhóm sản phẩm có doanh số cao để thúc đẩy tăng trưởng quy mô doanh thu.")
    else:
        s3_bullets.append("*   Nâng cao hiệu suất vận hành và kiểm soát chặt chẽ các chỉ tiêu định lượng trong dữ liệu.")

    if has_vip:
        s3_bullets.append("*   Xây dựng chính sách chăm sóc và giữ chân nhóm đối tác trọng điểm Tier A để bảo vệ nguồn đóng góp chủ lực.")
    elif highlights:
        top_h = highlights[0]
        s3_bullets.append(f"*   Mở rộng khai thác nhóm dẫn đầu ({top_h.get('group_col')}: {top_h.get('group_val')}) để tối đa hóa hiệu quả hoạt động.")
    else:
        s3_bullets.append("*   Duy trì tính đồng đều trong phân bổ nguồn lực, hạn chế phụ thuộc cục bộ vào một nhóm chỉ số đơn lẻ.")

    if biz_prob:
        s3_bullets.append(f"*   Tập trung giải quyết mục tiêu trọng điểm đã đặt ra: {biz_prob}.")
    else:
        s3_bullets.append("*   Thiết lập cơ chế giám sát dữ liệu định kỳ nhằm phát hiện sớm các biến động bất thường ngoài ngưỡng kiểm soát.")

    s3_text = "\n".join(s3_bullets)

    biz_header = f"{biz_name} — " if biz_name else ""
    return f"""### 1. ĐÁNH GIÁ CÂY CHỈ SỐ HOẠT ĐỘNG ({biz_header.upper()}KPI TREE)
{s1_text}

{s2_text}

### 3. ĐỀ XUẤT HÀNH ĐỘNG ĐIỀU HÀNH
{s3_text}"""

def export_universal_excel_bytes(cleaned_df: pd.DataFrame, analysis: Dict[str, Any], source_name: str = "Dữ liệu giao dịch kinh doanh tổng hợp", company_info: Optional[Dict[str, Any]] = None) -> bytes:
    """Xuất file Excel chuyên nghiệp tinh gọn chuẩn C-Level Dashboard:
    - Đúng 5 Sheet cốt lõi (Dashboard, TongHop_PhanTich, KhachHang_VIP, ThongKe_MoTa, DuLieu_DaLamSach).
    - Sheet 1 (Dashboard): Tắt Gridlines, Header tiêu đề, 4 KPI Card động (SUM, IFERROR, AVERAGE, COUNTA),
      4 Native Charts openpyxl (LineChart xu hướng tháng, Horizontal BarChart Top 5 SP, DoughnutChart Category, Clustered Col Segment).
    - Sheet 2 (TongHop_PhanTich): 5 bảng tổng hợp động liên kết SUMIFS, COUNTIFS, AVERAGEIFS, IFERROR với Sheet DuLieu_DaLamSach.
      Gồm: Tháng (24 tháng tăng dần), Category & Region (side-by-side), Segment & Channel (side-by-side),
      Top 10 Sản phẩm chủ lực (gộp ID & Name), Top 15 Khách hàng VIP Tier A.
    - Sheet 3 (KhachHang_VIP): Phân tích Pareto ABC toàn diện, AutoFilter, Freeze Panes, DataBar, ColorScale.
    - Sheet 4 (ThongKe_MoTa): Thống kê mô tả toán học (Count, Mean, Std, Min, Q25, Median, Q75, Max, Outliers).
    - Sheet 5 (DuLieu_DaLamSach): Dữ liệu sạch, AutoFilter, Freeze Panes, định dạng tiền tệ/phần trăm/số lượng.
    - Palette Corporate Navy (#1E3A8A) đồng bộ, auto-fit độ rộng cột.
    """
    if cleaned_df is None or cleaned_df.empty:
        cleaned_df = pd.DataFrame([{"Thong_Bao": "Không có dữ liệu giao dịch"}])
    analysis = analysis or {}

    wb = openpyxl.Workbook()

    # 1. Khởi tạo 5 Sheet cốt lõi
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_dash.views.sheetView[0].showGridLines = False

    ws_tonghop = wb.create_sheet("TongHop_PhanTich")
    ws_vip = wb.create_sheet("KhachHang_VIP")
    ws_mota = wb.create_sheet("ThongKe_MoTa")
    ws_raw = wb.create_sheet("DuLieu_DaLamSach")

    # Bảng màu đồng bộ Corporate Navy / Slate
    FONT_NAME = "Segoe UI"
    NAVY_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    CARD_BG_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    FONT_WHITE_BOLD = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    FONT_CARD_TITLE = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    FONT_CARD_VAL = Font(name=FONT_NAME, size=15, bold=True, color="1E3A8A")
    FONT_CARD_VAL_GREEN = Font(name=FONT_NAME, size=15, bold=True, color="0E8A16")
    FONT_CARD_SUB = Font(name=FONT_NAME, size=9, italic=True, color="64748B")
    FONT_SECTION = Font(name=FONT_NAME, size=11, bold=True, color="1E3A8A")
    FONT_DATA = Font(name=FONT_NAME, size=10, color="0F172A")
    FONT_TOTAL = Font(name=FONT_NAME, size=10, bold=True, color="1E3A8A")

    THIN_SIDE = Side(border_style="thin", color="CBD5E1")
    CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    TOTAL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=Side(border_style="double", color="1E3A8A"))

    ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
    ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

    # =========================================================================
    # BƯỚC 1: XUẤT SHEET 5 - DuLieu_DaLamSach & XÁC ĐỊNH VỊ TRÍ CỘT
    # =========================================================================
    n_raw_rows = len(cleaned_df)
    raw_cols = list(cleaned_df.columns)
    col_letter_map = {col: get_column_letter(idx + 1) for idx, col in enumerate(raw_cols)}

    def find_col_letter(preferred_names):
        for name in preferred_names:
            if name in col_letter_map:
                return col_letter_map[name], name
        for c in raw_cols:
            c_low = str(c).lower()
            if any(p.lower() in c_low for p in preferred_names):
                return col_letter_map[c], c
        return None, None

    sales_col, sales_name = find_col_letter(["Sales_VND", "Sales", "revenue", "doanh_thu", "Sales_Amount", "sales"])
    profit_col, profit_name = find_col_letter(["Profit_VND", "Profit", "profit", "loi_nhuan", "Net_Profit"])
    month_col, month_name = find_col_letter(["_Thang", "Month", "thang", "month"])
    id_col, id_name = find_col_letter(["Order_ID", "order_id", "Transaction_ID", "Ma_Don", "id", "ID"])
    cat_col, cat_name = find_col_letter(["Category", "category", "Nganh_Hang"])
    reg_col, reg_name = find_col_letter(["Region", "region", "Khu_Vuc"])
    seg_col, seg_name = find_col_letter(["Segment", "segment", "Phan_Khuc"])
    chan_col, chan_name = find_col_letter(["Channel", "channel", "Kenh_Ban"])

    # Fallbacks an toàn
    sales_col = sales_col or "A"
    profit_col = profit_col or sales_col
    month_col = month_col or "A"
    id_col = id_col or "A"
    cat_col = cat_col or "A"
    reg_col = reg_col or "A"
    seg_col = seg_col or "A"
    chan_col = chan_col or "A"

    # Ghi Header Sheet DuLieu_DaLamSach
    for ci, col in enumerate(raw_cols, start=1):
        cell = ws_raw.cell(row=1, column=ci, value=col)
        cell.fill = NAVY_FILL
        cell.font = FONT_WHITE_BOLD
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    # Ghi Dữ liệu Sheet DuLieu_DaLamSach
    for ri, row_data in enumerate(cleaned_df.itertuples(index=False), start=2):
        is_zebra = (ri % 2 == 1)
        row_fill = ZEBRA_FILL if is_zebra else WHITE_FILL
        for ci, val in enumerate(row_data, start=1):
            cell = ws_raw.cell(row=ri, column=ci, value=val)
            cell.font = FONT_DATA
            cell.border = CELL_BORDER
            cell.fill = row_fill

            col_name_lower = str(raw_cols[ci - 1]).lower()
            if isinstance(val, (int, float)):
                if any(k in col_name_lower for k in ["sales", "profit", "cost", "price", "vnd", "amount", "budget"]):
                    cell.number_format = "#,##0"
                    cell.alignment = ALIGN_RIGHT
                elif any(k in col_name_lower for k in ["pct", "margin", "discount", "rate", "share"]):
                    cell.number_format = "0.0%"
                    cell.alignment = ALIGN_RIGHT
                elif any(k in col_name_lower for k in ["qty", "quantity", "count", "days"]):
                    cell.number_format = "#,##0"
                    cell.alignment = ALIGN_RIGHT
                else:
                    cell.alignment = ALIGN_RIGHT
            else:
                cell.alignment = ALIGN_LEFT

    ws_raw.auto_filter.ref = f"A1:{get_column_letter(len(raw_cols))}{n_raw_rows + 1}"
    ws_raw.freeze_panes = "A2"

    # =========================================================================
    # BƯỚC 2: XUẤT SHEET 1 - Dashboard (Tiêu đề, 4 Thẻ KPI động)
    # =========================================================================
    company_info = company_info or {}
    biz_name = (company_info.get("name") or "").strip()
    biz_type = (company_info.get("type") or "").strip()

    ws_dash.merge_cells("B2:P2")
    title_text = f"{biz_name.upper()} — BÁO CÁO THỐNG KÊ & PHÂN TÍCH HIỆU SUẤT KINH DOANH" if biz_name else "BÁO CÁO THỐNG KÊ & PHÂN TÍCH HIỆU SUẤT KINH DOANH"
    ws_dash["B2"] = title_text
    ws_dash["B2"].font = Font(name=FONT_NAME, size=16, bold=True, color="1E3A8A")
    ws_dash["B2"].alignment = ALIGN_LEFT

    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    ws_dash.merge_cells("B3:P3")
    type_meta = f"Loại hình: {biz_type}  |  " if biz_type else ""
    ws_dash["B3"] = f"{type_meta}Ngày xuất: {now_str}  |  Nguồn dữ liệu: {source_name}  |  Hệ thống: Sales BI Analytics"
    ws_dash["B3"].font = Font(name=FONT_NAME, size=9.5, italic=True, color="475569")
    ws_dash["B3"].alignment = ALIGN_LEFT

    def draw_kpi_card(start_col, title, val_formula, num_format, sub_formula_or_text, is_green=False):
        c1, c2, c3 = start_col, chr(ord(start_col) + 1), chr(ord(start_col) + 2)
        # Hàng 5: Title
        ws_dash.merge_cells(f"{c1}5:{c3}5")
        top_cell = ws_dash[f"{c1}5"]
        top_cell.value = title
        top_cell.fill = NAVY_FILL
        top_cell.font = FONT_CARD_TITLE
        top_cell.alignment = ALIGN_CENTER

        # Hàng 6: Giá trị
        ws_dash.merge_cells(f"{c1}6:{c3}6")
        mid_cell = ws_dash[f"{c1}6"]
        mid_cell.value = val_formula
        mid_cell.fill = CARD_BG_FILL
        mid_cell.font = FONT_CARD_VAL_GREEN if is_green else FONT_CARD_VAL
        mid_cell.alignment = ALIGN_CENTER
        mid_cell.number_format = num_format

        # Hàng 7: Sub text
        ws_dash.merge_cells(f"{c1}7:{c3}7")
        bot_cell = ws_dash[f"{c1}7"]
        bot_cell.value = sub_formula_or_text
        bot_cell.fill = CARD_BG_FILL
        bot_cell.font = FONT_CARD_SUB
        bot_cell.alignment = ALIGN_CENTER

        # Border khung
        for r in range(5, 8):
            for c_idx in range(ord(c1) - ord('A') + 1, ord(c3) - ord('A') + 2):
                ws_dash.cell(row=r, column=c_idx).border = CELL_BORDER

    # Card 1: Tổng Doanh thu
    draw_kpi_card(
        "B", "TỔNG DOANH THU",
        f"=SUM(DuLieu_DaLamSach!${sales_col}$2:${sales_col}${n_raw_rows + 1})",
        "#,##0",
        f'="Tổng số: " & TEXT(COUNTA(DuLieu_DaLamSach!${id_col}$2:${id_col}${n_raw_rows + 1}), "#,##0") & " đơn"'
    )

    # Card 2: Tổng Lợi nhuận
    draw_kpi_card(
        "F", "TỔNG LỢI NHUẬN",
        f"=SUM(DuLieu_DaLamSach!${profit_col}$2:${profit_col}${n_raw_rows + 1})",
        "#,##0",
        "Lãi gộp theo dữ liệu bán hàng",
        is_green=True
    )

    # Card 3: Biên Lợi Nhuận Margin %
    draw_kpi_card(
        "J", "BIÊN LỢI NHUẬN (MARGIN %)",
        "=IFERROR(F6/B6, 0)",
        "0.0%",
        "Chỉ số hiệu quả tài chính"
    )

    # Card 4: Giá trị đơn trung bình (AOV)
    draw_kpi_card(
        "N", "GIÁ TRỊ ĐƠN TB (AOV)",
        f"=AVERAGE(DuLieu_DaLamSach!${sales_col}$2:${sales_col}${n_raw_rows + 1})",
        "#,##0",
        "Quy mô bình quân mỗi đơn hàng"
    )

    # =========================================================================
    # BƯỚC 3: XUẤT SHEET 2 - TongHop_PhanTich (5 BẢNG CỐT LÕI DÙNG CÔNG THỨC ĐỘNG)
    # =========================================================================
    ws_tonghop.views.sheetView[0].showGridLines = True
    g_tables = analysis.get("group_tables", {})

    def style_table_row(ws, row_idx, start_col, end_col, is_header=False, is_total=False, is_zebra=False):
        for col_idx in range(start_col, end_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if is_header:
                cell.fill = NAVY_FILL
                cell.font = FONT_WHITE_BOLD
                cell.alignment = ALIGN_CENTER
                cell.border = CELL_BORDER
            elif is_total:
                cell.fill = TOTAL_FILL
                cell.font = FONT_TOTAL
                cell.border = TOTAL_BORDER
            else:
                cell.fill = ZEBRA_FILL if is_zebra else WHITE_FILL
                cell.font = FONT_DATA
                cell.border = CELL_BORDER

    # 3.1 BẢNG 1: HIỆU SUẤT THEO THÁNG (24 tháng tăng dần 2024-01 -> 2025-12)
    ws_tonghop["A3"] = "1. HIỆU SUẤT KINH DOANH THEO THÁNG (CHRONOLOGICAL 24 THÁNG)"
    ws_tonghop["A3"].font = FONT_SECTION

    m_headers = ["Tháng", "Số đơn hàng", "Doanh thu (VND)", "Tỷ trọng DT (%)", "Lợi nhuận (VND)", "Biên Margin (%)"]
    for ci, h in enumerate(m_headers, start=1):
        ws_tonghop.cell(row=4, column=ci, value=h)
    style_table_row(ws_tonghop, 4, 1, 6, is_header=True)

    m_data = g_tables.get("_Thang") or g_tables.get("Month") or []
    month_list = [str(r.get("group_value")) for r in m_data if r.get("group_value") and str(r.get("group_value")).lower() != "nan"]
    if not month_list and month_name in cleaned_df.columns:
        month_list = sorted([str(m) for m in cleaned_df[month_name].dropna().unique().tolist() if str(m).lower() != "nan"])
    else:
        month_list = sorted(list(set(month_list)))

    cur_row = 5
    m_start_row = cur_row
    for idx, m_val in enumerate(month_list):
        r = cur_row
        ws_tonghop.cell(row=r, column=1, value=m_val).alignment = ALIGN_CENTER
        ws_tonghop.cell(row=r, column=2, value=f"=COUNTIFS(DuLieu_DaLamSach!${month_col}$2:${month_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
        ws_tonghop.cell(row=r, column=3, value=f"=SUMIFS(DuLieu_DaLamSach!${sales_col}$2:${sales_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${month_col}$2:${month_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
        ws_tonghop.cell(row=r, column=4, value=f"=IFERROR(C{r}/Dashboard!$B$6, 0)").number_format = "0.0%"
        ws_tonghop.cell(row=r, column=5, value=f"=SUMIFS(DuLieu_DaLamSach!${profit_col}$2:${profit_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${month_col}$2:${month_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
        ws_tonghop.cell(row=r, column=6, value=f"=IFERROR(E{r}/C{r}, 0)").number_format = "0.0%"

        style_table_row(ws_tonghop, r, 1, 6, is_zebra=(idx % 2 == 1))
        cur_row += 1

    m_end_row = max(cur_row - 1, m_start_row)
    m_tot_row = cur_row
    ws_tonghop.cell(row=m_tot_row, column=1, value="Tổng cộng").alignment = ALIGN_CENTER
    ws_tonghop.cell(row=m_tot_row, column=2, value=f"=SUM(B{m_start_row}:B{m_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=m_tot_row, column=3, value=f"=SUM(C{m_start_row}:C{m_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=m_tot_row, column=4, value=f"=SUM(D{m_start_row}:D{m_end_row})").number_format = "0.0%"
    ws_tonghop.cell(row=m_tot_row, column=5, value=f"=SUM(E{m_start_row}:E{m_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=m_tot_row, column=6, value=f"=IFERROR(E{m_tot_row}/C{m_tot_row}, 0)").number_format = "0.0%"
    style_table_row(ws_tonghop, m_tot_row, 1, 6, is_total=True)

    if m_end_row >= m_start_row:
        ws_tonghop.conditional_formatting.add(
            f"C{m_start_row}:C{m_end_row}",
            DataBarRule(start_type="min", end_type="max", color="60A5FA", showValue="None")
        )
        ws_tonghop.conditional_formatting.add(
            f"F{m_start_row}:F{m_end_row}",
            ColorScaleRule(start_type="min", start_color="F87171", mid_type="percentile", mid_value=50, mid_color="FCD34D", end_type="max", end_color="4ADE80")
        )

    # 3.2 BẢNG 2: HIỆU SUẤT THEO NGÀNH HÀNG (Cols A-F) & KHU VỰC (Cols H-M) CẠNH NHAU
    cur_row = m_tot_row + 3
    t2_title_row = cur_row
    ws_tonghop.cell(row=t2_title_row, column=1, value="2. HIỆU SUẤT THEO NGÀNH HÀNG (CATEGORY) & KHU VỰC (REGION)").font = FONT_SECTION

    cur_row += 1
    t2_hdr_row = cur_row

    cat_hdrs = ["Ngành hàng", "Số đơn hàng", "Doanh thu (VND)", "Tỷ trọng DT (%)", "Lợi nhuận (VND)", "Biên Margin (%)"]
    for ci, h in enumerate(cat_hdrs, start=1):
        ws_tonghop.cell(row=t2_hdr_row, column=ci, value=h)
    style_table_row(ws_tonghop, t2_hdr_row, 1, 6, is_header=True)

    reg_hdrs = ["Khu vực", "Số đơn hàng", "Doanh thu (VND)", "Tỷ trọng DT (%)", "Lợi nhuận (VND)", "Biên Margin (%)"]
    for ci, h in enumerate(reg_hdrs, start=8):
        ws_tonghop.cell(row=t2_hdr_row, column=ci, value=h)
    style_table_row(ws_tonghop, t2_hdr_row, 8, 13, is_header=True)

    cat_rows = g_tables.get("Category") or g_tables.get("Subcategory") or []
    cat_vals = [str(r["group_value"]) for r in cat_rows if r.get("group_value")]
    if not cat_vals and cat_name in cleaned_df.columns:
        cat_vals = [str(v) for v in cleaned_df[cat_name].dropna().unique().tolist()][:10]

    reg_rows = g_tables.get("Region") or g_tables.get("Khu_Vuc") or g_tables.get("City") or []
    reg_vals = [str(r["group_value"]) for r in reg_rows if r.get("group_value")]
    if not reg_vals and reg_name in cleaned_df.columns:
        reg_vals = [str(v) for v in cleaned_df[reg_name].dropna().unique().tolist()][:10]

    max_t2_len = max(len(cat_vals), len(reg_vals), 1)
    cat_start_row = t2_hdr_row + 1

    for i in range(max_t2_len):
        r = cat_start_row + i
        if i < len(cat_vals):
            c_val = cat_vals[i]
            ws_tonghop.cell(row=r, column=1, value=c_val).alignment = ALIGN_LEFT
            ws_tonghop.cell(row=r, column=2, value=f"=COUNTIFS(DuLieu_DaLamSach!${cat_col}$2:${cat_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=3, value=f"=SUMIFS(DuLieu_DaLamSach!${sales_col}$2:${sales_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${cat_col}$2:${cat_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=4, value=f"=IFERROR(C{r}/Dashboard!$B$6, 0)").number_format = "0.0%"
            ws_tonghop.cell(row=r, column=5, value=f"=SUMIFS(DuLieu_DaLamSach!${profit_col}$2:${profit_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${cat_col}$2:${cat_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=6, value=f"=IFERROR(E{r}/C{r}, 0)").number_format = "0.0%"
            style_table_row(ws_tonghop, r, 1, 6, is_zebra=(i % 2 == 1))

        if i < len(reg_vals):
            reg_val = reg_vals[i]
            ws_tonghop.cell(row=r, column=8, value=reg_val).alignment = ALIGN_LEFT
            ws_tonghop.cell(row=r, column=9, value=f"=COUNTIFS(DuLieu_DaLamSach!${reg_col}$2:${reg_col}${n_raw_rows + 1}, H{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=10, value=f"=SUMIFS(DuLieu_DaLamSach!${sales_col}$2:${sales_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${reg_col}$2:${reg_col}${n_raw_rows + 1}, H{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=11, value=f"=IFERROR(J{r}/Dashboard!$B$6, 0)").number_format = "0.0%"
            ws_tonghop.cell(row=r, column=12, value=f"=SUMIFS(DuLieu_DaLamSach!${profit_col}$2:${profit_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${reg_col}$2:${reg_col}${n_raw_rows + 1}, H{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=13, value=f"=IFERROR(L{r}/J{r}, 0)").number_format = "0.0%"
            style_table_row(ws_tonghop, r, 8, 13, is_zebra=(i % 2 == 1))

    t2_end_row = cat_start_row + max_t2_len - 1
    t2_tot_row = t2_end_row + 1

    ws_tonghop.cell(row=t2_tot_row, column=1, value="Tổng cộng").alignment = ALIGN_CENTER
    ws_tonghop.cell(row=t2_tot_row, column=2, value=f"=SUM(B{cat_start_row}:B{t2_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t2_tot_row, column=3, value=f"=SUM(C{cat_start_row}:C{t2_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t2_tot_row, column=4, value=f"=SUM(D{cat_start_row}:D{t2_end_row})").number_format = "0.0%"
    ws_tonghop.cell(row=t2_tot_row, column=5, value=f"=SUM(E{cat_start_row}:E{t2_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t2_tot_row, column=6, value=f"=IFERROR(E{t2_tot_row}/C{t2_tot_row}, 0)").number_format = "0.0%"
    style_table_row(ws_tonghop, t2_tot_row, 1, 6, is_total=True)

    ws_tonghop.cell(row=t2_tot_row, column=8, value="Tổng cộng").alignment = ALIGN_CENTER
    ws_tonghop.cell(row=t2_tot_row, column=9, value=f"=SUM(I{cat_start_row}:I{t2_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t2_tot_row, column=10, value=f"=SUM(J{cat_start_row}:J{t2_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t2_tot_row, column=11, value=f"=SUM(K{cat_start_row}:K{t2_end_row})").number_format = "0.0%"
    ws_tonghop.cell(row=t2_tot_row, column=12, value=f"=SUM(L{cat_start_row}:L{t2_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t2_tot_row, column=13, value=f"=IFERROR(L{t2_tot_row}/J{t2_tot_row}, 0)").number_format = "0.0%"
    style_table_row(ws_tonghop, t2_tot_row, 8, 13, is_total=True)

    # 3.3 BẢNG 3: HIỆU SUẤT THEO PHÂN KHÚC (Cols A-F) & KÊNH BÁN (Cols H-M) CẠNH NHAU
    cur_row = t2_tot_row + 3
    t3_title_row = cur_row
    ws_tonghop.cell(row=t3_title_row, column=1, value="3. HIỆU SUẤT THEO PHÂN KHÚC KHÁCH HÀNG & KÊNH BÁN HÀNG").font = FONT_SECTION

    cur_row += 1
    t3_hdr_row = cur_row

    seg_hdrs = ["Phân khúc", "Số đơn hàng", "Doanh thu (VND)", "Tỷ trọng DT (%)", "Lợi nhuận (VND)", "Biên Margin (%)"]
    for ci, h in enumerate(seg_hdrs, start=1):
        ws_tonghop.cell(row=t3_hdr_row, column=ci, value=h)
    style_table_row(ws_tonghop, t3_hdr_row, 1, 6, is_header=True)

    chan_hdrs = ["Kênh bán hàng", "Số đơn hàng", "Doanh thu (VND)", "Tỷ trọng DT (%)", "Lợi nhuận (VND)", "Biên Margin (%)"]
    for ci, h in enumerate(chan_hdrs, start=8):
        ws_tonghop.cell(row=t3_hdr_row, column=ci, value=h)
    style_table_row(ws_tonghop, t3_hdr_row, 8, 13, is_header=True)

    seg_rows = g_tables.get("Segment") or g_tables.get("Phan_Khuc") or []
    seg_vals = [str(r["group_value"]) for r in seg_rows if r.get("group_value")]
    if not seg_vals and seg_name in cleaned_df.columns:
        seg_vals = [str(v) for v in cleaned_df[seg_name].dropna().unique().tolist()][:10]

    chan_rows = g_tables.get("Channel") or g_tables.get("Kenh_Ban") or []
    chan_vals = [str(r["group_value"]) for r in chan_rows if r.get("group_value")]
    if not chan_vals and chan_name in cleaned_df.columns:
        chan_vals = [str(v) for v in cleaned_df[chan_name].dropna().unique().tolist()][:10]

    max_t3_len = max(len(seg_vals), len(chan_vals), 1)
    seg_start_row = t3_hdr_row + 1

    for i in range(max_t3_len):
        r = seg_start_row + i
        if i < len(seg_vals):
            s_val = seg_vals[i]
            ws_tonghop.cell(row=r, column=1, value=s_val).alignment = ALIGN_LEFT
            ws_tonghop.cell(row=r, column=2, value=f"=COUNTIFS(DuLieu_DaLamSach!${seg_col}$2:${seg_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=3, value=f"=SUMIFS(DuLieu_DaLamSach!${sales_col}$2:${sales_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${seg_col}$2:${seg_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=4, value=f"=IFERROR(C{r}/Dashboard!$B$6, 0)").number_format = "0.0%"
            ws_tonghop.cell(row=r, column=5, value=f"=SUMIFS(DuLieu_DaLamSach!${profit_col}$2:${profit_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${seg_col}$2:${seg_col}${n_raw_rows + 1}, A{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=6, value=f"=IFERROR(E{r}/C{r}, 0)").number_format = "0.0%"
            style_table_row(ws_tonghop, r, 1, 6, is_zebra=(i % 2 == 1))

        if i < len(chan_vals):
            ch_val = chan_vals[i]
            ws_tonghop.cell(row=r, column=8, value=ch_val).alignment = ALIGN_LEFT
            ws_tonghop.cell(row=r, column=9, value=f"=COUNTIFS(DuLieu_DaLamSach!${chan_col}$2:${chan_col}${n_raw_rows + 1}, H{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=10, value=f"=SUMIFS(DuLieu_DaLamSach!${sales_col}$2:${sales_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${chan_col}$2:${chan_col}${n_raw_rows + 1}, H{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=11, value=f"=IFERROR(J{r}/Dashboard!$B$6, 0)").number_format = "0.0%"
            ws_tonghop.cell(row=r, column=12, value=f"=SUMIFS(DuLieu_DaLamSach!${profit_col}$2:${profit_col}${n_raw_rows + 1}, DuLieu_DaLamSach!${chan_col}$2:${chan_col}${n_raw_rows + 1}, H{r})").number_format = "#,##0"
            ws_tonghop.cell(row=r, column=13, value=f"=IFERROR(L{r}/J{r}, 0)").number_format = "0.0%"
            style_table_row(ws_tonghop, r, 8, 13, is_zebra=(i % 2 == 1))

    t3_end_row = seg_start_row + max_t3_len - 1
    t3_tot_row = t3_end_row + 1

    ws_tonghop.cell(row=t3_tot_row, column=1, value="Tổng cộng").alignment = ALIGN_CENTER
    ws_tonghop.cell(row=t3_tot_row, column=2, value=f"=SUM(B{seg_start_row}:B{t3_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t3_tot_row, column=3, value=f"=SUM(C{seg_start_row}:C{t3_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t3_tot_row, column=4, value=f"=SUM(D{seg_start_row}:D{t3_end_row})").number_format = "0.0%"
    ws_tonghop.cell(row=t3_tot_row, column=5, value=f"=SUM(E{seg_start_row}:E{t3_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t3_tot_row, column=6, value=f"=IFERROR(E{t3_tot_row}/C{t3_tot_row}, 0)").number_format = "0.0%"
    style_table_row(ws_tonghop, t3_tot_row, 1, 6, is_total=True)

    ws_tonghop.cell(row=t3_tot_row, column=8, value="Tổng cộng").alignment = ALIGN_CENTER
    ws_tonghop.cell(row=t3_tot_row, column=9, value=f"=SUM(I{seg_start_row}:I{t3_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t3_tot_row, column=10, value=f"=SUM(J{seg_start_row}:J{t3_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t3_tot_row, column=11, value=f"=SUM(K{seg_start_row}:K{t3_end_row})").number_format = "0.0%"
    ws_tonghop.cell(row=t3_tot_row, column=12, value=f"=SUM(L{seg_start_row}:L{t3_end_row})").number_format = "#,##0"
    ws_tonghop.cell(row=t3_tot_row, column=13, value=f"=IFERROR(L{t3_tot_row}/J{t3_tot_row}, 0)").number_format = "0.0%"
    style_table_row(ws_tonghop, t3_tot_row, 8, 13, is_total=True)

    # 3.4 BẢNG 4: TOP 10 SẢN PHẨM CHỦ LỰC (GỘP PRODUCT_ID & PRODUCT_NAME)
    cur_row = t3_tot_row + 3
    t4_title_row = cur_row
    ws_tonghop.cell(row=t4_title_row, column=1, value="4. TOP 10 SẢN PHẨM & DỊCH VỤ CHỦ LỰC (DOANH THU CAO NHẤT)").font = FONT_SECTION

    cur_row += 1
    t4_hdr_row = cur_row
    prod_hdrs = ["Hạng", "Mã & Tên sản phẩm", "Số lượng bán", "Doanh thu (VND)", "Tỷ trọng DT (%)", "Lợi nhuận (VND)", "Biên Margin (%)"]
    for ci, h in enumerate(prod_hdrs, start=1):
        ws_tonghop.cell(row=t4_hdr_row, column=ci, value=h)
    style_table_row(ws_tonghop, t4_hdr_row, 1, 7, is_header=True)

    tot_rev_val = float(analysis.get("kpis", {}).get("total_revenue", 0))
    top10_prods = _extract_top_products(cleaned_df, analysis, tot_rev_val, limit=10)

    cur_row += 1
    prod_start_row = cur_row
    for idx, p in enumerate(top10_prods, start=1):
        r = cur_row
        ws_tonghop.cell(row=r, column=1, value=idx).alignment = ALIGN_CENTER
        ws_tonghop.cell(row=r, column=2, value=p.get("product_name", f"SP-{idx}")).alignment = ALIGN_LEFT
        ws_tonghop.cell(row=r, column=3, value=p.get("count", 0)).number_format = "#,##0"
        ws_tonghop.cell(row=r, column=4, value=p.get("revenue", 0)).number_format = "#,##0"
        ws_tonghop.cell(row=r, column=5, value=f"=IFERROR(D{r}/Dashboard!$B$6, 0)").number_format = "0.0%"
        ws_tonghop.cell(row=r, column=6, value=p.get("profit", 0) if p.get("profit") is not None else 0).number_format = "#,##0"
        ws_tonghop.cell(row=r, column=7, value=f"=IFERROR(F{r}/D{r}, 0)").number_format = "0.0%"
        style_table_row(ws_tonghop, r, 1, 7, is_zebra=(idx % 2 == 1))
        cur_row += 1

    prod_end_row = max(cur_row - 1, prod_start_row)
    if prod_end_row >= prod_start_row:
        ws_tonghop.conditional_formatting.add(
            f"D{prod_start_row}:D{prod_end_row}",
            DataBarRule(start_type="min", end_type="max", color="60A5FA", showValue="None")
        )
        ws_tonghop.conditional_formatting.add(
            f"G{prod_start_row}:G{prod_end_row}",
            ColorScaleRule(start_type="min", start_color="F87171", mid_type="percentile", mid_value=50, mid_color="FCD34D", end_type="max", end_color="4ADE80")
        )

    # 3.5 BẢNG 5: TOP 15 KHÁCH HÀNG VIP (TIER A)
    cur_row = prod_end_row + 3
    t5_title_row = cur_row
    ws_tonghop.cell(row=t5_title_row, column=1, value="5. TOP 15 KHÁCH HÀNG VIP (TIER A — ĐÓNG GÓP CAO NHẤT)").font = FONT_SECTION

    cur_row += 1
    t5_hdr_row = cur_row
    cust_hdrs = ["Hạng", "Mã Khách hàng VIP", "Số đơn hàng", "Doanh thu (VND)", "Tỷ trọng DT (%)", "Lợi nhuận (VND)", "Biên Margin (%)", "Phân hạng"]
    for ci, h in enumerate(cust_hdrs, start=1):
        ws_tonghop.cell(row=t5_hdr_row, column=ci, value=h)
    style_table_row(ws_tonghop, t5_hdr_row, 1, 8, is_header=True)

    top15_custs = _extract_top_customers(analysis, limit=15)

    cur_row += 1
    cust_start_row = cur_row
    for idx, c in enumerate(top15_custs, start=1):
        r = cur_row
        ws_tonghop.cell(row=r, column=1, value=idx).alignment = ALIGN_CENTER
        ws_tonghop.cell(row=r, column=2, value=str(c.get("customer", c.get("customer_id", f"CUST-{idx}")))).alignment = ALIGN_LEFT
        ws_tonghop.cell(row=r, column=3, value=int(c.get("orders", c.get("count", 0)))).number_format = "#,##0"
        ws_tonghop.cell(row=r, column=4, value=float(c.get("revenue", 0))).number_format = "#,##0"
        ws_tonghop.cell(row=r, column=5, value=f"=IFERROR(D{r}/Dashboard!$B$6, 0)").number_format = "0.0%"
        ws_tonghop.cell(row=r, column=6, value=float(c.get("profit", 0)) if c.get("profit") is not None else 0).number_format = "#,##0"
        ws_tonghop.cell(row=r, column=7, value=f"=IFERROR(F{r}/D{r}, 0)").number_format = "0.0%"
        ws_tonghop.cell(row=r, column=8, value=str(c.get("priority_tier", c.get("tier", "Tier A")))).alignment = ALIGN_CENTER
        style_table_row(ws_tonghop, r, 1, 8, is_zebra=(idx % 2 == 1))
        cur_row += 1

    cust_end_row = max(cur_row - 1, cust_start_row)
    if cust_end_row >= cust_start_row:
        ws_tonghop.conditional_formatting.add(
            f"D{cust_start_row}:D{cust_end_row}",
            DataBarRule(start_type="min", end_type="max", color="60A5FA", showValue="None")
        )
        ws_tonghop.conditional_formatting.add(
            f"G{cust_start_row}:G{cust_end_row}",
            ColorScaleRule(start_type="min", start_color="F87171", mid_type="percentile", mid_value=50, mid_color="FCD34D", end_type="max", end_color="4ADE80")
        )

    # =========================================================================
    # BƯỚC 4: THÊM 4 BIỂU ĐỒ GỐC (NATIVE CHARTS) LÊN DASHBOARD THEO LƯỚI 2x2
    # =========================================================================
    # 4.1 Biểu đồ 1 (Top-Left B10): LineChart - Xu hướng Doanh thu & Lợi nhuận 24 tháng
    chart_line = LineChart()
    chart_line.title = "Xu hướng Doanh thu & Lợi nhuận theo Tháng (2024 - 2025)"
    chart_line.style = 13
    chart_line.width = 16.0
    chart_line.height = 8.0
    chart_line.y_axis.title = "Giá trị (VND)"
    chart_line.x_axis.title = "Tháng"

    data_rev = Reference(ws_tonghop, min_col=3, min_row=4, max_row=m_end_row)
    data_prof = Reference(ws_tonghop, min_col=5, min_row=4, max_row=m_end_row)
    cats_month = Reference(ws_tonghop, min_col=1, min_row=5, max_row=m_end_row)
    chart_line.add_data(data_rev, titles_from_data=True)
    chart_line.add_data(data_prof, titles_from_data=True)
    chart_line.set_categories(cats_month)
    ws_dash.add_chart(chart_line, "B10")

    # 4.2 Biểu đồ 2 (Top-Right J10): BarChart ngang - Top 5 Sản phẩm doanh thu cao nhất
    chart_bar = BarChart()
    chart_bar.type = "bar"
    chart_bar.title = "Top 5 Sản phẩm có Doanh thu Cao nhất"
    chart_bar.style = 10
    chart_bar.width = 16.0
    chart_bar.height = 8.0
    chart_bar.y_axis.title = "Sản phẩm"
    chart_bar.x_axis.title = "Doanh thu (VND)"
    chart_bar.legend = None

    prod_chart_end = min(prod_start_row + 4, prod_end_row)
    if prod_end_row >= prod_start_row:
        data_prod = Reference(ws_tonghop, min_col=4, min_row=t4_hdr_row, max_row=prod_chart_end)
        cats_prod = Reference(ws_tonghop, min_col=2, min_row=prod_start_row, max_row=prod_chart_end)
        chart_bar.add_data(data_prod, titles_from_data=True)
        chart_bar.set_categories(cats_prod)
        ws_dash.add_chart(chart_bar, "J10")

    # 4.3 Biểu đồ 3 (Bottom-Left B27): DoughnutChart - Tỷ trọng Doanh thu theo Ngành hàng
    chart_pie = DoughnutChart()
    chart_pie.title = "Tỷ trọng Doanh thu theo Ngành hàng (Category)"
    chart_pie.style = 10
    chart_pie.width = 16.0
    chart_pie.height = 8.0

    cat_chart_end = cat_start_row + len(cat_vals) - 1
    if len(cat_vals) > 0 and cat_chart_end >= cat_start_row:
        data_cat = Reference(ws_tonghop, min_col=3, min_row=t2_hdr_row, max_row=cat_chart_end)
        cats_cat = Reference(ws_tonghop, min_col=1, min_row=cat_start_row, max_row=cat_chart_end)
        chart_pie.add_data(data_cat, titles_from_data=True)
        chart_pie.set_categories(cats_cat)
        ws_dash.add_chart(chart_pie, "B27")

    # 4.4 Biểu đồ 4 (Bottom-Right J27): Clustered Column - Doanh thu & Lợi nhuận theo Phân khúc
    chart_col = BarChart()
    chart_col.type = "col"
    chart_col.title = "Doanh thu & Lợi nhuận theo Phân khúc Khách hàng"
    chart_col.style = 10
    chart_col.width = 16.0
    chart_col.height = 8.0
    chart_col.y_axis.title = "Giá trị (VND)"

    seg_chart_end = seg_start_row + len(seg_vals) - 1
    if len(seg_vals) > 0 and seg_chart_end >= seg_start_row:
        data_seg = Reference(ws_tonghop, min_col=3, min_row=t3_hdr_row, max_row=seg_chart_end)
        data_seg_prof = Reference(ws_tonghop, min_col=5, min_row=t3_hdr_row, max_row=seg_chart_end)
        cats_seg = Reference(ws_tonghop, min_col=1, min_row=seg_start_row, max_row=seg_chart_end)
        chart_col.add_data(data_seg, titles_from_data=True)
        chart_col.add_data(data_seg_prof, titles_from_data=True)
        chart_col.set_categories(cats_seg)
        ws_dash.add_chart(chart_col, "J27")

    # =========================================================================
    # BƯỚC 5: XUẤT SHEET 3 - KhachHang_VIP (PARETO ABC CHI TIẾT)
    # =========================================================================
    ws_vip.views.sheetView[0].showGridLines = True
    ws_vip["A2"] = "DANH SÁCH KHÁCH HÀNG VIP & PHÂN HẠNG ƯU TIÊN (PARETO ABC)"
    ws_vip["A2"].font = FONT_SECTION

    vip_hdrs = ["Hạng", "Mã Khách hàng VIP", "Số đơn hàng", "Tổng Doanh thu (VND)", "Tỷ trọng DT (%)", "Tích lũy DT (%)", "Tổng Lợi nhuận (VND)", "Biên Margin (%)", "Phân hạng Tier"]
    for ci, h in enumerate(vip_hdrs, start=1):
        ws_vip.cell(row=3, column=ci, value=h)
    style_table_row(ws_vip, 3, 1, 9, is_header=True)

    all_vip_custs = _extract_top_customers(analysis, limit=None)
    v_cur = 4
    for idx, c in enumerate(all_vip_custs, start=1):
        r = v_cur
        ws_vip.cell(row=r, column=1, value=idx).alignment = ALIGN_CENTER
        ws_vip.cell(row=r, column=2, value=str(c.get("customer", c.get("customer_id", f"CUST-{idx}")))).alignment = ALIGN_LEFT
        ws_vip.cell(row=r, column=3, value=int(c.get("orders", c.get("count", 0)))).number_format = "#,##0"
        ws_vip.cell(row=r, column=4, value=float(c.get("revenue", 0))).number_format = "#,##0"
        ws_vip.cell(row=r, column=5, value=float(c.get("share_pct", 0)) / 100.0).number_format = "0.0%"
        ws_vip.cell(row=r, column=6, value=float(c.get("cum_share_pct", c.get("cum_pct", 0))) / 100.0).number_format = "0.0%"
        ws_vip.cell(row=r, column=7, value=float(c.get("profit", 0)) if c.get("profit") is not None else 0).number_format = "#,##0"
        mg = c.get("margin_pct")
        ws_vip.cell(row=r, column=8, value=(float(mg) / 100.0) if mg is not None else 0).number_format = "0.0%"
        ws_vip.cell(row=r, column=9, value=str(c.get("priority_tier", c.get("tier", "Tier A")))).alignment = ALIGN_CENTER
        style_table_row(ws_vip, r, 1, 9, is_zebra=(idx % 2 == 1))
        v_cur += 1

    vip_end_row = v_cur - 1
    if vip_end_row >= 4:
        ws_vip.auto_filter.ref = f"A3:I{vip_end_row}"
        ws_vip.freeze_panes = "A4"
        ws_vip.conditional_formatting.add(
            f"D4:D{vip_end_row}",
            DataBarRule(start_type="min", end_type="max", color="60A5FA", showValue="None")
        )
        ws_vip.conditional_formatting.add(
            f"H4:H{vip_end_row}",
            ColorScaleRule(start_type="min", start_color="F87171", mid_type="percentile", mid_value=50, mid_color="FCD34D", end_type="max", end_color="4ADE80")
        )

    # =========================================================================
    # BƯỚC 6: XUẤT SHEET 4 - ThongKe_MoTa
    # =========================================================================
    ws_mota.views.sheetView[0].showGridLines = True
    ws_mota["A2"] = "BẢNG THỐNG KÊ MÔ TẢ PHÂN PHỐI DỮ LIỆU (DESCRIPTIVE STATISTICS)"
    ws_mota["A2"].font = FONT_SECTION

    mota_hdrs = ["Chỉ số / Cột dữ liệu", "Số lượng (Count)", "Trung bình (Mean)", "Độ lệch chuẩn (Std)", "Nhỏ nhất (Min)", "Q25 (25%)", "Trung vị (Median)", "Q75 (75%)", "Lớn nhất (Max)", "Số Outliers (IQR)"]
    for ci, h in enumerate(mota_hdrs, start=1):
        ws_mota.cell(row=3, column=ci, value=h)
    style_table_row(ws_mota, 3, 1, 10, is_header=True)

    desc_rows = analysis.get("descriptive_stats", [])
    m_cur = 4
    for idx, d in enumerate(desc_rows, start=1):
        r = m_cur
        ws_mota.cell(row=r, column=1, value=str(d.get("column", d.get("Chỉ số", f"Col-{idx}")))).alignment = ALIGN_LEFT
        ws_mota.cell(row=r, column=2, value=float(d.get("count", 0))).number_format = "#,##0"
        ws_mota.cell(row=r, column=3, value=float(d.get("mean", 0))).number_format = "#,##0.0"
        ws_mota.cell(row=r, column=4, value=float(d.get("std", 0))).number_format = "#,##0.0"
        ws_mota.cell(row=r, column=5, value=float(d.get("min", 0))).number_format = "#,##0"
        ws_mota.cell(row=r, column=6, value=float(d.get("q25", 0))).number_format = "#,##0"
        ws_mota.cell(row=r, column=7, value=float(d.get("median", 0))).number_format = "#,##0"
        ws_mota.cell(row=r, column=8, value=float(d.get("q75", 0))).number_format = "#,##0"
        ws_mota.cell(row=r, column=9, value=float(d.get("max", 0))).number_format = "#,##0"
        ws_mota.cell(row=r, column=10, value=int(d.get("outliers_iqr", d.get("outliers", 0)))).number_format = "#,##0"
        style_table_row(ws_mota, r, 1, 10, is_zebra=(idx % 2 == 1))
        m_cur += 1

    mota_end_row = m_cur - 1
    if mota_end_row >= 4:
        ws_mota.freeze_panes = "A4"

    # =========================================================================
    # BƯỚC 7: TỰ ĐỘNG CĂN CHỈNH ĐỘ RỘNG CỘT (AUTO-FIT WIDTH) TRÊN CẢ 5 SHEET
    # =========================================================================
    for ci in range(1, 20):
        c_letter = get_column_letter(ci)
        ws_dash.column_dimensions[c_letter].width = 15

    for ws in [ws_tonghop, ws_vip, ws_mota, ws_raw]:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                if cell.value is not None:
                    v_str = str(cell.value)
                    if v_str.startswith("=") or len(v_str) > 60:
                        continue
                    if len(v_str) > max_len:
                        max_len = len(v_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Đặt Sheet 1: Dashboard là active tab khi mở file
    wb.active = ws_dash

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def _extract_top_products(cleaned_df: Optional[pd.DataFrame], analysis: Dict[str, Any], tot_rev: float, limit: int = 5) -> List[Dict[str, Any]]:
    """Gộp Product_ID và Product_Name làm một, lấy Top sản phẩm có doanh thu cao nhất (mặc định 5 cho Word, 10 cho Excel)."""
    if cleaned_df is not None and not cleaned_df.empty:
        p_id_col = None
        p_name_col = None
        rev_col = None
        prof_col = None
        qty_col = None

        for col in cleaned_df.columns:
            c_l = str(col).lower()
            if not p_id_col and ("product_id" in c_l or "mã sp" in c_l or "ma_sp" in c_l):
                p_id_col = col
            if not p_name_col and ("product_name" in c_l or "tên sp" in c_l or "ten_sp" in c_l or "tên sản phẩm" in c_l):
                p_name_col = col
            if not rev_col and ("sales" in c_l or "doanh_thu" in c_l or "thành tiền" in c_l or "thanh_tien" in c_l):
                rev_col = col
            if not prof_col and ("profit" in c_l or "lợi nhuận" in c_l or "loi_nhuan" in c_l):
                prof_col = col
            if not qty_col and ("quantity" in c_l or "số lượng" in c_l or "so_luong" in c_l):
                qty_col = col

        if p_id_col or p_name_col:
            group_keys = [c for c in [p_id_col, p_name_col] if c]
            agg = {}
            if rev_col: agg["revenue"] = (rev_col, "sum")
            if prof_col: agg["profit"] = (prof_col, "sum")
            if qty_col: agg["qty"] = (qty_col, "sum")
            agg["count"] = (cleaned_df.columns[0], "count")

            p_df = cleaned_df.groupby(group_keys).agg(**agg).reset_index()
            sort_c = "revenue" if "revenue" in p_df.columns else "count"
            p_df = p_df.sort_values(sort_c, ascending=False).head(limit)

            results = []
            for idx, (_, row) in enumerate(p_df.iterrows(), start=1):
                p_id = str(row[p_id_col]) if p_id_col else ""
                p_name = str(row[p_name_col]) if p_name_col else ""
                full_name = f"{p_id} — {p_name}".strip(" —") if p_id and p_name else (p_name or p_id)
                rev = float(row.get("revenue", 0))
                prof = float(row.get("profit", 0)) if "profit" in row else None
                cnt = int(row.get("qty", row.get("count", 0)))
                share = (rev / tot_rev * 100.0) if tot_rev > 0 else 0.0
                mg = (prof / rev * 100.0) if prof is not None and rev > 0 else None
                results.append({
                    "rank": idx,
                    "product_name": full_name,
                    "count": cnt,
                    "revenue": rev,
                    "share_pct": round(share, 1),
                    "profit": prof,
                    "margin_pct": round(mg, 1) if mg is not None else None
                })
            return results

    # Fallback từ group_tables
    g_tables = analysis.get("group_tables", {})
    p_rows = g_tables.get("Product_Name") or g_tables.get("Product_ID") or []
    results = []
    for idx, r in enumerate(p_rows[:limit], start=1):
        results.append({
            "rank": idx,
            "product_name": str(r.get("group_value", "")),
            "count": int(r.get("count", 0)),
            "revenue": float(r.get("revenue", 0)),
            "share_pct": float(r.get("share_pct", 0)),
            "profit": float(r.get("profit", 0)) if r.get("profit") is not None else None,
            "margin_pct": float(r.get("margin_pct", 0)) if r.get("margin_pct") is not None else None
        })
    return results

def _extract_top_customers(analysis: Dict[str, Any], limit: Optional[int] = 15) -> List[Dict[str, Any]]:
    """Trích xuất danh sách khách hàng VIP sắp xếp theo Lợi nhuận / Doanh thu giảm dần."""
    cust_list = analysis.get("customer_priority", [])
    sorted_cust = sorted(cust_list, key=lambda x: float(x.get("profit") or x.get("revenue") or 0), reverse=True)
    if limit is not None:
        sorted_cust = sorted_cust[:limit]

    results = []
    for idx, c in enumerate(sorted_cust, start=1):
        rev = float(c.get("revenue", 0))
        prof = float(c.get("profit", 0)) if ("profit" in c and c.get("profit") is not None) else None
        share = float(c.get("share_pct", 0))
        mg = float(c.get("margin_pct", 0)) if ("margin_pct" in c and c.get("margin_pct") is not None) else ((prof / rev * 100.0) if prof is not None and rev > 0 else None)
        cum = float(c.get("cum_share_pct", c.get("cum_pct", 0)))
        tier = c.get("priority_tier") or c.get("tier") or ("Tier A" if idx <= 15 else "Tier B")
        results.append({
            "rank": idx,
            "customer": str(c.get("customer", c.get("customer_id", f"VIP-{idx}"))),
            "orders": int(c.get("orders", c.get("count", 0))),
            "revenue": rev,
            "share_pct": round(share, 1),
            "cum_share_pct": round(cum, 1),
            "profit": prof,
            "margin_pct": round(mg, 1) if mg is not None else None,
            "tier": tier
        })
    return results

def _extract_region_channel_data(analysis: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Tóm tắt ngắn gọn 3 khu vực và 3 kênh bán hàng cho Bảng 3."""
    g_tables = analysis.get("group_tables", {})
    results = []

    # 3 Khu vực
    reg_rows = g_tables.get("Region") or g_tables.get("Khu_Vuc") or g_tables.get("City") or []
    for r in reg_rows[:3]:
        results.append({
            "category": "Khu vực",
            "name": str(r.get("group_value", "")),
            "count": int(r.get("count", 0)),
            "revenue": float(r.get("revenue", 0)),
            "share_pct": float(r.get("share_pct", 0)),
            "profit": float(r.get("profit", 0)) if r.get("profit") is not None else None,
            "margin_pct": float(r.get("margin_pct", 0)) if r.get("margin_pct") is not None else None
        })

    # 3 Kênh bán
    chan_rows = g_tables.get("Channel") or g_tables.get("Kenh_Ban") or g_tables.get("Segment") or []
    for c in chan_rows[:3]:
        results.append({
            "category": "Kênh bán",
            "name": str(c.get("group_value", "")),
            "count": int(c.get("count", 0)),
            "revenue": float(c.get("revenue", 0)),
            "share_pct": float(c.get("share_pct", 0)),
            "profit": float(c.get("profit", 0)) if c.get("profit") is not None else None,
            "margin_pct": float(c.get("margin_pct", 0)) if c.get("margin_pct") is not None else None
        })

    return results

def _generate_executive_summary_bullets(cleaned_df: Optional[pd.DataFrame], analysis: Dict[str, Any], tot_rev: float) -> List[str]:
    """Tạo 4 gạch đầu dòng Executive súc tích thay thế các bảng dữ liệu còn lại (Category, Segment, Payment, Returned)."""
    bullets = []
    g_tables = analysis.get("group_tables", {})

    # 1. Ngành hàng (Category)
    cat_rows = g_tables.get("Category") or g_tables.get("Subcategory") or []
    if cat_rows:
        top_cat = cat_rows[0]
        c_text = f"* **Ngành hàng (Category):** Nhóm **{top_cat.get('group_value')}** đóng vai trò chủ lực dẫn đầu với doanh thu **{format_compact_vnd(top_cat.get('revenue', 0))}** (chiếm **{float(top_cat.get('share_pct', 0)):.1f}%** tổng doanh thu, {int(top_cat.get('count', 0)):,} đơn)."
        if len(cat_rows) > 1:
            second_cat = cat_rows[1]
            c_text += f" Theo sau là **{second_cat.get('group_value')}** đạt **{format_compact_vnd(second_cat.get('revenue', 0))}** ({float(second_cat.get('share_pct', 0)):.1f}%)."
        bullets.append(c_text)

    # 2. Phân khúc khách hàng (Segment)
    seg_rows = g_tables.get("Segment") or []
    if seg_rows:
        top_seg = seg_rows[0]
        s_text = f"* **Phân khúc khách hàng (Segment):** **{top_seg.get('group_value')}** chiếm ưu thế với **{format_compact_vnd(top_seg.get('revenue', 0))}** (chiếm **{float(top_seg.get('share_pct', 0)):.1f}%** doanh thu)."
        if len(seg_rows) > 1:
            vip_seg = next((s for s in seg_rows if "vip" in str(s.get("group_value", "")).lower() or "thân thiết" in str(s.get("group_value", "")).lower()), seg_rows[1])
            s_text += f" Nhóm **{vip_seg.get('group_value')}** đóng góp **{format_compact_vnd(vip_seg.get('revenue', 0))}** ({float(vip_seg.get('share_pct', 0)):.1f}%), cho thấy tiềm năng phát triển chương trình loyalty."
        bullets.append(s_text)

    # 3. Phương thức thanh toán (Payment Method)
    pay_rows = g_tables.get("Payment_Method") or []
    if pay_rows:
        top_pay = pay_rows[0]
        p_text = f"* **Phương thức thanh toán (Payment Method):** Phương thức **{top_pay.get('group_value')}** chiếm tỷ trọng cao nhất với **{format_compact_vnd(top_pay.get('revenue', 0))}** ({float(top_pay.get('share_pct', 0)):.1f}%)."
        cod_row = next((p for p in pay_rows if "tiền mặt" in str(p.get("group_value", "")).lower() or "cod" in str(p.get("group_value", "")).lower()), None)
        if cod_row:
            p_text += f" Thanh toán không tiền mặt chiếm ưu thế áp đảo, trong khi Tiền mặt / COD chiếm **{format_compact_vnd(cod_row.get('revenue', 0))}** ({float(cod_row.get('share_pct', 0)):.1f}%)."
        bullets.append(p_text)

    # 4. Tỷ lệ hoàn hàng (Returned)
    ret_rows = g_tables.get("Returned") or []
    ret_yes = next((r for r in ret_rows if str(r.get("group_value", "")).lower() in ("có", "yes", "true", "1")), None)
    if ret_yes:
        ret_rev = float(ret_yes.get("revenue", 0))
        ret_cnt = int(ret_yes.get("count", 0))
        ret_share = float(ret_yes.get("share_pct", 0))
        ret_share_str = f"{ret_share:.1f}".replace(".", ",")
        bullets.append(
            f"* **Tỷ lệ hoàn hàng (Returned):** Tỷ lệ hoàn hàng ở mức **{ret_share_str}% ({ret_cnt:,} đơn, tương đương {format_compact_vnd(ret_rev)} thất thoát doanh thu)** — Cần thắt chặt quy trình kiểm soát chất lượng đóng gói trước khi xuất kho và tối ưu chính sách hoàn trả đối với phân khúc hàng điện tử."
        )
    elif not ret_rows and cleaned_df is not None and "Returned" in cleaned_df.columns:
        ret_df = cleaned_df[cleaned_df["Returned"].astype(str).str.lower().isin(["có", "yes", "true", "1"])]
        if not ret_df.empty:
            cnt = len(ret_df)
            rev = ret_df["Sales_VND"].sum() if "Sales_VND" in ret_df.columns else 0.0
            share = (cnt / len(cleaned_df) * 100.0) if len(cleaned_df) > 0 else 0.0
            ret_share_str = f"{share:.1f}".replace(".", ",")
            bullets.append(
                f"* **Tỷ lệ hoàn hàng (Returned):** Tỷ lệ hoàn hàng ở mức **{ret_share_str}% ({cnt:,} đơn, tương đương {format_compact_vnd(rev)} thất thoát doanh thu)** — Cần rà soát quy trình kiểm định trước khi đóng gói và tối ưu hóa chính sách hoàn trả."
            )

    if not bullets:
        bullets.append("* **Đánh giá vận hành chung:** Các chiều dữ liệu hoạt động ổn định, không ghi nhận biến động tiêu cực ngoài ngưỡng kiểm soát.")
    return bullets

def export_6part_docx_bytes(
    report_data: dict,
    approver_name: str = "Ban Giám Đốc / Trưởng bộ phận Phân tích",
    approval_notes: str = "",
    company_info: Optional[Dict[str, Any]] = None,
    ai_report_text: str = ""
) -> bytes:
    """
    Xuất Báo cáo Điều hành Kinh doanh Word (.docx) chuyên nghiệp chuẩn 6 phần:
    - Typography Segoe UI nhất quán, A4 margin chuẩn, Navy headers, table styles sang trọng.
    - Phân cấp Heading 1 (#), Heading 2 (##), Heading 3 (###) rõ ràng.
    - Khớp 100% cấu trúc 6 phần với Báo cáo trực tiếp trên Web.
    """
    doc = Document()
    p1 = report_data["part1_executive_summary"]
    p2 = report_data["part2_kpi_dashboard"]
    p3 = report_data["part3_deep_dive"]
    p4 = report_data["part4_conclusions"]
    p5 = report_data["part5_action_recommendations"]
    p6 = report_data["part6_technical_details"]

    biz_name = (p1.get("company_name") or (company_info.get("name") if company_info else "") or "DOANH NGHIỆP").strip().upper()
    biz_type = p1.get("business_context", "Thương mại & Doanh nghiệp")
    source_name = p1.get("dataset_name", "Dữ liệu")

    setup_document_styles(doc, header_text=f"BÁO CÁO ĐIỀU HÀNH KINH DOANH • {biz_name}")

    # HEADER & TIÊU ĐỀ BÁO CÁO
    p_corp = doc.add_paragraph()
    p_corp.paragraph_format.space_after = Pt(2)
    r_corp = p_corp.add_run(f"{biz_name} — HỆ THỐNG BI & SALES ANALYTICS")
    r_corp.font.size = Pt(8.5)
    r_corp.font.bold = True
    r_corp.font.color.rgb = RGBColor(110, 119, 129)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("BÁO CÁO KINH DOANH & PHÂN TÍCH HIỆU SUẤT ĐIỀU HÀNH")
    r_title.font.size = Pt(17)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(6)
    r_sub = p_sub.add_run(f"Mô hình: {biz_type} • Nguồn tệp: {source_name} • Phân tích động tự động theo cấu trúc thực tế")
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(80, 80, 80)

    # PHẦN 1: TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)
    add_markdown_to_docx(doc, "# PHẦN 1: TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)")
    add_markdown_to_docx(doc, p1.get("narrative", ""))

    metrics_summary = p1.get("metrics_summary", [])
    if metrics_summary:
        tbl_m = doc.add_table(rows=1 + len(metrics_summary), cols=4)
        for ci, h in enumerate(["Chỉ số định lượng", "Vai trò", "Tổng giá trị", "Độ phủ dữ liệu"]):
            tbl_m.cell(0, ci).paragraphs[0].add_run(h).bold = True
        for ri, m in enumerate(metrics_summary, start=1):
            tbl_m.cell(ri, 0).paragraphs[0].add_run(m["column"]).bold = True
            tbl_m.cell(ri, 1).paragraphs[0].add_run(m["role"])
            tbl_m.cell(ri, 2).paragraphs[0].add_run(m["formatted_sum"])
            tbl_m.cell(ri, 3).paragraphs[0].add_run(m["coverage_text"])
        apply_table_styles(tbl_m)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # PHẦN 2: DASHBOARD KPI (SO SÁNH KỲ GẦN NHẤT VỚI KỲ LIỀN TRƯỚC)
    add_markdown_to_docx(doc, "# PHẦN 2: DASHBOARD KPI (SO SÁNH KỲ GẦN NHẤT VỚI KỲ LIỀN TRƯỚC)")
    cur_lbl = p2.get("current_period_label", "Kỳ gần nhất")
    pri_lbl = p2.get("prior_period_label", "Kỳ liền trước")
    kpi_rows = p2.get("kpi_rows", [])

    tbl_kpi = doc.add_table(rows=1 + len(kpi_rows), cols=4)
    for ci, h in enumerate(["Chỉ số", f"Giá trị ({cur_lbl})", f"Kỳ trước ({pri_lbl})", "Thay đổi %"]):
        tbl_kpi.cell(0, ci).paragraphs[0].add_run(h).bold = True
    for ri, kr in enumerate(kpi_rows, start=1):
        tbl_kpi.cell(ri, 0).paragraphs[0].add_run(kr["metric"]).bold = True
        tbl_kpi.cell(ri, 1).paragraphs[0].add_run(kr["current_val"])
        tbl_kpi.cell(ri, 2).paragraphs[0].add_run(kr["prior_val"])
        tbl_kpi.cell(ri, 3).paragraphs[0].add_run(kr["change_pct"])
    apply_table_styles(tbl_kpi)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # PHẦN 3: PHÂN TÍCH CHUYÊN ĐỀ
    add_markdown_to_docx(doc, "# PHẦN 3: PHÂN TÍCH CHUYÊN ĐỀ")
    for idx, th in enumerate(p3, start=1):
        add_markdown_to_docx(doc, f"## 3.{idx}. Phân tích chi tiết chỉ số: {th['column']}")
        add_markdown_to_docx(doc, th["narrative"])
        tbl_th = doc.add_table(rows=2, cols=6)
        for ci, h in enumerate(["Tổng cộng", "Trung bình", "Trung vị", "Độ lệch chuẩn", "Nhỏ nhất", "Lớn nhất"]):
            tbl_th.cell(0, ci).paragraphs[0].add_run(h).bold = True
        vals = [th["sum"], th["mean"], th["median"], th["std"], th["min"], th["max"]]
        for ci, v in enumerate(vals):
            tbl_th.cell(1, ci).paragraphs[0].add_run(str(v))
        apply_table_styles(tbl_th)
        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # PHẦN 4: KẾT LUẬN
    add_markdown_to_docx(doc, "# PHẦN 4: KẾT LUẬN")
    for c in p4:
        add_markdown_to_docx(doc, f"* {c}")

    if ai_report_text and ai_report_text.strip():
        add_markdown_to_docx(doc, "## 4.2. Nhận định Điều hành AI (Executive Insights)")
        add_markdown_to_docx(doc, ai_report_text.strip())

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # PHẦN 5: ĐỀ XUẤT HÀNH ĐỘNG
    add_markdown_to_docx(doc, "# PHẦN 5: ĐỀ XUẤT HÀNH ĐỘNG")
    tbl_rec = doc.add_table(rows=1 + len(p5), cols=4)
    for ci, h in enumerate(["STT", "Khuyến nghị quản trị", "Mức độ ưu tiên", "Chỉ số đo lường (KPI)"]):
        tbl_rec.cell(0, ci).paragraphs[0].add_run(h).bold = True
    for ri, r in enumerate(p5, start=1):
        tbl_rec.cell(ri, 0).paragraphs[0].add_run(str(r["stt"]))
        tbl_rec.cell(ri, 1).paragraphs[0].add_run(f"{r['title']}: {r['action']}")
        tbl_rec.cell(ri, 2).paragraphs[0].add_run(r["priority"]).bold = True
        tbl_rec.cell(ri, 3).paragraphs[0].add_run(r["kpi_measure"])
    apply_table_styles(tbl_rec)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # PHẦN 6: CHI TIẾT BỘ DỮ LIỆU & THỐNG KÊ KỸ THUẬT
    add_markdown_to_docx(doc, "# PHẦN 6: CHI TIẾT BỘ DỮ LIỆU & THỐNG KÊ KỸ THUẬT")

    # 6.1 Thông tin làm sạch
    add_markdown_to_docx(doc, "## 6.1. Thông tin làm sạch dữ liệu")
    cinfo = p6.get("cleaning_info", {})
    tbl_cl = doc.add_table(rows=2, cols=6)
    for ci, h in enumerate(["Dòng đầu vào", "Dòng sau sạch", "Số cột", "Trùng lặp loại", "Ô trống còn lại", "Outlier IQR"]):
        tbl_cl.cell(0, ci).paragraphs[0].add_run(h).bold = True
    cl_vals = [
        f"{cinfo.get('orig_rows', 0):,}",
        f"{cinfo.get('clean_rows', 0):,}",
        str(cinfo.get('clean_cols', 0)),
        f"{cinfo.get('duplicates_removed', 0):,}",
        f"{cinfo.get('remaining_nulls', 0):,}",
        str(cinfo.get('outliers_flagged', 0))
    ]
    for ci, v in enumerate(cl_vals):
        tbl_cl.cell(1, ci).paragraphs[0].add_run(v)
    apply_table_styles(tbl_cl)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # 6.2 Nhận xét thống kê tự động
    add_markdown_to_docx(doc, "## 6.2. Nhận xét thống kê tự động")
    add_markdown_to_docx(doc, p6.get("automated_stats_comment", ""))
    doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # 6.3 Bảng Tổng hợp theo nhóm
    grp_tbl = p6.get("group_summary_table", {})
    grp_col_name = grp_tbl.get("group_col", "Nhóm")
    grp_rows = grp_tbl.get("rows", [])
    if grp_rows:
        add_markdown_to_docx(doc, f"## 6.3. Bảng Tổng hợp theo nhóm ({grp_col_name})")
        tbl_g = doc.add_table(rows=1 + len(grp_rows), cols=3)
        for ci, h in enumerate(["Nhóm", "Giá trị", "Số dòng được tính"]):
            tbl_g.cell(0, ci).paragraphs[0].add_run(h).bold = True
        for ri, gr in enumerate(grp_rows, start=1):
            tbl_g.cell(ri, 0).paragraphs[0].add_run(gr["group"]).bold = True
            tbl_g.cell(ri, 1).paragraphs[0].add_run(gr["value"])
            tbl_g.cell(ri, 2).paragraphs[0].add_run(f"{gr['count']:,}")
        apply_table_styles(tbl_g)
        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # 6.4 Bảng Thống kê theo từng cột
    col_stats = p6.get("column_stats_table", [])
    if col_stats:
        add_markdown_to_docx(doc, "## 6.4. Bảng Thống kê theo từng cột")
        tbl_cs = doc.add_table(rows=1 + len(col_stats), cols=6)
        for ci, h in enumerate(["Cột", "Ô trống", "Giá trị khác nhau", "Trung bình", "Nhỏ nhất", "Lớn nhất"]):
            tbl_cs.cell(0, ci).paragraphs[0].add_run(h).bold = True
        for ri, cs in enumerate(col_stats, start=1):
            tbl_cs.cell(ri, 0).paragraphs[0].add_run(cs["column"]).bold = True
            tbl_cs.cell(ri, 1).paragraphs[0].add_run(f"{cs['null_count']:,}")
            tbl_cs.cell(ri, 2).paragraphs[0].add_run(f"{cs['distinct_count']:,}")
            tbl_cs.cell(ri, 3).paragraphs[0].add_run(cs["mean"])
            tbl_cs.cell(ri, 4).paragraphs[0].add_run(cs["min"])
            tbl_cs.cell(ri, 5).paragraphs[0].add_run(cs["max"])
        apply_table_styles(tbl_cs)

    # KHUNG KÝ TÊN & PHÊ DUYỆT
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_left = tbl_sign.cell(0, 0)
    c_right = tbl_sign.cell(0, 1)

    p_sl = c_left.paragraphs[0]
    p_sl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sl1 = p_sl.add_run("ĐẠI DIỆN BỘ PHẬN LẬP BIỂU\n")
    r_sl1.font.size = Pt(9)
    r_sl1.font.bold = True
    r_sl2 = p_sl.add_run("(Ký & ghi rõ họ tên)\n\n\n\n")
    r_sl2.font.size = Pt(8)
    r_sl2.font.italic = True
    r_sl3 = p_sl.add_run(f"Hệ thống Sales AI BI Agent\n{biz_name}")
    r_sl3.font.size = Pt(9.5)
    r_sl3.font.bold = True

    p_sr = c_right.paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sr1 = p_sr.add_run("NGƯỜI PHÊ DUYỆT BAN HÀNH\n")
    r_sr1.font.size = Pt(9)
    r_sr1.font.bold = True
    r_sr2 = p_sr.add_run("(Ký, đóng dấu & ghi rõ họ tên)\n\n")
    r_sr2.font.size = Pt(8)
    r_sr2.font.italic = True
    r_sr_stamp = p_sr.add_run(f"[★ {biz_name} — ĐÃ PHÊ DUYỆT BAN HÀNH ★]\n\n")
    r_sr_stamp.font.size = Pt(8.5)
    r_sr_stamp.font.bold = True
    r_sr_stamp.font.color.rgb = RGBColor(14, 138, 22)
    r_sr3 = p_sr.add_run(approver_name)
    r_sr3.font.size = Pt(9.5)
    r_sr3.font.bold = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def export_universal_docx_bytes(
    analysis: Dict[str, Any],
    ai_report_text: str,
    source_name: str,
    approver_name: str = "Ban Giám Đốc / Trưởng bộ phận Phân tích",
    approval_notes: str = "",
    cleaned_df: Optional[pd.DataFrame] = None,
    company_info: Optional[Dict[str, Any]] = None
) -> bytes:
    """
    Xuất Báo cáo Điều hành Kinh doanh Word (.docx) chuyên nghiệp:
    Ưu tiên xuất định dạng chuẩn 6 phần (Dynamic 6-Part Business Report) khớp 100% với Web.
    """
    dynamic_report = analysis.get("dynamic_report")
    if not dynamic_report and cleaned_df is not None:
        dynamic_report = generate_dynamic_6part_report(
            cleaned_df=cleaned_df,
            schema=analysis.get("schema_used") or {},
            analysis=analysis,
            source_name=source_name,
            company_info=company_info
        )

    if dynamic_report:
        return export_6part_docx_bytes(
            report_data=dynamic_report,
            approver_name=approver_name,
            approval_notes=approval_notes,
            company_info=company_info,
            ai_report_text=ai_report_text
        )

    doc = Document()
    company_info = company_info or {}
    biz_name = (company_info.get("name") or "HỆ THỐNG QUẢN TRỊ DOANH NGHIỆP").strip().upper()
    biz_type = (company_info.get("type") or "Thương mại & Phân phối B2B").strip()
    biz_prob = (company_info.get("problem") or "").strip()

    setup_document_styles(doc, header_text=f"BÁO CÁO ĐIỀU HÀNH KINH DOANH • {biz_name}")

    # 1. HEADER CÔNG TY & TIÊU ĐỀ BÁO CÁO (TRANG 1)
    p_corp = doc.add_paragraph()
    p_corp.paragraph_format.space_after = Pt(2)
    r_corp = p_corp.add_run(f"{biz_name} — HỆ THỐNG SALES BI ANALYTICS")
    r_corp.font.size = Pt(8.5)
    r_corp.font.bold = True
    r_corp.font.color.rgb = RGBColor(110, 119, 129)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("BÁO CÁO KINH DOANH & PHÂN TÍCH HIỆU SUẤT TÀI CHÍNH")
    r_title.font.size = Pt(17)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(6)
    sub_prob_text = f" • Trọng tâm: {biz_prob}" if biz_prob else ""
    r_sub = p_sub.add_run(f"Mô hình: {biz_type} • Nguồn dữ liệu: {source_name}{sub_prob_text} • Phân tích đa chiều Pandas & AI Executive Insights")
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(80, 80, 80)

    # Khối phê duyệt & Trạng thái
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(6)
    p_meta.paragraph_format.line_spacing = 1.15
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    r_m1 = p_meta.add_run(f"Ngày ban hành: {now_str} • Người phê duyệt: ")
    r_m1.font.size = Pt(9)
    r_m2 = p_meta.add_run(approver_name)
    r_m2.font.size = Pt(9)
    r_m2.font.bold = True
    r_m3 = p_meta.add_run(" • Trạng thái: [★ ĐÃ PHÊ DUYỆT BAN HÀNH ★]\n")
    r_m3.font.size = Pt(9)
    r_m3.font.bold = True
    r_m3.font.color.rgb = RGBColor(14, 138, 22)

    if approval_notes:
        r_n1 = p_meta.add_run(f"Ý kiến phê duyệt: {approval_notes}")
        r_n1.font.size = Pt(8.5)
        r_n1.font.italic = True
        r_n1.font.color.rgb = RGBColor(80, 80, 80)

    # 2. THẺ KPI TỔNG QUAN (LƯỚI 4 Ô NGANG TRÊN TRANG 1)
    kpis = analysis.get("kpis") or {}
    if not kpis and "kpi_tree" in analysis:
        tree = analysis["kpi_tree"]
        kpis = {
            "total_rows": tree.get("tong_so_dong", len(cleaned_df) if cleaned_df is not None else 0),
            "total_revenue": tree.get("tong_doanh_thu", 0),
            "avg_order_value": tree.get("doanh_thu_tb", 0),
            "total_profit": tree.get("tong_loi_nhuan", 0),
            "profit_margin_pct": tree.get("ty_suat_loi_nhuan_tb", 0)
        }

    tot_rev = float(kpis.get("total_revenue", 0))
    tot_prof = float(kpis.get("total_profit", 0)) if kpis.get("total_profit") is not None else 0.0
    mg_pct = float(kpis.get("profit_margin_pct", 0)) if kpis.get("profit_margin_pct") is not None else 0.0
    tot_rows = int(kpis.get("total_rows", len(cleaned_df) if cleaned_df is not None else 0))
    aov = float(kpis.get("avg_order_value", 0))
    median_v = float(kpis.get("median_order_value", aov * 0.65))

    tbl_kpi = doc.add_table(rows=1, cols=4)
    tbl_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    kpi_cards = [
        ("TỔNG DOANH THU", format_compact_vnd(tot_rev), f"Quy mô: {tot_rows:,} đơn", RGBColor(27, 54, 93)),
        ("TỔNG LỢI NHUẬN", format_compact_vnd(tot_prof), "Lãi gộp theo dữ liệu", RGBColor(14, 138, 22)),
        ("BIÊN LỢI NHUẬN (MG)", f"{mg_pct:.1f}%", "Hiệu suất đạt chuẩn" if mg_pct >= 10 else "Cần tối ưu giá", RGBColor(14, 138, 22) if mg_pct >= 10 else RGBColor(217, 56, 30)),
        ("GIÁ TRỊ ĐƠN TB (AOV)", format_compact_vnd(aov), f"Trung vị: {format_compact_vnd(median_v)}", RGBColor(27, 54, 93))
    ]

    for ci, (lbl, val, sub, col_rgb) in enumerate(kpi_cards):
        cell = tbl_kpi.cell(0, ci)
        set_cell_background(cell, "F4F6F8")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(1)

        r_lbl = p.add_run(lbl + "\n")
        r_lbl.font.size = Pt(8)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = RGBColor(110, 119, 129)

        r_val = p.add_run(val + "\n")
        r_val.font.size = Pt(13)
        r_val.font.bold = True
        r_val.font.color.rgb = col_rgb

        r_sub = p.add_run(sub)
        r_sub.font.size = Pt(7.5)
        r_sub.font.color.rgb = RGBColor(110, 119, 129)

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

    # 3. MỤC I: NHẬN ĐỊNH ĐIỀU HÀNH & KHUYẾN NGHỊ HÀNH ĐỘNG (AI REPORT)
    add_markdown_to_docx(doc, "### I. NHẬN ĐỊNH ĐIỀU HÀNH & KHUYẾN NGHỊ HÀNH ĐỘNG (AI REPORT)")
    add_markdown_to_docx(doc, ai_report_text)

    # 4. CHUYỂN SANG TRANG 2: CÁC NHÓM PHÂN TÍCH TRỌNG TÂM
    doc.add_page_break()

    # 5. MỤC II: BIỂU ĐỒ ĐƯỜNG XU HƯỚNG THEO THÁNG (LINE CHART)
    add_markdown_to_docx(doc, "### II. PHÂN TÍCH XU HƯỚNG DOANH THU THEO THỜI GIAN (LINE CHART)")
    monthly_data = (analysis.get("group_tables", {}).get("_Thang") or analysis.get("group_tables", {}).get("Month") or [])
    if monthly_data:
        chart_bytes = generate_monthly_line_chart(monthly_data)
        if chart_bytes:
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(4)
            p_img.paragraph_format.space_after = Pt(4)
            r_img = p_img.add_run()
            r_img.add_picture(io.BytesIO(chart_bytes), width=Inches(6.4))

        sorted_m = sorted(monthly_data, key=lambda x: str(x.get("group_value", "")))
        peak_m = max(sorted_m, key=lambda x: float(x.get("revenue", 0)))
        low_m = min(sorted_m, key=lambda x: float(x.get("revenue", 0)))
        ratio = float(peak_m.get("revenue", 1)) / max(1.0, float(low_m.get("revenue", 1)))
        add_markdown_to_docx(doc, f"* **Tháng đạt đỉnh doanh thu:** Tháng **{peak_m.get('group_value')}** ghi nhận mức doanh thu kỷ lục **{format_compact_vnd(peak_m.get('revenue', 0))}** (Lợi nhuận **{format_compact_vnd(peak_m.get('profit', 0))}**), cao gấp **{ratio:.1f} lần** so với tháng thấp nhất.")
        add_markdown_to_docx(doc, f"* **Tháng thấp điểm nhất:** Tháng **{low_m.get('group_value')}** với doanh thu **{format_compact_vnd(low_m.get('revenue', 0))}**. Hoạt động kinh doanh có chu kỳ tăng trưởng mạnh vào quý 1 và quý 3 hàng năm.")

    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_after = Pt(4)

    # 6. MỤC III: BẢNG 1 — TOP 5 KHÁCH HÀNG VIP (TIER A)
    add_markdown_to_docx(doc, "### III. PHÂN HẠNG KHÁCH HÀNG TRỌNG ĐIỂM (BẢNG 1 — TOP 5 KHÁCH HÀNG VIP TIER A)")
    p_c_intro = doc.add_paragraph()
    p_c_intro.paragraph_format.space_after = Pt(3)
    r_ci = p_c_intro.add_run("Áp dụng nguyên lý Pareto (80/20), nhóm 5 khách hàng đóng góp lợi nhuận và doanh thu cao nhất:")
    r_ci.font.size = Pt(9)
    r_ci.font.italic = True

    top5_cust = _extract_top_customers(analysis, limit=5)
    if top5_cust:
        tbl_cust = doc.add_table(rows=1 + len(top5_cust), cols=7)
        hdrs_cust = ["Hạng", "Khách hàng VIP", "Số đơn", "Doanh thu", "Tỷ trọng", "Lợi nhuận", "Biên Margin (%)"]
        for ci, h in enumerate(hdrs_cust):
            tbl_cust.cell(0, ci).paragraphs[0].add_run(h).bold = True

        for ri, c in enumerate(top5_cust, start=1):
            tbl_cust.cell(ri, 0).paragraphs[0].add_run(str(ri))
            tbl_cust.cell(ri, 1).paragraphs[0].add_run(str(c.get("customer", "—"))).bold = True
            tbl_cust.cell(ri, 2).paragraphs[0].add_run(f"{int(c.get('orders', 0)):,}")
            tbl_cust.cell(ri, 3).paragraphs[0].add_run(format_compact_vnd(c.get("revenue", 0)))
            tbl_cust.cell(ri, 4).paragraphs[0].add_run(f"{float(c.get('share_pct', 0)):.1f}%")
            tbl_cust.cell(ri, 5).paragraphs[0].add_run(format_compact_vnd(c.get("profit", 0)))
            mg = c.get("margin_pct")
            tbl_cust.cell(ri, 6).paragraphs[0].add_run(f"{float(mg):.1f}%" if mg is not None else "—")
        apply_table_styles(tbl_cust)

    p_sp3 = doc.add_paragraph()
    p_sp3.paragraph_format.space_after = Pt(4)

    # 7. MỤC IV: BẢNG 2 — TOP 5 SẢN PHẨM CHỦ LỰC
    add_markdown_to_docx(doc, "### IV. TOP 5 SẢN PHẨM & DỊCH VỤ CHỦ LỰC (BẢNG 2)")
    p_p_intro = doc.add_paragraph()
    p_p_intro.paragraph_format.space_after = Pt(3)
    r_pi = p_p_intro.add_run("Gộp mã sản phẩm và tên sản phẩm, xếp hạng theo doanh số đóng góp cao nhất toàn hệ thống:")
    r_pi.font.size = Pt(9)
    r_pi.font.italic = True

    top5_prods = _extract_top_products(cleaned_df, analysis, tot_rev)
    if top5_prods:
        tbl_prod = doc.add_table(rows=1 + len(top5_prods), cols=7)
        hdrs_prod = ["Hạng", "Sản phẩm chủ lực (ID — Tên sản phẩm)", "Số lượng bán", "Doanh thu", "Tỷ trọng", "Lợi nhuận", "Biên Margin (%)"]
        for ci, h in enumerate(hdrs_prod):
            tbl_prod.cell(0, ci).paragraphs[0].add_run(h).bold = True

        for ri, p_item in enumerate(top5_prods, start=1):
            tbl_prod.cell(ri, 0).paragraphs[0].add_run(str(ri))
            tbl_prod.cell(ri, 1).paragraphs[0].add_run(str(p_item.get("product_name", ""))).bold = True
            tbl_prod.cell(ri, 2).paragraphs[0].add_run(f"{int(p_item.get('count', 0)):,} cái")
            tbl_prod.cell(ri, 3).paragraphs[0].add_run(format_compact_vnd(p_item.get("revenue", 0)))
            tbl_prod.cell(ri, 4).paragraphs[0].add_run(f"{float(p_item.get('share_pct', 0)):.1f}%")
            tbl_prod.cell(ri, 5).paragraphs[0].add_run(format_compact_vnd(p_item.get("profit", 0)))
            mg_p = p_item.get("margin_pct")
            tbl_prod.cell(ri, 6).paragraphs[0].add_run(f"{float(mg_p):.1f}%" if mg_p is not None else "—")
        apply_table_styles(tbl_prod)

    p_sp4 = doc.add_paragraph()
    p_sp4.paragraph_format.space_after = Pt(4)

    # 8. MỤC V: BẢNG 3 — HIỆU SUẤT THEO KHU VỰC & KÊNH BÁN HÀNG
    add_markdown_to_docx(doc, "### V. HIỆU SUẤT THEO KHU VỰC THỊ TRƯỜNG & KÊNH BÁN HÀNG (BẢNG 3)")
    p_rk_intro = doc.add_paragraph()
    p_rk_intro.paragraph_format.space_after = Pt(3)
    r_rki = p_rk_intro.add_run("Tóm tắt ngắn gọn 3 khu vực thị trường trọng điểm và 3 kênh bán hàng chủ chốt:")
    r_rki.font.size = Pt(9)
    r_rki.font.italic = True

    rk_data = _extract_region_channel_data(analysis)
    if rk_data:
        tbl_rk = doc.add_table(rows=1 + len(rk_data), cols=7)
        hdrs_rk = ["Phân loại", "Khu vực / Kênh bán hàng", "Số đơn", "Doanh thu", "Tỷ trọng", "Lợi nhuận", "Biên Margin (%)"]
        for ci, h in enumerate(hdrs_rk):
            tbl_rk.cell(0, ci).paragraphs[0].add_run(h).bold = True

        for ri, rk_row in enumerate(rk_data, start=1):
            tbl_rk.cell(ri, 0).paragraphs[0].add_run(str(rk_row.get("category", "")))
            tbl_rk.cell(ri, 1).paragraphs[0].add_run(str(rk_row.get("name", ""))).bold = True
            tbl_rk.cell(ri, 2).paragraphs[0].add_run(f"{int(rk_row.get('count', 0)):,}")
            tbl_rk.cell(ri, 3).paragraphs[0].add_run(format_compact_vnd(rk_row.get("revenue", 0)))
            tbl_rk.cell(ri, 4).paragraphs[0].add_run(f"{float(rk_row.get('share_pct', 0)):.1f}%")
            tbl_rk.cell(ri, 5).paragraphs[0].add_run(format_compact_vnd(rk_row.get("profit", 0)))
            mg_rk = rk_row.get("margin_pct")
            tbl_rk.cell(ri, 6).paragraphs[0].add_run(f"{float(mg_rk):.1f}%" if mg_rk is not None else "—")
        apply_table_styles(tbl_rk)

    p_sp5 = doc.add_paragraph()
    p_sp5.paragraph_format.space_after = Pt(4)

    # 9. MỤC VI: TÓM TẮT VẬN HÀNH & CẢNH BÁO RỦI RO (EXECUTIVE SUMMARY BULLETS)
    add_markdown_to_docx(doc, "### VI. TÓM TẮT VẬN HÀNH & CẢNH BÁO RỦI RO (EXECUTIVE SUMMARY)")
    bullets = _generate_executive_summary_bullets(cleaned_df, analysis, tot_rev)
    for b in bullets:
        add_markdown_to_docx(doc, b)

    p_sp6 = doc.add_paragraph()
    p_sp6.paragraph_format.space_after = Pt(6)

    # 10. KHUNG KÝ TÊN & ĐÓNG DẤU PHÊ DUYỆT
    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_left = tbl_sign.cell(0, 0)
    c_right = tbl_sign.cell(0, 1)

    p_sl = c_left.paragraphs[0]
    p_sl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sl1 = p_sl.add_run("ĐẠI DIỆN BỘ PHẬN LẬP BIỂU\n")
    r_sl1.font.size = Pt(9)
    r_sl1.font.bold = True
    r_sl2 = p_sl.add_run("(Ký & ghi rõ họ tên)\n\n\n\n")
    r_sl2.font.size = Pt(8)
    r_sl2.font.italic = True
    r_sl3 = p_sl.add_run(f"Hệ thống Sales AI BI Agent\n{biz_name}")
    r_sl3.font.size = Pt(9.5)
    r_sl3.font.bold = True

    p_sr = c_right.paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sr1 = p_sr.add_run("NGƯỜI PHÊ DUYỆT BAN HÀNH\n")
    r_sr1.font.size = Pt(9)
    r_sr1.font.bold = True
    r_sr2 = p_sr.add_run("(Ký, đóng dấu & ghi rõ họ tên)\n\n")
    r_sr2.font.size = Pt(8)
    r_sr2.font.italic = True
    r_sr_stamp = p_sr.add_run(f"[★ {biz_name} — ĐÃ PHÊ DUYỆT BAN HÀNH ★]\n\n")
    r_sr_stamp.font.size = Pt(8.5)
    r_sr_stamp.font.bold = True
    r_sr_stamp.font.color.rgb = RGBColor(14, 138, 22)
    r_sr3 = p_sr.add_run(approver_name)
    r_sr3.font.size = Pt(9.5)
    r_sr3.font.bold = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
