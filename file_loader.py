"""
Smart Multi-Format Spreadsheet & File Loader cho Minh An AI Agent
Hỗ trợ: .xlsx, .xlsm, .xls, .csv (đa dấu phân cách & encoding), .tsv, .json
Tự động dò dòng tiêu đề (header row) kể cả khi file có dòng banner trang trí phía trên,
nạp lại sheet tham chiếu (DanhMuc_Gia, TonKho_TrienKhai) và gộp nhiều dòng cùng mã yêu cầu.
"""

import os
import io
import csv
import json
import base64
import pandas as pd
from typing import Dict, List, Tuple, Any

HEADER_KEYWORDS = {
    'khach_hang', 'ten_khach_hang', 'customer', 'cong_ty', 'noi_dung_yeu_cau',
    'noi_dung', 'yeu_cau', 'id_yeu_cau', 'ma_yeu_cau', 'ngan_sach', 'ngan_sach_cong_bo',
    'thoi_han', 'thoi_han_ngay', 'ma_sp', 'sku', 'email', 'email_kh'
}

def _detect_header_row(raw_df: pd.DataFrame) -> int:
    """Dò tìm dòng header thực sự trong 10 dòng đầu tiên của Excel/CSV."""
    max_scan = min(10, len(raw_df))
    best_row = 0
    best_score = -1

    for r_idx in range(max_scan):
        row_vals = [str(v).strip().lower() for v in raw_df.iloc[r_idx].tolist() if pd.notna(v)]
        score = sum(1 for v in row_vals if v in HEADER_KEYWORDS or any(k in v for k in ('khach', 'yeu_cau', 'noi_dung', 'ngan_sach', 'thoi_han', 'ma_sp', 'email')))
        if score > best_score:
            best_score = score
            best_row = r_idx

    return best_row if best_score >= 1 else 0

def _read_excel_smart(temp_path: str, engine) -> Tuple[pd.DataFrame, List[str]]:
    excel_file = pd.ExcelFile(temp_path)
    sheets = excel_file.sheet_names
    notes = []

    # Nạp lại Master Data nếu file Excel mới có kèm các sheet tham chiếu
    cat_sheet = next((s for s in sheets if s.lower() in ('danhmuc_gia', 'catalog', 'sanpham', 'bang_gia')), None)
    inv_sheet = next((s for s in sheets if s.lower() in ('tonkho_trienkhai', 'inventory', 'ton_kho', 'tonkho')), None)
    pol_sheet = next((s for s in sheets if s.lower() in ('chinhsach_duyet', 'policy', 'chinh_sach')), None)
    rep_sheet = next((s for s in sheets if s.lower() in ('nhanvienkinhdoanh', 'sales_reps', 'nvkd')), None)

    new_cat = pd.read_excel(temp_path, sheet_name=cat_sheet).to_dict('records') if cat_sheet else None
    new_inv = pd.read_excel(temp_path, sheet_name=inv_sheet).to_dict('records') if inv_sheet else None
    new_pol = pd.read_excel(temp_path, sheet_name=pol_sheet).to_dict('records') if pol_sheet else None
    new_rep = pd.read_excel(temp_path, sheet_name=rep_sheet).to_dict('records') if rep_sheet else None

    if new_cat or new_inv or new_pol or new_rep:
        engine.reload_master_data(new_cat, new_inv, new_pol, new_rep)
        notes.append("Đã cập nhật Master Data từ các sheet tham chiếu")

    # Tìm sheet chứa yêu cầu khách hàng
    req_sheet = next((s for s in sheets if s.lower() in ('yeucaukhachhang', 'yeu_cau', 'inquiries', 'requests', 'data')), sheets[0])
    raw_preview = pd.read_excel(temp_path, sheet_name=req_sheet, header=None, nrows=12)
    header_row = _detect_header_row(raw_preview)
    df = pd.read_excel(temp_path, sheet_name=req_sheet, header=header_row)
    df = df.dropna(how='all')

    if header_row > 0:
        notes.append(f"Tự động bỏ qua {header_row} dòng banner đầu")
    notes.append(f"Sheet [{req_sheet}] ({len(df)} dòng)")
    return df, notes

def _read_delimited_smart(raw_bytes: bytes, ext: str) -> Tuple[pd.DataFrame, List[str]]:
    notes = []
    text = None
    used_enc = 'utf-8-sig'
    for enc in ('utf-8-sig', 'utf-8', 'cp1258', 'latin1'):
        try:
            text = raw_bytes.decode(enc)
            used_enc = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw_bytes.decode('utf-8', errors='replace')

    delimiter = '\t' if ext == '.tsv' else ','
    try:
        sample = '\n'.join(text.splitlines()[:10])
        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
        delimiter = dialect.delimiter
    except Exception:
        if ';' in text.splitlines()[0] and ',' not in text.splitlines()[0]:
            delimiter = ';'

    raw_preview = pd.read_csv(io.StringIO(text), sep=delimiter, header=None, nrows=12)
    header_row = _detect_header_row(raw_preview)
    df = pd.read_csv(io.StringIO(text), sep=delimiter, header=header_row)
    df = df.dropna(how='all')
    notes.append(f"Encoding {used_enc}, phân cách '{delimiter}' ({len(df)} dòng)")
    return df, notes

def _group_multi_row_inquiries(rows: List[Dict], engine) -> List[Dict]:
    """
    Nếu file Excel có nhiều dòng lặp lại cùng `id_yeu_cau` (mỗi dòng 1 mã SP),
    tự động gộp chúng thành 1 cơ hội nhiều sản phẩm.
    """
    if not rows:
        return rows

    grouped = {}
    order_keys = []

    for idx, r in enumerate(rows):
        norm = engine._normalize_meta_aliases(r)
        req_id = str(norm.get('id_yeu_cau') or '').strip()
        # Chỉ gộp khi có id_yeu_cau rõ ràng (không rỗng)
        key = req_id if req_id and req_id.lower() not in ('nan', 'none', '') else f"__ROW_{idx}"

        if key not in grouped:
            grouped[key] = dict(norm)
            grouped[key]['_grouped_items'] = []
            order_keys.append(key)

        # Nếu dòng này có cột sản phẩm cấu trúc, thu thập vào _grouped_items
        struct_items = engine._parse_structured_items(norm)
        for si in struct_items:
            existing = next((x for x in grouped[key]['_grouped_items'] if x['ma_sp'] == si['ma_sp']), None)
            if existing:
                existing['so_luong'] += si['so_luong']
            else:
                grouped[key]['_grouped_items'].append(si)

    return [grouped[k] for k in order_keys]

def load_uploaded_file_smart(filename: str, content: Any, temp_dir: str, engine) -> Tuple[List[Dict], str]:
    """Điểm vào chính để đọc file .xlsx, .xlsm, .xls, .csv, .tsv, .json."""
    ext = os.path.splitext(filename.lower())[1]
    os.makedirs(temp_dir, exist_ok=True)
    temp_path = os.path.join(temp_dir, f"temp_smart_{os.path.basename(filename)}")
    notes = []

    try:
        if ext in ('.xlsx', '.xlsm', '.xls'):
            raw_bytes = base64.b64decode(content) if isinstance(content, str) else content
            with open(temp_path, "wb") as f:
                f.write(raw_bytes)
            df, notes = _read_excel_smart(temp_path, engine)
        elif ext in ('.csv', '.tsv', '.txt'):
            raw_bytes = content.encode('utf-8') if isinstance(content, str) else content
            df, notes = _read_delimited_smart(raw_bytes, ext)
        elif ext == '.json':
            parsed_json = json.loads(content) if isinstance(content, str) else content
            records = parsed_json.get('rows', parsed_json) if isinstance(parsed_json, dict) else parsed_json
            df = pd.DataFrame(records)
            notes.append(f"JSON ({len(df)} bản ghi)")
        else:
            raise ValueError(f"Định dạng '{ext}' chưa hỗ trợ. Vui lòng dùng .xlsx, .xlsm, .xls, .csv, .tsv hoặc .json.")

        # Chuyển NaN sang None sạch sẽ trước khi trả về
        df = df.where(pd.notnull(df), None)
        raw_rows = df.to_dict('records')
        merged_rows = _group_multi_row_inquiries(raw_rows, engine)

        if len(merged_rows) < len(raw_rows):
            notes.append(f"Đã gộp {len(raw_rows)} dòng thành {len(merged_rows)} yêu cầu")

        note_str = f"Đã đọc [{filename}]: " + " • ".join(notes)
        return merged_rows, note_str
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

