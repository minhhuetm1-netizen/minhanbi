"""
Agent Engine: Quy trình Xử lý Cơ hội Bán hàng B2B - Công ty Minh An (AI301 - Đề 01)
Hỗ trợ đa định dạng đầu vào (Excel nhiều cấu trúc, CSV/TSV đa dấu phân cách, JSON, Google Sheet, Text)
và đầy đủ các trường thông tin doanh nghiệp B2B (MST, Địa chỉ giao, Kênh, Ưu tiên, Thanh toán, Người duyệt KH, Nhiều SP).
"""

import os
import re
import json
import math
import uuid
import datetime
from typing import Dict, List, Any, Optional

AUDIT_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "agent_audit_logs.jsonl")

def _is_empty_or_nan(val) -> bool:
    if val is None:
        return True
    if isinstance(val, str):
        v = val.strip().lower()
        return v in ("", "nan", "none", "null", "n/a", "na", "-", "chưa rõ", "chua ro", "tbd", "không", "khong")
    if isinstance(val, (int, float)):
        return math.isnan(val)
    return False

class MinhAnAgentEngine:
    def __init__(self, catalog_path="data/catalog_pricing.csv", 
                 inventory_path="data/inventory_status.csv",
                 policy_path="data/discount_policy.csv",
                 sales_reps_path="data/sales_reps.csv"):
        self.catalog = self._load_csv_records(catalog_path)
        self.inventory = self._load_inventory(inventory_path)
        self.policies = self._load_csv_records(policy_path)
        self.sales_reps = self._load_csv_records(sales_reps_path)

    def _load_csv_records(self, path) -> List[Dict]:
        try:
            import pandas as pd
            df = pd.read_csv(path, encoding='utf-8-sig')
            return df.to_dict('records')
        except Exception:
            return []

    def _load_inventory(self, path) -> Dict[str, Dict]:
        records = self._load_csv_records(path)
        return {r['ma_sp']: r for r in records if 'ma_sp' in r}

    def reload_master_data(self, catalog_records=None, inventory_records=None, policy_records=None, reps_records=None):
        """Nạp lại bảng giá / tồn kho / chính sách / NVKD khi file Excel mới đi kèm các sheet tham chiếu."""
        if catalog_records:
            self.catalog = catalog_records
        if inventory_records:
            self.inventory = {r['ma_sp']: r for r in inventory_records if 'ma_sp' in r}
        if policy_records:
            self.policies = policy_records
        if reps_records:
            self.sales_reps = reps_records

    def _normalize_meta_aliases(self, meta: Dict) -> Dict:
        """Chuẩn hóa tên cột linh hoạt (tiếng Việt có/không dấu, tiếng Anh, viết tắt ERP/CRM)."""
        if not meta:
            return {}
        # Chuẩn hóa key về lowercase không khoảng trắng thừa để match linh hoạt hơn
        lower_map = {str(k).strip().lower(): v for k, v in meta.items()}
        m = dict(meta)

        alias_map = {
            'id_yeu_cau': ['id_yeu_cau', 'ma_yeu_cau', 'ma_yc', 'request_id', 'id', 'stt'],
            'ngay_nhan': ['ngay_nhan', 'ngay_tiep_nhan', 'ngay_gui', 'received_date', 'created_at', 'ngay'],
            'kenh_tiep_nhan': ['kenh_tiep_nhan', 'kenh', 'nguon', 'source', 'channel'],
            'khach_hang': ['khach_hang', 'ten_khach_hang', 'ten_cong_ty', 'cong_ty', 'doanh_nghiep', 'ten_kh', 'customer', 'company', 'client'],
            'ma_so_thue': ['ma_so_thue', 'mst', 'tax_id', 'tax_code'],
            'dia_chi_giao_hang': ['dia_chi_giao_hang', 'dia_chi', 'khu_vuc', 'tinh_thanh', 'address', 'delivery_address', 'location'],
            'nguoi_lien_he': ['nguoi_lien_he', 'ten_nguoi_lien_he', 'contact_person', 'contact', 'nguoi_phu_trach'],
            'email_kh': ['email_kh', 'email', 'email_khach_hang', 'mail'],
            'sdt_kh': ['sdt_kh', 'sdt', 'so_dien_thoai', 'dien_thoai', 'phone', 'mobile', 'tel'],
            'nguoi_phe_duyet_kh': ['nguoi_phe_duyet_kh', 'nguoi_phe_duyet', 'nguoi_duyet_kh', 'approver_client', 'client_approver'],
            'muc_do_uu_tien': ['muc_do_uu_tien', 'uu_tien', 'priority', 'do_khan'],
            'hinh_thuc_thanh_toan': ['hinh_thuc_thanh_toan', 'thanh_toan', 'dieu_khoan_thanh_toan', 'payment_terms', 'payment'],
            'noi_dung_yeu_cau': ['noi_dung_yeu_cau', 'noi_dung', 'yeu_cau', 'mo_ta', 'chi_tiet_yeu_cau', 'email_body', 'content', 'description'],
            'ngan_sach_cong_bo': ['ngan_sach_cong_bo', 'ngan_sach', 'ngan_sach_du_kien', 'ngan_sach_tong', 'budget', 'estimated_budget'],
            'thoi_han_ngay': ['thoi_han_ngay', 'thoi_han', 'thoi_han_trien_khai', 'deadline_days', 'so_ngay_giao', 'deadline', 'ngay_giao_mong_muon'],
            'chiet_khau_de_xuat_pct': ['chiet_khau_de_xuat_pct', 'chiet_khau', 'de_nghi_chiet_khau', 'muc_chiet_khau', 'discount', 'discount_pct', 'ck_pct', 'ck'],
            'ma_sp': ['ma_sp', 'sku', 'ma_san_pham', 'product_code', 'danh_sach_sp'],
            'ten_san_pham': ['ten_san_pham', 'san_pham', 'ten_sp', 'hang_hoa', 'product_name'],
            'so_luong': ['so_luong', 'sl', 'qty', 'quantity'],
            'yeu_cau_ky_thuat': ['yeu_cau_ky_thuat', 'ghi_chu_ky_thuat', 'ghi_chu', 'notes', 'remarks', 'tech_req']
        }

        for target, keys in alias_map.items():
            if target not in m or _is_empty_or_nan(m.get(target)):
                for k in keys:
                    if k in lower_map and not _is_empty_or_nan(lower_map[k]):
                        m[target] = lower_map[k]
                        break
        return m

    def _parse_budget_value(self, raw_val, text_lower: str) -> Optional[float]:
        """Hỗ trợ nhiều định dạng ngân sách: 270000000, '270.000.000 đ', '270 triệu', '270tr', '270M', '0.5 tỷ'."""
        if not _is_empty_or_nan(raw_val):
            if isinstance(raw_val, (int, float)):
                num = float(raw_val)
                if num <= 0:
                    return None
                return num * 1_000_000 if num < 5000 else num
            s = str(raw_val).strip().lower()
            m = re.search(r'(\d+(?:[\.,]\d+)?)\s*(tỷ|ty|triệu|trieu|tr|m|vnđ|vnd|đ|d)?', s)
            if m:
                raw_num = m.group(1)
                unit = m.group(2)
                # Nếu có nhiều dấu chấm phân cách hàng nghìn (vd 270.000.000)
                if raw_num.count('.') > 1 or (raw_num.count('.') == 1 and len(raw_num.split('.')[1]) == 3 and not unit):
                    clean_num = float(raw_num.replace('.', '').replace(',', ''))
                elif raw_num.count(',') > 1 or (raw_num.count(',') == 1 and len(raw_num.split(',')[1]) == 3 and not unit):
                    clean_num = float(raw_num.replace(',', '').replace('.', ''))
                else:
                    clean_num = float(raw_num.replace(',', '.'))

                if unit in ('tỷ', 'ty'):
                    return clean_num * 1_000_000_000
                elif unit in ('triệu', 'trieu', 'tr', 'm'):
                    return clean_num * 1_000_000
                else:
                    return clean_num * 1_000_000 if clean_num < 5000 else clean_num

        # Fallback: tìm trong văn bản
        budget_match = re.search(r'ngân sách[^\d]*(\d+(?:[\.,]\d+)*)\s*(tỷ|ty|triệu|trieu|tr|m|vnđ|vnd|đ)?', text_lower)
        if budget_match:
            val_str = budget_match.group(1)
            unit = budget_match.group(2)
            try:
                if val_str.count('.') > 1 or (val_str.count('.') == 1 and len(val_str.split('.')[1]) == 3 and unit not in ('tỷ', 'ty')):
                    num = float(val_str.replace('.', '').replace(',', ''))
                else:
                    num = float(val_str.replace(',', '.'))
                if unit in ('tỷ', 'ty'):
                    return num * 1_000_000_000
                elif unit in ('triệu', 'trieu', 'tr', 'm'):
                    return num * 1_000_000
                return num * 1_000_000 if num < 5000 else num
            except ValueError:
                return None
        return None

    def _parse_deadline_days(self, raw_val, text_lower: str, ref_date_val=None) -> Optional[int]:
        """Hỗ trợ nhiều định dạng thời hạn: số ngày, '2 tuần', '1 tháng', hoặc ngày dương lịch Excel (2026-11-25, 25/11/2026)."""
        ref_date = datetime.date.today()
        if not _is_empty_or_nan(ref_date_val):
            try:
                if hasattr(ref_date_val, 'date'):
                    ref_date = ref_date_val.date()
                else:
                    ref_str = str(ref_date_val).strip()[:10]
                    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
                        try:
                            ref_date = datetime.datetime.strptime(ref_str, fmt).date()
                            break
                        except ValueError:
                            continue
            except Exception:
                pass

        if not _is_empty_or_nan(raw_val):
            # Trường hợp Excel trả về Timestamp / datetime.date
            if hasattr(raw_val, 'date') or isinstance(raw_val, datetime.date):
                target_d = raw_val.date() if hasattr(raw_val, 'date') else raw_val
                diff = (target_d - ref_date).days
                return max(1, diff)
            if isinstance(raw_val, (int, float)):
                return max(1, int(float(raw_val))) if raw_val > 0 else None

            s = str(raw_val).strip().lower()
            # Kiểm tra định dạng ngày: dd/mm/yyyy hoặc yyyy-mm-dd
            date_m = re.search(r'(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}/\d{1,2}/\d{4}|\d{1,2}-\d{1,2}-\d{4})', s)
            if date_m:
                d_str = date_m.group(1)
                for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
                    try:
                        target_d = datetime.datetime.strptime(d_str, fmt).date()
                        return max(1, (target_d - ref_date).days)
                    except ValueError:
                        continue

            unit_m = re.search(r'(\d+)\s*(ngày|ngay|days|day|tuần|tuan|weeks|week|tháng|thang|months|month)?', s)
            if unit_m:
                n = int(unit_m.group(1))
                u = unit_m.group(2) or 'ngày'
                if 'tuần' in u or 'tuan' in u or 'week' in u:
                    return n * 7
                if 'tháng' in u or 'thang' in u or 'month' in u:
                    return n * 30
                return n

        deadline_match = re.search(r'(trong vòng|thời hạn|tiến độ|giao trong|trong)\s*(\d+)\s*(ngày|tuần|tháng)', text_lower)
        if deadline_match:
            d_num = int(deadline_match.group(2))
            d_unit = deadline_match.group(3)
            if d_unit == 'ngày':
                return d_num
            elif d_unit == 'tuần':
                return d_num * 7
            elif d_unit == 'tháng':
                return d_num * 30
        return None

    def _parse_discount_pct(self, raw_val, text_lower: str) -> Optional[float]:
        """Hỗ trợ: 8, '8%', '8.5 %', và định dạng phần trăm Excel (0.08 -> 8.0%)."""
        if raw_val == 0 or raw_val == "0" or raw_val == 0.0:
            return 0.0
        if not _is_empty_or_nan(raw_val):
            if isinstance(raw_val, (int, float)):
                val = float(raw_val)
                # Excel lưu 8% dưới dạng 0.08 -> tự động nhân 100
                if 0 < val < 1.0:
                    return round(val * 100.0, 2)
                return round(val, 2)
            s = str(raw_val).strip().replace('%', '').replace(',', '.')
            try:
                val = float(re.search(r'\d+(?:\.\d+)?', s).group(0))
                if 0 < val < 1.0 and '%' not in str(raw_val):
                    return round(val * 100.0, 2)
                return round(val, 2)
            except Exception:
                pass

        disc_match = re.search(r'chiết khấu[^\d]*(\d+(?:[\.,]\d+)?)\s*%', text_lower)
        if disc_match:
            try:
                return float(disc_match.group(1).replace(',', '.'))
            except ValueError:
                return None
        return None

    def _parse_structured_items(self, meta: Dict) -> List[Dict]:
        """Đọc sản phẩm nếu file Excel có cột ma_sp / san_pham / so_luong tách riêng (hoặc chuỗi 'LAP-DELL-5540:10; FW-FORTI-60F:1')."""
        items = []
        raw_sku = meta.get('ma_sp')
        raw_name = meta.get('ten_san_pham')
        raw_qty = meta.get('so_luong')

        if _is_empty_or_nan(raw_sku) and _is_empty_or_nan(raw_name):
            return items

        # Trường hợp nhiều SP trong 1 ô phân tách bằng dấu ; hoặc , hoặc xuống dòng
        sku_str = str(raw_sku or raw_name or '')
        parts = [p.strip() for p in re.split(r'[;\n]+', sku_str) if p.strip()]
        for part in parts:
            qty = 1
            # Check dạng "LAP-DELL-5540: 10" hoặc "LAP-DELL-5540 x 10"
            m = re.match(r'^([A-Za-z0-9\-_]+)\s*[:xX\*]\s*(\d+)$', part)
            if m:
                token, qty = m.group(1), int(m.group(2))
            else:
                token = part
                if not _is_empty_or_nan(raw_qty) and len(parts) == 1:
                    try:
                        qty = max(1, int(float(raw_qty)))
                    except Exception:
                        qty = 1

            token_lower = token.lower()
            cat_prod = next((p for p in self.catalog if p['ma_sp'].lower() == token_lower or token_lower in p['ten_san_pham'].lower()), None)
            if cat_prod:
                items.append({
                    "ma_sp": cat_prod['ma_sp'],
                    "ten_san_pham": cat_prod['ten_san_pham'],
                    "danh_muc": cat_prod.get('danh_muc', 'Thiết bị công nghệ'),
                    "so_luong": qty,
                    "don_gia_niem_yet": cat_prod['gia_niem_yet'],
                    "gia_von": cat_prod.get('gia_von', cat_prod['gia_niem_yet'] * 0.8)
                })
        return items

    def parse_raw_text(self, text: str, meta: Optional[Dict] = None) -> Dict[str, Any]:
        """Kỹ năng 1: Bóc tách đa trường dữ liệu và chuẩn hóa nhiều định dạng từ Excel/CSV/Text."""
        meta = self._normalize_meta_aliases(meta or {})
        text = str(text or meta.get('noi_dung_yeu_cau') or '')
        text_lower = text.lower()

        ten_kh = str(meta.get('khach_hang') or "Doanh nghiệp Khách hàng").strip()
        nguoi_lh = str(meta.get('nguoi_lien_he') or "").strip() if not _is_empty_or_nan(meta.get('nguoi_lien_he')) else ""
        email_kh = str(meta.get('email_kh') or "").strip() if not _is_empty_or_nan(meta.get('email_kh')) else ""
        sdt_kh = str(meta.get('sdt_kh') or "").strip() if not _is_empty_or_nan(meta.get('sdt_kh')) else ""

        # Chuẩn hóa SĐT nếu Excel đọc thành int mất số 0 đầu (vd 912345678 -> 0912345678)
        if sdt_kh and sdt_kh.isdigit() and len(sdt_kh) == 9 and not sdt_kh.startswith('0'):
            sdt_kh = '0' + sdt_kh

        if not email_kh:
            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
            if email_match:
                email_kh = email_match.group(0)
        if not sdt_kh:
            phone_match = re.search(r'(0\d{9,10}|\+84\d{9,10})', text)
            if phone_match:
                sdt_kh = phone_match.group(0)

        # Trường mở rộng B2B
        mst = str(meta.get('ma_so_thue') or "").strip() if not _is_empty_or_nan(meta.get('ma_so_thue')) else ""
        dia_chi = str(meta.get('dia_chi_giao_hang') or "").strip() if not _is_empty_or_nan(meta.get('dia_chi_giao_hang')) else ""
        kenh = str(meta.get('kenh_tiep_nhan') or "Email / Biểu mẫu").strip() if not _is_empty_or_nan(meta.get('kenh_tiep_nhan')) else "Email / Biểu mẫu"
        uu_tien = str(meta.get('muc_do_uu_tien') or "Tiêu chuẩn").strip() if not _is_empty_or_nan(meta.get('muc_do_uu_tien')) else "Tiêu chuẩn"
        thanh_toan = str(meta.get('hinh_thuc_thanh_toan') or "Tiêu chuẩn (Tạm ứng/TT)").strip() if not _is_empty_or_nan(meta.get('hinh_thuc_thanh_toan')) else "Tiêu chuẩn (Tạm ứng/TT)"
        nguoi_duyet_kh = str(meta.get('nguoi_phe_duyet_kh') or "").strip() if not _is_empty_or_nan(meta.get('nguoi_phe_duyet_kh')) else ""
        yeu_cau_kt = str(meta.get('yeu_cau_ky_thuat') or "").strip() if not _is_empty_or_nan(meta.get('yeu_cau_ky_thuat')) else ""

        budget = self._parse_budget_value(meta.get('ngan_sach_cong_bo'), text_lower)
        deadline_days = self._parse_deadline_days(meta.get('thoi_han_ngay'), text_lower, meta.get('ngay_nhan'))
        discount_req = self._parse_discount_pct(meta.get('chiet_khau_de_xuat_pct'), text_lower)

        # Ưu tiên lấy SP từ cột cấu trúc (hoặc danh sách gộp nhiều dòng Excel), sau đó bổ sung từ bóc tách văn bản
        detected_items = []
        matched_skus = set()

        # Nếu đã có items gộp sẵn từ multi-row grouping ở server
        if isinstance(meta.get('_grouped_items'), list):
            for gi in meta['_grouped_items']:
                detected_items.append(gi)
                matched_skus.add(gi['ma_sp'])

        for si in self._parse_structured_items(meta):
            if si['ma_sp'] not in matched_skus:
                detected_items.append(si)
                matched_skus.add(si['ma_sp'])

        cleaned_text = re.sub(r'\b24/7\b', '', text_lower)
        cleaned_text = re.sub(r'\b\d+\s*(tháng|năm)\b', '', cleaned_text)

        product_patterns = [
            ("LAP-DELL-5540", r'(dell\s*latitude\s*5540|latitude\s*5540|laptop\s*dell)'),
            ("LAP-LEN-T14", r'(lenovo\s*thinkpad\s*t14|thinkpad\s*t14|laptop\s*thinkpad|laptop\s*lenovo)'),
            ("SRV-DELL-R750", r'(dell\s*poweredge\s*r750|poweredge\s*r750|máy chủ dell|server dell)'),
            ("SW-CISCO-C9200", r'(cisco\s*catalyst\s*c9200|c9200l|switch\s*cisco|switch\s*catalyst)'),
            ("FW-FORTI-60F", r'(fortigate\s*60f|fortinet\s*60f|tường lửa fortinet|firewall fortinet)'),
            ("SVC-SEC-1Y", r'(dịch vụ triển khai & cấu hình an toàn|cấu hình an toàn mạng|triển khai bảo mật|cài đặt mạng)'),
            ("SVC-MAINT-247", r'(dịch vụ bảo trì & hỗ trợ kỹ thuật|bảo trì hệ thống|bảo trì it)'),
            ("LIC-M365-BUS", r'(microsoft\s*365|m365\s*business|bản quyền m365|license office 365|license microsoft)')
        ]

        for sku, pattern in product_patterns:
            if sku in matched_skus:
                continue
            match = re.search(pattern, cleaned_text)
            if match:
                qty = self._extract_quantity_around(cleaned_text, match.start(), match.end())
                cat_prod = next((p for p in self.catalog if p['ma_sp'] == sku), None)
                if cat_prod:
                    matched_skus.add(sku)
                    detected_items.append({
                        "ma_sp": sku,
                        "ten_san_pham": cat_prod['ten_san_pham'],
                        "danh_muc": cat_prod.get('danh_muc', 'Thiết bị công nghệ'),
                        "so_luong": qty,
                        "don_gia_niem_yet": cat_prod['gia_niem_yet'],
                        "gia_von": cat_prod.get('gia_von', cat_prod['gia_niem_yet'] * 0.8)
                    })

        for cat_prod in self.catalog:
            sku = str(cat_prod.get('ma_sp', ''))
            if not sku or sku in matched_skus:
                continue
            match = re.search(re.escape(sku.lower()), cleaned_text)
            if match:
                qty = self._extract_quantity_around(cleaned_text, match.start(), match.end())
                matched_skus.add(sku)
                detected_items.append({
                    "ma_sp": sku,
                    "ten_san_pham": cat_prod['ten_san_pham'],
                    "danh_muc": cat_prod.get('danh_muc', 'Thiết bị công nghệ'),
                    "so_luong": qty,
                    "don_gia_niem_yet": cat_prod['gia_niem_yet'],
                    "gia_von": cat_prod.get('gia_von', cat_prod['gia_niem_yet'] * 0.8)
                })

        return {
            "id_yeu_cau": str(meta.get('id_yeu_cau') or ''),
            "ngay_nhan": str(meta.get('ngay_nhan') or datetime.date.today().strftime('%Y-%m-%d'))[:10],
            "kenh_tiep_nhan": kenh,
            "khach_hang": ten_kh,
            "ma_so_thue": mst,
            "dia_chi_giao_hang": dia_chi,
            "nguoi_lien_he": nguoi_lh,
            "email_kh": email_kh,
            "sdt_kh": sdt_kh,
            "nguoi_phe_duyet_kh": nguoi_duyet_kh,
            "muc_do_uu_tien": uu_tien,
            "hinh_thuc_thanh_toan": thanh_toan,
            "yeu_cau_ky_thuat": yeu_cau_kt,
            "noi_dung_goc": text,
            "ngan_sach_cong_bo": budget,
            "thoi_han_ngay": deadline_days,
            "chiet_khau_de_xuat_pct": discount_req,
            "items_yeu_cau": detected_items
        }

    def _extract_quantity_around(self, cleaned_text: str, start: int, end: int) -> int:
        preceding_text = cleaned_text[max(0, start - 40):start]
        following_text = cleaned_text[end:min(len(cleaned_text), end + 20)]
        qty = 1
        qty_pre = re.findall(r'(\d+)\s*(?:chiếc|bộ|thiết bị|máy|gói|license|user|con)?\s*$', preceding_text.strip())
        if qty_pre:
            try:
                qty = int(qty_pre[-1])
            except ValueError:
                qty = 1
        else:
            qty_post = re.search(r'^\s*[:=x\-]?\s*(\d+)', following_text)
            if qty_post:
                try:
                    qty = int(qty_post.group(1))
                except ValueError:
                    qty = 1
        return max(1, qty)

    def _write_audit_log(self, trace_id: str, input_ref: str, logs: List[Dict], decision: Dict):
        try:
            os.makedirs(os.path.dirname(AUDIT_LOG_PATH), exist_ok=True)
            entry = {
                "trace_id": trace_id,
                "timestamp": datetime.datetime.now().isoformat(),
                "input_ref": input_ref,
                "steps": logs,
                "decision": decision
            }
            with open(AUDIT_LOG_PATH, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception:
            pass

    def run_full_pipeline(self, raw_input: Dict[str, Any]) -> Dict[str, Any]:
        logs = []
        trace_id = f"TRC-{uuid.uuid4().hex[:6].upper()}"
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

        items = raw_input.get("items_yeu_cau", [])
        budget = raw_input.get("ngan_sach_cong_bo")
        deadline = raw_input.get("thoi_han_ngay")
        discount_req = raw_input.get("chiet_khau_de_xuat_pct")
        payment_terms = str(raw_input.get("hinh_thuc_thanh_toan") or "")

        logs.append({
            "step": 1,
            "skill": "Skill_1_TiepNhan_PhanLoai",
            "tool": "Tool: Read_MultiFormat_Input",
            "title": "Tiếp nhận yêu cầu",
            "status": "SUCCESS",
            "desc": f"[{trace_id}] Kênh: {raw_input.get('kenh_tiep_nhan', 'Email')} | KH: '{raw_input.get('khach_hang')}' | {len(items)} SP."
        })

        missing_fields = []
        if not items:
            missing_fields.append("Danh mục sản phẩm / Số lượng cụ thể")
        if _is_empty_or_nan(budget):
            missing_fields.append("Ngân sách dự kiến được duyệt")
        if _is_empty_or_nan(deadline):
            missing_fields.append("Thời hạn triển khai / giao hàng mong muốn")

        if len(missing_fields) > 0:
            logs.append({
                "step": 2,
                "skill": "Skill_1_TiepNhan_PhanLoai",
                "tool": "Tool: Validate_Required_Fields",
                "title": "Phân loại & Kiểm tra thông tin",
                "status": "WARNING",
                "desc": f"Thiếu {', '.join(missing_fields)}. KHÔNG tự xác nhận bán hàng."
            })

            draft_clarification_email = self._generate_clarification_email(raw_input, missing_fields)
            logs.append({
                "step": 2.5,
                "skill": "Skill_3_PhanHoi_KhachHang",
                "tool": "Tool: Draft_Clarification_Email",
                "title": "Yêu cầu KH làm rõ / bổ sung",
                "status": "ACTION_REQUIRED",
                "desc": "Phân loại: [CẦN BỔ SUNG THÔNG TIN]. Đã tạo bản nháp hỏi lại KH."
            })

            result = {
                "trace_id": trace_id,
                "quy_trinh_hoan_tat": False,
                "ket_qua_kiem_tra": "THIEU_THONG_TIN",
                "phan_loai_co_hoi": "CAN_BO_SUNG_THONG_TIN",
                "nhan_phan_loai": "Cần bổ sung thông tin",
                "du_dieu_kien_tinh_tong": False,
                "xac_nhan_ban_hang_tu_dong": False,
                "truong_con_thieu": missing_fields,
                "logs": logs,
                "buoc_hien_tai": "YeuCau_KH_LamRo",
                "trang_thai_crm": "CHO_LAM_RO",
                "can_cu": f"Thiếu dữ kiện bắt buộc: {', '.join(missing_fields)}.",
                "buoc_xu_ly_tiep_theo": "NVKD gửi Email đề nghị khách hàng bổ sung ngân sách/thời hạn trước khi lập báo giá.",
                "ban_nhap_phan_hoi": draft_clarification_email,
                "ho_so_co_hoi": None,
                "tham_dinh_ton_kho": None,
                "bang_bao_gia": None,
                "trinh_duyet": None
            }
            self._write_audit_log(trace_id, raw_input.get("khach_hang", ""), logs, {"phan_loai": "CAN_BO_SUNG_THONG_TIN"})
            return result

        logs.append({
            "step": 2,
            "skill": "Skill_1_TiepNhan_PhanLoai",
            "tool": "Tool: Validate_Required_Fields",
            "title": "Phân loại & Kiểm tra thông tin",
            "status": "SUCCESS",
            "desc": "Đầy đủ dữ kiện bắt buộc (SP/SL, Ngân sách, Thời hạn)."
        })

        opp_id = f"OPP-{datetime.datetime.now().strftime('%Y%m')}-{uuid.uuid4().hex[:4].upper()}"
        primary_category = items[0]['danh_muc'] if items else "Thiết bị đầu cuối"
        assigned_rep = self._assign_sales_rep(primary_category)

        opp_dossier = {
            "ma_co_hoi": opp_id,
            "khach_hang": raw_input.get("khach_hang"),
            "ma_so_thue": raw_input.get("ma_so_thue", ""),
            "dia_chi_giao_hang": raw_input.get("dia_chi_giao_hang", ""),
            "kenh_tiep_nhan": raw_input.get("kenh_tiep_nhan", "Email"),
            "muc_do_uu_tien": raw_input.get("muc_do_uu_tien", "Tiêu chuẩn"),
            "hinh_thuc_thanh_toan": payment_terms,
            "nguoi_phe_duyet_kh": raw_input.get("nguoi_phe_duyet_kh", ""),
            "yeu_cau_ky_thuat": raw_input.get("yeu_cau_ky_thuat", ""),
            "nguoi_lien_he": raw_input.get("nguoi_lien_he"),
            "email_kh": raw_input.get("email_kh"),
            "sdt_kh": raw_input.get("sdt_kh"),
            "nvkd_phu_trach": assigned_rep.get("ho_ten", "Nguyễn Văn An"),
            "ma_nv": assigned_rep.get("ma_nv", "NVKD01"),
            "email_nvkd": assigned_rep.get("email", "kd@minhan.com.vn"),
            "chuyen_mon": assigned_rep.get("chuyen_mon", "Thiết bị B2B"),
            "ngay_tao": now_str,
            "sla_phan_hoi_gio": 2 if "khẩn" in str(raw_input.get("muc_do_uu_tien", "")).lower() else 4,
            "han_phan_hoi": (datetime.datetime.now() + datetime.timedelta(hours=2 if "khẩn" in str(raw_input.get("muc_do_uu_tien", "")).lower() else 4)).strftime("%Y-%m-%d %H:%M"),
            "trang_thai": "ĐANG_XỬ_LÝ"
        }

        logs.append({
            "step": 3,
            "skill": "Skill_1_TiepNhan_PhanLoai",
            "tool": "Tool: Assign_Sales_Rep_Workload",
            "title": "Tạo Hồ sơ Cơ hội & Gắn người phụ trách",
            "status": "SUCCESS",
            "desc": f"Khởi tạo {opp_id}, gán {opp_dossier['nvkd_phu_trach']} (Ưu tiên: {opp_dossier['muc_do_uu_tien']}), SLA {opp_dossier['sla_phan_hoi_gio']}h."
        })

        inventory_checks = []
        overall_stock_ok = True
        max_lead_time_days = 0
        stock_alerts = []

        for it in items:
            sku = it['ma_sp']
            qty = it['so_luong']
            inv_info = self.inventory.get(sku, {
                "ton_kha_dung": 0, "thoi_gian_giao_san_co_ngay": 3, "thoi_gian_nhap_hang_ngay": 15, "kho_luu_tru": "Kho Tổng"
            })
            ton_hien_co = int(inv_info.get("ton_kha_dung", 0))
            is_enough = ton_hien_co >= qty

            if is_enough:
                lead_days = int(inv_info.get("thoi_gian_giao_san_co_ngay", 1))
                tinh_trang = "Có sẵn hàng trong kho"
            else:
                overall_stock_ok = False
                shortage = qty - ton_hien_co
                lead_days = int(inv_info.get("thoi_gian_nhap_hang_ngay", 10))
                tinh_trang = f"Thiếu {shortage} chiếc, cần nhập ({lead_days} ngày)"
                stock_alerts.append(f"{it['ten_san_pham']}: Kho còn {ton_hien_co}/{qty} (Thiếu {shortage}), nhập mất {lead_days} ngày.")

            if lead_days > max_lead_time_days:
                max_lead_time_days = lead_days

            inventory_checks.append({
                "ma_sp": sku,
                "ten_san_pham": it['ten_san_pham'],
                "so_luong_can": qty,
                "ton_kha_dung": ton_hien_co,
                "du_hang": is_enough,
                "thoi_gian_cung_ung_ngay": lead_days,
                "kho": inv_info.get("kho_luu_tru", "Kho Tổng"),
                "tinh_trang": tinh_trang
            })

        lead_time_risk = max_lead_time_days > deadline
        if lead_time_risk:
            lead_time_note = f"CẢNH BÁO KHÔNG ĐÁP ỨNG TIẾN ĐỘ: Khách cần trong {deadline} ngày, nhưng lead-time thực tế cần {max_lead_time_days} ngày!"
        else:
            lead_time_note = f"Tiến độ khả thi: Giao hàng trong {max_lead_time_days} ngày (Đáp ứng thời hạn {deadline} ngày của khách)."

        logs.append({
            "step": 4,
            "skill": "Skill_2_ThamDinh_TonKho_Gia",
            "tool": "Tool: Check_Inventory_ATP & Query_PriceBook",
            "title": "Thẩm định tồn kho & giá",
            "status": "WARNING" if lead_time_risk else "SUCCESS",
            "desc": lead_time_note
        })

        total_list_price = 0
        total_cogs = 0
        line_items = []
        applied_discount_pct = discount_req if discount_req is not None else 0.0

        for it in items:
            unit_price = float(it['don_gia_niem_yet'])
            cogs = float(it.get('gia_von', unit_price * 0.8))
            sub_total = unit_price * it['so_luong']
            sub_cogs = cogs * it['so_luong']
            total_list_price += sub_total
            total_cogs += sub_cogs

            line_items.append({
                "ma_sp": it['ma_sp'],
                "ten_san_pham": it['ten_san_pham'],
                "so_luong": it['so_luong'],
                "don_gia_niem_yet": unit_price,
                "thanh_tien_niem_yet": sub_total,
                "gia_von": cogs
            })

        discount_amount = total_list_price * (applied_discount_pct / 100.0)
        net_total_before_vat = total_list_price - discount_amount
        vat_amount = net_total_before_vat * 0.10
        final_total_with_vat = net_total_before_vat + vat_amount
        gross_profit = net_total_before_vat - total_cogs
        gross_margin_pct = (gross_profit / net_total_before_vat * 100.0) if net_total_before_vat > 0 else 0

        budget_variance = (budget - final_total_with_vat) if budget is not None else 0
        budget_status = "TRONG_NGAN_SACH" if budget_variance >= 0 else "VUOT_NGAN_SACH"

        quotation_summary = {
            "tong_gia_niem_yet": total_list_price,
            "chiet_khau_pct": applied_discount_pct,
            "tien_chiet_khau": discount_amount,
            "tong_truoc_vat": net_total_before_vat,
            "thue_vat_10pct": vat_amount,
            "tong_sau_vat": final_total_with_vat,
            "tong_gia_von": total_cogs,
            "loi_nhuan_gop": gross_profit,
            "ty_suat_loi_nhuan_gop_pct": round(gross_margin_pct, 1),
            "ngan_sach_khach": budget,
            "chenh_lech_ngan_sach": budget_variance,
            "tinh_trang_ngan_sach": budget_status,
            "chi_tiet_hang_hoa": line_items
        }

        logs.append({
            "step": 5,
            "skill": "Skill_2_ThamDinh_TonKho_Gia",
            "tool": "Tool: Calculate_Quote_And_Margin",
            "title": "Lập Báo giá & Dự thảo đề xuất",
            "status": "SUCCESS",
            "desc": f"Tổng thanh toán: {final_total_with_vat:,.0f} đ | CK: {applied_discount_pct}% | Margin: {gross_margin_pct:.1f}%."
        })

        exceed_authority = False
        approval_reasons = []

        if applied_discount_pct > 5.0:
            exceed_authority = True
            approval_reasons.append(f"Chiết khấu {applied_discount_pct}% vượt hạn mức NVKD (tối đa 5.0%).")
        if final_total_with_vat > 200_000_000:
            exceed_authority = True
            approval_reasons.append(f"Giá trị đơn hàng ({final_total_with_vat:,.0f} đ) vượt ngưỡng tự duyệt của NVKD (200 triệu đ).")
        if lead_time_risk:
            exceed_authority = True
            approval_reasons.append(f"Không đáp ứng thời hạn gốc ({deadline} ngày < lead-time {max_lead_time_days} ngày).")
        if gross_margin_pct < 10.0:
            exceed_authority = True
            approval_reasons.append(f"Biên lợi nhuận gộp ({gross_margin_pct:.1f}%) dưới ngưỡng chuẩn 10%.")
        if "công nợ" in payment_terms.lower() or "trả chậm" in payment_terms.lower():
            exceed_authority = True
            approval_reasons.append(f"Điều khoản thanh toán đặc biệt ({payment_terms}) cần cấp quản lý duyệt.")
        if budget_status == "VUOT_NGAN_SACH":
            approval_reasons.append(f"Tổng báo giá vượt ngân sách khách công bố ({abs(budget_variance):,.0f} đ).")

        approver_role = "NVKD Tự quyết"
        if exceed_authority:
            if applied_discount_pct > 12.0 or final_total_with_vat > 800_000_000 or gross_margin_pct < 8.0:
                approver_role = "Ban Giám đốc (BGĐ)"
            else:
                approver_role = "Trưởng phòng Kinh doanh (TPKD)"

        if lead_time_risk or not overall_stock_ok:
            phan_loai = "CHUA_DAP_UNG"
            nhan_phan_loai = "Chưa đáp ứng"
            du_dieu_kien_tinh_tong = False
            final_crm_status = "CHUA_DAP_UNG"
            can_cu = f"Kho/Tiến độ chưa đáp ứng yêu cầu gốc: {'; '.join(stock_alerts or [lead_time_note])}"
            buoc_tiep_theo = f"Đề xuất giao hàng phân kỳ ({max_lead_time_days} ngày) và trình {approver_role} xem xét ngoại lệ."
        elif exceed_authority:
            phan_loai = "CHO_DUYET"
            nhan_phan_loai = "Chờ duyệt"
            du_dieu_kien_tinh_tong = True
            final_crm_status = "CHO_TPKD_DUYET" if "Trưởng phòng" in approver_role else "CHO_BGD_DUYET"
            can_cu = f"Vượt quyền NVKD: {'; '.join(approval_reasons)}"
            buoc_tiep_theo = f"Chuyển phiếu trình duyệt lên {approver_role} trước khi phát hành báo giá chính thức."
        else:
            phan_loai = "CO_THE_TIEP_TUC"
            nhan_phan_loai = "Có thể tiếp tục"
            du_dieu_kien_tinh_tong = True
            final_crm_status = "SAN_SANG_GUI_KH"
            can_cu = f"Đủ dữ kiện, tồn kho sẵn sàng ({max_lead_time_days} ngày <= {deadline} ngày), chiết khấu {applied_discount_pct}% <= 5.0%."
            buoc_tiep_theo = f"NVKD {opp_dossier['nvkd_phu_trach']} rà soát thư chào giá và gửi khách hàng trong SLA {opp_dossier['sla_phan_hoi_gio']}h."

        logs.append({
            "step": 6,
            "skill": "Skill_3_PhanHoi_KhachHang",
            "tool": "Tool: Evaluate_Approval_Rules",
            "title": "Kiểm soát Thẩm quyền & Phân loại Task 2",
            "status": "WARNING" if (exceed_authority or phan_loai == "CHUA_DAP_UNG") else "SUCCESS",
            "desc": f"Phân loại: [{nhan_phan_loai}] | Cấp duyệt: [{approver_role}]."
        })

        approval_ticket = {
            "ma_phieu": f"APR-{opp_id[-4:]}",
            "can_phe_duyet": exceed_authority or (phan_loai == "CHUA_DAP_UNG"),
            "cap_phe_duyet": approver_role,
            "ly_do": approval_reasons,
            "nguoi_trinh": opp_dossier["nvkd_phu_trach"],
            "ngay_trinh": now_str,
            "trang_thai": "CHỜ_DUYỆT" if (exceed_authority or phan_loai == "CHUA_DAP_UNG") else "ĐÃ_TỰ_DUYỆT"
        }

        draft_response_email = self._generate_proposal_email(
            raw_input, opp_dossier, quotation_summary, inventory_checks, 
            lead_time_risk, max_lead_time_days, exceed_authority, approver_role
        )

        logs.append({
            "step": 7,
            "skill": "Skill_3_PhanHoi_KhachHang",
            "tool": "Tool: Save_CRM_And_AuditLog",
            "title": "Kết thúc: Cập nhật & Theo dõi",
            "status": "SUCCESS",
            "desc": f"Cập nhật CRM [{final_crm_status}], lưu truy vết [{trace_id}]."
        })

        result = {
            "trace_id": trace_id,
            "quy_trinh_hoan_tat": True,
            "ket_qua_kiem_tra": "HOP_LE",
            "phan_loai_co_hoi": phan_loai,
            "nhan_phan_loai": nhan_phan_loai,
            "du_dieu_kien_tinh_tong": du_dieu_kien_tinh_tong,
            "xac_nhan_ban_hang_tu_dong": False,
            "can_cu": can_cu,
            "buoc_xu_ly_tiep_theo": buoc_tiep_theo,
            "logs": logs,
            "buoc_hien_tai": "KetThuc_CapNhat_TheoDoi",
            "trang_thai_crm": final_crm_status,
            "ho_so_co_hoi": opp_dossier,
            "tham_dinh_ton_kho": {
                "tinh_trang_kho": "Đủ hàng sẵn" if overall_stock_ok else "Thiếu hàng / Cần nhập",
                "thoi_gian_cung_ung_toi_thieu_ngay": max_lead_time_days,
                "rui_ro_tien_do": lead_time_risk,
                "ghi_chu_tien_do": lead_time_note,
                "canh_bao_ton": stock_alerts,
                "chi_tiet": inventory_checks
            },
            "bang_bao_gia": quotation_summary,
            "trinh_duyet": approval_ticket,
            "ban_nhap_phan_hoi": draft_response_email
        }

        self._write_audit_log(trace_id, opp_id, logs, {
            "phan_loai": phan_loai,
            "tong_sau_vat": final_total_with_vat,
            "cap_duyet": approver_role
        })
        return result

    def summarize_batch(self, analyzed_items: List[Dict[str, Any]]) -> Dict[str, Any]:
        counts = {"CO_THE_TIEP_TUC": 0, "CAN_BO_SUNG_THONG_TIN": 0, "CHUA_DAP_UNG": 0, "CHO_DUYET": 0}
        eligible_total_value = 0.0
        continue_value = 0.0
        pending_approval_value = 0.0
        rows_summary = []

        for idx, it in enumerate(analyzed_items):
            p = it.get("pipeline", it)
            parsed = it.get("parsed", {})
            cat = p.get("phan_loai_co_hoi", "CAN_BO_SUNG_THONG_TIN")
            if cat in counts:
                counts[cat] += 1

            q = p.get("bang_bao_gia") or {}
            val = float(q.get("tong_sau_vat", 0) or 0)
            if p.get("du_dieu_kien_tinh_tong"):
                eligible_total_value += val
                if cat == "CO_THE_TIEP_TUC":
                    continue_value += val
                elif cat == "CHO_DUYET":
                    pending_approval_value += val

            d = p.get("ho_so_co_hoi") or {}
            rows_summary.append({
                "index": idx,
                "id_yeu_cau": parsed.get("id_yeu_cau") or f"REQ-{idx+1:03d}",
                "khach_hang": parsed.get("khach_hang") or d.get("khach_hang") or "Khách hàng",
                "ma_co_hoi": d.get("ma_co_hoi", "-"),
                "nvkd": d.get("nvkd_phu_trach", "Chờ làm rõ"),
                "phan_loai": cat,
                "nhan_phan_loai": p.get("nhan_phan_loai", ""),
                "gia_tri_sau_vat": val,
                "du_dieu_kien_tinh_tong": p.get("du_dieu_kien_tinh_tong", False),
                "can_cu": p.get("can_cu", ""),
                "buoc_tiep_theo": p.get("buoc_xu_ly_tiep_theo", "")
            })

        return {
            "tong_so_yeu_cau": len(analyzed_items),
            "so_co_the_tiep_tuc": counts["CO_THE_TIEP_TUC"],
            "so_can_bo_sung_thong_tin": counts["CAN_BO_SUNG_THONG_TIN"],
            "so_chua_dap_ung": counts["CHUA_DAP_UNG"],
            "so_cho_duyet": counts["CHO_DUYET"],
            "tong_gia_tri_tiep_tuc": continue_value,
            "tong_gia_tri_cho_duyet": pending_approval_value,
            "tong_gia_tri_du_dieu_kien": eligible_total_value,
            "chi_tiet": rows_summary
        }

    def _assign_sales_rep(self, primary_category: str) -> Dict[str, Any]:
        if not self.sales_reps:
            return {"ma_nv": "NVKD01", "ho_ten": "Nguyễn Văn An", "chuyen_mon": "Máy chủ & Lưu trữ", "tai_hien_tai": 28, "email": "an.nv@minhan.com.vn"}
        p_cat_clean = str(primary_category).lower()
        matched = [r for r in self.sales_reps if p_cat_clean in str(r.get('chuyen_mon', '')).lower() or str(r.get('chuyen_mon', '')).lower() in p_cat_clean or any(k in str(r.get('chuyen_mon', '')).lower() for k in p_cat_clean.split())]
        if matched:
            return sorted(matched, key=lambda x: x.get('tai_hien_tai', 0))[0]
        return sorted(self.sales_reps, key=lambda x: x.get('tai_hien_tai', 0))[0]

    def _generate_clarification_email(self, raw_input: Dict, missing_fields: List[str]) -> str:
        kh = raw_input.get("khach_hang", "Quý Khách hàng")
        contact = raw_input.get("nguoi_lien_he", "Quý Anh/Chị")
        email_kh = raw_input.get("email_kh", "khachhang@domain.com")

        bullets = ""
        for i, f in enumerate(missing_fields, 1):
            if "Ngân sách" in f:
                bullets += f"  {i}. **Ngân sách dự kiến**: Khoảng ngân sách Quý công ty dự định phân bổ để Minh An tư vấn cấu hình phù hợp nhất.\n"
            elif "Thời hạn" in f:
                bullets += f"  {i}. **Thời hạn triển khai mong muốn**: Mốc thời gian cần bàn giao hoàn tất để Vận hành chuẩn bị kho và xếp lịch kỹ sư.\n"
            elif "Danh mục" in f or "Sản phẩm" in f:
                bullets += f"  {i}. **Danh mục & Số lượng cụ thể**: Chi tiết model, số lượng từng thiết bị hoặc bản quyền.\n"
            else:
                bullets += f"  {i}. **{f}**: Thông tin chi tiết để hoàn tất hồ sơ cơ hội.\n"

        if not raw_input.get("nguoi_phe_duyet_kh"):
            bullets += f"  {len(missing_fields)+1}. **Đầu mối phê duyệt phía Quý công ty**: Thông tin người phụ trách quyết định ngân sách để thuận tiện phối hợp.\n"

        return f"""Tiêu đề: [Minh An B2B] Phản hồi yêu cầu báo giá & Đề nghị xác nhận thông tin - {kh}
Gửi đến: {email_kh} (Kính gửi: {contact or 'Bộ phận Mua sắm'})

Kính gửi Quý khách hàng {kh},

Công ty Cổ phần Công nghệ Minh An xin trân trọng cảm ơn Quý công ty đã gửi yêu cầu hợp tác về thiết bị và giải pháp công nghệ.

Để xây dựng phương án kỹ thuật tối ưu và áp dụng chính sách giá tốt nhất, Minh An xin phép nhờ Quý công ty xác nhận thêm một số thông tin còn thiếu sau:

{bullets}
Ngay khi nhận được phản hồi, Minh An cam kết gửi Bảng Báo giá chính thức trong vòng 04 giờ làm việc.

Trân trọng,
Bộ phận Tiếp nhận B2B - Công ty CP Công nghệ Minh An | Hotline: (024) 3888 9999"""

    def _generate_proposal_email(self, raw_input: Dict, dossier: Dict, quote: Dict, 
                                 inv_checks: List[Dict], lead_time_risk: bool, 
                                 max_lead_days: int, exceed_authority: bool, 
                                 approver_role: str) -> str:
        kh = dossier.get("khach_hang", "Quý Khách hàng")
        mst = f" (MST: {dossier['ma_so_thue']})" if dossier.get("ma_so_thue") else ""
        contact = dossier.get("nguoi_lien_he", "Quý Anh/Chị")
        email_kh = dossier.get("email_kh", "khachhang@domain.com")
        rep_name = dossier.get("nvkd_phu_trach")
        rep_email = dossier.get("email_nvkd")

        table_rows = ""
        for i, item in enumerate(quote['chi_tiet_hang_hoa'], 1):
            table_rows += f"| {i} | {item['ten_san_pham']} | {item['so_luong']} | {item['don_gia_niem_yet']:,.0f} đ | {item['thanh_tien_niem_yet']:,.0f} đ |\n"

        if lead_time_risk:
            lead_commitment = f"- **Tiến độ cung ứng (Đề xuất điều chỉnh)**: Do có hạng mục cần nhập kho ({max_lead_days} ngày làm việc), Minh An đề xuất giao 2 đợt: Đợt 1 giao hàng sẵn kho trong 24-48h, Đợt 2 hoàn tất trong {max_lead_days} ngày."
        else:
            lead_commitment = f"- **Tiến độ cung ứng**: Toàn bộ hàng hóa sẵn kho, bàn giao tại {dossier.get('dia_chi_giao_hang') or 'văn phòng Quý công ty'} trong vòng {max_lead_days} ngày làm việc."

        tech_note = f"\n- **Yêu cầu kỹ thuật kèm theo**: {dossier['yeu_cau_ky_thuat']}" if dossier.get("yeu_cau_ky_thuat") else ""
        approval_notice = f"\n*(Lưu ý nội bộ: Hồ sơ đang trình {approver_role} phê duyệt chính thức trước khi phát hành.)*\n" if exceed_authority else ""

        return f"""Tiêu đề: [Minh An B2B] Thư chào giá & Dự thảo Đề xuất - {kh}{mst} (Mã cơ hội: {dossier['ma_co_hoi']})
Gửi đến: {email_kh} (Kính gửi: {contact})
Người phụ trách: {rep_name} ({rep_email})

Kính gửi Quý khách hàng {kh}{mst},

Công ty Cổ phần Công nghệ Minh An trân trọng gửi tới Quý công ty Bảng Dự thảo Đề xuất Báo giá chi tiết:

### I. DANH MỤC THIẾT BỊ VÀ DỊCH VỤ ĐỀ XUẤT
| STT | Tên Hàng hóa / Dịch vụ | SL | Đơn giá niêm yết | Thành tiền niêm yết |
| :-- | :--------------------- | :- | :--------------- | :------------------ |
{table_rows}
### II. TỔNG HỢP GIÁ TRỊ & ĐIỀU KHOẢN THƯƠNG MẠI
- **Tổng giá trị niêm yết**: {quote['tong_gia_niem_yet']:,.0f} VNĐ
- **Chiết khấu thương mại ({quote['chiet_khau_pct']:.1f}%)**: -{quote['tien_chiet_khau']:,.0f} VNĐ
- **Tổng giá trị trước thuế**: {quote['tong_truoc_vat']:,.0f} VNĐ
- **Thuế GTGT (VAT 10%)**: {quote['thue_vat_10pct']:,.0f} VNĐ
- **TỔNG GIÁ TRỊ THANH TOÁN**: **{quote['tong_sau_vat']:,.0f} VNĐ**
- **Hình thức thanh toán**: {dossier.get('hinh_thuc_thanh_toan', 'Theo hợp đồng')}

### III. CAM KẾT TRIỂN KHAI VÀ HẬU MÃI
{lead_commitment}{tech_note}
- **Chính sách bảo hành**: Bảo hành chính hãng tận nơi theo tiêu chuẩn Hãng.{approval_notice}
Trân trọng,
{rep_name} - Chuyên viên Tư vấn Giải pháp Doanh nghiệp | Công ty CP Công nghệ Minh An"""
