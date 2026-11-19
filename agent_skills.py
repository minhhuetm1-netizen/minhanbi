"""
agent_skills.py
Các kỹ năng nghiệp vụ (Skills) của AI Agent Minh An:
- Skill 1: AI Extraction & Data Cleaning (Hỗ trợ Gemini/OpenAI API hoặc Rule-based NLP)
- Skill 2: Business Logic & Calculation Engine (100% Python/Pandas chính xác, không để LLM tự cộng trừ)
- Skill 3: Document Drafting & Tone Rewriting (Hồ sơ cơ hội, Email phản hồi KH, Phiếu chuyển duyệt)
"""

import os
import re
import json
import uuid
import urllib.request
import datetime
from typing import Dict, List, Any, Optional
import pandas as pd
from data_processor import (
    is_missing, normalize_currency, normalize_deadline_days, normalize_percentage
)

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gemini_config.json")

def get_default_ai_config() -> Dict[str, str]:
    """Đọc cấu hình API Key & Model mặc định từ biến môi trường hoặc gemini_config.json."""
    env_key = os.environ.get("GEMINI_API_KEY", "").strip()
    env_model = os.environ.get("GEMINI_MODEL", "gemini-3-flash-preview").strip()
    if env_key:
        return {"provider": "Gemini", "gemini_api_key": env_key, "gemini_model": env_model}
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {"provider": "Rule-based + AI", "gemini_api_key": "", "gemini_model": "gemini-3-flash-preview"}

SALES_REPS = [
    {"ma_nv": "NVKD01", "ho_ten": "Nguyễn Văn An", "chuyen_mon": "Máy chủ & Lưu trữ", "tai": 28, "email": "an.nv@minhan.com.vn"},
    {"ma_nv": "NVKD02", "ho_ten": "Trần Thị Bích", "chuyen_mon": "Thiết bị đầu cuối", "tai": 30, "email": "bich.tt@minhan.com.vn"},
    {"ma_nv": "NVKD03", "ho_ten": "Lê Hoàng Cường", "chuyen_mon": "Hạ tầng Mạng", "tai": 29, "email": "cuong.lh@minhan.com.vn"},
    {"ma_nv": "NVKD04", "ho_ten": "Phạm Minh Dũng", "chuyen_mon": "Phần mềm & Bản quyền", "tai": 31, "email": "dung.pm@minhan.com.vn"},
    {"ma_nv": "NVKD05", "ho_ten": "Hoàng Thu Hà", "chuyen_mon": "Dịch vụ Kỹ thuật", "tai": 32, "email": "ha.ht@minhan.com.vn"},
    {"ma_nv": "NVKD06", "ho_ten": "Đỗ Quốc Hùng", "chuyen_mon": "An toàn thông tin", "tai": 27, "email": "hung.dq@minhan.com.vn"}
]

def _call_llm_json(prompt: str, api_key: str = "", provider: str = "Gemini") -> Optional[str]:
    """Gọi LLM (Gemini hoặc OpenAI) bằng thư viện chuẩn urllib; tự nạp key từ gemini_config.json nếu trống."""
    cfg = get_default_ai_config()
    eff_key = (api_key or "").strip() or (cfg.get("gemini_api_key", "") if provider == "Gemini" else "")
    if not eff_key:
        return None
    try:
        if provider == "Gemini":
            primary_model = cfg.get("gemini_model", "gemini-3-flash-preview")
            models_to_try = [primary_model, "gemini-3.1-flash-lite-preview", "gemini-flash-latest"]
            seen = set()
            for model in models_to_try:
                if model in seen:
                    continue
                seen.add(model)
                try:
                    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={eff_key}"
                    payload = {
                        "contents": [{"parts": [{"text": prompt}]}],
                        "generationConfig": {"thinkingConfig": {"thinkingBudget": 0}}
                    }
                    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"})
                    with urllib.request.urlopen(req, timeout=20) as resp:
                        data = json.loads(resp.read().decode("utf-8"))
                        return data["candidates"][0]["content"]["parts"][0]["text"]
                except Exception:
                    continue
        elif provider == "OpenAI":
            url = "https://api.openai.com/v1/chat/completions"
            payload = {
                "model": "gpt-4o-mini",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.2
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json", "Authorization": f"Bearer {eff_key}"}
            )
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                return data["choices"][0]["message"]["content"]
    except Exception:
        return None
    return None

def _extract_qty_near(cleaned_text: str, start: int, end: int) -> int:
    preceding = cleaned_text[max(0, start - 40):start]
    following = cleaned_text[end:min(len(cleaned_text), end + 20)]
    qty = 1
    m_pre = re.findall(r'(\d+)\s*(?:chiếc|bộ|thiết bị|máy|gói|license|user|con)?\s*$', preceding.strip())
    if m_pre:
        qty = int(m_pre[-1])
    else:
        m_post = re.search(r'^\s*[:=x\-]?\s*(\d+)', following)
        if m_post:
            qty = int(m_post.group(1))
    return max(1, qty)

# ==============================================================================
# SKILL 1: AI EXTRACTION & DATA CLEANING
# ==============================================================================
def skill1_extract_and_clean(
    inquiries_df: pd.DataFrame,
    catalog_df: pd.DataFrame,
    api_key: str = "",
    provider: str = "Rule-based + AI",
    extra_prompt: str = "",
    audit_log: Optional[List[Dict]] = None
) -> pd.DataFrame:
    """
    Bóc tách từ văn bản tự do + chuẩn hóa số, tiền tệ, ngày tháng, rỗng (null/NaN).
    Trả về DataFrame sạch để người dùng kiểm duyệt tại Bước 2.
    """
    catalog_records = catalog_df.to_dict("records") if catalog_df is not None and not catalog_df.empty else []
    cleaned_rows = []

    if audit_log is not None:
        audit_log.append({
            "timestamp": datetime.datetime.now().strftime("%H:%M:%S"),
            "step": "Bước 2 - Skill 1",
            "tool": f"skill1_extract_and_clean ({provider})",
            "detail": f"Bóc tách & làm sạch {len(inquiries_df)} yêu cầu. Ghi chú bổ sung: '{extra_prompt or 'Không'}'"
        })

    product_patterns = [
        ("LAP-DELL-5540", r'(dell\s*latitude\s*5540|latitude\s*5540|laptop\s*dell)'),
        ("LAP-LEN-T14", r'(lenovo\s*thinkpad\s*t14|thinkpad\s*t14|laptop\s*thinkpad|laptop\s*lenovo)'),
        ("SRV-DELL-R750", r'(dell\s*poweredge\s*r750|poweredge\s*r750|máy chủ dell|server dell)'),
        ("SW-CISCO-C9200", r'(cisco\s*catalyst\s*c9200|c9200l|switch\s*cisco|switch\s*catalyst)'),
        ("FW-FORTI-60F", r'(fortigate\s*60f|fortinet\s*60f|tường lửa fortinet|firewall fortinet)'),
        ("SVC-SEC-1Y", r'(dịch vụ triển khai[^\.]{0,25}an toàn mạng|triển khai an toàn mạng|cấu hình an toàn mạng|triển khai bảo mật|cài đặt mạng)'),
        ("SVC-MAINT-247", r'(dịch vụ bảo trì & hỗ trợ kỹ thuật|bảo trì hệ thống|bảo trì it)'),
        ("LIC-M365-BUS", r'(microsoft\s*365|m365\s*business|bản quyền m365|license office 365|license microsoft)')
    ]

    # Giới hạn tối đa 5 lượt gọi LLM mẫu để tránh tràn Context Window hoặc Rate-limit khi file có hàng nghìn dòng
    llm_calls_count = 0
    max_sample_llm_calls = 5

    for idx, row in inquiries_df.iterrows():
        r = row.to_dict()
        req_id = str(r.get("id_yeu_cau")) if not is_missing(r.get("id_yeu_cau")) else f"REQ-{idx+1:03d}"
        raw_text = str(r.get("noi_dung_yeu_cau") or "")
        text_lower = raw_text.lower()

        kh = str(r.get("khach_hang") or f"Khách hàng #{idx+1}").strip()
        if is_missing(kh):
            kh = f"Khách hàng #{idx+1}"

        # Chuẩn hóa ngân sách, thời hạn, chiết khấu từ cột hoặc trích từ văn bản
        budget = normalize_currency(r.get("ngan_sach_cong_bo"))
        if budget is None:
            bm = re.search(r'ngân sách[^\d]*(\d+(?:[\.,]\d+)*)\s*(tỷ|ty|triệu|trieu|tr|m|vnđ|vnd|đ)?', text_lower)
            if bm:
                budget = normalize_currency(f"{bm.group(1)} {bm.group(2) or ''}")

        deadline = normalize_deadline_days(r.get("thoi_han_ngay"), r.get("ngay_nhan"))
        if deadline is None:
            dm = re.search(r'(trong vòng|thời hạn|tiến độ|giao trong|trong)\s*(\d+\s*(?:ngày|tuần|tháng))', text_lower)
            if dm:
                deadline = normalize_deadline_days(dm.group(2), r.get("ngay_nhan"))

        discount = normalize_percentage(r.get("chiet_khau_de_xuat_pct"))
        if discount is None:
            dcm = re.search(r'chiết khấu[^\d]*(\d+(?:[\.,]\d+)?)\s*%', text_lower)
            if dcm:
                discount = normalize_percentage(dcm.group(1))

        # Bóc tách sản phẩm & số lượng (ưu tiên cột tách riêng nếu có, sau đó quét văn bản)
        items = []
        matched_skus = set()

        if not is_missing(r.get("ma_sp")):
            raw_sku_str = str(r.get("ma_sp"))
            for part in re.split(r'[;\n]+', raw_sku_str):
                part = part.strip()
                if not part:
                    continue
                qty = 1
                m_part = re.match(r'^([A-Za-z0-9\-_]+)\s*[:xX\*]\s*(\d+)$', part)
                if m_part:
                    tok, qty = m_part.group(1), int(m_part.group(2))
                else:
                    tok = part
                    if not is_missing(r.get("so_luong")):
                        try:
                            qty = max(1, int(float(r.get("so_luong"))))
                        except Exception:
                            qty = 1
                cat_match = next((c for c in catalog_records if str(c.get("ma_sp", "")).lower() == tok.lower() or tok.lower() in str(c.get("ten_san_pham", "")).lower()), None)
                if cat_match and cat_match["ma_sp"] not in matched_skus:
                    matched_skus.add(cat_match["ma_sp"])
                    items.append({"ma_sp": cat_match["ma_sp"], "ten_san_pham": cat_match["ten_san_pham"], "so_luong": qty})

        cleaned_text = re.sub(r'\b24/7\b', '', text_lower)
        cleaned_text = re.sub(r'\b\d+\s*(tháng|năm)\b', '', cleaned_text)

        for sku, pat in product_patterns:
            if sku in matched_skus:
                continue
            m = re.search(pat, cleaned_text)
            if m:
                qty = _extract_qty_near(cleaned_text, m.start(), m.end())
                cat_match = next((c for c in catalog_records if str(c.get("ma_sp", "")) == sku), None)
                if cat_match:
                    matched_skus.add(sku)
                    items.append({"ma_sp": sku, "ten_san_pham": cat_match["ten_san_pham"], "so_luong": qty})

        for cat_match in catalog_records:
            sku = str(cat_match.get("ma_sp", ""))
            if not sku or sku in matched_skus:
                continue
            m = re.search(re.escape(sku.lower()), cleaned_text)
            if m:
                qty = _extract_qty_near(cleaned_text, m.start(), m.end())
                matched_skus.add(sku)
                items.append({"ma_sp": sku, "ten_san_pham": cat_match["ten_san_pham"], "so_luong": qty})

        # Gọi LLM mẫu khi ở chế độ Gemini/OpenAI (chỉ gọi tối đa max_sample_llm_calls để bảo vệ RAM/Context)
        has_eff_key = bool((api_key or "").strip()) or (provider == "Gemini" and bool(get_default_ai_config().get("gemini_api_key")))
        needs_llm = (not items or budget is None or deadline is None or discount is None or bool((extra_prompt or "").strip()))
        if has_eff_key and provider in ("Gemini", "OpenAI") and raw_text.strip() and needs_llm and llm_calls_count < max_sample_llm_calls:
            llm_calls_count += 1
            llm_prompt = (
                f"Trích xuất JSON (chỉ JSON thuần): ngan_sach_vnd (số hoặc null), thoi_han_ngay (số nguyên hoặc null), "
                f"chiet_khau_pct (số hoặc null), nguoi_duyet_kh (chuỗi hoặc rỗng). "
                f"Ghi chú thêm: {extra_prompt}. Văn bản: {raw_text}"
            )
            llm_out = _call_llm_json(llm_prompt, api_key, provider)
            if llm_out:
                try:
                    j_str = re.search(r'\{.*\}', llm_out, re.S).group(0)
                    parsed_llm = json.loads(j_str)
                    if budget is None and parsed_llm.get("ngan_sach_vnd"):
                        budget = float(parsed_llm["ngan_sach_vnd"])
                    if deadline is None and parsed_llm.get("thoi_han_ngay"):
                        deadline = int(parsed_llm["thoi_han_ngay"])
                    if discount is None and parsed_llm.get("chiet_khau_pct") is not None:
                        discount = float(parsed_llm["chiet_khau_pct"])
                    if audit_log is not None:
                        audit_log.append({
                            "timestamp": datetime.datetime.now().strftime("%H:%M:%S"),
                            "step": "Bước 2 - LLM",
                            "tool": f"{provider} API ({get_default_ai_config().get('gemini_model', 'gemini-3-flash-preview')})",
                            "detail": f"Đã gọi Gemini phân tích bổ sung cho {req_id}"
                        })
                except Exception:
                    pass

        # Kiểm tra thông tin thiếu
        missing_list = []
        if not items:
            missing_list.append("Danh mục sản phẩm / Số lượng")
        if budget is None or budget <= 0:
            missing_list.append("Ngân sách dự kiến")
        if deadline is None or deadline <= 0:
            missing_list.append("Thời hạn triển khai")

        items_formatted = "; ".join([f"{it['ma_sp']}:{it['so_luong']}" for it in items])
        nguoi_duyet = "" if is_missing(r.get("nguoi_phe_duyet_kh")) else str(r.get("nguoi_phe_duyet_kh")).strip()

        cleaned_rows.append({
            "id_yeu_cau": req_id,
            "khach_hang": kh,
            "nguoi_lien_he": "" if is_missing(r.get("nguoi_lien_he")) else str(r.get("nguoi_lien_he")).strip(),
            "email_kh": "" if is_missing(r.get("email_kh")) else str(r.get("email_kh")).strip(),
            "nguoi_phe_duyet_kh": nguoi_duyet,
            "hinh_thuc_thanh_toan": "Tiêu chuẩn" if is_missing(r.get("hinh_thuc_thanh_toan")) else str(r.get("hinh_thuc_thanh_toan")).strip(),
            "noi_dung_goc": raw_text,
            "san_pham_va_sl": items_formatted,
            "ngan_sach_vnd": budget,
            "thoi_han_ngay": deadline,
            "chiet_khau_de_nghi_pct": discount if discount is not None else 0.0,
            "thieu_thong_tin": len(missing_list) > 0,
            "chi_tiet_thieu": ", ".join(missing_list) if missing_list else "Đầy đủ",
            "trang_thai_lam_sach": "Cần bổ sung thông tin" if missing_list else "Đủ dữ kiện"
        })

    return pd.DataFrame(cleaned_rows)

# ==============================================================================
# SKILL 2: BUSINESS LOGIC & CALCULATION ENGINE (PURE PYTHON/PANDAS)
# ==============================================================================
def skill2_calculate_and_classify(
    cleaned_df: pd.DataFrame,
    catalog_df: pd.DataFrame,
    inventory_df: pd.DataFrame,
    rules_config: Dict[str, Any],
    audit_log: Optional[List[Dict]] = None
) -> List[Dict[str, Any]]:
    """
    Tính toán giá, đối chiếu tồn kho, kiểm tra Lead-time & Phân quyền duyệt bằng Python/Pandas.
    Phân loại chính xác vào 4 nhóm:
      1. Có thể tiếp tục
      2. Cần bổ sung thông tin
      3. Chưa đáp ứng
      4. Chờ duyệt
    """
    cat_map = {str(r["ma_sp"]).strip(): r for r in catalog_df.to_dict("records")} if not catalog_df.empty else {}
    inv_map = {str(r["ma_sp"]).strip(): r for r in inventory_df.to_dict("records")} if not inventory_df.empty else {}

    nvkd_max_disc = float(rules_config.get("nvkd_max_discount", 5.0))
    tpkd_max_disc = float(rules_config.get("tpkd_max_discount", 12.0))
    nvkd_max_val = float(rules_config.get("nvkd_max_value", 200_000_000))
    vat_rate = float(rules_config.get("vat_rate", 0.10))

    if audit_log is not None:
        audit_log.append({
            "timestamp": datetime.datetime.now().strftime("%H:%M:%S"),
            "step": "Bước 3 - Skill 2",
            "tool": "skill2_calculate_and_classify (Python/Pandas Engine)",
            "detail": f"Áp dụng quy tắc: CK NVKD <= {nvkd_max_disc}%, CK TPKD <= {tpkd_max_disc}%, Giá trị NVKD <= {nvkd_max_val:,.0f} đ."
        })

    analyzed_results = []

    for idx, row in cleaned_df.iterrows():
        r = row.to_dict()
        req_id = r.get("id_yeu_cau", f"REQ-{idx+1:03d}")
        kh = r.get("khach_hang", "Khách hàng")
        budget = None if is_missing(r.get("ngan_sach_vnd")) else float(r.get("ngan_sach_vnd"))
        deadline = None if is_missing(r.get("thoi_han_ngay")) else int(float(r.get("thoi_han_ngay")))
        disc_pct = 0.0 if is_missing(r.get("chiet_khau_de_nghi_pct")) else float(r.get("chiet_khau_de_xuat_pct") if "chiet_khau_de_xuat_pct" in r else r.get("chiet_khau_de_nghi_pct"))
        payment_str = str(r.get("hinh_thuc_thanh_toan") or "Tiêu chuẩn")

        # Parse lại danh sách SP từ ô đã làm sạch (hỗ trợ người dùng vừa sửa tay ở Bước 2!)
        raw_items_str = str(r.get("san_pham_va_sl") or "")
        parsed_items = []
        for token in [t.strip() for t in raw_items_str.split(";") if t.strip()]:
            if ":" in token:
                sku, q_str = token.split(":", 1)
                sku = sku.strip()
                try:
                    qty = max(1, int(float(q_str.strip())))
                except Exception:
                    qty = 1
            else:
                sku, qty = token, 1
            if sku in cat_map:
                parsed_items.append({"ma_sp": sku, "so_luong": qty, "info": cat_map[sku]})

        # Kiểm tra thiếu thông tin sau khi người dùng có thể đã chỉnh tay ở Bước 2
        missing_fields = []
        if not parsed_items:
            missing_fields.append("Danh mục sản phẩm / Số lượng")
        if budget is None or budget <= 0:
            missing_fields.append("Ngân sách dự kiến")
        if deadline is None or deadline <= 0:
            missing_fields.append("Thời hạn triển khai")

        if missing_fields:
            analyzed_results.append({
                "id_yeu_cau": req_id,
                "ma_co_hoi": "-",
                "khach_hang": kh,
                "nguoi_lien_he": r.get("nguoi_lien_he", ""),
                "email_kh": r.get("email_kh", ""),
                "nvkd_phu_trach": "Chờ phân bổ (Thiếu TT)",
                "phan_loai": "Cần bổ sung thông tin",
                "du_dieu_kien_tinh_tong": False,
                "tu_dong_xac_nhan": False,
                "can_phieu_chuyen_duyet": False,
                "cap_phe_duyet": "Chưa xác định",
                "gia_tri_niem_yet": 0.0,
                "chiet_khau_pct": disc_pct,
                "gia_tri_sau_vat": 0.0,
                "loi_nhuan_gop": 0.0,
                "margin_pct": 0.0,
                "ngan_sach_kh": budget or 0.0,
                "can_cu": f"Thiếu dữ kiện bắt buộc: {', '.join(missing_fields)}.",
                "buoc_tiep_theo": "Gửi email đề nghị khách hàng bổ sung thông tin trước khi lập báo giá.",
                "chi_tiet_hang_hoa": [],
                "ly_do_duyet": [],
                "noi_dung_goc": r.get("noi_dung_goc", "")
            })
            continue

        # Đủ dữ kiện -> Tạo mã cơ hội & Phân bổ NVKD
        opp_id = f"OPP-{datetime.datetime.now().strftime('%Y%m')}-{uuid.uuid4().hex[:4].upper()}"
        primary_cat = str(parsed_items[0]["info"].get("danh_muc", "")).lower()
        rep = next((s for s in SALES_REPS if s["chuyen_mon"].lower() in primary_cat or primary_cat in s["chuyen_mon"].lower()), SALES_REPS[idx % len(SALES_REPS)])

        # Tính toán giá & kiểm tra tồn kho
        list_total = 0.0
        cogs_total = 0.0
        max_lead_days = 0
        stock_ok = True
        stock_issues = []
        line_details = []

        for it in parsed_items:
            sku = it["ma_sp"]
            qty = it["so_luong"]
            p_info = it["info"]
            inv = inv_map.get(sku, {"ton_kha_dung": 0, "thoi_gian_giao_san_co_ngay": 2, "thoi_gian_nhap_hang_ngay": 14})

            unit_price = float(p_info.get("gia_niem_yet", 0))
            cogs = float(p_info.get("gia_von", unit_price * 0.8))
            sub_list = unit_price * qty
            sub_cogs = cogs * qty
            list_total += sub_list
            cogs_total += sub_cogs

            ton = int(float(inv.get("ton_kha_dung", 0)))
            du_hang = ton >= qty
            lt = int(float(inv.get("thoi_gian_giao_san_co_ngay", 1))) if du_hang else int(float(inv.get("thoi_gian_nhap_hang_ngay", 14)))
            if lt > max_lead_days:
                max_lead_days = lt
            if not du_hang:
                stock_ok = False
                stock_issues.append(f"{sku} thiếu {qty - ton} (tồn {ton}/{qty}, nhập {lt} ngày)")

            line_details.append({
                "ma_sp": sku,
                "ten_san_pham": p_info.get("ten_san_pham", sku),
                "so_luong": qty,
                "ton_kha_dung": ton,
                "du_hang": du_hang,
                "lead_time_ngay": lt,
                "don_gia": unit_price,
                "thanh_tien": sub_list
            })

        disc_amt = list_total * (disc_pct / 100.0)
        net_total = list_total - disc_amt
        vat_amt = net_total * vat_rate
        grand_total = net_total + vat_amt
        gross_profit = net_total - cogs_total
        margin_pct = round((gross_profit / net_total * 100.0), 1) if net_total > 0 else 0.0
        lead_conflict = max_lead_days > deadline

        # Kiểm tra thẩm quyền & phân loại 4 nhóm
        exceed_reasons = []
        if disc_pct > nvkd_max_disc:
            exceed_reasons.append(f"Chiết khấu {disc_pct}% > hạn mức NVKD ({nvkd_max_disc}%)")
        if grand_total > nvkd_max_val:
            exceed_reasons.append(f"Giá trị {grand_total:,.0f} đ > ngưỡng tự duyệt NVKD ({nvkd_max_val:,.0f} đ)")
        if "công nợ" in payment_str.lower() or "trả chậm" in payment_str.lower():
            exceed_reasons.append(f"Thanh toán đặc biệt ({payment_str})")
        if margin_pct < 10.0:
            exceed_reasons.append(f"Biên lãi gộp mỏng ({margin_pct}% < 10%)")

        if lead_conflict or not stock_ok:
            phan_loai = "Chưa đáp ứng"
            eligible = False
            cap_duyet = "Ban Giám đốc (BGĐ)" if disc_pct > tpkd_max_disc else "Trưởng phòng Kinh doanh (TPKD)"
            can_phieu = True
            if lead_conflict:
                exceed_reasons.append(f"Thời hạn khách ({deadline} ngày) < Lead-time ({max_lead_days} ngày)")
            can_cu = f"Không đáp ứng kho/tiến độ: {'; '.join(stock_issues or [f'Cần {max_lead_days} ngày > {deadline} ngày'])}"
            buoc_tiep = f"Tư vấn giao 2 đợt ({max_lead_days} ngày) hoặc từ chối khéo; không tự xác nhận đơn."
        elif exceed_reasons:
            phan_loai = "Chờ duyệt"
            eligible = True
            cap_duyet = "Ban Giám đốc (BGĐ)" if (disc_pct > tpkd_max_disc or grand_total > 800_000_000) else "Trưởng phòng Kinh doanh (TPKD)"
            can_phieu = True
            can_cu = f"Đáp ứng kho & tiến độ ({max_lead_days}/{deadline} ngày) nhưng vượt quyền: {'; '.join(exceed_reasons)}."
            buoc_tiep = f"Lập phiếu chuyển duyệt gửi {cap_duyet}; chờ duyệt trước khi gửi chính thức."
        else:
            phan_loai = "Có thể tiếp tục"
            eligible = True
            cap_duyet = "NVKD Tự quyết"
            can_phieu = False
            can_cu = f"Đủ tồn kho, đáp ứng tiến độ ({max_lead_days}/{deadline} ngày), CK {disc_pct}% <= {nvkd_max_disc}%."
            buoc_tiep = f"NVKD {rep['ho_ten']} gửi báo giá chuẩn cho khách hàng trong SLA 4h."

        analyzed_results.append({
            "id_yeu_cau": req_id,
            "ma_co_hoi": opp_id,
            "khach_hang": kh,
            "nguoi_lien_he": r.get("nguoi_lien_he", ""),
            "email_kh": r.get("email_kh", ""),
            "nvkd_phu_trach": rep["ho_ten"],
            "email_nvkd": rep["email"],
            "phan_loai": phan_loai,
            "du_dieu_kien_tinh_tong": eligible,
            "tu_dong_xac_nhan": False,  # Quy tắc cứng: Không tự động xác nhận bán hàng
            "can_phieu_chuyen_duyet": can_phieu,
            "cap_phe_duyet": cap_duyet,
            "gia_tri_niem_yet": list_total,
            "chiet_khau_pct": disc_pct,
            "tien_chiet_khau": disc_amt,
            "tong_truoc_vat": net_total,
            "thue_vat": vat_amt,
            "gia_tri_sau_vat": grand_total,
            "tong_gia_von": cogs_total,
            "loi_nhuan_gop": gross_profit,
            "margin_pct": margin_pct,
            "ngan_sach_kh": budget,
            "chenh_lech_ngan_sach": budget - grand_total,
            "thoi_han_kh_ngay": deadline,
            "lead_time_thuc_te_ngay": max_lead_days,
            "can_cu": can_cu,
            "buoc_tiep_theo": buoc_tiep,
            "chi_tiet_hang_hoa": line_details,
            "ly_do_duyet": exceed_reasons,
            "noi_dung_goc": r.get("noi_dung_goc", "")
        })

    return analyzed_results

# ==============================================================================
# SKILL 3: DOCUMENT DRAFTING & TONE REWRITING
# ==============================================================================
def skill3_generate_documents(opp: Dict[str, Any]) -> Dict[str, str]:
    """Sinh 3 tài liệu cho từng cơ hội: Hồ sơ cơ hội, Email phản hồi KH, và Phiếu chuyển duyệt (nếu cần)."""
    kh = opp["khach_hang"]
    cat = opp["phan_loai"]

    # 1. Hồ sơ cơ hội
    dossier_txt = f"""HỒ SƠ CƠ HỘI BÁN HÀNG B2B - CÔNG TY MINH AN
==================================================
Mã yêu cầu       : {opp['id_yeu_cau']}
Mã cơ hội        : {opp['ma_co_hoi']}
Khách hàng       : {kh} ({opp.get('nguoi_lien_he', '')} - {opp.get('email_kh', '')})
NVKD phụ trách   : {opp['nvkd_phu_trach']}
Phân loại trạng thái: [{cat.upper()}]
Đủ ĐK tính tổng GT  : {'CÓ' if opp['du_dieu_kien_tinh_tong'] else 'KHÔNG'}
Tự xác nhận đơn  : KHÔNG (Tuân thủ kiểm duyệt bắt buộc)

KẾT QUẢ ĐỐI CHIẾU & TÍNH TOÁN:
- Giá trị đề xuất (Sau VAT): {opp['gia_tri_sau_vat']:,.0f} VNĐ (Chiết khấu: {opp['chiet_khau_pct']}%)
- Lợi nhuận gộp dự kiến    : {opp['loi_nhuan_gop']:,.0f} VNĐ (Gross Margin: {opp['margin_pct']}%)
- Ngân sách khách công bố  : {opp['ngan_sach_kh']:,.0f} VNĐ
- Căn cứ đối chiếu         : {opp['can_cu']}
- Bước xử lý tiếp theo     : {opp['buoc_tiep_theo']}
"""

    # 2. Bản nháp Email phản hồi KH
    if cat == "Cần bổ sung thông tin":
        email_txt = f"""Tiêu đề: [Minh An B2B] Phản hồi yêu cầu báo giá & Đề nghị bổ sung thông tin - {kh}
Kính gửi: {opp.get('nguoi_lien_he') or 'Quý Công ty'} ({kh}),

Công ty Cổ phần Công nghệ Minh An xin cảm ơn Quý công ty đã gửi yêu cầu hợp tác thiết bị/dịch vụ công nghệ.

Để đội ngũ kỹ thuật và kinh doanh xây dựng báo giá chính xác nhất theo tồn kho thực tế, kính nhờ Quý công ty xác nhận bổ sung thông tin còn thiếu sau:
- {opp['can_cu']}

Ngay khi nhận được phản hồi, Minh An cam kết gửi Báo giá chi tiết trong vòng 04 giờ làm việc.

Trân trọng,
Bộ phận Kinh doanh B2B - Công ty CP Công nghệ Minh An"""
    elif cat == "Chưa đáp ứng":
        email_txt = f"""Tiêu đề: [Minh An B2B] Phương án tiến độ & Đề xuất điều chỉnh đơn hàng - {kh} ({opp['ma_co_hoi']})
Kính gửi: {opp.get('nguoi_lien_he') or 'Quý Công ty'} ({kh}),

Minh An đã kiểm tra tồn kho thực tế cho gói thiết bị Quý công ty yêu cầu (Tổng dự toán: {opp['gia_tri_sau_vat']:,.0f} VNĐ đã gồm VAT).
Tuy nhiên, hiện tại có vướng mắc về tồn kho/tiến độ: {opp['can_cu']}.

Minh An xin trân trọng đề xuất 2 phương án để Quý công ty xem xét:
1. Giao hàng làm 2 đợt: Đợt 1 giao ngay số lượng có sẵn trong 24-48h, Đợt 2 bàn giao đủ sau {opp.get('lead_time_thuc_te_ngay', 14)} ngày.
2. Tư vấn cấu hình tương đương đang sẵn hàng 100% tại kho Hà Nội/TP.HCM.

Rất mong nhận được phản hồi của Quý công ty để chúng tôi hỗ trợ kịp thời!

Trân trọng,
{opp['nvkd_phu_trach']} - Công ty CP Công nghệ Minh An"""
    else:
        rows_md = "\n".join([f"- {d['ten_san_pham']} | SL: {d['so_luong']} | Thành tiền: {d['thanh_tien']:,.0f} đ" for d in opp["chi_tiet_hang_hoa"]])
        note_appr = f"\n*(Lưu ý nội bộ: Báo giá đang trình {opp['cap_phe_duyet']} duyệt trước khi gửi chính thức.)*\n" if opp["can_phieu_chuyen_duyet"] else ""
        email_txt = f"""Tiêu đề: [Minh An B2B] Thư chào giá & Dự thảo Đề xuất - {kh} ({opp['ma_co_hoi']})
Kính gửi: {opp.get('nguoi_lien_he') or 'Quý Công ty'} ({kh}),

Công ty Cổ phần Công nghệ Minh An trân trọng gửi tới Quý công ty Dự thảo Báo giá chi tiết:
{rows_md}

- Tổng giá trị niêm yết: {opp['gia_tri_niem_yet']:,.0f} VNĐ
- Chiết khấu ({opp['chiet_khau_pct']}%): -{opp.get('tien_chiet_khau', 0):,.0f} VNĐ
- TỔNG THANH TOÁN (Đã gồm VAT 10%): {opp['gia_tri_sau_vat']:,.0f} VNĐ
- Thời gian bàn giao cam kết: {opp.get('lead_time_thuc_te_ngay', 3)} ngày làm việc.{note_appr}
Trân trọng,
{opp['nvkd_phu_trach']} - Công ty CP Công nghệ Minh An"""

    # 3. Phiếu chuyển duyệt (Chỉ tạo khi cần duyệt vượt quyền / ngoại lệ)
    approval_txt = ""
    if opp["can_phieu_chuyen_duyet"]:
        reasons_str = "\n".join([f"  + {r}" for r in opp.get("ly_do_duyet", [])])
        approval_txt = f"""PHIẾU TRÌNH DUYỆT NGOẠI LỆ - MINH AN COMPANY
==================================================
Kính gửi       : {opp['cap_phe_duyet']}
Người trình     : {opp['nvkd_phu_trach']}
Mã cơ hội       : {opp['ma_co_hoi']} | Khách hàng: {kh}
Trạng thái      : CHỜ PHÊ DUYỆT (Chưa xác nhận với khách)

CHỈ SỐ TÀI CHÍNH & LÝ DO TRÌNH DUYỆT:
- Tổng giá trị đơn (Sau VAT): {opp['gia_tri_sau_vat']:,.0f} VNĐ
- Chiết khấu đề xuất       : {opp['chiet_khau_pct']}%
- Lợi nhuận gộp            : {opp['loi_nhuan_gop']:,.0f} VNĐ (Margin: {opp['margin_pct']}%)
- Các căn cứ vượt quyền:
{reasons_str}

Đề xuất của NVKD: Kính trình {opp['cap_phe_duyet']} xem xét phê duyệt để kịp gửi khách hàng trong SLA."""

    return {
        "ho_so_co_hoi": dossier_txt,
        "email_nhap": email_txt,
        "phieu_chuyen_duyet": approval_txt
    }

def rewrite_email_tone(current_email: str, style: str, api_key: str = "", provider: str = "Rule-based + AI") -> str:
    """Viết lại văn phong Email nháp theo yêu cầu người dùng ở Bước 4."""
    has_eff_key = bool((api_key or "").strip()) or (provider == "Gemini" and bool(get_default_ai_config().get("gemini_api_key")))
    if has_eff_key and provider in ("Gemini", "OpenAI"):
        prompt = f"Hãy viết lại email B2B tiếng Việt dưới đây theo văn phong '{style}', giữ nguyên toàn bộ số liệu, bảng giá và mã cơ hội:\n\n{current_email}"
        res = _call_llm_json(prompt, api_key, provider)
        if res:
            return res.strip()

    # Fallback thông minh theo văn phong chọn
    if "ngắn gọn" in style.lower():
        lines = [l for l in current_email.splitlines() if l.strip()]
        return "\n".join(lines[:3]) + "\n\n[Bản tóm tắt ngắn gọn]\n" + "\n".join([l for l in lines if any(k in l for k in ('-', 'TỔNG', 'Thời gian', 'Tiêu đề', 'Trân trọng'))])
    elif "thuyết phục" in style.lower():
        return current_email.replace("trân trọng gửi", "rất hân hạnh được đồng hành và trân trọng gửi") + "\n\nP/S: Minh An sẵn sàng hỗ trợ kỹ sư demo/khảo sát tận nơi miễn phí nếu Quý công ty cần triển khai gấp."
    else:
        return current_email.replace("Xin cảm ơn", "Xin trân trọng gửi lời cảm ơn sâu sắc nhất").replace("Chào", "Kính chào")
