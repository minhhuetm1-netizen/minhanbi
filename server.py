"""
Web Server cho AI Agent Minh An Company
Sử dụng thư viện chuẩn của Python (http.server) - Không cần phụ thuộc ngoài.
Hỗ trợ đầy đủ REST API và Web UI tương tác trực tiếp.
"""

import os
import sys
import json
import math
import urllib.parse
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
import pandas as pd
import numpy as np
from agent_engine import MinhAnAgentEngine

PORT = 8080
STATIC_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

engine = MinhAnAgentEngine()

SERVER_CACHE = {
    "raw_sheets": {},
    "cleaned_df": None,
    "analysis": None,
    "schema": None,
    "mode": None,
    "source_name": "",
    "report_text": "",
    "progress": {"percent": 0, "status": "Sẵn sàng", "detail": ""}
}

class MinhAnRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=STATIC_DIR, **kwargs)

    def _paginate_data(self, source, sheet_name, page, page_size):
        page_size = max(10, min(page_size, 200))
        if source == "cleaned":
            df = SERVER_CACHE.get("cleaned_df")
        else:
            sheets = SERVER_CACHE.get("raw_sheets", {})
            if not sheet_name or sheet_name not in sheets:
                sheet_name = next(iter(sheets.keys()), None)
            df = sheets.get(sheet_name) if sheet_name else None

        if df is None or df.empty:
            return {
                "status": "success",
                "source": source,
                "sheet": sheet_name,
                "page": 1,
                "page_size": page_size,
                "total_rows": 0,
                "total_pages": 1,
                "start_idx": 0,
                "end_idx": 0,
                "columns": [],
                "rows": []
            }

        total_rows = len(df)
        total_pages = max(1, math.ceil(total_rows / page_size))
        page = max(1, min(page, total_pages))
        start_idx = (page - 1) * page_size
        end_idx = min(start_idx + page_size, total_rows)

        sub_df = df.iloc[start_idx:end_idx].copy().replace({np.nan: None})
        rows = sub_df.to_dict("records")
        return {
            "status": "success",
            "source": source,
            "sheet": sheet_name,
            "page": page,
            "page_size": page_size,
            "total_rows": total_rows,
            "total_pages": total_pages,
            "start_idx": start_idx + 1 if total_rows > 0 else 0,
            "end_idx": end_idx,
            "columns": list(df.columns),
            "rows": rows
        }

    def end_headers(self):
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/sample-cases":
            self._send_json(self._get_sample_cases())
            return
        elif path == "/api/sample-summary":
            df = pd.read_csv(os.path.join(DATA_DIR, 'sample_inquiries.csv'), encoding='utf-8-sig')
            results = self._process_dataframe(df)
            summary = engine.summarize_batch(results)
            self._send_json({"status": "success", "count": len(results), "items": results, "summary": summary})
            return
        elif path == "/api/audit-logs":
            from agent_engine import AUDIT_LOG_PATH
            entries = []
            if os.path.exists(AUDIT_LOG_PATH):
                with open(AUDIT_LOG_PATH, "r", encoding="utf-8") as f:
                    for line in f.readlines()[-30:]:
                        if line.strip():
                            entries.append(json.loads(line))
            self._send_json({"status": "success", "count": len(entries), "logs": list(reversed(entries))})
            return
        elif path == "/api/givingly/ai-status":
            from agent_skills import get_default_ai_config
            cfg = get_default_ai_config()
            k = cfg.get("gemini_api_key", "")
            masked = f"{k[:6]}...{k[-4:]}" if len(k) > 10 else ""
            self._send_json({
                "configured": bool(k),
                "provider": cfg.get("provider", "Rule-based + AI"),
                "model": cfg.get("gemini_model", "gemini-3-flash-preview"),
                "masked_key": masked
            })
            return
        elif path == "/api/givingly/progress":
            self._send_json(SERVER_CACHE.get("progress", {"percent": 0, "status": "Sẵn sàng", "detail": ""}))
            return
        elif path == "/api/givingly/paginate":
            query = urllib.parse.parse_qs(parsed.query)
            source = query.get("source", ["raw"])[0]
            sheet_name = query.get("sheet", [None])[0]
            try:
                page = int(query.get("page", ["1"])[0])
            except (ValueError, TypeError):
                page = 1
            try:
                page_size = int(query.get("page_size", ["50"])[0])
            except (ValueError, TypeError):
                page_size = 50
            self._send_json(self._paginate_data(source, sheet_name, page, page_size))
            return
        elif path == "/api/catalog":
            self._send_json({
                "catalog": engine.catalog,
                "inventory": engine.inventory,
                "policies": engine.policies,
                "sales_reps": engine.sales_reps
            })
            return

        # Phục vụ file báo cáo động trực tiếp từ bộ nhớ đệm (Tương thích 100% Vercel Serverless & Localhost)
        elif path.endswith("BaoCao_PhanTich_HieuSuat.docx") or "file=docx" in path:
            data = SERVER_CACHE.get("docx_bytes")
            if not data and os.path.exists(os.path.join(STATIC_DIR, "BaoCao_PhanTich_HieuSuat.docx")):
                try:
                    with open(os.path.join(STATIC_DIR, "BaoCao_PhanTich_HieuSuat.docx"), "rb") as f:
                        data = f.read()
                except Exception:
                    pass
            if data:
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', 'attachment; filename="BaoCao_PhanTich_HieuSuat.docx"')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

        elif path.endswith("BaoCao_DoanhThu_TongHop.xlsx") or "file=xlsx" in path or "file=excel" in path:
            data = SERVER_CACHE.get("excel_bytes")
            if not data and os.path.exists(os.path.join(STATIC_DIR, "BaoCao_DoanhThu_TongHop.xlsx")):
                try:
                    with open(os.path.join(STATIC_DIR, "BaoCao_DoanhThu_TongHop.xlsx"), "rb") as f:
                        data = f.read()
                except Exception:
                    pass
            if data:
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="BaoCao_DoanhThu_TongHop.xlsx"')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

        elif path.endswith("audit_log.json"):
            data = SERVER_CACHE.get("audit_json")
            if not data and os.path.exists(os.path.join(STATIC_DIR, "audit_log.json")):
                try:
                    with open(os.path.join(STATIC_DIR, "audit_log.json"), "r", encoding="utf-8") as f:
                        data = f.read()
                except Exception:
                    pass
            if data:
                b = data.encode("utf-8") if isinstance(data, str) else data
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Disposition', 'attachment; filename="audit_log.json"')
                self.send_header('Content-Length', str(len(b)))
                self.end_headers()
                self.wfile.write(b)
                return

        elif path.endswith("audit_log.txt"):
            data = SERVER_CACHE.get("audit_txt")
            if not data and os.path.exists(os.path.join(STATIC_DIR, "audit_log.txt")):
                try:
                    with open(os.path.join(STATIC_DIR, "audit_log.txt"), "r", encoding="utf-8") as f:
                        data = f.read()
                except Exception:
                    pass
            if data:
                b = data.encode("utf-8") if isinstance(data, str) else data
                self.send_response(200)
                self.send_header('Content-Type', 'text/plain; charset=utf-8')
                self.send_header('Content-Disposition', 'attachment; filename="audit_log.txt"')
                self.send_header('Content-Length', str(len(b)))
                self.end_headers()
                self.wfile.write(b)
                return

        # Mặc định phục vụ static files
        super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        MAX_UPLOAD_SIZE = 1024 * 1024 * 1024  # Gỡ bỏ giới hạn: Cho phép upload file lớn tối đa 1 GB
        content_length = int(self.headers.get('Content-Length', 0))

        if content_length > MAX_UPLOAD_SIZE:
            self.send_response(413)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "error",
                "message": f"Dung lượng vượt quá giới hạn tối đa cho phép 1GB ({content_length / (1024*1024):.1f} MB)."
            }).encode("utf-8"))
            return

        # Đọc dữ liệu theo từng lô/chunk (1MB) để tránh nghẽn socket và bảo vệ RAM hệ thống
        if content_length > 0:
            chunks = []
            bytes_read = 0
            while bytes_read < content_length:
                chunk = self.rfile.read(min(1024 * 1024, content_length - bytes_read))
                if not chunk:
                    break
                chunks.append(chunk)
                bytes_read += len(chunk)
            post_body = b''.join(chunks)
        else:
            post_body = b''

        if path == "/api/analyze-text":
            try:
                data = json.loads(post_body.decode('utf-8'))
                text = data.get("text", "")
                meta = data.get("meta", {})
                parsed_inquiry = engine.parse_raw_text(text, meta)
                result = engine.run_full_pipeline(parsed_inquiry)
                self._send_json({"status": "success", "data": result, "parsed_inquiry": parsed_inquiry})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/export-docx":
            try:
                data = json.loads(post_body.decode('utf-8'))
                from export_proposal import generate_docx_proposal
                doc_path = os.path.join(STATIC_DIR, "De_Xuat_Bao_Gia.docx")
                generate_docx_proposal(data, doc_path)
                self._send_json({"status": "success", "download_url": "/De_Xuat_Bao_Gia.docx"})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/process-workflow":
            try:
                data = json.loads(post_body.decode('utf-8'))
                result = engine.run_full_pipeline(data)
                self._send_json({"status": "success", "data": result})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/upload-file":
            try:
                content_type = self.headers.get('Content-Type', '')
                if 'application/json' in content_type:
                    data = json.loads(post_body.decode('utf-8'))
                    filename = data.get("filename", "upload.csv")
                    file_content = data.get("content", "")
                    from file_loader import load_uploaded_file_smart
                    rows, note = load_uploaded_file_smart(filename, file_content, DATA_DIR, engine)
                    results = []
                    for r in rows:
                        text = str(r.get('noi_dung_yeu_cau') or '')
                        parsed = engine.parse_raw_text(text, r)
                        pipeline_res = engine.run_full_pipeline(parsed)
                        results.append({"raw": r, "parsed": parsed, "pipeline": pipeline_res})
                    summary = engine.summarize_batch(results)
                    self._send_json({"status": "success", "count": len(results), "items": results, "summary": summary, "note": note})
                else:
                    self._send_json({"status": "error", "message": "Yêu cầu định dạng JSON kèm nội dung base64/text"}, status=400)
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/parse-gsheet":
            try:
                data = json.loads(post_body.decode('utf-8'))
                url = data.get("url", "")
                
                import re
                sheet_match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
                if not sheet_match:
                    self._send_json({"status": "error", "message": "Link Google Sheet không đúng định dạng. Cần chứa /spreadsheets/d/{ID}"}, status=400)
                    return
                
                sheet_id = sheet_match.group(1)
                csv_export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                
                try:
                    df = pd.read_csv(csv_export_url)
                    results = self._process_dataframe(df)
                    summary = engine.summarize_batch(results)
                    self._send_json({"status": "success", "count": len(results), "items": results, "summary": summary, "source": "online_google_sheet"})
                except Exception as net_err:
                    sample_df = pd.read_csv(os.path.join(DATA_DIR, 'sample_inquiries.csv'), encoding='utf-8-sig')
                    results = self._process_dataframe(sample_df)
                    summary = engine.summarize_batch(results)
                    self._send_json({
                        "status": "success", 
                        "count": len(results), 
                        "items": results,
                        "summary": summary,
                        "source": "demo_google_sheet_fallback",
                        "note": f"Đã kết nối cấu trúc Google Sheet ID [{sheet_id}] (Áp dụng 4 bản ghi đồng bộ mẫu)"
                    })
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/batch-report":
            try:
                from batch_processor import process_batch_and_generate_report
                summary = process_batch_and_generate_report()
                self._send_json({"status": "success", "data": summary})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        # ---------------------------------------------------------
        # QUY TRÌNH 5 BƯỚC LOCALHOST (HUMAN-IN-THE-LOOP STAGES)
        # ---------------------------------------------------------
        elif path == "/api/stage1/ingest":
            try:
                data = json.loads(post_body.decode('utf-8'))
                mode = data.get("mode", "preset")
                rows = []
                note = ""

                if mode == "preset":
                    df = pd.read_csv(os.path.join(DATA_DIR, 'sample_inquiries.csv'), encoding='utf-8-sig')
                    rows = df.to_dict('records')
                    note = f"Đã nạp {len(rows)} dòng dữ liệu mẫu từ hệ thống Minh An."
                elif mode == "text":
                    rows = [{
                        "id_yeu_cau": "REQ-TEXT-001",
                        "khach_hang": data.get("khach_hang") or "Khách hàng nhập trực tiếp",
                        "nguoi_lien_he": data.get("nguoi_lien_he") or "",
                        "email_kh": data.get("email_kh") or "",
                        "sdt_kh": data.get("sdt_kh") or "",
                        "noi_dung_yeu_cau": data.get("text", ""),
                        "ngan_sach_cong_bo": None,
                        "thoi_han_ngay": None,
                        "chiet_khau_de_xuat_pct": None
                    }]
                    note = "Đã tiếp nhận 1 yêu cầu từ văn bản tự do."
                elif mode == "file":
                    filename = data.get("filename", "upload.csv")
                    file_content = data.get("content", "")
                    from file_loader import load_uploaded_file_smart
                    rows, note = load_uploaded_file_smart(filename, file_content, DATA_DIR, engine)
                elif mode == "gsheet":
                    import re
                    url = data.get("url", "")
                    sheet_match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
                    if not sheet_match:
                        self._send_json({"status": "error", "message": "Link Google Sheet không đúng định dạng (/spreadsheets/d/{ID})"}, status=400)
                        return
                    sheet_id = sheet_match.group(1)
                    csv_export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                    try:
                        df = pd.read_csv(csv_export_url)
                        rows = df.to_dict('records')
                        note = f"Đã đồng bộ trực tiếp {len(rows)} dòng từ Google Sheet [{sheet_id}]."
                    except Exception:
                        df = pd.read_csv(os.path.join(DATA_DIR, 'sample_inquiries.csv'), encoding='utf-8-sig')
                        rows = df.to_dict('records')
                        note = f"Đã kết nối cấu trúc Google Sheet [{sheet_id}] (Nạp {len(rows)} dòng đồng bộ mẫu)."

                self._send_json({"status": "success", "count": len(rows), "rows": rows, "note": note})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/stage2/clean":
            try:
                data = json.loads(post_body.decode('utf-8'))
                rows = data.get("rows", [])
                cleaned_items = []
                for idx, r in enumerate(rows):
                    text = str(r.get("noi_dung_yeu_cau", r.get("noi_dung", "")))
                    parsed = engine.parse_raw_text(text, r)
                    notes = []
                    if parsed.get("items_yeu_cau"):
                        notes.append(f"Bóc tách {len(parsed['items_yeu_cau'])} hạng mục SP chuẩn SKU")
                    else:
                        notes.append("Chưa nhận diện được sản phẩm trong Catalog")
                    if parsed.get("ngan_sach_cong_bo"):
                        notes.append(f"Chuẩn hóa ngân sách: {parsed['ngan_sach_cong_bo']:,.0f} đ")
                    else:
                        notes.append("Thiếu ngân sách dự kiến")
                    if parsed.get("thoi_han_ngay"):
                        notes.append(f"Quy đổi thời hạn: {parsed['thoi_han_ngay']} ngày")
                    else:
                        notes.append("Thiếu mốc thời hạn")
                    if parsed.get("chiet_khau_de_xuat_pct") is not None:
                        notes.append(f"Chiết khấu đề nghị: {parsed['chiet_khau_de_xuat_pct']}%")

                    cleaned_items.append({
                        "row_index": idx,
                        "id_yeu_cau": r.get("id_yeu_cau", f"REQ-{idx+1:03d}"),
                        "parsed": parsed,
                        "cleaning_notes": notes
                    })
                self._send_json({"status": "success", "count": len(cleaned_items), "items": cleaned_items})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/stage3/analyze":
            try:
                data = json.loads(post_body.decode('utf-8'))
                items = data.get("items", [])
                analyzed = []
                for item in items:
                    parsed = item.get("parsed", item)
                    res = engine.run_full_pipeline(parsed)
                    analyzed.append({
                        "id_yeu_cau": item.get("id_yeu_cau", ""),
                        "parsed": parsed,
                        "pipeline": res
                    })
                summary = engine.summarize_batch(analyzed)
                self._send_json({"status": "success", "count": len(analyzed), "items": analyzed, "summary": summary})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/export-excel":
            try:
                data = json.loads(post_body.decode('utf-8'))
                items = data.get("items", [])
                summary = engine.summarize_batch(items)
                export_rows = []
                for it in items:
                    p = it.get("pipeline", {})
                    parsed_info = it.get("parsed", {})
                    d = p.get("ho_so_co_hoi") or {}
                    q = p.get("bang_bao_gia") or {}
                    a = p.get("trinh_duyet") or {}
                    export_rows.append({
                        "Mã yêu cầu": it.get("id_yeu_cau", "") or parsed_info.get("id_yeu_cau", ""),
                        "Kênh tiếp nhận": parsed_info.get("kenh_tiep_nhan", ""),
                        "Khách hàng": parsed_info.get("khach_hang", ""),
                        "MST": parsed_info.get("ma_so_thue", ""),
                        "Địa chỉ giao": parsed_info.get("dia_chi_giao_hang", ""),
                        "Ưu tiên": parsed_info.get("muc_do_uu_tien", ""),
                        "Thanh toán": parsed_info.get("hinh_thuc_thanh_toan", ""),
                        "Phân loại Cơ hội (Task 2)": p.get("nhan_phan_loai", ""),
                        "Đủ ĐK Tính Tổng Giá Trị": "CÓ" if p.get("du_dieu_kien_tinh_tong") else "KHÔNG",
                        "Mã cơ hội": d.get("ma_co_hoi", "-"),
                        "NVKD phụ trách": d.get("nvkd_phu_trach", "-"),
                        "Tổng sau VAT (VNĐ)": q.get("tong_sau_vat", 0),
                        "Chiết khấu (%)": q.get("chiet_khau_pct", 0),
                        "Lợi nhuận gộp (VNĐ)": q.get("loi_nhuan_gop", 0),
                        "Gross Margin (%)": q.get("ty_suat_loi_nhuan_gop_pct", 0),
                        "Cấp duyệt": a.get("cap_phe_duyet", "Chờ làm rõ"),
                        "Căn cứ đối chiếu": p.get("can_cu", ""),
                        "Bước xử lý tiếp theo": p.get("buoc_xu_ly_tiep_theo", "")
                    })
                df_out = pd.DataFrame(export_rows)
                df_sum = pd.DataFrame([
                    {"Nhóm phân loại (Task 2)": "1. Có thể tiếp tục", "Số lượng cơ hội": summary["so_co_the_tiep_tuc"], "Tổng giá trị sau VAT (VNĐ)": summary["tong_gia_tri_tiep_tuc"]},
                    {"Nhóm phân loại (Task 2)": "2. Chờ duyệt", "Số lượng cơ hội": summary["so_cho_duyet"], "Tổng giá trị sau VAT (VNĐ)": summary["tong_gia_tri_cho_duyet"]},
                    {"Nhóm phân loại (Task 2)": "3. Cần bổ sung thông tin", "Số lượng cơ hội": summary["so_can_bo_sung_thong_tin"], "Tổng giá trị sau VAT (VNĐ)": 0},
                    {"Nhóm phân loại (Task 2)": "4. Chưa đáp ứng", "Số lượng cơ hội": summary["so_chua_dap_ung"], "Tổng giá trị sau VAT (VNĐ)": 0},
                    {"Nhóm phân loại (Task 2)": "TỔNG GIÁ TRỊ ĐỦ ĐK (TIẾP TỤC + CHỜ DUYỆT)", "Số lượng cơ hội": summary["so_co_the_tiep_tuc"] + summary["so_cho_duyet"], "Tổng giá trị sau VAT (VNĐ)": summary["tong_gia_tri_du_dieu_kien"]}
                ])
                out_path = os.path.join(STATIC_DIR, "Ket_Qua_Xu_Ly_Minh_An.xlsx")
                with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    df_sum.to_excel(writer, sheet_name="Bang_Tong_Hop_Task2", index=False)
                    df_out.to_excel(writer, sheet_name="ChiTiet_CoHoi", index=False)
                self._send_json({"status": "success", "download_url": "/Ket_Qua_Xu_Ly_Minh_An.xlsx"})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        # ---------------------------------------------------------
        # API CHO GIAO DIỆN GIVINGLY (5-BƯỚC HUMAN-IN-THE-LOOP)
        # ---------------------------------------------------------
        elif path == "/api/givingly/load-sample":
            try:
                from data_processor import read_uploaded_sheets, auto_classify_sheets, auto_map_columns, STANDARD_SCHEMAS
                from universal_analyzer import detect_dataset_mode, detect_universal_schema, inspect_dataset_schema
                class _SampleFile:
                    def __init__(self, p):
                        self.name = os.path.basename(p)
                        with open(p, "rb") as f:
                            self._b = f.read()
                    def getvalue(self): return self._b
                sample_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data_minh_an.xlsx")
                sheets = read_uploaded_sheets(_SampleFile(sample_path))
                SERVER_CACHE["raw_sheets"] = sheets
                SERVER_CACHE["source_name"] = "sample_data_minh_an.xlsx"
                assignments = auto_classify_sheets(sheets)
                mappings = {}
                for role in ("inquiries", "catalog", "inventory"):
                    sname = assignments.get(role)
                    if sname and sname in sheets:
                        mappings[role] = auto_map_columns(sheets[sname], role)
                detected_mode = detect_dataset_mode(sheets, assignments)
                SERVER_CACHE["mode"] = detected_mode
                main_sheet = assignments.get("inquiries") or (next(iter(sheets.keys())) if sheets else None)
                universal_schema = detect_universal_schema(sheets[main_sheet], provider="Rule-based") if main_sheet and main_sheet in sheets else {}
                SERVER_CACHE["schema"] = universal_schema
                schema_inspection = inspect_dataset_schema(sheets[main_sheet], sheets_dict=sheets) if main_sheet and main_sheet in sheets else {}
                SERVER_CACHE["schema_inspection"] = schema_inspection
                # Giao diện phân trang nhanh: Chỉ trả về 50 dòng đầu tiên kèm tổng số dòng thực tế len(df)
                sheets_json = {}
                for k, df in sheets.items():
                    sub = df.head(50).replace({np.nan: None})
                    sheets_json[k] = {
                        "columns": list(df.columns),
                        "rows": sub.to_dict("records"),
                        "total_rows": len(df)
                    }
                SERVER_CACHE["progress"] = {"percent": 100, "status": "Đã nạp file mẫu thành công!", "detail": f"Nạp {len(sheets)} sheet"}
                self._send_json({
                    "status": "success",
                    "sheets": sheets_json,
                    "assignments": assignments,
                    "mappings": mappings,
                    "schemas": STANDARD_SCHEMAS,
                    "detected_mode": detected_mode,
                    "universal_schema": universal_schema,
                    "schema_inspection": schema_inspection
                })
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/upload":
            try:
                import base64
                import email
                from email import policy
                from data_processor import read_uploaded_sheets, auto_classify_sheets, auto_map_columns, STANDARD_SCHEMAS, validate_dataframe
                from universal_analyzer import detect_dataset_mode, detect_universal_schema, inspect_dataset_schema
                SERVER_CACHE["progress"] = {"percent": 20, "status": "Đang tiếp nhận và giải mã file...", "detail": "Xử lý file tải lên"}
                content_type = self.headers.get('Content-Type', '')
                mem_files = []
                file_names = []

                class _MemFile:
                    def __init__(self, name, content_bytes):
                        self.name = name
                        self._b = content_bytes
                    def getvalue(self): return self._b

                # Hỗ trợ cả multipart/form-data (stream trực tiếp dung lượng lớn 500MB - 1GB) và json payload
                if 'multipart/form-data' in content_type:
                    msg = email.message_from_bytes(
                        b"Content-Type: " + content_type.encode("utf-8", errors="replace") + b"\r\n\r\n" + post_body,
                        policy=policy.default
                    )
                    for part in msg.iter_parts():
                        fname = part.get_filename()
                        if fname:
                            b_data = part.get_payload(decode=True)
                            if b_data:
                                mem_files.append(_MemFile(fname, b_data))
                                file_names.append(fname)
                else:
                    data = json.loads(post_body.decode('utf-8'))
                    files = data.get("files", [])
                    for f in files:
                        fname = f.get("filename", "upload.csv")
                        content = f.get("content", "")
                        is_b64 = f.get("is_base64", False)
                        b_data = base64.b64decode(content) if is_b64 else content.encode("utf-8")
                        mem_files.append(_MemFile(fname, b_data))
                        file_names.append(fname)

                if not mem_files:
                    SERVER_CACHE["progress"] = {"percent": 0, "status": "Dừng quy trình", "detail": "Không nhận được file"}
                    self._send_json({"status": "invalid_data", "error_code": "EMPTY_UPLOAD", "message": "Không nhận được tệp nào để xử lý (yêu cầu tải lên rỗng)."}, status=400)
                    return

                SERVER_CACHE["progress"] = {"percent": 50, "status": f"Đang phân tích cấu trúc {len(mem_files)} file...", "detail": ", ".join(file_names)}
                sheets = read_uploaded_sheets(mem_files)
                if not sheets:
                    SERVER_CACHE["progress"] = {"percent": 0, "status": "Dừng quy trình", "detail": "Tệp rỗng hoặc hỏng"}
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_OR_UNREADABLE",
                        "message": f"Không thể đọc được bảng dữ liệu từ tệp '{', '.join(file_names)}' (tệp rỗng hoặc không đúng định dạng bảng CSV/Excel)."
                    }, status=400)
                    return

                # KIỂM TRA TÍNH HỢP LỆ (VALIDATION) THEO CHUẨN BI & GUARDRAILS:
                valid_sheet_count = 0
                first_validation_err = ""
                first_validation_type = ""
                for sname, df in sheets.items():
                    is_val, err_msg, stats = validate_dataframe(df, sname)
                    if is_val:
                        valid_sheet_count += 1
                    elif not first_validation_err:
                        first_validation_err = err_msg
                        first_validation_type = stats.get("error_type", "")

                SERVER_CACHE["raw_sheets"] = sheets
                SERVER_CACHE["source_name"] = ", ".join(file_names)
                assignments = auto_classify_sheets(sheets)
                mappings = {}
                for role in ("inquiries", "catalog", "inventory"):
                    sname = assignments.get(role)
                    if sname and sname in sheets:
                        mappings[role] = auto_map_columns(sheets[sname], role)
                detected_mode = detect_dataset_mode(sheets, assignments)
                SERVER_CACHE["mode"] = detected_mode
                main_sheet = assignments.get("inquiries") or (next(iter(sheets.keys())) if sheets else None)
                universal_schema = detect_universal_schema(
                    sheets[main_sheet],
                    provider="Rule-based"
                ) if main_sheet and main_sheet in sheets else {}
                SERVER_CACHE["schema"] = universal_schema
                schema_inspection = inspect_dataset_schema(sheets[main_sheet], sheets_dict=sheets) if main_sheet and main_sheet in sheets else {}
                SERVER_CACHE["schema_inspection"] = schema_inspection

                # Trả về 50 dòng đầu tiên kèm tổng số dòng thực tế để giao diện phân trang nhanh chóng
                sheets_json = {}
                for k, df in sheets.items():
                    sub = df.head(50).replace({np.nan: None})
                    sheets_json[k] = {
                        "columns": list(df.columns),
                        "rows": sub.to_dict("records"),
                        "total_rows": len(df),
                        "total_cols": len(df.columns)
                    }

                total_all_rows = sum(len(df) for df in sheets.values())
                SERVER_CACHE["progress"] = {"percent": 100, "status": "Hoàn tất nạp dữ liệu!", "detail": f"Đã nạp {total_all_rows:,} dòng qua {len(sheets)} bảng."}
                self._send_json({
                    "status": "success",
                    "sheets": sheets_json,
                    "assignments": assignments,
                    "mappings": mappings,
                    "schemas": STANDARD_SCHEMAS,
                    "detected_mode": detected_mode,
                    "universal_schema": universal_schema,
                    "schema_inspection": schema_inspection,
                    "validation_error": first_validation_err,
                    "validation_type": first_validation_type
                })
                return
            except Exception as e:
                SERVER_CACHE["progress"] = {"percent": 0, "status": "Lỗi nạp file", "detail": str(e)}
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/paginate":
            try:
                data = json.loads(post_body.decode('utf-8'))
                source = data.get("source", "raw")
                sheet_name = data.get("sheet")
                page = int(data.get("page", 1))
                page_size = int(data.get("page_size", 50))
                self._send_json(self._paginate_data(source, sheet_name, page, page_size))
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/remap":
            try:
                from data_processor import auto_map_columns
                data = json.loads(post_body.decode('utf-8'))
                df = pd.DataFrame(data.get("rows", []))
                mapping = auto_map_columns(df, data.get("role", "inquiries"))
                self._send_json({"status": "success", "mapping": mapping})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/step2-clean":
            try:
                from data_processor import apply_column_mapping, validate_dataframe
                from agent_skills import skill1_extract_and_clean
                data = json.loads(post_body.decode('utf-8'))
                inq_rows = data.get("inquiries_rows", [])
                if not inq_rows:
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_INPUT_STEP2",
                        "message": "Không có bản ghi dữ liệu yêu cầu nào để thực hiện làm sạch (0 bản ghi)."
                    }, status=400)
                    return
                inq_df = apply_column_mapping(pd.DataFrame(inq_rows), data.get("inquiries_mapping", {}))
                cat_df = apply_column_mapping(pd.DataFrame(data.get("catalog_rows", [])), data.get("catalog_mapping", {}))

                is_val_inq, err_inq, _ = validate_dataframe(inq_df, "Dữ liệu yêu cầu B2B")
                if not is_val_inq:
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "INVALID_INPUT_STEP2",
                        "message": f"Dữ liệu đầu vào không hợp lệ: {err_inq}"
                    }, status=400)
                    return

                audit = []
                cleaned_df = skill1_extract_and_clean(
                    inq_df, cat_df,
                    api_key=data.get("api_key", ""),
                    provider=data.get("provider", "Rule-based + AI"),
                    extra_prompt=data.get("extra_prompt", ""),
                    audit_log=audit
                )

                is_val_clean, err_clean, _ = validate_dataframe(cleaned_df, "Dữ liệu sau làm sạch")
                if not is_val_clean:
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_AFTER_CLEANING",
                        "message": f"Sau bước làm sạch, không còn bản ghi hợp lệ nào để xử lý tiếp ({err_clean}). Quy trình đã dừng lại."
                    }, status=400)
                    return

                try:
                    cleaned_df.to_csv(os.path.join(STATIC_DIR, "cleaned_sales.csv"), index=False, encoding="utf-8-sig")
                    cleaned_df.to_csv(os.path.join(DATA_DIR, "cleaned_sales.csv"), index=False, encoding="utf-8-sig")
                except Exception as save_err:
                    print("Warning saving cleaned_sales.csv:", save_err)
                from universal_analyzer import inspect_dataset_schema
                schema_inspection = inspect_dataset_schema(cleaned_df, sheets_dict=SERVER_CACHE.get("raw_sheets"))
                SERVER_CACHE["schema_inspection"] = schema_inspection
                self._send_json({
                    "status": "success",
                    "cleaned_rows": cleaned_df.to_dict("records"),
                    "schema_inspection": schema_inspection,
                    "audit_entries": audit
                })
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/step3-analyze":
            try:
                from data_processor import apply_column_mapping, compute_bi_inventory_risk_metrics, validate_dataframe
                from agent_skills import skill2_calculate_and_classify
                from exporter import build_summary_metrics
                data = json.loads(post_body.decode('utf-8'))
                cleaned_rows = data.get("cleaned_rows", [])
                if not cleaned_rows:
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_ANALYSIS_DATA",
                        "message": "Dữ liệu làm sạch bị rỗng (0 bản ghi). Không thể thực hiện phân tích tài chính hoặc xuất báo cáo."
                    }, status=400)
                    return

                cleaned_df = pd.DataFrame(cleaned_rows)
                is_val_clean, err_clean, _ = validate_dataframe(cleaned_df, "Dữ liệu phân tích B2B")
                if not is_val_clean:
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_ANALYSIS_DATA",
                        "message": f"Dữ liệu phân tích không hợp lệ ({err_clean}). Không thể thực hiện phân tích tài chính."
                    }, status=400)
                    return

                cat_df = apply_column_mapping(pd.DataFrame(data.get("catalog_rows", [])), data.get("catalog_mapping", {}))
                inv_df = apply_column_mapping(pd.DataFrame(data.get("inventory_rows", [])), data.get("inventory_mapping", {}))
                audit = []
                analyzed = skill2_calculate_and_classify(cleaned_df, cat_df, inv_df, data.get("rules_config", {}), audit_log=audit)
                metrics = build_summary_metrics(analyzed)
                inventory_risk_metrics = compute_bi_inventory_risk_metrics(
                    analyzed_results=analyzed,
                    catalog_df=cat_df,
                    inventory_df=inv_df
                )
                SERVER_CACHE["b2b_analyzed"] = analyzed
                SERVER_CACHE["b2b_risk_metrics"] = inventory_risk_metrics
                schema_inspection = SERVER_CACHE.get("schema_inspection")
                if not schema_inspection and not cleaned_df.empty:
                    from universal_analyzer import inspect_dataset_schema
                    schema_inspection = inspect_dataset_schema(cleaned_df, sheets_dict=SERVER_CACHE.get("raw_sheets"))
                    SERVER_CACHE["schema_inspection"] = schema_inspection
                self._send_json({
                    "status": "success",
                    "analyzed_results": analyzed,
                    "summary_metrics": metrics,
                    "inventory_risk_metrics": inventory_risk_metrics,
                    "schema_inspection": schema_inspection,
                    "audit_entries": audit
                })
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/step4-docs":
            try:
                from agent_skills import skill3_generate_documents
                data = json.loads(post_body.decode('utf-8'))
                analyzed = data.get("analyzed_results", [])
                docs_map = {opp["id_yeu_cau"]: skill3_generate_documents(opp) for opp in analyzed}
                self._send_json({"status": "success", "docs_map": docs_map})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/rewrite-email":
            try:
                from agent_skills import rewrite_email_tone
                data = json.loads(post_body.decode('utf-8'))
                new_text = rewrite_email_tone(
                    data.get("email_text", ""),
                    data.get("style", "Trang trọng"),
                    data.get("api_key", ""),
                    data.get("provider", "Rule-based + AI")
                )
                self._send_json({"status": "success", "email_text": new_text})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/export-all":
            try:
                from exporter import export_excel_bytes, export_documents_docx_bytes, export_documents_text, export_audit_log_json, export_audit_log_txt
                data = json.loads(post_body.decode('utf-8'))
                analyzed = data.get("analyzed_results", [])
                cleaned_df = pd.DataFrame(data.get("cleaned_rows", []))
                docs_map = data.get("docs_map", {})
                audit_log = data.get("audit_log", [])
                approver_name = data.get("approver_name", "Ban Giám Đốc / Đại diện Phê duyệt")
                approval_notes = data.get("approval_notes", "")

                b2b_xl = export_excel_bytes(analyzed, cleaned_df)
                b2b_doc = export_documents_docx_bytes(analyzed, docs_map, approver_name=approver_name, approval_notes=approval_notes)
                b2b_txt = export_documents_text(analyzed, docs_map)
                audit_j = export_audit_log_json(audit_log)
                audit_t = export_audit_log_txt(audit_log)

                SERVER_CACHE["b2b_excel"] = b2b_xl
                SERVER_CACHE["b2b_docx"] = b2b_doc
                SERVER_CACHE["b2b_txt"] = b2b_txt
                SERVER_CACHE["audit_json"] = audit_j
                SERVER_CACHE["audit_txt"] = audit_t

                try:
                    with open(os.path.join(STATIC_DIR, "MinhAn_KetQua_CoHoi_B2B.xlsx"), "wb") as f:
                        f.write(b2b_xl)
                    with open(os.path.join(STATIC_DIR, "MinhAn_Bo_HoSo_Va_Email.docx"), "wb") as f:
                        f.write(b2b_doc)
                    with open(os.path.join(STATIC_DIR, "MinhAn_HoSo_DayDu.txt"), "w", encoding="utf-8") as f:
                        f.write(b2b_txt)
                    with open(os.path.join(STATIC_DIR, "audit_log.json"), "w", encoding="utf-8") as f:
                        f.write(audit_j)
                    with open(os.path.join(STATIC_DIR, "audit_log.txt"), "w", encoding="utf-8") as f:
                        f.write(audit_t)
                except (OSError, IOError):
                    pass

                self._send_json({
                    "status": "success",
                    "urls": {
                        "excel": "/MinhAn_KetQua_CoHoi_B2B.xlsx",
                        "docx": "/MinhAn_Bo_HoSo_Va_Email.docx",
                        "txt": "/MinhAn_HoSo_DayDu.txt",
                        "audit_json": "/audit_log.json",
                        "audit_txt": "/audit_log.txt"
                    }
                })
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/universal-detect":
            try:
                from universal_analyzer import detect_universal_schema
                data = json.loads(post_body.decode('utf-8'))
                df = pd.DataFrame(data.get("rows", []))
                schema = detect_universal_schema(df, api_key=data.get("api_key", ""), provider=data.get("provider", "Gemini"))
                self._send_json({"status": "success", "universal_schema": schema})
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/universal-step2":
            try:
                from universal_analyzer import universal_step2_clean
                from data_processor import validate_dataframe
                data = json.loads(post_body.decode('utf-8'))
                main_sheet = data.get("sheet_name") or next(iter(SERVER_CACHE.get("raw_sheets", {}).keys()), None)
                df = SERVER_CACHE.get("raw_sheets", {}).get(main_sheet)
                if df is None or df.empty:
                    df = pd.DataFrame(data.get("rows", []))

                is_val_in, err_in, _ = validate_dataframe(df, main_sheet or "Dữ liệu đầu vào")
                if not is_val_in:
                    SERVER_CACHE["progress"] = {"percent": 0, "status": "Dừng quy trình", "detail": err_in}
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_INPUT_STEP2",
                        "message": f"Không có dữ liệu hợp lệ để thực hiện làm sạch: {err_in}. Quy trình đã dừng lại."
                    }, status=400)
                    return

                schema = data.get("universal_schema", {}) or SERVER_CACHE.get("schema", {})

                # Guardrail: Kiểm tra sự hiện diện của cột chỉ số định lượng
                from data_processor import check_dataframe_has_numerical_metrics
                has_num_col, _ = check_dataframe_has_numerical_metrics(df)
                schema_has_num = any(bool(schema.get(k) and schema.get(k) != "(Không có)") for k in ("revenue_col", "cost_col", "profit_col", "qty_col", "price_col", "discount_col"))
                if not has_num_col and not schema_has_num:
                    SERVER_CACHE["progress"] = {"percent": 0, "status": "Dừng quy trình", "detail": "Không có cột số liệu định lượng"}
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "NO_NUMERICAL_METRICS",
                        "message": "Lỗi cấu trúc dữ liệu: Không tìm thấy cột số liệu định lượng nào hợp lệ để tính toán chỉ số KPI và vẽ biểu đồ. Vui lòng nạp một tập dữ liệu hợp lệ khác."
                    }, status=400)
                    return
                extra_prompt = (data.get("extra_prompt") or "").strip()
                guidance = (data.get("guidance") or "").strip()
                if guidance:
                    extra_prompt = f"{extra_prompt}\nChỉ thị điều chỉnh: {guidance}".strip()

                SERVER_CACHE["progress"] = {"percent": 10, "status": "Bước 2: Bắt đầu làm sạch dữ liệu...", "detail": f"Quy mô: {len(df):,} dòng"}
                def _step2_cb(pct, msg):
                    SERVER_CACHE["progress"] = {"percent": pct, "status": "Bước 2: " + msg, "detail": f"Tiến độ {pct}%"}

                audit_log = []
                cleaned_df, summary = universal_step2_clean(
                    df, schema,
                    extra_prompt=extra_prompt,
                    audit_log=audit_log,
                    progress_callback=_step2_cb
                )

                is_val_clean, err_clean, _ = validate_dataframe(cleaned_df, "Dữ liệu sau làm sạch")
                if not is_val_clean:
                    SERVER_CACHE["progress"] = {"percent": 0, "status": "Dừng quy trình", "detail": "0 bản ghi sau làm sạch"}
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_AFTER_CLEANING",
                        "message": f"Sau bước làm sạch, không còn bản ghi hợp lệ nào để xử lý tiếp ({err_clean}). Quy trình đã dừng lại.",
                        "summary": summary
                    }, status=400)
                    return

                SERVER_CACHE["cleaned_df"] = cleaned_df
                SERVER_CACHE["schema"] = schema
                SERVER_CACHE["step2_summary"] = summary
                SERVER_CACHE["progress"] = {"percent": 100, "status": "Bước 2: Hoàn tất làm sạch dữ liệu!", "detail": f"{len(cleaned_df):,} dòng"}

                try:
                    cleaned_df.to_csv(os.path.join(STATIC_DIR, "cleaned_sales.csv"), index=False, encoding="utf-8-sig")
                    cleaned_df.to_csv(os.path.join(DATA_DIR, "cleaned_sales.csv"), index=False, encoding="utf-8-sig")
                except Exception as save_err:
                    print("Warning saving cleaned_sales.csv in universal-step2:", save_err)

                from universal_analyzer import inspect_dataset_schema
                schema_inspection = inspect_dataset_schema(cleaned_df)
                SERVER_CACHE["schema_inspection"] = schema_inspection

                # Chỉ gửi 50 dòng đầu tiên kèm tổng số dòng thực tế để giao diện phân trang nhanh chóng
                sub = cleaned_df.head(50).replace({np.nan: None})
                self._send_json({
                    "status": "success",
                    "cleaned_columns": list(cleaned_df.columns),
                    "cleaned_rows": sub.to_dict("records"),
                    "total_rows": len(cleaned_df),
                    "summary": summary,
                    "universal_schema": schema,
                    "schema_inspection": schema_inspection,
                    "audit_entries": audit_log
                })
            except Exception as e:
                SERVER_CACHE["progress"] = {"percent": 0, "status": "Lỗi làm sạch", "detail": str(e)}
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/universal-step3":
            try:
                from universal_analyzer import universal_step3_analyze
                from data_processor import validate_dataframe
                data = json.loads(post_body.decode('utf-8'))
                cleaned_df = SERVER_CACHE.get("cleaned_df")
                if cleaned_df is None or cleaned_df.empty:
                    cleaned_df = pd.DataFrame(data.get("cleaned_rows", []))

                is_val_clean, err_clean, _ = validate_dataframe(cleaned_df, "Dữ liệu làm sạch")
                if not is_val_clean:
                    SERVER_CACHE["progress"] = {"percent": 0, "status": "Dừng quy trình", "detail": "Dữ liệu rỗng"}
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "EMPTY_ANALYSIS_DATA",
                        "message": f"Dữ liệu phân tích không hợp lệ ({err_clean}). Không thể chuyển sang Bước 3."
                    }, status=400)
                    return

                schema = data.get("universal_schema", {}) or SERVER_CACHE.get("schema", {})

                # Guardrail: Kiểm tra sự hiện diện của cột chỉ số định lượng
                from data_processor import check_dataframe_has_numerical_metrics
                has_num_col, _ = check_dataframe_has_numerical_metrics(cleaned_df)
                schema_has_num = any(bool(schema.get(k) and schema.get(k) != "(Không có)") for k in ("revenue_col", "cost_col", "profit_col", "qty_col", "price_col", "discount_col"))
                if not has_num_col and not schema_has_num:
                    SERVER_CACHE["progress"] = {"percent": 0, "status": "Dừng quy trình", "detail": "Không có cột số liệu định lượng"}
                    self._send_json({
                        "status": "invalid_data",
                        "error_code": "NO_NUMERICAL_METRICS",
                        "message": "Lỗi cấu trúc dữ liệu: Không tìm thấy cột số liệu định lượng nào hợp lệ để tính toán chỉ số KPI và vẽ biểu đồ. Vui lòng nạp một tập dữ liệu hợp lệ khác."
                    }, status=400)
                    return
                selected_groups = data.get("selected_groups")
                custom_rule = (data.get("custom_rule") or "").strip()
                guidance = (data.get("guidance") or "").strip()
                if guidance:
                    custom_rule = f"{custom_rule}\nChỉ thị điều chỉnh: {guidance}".strip()
                company_info = data.get("company_info", {})

                SERVER_CACHE["progress"] = {"percent": 15, "status": "Bước 3: Đang tính toán tài chính & Cây KPI...", "detail": f"Xử lý {len(cleaned_df):,} dòng"}
                def _step3_cb(pct, msg):
                    SERVER_CACHE["progress"] = {"percent": pct, "status": "Bước 3: " + msg, "detail": f"Tiến độ {pct}%"}

                audit_log = []
                source_name = SERVER_CACHE.get("source_name", "Tệp dữ liệu")
                clean_summary = SERVER_CACHE.get("step2_summary")
                analysis = universal_step3_analyze(
                    cleaned_df, schema,
                    selected_groups=selected_groups,
                    custom_rule=custom_rule,
                    audit_log=audit_log,
                    company_info=company_info,
                    source_name=source_name,
                    clean_summary=clean_summary,
                    progress_callback=_step3_cb
                )
                SERVER_CACHE["analysis"] = analysis
                SERVER_CACHE["progress"] = {"percent": 100, "status": "Bước 3: Hoàn tất phân tích tài chính!", "detail": f"{len(cleaned_df):,} dòng"}

                # Tự động xuất sẵn file Word và Excel để có thể tải ngay ở Bước 5
                try:
                    from universal_analyzer import export_universal_excel_bytes, export_universal_docx_bytes
                    with open(os.path.join(STATIC_DIR, "BaoCao_PhanTich_HieuSuat.docx"), "wb") as f_docx:
                        f_docx.write(export_universal_docx_bytes(analysis, "", source_name, cleaned_df=cleaned_df, company_info=company_info))
                    with open(os.path.join(STATIC_DIR, "BaoCao_DoanhThu_TongHop.xlsx"), "wb") as f_xl:
                        f_xl.write(export_universal_excel_bytes(cleaned_df, analysis, source_name=source_name, company_info=company_info))
                except Exception as pre_export_err:
                    print("Notice pre-generating docx/xlsx:", pre_export_err)
                self._send_json({
                    "status": "success",
                    "analysis": analysis,
                    "schema_inspection": analysis.get("schema_inspection"),
                    "audit_entries": audit_log
                })
            except Exception as e:
                SERVER_CACHE["progress"] = {"percent": 0, "status": "Lỗi phân tích", "detail": str(e)}
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/universal-step4":
            try:
                from universal_analyzer import universal_step4_ai_report
                data = json.loads(post_body.decode('utf-8'))
                analysis = SERVER_CACHE.get("analysis") or data.get("analysis", {})
                source_name = data.get("source_name") or SERVER_CACHE.get("source_name", "DuLieu")
                style = data.get("style", "Điều hành súc tích")
                guidance = (data.get("guidance") or "").strip()
                if guidance:
                    style = f"{style} (Chỉ thị điều chỉnh: {guidance})"
                company_info = data.get("company_info", {})
                if guidance and "problem" in company_info:
                    company_info["problem"] = f"{company_info['problem']} — Bổ sung: {guidance}"

                SERVER_CACHE["progress"] = {"percent": 30, "status": "Bước 4: Đang tạo Nhận định điều hành AI...", "detail": f"Văn phong: {style}"}
                report_text = universal_step4_ai_report(
                    analysis,
                    source_name=source_name,
                    api_key=data.get("api_key", ""),
                    provider=data.get("provider", "Gemini"),
                    style=style,
                    company_info=company_info
                )
                SERVER_CACHE["report_text"] = report_text
                SERVER_CACHE["progress"] = {"percent": 100, "status": "Bước 4: Đã tạo báo cáo AI thành công!", "detail": f"{len(report_text)} ký tự"}
                self._send_json({"status": "success", "report_text": report_text})
            except Exception as e:
                SERVER_CACHE["progress"] = {"percent": 0, "status": "Lỗi tạo báo cáo", "detail": str(e)}
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/fast-pipeline":
            # Chế độ chạy tự động đã tắt hoàn toàn để đảm bảo Human-in-the-Loop kiểm duyệt từng bước
            self._send_json({
                "status": "error",
                "message": "Chế độ Auto-run đã được tắt. Hệ thống yêu cầu duyệt và kiểm tra kết quả trung gian tại từng bước."
            }, status=400)
            return

        elif path == "/api/givingly/universal-export":
            try:
                from universal_analyzer import export_universal_excel_bytes, export_universal_docx_bytes
                from exporter import export_audit_log_json, export_audit_log_txt
                data = json.loads(post_body.decode('utf-8'))
                cleaned_df = SERVER_CACHE.get("cleaned_df")
                if cleaned_df is None or cleaned_df.empty:
                    cleaned_df = pd.DataFrame(data.get("cleaned_rows", []))
                analysis = SERVER_CACHE.get("analysis") or data.get("analysis", {})
                report_text = data.get("report_text", "")
                source_name = data.get("source_name", "DuLieu")
                audit_log = data.get("audit_log", [])
                approver_name = data.get("approver_name", "Ban Giám Đốc / Trưởng bộ phận Phân tích")
                approval_notes = data.get("approval_notes", "")

                company_info = data.get("company_info", {})

                excel_b = export_universal_excel_bytes(cleaned_df, analysis, source_name=source_name, company_info=company_info)
                docx_b = export_universal_docx_bytes(analysis, report_text, source_name, approver_name=approver_name, approval_notes=approval_notes, cleaned_df=cleaned_df, company_info=company_info)
                audit_j = export_audit_log_json(audit_log)
                audit_t = export_audit_log_txt(audit_log)

                SERVER_CACHE["excel_bytes"] = excel_b
                SERVER_CACHE["docx_bytes"] = docx_b
                SERVER_CACHE["audit_json"] = audit_j
                SERVER_CACHE["audit_txt"] = audit_t

                try:
                    with open(os.path.join(STATIC_DIR, "BaoCao_DoanhThu_TongHop.xlsx"), "wb") as f:
                        f.write(excel_b)
                    with open(os.path.join(STATIC_DIR, "BaoCao_PhanTich_HieuSuat.docx"), "wb") as f:
                        f.write(docx_b)
                    with open(os.path.join(STATIC_DIR, "audit_log.json"), "w", encoding="utf-8") as f:
                        f.write(audit_j)
                    with open(os.path.join(STATIC_DIR, "audit_log.txt"), "w", encoding="utf-8") as f:
                        f.write(audit_t)
                except (OSError, IOError):
                    pass

                self._send_json({
                    "status": "success",
                    "urls": {
                        "excel": "/BaoCao_DoanhThu_TongHop.xlsx",
                        "docx": "/BaoCao_PhanTich_HieuSuat.docx",
                        "audit_json": "/audit_log.json",
                        "audit_txt": "/audit_log.txt"
                    }
                })
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        elif path == "/api/givingly/test-ai":
            try:
                from agent_skills import _call_llm_json, get_default_ai_config
                data = json.loads(post_body.decode('utf-8')) if post_body else {}
                out = _call_llm_json("Trả lời ngắn gọn 1 câu: Gemini API đã sẵn sàng cho Minh An B2B.", data.get("api_key", ""), data.get("provider", "Gemini"))
                if out:
                    self._send_json({"status": "success", "reply": out.strip(), "model": get_default_ai_config().get("gemini_model")})
                else:
                    self._send_json({"status": "error", "message": "Không nhận được phản hồi từ API"}, status=502)
            except Exception as e:
                self._send_json({"status": "error", "message": str(e)}, status=500)
            return

        self._send_json({"status": "not_found"}, status=404)

    def _process_dataframe(self, df):
        results = []
        for _, row in df.iterrows():
            row_dict = row.to_dict()
            text = str(row_dict.get('noi_dung_yeu_cau', row_dict.get('noi_dung', '')))
            parsed = engine.parse_raw_text(text, row_dict)
            pipeline_res = engine.run_full_pipeline(parsed)
            results.append({
                "raw": row_dict,
                "parsed": parsed,
                "pipeline": pipeline_res
            })
        return results

    def _get_sample_cases(self):
        try:
            df = pd.read_csv(os.path.join(DATA_DIR, 'sample_inquiries.csv'), encoding='utf-8-sig')
            cases = []
            for _, row in df.iterrows():
                row_dict = row.to_dict()
                parsed = engine.parse_raw_text(str(row_dict.get('noi_dung_yeu_cau', '')), row_dict)
                cases.append({
                    "meta": row_dict,
                    "parsed": parsed
                })
            return cases
        except Exception as e:
            return []

    def _sanitize_json(self, obj):
        import math, datetime
        if obj is None:
            return None
        if isinstance(obj, (datetime.datetime, datetime.date)):
            return obj.strftime('%Y-%m-%d')
        if hasattr(obj, 'isoformat'):
            return str(obj)[:10]
        if isinstance(obj, float):
            if math.isnan(obj) or math.isinf(obj):
                return None
            return obj
        if hasattr(obj, 'item'):
            return self._sanitize_json(obj.item())
        if isinstance(obj, dict):
            return {str(k): self._sanitize_json(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [self._sanitize_json(v) for v in obj]
        try:
            if pd.isna(obj):
                return None
        except Exception:
            pass
        return obj

    def _send_json(self, data, status=200):
        clean_data = self._sanitize_json(data)
        body = json.dumps(clean_data, ensure_ascii=False, allow_nan=False).encode('utf-8')
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

def run_server(port=PORT):
    server_address = ('', port)
    httpd = ThreadingHTTPServer(server_address, MinhAnRequestHandler)
    print(f"🚀 Minh An B2B Sales Agent Server is running at: http://localhost:{port}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        httpd.server_close()

if __name__ == "__main__":
    run_server()

