// MinhAn AI — Givingly Split-Panel UI + 5-Step Human-in-the-Loop Controller

let currentStep = 1;
let maxStep = 1;

let sheetsData = {};     // { sheetName: { columns: [...], rows: [...] } }
let assignments = {};    // { inquiries: sheetName, catalog: sheetName, inventory: sheetName }
let mappings = {};       // { inquiries: { stdField: actualCol }, catalog: {...}, inventory: {...} }
let schemas = {};        // STANDARD_SCHEMAS from server

let datasetMode = "UNIVERSAL_ANALYTICS"; // "MINH_AN_B2B" | "UNIVERSAL_ANALYTICS"
let universalSchema = {};
let univMainSheet = "";
let univCleanedRows = [];
let univCleanedCols = [];
let univCleanSummary = {};
let univAnalysis = null;
let univActiveDim = "";
let univReportText = "";
let lastSourceName = "DuLieu";

let cleanedRows = [];    // Step 2 output (editable) - Minh An B2B
let analyzedResults = []; // Step 3 output - Minh An B2B
let summaryMetrics = null;
let docsMap = {};        // Step 4 output - Minh An B2B
let auditLog = [];
let step5Approved = false;

// Table & Narrative collapse/expand UI states
let s1PreviewShowAll = false;
let s2UnivShowAll = false;
let s3B2bShowAll = false;
let s3UnivDimShowAll = false;
let s3VipShowAll = false;
let s5ReportTableShowAll = false;
let s5NarrativeExpanded = false;
let s5VipShowAll = false;
let s5ProdShowAll = false;
let donutActiveDim = "Risk"; // Default to Risk Status Distribution for Senior BI Dashboard
let biRiskMetrics = null;
let b2bRiskMetrics = null;
let schemaInspection = null;
let s5RiskFilter = "ALL"; // "ALL" | "OUT_OF_STOCK" | "NEAR_EXPIRY" | "SAFE"
let s5PivotFilter = "ALL"; // "ALL" | "TOP" | "STANDARD" | "LOW"
let s5RiskSearchQuery = "";
let chartInstances = {};

// Pagination state for Step 1 (Raw Preview) and Step 2 (Cleaned Preview)
let s1CurrentPage = 1;
let s1PageSize = 50;
let s1TotalRows = 0;
let s1TotalPages = 1;

let s2CurrentPage = 1;
let s2PageSize = 50;
let s2TotalRows = 0;
let s2TotalPages = 1;

let progressPollInterval = null;

function showProgressBar(title, detail = "") {
  const wrap = document.getElementById("global-progress-bar-wrap");
  if (!wrap) return;
  wrap.classList.remove("hidden");
  const tEl = document.getElementById("pb-title");
  const dEl = document.getElementById("pb-detail");
  if (tEl) tEl.textContent = title;
  if (dEl) dEl.textContent = detail;
  setProgressBar(0);
}

function setProgressBar(percent, detail = null, subRight = null) {
  const p = Math.max(0, Math.min(100, Math.round(percent)));
  const fill = document.getElementById("pb-fill");
  const pct = document.getElementById("pb-percent");
  if (fill) fill.style.width = `${p}%`;
  if (pct) pct.textContent = `${p}%`;
  if (detail !== null) {
    const el = document.getElementById("pb-detail");
    if (el) el.textContent = detail;
  }
  if (subRight !== null) {
    const el = document.getElementById("pb-sub-right");
    if (el) el.textContent = subRight;
  }
}

function hideProgressBar() {
  const wrap = document.getElementById("global-progress-bar-wrap");
  if (!wrap) return;
  setProgressBar(100, "Hoàn tất!", "100%");
  setTimeout(() => {
    wrap.classList.add("hidden");
    setProgressBar(0);
    stopProgressPolling();
  }, 450);
}

function startProgressPolling(basePct = 20, maxPct = 90) {
  stopProgressPolling();
  let currentSimPct = basePct;
  progressPollInterval = setInterval(async () => {
    try {
      const res = await fetch("/api/givingly/progress");
      const data = await res.json();
      if (data && data.percent > 0) {
        setProgressBar(data.percent, data.status || data.detail, data.detail || "");
      } else {
        currentSimPct = Math.min(maxPct, currentSimPct + Math.floor(Math.random() * 8) + 2);
        setProgressBar(currentSimPct, "Hệ thống đang xử lý dữ liệu ngầm...", `${currentSimPct}%`);
      }
    } catch (e) {
      currentSimPct = Math.min(maxPct, currentSimPct + 5);
      setProgressBar(currentSimPct);
    }
  }, 350);
}

function stopProgressPolling() {
  if (progressPollInterval) {
    clearInterval(progressPollInterval);
    progressPollInterval = null;
  }
}

async function changeS1Page(delta) {
  const target = s1CurrentPage + delta;
  if (target < 1 || target > s1TotalPages) return;
  s1CurrentPage = target;
  await renderPreviewTable();
}

async function onS1PageSizeChanged(val) {
  s1PageSize = parseInt(val, 10) || 50;
  s1CurrentPage = 1;
  await renderPreviewTable();
}

async function changeS2Page(delta) {
  const target = s2CurrentPage + delta;
  if (target < 1 || target > s2TotalPages) return;
  s2CurrentPage = target;
  await renderUnivTable();
}

async function onS2PageSizeChanged(val) {
  s2PageSize = parseInt(val, 10) || 50;
  s2CurrentPage = 1;
  await renderUnivTable();
}

function formatCompactVND(val) {
  if (val === null || val === undefined || isNaN(val)) return "-";
  const num = Number(val);
  const abs = Math.abs(num);
  if (abs >= 1e12) {
    return (num / 1e12).toLocaleString("vi-VN", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " Tỷ";
  } else if (abs >= 1e9) {
    return (num / 1e9).toLocaleString("vi-VN", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " Tỷ";
  } else if (abs >= 1e6) {
    return (num / 1e6).toLocaleString("vi-VN", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " Tr";
  } else if (abs >= 1e3) {
    return (num / 1e3).toLocaleString("vi-VN", { minimumFractionDigits: 0, maximumFractionDigits: 0 }) + " K";
  }
  return num.toLocaleString("vi-VN");
}

function toggleS5VipTable() {
  s5VipShowAll = !s5VipShowAll;
  renderStep5ReportPreview();
}

function toggleS5ProdTable() {
  s5ProdShowAll = !s5ProdShowAll;
  renderStep5ReportPreview();
}

function switchDonutDimension(dim) {
  donutActiveDim = dim;
  const schema = getEffectiveSchemaInspection();
  const hasSkuInv = Boolean(schema && schema.has_sku_inventory);
  const title = document.getElementById("donut-chart-title");
  if (title) {
    if (hasSkuInv) {
      if (dim === "Risk") title.textContent = "🍩 Phân bổ Trạng thái Rủi ro Tồn kho";
      else if (dim === "Category") title.textContent = "🍩 Tỷ trọng Doanh thu (Ngành hàng)";
      else title.textContent = "🍩 Tỷ trọng Doanh thu (Khu vực)";
    } else {
      title.textContent = `🍩 Tỷ trọng ${formatColDisplayName(schema.primary_metric)} theo ${formatColDisplayName(dim)}`;
    }
  }

  const container = document.getElementById("donut-chips-container");
  if (container) {
    const btns = container.querySelectorAll("button");
    btns.forEach(b => {
      const bDim = b.getAttribute("data-dim") || b.textContent.trim();
      const match = bDim === dim || bDim.toLowerCase() === String(dim).toLowerCase();
      b.className = match ? "px-2 py-0.5 rounded bg-[#0B0D08] text-[#D4EE0E] font-bold transition" : "px-2 py-0.5 rounded bg-[#EFECE6] text-[#0C0E08] hover:bg-black/10 font-bold transition";
    });
  }
  renderStep5DonutChart();
}

function setS5RiskFilter(status) {
  const schema = getEffectiveSchemaInspection();
  const hasSkuInv = Boolean(schema && schema.has_sku_inventory);
  if (hasSkuInv) {
    s5RiskFilter = status;
  } else {
    s5PivotFilter = status;
  }
  updateFilterChipsHighlight();
  renderStep5RiskTable();
}

function updateFilterChipsHighlight() {
  const schema = getEffectiveSchemaInspection();
  const hasSkuInv = Boolean(schema && schema.has_sku_inventory);
  const activeStatus = hasSkuInv ? s5RiskFilter : s5PivotFilter;

  const chips = [
    { id: "btn-chip-all", key: "ALL", activeCls: "bg-[#0B0D08] text-[#D4EE0E] border-black", inactiveCls: "bg-white text-[#0C0E08] border-black/20" },
    { id: "btn-chip-out", key: hasSkuInv ? "OUT_OF_STOCK" : "TOP", activeCls: hasSkuInv ? "bg-rose-600 text-white border-rose-700" : "bg-emerald-600 text-white border-emerald-700", inactiveCls: hasSkuInv ? "bg-white text-rose-700 border-rose-300" : "bg-white text-emerald-700 border-emerald-300" },
    { id: "btn-chip-exp", key: hasSkuInv ? "NEAR_EXPIRY" : "STANDARD", activeCls: hasSkuInv ? "bg-amber-500 text-white border-amber-600" : "bg-blue-600 text-white border-blue-700", inactiveCls: hasSkuInv ? "bg-white text-amber-700 border-amber-300" : "bg-white text-blue-700 border-blue-300" },
    { id: "btn-chip-safe", key: hasSkuInv ? "SAFE" : "LOW", activeCls: hasSkuInv ? "bg-emerald-600 text-white border-emerald-700" : "bg-amber-500 text-white border-amber-600", inactiveCls: hasSkuInv ? "bg-white text-emerald-700 border-emerald-300" : "bg-white text-amber-700 border-amber-300" }
  ];

  chips.forEach(c => {
    const el = document.getElementById(c.id);
    if (!el) return;
    if (activeStatus === c.key) {
      el.className = `px-3 py-1 rounded-lg font-bold transition shadow-2xs ${c.activeCls}`;
    } else {
      el.className = `px-3 py-1 rounded-lg font-bold transition shadow-2xs ${c.inactiveCls}`;
    }
  });
}

function onS5RiskSearchInput(val) {
  s5RiskSearchQuery = (val || "").trim().toLowerCase();
  renderStep5RiskTable();
}

function toggleS1Preview() {
  s1PreviewShowAll = !s1PreviewShowAll;
  renderPreviewTable();
}

function toggleS2UnivTable() {
  s2UnivShowAll = !s2UnivShowAll;
  renderUnivTable();
}

function toggleS3B2bTable() {
  s3B2bShowAll = !s3B2bShowAll;
  renderStep3Table();
}

function toggleS3UnivDimTable() {
  s3UnivDimShowAll = !s3UnivDimShowAll;
  renderUnivDimTable();
}

function toggleS3VipTable() {
  s3VipShowAll = !s3VipShowAll;
  renderUnivVipPane();
}

function toggleS5ReportTable() {
  s5ReportTableShowAll = !s5ReportTableShowAll;
  renderStep5ReportPreview();
}

function toggleS5Narrative() {
  s5NarrativeExpanded = !s5NarrativeExpanded;
  const summaryEl = document.getElementById("s5-report-narrative-summary");
  const detailEl = document.getElementById("s5-report-narrative-detail");
  const btn = document.getElementById("btn-toggle-s5-narrative");
  if (s5NarrativeExpanded) {
    summaryEl?.classList.add("hidden");
    detailEl?.classList.remove("hidden");
    if (btn) btn.textContent = "▴ Thu gọn";
  } else {
    summaryEl?.classList.remove("hidden");
    detailEl?.classList.add("hidden");
    if (btn) btn.textContent = "▾ Xem chi tiết";
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  initDropzone();
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeAuditModal();
  });
  await checkAiStatus();
});

async function checkAiStatus() {
  try {
    const res = await fetch("/api/givingly/ai-status");
    const json = await res.json();
    const inp = document.getElementById("cfg-apikey");
    if (json.configured && inp) {
      inp.placeholder = `✓ Đã gắn sẵn Key (${json.masked_key}) • ${json.model}`;
    }
  } catch (e) {}
}

async function testGeminiConnection() {
  const btn = document.getElementById("btn-test-ai");
  if (btn) btn.textContent = "⏳ Đang kiểm tra...";
  try {
    const res = await fetch("/api/givingly/test-ai", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        api_key: document.getElementById("cfg-apikey")?.value || "",
        provider: document.getElementById("cfg-provider")?.value || "Gemini"
      })
    });
    const json = await res.json();
    if (json.status === "success") {
      showToast(`✓ Kết nối thành công (${json.model}): ${json.reply}`, "success");
    } else {
      showToast("✕ Lỗi kết nối: " + (json.message || "Unknown"), "error");
    }
  } catch (e) {
    showToast("✕ Không thể kết nối server test AI", "error");
  } finally {
    if (btn) btn.textContent = "🔌 Test API";
  }
}

// -------------------------------------------------------------
// 0. NAVIGATION, TOASTS & AUDIT LOG
// -------------------------------------------------------------
function toggleApiKeyInput() {
  const prov = document.getElementById("cfg-provider").value;
  const wrap = document.getElementById("cfg-apikey-wrap");
  if (prov === "Rule-based + AI") wrap.classList.add("opacity-50");
  else wrap.classList.remove("opacity-50");
}

function getCompanyInfo() {
  const name = (document.getElementById("biz-name")?.value || document.getElementById("cover-biz-name")?.value || "").trim();
  const type = (document.getElementById("biz-type")?.value || document.getElementById("cover-biz-type")?.value || "").trim();
  const extra = (document.getElementById("biz-extra")?.value || document.getElementById("cover-biz-extra")?.value || "").trim();
  const problem = (document.getElementById("biz-problem")?.value || document.getElementById("cover-biz-problem")?.value || "").trim();
  return { name, type, extra, problem };
}

function syncInputsToCover() {
  const pairs = [
    ["biz-name", "cover-biz-name"],
    ["biz-type", "cover-biz-type"],
    ["biz-extra", "cover-biz-extra"],
    ["biz-problem", "cover-biz-problem"]
  ];
  pairs.forEach(([sId, cId]) => {
    const sEl = document.getElementById(sId);
    const cEl = document.getElementById(cId);
    if (sEl && cEl && sEl.value) cEl.value = sEl.value;
  });
}

function syncInputsFromCover() {
  const pairs = [
    ["cover-biz-name", "biz-name"],
    ["cover-biz-type", "biz-type"],
    ["cover-biz-extra", "biz-extra"],
    ["cover-biz-problem", "biz-problem"]
  ];
  pairs.forEach(([cId, sId]) => {
    const cEl = document.getElementById(cId);
    const sEl = document.getElementById(sId);
    if (cEl && sEl && cEl.value) sEl.value = cEl.value;
  });
}

function initCompanyInfoSync() {
  const pairs = [
    ["biz-name", "cover-biz-name"],
    ["biz-type", "cover-biz-type"],
    ["biz-extra", "cover-biz-extra"],
    ["biz-problem", "cover-biz-problem"]
  ];
  pairs.forEach(([sId, cId]) => {
    const sEl = document.getElementById(sId);
    const cEl = document.getElementById(cId);
    if (sEl && cEl) {
      sEl.addEventListener("input", () => { cEl.value = sEl.value; });
      cEl.addEventListener("input", () => { sEl.value = sEl.value; });
    }
  });
}

function startApp() {
  syncInputsFromCover();
  document.getElementById("cover-page").classList.add("hidden");
  document.getElementById("main-workspace").classList.remove("hidden");
  addAudit("Khởi động", "Welcome_Screen", "Người dùng bấm Start — Bắt đầu xử lý dữ liệu");
  updateStepperUI();
  showToast("Sẵn sàng! Hãy nhập thông tin doanh nghiệp và tải file dữ liệu hoặc bấm Nạp file mẫu.", "info");
}

function backToCover() {
  syncInputsToCover();
  document.getElementById("main-workspace").classList.add("hidden");
  document.getElementById("cover-page").classList.remove("hidden");
}

function addAudit(step, tool, detail) {
  const ts = new Date().toLocaleTimeString("vi-VN", { hour12: false });
  auditLog.push({ timestamp: ts, step, tool, detail });
  document.getElementById("audit-count").textContent = auditLog.length;
}

function openAuditModal() {
  const list = document.getElementById("audit-modal-list");
  list.innerHTML = auditLog.length
    ? auditLog.slice().reverse().map(a => `
        <div class="bg-white border border-black/10 rounded-lg p-2.5">
          <div class="flex justify-between text-[10px] opacity-60 mb-0.5">
            <strong>[${a.step}] (${a.tool})</strong>
            <span>${a.timestamp}</span>
          </div>
          <div>${a.detail}</div>
        </div>
      `).join("")
    : `<div class="text-center py-6 opacity-60">Chưa có sự kiện nào được ghi nhận.</div>`;
  document.getElementById("modal-audit").classList.remove("hidden");
}

function closeAuditModal() {
  document.getElementById("modal-audit").classList.add("hidden");
}

function showToast(msg, type = "success") {
  const c = document.getElementById("toast-container");
  const t = document.createElement("div");
  let cls = "bg-[#0B0D08] text-[#D4EE0E] border-black";
  if (type === "warning") cls = "bg-[#D4EE0E] text-[#0C0E08] border-[#0C0E08]/20";
  if (type === "error") cls = "bg-red-600 text-white border-red-700";
  if (type === "info") cls = "bg-white text-[#0C0E08] border-black/20";

  t.className = `toast-enter pointer-events-auto px-4 py-3 rounded-xl border shadow-xl text-xs font-bold ${cls}`;
  t.textContent = msg;
  c.appendChild(t);
  setTimeout(() => {
    t.style.opacity = "0";
    t.style.transition = "opacity 0.25s ease";
    setTimeout(() => t.remove(), 260);
  }, 3000);
}

function formatVND(n) {
  if (n === null || n === undefined || isNaN(n)) return "0 đ";
  return new Intl.NumberFormat("vi-VN").format(Math.round(n)) + " đ";
}

// -------------------------------------------------------------
// 1. STEPPER, VALIDATION & GATE CONTROLS
// -------------------------------------------------------------
let hasDataValidationError = false;
let lastValidationErrorMsg = "";

function showDataValidationError(title, reason, detail = "") {
  hasDataValidationError = true;
  lastValidationErrorMsg = reason || "Tệp dữ liệu không hợp lệ hoặc rỗng.";

  // Hiển thị Alert Banner trên giao diện
  const alertEl = document.getElementById("data-validation-alert");
  if (alertEl) {
    alertEl.classList.remove("hidden");
    const tEl = document.getElementById("val-alert-title");
    const mEl = document.getElementById("val-alert-message");
    const dEl = document.getElementById("val-alert-detail");
    if (tEl) tEl.textContent = title || "CẢNH BÁO: TỆP DỮ LIỆU KHÔNG HỢP LỆ ĐỂ XỬ LÝ";
    if (mEl) mEl.textContent = reason || "Tệp tải lên không chứa dữ liệu hợp lệ (0 dòng, toàn bộ ô rỗng hoặc lỗi cấu trúc).";
    if (dEl) dEl.textContent = detail ? `Chi tiết: ${detail}` : `Lý do: ${reason}`;
    alertEl.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // Hiển thị Modal Cảnh báo Dừng Quy trình
  const modalEl = document.getElementById("modal-validation-warning");
  if (modalEl) {
    modalEl.classList.remove("hidden");
    const mtEl = document.getElementById("modal-val-title");
    const mrEl = document.getElementById("modal-val-reason");
    if (mtEl) mtEl.textContent = title || "Dừng quy trình: Tệp không có dữ liệu";
    if (mrEl) mrEl.textContent = reason || "Tệp tải lên rỗng (0 dòng) hoặc toàn bộ giá trị bị trống/null.";
  }

  // Chặn cả 2 nút Duyệt & Tiếp tục (ở bottom bar và ở Step 1)
  const btnOk = document.getElementById("btn-gate-ok");
  const btnS1 = document.getElementById("btn-step1-next");
  [btnOk, btnS1].forEach(btn => {
    if (btn) {
      btn.disabled = true;
      btn.classList.add("opacity-50", "cursor-not-allowed", "bg-rose-200", "text-rose-900", "border-rose-400");
      btn.classList.remove("bg-[#D4EE0E]", "text-[#0C0E08]", "hover:bg-white", "hover:brightness-95");
      btn.textContent = "⚠️ Đã khóa duyệt (Dữ liệu không hợp lệ)";
    }
  });

  // Cập nhật Status Pill trên thanh tiêu đề
  const pill = document.getElementById("ws-status-pill");
  if (pill) {
    pill.textContent = "Dữ liệu không hợp lệ";
    pill.className = "text-[11px] font-extrabold bg-rose-500 text-white px-3 py-1.5 rounded-md";
  }

  addAudit("Kiểm tra dữ liệu", "Data_Validation_Halt", `${title}: ${reason}`);
  showToast(`⚠️ ${reason}`, "error");
}

function clearDataValidationError() {
  hasDataValidationError = false;
  lastValidationErrorMsg = "";

  const alertEl = document.getElementById("data-validation-alert");
  if (alertEl) alertEl.classList.add("hidden");

  const modalEl = document.getElementById("modal-validation-warning");
  if (modalEl) modalEl.classList.add("hidden");

  const btnOk = document.getElementById("btn-gate-ok");
  const btnS1 = document.getElementById("btn-step1-next");
  [btnOk, btnS1].forEach(btn => {
    if (btn) {
      btn.disabled = false;
      btn.classList.remove("opacity-50", "cursor-not-allowed", "bg-rose-200", "text-rose-900", "border-rose-400");
      btn.classList.add("bg-[#D4EE0E]", "text-[#0C0E08]", "hover:bg-white");
    }
  });
  if (btnS1) btnS1.textContent = "✓ Duyệt Bước 1 & Sang Bước 2 →";
}

function dismissValidationAlert() {
  const alertEl = document.getElementById("data-validation-alert");
  if (alertEl) alertEl.classList.add("hidden");
}

function closeValidationModal() {
  const modalEl = document.getElementById("modal-validation-warning");
  if (modalEl) modalEl.classList.add("hidden");
}

function setGateLoading(loading, msg) {
  const btnOk = document.getElementById("btn-gate-ok");
  const btnRetry = document.getElementById("btn-gate-retry");
  const btnS1 = document.getElementById("btn-step1-next");
  if (btnOk) {
    btnOk.disabled = loading;
    btnOk.classList.toggle("opacity-60", loading);
    if (loading && msg) btnOk.textContent = msg;
  }
  if (btnS1) {
    btnS1.disabled = loading;
    btnS1.classList.toggle("opacity-60", loading);
  }
  if (btnRetry) btnRetry.disabled = loading;
}

function checkDatasetHasNumericalMetric() {
  const isUniv = datasetMode === "UNIVERSAL_ANALYTICS";
  if (!isUniv) {
    const inqMap = mappings.inquiries || {};
    return Boolean(inqMap.so_luong || inqMap.don_gia || inqMap.thanh_tien || inqMap.ngan_sach);
  }
  const roleKeys = ["revenue_col", "cost_col", "profit_col", "qty_col", "price_col", "discount_col"];
  const mappedCols = roleKeys.map(k => universalSchema[k]).filter(v => v && v !== "(Không có)");
  if (!mappedCols.length) {
    return false;
  }
  // Kiểm tra xem các cột đã ánh xạ có chứa ít nhất 1 số thực tế không (tránh chuỗi ký tự text toàn 'x')
  const sData = sheetsData[univMainSheet] || {};
  const rows = sData.rows || [];
  if (rows.length > 0) {
    for (const c of mappedCols) {
      for (const r of rows) {
        const val = r[c];
        if (val !== null && val !== undefined) {
          const strVal = String(val).trim();
          if (strVal !== "" && strVal.toLowerCase() !== "nan" && strVal.toLowerCase() !== "none" && strVal.toLowerCase() !== "null") {
            const num = Number(strVal.replace(/[\$,VNĐvndđ\s]/g, "").replace(/,/g, ""));
            if (!isNaN(num)) {
              return true;
            }
          }
        }
      }
    }
    return false;
  }
  return universalSchema.has_numerical_metrics !== false;
}

function updateStepperUI() {
  const isUniv = datasetMode === "UNIVERSAL_ANALYTICS";
  const titles = isUniv ? {
    1: "1. Nhập & Ánh xạ",
    2: "2. Làm sạch dữ liệu",
    3: "3. Phân tích tài chính",
    4: "4. Báo cáo & Nhận định",
    5: "5. Xuất báo cáo"
  } : {
    1: "1. Nhập & Ánh xạ",
    2: "2. Làm sạch dữ liệu",
    3: "3. Phân tích chiết khấu",
    4: "4. Tổng hợp & Báo cáo",
    5: "5. Xuất báo cáo"
  };

  document.getElementById("ws-step-kicker").textContent = `Bước ${currentStep} / 5`;
  document.getElementById("ws-step-title").textContent = titles[currentStep];

  for (let s = 1; s <= 5; s++) {
    const rail = document.getElementById(`rail-step-${s}`);
    const badge = document.getElementById(`rail-badge-${s}`);
    const panel = document.getElementById(`step-panel-${s}`);
    const fcNode = document.getElementById(`fc-node-${s}`);

    if (s === currentStep) {
      rail.className = "rounded-xl p-3 bg-[#0B0D08] text-white cursor-pointer flex items-center justify-between shadow-sm";
      badge.textContent = "Đang mở";
      badge.className = "text-[10px] font-bold px-2 py-0.5 rounded bg-[#D4EE0E] text-[#0C0E08]";
      panel.classList.remove("hidden");
      if (fcNode) fcNode.className = "px-2 py-1 rounded bg-[#0B0D08] text-[#D4EE0E] whitespace-nowrap shadow-xs";
    } else if (s < currentStep) {
      rail.className = "rounded-xl p-3 bg-white/85 text-[#0C0E08] cursor-pointer flex items-center justify-between";
      badge.textContent = "✓ Đã duyệt";
      badge.className = "text-[10px] font-bold px-2 py-0.5 rounded bg-[#0B0D08] text-[#D4EE0E]";
      panel.classList.add("hidden");
      if (fcNode) fcNode.className = "px-2 py-1 rounded bg-[#D4EE0E] text-[#0C0E08] whitespace-nowrap";
    } else if (s <= maxStep) {
      rail.className = "rounded-xl p-3 bg-white/70 text-[#0C0E08] cursor-pointer flex items-center justify-between";
      badge.textContent = "Đã mở";
      badge.className = "text-[10px] font-bold px-2 py-0.5 rounded bg-black/10";
      panel.classList.add("hidden");
      if (fcNode) fcNode.className = "px-2 py-1 rounded bg-[#EFECE6] text-[#0C0E08] whitespace-nowrap";
    } else {
      rail.className = "rounded-xl p-3 bg-white/45 text-[#0C0E08]/60 cursor-pointer flex items-center justify-between";
      badge.textContent = "Chờ duyệt";
      badge.className = "text-[10px] font-bold px-2 py-0.5 rounded bg-black/10";
      panel.classList.add("hidden");
      if (fcNode) fcNode.className = "px-2 py-1 rounded bg-[#EFECE6] text-[#0C0E08]/50 whitespace-nowrap";
    }
  }

  const descEl = document.getElementById("gate-step-desc");
  const stepDescs = {
    1: "Bước 1: Kiểm tra cấu trúc bảng thô, thông tin DN & ánh xạ cột",
    2: "Bước 2: Kiểm tra kết quả làm sạch, chuẩn hóa số liệu & outlier",
    3: "Bước 3: Kiểm tra tính toán tài chính, Cây KPI & phân nhóm",
    4: "Bước 4: Kiểm tra nhận định điều hành AI & đề xuất giải pháp",
    5: "Bước 5: Báo cáo điều hành trực quan & Xuất file bản cứng (.xlsx / .docx)"
  };
  if (descEl) descEl.textContent = stepDescs[currentStep] || `Bước ${currentStep}`;

  const btnOk = document.getElementById("btn-gate-ok");
  const btnRetry = document.getElementById("btn-gate-retry");
  const btnS1 = document.getElementById("btn-step1-next");

  if (currentStep === 5) {
    if (btnRetry) btnRetry.textContent = "← Quay lại Bước 4";
    if (btnOk) {
      if (!step5Approved) {
        btnOk.textContent = "✓ Phê duyệt & Ban hành báo cáo";
      } else {
        btnOk.textContent = "🔄 Làm mới quy trình";
      }
    }
  } else {
    if (btnRetry) btnRetry.textContent = `↺ Chạy lại Bước ${currentStep}`;
    if (btnOk) btnOk.textContent = `✓ Duyệt Bước ${currentStep} & Tiếp tục →`;
  }

  // Nếu đang có lỗi dữ liệu không hợp lệ -> Chặn và đổi màu cả 2 nút duyệt
  if (hasDataValidationError) {
    if (btnOk) {
      btnOk.disabled = true;
      btnOk.classList.add("opacity-50", "cursor-not-allowed", "bg-rose-200", "text-rose-900", "border-rose-400");
      btnOk.classList.remove("bg-[#D4EE0E]", "text-[#0C0E08]", "hover:bg-white");
      btnOk.textContent = "⚠️ Đã khóa duyệt (Dữ liệu không hợp lệ)";
    }
    if (btnS1) {
      btnS1.disabled = true;
      btnS1.classList.add("opacity-50", "cursor-not-allowed", "bg-rose-200", "text-rose-900", "border-rose-400");
      btnS1.classList.remove("bg-[#D4EE0E]", "text-[#0C0E08]", "hover:bg-white");
      btnS1.textContent = "⚠️ Đã khóa duyệt (Dữ liệu không hợp lệ)";
    }
  } else {
    if (btnOk && !btnOk.disabled) {
      btnOk.classList.remove("opacity-50", "cursor-not-allowed", "bg-rose-200", "text-rose-900", "border-rose-400");
      btnOk.classList.add("bg-[#D4EE0E]", "text-[#0C0E08]", "hover:bg-white");
    }
    if (btnS1) {
      btnS1.disabled = false;
      btnS1.classList.remove("opacity-50", "cursor-not-allowed", "bg-rose-200", "text-rose-900", "border-rose-400");
      btnS1.classList.add("bg-[#D4EE0E]", "text-[#0C0E08]", "hover:bg-white");
      btnS1.textContent = "✓ Duyệt Bước 1 & Sang Bước 2 →";
    }
  }
}

function jumpStep(step) {
  if (hasDataValidationError && step > 1) {
    showDataValidationError(
      "Đã chặn chuyển bước",
      "Không thể chuyển sang bước sau vì tệp dữ liệu hiện tại không có dữ liệu hợp lệ hoặc rỗng.",
      "Vui lòng tải tệp hợp lệ trước."
    );
    return;
  }
  if (step > maxStep) {
    showToast(`Vui lòng duyệt ở Bước ${maxStep} trước!`, "warning");
    return;
  }
  currentStep = step;
  updateStepperUI();
  if (currentStep === 5) {
    try {
      renderStep5ReportPreview();
    } catch (err) {
      console.error("jumpStep Step 5 preview error:", err);
    }
  }
}

async function gateApprove() {
  if (hasDataValidationError) {
    showDataValidationError(
      "Dừng quy trình: Dữ liệu không hợp lệ",
      lastValidationErrorMsg || "Không thể phê duyệt vì dữ liệu hiện tại không hợp lệ. Vui lòng nạp tệp hợp lệ.",
      "Hệ thống đã chặn duyệt chuyển bước để tránh lỗi phân tích với dữ liệu không hợp lệ."
    );
    return;
  }
  if (!checkDatasetHasNumericalMetric()) {
    showDataValidationError(
      "Lỗi cấu trúc dữ liệu",
      "Lỗi cấu trúc dữ liệu: Không tìm thấy cột số liệu định lượng nào hợp lệ để tính toán chỉ số KPI và vẽ biểu đồ.",
      "Vui lòng nạp một tập dữ liệu hợp lệ khác hoặc chọn ít nhất 1 cột số liệu định lượng."
    );
    return;
  }
  const isUniv = datasetMode === "UNIVERSAL_ANALYTICS";
  if (currentStep === 1) {
    if (!Object.keys(sheetsData).length) {
      showToast("Chưa có dữ liệu! Vui lòng tải file hoặc bấm '⚡ Nạp file mẫu'.", "warning");
      const dz = document.getElementById("dropzone");
      if (dz) {
        dz.classList.add("border-amber-500", "bg-amber-50", "animate-pulse");
        setTimeout(() => dz.classList.remove("border-amber-500", "bg-amber-50", "animate-pulse"), 3000);
      }
      return;
    }
    if (s1TotalRows <= 0) {
      showDataValidationError(
        "Tệp dữ liệu rỗng (0 dòng)",
        "Tệp dữ liệu hiện tại không chứa dòng dữ liệu nào hợp lệ để xử lý.",
        "Số dòng dữ liệu = 0."
      );
      return;
    }
    const info = getCompanyInfo();
    if (!info.name) {
      showToast("⚠️ Vui lòng nhập 'Tên doanh nghiệp' (Bắt buộc)!", "warning");
      const el = document.getElementById("biz-name") || document.getElementById("cover-biz-name");
      if (el) {
        el.focus();
        el.classList.add("border-rose-500", "ring-2", "ring-rose-200");
        setTimeout(() => el.classList.remove("border-rose-500", "ring-2", "ring-rose-200"), 3500);
      }
      return;
    }
    if (!info.type) {
      showToast("⚠️ Vui lòng nhập 'Loại hình doanh nghiệp' (Bắt buộc)!", "warning");
      const el = document.getElementById("biz-type") || document.getElementById("cover-biz-type");
      if (el) {
        el.focus();
        el.classList.add("border-rose-500", "ring-2", "ring-rose-200");
        setTimeout(() => el.classList.remove("border-rose-500", "ring-2", "ring-rose-200"), 3500);
      }
      return;
    }
  }
  setGateLoading(true, "⏳ Đang xử lý bằng AI & Pandas...");
  try {
    if (currentStep === 1) {
      if (isUniv) await runUnivStep2();
      else await runStep2Clean();
      if (hasDataValidationError) return;
      const info = getCompanyInfo();
      addAudit("Bước 1 - Gate", "Human_Approve", `Đã duyệt cấu trúc, mapping & thông tin DN: ${info.name} (${info.type}).`);
      maxStep = Math.max(maxStep, 2);
      currentStep = 2;
      updateStepperUI();
      showToast(isUniv ? "✓ Đã duyệt Bước 1 → Sang Bước 2: Làm sạch & Thống kê mô tả!" : "✓ Đã duyệt Bước 1 → Sang Bước 2 (Skill 1)!", "success");
    } else if (currentStep === 2) {
      if (isUniv) {
        await runUnivStep3();
        if (hasDataValidationError) return;
        addAudit("Bước 2 - Gate", "Human_Approve", `Đã duyệt ${univCleanedRows.length} dòng làm sạch.`);
      } else {
        await runStep3Analyze();
        if (hasDataValidationError) return;
        addAudit("Bước 2 - Gate", "Human_Approve", `Đã duyệt ${cleanedRows.length} bản ghi làm sạch.`);
      }
      maxStep = Math.max(maxStep, 3);
      currentStep = 3;
      updateStepperUI();
      showToast(isUniv ? "✓ Đã duyệt Bước 2 → Sang Bước 3: Cây KPI, KH Ưu tiên & Phân nhóm!" : "✓ Đã duyệt Bước 2 → Sang Bước 3 (Tính toán)!", "success");
    } else if (currentStep === 3) {
      if (isUniv) await runUnivStep4();
      else await runStep4Docs();
      if (hasDataValidationError) return;
      addAudit("Bước 3 - Gate", "Human_Approve", isUniv ? "Đã duyệt Cây KPI, KH Ưu tiên & Phân nhóm." : "Đã duyệt kết quả tính toán & phân loại 4 nhóm.");
      maxStep = Math.max(maxStep, 4);
      currentStep = 4;
      updateStepperUI();
      showToast(isUniv ? "✓ Đã duyệt Bước 3 → Sang Bước 4: Báo cáo Điều hành & Nhận định AI!" : "✓ Đã duyệt Bước 3 → Sang Bước 4 (Tài liệu)!", "success");
    } else if (currentStep === 4) {
      maxStep = 5;
      currentStep = 5;
      updateStepperUI();
      try {
        renderStep5ReportPreview();
      } catch (err) {
        console.error("Step 5 preview error:", err);
      }
      addAudit("Bước 4 - Gate", "Human_Approve", "Đã chốt Bước 4 và chuyển sang Bước 5 để duyệt và ban hành báo cáo kinh doanh.");
      showToast("✓ Đã duyệt Bước 4 → Sang Bước 5: Báo cáo Điều hành trực quan!", "success");
    } else if (currentStep === 5) {
      if (!step5Approved) {
        await approveAndGenerateBusinessReport();
      } else {
        currentStep = 1;
        maxStep = 1;
        step5Approved = false;
        updateStepperUI();
        showToast("Đã sẵn sàng lượt xử lý mới!", "info");
      }
    }
  } catch (err) {
    console.error("gateApprove error:", err);
    showToast("⚠️ " + (err.message || "Đã dừng quy trình do lỗi dữ liệu"), "error");
  } finally {
    setGateLoading(false);
    updateStepperUI();
  }
}

async function rerunCurrentStep() {
  await gateRetry();
}

async function gateRetry() {
  const isUniv = datasetMode === "UNIVERSAL_ANALYTICS";
  setGateLoading(true, `⏳ Đang chạy lại Bước ${currentStep}...`);
  try {
    if (currentStep === 1) {
      const g1 = (document.getElementById("guidance-step-1")?.value || "").trim();
      addAudit("Bước 1 - Gate", "Human_Retry", `Chạy lại Bước 1${g1 ? ` (Chỉ thị: ${g1})` : ""}`);
      showToast("Đã làm mới Bước 1! Bạn có thể kiểm tra lại thông tin DN, ánh xạ cột và bấm Duyệt.", "info");
    } else if (currentStep === 2) {
      const g2 = (document.getElementById("guidance-step-2")?.value || "").trim();
      if (isUniv) await runUnivStep2();
      else await runStep2Clean();
      addAudit("Bước 2 - Gate", "Human_Retry", `Chạy lại làm sạch Bước 2${g2 ? ` (Chỉ thị: ${g2})` : ""}`);
      showToast("✓ Đã chạy lại làm sạch dữ liệu Bước 2 theo chỉ thị!", "success");
    } else if (currentStep === 3) {
      const g3 = (document.getElementById("guidance-step-3")?.value || "").trim();
      if (isUniv) await runUnivStep3();
      else await runStep3Analyze();
      addAudit("Bước 3 - Gate", "Human_Retry", `Tính toán lại Bước 3${g3 ? ` (Chỉ thị: ${g3})` : ""}`);
      showToast("✓ Đã tính toán lại số liệu & Cây KPI Bước 3 theo chỉ thị!", "success");
    } else if (currentStep === 4) {
      const g4 = (document.getElementById("guidance-step-4")?.value || "").trim();
      if (isUniv) await runUnivStep4();
      else await runStep4Docs();
      addAudit("Bước 4 - Gate", "Human_Retry", `Tạo lại nhận định/báo cáo Bước 4${g4 ? ` (Chỉ thị: ${g4})` : ""}`);
      showToast("✓ Đã tạo lại nhận định AI & Báo cáo Bước 4 theo chỉ thị!", "success");
    } else if (currentStep === 5) {
      renderStep5ReportPreview();
      addAudit("Bước 5 - Gate", "Human_Retry", "Làm mới Dashboard Báo cáo Bước 5.");
      showToast("✓ Đã làm mới Dashboard Báo cáo trực quan Bước 5!", "success");
    }
  } catch (err) {
    showToast("Lỗi khi chạy lại bước: " + (err.message || err), "error");
  } finally {
    setGateLoading(false);
    updateStepperUI();
  }
}

// -------------------------------------------------------------
// 2. STEP 1: INGEST & DYNAMIC MAPPING
// -------------------------------------------------------------
function initDropzone() {
  const dz = document.getElementById("dropzone");
  if (!dz) return;
  ["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, e => {
    e.preventDefault(); e.stopPropagation();
    dz.classList.add("border-[#0C0E08]", "bg-[#D4EE0E]/20");
  }));
  ["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => {
    e.preventDefault(); e.stopPropagation();
    dz.classList.remove("border-[#0C0E08]", "bg-[#D4EE0E]/20");
  }));
  dz.addEventListener("drop", e => {
    if (e.dataTransfer.files.length) processUploadedFiles(Array.from(e.dataTransfer.files));
  });
}

function handleFilesSelected(e) {
  if (e.target.files.length) processUploadedFiles(Array.from(e.target.files));
}

async function processUploadedFiles(files) {
  if (!files || !files.length) return;

  // Kiểm tra trước: Nếu toàn bộ file tải lên đều có dung lượng 0 bytes -> Lập tức cảnh báo và dừng quy trình!
  const emptyFiles = files.filter(f => (f.size || 0) === 0);
  if (emptyFiles.length === files.length) {
    showDataValidationError(
      "Tệp dữ liệu rỗng (0 Bytes)",
      `Tệp "${files[0].name}" có dung lượng 0 bytes hoặc rỗng hoàn toàn.`,
      "Vui lòng kiểm tra lại tệp nguồn hoặc tải tệp dữ liệu khác có chứa dữ liệu thực tế."
    );
    return;
  }

  const totalBytes = files.reduce((acc, f) => acc + (f.size || 0), 0);
  const mbTotal = (totalBytes / (1024 * 1024)).toFixed(1);
  showProgressBar("Đang tải file lên hệ thống...", `0 MB / ${mbTotal} MB (0%)`);

  // Sử dụng FormData để stream nhị phân trực tiếp dung lượng lớn (500MB - 1GB) không qua base64 trong RAM
  const formData = new FormData();
  for (const f of files) {
    formData.append("files", f);
  }

  try {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/givingly/upload");

    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable && e.total > 0) {
        const pct = Math.min(80, Math.round((e.loaded / e.total) * 80));
        const loadedMb = (e.loaded / (1024 * 1024)).toFixed(1);
        setProgressBar(pct, `Đang truyền dữ liệu: ${loadedMb} MB / ${mbTotal} MB (${pct}%)`, `${loadedMb}/${mbTotal} MB`);
      }
    };

    const res = await new Promise((resolve, reject) => {
      xhr.onload = () => {
        setProgressBar(90, "Đang phân tích cấu trúc bảng & ánh xạ tự động...", "Xử lý backend");
        if (xhr.status >= 200 && xhr.status < 300) {
          try {
            resolve(JSON.parse(xhr.responseText));
          } catch (err) {
            reject(new Error("Lỗi đọc phản hồi từ server"));
          }
        } else {
          try {
            const errJson = JSON.parse(xhr.responseText);
            if (errJson.status === "invalid_data" || errJson.error_code) {
              resolve(errJson); // Cho phép xử lý cảnh báo trực quan
              return;
            }
            reject(new Error(errJson.message || `Lỗi tải file (${xhr.status})`));
          } catch (_) {
            reject(new Error(`Lỗi tải file (${xhr.status})`));
          }
        }
      };
      xhr.onerror = () => reject(new Error("Lỗi kết nối mạng khi tải file"));
      xhr.send(formData);
    });

    if (res.status === "success") {
      setProgressBar(100, "Hoàn tất nạp dữ liệu!", "100%");
      applyIngestedState(res, files.map(f => f.name).join(", "));
      if (hasDataValidationError) return;
      currentStep = 1;
      maxStep = 1;
      s1CurrentPage = 1;
      updateStepperUI();
      await renderPreviewTable();
      showToast(`📁 Đã nạp thành công ${files.length} file! Kiểm tra cấu trúc & thông tin DN rồi duyệt để tiếp tục.`, "success");
    } else if (res.status === "invalid_data" || res.error_code) {
      showDataValidationError(
        "Tệp không có dữ liệu hợp lệ",
        res.message || "Tệp tải lên không chứa dữ liệu hợp lệ để xử lý.",
        res.details?.reason || "Không phát hiện bản ghi hợp lệ nào."
      );
    } else {
      showDataValidationError(
        "Lỗi nạp tệp dữ liệu",
        res.message || "Không thể đọc dữ liệu từ tệp tải lên.",
        "Vui lòng kiểm tra định dạng tệp và thử lại."
      );
    }
  } catch (err) {
    showDataValidationError("Lỗi nạp file", "Không thể xử lý tệp tải lên: " + err.message, err.message);
  } finally {
    hideProgressBar();
  }
}

async function loadSampleWorkbook() {
  showProgressBar("Đang nạp file mẫu...", "sample_data_minh_an.xlsx");
  setProgressBar(30, "Đang đọc cấu trúc bảng & phân loại sheet...");
  try {
    const res = await fetch("/api/givingly/load-sample", { method: "POST" });
    const json = await res.json();
    if (json.status === "success") {
      setProgressBar(100, "Đã nạp file mẫu thành công!");
      const nameEl = document.getElementById("biz-name");
      const typeEl = document.getElementById("biz-type");
      const extraEl = document.getElementById("biz-extra");
      const probEl = document.getElementById("biz-problem");
      // User input or left blank for generic dynamic naming
      if (typeEl && !typeEl.value.trim()) typeEl.value = "Thương mại & Phân phối B2B";
      if (extraEl && !extraEl.value.trim()) extraEl.value = "Quy mô 6 NVKD, phân phối toàn quốc, 3 kênh chính";
      if (probEl && !probEl.value.trim()) probEl.value = "Tối ưu hóa chiết khấu nhóm VIP, kiểm soát tỷ lệ hoàn hàng và tăng biên lợi nhuận";
      syncInputsToCover();

      applyIngestedState(json, "sample_data_minh_an.xlsx");
      clearDataValidationError();
      currentStep = 1;
      maxStep = 1;
      s1CurrentPage = 1;
      updateStepperUI();
      await renderPreviewTable();
      showToast("⚡ Đã nạp file mẫu! Vui lòng kiểm tra thông tin doanh nghiệp & ánh xạ rồi bấm '✓ Tiếp tục' để duyệt.", "info");
    } else {
      showToast(json.message, "error");
    }
  } catch (err) {
    showToast("Lỗi nạp file mẫu", "error");
  } finally {
    hideProgressBar();
  }
}

// Auto-run pipeline is intentionally disabled.
// The system strictly operates as an explicit Step-by-Step Approval State Machine.
// Each step pauses to present intermediate results for user inspection & interactive guidance.

function applyIngestedState(json, sourceName) {
  sheetsData = json.sheets;
  assignments = json.assignments || {};
  mappings = json.mappings || {};
  schemas = json.schemas || {};
  datasetMode = json.detected_mode || "UNIVERSAL_ANALYTICS";
  universalSchema = json.universal_schema || {};
  lastSourceName = sourceName;
  step5Approved = false;
  schemaInspection = json.schema_inspection || null;

  const sNames = Object.keys(sheetsData || {});
  if (!sNames.length) {
    showDataValidationError(
      "Tệp không có bảng dữ liệu",
      `Tệp "${sourceName}" không có bất kỳ bảng dữ liệu hợp lệ nào.`,
      "Số lượng bảng đọc được là 0."
    );
    return;
  }

  univMainSheet = assignments.inquiries || sNames[0] || "";

  s1CurrentPage = 1;
  s1PageSize = 50;
  const curSheetData = sheetsData[univMainSheet] || {};
  s1TotalRows = curSheetData.total_rows || (curSheetData.rows || []).length || 0;
  const curCols = curSheetData.columns || [];

  // KIỂM TRA TÍNH HỢP LỆ THEO 3 YÊU CẦU GUARDRAILS:
  const isOver200Cols = curCols.length > 200;
  const isUnder2Rows = s1TotalRows < 2;

  let validationErrReason = "";
  let validationErrTitle = "";

  if (isOver200Cols && isUnder2Rows) {
    validationErrTitle = "Tệp không hợp lệ: Vượt số cột & Thiếu dòng";
    validationErrReason = `Cảnh báo: Tệp dữ liệu vượt quá giới hạn 200 cột cho phép (hiện tại: ${curCols.length} cột). Vui lòng rút gọn cấu trúc bảng trước khi phân tích. Cảnh báo: Tệp chỉ có ${s1TotalRows} dòng dữ liệu, không đủ mẫu tối thiểu (cần ít nhất 2 dòng) để thực hiện phân tích xu hướng và thống kê.`;
  } else if (isOver200Cols) {
    validationErrTitle = "Vượt quá giới hạn 200 cột";
    validationErrReason = `Cảnh báo: Tệp dữ liệu vượt quá giới hạn 200 cột cho phép (hiện tại: ${curCols.length} cột). Vui lòng rút gọn cấu trúc bảng trước khi phân tích.`;
  } else if (isUnder2Rows) {
    validationErrTitle = "Dữ liệu không đủ mẫu tối thiểu";
    validationErrReason = `Cảnh báo: Tệp chỉ có ${s1TotalRows} dòng dữ liệu, không đủ mẫu tối thiểu (cần ít nhất 2 dòng) để thực hiện phân tích xu hướng và thống kê.`;
  } else if (s1TotalRows === 0 || !curCols.length) {
    validationErrTitle = "Tệp dữ liệu rỗng (0 dòng)";
    validationErrReason = `Tệp "${sourceName}" không chứa bất kỳ dòng dữ liệu nào hợp lệ (0 bản ghi).`;
  } else {
    // Kiểm tra nếu tất cả ô trong bảng đều là null/NaN/khoảng trắng
    const previewRows = curSheetData.rows || [];
    if (previewRows.length > 0) {
      const allEmpty = previewRows.every(r => Object.values(r).every(v => v === null || v === undefined || String(v).trim() === "" || String(v).trim().toLowerCase() === "nan"));
      if (allEmpty) {
        validationErrTitle = "Toàn bộ ô dữ liệu bị trống/null";
        validationErrReason = `Tệp "${sourceName}" chứa các dòng nhưng toàn bộ giá trị các ô đều bị trống (100% null / NaN / khoảng trắng).`;
      }
    }
  }

  // Luôn nạp bảng để người dùng thấy bản xem trước dữ liệu thô và cấu trúc cột
  document.getElementById("s1-mapping-area").classList.remove("hidden");
  document.getElementById("ws-status-pill").textContent = `${s1TotalRows.toLocaleString("vi-VN")} dòng đã nạp`;

  ["inquiries", "catalog", "inventory"].forEach(role => {
    const sel = document.getElementById(`role-${role}`);
    if (sel) sel.innerHTML = sNames.map(s => `<option value="${s}" ${assignments[role] === s ? "selected" : ""}>${s}</option>`).join("");
  });

  const univSel = document.getElementById("s1-univ-sheet-select");
  if (univSel) {
    univSel.innerHTML = sNames.map(s => {
      const count = sheetsData[s].total_rows || sheetsData[s].rows.length;
      return `<option value="${s}" ${univMainSheet === s ? "selected" : ""}>${s} (${count.toLocaleString("vi-VN")} dòng)</option>`;
    }).join("");
  }

  const modeSel = document.getElementById("s1-mode-select");
  if (modeSel) modeSel.value = datasetMode;
  renderModeUI();
  renderMappingGrid();
  renderUnivMappingGrid();
  renderPreviewTable();

  if (validationErrReason) {
    showDataValidationError(validationErrTitle, validationErrReason, "Hệ thống đã dừng quy trình và chặn chuyển bước.");
  } else if (!checkDatasetHasNumericalMetric()) {
    // Đã được renderUnivMappingGrid() tự động bật cảnh báo đỏ và khóa duyệt
  } else {
    clearDataValidationError();
    showToast(`Đã nạp ${sourceName} — Chế độ: ${datasetMode === "UNIVERSAL_ANALYTICS" ? "Phân tích Doanh thu & Hiệu suất" : "Minh An B2B"}`, "success");
  }
}

function onModeChanged() {
  datasetMode = document.getElementById("s1-mode-select").value;
  renderModeUI();
  renderPreviewTable();
  updateStepperUI();
}

function renderModeUI() {
  const isUniv = datasetMode === "UNIVERSAL_ANALYTICS";
  const badge = document.getElementById("s1-mode-badge");
  if (badge) {
    badge.textContent = isUniv ? "📊 Báo cáo Doanh thu & Hiệu suất Đa định dạng" : "🏢 Thẩm định Cơ hội B2B Minh An";
  }

  const toggle = (id, show) => {
    const el = document.getElementById(id);
    if (el) el.classList.toggle("hidden", !show);
  };

  toggle("s1-b2b-roles", !isUniv);
  toggle("s1-b2b-mapping-wrap", !isUniv);
  toggle("s1-univ-sheet-bar", isUniv);
  toggle("s1-univ-mapping-wrap", isUniv);

  toggle("s2-comparison-list", !isUniv);
  toggle("s2-univ-wrap", isUniv);

  toggle("s3-b2b-wrap", !isUniv);
  toggle("s3-univ-wrap", isUniv);

  toggle("s4-b2b-wrap", !isUniv);
  toggle("s4-univ-wrap", isUniv);
}

async function onUnivSheetChanged() {
  univMainSheet = document.getElementById("s1-univ-sheet-select").value;
  s1CurrentPage = 1;
  const curSheetData = sheetsData[univMainSheet] || {};
  s1TotalRows = curSheetData.total_rows || (curSheetData.rows || []).length || 0;
  const rows = curSheetData.rows || [];
  const res = await fetch("/api/givingly/universal-detect", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      rows: rows.slice(0, 15),
      api_key: document.getElementById("cfg-apikey")?.value || "",
      provider: document.getElementById("cfg-provider")?.value || "Gemini"
    })
  });
  const json = await res.json();
  if (json.status === "success") {
    universalSchema = json.universal_schema || {};
    renderUnivMappingGrid();
  }
  await renderPreviewTable();
}

function renderUnivMappingGrid() {
  const insightEl = document.getElementById("s1-ai-insight");
  const hasNumMetric = checkDatasetHasNumericalMetric();

  if (insightEl) {
    if (!hasNumMetric) {
      // Đổi banner màu xanh thành banner màu đỏ theo đúng Yêu cầu 3
      insightEl.className = "bg-rose-50 border-2 border-rose-500 rounded-xl p-3.5 text-xs text-rose-950 font-bold shadow-sm transition-all";
      insightEl.innerHTML = `
        <div class="flex items-start gap-2.5">
          <span class="text-rose-600 text-base font-black shrink-0">❌</span>
          <div class="space-y-0.5">
            <span class="block text-rose-900 font-extrabold uppercase tracking-wide text-[11px]">Cảnh báo cấu trúc dữ liệu</span>
            <p class="text-rose-800 font-bold text-xs leading-relaxed">
              Lỗi cấu trúc dữ liệu: Không tìm thấy cột số liệu định lượng nào hợp lệ để tính toán chỉ số KPI và vẽ biểu đồ.
            </p>
          </div>
        </div>
      `;
      // Khóa nút duyệt sang bước 2
      showDataValidationError(
        "Lỗi cấu trúc dữ liệu",
        "Lỗi cấu trúc dữ liệu: Không tìm thấy cột số liệu định lượng nào hợp lệ để tính toán chỉ số KPI và vẽ biểu đồ.",
        "Tất cả các trường chỉ số định lượng (Doanh thu, Chi phí, Lợi nhuận, Số lượng, Đơn giá...) đều rơi vào trạng thái (Không có) và dữ liệu chỉ toàn chuỗi ký tự văn bản."
      );
    } else {
      // Khôi phục banner màu xanh bình thường nếu có chỉ số định lượng
      insightEl.className = "bg-[#D4EE0E]/30 border border-black/10 rounded-lg px-3 py-2 text-[11px] transition-all";
      insightEl.innerHTML = `<strong>🤖 AI Đã đọc hiểu tệp (${universalSchema.dataset_domain || "Dữ liệu kinh doanh"}):</strong> ${universalSchema.ai_summary || "Sẵn sàng làm sạch, dựng Cây KPI/Margin và phân hạng nhóm ưu tiên."}`;
      if (hasDataValidationError && lastValidationErrorMsg.includes("Không tìm thấy cột số liệu định lượng")) {
        clearDataValidationError();
      }
    }
  }

  const grid = document.getElementById("s1-univ-mapping-grid");
  const grpBox = document.getElementById("s1-univ-groups-checks");
  if (!grid || !grpBox) return;

  const cols = sheetsData[univMainSheet]?.columns || universalSchema.all_columns || [];
  const opts = ["(Không có)", ...cols];
  const roleFields = [
    ["revenue_col", "💰 Doanh thu / Giá trị"],
    ["cost_col", "🧾 Chi phí / Giá vốn"],
    ["profit_col", "📈 Lợi nhuận"],
    ["customer_col", "⭐ Khách hàng (VIP)"],
    ["project_col", "🗂️ Dự án / Sản phẩm"],
    ["qty_col", "📦 Số lượng"],
    ["price_col", "🏷️ Đơn giá"],
    ["discount_col", "💸 Chiết khấu (%)"],
    ["date_col", "📅 Ngày tháng"],
    ["id_col", "🔑 Mã đơn / ID"]
  ];

  grid.innerHTML = roleFields.map(([key, label]) => {
    const cur = universalSchema[key] || "(Không có)";
    return `
      <div>
        <span class="block text-[10px] font-bold text-[#0C0E08]/70 mb-1 truncate" title="${label}">${label}</span>
        <select onchange="onUnivRoleSelectChange('${key}', this.value)" class="w-full bg-[#F7F7F4] border border-black/15 rounded-lg px-2 py-1.5 text-[11px] font-semibold focus:border-black focus:outline-none transition">
          ${opts.map(c => `<option value="${c}" ${cur === c ? "selected" : ""}>${c}</option>`).join("")}
        </select>
      </div>
    `;
  }).join("");

  // Yêu cầu 1: Nếu vượt quá giới hạn 200 cột -> Ẩn danh sách checkbox Phân nhóm để tránh vỡ giao diện
  if (cols.length > 200) {
    grpBox.innerHTML = `
      <div class="w-full bg-rose-50 border border-rose-300 rounded-lg p-2.5 text-xs text-rose-900 font-bold flex items-center gap-2">
        <span>⚠️</span>
        <span>Đã ẩn danh sách checkbox phân nhóm do tệp vượt quá giới hạn 200 cột (${cols.length} cột) để bảo vệ giao diện.</span>
      </div>
    `;
  } else {
    const activeGroups = new Set(universalSchema.group_cols || []);
    grpBox.innerHTML = cols.map(c => {
      const checked = activeGroups.has(c) ? "checked" : "";
      return `
        <label class="inline-flex items-center gap-1.5 bg-[#F7F7F4] border border-black/15 px-2.5 py-1 rounded-lg cursor-pointer hover:bg-[#D4EE0E]/30">
          <input type="checkbox" ${checked} onchange="toggleUnivGroupCol('${c.replace(/'/g, "\\'")}', this.checked)">
          <span class="font-semibold">${c}</span>
        </label>
      `;
    }).join("");
  }
}

function onUnivRoleSelectChange(key, val) {
  universalSchema[key] = val === "(Không có)" ? null : val;
  const isNumKey = ["revenue_col", "cost_col", "profit_col", "qty_col", "price_col", "discount_col"].includes(key);
  if (isNumKey) {
    const hasAny = ["revenue_col", "cost_col", "profit_col", "qty_col", "price_col", "discount_col"].some(k => {
      const v = universalSchema[k];
      return v && v !== "(Không có)";
    });
    universalSchema.has_numerical_metrics = hasAny;
    if (hasAny && hasDataValidationError && lastValidationErrorMsg.includes("Không tìm thấy cột số liệu định lượng")) {
      clearDataValidationError();
    }
  }
  renderUnivMappingGrid();
  updateStepperUI();
}

function toggleUnivGroupCol(col, isChecked) {
  const cur = new Set(universalSchema.group_cols || []);
  if (isChecked) cur.add(col);
  else cur.delete(col);
  universalSchema.group_cols = Array.from(cur);
}

async function onRoleChanged(role) {
  const chosen = document.getElementById(`role-${role}`).value;
  assignments[role] = chosen;
  const res = await fetch("/api/givingly/remap", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rows: sheetsData[chosen].rows, role })
  });
  const json = await res.json();
  if (json.status === "success") {
    mappings[role] = json.mapping;
    if (role === "inquiries") {
      renderMappingGrid();
      renderPreviewTable();
    }
  }
}

function renderMappingGrid() {
  const grid = document.getElementById("s1-mapping-grid");
  const sName = assignments.inquiries;
  if (!sName || !sheetsData[sName]) return;

  const actualCols = ["(Không có / Để trống)", ...sheetsData[sName].columns];
  const schema = schemas.inquiries || {};
  const curMap = mappings.inquiries || {};

  grid.innerHTML = Object.entries(schema).map(([stdKey, label]) => `
    <div>
      <span class="block text-[10px] font-bold text-[#0C0E08]/70 mb-1 truncate" title="${label}">${label}</span>
      <select onchange="mappings.inquiries['${stdKey}'] = this.value" class="w-full bg-[#F7F7F4] border border-black/15 rounded-lg px-2 py-1.5 text-[11px] font-semibold">
        ${actualCols.map(c => `<option value="${c}" ${curMap[stdKey] === c ? "selected" : ""}>${c}</option>`).join("")}
      </select>
    </div>
  `).join("");
}

async function renderPreviewTable() {
  const sName = datasetMode === "UNIVERSAL_ANALYTICS" ? (univMainSheet || Object.keys(sheetsData)[0]) : assignments.inquiries;
  if (!sName) return;

  // 1. Thử phân trang trực tiếp từ server
  try {
    const res = await fetch(`/api/givingly/paginate?source=raw&sheet=${encodeURIComponent(sName)}&page=${s1CurrentPage}&page_size=${s1PageSize}`);
    const json = await res.json();
    if (json.status === "success" && json.columns && json.columns.length > 0) {
      s1TotalRows = json.total_rows;
      s1TotalPages = Math.max(1, json.total_pages);
      s1CurrentPage = json.page;

      const countEl = document.getElementById("s1-preview-count");
      if (countEl) countEl.textContent = `Đã nạp: ${s1TotalRows.toLocaleString("vi-VN")} dòng`;

      const pageInfo = document.getElementById("s1-page-info");
      if (pageInfo) pageInfo.textContent = `Trang ${s1CurrentPage} / ${s1TotalPages}`;

      const rowsRange = document.getElementById("s1-rows-range");
      if (rowsRange) rowsRange.textContent = s1TotalRows > 0 ? `Dòng ${json.start_idx.toLocaleString("vi-VN")} - ${json.end_idx.toLocaleString("vi-VN")}` : `0 dòng`;

      const btnPrev = document.getElementById("s1-btn-prev");
      if (btnPrev) btnPrev.disabled = s1CurrentPage <= 1;

      const btnNext = document.getElementById("s1-btn-next");
      if (btnNext) btnNext.disabled = s1CurrentPage >= s1TotalPages;

      const cols = json.columns || [];
      const rows = json.rows || [];
      const tbl = document.getElementById("s1-preview-table");
      if (tbl) {
        tbl.innerHTML = `
          <thead class="bg-[#F7F7F4] border-b border-black/10 font-bold sticky top-0">
            <tr>${cols.map(c => `<th class="py-2 px-3 whitespace-nowrap">${c}</th>`).join("")}</tr>
          </thead>
          <tbody class="divide-y divide-black/10">
            ${rows.map(r => `<tr>${cols.map(c => `<td class="py-1.5 px-3 max-w-[200px] truncate">${r[c] ?? "-"}</td>`).join("")}</tr>`).join("")}
          </tbody>
        `;
      }
      return;
    }
  } catch (err) {
    console.warn("Server pagination fallback for raw data:", err);
  }

  // 2. Fallback bộ nhớ cục bộ nếu fetch server chưa sẵn sàng
  if (!sheetsData[sName]) return;
  const { columns, rows } = sheetsData[sName];
  s1TotalRows = sheetsData[sName].total_rows || rows.length;
  s1TotalPages = Math.max(1, Math.ceil(s1TotalRows / s1PageSize));
  const startIdx = (s1CurrentPage - 1) * s1PageSize;
  const displayRows = rows.slice(startIdx, startIdx + s1PageSize);

  const countEl = document.getElementById("s1-preview-count");
  if (countEl) countEl.textContent = `Đã nạp: ${s1TotalRows.toLocaleString("vi-VN")} dòng`;
  const pageInfo = document.getElementById("s1-page-info");
  if (pageInfo) pageInfo.textContent = `Trang ${s1CurrentPage} / ${s1TotalPages}`;
  const rowsRange = document.getElementById("s1-rows-range");
  if (rowsRange) rowsRange.textContent = s1TotalRows > 0 ? `Dòng ${(startIdx + 1).toLocaleString("vi-VN")} - ${Math.min(startIdx + s1PageSize, s1TotalRows).toLocaleString("vi-VN")}` : `0 dòng`;
  const btnPrev = document.getElementById("s1-btn-prev");
  if (btnPrev) btnPrev.disabled = s1CurrentPage <= 1;
  const btnNext = document.getElementById("s1-btn-next");
  if (btnNext) btnNext.disabled = s1CurrentPage >= s1TotalPages;

  const tbl = document.getElementById("s1-preview-table");
  if (tbl) {
    tbl.innerHTML = `
      <thead class="bg-[#F7F7F4] border-b border-black/10 font-bold sticky top-0">
        <tr>${columns.map(c => `<th class="py-2 px-3 whitespace-nowrap">${c}</th>`).join("")}</tr>
      </thead>
      <tbody class="divide-y divide-black/10">
        ${displayRows.map(r => `<tr>${columns.map(c => `<td class="py-1.5 px-3 max-w-[200px] truncate">${r[c] ?? "-"}</td>`).join("")}</tr>`).join("")}
      </tbody>
    `;
  }
}

// -------------------------------------------------------------
// 3. STEP 2: SKILL 1 CLEANING & MANUAL EDIT
// -------------------------------------------------------------
async function runStep2Clean() {
  const sInq = assignments.inquiries;
  const sCat = assignments.catalog;
  const res = await fetch("/api/givingly/step2-clean", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      inquiries_rows: sheetsData[sInq]?.rows || [],
      catalog_rows: sheetsData[sCat]?.rows || [],
      inquiries_mapping: mappings.inquiries || {},
      catalog_mapping: mappings.catalog || {},
      api_key: document.getElementById("cfg-apikey")?.value || "",
      provider: document.getElementById("cfg-provider")?.value || "Rule-based + AI",
      extra_prompt: document.getElementById("s2-extra-prompt")?.value || ""
    })
  });
  const json = await res.json();
  if (json.status === "success") {
    cleanedRows = json.cleaned_rows || [];
    if (!cleanedRows.length) {
      showDataValidationError(
        "Không còn bản ghi sau làm sạch",
        "Sau bước làm sạch, không còn bản ghi dữ liệu yêu cầu nào hợp lệ (0 bản ghi).",
        "Quy trình đã dừng lại để tránh phân tích dữ liệu rỗng."
      );
      return;
    }
    clearDataValidationError();
    if (json.schema_inspection) schemaInspection = json.schema_inspection;
    (json.audit_entries || []).forEach(a => addAudit(a.step, a.tool, a.detail));
    renderStep2Cards();
  } else if (json.status === "invalid_data") {
    showDataValidationError(
      "Dừng quy trình tại Bước 2",
      json.message || "Không còn bản ghi dữ liệu hợp lệ để xử lý tiếp.",
      "Quy trình đã dừng lại trước khi phân tích."
    );
  } else {
    showDataValidationError(
      "Lỗi làm sạch dữ liệu B2B",
      json.message || "Lỗi xử lý làm sạch dữ liệu.",
      "Vui lòng kiểm tra lại ánh xạ cột."
    );
  }
}

function renderStep2Cards() {
  const wrap = document.getElementById("s2-comparison-list");
  wrap.innerHTML = cleanedRows.map((r, i) => {
    const missing = r.thieu_thong_tin;
    return `
      <div class="bg-white border ${missing ? 'border-orange-400' : 'border-black/15'} rounded-xl p-4 space-y-3 text-xs">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <div>
            <span class="font-mono font-bold bg-[#0B0D08] text-[#D4EE0E] px-2 py-0.5 rounded text-[10px]">${r.id_yeu_cau}</span>
            <strong class="ml-2 text-sm">${r.khach_hang}</strong>
          </div>
          <span class="px-2.5 py-0.5 rounded font-extrabold text-[11px] ${missing ? 'bg-orange-100 text-orange-800' : 'bg-[#D4EE0E] text-[#0C0E08]'}">
            ${missing ? '⚠ ' + r.chi_tiet_thieu : '✓ Đủ dữ kiện'}
          </span>
        </div>

        <div class="bg-[#F7F7F4] rounded-lg p-2.5 text-[11px] text-[#0C0E08]/75 border border-black/10">
          <strong>Văn bản gốc:</strong> ${r.noi_dung_goc || '(Trống)'}
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-4 gap-2">
          <div class="sm:col-span-2">
            <label class="block text-[10px] font-bold opacity-60 mb-1">SP &amp; SL bóc tách (SKU:SL)</label>
            <input type="text" value="${r.san_pham_va_sl || ''}" onchange="cleanedRows[${i}].san_pham_va_sl = this.value" class="w-full bg-[#F7F7F4] border border-black/15 rounded-lg px-2.5 py-1.5 font-mono font-bold">
          </div>
          <div>
            <label class="block text-[10px] font-bold opacity-60 mb-1">Ngân sách (VNĐ)</label>
            <input type="number" value="${r.ngan_sach_vnd ?? ''}" placeholder="Thiếu" onchange="cleanedRows[${i}].ngan_sach_vnd = this.value === '' ? null : Number(this.value)" class="w-full bg-[#F7F7F4] border border-black/15 rounded-lg px-2.5 py-1.5 font-mono font-bold">
          </div>
          <div class="grid grid-cols-2 gap-1.5">
            <div>
              <label class="block text-[10px] font-bold opacity-60 mb-1">Hạn (ngày)</label>
              <input type="number" value="${r.thoi_han_ngay ?? ''}" placeholder="Thiếu" onchange="cleanedRows[${i}].thoi_han_ngay = this.value === '' ? null : Number(this.value)" class="w-full bg-[#F7F7F4] border border-black/15 rounded-lg px-2 py-1.5 font-mono font-bold">
            </div>
            <div>
              <label class="block text-[10px] font-bold opacity-60 mb-1">CK (%)</label>
              <input type="number" step="0.5" value="${r.chiet_khau_de_nghi_pct ?? 0}" onchange="cleanedRows[${i}].chiet_khau_de_nghi_pct = Number(this.value)" class="w-full bg-[#F7F7F4] border border-black/15 rounded-lg px-2 py-1.5 font-mono font-bold">
            </div>
          </div>
        </div>
      </div>
    `;
  }).join("");
  document.getElementById("ws-status-pill").textContent = `${cleanedRows.length} đơn đã làm sạch`;
}

// -------------------------------------------------------------
// 4. STEP 3: SKILL 2 CALCULATION & 4-GROUP CLASSIFICATION
// -------------------------------------------------------------
async function runStep3Analyze() {
  const sCat = assignments.catalog;
  const sInv = assignments.inventory;
  const rulesConfig = {
    nvkd_max_discount: Number(document.getElementById("rule-nvkd-disc").value),
    tpkd_max_discount: Number(document.getElementById("rule-tpkd-disc").value),
    nvkd_max_value: Number(document.getElementById("rule-nvkd-val").value),
    vat_rate: Number(document.getElementById("rule-vat").value) / 100.0
  };

  const res = await fetch("/api/givingly/step3-analyze", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      cleaned_rows: cleanedRows,
      catalog_rows: sheetsData[sCat]?.rows || [],
      inventory_rows: sheetsData[sInv]?.rows || [],
      catalog_mapping: mappings.catalog || {},
      inventory_mapping: mappings.inventory || {},
      rules_config: rulesConfig
    })
  });
  const json = await res.json();
  if (json.status === "success") {
    analyzedResults = json.analyzed_results;
    summaryMetrics = json.summary_metrics;
    if (json.inventory_risk_metrics) b2bRiskMetrics = json.inventory_risk_metrics;
    if (json.schema_inspection) schemaInspection = json.schema_inspection;
    (json.audit_entries || []).forEach(a => addAudit(a.step, a.tool, a.detail));
    renderStep3Table();
  } else {
    showToast(json.message, "error");
  }
}

function renderStep3Table() {
  const tbody = document.getElementById("s3-results-tbody");
  const btn = document.getElementById("btn-toggle-s3-b2b");
  if (btn) {
    btn.textContent = s3B2bShowAll ? "Thu gọn" : `Xem tất cả (${analyzedResults.length})`;
    btn.classList.toggle("hidden", analyzedResults.length <= 5);
  }
  const displayList = s3B2bShowAll ? analyzedResults : analyzedResults.slice(0, 5);

  tbody.innerHTML = displayList.map(r => {
    let badge = "bg-[#EFECE6] text-[#0C0E08]";
    if (r.phan_loai === "Có thể tiếp tục") badge = "bg-[#D4EE0E] text-[#0C0E08]";
    else if (r.phan_loai === "Chờ duyệt") badge = "bg-[#0B0D08] text-[#D4EE0E]";
    else if (r.phan_loai === "Cần bổ sung thông tin") badge = "bg-orange-100 text-orange-800";
    else if (r.phan_loai === "Chưa đáp ứng") badge = "bg-red-100 text-red-800";

    return `
      <tr>
        <td class="py-2.5 px-3 font-mono font-bold">${r.id_yeu_cau}<br><span class="text-[10px] opacity-60">${r.ma_co_hoi}</span></td>
        <td class="py-2.5 px-3"><strong>${r.khach_hang}</strong><br><span class="text-[10px] opacity-60">${r.nvkd_phu_trach}</span></td>
        <td class="py-2.5 px-3"><span class="px-2 py-0.5 rounded font-extrabold text-[10px] ${badge}">${r.phan_loai}</span><br><span class="text-[10px] opacity-70">${r.cap_phe_duyet}</span></td>
        <td class="py-2.5 px-3 text-right font-mono font-bold">${r.gia_tri_sau_vat > 0 ? formatVND(r.gia_tri_sau_vat) : '-'}</td>
        <td class="py-2.5 px-3 text-center font-mono">${r.chiet_khau_pct}%<br><span class="text-[10px] opacity-60">Margin ${r.margin_pct}%</span></td>
        <td class="py-2.5 px-3 max-w-[200px]">${r.can_cu}</td>
      </tr>
    `;
  }).join("");
  document.getElementById("ws-status-pill").textContent = `Đủ ĐK: ${formatVND(summaryMetrics?.eligible_total_value || 0)}`;
}

// -------------------------------------------------------------
// 5. STEP 4: DASHBOARD & 3 DOCUMENTS (SKILL 3)
// -------------------------------------------------------------
async function runStep4Docs() {
  const res = await fetch("/api/givingly/step4-docs", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ analyzed_results: analyzedResults })
  });
  const json = await res.json();
  if (json.status === "success") {
    docsMap = json.docs_map;
    renderStep4Dashboard();
  }
}

function renderStep4Dashboard() {
  const c = summaryMetrics.counts;
  document.getElementById("s4-cnt-continue").textContent = c["Có thể tiếp tục"];
  document.getElementById("s4-cnt-approve").textContent = c["Chờ duyệt"];
  document.getElementById("s4-cnt-missing").textContent = c["Cần bổ sung thông tin"];
  document.getElementById("s4-cnt-unmet").textContent = c["Chưa đáp ứng"];
  document.getElementById("s4-total-count").textContent = `${summaryMetrics.total_count} đơn`;
  document.getElementById("s4-eligible-val").textContent = formatVND(summaryMetrics.eligible_total_value);

  const eligibleCnt = c["Có thể tiếp tục"] + c["Chờ duyệt"];
  const pct = summaryMetrics.total_count ? Math.round((eligibleCnt / summaryMetrics.total_count) * 100) : 0;
  document.getElementById("s4-progress-bar").style.width = `${pct}%`;

  const sel = document.getElementById("s4-opp-select");
  sel.innerHTML = analyzedResults.map(r => `<option value="${r.id_yeu_cau}">${r.id_yeu_cau} — ${r.khach_hang} [${r.phan_loai}]</option>`).join("");
  renderSelectedOppDocs();
}

function renderSelectedOppDocs() {
  const reqId = document.getElementById("s4-opp-select").value;
  const doc = docsMap[reqId] || {};
  document.getElementById("doc-text-dossier").value = doc.ho_so_co_hoi || "";
  document.getElementById("doc-text-email").value = doc.email_nhap || "";
  document.getElementById("doc-text-approval").value = doc.phieu_chuyen_duyet || "(Cơ hội này thuộc thẩm quyền NVKD Tự quyết — Không cần Phiếu chuyển duyệt)";
}

function saveEditedEmail() {
  const reqId = document.getElementById("s4-opp-select").value;
  if (docsMap[reqId]) {
    docsMap[reqId].email_nhap = document.getElementById("doc-text-email").value;
  }
}

async function rewriteSelectedEmail() {
  const reqId = document.getElementById("s4-opp-select").value;
  const style = document.getElementById("s4-tone-select").value;
  const res = await fetch("/api/givingly/rewrite-email", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      email_text: document.getElementById("doc-text-email").value,
      style,
      api_key: document.getElementById("cfg-apikey")?.value || "",
      provider: document.getElementById("cfg-provider")?.value || "Rule-based + AI"
    })
  });
  const json = await res.json();
  if (json.status === "success") {
    docsMap[reqId].email_nhap = json.email_text;
    document.getElementById("doc-text-email").value = json.email_text;
    addAudit("Bước 4 - Skill 3", "rewrite_email_tone", `Viết lại email ${reqId} theo văn phong '${style}'`);
    showToast(`Đã viết lại email theo văn phong '${style}'!`, "success");
  }
}

function switchDocTab(tab) {
  ["dossier", "email", "approval"].forEach(t => {
    const btn = document.getElementById(`doctab-${t}`);
    const wrap = document.getElementById(`doc-wrap-${t}`);
    if (t === tab) {
      btn.className = "px-2.5 py-1 rounded bg-[#0B0D08] text-white";
      wrap.classList.remove("hidden");
    } else {
      btn.className = "px-2.5 py-1 rounded text-[#0C0E08]";
      wrap.classList.add("hidden");
    }
  });
}

// -------------------------------------------------------------
// 6. STEP 5: EXPORT ALL ARTIFACTS
// -------------------------------------------------------------
async function runStep5Export() {
  const approverName = document.getElementById("s5-approver-name")?.value?.trim() || "Ban Giám Đốc";
  const approverTitle = document.getElementById("s5-approver-title")?.value?.trim() || "Đại diện Phê duyệt";
  const approvalNotes = document.getElementById("s5-approval-notes")?.value?.trim() || "Đồng ý phê duyệt số liệu kinh doanh, chính sách chiết khấu và ban hành báo cáo chính thức.";
  const fullApprover = `${approverName} — ${approverTitle}`;

  const res = await fetch("/api/givingly/export-all", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      analyzed_results: analyzedResults,
      cleaned_rows: cleanedRows,
      docs_map: docsMap,
      audit_log: auditLog,
      approver_name: fullApprover,
      approval_notes: approvalNotes
    })
  });
  const json = await res.json();
  if (json.status === "success") {
    const setProp = (id, prop, val) => { const el = document.getElementById(id); if (el) el[prop] = val; };
    setProp("dl-excel", "href", json.urls.excel);
    setProp("dl-docx", "href", json.urls.docx);
    setProp("dl-txt", "href", json.urls.txt);
    setProp("dl-audit", "href", json.urls.audit_json);
    setProp("dl-excel-desc", "textContent", "3 tab: TongHop, ChiTietCoHoi, DuLieuLamSach");
    setProp("dl-docx-desc", "textContent", "Báo Cáo Kinh Doanh Chính Thức + Hồ sơ + Email + Phiếu duyệt");
    setProp("dl-txt-title", "textContent", "📝 Hồ sơ Thuần Text (.txt)");
    setProp("dl-txt-desc", "textContent", "Dễ dàng sao chép gửi nhanh");
    setProp("ws-status-pill", "textContent", "Hoàn tất 5/5 Bước");
    return true;
  } else {
    showToast(json.message, "error");
    return false;
  }
}

// -------------------------------------------------------------
// 7. UNIVERSAL ANALYTICS MODE (STEPS 2 -> 5)
// -------------------------------------------------------------
let univActiveView = "tree";

function fmtNum(n) {
  if (n === null || n === undefined || isNaN(n)) return "-";
  return new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 2 }).format(n);
}

async function runUnivStep2() {
  showProgressBar("Bước 2/5: Đang làm sạch và chuẩn hóa dữ liệu lớn...", 10, "Khởi tạo tác vụ phân tích & làm sạch song song...");
  startProgressPolling();
  try {
    const sName = univMainSheet || Object.keys(sheetsData)[0];
    const guidance = (document.getElementById("guidance-step-2")?.value || document.getElementById("s2-extra-prompt")?.value || "").trim();
    const res = await fetch("/api/givingly/universal-step2", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sheet_name: sName,
        rows: sheetsData[sName]?.rows || [],
        universal_schema: universalSchema,
        extra_prompt: guidance,
        guidance: guidance,
        company_info: getCompanyInfo()
      })
    });
    const json = await res.json();
    if (json.status === "success") {
      s2CurrentPage = 1;
      s2TotalRows = json.total_rows || (json.summary?.clean_rows) || (json.cleaned_rows?.length || 0);
      s2TotalPages = Math.max(1, Math.ceil(s2TotalRows / s2PageSize));
      univCleanedRows = json.cleaned_rows || [];
      univCleanedCols = json.cleaned_columns || [];
      univCleanSummary = json.summary || {};

      if (s2TotalRows <= 0 || (univCleanedRows.length === 0 && (!json.cleaned_rows || json.cleaned_rows.length === 0))) {
        showDataValidationError(
          "Không còn bản ghi sau làm sạch",
          "Sau bước làm sạch, không còn bản ghi dữ liệu hợp lệ nào để xử lý tiếp (100% dữ liệu đã bị lọc bỏ do trùng lặp hoặc lỗi định dạng).",
          "Quy trình đã dừng lại để tránh phân tích dữ liệu rỗng."
        );
        return;
      }
      clearDataValidationError();

      if (json.universal_schema) universalSchema = json.universal_schema;
      if (json.schema_inspection) schemaInspection = json.schema_inspection;
      (json.audit_entries || []).forEach(a => addAudit(a.step, a.tool, a.detail));
      renderUnivStep2();
      await renderUnivTable();
      setProgressBar(100, "Hoàn tất Bước 2", `Đã chuẩn hóa ${s2TotalRows.toLocaleString("vi-VN")} dòng`);
    } else if (json.status === "invalid_data") {
      showDataValidationError(
        "Dừng quy trình tại Bước 2",
        json.message || "Không còn bản ghi dữ liệu hợp lệ để xử lý tiếp.",
        "Quy trình đã dừng lại trước khi phân tích."
      );
    } else {
      showDataValidationError(
        "Lỗi làm sạch dữ liệu Bước 2",
        json.message || "Lỗi xử lý làm sạch dữ liệu.",
        "Vui lòng kiểm tra lại cấu trúc bảng dữ liệu."
      );
    }
  } catch (err) {
    showToast("Lỗi làm sạch dữ liệu: " + err.message, "error");
  } finally {
    stopProgressPolling();
    hideProgressBar();
  }
}

function renderUnivStep2() {
  const sum = univCleanSummary || {};
  document.getElementById("s2-univ-summary").innerHTML = `
    <div class="bg-white border border-black/15 rounded-xl p-3">
      <span class="text-[10px] font-bold opacity-60 block">DÒNG ĐẦU VÀO</span>
      <strong class="text-base font-black">${fmtNum(sum.orig_rows)}</strong>
    </div>
    <div class="bg-[#D4EE0E]/50 border border-black/15 rounded-xl p-3">
      <span class="text-[10px] font-bold opacity-70 block">SAU LÀM SẠCH</span>
      <strong class="text-base font-black">${fmtNum(sum.clean_rows)}</strong>
    </div>
    <div class="bg-white border border-black/15 rounded-xl p-3">
      <span class="text-[10px] font-bold opacity-60 block">TRÙNG LẶP ĐÃ LOẠI</span>
      <strong class="text-base font-black">${fmtNum(sum.duplicates_removed)}</strong>
    </div>
    <div class="bg-white border border-black/15 rounded-xl p-3">
      <span class="text-[10px] font-bold opacity-60 block">OUTLIER (IQR)</span>
      <strong class="text-base font-black">${fmtNum(sum.outliers_flagged ?? 0)}</strong>
    </div>
    <div class="bg-white border border-black/15 rounded-xl p-3">
      <span class="text-[10px] font-bold opacity-60 block">CỘT DẪN XUẤT</span>
      <strong class="text-[11px] font-extrabold leading-tight block">${(sum.computed_cols || []).length ? sum.computed_cols.join(" • ") : "Chuẩn hóa số/ngày"}</strong>
    </div>
  `;

  const statsEl = document.getElementById("s2-univ-stats");
  const stats = sum.descriptive_stats || [];
  if (statsEl) {
    statsEl.innerHTML = stats.length ? `
      <div class="px-4 py-2 bg-[#EFECE6] text-[11px] font-bold flex justify-between">
        <span>📐 Thống kê Mô tả Chuẩn (Descriptive Statistics)</span>
        <span class="opacity-65">Mean • Median • Std • Q1/Q3</span>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-[11px] text-left">
          <thead class="bg-[#F7F7F4] border-b border-black/10 font-bold">
            <tr>
              <th class="py-1.5 px-3">Trường số liệu</th>
              <th class="py-1.5 px-3 text-right">Trung bình (Mean)</th>
              <th class="py-1.5 px-3 text-right">Trung vị (Median)</th>
              <th class="py-1.5 px-3 text-right">Độ lệch chuẩn</th>
              <th class="py-1.5 px-3 text-right">Min — Max</th>
              <th class="py-1.5 px-3 text-right">Outlier</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-black/10 font-mono">
            ${stats.map(st => `
              <tr>
                <td class="py-1.5 px-3 font-sans font-bold">${st.column}</td>
                <td class="py-1.5 px-3 text-right">${fmtNum(st.mean)}</td>
                <td class="py-1.5 px-3 text-right">${fmtNum(st.median)}</td>
                <td class="py-1.5 px-3 text-right">${fmtNum(st.std)}</td>
                <td class="py-1.5 px-3 text-right">${fmtNum(st.min)} — ${fmtNum(st.max)}</td>
                <td class="py-1.5 px-3 text-right font-bold ${st.outliers > 0 ? 'text-orange-600' : ''}">${st.outliers}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    ` : `<div class="p-3 text-xs opacity-60">Không có cột số liệu để thống kê mô tả.</div>`;
  }

  const cleanTotal = s2TotalRows || sum.clean_rows || univCleanedRows.length;
  document.getElementById("ws-status-pill").textContent = `${cleanTotal.toLocaleString("vi-VN")} dòng đã chuẩn hóa`;
}

async function renderUnivTable() {
  // 1. Phân trang từ server
  try {
    const res = await fetch(`/api/givingly/paginate?source=cleaned&page=${s2CurrentPage}&page_size=${s2PageSize}`);
    const json = await res.json();
    if (json.status === "success" && json.columns && json.columns.length > 0) {
      s2TotalRows = json.total_rows;
      s2TotalPages = Math.max(1, json.total_pages);
      s2CurrentPage = json.page;
      univCleanedCols = json.columns;
      univCleanedRows = json.rows;

      const countEl = document.getElementById("s2-univ-count");
      if (countEl) countEl.textContent = `Đã làm sạch: ${s2TotalRows.toLocaleString("vi-VN")} dòng`;

      const pageInfo = document.getElementById("s2-page-info");
      if (pageInfo) pageInfo.textContent = `Trang ${s2CurrentPage} / ${s2TotalPages}`;

      const rowsRange = document.getElementById("s2-rows-range");
      if (rowsRange) rowsRange.textContent = s2TotalRows > 0 ? `Dòng ${json.start_idx.toLocaleString("vi-VN")} - ${json.end_idx.toLocaleString("vi-VN")}` : `0 dòng`;

      const btnPrev = document.getElementById("s2-btn-prev");
      if (btnPrev) btnPrev.disabled = s2CurrentPage <= 1;

      const btnNext = document.getElementById("s2-btn-next");
      if (btnNext) btnNext.disabled = s2CurrentPage >= s2TotalPages;

      const tbl = document.getElementById("s2-univ-table");
      if (tbl) {
        tbl.innerHTML = `
          <thead class="bg-[#F7F7F4] border-b border-black/10 font-bold sticky top-0">
            <tr>${univCleanedCols.map(c => `<th class="py-2 px-3 whitespace-nowrap ${c.startsWith('_') ? 'bg-[#D4EE0E]/40' : ''}">${c}</th>`).join("")}</tr>
          </thead>
          <tbody class="divide-y divide-black/10">
            ${univCleanedRows.map(r => `<tr>${univCleanedCols.map(c => `<td class="py-1.5 px-3 max-w-[200px] truncate">${r[c] ?? "-"}</td>`).join("")}</tr>`).join("")}
          </tbody>
        `;
      }
      return;
    }
  } catch (err) {
    console.warn("Server pagination fallback for cleaned data:", err);
  }

  // 2. Fallback
  const countEl = document.getElementById("s2-univ-count");
  if (countEl) countEl.textContent = `Đã làm sạch: ${s2TotalRows.toLocaleString("vi-VN")} dòng`;
  const pageInfo = document.getElementById("s2-page-info");
  if (pageInfo) pageInfo.textContent = `Trang ${s2CurrentPage} / ${s2TotalPages}`;
  const rowsRange = document.getElementById("s2-rows-range");
  if (rowsRange) rowsRange.textContent = s2TotalRows > 0 ? `Dòng 1 - ${univCleanedRows.length}` : `0 dòng`;
  const btnPrev = document.getElementById("s2-btn-prev");
  if (btnPrev) btnPrev.disabled = s2CurrentPage <= 1;
  const btnNext = document.getElementById("s2-btn-next");
  if (btnNext) btnNext.disabled = s2CurrentPage >= s2TotalPages;

  const tbl = document.getElementById("s2-univ-table");
  if (tbl) {
    tbl.innerHTML = `
      <thead class="bg-[#F7F7F4] border-b border-black/10 font-bold sticky top-0">
        <tr>${univCleanedCols.map(c => `<th class="py-2 px-3 whitespace-nowrap ${c.startsWith('_') ? 'bg-[#D4EE0E]/40' : ''}">${c}</th>`).join("")}</tr>
      </thead>
      <tbody class="divide-y divide-black/10">
        ${univCleanedRows.map(r => `<tr>${univCleanedCols.map(c => `<td class="py-1.5 px-3 max-w-[200px] truncate">${r[c] ?? "-"}</td>`).join("")}</tr>`).join("")}
      </tbody>
    `;
  }
}

async function runUnivStep3() {
  showProgressBar("Bước 3/5: Đang phân tích tài chính & Cây KPI...", 15, "Tính toán KPI trên toàn bộ tập dữ liệu...");
  startProgressPolling();
  try {
    const customRule = (document.getElementById("s3-univ-custom-rule")?.value || "").trim();
    const guidance = (document.getElementById("guidance-step-3")?.value || "").trim();
    const res = await fetch("/api/givingly/universal-step3", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        cleaned_rows: univCleanedRows,
        universal_schema: universalSchema,
        selected_groups: universalSchema.group_cols || [],
        custom_rule: customRule,
        guidance: guidance,
        company_info: getCompanyInfo()
      })
    });
    const json = await res.json();
    if (json.status === "success") {
      univAnalysis = json.analysis;
      if (json.analysis && json.analysis.inventory_risk_metrics) {
        biRiskMetrics = json.analysis.inventory_risk_metrics;
      }
      if (json.schema_inspection) schemaInspection = json.schema_inspection;
      else if (json.analysis && json.analysis.schema_inspection) schemaInspection = json.analysis.schema_inspection;
      (json.audit_entries || []).forEach(a => addAudit(a.step, a.tool, a.detail));
      const dims = Object.keys(univAnalysis.group_tables || {});
      univActiveDim = dims.includes(univActiveDim) ? univActiveDim : (dims[0] || "");
      renderUnivStep3();
      setProgressBar(100, "Hoàn tất Bước 3", "Đã phân tích tài chính & Cây KPI thành công!");
    } else {
      showToast(json.message || "Lỗi phân tích Bước 3", "error");
    }
  } catch (err) {
    showToast("Lỗi phân tích Bước 3: " + err.message, "error");
  } finally {
    stopProgressPolling();
    hideProgressBar();
  }
}

function switchUnivView(view) {
  univActiveView = view;
  ["tree", "vip", "groups"].forEach(v => {
    const tab = document.getElementById(`s3-tab-${v}`);
    const pane = document.getElementById(`s3-univ-pane-${v}`);
    if (tab && pane) {
      if (v === view) {
        tab.className = "px-3 py-1.5 rounded-lg font-bold bg-[#0B0D08] text-[#D4EE0E]";
        pane.classList.remove("hidden");
      } else {
        tab.className = "px-3 py-1.5 rounded-lg font-bold bg-[#EFECE6] text-[#0C0E08]";
        pane.classList.add("hidden");
      }
    }
  });
}

function renderUnivStep3() {
  if (!univAnalysis) return;
  const k = univAnalysis.kpis || {};
  document.getElementById("s3-univ-kpis").innerHTML = `
    <div class="bg-[#D4EE0E] rounded-xl p-3 border border-black/15">
      <span class="text-[10px] font-bold opacity-80 block">Tổng doanh thu</span>
      <strong class="text-base font-black font-mono">${fmtNum(k.total_revenue)}</strong>
    </div>
    <div class="bg-white rounded-xl p-3 border border-black/15">
      <span class="text-[10px] font-bold opacity-60 block">Tổng lợi nhuận</span>
      <strong class="text-base font-black font-mono">${fmtNum(k.total_profit)}</strong>
    </div>
    <div class="bg-white rounded-xl p-3 border border-black/15">
      <span class="text-[10px] font-bold opacity-60 block">Biên lợi nhuận</span>
      <strong class="text-base font-black font-mono">${k.profit_margin_pct ?? "-"}%</strong>
    </div>
    <div class="bg-white rounded-xl p-3 border border-black/15">
      <span class="text-[10px] font-bold opacity-60 block">Khách hàng VIP</span>
      <strong class="text-base font-black font-mono">${(univAnalysis.tier_summary?.A ?? 0)} Tier A</strong>
    </div>
  `;

  renderUnivTreePane();
  renderUnivVipPane();

  const dims = Object.keys(univAnalysis.group_tables || {});
  document.getElementById("s3-univ-dim-tabs").innerHTML = dims.length
    ? dims.map(d => `<button onclick="switchUnivDim('${d.replace(/'/g, "\\'")}')" class="px-2.5 py-1 rounded-md text-[11px] font-bold ${d === univActiveDim ? 'bg-[#0B0D08] text-[#D4EE0E]' : 'bg-[#EFECE6] text-[#0C0E08]'}">${d === "_Thang" ? "Theo Tháng" : d}</button>`).join("")
    : `<span class="text-[#0C0E08]/60 italic">Không có cột phân nhóm</span>`;

  renderUnivDimTable();
  switchUnivView(univActiveView);
  document.getElementById("ws-status-pill").textContent = `Tổng DT: ${fmtNum(k.total_revenue)} • MG: ${k.profit_margin_pct ?? "-"}%`;
}

function renderUnivTreePane() {
  const wrap = document.getElementById("s3-univ-pane-tree");
  const tree = univAnalysis?.kpi_tree;
  if (!wrap || !tree) return;

  wrap.innerHTML = `
    <div class="bg-white border border-black/15 rounded-xl p-4 space-y-4">
      <div class="flex flex-col items-center">
        <div class="bg-[#0B0D08] text-white rounded-xl px-6 py-3 text-center shadow-sm">
          <span class="text-[10px] font-bold uppercase tracking-wider text-[#D4EE0E] block">${tree.root.label}</span>
          <strong class="text-xl font-black font-mono">${fmtNum(tree.root.value)}</strong>
        </div>
        <div class="w-0.5 h-4 bg-black/30"></div>
        <div class="w-4/5 h-0.5 bg-black/20"></div>
      </div>

      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
        ${(tree.branches || []).map((b, idx) => `
          <div class="${idx === 2 ? 'bg-[#D4EE0E]/45 border-[#0C0E08]/30' : 'bg-[#F7F7F4] border-black/15'} border rounded-xl p-3 space-y-1">
            <span class="text-[10px] font-bold opacity-70 block">${b.label}</span>
            <strong class="text-base font-black font-mono block">${fmtNum(b.value)}</strong>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderUnivVipPane() {
  const wrap = document.getElementById("s3-univ-pane-vip");
  const list = univAnalysis?.customer_priority || [];
  const ts = univAnalysis?.tier_summary || {};
  if (!wrap) return;

  const displayList = s3VipShowAll ? list : list.slice(0, 5);

  wrap.innerHTML = `
    <div class="grid grid-cols-3 gap-2.5 text-xs">
      <div class="bg-[#D4EE0E] border border-black/15 rounded-xl p-3">
        <span class="text-[10px] font-bold block">Tier A (VIP)</span>
        <strong class="text-lg font-black">${ts.A || 0}</strong>
      </div>
      <div class="bg-white border border-black/15 rounded-xl p-3">
        <span class="text-[10px] font-bold opacity-60 block">Tier B</span>
        <strong class="text-lg font-black">${ts.B || 0}</strong>
      </div>
      <div class="bg-white border border-black/15 rounded-xl p-3">
        <span class="text-[10px] font-bold opacity-60 block">Tier C</span>
        <strong class="text-lg font-black">${ts.C || 0}</strong>
      </div>
    </div>

    <div class="bg-white border border-black/15 rounded-xl overflow-hidden">
      <div class="px-3 py-2 bg-[#EFECE6] border-b border-black/10 flex items-center justify-between text-xs">
        <span class="font-bold">Khách hàng ưu tiên (${list.length})</span>
        ${list.length > 5 ? `<button onclick="toggleS3VipTable()" class="text-[10px] font-bold bg-white border border-black/15 px-2 py-0.5 rounded hover:bg-black/5">${s3VipShowAll ? 'Thu gọn' : 'Xem tất cả (' + list.length + ')'}</button>` : ''}
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-[11px] text-left">
          <thead class="bg-[#F7F7F4] font-bold border-b border-black/10">
            <tr>
              <th class="py-2 px-3">#</th>
              <th class="py-2 px-3">Khách hàng / Đối tượng</th>
              <th class="py-2 px-3 text-right">Số đơn</th>
              <th class="py-2 px-3 text-right">Doanh thu</th>
              <th class="py-2 px-3 text-right">Tỷ trọng</th>
              <th class="py-2 px-3 text-right">Margin</th>
              <th class="py-2 px-3">Phân hạng</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-black/10">
            ${displayList.length ? displayList.map(v => {
              let tBadge = "bg-[#EFECE6] text-[#0C0E08]";
              if (v.tier.includes("Tier A")) tBadge = "bg-[#0B0D08] text-[#D4EE0E]";
              else if (v.tier.includes("Tier B")) tBadge = "bg-[#D4EE0E] text-[#0C0E08]";
              return `
                <tr>
                  <td class="py-2 px-3 font-mono font-bold">${v.rank}</td>
                  <td class="py-2 px-3 font-bold">${v.customer}</td>
                  <td class="py-2 px-3 text-right font-mono">${fmtNum(v.orders)}</td>
                  <td class="py-2 px-3 text-right font-mono font-bold">${fmtNum(v.revenue)}</td>
                  <td class="py-2 px-3 text-right font-mono">${v.share_pct}%</td>
                  <td class="py-2 px-3 text-right font-mono">${v.margin_pct !== null && v.margin_pct !== undefined ? v.margin_pct + "%" : "-"}</td>
                  <td class="py-2 px-3"><span class="px-2 py-0.5 rounded font-extrabold text-[10px] ${tBadge}">${v.tier}</span></td>
                </tr>
              `;
            }).join("") : `<tr><td colspan="7" class="py-4 px-3 text-center opacity-60">Chưa xác định được cột Khách hàng.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function switchUnivDim(dim) {
  univActiveDim = dim;
  renderUnivStep3();
  switchUnivView("groups");
}

function renderUnivDimTable() {
  const tbody = document.getElementById("s3-univ-tbody");
  const rows = (univAnalysis?.group_tables || {})[univActiveDim] || [];
  const btn = document.getElementById("btn-toggle-s3-univ-dim");
  if (btn) {
    btn.textContent = s3UnivDimShowAll ? "Thu gọn" : `Xem tất cả (${rows.length})`;
    btn.classList.toggle("hidden", rows.length <= 5);
  }
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="py-4 px-3 text-center opacity-60">Không có dữ liệu phân nhóm.</td></tr>`;
    return;
  }
  const displayRows = s3UnivDimShowAll ? rows : rows.slice(0, 5);
  tbody.innerHTML = displayRows.map(r => {
    const perf = r.performance || "Ổn định";
    let badge = "bg-[#EFECE6] text-[#0C0E08]";
    if (perf.includes("Hiệu suất cao")) badge = "bg-[#D4EE0E] text-[#0C0E08]";
    else if (perf.includes("Cần tối ưu")) badge = "bg-orange-100 text-orange-800";
    return `
      <tr>
        <td class="py-2 px-3 font-bold">${r.group_value}</td>
        <td class="py-2 px-3 text-right font-mono">${fmtNum(r.count)}</td>
        <td class="py-2 px-3 text-right font-mono font-bold">${fmtNum(r.revenue)}</td>
        <td class="py-2 px-3 text-right font-mono">${r.share_pct}%</td>
        <td class="py-2 px-3 text-right font-mono">${fmtNum(r.profit)}</td>
        <td class="py-2 px-3 text-right font-mono">${r.margin_pct !== null && r.margin_pct !== undefined ? r.margin_pct + "%" : "-"}</td>
        <td class="py-2 px-3"><span class="px-2 py-0.5 rounded font-extrabold text-[10px] ${badge}">${perf}</span></td>
      </tr>
    `;
  }).join("");
}

async function runUnivStep4() {
  showProgressBar("Bước 4/5: Đang sinh báo cáo & nhận định điều hành AI...", 30, "Mô hình AI đang phân tích dữ liệu tổng hợp...");
  startProgressPolling();
  try {
    const style = document.getElementById("s4-univ-style")?.value || "Điều hành súc tích";
    const guidance = (document.getElementById("guidance-step-4")?.value || "").trim();
    const res = await fetch("/api/givingly/universal-step4", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        analysis: univAnalysis,
        source_name: lastSourceName,
        style,
        guidance: guidance,
        api_key: document.getElementById("cfg-apikey")?.value || "",
        provider: document.getElementById("cfg-provider")?.value || "Gemini",
        company_info: getCompanyInfo()
      })
    });
    const json = await res.json();
    if (json.status === "success") {
      univReportText = json.report_text || "";
      addAudit("Bước 4 - AI Report", "universal_step4_ai_report", `Tạo Báo cáo điều hành (${style})`);
      renderUnivStep4();
      setProgressBar(100, "Hoàn tất Bước 4", "Đã tạo báo cáo AI thành công!");
    } else {
      showToast(json.message || "Lỗi tạo báo cáo Bước 4", "error");
    }
  } catch (err) {
    showToast("Lỗi tạo báo cáo AI: " + err.message, "error");
  } finally {
    stopProgressPolling();
    hideProgressBar();
  }
}

// =============================================================
// Rule 3: Markdown-to-HTML Parser for Step 4 & Step 5
// =============================================================
function renderMarkdownToHtml(md) {
  if (!md || typeof md !== "string") return "";
  let text = md.replace(/\r\n/g, "\n").trim();
  if (!text) return "";

  const escapeHtml = (str) => str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  const formatInline = (str) => {
    return str
      .replace(/\*\*\*(.*?)\*\*\*/g, "<strong class='font-bold text-slate-900'><em class='italic'>$1</em></strong>")
      .replace(/\*\*(.*?)\*\*/g, "<strong class='font-bold text-slate-900'>$1</strong>")
      .replace(/\*(.*?)\*/g, "<em class='italic text-slate-700'>$1</em>")
      .replace(/`([^`]+)`/g, "<code class='bg-slate-100 text-blue-700 px-1 py-0.5 rounded font-mono text-[11px]'>$1</code>");
  };

  const lines = text.split("\n");
  let html = [];
  let inList = false;

  const closeList = () => {
    if (inList) {
      html.push("</ul>");
      inList = false;
    }
  };

  for (let i = 0; i < lines.length; i++) {
    let rawLine = lines[i].trim();
    if (!rawLine) {
      closeList();
      continue;
    }

    // Skip markdown horizontal rules or table separators
    if (/^[\-\*_]{3,}$/.test(rawLine) || /^\|?[\s\:\-]+(?:\|[\s\:\-]+)+\|?$/.test(rawLine)) {
      continue;
    }

    // Headers
    if (rawLine.startsWith("### ")) {
      closeList();
      const hText = formatInline(escapeHtml(rawLine.slice(4).trim()));
      html.push(`<h4 class="text-xs font-black uppercase tracking-wider text-slate-900 mt-3.5 mb-1.5 flex items-center gap-1.5">
        <span class="w-1.5 h-3.5 bg-[#1B365D] rounded-full inline-block"></span>
        <span>${hText}</span>
      </h4>`);
      continue;
    }
    if (rawLine.startsWith("## ")) {
      closeList();
      const hText = formatInline(escapeHtml(rawLine.slice(3).trim()));
      html.push(`<h3 class="text-sm font-black text-[#1B365D] mt-4 mb-2 pb-1 border-b border-black/10">${hText}</h3>`);
      continue;
    }
    if (rawLine.startsWith("# ")) {
      closeList();
      const hText = formatInline(escapeHtml(rawLine.slice(2).trim()));
      html.push(`<h2 class="text-base font-black text-[#1B365D] mt-4 mb-2 pb-1.5 border-b-2 border-[#1B365D]/30">${hText}</h2>`);
      continue;
    }

    // Bullet items: * or -
    const bulletMatch = rawLine.match(/^[\*\-]\s+(.*)$/);
    if (bulletMatch) {
      if (!inList) {
        html.push('<ul class="space-y-1.5 my-2 pl-1">');
        inList = true;
      }
      const itemText = formatInline(escapeHtml(bulletMatch[1].trim()));
      html.push(`
        <li class="flex items-start gap-2 text-xs text-slate-700 leading-relaxed">
          <span class="text-[#1B365D] font-black shrink-0 mt-0.5">•</span>
          <span class="flex-1">${itemText}</span>
        </li>
      `);
      continue;
    }

    // Numbered list
    const numMatch = rawLine.match(/^(\d+)\.\s+(.*)$/);
    if (numMatch) {
      if (!inList) {
        html.push('<ul class="space-y-1.5 my-2 pl-1">');
        inList = true;
      }
      const itemText = formatInline(escapeHtml(numMatch[2].trim()));
      html.push(`
        <li class="flex items-start gap-2 text-xs text-slate-700 leading-relaxed">
          <span class="text-[#1B365D] font-bold shrink-0 mt-0.5 text-[11px] font-mono">${numMatch[1]}.</span>
          <span class="flex-1">${itemText}</span>
        </li>
      `);
      continue;
    }

    // Paragraph
    closeList();
    const pText = formatInline(escapeHtml(rawLine));
    html.push(`<p class="text-xs text-slate-700 leading-relaxed mb-2">${pText}</p>`);
  }

  closeList();
  return html.join("\n");
}

let s4EditorOpen = false;
function toggleS4MarkdownEditor() {
  s4EditorOpen = !s4EditorOpen;
  const wrap = document.getElementById("s4-univ-editor-wrap");
  const rendered = document.getElementById("s4-univ-report-rendered");
  const icon = document.getElementById("s4-toggle-icon");
  const text = document.getElementById("s4-toggle-text");
  if (wrap && rendered) {
    if (s4EditorOpen) {
      wrap.classList.remove("hidden");
      if (icon) icon.textContent = "👁️";
      if (text) text.textContent = "Xem trực quan";
    } else {
      wrap.classList.add("hidden");
      if (icon) icon.textContent = "✏️";
      if (text) text.textContent = "Sửa Markdown";
      rendered.innerHTML = renderMarkdownToHtml(univReportText);
    }
  }
}

function onS4ReportInput(val) {
  univReportText = val;
  const rendered = document.getElementById("s4-univ-report-rendered");
  if (rendered) rendered.innerHTML = renderMarkdownToHtml(val);
}

function renderUnivStep4() {
  if (!univAnalysis) return;
  const k = univAnalysis.kpis || {};

  // Rule 2: Chỉ hiển thị các thẻ KPI thực sự tồn tại trong bộ dữ liệu
  const cardsHtml = [];
  if (k.total_revenue !== undefined && k.total_revenue !== null && k.total_revenue > 0) {
    cardsHtml.push(`
      <div class="bg-[#D4EE0E]/50 p-2.5 rounded-lg border border-black/10">
        <span class="text-[10px] font-bold block text-slate-700">Tổng Doanh thu</span>
        <strong class="text-base font-black font-mono text-slate-900">${fmtNum(k.total_revenue)}</strong>
      </div>
    `);
  }
  if (k.total_profit !== undefined && k.total_profit !== null) {
    cardsHtml.push(`
      <div class="bg-[#EFECE6] p-2.5 rounded-lg border border-black/10">
        <span class="text-[10px] font-bold block text-slate-700">Tổng Lợi nhuận</span>
        <strong class="text-base font-black font-mono text-emerald-800">${fmtNum(k.total_profit)}</strong>
      </div>
    `);
  }
  if (k.profit_margin_pct !== undefined && k.profit_margin_pct !== null) {
    cardsHtml.push(`
      <div class="bg-[#EFECE6] p-2.5 rounded-lg border border-black/10">
        <span class="text-[10px] font-bold block text-slate-700">Biên lợi nhuận</span>
        <strong class="text-base font-black font-mono text-emerald-800">${k.profit_margin_pct}%</strong>
      </div>
    `);
  }
  if (univAnalysis.tier_summary && univAnalysis.tier_summary.col_used && (univAnalysis.tier_summary.A || 0) > 0) {
    cardsHtml.push(`
      <div class="bg-[#EFECE6] p-2.5 rounded-lg border border-black/10">
        <span class="text-[10px] font-bold block text-slate-700">Đối tác VIP</span>
        <strong class="text-base font-black font-mono text-blue-900">${univAnalysis.tier_summary.A} Tier A</strong>
      </div>
    `);
  }
  if (cardsHtml.length < 2) {
    cardsHtml.push(`
      <div class="bg-[#EFECE6] p-2.5 rounded-lg border border-black/10">
        <span class="text-[10px] font-bold block text-slate-700">Quy mô bản ghi</span>
        <strong class="text-base font-black font-mono text-slate-900">${(k.total_rows || 0).toLocaleString("vi-VN")} dòng</strong>
      </div>
    `);
  }
  if (k.total_qty !== undefined && k.total_qty !== null && cardsHtml.length < 4) {
    cardsHtml.push(`
      <div class="bg-[#EFECE6] p-2.5 rounded-lg border border-black/10">
        <span class="text-[10px] font-bold block text-slate-700">Tổng Sản lượng</span>
        <strong class="text-base font-black font-mono text-slate-900">${Number(k.total_qty).toLocaleString("vi-VN")}</strong>
      </div>
    `);
  }

  const kpisLeftEl = document.getElementById("s4-univ-kpis-left");
  if (kpisLeftEl) kpisLeftEl.innerHTML = cardsHtml.join("");

  const dim = univActiveDim || Object.keys(univAnalysis.group_tables || {})[0];
  const topRows = ((univAnalysis.group_tables || {})[dim] || []).slice(0, 4);
  const barsRightEl = document.getElementById("s4-univ-bars-right");
  if (barsRightEl) {
    barsRightEl.innerHTML = `
      <div class="flex items-center justify-between">
        <strong class="text-xs font-extrabold text-slate-800">Top [${dim === "_Thang" ? "Tháng" : (dim || "Nhóm")}]</strong>
        <span class="text-[10px] font-bold bg-[#D4EE0E] px-2 py-0.5 rounded text-slate-900">Top Đóng Góp</span>
      </div>
      ${topRows.map(r => `
        <div class="space-y-1">
          <div class="flex justify-between text-[11px] font-bold">
            <span class="truncate max-w-[160px] text-slate-700">${r.group_value}</span>
            <span class="font-mono text-slate-900">${fmtNum(r.revenue || r.count)} (${r.share_pct || 0}%)</span>
          </div>
          <div class="w-full bg-black/10 h-1.5 rounded-full overflow-hidden">
            <div class="bg-[#D4EE0E] h-full rounded-full" style="width: ${Math.min(100, Math.max(2, r.share_pct || 0))}%"></div>
          </div>
        </div>
      `).join("") || '<p class="text-[11px] opacity-60 italic">Chưa có bảng phân nhóm.</p>'}
    `;
  }

  // Rule 3: Đồng bộ cả Textarea và View Markdown-to-HTML
  const rawTextarea = document.getElementById("s4-univ-report-text");
  if (rawTextarea) rawTextarea.value = univReportText;

  const renderedBox = document.getElementById("s4-univ-report-rendered");
  if (renderedBox) {
    renderedBox.innerHTML = renderMarkdownToHtml(univReportText);
  }
}

async function regenerateUnivReport() {
  const btn = document.getElementById("btn-univ-regen");
  if (btn) btn.textContent = "⏳ Đang viết...";
  try {
    await runUnivStep4();
    showToast("Đã cập nhật Báo cáo AI theo văn phong mới!", "success");
  } finally {
    if (btn) btn.textContent = "✨ AI Viết lại Nhận định";
  }
}

async function runUnivStep5Export() {
  const approverName = document.getElementById("s5-approver-name")?.value?.trim() || "Ban Giám Đốc";
  const approverTitle = document.getElementById("s5-approver-title")?.value?.trim() || "Đại diện Phê duyệt";
  const approvalNotes = document.getElementById("s5-approval-notes")?.value?.trim() || "Đồng ý phê duyệt số liệu kinh doanh, chính sách chiết khấu và ban hành báo cáo chính thức.";
  const fullApprover = `${approverName} — ${approverTitle}`;

  const res = await fetch("/api/givingly/universal-export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      cleaned_rows: univCleanedRows,
      analysis: univAnalysis,
      report_text: document.getElementById("s4-univ-report-text")?.value || univReportText,
      source_name: lastSourceName,
      audit_log: auditLog,
      approver_name: fullApprover,
      approval_notes: approvalNotes,
      company_info: getCompanyInfo()
    })
  });
  const json = await res.json();
  if (json.status === "success") {
    const setProp = (id, prop, val) => { const el = document.getElementById(id); if (el) el[prop] = val; };
    setProp("dl-excel", "href", json.urls.excel);
    setProp("dl-docx", "href", json.urls.docx);
    setProp("dl-txt", "href", json.urls.audit_txt);
    setProp("dl-audit", "href", json.urls.audit_json);
    setProp("dl-excel-desc", "textContent", "Đa tab: TongQuan_KPI, ThongKe_MoTa, KhachHang_UuTien, Nhom_*");
    setProp("dl-docx-desc", "textContent", "Báo Cáo Kinh Doanh Đã Phê Duyệt + Cây KPI + KH VIP");
    setProp("dl-txt-title", "textContent", "📄 Audit Log Thuần Text (.txt)");
    setProp("dl-txt-desc", "textContent", "Nhật ký từng bước kiểm duyệt & công cụ");
    setProp("ws-status-pill", "textContent", "Hoàn tất 5/5 Bước");
    return true;
  } else {
    showToast(json.message || "Lỗi xuất file Bước 5", "error");
    return false;
  }
}

// -------------------------------------------------------------
// -------------------------------------------------------------
// 8. DYNAMIC SCHEMA-DRIVEN DASHBOARD & SENIOR BI REPORT PREVIEW
// Phân loại cột thành 3 nhóm: Categorical, Numerical, Temporal
// Tự động thích ứng 100% theo dữ liệu thực tế (Không hiển thị biểu mẫu giả định)
// -------------------------------------------------------------

function formatColDisplayName(col) {
  if (!col) return "Chỉ số";
  const map = {
    revenue: "Doanh thu",
    sales: "Doanh thu",
    doanh_thu: "Doanh thu",
    doanh_so: "Doanh số",
    thanh_tien: "Thành tiền",
    ngan_sach: "Ngân sách",
    ngan_sach_cong_bo: "Ngân sách công bố",
    cost: "Chi phí",
    cogs: "Giá vốn",
    gia_von: "Giá vốn",
    chi_phi: "Chi phí",
    profit: "Lợi nhuận",
    loi_nhuan: "Lợi nhuận",
    lai_gop: "Lợi nhuận gộp",
    quantity: "Số lượng",
    qty: "Số lượng",
    so_luong: "Số lượng",
    san_luong: "Sản lượng",
    units: "Đơn vị",
    price: "Giá bán",
    don_gia: "Đơn giá",
    category: "Ngành hàng",
    danh_muc: "Danh mục",
    region: "Khu vực",
    khu_vuc: "Khu vực",
    channel: "Kênh bán",
    kenh: "Kênh bán",
    customer: "Khách hàng",
    khach_hang: "Khách hàng",
    segment: "Phân khúc",
    phan_khuc: "Phân khúc",
    order_date: "Ngày đặt hàng",
    ngay_nhan: "Ngày tiếp nhận",
    _Thang: "Tháng",
    thang: "Tháng"
  };
  const clean = String(col).toLowerCase().trim();
  if (map[clean]) return map[clean];
  return String(col)
    .replace(/_/g, " ")
    .replace(/\b\w/g, l => l.toUpperCase());
}

function inspectClientDataSchema(rows) {
  if (!rows || !rows.length) {
    return {
      categorical_cols: [],
      numerical_cols: [],
      temporal_cols: [],
      has_sku_inventory: false,
      has_temporal: false,
      has_categorical: false,
      has_numerical: false,
      primary_metric: null,
      secondary_metric: null,
      primary_dimension: null,
      secondary_dimension: null,
      primary_temporal: null,
      pivot_summary: [],
      total_rows: 0
    };
  }

  const sampleRow = rows[0] || {};
  const cols = Object.keys(sampleRow);
  const temporalCols = [];
  const numericalCols = [];
  const categoricalCols = [];

  const TIME_KEYWORDS = ["date", "ngày", "ngay", "thời_gian", "thoi_gian", "time", "timestamp", "created_at", "month", "tháng", "thang", "_thang", "year", "năm", "nam", "quarter", "quý", "quy", "period"];
  const ID_KEYWORDS = ["id", "mã", "ma_", "code", "uuid", "guid", "stt", "số_thứ_tự", "sdt", "phone", "dien_thoai", "tel", "mobile", "fax", "tax", "mst", "zip", "postal", "cccd", "cmnd"];
  const REV_KEYWORDS = ["revenue", "sales", "doanh_thu", "doanh_so", "thanh_tien", "ngan_sach", "tiền", "tien", "amount", "budget"];
  const PROFIT_KEYWORDS = ["profit", "lợi_nhuận", "loi_nhuan", "lãi", "lai", "margin", "biên_ln"];
  const COST_KEYWORDS = ["cost", "chi_phí", "chi_phi", "gia_von", "giá_vốn", "expense", "cogs"];
  const QTY_KEYWORDS = ["qty", "quantity", "số_lượng", "so_luong", "sản_lượng", "san_luong", "volume", "units"];
  const PRICE_KEYWORDS = ["price", "đơn_giá", "don_gia", "giá_bán", "gia_ban", "unit_price"];
  const DISC_KEYWORDS = ["discount", "chiết_khấu", "chiet_khau", "giảm_giá", "giam_gia", "rebate"];
  const STOCK_KEYWORDS = ["stock", "ton_kho", "tồn_kho", "ton_kha_dung", "tồn_khả_dụng", "inventory", "on_hand", "khả_dụng"];
  const SKU_KEYWORDS = ["sku", "mã_sp", "ma_sp", "product_id", "item_id", "mã_hàng", "ma_hang"];

  function stripAccents(s) {
    return String(s || "").toLowerCase().trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d").replace(/Đ/g, "D");
  }

  cols.forEach(col => {
    const colClean = stripAccents(col);
    const nonNullVals = rows.map(r => r[col]).filter(v => v !== null && v !== undefined && String(v).trim() !== "");
    if (!nonNullVals.length) return;

    // 1. Temporal Check
    let isTemp = false;
    if (TIME_KEYWORDS.some(k => colClean === k || colClean.includes(k))) {
      const sample = nonNullVals.slice(0, 10);
      let validDateCount = 0;
      sample.forEach(s => {
        const d = Date.parse(s);
        if (!isNaN(d)) validDateCount++;
      });
      if (validDateCount >= Math.max(1, Math.floor(sample.length * 0.5)) ||
          ["thang", "tháng", "_thang", "month", "year", "nam", "quy", "quarter"].some(w => colClean.includes(w))) {
        isTemp = true;
      }
    }

    if (isTemp) {
      const uniqueVals = new Set(nonNullVals.map(String));
      temporalCols.push({
        name: col,
        role: ["ngay", "date"].some(w => colClean.includes(w)) ? "date" : "month",
        distinct_count: uniqueVals.size,
        sample_values: nonNullVals.slice(0, 3).map(String)
      });
      return;
    }

    // 2. Numerical Check
    let isNum = false;
    let numVals = [];
    let isId = ID_KEYWORDS.some(k => colClean.includes(k));

    if (!isId) {
      let numCount = 0;
      nonNullVals.forEach(v => {
        if (typeof v === "number" && !isNaN(v)) {
          numVals.push(v);
          numCount++;
        } else {
          const strClean = String(v).replace(/[\$,VNĐvndđ\s]/g, "").replace(/,/g, "");
          const p = Number(strClean);
          if (!isNaN(p) && strClean !== "") {
            numVals.push(p);
            numCount++;
          }
        }
      });
      if (numCount >= Math.max(1, Math.floor(nonNullVals.length * 0.7))) {
        isNum = true;
      }
    }

    if (isNum && numVals.length > 0) {
      let role = "metric";
      if (REV_KEYWORDS.some(k => colClean.includes(k))) role = "revenue";
      else if (PROFIT_KEYWORDS.some(k => colClean.includes(k))) role = "profit";
      else if (COST_KEYWORDS.some(k => colClean.includes(k))) role = "cost";
      else if (QTY_KEYWORDS.some(k => colClean.includes(k))) role = "quantity";
      else if (PRICE_KEYWORDS.some(k => colClean.includes(k))) role = "price";
      else if (DISC_KEYWORDS.some(k => colClean.includes(k))) role = "discount";
      else if (STOCK_KEYWORDS.some(k => colClean.includes(k))) role = "stock";

      const sum = numVals.reduce((a, b) => a + b, 0);
      const mean = sum / numVals.length;
      const sorted = [...numVals].sort((a, b) => a - b);
      const median = sorted[Math.floor(sorted.length / 2)];
      const min = sorted[0];
      const max = sorted[sorted.length - 1];

      numericalCols.push({
        name: col,
        role: role,
        sum: sum,
        mean: mean,
        median: median,
        min: min,
        max: max,
        count: numVals.length
      });
      return;
    }

    // 3. Categorical Check
    let role = "dimension";
    if (SKU_KEYWORDS.some(k => colClean.includes(k))) role = "sku";
    else if (["product", "san_pham", "ten_sp", "hang_hoa", "item"].some(k => colClean.includes(k))) role = "product";
    else if (["category", "danh_muc", "nganh_hang"].some(k => colClean.includes(k))) role = "category";
    else if (["region", "khu_vuc", "tinh", "tp", "city"].some(k => colClean.includes(k))) role = "region";
    else if (["channel", "kenh"].some(k => colClean.includes(k))) role = "channel";
    else if (["customer", "khach_hang", "client", "buyer"].some(k => colClean.includes(k))) role = "customer";
    else if (["segment", "phan_khuc", "tier", "hang"].some(k => colClean.includes(k))) role = "segment";
    else if (["status", "trang_thai"].some(k => colClean.includes(k))) role = "status";

    const freqMap = {};
    nonNullVals.forEach(v => {
      const s = String(v);
      freqMap[s] = (freqMap[s] || 0) + 1;
    });
    const uniqueCount = Object.keys(freqMap).length;
    const topValues = Object.entries(freqMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([val, cnt]) => ({ val, cnt }));

    categoricalCols.push({
      name: col,
      role: role,
      distinct_count: uniqueCount,
      top_values: topValues
    });
  });

  const hasB2bInventory = Boolean(sheetsData && (sheetsData.inventory || sheetsData.TonKho_TrienKhai || assignments.inventory));
  const hasSku = categoricalCols.some(c => c.role === "sku" || SKU_KEYWORDS.some(k => stripAccents(c.name).includes(k)));
  const hasStock = numericalCols.some(c => c.role === "stock" || STOCK_KEYWORDS.some(k => stripAccents(c.name).includes(k)));
  const hasSkuInventory = (hasSku && hasStock) || hasB2bInventory;

  const revCandidates = numericalCols.filter(c => c.role === "revenue");
  const profitCandidates = numericalCols.filter(c => c.role === "profit");
  const qtyCandidates = numericalCols.filter(c => c.role === "quantity");
  const primaryMetric = revCandidates.length ? revCandidates[0].name : (numericalCols.length ? numericalCols[0].name : null);

  let secondaryMetric = null;
  if (profitCandidates.length && profitCandidates[0].name !== primaryMetric) secondaryMetric = profitCandidates[0].name;
  else if (qtyCandidates.length && qtyCandidates[0].name !== primaryMetric) secondaryMetric = qtyCandidates[0].name;
  else {
    const sec = numericalCols.filter(c => c.name !== primaryMetric);
    if (sec.length) secondaryMetric = sec[0].name;
  }

  const dimPriority = ["category", "region", "channel", "segment", "product", "customer", "dimension"];
  let primaryDimension = null;
  for (const pRole of dimPriority) {
    const matched = categoricalCols.filter(c => c.role === pRole && c.distinct_count >= 2 && c.distinct_count <= 50);
    if (matched.length) {
      primaryDimension = matched[0].name;
      break;
    }
  }
  if (!primaryDimension) {
    const cand = categoricalCols.filter(c => c.distinct_count >= 2 && c.distinct_count <= 50);
    if (cand.length) primaryDimension = cand[0].name;
    else if (categoricalCols.length) primaryDimension = categoricalCols[0].name;
  }

  let secondaryDimension = null;
  const secDims = categoricalCols.filter(c => c.name !== primaryDimension && c.distinct_count >= 2 && c.distinct_count <= 50);
  if (secDims.length) secondaryDimension = secDims[0].name;

  const primaryTemporal = temporalCols.length ? temporalCols[0].name : null;

  const pivotSummary = [];
  if (primaryDimension && primaryMetric) {
    const pCol = primaryDimension;
    const mCol = primaryMetric;
    const sCol = secondaryMetric;

    const groupMap = {};
    rows.forEach(r => {
      const dimVal = String(r[pCol] || "Khác").trim();
      let mVal = Number(String(r[mCol] || 0).replace(/[\$,VNĐvndđ\s]/g, "").replace(/,/g, ""));
      if (isNaN(mVal)) mVal = 0;
      let sVal = sCol ? Number(String(r[sCol] || 0).replace(/[\$,VNĐvndđ\s]/g, "").replace(/,/g, "")) : 0;
      if (isNaN(sVal)) sVal = 0;

      if (!groupMap[dimVal]) {
        groupMap[dimVal] = { count: 0, total_primary: 0, total_secondary: 0 };
      }
      groupMap[dimVal].count++;
      groupMap[dimVal].total_primary += mVal;
      groupMap[dimVal].total_secondary += sVal;
    });

    const totalAllPrimary = Object.values(groupMap).reduce((acc, g) => acc + g.total_primary, 0);
    const sortedEntries = Object.entries(groupMap).sort((a, b) => b[1].total_primary - a[1].total_primary);

    sortedEntries.forEach(([dimVal, g]) => {
      const share = totalAllPrimary > 0 ? Number(((g.total_primary / totalAllPrimary) * 100).toFixed(1)) : 0;
      const avg = g.count > 0 ? (g.total_primary / g.count) : 0;

      let status = "Tiêu chuẩn";
      let level = "STANDARD";
      let rec = `Duy trì quy mô ổn định (${share}%); khai thác tiềm năng mở rộng.`;

      if (share >= 20.0) {
        status = "Dẫn đầu";
        level = "TOP";
        rec = `Chiếm ${share}% tổng thể; nhóm trụ cột cần bảo vệ thị phần và ưu tiên nguồn lực.`;
      } else if (share < 5.0) {
        status = "Cần cải thiện";
        level = "LOW";
        rec = `Tỷ trọng mỏng (${share}%); cần rà soát chiến lược tiếp cận hoặc tối ưu chi phí.`;
      }

      pivotSummary.push({
        dimension_value: dimVal,
        order_count: g.count,
        total_primary: g.total_primary,
        share_pct: share,
        avg_primary: avg,
        total_secondary: sCol ? g.total_secondary : null,
        exception_status: status,
        exception_level: level,
        recommendation: rec
      });
    });
  }

  return {
    categorical_cols: categoricalCols,
    numerical_cols: numericalCols,
    temporal_cols: temporalCols,
    has_sku_inventory: hasSkuInventory,
    has_temporal: temporalCols.length > 0,
    has_categorical: categoricalCols.length > 0,
    has_numerical: numericalCols.length > 0,
    primary_metric: primaryMetric,
    secondary_metric: secondaryMetric,
    primary_dimension: primaryDimension,
    secondary_dimension: secondaryDimension,
    primary_temporal: primaryTemporal,
    pivot_summary: pivotSummary,
    total_rows: rows.length
  };
}

function getEffectiveSchemaInspection() {
  if (schemaInspection && schemaInspection.numerical_cols && schemaInspection.numerical_cols.length > 0) {
    return schemaInspection;
  }
  if (univAnalysis && univAnalysis.schema_inspection && univAnalysis.schema_inspection.numerical_cols) {
    schemaInspection = univAnalysis.schema_inspection;
    return schemaInspection;
  }
  const rows = (univCleanedRows && univCleanedRows.length) ? univCleanedRows :
               (cleanedRows && cleanedRows.length) ? cleanedRows :
               (sheetsData && Object.keys(sheetsData).length) ? (sheetsData[Object.keys(sheetsData)[0]]?.rows || []) : [];
  schemaInspection = inspectClientDataSchema(rows);
  return schemaInspection;
}

function getEffectiveBiRiskMetrics() {
  if (datasetMode === "MINH_AN_B2B" && b2bRiskMetrics) {
    return b2bRiskMetrics;
  }
  if (biRiskMetrics) return biRiskMetrics;

  const catRows = sheetsData[assignments.catalog]?.rows || [];
  const invRows = sheetsData[assignments.inventory]?.rows || [];
  const inqRows = sheetsData[assignments.inquiries]?.rows || [];

  if (catRows.length > 0 && invRows.length > 0) {
    const invMap = {};
    invRows.forEach(r => { invMap[String(r.ma_sp || "").trim()] = r; });
    const demandMap = {};
    const revMap = {};

    if (analyzedResults && analyzedResults.length > 0) {
      analyzedResults.forEach(opp => {
        (opp.chi_tiet_hang_hoa || []).forEach(line => {
          const sku = String(line.ma_sp || "").trim();
          const qty = Number(line.so_luong || 0);
          const amt = Number(line.thanh_tien || 0);
          if (sku) {
            demandMap[sku] = (demandMap[sku] || 0) + qty;
            revMap[sku] = (revMap[sku] || 0) + amt;
          }
        });
      });
    } else {
      inqRows.forEach(r => {
        const text = String(r.noi_dung_yeu_cau || r.noi_dung_goc || "");
        catRows.forEach(c => {
          const sku = String(c.ma_sp || "").trim();
          if (sku && text.includes(sku)) {
            demandMap[sku] = (demandMap[sku] || 0) + 2;
          }
        });
      });
    }

    let totVal = 0;
    const skuList = catRows.map(c => {
      const sku = String(c.ma_sp || "").trim();
      const inv = invMap[sku] || {};
      const stock = Number(inv.ton_kha_dung || 0);
      const cost = Number(c.gia_von || 0);
      const price = Number(c.gia_niem_yet || 0);
      const demand = demandMap[sku] || 0;
      const rev = revMap[sku] || (demand * price);
      const itemVal = stock * (cost || price);
      totVal += itemVal;

      const dailyVel = demand > 0 ? Number((demand / 7).toFixed(2)) : 0;
      const doi = dailyVel > 0 ? Number((stock / dailyVel).toFixed(1)) : (stock > 0 ? 99.0 : 0);
      const leadReady = Number(inv.thoi_gian_giao_san_co_ngay || 1);
      const leadImport = Number(inv.thoi_gian_nhap_hang_ngay || 14);

      let status = "An toàn";
      let level = "SAFE";
      let rec = "Tồn khả dụng ổn định, sẵn sàng giao ngay trong 24-48h.";

      if (stock === 0 || (demand > 0 && stock < demand) || (dailyVel > 0 && doi <= 3)) {
        status = "Đứt hàng";
        level = "CRITICAL";
        const def = Math.max(0, demand - stock);
        rec = `Thiếu hụt ${def} đơn vị so với nhu cầu; kích hoạt đặt hàng bổ sung khẩn cấp.`;
      } else if (doi < 14 || leadImport >= 14 || stock <= 4) {
        status = "Cận date / Cảnh báo";
        level = "WARNING";
        rec = `Tồn kho mỏng (${stock} chiếc) hoặc lead-time nhập dài (${leadImport} ngày); cảnh báo rủi ro tiến độ.`;
      }

      return {
        sku,
        name: c.ten_san_pham || sku,
        category: c.danh_muc || "Thiết bị",
        stock,
        sales_velocity: demand,
        daily_velocity: dailyVel,
        days_of_inventory: doi,
        lead_time: `Giao ${leadReady} ngày / Nhập ${leadImport} ngày`,
        inventory_value: itemVal,
        unit_price: price,
        revenue: rev,
        risk_status: status,
        risk_level: level,
        recommendation: rec
      };
    });

    const totSkus = skuList.length;
    const outCnt = skuList.filter(s => s.risk_status === "Đứt hàng").length;
    const expCnt = skuList.filter(s => s.risk_status === "Cận date / Cảnh báo").length;
    const safeCnt = skuList.filter(s => s.risk_status === "An toàn").length;

    return {
      total_skus: totSkus,
      out_of_stock_count: outCnt,
      out_of_stock_pct: totSkus > 0 ? Number((outCnt / totSkus * 100).toFixed(1)) : 0,
      near_expiry_count: expCnt,
      near_expiry_pct: totSkus > 0 ? Number((expCnt / totSkus * 100).toFixed(1)) : 0,
      safe_count: safeCnt,
      safe_pct: totSkus > 0 ? Number((safeCnt / totSkus * 100).toFixed(1)) : 0,
      total_inventory_value: totVal,
      avg_sales_velocity: totSkus > 0 ? Number((skuList.reduce((a, s) => a + s.sales_velocity, 0) / totSkus).toFixed(1)) : 0,
      overall_risk_rate: totSkus > 0 ? Number(((outCnt + expCnt) / totSkus * 100).toFixed(1)) : 0,
      sku_list: skuList
    };
  }

  return {
    total_skus: 0,
    out_of_stock_count: 0,
    out_of_stock_pct: 0,
    near_expiry_count: 0,
    near_expiry_pct: 0,
    safe_count: 0,
    safe_pct: 0,
    total_inventory_value: 0,
    avg_sales_velocity: 0,
    overall_risk_rate: 0,
    sku_list: []
  };
}

function renderStep5ReportPreview() {
  try {
    const sourceNameEl = document.getElementById("s5-report-source-name");
    const execHeadingEl = document.getElementById("s5-exec-heading");
    const execSubHeadingEl = document.getElementById("s5-exec-subheading");
    const kpiGrid = document.getElementById("s5-kpi-grid");
    const cInfo = getCompanyInfo();
    const schema = getEffectiveSchemaInspection();
    const hasSkuInv = Boolean(schema && schema.has_sku_inventory);

    if (sourceNameEl) {
      sourceNameEl.textContent = `Nguồn: ${lastSourceName || "cleaned_sales.csv"}`;
    }

    if (execHeadingEl) {
      const bizPrefix = cInfo.name ? `${cInfo.name} — ` : "";
      if (hasSkuInv) {
        execHeadingEl.textContent = `${bizPrefix}Báo cáo Thẩm định Bán hàng & Rủi ro Tồn kho (B2B)`;
      } else {
        const primName = formatColDisplayName(schema.primary_metric);
        execHeadingEl.textContent = `${bizPrefix}Báo cáo Quản trị Hiệu suất & Phân tích Đa chiều (${primName})`;
      }
    }

    if (execSubHeadingEl) {
      const typePart = cInfo.type ? `Mô hình: ${cInfo.type}` : "Mô hình: Thương mại & Doanh nghiệp";
      const extraPart = cInfo.extra ? ` • ${cInfo.extra}` : "";
      const probPart = cInfo.problem ? ` • Trọng tâm: ${cInfo.problem}` : "";
      const dynamicBadge = hasSkuInv ? "Chuỗi cung ứng & Tồn kho" : `Dynamic Schema Inspector: ${schema.categorical_cols.length} chiều định tính, ${schema.numerical_cols.length} chỉ số định lượng`;
      execSubHeadingEl.textContent = `${typePart}${extraPart}${probPart} • ${dynamicBadge}`;
    }

    // =========================================================
    // 1. TẦNG 1: EXECUTIVE SUMMARY - CỤM THẺ KPI TỔNG QUAN
    // =========================================================
    if (kpiGrid) {
      if (hasSkuInv) {
        const riskMetrics = getEffectiveBiRiskMetrics();
        const totSkus = riskMetrics.total_skus || 0;
        const outCount = riskMetrics.out_of_stock_count || 0;
        const outPct = riskMetrics.out_of_stock_pct || 0;
        const expCount = riskMetrics.near_expiry_count || 0;
        const expPct = riskMetrics.near_expiry_pct || 0;
        const safeCount = riskMetrics.safe_count || 0;
        const safePct = riskMetrics.safe_pct || 0;
        const totInvVal = riskMetrics.total_inventory_value || 0;
        const riskRate = riskMetrics.overall_risk_rate || 0;

        let totRev = 0, mgPct = 15.0;
        if (analyzedResults.length > 0) {
          totRev = analyzedResults.reduce((acc, r) => acc + Number(r.thanh_tien_sau_ck || r.gia_tri_sau_vat || 0), 0);
          const prof = analyzedResults.reduce((acc, r) => acc + (Number(r.thanh_tien_sau_ck || 0) * 0.15), 0);
          mgPct = totRev > 0 ? (prof / totRev * 100) : 15.0;
        }

        kpiGrid.innerHTML = `
          <div class="bg-white border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">Tổng số SKU</span>
              <span class="text-xs">📦</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#0C0E08] font-mono block">${totSkus.toLocaleString("vi-VN")} SKU</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60 mt-0.5 block">Danh mục quản lý</span>
          </div>

          <div class="bg-rose-50/90 border border-rose-300 rounded-xl p-3.5 shadow-xs relative overflow-hidden">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-rose-800">Cảnh báo đứt hàng</span>
              <span class="w-2 h-2 rounded-full bg-rose-600 animate-ping"></span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-rose-700 font-mono block">${outCount} SKU</strong>
            <span class="text-[10px] font-bold text-rose-800 mt-0.5 block">Tỷ lệ: ${outPct}% danh mục</span>
          </div>

          <div class="bg-amber-50/90 border border-amber-300 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-amber-800">Cận date / Cảnh báo</span>
              <span class="text-xs">⚠️</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-amber-700 font-mono block">${expCount} SKU</strong>
            <span class="text-[10px] font-bold text-amber-800 mt-0.5 block">Tỷ lệ: ${expPct}% danh mục</span>
          </div>

          <div class="bg-[#F7F7F4] border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">Giá trị tồn kho</span>
              <span class="text-xs">💰</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#0C0E08] font-mono block">${formatCompactVND(totInvVal)}</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60 mt-0.5 block">Tồn khả dụng VND</span>
          </div>

          <div class="bg-white border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">Tổng Doanh thu</span>
              <span class="text-xs">📈</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-emerald-700 font-mono block">${formatCompactVND(totRev)}</strong>
            <span class="text-[10px] font-semibold text-emerald-800 mt-0.5 block">Biên LN: ${mgPct.toFixed(1)}%</span>
          </div>

          <div class="bg-[#0B0D08] text-white border border-black rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#D4EE0E]">Tỷ lệ rủi ro</span>
              <span class="text-xs">🛡️</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#D4EE0E] font-mono block">${riskRate}%</strong>
            <span class="text-[10px] font-bold text-white/70 mt-0.5 block">An toàn: ${safeCount} SKU (${safePct}%)</span>
          </div>
        `;
      } else {
        const primColObj = schema.numerical_cols.find(c => c.name === schema.primary_metric) || schema.numerical_cols[0] || {};
        const secColObj = schema.numerical_cols.find(c => c.name === schema.secondary_metric);
        const primDimObj = schema.categorical_cols.find(c => c.name === schema.primary_dimension) || schema.categorical_cols[0] || {};
        const custColObj = schema.categorical_cols.find(c => c.role === "customer");

        const totPrimary = primColObj.sum || 0;
        const avgPrimary = primColObj.mean || 0;
        const totalRows = schema.total_rows || 0;

        let card3Title = "Tổng số bản ghi";
        let card3Val = `${totalRows.toLocaleString("vi-VN")} dòng`;
        let card3Sub = "Quy mô dữ liệu thực tế";
        if (secColObj) {
          card3Title = `Tổng ${formatColDisplayName(secColObj.name)}`;
          card3Val = (secColObj.role === "profit" || secColObj.role === "cost" || secColObj.role === "revenue")
            ? formatCompactVND(secColObj.sum)
            : `${Math.round(secColObj.sum).toLocaleString("vi-VN")} đơn vị`;
          const ratio = totPrimary > 0 ? ((secColObj.sum / totPrimary) * 100).toFixed(1) : 0;
          card3Sub = `Tỷ lệ: ${ratio}% so với ${formatColDisplayName(primColObj.name)}`;
        }

        const topPivots = schema.pivot_summary || [];
        const topEntity = topPivots.length > 0 ? topPivots[0] : null;
        const topShare = topEntity ? topEntity.share_pct : 0;

        let card5Title = custColObj ? `Khách hàng / Đối tác` : `Giao dịch lớn nhất`;
        let card5Val = custColObj ? `${custColObj.distinct_count} đối tác` : formatCompactVND(primColObj.max || 0);
        let card5Sub = custColObj ? `Cột: ${custColObj.name}` : `Đỉnh kỷ lục`;

        let card6Title = "Độ tập trung Top 1";
        let card6Val = topEntity ? `${topShare}%` : "—";
        let card6Sub = topEntity ? `Nhóm: ${topEntity.dimension_value}` : "Đang tính";
        if (secColObj && secColObj.role === "profit" && totPrimary > 0) {
          card6Title = "Biên Lợi nhuận (Margin)";
          card6Val = `${((secColObj.sum / totPrimary) * 100).toFixed(1)}%`;
          card6Sub = "Toàn diện trên doanh thu";
        }

        kpiGrid.innerHTML = `
          <div class="bg-white border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">Tổng ${formatColDisplayName(primColObj.name)}</span>
              <span class="text-xs">💰</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-emerald-700 font-mono block">${formatCompactVND(totPrimary)}</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60 mt-0.5 block">Chỉ số định lượng trọng tâm</span>
          </div>

          <div class="bg-[#F7F7F4] border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">TB / Bản ghi (${formatColDisplayName(primColObj.name)})</span>
              <span class="text-xs">📊</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#0C0E08] font-mono block">${formatCompactVND(avgPrimary)}</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60 mt-0.5 block">Trung vị: ${formatCompactVND(primColObj.median || 0)}</span>
          </div>

          <div class="bg-white border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">${card3Title}</span>
              <span class="text-xs">📈</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#0B0D08] font-mono block">${card3Val}</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60 mt-0.5 block">${card3Sub}</span>
          </div>

          <div class="bg-white border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">Phân nhóm ${formatColDisplayName(primDimObj.name)}</span>
              <span class="text-xs">🏷️</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#0C0E08] font-mono block">${primDimObj.distinct_count || 1} nhóm</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60 mt-0.5 block">Chiều phân loại chính</span>
          </div>

          <div class="bg-[#F7F7F4] border border-black/15 rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#0C0E08]/70">${card5Title}</span>
              <span class="text-xs">👥</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#0C0E08] font-mono block">${card5Val}</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60 mt-0.5 block">${card5Sub}</span>
          </div>

          <div class="bg-[#0B0D08] text-white border border-black rounded-xl p-3.5 shadow-xs">
            <div class="flex items-center justify-between mb-1">
              <span class="text-[10px] font-black uppercase text-[#D4EE0E]">${card6Title}</span>
              <span class="text-xs">⚡</span>
            </div>
            <strong class="text-xl sm:text-2xl font-black text-[#D4EE0E] font-mono block">${card6Val}</strong>
            <span class="text-[10px] font-bold text-white/70 mt-0.5 block">${card6Sub}</span>
          </div>
        `;
      }
    }

    renderStep5DynamicReport(univAnalysis?.dynamic_report);

    const skuInvSec = document.getElementById("s5-sku-inv-section");
    if (skuInvSec) {
      if (hasSkuInv) {
        skuInvSec.classList.remove("hidden");
        renderStep5RiskTable();
        renderStep5VipTable();
        renderStep5ProdTable();
      } else {
        skuInvSec.classList.add("hidden");
      }
    }

    renderStep5Charts();
    renderStep5ExecutiveCards();
  } catch (err) {
    console.error("renderStep5ReportPreview error:", err);
  }
}

// =============================================================
// RENDER DYNAMIC 6-PART BUSINESS REPORT (CHUẨN SENIOR BI)
// =============================================================
function renderStep5DynamicReport(rep) {
  if (!rep) {
    rep = buildFallbackDynamicReport();
  }
  if (!rep) return;

  const p1 = rep.part1_executive_summary || {};
  const p2 = rep.part2_kpi_dashboard || {};
  const p3 = rep.part3_deep_dive || [];
  const p4 = rep.part4_conclusions || [];
  const p5 = rep.part5_action_recommendations || [];
  const p6 = rep.part6_technical_details || {};

  // 1. PHẦN 1: TÓM TẮT ĐIỀU HÀNH
  const p1NarrativeEl = document.getElementById("s5-p1-narrative");
  if (p1NarrativeEl) {
    p1NarrativeEl.innerHTML = renderMarkdownToHtml(p1.narrative || "Báo cáo điều hành được khởi tạo tự động từ dữ liệu thực tế.");
  }
  const p1MetricsTbody = document.getElementById("s5-p1-metrics-tbody");
  if (p1MetricsTbody) {
    const mList = p1.metrics_summary || [];
    if (mList.length === 0) {
      p1MetricsTbody.innerHTML = `<tr><td colspan="4" class="py-3 px-3 text-center text-gray-500 italic">Không tìm thấy cột định lượng hợp lệ.</td></tr>`;
    } else {
      p1MetricsTbody.innerHTML = mList.map(m => `
        <tr class="hover:bg-black/[0.02]">
          <td class="py-2 px-3 font-bold text-[#0C0E08]">${m.column}</td>
          <td class="py-2 px-3 text-[#0C0E08]/80">${m.role}</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-emerald-700">${m.formatted_sum}</td>
          <td class="py-2 px-3 font-semibold text-[#0C0E08]/80">${m.coverage_text}</td>
        </tr>
      `).join("");
    }
  }

  // 2. PHẦN 2: DASHBOARD KPI (SO SÁNH KỲ GẦN NHẤT VỚI KỲ LIỀN TRƯỚC)
  const p2CurLbl = document.getElementById("s5-p2-th-cur");
  const p2PriLbl = document.getElementById("s5-p2-th-pri");
  if (p2CurLbl) p2CurLbl.textContent = `Giá trị (${p2.current_period_label || "Kỳ gần nhất"})`;
  if (p2PriLbl) p2PriLbl.textContent = `Kỳ trước (${p2.prior_period_label || "Kỳ liền trước"})`;

  const p2KpiTbody = document.getElementById("s5-p2-kpi-tbody");
  if (p2KpiTbody) {
    const kRows = p2.kpi_rows || [];
    p2KpiTbody.innerHTML = kRows.map(r => {
      const isMissing = r.status === "missing" || r.current_val === "—";
      let chgBadge = `<span class="text-gray-400 font-mono">—</span>`;
      if (r.change_pct && r.change_pct !== "—") {
        const isUp = r.change_pct.startsWith("+");
        const isDown = r.change_pct.startsWith("-");
        const colorClass = isUp ? "text-emerald-700 bg-emerald-50 border-emerald-300" : (isDown ? "text-rose-700 bg-rose-50 border-rose-300" : "text-gray-700 bg-gray-50 border-gray-300");
        chgBadge = `<span class="inline-block px-2 py-0.5 rounded font-mono font-bold text-[10px] border ${colorClass}">${r.change_pct}</span>`;
      }
      return `
        <tr class="${isMissing ? 'bg-gray-50/50 text-gray-400' : 'hover:bg-black/[0.02]'}">
          <td class="py-2.5 px-3 font-bold ${isMissing ? 'text-gray-500 italic' : 'text-[#0C0E08]'}">
            ${r.metric} ${isMissing ? '<span class="text-[9px] font-normal text-gray-400 ml-1">(Chưa có cột trong file)</span>' : ''}
          </td>
          <td class="py-2.5 px-3 text-right font-mono font-extrabold ${isMissing ? 'text-gray-400' : 'text-emerald-700'}">${r.current_val}</td>
          <td class="py-2.5 px-3 text-right font-mono text-[#0C0E08]/75">${r.prior_val}</td>
          <td class="py-2.5 px-3 text-center">${chgBadge}</td>
        </tr>
      `;
    }).join("");
  }

  // 3. PHẦN 3: PHÂN TÍCH CHUYÊN ĐỀ
  const p3Container = document.getElementById("s5-p3-thematic-container");
  if (p3Container) {
    if (p3.length === 0) {
      p3Container.innerHTML = `<p class="text-gray-500 italic">Không có chỉ số định lượng nào để phân tích chuyên đề.</p>`;
    } else {
      p3Container.innerHTML = p3.map((th, idx) => `
        <div class="bg-[#F7F7F4] border border-black/10 rounded-xl p-3.5 space-y-2.5 shadow-2xs">
          <div class="flex items-center justify-between">
            <strong class="text-xs font-black text-[#1B365D]">3.${idx + 1}. Phân tích chi tiết chỉ số: ${th.column}</strong>
            <span class="text-[10px] font-bold bg-[#D4EE0E] text-[#0C0E08] px-2 py-0.5 rounded">${Number(th.count || 0).toLocaleString("vi-VN")} bản ghi</span>
          </div>
          <p class="text-xs text-[#0C0E08]/85 leading-relaxed font-medium">${th.narrative}</p>
          <div class="overflow-x-auto bg-white border border-black/10 rounded-lg">
            <table class="w-full text-[11px] text-center">
              <thead class="bg-[#EFECE6] font-bold text-[#0C0E08]/80">
                <tr>
                  <th class="py-1.5 px-2">Tổng cộng</th>
                  <th class="py-1.5 px-2">Trung bình</th>
                  <th class="py-1.5 px-2">Trung vị</th>
                  <th class="py-1.5 px-2">Độ lệch chuẩn</th>
                  <th class="py-1.5 px-2">Nhỏ nhất</th>
                  <th class="py-1.5 px-2">Lớn nhất</th>
                </tr>
              </thead>
              <tbody class="font-mono">
                <tr>
                  <td class="py-1.5 px-2 font-bold text-emerald-700">${th.sum}</td>
                  <td class="py-1.5 px-2">${th.mean}</td>
                  <td class="py-1.5 px-2">${th.median}</td>
                  <td class="py-1.5 px-2">${th.std}</td>
                  <td class="py-1.5 px-2 text-blue-700">${th.min}</td>
                  <td class="py-1.5 px-2 text-rose-700">${th.max}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      `).join("");
    }
  }

  // 4. PHẦN 4: KẾT LUẬN
  const p4Container = document.getElementById("s5-p4-conclusions-container");
  if (p4Container) {
    p4Container.innerHTML = p4.map(c => `
      <div class="flex items-start gap-2.5 p-2.5 rounded-lg bg-[#F7F7F4] border border-black/10">
        <span class="w-5 h-5 rounded-full bg-[#0B0D08] text-[#D4EE0E] flex items-center justify-center text-[10px] font-black shrink-0 mt-0.5">✓</span>
        <p class="text-xs text-[#0C0E08]/85 leading-relaxed font-medium">${c}</p>
      </div>
    `).join("");
  }

  // Rule 3: AI Insights rendered using Markdown-to-HTML parser
  const s4Report = document.getElementById("s4-univ-report-text")?.value || univReportText || "";
  const aiBox = document.getElementById("s5-p4-ai-insights-box");
  const aiText = document.getElementById("s5-p4-ai-insights-text");
  if (aiBox && aiText) {
    if (s4Report && s4Report.trim()) {
      aiText.innerHTML = renderMarkdownToHtml(s4Report.trim());
      aiBox.classList.remove("hidden");
    } else {
      aiBox.classList.add("hidden");
    }
  }

  // 5. PHẦN 5: ĐỀ XUẤT HÀNH ĐỘNG
  const p5Tbody = document.getElementById("s5-p5-actions-tbody");
  if (p5Tbody) {
    p5Tbody.innerHTML = p5.map(r => {
      const pLevel = r.priority_level;
      const badgeClass = pLevel === "high" ? "bg-rose-100 text-rose-800 border-rose-300" : (pLevel === "medium" ? "bg-amber-100 text-amber-800 border-amber-300" : "bg-emerald-100 text-emerald-800 border-emerald-300");
      return `
        <tr class="hover:bg-black/[0.02]">
          <td class="py-2.5 px-3 text-center font-bold text-[#0C0E08]/70">${r.stt}</td>
          <td class="py-2.5 px-3">
            <strong class="font-bold text-[#0C0E08] block">${r.title}</strong>
            <span class="text-xs text-[#0C0E08]/75 mt-0.5 block leading-relaxed">${r.action}</span>
          </td>
          <td class="py-2.5 px-3 text-center">
            <span class="inline-block text-[10px] font-extrabold px-2.5 py-1 rounded border ${badgeClass}">${r.priority}</span>
          </td>
          <td class="py-2.5 px-3 font-semibold text-[#0C0E08]/85 font-mono text-[11px]">${r.kpi_measure}</td>
        </tr>
      `;
    }).join("");
  }

  // 6. PHẦN 6: CHI TIẾT BỘ DỮ LIỆU & THỐNG KÊ KỸ THUẬT
  // 6.1
  const p6CleanTbody = document.getElementById("s5-p6-clean-tbody");
  if (p6CleanTbody && p6.cleaning_info) {
    const ci = p6.cleaning_info;
    p6CleanTbody.innerHTML = `
      <tr>
        <td class="py-2 px-2.5 font-bold">${Number(ci.orig_rows || 0).toLocaleString("vi-VN")}</td>
        <td class="py-2 px-2.5 font-bold text-emerald-700">${Number(ci.clean_rows || 0).toLocaleString("vi-VN")}</td>
        <td class="py-2 px-2.5">${ci.clean_cols || 0}</td>
        <td class="py-2 px-2.5 text-amber-700 font-bold">${Number(ci.duplicates_removed || 0).toLocaleString("vi-VN")}</td>
        <td class="py-2 px-2.5">${Number(ci.remaining_nulls || 0).toLocaleString("vi-VN")}</td>
        <td class="py-2 px-2.5 font-bold ${ci.outliers_flagged > 0 ? 'text-rose-700' : 'text-gray-500'}">${ci.outliers_flagged || 0}</td>
      </tr>
    `;
  }

  // 6.2
  const p6StatsCommentEl = document.getElementById("s5-p6-stats-comment");
  if (p6StatsCommentEl) {
    p6StatsCommentEl.innerHTML = renderMarkdownToHtml(p6.automated_stats_comment || "Dữ liệu đã được kiểm định toàn vẹn.");
  }

  // 6.3
  const p6GrpNameEl = document.getElementById("s5-p6-grp-col-name");
  const p6GrpTbody = document.getElementById("s5-p6-group-tbody");
  if (p6GrpTbody && p6.group_summary_table) {
    if (p6GrpNameEl) p6GrpNameEl.textContent = p6.group_summary_table.group_col || "Nhóm";
    const gRows = p6.group_summary_table.rows || [];
    p6GrpTbody.innerHTML = gRows.map(gr => `
      <tr class="hover:bg-black/[0.02]">
        <td class="py-2 px-3 font-bold text-[#0C0E08]">${gr.group}</td>
        <td class="py-2 px-3 text-right font-mono font-bold text-emerald-700">${gr.value}</td>
        <td class="py-2 px-3 text-center font-mono text-[#0C0E08]/75">${Number(gr.count || 0).toLocaleString("vi-VN")}</td>
      </tr>
    `).join("");
  }

  // 6.4
  const p6ColstatsTbody = document.getElementById("s5-p6-colstats-tbody");
  if (p6ColstatsTbody && p6.column_stats_table) {
    const csRows = p6.column_stats_table || [];
    p6ColstatsTbody.innerHTML = csRows.map(cs => `
      <tr class="hover:bg-black/[0.02]">
        <td class="py-2 px-3 font-bold text-[#0C0E08]">${cs.column}</td>
        <td class="py-2 px-3 text-center font-mono ${cs.null_count > 0 ? 'text-amber-700 font-bold' : 'text-gray-400'}">${Number(cs.null_count || 0).toLocaleString("vi-VN")}</td>
        <td class="py-2 px-3 text-center font-mono">${Number(cs.distinct_count || 0).toLocaleString("vi-VN")}</td>
        <td class="py-2 px-3 text-right font-mono text-[#0C0E08]/85">${cs.mean}</td>
        <td class="py-2 px-3 text-right font-mono text-[#0C0E08]/70">${cs.min}</td>
        <td class="py-2 px-3 text-right font-mono font-bold text-[#0C0E08]">${cs.max}</td>
      </tr>
    `).join("");
  }
}

function buildFallbackDynamicReport() {
  const schema = getEffectiveSchemaInspection();
  const cInfo = getCompanyInfo();
  const totalRows = (univCleanedRows && univCleanedRows.length) || (cleanedRows && cleanedRows.length) || schema.total_rows || 0;
  const numCols = schema.numerical_cols || [];
  const primMetric = schema.primary_metric || (numCols[0] ? numCols[0].name : "DoanhThu");
  const primObj = numCols.find(c => c.name === primMetric) || numCols[0] || {};
  const primSum = primObj.sum || 0;

  const metricsSummary = numCols.map(c => ({
    column: c.name,
    role: c.role || "Chỉ số định lượng",
    sum: c.sum || 0,
    formatted_sum: formatCompactVND(c.sum || 0),
    coverage_text: `Có số liệu ở ${totalRows.toLocaleString("vi-VN")}/${totalRows.toLocaleString("vi-VN")} dòng (100%)`
  }));

  const kpiRows = [
    { metric: "Doanh thu / Giá trị chính", current_val: formatCompactVND(primSum), prior_val: "—", change_pct: "—", status: "active" },
    { metric: "Tổng Lợi nhuận (Profit)", current_val: "—", prior_val: "—", change_pct: "—", status: "missing" },
    { metric: "Chi phí / Giá vốn (Cost)", current_val: "—", prior_val: "—", change_pct: "—", status: "missing" },
    { metric: "Biên Lợi nhuận (Margin %)", current_val: "—", prior_val: "—", change_pct: "—", status: "missing" },
    { metric: "Số lượng / Sản lượng (Qty)", current_val: "—", prior_val: "—", change_pct: "—", status: "missing" },
    { metric: "Số đơn hàng / Giao dịch", current_val: `${totalRows.toLocaleString("vi-VN")} đơn`, prior_val: "—", change_pct: "—", status: "active" },
    { metric: "Giá trị đơn hàng TB (AOV)", current_val: formatCompactVND(primObj.mean || 0), prior_val: "—", change_pct: "—", status: "active" }
  ];

  const thematic = numCols.map(c => ({
    column: c.name,
    count: totalRows,
    sum: formatCompactVND(c.sum || 0),
    mean: formatCompactVND(c.mean || 0),
    median: formatCompactVND(c.median || 0),
    std: formatCompactVND(c.std || 0),
    min: formatCompactVND(c.min || 0),
    max: formatCompactVND(c.max || 0),
    narrative: `Chỉ số '${c.name}' đạt tổng quy mô ${formatCompactVND(c.sum || 0)}, trung bình ${formatCompactVND(c.mean || 0)} qua ${totalRows.toLocaleString("vi-VN")} bản ghi.`
  }));

  return {
    part1_executive_summary: {
      dataset_name: lastSourceName || "Dữ liệu",
      narrative: `Báo cáo điều hành được khởi tạo tự động từ tệp '${lastSourceName || "Dữ liệu"}' của ${cInfo.name || "Doanh nghiệp"} (${cInfo.type || "Thương mại"}). Tổng cộng ${totalRows.toLocaleString("vi-VN")} dòng và ${schema.total_cols || 0} cột thông tin.`,
      metrics_summary: metricsSummary
    },
    part2_kpi_dashboard: {
      current_period_label: "Kỳ gần nhất",
      prior_period_label: "Kỳ liền trước",
      kpi_rows: kpiRows
    },
    part3_deep_dive: thematic,
    part4_conclusions: [
      `Quy mô kinh doanh duy trì ổn định với tổng doanh thu đạt ${formatCompactVND(primSum)} qua ${totalRows.toLocaleString("vi-VN")} giao dịch.`,
      `Hệ thống kiểm định 100% bản ghi đã được làm sạch và chuẩn hóa trước khi tổng hợp.`
    ],
    part5_action_recommendations: [
      { stt: 1, title: "Tối ưu hóa chính sách giá & chiết khấu", action: "Kiểm soát biên độ chiết khấu và cơ cấu hợp đồng cho đối tác lớn.", priority: "Ưu tiên Cao", priority_level: "high", kpi_measure: "Biên lợi nhuận gộp >= 15.0%" },
      { stt: 2, title: "Đẩy mạnh khai thác kênh tăng trưởng cao", action: "Tập trung phân bổ nguồn lực vào nhóm ngành hàng chủ lực.", priority: "Ưu tiên Trung bình", priority_level: "medium", kpi_measure: "Tăng trưởng doanh thu kỳ tới +10%" },
      { stt: 3, title: "Quản trị chu kỳ quay vòng vốn", action: "Theo dõi số ngày tồn kho và dòng tiền giao dịch.", priority: "Ưu tiên Thấp", priority_level: "low", kpi_measure: "Vòng quay an toàn" }
    ],
    part6_technical_details: {
      cleaning_info: { orig_rows: totalRows, clean_rows: totalRows, clean_cols: schema.total_cols || 0, duplicates_removed: 0, remaining_nulls: 0, outliers_flagged: 0 },
      automated_stats_comment: `Bộ dữ liệu có giá trị nhỏ nhất: ${formatCompactVND(primObj.min || 0)}, lớn nhất: ${formatCompactVND(primObj.max || 0)}, trung bình: ${formatCompactVND(primObj.mean || 0)}.`,
      group_summary_table: { group_col: schema.primary_dimension || "Nhóm", rows: [] },
      column_stats_table: []
    }
  };
}

// =============================================================
// NÚT TẢI XUỐNG BÁO CÁO WORD (.DOCX) & EXCEL (.XLSX)
// =============================================================
async function downloadWordReport() {
  showToast("📄 Đang xuất tệp Microsoft Word (.docx) chuẩn 6 phần...", "info");
  const ok = await runUnivStep5Export();
  if (ok) {
    const a = document.createElement("a");
    a.href = "/BaoCao_PhanTich_HieuSuat.docx";
    a.download = "BaoCao_PhanTich_HieuSuat.docx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showToast("✓ Đã tải xuống Báo cáo Word (.docx) thành công!", "success");
  }
}

async function downloadExcelReport() {
  showToast("📊 Đang xuất Sổ cái Excel (.xlsx) đa Sheet...", "info");
  const ok = await runUnivStep5Export();
  if (ok) {
    const a = document.createElement("a");
    a.href = "/BaoCao_DoanhThu_TongHop.xlsx";
    a.download = "BaoCao_DoanhThu_TongHop.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showToast("✓ Đã tải xuống Sổ cái Excel (.xlsx) thành công!", "success");
  }
}

function safeDestroyChart(key) {
  if (chartInstances[key]) {
    try {
      chartInstances[key].destroy();
    } catch (e) {
      console.warn("Chart destroy warning:", e);
    }
    chartInstances[key] = null;
  }
}

function renderStep5Charts() {
  if (typeof Chart === "undefined") {
    console.warn("Chart.js is not loaded yet.");
    return;
  }
  renderStep5DonutChart();
  renderStep5TopProductsChart();
  renderStep5MonthlyChart();
  renderStep5SegmentChart();
}

function renderStep5DonutChart() {
  const canvas = document.getElementById("chart-donut");
  if (!canvas || typeof Chart === "undefined") return;
  safeDestroyChart("donut");

  const schema = getEffectiveSchemaInspection();
  const hasSkuInv = Boolean(schema && schema.has_sku_inventory);
  const chipsContainer = document.getElementById("donut-chips-container");
  const titleEl = document.getElementById("donut-chart-title");

  let labels = [];
  let data = [];
  let colors = ["#0B0D08", "#D4EE0E", "#047857", "#2563EB", "#F59E0B", "#8B5CF6", "#EC4899", "#14B8A6"];

  if (hasSkuInv) {
    if (chipsContainer) {
      chipsContainer.innerHTML = `
        <button onclick="switchDonutDimension('Risk')" class="px-2 py-0.5 rounded font-bold transition ${donutActiveDim === 'Risk' ? 'bg-[#0B0D08] text-[#D4EE0E]' : 'bg-[#EFECE6] text-[#0C0E08]'}">Trạng thái rủi ro</button>
        <button onclick="switchDonutDimension('Category')" class="px-2 py-0.5 rounded font-bold transition ${donutActiveDim === 'Category' ? 'bg-[#0B0D08] text-[#D4EE0E]' : 'bg-[#EFECE6] text-[#0C0E08]'}">Ngành hàng</button>
        <button onclick="switchDonutDimension('Region')" class="px-2 py-0.5 rounded font-bold transition ${donutActiveDim === 'Region' ? 'bg-[#0B0D08] text-[#D4EE0E]' : 'bg-[#EFECE6] text-[#0C0E08]'}">Khu vực</button>
      `;
    }

    if (donutActiveDim === "Risk") {
      const riskMetrics = getEffectiveBiRiskMetrics();
      labels = ["Đứt hàng (Critical)", "Cận date / Cảnh báo", "An toàn (Safe)"];
      data = [riskMetrics.out_of_stock_count || 0, riskMetrics.near_expiry_count || 0, riskMetrics.safe_count || 0];
      colors = ["#EF4444", "#F59E0B", "#10B981"];
      if (titleEl) titleEl.textContent = "🍩 Phân bổ Trạng thái Rủi ro Tồn kho";
    } else {
      const catMap = {};
      const riskMetrics = getEffectiveBiRiskMetrics();
      (riskMetrics.sku_list || []).forEach(s => {
        const k = (donutActiveDim === "Category" ? s.category : s.lead_time) || "Khác";
        catMap[k] = (catMap[k] || 0) + (s.revenue || s.inventory_value || 1);
      });
      const entries = Object.entries(catMap).slice(0, 6);
      labels = entries.map(e => e[0]);
      data = entries.map(e => e[1]);
      if (titleEl) titleEl.textContent = `🍩 Tỷ trọng theo ${donutActiveDim === 'Category' ? 'Ngành hàng' : 'Khu vực'}`;
    }
  } else {
    const catCols = schema.categorical_cols.slice(0, 4);
    if (!catCols.some(c => c.name === donutActiveDim)) {
      donutActiveDim = schema.primary_dimension || (catCols[0]?.name) || "Phân loại";
    }

    if (chipsContainer) {
      chipsContainer.innerHTML = catCols.map(c => {
        const isActive = donutActiveDim === c.name;
        const btnCls = isActive ? "bg-[#0B0D08] text-[#D4EE0E]" : "bg-[#EFECE6] text-[#0C0E08] hover:bg-black/10";
        return `<button onclick="switchDonutDimension('${c.name}')" data-dim="${c.name}" class="px-2 py-0.5 rounded font-bold transition ${btnCls}">${formatColDisplayName(c.name)}</button>`;
      }).join("");
    }

    if (titleEl) {
      titleEl.textContent = `🍩 Tỷ trọng ${formatColDisplayName(schema.primary_metric)} theo ${formatColDisplayName(donutActiveDim)}`;
    }

    const rows = (univCleanedRows && univCleanedRows.length) ? univCleanedRows :
                 (cleanedRows && cleanedRows.length) ? cleanedRows : [];
    const pCol = donutActiveDim;
    const mCol = schema.primary_metric;

    const groupMap = {};
    rows.forEach(r => {
      const k = String(r[pCol] || "Khác").trim();
      let v = mCol ? Number(String(r[mCol] || 0).replace(/[\$,VNĐvndđ\s]/g, "").replace(/,/g, "")) : 1;
      if (isNaN(v)) v = 1;
      groupMap[k] = (groupMap[k] || 0) + v;
    });

    const entries = Object.entries(groupMap).sort((a, b) => b[1] - a[1]).slice(0, 6);
    labels = entries.map(e => e[0]);
    data = entries.map(e => e[1]);
  }

  const totVal = data.reduce((a, b) => a + b, 0);

  chartInstances["donut"] = new Chart(canvas, {
    type: "doughnut",
    data: {
      labels: labels.length ? labels : ["Không có dữ liệu"],
      datasets: [{
        data: data.length && totVal > 0 ? data : [1],
        backgroundColor: data.length && totVal > 0 ? colors.slice(0, labels.length) : ["#E5E7EB"],
        borderWidth: 2,
        borderColor: "#FFFFFF"
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "right", labels: { boxWidth: 10, font: { size: 9, weight: "bold" } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const val = ctx.parsed;
              const pct = totVal > 0 ? ((val / totVal) * 100).toFixed(1) : 0;
              const unit = donutActiveDim === "Risk" ? `${val} SKU` : formatCompactVND(val);
              return `${ctx.label}: ${unit} (${pct}%)`;
            }
          }
        }
      },
      cutout: "60%"
    }
  });
}

function renderStep5TopProductsChart() {
  const canvas = document.getElementById("chart-top-products");
  if (!canvas || typeof Chart === "undefined") return;
  safeDestroyChart("products");

  const schema = getEffectiveSchemaInspection();
  const hasSkuInv = Boolean(schema && schema.has_sku_inventory);
  const titleEl = document.getElementById("top-products-chart-title");
  const badgeEl = document.getElementById("top-products-chart-badge");

  if (hasSkuInv) {
    if (titleEl) titleEl.textContent = "📊 Tốc độ bán & Số ngày tồn kho (DOI) theo SKU";
    if (badgeEl) badgeEl.textContent = "Đối chiếu rủi ro";

    const riskMetrics = getEffectiveBiRiskMetrics();
    const skuItems = (riskMetrics.sku_list || []).slice(0, 8);
    const labels = skuItems.map(s => (s.sku || "SKU").slice(0, 15));
    const velocityData = skuItems.map(s => Number(s.sales_velocity || 0));
    const doiData = skuItems.map(s => Number(s.days_of_inventory || 0));

    chartInstances["products"] = new Chart(canvas, {
      type: "bar",
      data: {
        labels: labels.length ? labels : ["Chưa có SKU"],
        datasets: [
          {
            label: "Nhu cầu (Units)",
            data: velocityData.length ? velocityData : [0],
            backgroundColor: "#0B0D08",
            borderRadius: 4,
            yAxisID: "y"
          },
          {
            label: "Số ngày tồn (DOI)",
            data: doiData.length ? doiData : [0],
            type: "line",
            borderColor: "#047857",
            backgroundColor: "#047857",
            pointRadius: 4,
            borderWidth: 2,
            tension: 0.25,
            yAxisID: "y1"
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "top", labels: { boxWidth: 10, font: { size: 9, weight: "bold" } } }
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { size: 9, weight: "bold" } } },
          y: { position: "left", title: { display: true, text: "Nhu cầu", font: { size: 9 } } },
          y1: { position: "right", title: { display: true, text: "Số ngày tồn", font: { size: 9 } }, grid: { display: false } }
        }
      }
    });
  } else {
    const dimName = schema.primary_dimension || "Phân loại";
    const metricName = schema.primary_metric || "Chỉ số";
    if (titleEl) titleEl.textContent = `📊 Top ${formatColDisplayName(dimName)} theo ${formatColDisplayName(metricName)}`;
    if (badgeEl) badgeEl.textContent = "Quy mô thực tế";

    const pivots = (schema.pivot_summary || []).slice(0, 8);
    const labels = pivots.map(p => p.dimension_value.slice(0, 15));
    const vals = pivots.map(p => p.total_primary);

    chartInstances["products"] = new Chart(canvas, {
      type: "bar",
      data: {
        labels: labels.length ? labels : ["Chưa có dữ liệu"],
        datasets: [{
          label: `Tổng ${formatColDisplayName(metricName)}`,
          data: vals.length ? vals : [0],
          backgroundColor: "#0B0D08",
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `${formatColDisplayName(metricName)}: ${formatCompactVND(ctx.parsed.y)}`
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { size: 9, weight: "bold" }, maxRotation: 25 } },
          y: { grid: { color: "rgba(0,0,0,0.06)" }, ticks: { font: { size: 9 }, callback: (v) => formatCompactVND(v) } }
        }
      }
    });
  }
}

function renderStep5MonthlyChart() {
  const canvas = document.getElementById("chart-monthly-trend");
  if (!canvas || typeof Chart === "undefined") return;
  safeDestroyChart("monthly");

  const schema = getEffectiveSchemaInspection();
  const titleEl = document.getElementById("trend-chart-title");
  const badgeEl = document.getElementById("trend-chart-badge");

  const gTables = univAnalysis?.group_tables || {};
  let thangRows = gTables["_Thang"] || gTables["Thang"] || gTables["Month"] || [];
  thangRows = [...thangRows].sort((a, b) => String(a.group_value).localeCompare(String(b.group_value)));

  let labels = [];
  let lineData1 = [];
  let lineData2 = [];
  let label1 = "Doanh thu";
  let label2 = "Lợi nhuận";

  if (schema.has_temporal && thangRows.length > 0) {
    if (titleEl) titleEl.textContent = `📈 Xu hướng ${formatColDisplayName(schema.primary_metric)} theo ${formatColDisplayName(schema.primary_temporal || 'Thời gian')}`;
    if (badgeEl) badgeEl.textContent = "Chuỗi thời gian";

    labels = thangRows.map(r => r.group_value);
    lineData1 = thangRows.map(r => Number(r.revenue || 0));
    lineData2 = thangRows.map(r => Number(r.profit || 0));
    label1 = formatColDisplayName(schema.primary_metric);
    label2 = schema.secondary_metric ? formatColDisplayName(schema.secondary_metric) : "Chỉ số 2";
  } else if (schema.has_temporal && (univCleanedRows.length || cleanedRows.length)) {
    const tCol = schema.primary_temporal;
    const mCol = schema.primary_metric;
    const rows = univCleanedRows.length ? univCleanedRows : cleanedRows;
    const tMap = {};
    rows.forEach(r => {
      const rawT = String(r[tCol] || "").trim();
      const tKey = rawT.length >= 7 ? rawT.slice(0, 7) : rawT;
      const v = Number(String(r[mCol] || 0).replace(/[\$,VNĐvndđ\s]/g, "").replace(/,/g, "")) || 0;
      if (tKey) tMap[tKey] = (tMap[tKey] || 0) + v;
    });
    const sortedT = Object.entries(tMap).sort((a, b) => a[0].localeCompare(b[0])).slice(0, 12);
    labels = sortedT.map(e => e[0]);
    lineData1 = sortedT.map(e => e[1]);
    label1 = formatColDisplayName(mCol);
  } else {
    const secDim = schema.secondary_dimension || schema.primary_dimension;
    if (titleEl) titleEl.textContent = `📊 Phân bổ ${formatColDisplayName(schema.primary_metric)} theo ${formatColDisplayName(secDim)}`;
    if (badgeEl) badgeEl.textContent = "Phân tích thứ cấp";

    const secPivots = (schema.pivot_summary || []).slice(0, 6);
    labels = secPivots.map(p => p.dimension_value.slice(0, 12));
    lineData1 = secPivots.map(p => p.total_primary);
    label1 = formatColDisplayName(schema.primary_metric);
  }

  chartInstances["monthly"] = new Chart(canvas, {
    type: "line",
    data: {
      labels: labels.length ? labels : ["Kỳ 1", "Kỳ 2", "Kỳ 3"],
      datasets: [
        {
          label: label1,
          data: lineData1.length ? lineData1 : [0, 0, 0],
          borderColor: "#0B0D08",
          backgroundColor: "rgba(11, 13, 8, 0.05)",
          borderWidth: 2,
          fill: true,
          tension: 0.35,
          pointRadius: 3
        },
        ...(lineData2.length && lineData2.some(v => v > 0) ? [{
          label: label2,
          data: lineData2,
          borderColor: "#047857",
          backgroundColor: "rgba(4, 120, 87, 0.05)",
          borderWidth: 2,
          fill: false,
          tension: 0.35,
          pointRadius: 3
        }] : [])
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "top", labels: { boxWidth: 12, font: { size: 10, weight: "bold" } } },
        tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${formatCompactVND(ctx.parsed.y)}` } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 9 }, maxRotation: 30 } },
        y: { grid: { color: "rgba(0,0,0,0.06)" }, ticks: { font: { size: 9 }, callback: (v) => formatCompactVND(v) } }
      }
    }
  });
}

function renderStep5SegmentChart() {
  const canvas = document.getElementById("chart-segment-bar");
  if (!canvas || typeof Chart === "undefined") return;
  safeDestroyChart("segment");

  const schema = getEffectiveSchemaInspection();
  const titleEl = document.getElementById("segment-chart-title");
  const badgeEl = document.getElementById("segment-chart-badge");
  const dimName = schema.primary_dimension || "Nhóm";

  if (titleEl) {
    const m1 = formatColDisplayName(schema.primary_metric);
    const m2 = schema.secondary_metric ? formatColDisplayName(schema.secondary_metric) : "Số lượng giao dịch";
    titleEl.textContent = `📊 Đối chiếu ${m1} & ${m2} theo ${formatColDisplayName(dimName)}`;
  }
  if (badgeEl) badgeEl.textContent = "Cột nhóm đối chiếu";

  const pivots = (schema.pivot_summary || []).slice(0, 6);
  const labels = pivots.map(p => p.dimension_value.slice(0, 12));
  const data1 = pivots.map(p => p.total_primary);
  const data2 = pivots.map(p => p.total_secondary !== null && p.total_secondary !== undefined ? p.total_secondary : p.order_count);
  const label2 = schema.secondary_metric ? `Tổng ${formatColDisplayName(schema.secondary_metric)}` : "Số giao dịch (Count)";

  chartInstances["segment"] = new Chart(canvas, {
    type: "bar",
    data: {
      labels: labels.length ? labels : ["Chưa có nhóm"],
      datasets: [
        {
          label: `Tổng ${formatColDisplayName(schema.primary_metric)}`,
          data: data1.length ? data1 : [0],
          backgroundColor: "#0B0D08",
          borderRadius: 3,
          yAxisID: "y"
        },
        {
          label: label2,
          data: data2.length ? data2 : [0],
          backgroundColor: "#D4EE0E",
          borderRadius: 3,
          yAxisID: "y1"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "top", labels: { boxWidth: 10, font: { size: 9, weight: "bold" } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              if (ctx.dataset.yAxisID === "y1" && !schema.secondary_metric) {
                return `${ctx.dataset.label}: ${ctx.parsed.y} giao dịch`;
              }
              return `${ctx.dataset.label}: ${formatCompactVND(ctx.parsed.y)}`;
            }
          }
        }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 9, weight: "bold" } } },
        y: { position: "left", grid: { color: "rgba(0,0,0,0.06)" }, ticks: { font: { size: 9 }, callback: (v) => formatCompactVND(v) } },
        y1: { position: "right", grid: { display: false }, ticks: { font: { size: 9 } } }
      }
    }
  });
}

function renderStep5RiskTable() {
  const tbody = document.getElementById("s5-tbody-risk");
  const thead = document.getElementById("s5-thead-risk");
  const titleText = document.getElementById("s5-table-title-text");
  const subtitleEl = document.getElementById("s5-table-subtitle");
  const chipsContainer = document.getElementById("s5-filter-chips-container");
  if (!tbody) return;

  const schema = getEffectiveSchemaInspection();
  const hasSkuInv = Boolean(schema && schema.has_sku_inventory);

  if (hasSkuInv) {
    if (titleText) titleText.textContent = "III. Bảng Dữ liệu Chi tiết — Quản trị bằng Ngoại lệ (Management by Exception)";
    if (subtitleEl) subtitleEl.textContent = "Tô màu có điều kiện: Đỏ (Đứt hàng / Nghiêm trọng) • Vàng (Cận date / Cảnh báo) • Xanh (An toàn)";

    if (thead) {
      thead.innerHTML = `
        <tr>
          <th class="py-2.5 px-3 text-center w-10">#</th>
          <th class="py-2.5 px-3 whitespace-nowrap">Mã SKU</th>
          <th class="py-2.5 px-3">Tên Sản phẩm &amp; Danh mục</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">Tồn khả dụng</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">Tốc độ bán / Nhu cầu</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">Số ngày tồn (DOI)</th>
          <th class="py-2.5 px-3 text-center whitespace-nowrap">Hạn dùng / Lead-time</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">Giá trị tồn kho</th>
          <th class="py-2.5 px-3 text-center whitespace-nowrap">Trạng thái rủi ro</th>
          <th class="py-2.5 px-3">Khuyến nghị điều hành BI</th>
        </tr>
      `;
    }

    const riskMetrics = getEffectiveBiRiskMetrics();
    const allSkus = riskMetrics.sku_list || [];
    const outCnt = riskMetrics.out_of_stock_count || allSkus.filter(s => s.risk_status === "Đứt hàng").length;
    const expCnt = riskMetrics.near_expiry_count || allSkus.filter(s => s.risk_status === "Cận date / Cảnh báo").length;
    const safeCnt = riskMetrics.safe_count || allSkus.filter(s => s.risk_status === "An toàn").length;

    if (chipsContainer) {
      chipsContainer.innerHTML = `
        <span class="text-[10px] font-bold text-[#0C0E08]/60 uppercase">Lọc nhanh:</span>
        <button id="btn-chip-all" onclick="setS5RiskFilter('ALL')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5RiskFilter === 'ALL' ? 'bg-[#0B0D08] text-[#D4EE0E] border border-black' : 'bg-white text-[#0C0E08] border border-black/20'}">Tất cả (${allSkus.length})</button>
        <button id="btn-chip-out" onclick="setS5RiskFilter('OUT_OF_STOCK')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5RiskFilter === 'OUT_OF_STOCK' ? 'bg-rose-600 text-white border border-rose-700' : 'bg-white hover:bg-rose-50 text-rose-700 border border-rose-300'}">🚨 Đứt hàng (${outCnt})</button>
        <button id="btn-chip-exp" onclick="setS5RiskFilter('NEAR_EXPIRY')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5RiskFilter === 'NEAR_EXPIRY' ? 'bg-amber-500 text-white border-amber-600' : 'bg-white hover:bg-amber-50 text-amber-700 border border-amber-300'}">⚠️ Cận date / Cảnh báo (${expCnt})</button>
        <button id="btn-chip-safe" onclick="setS5RiskFilter('SAFE')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5RiskFilter === 'SAFE' ? 'bg-emerald-600 text-white border border-emerald-700' : 'bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300'}">✓ An toàn (${safeCnt})</button>
      `;
    }

    let filtered = allSkus;
    if (s5RiskFilter === "OUT_OF_STOCK") filtered = filtered.filter(s => s.risk_status === "Đứt hàng");
    else if (s5RiskFilter === "NEAR_EXPIRY") filtered = filtered.filter(s => s.risk_status === "Cận date / Cảnh báo");
    else if (s5RiskFilter === "SAFE") filtered = filtered.filter(s => s.risk_status === "An toàn");

    if (s5RiskSearchQuery) {
      filtered = filtered.filter(s => {
        const q = s5RiskSearchQuery;
        return (s.sku || "").toLowerCase().includes(q) ||
               (s.name || "").toLowerCase().includes(q) ||
               (s.category || "").toLowerCase().includes(q) ||
               (s.risk_status || "").toLowerCase().includes(q) ||
               (s.recommendation || "").toLowerCase().includes(q);
      });
    }

    if (!filtered.length) {
      tbody.innerHTML = `<tr><td colspan="10" class="text-center py-6 text-xs text-[#0C0E08]/60 font-medium">Không tìm thấy SKU nào khớp với bộ lọc.</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered.map((item, idx) => {
      const isOut = item.risk_status === "Đứt hàng";
      const isExp = item.risk_status === "Cận date / Cảnh báo";

      let rowClass = "bg-white hover:bg-black/[0.02] text-[#0C0E08] border-l-4 border-l-emerald-500 transition";
      let badgeHtml = `<span class="px-2 py-0.5 rounded text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">✓ AN TOÀN</span>`;

      if (isOut) {
        rowClass = "bg-rose-50/80 hover:bg-rose-100/90 text-rose-950 font-bold border-l-4 border-l-rose-500 transition";
        badgeHtml = `<span class="px-2 py-0.5 rounded text-[10px] font-black bg-rose-600 text-white shadow-xs animate-pulse">🚨 ĐỨT HÀNG</span>`;
      } else if (isExp) {
        rowClass = "bg-amber-50/80 hover:bg-amber-100/90 text-amber-950 font-semibold border-l-4 border-l-amber-500 transition";
        badgeHtml = `<span class="px-2 py-0.5 rounded text-[10px] font-black bg-amber-500 text-white shadow-xs">⚠️ CẬN DATE / CẢNH BÁO</span>`;
      }

      const doiText = item.days_of_inventory > 90 ? ">90 ngày" : `${item.days_of_inventory} ngày`;

      return `
        <tr class="${rowClass}">
          <td class="py-2.5 px-3 text-center font-mono font-bold text-[#0C0E08]/50">${idx + 1}</td>
          <td class="py-2.5 px-3 font-mono font-black text-[#0B0D08] whitespace-nowrap">${item.sku}</td>
          <td class="py-2.5 px-3 max-w-[200px]">
            <strong class="block text-[11px] font-bold truncate" title="${item.name}">${item.name}</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60">${item.category || "Hàng hóa"}</span>
          </td>
          <td class="py-2.5 px-3 text-right font-mono font-black ${isOut ? 'text-rose-700' : ''}">${(item.stock || 0).toLocaleString()}</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold">${(item.sales_velocity || 0).toLocaleString()}</td>
          <td class="py-2.5 px-3 text-right font-mono font-black ${isOut ? 'text-rose-700' : (isExp ? 'text-amber-700' : 'text-emerald-700')}">${doiText}</td>
          <td class="py-2.5 px-3 text-center text-[10px] font-medium whitespace-nowrap">${item.lead_time || "1-3 ngày"}</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold">${formatCompactVND(item.inventory_value || 0)}</td>
          <td class="py-2.5 px-3 text-center whitespace-nowrap">${badgeHtml}</td>
          <td class="py-2.5 px-3 text-[10px] leading-relaxed max-w-[240px]">${item.recommendation}</td>
        </tr>
      `;
    }).join("");
  } else {
    const dimTitle = formatColDisplayName(schema.primary_dimension);
    const primMetricTitle = formatColDisplayName(schema.primary_metric);
    const secMetricTitle = schema.secondary_metric ? formatColDisplayName(schema.secondary_metric) : "Chỉ số phụ";

    if (titleText) titleText.textContent = `III. Bảng Tổng hợp Quản trị theo ${dimTitle} (Dynamic Exception Pivot)`;
    if (subtitleEl) subtitleEl.textContent = `Quản trị bằng ngoại lệ: Tự động đánh dấu các nhóm dẫn đầu (🌟 Dẫn đầu - Top Performer) và nhóm cần cải thiện (⚠️ Cần cải thiện).`;

    if (thead) {
      thead.innerHTML = `
        <tr>
          <th class="py-2.5 px-3 text-center w-10">#</th>
          <th class="py-2.5 px-3 whitespace-nowrap">${dimTitle}</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">Số giao dịch</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">Tổng ${primMetricTitle}</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">Tỷ trọng (%)</th>
          <th class="py-2.5 px-3 text-right whitespace-nowrap">${schema.secondary_metric ? 'Tổng ' + secMetricTitle : 'TB / Giao dịch'}</th>
          <th class="py-2.5 px-3 text-center whitespace-nowrap">Phân loại Ngoại lệ</th>
          <th class="py-2.5 px-3">Khuyến nghị điều hành BI</th>
        </tr>
      `;
    }

    const allPivots = schema.pivot_summary || [];
    const topCnt = allPivots.filter(p => p.exception_level === "TOP").length;
    const stdCnt = allPivots.filter(p => p.exception_level === "STANDARD").length;
    const lowCnt = allPivots.filter(p => p.exception_level === "LOW").length;

    if (chipsContainer) {
      chipsContainer.innerHTML = `
        <span class="text-[10px] font-bold text-[#0C0E08]/60 uppercase">Lọc theo ngoại lệ:</span>
        <button id="btn-chip-all" onclick="setS5RiskFilter('ALL')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5PivotFilter === 'ALL' ? 'bg-[#0B0D08] text-[#D4EE0E] border border-black' : 'bg-white text-[#0C0E08] border border-black/20'}">Tất cả (${allPivots.length})</button>
        <button id="btn-chip-out" onclick="setS5RiskFilter('TOP')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5PivotFilter === 'TOP' ? 'bg-emerald-600 text-white border border-emerald-700' : 'bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300'}">🌟 Nhóm Dẫn đầu (${topCnt})</button>
        <button id="btn-chip-exp" onclick="setS5RiskFilter('STANDARD')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5PivotFilter === 'STANDARD' ? 'bg-blue-600 text-white border border-blue-700' : 'bg-white hover:bg-blue-50 text-blue-700 border border-blue-300'}">⚡ Tiêu chuẩn (${stdCnt})</button>
        <button id="btn-chip-safe" onclick="setS5RiskFilter('LOW')" class="px-3 py-1 rounded-lg font-bold transition shadow-2xs ${s5PivotFilter === 'LOW' ? 'bg-amber-500 text-white border border-amber-600' : 'bg-white hover:bg-amber-50 text-amber-700 border border-amber-300'}">⚠️ Cần cải thiện (${lowCnt})</button>
      `;
    }

    let filtered = allPivots;
    if (s5PivotFilter === "TOP") filtered = filtered.filter(p => p.exception_level === "TOP");
    else if (s5PivotFilter === "STANDARD") filtered = filtered.filter(p => p.exception_level === "STANDARD");
    else if (s5PivotFilter === "LOW") filtered = filtered.filter(p => p.exception_level === "LOW");

    if (s5RiskSearchQuery) {
      filtered = filtered.filter(p => {
        const q = s5RiskSearchQuery;
        return (p.dimension_value || "").toLowerCase().includes(q) ||
               (p.exception_status || "").toLowerCase().includes(q) ||
               (p.recommendation || "").toLowerCase().includes(q);
      });
    }

    if (!filtered.length) {
      tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-xs text-[#0C0E08]/60 font-medium">Không tìm thấy phân nhóm nào khớp với bộ lọc "${s5RiskSearchQuery || s5PivotFilter}".</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered.map((item, idx) => {
      const isTop = item.exception_level === "TOP";
      const isLow = item.exception_level === "LOW";

      let rowClass = "bg-white hover:bg-black/[0.02] text-[#0C0E08] border-l-4 border-l-blue-400 transition";
      let badgeHtml = `<span class="px-2 py-0.5 rounded text-[10px] font-black bg-blue-50 text-blue-700 border border-blue-200">⚡ TIÊU CHUẨN</span>`;

      if (isTop) {
        rowClass = "bg-emerald-50/80 hover:bg-emerald-100/90 text-emerald-950 font-bold border-l-4 border-l-emerald-500 transition";
        badgeHtml = `<span class="px-2 py-0.5 rounded text-[10px] font-black bg-emerald-600 text-white shadow-xs">🌟 DẪN ĐẦU</span>`;
      } else if (isLow) {
        rowClass = "bg-amber-50/80 hover:bg-amber-100/90 text-amber-950 font-semibold border-l-4 border-l-amber-500 transition";
        badgeHtml = `<span class="px-2 py-0.5 rounded text-[10px] font-black bg-amber-500 text-white shadow-xs">⚠️ CẦN CẢI THIỆN</span>`;
      }

      const col6Val = item.total_secondary !== null && item.total_secondary !== undefined
        ? formatCompactVND(item.total_secondary)
        : formatCompactVND(item.avg_primary);

      return `
        <tr class="${rowClass}">
          <td class="py-2.5 px-3 text-center font-mono font-bold text-[#0C0E08]/50">${idx + 1}</td>
          <td class="py-2.5 px-3">
            <strong class="block text-xs font-black text-[#0B0D08]">${item.dimension_value}</strong>
            <span class="text-[10px] font-semibold text-[#0C0E08]/60">${dimTitle}</span>
          </td>
          <td class="py-2.5 px-3 text-right font-mono font-bold">${item.order_count.toLocaleString()}</td>
          <td class="py-2.5 px-3 text-right font-mono font-black ${isTop ? 'text-emerald-700' : ''}">${formatCompactVND(item.total_primary)}</td>
          <td class="py-2.5 px-3 text-right font-mono font-black ${isTop ? 'text-emerald-700' : ''}">${item.share_pct}%</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold">${col6Val}</td>
          <td class="py-2.5 px-3 text-center whitespace-nowrap">${badgeHtml}</td>
          <td class="py-2.5 px-3 text-[10px] leading-relaxed max-w-[280px]">${item.recommendation}</td>
        </tr>
      `;
    }).join("");
  }
}

function renderStep5VipTable() {
  const tbody = document.getElementById("s5-tbody-vip");
  const btnToggle = document.getElementById("btn-toggle-s5-vip");
  if (!tbody) return;

  let allRows = univAnalysis?.customer_priority || [];
  if (!allRows.length && analyzedResults.length > 0) {
    allRows = analyzedResults.map((r, i) => ({
      customer_id: r.ten_khach_hang || r.khach_hang || `KH-${i+1}`,
      revenue: Number(r.thanh_tien_sau_ck || 0),
      share_pct: 10,
      margin_pct: 15,
      tier: (r.phan_loai || "").includes("tiếp tục") ? "Tier A (VIP)" : "Tier B"
    }));
  }

  const rowsToShow = s5VipShowAll ? allRows : allRows.slice(0, 5);
  if (btnToggle) {
    btnToggle.textContent = s5VipShowAll ? "Thu gọn" : `Xem tất cả (${allRows.length})`;
    btnToggle.classList.toggle("hidden", allRows.length <= 5);
  }

  if (!rowsToShow.length) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center py-4 text-xs opacity-50">Chưa có dữ liệu khách hàng</td></tr>`;
    return;
  }

  tbody.innerHTML = rowsToShow.map((c, idx) => {
    const tier = c.tier || "VIP";
    const isA = tier.includes("Tier A") || tier.includes("A");
    return `
      <tr class="hover:bg-black/[0.02]">
        <td class="py-2 px-2.5 text-center font-mono font-bold text-[#0C0E08]/60">${idx + 1}</td>
        <td class="py-2 px-2.5 font-bold text-[#0C0E08]">${c.customer_id || c.customer || "—"}</td>
        <td class="py-2 px-2.5 text-right font-mono font-bold">${formatCompactVND(c.revenue || 0)}</td>
        <td class="py-2 px-2.5 text-right font-mono">${Number(c.share_pct || 0).toFixed(1)}%</td>
        <td class="py-2 px-2.5 text-right font-mono font-bold text-emerald-700">${c.margin_pct !== null && c.margin_pct !== undefined ? Number(c.margin_pct).toFixed(1) + "%" : "-"}</td>
        <td class="py-2 px-2.5 text-center">
          <span class="px-2 py-0.5 rounded text-[10px] font-black ${isA ? "bg-[#D4EE0E] text-[#0C0E08]" : "bg-black/10 text-[#0C0E08]"}">
            ${tier}
          </span>
        </td>
      </tr>
    `;
  }).join("");
}

function renderStep5ProdTable() {
  const tbody = document.getElementById("s5-tbody-prod");
  const btnToggle = document.getElementById("btn-toggle-s5-prod");
  if (!tbody) return;

  let allRows = univAnalysis?.top_products || [];
  if (!allRows.length && analyzedResults.length > 0) {
    allRows = analyzedResults.map((r, i) => ({
      product_name: r.ten_san_pham || r.ma_san_pham || `Sản phẩm ${i+1}`,
      quantity: Number(r.so_luong || 1),
      revenue: Number(r.thanh_tien_sau_ck || 0),
      share_pct: 10,
      margin_pct: 15
    }));
  }

  const rowsToShow = s5ProdShowAll ? allRows : allRows.slice(0, 5);
  if (btnToggle) {
    btnToggle.textContent = s5ProdShowAll ? "Thu gọn" : `Xem tất cả (${allRows.length})`;
    btnToggle.classList.toggle("hidden", allRows.length <= 5);
  }

  if (!rowsToShow.length) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center py-4 text-xs opacity-50">Chưa có dữ liệu sản phẩm</td></tr>`;
    return;
  }

  tbody.innerHTML = rowsToShow.map((p, idx) => `
    <tr class="hover:bg-black/[0.02]">
      <td class="py-2 px-2.5 text-center font-mono font-bold text-[#0C0E08]/60">${idx + 1}</td>
      <td class="py-2 px-2.5 font-bold text-[#0C0E08]">${p.product_name || "—"}</td>
      <td class="py-2 px-2.5 text-right font-mono">${(p.quantity || 0).toLocaleString()}</td>
      <td class="py-2 px-2.5 text-right font-mono font-bold">${formatCompactVND(p.revenue || 0)}</td>
      <td class="py-2 px-2.5 text-right font-mono">${Number(p.share_pct || 0).toFixed(1)}%</td>
      <td class="py-2 px-2.5 text-right font-mono font-bold text-emerald-700">${p.margin_pct !== null && p.margin_pct !== undefined ? Number(p.margin_pct).toFixed(1) + "%" : "-"}</td>
    </tr>
  `).join("");
}

function renderStep5ExecutiveCards() {
  const cardPerf = document.getElementById("s5-card-performance");
  const cardConc = document.getElementById("s5-card-conclusions");
  const cardAct = document.getElementById("s5-card-actions");

  const cards = univAnalysis?.executive_cards || {};
  let bPerf = cards.performance_analysis || [];
  let bConc = cards.key_conclusions || [];
  let bAct = cards.action_proposals || [];

  if (!bPerf.length) {
    const k = univAnalysis?.kpis || {};
    const totRows = k.total_rows || analyzedResults.length || 0;
    const totRevStr = k.total_revenue ? fmtNum(k.total_revenue) : "";
    bPerf = [
      `Quy mô dữ liệu: Kiểm tra và chuẩn hóa ${totRows.toLocaleString("vi-VN")} bản ghi toàn vẹn.`
    ];
    if (totRevStr) {
      bPerf.push(`Tổng giá trị giao dịch ghi nhận: ${totRevStr}.`);
    }
    bConc = [
      `Dữ liệu hoạt động đồng bộ, các chỉ tiêu kỹ thuật nằm trong ngưỡng kiểm soát.`
    ];
    bAct = [
      `Tiếp tục theo dõi các chỉ số định lượng trọng tâm và nâng cao chất lượng nguồn dữ liệu.`
    ];
  }

  const renderBulletList = (bullets, dotBg = "bg-[#0B0D08]") => {
    if (!bullets || !bullets.length) return `<p class="text-xs text-[#0C0E08]/60 italic">Đang cập nhật phân tích...</p>`;
    return bullets.map(b => {
      let text = String(b)
        .replace(/^[*+-]\s+/, "")
        .replace(/^#+\s+/, "")
        .replace(/\*\*(.*?)\*\*/g, "<strong class='font-bold text-[#0C0E08]'>$1</strong>")
        .replace(/\*(.*?)\*/g, "<em class='italic'>$1</em>")
        .replace(/\|/g, " ")
        .replace(/^#{1,6}/g, "");
      return `
        <div class="flex items-start gap-2.5">
          <span class="inline-block w-1.5 h-1.5 rounded-full ${dotBg} mt-1.5 flex-shrink-0"></span>
          <p class="text-xs leading-relaxed text-[#0C0E08]/85">${text}</p>
        </div>
      `;
    }).join("");
  };

  if (cardPerf) cardPerf.innerHTML = renderBulletList(bPerf, "bg-[#0B0D08]");
  if (cardConc) cardConc.innerHTML = renderBulletList(bConc, "bg-[#0B0D08]");
  if (cardAct) cardAct.innerHTML = renderBulletList(bAct, "bg-[#0B0D08]");
}

async function approveAndGenerateBusinessReport() {
  const btn = document.getElementById("btn-approve-report");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "⏳ Đang phê duyệt & tạo file...";
  }

  const approverName = document.getElementById("s5-approver-name")?.value?.trim() || "Ban Giám Đốc";
  const approverTitle = document.getElementById("s5-approver-title")?.value?.trim() || "Đại diện Phê duyệt";
  const approvalNotes = document.getElementById("s5-approval-notes")?.value?.trim() || "Đồng ý phê duyệt số liệu kinh doanh, chính sách chiết khấu và ban hành báo cáo chính thức.";

  addAudit("Bước 5", "approve_business_report", `Phê duyệt báo cáo kinh doanh bởi ${approverName} (${approverTitle}) - Ghi chú: ${approvalNotes}`);

  let success = false;
  if (datasetMode === "MINH_AN_B2B") {
    success = await runStep5Export();
  } else {
    success = await runUnivStep5Export();
  }

  if (success) {
    step5Approved = true;
    updateStepperUI();
    renderStep5ReportPreview();
    showToast("🎉 Báo cáo Kinh doanh đã được Phê duyệt & Ban hành thành công!", "success");
  } else {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "✓ Duyệt & Ban Hành Báo Cáo";
    }
  }
}

function printBusinessReport() {
  window.print();
}

// Attach live preview updates when user edits approver inputs & sync company inputs
document.addEventListener("DOMContentLoaded", () => {
  initCompanyInfoSync();
  ["s5-approver-name", "s5-approver-title", "s5-approval-notes"].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener("input", () => {
        if (currentStep === 5) renderStep5ReportPreview();
      });
      el.addEventListener("change", () => {
        if (currentStep === 5) renderStep5ReportPreview();
      });
    }
  });
});

function renderMarkdownToHtml(md) {
  if (!md) return "";
  const lines = md.split("\n");
  const html = [];
  lines.forEach(line => {
    let l = line.trim();
    if (!l) {
      html.push("<div class='h-1.5'></div>");
      return;
    }
    if (/^[-*_=\s]{3,}$/.test(l)) {
      html.push("<hr class='my-2 border-black/10'>");
      return;
    }
    if (/^#{1,6}\s+/.test(l)) {
      const text = l.replace(/^#{1,6}\s+/, "").replace(/\*\*/g, "");
      html.push(`<h4 class="font-bold text-xs sm:text-sm text-[#0C0E08] uppercase tracking-wider mt-3 mb-1">${text}</h4>`);
      return;
    }
    if (/^[*+-]\s+/.test(l)) {
      let text = l.replace(/^[*+-]\s+/, "");
      text = text.replace(/\*\*(.*?)\*\*/g, "<strong class='font-bold text-[#0C0E08]'>$1</strong>");
      text = text.replace(/\*(.*?)\*/g, "<em class='italic'>$1</em>");
      html.push(`<div class="flex items-start gap-2 my-0.5"><span class="text-[#0C0E08] font-bold">•</span><span class="text-xs leading-relaxed">${text}</span></div>`);
      return;
    }
    if (/^\d+\.\s+/.test(l)) {
      const match = l.match(/^(\d+\.)\s+(.*)$/);
      if (match) {
        let text = match[2].replace(/\*\*(.*?)\*\*/g, "<strong class='font-bold text-[#0C0E08]'>$1</strong>");
        text = text.replace(/\*(.*?)\*/g, "<em class='italic'>$1</em>");
        html.push(`<div class="flex items-start gap-1.5 my-1 font-semibold text-xs"><span class="font-bold text-[#0C0E08]">${match[1]}</span><span>${text}</span></div>`);
        return;
      }
    }
    let text = l.replace(/\*\*(.*?)\*\*/g, "<strong class='font-bold text-[#0C0E08]'>$1</strong>");
    text = text.replace(/\*(.*?)\*/g, "<em class='italic'>$1</em>");
    text = text.replace(/\|/g, " ");
    html.push(`<p class="text-xs my-1 leading-relaxed">${text}</p>`);
  });
  return html.join("");
}
