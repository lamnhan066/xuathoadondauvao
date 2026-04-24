const state = {
  invoices: [],
  selectedIds: new Set(),
  syncState: {
    lastSyncAt: null,
    lastDateFrom: null,
    lastDateTo: null,
    invoiceCount: 0,
    dataOrigin: ""
  },
  dataOrigin: "",
  productCatalog: {
    fileName: "",
    loadedAt: null,
    products: [],
    selections: {}
  },
  exportReview: {
    isOpen: false,
    pendingAction: "idle",
    entries: [],
    resolve: null
  },
  exportBusy: {
    active: false,
    label: "Xuất Tendoo"
  }
};

const TENDOO_TEMPLATE_HEADERS = [
  "Mã sản phẩm/nvl (*)",
  "Loại hàng",
  "Tên sản phẩm/nvl",
  "Đơn vị tính",
  "Số lượng (*)",
  "Đơn giá (*)",
  "Giảm giá sản phẩm %",
  "Giảm giá sản phẩm VNĐ",
  "Thuế GTGT"
];

const TENDOO_HEADER_KEYS = TENDOO_TEMPLATE_HEADERS.map(normalizeText);

const PRODUCT_CODE_HEADERS = buildNormalizedHeaderSet([
  "Mã sản phẩm(SKU)",
  "Mã sản phẩm",
  "Mã sản phẩm/nvl",
  "Mã sp",
  "Mã sp/nvl",
  "Mã hàng",
  "Mã hàng hóa",
  "Mã vật tư",
  "Mã vt",
  "SKU",
  "Code",
  "Product code",
  "Item code"
]);

const PRODUCT_NAME_HEADERS = buildNormalizedHeaderSet([
  "Tên sản phẩm",
  "Tên sản phẩm/nvl",
  "Tên sp",
  "Tên sp/nvl",
  "Tên hàng",
  "Tên hàng hóa",
  "Tên vật tư",
  "Tên vt",
  "Product name",
  "Name",
  "Item name"
]);

const PRODUCT_UNIT_HEADERS = buildNormalizedHeaderSet([
  "Tên đơn vị",
  "Đơn vị tính",
  "ĐVT",
  "Unit"
]);

const PRODUCT_TEMPLATE_SCAN_ROWS = 80;
const APP_STATE_STORAGE_KEY = "appState";

const el = {
  authStatus: document.getElementById("authStatus"),
  portalFlowStatus: document.getElementById("portalFlowStatus"),
  fetchStatus: document.getElementById("fetchStatus"),
  checkAuthBtn: document.getElementById("checkAuthBtn"),
  openPortalBtn: document.getElementById("openPortalBtn"),
  fetchBtn: document.getElementById("fetchBtn"),
  exportBtn: document.getElementById("exportBtn"),
  productTemplateInput: document.getElementById("productTemplateInput"),
  productTemplateBtn: document.getElementById("productTemplateBtn"),
  clearProductTemplateBtn: document.getElementById("clearProductTemplateBtn"),
  productTemplateStatus: document.getElementById("productTemplateStatus"),
  productResolutionList: document.getElementById("productResolutionList"),
  exportReviewDialog: document.getElementById("exportReviewDialog"),
  exportReviewSelectAll: document.getElementById("exportReviewSelectAll"),
  exportReviewList: document.getElementById("exportReviewList"),
  exportReviewSummary: document.getElementById("exportReviewSummary"),
  exportReviewStatus: document.getElementById("exportReviewStatus"),
  exportReviewCancelBtn: document.getElementById("exportReviewCancelBtn"),
  exportReviewConfirmBtn: document.getElementById("exportReviewConfirmBtn"),
  dateFrom: document.getElementById("dateFrom"),
  dateTo: document.getElementById("dateTo"),
  queryStatus: document.getElementById("queryStatus"),
  scoStatus: document.getElementById("scoStatus"),
  invoiceRows: document.getElementById("invoiceRows"),
  invoiceDetail: document.getElementById("invoiceDetail"),
  checkAll: document.getElementById("checkAll"),
  invoiceSourceBadge: document.getElementById("invoiceSourceBadge"),
  invoiceListBadge: document.getElementById("invoiceListBadge"),
  btnToday: document.getElementById("btnToday"),
  btnLast7: document.getElementById("btnLast7"),
  btnThisMonth: document.getElementById("btnThisMonth"),
  supportBtn: document.getElementById("supportBtn"),
  githubBtn: document.getElementById("githubBtn"),
  versionLabel: document.getElementById("versionLabel"),
  supportDialog: document.getElementById("supportDialog"),
  supportCloseBtn: document.getElementById("supportCloseBtn"),
  supportStatus: document.getElementById("supportStatus"),
  momoQr: document.getElementById("momoQr"),
  paypalOpenBtn: document.getElementById("paypalOpenBtn"),
  drawerToggleBtn: document.getElementById("drawerToggleBtn"),
  drawerBackdrop: document.getElementById("drawerBackdrop")
};

init();

async function init() {
  bindEvents();
  updateDrawerToggleVisibility();
  renderVersion();
  await loadCachedState();
  await checkAuth();
}

function renderVersion() {
  try {
    const version = (chrome && chrome.runtime && chrome.runtime.getManifest && chrome.runtime.getManifest().version) || "";
    if (el.versionLabel) {
      el.versionLabel.textContent = version ? `v${version}` : "";
      el.versionLabel.title = version ? `Phiên bản ${version}` : "";
    }
  } catch (e) {
    // ignore
  }
}

function bindEvents() {
  el.checkAuthBtn.addEventListener("click", checkAuth);
  el.openPortalBtn.addEventListener("click", async () => {
    try {
      await openPortal();
      setStatus(el.authStatus, "Đã mở cổng. Vui lòng đăng nhập.");
    } catch (error) {
      setStatus(el.authStatus, error.message, true);
    }
  });
  el.fetchBtn.addEventListener("click", fetchInvoices);
  el.exportBtn.addEventListener("click", exportSelected);
  el.productTemplateBtn.addEventListener("click", () => {
    el.productTemplateInput.click();
  });
  el.productTemplateInput.addEventListener("change", onProductTemplateSelected);
  el.clearProductTemplateBtn.addEventListener("click", clearProductCatalog);
  el.checkAll.addEventListener("change", toggleAll);
  el.exportReviewDialog.addEventListener("close", onExportReviewDialogClose);
  el.exportReviewDialog.addEventListener("cancel", onExportReviewDialogCancel);
  el.exportReviewDialog.addEventListener("click", onExportReviewDialogClick);
  el.exportReviewDialog.addEventListener("input", onExportReviewDialogFieldChange);
  el.exportReviewDialog.addEventListener("change", onExportReviewDialogFieldChange);
  el.exportReviewSelectAll.addEventListener("change", onExportReviewSelectAllChange);
  el.exportReviewCancelBtn.addEventListener("click", () => closeExportReviewDialog(false));
  el.exportReviewConfirmBtn.addEventListener("click", () => closeExportReviewDialog(true));

  el.btnToday?.addEventListener("click", () => setQuickDate("today"));
  el.btnLast7?.addEventListener("click", () => setQuickDate("last7"));
  el.btnThisMonth?.addEventListener("click", () => setQuickDate("thisMonth"));

  // Drawer sidebar handlers (small screens)
  el.drawerToggleBtn?.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleSidebarDrawer();
  });

  el.drawerBackdrop?.addEventListener("click", () => closeSidebarDrawer());

  // Close drawer with Escape
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closeSidebarDrawer();
    }
  });

  // Update visibility of the floating toggle on resize
  window.addEventListener('resize', updateDrawerToggleVisibility);

  // Support (Ủng hộ) dialog handlers
  el.supportBtn?.addEventListener("click", openSupportDialog);
  el.supportCloseBtn?.addEventListener("click", () => closeSupportDialog());
  el.paypalOpenBtn?.addEventListener("click", async () => {
    const url = "https://paypal.me/lamnhan066";
    try {
      await sendRuntimeMessage({ type: "OPEN_EXTERNAL_URL", payload: { url } });
    } catch (err) {
      try { window.open(url, "_blank"); } catch (e) { }
    }
  });
  el.githubBtn?.addEventListener("click", async () => {
    const url = "https://github.com/lamnhan066/xuathoadondauvao";
    try {
      await sendRuntimeMessage({ type: "OPEN_EXTERNAL_URL", payload: { url } });
    } catch (err) {
      try { window.open(url, "_blank"); } catch (e) { }
    }
  });
  el.supportDialog?.addEventListener("click", (e) => {
    if (e.target === el.supportDialog) {
      closeSupportDialog();
    }
  });

  // Note: floating drawer toggle is used; no in-sidebar close button
}

function toggleSidebarDrawer() {
  const sb = document.querySelector('.sidebar');
  const backdrop = el.drawerBackdrop;
  if (!sb) return;
  const opening = !sb.classList.contains('open');
  if (opening) {
    sb.classList.add('open');
    if (backdrop) {
      backdrop.hidden = false;
      // force reflow to allow transition
      void backdrop.offsetWidth;
      backdrop.classList.add('visible');
    }
  } else {
    closeSidebarDrawer();
  }
  updateDrawerToggleVisibility();
}

function closeSidebarDrawer() {
  const sb = document.querySelector('.sidebar');
  const backdrop = el.drawerBackdrop;
  if (!sb) return;
  sb.classList.remove('open');
  if (backdrop) {
    backdrop.classList.remove('visible');
    // hide after transition
    setTimeout(() => { backdrop.hidden = true; }, 220);
  }
  updateDrawerToggleVisibility();
}

function updateDrawerToggleVisibility() {
  const btn = el.drawerToggleBtn;
  const sb = document.querySelector('.sidebar');
  if (!btn) return;

  const isLarge = window.matchMedia('(min-width:1025px)').matches;
  if (isLarge) {
    btn.style.display = 'none';
    return;
  }

  // On small screens, show the button when the sidebar is closable (i.e., available as drawer)
  btn.style.display = 'inline-flex';
  const open = sb && sb.classList.contains('open');
  if (open) {
    btn.classList.add('open');
    btn.textContent = '←';
    // position the floating button to the right of the sidebar so it's always visible
    const sidebarWidth = getSidebarWidth();
    btn.style.left = `${sidebarWidth + 12}px`;
  } else {
    btn.classList.remove('open');
    btn.textContent = '→';
    btn.style.left = '12px';
  }
}

function getSidebarWidth() {
  const sb = document.querySelector('.sidebar');
  if (!sb) return 320;
  const w = sb.getBoundingClientRect().width || 320;
  return Math.min(w, Math.max(200, w));
}

function setQuickDate(type) {
  const now = new Date();
  let from = new Date();

  if (type === "today") {
    from = now;
  } else if (type === "last7") {
    from.setDate(now.getDate() - 7);
  } else if (type === "thisMonth") {
    from = new Date(now.getFullYear(), now.getMonth(), 1);
  }

  el.dateFrom.value = toInputDate(from);
  el.dateTo.value = toInputDate(now);
}

function setExportBusy(isBusy, label) {
  state.exportBusy.active = isBusy;
  state.exportBusy.label = label || "Xuất Tendoo";

  if (el.exportBtn) {
    el.exportBtn.disabled = isBusy;
    el.exportBtn.classList.toggle("is-loading", isBusy);
    const labelEl = el.exportBtn.querySelector(".label");
    if (labelEl) {
      labelEl.textContent = state.exportBusy.label;
    } else {
      el.exportBtn.textContent = state.exportBusy.label;
    }
  }

  document.body.classList.toggle("export-busy", isBusy);
}

function setDefaultDateRange() {
  const now = new Date();
  const before = new Date(now);
  before.setDate(now.getDate() - 30);

  el.dateFrom.value = toInputDate(before);
  el.dateTo.value = toInputDate(now);
}

function toInputDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

async function checkAuth() {
  setStatus(el.authStatus, "Đang kiểm tra luồng đăng nhập trên cổng...");
  setStatus(el.portalFlowStatus, "Đang đọc trạng thái cổng...");

  try {
    const response = await sendRuntimeMessage({ type: "CHECK_AUTH" });

    if (!response.ok) {
      throw new Error(response.error || "Lỗi kiểm tra đăng nhập không xác định");
    }

    renderPortalFlowStatus(response.portalFlow, response.phase);

    if (response.isAuthenticated) {
      setStatus(
        el.authStatus,
        response.reason || "Phiên hợp lệ. Extension sẽ dùng session hiện tại."
      );
      return response;
    }

    const serverMsg =
      response.reason || "Chưa đăng nhập. Vui lòng đăng nhập trên trang hoadondientu.gdt.gov.vn";
    setStatus(el.authStatus, serverMsg, true);
    return response;
  } catch (error) {
    setStatus(el.authStatus, `Không thể kiểm tra session: ${error.message}`, true);
    setStatus(el.portalFlowStatus, "Không đọc được trạng thái cổng.", true);
    return null;
  }
}

async function openPortal() {
  const response = await sendRuntimeMessage({ type: "OPEN_LOGIN_POPUP" });
  if (!response?.ok) {
    throw new Error(response?.error || "Không thể mở cửa sổ đăng nhập");
  }
}

async function fetchInvoices() {
  const auth = await checkAuth();
  if (!auth?.isAuthenticated) {
    setStatus(el.fetchStatus, "Cần đăng nhập hoặc mở đúng màn hình tra cứu hóa đơn trước khi tải dữ liệu.", true);
    return;
  }

  const dateFromInput = el.dateFrom.value;
  const dateToInput = el.dateTo.value;

  if (!dateFromInput || !dateToInput) {
    setStatus(el.fetchStatus, "Vui lòng nhập đầy đủ ngày từ-đến.", true);
    return;
  }

  const payload = {
    dateFrom: toApiDate(dateFromInput, false),
    dateTo: toApiDate(dateToInput, true),
    queryStatus: Number(el.queryStatus.value || 5),
    scoStatus: Number(el.scoStatus.value || 8)
  };

  setStatus(el.fetchStatus, "Đang tải dữ liệu từ Hóa đơn điện tử và Hóa đơn có mã khởi tạo từ máy tính tiền...");

  try {
    const response = await sendRuntimeMessage({
      type: "FETCH_INVOICES",
      payload
    });

    if (!response.ok) {
      throw new Error(response.error || "Lỗi tải dữ liệu không xác định");
    }

    state.invoices = response.data.invoices || [];
    state.dataOrigin = response.data.servedFromLocal ? "local" : "network";
    state.syncState = {
      ...state.syncState,
      lastSyncAt: response.data.lastSyncAt || new Date().toISOString(),
      lastDateFrom: response.data.requestedDateFrom || payload.dateFrom,
      lastDateTo: response.data.requestedDateTo || payload.dateTo,
      invoiceCount: response.data.stats?.total || state.invoices.length,
      dataOrigin: state.dataOrigin
    };
    state.selectedIds.clear();

    renderRows();
    renderInvoiceSourceBadge();

    const { queryCount, scoCount, total } = response.data.stats;
    const errorMessages = Array.isArray(response.data.errors) ? response.data.errors.filter(Boolean) : [];
    const fallbackNote = response.data.servedFromLocal
      ? " Đang hiển thị dữ liệu cục bộ do lỗi kết nối hoặc không có phản hồi từ cổng."
      : "";

    if (errorMessages.length) {
      setStatus(
        el.fetchStatus,
        `Tải dữ liệu thất bại.${fallbackNote} Lỗi: ${errorMessages.join("; ")}`,
        true
      );
    } else {
      setStatus(
        el.fetchStatus,
        `Tải xong: Hóa đơn điện tử=${queryCount}, Hóa đơn có mã khởi tạo từ máy tính tiền=${scoCount}, tổng sau gộp=${total}.${fallbackNote}`
      );
    }
  } catch (error) {
    setStatus(el.fetchStatus, `Tải dữ liệu thất bại: ${error.message}`, true);
  }
}

function renderRows() {
  if (!state.invoices.length) {
    el.invoiceRows.innerHTML =
      '<tr><td colspan="7" class="empty-state">Không có dữ liệu trong khoảng ngày đã chọn.</td></tr>';
    return;
  }

  const html = state.invoices
    .map((item) => {
      const id = getInvoiceId(item);
      const checked = state.selectedIds.has(id) ? "checked" : "";
      const date = formatDateTime(item.tdlap || item.ntao || item.ncnhat);
      const total = toMoney(item.tgtttbso);
      const seller = escapeHtml(item.nbten || "-");
      const invoiceNo = escapeHtml(`${item.khhdon || ""}-${item.shdon ?? ""}`);
      const rawSource = item._source || "-";
      const sourceLabel =
        rawSource === "query"
          ? "HĐ điện tử"
          : rawSource === "sco-query"
            ? "HĐ máy tính tiền"
            : rawSource;
      const source = escapeHtml(sourceLabel);

      return `
        <tr data-id="${escapeHtml(id)}" class="invoice-row">
          <td class="col-check"><input type="checkbox" class="row-check" data-id="${escapeHtml(id)}" ${checked} /></td>
          <td class="col-source">${source}</td>
          <td class="col-no">${invoiceNo}</td>
          <td class="col-seller">${seller}</td>
          <td class="col-date">${escapeHtml(date)}</td>
          <td class="col-amount">${escapeHtml(total)}</td>
          <td class="col-action"><span class="view-indicator">→</span></td>
        </tr>
      `;
    })
    .join("");

  el.invoiceRows.innerHTML = html;

  el.invoiceRows.querySelectorAll(".invoice-row").forEach((row) => {
    row.addEventListener("click", (e) => {
      if (e.target.classList.contains("row-check") || e.target.type === "checkbox") {
        return;
      }
      onViewDetail(row.dataset.id);
    });
  });

  el.invoiceRows.querySelectorAll(".row-check").forEach((checkbox) => {
    checkbox.addEventListener("change", onRowCheck);
  });
}

function onRowCheck(event) {
  const id = event.target.dataset.id;
  if (!id) return;

  if (event.target.checked) {
    state.selectedIds.add(id);
  } else {
    state.selectedIds.delete(id);
  }

  el.checkAll.checked = state.selectedIds.size === state.invoices.length;
}

async function onViewDetail(id) {
  if (!id) return;

  const invoice = state.invoices.find((item) => getInvoiceId(item) === id);
  if (!invoice) return;

  // Highlight active row
  el.invoiceRows.querySelectorAll(".invoice-row").forEach((row) => {
    row.classList.toggle("active", row.dataset.id === id);
  });

  try {
    el.invoiceDetail.innerHTML = '<div class="empty-detail"><span class="btn-spinner" style="display:block;border-top-color:var(--brand)"></span><p>Đang tải chi tiết...</p></div>';

    // Smooth scroll to detail pane
    el.invoiceDetail.closest('.detail-pane')?.scrollIntoView({ behavior: 'smooth', block: 'start' });

    const detail = await loadInvoiceDetail(invoice);
    el.invoiceDetail.innerHTML = renderDetailView(detail);
  } catch (error) {
    el.invoiceDetail.innerHTML = renderDetailView(invoice);
    console.error("Không tải được chi tiết hóa đơn:", error);
  }
}

function toggleAll(event) {
  const checked = event.target.checked;

  state.selectedIds.clear();
  if (checked) {
    state.invoices.forEach((item) => state.selectedIds.add(getInvoiceId(item)));
  }

  el.invoiceRows.querySelectorAll(".row-check").forEach((input) => {
    input.checked = checked;
  });
}

async function exportSelected() {
  const selected = state.invoices.filter((item) =>
    state.selectedIds.has(getInvoiceId(item))
  );

  if (!selected.length) {
    setStatus(el.fetchStatus, "Vui lòng chọn ít nhất 1 hóa đơn để xuất.", true);
    return;
  }

  if (!window.XLSX) {
    setStatus(el.fetchStatus, "Thiếu thư viện XLSX trong extension.", true);
    return;
  }

  // --- Mới: tải chi tiết trước để tránh lỗi dự phòng khi rà soát ---
  const detailInvoices = [];
  setExportBusy(true, `Đang tải ${selected.length} hóa đơn để xuất...`);

  try {
    for (const [index, invoice] of selected.entries()) {
      setExportBusy(true, `Đang tải hóa đơn ${index + 1}/${selected.length}...`);
      const detailInvoice = await loadInvoiceDetail(invoice).catch(() => invoice);
      detailInvoices.push({ sourceInvoice: invoice, detailInvoice });
    }
  } catch (error) {
    setExportBusy(false);
    setStatus(el.fetchStatus, `Không thể tải dữ liệu hóa đơn để xuất: ${error.message}`, true);
    return;
  }

  // --- Now perform audits on the DETAILED data ---
  const flatDetailInvoices = detailInvoices.map((item) => item.detailInvoice);
  const productAudit = collectProductResolutionNeeds(flatDetailInvoices, state.productCatalog);

  if (!state.productCatalog.products.length) {
    setStatus(
      el.fetchStatus,
      "Chưa tải mẫu BC_San_Pham từ Tendoo. Tiện ích sẽ xuất bằng mã tự sinh nếu cần (khuyến nghị: tải mẫu để tự động xuất mã sản phẩm chính xác)."
    );
    renderProductCatalogStatus(productAudit);
  }

  renderProductCatalogStatus(productAudit);

  if (productAudit.ambiguous.length > 0) {
    setExportBusy(false);
    setStatus(
      el.fetchStatus,
      "Có tên sản phẩm trùng trong mẫu. Hãy chọn đúng mã ở khối Mẫu sản phẩm trước khi xuất.",
      true
    );
    return;
  }

  const exportReviewEntries = state.productCatalog.products.length
    ? collectExportReviewNeeds(flatDetailInvoices, state.productCatalog)
    : [];
  let exportReview = null;

  if (exportReviewEntries.length > 0) {
    setExportBusy(false);
    setStatus(
      el.fetchStatus,
      `Có ${exportReviewEntries.length} dòng chưa xác định được mã sản phẩm. Hãy rà soát trong hộp thoại trước khi xuất.`
    );

    exportReview = await openExportReviewDialog(exportReviewEntries);
    if (!exportReview) {
      setStatus(el.fetchStatus, "Đã hủy xuất để rà soát lại các dòng chưa xác định mã sản phẩm.");
      return;
    }

    setExportBusy(true, "Đang tạo file XLSX Tendoo...");
  }

  const includedReviewCount = exportReview
    ? exportReview.entries.filter((item) => item.include).length
    : 0;

  if (exportReviewEntries.length > 0) {
    if (includedReviewCount > 0) {
      setStatus(
        el.fetchStatus,
        `Đã xác nhận ${includedReviewCount}/${exportReviewEntries.length} dòng chưa xác định mã để tiếp tục xuất.`
      );
    } else {
      setStatus(
        el.fetchStatus,
        "Bạn đã loại bỏ toàn bộ các dòng chưa xác định mã. Chỉ các dòng còn lại sẽ được xuất."
      );
    }
  }

  setStatus(
    el.fetchStatus,
    selected.length === 1
      ? "Đang tạo file XLSX Tendoo..."
      : `Đang tạo ${selected.length} file XLSX Tendoo riêng biệt...`
  );

  try {
    const savedFiles = [];

    for (const [index, item] of detailInvoices.entries()) {
      const { sourceInvoice, detailInvoice } = item;
      const filename = await exportSingleInvoiceToXlsx(
        detailInvoice,
        sourceInvoice,
        index,
        detailInvoices.length,
        state.productCatalog,
        exportReview
      );

      if (!filename) {
        continue;
      }

      savedFiles.push(filename);

      if (selected.length > 1) {
        setStatus(
          el.fetchStatus,
          `Đã xuất ${savedFiles.length}/${selected.length} file: ${filename}`
        );
      }
    }

    if (!savedFiles.length) {
      setStatus(
        el.fetchStatus,
        "Không có dòng nào được xuất sau khi loại bỏ các dòng chưa xác định mã.",
        true
      );
      return;
    }

    setStatus(
      el.fetchStatus,
      selected.length === 1
        ? `Đã tạo file XLSX: ${savedFiles[0]}.`
        : `Đã tạo xong ${savedFiles.length} file XLSX riêng biệt.`
    );
  } catch (error) {
    setStatus(el.fetchStatus, `Xuất XLSX thất bại: ${error.message}`, true);
  } finally {
    setExportBusy(false);
  }
}

async function exportSingleInvoiceToXlsx(
  detailInvoice,
  sourceInvoice,
  exportIndex,
  totalCount,
  productCatalog,
  exportReview
) {
  const exportRows = buildTendooRows([detailInvoice], productCatalog, exportReview);
  if (!exportRows.length) {
    return null;
  }

  const workbook = createTendooWorkbook(exportRows);

  const xlsxBytes = window.XLSX.write(workbook, {
    type: "array",
    bookType: "xlsx",
    cellStyles: true
  });

  const blob = new Blob([xlsxBytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const filename = buildInvoiceFilename(sourceInvoice, exportIndex, totalCount);

  try {
    await chrome.downloads.download({
      url,
      filename,
      saveAs: totalCount === 1
    });
    return filename;
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }
}

function buildInvoiceFilename(invoice, exportIndex, totalCount) {
  const invoiceNo = buildInvoiceDisplayNumber(invoice);
  const sellerName = sanitizeFilePart(invoice?.nbten || invoice?.nmten || "nguoi-ban");
  const issuedAt = formatInvoiceTimestampForFile(invoice?.tdlap || invoice?.ncnhat || invoice?.ntao);
  const invoicePart = sanitizeFilePart(invoiceNo || `hoa-don-${exportIndex + 1}`);
  const sequencePart = String(exportIndex + 1).padStart(Math.max(2, String(totalCount).length), "0");

  return `${sequencePart}-${invoicePart}-${sellerName}-${issuedAt}.xlsx`;
}

function createTendooWorkbook(rows) {
  const worksheetData = [
    TENDOO_TEMPLATE_HEADERS,
    ...rows.map((row) => [
      row.itemCode,
      row.itemType,
      row.itemName,
      row.unit,
      row.quantity,
      row.price,
      row.discountPercent,
      row.discountVnd,
      row.vat
    ])
  ];

  const worksheet = window.XLSX.utils.aoa_to_sheet(worksheetData);
  worksheet["!cols"] = [
    { wch: 24 },
    { wch: 14 },
    { wch: 34 },
    { wch: 14 },
    { wch: 12 },
    { wch: 16 },
    { wch: 18 },
    { wch: 18 },
    { wch: 14 }
  ];

  const workbook = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(workbook, worksheet, "Tendoo");
  return workbook;
}

function buildInvoiceDisplayNumber(invoice) {
  const invoiceNo = [invoice?.khhdon, invoice?.shdon].filter(Boolean).join("-");
  if (invoiceNo) {
    return invoiceNo;
  }

  if (invoice?.mhdon) {
    return invoice.mhdon;
  }

  return "hoa-don";
}

function formatInvoiceTimestampForFile(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return formatCompactTimestamp(new Date());
  }

  return formatCompactTimestamp(date);
}

function formatCompactTimestamp(date) {
  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");

  return `${yyyy}${mm}${dd}-${hh}${mi}${ss}`;
}

function sanitizeFilePart(value) {
  return (
    String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d")
      .replace(/[^a-zA-Z0-9._-]+/g, "-")
      .replace(/-{2,}/g, "-")
      .replace(/^[-_.]+|[-_.]+$/g, "")
      .slice(0, 120) || "hoa-don"
  );
}

async function loadInvoiceDetail(invoice) {
  try {
    const response = await sendRuntimeMessage({
      type: "FETCH_INVOICE_DETAIL",
      payload: {
        nbmst: invoice?.nbmst,
        khhdon: invoice?.khhdon,
        shdon: invoice?.shdon,
        khmshdon: invoice?.khmshdon,
        source: invoice?._source === "sco-query" ? "sco-query" : "query"
      }
    });

    if (response?.ok && response.data) {
      // If we got data through background fetch, it's from the network
      state.detailOrigin = "network";
      renderInvoiceSourceBadge();

      // Attempt to cache the fetched detail into local DB via background
      try {
        sendRuntimeMessage({ type: "UPSERT_INVOICE", payload: { invoice: response.data } }).catch(() => { });
      } catch (e) {
        // ignore caching errors
      }

      return response.data;
    }

    // Fallback to local invoice when response not ok or missing data:
    state.detailOrigin = "local";
    renderInvoiceSourceBadge();

    try {
      const cachedResp = await sendRuntimeMessage({ type: "GET_CACHED_INVOICE", payload: { invoice } }).catch(() => null);
      if (cachedResp?.ok && cachedResp.data) {
        return cachedResp.data;
      }
    } catch (e) {
      // ignore
    }

    return invoice;
  } catch (err) {
    // Network or runtime error: show local invoice as fallback
    state.detailOrigin = "local";
    renderInvoiceSourceBadge();

    try {
      const cachedResp = await sendRuntimeMessage({ type: "GET_CACHED_INVOICE", payload: { invoice } }).catch(() => null);
      if (cachedResp?.ok && cachedResp.data) {
        return cachedResp.data;
      }
    } catch (e) {
      // ignore
    }

    return invoice;
  }
}

function buildTendooRows(invoices, productCatalog = state.productCatalog, exportReview = null) {
  const rows = [];
  const reviewMap = buildExportReviewDecisionMap(exportReview);

  invoices.forEach((invoice, invoiceIndex) => {
    const exportEntries = buildInvoiceExportEntries(invoice, invoiceIndex, productCatalog);

    exportEntries.forEach((entry) => {
      const reviewDecision = reviewMap.get(entry.key);
      if (reviewDecision && !reviewDecision.include) {
        return;
      }

      rows.push(reviewDecision?.code ? { ...entry.row, itemCode: reviewDecision.code } : entry.row);
    });
  });

  return rows;
}

function collectExportReviewNeeds(invoices, productCatalog = state.productCatalog) {
  const reviewEntries = [];

  invoices.forEach((invoice, invoiceIndex) => {
    buildInvoiceExportEntries(invoice, invoiceIndex, productCatalog).forEach((entry) => {
      if (entry.needsReview) {
        reviewEntries.push(entry);
      }
    });
  });

  return reviewEntries;
}

function buildExportReviewDecisionMap(exportReview) {
  const reviewMap = new Map();

  (exportReview?.entries || []).forEach((entry) => {
    reviewMap.set(entry.key, {
      include: entry.include !== false,
      code: normalizeCellText(entry.code || entry.currentCode || "")
    });
  });

  return reviewMap;
}

function buildInvoiceExportEntries(invoice, invoiceIndex, productCatalog = state.productCatalog) {
  const lineItems = extractInvoiceLineItems(invoice);

  if (lineItems.length === 0) {
    return [buildFallbackInvoiceExportEntry(invoice, invoiceIndex, productCatalog)];
  }

  return lineItems.map((lineItem, lineIndex) => buildLineItemExportEntry(invoice, lineItem, invoiceIndex, lineIndex, productCatalog));
}

function buildLineItemExportEntry(invoice, lineItem, invoiceIndex, lineIndex, productCatalog) {
  const itemName = extractInvoiceProductName(lineItem) || invoice.thdon || invoice.tlhdon || "Hàng hóa từ hóa đơn";
  const resolvedProduct = resolveProductCatalogEntry(itemName, productCatalog);
  const fallbackCode = buildGeneratedItemCode(invoice, invoiceIndex, lineIndex);
  const currentCode =
    resolvedProduct?.code ||
    pickFirstText(lineItem, ["mhhhoa", "mhang", "mahang", "msp", "msanpham", "sku", "code"]) ||
    fallbackCode;

  return {
    key: buildExportReviewKey(invoice, invoiceIndex, lineIndex, "line"),
    invoiceId: getInvoiceId(invoice) || `${invoiceIndex + 1}`,
    invoiceLabel: buildInvoiceDisplayNumber(invoice) || `Hóa đơn ${invoiceIndex + 1}`,
    lineIndex,
    rowType: "line",
    itemName,
    resolvedCode: resolvedProduct?.code || "",
    currentCode,
    needsReview: !resolvedProduct?.code,
    row: buildTendooRowFromLineItem(invoice, lineItem, invoiceIndex, lineIndex, productCatalog, currentCode)
  };
}

function buildFallbackInvoiceExportEntry(invoice, invoiceIndex, productCatalog) {
  const itemName = invoice.thdon || invoice.tlhdon || "Hóa đơn mua vào";
  const resolvedProduct = resolveProductCatalogEntry(itemName, productCatalog);
  const fallbackCode = `INV-${invoiceIndex + 1}`;
  const currentCode = resolvedProduct?.code || fallbackCode;

  return {
    key: buildExportReviewKey(invoice, invoiceIndex, 0, "fallback"),
    invoiceId: getInvoiceId(invoice) || `${invoiceIndex + 1}`,
    invoiceLabel: buildInvoiceDisplayNumber(invoice) || `Hóa đơn ${invoiceIndex + 1}`,
    lineIndex: 0,
    rowType: "fallback",
    itemName,
    resolvedCode: resolvedProduct?.code || "",
    currentCode,
    needsReview: !resolvedProduct?.code,
    row: buildFallbackInvoiceRow(invoice, invoiceIndex, productCatalog, currentCode)
  };
}

function buildExportReviewKey(invoice, invoiceIndex, lineIndex, rowType) {
  const invoiceId = getInvoiceId(invoice) || `invoice-${invoiceIndex + 1}`;
  return `${invoiceId}::${rowType}::${lineIndex}`;
}

function buildTendooRowFromLineItem(
  invoice,
  lineItem,
  invoiceIndex,
  lineIndex,
  productCatalog,
  codeOverride
) {
  const baseRow = mapLineItemToTendooRow(invoice, lineItem, invoiceIndex, lineIndex, productCatalog);
  return {
    ...baseRow,
    itemCode: normalizeCellText(codeOverride || baseRow.itemCode)
  };
}

function extractInvoiceLineItems(invoice) {
  const directCandidates = [
    invoice?.hdhhdvu,
    invoice?.dshhdvu,
    invoice?.hhdvu,
    invoice?.hhdvus,
    invoice?.hanghoas,
    invoice?.hangHoas,
    invoice?.dsHangHoa,
    invoice?.chiTietHangHoa,
    invoice?.cthhdvu
  ];

  for (const candidate of directCandidates) {
    if (Array.isArray(candidate) && candidate.length > 0) {
      return candidate;
    }
  }

  // Fallback: detect arrays whose key suggests item-level details.
  for (const [key, value] of Object.entries(invoice || {})) {
    if (!Array.isArray(value) || value.length === 0) {
      continue;
    }

    const normalizedKey = normalizeText(key);
    if (/(hhdvu|hang|chi tiet|item|detail)/.test(normalizedKey)) {
      return value;
    }
  }

  return [];
}

function mapLineItemToTendooRow(invoice, lineItem, invoiceIndex, lineIndex, productCatalog) {
  const itemName =
    extractInvoiceProductName(lineItem) || invoice.thdon || invoice.tlhdon || "Hàng hóa từ hóa đơn";
  const resolvedProduct = resolveProductCatalogEntry(itemName, productCatalog);
  const itemCode =
    resolvedProduct?.code ||
    pickFirstText(lineItem, ["mhhhoa", "mhang", "mahang", "msp", "msanpham", "sku", "code"]) ||
    buildGeneratedItemCode(invoice, invoiceIndex, lineIndex);

  const unit =
    pickFirstText(lineItem, ["dvtinh", "dvt", "dvtien", "donvitinh", "unit"]) || "Cái";

  const quantity = pickFirstNumber(lineItem, ["sluong", "soluong", "qty", "quantity", "kluong"], 1);

  const price = pickFirstNumber(lineItem, ["dgia", "dongia", "dgiacthue", "price"], null);
  const fallbackPrice = toNumber(invoice.tgtcthue || invoice.tgtttbso || 0);

  const discountPercent = pickFirstNumber(
    lineItem,
    ["tlckhau", "ptckhau", "ptgiam", "discountpercent"],
    ""
  );

  const discountVnd = pickFirstNumber(
    lineItem,
    ["stckhau", "tienckhau", "giamgia", "discountamount"],
    ""
  );

  const vat = normalizeVat(invoice, lineItem);

  return {
    itemCode,
    // Leave itemType empty for Tendoo export (and preview)
    itemType: "",
    itemName,
    unit,
    quantity,
    price: price === null ? fallbackPrice : price,
    discountPercent: formatPercentDisplayOrBlankIfZero(discountPercent),
    discountVnd: formatBlankIfZero(discountVnd),
    vat: formatPercentDisplay(vat)
  };
}

function buildFallbackInvoiceRow(invoice, index, productCatalog, codeOverride) {
  const itemName = invoice.thdon || invoice.tlhdon || "Hóa đơn mua vào";
  const resolvedProduct = resolveProductCatalogEntry(itemName, productCatalog);
  const discountVnd = toNumber(invoice.ttcktmai || 0);

  return {
    itemCode: normalizeCellText(codeOverride || resolvedProduct?.code || `INV-${index + 1}`),
    itemType: "",
    itemName,
    unit: "Cái",
    quantity: 1,
    price: toNumber(invoice.tgtcthue || invoice.tgtttbso || 0),
    discountPercent: "",
    discountVnd: formatBlankIfZero(discountVnd),
    vat: normalizeVat(invoice)
  };
}

function buildGeneratedItemCode(invoice, invoiceIndex, lineIndex) {
  const invoiceNo = [invoice?.khhdon, invoice?.shdon].filter(Boolean).join("-");
  if (invoiceNo) {
    return `${invoiceNo}-${lineIndex + 1}`;
  }

  return `INV-${invoiceIndex + 1}-${lineIndex + 1}`;
}

function extractInvoiceProductName(lineItem) {
  return (
    pickFirstText(lineItem, ["thhdvu", "tenhang", "tenhh", "tenhhdvu", "ten", "name"]) ||
    ""
  );
}

function collectProductResolutionNeeds(invoices, productCatalog) {
  const seenNames = new Set();
  const ambiguous = [];
  const missing = [];
  const noCode = [];

  invoices.forEach((invoice, invoiceIndex) => {
    const lineItems = extractInvoiceLineItems(invoice);

    if (lineItems.length === 0) {
      const fallbackName = normalizeProductName(invoice.thdon || invoice.tlhdon || "Hóa đơn mua vào");
      if (!seenNames.has(fallbackName)) {
        seenNames.add(fallbackName);
        const match = findProductMatches(invoice.thdon || invoice.tlhdon || "Hóa đơn mua vào", productCatalog);
        match.source = buildInvoiceDisplayNumber(invoice);
        if (match.matches.length > 1 && !match.selection) {
          ambiguous.push(match);
        } else if (match.matches.length > 0 && !match.selection?.code) {
          noCode.push({
            ...match,
            source: buildInvoiceDisplayNumber(invoice),
            previewRow: buildFallbackInvoiceRow(invoice, invoiceIndex, productCatalog)
          });
        } else if (match.matches.length === 0) {
          missing.push(match);
        }
      }
      return;
    }

    lineItems.forEach((lineItem, lineIndex) => {
      const productName = extractInvoiceProductName(lineItem);
      const normalizedName = normalizeProductName(productName);
      if (!productName || seenNames.has(normalizedName)) {
        return;
      }

      seenNames.add(normalizedName);
      const match = findProductMatches(productName, productCatalog);
      match.source = buildInvoiceDisplayNumber(invoice);
      if (match.matches.length > 1 && !match.selection) {
        ambiguous.push(match);
      } else if (match.matches.length > 0 && !match.selection?.code) {
        noCode.push({
          ...match,
          source: buildInvoiceDisplayNumber(invoice),
          previewRow: mapLineItemToTendooRow(invoice, lineItem, invoiceIndex, lineIndex, productCatalog)
        });
      } else if (match.matches.length === 0) {
        missing.push(match);
      }
    });
  });

  return {
    ambiguous,
    missing,
    noCode
  };
}

function findProductMatches(productName, productCatalog) {
  const normalizedName = normalizeProductName(productName);

  // Exact candidates first
  let candidates = getProductCandidatesForName(normalizedName, productCatalog);

  // If no exact match, try more permissive substring matches to avoid false negatives
  if (candidates.length === 0 && normalizedName) {
    const lowered = normalizedName;
    const fuzzy = (productCatalog?.products || []).filter((p) => {
      const pn = String(p.normalizedName || "");
      return pn.includes(lowered) || lowered.includes(pn);
    });

    // Deduplicate fuzzy results (by code + name)
    const seen = new Set();
    candidates = [];
    for (const f of fuzzy) {
      const key = `${f.normalizedName}::${f.code}`;
      if (!seen.has(key)) {
        seen.add(key);
        candidates.push(f);
      }
    }
  }

  const selectedCode = productCatalog.selections?.[normalizedName] || "";
  const selection = candidates.find((candidate) => candidate.code === selectedCode) || null;

  return {
    name: productName,
    normalizedName,
    matches: candidates,
    selection
  };
}

function resolveProductCatalogEntry(productName, productCatalog) {
  const match = findProductMatches(productName, productCatalog);

  if (match.selection) {
    return match.selection;
  }

  if (match.matches.length === 1) {
    return match.matches[0];
  }

  return null;
}

function getProductCandidatesForName(normalizedName, productCatalog) {
  if (!normalizedName || !Array.isArray(productCatalog?.products)) {
    return [];
  }

  return productCatalog.products.filter((product) => product.normalizedName === normalizedName);
}

function pickFirstText(record, keys) {
  for (const key of keys) {
    const value = record?.[key];
    if (value === null || value === undefined) {
      continue;
    }

    const text = String(value).trim();
    if (text) {
      return text;
    }
  }

  return "";
}

function pickFirstNumber(record, keys, defaultValue) {
  for (const key of keys) {
    const value = record?.[key];
    const parsed = parseOptionalNumber(value);
    if (parsed !== null) {
      return parsed;
    }
  }

  return defaultValue;
}

function parseOptionalNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value === "string") {
    const cleaned = value.trim().replace(/%$/, "").replaceAll(",", "");
    if (!cleaned) {
      return null;
    }

    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  return null;
}

function resolveItemType(invoice, lineItem) {
  const lineType = normalizeText(
    pickFirstText(lineItem, ["lhdvu", "loaihang", "itemtype", "type"])
  );
  if (lineType.includes("dich vu") || lineType === "dv") {
    return "Dịch vụ";
  }

  const invoiceTypeText = normalizeText(invoice?.thdon || invoice?.tlhdon || "");
  if (invoiceTypeText.includes("dich vu")) {
    return "Dịch vụ";
  }

  return "Sản phẩm";
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeHeaderDisplay(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function cloneValue(value) {
  if (value === null || value === undefined) {
    return value;
  }

  return JSON.parse(JSON.stringify(value));
}

function normalizeVat(invoice, lineItem = null) {
  const lineTax = pickFirstText(lineItem, ["ltsuat", "tsuat", "thuesuat", "vat", "taxrate"]);
  if (lineTax) {
    return normalizeVatText(lineTax);
  }

  const tax = invoice?.thttltsuat?.[0]?.tsuat;
  if (tax) {
    return normalizeVatText(tax);
  }

  if (toNumber(invoice?.tgtthue) === 0) {
    return "Không chịu thuế";
  }

  return "";
}

function normalizeVatText(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  if (/^kct$/i.test(text)) {
    return "Không chịu thuế";
  }

  return text;
}

function formatPercentDisplay(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return "";
    }

    if (/^không chịu thuế$/i.test(trimmed)) {
      return trimmed;
    }

    if (/%$/.test(trimmed)) {
      return trimmed;
    }

    if (/^kct$/i.test(trimmed)) {
      return "Không chịu thuế";
    }

    const parsed = Number(trimmed.replace(/,/g, ""));
    if (Number.isFinite(parsed)) {
      return `${trimmed}%`;
    }

    return trimmed;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return `${value}%`;
  }

  return String(value);
}

function formatPercentDisplayOrBlankIfZero(value) {
  const parsed = parseOptionalNumber(value);
  if (parsed === 0) {
    return "";
  }

  if (parsed !== null) {
    return `${parsed}%`;
  }

  return formatPercentDisplay(value);
}

function formatBlankIfZero(value) {
  const parsed = parseOptionalNumber(value);
  if (parsed === null || parsed === 0) {
    return "";
  }

  return parsed;
}

function buildNormalizedHeaderSet(values) {
  return new Set(values.map(normalizeHeaderKey));
}

function normalizeProductName(value) {
  return normalizeHeaderKey(value);
}

function normalizeHeaderKey(value) {
  return normalizeText(value)
    .replace(/[^a-z0-9/]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

async function onProductTemplateSelected(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    setStatus(el.productTemplateStatus, `Đang đọc mẫu sản phẩm ${file.name}...`);
    const buffer = await file.arrayBuffer();
    const parsed = parseProductCatalogWorkbook(buffer, file.name);

    state.productCatalog = {
      fileName: parsed.fileName,
      loadedAt: new Date().toISOString(),
      products: parsed.products,
      selections: mergeProductSelections(state.productCatalog.selections, parsed.products)
    };

    // Persist parsed catalog and also save original file bytes to IndexedDB for reuse
    await persistProductCatalogState();
    try {
      await saveProductFileToIdb(parsed.fileName, buffer);
    } catch (err) {
      // Non-fatal: if IndexedDB save fails, we still keep parsed data in chrome.storage
      console.warn("Không lưu được tệp mẫu sản phẩm gốc vào IndexedDB:", err);
    }
    renderProductCatalogStatus();
  } catch (error) {
    setStatus(el.productTemplateStatus, `Không thể đọc mẫu sản phẩm: ${error.message}`, true);
  } finally {
    event.target.value = "";
  }
}

async function clearProductCatalog() {
  state.productCatalog = {
    fileName: "",
    loadedAt: null,
    products: [],
    selections: {}
  };

  await persistProductCatalogState();
  try {
    await removeProductFileFromIdb();
  } catch (err) {
    console.warn("Không xóa được tệp mẫu sản phẩm khỏi IndexedDB:", err);
  }
  renderProductCatalogStatus();
  setStatus(el.productTemplateStatus, "Đã xóa mẫu sản phẩm đã lưu.");
}

function parseProductCatalogWorkbook(buffer, fileName) {
  // Accept ArrayBuffer or Uint8Array. Prefer Uint8Array for 'array' type.
  let data = buffer;
  if (buffer instanceof ArrayBuffer) {
    data = new Uint8Array(buffer);
  }

  // Try reading as array first (works for .xlsx/.xlsm). If that fails (corrupted zip / bad uncompressed size),
  // and the filename suggests an old .xls, fallback to binary read.
  let workbook;
  try {
    workbook = window.XLSX.read(data, { type: "array", cellStyles: true });
  } catch (err) {
    const isOldXls = /\.xls$/i.test(String(fileName || "")) && !/\.xlsx$/i.test(String(fileName || ""));
    if (isOldXls) {
      try {
        // Convert Uint8Array to binary string in chunks to avoid call stack limits
        const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
        let binary = "";
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
          binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunkSize)));
        }
        workbook = window.XLSX.read(binary, { type: "binary", cellStyles: true });
      } catch (err2) {
        throw new Error(`Không thể đọc file mẫu (xls): ${err2?.message || String(err2)}`);
      }
    } else {
      throw new Error(`Không thể đọc file mẫu: ${err?.message || String(err)}`);
    }
  }

  const products = [];
  const seenKeys = new Set();

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      return;
    }

    const sheetProducts = extractProductsFromSheet(sheet, sheetName);
    sheetProducts.forEach((product) => {
      const dedupeKey = product.code
        ? `${product.normalizedName}::${normalizeProductName(product.code)}`
        : `${product.normalizedName}::row-${product.sheetName}-${product.rowNumber}`;
      if (seenKeys.has(dedupeKey)) {
        return;
      }

      seenKeys.add(dedupeKey);
      products.push(product);
    });
  });

  if (products.length === 0) {
    throw new Error("Không tìm thấy cột mã/tên sản phẩm trong file đã chọn");
  }

  return {
    fileName,
    products
  };
}

function extractProductsFromSheet(sheet, sheetName) {
  const matrix = window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: ""
  });

  const headerRowIndex = findProductHeaderRowIndex(matrix);
  if (headerRowIndex < 0) {
    return [];
  }

  const headerRow = matrix[headerRowIndex] || [];
  const codeColumn = findHeaderColumnIndex(headerRow, PRODUCT_CODE_HEADERS);
  const nameColumn = findHeaderColumnIndex(headerRow, PRODUCT_NAME_HEADERS);
  const unitColumn = findHeaderColumnIndex(headerRow, PRODUCT_UNIT_HEADERS);

  if (nameColumn < 0) {
    return [];
  }

  const resolvedCodeColumn = codeColumn >= 0 ? codeColumn : Math.max(0, nameColumn - 1);

  const products = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    const code = normalizeCellText(row[resolvedCodeColumn]);
    const name = normalizeCellText(row[nameColumn]);

    if (!name) {
      continue;
    }

    products.push({
      code,
      hasCode: Boolean(code),
      name,
      normalizedName: normalizeProductName(name),
      unit: unitColumn >= 0 ? normalizeCellText(row[unitColumn]) : "",
      sheetName,
      rowNumber: rowIndex + 1
    });
  }

  return products;
}

function findProductHeaderRowIndex(matrix) {
  for (let rowIndex = 0; rowIndex < Math.min(matrix.length, PRODUCT_TEMPLATE_SCAN_ROWS); rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    const normalized = row.map(normalizeCellText).map(normalizeHeaderKey);
    const hasCodeHeader = normalized.some((value) => PRODUCT_CODE_HEADERS.has(value));
    const hasNameHeader = normalized.some((value) => PRODUCT_NAME_HEADERS.has(value));
    const hasSttHeader = normalized.some((value) => value === "stt");

    if ((hasCodeHeader && hasNameHeader) || (hasSttHeader && hasNameHeader && hasCodeHeader)) {
      return rowIndex;
    }

    if (hasNameHeader && normalized.some((value) => /(^|\s)ma(\s|$)/.test(value))) {
      return rowIndex;
    }
  }

  return -1;
}

function findHeaderColumnIndex(row, headerSet) {
  for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
    const value = normalizeHeaderKey(row[columnIndex]);
    if (headerSet.has(value)) {
      return columnIndex;
    }
  }

  for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
    const value = normalizeHeaderKey(row[columnIndex]);
    if (headerSet === PRODUCT_CODE_HEADERS && /(^|\s)ma(\s|$)/.test(value)) {
      return columnIndex;
    }
    if (headerSet === PRODUCT_NAME_HEADERS && /(^|\s)ten(\s|$)/.test(value)) {
      return columnIndex;
    }
  }

  return -1;
}

function normalizeCellText(value) {
  return String(value ?? "").trim();
}

function mergeProductSelections(existingSelections, products) {
  const nextSelections = {};
  const availableCodesByName = new Map();

  products.forEach((product) => {
    const list = availableCodesByName.get(product.normalizedName) || [];
    list.push(product.code);
    availableCodesByName.set(product.normalizedName, list);
  });

  Object.entries(existingSelections || {}).forEach(([normalizedName, code]) => {
    const availableCodes = availableCodesByName.get(normalizedName) || [];
    if (availableCodes.includes(code)) {
      nextSelections[normalizedName] = code;
    }
  });

  return nextSelections;
}

function renderProductCatalogStatus(productAudit = { ambiguous: [], missing: [], noCode: [] }) {
  if (!el.productTemplateStatus || !el.productResolutionList) {
    return;
  }

  const catalog = state.productCatalog;
  if (!catalog.products.length) {
    setStatus(
      el.productTemplateStatus,
      "Chưa tải mẫu sản phẩm. Hãy chọn file BC_San_Pham để tra mã trước khi xuất."
    );
    el.productResolutionList.innerHTML = "";
    return;
  }

  const grouped = groupProductsByName(catalog.products);
  const duplicateGroups = Array.from(grouped.entries()).filter(([, products]) => products.length > 1);
  const ambiguousSet = new Set(productAudit.ambiguous.map((item) => item.normalizedName));
  const noCodeProducts = catalog.products.filter((product) => !product.code);
  const noCodeRows = productAudit.noCode?.length
    ? productAudit.noCode
    : noCodeProducts.map((product) => ({
      name: product.name,
      normalizedName: product.normalizedName,
      source: `${product.sheetName}:${product.rowNumber}`,
      previewRow: {
        itemCode: product.code || "(sinh tự động)",
        itemType: "",
        itemName: product.name,
        unit: product.unit || "",
        quantity: "",
        price: "",
        discountPercent: "",
        discountVnd: "",
        vat: ""
      }
    }));

  const duplicateRows = duplicateGroups
    .map(([normalizedName, products]) => {
      const selectedCode = catalog.selections?.[normalizedName] || products[0].code;
      const hasStoredSelection = Boolean(catalog.selections?.[normalizedName]);
      const options = products
        .map((product) => {
          const selected = product.code === selectedCode ? "selected" : "";
          return `<option value="${escapeHtml(product.code)}" ${selected}>${escapeHtml(product.code)} - ${escapeHtml(product.name)}${product.unit ? ` (${escapeHtml(product.unit)})` : ""} [${escapeHtml(product.sheetName)}:${product.rowNumber}]</option>`;
        })
        .join("");

      const isAmbiguous = ambiguousSet.has(normalizedName);
      return `
        <tr>
          <td>
            <div class="product-name">${escapeHtml(products[0].name)}</div>
            <div class="product-meta">${escapeHtml(normalizedName)}</div>
          </td>
          <td>
            <select class="product-select" data-product-name="${escapeHtml(normalizedName)}">
              ${options}
            </select>
          </td>
          <td>${products.length}</td>
          <td>${isAmbiguous || !hasStoredSelection ? "Cần chọn" : "Đã chọn"}</td>
        </tr>
      `;
    })
    .join("");

  const duplicateSection = duplicateGroups.length
    ? `
        <table class="product-table">
          <thead>
            <tr>
              <th>Tên sản phẩm</th>
              <th>Mã được chọn</th>
              <th>Số mã trùng</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>${duplicateRows}</tbody>
        </table>
      `
    : '<p class="muted-note">Không có tên sản phẩm trùng trong mẫu đã tải.</p>';

  el.productResolutionList.innerHTML = `
    <div class="product-summary">
      <div><strong>${catalog.fileName || "Mẫu sản phẩm"}</strong></div>
      <div>${catalog.products.length} dòng sản phẩm đã tải</div>
      <div>${duplicateGroups.length} tên trùng</div>
    </div>
    ${duplicateSection}
  `;

  el.productResolutionList.querySelectorAll(".product-select").forEach((select) => {
    select.addEventListener("change", onProductSelectionChanged);
  });

  const duplicateCount = duplicateGroups.length;
  const statusParts = [`Đã tải ${catalog.products.length} dòng sản phẩm từ ${catalog.fileName || "mẫu local"}.`];
  if (duplicateCount > 0) {
    statusParts.push(`${duplicateCount} tên đang có nhiều mã, cần chọn trước khi xuất.`);
  }

  setStatus(el.productTemplateStatus, statusParts.join(" "));
}

function openExportReviewDialog(entries) {
  if (!el.exportReviewDialog) {
    return Promise.resolve(null);
  }

  const normalizedEntries = entries.map((entry) => ({
    ...entry,
    include: true,
    code: normalizeCellText(entry.currentCode || entry.row?.itemCode || entry.resolvedCode || "")
  }));

  state.exportReview = {
    isOpen: true,
    pendingAction: "idle",
    entries: normalizedEntries,
    resolve: null
  };

  renderExportReviewDialog();

  return new Promise((resolve) => {
    state.exportReview.resolve = resolve;
    el.exportReviewDialog.showModal();
  });
}

function closeExportReviewDialog(accepted) {
  if (!el.exportReviewDialog || !state.exportReview.isOpen) {
    return;
  }

  state.exportReview.pendingAction = accepted ? "accept" : "cancel";
  if (el.exportReviewDialog.open) {
    el.exportReviewDialog.close();
  }
}

function onExportReviewDialogClose() {
  if (!state.exportReview.resolve) {
    resetExportReviewDialogState();
    return;
  }

  const resolver = state.exportReview.resolve;
  const accepted = state.exportReview.pendingAction === "accept";
  const decision = accepted ? readExportReviewDecision() : null;

  resetExportReviewDialogState();
  resolver(decision);
}

function onExportReviewDialogCancel(event) {
  event.preventDefault();
  closeExportReviewDialog(false);
}

function onExportReviewDialogClick(event) {
  if (event.target === el.exportReviewDialog) {
    closeExportReviewDialog(false);
  }
}

function onExportReviewDialogFieldChange(event) {
  const target = event.target;
  const row = target.closest?.("[data-review-key]");
  if (!row) {
    return;
  }

  const reviewKey = row.dataset.reviewKey;
  const entry = state.exportReview.entries.find((item) => item.key === reviewKey);
  if (!entry) {
    return;
  }

  if (target.matches("input[type='checkbox']")) {
    entry.include = target.checked;
    syncExportReviewSelectAllControl();
  }

  if (target.matches("input[type='text']")) {
    entry.code = target.value;
  }

  updateExportReviewSummary();
}

function onExportReviewSelectAllChange(event) {
  const checked = event.target.checked;
  const entries = state.exportReview.entries || [];

  entries.forEach((entry) => {
    entry.include = checked;
  });

  if (el.exportReviewList) {
    el.exportReviewList.querySelectorAll(".export-review-include").forEach((checkbox) => {
      checkbox.checked = checked;
    });
  }

  syncExportReviewSelectAllControl();
  updateExportReviewSummary();
}

function renderExportReviewDialog() {
  if (!el.exportReviewDialog || !el.exportReviewList || !el.exportReviewSummary || !el.exportReviewStatus) {
    return;
  }

  const entries = state.exportReview.entries || [];
  if (!entries.length) {
    el.exportReviewList.innerHTML = "";
    el.exportReviewSummary.textContent = "Không có dòng nào cần review.";
    el.exportReviewStatus.textContent = "";
    return;
  }

  el.exportReviewList.innerHTML = entries
    .map((entry, index) => {
      const rowNumber = entry.lineIndex >= 0 ? entry.lineIndex + 1 : 1;
      const reason = entry.resolvedCode
        ? "Mã SP trong mẫu còn trống"
        : "Không tìm thấy mã SP trong mẫu";
      const row = entry.row || {};

      return `
        <tr data-review-key="${escapeHtml(entry.key)}">
          <td>
            <input type="checkbox" class="export-review-include" ${entry.include ? "checked" : ""} />
          </td>
          <td>
            <div class="export-review-invoice">${escapeHtml(entry.invoiceLabel)}</div>
          </td>
          <td>
            <div class="export-review-code-row">
              <input type="text" class="export-review-code" value="${escapeHtml(entry.code)}" placeholder="Mã sản phẩm" />
              <div class="export-review-meta export-review-current-code">Mã hiện tại: ${escapeHtml(entry.currentCode || "")}</div>
              <div class="export-review-meta export-review-reason">Lý do: ${escapeHtml(reason)}</div>
            </div>
          </td>
          <td>
            <div class="export-review-name">${escapeHtml(entry.itemName)}</div>
          </td>
          <td>${escapeHtml(row.unit || "")}</td>
          <td>${escapeHtml(String(row.quantity ?? ""))}</td>
          <td>${escapeHtml(String(row.price ?? ""))}</td>
          <td>${escapeHtml([row.discountPercent, row.discountVnd].filter(Boolean).join(" / "))}</td>
          <td>${escapeHtml(String(row.vat ?? ""))}</td>
        </tr>
      `;
    })
    .join("");

  syncExportReviewSelectAllControl();
  updateExportReviewSummary();
}

function syncExportReviewSelectAllControl() {
  if (!el.exportReviewSelectAll) {
    return;
  }

  const entries = state.exportReview.entries || [];
  const includedCount = entries.filter((entry) => entry.include).length;

  el.exportReviewSelectAll.checked = entries.length > 0 && includedCount === entries.length;
  el.exportReviewSelectAll.indeterminate = includedCount > 0 && includedCount < entries.length;
}

function updateExportReviewSummary() {
  if (!el.exportReviewSummary || !el.exportReviewStatus) {
    return;
  }

  const entries = state.exportReview.entries || [];
  const includedCount = entries.filter((entry) => entry.include).length;
  const excludedCount = entries.length - includedCount;
  const invoiceCount = new Set(entries.map((entry) => entry.invoiceId)).size;

  el.exportReviewSummary.textContent = `${entries.length} dòng cần review từ ${invoiceCount} hóa đơn. ${includedCount} dòng đang được chọn xuất, ${excludedCount} dòng sẽ bị loại bỏ.`;
  el.exportReviewStatus.textContent = entries.some((entry) => !normalizeCellText(entry.code))
    ? "Hãy nhập mã sản phẩm cho các dòng được chọn xuất nếu muốn thay mã hiện tại."
    : "Có thể tiếp tục xuất ngay với mã hiện tại hoặc mã bạn vừa chỉnh sửa.";
}

function readExportReviewDecision() {
  const entries = (state.exportReview.entries || []).map((entry) => ({
    ...entry,
    code: normalizeCellText(entry.code || entry.currentCode || "")
  }));

  return {
    entries,
    entriesByKey: buildExportReviewDecisionMap({ entries })
  };
}

function resetExportReviewDialogState() {
  state.exportReview = {
    isOpen: false,
    pendingAction: "idle",
    entries: [],
    resolve: null
  };

  if (el.exportReviewDialog?.open) {
    el.exportReviewDialog.close();
  }

  if (el.exportReviewList) {
    el.exportReviewList.innerHTML = "";
  }
  if (el.exportReviewSummary) {
    el.exportReviewSummary.textContent = "";
  }
  if (el.exportReviewStatus) {
    el.exportReviewStatus.textContent = "";
  }
}

function groupProductsByName(products) {
  const grouped = new Map();

  products.forEach((product) => {
    const list = grouped.get(product.normalizedName) || [];
    list.push(product);
    grouped.set(product.normalizedName, list);
  });

  return grouped;
}

async function onProductSelectionChanged(event) {
  const normalizedName = event.target.dataset.productName;
  const code = event.target.value;

  if (!normalizedName) {
    return;
  }

  state.productCatalog = {
    ...state.productCatalog,
    selections: {
      ...(state.productCatalog.selections || {}),
      [normalizedName]: code
    }
  };

  await persistProductCatalogState();
  renderProductCatalogStatus();
}

async function persistProductCatalogState() {
  const { [APP_STATE_STORAGE_KEY]: appState = {} } = await chrome.storage.local.get(APP_STATE_STORAGE_KEY);

  await chrome.storage.local.set({
    [APP_STATE_STORAGE_KEY]: {
      ...appState,
      productCatalog: state.productCatalog
    }
  });
}

// IndexedDB helpers to persist the original uploaded product file (blob)
const IDB_DB_NAME = "hoadondientu_files";
const IDB_STORE_NAME = "files";
function openIdb() {
  return new Promise((resolve, reject) => {
    const req = window.indexedDB.open(IDB_DB_NAME, 1);
    req.onupgradeneeded = (ev) => {
      const db = ev.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE_NAME)) {
        db.createObjectStore(IDB_STORE_NAME, { keyPath: "id" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("IndexedDB open error"));
  });
}

async function saveProductFileToIdb(fileName, arrayBuffer) {
  const db = await openIdb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE_NAME, "readwrite");
    const store = tx.objectStore(IDB_STORE_NAME);
    const blob = new Blob([arrayBuffer], { type: "application/octet-stream" });
    const record = {
      id: "productCatalog",
      fileName,
      blob,
      savedAt: new Date().toISOString()
    };
    const req = store.put(record);
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error || new Error("IndexedDB put error"));
  });
}

async function removeProductFileFromIdb() {
  const db = await openIdb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE_NAME, "readwrite");
    const store = tx.objectStore(IDB_STORE_NAME);
    const req = store.delete("productCatalog");
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error || new Error("IndexedDB delete error"));
  });
}

async function getProductFileFromIdb() {
  const db = await openIdb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE_NAME, "readonly");
    const store = tx.objectStore(IDB_STORE_NAME);
    const req = store.get("productCatalog");
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error || new Error("IndexedDB get error"));
  });
}

function toApiDate(inputValue, endOfDay) {
  const [yyyy, mm, dd] = inputValue.split("-");
  const suffix = endOfDay ? "23:59:59" : "00:00:00";
  return `${dd}/${mm}/${yyyy}T${suffix}`;
}

function toMoney(value) {
  const num = toNumber(value);
  if (Number.isNaN(num)) return "-";
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0
  }).format(num);
}

function toNumber(value) {
  if (typeof value === "number") {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Number(value.replaceAll(",", ""));
    return Number.isNaN(parsed) ? 0 : parsed;
  }

  return 0;
}

function formatDateTime(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "short",
    timeStyle: "short"
  }).format(date);
}

function getInvoiceId(item) {
  return item?.id || item?.mhdon || `${item?.khhdon || ""}-${item?.shdon || ""}-${item?.tdlap || ""}`;
}

function setStatus(target, message, isError = false) {
  target.textContent = message;
  target.classList.toggle("error", Boolean(isError));
}

// --- Support dialog helpers ---
function openSupportDialog() {
  if (!el.supportDialog) return;
  if (el.supportStatus) el.supportStatus.textContent = "";
  try {
    el.supportDialog.showModal();
  } catch (e) {
    // fallback for browsers that don't support dialog.showModal
    el.supportDialog.style.display = "block";
  }
}

function closeSupportDialog() {
  if (!el.supportDialog) return;
  try {
    if (el.supportDialog.open) el.supportDialog.close();
    else el.supportDialog.style.display = "none";
  } catch (e) {
    el.supportDialog.style.display = "none";
  }
}

function renderPortalFlowStatus(portalFlow, phase) {
  if (!el.portalFlowStatus) return;

  const effectivePhase = phase || portalFlow?.phase || "unknown";
  const phaseLabel = describePortalPhase(effectivePhase);
  const detailParts = [];

  if (portalFlow?.title) {
    detailParts.push(portalFlow.title);
  }

  if (portalFlow?.hasLoginModal) {
    detailParts.push("đang ở modal đăng nhập");
  }

  if (portalFlow?.hasInvoiceSearch) {
    detailParts.push("đã thấy luồng tra cứu hóa đơn");
  }

  if (portalFlow?.hasLoggedInActionButton) {
    detailParts.push("thấy nút hành động (khả năng đã đăng nhập)");
  }

  const suffix = detailParts.length ? ` - ${detailParts.join("; ")}` : "";
  setStatus(el.portalFlowStatus, `${phaseLabel}${suffix}`);
}

function describePortalPhase(phase) {
  switch (phase) {
    case "login-modal":
      return "Màn hình đăng nhập";
    case "invoice-search":
      return "Màn hình tra cứu hóa đơn";
    case "authenticated-shell":
      return "Đã xác thực";
    case "login-required":
      return "Phiên hết hạn hoặc chưa xác thực";
    default:
      return "Chưa xác định";
  }
}

function escapeHtml(input) {
  return String(input)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderDetailView(detail) {
  if (!detail || typeof detail !== "object") {
    return `<div class="detail-empty">${escapeHtml(String(detail ?? "Không có dữ liệu chi tiết."))}</div>`;
  }


  const seller = buildPreviewPartyInfo(detail, "seller");
  const buyer = buildPreviewPartyInfo(detail, "buyer");
  const lineItems = extractInvoiceLineItems(detail);
  const previewRows = (lineItems.length > 0 ? lineItems : [buildPreviewFallbackLineItem(detail)]).map(
    (lineItem, index) => buildPreviewProductRow(detail, lineItem, index, state.productCatalog)
  );

  return `
    <div class="preview-layout">
      <section class="preview-section">
        <div class="preview-section-title">Người bán</div>
        <div class="preview-party-grid">
          ${renderPreviewPartyField("Tên người bán", seller.name)}
          ${renderPreviewPartyField("Mã số thuế", seller.taxCode)}
          ${renderPreviewPartyField("Địa chỉ", seller.address)}
        </div>
      </section>

      <section class="preview-section">
        <div class="preview-section-title">Người mua</div>
        <div class="preview-party-grid">
          ${renderPreviewPartyField("Tên người mua", buyer.name)}
          ${renderPreviewPartyField("Mã số thuế", buyer.taxCode)}
          ${renderPreviewPartyField("Địa chỉ", buyer.address)}
        </div>
      </section>

      <section class="preview-section">
        <div class="preview-section-title">Thông tin sản phẩm</div>
        <div class="preview-table-wrap">
          <table class="preview-table">
            <colgroup>
              <col class="col-stt" />
              <col class="col-product-code" />
              <col class="col-nature" />
              <col class="col-product-type" />
              <col class="col-product-name" />
              <col class="col-unit" />
              <col class="col-quantity" />
              <col class="col-price" />
              <col class="col-discount" />
              <col class="col-tax" />
              <col class="col-amount" />
            </colgroup>
            <thead>
              <tr>
                <th>STT</th>
                <th>Mã sản phẩm</th>
                <th>Tính chất</th>
                <th>Loại hàng hoá đặc trưng</th>
                <th>Tên hàng hóa, dịch vụ</th>
                <th>Đơn vị tính</th>
                <th>Số lượng</th>
                <th>Đơn giá</th>
                <th>Chiết khấu</th>
                <th>Thuế suất</th>
                <th>Thành tiền chưa có thuế GTGT</th>
              </tr>
            </thead>
            <tbody>
              ${previewRows.join("")}
            </tbody>
          </table>
        </div>
      </section>
      <section class="preview-section">
        <div class="preview-section-title">Tổng hợp thanh toán</div>
        <div class="preview-party-grid">
          ${(() => {
      const totals = buildPreviewTotals(detail);
      return `
            ${renderPreviewPartyField("Thuế suất", totals.taxRate)}
            ${renderPreviewPartyField("Tổng tiền chưa thuế", totals.totalBeforeTax)}
            ${renderPreviewPartyField("Tổng tiền thuế", totals.totalTax)}
            ${renderPreviewPartyField("Tổng tiền phí", totals.totalFees)}
            ${renderPreviewPartyField("Tổng tiền chiết khấu thương mại", totals.totalDiscount)}
            ${renderPreviewPartyField("Tổng tiền thanh toán bằng số", totals.totalNumeric)}
            ${renderPreviewPartyField("Tổng tiền thanh toán bằng chữ", totals.totalText)}
          `;
    })()}
        </div>
      </section>
    </div>
  `;
}

function buildPreviewPartyInfo(detail, role) {
  const config = role === "seller"
    ? {
      nameKeys: ["nbten", "tennguoiban", "sellername", "seller_name", "name"],
      taxKeys: ["nbmst", "sellerTaxCode", "seller_tax_code"],
      addressKeys: ["nbdchi", "selleraddress", "seller_address", "address"]
    }
    : {
      nameKeys: ["nmten", "nmtnmua", "tennguoinmua", "buyername", "buyer_name", "name"],
      taxKeys: ["nmmst", "buyerTaxCode", "buyer_tax_code"],
      addressKeys: ["nmdchi", "buyeraddress", "buyer_address", "address"]
    };

  return {
    name: pickFirstText(detail, config.nameKeys) || "-",
    taxCode: pickFirstText(detail, config.taxKeys) || "-",
    address: pickFirstText(detail, config.addressKeys) || "-"
  };
}

function buildPreviewTotals(detail) {
  const totalBefore = pickFirstNumber(detail, ["tgtcthue", "tgtttbso", "tgttbso", "tgtctchu"], null);
  const totalTax = pickFirstNumber(detail, ["tgtthue", "tthue"], null);
  const totalFees = pickFirstNumber(detail, ["tgtphi", "tgtphi"], null);
  const totalDiscount = pickFirstNumber(detail, ["ttcktmai", "ttcktmai"], null);
  const totalNumeric = pickFirstNumber(detail, ["tgtttbso", "tgttbso", "tgtcthue"], null);
  const totalText = pickFirstText(detail, ["tgtttbchu", "tgtttbchu"]) || "-";

  return {
    taxRate: formatPreviewTaxRate(detail, {}),
    totalBeforeTax: totalBefore !== null ? formatPreviewMoney(totalBefore) : "-",
    totalTax: totalTax !== null ? formatPreviewMoney(totalTax) : "0",
    totalFees: totalFees !== null ? formatPreviewMoney(totalFees) : "0",
    totalDiscount: totalDiscount !== null ? formatPreviewMoney(totalDiscount) : "0",
    totalNumeric: totalNumeric !== null ? formatPreviewMoney(totalNumeric) : "0",
    totalText: totalText
  };
}

function buildPreviewFallbackLineItem(detail) {
  return {
    stt: 1,
    msp: "",
    tchat: detail?.tchat,
    loaihanghoa: "",
    ten: detail?.thdon || detail?.tlhdon || "Hóa đơn mua vào",
    dvtinh: "-",
    sluong: 1,
    dgia: toNumber(detail?.tgtcthue || detail?.tgtttbso || 0),
    stckhau: detail?.ttcktmai || 0,
    tlckhau: detail?.tlckhau || null,
    ltsuat: detail?.thttltsuat?.[0]?.tsuat ?? (toNumber(detail?.tgtthue) === 0 ? "KCT" : ""),
    thtcthue: detail?.tgtcthue || detail?.tgtttbso || 0
  };
}

function buildPreviewProductRow(detail, lineItem, index, productCatalog) {
  const stt = pickFirstNumber(lineItem, ["stt", "idx", "index"], index + 1);
  const itemNature = formatPreviewItemNature(pickFirstText(lineItem, ["tchat", "tinhchat", "nature"]) || lineItem?.tchat);
  const productType = pickFirstText(lineItem, ["loaihanghoa", "loaihang", "lhdvu", "itemtype", "type"]) || "";
  // Do not display Loại hàng in preview per request — keep blank
  const previewProductType = "";
  const itemName =
    pickFirstText(lineItem, ["ten", "thhdvu", "tenhang", "tenhh", "tenhhdvu", "name"]) || "-";
  const resolvedProduct = resolveProductCatalogEntry(itemName, productCatalog);
  const productCode = resolvedProduct?.code || "-";
  const unit = pickFirstText(lineItem, ["dvtinh", "dvt", "dvtien", "donvitinh", "unit"]) || "-";
  const quantity = formatPreviewNumber(pickFirstNumber(lineItem, ["sluong", "soluong", "qty", "quantity", "kluong"], null));
  const price = formatPreviewMoney(pickFirstNumber(lineItem, ["dgia", "dongia", "dgiacthue", "price"], null));
  const discount = formatPreviewDiscount(lineItem);
  const taxRate = formatPreviewTaxRate(detail, lineItem);
  const amount = formatPreviewMoney(
    pickFirstNumber(lineItem, ["thtcthue", "thtien", "thanhTien", "thanhtien", "amount"], null) ??
    computePreviewLineAmount(lineItem)
  );

  return `
    <tr>
      <td>${escapeHtml(String(stt || index + 1))}</td>
      <td>${escapeHtml(productCode)}</td>
      <td>${escapeHtml(itemNature)}</td>
      <td>${escapeHtml(previewProductType)}</td>
      <td>${escapeHtml(itemName)}</td>
      <td>${escapeHtml(unit)}</td>
      <td class="preview-number">${escapeHtml(quantity)}</td>
      <td class="preview-number">${escapeHtml(price)}</td>
      <td class="preview-number">${escapeHtml(discount)}</td>
      <td>${escapeHtml(taxRate)}</td>
      <td class="preview-number">${escapeHtml(amount)}</td>
    </tr>
  `;
}

function computePreviewLineAmount(lineItem) {
  const quantity = pickFirstNumber(lineItem, ["sluong", "soluong", "qty", "quantity", "kluong"], null);
  const price = pickFirstNumber(lineItem, ["dgia", "dongia", "dgiacthue", "price"], null);
  if (quantity === null || price === null) {
    return null;
  }

  return quantity * price;
}

function formatPreviewItemNature(value) {
  const text = String(value ?? "").trim();
  if (!text) {
    return "Hàng hóa, dịch vụ";
  }

  const normalized = normalizeText(text);
  if (normalized === "1" || normalized === "hang hoa dich vu" || normalized.includes("hang hoa")) {
    return "Hàng hóa, dịch vụ";
  }

  if (normalized === "2") {
    return "Dịch vụ";
  }

  return text;
}

function formatPreviewDiscount(lineItem) {
  const discountAmount = pickFirstNumber(lineItem, ["stckhau", "tienckhau", "giamgia", "discountamount"], null);
  if (discountAmount !== null && discountAmount !== 0) {
    return formatPreviewMoney(discountAmount);
  }

  const discountPercent = pickFirstNumber(lineItem, ["tlckhau", "ptckhau", "ptgiam", "discountpercent"], null);
  if (discountPercent !== null && discountPercent !== 0) {
    return formatPreviewPercent(discountPercent);
  }

  return "0";
}

function formatPreviewTaxRate(detail, lineItem) {
  const lineTax = pickFirstText(lineItem, ["ltsuat", "tsuat", "thuesuat", "vat", "taxrate"]);
  if (lineTax) {
    const trimmed = lineTax.trim();
    if (trimmed === "0" || trimmed === "0.0") {
      return "KCT";
    }

    return trimmed;
  }

  const invoiceTax = detail?.thttltsuat?.[0]?.tsuat;
  if (invoiceTax !== null && invoiceTax !== undefined && invoiceTax !== "") {
    const normalizedTax = String(invoiceTax).trim();
    if (normalizedTax === "0" || normalizedTax === "0.0") {
      return "KCT";
    }

    return normalizedTax;
  }

  if (toNumber(detail?.tgtthue) === 0) {
    return "KCT";
  }

  return "";
}

function formatPreviewMoney(value) {
  const parsed = parseOptionalNumber(value);
  if (parsed === null) {
    return "-";
  }

  return new Intl.NumberFormat("vi-VN", {
    maximumFractionDigits: 0
  }).format(parsed);
}

function formatPreviewNumber(value) {
  const parsed = parseOptionalNumber(value);
  if (parsed === null) {
    return "-";
  }

  return new Intl.NumberFormat("vi-VN", {
    maximumFractionDigits: 3
  }).format(parsed);
}

function formatPreviewPercent(value) {
  const parsed = parseOptionalNumber(value);
  if (parsed === null) {
    return "-";
  }

  return `${new Intl.NumberFormat("vi-VN", {
    maximumFractionDigits: 3
  }).format(parsed)}%`;
}

function renderPreviewPartyField(label, value) {
  return `
    <div class="preview-party-field">
      <strong>${escapeHtml(label)}</strong>
      <span>${escapeHtml(value || "-")}</span>
    </div>
  `;
}

function shouldRenderAsNestedSection(value) {
  return Array.isArray(value) || (value && typeof value === "object");
}

function renderNestedSection(title, value) {
  if (Array.isArray(value)) {
    if (value.length === 0) {
      return `
        <div class="detail-section">
          <div class="detail-section-title">${escapeHtml(prettifyDetailKey(title))}</div>
          <div class="detail-empty">Không có dữ liệu.</div>
        </div>
      `;
    }

    if (value.every((item) => item && typeof item === "object" && !Array.isArray(item))) {
      const columns = collectTableColumns(value);
      const headerCells = columns.map((column) => `<th>${escapeHtml(prettifyDetailKey(column))}</th>`).join("");
      const bodyRows = value
        .map((item, index) => {
          const cells = columns
            .map((column) => `<td class="detail-value">${renderDetailCell(item[column], `${title}[${index}].${column}`)}</td>`)
            .join("");
          return `<tr>${cells}</tr>`;
        })
        .join("");

      return `
        <div class="detail-section">
          <div class="detail-section-title">${escapeHtml(prettifyDetailKey(title))}</div>
          <div class="detail-subtable">
            <table class="detail-table">
              <thead><tr>${headerCells}</tr></thead>
              <tbody>${bodyRows}</tbody>
            </table>
          </div>
        </div>
      `;
    }

    const listItems = value.map((item) => `<li>${renderDetailCell(item, title)}</li>`).join("");
    return `
      <div class="detail-section">
        <div class="detail-section-title">${escapeHtml(prettifyDetailKey(title))}</div>
        <div class="detail-empty"><ul>${listItems}</ul></div>
      </div>
    `;
  }

  const entries = Object.entries(value);
  if (entries.length === 0) {
    return `
      <div class="detail-section">
        <div class="detail-section-title">${escapeHtml(prettifyDetailKey(title))}</div>
        <div class="detail-empty">Không có dữ liệu.</div>
      </div>
    `;
  }

  const rows = entries
    .map(([key, childValue]) => `
      <tr>
        <td class="detail-key">${escapeHtml(prettifyDetailKey(key))}</td>
        <td class="detail-value">${renderDetailCell(childValue, `${title}.${key}`)}</td>
      </tr>
    `)
    .join("");

  return `
    <div class="detail-section">
      <div class="detail-section-title">${escapeHtml(prettifyDetailKey(title))}</div>
      <table class="detail-table">
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderDetailCell(value, path) {
  if (value === null || value === undefined || value === "") {
    return `<span class="detail-empty">-</span>`;
  }

  if (Array.isArray(value) || (value && typeof value === "object")) {
    return renderCompactDetailValue(value, path);
  }

  return escapeHtml(formatDetailPrimitive(value));
}

function renderCompactDetailValue(value, path) {
  if (Array.isArray(value)) {
    if (value.length === 0) {
      return `<span class="detail-empty">[]</span>`;
    }

    if (value.every((item) => item && typeof item === "object" && !Array.isArray(item))) {
      const rows = value
        .map((item, index) => {
          const compact = Object.entries(item)
            .map(([key, childValue]) => `<div><span class="detail-tag">${escapeHtml(prettifyDetailKey(key))}</span> ${renderDetailCell(childValue, `${path}[${index}].${key}`)}</div>`)
            .join("");
          return `<div style="margin-bottom:8px;">${compact}</div>`;
        })
        .join("");
      return rows;
    }

    return `<div>${value.map((item) => renderDetailCell(item, path)).join("<br/>")}</div>`;
  }

  const entries = Object.entries(value);
  if (entries.length === 0) {
    return `<span class="detail-empty">{}</span>`;
  }

  return `<div>${entries
    .map(([key, childValue]) => `<div><span class="detail-tag">${escapeHtml(prettifyDetailKey(key))}</span> ${renderDetailCell(childValue, `${path}.${key}`)}</div>`)
    .join("")}</div>`;
}

function collectTableColumns(items) {
  const columns = [];
  const seen = new Set();

  items.forEach((item) => {
    Object.keys(item).forEach((key) => {
      if (seen.has(key)) {
        return;
      }

      seen.add(key);
      columns.push(key);
    });
  });

  return columns;
}

function formatDetailPrimitive(value) {
  if (value === null || value === undefined) {
    return "-";
  }

  if (typeof value === "string") {
    return value;
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  return String(value);
}

function prettifyDetailKey(key) {
  return String(key)
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^./, (char) => char.toUpperCase());
}

function sendRuntimeMessage(message) {
  return new Promise((resolve, reject) => {
    try {
      if (chrome && chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage(message, (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }
          resolve(response);
        });
      } else {
        reject(new Error("chrome.runtime.sendMessage not available"));
      }
    } catch (e) {
      reject(e instanceof Error ? e : new Error(String(e)));
    }
  });
}

async function loadCachedState() {
  setDefaultDateRange();

  const [syncResponse, appStateResponse] = await Promise.all([
    sendRuntimeMessage({ type: "GET_SYNC_STATE" }).catch(() => null),
    chrome.storage.local.get(APP_STATE_STORAGE_KEY)
  ]);

  const appState = appStateResponse?.[APP_STATE_STORAGE_KEY] || {};
  state.syncState = syncResponse?.ok ? (syncResponse.data || state.syncState) : state.syncState;

  const shouldLoadCachedInvoices = Boolean(state.syncState?.lastDateFrom && state.syncState?.lastDateTo);
  if (shouldLoadCachedInvoices) {
    const cachedResponse = await sendRuntimeMessage({
      type: "GET_CACHED_INVOICES",
      payload: {
        dateFrom: state.syncState.lastDateFrom,
        dateTo: state.syncState.lastDateTo
      }
    }).catch(() => null);

    if (cachedResponse?.ok) {
      state.invoices = cachedResponse.data?.invoices || [];
      state.dataOrigin = state.invoices.length ? "local" : "";
      if (!state.syncState.lastSyncAt && cachedResponse.data?.syncState?.lastSyncAt) {
        state.syncState = {
          ...state.syncState,
          ...cachedResponse.data.syncState
        };
      }
      setStatus(
        el.fetchStatus,
        state.invoices.length
          ? `Đã phục hồi ${state.invoices.length} hóa đơn đã lưu local (${state.syncState.lastSyncAt || "n/a"}).`
          : "Không có hóa đơn local trong khoảng đã lưu."
      );
    } else {
      state.invoices = [];
    }
  } else {
    state.invoices = [];
  }

  state.productCatalog = normalizeLoadedProductCatalog(appState?.productCatalog);
  renderProductCatalogStatus();
  renderRows();
  renderInvoiceSourceBadge();
}

function renderInvoiceSourceBadge() {
  // Update both list and detail badges if present
  const listEl = el.invoiceListBadge;
  const detailEl = el.invoiceSourceBadge;

  const isListLocal = state.dataOrigin === "local" || state.syncState?.dataOrigin === "local";
  const listTextLocal = state.syncState?.lastSyncAt ? ` · ${formatDateTime(state.syncState.lastSyncAt)}` : "";

  if (listEl) {
    listEl.hidden = false;
    listEl.classList.remove("local", "online");
    if (isListLocal) {
      listEl.classList.add("local");
      listEl.textContent = `Dữ liệu cục bộ${listTextLocal}`;
    } else {
      listEl.classList.add("online");
      listEl.textContent = `Dữ liệu trực tuyến`;
    }
  }

  if (detailEl) {
    detailEl.hidden = false;
    detailEl.classList.remove("local", "online");

    // Decouple: Only show local for detail if state.detailOrigin is explicitly local
    const isDetailLocal = state.detailOrigin === "local";
    const detailTextLocal = state.syncState?.lastSyncAt ? ` · ${formatDateTime(state.syncState.lastSyncAt)}` : "";

    if (isDetailLocal) {
      detailEl.classList.add("local");
      detailEl.textContent = `Dữ liệu cục bộ${detailTextLocal}`;
    } else {
      detailEl.classList.add("online");
      detailEl.textContent = `Dữ liệu trực tuyến`;
    }
  }
}

function fromApiDateToInput(value) {
  const match = String(value || "").match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (!match) {
    return "";
  }

  const [, dd, mm, yyyy] = match;
  return `${yyyy}-${mm}-${dd}`;
}

function normalizeLoadedProductCatalog(productCatalog) {
  if (!productCatalog?.products?.length) {
    return {
      fileName: "",
      loadedAt: null,
      products: [],
      selections: {}
    };
  }

  return {
    fileName: productCatalog.fileName || "",
    loadedAt: productCatalog.loadedAt || null,
    products: productCatalog.products,
    selections: productCatalog.selections || {}
  };
}
