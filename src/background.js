const PORTAL_BASE_URL = "https://hoadondientu.gdt.gov.vn";
const QUERY_BASE = `${PORTAL_BASE_URL}:30000`;
const PORTAL_SEARCH_URL = `${PORTAL_BASE_URL}/tra-cuu/tra-cuu-hoa-don`;
const QUERY_PATH = "/query/invoices/purchase";
const SCO_QUERY_PATH = "/sco-query/invoices/purchase";
const DETAIL_PATHS = {
  query: "/query/invoices/detail",
  "sco-query": "/sco-query/invoices/detail"
};
const MAX_PURCHASE_PAGE_SIZE = 50;
const PORTAL_REQUEST_DELAY_MS = 1000;
const PORTAL_PAGE_FETCH_TIMEOUT_MS = 12000;
const INVOICE_DB_NAME = "hoadondientu_invoices";
const INVOICE_DB_VERSION = 1;
const INVOICE_STORE_NAME = "invoices";
const INVOICE_META_STORE_NAME = "sync_meta";
const INVOICE_SYNC_META_KEY = "invoice_sync_meta";

let portalRequestChain = Promise.resolve();
let lastPortalRequestAt = 0;

function createEmptyPortalFlow() {
  return {
    phase: "unknown",
    hasLoginModal: null,
    hasLoggedInActionButton: null,
    hasInvoiceSearch: null,
    isLoggedInUi: null,
    tabId: null,
    url: null,
    title: null,
    evidence: [],
    updatedAt: null
  };
}

function normalizePortalFlow(payload = {}, tabId = null) {
  const hasLoginModal = payload?.hasLoginModal === true;
  const hasLoggedInActionButton = payload?.hasLoggedInActionButton === true;
  const hasInvoiceSearch = payload?.hasInvoiceSearch === true;
  const phase =
    payload?.phase ||
    (hasLoginModal
      ? "login-modal"
      : hasInvoiceSearch
        ? "invoice-search"
        : hasLoggedInActionButton || payload?.isLoggedInUi === true
          ? "authenticated-shell"
          : "unknown");

  return {
    ...createEmptyPortalFlow(),
    ...payload,
    phase,
    hasLoginModal,
    hasLoggedInActionButton,
    hasInvoiceSearch,
    isLoggedInUi: payload?.isLoggedInUi ?? (phase !== "login-modal" && (hasLoggedInActionButton || hasInvoiceSearch)),
    tabId,
    evidence: Array.isArray(payload?.evidence) ? payload.evidence.slice(0, 10) : [],
    updatedAt: payload?.updatedAt || new Date().toISOString()
  };
}

function userFriendlyServerError(err) {
  const raw = err?.message || String(err);
  const needsHostPermission = /access contents of the page|request permission to access the respective host|must request permission to access/i.test(raw);

  if (needsHostPermission) {
    console.warn("Detected potential missing host permission error:", raw);
    return `Không thể kiểm tra server: Tiện ích chưa thể truy cập trang Hóa đơn điện tử. Vui lòng mở hoặc làm mới trang Hóa đơn Điện tử (hoadondientu.gdt.gov.vn) rồi thử lại.`;
  }

  console.error("Server check error:", err);
  return `Không thể kiểm tra server. Vui lòng thử lại sau.`;
}

function isHostPermissionError(err) {
  const raw = err?.message || String(err);
  return /access contents of the page|request permission to access the respective host|must request permission to access/i.test(raw);
}
async function readPortalFlow() {
  const { portalFlow, authHint } = await chrome.storage.local.get(["portalFlow", "authHint"]);
  return normalizePortalFlow(portalFlow || authHint || {});
}

function buildAuthProbeUrl() {
  const now = new Date();
  return buildUrl(QUERY_PATH, {
    dateFrom: formatDateForApi(now),
    dateTo: formatDateForApi(now, true),
    status: 5,
    size: 1
  });
}

chrome.runtime.onInstalled.addListener((details) => {
  if (details?.reason !== "install") {
    return;
  }

  chrome.storage.local.set({
    appState: {
      lastSyncAt: null,
      lastDateFrom: null,
      lastDateTo: null,
      invoiceCount: 0,
      dataOrigin: "",
      productCatalog: {
        fileName: "",
        loadedAt: null,
        products: [],
        selections: {}
      }
    },
    portalFlow: createEmptyPortalFlow(),
    authHint: createEmptyPortalFlow()
  });
});

chrome.action.onClicked.addListener(() => {
  openExtensionTab().catch((error) => {
    console.error("Không mở được tab extension:", error);
  });
});


chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  try {
    if (!message || !message.type) {
      return false;
    }

    if (message.type === "CHECK_AUTH") {
      resolveAuthState()
        .then((result) => sendResponse({ ok: true, ...result }))
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    if (message.type === "PORTAL_FLOW" || message.type === "AUTH_HINT") {
      const portalFlow = normalizePortalFlow(message.payload || {}, _sender?.tab?.id ?? null);

      chrome.storage.local
        .set({
          portalFlow,
          authHint: portalFlow
        })
        .then(() => sendResponse({ ok: true, data: portalFlow }))
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    if (message.type === "GET_PORTAL_FLOW") {
      readPortalFlow()
        .then((portalFlow) => sendResponse({ ok: true, data: portalFlow }))
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    if (message.type === "OPEN_LOGIN_POPUP") {
      openLoginWindow()
        .then(() => sendResponse({ ok: true }))
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    if (message.type === "OPEN_EXTERNAL_URL") {
      const url = message.payload?.url;
      if (!url) {
        sendResponse({ ok: false, error: "Missing url payload" });
        return true;
      }

      openExternalUrl(url)
        .then(() => sendResponse({ ok: true }))
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

  } catch (e) {
    try {
      console.error("onMessage handler unexpected error:", e);
      sendResponse({ ok: false, error: e?.message || String(e) });
      return true;
    } catch (inner) {
      console.error("onMessage handler failed to sendResponse:", inner);
      return false;
    }
  }

  if (message.type === "FETCH_INVOICES") {
    fetchInvoices(message.payload)
      .then((result) => sendResponse({ ok: true, data: result }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "GET_SYNC_STATE") {
    getInvoiceSyncState()
      .then((result) => sendResponse({ ok: true, data: result }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "GET_CACHED_INVOICES") {
    getCachedInvoices(message.payload || {})
      .then((result) => sendResponse({ ok: true, data: result }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "GET_CACHED_INVOICE") {
    const invoice = message.payload?.invoice;
    if (!invoice) {
      sendResponse({ ok: false, error: "Missing invoice payload" });
      return true;
    }

    getCachedInvoiceByKey(invoice)
      .then((result) => sendResponse({ ok: true, data: result }))
      .catch((err) => sendResponse({ ok: false, error: err?.message || String(err) }));
    return true;
  }

  if (message.type === "FETCH_INVOICE_DETAIL") {
    fetchInvoiceDetail(message.payload)
      .then((result) => sendResponse({ ok: true, data: result }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "UPSERT_INVOICE") {
    const invoice = message.payload?.invoice;
    if (!invoice) {
      sendResponse({ ok: false, error: "Missing invoice payload" });
      return true;
    }

    // Use existing syncInvoicesToLocal to upsert a single invoice into IndexedDB.
    syncInvoicesToLocal([invoice], {})
      .then((res) => sendResponse({ ok: true, data: res }))
      .catch((err) => sendResponse({ ok: false, error: err?.message || String(err) }));
    return true;
  }


  return false;
});

async function openLoginWindow() {
  try {
    await chrome.tabs.create({ url: PORTAL_SEARCH_URL, active: true });
  } catch (e) {
    try {
      // fallback to windows.create if tabs.create fails for some reason
      await chrome.windows.create({ url: PORTAL_SEARCH_URL, type: "popup", width: 1180, height: 900, focused: true });
    } catch (err) {
      throw new Error(err?.message || e?.message || "Không thể mở cổng đăng nhập");
    }
  }
}

async function openExternalUrl(url) {
  try {
    await chrome.tabs.create({ url, active: true });
  } catch (e) {
    try {
      await chrome.windows.create({ url, type: "popup", width: 900, height: 700, focused: true });
    } catch (err) {
      throw new Error(err?.message || e?.message || "Không thể mở liên kết");
    }
  }
}

async function openExtensionTab() {
  const targetUrl = chrome.runtime.getURL("src/popup/popup.html");
  const tabs = await chrome.tabs.query({});
  const existingTab = tabs.find((tab) => tab.url === targetUrl);

  if (existingTab?.id !== undefined) {
    await chrome.tabs.update(existingTab.id, { active: true });
    if (existingTab.windowId !== chrome.windows.WINDOW_ID_CURRENT) {
      await chrome.windows.update(existingTab.windowId, { focused: true });
    }
    return;
  }

  await chrome.tabs.create({ url: targetUrl, active: true });
}

async function resolveAuthState() {
  const portalFlow = await readPortalFlow();
  const url = buildAuthProbeUrl();

  try {
    const response = await withPortalRequestDelay(async () => {
      const requestHeaders = await buildPortalRequestHeaders("query", QUERY_PATH);
      return fetch(url, {
        method: "GET",
        credentials: "include",
        headers: requestHeaders
      });
    });

    return buildAuthProbeResult(response, portalFlow, "server");
  } catch (err) {
    if (isHostPermissionError(err)) {
      try {
        const requestHeaders = await buildPortalRequestHeaders("query", QUERY_PATH);
        const response = await fetchViaPortalPage(url, {
          headers: requestHeaders,
          retries: 1,
          timeout: 3000
        });

        return buildAuthProbeResult(response, portalFlow, "page");
      } catch (pageErr) {
        return {
          isAuthenticated: false,
          phase: portalFlow?.phase || "unknown",
          portalFlow,
          reason: userFriendlyServerError(pageErr)
        };
      }
    }

    return {
      isAuthenticated: false,
      phase: portalFlow?.phase || "unknown",
      portalFlow,
      reason: userFriendlyServerError(err)
    };
  }
}


function buildAuthProbeResult(response, portalFlow, source) {
  if (response.status === 401 || response.status === 403) {
    return {
      isAuthenticated: false,
      phase: "login-required",
      portalFlow,
      status: response.status,
      reason: `Phiên đăng nhập không hợp lệ hoặc đã hết hạn`
    };
  }

  if (!response.ok) {
    return {
      isAuthenticated: false,
      phase: portalFlow?.phase || "unknown",
      portalFlow,
      status: response.status,
      reason: `Kiểm tra đăng nhập thất bại (${source} HTTP ${response.status})`
    };
  }

  return {
    isAuthenticated: true,
    phase: portalFlow?.phase && portalFlow.phase !== "unknown" ? portalFlow.phase : "authenticated-shell",
    portalFlow,
    status: response.status,
    reason:
      portalFlow?.phase === "invoice-search"
        ? "Đã ở màn hình tra cứu hóa đơn"
        : `Phiên đăng nhập đang hoạt động`
  };
}

async function fetchInvoices(payload) {
  if (!payload?.dateFrom || !payload?.dateTo) {
    throw new Error("Thiếu tham số dateFrom/dateTo");
  }

  const syncState = await getInvoiceSyncState();
  const queryStatus = payload.queryStatus ?? 5;
  const scoStatus = payload.scoStatus ?? 8;
  const size = clampPageSize(payload.size ?? MAX_PURCHASE_PAGE_SIZE);

  const [queryResult, scoResult] = await Promise.allSettled([
    fetchInvoiceSource(QUERY_PATH, payload.dateFrom, payload.dateTo, queryStatus, size, "query"),
    fetchInvoiceSource(SCO_QUERY_PATH, payload.dateFrom, payload.dateTo, scoStatus, size, "sco-query")
  ]);

  const queryInvoices = queryResult.status === "fulfilled" ? queryResult.value : [];
  const scoInvoices = scoResult.status === "fulfilled" ? scoResult.value : [];
  const errors = [];

  if (queryResult.status === "rejected") {
    errors.push(`query: ${queryResult.reason?.message || String(queryResult.reason)}`);
  }

  if (scoResult.status === "rejected") {
    errors.push(`sco-query: ${scoResult.reason?.message || String(scoResult.reason)}`);
  }

  if (!queryInvoices.length && !scoInvoices.length) {
    const shouldFallbackLocal = isLikelyNetworkFailure(queryResult.reason) || isLikelyNetworkFailure(scoResult.reason);
    if (shouldFallbackLocal) {
      const cached = await getCachedInvoices({
        dateFrom: payload.dateFrom,
        dateTo: payload.dateTo
      });

      if (cached.invoices.length) {
        return {
          invoices: cached.invoices,
          stats: {
            queryCount: 0,
            scoCount: 0,
            total: cached.invoices.length,
            cachedCount: cached.invoices.length,
            inserted: 0,
            updated: 0,
            skipped: 0
          },
          errors,
          servedFromLocal: true,
          source: "local",
          requestedDateFrom: payload.dateFrom,
          requestedDateTo: payload.dateTo,
          lastSyncAt: cached.syncState?.lastSyncAt || null
        };
      }
    }

    throw new Error(errors.length ? `Không tải được dữ liệu: ${errors.join(" | ")}` : "Không tải được dữ liệu hóa đơn");
  }

  const merged = dedupeInvoices([...queryInvoices, ...scoInvoices]);
  const syncResult = await syncInvoicesToLocal(merged, {
    requestedDateFrom: payload.dateFrom,
    requestedDateTo: payload.dateTo,
    queryCount: queryInvoices.length,
    scoCount: scoInvoices.length
  });

  const { appState: existingAppState = {} } = await chrome.storage.local.get("appState");

  await chrome.storage.local.set({
    appState: {
      ...existingAppState,
      lastSyncAt: new Date().toISOString(),
      lastDateFrom: payload.dateFrom,
      lastDateTo: payload.dateTo,
      invoiceCount: syncResult.totalCount,
      dataOrigin: "network",
      productCatalog: existingAppState.productCatalog || {
        fileName: "",
        loadedAt: null,
        products: [],
        selections: {}
      }
    }
  });

  return {
    invoices: syncResult.invoices,
    stats: {
      queryCount: queryInvoices.length,
      scoCount: scoInvoices.length,
      total: syncResult.totalCount,
      inserted: syncResult.inserted,
      updated: syncResult.updated,
      skipped: syncResult.skipped
    },
    errors,
    servedFromLocal: false,
    source: "network",
    requestedDateFrom: payload.dateFrom,
    requestedDateTo: payload.dateTo,
    lastSyncAt: new Date().toISOString()
  };
}

async function getInvoiceSyncState() {
  const [meta, { appState = {} }] = await Promise.all([
    getInvoiceSyncMeta(),
    chrome.storage.local.get("appState")
  ]);

  return {
    lastSyncAt: meta?.lastSyncAt || appState.lastSyncAt || null,
    lastDateFrom: meta?.lastDateFrom || appState.lastDateFrom || null,
    lastDateTo: meta?.lastDateTo || appState.lastDateTo || null,
    invoiceCount: meta?.invoiceCount ?? appState.invoiceCount ?? 0,
    dataOrigin: meta?.dataOrigin || appState.dataOrigin || "",
    updatedAt: meta?.updatedAt || appState.lastSyncAt || null
  };
}

async function getCachedInvoices({ dateFrom, dateTo, limit } = {}) {
  const invoices = await queryInvoicesByDateRange(dateFrom, dateTo, limit);
  const syncState = await getInvoiceSyncState();

  return {
    invoices,
    syncState
  };
}

async function syncInvoicesToLocal(invoices, syncMeta) {
  const db = await openInvoiceDatabase();
  const now = new Date().toISOString();
  const ordered = Array.isArray(invoices) ? invoices : [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  await new Promise((resolve, reject) => {
    const tx = db.transaction([INVOICE_STORE_NAME, INVOICE_META_STORE_NAME], "readwrite");
    const store = tx.objectStore(INVOICE_STORE_NAME);
    const metaStore = tx.objectStore(INVOICE_META_STORE_NAME);
    let index = 0;

    const processNext = () => {
      if (index >= ordered.length) {
        const metaReq = metaStore.put({
          key: INVOICE_SYNC_META_KEY,
          ...syncMeta,
          totalCount: ordered.length,
          inserted,
          updated,
          skipped,
          dataOrigin: "network",
          updatedAt: now
        });

        metaReq.onerror = () => reject(metaReq.error || new Error("Không ghi được metadata đồng bộ"));
        return;
      }

      const invoice = ordered[index++];
      const invoiceKey = buildInvoiceKey(invoice);
      const nextRecord = buildInvoiceRecord(invoice, now);
      const getReq = store.get(invoiceKey);

      getReq.onerror = () => reject(getReq.error || new Error("Không đọc được hóa đơn đã lưu"));
      getReq.onsuccess = () => {
        const existing = getReq.result;
        if (!existing) {
          store.put(nextRecord);
          inserted += 1;
          processNext();
          return;
        }

        if (shouldReplaceStoredInvoice(existing, nextRecord)) {
          store.put(nextRecord);
          updated += 1;
        } else {
          skipped += 1;
        }

        processNext();
      };
    };

    tx.onerror = () => reject(tx.error || new Error("Không ghi được hóa đơn vào IndexedDB"));
    tx.onabort = () => reject(tx.error || new Error("Giao dịch IndexedDB bị hủy"));
    tx.oncomplete = () => resolve();

    processNext();
  });

  return {
    invoices: ordered,
    totalCount: ordered.length,
    inserted,
    updated,
    skipped
  };
}

async function queryInvoicesByDateRange(dateFrom, dateTo, limit) {
  const db = await openInvoiceDatabase();
  const lowerBound = parseApiDateToTimestamp(dateFrom, false);
  const upperBound = parseApiDateToTimestamp(dateTo, true);
  const results = [];

  await new Promise((resolve, reject) => {
    const tx = db.transaction(INVOICE_STORE_NAME, "readonly");
    const store = tx.objectStore(INVOICE_STORE_NAME);
    const index = store.index("issuedAtMs");
    const range = IDBKeyRange.bound(lowerBound, upperBound);
    const cursorReq = index.openCursor(range, "prev");

    cursorReq.onerror = () => reject(cursorReq.error || new Error("Không đọc được hóa đơn local"));
    cursorReq.onsuccess = () => {
      const cursor = cursorReq.result;
      if (!cursor) {
        resolve();
        return;
      }

      results.push(cursor.value.invoice);
      if (Number.isFinite(limit) && limit > 0 && results.length >= limit) {
        resolve();
        return;
      }

      cursor.continue();
    };

    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Không đọc được hóa đơn local"));
  });

  return results;
}

async function getCachedInvoiceByKey(invoice) {
  const db = await openInvoiceDatabase();
  const invoiceKey = buildInvoiceKey(invoice);

  return new Promise((resolve, reject) => {
    const tx = db.transaction(INVOICE_STORE_NAME, "readonly");
    const store = tx.objectStore(INVOICE_STORE_NAME);
    const req = store.get(invoiceKey);

    req.onerror = () => reject(req.error || new Error("Không đọc được hóa đơn local"));
    req.onsuccess = () => {
      const record = req.result;
      if (!record) {
        resolve(null);
        return;
      }
      resolve(record.invoice || null);
    };
    tx.onerror = () => reject(tx.error || new Error("Không đọc được hóa đơn local"));
  });
}

async function getInvoiceSyncMeta() {
  const db = await openInvoiceDatabase();

  return new Promise((resolve, reject) => {
    const tx = db.transaction(INVOICE_META_STORE_NAME, "readonly");
    const store = tx.objectStore(INVOICE_META_STORE_NAME);
    const req = store.get(INVOICE_SYNC_META_KEY);

    req.onerror = () => reject(req.error || new Error("Không đọc được metadata đồng bộ"));
    req.onsuccess = () => resolve(req.result || null);
    tx.onerror = () => reject(tx.error || new Error("Không đọc được metadata đồng bộ"));
  });
}

async function openInvoiceDatabase() {
  if (typeof indexedDB === "undefined") {
    throw new Error("Trình duyệt không hỗ trợ IndexedDB");
  }

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(INVOICE_DB_NAME, INVOICE_DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = event.target.result;

      if (!db.objectStoreNames.contains(INVOICE_STORE_NAME)) {
        const store = db.createObjectStore(INVOICE_STORE_NAME, { keyPath: "invoiceKey" });
        store.createIndex("issuedAtMs", "issuedAtMs", { unique: false });
        store.createIndex("syncedAt", "syncedAt", { unique: false });
      }

      if (!db.objectStoreNames.contains(INVOICE_META_STORE_NAME)) {
        db.createObjectStore(INVOICE_META_STORE_NAME, { keyPath: "key" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Không mở được IndexedDB hóa đơn"));
  });
}

function buildInvoiceRecord(invoice, syncedAt) {
  const invoiceKey = buildInvoiceKey(invoice);
  const issuedAtMs = parseInvoiceTimestamp(invoice?.tdlap || invoice?.ntao || invoice?.ncnhat);
  const updatedAtMs = parseInvoiceTimestamp(
    invoice?.ncnhat || invoice?.ncnhattt || invoice?.ncnhatdl || invoice?.ncnhatts || invoice?.updatedAt || invoice?.modifiedAt
  );
  const contentSignature = buildInvoiceContentSignature(invoice);

  return {
    invoiceKey,
    invoice: {
      ...invoice,
      _source: invoice?._source || "network"
    },
    issuedAtMs,
    updatedAtMs,
    contentSignature,
    syncedAt,
    source: invoice?._source || "network"
  };
}

function shouldReplaceStoredInvoice(existingRecord, nextRecord) {
  const existingInvoice = existingRecord?.invoice || existingRecord;
  const nextInvoice = nextRecord?.invoice || nextRecord;

  if (!existingInvoice) {
    return true;
  }

  if (existingRecord?.contentSignature && nextRecord?.contentSignature && existingRecord.contentSignature === nextRecord.contentSignature) {
    return false;
  }

  const existingScore = scoreInvoice(existingInvoice);
  const nextScore = scoreInvoice(nextInvoice);

  if (nextScore !== existingScore) {
    return nextScore > existingScore;
  }

  const existingUpdatedAtMs = Number(existingRecord?.updatedAtMs || 0);
  const nextUpdatedAtMs = Number(nextRecord?.updatedAtMs || 0);
  if (nextUpdatedAtMs !== existingUpdatedAtMs) {
    return nextUpdatedAtMs > existingUpdatedAtMs;
  }

  const existingIssuedAtMs = Number(existingRecord?.issuedAtMs || 0);
  const nextIssuedAtMs = Number(nextRecord?.issuedAtMs || 0);
  if (nextIssuedAtMs !== existingIssuedAtMs) {
    return nextIssuedAtMs > existingIssuedAtMs;
  }

  const existingSyncedAt = parseInvoiceTimestamp(existingRecord?.syncedAt);
  const nextSyncedAt = parseInvoiceTimestamp(nextRecord?.syncedAt);
  if (nextSyncedAt !== existingSyncedAt) {
    return nextSyncedAt > existingSyncedAt;
  }

  return false;
}

function buildInvoiceContentSignature(invoice) {
  const signatureParts = [
    invoice?.mhdon ?? "",
    invoice?.id ?? "",
    invoice?.khhdon ?? "",
    invoice?.shdon ?? "",
    invoice?.nbmst ?? "",
    invoice?.tdlap ?? "",
    invoice?.ncnhat ?? "",
    invoice?.tgtttbso ?? "",
    invoice?.tgtcthue ?? "",
    invoice?.tgtthue ?? "",
    Array.isArray(invoice?.thttltsuat)
      ? invoice.thttltsuat.map((item) => `${item?.tsuat ?? ""}:${item?.tthue ?? ""}`).join("|")
      : "",
    invoice?.nky ?? "",
    invoice?.nbcks ?? ""
  ];

  return signatureParts
    .map((part) => normalizeInvoiceSignatureValue(part))
    .join("~");
}

function normalizeInvoiceSignatureValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (Array.isArray(value)) {
    return value.map((item) => normalizeInvoiceSignatureValue(item)).join(",");
  }

  if (typeof value === "object") {
    return JSON.stringify(value);
  }

  return String(value).trim();
}

function parseInvoiceTimestamp(value) {
  if (!value) {
    return 0;
  }

  const time = new Date(value).getTime();
  return Number.isFinite(time) ? time : 0;
}

function parseApiDateToTimestamp(value, endOfDay = false) {
  const normalized = normalizeApiDateRangeValue(value, endOfDay);
  return normalized ? new Date(normalized).getTime() : 0;
}

function normalizeApiDateRangeValue(value, endOfDay = false) {
  const text = String(value || "").trim();
  const match = text.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (!match) {
    return "";
  }

  const [, dd, mm, yyyy] = match;
  return `${yyyy}-${mm}-${dd}T${endOfDay ? "23:59:59" : "00:00:00"}`;
}

function isLikelyNetworkFailure(error) {
  const message = String(error?.message || error || "").toLowerCase();
  return (
    message.includes("timeout") ||
    message.includes("failed to fetch") ||
    message.includes("network") ||
    message.includes("fetch failed") ||
    message.includes("load failed") ||
    message.includes("offline")
  );
}

async function fetchInvoiceSource(path, dateFrom, dateTo, status, size, source) {
  const url = buildUrl(path, {
    dateFrom,
    dateTo,
    status,
    size
  });

  try {
    const response = await fetchWithFreshPortalHeaders(url, {
      source,
      path,
      retries: 1,
      timeout: 3000
    });

    if (response.status === 401 || response.status === 403) {
      throw new Error("Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.");
    }

    if (!response.ok) {
      throw new Error(formatPortalFetchFailureMessage(source, response));
    }

    const json = safeParseJson(response.body, source);
    const invoices = Array.isArray(json?.datas) ? json.datas : [];

    return invoices.map((item) => ({
      ...item,
      _source: source
    }));
  } catch (error) {
    if (error?.status === 401 || error?.status === 403) {
      throw new Error("Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.");
    }

    throw error;
  }
}

async function fetchInvoiceDetail(payload) {
  const nbmst = String(payload?.nbmst || "").trim();
  const khhdon = String(payload?.khhdon || "").trim();
  const shdon = String(payload?.shdon ?? "").trim();
  const khmshdon = String(payload?.khmshdon ?? "").trim();
  const source = payload?.source === "sco-query" ? "sco-query" : "query";
  const detailPath = DETAIL_PATHS[source];

  if (!nbmst || !khhdon || !shdon || !khmshdon) {
    throw new Error("Thiếu tham số để lấy chi tiết hóa đơn");
  }

  const url = buildDetailUrl({ nbmst, khhdon, shdon, khmshdon, detailPath });

  try {
    const response = await fetchWithFreshPortalHeaders(url, {
      source,
      path: detailPath,
      actionOverride: "Xem hóa đơn (hóa đơn máy tính tiền mua vào)",
      retries: 1,
      timeout: 3000
    });

    if (response.status === 401 || response.status === 403) {
      throw new Error("Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.");
    }

    if (!response.ok) {
      throw new Error(formatPortalFetchFailureMessage("query", response, "Tải chi tiết hóa đơn thất bại"));
    }

    return safeParseJson(response.body, "detail");
  } catch (error) {
    if (error?.status === 401 || error?.status === 403) {
      throw new Error("Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.");
    }

    throw error;
  }
}

async function fetchViaPortalPage(url, { retries = 3, timeout = PORTAL_PAGE_FETCH_TIMEOUT_MS, headers = {} } = {}) {
  const tab = await findPortalTab();
  if (!tab?.id) {
    const err = "Không tìm thấy tab hoadondientu đang mở để lấy phiên đăng nhập.";
    console.warn("fetchViaPortalPage: no portal tab found", { url });
    throw new Error(err);
  }

  const message = {
    type: "PAGE_FETCH",
    url,
    method: "GET",
    headers,
    retries: Math.max(1, Math.trunc(retries)),
    timeout
  };

  let response = await sendPortalMessage(tab.id, message);
  if (response && !isPortalTimeoutResponse(response) && response.ok !== false) {
    return response;
  }

  // If bridge is not ready yet (or page fetch timed out), inject and retry once.
  console.warn("fetchViaPortalPage: initial portal bridge missing or timed out, attempting inject and retry", {
    tabId: tab.id,
    url,
    response
  });

  await ensurePortalBridgeInjected(tab.id);
  response = await sendPortalMessage(tab.id, message);
  if (response && response.ok !== false) {
    return response;
  }

  const errorDetail = response && response.error ? String(response.error) : "no response from portal tab";
  console.warn("fetchViaPortalPage: failed after retry", { tabId: tab.id, url, error: errorDetail, response });
  throw new Error(`Không nhận được phản hồi từ tab Hóa đơn điện tử: ${errorDetail}`);
}

function isPortalTimeoutResponse(response) {
  const error = String(response?.error || "").toLowerCase().trim();
  return error === "timeout" || error.includes("timeout");
}

async function buildPortalRequestHeaders(source, path, actionOverride = null) {
  const baseHeaders = {
    Accept: "application/json, text/plain, */*",
    "Accept-Language": "vi"
  };

  const requestContext = await getPortalRequestContext();
  if (requestContext?.authorization) {
    baseHeaders.Authorization = requestContext.authorization;
  }

  const actionValue =
    actionOverride ||
    (source === "sco-query"
      ? requestContext?.actionScoQuery || "Tìm kiếm (hóa đơn máy tính tiền mua vào)"
      : requestContext?.actionQuery || "Tìm kiếm (hóa đơn mua vào)");

  baseHeaders.action = encodeURIComponent(actionValue);
  baseHeaders["end-point"] = requestContext?.endPoint || "/tra-cuu/tra-cuu-hoa-don";

  return baseHeaders;
}

function buildDetailUrl({ nbmst, khhdon, shdon, khmshdon, detailPath }) {
  const params = new URLSearchParams({
    nbmst,
    khhdon,
    shdon,
    khmshdon
  });

  return `${QUERY_BASE}${detailPath}?${params.toString()}`;
}

function withPortalRequestDelay(executor) {
  const run = async () => {
    const now = Date.now();
    const elapsed = now - lastPortalRequestAt;
    if (elapsed < PORTAL_REQUEST_DELAY_MS) {
      await delay(PORTAL_REQUEST_DELAY_MS - elapsed);
    }

    lastPortalRequestAt = Date.now();
    return executor();
  };

  const next = portalRequestChain.then(run, run);
  portalRequestChain = next.then(() => undefined, () => undefined);
  return next;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function promiseWithTimeout(promise, ms, message = "timeout") {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      reject(new Error(message));
    }, ms);

    promise
      .then((v) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve(v);
      })
      .catch((e) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(e);
      });
  });
}

async function fetchWithTimeout(url, init = {}, timeout = 5000) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeout);
  try {
    const merged = { ...init, signal: controller.signal };
    const resp = await fetch(url, merged);
    clearTimeout(id);
    return resp;
  } catch (e) {
    clearTimeout(id);
    if (e && e.name === "AbortError") {
      throw new Error("timeout");
    }
    throw e;
  }
}

async function getPortalRequestContext() {
  const tab = await findPortalTab();
  if (!tab?.id) {
    return null;
  }

  const message = { type: "GET_PORTAL_REQUEST_CONTEXT" };
  let response = await sendPortalMessage(tab.id, message);
  if (response?.ok) {
    return response.data || null;
  }

  await ensurePortalBridgeInjected(tab.id);
  response = await sendPortalMessage(tab.id, message);
  if (response?.ok) {
    return response.data || null;
  }

  return null;
}

async function sendPortalMessage(tabId, message) {
  return new Promise((resolve) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      if (chrome.runtime.lastError) {
        const lastErr = chrome.runtime.lastError;
        const errorMessage = lastErr && lastErr.message ? lastErr.message : String(lastErr || "");

        // Return a structured error object instead of null so callers can distinguish cases
        if (isMissingReceiverError(errorMessage)) {
          resolve({ ok: false, error: "missing_receiver", details: errorMessage });
          return;
        }

        console.warn("sendPortalMessage: chrome.runtime.lastError", { tabId, messageType: message?.type, errorMessage });

        resolve({ ok: false, error: "runtime_error", details: errorMessage });
        return;
      }

      // Normalize response shape to always include ok flag when possible
      if (response && typeof response === "object" && !Object.prototype.hasOwnProperty.call(response, "ok")) {
        resolve({ ok: true, ...response });
        return;
      }

      resolve(response || { ok: false, error: "no_response" });
    });
  });
}

async function ensurePortalBridgeInjected(tabId) {
  const files = ["src/content/content.js", "src/content/page-fetcher.js"];
  try {
    // Inject into all frames to maximize chance bridge is available for framed apps
    await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      files
    });

    // Verify injection by executing a small check in page context. Retry a few times.
    const maxChecks = 3;
    for (let attempt = 1; attempt <= maxChecks; attempt += 1) {
      try {
        const results = await chrome.scripting.executeScript({
          target: { tabId },
          func: () => {
            try {
              return Boolean(window.__EXT_PAGE_FETCH_INJECTED) && Array.isArray(window.__HOADON_LOGIN_MODAL_SELECTORS);
            } catch (e) {
              return false;
            }
          }
        });

        // `results` is an array of InjectionResult; consider success if any frame returned truthy
        const ok = Array.isArray(results) && results.some((r) => r && r.result === true);
        if (ok) return;
      } catch (inner) {
        // ignore and retry
      }

      // small backoff between checks
      await new Promise((res) => setTimeout(res, 250 * attempt));
    }

    // If we reach here, injection didn't appear to take effect
    console.warn("ensurePortalBridgeInjected: injection completed but bridge not detected after retries", { tabId, files });
  } catch (error) {
    const msg = String(error?.message || error || "");
    if (!isMissingReceiverError(msg)) {
      throw error;
    }
    // If missing receiver, log and return — caller will handle fallback
    console.warn("ensurePortalBridgeInjected: scripting.executeScript missing receiver", { tabId, message: msg });
  }
}

function isMissingReceiverError(message) {
  const text = String(message || "").toLowerCase();
  return text.includes("receiving end does not exist") || text.includes("could not establish connection");
}

function buildPortalFetchInit(url, headers = {}) {
  return {
    method: "GET",
    credentials: "include",
    mode: "cors",
    referrer: "https://hoadondientu.gdt.gov.vn/",
    referrerPolicy: "strict-origin-when-cross-origin",
    headers: {
      ...headers
    }
  };
}

async function fetchWithFreshPortalHeaders(url, { source, path, actionOverride = null, retries = 1, timeout = 15000 } = {}) {
  return withPortalRequestDelay(async () => {
    const requestHeaders = await buildPortalRequestHeaders(source, path, actionOverride);
    return fetchViaPortalPage(url, {
      retries,
      timeout,
      headers: requestHeaders
    });
  });
}

async function findPortalTab() {
  const portalFlow = await readPortalFlow();

  if (portalFlow?.tabId !== null && portalFlow?.tabId !== undefined) {
    try {
      const tab = await chrome.tabs.get(portalFlow.tabId);
      if (tab?.id && tab.url && tab.url.includes("hoadondientu.gdt.gov.vn")) {
        return tab;
      }
    } catch (e) {
      // fall through to URL search
    }
  }

  const tabs = await chrome.tabs.query({});
  return (
    tabs.find((tab) => tab.active && tab.url && tab.url.includes("hoadondientu.gdt.gov.vn")) ||
    tabs.find((tab) => tab.url && tab.url.includes("hoadondientu.gdt.gov.vn")) ||
    null
  );
}

function safeParseJson(text, source) {
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Phản hồi ${source} không phải JSON hợp lệ`);
  }
}

function formatPortalFetchFailureMessage(source, response, prefix = "Tải dữ liệu thất bại") {
  if (response?.status === 401 || response?.status === 403) {
    return "Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.";
  }

  const bodySnippet = typeof response?.body === "string" && response.body.trim()
    ? ` - ${response.body.trim().slice(0, 200)}`
    : "";
  const responseError = String(response?.error || "").trim();

  if (typeof response?.status === "number") {
    return `${prefix} (${source}): ${response.status}${bodySnippet}`;
  }

  if (responseError) {
    return `${prefix} (${source}): ${responseError}`;
  }

  return `${prefix} (${source}): không nhận được phản hồi từ portal`;
}

function dedupeInvoices(invoices) {
  const map = new Map();

  for (const invoice of invoices) {
    const key = buildInvoiceKey(invoice);
    if (!map.has(key)) {
      map.set(key, invoice);
      continue;
    }

    // Prefer signed or richer record if duplicate appears from both endpoints.
    const existing = map.get(key);
    const existingScore = scoreInvoice(existing);
    const nextScore = scoreInvoice(invoice);
    if (nextScore > existingScore) {
      map.set(key, invoice);
    }
  }

  return Array.from(map.values()).sort((a, b) => {
    const da = new Date(a?.tdlap || 0).getTime();
    const db = new Date(b?.tdlap || 0).getTime();
    return db - da;
  });
}

function buildInvoiceKey(invoice) {
  if (invoice?.mhdon) {
    return `mhdon:${invoice.mhdon}`;
  }

  if (invoice?.id) {
    return `id:${invoice.id}`;
  }

  return [
    invoice?.khhdon ?? "",
    invoice?.shdon ?? "",
    invoice?.nbmst ?? "",
    invoice?.tdlap ?? ""
  ].join("|");
}

function scoreInvoice(invoice) {
  let score = 0;
  if (invoice?.nky) score += 2;
  if (invoice?.nbcks) score += 2;
  if (invoice?.thttltsuat?.length) score += 1;
  if (invoice?._source === "sco-query") score += 1;
  return score;
}

function buildUrl(path, { dateFrom, dateTo, status, size }) {
  const search = `tdlap=ge=${dateFrom};tdlap=le=${dateTo};ttxly==${status}`;
  const pageSize = clampPageSize(size);

  // Use raw query format to match the portal's own successful requests.
  return `${QUERY_BASE}${path}?sort=tdlap:desc&size=${pageSize}&search=${search}`;
}

function clampPageSize(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return MAX_PURCHASE_PAGE_SIZE;
  }

  return Math.min(Math.trunc(parsed), MAX_PURCHASE_PAGE_SIZE);
}

function formatDateForApi(value, endOfDay = false) {
  const date = new Date(value);
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = String(date.getFullYear());
  const suffix = endOfDay ? "23:59:59" : "00:00:00";

  return `${dd}/${mm}/${yyyy}T${suffix}`;
}
