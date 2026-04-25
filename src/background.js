// =============================================================================
// CONSTANTS
// =============================================================================

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
const AUTH_PROBE_TIMEOUT_MS = 8000;
const BRIDGE_INJECT_TIMEOUT_MS = 5000;
const BRIDGE_CHECK_RETRIES = 3;
const BRIDGE_CHECK_DELAY_MS = 250;

const INVOICE_DB_NAME = "hoadondientu_invoices";
const INVOICE_DB_VERSION = 1;
const INVOICE_STORE_NAME = "invoices";
const INVOICE_META_STORE_NAME = "sync_meta";
const INVOICE_SYNC_META_KEY = "invoice_sync_meta";

// =============================================================================
// REQUEST QUEUE — serialise portal fetches with rate-limiting
// =============================================================================

let portalRequestChain = Promise.resolve();
let lastPortalRequestAt = 0;

// =============================================================================
// ERROR CLASSIFICATION
// =============================================================================

/**
 * Returns true only for the specific Chrome scripting host-permission error.
 * Deliberately narrow so ordinary CORS/network errors are not mis-classified.
 */
function isHostPermissionError(err) {
  const raw = err?.message || String(err);
  return /Extension manifest must request permission to access the respective host|must request permission to access the respective host/i.test(raw);
}

function isNetworkError(err) {
  const msg = String(err?.message || err || "").toLowerCase();
  return (
    msg.includes("failed to fetch") ||
    msg.includes("fetch failed") ||
    msg.includes("load failed") ||
    msg.includes("network") ||
    msg.includes("offline")
  );
}

function isTimeoutError(err) {
  const msg = String(err?.message || err || "").toLowerCase();
  return msg === "timeout" || msg.includes("timeout") || msg.includes("timed out");
}

function isAuthError(status) {
  return status === 401 || status === 403;
}

function classifyFetchError(err) {
  if (isTimeoutError(err)) return "timeout";
  if (isHostPermissionError(err)) return "host_permission";
  if (isNetworkError(err)) return "network";
  return "unknown";
}

function isMissingReceiverError(message) {
  const text = String(message || "").toLowerCase();
  return (
    text.includes("receiving end does not exist") ||
    text.includes("could not establish connection")
  );
}

// =============================================================================
// USER-FACING ERROR MESSAGES
// =============================================================================

function userFriendlyError(err) {
  const kind = classifyFetchError(err);

  switch (kind) {
    case "timeout":
      return "Yêu cầu quá thời gian chờ. Vui lòng thử lại.";
    case "host_permission":
      return "Tiện ích chưa thể truy cập trang Hóa đơn điện tử. Vui lòng mở hoặc làm mới trang hoadondientu.gdt.gov.vn rồi thử lại.";
    case "network":
      return "Không có kết nối mạng hoặc máy chủ không phản hồi. Vui lòng kiểm tra kết nối và thử lại.";
    default:
      return `Không thể kết nối máy chủ. Vui lòng thử lại sau. (${err?.message || err})`;
  }
}

function formatPortalFetchFailureMessage(source, response, prefix = "Tải dữ liệu thất bại") {
  if (isAuthError(response?.status)) {
    return "Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.";
  }

  const bodySnippet =
    typeof response?.body === "string" && response.body.trim()
      ? ` — ${response.body.trim().slice(0, 200)}`
      : "";
  const responseError = String(response?.error || "").trim();

  if (typeof response?.status === "number") {
    return `${prefix} (${source}): HTTP ${response.status}${bodySnippet}`;
  }

  if (responseError) {
    return `${prefix} (${source}): ${responseError}`;
  }

  return `${prefix} (${source}): không nhận được phản hồi từ portal`;
}

// =============================================================================
// PORTAL FLOW STATE
// =============================================================================

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
    isLoggedInUi:
      payload?.isLoggedInUi ??
      (phase !== "login-modal" && (hasLoggedInActionButton || hasInvoiceSearch)),
    tabId,
    evidence: Array.isArray(payload?.evidence) ? payload.evidence.slice(0, 10) : [],
    updatedAt: payload?.updatedAt || new Date().toISOString()
  };
}

async function readPortalFlow() {
  const { portalFlow, authHint } = await chrome.storage.local.get(["portalFlow", "authHint"]);
  return normalizePortalFlow(portalFlow || authHint || {});
}

// =============================================================================
// TIMING UTILITIES
// =============================================================================

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Wraps a promise with a hard timeout.
 * Rejects with a typed timeout error so callers can distinguish it cleanly.
 */
function withTimeout(promise, ms, label = "timeout") {
  return new Promise((resolve, reject) => {
    let done = false;

    const timer = setTimeout(() => {
      if (done) return;
      done = true;
      reject(Object.assign(new Error(`timeout: ${label}`), { kind: "timeout" }));
    }, ms);

    promise.then(
      (v) => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        resolve(v);
      },
      (e) => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        reject(e);
      }
    );
  });
}

/**
 * Serialises portal requests with a minimum inter-request delay.
 */
function withPortalRequestDelay(executor) {
  const run = async () => {
    const elapsed = Date.now() - lastPortalRequestAt;
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

// =============================================================================
// FETCH HELPERS
// =============================================================================

/**
 * Builds the standard fetch init object for direct service-worker requests.
 * Includes mode:"cors" which is required when using credentials:"include"
 * cross-origin — omitting it causes a browser error that looks like a host
 * permission error.
 */
function buildPortalFetchInit(headers = {}) {
  return {
    method: "GET",
    credentials: "include",
    mode: "cors",
    referrer: `${PORTAL_BASE_URL}/`,
    referrerPolicy: "strict-origin-when-cross-origin",
    headers: { ...headers }
  };
}

/**
 * fetch() with an AbortController-based timeout.
 * Throws a typed timeout error on expiry.
 */
async function fetchWithTimeout(url, init = {}, timeoutMs = 5000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, { ...init, signal: controller.signal });
    clearTimeout(timer);
    return response;
  } catch (err) {
    clearTimeout(timer);
    if (err?.name === "AbortError") {
      throw Object.assign(new Error(`timeout: ${url}`), { kind: "timeout" });
    }
    throw err;
  }
}

// =============================================================================
// PORTAL REQUEST HEADERS
// =============================================================================

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

// =============================================================================
// AUTH PROBE
// =============================================================================

function buildAuthProbeUrl() {
  const now = new Date();
  return buildUrl(QUERY_PATH, {
    dateFrom: formatDateForApi(now),
    dateTo: formatDateForApi(now, true),
    status: 5,
    size: 1
  });
}

/**
 * Resolves current auth state by probing the API directly from the service
 * worker first, then falling back to the portal page bridge if needed.
 *
 * FIX: Uses buildPortalFetchInit() which includes mode:"cors" — previously the
 * direct fetch omitted this, causing a browser CORS rejection that was
 * mis-classified as a host-permission error and sent the flow into the
 * page-bridge fallback unconditionally.
 */
async function resolveAuthState() {
  const portalFlow = await readPortalFlow();
  const url = buildAuthProbeUrl();

  // --- Attempt 1: direct fetch from service worker ---
  try {
    const response = await withTimeout(
      withPortalRequestDelay(async () => {
        const headers = await buildPortalRequestHeaders("query", QUERY_PATH);
        return fetchWithTimeout(url, buildPortalFetchInit(headers), AUTH_PROBE_TIMEOUT_MS);
      }),
      AUTH_PROBE_TIMEOUT_MS + PORTAL_REQUEST_DELAY_MS,
      "auth-probe-direct"
    );

    return buildAuthProbeResult(response, portalFlow, "direct");
  } catch (directErr) {
    const kind = classifyFetchError(directErr);

    // Only fall back through the page bridge for genuine host-permission errors.
    // CORS errors (which no longer occur after the mode:"cors" fix), timeouts,
    // and network errors are surfaced directly.
    if (kind !== "host_permission") {
      console.warn("resolveAuthState: direct fetch failed", { kind, message: directErr?.message });
      return {
        isAuthenticated: false,
        phase: portalFlow?.phase || "unknown",
        portalFlow,
        portalFlowTrusted: false,
        reason: userFriendlyError(directErr)
      };
    }

    console.warn("resolveAuthState: host permission error on direct fetch, trying page bridge", {
      message: directErr?.message
    });
  }

  // --- Attempt 2: page bridge fallback (host permission scenario) ---
  try {
    const headers = await buildPortalRequestHeaders("query", QUERY_PATH);
    const response = await withTimeout(
      fetchViaPortalPage(url, { headers, retries: 1, timeout: AUTH_PROBE_TIMEOUT_MS }),
      AUTH_PROBE_TIMEOUT_MS + 2000,
      "auth-probe-bridge"
    );

    return buildAuthProbeResult(response, portalFlow, "bridge");
  } catch (bridgeErr) {
    console.warn("resolveAuthState: bridge fallback also failed", { message: bridgeErr?.message });
    return {
      isAuthenticated: false,
      phase: portalFlow?.phase || "unknown",
      portalFlow,
      portalFlowTrusted: false,
      reason: userFriendlyError(bridgeErr)
    };
  }
}

function buildAuthProbeResult(response, portalFlow, source) {
  if (isAuthError(response?.status)) {
    return {
      isAuthenticated: false,
      phase: "login-required",
      portalFlow,
      status: response.status,
      portalFlowTrusted: true,
      reason: "Phiên đăng nhập không hợp lệ hoặc đã hết hạn"
    };
  }

  if (!response?.ok) {
    return {
      isAuthenticated: false,
      phase: portalFlow?.phase || "unknown",
      portalFlow,
      status: response?.status,
      portalFlowTrusted: true,
      reason: `Kiểm tra đăng nhập thất bại (${source} HTTP ${response?.status ?? "?"})`
    };
  }

  return {
    isAuthenticated: true,
    phase:
      portalFlow?.phase && portalFlow.phase !== "unknown"
        ? portalFlow.phase
        : "authenticated-shell",
    portalFlow,
    status: response.status,
    portalFlowTrusted: true,
    reason:
      portalFlow?.phase === "invoice-search"
        ? "Đã ở màn hình tra cứu hóa đơn"
        : "Phiên đăng nhập đang hoạt động"
  };
}

// =============================================================================
// PAGE BRIDGE
// =============================================================================

async function fetchViaPortalPage(url, { retries = 3, timeout = PORTAL_PAGE_FETCH_TIMEOUT_MS, headers = {} } = {}) {
  const tab = await findPortalTab();
  if (!tab?.id) {
    throw new Error(
      "Không tìm thấy tab hoadondientu đang mở để lấy phiên đăng nhập. Vui lòng mở trang hoadondientu.gdt.gov.vn."
    );
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

  // Bridge not ready — inject and retry once
  console.warn("fetchViaPortalPage: bridge missing or timed out, injecting and retrying", {
    tabId: tab.id,
    url
  });

  await withTimeout(
    ensurePortalBridgeInjected(tab.id),
    BRIDGE_INJECT_TIMEOUT_MS,
    "bridge-inject"
  );

  response = await sendPortalMessage(tab.id, message);

  if (response && response.ok !== false) {
    return response;
  }

  const errDetail = response?.error ? String(response.error) : "no response from portal tab";
  throw new Error(`Không nhận được phản hồi từ tab Hóa đơn điện tử: ${errDetail}`);
}

function isPortalTimeoutResponse(response) {
  const error = String(response?.error || "").toLowerCase().trim();
  return error === "timeout" || error.includes("timeout");
}

async function sendPortalMessage(tabId, message) {
  return new Promise((resolve) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      if (chrome.runtime.lastError) {
        const errMsg = chrome.runtime.lastError?.message || String(chrome.runtime.lastError || "");

        if (isMissingReceiverError(errMsg)) {
          resolve({ ok: false, error: "missing_receiver", details: errMsg });
          return;
        }

        console.warn("sendPortalMessage: runtime error", { tabId, type: message?.type, errMsg });
        resolve({ ok: false, error: "runtime_error", details: errMsg });
        return;
      }

      // Normalise: ensure response always has an `ok` field
      if (response && typeof response === "object" && !Object.prototype.hasOwnProperty.call(response, "ok")) {
        resolve({ ok: true, ...response });
        return;
      }

      resolve(response ?? { ok: false, error: "no_response" });
    });
  });
}

async function ensurePortalBridgeInjected(tabId) {
  const files = ["src/content/content.js", "src/content/page-fetcher.js"];

  try {
    await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      files
    });

    for (let attempt = 1; attempt <= BRIDGE_CHECK_RETRIES; attempt++) {
      try {
        const results = await chrome.scripting.executeScript({
          target: { tabId },
          func: () => {
            try {
              return (
                Boolean(window.__EXT_PAGE_FETCH_INJECTED) &&
                Array.isArray(window.__HOADON_LOGIN_MODAL_SELECTORS)
              );
            } catch (e) { console.error(e); return false; }
          }
        });

        if (Array.isArray(results) && results.some((r) => r?.result === true)) {
          return; // Bridge confirmed
        }
      } catch (e) {
        console.error(e);
        // ignore, retry
      }

      await delay(BRIDGE_CHECK_DELAY_MS * attempt);
    }

    console.warn("ensurePortalBridgeInjected: bridge not detected after retries", { tabId, files });
  } catch (err) {
    const msg = String(err?.message || err || "");
    if (!isMissingReceiverError(msg)) {
      throw err;
    }
    console.warn("ensurePortalBridgeInjected: missing receiver", { tabId, message: msg });
  }
}

async function getPortalRequestContext() {
  const tab = await findPortalTab();
  if (!tab?.id) return null;

  const message = { type: "GET_PORTAL_REQUEST_CONTEXT" };
  let response = await sendPortalMessage(tab.id, message);

  if (response?.ok) return response.data || null;

  await ensurePortalBridgeInjected(tab.id);
  response = await sendPortalMessage(tab.id, message);

  return response?.ok ? response.data || null : null;
}

async function findPortalTab() {
  const portalFlow = await readPortalFlow();

  if (portalFlow?.tabId != null) {
    try {
      const tab = await chrome.tabs.get(portalFlow.tabId);
      if (tab?.id && tab.url?.includes("hoadondientu.gdt.gov.vn")) {
        return tab;
      }
    } catch (e) {
      console.error(e);
      // tab no longer exists, fall through
    }
  }

  const tabs = await chrome.tabs.query({});
  return (
    tabs.find((t) => t.active && t.url?.includes("hoadondientu.gdt.gov.vn")) ||
    tabs.find((t) => t.url?.includes("hoadondientu.gdt.gov.vn")) ||
    null
  );
}

// =============================================================================
// INVOICE FETCHING
// =============================================================================

async function fetchWithFreshPortalHeaders(url, { source, path, actionOverride = null, retries = 1, timeout = 15000 } = {}) {
  return withPortalRequestDelay(async () => {
    const headers = await buildPortalRequestHeaders(source, path, actionOverride);
    return fetchViaPortalPage(url, { retries, timeout, headers });
  });
}

async function fetchInvoices(payload) {
  if (!payload?.dateFrom || !payload?.dateTo) {
    throw new Error("Thiếu tham số dateFrom/dateTo");
  }

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

  // Both sources failed — attempt local cache fallback
  if (!queryInvoices.length && !scoInvoices.length) {
    const isNetworkFailure =
      isLikelyNetworkOrTimeoutFailure(queryResult.reason) ||
      isLikelyNetworkOrTimeoutFailure(scoResult.reason);

    if (isNetworkFailure) {
      const cached = await getCachedInvoices({ dateFrom: payload.dateFrom, dateTo: payload.dateTo });
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

    throw new Error(
      errors.length
        ? `Không tải được dữ liệu: ${errors.join(" | ")}`
        : "Không tải được dữ liệu hóa đơn"
    );
  }

  const merged = dedupeInvoices([...queryInvoices, ...scoInvoices]);
  const syncResult = await syncInvoicesToLocal(merged, {
    requestedDateFrom: payload.dateFrom,
    requestedDateTo: payload.dateTo,
    queryCount: queryInvoices.length,
    scoCount: scoInvoices.length
  });

  const { appState: existing = {} } = await chrome.storage.local.get("appState");
  const now = new Date().toISOString();

  await chrome.storage.local.set({
    appState: {
      ...existing,
      lastSyncAt: now,
      lastDateFrom: payload.dateFrom,
      lastDateTo: payload.dateTo,
      invoiceCount: syncResult.totalCount,
      dataOrigin: "network",
      productCatalog: existing.productCatalog || {
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
    lastSyncAt: now
  };
}

async function fetchInvoiceSource(path, dateFrom, dateTo, status, size, source) {
  const pageSize = clampPageSize(size);
  let page = 0;
  let state = null;
  const seenStates = new Set();
  const all = [];

  while (true) {
    const url = buildUrl(path, { dateFrom, dateTo, status, size: pageSize, page, state });

    const response = await fetchWithFreshPortalHeaders(url, {
      source,
      path,
      retries: 1,
      timeout: PORTAL_PAGE_FETCH_TIMEOUT_MS
    });

    if (isAuthError(response?.status)) {
      throw new Error("Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.");
    }

    if (!response?.ok) {
      throw new Error(formatPortalFetchFailureMessage(source, response));
    }

    const json = safeParseJson(response.body, source);
    const invoices = Array.isArray(json?.datas) ? json.datas : [];

    if (invoices.length) {
      for (const item of invoices) all.push({ ...item, _source: source });
    }

    if (Object.prototype.hasOwnProperty.call(json, "state") && json.state === null) {
      break;
    }

    const nextState = typeof json?.state === "string" && json.state ? String(json.state) : null;
    if (nextState) {
      if (seenStates.has(nextState)) break;
      seenStates.add(nextState);
      state = nextState;
      await delay(PORTAL_REQUEST_DELAY_MS);
      continue;
    }

    break;
  }

  return all;
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

  const response = await fetchWithFreshPortalHeaders(url, {
    source,
    path: detailPath,
    actionOverride: "Xem hóa đơn (hóa đơn máy tính tiền mua vào)",
    retries: 1,
    timeout: PORTAL_PAGE_FETCH_TIMEOUT_MS
  });

  if (isAuthError(response?.status)) {
    throw new Error("Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.");
  }

  if (!response?.ok) {
    throw new Error(formatPortalFetchFailureMessage("query", response, "Tải chi tiết hóa đơn thất bại"));
  }

  return safeParseJson(response.body, "detail");
}

// =============================================================================
// INDEXEDDB — SYNC & QUERY
// =============================================================================

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
    request.onerror = () =>
      reject(request.error || new Error("Không mở được IndexedDB hóa đơn"));
  });
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

    tx.onerror = () => reject(tx.error || new Error("Không ghi được hóa đơn vào IndexedDB"));
    tx.onabort = () => reject(tx.error || new Error("Giao dịch IndexedDB bị hủy"));
    tx.oncomplete = () => resolve();

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
        metaReq.onerror = () =>
          reject(metaReq.error || new Error("Không ghi được metadata đồng bộ"));
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
          inserted++;
        } else if (shouldReplaceStoredInvoice(existing, nextRecord)) {
          store.put(nextRecord);
          updated++;
        } else {
          skipped++;
        }
        processNext();
      };
    };

    processNext();
  });

  return { invoices: ordered, totalCount: ordered.length, inserted, updated, skipped };
}

async function queryInvoicesByDateRange(dateFrom, dateTo, limit) {
  const db = await openInvoiceDatabase();
  const lowerBound = parseApiDateToTimestamp(dateFrom, false);
  const upperBound = parseApiDateToTimestamp(dateTo, true);
  const results = [];

  await new Promise((resolve, reject) => {
    const tx = db.transaction(INVOICE_STORE_NAME, "readonly");
    const store = tx.objectStore(INVOICE_STORE_NAME);
    const idx = store.index("issuedAtMs");
    const range = IDBKeyRange.bound(lowerBound, upperBound);
    const cursorReq = idx.openCursor(range, "prev");

    cursorReq.onerror = () =>
      reject(cursorReq.error || new Error("Không đọc được hóa đơn local"));
    cursorReq.onsuccess = () => {
      const cursor = cursorReq.result;
      if (!cursor) { resolve(); return; }

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
    req.onsuccess = () => resolve(req.result?.invoice || null);
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
  const [invoices, syncState] = await Promise.all([
    queryInvoicesByDateRange(dateFrom, dateTo, limit),
    getInvoiceSyncState()
  ]);
  return { invoices, syncState };
}

// =============================================================================
// INVOICE RECORD HELPERS
// =============================================================================

function buildInvoiceRecord(invoice, syncedAt) {
  return {
    invoiceKey: buildInvoiceKey(invoice),
    invoice: { ...invoice, _source: invoice?._source || "network" },
    issuedAtMs: parseInvoiceTimestamp(invoice?.tdlap || invoice?.ntao || invoice?.ncnhat),
    updatedAtMs: parseInvoiceTimestamp(
      invoice?.ncnhat || invoice?.ncnhattt || invoice?.ncnhatdl || invoice?.ncnhatts || invoice?.updatedAt || invoice?.modifiedAt
    ),
    contentSignature: buildInvoiceContentSignature(invoice),
    syncedAt,
    source: invoice?._source || "network"
  };
}

function shouldReplaceStoredInvoice(existingRecord, nextRecord) {
  if (!existingRecord?.invoice) return true;

  if (
    existingRecord.contentSignature &&
    nextRecord.contentSignature &&
    existingRecord.contentSignature === nextRecord.contentSignature
  ) {
    return false;
  }

  const existingScore = scoreInvoice(existingRecord.invoice);
  const nextScore = scoreInvoice(nextRecord.invoice);
  if (nextScore !== existingScore) return nextScore > existingScore;

  const existingUpdatedAtMs = Number(existingRecord.updatedAtMs || 0);
  const nextUpdatedAtMs = Number(nextRecord.updatedAtMs || 0);
  if (nextUpdatedAtMs !== existingUpdatedAtMs) return nextUpdatedAtMs > existingUpdatedAtMs;

  const existingIssuedAtMs = Number(existingRecord.issuedAtMs || 0);
  const nextIssuedAtMs = Number(nextRecord.issuedAtMs || 0);
  if (nextIssuedAtMs !== existingIssuedAtMs) return nextIssuedAtMs > existingIssuedAtMs;

  const existingSyncedAt = parseInvoiceTimestamp(existingRecord.syncedAt);
  const nextSyncedAt = parseInvoiceTimestamp(nextRecord.syncedAt);
  return nextSyncedAt > existingSyncedAt;
}

function buildInvoiceContentSignature(invoice) {
  const parts = [
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

  return parts.map(normalizeInvoiceSignatureValue).join("~");
}

function normalizeInvoiceSignatureValue(value) {
  if (value == null) return "";
  if (Array.isArray(value)) return value.map(normalizeInvoiceSignatureValue).join(",");
  if (typeof value === "object") return JSON.stringify(value);
  return String(value).trim();
}

function buildInvoiceKey(invoice) {
  if (invoice?.mhdon) return `mhdon:${invoice.mhdon}`;
  if (invoice?.id) return `id:${invoice.id}`;
  return [invoice?.khhdon ?? "", invoice?.shdon ?? "", invoice?.nbmst ?? "", invoice?.tdlap ?? ""].join("|");
}

function scoreInvoice(invoice) {
  let score = 0;
  if (invoice?.nky) score += 2;
  if (invoice?.nbcks) score += 2;
  if (invoice?.thttltsuat?.length) score += 1;
  if (invoice?._source === "sco-query") score += 1;
  return score;
}

function dedupeInvoices(invoices) {
  const map = new Map();

  for (const invoice of invoices) {
    const key = buildInvoiceKey(invoice);
    if (!map.has(key)) {
      map.set(key, invoice);
      continue;
    }
    if (scoreInvoice(invoice) > scoreInvoice(map.get(key))) {
      map.set(key, invoice);
    }
  }

  return Array.from(map.values()).sort((a, b) => {
    const da = new Date(a?.tdlap || 0).getTime();
    const db = new Date(b?.tdlap || 0).getTime();
    return db - da;
  });
}

// =============================================================================
// DATE / URL UTILITIES
// =============================================================================

function buildUrl(path, { dateFrom, dateTo, status, size, page, state } = {}) {
  const search = `tdlap=ge=${dateFrom};tdlap=le=${dateTo};ttxly==${status}`;
  const qs = new URLSearchParams();
  qs.set("sort", "tdlap:desc");
  qs.set("size", String(clampPageSize(size)));
  qs.set("search", search);
  if (state) qs.set("state", state);
  else if (Number.isFinite(Number(page))) qs.set("page", String(Number(page)));
  return `${QUERY_BASE}${path}?${qs.toString()}`;
}

function buildDetailUrl({ nbmst, khhdon, shdon, khmshdon, detailPath }) {
  const params = new URLSearchParams({ nbmst, khhdon, shdon, khmshdon });
  return `${QUERY_BASE}${detailPath}?${params.toString()}`;
}

function clampPageSize(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0
    ? Math.min(Math.trunc(parsed), MAX_PURCHASE_PAGE_SIZE)
    : MAX_PURCHASE_PAGE_SIZE;
}

function formatDateForApi(value, endOfDay = false) {
  const date = new Date(value);
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = String(date.getFullYear());
  return `${dd}/${mm}/${yyyy}T${endOfDay ? "23:59:59" : "00:00:00"}`;
}

function parseApiDateToTimestamp(value, endOfDay = false) {
  const text = String(value || "").trim();
  const match = text.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (!match) return 0;
  const [, dd, mm, yyyy] = match;
  return new Date(`${yyyy}-${mm}-${dd}T${endOfDay ? "23:59:59" : "00:00:00"}`).getTime();
}

function parseInvoiceTimestamp(value) {
  if (!value) return 0;
  const t = new Date(value).getTime();
  return Number.isFinite(t) ? t : 0;
}

function safeParseJson(text, source) {
  try {
    return JSON.parse(text);
  } catch (e) {
    console.error('safeParseJson failed:', e);
    throw new Error(`Phản hồi ${source} không phải JSON hợp lệ`);
  }
}

function isLikelyNetworkOrTimeoutFailure(error) {
  return isNetworkError(error) || isTimeoutError(error);
}

// =============================================================================
// TAB / WINDOW HELPERS
// =============================================================================

async function openLoginWindow() {
  try {
    await chrome.tabs.create({ url: PORTAL_SEARCH_URL, active: true });
  } catch (e) {
    try {
      await chrome.windows.create({
        url: PORTAL_SEARCH_URL,
        type: "popup",
        width: 1180,
        height: 900,
        focused: true
      });
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
  const existing = tabs.find((t) => t.url === targetUrl);

  if (existing?.id !== undefined) {
    await chrome.tabs.update(existing.id, { active: true });
    if (existing.windowId !== chrome.windows.WINDOW_ID_CURRENT) {
      await chrome.windows.update(existing.windowId, { focused: true });
    }
    return;
  }

  await chrome.tabs.create({ url: targetUrl, active: true });
}

async function openExtensionWindow(windowBounds = {}) {
  const targetUrl = chrome.runtime.getURL("src/popup/popup.html");
  const width = Number.isFinite(Number(windowBounds.width)) ? Math.max(320, Math.round(Number(windowBounds.width))) : 1180;
  const height = Number.isFinite(Number(windowBounds.height)) ? Math.max(240, Math.round(Number(windowBounds.height))) : 900;
  const left = Number.isFinite(Number(windowBounds.left)) ? Math.round(Number(windowBounds.left)) : undefined;
  const top = Number.isFinite(Number(windowBounds.top)) ? Math.round(Number(windowBounds.top)) : undefined;

  await chrome.windows.create({
    url: targetUrl,
    type: "popup",
    width,
    height,
    ...(left !== undefined ? { left } : {}),
    ...(top !== undefined ? { top } : {}),
    focused: true
  });
}

// =============================================================================
// LIFECYCLE & MESSAGE HANDLER
// =============================================================================

chrome.runtime.onInstalled.addListener((details) => {
  if (details?.reason !== "install") return;

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
  openExtensionTab().catch((err) => {
    console.error("Không mở được tab extension:", err);
  });
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message?.type) return false;

  const tabId = sender?.tab?.id ?? null;

  switch (message.type) {
    case "CHECK_AUTH":
      resolveAuthState()
        .then((result) => sendResponse({ ok: true, ...result }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "PORTAL_FLOW":
    case "AUTH_HINT": {
      const portalFlow = normalizePortalFlow(message.payload || {}, tabId);
      chrome.storage.local
        .set({ portalFlow, authHint: portalFlow })
        .then(() => sendResponse({ ok: true, data: portalFlow }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;
    }

    case "GET_PORTAL_FLOW":
      readPortalFlow()
        .then((portalFlow) => sendResponse({ ok: true, data: portalFlow }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "OPEN_LOGIN_POPUP":
      openLoginWindow()
        .then(() => sendResponse({ ok: true }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "OPEN_EXTENSION_WINDOW":
      openExtensionWindow(message.payload || {})
        .then(() => sendResponse({ ok: true }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "OPEN_EXTERNAL_URL": {
      const url = message.payload?.url;
      if (!url) {
        sendResponse({ ok: false, error: "Missing url payload" });
        return true;
      }
      openExternalUrl(url)
        .then(() => sendResponse({ ok: true }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;
    }

    case "FETCH_INVOICES":
      fetchInvoices(message.payload)
        .then((result) => sendResponse({ ok: true, data: result }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "GET_SYNC_STATE":
      getInvoiceSyncState()
        .then((result) => sendResponse({ ok: true, data: result }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "GET_CACHED_INVOICES":
      getCachedInvoices(message.payload || {})
        .then((result) => sendResponse({ ok: true, data: result }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "GET_CACHED_INVOICE": {
      const invoice = message.payload?.invoice;
      if (!invoice) {
        sendResponse({ ok: false, error: "Missing invoice payload" });
        return true;
      }
      getCachedInvoiceByKey(invoice)
        .then((result) => sendResponse({ ok: true, data: result }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;
    }

    case "FETCH_INVOICE_DETAIL":
      fetchInvoiceDetail(message.payload)
        .then((result) => sendResponse({ ok: true, data: result }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;

    case "UPSERT_INVOICE": {
      const invoice = message.payload?.invoice;
      if (!invoice) {
        sendResponse({ ok: false, error: "Missing invoice payload" });
        return true;
      }
      syncInvoicesToLocal([invoice], {})
        .then((res) => sendResponse({ ok: true, data: res }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;
    }

    default:
      return false;
  }
});