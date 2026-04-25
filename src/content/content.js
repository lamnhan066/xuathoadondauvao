if (window.__HOADON_CONTENT_INJECTED) {
  // already injected by another mechanism (manifest or scripting.executeScript)
} else {
  window.__HOADON_CONTENT_INJECTED = true;

  let publishTimer = null;
  let lastSignature = "";

  // Use a single fetch probe to determine authentication state.
  async function buildPortalFlow() {
    const storageSnapshot = collectStorageSnapshot();
    const authorization =
      findBearerValue(storageSnapshot) ||
      buildBearerFromValue(storageSnapshot.jwt) ||
      buildBearerFromValue(storageSnapshot.token) ||
      buildBearerFromValue(storageSnapshot.accessToken) ||
      buildBearerFromValue(findJwtCookieValue());

    const payload = {
      phase: authorization ? "authenticated-shell" : "unknown",
      hasLoginModal: false,
      hasInvoiceSearch: false,
      hasLoggedInActionButton: false,
      isLoggedInUi: Boolean(authorization),
      title: document.title || "",
      url: window.location.href,
      evidence: [],
      updatedAt: new Date().toISOString()
    };

    if (authorization) {
      payload.evidence.push({ type: "auth", source: "storage-or-cookie" });
    } else {
      payload.evidence.push({ type: "auth", source: "missing-token" });
    }

    return payload;
  }

  async function publishPortalFlow() {
    try {
      const payload = await buildPortalFlow();
      const signature = JSON.stringify(payload);
      if (signature === lastSignature) return;
      lastSignature = signature;

      try {
        if (chrome && chrome.runtime && chrome.runtime.sendMessage) {
          chrome.runtime.sendMessage({ type: "PORTAL_FLOW", payload }, (response) => {
            if (chrome.runtime.lastError) {
              console.warn("PORTAL_FLOW sendMessage error:", chrome.runtime.lastError);
            }
          });
        }
      } catch (e) {
        console.warn("PORTAL_FLOW sendMessage exception:", e);
      }
    } catch (e) {
      console.warn("publishPortalFlow error:", e);
    }
  }

  function readPortalRequestContext() {
    const storageSnapshot = collectStorageSnapshot();
    const authorization =
      findBearerValue(storageSnapshot) ||
      buildBearerFromValue(storageSnapshot.jwt) ||
      buildBearerFromValue(storageSnapshot.token) ||
      buildBearerFromValue(storageSnapshot.accessToken) ||
      buildBearerFromValue(findJwtCookieValue());

    return {
      authorization,
      acceptLanguage: storageSnapshot.language || "vi",
      actionQuery: "Tìm kiếm (hóa đơn mua vào)",
      actionScoQuery: "Tìm kiếm (hóa đơn máy tính tiền mua vào)",
      endPoint: "/tra-cuu/tra-cuu-hoa-don",
      cookieJwt: findJwtCookieValue() || null,
      source: authorization ? "page-storage" : "cookie"
    };
  }

  function collectStorageSnapshot() {
    const snapshot = {};

    try {
      for (let index = 0; index < localStorage.length; index += 1) {
        const key = localStorage.key(index);
        if (key) {
          snapshot[key] = localStorage.getItem(key);
        }
      }
    } catch (error) {
      console.error('collectStorageSnapshot localStorage error:', error);
    }

    try {
      for (let index = 0; index < sessionStorage.length; index += 1) {
        const key = sessionStorage.key(index);
        if (key) {
          snapshot[key] = sessionStorage.getItem(key);
        }
      }
    } catch (error) {
      console.error('collectStorageSnapshot sessionStorage error:', error);
    }

    return snapshot;
  }

  function findBearerValue(snapshot) {
    for (const value of Object.values(snapshot || {})) {
      if (typeof value !== "string") continue;

      const bearerMatch = value.match(/Bearer\s+([A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+)/i);
      if (bearerMatch) {
        return `Bearer ${bearerMatch[1]}`;
      }
    }

    return "";
  }

  function buildBearerFromValue(value) {
    if (typeof value !== "string") return "";
    const trimmed = value.trim();
    if (!trimmed) return "";
    if (/^Bearer\s+/i.test(trimmed)) return trimmed;
    if (/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(trimmed)) {
      return `Bearer ${trimmed}`;
    }
    return "";
  }

  function findJwtCookieValue() {
    const cookies = String(document.cookie || "").split(";");
    for (const cookie of cookies) {
      const [rawKey, ...rest] = cookie.split("=");
      const key = rawKey.trim();
      if (key.toLowerCase() === "jwt") {
        return rest.join("=").trim();
      }
    }

    return "";
  }

  function schedulePublishPortalFlow() {
    if (publishTimer) {
      clearTimeout(publishTimer);
    }

    publishTimer = setTimeout(() => {
      publishTimer = null;
      publishPortalFlow();
    }, 200);
  }

  function isElementVisible(element) {
    if (!element) return false;
    const style = window.getComputedStyle(element);
    return style.display !== "none" && style.visibility !== "hidden" && style.opacity !== "0";
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

  const observer = new MutationObserver(() => {
    schedulePublishPortalFlow();
  });

  observer.observe(document.documentElement, {
    childList: true,
    subtree: true,
    attributes: true,
    characterData: true
  });

  try {
    publishPortalFlow();
  } catch (e) {
    console.warn("publishPortalFlow initial invocation failed:", e);
  }

  // Listen for page fetch requests from the extension and perform them in page context
  try {
    if (chrome && chrome.runtime && chrome.runtime.onMessage && chrome.runtime.onMessage.addListener) {
      chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
        try {
          if (!message) return false;

          if (message.type === "GET_PORTAL_REQUEST_CONTEXT") {
            sendResponse({ ok: true, data: readPortalRequestContext() });
            return false;
          }

          if (!message || message.type !== "PAGE_FETCH") return false;

          const url = message.url;
          const method = message.method || "GET";
          const requestHeaders = {
            Accept: "application/json, text/plain, */*",
            "Accept-Language": "vi",
            ...(message.headers || {})
          };

          const id = `ext_page_fetch_${Date.now()}_${Math.random().toString(36).slice(2)}`;

          function handleMessage(event) {
            if (event.source !== window || !event.data || event.data.source !== "EXT_PAGE_FETCH" || event.data.id !== id) return;
            window.removeEventListener("message", handleMessage);
            const data = event.data;
            sendResponse({ ok: !data.error && data.status >= 200 && data.status < 400, status: data.status, body: data.body, error: data.error });
          }

          window.addEventListener("message", handleMessage);

          // Perform the page fetch with retries and per-attempt timeout
          const runPageFetch = async () => {
            const maxRetries = message.retries ?? 3;
            const attemptTimeout = message.timeout ?? 1500; // ms per attempt

            try {
              const result = await doFetchAttempt(id, url, method, requestHeaders, maxRetries, attemptTimeout);
              sendResponse({ ok: !result.error && result.status >= 200 && result.status < 400, status: result.status, body: result.body, error: result.error });
            } catch (e) {
              sendResponse({ ok: false, error: e?.message || String(e) });
            }
          };

          runPageFetch();

          return true; // async
        } catch (e) {
          console.warn("onMessage handler error:", e);
          return false;
        }
      });
    }
  } catch (e) {
    console.warn("Failed to register chrome.runtime.onMessage listener:", e);
  }

  async function doFetchAttempt(id, url, method, requestHeaders, maxRetries, attemptTimeout) {
    return new Promise((resolve) => {
      let attempts = 0;
      let attemptTimer = null;
      let overallTimer = null;

      function cleanup() {
        window.removeEventListener("message", onMessage);
        if (attemptTimer) clearTimeout(attemptTimer);
        if (overallTimer) clearTimeout(overallTimer);
      }

      function onMessage(event) {
        if (event.source !== window || !event.data || event.data.source !== "EXT_PAGE_FETCH" || event.data.id !== id) return;
        cleanup();
        const data = event.data;
        if (data.error) {
          resolve({ error: data.error });
        } else {
          resolve({ status: data.status, body: data.body });
        }
      }

      window.addEventListener("message", onMessage);

      function attempt() {
        attempts += 1;
        window.postMessage({ source: "EXT_PAGE_FETCH_REQUEST", id, url, method, headers: requestHeaders }, "*");
        attemptTimer = setTimeout(() => {
          if (attempts < maxRetries) {
            attempt();
          } else {
            cleanup();
            resolve({ error: "timeout" });
          }
        }, attemptTimeout);
      }

      // Overall safety timeout (a bit longer than attempts total)
      overallTimer = setTimeout(() => {
        cleanup();
        resolve({ error: "timeout" });
      }, maxRetries * attemptTimeout + 500);

      attempt();
    });
  }
}
