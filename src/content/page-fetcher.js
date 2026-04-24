if (!window.__EXT_PAGE_FETCH_INJECTED) {
  window.__EXT_PAGE_FETCH_INJECTED = true;

  window.addEventListener(
    "message",
    async function (event) {
      if (
        event.source !== window ||
        !event.data ||
        event.data.source !== "EXT_PAGE_FETCH_REQUEST"
      ) {
        return;
      }

      const { id, url, method = "GET", headers = {} } = event.data;
      try {
        const res = await fetch(url, {
          method,
          credentials: "include",
          mode: "cors",
          cache: "default",
          redirect: "follow",
          referrer: "https://hoadondientu.gdt.gov.vn/",
          referrerPolicy: "strict-origin-when-cross-origin",
          headers: {
            Accept: "application/json, text/plain, */*",
            "Accept-Language": "vi",
            ...headers
          }
        });
        const text = await res.text();
        window.postMessage(
          { source: "EXT_PAGE_FETCH", id, status: res.status, body: text },
          "*"
        );
      } catch (e) {
        window.postMessage(
          {
            source: "EXT_PAGE_FETCH",
            id,
            error: e && e.message ? e.message : String(e)
          },
          "*"
        );
      }
    },
    false
  );
}
