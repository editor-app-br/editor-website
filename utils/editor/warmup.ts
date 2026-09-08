import { WARMUP_CACHE_PREFIX } from "@/utils/embed-protocol";
import { API_JS, getAppRoot, PRELOAD_HTML } from "@/utils/editor/utils";

const X2T_ASSETS = ["/x2t/x2t.js", "/x2t/x2t.wasm"];

const EDITOR_SHELLS = [
  "/web-apps/apps/documenteditor/main/index.html",
  "/web-apps/apps/spreadsheeteditor/main/index.html",
  "/web-apps/apps/presentationeditor/main/index.html",
  "/web-apps/apps/pdfeditor/main/index.html",
];

/** Must be present before DocsAPI can start reliably after a cold cache wipe. */
const CRITICAL_RELATIVE = [API_JS, "/x2t/x2t.js", "/x2t/x2t.wasm"] as const;

const FETCH_HOOK = "__jiEditorAssetFetch";

function seedUrls(appRoot: string): string[] {
  const root = appRoot.replace(/\/$/, "");
  return [
    `${root}${API_JS}`,
    `${root}${PRELOAD_HTML}`,
    `${root}/sdkjs-plugins/v1/plugins.js`,
    ...EDITOR_SHELLS.map((path) => `${root}${path}`),
    ...X2T_ASSETS,
    "/office-plugins/agent/config.json",
    "/office-plugins/agent/index.html",
    "/office-plugins/agent/plugin.js",
  ];
}

export function warmupCacheName(appRoot: string): string {
  return `${WARMUP_CACHE_PREFIX}${appRoot.replace(/^\//, "")}`;
}

function cacheNameForUrl(url: string): string {
  try {
    const path = new URL(url, location.origin).pathname;
    const match = path.match(/^\/(v[^/]+)\//);
    if (match) return warmupCacheName(`/${match[1]}`);
  } catch {
    /* fall through */
  }
  return warmupCacheName(getAppRoot());
}

function isVersionedStatic(url: string): boolean {
  try {
    const path = new URL(url, location.origin).pathname;
    if (path.endsWith("plugins.json")) return false;
    if (path.endsWith(".html")) {
      // Editor shells under /v* are safe to keep offline; skip generic HTML.
      return /^\/v[^/]+\//.test(path) || path.startsWith("/office-plugins/");
    }
    return (
      /^\/v[^/]+\//.test(path) ||
      path.startsWith("/x2t/") ||
      path.startsWith("/office-plugins/")
    );
  } catch {
    return false;
  }
}

function allowCachedHtml(path: string): boolean {
  return (
    path.startsWith("/office-plugins/") ||
    (/^\/v[^/]+\//.test(path) && path.endsWith(".html"))
  );
}

async function putIfOk(cache: Cache, requestUrl: string): Promise<boolean> {
  // Bypass HTTP cache so a wiped Cache Storage entry is re-downloaded.
  const response = await fetch(requestUrl, {
    credentials: "same-origin",
    cache: "reload",
  });
  if (!response.ok) return false;
  const contentType = response.headers.get("content-type") || "";
  const path = new URL(requestUrl, location.origin).pathname;
  if (contentType.includes("text/html") && !allowCachedHtml(path)) {
    return false;
  }
  await cache.put(requestUrl, response.clone());
  return true;
}

/** Cache Storage hit for a versioned editor asset, or null. */
export async function matchEditorAsset(url: string): Promise<Response | null> {
  if (typeof caches === "undefined" || !isVersionedStatic(url)) return null;
  try {
    const href = new URL(url, location.origin).href;
    const hit = await caches.match(href);
    return hit?.ok ? hit : null;
  } catch {
    return null;
  }
}

/**
 * Serve from Cache Storage, else download from this origin and recache.
 * Used when an asset was deleted from Cache Storage.
 */
export async function fetchThroughEditorCache(
  input: RequestInfo | URL,
  init?: RequestInit,
  nativeFetch: typeof fetch = fetch,
): Promise<Response> {
  const request = input instanceof Request && !init ? input : new Request(input, init);
  if (request.method !== "GET" || !isVersionedStatic(request.url)) {
    return nativeFetch(request);
  }
  const href = new URL(request.url, location.origin).href;
  const cached = await matchEditorAsset(href);
  if (cached) return cached;
  try {
    const response = await nativeFetch(
      new Request(href, {
        method: "GET",
        credentials: request.credentials,
        cache: "reload",
      }),
    );
    if (response.ok && typeof caches !== "undefined") {
      const path = new URL(href).pathname;
      const contentType = response.headers.get("content-type") || "";
      if (!contentType.includes("text/html") || allowCachedHtml(path)) {
        const cache = await caches.open(cacheNameForUrl(href));
        void cache.put(href, response.clone()).catch(() => undefined);
      }
    }
    return response;
  } catch (err) {
    const fallback = await matchEditorAsset(href);
    if (fallback) return fallback;
    throw err;
  }
}

/** Patch window.fetch so x2t WASM / same-origin assets refill Cache Storage on miss. */
export function installEditorAssetFetchHook(target: Window = window): void {
  const win = target as Window & { [FETCH_HOOK]?: boolean; fetch: typeof fetch };
  if (win[FETCH_HOOK] || typeof win.fetch !== "function") return;
  const native = win.fetch.bind(win);
  win.fetch = ((input: RequestInfo | URL, init?: RequestInit) =>
    fetchThroughEditorCache(input, init, native)) as typeof fetch;
  win[FETCH_HOOK] = true;
}

export async function prefetchEditorAssets(
  appRoot: string,
  onProgress?: (done: number, total: number) => void,
): Promise<{ cached: number; failed: number }> {
  if (typeof caches === "undefined") {
    return { cached: 0, failed: seedUrls(appRoot).length };
  }
  const cache = await caches.open(warmupCacheName(appRoot));
  const urls = seedUrls(appRoot);
  let cached = 0;
  let failed = 0;
  const batchSize = 4;
  for (let i = 0; i < urls.length; i += batchSize) {
    const slice = urls.slice(i, i + batchSize);
    const results = await Promise.all(
      slice.map(async (path) => {
        try {
          const href = new URL(path, location.origin).href;
          const existing = await caches.match(href);
          if (existing?.ok) return true;
          return await putIfOk(cache, href);
        } catch {
          return false;
        }
      }),
    );
    for (const ok of results) {
      if (ok) cached += 1;
      else failed += 1;
    }
    onProgress?.(Math.min(i + slice.length, urls.length), urls.length);
  }
  return { cached, failed };
}

export async function isEditorCacheReady(appRoot: string): Promise<boolean> {
  if (typeof caches === "undefined") return false;
  const root = appRoot.replace(/\/$/, "");
  try {
    for (const rel of CRITICAL_RELATIVE) {
      const href = new URL(
        rel.startsWith("/x2t/") ? rel : `${root}${rel}`,
        location.origin,
      ).href;
      const hit = await caches.match(href);
      if (!hit || !hit.ok) return false;
    }
    return true;
  } catch {
    return false;
  }
}

/**
 * If the static cache was wiped, refill it. Safe to call on every /embed open.
 * Returns whether critical assets are ready after the call.
 */
export async function ensureEditorAssets(
  appRoot: string,
  onProgress?: (done: number, total: number) => void,
): Promise<{ ready: boolean; cached: number; failed: number; skipped: boolean }> {
  if (await isEditorCacheReady(appRoot)) {
    return { ready: true, cached: 0, failed: 0, skipped: true };
  }
  const result = await prefetchEditorAssets(appRoot, onProgress);
  const ready = await isEditorCacheReady(appRoot);
  return { ready, cached: result.cached, failed: result.failed, skipped: false };
}

/**
 * Load DocsAPI with a real script URL. Blob URLs break OnlyOffice getBasePath()
 * (`script.src` must match `api/documents/api.js`).
 */
export async function loadScriptCached(url: string): Promise<void> {
  if ((window as Window & { DocsAPI?: { DocEditor?: unknown } }).DocsAPI?.DocEditor) {
    return;
  }
  await new Promise<void>((resolve, reject) => {
    const script = document.createElement("script");
    script.src = url;
    script.dataset.jiSrc = url;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load script ${url}`));
    document.head.appendChild(script);
  });
}

export function shouldPreserveEditorCache(name: string): boolean {
  return name.startsWith(WARMUP_CACHE_PREFIX) || name.startsWith("document_editor_static_");
}

export function warmupSeedCount(appRoot: string): number {
  return seedUrls(appRoot).length;
}

export { isVersionedStatic };
