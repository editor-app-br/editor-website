import { WARMUP_CACHE_PREFIX } from "@/utils/embed-protocol";
import { API_JS, PRELOAD_HTML } from "@/utils/editor/utils";

const X2T_ASSETS = ["/x2t/x2t.js", "/x2t/x2t.wasm"];

const EDITOR_SHELLS = [
  "/web-apps/apps/documenteditor/main/index.html",
  "/web-apps/apps/spreadsheeteditor/main/index.html",
  "/web-apps/apps/presentationeditor/main/index.html",
  "/web-apps/apps/pdfeditor/main/index.html",
];

/** Must be present before DocsAPI can start reliably after a cold cache wipe. */
const CRITICAL_RELATIVE = [API_JS, "/x2t/x2t.js", "/x2t/x2t.wasm"] as const;

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

function isVersionedStatic(url: string): boolean {
  try {
    const path = new URL(url, location.origin).pathname;
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

async function putIfOk(cache: Cache, requestUrl: string): Promise<boolean> {
  const response = await fetch(requestUrl, { credentials: "same-origin" });
  if (!response.ok) return false;
  const contentType = response.headers.get("content-type") || "";
  const path = new URL(requestUrl, location.origin).pathname;
  const allowHtml =
    path.startsWith("/office-plugins/") ||
    (/^\/v[^/]+\//.test(path) && path.endsWith(".html"));
  if (contentType.includes("text/html") && !allowHtml) {
    return false;
  }
  await cache.put(requestUrl, response.clone());
  return true;
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
          const ok = await putIfOk(cache, new URL(path, location.origin).href);
          return ok;
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
    const cache = await caches.open(warmupCacheName(appRoot));
    for (const rel of CRITICAL_RELATIVE) {
      const href = new URL(rel.startsWith("/x2t/") ? rel : `${root}${rel}`, location.origin).href;
      const hit = await cache.match(href);
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

/** Load DocsAPI (or any script URL) from Cache Storage when present, else network. */
export async function loadScriptCached(url: string): Promise<void> {
  const tryLoad = (src: string, mark: string) =>
    new Promise<void>((resolve, reject) => {
      const script = document.createElement("script");
      script.src = src;
      script.dataset.jiSrc = mark;
      script.onload = () => resolve();
      script.onerror = () => reject(new Error(`Failed to load script ${mark}`));
      document.head.appendChild(script);
    });

  if ((window as Window & { DocsAPI?: { DocEditor?: unknown } }).DocsAPI?.DocEditor) {
    return;
  }

  try {
    if (typeof caches !== "undefined") {
      const hit = await caches.match(url);
      if (hit?.ok) {
        const blob = await hit.blob();
        const objectUrl = URL.createObjectURL(blob);
        try {
          await tryLoad(objectUrl, url);
          URL.revokeObjectURL(objectUrl);
          return;
        } catch {
          URL.revokeObjectURL(objectUrl);
        }
      }
    }
  } catch {
    /* network path below */
  }
  await tryLoad(url, url);
}

export function shouldPreserveEditorCache(name: string): boolean {
  return name.startsWith(WARMUP_CACHE_PREFIX) || name.startsWith("document_editor_static_");
}

export function warmupSeedCount(appRoot: string): number {
  return seedUrls(appRoot).length;
}

export { isVersionedStatic };
