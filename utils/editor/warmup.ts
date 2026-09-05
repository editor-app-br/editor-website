import { WARMUP_CACHE_PREFIX } from "@/utils/embed-protocol";
import { API_JS, PRELOAD_HTML } from "@/utils/editor/utils";

const X2T_ASSETS = ["/x2t/x2t.js", "/x2t/x2t.wasm"];

const EDITOR_SHELLS = [
  "/web-apps/apps/documenteditor/main/index.html",
  "/web-apps/apps/spreadsheeteditor/main/index.html",
  "/web-apps/apps/presentationeditor/main/index.html",
  "/web-apps/apps/pdfeditor/main/index.html",
];

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

function isVersionedStatic(url: string): boolean {
  try {
    const path = new URL(url, location.origin).pathname;
    if (path.endsWith(".html")) return false;
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
  if (contentType.includes("text/html") && !requestUrl.includes("/office-plugins/")) {
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
  const cacheName = `${WARMUP_CACHE_PREFIX}${appRoot.replace(/^\//, "")}`;
  const cache = await caches.open(cacheName);
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

export function shouldPreserveEditorCache(name: string): boolean {
  return name.startsWith(WARMUP_CACHE_PREFIX) || name.startsWith("document_editor_static_");
}

export function warmupSeedCount(appRoot: string): number {
  return seedUrls(appRoot).length;
}

export { isVersionedStatic };
