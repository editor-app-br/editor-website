"use client";

import { useLayoutEffect, useRef, useEffect, useState } from "react";
import { useResolvedLanguage, useHasHydrated } from "@/store";
import {
  API_JS,
  getAppRoot,
  getDocumentType,
  PRELOAD_HTML,
  toOnlyOfficeLang,
} from "@/utils/editor/utils";
import io, { MockSocket } from "@/utils/editor/socket";
import { createFetchProxy } from "@/utils/editor/fetch";
import { createXHRProxy } from "@/utils/editor/xhr";
import { DocEditor, PluginMode } from "@/utils/editor/types";
import { EditorServer } from "@/utils/editor/server";
import {
  EMBED_PROTOCOL_VERSION,
  embedMimeForType,
  EmbedPluginDiag,
  EditorToHostMessage,
  HostToEditorMessage,
} from "@/utils/embed-protocol";
import { isAllowedEmbedHostOrigin, loadEmbedPartnerConfig } from "@/utils/embed-origins";
import { injectEmbedChrome, injectPreviewChrome } from "@/utils/embed-chrome";
import { prefetchEditorAssets } from "@/utils/editor/warmup";
import {
  AGENT_PLUGIN_GUID,
  AGENT_PLUGIN_MANIFEST,
  getAgentPluginsData,
  isAgentPluginConfigPath,
  isPluginsJsonPath,
} from "@/utils/editor/plugins";
import {
  LEGAL_CONTACT,
  PUBLIC_ABOUT_URL,
  PUBLIC_CUBE_LOGO,
  PUBLIC_EDITOR_URL,
} from "@/utils/attribution";
import { Loader2 } from "lucide-react";

type PluginHost = {
  onExternalPluginMessage?: (data: Record<string, unknown>) => void;
};

type PluginsController = {
  onPluginsInit?: (plugins: unknown, fromManager?: boolean) => void;
  parsePlugins?: (
    plugins: unknown[],
    uiCustomize?: boolean,
    forceUpdate?: boolean,
    fromManager?: boolean,
  ) => void;
  runAutoStartPlugins?: () => void;
  autostart?: string[];
  api?: { asc_pluginRun?: (guid: string, variation: number, data: string) => void };
  appOptions?: { canPlugins?: boolean; isEdit?: boolean };
};

type AscPluginManager = {
  plugins?: unknown[];
  e2?: Record<string, AscPluginEntry | undefined>;
  GH?: Record<string, { K9?: string } | undefined>;
  SNb?: unknown;
  f9i?: boolean;
  path?: string;
  wm?: (guid: string, variation: number, data: unknown, force?: boolean) => void;
  ZDi?: (plugin: AscPluginEntry, variation: number) => boolean;
  show?: (guid: string) => void;
};

type AscPluginEntry = {
  Dx?: string;
  $M?: string;
  Bsa?: boolean;
  OG?: Array<{
    url?: string;
    Bcc?: string[];
    get_Visual?: () => boolean;
  }>;
  y5b?: () => boolean;
};

type FrameEditorWindow = Window & {
  Asc?: {
    editor?: {
      asc_pluginRun?: (guid: string, variation: number, data: string) => void;
      t7?: AscPluginManager;
    };
    CPlugin?: new () => { deserialize: (raw: unknown) => void };
  };
  DE?: {
    getController?: (name: string) => PluginsController | undefined;
  };
  g_asc_plugins?: AscPluginManager;
};

/** Cross-realm safe: frameEditor nodes fail `instanceof HTMLIFrameElement` in embed. */
function isHtmlIframe(node: Element | null | undefined): node is HTMLIFrameElement {
  return !!node && node.tagName === "IFRAME";
}

function isAgentPluginFrame(node: Element | null | undefined): node is HTMLIFrameElement {
  if (!isHtmlIframe(node)) return false;
  const id = node.id || "";
  const src = node.getAttribute("src") || "";
  return id.includes("7E4A1C90") || src.includes("/office-plugins/agent");
}

function topIsSameOrigin(): boolean {
  try {
    return window.top === window || window.top?.location.origin === window.location.origin;
  } catch {
    return false;
  }
}

function collectPluginDiag(extra?: Partial<EmbedPluginDiag>): EmbedPluginDiag {
  const iframe = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
  let frameDoc = false;
  let pluginFrames = 0;
  try {
    const doc = iframe?.contentDocument || null;
    frameDoc = !!doc;
    if (doc) {
      pluginFrames = countAgentPluginFrames(doc);
    }
  } catch {
    frameDoc = false;
  }
  return {
    crossOriginIsolated: window.crossOriginIsolated === true,
    sab: typeof SharedArrayBuffer,
    topSameOrigin: topIsSameOrigin(),
    jiPatched: iframe?.dataset.jiPatched === "1",
    frameEditor: !!iframe,
    frameDoc,
    pluginFrames,
    pluginReady: false,
    ...extra,
  };
}

/** Build values inside frameEditor's realm — parent Arrays fail `instanceof Array` there. */
function frameEvalJson<T>(win: Window, value: unknown): T {
  return (win as Window & { eval: (code: string) => T }).eval(`(${JSON.stringify(value)})`);
}

function agentPluginFrameId(): string {
  return `iframe_${AGENT_PLUGIN_GUID}`;
}

function countAgentPluginFrames(doc: Document | null | undefined): number {
  if (!doc) return 0;
  const seen = new Set<Element>();
  const direct = doc.getElementById(agentPluginFrameId());
  if (direct) seen.add(direct);
  for (const node of Array.from(doc.querySelectorAll("iframe"))) {
    if (isAgentPluginFrame(node)) seen.add(node);
  }
  return seen.size;
}

/** Mirror OnlyOffice show() for invisible plugins when wm() still refuses. */
function injectAgentPluginIframe(
  win: FrameEditorWindow,
  doc: Document,
  mgr: AscPluginManager,
): string {
  const id = agentPluginFrameId();
  if (doc.getElementById(id)) return "exists";

  const plugin = mgr.e2?.[AGENT_PLUGIN_GUID];
  const variation = plugin?.OG?.[0];
  const baseUrl =
    (plugin?.$M && plugin.$M.length > 0 ? plugin.$M : null) ||
    `${location.origin}/office-plugins/agent/`;
  const relUrl = variation?.url || "index.html";
  const lang = "en";
  let theme = "light";
  try {
    const dq = (win as unknown as { AscCommon?: { Dq?: { type?: string } } }).AscCommon?.Dq;
    if (dq?.type) theme = String(dq.type);
  } catch {
    /* ignore */
  }

  // GH entry is required: plugins.js "initialize" is ignored without it.
  if (!mgr.GH) mgr.GH = {};
  if (!mgr.GH[AGENT_PLUGIN_GUID]) {
    const stub = doc.createElement("span");
    stub.setAttribute("guid", AGENT_PLUGIN_GUID);
    mgr.GH[AGENT_PLUGIN_GUID] = {
      K9: id,
    };
    (mgr.GH[AGENT_PLUGIN_GUID] as { E4?: HTMLElement; $na?: number; yzd?: boolean }).E4 = stub;
    (mgr.GH[AGENT_PLUGIN_GUID] as { $na?: number }).$na = 0;
    (mgr.GH[AGENT_PLUGIN_GUID] as { yzd?: boolean }).yzd = false;
  }

  const frame = doc.createElement("iframe");
  frame.name = id;
  frame.id = id;
  frame.src = `${baseUrl}${relUrl}?lang=${lang}&theme-type=${theme}`;
  frame.style.cssText =
    "position:absolute;top:-100px;left:0;width:10000px;height:100px;overflow:hidden;z-index:-1000";
  frame.setAttribute("frameBorder", "0");
  frame.setAttribute("allow", "autoplay");
  doc.body.appendChild(frame);
  return doc.getElementById(id) ? "injected" : "inject-failed";
}

function forceAgentPluginRun(): string {
  const iframe = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
  const win = iframe?.contentWindow as FrameEditorWindow | null | undefined;
  const doc = iframe?.contentDocument || null;
  if (!win) return "no-frame";
  try {
    const pluginsCtrl = win.DE?.getController?.("Common.Controllers.Plugins");
    const baseUrl = `${location.origin}/office-plugins/agent/`;
    // Must be frame-local: OnlyOffice parsePlugins uses `pluginsdata instanceof Array`
    // against frameEditor's Array ctor; a parent-window array always fails that check,
    // leaves canPlugins=false, and never calls asc_pluginsRegister.
    const manifestList = frameEvalJson<unknown[]>(win, [
      { ...AGENT_PLUGIN_MANIFEST, baseUrl },
    ]);
    const parts: string[] = [];

    if (pluginsCtrl) {
      parts.push("ctrl");
      if (typeof pluginsCtrl.onPluginsInit === "function") {
        pluginsCtrl.onPluginsInit(manifestList, true);
        parts.push("init");
      } else if (typeof pluginsCtrl.parsePlugins === "function") {
        pluginsCtrl.parsePlugins(manifestList as unknown[], false, true, true);
        parts.push("parse");
      }
      pluginsCtrl.autostart = frameEvalJson<string[]>(win, [AGENT_PLUGIN_GUID]);
      if (typeof pluginsCtrl.runAutoStartPlugins === "function") {
        pluginsCtrl.runAutoStartPlugins();
        parts.push("autostart");
      }
      if (pluginsCtrl.appOptions) {
        parts.push(`canPlugins=${String(pluginsCtrl.appOptions.canPlugins)}`);
      }
    } else {
      parts.push("no-ctrl");
    }

    const mgr: AscPluginManager | undefined =
      win.g_asc_plugins || win.Asc?.editor?.t7 || undefined;
    // Stuck SNb makes every later wm() a no-op.
    if (mgr && mgr.SNb) {
      mgr.SNb = null;
      parts.push("clearSNb");
    }

    const api = win.Asc?.editor || pluginsCtrl?.api;
    const run = api?.asc_pluginRun;
    if (typeof run === "function" && api) {
      run.call(api, AGENT_PLUGIN_GUID, 0, "");
      parts.push("run");
    } else {
      parts.push("no-api");
    }

    // Run wm entirely inside frameEditor so `this` / closures match the SDK.
    const wmDiag = (win as Window & { eval: (code: string) => unknown }).eval(`
      (function(){
        var g = ${JSON.stringify(AGENT_PLUGIN_GUID)};
        var mgr = window.g_asc_plugins;
        if (!mgr) return "no-mgr";
        try {
          mgr.SNb = null;
          if (typeof mgr.wm === "function") mgr.wm(g, 0, "", true);
          return [
            "e2=" + (mgr.e2 && mgr.e2[g] ? 1 : 0),
            "gh=" + (mgr.GH && mgr.GH[g] ? 1 : 0),
            "frame=" + (document.getElementById("iframe_" + g) ? 1 : 0)
          ].join(",");
        } catch (err) {
          return "wm-err:" + (err && err.message ? err.message : String(err));
        }
      })()
    `);
    parts.push(`inframe:${String(wmDiag)}`);

    const plugin = mgr?.e2?.[AGENT_PLUGIN_GUID];
    const bcc = plugin?.OG?.[0]?.Bcc;
    const zdi =
      mgr && plugin && typeof mgr.ZDi === "function" ? mgr.ZDi(plugin, 0) : null;
    parts.push(
      `bcc=${Array.isArray(bcc) ? bcc.join("|") : "?"}`,
      `base=${plugin?.$M ? "1" : "0"}`,
      `zdi=${zdi === null ? "?" : String(zdi)}`,
      `frames=${countAgentPluginFrames(doc)}`,
    );

    if (countAgentPluginFrames(doc) === 0 && mgr && doc) {
      parts.push(`fallback=${injectAgentPluginIframe(win, doc, mgr)}`);
      parts.push(`frames2=${countAgentPluginFrames(doc)}`);
    }

    parts.push(`gPlugins=${mgr?.plugins?.length ?? "?"}`);
    return parts.join(",");
  } catch (err) {
    return err instanceof Error ? err.message : String(err);
  }
}

function parseAgentPayload(raw: unknown): {
  type?: string;
  event?: string;
  requestId?: string;
  result?: unknown;
  error?: string;
} | null {
  let data = raw;
  if (typeof raw === "string") {
    try {
      data = JSON.parse(raw) as unknown;
    } catch {
      return null;
    }
  }
  if (!data || typeof data !== "object") return null;
  const row = data as { type?: string };
  if (row.type !== "agent-doc") return null;
  return data as {
    type?: string;
    event?: string;
    requestId?: string;
    result?: unknown;
    error?: string;
  };
}

export default function EmbedPage() {
  const defaultLanguage = useResolvedLanguage();
  const hasHydrated = useHasHydrated();
  const [loading, setLoading] = useState(true);
  const [appRoot, setAppRoot] = useState<string>(() => getAppRoot());
  const [warmupLabel, setWarmupLabel] = useState<string | null>(null);
  const isDirty = useRef(false);
  const editorRef = useRef<DocEditor | null>(null);
  const serverRef = useRef<EditorServer | null>(null);
  const pluginModeRef = useRef<PluginMode>("agent");
  const hostOriginRef = useRef<string>("*");
  const fileMetaRef = useRef({ fileName: "document.docx", fileType: "docx" });
  const saveRequestIdRef = useRef<string | null>(null);
  const hideChromeRef = useRef(false);
  const readyTimerRef = useRef<number | null>(null);
  const openingRef = useRef(false);
  const pluginReadyFromChild = useRef(false);
  const pluginSourceRef = useRef<MessageEventSource | null>(null);

  const stopReadyPing = () => {
    if (readyTimerRef.current != null) {
      window.clearInterval(readyTimerRef.current);
      readyTimerRef.current = null;
    }
  };

  const postToHost = (msg: EditorToHostMessage, transferables?: Transferable[]) => {
    if (typeof window === "undefined" || !window.parent || window.parent === window) return;
    const target = hostOriginRef.current || "*";
    const send = (dest: string) => {
      if (transferables && transferables.length > 0) {
        window.parent.postMessage(msg, dest, transferables);
      } else {
        window.parent.postMessage(msg, dest);
      }
    };
    try {
      send(target);
    } catch {
      if (target !== "*") send("*");
    }
    // Handshake events: also send "*" so a mismatched target origin cannot drop them.
    if (
      (!transferables || transferables.length === 0) &&
      target !== "*" &&
      (msg.type === "ready" || msg.type === "appReady" || msg.type === "documentReady" || msg.type === "error")
    ) {
      try {
        send("*");
      } catch {
        /* ignore */
      }
    }
  };

  useEffect(() => {
    const searchParams = new URLSearchParams(window.location.search);
    const isWarmup = searchParams.get("warmup") === "1";
    void loadEmbedPartnerConfig();
    if (!isWarmup) {
      const pingReady = () =>
        postToHost({
          type: "ready",
          version: EMBED_PROTOCOL_VERSION,
          diag: collectPluginDiag({ pluginReady: pluginReadyFromChild.current }),
        });
      pingReady();
      readyTimerRef.current = window.setInterval(pingReady, 500);
    }

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") postToHost({ type: "escape" });
    };
    window.addEventListener("keydown", handleKeyDown);

    const onAgentMessage = (event: MessageEvent) => {
      const body = parseAgentPayload(event.data);
      if (!body) return;
      if (body.event === "completion/request") {
        const row =
          body.result && typeof body.result === "object"
            ? (body.result as { prefix?: string; suffix?: string })
            : {};
        postToHost({
          type: "completionRequest",
          prefix: String(row.prefix || ""),
          suffix: String(row.suffix || ""),
        });
        return;
      }
      if (body.event === "ghost") {
        const row =
          body.result && typeof body.result === "object"
            ? (body.result as { shown?: boolean; text?: string; accepted?: boolean })
            : {};
        postToHost({
          type: "ghost",
          shown: row.shown === true,
          text: row.text,
          accepted: row.accepted,
        });
        return;
      }
      if (body.requestId === "ready") {
        pluginReadyFromChild.current = true;
        if (event.source) pluginSourceRef.current = event.source;
        return;
      }
      if (body.requestId) {
        postToHost({
          type: "commandResult",
          requestId: String(body.requestId),
          result: body.result,
          error: body.error || undefined,
        });
      }
    };
    window.addEventListener("message", onAgentMessage);

    return () => {
      stopReadyPing();
      window.removeEventListener("keydown", handleKeyDown);
      window.removeEventListener("message", onAgentMessage);
    };
  }, []);

  useLayoutEffect(() => {
    if (!hasHydrated) return;

    const searchParams = new URLSearchParams(window.location.search);
    const resolvedAppRoot = getAppRoot(searchParams.get("v") || searchParams.get("version"));
    setAppRoot(resolvedAppRoot);
    const apiUrl = resolvedAppRoot + API_JS;
    const isWarmup = searchParams.get("warmup") === "1";

    if (isWarmup) {
      setWarmupLabel("Preparando cache do editor…");
      void prefetchEditorAssets(resolvedAppRoot, (done, total) => {
        setWarmupLabel(`Preparando cache do editor… ${done}/${total}`);
      })
        .then((result) => {
          postToHost({ type: "warmed", cached: result.cached, failed: result.failed });
          setLoading(false);
          setWarmupLabel("Cache do editor pronto.");
        })
        .catch((err: unknown) => {
          postToHost({
            type: "error",
            message: err instanceof Error ? err.message : "Warmup failed",
          });
          setLoading(false);
        });
      return;
    }

    const server = new EditorServer({
      getState: () => ({ plugins: pluginModeRef.current }),
      persistFile: async (blob, fileName) => {
        const bytes = await blob.arrayBuffer();
        const requestId = saveRequestIdRef.current || undefined;
        saveRequestIdRef.current = null;
        isDirty.current = false;
        postToHost(
          {
            type: "saved",
            requestId,
            fileName: fileName || fileMetaRef.current.fileName,
            fileType: fileMetaRef.current.fileType,
            mime: embedMimeForType(fileMetaRef.current.fileType),
            bytes,
          },
          [bytes],
        );
        postToHost({ type: "dirty", value: false });
      },
    });
    serverRef.current = server;

    MockSocket.on("connect", server.handleConnect);
    MockSocket.on("disconnect", server.handleDisconnect);

    const applyEditorChrome = (doc: Document | null | undefined) => {
      if (hideChromeRef.current) injectPreviewChrome(doc);
      else injectEmbedChrome(doc);
    };

    const patchEditorFrame = () => {
      const iframe = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
      const win = iframe?.contentWindow as typeof window | undefined;
      const iframeDoc = iframe?.contentDocument;
      if (!iframe || !iframeDoc || !win) throw new Error("Iframe not loaded");
      if (iframe.dataset.jiPatched === "1") {
        applyEditorChrome(iframeDoc);
        return;
      }

      const xhr = createXHRProxy(win.XMLHttpRequest, win);
      const fetchProxy = createFetchProxy(win);
      const NativeWorker = win.Worker;

      xhr.use((request: Request) => server.handleRequest(request));
      xhr.useSync((url) => {
        try {
          const u = new URL(url, location.origin);
          if (isPluginsJsonPath(u.pathname)) {
            const state = pluginModeRef.current;
            if (state === "none") {
              return {
                body: JSON.stringify({ url: "", pluginsData: [], autostart: [] }),
                contentType: "application/json",
              };
            }
            if (state === "agent") {
              return {
                body: JSON.stringify(getAgentPluginsData(location.origin)),
                contentType: "application/json",
              };
            }
            return null;
          }
          if (isAgentPluginConfigPath(u.pathname)) {
            const baseUrl = `${location.origin}/office-plugins/agent/`;
            return {
              body: JSON.stringify({ ...AGENT_PLUGIN_MANIFEST, baseUrl }),
              contentType: "application/json",
            };
          }
        } catch {
          return null;
        }
        return null;
      });
      fetchProxy.use((request: Request) => server.handleRequest(request));

      Object.assign(win, {
        io,
        XMLHttpRequest: xhr,
        fetch: fetchProxy,
        Worker: function WorkerProxy(url: string, options?: WorkerOptions) {
          const parsed = new URL(url, location.origin);
          return new NativeWorker(
            parsed.href.replace(parsed.origin, location.origin),
            options,
          );
        },
      });

      win.addEventListener("keydown", (e: KeyboardEvent) => {
        if (e.key === "Escape") postToHost({ type: "escape" });
      });

      applyEditorChrome(iframeDoc);
      iframe.dataset.jiPatched = "1";

      // Unblock OnlyOffice mergePlugins if the async plugins.json fetch stalls
      // (observed when /embed is a third-party iframe under workspace).
      const ensurePluginLoadConfig = () => {
        const common = (win as Window & { Common?: { Utils?: { loadConfig?: unknown } } }).Common;
        if (!common?.Utils || typeof common.Utils.loadConfig !== "function") return false;
        if ((common.Utils as { __jiAgentLoadConfig?: boolean }).__jiAgentLoadConfig) return true;
        const nativeLoad = common.Utils.loadConfig.bind(common.Utils) as (
          url: string,
          cb: (obj: unknown) => void,
        ) => void;
        common.Utils.loadConfig = (url: string, cb: (obj: unknown) => void) => {
          try {
            const resolved = new URL(url, win.location.href);
            if (isPluginsJsonPath(resolved.pathname) && pluginModeRef.current === "agent") {
              cb(getAgentPluginsData(location.origin));
              return;
            }
            if (isPluginsJsonPath(resolved.pathname) && pluginModeRef.current === "none") {
              cb({ url: "", pluginsData: [], autostart: [] });
              return;
            }
          } catch {
            /* fall through */
          }
          nativeLoad(url, cb);
        };
        (common.Utils as { __jiAgentLoadConfig?: boolean }).__jiAgentLoadConfig = true;
        return true;
      };
      if (!ensurePluginLoadConfig()) {
        let n = 0;
        const iv = win.setInterval(() => {
          if (ensurePluginLoadConfig() || ++n > 200) win.clearInterval(iv);
        }, 25);
      }
    };

    const watchEditorFrame = () => {
      const tryPatch = () => {
        try {
          patchEditorFrame();
          return true;
        } catch {
          return false;
        }
      };
      if (tryPatch()) return;
      const root = document.getElementById("placeholder") || document.body;
      const observer = new MutationObserver(() => {
        if (tryPatch()) observer.disconnect();
      });
      observer.observe(root, { childList: true, subtree: true });
      // Sync plugins.json XHR races ahead of MutationObserver; poll hard for 2s.
      let n = 0;
      const poll = window.setInterval(() => {
        if (tryPatch() || ++n > 200) {
          window.clearInterval(poll);
          if (tryPatch()) observer.disconnect();
        }
      }, 10);
      window.setTimeout(() => {
        window.clearInterval(poll);
        observer.disconnect();
      }, 30_000);
    };

    const destroyEditor = () => {
      pluginReadyFromChild.current = false;
      pluginSourceRef.current = null;
      try {
        editorRef.current?.destroyEditor?.();
      } catch {
        /* already torn down */
      }
      editorRef.current = null;
      const placeholder = document.getElementById("placeholder");
      if (placeholder) placeholder.innerHTML = "";
    };

    const createEditorInstance = (options: {
      editing: boolean;
      lang: string;
      uiTheme: string;
      fileName: string;
      hideChrome: boolean;
    }) => {
      const doc = server.getDocument();
      const user = server.getUser();
      const documentType = getDocumentType(doc.fileType);
      const canEdit = options.editing && doc.fileType !== "pdf";
      const isPreview = options.hideChrome || !options.editing;

      server.setClient({
        buildVersion: window.DocsAPI?.DocEditor?.version() || "9.3.0",
      });

      // Start observing before DocEditor creates frameEditor — otherwise the
      // first sync plugins.json XHR misses our proxy and the Agent never mounts.
      watchEditorFrame();

      const editor = new window.DocsAPI!.DocEditor("placeholder", {
        document: {
          fileType: doc.fileType,
          key: doc.key,
          title: options.fileName || doc.title,
          url: doc.url,
          permissions: {
            edit: canEdit,
            chat: false,
            rename: canEdit,
            protect: canEdit,
            review: false,
            print: false,
            download: !isPreview,
            plugins: pluginModeRef.current === "agent",
          },
        },
        documentType,
        editorConfig: {
          mode: canEdit ? "edit" : "view",
          lang: toOnlyOfficeLang(options.lang || defaultLanguage),
          coEditing: {
            mode: "fast",
            change: false,
          },
          user: {
            ...user,
          },
          ...(pluginModeRef.current === "agent" ? { plugins: getAgentPluginsData(location.origin) } : {}),
          customization: {
            uiTheme: options.uiTheme,
            autosave: canEdit,
            compactHeader: true,
            toolbarHideFileName: true,
            hideRightMenu: true,
            hideToolbar: isPreview,
            zoom: isPreview ? -1 : 100,
            layout: {
              header: isPreview ? false : { user: false, users: false },
              leftMenu: !isPreview,
              rightMenu: false,
              statusBar: !isPreview,
              toolbar: !isPreview,
            },
            features: {
              spellcheck: {
                change: false,
              },
            },
            plugins: pluginModeRef.current === "agent",
            about: !isPreview,
            feedback: false,
            customer: isPreview
              ? undefined
              : {
                  name: "EDITOR GRATUITO (editor.app.br / editor.com.br)",
                  www: PUBLIC_EDITOR_URL,
                  mail: LEGAL_CONTACT,
                  info: "Software livre GNU AGPL v3. Motor ONLYOFFICE Community Edition © Ascensio System SIA. Código-fonte: editor.app.br/source",
                  logo: `${location.origin}${PUBLIC_CUBE_LOGO}`,
                  logoDark: `${location.origin}${PUBLIC_CUBE_LOGO}`,
                },
            logo: {
              visible: false,
              url: PUBLIC_ABOUT_URL,
            },
          },
        },
        events: {
          onAppReady: async () => {
            try {
              patchEditorFrame();
            } catch (err) {
              console.error(err);
            }
            setLoading(false);
            postToHost({ type: "appReady" });
          },
          onDocumentReady: () => {
            openingRef.current = false;
            setLoading(false);
            const iframe = document.querySelector<HTMLIFrameElement>(
              'iframe[name="frameEditor"]',
            );
            applyEditorChrome(iframe?.contentDocument);
            if (pluginModeRef.current !== "agent") {
              postToHost({ type: "documentReady" });
              return;
            }
            // Host already shows the UI on appReady. Only advertise documentReady
            // once the Agent plugin can receive commands — an early soft ready made
            // the workspace bind tools while iframe_asc.* never existed, then
            // retries remounted the host and hit the 45s handshake sad-pane.
            let finished = false;
            let forceAttempted = false;
            let forceResult: string | undefined;
            const finish = (forceRun?: string) => {
              if (finished) return;
              finished = true;
              postToHost({
                type: "documentReady",
                diag: collectPluginDiag({
                  // Frames alone are not enough — callCommand may still be missing.
                  pluginReady: pluginCanReceive(),
                  forceRun: forceRun ?? forceResult,
                }),
              });
            };
            // Agent tools need Asc.plugin.callCommand (ready ping), not just iframe_asc.*.
            const pluginLive = () => pluginCanReceive();
            if (pluginLive()) {
              finish();
              return;
            }
            const tick = window.setInterval(() => {
              if (!pluginLive()) return;
              window.clearInterval(tick);
              finish(forceAttempted ? forceResult || "ran" : undefined);
            }, 200);
            // Third-party /embed (workspace): OnlyOffice autostart sometimes never
            // mounts iframe_asc.* — register Agent + nudge asc_pluginRun.
            window.setTimeout(() => {
              if (finished || forceAttempted || pluginLive()) return;
              forceAttempted = true;
              forceResult = forceAgentPluginRun();
              console.warn("[embed] forceAgentPluginRun@1.5s:", forceResult);
            }, 1_500);
            // Second attempt if the controller was not ready yet.
            window.setTimeout(() => {
              if (finished || pluginLive()) return;
              forceAttempted = true;
              forceResult = forceAgentPluginRun();
              console.warn("[embed] forceAgentPluginRun@5s:", forceResult);
            }, 5_000);
            // Keep nudging while waiting for the Agent ready ping (callCommand).
            window.setTimeout(() => {
              if (finished || pluginLive()) return;
              forceResult = forceAgentPluginRun();
              console.warn("[embed] forceAgentPluginRun@15s:", forceResult);
            }, 15_000);
            // Safety: never fatal-error the human editor. After 45s still notify
            // so Agent tools can surface a clear timeout instead of hanging forever.
            window.setTimeout(() => {
              window.clearInterval(tick);
              finish(forceAttempted ? `timeout-after-force:${forceResult || "?"}` : "timeout");
            }, 45_000);
          },
          onDocumentStateChange: (e: { data: boolean }) => {
            if (e.data) {
              isDirty.current = true;
              postToHost({ type: "dirty", value: true });
            }
          },
          onError: (e: unknown) => {
            console.error("Editor error", e);
            postToHost({ type: "error", message: String(e) });
          },
          onDownloadAs: () => {
            postToHost({ type: "print" });
          },
        },
        type: "desktop",
        width: "100%",
        height: "100%",
      });

      editorRef.current = editor;
      // watchEditorFrame already started before DocEditor; patch again in case
      // the frame appeared synchronously during construction.
      watchEditorFrame();
      setLoading(false);
      postToHost({ type: "appReady" });
      return editor;
    };

    const loadEditorScript = (callback: () => void) => {
      if (window.DocsAPI && window.DocsAPI.DocEditor) {
        callback();
        return;
      }
      let script = document.querySelector<HTMLScriptElement>(`script[src="${apiUrl}"]`);
      if (!script) {
        script = document.createElement("script");
        script.src = apiUrl;
        document.head.appendChild(script);
      }
      script.onload = () => callback();
      script.onerror = () => {
        openingRef.current = false;
        postToHost({ type: "error", message: "Failed to load DocsAPI script" });
      };
    };

    const findPluginFrames = (): HTMLIFrameElement[] => {
      const editor = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
      let doc: Document | null = null;
      try {
        doc = editor?.contentDocument || null;
      } catch {
        doc = null;
      }
      if (!doc) return [];
      const guid = AGENT_PLUGIN_GUID;
      const shortGuid = guid.replace(/^asc\./, "");
      const nodes = [
        doc.getElementById(`iframe_${guid}`),
        doc.getElementById(`iframe_${shortGuid}`),
        doc.getElementById(guid),
        ...Array.from(doc.querySelectorAll("iframe")),
      ];
      const seen = new Set<HTMLIFrameElement>();
      for (const node of nodes) {
        // Do not use instanceof — frameEditor nodes are a different realm.
        if (!isHtmlIframe(node)) continue;
        const id = node.id || "";
        const src = node.getAttribute("src") || "";
        if (
          id.includes("7E4A1C90") ||
          id.startsWith("iframe_asc.") ||
          src.includes("/office-plugins/agent") ||
          node === doc.getElementById(`iframe_${guid}`)
        ) {
          seen.add(node);
        }
      }
      return [...seen];
    };

    const deliverPluginCommand = (data: Record<string, unknown>): boolean => {
      const targets = new Set<Window>();
      const source = pluginSourceRef.current;
      if (source && typeof (source as Window).postMessage === "function") {
        targets.add(source as Window);
      }
      for (const plugin of findPluginFrames()) {
        if (plugin.contentWindow) targets.add(plugin.contentWindow);
      }
      if (!targets.size) return false;
      let delivered = false;
      for (const win of targets) {
        try {
          // Prefer postMessage (same path as the ready ping). Direct
          // onExternalPluginMessage across windows can leave callCommand hanging.
          win.postMessage(data, "*");
          delivered = true;
        } catch {
          /* cross-origin / detached */
        }
      }
      return delivered;
    };

    const pluginCanReceive = () => pluginReadyFromChild.current;

    const deliverPluginCommandWhenReady = (data: Record<string, unknown>, requestId?: string) => {
      if (deliverPluginCommand(data)) return;
      const started = Date.now();
      const tick = window.setInterval(() => {
        if (deliverPluginCommand(data)) {
          window.clearInterval(tick);
          return;
        }
        if (Date.now() - started < 20_000) return;
        window.clearInterval(tick);
        postToHost({
          type: "commandResult",
          requestId: requestId || "",
          error: "OnlyOffice agent plugin is not loaded.",
        });
      }, 250);
    };

    const handleHostMessage = async (event: MessageEvent<HostToEditorMessage>) => {
      if (!isAllowedEmbedHostOrigin(event.origin)) {
        await loadEmbedPartnerConfig();
        if (!isAllowedEmbedHostOrigin(event.origin)) return;
      }
      hostOriginRef.current = event.origin;
      const data = event.data;
      if (!data || typeof data !== "object") return;

      if (data.type === "open") {
        if (openingRef.current || editorRef.current) return;
        try {
          stopReadyPing();
          setLoading(true);
          destroyEditor();
          openingRef.current = true;
          pluginModeRef.current = data.plugins === "none" || data.variant === "preview" ? "none" : "agent";
          hideChromeRef.current = data.hideChrome === true || data.variant === "preview";
          fileMetaRef.current = { fileName: data.fileName, fileType: String(data.fileType) };
          const blob = new Blob([data.bytes]);
          const file = new File([blob], data.fileName, { type: "application/octet-stream" });
          await server.open(file, {
            fileName: data.fileName,
            fileType: data.fileType,
          });
          const uiTheme =
            data.theme === "dark" ? "theme-dark" : data.theme === "light" ? "theme-light" : "theme-white";
          loadEditorScript(() => {
            createEditorInstance({
              editing: data.editing,
              lang: toOnlyOfficeLang(data.lang || "pt"),
              uiTheme,
              fileName: data.fileName,
              hideChrome: hideChromeRef.current,
            });
          });
        } catch (err: unknown) {
          openingRef.current = false;
          const message = err instanceof Error ? err.message : "Failed to open document";
          postToHost({ type: "error", message, requestId: data.requestId });
          setLoading(false);
        }
        return;
      }

      if (data.type === "save") {
        const editor = editorRef.current;
        if (!editor) {
          postToHost({ type: "error", message: "Editor is not ready.", requestId: data.requestId });
          return;
        }
        saveRequestIdRef.current = data.requestId;
        try {
          editor.downloadAs(fileMetaRef.current.fileType);
        } catch (err: unknown) {
          saveRequestIdRef.current = null;
          postToHost({
            type: "error",
            message: err instanceof Error ? err.message : "Failed to save",
            requestId: data.requestId,
          });
        }
        return;
      }

      if (data.type === "setTheme") {
        const theme = data.theme === "dark" ? "theme-dark" : "theme-light";
        try {
          editorRef.current?.serviceCommand?.("theme", theme);
        } catch {
          /* theme applies on next open */
        }
        return;
      }

      if (data.type === "print") {
        postToHost({ type: "print" });
        return;
      }

      if (data.type === "destroy") {
        openingRef.current = false;
        destroyEditor();
        return;
      }

      if (data.type === "command") {
        deliverPluginCommandWhenReady(
          {
            type: data.name,
            requestId: data.requestId,
            payload: data.payload || {},
          },
          data.requestId,
        );
      }
    };

    window.addEventListener("message", handleHostMessage);

    return () => {
      window.removeEventListener("message", handleHostMessage);
      MockSocket.off("connect", server.handleConnect);
      MockSocket.off("disconnect", server.handleDisconnect);
      destroyEditor();
      serverRef.current = null;
    };
  }, [hasHydrated, defaultLanguage]);

  return (
    <div className="relative w-screen h-screen overflow-hidden bg-background m-0 p-0">
      {loading && (
        <div className="absolute inset-0 z-50 flex flex-col items-center justify-center bg-background/80 backdrop-blur-xs">
          <Loader2 className="w-8 h-8 animate-spin text-emerald-600 mb-2" />
          <p className="text-xs text-text-secondary">
            {warmupLabel || "Carregando editor seguro..."}
          </p>
        </div>
      )}
      <div id="placeholder" className="w-full h-full">
        <iframe className="w-0 h-0 hidden" src={appRoot + PRELOAD_HTML} />
      </div>
    </div>
  );
}
