"use client";

import { useLayoutEffect, useRef, useEffect, useState } from "react";
import { useResolvedLanguage, useHasHydrated } from "@/store";
import {
  API_JS,
  getAppRoot,
  getDocumentType,
  PRELOAD_HTML,
} from "@/utils/editor/utils";
import io, { MockSocket } from "@/utils/editor/socket";
import { createFetchProxy } from "@/utils/editor/fetch";
import { createXHRProxy } from "@/utils/editor/xhr";
import { DocEditor, PluginMode } from "@/utils/editor/types";
import { EditorServer } from "@/utils/editor/server";
import {
  EMBED_PROTOCOL_VERSION,
  embedMimeForType,
  EditorToHostMessage,
  HostToEditorMessage,
} from "@/utils/embed-protocol";
import { isAllowedEmbedHostOrigin, loadEmbedPartnerConfig } from "@/utils/embed-origins";
import { injectEmbedChrome, injectPreviewChrome } from "@/utils/embed-chrome";
import { prefetchEditorAssets } from "@/utils/editor/warmup";
import { AGENT_PLUGIN_GUID, getAgentPluginsData } from "@/utils/editor/plugins";
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
      const pingReady = () => postToHost({ type: "ready", version: EMBED_PROTOCOL_VERSION });
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
      if (body.requestId && body.requestId !== "ready") {
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

      const xhr = createXHRProxy(win.XMLHttpRequest);
      const fetchProxy = createFetchProxy(win);
      const NativeWorker = win.Worker;

      xhr.use((request: Request) => server.handleRequest(request));
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
      window.setTimeout(() => observer.disconnect(), 30_000);
    };

    const destroyEditor = () => {
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
          },
        },
        documentType,
        editorConfig: {
          mode: canEdit ? "edit" : "view",
          lang: options.lang || defaultLanguage,
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
            postToHost({ type: "documentReady" });
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
      // Do not wait for onAppReady: the host handshake is 45s, and agent plugin
      // load can delay that callback. Patch as soon as frameEditor exists.
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

    const deliverPluginCommand = (data: Record<string, unknown>): boolean => {
      const iframe = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
      let plugin: HTMLIFrameElement | null = null;
      try {
        plugin = iframe?.contentDocument?.getElementById(
          `iframe_${AGENT_PLUGIN_GUID}`,
        ) as HTMLIFrameElement | null;
      } catch {
        plugin = null;
      }
      const api = plugin?.contentWindow as (Window & { Asc?: { plugin?: PluginHost } }) | null;
      const host = api?.Asc?.plugin;
      if (!host?.onExternalPluginMessage) return false;
      host.onExternalPluginMessage(data);
      return true;
    };

    const deliverPluginCommandWhenReady = (data: Record<string, unknown>, requestId?: string) => {
      if (deliverPluginCommand(data)) return;
      const started = Date.now();
      const tick = window.setInterval(() => {
        if (deliverPluginCommand(data)) {
          window.clearInterval(tick);
          return;
        }
        if (Date.now() - started < 12_000) return;
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
              lang: data.lang || "pt-BR",
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
