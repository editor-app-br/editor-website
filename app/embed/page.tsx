"use client";

import { useLayoutEffect, useRef, useEffect, useState } from "react";
import { useAppStore, useResolvedLanguage, useHasHydrated } from "@/store";
import {
  API_JS,
  getAppRoot,
  getDocumentType,
  PRELOAD_HTML,
} from "@/utils/editor/utils";
import io, { MockSocket } from "@/utils/editor/socket";
import { createFetchProxy } from "@/utils/editor/fetch";
import { createXHRProxy } from "@/utils/editor/xhr";
import { DocEditor } from "@/utils/editor/types";
import { HostToEditorMessage, EditorToHostMessage } from "@/utils/embed-protocol";
import {
  EDITOR_WORDMARK_BLACK,
  EDITOR_WORDMARK_WHITE,
} from "@/utils/branding";
import { Loader2 } from "lucide-react";

export default function EmbedPage() {
  const server = useAppStore((state) => state.server);
  const defaultLanguage = useResolvedLanguage();
  const defaultTheme = useAppStore((state) => state.theme);
  const hasHydrated = useHasHydrated();
  const [loading, setLoading] = useState(true);
  const [appRoot, setAppRoot] = useState<string>("/v9.3.1-1");
  const isDirty = useRef(false);
  const editorRef = useRef<DocEditor | null>(null);

  const postToHost = (msg: EditorToHostMessage, transferables?: Transferable[]) => {
    if (typeof window !== "undefined" && window.parent && window.parent !== window) {
      if (transferables && transferables.length > 0) {
        window.parent.postMessage(msg, "*", transferables);
      } else {
        window.parent.postMessage(msg, "*");
      }
    }
  };

  useEffect(() => {
    postToHost({ type: "ready", version: "1.0.0" });

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        postToHost({ type: "escape" });
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  useLayoutEffect(() => {
    if (!hasHydrated) return;

    const searchParams = new URLSearchParams(window.location.search);
    const resolvedAppRoot = getAppRoot(searchParams.get("v") || searchParams.get("version"));
    setAppRoot(resolvedAppRoot);

    const apiUrl = resolvedAppRoot + API_JS;

    MockSocket.on("connect", server.handleConnect);
    MockSocket.on("disconnect", server.handleDisconnect);

    const onAppReady = () => {
      const iframe = document.querySelector<HTMLIFrameElement>(
        'iframe[name="frameEditor"]',
      );
      const win = iframe?.contentWindow as typeof window;
      const iframeDoc = iframe?.contentDocument;
      if (!iframeDoc || !win) {
        throw new Error("Iframe not loaded");
      }

      const xhr = createXHRProxy(win.XMLHttpRequest);
      const fetchProxy = createFetchProxy(win);
      const _Worker = win.Worker;

      xhr.use((request: Request) => server.handleRequest(request));
      fetchProxy.use((request: Request) => server.handleRequest(request));

      Object.assign(win, {
        io: io,
        XMLHttpRequest: xhr,
        fetch: fetchProxy,
        Worker: function (url: string, options?: WorkerOptions) {
          const u = new URL(url, location.origin);
          return new _Worker(
            u.href.replace(u.origin, location.origin),
            options,
          );
        },
      });

      win.addEventListener("keydown", (e: KeyboardEvent) => {
        if (e.key === "Escape") {
          postToHost({ type: "escape" });
        }
      });
    };

    const createEditorInstance = (options: {
      editing: boolean;
      lang: string;
      uiTheme: string;
      fileName: string;
    }) => {
      const doc = server.getDocument();
      const user = server.getUser();
      const documentType = getDocumentType(doc.fileType);

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
            edit: options.editing && doc.fileType !== "pdf",
            chat: false,
            rename: options.editing,
            protect: options.editing,
            review: false,
            print: false,
          },
        },
        documentType: documentType,
        editorConfig: {
          lang: options.lang || defaultLanguage,
          coEditing: {
            mode: "fast",
            change: false,
          },
          user: {
            ...user,
          },
          customization: {
            uiTheme: options.uiTheme || defaultTheme,
            features: {
              spellcheck: {
                change: false,
              },
            },
            logo: {
              image: location.origin + EDITOR_WORDMARK_BLACK,
              imageDark: location.origin + EDITOR_WORDMARK_WHITE,
              url: location.origin,
            },
          },
        },
        events: {
          onAppReady: async () => {
            onAppReady();
            setLoading(false);
          },
          onDocumentReady: () => {
            setLoading(false);
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
          onSaveDocument: () => {
            isDirty.current = false;
            postToHost({ type: "dirty", value: false });
          },
          onSave: () => {
            isDirty.current = false;
            postToHost({ type: "dirty", value: false });
          },
        },
        type: "desktop",
        width: "100%",
        height: "100%",
      });

      editorRef.current = editor;
      return editor;
    };

    const loadEditorScript = (callback: () => void) => {
      if (window.DocsAPI && window.DocsAPI.DocEditor) {
        callback();
        return;
      }
      let script = document.querySelector<HTMLScriptElement>(
        `script[src="${apiUrl}"]`,
      );
      if (!script) {
        script = document.createElement("script");
        script.src = apiUrl;
        document.head.appendChild(script);
      }
      script.onload = () => callback();
      script.onerror = (err) => {
        console.error("Failed to load DocsAPI script", err);
        postToHost({ type: "error", message: "Failed to load DocsAPI script" });
      };
    };

    const handleHostMessage = async (event: MessageEvent<HostToEditorMessage>) => {
      const data = event.data;
      if (!data || typeof data !== "object") return;

      if (data.type === "open") {
        try {
          setLoading(true);
          const blob = new Blob([data.bytes]);
          const file = new File([blob], data.fileName, { type: "application/octet-stream" });
          
          await server.open(file, {
            fileName: data.fileName,
            fileType: data.fileType,
          });

          loadEditorScript(() => {
            createEditorInstance({
              editing: data.editing,
              lang: data.lang || "pt-BR",
              uiTheme: data.theme || "light",
              fileName: data.fileName,
            });
          });
        } catch (err: any) {
          postToHost({ type: "error", message: err?.message || "Failed to open document" });
          setLoading(false);
        }
      } else if (data.type === "destroy") {
        editorRef.current?.destroyEditor?.();
        editorRef.current = null;
      }
    };

    window.addEventListener("message", handleHostMessage);

    return () => {
      window.removeEventListener("message", handleHostMessage);
      MockSocket.off("connect", server.handleConnect);
      MockSocket.off("disconnect", server.handleDisconnect);
      editorRef.current?.destroyEditor?.();
    };
  }, [hasHydrated]);

  return (
    <div className="relative w-screen h-screen overflow-hidden bg-background m-0 p-0">
      {loading && (
        <div className="absolute inset-0 z-50 flex flex-col items-center justify-center bg-background/80 backdrop-blur-xs">
          <Loader2 className="w-8 h-8 animate-spin text-emerald-600 mb-2" />
          <p className="text-xs text-text-secondary">Carregando editor seguro...</p>
        </div>
      )}
      <div id="placeholder" className="w-full h-full">
        <iframe
          className="w-0 h-0 hidden"
          src={appRoot + PRELOAD_HTML}
        />
      </div>
    </div>
  );
}
