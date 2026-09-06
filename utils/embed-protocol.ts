/**
 * Thin-embed protocol between a host app and editor.app.br /embed.
 * Host exchanges document bytes and JSON commands only (AGPL isolation).
 */

export const EMBED_PROTOCOL_VERSION = "1.1.0";

export type EmbedDocumentType =
  | "docx"
  | "xlsx"
  | "pptx"
  | "pdf"
  | "doc"
  | "xls"
  | "ppt"
  | "txt"
  | "csv"
  | "odt"
  | "ods"
  | "odp"
  | "rtf"
  | "docm"
  | "xlsm"
  | "pptm";

export type EmbedTheme = "dark" | "light" | "system";
export type EmbedVariant = "editor" | "preview";
export type EmbedPluginMode = "agent" | "none";

export type HostToEditorMessage =
  | {
      type: "open";
      requestId: string;
      fileName: string;
      fileType: EmbedDocumentType | string;
      editing: boolean;
      theme?: EmbedTheme;
      lang?: string;
      bytes: ArrayBuffer;
      variant?: EmbedVariant;
      hideChrome?: boolean;
      plugins?: EmbedPluginMode;
    }
  | {
      type: "save";
      requestId: string;
    }
  | {
      type: "setTheme";
      theme: "dark" | "light";
    }
  | {
      type: "print";
    }
  | {
      type: "destroy";
    }
  | {
      type: "command";
      requestId: string;
      name: string;
      payload?: Record<string, unknown>;
    };

/** Debug fields for Agent plugin mount failures (third-party /embed). */
export type EmbedPluginDiag = {
  crossOriginIsolated: boolean;
  sab: string;
  topSameOrigin: boolean;
  jiPatched: boolean;
  frameEditor: boolean;
  frameDoc: boolean;
  pluginFrames: number;
  pluginReady: boolean;
  forceRun?: string;
};

export type EditorToHostMessage =
  | {
      type: "ready";
      version: string;
      diag?: EmbedPluginDiag;
    }
  | {
      type: "dirty";
      value: boolean;
    }
  | {
      type: "saved";
      requestId?: string;
      fileName: string;
      fileType: string;
      mime: string;
      bytes: ArrayBuffer;
    }
  | {
      type: "error";
      message: string;
      requestId?: string;
    }
  | {
      type: "escape";
    }
  | {
      type: "print";
    }
  | {
      type: "commandResult";
      requestId: string;
      result?: unknown;
      error?: string;
    }
  | {
      type: "completionRequest";
      prefix: string;
      suffix?: string;
    }
  | {
      type: "ghost";
      shown: boolean;
      text?: string;
      accepted?: boolean;
    }
  | {
      type: "appReady";
    }
  | {
      type: "documentReady";
      diag?: EmbedPluginDiag;
    }
  | {
      type: "warmed";
      cached: number;
      failed: number;
    };

export const WARMUP_CACHE_PREFIX = "editor-static-";

export function embedMimeForType(fileType: string): string {
  const ext = fileType.replace(/^\./, "").toLowerCase();
  if (ext === "docx") {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  }
  if (ext === "xlsx") {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
  if (ext === "pptx") {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  if (ext === "pdf") return "application/pdf";
  if (ext === "odt") return "application/vnd.oasis.opendocument.text";
  if (ext === "ods") return "application/vnd.oasis.opendocument.spreadsheet";
  if (ext === "odp") return "application/vnd.oasis.opendocument.presentation";
  if (ext === "csv") return "text/csv";
  if (ext === "txt") return "text/plain";
  return "application/octet-stream";
}
