import { converter } from "./x2t";
import type { EditorServer } from "./server";

type CompareHost = Window & {
  AscCommonWord?: {
    CompareDocuments: (file: {
      GetBinary: () => Uint8Array;
      GetImageMap: () => Record<string, string>;
      IsValid: () => boolean;
    }) => void;
  };
  Asc?: {
    c_oAscDisplayModeInReview?: { Edit: number };
    editor?: { asc_SetDisplayModeInReview?: (mode: number) => void };
  };
  AscCommon?: {
    CollaborativeEditing?: { m_nUseType?: number };
  };
};

/**
 * Compare inside the local WASM editor.
 * setRevisedFile asks a document server to convert the other file; this build
 * has no document server, so that call returns without redlines.
 */
export async function compareWordDocument(options: {
  bytes: ArrayBuffer;
  fileType: string;
  server: EditorServer;
}): Promise<void> {
  const fileType = options.fileType.replace(/^\./, "").toLowerCase();
  const converted = await converter.convert({
    data: options.bytes,
    fileFrom: `doc.${fileType}`,
    fileTo: "Editor.bin",
  });
  if (!converted.output || converted.output.byteLength === 0) {
    throw new Error("Não foi possível ler a outra versão.");
  }

  const frame = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
  const win = frame?.contentWindow as CompareHost | null | undefined;
  const compare = win?.AscCommonWord?.CompareDocuments;
  if (!win || !compare) throw new Error("O editor ainda não está pronto para comparar.");

  const binary = converted.output;
  const imageMap = options.server.urlsForCompareMedia(converted.media || {});
  const collab = win.AscCommon?.CollaborativeEditing;
  const previousUse = collab?.m_nUseType;
  // CompareBinary refuses the call unless the editor is in single-user mode.
  if (collab && previousUse !== 1) collab.m_nUseType = 1;
  try {
    compare({
      GetBinary: () => binary,
      GetImageMap: () => imageMap,
      IsValid: () => true,
    });
    const editMode = win.Asc?.c_oAscDisplayModeInReview?.Edit;
    if (typeof editMode === "number") {
      win.Asc?.editor?.asc_SetDisplayModeInReview?.(editMode);
    }
  } finally {
    if (collab && previousUse !== undefined && previousUse !== 1) {
      collab.m_nUseType = previousUse;
    }
  }
}
