import { AvsFileType } from "./types";

export type DownloadAsCmd = {
  title?: string;
  outputformat?: number;
  format?: string;
  savetype?: number | string;
};

const PDF_OUTPUT = new Set([
  AvsFileType.AVS_FILE_CROSSPLATFORM_PDF,
  AvsFileType.AVS_FILE_CROSSPLATFORM_PDFA,
]);

export function nativeFormatFromExt(ext: string): number | undefined {
  switch (ext.replace(/^\./, "").toLowerCase()) {
    case "docx":
      return AvsFileType.AVS_FILE_DOCUMENT_DOCX;
    case "xlsx":
      return AvsFileType.AVS_FILE_SPREADSHEET_XLSX;
    case "pptx":
      return AvsFileType.AVS_FILE_PRESENTATION_PPTX;
    case "pdf":
      return AvsFileType.AVS_FILE_CROSSPLATFORM_PDF;
    default:
      return undefined;
  }
}

function extOf(value?: string): string {
  if (!value) return "";
  const base = value.includes(".") ? value.split(".").pop() || "" : value;
  return base.replace(/^\./, "").toLowerCase();
}

export function isPdfDownloadAs(cmd: DownloadAsCmd): boolean {
  const format = extOf(cmd.format);
  const titleExt = extOf(cmd.title);
  if (format === "pdf" || format === "pdfa") return true;
  if (titleExt === "pdf" || titleExt === "pdfa") return true;
  if (cmd.outputformat != null && PDF_OUTPUT.has(cmd.outputformat)) return true;
  return false;
}

/**
 * True when downloadAs is a save of the original Office file.
 * Print / Download as PDF / other exports must not overwrite the document.
 */
export function isNativeOfficePersistExport(
  originalExt: string,
  cmd: DownloadAsCmd,
): boolean {
  const original = extOf(originalExt);
  if (!original) return false;
  if (isPdfDownloadAs(cmd) && original !== "pdf") return false;

  const format = extOf(cmd.format);
  if (format && format !== original) return false;

  const titleExt = cmd.title?.includes(".") ? extOf(cmd.title) : "";
  if (titleExt && titleExt !== original) return false;

  const native = nativeFormatFromExt(original);
  if (cmd.outputformat != null && native != null && cmd.outputformat !== native) {
    return false;
  }

  return true;
}
