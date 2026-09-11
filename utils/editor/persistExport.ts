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

const XLSX_FAMILY = new Set(["xlsx", "xls", "xlsm", "xlsb", "xltx", "xltm"]);
const PPTX_FAMILY = new Set(["pptx", "ppt", "pptm", "ppsx"]);

function outputMatchesOriginal(original: string, outputformat: number): boolean {
  const native = nativeFormatFromExt(original);
  if (native != null && outputformat === native) return true;
  if (XLSX_FAMILY.has(original)) {
    return (
      outputformat === AvsFileType.AVS_FILE_SPREADSHEET_XLSX ||
      outputformat === AvsFileType.AVS_FILE_SPREADSHEET_XLSX_FLAT ||
      outputformat === AvsFileType.AVS_FILE_SPREADSHEET_XLSX_PACKAGE ||
      outputformat === AvsFileType.AVS_FILE_SPREADSHEET_XLSM ||
      outputformat === AvsFileType.AVS_FILE_SPREADSHEET_XLS
    );
  }
  if (PPTX_FAMILY.has(original)) {
    return (
      outputformat === AvsFileType.AVS_FILE_PRESENTATION_PPTX ||
      outputformat === AvsFileType.AVS_FILE_PRESENTATION_PPTX_PACKAGE
    );
  }
  return false;
}

function formatMatchesOriginal(original: string, format: string): boolean {
  if (!format || format === original) return true;
  if (XLSX_FAMILY.has(original) && XLSX_FAMILY.has(format)) return true;
  if (PPTX_FAMILY.has(original) && PPTX_FAMILY.has(format)) return true;
  return false;
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
 * Spreadsheet editor often sends XLSX_FLAT / XLSX_PACKAGE instead of XLSX.
 */
export function isNativeOfficePersistExport(
  originalExt: string,
  cmd: DownloadAsCmd,
): boolean {
  const original = extOf(originalExt);
  if (!original) return false;
  if (isPdfDownloadAs(cmd) && original !== "pdf") return false;

  const format = extOf(cmd.format);
  if (format && !formatMatchesOriginal(original, format)) return false;

  const titleExt = cmd.title?.includes(".") ? extOf(cmd.title) : "";
  if (titleExt && !formatMatchesOriginal(original, titleExt)) return false;

  if (cmd.outputformat != null && !outputMatchesOriginal(original, cmd.outputformat)) {
    return false;
  }

  return true;
}
