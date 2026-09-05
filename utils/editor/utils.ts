import { DocumentType } from "./types";

export function getFileExt(name: string) {
  const type = name.split(".").pop() || "";
  return type.toLowerCase();
}

export enum AppType {
  word = 1,
  slide = 3,
  cell = 2,
  draw = 4,
  pdf = 5,
}

export const docTypeMap = {
  // Document
  docx: AppType.word,
  doc: AppType.word,
  odt: AppType.word,
  rtf: AppType.word,
  txt: AppType.word,
  html: AppType.word,
  mht: AppType.word,
  epub: AppType.word,
  fb2: AppType.word,
  mobi: AppType.word,
  docm: AppType.word,
  dotx: AppType.word,
  dotm: AppType.word,
  oform: AppType.word,
  docxf: AppType.word,

  // Presentation
  pptx: AppType.slide,
  ppt: AppType.slide,
  odp: AppType.slide,
  ppsx: AppType.slide,
  pptm: AppType.slide,
  ppsm: AppType.slide,
  potx: AppType.slide,
  potm: AppType.slide,
  otp: AppType.slide,
  odg: AppType.slide,

  // Spreadsheet
  xlsx: AppType.cell,
  xls: AppType.cell,
  ods: AppType.cell,
  csv: AppType.cell,
  xlsm: AppType.cell,
  xltx: AppType.cell,
  xltm: AppType.cell,
  xlsb: AppType.cell,
  ots: AppType.cell,

  // Draw
  vsdx: AppType.draw,
  vssx: AppType.draw,
  vstx: AppType.draw,
  vsdm: AppType.draw,
  vssm: AppType.draw,
  vstm: AppType.draw,

  // PDF
  pdf: AppType.pdf,
};

export function getDocumentType(ext: string) {
  const code = docTypeMap[ext.toLowerCase() as keyof typeof docTypeMap];
  const type = AppType[code] as DocumentType;
  return type || DocumentType.Word;
}

/** Returns the URL for creating a new document of the given type. */
export function getNewUrl(type: string) {
  return `/editor?new=${type}`;
}

export function getAppRoot(versionParam?: string | null): string {
  if (versionParam) {
    const clean = versionParam.trim();
    if (clean.startsWith("/")) return clean;
    if (clean.startsWith("v")) return `/${clean}`;
    return `/v${clean}`;
  }
  return process.env.NEXT_PUBLIC_APP_ROOT || "/v9.3.1-2";
}

export const APP_ROOT = getAppRoot();
export const PRELOAD_HTML = "/web-apps/apps/api/documents/preload.html";
export const API_JS = "/web-apps/apps/api/documents/api.js";

/**
 * OnlyOffice locale files are `pt.json` / `en.json`, not `pt-BR.json`.
 * Browser and workspace send BCP-47 tags; those 404 and the editor shows
 * "The interface language is not loaded".
 */
export function toOnlyOfficeLang(lang?: string | null): string {
  const primary = (lang || "pt").trim().split(/[-_]/)[0]?.toLowerCase();
  return primary || "pt";
}
